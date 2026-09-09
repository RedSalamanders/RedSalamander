#include "TerminalKittyImagePipeline.h"

#include <algorithm>
#include <chrono>
#include <limits>
#include <new>
#include <system_error>
#include <utility>

#include "Helpers.h"

namespace
{
[[nodiscard]] size_t SourceBytesPerPixel(TerminalKittyPixelFormat format) noexcept
{
    switch (format)
    {
    case TerminalKittyPixelFormat::Rgb:
        return 3u;
    case TerminalKittyPixelFormat::Rgba:
        return 4u;
    case TerminalKittyPixelFormat::GrayAlpha:
        return 2u;
    case TerminalKittyPixelFormat::Gray:
        return 1u;
    }
    return 0u;
}

[[nodiscard]] uint8_t Premultiply(uint8_t component, uint8_t alpha) noexcept
{
    return static_cast<uint8_t>(
        (static_cast<uint32_t>(component) * static_cast<uint32_t>(alpha) + 127u) / 255u);
}
} // namespace

TerminalKittyImagePipeline::~TerminalKittyImagePipeline()
{
    Stop();
}

HRESULT TerminalKittyImagePipeline::Start(NotifyReady notifyReady, void* notifyContext) noexcept
{
    std::scoped_lock lock(_mutex);
    if (_started)
    {
        return S_FALSE;
    }
    if (notifyReady == nullptr || notifyContext == nullptr)
    {
        return E_INVALIDARG;
    }

    _notifyReady = notifyReady;
    _notifyContext = notifyContext;
    _stopping = false;
    try
    {
        _worker = std::jthread([this](std::stop_token stopToken) noexcept { WorkerMain(stopToken); });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::system_error& error)
    {
        _notifyReady = nullptr;
        _notifyContext = nullptr;
        return HRESULT_FROM_WIN32(static_cast<DWORD>(error.code().value()));
    }
    _started = true;
    return S_OK;
}

bool TerminalKittyImagePipeline::ValidateWork(
    const TerminalKittyGenerationWork& work, size_t& sourceBytes, size_t& convertedBytes) noexcept
{
    sourceBytes = 0u;
    convertedBytes = 0u;
    if (work.requestId == 0u || work.storageGeneration == 0u)
    {
        return false;
    }
    for (const TerminalKittyImageWork& image : work.images)
    {
        const size_t bytesPerPixel = SourceBytesPerPixel(image.format);
        const uint64_t pixels = static_cast<uint64_t>(image.width) * static_cast<uint64_t>(image.height);
        const uint64_t expectedSource = pixels * bytesPerPixel;
        const uint64_t expectedConverted = pixels * 4u;
        if (image.imageId == 0u || image.imageGeneration == 0u || image.width == 0u || image.height == 0u ||
            bytesPerPixel == 0u || pixels > MaximumPixels || expectedSource != image.source.size() ||
            expectedSource > MaximumSourceBytes || expectedConverted > MaximumConvertedBytes ||
            sourceBytes > MaximumSourceBytes - static_cast<size_t>(expectedSource) ||
            convertedBytes > MaximumConvertedBytes - static_cast<size_t>(expectedConverted))
        {
            return false;
        }
        sourceBytes += static_cast<size_t>(expectedSource);
        convertedBytes += static_cast<size_t>(expectedConverted);
    }
    return true;
}

bool TerminalKittyImagePipeline::Submit(TerminalKittyGenerationWork work) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    size_t sourceBytes = 0u;
    size_t convertedBytes = 0u;
    if (! ValidateWork(work, sourceBytes, convertedBytes))
    {
        return false;
    }

    {
        std::scoped_lock lock(_mutex);
        if (! _started || _stopping || work.requestId <= _latestRequestId)
        {
            return false;
        }
        if (_pending.has_value())
        {
            ++_droppedPendingGenerations;
        }
        _latestRequestId = work.requestId;
        _pending = std::move(work);
        _queuedBytes = sourceBytes;
        if (_ready.has_value() && _ready->requestId != _latestRequestId)
        {
            _ready.reset();
            _readyBytes = 0u;
        }
        UpdateMemoryHighWaterLocked();
    }
    _workReady.notify_all();
    Debug::Perf::EmitDurationUs(
        L"terminal.kitty.submit_us", std::max<uint64_t>(1u, Debug::Perf::ElapsedUs(startedAt)), sourceBytes, convertedBytes);
    Debug::Perf::EmitValue(L"terminal.kitty.queue_bytes", sourceBytes);
    static_cast<void>(convertedBytes);
    return true;
}

std::optional<TerminalKittyGenerationResult> TerminalKittyImagePipeline::TakeReady(uint64_t requestId) noexcept
{
    std::scoped_lock lock(_mutex);
    if (! _ready.has_value())
    {
        return std::nullopt;
    }
    if (_ready->requestId != requestId)
    {
        _ready.reset();
        _readyBytes = 0u;
        return std::nullopt;
    }
    std::optional<TerminalKittyGenerationResult> result(std::move(_ready));
    _ready.reset();
    _readyBytes = 0u;
    return result;
}

void TerminalKittyImagePipeline::Invalidate(uint64_t requestId) noexcept
{
    {
        std::scoped_lock lock(_mutex);
        _latestRequestId = std::max(_latestRequestId, requestId);
        _pending.reset();
        _ready.reset();
        _queuedBytes = 0u;
        _readyBytes = 0u;
    }
    _workReady.notify_all();
}

void TerminalKittyImagePipeline::RequestStop() noexcept
{
    {
        std::scoped_lock lock(_mutex);
        _stopping = true;
        _pending.reset();
        _ready.reset();
        _queuedBytes = 0u;
        _readyBytes = 0u;
    }
    if (_worker.joinable())
    {
        _worker.request_stop();
    }
    _workReady.notify_all();
}

void TerminalKittyImagePipeline::Stop() noexcept
{
    RequestStop();
    if (_worker.joinable())
    {
        _worker.join();
    }
    std::scoped_lock lock(_mutex);
    _started = false;
    _notifyReady = nullptr;
    _notifyContext = nullptr;
    _activeSourceBytes = 0u;
    _activeConvertedBytes = 0u;
    _activeRequestId = 0u;
    _activeWorkers = 0u;
}

TerminalKittyPipelineStats TerminalKittyImagePipeline::GetStats() const noexcept
{
    std::scoped_lock lock(_mutex);
    return TerminalKittyPipelineStats{
        .queuedBytes = _queuedBytes,
        .activeSourceBytes = _activeSourceBytes,
        .activeConvertedBytes = _activeConvertedBytes,
        .readyBytes = _readyBytes,
        .residentBytes = _queuedBytes + _activeSourceBytes + _activeConvertedBytes + _readyBytes,
        .memoryHighWaterBytes = _memoryHighWaterBytes,
        .latestRequestId = _latestRequestId,
        .activeRequestId = _activeRequestId,
        .droppedPendingGenerations = _droppedPendingGenerations,
        .rejectedStaleGenerations = _rejectedStaleGenerations,
        .activeWorkers = _activeWorkers,
        .maximumActiveWorkers = _maximumActiveWorkers};
}

void TerminalKittyImagePipeline::UpdateMemoryHighWaterLocked() noexcept
{
    const size_t residentBytes = _queuedBytes + _activeSourceBytes + _activeConvertedBytes + _readyBytes;
    _memoryHighWaterBytes = std::max(_memoryHighWaterBytes, residentBytes);
}

bool TerminalKittyImagePipeline::Convert(
    TerminalKittyGenerationWork work, TerminalKittyGenerationResult& result) noexcept
{
    size_t sourceBytes = 0u;
    size_t convertedBytes = 0u;
    if (! ValidateWork(work, sourceBytes, convertedBytes))
    {
        return false;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    result = {};
    result.requestId = work.requestId;
    result.storageGeneration = work.storageGeneration;
    result.sourceBytes = sourceBytes;
    result.convertedBytes = convertedBytes;
#if defined(ENABLE_TESTS)
    {
        std::unique_lock lock(_mutex);
        _conversionStartedForTests = true;
        _workReady.notify_all();
        _workReady.wait(lock, [this] noexcept { return _stopping || ! _pauseAfterConversionStart; });
    }
#endif
    try
    {
        result.images.reserve(work.images.size());
        for (TerminalKittyImageWork& image : work.images)
        {
            const size_t sourceBytesPerPixel = SourceBytesPerPixel(image.format);
            const size_t pixelCount = static_cast<size_t>(image.width) * image.height;
            TerminalKittyConvertedImage converted{};
            converted.imageId = image.imageId;
            converted.width = image.width;
            converted.height = image.height;
            converted.imageGeneration = image.imageGeneration;
            converted.bgra.resize(pixelCount * 4u);
            for (size_t pixel = 0u; pixel < pixelCount; ++pixel)
            {
                const size_t sourceOffset = pixel * sourceBytesPerPixel;
                uint8_t red = image.source[sourceOffset];
                uint8_t green = red;
                uint8_t blue = red;
                uint8_t alpha = 0xFFu;
                if (image.format == TerminalKittyPixelFormat::Rgb || image.format == TerminalKittyPixelFormat::Rgba)
                {
                    green = image.source[sourceOffset + 1u];
                    blue = image.source[sourceOffset + 2u];
                    if (image.format == TerminalKittyPixelFormat::Rgba)
                    {
                        alpha = image.source[sourceOffset + 3u];
                    }
                }
                else if (image.format == TerminalKittyPixelFormat::GrayAlpha)
                {
                    alpha = image.source[sourceOffset + 1u];
                }
                const size_t destination = pixel * 4u;
                converted.bgra[destination] = Premultiply(blue, alpha);
                converted.bgra[destination + 1u] = Premultiply(green, alpha);
                converted.bgra[destination + 2u] = Premultiply(red, alpha);
                converted.bgra[destination + 3u] = alpha;
            }
            result.images.push_back(std::move(converted));
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    result.conversionDurationUs = Debug::Perf::ElapsedUs(startedAt);
    return true;
}

void TerminalKittyImagePipeline::WorkerMain(std::stop_token stopToken) noexcept
{
    while (! stopToken.stop_requested())
    {
        TerminalKittyGenerationWork work;
        size_t sourceBytes = 0u;
        size_t convertedBytes = 0u;
        {
            std::unique_lock lock(_mutex);
            _workReady.wait(lock, stopToken, [this] noexcept { return _stopping || _pending.has_value(); });
            if (_stopping || stopToken.stop_requested())
            {
                break;
            }
#if defined(ENABLE_TESTS)
            _workReady.wait(lock, stopToken, [this] noexcept { return _stopping || ! _pauseBeforeTake; });
            if (_stopping || stopToken.stop_requested())
            {
                break;
            }
#endif
            work = std::move(_pending.value());
            _pending.reset();
            _queuedBytes = 0u;
            static_cast<void>(ValidateWork(work, sourceBytes, convertedBytes));
            _activeSourceBytes = sourceBytes;
            _activeConvertedBytes = 0u;
            _activeRequestId = work.requestId;
            ++_activeWorkers;
            _maximumActiveWorkers = std::max(_maximumActiveWorkers, _activeWorkers);
            UpdateMemoryHighWaterLocked();
        }

        TerminalKittyGenerationResult result;
        const bool converted = Convert(std::move(work), result);
        const uint64_t resultRequestId = result.requestId;
        const uint64_t conversionDurationUs = result.conversionDurationUs;
        Debug::Perf::EmitValue(L"terminal.kitty.active_source_bytes", sourceBytes);
        Debug::Perf::EmitValue(
            L"terminal.kitty.active_converted_bytes", converted ? result.convertedBytes : 0u);
        NotifyReady notifyReady = nullptr;
        void* notifyContext = nullptr;
        bool published = false;
        {
            std::unique_lock lock(_mutex);
            _activeConvertedBytes = converted ? result.convertedBytes : 0u;
            UpdateMemoryHighWaterLocked();
#if defined(ENABLE_TESTS)
            _workReady.wait(lock, stopToken, [this] noexcept { return _stopping || ! _pauseBeforePublish; });
#endif
            const bool stopping = _stopping || stopToken.stop_requested();
            const bool stale = ! stopping && result.requestId != _latestRequestId;
            if (stale)
            {
                ++_rejectedStaleGenerations;
            }
            else if (! stopping && converted)
            {
                _ready = std::move(result);
                _readyBytes = _ready->convertedBytes;
                notifyReady = _notifyReady;
                notifyContext = _notifyContext;
                published = true;
            }
            _activeSourceBytes = 0u;
            _activeConvertedBytes = 0u;
            _activeRequestId = 0u;
            --_activeWorkers;
            UpdateMemoryHighWaterLocked();
            if (stopping)
            {
                break;
            }
        }

        Debug::Perf::EmitDurationUs(
            L"terminal.kitty.convert_us",
            std::max<uint64_t>(1u, conversionDurationUs),
            sourceBytes,
            convertedBytes,
            converted ? S_OK : E_INVALIDARG);
        const TerminalKittyPipelineStats stats = GetStats();
        Debug::Perf::EmitValue(L"terminal.kitty.memory_high_water_bytes", stats.memoryHighWaterBytes);
        if (published)
        {
            Debug::Perf::EmitValue(L"terminal.kitty.ready_bytes", convertedBytes);
        }
        if (published && resultRequestId == stats.latestRequestId && notifyReady != nullptr)
        {
            notifyReady(notifyContext);
        }
    }
}

#if defined(ENABLE_TESTS)
void TerminalKittyImagePipeline::SetPauseBeforeTakeForTests(bool pause) noexcept
{
    {
        std::scoped_lock lock(_mutex);
        _pauseBeforeTake = pause;
    }
    _workReady.notify_all();
}

void TerminalKittyImagePipeline::SetPauseBeforePublishForTests(bool pause) noexcept
{
    {
        std::scoped_lock lock(_mutex);
        _pauseBeforePublish = pause;
    }
    _workReady.notify_all();
}

void TerminalKittyImagePipeline::SetPauseAfterConversionStartForTests(bool pause) noexcept
{
    {
        std::scoped_lock lock(_mutex);
        _pauseAfterConversionStart = pause;
        if (pause)
        {
            _conversionStartedForTests = false;
        }
    }
    _workReady.notify_all();
}

bool TerminalKittyImagePipeline::WaitForConversionStartForTests(DWORD timeoutMilliseconds) noexcept
{
    std::unique_lock lock(_mutex);
    return _workReady.wait_for(lock,
                               std::chrono::milliseconds(timeoutMilliseconds),
                               [this] noexcept { return _conversionStartedForTests; });
}

#endif
