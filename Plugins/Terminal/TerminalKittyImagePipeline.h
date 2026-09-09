#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <stop_token>
#include <thread>
#include <vector>

enum class TerminalKittyPixelFormat : uint8_t
{
    Rgb,
    Rgba,
    GrayAlpha,
    Gray,
};

struct TerminalKittyImageWork final
{
    uint32_t imageId                = 0u;
    uint32_t width                  = 0u;
    uint32_t height                 = 0u;
    uint64_t imageGeneration        = 0u;
    TerminalKittyPixelFormat format = TerminalKittyPixelFormat::Rgba;
    std::vector<uint8_t> source;
};

struct TerminalKittyGenerationWork final
{
    uint64_t requestId         = 0u;
    uint64_t storageGeneration = 0u;
    std::vector<TerminalKittyImageWork> images;
};

struct TerminalKittyConvertedImage final
{
    uint32_t imageId         = 0u;
    uint32_t width           = 0u;
    uint32_t height          = 0u;
    uint64_t imageGeneration = 0u;
    std::vector<uint8_t> bgra;
};

struct TerminalKittyGenerationResult final
{
    uint64_t requestId            = 0u;
    uint64_t storageGeneration    = 0u;
    uint64_t conversionDurationUs = 0u;
    size_t sourceBytes            = 0u;
    size_t convertedBytes         = 0u;
    std::vector<TerminalKittyConvertedImage> images;
};

struct TerminalKittyPipelineStats final
{
    size_t queuedBytes                 = 0u;
    size_t activeSourceBytes           = 0u;
    size_t activeConvertedBytes        = 0u;
    size_t readyBytes                  = 0u;
    size_t residentBytes               = 0u;
    size_t memoryHighWaterBytes        = 0u;
    uint64_t latestRequestId           = 0u;
    uint64_t activeRequestId           = 0u;
    uint64_t droppedPendingGenerations = 0u;
    uint64_t rejectedStaleGenerations  = 0u;
    uint32_t activeWorkers             = 0u;
    uint32_t maximumActiveWorkers      = 0u;
};

class TerminalKittyImagePipeline final
{
public:
    using NotifyReady = void (*)(void* context) noexcept;

    static constexpr uint64_t MaximumPixels       = 16u * 1024u * 1024u;
    static constexpr size_t MaximumSourceBytes    = 64u * 1024u * 1024u;
    static constexpr size_t MaximumConvertedBytes = 64u * 1024u * 1024u;

    TerminalKittyImagePipeline() = default;
    ~TerminalKittyImagePipeline();

    TerminalKittyImagePipeline(const TerminalKittyImagePipeline&)            = delete;
    TerminalKittyImagePipeline(TerminalKittyImagePipeline&&)                 = delete;
    TerminalKittyImagePipeline& operator=(const TerminalKittyImagePipeline&) = delete;
    TerminalKittyImagePipeline& operator=(TerminalKittyImagePipeline&&)      = delete;

    [[nodiscard]] HRESULT Start(NotifyReady notifyReady, void* notifyContext) noexcept;
    [[nodiscard]] bool Submit(TerminalKittyGenerationWork work) noexcept;
    [[nodiscard]] std::optional<TerminalKittyGenerationResult> TakeReady(uint64_t requestId) noexcept;
    void Invalidate(uint64_t requestId) noexcept;
    void RequestStop() noexcept;
    void Stop() noexcept;
    [[nodiscard]] TerminalKittyPipelineStats GetStats() const noexcept;

#if defined(ENABLE_TESTS)
    void SetPauseBeforeTakeForTests(bool pause) noexcept;
    void SetPauseBeforePublishForTests(bool pause) noexcept;
    void SetPauseAfterConversionStartForTests(bool pause) noexcept;
    [[nodiscard]] bool WaitForConversionStartForTests(DWORD timeoutMilliseconds) noexcept;
#endif

private:
    [[nodiscard]] static bool ValidateWork(const TerminalKittyGenerationWork& work, size_t& sourceBytes, size_t& convertedBytes) noexcept;
    [[nodiscard]] bool Convert(TerminalKittyGenerationWork work, TerminalKittyGenerationResult& result) noexcept;
    void WorkerMain(std::stop_token stopToken) noexcept;
    void UpdateMemoryHighWaterLocked() noexcept;

    mutable std::mutex _mutex;
    std::condition_variable_any _workReady;
    std::optional<TerminalKittyGenerationWork> _pending;
    std::optional<TerminalKittyGenerationResult> _ready;
    std::jthread _worker;
    NotifyReady _notifyReady            = nullptr;
    void* _notifyContext                = nullptr;
    size_t _queuedBytes                 = 0u;
    size_t _activeSourceBytes           = 0u;
    size_t _activeConvertedBytes        = 0u;
    size_t _readyBytes                  = 0u;
    size_t _memoryHighWaterBytes        = 0u;
    uint64_t _latestRequestId           = 0u;
    uint64_t _activeRequestId           = 0u;
    uint64_t _droppedPendingGenerations = 0u;
    uint64_t _rejectedStaleGenerations  = 0u;
    uint32_t _activeWorkers             = 0u;
    uint32_t _maximumActiveWorkers      = 0u;
    bool _started                       = false;
    bool _stopping                      = false;
#if defined(ENABLE_TESTS)
    bool _pauseBeforeTake           = false;
    bool _pauseBeforePublish        = false;
    bool _pauseAfterConversionStart = false;
    bool _conversionStartedForTests = false;
#endif
};
