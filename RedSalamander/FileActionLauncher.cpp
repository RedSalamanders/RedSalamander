#include "FileActionLauncher.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <limits>
#include <memory>
#include <new>
#include <shellapi.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"

namespace
{
[[nodiscard]] HRESULT MissingMacroValue() noexcept
{
    return HRESULT_FROM_WIN32(ERROR_BAD_ARGUMENTS);
}

[[nodiscard]] HRESULT InvalidMacroSyntax() noexcept
{
    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

void AppendWindowsQuotedArgumentContent(std::wstring& out, std::wstring_view value)
{
    size_t backslashes = 0u;
    for (const wchar_t ch : value)
    {
        if (ch == L'\\')
        {
            ++backslashes;
            continue;
        }

        if (ch == L'"')
        {
            out.append((backslashes * 2u) + 1u, L'\\');
            out.push_back(L'"');
            backslashes = 0u;
            continue;
        }

        out.append(backslashes, L'\\');
        backslashes = 0u;
        out.push_back(ch);
    }

    out.append(backslashes * 2u, L'\\');
}

void AppendWindowsQuotedArgument(std::wstring& out, std::wstring_view value)
{
    out.push_back(L'"');
    AppendWindowsQuotedArgumentContent(out, value);
    out.push_back(L'"');
}

[[nodiscard]] bool IsMacroWrappedByTemplateQuotes(std::wstring_view templateText, size_t open, size_t close) noexcept
{
    return open > 0u && close + 1u < templateText.size() && templateText[open - 1u] == L'"' && templateText[close + 1u] == L'"';
}

[[nodiscard]] HRESULT WriteAllBytes(HANDLE file, const void* data, size_t byteCount) noexcept
{
    const auto* cursor  = static_cast<const std::byte*>(data);
    size_t totalWritten = 0u;
    while (totalWritten < byteCount)
    {
        const size_t remaining = byteCount - totalWritten;
        const DWORD chunk      = static_cast<DWORD>(std::min<size_t>(remaining, static_cast<size_t>(std::numeric_limits<DWORD>::max())));
        DWORD written          = 0u;
        if (WriteFile(file, cursor + totalWritten, chunk, &written, nullptr) == FALSE || written == 0u)
        {
            const DWORD error = GetLastError();
            return HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_WRITE_FAULT : error);
        }
        totalWritten += written;
    }

    return S_OK;
}

[[nodiscard]] std::filesystem::path CurrentDirectoryForContext(const FileActionLauncher::MacroContext& context)
{
    if (! context.currentDirectory.empty())
    {
        return context.currentDirectory;
    }
    if (! context.itemPath.empty())
    {
        return context.itemPath.parent_path();
    }
    return {};
}

[[nodiscard]] HRESULT ResolveMacro(std::wstring_view name, const FileActionLauncher::MacroContext& context, std::wstring& value) noexcept
{
    value.clear();

    if (name == L"Path")
    {
        const std::filesystem::path path = CurrentDirectoryForContext(context);
        if (path.empty())
        {
            return MissingMacroValue();
        }
        value = path.wstring();
        return S_OK;
    }

    if (name == L"FullPath" || name == L"PathAndFilename")
    {
        if (context.itemPath.empty())
        {
            return MissingMacroValue();
        }
        value = context.itemPath.wstring();
        return S_OK;
    }

    if (name == L"Filename")
    {
        if (context.itemPath.empty())
        {
            return MissingMacroValue();
        }
        value = context.itemPath.filename().wstring();
        if (value.empty())
        {
            return MissingMacroValue();
        }
        return S_OK;
    }

    if (name == L"SelectedPathsFile")
    {
        if (context.selectedPathsFile.empty())
        {
            return MissingMacroValue();
        }
        value = context.selectedPathsFile.wstring();
        return S_OK;
    }

    if (name == L"OppositePanePath")
    {
        if (context.oppositePanePath.empty())
        {
            return MissingMacroValue();
        }
        value = context.oppositePanePath.wstring();
        return S_OK;
    }

    if (name == L"ComputerName")
    {
        if (context.computerName.empty())
        {
            return MissingMacroValue();
        }
        value = context.computerName;
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
}

[[nodiscard]] HRESULT TemplateReferencesMacro(std::wstring_view templateText, std::wstring_view macroName, bool& references) noexcept
{
    references = false;

    for (size_t index = 0; index < templateText.size();)
    {
        const wchar_t ch = templateText[index];
        if (ch == L'{')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'{')
            {
                index += 2u;
                continue;
            }

            const size_t close = templateText.find(L'}', index + 1u);
            if (close == std::wstring_view::npos)
            {
                return InvalidMacroSyntax();
            }

            const std::wstring_view name = templateText.substr(index + 1u, close - index - 1u);
            if (name.empty())
            {
                return InvalidMacroSyntax();
            }
            if (name == macroName)
            {
                references = true;
                return S_OK;
            }

            index = close + 1u;
            continue;
        }

        if (ch == L'}')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'}')
            {
                index += 2u;
                continue;
            }
            return InvalidMacroSyntax();
        }

        ++index;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ActionReferencesSelectedPathsFile(const Common::Settings::FileActionDefinition& action, bool& references) noexcept
{
    references = false;

    bool currentReferences = false;
    if (const HRESULT hr = TemplateReferencesMacro(action.executablePath, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    if (const HRESULT hr = TemplateReferencesMacro(action.arguments, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    if (const HRESULT hr = TemplateReferencesMacro(action.workingDirectory, L"SelectedPathsFile", currentReferences); FAILED(hr))
    {
        return hr;
    }
    references = references || currentReferences;

    return S_OK;
}

[[nodiscard]] HRESULT ExpandMacrosInternal(std::wstring_view templateText,
                                           const FileActionLauncher::MacroContext& context,
                                           bool quoteArgumentMacros,
                                           std::wstring& out) noexcept
{
    out.clear();
    out.reserve(templateText.size());

    for (size_t index = 0; index < templateText.size();)
    {
        const wchar_t ch = templateText[index];
        if (ch == L'{')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'{')
            {
                out.push_back(L'{');
                index += 2u;
                continue;
            }

            const size_t close = templateText.find(L'}', index + 1u);
            if (close == std::wstring_view::npos)
            {
                return InvalidMacroSyntax();
            }

            const std::wstring_view name = templateText.substr(index + 1u, close - index - 1u);
            if (name.empty())
            {
                return InvalidMacroSyntax();
            }

            std::wstring value;
            const HRESULT hr = ResolveMacro(name, context, value);
            if (FAILED(hr))
            {
                return hr;
            }

            if (! quoteArgumentMacros)
            {
                out.append(value);
            }
            else if (IsMacroWrappedByTemplateQuotes(templateText, index, close))
            {
                AppendWindowsQuotedArgumentContent(out, value);
            }
            else
            {
                AppendWindowsQuotedArgument(out, value);
            }

            index = close + 1u;
            continue;
        }

        if (ch == L'}')
        {
            if (index + 1u < templateText.size() && templateText[index + 1u] == L'}')
            {
                out.push_back(L'}');
                index += 2u;
                continue;
            }

            return InvalidMacroSyntax();
        }

        out.push_back(ch);
        ++index;
    }

    return S_OK;
}

[[nodiscard]] HRESULT ExpandArgumentMacros(std::wstring_view templateText, const FileActionLauncher::MacroContext& context, std::wstring& out) noexcept
{
    return ExpandMacrosInternal(templateText, context, true, out);
}

void CleanupFiles(const std::vector<std::filesystem::path>& files) noexcept
{
    if (files.empty())
    {
        return;
    }

    Debug::Perf::Scope perf(L"fileaction.selected_paths_file.cleanup_us");
    perf.SetValue0(static_cast<uint64_t>(files.size()));

    uint64_t failureCount = 0u;
    for (const std::filesystem::path& file : files)
    {
        if (file.empty())
        {
            continue;
        }

        std::error_code ec;
        static_cast<void>(std::filesystem::remove(file, ec));
        if (ec)
        {
            ++failureCount;
            Debug::Warning(L"FileActionLauncher: failed to delete temporary selected-paths file '{}' (error={})", file.wstring(), ec.value());
        }
    }

    perf.SetValue1(failureCount);
    if (failureCount != 0u)
    {
        perf.SetHr(E_FAIL);
    }
}

[[nodiscard]] HRESULT CreateSelectedPathsFile(const std::vector<std::filesystem::path>& selectedPaths, std::filesystem::path& outPath) noexcept
{
    outPath.clear();

    Debug::Perf::Scope perf(L"fileaction.selected_paths_file.write_us");
    perf.SetValue0(static_cast<uint64_t>(selectedPaths.size()));

    if (selectedPaths.empty())
    {
        perf.SetHr(MissingMacroValue());
        return MissingMacroValue();
    }

    std::array<wchar_t, MAX_PATH + 1u> tempRoot{};
    const DWORD rootLength = GetTempPathW(static_cast<DWORD>(tempRoot.size()), tempRoot.data());
    if (rootLength == 0u || rootLength >= tempRoot.size())
    {
        const HRESULT hr = HRESULT_FROM_WIN32(rootLength == 0u ? GetLastError() : ERROR_BUFFER_OVERFLOW);
        perf.SetHr(hr);
        return hr;
    }

    std::array<wchar_t, MAX_PATH + 1u> tempFile{};
    if (GetTempFileNameW(tempRoot.data(), L"rsa", 0u, tempFile.data()) == 0u)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        return hr;
    }

    const std::filesystem::path filePath(tempFile.data());
    bool keepFile              = false;
    const auto deleteOnFailure = wil::scope_exit([&]() noexcept
    {
        if (! keepFile)
        {
            DeleteFileW(filePath.c_str());
        }
    });

    wil::unique_handle file(CreateFileW(filePath.c_str(), GENERIC_WRITE, 0u, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY, nullptr));
    if (! file)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        return hr;
    }

    std::wstring text;
    for (const std::filesystem::path& selectedPath : selectedPaths)
    {
        if (selectedPath.empty())
        {
            continue;
        }
        text.append(selectedPath.wstring());
        text.append(L"\r\n");
    }
    if (text.empty())
    {
        const HRESULT hr = MissingMacroValue();
        perf.SetHr(hr);
        return hr;
    }

    constexpr std::array<std::byte, 2u> kUtf16LeBom{{std::byte{0xFFu}, std::byte{0xFEu}}};
    if (const HRESULT hr = WriteAllBytes(file.get(), kUtf16LeBom.data(), kUtf16LeBom.size()); FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }

    const size_t textBytes = text.size() * sizeof(wchar_t);
    perf.SetValue1(static_cast<uint64_t>(textBytes + kUtf16LeBom.size()));
    if (const HRESULT hr = WriteAllBytes(file.get(), text.data(), textBytes); FAILED(hr))
    {
        perf.SetHr(hr);
        return hr;
    }

    keepFile = true;
    outPath  = filePath;
    return S_OK;
}

struct DeferredCleanupContext
{
    DeferredCleanupContext()                                         = default;
    DeferredCleanupContext(const DeferredCleanupContext&)            = delete;
    DeferredCleanupContext& operator=(const DeferredCleanupContext&) = delete;
    DeferredCleanupContext(DeferredCleanupContext&&)                 = delete;
    DeferredCleanupContext& operator=(DeferredCleanupContext&&)      = delete;

    wil::unique_threadpool_wait_nowait wait;
    wil::unique_handle process;
    std::vector<std::filesystem::path> files;
};

void CALLBACK DeferredCleanupCallback(PTP_CALLBACK_INSTANCE instance, void* rawContext, PTP_WAIT /*wait*/, TP_WAIT_RESULT /*waitResult*/) noexcept
{
    std::unique_ptr<DeferredCleanupContext> context(static_cast<DeferredCleanupContext*>(rawContext));
    if (! context)
    {
        return;
    }

    DisassociateCurrentThreadFromCallback(instance);
    context->wait.reset();
    context->process.reset();
    CleanupFiles(context->files);
}

[[nodiscard]] bool TryScheduleCleanupAfterProcessExit(wil::unique_handle process, const std::vector<std::filesystem::path>& files) noexcept
{
    if (! process || files.empty())
    {
        return true;
    }

    auto context = wil::make_unique_nothrow<DeferredCleanupContext>();
    if (! context)
    {
        return false;
    }

    context->process = std::move(process);
    context->files   = files;
    context->wait.reset(CreateThreadpoolWait(DeferredCleanupCallback, context.get(), nullptr));
    if (! context->wait)
    {
        return false;
    }

    SetThreadpoolWait(context->wait.get(), context->process.get(), nullptr);
    static_cast<void>(context.release());
    return true;
}

} // namespace

namespace FileActionLauncher
{
HRESULT ExpandMacros(std::wstring_view templateText, const MacroContext& context, std::wstring& out) noexcept
{
    return ExpandMacrosInternal(templateText, context, false, out);
}

HRESULT BuildExternalLaunchPlan(const Common::Settings::FileActionDefinition& action, const MacroContext& context, LaunchPlan& out) noexcept
{
    out = LaunchPlan{};

    if (action.kind != Common::Settings::FileActionKind::ExternalProgram || action.executablePath.empty())
    {
        return E_INVALIDARG;
    }

    MacroContext effectiveContext = context;
    std::vector<std::filesystem::path> createdCleanupFiles;
    bool keepCreatedFiles       = false;
    const auto cleanupOnFailure = wil::scope_exit([&]() noexcept
    {
        if (! keepCreatedFiles)
        {
            CleanupFiles(createdCleanupFiles);
        }
    });

    bool referencesSelectedPathsFile = false;
    if (const HRESULT hr = ActionReferencesSelectedPathsFile(action, referencesSelectedPathsFile); FAILED(hr))
    {
        return hr;
    }
    if (referencesSelectedPathsFile && effectiveContext.selectedPathsFile.empty())
    {
        std::vector<std::filesystem::path> selectedPaths = effectiveContext.selectedPaths;
        if (selectedPaths.empty() && ! effectiveContext.itemPath.empty())
        {
            selectedPaths.push_back(effectiveContext.itemPath);
        }

        std::filesystem::path selectedPathsFile;
        if (const HRESULT hr = CreateSelectedPathsFile(selectedPaths, selectedPathsFile); FAILED(hr))
        {
            return hr;
        }

        effectiveContext.selectedPathsFile = selectedPathsFile;
        createdCleanupFiles.push_back(selectedPathsFile);
        out.cleanupFilesAfterExit.push_back(selectedPathsFile);
    }

    if (const HRESULT hr = ExpandMacros(action.executablePath, effectiveContext, out.executablePath); FAILED(hr))
    {
        out = LaunchPlan{};
        return hr;
    }
    if (out.executablePath.empty())
    {
        out = LaunchPlan{};
        return E_INVALIDARG;
    }

    if (const HRESULT hr = ExpandArgumentMacros(action.arguments, effectiveContext, out.arguments); FAILED(hr))
    {
        out = LaunchPlan{};
        return hr;
    }

    if (! action.workingDirectory.empty())
    {
        if (const HRESULT hr = ExpandMacros(action.workingDirectory, effectiveContext, out.workingDirectory); FAILED(hr))
        {
            out = LaunchPlan{};
            return hr;
        }
    }
    else
    {
        out.workingDirectory = CurrentDirectoryForContext(effectiveContext).wstring();
    }

    keepCreatedFiles = true;
    return S_OK;
}

HRESULT LaunchExternalPlan(const LaunchPlan& plan, const LaunchOptions& options, LaunchResult* result) noexcept
{
    Debug::Perf::Scope perf(L"fileaction.external.launch_us");
    perf.SetDetail(options.waitForExit ? L"wait" : L"start");
    perf.SetValue0(static_cast<uint64_t>(plan.arguments.size()));
    perf.SetValue1(options.waitForExit ? options.waitTimeoutMs : 0u);

    if (result != nullptr)
    {
        *result = LaunchResult{};
    }

    if (plan.executablePath.empty())
    {
        perf.SetHr(E_INVALIDARG);
        return E_INVALIDARG;
    }

    SHELLEXECUTEINFOW executeInfo{};
    executeInfo.cbSize       = sizeof(executeInfo);
    executeInfo.fMask        = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI;
    executeInfo.hwnd         = options.ownerWindow;
    executeInfo.lpVerb       = L"open";
    executeInfo.lpFile       = plan.executablePath.c_str();
    executeInfo.lpParameters = plan.arguments.empty() ? nullptr : plan.arguments.c_str();
    executeInfo.lpDirectory  = plan.workingDirectory.empty() ? nullptr : plan.workingDirectory.c_str();
    executeInfo.nShow        = options.showCommand;

    if (ShellExecuteExW(&executeInfo) == FALSE)
    {
        const DWORD error = Debug::ErrorWithLastError(L"FileActionLauncher: ShellExecuteExW failed for external action '{}'", plan.executablePath);
        const HRESULT hr  = HRESULT_FROM_WIN32(error == ERROR_SUCCESS ? ERROR_FILE_NOT_FOUND : error);
        perf.SetHr(hr);
        CleanupFiles(plan.cleanupFilesAfterExit);
        return hr;
    }

    wil::unique_handle processHandle(executeInfo.hProcess);
    if (processHandle)
    {
        if (result != nullptr)
        {
            result->processId = GetProcessId(processHandle.get());
        }
        if (result != nullptr && options.captureProcessHandle)
        {
            HANDLE duplicatedHandle = nullptr;
            if (DuplicateHandle(GetCurrentProcess(),
                                processHandle.get(),
                                GetCurrentProcess(),
                                &duplicatedHandle,
                                SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION,
                                FALSE,
                                0) != FALSE)
            {
                result->processHandle.reset(duplicatedHandle);
            }
            else
            {
                Debug::Warning(L"FileActionLauncher: failed to duplicate external process handle for tracking.");
            }
        }
    }
    if (! options.waitForExit)
    {
        if (! TryScheduleCleanupAfterProcessExit(std::move(processHandle), plan.cleanupFilesAfterExit))
        {
            Debug::Warning(L"FileActionLauncher: failed to schedule temporary selected-paths cleanup after external process exit.");
        }
        return S_OK;
    }
    if (! processHandle)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        perf.SetHr(hr);
        return hr;
    }

    const DWORD waitResult = WaitForSingleObject(processHandle.get(), options.waitTimeoutMs);
    if (waitResult == WAIT_TIMEOUT)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        perf.SetHr(hr);
        return hr;
    }
    if (waitResult == WAIT_FAILED)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        return hr;
    }
    if (waitResult != WAIT_OBJECT_0)
    {
        perf.SetHr(E_FAIL);
        return E_FAIL;
    }

    DWORD exitCode = 0;
    if (GetExitCodeProcess(processHandle.get(), &exitCode) == FALSE)
    {
        const HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
        perf.SetHr(hr);
        return hr;
    }

    if (result != nullptr)
    {
        result->exitCodeAvailable = true;
        result->exitCode          = exitCode;
    }

    CleanupFiles(plan.cleanupFilesAfterExit);
    return S_OK;
}
} // namespace FileActionLauncher
