#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <shellapi.h>

#include <algorithm>
#include <cwchar>
#include <format>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#pragma comment(lib, "Shell32.lib")

namespace
{
constexpr wchar_t kTargetExeName[] = L"RedSalamander.exe";
constexpr wchar_t kErrorCaption[]  = L"RedSalamander Launcher";

[[nodiscard]] std::wstring StripLongPathPrefix(std::wstring path)
{
    static constexpr std::wstring_view kDosPrefix = LR"(\\?\)";
    static constexpr std::wstring_view kUncPrefix = LR"(\\?\UNC\)";

    if (path.rfind(kUncPrefix, 0u) == 0u)
    {
        std::wstring stripped(LR"(\\)");
        stripped.append(path.substr(kUncPrefix.size()));
        return stripped;
    }

    if (path.rfind(kDosPrefix, 0u) == 0u)
    {
        return path.substr(kDosPrefix.size());
    }

    return path;
}

[[nodiscard]] std::wstring GetCurrentModulePath()
{
    std::wstring path(MAX_PATH, L'\0');
    for (;;)
    {
        const DWORD written = ::GetModuleFileNameW(nullptr, path.data(), static_cast<DWORD>(path.size()));
        if (written == 0u)
        {
            return {};
        }

        const size_t writtenChars = static_cast<size_t>(written);
        if (writtenChars < path.size())
        {
            path.resize(writtenChars);
            return path;
        }

        path.resize(path.size() * 2u);
    }
}

[[nodiscard]] std::wstring GetFinalPathFromHandle(HANDLE file)
{
    std::wstring path(MAX_PATH, L'\0');
    for (;;)
    {
        const DWORD written =
            ::GetFinalPathNameByHandleW(file, path.data(), static_cast<DWORD>(path.size()), FILE_NAME_NORMALIZED | VOLUME_NAME_DOS);
        if (written == 0u)
        {
            return {};
        }

        const size_t writtenChars = static_cast<size_t>(written);
        if (writtenChars < path.size())
        {
            path.resize(writtenChars);
            return StripLongPathPrefix(std::move(path));
        }

        path.resize(writtenChars + 1u);
    }
}

[[nodiscard]] std::wstring ResolveFinalFilePath(std::wstring_view path)
{
    const std::wstring pathText(path);
    wil::unique_handle file(::CreateFileW(pathText.c_str(),
                                          FILE_READ_ATTRIBUTES,
                                          FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                          nullptr,
                                          OPEN_EXISTING,
                                          FILE_ATTRIBUTE_NORMAL,
                                          nullptr));
    if (! file)
    {
        return {};
    }

    return GetFinalPathFromHandle(file.get());
}

[[nodiscard]] std::wstring GetParentDirectory(std::wstring_view path)
{
    const size_t separator = path.find_last_of(L"\\/");
    if (separator == std::wstring_view::npos)
    {
        return {};
    }

    return std::wstring(path.substr(0u, separator));
}

[[nodiscard]] std::wstring CombinePath(std::wstring_view directory, std::wstring_view leaf)
{
    std::wstring path(directory);
    if (! path.empty() && path.back() != L'\\' && path.back() != L'/')
    {
        path.push_back(L'\\');
    }
    path.append(leaf);
    return path;
}

[[nodiscard]] std::wstring QuoteCommandLineArgument(std::wstring_view argument)
{
    if (argument.empty())
    {
        return L"\"\"";
    }

    if (argument.find_first_of(L" \t\n\v\"") == std::wstring_view::npos)
    {
        return std::wstring(argument);
    }

    std::wstring quoted;
    quoted.push_back(L'"');

    size_t pendingBackslashes = 0u;
    for (const wchar_t ch : argument)
    {
        if (ch == L'\\')
        {
            ++pendingBackslashes;
            continue;
        }

        if (ch == L'"')
        {
            quoted.append((pendingBackslashes * 2u) + 1u, L'\\');
            quoted.push_back(ch);
            pendingBackslashes = 0u;
            continue;
        }

        quoted.append(pendingBackslashes, L'\\');
        pendingBackslashes = 0u;
        quoted.push_back(ch);
    }

    quoted.append(pendingBackslashes * 2u, L'\\');
    quoted.push_back(L'"');
    return quoted;
}

[[nodiscard]] std::wstring BuildTargetCommandLine(std::wstring_view targetPath, int argc, wchar_t** argv)
{
    std::wstring commandLine = QuoteCommandLineArgument(targetPath);
    for (int index = 1; index < argc; ++index)
    {
        commandLine.push_back(L' ');
        commandLine.append(QuoteCommandLineArgument(argv[index] ? std::wstring_view(argv[index]) : std::wstring_view()));
    }

    return commandLine;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
        right.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }

    if (left.empty() || right.empty())
    {
        return left.empty() && right.empty();
    }

    return ::CompareStringOrdinal(left.data(),
                                  static_cast<int>(left.size()),
                                  right.data(),
                                  static_cast<int>(right.size()),
                                  TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool ShouldWaitForTargetExit(int argc, wchar_t** argv) noexcept
{
    if (argc <= 1 || argv == nullptr)
    {
        return false;
    }

    constexpr std::wstring_view kForegroundWaitArgs[] = {
        L"--selftest",
        L"--compare-selftest",
        L"--commands-selftest",
        L"--fileops-selftest",
        L"--selftest-list-cases",
    };

    for (int index = 1; index < argc; ++index)
    {
        const wchar_t* arg = argv[index];
        if (arg == nullptr || arg[0] == L'\0')
        {
            continue;
        }

        const std::wstring_view argView(arg);
        for (const std::wstring_view waitArg : kForegroundWaitArgs)
        {
            if (EqualsNoCase(argView, waitArg))
            {
                return true;
            }
        }
    }

    return false;
}

void ShowError(std::wstring_view message)
{
    const std::wstring messageText(message);
    static_cast<void>(::MessageBoxW(nullptr, messageText.c_str(), kErrorCaption, MB_OK | MB_ICONERROR));
}

int LaunchRedSalamander()
{
    const std::wstring modulePath = GetCurrentModulePath();
    if (modulePath.empty())
    {
        ShowError(L"Failed to resolve the launcher path.");
        return 1;
    }

    const std::wstring finalLauncherPath = ResolveFinalFilePath(modulePath);
    if (finalLauncherPath.empty())
    {
        ShowError(std::format(L"Failed to resolve the launcher target path for '{}'.", modulePath));
        return 1;
    }

    const std::wstring packageDirectory = GetParentDirectory(finalLauncherPath);
    if (packageDirectory.empty())
    {
        ShowError(std::format(L"Failed to resolve the package directory from '{}'.", finalLauncherPath));
        return 1;
    }

    const std::wstring targetPath = CombinePath(packageDirectory, kTargetExeName);

    int argc = 0;
    wil::unique_hlocal_ptr<wchar_t*> argv(::CommandLineToArgvW(::GetCommandLineW(), &argc));
    if (! argv)
    {
        ShowError(L"Failed to parse the launcher command line.");
        return 1;
    }
    const bool waitForTargetExit = ShouldWaitForTargetExit(argc, argv.get());

    std::wstring commandLine = BuildTargetCommandLine(targetPath, argc, argv.get());
    std::vector<wchar_t> commandLineBuffer(commandLine.begin(), commandLine.end());
    commandLineBuffer.push_back(L'\0');

    STARTUPINFOW startupInfo{};
    ::GetStartupInfoW(&startupInfo);

    PROCESS_INFORMATION processInfo{};
    if (::CreateProcessW(targetPath.c_str(), commandLineBuffer.data(), nullptr, nullptr, FALSE, 0u, nullptr, nullptr, &startupInfo, &processInfo) == FALSE)
    {
        const DWORD error = ::GetLastError();
        ShowError(std::format(L"Failed to launch '{}'. Win32 error {}.", targetPath, error));
        return 1;
    }

    wil::unique_handle process(processInfo.hProcess);
    wil::unique_handle thread(processInfo.hThread);
    if (waitForTargetExit)
    {
        if (::WaitForSingleObject(process.get(), INFINITE) == WAIT_FAILED)
        {
            const DWORD error = ::GetLastError();
            ShowError(std::format(L"Failed while waiting for '{}'. Win32 error {}.", targetPath, error));
            return 1;
        }

        DWORD exitCode = 1u;
        if (::GetExitCodeProcess(process.get(), &exitCode) == FALSE)
        {
            const DWORD error = ::GetLastError();
            ShowError(std::format(L"Failed to read the exit code from '{}'. Win32 error {}.", targetPath, error));
            return 1;
        }

        return exitCode <= static_cast<DWORD>((std::numeric_limits<int>::max)()) ? static_cast<int>(exitCode) : 1;
    }

    return 0;
}
} // namespace

#if defined(REDLAUNCHER_GUI)
int WINAPI wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    return LaunchRedSalamander();
}
#else
int wmain()
{
    return LaunchRedSalamander();
}
#endif
