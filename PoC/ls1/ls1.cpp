// fastwalk.cpp - ultra-fast recursive file+folder lister for NTFS/ReFS
// Build: cl /O2 /std:c++20 fastwalk.cpp
#include <cstdio>
#include <deque>
#include <string>
#include <vector>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

static inline bool IsDotOrDotDot(const wchar_t* name)
{
    return (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0')));
}

// Convert to extended-length path so we skip legacy MAX_PATH checks.
// Normal:  \\?\C:\path
// UNC:     \\?\UNC\server\share\path
static std::wstring ToExtendedPath(const std::wstring& p)
{
    if (p.rfind(L"\\\\?\\", 0) == 0)
        return p; // already extended
    if (p.rfind(L"\\\\", 0) == 0)
    {
        // \\server\share\... -> \\?\UNC\server\share\...
        return L"\\\\?\\UNC\\" + p.substr(2);
    }
    return L"\\\\?\\" + p;
}

static inline void EnsureConsoleUtf8()
{
    // Makes console accept UTF-8 if not redirected. (No harm if redirected.)
    SetConsoleOutputCP(CP_UTF8);
}

// Write one UTF-8 line to stdout efficiently (works for console or redirection).
static void WriteLineUtf8(const std::wstring& wline)
{
    // Convert UTF-16 -> UTF-8
    int need = WideCharToMultiByte(CP_UTF8, 0, wline.c_str(), (int)wline.size(), nullptr, 0, nullptr, nullptr);
    if (need <= 0)
        return;
    std::string utf8;
    utf8.resize((size_t)need);
    WideCharToMultiByte(CP_UTF8, 0, wline.c_str(), (int)wline.size(), (LPSTR)(utf8.data()), need, nullptr, nullptr);

    // add newline
    static const char nl[2] = {'\n', '\0'};
    HANDLE hout             = GetStdHandle(STD_OUTPUT_HANDLE);
    DWORD written           = 0;
    WriteFile(hout, utf8.data(), (DWORD)utf8.size(), &written, nullptr);
    WriteFile(hout, nl, 1, &written, nullptr);
}

// Append "child" to "base" making sure there is exactly one backslash.
static std::wstring JoinPath(const std::wstring& base, const std::wstring& child)
{
    if (base.empty())
        return child;
    wchar_t last   = base.back();
    bool needSlash = (last != L'\\' && last != L'/');
    if (needSlash)
        return base + L'\\' + child;
    return base + child;
}

static void WalkTree(const std::wstring& startInput)
{
    std::wstring start = startInput.empty() ? L"." : startInput;
    // Normalize to extended-length path once at the root.
    std::wstring root = ToExtendedPath(start);

    // Reduce system error popups (e.g., inaccessible folders).
    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);

    // Non-recursive DFS using a stack for cache-friendliness.
    std::vector<std::wstring> stack;
    stack.reserve(4096);
    stack.push_back(root);

    // Print the starting directory as well (the user asked for files + folders).
    WriteLineUtf8(start);

    WIN32_FIND_DATAW ffd;

    while (! stack.empty())
    {
        std::wstring dir = std::move(stack.back());
        stack.pop_back();

        // Enumerate everything in dir
        std::wstring pattern = JoinPath(dir, L"*");

        HANDLE hFind = FindFirstFileExW(pattern.c_str(),
                                        FindExInfoBasic, // cheaper than standard; skips 8.3 and extras
                                        &ffd,
                                        FindExSearchNameMatch,
                                        nullptr,
                                        FIND_FIRST_EX_LARGE_FETCH // batch results to reduce syscalls (local only)
        );

        if (hFind == INVALID_HANDLE_VALUE)
        {
            // Access denied or transient error: just continue.
            continue;
        }

        do
        {
            const wchar_t* name = ffd.cFileName;
            if (IsDotOrDotDot(name))
                continue;

            std::wstring full = JoinPath(dir, name);

            // Print every item (directories + files)
            // Strip the \\?\ for display niceness if desired; here we keep paths as-is for speed.
            // If you prefer pretty output, remove the \\?\ prefix when present.
            // For consistent encoding, we already output UTF-8.
            // Note: printing is often the bottleneck; redirect to a file for maximum throughput.
            //       e.g., fastwalk.exe D:\ > list.txt
            WriteLineUtf8(full);

            // Recurse into directories, but avoid reparse points (symlinks/junctions) to prevent loops.
            if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
            {
                if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0)
                {
                    stack.push_back(full);
                }
            }
        } while (FindNextFileW(hFind, &ffd));

        FindClose(hFind);
    }
}

int wmain(int argc, wchar_t** argv)
{
    EnsureConsoleUtf8();

    std::wstring start;
    if (argc >= 2)
    {
        start = argv[1];
    }
    else
    {
        start = L".";
    }
    WalkTree(start);
    return 0;
}
