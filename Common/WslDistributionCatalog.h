#pragma once

#include <string>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#ifndef COMMON_API
#ifdef COMMON_EXPORTS
#define COMMON_API __declspec(dllexport)
#else
#define COMMON_API __declspec(dllimport)
#endif
#endif

namespace Common::Wsl
{
inline constexpr wchar_t kDistributionRegistryPath[] = L"Software\\Microsoft\\Windows\\CurrentVersion\\Lxss";

struct DistributionRecord
{
    std::wstring name;
    std::wstring guid;
    std::wstring basePath;
    bool isWsl2 = false;
};

[[nodiscard]] COMMON_API std::vector<DistributionRecord> EnumerateDistributions() noexcept;
[[nodiscard]] COMMON_API std::vector<DistributionRecord> EnumerateDistributions(HKEY root, const wchar_t* subKey) noexcept;
[[nodiscard]] COMMON_API bool IsInstalled() noexcept;
} // namespace Common::Wsl
