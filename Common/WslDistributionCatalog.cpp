#include "WslDistributionCatalog.h"

#include "Helpers.h"
#include "RegistryUtils.h"

#include <algorithm>
#include <chrono>
#include <iterator>
#include <string_view>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/registry.h>
#include <wil/resource.h>
#pragma warning(pop)

namespace
{
constexpr DWORD kMaximumDistributionSubkeys = 1024u;
constexpr size_t kGuidChars                  = 38u;
constexpr std::wstring_view kDockerPrefix    = L"docker-desktop";
constexpr std::wstring_view kRancherPrefix   = L"rancher-desktop";

[[nodiscard]] bool StartsWithNoCase(std::wstring_view value, std::wstring_view prefix) noexcept
{
    return value.size() >= prefix.size() &&
           CompareStringOrdinal(value.data(), static_cast<int>(prefix.size()), prefix.data(), static_cast<int>(prefix.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool ShouldExclude(std::wstring_view name) noexcept
{
    return StartsWithNoCase(name, kDockerPrefix) || StartsWithNoCase(name, kRancherPrefix);
}

[[nodiscard]] bool IsGuidSubkey(std::wstring_view name) noexcept
{
    return name.size() == kGuidChars && name.front() == L'{' && name.back() == L'}';
}

[[nodiscard]] bool ReadModernFlag(HKEY key) noexcept
{
    uint32_t value = 0u;
    return SUCCEEDED(wil::reg::get_value_dword_nothrow(key, L"Modern", &value)) && value == 1u;
}

[[nodiscard]] int CompareNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE);
}
} // namespace

namespace Common::Wsl
{
std::vector<DistributionRecord> EnumerateDistributions() noexcept
{
    return EnumerateDistributions(HKEY_CURRENT_USER, kDistributionRegistryPath);
}

std::vector<DistributionRecord> EnumerateDistributions(HKEY root, const wchar_t* subKey) noexcept
{
    const auto startedAt = std::chrono::steady_clock::now();
    std::vector<DistributionRecord> distributions;
    uint64_t inspectedSubkeys = 0u;
    HRESULT resultStatus      = S_OK;
    const auto emitMetric     = wil::scope_exit([&]() noexcept
    {
        Debug::Perf::Emit(L"wsl.catalog.enumerate_us",
                          L"registry",
                          Debug::Perf::ElapsedUs(startedAt),
                          distributions.size(),
                          inspectedSubkeys,
                          resultStatus);
    });

    if (root == nullptr || subKey == nullptr)
    {
        resultStatus = E_INVALIDARG;
        return distributions;
    }

    wil::unique_hkey catalogKey;
    const HRESULT openHr = wil::reg::open_unique_key_nothrow(root, subKey, catalogKey, wil::reg::key_access::read);
    if (FAILED(openHr))
    {
        resultStatus = openHr;
        return distributions;
    }

    for (DWORD index = 0u; index < kMaximumDistributionSubkeys; ++index)
    {
        wchar_t guidBuffer[kGuidChars + 1u]{};
        DWORD guidChars      = static_cast<DWORD>(std::size(guidBuffer));
        const LSTATUS status = RegEnumKeyExW(catalogKey.get(), index, guidBuffer, &guidChars, nullptr, nullptr, nullptr, nullptr);
        if (status == ERROR_NO_MORE_ITEMS)
        {
            break;
        }
        ++inspectedSubkeys;
        if (status != ERROR_SUCCESS)
        {
            continue;
        }

        const std::wstring_view guidView(guidBuffer, guidChars);
        if (! IsGuidSubkey(guidView))
        {
            continue;
        }

        wil::unique_hkey distributionKey;
        if (FAILED(wil::reg::open_unique_key_nothrow(catalogKey.get(), guidBuffer, distributionKey, wil::reg::key_access::read)))
        {
            continue;
        }

        const auto name     = Registry::ReadBoundedStringValue(distributionKey.get(), L"DistributionName");
        const auto basePath = Registry::ReadBoundedStringValue(distributionKey.get(), L"BasePath");
        if (! name.has_value() || ! basePath.has_value() || ShouldExclude(name.value()))
        {
            continue;
        }

        DistributionRecord record;
        record.name     = name.value();
        record.guid     = guidView;
        record.basePath = basePath.value();
        record.isWsl2   = ReadModernFlag(distributionKey.get());
        distributions.push_back(std::move(record));
    }

    std::sort(distributions.begin(), distributions.end(), [](const DistributionRecord& left, const DistributionRecord& right) noexcept
    {
        const int nameComparison = CompareNoCase(left.name, right.name);
        if (nameComparison != CSTR_EQUAL)
        {
            return nameComparison == CSTR_LESS_THAN;
        }

        const int guidComparison = CompareNoCase(left.guid, right.guid);
        if (guidComparison != CSTR_EQUAL)
        {
            return guidComparison == CSTR_LESS_THAN;
        }
        return left.guid < right.guid;
    });

    return distributions;
}

bool IsInstalled() noexcept
{
    wil::unique_hkey catalogKey;
    return SUCCEEDED(wil::reg::open_unique_key_nothrow(
        HKEY_CURRENT_USER, kDistributionRegistryPath, catalogKey, wil::reg::key_access::read));
}
} // namespace Common::Wsl
