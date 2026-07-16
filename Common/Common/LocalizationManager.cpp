#include "LocalizationManager.h"

#include <algorithm>
#include <cstddef>
#include <filesystem>
#include <limits>
#include <mutex>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#pragma warning(pop)

#include "Helpers.h"

namespace Localization
{
namespace
{
struct ResourceOwner final
{
    ResourceOwner()                                    = default;
    ResourceOwner(const ResourceOwner&)                = delete;
    ResourceOwner& operator=(const ResourceOwner&)     = delete;
    ResourceOwner(ResourceOwner&&) noexcept            = default;
    ResourceOwner& operator=(ResourceOwner&&) noexcept = default;
    ~ResourceOwner()                                   = default;

    std::wstring moduleName;
    HINSTANCE embeddedInstance = nullptr;
    wil::unique_hmodule satellite;
    std::wstring loadedCulture;
    size_t registrationCount = 1u;
};

struct LocalizationState final
{
    LocalizationState()                                    = default;
    LocalizationState(const LocalizationState&)            = delete;
    LocalizationState& operator=(const LocalizationState&) = delete;
    LocalizationState(LocalizationState&&)                 = delete;
    LocalizationState& operator=(LocalizationState&&)      = delete;
    ~LocalizationState()                                   = default;

    std::mutex mutex;
    LanguagePreference preference;
    std::unordered_map<HINSTANCE, ResourceOwner> ownersByInstance;
    std::unordered_set<std::wstring> warnedLoads;
};

[[nodiscard]] LocalizationState& State() noexcept
{
    static LocalizationState* state = new LocalizationState();
    return *state;
}

[[nodiscard]] HINSTANCE NormalizeInstance(HINSTANCE instance) noexcept
{
    return instance ? instance : GetModuleHandleW(nullptr);
}

[[nodiscard]] std::filesystem::path GetModuleDirectory(HINSTANCE instance) noexcept
{
    std::wstring path(MAX_PATH, L'\0');
    for (;;)
    {
        const DWORD copied = GetModuleFileNameW(reinterpret_cast<HMODULE>(instance), path.data(), static_cast<DWORD>(path.size()));
        if (copied == 0)
        {
            return {};
        }
        if (copied < path.size() - 1u)
        {
            path.resize(copied);
            break;
        }
        path.resize(path.size() * 2u);
    }

    return std::filesystem::path(path).parent_path();
}

[[nodiscard]] std::wstring FoldCulture(std::wstring_view culture)
{
    std::wstring result(culture);
    std::replace(result.begin(), result.end(), L'_', L'-');
    return result;
}

[[nodiscard]] bool IsEnglishCulture(std::wstring_view culture) noexcept
{
    return culture.size() >= 2u && (culture[0] == L'e' || culture[0] == L'E') && (culture[1] == L'n' || culture[1] == L'N') &&
           (culture.size() == 2u || culture[2] == L'-');
}

void AddCultureWithParents(std::vector<std::wstring>& cultures, std::wstring_view culture)
{
    std::wstring current = FoldCulture(culture);
    while (! current.empty())
    {
        if (! IsEnglishCulture(current) && std::find(cultures.begin(), cultures.end(), current) == cultures.end())
        {
            cultures.push_back(current);
        }

        const size_t dash = current.rfind(L'-');
        if (dash == std::wstring::npos)
        {
            break;
        }
        current.resize(dash);
    }
}

[[nodiscard]] std::vector<std::wstring> ResolveCultureChain(const LanguagePreference& preference)
{
    std::vector<std::wstring> cultures;
    if (preference.kind == LanguagePreferenceKind::Culture)
    {
        AddCultureWithParents(cultures, preference.culture);
        return cultures;
    }

    ULONG count        = 0;
    ULONG bufferLength = 0;
    if (GetUserPreferredUILanguages(MUI_LANGUAGE_NAME, &count, nullptr, &bufferLength) == FALSE || bufferLength == 0)
    {
        return cultures;
    }

    std::wstring buffer(bufferLength, L'\0');
    if (GetUserPreferredUILanguages(MUI_LANGUAGE_NAME, &count, buffer.data(), &bufferLength) == FALSE)
    {
        return cultures;
    }

    PCWSTR cursor = buffer.c_str();
    for (ULONG index = 0; index < count && cursor && *cursor != L'\0'; ++index)
    {
        const std::wstring_view culture(cursor);
        AddCultureWithParents(cultures, culture);
        cursor += culture.size() + 1u;
    }
    return cultures;
}

[[nodiscard]] wil::unique_hmodule TryLoadSatelliteForCulture(const ResourceOwner& owner,
                                                             std::wstring_view culture,
                                                             std::unordered_set<std::wstring>& warnedLoads) noexcept
{
    const std::filesystem::path moduleDirectory = GetModuleDirectory(owner.embeddedInstance);
    if (moduleDirectory.empty())
    {
        return {};
    }

    const std::wstring satelliteName = std::format(L"{}-{}.dll", owner.moduleName, culture);

    std::vector<std::filesystem::path> candidateDirectories;
    candidateDirectories.push_back(moduleDirectory);
    const std::filesystem::path hostDirectory = GetModuleDirectory(GetModuleHandleW(nullptr));
    if (! hostDirectory.empty() && hostDirectory != moduleDirectory)
    {
        candidateDirectories.push_back(hostDirectory);
    }

    std::filesystem::path firstSatellitePath;
    for (const std::filesystem::path& candidateDirectory : candidateDirectories)
    {
        const std::filesystem::path satellitePath = candidateDirectory / L"Lang" / satelliteName;
        if (firstSatellitePath.empty())
        {
            firstSatellitePath = satellitePath;
        }

        wil::unique_hmodule satellite(LoadLibraryExW(satellitePath.c_str(), nullptr, LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE));
        if (satellite)
        {
            return satellite;
        }
    }

    const std::wstring warnKey = owner.moduleName + L"|" + std::wstring(culture);
    if (warnedLoads.insert(warnKey).second)
    {
        Debug::Warning(L"Localization satellite '{}' could not be loaded; embedded English resources will be used.", firstSatellitePath.c_str());
    }
    return {};
}

void ReloadOwnerSatellite(ResourceOwner& owner, const LanguagePreference& preference, std::unordered_set<std::wstring>& warnedLoads) noexcept
{
    owner.satellite.reset();
    owner.loadedCulture.clear();

    for (const std::wstring& culture : ResolveCultureChain(preference))
    {
        wil::unique_hmodule satellite = TryLoadSatelliteForCulture(owner, culture, warnedLoads);
        if (satellite)
        {
            owner.loadedCulture = culture;
            owner.satellite     = std::move(satellite);
            return;
        }
    }
}

[[nodiscard]] int LoadStringFromInstance(HINSTANCE instance, UINT id, std::wstring& result) noexcept
{
    PCWSTR ptr       = nullptr;
    const int length = ::LoadStringW(instance, id, reinterpret_cast<LPWSTR>(&ptr), 0);
    if (length <= 0 || ! ptr)
    {
        result.clear();
        return 0;
    }

    result.assign(ptr, static_cast<size_t>(length));
    return length;
}

[[nodiscard]] ResourceOwner* FindOwnerLocked(LocalizationState& state, HINSTANCE embeddedInstance) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    const auto it            = state.ownersByInstance.find(instance);
    return (it == state.ownersByInstance.end()) ? nullptr : &it->second;
}
} // namespace

HRESULT RegisterResourceOwner(std::wstring_view ownerName, HINSTANCE embeddedInstance) noexcept
{
    if (ownerName.empty())
    {
        return E_INVALIDARG;
    }

    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance)
    {
        return E_HANDLE;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);

    if (const auto existing = state.ownersByInstance.find(instance); existing != state.ownersByInstance.end())
    {
        if (! OrdinalString::EqualsNoCase(existing->second.moduleName, ownerName))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if (existing->second.registrationCount == (std::numeric_limits<size_t>::max)())
        {
            return HRESULT_FROM_WIN32(ERROR_TOO_MANY_NAMES);
        }

        ++existing->second.registrationCount;
        return S_OK;
    }

    ResourceOwner owner;
    owner.moduleName       = std::wstring(ownerName);
    owner.embeddedInstance = instance;
    ReloadOwnerSatellite(owner, state.preference, state.warnedLoads);
    state.ownersByInstance[instance] = std::move(owner);
    return S_OK;
}

void UnregisterResourceOwner(HINSTANCE embeddedInstance) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance)
    {
        return;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    const auto existing = state.ownersByInstance.find(instance);
    if (existing == state.ownersByInstance.end())
    {
        return;
    }

    if (existing->second.registrationCount > 1u)
    {
        --existing->second.registrationCount;
        return;
    }

    state.ownersByInstance.erase(existing);
}

HRESULT ApplyLanguagePreference(const LanguagePreference& preference) noexcept
{
    if (preference.kind == LanguagePreferenceKind::Culture && preference.culture.empty())
    {
        return E_INVALIDARG;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    state.preference = preference;
    for (auto& [_, owner] : state.ownersByInstance)
    {
        ReloadOwnerSatellite(owner, state.preference, state.warnedLoads);
    }
    return S_OK;
}

int LoadString(HINSTANCE embeddedInstance, UINT id, std::wstring& result) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance)
    {
        result.clear();
        return 0;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    if (ResourceOwner* owner = FindOwnerLocked(state, instance); owner && owner->satellite)
    {
        const int satelliteLength = LoadStringFromInstance(owner->satellite.get(), id, result);
        if (satelliteLength > 0)
        {
            return satelliteLength;
        }
    }

    return LoadStringFromInstance(instance, id, result);
}

HMENU LoadMenuResource(HINSTANCE embeddedInstance, UINT menuId) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance)
    {
        return nullptr;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    if (ResourceOwner* owner = FindOwnerLocked(state, instance); owner && owner->satellite)
    {
        if (HMENU menu = LoadMenuW(owner->satellite.get(), MAKEINTRESOURCEW(menuId)))
        {
            return menu;
        }
    }

    return LoadMenuW(instance, MAKEINTRESOURCEW(menuId));
}

HACCEL LoadAcceleratorsResource(HINSTANCE embeddedInstance, PCWSTR tableName) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance || ! tableName)
    {
        return nullptr;
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    if (ResourceOwner* owner = FindOwnerLocked(state, instance); owner && owner->satellite)
    {
        if (HACCEL accel = LoadAcceleratorsW(owner->satellite.get(), tableName))
        {
            return accel;
        }
    }

    return LoadAcceleratorsW(instance, tableName);
}

ResourceLookupResult FindLocalizedResourceHandle(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept
{
    const HINSTANCE instance = NormalizeInstance(embeddedInstance);
    if (! instance || ! name || ! type)
    {
        return {};
    }

    auto& state = State();
    std::scoped_lock lock(state.mutex);
    if (ResourceOwner* owner = FindOwnerLocked(state, instance); owner && owner->satellite)
    {
        if (HRSRC resource = FindResourceW(owner->satellite.get(), name, type))
        {
            return {owner->satellite.get(), resource};
        }
    }

    if (HRSRC resource = FindResourceW(instance, name, type))
    {
        return {instance, resource};
    }
    return {};
}

HRSRC FindLocalizedResource(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept
{
    return FindLocalizedResourceHandle(embeddedInstance, name, type).resource;
}

HINSTANCE ResolveResourceInstance(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type) noexcept
{
    if (const ResourceLookupResult found = FindLocalizedResourceHandle(embeddedInstance, name, type))
    {
        return found.instance;
    }
    return nullptr;
}

HANDLE LoadImageResource(HINSTANCE embeddedInstance, PCWSTR name, UINT type, int cx, int cy, UINT flags) noexcept
{
    PCWSTR resourceType = nullptr;
    switch (type)
    {
        case IMAGE_BITMAP: resourceType = RT_BITMAP; break;
        case IMAGE_CURSOR: resourceType = RT_GROUP_CURSOR; break;
        case IMAGE_ICON: resourceType = RT_GROUP_ICON; break;
        default: break;
    }

    const HINSTANCE embedded = NormalizeInstance(embeddedInstance);
    const HINSTANCE instance = resourceType ? ResolveResourceInstance(embedded, name, resourceType) : embedded;
    if (! instance)
    {
        return nullptr;
    }

    if (HANDLE image = LoadImageW(instance, name, type, cx, cy, flags))
    {
        return image;
    }

    return instance != embedded && embedded ? LoadImageW(embedded, name, type, cx, cy, flags) : nullptr;
}

bool LoadResourceBytes(HINSTANCE embeddedInstance, PCWSTR name, PCWSTR type, std::vector<std::byte>& result) noexcept
{
    result.clear();
    auto tryLoad = [&result](const ResourceLookupResult found) noexcept -> bool
    {
        if (! found)
        {
            return false;
        }

        const DWORD size = SizeofResource(found.instance, found.resource);
        if (size == 0)
        {
            return false;
        }

        HGLOBAL loaded = LoadResource(found.instance, found.resource);
        if (! loaded)
        {
            return false;
        }

        const void* bytes = LockResource(loaded);
        if (! bytes)
        {
            return false;
        }

        const auto* first = static_cast<const std::byte*>(bytes);
        result.assign(first, first + size);
        return true;
    };

    const ResourceLookupResult found = FindLocalizedResourceHandle(embeddedInstance, name, type);
    if (tryLoad(found))
    {
        return true;
    }

    const HINSTANCE embedded = NormalizeInstance(embeddedInstance);
    if (embedded && found.instance != embedded)
    {
        if (HRSRC embeddedResource = FindResourceW(embedded, name, type))
        {
            return tryLoad({embedded, embeddedResource});
        }
    }
    return false;
}

std::vector<std::wstring> DiscoverAvailableCultures() noexcept
{
    std::vector<std::wstring> cultures;
    const std::filesystem::path langDirectory = GetModuleDirectory(GetModuleHandleW(nullptr)) / L"Lang";

    std::error_code ec;
    if (! std::filesystem::is_directory(langDirectory, ec))
    {
        return cultures;
    }

    for (std::filesystem::directory_iterator it(langDirectory, ec), end; ! ec && it != end; it.increment(ec))
    {
        if (! it->is_regular_file(ec))
        {
            ec.clear();
            continue;
        }

        const std::wstring stem = it->path().stem().wstring();
        const size_t dash       = stem.rfind(L'-');
        if (dash == std::wstring::npos || dash + 1u >= stem.size())
        {
            continue;
        }

        std::wstring culture      = stem.substr(dash + 1u);
        const size_t previousDash = stem.rfind(L'-', dash - 1u);
        if (previousDash != std::wstring::npos && dash - previousDash <= 3u)
        {
            culture = stem.substr(previousDash + 1u);
        }

        if (std::find(cultures.begin(), cultures.end(), culture) == cultures.end())
        {
            cultures.push_back(std::move(culture));
        }
    }

    std::sort(cultures.begin(), cultures.end());
    return cultures;
}
} // namespace Localization
