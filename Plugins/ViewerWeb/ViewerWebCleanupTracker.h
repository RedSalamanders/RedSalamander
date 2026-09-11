#pragma once

#include <algorithm>
#include <cstddef>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace ViewerWebSecurity
{
class StagedCleanupTracker final
{
public:
    using CleanupOperation = bool (*)(void* context, std::wstring_view path) noexcept;

    struct Operations
    {
        void* context = nullptr;
        CleanupOperation deleteNow = nullptr;
        CleanupOperation scheduleLater = nullptr;
    };

    StagedCleanupTracker()                                      = default;
    StagedCleanupTracker(const StagedCleanupTracker&)            = delete;
    StagedCleanupTracker(StagedCleanupTracker&&)                 = delete;
    StagedCleanupTracker& operator=(const StagedCleanupTracker&) = delete;
    StagedCleanupTracker& operator=(StagedCleanupTracker&&)      = delete;

    void Track(std::wstring path)
    {
        if (path.empty())
        {
            return;
        }

        std::scoped_lock lock(_mutex);
        if (std::ranges::find(_pendingPaths, path) == _pendingPaths.end())
        {
            _pendingPaths.push_back(std::move(path));
        }
    }

    void Retry(const Operations& operations)
    {
        std::vector<std::wstring> current;
        {
            std::scoped_lock lock(_mutex);
            current.swap(_pendingPaths);
        }

        std::vector<std::wstring> failed;
        failed.reserve(current.size());
        for (std::wstring& path : current)
        {
            const bool deleted = operations.deleteNow && operations.deleteNow(operations.context, path);
            const bool scheduled = ! deleted && operations.scheduleLater && operations.scheduleLater(operations.context, path);
            if (! deleted && ! scheduled)
            {
                failed.push_back(std::move(path));
            }
        }

        if (! failed.empty())
        {
            std::scoped_lock lock(_mutex);
            for (std::wstring& path : failed)
            {
                if (std::ranges::find(_pendingPaths, path) == _pendingPaths.end())
                {
                    _pendingPaths.push_back(std::move(path));
                }
            }
        }
    }

    [[nodiscard]] size_t PendingCount() const noexcept
    {
        std::scoped_lock lock(_mutex);
        return _pendingPaths.size();
    }

private:
    mutable std::mutex _mutex;
    std::vector<std::wstring> _pendingPaths;
};
} // namespace ViewerWebSecurity
