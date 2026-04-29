# Callback Teardown Guard Pattern

## Problem
Async callbacks (threadpool, I/O completion) can fire after the host object that registered them has begun teardown. Checking a `_stopping` flag is insufficient when it's set concurrently with cleanup — there's a window where the callback passes the check but the host is already partially destroyed.

## Pattern

```cpp
// 1. Define guard (non-copyable because of std::atomic)
struct CallbackGuard
{
    CallbackGuard()                              = default;
    CallbackGuard(const CallbackGuard&)            = delete;
    CallbackGuard(CallbackGuard&&)                 = delete;
    CallbackGuard& operator=(const CallbackGuard&) = delete;
    CallbackGuard& operator=(CallbackGuard&&)      = delete;

    std::atomic<bool> valid{true};
};

// 2. Owner holds shared_ptr
std::shared_ptr<CallbackGuard> _callbackGuard = std::make_shared<CallbackGuard>();

// 3. On teardown — invalidate FIRST, before any other cleanup
void Stop() noexcept
{
    _callbackGuard->valid.store(false, std::memory_order_release);
    _stopping.store(true, std::memory_order_release);
    // ... CancelIo, WaitForThreadpoolCallbacks, resource cleanup ...
}

// 4. In every callback — check guard before invoking host
void NotifyHost() noexcept
{
    if (_stopping.load(std::memory_order_acquire)) return;
    if (!_callbackGuard->valid.load(std::memory_order_acquire)) return;
    if (!_callback) return;

    _callback->OnEvent(&data, _cookie);
}

// 5. On restart — re-enable the guard
void Start() noexcept
{
    _callbackGuard->valid.store(true, std::memory_order_release);
    // ...
}
```

## Memory Ordering
- `store(false, release)` in Stop() → `load(acquire)` in callback
- This forms a happens-before relationship: if the callback sees `valid == false`, all prior writes in Stop() are visible

## Key Rules
- Guard invalidation must be the **first** operation in teardown
- Both `_stopping` and `_callbackGuard` should be checked (defense in depth)
- Explicit deleted copy/move satisfies MSVC /W4 warnings (C4625, C4626, C5026, C5027)
- Use `std::make_shared` in constructor (tiny allocation, effectively can't fail)

## When to Use
- Any plugin with async callbacks that could outlive the host reference
- Threadpool I/O callbacks (ReadDirectoryChangesW, etc.)
- Background worker threads invoking registered callbacks
- Any scenario where `SetCallback(nullptr)` must be a hard barrier

## Reference Implementation
- `Plugins/FileSystem/FileSystem.Watch.cpp` — DirectoryWatch class
