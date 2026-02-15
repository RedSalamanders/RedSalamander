---
name: async-threading
description: Threading model and async operation patterns for Windows applications. Use when implementing background operations, thread synchronization, posting to UI thread, or using thread pools.
metadata:
  author: DualTail
  version: "1.0"
---

# Threading and Async Operations

## Threading Model

- **UI operations** → Main thread only
- **Heavy computations** → Worker threads
- **Never block** UI thread with synchronous waits

## Synchronization with WIL

```cpp
wil::srwlock _lock;

// Exclusive access
void WriteOperation() 
{
    auto guard = _lock.lock_exclusive();
    // Protected write
}

// Shared read access
void ReadOperation() 
{
    auto guard = _lock.lock_shared();
    // Read-only access
}
```

## Event Signaling

```cpp
wil::unique_event _completionEvent;
_completionEvent.create();

// Worker signals
_completionEvent.SetEvent();

// Wait with timeout
if (_completionEvent.wait(100) == WAIT_OBJECT_0) 
{
    // Completed
}
```

## Post to UI Thread

```cpp
constexpr UINT WM_APP_WORK = WM_APP + 1;

struct UiWorkItem
{
    int kind;
    std::wstring text;
};

void PostToUIThread(UiWorkItem item)
{
    auto payload = std::make_unique<UiWorkItem>(std::move(item));
    // `PostMessagePayload` reclaims the payload automatically if PostMessageW fails.
    static_cast<void>(PostMessagePayload(_hWnd, WM_APP_WORK, 0, std::move(payload)));
}

// In WndProc
case WM_APP_WORK: 
{
    auto payload = TakeMessagePayload<UiWorkItem>(lParam);
    if (! payload)
    {
        return 0;
    }
    // Use payload->kind/payload->text to update UI
    return 0;
} break;
```

### `HWND` Teardown and Posted Payloads

If an `HWND` can be destroyed while payload messages are still queued, Windows may discard those messages without delivering them. To prevent payload leaks (and to keep the `PostMessagePayload` registry consistent):

- Call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`).
- Call `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`.
- Always receive with `TakeMessagePayload<T>(lParam)` (not a manual `unique_ptr` wrap).

## Async Icon Pattern

```cpp
// 1. Background work: extract (use threadpool; avoid std::thread::detach in plugins)
constexpr UINT WM_APP_ICON_READY = WM_APP + 2;

struct IconReadyPayload
{
    int idx = 0;
    wil::unique_hicon icon;
};

struct IconWorkItem
{
    HWND hwnd = nullptr;
    int idx   = 0;
};

void RequestIconAsync(int idx)
{
    const HWND hwnd = _hWnd;
    auto ctx  = std::make_unique<IconWorkItem>();
    ctx->hwnd = hwnd;
    ctx->idx  = idx;

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
         {
             std::unique_ptr<IconWorkItem> ctx(static_cast<IconWorkItem*>(context));
             if (! ctx || ! ctx->hwnd)
             {
                 return;
             }

             [[maybe_unused]] auto coInit = wil::CoInitializeEx(COINIT_MULTITHREADED);
             wil::unique_hicon icon = IconCache::GetInstance().ExtractSystemIcon(ctx->idx);
             auto payload           = std::make_unique<IconReadyPayload>();
             payload->idx           = ctx->idx;
             payload->icon          = std::move(icon);

             static_cast<void>(PostMessagePayload(ctx->hwnd, WM_APP_ICON_READY, 0, std::move(payload)));
        },
        ctx.get(),
        nullptr);

    if (queued != 0)
    {
        ctx.release(); // Ownership transfers to the callback
    }
}

// In WndProc
case WM_APP_ICON_READY:
{
    auto payload = TakeMessagePayload<IconReadyPayload>(lParam);
    if (! payload)
    {
        return 0;
    }

    const int idx = payload->idx;
    auto bitmap   = IconCache::ConvertIconToBitmapOnUIThread(std::move(payload->icon));
    // Cache bitmap + invalidate
    return 0;
} break;
```

## Plugin Unload Safety (DLLs)

When scheduling background work from a **plugin DLL** (viewers/file systems):

1. Avoid `std::thread(...).detach()`.
2. Prefer `TrySubmitThreadpoolCallback` or `std::jthread` tracked on the instance and stopped/joined on unload.
3. If a callback can outlive the host’s last reference to the DLL, **pin the module** for the callback lifetime:
   - add a static anchor in the module: `static const int kMyPluginModuleAnchor = 0;`
   - store `wil::unique_hmodule moduleKeepAlive = AcquireModuleReferenceFromAddress(&kMyPluginModuleAnchor);` in the callback context
   - touch it in the callback: `static_cast<void>(ctx->moduleKeepAlive);`

## Best Practices

1. Use message queues for cross-thread communication
2. Prefer lock-free patterns where possible
3. Handle thread cancellation gracefully
4. Profile multi-threaded code paths
5. Implement cleanup/cancel mechanisms
6. `catch (...)` is FORBIDDEN; if catching is mandatory at a thread/callback boundary, catch only explicitly named exception types and add a short comment explaining why.

## Resource Management

### Caching
- Implement LRU caches for expensive resources
- Use weak references to break cycles
- Monitor resource usage and implement limits

### Thread Pools
```cpp
// Keep the work object alive until callbacks complete, and close it via RAII.
wil::unique_any<PTP_WORK, decltype(&::CloseThreadpoolWork), ::CloseThreadpoolWork> _work;

void StartWork(void* context)
{
    _work.reset(CreateThreadpoolWork(&WorkCallback, context, nullptr));
    SubmitThreadpoolWork(_work.get());
}

void CancelWork()
{
    if (_work)
    {
        WaitForThreadpoolWorkCallbacks(_work.get(), /*cancelPendingCallbacks*/ TRUE);
        _work.reset(); // CloseThreadpoolWork
    }
}
```

## Architecture Patterns

### Component Design
- Single responsibility principle
- Composition over inheritance
- Proper separation of concerns
- Design for testability

### Device Loss Handling
- Implement proper D2D device loss handling
- Cache expensive resources with LRU eviction
- Use weak references to break cycles
