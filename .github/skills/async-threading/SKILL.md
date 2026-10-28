---
name: async-threading
description: Threading model and async operation patterns for Windows applications. Use when implementing background operations, thread synchronization, posting to UI thread, or using thread pools.
metadata:
  author: RedSalamander
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
    auto payload = TakeMessagePayload<UiWorkItem>(hwnd, message, lParam);
    if (! payload)
    {
        return 0;
    }
    // Use payload->kind/payload->text to update UI
    return 0;
} break;
```

### `HWND` Teardown and Posted Payloads

If an `HWND` can be destroyed while payload messages are still queued, Windows may discard those messages without
delivering them. `PostMessagePayload` places an opaque registry token—not a payload address—in nonzero `lParam`.
The registry owns the payload until a receiver takes it or teardown drains it:

- Call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`).
- Call `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`; it fences new posts, invalidates all tokens for the
  window, and deletes their storage exactly once.
- Always receive with `TakeMessagePayload<T>(hwnd, message, lParam)` (not a cast or manual `unique_ptr` wrap). A null result from a
  nonzero token means stale, drained, unregistered, or type-mismatched input and must be ignored before
  message-specific work.
- Tokens use cryptographic entropy and live-entry collision rejection, with bounded retries and no predictable fallback. Validate the actual receiving HWND, message and type before removal; mismatches preserve registry ownership. A drained entry is invalidated; random token reuse has negligible collision probability, not a mathematical never-reuse guarantee.
- Reserve `lParam == 0` only for a payload-less fallback explicitly defined by that message.

### UI-host cross-thread checklist

For every UI-host change that posts work across threads, review and cover:

- Payload ownership: every heap payload is posted with `PostMessagePayload(...)`, received with
  `TakeMessagePayload<T>(hwnd, message, lParam)`, and ignored when a nonzero token no longer resolves.
- UI-thread boundary: UI state is mutated only from the owning UI thread or explicitly marshaled there.
- Cancellation path: producers stop or observe cancellation before teardown begins.
- Teardown drain: receiving windows initialize/drain the posted-payload registry.
- Regression guard: add a focused teardown stress/selftest when the touched host can queue payloads.

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
    auto payload = TakeMessagePayload<IconReadyPayload>(hwnd, message, lParam);
    if (! payload)
    {
        return 0;
    }

    const int idx = payload->idx;
    auto bitmap   = IconCache::GetInstance().ConvertIconToBitmapOnUIThread(payload->icon.get(), idx, _d2dContext.get());
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
   - if submission fails, let the context destroy the pin normally on the submitting thread
   - at the start of a threadpool callback, transfer the owning pin with
     `TransferModulePinToCallbackReturn(instance, ctx->moduleKeepAlive)`

Do not let the callback context destroy its last `wil::unique_hmodule` while the callback body is returning. `FreeLibrary` can otherwise unmap the plugin's callback code before Windows crosses the callback return boundary. This transfer rule is specific to Windows threadpool callbacks; other callback/thread systems must retain their documented module ownership until their own quiet point.

## Callback teardown

Atomic stopping/valid flags suppress later work; they are not a lifetime lease. A callback can pass
its checks just before teardown starts. Keep the instance, callback target and resources alive until
actual callback quiescence, and synchronize callback/cookie registration with notification under the
owning interface contract. A shared guard alone does not protect a raw host pointer.

Review `Plugins/FileSystem/FileSystem.Watch.cpp` (DirectoryWatch) before adapting the pattern: its
stop path invalidates the guard, cancels production, and waits for threadpool I/O/work callbacks before
releasing their resources. Use the appropriate non-UI join/drain/barrier for the touched host. Do not
replace that barrier with acquire/release ordering or re-enable a guard while old callbacks can run.

## Scheduling and resource review

Reuse the existing synchronized ownership model. Introduce lock-free code only with a demonstrated need and
reviewed memory-ordering proof. Bound queues, coalesce stale UI work and measure the protected scenario under
the repository performance contract instead of prescribing a new cache or thread for every component.

Check threadpool allocation/submission before use; keep work/context ownership until cancellation and callback
quiescence are complete. Wait on an appropriate non-UI owner and never wait for the current callback itself.
Apply the plugin module-pin rule above independently of work-object lifetime.

No exception crosses a Win32/threadpool/COM boundary. `catch (...)` is forbidden; catch only explicitly named
exception types where required, document the fallback, and preserve the repository's fatal allocation policy.
