---
name: plugin-callbacks
description: Plugin callback patterns with cookie-based context for INavigationMenuCallback, IFileSystemCallback, and IFileSystemSearchCallback interfaces. Use when implementing or consuming plugin interfaces.
metadata:
  author: DualTail
  version: "1.0"
---

# Plugin Callback Pattern

## Key Principles

Plugin callbacks are **raw vtable interfaces** (NOT COM):
- Do NOT inherit from `IUnknown`
- Do NOT use `AddRef`/`Release`
- Do NOT store in `wil::com_ptr`

## Cookie Pattern

### Registration-Style (SetCallback)

```cpp
// Host registers callback
plugin->SetCallback(this, reinterpret_cast<void*>(static_cast<std::uintptr_t>(paneId)));

// Plugin stores weak reference
void SetCallback(ICallback* callback, void* cookie) 
{
    _callback = callback;  // Weak pointer - host owns lifetime
    _cookie = cookie;      // Pass back verbatim
}

// Plugin invokes with cookie
if (_callback) {
    _callback->OnEvent(_cookie, data);
}
```

### Per-Call Callbacks

```cpp
// Each operation passes callback + cookie
void Search(const wchar_t* query, ISearchCallback* callback, void* cookie) 
{
    // Forward same cookie to every callback
    callback->OnItemFound(item, cookie);
    callback->OnComplete(S_OK, cookie);
}
```

## Cleanup Requirements

1. **Host** must clear callbacks before unloading:
   ```cpp
   plugin->SetCallback(nullptr, nullptr);
   plugin.reset();
   ```

2. **Plugin** must stop calling after cleared:
   ```cpp
   void Cancel() 
   {
       _callback = nullptr;
       _cookie = nullptr;
   }
   ```

## Async / Unload Safety (Background Work)

If a plugin performs background work that may invoke callbacks later (threads, threadpool, async IO):

1. Prefer **cooperative cancellation** (`std::jthread` + `stop_token`) and **join/stop on unload**.
2. Avoid `std::thread(...).detach()` in plugin DLLs. Detached threads can outlive the plugin unload boundary and crash the process.
3. If detaching is truly unavoidable:
   - Keep the plugin module loaded for the worker lifetime (pin the DLL refcount via `AcquireModuleReferenceFromAddress(...)` from `Common/Helpers.h`).
   - Ensure the worker cannot run on a destroyed plugin instance (hold a strong lifetime reference for the worker and release it on exit).
   - Make thread start exception-safe in `noexcept` paths: if catching is required, `catch (...)` is FORBIDDEN; catch only explicitly named exception types and add a short comment explaining why catching is mandatory at that boundary.

### Preferred Pattern: Threadpool Callback + Module Pin

```cpp
// In the plugin DLL:
static const int kMyPluginModuleAnchor = 0;

struct WorkCtx
{
    wil::unique_hmodule moduleKeepAlive;
    HWND hwnd = nullptr;
};

void StartWorkAsync(HWND hwnd)
{
    auto ctx = std::make_unique<WorkCtx>();
    ctx->hwnd = hwnd;
    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kMyPluginModuleAnchor);

    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE /*instance*/, void* context) noexcept
        {
            std::unique_ptr<WorkCtx> ctx(static_cast<WorkCtx*>(context));
            static_cast<void>(ctx->moduleKeepAlive); // Keep DLL loaded for callback lifetime

            // Do background work, then post back to UI thread using PostMessagePayload(...)
        },
        ctx.get(),
        nullptr);

    if (queued != 0)
    {
        ctx.release(); // Ownership transfers to the callback
    }
}
```

**Rule of thumb:** unloading must be a “quiet point” where no background worker can still call a cleared callback or touch freed memory.

## “Quiet Point” Shutdown / Unload Ordering

Define a **quiet point** as the moment after which:
- no background worker can touch plugin instance memory,
- no worker can call a cleared callback,
- no worker can run code in an unloaded DLL,
- no in-flight posted payload depends on an `HWND` that may be destroyed.

### Host-side ordering (recommended)
1. **Stop producers first**: stop timers / stop scheduling new background work / stop posting new payload messages.
2. **Request plugin shutdown**: call `Close()` / cancel APIs on the plugin object(s).
3. **Clear callbacks**: `SetCallback(nullptr, nullptr)` (do this before releasing/unloading).
4. **Release instances**: drop the last `wil::com_ptr` to the plugin COM objects before unloading the module.

### Plugin-side requirements
- Treat `SetCallback(nullptr, nullptr)` as a hard barrier: after it returns, **no thread may call** the previous callback pointer.
- On `Close()` / `Cancel()` / window `WM_DESTROY`:
  - signal cancellation (stop tokens/atomics),
  - ensure any threadpool callbacks cannot outlive the instance (wait/cancel in a non-UI context, or keep a strong self lifetime for the callback),
  - if scheduling threadpool callbacks that may outlive host references, **pin the module** with `AcquireModuleReferenceFromAddress(...)`.
- **Posted payloads:** `PostMessagePayload(...)` only reclaims payloads when `PostMessageW` fails; it does not protect against an `HWND` being destroyed before the message is processed. Quiet-point shutdown must therefore stop producers and/or route payloads to a dispatcher that outlives view windows.

## Interface Example

```cpp
// NOT COM - no IUnknown inheritance
class IFileSystemCallback 
{
public:
    virtual void OnItemFound(const FileInfo& info, void* cookie) = 0;
    virtual void OnComplete(HRESULT result, void* cookie) = 0;
    virtual void OnProgress(int percent, void* cookie) = 0;
};
```

## Cookie Usage

The cookie disambiguates context (e.g., which pane triggered the operation):

```cpp
void MyHost::OnItemFound(const FileInfo& info, void* cookie) 
{
    auto paneId = static_cast<PaneId>(reinterpret_cast<std::uintptr_t>(cookie));
    GetPane(paneId)->AddItem(info);
}
```
