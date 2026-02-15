---
name: win32-wndproc
description: Win32 window procedure and message handling patterns. Use when implementing WndProc, handling Windows messages like WM_PAINT WM_SIZE WM_NOTIFY WM_COMMAND, or creating message handler methods.
metadata:
  author: DualTail
  version: "1.0"
---

# Win32 Message Handling

## Mandatory Guidelines

- Keep switch cases **≤ 6 lines**
- Route messages to dedicated `On*` handlers
- Cast `WPARAM`/`LPARAM` in WndProc before calling handlers
- Put all logic in handler methods
- Custom `WM_APP` / `WM_USER` message IDs must be declared in `Common/WindowMessages.h` (namespace `WndMsg`)
- `catch (...)` is FORBIDDEN; if catching is mandatory at a callback boundary, catch only explicitly named exception types and add a short comment explaining why.

## Window Procedure Pattern

```cpp
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    auto* self = reinterpret_cast<MyWindow*>(GetWindowLongPtr(hWnd, GWLP_USERDATA));
    
    switch (message) {
    case WM_CREATE:
        return self->OnCreate(reinterpret_cast<CREATESTRUCT*>(lParam));
    case WM_PAINT:
        return self->OnPaint();
    case WM_SIZE:
        return self->OnSize(LOWORD(lParam), HIWORD(lParam));
    case WM_NOTIFY:
        return self->OnNotify(static_cast<int>(wParam), 
                              reinterpret_cast<NMHDR*>(lParam));
    case WM_COMMAND:
        return self->OnCommand(LOWORD(wParam), HIWORD(wParam), 
                               reinterpret_cast<HWND>(lParam));
    case WM_DESTROY:
        return self->OnDestroy();
    }
    return DefWindowProc(hWnd, message, wParam, lParam);
}
```

## Handler Signatures

```cpp
LRESULT OnCreate(CREATESTRUCT* cs);
LRESULT OnPaint();
LRESULT OnSize(int width, int height);
LRESULT OnNotify(int controlId, NMHDR* nmhdr);
LRESULT OnCommand(WORD commandId, WORD notifyCode, HWND controlHwnd);
LRESULT OnDestroy();
LRESULT OnDpiChanged(WORD dpi, RECT* newRect);
LRESULT OnMouseMove(int x, int y, WPARAM keys);
```

## OnPaint with RAII

```cpp
LRESULT MyWindow::OnPaint() {
    PAINTSTRUCT ps;
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd, &ps);
    FillRect(hdc.get(), &ps.rcPaint, _backgroundBrush.get());
    return 0; // Automatic EndPaint
}
```

## Closing Owned Windows (`wil::unique_hwnd`)

If your window handle is owned by `wil::unique_hwnd`, destroy via the wrapper:

```cpp
// ✅ Correct
_hWnd.reset();

// ❌ Wrong (double-destroy risk)
DestroyWindow(_hWnd.get());
```

## Message Payload Ownership (`PostMessageW`)

When posting heap payloads to a window, prefer the `PostMessagePayload(...)` / `TakeMessagePayload<T>(lParam)` helpers (see `Common/Helpers.h`) to avoid leak-on-failure and standardize ownership transfer.

```cpp
// Sender
// (Declare your message ID in `Common/WindowMessages.h` as `WndMsg::k...`)
auto payload = std::make_unique<MyPayload>();
static_cast<void>(PostMessagePayload(hwnd, WndMsg::kMyPayloadMessage, 0, std::move(payload)));

// Receiver (WndProc)
case WndMsg::kMyPayloadMessage:
{
    auto payload = TakeMessagePayload<MyPayload>(lParam);
    if (! payload) { return 0; }
    // ... use payload ...
    return 0;
}
```

### `WM_NCDESTROY` Drain Policy (Leak-on-Destroy)

If the target `HWND` can be destroyed while payload messages are still queued, Windows may discard those messages without delivering them. To prevent leaks:

- Call `InitPostedPayloadWindow(hwnd)` during window creation (`WM_NCCREATE`/`WM_CREATE`) for windows that receive payload messages.
- Call `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`.
- Always use `TakeMessagePayload<T>(lParam)` in the receiver so the registry can unregister; do not manually wrap `lParam` into a `std::unique_ptr`.

## OnNotify Pattern

```cpp
LRESULT OnNotify(int controlId, NMHDR* nmhdr) {
    switch (nmhdr->code) {
    case NM_CLICK:
        return OnItemClick(reinterpret_cast<NMITEMACTIVATE*>(nmhdr));
    case LVN_GETDISPINFO:
        return OnGetDispInfo(reinterpret_cast<NMLVDISPINFO*>(nmhdr));
    }
    return 0;
}
```
