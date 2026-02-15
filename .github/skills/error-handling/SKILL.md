---
name: error-handling
description: Error handling patterns for Windows APIs and HRESULT. Use when handling errors, logging with Debug class, implementing error propagation, or dealing with Win32 API failures.
metadata:
  author: DualTail
  version: "1.0"
---

# Error Handling

## HRESULT Pattern

Use HRESULT for all Windows API calls:

```cpp
HRESULT hr = SomeWindowsApi();
if (FAILED(hr)) 
{
    Debug::Error(L"SomeWindowsApi failed: {:#x}", hr);
    return hr;
}
```

## No `goto` Cleanup

`goto` is prohibited. Use early `return` and RAII (`wil::unique_*`, `wil::com_ptr`, `wil::scope_exit`) to handle cleanup and common exit paths.

```cpp
HRESULT DoWork() noexcept
{
    const bool wasBusy = _busy;
    _busy = true;
    auto restoreBusy = wil::scope_exit([&] { _busy = wasBusy; });

    HRESULT hr = Step1();
    if (FAILED(hr)) { return hr; }

    hr = Step2();
    if (FAILED(hr)) { return hr; }

    return S_OK;
}
```

## Debug Logging Functions

| Function | When to Use |
|----------|-------------|
| `Debug::ErrorWithLastError(...)` | Win32 API fails, propagate via `HRESULT_FROM_WIN32` |
| `Debug::Error(...)` | Unexpected failures that abort an operation |
| `Debug::Warning(...)` | Recoverable failures with fallback |
| `Debug::Info(...)` | Informational messages |

## Win32 API Failures

Capture `lastError` immediately after the API call:

```cpp
BOOL result = SomeWin32Api();
if (!result) 
{
    DWORD lastError = Debug::ErrorWithLastError(L"SomeWin32Api failed");
    return HRESULT_FROM_WIN32(lastError);
}
```

## What NOT to Log

Don't log normal control-flow:
- Window size 0 during minimize
- Cancellation/abort codes
- D2D/DXGI "recreate target" paths that are handled
- User-initiated cancellations

## Logging Best Practices

```cpp
// ✅ Pass format args directly
Debug::Error(L"Failed to open {}: {:#x}", filename, hr);

// ❌ Avoid std::format wrapper
Debug::Error(std::format(L"Failed: {}", msg).c_str());

// ❌ Avoid embedded newlines
Debug::Error(L"Line1\nLine2");  // BAD
```

## Error Propagation

Avoid "log at every layer" - log once with context:

```cpp
// ✅ Log at the point of failure with full context
HRESULT LoadFile(const std::wstring& path) 
{
    wil::unique_handle file(CreateFileW(path.c_str(), ...));
    if (! file) 
    {
        auto err = Debug::ErrorWithLastError(L"LoadFile: CreateFileW failed for {}", path);
        return HRESULT_FROM_WIN32(err);
    }
    // ...
}

// ❌ Don't log again in caller
HRESULT hr = LoadFile(path);
if (FAILED(hr)) 
{
    Debug::Error(L"LoadFile failed");  // Redundant!
    return hr;
}
```

## Exceptions (No `catch (...)`)

Prefer code that **does not throw** in normal control-flow:
- Use `HRESULT` / Win32 error codes for Windows APIs.
- Prefer `std::filesystem` overloads that take a `std::error_code&` instead of throwing.

`catch (...)` is **FORBIDDEN** in this codebase.

If exception handling is mandatory, only catch explicitly named exception types at **ABI / `noexcept` boundaries** where an escaping exception would terminate the process or cause UB:
- `noexcept` COM methods (`STDMETHODCALLTYPE ... noexcept`)
- Win32 callbacks / `WndProc` / threadpool callbacks
- Thread entrypoints (`std::jthread`, `TrySubmitThreadpoolCallback`, etc.)

If you need to translate exceptions into failure codes, prefer:
```cpp
HRESULT DoThing() noexcept
{
    // Mandatory: `noexcept` boundary. Catch known exception types and convert to HRESULT.
    try
    {
        // ... work that may allocate/throw ...
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal in this codebase. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"DoThing: std::exception");
        return E_FAIL;
    }
}
```

When catching exceptions:
- **Catch-and-fail:** log once with `Debug::Error(...)` and include enough context to diagnose.
- **`std::bad_alloc`:** prefer fail-fast (`std::terminate()`); avoid doing additional allocations/log formatting on the OOM path.
- **Avoid redundant micro-catches:** in a function that is already `noexcept`, you do not need a standalone `catch (const std::bad_alloc&) { std::terminate(); }` unless you also have a broader catch (e.g. `catch (const std::exception&)`) that would otherwise swallow OOM.
- **Catch-and-continue:** only if it’s truly best-effort; document the fallback (avoid empty catch without an explanation).
- Never log sensitive data (passwords/secrets).

## User-Facing Errors

- Fail gracefully with user-friendly messages
- Use resource strings (see localization skill)
- Provide actionable information when possible

```cpp
if (FAILED(hr)) 
{
    const UINT messageId = operation == FILESYSTEM_COPY ? static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_COPY) : static_cast<UINT>(IDS_FMT_FILEOPS_CONFIRM_MOVE);
    const std::wstring message = FormatStringResource(nullptr, messageId, what, fromText, toText);

    const std::wstring caption = LoadStringResource(nullptr, IDS_CAPTION_CONFIRM);
    HostPromptRequest prompt{};
    prompt.version = 1;
    prompt.sizeBytes = sizeof(prompt);
    prompt.scope = (owner && IsWindow(owner)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    prompt.severity = HOST_ALERT_WARNING;
    prompt.buttons = HOST_PROMPT_BUTTONS_OK_CANCEL;
    prompt.targetWindow = (prompt.scope == HOST_ALERT_SCOPE_WINDOW) ? owner : nullptr;
    prompt.title = caption.c_str();
    prompt.message = message.c_str();
    prompt.defaultResult = HOST_PROMPT_RESULT_OK;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT hr = HostShowPrompt(prompt, nullptr, &promptResult);
    if (FAILED(hr))
    {
        return false;
    }

    return promptResult == HOST_PROMPT_RESULT_OK;
}
```


