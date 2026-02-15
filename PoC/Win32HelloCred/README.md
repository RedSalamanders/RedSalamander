# Win32HelloCred

A minimal **Win32 desktop** C++ sample that:

1. Saves a secret using **Windows Credential Manager** (`CredWriteW` / `CredReadW`).
2. Requires **Windows Hello/PIN** re-verification before reading the stored secret, using
   `IUserConsentVerifierInterop::RequestVerificationForWindowAsync` (WinRT interop).

## Build requirements

- Visual Studio 2019/2022 (Desktop development with C++)
- Windows 10/11 SDK installed (for C++/WinRT headers and `userconsentverifierinterop.h`)

## How to run

1. Build and run.
2. Enter:
   - Target: a unique name, e.g. `MyApp:demo` or `MyApp:service:alice`
   - Username: optional
   - Secret: password/token
3. Click **Save**.
4. Click **Load (requires Hello)** to trigger Windows Hello verification and then read the credential.

## Notes

- This is a demonstration. In real apps:
  - Prefer recallable tokens over passwords.
  - Do not display secrets in UI; this sample does only to prove the flow.
  - Minimize secret lifetime in memory and avoid copying into immutable strings where possible.
