#pragma once

#include <cstdint>
#include <unknwn.h>
#include <wchar.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure

// Root host services object (extensible via QueryInterface).
// UUID: {c7191bad-276e-4f7b-91ec-4803315413a7}
interface __declspec(uuid("c7191bad-276e-4f7b-91ec-4803315413a7")) __declspec(novtable) IHost : public IUnknown{};

enum HostAlertScope : uint32_t
{
    HOST_ALERT_SCOPE_PANE_CONTENT = 1, // pane without navigation bar
    HOST_ALERT_SCOPE_PANE         = 2, // pane with navigation bar
    HOST_ALERT_SCOPE_APPLICATION  = 3, // application window
    HOST_ALERT_SCOPE_WINDOW       = 4, // specific HWND (request.targetWindow)
};

enum HostAlertModality : uint32_t
{
    HOST_ALERT_MODELESS = 1,
    HOST_ALERT_MODAL    = 2,
};

enum HostAlertSeverity : uint32_t
{
    HOST_ALERT_ERROR   = 1,
    HOST_ALERT_WARNING = 2,
    HOST_ALERT_INFO    = 3,
    HOST_ALERT_BUSY    = 4,
};

struct HostAlertRequest
{
    // ABI versioning.
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostAlertRequest)

    HostAlertScope scope;
    HostAlertModality modality;
    HostAlertSeverity severity;

    // Used only when scope == HOST_ALERT_SCOPE_WINDOW.
    HWND targetWindow;

    // Strings are UTF-16, NUL-terminated, caller-owned, and only valid for the duration of the call.
    const wchar_t* title;   // optional (nullptr/empty allowed)
    const wchar_t* message; // required for user-visible alerts

    // If FALSE, the host does not expose a close “X” and does not dismiss the alert on Esc.
    // Typical use: fatal errors that require navigation away, or busy states.
    BOOL closable;

    uint32_t reserved[8];
};

// UUID: {06da0f05-fe31-4273-9029-22037e3b1ea8}
interface __declspec(uuid("06da0f05-fe31-4273-9029-22037e3b1ea8")) __declspec(novtable) IHostAlerts : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE ShowAlert(const HostAlertRequest* request, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE ClearAlert(HostAlertScope scope, void* cookie) noexcept           = 0;
};

enum HostPromptButtons : uint32_t
{
    HOST_PROMPT_BUTTONS_OK            = 1,
    HOST_PROMPT_BUTTONS_OK_CANCEL     = 2,
    HOST_PROMPT_BUTTONS_YES_NO        = 3,
    HOST_PROMPT_BUTTONS_YES_NO_CANCEL = 4,
};

// Result values intentionally match Win32 MessageBox IDs (IDOK/IDCANCEL/IDYES/IDNO).
enum HostPromptResult : uint32_t
{
    HOST_PROMPT_RESULT_NONE   = 0,
    HOST_PROMPT_RESULT_OK     = 1, // IDOK
    HOST_PROMPT_RESULT_CANCEL = 2, // IDCANCEL
    HOST_PROMPT_RESULT_YES    = 6, // IDYES
    HOST_PROMPT_RESULT_NO     = 7, // IDNO
};

struct HostPromptRequest
{
    // ABI versioning.
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostPromptRequest)

    HostAlertScope scope;
    HostAlertSeverity severity;
    HostPromptButtons buttons;

    // Used only when scope == HOST_ALERT_SCOPE_WINDOW.
    HWND targetWindow;

    // Strings are UTF-16, NUL-terminated, caller-owned, and only valid for the duration of the call.
    const wchar_t* title;   // optional (nullptr/empty allowed)
    const wchar_t* message; // required

    // Optional: if set to a value that exists in `buttons`, the host SHOULD default-focus it.
    // Use HOST_PROMPT_RESULT_NONE to indicate no preference.
    HostPromptResult defaultResult;

    uint32_t reserved[8];
};

// UUID: {afb5a715-1110-41f3-b7bb-133d6ca735fd}
interface __declspec(uuid("afb5a715-1110-41f3-b7bb-133d6ca735fd")) __declspec(novtable) IHostPrompts : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE ShowPrompt(const HostPromptRequest* request, void* cookie, HostPromptResult* result) noexcept = 0;
};

enum HostConnectionSecretKind : uint32_t
{
    HOST_CONNECTION_SECRET_PASSWORD           = 1,
    HOST_CONNECTION_SECRET_SSH_KEY_PASSPHRASE = 2
};

struct HostConnectionManagerRequest
{
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostConnectionManagerRequest)

    // Optional filter: only show connections whose ConnectionProfile.pluginId matches this value.
    // nullptr/empty means "all connections".
    const wchar_t* filterPluginId;

    // Optional owner window for the dialog. If null, the host chooses an appropriate parent.
    HWND ownerWindow;

    uint32_t reserved[8];
};

struct HostConnectionManagerResult
{
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostConnectionManagerResult)

    // On S_OK, the host allocates a NUL-terminated UTF-16 string with CoTaskMemAlloc and stores it here.
    // Callers must free it with CoTaskMemFree().
    // On S_FALSE, this is set to nullptr.
    //
    // This is the unique user-visible connection name (ConnectionProfile.name).
    wchar_t* connectionName;

    uint32_t reserved[8];
};

// UUID: {018b09cf-dd4e-47ac-b013-baef06220cff}
interface __declspec(uuid("018b09cf-dd4e-47ac-b013-baef06220cff")) __declspec(novtable) IHostConnections : public IUnknown
{
    // Shows the Connection Manager dialog and returns the selected connection name.
    // Returns:
    // - S_OK: user selected a connection (result->connectionName is non-null)
    // - S_FALSE: user cancelled (result->connectionName is null)
    virtual HRESULT STDMETHODCALLTYPE ShowConnectionManager(const HostConnectionManagerRequest* request, HostConnectionManagerResult* result) noexcept = 0;

    // Returns a UTF-8 JSON object describing a saved connection (non-secret fields).
    // connectionName is the (case-insensitive) unique ConnectionProfile.name.
    // On success, the host allocates a NUL-terminated UTF-8 string with CoTaskMemAlloc and stores it in *jsonUtf8.
    // Callers must free it with CoTaskMemFree().
    virtual HRESULT STDMETHODCALLTYPE GetConnectionJsonUtf8(const wchar_t* connectionName, char** jsonUtf8) noexcept = 0;

    // Returns the requested secret (password/passphrase), optionally protected by Windows Hello (host policy).
    // On success, the host allocates a NUL-terminated UTF-16 string with CoTaskMemAlloc and stores it in *secretOut.
    // Callers must free it with CoTaskMemFree().
    //
    // This function does NOT prompt. If the secret is not available (not saved and no session-cached secret),
    // it returns HRESULT_FROM_WIN32(ERROR_NOT_FOUND).
    virtual HRESULT STDMETHODCALLTYPE GetConnectionSecret(
        const wchar_t* connectionName, HostConnectionSecretKind kind, HWND ownerWindow, wchar_t** secretOut) noexcept = 0;

    // Prompts the user (themed) for a secret (password/passphrase) and stores it in a per-session in-memory cache keyed
    // by (connectionId, secretKind). The secret is NOT persisted to WinCred.
    //
    // Returns:
    // - S_OK: secretOut is non-null (may be empty for SSH key passphrase to indicate "no passphrase")
    // - S_FALSE: user cancelled (secretOut is null)
    virtual HRESULT STDMETHODCALLTYPE PromptForConnectionSecret(
        const wchar_t* connectionName, HostConnectionSecretKind kind, HWND ownerWindow, wchar_t** secretOut) noexcept = 0;

    // Clears any per-session cached secret for this connection (does not modify WinCred).
    virtual HRESULT STDMETHODCALLTYPE ClearCachedConnectionSecret(const wchar_t* connectionName, HostConnectionSecretKind kind) noexcept = 0;

    // FTP-only: if a server rejects anonymous login, the plugin may ask the host to:
    // - prompt the user for credentials,
    // - persistently flip the profile to authMode=password + update userName,
    // - keep the password session-only (unless the user later saves it via Connection Manager).
    //
    // Returns:
    // - S_OK: profile updated and a session password is available
    // - S_FALSE: user cancelled
    virtual HRESULT STDMETHODCALLTYPE UpgradeFtpAnonymousToPassword(const wchar_t* connectionName, HWND ownerWindow) noexcept = 0;
};

enum HostPaneExecuteFlags : uint32_t
{
    HOST_PANE_EXECUTE_FLAG_NONE            = 0,
    HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW = 0x1,
};

struct HostPaneExecuteRequest
{
    uint32_t version;   // 1
    uint32_t sizeBytes; // sizeof(HostPaneExecuteRequest)

    HostPaneExecuteFlags flags;

    // Folder path to navigate to in the active pane (UTF-16, NUL-terminated, caller-owned).
    const wchar_t* folderPath;

    // Optional: leaf display name to focus after navigation (UTF-16, NUL-terminated, caller-owned).
    // This is NOT a path; it must not contain separators.
    const wchar_t* focusItemDisplayName;

    // Optional: FolderView command id to execute after navigation completes (e.g. IDM_FOLDERVIEW_CONTEXT_*). 0 = none.
    unsigned int folderViewCommandId;

    uint32_t reserved[8];
};

// UUID: {2f1a61a6-6e8c-4c1e-ae33-0f2cfb42e3b9}
interface __declspec(uuid("2f1a61a6-6e8c-4c1e-ae33-0f2cfb42e3b9")) __declspec(novtable) IHostPaneExecute : public IUnknown
{
    // Executes a request in the active pane. The host may activate its main window and navigate/focus as needed.
    virtual HRESULT STDMETHODCALLTYPE ExecuteInActivePane(const HostPaneExecuteRequest* request) noexcept = 0;
};

#pragma warning(pop)
