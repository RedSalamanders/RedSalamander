Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSTerminalText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
}

function Get-RSTerminalPluginImplementationSource {
    $parts = @(
        'Plugins\Terminal\Terminal.cpp'
        'Plugins\Terminal\TerminalInternal.h'
        'Plugins\Terminal\TerminalInternal.cpp'
        'Plugins\Terminal\TerminalSecurity.cpp'
        'Plugins\Terminal\TerminalSession.cpp'
        'Plugins\Terminal\TerminalVt.cpp'
        'Plugins\Terminal\TerminalVtUpgradeTests.cpp'
        'Plugins\Terminal\TerminalVtUpgradeTestContract.h'
    )
    return ($parts | ForEach-Object { Get-RSTerminalText -Path $_ }) -join "`n"
}

Describe 'Embedded Terminal plugin source contracts' {
    BeforeAll {
        $source = Get-RSTerminalPluginImplementationSource
        $header = Get-RSTerminalText -Path 'Plugins\Terminal\Terminal.h'
        $commandSurfaceModel = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalCommandSurfaceModel.cpp'
        $accessibility = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalAccessibility.cpp'
        $loader = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalRuntimeLoader.cpp'
        $loaderHeader = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalRuntimeLoader.h'
        $terminalProject = Get-RSTerminalText -Path 'Plugins\Terminal\Terminal.vcxproj'
        $kittyPipeline = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalKittyImagePipeline.cpp'
        $kittyPipelineHeader = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalKittyImagePipeline.h'
        $interfaceHeader = Get-RSTerminalText -Path 'Common\PlugInterfaces\Terminal.h'
        $runtimeBuilder = @(
            Get-RSTerminalText -Path 'Tools\Build-GhosttyTerminalRuntime.ps1'
            Get-RSTerminalText -Path 'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1'
        ) -join "`n"
        $runtimeLockText = Get-RSTerminalText -Path 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
        $runtimeLock = $runtimeLockText | ConvertFrom-Json -Depth 100
        $runtimeOverlay = Get-RSTerminalText -Path 'Tools\TerminalEngine\Patches\Ghostty-WindowsPrivateRuntime.patch'
        $runtimeManifest = Get-RSTerminalText -Path 'RuntimeDependencies.props'
        $pluginContractTests = Get-RSTerminalText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'
        $resources = Get-RSTerminalText -Path 'Plugins\Terminal\TerminalResources.rc'
        $folderNavigation = Get-RSTerminalText -Path 'RedSalamander\FolderWindow.FileSystem.Navigation.cpp'
        $commandsNavigation = Get-RSTerminalText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp'
        $folderHeader = Get-RSTerminalText -Path 'RedSalamander\FolderWindow.h'
        $folderLayout = Get-RSTerminalText -Path 'RedSalamander\FolderWindow.Layout.cpp'
        $folderInteraction = Get-RSTerminalText -Path 'RedSalamander\FolderWindow.Interaction.cpp'
        $floatingTerminal = Get-RSTerminalText -Path 'RedSalamander\FloatingTerminalWindow.cpp'
        $applicationSource = Get-RSTerminalText -Path 'RedSalamander\RedSalamander.cpp'
        $windowMessages = Get-RSTerminalText -Path 'Common\WindowMessages.h'
        $attributes = Get-RSTerminalText -Path '.gitattributes'
        $terminalSpec = Get-RSTerminalText -Path 'Specs\Terminal\Terminal_EmbeddedPlugin.md'
        $vtUpgradeCorpusText = Get-RSTerminalText -Path 'Specs\Terminal\TerminalVtUpgradeCorpus.v1.json'
        $vtUpgradeCorpus = $vtUpgradeCorpusText | ConvertFrom-Json
        $vtUpgradeEvidenceSchemaPath = Join-Path $repoRoot 'Specs\Terminal\TerminalVtUpgradeEvidence.schema.json'
    }

    It 'keeps Floating Terminal messages in the unique central registry' {
        $floatingTerminal | Should Not Match 'WM_APP\s*\+'
        $floatingTerminal | Should Match 'WndMsg::kFloatingTerminalCloseEmpty'
        $floatingTerminal | Should Match 'WndMsg::kFloatingTerminalRestoreNextTab'
        $windowMessages | Should Match 'kFloatingTerminalCloseEmpty\s*=\s*WM_APP\s*\+\s*0x188'
        $windowMessages | Should Match 'kFloatingTerminalRestoreNextTab\s*=\s*WM_APP\s*\+\s*0x189'

        $ids = @{}
        foreach ($match in [regex]::Matches($windowMessages, '(?m)^inline constexpr UINT\s+(?<name>\w+)\s*=\s*(?<family>WM_APP|WM_USER)\s*\+\s*(?<offset>0x[0-9A-Fa-f]+|\d+)u?\s*;')) {
            $offsetText = $match.Groups['offset'].Value
            $offset = if ($offsetText.StartsWith('0x', [System.StringComparison]::OrdinalIgnoreCase)) {
                [Convert]::ToUInt32($offsetText.Substring(2), 16)
            } else {
                [Convert]::ToUInt32($offsetText, 10)
            }
            $key = "{0}:{1}" -f $match.Groups['family'].Value, $offset
            if ($ids.ContainsKey($key)) {
                throw "Duplicate central window message value $key for $($ids[$key]) and $($match.Groups['name'].Value)."
            }
            $ids[$key] = $match.Groups['name'].Value
        }
        $ids.Count | Should BeGreaterThan 0
    }

    It 'renders from a reusable Ghostty structured snapshot instead of flattening styled cells' {
        $header | Should Match 'GhosttyRenderState\s+_renderState'
        $header | Should Match 'GhosttyRenderStateRowIterator\s+_renderRows'
        $header | Should Match 'GhosttyRenderStateRowCells\s+_renderCells'
        $source | Should Match 'renderStateBeginUpdate\(_renderState, _ghosttyTerminal\)'
        $source | Should Match 'renderStateEndUpdate\(_renderState\)'
        $source | Should Match 'GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_GRAPHEMES_UTF8'
        $source | Should Match 'GHOSTTY_RENDER_STATE_ROW_CELLS_DATA_STYLE'
        $source | Should Match 'GHOSTTY_RENDER_STATE_DATA_CURSOR_VISUAL_STYLE'
        $source | Should Match 'clearRenderDirtyState\(\)'
    }

    It 'uses measured pixel-aligned cells and seam-free background fills' {
        $source | Should Match 'CreateTextLayout\(kCellMeasureText'
        $source | Should Match 'layout->GetMetrics\(&metrics\)'
        $source | Should Match 'cellWidthPx\s*=\s*std::max\(1, static_cast<int>\(std::lround'
        $source | Should Match 'cellHeightPx\s*=\s*std::max\(1, static_cast<int>\(std::lround'
        $source | Should Match 'SetAntialiasMode\(D2D1_ANTIALIAS_MODE_ALIASED\)'
        $source | Should Match 'Debug::Perf::Scope renderPerf\(L"terminal\.render\.frame_us"\)'
    }

    It 'shares one plugin-action ID list and does not reimplement default chords in WindowProc' {
        $interfaceHeader | Should Match 'IsTerminalPluginActionId'
        $interfaceHeader | Should Match 'cmd/terminal/copySelectionOrPassthrough'
        $source | Should Match 'return IsTerminalPluginActionId\(commandId\)'
        $folderInteraction | Should Match 'IsTerminalPluginActionId\(commandId\)'
        $floatingTerminal | Should Match 'IsTerminalPluginActionId\(commandId\)'
        $routeInput = [regex]::Match(
            $source,
            '(?s)Terminal::MessageRouteResult Terminal::routeInputMessage\(.*?\n\}').Value
        $routeInput.Length | Should BeGreaterThan 0
        $routeInput | Should Not Match 'copySelection\(&selectionPresent\)'
        $routeInput | Should Not Match "wParam == static_cast<WPARAM>\('C'\)"
        $routeInput | Should Not Match "wParam == static_cast<WPARAM>\('A'\)"
        $routeInput | Should Not Match "wParam == static_cast<WPARAM>\('V'\)"
        $encode = [regex]::Match(
            $source,
            '(?s)bool\s+Terminal::encodeSpecialKey\(.*?\n\}').Value
        $encode.Length | Should BeGreaterThan 0
        $encode | Should Match 'VirtualKeyProducesTranslatedCharacter\(virtualKey\)'
        $encode | Should Match 'MarkTranslatedCharacterConsumed\(_translatedCharacterSuppressionPending\)'
        $source | Should Match 'routeInputMessage\(nullptr, WM_KEYDOWN, VK_RETURN'
        $source | Should Match 'routeInputMessage\(nullptr, WM_CHAR, static_cast<WPARAM>\(L''\\r''\)'
        $source | Should Match 'routeInputMessage\(nullptr, WM_KEYDOWN, VK_UP'
        $source | Should Match 'routeInputMessage\(nullptr, WM_CHAR, static_cast<WPARAM>\(L''a''\)'
        $source | Should Match 'confirmationEscapeKeydown\s*=\s*routeInputMessage\(nullptr, WM_KEYDOWN, VK_ESCAPE'
        $source | Should Match 'confirmationEscapeChar\s*=\s*routeInputMessage\(nullptr, WM_CHAR, static_cast<WPARAM>\(L''\\x1B''\)'
        $source | Should Match 'surfaceEscapeKeydown\s*=\s*routeInputMessage\(nullptr, WM_KEYDOWN, VK_ESCAPE'
        $source | Should Match 'surfaceEscapeChar\s*=\s*routeInputMessage\(nullptr, WM_CHAR, static_cast<WPARAM>\(L''\\x1B''\)'
        $source | Should Match 'case\s+VK_ESCAPE:\s*MarkTranslatedCharacterConsumed\(_translatedCharacterSuppressionPending\);\s*closeCommandSurface\(\)'
        $source | Should Match 'else if\s*\(wParam == VK_ESCAPE\)\s*\{\s*MarkTranslatedCharacterConsumed\(_translatedCharacterSuppressionPending\);\s*resolveConfirmation\(false\)'
    }

    It 'keeps the UI-facing Close method free of joins and wait APIs' {
        $close = [regex]::Match(
            $source,
            '(?s)HRESULT\s+STDMETHODCALLTYPE\s+Terminal::Close\(\)\s+noexcept\s*\{.*?\n\}').Value
        $close.Length | Should BeGreaterThan 0
        $close | Should Not Match '\.join\s*\('
        $close | Should Not Match 'WaitFor(Single|Multiple)Objects?'
        $close | Should Match 'SubmitThreadpoolWork\(_closeWork\.get\(\)\)'
        $source | Should Match 'Terminal::Terminal\(\)[\s\S]*?ensureCloseWork\(\)'
        $source | Should Match 'HRESULT\s+Terminal::ensureCloseWork\(\)[\s\S]*?CreateThreadpoolWork\(&Terminal::CloseThreadpoolCallback'
        $source | Should Match 'Terminal::~Terminal\(\)[\s\S]*?_closing\.exchange[\s\S]*?prepareClose\(\);[\s\S]*?finishClose\(\);'
        $source | Should Match 'TransferModulePinToCallbackReturn\(instance, self->_closeModulePin\)'
        $source | Should Match 'CloseThreadpoolCallback[\s\S]*?finishClose\(\)'
        $prepareClose = [regex]::Match(
            $source,
            '(?s)void\s+Terminal::prepareClose\(\)\s+noexcept\s*\{.*?\nvoid\s+CALLBACK').Value
        if ($prepareClose.Length -eq 0) {
            $prepareClose = [regex]::Match(
                $source,
                '(?s)void\s+Terminal::prepareClose\(\)\s+noexcept\s*\{.*?\nvoid\s+Terminal::finishClose').Value
        }
        $prepareClose.Length | Should BeGreaterThan 0
        $prepareClose | Should Not Match '\.join\s*\('
        $prepareClose | Should Not Match 'WaitFor(Single|Multiple)Objects?'
        $prepareClose | Should Match 'retireCommandSurfaceWorker\(\)'
        $source | Should Match 'void\s+Terminal::finishClose\(\)[\s\S]*?joinCommandSurfaceWorkers\(\)'
        $source | Should Match 'openCommandSurface[\s\S]*?retireCommandSurfaceWorker\(\)'
        $source | Should Match 'void\s+Terminal::retireCommandSurfaceWorker\(\)[\s\S]*?TrySubmitThreadpoolCallback\(&Terminal::CommandSurfaceJoinCallback'
        $source | Should Match 'CommandSurfaceJoinCallback[\s\S]*?TransferModulePinToCallbackReturn'
        $header | Should Match 'std::condition_variable\s+_commandSurfaceJoinCv'
        $source | Should Match '_commandSurfaceJoinCv\.wait\(lock,\s*\[this\]\(\)\s+noexcept\s*\{\s*return\s+_commandSurfaceJoinOutstanding\s*==\s*0u;\s*\}\)'
        ($header + $source) | Should Not Match '_commandSurfaceJoinsIdle|ResetEvent\(\)|SetEvent\(\)'
        $source | Should Match 'closeCommandSurface[\s\S]*?retireCommandSurfaceWorker\(\)'
    }

    It 'queues bounded input and performs pipe writes only on a pinned worker' {
        $header | Should Match 'std::deque<std::string>\s+_inputQueue'
        $source | Should Match 'kMaximumInputBytes\s*=\s*8u\s*\*\s*1024u\s*\*\s*1024u'
        $source | Should Match 'kMaximumInputRequests\s*=\s*3840u'
        $source | Should Match 'TrySubmitThreadpoolCallback\(&Terminal::InputDrainCallback'
        $source | Should Match 'InputDrainCallback[\s\S]*?TransferModulePinToCallbackReturn'
        $source | Should Match 'void\s+Terminal::drainInput\(\)[\s\S]*?WriteFile\('
        $writeInput = [regex]::Match(
            $source,
            '(?s)bool\s+Terminal::writeInput\(.*?\n\}').Value
        $writeInput.Length | Should BeGreaterThan 0
        $writeInput | Should Not Match 'WriteFile\('
        $source | Should Match 'clearQueuedInput\(\)[\s\S]*?SecureZeroMemory'
    }

    It 'admits a suspended shell into a plugin-owned kill-on-close job before it can run' {
        $header | Should Match 'wil::unique_handle\s+_processJob'
        $source | Should Match 'PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,\s*pseudoConsole,\s*sizeof\(pseudoConsole\)'
        $source | Should Not Match 'PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,\s*(?:const_cast<[^>]+>\s*\()?&pseudoConsole'
        $source | Should Match 'PIPE_ACCESS_DUPLEX\s*\|\s*FILE_FLAG_OVERLAPPED\s*\|\s*FILE_FLAG_FIRST_PIPE_INSTANCE'
        $source | Should Match 'GetNamedPipeClientProcessId\(pipe\.get\(\), &clientProcessId\)[\s\S]*?clientProcessId != rootProcessId'
        $source | Should Match 'WriteIntegrationReply\([\s\r\n]*pipe\.get\(\), stopToken, request\.versioned, reply\.followTarget, reply\.promptReplacement\)'
        $source | Should Not Match 'Set-Location -LiteralPath " \+ quotePath\(followTarget\)'
        $source | Should Match 'JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE'
        $source | Should Match 'EXTENDED_STARTUPINFO_PRESENT \| CREATE_UNICODE_ENVIRONMENT \| CREATE_SUSPENDED,'
        $source | Should Match 'startup\.StartupInfo\.dwFlags = STARTF_USESTDHANDLES'
        $source | Should Match 'AssignProcessToJobObject\(processJob\.get\(\), process\.get\(\)\)'
        $source | Should Match 'ResumeThread\(processThread\.get\(\)\)'
        $source | Should Match '_processJob\.reset\(\);[\s\S]*?TerminateProcess'
        $startPseudoConsole = [regex]::Match(
            $source,
            '(?s)HRESULT\s+Terminal::startPseudoConsole\(.*?\nvoid\s+Terminal::integrationMain').Value
        $startPseudoConsole.Length | Should BeGreaterThan 0
        $startPseudoConsole | Should Not Match 'COORD\s*\{\s*1\s*,\s*1\s*\}'
        $resize = [regex]::Match(
            $source,
            '(?s)void\s+Terminal::resizeTerminal\(\)\s+noexcept\s*\{.*?\nvoid\s+Terminal::updateCellMetrics').Value
        $resize.Length | Should BeGreaterThan 0
        $resize | Should Match 'if \(emptyClient\)\s*\{[^}]*return;'
        $resize | Should Not Match 'emptyClient\s*\?\s*1u'
        $source | Should Match 'commandLine\.append\(L" -NoExit"\)'
        $source | Should Match 'commandLine\.append\(L" -Interactive"\)'
        $source | Should Match 'EqualsNoCase\(std::wstring_view\(executable\)\.substr\(executable\.size\(\) - 8u\), L"pwsh\.exe"\)'
        $startPseudoConsole | Should Not Match 'ResizePseudoConsole\('
        $resize | Should Match '_ptyClientSizeSynced && columns == _columns && rows == _rows'
        $source | Should Match 'hr = startPseudoConsole\(context->launchLocation\);[\s\S]*?resizeTerminal\(\);'
        $folderLayout | Should Not Match 'SWP_NOSIZE'
    }

    It 'routes every legacy path insertion command through the opposite terminal and removes the pseudo prompt' {
        $currentDirectory = [regex]::Match(
            $folderNavigation,
            '(?s)void\s+FolderWindow::CommandBringCurrentDirToCommandLine\(.*?\n\}').Value
        $currentDirectory.Length | Should BeGreaterThan 0
        $currentDirectory | Should Match 'OpenTerminalPane\(pane, workingDirectory\.value\(\)\)'
        $currentDirectory | Should Match 'TerminalPathInsertionMode::CurrentDirectoryFull'
        $currentDirectory | Should Match 'initiatingSourceLocation\s*=\s*sourceLocation\.View\(\)'
        ([regex]::Matches(
                $folderNavigation,
                'insertion\.initiatingSource\.paneInstanceId\s*=\s*host\.terminalOriginalSourcePaneInstanceId')).Count |
            Should Be 2
        $folderNavigation | Should Not Match 'insertion\.initiatingSource\.paneInstanceId\s*=\s*pane\s*=='
        $commandsNavigation | Should Match 'CommandBringCurrentDirToCommandLine\(FolderWindow::Pane::Right\)[\s\S]{0,1800}CommandInsertFocusedPathInTerminal\(FolderWindow::Pane::Right, false\)'
        $source | Should Match 'TerminalPathInsertionMode::CurrentDirectoryFull[\s\S]*?insertion->initiatingSourceLocation'
        $folderNavigation | Should Not Match 'ShowCommandLine\(|LaunchCommandLine\(|ExecuteCommandLineFromEdit'
        $folderHeader | Should Not Match 'CommandLineDebugSnapshot|_hCommandLineHost|_commandLineField'
        $folderLayout | Should Not Match '_commandLineRect|UpdateCommandLineHostLayout'
    }

    It 'keeps inactive pane tabs out of the header and gives terminal input priority over host shortcuts' {
        $folderLayout | Should Match 'SetTabVisible\(1u, previewAvailable\)'
        $folderLayout | Should Match 'SetTabVisible\(2u, host\.terminalOpen\)'
        $folderLayout | Should Match '_leftPreviewContentRect : _rightPreviewContentRect'
        $folderHeader | Should Match 'IsTerminalInputTarget\(HWND targetWindow\) const noexcept'
        $folderInteraction | Should Match 'targetWindow == terminalWindow \|\| IsChild\(terminalWindow, targetWindow\)'
        $applicationSource | Should Match 'g_folderWindow\.IsTerminalInputTarget\(msg\.hwnd\)'
        $applicationSource | Should Match '! terminalInputTarget && TryReclaimMainFolderViewFocusOnEscape'
        $applicationSource | Should Match 'isMainWindowMessage && ! editFocused && ! terminalInputTarget'
        $applicationSource | Should Match 'terminalInputTarget \|\| ! isMainWindowMessage \|\| editFocused \|\| ! TranslateAccelerator'
    }

    It 'binds path insertion to the latest accepted source and marks input only after admission' {
        $source | Should Match 'insertion->initiatingSource\.folderWindowInstanceId != _originalSource\.folderWindowInstanceId'
        $source | Should Match 'insertion->initiatingSourceGeneration != _sourceGeneration'
        $source | Should Match 'insertion->initiatingSourceLocation\.kind != _sourceLocationKind'
        $source | Should Match '_hasPendingUserInput, _trustedCurrentDirectory'
        $folderNavigation | Should Match 'UpdateSourceLocation\(&update\);[\s\S]*?if \(SUCCEEDED\(updateHr\)\)[\s\S]*?terminalSourceGeneration = update\.sourceGeneration'

        $writeTextInput = [regex]::Match(
            $source,
            '(?s)bool\s+Terminal::writeTextInput\(.*?\n\}').Value
        $writeTextInput.Length | Should BeGreaterThan 0
        $writeTextInput | Should Match 'if \(! writeInput\(utf8\)\)[\s\S]*?markUserInput\(utf8\)'
    }

    It 'keeps clipboard reading, paste policy, and bracketed encoding inside Terminal.dll' {
        $source | Should Match 'ReadBoundedClipboardText'
        $source | Should Match 'OpenClipboard\(ownerWindow\)'
        $source | Should Match 'GlobalSize\(handle\)'
        $source | Should Match 'TryUtf8FromUtf16Strict\(normalized\)'
        $source | Should Match 'pasteIsSafe\(source\.data\(\), source\.size\(\)\)'
        $source | Should Match 'GetTerminalMode\(_ghosttyTerminal, GHOSTTY_MODE_BRACKETED_PASTE, bracketed\)'
        $source | Should Match 'pasteEncode\('
        $source | Should Match 'case WM_PASTE:'
        $source | Should Match 'pasteClipboard\(\)'
        $source | Should Match 'SecureZeroMemory\(encoded\.data\(\), encoded\.size\(\)\)'
        $source | Should Match '_pendingUnsafePaste\s*=\s*std::move\(source\)'
        $source | Should Match '_unsafePasteConfirmationVisible\s*=\s*true'
        $source | Should Match 'void\s+Terminal::resolveConfirmation\(bool accept\)'
        $source | Should Match 'drawConfirmation\(widthDip, heightDip\)'
        $source | Should Match 'confirmationVisible\(\)'
        $source | Should Match 'SecureZeroMemory\(_pendingUnsafePaste\.data\(\), _pendingUnsafePaste\.size\(\)\)'
        $resources | Should Match 'IDS_TERMINAL_UNSAFE_PASTE_MESSAGE'
    }

    It 'defers bounded OSC 52 text writes to the UI thread behind deny ask allow policy' {
        $header | Should Match 'std::wstring\s+osc52Policy\s*=\s*L"ask"'
        $header | Should Match 'uint32_t\s+osc52MaxBytes\s*=\s*256u\s*\*\s*1024u'
        $header | Should Match 'std::atomic_bool\s+_osc52PostPending'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE'
        $header | Should Match 'static void GhosttyClipboardWrite'
        $header | Should Not Match 'static GhosttyClipboardWriteResult GhosttyClipboardWrite'
        $source | Should Match 'GhosttyClipboardWriteResult\s+Terminal::handleGhosttyClipboardWrite'
        $source | Should Match 'offsetof\(::GhosttyClipboardWrite, reply\)'
        $source | Should Match 'GhosttyClipboardWriteReply reply\{\}'
        $source | Should Match 'write->reply\(write, &reply\)'
        $source | Should Match 'reply\.remember\s*=\s*false'
        $source | Should Match 'kGhosttyClipboardWriteMaxBytes\s*=\s*256u\s*\*\s*1024u'
        $source | Should Match 'kMaximumOsc52Bytes\s*=\s*256u\s*\*\s*1024u'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_CLIPBOARD_READ, nullptr'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED'
        $source | Should Match 'kittyClipboardWriteEnabled\s*=\s*false'
        $source | Should Match 'GHOSTTY_CLIPBOARD_LOCATION_STANDARD'
        $source | Should Match 'kMaximumClipboardRepresentations\s*=\s*16u'
        $source | Should Match 'mime == "text/plain"'
        $source | Should Match 'mime == "text/plain;charset=utf-8"'
        $source | Should Match '_osc52MaxBytes\.load'
        $source | Should Match 'std::memchr\(selectedView\.data, 0, selectedView\.length\)'
        $source | Should Match 'IsValidUtf8Strict\(selectedText\)'
        $source | Should Match 'payload->utf8\.assign'
        $source | Should Match 'PostMessagePayload\(hwnd, WndMsg::kTerminalClipboardWrite'
        $source | Should Match 'InitPostedPayloadWindow\(hwnd\)'
        $source | Should Match 'TakeMessagePayload<Osc52ClipboardPayload>\(lParam\)'
        $source | Should Match 'DrainPostedPayloadsForWindow\(hwnd\)'
        $source | Should Match 'TryUtf16FromUtf8Strict\(payload->utf8\)'
        $source | Should Match 'TrySetUnicodeText\('
        $source | Should Match '_osc52ConfirmationVisible\s*=\s*true'
        $resources | Should Match 'IDS_TERMINAL_OSC52_MESSAGE'

        $callback = [regex]::Match(
            $source,
            '(?s)GhosttyClipboardWriteResult\s+Terminal::handleGhosttyClipboardWrite\(.*?\n\}').Value
        $callback.Length | Should BeGreaterThan 0
        $callback | Should Not Match 'OpenClipboard|EmptyClipboard|SetClipboardData|TrySetUnicodeText'
        $callback | Should Match 'report BUSY rather than claiming that consent or clipboard I/O succeeded'
        $callback | Should Match 'return GHOSTTY_CLIPBOARD_WRITE_RESULT_BUSY;\s*\}'

        $source | Should Match 'kGhosttyClipboardWriteMaxBytes - 1u'
        $source | Should Match 'kGhosttyClipboardWriteMaxBytes \+ 1u'
        $source | Should Match 'OSC 52 ask policy rejects overlapping UI work as busy'
        $source | Should Match 'OSC 52 callback rejects invalid UTF-8 before posting UI work'
        $source | Should Match 'OSC 52 callback denies teardown-before-UI work'
        $source | Should Match 'clipboardCapture\.replyCount == 1u'
        $source | Should Match 'type=write:status=ENOSYS:id=x'
        $source | Should Match 'type=read:status=EPERM:id=r'
    }

    It 'activates bounded OSC 8 hyperlinks only through the configured safe policy' {
        $header | Should Match 'std::wstring\s+hyperlinkPolicy\s*=\s*L"ask"'
        $source | Should Match 'kMaximumHyperlinkBytes\s*=\s*8192u'
        $source | Should Match 'gridRefHyperlinkUri\(&cell, nullptr, 0u, &required\)'
        $source | Should Match 'TryUtf16FromUtf8Strict'
        $source | Should Match 'StartsWithNoCase\(uri, L"http://"\)'
        $source | Should Match 'StartsWithNoCase\(uri, L"https://"\)'
        $source | Should Match 'StartsWithNoCase\(uri, L"mailto:"\)'
        $source | Should Match 'ShellExecuteW\(hwnd, L"open"'
        $source | Should Match 'GetKeyState\(VK_CONTROL\)[\s\S]*?activateHyperlinkAt\(point\)'
        $source | Should Match '_hyperlinkConfirmationVisible\s*=\s*true'
        $resources | Should Match 'IDS_TERMINAL_HYPERLINK_MESSAGE'
    }

    It 'uses Ghostty-owned selection semantics for pointer selection and copy' {
        $source | Should Match 'terminalGridRef\(_ghosttyTerminal'
        $source | Should Match 'terminalGridRefTrack\([\s\r\n]*_ghosttyTerminal'
        $source | Should Match 'trackedGridRefSnapshot\(_selectionAnchor'
        $source | Should Match 'trackedGridRefFree\(_selectionAnchor\)'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_SELECTION'
        $source | Should Match 'terminalSelectWord\(_ghosttyTerminal'
        $source | Should Match 'terminalSelectAll\(_ghosttyTerminal'
        $source | Should Match 'terminalSelectionFormatBuffer\(_ghosttyTerminal'
        $source | Should Match 'Common::Clipboard::TrySetUnicodeText'
        $source | Should Match 'case WM_LBUTTONDBLCLK:'
        $source | Should Match 'case WM_COPY:'
        $source | Should Match 'timerId == kSelectionAutoScrollTimer'
        $source | Should Match 'SetTimer\(hwnd, kSelectionAutoScrollTimer, 50u, nullptr\)'
        $source | Should Match 'const bool selectionPresent = hasSelection\(\);[\s\S]*?copySelection\(\)[\s\S]*?TerminalShortcutRoute::Handled'
    }

    It 'routes application mouse modes through Ghostty and otherwise scrolls the viewport' {
        $header | Should Match 'GhosttyMouseEncoder\s+_mouseEncoder'
        $header | Should Match 'GhosttyMouseEvent\s+_mouseEvent'
        $source | Should Match 'mouseEncoderSetFromTerminal\(_mouseEncoder, _ghosttyTerminal\)'
        $source | Should Match 'mouseEncoderEncode\('
        $source | Should Match 'GHOSTTY_MOUSE_BUTTON_FOUR'
        $source | Should Match 'GHOSTTY_MOUSE_BUTTON_FIVE'
        $source | Should Match 'terminalScrollViewport\(_ghosttyTerminal, scroll\)'
        $source | Should Match 'case WM_MOUSEWHEEL:'
    }

    It 'round-trips one visible native scrollbar through Ghostty absolute rows' {
        $source | Should Match 'WS_TABSTOP \| WS_VSCROLL'
        $source | Should Match 'GHOSTTY_TERMINAL_DATA_SCROLLBAR'
        $source | Should Match 'GHOSTTY_SCROLL_VIEWPORT_ROW'
        $source | Should Match 'case WM_VSCROLL:'
        $source | Should Match 'case SB_THUMBTRACK:'
        $source | Should Match 'SetScrollInfo\(hwnd, SB_VERT, &info, TRUE\)'
        $source | Should Match '_themeHighContrast \? L"" : \(_themeDark \? L"DarkMode_Explorer" : L"Explorer"\)'
        $source | Should Match 'SetWindowTheme\(hwnd, themeName, nullptr\)'
        $source | Should Match 'SendMessageW\(hwnd, WM_THEMECHANGED, 0, 0\)'
        $source | Should Match 'RedrawWindow\(hwnd, nullptr, nullptr, RDW_INVALIDATE \| RDW_FRAME\)'
        $source | Should Match '_windowHandle\.store[\s\S]*?applyNativeScrollbarTheme\(_window\.get\(\)\)'
        $source | Should Match 'Terminal::SetTheme[\s\S]*?applyNativeScrollbarTheme\(hwnd\)'
    }

    It 'dims the complete unfocused terminal scene through the shared pane policy' {
        $source | Should Match '#include "PaneVisualState\.h"'
        $header | Should Match 'void\s+drawInactivePaneOverlay\(float widthDip, float heightDip, bool paneFocused\) noexcept'
        $source | Should Match 'renderPerf\.SetDetail\(paneFocused \? L"focused" : L"unfocused"\)'
        $source | Should Match 'drawConfirmation\(widthDip, heightDip\);\s*drawInactivePaneOverlay\(widthDip, heightDip, paneFocused\);'
        $source | Should Match 'Common::PaneVisualState::ResolveInactiveContentOverlayAlpha\(paneFocused, _themeHighContrast\)'
        $source | Should Match 'FillRectangle\(D2D1::RectF\(0\.0f, 0\.0f, widthDip, heightDip\), _foregroundBrush\.get\(\)\)'
        $source | Should Match 'Terminal::routeAccessibilityMessage[\s\S]*?case WM_SETFOCUS:[\s\S]*?InvalidateRect\(hwnd, nullptr, FALSE\)[\s\S]*?case WM_KILLFOCUS:[\s\S]*?InvalidateRect\(hwnd, nullptr, FALSE\)'
    }

    It 'keeps IME composition, commit, and candidate placement in the plugin control' {
        $header | Should Match 'std::wstring\s+_imeComposition'
        $source | Should Match 'case WM_IME_STARTCOMPOSITION:'
        $source | Should Match 'case WM_IME_COMPOSITION:'
        $source | Should Match 'GCS_RESULTSTR'
        $source | Should Match 'GCS_COMPSTR'
        $source | Should Match 'ImmSetCandidateWindow'
        $source | Should Match 'ImmSetCompositionWindow'
        $source | Should Match 'DrawTextW\(_imeComposition\.data\(\)'
    }

    It 'routes WindowProc domains through handled-result contracts and retires the window last' {
        $header | Should Match 'struct\s+MessageRouteResult\s+final[\s\S]*?bool\s+handled[\s\S]*?LRESULT\s+result'
        foreach ($route in @(
                'routeInputMessage',
                'routeMouseMessage',
                'routeImeMessage',
                'routeAccessibilityMessage',
                'routeSessionMessage',
                'routeClipboardMessage',
                'routeRenderMessage',
                'routeTeardownMessage')) {
            $header | Should Match ([regex]::Escape($route + '('))
            $source | Should Match ([regex]::Escape('Terminal::' + $route + '('))
        }

        $windowProc = [regex]::Match(
            $source,
            '(?s)LRESULT\s+CALLBACK\s+Terminal::WindowProc\(.*?\n\}').Value
        $windowProc.Length | Should BeGreaterThan 0
        $windowProc | Should Not Match 'switch\s*\('
        $windowProc | Should Match 'if\s*\(message == WM_NCCREATE\)[\s\S]*?InitPostedPayloadWindow\(hwnd\)'
        $windowProc | Should Match 'if\s*\(self == nullptr\)[\s\S]*?DefWindowProcW\(hwnd, message, wParam, lParam\)'

        $previousRouteIndex = -1
        foreach ($route in @(
                'routeInputMessage',
                'routeMouseMessage',
                'routeImeMessage',
                'routeAccessibilityMessage',
                'routeSessionMessage',
                'routeClipboardMessage',
                'routeRenderMessage',
                'routeTeardownMessage')) {
            $routeIndex = $windowProc.IndexOf($route, [StringComparison]::Ordinal)
            $routeIndex | Should BeGreaterThan $previousRouteIndex
            $previousRouteIndex = $routeIndex
        }
        $windowProc.LastIndexOf('DefWindowProcW', [StringComparison]::Ordinal) | Should BeGreaterThan $previousRouteIndex

        $teardown = [regex]::Match(
            $source,
            '(?s)Terminal::MessageRouteResult\s+Terminal::routeTeardownMessage\(.*?\n\}').Value
        $teardown.Length | Should BeGreaterThan 0
        $drainIndex = $teardown.IndexOf('DrainPostedPayloadsForWindow', [StringComparison]::Ordinal)
        $pendingIndex = $teardown.IndexOf('_osc52PostPending.store', [StringComparison]::Ordinal)
        $retireIndex = $teardown.IndexOf('_accessibility->Retire()', [StringComparison]::Ordinal)
        $windowIndex = $teardown.IndexOf('_windowHandle.store(nullptr', [StringComparison]::Ordinal)
        $userDataIndex = $teardown.IndexOf('SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0)', [StringComparison]::Ordinal)
        $countIndex = $teardown.IndexOf('g_terminalWindowCount.fetch_sub', [StringComparison]::Ordinal)
        $defaultIndex = $teardown.IndexOf('DefWindowProcW(hwnd, message, wParam, lParam)', [StringComparison]::Ordinal)
        $drainIndex | Should BeGreaterThan -1
        $drainIndex | Should BeLessThan $pendingIndex
        $pendingIndex | Should BeLessThan $retireIndex
        $retireIndex | Should BeLessThan $windowIndex
        $windowIndex | Should BeLessThan $userDataIndex
        $userDataIndex | Should BeLessThan $countIndex
        $countIndex | Should BeLessThan $defaultIndex
    }

    It 'copies a bounded exact-key Kitty visible delta under engine ownership and converts it on one latest-only worker' {
        $header | Should Match 'GhosttyKittyGraphicsPlacementIterator\s+_kittyPlacementIterator'
        $header | Should Match 'TerminalKittyImagePipeline\s+_kittyPipeline'
        $header | Should Match 'std::unordered_map<KittyImageKey,\s*KittyImageSnapshot,\s*KittyImageKeyHash>\s+_kittyImages'
        $header | Should Match 'uint32_t\s+imageId[\s\S]*?uint64_t\s+imageGeneration[\s\S]*?operator=='
        $header | Should Match 'enum class KittyImageState[\s\S]*?EnginePending[\s\S]*?ConversionPending[\s\S]*?Ready[\s\S]*?Rejected'
        $header | Should Match 'needsKittyReplacementWork'
        $kittyPipelineHeader | Should Match 'MaximumPixels\s*=\s*16u\s*\*\s*1024u\s*\*\s*1024u'
        $kittyPipelineHeader | Should Match 'MaximumSourceBytes\s*=\s*64u\s*\*\s*1024u\s*\*\s*1024u'
        $kittyPipelineHeader | Should Match 'MaximumConvertedBytes\s*=\s*64u\s*\*\s*1024u\s*\*\s*1024u'
        $kittyPipelineHeader | Should Match 'std::optional<TerminalKittyGenerationWork>\s+_pending'
        $kittyPipelineHeader | Should Match 'std::optional<TerminalKittyGenerationResult>\s+_ready'
        $kittyPipelineHeader | Should Match 'std::jthread\s+_worker'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_FILE, &externalMedium'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_TEMP_FILE, nullptr'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_IMAGE_MEDIUM_SHARED_MEM, &externalMedium'
        $source | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS, &placementLimits'
        $source | Should Match 'captureKittyGraphicsLocked\(kittyWork, kittyPlacements, submitKittyWork\)'
        $source | Should Match '_kittyImages\.try_emplace\(imageKey\)'
        $source | Should Match 'imageIterator->first\.imageId != imageId[\s\S]*?_kittyImages\.erase\(imageIterator\)[\s\S]*?_kittySupersededImageCount'
        $source | Should Match '_kittyImages\.find\(placement\.imageKey\)'
        $source | Should Match 'imageWork\.source\.assign\(source, source \+ sourceLength\)'
        $source | Should Match 'result == GHOSTTY_NO_VALUE \? KittyImageState::EnginePending : KittyImageState::Rejected'
        $source | Should Match 'image\.state == KittyImageState::EnginePending \|\| retryDeferred'
        $source | Should Match 'image\.state == KittyImageState::Rejected'
        $source | Should Match 'resetKittyPendingForStorageChange\(image\)'
        $source | Should Match 'isCurrentKittyConversionResult\(imageIterator->second, result->requestId\)'
        $source | Should Match '_kittyPipeline\.Submit\(std::move\(kittyWork\)\)'
        $kittyPipeline | Should Match 'work = std::move\(_pending\.value\(\)\)'
        $kittyPipeline | Should Match 'const bool converted = Convert\(std::move\(work\), result\)'
        $kittyPipeline | Should Match 'result\.requestId != _latestRequestId'
        $kittyPipeline | Should Match '\+\+_droppedPendingGenerations'
        $kittyPipeline | Should Match '\+\+_rejectedStaleGenerations'
        $kittyPipeline | Should Not Match 'CreateBitmap|ID2D1|ID2D'
        $source | Should Not Match 'std::ranges::find_if\(\s*_kittyImages'
        $source | Should Not Match 'std::ranges::any_of\(\s*work\.images'
        $source | Should Match '_kittyPipeline\.RequestStop\(\)[\s\S]*?_window\.reset\(\)'
        $source | Should Match '_kittyPipeline\.Stop\(\)[\s\S]*?_runtime\.Unload\(\)'
        $source | Should Match 'GHOSTTY_KITTY_IMAGE_FORMAT_GRAY_ALPHA'
        $source | Should Match 'D2D1_ALPHA_MODE_PREMULTIPLIED'
        $source | Should Match 'drawKittyLayer\(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_BG\)'
        $source | Should Match 'drawKittyLayer\(GHOSTTY_KITTY_PLACEMENT_LAYER_BELOW_TEXT\)'
        $source | Should Match 'drawKittyLayer\(GHOSTTY_KITTY_PLACEMENT_LAYER_ABOVE_TEXT\)'
        $loader | Should Match 'GUID_WICPixelFormat32bppRGBA'
        $loader | Should Match 'pixels > kMaximumKittyPixels'
        $loader | Should Match 'sysSet\(GHOSTTY_SYS_OPT_DECODE_PNG, nullptr\)'
        $loader | Should Match 'thread_local IWICImagingFactory\* g_pngDecodeWicFactory'
        $loader | Should Match 'TerminalPngDecodeThreadContext::TerminalPngDecodeThreadContext\(\) noexcept'
        $source | Should Match 'void Terminal::readerMain\([\s\S]*?TerminalPngDecodeThreadContext pngDecodeContext;'
        $decodePng = [regex]::Match(
            $loader,
            '(?s)\[\[nodiscard\]\] bool DecodePng\(.*?\r?\n\}\r?\n\r?\n\[\[nodiscard\]\] std::wstring GetRuntimePath').Value
        $decodePng | Should Match 'g_pngDecodeWicFactory'
        $decodePng | Should Not Match 'CoInitializeEx|CoCreateInstance'
        $runtimeOverlay | Should Match 'GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS\s*=\s*40'
        foreach ($dataKey in @(
                'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT = 3',
                'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES = 4',
                'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT_LIMIT = 5',
                'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_BYTES_LIMIT = 6')) {
            $runtimeOverlay | Should Match ([regex]::Escape($dataKey))
        }
        $runtimeOverlay | Should Match 'maximum_placement_count:\s*usize\s*=\s*4096'
        $runtimeOverlay | Should Match 'maximum_placement_bytes:\s*usize\s*=\s*1024\s*\*\s*1024'
    }

    It 'bounds Kitty upload work per frame, preserves CPU generations on device loss, and emits stage metrics' {
        $source | Should Match 'kMaximumKittyUploadBytesPerFrame\s*=\s*64u\s*\*\s*1024u\s*\*\s*1024u'
        $source | Should Match 'kMaximumKittyUploadsPerFrame\s*=\s*1u'
        $source | Should Match '_kittyFrameUploadCount >= kMaximumKittyUploadsPerFrame'
        $source | Should Match 'uploadBytes > kMaximumKittyUploadBytesPerFrame - _kittyFrameUploadBytes'
        $discard = [regex]::Match(
            $source,
            '(?s)void\s+Terminal::discardDeviceResources\(\)\s+noexcept\s*\{.*?\n\}').Value
        $discard | Should Match 'image\.bitmap\.reset\(\)'
        $discard | Should Not Match 'image\.bgra\.clear|_kittyImages\.clear'
        $header | Should Match 'enum class KittyBitmapUploadDisposition[\s\S]*?RetryScheduled[\s\S]*?RecreateTarget[\s\S]*?PermanentFailure'
        $source | Should Match 'Terminal::KittyBitmapUploadResult\s+Terminal::ensureKittyBitmap'
        $source | Should Match 'bool\s+Terminal::recoverKittyDeviceAfterFrame'
        foreach ($metric in @(
                'terminal.kitty.capture_lock_us',
                'terminal.kitty.cache_reuse_capture_lock_us',
                'terminal.kitty.submit_us',
                'terminal.kitty.queue_bytes',
                'terminal.kitty.convert_us',
                'terminal.kitty.upload_us',
                'terminal.kitty.memory_high_water_bytes',
                'terminal.kitty.placement_total_count',
                'terminal.kitty.placement_visible_count',
                'terminal.kitty.placement_offscreen_count',
                'terminal.kitty.placement_virtual_count',
                'terminal.kitty.visible_image_key_count',
                'terminal.kitty.engine_placement_count',
                'terminal.kitty.engine_placement_bytes',
                'terminal.kitty.active_source_bytes',
                'terminal.kitty.active_converted_bytes',
                'terminal.kitty.ready_bytes',
                'terminal.kitty.converted_bytes',
                'terminal.kitty.cache_bytes',
                'terminal.kitty.pinned_bytes',
                'terminal.kitty.evicted_bytes',
                'terminal.kitty.engine_pending_image_key_count',
                'terminal.kitty.superseded_image_count',
                'terminal.kitty.removed_image_count',
                'terminal.kitty.legacy_linear_capture_comparison_count',
                'terminal.kitty.legacy_linear_draw_comparison_count',
                'terminal.kitty.legacy_linear_draw_lookup_us',
                'terminal.kitty.image_lookup_count',
                'terminal.kitty.keyed_draw_lookup_us',
                'terminal.kitty.frame_upload_bytes',
                'terminal.kitty.frame_upload_count',
                'terminal.kitty.active_close_us',
                'terminal.kitty.upload_retry_scheduled_count',
                'terminal.kitty.upload_stable_failure_count',
                'terminal.kitty.device_recovery_count',
                'terminal.kitty.repaint_request_count',
                'terminal.kitty.resize_us',
                'terminal.kitty.resize_count',
                'terminal.kitty.resize_recreate_count',
                'terminal.kitty.resize_retained_bitmap_count',
                'terminal.kitty.resize_retained_bitmap_bytes',
                'terminal.frame_total_us',
                'terminal.input_to_frame_us',
                'terminal.output_to_frame_us')) {
            ($source + $kittyPipeline) | Should Match ([regex]::Escape('L"' + $metric + '"'))
        }
        $pluginContractTests | Should Match '--terminal-selftests'
        $pluginContractTests | Should Match 'RedSalamanderTerminalDebugSelfTests[\s\S]*?true,[\s\S]*?focusedSuccess'
    }

    It 'retains Terminal device resources across ordinary WM_SIZE and reserves full discard for policy or device loss' {
        $renderRoute = [regex]::Match(
            $source,
            '(?s)Terminal::MessageRouteResult\s+Terminal::routeRenderMessage\(.*?\n\}').Value
        $renderRoute.Length | Should BeGreaterThan 0
        $sizeCase = [regex]::Match($renderRoute, '(?s)case\s+WM_SIZE:.*?return\s+\{true,\s*0\};').Value
        $sizeCase | Should Match 'handleResize\(hwnd\)'
        $sizeCase | Should Not Match 'discardDeviceResources\(\)'
        $resizeHandler = [regex]::Match(
            $source,
            '(?s)void\s+Terminal::handleResize\(HWND hwnd\)\s+noexcept\s*\{.*?\n\}').Value
        $resizeHandler | Should Match 'resizeRenderTargetToClient\(\)'
        $resizeHandler | Should Not Match 'discardDeviceResources\(\)'
    }

    It 'keeps every authoritative Terminal TestRuns reference resolvable' {
        $terminalSpec | Should Not Match '2026-08-09_150404'
        $references = @([regex]::Matches(
                $terminalSpec,
                'Specs/TestRuns/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+') |
                ForEach-Object Value |
                Sort-Object -Unique)
        $references.Count | Should BeGreaterThan 0
        foreach ($reference in $references) {
            Test-Path -LiteralPath (Join-Path $repoRoot $reference) -PathType Container | Should Be $true
        }
    }

    It 'publishes an immutable UI Automation Text snapshot and retires passive providers safely' {
        $accessibility | Should Match 'class TerminalProvider final : public IRawElementProviderSimple, public ITextProvider'
        $accessibility | Should Match 'class TerminalTextRange final : public ITextRangeProvider'
        $accessibility | Should Match 'std::shared_ptr<const Snapshot>\s+snapshot'
        $accessibility | Should Match 'UIA_E_ELEMENTNOTAVAILABLE'
        $accessibility | Should Match 'g_accessibilityObjectCount\.fetch_add'
        $accessibility | Should Match 'CanUnloadTerminalAccessibilityProviders'
        $accessibility | Should Match 'UiaDisconnectProvider'
        $source | Should Match 'case WM_GETOBJECT:'
        $source | Should Match 'void\s+Terminal::publishAccessibilitySnapshot\(\)\s+noexcept'
        $source | Should Match 'UiaClientsAreListening\(\)'
        $source | Should Match '_accessibility->Publish\(std::move\(text\),'
        $source | Should Match '_accessibility->Retire\(\)'
        $source | Should Match 'CanUnloadTerminalAccessibilityProviders\(\)'
        $accessibility | Should Not Match 'Ghostty'
        $accessibility | Should Not Match 'TerminalRuntimeLoader'
    }

    It 'fails the private runtime build and loader closed when required terminal APIs disappear' {
        foreach ($export in @(
                'ghostty_render_state_begin_update',
                'ghostty_render_state_end_update',
                'ghostty_render_state_row_cells_get',
                'ghostty_cell_get',
                'ghostty_key_encoder_setopt_from_terminal',
                'ghostty_key_encoder_encode',
                'ghostty_paste_is_safe',
                'ghostty_paste_encode',
                'ghostty_terminal_grid_ref',
                'ghostty_grid_ref_hyperlink_uri',
                'ghostty_terminal_grid_ref_track',
                'ghostty_tracked_grid_ref_free',
                'ghostty_tracked_grid_ref_snapshot',
                'ghostty_terminal_select_word',
                'ghostty_terminal_select_all',
                'ghostty_terminal_selection_format_buf',
                'ghostty_terminal_scroll_viewport',
                'ghostty_mouse_encoder_setopt_from_terminal',
                'ghostty_mouse_encoder_encode',
                'ghostty_mouse_event_set_position',
                'ghostty_terminal_get',
                'ghostty_alloc',
                'ghostty_free',
                'ghostty_sys_set',
                'ghostty_kitty_graphics_get',
                'ghostty_kitty_graphics_image',
                'ghostty_kitty_graphics_image_get_multi',
                'ghostty_kitty_graphics_placement_iterator_new',
                'ghostty_kitty_graphics_placement_iterator_free',
                'ghostty_kitty_graphics_placement_next',
                'ghostty_kitty_graphics_placement_get',
                'ghostty_kitty_graphics_placement_render_info')) {
            $loader | Should Match ([regex]::Escape('"' + $export + '"'))
            (@($runtimeLock.abi.loaderRequiredExports) -ccontains $export) | Should Be $true
        }
    }

    It 'keeps the reviewed mode ABI as the single admitted Product contract' {
        $terminalProject | Should Not Match 'RSTerminalGhosttyAbi|RS_TERMINAL_GHOSTTY_CANDIDATE_ABI'
        $loaderHeader | Should Not Match 'ghostty_terminal_mode_get|RS_TERMINAL_GHOSTTY_CANDIDATE_ABI'
        $loader | Should Not Match 'ghostty_terminal_mode_get|RS_TERMINAL_GHOSTTY_CANDIDATE_ABI'
        $loader | Should Match '(?s)GhosttyTerminalModeConfig config\{\};\s*config\.mode = mode;\s*const GhosttyResult result = terminalGet\(terminal, GHOSTTY_TERMINAL_DATA_MODE, &config\);'
        $loader | Should Match '(?s)if \(result == GHOSTTY_SUCCESS\)\s*\{\s*value = config\.value;'
        $loader | Should Match 'static_assert\(! HasSizeMember<GhosttyTerminalModeConfig>\)'
        $loader | Should Match 'offsetof\(GhosttyTerminalModeConfig, value\) == sizeof\(GhosttyMode\)'
    }

    It 'initializes every versioned Ghostty record used by the adapter and exercises exact terminal modes' {
        foreach ($record in @(
                'GhosttyStyle'
                'GhosttyGridRef'
                'GhosttySelection'
                'GhosttyTerminalSelectWordOptions'
                'GhosttyTerminalSelectionFormatOptions'
                'GhosttyMouseEncoderSize'
                'GhosttyKittyGraphicsPlacementRenderInfo'
                'GhosttyFormatterTerminalOptions')) {
            $source | Should Match ("$record\s+\w+\{\};[\s\S]{0,160}?\.size\s*=\s*sizeof\(")
        }
        $source | Should Match 'sizeof\(GhosttyKittyPlacementLimits\)'
        foreach ($mode in @(
                'GHOSTTY_MODE_DECCKM'
                'GHOSTTY_MODE_NORMAL_MOUSE'
                'GHOSTTY_MODE_SGR_MOUSE'
                'GHOSTTY_MODE_FOCUS_EVENT'
                'GHOSTTY_MODE_ALT_SCREEN_SAVE'
                'GHOSTTY_MODE_BRACKETED_PASTE')) {
            $source | Should Match $mode
        }
    }

    It 'keeps the admitted regression for DCS payload bytes above 0x7F' {
        $source | Should Match '(?s)dcsHighBytes.*?0xC3u, 0x9Bu.*?== "Y"'
        $source | Should Match 'postlude\.push_back\(static_cast<char>\(0x80u\)\)'
        $source | Should Match 'postlude\.push_back\(static_cast<char>\(0xFEu\)\)'
        $source | Should Match 'postlude\.push_back\(static_cast<char>\(0xFFu\)\)'
    }

    It 'keeps the host ABI size-evolvable with the reviewed weak event callback' {
        $interfaceHeader | Should Match 'struct\s+TerminalOpenContext\s*\{\s*uint32_t\s+sizeBytes'
        $interfaceHeader | Should Match 'interface[\s\S]+ITerminal[\s\S]+Open\(const TerminalOpenContext\* context\)'
        $interfaceHeader | Should Match 'GetChildWindow[\s\S]+SetTheme[\s\S]+UpdateSourceLocation[\s\S]+InsertPath[\s\S]+GetViewState[\s\S]+SetCallback[\s\S]+Close'
        $interfaceHeader | Should Match 'SetCallback\(ITerminalEventCallback\* callback, void\* cookie\) noexcept = 0'
        $interfaceHeader | Should Match 'OnTerminalEvent\(const TerminalEvent\* event, void\* cookie\) noexcept = 0'
        $interfaceHeader | Should Not Match 'OnTerminalEvent\(void\* cookie, const TerminalEvent\* event\)'
        $interfaceHeader | Should Match 'int64_t\s+rootExitObservedTimestampNs'
        $source | Should Match '_rootExitObservedTimestampNs\.store\(SteadyTimestampNs\(\)'
        $source | Should Match 'event\.rootExitObservedTimestampNs\s*=\s*_rootExitObservedTimestampNs\.load'
        $floatingTerminal | Should Match 'terminal\.session\.root_exit_to_host_close_us'
        $interfaceHeader | Should Not Match 'ITerminalCallback|ResolvePendingOpen|RequestClose|RequestRestart|pendingInsertion'
        $interfaceHeader | Should Not Match 'TerminalPhysicalHostKey|TerminalProfileSelectionMode|TerminalOpenPurpose'
    }

    It 'drains floating Terminal payloads before the shared host can consume final teardown' {
        $teardown = $floatingTerminal.IndexOf('if (message == WM_NCDESTROY)', [System.StringComparison]::Ordinal)
        $sharedHost = $floatingTerminal.IndexOf('bool handled = false;', [System.StringComparison]::Ordinal)
        $teardown | Should BeGreaterThan -1
        $sharedHost | Should BeGreaterThan $teardown
        $floatingTerminal | Should Match 'DrainPostedPayloadsForWindow\(hwnd\)'
        $floatingTerminal | Should Match 'terminal\.floating\.retained_payloads'
    }

    It 'uses only the authenticated PowerShell provider for Suggestions history and insertion' {
        $source | Should Match 'Get-PSReadLineOption\)\.HistorySavePath'
        $source | Should Match 'PSConsoleReadLine\]::GetBufferState'
        $source | Should Match 'Microsoft\.PowerShell\.Core\\Get-History -Count 256'
        $source | Should Match 'GetNamedPipeClientProcessId\(pipe\.get\(\), &clientProcessId\)[\s\S]*?clientProcessId != rootProcessId'
        $source | Should Match 'ParseIntegrationRequest\(utf8, request\)'
        $source | Should Match 'PSConsoleReadLine\]::RevertLine\(\)[\s\S]*?PSConsoleReadLine\]::Insert\(\$rsReplacement\)'
        $source | Should Match '_trustedPromptBuffer == _commandSurfacePromptSnapshot'
        $source | Should Not Match '_currentCommandLine|recordSubmittedCommand'
        $commandSurfaceModel | Should Match 'LoadPowerShellHistory\(std::wstring_view trustedPath'
        $commandSurfaceModel | Should Not Match 'GetEnvironmentVariableW|ConsoleHost_history\.txt'
    }

    It 'includes windows.h before bcrypt.h in first-party headers so clang-format cannot alphabetize CNG first' {
        # bcrypt.h needs ULONG/NTSTATUS from windows.h. IncludeBlocks: Preserve plus SortIncludes
        # would put bcrypt.h first inside a shared block, so windows.h must occupy an earlier block.
        $windowsInclude = [regex]::new('(?im)^#include\s+<windows\.h>')
        $bcryptInclude = [regex]::new('(?im)^#include\s+<bcrypt\.h>')
        $headerRoots = @(
            'Common'
            'Plugins'
            'RedSalamander'
            'RedConfigure'
            'RedLauncher'
            'RedSalamanderMonitor'
            'RedSalamanderSearchService'
            'Tests'
        )
        foreach ($rootName in $headerRoots) {
            $root = Join-Path $repoRoot $rootName
            if (-not (Test-Path -LiteralPath $root)) {
                continue
            }
            Get-ChildItem -LiteralPath $root -Recurse -File |
                Where-Object { $_.Extension -eq '.h' -or $_.Extension -eq '.hpp' } |
                ForEach-Object {
                    $text = Get-Content -LiteralPath $_.FullName -Raw
                    $bcryptMatches = @($bcryptInclude.Matches($text))
                    if ($bcryptMatches.Count -eq 0) {
                        return
                    }
                    $windowsMatches = @($windowsInclude.Matches($text))
                    $relative = $_.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
                    foreach ($match in $bcryptMatches) {
                        $prior = $windowsMatches | Where-Object { $_.Index -lt $match.Index } | Select-Object -Last 1
                        if ($null -eq $prior) {
                            throw "$relative includes <bcrypt.h> before <windows.h>."
                        }
                    }
                    $text | Should Match '(?ms)#include <windows.h>\s*\r?\n\s*\r?\n#include <bcrypt.h>'
                }
        }

        $loaderHeader | Should Match '(?ms)#include <windows.h>\s*\r?\n\s*\r?\n#include <bcrypt.h>\s*\r?\n#include <wincodec.h>'
    }

    It 'locks the exact private runtime and Terminal plugin as an atomic generated-hash pair' {
        $loader | Should Match 'CreateFileW\([\s\S]+GENERIC_READ,[\s\r\n]+FILE_SHARE_READ,[\s\S]+FILE_FLAG_OPEN_REPARSE_POINT'
        $loader | Should Not Match 'FILE_SHARE_WRITE|FILE_SHARE_DELETE'
        $loader | Should Match 'GetFileInformationByHandle'
        $loader | Should Match 'FILE_ATTRIBUTE_DIRECTORY\s*\|\s*FILE_ATTRIBUTE_REPARSE_POINT'
        $loader | Should Match 'BCryptOpenAlgorithmProvider[\s\S]+BCRYPT_SHA256_ALGORITHM'
        $loader | Should Match 'fileHeader\.Machine\s*!=\s*kExpectedRuntimeMachine'
        $loader | Should Match 'digestText\s*!=\s*kExpectedRuntimeSha256'
        $loader | Should Match 'terminal_runtime_identity\.generated\.h'
        $loader | Should Match 'RSTerminalRuntimeIdentity::kSizeBytes'
        $loader | Should Match 'RSTerminalRuntimeIdentity::kSha256'
        $loader | Should Match 'GetFinalPathNameByHandleW'
        $loader | Should Match 'LoadLibraryExW\([\s\r\n]*lockedRuntimePath\.c_str\(\)'
        $loader | Should Match '_module\.reset\(\);\s*_runtimeFile\.reset\(\);'
        $loaderHeader | Should Match 'wil::unique_hfile\s+_runtimeFile;\s*wil::unique_hmodule\s+_module;'
        $loader | Should Not Match '[0-9a-f]{64}'
        $runtimeBuilder | Should Match "pairingContract\s*=\s*'generated-raw-sha256-v1'"
        $runtimeBuilder | Should Match 'Get-TerminalRuntimeIdentityHeaderText'
        $runtimeBuilder | Should Match 'dllSha256\s*=\s*\$peIdentity\.sha256'
        $runtimeBuilder | Should Match 'dllSizeBytes\s*=\s*\[string\]\$peIdentity\.sizeBytes'
        $runtimeBuilder | Should Match 'generatedIdentityHeader'
        $runtimeBuilder | Should Match "PSObject\.Properties\['exportCount'\]"
        $runtimeBuilder | Should Match "PSObject\.Properties\['exportSetSha256'\]"
    }

    It 'supports offline reproduction and stages the complete private-runtime notice closure' {
        $terminalSpec | Should Match 'all\s+six notices as one\s+atomic\s+manifest unit'
        $runtimeBuilder | Should Match '\[switch\]\$Offline'
        $runtimeBuilder | Should Match 'Offline verification requires the pinned Zig archive cache'
        $runtimeBuilder | Should Match 'Offline verification requires the pinned Ghostty source checkout'
        $runtimeBuilder | Should Match 'deps\.files\.ghostty\.org serves source tarballs'
        $runtimeBuilder | Should Match 'ArgumentList @\(''fetch'', ''--global-cache-dir'', \$fetchCacheRoot, \$sourceArchive\)'
        $runtimeBuilder | Should Match '\$actualPackageName\s+-cne\s+\$packageName'
        $runtimeBuilder | Should Match '\$dependencyId package content hash mismatch'
        $runtimeBuilder | Should Match '\$dependencyId canonical package archive SHA-256 mismatch'
        $runtimeBuilder | Should Match '\$uucodeSourceArchiveSha256\s*=\s*\[string\]\$uucodeLock\.sourceArchive\.sha256'
        $runtimeBuilder | Should Match '\$uucodeSourceInternalRoot\s*=\s*\[string\]\$uucodeLock\.sourceArchive\.internalRoot'
        $runtimeBuilder | Should Match '\$uucodePackageCacheFileName\s*=\s*\[string\]\$uucodeLock\.contentArchive\.fileName'
        $runtimeBuilder | Should Match '\$uucodePackageInternalRoot\s*=\s*\[string\]\$uucodeLock\.contentArchive\.internalRoot'
        $runtimeBuilder | Should Match '\$uucodeLicenseSha256\s*=\s*\[string\]\$uucodeLock\.licenseSha256'
        foreach ($dependency in @($runtimeLock.dependencies)) {
            [string]$dependency.sourceArchive.sha256 | Should Match '^[0-9a-f]{64}$'
            [string]$dependency.contentArchive.sha256 | Should Match '^[0-9a-f]{64}$'
            [string]$dependency.licenseSha256 | Should Match '^[0-9a-f]{64}$'
        }
        $runtimeBuilder | Should Match 'Assert-RSTerminalEngineArchiveLayout'
        $runtimeBuilder | Should Match 'Get-RSTerminalRuntimeImportModules'
        ([regex]::Matches(
                $runtimeBuilder,
                'Test-RSTerminalRuntimeImportModules')).Count | Should Be 3
        ([regex]::Matches(
                $runtimeBuilder,
                'function\s+Get-TerminalRuntimeIdentityHeaderText')).Count |
            Should Be 1
        $runtimeBuilder | Should Match 'Get-RSTerminalLfTextIdentity'
        $runtimeBuilder | Should Match 'foreach \(\$dependencyState in \$dependencyStates\)'
        $runtimeBuilder | Should Match 'foreach \(\$noticeState in \$noticeStates\)'
        $runtimeBuilder | Should Match 'WriteAllBytes\(\$normalizedPrivateRuntimeOverlayPatch, \$privateRuntimeOverlayPatchBytes\)'
        $attributes | Should Match 'Tools/TerminalEngine/Patches/\*\.patch\s+text\s+eol=lf'
        $attributes | Should Match 'External/TerminalEngine/Notices/\*\.txt\s+text\s+eol=lf'
        $attributes | Should Match 'Specs/Terminal/TerminalEngine\*\s+text\s+eol=lf'
        $attributes | Should Match 'Tools/TerminalEngine/\*\*\s+text\s+eol=lf'
        $attributes | Should Match 'Tools/Tests/TerminalEngine\*\.Tests\.ps1\s+text\s+eol=lf'
        $runtimeBuilder | Should Match '@\(''apply'', ''--check'', \$normalizedPrivateRuntimeOverlayPatch\)'
        $gitApplyLines = [regex]::Matches($runtimeBuilder, "@\('apply',.*")
        $gitApplyLines.Count | Should Be 5
        foreach ($gitApplyArguments in $gitApplyLines) {
            $gitApplyArguments.Value | Should Match '\$normalizedPrivateRuntimeOverlayPatch'
        }
        $runtimeBuilder | Should Match '\$sourceStatus\s*=\s*@\('
        $runtimeBuilder | Should Match '\$sourceStatus\.Count\s*-ne\s*\$allowedStatus\.Count'
        foreach ($notice in @($runtimeLock.notices)) {
            $noticeName = Split-Path ([string]$notice.outputRelativePath) -Leaf
            $runtimeManifest | Should Match ([regex]::Escape("TerminalRuntime\$noticeName"))
        }
        $runtimeBuilder | Should Match 'schemaVersion\s*=\s*5'
        $runtimeBuilder | Should Match 'licenseSha256\s*=\s*\$expectedLicenseHashes'
    }

    It 'freezes the measurement-only VT upgrade corpus before changing the product pin' {
        $vtUpgradeCorpusText | Should Not BeNullOrEmpty
        $vtUpgradeCorpus.formatId | Should Be 'red-salamander-terminal-vt-upgrade-corpus'
        $vtUpgradeCorpus.version | Should Be 1
        $vtUpgradeCorpus.sampling.warmupCount | Should Be 5
        $vtUpgradeCorpus.sampling.sampleCount | Should Be 200
        $vtUpgradeCorpus.sampling.candidateMaximumP95Ratio | Should Be '1.10'
        @($vtUpgradeCorpus.steps.id) | Should Be @(
            'reset-delete',
            'printable-wide',
            'sgr-hyperlink',
            'screen-state',
            'osc52-below',
            'osc52-at',
            'osc52-above',
            'resize-intermediate',
            'kitty-partial',
            'kitty-pending-snapshot',
            'kitty-completion',
            'dcs-high-byte',
            'pty-responses',
            'resize-final',
            'semantic-snapshot',
            'hosted-render')
        @($vtUpgradeCorpus.steps | Where-Object id -like 'osc52-*' | ForEach-Object decodedBytes) |
            Should Be @(262143, 262144, 262145)
        @($vtUpgradeCorpus.expectedDeltas).Count | Should Be 1
        $vtUpgradeCorpus.expectedDeltas[0].id | Should Be 'dcs-high-byte-preservation'

        $source | Should Match 'DebugRunVtUpgradeCorpus'
        $source | Should Match 'kSampleCount\s*=\s*200u'
        $pluginContractTests | Should Match 'terminal\.vt_upgrade\.vt_write_us'
        $pluginContractTests | Should Match 'terminal\.vt_upgrade\.snapshot_us'
        $pluginContractTests | Should Match 'terminal\.vt_upgrade\.hosted_render_us'
        $source | Should Match 'kittyPendingProbeCount'
        $source | Should Match 'peakPrivateBytes'
        $source | Should Match 'ptyOutputWithoutExpectedDeltaSha256'
        $source | Should Match 'snapshotWithoutExpectedDeltaSha256'
        $source | Should Match 'Allocation bytes are an engine storage detail'
        $pluginContractTests | Should Match '--terminal-vt-upgrade-corpus'
        $pluginContractTests | Should Match 'TerminalVtUpgradeEvidence\.schema\.json|red-salamander\.terminal-vt-upgrade-evidence\.v1'

        $example = @{
            schema = 'red-salamander.terminal-vt-upgrade-evidence.v1'
            runId = '20260825_210000_current-pin'
            createdUtc = '2026-08-25T21:00:00.000Z'
            repositoryCommit = '0' * 40
            repositoryBranch = 'codex/ghostty-upgrade'
            machineHash = '0' * 12
            configuration = 'Release'
            buildReceiptId = '0' * 64
            runtimeIdentitySha256 = '0' * 64
            environment = @{
                windowsBuild = '26200'; cpu = 'test'; architecture = 'AMD64'; compiler = '19.51'
                windowsSdk = '10.0.26100.0'; zig = '0.15.1'; vcpkgCommit = '0' * 40
            }
            corpus = @{ formatId = 'red-salamander-terminal-vt-upgrade-corpus'; version = 1; sha256 = '0' * 64; inputByteCount = 1 }
            behavior = @{
                ptyOutputSha256 = '0' * 64; snapshotSha256 = '0' * 64
                withoutExpectedDeltas = @{ ptyOutputSha256 = '0' * 64; snapshotSha256 = '0' * 64 }
                expectedDeltas = @(@{ id = 'dcs-high-byte-preservation'; disposition = 'candidate-only-if-reviewed' })
            }
            measurements = @{
                sampleCount = 200; percentileRule = 'nearest-rank'
                vtWrite = @{ unit = 'us'; p50 = 1; p95 = 1; candidateMaximumP95 = 2; samples = @(1) * 200 }
                snapshot = @{ unit = 'us'; p50 = 1; p95 = 1; candidateMaximumP95 = 2; samples = @(1) * 200 }
                hostedRender = @{ unit = 'us'; p50 = 1; p95 = 1; candidateMaximumP95 = 2; samples = @(1) * 200 }
            }
            kitty = @{ placementCount = 1 }
            memory = @{ peakPrivateBytes = 1 }
            assertionCount = 1
            outcome = 'pass'
        } | ConvertTo-Json -Depth 8
        Test-Json -Json $example -SchemaFile $vtUpgradeEvidenceSchemaPath | Should Be $true
    }
}
