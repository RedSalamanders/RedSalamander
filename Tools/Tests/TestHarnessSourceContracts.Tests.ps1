Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
}

Describe 'IMAP UIDVALIDITY and single-message delete source contracts' {
    BeforeAll {
        $imapSource = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.Imap.cpp'
        $helpersHeader = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.ImapHelpers.h'
        $helpersSource = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.ImapHelpers.cpp'
    }

    It 'never issues mailbox-wide EXPUNGE and requires UIDPLUS before the mark-expunge-rollback state machine' {
        $imapSource | Should Not Match '(?s)CurlPerformImapCustomRequest\([^;]{0,500},\s*"EXPUNGE"\s*,'
        $imapSource | Should Match 'ImapFetchCapabilities\(conn, capabilities\)'
        $imapSource | Should Match 'ExecuteImapSingleMessageDelete\(capabilities\.uidPlus'
        $imapSource | Should Match 'ImapDeleteCommand::UidExpunge'
        $imapSource | Should Match 'ImapDeleteCommand::RemoveDeletedFlag'
        $helpersSource | Should Match 'if \(! uidPlusAvailable\)[\s\S]{0,120}ERROR_NOT_SUPPORTED'
        $helpersSource | Should Match 'rollbackHr\s*=\s*executor\(context, ImapDeleteCommand::RemoveDeletedFlag'
    }

    It 'carries UIDVALIDITY in listed identity and revalidates it before message fetch or delete' {
        $helpersHeader | Should Match 'struct ImapMessageIdentity[\s\S]*?uidValidity[\s\S]*?uid'
        $imapSource | Should Match 'BuildImapMessageLeafName\(meta\.subject, meta\.from, uidValidity, uid\)'
        $imapSource | Should Match 'TryParseImapMessageIdentityFromLeafName\(leaf, identity\)'
        ([regex]::Matches($imapSource, 'ImapValidateMessageUidValidity\(conn, mailboxPath, delimiter, identity\.uidValidity\)')).Count | Should BeGreaterThan 2
        $helpersSource | Should Match 'ERROR_REVISION_MISMATCH'

        $listing = [regex]::Match(
            $imapSource,
            '(?s)HRESULT\s+ImapReadDirectoryEntries\(.*?(?=HRESULT\s+ReadDirectoryEntries\()').Value
        $listing.Length | Should BeGreaterThan 0
        ([regex]::Matches($listing, 'ImapFetchMailboxStatus\(')).Count | Should Be 1
        $listing.IndexOf('ImapFetchMailboxStatus(') | Should BeLessThan $listing.IndexOf('ImapListMessageUids(')
    }
}

Describe 'Microsoft Drive credential-boundary source contracts' {
    BeforeAll {
        $source = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
    }

    It 'validates Graph and preauthenticated upload targets before credential-bound dispatch' {
        $source | Should Match 'struct ValidatedGraphApiUrl'
        $source | Should Match 'struct ValidatedPreauthenticatedUploadUrl'
        $source | Should Match 'ValidateGraphApiUrl\(rawUrl, graphUrl\)'
        $source | Should Match 'ValidatePreauthenticatedUploadUrl\(rawUploadUrl, uploadUrlOut\)'
        $source | Should Match 'SendPreauthenticatedUploadRequest'
        $source | Should Match 'WINHTTP_DISABLE_REDIRECTS'
        $source | Should Match 'Common::Paging::WideContinuationGuard pager'
        $source | Should Match 'pager\.BeginContinuation\(nextUrl'
        $source | Should Match 'foreign Graph continuation should fail before a second request'
        $source | Should Match 'repeated Graph continuation should fail before a third request'
    }

    It 'keeps bearer values and opaque query values out of diagnostics and serialized header blocks' {
        $source | Should Match 'DescribeHttpRequestTarget'
        $source | Should Match 'SecureClear\(authorizationHeader\)'
        $source | Should Match 'WinHttpAddRequestHeaders'
        $source | Should Not Match 'headerBlock\.append\(L"Authorization:'
        $source | Should Not Match 'headers=''\{\}'''
        $source | Should Not Match 'url=''\{\}'''
        $source | Should Match 'captured diagnostics must not contain the bearer sentinel'
        $source | Should Match 'captured diagnostics must not contain Graph or upload-session query sentinels'
    }
}

function Get-RSTestSourceContractFiles {
    $extensions = @('.cpp', '.h', '.hpp', '.ps1')
    $roots = @(
        'RedSalamander\SelfTest',
        'Tests',
        'Tools\Tests'
    )

    foreach ($root in $roots) {
        Get-ChildItem -LiteralPath (Join-Path $repoRoot $root) -Recurse -File |
            Where-Object { $extensions -contains $_.Extension } |
            Where-Object { -not $_.FullName.EndsWith('Tools\Tests\TestHarnessSourceContracts.Tests.ps1', [StringComparison]::OrdinalIgnoreCase) }
    }
}

function Get-RSRelativeTestSourcePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $getRelativePath = [System.IO.Path].GetMethod('GetRelativePath', [type[]]@([string], [string]))
    if ($null -ne $getRelativePath) {
        return [System.IO.Path]::GetRelativePath($Root, $Path)
    }

    $rootFull = [System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    $rootUri = [Uri]::new($rootFull)
    $pathUri = [Uri]::new($pathFull)
    return [Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('/', '\')
}

Describe 'Test harness source contracts' {
    It 'rejects unknown DxUiTests switches instead of silently running every suite' {
        $source = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.cpp'

        $source | Should Match 'Unknown argument'
        $source | Should Match 'arg\[0\]\s*==\s*L''-'''
    }

    It 'rejects empty DxUiTests option values with targeted diagnostics' {
        $source = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.cpp'

        $source | Should Match 'Missing suite name'
        $source | Should Match 'Missing perf JSONL path'
    }

    It 'documents and enforces bounded self-test timeout multipliers' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $source | Should Match 'kSelfTestTimeoutMultiplierMin'
        $source | Should Match 'kSelfTestTimeoutMultiplierMin\s*=\s*0\.1'
        $source | Should Match 'kSelfTestTimeoutMultiplierMax'
        $source | Should Match 'kSelfTestTimeoutMultiplierMax\s*=\s*100\.0'
        $source | Should Match 'std::isfinite'
        $source | Should Match 'std::clamp\(parsed,\s*kSelfTestTimeoutMultiplierMin,\s*kSelfTestTimeoutMultiplierMax\)'
        $source | Should Match 'Clamped --selftest-timeout-multiplier'
        $source | Should Match 'Invalid --selftest-timeout-multiplier'
    }

    It 'keeps scaled self-test timeouts finite and nonzero for nonzero bases' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'std::isfinite\(scaled\)'
        $source | Should Match 'baseMs\s*>\s*0u'
        $source | Should Match 'return 1u;'
    }

    It 'waits for runner-native self-test processes by direct process handle' {
        $source = Get-RSText -Path 'Tools\Run-AllTests.ps1'

        $source | Should Match 'function Invoke-RSSelfTestProcess'
        $source | Should Match '\[System\.Diagnostics\.ProcessStartInfo\]::new\(\)'
        $source | Should Match '\$startInfo\.UseShellExecute = \$false'
        $source | Should Match '\$process\.WaitForExit\(\)'
        $source | Should Not Match 'Start-Process -FilePath \$sc\.Path'
    }

    It 'keeps the nightly shuffle repeat lane separate from the PR gate' {
        $nightly = Get-RSText -Path '.github\workflows\nightly-flake.yml'
        $ci = Get-RSText -Path '.github\workflows\ci.yml'

        $nightly | Should Match 'schedule:'
        $nightly | Should Match 'workflow_dispatch:'
        $nightly | Should Match 'REDSALAMANDER_TEST_ROOT'
        $nightly | Should Match 'Run-AllTests\.ps1\s+-Suite\s+All'
        $nightly | Should Match '-SelfTestRepeat\s+5'
        $nightly | Should Match '-SelfTestShuffleSeed\s+\$seed'
        $nightly | Should Match '-ClassifyFailures'
        $nightly | Should Match 'selftest-artifacts-nightly-shuffle'
        $ci | Should Not Match '-SelfTestRepeat\s+5'
        $ci | Should Not Match 'selftest-artifacts-nightly-shuffle'
    }

    It 'gives ViewerPETests nested shell churn its own timeout budget' {
        $source = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $source | Should Match 'kViewerHarnessDefaultTimeout\s*=\s*120000ms'
        $source | Should Match 'kViewerShellComboLongRunTimeout\s*=\s*600000ms'
        $source | Should Match 'kViewerShellComboLongRunTimeout\s*>\s*kViewerHarnessDefaultTimeout'
        $source | Should Match 'TestViewerShellComboHostsLongRunOpenCloseStayStable",\s*kViewerShellComboLongRunTimeout'
    }

    It 'keeps ViewerVLC dependency loading isolated from the process DLL directory' {
        $source = Get-RSText -Path 'Plugins\ViewerVLC\ViewerVLC.cpp'

        $source | Should Not Match 'SetDllDirectoryW|GetDllDirectoryW|AddDllDirectory|RemoveDllDirectory'
        $source | Should Not Match 'SearchPathW'
        $source | Should Not Match 'GetEnvironmentVariableW'
        $source | Should Match 'FOLDERID_ProgramFiles'
        $source | Should Match 'FOLDERID_ProgramFilesX86'
        $source | Should Match 'kVlcModuleLoadFlags\s*=\s*LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR\s*\|\s*LOAD_LIBRARY_SEARCH_DEFAULT_DIRS'
        ([regex]::Matches($source, 'LoadLibraryExW\(dllPath\.c_str\(\),\s*nullptr,\s*kVlcModuleLoadFlags\)')).Count | Should Be 1
    }

    It 'keeps ViewerVLC probing, loader fallback, drawable teardown, and option boundaries safe' {
        $source = Get-RSText -Path 'Plugins\ViewerVLC\ViewerVLC.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'
        $commands = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        $startPlayback = [regex]::Match($source, 'bool\s+ViewerVLC::StartPlayback\([\s\S]*?\n\}')
        $submitFailure = [regex]::Match($source, 'if\s*\(queued\s*==\s*0\)[\s\S]{0,1800}?\n\s*\}')

        $startPlayback.Success | Should Be $true
        $startPlayback.Value | Should Not Match 'AutoDetectVlcInstallDir|IsVlcInstallDir|SearchPathW|LoadLibraryExW|LoadVlcState|filesystem::exists|filesystem::is_regular_file'
        $submitFailure.Success | Should Be $true
        $submitFailure.Value | Should Not Match 'LoadLibraryExW|LoadVlcState|libvlc_new'
        $source | Should Match 'work->config\s*=\s*_config'
        $source | Should Match 'BuildVlcLoadSpec\(work->config'
        $source | Should Match 'IDS_VIEWERVLC_DETAILS_ASYNC_SUBMIT_FAILED'
        $source | Should Match 'result->windowIdentity\s*==\s*_windowIdentity'
        ([regex]::Matches($source, 'result->moduleKeepAlive\s*=\s*AcquireModuleReferenceFromAddress')).Count | Should Be 2
        $source | Should Match 'CreateThreadpoolWork\(&ViewerVLC::VlcCleanupDispatcherCallback'
        $source | Should Match 'SubmitThreadpoolWork\(_cleanupDispatcherWork\.get\(\)\)'
        $source | Should Match 'SendMessageTimeoutW\(hwnd[\s\S]{0,300}kViewerVlcAsyncFallbackReady'
        $source | Should Match 'key\.size\(\)\s*>\s*3u\s*&&\s*EqualsAsciiIgnoreCase\(key\.substr\(0u,\s*3u\),\s*"no-"\)[\s\S]{0,160}key\.remove_prefix\(3u\)'
        $source | Should Match '"logfile"[\s\S]{0,100}"file-logging"[\s\S]{0,100}"pidfile"'
        $source | Should Not Match '~VlcAsyncLoadResult\(\)[\s\S]{0,300}libvlc_media_player_stop'
        $source | Should Match 'RecordAsyncFallbackCompletion'
        $source | Should Match 'kViewerVlcAsyncFallbackReady'
        $source | Should Match 'load-post-fallback-terminal'
        $source | Should Match 'cleanup-post-fallback-terminal'
        $source | Should Not Match 'libvlc_media_player_set_hwnd\([^\r\n]*nullptr'
        $source | Should Match 'ShowWindow\(_hVideo\.get\(\),\s*SW_HIDE\)'
        $source | Should Match 'work->state\s*=\s*std::move\(state\)'
        $source | Should Match '_pendingCloseCleanupCount\s*!=\s*0'
        $source | Should Match '_destroyingAfterCleanup\s*=\s*true[\s\S]{0,200}_hWnd\.reset\(\)'
        $source | Should Match 'IsValidVlcModuleToken'
        $source | Should Match 'IsDangerousVlcExtraOption'
        $tests | Should Match 'loader-submit failure never runs loader work synchronously'
        $tests | Should Match 'retains parent and video HWND identities'
        $tests | Should Match 'assigns a new identity to every reopened window'
        $tests | Should Match 'load-completion post failure reaches the identity-bound terminal UI fallback'
        $tests | Should Match 'close-completion post failure drained through the identity-bound UI fallback'
        $tests | Should Match 'pre-created dispatcher drains cleanup after an injected allocation failure'
        $tests | Should Match 'pre-created dispatcher drains cleanup after an injected submit failure'
        $tests | Should Match 'forced parent destruction does not wait for delayed cleanup'
        $tests | Should Match 'unloads after the HWND-independent dispatcher completes stop/release'
        $commands | Should Match 'retirementDeadline\s*=\s*std::chrono::steady_clock::now\(\)\s*\+\s*SelfTest::Scale\(30''000ms\)'
        $commands | Should Match 'retirementVlcSnapshot\.cleanupCompletions'
        $commands | Should Match 'retirementVlcSnapshot\.deferredCleanupCount'
    }

    It 'bounds multi-plugin metadata before host indexing and reference-counts shared localization owners' {
        $factoryHeader = Get-RSText -Path 'Common\PlugInterfaces\Factory.h'
        $factoryImpl = Get-RSText -Path 'Common\PlugInterfaces\FactoryImpl.h'
        $viewerManager = Get-RSText -Path 'RedSalamander\ViewerPluginManager.cpp'
        $fileSystemManager = Get-RSText -Path 'RedSalamander\FileSystemPluginManager.cpp'
        $localization = Get-RSText -Path 'Common\Common\LocalizationManager.cpp'
        $pluginTests = Get-RSText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'
        $localizationTests = Get-RSText -Path 'Tests\LocalizationTests\LocalizationTests.cpp'

        $factoryHeader | Should Match 'kMaxEnumeratedPluginsPerModule\s*=\s*256u'
        $factoryHeader | Should Match 'return\s+metaData\s*!=\s*nullptr\s*&&\s*IsValidEnumeratedPluginCount\(count\)'
        $factoryImpl | Should Match 'entries\.size\(\)\s*>\s*kMaxEnumeratedPluginsPerModule'
        $factoryImpl | Should Match 'reinterpret_cast<std::uintptr_t>\(currentMetaData\)\s*!=\s*expectedAddress'

        foreach ($manager in @($viewerManager, $fileSystemManager)) {
            $manager | Should Match 'std::filesystem::directory_iterator\s+item\(optionalDir,\s*ec\)'
            $manager | Should Match 'while\s*\(!\s*ec\s*&&\s*item\s*!=\s*end\)'
            $manager | Should Match 'item\.increment\(ec\)'
            $manager | Should Not Match 'for\s*\([^\r\n]*std::filesystem::directory_iterator\(optionalDir,\s*ec\)'
            $manager | Should Match 'IsValidEnumeratedPluginRange\(metaData,\s*count\)'
        }

        $fileSystemManager | Should Match 'if\s*\(!\s*IsValidEnumeratedPluginRange\(metaData,\s*count\)\)[\s\S]{0,700}result\.entries\.reserve\(count\)'
        $viewerManager | Should Match 'if\s*\(!\s*IsValidEnumeratedPluginRange\(metaData,\s*count\)\)[\s\S]{0,900}for\s*\(unsigned int i\s*=\s*0;\s*i\s*<\s*count;\s*\+\+i\)'

        $localization | Should Match 'size_t\s+registrationCount\s*=\s*1u'
        $localization | Should Match 'existing->second\.registrationCount\s*==\s*\(std::numeric_limits<size_t>::max\)\(\)'
        $localization | Should Match '\+\+existing->second\.registrationCount'
        $localization | Should Match 'existing->second\.registrationCount\s*>\s*1u[\s\S]{0,120}--existing->second\.registrationCount'
        $localization | Should Match 'OrdinalString::EqualsNoCase\(existing->second\.moduleName,\s*ownerName\)'

        $pluginTests | Should Match 'IsValidEnumeratedPluginCount\(kMaxEnumeratedPluginsPerModule\)'
        $pluginTests | Should Match 'factory enumeration rejects a producer count above the ABI maximum'
        $pluginTests | Should Match 'factory enumeration rejects non-contiguous producer metadata'
        $localizationTests | Should Match 'TestNestedResourceOwnerRegistrationKeepsSatelliteUntilLastRelease'
        $localizationTests | Should Match 'first unregister keeps the shared module resource owner registered'
    }

    It 'keeps viewer threadpool module pins alive through the callback return boundary' {
        $callbackFiles = @(
            @{ Path = 'Plugins\ViewerPE\ViewerPE.cpp'; ExpectedCallbacks = 1 },
            @{ Path = 'Plugins\ViewerWeb\ViewerWeb.cpp'; ExpectedCallbacks = 2 },
            @{ Path = 'Plugins\ViewerImgRaw\ViewerImgRaw.Decode.cpp'; ExpectedCallbacks = 2 },
            @{ Path = 'Plugins\ViewerImgRaw\ViewerImgRaw.Export.cpp'; ExpectedCallbacks = 1 },
            @{ Path = 'Plugins\ViewerText\ViewerText.cpp'; ExpectedCallbacks = 1 },
            @{ Path = 'Plugins\ViewerText\ViewerText.Text.cpp'; ExpectedCallbacks = 1 },
            @{ Path = 'Plugins\ViewerVLC\ViewerVLC.cpp'; ExpectedCallbacks = 2 },
            @{ Path = 'Plugins\ViewerSqlite\ViewerSqlite.cpp'; ExpectedCallbacks = 1 }
        )

        foreach ($callbackFile in $callbackFiles) {
            $source = Get-RSText -Path $callbackFile.Path
            $submitCount = ([regex]::Matches($source, 'TrySubmitThreadpoolCallback\s*\(')).Count
            $returnTransferCount = ([regex]::Matches(
                    $source,
                    'TransferModulePinToCallbackReturn\s*\(\s*instance\s*,\s*[A-Za-z_][A-Za-z0-9_]*->moduleKeepAlive\s*\)')).Count

            $submitCount | Should Be $callbackFile.ExpectedCallbacks
            $returnTransferCount | Should Be $callbackFile.ExpectedCallbacks
        }

        $helpersSource = Get-RSText -Path 'Common\Helpers.h'
        $helpersSource | Should Match 'TransferModulePinToCallbackReturn[\s\S]{0,400}FreeLibraryWhenCallbackReturns\(instance,\s*modulePin\.release\(\)\)'

        $vlcSource = Get-RSText -Path 'Plugins\ViewerVLC\ViewerVLC.cpp'
        $vlcSource | Should Match 'VlcCleanupDispatcherCallback[\s\S]{0,800}TransferModulePinToCallbackReturn\(instance,\s*callbackPin\)'
    }

    It 'rejects viewer async work without module pins and impossible provider read counts' {
        $web = Get-RSText -Path 'Plugins\ViewerWeb\ViewerWeb.cpp'
        $imgRawDecode = Get-RSText -Path 'Plugins\ViewerImgRaw\ViewerImgRaw.Decode.cpp'
        $pe = Get-RSText -Path 'Plugins\ViewerPE\ViewerPE.cpp'

        ([regex]::Matches($web, 'AcquireModuleReferenceFromAddress\(&kViewerWebModuleAnchor\)')).Count | Should Be 2
        $web | Should Match 'ctx->moduleKeepAlive\s*=\s*AcquireModuleReferenceFromAddress\(&kViewerWebModuleAnchor\);\s*if\s*\(!\s*ctx->moduleKeepAlive\)'
        $web | Should Match 'work->moduleKeepAlive\s*=\s*AcquireModuleReferenceFromAddress\(&kViewerWebModuleAnchor\);\s*if\s*\(!\s*work->moduleKeepAlive\)'

        ([regex]::Matches($imgRawDecode, 'AcquireModuleReferenceFromAddress\(&kViewerImgRawModuleAnchor\)')).Count | Should Be 2
        ([regex]::Matches($imgRawDecode, 'if\s*\(!\s*ctx->moduleKeepAlive\)')).Count | Should Be 2
        $imgRawDecode | Should Match 'Failed to pin the plugin module for async open'
        $imgRawDecode | Should Match 'failAsyncSubmission\(\)'
        $pe | Should Match 'if\s*\(read\s*>\s*want\)\s*\{[\s\S]{0,240}ERROR_INVALID_DATA'
    }

    It 'keeps ViewerImgRaw main decode scheduling bounded, exact, terminal, and ordered' {
        $source = Get-RSText -Path 'Plugins\ViewerImgRaw\ViewerImgRaw.Decode.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $source | Should Match 'struct\s+ViewerImgRaw::AsyncOpenSchedulerState[\s\S]{0,900}std::unique_ptr<AsyncOpenRequest>\s+pending'
        $source | Should Match 'bool\s+workerActive\s*=\s*false'
        $source | Should Match 'uint64_t\s+replacedPendingCount\s*=\s*0u'
        $source | Should Match 'viewer\.imgraw\.open\.queue\.replaced_count'
        $source | Should Match 'scheduler->pending\s*=\s*std::move\(request\)'
        $source | Should Match 'reader->Seek\(0,\s*FILE_BEGIN,\s*&pos\)'
        $source | Should Match 'FAILED\(seekHr\)\s*\|\|\s*pos\s*!=\s*0u'
        $source | Should Match 'if\s*\(got\s*==\s*0\)[\s\S]{0,220}ERROR_HANDLE_EOF'
        $source | Should Match 'trailingRead\s*!=\s*0u'
        $source | Should Match 'terminalFallbackPending\.store\(true,\s*std::memory_order_release\)'
        $source | Should Match 'PollAsyncOpenTerminalFallback'
        $source | Should Match '_debugLastPreviewApplyOrdinal\s*=\s*applyOrdinal'
        $source | Should Match '_debugLastFinalApplyOrdinal\s*=\s*applyOrdinal'

        ([regex]::Matches($tests, 'TestViewerImgRawLatestWinsExactReaderAndCloseSafety')).Count | Should Be 4
        ([regex]::Matches($tests, 'TestViewerImgRawEmbeddedThumbnailTerminalSequencing')).Count | Should Be 4
        $tests | Should Match 'lastPreviewApplyOrdinal\s*<\s*resource\.lastFinalApplyOrdinal'
        $tests | Should Match 'blockedViewer->SetCallback\(nullptr,\s*nullptr\)'
        $tests | Should Match 'callback-return module pin unloads only after blocked work returns'
    }

    It 'keeps ViewerWeb security setup fail-closed and teardown outside DllMain' {
        $source = Get-RSText -Path 'Plugins\ViewerWeb\ViewerWeb.cpp'
        $header = Get-RSText -Path 'Plugins\ViewerWeb\ViewerWeb.h'
        $dllMain = Get-RSText -Path 'Plugins\ViewerWeb\dllmain.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $header | Should Match '\[\[nodiscard\]\]\s+HRESULT\s+ConfigureWebViewSettings\(\)\s+noexcept'
        $source | Should Match 'HRESULT\s+ViewerWeb::ConfigureWebViewSettings\(\)\s+noexcept'
        $source | Should Match 'HRESULT\s+hr\s*=\s*settings->put_IsScriptEnabled'
        $source | Should Match 'hr\s*=\s*settings->put_IsWebMessageEnabled'
        $source | Should Match 'hr\s*=\s*settings->put_AreDevToolsEnabled'
        $source | Should Not Match 'static_cast<void>\(settings->put_(IsScriptEnabled|IsWebMessageEnabled|AreDevToolsEnabled)'

        $source | Should Match 'securityHr\s*=\s*_webView->add_NavigationStarting'
        $source | Should Match 'securityHr\s*=\s*_webView->add_FrameNavigationStarting'
        $source | Should Match 'securityHr\s*=\s*_webView->add_NewWindowRequested'
        $source | Should Match 'securityHr\s*=\s*_webView->add_WebResourceRequested'
        $source | Should Match 'securityHr\s*=\s*[\r\n\s]*_webView->AddWebResourceRequestedFilter'
        $source | Should Match 'if\s*\(FAILED\(securityHr\)\)\s*\{\s*return failSecuritySetup\(securityHr\)'
        $source | Should Match 'failInternalRequest[\s\S]{0,500}DiscardWebView2\(\)'
        $source | Should Match 'FAILED\(installResponseHr\)\s*\?\s*failInternalRequest\(installResponseHr\)'
        $source | Should Match 'const HRESULT cancelHr\s*=\s*args->put_Cancel\(TRUE\)'
        $source | Should Match 'const HRESULT handledHr\s*=\s*args->put_Handled\(TRUE\)'
        $source | Should Match 'if\s*\(FAILED\(cancelHr\)\)[\s\S]{0,700}DiscardWebView2\(\)'
        $source | Should Match 'if\s*\(FAILED\(handledHr\)\)[\s\S]{0,700}DiscardWebView2\(\)'

        ([regex]::Matches($source, 'reader->Read\(')).Count | Should Be 2
        ([regex]::Matches($source, 'ViewerWebSecurity::IsProviderReadCountValid\(')).Count | Should Be 2
        ([regex]::Matches($source, '(?:const HRESULT readHr|const HRESULT copyHr)\s*=\s*ReadProviderExactly\(')).Count | Should Be 4
        $source | Should Match 'g_activeAsyncWorkerCount\.fetch_add\(1u,\s*std::memory_order_acq_rel\)'
        $source | Should Match 'g_activeAsyncWorkerCount\.fetch_sub\(1u,\s*std::memory_order_acq_rel\)'
        $source | Should Match 'g_liveComCallbackCount\.fetch_add\(1u,\s*std::memory_order_relaxed\)'
        $source | Should Match 'g_comCallbackReleaseEpoch\.fetch_add\(1u,\s*std::memory_order_release\)[\s\S]{0,160}g_liveComCallbackCount\.fetch_sub'
        $source | Should Match 'callbackOwnerThread\s*!=\s*0u\s*&&\s*callbackOwnerThread\s*!=\s*GetCurrentThreadId\(\)'
        $source | Should Match 'if\s*\(releaseEpoch\s*!=\s*observedEpoch\)[\s\S]{0,220}return false'
        $dllMain | Should Not Match 'reason\s*==\s*DLL_PROCESS_DETACH'
        $dllMain | Should Not Match 'ResetSharedEnvironment\s*\('
        $tests | Should Match 'ViewerWeb provider Read failure reaches a terminal UI state'
        $tests | Should Match 'ViewerWeb locked destination preserves every pre-existing byte'
        $tests | Should Match 'ViewerWeb callback release epoch requires a later quiescent unload observation'
        $tests | Should Match 'ViewerWeb Save As worker keeps its DLL mapped after the caller releases the module handle'
        $tests | Should Match 'ViewerWeb module unloads only after the blocked Save As callback returns'
    }

    It 'keeps ViewerPE reader commitments exact and current terminal delivery recoverable' {
        $source = Get-RSText -Path 'Plugins\ViewerPE\ViewerPE.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $source | Should Match 'reader->Seek\(0,\s*FILE_BEGIN,\s*&position\)'
        $source | Should Match 'FAILED\(hr\)\s*\|\|\s*position\s*!=\s*0u'
        $source | Should Match 'if\s*\(read\s*==\s*0\)[\s\S]{0,220}ERROR_HANDLE_EOF'
        $source | Should Match 'if\s*\(offset\s*!=\s*bytes\.size\(\)\)'
        $source | Should Match 'trailingRead\s*>\s*1u\s*\|\|\s*trailingRead\s*!=\s*0u'
        $source | Should Match 'terminalFallbackRequestId\.store\(requestId,\s*std::memory_order_release\)'
        $source | Should Match 'SetTimer\(hwnd,\s*kAsyncParseFallbackTimerId'
        $source | Should Match 'PollAsyncParseTerminalFallback'
        $tests | Should Match 'ViewerPeDebugAsyncFault::ResultAllocation'
        $tests | Should Match 'ViewerPeDebugAsyncFault::PayloadPost'
        $tests | Should Match 'rejects premature EOF before the committed GetSize byte count'
        $tests | Should Match 'rejects bytes beyond the committed GetSize byte count'
        $tests | Should Match 'rejects Read counts larger than the requested buffer'
    }

    It 'keeps ViewerSpace close non-waiting with an honest runtime unload and shutdown gate' {
        $source = Get-RSText -Path 'Plugins\ViewerSpace\ViewerSpace.cpp'
        $header = Get-RSText -Path 'Plugins\ViewerSpace\ViewerSpace.h'
        $factory = Get-RSText -Path 'Plugins\ViewerSpace\Factory.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $header | Should Match 'void\s+AbandonScanWorkers\(\)\s+noexcept'
        $source | Should Not Match 'CancelScanAndWait|ReapFinishedScanWorkers\(bool\s+wait\)'
        $source | Should Match 'worker\.done[\s\S]{0,180}worker\.thread\.join\(\)[\s\S]{0,180}g_scanWorkerUnloadGates\.fetch_sub'
        $source | Should Match 'worker\.thread\.detach\(\)[\s\S]{0,300}abandonedWorkers'
        $source | Should Match 'g_scanWorkerUnloadGates\.fetch_add\(1u'
        $source | Should Match 'CanUnloadViewerSpaceModuleNow[\s\S]{0,180}g_scanWorkerUnloadGates\.load'
        $factory | Should Match 'RedSalamanderPluginCanUnloadNow[\s\S]{0,180}CanUnloadViewerSpaceModuleNow'
        $source | Should Match 'g_scanSchedulerShutdown\s*=\s*true'
        $source | Should Match '_shutdownRequested\.store\(true,\s*std::memory_order_release\)'
        ([regex]::Matches($source, 'update\.generation\s*!=\s*_scanGeneration\.load')).Count | Should BeGreaterThan 1
        $tests | Should Match 'TestViewerSpaceBlockedProviderCloseAndPostUpdateRace'
        $tests | Should Match 'kViewerSpaceDebugPauseNextPostUpdate'
        $tests | Should Match 'runtime unload gate rejects refresh while an abandoned worker exists'
    }

    It 'captures viewer payload fields before moving the payload into PostMessagePayload' {
        $viewerSources = Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Plugins') -Recurse -File -Filter 'Viewer*.cpp'
        $unsafeSiblingArgumentPattern = [regex]::new(
            'PostMessagePayload\([^;]{0,500}(?<payload>[A-Za-z_][A-Za-z0-9_]*)->[A-Za-z_][A-Za-z0-9_]*[^;]{0,500}std::move\(\k<payload>\)',
            [System.Text.RegularExpressions.RegexOptions]::Singleline)

        foreach ($viewerSource in $viewerSources) {
            $source = Get-Content -LiteralPath $viewerSource.FullName -Raw
            $unsafeSiblingArgumentPattern.IsMatch($source) | Should Be $false
        }
    }

    It 'keeps ViewerText Save As source-first and transactionally committed' {
        $source = Get-RSText -Path 'Plugins\ViewerText\ViewerText.cpp'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'
        $saveMatch = [regex]::Match($source, 'HRESULT\s+ViewerText::SaveAsToPath[\s\S]*?\n\}\r?\n\r?\nvoid\s+ViewerText::CommandSaveAs')

        $saveMatch.Success | Should Be $true
        $body = $saveMatch.Value
        $body | Should Not Match 'CREATE_ALWAYS'
        $body | Should Match '_isLoading[\s\S]{0,120}ERROR_BUSY'
        $body | Should Match '_textStreamActive[\s\S]{0,160}ERROR_NOT_SUPPORTED'
        $body | Should Match 'CreateSiblingSaveTemp'
        $body | Should Match 'sourceReader->GetSize\(&expectedSourceBytes\)'
        $body | Should Match 'copiedSourceBytes\s*!=\s*expectedSourceBytes[\s\S]{0,120}ERROR_HANDLE_EOF'
        $body | Should Match 'FlushFileBuffers\(tempFile\.get\(\)\)'
        $body | Should Match 'MoveFileExW\(tempPath\.c_str\(\),\s*normalizedDestination\.c_str\(\),\s*MOVEFILE_REPLACE_EXISTING\s*\|\s*MOVEFILE_WRITE_THROUGH\)'
        $readerOpenIndex = $body.IndexOf('CreateFileReader(_currentPath.c_str(), sourceReader.put())')
        $readerSeekIndex = $body.IndexOf('ViewerTextSafety::SeekExact(sourceReader.get(), 0u)')
        $tempCreateIndex = $body.IndexOf('CreateSiblingSaveTemp')
        $readerOpenIndex | Should BeGreaterThan -1
        $readerSeekIndex | Should BeGreaterThan -1
        $tempCreateIndex | Should BeGreaterThan -1
        $readerOpenIndex | Should BeLessThan $tempCreateIndex
        $readerSeekIndex | Should BeLessThan $tempCreateIndex

        $tests | Should Match 'TestViewerTextSaveAsPreservesDataOnFailures'
        $tests | Should Match 'ViewerTextDebugSaveFault::SourceOpen'
        $tests | Should Match 'ViewerTextDebugSaveFault::SourceRead'
        $tests | Should Match 'ViewerTextDebugSaveFault::Encode'
        $tests | Should Match 'ViewerTextDebugSaveFault::Write'
        $tests | Should Match 'ViewerTextDebugSaveFault::Flush'
        $tests | Should Match 'ViewerTextDebugSaveFault::Commit'
    }

    It 'keeps ViewerText async decode and clipboard work terminal and bounded' {
        $openSource = Get-RSText -Path 'Plugins\ViewerText\ViewerText.cpp'
        $textSource = Get-RSText -Path 'Plugins\ViewerText\ViewerText.Text.cpp'
        $hexSource = Get-RSText -Path 'Plugins\ViewerText\ViewerText.Hex.cpp'
        $helpers = Get-RSText -Path 'Plugins\ViewerText\ViewerText.SafetyHelpers.h'
        $header = Get-RSText -Path 'Plugins\ViewerText\ViewerText.h'
        $tests = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'

        $openSource | Should Match 'auto postTerminal = wil::scope_exit'
        $openSource | Should Match 'PostMessagePayload\(hwnd,\s*kAsyncOpenCompleteMessage,\s*static_cast<WPARAM>\(requestId\)'
        $openSource | Should Match 'SendMessageTimeoutW\(hwnd,[\s\S]{0,120}kAsyncOpenFailureMessage'
        $openSource | Should Match 'OnAsyncOpenFailure\(requestId,\s*E_OUTOFMEMORY\)'
        $openSource | Should Match 'OnAsyncOpenFailure\(requestId,\s*HRESULT_FROM_WIN32\(ERROR_NOT_ENOUGH_MEMORY\)\)'
        $openSource | Should Not Match 'falling back to HEX view|allowHexFallback'
        $openSource | Should Not Match 'MultiByteToWideChar failed[^\r\n]+result->hr'
        $textSource | Should Match 'const uint64_t terminalRequestId = result->requestId;[\s\S]{0,420}PostMessagePayload\(work->hwnd,\s*WndMsg::kViewerTextAsyncStreamComplete,\s*static_cast<WPARAM>\(terminalRequestId\),\s*std::move\(result\)\)'
        $textSource | Should Not Match 'PostMessagePayload\([^;]{0,300}result->requestId[^;]{0,300}std::move\(result\)'

        $helpers | Should Match 'kShiftJisCodePage\s*=\s*932u'
        $helpers | Should Match 'kGbkCodePage\s*=\s*936u'
        $helpers | Should Match 'kBig5CodePage\s*=\s*950u'
        $helpers | Should Match 'IncompleteDbcsTailSize'
        $openSource | Should Match 'IncompleteDbcsTailSize\(bytes\.data\(\),\s*bytes\.size\(\),\s*displayCodePage\)'
        $textSource | Should Match 'IncompleteDbcsTailSize\(bytes\.data\(\),\s*bytes\.size\(\),\s*displayCodePage\)'

        $hexSource | Should Match 'DecodeUtf8Scalar'
        $hexSource | Should Match '0xD800u \+ \(value >> 10u\)'
        $hexSource | Should Match '0xDC00u \+ \(value & 0x3FFu\)'
        $hexSource | Should Match 'ComputeHexClipboardPlan\(_fileSize,\s*startLine,\s*endLine\)'
        $planIndex = $hexSource.IndexOf('ComputeHexClipboardPlan(_fileSize, startLine, endLine)')
        $csvIndex = $hexSource.IndexOf('std::wstring csv;', $planIndex)
        $planIndex | Should BeGreaterThan -1
        $csvIndex | Should BeGreaterThan $planIndex
        $hexSource | Should Match 'viewer\.hex\.clipboard_accepted_bytes'
        $hexSource | Should Match 'viewer\.hex\.clipboard_rejected_bytes'
        $hexSource | Should Match 'viewer\.hex\.clipboard_format_us'
        $hexSource | Should Not Match 'catch\s*\(\s*const\s+std::bad_alloc'

        ($openSource + $textSource + $hexSource + $header) | Should Not Match 'StreamOutCallback|_msftEditModule|DetectEncodingAndSize'
        $tests | Should Match 'TestViewerTextDecodeAndClipboardSafetyHelpers'
        $tests | Should Match 'TestViewerTextAsyncOpenAndUtf8HexTerminalContracts'
        $tests | Should Match 'ViewerTextDebugAsyncOpenFault::PayloadPost'
        $tests | Should Match 'ViewerTextDebugAsyncOpenFault::Submit'
    }

    It 'requires git identity before accepting self-test archive repo roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'kRepoGitDirName'
        $source | Should Match 'IsRepoRootCandidate'
        $source | Should Match 'candidate\s*/\s*kRepoGitDirName'
    }

    It 'keeps self-test archive parent walking bounded by a named limit' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'

        $source | Should Match 'kRepoRootParentWalkLimit'
        $source | Should Match 'i\s*<\s*kRepoRootParentWalkLimit'
    }

    It 'scopes the self-test run mutex to the isolated self-test root' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $source | Should Match 'kSelfTestRunMutexPrefix'
        $source | Should Match 'BuildSelfTestRunMutexName'
        $source | Should Match 'SelfTest::SelfTestRoot\(\)\.native\(\)'
        $source | Should Match 'HashSelfTestRootForMutex'
        $source | Should Not Match 'constexpr\s+wchar_t\s+kSelfTestRunMutexName\[\]'
    }

    It 'lets native selftests consume the unified test root and run id directly' {
        $common = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $runner = Get-RSText -Path 'Tools\Run-AllTests.ps1'

        $common | Should Match 'REDSALAMANDER_TEST_ROOT'
        $common | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $common | Should Match 'GetUnifiedSelfTestRootFromEnvironment'
        $common | Should Match 'kRunsDirName'
        $common | Should Match 'kArtifactsDirName'
        $common | Should Match 'kSelfTestArtifactDirName'
        $common.IndexOf('GetUnifiedSelfTestRootFromEnvironment') | Should BeLessThan $common.IndexOf('GetSelfTestRootOverrideFromEnvironment')

        $runner | Should Match '\$env:REDSALAMANDER_TEST_ROOT\s*=\s*\$testRunContext\.TestRoot'
        $runner | Should Match '\$env:REDSALAMANDER_TEST_RUN_ID\s*=\s*\$testRunContext\.RunId'
        $runner | Should Match '-SelfTestRootOverride\s+''''\s*'
        $runner | Should Match '\[Environment\]::SetEnvironmentVariable\(''REDSALAMANDER_SELFTEST_ROOT'',\s*\$null,\s*''Process''\)'
        $runner | Should Not Match '\$env:REDSALAMANDER_SELFTEST_ROOT\s*=\s*\$testRunContext\.LegacySelfTestRoot'
    }

    It 'exposes native TestSandbox scratch acquisition separate from last_run artifacts' {
        $common = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $header = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'

        $header | Should Match 'struct\s+TestSandbox'
        $header | Should Match 'AcquireTestSandbox'
        $common | Should Match 'kScratchDirName\{L"scratch"\}'
        $common | Should Match 'GetUnifiedTestRunScratchRootFromEnvironment'
        $common | Should Match 'SelfTestRoot\(\)\s*/\s*kLastRunDirName'
        $common | Should Match 'AcquireTestSandbox'
        $common | Should Match 'std::filesystem::create_directories\(sandboxRoot'
        $common | Should Match 'CreateTestSandboxAtScratchRoot'
        $common | Should Match 'AppendSelfTestTrace\(std::format'
    }

    It 'routes FileOperations cross-volume scratch through the sanctioned alternate TestSandbox root' {
        $common = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $header = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $fileOps = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $header | Should Match 'AcquireTestSandboxOnVolume'
        $common | Should Match 'kAlternateVolumeTestSandboxDirName\{L"RedSalamanderTestSandbox"\}'
        $common | Should Match 'AcquireTestSandboxOnVolume'
        $common | Should Match 'normalizedVolumeRoot\s*/\s*std::wstring\(kAlternateVolumeTestSandboxDirName\)\s*/\s*std::wstring\(kRunsDirName\)'
        $common | Should Match 'std::wstring\(kScratchDirName\)'

        $fileOps | Should Match 'AcquireTestSandboxOnVolume\(SelfTest::SelfTestSuite::FileOperations,\s*L"real_cross_volume_move"'
        $fileOps | Should Match 'alternateVolumeSandbox'
        $fileOps | Should Match 'PruneEmptyAlternateVolumeSandboxParents'
        $fileOps | Should Not Match 'RedSalamanderCrossVolumeSelfTest_'
    }

    It 'routes RedConfigureTests scratch through the unified TestSandbox root' {
        $source = Get-RSText -Path 'Tests\RedConfigureTests\RedConfigureTests.cpp'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'

        $source | Should Match 'AcquireRedConfigureTestSandbox'
        $source | Should Match 'kRedConfigureHarnessSegment\{L"redconfigure"\}'
        $source | Should Match 'TestSupport::AcquireTestDirectory'
        $source | Should Match '\.cleanExisting\s*=\s*false'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kScratchDirectoryName'
        $source | Should Not Match 'std::filesystem::temp_directory_path'
    }

    It 'routes ViewerSqliteTests scratch through the unified TestSandbox root' {
        $source = Get-RSText -Path 'Tests\ViewerSqliteTests\ViewerSqliteTests.cpp'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'

        $source | Should Match 'AcquireViewerSqliteTestSandbox'
        $source | Should Match 'kViewerSqliteHarnessSegment\{L"viewer-sqlite"\}'
        $source | Should Match 'TestSupport::AcquireTestDirectory'
        $source | Should Match '\.cleanExisting\s*=\s*false'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kScratchDirectoryName'
        $source | Should Not Match 'std::filesystem::temp_directory_path'
    }

    It 'routes CrashHandlingTests scratch through the unified TestSandbox root' {
        $source = Get-RSText -Path 'Tests\CrashHandlingTests\CrashHandlingTests.cpp'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'

        $source | Should Match 'AcquireCrashHandlingTestSandbox'
        $source | Should Match 'kCrashHandlingHarnessSegment\{L"crash-handling"\}'
        $source | Should Match 'TestSupport::AcquireTestDirectory'
        $source | Should Match '\.cleanExisting\s*=\s*false'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kScratchDirectoryName'
        $source | Should Not Match 'std::filesystem::temp_directory_path'
    }

    It 'routes Commands plugin-config scratch through native TestSandbox roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp'

        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"viewer_text_hex_byte_color_perf"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"viewer_text_diff_perf"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"viewer_imgraw_close_roundtrip"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"viewer_web_close_roundtrip"\)'
        $source | Should Not Match 'std::filesystem::temp_directory_path'
    }

    It 'recovers plugin-config debug tab traversal from logical focus position before first-control fallback' {
        $source = Get-RSText -Path 'RedSalamander\ManagePluginsDialog.cpp'

        $source | Should Match 'lastDebugFocusedHostIndex'
        $source | Should Match 'TryResolvePluginConfigDebugFocusHostByIndex'
        $source | Should Match 'TryRecoverPluginConfigDebugFocusedHost'
        $source | Should Match 'MovePluginConfigDialogTabFocusFromHost[\s\S]*RememberPluginConfigDebugFocusedHost\(\*state,\s*hosts\[nextIndex\]\.host,\s*nextIndex\)'
        $source | Should Match 'FillPluginConfigFocusSnapshot[\s\S]*TryFillPluginConfigFocusSnapshotFromDebugHost'
        $source | Should Match 'case PluginConfigDebugCommand::AdvanceTab:[\s\S]*TryRecoverPluginConfigDebugFocusedHost'
    }

    It 'keeps plugin-config tab traversal retries from sending an extra tab after a successful advance' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.PluginConfig.cpp'

        $source | Should Match 'waitForExpectedSnapshot'
        $source | Should Match 'const bool advanced = DebugAdvancePluginConfigurationDialogTab\(reverse\)'
        $source | Should Not Match 'if \(! reached\)\s*\{\s*PumpPendingMessages\(\);\s*reached = advanceAndWait\(\);'
    }

    It 'waits for preview fallback text and source-pane focus together' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'

        $source | Should Match 'WaitForPreviewPaneTextAndFocus'
        $source | Should Match 'WaitForPreviewPaneTextAndFocus\(L"Name: mystery\.no-preview-props"'
        $source | Should Match 'WaitForPreviewPaneTextAndFocus\(L"Name: folder-no-preview-props"'
        $source | Should Match 'g_folderWindow\.GetFocusedFolderViewHwnd\(\) == expectedFocus'
    }

    It 'makes quick-search tests own pane sort and extension visibility' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
        $viewSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $source | Should Match 'const FolderView::SortBy leftSortByBefore\s*=\s*g_folderWindow\.GetSortBy\(FolderWindow::Pane::Left\)'
        $source | Should Match 'const FolderView::SortDirection leftSortDirectionBefore\s*=\s*g_folderWindow\.GetSortDirection\(FolderWindow::Pane::Left\)'
        $source | Should Match 'g_folderWindow\.SetSort\(FolderWindow::Pane::Left,\s*FolderView::SortBy::Name,\s*FolderView::SortDirection::Ascending\)'
        $source | Should Match 'g_folderWindow\.SetFileExtensionsVisible\(FolderWindow::Pane::Left,\s*true\)'
        $source | Should Match 'g_folderWindow\.SetFileExtensionsVisible\(FolderWindow::Pane::Left,\s*false\)'
        $viewSource | Should Match '\[\[nodiscard\]\]\s+bool\s+ClearSyntheticTextInputModifiersForFolderViewPerf\(\)\s+noexcept'
        $viewSource | Should Match 'SetKeyboardState\(keyboardState\.data\(\)\)\s*==\s*FALSE'
        $viewSource | Should Match 'GetKeyState\(VK_CONTROL\)[\s\S]*GetKeyState\(VK_MENU\)'
        $viewSource | Should Match 'state\.Require\(SendFolderViewQuickSearchCharForPerf\(folderView,\s*ch\)'

        $hugeStart = $viewSource.IndexOf('[[nodiscard]] bool TestFolderViewPerfHugeFolderScale')
        $hugeNext = $viewSource.IndexOf('[[nodiscard]] bool TestFolderViewPerfColdFirstVisit', $hugeStart)
        $hugeStart | Should BeGreaterThan -1
        $hugeNext | Should BeGreaterThan $hugeStart
        $hugeBlock = $viewSource.Substring($hugeStart, $hugeNext - $hugeStart)
        $hugeBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"huge FolderView scale Quick Search perf"\)'
        $hugeBlock | Should Match 'focusedFolderView=0x\{:X\}, win32Focus=0x\{:X\}, foreground=0x\{:X\}'
    }

    It 'keeps quick-search navigation expectations tied to visible match order' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPaneQuickSearchIntegratedNavigation')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestPaneQuickSearchUsesVisualNamesWhenExtensionsHidden', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'quickSearchMatchDisplayOrder'
        $testBlock | Should Match 'findNextQuickSearchMatchName'
        $testBlock | Should Match 'expectedAfterFirstDown'
        $testBlock | Should Match 'expectedAfterSecondDown'
        $testBlock | Should Match 'focusRestored\s*=\s*activated\s*&&\s*waitForLeftFolderViewFocusPassive\(SelfTest::Scale\(500ms\)\)'
        $testBlock | Should Match 'stillEmpty\s*=\s*focusRestored\s*&&\s*waitForQuickSearchSnapshot'
        $testBlock | Should Match 'activateQuickSearchThroughCommand\(L"Quick Search no-match reactivation",\s*snapshot\)'
        $testBlock | Should Match 'activateQuickSearchThroughCommand\(L"Quick Search shortcut-routed Space reactivation",\s*snapshot\)'
        $activationHelperStart = $testBlock.IndexOf('const auto activateQuickSearchThroughCommand')
        $matchOrderStart = $testBlock.IndexOf('const auto quickSearchMatchDisplayOrder', $activationHelperStart)
        $activationHelperStart | Should BeGreaterThan -1
        $matchOrderStart | Should BeGreaterThan $activationHelperStart
        $activationHelperBlock = $testBlock.Substring($activationHelperStart, $matchOrderStart - $activationHelperStart)
        $activationHelperBlock | Should Match 'SendMessageW\(mainWindow,\s*WM_COMMAND,\s*MAKEWPARAM\(IDM_PANE_QUICK_SEARCH,\s*0\),\s*0\)'
        $activationHelperBlock | Should Not Match 'g_folderWindow\.CommandQuickSearch'

        $noMatchStart = $testBlock.IndexOf('activateQuickSearchThroughCommand(L"Quick Search no-match reactivation", snapshot)')
        $noMatchEnd = $testBlock.IndexOf('SendMessageW(folderView, WM_KEYDOWN, VK_ESCAPE, 0)', $noMatchStart)
        $noMatchStart | Should BeGreaterThan -1
        $noMatchEnd | Should BeGreaterThan $noMatchStart
        $noMatchBlock = $testBlock.Substring($noMatchStart, $noMatchEnd - $noMatchStart)
        $noMatchBlock | Should Not Match 'g_folderWindow\.CommandQuickSearch'
        $testBlock | Should Match 'Quick Search should select one of the starts-with matches'
        $testBlock | Should Not Match 'snapshot\.focusedDisplayName == L"alpha\.txt"[\s\S]*Quick Search should select the first starts-with match'
        $testBlock | Should Not Match 'snapshot\.focusedDisplayName == L"alpine\.log"[\s\S]*Quick Search Down should move to the next match'
        $testBlock | Should Not Match 'snapshot\.focusedDisplayName == L"beta-alpha\.txt"[\s\S]*Quick Search Down should include contained matches'
    }

    It 'keeps Hot Paths browse page-settle failures diagnostic-rich' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogHotPathsBrowseLiveDxInteraction')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogHotPathsTabTraversalLiveDxInteraction', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'const bool pageReady = waitForSnapshot'
        $testBlock | Should Match 'Preferences Hot Paths page did not settle to the active DX surface before browse interaction validation; \{\}'
        $testBlock | Should Match 'DescribeHotPathsSnapshot\(snapshot\)'
    }

    It 'uses condition-based focus waiting for Viewers roundtrip category navigation' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogViewersRoundTripRestoresDxUiSurface')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogViewersThemeCycleKeepsSurfaceLegible', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'const auto focusCategoryTreeHost'
        $testBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $testBlock | Should Match 'nativeFocus=0x\{:X\}, categoryHost=0x\{:X\}, activePage=0x\{:X\}'
        $testBlock | Should Match 'focusCategoryTreeHost\(std::format\(L"before \{\}",\s*context\)\)'
        $testBlock | Should Not Match 'SetFocus\(categoryTreeHost\)\s*==\s*categoryTreeHost'
    }

    It 'retries Preferences focus acquisition during condition-based waits' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.cpp'

        $helperStart = $source.IndexOf('[[nodiscard]] bool FocusWindowAndWait')
        $nextBlock = $source.IndexOf('struct ScopedSettingsArtifactBackup', $helperStart)
        $helperStart | Should BeGreaterThan -1
        $nextBlock | Should BeGreaterThan $helperStart

        $helperBlock = $source.Substring($helperStart, $nextBlock - $helperStart)
        $helperBlock | Should Match 'const auto targetHasFocus'
        $helperBlock | Should Match 'IsChild\(hwnd,\s*focused\)'
        $helperBlock | Should Match 'AttachThreadInput\(foregroundThreadId,\s*currentThreadId,\s*TRUE\)'
        $helperBlock | Should Match 'AttachThreadInput\(foregroundThreadId,\s*currentThreadId,\s*FALSE\)'
        $helperBlock | Should Match 'ShowWindow\(root,\s*SW_SHOWNORMAL\)'
        $helperBlock | Should Match 'BringWindowToTop\(root\)'
        $helperBlock | Should Match 'SetForegroundWindow\(root\)'
        $helperBlock | Should Match 'SetWindowPos\(root,\s*HWND_TOPMOST'
        $helperBlock | Should Match 'SetWindowPos\(root,\s*HWND_NOTOPMOST'
        $helperBlock | Should Match 'const auto tryFocus'
        $helperBlock | Should Match 'SetFocus\(hwnd\)'
        $helperBlock | Should Match 'while \([^\r\n]*deadline\)[\s\S]*static_cast<void>\(tryFocus\(\)\);'
    }

    It 'waits for Plugins sort-search header rects before reorder-resize interactions' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp'

        $plainTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSearchRoundTrip')
        $plainNextTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSearchRoundTrip', $plainTestStart)
        $plainTestStart | Should BeGreaterThan -1
        $plainNextTestStart | Should BeGreaterThan $plainTestStart

        $plainTestBlock = $source.Substring($plainTestStart, $plainNextTestStart - $plainTestStart)
        $plainTestBlock | Should Match 'const auto waitForInitialMainHeaderRects'
        $plainTestBlock | Should Match 'DebugGetPreferencesPluginsMainListHeaderClientRect\(0u,\s*currentNameHeaderRect\)'
        $plainTestBlock | Should Match 'DebugGetPreferencesPluginsMainListHeaderClientRect\(1u,\s*currentTypeHeaderRect\)'
        $plainTestBlock | Should Match 'value\.pluginsMainListVisibleColumnCount >= 2u'
        $plainTestBlock | Should Match 'currentNameHeaderRect\.left < currentTypeHeaderRect\.left'
        $plainTestBlock | Should Match 'Name/Type header rects before combined reordered-resized/search validation'
        $plainTestBlock | Should Not Match 'Failed to capture the visible Preferences Plugins Name header rect before combined reordered-resized/search validation'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogPluginsReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $testBlock | Should Match 'const auto waitForInitialMainHeaderRects'
        $testBlock | Should Match 'DebugGetPreferencesPluginsMainListHeaderClientRect\(0u,\s*currentNameHeaderRect\)'
        $testBlock | Should Match 'DebugGetPreferencesPluginsMainListHeaderClientRect\(1u,\s*currentTypeHeaderRect\)'
        $testBlock | Should Match 'value\.pluginsMainListVisibleColumnCount >= 2u'
        $testBlock | Should Match 'currentNameHeaderRect\.left < currentTypeHeaderRect\.left'
        $testBlock | Should Match 'Name/Type header rects before reordered-resized-sort/search validation'
        $testBlock | Should Not Match 'state\.Require\(DebugGetPreferencesPluginsMainListHeaderClientRect\(0u,\s*nameHeaderRect\)'
    }

    It 'waits for Viewers sort-search header rects before reorder-resize interactions' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedColumnsSurviveSortCyclesAndSearchRoundTrip')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogViewersReorderedResizedCopyFollowsVisibleColumnsAfterSortCyclesAndSearchRoundTrip', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'const auto waitForInitialMainHeaderRects'
        $testBlock | Should Match 'DebugGetPreferencesViewersListHeaderClientRect\(kViewersAssociationMatchColumn,\s*currentExtensionHeaderRect\)'
        $testBlock | Should Match 'DebugGetPreferencesViewersListHeaderClientRect\(kViewersAssociationPrimaryActionColumn,\s*currentViewerHeaderRect\)'
        $testBlock | Should Match 'value\.viewersListVisibleColumnCount >= 2u'
        $testBlock | Should Match 'currentExtensionHeaderRect\.left < currentViewerHeaderRect\.left'
        $testBlock | Should Match 'Match/F3 View header rects before'
        $testBlock | Should Not Match 'state\.Require\(DebugGetPreferencesViewersListHeaderClientRect\(kViewersAssociationMatchColumn,\s*extensionHeaderRect\)'
    }

    It 'keeps broad Preferences focus paths on the shared focus helper' {
        $themes = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp'
        $compare = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp'

        @(
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp',
            'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp'
        ) | ForEach-Object {
            $focusSource = Get-RSText -Path $_
            $focusSource | Should Match 'FocusWindowAndWait\('
            $focusSource | Should Not Match 'SetFocus\((?<target>\w+)\)\s*==\s*\k<target>'
        }

        $themesStart = $themes.IndexOf('[[nodiscard]] bool TestPreferencesDialogThemesSearchRoundTripPreservesRetainedState')
        $themesNext = $themes.IndexOf('[[nodiscard]] bool TestPreferencesDialogThemesSelectionSurvivesLegacyComboClear', $themesStart)
        $themesBlock = $themes.Substring($themesStart, $themesNext - $themesStart)
        $themesBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $themesBlock | Should Not Match 'SetFocus\(categoryTreeHost\) == categoryTreeHost'

        $compareStart = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesPageUsesDxUiStatics')
        $compareNext = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesLiveDxInteraction', $compareStart)
        $compareBlock = $compare.Substring($compareStart, $compareNext - $compareStart)
        $compareBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $compareBlock | Should Not Match 'SetFocus\(categoryTreeHost\) == categoryTreeHost'

        $compareRoundTripStart = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesRoundTripRestoresDxUiSurface')
        $compareRoundTripNext = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCompareDirectoriesThemeCycleKeepsSurfaceLegible', $compareRoundTripStart)
        $compareRoundTripBlock = $compare.Substring($compareRoundTripStart, $compareRoundTripNext - $compareRoundTripStart)
        $compareRoundTripBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $compareRoundTripBlock | Should Match 'isStablePage'
        $compareRoundTripBlock | Should Match 'pageHostUsesDxUiHost'
        $compareRoundTripBlock | Should Match 'pageTitle == expectedTitle'
        $compareRoundTripBlock | Should Match 'pageDescription == expectedDescription'
        $compareRoundTripBlock | Should Match 'describeSnapshot\(snapshot\)'
        $compareRoundTripBlock | Should Match 'DebugFocusPreferencesCategoryTree\(\)'
        $compareRoundTripBlock | Should Not Match 'Phase 8: removed field'
        $compareRoundTripBlock | Should Not Match 'SetFocus\(categoryTreeHost\) == categoryTreeHost'

        $churnStart = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategorySwitchesDoNotChurnTreeHost')
        $churnNext = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogScrollHostPreservesRetainedPageState', $churnStart)
        $churnStart | Should BeGreaterThan -1
        $churnNext | Should BeGreaterThan $churnStart
        $churnBlock = $compare.Substring($churnStart, $churnNext - $churnStart)
        $churnBlock | Should Match 'const auto focusCategoryTreeHost'
        $churnBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(std::chrono::milliseconds\{1000\}\)\)'
        $churnBlock | Should Match 'DebugFocusPreferencesCategoryTree\(\)'
        $churnBlock | Should Match 'categoryTreeDxHostFocusControlActive'

        $tabStart = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogEditorsAndMouseTabSkipNoteSurface')
        $tabNext = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogViewersEditorsFileActionSettingsApply', $tabStart)
        $tabBlock = $compare.Substring($tabStart, $tabNext - $tabStart)
        $tabBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $tabBlock | Should Match 'FocusWindowAndWait\(target,\s*SelfTest::Scale\(1000ms\)\)'
        $tabBlock | Should Match 'SelfTest::Scale\(5000ms\)'
        $tabBlock | Should Match 'expectedCategoryTreeFocused=\{\}'
        $tabBlock | Should Match 'expectedCategory=\{\}'
        $tabBlock | Should Match 'createdPaneWindowCount=\{\}'
        $tabBlock | Should Match 'visiblePaneWindowCount=\{\}'
        $tabBlock | Should Not Match 'SetFocus\(categoryTreeHost\) == categoryTreeHost'

        $reverseStart = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategoryTreeReverseKeyboardNavigation')
        $reverseNext = $compare.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategoryTreePageNavigation', $reverseStart)
        $reverseBlock = $compare.Substring($reverseStart, $reverseNext - $reverseStart)
        $reverseBlock | Should Match 'FocusWindowAndWait\(categoryTreeHost,\s*SelfTest::Scale\(1000ms\)\)'
        $reverseBlock | Should Match 'DebugSelectPreferencesCategory\(kPrefCategoryPlugins\)'
        $reverseBlock | Should Match 'waitForPluginsCollapsed'
        $reverseBlock | Should Match 'snapshot\.pluginsExpanded'
        $reverseBlock | Should Match 'DebugSendPreferencesCategoryTreeKey\(VK_LEFT\)'
        $reverseBlock | Should Match 'DebugSelectPreferencesCategory\(kPrefCategoryGeneral\)'
        $reverseBlock | Should Match 'DebugFocusPreferencesCategoryTree\(\)'
        $reverseBlock | Should Match 'categoryTreeDxHostFocusControlActive'
        $reverseBlock | Should Match 'categoryTreeHasSelectedItem'
        $reverseBlock | Should Match 'categoryTreeDxHostRenderCount != 0u'
        $reverseBlock | Should Match 'restoreCategoryTreePrecondition'
        $reverseBlock | Should Match 'allowCategoryReselect'
        $reverseBlock | Should Match 'WaitForPreferencesCategoryTreeRenderCountToSettle\(settledSnapshot\)'
        $reverseBlock | Should Match 'waitForCategoryTreeReadyForKey'
        $reverseBlock | Should Match 'kRequiredStableSamples = 3u'
        $reverseBlock | Should Match 'SelfTest::Scale\(3000ms\)'
        $reverseBlock | Should Match 'stableSamples = 0u'
        $reverseBlock | Should Match 'initial reverse-navigation reset", true\)'
        $reverseBlock | Should Match 'pre-key", context\), false\)'
        $reverseBlock | Should Match 'sendCategoryTreeKey\(kPrefCategoryGeneral'
        $reverseBlock | Should Match 'sendCategoryTreeKey\(kPrefCategoryAdvanced'
        $reverseBlock | Should Match 'sendCategoryTreeKey\(kPrefCategoryMonitor'
        $reverseBlock | Should Match 'sendCategoryTreeKey\(kPrefCategoryHotPaths'
        $reverseBlock | Should Match 'sendCategoryTreeKey\([^\r\n]+VK_HOME'
        $reverseBlock | Should Match 'sendCategoryTreeKey\([^\r\n]+VK_END'
        $reverseBlock | Should Match 'sendCategoryTreeKey\([^\r\n]+VK_UP'
        $reverseBlock | Should Match 'sendCategoryTreeKey\([^\r\n]+VK_DOWN'
        $reverseBlock | Should Match 'DebugSendPreferencesCategoryTreeKey\(virtualKey\)'
        $reverseBlock | Should Match 'pre-key'
        $reverseBlock | Should Not Match 'SendMessageW\(categoryTreeHost,\s*WM_KEYDOWN'
        $reverseBlock | Should Not Match 'SetFocus\(categoryTreeHost\) == categoryTreeHost'

        $prefsHeader = Get-RSText -Path 'RedSalamander\Preferences.h'
        $prefsHeader | Should Match 'DebugFocusPreferencesCategoryTree\(\)'
        $prefsHeader | Should Match 'DebugSendPreferencesCategoryTreeKey\(UINT virtualKey\)'
        $prefsHeader | Should Match 'categoryTreeDxHostFocusControlActive'

        $prefsDialog = Get-RSText -Path 'RedSalamander\Preferences.Dialog.cpp'
        $mainWindowSource = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $navigationSource = Get-RSText -Path 'RedSalamander\NavigationView.cpp'
        $folderInteractionSource = Get-RSText -Path 'RedSalamander\FolderWindow.Interaction.cpp'
        $categoryHostStart = $prefsDialog.IndexOf('LRESULT CALLBACK PreferencesDxCategoryHostWndProc')
        $categoryHostNext = $prefsDialog.IndexOf('LRESULT CALLBACK PreferencesDxShellHostWndProc', $categoryHostStart)
        $categoryHostStart | Should BeGreaterThan -1
        $categoryHostNext | Should BeGreaterThan $categoryHostStart

        $categoryHostBlock = $prefsDialog.Substring($categoryHostStart, $categoryHostNext - $categoryHostStart)
        $categoryHostBlock | Should Match 'msg == WM_LBUTTONDOWN[\s\S]*SetFocus\(hwnd\)[\s\S]*_categoryTreeHost\.HandleMessage'

        $restoreStart = $mainWindowSource.IndexOf('case WndMsg::kPaneRestoreFolderFocus:')
        $restoreNext = $mainWindowSource.IndexOf('case WM_TIMER:', $restoreStart)
        $restoreStart | Should BeGreaterThan -1
        $restoreNext | Should BeGreaterThan $restoreStart
        $restoreBlock = $mainWindowSource.Substring($restoreStart, $restoreNext - $restoreStart)
        $restoreBlock | Should Match 'IsWindowEnabled\(hWnd\) == FALSE'
        $restoreBlock | Should Match 'GetActiveWindow\(\)[\s\S]*activeWindow != hWnd'
        $restoreBlock | Should Match 'GetForegroundWindow\(\)[\s\S]*foregroundWindow != hWnd'

        $navigationRestoreStart = $navigationSource.IndexOf('case WndMsg::kNavigationViewRestoreFolderFocus:')
        $navigationRestoreNext = $navigationSource.IndexOf('return DefWindowProcW', $navigationRestoreStart)
        $navigationRestoreStart | Should BeGreaterThan -1
        $navigationRestoreNext | Should BeGreaterThan $navigationRestoreStart
        $navigationRestoreBlock = $navigationSource.Substring($navigationRestoreStart, $navigationRestoreNext - $navigationRestoreStart)
        $navigationRestoreBlock | Should Match 'GetActiveWindow\(\)[\s\S]*activeWindow != root'
        $navigationRestoreBlock | Should Match 'GetForegroundWindow\(\)[\s\S]*foregroundWindow != root'
        $navigationRestoreBlock | Should Not Match 'SetActiveWindow\(root\)'

        $folderRestoreStart = $folderInteractionSource.IndexOf('bool FolderWindow::TryRestoreActivePaneFolderViewFocus')
        $folderRestoreNext = $folderInteractionSource.IndexOf('FolderWindow::Pane FolderWindow::GetPaneFromChild', $folderRestoreStart)
        $folderRestoreStart | Should BeGreaterThan -1
        $folderRestoreNext | Should BeGreaterThan $folderRestoreStart
        $folderRestoreBlock = $folderInteractionSource.Substring($folderRestoreStart, $folderRestoreNext - $folderRestoreStart)
        $folderRestoreBlock | Should Match 'GetActiveWindow\(\)[\s\S]*activeWindow != root'
        $folderRestoreBlock | Should Match 'GetForegroundWindow\(\)[\s\S]*foregroundWindow != root'

        $focusStart = $prefsDialog.IndexOf('bool PreferencesDialog::DebugFocusCategoryTree')
        $focusNext = $prefsDialog.IndexOf('bool PreferencesDialog::DebugSendCategoryTreeKey', $focusStart)
        $focusStart | Should BeGreaterThan -1
        $focusNext | Should BeGreaterThan $focusStart

        $focusBlock = $prefsDialog.Substring($focusStart, $focusNext - $focusStart)
        $focusBlock | Should Match '_categoryTreeHost\.ResetInteractionState\(\)'
        $focusBlock | Should Match '_categoryTreeHost\.SetFocusControl\(hostState\._categoryTreeControl\)'

        $keyStart = $prefsDialog.IndexOf('bool PreferencesDialog::DebugSendCategoryTreeKey')
        $keyNext = $prefsDialog.IndexOf('namespace', $keyStart)
        $keyStart | Should BeGreaterThan -1
        $keyNext | Should BeGreaterThan $keyStart

        $keyBlock = $prefsDialog.Substring($keyStart, $keyNext - $keyStart)
        $keyBlock | Should Match 'DebugFocusCategoryTree\(\)'
        $keyBlock | Should Match 'MapVirtualKeyW\(virtualKey,\s*MAPVK_VK_TO_VSC_EX\)'
        $keyBlock | Should Match 'scanCode\s*&\s*0xFF00u[\s\S]*1ull\s*<<\s*24u'
        $keyBlock | Should Match 'SendMessageW\(state->categoryTreeWindow,\s*WM_KEYDOWN,\s*virtualKey,\s*keyDownLParam\)'
        $keyBlock | Should Match 'SendMessageW\(state->categoryTreeWindow,\s*WM_KEYUP,\s*virtualKey,\s*keyUpLParam\)'
        $keyBlock | Should Not Match '_categoryTreeControl->OnKeyDown'
    }

    It 'keeps broad Commands app navigation shell focus checks condition-based' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $driveStart = $source.IndexOf('[[nodiscard]] bool TestAppOpenDriveMenusKeepsNavigationShellStable')
        $driveNext = $source.IndexOf('[[nodiscard]] bool TestPaneRefresh', $driveStart)
        $driveStart | Should BeGreaterThan -1
        $driveNext | Should BeGreaterThan $driveStart

        $driveBlock = $source.Substring($driveStart, $driveNext - $driveStart)
        $driveBlock | Should Match 'waitForTargetPaneFocus'
        $driveBlock | Should Match 'std::chrono::steady_clock::now\(\) \+ SelfTest::Scale\(3000ms\)'
        $driveBlock | Should Match 'DescribeWindowHandleForSelfTest\(GetFocus\(\)\)'
        $driveBlock | Should Not Match 'state\.Require\(GetFocus\(\) == folderView \|\| GetFocus\(\) == navigationView'

        $swapStart = $source.IndexOf('[[nodiscard]] bool TestSwapPanesKeepsNavigationShellStable')
        $swapNext = $source.IndexOf('[[nodiscard]] bool TestToggleUiChromeKeepsNavigationShellStable', $swapStart)
        $swapStart | Should BeGreaterThan -1
        $swapNext | Should BeGreaterThan $swapStart

        $swapBlock = $source.Substring($swapStart, $swapNext - $swapStart)
        $swapBlock | Should Match 'expectedFocusedFolderView\s*=\s*g_folderWindow\.GetFolderViewHwnd\(FolderWindow::Pane::Left\)'
        $swapBlock | Should Match 'focusedFolderView=0x\{:X\}'
        $swapBlock | Should Match 'DescribeWindowHandleForSelfTest\(GetFocus\(\)\)'
        $swapBlock | Should Match 'paneRestoreAttempted\s*=\s*true'
        $swapBlock | Should Match 'WaitForAtomicAtLeast\(leftRestoreEnumCount'
        $swapBlock | Should Match 'WaitForAtomicAtLeast\(rightRestoreEnumCount'
        $swapBlock | Should Match 'FocusFolderViewPane\(activePaneBefore\)'
        $swapBlock | Should Not Match 'g_folderWindow\.GetFocusedFolderViewHwnd\(\) == leftFolderView'

        $historyStart = $source.IndexOf('[[nodiscard]] bool TestPaneNavigationViewHistoryDropdownKeyboardNavigation')
        $historyNext = $source.IndexOf('[[nodiscard]] bool TestPaneNavigationViewHistoryDropdownEscapeReturnsFocusToFolderView', $historyStart)
        $historyStart | Should BeGreaterThan -1
        $historyNext | Should BeGreaterThan $historyStart

        $historyBlock = $source.Substring($historyStart, $historyNext - $historyStart)
        $historyBlock | Should Match 'popupStateDeadline\s*=\s*std::chrono::steady_clock::now\(\)\s*\+\s*SelfTest::Scale\(3000ms\)'
        $historyBlock | Should Match 'DebugGetContextMenuPopupState\(popup,\s*popupState\)\s*&&\s*popupState\.keyboardIndex\.has_value\(\)'
        $historyBlock | Should Match 'popupResult\.popupStateCaptured\s*=\s*true'
        $historyBlock | Should Not Match 'if\s*\(!\s*RedSalamander::DxUi::DebugGetContextMenuPopupState\(popup,\s*popupState\)'
    }

    It 'makes the app Shortcuts navigation-shell fixture own its settings and window isolation' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $testStart = $source.IndexOf('[[nodiscard]] bool TestAppShowShortcutsKeepsNavigationShellStable')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestSwapPanesCommand', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"App Shortcuts navigation-shell validation"\)'
        $testBlock | Should Match 'shortcutsBefore\s*=\s*g_settings\.shortcuts'
        $testBlock | Should Match 'g_settings\.shortcuts\s*=\s*ShortcutDefaults::CreateDefaultShortcuts\(\)'
        $testBlock | Should Match 'g_settings\.shortcuts\s*=\s*shortcutsBefore;[\s\S]*DebugReloadShortcutsFromSettings\(\)'
    }

    It 'isolates the Preferences Keyboard reset live interaction fixture' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp'
        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogKeyboardResetLiveDxInteraction')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogKeyboardExportLiveDxInteraction', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"Preferences Keyboard reset live interaction validation"\)'
        $testBlock | Should Match 'InvokeVisibleDxAction\(getShellHost\(\),\s*UIA_ButtonControlTypeId,\s*cancelButtonText\)'
        $testBlock | Should Match 'WaitForWindowClosed\(prefs,\s*SelfTest::Scale\(3000ms\)\)'
    }

    It 'isolates the pane-filter DxUi surface fixture before global prompt discovery' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp'
        $testStart = $source.IndexOf('[[nodiscard]] bool TestPaneFilterPromptUsesDxUiSurface')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestPaneFilterPromptLiveDxInteraction', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"pane-filter DxUi surface validation"\)'
        $testBlock | Should Match 'DebugSetFolderViewPaneFilterPromptHelpExpanded\(true\)'
        $testBlock | Should Match 'DebugGetFolderViewPaneFilterPromptSnapshot\(cycle\.expandedHelpSnapshot\)'
    }

    It 'isolates both Compare navigation-shell fixtures before asserting restored focus' {
        $navigation = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Navigation.cpp'
        $viewCommands = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $navigationStart = $navigation.IndexOf('[[nodiscard]] bool TestCompareDirectoriesKeepsNavigationShellStable')
        $navigationNext = $navigation.IndexOf('[[nodiscard]] bool TestSwitchPaneFocusKeepsNavigationShellStable', $navigationStart)
        $viewStart = $viewCommands.IndexOf('[[nodiscard]] bool TestAppCompareKeepsNavigationShellStable')
        $viewNext = $viewCommands.IndexOf('[[nodiscard]] bool TestSwapPanesKeepsNavigationShellStable', $viewStart)
        $navigationStart | Should BeGreaterThan -1
        $navigationNext | Should BeGreaterThan $navigationStart
        $viewStart | Should BeGreaterThan -1
        $viewNext | Should BeGreaterThan $viewStart

        $navigationBlock = $navigation.Substring($navigationStart, $navigationNext - $navigationStart)
        $viewBlock = $viewCommands.Substring($viewStart, $viewNext - $viewStart)
        $navigationBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"Compare Directories navigation-shell validation"\)'
        $navigationBlock | Should Match 'WaitForFolderViewPaneFocus\(FolderWindow::Pane::Left,\s*folderView,\s*SelfTest::Scale\(3000ms\)\)'
        $viewBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"App Compare Directories navigation-shell validation"\)'
        $viewBlock | Should Match 'WaitForFolderViewPaneFocus\(FolderWindow::Pane::Left,\s*folderView,\s*SelfTest::Scale\(3000ms\)\)'
        $viewBlock | Should Match 'foreground=0x\{:X\}'
    }

    It 'isolates mouse-opened menu keyboard ordering and preserves split post diagnostics' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $testStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuMouseOpenedPopupProcessesKeyboardBeforeMouseMove')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestMainMenuTopLevelMappingMatchesRawMenu', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"mouse-opened main-menu keyboard ordering validation"\)'
        $testBlock | Should Match 'keyTargetValid'
        $testBlock | Should Match 'keyDownMessageSent'
        $testBlock | Should Match 'keyUpMessageSent'
        $testBlock | Should Match 'keyboardAppliedBeforeMouseMove'
        $testBlock | Should Match 'targetValid=\{\}, keyDownSent=\{\}, keyUpSent=\{\}'
        $testBlock | Should Not Match 'PostMessageW\(keyTarget,\s*WM_KEYDOWN,\s*VK_DOWN,\s*0\)\s*!=\s*FALSE\s*&&\s*PostMessageW'
    }

    It 'reacquires the submenu placement cascade target after root-menu initialization' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $testStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuSubmenuPlacementMatchesSpec')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestAppPreferencesKeepsNavigationShellStable', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"main-menu submenu placement validation"\)'
        $testBlock | Should Match 'RaiseWindowForDirectedSelfTestInput'
        $testBlock | Should Match 'initializedChildIndex\s*=\s*FindFirstCascadeMenuItemIndex\(popupMenu\)'
        $testBlock | Should Match 'childIndexAfterOpen'
        $testBlock | Should Match 'rootPopupValidAfterNavigation'
        $testBlock | Should Match 'navigationKeyDownPosted'
        $testBlock | Should Match 'navigationKeyUpPosted'
    }

    It 'keeps Preferences category-tree UIA selection pinned to category-id selection' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.ChromeAndPlugins.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategoryTreeExposesLiveUiaSelection')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogShellUsesDxUiChrome', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'DebugSelectPreferencesCategory\(kPrefCategoryGeneral\)'
        $testBlock | Should Match 'DebugSelectPreferencesCategory\(kPrefCategoryViewers\)'
        $testBlock | Should Match 'SelfTest::Scale\(3000ms\)'
        $testBlock | Should Match 'CollectUiaSelectionPatternState\(categoryTreeHost\)'
        $testBlock | Should Match 'selectionMatches\(\)'
        $testBlock | Should Match 'expected=\{\} actual=\{\} title='
        $testBlock | Should Not Match 'clickViewers'
        $testBlock | Should Not Match 'WM_LBUTTONDOWN'
    }

    It 'keeps Preferences debug category selection independent from transient snapshot capture' {
        $source = Get-RSText -Path 'RedSalamander\Preferences.Dialog.cpp'

        $selectorStart = $source.IndexOf('bool PreferencesDialog::DebugSelectCategory')
        $selectorNext = $source.IndexOf('bool PreferencesDialog::DebugSelectPluginsTreeChild', $selectorStart)
        $selectorStart | Should BeGreaterThan -1
        $selectorNext | Should BeGreaterThan $selectorStart

        $selectorBlock = $source.Substring($selectorStart, $selectorNext - $selectorStart)
        $selectorBlock | Should Match 'state->categoryTreeUsesDxUi'
        $selectorBlock | Should Match 'EncodeCategoryNodeId\(category\)'
        $selectorBlock | Should Match 'RequestSelectVisibleItem\(visibleIndex\)'
        $selectorBlock | Should Match '! snapshot\.pluginItemSelected'
        $selectorBlock | Should Match 'SelectCategory\(dlg, \*state, category\);[\s\S]*state->currentCategory == category'
        $selectorBlock | Should Not Match 'return DebugGetSnapshot\(snapshot\) && snapshot\.currentCategory == category;'
    }

    It 'keeps Preferences category churn settled, attributable, and bounded under uncontended focus' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp'
        $preferencesHeader = Get-RSText -Path 'RedSalamander\Preferences.h'
        $preferencesDialog = Get-RSText -Path 'RedSalamander\Preferences.Dialog.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategorySwitchesDoNotChurnTreeHost')
        $testNext = $source.IndexOf('[[nodiscard]] bool TestPreferencesDialogScrollHostPreservesRetainedPageState', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'WaitForPreferencesCategoryTreeRenderCountToSettle\(before\)'
        $testBlock | Should Match 'WaitForPreferencesCategoryTreeRenderCountToSettle\(after\)'
        $testBlock | Should Match 'kMaxExpectedCategoryTreeRenderDelta = 12u'
        $testBlock | Should Match 'treeInvalidateDelta\s*=\s*after\.categoryTreeDxHostInvalidateCount\s*-\s*before\.categoryTreeDxHostInvalidateCount'
        $testBlock | Should Match 'treeRenderDelta <= treeInvalidateDelta \+ 1u'
        $testBlock | Should Match 'treeRenderDelta > 0u && treeRenderDelta <= kMaxExpectedCategoryTreeRenderDelta'
        $testBlock | Should Match 'after\.categoryTreeFocused'
        $testBlock | Should Match 'after\.categoryTreeHasSelectedItem'
        $testBlock | Should Match 'after\.visibleCurrentPageChildWindowCount == 1u'
        $testBlock | Should Match 'after\.currentPageRenderedDxHostCount == 1u'
        $testBlock | Should Match 'without runaway repaint'
        $testBlock | Should Match 'renderDelta=\{\}, invalidateDelta=\{\}'
        $testBlock | Should Not Match 'kMaxExpectedCategoryTreeRenderDelta = (4|6)u'
        $preferencesHeader | Should Match 'uint64_t\s+categoryTreeDxHostInvalidateCount\s*=\s*0u'
        $preferencesDialog | Should Match 'out\.categoryTreeDxHostInvalidateCount\s*=\s*hostState\._categoryTreeHost\.DebugGetInvalidateCount\(\)'
        $testBlock | Should Match 'PostMessageW\(mainWindow, WndMsg::kPaneRestoreFolderFocus, 0, 0\)'
        $testBlock | Should Match 'stale main-window focus restore must not steal focus from Preferences'
        $testBlock | Should Match 'after\.foregroundWindow != before\.foregroundWindow'
        $testBlock | Should Match 'requires a stable foreground window'
        $source | Should Match 'page-navigation focus assertions require a stable foreground window'
    }

    It 'carries selected FolderView rename identity across bounded refresh interleavings' {
        $header = Get-RSText -Path 'RedSalamander\FolderView.h'
        $source = Get-RSText -Path 'RedSalamander\FolderView.Enumeration.cpp'

        $header | Should Match 'PendingRefreshSelectionRename[\s\S]*std::chrono::steady_clock::time_point\s+expiresAt'
        $source | Should Match 'rename\.toDisplayName\.assign\(impact->renamedToDisplayName\)'
        $source | Should Match 'const\s+auto\s+now\s*=\s*std::chrono::steady_clock::now\(\);[\s\S]{0,120}const\s+auto\s+expiresAt\s*=\s*now\s*\+\s*std::chrono::seconds\{30\}'
        $source | Should Match 'rename\.expiresAt\s*=\s*expiresAt'
        $source | Should Match '!\s*refreshSelectionRenameTargetsObserved\[renameIndex\][\s\S]*rename\.fromWasSelected[\s\S]*std::chrono::steady_clock::now\(\)\s*<\s*rename\.expiresAt'
        $source | Should Match '_pendingRefreshSelectionRenames\.push_back\(std::move\(rename\)\)'
        $source | Should Not Match 'unresolvedRefreshesRemaining|kRefreshBudget'
    }

    It 'routes Commands long-path and high-cardinality fixtures through native TestSandbox roots' {
        $search = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
        $dialogs = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp'
        $viewCommands = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $search | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"search_index_stream"\)'
        $dialogs | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"pane_rename_long_selection"\)'
        $dialogs | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"item_properties_scroll"\)'
        $viewCommands | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"folder_sort_toggle_stress"\)'
        $viewCommands | Should Match 'WaitForFolderViewSortAndWarmRender\([\s\S]*SelfTest::Scale\(5000ms\)\)'
        $viewCommands | Should Match 'Sort-toggle stress did not reach sort-ready warm rendering'
        $viewCommands | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"viewer_space_synthetic_bucket_metrics"\)'
    }

    It 'routes ShellCommands shortcut save temp files through native TestSandbox roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ShellCommands.cpp'

        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands,\s*L"shell_shortcut_save_temp"\)'
        $source | Should Not Match 'GetTempPathW'
        $source | Should Not Match 'GetTempFileNameW'
    }

    It 'routes PerformanceTests2 scratch through the unified TestSandbox root' {
        $header = Get-RSText -Path 'Tests\PerformanceTests2\pch.h'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'
        $sources = @(
            'Tests\PerformanceTests2\FolderIconEnumerationPerfTest.cpp',
            'Tests\PerformanceTests2\FolderIconEnumerationDuplicatePathPerfTest.cpp',
            'Tests\PerformanceTests2\FolderViewRefreshDuplicatePathPerfTest.cpp'
        ) | ForEach-Object { Get-RSText -Path $_ }
        $combined = $sources -join "`n"

        $header | Should Match 'AcquirePerformanceTestSandbox'
        $header | Should Match 'kPerformanceTests2HarnessSegment\{L"performance-tests2"\}'
        $header | Should Match 'TestSupport::AcquireTestDirectory'
        $header | Should Match '\.emptyLeafFallback\s*=\s*L"default"'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kScratchDirectoryName'
        $combined | Should Match 'AcquirePerformanceTestSandbox\(L"folder_icon_enumeration_perf"'
        $combined | Should Match 'AcquirePerformanceTestSandbox\(L"folder_icon_enumeration_duplicate_path_perf"'
        $combined | Should Match 'AcquirePerformanceTestSandbox\(L"folder_view_refresh_duplicate_path_perf_shared"'
        $combined | Should Not Match 'std::filesystem::temp_directory_path'
    }

    It 'routes BatchRename window fixture roots through native TestSandbox roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.BatchRename.cpp'

        $source | Should Match 'AcquireBatchRenameCommandsSandboxRoot'
        $source | Should Match 'SelfTest::AcquireTestSandbox\(SelfTest::SelfTestSuite::Commands'
        $source | Should Match 'batch_rename_window_open'
        $source | Should Match 'batch_rename_rules_preview'
        $source | Should Match 'batch_rename_duplicate_source_target_refresh'
        $source | Should Not Match 'L"C:\\\\BatchRename[A-Za-z]*SelfTest"'
        $source | Should Not Match 'const\s+std::filesystem::path\s+root\s*=\s*L"C:\\\\BatchRename'
    }

    It 'routes ViewerPETests fixture files through the unified TestSandbox root' {
        $source = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'

        $source | Should Match 'AcquireViewerPETestSandbox'
        $source | Should Match 'kViewerPEHarnessSegment\{L"viewer-pe"\}'
        $source | Should Match 'TestSupport::AcquireTestDirectory'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kScratchDirectoryName'
        $source | Should Not Match 'buildDir\s*/\s*L"Viewer(Web|ImgRaw|ImgRawPng|Text|Space|VLC)Tests"'
    }

    It 'routes DxUiTests generated artifact defaults through the unified TestSandbox root' {
        $main = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.cpp'
        $helpers = Get-RSText -Path 'Tests\DxUiTests\DxUiTestHelpers.h'
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'
        $animation = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Animation.cpp'
        $windowHost = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.WindowHost.cpp'

        $helpers | Should Match 'GetDxUiTestArtifactPath'
        $helpers | Should Match 'GetDxUiTestArtifactDirectory'
        $helpers | Should Match 'TestSupport::AcquireTestDirectory'
        $helpers | Should Match 'TestDirectoryKind::Artifacts'
        $helpers | Should Match '\.includeLeafSegment\s*=\s*false'
        $helpers | Should Match '\.cleanExisting\s*=\s*false'
        $support | Should Match 'REDSALAMANDER_TEST_ROOT'
        $support | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $support | Should Match 'kRunsDirectoryName'
        $support | Should Match 'kArtifactsDirectoryName'
        $main | Should Match 'GetDxUiTestArtifactPath\(L"DxUiControlGallery\.png"\)'
        $main | Should Match 'GetDxUiTestArtifactPath\(L"DxUiButtonContrast\.png"\)'
        $animation | Should Match 'GetDxUiTestArtifactPath\(L"dxui_animation_scheduler_testlocal\.jsonl"\)'
        $windowHost | Should Match 'GetDxUiTestArtifactPath\(L"dxui_windowhost_stage_metrics_testlocal\.jsonl"\)'
        $main | Should Not Match 'Specs" / L"TestRuns" / L"DxUiGallery"'
        $animation | Should Not Match 'Specs" / L"TestRuns" / L"local_scratch"'
        $windowHost | Should Not Match 'Specs" / L"TestRuns" / L"local_scratch"'
    }

    It 'keeps shared TestSupport sandbox and environment policies explicit' {
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'

        $support | Should Match 'class ScopedEnvironmentVariable final'
        $support | Should Match 'ReadEnvironmentValue'
        $support | Should Match 'text\.size\(\) > 160u'
        $support | Should Match 'std::wstring_view emptyLeafFallback = L"case"'
        $support | Should Match 'bool includeLeafSegment\s*=\s*true'
        $support | Should Match 'bool cleanExisting\s*=\s*true'
        $support | Should Match 'if \(options\.cleanExisting\)[\s\S]*std::filesystem::remove_all'
    }

    It 'keeps viewer message pumping and typed snapshot polling on bounded shared support' {
        $support = Get-RSText -Path 'Tests\TestSupport\TestSupport.h'
        $viewerPe = Get-RSText -Path 'Tests\ViewerPETests\ViewerPETests.cpp'
        $viewerSqlite = Get-RSText -Path 'Tests\ViewerSqliteTests\ViewerSqliteTests.cpp'
        $windowHost = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.WindowHost.cpp'

        $support | Should Match 'size_t PumpPendingMessages\(size_t maxMessageCount = 1024u\)'
        $support | Should Match 'MessagePumpWaitResult PumpMessagesUntil'
        $support | Should Match 'timeoutDiagnostic\.append\(L" timed out after "\)'
        $support | Should Match 'bool WaitForSnapshot'
        $viewerPe | Should Match 'TestSupport::PumpMessagesUntil'
        $viewerPe | Should Match 'TestSupport::WaitForSnapshot<WndMsg::ViewerTextDebugSnapshot>'
        $viewerSqlite | Should Match 'TestSupport::PumpMessagesUntil'
        $viewerSqlite | Should Match 'TestSupport::WaitForSnapshot<WndMsg::ViewerSqliteDebugSnapshot>'
        $viewerPe | Should Not Match 'while \(std::chrono::steady_clock::now\(\) < deadline\)'
        $viewerSqlite | Should Not Match 'while \(std::chrono::steady_clock::now\(\) < deadline\)'
        $windowHost | Should Match 'TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling'
        $windowHost | Should Match 'WM_APP \+ 73u'
        $windowHost | Should Match 'budget 25 ms'
    }

    It 'routes MTP fake journal LOCALAPPDATA state through native TestSandbox roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp'

        $source | Should Match 'AcquireMtpJournalLocalAppDataSandbox'
        $source | Should Match 'SelfTest::AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories'
        $source | Should Match 'SetEnvironmentVariableW\(L"LOCALAPPDATA",\s*localAppData\.c_str\(\)\)'
        $source | Should Match 'L"mtp_journal_temp_retry"'
        $source | Should Match 'L"mtp_journal_no_temp_puid"'
        $source | Should Match 'L"mtp_overwrite_safety_matrix"'
        $source | Should Not Match 'GetEnvironmentVariableW\(L"LOCALAPPDATA"'
        $source | Should Not Match 'getLocalAppDataPath'
    }

    It 'routes Compare dummy filesystem scratch through native TestSandbox roots' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.CoreDiffs.cpp'

        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories,\s*L"dummy_content"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories,\s*L"normalized_name_collision_preserves_same_side_entries"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories,\s*L"deep_tree"\)'
        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories,\s*L"invalidate"\)'
        $source | Should Not Match 'CompareSelfTest_'
        $source | Should Not Match 'std::filesystem::path\(L"[A-Z]:\\\\\"\)\s*/\s*\(L"CompareSelfTest_"'
    }

    It 'routes Tools Pester scratch through the unified TestSandbox root' {
        $testRunPlan = Get-RSText -Path 'Tools\TestRunPlan.ps1'
        $scripts = @(
            'Tools\Tests\WingetValidation.Tests.ps1',
            'Tools\Tests\VcpkgInstallSafety.Tests.ps1',
            'Tools\Tests\ShowPerfRuns.Tests.ps1'
        ) | ForEach-Object { Get-RSText -Path $_ }
        $combined = $scripts -join "`n"

        $testRunPlan | Should Match 'function\s+New-RSTestSandboxScratchDirectory'
        $testRunPlan | Should Match '\$scratchRoot\s*=\s*Join-Path\s+\$runRoot\s+''scratch'''
        $testRunPlan | Should Match 'RSWingetValidationTest_\*'
        $testRunPlan | Should Match 'RSVcRuntimeTest_\*'
        $testRunPlan | Should Match 'rs-vcpkg-install-root-test\*'
        $testRunPlan | Should Match 'rs-vcpkg-single-file-merge-\*'
        $testRunPlan | Should Match 'rs-show-perfruns-tests-\*'

        $combined | Should Match 'New-RSTestSandboxScratchDirectory'
        $combined | Should Match "Harness 'tools-pester'"
        $combined | Should Not Match '\[System\.IO\.Path\]::GetTempPath\(\)'
    }

    It 'keeps gated test sources inside TestSandbox instead of raw temp or profile roots' {
        $forbiddenPatterns = [ordered]@{
            'std::filesystem::temp_directory_path' = 'std::filesystem::temp_directory_path\s*\('
            'GetTempPath' = '\bGetTempPath(?:W|A)?\s*\('
            'GetTempFileName' = '\bGetTempFileName(?:W|A)?\s*\('
            'PowerShell GetTempPath' = '\[System\.IO\.Path\]::GetTempPath\(\)'
            'LOCALAPPDATA getenv' = 'GetEnvironmentVariableW\s*\(\s*L"LOCALAPPDATA"'
            'C getenv temp/profile' = 'std::getenv\s*\(\s*"(?:LOCALAPPDATA|TEMP|TMP)"'
            'legacy Compare drive root' = 'CompareSelfTest_'
            'legacy cross-volume drive root' = 'RedSalamanderCrossVolumeSelfTest_'
            'legacy BatchRename drive root' = '[A-Z]:\\+BatchRename[A-Za-z]*SelfTest'
            'legacy DxUiGallery artifact root' = 'Specs\\TestRuns\\DxUiGallery'
            'legacy local_scratch artifact root' = 'Specs\\TestRuns\\local_scratch'
        }

        $syntheticViolation = @'
std::filesystem::temp_directory_path();
GetTempPathW(0, nullptr);
GetTempFileNameW(nullptr, nullptr, 0, nullptr);
[System.IO.Path]::GetTempPath()
GetEnvironmentVariableW(L"LOCALAPPDATA", nullptr, 0);
std::getenv("TEMP");
Y:\CompareSelfTest_dummy
C:\RedSalamanderCrossVolumeSelfTest_123
C:\\BatchRenameWindowSelfTest
Specs\TestRuns\DxUiGallery
Specs\TestRuns\local_scratch
'@
        foreach ($entry in $forbiddenPatterns.GetEnumerator()) {
            $syntheticViolation | Should Match $entry.Value
        }

        $findings = foreach ($file in Get-RSTestSourceContractFiles) {
            $text = Get-Content -LiteralPath $file.FullName -Raw
            $relative = Get-RSRelativeTestSourcePath -Root $repoRoot -Path $file.FullName
            foreach ($entry in $forbiddenPatterns.GetEnumerator()) {
                if ($relative -ieq 'Tools\Tests\RunAllTestsPlan.Tests.ps1' -and $entry.Key -like 'legacy *') {
                    continue
                }

                if ($text -match $entry.Value) {
                    "$relative contains forbidden test root acquisition pattern: $($entry.Key)"
                }
            }
        }

        ($findings -join "`n") | Should Be ''
    }

    It 'routes SettingsStore app files through the unified TestSandbox root during test runs' {
        $source = Get-RSText -Path 'Common\Common\SettingsStore.cpp'

        $source | Should Match 'REDSALAMANDER_TEST_ROOT'
        $source | Should Match 'REDSALAMANDER_TEST_RUN_ID'
        $source | Should Match 'settings-store'
        $source | Should Match 'runs'
        $source | Should Match 'scratch'
        $source | Should Match 'GetUnifiedTestSettingsDirectoryPathFromEnvironment'
        $source | Should Match 'GetSettingsDirectoryPath\(\)[\s\S]*GetUnifiedTestSettingsDirectoryPathFromEnvironment\(\)'
    }

    It 'keeps SettingsStore TestSandbox path overrides behind ENABLE_TESTS' {
        $source = Get-RSText -Path 'Common\Common\SettingsStore.cpp'

        $source | Should Match '#ifdef ENABLE_TESTS[\s\S]*GetUnifiedTestSettingsDirectoryPathFromEnvironment'
        $source | Should Match 'GetSettingsDirectoryPath\(\)[\s\S]*#ifdef ENABLE_TESTS[\s\S]*GetUnifiedTestSettingsDirectoryPathFromEnvironment\(\)[\s\S]*#endif'
    }

    It 'moves Compare foreground service capture files into TestSandbox instead of process temp' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'

        $source | Should Match 'AcquireTestSandbox\(SelfTest::SelfTestSuite::CompareDirectories,\s*L"foreground_search_service_stdout"'
        $source | Should Match 'foreground_search_service_stdout'
        $source | Should Not Match 'GetTempPathW'
        $source | Should Not Match 'GetTempFileNameW'
    }

    It 'wraps foreground search service selftest processes in a kill-on-close JobObject' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'

        $source | Should Match 'CreateKillOnCloseJob'
        $source | Should Match 'JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE'
        $source | Should Match 'SetInformationJobObject'
        $source | Should Match 'AssignProcessToJobObject'
        $source | Should Match 'CREATE_SUSPENDED'
        $source | Should Match 'ResumeThread'
        $source | Should Match 'wil::unique_handle\s+_job'
        $source | Should Match '_job\.reset\(\)'
        $source | Should Match 'Foreground search service JobObject'

        $readyStart = $source.IndexOf('[[nodiscard]] bool WaitUntilReady')
        $readyNext = $source.IndexOf('[[nodiscard]] bool DrainCapturedOutput', $readyStart)
        $readyStart | Should BeGreaterThan -1
        $readyNext | Should BeGreaterThan $readyStart

        $readyBlock = $source.Substring($readyStart, $readyNext - $readyStart)
        $readyBlock | Should Match 'WaitForPipeReady\(outError\)'
        $readyBlock | Should Match 'SearchServiceBroker::GetStatus\(status\)'
        $readyBlock | Should Match 'EqualsIgnoreCase\(status\.pipeName,\s*_pipeName\)'
        $readyBlock | Should Match 'expectedPipe=''\{\}'' lastPipe=''\{\}'''
        $source | Should Match 'SearchServiceBroker::RequestShutdown\(_pipeName,\s*timeoutMs\)'
        $source | Should Match 'ShutdownAndWaitForExitAndCapture'
        $source | Should Not Match '--max-requests='
    }

    It 'keeps foreground Search service stdout tests independent from hidden readiness requests' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'

        $idleStart = $source.IndexOf('L"search_service_sqlite_idle_maintenance_queue_and_completion"')
        $idleNext = $source.IndexOf('L"search_service_sqlite_delete_burst_maintenance_preserves_query_parity"', $idleStart)
        $idleStart | Should BeGreaterThan -1
        $idleNext | Should BeGreaterThan $idleStart
        $idleBlock = $source.Substring($idleStart, $idleNext - $idleStart)
        $idleBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $idleBlock | Should Match 'SearchServiceBroker::GetStatus\(queuedStatus\)'
        $idleBlock | Should Match 'SearchServiceBroker::GetStatus\(completedStatus\)'
        $idleBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'

        $invalidStart = $source.IndexOf('L"search_service_sqlite_invalid_store_falls_back_live_scan"')
        $invalidNext = $source.IndexOf('L"search_service_sqlite_query_failure_falls_back_live_scan"', $invalidStart)
        $invalidStart | Should BeGreaterThan -1
        $invalidNext | Should BeGreaterThan $invalidStart
        $invalidBlock = $source.Substring($invalidStart, $invalidNext - $invalidStart)
        $invalidBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $invalidBlock | Should Match 'SearchServiceBroker::GetStatus\(status\)'
        $invalidBlock | Should Match 'SearchServiceBroker::Query\(request'
        $invalidBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'

        $pendingStart = $source.IndexOf('L"search_service_sqlite_status_reports_pending_legacy_import"')
        $pendingNext = $source.IndexOf('L"search_service_binary_uses_console_subsystem"', $pendingStart)
        $pendingStart | Should BeGreaterThan -1
        $pendingNext | Should BeGreaterThan $pendingStart
        $pendingBlock = $source.Substring($pendingStart, $pendingNext - $pendingStart)
        $pendingBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $pendingBlock | Should Match 'SearchServiceBroker::GetStatus\(status\)'
        $pendingBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'

        $warmStart = $source.IndexOf('L"search_service_sqlite_startup_warms_overridden_roots"')
        $warmNext = $source.IndexOf('L"search_service_sqlite_startup_warmup_failure_status_roundtrip"', $warmStart)
        $warmStart | Should BeGreaterThan -1
        $warmNext | Should BeGreaterThan $warmStart
        $warmBlock = $source.Substring($warmStart, $warmNext - $warmStart)
        $warmBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $warmBlock | Should Match 'SearchServiceBroker::GetStatus\(status\)'
        $warmBlock | Should Match 'SearchServiceBroker::Query\(request'
        $warmBlock | Should Match 'SearchServiceBroker::GetStatus\(afterQueryStatus\)'
        $warmBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'

        $logStart = $source.IndexOf('L"search_service_foreground_logs_request_status"')
        $logNext = $source.IndexOf('L"search_service_query_reports_live_progress"', $logStart)
        $logStart | Should BeGreaterThan -1
        $logNext | Should BeGreaterThan $logStart
        $logBlock = $source.Substring($logStart, $logNext - $logStart)
        $logBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $logBlock | Should Match 'SearchServiceBroker::Query\(request'
        $logBlock | Should Match 'SearchServiceBroker::GetStatus\(status\)'
        $logBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'
    }

    It 'keeps Compare live compact requests bound to an explicit pipe with captured-output diagnostics' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'

        $testStart = $source.IndexOf('L"search_service_compact_request_roundtrip"')
        $testNext = $source.IndexOf('L"search_service_sqlite_status_reports_maintenance_history"', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'kExplicitCompactLargeEntryCount\s*=\s*2048u'
        $testBlock | Should Match 'kExplicitCompactRetainedEntryCount\s*=\s*1800u'
        $testBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $testBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*preCompactStatus,\s*pipeName,\s*L"live compact request preflight service"\)'
        $testBlock | Should Match '! preCompactStatus\.maintenanceQueued'
        $testBlock | Should Match '! preCompactStatus\.maintenanceRunning'
        $testBlock | Should Match 'requestCompactArguments\{L"--request-compact",\s*std::format\(L"--pipe-name=\{\}",\s*pipeName\)\}'
        $testBlock | Should Match 'RunProcessAndCaptureOutput\([\s\S]*servicePath\.wstring\(\),\s*requestCompactArguments'
        $testBlock | Should Match 'requestCompactOutput\(result\.output\.begin\(\),\s*result\.output\.end\(\)\)'
        $testBlock | Should Match 'Expected --request-compact exit code 0, got \{\}\. output=''\{\}'''
        $testBlock | Should Match 'service\.ShutdownAndWaitForExitAndCapture'
        $testBlock | Should Match 'serviceOutput\.contains\("Maintenance running"\)'
        $testBlock | Should Match 'serviceOutput\.contains\("Maintenance completed"\)'
        $testBlock | Should Not Match 'RunProcessAndCaptureOutput\(servicePath\.wstring\(\),\s*L"--request-compact"'
    }

    It 'keeps the shared child-process runner contained, concurrently drained, and parity-tested' {
        $header = Get-RSText -Path 'Tests\TestSupport\ChildProcess.h'
        $compare = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $cases = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'
        $probe = Get-RSText -Path 'RedSalamanderSearchService\Main.cpp'

        $header | Should Match 'JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE'
        $header | Should Match 'PROC_THREAD_ATTRIBUTE_HANDLE_LIST'
        $header | Should Match 'CREATE_SUSPENDED \| EXTENDED_STARTUPINFO_PRESENT'
        $header | Should Match 'AssignProcessToJobObject[\s\S]*ResumeThread'
        $header | Should Match 'std::jthread stdoutDrain'
        $header | Should Match 'std::jthread stderrDrain'
        $header | Should Match 'maxStdoutBytes'
        $header | Should Match 'maxStderrBytes'
        $header | Should Match 'QuoteWindowsCommandLineArgument'
        $compare | Should Match 'TestSupport::RunChildProcess'
        $compare | Should Not Match 'WaitForSingleObject\(process\.get\(\),\s*timeoutMs\)[\s\S]*ReadFile\(readPipe\.get\(\)'
        $cases | Should Match 'L"test_support_child_process_runner_contract"'
        $cases | Should Match 'ChildProcessStatus::TimedOut'
        $cases | Should Match 'ChildProcessStatus::Cancelled'
        $cases | Should Match 'HANDLE_CLOSED'
        $probe | Should Match '--test-support-child-probe=delayed-marker'
        $probe | Should Match '#ifdef ENABLE_TESTS[\s\S]*TryRunTestSupportChildProbe'
    }

    It 'keeps local index snapshot reload timing advisory instead of a correctness ceiling' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'

        $source | Should Match 'compare\.selftest\.local_index\.snapshot_reload_us'
        $source | Should Match 'Debug::Perf::Emit'
        $source | Should Not Match 'warmElapsedMs\s*<\s*1000u'
        $source | Should Not Match 'Warm indexed query took too long'
    }

    It 'keeps SQLite traversal-seed currentness-unproven state from being downgraded by metadata probes' {
        $source = Get-RSText -Path 'Common\LocalSearchIndexCore.cpp'

        $source | Should Match 'ResolveSqliteVolumeStateForMetadataCompleteness'
        $source | Should Match 'desiredState == SqliteIndexStore::kVolumeStateCurrentnessUnproven'
        $source | Should Match 'BuildSqliteReplaceRequest[\s\S]*ResolveSqliteVolumeStateForMetadataCompleteness\(desiredState,\s*outMetadataComplete\)'
        $source | Should Match 'BuildSqliteApplyJournalDeltaRequest[\s\S]*ResolveSqliteVolumeStateForMetadataCompleteness\(outRequest\.state,\s*metadataComplete\)'
        $source | Should Match 'BuildSqliteApplyJournalDeltaRequest[\s\S]*ResolveSqliteVolumeStateForMetadataCompleteness\(outRequest\.state,\s*seedMetadataComplete\)'
        $source | Should Match 'SqliteVolumeStore[\s\S]*ResolveSqliteVolumeStateForMetadataCompleteness\(state\.state,\s*metadataComplete\)'
    }

    It 'keeps foreground Search service status probes behind a condition-based status wait' {
        $harness = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'

        $harness | Should Match 'WaitForSearchServiceStatus'
        $harness | Should Match 'SearchServiceBroker::GetStatus\(candidate\)'
        $harness | Should Match 'SUCCEEDED\(lastHr\)\s*&&\s*\(expectedPipeName\.empty\(\)\s*\|\|\s*EqualsIgnoreCase\(candidate\.pipeName,\s*expectedPipeName\)\)'
        $harness | Should Match "expectedPipe='\{\}'"
        $harness | Should Match 'SelfTest::ScaleTimeout\(timeoutMs\)'
        $harness | Should Match 'kForegroundSearchServiceReadinessTimeoutMs\s*=\s*30''000u'
        $harness | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics'
        $harness | Should Match 'TryCaptureExitedOutputForFailure'
        $harness | Should Match 'Foreground search service exited before SearchServiceBroker::GetStatus'

        $maintenanceStart = $source.IndexOf('L"search_service_sqlite_status_reports_maintenance_history"')
        $maintenanceNext = $source.IndexOf('L"search_service_sqlite_idle_maintenance_queue_and_completion"', $maintenanceStart)
        $maintenanceStart | Should BeGreaterThan -1
        $maintenanceNext | Should BeGreaterThan $maintenanceStart

        $maintenanceBlock = $source.Substring($maintenanceStart, $maintenanceNext - $maintenanceStart)
        $maintenanceBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*status,\s*pipeName,\s*L"maintenance-history SQLite service"\)'
        $maintenanceBlock | Should Not Match 'hr\s*=\s*SearchServiceBroker::GetStatus\(status\);\s*state\.Require\(SUCCEEDED\(hr\)'

        $rootDiscoveryStart = $source.IndexOf('L"search_service_discovers_fixed_local_roots_on_start"')
        $rootDiscoveryNext = $source.IndexOf('L"search_service_sqlite_startup_warms_overridden_roots"', $rootDiscoveryStart)
        $rootDiscoveryStart | Should BeGreaterThan -1
        $rootDiscoveryNext | Should BeGreaterThan $rootDiscoveryStart

        $rootDiscoveryBlock = $source.Substring($rootDiscoveryStart, $rootDiscoveryNext - $rootDiscoveryStart)
        $rootDiscoveryBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*status,\s*pipeName,\s*L"root-discovery service"\)'
        $rootDiscoveryBlock | Should Not Match 'hr\s*=\s*SearchServiceBroker::GetStatus\(status\);\s*state\.Require\(SUCCEEDED\(hr\)'

        $defaultStoreStart = $source.IndexOf('L"search_service_sqlite_default_store_uses_build_specific_programdata_root"')
        $seededDefaultStoreStart = $source.IndexOf('L"search_service_sqlite_seeded_default_store_reuses_build_specific_programdata_root"', $defaultStoreStart)
        $defaultStoreStart | Should BeGreaterThan -1
        $seededDefaultStoreStart | Should BeGreaterThan $defaultStoreStart
        $rootDiscoveryStart | Should BeGreaterThan $seededDefaultStoreStart

        $defaultStoreBlock = $source.Substring($defaultStoreStart, $seededDefaultStoreStart - $defaultStoreStart)
        $defaultStoreBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*L"--store-backend=sqlite"\)'
        $defaultStoreBlock | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics\(state,\s*status,\s*pipeName,\s*L"default-store SQLite service",\s*service\)'
        $defaultStoreBlock | Should Not Match 'SearchServiceBroker::GetStatus\(status\)'

        $seededDefaultStoreBlock = $source.Substring($seededDefaultStoreStart, $rootDiscoveryStart - $seededDefaultStoreStart)
        $seededDefaultStoreBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*L"--store-backend=sqlite"\)'
        $seededDefaultStoreBlock | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics\(state,\s*status,\s*pipeName,\s*L"seeded default-store SQLite service",\s*service\)'
        $seededDefaultStoreBlock | Should Not Match 'statusDeadline'
        $seededDefaultStoreBlock | Should Not Match 'SearchServiceBroker::GetStatus\(status\)'

        $latencyStart = $source.IndexOf('L"local_search_service_indexed_name_latency_and_parity"')
        $latencyNext = $source.IndexOf('L"local_search_service_matches_host_fallback"', $latencyStart)
        $latencyStart | Should BeGreaterThan -1
        $latencyNext | Should BeGreaterThan $latencyStart

        $latencyBlock = $source.Substring($latencyStart, $latencyNext - $latencyStart)
        $latencyBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $latencyBlock | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics\(state,\s*readyStatus,\s*pipeName,\s*L"indexed name latency service",\s*service\)'
        $latencyBlock.IndexOf('WaitForSearchServiceStatusWithProcessDiagnostics') | Should BeLessThan $latencyBlock.IndexOf('SearchServiceBroker::Query(brokerRequest')
        $latencyBlock | Should Match 'Warmup broker query failed\. hr=0x\{:08X\}\{\}'

        $deviceRootStart = $source.IndexOf('L"search_service_rejects_device_root_and_continues"')
        $deviceRootNext = $source.IndexOf('L"search_service_slow_partial_client_does_not_block_next_client"', $deviceRootStart)
        $deviceRootStart | Should BeGreaterThan -1
        $deviceRootNext | Should BeGreaterThan $deviceRootStart

        $deviceRootBlock = $source.Substring($deviceRootStart, $deviceRootNext - $deviceRootStart)
        $deviceRootBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*status,\s*pipeName,\s*L"device-root rejection recovery service"\)'
        $deviceRootBlock | Should Not Match 'const HRESULT statusHr = SearchServiceBroker::GetStatus\(status\)'

        $slowPartialStart = $source.IndexOf('L"search_service_slow_partial_client_does_not_block_next_client"')
        $slowPartialNext = $source.IndexOf('L"search_service_sqlite_legacy_auto_vacuum_queues_idle_maintenance"', $slowPartialStart)
        $slowPartialStart | Should BeGreaterThan -1
        $slowPartialNext | Should BeGreaterThan $slowPartialStart

        $slowPartialBlock = $source.Substring($slowPartialStart, $slowPartialNext - $slowPartialStart)
        $slowPartialBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $slowPartialBlock | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics\(state,\s*status,\s*pipeName,\s*L"slow partial client recovery service",\s*service\)'
        $slowPartialBlock | Should Not Match 'const HRESULT statusHr = SearchServiceBroker::GetStatus\(status\)'

        $legacyAutoVacuumStart = $slowPartialNext
        $legacyAutoVacuumNext = $source.IndexOf('L"sqlite_index_store_upgrade_paths"', $legacyAutoVacuumStart)
        $legacyAutoVacuumNext | Should BeGreaterThan $legacyAutoVacuumStart

        $legacyAutoVacuumBlock = $source.Substring($legacyAutoVacuumStart, $legacyAutoVacuumNext - $legacyAutoVacuumStart)
        $legacyAutoVacuumBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*queuedStatus,\s*pipeName,\s*L"legacy auto_vacuum queued maintenance service"\)'
        $legacyAutoVacuumBlock | Should Match 'SelfTest::ScaleTimeout\(30''000\)'
        $legacyAutoVacuumBlock | Should Match '! completedStatus\.maintenanceQueued\s*&&\s*! completedStatus\.maintenanceRunning'
        $legacyAutoVacuumBlock | Should Not Match 'HRESULT hr = SearchServiceBroker::GetStatus\(queuedStatus\)'

        $testStart = $source.IndexOf('L"search_service_status_and_query_roundtrip"')
        $testNext = $source.IndexOf('L"local_search_service_single_request_uses_query"', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*status,\s*pipeName,\s*L"status roundtrip service"\)'
        $testBlock | Should Not Match 'hr\s*=\s*SearchServiceBroker::GetStatus\(status\);\s*state\.Require\(SUCCEEDED\(hr\)'

        $multiClientStart = $source.IndexOf('L"search_service_multi_client_and_rebuild_control"')
        $multiClientNext = $source.IndexOf('L"search_service_rebuild_deleted_root_purges_index"', $multiClientStart)
        $multiClientStart | Should BeGreaterThan -1
        $multiClientNext | Should BeGreaterThan $multiClientStart

        $multiClientBlock = $source.Substring($multiClientStart, $multiClientNext - $multiClientStart)
        $multiClientBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion'
        $multiClientBlock | Should Match 'WaitForSearchServiceStatus\(state,\s*initialStatus,\s*pipeName,\s*L"multi-client rebuild service"\)'
        $multiClientBlock | Should Match 'std::jthread clientA'

        $deletedRootStart = $multiClientNext
        $deletedRootNext = $source.IndexOf('L"search_text_helpers_decoding_and_binary"', $deletedRootStart)
        $deletedRootNext | Should BeGreaterThan $deletedRootStart

        $deletedRootBlock = $source.Substring($deletedRootStart, $deletedRootNext - $deletedRootStart)
        $deletedRootBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*true,\s*serviceError,\s*true,\s*extraArgs\)'
        $deletedRootBlock | Should Match 'WaitForSearchServiceStatusWithProcessDiagnostics\(state,\s*readyStatus,\s*pipeName,\s*L"deleted-root rebuild service",\s*service\)'
        $deletedRootBlock.IndexOf('WaitForSearchServiceStatusWithProcessDiagnostics') | Should BeLessThan $deletedRootBlock.IndexOf('SearchServiceBroker::Query(request')
    }

    It 'lets delete-burst maintenance tests accept an already completed first idle window' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp'

        $testStart = $source.IndexOf('L"search_service_sqlite_delete_burst_maintenance_preserves_query_parity"')
        $testNext = $source.IndexOf('L"search_service_sqlite_bootstrap_status_roundtrip"', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'formatMaintenanceStatus'
        $testBlock | Should Match 'service\.Start\(pipeName,\s*0u,\s*SearchServiceBroker::kProtocolVersion,\s*false,\s*serviceError,\s*true,\s*extraArgs\)'
        $testBlock | Should Match 'SelfTest::ScaleTimeout\(kForegroundSearchServiceReadinessTimeoutMs\)'
        $testBlock | Should Match 'service\.TryCaptureExitedOutputForFailure\(\)'
        $testBlock | Should Match 'status\.maintenanceQueued\s*&&\s*status\.persistentStoreFreelistPageCount\s*>=\s*beforeInfo\.freelistPageCount'
        $testBlock | Should Match 'status\.persistentStoreFreelistPageCount\s*<\s*beforeInfo\.freelistPageCount'
        $testBlock | Should Match 'firstWindowAlreadyObserved\s*=\s*initialStatus\.persistentStoreFreelistPageCount\s*<\s*beforeInfo\.freelistPageCount'
        $testBlock | Should Match 'afterFirstWindow\s*=\s*initialStatus'
        $testBlock | Should Match 'afterFirstWindow\.persistentStoreFreelistPageCount\s*<\s*beforeInfo\.freelistPageCount'
        $testBlock | Should Match 'afterFirstWindow\.persistentStoreBytes\s*<\s*beforeInfo\.databaseBytes'
        $testBlock | Should Not Match 'afterFirstWindow\.persistentStoreBytes\s*<\s*initialStatus\.persistentStoreBytes'
    }

    It 'scales named raw self-test waits through the timeout multiplier' {
        $search = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
        $batchRename = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.BatchRename.cpp'
        $fileOps = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $fileOpsPhase05 = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Phases05_06.cpp'

        $waitForFlagMatch = [regex]::Match($search, '\[\[nodiscard\]\]\s+bool\s+WaitForFlag\([\s\S]*?\n\}')
        $waitForFlagMatch.Success | Should Be $true
        $waitForFlag = $waitForFlagMatch.Value
        $waitForFlag | Should Match 'SelfTest::ScaleTimeout\(timeoutMs\)'
        $waitForFlag | Should Not Match 'GetTickCount64\(\)\s*\+\s*timeoutMs'
        $search | Should Match 'WaitForFlag\(callback\.firstCallbackEntered,\s*5000u\)'

        $batchRename | Should Match 'WaitForSingleObject\(_renameItemsGate,\s*static_cast<DWORD>\(SelfTest::ScaleTimeout\(30''000u\)\)\)'
        $batchRename | Should Not Match 'WaitForSingleObject\(_renameItemsGate,\s*30000u\)'

        $fileOps | Should Match 'constexpr\s+DWORD\s+kRetryDelayMs\s*=\s*50u'
        $fileOps | Should Match 'GetTickCount64\(\)\s*\+\s*SelfTest::ScaleTimeout\(6''000u\)'
        $fileOps | Should Match 'while\s*\(::GetTickCount64\(\)\s*<\s*deadline\)'
        $fileOps | Should Not Match 'kMaxAttempts\s*=\s*120'
        $fileOps | Should Match 'TickPhases05To06\(SelfTestState& state\)'
        $fileOps | Should Match 'TickPhases07To09\(SelfTestState& state\)'
        $fileOps | Should Match 'placing every included phase in this one switch made Tick reserve roughly 780 KiB'

        $popupSmokeStart = $fileOpsPhase05.IndexOf('case SelfTestState::Step::Phase6_PopupSmokeResizeAndPause:')
        $popupSmokeNext = $fileOpsPhase05.IndexOf('case SelfTestState::Step::Phase6_DeleteBytesMeaningful:', $popupSmokeStart)
        $popupSmokeStart | Should BeGreaterThan -1
        $popupSmokeNext | Should BeGreaterThan $popupSmokeStart
        $popupSmokeBlock = $fileOpsPhase05.Substring($popupSmokeStart, $popupSmokeNext - $popupSmokeStart)
        $popupSmokeBlock | Should Match 'kSourceFileCount\s*=\s*16u'
        $popupSmokeBlock | Should Match 'std::vector<std::filesystem::path>\s+sources'
        $popupSmokeBlock | Should Match 'observed resumed progress items='

        $queuedCancelStart = $fileOpsPhase05.IndexOf('case SelfTestState::Step::Phase5_CancelQueuedTask:')
        $queuedCancelNext = $fileOpsPhase05.IndexOf('case SelfTestState::Step::Phase5_SwitchParallelToWaitDuringPreCalc:', $queuedCancelStart)
        $queuedCancelStart | Should BeGreaterThan -1
        $queuedCancelNext | Should BeGreaterThan $queuedCancelStart

        $queuedCancelBlock = $fileOpsPhase05.Substring($queuedCancelStart, $queuedCancelNext - $queuedCancelStart)
        $queuedCancelBlock | Should Match 'Phase5_CancelQueuedTask timed out\. stepState=\{\} A: \{\} B: \{\} C: \{\}'
        $queuedCancelBlock | Should Match 'kQueueHoldSpeedLimitBytesPerSecond\s*=\s*8ull\s*\*\s*1024ull'
        $queuedCancelBlock | Should Match 'taskA->SetDesiredSpeedLimit\(kQueueHoldSpeedLimitBytesPerSecond\)'
        $queuedCancelBlock | Should Match 'taskA->SkipPreCalculation\(\)'
        $queuedCancelBlock | Should Match 'Active queue holder completed before the queued task could be created'
        $queuedCancelBlock | Should Match '! taskA->HasEnteredOperation\(\)'
        $queuedCancelBlock | Should Match 'Failed to start the queued dummy copy task for queued-cancel test'
        $queuedCancelBlock | Should Match 'Queued task entered operation before queued cancellation'
        $queuedCancelBlock | Should Match 'if\s*\(taskB->IsWaitingInQueue\(\)\)\s*\{\s*taskB->RequestCancel\(\);\s*state\.stepState\s*=\s*3;'
        $queuedCancelBlock | Should Match 'Queued task cancellation expected cancel hr'
        $queuedCancelBlock | Should Match 'Active queue holder expected cancel hr'
        $queuedCancelBlock | Should Not Match 'Fail\(L"Phase5_CancelQueuedTask timed out\."\)'

        $activeEnteredIndex = $queuedCancelBlock.IndexOf('if (! taskA->HasEnteredOperation())')
        $queuedStartIndex = $queuedCancelBlock.IndexOf('state.taskB = StartFileOperationAndGetId')
        $queuedObservedIndex = $queuedCancelBlock.IndexOf('if (taskB->IsWaitingInQueue())')
        $queuedCancelIndex = $queuedCancelBlock.IndexOf('taskB->RequestCancel()', $queuedObservedIndex)
        $activeEnteredIndex | Should BeGreaterThan -1
        $queuedStartIndex | Should BeGreaterThan $activeEnteredIndex
        $queuedObservedIndex | Should BeGreaterThan $queuedStartIndex
        $queuedCancelIndex | Should BeGreaterThan $queuedObservedIndex
    }

    It 'keeps Commands UIA action helpers deterministic under DxUi dispatch timeouts' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        $search = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'

        $threadContextStart = $source.IndexOf('struct UiaThreadContext final')
        $lifetimeHelperStart = $source.IndexOf('struct UiaOperationLifetime final', $threadContextStart)
        $runHelperStart = $source.IndexOf('[[nodiscard]] bool RunUiaActionWithMessagePump', $lifetimeHelperStart)
        $rawSetterStart = $source.IndexOf('[[nodiscard]] bool SetWindowHostRawProviderValueByNameWithMessagePump', $runHelperStart)
        $visibleSetterStart = $source.IndexOf('[[nodiscard]] bool SetVisibleDescendantValueByNameWithMessagePump', $rawSetterStart)
        $threadContextStart | Should BeGreaterThan -1
        $lifetimeHelperStart | Should BeGreaterThan $threadContextStart
        $runHelperStart | Should BeGreaterThan -1
        $rawSetterStart | Should BeGreaterThan $runHelperStart
        $visibleSetterStart | Should BeGreaterThan $rawSetterStart

        $threadLifetimeBlock = $source.Substring($threadContextStart, $runHelperStart - $threadContextStart)
        $runHelperBlock = $source.Substring($runHelperStart, $rawSetterStart - $runHelperStart)
        $rawSetterBlock = $source.Substring($rawSetterStart, $visibleSetterStart - $rawSetterStart)

        $threadLifetimeBlock | Should Match 'IUIAutomation2'
        $threadLifetimeBlock | Should Match 'CLSID_CUIAutomation8'
        $threadLifetimeBlock | Should Match 'connectionHr\s*=\s*timeoutAutomation->put_ConnectionTimeout\(timeoutMs\)'
        $threadLifetimeBlock | Should Match 'transactionHr\s*=\s*timeoutAutomation->put_TransactionTimeout\(timeoutMs\)'
        $threadLifetimeBlock | Should Match 'return\s+SUCCEEDED\(connectionHr\)\s*&&\s*SUCCEEDED\(transactionHr\)'
        $threadLifetimeBlock | Should Match 'CoEnableCallCancellation\(nullptr\)'
        $threadLifetimeBlock | Should Match 'CoCancelCall\(workerThreadId,\s*0u\)'
        $threadLifetimeBlock | Should Match 'struct\s+UiaDeadlineBudget\s+final'
        $threadLifetimeBlock | Should Match 'operationDeadline\s*=\s*started\s*\+\s*deadlineBudget\.operation'
        $threadLifetimeBlock | Should Match 'overallDeadline\s*=\s*started\s*\+\s*deadlineBudget\.total'
        $threadLifetimeBlock | Should Match 'if\s*\(! GetThreadUiAutomationContext\(\)\.ConfigureTimeouts\(deadlineBudget\.clientTimeoutMs\)\)'
        $threadLifetimeBlock | Should Match 'if\s*\(! lifetime\.done\.load[\s\S]*std::terminate\(\)'
        $threadLifetimeBlock | Should Match 'worker\.join\(\)'
        $threadLifetimeBlock | Should Not Match 'worker\.detach\('

        $runHelperBlock | Should Match 'kUiaActionDispatchTimeoutBudgetMs'
        $runHelperBlock | Should Match 'WaitForBoundedUiaWorker'
        $runHelperBlock | Should Match 'kBlockedOperationBudgetMs\s*=\s*1000u'
        $runHelperBlock | Should Match 'deadlineBudget\s*=\s*MakeUiaDeadlineBudget\(kBlockedOperationBudgetMs\)'
        $runHelperBlock | Should Match 'blockingDuration\s*=\s*deadlineBudget\.operation\s*\+\s*\(cancellationReserve\s*/\s*2\)'
        $runHelperBlock | Should Match 'sleep_for\(blockingDuration\)'
        $runHelperBlock | Should Match 'return\s+true;'
        $runHelperBlock | Should Match 'elapsed\s*<\s*deadlineBudget\.total'
        $runHelperBlock | Should Not Match 'worker\.detach\('
        $source | Should Match 'settings_uia_blocked_operation_deadline_is_bounded'

        $rawSetterBlock | Should Match 'GetWindowThreadProcessId\(hwnd,\s*nullptr\)\s*==\s*GetCurrentThreadId\(\)'
        $rawSetterBlock | Should Match 'SetWindowHostRawProviderValueByName\(hwnd'
        $rawSetterBlock.IndexOf('SetWindowHostRawProviderValueByName(hwnd') | Should BeLessThan $rawSetterBlock.IndexOf('RunUiaActionWithMessagePump')

        $restoredActionStart = $search.IndexOf('[[nodiscard]] bool TestFindDialogRestoredCombinedViewStateActionButtonsActivateSelection')
        $themeCycleStart = $search.IndexOf('[[nodiscard]] bool TestFindDialogThemeCycleKeepsGridLegible', $restoredActionStart)
        $restoredActionStart | Should BeGreaterThan -1
        $themeCycleStart | Should BeGreaterThan $restoredActionStart

        $restoredActionBlock = $search.Substring($restoredActionStart, $themeCycleStart - $restoredActionStart)
        $restoredActionBlock | Should Match 'InvokeVisibleDescendantByNameWithMessagePump\(\s*findWindow,\s*UIA_ButtonControlTypeId,\s*openButtonText,\s*L"restored combined Find Open action button'
        $restoredActionBlock | Should Match 'InvokeVisibleDescendantByNameWithMessagePump\(\s*findWindow,\s*UIA_ButtonControlTypeId,\s*parentButtonText,\s*L"restored combined Find Go to folder action button'
        $restoredActionBlock | Should Not Match 'InvokeVisibleDescendantByName\(\s*findWindow,\s*UIA_ButtonControlTypeId,\s*(openButtonText|parentButtonText)\)'
    }

    It 'suppresses mismatched live cursor points during persistent View-to-Plugins menu hover selftests' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuPersistentViewToPluginsHoverSwitchesPopup')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuPersistentViewToFilesHoverHighlightFollowsPointer', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'ResolveMenuBarHoverSuppressionPoint\(mainWindow,\s*viewScreenPoint,\s*viewMapping->visualIndex,\s*pluginsMapping->visualIndex\)'
        $testBlock | Should Not Match ':\s*std::pair<POINT,\s*int>\{viewScreenPoint,\s*static_cast<int>\(viewMapping->visualIndex\)\}'

        $cursorParkIndex = $testBlock.IndexOf('SetCursorPos(viewScreenPoint.x, viewScreenPoint.y)')
        $resolveIndex = $testBlock.IndexOf('ResolveMenuBarHoverSuppressionPoint(mainWindow, viewScreenPoint, viewMapping->visualIndex, pluginsMapping->visualIndex)')
        $cursorParkIndex | Should BeGreaterThan -1
        $resolveIndex | Should BeGreaterThan -1
        $cursorParkIndex | Should BeLessThan $resolveIndex
        $testBlock | Should Match 'DirectedSelfTestInputWarning[^\r\n]*inputWarning'
        $testBlock | Should Match 'restoreCursorOnExit\s*=\s*wil::scope_exit'
    }

    It 'routes persistent View-to-Files menu hover through the menu bar root-switch path' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuPersistentViewToFilesHoverHighlightFollowsPointer')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuMouseOpenKeepsPopupSelectionClear', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'ResolveMenuBarHoverSuppressionPoint\(mainWindow,\s*viewScreenPoint,\s*viewMapping->visualIndex,\s*filesMapping->visualIndex\)'
        $testBlock | Should Match 'DebugHitTestMainMenuBarScreenPoint\(mainWindow,\s*filesScreenPoint'
        $testBlock | Should Match 'settledOnFiles\s*=\s*replacementPopup\s*&&\s*selectedIndexAfterHover\.load'
        $testBlock | Should Not Match 'PostMessageW\(initialPopup,\s*WM_MOUSEMOVE'
        $testBlock | Should Match 'DirectedSelfTestInputWarning[^\r\n]*inputWarning'
        $testBlock | Should Match 'restoreCursorOnExit\s*=\s*wil::scope_exit'
    }

    It 'centralizes the Commands real-input warning and guards every cursor-warp family' {
        $commands = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'
        $view = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $search = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'
        $preferences = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp'

        $commands | Should Match 'class\s+DirectedSelfTestInputWarning'
        $view | Should Not Match 'class\s+DirectedSelfTestInputWarning'
        $search | Should Not Match 'class\s+SearchDirectedSelfTestInputWarning'
        ($view + $search + $preferences) | Should Not Match '\bSendInput\s*\('

        $preferencesStart = $preferences.IndexOf('[[nodiscard]] bool TestPreferencesDialogCategorySwitchesDoNotChurnTreeHost')
        $preferencesNext = $preferences.IndexOf('[[nodiscard]] bool TestPreferencesDialogScrollHostPreservesRetainedPageState', $preferencesStart)
        $preferencesStart | Should BeGreaterThan -1
        $preferencesNext | Should BeGreaterThan $preferencesStart
        $preferencesBlock = $preferences.Substring($preferencesStart, $preferencesNext - $preferencesStart)
        $preferencesBlock | Should Match 'DirectedSelfTestInputWarning\s+inputWarning'
        $preferencesBlock | Should Match 'restoreCursor\s*=\s*wil::scope_exit'

        $splitStart = $search.IndexOf('[[nodiscard]] bool ProbeFindSplitMenuStationaryHover')
        $splitNext = $search.IndexOf('[[nodiscard]] bool PostSelfTestOutsideClickToPopup', $splitStart)
        $splitStart | Should BeGreaterThan -1
        $splitNext | Should BeGreaterThan $splitStart
        $splitBlock = $search.Substring($splitStart, $splitNext - $splitStart)
        $splitBlock | Should Match 'DirectedSelfTestInputWarning\s+inputWarning'
        $splitBlock | Should Match 'restoreCursor\s*=\s*wil::scope_exit'

        $historyStart = $search.IndexOf('[[nodiscard]] bool ProbeFindDestinationHistoryMenu')
        $historyNext = $search.IndexOf('[[nodiscard]] bool ProbeFindDestinationHistoryMenuFromActiveEditMode', $historyStart)
        $historyStart | Should BeGreaterThan -1
        $historyNext | Should BeGreaterThan $historyStart
        $historyBlock = $search.Substring($historyStart, $historyNext - $historyStart)
        $historyBlock | Should Match 'DirectedSelfTestInputWarning\s+inputWarning'
        $historyBlock | Should Match 'restoreCursor\s*=\s*wil::scope_exit'
    }

    It 'opens the temporary menu-bar hover-switch popup through keyboard activation' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuHoverSwitchesTopLevelPopup')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuPersistentDirectHoverSwitchesTopLevelPopup', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'openInputSent'
        $testBlock | Should Match 'stalePointerSeedDelivered'
        $testBlock | Should Match 'SetFocus\(menuBarWindow\)'
        $testBlock | Should Match 'PostMessageW\(menuBarWindow,\s*WM_KEYDOWN,\s*VK_DOWN'
        $testBlock | Should Match 'stale pointer seed delivered'
        $testBlock | Should Match 'keyboard-open stale-pointer guard'
        $testBlock | Should Match 'Temporary menu-bar keyboard open did not open the initial DxUI popup'
        $testBlock | Should Not Match 'PostMessageW\(menuBarWindow,\s*WM_LBUTTONDOWN'
        $testBlock | Should Not Match 'Menu-bar click did not select the expected first top-level menu'
    }

    It 'parks real cursor away from DxUi popup initial hover in menu selftests' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $source | Should Match '\[\[nodiscard\]\]\s+bool\s+TryChooseMainWindowNeutralCursorPoint'

        $mouseStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuMouseOpenKeepsPopupSelectionClear')
        $mouseNext = $source.IndexOf('[[nodiscard]] bool TestMainMenuMouseOpenedPopupProcessesKeyboardBeforeMouseMove', $mouseStart)
        $mouseStart | Should BeGreaterThan -1
        $mouseNext | Should BeGreaterThan $mouseStart

        $mouseBlock = $source.Substring($mouseStart, $mouseNext - $mouseStart)
        $mouseBlock | Should Match 'TryChooseMainWindowNeutralCursorPoint\(mainWindow,\s*neutralCursorPoint\)'
        $mouseBlock | Should Match 'SetCursorPos\(neutralCursorPoint\.x,\s*neutralCursorPoint\.y\)'
        $mouseBlock | Should Match 'DirectedSelfTestInputWarning[^\r\n]*inputWarning'
        $mouseBlock | Should Match 'restoreCursor\s*=\s*wil::scope_exit'
        $mouseBlock | Should Match 'DebugSetMainMenuBarHoverSuppressionCursorOverride\(clickScreenPoint\)'
        $mouseBlock | Should Not Match 'SetCursorPos\(clickScreenPoint\.x,\s*clickScreenPoint\.y\)'

        $submenuStart = $source.IndexOf('[[nodiscard]] bool TestMainMenuSubmenuPlacementMatchesSpec')
        $submenuNext = $source.IndexOf('[[nodiscard]] bool TestAppPreferencesKeepsNavigationShellStable', $submenuStart)
        $submenuStart | Should BeGreaterThan -1
        $submenuNext | Should BeGreaterThan $submenuStart

        $submenuBlock = $source.Substring($submenuStart, $submenuNext - $submenuStart)
        $submenuBlock | Should Match 'TryChooseMainWindowNeutralCursorPoint\(mainWindow,\s*neutralCursorPoint\)'
        $submenuBlock | Should Match 'SetCursorPos\(neutralCursorPoint\.x,\s*neutralCursorPoint\.y\)'
        $submenuBlock | Should Match 'DirectedSelfTestInputWarning[^\r\n]*inputWarning'
        $submenuBlock | Should Match 'restoreCursor\s*=\s*wil::scope_exit'
        $submenuBlock | Should Match 'DebugSetMainMenuBarHoverSuppressionCursorOverride\(neutralCursorPoint\)'
    }

    It 'restores broad command-probe UI state and isolates modeless ownership fixtures' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestDispatchAllCommandsSmoke')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestModelessWindowOwnership', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'baselineZoomedPane\s*=\s*g_folderWindow\.GetZoomedPane\(\)'
        $testBlock | Should Match 'baselineZoomRestoreSplitRatio\s*=\s*g_folderWindow\.GetZoomRestoreSplitRatio\(\)'
        $testBlock | Should Match 'baselineSplitRatio\s*=\s*g_folderWindow\.GetSplitRatio\(\)'
        $testBlock | Should Match 'g_folderWindow\.SetZoomState\(baselineZoomedPane,\s*baselineZoomRestoreSplitRatio\)'
        $testBlock | Should Match 'if\s*\(\s*!\s*baselineZoomedPane\.has_value\(\)\s*&&\s*!\s*baselineZoomRestoreSplitRatio\.has_value\(\)\s*\)[\s\S]{0,180}g_folderWindow\.SetSplitRatio\(baselineSplitRatio\)'
        $testBlock | Should Match 'baselineFunctionBarVisible\s*=\s*g_folderWindow\.GetFunctionBarVisible\(\)'
        $testBlock | Should Match 'baselinePreviewOk\s*=\s*g_folderWindow\.DebugGetPreviewPaneSnapshot\(baselinePreview\)'
        $testBlock | Should Match 'g_folderWindow\.TogglePreviewPane\(currentPreview\.sourcePane\)'
        $testBlock | Should Match 'g_folderWindow\.SetFunctionBarVisible\(baselineFunctionBarVisible\)'
        $testBlock | Should Match 'g_folderWindow\.GetFunctionBarVisible\(\)\s*==\s*baselineFunctionBarVisible'

        $modelessStart = $nextTestStart
        $modelessNext = $source.IndexOf('[[nodiscard]] bool TestFullScreenToggle', $modelessStart)
        $modelessNext | Should BeGreaterThan $modelessStart

        $modelessBlock = $source.Substring($modelessStart, $modelessNext - $modelessStart)
        $modelessBlock | Should Match 'shortcutsBefore\s*=\s*g_settings\.shortcuts'
        $modelessBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow,\s*state,\s*L"modeless window ownership validation"\)'
        $modelessBlock | Should Match 'g_settings\.shortcuts\s*=\s*ShortcutDefaults::CreateDefaultShortcuts\(\)'
        $modelessBlock | Should Match 'g_settings\.shortcuts\s*=\s*shortcutsBefore'
        $modelessBlock | Should Match 'DebugReloadShortcutsFromSettings\(\)'
    }

    It 'keeps Shortcuts live UIA search interaction message-pumped under broad order' {
        $shortcutsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Shortcuts.cpp'
        $settingsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'

        $testStart = $shortcutsSource.IndexOf('[[nodiscard]] bool TestShortcutsWindowLiveDxSearchInteraction')
        $testNext = $shortcutsSource.IndexOf('[[nodiscard]] bool TestShortcutsWindowLongRunScrollingStaysBounded', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $shortcutsSource.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'CollectVisibleDescendantValuePatternStateWithMessagePump'
        $testBlock | Should Match 'CollectVisibleDescendantSelectionPatternStateWithMessagePump'
        $testBlock | Should Match 'SetVisibleDescendantValueWithMessagePump'
        $testBlock | Should Match 'Shortcuts \{\} search SetValue'
        $testBlock | Should Match 'Shortcuts reopened live search clear SetValue'
        $testBlock | Should Not Match 'SetVisibleDescendantValue\(shortcuts,\s*UIA_EditControlTypeId'
        $testBlock | Should Not Match 'CollectVisibleDescendantValuePatternState\(shortcuts,\s*UIA_EditControlTypeId'
        $testBlock | Should Not Match 'CollectVisibleDescendantSelectionPatternState\(shortcuts,\s*UIA_DataGridControlTypeId'

        $settingsSource | Should Match 'CollectVisibleDescendantSelectionPatternStateWithMessagePump'
        $settingsSource | Should Match 'SetVisibleDescendantValueWithMessagePump'
        $settingsSource | Should Match 'WaitForBoundedUiaWorker\(worker,\s*sharedState->lifetime,\s*3000u,\s*L"SelectionPattern read"'
        $settingsSource | Should Match 'UIA helper: \{\} timed out during ''\{\}''\.'
    }

    It 'keeps Compare Options close-reopen churn globally settled before live UIA reuse' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp'

        $source | Should Match 'WaitForNoCompareDirectoriesWindowForOptionsSelfTest'
        $source | Should Match 'CloseCompareDirectoriesWindowsForOptionsSelfTest'
        $source | Should Match 'SendMessageW\(existing,\s*WM_CLOSE'
        $source | Should Match 'WaitForNamedCompareOptionsEditValueState'
        $source | Should Match 'CollectWindowHostRawProviderValuePatternStates\(hwnd,\s*UIA_EditControlTypeId\)'

        $longRunStart = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsLongRunOpenCloseStaysStable')
        $liveStart = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsLiveDxBodyInteraction', $longRunStart)
        $nextLive = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsPointerClickTogglesLiveDxInteraction', $liveStart)
        $longRunStart | Should BeGreaterThan -1
        $liveStart | Should BeGreaterThan $longRunStart
        $nextLive | Should BeGreaterThan $liveStart

        $longRunBlock = $source.Substring($longRunStart, $liveStart - $longRunStart)
        $liveBlock = $source.Substring($liveStart, $nextLive - $liveStart)

        $longRunBlock | Should Match 'closeCompareWindow\(L"long-run open/close setup"\)'
        $longRunBlock | Should Match 'closeCompareWindow\(std::format\(L"long-run open/close cycle \{\} setup"'
        $longRunBlock | Should Match 'closeCompareWindow\(std::format\(L"long-run open/close cycle \{\} cleanup"'
        $longRunBlock | Should Match 'WaitForNoCompareDirectoriesWindowForOptionsSelfTest\(SelfTest::Scale\(1000ms\)\)'
        $longRunBlock | Should Not Match 'PostMessageW\(compare,\s*WM_CLOSE'

        $liveBlock | Should Match 'closeCompareWindow\(L"live DX body interaction setup"\)'
        $liveBlock | Should Match 'WaitForNoCompareDirectoriesWindowForOptionsSelfTest\(SelfTest::Scale\(1000ms\)\)'
        $liveBlock | Should Match 'Compare Directories window did not settle closed after live UIA InvokePattern Cancel action'

        $themeStart = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsThemeCycleKeepsSurfaceLegible')
        $themeNext = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsPointerClickTogglesLiveDxInteraction', $themeStart)
        $themeStart | Should BeGreaterThan -1
        $themeNext | Should BeGreaterThan $themeStart
        $themeBlock = $source.Substring($themeStart, $themeNext - $themeStart)
        $themeBlock | Should Match 'WaitForNamedCompareOptionsEditValueState\(compare,\s*ignoreFilesEditName'
        $themeBlock | Should Not Match 'CollectVisibleDescendantValuePatternStateWithMessagePump\(compare,\s*UIA_EditControlTypeId'
    }

    It 'keeps File Operations speed-limit prompt churn ValuePattern checks condition-based' {
        $fileOpsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.FileOps.cpp'
        $settingsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'

        $testStart = $fileOpsSource.IndexOf('[[nodiscard]] bool TestFileOperationsSpeedLimitPromptLongRunOpenCloseStaysStable')
        $testNext = $fileOpsSource.IndexOf('[[nodiscard]] bool TestFileOperationsSpeedLimitPromptKeepsNavigationShellStable', $testStart)
        $testStart | Should BeGreaterThan -1
        $testNext | Should BeGreaterThan $testStart

        $testBlock = $fileOpsSource.Substring($testStart, $testNext - $testStart)
        $testBlock | Should Match 'WaitForVisibleDescendantValuePatternState'
        $testBlock | Should Match 'File Operations speed-limit churn cycle \{\} initial ValuePattern read'
        $testBlock | Should Match 'Custom speed-limit prompt ValuePattern should settle to'
        $testBlock | Should Not Match 'const auto valueState\s*=\s*CollectVisibleDescendantValuePatternState\(prompt,\s*UIA_EditControlTypeId\)'

        $settingsSource | Should Match 'WaitForVisibleDescendantValuePatternState'
        $settingsSource | Should Match 'CollectVisibleDescendantValuePatternStateWithMessagePump'
    }

    It 'keeps dialog prompt model and UIA checks condition-based under broad order' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp'

        $changeStart = $source.IndexOf('[[nodiscard]] bool TestPaneChangeCasePromptLiveDxInteraction')
        $changeNext = $source.IndexOf('[[nodiscard]] bool TestPaneChangeCasePromptLongRunOpenCloseStaysStable', $changeStart)
        $createStart = $source.IndexOf('[[nodiscard]] bool TestCreateDirectoryPromptLongRunOpenCloseStaysStable')
        $createNext = $source.IndexOf('[[nodiscard]] bool TestPaneItemPropertiesUsesDxUiSurface', $createStart)
        $filterStart = $source.IndexOf('[[nodiscard]] bool TestPaneFilterPromptUsesDxUiSurface')
        $filterNext = $source.IndexOf('[[nodiscard]] bool TestPaneFilterPromptLiveDxInteraction', $filterStart)
        $changeStart | Should BeGreaterThan -1
        $changeNext | Should BeGreaterThan $changeStart
        $createStart | Should BeGreaterThan -1
        $createNext | Should BeGreaterThan $createStart
        $filterStart | Should BeGreaterThan -1
        $filterNext | Should BeGreaterThan $filterStart

        $changeBlock = $source.Substring($changeStart, $changeNext - $changeStart)
        $createBlock = $source.Substring($createStart, $createNext - $createStart)
        $filterBlock = $source.Substring($filterStart, $filterNext - $filterStart)

        $changeBlock | Should Match 'snapshotDeadline'
        $changeBlock | Should Match 'editedSnapshot\.includeSubdirsChecked\s*==\s*\(expectedState\s*==\s*ToggleState_On\)'
        $changeBlock | Should Match 'PumpPendingMessages\(\)'

        $createBlock | Should Match 'WaitForVisibleDescendantValuePatternState'
        $source | Should Match 'WaitForCreateDirectoryPromptSelectedTextSnapshot'
        $createBlock | Should Match 'WaitForCreateDirectoryPromptSelectedTextSnapshot\('
        $createBlock | Should Match 'Create-directory prompt cycle \{\} initial ValuePattern read'
        $createBlock | Should Match 'Create-directory prompt ValuePattern should settle to'
        $createBlock | Should Not Match 'cycleResult\.valueState\s*=\s*CollectVisibleDescendantValuePatternState\(prompt,\s*UIA_EditControlTypeId\)'

        $filterBlock | Should Match 'selectionMasksBefore\s*=\s*g_settings\.selectionMasks'
        $filterBlock | Should Match 'g_settings\.selectionMasks\s*=\s*selectionMasksBefore'
        $filterBlock | Should Match 'selectionMasks\.filterHistory\s*=\s*\{L"\*\.txt",\s*L"\*\.log"\}'
    }

    It 'keeps Find destination/action button probes pinned to deterministic input state' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'

        $historyStart = $source.IndexOf('[[nodiscard]] bool ProbeFindDestinationHistoryMenu')
        $historyNext = $source.IndexOf('[[nodiscard]] bool ProbeFindDestinationHistoryMenuFromActiveEditMode', $historyStart)
        $historyStart | Should BeGreaterThan -1
        $historyNext | Should BeGreaterThan $historyStart

        $historyBlock = $source.Substring($historyStart, $historyNext - $historyStart)
        $historyBlock | Should Match 'std::atomic<bool>\s+cursorMovedToItem\{false\}'
        $historyBlock | Should Match 'SetCursorPos\(itemCenter\.x,\s*itemCenter\.y\)'
        $historyBlock | Should Match 'PostMessageW\(popup,\s*WM_MOUSEMOVE,\s*0,\s*deliveredMove\)'
        $historyBlock | Should Match 'Find destination history menu did not repaint hover highlight for the item under the delivered pointer'
        $historyBlock | Should Match 'hoverFailureDetails'

        $splitStart = $source.IndexOf('[[nodiscard]] bool ProbeFindSplitMenuStationaryHover')
        $splitNext = $source.IndexOf('// Light-dismisses an owned DxUi context-menu popup', $splitStart)
        $splitStart | Should BeGreaterThan -1
        $splitNext | Should BeGreaterThan $splitStart

        $splitBlock = $source.Substring($splitStart, $splitNext - $splitStart)
        $splitBlock | Should Match 'WaitForNoVisibleOwnedDxUiContextMenusForSearchTest\(ownerWindow,\s*SelfTest::Scale\(1500ms\)\)'
        $splitBlock | Should Match 'std::atomic<bool>\s+cursorMovedToItem\{false\}'
        $splitBlock | Should Match 'PostMessageW\(popup,\s*WM_MOUSEMOVE,\s*0,\s*deliveredMove\)'
        $splitBlock | Should Match 'Find split action menu did not repaint hover highlight for the delivered stationary pointer'
        $splitBlock | Should Match 'hoverFailureDetails'

        $shortcutsStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogResultShortcutsUseShellClipboardAndFileActions')
        $shortcutsNext = $source.IndexOf('[[nodiscard]] bool TestFindDialogLargeLocalSearchUsesIncrementalUpdates', $shortcutsStart)
        $shortcutsStart | Should BeGreaterThan -1
        $shortcutsNext | Should BeGreaterThan $shortcutsStart

        $shortcutsBlock = $source.Substring($shortcutsStart, $shortcutsNext - $shortcutsStart)
        $shortcutsBlock | Should Match 'findSnapshotContainsPath'
        $shortcutsBlock | Should Match 'selectFindShortcutResult'
        $shortcutsBlock | Should Match 'Find shortcut result ''\{\}'' was not available before \{\}'
        $shortcutsBlock | Should Match 'Find shortcut result ''\{\}'' did not become the stable single selection before \{\}'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(file\.native\(\)\)'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(explicitCopyFile\.native\(\)\)'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(copyFile\.native\(\)\)'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(moveFile\.native\(\)\)'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(deleteFile\.native\(\)\)'
        $shortcutsBlock | Should Not Match 'DebugSelectFindFilesWindowResult\(permanentFile\.native\(\)\)'

        $actionStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogActionButtonsActivateExpectedCommands')
        $actionNext = $source.IndexOf('[[nodiscard]] bool TestFindDialogRestoresPersistedGridLayout', $actionStart)
        $actionStart | Should BeGreaterThan -1
        $actionNext | Should BeGreaterThan $actionStart

        $actionBlock = $source.Substring($actionStart, $actionNext - $actionStart)
        $actionBlock | Should Match 'PrepareMainWindowForIsolatedUiCase\(mainWindow, state, L"Find action-button command activation validation"\)'
        $openTrace = $actionBlock.IndexOf('Trace(L"action-buttons: invoking Open button")')
        $parentTrace = $actionBlock.IndexOf('Trace(L"action-buttons: invoking Go to folder button")')
        $openActivePane = $actionBlock.IndexOf('g_folderWindow.SetActivePane(FolderWindow::Pane::Left)', $actionBlock.IndexOf('Selecting the directory result did not enable Open/Parent'))
        $parentActivePane = $actionBlock.IndexOf('g_folderWindow.SetActivePane(FolderWindow::Pane::Left)', $openTrace + 1)
        $openActivePane | Should BeGreaterThan -1
        $parentActivePane | Should BeGreaterThan -1
        $openActivePane | Should BeLessThan $openTrace
        $parentActivePane | Should BeLessThan $parentTrace
        $actionBlock | Should Match 'Open button delivered click did not navigate the focused pane into the selected directory; focusedPane='
        $actionBlock | Should Match 'Parent button delivered click did not return the focused pane to the selected directory parent; focusedPane='
    }

    It 'keeps Find header-drag reorder failures diagnostic-rich' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogHeaderDragReordersColumnsWithoutSort')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogCopyFollowsReorderedColumns', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'const bool reordered = WaitForFindSnapshot'
        $testBlock | Should Match 'value\.usesDxUiHost'
        $testBlock | Should Match 'Find header drag should reorder visible columns without losing selection or bounded visible work\. \{\}'
        $testBlock | Should Match 'DescribeFindSnapshotBrief\(snapshot\)'
        $testBlock | Should Not Match 'state\.Require\(WaitForFindSnapshot\([\s\S]{0,900}Find header drag should reorder visible columns without losing selection or bounded visible work\."\);'
    }

    It 'keeps Find reordered-sort restore checks diagnostic-rich under broad order' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Search.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogRestoresReorderedSortedGridLayout')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestFindDialogRestoresCombinedViewState', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'Common::Settings::SearchDialogSettings\s+search\{\}'
        $testBlock | Should Match 'search\.includeFiles\s*=\s*true'
        $testBlock | Should Match 'search\.includeDirectories\s*=\s*false'
        $testBlock | Should Match 'search\.preferIndex\s*=\s*false'
        $testBlock | Should Match 'ApplyFindVisibleHeaderReorderViaDebug\(snapshot,\s*SelfTest::Scale\(3000ms\)\)'
        $testBlock | Should Match 'DebugSetFindFilesWindowResultSort\(0u,\s*true\)'
        $testBlock | Should Match 'SelfTest::Scale\(5000ms\)'
        $testBlock | Should Match 'Find logical Name sort did not stay correct after visible header reorder\. \{\}'
        $testBlock | Should Match 'DescribeFindSnapshotBrief\(snapshot\)'
    }

    It 'keeps Compare Directories lower-card scroll focus failures diagnostic-rich' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.CompareOptions.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsScrollToLowerCardsStaysStable')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestCompareDirectoriesOptionsEnterAndEscapeRouteDefaultCancel', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'scrollableResizeCandidates'
        $testBlock | Should Match 'resizeAttempts'
        $testBlock | Should Match 'bodyScrollMax > 0'
        $testBlock | Should Match 'dipsToPixels\(760\),\s*dipsToPixels\(480\)'
        $testBlock | Should Match 'Compare Directories options lower toggle did not stay in view after focus moved into the scrolled body'
        $testBlock | Should Match 'DescribeCompareOptionsThemeSnapshot\(snapshot\)'
        $testBlock | Should Match 'scrolledOffset'
        $testBlock | Should Match 'focusTarget'
        $testBlock | Should Match 'bodyScroll'
        $testBlock | Should Not Match 'reducedHeightPx'
    }

    It 'keeps Connection Manager long-run open-close settled and diagnostic-rich' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Connections.cpp'

        $testStart = $source.IndexOf('[[nodiscard]] bool TestConnectionManagerWindowLongRunOpenCloseStaysStable')
        $nextTestStart = $source.IndexOf('[[nodiscard]] bool TestConnectionManagerWindowThemeCycleKeepsFormAndSelectionLegible', $testStart)
        $testStart | Should BeGreaterThan -1
        $nextTestStart | Should BeGreaterThan $testStart

        $testBlock = $source.Substring($testStart, $nextTestStart - $testStart)
        $testBlock | Should Match 'closeExistingWindow\(std::format\(L"cycle \{\} open", cycle\)\)'
        $testBlock | Should Match 'DebugDispatchShortcutCommand\(mainWindow,\s*L"cmd/pane/connections"\)'
        $testBlock | Should Match 'Connection Manager window did not open during cycle \{\}\. commandDispatched=\{\} currentHandle=0x\{:X\}'
        $testBlock | Should Match 'DescribeConnectionManagerSnapshot\(diagnostic\)'
        $testBlock | Should Match 'Connection Manager singleton still has a live window after cycle \{\} close\.'
        $testBlock | Should Not Match 'SendMessageW\(mainWindow,\s*WM_COMMAND,\s*MAKEWPARAM\(IDM_PANE_CONNECTION_MANAGER,\s*0\),\s*0\)'
    }

    It 'keeps legacy TestSandbox cleanup dry-run gated and ShouldProcess protected' {
        $source = Get-RSText -Path 'Tools\Clean-TestSandbox.ps1'

        $source | Should Match 'SupportsShouldProcess'
        $source | Should Match '\[switch\]\$Apply'
        $source | Should Match 'Get-RSTestSandboxLegacyCleanupPlan'
        $source | Should Match 'Resolve-RSTestSandboxCleanupTargets'
        $source | Should Match 'if\s*\(-not\s+\$Apply\)'
        $source | Should Match '\$PSCmdlet\.ShouldProcess'
        $source | Should Match 'Remove-Item\s+-LiteralPath'
        $source | Should Match '-ErrorAction\s+Stop'
        $source | Should Match 'catch\s*\{'
        $source | Should Match 'Write-Warning'
        $source | Should Match "Status.*'Failed'"
        $source | Should Not Match 'Remove-Item\s+-Path'
    }

    It 'runs legacy TestSandbox cleanup from the unified runner before child tests' {
        $runner = Get-RSText -Path 'Tools\Run-AllTests.ps1'

        $runner | Should Match '\[switch\]\$SkipLegacySandboxCleanup'
        $runner | Should Match 'Clean-TestSandbox\.ps1'
        $runner | Should Match '-Apply'
        $runner | Should Match 'if\s*\(-not\s+\$SkipLegacySandboxCleanup\)'
        $runner.IndexOf('Clean-TestSandbox.ps1') | Should BeLessThan $runner.IndexOf('# --- Build step ---')
    }

    It 'bypasses fatal error modal dialogs while self-tests are running' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+ShowFatalErrorDialog\([\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true

        $body = $bodyMatch.Value
        $body | Should Match 'IsRunningAnySelfTest\(\)'
        $body | Should Match 'SelfTest::AppendSelfTestTrace'
        $body | Should Match 'return;'
        $body.IndexOf('IsRunningAnySelfTest()') | Should BeLessThan $body.IndexOf('dialog.ShowModal()')
    }

    It 'keeps Commands Preferences chunks namespace-self-contained' {
        $coordinator = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Preferences.cpp'
        $coordinator.TrimEnd() | Should Not Match 'namespace\s*\{\s*$'

        $chunkNames = @(
            'Commands.SelfTest.Preferences.ChromeAndPlugins.cpp',
            'Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp',
            'Commands.SelfTest.Preferences.HotPathsAndKeyboard.cpp',
            'Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp',
            'Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp',
            'Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp',
            'Commands.SelfTest.Preferences.Dispatch.cpp'
        )

        foreach ($chunkName in $chunkNames) {
            $chunk = Get-RSText -Path "RedSalamander\SelfTest\Commands\$chunkName"
            $chunk | Should Match '^namespace\s*\{'
            $chunk.TrimEnd() | Should Match '\}\s*//\s*namespace'
        }
    }

    It 'keeps IconCache association LRU perf emission outside the association cache lock helper' {
        $source = Get-RSText -Path 'RedSalamander\IconCache.cpp'
        $bodyMatch = [regex]::Match($source, 'EvictAssociationQueryBatch\(\)[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Not Match 'PerfEmit'
        $source | Should Match 'EmitAssociationLruEvictScanMetric'
    }

    It 'bounds DirectoryInfoCache eviction when protected entries exceed the cache budget' {
        $source = Get-RSText -Path 'RedSalamander\DirectoryInfoCache.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+DirectoryInfoCache::MaybeEvictLocked\([\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Match 'scanLimit\s*=\s*_lru\.size\(\)'
        $bodyMatch.Value | Should Match 'scanned\s*<\s*scanLimit'
        $bodyMatch.Value | Should Match '\+\+scanned'
        $bodyMatch.Value | Should Match 'only pinned/borrowed/loading entries cannot spin'
    }

    It 'keeps FolderView thumbnail WIC factory ownership local to the processing thread' {
        $header = Get-RSText -Path 'RedSalamander\FolderView.h'
        $source = Get-RSText -Path 'RedSalamander\FolderView.Icons.cpp'

        $header | Should Not Match '_thumbnailWicFactory'
        $source | Should Match 'EnsureThumbnailWicFactory\(\s*wil::com_ptr<IWICImagingFactory>&'
        $source | Should Match 'thumbnailWicFactory'
    }

    It 'keeps IMAP summary repair budget exhaustion local to the current repair pass' {
        $source = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.Imap.cpp'
        $bodyMatch = [regex]::Match($source, 'auto\s+repairMissingSummaries\s*=\s*\[[\s\S]*?return\s+repairResult;\s*\n\s*\};')
        $bodyMatch.Success | Should Be $true

        $bodyMatch.Value | Should Not Match 'kRepairBudgetExceededHr'
        $bodyMatch.Value | Should Not Match 'tryConsumeRepairFetch[\s\S]*?noexcept'
        $bodyMatch.Value | Should Match 'break;'
    }

    It 'clears FolderView incremental-search layout effects over the visual label text' {
        # Highlights are matched and applied against GetVisualDisplayName (see
        # the sibling contract below), so clearing must use the same text the
        # label layout and highlight ranges were built from.
        $source = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+FolderView::ClearIncrementalSearchLayoutEffects\(\)\s+noexcept[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true
        $bodyMatch.Value | Should Match 'GetVisualDisplayName\(item\)'
        $bodyMatch.Value | Should Not Match 'item\.displayName'
    }

    It 'matches FolderView incremental-search highlights against visual display names' {
        $source = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $bodyMatch = [regex]::Match($source, 'void\s+FolderView::UpdateIncrementalSearchHighlightForFocusedItem\(\)[\s\S]*?\n\}')
        $bodyMatch.Success | Should Be $true

        $bodyMatch.Value | Should Match 'GetVisualDisplayName\(item\)'
        $bodyMatch.Value | Should Not Match 'FindIncrementalSearchMatchOffset\(item\.displayName\)'
    }

    It 'cleans stale FolderView incremental-search drawing effects when search is inactive' {
        $header = Get-RSText -Path 'RedSalamander\FolderView.h'
        $rendering = Get-RSText -Path 'RedSalamander\FolderView.Rendering.cpp'
        $interaction = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'

        $header | Should Match '_incrementalSearchLayoutEffectsDirty'
        $rendering | Should Match '_incrementalSearchLayoutEffectsDirty'
        $rendering | Should Match 'ClearIncrementalSearchLayoutEffects\(\)'
        $interaction | Should Match '_incrementalSearchLayoutEffectsDirty\s*=\s*true'
        $interaction | Should Match '_incrementalSearchLayoutEffectsDirty\s*=\s*false'
    }

    It 'exposes runner-native self-test case listing without executing case bodies' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $common = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commandsHeader = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.h'
        $commandsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'
        $compareHeader = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.h'
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'

        $main | Should Match '--selftest-list-cases'
        $main | Should Match 'BuildSelfTestCaseListJson'
        $main | Should Match 'SelfTest::GetSelfTestOptions\(\)\s*=\s*g_selfTestOptions;'
        $main | Should Match 'FileOperationsSelfTest::BuildExpectedCaseNames'
        $common | Should Match 'listCasesOnly'
        $commandsHeader | Should Match 'ListCases'
        $commandsSource | Should Match 'listCasesOnly'
        $compareHeader | Should Match 'ListCases'
        $compareSource | Should Match 'kCompareCaseNames'
    }

    It 'exposes native self-test repeat and seeded shuffle execution controls' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $commonHeader = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commonSource = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $commandsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'

        $main | Should Match '--selftest-repeat=N'
        $main | Should Match '--selftest-shuffle=SEED'
        $main | Should Match 'getArgValue\(L"--selftest-repeat=",\s*repeatArg\)'
        $main | Should Match 'getArgValue\(L"--selftest-shuffle=",\s*shuffleArg\)'
        $main | Should Match 'ParseSelfTestRepeatCount'
        $main | Should Match 'ParseSelfTestShuffleSeed'
        $main | Should Match '--selftest-shuffle is currently supported for --commands-selftest, --compare-selftest, and --fileops-selftest'

        $commonHeader | Should Match 'repeatCount'
        $commonHeader | Should Match 'shuffleSeed'
        $commonHeader | Should Match 'BuildSelfTestCaseExecutionOrder'
        $commonHeader | Should Match 'ShouldUseExplicitCaseExecutionOrder'
        $commonHeader | Should Match 'RunCaseAttempt'

        $commonSource | Should Match '#include <random>'
        $commonSource | Should Match 'std::shuffle'
        $commonSource | Should Match 'std::mt19937_64'
        $commonSource | Should Match 'repeatIndex'
        $commonSource | Should Match 'shuffleSeed'

        $commandsSource | Should Match 'BuildSelfTestCaseExecutionOrder'
    }

    It 'reinitializes Commands UIA cache after explicit-order case listing' {
        $settingsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'

        $contextStart = $settingsSource.IndexOf('struct UiaThreadContext final')
        $contextNext = $settingsSource.IndexOf('[[nodiscard]] UiaThreadContext& GetThreadUiAutomationContext', $contextStart)
        $contextStart | Should BeGreaterThan -1
        $contextNext | Should BeGreaterThan $contextStart

        $contextBlock = $settingsSource.Substring($contextStart, $contextNext - $contextStart)
        $contextBlock | Should Match 'void EnsureInitialized\(\) noexcept'
        $contextBlock | Should Match 'if \(automation\)'
        $contextBlock | Should Match 'automation\s*=\s*std::move\(createdAutomation\)'

        $getterStart = $settingsSource.IndexOf('IUIAutomation* GetThreadUiAutomation()')
        $getterNext = $settingsSource.IndexOf('void ReleaseThreadUiAutomationForSelfTest', $getterStart)
        $getterStart | Should BeGreaterThan -1
        $getterNext | Should BeGreaterThan $getterStart

        $getterBlock = $settingsSource.Substring($getterStart, $getterNext - $getterStart)
        $getterBlock | Should Match 'context\.EnsureInitialized\(\)'
    }

    It 'isolates every Commands explicit-order case before fixture dispatch' {
        $source = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'
        $loopStart = $source.IndexOf('for (const SelfTest::SelfTestCaseExecution& execution : executionOrder)')
        $loopNext = $source.IndexOf('else', $loopStart)
        $loopStart | Should BeGreaterThan -1
        $loopNext | Should BeGreaterThan $loopStart

        $loopBlock = $source.Substring($loopStart, $loopNext - $loopStart)
        $prepareIndex = $loopBlock.IndexOf('PrepareMainWindowForIsolatedUiCase(mainWindow, isolationState, isolationContext)')
        $dispatchIndex = $loopBlock.LastIndexOf('runRegisteredCases(caseOptions)')
        $prepareIndex | Should BeGreaterThan -1
        $dispatchIndex | Should BeGreaterThan $prepareIndex
        $loopBlock | Should Match 'SelfTest::RunCase\(caseOptions,\s*suite,\s*execution\.name'
        $loopBlock | Should Match 'state\.Require\(false,\s*failure\)'
    }

    It 'keeps CompareDirectories seeded shuffle native through explicit case order' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $main | Should Match 'RunCompareDirectoriesSelfTestPlan'
        $main | Should Match 'CompareDirectoriesSelfTest::ListCases'
        $main | Should Match 'BuildSelfTestCaseExecutionOrder'
        $main | Should Match 'CompareSelfTest: explicit execution order'
        $main | Should Match 'caseOptions\.caseFilter\s*=\s*execution\.name'
        $main | Should Match 'caseOptions\.repeatCount\s*=\s*1u'
        $main | Should Match 'caseOptions\.repeatIndex\s*=\s*execution\.repeatIndex'
        $main | Should Match 'caseOptions\.writeJsonSummary\s*=\s*false'
        $main | Should Match 'RunCompareDirectoriesSelfTestPlan\(g_selfTestOptions,\s*&compareResult\)'
    }

    It 'keeps FileOperations repeat coverage native instead of runner-only' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $main | Should Match 'g_fileOpsSelfTestRunRepeatIndexes'
        $main | Should Match 'g_fileOpsSelfTestExpectedCases'
        $main | Should Match 'SelfTest::SelfTestCaseExecution'
        $main | Should Match 'BuildFileOpsRepeatedRunPlan'
        $main | Should Match 'BuildFileOpsExecutionPlan'
        $main | Should Match 'FileOpsSelfTest: explicit execution order'
        $main | Should Match 'IsFileOpsStructuralCaseName'
        $main | Should Match 'orderOptions\.caseFilter\.clear\(\)'
        $main | Should Match 'BuildFileOpsRepeatedExpectedCases'
        $main | Should Match 'MakeFileOpsRunOptions\(runFilter,\s*repeatIndex\)'
        $main | Should Match 'item\.repeatIndex\s*=\s*repeatIndex'
        $main | Should Match 'existing\.name == item\.name && existing\.repeatIndex == item\.repeatIndex'
        $main | Should Match 'skipped\.repeatIndex\s*=\s*expectedCase\.repeatIndex'
    }

    It 'keeps FolderView overlay perf sample sufficiency advisory' {
        $viewCommands = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $perfSpec = Get-RSText -Path 'Specs\Testing\Testing_PerformanceValidation.md'

        $caseStart = $viewCommands.IndexOf('[[nodiscard]] bool TestFolderViewPerfOverlayInvalidationStress')
        $caseEnd = $viewCommands.IndexOf('[[nodiscard]] bool TestFolderViewPerfScrollRenderStress')
        $caseStart | Should Not Be -1
        $caseEnd | Should BeGreaterThan $caseStart
        $caseBlock = $viewCommands.Substring($caseStart, $caseEnd - $caseStart)

        $caseBlock | Should Not Match 'Scale\s*\(\s*4200ms\s*\)'
        $caseBlock | Should Not Match 'state\.Require\s*\(\s*overlaySamplesEnoughForP95'
        $caseBlock | Should Match 'AppendFolderViewMetricQualityJson\s*\(\s*json\s*,\s*overlayTotalFrameCount,\s*overlayPresentFrameCount'
        $caseBlock | Should Match 'folderView_perf_overlay_invalidation_stress_metrics\.json'
        $caseBlock | Should Match 'overlayFrameCollectionPassCount'

        $perfSpec | Should Match 'folderView_perf_overlay_invalidation_stress'
        $perfSpec | Should Match 'samplesEnoughForP95.*advisory'
    }

    It 'uses shared self-test result emission for ad hoc and state-machine cases' {
        $commonHeader = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commonSource = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $fileOpsSource = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $commonHeader | Should Match 'AppendCaseResult'
        $commonSource | Should Match 'void\s+AppendCaseResult'
        $compareSource | Should Match 'SelfTest::AppendCaseResult'
        $fileOpsSource | Should Match 'SelfTest::AppendCaseResult'
    }

    It 'preserves partial self-test results when a case or suite crashes' {
        $commonHeader = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commonSource = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.cpp'
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $runner = Get-RSText -Path 'Tools\Run-AllTests.ps1'

        $commonHeader | Should Match 'BeginInFlightSelfTestCase'
        $commonHeader | Should Match 'EndInFlightSelfTestCase'
        $commonHeader | Should Match 'MarkInFlightSelfTestCaseCrashed'
        $commonHeader | Should Match 'Status\s*\{[\s\S]{0,120}crashed'
        $commonHeader | Should Match 'crashCaseName'
        $commonHeader | Should Match 'TriggerSelfTestCaseCrashInjection'

        $runCaseMatch = [regex]::Match($commonHeader, 'template\s*<typename\s+Func>\s*void\s+RunCase[\s\S]*?\n\}')
        $runCaseMatch.Success | Should Be $true
        $runCaseBody = $runCaseMatch.Value
        $runCaseBody | Should Match 'BeginInFlightSelfTestCase\s*\(\s*suite\.suite,\s*name\s*\)'
        $runCaseBody | Should Match 'SelfTestCaseNameEquals\s*\(\s*options\.crashCaseName,\s*name\s*\)'
        $runCaseBody | Should Match 'TriggerSelfTestCaseCrashInjection\s*\(\s*suite\.suite,\s*name\s*\)'
        $runCaseBody | Should Match 'EndInFlightSelfTestCase\s*\(\s*suite\.suite,\s*name\s*\)'

        $appendCaseMatch = [regex]::Match($commonSource, 'void\s+AppendCaseResult\s*\(\s*SelfTestSuiteResult&\s+suite,\s*SelfTestCaseResult\s+result\s*\)[\s\S]*?\n\}')
        $appendCaseMatch.Success | Should Be $true
        $appendCaseMatch.Value | Should Match 'FlushSuiteJsonAfterCase\s*\(\s*suite\s*\)'
        $commonSource | Should Match 'void\s+TriggerSelfTestCaseCrashInjection'
        $commonSource | Should Match 'EXCEPTION_ACCESS_VIOLATION'
        $commonSource | Should Match 'void\s+MarkInFlightSelfTestCaseCrashed'
        $commonSource | Should Match 'Status::crashed'

        $crashHelperStart = $main.IndexOf('void RecordSelfTestUnhandledExceptionCrash')
        $crashHelperEnd = $main.IndexOf('void TraceSelfTestExitCode', $crashHelperStart)
        $crashHelperStart | Should Not Be -1
        $crashHelperEnd | Should BeGreaterThan $crashHelperStart
        $crashHelperBody = $main.Substring($crashHelperStart, $crashHelperEnd - $crashHelperStart)
        $crashHelperBody | Should Match 'MarkInFlightSelfTestCaseCrashed\s*\(\s*g_selfTestRunResult'
        $crashHelperBody | Should Match 'WriteRunJson\s*\(\s*g_selfTestRunResult'

        $sehMatch = [regex]::Match($main, '__except\s*\(CrashHandler::WriteDumpForException\(GetExceptionInformation\(\)\)\)[\s\S]*?return\s+-1;')
        $sehMatch.Success | Should Be $true
        $sehBody = $sehMatch.Value
        $sehBody | Should Match 'RecordSelfTestUnhandledExceptionCrash\s*\(\s*exceptionCode,\s*exceptionName\s*\)'
        $sehBody.IndexOf('RecordSelfTestUnhandledExceptionCrash') | Should BeLessThan $sehBody.IndexOf('return -1;')

        $main | Should Match '--selftest-crash-case=NAME'
        $main | Should Match 'getArgValue\(L"--selftest-crash-case=",\s*crashCaseArg\)'
        $main | Should Match 'g_selfTestOptions\.crashCaseName\s*=\s*std::move\(crashCaseArg\)'
        $runner | Should Match "'crashed'\s*\{\s*return\s*'Red'"
        $runner | Should Match "-in\s+@\('failed',\s*'crashed'\)"
    }

    It 'keeps injected classifier proof hooks explicit and debug-only' {
        $main = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $commonHeader = Get-RSText -Path 'RedSalamander\SelfTest\Common\SelfTestCommon.h'
        $commandsSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.cpp'

        $main | Should Match '--selftest-flaky-proof-case=NAME'
        $main | Should Match '--selftest-order-proof-case=NAME'
        $main | Should Match 'getArgValue\(L"--selftest-flaky-proof-case=",\s*flakyProofCaseArg\)'
        $main | Should Match 'getArgValue\(L"--selftest-order-proof-case=",\s*orderProofCaseArg\)'
        $main | Should Match 'g_selfTestOptions\.flakyProofCaseName\s*=\s*std::move\(flakyProofCaseArg\)'
        $main | Should Match 'g_selfTestOptions\.orderProofCaseName\s*=\s*std::move\(orderProofCaseArg\)'

        $commonHeader | Should Match 'flakyProofCaseName'
        $commonHeader | Should Match 'orderProofCaseName'
        $commonHeader | Should Match 'TryInjectSelfTestClassifierProofFailure'
        $commonHeader | Should Match 'injected flaky classifier proof failure'
        $commonHeader | Should Match 'injected order-dependent classifier proof failure'
        $commonHeader | Should Match 'classifierProofSuiteContext'
        $commonHeader | Should Match 'classifierProofShuffleContext'
        $commonHeader | Should Match 'const bool isolatedCaseRerun\s*=\s*! options\.classifierProofSuiteContext'
        $commonHeader | Should Match '! options\.classifierProofShuffleContext'

        $commandsSource | Should Match 'caseOptions\.classifierProofSuiteContext\s*=\s*true'
        $commandsSource | Should Match 'caseOptions\.classifierProofShuffleContext\s*=\s*options\.shuffleSeed\.has_value\(\)'
        $main | Should Match 'caseOptions\.classifierProofSuiteContext\s*=\s*true'
        $main | Should Match 'caseOptions\.classifierProofShuffleContext\s*=\s*options\.shuffleSeed\.has_value\(\)'
        $main | Should Match 'options\.classifierProofSuiteContext\s*=\s*SelfTest::ShouldUseExplicitCaseExecutionOrder\(g_selfTestOptions\)'
        $main | Should Match 'options\.classifierProofShuffleContext\s*=\s*g_selfTestOptions\.shuffleSeed\.has_value\(\)'
    }

    It 'keeps CompareDirectories runner-listed case names unique' {
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $arrayMatch = [regex]::Match($compareSource, 'kCompareCaseNames[\s\S]*?=\s*\{(?<body>[\s\S]*?)\};')
        $arrayMatch.Success | Should Be $true

        $names = @([regex]::Matches($arrayMatch.Groups['body'].Value, 'L"(?<name>[^"]+)"') | ForEach-Object { $_.Groups['name'].Value })
        $duplicates = @($names | Group-Object | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name)

        @($duplicates).Count | Should Be 0
    }

    It 'keeps CompareDirectories literal RunCase names listed for runner coverage' {
        $compareSource = Get-RSText -Path 'RedSalamander\SelfTest\CompareDirectories\CompareDirectoriesEngine.SelfTest.cpp'
        $arrayMatch = [regex]::Match($compareSource, 'kCompareCaseNames[\s\S]*?=\s*\{(?<body>[\s\S]*?)\};')
        $arrayMatch.Success | Should Be $true

        $listedNames = @([regex]::Matches($arrayMatch.Groups['body'].Value, 'L"(?<name>[^"]+)"') | ForEach-Object { $_.Groups['name'].Value })
        $listedSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($name in $listedNames) {
            [void]$listedSet.Add($name)
        }

        $compareDir = Join-Path $repoRoot 'RedSalamander\SelfTest\CompareDirectories'
        $literalRunCaseNames = @(
            Get-ChildItem -LiteralPath $compareDir -Filter '*.cpp' |
                ForEach-Object {
                    $source = Get-Content -LiteralPath $_.FullName -Raw
                    [regex]::Matches($source, 'SelfTest::RunCase\(\s*options,\s*suite,\s*L"(?<name>[^"]+)"') |
                        ForEach-Object { $_.Groups['name'].Value }
                } |
                Sort-Object -Unique
        )

        $missing = @($literalRunCaseNames | Where-Object { -not $listedSet.Contains($_) })
        @($missing).Count | Should Be 0
    }

    It 'opens the file-operations custom speed-limit prompt asynchronously for self-tests' {
        $popupSource = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'

        $popupSource | Should Match 'kFileOperationsPopupDeferredSpeedLimitPromptMessage'
        $popupSource | Should Match 'Let self-test callers advance their state before the modal prompt loop starts'
        $popupSource | Should Match 'PostMessageW\(hwnd,\s*kFileOperationsPopupDeferredSpeedLimitPromptMessage'
        $popupSource | Should Not Match 'payload->kind\s*==\s*PopupHitTest::Kind::TaskSpeedLimit\s*&&\s*payload->data\s*==\s*1u\)[\s\S]{0,160}ShowCustomSpeedLimitPromptForTask'
    }

    It 'keeps FileOperations self-test case prefixes runnable' {
        $fileOpsSource = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $fileOpsSource | Should Match 'FindPhasesByPrefix'
        $fileOpsSource | Should Match 'StartsWithIgnoreCase\(StepToString\(step\),\s*prefix\)'
        $fileOpsSource | Should Match 'selection\.activePhases\s*=\s*prefixMatches'
        $fileOpsSource | Should Match 'ResolveRunSelection\(filter\)\.recognized'
    }

    It 'keeps FileOperations pre-calc perf counters race-free under worker concurrency' {
        $header = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperationsInternal.h'
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $header | Should Match 'std::atomic<uint64_t>\s+preCalcUs'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcCallbackCount'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcCallbackUs'
        $header | Should Match 'std::atomic<uint64_t>\s+preCalcLockWaitUs'

        $source | Should Match '_perf\.preCalcUs\.fetch_add'
        $source | Should Match '_perf\.preCalcCallbackCount\.fetch_add'
        $source | Should Match '_perf\.preCalcCallbackUs\.fetch_add'
        $source | Should Match '_perf\.preCalcLockWaitUs\.fetch_add'

        $source | Should Not Match '\+\+_perf\.preCalcCallbackCount'
        $source | Should Not Match '_perf\.preCalcCallbackCount\+\+'
        $source | Should Not Match '_perf\.preCalcUs\s*\+='
        $source | Should Not Match '_perf\.preCalcCallbackUs\s*\+='
        $source | Should Not Match '_perf\.preCalcLockWaitUs\s*\+='
    }

    It 'keeps FileSystem dynamic scheduler cleanup explicit and recursive-copy queueing unbounded' {
        $source = Get-RSText -Path 'Plugins\FileSystem\FileSystem.FileOps.cpp'

        $source | Should Not Match 'maxQueuedItems'
        $source | Should Not Match 'queue\.items\.size\(\)\s*<'
        $source | Should Not Match 'queue\.cv\.wait\(lock'

        $schedulerMatch = [regex]::Match($source, 'class\s+SharedFileOpsJobScheduler\s+final[\s\S]*?\n\};\s*\n\s*SharedFileOpsJobScheduler&')
        $schedulerMatch.Success | Should Be $true
        $scheduler = $schedulerMatch.Value

        $hasSchedulableMatch = [regex]::Match($scheduler, '\[\[nodiscard\]\]\s+bool\s+hasSchedulableWorkLocked\(\)\s+noexcept\s*\{(?<body>[\s\S]*?)\n\s*\}')
        $hasSchedulableMatch.Success | Should Be $true
        $hasSchedulableMatch.Groups['body'].Value | Should Not Match 'cleanupJobsLocked'

        $scheduler | Should Match 'while\s*\(!\s*job->done\.load[\s\S]*?cleanupJobsLocked\(\);[\s\S]*?_cv\.wait\(lock\);[\s\S]*?cleanupJobsLocked\(\);'
    }

    It 'keeps cross-filesystem MOVE source deletion behind one bridge-copy completion helper' {
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $source | Should Match 'ShouldDeleteMoveSourceAfterBridgeCopy'
        $source | Should Match 'SUCCEEDED\(copyHr\)[\s\S]{0,360}skippedFileConflictCount\s*==\s*0'
        $source | Should Not Match 'moveCopyCompleted\s*=\s*bridgeSkippedDirectoryReparseCount\s*==\s*0\s*&&\s*!\s*bridgeRootDirectoryReparseSkipped'
        $source | Should Not Match 'bridgeLeftSourceAuthoritative\s*=\s*bridge\.skippedDirectoryReparseCount\s*>\s*0'
    }

    It 'keeps FileOperations conflict routing in one arbiter and one action-layout builder' {
        $header = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperationsInternal.h'
        $source = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'

        $header | Should Match 'struct\s+ConflictArbiter'
        $header | Should Match 'ConflictArbiter\s+_conflictArbiter'

        ([regex]::Matches($source, '\[\[nodiscard\]\]\s+ConflictActionLayout\s+BuildConflictActionLayout')).Count | Should Be 1
        $source | Should Not Match 'clearCachedDecision'
        $source | Should Not Match 'const\s+auto\s+setConflictPromptLocked'
        $source | Should Not Match 'const\s+auto\s+waitForConflictDecision'
        $source | Should Not Match 'const\s+auto\s+getCachedDecision'
        $source | Should Not Match 'const\s+auto\s+setCachedDecision'
    }

    It 'keeps Riptide low-risk hardening invariants wired' {
        $fileOps = Get-RSText -Path 'Plugins\FileSystem\FileSystem.FileOps.cpp'
        $runtime = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.Runtime.cpp'
        $bridge = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'
        $storage = Get-RSText -Path 'Plugins\FileSystem\FileSystem.cpp'
        $popup = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        $microsoftDrive = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $findFiles = Get-RSText -Path 'RedSalamander\FindFilesWindow.cpp'
        $folderWindowHeader = Get-RSText -Path 'RedSalamander\FolderWindow.h'
        $folderWindow = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.cpp'

        $fileOps | Should Not Match 'IsDirectoryEmpty'
        $runtime | Should Match 'sourceCandidates'
        $runtime | Should Match 'destinationCandidates'
        $runtime | Should Match 'for\s*\(const std::wstring& sourceCandidate'
        $runtime | Should Match 'for\s*\(const std::wstring& destinationCandidate'
        $bridge | Should Match 'const\s+FileSystemFlags\s+tempFlags[\s\S]{0,260}~[\s\S]{0,160}FILESYSTEM_FLAG_ALLOW_OVERWRITE'
        $bridge | Should Not Match 'const\s+FileSystemFlags\s+tempFlags\s*=\s*static_cast<FileSystemFlags>\([\s\S]{0,220}\|\s*static_cast<uint32_t>\(FILESYSTEM_FLAG_ALLOW_OVERWRITE\)'
        $storage | Should Match 'else\s+if\s*\(\s*busTypeKnown\s*&&\s*busType\s*==\s*BusTypeNvme\s*\)'
        $popup | Should Match 'if\s*\(\s*weightCount\s*<\s*weights\.size\(\)\s*\)\s*\{[\s\S]{0,160}\+\+weightCount;'
        $popup | Should Not Match '\}\s*\+\+weightCount;\s*\}'

        $microsoftDrive | Should Match 'using\s+CancelProbe\s*=\s*std::function<HRESULT\(\)>'
        $microsoftDrive | Should Match 'MergeMoveFolderIntoExisting[\s\S]{0,800}const\s+CancelProbe&\s+checkCancel'
        $microsoftDrive | Should Match 'MoveOrRenameItem[\s\S]{0,800}const\s+CancelProbe&\s+checkCancel'
        $microsoftDrive | Should Match 'const\s+CancelProbe\s+checkCancel'
        $microsoftDrive | Should Match 'MoveOrRenameItem[\s\S]{0,520}checkCancel'

        $s3 | Should Match 'kMaxS3PerObjectConflictRetries'
        $s3 | Should Match 'unsigned\s+int\s+retryCount\s*=\s*0'
        $s3 | Should Match 'case\s+FileSystemIssueAction::Retry:[\s\S]{0,520}checkCancel[\s\S]{0,520}retryCount'

        $findFiles | Should Match 'uint64_t\s+taskId\s*=\s*0'
        $findFiles | Should Match 'uint64_t\s+createdTickMs\s*=\s*0'
        $findFiles | Should Match 'ReapExpiredPendingResultRemovals'
        $findFiles | Should Match 'PendingResultRemoval\{\.taskId\s*='
        $folderWindowHeader | Should Match 'std::weak_ptr<void>\s+lifetimeGuard'
        $folderWindowHeader | Should Match 'bool\s+hasLifetimeGuard\s*=\s*false'
        $folderWindow | Should Match 'hasLifetimeGuard\s*&&\s*subscription\.lifetimeGuard\.expired\(\)'
    }

    It 'keeps Floodgate data-safety invariants wired' {
        $bridge = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'
        $bridgeQueue = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.Queue.cpp'
        $sevenZip = Get-RSText -Path 'Plugins\FileSystem7z\FileSystem7z.cpp'
        $fileOpsSelfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'
        $fairstreamSelfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Fairstream.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $s3Io = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.IO.cpp'
        $curl = Get-RSText -Path 'Plugins\FileSystemCurl\FileSystemCurl.CopyMove.cpp'
        $microsoftDrive = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $directoryOps = Get-RSText -Path 'Plugins\FileSystem\FileSystem.DirectoryOps.cpp'
        $watch = Get-RSText -Path 'Plugins\FileSystem\FileSystem.Watch.cpp'
        $fileOpsDiagnostics = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.Diagnostics.cpp'
        $issuesPane = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.IssuesPane.cpp'
        $localIndex = Get-RSText -Path 'Common\LocalSearchIndexCore.cpp'
        $folderView = Get-RSText -Path 'RedSalamander\FolderView.cpp'
        $folderViewEnumeration = Get-RSText -Path 'RedSalamander\FolderView.Enumeration.cpp'
        $folderViewInteraction = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $navigationCommands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Navigation.cpp'
        $mainWindow = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'

        $s3 | Should Match 'FindAncestorObjectConflict'
        $s3 | Should Match 'plannedDestinationKeys'
        $s3 | Should Match 'planned sibling'
        $s3 | Should Match 'plannedDestinationKeys\.contains\(destinationStates\[i\]\.ancestorKey\)'
        $s3 | Should Match 'destinationStates\[i\]\.ancestorConflict\s*&&\s*!\s*destinationStates\[i\]\.ancestorKey\.empty\(\)\s*&&[\s\S]{0,160}plannedDestinationKeys\.contains\(destinationStates\[i\]\.ancestorKey\)[\s\S]{0,120}objectSkipped\s*=\s*true'
        $s3 | Should Match 'destinationStates\[i\]\.ancestorConflict[\s\S]{0,220}plannedDestinationKeys\.contains\(destinationStates\[i\]\.ancestorKey\)'
        $s3 | Should Match 'return\s+hadSkipped\s*\?\s*HRESULT_FROM_WIN32\(ERROR_PARTIAL_COPY\)\s*:\s*S_OK'
        $s3 | Should Match 'RunDebugPlannedDestinationAncestorCollisionSelfTest'
        $s3 | Should Match 'src/a/b'

        $bridge | Should Match 'hrReaderSize\s*=\s*reader->GetSize\(&fileTotalBytes\)'
        $bridge | Should Match 'const\s+bool\s+isMove\s*=\s*task\._operation\s*==\s*FILESYSTEM_MOVE[\s\S]{0,260}bridge\.integrity\.sourceSizeUnknown'
        $bridge | Should Match 'PromoteTempToFinalPath[\s\S]{0,2200}CreateFileReader\(destinationPath\.c_str\(\),\s*destinationReader\.addressof\(\)\)'
        $bridge | Should Match 'bridge\.integrity\.destinationSizeMismatch'
        $bridgeQueue | Should Match 'SetFileOpsBridgeFailNextSourceGetSizeForSelfTest'
        $fairstreamSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource[\s\S]{0,1800}SetFileOpsBridgeFailNextSourceGetSizeForSelfTest\(1\)'
        $fairstreamSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource'
        $fairstreamSelfTest | Should Match 'ERROR_PARTIAL_COPY'
        $fairstreamSelfTest | Should Match 'sourceText\s*!=\s*kSourceContents'
        $fileOpsSelfTest | Should Match 'Floodgate_CrossFsMoveGetSizeFailurePreservesSource'
        # The full per-family suite only runs cases that are MEMBERS of a kFileOpsFamily* array; a case present
        # only in the enum/name-map/kFileOpsPhaseOrder is reported but silently skipped. Assert the data-safety
        # case is actually inside the Fairstream family literal (between its '{{' and closing '}}').
        $fileOpsSelfTest | Should Match 'kFileOpsFamilyFairstream\{\{(?:(?!\}\})[\s\S])*?Floodgate_CrossFsMoveGetSizeFailurePreservesSource'

        $curl | Should Match 'CurlUploadFromFile\(destinationConn,\s*stagedRemotePath'
        $curl | Should Match 'CurlProbeRemoteFileSize\(destinationConn,\s*stagedRemotePath,\s*stagedProbeSize,\s*stagedProbeSizeKnown\)'
        $curl | Should Match 'stagedProbeSizeKnown\s*&&\s*stagedProbeSize\s*!=\s*fileSize[\s\S]{0,160}RemoteDeleteFile\(destinationConn,\s*stagedRemotePath\)[\s\S]{0,120}ERROR_PARTIAL_COPY'
        $curl | Should Match 'GetEntryInfo\(destinationConn,\s*stagedRemotePath,\s*stagedInfo\)'
        $curl | Should Match 'stagedInfo\.sizeKnown\s*&&\s*stagedInfo\.sizeBytes\s*!=\s*fileSize[\s\S]{0,180}RemoteDeleteFile\(destinationConn,\s*stagedRemotePath\)[\s\S]{0,120}ERROR_PARTIAL_COPY'

        $sevenZip | Should Match 'NormalizeArchiveEntryKey[\s\S]{0,900}key\.size\(\)\s*>=\s*2u\s*&&\s*key\[1\]\s*==\s*L.:.[\s\S]{0,160}return\s+\{\}'
        $sevenZip | Should Match 'component\.empty\(\)\s*\|\|\s*component\s*==\s*L"\."\s*\|\|\s*component\s*==\s*L"\.\."'
        $sevenZip | Should Match 'component\.find\(L'':''\)\s*!=\s*std::wstring_view::npos'
        $sevenZip | Should Match 'ensureDir[\s\S]{0,420}if\s*\(\s*!\s*existing->second\.isDirectory\s*\)[\s\S]{0,220}return\s+S_FALSE[\s\S]{0,120}return\s+S_OK'
        $sevenZip | Should Not Match 'outEntries\[raw\.key\]\s*='
        $sevenZip | Should Match 'if\s*\(\s*outEntries\.contains\(raw\.key\)\s*\)[\s\S]{0,180}first indexed entry wins[\s\S]{0,80}continue'

        $s3Io | Should Match 'ValidateS3RangeResponseLength'
        $s3Io | Should Match 'responseBytes\s*!=\s*expectedBytes\s*\|\|\s*bodyBytesRead\s*!=\s*expectedBytes[\s\S]{0,120}ERROR_PARTIAL_COPY'
        $s3Io | Should Match 'RunDebugRangeReadContractSelfTest[\s\S]{0,1200}ERROR_PARTIAL_COPY'
        $s3 | Should Match 'RunDebugRangeReadContractSelfTest\(\*passed,\s*\*failed\)'

        $directoryOps | Should Match 'auto\s+markPartial[\s\S]{0,180}ERROR_PARTIAL_COPY'
        $directoryOps | Should Match '!\s*rootDirectory\s*&&\s*IsNonFatalDirectorySizeChildError\(lastError\)[\s\S]{0,120}markPartial\(\)'
        $watch | Should Match 'else\s*\{\s*EnqueueOverflowLocked\(\);\s*\}[\s\S]{0,180}// Re-arm read immediately before dispatching callbacks'
        $watch | Should Match 'Failed to re-issue directory watch[\s\S]{0,180}EnqueueOverflowLocked\(\)'
        $fileOpsDiagnostics | Should Match 'WriteFile\(file\.get\(\),\s*&bom[\s\S]{0,120}written\s*!=\s*sizeof\(bom\)'
        $fileOpsDiagnostics | Should Match 'WriteFile\(file\.get\(\),\s*header\.data\(\)[\s\S]{0,160}written\s*!=\s*static_cast<DWORD>\(headerBytes\)'
        $fileOpsDiagnostics | Should Match 'WriteFile\(file\.get\(\),\s*line\.data\(\)[\s\S]{0,160}written\s*!=\s*static_cast<DWORD>\(bytesToWrite\)'
        $issuesPane | Should Match 'std::format\(L"\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}\|\{\}"'
        $issuesPane | Should Match 'issue\.operation'
        $issuesPane | Should Match 'issue\.storageType'
        $issuesPane | Should Match 'issue\.destinationStorageType'
        $localIndex | Should Match '\.rs_tmp_'

        $microsoftDrive | Should Match 'incompleteDueToInvalidChildNameOut'
        $microsoftDrive | Should Match 'sourceEnumerationIncomplete[\s\S]{0,260}allChildrenMoved\s*=\s*!\s*sourceEnumerationIncomplete'
        $microsoftDrive | Should Match 'RunDebugEmptyNameChildBlocksSourceDeleteSelfTest'
        $microsoftDrive | Should Match 'RunDebugMoveRejectsInvalidOptionsSizeSelfTest'
        $microsoftDrive | Should Match 'options\s*!=\s*nullptr\s*&&\s*options->sizeBytes\s*!=\s*sizeof\(FileSystemOptions\)'

        $folderView | Should Match 'const\s+bool\s+folderChanged[\s\S]{0,420}if\s*\(\s*folderChanged\s*\)[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $folderViewEnumeration | Should Match 'bool\s+isRefresh\s*=\s*false[\s\S]{0,420}if\s*\(\s*!\s*isRefresh\s*\)[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $folderViewInteraction | Should Match 'focusStayedInView[\s\S]{0,220}newFocus\s*==\s*_hWnd\.get\(\)[\s\S]{0,160}IsChild\(_hWnd\.get\(\),\s*newFocus\)'
        $folderViewInteraction | Should Match 'newFocus\s*!=\s*nullptr\s*&&\s*!\s*focusStayedInView[\s\S]{0,120}ExitIncrementalSearch\(\)'
        $navigationCommands | Should Match 'void\s+FolderWindow::CommandQuickSearch[\s\S]*?PostMessageW\(_hWnd\.get\(\),\s*WndMsg::kPaneRestoreFolderFocus'
        $folderViewInteraction | Should Not Match 'QuickSearch OnKillFocus|QuickSearch Exit|DescribeQuickSearchFocusTarget'
        $navigationCommands | Should Not Match 'CommandQuickSearch:'
        $mainWindow | Should Not Match 'WM_COMMAND QuickSearch'
    }

    It 'keeps Delta reverify file-system safety fixes wired' {
        $bridge = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.State.cpp'
        $phase12 = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Phases10_13.cpp'
        $pathOps = Get-RSText -Path 'Plugins\FileSystem\FileSystem.Path.cpp'
        $search = Get-RSText -Path 'Plugins\FileSystem\FileSystem.Search.cpp'
        $fileOps = Get-RSText -Path 'Plugins\FileSystem\FileSystem.FileOps.cpp'
        $localIndex = Get-RSText -Path 'Common\LocalSearchIndexCore.cpp'

        $pathOps | Should Match 'const\s+DWORD\s+replaceFlags\s*=\s*options\.ignoreReplaceMergeErrors\s*\?\s*REPLACEFILE_IGNORE_MERGE_ERRORS\s*:\s*0u'
        $pathOps | Should Match 'Debug::Warning\(L"FileSystem: promoted ''\{\}'' but failed to set final attributes'

        $search | Should Match 'queued\.status\s*=\s*result\.status'
        $search | Should Match 'if\s*\(\s*SUCCEEDED\(queued\.status\)\s*\)\s*\{\s*result\.status\s*=\s*S_OK;'

        $bridge | Should Match 'RetryTransientCleanupProbe'
        $bridge | Should Match 'const\s+bool\s+sourceIsReparse[\s\S]{0,3200}if\s*\(\s*sourceIsReparse\s*\)[\s\S]{0,420}DeleteCopiedSourcePathNoRecursive'
        $bridge | Should Match 'ReaderMatchesHash[\s\S]{0,3600}while\s*\(\s*readTotal\s*<\s*sizeBytes\s*\)[\s\S]{0,1200}HashBytes\(hash,\s*hashBuffer,\s*bytesRead\)[\s\S]{0,160}readTotal\s*\+=\s*bytesRead'
        $bridge | Should Not Match 'ReadersHaveEqualContent|sourceRead\s*!=\s*destinationRead'
        ([regex]::Matches($bridge, 'moveBridge->cookie\s*=\s*static_cast<void\*>\(&cookie\);')).Count | Should Be 2
        $bridge | Should Match 'bridge\.move\.cleanup\.cancelled'

        $phase12 | Should Match 'bridgeFollowRootReparse'
        $phase12 | Should Match 'reparsePointPolicy":"followTargets'
        $phase12 | Should Match 'Bridge follow-target move reparse removed the junction target file'

        $fileOps | Should Match 'RemoveDirectoryW\(tempPath\.c_str\(\)\)\s*==\s*0[\s\S]{0,180}DeleteFileW\(tempPath\.c_str\(\)\)'
        $fileOps | Should Match 'DeleteFileW\(tempPath\.c_str\(\)\)\s*==\s*0[\s\S]{0,180}RemoveDirectoryW\(tempPath\.c_str\(\)\)'
        $fileOps | Should Match 'if\s*\(\s*!\s*context\.parallel\s*\)[\s\S]{0,180}context\.completedBytes\s*=\s*progress\.itemBaseBytes'
        $fileOps | Should Match 'NormalizeReparseCopyFailure[\s\S]{0,280}!\s*allowOverwrite\s*&&\s*failure\s*==\s*HRESULT_FROM_WIN32\(ERROR_INVALID_PARAMETER\)[\s\S]{0,160}ERROR_NOT_SUPPORTED'
        $fileOps | Should Match 'copyHr\s*=\s*NormalizeReparseCopyFailure\(copyHr,\s*allowOverwriteEffective\)'

        $localIndex | Should Match 'HasGeneratedBridgeTempSuffix'
        $localIndex | Should Match 'HasGeneratedCopyTempSuffix'
        $localIndex | Should Match 'IsGeneratedGuidWithBraces'
        $localIndex | Should Match 'IsRedSalamanderStagedTempName[\s\S]{0,220}HasGeneratedTempTail'
    }

    It 'keeps HostServices connection-secret lifecycle contracts wired' {
        $hostServices = Get-RSText -Path 'RedSalamander\HostServices.cpp'

        $hostServices | Should Match 'HRESULT STDMETHODCALLTYPE GetConnectionJsonUtf8[\s\S]{0,700}if\s*\(\s*!\s*hostWindow\s*\)[\s\S]{0,120}ERROR_INVALID_WINDOW_HANDLE'
        $hostServices | Should Match 'HRESULT STDMETHODCALLTYPE GetConnectionSecret[\s\S]{0,760}if\s*\(\s*!\s*hostWindow\s*\)[\s\S]{0,120}ERROR_INVALID_WINDOW_HANDLE'
        $hostServices | Should Match 'HRESULT STDMETHODCALLTYPE SetConnectionSecret[\s\S]{0,620}if\s*\(\s*!\s*hostWindow\s*\)[\s\S]{0,120}ERROR_INVALID_WINDOW_HANDLE'
        $hostServices | Should Match 'HRESULT STDMETHODCALLTYPE DeleteConnectionSecret[\s\S]{0,520}if\s*\(\s*!\s*hostWindow\s*\)[\s\S]{0,120}ERROR_INVALID_WINDOW_HANDLE'
        $hostServices | Should Match 'HRESULT STDMETHODCALLTYPE ClearCachedConnectionSecret[\s\S]{0,520}if\s*\(\s*!\s*hostWindow\s*\)[\s\S]{0,120}ERROR_INVALID_WINDOW_HANDLE'

        $hostServices | Should Match 'HRESULT HostServices::SetConnectionSecretOnUiThread[\s\S]{0,220}EnsureHostUiThreadReady\(\)'
        $hostServices | Should Match 'SessionSecretEntry&\s+entry\s*=\s*\(\*sessionMap\)\[profile->id\][\s\S]{0,2400}SaveGenericCredential\(targetName,\s*profile->userName,\s*secretView\)'
        $hostServices | Should Match 'entry\.present\s*=\s*false[\s\S]{0,1800}DeleteGenericCredential\(targetName\)[\s\S]{0,360}if\s*\(\s*FAILED\(deleteHr\)'
        $hostServices | Should Match 'SetQuickConnectSecret\(secretKind,\s*secretView\)[\s\S]{0,460}SessionSecretEntry&\s+entry\s*=\s*\(\*sessionMap\)\[profile->id\]'
    }

    It 'keeps Riptide test-credibility gaps closed' {
        $fairstream = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.Fairstream.cpp'
        $popup = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.cpp'
        $popupHeader = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.Popup.h'
        $selfTest = Get-RSText -Path 'RedSalamander\SelfTest\FileOperations\FolderWindow.FileOperations.SelfTest.cpp'

        $fairstream | Should Match 'shouldOverwriteConflict'
        $fairstream | Should Match 'Task::ConflictAction::Overwrite'
        $fairstream | Should Match 'Task::ConflictAction::Skip'
        $fairstream | Should Match 'kSourceBytes'
        $fairstream | Should Match 'kDestinationBytes'
        $fairstream | Should Not Match 'after four skips'

        $popupHeader | Should Match 'DebugBuildFileOperationsGraphFairnessHistorySnapshot'
        $popup | Should Match 'DebugBuildFileOperationsGraphFairnessHistorySnapshot'
        $popup | Should Match 'PopulateGraphHueDebugSummary'
        $fairstream | Should Match 'shape=4-equal-streams-deterministic-history'
        $fairstream | Should Not Match 'ShowWindow\(popup,\s*SW_SHOWNOACTIVATE\)'
        $fairstream | Should Not Match 'shape=4-equal-streams-live-history'

        $selfTest | Should Match 'TryReadPerfMetricSamples'
        $selfTest | Should Match 'fairstreamSaturationDispatchMetricBaseline'
        $selfTest | Should Match 'fairstreamSaturationMaxDistinctInFlightTrees'
        $fairstream | Should Match 'FileOps\.CopyRecursiveParallel\.WorkItemDispatches'
        $fairstream | Should Match 'fairstreamSaturationDispatchMetricBaseline'
        $fairstream | Should Match 'fairstreamSaturationMaxDistinctInFlightTrees'
        $fairstream | Should Match 'Fairstream_SaturationConcurrentCopiesMakeProgress[\s\S]*copyMoveMaxConcurrency":4'
        $fairstream | Should Match 'distinct source trees'
        $fairstream | Should Match 'robust dispatch'
    }

    It 'keeps native DxUi text hosts compatible with Win32 edit text and selection messages' {
        $windowHostSource = Get-RSText -Path 'Common\DxUi\DxUi.WindowHost.cpp'
        $nativeTextInputSource = Get-RSText -Path 'Common\DxUi\DxUi.NativeTextInput.cpp'

        $windowHostSource | Should Match 'case WM_GETTEXTLENGTH:'
        $windowHostSource | Should Match 'case WM_GETTEXT:'
        $windowHostSource | Should Match 'case WM_SETTEXT:'
        $windowHostSource | Should Match 'case EM_GETSEL:'
        $windowHostSource | Should Match 'case EM_SETSEL:'
        $windowHostSource | Should Match 'case EM_REPLACESEL:'

        $nativeTextInputSource | Should Match 'case WM_GETTEXTLENGTH:'
        $nativeTextInputSource | Should Match 'case WM_GETTEXT:'
        $nativeTextInputSource | Should Match 'case WM_SETTEXT:'
        $nativeTextInputSource | Should Match 'case EM_GETSEL:'
        $nativeTextInputSource | Should Match 'case EM_SETSEL:'
        $nativeTextInputSource | Should Match 'case EM_REPLACESEL:'
    }

    It 'gates DxUi focus-sensitive Win32 assertions behind interactive desktop probes' {
        $helpers = Get-RSText -Path 'Tests\DxUiTests\DxUiTestHelpers.h'
        $nativeTextInput = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.NativeTextInput.cpp'
        $menu = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Menu.cpp'

        $helpers | Should Match 'SkipDxUiTest'
        $helpers | Should Match 'TryFocusDxUiTestWindow'
        $helpers | Should Match 'TryActivateDxUiTestWindow'
        $helpers | Should Match 'WaitForDxUiThreadFocus'

        $nativeTextInput | Should Match 'TestNativeTextInputBackendFocusesHostWithoutBridgeChild[\s\S]{0,900}TryFocusDxUiTestWindow\(window\.Hwnd\(\)\)'
        $nativeTextInput | Should Match 'TestNativeTextInputBackendOwnsSystemCaretOnHostHwnd[\s\S]{0,900}TryFocusDxUiTestWindow\(window\.Hwnd\(\)\)'
        $nativeTextInput | Should Match 'SkipDxUiTest\("native text input requires an interactive desktop'

        $menu | Should Match 'TryActivateDxUiTestWindow\(ownerWindow\.Hwnd\(\)\)'
        $menu | Should Match 'SkipDxUiTest\("DxUi menu popup requires an interactive desktop'
        $menu | Should Not Match 'static_cast<void>\(SetForegroundWindow\(ownerWindow\.Hwnd\(\)\)\)'
    }

    It 'keeps MainMenuBarHost HWND teardown inside the DxUi host lifetime' {
        $source = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $classMatch = [regex]::Match($source, 'class\s+MainMenuBarHost\s+final[\s\S]*?\n\};')
        $classMatch.Success | Should Be $true
        $body = $classMatch.Value

        $body | Should Match '~MainMenuBarHost\(\)\s+noexcept[\s\S]{0,80}Destroy\(\);'
        $body | Should Match 'void\s+OnNcDestroy\(HWND\s+hwnd\)\s+noexcept[\s\S]*?_hwnd\.release\(\)'
        $body | Should Match 'SetWindowLongPtrW\(hwnd,\s*GWLP_USERDATA,\s*0\)'

        $hostIndex = $body.IndexOf('RedSalamander::DxUi::WindowHost _host;')
        $hwndIndex = $body.IndexOf('wil::unique_hwnd _hwnd;')
        ($hostIndex -ge 0) | Should Be $true
        ($hwndIndex -ge 0) | Should Be $true
        ($hostIndex -lt $hwndIndex) | Should Be $true
    }

    It 'serializes build and test artifacts and contains every shared process launch tree' {
        $artifactLock = Get-RSText -Path 'Tools\ArtifactOperationLock.ps1'
        $build = Get-RSText -Path 'build.ps1'
        $runner = Get-RSText -Path 'Tools\Run-AllTests.ps1'
        $environment = Get-RSText -Path 'Tools\SanitizedEnvironment.ps1'
        $streaming = Get-RSText -Path 'Tools\ProcessStreaming.ps1'
        $sanitizedMsbuild = Get-RSText -Path 'Tools\Invoke-SanitizedMsbuild.ps1'
        $deploymentTest = Get-RSText -Path 'Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1'

        $artifactLock | Should Match 'artifact-operation\.lock'
        $artifactLock | Should Match '\[System\.IO\.FileShare\]::None'
        $artifactLock | Should Match 'red-salamander\.artifact-operation-lock\.v1'
        $artifactLock | Should Match 'artifact-operation-owner\.json'
        $artifactLock | Should Match '\[System\.IO\.File\]::Replace\(\$temporaryPath, \$metadataPath, \$backupPath\)'
        $artifactLock | Should Match 'owner_pid = \$PID'
        $artifactLock | Should Match 'owner_process_start_utc_ticks = \$OwnerProcessStartUtcTicks'
        $artifactLock | Should Match 'operation = \$Operation'
        $artifactLock | Should Match 'started_utc = \$StartedUtc'
        $artifactLock | Should Match 'REDSALAMANDER_ARTIFACT_OPERATION_TOKEN'
        $artifactLock | Should Match 'OwnerManagedThreadId -eq \$managedThreadId'
        $artifactLock | Should Match 'REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT'
        $artifactLock | Should Match 'Get-RSImmediateParentProcessId -ProcessId \$PID'
        $artifactLock | Should Match 'IsCurrentProcessInJob\(\)'
        $artifactLock | Should Match 'ActiveDelegations\.Count -gt 0'
        $artifactLock | Should Match 'artifact-operation-contaminated\.json'
        $artifactLock | Should Match "Name='MSBuild\.exe' OR Name='cl\.exe' OR Name='link\.exe'"

        $buildLockIndex = $build.IndexOf('Enter-RSArtifactOperationLock')
        $buildProcessPreflightIndex = $build.IndexOf('Stop-BuildOutputProcess -ProcessName')
        $buildLockIndex | Should BeGreaterThan -1
        $buildProcessPreflightIndex | Should BeGreaterThan $buildLockIndex
        $build | Should Match 'Assert-RSNoResidualArtifactToolProcesses -RepoRoot \$SolutionDir'
        $build | Should Match 'finally\s*\{\s*Exit-RSArtifactOperationLock -Lock \$artifactOperationLock'

        $runnerLockIndex = $runner.IndexOf('Enter-RSArtifactOperationLock')
        $runnerContextIndex = $runner.IndexOf('$testRunContext = New-RSTestRunContext')
        $runnerLockIndex | Should BeGreaterThan -1
        $runnerContextIndex | Should BeGreaterThan $runnerLockIndex
        $runner | Should Match 'Assert-RSNoResidualArtifactToolProcesses -RepoRoot \$repoRoot'
        $runner | Should Match 'finally\s*\{\s*Exit-RSArtifactOperationLock -Lock \$artifactOperationLock'

        $environment | Should Match 'JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE|JobObjectLimitKillOnJobClose'
        $environment | Should Match 'ProcThreadAttributeJobList'
        $environment | Should Match 'ProcThreadAttributeHandleList'
        $environment | Should Match 'UpdateAttribute\(attributeList,\s*ProcThreadAttributeJobList'
        ([regex]::Matches($environment, 'new FileStream\((?:parentInput|parentOutput|parentError),\s*FileAccess\.(?:Write|Read),\s*4096,\s*false\)')).Count | Should Be 3
        $jobAttributeIndex = $environment.IndexOf('UpdateAttribute(attributeList, ProcThreadAttributeJobList')
        $createProcessIndex = $environment.IndexOf('if (!CreateProcessW(')
        $jobAttributeIndex | Should BeGreaterThan -1
        $createProcessIndex | Should BeGreaterThan $jobAttributeIndex
        $environment | Should Match 'IsProcessInJob'
        $environment | Should Match 'function Start-RSContainedProcess'
        $environment | Should Match '\[RedSalamander\.Tooling\.ContainedProcess\]::Start\(\$ProcessStartInfo,\s*\$commandLine\)'
        $environment | Should Match 'Start-RSContainedProcess -ProcessStartInfo \$psi -DelegateArtifactOperation'
        $environment | Should Not Match '\.Start\(\)[\s\S]{0,240}AssignProcessToJobObject'
        $streaming | Should Match '\$process = Start-RSContainedProcess -ProcessStartInfo \$psi'
        $streaming | Should Match 'Close-RSContainedProcess -Process \$process'
        $runner | Should Match 'function Invoke-RSSelfTestCaseList[\s\S]*?Start-RSContainedProcess -ProcessStartInfo \$startInfo'
        $runner | Should Match 'function Invoke-RSSelfTestProcess[\s\S]*?Start-RSContainedProcess -ProcessStartInfo \$startInfo'
        $sanitizedMsbuild | Should Match '\$exitCode = Invoke-RSProcess'
        $sanitizedMsbuild | Should Match 'Enter-RSArtifactOperationLock[\s\S]*?finally\s*\{\s*Exit-RSArtifactOperationLock'
        $deploymentTest | Should Match 'Start-RSContainedProcess[\s\S]*?-DelegateArtifactOperation'
        $deploymentTest | Should Match 'Close-RSContainedProcess -Process \$process'
        $deploymentTest | Should Match 'Enter-RSArtifactOperationLock[\s\S]*?AfterAll[\s\S]*?Exit-RSArtifactOperationLock'
    }

    It 'keeps IronLedger FolderView data-integrity guards wired' {
        $dragDrop = Get-RSText -Path 'RedSalamander\FolderView.DragDrop.cpp'
        $fileOps = Get-RSText -Path 'RedSalamander\FolderView.FileOps.cpp'
        $interaction = Get-RSText -Path 'RedSalamander\FolderView.Interaction.cpp'
        $menus = Get-RSText -Path 'RedSalamander\FolderView.Menus.cpp'
        $enumeration = Get-RSText -Path 'RedSalamander\FolderView.Enumeration.cpp'
        $rendering = Get-RSText -Path 'RedSalamander\FolderView.Rendering.cpp'
        $folderWindowFileOps = Get-RSText -Path 'RedSalamander\FolderWindow.FileOperations.cpp'

        $dragDrop | Should Match 'ScreenToClient\(_owner\.GetHWND\(\),\s*&clientPoint\)'
        $dragDrop | Should Match '!\s*IsCurrentFolderEnumerated\(\)[\s\S]{0,120}DRAGDROP_S_CANCEL'
        $dragDrop | Should Match 'HitTest\(clientPoint\)[\s\S]{0,240}GetItemFullPath'
        $dragDrop | Should Match 'header->pathCount\s*>\s*kMaxDropPathCount'
        $dragDrop | Should Match 'GlobalSize\(medium\.hGlobal\)[\s\S]{0,700}dropFiles->pFiles\s*>\s*bytesAvailable'
        $dragDrop | Should Match 'DragQueryFileW\(drop,\s*0xFFFFFFFFu'
        $dragDrop | Should Match 'IsSameOrDescendantPath\(normalizedSource\.native\(\),\s*normalizedDestination\.native\(\)\)'
        $dragDrop | Should Match 'effect\s*==\s*DROPEFFECT_MOVE\s*&&\s*!\s*internalDrop[\s\S]{0,360}reportedEffect\s*=\s*DROPEFFECT_COPY'
        $dragDrop | Should Match '\*performedEffect\s*=\s*reportedEffect'

        $fileOps | Should Match 'totalChars\s*>\s*\(\(std::numeric_limits<size_t>::max\)\(\)\s*-\s*sizeof\(DROPFILES\)\)\s*/\s*sizeof\(wchar_t\)'
        $fileOps | Should Match 'fileCount\s*>\s*kMaxClipboardDropPaths'
        $fileOps | Should Match '!\s*IsCurrentFolderEnumerated\(\)'
        $fileOps | Should Match 'request\.completionCallback[\s\S]{0,420}InvalidateMoveClipboardAfterVerifiedCompletion'

        $interaction | Should Match '!\s*_items\[\*hit\]\.selected[\s\S]{0,100}SelectSingle\(\*hit\)'
        $interaction | Should Match 'ClearSelection\(\);[\s\S]{0,180}_drag\.dragging\s*=\s*false[\s\S]{0,180}_focusedIndex\s*=\s*static_cast<size_t>\(-1\)'
        $menus | Should Match 'targetSnapshot\s*=\s*GetSelectedOrFocusedDisplayNames\(\)'
        $menus | Should Match 'targetBoundCommand[\s\S]{0,700}snapshotStillMatches\(\)'

        $enumeration | Should Match 'kEnumerationReserveCeiling[\s\S]{0,520}clampedCount'
        $enumeration | Should Match 'focusedName\.empty\(\)\s*&&\s*fallbackFocusIndex\s*!=\s*invalidIndex'
        $enumeration | Should Match 'OnCommandMessage\(pending\.commandId\)'
        $rendering | Should Match '_hoveredIndex\s*<\s*_items\.size\(\)[\s\S]{0,100}_items\[_hoveredIndex\]'

        $folderWindowFileOps | Should Match '!\s*request\.sourceContextSpecified[\s\S]{0,900}FindPluginById\(L"builtin/file-system"\)'
        $folderWindowFileOps | Should Match '_fileOperationRequestCompletionCallbacks\.insert_or_assign'
        $folderWindowFileOps | Should Match '!\s*dest\.folderView\.IsCurrentFolderEnumerated\(\)[\s\S]{0,220}destinationUnsettled'
        $folderWindowFileOps | Should Match 'destinationUnsettled\s*\?\s*IDS_MSG_PANE_OP_DESTINATION_LOADING'
    }

    It 'keeps Observatory Track 5 pack-and-delete output safety wired' {
        $commands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Commands.cpp'
        $paths = Get-RSText -Path 'Common\PathUtils.h'

        $paths | Should Match 'NormalizedWindowsPathEqualsNoCase'
        $paths | Should Match 'IsSameOrDescendantNormalizedWindowsPath'
        $commands | Should Match 'ValidateArchiveOutputOutsideSelectedSources\(selectedPaths,\s*archivePath\.value\(\)\)'
        $commands | Should Match 'sourceIsDirectory[\s\S]{0,220}IsSameOrDescendantNormalizedWindowsPath'
        $commands | Should Match 'IDS_MSG_ARCHIVE_OUTPUT_INSIDE_SOURCE'
        $commands | Should Match 'if\s*\(FAILED\(result\.hr\)\)[\s\S]{0,1600}if\s*\(deleteSourcesAfterPack\)'
        $commands | Should Match 'deleteRequest\.operation\s*=\s*FILESYSTEM_DELETE[\s\S]{0,320}StartFileOperationFromFolderView\(pane,\s*std::move\(deleteRequest\)\)'
        $commands | Should Not Match 'DeletePackedSources\('
    }

    It 'keeps Observatory Track 5 local outputs transactional' {
        $transaction = Get-RSText -Path 'Common\LocalFileTransaction.h'
        $commands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Commands.cpp'
        $monitor = Get-RSText -Path 'RedSalamanderMonitor\Document.cpp'

        $transaction | Should Match 'CreateUniqueSiblingFile\(normalizedTarget\.native\(\)'
        $transaction | Should Match 'Common::HandleIo::WriteAll'
        $transaction | Should Match 'FlushFileBuffers\(_file\.get\(\)\)'
        $transaction | Should Match 'MoveFileExW\(_temporaryPath\.c_str\(\),\s*_targetPath\.c_str\(\),\s*moveFlags\)'
        $transaction | Should Match '~LocalFileTransaction\(\)\s+noexcept[\s\S]{0,80}Abort\(\);'
        $transaction | Should Match '!\s*_committed\s*&&\s*!\s*_temporaryPath\.empty\(\)[\s\S]{0,120}DeleteFileW'
        $transaction | Should Match '#ifdef ENABLE_TESTS[\s\S]{0,900}FailNextLocalFileTransactionWrite'
        $transaction | Should Match 'FailNextLocalFileTransactionFlush'

        $commands | Should Match 'WriteMakeFileListUtf8File[\s\S]{0,500}LocalFileTransaction::Create'
        $commands | Should Match 'transaction\.Commit\(static_cast<uint64_t>\(bytes\.size\(\)\)\)'
        $commands | Should Not Match 'WriteMakeFileListUtf8File[\s\S]{0,900}CREATE_ALWAYS'

        $monitor | Should Match 'SaveTextToFile[\s\S]{0,420}LocalFileTransaction::Create'
        $monitor | Should Match 'TryUtf8FromUtf16Strict'
        $monitor | Should Match 'transaction\.Commit\(expectedBytes\)'
        $monitor | Should Not Match 'SaveTextToFile[\s\S]{0,1200}std::ofstream'

        $monitorTest = Get-RSText -Path 'Tests\MonitorTest\MonitorTest.cpp'
        $monitorTest | Should Match 'RunLocalFileTransactionFaultSelfTest'
        $monitorTest | Should Match 'FailNextLocalFileTransactionWrite\(diskFull\)'
        $monitorTest | Should Match 'FailNextLocalFileTransactionFlush\(flushFault\)'
        $monitorTest | Should Match 'Commit\(replacement\.size\(\) \+ 1u\)'
    }

    It 'keeps Observatory Track 5 Make File List responsive, cancellable, and explicit about its destination' {
        $commands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Commands.cpp'
        $folderWindow = Get-RSText -Path 'RedSalamander\FolderWindow.h'
        $messages = Get-RSText -Path 'Common\WindowMessages.h'
        $selftest = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'

        $commands | Should Match 'PromptForMakeFileListOutputFile[\s\S]{0,1800}GetSaveFileNameW'
        $commands | Should Match 'OFN_OVERWRITEPROMPT\s*\|\s*OFN_PATHMUSTEXIST'
        $commands | Should Match 'state\.makeFileListThread\s*=\s*std::jthread'
        $commands | Should Match 'CollectMakeFileListEntries[\s\S]{0,600}const std::stop_token& stopToken'
        $commands | Should Match 'RenderMakeFileListOutput[\s\S]{0,420}const std::stop_token& stopToken'
        $commands | Should Match 'WriteMakeFileListUtf8File[\s\S]{0,850}stopToken\.stop_requested\(\)[\s\S]{0,280}transaction\.Commit'
        $commands | Should Match 'state\.makeFileListThread\.request_stop\(\)'
        $commands | Should Match 'PostMessagePayload\(ownerHwnd,\s*WndMsg::kMakeFileListCompleted'
        $commands | Should Match 'OnMakeFileListCompleted[\s\S]{0,900}Common::Clipboard::TrySetUnicodeText'
        $folderWindow | Should Match 'std::jthread makeFileListThread'
        $messages | Should Match 'kMakeFileListTaskUpdate[\s\S]{0,120}kMakeFileListCompleted'

        $selftest | Should Match 'bulkFileCount\s*=\s*1024u'
        $selftest | Should Match 'DebugSetMakeFileListWorkerDelay\(500u\)'
        $selftest | Should Match 'Cancelled Make File List must not publish an output file'
        $selftest | Should Match 'Make File List command should return promptly'
    }

    It 'keeps Observatory Track 5 archive extraction conflicts explicit and transactional' {
        $commands = Get-RSText -Path 'RedSalamander\FolderWindow.FileSystem.Commands.cpp'
        $settingsSelftest = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        $dialogSelftest = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Dialogs.cpp'

        $commands | Should Match 'enum class ArchiveExistingTargetPolicy[\s\S]{0,160}Skip,[\s\S]{0,80}Replace'
        $commands | Should Match 'BuildArchiveConflictPolicyComboItems[\s\S]{0,260}IDS_ARCHIVE_UNPACK_CONFLICT_SKIP[\s\S]{0,180}IDS_ARCHIVE_UNPACK_CONFLICT_REPLACE'
        $commands | Should Match '_conflictPolicyCombo->SetSelectedIndex\(0u\)'
        $commands | Should Match 'ClassifyArchiveTarget[\s\S]{0,900}ArchiveTargetDecision::Skip'
        $commands | Should Match 'ExtractStoredZipFileEntry[\s\S]{0,900}LocalFileTransaction::Create'
        $commands | Should Match 'SevenZipExtractFileOutStream[\s\S]{0,1400}LocalFileTransaction::Create'
        $commands | Should Match 'skippedConflictCount'

        $settingsSelftest | Should Match 'Unpack skip-existing policy should preserve existing destination file'
        $settingsSelftest | Should Match 'Unpack replace-existing policy should publish the archive content'
        $dialogSelftest | Should Match 'default to skipping existing files'
        $dialogSelftest | Should Match 'DebugSetArchiveUnpackPromptReplaceExisting\(true\)'
    }

    It 'keeps Observatory Track 6 Monitor ownership and bounded-pipeline guards wired' {
        $clipboard = Get-RSText -Path 'Common\UnicodeClipboard.h'
        $colorView = Get-RSText -Path 'RedSalamanderMonitor\ColorTextView.cpp'
        $document = Get-RSText -Path 'RedSalamanderMonitor\Document.h'
        $listener = Get-RSText -Path 'RedSalamanderMonitor\EtwListener.cpp'
        $reader = Get-RSText -Path 'RedSalamanderMonitor\MonitorFileReader.cpp'
        $monitorTest = Get-RSText -Path 'Tests\MonitorTest\MonitorTest.cpp'

        $clipboard | Should Match 'wil::unique_hglobal storage\(GlobalAlloc\(GMEM_MOVEABLE,\s*bytes\)\)'
        $clipboard | Should Match 'SetClipboardData\(CF_UNICODETEXT,\s*storage\.get\(\)\)[\s\S]{0,180}storage\.release\(\)'
        $clipboard | Should Not Match 'Win32UnicodeClipboardApi|TrySetUnicodeTextWithApi|HGLOBAL storage\s*=\s*api\.'

        $colorView | Should Match '_etwEventQueue\.size\(\)\s*>=\s*maxQueuedEvents[\s\S]{0,160}_etwEventQueue\.pop_front\(\)'
        $colorView | Should Match 'EnforceRetentionLimits\(_maxRetainedLines,\s*_maxRetainedTextBytes\)'
        $colorView | Should Match 'monitor\.etw\.queue_high_water_mark'
        $colorView | Should Match 'monitor\.etw\.dropped_count'
        $colorView | Should Match 'monitor\.search\.match_update_us'
        $document | Should Not Match 'const\s+std::(?:vector|deque)<Line>\s*&\s*Lines\('

        $listener | Should Match 'logfile\.Context\s*=\s*callbackState\.get\(\)'
        $listener | Should Match 'TRACEHANDLE stableTraceHandle\s*=\s*traceHandle[\s\S]{0,120}ProcessTrace\(&stableTraceHandle'
        $listener | Should Match 'WaitForSingleObject\(workerHandle,\s*shutdownTimeoutMs\)'
        $listener | Should Match '_workerThread\.detach\(\)'
        $listener | Should Not Match 's_instance|ProcessTrace\(&_traceHandle'

        $reader | Should Match 'GetFileSizeEx[\s\S]{0,700}ERROR_FILE_TOO_LARGE[\s\S]{0,700}bytes\.reserve'
        $reader | Should Match 'stopToken\.stop_requested\(\)'
        $reader | Should Match 'TryUtf16FromUtf8Strict'
        $monitorTest | Should Match 'RunEtwConsumerShutdownSelfTest'
        $monitorTest | Should Match 'kDetachedHandle[\s\S]{0,1800}20u'
        $monitorTest | Should Match 'sparse-over-2gib\.txt'
        $monitorTest | Should Match 'RunUnicodeClipboardContractSelfTest'
    }

    It 'keeps Observatory Track 7 settings commits conflict-aware and recoverable' {
        $header = Get-RSText -Path 'Common\SettingsStore.h'
        $store = Get-RSText -Path 'Common\Common\SettingsStore.cpp'
        $transaction = Get-RSText -Path 'Common\LocalFileTransaction.h'
        $hotReload = Get-RSText -Path 'RedSalamander\SettingsHotReload.cpp'
        $preferences = Get-RSText -Path 'RedSalamander\Preferences.Dialog.cpp'
        $resource = Get-RSText -Path 'RedSalamander\RedSalamander.rc'
        $tests = Get-RSText -Path 'Tests\SettingsSchemaTests\SettingsSchemaTests.cpp'

        $header | Should Match 'SettingsPersistenceState[\s\S]{0,700}std::optional<SettingsFileStamp> expectedFileStamp'
        $header | Should Match 'SaveSettings\(std::wstring_view appId,\s*Settings& settings\)'
        $store | Should Match 'LocalFileTransaction::Create\(path[\s\S]{0,1200}SettingsCommitLock::Acquire'
        $store | Should Match 'currentStamp\s*!=\s*\*expectedStamp[\s\S]{0,160}ERROR_REVISION_MISMATCH'
        $store | Should Match 'settings\.persistence\.expectedFileStamp\s*=\s*committedStamp'
        $store | Should Not Match 'tmpPath\s*\+=\s*L"\.tmp"'
        $transaction | Should Match 'Commit\(std::optional<uint64_t> expectedSize\s*=\s*std::nullopt,\s*BY_HANDLE_FILE_INFORMATION\* committedFileInformation'

        $store | Should Match 'value\s*>\s*\(std::numeric_limits<uint32_t>::max\)\(\)'
        $store | Should Match 'TryUtf8FromUtf16Strict[\s\S]{0,240}ERROR_NO_UNICODE_TRANSLATION'
        $store | Should Match 'sourcePreserved[\s\S]{0,1000}SettingsSavePermission::ExplicitReplacementRequired'
        $store | Should Match 'std::filesystem::exists\(candidate,\s*ec\)'

        $hotReload | Should Match 'struct StampLineage[\s\S]{0,700}expected\s*==\s*lineage->second\.source\s*\|\|\s*expected\s*==\s*lineage->second\.committed'
        $preferences | Should Match 'struct PreferencesSaveResult[\s\S]{0,180}bool mainSaved[\s\S]{0,100}bool monitorSaved'
        $preferences | Should Match 'if\s*\(saveResult\.monitorSaved\)[\s\S]{0,180}monitorBaselineSettings\s*=\s*state\.workingMonitorSettings'
        $preferences | Should Match 'IDS_FMT_PREFS_MONITOR_SAVE_FAILED_AFTER_MAIN'
        $resource | Should Match 'IDS_FMT_PREFS_MONITOR_SAVE_FAILED_AFTER_MAIN[\s\S]{0,500}\{0\}[\s\S]{0,220}\{1:08X\}'

        $tests | Should Match '--settings-store-cas-child'
        $tests | Should Match 'ERROR_REVISION_MISMATCH'
        $tests | Should Match 'ERROR_NO_UNICODE_TRANSLATION'
        $tests | Should Match 'values above UINT32_MAX are rejected instead of truncated'
        $tests | Should Match 'failed backup leaves defaults save-blocked'
    }

    It 'keeps Observatory Track 8 connection identity and authorization boundaries wired' {
        $settingsHeader = Get-RSText -Path 'Common\SettingsStore.h'
        $settingsStore = Get-RSText -Path 'Common\Common\SettingsStore.cpp'
        $secretsHeader = Get-RSText -Path 'RedSalamander\ConnectionSecrets.h'
        $secrets = Get-RSText -Path 'RedSalamander\ConnectionSecrets.cpp'
        $manager = Get-RSText -Path 'RedSalamander\ConnectionManagerWindow.cpp'
        $hostServices = Get-RSText -Path 'RedSalamander\HostServices.cpp'
        $app = Get-RSText -Path 'RedSalamander\RedSalamander.cpp'
        $settingsTests = Get-RSText -Path 'Tests\SettingsSchemaTests\SettingsSchemaTests.cpp'
        $connectionTests = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Connections.cpp'

        $settingsHeader | Should Match 'CreateConnectionProfileId'
        $settingsHeader | Should Match 'NormalizeConnectionProfileId'
        $settingsHeader | Should Match 'ConnectionProfileIdMigration'
        $settingsStore | Should Match 'normalizedIdCounts\[canonicalId\]\s*==\s*1u'
        $settingsStore | Should Match 'profile\.savePassword\s*=\s*false'
        $settingsStore | Should Match 'ValidateConnectionProfileIds\(settings\.connections\.value\(\)\)'
        $manager | Should Match 'ValidateConnectionProfileIds\(connSettings\)'
        $secrets | Should Match 'NormalizeConnectionProfileId\(connectionId,\s*canonicalId\)'

        $secretsHeader | Should Match 'enum class SecretAccessPurpose[\s\S]{0,120}Interactive[\s\S]{0,80}Background'
        $secrets | Should Match 'SecretAccessAuthorizationKey[\s\S]{0,260}SecretKind kind[\s\S]{0,120}SecretAccessPurpose purpose'
        $manager | Should Match 'SecretAccessPurpose::Interactive'
        $manager | Should Not Match 'HasSecretAccessAuthorization'
        $hostServices | Should Match 'SecretAccessPurpose::Background'
        $hostServices | Should Match 'ClearAllSecretAccessAuthorizations\(\)'
        $app | Should Match 'WTSRegisterSessionNotification'
        $app | Should Match 'WM_WTSSESSION_CHANGE'
        $app | Should Match 'HostClearConnectionSessionState\(\)'

        $settingsTests | Should Match 'strict reload rejects case-colliding IDs'
        $settingsTests | Should Match 'ambiguous saved-secret references are cleared instead of copied or aliased'
        $connectionTests | Should Match 'connection_secret_authorization_scopes'
        $connectionTests | Should Match 'unsigned tick wrap'
        $connectionTests | Should Match 'background continuation grant'
    }

    It 'keeps Observatory Track 9 DxUi identity, reentrancy, and text-input guards wired' {
        $controls = Get-RSText -Path 'Common\DxUi\DxUi.Controls.cpp'
        $accessibility = Get-RSText -Path 'Common\DxUi\DxUi.Accessibility.cpp'
        $nativeMenu = Get-RSText -Path 'Common\DxUi\DxUiNativeMenuInterop.h'
        $nativeTextInput = Get-RSText -Path 'Common\DxUi\DxUi.NativeTextInput.cpp'
        $textInput = Get-RSText -Path 'Common\DxUi\DxUi.TextInput.cpp'
        $tree = Get-RSText -Path 'Common\DxUi\DxUi.Tree.cpp'
        $navigationPopup = Get-RSText -Path 'RedSalamander\NavigationView.FullPathPopup.cpp'
        $controlTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Controls.cpp'
        $menuTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Menu.cpp'
        $treeTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Tree.cpp'
        $textFieldTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.TextField.cpp'
        $nativeTextInputTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.NativeTextInput.cpp'
        $accessibilityTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Accessibility.cpp'
        $commandTests = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'

        $controls | Should Match 'const auto openItem\s*=\s*_onOpenItem;[\s\S]{0,120}RequestInvalidate\(\);[\s\S]{0,80}openItem\('

        $nativeMenu | Should Match 'SyncMenuModelInternal\(bool invokeRefresh\)[\s\S]{0,500}menuBarLifetime\.expired\(\)'
        $nativeMenu | Should Match 'ContextMenu::Show\(ownerWindow[\s\S]{0,320}menuBarLifetime\.expired\(\)\s*\|\|\s*_menuBar\s*!=\s*menuBar'
        $nativeMenu | Should Match 'message\s*!=\s*WM_NCDESTROY\s*&&\s*hadMenuBar\s*&&\s*menuBarLifetime\.expired\(\)'

        $tree | Should Match 'OnTreeSelectionChanged\(item\.id\);[\s\S]{0,120}selfLifetime\.expired\(\)'
        $tree | Should Match 'FindVisibleItemById\(hitItem\.id\)'
        $tree | Should Match 'FindVisibleItemById\(item\.id\)\.has_value\(\)'

        $accessibility | Should Match 'SetTreeItemRuntimeId\([\s\S]{0,500}itemId\s*&\s*0xFFFFFFFFull[\s\S]{0,200}itemId\s*>>\s*32u'
        $accessibility | Should Match 'ResolveTreeVisibleIndex[\s\S]{0,500}FindVisibleItemById\(_treeItemId\)'
        $accessibility | Should Match 'UIA_E_ELEMENTNOTAVAILABLE'
        $accessibility | Should Match 'ExpandToEnclosingUnit\(TextUnit unit\)[\s\S]{0,4200}GetEnclosingTextRangeCharacterSpan[\s\S]{0,800}GetEnclosingTextRangeWordSpan[\s\S]{0,1200}TryGetEnclosingTextRangeVisualLineSpan'

        $nativeTextInput | Should Match '_nativeTextInputImeComposing[\s\S]{0,500}ImportTextInputState\(\*this,\s*_nativeTextInputImeBaseState\.value\(\),\s*false\)[\s\S]{0,500}DeactivateNativeTextInputTsf\(\)'
        $textInput | Should Match 'ReplaceSelectionAndNotify[\s\S]{0,1200}SyncTextInput\(this\)[\s\S]{0,160}RefreshAccessibilitySnapshot\(\)[\s\S]{0,160}RequestInvalidate\(\)[\s\S]{0,100}NotifyChanged\(\)'
        $textInput | Should Match 'bool TextField::NotifyChanged\(\)[\s\S]{0,700}selfLifetime[\s\S]{0,200}onTextChanged\(textSnapshot\)[\s\S]{0,200}selfLifetime\.expired\(\)'
        $textInput | Should Match 'ControlTextIndexToDisplayTextIndex'
        $textInput | Should Match 'DisplayTextIndexToControlTextIndex'
        $textInput | Should Match 'ControlTextRangeToDisplayTextRange'
        $textInput | Should Match 'dxui\.textinput\.masked_index_map_rebuild_us'

        $navigationPopup | Should Match 'GetWindow\(activatingWindow,\s*GW_OWNER\)\s*==\s*popupHwnd'
        $controlTests | Should Match 'TestMenuBarActivationCanReplaceRootSafely'
        $menuTests | Should Match 'TestNativeMenuBarNestedPopupCanDestroyHostSafely'
        $treeTests | Should Match 'TestTreeExpanderReResolvesStableItemAfterSelectionReorder'
        $treeTests | Should Match 'TestTreeSelectionDelegateCanReplaceRootSafely'
        $textFieldTests | Should Match 'TestTextFieldReplaceSelectionSynchronizesBeforeTerminalNotification'
        $textFieldTests | Should Match 'TestMaskedTextFieldGeometryMapsUtf16SourceToDisplayElements'
        $nativeTextInputTests | Should Match 'native ime window deactivation restores the pre-composition text'
        $nativeTextInputTests | Should Match 'native ime app deactivation restores the pre-composition text'
        $accessibilityTests | Should Match 'TestAccessibilityTreeItemProviderKeepsStableIdentityAcrossReorder'
        $accessibilityTests | Should Match 'removed retained tree-item provider reports element-not-available'
        $accessibilityTests | Should Match 'character expansion keeps the complete ZWJ emoji cluster'
        $accessibilityTests | Should Match 'cross-thread TextRange line expansion normalizes to the visual line'
        $commandTests | Should Match 'cmd_pane_navigationView_full_path_popup_owned_window_activation'
    }

    It 'keeps Observatory Track 10 cloud paging, upload acknowledgement, commit, and identity guards wired' {
        $pager = Get-RSText -Path 'Common\PaginationGuard.h'
        $microsoft = Get-RSText -Path 'Plugins\FileSystemMicrosoftDrive\FileSystemMicrosoftDrive.cpp'
        $google = Get-RSText -Path 'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.cpp'
        $googleHeader = Get-RSText -Path 'Plugins\FileSystemGoogleDrive\FileSystemGoogleDrive.h'
        $s3Directory = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $s3DirectoryOps = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.DirectoryOps.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.S3.cpp'
        $s3Table = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.S3Table.cpp'
        $themeTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Theme.cpp'
        $pluginTests = Get-RSText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'

        $pager | Should Match 'class ContinuationGuard final'
        $pager | Should Match 'maxPages'
        $pager | Should Match 'maxItems'
        $pager | Should Match 'maxBytes'
        $pager | Should Match 'deadlineTickMs'
        $pager | Should Match 'cancellationProbe'
        $pager | Should Match '_seenTokens\.emplace'
        $themeTests | Should Match 'TestCloudPaginationGuardBoundsProgressAndCancellation'

        $microsoft | Should Match 'ParseNextExpectedUploadOffset'
        ([regex]::Matches($microsoft, 'ParseNextExpectedUploadOffset\(uploadResponse\.body')).Count | Should Be 2
        $microsoft | Should Match 'Graph item DELETE is recursive and this provider has no atomic'
        $microsoft | Should Match 'MoveCommitResult'
        $microsoft | Should Match 'move committed but overwrite-backup cleanup remains pending'
        $microsoft | Should Match 'RunDebugMergeNeverRecursivelyDeletesSourceFolderSelfTest'
        $microsoft | Should Match 'RunDebugCommittedMoveCleanupFailureIsWarningSelfTest'

        $google | Should Match 'CURLOPT_TIMEOUT_MS'
        $google | Should Match 'kMaxJsonResponseBytes'
        $google | Should Match 'kMaxAuthorizedRetries'
        $google | Should Match '_tokenRefreshesInFlight\.contains'
        $google | Should Match 'MakeExposedItemName'
        $google | Should Match 'same \[id:AbC\] \[id:literal\]'
        $googleHeader | Should Match 'std::condition_variable _tokenCv'
        $pluginTests | Should Match 'RedSalamanderGoogleDriveDebugSelfTests'

        $s3Directory | Should Match 'Common::Paging::Utf8ContinuationGuard pager'
        $s3DirectoryOps | Should Match 'Common::Paging::Utf8ContinuationGuard pager'
        $s3 | Should Match 'Common::Paging::Utf8ContinuationGuard pager'
        ([regex]::Matches($s3Table, 'Common::Paging::Utf8ContinuationGuard pager')).Count | Should Be 2
        $s3Directory | Should Match 'RunDebugPaginationGuardSelfTest'
    }

    It 'keeps Observatory Track 11 S3 transfer, cleanup, and unload contracts wired' {
        $shared = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Shared.cpp'
        $factory = Get-RSText -Path 'Plugins\FileSystemS3\Factory.cpp'
        $directory = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.Directory.cpp'
        $directoryOps = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.DirectoryOps.cpp'
        $io = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.IO.cpp'
        $s3 = Get-RSText -Path 'Plugins\FileSystemS3\FileSystemS3.S3.cpp'
        $pluginTests = Get-RSText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'

        $shared | Should Match 'State::Initializing'
        $shared | Should Match '_refCount\s*=\s*1u'
        $shared | Should Match 'RunDebugAwsSdkLifetimeContractSelfTest'
        $factory | Should Match 'SchedulePendingMultipartAbortCleanup'
        $factory | Should Match 'CanUnloadPendingMultipartAbortCleanup[\s\S]{0,180}AwsSdkLifetime::CanUnloadNow'
        $s3 | Should Match 'ValidateS3UploadReadResult\(sizeBytes,\s*body->GetConsumedBytes\(\)'
        $io | Should Match 'class PendingMultipartAbortQueue final'
        $io | Should Match 'TransferModulePinToCallbackReturn'
        $io | Should Match 'failed destructor abort should be retried asynchronously'
        $directoryOps | Should Match 'RunDebugDirectorySizeCallbackContractSelfTest'
        $directory | Should Match 'RunDebugRecursiveDeleteConvergenceSelfTest'
        $directory | Should Match 'RunDebugCommittedCleanupDebtSelfTest'
        $pluginTests | Should Match 'TestS3RuntimeUnloadContract'
    }

    It 'keeps Observatory Track 12 plugin configuration, secret, runtime, and cancellation contracts wired' {
        $jsonHelpers = Get-RSText -Path 'Common\YyjsonHelpers.h'
        $curlRuntime = Get-RSText -Path 'Common\CurlProcessRuntime.h'
        $fileSystemContract = Get-RSText -Path 'Common\PlugInterfaces\FileSystem.h'
        $directoryCache = Get-RSText -Path 'RedSalamander\DirectoryInfoCache.cpp'
        $sevenZipHeader = Get-RSText -Path 'Plugins\FileSystem7z\FileSystem7z.h'
        $sevenZip = Get-RSText -Path 'Plugins\FileSystem7z\FileSystem7z.cpp'
        $curlFactory = Get-RSText -Path 'Plugins\FileSystemCurl\Factory.cpp'
        $googleFactory = Get-RSText -Path 'Plugins\FileSystemGoogleDrive\Factory.cpp'
        $pluginTests = Get-RSText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'

        $jsonHelpers | Should Match 'ParseObjectDocument'
        $jsonHelpers | Should Match 'WriteObjectWithoutMembers'
        $pluginTests | Should Match 'malformed configuration preserves live state'
        $pluginTests | Should Match 'wrong-root configuration preserves live state'
        $pluginTests | Should Match 'legacy configuration secrets are imported but not persisted'

        $curlRuntime | Should Match 'class ProcessLease final'
        $curlRuntime | Should Match 'final participant performs curl_global_cleanup outside loader lock'
        $curlFactory | Should Match 'RedSalamanderPluginShutdown'
        $curlFactory | Should Match 'RedSalamanderPluginCanUnloadNow'
        $googleFactory | Should Match 'RedSalamanderPluginShutdown'
        $googleFactory | Should Match 'RedSalamanderPluginCanUnloadNow'
        $pluginTests | Should Match 'TestCrossPluginCurlRuntimeUnloadContract'
        $pluginTests | Should Match 'still creates a libcurl easy handle after peer unload'

        $fileSystemContract | Should Match 'IFileSystemCancellableDirectoryEnumeration'
        $directoryCache | Should Match 'StopTokenDirectoryEnumerationCallback'
        $directoryCache | Should Match 'ReadDirectoryInfoCancellable'
        $sevenZipHeader | Should Match 'std::atomic<uint64_t> _indexGeneration'
        $sevenZip | Should Match 'RunDebugArchiveIndexCancellationSelfTest'
        $sevenZip | Should Match "kMaxIndexItems\s*=\s*1'000'000u"
        $sevenZip | Should Match 'kMaxIndexTextBytes\s*=\s*128u\s*\*\s*1024u\s*\*\s*1024u'
    }

    It 'keeps Observatory Track 13 local and Dummy provider mutations fail-before-change' {
        $local = Get-RSText -Path 'Plugins\FileSystem\FileSystem.cpp'
        $localPath = Get-RSText -Path 'Plugins\FileSystem\FileSystem.Path.cpp'
        $dummy = Get-RSText -Path 'Plugins\FileSystemDummy\FileSystemDummy.cpp'
        $pluginTests = Get-RSText -Path 'Tests\PluginContractTests\PluginContractTests.cpp'

        $local | Should Match 'allowReplaceReadOnly\s*&&\s*!\s*allowOverwrite[\s\S]{0,100}return E_INVALIDARG'
        $local | Should Not Match 'SetFileAttributesW\(filePath\.c_str\(\),\s*attributes\s*&\s*~FILE_ATTRIBUTE_READONLY\)'
        $localPath | Should Match 'for \(unsigned int attempt = 0u; attempt < 8u; \+\+attempt\)'
        $localPath | Should Match 'written < absolute\.size\(\)'
        $localPath | Should Match 'required\s*=\s*written'
        $localPath | Should Match 'kMaximumAbsolutePathChars'

        $dummy | Should Match 'ValidateDirectoryMerge\(sourceDirectory,\s*destinationDirectory,\s*flags\)'
        $dummy | Should Match 'ValidateDeleteNode\(target,\s*flags\)'
        $dummy | Should Match 'Lazy descendants can contain READONLY nodes'
        $pluginTests | Should Match 'local writer rejects replace-readonly without overwrite and preserves destination attributes'
        $pluginTests | Should Match 'dummy directory copy preflights a late collision without partial destination mutation'
        $pluginTests | Should Match 'dummy directory move preflights a late collision without partial source or destination mutation'
        $pluginTests | Should Match 'dummy recursive delete applies read-only policy to descendants before mutation'
    }

    It 'keeps Observatory Track 14 final-save, stale-suggestion, and bounded-menu contracts wired' {
        $settings = Get-RSText -Path 'RedSalamander\SettingsHotReload.cpp'
        $settingsTests = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.Settings.cpp'
        $navigationHeader = Get-RSText -Path 'RedSalamander\NavigationView.h'
        $navigationEdit = Get-RSText -Path 'RedSalamander\NavigationView.Edit.cpp'
        $navigationMenus = Get-RSText -Path 'RedSalamander\NavigationView.Menus.cpp'
        $navigationTests = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.SelfTest.ViewCommands.cpp'
        $menu = Get-RSText -Path 'Common\DxUi\DxUi.Menu.cpp'
        $menuTests = Get-RSText -Path 'Tests\DxUiTests\DxUiTests.Menu.cpp'

        $settings | Should Match 'enum class SettingsSaveShutdownState[\s\S]{0,220}Running[\s\S]{0,100}FinalSavePending[\s\S]{0,100}FinalSaveQueued[\s\S]{0,100}ShuttingDown'
        $settings | Should Match 'BeginProcessShutdown\(\)[\s\S]{0,180}_submissionMutex[\s\S]{0,220}FinalSavePending'
        $settings | Should Match '_finalSaveCompletion\s*=\s*completion[\s\S]{0,140}FinalSaveQueued'
        $settings | Should Not Match '_processShutdownStarted'
        $settingsTests | Should Match 'SettingsSaveChildMode::FinalSaveOrdering'
        $settingsTests | Should Match 'Process finalization should admit exactly one settings snapshot'

        $navigationHeader | Should Match 'struct EditSuggestResultsPayload[\s\S]{0,260}requestId[\s\S]{0,100}editSessionId[\s\S]{0,180}queryText'
        $navigationEdit | Should Match 'owned->requestId\s*!=\s*_editSuggestRequestId\.load[\s\S]{0,180}owned->editSessionId\s*!=\s*_editSuggestEditSessionId[\s\S]{0,120}owned->queryText\s*!=\s*_pathEdit->field->GetText'
        $navigationEdit | Should Match 'ApplyEditSuggestIndex[\s\S]{0,1200}_editSuggestRequestId\.fetch_add'
        $navigationTests | Should Match 'DebugPostCurrentNavigationEditSuggestResult'
        $navigationMenus | Should Match 'kMaxSiblingItems\s*=\s*static_cast<size_t>\(ID_SIBLING_SEARCH\s*-\s*ID_SIBLING_BASE\)'
        $navigationHeader | Should Match 'ID_SIBLING_SEARCH\s*=\s*699'

        $menu | Should Match 'std::vector<float> itemOffsetsDip'
        ([regex]::Matches($menu, 'std::upper_bound\(popup(?:\.|->)itemOffsetsDip')).Count | Should BeGreaterThan 1
        $menu | Should Match 'dxui\.menu\.popup\.visible_rows'
        $menuTests | Should Match 'TestLargeMenuPaintsOnlyVisibleRowsWithCachedOffsets'
        $menuTests | Should Match 'kItemCount\s*=\s*4096u'
        $menuTests | Should Match 'lastPaintedItemCount\s*<=\s*32u'
    }

    It 'keeps production translation units independently compiled' {
        $productionRoots = @(
            'Common',
            'Plugins',
            'RedConfigure',
            'RedLauncher',
            'RedSalamander',
            'RedSalamanderMonitor',
            'RedSalamanderSearchService'
        )
        $implementationIncludePattern = '(?m)^\s*#\s*include\s*"[^"\r\n]+\.cpp"'
        $violations = @(
            foreach ($root in $productionRoots) {
                Get-ChildItem -LiteralPath (Join-Path $repoRoot $root) -Recurse -File |
                    Where-Object { $_.Extension -in @('.cpp', '.h', '.hpp') } |
                    Where-Object { $_.FullName -notmatch '[\\/]SelfTest[\\/]' } |
                    ForEach-Object {
                        $source = Get-Content -LiteralPath $_.FullName -Raw
                        if ($source -match $implementationIncludePattern) {
                            [System.IO.Path]::GetRelativePath($repoRoot, $_.FullName)
                        }
                    }
            }
        )

        $violations.Count | Should Be 0
    }
}
