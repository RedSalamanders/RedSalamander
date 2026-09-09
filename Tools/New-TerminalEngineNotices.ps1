<#
.SYNOPSIS
Materializes a deterministic third-party notices file from a Terminal manifest.

.DESCRIPTION
Reads the reviewed Terminal notice manifest, requires the named dependency records, writes the combined notices file, and writes an identity receipt binding the output to its inputs.

.PARAMETER ManifestFile
Specifies the reviewed notice manifest.

.PARAMETER OutputFile
Specifies the notices text file to create or replace.

.PARAMETER ReceiptFile
Specifies the machine-readable receipt file to create or replace.

.PARAMETER RequiredDependencyId
Lists every dependency identifier that must be represented in the notice output.

.PARAMETER PassThruReceipt
Returns the resolved receipt path after successful materialization.

.OUTPUTS
With -PassThruReceipt, a string containing the receipt path. Otherwise no pipeline output. Always writes OutputFile and ReceiptFile.

.NOTES
Prerequisites: a valid reviewed notice manifest and all referenced source texts. Side effects: replaces the requested output and receipt files; callers must choose repository build/evidence destinations deliberately. Exit is nonzero on missing dependencies, unsafe/invalid content, or identity mismatch. Primary consumers: historical Terminal qualification and notice-materialization tests; the current Ghostty builder consumes its reviewed notice texts directly.

.EXAMPLE
.\Tools\New-TerminalEngineNotices.ps1 -ManifestFile .\.build\notices.json -OutputFile .\.build\THIRD-PARTY-NOTICES.txt -ReceiptFile .\.build\notices-receipt.json -RequiredDependencyId ghostty -PassThruReceipt
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ManifestFile,

    [Parameter(Mandatory)]
    [string]$OutputFile,

    [Parameter(Mandatory)]
    [string]$ReceiptFile,

    [Parameter(Mandatory)]
    [string[]]$RequiredDependencyId,

    [switch]$PassThruReceipt
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$modulePath = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineNotices.psm1')
Import-Module $modulePath -Force -ErrorAction Stop

$result = Invoke-RSTerminalNoticeMaterialization `
    -ManifestFile $ManifestFile `
    -OutputFile $OutputFile `
    -ReceiptFile $ReceiptFile `
    -RequiredDependencyId $RequiredDependencyId `
    -RepoRoot $repoRoot

if ($PassThruReceipt) {
    $result.ReceiptPath
}
