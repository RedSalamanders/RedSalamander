Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:NoticeFormatId = 'red-salamander-terminal-engine-notices'
$script:NoticeReceiptFormatId = 'red-salamander-terminal-engine-notice-receipt'
$script:NoticeSchemaVersion = 1

$script:RepoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
Import-Module (
    Join-Path $PSScriptRoot 'TerminalEngineLifecycle.psm1') -Force -ErrorAction Stop
Import-Module (
    Join-Path $script:RepoRoot 'Tools\TerminalEvidence.psm1') -Force -ErrorAction Stop

function Get-RSNoticeProperty {
    param(
        [Parameter(Mandatory)]
        [Collections.IDictionary]$Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not $Object.Contains($Name)) {
        throw "Required notice property '$Name' is missing."
    }
    return $Object[$Name]
}

function Assert-RSNoticeExactProperties {
    param(
        [Parameter(Mandatory)]
        [Collections.IDictionary]$Object,

        [Parameter(Mandatory)]
        [string[]]$Expected,

        [Parameter(Mandatory)]
        [string]$Context
    )

    $expectedSet = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::Ordinal)
    foreach ($name in $Expected) {
        [void]$expectedSet.Add($name)
    }
    foreach ($nameValue in $Object.Keys) {
        $name = [string]$nameValue
        if (-not $expectedSet.Contains($name)) {
            throw "$Context contains unsupported property '$name'."
        }
    }
    foreach ($name in $Expected) {
        if (-not $Object.Contains($name)) {
            throw "$Context is missing required property '$name'."
        }
    }
}

function Assert-RSNoticeIdentifier {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($Value -cnotmatch '^[A-Za-z0-9][A-Za-z0-9._+-]{0,127}$') {
        throw "$Name must be a nonempty stable identifier."
    }
}

function Assert-RSNoticeSha256 {
    param(
        [Parameter(Mandatory)]
        [string]$Value,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($Value -cnotmatch '^[0-9a-f]{64}$') {
        throw "$Name must be exactly 64 lowercase hexadecimal characters."
    }
}

function Assert-RSNoticeProvenance {
    param(
        [Parameter(Mandatory)]
        [Collections.IDictionary]$InputRecord,

        [Parameter(Mandatory)]
        [string]$Context
    )

    $sourceUrl = [string](Get-RSNoticeProperty `
            -Object $InputRecord `
            -Name 'sourceUrl')
    $uri = $null
    if (-not [Uri]::TryCreate(
            $sourceUrl,
            [UriKind]::Absolute,
            [ref]$uri) -or
        $uri.Scheme -cne 'https' -or
        -not [string]::IsNullOrEmpty($uri.UserInfo)) {
        throw "$Context sourceUrl must be one credential-free HTTPS URL."
    }

    $retrievedUtc = [string](Get-RSNoticeProperty `
            -Object $InputRecord `
            -Name 'retrievedUtc')
    $retrieved = [DateTimeOffset]::MinValue
    if ($retrievedUtc -cnotmatch
            '^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$' -or
        -not [DateTimeOffset]::TryParseExact(
            $retrievedUtc,
            'yyyy-MM-ddTHH:mm:ssZ',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::AssumeUniversal,
            [ref]$retrieved)) {
        throw "$Context retrievedUtc must be an exact whole-second UTC timestamp."
    }

    Assert-RSNoticeSha256 `
        -Value ([string](Get-RSNoticeProperty `
                -Object $InputRecord `
                -Name 'sourceSha256')) `
        -Name "$Context sourceSha256"

    $sourceMember = [string](Get-RSNoticeProperty `
            -Object $InputRecord `
            -Name 'sourceMember')
    if ($sourceMember.Contains('\') -or
        [IO.Path]::IsPathFullyQualified($sourceMember) -or
        $sourceMember.StartsWith('/', [StringComparison]::Ordinal) -or
        @($sourceMember.Split('/') | Where-Object { $_ -ceq '..' }).Count -ne 0) {
        throw "$Context sourceMember must be empty or one safe archive-relative path."
    }
}

function Resolve-RSNoticeRepositoryPath {
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [string]$Path,

        [switch]$MustExist
    )

    $canonicalRoot = [IO.Path]::GetFullPath($RepoRoot)
    $candidate = if ([IO.Path]::IsPathFullyQualified($Path)) {
        [IO.Path]::GetFullPath($Path)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $canonicalRoot $Path))
    }
    if (-not (Test-RSPathIsWithin -Path $candidate -Root $canonicalRoot)) {
        throw "Notice path escapes the repository root: $Path"
    }
    $relative = [IO.Path]::GetRelativePath($canonicalRoot, $candidate)
    return Resolve-RSTerminalPathInRoot `
        -Root $canonicalRoot `
        -Path $relative `
        -MustExist:$MustExist
}

function Get-RSNoticeBytesSha256 {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes
    )

    $hasher = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = $hasher.ComputeHash($Bytes)
        return ([Convert]::ToHexString($digest)).ToLowerInvariant()
    }
    finally {
        $hasher.Dispose()
    }
}

function Read-RSNoticeUtf8Text {
    param(
        [Parameter(Mandatory)]
        [byte[]]$Bytes,

        [Parameter(Mandatory)]
        [string]$InputId
    )

    if ($Bytes.Length -ge 3 -and
        $Bytes[0] -eq 0xEF -and
        $Bytes[1] -eq 0xBB -and
        $Bytes[2] -eq 0xBF) {
        throw "Notice input '$InputId' must be UTF-8 without a byte-order mark."
    }

    $encoding = [Text.UTF8Encoding]::new($false, $true)
    try {
        $text = $encoding.GetString($Bytes)
    }
    catch [Text.DecoderFallbackException] {
        throw "Notice input '$InputId' is not valid UTF-8."
    }
    if ($text.IndexOf([char]0) -ge 0) {
        throw "Notice input '$InputId' contains a NUL character."
    }
    return $text.Replace("`r`n", "`n").Replace("`r", "`n").
        TrimEnd([char]"`n")
}

function Add-RSNoticeLine {
    param(
        [Parameter(Mandatory)]
        [Text.StringBuilder]$Builder,

        [AllowEmptyString()]
        [string]$Text = ''
    )

    [void]$Builder.Append($Text)
    [void]$Builder.Append("`n")
}

function Write-RSNoticeFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [byte[]]$Bytes
    )

    $parent = [IO.Path]::GetDirectoryName($Path)
    if (-not [string]::IsNullOrEmpty($parent)) {
        [void](New-Item -ItemType Directory -Path $parent -Force)
    }
    [IO.File]::WriteAllBytes($Path, $Bytes)
}

function Invoke-RSTerminalNoticeMaterialization {
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

        [string]$RepoRoot = $script:RepoRoot
    )

    $canonicalRoot = [IO.Path]::GetFullPath($RepoRoot)
    $manifestPath = Resolve-RSNoticeRepositoryPath `
        -RepoRoot $canonicalRoot `
        -Path $ManifestFile `
        -MustExist
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw 'The terminal-engine notice manifest must be a file.'
    }
    $outputPath = Resolve-RSNoticeRepositoryPath `
        -RepoRoot $canonicalRoot `
        -Path $OutputFile
    $receiptPath = Resolve-RSNoticeRepositoryPath `
        -RepoRoot $canonicalRoot `
        -Path $ReceiptFile
    $reservedPaths = [Collections.Generic.HashSet[string]]::new(
        [StringComparer]::OrdinalIgnoreCase)
    [void]$reservedPaths.Add($manifestPath)
    if (-not $reservedPaths.Add($outputPath) -or
        -not $reservedPaths.Add($receiptPath)) {
        throw 'ManifestFile, OutputFile, and ReceiptFile must be distinct paths.'
    }

    $requiredIds = @($RequiredDependencyId)
    if ($requiredIds.Count -eq 0) {
        throw 'RequiredDependencyId must name the complete frozen dependency closure.'
    }
    $lastRequiredId = $null
    foreach ($dependencyId in $requiredIds) {
        Assert-RSNoticeIdentifier `
            -Value $dependencyId `
            -Name 'RequiredDependencyId'
        if ($null -ne $lastRequiredId -and
            [StringComparer]::Ordinal.Compare(
                $lastRequiredId,
                $dependencyId) -ge 0) {
            throw 'RequiredDependencyId values must be unique and sorted ordinally.'
        }
        $lastRequiredId = $dependencyId
    }

    $manifestBytes = [IO.File]::ReadAllBytes($manifestPath)
    if ($manifestBytes.Length -ge 3 -and
        $manifestBytes[0] -eq 0xEF -and
        $manifestBytes[1] -eq 0xBB -and
        $manifestBytes[2] -eq 0xBF) {
        throw 'The terminal-engine notice manifest must be UTF-8 without a byte-order mark.'
    }
    $manifestText = [Text.UTF8Encoding]::new($false, $true).GetString(
        $manifestBytes)
    $manifest = $manifestText |
        ConvertFrom-Json `
            -AsHashtable `
            -Depth 100 `
            -NoEnumerate `
            -DateKind String `
            -ErrorAction Stop
    if ($manifest -isnot [Collections.IDictionary]) {
        throw 'The terminal-engine notice manifest root must be an object.'
    }
    Assert-RSNoticeExactProperties `
        -Object $manifest `
        -Expected @('formatId', 'schemaVersion', 'closureId', 'dependencies') `
        -Context 'Notice manifest'
    if ([string](Get-RSNoticeProperty -Object $manifest -Name 'formatId') -cne
        $script:NoticeFormatId) {
        throw "Notice manifest formatId must be '$($script:NoticeFormatId)'."
    }
    if ([int](Get-RSNoticeProperty -Object $manifest -Name 'schemaVersion') -ne
        $script:NoticeSchemaVersion) {
        throw "Notice manifest schemaVersion must be $($script:NoticeSchemaVersion)."
    }
    $closureId = [string](Get-RSNoticeProperty -Object $manifest -Name 'closureId')
    Assert-RSNoticeIdentifier -Value $closureId -Name 'closureId'

    $dependencies = @(Get-RSNoticeProperty `
            -Object $manifest `
            -Name 'dependencies')
    if ($dependencies.Count -ne $requiredIds.Count) {
        throw (
            "Notice dependency closure has $($dependencies.Count) entries; " +
            "the required closure has $($requiredIds.Count).")
    }

    $materializedDependencies = [Collections.Generic.List[object]]::new()
    $noticeBuilder = [Text.StringBuilder]::new()
    $manifestFileSha256 = Get-RSNoticeBytesSha256 -Bytes $manifestBytes
    $manifestJcsSha256 = Get-RSJcsSha256 -InputObject $manifest
    Add-RSNoticeLine `
        -Builder $noticeBuilder `
        -Text 'RED SALAMANDER TERMINAL ENGINE THIRD-PARTY NOTICES'
    Add-RSNoticeLine -Builder $noticeBuilder
    Add-RSNoticeLine `
        -Builder $noticeBuilder `
        -Text "Closure: $closureId"
    Add-RSNoticeLine `
        -Builder $noticeBuilder `
        -Text "Manifest JCS SHA-256: $manifestJcsSha256"
    Add-RSNoticeLine `
        -Builder $noticeBuilder `
        -Text 'Generated deterministically from exact hash-verified UTF-8 inputs.'

    for ($dependencyIndex = 0;
        $dependencyIndex -lt $dependencies.Count;
        ++$dependencyIndex) {
        $dependency = $dependencies[$dependencyIndex]
        if ($dependency -isnot [Collections.IDictionary]) {
            throw "dependencies[$dependencyIndex] must be an object."
        }
        Assert-RSNoticeExactProperties `
            -Object $dependency `
            -Expected @('id', 'pin', 'spdx', 'inputs') `
            -Context "dependencies[$dependencyIndex]"
        $dependencyId = [string](Get-RSNoticeProperty `
                -Object $dependency `
                -Name 'id')
        Assert-RSNoticeIdentifier `
            -Value $dependencyId `
            -Name "dependencies[$dependencyIndex].id"
        if ($dependencyId -cne $requiredIds[$dependencyIndex]) {
            throw (
                "Notice dependency '$dependencyId' does not match required " +
                "ordinal closure entry '$($requiredIds[$dependencyIndex])'.")
        }
        $pin = [string](Get-RSNoticeProperty -Object $dependency -Name 'pin')
        $spdx = [string](Get-RSNoticeProperty -Object $dependency -Name 'spdx')
        if ([string]::IsNullOrWhiteSpace($pin) -or
            [string]::IsNullOrWhiteSpace($spdx)) {
            throw "Dependency '$dependencyId' must declare nonempty pin and spdx."
        }

        $inputs = @(Get-RSNoticeProperty -Object $dependency -Name 'inputs')
        if ($inputs.Count -eq 0) {
            throw "Dependency '$dependencyId' must have at least one license input."
        }
        $materializedInputs = [Collections.Generic.List[object]]::new()
        $inputContents = [Collections.Generic.List[object]]::new()
        $lastInputKey = $null
        $licenseCount = 0
        for ($inputIndex = 0; $inputIndex -lt $inputs.Count; ++$inputIndex) {
            $inputRecord = $inputs[$inputIndex]
            if ($inputRecord -isnot [Collections.IDictionary]) {
                throw "Dependency '$dependencyId' input $inputIndex must be an object."
            }
            Assert-RSNoticeExactProperties `
                -Object $inputRecord `
                -Expected @(
                    'id',
                    'kind',
                    'path',
                    'sha256',
                    'sizeBytes',
                    'sourceUrl',
                    'retrievedUtc',
                    'sourceSha256',
                    'sourceMember') `
                -Context "Dependency '$dependencyId' input $inputIndex"
            $inputId = [string](Get-RSNoticeProperty `
                    -Object $inputRecord `
                    -Name 'id')
            $kind = [string](Get-RSNoticeProperty `
                    -Object $inputRecord `
                    -Name 'kind')
            Assert-RSNoticeIdentifier `
                -Value $inputId `
                -Name "Dependency '$dependencyId' input id"
            if ($kind -cnotin @('license', 'notice')) {
                throw "Dependency '$dependencyId' input '$inputId' has unsupported kind '$kind'."
            }
            if ($kind -ceq 'license') {
                ++$licenseCount
            }
            $inputKey = "$kind`0$inputId"
            if ($null -ne $lastInputKey -and
                [StringComparer]::Ordinal.Compare(
                    $lastInputKey,
                    $inputKey) -ge 0) {
                throw (
                    "Dependency '$dependencyId' inputs must be unique and " +
                    'sorted ordinally by kind then id.')
            }
            $lastInputKey = $inputKey

            $relativePath = [string](Get-RSNoticeProperty `
                    -Object $inputRecord `
                    -Name 'path')
            if ($relativePath.Contains('\') -or
                [IO.Path]::IsPathFullyQualified($relativePath)) {
                throw "Dependency '$dependencyId' input '$inputId' path must be repository-relative with '/' separators."
            }
            $inputPath = Resolve-RSNoticeRepositoryPath `
                -RepoRoot $canonicalRoot `
                -Path $relativePath `
                -MustExist
            if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) {
                throw "Dependency '$dependencyId' input '$inputId' must be a file."
            }
            if (-not $reservedPaths.Add($inputPath)) {
                throw "Notice input path is duplicated or aliases an output: $relativePath"
            }
            $expectedSha256 = [string](Get-RSNoticeProperty `
                    -Object $inputRecord `
                    -Name 'sha256')
            Assert-RSNoticeSha256 `
                -Value $expectedSha256 `
                -Name "Dependency '$dependencyId' input '$inputId' sha256"
            Assert-RSNoticeProvenance `
                -InputRecord $inputRecord `
                -Context "Dependency '$dependencyId' input '$inputId'"
            $expectedSize = [string](Get-RSNoticeProperty `
                    -Object $inputRecord `
                    -Name 'sizeBytes')
            if ($expectedSize -cnotmatch '^(0|[1-9][0-9]*)$') {
                throw "Dependency '$dependencyId' input '$inputId' sizeBytes must be an unsigned decimal string."
            }
            $bytes = [IO.File]::ReadAllBytes($inputPath)
            if ([uint64]$expectedSize -ne [uint64]$bytes.Length) {
                throw "Dependency '$dependencyId' input '$inputId' size differs from the manifest."
            }
            $actualSha256 = Get-RSNoticeBytesSha256 -Bytes $bytes
            if ($actualSha256 -cne $expectedSha256) {
                throw "Dependency '$dependencyId' input '$inputId' SHA-256 differs from the manifest."
            }
            $text = Read-RSNoticeUtf8Text -Bytes $bytes -InputId $inputId
            $materializedInputs.Add([ordered]@{
                    inputId = $inputId
                    kind = $kind
                    sha256 = $actualSha256
                    sizeBytes = $expectedSize
                    sourceSha256 = [string]$inputRecord['sourceSha256']
                })
            $inputContents.Add([pscustomobject]@{
                    inputId = $inputId
                    kind = $kind
                    sha256 = $actualSha256
                    text = $text
                })
        }
        if ($licenseCount -eq 0) {
            throw "Dependency '$dependencyId' has no input with kind 'license'."
        }

        Add-RSNoticeLine -Builder $noticeBuilder
        Add-RSNoticeLine `
            -Builder $noticeBuilder `
            -Text ('=' * 79)
        Add-RSNoticeLine `
            -Builder $noticeBuilder `
            -Text "Dependency: $dependencyId"
        Add-RSNoticeLine -Builder $noticeBuilder -Text "Pin: $pin"
        Add-RSNoticeLine -Builder $noticeBuilder -Text "SPDX: $spdx"
        foreach ($inputContent in $inputContents) {
            Add-RSNoticeLine -Builder $noticeBuilder
            Add-RSNoticeLine `
                -Builder $noticeBuilder `
                -Text "[$($inputContent.kind):$($inputContent.inputId)]"
            Add-RSNoticeLine `
                -Builder $noticeBuilder `
                -Text "SHA-256: $($inputContent.sha256)"
            Add-RSNoticeLine -Builder $noticeBuilder
            Add-RSNoticeLine `
                -Builder $noticeBuilder `
                -Text $inputContent.text
        }
        $materializedDependencies.Add([ordered]@{
                dependencyId = $dependencyId
                pin = $pin
                spdx = $spdx
                inputs = [object[]]$materializedInputs.ToArray()
            })
    }
    Add-RSNoticeLine -Builder $noticeBuilder

    $noticeBytes = [Text.UTF8Encoding]::new($false, $true).GetBytes(
        $noticeBuilder.ToString())
    $noticeSha256 = Get-RSNoticeBytesSha256 -Bytes $noticeBytes
    $receipt = [ordered]@{
        formatId = $script:NoticeReceiptFormatId
        schemaVersion = $script:NoticeSchemaVersion
        closureId = $closureId
        manifestFileSha256 = $manifestFileSha256
        manifestJcsSha256 = $manifestJcsSha256
        requiredDependencyIds = [object[]]$requiredIds
        dependencyCount = $requiredIds.Count.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        dependencies = [object[]]$materializedDependencies.ToArray()
        notice = [ordered]@{
            sha256 = $noticeSha256
            sizeBytes = $noticeBytes.Length.ToString(
                [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    $receiptBytes = ConvertTo-RSJcsUtf8Bytes -InputObject $receipt

    Write-RSNoticeFile -Path $outputPath -Bytes $noticeBytes
    Write-RSNoticeFile -Path $receiptPath -Bytes $receiptBytes

    return [pscustomobject]@{
        OutputPath = $outputPath
        OutputSha256 = $noticeSha256
        OutputSizeBytes = [uint64]$noticeBytes.Length
        ReceiptPath = $receiptPath
        ReceiptSha256 = Get-RSNoticeBytesSha256 -Bytes $receiptBytes
        ReceiptSizeBytes = [uint64]$receiptBytes.Length
        ManifestFileSha256 = $manifestFileSha256
        ManifestJcsSha256 = $manifestJcsSha256
    }
}

Export-ModuleMember -Function Invoke-RSTerminalNoticeMaterialization
