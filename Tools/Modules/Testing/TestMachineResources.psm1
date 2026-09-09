Set-StrictMode -Version Latest

$script:RSTestMachineResourceSchema = 'red-salamander.test-machine-resources.v1'
$script:RSManagedResourceEnvironmentNames = @(
    'REDSALAMANDER_SELFTEST_CONN_FTP',
    'REDSALAMANDER_SELFTEST_CONN_SFTP',
    'REDSALAMANDER_SELFTEST_CONN_SCP',
    'REDSALAMANDER_SELFTEST_CONN_IMAP',
    'REDSALAMANDER_SELFTEST_CONN_S3',
    'REDSALAMANDER_SELFTEST_CONN_S3_ALT',
    'REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL',
    'REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS',
    'REDSALAMANDER_SELFTEST_CONN_SHAREPOINT',
    'REDSALAMANDER_SELFTEST_SMB_ROOT',
    'REDSALAMANDER_SELFTEST_MTP_DEVICE',
    'REDSALAMANDER_SELFTEST_MTP_PROFILE',
    'REDSALAMANDER_SELFTEST_MTP_ROOT',
    'REDSALAMANDER_SELFTEST_MTP_SCRATCH',
    'REDSALAMANDER_TEST_ALTERNATE_ROOT'
)
$script:RSConnectionProfileEnvironmentById = @{
    'connection.ftp.primary' = 'REDSALAMANDER_SELFTEST_CONN_FTP'
    'connection.sftp.primary' = 'REDSALAMANDER_SELFTEST_CONN_SFTP'
    'connection.scp.primary' = 'REDSALAMANDER_SELFTEST_CONN_SCP'
    'connection.imap.primary' = 'REDSALAMANDER_SELFTEST_CONN_IMAP'
    'connection.s3.primary' = 'REDSALAMANDER_SELFTEST_CONN_S3'
    'connection.s3.alternate' = 'REDSALAMANDER_SELFTEST_CONN_S3_ALT'
    'connection.onedrive.personal' = 'REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_PERSONAL'
    'connection.onedrive.business' = 'REDSALAMANDER_SELFTEST_CONN_ONEDRIVE_BUSINESS'
    'connection.sharepoint.primary' = 'REDSALAMANDER_SELFTEST_CONN_SHAREPOINT'
}

function Get-RSTestMachineResourceManifestPath {
    param([Parameter(Mandatory = $true)][string]$TestRoot)

    return [System.IO.Path]::GetFullPath((Join-Path $TestRoot 'config\machine-resources.json'))
}

function Test-RSTestMachineResourceText {
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string]$Description
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value.Length -gt 1024 -or $Value.IndexOf([char]0) -ge 0 -or $Value -match '[\r\n]') {
        throw "$Description must be non-empty, at most 1024 characters, and contain no NUL or newline."
    }
    return $Value
}

function Assert-RSSmbSelfTestRoot {
    param([Parameter(Mandatory = $true)][string]$Path)

    $root = Test-RSTestMachineResourceText -Value $Path -Description 'SMB root'
    if (-not $root.StartsWith('\\')) {
        throw "SMB root must be an absolute UNC path: $root"
    }

    $segments = @($root.Trim('\').Split('\', [System.StringSplitOptions]::RemoveEmptyEntries))
    if ($segments.Count -lt 3 -or @($segments | Where-Object { $_ -in @('.', '..') }).Count -gt 0) {
        throw "SMB root must name a directory below a share and must not contain '.' or '..': $root"
    }
    if (@($segments | Where-Object { $_ -match '(?i)selftest' }).Count -eq 0) {
        throw "SMB root must include a path segment containing 'selftest': $root"
    }
    return $root.TrimEnd('\')
}

function Assert-RSAlternateLocalTestRoot {
    param([Parameter(Mandatory = $true)][string]$Path)

    $root = Test-RSTestMachineResourceText -Value $Path -Description 'Alternate local test root'
    if ($root -notmatch '^[A-Za-z]:\\RedSalamander\.Perf$') {
        throw "Alternate local test root must be exact X:\RedSalamander.Perf: $root"
    }
    return [System.IO.Path]::GetFullPath($root)
}

function Get-RSTestMachineResourceDigest {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Resources)

    $projection = @($Resources | Sort-Object id | ForEach-Object {
            $row = [ordered]@{
                id = [string]$_.id
                kind = [string]$_.kind
                enabled = [bool]$_.enabled
            }
            if ([bool]$_.enabled) {
                foreach ($propertyName in @('profile', 'root', 'device', 'scratch')) {
                    $property = $_.PSObject.Properties[$propertyName]
                    if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $row[$propertyName] = [string]$property.Value
                    }
                }
            }
            [pscustomobject]$row
        })
    $json = ConvertTo-Json -InputObject @($projection) -Depth 8 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString($sha256.ComputeHash($bytes)).ToLowerInvariant()
    } finally {
        $sha256.Dispose()
    }
}

function Read-RSTestMachineResources {
    param(
        [Parameter(Mandatory = $true)][string]$TestRoot,
        [Parameter(Mandatory = $true)][string]$SchemaPath
    )

    $manifestPath = Get-RSTestMachineResourceManifestPath -TestRoot $TestRoot
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        return [pscustomobject]@{
            Present = $false
            Path = $manifestPath
            Resources = @()
            Environment = [ordered]@{}
            ManagedEnvironmentNames = @()
            CapabilityDigest = Get-RSTestMachineResourceDigest -Resources @()
        }
    }

    $manifestItem = Get-Item -LiteralPath $manifestPath -Force -ErrorAction Stop
    if (($manifestItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Test machine resource manifest must not be a reparse point: $manifestPath"
    }

    $json = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop
    if (-not (Test-Json -Json $json -SchemaFile $SchemaPath -ErrorAction SilentlyContinue)) {
        throw "Test machine resource manifest does not match the schema: $manifestPath"
    }
    $manifest = $json | ConvertFrom-Json -Depth 20
    if ([string]$manifest.schema -cne $script:RSTestMachineResourceSchema) {
        throw "Unsupported test machine resource schema '$($manifest.schema)'."
    }

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $environment = [ordered]@{}
    foreach ($resource in @($manifest.resources)) {
        $id = [string]$resource.id
        $kind = [string]$resource.kind
        if (-not $ids.Add($id)) {
            throw "Duplicate test machine resource id '$id'."
        }
        if (-not [bool]$resource.enabled) {
            continue
        }

        switch ($kind) {
            'connection-profile' {
                if (-not $script:RSConnectionProfileEnvironmentById.ContainsKey($id)) {
                    throw "Connection-profile resource id '$id' is not supported."
                }
                $profile = Test-RSTestMachineResourceText -Value ([string]$resource.profile) -Description "Profile for '$id'"
                $environment[$script:RSConnectionProfileEnvironmentById[$id]] = $profile
            }
            'smb-share' {
                if ($id -cne 'path.smb.primary') {
                    throw "SMB resource id must be 'path.smb.primary'."
                }
                $environment['REDSALAMANDER_SELFTEST_SMB_ROOT'] = Assert-RSSmbSelfTestRoot -Path ([string]$resource.root)
            }
            'mtp-device' {
                if ($id -cne 'device.mtp.primary') {
                    throw "MTP resource id must be 'device.mtp.primary'."
                }
                foreach ($mapping in @(
                        @{ Property = 'device'; Environment = 'REDSALAMANDER_SELFTEST_MTP_DEVICE' },
                        @{ Property = 'profile'; Environment = 'REDSALAMANDER_SELFTEST_MTP_PROFILE' },
                        @{ Property = 'root'; Environment = 'REDSALAMANDER_SELFTEST_MTP_ROOT' },
                        @{ Property = 'scratch'; Environment = 'REDSALAMANDER_SELFTEST_MTP_SCRATCH' }
                    )) {
                    $property = $resource.PSObject.Properties[$mapping.Property]
                    if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $environment[$mapping.Environment] = Test-RSTestMachineResourceText `
                            -Value ([string]$property.Value) -Description "MTP $($mapping.Property)"
                    }
                }
            }
            'local-test-root' {
                if ($id -cne 'path.local.alternate') {
                    throw "Local test-root resource id must be 'path.local.alternate'."
                }
                $environment['REDSALAMANDER_TEST_ALTERNATE_ROOT'] =
                    Assert-RSAlternateLocalTestRoot -Path ([string]$resource.root)
            }
            default { throw "Unsupported test machine resource kind '$kind'." }
        }
    }

    return [pscustomobject]@{
        Present = $true
        Path = $manifestPath
        Resources = @($manifest.resources)
        Environment = $environment
        ManagedEnvironmentNames = @($script:RSManagedResourceEnvironmentNames)
        CapabilityDigest = Get-RSTestMachineResourceDigest -Resources @($manifest.resources)
    }
}

function Set-RSTestMachineResourceEnvironment {
    param([Parameter(Mandatory = $true)][object]$Configuration)

    $names = if ([bool]$Configuration.Present) {
        @($Configuration.ManagedEnvironmentNames)
    } else {
        @($Configuration.Environment.Keys)
    }
    $previous = [ordered]@{}
    foreach ($name in $names) {
        $previous[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        [Environment]::SetEnvironmentVariable($name, $null, 'Process')
    }
    foreach ($entry in $Configuration.Environment.GetEnumerator()) {
        [Environment]::SetEnvironmentVariable([string]$entry.Key, [string]$entry.Value, 'Process')
    }

    return [pscustomobject]@{ Previous = $previous }
}

function Restore-RSTestMachineResourceEnvironment {
    param([AllowNull()][object]$Lease)

    if ($null -eq $Lease) {
        return
    }
    foreach ($entry in $Lease.Previous.GetEnumerator()) {
        [Environment]::SetEnvironmentVariable([string]$entry.Key, $entry.Value, 'Process')
    }
}

Export-ModuleMember -Function @(
    'Get-RSTestMachineResourceManifestPath',
    'Read-RSTestMachineResources',
    'Set-RSTestMachineResourceEnvironment',
    'Restore-RSTestMachineResourceEnvironment'
)
