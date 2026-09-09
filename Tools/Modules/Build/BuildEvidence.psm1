Set-StrictMode -Version Latest

$validationModule = Join-Path (Split-Path -Parent $PSScriptRoot) 'Testing\ValidationFingerprint.psm1'
$runtimeDependencyModule = Join-Path (Split-Path -Parent $PSScriptRoot) 'Packaging\RuntimeDependencies.psm1'
$artifactOperationModule = Join-Path $PSScriptRoot 'ArtifactOperationLock.psm1'
# Dependency imports must not use -Force: this module is re-imported by build.ps1
# inside Run-AllTests, and forcing a nested dependency would unload the runner's
# own ValidationFingerprint command surface from its parent scope.
Import-Module $validationModule -Scope Local -ErrorAction Stop
Import-Module $runtimeDependencyModule -Scope Local -ErrorAction Stop
Import-Module $artifactOperationModule -Scope Local -ErrorAction Stop

function Get-RSFileSha256 {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)

    $stream = [IO.File]::Open($LiteralPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        return ([Convert]::ToHexString($sha256.ComputeHash($stream))).ToLowerInvariant()
    } finally {
        $sha256.Dispose()
        $stream.Dispose()
    }
}

function Get-RSBytesSha256 {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes)

    return ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes))).ToLowerInvariant()
}

function ConvertTo-RSBuildRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$LiteralPath
    )

    $root = [IO.Path]::GetFullPath($RepoRoot)
    $path = [IO.Path]::GetFullPath($LiteralPath)
    $relative = [IO.Path]::GetRelativePath($root, $path).Replace('\', '/')
    if (-not (Test-RSValidationRelativePath -Path $relative)) {
        throw "Build artifact path must remain repository-relative: $path"
    }
    return $relative
}

function Resolve-RSBuildArtifactPath {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath
    )

    if (-not (Test-RSValidationRelativePath -Path $RelativePath)) {
        throw "Build artifact path is not a safe repository-relative path: $RelativePath"
    }
    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $path = [IO.Path]::GetFullPath((Join-Path $root $RelativePath))
    if (-not $path.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Build artifact path escaped the repository: $RelativePath"
    }
    $current = Get-Item -LiteralPath $root -Force -ErrorAction Stop
    if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Build artifact repository root must not be a reparse point: $root"
    }
    foreach ($segment in [IO.Path]::GetRelativePath($root, $path).Split([IO.Path]::DirectorySeparatorChar)) {
        $current = Get-Item -LiteralPath (Join-Path $current.FullName $segment) -Force -ErrorAction Stop
        if (($current.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Build artifact path contains a reparse point: $($current.FullName)"
        }
    }
    return $path
}

function Format-RSBuildDuration {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][TimeSpan]$Duration)

    $totalSeconds = [Math]::Max(0, [Math]::Floor($Duration.TotalSeconds))
    $hours = [Math]::Floor($totalSeconds / 3600)
    $minutes = [Math]::Floor(($totalSeconds % 3600) / 60)
    $seconds = $totalSeconds % 60
    if ($hours -gt 0) {
        return '{0}:{1:00}:{2:00}' -f $hours, $minutes, $seconds
    }
    return '{0:00}:{1:00}' -f $minutes, $seconds
}

function Format-RSBuildSizeMiB {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateRange(0, [long]::MaxValue)][long]$SizeBytes)

    return ([double]$SizeBytes / 1MB).ToString('0.00', [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-RSPowerShellDoubleQuotedLiteral {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)

    return '"' + $Value.Replace('`', '``').Replace('$', '`$').Replace('"', '`"') + '"'
}

function ConvertTo-RSBuildOutputLine {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object]$Artifact)

    $literal = ConvertTo-RSPowerShellDoubleQuotedLiteral -Value ([string]$Artifact.absolute_path_diagnostic)
    $size = Format-RSBuildSizeMiB -SizeBytes ([long]$Artifact.size_bytes)
    if ([string]$Artifact.role -eq 'launchable') {
        return "& $literal # $size MiB"
    }
    return "Output: $literal # $size MiB"
}

function Get-RSBuildArtifactRole {
    param([Parameter(Mandatory = $true)][IO.FileInfo]$File)

    $extension = $File.Extension.ToLowerInvariant()
    if ($extension -eq '.exe' -and $File.BaseName -in @('RedLauncher', 'RedSalamander', 'RedSalamanderMonitor')) { return 'launchable' }
    if ($extension -eq '.exe') { return 'test' }
    if ($extension -eq '.pdb') { return 'symbols' }
    if ($extension -in @('.msix', '.msi', '.zip')) { return 'package' }
    if ($extension -eq '.dll' -and $File.Directory.Name -eq 'Plugins') { return 'plugin' }
    if ($extension -eq '.dll') { return 'library' }
    return 'runtime-data'
}

function Get-RSBuildArtifactId {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $digest = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ relative_path = $RelativePath.ToLowerInvariant() })
    return 'artifact.' + $digest.Substring(0, 24)
}

function New-RSBuildArtifactRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [string[]]$RuntimeClosure = @()
    )

    $file = Get-Item -LiteralPath $LiteralPath -ErrorAction Stop
    if ($file.PSIsContainer) { throw "Build artifact must be a file: $LiteralPath" }
    if (($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Build artifact must not be a reparse point: $LiteralPath"
    }
    $relative = ConvertTo-RSBuildRelativePath -RepoRoot $RepoRoot -LiteralPath $file.FullName
    [pscustomobject][ordered]@{
        id = Get-RSBuildArtifactId -RelativePath $relative
        role = Get-RSBuildArtifactRole -File $file
        relative_path = $relative
        absolute_path_diagnostic = $file.FullName
        size_bytes = [long]$file.Length
        sha256 = Get-RSFileSha256 -LiteralPath $file.FullName
        runtime_closure = @($RuntimeClosure | Sort-Object -Unique)
    }
}

function Get-RSBuildManifestScalar {
    param(
        [Parameter(Mandatory = $true)][string[]]$Line,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$ManifestPath
    )

    $prefix = $Name + '='
    $values = @($Line | Where-Object { $_.StartsWith($prefix, [StringComparison]::Ordinal) } |
        ForEach-Object { $_.Substring($prefix.Length) })
    if ($values.Count -ne 1 -or [string]::IsNullOrWhiteSpace($values[0])) {
        throw "Build artifact manifest must contain exactly one non-empty '$Name' row: $ManifestPath"
    }
    return [string]$values[0]
}

function Invoke-RSNativeTextOutput {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [Parameter(Mandatory = $true)][string[]]$Argument,
        [Parameter(Mandatory = $true)][string]$WorkingDirectory
    )

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $FilePath
    $start.WorkingDirectory = $WorkingDirectory
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    foreach ($value in $Argument) { [void]$start.ArgumentList.Add($value) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (-not $process.Start()) { throw "Unable to start native identity probe: $FilePath" }
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $process.WaitForExit()
        $stdout = $stdoutTask.GetAwaiter().GetResult()
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) {
            throw "Native identity probe failed with exit $($process.ExitCode): $stderr"
        }
        return $stdout
    }
    finally {
        $process.Dispose()
    }
}

function New-RSBuildToolFileIdentity {
    param(
        [Parameter(Mandatory = $true)][string]$Role,
        [Parameter(Mandatory = $true)][string]$LiteralPath
    )

    $file = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Resolved toolchain input must be a regular non-reparse file: $LiteralPath"
    }
    $version = [string]$file.VersionInfo.FileVersion
    if ([string]::IsNullOrWhiteSpace($version)) { $version = 'unknown' }
    return [pscustomobject][ordered]@{
        role = $Role
        file_name = $file.Name
        file_version = $version
        size_bytes = [int64]$file.Length
        sha256 = Get-RSFileSha256 -LiteralPath $file.FullName
    }
}

function New-RSBuildSdkFileIdentity {
    param(
        [Parameter(Mandatory = $true)][string]$Role,
        [Parameter(Mandatory = $true)][string]$LiteralPath
    )

    $file = Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop
    if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Resolved SDK input must be a regular non-reparse file: $LiteralPath"
    }
    return [pscustomobject][ordered]@{
        role = $Role
        file_name = $file.Name
        size_bytes = [int64]$file.Length
        sha256 = Get-RSFileSha256 -LiteralPath $file.FullName
    }
}

function New-RSBuildToolchainIdentityRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [Parameter(Mandatory = $true)][string]$VcToolsVersion,
        [Parameter(Mandatory = $true)][string]$WindowsSdkVersion,
        [Parameter(Mandatory = $true)][string]$MSBuildPath,
        [Parameter(Mandatory = $true)][string]$CompilerPath,
        [Parameter(Mandatory = $true)][string]$LinkerPath,
        [Parameter(Mandatory = $true)][string]$ResourceCompilerPath,
        [Parameter(Mandatory = $true)][string]$WindowsHeaderPath,
        [Parameter(Mandatory = $true)][string]$UmImportLibraryPath,
        [Parameter(Mandatory = $true)][string]$UcrtImportLibraryPath
    )

    return [pscustomobject][ordered]@{
        schema = 'red-salamander.toolchain-identity.v1'
        canonicalization = 'rs-canonical-json-v1'
        platform = $Platform
        vc_tools_version = $VcToolsVersion.Trim().TrimEnd('\')
        windows_sdk_version = $WindowsSdkVersion.Trim().TrimEnd('\')
        msbuild = New-RSBuildToolFileIdentity -Role msbuild -LiteralPath $MSBuildPath
        compiler = New-RSBuildToolFileIdentity -Role compiler -LiteralPath $CompilerPath
        linker = New-RSBuildToolFileIdentity -Role linker -LiteralPath $LinkerPath
        resource_compiler = New-RSBuildToolFileIdentity -Role resource-compiler -LiteralPath $ResourceCompilerPath
        sdk_files = @(
            New-RSBuildSdkFileIdentity -Role windows-header -LiteralPath $WindowsHeaderPath
            New-RSBuildSdkFileIdentity -Role um-import-library -LiteralPath $UmImportLibraryPath
            New-RSBuildSdkFileIdentity -Role ucrt-import-library -LiteralPath $UcrtImportLibraryPath
        )
    }
}

function Get-RSBuildToolchainIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$MSBuildPath,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration
    )

    $root = [IO.Path]::GetFullPath($RepoRoot)
    $probeProject = Join-Path $root 'RedSalamander\RedSalamander.vcxproj'
    $propertyNames = 'VCToolsInstallDir,VCToolsVersion,WindowsSdkDir,WindowsTargetPlatformVersion'
    $jsonText = Invoke-RSNativeTextOutput -FilePath $MSBuildPath -WorkingDirectory $root -Argument @(
        $probeProject,
        "/p:Configuration=$Configuration",
        "/p:Platform=$Platform",
        "-getProperty:$propertyNames",
        '-nologo'
    )
    $resolved = $jsonText | ConvertFrom-Json -DateKind String
    $properties = $resolved.Properties
    foreach ($name in $propertyNames.Split(',')) {
        if ([string]::IsNullOrWhiteSpace([string]$properties.$name)) {
            throw "MSBuild did not resolve required toolchain property '$name'."
        }
    }

    $target = $Platform.ToLowerInvariant()
    $vcToolsRoot = [IO.Path]::GetFullPath([string]$properties.VCToolsInstallDir)
    $sdkRoot = [IO.Path]::GetFullPath([string]$properties.WindowsSdkDir)
    $sdkVersion = ([string]$properties.WindowsTargetPlatformVersion).Trim().TrimEnd('\')
    $record = New-RSBuildToolchainIdentityRecord -Platform $Platform `
        -VcToolsVersion ([string]$properties.VCToolsVersion) -WindowsSdkVersion $sdkVersion `
        -MSBuildPath $MSBuildPath `
        -CompilerPath (Join-Path $vcToolsRoot "bin\Hostx64\$target\cl.exe") `
        -LinkerPath (Join-Path $vcToolsRoot "bin\Hostx64\$target\link.exe") `
        -ResourceCompilerPath (Join-Path $sdkRoot "bin\$sdkVersion\x64\rc.exe") `
        -WindowsHeaderPath (Join-Path $sdkRoot "Include\$sdkVersion\um\Windows.h") `
        -UmImportLibraryPath (Join-Path $sdkRoot "Lib\$sdkVersion\um\$target\kernel32.lib") `
        -UcrtImportLibraryPath (Join-Path $sdkRoot "Lib\$sdkVersion\ucrt\$target\ucrt.lib")
    return [pscustomobject]@{
        Record = $record
        Digest = Get-RSCanonicalJsonDigest -InputObject $record
    }
}

function Publish-RSBuildToolchainIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][object]$IdentityRecord,
        [Parameter(Mandatory = $true)][string]$SchemaPath
    )

    [void][IO.Directory]::CreateDirectory([IO.Path]::GetFullPath($BuildOutputDir))
    $path = Join-Path $BuildOutputDir 'toolchain-identity.json'
    Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $IdentityRecord `
        -SchemaPath $SchemaPath -AllowedRoot $BuildOutputDir
    return $path
}

function Resolve-RSDeclaredBuildOutputPath {
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $output = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $path = [IO.Path]::GetFullPath($LiteralPath)
    if (-not $path.StartsWith($output + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Description is outside the selected build output: $path"
    }
    $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$Description must be a regular non-reparse file: $path"
    }
    return $item.FullName
}

function Resolve-RSVcpkgRuntimeOutputPath {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$ProjectTargetPath,
        [Parameter(Mandatory = $true)][string]$LiteralPath,
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform
    )

    $description = "vcpkg runtime output for '$Project'"
    $outputRoot = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $path = [IO.Path]::GetFullPath($LiteralPath)
    if ($path.StartsWith($outputRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        return Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $outputRoot `
            -LiteralPath $path -Description $description
    }

    # Legacy applocal.ps1 writes destination paths, while current vcpkg z-applocal
    # writes source paths. Only canonical platform-scoped vcpkg sources may be
    # rebased to the directory containing the already-validated project target.
    $vcpkgRoot = [IO.Path]::GetFullPath(
        (Join-Path $RepoRoot ".build\vcpkg_installed\$Platform")).TrimEnd('\')
    if (-not $path.StartsWith($vcpkgRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "$description is outside the selected build output and canonical vcpkg install root: $path"
    }

    $source = Get-Item -LiteralPath $path -Force -ErrorAction Stop
    if ($source.PSIsContainer -or ($source.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "$description source must be a regular non-reparse file: $path"
    }

    $deployedPath = Join-Path (Split-Path -Parent $ProjectTargetPath) $source.Name
    return Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $outputRoot `
        -LiteralPath $deployedPath -Description "$description deployed from '$path'"
}

function Get-RSBuildArtifactManifestDeclarations {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$ArtifactManifestDir,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform
    )

    $manifestRoot = [IO.Path]::GetFullPath($ArtifactManifestDir)
    $manifestRootItem = Get-Item -LiteralPath $manifestRoot -Force -ErrorAction Stop
    if (-not $manifestRootItem.PSIsContainer -or
        ($manifestRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Build artifact manifest root must be a regular directory: $manifestRoot"
    }
    $manifestFiles = @(Get-ChildItem -LiteralPath $manifestRoot -File -Filter '*.manifest' |
        Sort-Object Name)
    if ($manifestFiles.Count -eq 0) { throw "Build produced no artifact manifests: $manifestRoot" }

    $runtimeManifest = Get-RSRuntimeDependencyManifest `
        -RepoRoot $RepoRoot -Configuration $Configuration -Platform $Platform
    $projects = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $declarations = foreach ($manifestFile in $manifestFiles) {
        if (($manifestFile.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Build artifact manifest must not be a reparse point: $($manifestFile.FullName)"
        }
        $lines = @(Get-Content -LiteralPath $manifestFile.FullName -Encoding UTF8 |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $project = Get-RSBuildManifestScalar -Line $lines -Name project -ManifestPath $manifestFile.FullName
        if (-not $projects.Add($project)) { throw "Duplicate build artifact project manifest: $project" }
        $testsText = Get-RSBuildManifestScalar -Line $lines -Name tests_enabled -ManifestPath $manifestFile.FullName
        if ($testsText -notin @('true', 'false')) {
            throw "Build artifact manifest has invalid tests_enabled value '$testsText': $($manifestFile.FullName)"
        }
        $target = Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $BuildOutputDir `
            -LiteralPath (Get-RSBuildManifestScalar -Line $lines -Name target -ManifestPath $manifestFile.FullName) `
            -Description "Declared target for '$project'"

        $runtime = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $runtimePrefix = 'runtime='
        foreach ($line in @($lines | Where-Object { $_.StartsWith($runtimePrefix, [StringComparison]::Ordinal) })) {
            $candidate = $line.Substring($runtimePrefix.Length)
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $resolved = Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $BuildOutputDir `
                -LiteralPath $candidate -Description "Declared runtime reference for '$project'"
            [void]$runtime.Add($resolved)
        }

        $tlogPath = Get-RSBuildManifestScalar -Line $lines -Name vcpkg_tlog -ManifestPath $manifestFile.FullName
        if (Test-Path -LiteralPath $tlogPath -PathType Leaf) {
            foreach ($candidate in @(Get-Content -LiteralPath $tlogPath -Encoding Unicode |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                $resolved = Resolve-RSVcpkgRuntimeOutputPath -RepoRoot $RepoRoot `
                    -BuildOutputDir $BuildOutputDir -ProjectTargetPath $target `
                    -LiteralPath $candidate -Project $project -Platform $Platform
                [void]$runtime.Add($resolved)
            }
        }

        foreach ($dependency in @($runtimeManifest.Dependencies | Where-Object { $_.Projects -contains $project })) {
            $destinationRoot = if ($dependency.OutputRoot -eq 'BuildOutput') {
                $BuildOutputDir
            } else {
                Join-Path $BuildOutputDir 'Plugins'
            }
            $candidate = Join-Path $destinationRoot $dependency.OutputName
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                $resolved = Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $BuildOutputDir `
                    -LiteralPath $candidate -Description "Manifest runtime output '$($dependency.Id)' for '$project'"
                [void]$runtime.Add($resolved)
            } elseif ($dependency.Required) {
                throw "Required manifest runtime output '$($dependency.Id)' is missing for '$project': $candidate"
            }
        }

        [pscustomobject]@{
            Project = $project
            Target = $target
            TestsEnabled = ($testsText -eq 'true')
            Runtime = @($runtime | Sort-Object)
            ManifestPath = $manifestFile.FullName
        }
    }
    return @($declarations)
}

function Test-RSBuildArtifactManifestTestsEnabled {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ArtifactManifestDir)

    $files = @(Get-ChildItem -LiteralPath $ArtifactManifestDir -File -Filter '*.manifest' -ErrorAction Stop)
    if ($files.Count -eq 0) { throw "Build produced no artifact manifests: $ArtifactManifestDir" }
    foreach ($file in $files) {
        $lines = @(Get-Content -LiteralPath $file.FullName -Encoding UTF8)
        if ((Get-RSBuildManifestScalar -Line $lines -Name tests_enabled -ManifestPath $file.FullName) -ne 'true') {
            return $false
        }
    }
    return $true
}

function Get-RSBuildArtifactRecords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$ArtifactManifestDir,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [AllowEmptyString()][string]$ProjectName = '',
        [AllowEmptyString()][string]$ResolvedProjectPath = ''
    )

    $output = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $declarations = @(Get-RSBuildArtifactManifestDeclarations -RepoRoot $RepoRoot `
        -BuildOutputDir $output -ArtifactManifestDir $ArtifactManifestDir `
        -Configuration $Configuration -Platform $Platform)
    if (-not [string]::IsNullOrWhiteSpace($ProjectName) -and
        @($declarations | Where-Object Project -eq $ProjectName).Count -eq 0) {
        throw "No output declaration was produced for selected project '$ProjectName'."
    }

    $extraByProject = @{}
    foreach ($declaration in $declarations) { $extraByProject[$declaration.Project] = @() }
    if ($extraByProject.ContainsKey('RedLauncher')) {
        $extraByProject['RedLauncher'] += Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $output `
            -LiteralPath (Join-Path $output 'red.exe') -Description 'RedLauncher command alias'
    }
    foreach ($appProject in @('RedSalamander', 'RedSalamanderMonitor')) {
        if (-not $extraByProject.ContainsKey($appProject)) { continue }
        $extraByProject[$appProject] += Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $output `
            -LiteralPath (Join-Path $output 'SettingsStore.schema.json') -Description "$appProject settings schema"
        foreach ($source in @(Get-ChildItem -LiteralPath (Join-Path $RepoRoot 'Specs\Themes') -File -Recurse |
                Where-Object { $_.Name -like '*.theme.json5' -or $_.Name -eq 'THIRD-PARTY-NOTICES.md' -or $_.Name -like '*.LICENSE.txt' })) {
            $themeRelative = [IO.Path]::GetRelativePath((Join-Path $RepoRoot 'Specs\Themes'), $source.FullName)
            $extraByProject[$appProject] += Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $output `
                -LiteralPath (Join-Path (Join-Path $output 'Themes') $themeRelative) -Description "$appProject theme runtime data"
        }
    }
    if ($extraByProject.ContainsKey('RedSalamander')) {
        $extraByProject['RedSalamander'] += Resolve-RSDeclaredBuildOutputPath -BuildOutputDir $output `
            -LiteralPath (Join-Path $output 'asan.supp') -Description 'RedSalamander ASan suppressions'
    }

    $targetByPath = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $allPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($declaration in $declarations) {
        if (-not $targetByPath.TryAdd($declaration.Target, $declaration)) {
            throw "Multiple project manifests declare the same target: $($declaration.Target)"
        }
        [void]$allPaths.Add($declaration.Target)
        foreach ($runtime in @($declaration.Runtime) + @($extraByProject[$declaration.Project])) {
            [void]$allPaths.Add($runtime)
        }
    }

    $redSalamander = @($declarations | Where-Object Project -eq 'RedSalamander' | Select-Object -First 1)
    if ($redSalamander.Count -eq 1) {
        foreach ($declaration in $declarations) {
            $relativeTarget = [IO.Path]::GetRelativePath($output, $declaration.Target).Replace('\', '/')
            if ($relativeTarget.StartsWith('Plugins/', [StringComparison]::OrdinalIgnoreCase) -or
                $relativeTarget.StartsWith('Lang/', [StringComparison]::OrdinalIgnoreCase)) {
                $extraByProject['RedSalamander'] += $declaration.Target
            }
        }
    }

    $closureByProject = @{}
    foreach ($declaration in $declarations) {
        $closure = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        [void]$closure.Add($declaration.Target)
        foreach ($runtime in @($declaration.Runtime) + @($extraByProject[$declaration.Project])) {
            [void]$closure.Add($runtime)
        }
        $closureByProject[$declaration.Project] = $closure
    }
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($declaration in $declarations) {
            $closure = $closureByProject[$declaration.Project]
            foreach ($path in @($closure)) {
                if (-not $targetByPath.ContainsKey($path)) { continue }
                $referenced = $targetByPath[$path]
                foreach ($transitive in @($closureByProject[$referenced.Project])) {
                    if ($closure.Add($transitive)) { $changed = $true }
                }
            }
        }
    }

    $declaredBinaryPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in $allPaths) {
        if ([IO.Path]::GetExtension($path) -in @('.exe', '.dll')) { [void]$declaredBinaryPaths.Add($path) }
    }
    $unknownBinaries = @(Get-ChildItem -LiteralPath $output -File -Recurse | Where-Object {
            $_.FullName -notlike "*\receipts\*" -and $_.Extension -in @('.exe', '.dll') -and
            -not $declaredBinaryPaths.Contains($_.FullName)
        } | Sort-Object FullName)
    if ([string]::IsNullOrWhiteSpace($ProjectName) -and $unknownBinaries.Count -gt 0) {
        throw "Ungoverned executable or DLL remains in the selected build output: $($unknownBinaries.FullName -join ', ')"
    }

    $records = foreach ($path in @($allPaths | Sort-Object)) {
        $recordClosure = @()
        if ($targetByPath.ContainsKey($path)) {
            $recordClosure = @($closureByProject[$targetByPath[$path].Project] |
                ForEach-Object { ConvertTo-RSBuildRelativePath -RepoRoot $RepoRoot -LiteralPath $_ } |
                Sort-Object -Unique)
        }
        $record = New-RSBuildArtifactRecord -RepoRoot $RepoRoot -LiteralPath $path -RuntimeClosure $recordClosure
        if (-not $targetByPath.ContainsKey($path) -and $record.role -eq 'plugin') { $record.role = 'library' }
        if (-not [string]::IsNullOrWhiteSpace($ProjectName) -and
            $targetByPath.ContainsKey($path) -and $targetByPath[$path].Project -eq $ProjectName -and
            [IO.Path]::GetExtension($path) -eq '.exe') {
            $record.role = 'launchable'
        }
        $record
    }
    return @($records | Sort-Object relative_path)
}

function Invoke-RSGitBinaryOutput {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$Argument
    )

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'git.exe'
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    foreach ($value in @('-C', $RepoRoot) + $Argument) { [void]$start.ArgumentList.Add($value) }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $memory = [IO.MemoryStream]::new()
    try {
        if (-not $process.Start()) { throw 'Unable to start Git.' }
        $process.StandardOutput.BaseStream.CopyTo($memory)
        $errorText = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "Git failed with exit $($process.ExitCode): $errorText" }
        return ,$memory.ToArray()
    } finally {
        $memory.Dispose()
        $process.Dispose()
    }
}

function Get-RSGitSourcePaths {
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    $bytes = Invoke-RSGitBinaryOutput -RepoRoot $RepoRoot -Argument @('ls-files', '-co', '--exclude-standard', '-z', '--', '.')
    $paths = @([Text.Encoding]::UTF8.GetString([byte[]]$bytes).Split([char]0, [StringSplitOptions]::RemoveEmptyEntries))
    [Array]::Sort($paths, [StringComparer]::Ordinal)
    Assert-RSValidationPathSet -Path $paths
    return $paths
}

function Get-RSGitPathList {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$Argument
    )

    $bytes = Invoke-RSGitBinaryOutput -RepoRoot $RepoRoot -Argument $Argument
    return @([Text.Encoding]::UTF8.GetString([byte[]]$bytes).Split([char]0, [StringSplitOptions]::RemoveEmptyEntries))
}

function Get-RSBuildSourceSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    $snapshot = Get-RSWorkspaceSnapshot -RepoRoot $RepoRoot
    [pscustomobject]@{
        GitHead = $snapshot.git_head
        SnapshotId = $snapshot.snapshot_id
        FileCount = @($snapshot.files).Count
        WorkspaceSnapshot = $snapshot
    }
}

function Get-RSFileSetDigest {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string[]]$RelativePath
    )

    $rows = foreach ($relative in @($RelativePath | Sort-Object -Unique)) {
        $path = Join-Path $RepoRoot $relative
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            [ordered]@{ path = $relative.Replace('\', '/'); sha256 = Get-RSFileSha256 -LiteralPath $path }
        }
    }
    return Get-RSCanonicalJsonDigest -InputObject @($rows)
}

function Get-RSGhosttyRuntimeIdentityInputPaths {
    return [string[]]@(
        'External\TerminalEngine\GhosttyRuntimeLock.v1.json',
        'Specs\Terminal\GhosttyRuntimeLock.schema.json',
        'Tools\Build-GhosttyTerminalRuntime.ps1',
        'Tools\Modules\Build\SanitizedEnvironment.psm1',
        'Tools\Modules\Testing\ValidationFingerprint.psm1',
        'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1',
        'Tools\TerminalEngine\Patches\Ghostty-WindowsPrivateRuntime.patch',
        'Tools\TerminalEngine\TerminalEngineArchive.psm1',
        'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1',
        'Tools\TerminalEngine\TerminalEngineTextIdentity.psm1',
        'Tools\TerminalEngine\TerminalRuntimeExportIdentity.psm1',
        'Tools\TerminalEngine\TerminalRuntimeImportPolicy.psm1',
        'External\TerminalEngine\Notices\uucode-bjoern-hoehrmann.txt',
        'External\TerminalEngine\Notices\unicode-license-3.0-2025.txt',
        'RuntimeDependencies.props')
}

function Get-RSGhosttyRuntimeIdentityDigest {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    return Get-RSFileSetDigest `
        -RepoRoot ([IO.Path]::GetFullPath($RepoRoot)) `
        -RelativePath (Get-RSGhosttyRuntimeIdentityInputPaths)
}

function Get-RSBuildIdentityBundle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$MSBuildPath,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration
    )

    $root = [IO.Path]::GetFullPath($RepoRoot)
    $toolchain = Get-RSBuildToolchainIdentity -RepoRoot $root -MSBuildPath $MSBuildPath `
        -Platform $Platform -Configuration $Configuration
    $sourcePaths = Get-RSGitSourcePaths -RepoRoot $root
    $projectExtensions = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($extension in @('.sln', '.vcxproj', '.props', '.targets')) { [void]$projectExtensions.Add($extension) }
    $projectFiles = @($sourcePaths | Where-Object {
            $projectExtensions.Contains([IO.Path]::GetExtension($_)) -and
            -not $_.StartsWith('Specs/Plans/Done/', [StringComparison]::OrdinalIgnoreCase)
        })
    $runtimeFiles = @('RuntimeDependencies.props', 'Directory.Build.targets') + @($sourcePaths | Where-Object {
            $_.StartsWith('Specs/Themes/', [StringComparison]::OrdinalIgnoreCase)
        })
    [pscustomobject]@{
        RepoIdentity = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ physical_root = $root.ToLowerInvariant() })
        ToolchainIdentityRecord = $toolchain.Record
        ToolchainIdentityDigest = $toolchain.Digest
        VcpkgIdentityDigest = Get-RSFileSetDigest -RepoRoot $root -RelativePath @('vcpkg.json', 'vcpkg-configuration.json', 'vcpkg-tool.json')
        GhosttyRuntimeIdentity = Get-RSGhosttyRuntimeIdentityDigest -RepoRoot $root
        ProjectGraphDigest = Get-RSFileSetDigest -RepoRoot $root -RelativePath $projectFiles
        RuntimeDependencyDigest = Get-RSFileSetDigest -RepoRoot $root -RelativePath $runtimeFiles
    }
}

function Get-RSBuildReceiptProjection {
    param([Parameter(Mandatory = $true)][object]$Receipt)

    [ordered]@{
        schema = $Receipt.schema
        canonicalization = $Receipt.canonicalization
        trust_scope = $Receipt.trust_scope
        repo_identity = $Receipt.repo_identity
        source_snapshot_id = $Receipt.source_snapshot_id
        git_head = $Receipt.git_head
        platform = $Receipt.platform
        configuration = $Receipt.configuration
        target = $Receipt.target
        project_selection = @($Receipt.project_selection)
        build_arguments = @($Receipt.build_arguments)
        tests_enabled = [bool]$Receipt.tests_enabled
        official_release = [bool]$Receipt.official_release
        toolchain_identity_digest = $Receipt.toolchain_identity_digest
        vcpkg_identity_digest = $Receipt.vcpkg_identity_digest
        ghostty_runtime_identity = $Receipt.ghostty_runtime_identity
        project_graph_digest = $Receipt.project_graph_digest
        runtime_dependency_digest = $Receipt.runtime_dependency_digest
        exit_code = [int64]$Receipt.exit_code
        outputs = @($Receipt.outputs | ForEach-Object {
                [ordered]@{
                    id = $_.id; role = $_.role; relative_path = $_.relative_path
                    size_bytes = [int64]$_.size_bytes; sha256 = $_.sha256
                    runtime_closure = @($_.runtime_closure)
                }
            })
    }
}

function New-RSBuildReceipt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$SourceSnapshot,
        [Parameter(Mandatory = $true)][object]$Identity,
        [Parameter(Mandatory = $true)][ValidateSet('x64', 'ARM64')][string]$Platform,
        [Parameter(Mandatory = $true)][ValidateSet('Debug', 'Release', 'ASan Debug')][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][string[]]$ProjectSelection,
        [string[]]$BuildArguments = @(),
        [Parameter(Mandatory = $true)][bool]$TestsEnabled,
        [Parameter(Mandatory = $true)][bool]$OfficialRelease,
        [Parameter(Mandatory = $true)][datetime]$StartedUtc,
        [Parameter(Mandatory = $true)][datetime]$EndedUtc,
        [Parameter(Mandatory = $true)][TimeSpan]$Duration,
        [Parameter(Mandatory = $true)][object[]]$Outputs
    )

    Assert-RSValidationIdentifier -Id $Target
    $receipt = [ordered]@{
        schema = 'red-salamander.build-receipt.v1'
        canonicalization = 'rs-canonical-json-v1'
        receipt_id = ''
        trust_scope = 'local-checkout'
        repo_identity = $Identity.RepoIdentity
        source_snapshot_id = $SourceSnapshot.SnapshotId
        git_head = $SourceSnapshot.GitHead
        platform = $Platform
        configuration = $Configuration
        target = $Target
        project_selection = @($ProjectSelection | ForEach-Object { $_.Replace('\', '/') })
        build_arguments = @($BuildArguments)
        tests_enabled = $TestsEnabled
        official_release = $OfficialRelease
        toolchain_identity_digest = $Identity.ToolchainIdentityDigest
        vcpkg_identity_digest = $Identity.VcpkgIdentityDigest
        ghostty_runtime_identity = $Identity.GhosttyRuntimeIdentity
        project_graph_digest = $Identity.ProjectGraphDigest
        runtime_dependency_digest = $Identity.RuntimeDependencyDigest
        started_utc = $StartedUtc.ToUniversalTime().ToString('o')
        ended_utc = $EndedUtc.ToUniversalTime().ToString('o')
        duration_ms = [int64][Math]::Max(0, [Math]::Round($Duration.TotalMilliseconds))
        exit_code = [int64]0
        outputs = @($Outputs)
    }
    $receipt.receipt_id = Get-RSCanonicalJsonDigest -InputObject (Get-RSBuildReceiptProjection -Receipt ([pscustomobject]$receipt))
    return [pscustomobject]$receipt
}

function Get-RSBuildReceiptStore {
    param([Parameter(Mandatory = $true)][string]$BuildOutputDir)

    $root = [IO.Path]::GetFullPath($BuildOutputDir)
    [pscustomobject]@{
        Root = $root
        ReceiptDirectory = Join-Path $root 'receipts'
        PointerPath = Join-Path $root 'build-receipt.json'
    }
}

function Assert-RSBuildReceiptSemantics {
    param([Parameter(Mandatory = $true)][object]$Receipt)

    $computedReceiptId = Get-RSCanonicalJsonDigest -InputObject (Get-RSBuildReceiptProjection -Receipt $Receipt)
    if ([string]$Receipt.receipt_id -cne $computedReceiptId) {
        throw "Build receipt projection digest does not match receipt_id '$($Receipt.receipt_id)'."
    }

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $paths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $outputsByPath = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($output in @($Receipt.outputs)) {
        $relative = [string]$output.relative_path
        if (-not $ids.Add([string]$output.id)) { throw "Duplicate build artifact ID: $($output.id)" }
        if (-not $paths.Add($relative)) { throw "Duplicate build artifact path: $relative" }
        if ([string]$output.id -cne (Get-RSBuildArtifactId -RelativePath $relative)) {
            throw "Build artifact ID does not match its governed relative path: $relative"
        }
        $closure = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($closurePath in @($output.runtime_closure)) {
            if (-not $closure.Add([string]$closurePath)) {
                throw "Duplicate runtime-closure path for '$relative': $closurePath"
            }
        }
        $outputsByPath.Add($relative, $output)
    }
    foreach ($output in @($Receipt.outputs)) {
        foreach ($closurePath in @($output.runtime_closure)) {
            if (-not $outputsByPath.ContainsKey([string]$closurePath)) {
                throw "Runtime closure for '$($output.relative_path)' references an unattested output: $closurePath"
            }
        }
        if (@($output.runtime_closure).Count -gt 0 -and
            -not (@($output.runtime_closure) -contains [string]$output.relative_path)) {
            throw "Runtime closure for '$($output.relative_path)' must include the artifact itself."
        }
    }
    return $true
}

function Publish-RSBuildReceipt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][object]$Receipt,
        [Parameter(Mandatory = $true)][string]$SchemaPath
    )

    $null = Assert-RSBuildReceiptSemantics -Receipt $Receipt
    $store = Get-RSBuildReceiptStore -BuildOutputDir $BuildOutputDir
    if (-not (Test-Path -LiteralPath $store.ReceiptDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $store.ReceiptDirectory)
    }
    $recordPath = Join-Path $store.ReceiptDirectory ($Receipt.receipt_id + '.json')
    if (-not (Test-Path -LiteralPath $recordPath -PathType Leaf)) {
        Write-RSAtomicCanonicalJsonFile -LiteralPath $recordPath -InputObject $Receipt -SchemaPath $SchemaPath -AllowedRoot $store.Root
    }
    $stored = Read-RSValidationJsonFile -LiteralPath $recordPath -SchemaPath $SchemaPath -AllowedRoot $store.Root
    $null = Assert-RSBuildReceiptSemantics -Receipt $stored
    if ((Get-RSCanonicalJsonDigest -InputObject $stored) -cne (Get-RSCanonicalJsonDigest -InputObject $Receipt)) {
        throw "Stored build receipt record does not exactly match the receipt being published: $recordPath"
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath $store.PointerPath -InputObject $stored -SchemaPath $SchemaPath -AllowedRoot $store.Root
    return $stored
}

function Read-RSBuildReceipt {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$SchemaPath
    )

    $store = Get-RSBuildReceiptStore -BuildOutputDir $BuildOutputDir
    if (-not (Test-Path -LiteralPath $store.PointerPath -PathType Leaf)) { return $null }
    $pointer = Read-RSValidationJsonFile -LiteralPath $store.PointerPath -SchemaPath $SchemaPath -AllowedRoot $store.Root
    $null = Assert-RSBuildReceiptSemantics -Receipt $pointer
    $recordPath = Join-Path $store.ReceiptDirectory ([string]$pointer.receipt_id + '.json')
    $record = Read-RSValidationJsonFile -LiteralPath $recordPath -SchemaPath $SchemaPath -AllowedRoot $store.Root
    $null = Assert-RSBuildReceiptSemantics -Receipt $record
    if ((Get-RSCanonicalJsonDigest -InputObject $pointer) -cne (Get-RSCanonicalJsonDigest -InputObject $record) -or
        (Get-RSFileSha256 -LiteralPath $store.PointerPath) -cne (Get-RSFileSha256 -LiteralPath $recordPath)) {
        throw "Build receipt pointer does not byte-for-byte and semantically match its immutable record: $recordPath"
    }
    return $record
}

function Suspend-RSBuildReceipt {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$BuildOutputDir)

    $store = Get-RSBuildReceiptStore -BuildOutputDir $BuildOutputDir
    if (Test-Path -LiteralPath $store.PointerPath -PathType Leaf) {
        Remove-Item -LiteralPath $store.PointerPath -Force
    }
}

function Test-RSBuildReceiptCompatible {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Receipt,
        [Parameter(Mandatory = $true)][object]$SourceSnapshot,
        [Parameter(Mandatory = $true)][object]$Identity,
        [Parameter(Mandatory = $true)][string]$Platform,
        [Parameter(Mandatory = $true)][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [string[]]$ProjectSelection = @(),
        [string[]]$BuildArguments = @(),
        [bool]$TestsEnabled = $true,
        [bool]$OfficialRelease = $false
    )

    if ($null -eq $Receipt) { return $false }
    try { $null = Assert-RSBuildReceiptSemantics -Receipt $Receipt } catch { return $false }
    if ($Receipt.source_snapshot_id -ne $SourceSnapshot.SnapshotId -or
        $Receipt.repo_identity -ne $Identity.RepoIdentity -or
        $Receipt.toolchain_identity_digest -ne $Identity.ToolchainIdentityDigest -or
        $Receipt.vcpkg_identity_digest -ne $Identity.VcpkgIdentityDigest -or
        $Receipt.project_graph_digest -ne $Identity.ProjectGraphDigest -or
        $Receipt.runtime_dependency_digest -ne $Identity.RuntimeDependencyDigest -or
        $Receipt.platform -ne $Platform -or $Receipt.configuration -ne $Configuration -or
        $Receipt.target -ne $Target -or
        [bool]$Receipt.tests_enabled -ne $TestsEnabled -or
        [bool]$Receipt.official_release -ne $OfficialRelease -or
        (@($Receipt.project_selection) -join "`n") -cne (@($ProjectSelection) -join "`n") -or
        (@($Receipt.build_arguments) -join "`n") -cne (@($BuildArguments) -join "`n")) { return $false }
    foreach ($artifact in @($Receipt.outputs)) {
        $path = Join-Path $RepoRoot ([string]$artifact.relative_path)
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $false }
        if ((Get-RSFileSha256 -LiteralPath $path) -ne $artifact.sha256) { return $false }
    }
    return $true
}

function Test-RSBuildReceiptForArtifactUse {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$Receipt,
        [Parameter(Mandatory = $true)][object]$SourceSnapshot,
        [Parameter(Mandatory = $true)][string]$Platform,
        [Parameter(Mandatory = $true)][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [string[]]$RequiredRelativePath = @()
    )

    if ($null -eq $Receipt) { return $false }
    try { $null = Assert-RSBuildReceiptSemantics -Receipt $Receipt } catch { return $false }
    if ($Receipt.target -ne 'solution' -or
        -not [bool]$Receipt.tests_enabled -or
        $Receipt.source_snapshot_id -ne $SourceSnapshot.SnapshotId -or
        $Receipt.git_head -ne $SourceSnapshot.GitHead -or
        $Receipt.platform -ne $Platform -or $Receipt.configuration -ne $Configuration) { return $false }
    $artifactPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($artifact in @($Receipt.outputs)) {
        if (-not $artifactPaths.Add([string]$artifact.relative_path)) { return $false }
        try { $path = Resolve-RSBuildArtifactPath -RepoRoot $RepoRoot -RelativePath ([string]$artifact.relative_path) }
        catch { return $false }
        $item = Get-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        if ($null -eq $item -or $item.PSIsContainer -or
            ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            [int64]$item.Length -ne [int64]$artifact.size_bytes -or
            (Get-RSFileSha256 -LiteralPath $path) -ne $artifact.sha256) { return $false }
    }
    foreach ($relative in $RequiredRelativePath) {
        if (-not $artifactPaths.Contains($relative.Replace('\', '/'))) { return $false }
    }
    return $true
}

function Get-RSPortableBuildAttestationProjection {
    param([Parameter(Mandatory = $true)][object]$Attestation)

    [ordered]@{
        schema = $Attestation.schema
        canonicalization = $Attestation.canonicalization
        trust_scope = $Attestation.trust_scope
        source_commit = $Attestation.source_commit
        source_snapshot_id = $Attestation.source_snapshot_id
        producer = [ordered]@{
            repository = $Attestation.producer.repository
            workflow = $Attestation.producer.workflow
            run_id = $Attestation.producer.run_id
            run_attempt = [int64]$Attestation.producer.run_attempt
            job = $Attestation.producer.job
        }
        platform = $Attestation.platform
        configuration = $Attestation.configuration
        target = $Attestation.target
        project_selection = @($Attestation.project_selection)
        build_arguments = @($Attestation.build_arguments)
        tests_enabled = [bool]$Attestation.tests_enabled
        official_release = [bool]$Attestation.official_release
        toolchain_identity_digest = $Attestation.toolchain_identity_digest
        vcpkg_identity_digest = $Attestation.vcpkg_identity_digest
        ghostty_runtime_identity = $Attestation.ghostty_runtime_identity
        project_graph_digest = $Attestation.project_graph_digest
        runtime_dependency_digest = $Attestation.runtime_dependency_digest
        artifact_manifest_digest = $Attestation.artifact_manifest_digest
        artifacts = @($Attestation.artifacts | ForEach-Object {
                [ordered]@{
                    id = $_.id; role = $_.role; relative_path = $_.relative_path
                    size_bytes = [int64]$_.size_bytes; sha256 = $_.sha256
                    runtime_closure = @($_.runtime_closure)
                }
            })
    }
}

function Get-RSVerifiedBuildReceiptArtifacts {
    param(
        [Parameter(Mandatory = $true)][object]$Receipt,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir
    )

    $null = Assert-RSBuildReceiptSemantics -Receipt $Receipt
    $outputRoot = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $artifacts = foreach ($artifact in @($Receipt.outputs | Sort-Object relative_path)) {
        $path = Resolve-RSBuildArtifactPath -RepoRoot $RepoRoot -RelativePath ([string]$artifact.relative_path)
        if (-not $path.StartsWith($outputRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Receipt output is outside the selected build output: $($artifact.relative_path)"
        }
        $file = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        if ($file.PSIsContainer -or ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            [int64]$file.Length -ne [int64]$artifact.size_bytes -or
            (Get-RSFileSha256 -LiteralPath $file.FullName) -ne $artifact.sha256) {
            throw "Receipt output changed before portable publication: $($artifact.relative_path)"
        }
        [pscustomobject][ordered]@{
            id = $artifact.id
            role = $artifact.role
            relative_path = $artifact.relative_path
            size_bytes = [int64]$artifact.size_bytes
            sha256 = $artifact.sha256
            runtime_closure = @($artifact.runtime_closure)
        }
    }
    return @($artifacts)
}

function New-RSPortableBuildAttestation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Receipt,
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$ToolchainSchemaPath,
        [Parameter(Mandatory = $true)][string]$ProducerRepository,
        [Parameter(Mandatory = $true)][string]$ProducerWorkflow,
        [Parameter(Mandatory = $true)][string]$ProducerRunId,
        [Parameter(Mandatory = $true)][ValidateRange(1, [int]::MaxValue)][int]$ProducerRunAttempt,
        [Parameter(Mandatory = $true)][string]$ProducerJob,
        [datetime]$CreatedUtc = [datetime]::UtcNow
    )

    if ($ProducerRunId -notmatch '^[0-9]+$') { throw 'Portable producer run ID must contain only decimal digits.' }
    $portableArtifacts = @(Get-RSVerifiedBuildReceiptArtifacts -Receipt $Receipt `
        -RepoRoot $RepoRoot -BuildOutputDir $BuildOutputDir)
    $toolchainRelativePath = ".build/$($Receipt.platform)/$($Receipt.configuration)/toolchain-identity.json"
    $toolchainArtifact = @($portableArtifacts | Where-Object {
            [string]::Equals($_.relative_path, $toolchainRelativePath, [StringComparison]::OrdinalIgnoreCase)
        })
    if ($toolchainArtifact.Count -ne 1) {
        throw "Receipt must attest exactly one canonical toolchain identity record: $toolchainRelativePath"
    }
    $toolchainPath = Resolve-RSBuildArtifactPath -RepoRoot $RepoRoot -RelativePath $toolchainRelativePath
    $toolchainRecord = Read-RSValidationJsonFile -LiteralPath $toolchainPath `
        -SchemaPath $ToolchainSchemaPath -AllowedRoot ([IO.Path]::GetFullPath($BuildOutputDir))
    if ((Get-RSCanonicalJsonDigest -InputObject $toolchainRecord) -ne $Receipt.toolchain_identity_digest) {
        throw 'Canonical toolchain record does not match the build receipt toolchain identity digest.'
    }
    $attestation = [ordered]@{
        schema = 'red-salamander.portable-build-attestation.v1'
        canonicalization = 'rs-canonical-json-v1'
        attestation_id = ''
        trust_scope = 'same-workflow-run'
        source_commit = $Receipt.git_head
        source_snapshot_id = $Receipt.source_snapshot_id
        producer = [ordered]@{
            repository = $ProducerRepository
            workflow = $ProducerWorkflow
            run_id = $ProducerRunId
            run_attempt = [int64]$ProducerRunAttempt
            job = $ProducerJob
        }
        platform = $Receipt.platform
        configuration = $Receipt.configuration
        target = $Receipt.target
        project_selection = @($Receipt.project_selection)
        build_arguments = @($Receipt.build_arguments)
        tests_enabled = [bool]$Receipt.tests_enabled
        official_release = [bool]$Receipt.official_release
        toolchain_identity_digest = $Receipt.toolchain_identity_digest
        vcpkg_identity_digest = $Receipt.vcpkg_identity_digest
        ghostty_runtime_identity = $Receipt.ghostty_runtime_identity
        project_graph_digest = $Receipt.project_graph_digest
        runtime_dependency_digest = $Receipt.runtime_dependency_digest
        artifact_manifest_digest = Get-RSCanonicalJsonDigest -InputObject $portableArtifacts
        artifacts = $portableArtifacts
        created_utc = $CreatedUtc.ToUniversalTime().ToString('o')
    }
    $attestation.attestation_id = Get-RSCanonicalJsonDigest -InputObject (
        Get-RSPortableBuildAttestationProjection -Attestation ([pscustomobject]$attestation))
    return [pscustomobject]$attestation
}

function Publish-RSPortableBuildAttestation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][object]$Attestation,
        [Parameter(Mandatory = $true)][string]$SchemaPath
    )

    $root = [IO.Path]::GetFullPath($BuildOutputDir)
    $path = Join-Path $root 'portable-build-attestation.json'
    Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $Attestation -SchemaPath $SchemaPath -AllowedRoot $root
    return $path
}

function Publish-RSPortableBuildPayload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][object]$Receipt,
        [Parameter(Mandatory = $true)][object]$Attestation,
        [Parameter(Mandatory = $true)][string]$PortableSchemaPath,
        [string]$StagingRoot
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $outputRoot = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $verifiedArtifacts = @(Get-RSVerifiedBuildReceiptArtifacts -Receipt $Receipt `
        -RepoRoot $root -BuildOutputDir $outputRoot)
    if ((Get-RSCanonicalJsonDigest -InputObject $verifiedArtifacts) -ne $Attestation.artifact_manifest_digest -or
        (Get-RSCanonicalJsonDigest -InputObject @($Attestation.artifacts)) -ne $Attestation.artifact_manifest_digest) {
        throw 'Portable payload artifacts do not exactly match the verified receipt output set.'
    }

    if ([string]::IsNullOrWhiteSpace($StagingRoot)) {
        $StagingRoot = Join-Path $root `
            ('.build\PortablePayload\{0}' -f [Guid]::NewGuid().ToString('N'))
    }
    $stagingParent = Resolve-RSGuardedRepositoryPath -RepoRoot $root `
        -GuardedRelativeRoot '.build\PortablePayload' -Path (Split-Path $StagingRoot -Parent) -CreateDirectory
    $stage = Join-Path $stagingParent (Split-Path $StagingRoot -Leaf)
    $stage = Resolve-RSGuardedRepositoryPath -RepoRoot $root `
        -GuardedRelativeRoot '.build\PortablePayload' -Path $stage
    if (Test-Path -LiteralPath $stage) {
        throw "Portable payload staging root must be unique and absent: $stage"
    }
    [void][IO.Directory]::CreateDirectory($stage)
    $stage = Resolve-RSGuardedRepositoryPath -RepoRoot $root `
        -GuardedRelativeRoot '.build\PortablePayload' -Path $stage -CreateDirectory

    foreach ($artifact in $verifiedArtifacts) {
        $source = Resolve-RSBuildArtifactPath -RepoRoot $root -RelativePath ([string]$artifact.relative_path)
        $relativeWithinOutput = [IO.Path]::GetRelativePath($outputRoot, $source)
        if ($relativeWithinOutput.StartsWith('..', [StringComparison]::Ordinal)) {
            throw "Portable payload source escaped the build output: $($artifact.relative_path)"
        }
        $destination = Join-Path $stage $relativeWithinOutput
        $destinationParent = Resolve-RSGuardedRepositoryPath -RepoRoot $root `
            -GuardedRelativeRoot '.build\PortablePayload' -Path (Split-Path $destination -Parent) -CreateDirectory
        $destination = Resolve-RSGuardedRepositoryPath -RepoRoot $root `
            -GuardedRelativeRoot '.build\PortablePayload' `
            -Path (Join-Path $destinationParent (Split-Path $destination -Leaf))
        [IO.File]::Copy($source, $destination, $false)
        if ((Get-RSFileSha256 -LiteralPath $destination) -ne $artifact.sha256) {
            throw "Portable payload copy verification failed: $($artifact.relative_path)"
        }
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $stage 'portable-build-attestation.json') `
        -InputObject $Attestation -SchemaPath $PortableSchemaPath -AllowedRoot $stage
    return $stage
}

function Import-RSPortableBuildAttestation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$BuildOutputDir,
        [Parameter(Mandatory = $true)][string]$PortableSchemaPath,
        [Parameter(Mandatory = $true)][string]$BuildReceiptSchemaPath,
        [Parameter(Mandatory = $true)][string]$ToolchainSchemaPath,
        [Parameter(Mandatory = $true)][string]$ExpectedRepository,
        [Parameter(Mandatory = $true)][string]$ExpectedWorkflow,
        [Parameter(Mandatory = $true)][string]$ExpectedRunId,
        [Parameter(Mandatory = $true)][int]$ExpectedRunAttempt,
        [Parameter(Mandatory = $true)][string]$ExpectedProducerJob,
        [Parameter(Mandatory = $true)][string]$ExpectedAttestationId,
        [Parameter(Mandatory = $true)][string]$ExpectedToolchainIdentityDigest,
        [Parameter(Mandatory = $true)][string]$ExpectedSourceCommit,
        [Parameter(Mandatory = $true)][string]$ExpectedPlatform,
        [Parameter(Mandatory = $true)][string]$ExpectedConfiguration
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $outputRoot = [IO.Path]::GetFullPath($BuildOutputDir).TrimEnd('\')
    $attestationPath = Join-Path $outputRoot 'portable-build-attestation.json'
    $attestation = Read-RSValidationJsonFile -LiteralPath $attestationPath -SchemaPath $PortableSchemaPath -AllowedRoot $outputRoot
    $computedAttestationId = Get-RSCanonicalJsonDigest -InputObject (Get-RSPortableBuildAttestationProjection -Attestation $attestation)
    if ($attestation.attestation_id -ne $computedAttestationId -or
        $attestation.attestation_id -ne $ExpectedAttestationId -or
        $attestation.toolchain_identity_digest -ne $ExpectedToolchainIdentityDigest) {
        throw 'Portable build attestation identity or trusted toolchain output mismatch.'
    }
    if ($attestation.producer.repository -cne $ExpectedRepository -or
        $attestation.producer.workflow -cne $ExpectedWorkflow -or
        $attestation.producer.run_id -cne $ExpectedRunId -or
        [int64]$attestation.producer.run_attempt -ne $ExpectedRunAttempt -or
        $attestation.producer.job -cne $ExpectedProducerJob -or
        $attestation.source_commit -cne $ExpectedSourceCommit.ToLowerInvariant() -or
        $attestation.platform -cne $ExpectedPlatform -or
        $attestation.configuration -cne $ExpectedConfiguration) {
        throw 'Portable build attestation provenance does not match this consumer job.'
    }
    $snapshot = Get-RSBuildSourceSnapshot -RepoRoot $root
    if ($snapshot.GitHead -ne $attestation.source_commit -or $snapshot.SnapshotId -ne $attestation.source_snapshot_id) {
        throw 'Portable build attestation source snapshot does not match this checkout.'
    }
    if ((Get-RSCanonicalJsonDigest -InputObject @($attestation.artifacts)) -ne $attestation.artifact_manifest_digest) {
        throw 'Portable build artifact manifest digest mismatch.'
    }

    $expectedPaths = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $localArtifacts = foreach ($artifact in @($attestation.artifacts)) {
        if (-not $expectedPaths.Add([string]$artifact.relative_path)) { throw "Duplicate portable artifact path: $($artifact.relative_path)" }
        $path = Resolve-RSBuildArtifactPath -RepoRoot $root -RelativePath ([string]$artifact.relative_path)
        if (-not $path.StartsWith($outputRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Portable artifact is outside the expected output directory: $($artifact.relative_path)"
        }
        $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "Portable artifact must be a regular file: $($artifact.relative_path)"
        }
        if ([int64]$item.Length -ne [int64]$artifact.size_bytes -or (Get-RSFileSha256 -LiteralPath $path) -ne $artifact.sha256) {
            throw "Portable artifact digest mismatch: $($artifact.relative_path)"
        }
        [pscustomobject][ordered]@{
            id = $artifact.id; role = $artifact.role; relative_path = $artifact.relative_path
            absolute_path_diagnostic = $path; size_bytes = [int64]$artifact.size_bytes
            sha256 = $artifact.sha256; runtime_closure = @($artifact.runtime_closure)
        }
    }
    $actualPaths = @(Get-ChildItem -LiteralPath $outputRoot -File -Recurse | Where-Object {
            $_.FullName -notlike "*\receipts\*" -and
            $_.Name -notin @('build-receipt.json', 'portable-build-attestation.json')
        } | ForEach-Object { ConvertTo-RSBuildRelativePath -RepoRoot $root -LiteralPath $_.FullName })
    foreach ($path in $actualPaths) {
        if (-not $expectedPaths.Contains($path)) { throw "Unattested file exists in portable build payload: $path" }
    }
    if ($actualPaths.Count -ne $expectedPaths.Count) { throw 'Portable build payload file count does not match its manifest.' }

    $toolchainRelativePath = ".build/$ExpectedPlatform/$ExpectedConfiguration/toolchain-identity.json"
    if (-not $expectedPaths.Contains($toolchainRelativePath)) {
        throw "Portable build payload omits the canonical toolchain identity record: $toolchainRelativePath"
    }
    $toolchainPath = Resolve-RSBuildArtifactPath -RepoRoot $root -RelativePath $toolchainRelativePath
    $toolchainRecord = Read-RSValidationJsonFile -LiteralPath $toolchainPath `
        -SchemaPath $ToolchainSchemaPath -AllowedRoot $outputRoot
    if ((Get-RSCanonicalJsonDigest -InputObject $toolchainRecord) -ne $attestation.toolchain_identity_digest) {
        throw 'Portable toolchain record does not match the attested toolchain identity digest.'
    }

    $identity = [pscustomobject]@{
        RepoIdentity = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ physical_root = $root.ToLowerInvariant() })
        ToolchainIdentityDigest = $attestation.toolchain_identity_digest
        VcpkgIdentityDigest = $attestation.vcpkg_identity_digest
        GhosttyRuntimeIdentity = $attestation.ghostty_runtime_identity
        ProjectGraphDigest = $attestation.project_graph_digest
        RuntimeDependencyDigest = $attestation.runtime_dependency_digest
    }
    $created = [datetime]::Parse($attestation.created_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    $receipt = New-RSBuildReceipt -SourceSnapshot $snapshot -Identity $identity `
        -Platform $attestation.platform -Configuration $attestation.configuration -Target $attestation.target `
        -ProjectSelection @($attestation.project_selection) -BuildArguments @($attestation.build_arguments) `
        -TestsEnabled ([bool]$attestation.tests_enabled) -OfficialRelease ([bool]$attestation.official_release) `
        -StartedUtc $created -EndedUtc $created -Duration ([TimeSpan]::Zero) -Outputs @($localArtifacts)
    return Publish-RSBuildReceipt -BuildOutputDir $outputRoot -Receipt $receipt -SchemaPath $BuildReceiptSchemaPath
}

function ConvertTo-RSBuildSummaryLines {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Configuration,
        [Parameter(Mandatory = $true)][string]$Platform,
        [Parameter(Mandatory = $true)][TimeSpan]$Duration,
        [Parameter(Mandatory = $true)][object[]]$Artifacts
    )

    $lines = @(
        '========================================',
        'Build completed successfully!',
        "Configuration: $Configuration | Platform: $Platform",
        ('Build time: ' + (Format-RSBuildDuration -Duration $Duration)),
        '========================================',
        'Outputs (copy/paste any PowerShell command to run):'
    )
    $lines += @($Artifacts | Sort-Object relative_path | ForEach-Object { ConvertTo-RSBuildOutputLine -Artifact $_ })
    return $lines
}

Export-ModuleMember -Function @(
    'Format-RSBuildDuration',
    'Format-RSBuildSizeMiB',
    'ConvertTo-RSBuildOutputLine',
    'New-RSBuildArtifactRecord',
    'Get-RSBuildArtifactRecords',
    'Test-RSBuildArtifactManifestTestsEnabled',
    'Get-RSBuildSourceSnapshot',
    'New-RSBuildToolchainIdentityRecord',
    'Get-RSBuildToolchainIdentity',
    'Publish-RSBuildToolchainIdentity',
    'Get-RSGhosttyRuntimeIdentityInputPaths',
    'Get-RSGhosttyRuntimeIdentityDigest',
    'Get-RSBuildIdentityBundle',
    'New-RSBuildReceipt',
    'Publish-RSBuildReceipt',
    'Read-RSBuildReceipt',
    'Suspend-RSBuildReceipt',
    'Test-RSBuildReceiptCompatible',
    'Test-RSBuildReceiptForArtifactUse',
    'New-RSPortableBuildAttestation',
    'Publish-RSPortableBuildAttestation',
    'Publish-RSPortableBuildPayload',
    'Import-RSPortableBuildAttestation',
    'ConvertTo-RSBuildSummaryLines'
)
