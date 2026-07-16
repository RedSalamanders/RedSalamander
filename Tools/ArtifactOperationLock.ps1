Set-StrictMode -Version Latest

$artifactLockStateVariable = Get-Variable -Name RSArtifactOperationLocks -Scope Global -ErrorAction SilentlyContinue
if ($null -eq $artifactLockStateVariable) {
    $global:RSArtifactOperationLocks = @{}
}

function Get-RSArtifactOperationLockPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path ([System.IO.Path]::GetFullPath($RepoRoot)) '.build\artifact-operation.lock'
}

function Get-RSArtifactOperationStateKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LockPath
    )

    return [System.IO.Path]::GetFullPath($LockPath).ToUpperInvariant()
}

function Get-RSArtifactOperationOwnerMetadataPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path ([System.IO.Path]::GetFullPath($RepoRoot)) '.build\artifact-operation-owner.json'
}

function Read-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $null
    }

    try {
        $metadata = Get-Content -LiteralPath $metadataPath -Raw | ConvertFrom-Json
        $requiredProperties = @(
            'schema',
            'lock_path',
            'repo_root',
            'owner_pid',
            'owner_process_start_utc_ticks',
            'operation',
            'started_utc',
            'token'
        )
        foreach ($propertyName in $requiredProperties) {
            if ($null -eq $metadata.PSObject.Properties[$propertyName]) {
                return $null
            }
        }

        $schema = [string]$metadata.schema
        if ($schema -notin @(
                'red-salamander.artifact-operation-owner.v2',
                'red-salamander.artifact-operation-owner.v3')) {
            return $null
        }

        if ($schema -eq 'red-salamander.artifact-operation-owner.v3' -and
            $null -eq $metadata.PSObject.Properties['scope']) {
            return $null
        }

        $ownerPid = 0
        $ownerProcessStartUtcTicks = 0L
        if (-not [int]::TryParse([string]$metadata.owner_pid, [ref]$ownerPid) -or $ownerPid -le 0 -or
            -not [long]::TryParse([string]$metadata.owner_process_start_utc_ticks, [ref]$ownerProcessStartUtcTicks) -or
            $ownerProcessStartUtcTicks -le 0 -or
            [string]::IsNullOrWhiteSpace([string]$metadata.lock_path) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.operation) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.started_utc) -or
            [string]::IsNullOrWhiteSpace([string]$metadata.token)) {
            return $null
        }

        return $metadata
    }
    catch {
        return $null
    }
}

function Write-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$LockPath,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Token,

        [Parameter(Mandatory = $true)]
        [string]$StartedUtc,

        [Parameter(Mandatory = $true)]
        [long]$OwnerProcessStartUtcTicks,

        [hashtable]$Scope = @{}
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $normalizedRoot
    $metadataDirectory = Split-Path -Parent $metadataPath
    [void](New-Item -ItemType Directory -Path $metadataDirectory -Force)
    $temporaryPath = "$metadataPath.$PID.$([Guid]::NewGuid().ToString('N')).tmp"
    $backupPath = "$metadataPath.$PID.$([Guid]::NewGuid().ToString('N')).bak"
    $normalizedScope = [ordered]@{}
    foreach ($scopeName in @('kind', 'target', 'configuration', 'platform', 'suite')) {
        if ($Scope.ContainsKey($scopeName) -and
            -not [string]::IsNullOrWhiteSpace([string]$Scope[$scopeName])) {
            $normalizedScope[$scopeName] = [string]$Scope[$scopeName]
        }
    }

    $payload = [ordered]@{
        schema = 'red-salamander.artifact-operation-owner.v3'
        lock_path = [System.IO.Path]::GetFullPath($LockPath)
        repo_root = $normalizedRoot
        owner_pid = $PID
        owner_process_start_utc_ticks = $OwnerProcessStartUtcTicks
        operation = $Operation
        started_utc = $StartedUtc
        token = $Token
        scope = $normalizedScope
    }

    try {
        $json = $payload | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($temporaryPath, $json, [System.Text.UTF8Encoding]::new($false))
        if (Test-Path -LiteralPath $metadataPath) {
            [System.IO.File]::Replace($temporaryPath, $metadataPath, $backupPath)
        }
        else {
            [System.IO.File]::Move($temporaryPath, $metadataPath)
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }

    return $metadataPath
}

function Remove-RSArtifactOperationOwnerMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $metadataPath)) {
        return $false
    }

    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot
    if ($null -eq $metadata -or
        -not [string]::Equals([string]$metadata.token, $Token, [System.StringComparison]::Ordinal)) {
        return $false
    }

    Remove-Item -LiteralPath $metadataPath -Force
    return $true
}

function Set-RSArtifactOperationLockMarker {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileStream]$LockStream,

        [Parameter(Mandatory = $true)]
        [string]$Token
    )

    $markerBytes = [System.Text.Encoding]::UTF8.GetBytes(
        "red-salamander.artifact-operation-lock.v1`ntoken=$Token`n")
    $LockStream.SetLength(0)
    $LockStream.Position = 0
    $LockStream.Write($markerBytes, 0, $markerBytes.Length)
    $LockStream.Flush($true)
}

function Clear-RSArtifactOperationLockMarker {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileStream]$LockStream
    )

    $LockStream.SetLength(0)
    $LockStream.Flush($true)
}

function Get-RSImmediateParentProcessId {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ProcessId
    )

    try {
        $process = @(Get-CimInstance Win32_Process -Filter "ProcessId=$ProcessId" -ErrorAction Stop |
                Select-Object -First 1)
        if ($process.Count -eq 0) {
            return 0
        }

        return [int]$process[0].ParentProcessId
    }
    catch {
        return 0
    }
}

function Test-RSArtifactOperationOwnerProcess {
    param(
        [Parameter(Mandatory = $true)]
        [int]$OwnerProcessId,

        [Parameter(Mandatory = $true)]
        [long]$OwnerProcessStartUtcTicks
    )

    $ownerProcess = $null
    try {
        $ownerProcess = Get-Process -Id $OwnerProcessId -ErrorAction Stop
        return $ownerProcess.StartTime.ToUniversalTime().Ticks -eq $OwnerProcessStartUtcTicks
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $ownerProcess) {
            $ownerProcess.Dispose()
        }
    }
}

function Get-RSDelegatedArtifactOperationOwner {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$LockPath
    )

    $token = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $ownerPidText = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', 'Process')
    $inheritedLockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    $delegationToken = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN', 'Process')
    $delegationParentPidText = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID', 'Process')
    $delegationReadyEvent = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT', 'Process')
    if ([string]::IsNullOrWhiteSpace($token) -or
        [string]::IsNullOrWhiteSpace($ownerPidText) -or
        [string]::IsNullOrWhiteSpace($inheritedLockPath) -or
        [string]::IsNullOrWhiteSpace($delegationToken) -or
        [string]::IsNullOrWhiteSpace($delegationParentPidText) -or
        [string]::IsNullOrWhiteSpace($delegationReadyEvent) -or
        -not [string]::Equals(
            $delegationReadyEvent,
            "Local\RedSalamander.ArtifactDelegation.$delegationToken",
            [System.StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [System.IO.Path]::GetFullPath($inheritedLockPath),
            [System.IO.Path]::GetFullPath($LockPath),
            [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    $ownerPid = 0
    $delegationParentPid = 0
    if (-not [int]::TryParse($ownerPidText, [ref]$ownerPid) -or
        -not [int]::TryParse($delegationParentPidText, [ref]$delegationParentPid) -or
        $ownerPid -le 0 -or
        $ownerPid -eq $PID -or
        $delegationParentPid -ne $ownerPid -or
        (Get-RSImmediateParentProcessId -ProcessId $PID) -ne $ownerPid) {
        return $null
    }

    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot
    if ($null -eq $metadata -or
        -not [string]::Equals([string]$metadata.token, $token, [System.StringComparison]::Ordinal) -or
        -not [string]::Equals(
            [System.IO.Path]::GetFullPath([string]$metadata.lock_path),
            [System.IO.Path]::GetFullPath($LockPath),
            [System.StringComparison]::OrdinalIgnoreCase) -or
        [int]$metadata.owner_pid -ne $ownerPid -or
        -not (Test-RSArtifactOperationOwnerProcess `
            -OwnerProcessId $ownerPid `
            -OwnerProcessStartUtcTicks ([long]$metadata.owner_process_start_utc_ticks))) {
        return $null
    }

    $readyEvent = $null
    try {
        $readyEvent = [System.Threading.EventWaitHandle]::OpenExisting($delegationReadyEvent)
        if (-not $readyEvent.WaitOne(5000)) {
            return $null
        }
    }
    catch {
        return $null
    }
    finally {
        if ($null -ne $readyEvent) {
            $readyEvent.Dispose()
        }
    }

    $jobType = 'RedSalamander.Tooling.KillOnCloseJob' -as [type]
    if ($null -eq $jobType -or -not [RedSalamander.Tooling.KillOnCloseJob]::IsCurrentProcessInJob()) {
        return $null
    }

    if (-not (Test-RSArtifactOperationOwnerProcess `
            -OwnerProcessId $ownerPid `
            -OwnerProcessStartUtcTicks ([long]$metadata.owner_process_start_utc_ticks))) {
        return $null
    }

    foreach ($environmentName in @(
            'REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN',
            'REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID',
            'REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT')) {
        [Environment]::SetEnvironmentVariable($environmentName, $null, 'Process')
    }

    return $metadata
}

function Format-RSArtifactOperationOwnerDiagnostic {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $metadataPath = Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $RepoRoot
    $metadata = Read-RSArtifactOperationOwnerMetadata -RepoRoot $RepoRoot
    if ($null -eq $metadata) {
        return "Owner metadata is unavailable or malformed at '$metadataPath'."
    }

    return "Owner PID=$($metadata.owner_pid); Operation='$($metadata.operation)'; StartedUtc='$($metadata.started_utc)'."
}

function Open-RSArtifactOperationLockStream {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LockPath,

        [ValidateRange(0, 60000)]
        [int]$TimeoutMilliseconds = 0
    )

    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMilliseconds)
    do {
        try {
            return [System.IO.FileStream]::new(
                $LockPath,
                [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            $nativeError = $_.Exception.HResult -band 0xFFFF
            if ($nativeError -notin @(32, 33)) {
                throw
            }

            if ($TimeoutMilliseconds -eq 0 -or [DateTime]::UtcNow -ge $deadline) {
                return $null
            }

            Start-Sleep -Milliseconds 25
        }
    } while ($true)
}

function Enter-RSArtifactOperationLock {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Operation,

        [hashtable]$Scope = @{},

        [ValidateRange(0, 60000)]
        [int]$TimeoutMilliseconds = 0
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $lockPath = Get-RSArtifactOperationLockPath -RepoRoot $normalizedRoot
    $lockDirectory = Split-Path -Parent $lockPath
    [void](New-Item -ItemType Directory -Path $lockDirectory -Force)
    $stateKey = Get-RSArtifactOperationStateKey -LockPath $lockPath
    $managedThreadId = [Environment]::CurrentManagedThreadId
    if ($global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        $state = $global:RSArtifactOperationLocks[$stateKey]
        if ([int]$state.OwnerManagedThreadId -eq $managedThreadId) {
            $leaseId = [Guid]::NewGuid().ToString('N')
            $state.Leases[$leaseId] = $true
            return [pscustomobject]@{
                Name = $lockPath
                StateKey = $stateKey
                LeaseId = $leaseId
                Operation = $Operation
                WasAbandoned = $false
                AbandonedOwner = $null
                IsDelegated = $false
                IsRootOwner = $false
            }
        }
    }

    $delegatedOwner = Get-RSDelegatedArtifactOperationOwner -RepoRoot $normalizedRoot -LockPath $lockPath
    if ($null -ne $delegatedOwner) {
        return [pscustomobject]@{
            Name = $lockPath
            StateKey = $stateKey
            LeaseId = ''
            Operation = $Operation
            WasAbandoned = $false
            AbandonedOwner = $null
            IsDelegated = $true
            IsRootOwner = $false
        }
    }

    $lockStream = Open-RSArtifactOperationLockStream `
        -LockPath $lockPath `
        -TimeoutMilliseconds $TimeoutMilliseconds
    if ($null -eq $lockStream) {
        $ownerDiagnostic = Format-RSArtifactOperationOwnerDiagnostic -RepoRoot $normalizedRoot
        throw "Cannot start '$Operation' because another RedSalamander build or test operation owns '$lockPath'. $ownerDiagnostic Wait for it to finish and retry."
    }

    $abandonedOwner = Read-RSArtifactOperationOwnerMetadata -RepoRoot $normalizedRoot
    $wasAbandoned = $lockStream.Length -gt 0 -or $null -ne $abandonedOwner -or
        (Test-Path -LiteralPath (Get-RSArtifactOperationOwnerMetadataPath -RepoRoot $normalizedRoot))
    $token = [Guid]::NewGuid().ToString('N')
    $leaseId = [Guid]::NewGuid().ToString('N')
    $startedUtc = [DateTime]::UtcNow.ToString('o')
    $currentProcess = [System.Diagnostics.Process]::GetCurrentProcess()
    try {
        $ownerProcessStartUtcTicks = $currentProcess.StartTime.ToUniversalTime().Ticks
    }
    finally {
        $currentProcess.Dispose()
    }

    $previousToken = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $previousOwnerPid = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', 'Process')
    $previousLockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    try {
        Set-RSArtifactOperationLockMarker -LockStream $lockStream -Token $token
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $token, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $PID.ToString(), 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $lockPath, 'Process')
        [void](Write-RSArtifactOperationOwnerMetadata `
                -RepoRoot $normalizedRoot `
                -LockPath $lockPath `
                -Operation $Operation `
                -Token $token `
                -StartedUtc $startedUtc `
                -OwnerProcessStartUtcTicks $ownerProcessStartUtcTicks `
                -Scope $Scope)

        $leases = @{}
        $leases[$leaseId] = $true
        $global:RSArtifactOperationLocks[$stateKey] = [pscustomobject]@{
            LockStream = $lockStream
            LockPath = $lockPath
            RepoRoot = $normalizedRoot
            Token = $token
            OwnerManagedThreadId = $managedThreadId
            Leases = $leases
            ActiveDelegations = @{}
            PreviousToken = $previousToken
            PreviousOwnerPid = $previousOwnerPid
            PreviousLockPath = $previousLockPath
        }

        return [pscustomobject]@{
            Name = $lockPath
            StateKey = $stateKey
            LeaseId = $leaseId
            Operation = $Operation
            WasAbandoned = $wasAbandoned
            AbandonedOwner = $abandonedOwner
            IsDelegated = $false
            IsRootOwner = $true
        }
    }
    catch {
        [void]$global:RSArtifactOperationLocks.Remove($stateKey)
        [void](Remove-RSArtifactOperationOwnerMetadata -RepoRoot $normalizedRoot -Token $token)
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $previousToken, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $previousOwnerPid, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $previousLockPath, 'Process')
        try {
            Clear-RSArtifactOperationLockMarker -LockStream $lockStream
        }
        finally {
            $lockStream.Dispose()
        }
        throw
    }
}

function New-RSArtifactOperationChildDelegation {
    $token = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', 'Process')
    $lockPath = [Environment]::GetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', 'Process')
    if ([string]::IsNullOrWhiteSpace($token) -or [string]::IsNullOrWhiteSpace($lockPath)) {
        return $null
    }

    $stateKey = Get-RSArtifactOperationStateKey -LockPath $lockPath
    if (-not $global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        return $null
    }

    $state = $global:RSArtifactOperationLocks[$stateKey]
    if (-not [string]::Equals([string]$state.Token, $token, [System.StringComparison]::Ordinal) -or
        [int]$state.OwnerManagedThreadId -ne [Environment]::CurrentManagedThreadId) {
        return $null
    }

    $delegationId = [Guid]::NewGuid().ToString('N')
    $readyEventName = "Local\RedSalamander.ArtifactDelegation.$delegationId"
    $readyEvent = [System.Threading.EventWaitHandle]::new(
        $false,
        [System.Threading.EventResetMode]::ManualReset,
        $readyEventName)
    $delegation = [pscustomobject]@{
        Id = $delegationId
        StateKey = $stateKey
        Token = $token
        ReadyEvent = $readyEvent
        Environment = [ordered]@{
            REDSALAMANDER_ARTIFACT_DELEGATION_TOKEN = $delegationId
            REDSALAMANDER_ARTIFACT_DELEGATION_PARENT_PID = $PID.ToString()
            REDSALAMANDER_ARTIFACT_DELEGATION_READY_EVENT = $readyEventName
        }
    }
    $state.ActiveDelegations[$delegationId] = $delegation
    return $delegation
}

function Complete-RSArtifactOperationChildDelegation {
    param(
        [AllowNull()]
        [object]$Delegation
    )

    if ($null -eq $Delegation) {
        return
    }

    $stateKey = [string]$Delegation.StateKey
    if ($global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        $state = $global:RSArtifactOperationLocks[$stateKey]
        if ([string]::Equals([string]$state.Token, [string]$Delegation.Token, [System.StringComparison]::Ordinal)) {
            [void]$state.ActiveDelegations.Remove([string]$Delegation.Id)
        }
    }

    $Delegation.ReadyEvent.Dispose()
}

function Exit-RSArtifactOperationLock {
    param(
        [AllowNull()]
        [object]$Lock
    )

    if ($null -eq $Lock -or [bool]$Lock.IsDelegated) {
        return
    }

    $stateKey = [string]$Lock.StateKey
    if (-not $global:RSArtifactOperationLocks.ContainsKey($stateKey)) {
        return
    }

    $state = $global:RSArtifactOperationLocks[$stateKey]
    if ([int]$state.OwnerManagedThreadId -ne [Environment]::CurrentManagedThreadId) {
        throw "Artifact-operation lease '$($Lock.LeaseId)' must be released on its owning managed thread."
    }

    $leaseId = [string]$Lock.LeaseId
    if (-not $state.Leases.ContainsKey($leaseId)) {
        return
    }

    if ($state.Leases.Count -eq 1 -and $state.ActiveDelegations.Count -gt 0) {
        throw "Cannot release the root artifact-operation lease while $($state.ActiveDelegations.Count) contained child delegation(s) remain active."
    }

    [void]$state.Leases.Remove($leaseId)
    if ($state.Leases.Count -gt 0) {
        return
    }

    try {
        [void](Remove-RSArtifactOperationOwnerMetadata -RepoRoot $state.RepoRoot -Token $state.Token)
        Clear-RSArtifactOperationLockMarker -LockStream $state.LockStream
    }
    finally {
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_TOKEN', $state.PreviousToken, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_OWNER_PID', $state.PreviousOwnerPid, 'Process')
        [Environment]::SetEnvironmentVariable('REDSALAMANDER_ARTIFACT_OPERATION_LOCK_PATH', $state.PreviousLockPath, 'Process')
        $state.LockStream.Dispose()
        [void]$global:RSArtifactOperationLocks.Remove($stateKey)
    }
}

function Get-RSArtifactContaminationMarkerPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Join-Path ([System.IO.Path]::GetFullPath($RepoRoot)) '.build\artifact-operation-contaminated.json'
}

function Set-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Reason,

        [AllowNull()]
        [object]$AbandonedOwner = $null
    )

    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot
    $markerDirectory = Split-Path -Parent $markerPath
    [void](New-Item -ItemType Directory -Path $markerDirectory -Force)

    $abandonedOperation = ''
    $abandonedStartedUtc = ''
    $abandonedScope = [ordered]@{}
    if ($null -ne $AbandonedOwner) {
        if ($null -ne $AbandonedOwner.PSObject.Properties['operation']) {
            $abandonedOperation = [string]$AbandonedOwner.operation
        }
        if ($null -ne $AbandonedOwner.PSObject.Properties['started_utc']) {
            $abandonedStartedUtc = [string]$AbandonedOwner.started_utc
        }
        if ($null -ne $AbandonedOwner.PSObject.Properties['scope'] -and
            $null -ne $AbandonedOwner.scope) {
            foreach ($scopeName in @('kind', 'target', 'configuration', 'platform', 'suite')) {
                if ($null -ne $AbandonedOwner.scope.PSObject.Properties[$scopeName] -and
                    -not [string]::IsNullOrWhiteSpace([string]$AbandonedOwner.scope.$scopeName)) {
                    $abandonedScope[$scopeName] = [string]$AbandonedOwner.scope.$scopeName
                }
            }
        }
    }

    $payload = [ordered]@{
        schema = 'red-salamander.artifact-operation-contamination.v2'
        recorded_utc = [DateTime]::UtcNow.ToString('o')
        process_id = $PID
        reason = $Reason
        abandoned_operation = $abandonedOperation
        abandoned_started_utc = $abandonedStartedUtc
        abandoned_scope = $abandonedScope
    }
    $json = $payload | ConvertTo-Json -Depth 4
    [System.IO.File]::WriteAllText($markerPath, $json, [System.Text.UTF8Encoding]::new($false))
    return $markerPath
}

function Read-RSArtifactOperationContamination {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot
    if (-not (Test-Path -LiteralPath $markerPath)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $markerPath -Raw | ConvertFrom-Json
    }
    catch {
        return [pscustomobject]@{
            schema = 'red-salamander.artifact-operation-contamination.unknown'
            reason = 'The contamination marker is malformed.'
            abandoned_operation = ''
            abandoned_started_utc = ''
            abandoned_scope = [pscustomobject]@{}
        }
    }
}

function Test-RSArtifactOperationRepairAllowed {
    param(
        [AllowNull()]
        [object]$Contamination,

        [switch]$Rebuild,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$ProjectName = '',

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$Platform
    )

    if ($null -eq $Contamination -or -not $Rebuild -or
        -not [string]::IsNullOrWhiteSpace($ProjectName)) {
        return $false
    }

    $scopeProperty = $Contamination.PSObject.Properties['abandoned_scope']
    if ($null -eq $scopeProperty -or $null -eq $scopeProperty.Value) {
        # Legacy markers did not record scope. A full-solution rebuild is the only
        # safe automated recovery available for those markers.
        return $true
    }

    $scope = $scopeProperty.Value
    $recordedConfiguration = if ($null -ne $scope.PSObject.Properties['configuration']) {
        [string]$scope.configuration
    } else { '' }
    $recordedPlatform = if ($null -ne $scope.PSObject.Properties['platform']) {
        [string]$scope.platform
    } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($recordedConfiguration) -and
        -not [string]::Equals($recordedConfiguration, $Configuration, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    if (-not [string]::IsNullOrWhiteSpace($recordedPlatform) -and
        -not [string]::Equals($recordedPlatform, $Platform, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    return $true
}

function Test-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    return Test-Path -LiteralPath (Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot)
}

function Clear-RSArtifactOperationContaminated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $markerPath = Get-RSArtifactContaminationMarkerPath -RepoRoot $RepoRoot
    if (Test-Path -LiteralPath $markerPath) {
        Remove-Item -LiteralPath $markerPath -Force
    }
}

function Get-RSResidualArtifactToolProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $normalizedRoot = [System.IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $normalizedBuildRoot = Join-Path $normalizedRoot '.build'
    $processes = @()
    try {
        $processes = @(Get-CimInstance Win32_Process `
                -Filter "Name='MSBuild.exe' OR Name='cl.exe' OR Name='link.exe'" `
                -ErrorAction Stop)
    }
    catch {
        throw "Unable to verify that no residual RedSalamander build tools are active: $($_.Exception.Message)"
    }

    return @($processes | Where-Object {
            $commandLine = [string]$_.CommandLine
            if ([string]::IsNullOrWhiteSpace($commandLine)) {
                return $false
            }

            return $commandLine.IndexOf($normalizedBuildRoot, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $commandLine.IndexOf($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
        })
}

function Assert-RSNoResidualArtifactToolProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $processes = @(Get-RSResidualArtifactToolProcesses -RepoRoot $RepoRoot)
    if ($processes.Count -eq 0) {
        return
    }

    $diagnostics = @($processes | ForEach-Object {
            "  PID=$($_.ProcessId); ParentPID=$($_.ParentProcessId); Name='$($_.Name)'; CommandLine='$($_.CommandLine)'"
        })
    throw ("Cannot start while a build tool is still touching this repository. " +
           "Wait for it to exit; do not start another incremental build against shared intermediates.`n" +
           ($diagnostics -join "`n"))
}
