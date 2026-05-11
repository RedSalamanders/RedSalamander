function Get-RSNormalizedEnvironmentMap {
    $environment = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pathSegments = [System.Collections.Generic.List[string]]::new()
    $seenPathSegments = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($scope in @('Machine', 'User', 'Process')) {
        foreach ($entry in [System.Environment]::GetEnvironmentVariables($scope).GetEnumerator()) {
            $name = [string]$entry.Key
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }

            $canonicalName = if ([string]::Equals($name, 'PATH', [System.StringComparison]::OrdinalIgnoreCase)) { 'Path' } else { $name }
            $value = [string]$entry.Value
            if ([string]::Equals($canonicalName, 'Path', [System.StringComparison]::OrdinalIgnoreCase)) {
                foreach ($segment in ($value -split ';')) {
                    $trimmedSegment = $segment.Trim()
                    if ([string]::IsNullOrWhiteSpace($trimmedSegment)) {
                        continue
                    }

                    if ($seenPathSegments.Add($trimmedSegment)) {
                        [void]$pathSegments.Add($trimmedSegment)
                    }
                }
            }
            else {
                $environment[$canonicalName] = $value
            }
        }
    }

    if ($pathSegments.Count -gt 0) {
        $environment['Path'] = ($pathSegments -join ';')
    }

    return $environment
}

function ConvertTo-RSQuotedProcessArgument {
    param(
        [AllowNull()]
        [string]$Argument
    )

    if ($null -eq $Argument -or $Argument.Length -eq 0) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }

        if ($character -eq '"') {
            if ($backslashCount -gt 0) {
                [void]$builder.Append(('\' * ($backslashCount * 2)))
                $backslashCount = 0
            }

            [void]$builder.Append('\"')
            continue
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append(('\' * $backslashCount))
            $backslashCount = 0
        }

        [void]$builder.Append($character)
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append(('\' * ($backslashCount * 2)))
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

function Set-RSProcessArguments {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo,

        [string[]]$Arguments = @()
    )

    $argumentListProperty = $ProcessStartInfo.PSObject.Properties['ArgumentList']
    if ($argumentListProperty -and $null -ne $ProcessStartInfo.ArgumentList) {
        foreach ($argument in $Arguments) {
            [void]$ProcessStartInfo.ArgumentList.Add($argument)
        }
        return
    }

    $ProcessStartInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-RSQuotedProcessArgument $_ }) -join ' ')
}

function Get-RSProcessEnvironmentDictionary {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo
    )

    $environmentProperty = $ProcessStartInfo.PSObject.Properties['Environment']
    if ($environmentProperty -and $null -ne $ProcessStartInfo.Environment) {
        return ,$ProcessStartInfo.Environment
    }

    $environmentVariablesProperty = $ProcessStartInfo.PSObject.Properties['EnvironmentVariables']
    if ($environmentVariablesProperty -and $null -ne $ProcessStartInfo.EnvironmentVariables) {
        return ,$ProcessStartInfo.EnvironmentVariables
    }

    return ,$ProcessStartInfo.EnvironmentVariables
}

function Set-RSProcessEnvironmentValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$EnvironmentDictionary,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [string]$Value
    )

    if ($null -ne $Value) {
        $EnvironmentDictionary[$Name] = $Value
        return
    }

    if ($EnvironmentDictionary.Contains($Name)) {
        [void]$EnvironmentDictionary.Remove($Name)
    }
}

function New-RSProcessStartInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = (Get-Location).Path,

        [hashtable]$AdditionalEnvironment = @{}
    )

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $FilePath
    Set-RSProcessArguments -ProcessStartInfo $psi -Arguments $Arguments

    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $environmentDictionary = Get-RSProcessEnvironmentDictionary -ProcessStartInfo $psi
    $environmentDictionary.Clear()
    foreach ($entry in (Get-RSNormalizedEnvironmentMap).GetEnumerator()) {
        Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $entry.Key -Value $entry.Value
    }

    foreach ($entry in $AdditionalEnvironment.GetEnumerator()) {
        $name = [string]$entry.Key
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($null -eq $entry.Value) {
            Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $name -Value $null
            continue
        }

        Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $name -Value ([string]$entry.Value)
    }

    return $psi
}
