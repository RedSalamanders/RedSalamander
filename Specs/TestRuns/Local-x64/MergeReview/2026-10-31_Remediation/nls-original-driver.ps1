$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$temp = Join-Path $root 'temp'
New-Item -ItemType Directory -Path $temp -Force | Out-Null
$env:TEMP = $temp
$env:TMP = $temp
$sourcePath = 'Z:\src\RedSalamander\Common\FileSystemPathIdentity.cpp'
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$head = (& git -C 'Z:\src\RedSalamander' rev-parse HEAD).Trim()
Add-Type -TypeDefinition ([System.IO.File]::ReadAllText((Join-Path $root 'NlsParityOracle.cs'))) -Language CSharp
$result = [RedSalamanderNlsOracle]::Run()
$output = [ordered]@{
    provenance = [ordered]@{
        utc = [DateTime]::UtcNow.ToString('o')
        repoHead = $head
        identitySource = $sourcePath
        identitySourceSha256 = $sourceHash
        oracleSource = (Join-Path $root 'NlsParityOracle.cs')
        powershell = $PSVersionTable.PSVersion.ToString()
        os = [Environment]::OSVersion.VersionString
        temp = $temp
        api = 'LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE, explicit UTF-16 lengths); CompareStringOrdinal(ignoreCase=true, explicit UTF-16 lengths)'
        scope = 'All 65,536 single UTF-16 code units, including isolated surrogates, plus representative supplementary and contextual strings; not a proof for arbitrary strings.'
    }
    result = $result
}
$json = ConvertTo-Json -InputObject $output -Depth 12
Set-Content -LiteralPath (Join-Path $root 'result.json') -Value $json -Encoding utf8
$json
if (-not $result.pass) { exit 2 }