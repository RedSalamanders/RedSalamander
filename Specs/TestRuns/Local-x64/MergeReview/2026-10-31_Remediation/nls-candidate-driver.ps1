$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$temp = Join-Path $root 'temp'
New-Item -ItemType Directory -Path $temp -Force | Out-Null
$env:TEMP = $temp; $env:TMP = $temp
$sourcePath = 'Z:\src\RedSalamander\Common\FileSystemPathIdentity.cpp'
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash
$head = (& git -C 'Z:\src\RedSalamander' rev-parse HEAD).Trim()
Add-Type -TypeDefinition ([System.IO.File]::ReadAllText((Join-Path $root 'CandidateNlsParityOracle.cs'))) -Language CSharp
$result = [RedSalamanderCandidateNlsOracle]::Run()
$output = [ordered]@{ provenance=[ordered]@{utc=[DateTime]::UtcNow.ToString('o');repoHead=$head;sourcePath=$sourcePath;sourceSha256=$sourceHash;oracleSource=(Join-Path $root 'CandidateNlsParityOracle.cs');powershell=$PSVersionTable.PSVersion.ToString();os=[Environment]::OSVersion.VersionString;temp=$temp;api='LCMapStringEx(LOCALE_NAME_INVARIANT, LCMAP_UPPERCASE, explicit one-unit UTF-16 lengths) plus CompareStringOrdinal(ignoreCase=true, explicit lengths), matching candidate folding';scope='All 65,536 single UTF-16 units with native equivalence classes and accepted-key collision checks, plus required supplementary Deseret/context/emoji controls. This is not a proof for arbitrary multi-unit strings.'}; result=$result }
$json = ConvertTo-Json -InputObject $output -Depth 16
Set-Content -LiteralPath (Join-Path $root 'candidate-result.json') -Value $json -Encoding utf8
$json
if (-not $result.pass) { exit 2 }
