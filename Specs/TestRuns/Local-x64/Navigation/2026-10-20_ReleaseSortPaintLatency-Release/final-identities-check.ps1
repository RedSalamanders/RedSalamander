$ErrorActionPreference='Stop'
$root=$PSScriptRoot
$repo='Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation'
$main='Z:/src/RedSalamander'
$accepted=Get-Content (Join-Path $root 'accepted-candidates.json') -Raw|ConvertFrom-Json
$sourceCopies=Get-Content (Join-Path $root 'candidate-source-copies.json') -Raw|ConvertFrom-Json
$code=@('RedSalamander/FolderView.cpp','RedSalamander/FolderView.h','RedSalamander/FolderView.Enumeration.cpp','RedSalamander/FolderView.Rendering.cpp','RedSalamander/SelfTest/Commands/Commands.Tests.ViewCommands.cpp')
$sources=@(foreach($relative in $code){
 $frozen=@($sourceCopies|Where-Object path -eq $relative)
 $isolated=(Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash
 $current=(Get-FileHash -LiteralPath (Join-Path $main $relative)).Hash
 if($frozen.Count-ne1-or$isolated-ne$frozen[0].sha256-or$current-ne$isolated){throw "Final sorting code identity differs: $relative"}
 [ordered]@{path=$relative;sha256=$isolated;main_matches_qualified_isolated_source=$true}
})
Import-Module (Join-Path $repo 'Tools/Modules/Build/ArtifactOperationLock.psm1') -Force
$profiles=@()
foreach($configuration in @('Debug','Release')){
 $lease=Enter-RSArtifactOperationLock -RepoRoot $repo -Operation 'sorting-final-output-readback' -Scope @{configuration=$configuration;platform='x64'}
 try{
  $phase=@($accepted.blocks|Where-Object configuration -eq $configuration|ForEach-Object phase)[0]
  $receipt=Get-Content (Join-Path $root "$phase/build-receipt.json") -Raw|ConvertFrom-Json
  $live=Get-Content (Join-Path $repo ".build/x64/$configuration/build-receipt.json") -Raw|ConvertFrom-Json
  if($receipt.receipt_id-ne$live.receipt_id-or!$receipt.tests_enabled-or$receipt.exit_code-ne0){throw 'Final build receipt differs'}
  $snapshot=Get-Content (Join-Path $root "$phase/workspace-final.json") -Raw|ConvertFrom-Json
  if($snapshot.snapshot_id-ne$receipt.source_snapshot_id){throw 'Test final source snapshot differs from compiled source'}
  foreach($source in $sources){
   $input=@($snapshot.files|Where-Object path -eq $source.path)
   if($input.Count-ne1-or$input[0].worktree_sha256-ne$source.sha256){throw 'Compiled sorting input differs'}
  }
  $exe=@($receipt.outputs|Where-Object relative_path -eq ".build/x64/$configuration/RedSalamander.exe")
  if($exe.Count-ne1){throw 'Application executable receipt missing'}
  $paths=@($exe[0].runtime_closure)+$exe[0].relative_path|Sort-Object -Unique
  $outputs=@(foreach($relative in $paths){
   $record=@($receipt.outputs|Where-Object relative_path -eq $relative)
   if($record.Count-ne1){throw 'Application runtime closure artifact missing'}
   $actual=(Get-FileHash -LiteralPath (Join-Path $repo $relative)).Hash
   if($actual-ne$record[0].sha256){throw "Current application artifact changed: $relative"}
   [ordered]@{path=$relative;sha256=$actual;bytes=$record[0].size_bytes}
  })
  $profiles += [ordered]@{configuration=$configuration;platform='x64';receipt_id=$receipt.receipt_id;source_snapshot_id=$receipt.source_snapshot_id;toolchain_identity_digest=$receipt.toolchain_identity_digest;runtime_dependency_digest=$receipt.runtime_dependency_digest;tests_enabled=$true;application_runtime_closure=$outputs}
 }finally{Exit-RSArtifactOperationLock -Lock $lease}
}
[ordered]@{verified_utc=[DateTime]::UtcNow.ToString('o');status='passed';sources=$sources;profiles=$profiles;scope='Exact compiled C++ input and current application/runtime output readback in the isolated checkout under each profile artifact lock. Main sorting C++ bytes match. Later plan/evidence closeout is documentation-only; the frozen qualification base excludes unrelated ViewerSpace changes.'}|ConvertTo-Json -Depth 10|Set-Content (Join-Path $root 'final-identities.json')
[pscustomobject]@{status='passed';profiles=$profiles.Count;matching_cpp_inputs=$sources.Count}
