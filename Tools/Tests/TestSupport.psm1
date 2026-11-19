Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepositoryRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Invoke-RSTestFixtureNative {
    <# Explicit owned fixture transport; not a production command shim. No scheduled caller supplies stdin. #>
    param([Parameter(Mandatory)][string]$FilePath, [string[]]$Arguments = @(),
        [string]$WorkingDirectory = (Get-Location).ProviderPath, [System.Text.Encoding]$StandardOutputEncoding)
    if ($MyInvocation.ExpectingInput) { throw 'Scheduled native fixture stdin requires its explicit reviewed transport.' }
    Import-Module (Join-Path $PSScriptRoot '../Modules/Build/ProcessStreaming.psm1') -Scope Local -ErrorAction Stop
    $events=[Collections.Generic.List[object]]::new()
    $capture={param($Line,$IsError)$events.Add([pscustomobject]@{Text=[string]$Line;IsError=[bool]$IsError})}.GetNewClosure()
    $exitCode=Invoke-RSStreamingProcess -FilePath $FilePath -Arguments $Arguments -WorkingDirectory $WorkingDirectory `
        -StandardOutputEncoding $StandardOutputEncoding -OutputLineCallback $capture
    foreach($event in $events){
        if($event.IsError){Write-Error -Message $event.Text -ErrorId NativeCommandError -Category NotSpecified -ErrorAction Continue}
        else{Write-Output ([string]$event.Text)}
    }
    $global:LASTEXITCODE=[int]$exitCode
}

function Invoke-RSTestFixtureGit {
    <# Git is the explicit UTF-8 specialization of the owned native fixture transport. #>
    param([string[]]$Arguments)
    if ($MyInvocation.ExpectingInput) { throw 'Scheduled Git fixture stdin requires its explicit reviewed transport.' }
    Invoke-RSTestFixtureNative -FilePath 'git' -Arguments $Arguments -WorkingDirectory (Get-Location).ProviderPath `
        -StandardOutputEncoding ([Text.UTF8Encoding]::new($false))
}

function New-RSTestFormatterFixture {
    <# Test-only native CUI formatter. Reuses the existing managed bootstrapper compilation pattern;
       production preview acquires no compiler or engine dependency. No source formatting occurs. #>
    param([string]$Root = '')
    Import-Module (Join-Path $PSScriptRoot '../Modules/Testing/TestSandbox.psm1') -Scope Local -ErrorAction Stop
    $context=New-RSTestRunContext -RepoRoot $script:RepositoryRoot
    if([string]::IsNullOrWhiteSpace($Root)){
        $Root=New-RSTestSandboxScratchDirectory -RepoRoot $script:RepositoryRoot -Harness 'ps' -Case ('cf-'+[guid]::NewGuid().ToString('N'))
    }
    $Root=Assert-RSTestSandboxContainedPath -TestRoot $context.TestRoot -Path $Root -RequireDescendant
    if((Test-Path -LiteralPath $Root) -and @(Get-ChildItem -LiteralPath $Root -Force).Count){throw 'Native formatter fixture requires an empty owned child.'}
    $source=Join-Path $Root 'Formatter.cs';$compile=Join-Path $Root 'Compile.ps1';$exe=Join-Path $Root 'native formatter.exe'
    $log=Join-Path $Root 'formatter.args.jsonl';$audit=Join-Path $Root 'formatter.audit.jsonl'
    $compileLog=Join-Path $Root 'compile.stdout.log';$compileError=Join-Path $Root 'compile.stderr.log'
    foreach($path in @($source,$compile,$exe,($exe+'.config'),$log,$audit,$compileLog,$compileError)){
        if($path.Length -ge 240){throw 'Managed native formatter fixture exceeds its conservative loader path budget.'}
    }
    [void][IO.Directory]::CreateDirectory($Root)
    $body=@'
using System;
using System.IO;
using System.Text;
using System.Runtime.InteropServices;
class FormatterFixtureProgram {
    [DllImport("kernel32.dll")] static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindowVisible(IntPtr window);
    static string Json(string value) { return "\""+value.Replace("\\","\\\\").Replace("\"","\\\"").Replace("\r","\\r").Replace("\n","\\n").Replace("\t","\\t")+"\""; }
    static string ArgsJson(string[] args) { string[] values=new string[args.Length]; for(int i=0;i<args.Length;i++)values[i]=Json(args[i]); return "["+String.Join(",",values)+"]"; }
    // Concurrent fixture processes serialize their appends instead of failing on a sharing violation.
    static void Append(string path,string line) {
        byte[] bytes=new UTF8Encoding(false).GetBytes(line+Environment.NewLine);
        for(int attempt=0;;attempt++){
            try{ using(FileStream stream=new FileStream(path,FileMode.Append,FileAccess.Write,FileShare.Read)){stream.Write(bytes,0,bytes.Length);} return; }
            catch(IOException){ if(attempt>=400)throw; System.Threading.Thread.Sleep(5); }
        }
    }
    static int Main(string[] args) {
        Console.OutputEncoding=new UTF8Encoding(false);
        string root=Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
        string audit=Path.Combine(root,"formatter.audit.jsonl");
        IntPtr window=GetConsoleWindow();
        string record="{\"processId\":"+System.Diagnostics.Process.GetCurrentProcess().Id+",\"consolePresent\":"+(window!=IntPtr.Zero?"true":"false")+",\"consoleVisible\":"+(window!=IntPtr.Zero&&IsWindowVisible(window)?"true":"false")+",\"workingDirectory\":"+Json(Environment.CurrentDirectory)+",\"arguments\":"+ArgsJson(args)+"}";
        Append(audit,record);
        bool version=args.Length==1&&args[0]=="--version";
        string custom=Environment.GetEnvironmentVariable("RS_FORMATTER_STDOUT");
        if(custom!=null)Console.Out.WriteLine(custom);
        else if(version)Console.Out.WriteLine("clang-format version "+(Environment.GetEnvironmentVariable("RS_FORMATTER_VERSION")??"22.1.3"));
        if(!version){ string log=Environment.GetEnvironmentVariable("RS_FORMATTER_LOG")??Path.Combine(root,"formatter.args.jsonl"); Append(log,ArgsJson(args)); }
        // A --dry-run check reports each file after "--" that contains the marker the way clang-format
        // --Werror reports drift. The probe still never rewrites a file.
        int violations=0;
        string marker=Environment.GetEnvironmentVariable("RS_FORMATTER_UNFORMATTED");
        int files=Array.IndexOf(args,"--");
        if(!version&&!String.IsNullOrEmpty(marker)&&files>=0&&Array.IndexOf(args,"--dry-run")>=0){
            for(int i=files+1;i<args.Length;i++){
                if(File.Exists(args[i])&&File.ReadAllText(args[i]).Contains(marker)){Console.Error.WriteLine(args[i]+":1:1: error: code should be clang-formatted [-Wclang-format-violations]");violations++;}
            }
        }
        int count=0;Int32.TryParse(Environment.GetEnvironmentVariable("RS_FORMATTER_STDERR_COUNT"),out count);
        if(count<0||count>262144)return 98;
        if(count>0)Console.Error.Write(new String('E',count));
        int code=0;Int32.TryParse(Environment.GetEnvironmentVariable("RS_FORMATTER_EXIT"),out code);
        return violations>0?1:code;
    }
}
'@
    [IO.File]::WriteAllText($source,$body,[Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($compile,'param($SourcePath,$OutputPath); $ErrorActionPreference=''Stop''; Add-Type -Path $SourcePath -OutputAssembly $OutputPath -OutputType ConsoleApplication',[Text.UTF8Encoding]::new($false))
    $records=@(Invoke-RSTestFixtureNative -FilePath 'powershell.exe' -Arguments @('-NoProfile','-NonInteractive','-File',$compile,$source,$exe) -WorkingDirectory $Root 2>&1)
    $code=$global:LASTEXITCODE
    @($records|Where-Object {$_ -isnot [Management.Automation.ErrorRecord]}|ForEach-Object{[string]$_})|Set-Content -LiteralPath $compileLog
    @($records|Where-Object {$_ -is [Management.Automation.ErrorRecord]}|ForEach-Object{[string]$_})|Set-Content -LiteralPath $compileError
    if($code -ne 0 -or -not (Test-Path -LiteralPath $exe -PathType Leaf)){throw "Owned native formatter compilation failed ($code); see $compileError"}
    $bytes=[IO.File]::ReadAllBytes($exe);$pe=[BitConverter]::ToInt32($bytes,0x3c)
    if([BitConverter]::ToUInt16($bytes,$pe+24+68) -ne 3){throw 'Native formatter fixture must be CUI subsystem3.'}
    [pscustomobject]@{Root=$Root;Executable=$exe;LogPath=$log;AuditPath=$audit;SourcePath=$source;CompileLog=$compileLog;CompileErrorLog=$compileError;ExecutableSHA256=(Get-FileHash $exe).Hash;Subsystem=3}
}


function Get-RSText {
    <#
    .SYNOPSIS
        Reads a repository-relative text file for a source-contract test.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Import-Module (Join-Path $PSScriptRoot '../Modules/Build/RedMonitorDependency.psm1') -Force
    $resolved=Resolve-RSRedMonitorTestInput -RepoRoot $script:RepositoryRoot -Path $Path
    return Get-Content -LiteralPath $resolved -Raw
}

function Assert-RSEqual {
    <#
    .SYNOPSIS
        Raises a focused test error when two scalar values differ.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function Assert-RSSequenceEqual {
    <#
    .SYNOPSIS
        Raises a focused test error when two ordered sequences differ.
    #>
    [CmdletBinding()]
    param(
        [AllowNull()]
        [object[]]$Actual,

        [AllowNull()]
        [object[]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $actualItems = @($Actual)
    $expectedItems = @($Expected)
    if ($actualItems.Count -ne $expectedItems.Count) {
        throw "$Message Expected $($expectedItems.Count) item(s) but got $($actualItems.Count)."
    }

    for ($index = 0; $index -lt $expectedItems.Count; $index++) {
        if ($actualItems[$index] -ne $expectedItems[$index]) {
            throw "$Message Item $index expected '$($expectedItems[$index])' but got '$($actualItems[$index])'."
        }
    }
}

function Invoke-RSTestBlock {
    <#
    .SYNOPSIS
        Executes a positive test action and reports terminating or promoted errors explicitly.
    .DESCRIPTION
        Returns the action output and failure record so positive success assertions do not
        depend on Pester's version-specific `Should Not Throw` behavior.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $previousErrorActionPreference = $ErrorActionPreference
    $output = [System.Collections.Generic.List[object]]::new()
    try {
        $ErrorActionPreference = 'Stop'
        foreach ($item in (& $ScriptBlock)) {
            $output.Add($item)
        }
        return [pscustomobject]@{
            Succeeded = $true
            Output = @($output.ToArray())
            Error = $null
        }
    } catch {
        return [pscustomobject]@{
            Succeeded = $false
            Output = @($output.ToArray())
            Error = $_
        }
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function New-RSValidationGitFixture {
    <#
    .SYNOPSIS
        Creates the shared dirty/staged/untracked/delete/rename Git fixture used by
        Operation Startrail contract tests.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Root)

    $fixtureRoot = [IO.Path]::GetFullPath($Root)
    if (Test-Path -LiteralPath $fixtureRoot) {
        throw "Validation fixture root already exists: $fixtureRoot"
    }
    [void](New-Item -ItemType Directory -Path $fixtureRoot)
    [void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'path with spaces'))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'dirty.txt'), "base`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'delete.txt'), "delete`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'rename.txt'), "rename`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'path with spaces\tracked.txt'), "space`n", [Text.UTF8Encoding]::new($false))

    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'init'; '--quiet')
    if ($LASTEXITCODE -ne 0) { throw 'Unable to initialize validation Git fixture.' }
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'config'; 'user.email'; 'startrail-fixture@invalid.example')
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'config'; 'user.name'; 'Startrail Fixture')
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'add'; '--'; '.')
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'commit'; '--quiet'; '-m'; 'fixture base')
    if ($LASTEXITCODE -ne 0) { throw 'Unable to commit validation Git fixture base.' }

    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'dirty.txt'), "dirty`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'staged.txt'), "staged`n", [Text.UTF8Encoding]::new($false))
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'add'; '--'; 'staged.txt')
    [IO.File]::WriteAllText((Join-Path $fixtureRoot 'untracked.txt'), "untracked`n", [Text.UTF8Encoding]::new($false))
    Remove-Item -LiteralPath (Join-Path $fixtureRoot 'delete.txt')
    Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'mv'; '--'; 'rename.txt'; 'renamed.txt')
    if ($LASTEXITCODE -ne 0) { throw 'Unable to create validation Git fixture rename.' }

    return [pscustomobject]@{
        Root = $fixtureRoot
        DirtyPath = 'dirty.txt'
        StagedPath = 'staged.txt'
        UntrackedPath = 'untracked.txt'
        DeletedPath = 'delete.txt'
        RenamedOldPath = 'rename.txt'
        RenamedNewPath = 'renamed.txt'
        SpacePath = 'path with spaces/tracked.txt'
    }
}

function New-RSTestRunArchiveFixture {
    <#
    .SYNOPSIS
        Allocates a unique, caller-owned archive transaction fixture in an already marked sandbox.
    #>
    param([Parameter(Mandatory = $true)][string]$TestRoot)
    $id = (New-RSTestRunContext -RepoRoot $script:RepositoryRoot -TestRootOverride $TestRoot).RunId
    $source = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $TestRoot "runs\$id") -RequireDescendant
    [void][IO.Directory]::CreateDirectory($source)
    $payload = Join-Path $source 'unique.txt'
    [IO.File]::WriteAllText($payload, 'synthetic archive regression evidence')
    return [pscustomobject]@{
        Root = $TestRoot; Id = $id; Source = $source; Payload = $payload
        Destination = Join-Path $TestRoot "evidence\completed-runs\$id"
        Records = Join-Path $TestRoot "evidence\run-archive-records\$id"
    }
}

Export-ModuleMember -Function @(
    'Invoke-RSTestFixtureNative',
    'Invoke-RSTestFixtureGit',
    'New-RSTestFormatterFixture',
    'Get-RSText',
    'Invoke-RSTestBlock',
    'Assert-RSEqual',
    'Assert-RSSequenceEqual',
    'New-RSValidationGitFixture',
    'New-RSTestRunArchiveFixture'
)
