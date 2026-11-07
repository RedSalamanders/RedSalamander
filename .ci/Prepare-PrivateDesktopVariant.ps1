# Private transformed fixture; the candidate source is never changed.
function Get-PrivateBoundaryVariant {
 param([Parameter(Mandatory)][string]$ModulePath,[Parameter(Mandatory)][ValidateSet('Explicit','Inherited')][string]$Variant)
 $tokens=$null;$errors=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($ModulePath,[ref]$tokens,[ref]$errors)
 if($errors.Count){throw 'Candidate module has parse errors.'}
 $funcs=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Initialize-RSHostedDesktopNative'},$true))
 if($funcs.Count -ne 1){throw 'Candidate initializer identity is ambiguous.'}
 $commands=@($funcs[0].Body.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Add-Type'},$true))
 if($commands.Count -ne 1){throw 'Candidate initializer Add-Type identity is ambiguous.'}
 $strings=@($commands[0].CommandElements|Where-Object {$_ -is [Management.Automation.Language.StringConstantExpressionAst] -and $_.StringConstantType -eq [Management.Automation.Language.StringConstantType]::SingleQuotedHereString})
 if($strings.Count -ne 1){throw 'Candidate initializer source literal is ambiguous.'}
 $original=$strings[0].Value;$code=$original
 $anchor='        private void Run() {'
 $diagnostics=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'DesktopVariantInstrumentation.txt'))
 if(([regex]::Matches($code,[regex]::Escape($anchor))).Count -ne 1){throw 'Candidate worker anchor changed.'}
 $code=$code.Replace($anchor,$diagnostics+"`n"+$anchor)
 $old="                if (!SetThreadDesktop(desktop))`n                    throw new Win32Exception(Marshal.GetLastWin32Error(), `"Cannot attach the hosted foreground thread to Default.`");"
 if($code.Contains("`r`n")){$old=$old.Replace("`n","`r`n")}
 if(([regex]::Matches($code,[regex]::Escape($old))).Count -ne 1){throw 'Candidate desktop attachment anchor changed.'}
 $selected=if($Variant -ceq 'Explicit'){$old}else{'                var inherited=HostedDesktopNative.ReadDesktopContext(0);'+"`n"+'                if (inherited.WindowStation != "WinSta0" || inherited.ThreadDesktop != "Default" || inherited.InputDesktop != "Default" || inherited.WindowStationError != 0 || inherited.ThreadDesktopError != 0 || inherited.InputDesktopError != 0) throw new InvalidOperationException("Inherited hosted worker desktop is not positively WinSta0/Default.");'}
 $replacement='                DiagnosticDesktopAccess("before", desktop);'+"`n"+$selected+"`n"+'                DiagnosticDesktopAccess("after", desktop);'
 $code=$code.Replace($old,$replacement)
 $hash=[Security.Cryptography.SHA256]::Create()
 try {$originalHash=[Convert]::ToHexString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($original)));$compiledHash=[Convert]::ToHexString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($code)))}finally{$hash.Dispose()}
 [pscustomobject]@{Original=$original;Compiled=$code;OriginalSha256=$originalHash;CompiledSha256=$compiledHash;Variant=$Variant;
    Replacements=@(@{anchor=$anchor;replacement=($diagnostics+"`n"+$anchor);purpose='shared worker handle access observation'},@{anchor=$old;replacement=$replacement;purpose='attachment variant with before/after observations'})}
}
