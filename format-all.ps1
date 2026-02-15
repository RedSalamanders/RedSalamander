# Format all C++ files, excluding build/package directories
Get-ChildItem -Recurse -Include *.cpp,*.h |
    Where-Object { 
        $_.FullName -notmatch '\\(packages|vcpkg_installed|\.vs|Debug|Release|x64|out|\.git)\\' 
    } |
    ForEach-Object {
        Write-Host "Formatting: $($_.Name)" -ForegroundColor Cyan
        clang-format -i --style=file $_.FullName
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "  ✓ Success" -ForegroundColor Green
        } else {
            Write-Host "  ✗ Failed" -ForegroundColor Red
        }
    }

Write-Host "`nFormatting complete!" -ForegroundColor Yellow
Write-Host "Review changes with: git diff" -ForegroundColor Yellow