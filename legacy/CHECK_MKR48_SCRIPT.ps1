#requires -version 5.1
param(
    [Parameter(Mandatory=$true)]
    [string]$ScriptPath
)

if(-not (Test-Path -LiteralPath $ScriptPath)){
    Write-Host '[ERROR] Main PowerShell script not found.' -ForegroundColor Red
    exit 2
}

$tokens = $null
$parseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    $ScriptPath,
    [ref]$tokens,
    [ref]$parseErrors
)

if($parseErrors.Count -gt 0){
    Write-Host '[ERROR] PowerShell script syntax check failed.' -ForegroundColor Red
    foreach($parseError in $parseErrors){
        Write-Host (' - ' + $parseError.Message) -ForegroundColor Red
    }
    exit 2
}

Write-Host '[OK] PowerShell 5.1 script syntax check passed.' -ForegroundColor Green
exit 0
