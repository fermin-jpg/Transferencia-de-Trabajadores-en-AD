$f = Get-Content .\CopiarTrabajadores.bat -Encoding Default
$n = ($f | Select-String '^:PS1START').LineNumber
$psCode = $f | Select-Object -Skip $n | Out-String
$errors = $null
$tokens = $null
[System.Management.Automation.Language.Parser]::ParseInput($psCode, [ref]$tokens, [ref]$errors)
if ($errors) {
    Write-Host "SINTAXIS ERROR ENCONTRADO:" -ForegroundColor Red
    $errors | Format-List
} else {
    Write-Host "Sintaxis de PowerShell CORRECTA!" -ForegroundColor Green
}
