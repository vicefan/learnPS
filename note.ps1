.\SettingModules.ps1

$xlsx_path = Join-Path $env:USERPROFILE 'Desktop\tst.xlsx'

$x = Get-SheetNames -Path $xlsx_path
$x, ($x.GetType().FullName)

try {
    $SheetNames = Get-SheetNames -Path $xlsx_path
} catch {
    Write-Error "Failed to get sheet names: $_"
    return
}

Write-Host "Available sheets:"
for ($i = 0; $i -lt $SheetNames.Count; $i++) {
    Write-Host "[$($i+1)] $($SheetNames[$i])"
}

# 예: 첫 번째 시트의 상위 1 행 출력
$firstSheet = $SheetNames[0]
$data = Read-Xlsx -Path $xlsx_path -SheetName $firstSheet
$data | Select-Object -First 1 | Format-Table -AutoSize