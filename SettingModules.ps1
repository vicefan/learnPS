$ModulePath = ".\modules"
$ModuleFiles = Get-ChildItem -Path $ModulePath -Filter *.psm1
Clear-Host
Write-Host "Found module files: $($ModuleFiles.Name)"
$ENVPATH = $env:PSModulePath -split ';'
$EnvPath = $ENVPATH | Out-GridView -Title "Select a module path" -PassThru
Write-Host "Selected module path: $EnvPath"

foreach ($ModuleFile in $ModuleFiles) {
    $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension($ModuleFile.Name)
    if (-not (Test-Path -Path (Join-Path -Path $EnvPath -ChildPath $ModuleName))) {
        Write-Host "Importing module: $ModuleName from $($ModuleFile.FullName) to $EnvPath"
        New-Item -Path (Join-Path -Path $EnvPath -ChildPath $ModuleName) -ItemType Directory -Force | Out-Null
        Copy-Item -Path $ModuleFile.FullName -Destination (Join-Path -Path $EnvPath -ChildPath $ModuleName) -Force
        Write-Host "Module $ModuleName imported successfully."
    } else {
        Write-Host "Module $ModuleName already exists in $EnvPath. Skipping import."
    }
}

# 최종 결과
Write-Host "`nFinal Result:"
Get-ChildItem -Path $EnvPath -Name | Format-Wide -Column 2