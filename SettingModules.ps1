$ModulePath = ".\modules"
$ModuleFiles = Get-ChildItem -Path $ModulePath -Filter *.psm1
Clear-Host
Write-Host "Found module files: $($ModuleFiles.Name)"
$ENVPATH = $env:PSModulePath -split ';'
$EnvPath = $ENVPATH | Out-GridView -Title "Select a module path" -PassThru
Write-Host "Selected module path: $EnvPath"

foreach ($ModuleFile in $ModuleFiles) {
    $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension($ModuleFile.Name)
    $targetDir = Join-Path -Path $EnvPath -ChildPath $ModuleName

    if (-not (Test-Path -Path $targetDir)) {
        Write-Host "Importing module: $ModuleName from $($ModuleFile.FullName) to $EnvPath"
        New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
        Copy-Item -Path $ModuleFile.FullName -Destination $targetDir -Force
        Write-Host "Module $ModuleName imported successfully."
    } else {
        # 기존 모듈이 존재할 때 사용자에게 재설치 여부 확인
        $answer = Read-Host "Module $ModuleName already exists in $EnvPath. Reinstall (overwrite)? [y/N]"
        if ($answer -match '^[Yy]') {
            try {
                Write-Host "Reinstalling $ModuleName..."
                Remove-Item -Path $targetDir -Recurse -Force -ErrorAction Stop
                New-Item -Path $targetDir -ItemType Directory -Force | Out-Null
                Copy-Item -Path $ModuleFile.FullName -Destination $targetDir -Force
                Write-Host "Module $ModuleName reinstalled successfully."
            } catch {
                # 파서 오류를 피하도록 문자열 포맷팅 사용
                Write-Warning ("Failed to reinstall {0}: {1}" -f $ModuleName, $_.Exception.Message)
            }
        } else {
            Write-Host "Module $ModuleName already exists in $EnvPath. Skipping import."
        }
    }
}

# 최종 결과
Write-Host "`nFinal Result:"
Get-ChildItem -Path $EnvPath -Name | Format-Wide -Column 2