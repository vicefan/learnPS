@echo off
rem --- 관리자 권한 확인 및 승격 (없으면 UAC로 현재 배치 다시 실행) ---
powershell -NoProfile -Command "if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','\"%~f0\" %*' -Verb RunAs; exit 0 }"

set "REPO=learnPS"
set "BRANCH=main"
set "TARGET=%APPDATA%\MergePDF"

rem 1) repo 전체 가져오기 (다운로드 -> 압축 해제 -> 복사) with retry for Expand-Archive
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$zip = Join-Path $env:TEMP ('%REPO%-%BRANCH%-{0}.zip' -f ([guid]::NewGuid()));" ^
  "try {" ^
  "  Invoke-WebRequest 'https://github.com/vicefan/%REPO%/archive/refs/heads/%BRANCH%.zip' -OutFile $zip -UseBasicParsing -ErrorAction Stop;" ^
  "  $work = Join-Path $env:TEMP ('%REPO%-%BRANCH%'); if (Test-Path $work) { Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue };" ^
  "  $attempt = 0; $max = 6; while ($true) { try { Expand-Archive -Path $zip -DestinationPath $work -Force -ErrorAction Stop; break } catch { $attempt++; if ($attempt -ge $max) { throw $_ } ; Start-Sleep -Seconds 1 } };" ^
  "  $repoRoot = Join-Path $work '%REPO%-%BRANCH%'; if (-not (Test-Path $repoRoot)) { throw 'Not found' };" ^
  "  Remove-Item -Recurse -Force '%TARGET%' -ErrorAction SilentlyContinue; New-Item -ItemType Directory -Path '%TARGET%' -Force | Out-Null;" ^
  "  Copy-Item -Path (Join-Path $repoRoot '*') -Destination '%TARGET%' -Recurse -Force;" ^
  "  Remove-Item -Force $zip -ErrorAction SilentlyContinue; Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue;" ^
  "  exit 0" ^
  "} catch { Write-Error $_; Remove-Item -Force $zip -ErrorAction SilentlyContinue; Remove-Item -Recurse -Force $work -ErrorAction SilentlyContinue; exit 1 }"

if errorlevel 1 (
  echo download or extract failed
  pause
  exit /b 1
)

rem 2) GUI 실행 (TARGET에 복사된 repo 안의 MergePDF\src\MergePDF.ps1 실행)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -FilePath 'powershell' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""%TARGET%\MergePDF\src\MergePDF.ps1""' -WindowStyle Normal"

exit /b 0