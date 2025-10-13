@echo off
set "REPO=learnPS"
set "BRANCH=main"
set "SUBDIR=MergePDF"
set "TARGET=%APPDATA%\MergePDF"

rem 1) src 가져오기
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$zip=Join-Path $env:TEMP ('%REPO%-%BRANCH%.zip');" ^
  "Invoke-WebRequest 'https://github.com/vicefan/%REPO%/archive/refs/heads/%BRANCH%.zip' -OutFile $zip -UseBasicParsing;" ^
  "$work=Join-Path $env:TEMP ('%REPO%-%BRANCH%'); Expand-Archive -Path $zip -DestinationPath $work -Force;" ^
  "$src=Join-Path $work '%REPO%-%BRANCH%\%SUBDIR%'; if(-not (Test-Path $src)){throw 'Not found'}; Remove-Item -Recurse -Force '%TARGET%' -ErrorAction SilentlyContinue; Copy-Item $src '%TARGET%' -Recurse -Force; Remove-Item $zip -Force; Remove-Item -Recurse -Force $work;"

if errorlevel 1 (
  echo download or extract failed
  pause
  exit /b 1
)

rem 2) GUI 실행
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -FilePath 'powershell' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""C:\tmp\MergePDF\src\MergePDF.ps1""' -WindowStyle Hidden"

exit /b 0