@echo off
set "REPO=learnPS"
set "BRANCH=main"
set "SUBDIR=MergePDF"
set "TARGET=C:\tmp\MergePDF"
powershell -NoProfile -Command ^
  "$zip=Join-Path $env:TEMP ('%REPO%-%BRANCH%.zip');" ^
  "Invoke-WebRequest 'https://github.com/vicefan/%REPO%/archive/refs/heads/%BRANCH%.zip' -OutFile $zip -UseBasicParsing;" ^
  "$work=Join-Path $env:TEMP ('%REPO%-%BRANCH%'); Expand-Archive -Path $zip -DestinationPath $work -Force;" ^
  "$src=Join-Path $work '%REPO%-%BRANCH%\%SUBDIR%'; if(-not (Test-Path $src)){throw 'Not found'}; Remove-Item -Recurse -Force '%TARGET%' -ErrorAction SilentlyContinue; Copy-Item $src '%TARGET%' -Recurse -Force; Remove-Item $zip -Force; Remove-Item -Recurse -Force $work; Write-Host 'Done';"

powershell -NoProfile -ExecutionPolicy Bypass -File "C:\tmp\MergePDF\src\MergePDF.ps1"