@echo off
rem --- 관리자 권한 확인 및 승격 (없으면 UAC로 현재 배치 다시 실행) ---
powershell -NoProfile -Command "if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Start-Process -FilePath 'cmd.exe' -ArgumentList '/c','\"%~f0\" %*' -Verb RunAs; exit 0 }"

set "REPO=learnPS"
set "BRANCH=main"
set "SUBDIR=MergePDF"
set "TARGET=%USERPROFILE%\AppData\Local\Temp\%SUBDIR%"

 :: BatchGotAdmin
 :-------------------------------------
 REM  --> Check for permissions
 >nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
 if '%errorlevel%' NEQ '0' (
     echo Requesting administrative privileges...
     goto UACPrompt
 ) else ( goto gotAdmin )

:UACPrompt
     echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
     echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
     exit /B

:gotAdmin
     if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
     pushd "%CD%"
     CD /D "%~dp0"
 :--------------------------------------

rem 파워셸 스크립트 실행 권한 설정
powershell -NoProfile -Command "Set-ExecutionPolicy RemoteSigned"

rem 1) ZIP 다운로드 및 서브폴더 추출
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$zip=Join-Path $env:TEMP ('%REPO%-%BRANCH%.zip');" ^
  "Invoke-WebRequest 'https://github.com/vicefan/%REPO%/archive/refs/heads/%BRANCH%.zip' -OutFile $zip -UseBasicParsing;" ^
  "$work=Join-Path $env:TEMP ('%REPO%-%BRANCH%'); Expand-Archive -Path $zip -DestinationPath $work -Force;" ^
  "$src=Join-Path $work '%REPO%-%BRANCH%\%SUBDIR%'; if(-not (Test-Path $src)){throw 'Not found'}; Remove-Item -Recurse -Force '%TARGET%' -ErrorAction SilentlyContinue; Copy-Item $src '%TARGET%' -Recurse -Force; Remove-Item $zip -Force; Remove-Item -Recurse -Force $work;"

if errorlevel 1 (
  echo Download or Extract Failed
  pause
  exit /b 1
)

rem 2) GUI 스크립트 숨김으로 실행 (새로운 PowerShell 프로세스, 콘솔 창 없음)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Process -FilePath 'powershell' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File ""C:\tmp\MergePDF\src\MergePDF.ps1""' -WindowStyle Hidden"

exit /b 0