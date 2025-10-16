function Set-Admin {
    <#
    .SYNOPSIS
      관리자 권한이 아니면 동일 스크립트를 관리자 권한으로 재실행하고 현재 프로세스를 종료
      (스크립트에서 맨 처음에 호출하면 자동으로 상승 실행을 보장할 수 있음)
    #>
    param()
    if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) { return $true }

    $psExe = Join-Path $PSHome 'powershell.exe'
    # 호출한 스크립트 파일 경로가 있으면 그 파일을 -File로 재실행, 없으면 새 관리자 셸만 실행
    $callerScript = $MyInvocation.MyCommand.Path

    if ($callerScript) {
        $args = @('-NoProfile','-ExecutionPolicy','Bypass','-File', $callerScript)
    } else {
        $args = @('-NoProfile','-ExecutionPolicy','Bypass')
    }

    try {
        Start-Process -FilePath $psExe -Verb RunAs -ArgumentList $args
        # 관리자 권한으로 재실행 명령을 보냈으므로 현재 프로세스/스크립트는 종료
        exit 0
    } catch {
        Write-Warning ("Failed to elevate: {0}" -f $_.Exception.Message)
        return $false
    }
}

Export-ModuleMember -Function Set-Admin