# AI排单系统 - 自动守护脚本
# 用法: 在项目目录下 PowerShell 运行 .\watchdog.ps1
# 进程崩溃/退出后 5 秒自动重启，重启记录写入 watchdog.log

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $ScriptDir "watchdog.log"
$Delay = 5

function Write-Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Out-File -Append -FilePath $LogFile -Encoding utf8
    Write-Host "$ts  $msg"
}

Write-Log "===== Watchdog 启动 ====="
Write-Log "工作目录: $ScriptDir"

# ---- API 服务守护 ----
$apiJob = Start-Job -Name "api" -ArgumentList $ScriptDir, $LogFile, $Delay -ScriptBlock {
    param($dir, $log, $delay)
    Set-Location $dir
    while ($true) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$ts  [API服务] 正在启动..." | Out-File -Append -FilePath $log -Encoding utf8
        $proc = Start-Process -FilePath "python" -ArgumentList "run.py" -PassThru -NoNewWindow
        $proc | Wait-Process
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$ts  [API服务] 退出 (code: $($proc.ExitCode))，${delay}s 后重启..." | Out-File -Append -FilePath $log -Encoding utf8
        Start-Sleep -Seconds $delay
    }
}

# ---- 客户机器人守护 ----
$custJob = Start-Job -Name "cust" -ArgumentList $ScriptDir, $LogFile, $Delay -ScriptBlock {
    param($dir, $log, $delay)
    Set-Location $dir
    while ($true) {
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$ts  [客户机器人] 正在启动..." | Out-File -Append -FilePath $log -Encoding utf8
        $proc = Start-Process -FilePath "python" -ArgumentList "customer_agent.py" -PassThru -NoNewWindow
        $proc | Wait-Process
        $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$ts  [客户机器人] 退出 (code: $($proc.ExitCode))，${delay}s 后重启..." | Out-File -Append -FilePath $log -Encoding utf8
        Start-Sleep -Seconds $delay
    }
}

Write-Log "守护已就绪: API服务 + 客户机器人"
Write-Host ""
Write-Host "按 Ctrl+C 停止 watchdog..."
Write-Host ""

try {
    while ($true) {
        Start-Sleep -Seconds 60
        # 心跳检查
        foreach ($j in @($apiJob, $custJob)) {
            if ($j.State -eq "Failed") {
                $err = $j.ChildJobs[0].Error -join "; "
                Write-Log "[紧急] 守护 Job '$($j.Name)' 异常: $err"
            }
        }
    }
}
finally {
    Write-Log "===== Watchdog 停止 ====="
    $apiJob  | Stop-Job -ErrorAction SilentlyContinue | Remove-Job
    $custJob | Stop-Job -ErrorAction SilentlyContinue | Remove-Job
}
