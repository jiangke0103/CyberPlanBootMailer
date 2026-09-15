$ErrorActionPreference='Stop'
$root=$PSScriptRoot
. (Join-Path $root 'scripts\common.ps1')
$python=Find-PythonExe
Write-Host "使用 Python: $python" -ForegroundColor Cyan
& $python -m pip install --disable-pip-version-check --quiet --user -r (Join-Path $root 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'python-docx 安装失败，请检查网络或 Python 环境。' }
$python=Resolve-ProjectPython

$entry=Join-Path $root 'scripts\entrypoint.ps1'
$powershell='C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$user=$env:USERDOMAIN+'\'+$env:USERNAME
$settings=New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
$principal=New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$morningAction=New-ScheduledTaskAction -Execute $powershell -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$entry+'" -Mode Morning')
$reviewAction=New-ScheduledTaskAction -Execute $powershell -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+$entry+'" -Mode Review')
$logonTrigger=New-ScheduledTaskTrigger -AtLogOn -User $user
$logonTrigger.Delay='PT20S'
$reviewTrigger=New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday -At '21:00'

$oldTasks=@('网络安全就业计划-开机提醒','网络安全就业计划-每日任务生成','网络安全就业计划-早间提醒','网络安全就业计划-晚间复盘提醒')
foreach($name in $oldTasks){ if(Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue){ Unregister-ScheduledTask -TaskName $name -Confirm:$false } }

Register-ScheduledTask -TaskName 'CyberPlanBootMailer-Logon' -Description '登录后生成当天计划、发送邮件并弹出桌面通知' -Action $morningAction -Trigger $logonTrigger -Settings $settings -Principal $principal -Force | Out-Null
Register-ScheduledTask -TaskName 'CyberPlanBootMailer-Review' -Description '周二至周日21:00发送复盘提醒邮件并弹出桌面通知' -Action $reviewAction -Trigger $reviewTrigger -Settings $settings -Principal $principal -Force | Out-Null
$pythonw=Join-Path (Split-Path -Parent $python) 'pythonw.exe'
if(-not (Test-Path -LiteralPath $pythonw)){$pythonw=$python}
$ws=New-Object -ComObject WScript.Shell
$desktop=[Environment]::GetFolderPath('Desktop')
$programs=[Environment]::GetFolderPath('Programs')
foreach($dir in @($desktop,$programs)){
    if($dir){
        $lnk=$ws.CreateShortcut((Join-Path $dir 'CyberPlan Boot Mailer.lnk'))
        $lnk.TargetPath=$pythonw
        $lnk.Arguments='"'+ (Join-Path $root 'desktop_app.pyw') + '"'
        $lnk.WorkingDirectory=$root
        $lnk.Description='每日计划、开机自启和邮件桌面应用'
        $lnk.IconLocation=($pythonw+',0')
        $lnk.Save()
    }
}

Write-Host '安装完成。开机登录后自动运行，Codex 关闭也能工作。' -ForegroundColor Green
Write-Host ('邮件配置入口：'+(Join-Path  '配置邮件.lnk'))
Write-Host ('卸载入口：'+(Join-Path  '卸载开机自启.lnk'))