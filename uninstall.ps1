$tasks=@('CyberPlanBootMailer-Logon','CyberPlanBootMailer-Review')
foreach($name in $tasks){ if(Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue){ Unregister-ScheduledTask -TaskName $name -Confirm:$false; Write-Host "已移除计划任务: $name" -ForegroundColor Yellow } }
$desktop=[Environment]::GetFolderPath('Desktop')
$programs=[Environment]::GetFolderPath('Programs')
foreach($dir in @($desktop,$programs)){
    if($dir){
        $lnk=Join-Path $dir 'CyberPlan Boot Mailer.lnk'
        if(Test-Path -LiteralPath $lnk){Remove-Item -LiteralPath $lnk -Force}
    }
}
Write-Host '开机自启已卸载。邮件配置和日志仍保留在项目目录中。' -ForegroundColor Green