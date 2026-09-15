$ErrorActionPreference='Stop'
$systemModules='C:\Windows\System32\WindowsPowerShell\v1.0\Modules'
if($env:PSModulePath -notlike ('*'+$systemModules+'*')){$env:PSModulePath=$systemModules+';'+$env:PSModulePath}
$root=Split-Path -Parent $PSScriptRoot
$configPath=Join-Path $root 'config\email_config.json'
$secretPath=Join-Path $root 'config\email_secret.txt'

Write-Host '配置每日计划邮件' -ForegroundColor Cyan
Write-Host '注意：这里要填写邮箱的 SMTP 授权码/应用专用密码，不是网页登录密码。'
Write-Host ''
Write-Host '1. QQ邮箱'
Write-Host '2. 163邮箱'
Write-Host '3. Gmail'
Write-Host '4. Outlook/Hotmail'
Write-Host '5. 其他邮箱'
$choice=Read-Host '请选择邮箱类型（1-5）'

switch($choice){
    '1' {$hostName='smtp.qq.com';$port=587;$security='starttls'}
    '2' {$hostName='smtp.163.com';$port=465;$security='ssl'}
    '3' {$hostName='smtp.gmail.com';$port=587;$security='starttls'}
    '4' {$hostName='smtp-mail.outlook.com';$port=587;$security='starttls'}
    default {
        $hostName=Read-Host 'SMTP服务器地址'
        $port=[int](Read-Host 'SMTP端口，例如 587 或 465')
        $security=Read-Host '加密方式：starttls 或 ssl'
    }
}

$sender=(Read-Host '发件邮箱地址').Trim()
if ([string]::IsNullOrWhiteSpace($sender)) { throw '邮箱地址不能为空' }
$recipient=(Read-Host '收件邮箱，直接回车表示发给自己').Trim()
if ([string]::IsNullOrWhiteSpace($recipient)) { $recipient=$sender }
$secure=Read-Host '请输入 SMTP 授权码/应用专用密码' -AsSecureString

$config=[ordered]@{
    provider=$choice
    smtp_host=$hostName
    smtp_port=[int]$port
    security=$security
    username=$sender
    sender=$sender
    recipient=$recipient
}
$json=$config | ConvertTo-Json -Depth 3
$utf8NoBom=New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($configPath,$json,$utf8NoBom)
[IO.File]::WriteAllText($secretPath,($secure | ConvertFrom-SecureString),$utf8NoBom)
Write-Host '配置已保存，正在发送测试邮件……' -ForegroundColor Yellow
& (Join-Path $PSScriptRoot 'send_email.ps1') -Mode Test
if ($LASTEXITCODE -ne 0) {
    Write-Host '测试邮件发送失败，请检查授权码、SMTP服务和网络。' -ForegroundColor Red
    exit 1
}
Write-Host '测试邮件发送成功。以后开机后会发送今日任务，21:00 会发送复盘提醒。' -ForegroundColor Green
