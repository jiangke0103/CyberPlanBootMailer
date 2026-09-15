param(
    [ValidateSet('Morning','Review')]
    [string]$Mode = 'Morning',
    [int]$PopupTimeout = 180,
    [switch]$Force
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'common.ps1')
$root=Split-Path -Parent $PSScriptRoot
$today=Get-Date -Format 'yyyy-MM-dd'
$weekday=(Get-Date).DayOfWeek
$dataRoot=Join-Path $root 'data\每日任务'
$stateDir=Join-Path $dataRoot '.state'
$logDir=Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $stateDir,$logDir | Out-Null
$stateFile=Join-Path $stateDir ("{0}_{1}.done" -f $Mode.ToLower(),$today)
$logFile=Join-Path $logDir 'service.log'
if ((Test-Path -LiteralPath $stateFile) -and -not $Force) {
    Add-Content -LiteralPath $logFile -Value ("[{0}] {1} skipped: already shown" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Mode) -Encoding utf8
    exit 0
}
if ($Mode -eq 'Review' -and $weekday -eq 'Monday') { exit 0 }
$python=Resolve-ProjectPython
$generator=Join-Path $root 'src\generate_daily_task.py'
$dayDir=Join-Path $dataRoot $today
$taskMd=Join-Path $dayDir ("{0}_今日任务.md" -f $today)
$taskDocx=Join-Path $dayDir ("{0}_今日任务.docx" -f $today)
$reviewMd=Join-Path $dayDir ("{0}_晚间复盘.md" -f $today)
$reviewDocx=Join-Path $dayDir ("{0}_晚间复盘.docx" -f $today)
if (-not (Test-Path -LiteralPath $taskDocx) -or -not (Test-Path -LiteralPath $reviewDocx)) {
    & $python $generator --mode both
    if ($LASTEXITCODE -ne 0) { Add-Content -LiteralPath $logFile -Value ("[{0}] generator failed exit={1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$LASTEXITCODE) -Encoding utf8 }
}
if (-not (Test-Path -LiteralPath $taskMd)) { throw "今日任务文件不存在: $taskMd" }
$lines=Get-Content -LiteralPath $taskMd -Encoding utf8
$phaseLine=($lines | Where-Object {$_ -match '^> 阶段：' } | Select-Object -First 1)
$resultLines=@()
$inResults=$false
foreach($line in $lines){
    if($line -match '^## 今日三个结果'){$inResults=$true;continue}
    if($inResults -and $line -match '^## '){break}
    if($inResults -and $line -match '^\d+\.\s+'){$resultLines+=$line.Trim()}
}
$taskRows=@()
foreach($line in $lines){
    if($line -match '^\|\s*\d{2}:\d{2}-\d{2}:\d{2}\s*\|'){
        $cols=$line.Trim().Trim('|').Split('|')|ForEach-Object{$_.Trim()}
        if($cols.Count -ge 4){$taskRows+=[pscustomobject]@{Time=$cols[0];Task=$cols[1];How=$cols[2];Output=$cols[3]}}
    }
    if($taskRows.Count -ge 3){break}
}
if($weekday -ne 'Monday'){
    try { & (Join-Path $PSScriptRoot 'send_email.ps1') -Mode $Mode -Day $today | Out-Null }
    catch { Add-Content -LiteralPath $logFile -Value ("[{0}] email failed: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding utf8 }
}

$shell=New-Object -ComObject WScript.Shell
if($Mode -eq 'Morning'){
    if($weekday -eq 'Monday'){
        $title='网络安全就业计划'
        $message="今天是周一，完全休息。`n`n不学习、不做 CTF、不安排任务。`n`n周二再恢复计划。"
        $result=$shell.Popup($message,$PopupTimeout,$title,64)
        $openPath=$null
    }
    else {
        $title="今日网络安全计划 $today"
        $summary=@()
        if($phaseLine){$summary+=$phaseLine.TrimStart('> ').Trim()}
        if($resultLines.Count -gt 0){$summary+='';$summary+='今日三个结果：';$summary+=$resultLines}
        if($taskRows.Count -gt 0){
            $summary+='';$summary+='前三步：'
            foreach($row in $taskRows){$summary+=("{0}  {1}  →  {2}" -f $row.Time,$row.Task,$row.Output)}
            $summary+='';$summary+=("现在先做：{0}" -f $taskRows[0].Task);$summary+=("做法：{0}" -f $taskRows[0].How)
        }
        $summary+='';$summary+='单击 是 打开今天的 Word 任务表。'
        $message=$summary -join "`n"
        $result=$shell.Popup($message,$PopupTimeout,$title,4 -bor 64)
        $openPath=$taskDocx
    }
}
else {
    $title="网络安全计划复盘提醒 $today"
    $message="21:40 开始复盘，20 分钟内完成。`n`n必须回答：`n1. 今日完成`n2. 今日产出`n3. 最大卡点`n4. 明天三件事`n5. 今日得分/10`n`n单击 是 打开 Word 复盘表。"
    $result=$shell.Popup($message,$PopupTimeout,$title,4 -bor 64)
    $openPath=$reviewDocx
}
if($result -eq 6 -and $openPath -and (Test-Path -LiteralPath $openPath)){Start-Process -FilePath $openPath}
Set-Content -LiteralPath $stateFile -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -Encoding utf8
Add-Content -LiteralPath $logFile -Value ("[{0}] {1} shown, result={2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Mode,$result) -Encoding utf8
