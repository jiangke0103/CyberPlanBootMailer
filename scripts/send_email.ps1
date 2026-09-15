param(
    [ValidateSet('Morning','Review','Test')]
    [string]$Mode = 'Morning',
    [string]$Day = (Get-Date -Format 'yyyy-MM-dd')
)
$ErrorActionPreference='Stop'
$systemModules='C:\Windows\System32\WindowsPowerShell\v1.0\Modules'
if($env:PSModulePath -notlike ('*'+$systemModules+'*')){$env:PSModulePath=$systemModules+';'+$env:PSModulePath}
. (Join-Path $PSScriptRoot 'common.ps1')
$root=Split-Path -Parent $PSScriptRoot
$configPath=Join-Path $root 'config\email_config.json'
$secretPath=Join-Path $root 'config\email_secret.txt'
$python=Resolve-ProjectPython
$sender=Join-Path $root 'src\send_daily_email.py'
$logDir=Join-Path $root 'logs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
if (-not (Test-Path -LiteralPath $configPath) -or -not (Test-Path -LiteralPath $secretPath)) {
    Add-Content -LiteralPath (Join-Path $logDir 'email.log') -Value ("[{0}] email skipped: config missing" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding utf8
    exit 0
}
$secure=Get-Content -LiteralPath $secretPath -Raw | ConvertTo-SecureString
$bstr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
try {
    $plain=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    $env:CYBER_SMTP_PASSWORD=$plain
    & $python $sender --mode $Mode --date $Day
    exit $LASTEXITCODE
}
finally {
    Remove-Item Env:\CYBER_SMTP_PASSWORD -ErrorAction SilentlyContinue
    if ($bstr -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}
