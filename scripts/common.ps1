function Find-PythonExe {
    $candidates = @()
    if ($env:LOCALAPPDATA) {
        $candidates += Get-ChildItem (Join-Path $env:LOCALAPPDATA 'Programs\Python') -Filter 'python.exe' -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -ExpandProperty FullName
    }
    $candidates += 'C:\Python312\python.exe','C:\Python311\python.exe'
    $cmd = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($cmd) { $candidates += $cmd.Source }
    foreach ($p in ($candidates | Where-Object { $_ } | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    throw '未找到 Python。请先安装 Python 3.11 或更高版本。'
}

function Resolve-ProjectPython {
    $p = Find-PythonExe
    & $p -c 'import docx' 2>$null
    if ($LASTEXITCODE -eq 0) { return $p }
    throw 'Python 已找到，但 python-docx 未安装。请运行 install.ps1 安装依赖。'
}