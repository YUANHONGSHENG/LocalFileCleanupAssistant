$ErrorActionPreference = 'Stop'

try {
    $app = Join-Path $PSScriptRoot 'App.ps1'
    & $app @args
}
catch {
    $log = Join-Path (Split-Path -Parent $PSScriptRoot) 'launcher-error.txt'
    try {
        [IO.File]::WriteAllText($log, ($_ | Out-String), [Text.UTF8Encoding]::new($false))
    }
    catch { }

    try {
        Add-Type -AssemblyName PresentationFramework
        [Windows.MessageBox]::Show(
            "启动失败：`n$($_.Exception.Message)`n`n日志：$log",
            '本地文件清理助手',
            [Windows.MessageBoxButton]::OK,
            [Windows.MessageBoxImage]::Error) | Out-Null
    }
    catch { }
}
