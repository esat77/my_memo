# create_shortcut.ps1
# hello.exe と同じフォルダ内に run_hello.bat がある前提

$baseDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath   = Join-Path $baseDir "run_hello.bat"
$exePath   = Join-Path $baseDir "hello.exe"
$lnkPath   = Join-Path $baseDir "run_hello.lnk"

$WshShell  = New-Object -ComObject WScript.Shell
$shortcut  = $WshShell.CreateShortcut($lnkPath)

$shortcut.TargetPath   = $batPath
$shortcut.WorkingDirectory = $baseDir
$shortcut.WindowStyle  = 7   # 最小化
$shortcut.IconLocation = $exePath  # hello.exe のアイコンを使用

$shortcut.Save()

Write-Output "Shortcut created: $lnkPath"
