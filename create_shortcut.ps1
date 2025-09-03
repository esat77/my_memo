param(
    [string]$ExePath = ".\hello.exe",      # exe のパス（デフォルト: カレントにある myapp.exe）
    [string]$ShortcutPath = "$env:USERPROFILE\Desktop\hello.lnk"  # ショートカット保存先（デフォルト: デスクトップ）
)

# WScript.Shell COM オブジェクトを作成
$WshShell = New-Object -ComObject WScript.Shell

# ショートカットを作成
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)

# exe のフルパスを設定
$Shortcut.TargetPath = (Resolve-Path $ExePath).Path

# 作業フォルダを exe のフォルダに
$Shortcut.WorkingDirectory = Split-Path $ExePath -Parent

# 最小化で起動するように設定（7 = 最小化）
$Shortcut.WindowStyle = 7

# アイコンを exe と同じにする
$Shortcut.IconLocation = $Shortcut.TargetPath

# 保存
$Shortcut.Save()
