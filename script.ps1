$OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("UTF-8")

$targetPath = "C:\Windows\System32\calc.exe"
$linkPath = "$env:APPDATA\$([char]0x0421)$([char]0x043F)$([char]0x0440)$([char]0x0430)$([char]0x0432)$([char]0x043A)$([char]0x0430).lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($linkPath)
$shortcut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE, 0"
$shortcut.TargetPath = $targetPath
$shortcut.Save()
