$targetPath = "C:\Windows\System32\calc.exe"
$linkPath = "$env:APPDATA\Help.lnk"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($linkPath)
$shortcut.IconLocation = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE, 0"
$shortcut.TargetPath = $targetPath
$shortcut.Save()
