$targetPath = "C:\Windows\System32\calc.exe"
$linkPath = "$env:APPDATA\Help.lnk"

$wordRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe"
$wordPath = (Get-ItemProperty -Path $wordRegistryPath).'(default)'
echo $wordPath

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($linkPath)
$shortcut.IconLocation = "$wordpath, 0"
$shortcut.TargetPath = $targetPath
$shortcut.Save()
