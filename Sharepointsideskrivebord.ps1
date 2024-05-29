$SharePointURL = "https://ronjafagprove.sharepoint.com/"


$ShortcutPath = "$env:Public\Desktop\SharePoint Snarvei.lnk"

$WScriptShell = New-Object -ComObject WScript.Shell

$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $SharePointURL
$Shortcut.Save()