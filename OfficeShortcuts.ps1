#OfficeShortcuts -- Creates shortcuts for the main four Office 2016 applications.
$ExeName = "WINWORD.exe"
$ExeAlias = "Word"


#Creates a WScript.Shell object, sets the target path and destination, and saves the shortcut with a more readable name.
#Accepts the .exe file name and readable shortcut name as arguments.
function CreateOfficeShortcut([string]$MyName, [string]$MyAlias)
{
    $TargetFile = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\$MyName"
    $ShortcutFile = "$env:USERPROFILE\Desktop\$MyAlias 2016.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.Save()
}

CreateOfficeShortcut $ExeName $ExeAlias