#OfficeShortcuts -- Creates shortcuts for the main four Office 2016 applications.

[string[]] $Office16ExeNames = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "OUTLOOK.EXE")
[string[]] $Office16ExeAliases = @("Word", "Excel", "PowerPoint", "Outlook")

#Creates a WScript.Shell object, sets the target path and destination, and saves the shortcut with a more readable name.
#Accepts the .exe file name and readable shortcut name as arguments.
function CreateOfficeShortcut([string]$ExeName, [string]$ExeAlias)
{
    $TargetFile = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\$ExeName"
    $ShortcutFile = "$env:USERPROFILE\Desktop\$ExeAlias 2016.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.Save()
}

#Loop through the array of executable names, and create a shortcut with a matching alias
for ($i=0; $i -lt $Office16ExeNames.Length; $i++)
{
    CreateOfficeShortcut $Office16ExeNames[$i] $Office16ExeAliases[$i]
}