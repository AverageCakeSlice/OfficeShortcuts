#OfficeShortcuts -- Creates shortcuts for the main four Office 2016 applications.

#Optional command-line parameters that can be passed to create icons for All (-a) icons, or a CUSTOM set of icons (-c)
param([switch] $a, [switch] $c)

#Creates a shortcut on the desktop that copies from the start menu shortcuts.
function CreateOfficeDesktopShortcut([string] $ShortcutName)
{
    Copy-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$ShortcutName.lnk" "$env:USERPROFILE\Desktop"
}

function Show-Menu
{
     Clear-Host
     Write-Host "Select which items you would like placed on the desktop:"
     
     Write-Host "1: Word"
     Write-Host "2: Excel"
     Write-Host "3: PowerPoint"
     Write-Host "4: Outlook"
     Write-Host "5: Access"
     Write-Host "6: OneNote"
     Write-Host "7: Publisher"
     Write-Host "D: Press 'D' when done."
}

function GetCustomUserSelection()
{
    [string[]]$CustomSelection = @()
    do
    {
        Clear-Host
        Show-Menu
        Write-Host "You have selected: $CustomSelection"
        $input = Read-Host "Please make a selection:"
        switch($input)
        {
            '1'{
                $CustomSelection += "Word"
            }
            '2'{
                $CustomSelection += "Excel"
            }
            '3'{
                $CustomSelection += "PowerPoint"
            }
            '4'{
                $CustomSelection += "Outlook"
            }
            '5'{
                $CustomSelection += "Access"
            }
            '6'{
                $CustomSelection += "OneNote"
            }
            '7'{
                CustomSelection += "Publisher"
            }
            'd'{
                return
            }
        }
    }
    until ($input -eq 'd')
    "Now do something else"
}

#Loop through the array of Shortcut names, and create a shortcut with a matching alias
function SelectOfficeShortcuts($arg1, $arg2)
{
    [string[]] $OfficeProgramNames = @("Word", "Excel", "PowerPoint", "Outlook")

    if ($a -eq $True)
    {
        $OfficeProgramNames += "Access", "OneNote*", "Publisher"
    }
    elseif($c -eq $True)
    {
        $OfficeProgramNames = @()
        GetCustomUserSelection $OfficeProgramNames
    }
    for ($i=0; $i -lt $OfficeProgramNames.Length; $i++)
    {
        CreateOfficeDesktopShortcut $OfficeProgramNames[$i]
    }
}

SelectOfficeShortcuts($a, $c)
