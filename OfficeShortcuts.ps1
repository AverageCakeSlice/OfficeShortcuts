#OfficeShortcuts -- Creates shortcuts for the main four Office 2016 applications.

#Optional command-line parameters that can be passed to create icons for all (-a) icons, or a custom set of icons (-c)
param([switch] $a, [switch] $c, [switch] $h)

#Creates a shortcut on the desktop that copies from the start menu shortcuts.
function CreateOfficeDesktopShortcut([string] $ShortcutName)
{
    Copy-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$ShortcutName.lnk" "$env:USERPROFILE\Desktop"
}

#Prints menu for selecting individual programs
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
     Write-Host "8: Skype for Business"
     Write-Host "D: Press 'D' when done."
}

#Requests custom user input
function GetCustomUserSelection()
{
    [string[]] $CustomSelectionGroup = @()
    do
    {
        Show-Menu
        Write-Host "You have selected: $CustomSelectionGroup"
        $ProgramSelection = Read-Host "Please make a selection"
        switch($ProgramSelection)
        {
            '1'{
                $CustomSelectionGroup += "Word*"
            }
            '2'{
                $CustomSelectionGroup += "Excel*"
            }
            '3'{
                $CustomSelectionGroup += "PowerPoint*"
            }
            '4'{
                $CustomSelectionGroup += "Outlook*"
            }
            '5'{
                $CustomSelectionGroup += "Access*"
            }
            '6'{
                $CustomSelectionGroup += "OneNote*"
            }
            '7'{
                $CustomSelectionGroup += "Publisher*"
            }
            '8'{
                $CustomSelectionGroup += "Skype*"
            }
            'd'{
            }
        }
    }
    until ($ProgramSelection -eq 'd')
    $SelectedPrograms = $CustomSelectionGroup
    $SelectedPrograms
}

#Function that changes based on the selected switches, defaults to the four most common office programs if no switch is input.
function SelectOfficeShortcuts([bool]$a, [bool]$c, [bool]$h)
{
    [string[]] $SelectedPrograms = @("Word*", "Excel*", "PowerPoint*", "Outlook*")

    if ($a)
    {
        $SelectedPrograms += "Access*", "OneNote*", "Publisher*", "Skype*"
    }
    elseif($c)
    {
        $SelectedPrograms = GetCustomUserSelection
    }
    elseif($h)
    {
        Write-Host "`n========================= SYNTAX =========================`n"
        Write-Host "Default    Copy Word, Excel, Powerpoint, and Outlook 2016"
        Write-Host "     -a    Copy all shortcuts for Office 2016 products"
        Write-Host "     -c    Copy a custom list of Office shortcuts, specified by the user"
        Write-Host "     -h    Display help for this script"
        exit
    }
    else {
        $SelectedPrograms
    }

    Write-Host "Got it! Copying shortcuts... $SelectedPrograms"
    
    for ($i=0; $i -lt $SelectedPrograms.Length; $i++)
    {
        CreateOfficeDesktopShortcut $SelectedPrograms[$i]
    }
}

#Run 
SelectOfficeShortcuts($a, $c, $h)
