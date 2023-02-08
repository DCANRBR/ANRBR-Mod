Import-Module ".\ANRBR-Mod.PSD1" -Force


Copy-ModuleToUserMods -ModulePSMPath '.\ANRBR-Mod.PSM1'
<#
Measure-Command {
Get-Process | Export-ExcelANRBR -Path '.\testsheet.xlsx' -TimeStamp
}


$myPath = "C:\Users\anrbr\Downloads"

$items = Get-ChildItem -Path $myPath -Recurse -Force | Measure-Object -Property
#>