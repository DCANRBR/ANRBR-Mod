Import-Module ".\ANRBR-Mod.PSM1" -Force

Measure-Command {
Get-Process | Export-ExcelANRBR -Path '.\testsheet.xlsx' -TimeStamp
}
