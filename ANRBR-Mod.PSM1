<#
    Welcome to my library

    V. 1.1
#>

function Export-ExcelANRBR {
    param (
        [Parameter(Mandatory, ValueFromPipeline)]$inputObject,
        [Parameter(Mandatory)][string]$Path,
        [switch]$TimeStamp,
        [string]$TableStyle
    )
    begin {
        if ($TimeStamp) {
            $dirName = [io.path]::GetDirectoryName($path)
            $filename = [io.path]::GetFileNameWithoutExtension($path)
            $ext = [io.path]::GetExtension($path)
            $filenameStamped = $filename + "_" + (Get-TimeFileANRBR)
            $Path = "$dirName\$filenameStamped$ext"
        }
        if (-not $TableStyle) {
            $TableStyle = 'Medium2'
        }
    
        $flags = @{
            Path       = $Path
            TableStyle = $TableStyle
            AutoSize   = $true
            ClearSheet = $true
        }
        $piped = [collections.Arraylist]@()
    }
    process {
        [void]$piped.Add($inputObject)
    }
    end {
        $piped | Export-Excel @flags
    }    
}

function Get-TimeFileANRBR {
    param (
        #OptionalParameters
    )
    Get-Date -f yyyyMMdd-HHmmss
}

function Copy-ModuleToUserMods {
    param (
        [Parameter(Mandatory)][string]$ModulePSMPath
    )
    $userModulesDir = ($env:PSModulePath -split ";")[0]
    $moduleName = [io.path]::GetFileNameWithoutExtension($ModulePSMPath)
    $moduleDir = "$userModulesDir\$moduleName"
    $modulePSMDir = [io.path]::GetDirectoryName($ModulePSMPath)
    $modulePSD = "$modulePSMDir\$moduleName.psd1"

    if (-Not (Test-Path $moduleDir)) {
        New-Item -ItemType Directory -Path $moduleDir
    }
    Copy-Item -Path $ModulePSMPath -Destination $moduleDir -Force
    Copy-Item -Path $modulePSD -Destination $moduleDir -Force
    if (Get-Module -Name $moduleName) {
        Remove-Module -Name $moduleName
    }
}
#Copy-ModuleToUserMods -ModulePSMPath '.\ANRBR-Mod.PSM1'

#Export-ModuleMember -Function Export-ExcelANRBR
#Export-ModuleMember -Function Get-TimeFileANRBR
#Export-ModuleMember -Function Copy-ModuleToUserMods