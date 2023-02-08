<#
    Welcome to my library

    V. 1.0
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

Export-ModuleMember -Function Export-ExcelANRBR
Export-ModuleMember -Function Get-TimeFileANRBR