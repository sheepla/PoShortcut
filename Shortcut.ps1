
<#
.DESCRIPTION
Get shortcut object from *.lnk file
.USAGE
Get-Shortcut foo.lnk
Get-Shortcut foo.lnk -GetTarget
#>
function Get-Shortcut
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$false, Position=0)]
        [String] $Path,

        [Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)]
        [Switch] $GetTarget
    )

    # Create link object
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut($Path)

    if ($GetTarget)
    {
        Write-Output $lnk.TargetPath
    }
    else
    {
        Write-Output $lnk
    }
}

<#
.DESCRIPTION
Create desktop shortcut (*.lnk) file

.
New-Shortcut -Path ~/Desktop/notepad.lnk -TargetPath C:/Windows/System32/notepad.exe -Description "Notepad"
#>
function New-Shortcut
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$false, Position=0, HelpMessage="Path of *.lnk or *.url file")]
        [String] $Path,

        [Parameter(Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Position=1)]
        [String] $TargetPath,

        [Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)]
        [String] $Description
    )

    # Parse $Path argument
    [String] $dirName = [IO.Path]::GetDirectoryName($Path)
    [String] $fileName = [IO.Path]::GetFileNameWithoutExtension($Path)
    [String] $ext = [IO.Path]::GetExtension($Path)
    [Bool] $pathIsValid = ($ext -eq ".lnk") -and (-not (Test-Path $filename)) -and (Test-Path $dirName)

    # Create link object
    $sh = New-Object -ComObject WScript.Shell
    $lnk = $sh.CreateShortcut($Path)
    $lnk.TargetPath = $TargetPath

    if ($null -ne $Description)
    {
        $lnk.Description = $Description
    }

    $lnk.Save()
    Write-Verbose "Created link: $lnk"
}
