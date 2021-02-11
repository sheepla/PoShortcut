
function New-SpecialShortcut
{
    param
    (
        [Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false,
        HelpMessage="OutDirectory")]
        [String] $OutDirectory = (Join-Path -Path $env:USERPROFILE -ChildPath "Desktop"),

        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, Position=1)]
        [String[]] $Name,

        [Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)]
        [String] $DisplayName,

        [Parameter(Mandatory=$false, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false)]
        [Switch] $WhatIf
    )

    process
    {
        foreach ($n in $Name)
        {
            $shortcut = Get-SpecialShortcut -Name $n
        }

        foreach ($s in $shortcut)
        {
            $shortcutName = "{0}.{1}" -f $s.Name, $s.CLSID
            $shortcutPath = Join-Path -Path $OutDirectory -ChildPath $shortcutName
            Write-Output $shortcutPath
            New-Item -Type Directory -Path $shortcutPath | Out-Null
        }
    }
}

function Get-SpecialShortcut
{
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Position=0)]
        [AllowEmptyString()]
        [String] $Name
    )

    if ($Name -eq "")
    {
        Import-Csv ./clsid.csv
    }
    else 
    {
        Import-Csv ./clsid.csv | Where-Object {$_.Name -ILike $Name}
    }
}
