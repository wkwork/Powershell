function Write-ChangeLog
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $Text,
        $ObjectIdentifier,
        $LogfileName = "Temp.log"
    )

    Process{
        $DateTime = Get-Date
        "{0} : {1} : {2}" -f $DateTime, $ObjectIdentifier, $Text | Tee-Object -Append "C:\scripts-utilities\Logs\$LogfileName"
    }
}

function prompt {
    "PS ...\" + ((Get-Location).path.ToString() -split "\\")[-2] + "\" + ((Get-Location).path.ToString() -split "\\")[-1] + "> "
}