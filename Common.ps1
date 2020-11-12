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

<#
.DESCRIPTION

Takes a simple store number like 10766 and
returns all users in the store groups for that store

.EXAMPLE

Get-StoreUsers 10766
#>
function Get-StoreUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        $StoreNumber
    )

    begin{
        $Result = $null
    }
    
    process {
        [array]$Result += Get-ADGroupMember -Identity "store manager $StoreNumber" | Get-ADUser -Properties Title, Mail
        [array]$Result += Get-ADGroupMember -Identity "designee $StoreNumber" | Get-ADUser -Properties Title, Mail
    }

    end{
        $Result
    }
}

