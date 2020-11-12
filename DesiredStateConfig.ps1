
<#
    Desired State Configuration
    -----------------------------------
    To make sure all desired settings for AD, Exchange and other objects in
    the environment are meeting requirements - store mailboxes have the correct
    licenses, store DLs are not hidden from the address book, FC groups
    are not restricted from receiving outside email, etc.
#>

function Log-DSCChanges {
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline, Position=0)]
        $ObjectID,
        [Parameter(Mandatory=$true, ValueFromPipeline, Position=1)]
        $StatusMessage,
        $LogsPath="c:\temp",
        $LogFileName="DesiredStateConfig_$(Get-Date -f "yyyy-MM-dd").log"
    )
    
    Begin{
        if (! (Test-Path $LogsPath\$LogFileName)) {
            New-Item $LogsPath\$LogFileName -ItemType File -Force -Confirm
        }
    }
    
    Process{
        Write-Host "$(Get-Date -f "yyyy-MM-dd HH:mm:ss") : $ObjectID : $StatusMessage" -ForegroundColor Yellow
        "$(Get-Date -f "yyyy-MM-dd HH:mm:ss") : $ObjectID : $StatusMessage" | Out-File -FilePath $LogsPath\$LogFileName -Append
    }

    End{}
}


#        ------------------ RULES ---------------------

# Skype
#   Unified Messaging requirements
#       All AD accounts
#           If HostedVoicemailPolicy = CloudUM, then HostedVoiceMail = True

Get-ADUser -Filter "HostedVoicemailPolicy -eq 'CloudUM'" -ResultSetSize unl