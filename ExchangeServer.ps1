# Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'
Import-Module .\common.ps1

function Load-ExchangeCommands {
    . "C:\Program Files\Microsoft\Exchange Server\V15\RemoteScripts\ConsoleInitialize.ps1"
    Connect-ExchangeServer -auto -ClientApplication:ManagementShell
}


function Connect-Exchange {

    . 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
    Connect-ExchangeServer -auto
    return $Session
}


# More info:https://www.ultimatewindowssecurity.com/exchange/mailboxaudit/configure.aspx
function Enable-AuditLogging
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [Microsoft.Exchange.Data.Directory.Management.Mailbox]$Mailbox
    )

    Process
    {

        try{
            Set-Mailbox -Identity $Mailbox.Name -AuditAdmin FolderBind -AuditDelegate FolderBind -AuditOwner None -AuditEnabled $true
        } catch {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text $_.exception.message -LogfileName ExchangeServer.log
            Pause
        }

        If (Get-Mailbox $Mailbox.Name | Select-Object AuditEnabled) {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text "Logging ENABLED" -LogfileName ExchangeServer.log
        } else {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text "FAILED" -LogfileName ExchangeServer.log
        }
        
    }
}


# Get-MailboxDatabase ALMB01_DB1_ST | ForEach-Object {get-mailbox -Database $_.Name | Disable-AuditLogging}
function Disable-AuditLogging
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        [Microsoft.Exchange.Data.Directory.Management.Mailbox]$Mailbox
    )

    Process
    {
        try{
            do{
                Set-Mailbox -Identity $Mailbox.Name -AuditEnabled $false
                Start-Sleep 2
            }
            while ((Get-Mailbox $Mailbox.Name).AuditEnabled)
            
        } catch {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text $_.exception.message -LogfileName ExchangeServer.log
            Pause
        }

        "Auditing enabled : $((Get-Mailbox $Mailbox.Name).AuditEnabled)"
        "False : $False"
        "Equality : $((Get-Mailbox $Mailbox.Name).AuditEnabled -eq $false)"

        If ((Get-Mailbox $Mailbox.Name).AuditEnabled -eq $false) {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text "Logging DISABLED" -LogfileName ExchangeServer.log
        } else {
            Write-ChangeLog -ObjectIdentifier $Mailbox.Name -Text "FAILED" -LogfileName ExchangeServer.log
        }
    }
}


function Delete-ExchangeLogs
{
    # RUN AS ADMIN!

    $CDrive = Get-Volume C
    $FreeRatio = $CDrive.SizeRemaining/$CDrive.Size
    $OldFreePercent = "{0:p}" -f $FreeRatio

    # Remove log files
    [array]$TargetFiles = $null
    $TargetFiles += Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\DailyPerformanceLogs"
    $TargetFiles += Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\PerformanceLogsToBeProcessed"
    $TargetFiles += Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Bin\Search\Ceres\Diagnostics\ETLTraces"
    Write-Host $TargetFiles.count "files found"
    $TargetFiles | Remove-Item -Confirm

    $CDrive = Get-Volume C
    $FreeRatio = $CDrive.SizeRemaining/$CDrive.Size
    $FreePercent = "{0:p}" -f $FreeRatio

    "Original Percent free space: $OldFreePercent"
    "Current Percent free space:  $FreePercent"
}



function Get-DynamicDistributionGroupMembers {

    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $GroupName
    )

    $Group = Get-DynamicDistributionGroup $GroupName
    Get-Recipient -ResultSize Unlimited -RecipientPreviewFilter $Group.LdapRecipientFilter -OrganizationalUnit $Group.RecipientContainer
}


function Get-DynamicDistributionGroupFilter {

    Param
    ($GroupName)

    (Get-DynamicDistributionGroup $GroupName).LdapRecipientFilter
}


function Enable-MailboxArchive {
    param (
    )
    # Run on prem for summary
    # Get-RemoteMailbox -ResultSize unlimited -Filter {ArchiveState -Eq "None" -AND RecipientTypeDetails -eq "RemoteUserMailbox"} | Measure-Object

    # Run on prem to enable mailboxes:
    Get-RemoteMailbox -ResultSize 50 -Filter {ArchiveState -Eq "None" -AND RecipientTypeDetails -eq "RemoteUserMailbox" -and ExchangeUserAccountControl -ne 'AccountDisabled'}  | Enable-RemoteMailbox -Archive | Select-Object name, recipienttypedetails, archivestate, archivestatus

}




function Add-MailRelayIPAddress {

    [CmdletBinding()]
    param (
        $IPAddress
    )

    [array]$RelayConnectors = "USTXALMHUB1\Relay Connector","USTXALMHUB2\Relay Connector","USTXALMHUB3\Relay Connector","ustxalmmb01\SMTP Relay","ustxalmmb02\SMTP Relay","ustxalmmb03\SMTP Relay","ustxalmmb04\SMTP Relay","ustxalmmb05\SMTP Relay","ustxalmmb06\SMTP Relay","ustxirmmb01\SMTP Relay","ustxirmmb02\SMTP Relay","ustxirmmb03\SMTP Relay","ustxirmmb04\SMTP Relay","ustxirmmb05\SMTP Relay","ustxirmmb06\SMTP Relay"
    
    foreach ($Connector in $RelayConnectors) {
        $RecvConn = Get-ReceiveConnector $Connector
        $RecvConn.RemoteIPRanges += $IPAddress
        Set-ReceiveConnector $Connector -RemoteIPRanges $RecvConn.RemoteIPRanges
        $RecvConn.RemoteIPRanges
    }
}

function Remove-EmailFromPublicFolder {
    param (
        $Folder,
        $Email
    )
    Get-MailPublicFolder -Identity $Folder | Select-Object -ExpandProperty emailaddresses
    Set-MailPublicFolder -Identity $Folder -emailaddresses @{Remove=$Email} -EmailAddressPolicyEnabled $false
}



function Add-EmailToPublicFolder {
    param (
        $Folder,
        $Email
    )
    Get-MailPublicFolder -Identity $Folder | Select-Object -ExpandProperty emailaddresses
    Set-MailPublicFolder -Identity $Folder -emailaddresses $Email -EmailAddressPolicyEnabled $false
}


function Update-MailboxQuotas {
    param (
        $EmailAddress,
        $WarningQuota = 4gb,
        $ProhibitSendQuota = 5gb,
        $Confirm = $false
    )

    Set-Mailbox $EmailAddress -UseDatabaseQuotaDefaults $false -IssueWarningQuota $WarningQuota -ProhibitSendQuota $ProhibitSendQuota -ProhibitSendReceiveQuota Unlimited -Confirm:$Confirm
}


function Add-SafeSenders {
    [CmdletBinding()]
    param (
        $Mailbox,
        [array]$SafeSenderAddresses
    )
    
    process {

        # Get the current safe senders and add the new one(s) to that list
        $CurrentSafeSenders = (Get-Mailbox $Mailbox | Get-MailboxJunkEmailConfiguration).TrustedSendersAndDomains
        $NewSafeSenders = $CurrentSafeSenders + $SafeSenderAddresses
        $NewSafeSenders = $NewSafeSenders | Select-Object -Unique

        # Check the addresses we're adding and discard if they already exist
        $SafeSenderAddresses | ForEach-Object {
            if ($CurrentSafeSenders -contains $_) {
                return "$Mailbox : Safe sender $_ already in the list"
            } else {
                # Add the safe sender to the new list
                "$Mailbox : Adding safe sender $_"
                Get-Mailbox $Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $NewSafeSenders
            }
        }
    }
}

function Enable-ArchiveForMigratedUsers {
    param (
        $TargetOU
    )
    
    if ($TargetOU) {
        Get-RemoteMailbox -Filter {archiveguid -eq $Null} | 
        where remoterecipienttype -eq "Migrated" | 
        Enable-RemoteMailbox -Archive -Confirm   
    } else {
        Get-RemoteMailbox -Filter {archiveguid -eq $Null} | 
        where remoterecipienttype -eq "Migrated" | 
        Enable-RemoteMailbox -Archive -Confirm 
    }
}


function Get-MailboxUsage {
    param (
        $Email
    )
    $User = Get-ADUser -Filter {mail -eq $Email} -Properties msexchrecipienttypedetails
    if ($User.msexchrecipienttypedetails -eq 1){
        Get-MailboxFolderStatistics $Email |
        Select-Object FolderPath, @{
            N = "TotalSize";
            E = {
                "{0:N2}" -f ((($_.FolderSize -replace "[0-9\.]+ [A-Z]* \(([0-9,]+) bytes\)","`$1") -replace ",","") / 1MB)
            }
        } | Measure-Object -Property TotalSize -sum | select Sum
    } else {
        "$Email - Recipient type: $($User.msexchrecipienttypedetails)"
    }
}


# Enable the on-prem mailbox and set the GUID to match O365
# You must connect to O365 first and get the GUID
# You can run it for an individual or with a CSV input - column names should be "UserID" and "GUID"
# Usage: Enable-ExistingRemoteMailbox "tharw001" "200eb482-3d08-45b5-a95c-00127dcf168b"
# Usage: Enable-ExistingRemoteMailbox -CSVFilePath C:\Temp\GUIDs.csv
function Enable-ExistingRemoteMailbox {
    param (
        [string]$CSVFilePath,
        [Parameter(ValueFromPipeline=$true,
                   Position=0)]
        [string]$UserID,
        [Parameter(ValueFromPipeline=$true,
                   Position=1)]
        [string]$GUID
    )

    if ($CSVFilePath) {
        $CSVFile = Import-Csv $FilePath
        foreach ($Row in $CSVFile) {
            Enable-ExistingRemoteMailbox $CSVFile.UserID $CSVFile.GUID
        }
    } else {
        $User = Get-ADUser $UserID
        Enable-RemoteMailbox -Identity $User.sAMAccountName -RemoteRoutingAddress "$($User.sAMAccountName)@711com.mail.onmicrosoft.com" | Out-Null
        Set-RemoteMailbox -Identity $User.sAMAccountName -ExchangeGuid $GUID
        sleep 2
        Get-RemoteMailbox $User.sAMAccountName | select DisplayName, ExchangeGuid
    }
}


# Usage: New-ConferenceRoom -RoomNumber R4D-33C -Location RD -Description "Presentation Only"
function New-ConferenceRoom {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        $RoomNumber,
        [ValidateSet("SSC","1601","1603","RD")]
        $Location,
        $Description,
        $BookinPolicy
    )
    
    process {

        if ($Description){
            $DisplayName = "Conf $Location $RoomNumber $Description"
        } else {
            $DisplayName = "Conf $Location $RoomNumber"
        }
        
        $UserPrincipalName = "$Location-$RoomNumber@7-11.com"
        $SamAccountName = "$Location-$RoomNumber"
        $NewRoom = New-RemoteMailbox $DisplayName -UserPrincipalName $UserPrincipalName -SamAccountName $SamAccountName -PrimarySmtpAddress $UserPrincipalName -OnPremisesOrganizationalUnit "7-11.com/corp/user groups/conference rooms" -room
        Write-Warning "Be sure to log on to O365 after syncing and set the bookin policy!"
        # Set-CalendarProcessing -Identity $NewRoom.mail -bookingwindowindays 425 -maximumdurationinminutes 720 -maximumconflictinstances 25 -conflictpercentageallowed 25 -allowconflicts $false -allowrecurringmeetings $true -scheduleonlyduringworkhours $false -enforceschedulinghorizon $true -deletesubject $false -forwardrequeststodelegates $false 

        return $NewRoom
    }
}


function Search-ExchangeLogs {
    param (
        $Sender,
        $Recipients
    )
    Get-MailboxServer USTXALMMB02 | Get-MessageTrackingLog -Sender $Sender -Recipients $Recipients -EventId Receive
}

Connect-Exchange