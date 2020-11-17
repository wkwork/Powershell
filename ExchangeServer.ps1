# Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'
. .\common.ps1

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

# Usage: Search-ExchangeLogs -Sender corpcomm@7-11.com
# Usage: Search-ExchangeLogs -Sender corpcomm@7-11.com | group EventId
function Search-ExchangeLogs {
    param (
        $Sender,
        $Recipients,
        $DaysToReturn = 30
    )
    $StartDate = (Get-Date).AddDays(-$DaysToReturn)
    $Logs = Get-ExchangeServer | where { $_.serverrole -eq 'Mailbox' } | Get-MessageTrackingLog -Sender $Sender -Recipients $Recipients -ResultSize Unlimited -Start $StartDate | where {($_.EventId -eq "Receive") -or ($_.EventId -eq "Deliver") -or ($_.EventId -eq "Fail")} | sort TimeStamp -Descending
    "Found $($Logs.count) log entries"
    $Logs
}




function Update-MarketManagerGroups {
    [CmdletBinding()]
    param ()
    
    begin {$Corrected = 0}
    
    process {

        Write-Host "Checking MM groups"

        Get-DistributionGroup -Filter "Name -like 'Mkt Mgr MKT-*'" -ResultSize 20000 | where RequireSenderAuthenticationEnabled -eq $True | foreach-object {

            Write-Host " "
            Write-Host "UPDATING $($_.Name) (created $($_.WhenCreated))..."
            Set-DistributionGroup -Identity "$($_.Name)" -RequireSenderAuthenticationEnabled $False
            $Corrected++
        }
    }

    end {
        Write-Warning "$Corrected corrected."
    }
}



function Update-FieldConsultantGroups {
    [CmdletBinding()]
    param ()
    
    begin {$Corrected = 0}
    
    process {

        Write-Host "Checking FC groups"

        Get-DistributionGroup -Filter "Name -like 'field consultant *'" -ResultSize 20000 | where Name -match "Field Consultant \d*" | where RequireSenderAuthenticationEnabled -eq $True | foreach-object {

            Write-Host " "
            Write-Host "UPDATING $($_.Name) (created $($_.WhenCreated))..."
            Set-DistributionGroup -Identity "$($_.Name)" -RequireSenderAuthenticationEnabled $False
            $Corrected++
        }
    }

    end {
        Write-Warning "$Corrected corrected."
    }
}


function Update-StoreManagerGroups {
    [CmdletBinding()]
    param ()
    
    begin {$Corrected = 0}
    
    process {

        Write-Host "Checking store groups"

        Get-DistributionGroup -Filter "Name -like 'Store Manager *'" -ResultSize 20000 | where Name -match "Store Manager \d*" | where RequireSenderAuthenticationEnabled -eq $True | foreach-object {
            Write-Host " "
            Write-Host "UPDATING $($_.Name) (created $($_.WhenCreated))..."
            Set-DistributionGroup -Identity "$($_.Name)" -RequireSenderAuthenticationEnabled $False
            $Corrected++
        }
    }
    
    end {
        Write-Warning "$Corrected corrected."
    }
}


function Update-DistributionGroups {
    param ()
    
    process{
        Update-StoreManagerGroups
        Update-FieldConsultantGroups
        Update-MarketManagerGroups
    }
}


function Search-ExchangeObjects {
    param (
        $SearchTerm
    )
    "Searching for $SearchTerm"
    Get-Recipient -Filter "name -like '*$SearchTerm*' -or PrimarySmtpAddress -like '*$SearchTerm*'" | select Name, PrimarySmtpAddress, RecipientType, RecipientTypeDetails | Format-Table -AutoSize
}

<#
.Synopsis
   Creates new account and mailbox for Horizon stores
.DESCRIPTION
   Creates a new remote mailbox, sets the description
   and name and adds the AD account to the Horizon groups.
.Example
    New-HorizonStore 38592
.Example
    38592, 31241, 45658 | New-HorizonStore
#>
function New-HorizonStore {
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $StoreNumber
    )
    
    process{
        # Create the remote mailbox
        $DisplayName = "COOP " + $StoreNumber
        $Alias = "COOP" + $StoreNumber
        $Password = ConvertTo-SecureString "&vZYS5xQ&^6*" -AsPlainText -Force
        $OU = "7-11.com/Stores/Sunoco"

        $User = $null
        if ($User = Get-ADUser $Alias){
            "{0} already exists" -f $User.Name
        } else {
            "Creating $Alias@7-11.com"
            New-RemoteMailbox -Name $DisplayName -UserPrincipalName "$Alias@7-11.com" -SamAccountName $Alias -RemoteRoutingAddress "$Alias@711com.mail.onmicrosoft.com" -Password $Password -DisplayName $DisplayName -OnPremisesOrganizationalUnit $OU
            $User = $null
            do {
                "Confirming user..."
                $User = Get-ADUser $Alias
                sleep 1
            } while (! $User)
            "Confirmed" + $User.DisplayName
        }

        # Set description and last name
        "Setting properties..."
        Set-ADUser $Alias -Surname $Alias -Description "Stripes - West" -Credential $AdminCred

        # Assign to AD groups
        "NAC-SunocoInStore",
        "su_SunocoStores-1-20016944",
        "su-Stores-Texas_Locations-1311725592",
        "USER-MS-HORIZON-IPADS",
        "USER-MS-Sub-AADPP1-F1",
        "USER-MS-Sub-O365-F1-COOP_West" |ForEach-Object {
            "Adding to $_..."
            Add-ADGroupMember -Identity $_ -Members $Alias -Credential $AdminCred
        }
    }
}




<#
.Synopsis
   Creates new remote mailbox
.DESCRIPTION
   Creates new remote mailbox. Use -shared for a shared mailbox.
#>
function Create-RemoteMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $DisplayName,
        [Parameter(Mandatory=$true)]
        $Alias,
        # Like "7-11.com/Corp/User Service Accounts"
        $OU = "7-11.com/Corp/User Service Accounts",
        [switch]$Shared
    )
    begin{
        $Password = ConvertTo-SecureString "Inc0gn1t0!" -AsPlainText -Force
    }

    process {
        "Creating $Alias@7-11.com"
        if ($Shared) {
            New-RemoteMailbox -Name $DisplayName -Shared -UserPrincipalName "$Alias@7-11.com" -SamAccountName $Alias -RemoteRoutingAddress "$Alias@711com.mail.onmicrosoft.com" -Password $Password -DisplayName $DisplayName -OnPremisesOrganizationalUnit $OU
        } else {
            New-RemoteMailbox -Name $DisplayName -UserPrincipalName "$Alias@7-11.com" -SamAccountName $Alias -RemoteRoutingAddress "$Alias@711com.mail.onmicrosoft.com" -Password $Password -DisplayName $DisplayName -OnPremisesOrganizationalUnit $OU
        }
    }
}


Connect-Exchange




########################### working ###############################


<#

Contract to Perm
A novel in 12 parts

Confirm old and new accounts
If old account is on-prem
    Migrate to O365
Soft-Delete mailbox
    Disable remote mailbox on-prem
    Remove O365 license
    Sync AAD
Restore inactive old mailbox to new mailbox
Remove inactive account/mailbox completely
Assign old address as alias to new mailbox
#>

function Convert-ContractMailboxToFteMailbox {
    param (
        [string]$OldAccountName = "skees002",
        [string]$NewAccountName = "s1922843"
    )
    
    # Confirm users exist
    If ($OldUser = Get-ADUser $OldAccountName -Properties Mail, MSExchRecipientTypeDetails) {"Found $Oldaccountname"} else {Write-Warning "$Oldaccountname not found in AD!"}
    If ($NewUser = Get-ADUser $NewAccountName -Properties Mail, MSExchRecipientTypeDetails) {"Found $NewAccountName"} else {Write-Warning "$NewAccountName not found in AD!"}
    
    If ($OldUser -and $NewUser){
    
        # Connect to Exchange
        $OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionURI http://USTXALMMB01.7-11.com/PowerShell/ -Authentication Kerberos -WarningAction Silentlycontinue
        Import-PSSession $OnPremSession -WarningAction SilentlyContinue -Prefix OnPrem
    
        # Get mailboxes for both accounts
        # TODO: Need to confirm where each mailbox is. If old mailbox is on-prem, connect to O365 and migrate it.
        $OldUser, $NewUser | ForEach-Object {
            if ($_.Mail){
                If ($_.MSExchRecipientTypeDetails -eq 1){Get-OnPremMailbox $_.Mail -}
                If ($_.MSExchRecipientTypeDetails -eq 2147483648){Get-OnPremRemoteMailbox $_.Mail}
            } else {
                Write-Warning "$($_.Name) does not have an email address"
                Exit
            }
        }
    
        Remove-PSSession $OnPremSession
    
    
    } else {
        exit
    }

}