# Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'

Import-Module .\ActiveDirectory.ps1

if ($Office365credentials) {
    Write-Warning "Using saved credentials..."
} else {
    [System.Management.Automation.PSCredential]$Office365credentials = Get-Credential -Message "Office 365 Credential"
}

<#
.Synopsis
   Connects to O365 tenant
.DESCRIPTION
   Connects to the 7-11 O365 tenant for further work 
#>
function Connect-O365 {
    param ()

    if (Get-Module -Name MSOnline) {
        Write-Warning "MSOnline module already loaded"
    } else {
        import-module MSOnline
    } 


    # MSOL Service

    if (Get-MsolDomain -ErrorAction SilentlyContinue) {

        Write-Warning "Already connected to Office 365"

    } else {

        if ($Office365credentials) {
            Write-Warning "Using saved credentials..."
        } else {
            $Office365credentials = Get-Credential $Credential -Message "Office 365 credential"
        }

        try{
            Connect-MsolService -Credential $Office365credentials
            Connect-SPOService -Url https://711com-admin.sharepoint.com -Credential $Office365credentials
            Get-MsolCompanyInformation
            Write-Host "Connected to Office 365"
        } catch {
            Write-Warning "Unable to connect to Office 365 : $($_.exception.message)"
        }
    }


    # O365 Exchange (Exchange Online PowerShell Module required!)
    # https://social.technet.microsoft.com/Forums/en-US/6673b735-3b60-49b2-948c-930dac9c3548/how-to-import-mfa-enabled-exchange-online-powershell-module-in-ise?forum=onlineservicesexchange

    if (Get-PSSession -Name winrm*) {

        Write-Warning "Already connected to Office 365 Exchange"

    } else {

        if ($Office365credentials) {
            Write-Warning "Using saved credentials..."
        } else {
            $Office365credentials = Get-Credential $Credential -Message "Office 365 credential"
        }
        
        try{
            Import-Module $((Get-ChildItem -Path "C:\Users\kwork002\AppData\Local\Apps\2.0\" -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName| Where-Object{$_ -notmatch "_none_"} | Select-Object -First 1)
            $EXOSession = New-ExoPSSession -Credential $Office365credentials
            Import-PSSession $EXOSession
            Write-Host "Connected to Office 365 Exchange"
        } catch {
            Write-Warning "Unable to connect to Office 365 Exchange: $($_.exception.message)"
        }
    }
}



<#
.Synopsis
   Resolves the O365 product name
.DESCRIPTION
   Helper function to resolve the O365 product display name
#>
function Get-ProductDisplayName {
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline)]
        [string]$SKU
    )

    Process
    {
        switch ($SKU) {
            '711com:AAD_PREMIUM' {'Azure Active Directory Premium P1'}
            '711com:DESKLESSPACK' {'Office 365 F1'}
            '711com:DYN365_ENTERPRISE_PLAN1' {'Dynamics 365 Customer Engagement Plan Enterprise Edition'}
            '711com:EMS' {'Enterprise Mobility + Security E3'}
            '711com:EMSPREMIUM' {'Enterprise Mobility + Security E5'}
            '711com:ENTERPRISEPACK' {'Office 365 Enterprise E3'}
            '711com:ENTERPRISEPREMIUM' {'Office 365 Enterprise E5'}
            '711com:FLOW_FREE' {'Microsoft Flow Free'}
            '711com:INTUNE_A_VL' {'Microsoft Intune'}
            '711com:MCOMEETADV' {'Audio Conferencing'}
            '711com:PBI_PREMIUM_P1_ADDON' {'Power BI AddOn'}
            '711com:PBI_PREMIUM_P2_ADDON' {'Power BI AddOn'}
            '711com:POWER_BI_PRO' {'Power BI Pro'}
            '711com:POWER_BI_STANDARD' {'Power BI (free)'}
            '711com:POWERAPPS_VIRAL' {'Microsoft PowerApps Plan 2 Trial'}
            '711com:PROJECTPROFESSIONAL' {'Project Online Professional'}
            '711com:SMB_APPS' {'Business Apps (free)'}
            '711com:SPZA_IW' {'App Connect'}
            '711com:STANDARDPACK' {'Office 365 Enterprise E1'}
            '711com:STREAM' {'Microsoft Stream Trial '}
            '711com:VISIOCLIENT' {'Visio Online Plan 2'}
            '711com:WIN_DEF_ATP' {'Windows Defender Advanced Threat Protection'}
            '711com:WINDOWS_STORE' {'Windows Store'}
            default {'Unknown'}
        }
    }
}



<#
.Synopsis
   Emails a report on current license usage
.DESCRIPTION
   Emails a report of license usage if there are
   less than 10% of total licenses available for a product.
#>
function Send-LicenseReport {
    param (
        [string[]]$To = @('keith.work@7-11.com', 'Cara.Packer@7-11.com'),
        [string]$From = "keith.work@7-11.com",
        [string]$Subject = "Office 365 License Report",
        [switch]$All,
        [switch]$Test
    )

    Connect-O365

    $Products = Get-msolAccountSku | Sort-Object AccountSKUID
    [array]$Result = $Null

    Foreach ($Product in $Products){

        $DisplayName = $Product.AccountSkuId | Get-ProductDisplayName
        $AvailableUnits = $Product.ActiveUnits - $Product.ConsumedUnits
        $AvailablePercentage = ($AvailableUnits/$Product.ActiveUnits)*100

        if ($All){
            # Don't restrict results
            $Result += $Product | Select-Object @{N='Name';E={$DisplayName}},
                @{N='Active';E={$_.ActiveUnits}},
                @{N='Assigned';E={$_.ConsumedUnits}},
                @{N='Available'; E={$AvailableUnits}},
                @{N='PercentAvailable'; E={"$([math]::Round($AvailablePercentage))%"}}
        } else {
            # Restrict results
            if ($AvailablePercentage -lt 10) {
                $Result += $Product | Select-Object @{N='Name';E={$DisplayName}},
                    @{N='Active';E={$_.ActiveUnits}},
                    @{N='Assigned';E={$_.ConsumedUnits}},
                    @{N='Available'; E={$AvailableUnits}},
                    @{N='PercentAvailable'; E={"$([math]::Round($AvailablePercentage))%"}}
            }
        }
    }

    $Result | Format-List
    $Body = $Result  | ConvertTo-Html
    $Body = [string]::Join(" ",$Body)
    if ($Test){$Subject = "$Subject - TEST ONLY"}
    Send-MailMessage -SmtpServer USTXALMMB01 -To $To -From $From -Subject $Subject -Body $Body -BodyAsHtml
}



<#
.Synopsis
   Creates lists of all delegation and forwarding rules
.DESCRIPTION
   Creates 3 CSV files listing all mailboxes that are delegated, that have forwarding rules in place,
   or that use SMTP forwarding. From Microsoft.
   https://github.com/OfficeDev/O365-InvestigationTooling/blob/master/DumpDelegatesandForwardingRules.ps1
#>
function Export-DelegatesAndForwardingRules
{
    [CmdletBinding()]
    Param()

    Begin{Connect-O365}

    Process
    {
        $allUsers = @()
        $AllUsers = Get-MsolUser -All -EnabledFilter EnabledOnly |
            Select-Object ObjectID, UserPrincipalName, FirstName, LastName, StrongAuthenticationRequirements, StsRefreshTokensValidFrom, StrongPasswordRequired, LastPasswordChangeTimestamp |
            Where-Object {($_.UserPrincipalName -notlike "*#EXT#*")}

        $UserInboxRules = @()
        $UserDelegates = @()

        foreach ($User in $allUsers)
        {
            Write-Host "Checking inbox rules and delegates for user: " $User.UserPrincipalName;
            $UserInboxRules += Get-InboxRule -Mailbox $User.UserPrincipalname |
                Select-Object Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage |
                Where-Object {($Null -ne $_.ForwardTo) -or ($Null -ne $_.ForwardAsAttachmentTo) -or ($Null -ne $_.RedirectsTo)}
            $UserDelegates += Get-MailboxPermission -Identity $User.UserPrincipalName |
                Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
        }

        $SMTPForwarding = Get-Mailbox -ResultSize Unlimited |
            Select-Object DisplayName,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxandForward |
            Where-Object {$Null -ne $_.ForwardingSMTPAddress}

        $UserInboxRules | Export-Csv $env:TEMP\MailForwardingRulesToExternalDomains.csv
        $UserDelegates | Export-Csv $env:TEMP\MailboxDelegatePermissions.csv
        $SMTPForwarding | Export-Csv $env:TEMP\Mailboxsmtpforwarding.csv
    }

    End{}
}

<#
.Synopsis
   Resolves AD groups for O365 licensing
.DESCRIPTION
   Aligns membership for the E3 and E5 groups to make sure
   that each E3 and E5 license has a Mobility + Security license.
#>
function Resolve-LicenseGroups {

    # Move E3 licensees to E5
    Copy-GroupMembers -NewGroup USER-MS-Sub-SPE-E5-AdvanceFeatureSet -OldGroup USER-MS-Sub-O365-E3-AdvanceFeatureSet
    Copy-GroupMembers -NewGroup USER-MS-Sub-SPE-E5-DefaultFeatureSet -OldGroup USER-MS-Sub-O365-E3-DefaultFeatureSet
    Copy-GroupMembers -NewGroup USER-ms-sub-o365-E5-FS1 -OldGroup USER-MS-Sub-O365-E3-FS1
    Copy-GroupMembers -NewGroup USER-MS-Sub-EMS-E5 -OldGroup USER-MS-Sub-EMS-E3-DefaultFeatureSet

    # Assign E3 Mobility + Security - E3 HAS BEEN RETIRED
    # Copy-GroupMembers -OldGroupName USER-MS-Sub-O365-E3-AdvanceFeatureSet -NewGroupName USER-MS-Sub-EMS-E3-DefaultFeatureSet
    # Copy-GroupMembers -OldGroupName USER-MS-Sub-O365-E3-COOP_East -NewGroupName USER-MS-Sub-EMS-E3-DefaultFeatureSet
    # Copy-GroupMembers -OldGroupName USER-MS-Sub-O365-E3-DefaultFeatureSet -NewGroupName USER-MS-Sub-EMS-E3-DefaultFeatureSet
    # Copy-GroupMembers -OldGroupName USER-MS-Sub-O365-E3-FS1 -NewGroupName USER-MS-Sub-EMS-E3-DefaultFeatureSet

    # Assign E5 Mobility + Security
    Copy-GroupMembers -OldGroupName USER-MS-Sub-SPE-E5-AdvanceFeatureSet -NewGroupName USER-MS-Sub-EMS-E5
    Copy-GroupMembers -OldGroupName USER-MS-Sub-SPE-E5-DefaultFeatureSet -NewGroupName USER-MS-Sub-EMS-E5
    Copy-GroupMembers -OldGroupName USER-MS-Sub-SPE-E5 -NewGroupName USER-MS-Sub-EMS-E5
    Copy-GroupMembers -OldGroupName USER-MS-Sub-O365-E5-FS1 -NewGroupName USER-MS-Sub-EMS-E5

}


function Enable-OneDrive {
    [CmdletBinding()]
    param (
        [string] $UserEmail,
        [string] $SPOAdminUrl = "https://711com-admin.sharepoint.com",
        [System.Management.Automation.PSCredential]$Cred
    )
    
    begin {
        if (Get-Module Microsoft.Online.SharePoint.PowerShell) {
            Import-Module Microsoft.Online.SharePoint.PowerShell
        } else {
            Write-Error "SharePoint module required! Install using 'Install-Module Microsoft.Online.SharePoint.PowerShell' from an elevated PowerShell session"
        }
        Connect-SPOService -Url $SPOAdminUrl -Credential $Cred
    }
    
    process {
        Request-SPOPersonalSite -UserEmails $UserEmail
    }
}


function Reset-MFASettings {
     
    Param(
    [Parameter (Mandatory = $true)]
    [string] $UserPrincipalName,
    [Parameter (Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credentials
    )
 
    Connect-MSOLService -credential $credentials
 
    $MSOLUser = Get-MSOLUser -UserPrincipalName $UserPrincipalName | Select-Object FirstName
    $FirstName = $MSOLUser.FirstName
 
    if($null -eq $MSOLUser)
    {
        # $RoleAssigned = Get-MsolUserRole -UserPrincipalName $UserPrincipalName
        if($null -eq $RoleAssigned)
        {
            Try
            {
                Login-AzureRmAccount -Credential $credentials
                Connect-AzureAD -Credential $credentials
                $UserManagerMail = (Get-AzureADUserManager -ObjectId $UserPrincipalName).Mail
            }
            Catch
            {
                $ExcMessage = $_.Exception.Message
                throw "Error: Can not connect to Azure AD!. Exception: $ExcMessage"
            }
             
            Try
            {
                Set-MSOLUser -UserPrincipalName $UserPrincipalName -StrongAuthenticationMethods @()
                Write-output "MFA settings have been removed for account $UserPrincipalName"
            }
            Catch
            {
                $ExcMessage = $_.Exception.Message
                throw "Error: Can not remove MFA settings!. Exception: $ExcMessage"
            }
 
            Try
            {    
                $Body= "Dear $FirstName, <br/><br/>Please be informed that MFA settings for your account $UserPrincipalName have been reset. The next time you login to Office 365, you should be prompted to set up your multi-factor authentication. If you have any questions, please contact the Help Desk. <br /><br/> Thank You,<br /> 7 Eleven Infrastructure/Operations Team"
                $Subject = "MFA Settings Updated"
                $CredUserName = $Credentials.UserName
             
                Send-MailMessage `
                    -To $UserPrincipalName  `
                    -Cc $UserManagerMail  `
                    -Subject $Subject  `
                    -Body $Body `
                    -SmtpServer 'USTXALMMB01' `
                    -From $CredUserName `
                    -BodyAsHtml
                             
             
                Write-Output "Mail has been send!"
            }
            Catch
            {
                $ExcMessage = $_.Exception.Message
                throw "Error: Can not send email!. Exception: $ExcMessage"
            }
        }
        else
        {
                Write-Output "Account $UserPrincipalName has special role assigned and MFA settings can not be removed."
        }
         
    }
    else
    {
        Write-output "User with UPN $UserPrincipalName does not exist!"
    }
}



Function Enable-O365Users {

    param(
        $UserQueryString,
        [Parameter(Mandatory = $true, ParameterSetName = 'F1')][switch]$F1License,
        [Parameter(Mandatory = $true, ParameterSetName = 'E1')][switch]$E1License,
        [Parameter(Mandatory = $true, ParameterSetName = 'E3')][switch]$E3License,
        [Parameter(Mandatory = $true, ParameterSetName = 'E5')][switch]$E5License
    )

    $Cred = Get-Credential -Message "Domain admin credentials:" -UserName $env:USERNAME + "@7-11.com"
    
    if ($F1License){
        $Group1 = "USER-MS-SUB-SA-F1"
    }

    if ($E1License){
        $Group1 = "USER-MS-Sub-O365-E1-StoreManager "
    }

    if ($E3License){
        $Group1 = "USER-MS-Sub-EMS-E3-DefaultFeatureSet"
        $Group2 = "USER-MS-Sub-O365-E3-DefaultFeatureSet"
    }

    if ($E5License){
        $Group1 = "USER-MS-Sub-EMS-E3-DefaultFeatureSet"
        $Group2 = "USER-MS-Sub-O365-E3-DefaultFeatureSet"
    }

    Get-ADUser -Filter {$UserQueryString} -Properties mail | ForEach-Object {

        $_.name

        if ($_.mail -ne $_.userprincipalname){
            "{0} --> {1}" -f $_.userprincipalname, $_.mail
            Set-ADUser $_ -UserPrincipalName $_.mail -Confirm
        }

        Remove-ADGroupMember -Credential $Cred -Identity "nointernet-send" -Members $_ -Confirm:$false
        Add-ADGroupMember -Credential $Cred  -Identity $Group1 -Members $_ -Confirm:$false
        Add-ADGroupMember -Credential $Cred  -Identity $Group2 -Members $_ -Confirm:$false

        Enable-OneDrive -UserEmail $_.mail -Cred $Cred
    }
}




Function Complete-O365Migration {

    param(
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        $UserEmail
    )

    while ((Get-MoveRequest -Identity $UserEmail).status -ne "Completed"){
        Get-MoveRequest -Identity $_ | Set-MoveRequest -SuspendWhenReadyToComplete:$false -preventCompletion:$false -CompleteAfter (Get-Date) 
        Resume-MoveRequest -Identity $_
        Start-Sleep 5
        return (Get-MoveRequest -Identity $UserEmail)
    }
    
}



function New-O365Group {
    [CmdletBinding()]

    param (
        $DisplayName,
        $Alias,
        $OwnerEmail,
        $CreatorEmail = "keith.work@7-11.com",
        $AdditionalAlias
    )
    
    New-UnifiedGroup –DisplayName $DisplayName –Alias $Alias -EmailAddress "$Alias@7-11.com"
    Start-Sleep 5
    Set-UnifiedGroup -Identity $Alias -AccessType Private
    Add-UnifiedGroupLinks –Identity $Alias –LinkType Members -links $OwnerEmail
    Add-UnifiedGroupLinks –Identity $Alias –LinkType Owners -links $OwnerEmail
    Remove-UnifiedGroupLinks –Identity $Alias –LinkType Owners -links $CreatorEmail -Confirm:$false
    Remove-UnifiedGroupLinks –Identity $Alias –LinkType Members -links $CreatorEmail -Confirm:$false
}

function Remove-EmailFromO365Group {
    param (
        $Group,
        $EmailAddress
    )
    Get-UnifiedGroup -Identity $Group | Select-Object -ExpandProperty emailaddresses
    Set-UnifiedGroup -Identity $Group -EmailAddresses @{Remove=$EmailAddress}
}



function Export-TeamsList
{   
     param (   
           $ExportPath
           )   
    process
    {
        Connect-PnPMicrosoftGraph -Scopes "Group.Read.All","User.ReadBasic.All"
        $accesstoken =Get-PnPAccessToken
        $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/any(c:c+eq+`'Team`')" -Method Get
        $TeamsList = @()
        do
        {
            foreach($value in $group.value)
            {
		     "Group Name: " + $value.displayName + " Group Type: " + $value.groupTypes
		     $id= $value.id
		     Try
		     {
		     $team = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  https://graph.microsoft.com/beta/Groups/$id/channels -Method Get
		     "Channel count for " + $value.displayName + " is " + $team.value.id.count
		     }
		     Catch
		     {
		     "Could not get channels for " + $value.displayName + ". " + $_.Exception.Message
		     $team = $null
		     }
		     If($team.value.id.count -ge 1)
		     {
			 $Owner = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri  https://graph.microsoft.com/v1.0/Groups/$id/owners -Method Get
			 $Teams = "" | Select-Object "TeamsName","TeamType","Channelcount","ChannelName","Owners"
			 $Teams.TeamsName = $value.displayname
			 $Teams.TeamType = $value.visibility
			 $Teams.ChannelCount = $team.value.id.count
			 $Teams.ChannelName = $team.value.displayName -join ";"
			 $Teams.Owners = $Owner.value.userPrincipalName -join ";"
			 $TeamsList+= $Teams
			 $Teams=$null
		     }
             
            }
            if ($null -eq $group.'@odata.nextLink' )
            {
              break
            }
            else
            {
              $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken"} -Uri $group.'@odata.nextLink' -Method Get
            }
        }while($true);
        $TeamsList
        $TeamsList |Export-csv $ExportPath -NoTypeInformation
    }
}



function Confirm-O365License {
    param (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName)]
        $Mail,
        [switch]$F1,
        [switch]$E3,
        [switch]$E5
    )
    
    $E3 = $false
    $E3M = $false
    $E5 = $false
    $E5M = $false
    $ValidE3 = $false
    $ValidE5 = $false
    $ValidLicense = $false

    $User = Get-ADUser -Filter {mail -eq $Mail}

    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-O365-E3-DefaultFeatureSet") {$E3 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-O365-E3-AdvanceFeatureSet") {$E3 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-O365-E3-COOP_East") {$E3 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-O365-E3-FS1") {$E3 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-EMS-E3-DefaultFeatureSet") {$E3M = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-SPE-E5-DefaultFeatureSet") {$E5 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-ms-sub-o365-E5-FS1") {$E5 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-SPE-E5") {$E5 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-SPE-E5-AdvanceFeatureSet") {$E5 = $True}
    if (Confirm-GroupMembership -User $User -GroupName "USER-MS-Sub-EMS-E5") {$E5M = $True}
    if ($E3 -and $E3M){$ValidE3 = $True}
    if ($E5 -and $E5M){$ValidE5= $True}
    if ($ValidE5){$ValidLicense = $True}
    if ($ValidLicense){return $True} else {
        Write-Warning "No valid license assigned to $Mail. Use Add-O365License to add one."
        return $false}
}




function Assign-O365License {
    param (
        $UserPrincipalName,
        [switch]$F1,
        [switch]$E3,
        [switch]$E5
    )
    
    Process{

        $User = Get-ADUser -Filter {UserPrincipalName -eq $UserPrincipalName}

        if ($F1){
            # Add F1
            Write-Warning "Assigning F1 to $UserPrincipalName"
            Add-ADGroupMember -Identity USER-MS-SUB-SA-F1 -Members $User -Credential $AdminCred
        }

        if ($E3){
            # Add E3
            Write-Warning "Assigning E3 to $UserPrincipalName"
            Add-ADGroupMember -Identity USER-MS-Sub-O365-E3-DefaultFeatureSet -Members $User -Credential $AdminCred
            Add-ADGroupMember -Identity USER-MS-Sub-EMS-E3-DefaultFeatureSet -Members $User -Credential $AdminCred
        }

        if ($E5){
            # Add E5
            Write-Warning "Assigning E5 to $UserPrincipalName"
            Add-ADGroupMember -Identity USER-MS-Sub-SPE-E5-DefaultFeatureSet -Members $User -Credential $AdminCred
            Add-ADGroupMember -Identity USER-MS-Sub-EMS-E5 -Members $User -Credential $AdminCred
        }
    }
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
                return "Safe sender $_ already in the list"
            } else {
                # Add the safe sender to the new list
                "Adding safe sender $_"
                Get-Mailbox $Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $NewSafeSenders
            }
        }
    }
}



function Add-UsersToBookinPolicy {
    [CmdletBinding()]
    param (
        # List of email addresses in string format
        [array]$UserEmailAddresses,

        # Display name of the room to add the users to
        [string]$RoomName
    )
    
    foreach($Address in $UserEmailAddresses){
        Set-CalendarProcessing $RoomName -BookInPolicy ((Get-CalendarProcessing -Identity $RoomName).BookInPolicy += $Address) 
    }
}


function Set-StandardBookinPolicy {
    param (
        $RoomEmail
    )
    if (Get-CalendarProcessing $RoomEmail){
        Set-CalendarProcessing -Identity $RoomEmail -bookingwindowindays 425 -maximumdurationinminutes 720 -maximumconflictinstances 25 -conflictpercentageallowed 25 -allowconflicts $false -allowrecurringmeetings $true -scheduleonlyduringworkhours $false -enforceschedulinghorizon $true -deletesubject $false -forwardrequeststodelegates $false 
    } else {
        Write-Warning "No policy found for $RoomEmail. Does it exist? Is it a room?"
    }
}


# Usage: Get-StoreForwardingAddress 25290
# Usage: 25290, 26803 | Get-StoreForwardingAddress -verbose
# Returns an email address (string value) like boardwalk_ftcollins@monfortcompanies.com. Use -verbose for more info.
function Get-StoreForwardingAddress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]$StoreNumber
    )
    process{
        $Recipients = Get-DistributionGroupMember "storemanager$StoreNumber@7-11.com"
        foreach ($Recipient in $Recipients) {
            If ($Recipient.RecipientType -eq "UserMailbox"){
                $RecipientMailbox = $Recipient | Get-Mailbox
                if ($RecipientMailbox.ForwardingAddress){
                    $ForwardingTarget = Get-MailContact $RecipientMailbox.ForwardingAddress
                    Write-Host -ForegroundColor Yellow "Group: storemanager$StoreNumber@7-11.com --> Member: $($RecipientMailbox.Name) ($($RecipientMailbox.PrimarySmtpAddress)) --> Contact: $($RecipientMailbox.ForwardingAddress) ($($ForwardingTarget.PrimarySmtpAddress))"
                } else {
                    Write-Host -ForegroundColor Yellow "Group: storemanager$StoreNumber@7-11.com --> Member: $($RecipientMailbox.Name) ($($RecipientMailbox.PrimarySmtpAddress)) --> NO FORWARDING ENABLED"
                }
            } elseif ($Recipient.RecipientType -eq "MailUser") {
                $RecipientContact = $Recipient | Get-MailUser 
                Write-Host -ForegroundColor Yellow "Group: storemanager$StoreNumber@7-11.com --> Member: $($RecipientContact.PrimarySmtpAddress) (Mail user, not a mailbox - Possibly ON-PREM)"
            } else {
                Write-Warning "Invalid recipient type: $($Recipient.RecipientType)"
                $Recipient
            }
        }
    }
}


Connect-O365