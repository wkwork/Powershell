Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'

Import-Module activedirectory

if ($AdminCred) {
    Write-Warning "Using saved credentials..."
} else {
    [System.Management.Automation.PSCredential]$AdminCred = Get-Credential -Message "Domain Admin Credential"
}

<#
.Synopsis
   Reads the membership of one group and adds those members to another group
.DESCRIPTION
   Reads the membership of the old group (up to the limit if any) and adds each of
   those members to the new group. Log is written at C:\Temp\Copy-GroupMembership.log
   Keith Work 11/1/18
   Added admin group to script 8/30/19
#>


<#
.Synopsis
   Removes all group members from an AD group
.DESCRIPTION
   Clears all members from the given group. If a limit is
   specified, it only clears that many users, first come first served.
   Log is written at C:\Temp\Clear-GroupMembership.log
   Keith Work 11/1/18
   Added admin group to script 8/30/19
#>
function Clear-GroupMembership {
    [CmdletBinding()]
    param (
        $GroupName,
        $Limit,
        [switch]$DisabledOnly,
        $LogPath = "C:\Temp\Clear-GroupMembership.log"
    )
    
    begin {}
    
    process {

        try {
            if ($Limit){(Get-ADGroupMember $GroupName)[0..$Limit-1] | Remove-ADGroupMember -Credential $AdminCred -Identity $GroupName -Members $_ -Confirm}
            else {Get-ADGroupMember $GroupName | Remove-ADGroupMember -Credential $AdminCred -Identity $GroupName -Members $_ -Confirm}
            "$(Get-Date) : $GroupName - $User : SUCCESS" | Tee-Object -FilePath $LogPath -Append
        } catch {
            "$(Get-Date) : $GroupName - $User : $($_.exception.message)" | Tee-Object -FilePath $LogPath -Append
        }
    }
    
    end {}
}


<#
.Synopsis
   Confirms the given user is a member of the given group
.DESCRIPTION
   Reads the group membership for the user (rather than reading
   the group members which could be bajillions) and compares each group
   they are in to the target group given.
   Keith Work 8/30/2019
#>
function Confirm-GroupMembership {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        [Microsoft.ActiveDirectory.Management.ADUser]$User,
        $GroupName
    )
    
    process {
        [array]$grps=Get-ADUser $user -Property memberOf | Select-Object -ExpandProperty memberOf | Get-ADGroup | Select-Object Name

        $GroupCheckResult = $null
        foreach($grp in $grps){
            write-verbose "Comparing $($grp.name) to target $Groupname"
            if($grp.Name -eq $GroupName){
                $GroupCheckResult = $true 
                write-verbose "Member"
            }
        }
        if ($GroupCheckResult -eq $true){
            return $true
        } else {
            return $false
        }
    }
}




# Move users from one AD group to another
#  Added admin group to script 8/30/19
function Move-GroupMembership {
    [CmdletBinding()]
    param (
            $NewGroup,
            $OldGroup
    )
    
    begin {
        Import-Module "\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules\ActiveDirectory.ps1"
    }
    
    process {
        # Get members of deprecated license group
        Get-ADGroupMember $OldGroup | Get-ADUser | ForEach-Object {
            "{0} : {1} -> {2}" -f $_.Name, $OldGroup, $NewGroup
            Pause
            $Confirmed = Confirm-GroupMembership $_ -GroupName $NewGroup
            # If not in new group, add user
            if (! $Confirmed) {
                # Add to preferred group and confirm it
                Add-ADGroupMember -Credential $AdminCred -Identity $NewGroup -Members $_ -Confirm:$false
                Start-Sleep 5        
                $Confirmed = Confirm-GroupMembership $_ -GroupName $NewGroup
            }

            # If confirmed in the new group
            if ($Confirmed) {
                # Remove from old group
                Remove-ADGroupMember -Credential $AdminCred -Identity $OldGroup -Members $_  -Confirm:$false
                "Confirmed"
            }
        }
    }
    
    end {
    }
}


# Add one user to each group that another user is assigned to, duplicating that user's group membership
function Copy-GroupMembership {
    [CmdletBinding()]
    param (
        [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser,
        [Microsoft.ActiveDirectory.Management.ADUser]$TargetUser
    )
    
    $Groups = (Get-ADUser $SourceUser -Properties memberOf).memberOf

    Foreach($Group in $Groups) {
        "Adding {0} to {1}" -f $TargetUser.Name, $Group
        Get-ADGroup -Identity $Group
        Add-ADGroupMember -Identity $Group -Members $TargetUser -Confirm
    }
}



# Copy members of one group to another
function Copy-GroupMembers {
    [CmdletBinding()]
    param (
            $NewGroup,
            $OldGroup
    )
    
    begin {
        Import-Module "\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules\ActiveDirectory.ps1"
    }
    
    process {
        # Get members of old group
        Get-ADGroupMember $OldGroup | Get-ADUser | ForEach-Object {
            "{0} : {1} + {2}" -f $_.Name, $OldGroup, $NewGroup
            Add-ADGroupMember -Credential $AdminCred -Identity $NewGroup -Members $_
            Start-Sleep 5        
            $Confirmed = Confirm-GroupMembership $_ -GroupName $NewGroup

            # If confirmed in the new group
            if ($Confirmed) {
                "Copy Confirmed"
            }
        }
    }
}



# Find members of 2 groups
function Find-DualGroupMembership {
    [CmdletBinding()]
    param (
            $Group1,
            $Group2
    )
    
    begin {
        Import-Module "\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules\ActiveDirectory.ps1"
    }
    
    process {
        # Get members of old group
        Get-ADGroupMember $Group1 | Get-ADUser -properties Title, Mail | ForEach-Object {
            if (Confirm-GroupMembership $_ -GroupName $Group2) {
                return $_ 
            }
        }
    }
}


# Removed disabled users from a given group - MUST BE RUN AS ADMIN
function Remove-DisabledUsersFromGroup
{

    Param
    (
        $TargetGroupName = "USER-MS-Sub-O365-E3-DefaultFeatureSet",
        $SearchBase
    )

    Get-ADGroupMember $TargetGroupName | Get-ADUser -Properties Description | Where-Object Enabled -eq $False | ForEach-Object{

        if ($_.Description -notlike "*leave of absence*"){
            Write-Host " "
            "$(get-date) : Removing $($_.name) ($($_.description)) from $TargetGroupName" | Tee-Object "C:\Temp\Remove-DisabledUsersFromGroup.log" -Append
            Remove-ADGroupMember -Credential $AdminCred -Identity $TargetGroupName -Members $_.Samaccountname -Confirm
        } else {
            "$(get-date) : $($_.name) is on LOA. NOT REMOVING from $TargetGroupName"
        }
    } 
}


# Remove disabled from users specifically from O365 groups
function Remove-DisabledUsersFromO365LicenseGroups
{

    "USER-MS-Sub-O365-E3-DefaultFeatureSet",
    "USER-MS-Sub-EMS-E3-DefaultFeatureSet",
    "USER-MS-Sub-SPE-E5-DefaultFeatureSet",
    "USER-MS-Sub-EMS-E5" | ForEach-Object {

        Remove-DisabledUsersFromGroup -TargetGroupName $_ -SearchBase "OU=User Employees,OU=Corp,DC=7-11,DC=com"
        Remove-DisabledUsersFromGroup -TargetGroupName $_ -SearchBase "OU=User Contractors,OU=Corp,DC=7-11,DC=com"
    }
}




function Remove-RedundantGroupMembers
{
    Param
    ($PreferredGroup,
    $ReduntantGroup)
    
    Find-DualGroupMembership $PreferredGroup $ReduntantGroup | ForEach-Object {
        "Removing $($_.name) from $ReduntantGroup" | Tee-Object "C:\Temp\Remove-RedundantGroupMembers.log" -Append
        Remove-ADGroupMember -Credential $AdminCred -Identity $ReduntantGroup -Members $_ -Confirm:$False
    }
}





# Find and fix users with incorrect UserPrincipalName
function Update-UserPrincipalName
{
    Get-ADUser -Filter {(ExtensionAttribute8 -eq "Corporate Employee" -or ExtensionAttribute8 -eq "Partner Account") -and (Enabled -eq $True)} -Properties mail, CN | Where-Object mail -Like "*7-11.com" | ForEach-Object {
        if ($_.userprincipalname -ne $_.mail){
            "{0} : Updating UPN from {1} to {2}" -f $_.Name, $_.UserPrincipalName, $_.Mail | Tee-Object "c:\temp\ADUpdates.log" -Append
            Try {
                $_ | Set-ADUser -Credential $AdminCred -UserPrincipalName $_.Mail -Confirm -ErrorAction Stop
            } catch {
                $_
                "FAILED $($_.Exception.Message)" | Tee-Object "c:\temp\ADUpdates.log" -Append
            }
        }
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




# Useage Get-LockoutServer s1865005
function Get-LockoutServer {
    param (
        $Username
    )
    ## Find the domain controller PDCe role
    $Pdce = (Get-AdDomain).PDCEmulator

    ## Query the security event log
    $Events = Get-WinEvent -Credential $AdminCred -ComputerName $Pdce -LogName 'Security' `
    -FilterXPath "*[System[EventID=4740] and EventData[Data[@Name='TargetUserName']='$Username']]"

    $Events | ForEach-Object {
        $_.properties[1].value
    } 
}



function Get-RemoteRecipientType {
    param (
        $SamAccountName
    )
    
    if ($User = Get-ADUser $SamAccountName -Properties msExchRemoteRecipientType) {

        switch ($User.msExchRemoteRecipientType) {
            "1"     {return "ProvisionMailbox"}                                             # Users
            "2"     {return "ProvisionArchive (On-Prem Mailbox)"}
            "3"     {return "ProvisionMailbox, ProvisionArchive"}
            "4"     {return "Migrated (UserMailbox)"}
            "6"     {return "ProvisionArchive, Migrated"}
            "8"     {return "DeprovisionMailbox"}
            "10"    {return "ProvisionArchive, DeprovisionMailbox"}
            "16"    {return "DeprovisionArchive (On-Prem Mailbox)"}
            "17"    {return "ProvisionMailbox, DeprovisionArchive"}
            "20"    {return "Migrated, DeprovisionArchive"}
            "24"    {return "DeprovisionMailbox, DeprovisionArchive"}
            "32"    {return "RoomMailbox"}                                                  # Rooms
            "33"    {return "ProvisionMailbox, RoomMailbox"}
            "35"    {return "ProvisionMailbox, ProvisionArchive, RoomMailbox"}
            "36"    {return "Migrated, RoomMailbox"}
            "38"    {return "ProvisionArchive, Migrated, RoomMailbox"}
            "49"    {return "ProvisionMailbox, DeprovisionArchive, RoomMailbox"}
            "52"    {return "Migrated, DeprovisionArchive, RoomMailbox"}
            "64"    {return "EquipmentMailbox"}                                             # Equipment
            "65"    {return "ProvisionMailbox, EquipmentMailbox"}
            "67"    {return "ProvisionMailbox, ProvisionArchive, EquipmentMailbox"}
            "68"    {return "Migrated, EquipmentMailbox"}
            "70"    {return "ProvisionArchive, Migrated, EquipmentMailbox"}
            "81"    {return "ProvisionMailbox, DeprovisionArchive, EquipmentMailbox"}
            "84"    {return "Migrated, DeprovisionArchive, EquipmentMailbox"}
            "96"    {return "SharedMailbox"}                                                # Shared Mailboxes
            "100"   {return "Migrated, SharedMailbox"}
            "102"   {return "ProvisionArchive, Migrated, SharedMailbox"}
            "116"   {return "Migrated, DeprovisionArchive, SharedMailbox"}
        }
    }

}



function Get-RecipientType {
    param (
        $SamAccountName
    )
    
    if ($User = Get-ADUser $SamAccountName -Properties msExchRecipientDisplayType) {

        switch ($User.msExchRecipientDisplayType) {
            "-2147483642"   {return "MailUser (RemoteUserMailbox)"}
            "-2147481850"   {return "MailUser (RemoteRoomMailbox)"}
            "-2147481594"   {return "MailUser (RemoteEquipmentMailbox)"}
            "0"             {return "UserMailbox (shared)"}
            "1"             {return "MailUniversalDistributionGroup"}
            "2"             {return "Public Folder"}
            "3"             {return "Dynamic Distribution Group"}
            "4"             {return "Organization"}
            "5"             {return "Private Distribution List"}
            "6"             {return "MailContact"}
            "7"             {return "UserMailbox (room)"}
            "8"             {return "UserMailbox (equipment)"}
            "1073741824"    {return "ACL able Mailbox User"}
            "1073741833"    {return "MailUniversalSecurityGroup"}
        }
    }
}



function Get-RecipientTypeDetails {
    param (
        $SamAccountName
    )
    
    if ($User = Get-ADUser $SamAccountName -Properties msExchRecipientTypeDetails) {

        switch ($User.msExchRecipientTypeDetails) {
            "1"             {return "User Mailbox"}
            "2"             {return "Linked Mailbox"}
            "4"             {return "Shared Mailbox"}
            "8"             {return "Legacy Mailbox"}
            "16"            {return "Room Mailbox"}
            "32"            {return "Equipment Mailbox"}
            "64"            {return "Mail Contact"}
            "128"           {return "Mail User"}
            "256"           {return "Mail-Enabled Universal Distribution Group"}
            "512"           {return "Mail-Enabled Non-Universal Distribution Group"}
            "1024"          {return "Mail-Enabled Universal Security Group"}
            "2048"          {return "Dynamic Distribution Group"}
            "4096"          {return "Public Folder"}
            "8192"          {return "System Attendant Mailbox"}
            "16384"         {return "System Mailbox"}
            "32768"         {return "Cross-Forest Mail Contact"}
            "65536"         {return "User"}
            "131072"        {return "Contact"}
            "262144"        {return "Universal Distribution Group"}
            "524288"        {return "Universal Security Group"}
            "1048576"       {return "Non-Universal Group"}
            "2097152"       {return "Disabled User"}
            "4194304"       {return "Microsoft Exchange"}
            "8388608"       {return "Arbitration Mailbox"}
            "16777216"      {return "Mailbox Plan"}
            "33554432"      {return "Linked User"}
            "268435456"     {return "Room List"}
            "536870912"     {return "Discovery Mailbox"}
            "1073741824"    {return "Role Group"}
            "2147483648"    {return "Remote Mailbox"}
            "137438953472"  {return "Team Mailbox"}
        }
    }

}


# ---------------------------------- Working ------------------------------
<#
    NOT TESTED    
    "34069","36043","11160","11232","23924", "32120","34083"
#>
function Set-StoreIpadPassword {
    param (
        $StoreNumber,
        $NewPassword = "Pa55word",
        $AdminCred = (Get-Credential -UserName "7-11\kwork006" -Message "Domain Admin Creds")
    )

    if ($Account = Get-ADUser ("sa-s" + $StoreNumber + "device")) {
        "Setting {0} to {1}" -f $Account.DistinguishedName, $NewPassword
        $Account | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$NewPassword" -Force) -Credential $AdminCred
        Pause
    } else {
        Write-Warning "sa-s" + $StoreNumber + "device not found"
    }
}