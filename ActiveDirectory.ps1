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

function Copy-GroupMembership
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        $OldGroupName,
        $NewGroupName,
        $Limit,
        $LogPath = "C:\Temp\Copy-GroupMembership.log"
    )

    Begin{}

    Process{
    
        $Count = 0


        $searchRoot = New-Object System.DirectoryServices.DirectoryEntry
        $adSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $adSearcher.SearchRoot = $searchRoot

        $adSearcher.Filter = "(cn=$OldGroupName)"

        $adSearcher.PropertiesToLoad.Add("member")
        $samResult = $adSearcher.FindOne()

        if($samResult)
        {
            $adAccount = $samResult.GetDirectoryEntry()
            $OldGroupMembers = $adAccount.Properties["member"]
        }
        
        if ($Limit) {$OldGroupMembers = $OldGroupMembers[0..($Limit-1)]}

        $OldGroupMembersSum = ($OldGroupMembers | Measure-Object).count
        $OldGroupMembers | Get-ADUser | ForEach-Object{

            $User = $_.Samaccountname
            $Count++

            try{
                Add-ADGroupMember -Credential $AdminCred -GroupName $NewGroupName -Members $_
                "$(Get-Date) : $Count of $OldGroupMembersSum : $NewGroupName + $User : SUCCESS" | Tee-Object -FilePath $LogPath -Append
            } catch {
                "$(Get-Date) : $Count of $OldGroupMembersSum : $NewGroupName + $User  : $($_.exception.message)" | Tee-Object -FilePath $LogPath -Append
            }

            Start-Sleep 1
        }
    }

    End{}
}

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
            $Confirmed = Confirm-GroupMembership $_ -GroupName $NewGroup
            # If not in new group, add user
            if (! $Confirmed) {
                # Add to preferred group and confirm it
                Add-ADGroupMember -Credential $AdminCred -Identity $NewGroup -Members $_
                Start-Sleep 5        
                $Confirmed = Confirm-GroupMembership $_ -GroupName $NewGroup
            }

            # If confirmed in the new group
            if ($Confirmed) {
                # Remove from old group
                Remove-ADGroupMember -Credential $AdminCred -Identity $OldGroup -Members $_ -Confirm
                "Confirmed"
            }
        }
    }
    
    end {
    }
}



# Copy members of one group to another
function Copy-GroupMembership {
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
        [Parameter(Mandatory=$true, ValueFromPipeline)]
        $EmailAddress,
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

    $User = Get-ADUser -Filter {mail -eq $EmailAddress}

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
    if ($ValidE3 -or $ValidE5){$ValidLicense = $True}
    if ($ValidLicense){return $True} else {
        Write-Warning "No license assigned to $EmailAddress. Use Add-O365License to add one."
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