Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'
Import-Module .\common.ps1

[System.Management.Automation.PSCredential]$AdminCred = Get-Credential -Message "Domain Admin Credential"


function Connect-SkypeForBusiness {
    [CmdletBinding()]
    param (
        $Pool = "lyncpool2",
        [System.Management.Automation.PSCredential]$Cred = $AdminCred
    )
    
    process {
        $session = New-PSSession -Name "Skype2015" -ConnectionURI "https://$Pool.7-11.com/OcsPowershell" -Credential $Cred
        Import-PsSession $session
    }
    
}


# Example: Enable-SkypeVoice -User Anjeer.Amol@7-11.com -Extension 86001
function Enable-SkypeUserVoice
{
    Param
    (
        # Target user's email address
        [Parameter(Mandatory=$true)]
        [string]$TargetUserEmail,

        # 5 digit extension
        [Parameter(Mandatory=$true)]
        [string]$Extension
    )

    "Enabling $TargetUserEmail..."

    if (Get-CsUser $User){
        # Good to go
    } else {
        Get-CsAdUser aamol001 | Enable-CsUser -RegistrarPool lyncpool2.7-11.com -SipAddress "sip:$TargetUserEmail"
        sleep 10
    }

    if (Get-CsUser $TargetUserEmail) {        
        Get-CsUser $TargetUserEmail | Set-CsUser -EnterpriseVoiceEnabled $true -LineUri "tel:+197282$Extension;ext=$Extension"
        Get-CsUser $TargetUserEmail | Grant-CsVoicePolicy -PolicyName "7-11 National"
        Get-CsUser $TargetUserEmail | Grant-CsConferencingPolicy -PolicyName "Video Enabled"
        Get-CsUser $TargetUserEmail | Grant-CsLocationPolicy -PolicyName "E911"
    } else {
        return "$TargetUserEmail not found!"
    }
}



# Finds users who are migrated to O365 but are missing the CloudUM voice policy

function Get-SkypeUsersWithMissingPolicy
{
    [CmdletBinding()]
    Param
    (
    )

    Process
    {
        $CSUser = $Null

        Get-ADUser -Filter * -Properties displayname, msExchRecipientTypeDetails, mail |where msExchRecipientTypeDetails -gt 1 | ForEach-Object{
            if ($CSUser = Get-CsUser "sip:$($_.mail)" -ErrorAction SilentlyContinue){
                $CSUser | where enabled -eq "True" | where EnterpriseVoiceEnabled -eq "True" | where HostedVoicemailPolicy -eq $Null
            }
        }
    }
}


# For users migrated to O365, they need their voice mail policy updated

function Grant-SkypeUserCloudPolicy
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$TargetUserEmail
    )

    Process
    {
        Grant-CsHostedVoicemailPolicy $TargetUserEmail -PolicyName Cloudum 
        do {$User = get-csuser "sip:$TargetUserEmail"; sleep 5} while ($User.HostedVoicemailPolicy -eq $Null)
        return $User
    }
}


Connect-SkypeForBusiness