Push-Location '\\7-encrypt\cssdocs$\Script Repository\PowerShell\Modules'

Import-Module .\ActiveDirectory.ps1
Import-Module .\Office365.ps1

function Write-LogFile
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   Position=0)]
        $String,

        $LogFile = "\\ustxirscifs01\Field_Tablet_Refresh$\_Automation\Logs\Staging.log"
    )

    Begin{}

    Process
    {
        $Timestamp = Get-Date -Format "yyyyMMdd-hhmmss"
        "$Timestamp | $String" | tee-object $LogFile -Append
    }

    End{}
}


function Reset-UserPassword
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Param1 help description
        $SourceFile = "\\ustxirscifs01\Field_Tablet_Refresh$\_UPS_Shipping_Share\Field Tablet Refresh_TEST_Password_Reset.csv",
        [System.Management.Automation.PSCredential]$Cred

    )

    $UserRecords = Import-Csv $SourceFile

    foreach ($UserRecord in $UserRecords) {

        if ($Null -eq $Cred) {$Cred = Get-Credential}

        $Password = ConvertTo-SecureString -String $UserRecord.KIRKFinal1 -AsPlainText –Force

        Get-ADUser $UserRecord.User_ID | Set-ADAccountPassword -NewPassword $Password -Reset -Confirm -Credential $Cred
        Write-LogFile "$($UserRecord.User_ID) : Resetting password"
    }
}



function New-CSVForO365
{
    [CmdletBinding()]
    Param
    (
        # Path to the folder containing the PST files
        $RootPath
    )


    if (Test-Path $RootPath) {

        [array]$FileCollection = $Null
        $Timestamp = Get-Date -Format "yyyyMMdd"

        Get-ChildItem "$RootPath\*.pst" | ForEach-Object {
            $UserName = ($_.Name -split "-")[0]
                
            if ($UPN = (Get-ADUser $UserName).UserPrincipalName) {
                $FileCollection += $_ | Select-Object @{n="Workload";e={"Exchange"}}, `
                            @{n="FilePath";e={$TimeStamp}}, `
                            Name, `
                            @{n="Mailbox";e={$UPN}}, `
                            @{n="IsArchive";e={"TRUE"}}, `
                            @{n="TargetRootFolder";e={$_.Name}}, `
                            @{n="ContentCodePage";e={}}, `
                            @{n="SPFileContainer";e={}}, `
                            @{n="SPManifestContainer";e={}}, `
                            @{n="SPSiteUrl";e={}}
            } else {
                Write-LogFile "$Username not found in AD!"
            }
        }
        $FileCollection | Export-Csv "$RootPath\$Timestamp.csv" -NoTypeInformation -Force

    } else {
        Write-LogFile "$RootPath does not exist!"
    }
}



function Confirm-FileMoveComplete
{
    [CmdletBinding()]
    Param
    (
        # The file to evaluate
        $FileName
    )

    Process
    {
        if (test-path $FileName){
            $FirstCheck = Get-item $FileName
            Start-Sleep 5
            $SecondCheck = Get-item $FileName

            if ($FirstCheck.LastWriteTime -eq $SecondCheck.LastWriteTime){
                return $True
            } else {
                return $False
            }
        }
    }

}


function Move-PstFilesForImport
{
    Param
    (
        $RootPath = "\\ustxirscifs01\Field_Tablet_Refresh$"
    )

    Begin
    {
        [array]$PSTFiles = $null
    }
    Process
    {
        $StagingFolder = Get-Date -Format "yyyyMMdd"
        $StagingFolder = "$RootPath\_Automation\Staging-$StagingFolder"
        New-Item $StagingFolder -ItemType Directory

        # Get the PST files
        $PSTFiles += Get-ChildItem $RootPath\*.pst -Recurse | 
            Where-Object fullname -NotLike "*_Automation*" | 
            Where-Object fullname -NotLike "Sharepoint.pst" |
            Where-Object length -GT 255kb

        # for each
        foreach($PSTFile in $PSTFiles){
            # Verify completion
            if (Confirm-FileMoveComplete $PSTFile.FullName) {

                Write-LogFile "File is ready : $($PSTFile.fullname)"
                # Rename
                $Username = ($PSTFile.fullname -split ("\\"))[4]
                $Timestamp = Get-Date -Format "hhmmss"
                $TargetName = $Username + "-" + $Timestamp + "-" + $PSTFile.Name

                # Move  file
                if  (Get-ADUser $UserName -erroraction silentlycontinue) {
                    Move-Item $PSTFile.fullname -Destination "$StagingFolder\$TargetName"
                    Write-LogFile "Moved $($PSTFile.fullname)"
                    Write-LogFile "New file: $("$StagingFolder\$TargetName")"
                    Start-Sleep 2
                } else {
                    Write-LogFile "ERROR Invalid user: $Username"
                }
            }
        }

        Generate-CSVForO365 $StagingFolder
    }

    End{}
}



function Confirm-MoveRequestPrep {
    param (
        [parameter(ValueFromPipeline)]
        $Email,
        $CompletionDateTime
    )

    $Result = [pscustomobject]@{
        User = $Email
        Intune = "ERROR"
        Office = "ERROR"
        Migration = "ERROR"
    }

    if ($User = Get-ADUser -Filter {mail -eq $Email} -Properties Mail, Title, Description) {
        $Result.User = $User.Mail
        $Result.Intune = Confirm-GroupMembership -User $User -GroupName "App-Intune"
        $Result.Office = Confirm-O365License -Mail $User.Mail
        $Result.Migration = (Get-MoveRequest -Identity $User.Mail -ErrorAction silentlycontinue).Status
    }

    if ($Result.Intune -ne $true){
        $Result.Intune = "Add"
        if ($User = Get-ADUser -Filter {mail -eq $Email}){
            Add-ADGroupMember -Identity app-intune -Members $User -Credential $AdminCred
        } else {Write-Warning "Unable to add [$Email] to App-Intune"}
    }

    $Result | Select-Object User, Office, Intune, Migration
}