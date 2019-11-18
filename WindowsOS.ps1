

<#
.Synopsis
   Queries the given computer for the given registry value
.EXAMPLE
   Get-RegistryValue -Computername ustxalmmb01 -RegKeyPath 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -RegKeyName 'Enabled'
   Checks the value given an explicit machine name
.EXAMPLE
   'ustxalmmb01' | Check-RegistryKey -RegKeyPath 'SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client' -RegKeyName 'Enabled'
   Checks a registry value froma computer name passed in as a string
.NOTES
   Keith Work - 10/18/18

#>
function Get-RegistryValue
{
    [CmdletBinding()]
    Param
    (
        [Parameter(ValueFromPipeline=$True)]
        [string]$Computername = $env:COMPUTERNAME,

        # Like 'LocalMachine'
        [string]$RegKeyHive = 'LocalMachine',

        # Like 'SOFTWARE\Microsoft\Internet Explorer'
        [string]$RegKeyPath,

        # Like 'Build'
        [string]$RegKeyName
    )

    Begin{}

    Process
    {
        try{
            ([Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegKeyHive, $Computername)).OpenSubKey($RegKeyPath).getvalue($RegKeyName)
        } catch {
            return "Value not found: $RegKeyPath\$RegKeyName"
        }
    }

    End{}
}




<#
.Synopsis
Queries a computer to check for interactive sessions

.DESCRIPTION
This script takes the output from the quser program and parses this to PowerShell objects

.NOTES   
Name: Get-LoggedOnUser
Author: Jaap Brasser
Version: 1.2.1
DateUpdated: 2015-09-23

.LINK
http://www.jaapbrasser.com

.PARAMETER ComputerName
The string or array of string for which a query will be executed

.EXAMPLE
.\Get-LoggedOnUser.ps1 -ComputerName server01,server02

Description:
Will display the session information on server01 and server02

.EXAMPLE
'server01','server02' | .\Get-LoggedOnUser.ps1

Description:
Will display the session information on server01 and server02
#>
function Get-ServerSessions {
    
    [CmdletBinding()] 

    param(
        [Parameter(ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
        [string[]]$ComputerName = 'localhost'
    )
    begin {
        $ErrorActionPreference = 'Stop'
    }

    process {
        foreach ($Computer in $ComputerName) {
            try {
                quser /server:$Computer 2>&1 | Select-Object -Skip 1 | ForEach-Object {
                    $CurrentLine = $_.Trim() -Replace '\s+',' ' -Split '\s'
                    $HashProps = @{
                        UserName = $CurrentLine[0]
                        ComputerName = $Computer
                    }

                    # If session is disconnected different fields will be selected
                    if ($CurrentLine[2] -eq 'Disc') {
                            $HashProps.SessionName = $null
                            $HashProps.Id = $CurrentLine[1]
                            $HashProps.State = $CurrentLine[2]
                            $HashProps.IdleTime = $CurrentLine[3]
                            $HashProps.LogonTime = $CurrentLine[4..6] -join ' '
                            $HashProps.LogonTime = $CurrentLine[4..($CurrentLine.GetUpperBound(0))] -join ' '
                    } else {
                            $HashProps.SessionName = $CurrentLine[1]
                            $HashProps.Id = $CurrentLine[2]
                            $HashProps.State = $CurrentLine[3]
                            $HashProps.IdleTime = $CurrentLine[4]
                            $HashProps.LogonTime = $CurrentLine[5..($CurrentLine.GetUpperBound(0))] -join ' '
                    }

                    New-Object -TypeName PSCustomObject -Property $HashProps |
                    Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
                }
            } catch {
                New-Object -TypeName PSCustomObject -Property @{
                    ComputerName = $Computer
                    Error = $_.Exception.Message
                } | Select-Object -Property UserName,ComputerName,SessionName,Id,State,IdleTime,LogonTime,Error
            }
        }
    }
}



function Get-ServerProfile {

param(
    $UserID,
    $Server
)
    $Result = $Null
    $Directory = "\\" + $Server + "\C$\Users\$UserID"
    if ($Result = Get-Item $Directory -ErrorAction SilentlyContinue){
        $Result.FullName
    } else {
        Write-Host "$Directory not found" -ForegroundColor Yellow
    }
}


function Get-ServerProfiles {
    param (
        $User,
        [string[]]$Servers = @("ustxalajts01", "ustxalajts02", "ustxalajts03", "ustxalajts04", "ustxalajts05", "ustxalajts06", "ustxalajts07", "ustxalajts08", "ustxalajts09", "ustxalajts10",
        "ustxalajts11", "ustxalajts12", "ustxalajts13", "ustxalajts14", "ustxalajts15", "ustxalajts16", "ustxalajts17", "ustxalajts18", "ustxalajts19", "ustxalajts20")
    )
    
    foreach ($Server in $Servers){
        Get-ServerProfile $User, $Server
    }
}