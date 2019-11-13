

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




