$MailData = import-csv c:\users\jowen035\documents\input.csv
foreach($m in $MailData)
{
Set-Mailbox $m.user -LitigationHoldEnabled $true
Get-Mailbox $m.user | where-object {($_.litigationholdenabled -eq "true")} | fl Name
}