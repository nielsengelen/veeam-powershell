# Script to gather mailbox size in Office 365 
#
# Last update 26/09/2017 - Niels Engelen - @nielsengelen

$LiveCred = Get-Credential 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic –AllowRedirection
Import-PSSession $Session
Get-Mailbox | Get-Mailboxstatistics | ft DisplayName, TotalItemSize
Remove-PSSession $Session 