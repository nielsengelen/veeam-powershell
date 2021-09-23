<# 
.NAME
    Veeam Backup for Microsoft Office 365 Mail per job 
.SYNOPSIS
    Script to use for reporting which organizations don't have a contact email
.DESCRIPTION
    Script to use for reporting which organizations don't have a contact email

    Utilises CSV files as a databases
    - organizations.csv (contains contact information)

    Released under the MIT license.
.LINK
    http://www.github.com/nielsengelen
#>


# Modify the values below to your needs
# Mail server configuration
$from = "vbo365@company.com"
$to = "mailbox@company.com"
$smtpserver = "mail.company.com"
$mailSubject = "[VB365] Organizations with missing contact person"
$port = "587" # default: 25
$usessl = $True # Use SSL ($True) or not ($False)

# Authentication against the mail server
$username = "authentication@company.com"
$password = ConvertTo-SecureString "AUTHPASSWORD" -AsPlainText -Force

# Do not change below unless you know what you are doing
Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

function Create-AdminReport() {
  param($orgs)

  $now = (Get-Date).ToString('dddd dd MMMM yyyy HH:mm:ss')
  
  foreach ($org in $orgs) {
    $orgsTable += '<tr style="height:12.75pt">
      <td colspan="2" nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;">
      <span style="font-size:10pt;">' + $org.Name + '</span>
      </td>
      <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;">
      <span style="font-size:10pt;">' + $org.IsBackedUp + '</span>
      </td>
    </tr>'
  }

  $EmailBody = '<html>
  <body style="font-family:Tahoma,sans-serif;">
  <table>
   <tr>
    <td>
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
     <tr>
      <td colspan="3" style="width:80%;background:#F5BD4C;padding:0cm 0cm 12.75pt 11.25pt;height:50pt">
      <b><span style="font-size:18pt;color:white">Organizations with missing contact person</span></b>
      </td>
     </tr>
     <tr style="height:26.25pt">
      <td colspan="3" style="width:80%;border:solid #A7A9AC 1pt;background:#F3F4F4;padding:3.75pt 0cm 0cm 11.25pt;height:26.25pt">
      <span style="color:#626365">Details (Generated on ' + $now + ')</span>
      </td>
     </tr>
     <tr style="height:12.75pt">
      <td nowrap colspan="2" style="width:25%;border:solid #A7A9AC 1pt;background:#E3E3E3">
      <b><span style="font-size:10pt;color:black;">Name</span></b>
      </td>
      <td nowrap style="width:25%;border:solid #A7A9AC 1pt;background:#E3E3E3">
      <b><span style="font-size:10pt;color:black;">Backed up</span></b>
      </td">
     </tr>
     ' + $orgsTable + ' 
    </table>
    </td>
   </tr>
  </table>
  </body>
  </html>'

  return $EmailBody
}


$companies = Import-CSV “$PSScriptRoot\organizations.csv” -Delimiter ";"
$organizations = Get-VBOOrganization
[System.Collections.ArrayList]$orgList = $organizations

foreach ($org in $organizations) {
  foreach ($comp in $companies) {
    if ($comp.Organization -eq $org.Name) {
      $orgList.Remove($org)
    }
  }
}

# Sending email to admin 
Write-Host "Sending email to admin"

$adminEmailReport = Create-AdminReport -orgs $orgList

if ($usessl) {
  Send-MailMessage -From $from -To $to -Subject $mailSubject -BodyAsHtml $adminEmailReport -SmtpServer $smtpserver -Port $port -Credential $cred -UseSsl -DeliveryNotificationOption OnFailure,OnSuccess
} else {
  Send-MailMessage -From $from -To $to -Subject $mailSubject -BodyAsHtml $adminEmailReport -SmtpServer $smtpserver -Port $port -Credential $cred -DeliveryNotificationOption OnFailure,OnSuccess
}