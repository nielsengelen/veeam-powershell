<# 
.NAME
    Veeam Backup for Microsoft Office 365 Mail per job 
.SYNOPSIS
    Script to use for reporting the latest session per job to a specific email
.DESCRIPTION
    Script to use for reporting the latest session per job to a specific email

    Utilises CSV files as a databases
    - organizations.csv (contains contact information)
    - jobs-result.csv (contains latest status per job)

    Released under the MIT license.
.LINK
    http://www.github.com/nielsengelen
#>

$adminEmail = "niels@foonet.be" # Configure if you set $adminReport to $True

# Modify the values below to your needs
# Mail server configuration
$from = "vbo365@company.com"
$smtpserver = "mail.company.com"
$subject = "[VB365] Report for" # This will be added before the actual report title
$port = "587" # default: 25
$usessl = $True # Use SSL ($True) or not ($False)

# Authentication against the mail server
$username = "authentication@company.com"
$password = ConvertTo-SecureString "AUTHPASSWORD" -AsPlainText -Force

# Set to $True if you want to show output - should only be used for debugging the script
$debug = $False

# Do not change below unless you know what you are doing
Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
$cred = New-Object System.Management.Automation.PSCredential ($username, $password)

function Create-Report() {
  param($name, $stats)
   
  $date = Get-Date -Date $stats.EndTime -Format “dddd dd MMMM yyyy HH:mm”
  $duration = New-TimeSpan -Start $stats.CreationTime -End $stats.EndTime
  $log = $stats.Log
  $jobLog = @()

  if ($stats.Status.ToString().ToLower() -eq 'success') {
    $color = '#00B050'
  } elseif ($stats.Status.ToString().ToLower() -eq 'warning') {
    $color = '#F5BD4C'
  } else {
    $color = '#FB9895'
  }
  
  foreach ($line in $log) {
    $joblog += '<tr style="height:12.75pt">
      <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;">
      <span style="font-size:10pt;">' + $line.CreationTime + '</span>
      </td>
      <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;">
      <span style="font-size:10pt;">' + $line.EndTime + '</span>
      </td>
      <td colspan="2" nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt;">
      <span style="font-size:10pt;">' + $line.Title + '</span>
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
         <td colspan="4" style="width:80%;background:' + $color + ';padding:0cm 0cm 12.75pt 11.25pt;height:50pt">
         <b><span style="font-size:18pt;color:white">Backup job: ' + $name + ' (' + $stats.Status + ')</span></b>
         </td>
        </tr>
        <tr style="height:26.25pt">
         <td colspan="4" style="width:80%;border:solid #A7A9AC 1pt;background:#F3F4F4;padding:3.75pt 0cm 0cm 11.25pt;height:26.25pt">
         <span style="color:#626365">' + $date + '</span>
         </td>
        </tr>
        <tr>
         <td nowrap style="width:25%;border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <b><span style="font-size:10pt;">Start time</span></b>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt; height:12.75pt">
         <span style="font-size:10pt;">' + $stats.CreationTime + '</span>
         </td>
         <td nowrap style="width:25%;border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <b><span style="font-size:10pt;">Objects processed</span></b>
         </td>
         <td nowrap style="width:25%;border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <span style="font-size:10pt;">' + $stats.Statistics.ProcessedObjects + '</span>
         </td>
        </tr>
        <tr>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <b><span style="font-size:10pt;">End time</span></b>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <span style="font-size:10pt;">' + $stats.EndTime + '</span>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <p><b><span style="font-size:10pt;">Processing rate</span></b>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <span style="font-size:10pt;">' + $stats.Statistics.ProcessingRate + '</span>
         </td>
        </tr>
	    <tr style="height:12.75pt">
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <b><span style="font-size:10pt;">Duration</span></b>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <span style="font-size:10pt;">' + $duration + '</span>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <b><span style="font-size:10pt;">Data transferred</span></b>
         </td>
         <td nowrap style="border:solid #A7A9AC 1pt;padding:1.5pt 2.25pt 1.5pt 2.25pt">
         <span style="font-size:10pt;">' + $stats.Statistics.TransferredData + '</span></p>
         </td>
        </tr>
        <tr style="height:26.25pt">
         <td nowrap colspan="4" style="width:80%;border:solid #A7A9AC 1pt;background:#F3F4F4;padding:3.75pt 0cm 0cm 11.25pt;height:26.25pt">
         <span style="color:#626365">Details</span>
         </td>
        </tr>
        <tr style="height:17.25pt">
         <td nowrap style="width:25%;border:solid #A7A9AC 1pt;background:#E3E3E3">
         <b><span style="font-size:10pt;color:black;">Start Time</span></b>
         </td>
         <td nowrap style="width:25%;border:solid #A7A9AC 1pt;background:#E3E3E3">
         <b><span style="font-size:10pt;color:black;">End Time</span></b>
         </td">
         <td colspan="2" nowrap style="width:50%;border:solid #A7A9AC 1pt;background:#E3E3E3">
         <b><span style="font-size:10pt;color:black;">Details</span></b>
         </td>
        </tr>
        ' + $jobLog + ' 
       </table>
       </td>
      </tr>
   </table>
  </body>
  </html>'

  return $EmailBody
}

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
$importedJobs = Import-CSV “$PSScriptRoot\jobs-result.csv” -Delimiter "," | Sort-Object Organization,Name -Unique
$exportedJobs = @()

foreach ($comp in $companies) {
    $org = Get-VBOOrganization -Name $comp.Organization
    $currentJobs = Get-VBOJob -Organization $org
    $now = (Get-Date).ToString('dd/MM/yyyy HH:mm:ss')

    foreach ($currentJob in $currentJobs) {
        if ($debug) {
          Write-Host "Checking job" $currentJob.Name "("$currentJob.Organization.Name")"
        }

        $session = Get-VBOJobSession -Job $currentJob -Last
        
        if ($session.Status -ne $null -and $session.Status.ToString().ToLower() -ne 'running') {
            # Check if we have the job listed in the jobs database file
            foreach ($checkJob in $importedJobs) {
                if ($debug) {
                  Write-Host "Checking existing job (" $currentJob ") against imported job (" $checkJob ")"
                }

                if ($currentJob.Name -eq $checkJob.Name -and $checkJob.Organization -eq $currentJob.Organization.Name) {
                    # Already included in the file - check if we need to update the time
                    if ((Get-Date $checkJob.LastEmail.ToString()) -lt (Get-Date $currentJob.LastRun.ToString())) {
                        if ($debug) {
                          Write-Host "Updating last email and sending email to"$comp.Contact for $checkJob.Name "("$checkJob.Organization")"
                        }

                        $report = Create-Report -name $currentJob.Name -stats $session
                        $mailSubject = $subject + " " + $currentJob.Name + " (" + $currentJob.Organization.Name + ")"
                        
                        if ($usessl) {
                            Send-MailMessage -From $from -To $comp.Contact -Subject $mailSubject -BodyAsHtml $report -SmtpServer $smtpserver -Port $port -Credential $cred -UseSsl -DeliveryNotificationOption OnFailure,OnSuccess
                        } else {
                            Send-MailMessage -From $from -To $comp.Contact -Subject $mailSubject -BodyAsHtml $report -SmtpServer $smtpserver -Port $port -Credential $cred -DeliveryNotificationOption OnFailure,OnSuccess
                        }
                    }
                }
            }

            $exportJob = $currentJob | Select-Object Organization, Name, @{Name="LastEmail";Expression={$_.LastRun}}
            $exportedJobs += $exportJob
        }
    }
}

# Updating jobs file
if ($debug) {
  Write-Host "Update jobs file"
}

$exportedJobs | Export-Csv -Path "$PSScriptRoot\jobs-result.csv" -NoTypeInformation -Force