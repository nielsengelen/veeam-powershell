<# 
.NAME
    Veeam Backup for Microsoft Office 365 Data Mover
.SYNOPSIS
    Leverage this free tool to move user data from 1 repository to another repository.
.DESCRIPTION
    Leverage this free tool to move user data from 1 repository to another repository.

    Released under the MIT license.
.LINK
    http://www.github.com/nielsengelen
    http://www.github.com/veeamhub
#>

Add-Type -AssemblyName System.Windows.Forms 
[System.Windows.Forms.Application]::EnableVisualStyles()

Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"

#region begin GUI{ 
$VBO365                          = New-Object system.Windows.Forms.Form
$VBO365.ClientSize               = '300,275'
$VBO365.text                     = "VBO365 Data Mover"
$VBO365.TopMost                  = $false
$VBO365.FormBorderStyle          = 'Fixed3D'
$VBO365.MaximizeBox              = $false
$VBO365.MinimizeBox              = $false

$lblInfo                         = New-Object system.Windows.Forms.Label
$lblInfo.text                    = "Use this tool to migrate user data`nto another repository."
$lblInfo.AutoSize                = $true
$lblInfo.location                = New-Object System.Drawing.Point(15,15)
$lblInfo.Font                    = 'Microsoft Sans Serif,10'

$lblSourceRepo                   = New-Object system.Windows.Forms.Label
$lblSourceRepo.text              = "Source repository:"
$lblSourceRepo.AutoSize          = $true
$lblSourceRepo.location          = New-Object System.Drawing.Point(15,75)

$lblTargetRepo                   = New-Object system.Windows.Forms.Label
$lblTargetRepo.text              = "Target repository:"
$lblTargetRepo.AutoSize          = $true
$lblTargetRepo.location          = New-Object System.Drawing.Point(15,125)

$lblUser                         = New-Object system.Windows.Forms.Label
$lblUser.text                    = "Select user:"
$lblUser.AutoSize                = $true
$lblUser.location                = New-Object System.Drawing.Point(15,100)

$cmbSourceRepo                   = New-Object system.Windows.Forms.ComboBox
$cmbSourceRepo.width             = 150
$cmbSourceRepo.location          = New-Object System.Drawing.Point(135,75)

$cmbTargetRepo                   = New-Object system.Windows.Forms.ComboBox
$cmbTargetRepo.width             = 150
$cmbTargetRepo.location          = New-Object System.Drawing.Point(135,125)

$cmbUsers                        = New-Object system.Windows.Forms.ComboBox
$cmbUsers.width                  = 150
$cmbUsers.location               = New-Object System.Drawing.Point(135,100)

$btnSubmit                       = New-Object system.Windows.Forms.Button
$btnSubmit.text                  = "Migrate"
$btnSubmit.width                 = 80
$btnSubmit.height                = 30
$btnSubmit.location              = New-Object System.Drawing.Point(110,155)

$lblDisclaimer                   = New-Object system.Windows.Forms.Label
$lblDisclaimer.text              = "Copyright (c) 2018 VeeamHub`n`nDistributed under MIT license."
$lblDisclaimer.AutoSize          = $true
$lblDisclaimer.location          = New-Object System.Drawing.Point(35,200)
$lblDisclaimer.Font              = 'Microsoft Sans Serif,10'

$VBO365.Controls.AddRange(@($lblInfo,$lblSourceRepo,$lblTargetRepo,$lblUser,$cmbSourceRepo,$cmbTargetRepo,$cmbUsers,$btnSubmit, $lblDisclaimer))

#region gui events {
$reposList = Get-VBORepository

foreach ($repos in $reposList) {
 [void] $cmbSourceRepo.Items.Add($repos.Name)
 [void] $cmbTargetRepo.Items.Add($repos.Name)
}

$cmbSourceRepo.Add_SelectedIndexChanged({
  $cmbUsers.Items.Clear()
  $cmbSourceRepo.Text = ""
  $cmbTargetRepo.Text = ""
  $cmbUsers.Text = ""

  $repo = Get-VBORepository -Name $cmbSourceRepo.SelectedItem
  $usersList = Get-VBOEntityData -Type User -Repository $repo

  foreach ($users in $usersList) {
   [void] $cmbUsers.Items.Add($users.DisplayName)
  }
})

$btnSubmit.Add_Click({
  $sourceRepo = $cmbSourceRepo.SelectedItem
  $targetRepo = $cmbTargetRepo.SelectedItem

  if (!$sourceRepo) {
    [System.Windows.Forms.MessageBox]::Show("Please select a source repository.", "Error", 0, 48)
  } elseif (!$targetRepo) {
    [System.Windows.Forms.MessageBox]::Show("Please select a target repository.", "Error", 0, 48)
  } else {
    if ($sourceRepo -eq $targetRepo) {
      [System.Windows.Forms.MessageBox]::Show("Source and target repository are the same.", "Error" , 0, 48)
    } else {
      $user = $cmbUsers.SelectedItem

      if (!$user) {
        [System.Windows.Forms.MessageBox]::Show("No user selected.", "Error" , 0, 48)
      } else {
        $source = Get-VBORepository -Name $sourceRepo
        $target = Get-VBORepository -Name $targetRepo
        $userdata = Get-VBOEntityData -Type User -Repository $source -Name $user

        Move-VBOEntityData -From $source -To $target -User $userdata -Mailbox -ArchiveMailbox -OneDrive -Sites

        $cmbTargetRepo.Items.clear()
        $cmbUsers.Items.Clear()
        $cmbTargetRepo.Text = ""
        $cmbUsers.Text = ""

        [System.Windows.Forms.MessageBox]::Show("Data has been moved.", "Success", 0, 64)
      }
    }
  }
})
#endregion events }
#endregion GUI }


[void]$VBO365.ShowDialog()