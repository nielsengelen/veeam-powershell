<# 
.NAME
    Veeam Backup for Microsoft Office 365 Auxiliary Backup Accounts creator
.SYNOPSIS
    Script to use for automatically creating auxiliary backup accounts
.DESCRIPTION
    Script to use for automatically creating auxiliary backup accounts for backing up SharePoint/OneDrive for Business
    Created for Veeam Backup for Microsoft Office 365 v5

    Requires AzureAD Module (will be installed if missing)

    The script will perform following steps:
    - Add accounts to your Office 365 subscription and a security group
    - Configure accounts as backup accounts within Veeam Backup for Microsoft Office 365

    Released under the MIT license.
.LINK
    http://www.github.com/nielsengelen
#>

# Modify the values below to your needs
# Number of accounts to add - advised is to add in bulk of 8 accounts
[Int]$Accounts = 8

# Number to start from (change this if you are adding extra accounts)
[Int]$StartFrom = 1

# Display Name for the accounts (these will get a number at the end, eg VeeamBackupAccount1, VeeamBackupAccount2)
$DisplayName = "VeeamBackupAccount"

# Your domain name
$Domain = "yourdomain(.onmicrosoft).com"

# Your security group name
$SecurityGroup = "VBO"

# Organization name as configured in Veeam Backup for Microsoft Office 365
$OrganizationName = "yourdomain(.onmicrosoft).com"

# Do not change below unless you know what you are doing
if ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -eq $null) {
  Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Installing AzureAD module..." -BackgroundColor DarkGreen
  Install-Module -Name AzureAD
}

Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"
Import-Module -Name AzureAD

# Connect to Office 365
Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Connecting to AzureAD..." -BackgroundColor DarkGreen
Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Provide your Office365 Admin Account credentials" -BackgroundColor DarkGreen
Connect-AzureAD

# Checking for the security group if it exists or if we need to create it
$secGroup = Get-AzureADGroup | Where { $_.DisplayName -eq $SecurityGroup }

if (!$secGroup) {
  Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Adding security group..." -BackgroundColor DarkGreen
  $secGroup = New-AzureADGroup -DisplayName $SecurityGroup -MailEnabled $False -MailNickName $SecurityGroup -SecurityEnabled $True -Description "Veeam Backup for Microsoft Office 365 Auxiliary Backup Accounts group" -ErrorAction SilentlyContinue
} else {
  Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Security group already exists..." -BackgroundColor DarkGray
}

Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Adding accounts..." -BackgroundColor DarkGreen

[Int]$TotalAccounts = $StartFrom + $Accounts - 1
$AccountsArray = @()

For ($i = $StartFrom; $i -le $TotalAccounts; $i++) {
  $PrincipalName = $DisplayName.ToLower() + $i + "@" + $Domain
  $Length = Get-Random -Minimum 8 -Maximum 16
  $NonAlphaChars = 3
  $Password = [System.Web.Security.Membership]::GeneratePassword($Length, $NonAlphaChars)

  $CheckUser = Get-AzureADUser -SearchString $PrincipalName -ErrorAction SilentlyContinue
  
  if ($checkUser) {
    Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Account $DisplayName$i already exists. Skipping..." -BackgroundColor DarkRed
   } else {
    Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Adding account: $DisplayName$i" -BackgroundColor DarkGreen

    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.EnforceChangePasswordPolicy = $False
    $PasswordProfile.ForceChangePasswordNextLogin = $False
    $PasswordProfile.Password = $Password

    $newUser = New-AzureADUser -AccountEnabled $True -DisplayName $DisplayName$i -MailNickName $DisplayName$i -PasswordPolicies "DisablePasswordExpiration" -PasswordProfile $PasswordProfile -UserPrincipalName $PrincipalName -ErrorAction SilentlyContinue
    $AccountsArray += ,@($PrincipalName, $Password)

    Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Adding account $DisplayName$i to security group $SecurityGroup." -BackgroundColor DarkGreen
    Add-AzureADGroupMember -ObjectId $SecGroup.ObjectId -RefObjectId $newUser.ObjectId
  }
}

Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Sleeping for 30 seconds to prevent sync issues..." -BackgroundColor DarkGray
Start-Sleep -Seconds 30

# Connect to Veeam Backup for Microsoft Office 365
Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Connecting to Veeam Backup for Microsoft Office 365..." -BackgroundColor DarkGreen

$Org = Get-VBOOrganization -Name $OrganizationName
$Group = Get-VBOOrganizationGroup -Organization $Org -DisplayName $SecurityGroup
$Members = Get-VBOOrganizationGroupMember -Group $Group
$BackupAccounts = @()

Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Configuring Backup Account for Veeam Backup for Microsoft Office 365..." -BackgroundColor DarkGreen

For ($j = 0; $j -lt $AccountsArray.Length; $j++) {
  ForEach ($Member in $Members) {
    if ($Member.Login -eq $AccountsArray[$j][0]) {
      Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Setting password for Backup Account $Member." -BackgroundColor DarkGreen
      $SecurePassword = ConvertTo-SecureString -String $AccountsArray[$j][1] -AsPlainText -Force
      $BackupAccounts += New-VBOBackupAccount -SecurityGroupMember $Member -Password $SecurePassword
    }
  }
}

Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Sleeping for 30 seconds to prevent sync issues..." -BackgroundColor DarkGray
Start-Sleep -Seconds 30

Write-Host (Get-Date -Format "hh:mm:ss dd/MM/yyyy") "Enabling Backup Accounts for Organization $OrganizationName..." -BackgroundColor DarkGreen
Set-VBOOrganization -Organization $Org -BackupAccounts $BackupAccounts | Out-Null

# Wipe the Accounts array
$AccountsArray = @()