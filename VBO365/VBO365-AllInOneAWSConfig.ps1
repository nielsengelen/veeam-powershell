Import-Module "C:\Program Files\Veeam\Backup365\Veeam.Archiver.PowerShell\Veeam.Archiver.PowerShell.psd1"

# Organization settings
$O365Username = "user@domain.com"
$O365Password = "password"
$Organization = "domain.com"
$AppID = "AAAAAAAA-XXXX-YYYY-ZZZZ-00000000000"
$AppSecret = "APPSECRET"

# Object storage settings
$S3AccessKey = "AMAZONACCESKEY"
$S3SecurityKey = "AMAZONSECURITYKEY"
$ObjectStorageName = "Object Storage Repository"
$BucketName = "BucketName"
$FolderName = "FolderName"

# Backup repository settings
$LocalFolder = "c:\localrepo"
$RepositoryName = "Backup Repository"
$RepositoryDesc = "Backup Repository"

# Combine it all together - do not touch below
$O365PasswordConverted = ConvertTo-SecureString $O365Password -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential ($O365Username, $O365PasswordConverted)

# Add the organization
$ApplicationSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
$Connection = New-VBOOffice365ConnectionSettings -AppCredential $Credentials -ApplicationId $AppID -ApplicationSecret $ApplicationSecret -GrantRolesAndPermissions
$Org = Add-VBOOrganization -Name $Organization -Office365ExchangeConnectionsSettings $Connection -Office365SharePointConnectionsSettings $Connection

# Add Object Storage
$SecurityKey = ConvertTo-SecureString -String $S3SecurityKey -AsPlainText -Force
$Account = Add-VBOAmazonS3Account -AccessKey $S3AccessKey -SecurityKey $SecurityKey
$AWSconn = New-VBOAmazonS3ServiceConnectionSettings -Account $Account -RegionType Global
$Bucket = Get-VBOAmazonS3Bucket -AmazonS3ConnectionSettings $AWSconn
$Folder = Get-VBOAmazonS3Folder -Bucket $Bucket -Name $FolderName
$ObjectStorage = Add-VBOAmazonS3ObjectStorageRepository -Folder $Folder -Name $ObjectStorageName

# Add Backup Repository
$Proxy = Get-VBOProxy
$Repository = Add-VBORepository -Proxy $Proxy -Path $LocalFolder -Name $RepositoryName -ObjectStorageRepository $ObjectStorage -RetentionPeriod Years3 -RetentionFrequencyType Daily -DailyTime "10:00" -DailyType Everyday -RetentionType ItemLevel -Description $RepositoryDesc

# Add the backup jobs
$MailJobItems = New-VBOBackupItem -Organization $Org -Mailbox -ArchiveMailbox
Add-VBOJob -Organization $Org -Repository $Repository -Name "E-mail" -SelectedItems $MailJobItems -Description "E-mail backup" -RunJob

$OneDriveItems = New-VBOBackupItem -Organization $Org -OneDrive
Add-VBOJob -Organization $Org -Repository $Repository -Name "OneDrive" -SelectedItems $OneDriveItems -Description "OneDrive backup" -RunJob

$SharePointItems = New-VBOBackupItem -Organization $Org -Sites
Add-VBOJob -Organization $Org -Repository $Repository -Name "SharePoint" -SelectedItems $SharePointItems -Description "SharePoint backup" -RunJob

Add-VBOJob -Organization $Org -Repository $Repository -Name "VIP Backup Job" -EntireOrganization -Description "VIP Users"

$oldRepo = Get-VBORepository -Name "Default Backup Repository"
Remove-VBORepository -Repository $oldRepo -Confirm:$false