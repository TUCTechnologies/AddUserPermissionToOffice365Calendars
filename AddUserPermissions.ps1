$Cred = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
 
Import-PSSession $Session
$userToAdd = ""
$PermissionType = "Editor"

$users = Get-Mailbox | Select -ExpandProperty PrimarySmtpAddress
Foreach ($u in $users) {
  $ExistingPermission = Get-MailboxFolderPermission -Identity $u":\calendar" -User $userToAdd -EA SilentlyContinue
  if ($ExistingPermission) {
    Remove-MailboxFolderPermission -Identity $u":\calendar" -User $userToAdd -Confirm:$False
  }
  if ($u -ne $userToAdd) {
    Add-MailboxFolderPermission $u":\Calendar" -user $userToAdd -accessrights $PermissionType
  }
}
 
Remove-PSSession $Session
