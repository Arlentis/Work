$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
Import-PSSession $session -AllowClobber -Verbose

$Username1 =  Get-ADUser -Filter * -Properties * | Select-Object "Name","SamAccountName","Title","Mail" | Sort-Object "Name" | Out-Gridview -Title "User" -PassThru | Select-Object -ExpandProperty SamAccountName
$Mailbox1 =  Get-ADUser -Filter * -Properties * | Select-Object "Name","SamAccountName","Title","Mail" | Sort-Object "Name" | Out-Gridview -Title "Mailbox" -PassThru | Select-Object -ExpandProperty SamAccountName


ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}

Remove-PSSession $session
