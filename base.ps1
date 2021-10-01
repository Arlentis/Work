$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
Import-PSSession $session -AllowClobber -Verbose


#region MAILBOX PERMISSIONS

### ADD SEND-AS PERMISSION
#Add-ADPermission -Identity 'IM' -User 'DanielM' -ExtendedRights 'Send-as'

### ADD FULL ACCESS PERMISSION
#$Users1 = "SarahJVP"
#ForEach ($User1 in $Users1) {Add-MailboxPermission -Identity "Managed" -User $User1 -AccessRights FullAccess -InheritanceType All}

### CHECK FULL ACCESS PERMISSION
#Get-MailboxPermission -Identity mayfairassist@healix.com | select-object "User" | Format-List

### GRANT SEND-ON-BEHALF TO DISTRIBUTION GROUP
#$Users1 = "Niamh","HR","DebbiePa","Tania","Trainingandqualityhi"
#Foreach ($User1 in $Users1) {Set-DistributionGroup "LMSSquad" -GrantSendOnBehalfTo $User1}

### GRANT SEND-ON-BEHALF OF MAILBOX
#Set-mailbox ‘managed’ –Grantsendonbehalfto @{add="Julie.Debenham@healix.com"}

### GRANT SEND-AS PERMISSION
#Add-ADPermission -Identity 'managed' -User 'JulieD' -ExtendedRights 'Send-as'

### CHECK WHICH MAILBOXES A USER HAS ACCESS TO
#Get-Mailbox -ErrorAction SilentlyContinue | Get-MailboxPermission -User Monika -ErrorAction SilentlyContinue | Select-Object -Property "AccessRights","User","Identity" | Format-List -ErrorAction SilentlyContinue

### REMOVE MAILBOX ACCESS
#$Username1 = "ConnieB"
#$Mailboxes1 = "MiabInsurance","Zurich"
#ForEach ($Mailbox1 in $Mailboxes1) {Remove-MailboxPermission -Identity $Mailbox1 -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{remove="$Username1"}} ; Remove-ADPermission -Identity $Mailbox1 -User $Username1 -ExtendedRights "Send As"


### GRANT COMPLETE MAILBOX ACCESS TO MULTIPLE PEOPLE
#$Mailbox = "HHSremittances"
#$Users1 = "Inmaculada","EmilyS","Lana"
#ForEach ($User1 in $Users1) {Add-MailboxPermission -Identity $Mailbox -User $User1 -AccessRights FullAccess -InheritanceType All  ; Set-mailbox $Mailbox –Grantsendonbehalfto @{Add=$User1} ; Add-ADPermission -Identity $Mailbox -User $User1 -ExtendedRights 'Send-as' }

### GRANT COMPLETE ACCESS TO MAILBOXES FOR ONE USER
#$Username1 = "ConnieB"
#$Mailboxes1 = "GlobalNetwork"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}


#endregion 


#region EMAIL ALIASES

### ADD EMAIL ALIAS
#Set-Mailbox "internationalhealthc" -EmailAddresses @{Add="cobhamaerospace@healix.com"}

### REMOVE EMAIL ALIAS
#Set-Mailbox "internationalhealthc" -EmailAddresses @{remove="AFL@healix.com"}

### FIND ALIAS
#Get-Recipient "westburtonb@healix.com" | Select-Object PrimarySmtpAddress

#endregion


#region CALENDAR PERMISSIONS (Remove entries before adding new ones)
### REMOVE CALENDAR PERMISSION
#Remove-MailboxFolderPermission -Identity "HHSClientTeamsAL@healix.com:\Calendar" -User "Charlotte"

### ADD CALENDAR PERMISSION
#Add-MailboxFolderPermission -Identity "HHSExec@healix.com:\Calendar" -User john.pugh@healix.com -AccessRights Editor

### GET CALENDAR PERMISSION
#Get-MailboxFolderPermission –Identity HHSExec@healix.com:\calendar | Select-Object -Property "User","AccessRights" | Format-Table

#endregion


#region GROUP MEMBERSHIP 
### GET ALL MEMBERS OF GROUP/S AND SUBGROUP/S AND EXPORT TO FILE
#$groups = "FCO security group"
#$results = foreach ($group in $groups) {
#    Get-ADGroupMember $group | select name,emailaddress,company #@{n='GroupName';e={$group}}, @{n='Description';e={(Get-ADGroup $group -Properties description).description}}
#}
#$results
#$results | Export-csv C:\Users\CLAdmin\Desktop\GroupMemberShip.txt -NoTypeInformation


### GET LIST OF USERS CREATION DATES
#$Users = "Name1", "Name2" 
#Foreach ($User in $Users) {
#Get-aduser $User -properties whencreated -ErrorAction SilentlyContinue| Select-Object SamAccountName,GivenName,Surname,Whencreated -ErrorAction SilentlyContinue | FT
#}


### GET GROUP MEMBERSHIP
#Get-ADGroupMember -Identity distallHHS | select-object -property "Name"


### GET AD USER CREATION DATE
# Get-aduser CallumL -properties whencreated | Select-Object SamAccountName,GivenName,Surname,Whencreated | FL


### GET DETAILED GROUP MEMBER INFORMATION
#$Group = Get-ADGroupMember "Domain Users"
#$Output = Foreach ($User in $group) {get-aduser $User -Properties Name, EmailAddress, Company | Where {$_.EmailAddress -ne $null} | Where { $_.Enabled -eq $True} | Select-Object Name, EmailAddress, Company}
#$Output | Export-csv C:\Users\CLAdmin\Desktop\Output.csv -NoTypeInformation


### LIST ALL GROUPS A MEMBER IS PART OF
#Get-ADPrincipalGroupMembership -Identity Susie | Format-Table -Property name

#endregion


#region USER/GROUP INFORMATION 
### LAST LOGON DATE
#Get-ADuser -Identity Facilities -Properties "LastLogonDate"

### GET EMAILS OF GROUP MEMBERS
#Get-ADGroup -filter {name -like 'distOfficeWindsorHouse'} | Get-ADGroupMember -Recursive | Get-ADUser -Properties Mail |select -ExpandProperty Mail

#endregion


#region MAILBOX INFORMATION

### GET MX500 RECORD FOR TRANSFERING OUTLOOK AUTOREPLY TO NEW MAILBOX
#Get-Mailbox CallumL | FL LegacyExchangeDN

### SEARCH MAILBOX FOR EMAL WITH SPECIFIC SUBJECT
#Search-Mailbox -Identity peter -SearchQuery ‘Subject:”Your Audley itinerary”‘ -TargetMailbox “CallumL” -TargetFolder “inbox” -LogOnly -LogLevel Full

### CHECK IF MAILBOX REPAIR IS NECESSARY
#New-MailboxRepairRequest -Mailbox HealixMedical -CorruptionType ProvisionedFolder,SearchFolder,AggregateCounts,Folderview -DetectOnly

#endregion



Remove-PSSession $session
