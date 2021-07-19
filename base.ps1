$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
Import-PSSession $session -AllowClobber -Verbose


#################################################################### MAILBOX PERMISSIONS ####################################################################
### ADD SEND-AS PERMISSION
#Add-ADPermission -Identity ______ -User 'HLX\Username' -ExtendedRights 'Send-as'

### ADD FULL ACCESS PERMISSION
#$Users1 = "AnnaG","Keira","DeborahGr","AnnI","LouiseG","GeorginaM"
#ForEach ($User1 in $Users1) {Add-MailboxPermission -Identity "covidtesting" -User $User1 -AccessRights FullAccess -InheritanceType All}

### CHECK FULL ACCESS PERMISSION
#Get-MailboxPermission -Identity CovidVaccineEnquiries-ENZ@healix.com | select-object "User" | Format-List

### GRANT SEND-ON-BEHALF TO DISTRIBUTION GROUP
#$Users1 = "Niamh","HR","DebbiePa","Tania","Trainingandqualityhi"
#Foreach ($User1 in $Users1) {Set-DistributionGroup "LMSSquad" -GrantSendOnBehalfTo $User1}




#################################################################### EMAIL ALIASES ##########################################################################
### ADD EMAIL ALIAS
#Set-Mailbox "managed" -EmailAddresses @{Add="clearbank@healix.com"}

### REMOVE EMAIL ALIAS
#Set-Mailbox "healthline" -EmailAddresses @{remove="ONS_Hub@healix.com"}

### FIND ALIAS
#Get-Recipient "Enquiries@healix.com" | Select-Object PrimarySmtpAddress




#################################################################### CALENDAR PERMISSIONS (Remove entries before adding new ones) ###########################
### REMOVE CALENDAR PERMISSION
#Remove-MailboxFolderPermission -Identity "Callum.Limb@healix.com:\Calendar" -User "Jayson.Glover@healix.com"

### ADD CALENDAR PERMISSION
#Add-MailboxFolderPermission -Identity "Callum.Limb@healix.com:\Calendar" -User "Jayson.Glover@healix.com"  -AccessRights Editor




#################################################################### GROUP MEMBERSHIP ########################################################################
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
#Get-ADGroupMember -Identity ConfGPNViewers | select-object -property "Name"

### GET AD USER CREATION DATE
# Get-aduser CallumL -properties whencreated | Select-Object SamAccountName,GivenName,Surname,Whencreated | FL


### GET DETAILED GROUP MEMBER INFORMATION
#$Group = Get-ADGroupMember "Domain Users"
#$Output = Foreach ($User in $group) {get-aduser $User -Properties Name, EmailAddress, Company | Where {$_.EmailAddress -ne $null} | Where { $_.Enabled -eq $True} | Select-Object Name, EmailAddress, Company}
#$Output | Export-csv C:\Users\CLAdmin\Desktop\Output.csv -NoTypeInformation



####################################################################  USER/GROUP INFORMATION  #################################################################
### LAST LOGON DATE
#Get-ADuser -Identity Facilities -Properties "LastLogonDate"

### GET EMAILS OF GROUP MEMBERS
#Get-ADGroup -filter {name -like 'distOfficeWindsorHouse'} | Get-ADGroupMember -Recursive | Get-ADUser -Properties Mail |select -ExpandProperty Mail




#################################################################### MAILBOX INFORMATION ######################################################################
### GET MX500 RECORD FOR TRANSFERING OUTLOOK AUTOREPLY TO NEW MAILBOX
#Get-Mailbox CallumL | FL LegacyExchangeDN

### SEARCH MAILBOX FOR EMAL WITH SPECIFIC SUBJECT
#Search-Mailbox -Identity peter -SearchQuery ‘Subject:”Your Audley itinerary”‘ -TargetMailbox “CallumL” -TargetFolder “inbox” -LogOnly -LogLevel Full

### CHECK IF MAILBOX REPAIR IS NECESSARY
#New-MailboxRepairRequest -Mailbox HealixMedical -CorruptionType ProvisionedFolder,SearchFolder,AggregateCounts,Folderview -DetectOnly



Remove-PSSession $session
