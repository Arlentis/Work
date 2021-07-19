##### CHANGES TO BE MADE
# Automate the copying of mailboxes? -- not sure whether this is the best idea really.. 

##### ONLY CHANGE THE FOLLOWING 4 LINES
$CopyFrom1 = "HameedT"
$FirstName1 = "Callum"
$LastName1 = "Deleteme"
$Username1 = "CallumD"
$UserPassword = ConvertTo-SecureString -AsPlainText "H34l1x!12" -Force

##### GETS TEMPLATE INFORMATION FROM EXISTING USER - OU, SELECTED PROPERTIES AND DATE FOR DESCRIPTION
$ou1 = (((Get-ADUser -identity $CopyFrom1 -Properties CanonicalName | select-object -expandproperty DistinguishedName) -split",") | select -Skip 1) -join ','
$DateTime = Get-Date -Format "dd/MM/yyyy HH:mm"
$template_account = Get-ADUser -Identity $CopyFrom1 -Properties State,Department,Country,City,wWWHomePage,Title,HomePage,OfficePhone,StreetAddress,MemberOf,Organization,Manager,HomePhone,Fax,City,Company,ScriptPath
$template_account.UserPrincipalName = $null

##### CREATE NEW USER USING TEMPLATE FROM ABOVE PLUS ADDITIONAL FIELDS LISTED BELOW
New-ADUser `
    -Instance $template_account `
    -Name "$FirstName1 $LastName1" `
    -SamAccountName "$Username1" `
    -AccountPassword $UserPassword `
    -Enabled $True `
    -Description "Created on $DateTime" `
    -DisplayName "$FirstName1 $LastName1" `
    -UserPrincipalName "$Username1@healix.com" `
    -GivenName "$FirstName1" `
    -Surname "$LastName1"

##### COPY GROUP MEMBERSHIP FROM TEMPLATE USER TO NEW USER
Start-Sleep 3
Get-ADUser -Identity $CopyFrom1 -Properties memberof | Select-Object -ExpandProperty memberof |  Add-ADGroupMember -Members $Username1

##### MOVE NEW USER INTO SAME OU AS TEMPLATE USER
Start-Sleep 3
Move-ADObject -Identity "CN=$FirstName1 $LastName1,CN=Users,DC=hlx,DC=int" -TargetPath $ou1
Remove-ADGroupMember -Identity "FCO Viewers" -Member $Username1 -ErrorAction SilentlyContinue -Confirm:$false
Remove-ADGroupMember -Identity "FCO Users" -Member $Username1 -ErrorAction SilentlyContinue -Confirm:$false

##### CONNECT TO EMC
$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
Import-PSSession $session -AllowClobber -Verbose

##### CREATE NEW MAILBOX FOR NEW USER -- DEFAULTS TO FIRSTNAME.LASTNAME@HEALIX.COM
Enable-Mailbox -identity $Username1 -Alias $Username1 -Database 'Archive Database'

##### LIST DELEGATE MAILBOX ACCESS OF TEMPLATE USER TO ADD TO NEW USER MANUALLY (THIS SHOULD BE CONTROLLED BY GROUPS IDEALLY...)
Get-Mailbox -ResultSize Unlimited -ErrorAction SilentlyContinue | Get-MailboxPermission -User $CopyFrom1 -ErrorAction SilentlyContinue | ft Identity -AutoSize -Wrap

##### EXIT EMC SESSION
Exit-PSSession

##### IF YOU WANT TO ADD MAILBOX PERMISSIONS THEN FILL IN THE IDENTITY FIELD BELOW, UNCOMMENT THE LINES AND RUN FROM HERE DOWN (HIGHLIGHT BELOW AND PRESS F8)

##### CONNECT TO EMC
$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
Import-PSSession $session -AllowClobber -Verbose


##### FILL IN IDENTITY FIELDS AS REQUIRED

## --------------------------------------- FCO ---------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity Healthline -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity MFAT -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity CovidVaccineENZ -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Covid19VaccinationEn -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity CovidVaccineEnquirie -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- GPN ---------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity HealthlineProviderIntelligence -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity HealixMedical -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity RepatExpenses -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- GSOC ---------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity GSOC -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity IMT -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- IPT ----------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity HINTAdmin -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- HINT ---------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity InternationalAssista -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity InternationalHealthc -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity MedicalAssessments -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- PROJECTS ------------------------------------------------------------------------------ ##
#Add-MailboxPermission -Identity Implementation -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- TRAINING & QUALITY -------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity TrainingAndQualityHi -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- MARKETING ----------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity Communications -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Marketing -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity SalesAndEnquiries -User $Username1 -AccessRights FullAccess -InheritanceType All
## --------------------------------------- HHS ----------------------------------------------------------------------------------- ##
#Add-MailboxPermission -Identity AprilUK -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity EduHealth -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity GeneralFDA -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity GroupHealth -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity HHSBankDetails -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity HHSInvoices -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Lorium -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Managed -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity MayfairAssist -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity MEC -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity miabinsurance -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Romif -User $Username1 -AccessRights FullAccess -InheritanceType All
#Add-MailboxPermission -Identity Zurich -User $Username1 -AccessRights FullAccess -InheritanceType All

##### EXIT EMC SESSION
Exit-PSSession

