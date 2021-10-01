##### ONLY CHANGE THE FOLLOWING 4 LINES + MAILBOX SECTION
$CopyFrom1 = ""
$FirstName1 = ""
$LastName1 = ""
$Username1 = ""
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
#Exit-PSSession

##### IF YOU WANT TO ADD MAILBOX PERMISSIONS THEN FILL IN THE IDENTITY FIELD BELOW, UNCOMMENT THE LINES AND RUN FROM HERE DOWN (HIGHLIGHT BELOW AND PRESS F8)

##### CONNECT TO EMC
#$password = Get-Content "C:\Users\CLAdmin\Documents\cred.txt" | ConvertTo-SecureString
#$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "HLX\CLAdmin",$Password
#$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://iris.hlx.int/powershell/ -Credential $Credential -Authentication Kerberos
#Import-PSSession $session -AllowClobber -Verbose

##### FILL IN IDENTITY FIELDS AS REQUIRED

#region FCO (Foreign Commonwealth Office)
#$Mailboxes1 = "Healthline","MFAT","CovidVaccineENZ","Covid19VaccinationEn","CovidVaccineEnquirie"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region HMS (Healix Medical Services)
#$Mailboxes1 = "HealthlineProviderIntelligence","HealixMedical","RepatExpenses","MedicalAssessments"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region GPN (Global Provider Network)
#$Mailboxes1 = "Research","GlobalProviderNetwor","Network.Development","GlobalNetwork"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region GSOC (Global Security Operations Center)
#$Mailboxes1 = "GSOC","IMT","GSOCTraining"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region Invoice Processing Team
#$Mailboxes1 = "HINTAdmin"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region HINT (Healix International)
#$Mailboxes1 = "InternationalAssista","InternationalHealthc"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region Projects
#$Mailboxes1 = "Implementation"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region T&Q (Training & Quality)
#$Mailboxes1 = "TrainingAndQualityHi"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region Marketing
#$Mailboxes1 = "Communications","Marketing","SalesAndEnquiries"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region Sales
#$Mailboxes1 = "SalesSupport"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region HHS (Healix Health Services, Marcus's team)
#$Mailboxes1 = "GeneralFDA","HHSBankDetails","HHSInvoices","HISPMI","Managed","MayfairAssist","MEC","Romif","PrioryGroup","techmahindra"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region HHS (Healix Health Services)
#$Mailboxes1 = "AprilUK","EduHealth","GeneralFDA","GroupHealth","HHSBankDetails","HHSInvoices","Lorium","Managed","MayfairAssist","MEC","MiabInsurance","Romif","Zurich","HISPMI","PrioryGroup","ShiftLeader","AprilMembership","CignPost","ClinicalQueries"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 –Grantsendonbehalfto @{add="$Username1"}}
#endregion

#region HHS 2 (Healix Health Services, David Butler's team)
#$Mailboxes1 = "Aetna","Chevron","CSHealthcare","GeneralFDA","Henner","HHSInvoices","Managed","MayfairAssist","MEC","MSHINT","Network.Development","RedArc","Techmahindra","William-Russell","Zurich"
#ForEach ($Mailbox1 in $Mailboxes1) {Add-MailboxPermission -Identity "$Mailbox1" -User $Username1 -AccessRights FullAccess -InheritanceType All ; Set-Mailbox $Mailbox1 -Grantsendonbehalfto @{add="$Username1"}}
#endregion

##### EXIT EMC SESSION
Remove-PSSession $session -ErrorAction SilentlyContinue
Exit-PSSession

