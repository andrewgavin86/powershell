===========================================================================
 The variables you will want to change are the SMTP Host, From Email and Expireindays.
 SMTPHOST: The smtp host it will use to send mail in this case we are using Office 365
 FromEmail: Who the script will send the e-mail from (you will need to utilize a office 365 license for this, make sure and set the password to never expire!)
 ExpireInDays: How many days before you want users alerted of passwords expiring

 Please note, in order for this to work, you must first ensure every user has an email address associated with their AD account, add emails under user account/email and proxyaddress SMTP: email@email.com
 
 Run the script manually first as it will ask for credentials to send email and then safely store them for future use.
 ===========================================================================


#SMTP Host
$SMTPHost = "smtp.office365.com"
#Who is the e-mail from
$FromEmail = "email-address@hostname.com"
#Password expiry days
$expireindays = 7
 
#Program File Path
$DirPath = "C:\Automation\PasswordExpiry"
 
$Date = Get-Date
#Check if program dir is present
$DirPathCheck = Test-Path -Path $DirPath
If (!($DirPathCheck))
{
 Try
 {
 #If not present then create the dir
 New-Item -ItemType Directory $DirPath -Force
 }
 Catch
 {
 $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
 }
}
#CredObj path
$CredObj = ($DirPath + "\" + "EmailExpiry.cred")
#Check if CredObj is present
$CredObjCheck = Test-Path -Path $CredObj
If (!($CredObjCheck))
{
 "$Date - INFO: creating cred object" | Out-File ($DirPath + "\" + "Log.txt") -Append
 #If not present get office 365 cred to save and store
 $Credential = Get-Credential -Message "Please enter your Office 365 credential that you will use to send e-mail from $FromEmail. If you are not using the account $FromEmail make sure this account has 'Send As' rights on $FromEmail."
 #Export cred obj
 $Credential | Export-CliXml -Path $CredObj
}
 
Write-Host "Importing Cred object..." -ForegroundColor Yellow
$Cred = (Import-CliXml -Path $CredObj)
 
 
# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
"$Date - INFO: Importing AD Module" | Out-File ($DirPath + "\" + "Log.txt") -Append
Import-Module ActiveDirectory
"$Date - INFO: Getting users" | Out-File ($DirPath + "\" + "Log.txt") -Append
$users = Get-Aduser -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress -filter { (Enabled -eq 'True') -and (PasswordNeverExpires -eq 'False') } | Where-Object { $_.PasswordExpired -eq $False }
 
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
 
# Process Each User for Password Expiry
foreach ($user in $users)
{
 $Name = (Get-ADUser $user | ForEach-Object { $_.Name })
 Write-Host "Working on $Name..." -ForegroundColor White
 Write-Host "Getting e-mail address for $Name..." -ForegroundColor Yellow
 $emailaddress = $user.emailaddress
 If (!($emailaddress))
 {
 Write-Host "$Name has no E-Mail address listed, looking at their proxyaddresses attribute..." -ForegroundColor Red
 Try
 {
 $emailaddress = (Get-ADUser $user -Properties proxyaddresses | Select-Object -ExpandProperty proxyaddresses | Where-Object { $_ -cmatch '^SMTP' }).Trim("SMTP:")
 }
 Catch
 {
 $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
 }
 If (!($emailaddress))
 {
 Write-Host "$Name has no email addresses to send an e-mail to!" -ForegroundColor Red
 #Don't continue on as we can't email $Null, but if there is an e-mail found it will email that address
 "$Date - WARNING: No email found for $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
 }
 
 }
 #Get Password last set date
 $passwordSetDate = (Get-ADUser $user -properties * | ForEach-Object { $_.PasswordLastSet })
 #Check for Fine Grained Passwords
 $PasswordPol = (Get-ADUserResultantPasswordPolicy $user)
 if (($PasswordPol) -ne $null)
 {
 $maxPasswordAge = ($PasswordPol).MaxPasswordAge
 }
 
 $expireson = $passwordsetdate + $maxPasswordAge
 $today = (get-date)
 #Gets the count on how many days until the password expires and stores it in the $daystoexpire var
 $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
 
 If (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays))
 {
 "$Date - INFO: Sending expiry notice email to $Name" | Out-File ($DirPath + "\" + "Log.txt") -Append
 Write-Host "Sending Password expiry email to $name" -ForegroundColor Yellow
 
 $SmtpClient = new-object system.net.mail.smtpClient
 $MailMessage = New-Object system.net.mail.mailmessage
 
 #Who is the e-mail sent from
 $mailmessage.From = $FromEmail
 #SMTP server to send email
 $SmtpClient.Host = $SMTPHost
 #SMTP SSL
 $SMTPClient.EnableSsl = $true
 #SMTP credentials
 $SMTPClient.Credentials = $cred
 #Send e-mail to the users email
 $mailmessage.To.add("$emailaddress")
 #Email subject
 $mailmessage.Subject = "Your password will expire $daystoexpire days"
 #Notification email on delivery / failure
 $MailMessage.DeliveryNotificationOptions = ("onSuccess", "onFailure")
 #Send e-mail with high priority
 $MailMessage.Priority = "High"
 $mailmessage.Body =
 "Dear $Name,
Your Domain password will expire in $daystoexpire days. Please change it as soon as possible.
 
To change your password, follow the method below:
 
1. On your Windows computer
 a.	If you are not in the office, logon and connect to VPN. 
 b.	Log onto your computer as usual and make sure you are connected to the internet.
 c.	Press Ctrl-Alt-Del and click on ""Change Password"".
 d.	Fill in your old password and set a new password.  See the password requirements below.
 e.	Press OK to return to your desktop. 
 
The new password must meet the minimum requirements set forth in our corporate policies including:
 1.	It must be at least 8 characters long.
 2.	It must contain at least one character from 3 of the 4 following groups of characters:
 a.  Uppercase letters (A-Z)
 b.  Lowercase letters (a-z)
 c.  Numbers (0-9)
 d.  Symbols (!@#$%^&*...)
 3.	It cannot match any of your past 24 passwords.
 4.	It cannot contain characters which match 3 or more consecutive characters of your username.
 5.	You cannot change your password more often than once in a 24 hour period.
 
If you have any questions please contact our Support team at helpdesk@frontierit.com or call us at 719.888.4357
 
Thanks,
YOUR NAME HERE
"
 Write-Host "Sending E-mail to $emailaddress..." -ForegroundColor Green
 Try
 {
 $smtpclient.Send($mailmessage)
 }
 Catch
 {
 $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
 }
 }
 Else
 {
 "$Date - INFO: Password for $Name not expiring for $daystoexpire days" | Out-File ($DirPath + "\" + "Log.txt") -Append
 Write-Host "Password for $Name does not expire for $daystoexpire days" -ForegroundColor White
 }
}
