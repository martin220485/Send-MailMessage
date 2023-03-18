# Import the EWS module
Import-Module -Name EWSPowerShell

# Create an EWS session with your Exchange credentials
$Session = New-EWSSession -Credential (Get-Credential) -ExchangeVersion 2013

# Create an EmailMessage object with the subject, body and recipient
$EmailMessage = New-EWSEmailMessage -Subject "Test mail" -Body "This is a test mail with a reminder" -To "user02@fabrikam.com"

# Set the PidLidReminderFileParameter property on the EmailMessage object to specify the sound file name
Set-EWSExtendedProperty -Item $EmailMessage -PropertyType String -PropertyName "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/851F001F" -PropertyValue "Reminder.wav"

# Set other reminder properties on the EmailMessage object as desired, such as the reminder date and time, whether it is active or not, etc.
Set-EWSExtendedProperty -Item $EmailMessage -PropertyType SystemTime -PropertyName "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85020040" -PropertyValue (Get-Date).AddMinutes(10) # Reminder due in 10 minutes
Set-EWSExtendedProperty -Item $EmailMessage -PropertyType Boolean -PropertyName "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8503000B" -PropertyValue $true # Reminder is active

# Send the EmailMessage object
Send-EWSEmailMessage -Item $EmailMessage
