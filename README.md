# Send-MailKitMessage

A replacement for PowerShell's [obsolete Send-MailMessage](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.1#description) implementing the [Microsoft-recommended MailKit library](https://docs.microsoft.com/en-us/dotnet/api/system.net.mail.smtpclient?view=net-5.0#remarks).

# Installation  

**For current user only** (does not require elevated privileges): ```Install-Module -Name "Send-MailKitMessage" -Scope CurrentUser```  
 
**For all users** (requires elevated privileges): ```Install-Module -Name "Send-MailKitMessage" -Scope AllUsers```  

# Usage

```
using module Send-MailKitMessage

#use secure connection if available ([bool], optional)
$UseSecureConnectionIfAvailable = $true

#authentication ([System.Management.Automation.PSCredential], optional)
$Credential = [System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

#SMTP server ([string], required)
$SMTPServer = "SMTPServer"

#port ([int], required)
$Port = PortNumber

#priority ([string], optional)
$Priority = [string]"Priority"

#sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
$From = [MimeKit.MailboxAddress]"SenderEmailAddress"

#recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
$RecipientList = [MimeKit.InternetAddressList]::new()
$RecipientList.Add([MimeKit.InternetAddress]"Recipient1EmailAddress")

#cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$CCList = [MimeKit.InternetAddressList]::new()
$CCList.Add([MimeKit.InternetAddress]"CCRecipient1EmailAddress")

#bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
$BCCList = [MimeKit.InternetAddressList]::new()
$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")

#subject ([string], required)
$Subject = [string]"Subject"

#text body ([string], optional)
$TextBody = [string]"TextBody"

#HTML body ([string], optional)
$HTMLBody = [string]"HTMLBody"

#attachment list ([System.Collections.Generic.List[string]], optional)
$AttachmentList = [System.Collections.Generic.List[string]]::new()
$AttachmentList.Add("Attachment1FilePath")

#splat parameters
$Parameters = @{
    "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable    
    "Credential" = $Credential
    "SMTPServer" = $SMTPServer
    "Port" = $Port
    "Priority" = $Priority
    "From" = $From
    "RecipientList" = $RecipientList
    "CCList" = $CCList
    "BCCList" = $BCCList
    "Subject" = $Subject
    "TextBody" = $TextBody
    "HTMLBody" = $HTMLBody
    "AttachmentList" = $AttachmentList
}

#send message
Send-MailKitMessage @Parameters
```

# Releases
### 3.2.0-preview1
* Add support for Windows PowerShell

### 3.1.0
* Changed UseSecureConnectionIfAvailable parameter type from [bool] to [switch]

### 3.0.0
* Added credential support
* Added parameter to use secure connection if available
* Removed extended classes
* Changed ToList parameter to RecipientList
* Properly return exceptions from the module to the caller
* Switched from BouncyCastle to Portable.BouncyCastle
* Updated MailKit to 2.10.1
* Updated MimeKit to 2.10.1
