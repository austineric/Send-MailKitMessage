# Send-MailKitMessage

A replacement for PowerShell's [obsolete Send-MailMessage](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7.1#description) implementing the [Microsoft-recommended MailKit library](https://docs.microsoft.com/en-us/dotnet/api/system.net.mail.smtpclient?view=net-5.0#remarks).

# Installation  

**For current user only** (does not require elevated privileges): ```Install-Module -Name "Send-MailKitMessage" -Scope CurrentUser```  
 
**For all users** (requires elevated prvileges: ```Install-Module -Name "Send-MailKitMessage" -Scope AllUsers```  

# Usage

<details open>
 <summary><a href="https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_splatting?view=powershell-7.1">Splatting</a> example</summary>
  
```
using module Send-MailKitMessage

$Parameters=@{

    #Use secure connection if available (optional; [bool])
    "UseSecureConnectionIfAvailable=$true

    #Authentication (optional; [System.Management.Automation.PSCredential])
    "Credential"=[System.Management.Automation.PSCredential]::new("Username", (ConvertTo-SecureString -String "Password" -AsPlainText -Force))

    #SMTP server (required; [string])
    "SMTPServer"="SMTPServer"

    #Port (required; [int])
    "Port"=PortNumber

    #Sender (required; [MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm)
    "From"=[MimeKit.MailboxAddress]"SenderEmailAddress"

    #Recipient list (at least one recipient required; [MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "RecipientList"=[MimeKit.InternetAddressList]"RecipientEmailAddress" #single recipient
    "RecipientList"=($EmailList | ForEach-Object { [MimeKit.InternetAddressList]$_ }) #multiple recipients contained in a list named $EmailList

    #CC list (optional; [MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "CCList"=[MimeKit.InternetAddressList]"CCRecipientEmailAddress" #single CC recipient
    "CCList"=($EmailList | ForEach-Object { [MimeKit.InternetAddressList]$_ }) #multiple CC recipients contained in a list named $EmailList

    #BCC list (optional; [MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "BCCList"=[MimeKit.InternetAddressList]"BCCRecipientEmailAddress" #single BCC recipient
    "BCCList"=($EmailList | ForEach-Object { [MimeKit.InternetAddressList]$_ }) #multiple BCC recipients contained in a list named $EmailList
    
    #Subject (required; [string])
    "Subject"="Subject"
    
    #Text body (optional; [string])
    "TextBody"="TextBody"
    
    #HTML body (optional; [string])
    "HTMLBody"="HTMLBody"
    
    #Attachment list (optional; [System.Collections.Generic.List[string]] or array)
    "AttachmentList"="AttachmentPath" #single attachment
    "AttachmentList=($AttachmentFilePathList | ForEach-Object { $_ }) #multiple attachment filepaths contained in a list named $AttachmentFilePathList

}

Send-MailKitMessage @Parameters
```
  
</details>




# Releases
### 3.0
* Added credential support
* Added parameter to use secure connection if available
* Removed extended classes
* Changed ToList parameter to RecipientList
* Properly return exceptions from the module to the caller
* Switched from BouncyCastle to Portable.BouncyCastle
* Updated MailKit to 2.10.1
* Updated MimeKit to 2.10.1
