

Set-Location -Path ".\Dev"

#ensure this script is being run from the Dev directory
if ((Get-Item -Path (Get-Location).Path).Name -ne "Dev")
{
    Throw "This script needs to be run from the Dev directory"
}

#when making changes to the script remove the module prior to re-running the using statment
Remove-Module -Name "Send-MailKitMessage"

#IMPORTANT: the dll's used by Send-MailKitMessage do not get unloaded unless the session is restarted (the "Kill Terminal" button in Visual Studio Code), failure to restart the session can cause "assembly is already loaded" errors

Add-Type -Path "..\.\Libraries\BouncyCastle.Crypto.dll"
Add-Type -Path "..\.\Libraries\MailKit.dll"
Add-Type -Path "..\.\Libraries\MimeKit.dll"

#use parameters file created by Dev script
$ParametersFileLocation=".\Parameters.csv"

#get parameters (choose one row and skip a set number of rows to allow multiple parameter lines in the csv file)
$ParametersFile=(Import-Csv -Path $ParametersFileLocation | Select-Object -First 1 -Skip 2)

#message
$UseSecureConnectionIfAvailable=[string]::IsNullOrWhiteSpace($ParametersFile."UseSecureConnectionIfAvailable") ? $false : [bool]$ParametersFile."UseSecureConnectionIfAvailable"
$Credential=([string]::IsNullOrWhiteSpace($ParametersFile."Username") -or [string]::IsNullOrWhiteSpace($ParametersFile."Password")) ? $null : [System.Management.Automation.PSCredential]::new($ParametersFile."Username", (ConvertTo-SecureString -String $ParametersFile."Password" -AsPlainText -Force))
$Message=[MimeKit.MimeMessage]::new()
$Message.From.Add($ParametersFile."From")
$Message.To.Add($ParametersFile."To")
if (-not ([string]::IsNullOrWhiteSpace($ParametersFile."CC"))) { $CCList.Add($ParametersFile."CC") }
if (-not ([string]::IsNullOrWhiteSpace($ParametersFile."BCC"))) { $CCList.Add($ParametersFile."BCC") }
$Message.Subject=$ParametersFile."Subject"
$BodyBuilder=[MimeKit.BodyBuilder]::new()
$BodyBuilder.TextBody=$ParametersFile."TextBody"
$BodyBuilder.HtmlBody=[System.Web.HttpUtility]::HtmlDecode($ParametersFile."HTMLBody")
[System.Collections.Generic.List[string]]$AttachmentList=[System.Collections.Generic.List[string]]::new()
if (-not ([string]::IsNullOrWhiteSpace($ParametersFile."AttachmentLocation"))) { $AttachmentList.Add($ParametersFile."AttachmentLocation") }

#add bodybuilder to message body
$Message.Body=$BodyBuilder.ToMessageBody()

#smtp send
$Client=New-Object MailKit.Net.Smtp.SmtpClient
$Client.Connect($ParametersFile."SMTPServer", $ParametersFile."Port", ($UseSecureConnectionIfAvailable ? [MailKit.Security.SecureSocketOptions]::Auto : [MailKit.Security.SecureSocketOptions]::None))
if ($Credential)
{
    $Client.Authenticate($Credential.UserName, ($Credential.Password | ConvertFrom-SecureString -AsPlainText))
}
$Client.Send($Message)
