
#ensure this script is being run from the Dev directory
Set-Location -Path ".\Dev"
if ((Get-Item -Path (Get-Location).Path).Name -ne "Dev")
{
    Throw "This script needs to be run from the Dev directory"
}

#when making changes to the script remove the module prior to re-running the using statment
Remove-Module -Name "Send-MailKitMessage"

#IMPORTANT: the dll's used by Send-MailKitMessage do not get unloaded unless the session is restarted (the "Kill Terminal" button in Visual Studio Code), failure to restart the session can cause "assembly is already loaded" errors
Invoke-Expression "using module `"..\..\Send-MailKitMessage`""

#ensure the repo version of the module is being used (not the real version if it is installed as a module)
Get-Module -Name "Send-MailKitMessage" | Select-Object -Property "ModuleBase"

#set parameter file location
$ParametersFileLocation=".\Parameters.csv"

#create parameter file
if (-not (Test-Path -Path $ParametersFileLocation))
{
    [PSCustomObject]@{
        UseSecureConnectionIfAvailable=[string]::Empty
        Username=[string]::Empty
        Password=[string]::Empty
        SMTPServer=[string]::Empty
        Port=[string]::Empty
        From=[string]::Empty
        To=[string]::Empty
        CC=[string]::Empty
        BCC=[string]::Empty
        Subject=[string]::Empty
        TextBody=[string]::Empty
        HTMLBody=[string]::Empty
        AttachmentLocation=[string]::Empty
    } | Export-Csv -Path ".\Parameters.csv"
}

#get parameters (choose one row and skip a set number of rows to allow multiple parameter lines in the csv file)
$ParametersFile=(Import-Csv -Path $ParametersFileLocation | Select-Object -First 1 -Skip 0)

#the below code is (mostly) from the README, using it here helps ensure the documentation is valid
#any material changes required here should be made in the README as well

$Parameters=@{

    #Use secure connection if available (optional)
    "UseSecureConnectionIfAvailable"=[string]::IsNullOrWhiteSpace($ParametersFile."UseSecureConnectionIfAvailable") ? $false : [bool]$ParametersFile."UseSecureConnectionIfAvailable"

    #authentication (optional)
    "Credential"=([string]::IsNullOrWhiteSpace($ParametersFile."Username") -or [string]::IsNullOrWhiteSpace($ParametersFile."Password")) ? $null : [System.Management.Automation.PSCredential]::new($ParametersFile."Username", (ConvertTo-SecureString -String $ParametersFile."Password" -AsPlainText -Force))

    #SMTP server (required)
    "SMTPServer"=$ParametersFile."SMTPServer"

    #Port (required)
    "Port"=$ParametersFile."Port"

    #Sender (required) (http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm)
    "From"=[MimeKit.MailboxAddress]$ParametersFile."From"

    #Recipient list (at least one recipient required) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "RecipientList"=[MimeKit.InternetAddressList]$ParametersFile."To"  #I actually think this works

    #CC list (optional) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "CCList"=[MimeKit.InternetAddressList]([string]::IsNullOrWhiteSpace($ParametersFile."CC") ? $null : $ParametersFile."CC")

    #BCC list (optional) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "BCCList"=[MimeKit.InternetAddressList]([string]::IsNullOrWhiteSpace($ParametersFile."BCC") ? $null : $ParametersFile."BCC")
    
    #Subject (required)
    "Subject"=$ParametersFile."Subject"
    
    #Text body (optional)
    "TextBody"=$ParametersFile."TextBody"
    
    #HTML body (optional)
    "HTMLBody"=$ParametersFile."HTMLBody"
    
    #Attachment list (optional)
    "AttachmentList"=[System.Collections.Generic.List[string]]([string]::IsNullOrWhiteSpace($ParametersFile."AttachmentLocation") ? $null : $ParametersFile."AttachmentLocation")

}

Send-MailKitMessage @Parameters