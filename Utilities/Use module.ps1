
#ensure this script is being run from the Dev directory
if ((Get-Item -Path (Get-Location).Path).Name -ne "Utilities")
{
    Throw "This script needs to be run from the Utilities directory"
}

#when making changes to the script remove the module prior to re-running the using statment (annoyingly the if statement has to be on one long line)
$PowerShellEdition = $PSVersionTable["PSEdition"]
if (($PowerShellEdition -eq "Desktop" -and (Get-Module -ListAvailable | Where-Object { $_."Name" -eq "Send-MailKitMessage" })) -or ($PowerShellEdition -eq "Core" -and (Get-Module -ListAvailable | Where-Object { $_."Name" -eq "Send-MailKitMessage" -and $_."PSEdition" -eq $PowerShellEdition }))
    )
{
    Remove-Module -Name "Send-MailKitMessage"
}

#IMPORTANT: the dll's used by Send-MailKitMessage do not get unloaded unless the session is restarted (the "Kill Terminal" button in Visual Studio Code), failure to restart the session can cause "assembly is already loaded" errors
Invoke-Expression "using module `"..\Project\bin\Debug\netstandard2.0\Send-MailKitMessage.psd1`""

#ensure the development version of the module is being used (not the real version if it is installed)
Get-Module -Name "Send-MailKitMessage" | Select-Object -Property "ModuleBase"

#set parameter file location
$ParametersFileLocation = ".\Parameters.csv"

#create parameter file
if (-not (Test-Path -Path $ParametersFileLocation))
{
    [PSCustomObject]@{
        UseSecureConnectionIfAvailable = [string]::Empty
        Username = [string]::Empty
        Password = [string]::Empty
        SMTPServer = [string]::Empty
        Port = [string]::Empty
        Priority = [string]::Empty
        From = [string]::Empty
        To = [string]::Empty
        CC = [string]::Empty
        BCC = [string]::Empty
        Subject = [string]::Empty
        TextBody = [string]::Empty
        HTMLBody = [string]::Empty
        AttachmentLocation = [string]::Empty
    } | Export-Csv -Path ".\Parameters.csv"
}

#get parameters (choose one row and skip a set number of rows to allow multiple parameter lines in the csv file)
$ParametersFile = (Import-Csv -Path $ParametersFileLocation | Select-Object -First 1 -Skip 0)

#the below code is (mostly) from the README, using it here helps ensure the documentation is valid
#any material changes required here should be made in the README as well

$Parameters = @{

    #Use secure connection if available (optional)
    "UseSecureConnectionIfAvailable" = if ([string]::IsNullOrWhiteSpace($ParametersFile."UseSecureConnectionIfAvailable")) { $false } else { [bool]$ParametersFile."UseSecureConnectionIfAvailable" }

    #authentication (optional)
    "Credential" = if ([string]::IsNullOrWhiteSpace($ParametersFile."Username") -or [string]::IsNullOrWhiteSpace($ParametersFile."Password")) { $null } else { [System.Management.Automation.PSCredential]::new($ParametersFile."Username", (ConvertTo-SecureString -String $ParametersFile."Password" -AsPlainText -Force)) }

    #SMTP server (required)
    "SMTPServer" = $ParametersFile."SMTPServer"

    #Port (required)
    "Port" = $ParametersFile."Port"

    #Priority (optional)
    "Priority" = if ([string]::IsNullOrWhiteSpace($ParametersFile."Priority")) { $null } else { [string]$ParametersFile."Priority" }

    #Sender (required) (http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm)
    "From" = [MimeKit.MailboxAddress]$ParametersFile."From"

    #Recipient list (at least one recipient required) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "RecipientList" = [MimeKit.InternetAddressList]$ParametersFile."To"
    #"RecipientList" = [MimeKit.InternetAddressList]::new([System.Collections.Generic.List[MimeKit.InternetAddress]](Import-Csv -Path "FileWithEmailAddresses.csv" | ForEach-Object { [MimeKit.InternetAddress]$_."Email" })) #example of importing a list of email addresses

    #CC list (optional) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "CCList" = if ([string]::IsNullOrWhiteSpace($ParametersFile."CC")) { $null } else { [MimeKit.InternetAddressList]$ParametersFile."CC" }

    #BCC list (optional) (http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm)
    "BCCList" = if ([string]::IsNullOrWhiteSpace($ParametersFile."BCC")) { $null } else { [MimeKit.InternetAddressList]$ParametersFile."BCC" }
    
    #Subject (required)
    "Subject" = $ParametersFile."Subject"
    
    #Text body (optional)
    "TextBody" = $ParametersFile."TextBody"
    
    #HTML body (optional)
    "HTMLBody" = $ParametersFile."HTMLBody"
    
    #Attachment list (optional)
    "AttachmentList" = if ([string]::IsNullOrWhiteSpace($ParametersFile."AttachmentLocation")) { $null } else { [System.Collections.Generic.List[string]]$ParametersFile."AttachmentLocation" }

}

#use Invoke-Expression so that PowerShell doesn't auto-load the Send-MailKitMessage module (which causes problems with trying to use the development version)
Invoke-Expression "Send-MailKitMessage @Parameters"