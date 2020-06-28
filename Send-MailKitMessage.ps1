

####################################
# Author:       Eric Austin
# Create date:  June 2020
# Description:  Functions as a central location for sending email
#               Uses MailKit because Microsoft's System.Net.Mail.SmtpClient is marked obsolete; it still works but it's probably better to use MailKit (the recommended alternative)
#               The best way to get the latest version for each assembly is to download the package from NuGet and unzip the .nupkg using 7Zip
####################################

function SendEmail(){
    param(
        [Parameter(Mandatory=$true)][string]$From,
        [Parameter(Mandatory=$true)][string[]]$ToList,
        [Parameter(Mandatory=$false)][string[]]$BCCList,
        [Parameter(Mandatory=$false)][string[]]$Subject,
        [Parameter(Mandatory=$false)][string[]]$HTMLBody,
        [Parameter(Mandatory=$false)][string[]]$AttachmentList
    )

    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"

    Try {

        #load assemblies in specific order
        Add-Type -Path (Join-Path -Path $CurrentDirectory -ChildPath (Join-Path -Path "MailKit" -ChildPath "BouncyCastle.Crypto.dll"))
        Add-Type -Path (Join-Path -Path $CurrentDirectory -ChildPath (Join-Path -Path "MailKit" -ChildPath "MimeKit.dll"))
        Add-Type -Path (Join-Path -Path $CurrentDirectory -ChildPath (Join-Path -Path "MailKit" -ChildPath "MailKit.dll"))

        #message
        $Message=New-Object MimeMessage

        #from
        $From=New-Object MailboxAddress($From)
        $Message.From.Add($From)

        #to
        $To=New-Object InternetAddressList
        $ToList | ForEach-Object {
            $To.Add($_)
        }
        $Message.To.AddRange($To)

        #bcc
        if ($BCCList.Count -gt 0)
        {
            $BCC=New-Object InternetAddressList
            $BCCList | ForEach-Object {
                $BCC.Add($_)
            }
            $Message.Bcc.AddRange($BCC)
        }

        #subject
        if ( -not ([string]::IsNullOrWhiteSpace($Message)))
        {
            $Message.Subject=$Subject
        }

        #html body (use [System.Web.HttpUtility]::HtmlDecode(TextToDecode) in case there are html elements present that have been escaped)
        $BodyBuilder=New-Object BodyBuilder
        if ( -not ([string]::IsNullOrWhiteSpace($Message)))
        {
            $BodyBuilder.HtmlBody=[System.Web.HttpUtility]::HtmlDecode($HTMLBody)
        }
        
        #attachment(s)
        if ($AttachmentList.Count -gt 0)
        {
            $AttachmentList | ForEach-Object {
                $BodyBuilder.Attachments.Add($_) | Out-Null
            }
        }

        #add bodybuilder to message body
        $Message.Body=$BodyBuilder.ToMessageBody()

        #smtp send
        $Client=New-Object MailKit.Net.Smtp.SmtpClient
        $Client.Connect("SMTPServerName", "PortNumber")
        $Client.Send($Message)

    }

    Catch {
        
        Throw $Error[0] #it may be redundant to have just a throw inside the catch but it correctly returns exceptions to the calling script so it's good

    }

    Finally {
        
        if ($Client.IsConnected)
        {
            $Client.Disconnect($true)
        }

    }

}