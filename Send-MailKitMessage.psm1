

####################################
# Author:       Eric Austin
# Repo:         https://github.com/austineric/Send-MailKitMessage
# Description:  A replacement for PowerShell's obsolete Send-MailMessage implementing the Microsoft-recommended MailKit library.
####################################

function Send-MailKitMessage(){
    param(
        [Parameter(Mandatory=$false)][bool]$UseSecureConnectionIfAvailable,
        [Parameter(Mandatory=$false)][pscredential]$Credential,
        [Parameter(Mandatory=$true)][string]$SMTPServer,
        [Parameter(Mandatory=$true)][int]$Port,
        [Parameter(Mandatory=$true)][MimeKit.MailboxAddress]$From,
        [Parameter(Mandatory=$true)][MimeKit.InternetAddressList]$RecipientList,
        [Parameter(Mandatory=$false)][MimeKit.InternetAddressList]$CCList,
        [Parameter(Mandatory=$false)][MimeKit.InternetAddressList]$BCCList,
        [Parameter(Mandatory=$false)][string]$Subject,
        [Parameter(Mandatory=$false)][string]$TextBody,
        [Parameter(Mandatory=$false)][string]$HTMLBody,
        [Parameter(Mandatory=$false)][string[]]$AttachmentList
    )

    Try {

        $ErrorActionPreference="Stop"

        #message
        $Message=[MimeKit.MimeMessage]::new()

        #from
        $Message.From.Add($From)

        #to
        $Message.To.AddRange($RecipientList)

        #cc
        if ($CCList.Count -gt 0)
        {
            $Message.Cc.AddRange($CCList)
        }
        
        #bcc
        if ($BCCList.Count -gt 0)
        {
            $Message.Bcc.AddRange($BCCList)
        }

        #subject
        if ( -not ([string]::IsNullOrWhiteSpace($Message)))
        {
            $Message.Subject=$Subject
        }

        #body
        #$BodyBuilder=New-Object BodyBuilder
        $BodyBuilder=[MimeKit.BodyBuilder]::new()
        
        #text body
        if (-not ([string]::IsNullOrWhiteSpace($TextBody)))
        {
            $BodyBuilder.TextBody=$TextBody
        }

        #html body
        if (-not ([string]::IsNullOrWhiteSpace($HTMLBody)))
        {
            $BodyBuilder.HtmlBody=[System.Web.HttpUtility]::HtmlDecode($HTMLBody)   #use [System.Web.HttpUtility]::HtmlDecode() in case there are html elements present that have been escaped
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
        $Client.Connect($SMTPServer, $Port, ($UseSecureConnectionIfAvailable ? [MailKit.Security.SecureSocketOptions]::Auto : [MailKit.Security.SecureSocketOptions]::None))
        if ($Credential)
        {
            $Client.Authenticate($Credential.UserName, ($Credential.Password | ConvertFrom-SecureString -AsPlainText))
        }
        $Client.Send($Message)

    }

    Catch {
        
        Throw $Global:Error[0]

    }

    Finally {
        
        if ($Client.IsConnected)
        {
            $Client.Disconnect($true)
        }

    }

}