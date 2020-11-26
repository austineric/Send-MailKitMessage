

####################################
# Author:       Eric Austin - https://github.com/austineric/Send-MailKitMessage
# Create date:  June 2020
# Description:  Uses MailKit to send email because Send-MailMessage is marked obsolete (https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage)
####################################

using namespace MailKit
using namespace MimeKit

#extend classes to be available to the calling script (MimeKit assembly is loaded first from the manifest file so is available when this module loads)
class MailboxAddressExtended : MailboxAddress { #does not allow parameterless construction
    MailboxAddressExtended([string]$Name, [string]$Address) : base($Name, $Address) {
        [string]$Name,      #can be null
        [string]$Address    #cannot be null
    }
}
class InternetAddressListExtended : InternetAddressList {}

function Send-MailKitMessage(){
    param(
        [Parameter(Mandatory=$false)][pscredential]$Credential,
        [Parameter(Mandatory=$true)][string]$SMTPServer,
        [Parameter(Mandatory=$true)][string]$Port,
        [Parameter(Mandatory=$true)][MailboxAddress]$From,
        [Parameter(Mandatory=$true)][InternetAddressList]$ToList,
        [Parameter(Mandatory=$false)][InternetAddressList]$CCList,
        [Parameter(Mandatory=$false)][InternetAddressList]$BCCList,
        [Parameter(Mandatory=$false)][string]$Subject,
        [Parameter(Mandatory=$false)][string]$TextBody,
        [Parameter(Mandatory=$false)][string]$HTMLBody,
        [Parameter(Mandatory=$false)][string[]]$AttachmentList
    )

    Try {

        $ErrorActionPreference="Stop"

        #message
        $Message=New-Object MimeMessage

        #from
        $Message.From.Add($From)

        #to
        $Message.To.AddRange($ToList)

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
        $BodyBuilder=New-Object BodyBuilder
        
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
        $Client.Connect($SMTPServer, $Port, [Security.SecureSocketOptions]::Auto)
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