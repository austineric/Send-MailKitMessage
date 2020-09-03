

####################################
# Author:       Eric Austin - https://github.com/austineric/Send-MailKitMessage
# Create date:  June 2020
# Description:  Uses MailKit to send email because Send-MailMessage is marked obsolete (https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage)
#               Works with .Net Core
####################################

using namespace MailKit
using namespace MimeKit

#extend classes to be available to the calling script (MimeKit assembly is loaded first from the manifest file so is available when this module loads)
class MailboxAddressExtended : MailboxAddress { #does not allow parameterless construction
    MailboxAddressExtended([string]$Name, [string]$Address ) : base($Name, $Address) {
        [string]$Name,      #can be null
        [string]$Address    #cannot be null
    }
}
class InternetAddressListExtended : InternetAddressList {}

function Send-MailKitMessage(){
    param(
        [Parameter(Mandatory=$true)][string[]]$SMTPServer,
        [Parameter(Mandatory=$true)][string[]]$Port,
        [Parameter(Mandatory=$true)][MailboxAddress]$From,
        [Parameter(Mandatory=$true)][InternetAddressList]$ToList,
        [Parameter(Mandatory=$false)][InternetAddressList]$CCList,
        [Parameter(Mandatory=$false)][InternetAddressList]$BCCList,
        [Parameter(Mandatory=$false)][string]$Subject,
        [Parameter(Mandatory=$false)][string]$HTMLBody,
        [Parameter(Mandatory=$false)][string[]]$AttachmentList
    )

    Try {

        #common variables
        $ErrorActionPreference="Stop"

        #message
        $Message=New-Object MimeMessage

        #from
        $From=New-Object MailboxAddress($From)
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
        $Client.Connect($SMTPServer, $Port)
        $Client.Send($Message)

    }

    Catch {
        
        Throw $Error[0]

    }

    Finally {
        
        if ($Client.IsConnected)
        {
            $Client.Disconnect($true)
        }

    }

}