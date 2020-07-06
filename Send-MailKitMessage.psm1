

####################################
# Author:       Eric Austin
# Create date:  June 2020
# Description:  Uses MailKit to send email because Send-MailMessage is marked obsolete (https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/send-mailmessage?view=powershell-7)
#               Works with .Net Core
####################################

using namespace MailKit
using namespace MimeKit

#load assemblies in specific order
Add-Type -Path ".\MailKit\BouncyCastle.Crypto.dll"
Add-Type -Path ".\MailKit\MailKit.dll"
Add-Type -Path ".\MailKit\MimeKit.dll"

#extend InternetAddressList class so it can be available to the calling script
class InternetAddressListExtended : InternetAddressList {}

function Send-MailKitMessage(){
    param(
        [Parameter(Mandatory=$true)][string]$From,
        [Parameter(Mandatory=$true)][InternetAddressList]$ToList,
        [Parameter(Mandatory=$false)][InternetAddressList]$BCCList,
        [Parameter(Mandatory=$false)][string[]]$Subject,
        [Parameter(Mandatory=$false)][string[]]$HTMLBody,
        [Parameter(Mandatory=$false)][string[]]$AttachmentList
    )

    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"

    Try {

        #create the SMTPCredentials.csv file if it does not exist (the nickname field allows multiple servers to be added to the file and selected programmatically ie prod vs test)
        if (-not (Test-Path -Path (Join-Path -Path $CurrentDirectory -ChildPath "SMTPCredentials.csv")))
        {
            New-Object -TypeName PSCustomObject -Property @{ "Nickname"=""; "SMTPServer"=""; "Port"=""} | Select-Object -Property "Nickname", "SMTPServer", "Port" | Export-Csv -Path (Join-Path -Path $CurrentDirectory -ChildPath "SMTPCredentials.csv")
        }

        #message
        $Message=New-Object MimeMessage

        #from
        $From=New-Object MailboxAddress($From)
        $Message.From.Add($From)

        #to
        $Message.To.AddRange($ToList)

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