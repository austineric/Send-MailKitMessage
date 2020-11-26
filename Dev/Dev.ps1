

#set location to dev directory
Set-Location -Path (Join-Path -Path "." -ChildPath "Dev")

#remove module from current session and load module
Remove-Module -Name "Send-MailKitMessage"
Invoke-Expression "using module `"..\..\Send-MailKitMessage`""

#create parameter file
$ParameterFileLocation=".\Parameters.csv"
if (-not (Test-Path -Path $ParameterFileLocation))
{
    [PSCustomObject]@{
        Username = ""
        Password = ""
        SMTPServer = ""
        Port = ""
        From = ""
        To = ""
        CC = ""
        BCC = ""
        Subject = ""
        TextBody = ""
        AttachmentLocation = ""
    } | Export-Csv -Path ".\Parameters.csv"
}

#get parameters
$Parameters=Import-Csv -Path $ParameterFileLocation

#set parameters
$Username=$Parameters."Username"
$Password=ConvertTo-SecureString -String $Parameters."Password" -AsPlainText -Force
$Credential=New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password
$SMTPServer=$Parameters."SMTPServer"
$Port=$Parameters."Port"
$From=New-Object MailboxAddressExtended($null, "$($Parameters.From)")
($ToList=New-Object InternetAddressListExtended).Add("$($Parameters.To)")
#($CCList=New-Object InternetAddressListExtended).Add("$($Parameters.CC)")
#($BCCList=New-Object InternetAddressListExtended).Add("$($Parameters.BCC)")
$Subject=$Parameters."Subject"
$TextBody=$Parameters."TextBody"
$HTMLBody=$Parameters."HTMLBody"
#$AttachmentList.Add($Parameters."AttachmentLocation")

Send-MailKitMessage -Credential $Credential -SMTPServer $SMTPServer -Port $Port -From $From -ToList $ToList -CCList $CCList -BCCList $BCCList -Subject $Subject -TextBody $TextBody -HTMLBody $HTMLBody -AttachmentList $AttachmentList

