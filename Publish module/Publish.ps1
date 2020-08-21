

Try {
    Publish-Module -Path "..\Send-MailKitMessage" -NuGetApiKey $env:PowerShellGalleryAPIKey
}
Catch{
    Throw $Error[0]
}

