

####################################
# Author:       Eric Austin
# Create date:  August 2020
# Description:  Publish module to the PowerShell gallery including version prompt
####################################

using namespace System.Collections.Generic

Try {

    #script variables
    $ManifestLocation="..\Send-MailKitMessage.psd1"
    $NewModuleVersion=[string]::Empty
    $Confirm=[string]::Empty

    function Get-ModuleVersion {
        return [PSCustomObject]@{
            "LocalManifestVersion"=((Get-Content -Path $ManifestLocation) | Where-Object { $_.Contains("ModuleVersion =") }).Replace("'", "").Replace("ModuleVersion = ", "")
            "PSGalleryVersion"=(Find-Module -Name "Send-MailKitMessage")."Version"
        }
    }

    Clear-Host
    Write-Host ""

    #list semantic versioning
    Write-Host ""
    Write-Host "Versioning is MAJOR.MINOR.PATCH"
    Write-Host "MAJOR version when you make incompatible API changes"
    Write-Host "MINOR version when you add functionality in a backwards compatible manner"
    Write-Host "PATCH version when you make backwards compatible bug fixes"

    #get local manifest version and PSGallery module version (use Out-Host to force it to write to the host before displaying the Read-Host below)
    Write-Host ""
    Write-Host "Getting module versions:"
    (Get-ModuleVersion) | Select-Object -Property "LocalManifestVersion", "PSGalleryVersion" | Out-Host

    #prompt for new version number
    Do {
        Write-Host ""
        $NewModuleVersion=(Read-Host "Enter new module version number")
        $Confirm=(Read-Host "New version will be $NewModuleVersion. Continue? (y/n)")
    }
    Until ($Confirm -eq "y")

    #update manifest
    Write-Host "Updating module manifest..."
    Update-ModuleManifest -Path $ManifestLocation -ModuleVersion $NewModuleVersion

    #publish module
    Write-Host "Publishing module to the PSGallery..."
    Publish-Module -Path "..\..\Send-MailKitMessage" -NuGetApiKey $env:PowerShellGalleryAPIKey  #yes this goes outside the root Send-MailKitMessage directory, you just pass Publish-Module the directory which contains the module

    #display current local manifest version and PSGallery module version
    Write-Host "Getting module versions:"
    (Get-ModuleVersion) | Select-Object -Property "LocalManifestVersion", "PSGalleryVersion" | Out-Host

    #open module page in PowerShell gallery
    Start-Process "https://www.powershellgallery.com/packages/Send-MailKitMessage"

    Write-Host "If all looks as it should the update can be committed to source control"

}
Catch{

    Throw $Error[0]

}





