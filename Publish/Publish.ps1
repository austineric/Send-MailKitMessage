

####################################
# Author:       Eric Austin
# Create date:  August 2020
# Description:  Publish module to the PowerShell gallery including version prompt
####################################

using namespace System.Collections.Generic

Try {

    #script variables
    $ManifestLocation="..\Send-MailKitMessage.psd1"
    [List[string]]$Manifest=New-Object -TypeName List[string]
    [int]$VersionLineIndex=0
    $CurrentModuleVersion=[string]::Empty
    $NewModuleVersion=[string]::Empty
    $Confirm=[string]::Empty

    #get manifest content
    $Manifest=Get-Content -Path $ManifestLocation
    
    #get index of the line with the current module version
    $VersionLineIndex=$Manifest.IndexOf(($Manifest | Where-Object { $_ -like "ModuleVersion = '*'"}))

    #parse out the current module version
    $CurrentModuleVersion=$Manifest[$VersionLineIndex].Replace("'", "").Replace("ModuleVersion = ", "")

    Clear-Host
    
    #list semantic versioning
    Write-Host ""
    Write-Host "Versioning is MAJOR.MINOR.PATCH"
    Write-Host "MAJOR version when you make incompatible API changes"
    Write-Host "MINOR version when you add functionality in a backwards compatible manner"
    Write-Host "PATCH version when you make backwards compatible bug fixes"

    #prompt for new version number
    Do {
        Write-Host ""
        Write-Host "Current module version is $CurrentModuleVersion."
        $NewModuleVersion=(Read-Host "Enter new version number")
        $Confirm=(Read-Host "New version will be $NewModuleVersion. Continue? (y/n)")
    }
    Until ($Confirm -eq "y")

    #update manifest
    $Manifest[$VersionLineIndex]="ModuleVersion = '$NewModuleVersion'"
    Set-Content -Path $ManifestLocation -Value $Manifest
    
    #publish module
    Write-Host "Publishing..."
    Publish-Module -Path "..\..\Send-MailKitMessage" -NuGetApiKey $env:PowerShellGalleryAPIKey

    #open module page in PowerShell gallery
    Start-Process "https://www.powershellgallery.com/packages/Send-MailKitMessage"

    Write-Host "Completed"

}
Catch{

    Throw $Error[0]

}





