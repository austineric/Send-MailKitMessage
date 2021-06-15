

####################################
# Author:       Eric Austin
# Create date:  May 2021
# Description:  Publish Send-MailKitMessage to the PowerShell Gallery
####################################

#namespaces
#using namespace System.Data     #required for DataTable
#using namespace System.Data.SqlClient
#using namespace System.Collections.Generic  #required for List<T>
#using module Send-MailKitMessage

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"

    #project elements
    $ProjectModuleManifestPath=(Join-Path -Path $CurrentDirectory -ChildPath "Project" -AdditionalChildPath "Send-MailKitMessage.psd1")
    $ProjectCSProjFilePath=(Join-Path -Path ".." -ChildPath "Project" -AdditionalChildPath "Send-MailKitMessage.csproj")
    $ProjectCSProjFile=[xml]::new()

    #published elements
    $PublishedModuleManifest=@{}
    $PublishedCSProjFile=[xml]::new()
    $PublishedModuleManifestPath=(Join-Path -Path $CurrentDirectory -ChildPath "Publish" -AdditionalChildPath "Send-MailKitMessage.psd1")
    $PublishedCSProjPath=(Join-Path -Path $CurrentDirectory -ChildPath "Publish" -AdditionalChildPath "Send-MailKitMessage.csproj")

    #updated version values
    $UpdatedModuleManifestVersion=[string]::Empty
    $UpdatedModuleManifestPrereleaseString=[string]::Empty
    $UpdatedCSProjVersion=[string]::Empty
    $UpdatedCSProjAssemblyVersion=[string]::Empty
    $UpdatedCSProjFileVersion=[string]::Empty

    #--------------#

    #get the published module manifest
    $PublishedModuleManifest=(Import-PowerShellDataFile -Path $PublishedModuleManifestPath).PrivateData.PSData
    $PublishedModuleManifest=(Import-PowerShellDataFile -Path "C:\Users\Eric\Desktop\2021.05.20T13.16.01\Send-MailKitMessage.psd1").PrivateData.PSData
    $PublishedModuleManifest.PrivateData.PSData.Prerelease
    #get the published csproj file
    $CSProjFile=[xml](Get-Content -Raw -Path $PublishedCSProjFile)

    Write-Host ""
    Write-Host "The published module manifest version is $($PublishedModuleManifest."ModuleVersion" + ([string]::IsNullOrWhiteSpace($PublishedModuleManifest.PrivateData.PSData.Prerelease) ? """" : "-" + $PublishedModuleManifest.PrivateData.PSData.Prerelease))"
    Write-Host ""   #I should make a copy of the manifest, set a prerelease string, import it, and make sure I can display it correctly
    Write-Host ""

    # So the script needs to:
        # display current values
        # prompt for new values (including prerelease)
        # Update the module manifest "ModuleVersion"
        # Update .csproj "Version"
        # Update .csproj "AssemblyVersion"
        # Update .csproj "FileVersion"
        # Update the module manifest "ModuleVersion"
        # Then "dotnet publish /p:Version=3.3.3 /p:AssemblyVersion=4.4.4 /p:FileVersion=5.5.5" (dotnet publish should copy the updated module manifest to the publish destination)









    #get the module version from the manifest so everything is coming from the downloaded version
    #get the module version from PSGallery (accounts for prerelease module; output includes prerelease string if present)
    #$PublishedModuleManifestVersion=(Find-Module -Name "Send-MailKitMessage" -AllowPrerelease -Repository "PSGallery").Version
    #(Find-Module -Name "azure.databricks.cicd.tools" -AllowPrerelease).Version

    #get the published moduled manifest version
    $Data=(Get-Content -Path $PublishedModuleManifestPath)
    $Data=@{}
    $Data.GetType()
    $Data[3]
    $Data["ModuleVersion"]
    
    

    #hmm, it looks like there isn't really a good way to get the csproj data (I mean some of it comes from the file itself but not all of it)
    #perhaps the build should copy the csproj file to the publish folder and then that file should be excluded from the psgallery publish

    #get the module version from the manifest
    $PublishedModuleManifestVersion=([System.IO.FileInfo](Join-Path -Path $PSGalleryModuleDownloadDirectory -ChildPath "Send-MailKitMessage" -AdditionalChildPath "*","Send_MailKitMessage.dll" -Resolve)).VersionInfo.ProductVersion

    #get the prerelease string from the manifest



    #get the published assembly version and the published file version
    $PublishedCSProjAssemblyVersion=([System.IO.FileInfo](Join-Path -Path $PSGalleryModuleDownloadDirectory -ChildPath "Send-MailKitMessage" -AdditionalChildPath "Send_MailKitMessage.dll")).VersionInfo.ProductVersion
    $PublishedCSProjAssemblyVersion=([System.IO.FileInfo](Join-Path -Path $PSGalleryModuleDownloadDirectory -ChildPath "Send-MailKitMessage" -AdditionalChildPath "*","Send_MailKitMessage.dll" -Resolve)).VersionInfo.

    $PublishedCSProjFileVersion=([System.IO.FileInfo](Join-Path -Path $PSGalleryModuleDownloadDirectory -ChildPath "Send-MailKitMessage" -AdditionalChildPath "Send_MailKitMessage.dll")).VersionInfo.FileVersion

    ([System.IO.FileInfo]"C:\Users\Eric\AppData\Local\Temp\Send-MailKitMessage\20210517190616\Send-MailKitMessage\3.1.0\Send_MailKitMessage.dll").VersionInfo

    #display the current versions and prompt for new value(s), including prerelease, major/minor/build/revision, etc
    
    #module manifest version
    Write-Host ""
    Write-Host "Versioning is MAJOR.MINOR.PATCH"
    Write-Host "MAJOR version when you make incompatible API changes"
    Write-Host "MINOR version when you add functionality in a backwards compatible manner"
    Write-Host "PATCH version when you make backwards compatible bug fixes"
    Write-Host "The published module version is $PublishedModuleManifestVersion"
    $UpdatedModuleManifestVersion=(Read-Host "Enter new module version")

    #module manifest prerelease string
    #does this already get displayed in the published module manifest version?
    #okay, the version and the module prerelease string are separate things

    #updated csproj version

    #csproj prerelease element?

    #updated csproj assembly version

    #updated csproj file version

    #update the module manifest
    Update-ModuleManifest -Path $ModuleManifestPath -ModuleVersion $UpdatedModuleManifestVersion -Prerelease $UpdatedModuleManifestPrereleaseString

    #update the .csproj file
    $CSProjFile=[xml](Get-Content -Raw -Path $CSProjFilePath)
    $CSProjFile.Project.PropertyGroup.Version=$UpdatedCSProjVersion
    $CSProjFile.Project.PropertyGroup.AssemblyVersion=$UpdatedCSProjAssemblyVersion
    $CSProjFile.Project.PropertyGroup.FileVersion=$UpdatedCSProjFileVersion
    $CSProjFile.Save($CSProjFilePath)

    #now I think I can do the publish (I'm not sure if I can put PowerShell variables in or if I need to build the whole string first or something)
    #I forget there may be a prerelease thing involved in the csproj file
    dotnet publish --configuration "Release" --output $PublishDirectory /p:Version=$UpdatedCSProjVersion /p:AssemblyVersion=$UpdatedCSProjAssemblyVersion /p:FileVersion=$UpdatedCSProjFileVersion

    #then if the publish succeeds I think I could publish to the PSGallery

    #probably clean up the module download directory and any parents I created

}

Catch {

    #error log
    #$ErrorData+=New-Object -TypeName PSCustomObject -Property @{"Date"=(Get-Date).ToString(); "ErrorMessage"=$Error[0].ToString()}    #don't use @Date for the date, this section needs to be completely independent so nothing can ever interfere with the error log being created
    #$ErrorData | Select-Object Date,ErrorMessage | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation

    #return value
    Exit 1
    
}

Finally {


}