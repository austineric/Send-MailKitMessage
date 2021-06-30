

####################################
# Author:       Eric Austin
# Create date:  May 2021
# Description:  Publish Send-MailKitMessage to the PowerShell Gallery
####################################

#versioning is not particularly simple (article is https://docs.microsoft.com/en-us/dotnet/standard/library-guidance/versioning)
# Article                           | What I see
# --------------------------------- | ----------
# PSGallery module manifest version | PSGallery module manifest version
# Not referenced                    | .csproj "Version" / "Product Version" in Windows Explorer; does not affect runtime
# Assembly version                  | .csproj "AssemblyVersion"; not shown in Windows Explorer; does affect runtime
# Assembly file version             | .csproj "FileVersion"; shown in Windows Explorer; does not affect runtime
# Assembly informational version    | Ignore

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"
    
    #project elements
    $ProjectCSProjFilePath=(Join-Path -Path $CurrentDirectory -ChildPath ".." -AdditionalChildPath "Project", "Send-MailKitMessage.csproj")

    #published elements
    $PublishedModuleManifest=@{}
    $PublishedCSProjFile=[xml]::new()
    $PublishDirectory=(Join-Path -Path $CurrentDirectory -ChildPath ".." -AdditionalChildPath "Publish", "Send-MailKitMessage")
    $PublishedModuleManifestPath=(Join-Path -Path $PublishDirectory -ChildPath "Send-MailKitMessage.psd1")
    $PublishedCSProjPath=(Join-Path -Path $PublishDirectory -ChildPath "Send-MailKitMessage.csproj")

    #updated version values
    $UpdatedModuleManifestVersion=[string]::Empty
    $UpdatedModuleManifestPrereleaseString=[string]::Empty
    $UpdatedCSProjVersion=[string]::Empty
    $UpdatedCSProjAssemblyVersion=[string]::Empty
    $UpdatedCSProjFileVersion=[string]::Empty

    #script elements
    $ValuesConfirmed=[string]::Empty

    #--------------#

    #get the published module manifest
    $PublishedModuleManifest=(Import-PowerShellDataFile -Path $PublishedModuleManifestPath)

    #get the published csproj file
    $PublishedCSProjFile=[xml](Get-Content -Raw -Path $PublishedCSProjPath)
    
    #display the current versions
    Write-Host ""
    Write-Host "The published module manifest version is $($PublishedModuleManifest."ModuleVersion" + ([string]::IsNullOrWhiteSpace($PublishedModuleManifest.PrivateData.PSData.Prerelease) ? [string]::Empty : ($PublishedModuleManifest.PrivateData.PSData.Prerelease.Contains("-") ? [string]::Empty : "-") + $PublishedModuleManifest.PrivateData.PSData.Prerelease))"
    Write-Host "The published csproj version is $($PublishedCSProjFile.Project.PropertyGroup.Version)"
    Write-Host "The published csproj assembly version is $($PublishedCSProjFile.Project.PropertyGroup.AssemblyVersion)"
    Write-Host "The published csproj file version is $($PublishedCSProjFile.Project.PropertyGroup.FileVersion)"
    
    #prompt for new values
    Do {

        #updated module manifest version
        Do {
            Write-Host ""
            Write-Host "First: the module manifest version (the version the PSGallery uses)"
            Write-Host "Versioning is MAJOR.MINOR.PATCH"
            Write-Host "MAJOR version when you make incompatible API changes"
            Write-Host "MINOR version when you add functionality in a backwards compatible manner"
            Write-Host "PATCH version when you make backwards compatible bug fixes"
            Write-Host "The published module manifest version is $($PublishedModuleManifest."ModuleVersion")"
            $UpdatedModuleManifestVersion=(Read-Host "Enter new module version (the prerelease value, if applicable, will be obtained next)")
        }
        Until (-not ([string]::IsNullOrWhiteSpace($UpdatedModuleManifestVersion)))
        
        #module manifest prerelease string (dash is not required https://docs.microsoft.com/en-us/powershell/scripting/gallery/concepts/module-prerelease-support?view=powershell-7.1)
        Write-Host ""
        Write-Host "Second, the prerelease value (if applicable; gets appended on to the module manifest version)"
        if ([string]::IsNullOrWhiteSpace($PublishedModuleManifest.PrivateData.PSData.Prerelease))
        {
            Write-Host "The published module manifest does not have a prerelease value"        

        }
        else
        {
            Write-Host "The published module prerelease value is `"$($PublishedModuleManifest.PrivateData.PSData.Prerelease)`""
        }
        $UpdatedModuleManifestPrereleaseString=(Read-Host "Enter the new prerelease value, ie `"preview1`" (no dash) (if no prerelease value is applicable just hit Enter)")
        
        #updated csproj version (shows as "Product Version" when viewing the file properties in Windows Explorer; does not affect runtime; just set to the same value as the $UpdatedModuleManifestVersion and include the prerelease value if applicable)
        Write-Host ""
        Write-Host "The csproj version (doesn't affect anything; shows as `"Product Version`" when viewing the file properties in Windows Explorer) has no conventions and will be set to the same value as the updated module manifest version"
        Write-Host "The published csproj version is $($PublishedCSProjFile.Project.PropertyGroup.Version)"
        $UpdatedCSProjVersion=$UpdatedModuleManifestVersion + ([string]::IsNullOrWhiteSpace($UpdatedModuleManifestPrereleaseString) ? [string]::Empty : ($UpdatedModuleManifestPrereleaseString.Contains("-") ? [string]::Empty : "-") + $UpdatedModuleManifestPrereleaseString)
        Write-Host "The updated csproj version will be $UpdatedCSProjVersion"

        #updated csproj assembly version
        Write-Host ""
        Write-Host "The csproj assembly version (affects the runtime; does not get shown in Windows Explorer) should be set to the MAJOR version of the updated module manifest version"
        Write-Host "The published csproj assembly version is $($PublishedCSProjFile.Project.PropertyGroup.AssemblyVersion)"
        $UpdatedCSProjAssemblyVersion=$UpdatedModuleManifestVersion.Substring(0, ($UpdatedModuleManifestVersion).IndexOf(".")) + ".0.0"
        Write-Host "The updated csproj assembly version will be $UpdatedCSProjAssemblyVersion"

        #updated file version
        Write-Host ""
        Write-Host "The csproj file version (doesn't affect anything; shows as `"File Version`" when viewing the file properties in Windows Explorer) should be set to MAJOR.MINOR.BUILD.REVISION"
        Write-Host "The published csproj file version is $($PublishedCSProjFile.Project.PropertyGroup.FileVersion)"
        [int]$Revision=($PublishedCSProjFile.Project.PropertyGroup.FileVersion).Substring($PublishedCSProjFile.Project.PropertyGroup.FileVersion.LastIndexOf(".") + 1)
        $Revision++
        $UpdatedCSProjFileVersion=$UpdatedModuleManifestVersion + "." + $Revision.ToString()
        Write-Host "The updated csproj file version will be $UpdatedCSProjFileVersion"
        
        Write-Host ""
        Write-Host "The new module manifest version will be $($UpdatedModuleManifestVersion + ([string]::IsNullOrWhiteSpace($UpdatedModuleManifestPrereleaseString) ? [string]::Empty : ($UpdatedModuleManifestPrereleaseString.Contains("-") ? [string]::Empty : "-") + $UpdatedModuleManifestPrereleaseString))"
        Write-Host "The new csproj version will be $UpdatedCSProjVersion"
        Write-Host "The new csproj assembly version will be $UpdatedCSProjAssemblyVersion"
        Write-Host "The new csproj file version will be $UpdatedCSProjFileVersion"

        Write-Host ""
        $ValuesConfirmed=(Read-Host "Proceed with publishing? (`"y`" to proceed, `"n`" to re-enter values)")

    }
    Until ($ValuesConfirmed -eq "y")

    #update the .csproj file
    $CSProjFile=[xml](Get-Content -Raw -Path $ProjectCSProjFilePath)
    $CSProjFile.Project.PropertyGroup.Version=$UpdatedCSProjVersion
    $CSProjFile.Project.PropertyGroup.AssemblyVersion=$UpdatedCSProjAssemblyVersion
    $CSProjFile.Project.PropertyGroup.FileVersion=$UpdatedCSProjFileVersion
    $CSProjFile.Save($ProjectCSProjFilePath)

    #run dotnet publish (the build part of dotnet publish copies the .csproj file into the "Publish" directory to become the source of published version information later)
    dotnet publish $ProjectCSProjFilePath --configuration "Release" --output $PublishDirectory /p:Version=$UpdatedCSProjVersion /p:AssemblyVersion=$UpdatedCSProjAssemblyVersion /p:FileVersion=$UpdatedCSProjFileVersion

    #update the module manifest (run in the "Publish" directory after dotnet build since the module manifest update process checks for the included assemblies)
    Update-ModuleManifest -Path $PublishedModuleManifestPath -ModuleVersion $UpdatedModuleManifestVersion -Prerelease $UpdatedModuleManifestPrereleaseString

    #publish to the PSGallery
    Publish-Module -Path $PublishDirectory -Name "Send-MailKitMessage" -Exclude "Send-MailKitMessage.csproj" -AllowPrerelease -NuGetApiKey $env:PSGalleryAPIKey -WhatIf -Verbose #I wonder if something is broke here, because it seems like this should work
    Publish-Module -Path $PublishDirectory -Repository PSGallery -NuGetApiKey $env:PSGalleryAPIKey -WhatIf #so this works
    Publish-Module -Path $PublishDirectory -Exclude "Send-MailKitMessage.csproj" -Repository PSGallery -NuGetApiKey $env:PSGalleryAPIKey -WhatIf -Verbose
    Publish-Module -Path $PublishDirectory -Repository PSGallery -NuGetApiKey $env:PSGalleryAPIKey -Exclude "Send-MailKitMessage.csproj" -WhatIf #doesn't work
    Publish-Module -Path $PublishDirectory -Repository PSGallery -AllowPrerelease -NuGetApiKey $env:PSGalleryAPIKey -WhatIf #doesn't work
     -Exclude "Send-MailKitMessage.csproj" -AllowPrerelease

}

Catch {

    Throw $Error[0]
	
}