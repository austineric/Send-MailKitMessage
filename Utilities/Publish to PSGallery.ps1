

####################################
# Author:       Eric Austin
# Create date:  May 2021
# Description:  Publish Send-MailKitMessage to the PowerShell Gallery
####################################

#versioning is not particularly simple (https://docs.microsoft.com/en-us/dotnet/standard/library-guidance/versioning)
# Listed in documentation           | What I see
# --------------------------------- | ----------
# PSGallery module manifest version | PSGallery module manifest version
# Not referenced                    | .csproj "Version" / "Product Version" in Windows Explorer; does not affect runtime
# Assembly version                  | .csproj "AssemblyVersion"; not shown in Windows Explorer; does affect runtime
# Assembly file version             | .csproj "FileVersion"; shown in Windows Explorer; does not affect runtime
# Assembly informational version    | Ignore

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot;
    $ErrorActionPreference = "Stop";

    #project elements
    $ProjectDirectory = Join-Path -Path $CurrentDirectory -ChildPath ".." -AdditionalChildPath "Project";
    $PublishDirectory = Join-Path -Path $CurrentDirectory -ChildPath ".." -AdditionalChildPath "Project", "bin", "Publish";

    #published module elements
    $PublishedModuleDownloadDirectory = Join-Path -Path $CurrentDirectory -ChildPath "Published module";
    $PublishedModuleDirectory = [string]::Empty;
    $PublishedManifestData = [hashtable]::new();
    $PublishedCSProjFileData = [xml]::new();

    #updated module elements
    $UpdatedManifestVersion = [string]::Empty;
    $UpdatedManifestPrereleaseString = [string]::Empty;
    $UpdatedCSProjFileData = [xml]::new();
    $UpdatedCSProjVersion = [string]::Empty;
    $UpdatedCSProjAssemblyVersion = [string]::Empty;
    $UpdatedCSProjFileVersion = [string]::Empty;

    #script elements
    $ModuleName = "Send-MailKitMessage";
    $ManifestFileName = "Send-MailKitMessage.psd1";
    $CSProjFileName = "Send-MailKitMessage.csproj";
    $DefaultProgressPreferenceValue = "Continue";
    $ValuesConfirmed = [string]::Empty;

    #--------------#

    Write-Host ""
    Write-Host "Downloading latest module published to the PowerShell Gallery...";

    #clear out the published module download directory if it exists
    if (Test-Path -Path $PublishedModuleDownloadDirectory)
    {
        foreach ($File in (Get-ChildItem -Path $PublishedModuleDownloadDirectory))
        {
            Remove-Item -Path $File."FullName" -Recurse -Force;
        }
    }
    else    #create the directory if it does not already exist
    {
        New-Item -ItemType Directory -Path $PublishedModuleDownloadDirectory | Out-Null;
    }

    #ensure the project publish directory exists
    if (-not (Test-Path -Path $PublishDirectory))
    {
        New-Item -ItemType Directory -Path $PublishDirectory | Out-Null;
    }

    #download the module from the PowerShell gallery ("Find-Module returns the newest version of a module if no parameters are used that limit the version" - https://learn.microsoft.com/en-us/powershell/module/powershellget/find-module?view=powershell-7.3")
    $ProgressPreference = "SilentlyContinue";   #hide the progress indicator
    Find-Module -Repository "PSGallery" -Name $ModuleName -AllowPrerelease | Save-Module -Path $PublishedModuleDownloadDirectory;
    $ProgressPreference = $DefaultProgressPreferenceValue;  #restore the progress indicator

    #throw an exception if the module is not present
    if ((Get-ChildItem -Path $PublishedModuleDownloadDirectory | Measure-Object | Select-Object -Property "Count" -ExpandProperty "Count") -eq 0)
    {
        Throw "The downloaded module could not be found";
    }

    #get the published module directory (the use of Get-ChildItem is due to Save-Module saving the downloaded module using the module version without preview suffixes as the directory name)
    $PublishedModuleDirectory = Get-ChildItem -Path (Join-Path $PublishedModuleDownloadDirectory -ChildPath $ModuleName) | Select-Object -Property "FullName" -ExpandProperty "FullName";

    #get the published manifest data
    $PublishedManifestData = Import-PowerShellDataFile -Path (Join-Path -Path $PublishedModuleDirectory -ChildPath $ManifestFileName);

    #get the published .csproj file data
    $PublishedCSProjFileData = [xml](Get-Content -Raw -Path (Join-Path -Path $PublishedModuleDirectory -ChildPath $CSProjFileName));
    
    #display the current versions
    Write-Host "The published module manifest version is $($PublishedManifestData."ModuleVersion" + ([string]::IsNullOrWhiteSpace($PublishedManifestData."PrivateData"."PSData"."Prerelease") ? [string]::Empty : ($PublishedManifestData."PrivateData"."PSData"."Prerelease".Contains("-") ? [string]::Empty : "-") + $PublishedManifestData."PrivateData"."PSData"."Prerelease"))";
    Write-Host "The published .csproj version is $($PublishedCSProjFileData."Project"."PropertyGroup"."Version")";
    Write-Host "The published .csproj assembly version is $($PublishedCSProjFileData."Project"."PropertyGroup"."AssemblyVersion")";
    Write-Host "The published .csproj file version is $($PublishedCSProjFileData."Project"."PropertyGroup"."FileVersion")";
    
    #prompt for new values
    Do {
        #updated manifest version
        Do {
            Write-Host "";
            Write-Host "First: the module manifest version (the version the PSGallery uses)";
            Write-Host "Versioning is MAJOR.MINOR.PATCH";
            Write-Host "MAJOR version when you make incompatible API changes";
            Write-Host "MINOR version when you add functionality in a backwards compatible manner";
            Write-Host "PATCH version when you make backwards compatible bug fixes";
            Write-Host "The published module manifest version (without prerelease indicator) is $($PublishedManifestData."ModuleVersion")";
            $UpdatedManifestVersion = (Read-Host "Enter new module version (the prerelease value, if applicable, will be obtained next)");
        }
        Until (-not ([string]::IsNullOrWhiteSpace($UpdatedManifestVersion)));
        
        #manifest prerelease string (dash is not required https://docs.microsoft.com/en-us/powershell/scripting/gallery/concepts/module-prerelease-support?view=powershell-7.1)
        Write-Host "";
        Write-Host "Second, the prerelease value (if applicable; gets appended on to the module manifest version)";
        if ([string]::IsNullOrWhiteSpace($PublishedManifestData."PrivateData"."PSData"."Prerelease"))
        {
            Write-Host "The published module manifest does not have a prerelease value";

        }
        else
        {
            Write-Host "The published module prerelease value is `"$($PublishedManifestData."PrivateData"."PSData"."Prerelease")`"";
        }
        $UpdatedManifestPrereleaseString = (Read-Host "Enter the new prerelease value, ie `"preview1`" (no dash) (if no prerelease value is applicable just hit Enter)");
        
        #updated csproj version (shows as "Product Version" when viewing the file properties in Windows Explorer; does not affect runtime; just set to the same value as the $UpdatedManifestVersion and include the prerelease value if applicable)
        Write-Host "";
        Write-Host "The .csproj version (doesn't affect anything; shows as `"Product Version`" when viewing the file properties in Windows Explorer) has no conventions and will be set to the same value as the updated module manifest version";
        Write-Host "The published .csproj version is $($PublishedCSProjFileData."Project"."PropertyGroup"."Version")";
        $UpdatedCSProjVersion = $UpdatedManifestVersion + ([string]::IsNullOrWhiteSpace($UpdatedManifestPrereleaseString) ? [string]::Empty : ($UpdatedManifestPrereleaseString.Contains("-") ? [string]::Empty : "-") + $UpdatedManifestPrereleaseString);
        Write-Host "The updated .csproj version will be $UpdatedCSProjVersion";

        #updated csproj assembly version
        Write-Host "";
        Write-Host "The .csproj assembly version (affects the runtime; does not get shown in Windows Explorer) should be set to the MAJOR version of the updated module manifest version";
        Write-Host "The published .csproj assembly version is $($PublishedCSProjFileData."Project"."PropertyGroup"."AssemblyVersion")";
        $UpdatedCSProjAssemblyVersion = $UpdatedManifestVersion.Substring(0, ($UpdatedManifestVersion).IndexOf(".")) + ".0.0";
        Write-Host "The updated .csproj assembly version will be $UpdatedCSProjAssemblyVersion";

        #updated file version
        Write-Host "";
        Write-Host "The .csproj file version (doesn't affect anything; shows as `"File Version`" when viewing the file properties in Windows Explorer) should be set to MAJOR.MINOR.BUILD.REVISION";
        Write-Host "The published .csproj file version is $($PublishedCSProjFileData."Project"."PropertyGroup"."FileVersion")";
        [int]$Revision = ($PublishedCSProjFileData."Project"."PropertyGroup"."FileVersion").Substring($PublishedCSProjFileData."Project"."PropertyGroup"."FileVersion".LastIndexOf(".") + 1);
        $Revision++;
        $UpdatedCSProjFileVersion = $UpdatedManifestVersion + "." + $Revision.ToString();
        Write-Host "The updated .csproj file version will be $UpdatedCSProjFileVersion";
        
        #summary
        Write-Host "";
        Write-Host "The new module manifest version will be $($UpdatedManifestVersion + ([string]::IsNullOrWhiteSpace($UpdatedManifestPrereleaseString) ? [string]::Empty : ($UpdatedManifestPrereleaseString.Contains("-") ? [string]::Empty : "-") + $UpdatedManifestPrereleaseString))";
        Write-Host "The new .csproj version will be $UpdatedCSProjVersion";
        Write-Host "The new .csproj assembly version will be $UpdatedCSProjAssemblyVersion";
        Write-Host "The new .csproj file version will be $UpdatedCSProjFileVersion";

        Write-Host "";
        $ValuesConfirmed = (Read-Host "Proceed with publishing? (`"y`" to proceed, any other key to re-enter values)");
    }
    Until ($ValuesConfirmed -eq "y");

    #empty out the publish directory
    foreach ($File in (Get-ChildItem -Path $PublishDirectory))
    {
        Remove-Item -Path $File."FullName" -Recurse -Force;
    }

    #publish the project to the publish directory
    Write-Host "";
    Write-Host "Building the project...";
    dotnet publish $ProjectDirectory --nologo --configuration "Release" --output $PublishDirectory /p:Version=$UpdatedCSProjVersion /p:AssemblyVersion=$UpdatedCSProjAssemblyVersion /p:FileVersion=$UpdatedCSProjFileVersion;
    Write-Host "Build completed";

    #copy the project manifest to the publish directory (copy the project's manifest rather than the published one since the project one may have updated information that needs to be present in the future published one)
    Copy-Item -Path (Join-Path -Path $ProjectDirectory -ChildPath $ManifestFileName) -Destination $PublishDirectory;

    #update the manifest in the publish directory (note that the module manifest update process ensures all required assemblies are present)
    Write-Host "";
    Write-Host "Updating the module manifest...";
    Update-ModuleManifest -Path (Join-Path -Path $PublishDirectory -ChildPath $ManifestFileName) -ModuleVersion $UpdatedManifestVersion -Prerelease $UpdatedManifestPrereleaseString;

    #copy the project .csproj file to the Publish directory (copy the project's .csproj file rather than the published one since the project one may have updated information that needs to be present in the future published one)
    Copy-Item -Path (Join-Path -Path $ProjectDirectory -ChildPath $CSProjFileName) -Destination $PublishDirectory;

    #update the .csproj file
    Write-Host "Updating the .csproj file...";
    $UpdatedCSProjFileData = [xml](Get-Content -Raw -Path (Join-Path -Path $ProjectDirectory -ChildPath $CSProjFileName));
    $UpdatedCSProjFileData."Project"."PropertyGroup"."Version" = $UpdatedCSProjVersion;
    $UpdatedCSProjFileData."Project"."PropertyGroup"."AssemblyVersion" = $UpdatedCSProjAssemblyVersion;
    $UpdatedCSProjFileData."Project"."PropertyGroup"."FileVersion" = $UpdatedCSProjFileVersion;
    $UpdatedCSProjFileData.Save("$(Join-Path -Path $PublishDirectory -ChildPath $CSProjFileName)");

    #publish the module
    #Write-Host "Publishing to the PSGallery...";
    #Publish-Module -Path $PublishDirectory -Repository PSGallery -NuGetApiKey $env:PowerShellGalleryAPIKey;

    #overwrite the project's manifest with the updated one from the publish directory (Copy-Item's default behavior is to overwrite files if they already exist in the destination)
    #Copy-Item -Path (Join-Path $PublishDirectory -ChildPath $ManifestFileName) -Destination $ProjectDirectory;

    #overwrite the project's .csproj file with the updated one from the publish directory (Copy-Item's default behavior is to overwrite files if they already exist in the destination)
    #Copy-Item -Path (Join-Path -Path $PublishDirectory -ChildPath $CSProjFileName) -Destination $ProjectDirectory;

    Write-Host "";
    Write-Host "Success";

}
Catch {
    Throw $Error[0];
}
Finally {
    #clean up published module download
    foreach ($File in (Get-ChildItem -Path $PublishedModuleDownloadDirectory))
    {
        Remove-Item -Path $File."FullName" -Recurse -Force;
    }

    #reset the progress preference variable
    $ProgressPreference = $DefaultProgressPreferenceValue;
}
