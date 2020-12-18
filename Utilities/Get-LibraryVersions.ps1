

####################################
# Author:       Eric Austin
# Create date:  November 2020
# Description:  Centralized function for getting library version information 
####################################

function Get-LibraryVersions {

    try {

        #function elements
        $LibraryDirectory="..\.\Libraries"
        [System.Collections.Generic.List[Object]]$LibraryList=[System.Collections.Generic.List[Object]]::new()

        #get Send-MailKitMessage libraries and versions
        Get-ChildItem -Path $LibraryDirectory | Where-Object { $_."Extension" -eq ".nuspec" } | ForEach-Object {
            [xml]$NuspecFile=Get-Content -Raw -Path $_."FullName"
            $LibraryList.Add(
                [PSCustomObject]@{
                    "LibraryName"=$NuspecFile."package"."metadata"."id"
                    "Send-MailKitMessageVersion"=$NuspecFile."package"."metadata"."version"
                    "NuGetVersion"=[string]::Empty
                    "UpdateIsAvailable"=[string]::Empty
                }
            )
        }

        #ensure NuGet is registered as a package source
        if (-not (Get-PackageSource | Where-Object { $_.Location -eq "https://www.nuget.org/api/v2" }))
        {
            Register-PackageSource -Name "NuGet" -Location "https://www.nuget.org/api/v2" -ProviderName "NuGet"
        }

        #get NuGet library versions
        $LibraryList | ForEach-Object {
            $_."NuGetVersion"=(Find-Package -Name $_."LibraryName")."Version"
        }

        #set UpdateIsAvailable
        $LibraryList | ForEach-Object {
            if ($_."Send-MailKitMessageVersion" -ne $_."NuGetVersion")
            {
                $_."UpdateIsAvailable"="True"
            }
        }

        return $LibraryList

    }
    catch {
        Throw $Error[0]
    }

}
    