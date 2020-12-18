

####################################
# Author:       Eric Austin
# Create date:  November 2020
# Description:  Update libraries used in the Send-MailKitMessage module (meant to be run manually)
####################################

using namespace System.Collections.Generic

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation=Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv"

    #script elements
    $LibraryDirectory=".\..\Libraries"
    [List[Object]]$LibraryList=[List[Object]]::new()
    $LibraryDownloadDirectory=(Join-Path -Path $CurrentDirectory -ChildPath "Library download")

    #--------------#

    Clear-Host

    Write-Host ""

    #ensure "Library download" directory exists
    Write-Host "Ensuring `"Library download`" directory exists..."
    if ( -not (Test-Path $LibraryDownloadDirectory))
    {
        New-Item -ItemType Directory -Path $LibraryDownloadDirectory | Out-Null
    }
    
    #reference Get-LibraryVersions function
    . ".\Get-LibraryVersions.ps1"

    #get library versions
    Write-Host "Getting libraries and versions..."
    $LibraryList=Get-LibraryVersions

    #download packages based on user selection
    foreach ($Library in (($LibraryList | Out-GridView -Title "Choose library(ies) to update" -PassThru)))
    {
        
        #ensure updated version is available
        if ($Library."Send-MailKitMessageVersion" -eq $Library."NuGetVersion")
        {
            Write-Host "$($Library."LibraryName"): no updated version available"
        }
        else {

            #ensure "Library download" directory is empty
            Write-Host "Emptying `"Library download`" directory..."
            Get-ChildItem -Path $LibraryDownloadDirectory | ForEach-Object {
               Remove-Item -Path $_."FullName" -Recurse
            }

            #download package
            Write-Host "Downloading `"$($Library."LibraryName")`"..."
            Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/$($Library."LibraryName")/$($Library."NuGetVersion")" -OutFile (Join-Path -Path $LibraryDownloadDirectory -ChildPath "$($Library."LibraryName").zip")

            #unzip package
            Write-Host "Extracting `"$($Library."LibraryName")`"..."
            Expand-Archive -Path (Join-Path -Path $LibraryDownloadDirectory -ChildPath "$($Library."LibraryName").zip") -DestinationPath (Join-Path -Path $LibraryDownloadDirectory -ChildPath $Library."LibraryName")

            #copy nuspec manifest and dll
            Write-Host "Copying nuspec manifest and dll..."
            Copy-Item -Path (Join-Path -Path $LibraryDownloadDirectory -ChildPath $Library."LibraryName" -AdditionalChildPath "$($Library."LibraryName").nuspec") -Destination $LibraryDirectory  #Copy-Item overwrites by default
            Copy-Item -Path (Join-Path -Path $LibraryDownloadDirectory -ChildPath $Library."LibraryName" -AdditionalChildPath "lib", "netstandard2.0", "*.dll") -Destination $LibraryDirectory  #Copy-Item overwrites by default (use a wildcard because Portable.BouncyCastle has a different dll name than the library, and there would only be one dll in the directory)
            
        }

    }

    #Get library versions
    Write-Host "Getting libraries and versions..."
    Get-LibraryVersions | Format-Table

    Write-Host "If all looks as it should the updates can be committed to source control and an updated module can be published to the PowerShell Gallery"
    
}

Catch {

    Write-Host $Error[0] -ForegroundColor "Red"

    #error log
    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{ "Date"=(Get-Date).ToString(); "Script"="Update libraries.ps1"; "ErrorMessage"=$Error[0].ToString() }
    $ErrorData | Select-Object "Date", "Script", "ErrorMessage" | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation

    #return value
    Exit 1
    
}

Finally {



}