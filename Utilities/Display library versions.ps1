

####################################
# Author:       Eric Austin
# Create date:  December 2020
# Description:  Displays library version information (meant to be run manually)
####################################

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot #$PSScriptRoot is an empty string when not run from a script, and null coalescing doens't work with empty strings
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation=Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv"

    #--------------#

    Clear-Host

    Write-Host ""
    Write-Host "This script displays the current Send-MailKitMessage libraries and their versions, as well as the latest NuGet versions"

    #reference Get-LibraryVersions function
    . ".\Get-LibraryVersions.ps1"

    #get library versions
    Write-Host "Getting libraries and versions..."
    Get-LibraryVersions | Format-Table

}

Catch {

    Write-Host $Error[0] -ForegroundColor "Red"

    #error log
    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{ "Date"=(Get-Date).ToString(); "Script"="Display library versions.ps1"; "ErrorMessage"=$Error[0].ToString() }
    $ErrorData | Select-Object "Date", "Script", "ErrorMessage" | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation

    #return value
    Exit 1
    
}