

####################################
# Author:       Eric Austin
# Create date:  November 2020
# Description:  Creates GitHub issues for outdated libraries
####################################

using namespace System.Collections.Generic

Try {

    #common variables
    $CurrentDirectory=[string]::IsNullOrWhiteSpace($PSScriptRoot) ? (Get-Location).Path : $PSScriptRoot
    $ErrorActionPreference="Stop"
    $ErrorData=@()
    $ErrorLogLocation=Join-Path -Path $CurrentDirectory -ChildPath "ErrorLog.csv"

    #script elements
    [string]$IssueTitle=[string]::Empty
    [List[Object]]$LibraryList=[List[Object]]::new()
    [List[string]]$ExistingIssuesList=[List[string]]::new()

    #--------------#

    #reference Get-LibraryVersions function
    . ".\Get-LibraryVersions.ps1"

    #get library versions
    $LibraryList=(Get-LibraryVersions)

    #build authentication header
    $Token=([System.Convert]::ToBase64String(([char[]]"Username:$env:GitHubServiceAccountPAT")))
    $Header=@{Authorization = "Basic $Token"}

    #get existing open issues
    $ExistingIssuesList=(Invoke-RestMethod -Method "Get" -Uri "https://api.github.com/repos/austineric/Send-MailKitMessage/issues" | Where-Object { $_."state" -eq "open" }).title

    #open a GitHub issue for each library that has an updated version available
    $LibraryList | Where-Object { $_."UpdateIsAvailable" -eq "True" } | ForEach-Object {
        
        #issue title
        $IssueTitle="Update $($_."LibraryName") from $($_."Send-MailKitMessageVersion") to $($_."NuGetVersion")"
        
        #create an issue so long as one with the same title does not already exist
        if (-not ($ExistingIssuesList.Contains($IssueTitle)))
        {
            $Body=@{
                "title"=$IssueTitle
            } | ConvertTo-Json
            Invoke-RestMethod -Method "Post" -Uri "https://api.github.com/repos/austineric/Send-MailKitMessage/issues" -Headers $Header -Body $Body | Out-Null
        }
        
    }
    
}

Catch {

    #error log
    $ErrorData+=New-Object -TypeName PSCustomObject -Property @{ "Date"=(Get-Date).ToString(); "Script"="Create GitHub issues for outdates libraries.ps1"; "ErrorMessage"=$Error[0].ToString() }
    $ErrorData | Select-Object "Date", "Script", "ErrorMessage" | Export-Csv -Path $ErrorLogLocation -Append -NoTypeInformation
    
    #return value
    Exit 1
    
}