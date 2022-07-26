#### Functions, modules, and starting variables ####
$ModuleCheck = Get-InstalledModule | Where-Object {$_.name -eq "HuduAPI"}
if ($null -eq $ModuleCheck) {
    Install-Module HuduAPI -Scope CurrentUser
}
Import-Module HuduAPI
function New-HuduITGlueImportedArticle {
    param(
        [string]$FilePathToImport,
        [string]$OrganizationNameToImportTo,
        [string]$NewKBArticleName
    )

    #Match the organization by name
    $HuduCompany = Get-HuduCompanies | Where-Object {$_.name -match $OrganizationNameToImportTo}
    
    #Confirm only one result was returned
    if (($HuduCompany|Measure-Object).count -eq 1) {
        #Grab the content of the file to create
        $KBContent = Get-Content ("$($FilePathToImport.replace('/',''))\$($NewKBArticleName.replace('/','')).html").replace('`n','').replace('`r','').ToString()
        Write-Verbose "Import document $FilePathToImport for organization id $($HuduCompany.id)"

        #Post new article to Hudu
        $Article = New-HuduArticle -name $NewKBArticleName -content ($KBContent|Out-String) -company_id $HuduCompany.id

        #Confirm article was posted with result, if no result indicate failure and move to Failed status
        if ($Article) {
            Write-Verbose "Migration of KB Article completed. Moving record to 'Migrated' path"
            Move-Item -Path ($($FilePathToImport.replace('/',''))).replace('`n','').replace('`r','').ToString() -Destination "$($FilePathToImport.replace('/',''))\..\1) Migration Results\Migrated"
        }else{
            Write-Warning "Migration of KB Article failed. Moving record to 'Failed' path"
            Move-Item -Path ($($FilePathToImport.replace('/',''))).replace('`n','').replace('`r','').ToString() -Destination "$($FilePathToImport.replace('/',''))\..\1) Migration Results\Failed"
        }
    }else{
        Write-Warning "More than one Company was returned from Company name search: $($OrganizationNameToImportTo). Please address manually"
    }
}
function Start-ITGlueKbImport {
    <#
    .DESCRIPTION
    Start-ITGlueKbImport is a function designed to parse the CSV from an ITGlue full export.
    Using the passed in variable for the Root location of the documents, and combining parts of the CSV the script will gather the content and post it to Hudu using the HuduAPI.
    This function was created specifically for cases where Luke's migration script on mspp.io isn't feasible because (for example) API Access isn't available to ITGlue.
    Note that when a migration is done, the migrated folder of that KB Article will be moved into the "Migrated" path (located in the Export Root) or the "Failed" path if it failed.
    To try again, just move the folder back into the root, and ensure that record is in the main CSV being parsed.

    .PARAMETER MainCSVExportPath
    This parameter is used to specify the full path to the CSV that ITGlue provides along with the export. The CSV will contain the location of all the KB Articles, the KB titles, and the Organization name.

    .PARAMETER RootFolderExportPath
    This parameter is used to specify the location where the ITGlue Export was extracted to

    .PARAMETER HuduBaseUrl
    This function relies on the HuduAPI powershell module created by Luke (mspp.io). Specify your Hudu instance URL (without the trailing '/') here to use that module to connect to the Hudu API.

    .PARAMETER HuduAPIKey
    This function relies on the HuduAPI powershell module created by Luke (mspp.io). Specify your HuduAPIKey here to use that module to connect to the Hudu API.

    .EXAMPLE
    Start-ITGlueKbImport -MainCSVExportPath $MainCSVExportPath -RootFolderExportPath $RootFolderExportPath -SkipDocumentsWithAttachments -HuduBaseUrl $HuduBaseURL -HuduAPIKey $HuduApiKey
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$MainCSVExportPath,
        [string]$RootFolderExportPath,
        [string]$HuduBaseUrl,
        [string]$HuduAPIKey
    )
    New-HuduAPIKey -ApiKey $HuduApiKey
    New-HuduBaseURL -BaseURL $HuduBaseURL

    #Import the CSV for processing 
    $ImportDetails = Import-Csv -Path $MainCSVExportPath
    foreach ($record in $ImportDetails) {
        Write-Host "Processing import for $RootFolderExportPath\$($record.locator) $($record.name.replace('/', ''))" -ForegroundColor Cyan
        New-HuduITGlueImportedArticle -FilePathToImport "$RootFolderExportPath\$($record.locator.replace('"','')) $($record.name.replace('/', '').replace('"',''))" -OrganizationNameToImportTo "$($record.organization)" -NewKBArticleName "$($record.name.replace('"',''))" -RootFolderExportPath $RootFolderExportPath  
    }
}

$MainCSVExportPath = (Read-Host -Prompt "Enter the path to your CSV").replace('"','')
$RootFolderExportPath = (Read-Host -Prompt "Enter the path to your documents exported from IT Glue").replace('"','')
$HuduBaseUrl = Read-Host -Prompt "Enter your Hudu URL without the trailing space (/)"
$HuduAPIKey = Read-Host -Prompt "Enter your Hudu API key"

#### End of functions, modules, and starting variables ####

# Create the folder structure for successful migrations and failures
New-Item -Path "$($RootFolderExportPath)" -Name "1) Migration Results" -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path "$($RootFolderExportPath)\1) Migration Results" -Name Migrated -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path "$($RootFolderExportPath)\1) Migration Results\Migrated" -Name "1) WithAttachment" -ItemType Directory -ErrorAction SilentlyContinue
New-Item -Path "$($RootFolderExportPath)\1) Migration Results" -Name Failed -ItemType Directory -ErrorAction SilentlyContinue

# Begins migration
Start-ITGlueKbImport -MainCSVExportPath $MainCSVExportPath -RootFolderExportPath $RootFolderExportPath -HuduBaseUrl $HuduBaseUrl -HuduAPIKey $HuduAPIKey

# Create copy of MainCSV in Failed directory for easy re-uploading. Successfully imported documents are removed from this copy
Copy-Item $MainCSVExportPath -Destination "$($RootFolderExportPath)\1) Migration Results\Failed"

# Gets names of all failed imports for comparison
$FailedImports = Get-ChildItem "$($RootFolderExportPath)\1) Migration Results\Failed" | Select-Object -ExpandProperty Name

$FailedCSV = Get-ChildItem "$($RootFolderExportPath)\1) Migration Results\Failed" | Where-Object {$_.Extension -eq ".csv"} | Select-Object -ExpandProperty FullName

# Removes successful imports from the copied CSV leaving only a list of the failed items to try again
  foreach ($record in $ImportDetails){
    if ("$($record.locator) $($record.name.replace('/', ''))" -notin $FailedImports){
        Set-Content $FailedCSV -Value (Get-Content $FailedCSV | Select-String -Pattern "$($record.locator)" -NotMatch)
    }
}

# Moves successfully imported articles with subfolders (likely for attachments) to a separate folder for easy access to documents that need attachments manually uploaded
Get-ChildItem "$($RootFolderExportPath)\1) Migration Results\Migrated" | Where-Object {$_.name -ne "1) WithAttachment"} | ForEach-Object{
    if ((Get-ChildItem -Directory $_.fullname).Count -gt 0){
        Move-Item $_.fullname -Destination "$($RootFolderExportPath)\1) Migration Results\Migrated\1) WithAttachment"
    }
}
