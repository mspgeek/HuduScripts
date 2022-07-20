function New-HuduITGlueImportedArticle {
param(
[string]$FilePathToImport,
[string]$OrganizationNameToImportTo,
[string]$NewKBArticleName
)


    #Match the organization by name
    $HuduCompany = Get-HuduCompanies -name $OrganizationNameToImportTo

    #Confirm only one result was returned
    if (($HuduCompany|measure).count -eq 1) {
    
        #Grab the content of the file to create
        $KBContent = Get-Content "$FilePathToImport\$NewKBArticleName.html"
        Write-Verbose "Import document $FilePathToImport for organization id $($HuduCompany.id)"

        #Post new article to Hudu
        $Article = New-HuduArticle -name $NewKBArticleName -content ($KBContent|Out-String) -company_id $HuduCompany.id

        #Confirm article was posted with result, if no result indicate failure and move to Failed status
        if ($Article) {
            Write-Verbose "Migration of KB Article completed. Moving record to 'Migrated' path"
            Move-Item -Path $FilePathToImport -Destination "$FilePathToImport\..\migrated"
        }
        else {
            Write-Warning "Migration of KB Article failed. Moving record to 'Failed' path"
            Move-Item -Path $FilePathToImport -Destination "$FilePathToImport\..\failed"
            }
    }
    else {
        Write-Warning "More than one Company was returned from Company name search: $($OrganizationNameToImportTo). Please address manually"
        }
}

function Start-ITGlueKbImport {
param(
[string]$MainCSVExportPath,
[string]$RootFolderExportPath,
[boolean]$SkipDocumentsWithAttachments,
[string]$HuduBaseUrl,
[string]$HuduAPIKey
)

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


.PARAMETER SkipDocumentsWithAttachments
For exports that have attachments, these folders contain more than just the single HTML file of the KB. This parameter will skip those documents so that you can do all the "easy" ones first and work through the remaining ones.


.PARAMETER HuduBaseUrl
This function relies on the HuduAPI powershell module created by Luke (mspp.io). Specify your Hudu instance URL (without the trailing '/') here to use that module to connect to the Hudu API.

.PARAMETER HuduAPIKey
This function relies on the HuduAPI powershell module created by Luke (mspp.io). Specify your HuduAPIKey here to use that module to connect to the Hudu API.

.EXAMPLE
Start-ITGlueKbImport -MainCSVExportPath $MainCSVExportPath -RootFolderExportPath $RootFolderExportPath -SkipDocumentsWithAttachments -HuduBaseUrl $HuduBaseURL -HuduAPIKey $HuduApiKey
#>

Import-Module HuduAPI

New-HuduAPIKey -ApiKey $HuduApiKey
New-HuduBaseURL -BaseURL $HuduBaseURL

#Import the CSV for processing 
$ImportDetails = Import-Csv -Path $MainCSVExportPath

# Create the folder structure for migrations and failures
New-Item -Path "$($RootFolderExportPath)" -Name Migrated -ItemType Directory
New-Item -Path "$($RootFolderExportPath)" -Name Failed -ItemType Directory

foreach ($record in $ImportDetails) {

    # Check to see if there is more than one file in the folder for processing
    if ($null -ne $SkipDocumentsWithAttachments) {
        if (((Get-ChildItem -Path "$RootFolderExportPath\$($record.locator) $($record.name)").count) -eq 1) {
        Write-Verbose "There is only one file, processing import"
        New-HuduITGlueImportedArticle -FilePathToImport "$RootFolderExportPath\$($record.locator) $($record.name)" -OrganizationNameToImportTo "$($record.organization)" -NewKBArticleName "$($record.name)"
        }
        else {
        Write-Warning "More than one document found in the folder, and parameter was specified to skip attachments"
        }
    } 
    else {
        "Processing all documents even with attachments"
         New-HuduITGlueImportedArticle -FilePathToImport "$RootFolderExportPath\$($record.locator) $($record.name)" -OrganizationNameToImportTo "$($record.organization)" -NewKBArticleName "$($record.name)"
    }



    }

}
