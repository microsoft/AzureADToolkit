<#
.SYNOPSIS
    Lists and categorizes risk for delegated permissions (OAuth2PermissionGrants) and application permissions (AppRoleAssignments).
.DESCRIPTION
    Long description
.EXAMPLE
    Build-AzureADAppConsentGrantReport -ReportOutputType ExcelWorkbook -ExcelWorkbookPath C:\temp\report.xlsx
    Output a report in Excel format
.EXAMPLE
    Build-AzureADAppConsentGrantReport -ReportOutputType ExcelWorkbook -ExcelWorkbookPath C:\temp\report.xlsx -PermissionsTableCsvPath .\table.csv
    Output a report in Excel format and specify a local path for a customized CSV containing consent risk categorizations
.INPUTS
    Inputs to this cmdlet (if any)
.OUTPUTS
    Output from this cmdlet (if any)
.NOTES
    General notes
.COMPONENT
    The component this cmdlet belongs to
.ROLE
    The role this cmdlet belongs to
.FUNCTIONALITY
    The functionality that best describes this cmdlet
#>
function Build-AzureADAppConsentGrantReport {
    [CmdletBinding(DefaultParameterSetName = 'Download Permissions Table Data',
        SupportsShouldProcess = $true,
        PositionalBinding = $false,
        HelpUri = 'http://www.microsoft.com/',
        ConfirmImpact = 'Medium')]
    [Alias()]
    [OutputType([String])]
    Param (

        # Output type for the report.
        [ValidateSet("ExcelWorkbook", "PowerShellObjects")]
        [string]
        $ReportOutputType = "ExcelWorkbook",
        # Output file location for Excel Workbook
        [Parameter(ParameterSetName = 'Excel Workbook Output')]
        [Parameter(Mandatory = $false)]
        [string]
        $ExcelWorkbookPath,
        # Path to CSV file for Permissions Table
        [string]
        $PermissionsTableCsvPath
    )

    begin {

        function checkMSGraphConnection {


        $apiPermissionScopes = @("Application.Read.All", "Application.ReadWrite.All")
 
        if ($null -eq (Get-MgContext) -or $null -eq (Get-MgContext).Scopes) {
            Write-Error "Please Connect to MS Graph API with the Connect-mgGraph cmdlet from the Microsoft.Graph.Authentication module first before calling functions! Application.Read.All is the recommended scope. Sign in with 'Connect-MgGraph -Scopes Application.Read.All'" -ErrorAction Stop
        }
        else {
            # Check that the current mgContext has the needed permissions for the call
            $authorized = $false
            foreach ($checkPermission in $apiPermissionScopes) {
                if ((Get-MgContext).Scopes.Contains($checkPermission)) {
                    $authorized = $true
                }
            }
            if ($true -ne $authorized) {
                Write-Error "Current MS Graph Context does not contain the proper scopes required to call the Application/Service Principal API.  Please ensure you are connecting with an identity that has this permission scope (Application.Read.All)!" -ErrorAction Stop
            }


        }



        }

        function GenerateExcelReport {
            param (
                $evaluatedData,
                $Path
            )

            # Delete the existing output file if it already exists
            $OutputFileExists = Test-Path $Path
            if ($OutputFileExists -eq $true) {
                Get-ChildItem $Path | Remove-Item -Force
            }

            $count = 0
            $highriskobjects = $evaluatedData | Where-Object { $_.Risk -eq "High" }
            $highriskobjects | ForEach-Object {
                $userAssignmentRequired = @()
                $userAssignments = @()
                $userAssignmentsCount = @()
                $userAssignmentRequired = Get-MgServicePrincipal -ServicePrincipalId $_.ClientObjectId

                if ($userAssignmentRequired.AppRoleAssignmentRequired -eq $true) {
                    $userAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $_.ClientObjectId -All:$true
                    $userAssignmentsCount = $userAssignments.count
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name UsersAssignedCount -Value $userAssignmentsCount
                }
                elseif ($userAssignmentRequired.AppRoleAssignmentRequired -eq $false) {
                    $userAssignmentsCount = "AllUsers"
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name UsersAssignedCount -Value $userAssignmentsCount
                }

                $count++
                Write-Progress -activity "Counting users assigned to high risk apps . . ." -status "Apps Counted: $count of $($highriskobjects.Count)" -percentComplete (($count / $highriskobjects.Count) * 100)
            }
            $highriskusers = $highriskobjects | Where-Object { $null -ne $_.PrincipalObjectId } | Select-Object PrincipalDisplayName, Risk | Sort-Object PrincipalDisplayName -Unique
            $highriskapps = $highriskobjects | Select-Object ClientDisplayName, Risk, UsersAssignedCount, MicrosoftRegisteredClientApp | Sort-Object ClientDisplayName -Unique | Sort-Object UsersAssignedCount -Descending

            # Pivot table by user
            $pt = New-PivotTableDefinition -SourceWorkSheet ConsentGrantData `
                -PivotTableName "PermissionsByUser" `
                -PivotFilter RiskFilter, PermissionFilter, ResourceDisplayNameFilter, ConsentTypeFilter, ClientDisplayName, MicrosoftRegisteredClientApp `
                -PivotRows PrincipalDisplayName `
                -PivotColumns Risk, PermissionType `
                -PivotData @{Permission = 'Count' } `
                -IncludePivotChart `
                -ChartType ColumnStacked `
                -ChartHeight 800 `
                -ChartWidth 1200 `
                -ChartRow 4 `
                -ChartColumn 14

            # Pivot table by resource
            $pt += New-PivotTableDefinition -SourceWorkSheet ConsentGrantData `
                -PivotTableName "PermissionsByResource" `
                -PivotFilter RiskFilter, ResourceDisplayNameFilter, ConsentTypeFilter, PrincipalDisplayName, MicrosoftRegisteredClientApp `
                -PivotRows ResourceDisplayName, PermissionFilter `
                -PivotColumns Risk, PermissionType `
                -PivotData @{Permission = 'Count' } `
                -IncludePivotChart `
                -ChartType ColumnStacked `
                -ChartHeight 800 `
                -ChartWidth 1200 `
                -ChartRow 4 `
                -ChartColumn 14

            # Pivot table by risk rating
            $pt += New-PivotTableDefinition -SourceWorkSheet ConsentGrantData `
                -PivotTableName "PermissionsByRiskRating" `
                -PivotFilter RiskFilter, PermissionFilter, ResourceDisplayNameFilter, ConsentTypeFilter, PrincipalDisplayName, MicrosoftRegisteredClientApp `
                -PivotRows Risk, ResourceDisplayName `
                -PivotColumns PermissionType `
                -PivotData @{Permission = 'Count' } `
                -IncludePivotChart `
                -ChartType ColumnStacked `
                -ChartHeight 800 `
                -ChartWidth 1200 `
                -ChartRow 4 `
                -ChartColumn 5

            $excel = $data | Export-Excel -Path $Path -WorksheetName ConsentGrantData `
                -PivotTableDefinition $pt `
                -AutoSize `
                -Activate `
                -HideSheet "None" `
                -UnHideSheet "PermissionsByRiskRating" `
                -PassThru

            # Create temporary Excel file and add High Risk Users sheet
            $xlTempFile = "$env:TEMP\ImportExcelTempFile.xlsx"
            Remove-Item $xlTempFile -ErrorAction Ignore
            $exceltemp = $highriskusers | Export-Excel $xlTempFile -PassThru
            Add-Worksheet -ExcelPackage $excel -WorksheetName HighRiskUsers -CopySource $exceltemp.Workbook.Worksheets["Sheet1"]

            # Create temporary Excel file and add High Risk Apps sheet
            $xlTempFile = "$env:TEMP\ImportExcelTempFile.xlsx"
            Remove-Item $xlTempFile -ErrorAction Ignore
            $exceltemp = $highriskapps | Export-Excel $xlTempFile -PassThru
            Add-Worksheet -ExcelPackage $excel -WorksheetName HighRiskApps -CopySource $exceltemp.Workbook.Worksheets["Sheet1"] -Activate

            $sheet = $excel.Workbook.Worksheets["ConsentGrantData"]
            Add-ConditionalFormatting -Worksheet $sheet -Range "A1:N1048576" -RuleType Equal -ConditionValue "High" -ForeGroundColor White -BackgroundColor Red -Bold -Underline
            Add-ConditionalFormatting -Worksheet $sheet -Range "A1:N1048576" -RuleType Equal -ConditionValue "Medium" -ForeGroundColor Black -BackgroundColor Orange -Bold -Underline
            Add-ConditionalFormatting -Worksheet $sheet -Range "A1:N1048576" -RuleType Equal -ConditionValue "Low" -ForeGroundColor Black -BackgroundColor Yellow -Bold -Underline

            $sheet = $excel.Workbook.Worksheets["HighRiskUsers"]
            Add-ConditionalFormatting -Worksheet $sheet -Range "B1:B1048576" -RuleType Equal -ConditionValue "High" -ForeGroundColor White -BackgroundColor Red -Bold -Underline
            Set-ExcelRange -Worksheet $sheet -Range A1:C1048576 -AutoSize

            $sheet = $excel.Workbook.Worksheets["HighRiskApps"]
            Add-ConditionalFormatting -Worksheet $sheet -Range "B1:B1048576" -RuleType Equal -ConditionValue "High" -ForeGroundColor White -BackgroundColor Red -Bold -Underline
            Set-ExcelRange -Worksheet $sheet -Range A1:C1048576 -AutoSize

            Export-Excel -ExcelPackage $excel|Out-Null
            Write-Verbose ("Excel workbook {0}" -f $ExcelWorkbookPath)
        }

        function Get-MSCloudIdConsentGrantList {
            [CmdletBinding()]
            param(
                [int] $PrecacheSize = 999
            )
            # An in-memory cache of objects by {object ID} andy by {object class, object ID}
            $script:ObjectByObjectId = @{}
            $script:ObjectByObjectClassId = @{}
            $script:KnownMSTenantIds = @("f8cdef31-a31e-4b4a-93e4-5f571e91255a", "72f988bf-86f1-41af-91ab-2d7cd011db47")

            # Function to add an object to the cache
            function CacheObject($Object) {
                if ($Object) {
                    if (-not $script:ObjectByObjectClassId.ContainsKey($Object.GetType().name)) {
                        $script:ObjectByObjectClassId[$Object.GetType().name] = @{}
                    }
                    $script:ObjectByObjectClassId[$Object.GetType().name][$Object.Id] = $Object
                    $script:ObjectByObjectId[$Object.Id] = $Object
                }
            }

            # Function to retrieve an object from the cache (if it's there), or from Azure AD (if not).
            function GetObjectByObjectId($ObjectId) {
                if (-not $script:ObjectByObjectId.ContainsKey($ObjectId)) {
                    Write-Verbose ("Querying Azure AD for object '{0}'" -f $ObjectId)
                    try {
                        $object = Get-AzureADObjectByObjectId -ObjectId $ObjectId
                        CacheObject -Object $object
                    }
                    catch {
                        Write-Verbose "Object not found."
                    }
                }
                return $script:ObjectByObjectId[$ObjectId]
            }

            # Get all ServicePrincipal objects and add to the cache
            Write-Verbose "Retrieving ServicePrincipal objects..."
            $servicePrincipals = Get-MgServicePrincipal -All:$true
            $Oauth2PermGrants = @()

            $count = 0
            foreach ($sp in $servicePrincipals) {
                CacheObject -Object $sp
                $spPermGrants = Get-MgServicePrincipalOauth2PermissionGrant -ServicePrincipalId $sp.Id -All:$true
                $Oauth2PermGrants += $spPermGrants
                $count++
                Write-Progress -activity "Caching Objects from Azure AD . . ." -status "Cached: $count of $($servicePrincipals.Count)" -percentComplete (($count / $servicePrincipals.Count) * 100)
            }

            # Get one page of User objects and add to the cache
            Write-Verbose "Retrieving User objects..."
            Get-MgUser -Top $PrecacheSize | ForEach-Object { CacheObject -Object $_ }

            # Get all existing OAuth2 permission grants, get the client, resource and scope details
            Write-Progress -Activity "Processing Delegated Permission Grants..."
            foreach ($grant in $Oauth2PermGrants) {
                if ($grant.Scope) {
                    $grant.Scope.Split(" ") | Where-Object { $_ } | ForEach-Object {
                        $scope = $_
                        $client = GetObjectByObjectId -ObjectId $grant.ClientId

                        # Determine if the object comes from the Microsoft Services tenant, and flag it if true
                        $MicrosoftRegisteredClientApp = @()
                        if ($client.AppOwnerTenantId -in $script:KnownMSTenantIds) {
                            $MicrosoftRegisteredClientApp = $true
                        }
                        else {
                            $MicrosoftRegisteredClientApp = $false
                        }

                        $resource = GetObjectByObjectId -ObjectId $grant.ResourceId
                        $principalDisplayName = ""
                        if ($grant.PrincipalId) {
                            $principal = GetObjectByObjectId -ObjectId $grant.PrincipalId
                            $principalDisplayName = $principal.DisplayName
                        }

                        if ($grant.ConsentType -eq "AllPrincipals") {
                            $simplifiedgranttype = "Delegated-AllPrincipals"
                        }
                        elseif ($grant.ConsentType -eq "Principal") {
                            $simplifiedgranttype = "Delegated-Principal"
                        }

                        New-Object PSObject -Property ([ordered]@{
                                "PermissionType"               = $simplifiedgranttype
                                "ConsentTypeFilter"            = $simplifiedgranttype
                                "ClientObjectId"               = $grant.ClientId
                                "ClientDisplayName"            = $client.DisplayName
                                "ResourceObjectId"             = $grant.ResourceId
                                "ResourceObjectIdFilter"       = $grant.ResourceId
                                "ResourceDisplayName"          = $resource.DisplayName
                                "ResourceDisplayNameFilter"    = $resource.DisplayName
                                "Permission"                   = $scope
                                "PermissionFilter"             = $scope
                                "PrincipalObjectId"            = $grant.PrincipalId
                                "PrincipalDisplayName"         = $principalDisplayName
                                "MicrosoftRegisteredClientApp" = $MicrosoftRegisteredClientApp
                            })
                    }
                }
            }

            # Iterate over all ServicePrincipal objects and get app permissions
            Write-Progress -Activity "Processing Application Permission Grants..."
            $script:ObjectByObjectClassId['MicrosoftGraphServicePrincipal'].GetEnumerator() | ForEach-Object {
                $sp = $_.Value

                Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id  -All:$true `
                | Where-Object { $_.PrincipalType -eq "ServicePrincipal" } | ForEach-Object {
                    $assignment = $_

                    $client = GetObjectByObjectId -ObjectId $assignment.PrincipalId

                    # Determine if the object comes from the Microsoft Services tenant, and flag it if true
                    $MicrosoftRegisteredClientApp = @()
                    if ($client.AppOwnerTenantId -in $script:KnownMSTenantIds) {
                        $MicrosoftRegisteredClientApp = $true
                    }
                    else {
                        $MicrosoftRegisteredClientApp = $false
                    }

                    $resource = GetObjectByObjectId -ObjectId $assignment.ResourceId
                    $appRole = $resource.AppRoles | Where-Object { $_.Id -eq $assignment.Id }

                    New-Object PSObject -Property ([ordered]@{
                            "PermissionType"               = "Application"
                            "ClientObjectId"               = $assignment.PrincipalId
                            "ClientDisplayName"            = $client.DisplayName
                            "ResourceObjectId"             = $assignment.ResourceId
                            "ResourceObjectIdFilter"       = $grant.ResourceId
                            "ResourceDisplayName"          = $resource.DisplayName
                            "ResourceDisplayNameFilter"    = $resource.DisplayName
                            "Permission"                   = $appRole.Value
                            "PermissionFilter"             = $appRole.Value
                            "ConsentTypeFilter"            = "Application"
                            "MicrosoftRegisteredClientApp" = $MicrosoftRegisteredClientApp
                        })
                }
            }


        }

        function EvaluateConsentGrants {
            param (
                $data
            )


            # Process Risk for gathered data
            $count = 0
            $data | ForEach-Object {

                try {
                   
                    $count++
                    Write-Progress -activity "Processing risk for each permission . . ." -status "Processed: $count of $($data.Count)" -percentComplete (($count / $data.Count) * 100)

                    $scope = $_.Permission
                    if ($_.PermissionType -eq "Delegated-AllPrincipals" -or "Delegated-Principal") {
                        $type = "Delegated"
                    }
                    elseif ($_.PermissionType -eq "Application") {
                        $type = "Application"
                    }

                    # Check permission table for an exact match
                    $risk = $null
                    $scoperoot = @()
                    Write-Debug ("Permission Scope: $Scope")

                    if ($scope -match '.') {
                        $scoperoot = $scope.Split(".")[0]
                    }
                    else {
                        $scoperoot = $scope
                    }

                    $test = ($permstable | Where-Object { $_.Permission -eq "$scoperoot" -and $_.Type -eq $type }).Risk # checking if there is a matching root in the CSV
                    $risk = ($permstable | Where-Object { $_.Permission -eq "$scope" -and $_.Type -eq $type }).Risk # Checking for an exact match

                    # Search for matching root level permission if there was no exact match
                    if (!$risk -and $test) {
                        # No exact match, but there is a root match
                        $risk = ($permstable | Where-Object { $_.Permission -eq "$scoperoot" -and $_.Type -eq $type }).Risk
                    }
                    elseif (!$risk -and !$test -and $type -eq "Application" -and $scope -like "*Write*") {
                        # Application permissions without exact or root matches with write scope
                        $risk = "High"
                    }
                    elseif (!$risk -and !$test -and $type -eq "Application" -and $scope -notlike "*Write*") {
                        # Application permissions without exact or root matches without write scope
                        $risk = "Medium"
                    }
                    elseif ($risk) {

                    }
                    else {
                        # Any permissions without a match, should be primarily Delegated permissions
                        $risk = "Unranked"
                    }

                    # Add the risk to the current object
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name Risk -Value $risk
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name RiskFilter -Value $risk
                    Add-Member -InputObject $_ -MemberType NoteProperty -Name Reason -Value $reason
                }
                catch {
                    Write-Error "Error Processing Permission for $_"
                }
                finally {
                    Write-Output $_
                }
            }

        }

        function loadPermisionsTable {
            param (
                $PermissionsTableCsvPath
            )

            if ($null -like $PermissionsTableCsvPath) {
                # Create hash table of permissions and permissions risk
                Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/microsoft/AzureADToolkit/main/src/data/aadconsentgrantpermissiontable.csv' -OutFile .\aadconsentgrantpermissiontable.csv
                $permstable = Import-Csv .\aadconsentgrantpermissiontable.csv -Delimiter ','
            }
            else {

                $permstable = Import-Csv $PermissionsTableCsvPath -Delimiter ','
            }

            Write-Output $permstable

        }

        checkMSGraphConnection

        if ("ExcelWorkbook" -eq $ReportOutputType) {
            # Determine if the ImportExcel module is installed since the parameter was included
            if ($null -eq (get-module -Name ImportExcel -ListAvailable)) {
                throw "The ImportExcel module is not installed.   This is used to export the results to an Excel worksheet.  Please install the ImportExcel Module before using this parameter or run without this parameter."
            }
        }
        
    }
    process {

        $permstable = loadPermisionsTable -PermissionsTableCsvPath $PermissionsTableCsvPath

        Write-Verbose "Retrieving Permission Grants from Azure AD Tenant..."
        $data = Get-MSCloudIdConsentGrantList
        if ($null -ne $data) {
            $evaluatedData = EvaluateConsentGrants -data $data
        }

    }
    end {

        if ("ExcelWorkbook" -eq $ReportOutputType) {
            
            Write-Verbose "Generating Excel Workbook at $ExcelWorkbookPath"
            GenerateExcelReport -evaluatedData $evaluatedData -Path $ExcelWorkbookPath
            
        }
        else {
            Write-Output $evaluatedData
        }

    }
}
