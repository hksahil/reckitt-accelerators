# Ensure Power BI module is installed
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force
}

# Ensure ImportExcel module is installed (optional, not used in this script)
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module -Name ImportExcel -Scope CurrentUser -Force
}

# Authenticate interactively
Login-PowerBI

# Output file path
$outputCsvPath = "$env:USERPROFILE\Desktop\PowerBI_DatasetDetails.csv"

# Initialize collection array
$dataCollection = @()

# 1️⃣ Fetch all workspaces
$workspacesUrl = "https://api.powerbi.com/v1.0/myorg/groups"
$workspacesResponse = Invoke-PowerBIRestMethod -Url $workspacesUrl -Method Get | ConvertFrom-Json

# 2️⃣ Target a specific workspace by name
$targetWorkspaceName = "YOUR_WORKSPACE_NAME_HERE"  # 🔁 Replace with your workspace name

$workspace = $workspacesResponse.value | Where-Object { $_.name -eq $targetWorkspaceName }

if ($null -eq $workspace) {
    Write-Host "❌ Workspace '$targetWorkspaceName' not found." -ForegroundColor Red
    Logout-PowerBI
    exit
}

$workspaceId = $workspace.id
$workspaceName = $workspace.name

Write-Host "`n✅ Workspace selected: $workspaceName ($workspaceId)`n"

# 3️⃣ Fetch all datasets in the selected workspace
$datasetsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets"
$datasetsResponse = Invoke-PowerBIRestMethod -Url $datasetsUrl -Method Get | ConvertFrom-Json

foreach ($dataset in $datasetsResponse.value) {
    $datasetId = $dataset.id

    Write-Host "🔄 Dataset: $($dataset.name) [$datasetId]"

    # Metadata
    $datasetUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId"
    $dataset = Invoke-PowerBIRestMethod -Url $datasetUrl -Method Get | ConvertFrom-Json

    $connectionType = switch ($dataset.targetStorageMode) {
        "Abf" { "Import" }
        "PremiumFiles" { "DirectQuery" }
        "Sas" { "DirectQuery" }
        "Mixed" { "Dual" }
        "Push" { "Push Dataset" }
        "Streaming" { "Streaming Dataset" }
        default { "Unknown" }
    }

    # Reports
    $reportsUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/reports"
    $reportsResponse = Invoke-PowerBIRestMethod -Url $reportsUrl -Method Get | ConvertFrom-Json
    $report = $reportsResponse.value | Where-Object { $_.datasetId -eq $datasetId }

    if ($null -eq $report) {
        Write-Host "⚠ No report found for dataset $($dataset.name)" -ForegroundColor Yellow
        continue
    }

    $reportId = $report.id
    $reportUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/reports/$reportId"
    $reportResponse = Invoke-PowerBIRestMethod -Url $reportUrl -Method Get | ConvertFrom-Json

    # Parameters
    $parametersUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/parameters"
    $parametersResponse = Invoke-PowerBIRestMethod -Url $parametersUrl -Method Get | ConvertFrom-Json
    $parameterNames = ($parametersResponse.value.name -join ", ")
    $parameterTypes = ($parametersResponse.value.type -join ", ")
    $parameterValues = ($parametersResponse.value.currentValue -join ", ")
    $isRequiredValues = ($parametersResponse.value.isRequired -join ", ")
    $suggestedValuesList = ($parametersResponse.value.suggestedValues -join "; ")

    # Refresh Schedule
    $refreshUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/refreshSchedule"
    $refreshResponse = Invoke-PowerBIRestMethod -Url $refreshUrl -Method Get | ConvertFrom-Json

    # Data Sources
    $datasourceUrl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/datasources"
    $datasourceResponse = Invoke-PowerBIRestMethod -Url $datasourceUrl -Method Get | ConvertFrom-Json

    $datasourceNames = ($datasourceResponse.value.name -join ", ")
    $datasourceTypes = ($datasourceResponse.value.datasourceType -join ", ")
    $datasourceIds = ($datasourceResponse.value.datasourceId -join ", ")
    $gatewayIds = ($datasourceResponse.value.gatewayId -join ", ")
    $connectionStrings = ($datasourceResponse.value.connectionDetails.connectionString -join ", ")
    $datasourcePaths = ($datasourceResponse.value.connectionDetails.path -join ", ")
    $datasourceKinds = ($datasourceResponse.value.connectionDetails.kind -join ", ")

    # Final record
    $datasetInfo = [PSCustomObject]@{
        WorkspaceName                   = $workspaceName
        WorkspaceID                     = $workspaceId
        Name                            = $dataset.name
        ID                              = $dataset.id
        Description                     = $dataset.description
        ConfiguredBy                    = $dataset.configuredBy
        CreatedDate                     = $dataset.createdDate
        TargetStorageMode               = $dataset.targetStorageMode
        ConnectionType                  = $connectionType
        IsRefreshable                   = if ($dataset.isRefreshable) { "Yes" } else { "No" }
        IsOnPremGatewayRequired         = if ($dataset.isOnPremGatewayRequired) { "Yes" } else { "No" }
        IsEffectiveIdentityRequired     = if ($dataset.isEffectiveIdentityRequired) { "Yes" } else { "No" }
        IsEffectiveIdentityRolesRequired = if ($dataset.isEffectiveIdentityRolesRequired) { "Yes" } else { "No" }
        IsInPlaceSharingEnabled         = if ($dataset.isInPlaceSharingEnabled) { "Yes" } else { "No" }
        AddRowsAPIEnabled               = if ($dataset.addRowsAPIEnabled) { "Yes" } else { "No" }
        CreateReportEmbedURL            = $dataset.createReportEmbedURL
        QnaEmbedURL                     = $dataset.qnaEmbedURL
        WebURL                          = $dataset.webUrl

        ParameterNames                  = $parameterNames
        ParameterTypes                  = $parameterTypes
        ParameterValues                 = $parameterValues
        IsRequired                      = $isRequiredValues
        SuggestedValues                 = $suggestedValuesList

        RefreshEnabled                  = if ($refreshResponse.enabled) { "Yes" } else { "No" }
        RefreshDays                     = ($refreshResponse.days -join ", ")
        RefreshTimes                    = ($refreshResponse.times -join ", ")
        RefreshTimeZone                 = $refreshResponse.localTimeZoneId
        RefreshNotifyOption             = $refreshResponse.notifyOption

        DatasourceNames                 = $datasourceNames
        DatasourceNamepath              = $datasourcePaths
        DatasourceNamekind              = $datasourceKinds
        DatasourceTypes                 = $datasourceTypes
        DatasourceIds                   = $datasourceIds
        GatewayIds                      = $gatewayIds
        ConnectionStrings               = $connectionStrings

        ReportID                        = $reportResponse.id
        ReportName                      = $reportResponse.name
        ReportDescription               = $reportResponse.description
        ReportEmbedUrl                  = $reportResponse.embedUrl
        ReportWebUrl                    = $reportResponse.webUrl
        ReportType                      = $reportResponse.reportType
        ReportAppId                     = $reportResponse.appId
        ReportDatasetId                 = $reportResponse.datasetId
        ReportIsOwnedByMe               = if ($reportResponse.isOwnedByMe) { "Yes" } else { "No" }
        ReportOriginalId                = $reportResponse.originalReportId
    }

    $dataCollection += $datasetInfo
}

# 4️⃣ Export to CSV
Write-Host "`n📁 Exporting details to CSV..."
$dataCollection | Export-Csv -Path $outputCsvPath -NoTypeInformation
Write-Host "✅ Export complete: $outputCsvPath"

# 5️⃣ Optional logout
Logout-PowerBI
