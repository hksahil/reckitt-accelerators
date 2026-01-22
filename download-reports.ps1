Import-Module MicrosoftPowerBIMgmt
Connect-PowerBIServiceAccount

# Root export folder
$root = Join-Path $env:USERPROFILE "Desktop\PowerBI_Exports_$(Get-Date -Format yyyyMMdd_HHmmss)"
New-Item -ItemType Directory -Path $root -Force | Out-Null

function Sanitize-Name {
    param([string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $Name = $Name.Replace($c, '_') }
    return $Name.Trim()
}

##################FOR SINGLE WORKSPACE###################

#$targetWorkspaceName = "Finance - Prod"   # <-- change

#$workspaces = Get-PowerBIWorkspace -All |
#              Where-Object { $_.Name -eq $targetWorkspaceName }

#if (-not $workspaces) {
#    throw "Workspace not found or you don't have access: $targetWorkspaceName"
#}

##################FOR Multiple WORKSPACES###################
#$targetWorkspaceNames = @(
#  "Finance - Prod",
#  "Sales - Analytics",
#  "HR - Reporting"
#)

#$workspaces = Get-PowerBIWorkspace -All |
#             Where-Object { $targetWorkspaceNames -contains $_.Name }

##################FOR ALL WORKSPACES###################

# Get workspaces you have access to (user scope)
$workspaces = Get-PowerBIWorkspace -All

$inventory = New-Object System.Collections.Generic.List[object]
$exportLog = New-Object System.Collections.Generic.List[object]

foreach ($ws in $workspaces) {
    $wsNameSafe = Sanitize-Name $ws.Name
    $wsFolder = Join-Path $root $wsNameSafe
    New-Item -ItemType Directory -Path $wsFolder -Force | Out-Null

    Write-Host "Workspace: $($ws.Name)" -ForegroundColor Cyan

    try {
        $reports = Get-PowerBIReport -WorkspaceId $ws.Id

        foreach ($r in $reports) {
            # Inventory row
            $inventory.Add([pscustomobject]@{
                WorkspaceId   = $ws.Id
                WorkspaceName = $ws.Name
                ReportId      = $r.Id
                ReportName    = $r.Name
                ReportWebUrl  = $r.WebUrl
                DatasetId     = $r.DatasetId
                DatasetName   = $r.DatasetName
            })

            # Export attempt
            $reportNameSafe = Sanitize-Name $r.Name
            $outFile = Join-Path $wsFolder "$reportNameSafe.pbix"

            # Export-PowerBIReport requires file path NOT to exist.
            if (Test-Path $outFile) {
                Remove-Item $outFile -Force
            }

            try {
                Export-PowerBIReport -WorkspaceId $ws.Id -Id $r.Id -OutFile $outFile

                $exportLog.Add([pscustomobject]@{
                    WorkspaceName = $ws.Name
                    WorkspaceId   = $ws.Id
                    ReportName    = $r.Name
                    ReportId      = $r.Id
                    Status        = "Success"
                    OutputFile    = $outFile
                    Message       = ""
                })

                Write-Host "  Exported: $($r.Name)" -ForegroundColor Green
            }
            catch {
                $exportLog.Add([pscustomobject]@{
                    WorkspaceName = $ws.Name
                    WorkspaceId   = $ws.Id
                    ReportName    = $r.Name
                    ReportId      = $r.Id
                    Status        = "Failed"
                    OutputFile    = $outFile
                    Message       = $_.Exception.Message
                })

                Write-Warning "  Failed export: $($r.Name) -> $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Warning "Failed workspace $($ws.Name): $($_.Exception.Message)"
    }
}

# Save outputs
$inventoryFile = Join-Path $root "PowerBI_Report_Inventory.csv"
$exportFile    = Join-Path $root "PowerBI_Export_Results.csv"

$inventory | Export-Csv -Path $inventoryFile -NoTypeInformation -Encoding UTF8
$exportLog  | Export-Csv -Path $exportFile    -NoTypeInformation -Encoding UTF8

Write-Host "`nDONE" -ForegroundColor Yellow
Write-Host "Inventory: $inventoryFile"
Write-Host "Export log: $exportFile"
Write-Host "Export root: $root"
