[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "sharepoint-list-schema-config.json"),
    [string]$ClientId = $env:ENTRAID_APP_ID,
    [string]$TemplatePath = (Join-Path $PSScriptRoot "sharepoint-list-import-template.xml"),
    [string]$InventoryPath = (Join-Path $PSScriptRoot "sharepoint-list-export-inventory.csv"),
    [switch]$IncludeHidden,
    [switch]$ForceAuthentication,
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$sourceSiteUrl = [string]$config.SourceSiteUrl
$excludedListUrls = [string[]]@($config.ExcludedListUrls)
$includedDocumentLibraryTitles = [string[]]@($config.IncludedDocumentLibraryTitles)

if ([string]::IsNullOrWhiteSpace($sourceSiteUrl)) {
    throw "Set SourceSiteUrl in $ConfigPath."
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw "Provide -ClientId or set ENTRAID_APP_ID to an Entra app ID configured for PnP interactive login."
}

Import-Module PnP.PowerShell -MinimumVersion 3.1.0

$connection = Connect-PnPOnline `
    -Url $sourceSiteUrl `
    -Interactive `
    -ClientId $ClientId `
    -ForceAuthentication:$ForceAuthentication `
    -ReturnConnection

$lists = @(Get-PnPList -Includes RootFolder -Connection $connection)
$inventory = @(
    foreach ($list in $lists) {
        $listUrl = [Uri]::UnescapeDataString($list.RootFolder.ServerRelativeUrl).TrimEnd("/")
        $isExplicitlyExcluded = $excludedListUrls -contains $listUrl
        $isExcludedDocumentLibrary =
            $list.BaseType -eq [Microsoft.SharePoint.Client.BaseType]::DocumentLibrary -and
            $list.Title -notin $includedDocumentLibraryTitles
        $isIncluded =
            -not $isExplicitlyExcluded -and
            -not $isExcludedDocumentLibrary -and
            ($IncludeHidden -or -not $list.Hidden)
        $reason = if ($isExplicitlyExcluded) {
            "Explicit exclusion"
        }
        elseif ($isExcludedDocumentLibrary) {
            "Document library not selected in IncludedDocumentLibraryTitles"
        }
        elseif ($list.Hidden -and -not $IncludeHidden) {
            "Hidden/system list"
        }
        else {
            "Included"
        }

        [pscustomobject]@{
            Title = $list.Title
            Url = $listUrl
            BaseTemplate = $list.BaseTemplate
            BaseType = $list.BaseType
            Hidden = $list.Hidden
            ItemCount = $list.ItemCount
            Included = $isIncluded
            Reason = $reason
        }
    }
)

$includedTitles = [string[]]@(
    $inventory |
        Where-Object Included |
        Sort-Object Title |
        Select-Object -ExpandProperty Title
)

if ($includedTitles.Count -eq 0) {
    throw "No lists were selected for export."
}

foreach ($path in @($TemplatePath, $InventoryPath)) {
    $parent = Split-Path -Parent $path
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    if ((Test-Path -LiteralPath $path) -and -not $Force) {
        throw "Output already exists: $path. Use -Force to replace it."
    }
}

$inventory |
    Sort-Object @{ Expression = "Included"; Descending = $true }, Title |
    Export-Csv -LiteralPath $InventoryPath -NoTypeInformation -Encoding utf8

Get-PnPSiteTemplate `
    -Out $TemplatePath `
    -Handlers Lists `
    -ListsToExtract $includedTitles `
    -Connection $connection `
    -Force:$Force

Write-Host "Exported $($includedTitles.Count) list schemas to $TemplatePath"
Write-Host "Wrote the complete included/excluded inventory to $InventoryPath"
