[CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
param(
    [string]$TargetSiteUrl,
    [string]$ConfigPath = (Join-Path $PSScriptRoot "sharepoint-list-schema-config.json"),
    [string]$ClientId = $env:ENTRAID_APP_ID,
    [string]$TemplatePath = (Join-Path $PSScriptRoot "sharepoint-list-import-template.xml"),
    [switch]$ForceAuthentication
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Configuration file not found: $ConfigPath"
}

$config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace($TargetSiteUrl)) {
    $TargetSiteUrl = [string]$config.TargetSiteUrl
}

if ([string]::IsNullOrWhiteSpace($TargetSiteUrl)) {
    throw "Set TargetSiteUrl in $ConfigPath or provide -TargetSiteUrl."
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw "Provide -ClientId or set ENTRAID_APP_ID to an Entra app ID configured for PnP interactive login."
}

if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) {
    throw "Provisioning template not found: $TemplatePath"
}

Import-Module PnP.PowerShell -MinimumVersion 3.1.0

$connection = Connect-PnPOnline `
    -Url $TargetSiteUrl `
    -Interactive `
    -ClientId $ClientId `
    -ForceAuthentication:$ForceAuthentication `
    -ReturnConnection

if ($PSCmdlet.ShouldProcess($TargetSiteUrl, "Apply list schema from $TemplatePath")) {
    Invoke-PnPSiteTemplate `
        -Path $TemplatePath `
        -Connection $connection

    Write-Host "Applied the SharePoint list schema to $TargetSiteUrl"
}
