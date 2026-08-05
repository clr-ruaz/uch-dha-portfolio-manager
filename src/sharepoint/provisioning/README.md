# SharePoint list schema provisioning

This workflow uses PnP PowerShell to export list and selected document-library schemas from one SharePoint site and apply them to another site. It exports definitions such as fields, internal names, content type bindings, views, settings, and supported lookup relationships. List items and documents are not exported.

Export reads SharePoint schema and writes local files only. Import changes the configured target site, so always review the template and run with `-WhatIf` first.

## Files

- `sharepoint-list-schema-config.example.json` is the commit-safe configuration example.
- `sharepoint-list-schema-config.json` contains local settings for a specific source and target and is ignored by Git.
- `Export-SharePointListSchema.ps1` generates the XML template and CSV inventory.
- `Import-SharePointListSchema.ps1` applies the XML template to a target site.
- `sharepoint-list-import-template.xml` is the generated PnP template consumed by the import script and is ignored by Git.
- `sharepoint-list-export-inventory.csv` records every discovered list or library and why it was included or omitted and is ignored by Git.

## Configure a run

Copy the tracked example to the ignored local filename:

```powershell
Copy-Item `
        -LiteralPath ".\sharepoint-list-schema-config.example.json" `
        -Destination ".\sharepoint-list-schema-config.json"
```

Then edit `sharepoint-list-schema-config.json` before exporting or importing:

```json
{
        "_note": "Replace all placeholder values before running the export or import script.",
    "SourceSiteUrl": "https://contoso.sharepoint.com/sites/source",
    "TargetSiteUrl": "https://contoso.sharepoint.com/sites/target",
    "ExcludedListUrls": [
        "/sites/source/Lists/Archive"
    ],
    "IncludedDocumentLibraryTitles": [
        "Project Files"
    ]
}
```

Configuration fields:

- `SourceSiteUrl` is the site whose list schemas are exported.
- `TargetSiteUrl` is the existing site to which the template may be applied. Leave it empty until ready to test an import.
- `ExcludedListUrls` contains exact, decoded server-relative URLs of lists or libraries to omit. Use values such as `/sites/source/Lists/Archive`, not a browser URL ending in `AllItems.aspx`.
- `IncludedDocumentLibraryTitles` is the allowlist of document-library display titles to export. Other document libraries are omitted. Use an empty array to export no document libraries.

The export operates on the single web identified by `SourceSiteUrl`; it does not traverse subsites. `ExcludedListUrls` therefore excludes lists and libraries within that source site, rather than excluding separate SharePoint sites.

Hidden SharePoint system lists are omitted by default. The optional `-IncludeHidden` switch includes hidden lists, but document libraries still must appear in `IncludedDocumentLibraryTitles`.

## Authentication

PnP.PowerShell 3.1 or later requires an Entra application registration for interactive login. Set its application ID in the current PowerShell session:

```powershell
$env:ENTRAID_APP_ID = "<application-client-id>"
```

The application and signed-in account need permission to read schema from the source site. Import additionally requires permission to manage lists on the target site. Do not place passwords, MFA codes, client secrets, or other credentials in commands or source control.

## Export

From this directory, generate the template and inventory:

```powershell
./Export-SharePointListSchema.ps1
```

Replace existing generated files intentionally:

```powershell
./Export-SharePointListSchema.ps1 -Force
```

Force a fresh account-selection prompt when PnP has cached the wrong account:

```powershell
./Export-SharePointListSchema.ps1 -ForceAuthentication -Force
```

To use a configuration file stored elsewhere:

```powershell
./Export-SharePointListSchema.ps1 `
        -ConfigPath "C:\provisioning\customer-site.json" `
        -Force
```

Review `sharepoint-list-export-inventory.csv` and `sharepoint-list-import-template.xml` after every export. The inventory is the definitive record of what was selected, and generated templates can contain source-specific settings that deserve review before import.

## Import

The target site must already exist. Set `TargetSiteUrl` in the JSON, then preview the operation without applying the template:

```powershell
./Import-SharePointListSchema.ps1 -WhatIf
```

The target can also be supplied as a one-time command-line override:

```powershell
./Import-SharePointListSchema.ps1 `
        -TargetSiteUrl "https://contoso.sharepoint.com/sites/test-target" `
        -WhatIf
```

After validating against a non-production target, apply the template:

```powershell
./Import-SharePointListSchema.ps1
```

PnP templates can be applied repeatedly to synchronize supported schema changes. Retain a backup before updating an existing site and use a separate test site before any production import.

## Custom output paths

Both scripts accept path overrides. For example:

```powershell
./Export-SharePointListSchema.ps1 `
        -TemplatePath ".\artifacts\lists.xml" `
        -InventoryPath ".\artifacts\inventory.csv" `
        -Force

./Import-SharePointListSchema.ps1 `
        -TemplatePath ".\artifacts\lists.xml" `
        -WhatIf
```