# UPM DHA Dashboard

## Overview

The UPM DHA Dashboard repository contains the Dallas Housing Authority portfolio-management solution and its supporting integrations. It is a monorepo: each delivery component has its own runtime, dependencies, and release process while sharing solution configuration and documentation.

## Repository structure

```text
UPM DHA Dashboard/
├── discovery/                         # Discovery artifacts
├── docs/                              # Solution-level documentation
├── src/
│   ├── azure-functions/               # Python HTTP API for document processing
│   ├── powerplatform/DallasHousingAuthority/
│   │   └── src/                       # Unpacked Power Platform solution
│   └── sharepoint/
│       ├── provisioning/              # PnP PowerShell schema export/import tooling
│       └── spfx/                      # DHA Portfolio Manager SPFx package
├── LICENSE
├── README.md
└── SECURITY.md
```

## Components

### Azure Functions API

Location: `src/azure-functions/`

Python Azure Functions v4 app using the decorator-based `FunctionApp` model. `function_app.py` provides function-key-protected HTTP endpoints for document and presentation operations:

| Route | Purpose |
| --- | --- |
| `DqsExtractZip` | Extract ZIP archive entries as Base64 payloads. |
| `DqsConvertCsvToXlsx` | Convert CSV content into XLSX, with optional headers and table formatting. |
| `DqsSplitPdf` | Split a PDF into one Base64-encoded file per page. |
| `DqsFillPpt` | Populate and insert PowerPoint slides from structured content. |
| `DqsStrToDoc` | Convert structured text content into a Word document. |
| `DqsExtractPpt` | Extract PowerPoint presentation content. |
| `DqsRegEx` | Perform bounded regular-expression operations. |
| `DqsReplaceTxtInPpt` | Replace text in a PowerPoint presentation. |

Key files: `function_app.py`, `host.json`, and `requirements.txt`. `local.settings.json` is local-only configuration and must not contain committed secrets.

Run locally:

```powershell
cd src/azure-functions
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
func start
```

### SharePoint Framework (SPFx)

Location: `src/sharepoint/spfx/`

The DHA Portfolio Manager SPFx project (`dha-portfolio-manager-spap`) supports a modern web part and a single-part app page experience. It uses SPFx `1.21.1` and React `17.0.1`; generated deployment packages are written to `sharepoint/solution/`.

```powershell
cd src/sharepoint/spfx
npm install
npm run build
npm run package-solution
# Local SharePoint workbench development
npm run serve
```

Upload the generated `.sppkg` file from `src/sharepoint/spfx/sharepoint/solution/` to the SharePoint App Catalog. See `src/sharepoint/spfx/README.md` for the component-specific details.

### Power Platform Solution

Location: `src/powerplatform/DallasHousingAuthority/`

`DallasHousingAuthority.cdsproj` is the Power Platform solution project. Its unpacked `src/` directory contains cloud workflow definitions, solution customizations, and environment-variable definitions for SharePoint, ResMan, Azure Functions, Azure OpenAI, notifications, and related integration settings.

### SharePoint Schema Provisioning

Location: `src/sharepoint/provisioning/`

PnP PowerShell scripts export SharePoint list and selected document-library schemas and import them into an existing target site. The tooling includes configuration examples, a generated PnP template, and an inventory report. Before applying changes, configure a local `sharepoint-list-schema-config.json` file and run the import with `-WhatIf`.

```powershell
cd src/sharepoint/provisioning
./Export-SharePointListSchema.ps1
./Import-SharePointListSchema.ps1 -WhatIf
```

See `src/sharepoint/provisioning/README.md` for authentication requirements, configuration fields, and safe import guidance.

## Prerequisites

Install only the tooling needed for the component you are working on:

- Git and Visual Studio Code
- Node.js and npm for the SPFx project
- Python and Azure Functions Core Tools for the Functions API
- PnP.PowerShell for SharePoint schema provisioning
- Power Platform CLI or Power Apps build tools for packaging and importing the Power Platform solution

## Source control and security

Keep components self-contained and keep solution-level guidance in `docs/`. Do not commit secrets, passwords, API keys, certificates, `local.settings.json`, environment-specific configuration, or dependency/build output such as `node_modules`.

## License

This repository is intended for internal development and maintenance of the UPM DHA Dashboard solution unless otherwise specified.
