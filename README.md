# UPM DHA Dashboard

## Overview

The **UPM DHA Dashboard** repository contains the source code and documentation for the UPM DHA Dashboard solution.

This repository is organized as a **monorepo**, with multiple independent projects maintained under a single Git repository. Each project has its own source code, configuration, dependencies, and build process while sharing a common repository for version control and documentation.

## Repository Structure

```text
UPM DHA Dashboard/
│
├── .gitignore
├── README.md
│
├── src/
│   ├── azure-functions/
│   │   ├── function_app.py
│   │   ├── host.json
│   │   ├── local.settings.json
│   │   └── requirements.txt
│   │
│   └── sharepoint/
│       └── spfx/
│           ├── package.json
│           ├── gulpfile.js
│           ├── tsconfig.json
│           └── README.md
│
└── docs/
    └── Project documentation
```

---

# Components

## Azure Functions

**Location**

```text
src/azure-functions/
```

This folder contains a Python Azure Functions app built with Azure Functions v4 and the decorator-based `FunctionApp` model.

Key files:

* `function_app.py` - contains HTTP-triggered functions.
* `host.json` - Azure Functions host configuration.
* `requirements.txt` - Python package dependencies.
* `local.settings.json` - local environment settings for development.

Current HTTP-triggered routes implemented in `function_app.py` include:

* `DqsExtractZip` — extracts ZIP archives and returns file entries as base64 payloads.
* `DqsConvertCsvToXlsx` — converts CSV content into XLSX format with optional header and table formatting.
* `DqsSplitPdf` — splits a PDF into individual page files encoded as base64.

Dependencies include `azure-functions`, `beautifulsoup4`, `markdown2`, `openpyxl`, `pypandoc`, `PyPDF2`, `python-pptx`, and `python-docx`.

> **Note**
>
> `local.settings.json` contains local development secrets and settings. Do not commit this file to source control.

Local development:

```powershell
cd src/azure-functions
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
func start
```

Use `func start` to launch the local Azure Functions runtime and test the HTTP routes defined in `function_app.py`.

---

## SharePoint Framework (SPFx)

**Location**

```text
src/sharepoint/spfx/
```

This folder contains a SharePoint Framework project for the **DHA Portfolio Manager** solution.

Project details:

* SPFx version `1.21.1`
* React `17.0.1`
* Package name: `dha-portfolio-manager-spap`
* Supports a modern web part and a single part app page experience

Key files:

* `package.json` — project dependencies and build scripts
* `gulpfile.js` — SPFx build tasks
* `tsconfig.json` — TypeScript compiler configuration
* `config/serve.json` — local serving configuration
* `sharepoint/solution/` — generated package output

Build and package commands:

```powershell
cd src/sharepoint/spfx
npm install
npm run build
npm run package-solution
```

Local development:

```powershell
cd src/sharepoint/spfx
npm install
npm run serve
```

Update `config/serve.json` if needed to point to your local SharePoint workbench or development site URL.

The deployable SharePoint package is generated under `src/sharepoint/spfx/sharepoint/solution/`.

Deploy to the SharePoint App Catalog:

1. Build and package the solution:

```powershell
cd src/sharepoint/spfx
npm install
npm run package-solution
```

2. Upload the generated `.sppkg` file from `src/sharepoint/spfx/sharepoint/solution/` to your tenant App Catalog.
3. If the package is tenant-deployable, enable tenant-wide deployment; otherwise add the app to the target site.
4. Add the app to the target site, then add the web part or single part app page to a SharePoint page.

For more details, see `src/sharepoint/spfx/README.md`.

---

# Documentation

Project documentation is stored in:

```text
docs/
```

Suggested documentation includes:

* Solution architecture
* Development guide
* Environment configuration
* Troubleshooting
* Release notes

---

# Development Prerequisites

Depending on the project being developed, the following tools may be required:

* Git
* Visual Studio Code
* Node.js and npm
* SharePoint Framework toolchain
* Azure Functions Core Tools
* Azure CLI
* Microsoft 365 CLI (optional)
* PnP PowerShell (optional)

Refer to the documentation in the `docs` folder for detailed setup instructions.

---

# Source Control Guidelines

## Repository

This repository contains multiple independent projects organized within a single Git repository.

Projects should remain self-contained, with their own configuration and dependencies.

## Do Not Commit

The following items should not be committed:

* Secrets
* Passwords
* API keys
* Certificates
* Environment-specific configuration
* Build output
* `node_modules`
* `local.settings.json`

---

# Recommended Branch Strategy

| Branch      | Purpose                   |
| ----------- | ------------------------- |
| `main`      | Production-ready code     |
| `develop`   | Active integration branch |
| `feature/*` | New features              |
| `bugfix/*`  | Bug fixes                 |
| `hotfix/*`  | Production fixes          |

---

# Repository Conventions

* Keep each project self-contained.
* Store solution-level documentation in the `docs` folder.
* Keep repository-level configuration in the repository root.
* Follow consistent commit messages and branching practices.

---

# Getting Started

Clone the repository:

```bash
git clone <repository-url>
```

Navigate to the repository:

```bash
cd "UPM DHA Dashboard"
```

Open the repository in Visual Studio Code:

```bash
code .
```

Refer to the documentation for project-specific setup instructions.

---

# License

This repository is intended for internal development and maintenance of the UPM DHA Dashboard solution unless otherwise specified.
