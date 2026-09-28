# DHA Portfolio Manager — agent guidance

## Project and scope

This repository maintains the DHA Portfolio Manager (also called the DHA Dashboard) for Upside Capital Holdings (UCH). The user-facing application is a SharePoint Framework (SPFx) web part and full-page app.

Maintained areas, with paths relative to the repository root:

| Area | Location | Responsibility |
| --- | --- | --- |
| Dashboard | `src/sharepoint/spfx/` | React UI, SharePoint data access, and authenticated flow requests |
| Azure Functions | `src/azure-functions/` | Python HTTP endpoints for document processing and regular expressions |
| Power Platform | `src/powerplatform/DallasHousingAuthority/` | Unpacked solution containing synchronization, transformation, analysis, and notification flows |
| SharePoint provisioning | `src/sharepoint/provisioning/` | PnP PowerShell export and import of list and library schemas |

- Skip `src/power-apps/` and `src/power-automate/` during searches, implementation, and validation unless the user explicitly requests them. The Power Apps implementation was abandoned because of licensing requirements. The Power Automate folder holds exported solution ZIPs; the editable source is the unpacked Power Platform solution above.
- Scope searches to maintained areas; exclude dependencies, build output, and local configuration unless specifically relevant.
- Other folders are supporting references, not additional maintained components by default. `docs/` and `discovery/` are ignored by Git; put shared project guidance in tracked files.
- Consult the root README and the relevant component README or existing instructions, then verify behavior against source. Documentation may lag implementation.

## Architecture and shared contracts

SPFx reads and updates SharePoint lists on the current site. Power Automate synchronizes ResMan properties, units, leases, people, and receivables with SharePoint, processes documents and communications, and coordinates analysis and notifications. Selected flows call Azure Functions for conversions, ZIP extraction, and regular-expression operations; analysis flows also integrate with Azure OpenAI. Provisioning supplies SharePoint schemas separately from application deployment.

- Trace producers and consumers across components before changing a SharePoint field, lookup, endpoint, payload, or response format.
- Preserve SharePoint internal field names and distinguish numeric SharePoint item/lookup IDs from ResMan domain identifiers such as property and billing-account IDs.
- Keep environment-specific endpoints and settings in the existing configuration mechanisms. Do not hardcode credentials or replace authenticated requests with embedded keys or signed callback URLs in browser code.

## SPFx dashboard

Working directory: `src/sharepoint/spfx/`.

- Preserve compatibility with the repository's SPFx `1.21.1` and React `17.0.1` dependencies. Use the existing npm/Gulp toolchain and lockfile; dependency upgrades are separate work unless required by the task.
- The entry point is `src/webparts/dhaPortfolioManager/DhaPortfolioManagerWebPart.ts`. Most UI and data logic is in `components/Dashboard.tsx` beneath that directory; styles are in `components/Dashboard.module.scss`. Edit the SCSS source rather than hand-editing generated style mappings.
- Preserve `SharePointWebPart` and `SharePointFullPage` manifest support and `supportsFullBleed: true`. Check both hosts when changing layout or navigation; the local workbench does not establish full-page behavior.
- Keep list names configurable through the web-part properties. Defaults are `DHA Intake`, `ResMan People`, and `ResMan TLedger`.
- `DHA Configuration` uses `Title` and a multiline text column with internal name `Value`. Its current keys are `Refresh Balances Endpoint`, `LLM Request Endpoint`, `Summarize Transactions Endpoint`, and `Analyze Balance Instructions`. Preserve the permission checks around configuration management.
- SharePoint requests use the SPFx context. Authenticated flow requests acquire a token for `https://service.flow.microsoft.com/`; `config/package-solution.json` requests Microsoft Flow Service `user_impersonation`. Preserve this authentication contract.
- Preserve list-query escaping, paging where implemented, lookup resolution, date handling, and ledger calculations when changing data access or rendering.

Commands from this working directory:

```powershell
npm install
npm run build
npm run package-solution
# For interactive development with a configured SharePoint workbench:
npm run serve
```

`build` runs `gulp bundle --ship`; `package-solution` runs `gulp package-solution --ship`; `serve` runs `gulp serve-deprecated`. Configure the development workbench URL in `config/serve.json` before serving. The package is generated at `sharepoint/solution/dha-portfolio-manager-spap.sppkg`. Building and packaging do not deploy it; deployment instructions are in the component README.

## Azure Functions

Working directory: `src/azure-functions/`.

- `function_app.py` contains the Python decorator-based `FunctionApp`; `host.json` configures the host and `requirements.txt` declares dependencies. Function-key authentication is the app default.
- Existing routes are `DqsExtractZip`, `DqsConvertCsvToXlsx`, `DqsSplitPdf`, `DqsFillPpt`, `DqsStrToDoc`, `DqsExtractPpt`, `DqsRegEx`, and `DqsReplaceTxtInPpt`.
- Preserve route names, request keys, response shapes, status codes, and Base64 conventions consumed by flows. Check callers before changing them.
- Preserve ZIP path filtering and input validation. Regular-expression limits currently cap patterns at 1,024 characters and input text at 20,000 characters; these are size limits, not an execution timeout.
- Applicable document conversions require a Pandoc executable under `tools/pandoc.exe` on Windows or `tools/pandoc` on other platforms. The tools folder is ignored by Git, so do not assume a fresh checkout includes this runtime dependency.
- Keep `local.settings.json`, function keys, virtual environments, and Azurite data local. Do not add `azure-functions-worker` to requirements; the platform manages it.

Local setup and syntax validation:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m py_compile function_app.py
# Start the local host when runtime testing is needed and settings are configured:
func start
```

For changed endpoints, test targeted synthetic requests covering normal input, malformed or missing input, and relevant limits. Verify the response contract, not only successful execution. Avoid invoking unrelated live flows as a test.

## Power Platform solution

Working directory: `src/powerplatform/DallasHousingAuthority/`.

- `DallasHousingAuthority.cdsproj` defines the solution project. Editable components are under `src/Workflows/`, `src/environmentvariabledefinitions/`, and `src/Other/`.
- Treat workflow JSON, companion `.json.data.xml` files, and solution metadata as a coherent source set. Preserve workflow GUIDs, filenames, child-flow references, connection references, and environment-variable schema names unless deliberately changing the corresponding component.
- Action names can be referenced by expressions and `runAfter`; review those references when renaming or moving actions. Preserve trigger authentication, conditions, schedules, retry behavior, and concurrency unless they are part of the requested change.
- Use existing environment variables and Key Vault integration for configuration and secrets. Do not copy exported environment values into examples or introduce literal credentials.
- Validate edited JSON and XML, inspect changed expressions and references, and check callers and downstream actions. Report unresolved dependencies rather than inventing replacement IDs or connections.
- Parsing alone does not verify flow behavior or solution importability. Use the Power Platform CLI/build tools and a configured non-production environment when packaging or integration validation is part of the task; no repository-specific packaging script or dedicated automated test suite is established here.
- Keep exported ZIPs and `bin/` or `obj/` output out of source changes. Do not use the deprecated export folder as the packaging source.

## SharePoint provisioning

Working directory: `src/sharepoint/provisioning/`. Follow its README for configuration and authentication.

- `Export-SharePointListSchema.ps1` exports schemas and an inventory; `Import-SharePointListSchema.ps1` applies the template to an existing target site. They do not transfer list items or documents.
- Preserve exact list-URL exclusions, the document-library title allowlist, and hidden-list handling. Export operates on one web, not a recursive traversal of subsites.
- Use the tracked `sharepoint-list-schema-config.example.json` as the starting point for ignored local configuration. PnP.PowerShell `3.1.0` or later and an Entra application ID are required; the scripts accept `-ClientId` or use `ENTRAID_APP_ID`.
- Syntax-check changed scripts without executing them. For a configured run, export with `./Export-SharePointListSchema.ps1`, review the generated template and inventory, then preview with `./Import-SharePointListSchema.ps1 -WhatIf` before applying changes.
- `-WhatIf` still authenticates to the target site; it is not an offline check. Validate against a non-production target and retain a backup before updating an existing site, as the README describes.
- Keep local configuration, generated templates, and inventories out of source control. Use `-Force` only when intentionally replacing existing export output.

## Working agreements and completion checks

- Inspect the working tree before editing. Preserve unrelated user changes; do not reset or overwrite them as part of cleanup.
- Make focused changes using the surrounding code's conventions. Avoid bulk formatting, incidental upgrades, and unrelated refactoring.
- Do not commit dependencies or generated packages, build folders, caches, or local settings. Keep secrets, signed URLs, resident data, and document contents out of committed examples and diagnostic logs; use synthetic fixtures.
- Local validation and packaging are distinct from deployment, solution import, live data updates, or notification delivery. Perform live operations only within the user's authorized task, and reuse authorization already given.
- No dedicated automated test suite was found in the maintained source. Do not claim tests exist or pass without checking. Add focused regression coverage when warranted by the change, rather than scaffolding a test framework for documentation-only edits.
- For SPFx code changes, run `npm run build` and `npm run package-solution`; supplement relevant UI changes with checks in the appropriate SharePoint hosts when available. For other components, use the checks described above.
- For documentation-only changes, check Markdown, referenced paths and commands, and the final diff; application builds are unnecessary.
- Before finishing, inspect the diff for unintended files or generated changes. Report what changed, checks actually performed, and any runtime or live-environment behavior that remains unverified.
