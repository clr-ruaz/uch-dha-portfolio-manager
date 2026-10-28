# DHA Portfolio Manager

SharePoint Framework 1.21.1 web part and Single Part App Page for the DHA Portfolio Manager. It reads and updates **DHA Intake** and reads **ResMan People** on the current SharePoint site.

## Build and package

```powershell
npm install
npm run build
npm run package-solution
```

The deployable package is generated at `sharepoint/solution/dha-portfolio-manager-spap.sppkg`.

## Deploy

1. Upload the `.sppkg` file to the tenant App Catalog.
2. In the SharePoint admin center, approve the pending **Microsoft Flow Service** `user_impersonation` API access request.
3. Enable tenant-wide deployment if desired; otherwise add the app to the target site.
4. In the target SharePoint site, select **New > Page**.
5. Select the **Apps** tab and choose **DHA Portfolio Manager** for a full-page app, or add **DHA Portfolio Manager** to a standard modern page section.
6. Configure the list names through the property pane if the defaults differ.
7. Create a **DHA Configuration** list with a multiline text column whose internal name is `Value`. As a site owner, use **Manage** in the dashboard command bar to save the Entra-authenticated Refresh Balances endpoint in that list.

The manifest supports `SharePointWebPart` and `SharePointFullPage`, with full-bleed rendering enabled. Full-page behavior must be verified on a deployed SharePoint page.

## Local workbench setup

From this directory, create your local configuration once:

```bash
# macOS / Linux
cp config/serve.example.json config/serve.json
```

```powershell
# Windows PowerShell
Copy-Item config/serve.example.json config/serve.json
```

If `config/serve.json` already exists, keep it rather than overwriting your local settings. Set its `initialPage` to your development SharePoint site's URL followed by `/_layouts/workbench.aspx`. The local `config/serve.json` is ignored by Git; commit shared configuration changes to `config/serve.example.json` using placeholder URLs only.

Trust the development HTTPS certificate on a new machine with `npx gulp trust-dev-cert`, then run `npm run serve` to start local debugging.
