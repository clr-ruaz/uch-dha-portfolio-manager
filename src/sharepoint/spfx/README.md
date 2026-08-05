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
2. Enable tenant-wide deployment if desired; otherwise add the app to the target site.
3. In the target SharePoint site, select **New > Page**.
4. Select the **Apps** tab and choose **DHA Portfolio Manager** for a full-page app, or add **DHA Portfolio Manager** to a standard modern page section.
5. Configure the list names through the property pane if the defaults differ.

The manifest supports `SharePointWebPart` and `SharePointFullPage`, with full-bleed rendering enabled. Before running `npm run serve`, replace the example tenant and site in `config/serve.json` with your development workbench URL. Full-page behavior must be verified on a deployed SharePoint page.
