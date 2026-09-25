# Move page

A SharePoint Framework (SPFx) list view command set that adds a **Move page** button to the **Site Pages** library. Select one or more pages, pick a destination folder, and the pages are moved there.

![SPFx](https://img.shields.io/badge/SPFx-1.22.2-green.svg)

## How it works

- The button appears in the Site Pages command bar only when one or more **pages** are selected with the row checkbox. It stays hidden when nothing is selected or when a folder is selected.
- The dialog lets you browse the library's folders (the `Forms` and `Templates` system folders are hidden) and choose **Move here**.
- Pages are moved with `SP.MoveCopyUtil.MoveFileByPath` and **never overwrite** a page with the same name. If a page with that name already exists in the target folder, that page fails with SharePoint's error message and the others still move.
- You can't move pages into the folder they are already in.
- Moving a page changes its URL, so navigation links to it need updating. The dialog warns about this.

## Prerequisites

- Node.js `>=22.14.0 <23`
- Permission to upload to the tenant app catalog (SharePoint admin or app catalog owner)

## Build

```bash
npm install
npm run build
```

This runs the tests and produces `sharepoint/solution/move-page.sppkg`.

To debug locally, set your tenant's domain once. `config/serve.json` uses a `{tenantDomain}` placeholder so the tenant name isn't committed:

```powershell
setx SPFX_SERVE_TENANT_DOMAIN "<tenant>.sharepoint.com"   # then open a new terminal
npm start
```

`pageUrl` in `config/serve.json` is the SPFx default placeholder. Before debugging, point it at your Site Pages library locally, e.g. `https://{tenantDomain}/sites/<site>/SitePages/Forms/AllPages.aspx`, and don't commit that change.

## Deploy

The button is registered by a feature that activates when the app is added to a site (`sharepoint/assets/elements.xml`, list template `119` = Site Pages). **No PnP script or `Add-PnPCustomAction` is needed**, and running one creates a duplicate registration.

1. Bump the versions (see [Releasing an update](#releasing-an-update)) and run `npm run build`.
2. Upload `sharepoint/solution/move-page.sppkg` to the **tenant app catalog** (SharePoint admin center → More features → Apps → App Catalog → Apps for SharePoint).
3. Click **Deploy**. There's no "add to all sites" option: the app is added one site at a time.
4. Check that the catalog shows the version you just built.
5. On the target site: **Site Contents** → **New** → **App** → **From your organization** → add **move-button-client-side-solution**.
6. Open Site Pages, hard refresh (Ctrl+Shift+R), and tick a page's checkbox. **Move page** appears in the command bar.

Use either the tenant app catalog or a site collection app catalog, never both. Both packages have the same solution ID, and a site catalog copy overrides the tenant one for that site.

### Releasing an update

Bump all three versions together so SharePoint picks up changes to both the code and the feature:

| File | Field | Example |
| --- | --- | --- |
| `config/package-solution.json` | `solution.version` | `2.1.2.0` → `2.1.3.0` |
| `config/package-solution.json` | `features[0].version` | `1.1.2.0` → `1.1.3.0` |
| `package.json` | `version` (the component version) | `0.0.4` → `0.0.5` |

Then build, upload over the existing package in the app catalog, and click **Deploy**. If the site offers an update in Site Contents, apply it.

### Removing

Remove the app from the site's **Site Contents**. That deactivates the feature and removes the button's registration.

## Troubleshooting

**The button doesn't appear**

1. Select a *page* using the checkbox, not a folder. Clicking the page title opens the page instead of selecting it.
2. Open the browser console on Site Pages and look for `Could not load move-button-command-set`.
3. Check the site's registrations for duplicates. There should be exactly one:

   ```powershell
   Connect-PnPOnline -Url https://<tenant>.sharepoint.com/sites/<site> -Interactive -ClientId <client-id>
   Get-PnPCustomAction -Scope All | ? ClientSideComponentId -eq "302145c4-06df-4dab-87c7-238c813ecb2c" | ft Id, Name, Title, Scope
   # Remove any extra that isn't the app's own registration (Name is the registration's own {GUID}):
   Remove-PnPCustomAction -Identity <Id> -Scope Web -Force
   ```

**`Cannot destructure property 'id' of 'a' as it is undefined`**

A framework dependency in the component manifest points to a version that isn't in the tenant's component store. This happened with React 17.0.2: the tenant has 17.0.1 and 18.3.1. `react` and `react-dom` are pinned to exactly `17.0.1` in `package.json`. Don't let them float. After a build, you can confirm the version in `sharepoint/solution/debug/<feature-id>/Extension_*.xml`.

**The old icon or title still shows after an update**

That's the browser cache. Hard refresh, or check in a private window.

## Reference

| Item | Value |
| --- | --- |
| Solution ID | `11432589-3925-42a7-89f8-f85058d9ced6` |
| Component ID | `302145c4-06df-4dab-87c7-238c813ecb2c` |
| Feature ID | `b2d1b27d-7b73-46c1-9258-aa75d8593f40` |
| Command ID | `MOVE_PAGE` |
| Location | `ClientSideExtension.ListViewCommandSet.CommandBar` |
| Target list template | `119` (Site Pages) |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 2.1.2.0 | 2026-09-25 | Fluent icon and "Move page" title |
| 2.1.1.0 | 2026-09-25 | Pin React to 17.0.1 to fix the extension failing to load |
| 2.1.0.0 | 2026-09-25 | Register via feature on app install; safer moves (no overwrite, special characters, folder handling) |
