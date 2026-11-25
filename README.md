# SMEF Container Loader

**Open the 3D view in your browser:** https://shinohub-create.github.io/smef-container-loader/
Click **Import Excel** and pick your Container Loader workbook, or drag the file onto the page. The workbook is read in your browser and is never uploaded anywhere. The workbook's **Open 3D View** links point here.

The rest of this file covers the optional Excel add-in, which needs permission to install add-ins.

## Excel add-in (optional)

The add-in puts a **3D View** button on Excel's Home tab. The button opens a side panel that reads your Container Loader workbook and draws the load in 3D:

- It reads cargo from the **Input** sheet, container sizes from **Containers**, and the container type from **Loading Plan** cell C3.
- It redraws by itself when you change those sheets. **Refresh** forces a re-read.
- Picking 20ft, 40ft or 40ft HC in the panel also writes that type into Loading Plan C3, so the workbook and the 3D view stay in step.
- The view has drag to rotate, scroll to zoom, 3D/Top/Side/Door views, box labels, a scale in cm, see-through mode, a loading-sequence player, and tabs for each container when more than one is needed.

Your workbook data never leaves Excel. The panel's code runs on your own computer, and the web address only serves the program files.

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Tells Excel where the add-in lives. This is the file you sideload. |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | The 3D panel |
| `commands.html` | Needed by the ribbon button |
| `assets/` | Icons |
| `lib/` | three.js (3D) and SheetJS, bundled so nothing else is downloaded |

## Step 1 – Host the folder on HTTPS

Office add-ins must be served over HTTPS. Two easy options:

**A. GitHub Pages (free, about 5 minutes)**
1. On GitHub, create a new **public** repository named exactly `smef-container-loader` under the `shinohub-create` account.
2. Upload every file and folder from this package to the root of the repository.
3. Go to **Settings → Pages**. Under **Source**, choose *Deploy from a branch*, pick `main` and `/ (root)`, then click **Save**.
4. After a minute, `https://shinohub-create.github.io/smef-container-loader/taskpane.html` should open. That address is already in `manifest.xml`, so you don't need to edit anything.

The code is public on GitHub Pages. Your workbook data is not, because it stays in Excel.

**B. Your IT team's web server, Azure Static Web Apps, or similar**
Upload the folder and note its HTTPS address. Then open `manifest.xml` in Notepad and use *Replace all*: change `https://shinohub-create.github.io/smef-container-loader` to your address (no trailing slash). Also change the `<AppDomain>` line to your site's root, for example `https://tools.smef.com`.

## Step 2 – Add it to Excel (sideload)

**Excel on the web (quickest test)**
1. Open the workbook in Excel for the web.
2. Go to **Home → Add-ins → More Add-ins → Upload My Add-in** (in some versions: **Insert → Add-ins**, or **Home → Add-ins → Advanced**).
3. Browse to `manifest.xml` and click **Upload**. **3D View** appears on the Home tab.

**Excel for Windows (desktop)**
1. Put `manifest.xml` in a folder, for example `C:\SMEF-Addin`.
2. Right-click the folder, then choose **Properties → Sharing → Share**, and share it with yourself. Note the network path it shows, for example `\\YOUR-PC\SMEF-Addin`.
3. In Excel, go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Paste the network path into *Catalog Url* and click **Add catalog**. Tick **Show in Menu**, click **OK**, then restart Excel.
4. Go to **Home → Add-ins → More Add-ins → SHARED FOLDER**, select **SMEF Container Loader**, and click **Add**.

**Excel for Mac**
1. In Finder, press **Go → Go to Folder** and open `~/Library/Containers/com.microsoft.Excel/Data/Documents/`. Create a folder named `wef` if there isn't one.
2. Copy `manifest.xml` into `wef`, then restart Excel.
3. Go to **Home → Add-ins** (or **Insert → My Add-ins**). The add-in is listed under *Developer add-ins*.

**For the whole company**
An IT admin can deploy `manifest.xml` to everyone from **Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps**.

## Requirements

You need Microsoft 365 or Excel 2021 or later (Windows, Mac or web). The panel uses Excel API 1.7.

## Workbook layout it expects

- **Input**: a header row containing *Model*, *Quantity*, *Length*, *Width*, *Height*, *Weight* and *Max stack*, with cargo rows below it. Reading stops at the *TOTAL* row.
- **Containers**: the type in column A (20ft / 40ft / 40ft HC), and internal L/W/H cm in columns B–D. Max payload goes in column F.
- **Loading Plan**: the selected container type in C3.

Packing follows the workbook's Calc sheet. Units stay upright, and each model is tried lengthwise and turned 90°. Each row fills the container width, stacked up to the height or the max-stack limit. Rows load from the front wall towards the door. When cargo spills into a second container, the 3D view moves the whole row into the next container instead of splitting it.
