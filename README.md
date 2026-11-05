# Workspace Tools V2

Browser-based file tools. VRP Mapper and E-SIGN process in the browser; Field Result uses the local Node.js server for memory-efficient workbook streaming. The optional Lark integration also uses the local server so its app secret stays out of browser code.

## Run locally with the Lark notifications

1. Fill in `.dev.vars` (already created locally with the app ID and table/view IDs from your link). Replace `LARK_APP_SECRET` with the **rotated** secret. If `.dev.vars` is missing, copy `.dev.vars.example` first.
2. Keep `.dev.vars` private. It is ignored by Git. Never paste the app secret or dashboard password into chat or commit them.
3. Install the Field Result streaming processor dependency once:

   ```powershell
   python -m pip install -r requirements.txt
   ```

4. In the Lark custom app, enable and publish `wiki:wiki:readonly` and either `bitable:app:readonly` or `base:record:retrieve`. The Bitable must also be shared with the custom app: in the Base's **Share / Manage permissions**, add the app and grant read access to the target table/view. If advanced permissions are on, grant the app the required role on that Base/table as well. A user being signed in to Lark in the browser does not grant these permissions to the server-side app identity.
5. Start the server from the project folder:

   ```powershell
   npm.cmd start
   ```

6. In VS Code's **Ports** panel, set port `8001` visibility to **Private** and allow only your own signed-in account. Then change `TUNNEL_ACCESS_MODE=unconfirmed` to `TUNNEL_ACCESS_MODE=private` in `.dev.vars`, start the server, and forward port `8001`. Open `https://cfbbqbh6-8001.asse.devtunnels.ms`; the tunnel applies your VS Code/GitHub account access and the app opens without a separate username/password prompt. The bell uses the same private connection automatically—no Worker URL or key needs to be entered in the UI.

The server binds to `127.0.0.1:8001`; the **Private** Dev Tunnel is the access control for the app and its Lark API. The server refuses requests unless `TUNNEL_ACCESS_MODE=private` is set in the ignored `.dev.vars` file, and reads all Lark secrets from that file. It checks the Lark view on page load and every five minutes. Export downloads the selected view as `NEW VRP ACCOUNTS FOR TAG_PIF WITH DL_JOREM FOR UPLOAD.csv`, using the matching 34-column upload template (blank values are kept in their columns). Counts and exports support views with up to 24,000 records.

### Lark configuration

In `.dev.vars`, these settings are already filled in where available; use the following variable names:

- `PUBLIC_APP_ORIGIN`: exact HTTPS origin of the forwarded VS Code tunnel (no trailing slash).
- `TUNNEL_ACCESS_MODE=private`: set this only after the actual VS Code port visibility is **Private**. This configuration flag alone does not make the tunnel private.
- `LARK_APP_ID`, `LARK_APP_SECRET`: custom app credentials. Use the rotated secret, not the one previously shared in chat.
- `LARK_WIKI_TOKEN`, `LARK_TABLE_ID`, `LARK_VIEW_ID`: identifiers from the Lark wiki/Bitable link.

If the Dev Tunnel URL changes, update `PUBLIC_APP_ORIGIN` and restart the server. **Do not set the tunnel to Public while the app server is running**: there is intentionally no app password prompt in private-tunnel mode. Stopping `npm start` disconnects the Lark notifications; VS Code forwarding is not a permanent production backend.

### Lark reports `1254302: RolePermNotAllow`

Lark denied the custom app's role on the Bitable. Confirm the read scope is enabled and published in the Lark Developer Console, then share the Bitable itself with the app and grant read access to the table/view. For Bases with advanced permissions, make sure the app's role is allowed to query records. Restart the local server after changing `.dev.vars`; publishing permission-scope changes may also require republishing the Lark app.

## Tools

Click the theme icon in the top bar to cycle between **System setting**, **Light mode**, and **Dark mode**. The icon indicates the current setting, which is saved in this browser.

The browser tab title shows the Lark record count, and the tab icon displays a red dot when the count is greater than zero.

- **VRP Mapper:** process one or more MC2 CSVs or an FCL `.xlsx` workbook with a `SUMMARY` sheet. Use **Get REF CODE (Column M)** to copy unique codes from the selected source file(s) to the clipboard, or enter codes manually in the reference-code box to filter processing. Downloads the mapped template CSV and release CSV. Mapping uses `demand_letter_template.csv`, `cmd_id.xlsx`, and `amounts.csv`.
- **Field Result:** uploads the selected `.xlsx` workbook to the local Node server, which uses Python `openpyxl` read-only mode to stream the `RESULT` sheet row by row and writes a formatted workbook plus `fieldman.csv` to temporary storage. Select **PIF** (bank name in source column AG contains “PIF”) or **ALL CAMPAIGN** in settings before processing; the Process button stays disabled until a choice is made. Selecting either campaign clears any exact-bank filter so **ALL CAMPAIGN** includes every bank. The selection is sent as both a query parameter and a header so proxies that remove custom headers do not silently change the campaign. You can also enter an exact bank name from source column AG to process only that bank; bank names are not loaded from the workbook. The result notice identifies the campaign used, and the app checks that the server reports the same campaign before displaying output. Restart `npm.cmd start` after updating the server or Python processor. The browser only receives the first 100 preview rows. Uploads are limited to 512 MB; actual processing time depends on workbook contents and device performance. Output files expire after 30 minutes. GitHub Pages has no processing backend and uses the browser-only fallback for smaller workbooks.
- **E-SIGN FIXER:** detects signatures on a PNG/JPG page and downloads transparent PNGs in a ZIP file and a Word document.

## Publish the static tools with GitHub Pages

GitHub Pages can host the browser-only tools, but it cannot run the local Lark API server. The automatic Lark notifications work only when the protected Node server is running through the forwarded port (or when a separately secured production backend is configured). Do not publish private reference files such as `cmd_id.xlsx` or `amounts.csv` to a public repository.

The checked-in browser libraries and their licenses are in `vendor/`; local fonts are in `assets/fonts/`.

Run the API and login tests with:

```powershell
npm.cmd test
```
