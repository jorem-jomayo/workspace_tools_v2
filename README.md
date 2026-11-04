# Workspace Tools V2

Browser-based file tools. Uploaded files are processed locally in your browser and are not sent to a processing server.

## Run locally

Open `index.html` from a local web server so the reference workbooks and scripts can load correctly. For example, from the project folder run:

```powershell
python -m http.server 8001 --bind 127.0.0.1
```

Then open `http://localhost:8001`. The site uses the checked-in browser libraries in `vendor/` and local Inter font files in `assets/fonts/`; no Node.js setup or external font request is required.

## Tools

- **VRP Mapper:** process one or more MC2 CSVs or an FCL `.xlsx` workbook with a `SUMMARY` sheet. Optionally filter by reference code. Downloads the mapped template CSV and release CSV. Mapping uses `demand_letter_template.csv`, `cmd_id.xlsx`, and `amounts.csv`.
- **Field Result:** extracts columns from the `RESULT` sheet in an `.xlsx` workbook and downloads a formatted workbook plus `fieldman.csv`.
- **E-SIGN FIXER:** detects signatures on a PNG/JPG page and downloads transparent PNGs in a ZIP file and a Word document.

## Publish with GitHub Pages

1. Push the repository to GitHub.
2. Open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select the branch containing `index.html` and the `/ (root)` folder, then save.

The app and its file processing run in the visitor's browser; GitHub Pages does not need a Python backend. Keep the reference files in the repository root so VRP Mapper can fetch them. Do not publish private or confidential reference files to a public repository.

Bundled libraries and their licenses are included in `vendor/`.
