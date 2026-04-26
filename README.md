# External Calc Engine v01

Initial MVP for a Windows Desktop Excel Office.js add-in plus a Python FastAPI calc engine.

This version does not calculate payroll. It only proves that the Excel task pane can read the `Payroll` config row, bulk-load headers and data, filter rows where `Include in Load` equals `1`, and send a preview payload to the backend.

## Project Structure

- `excel-addin/` - Office.js task pane add-in.
- `calc-engine/` - FastAPI backend.
- `shared/` - Shared payload notes.

## Local Backend

From `calc-engine/`:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m uvicorn app.main:app --host localhost --port 8000 --ssl-keyfile "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" --ssl-certfile "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt"
```

Check:

```powershell
Invoke-RestMethod https://localhost:8000/health
```

## Local Excel Add-in

From `excel-addin/`:

```powershell
npm install
npm run certs
npm run dev
```

The local task pane URL is:

```text
https://localhost:3000/taskpane.html
```

The local manifest is:

```text
excel-addin/manifest.xml
```

## Excel Sideloading

For Windows Desktop Excel, the simple local sideload path is:

1. Create a trusted catalog folder, for example `C:\OfficeAddins\ExternalCalcEngine`.
2. Copy `excel-addin\manifest.xml` into that folder.
3. In Excel, go to `File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs`.
4. Add the folder path and tick `Show in Menu`.
5. Restart Excel.
6. Open the workbook, then use `Insert > My Add-ins > Shared Folder` and select the add-in.

## Workbook Assumptions

The workbook must contain:

- A `Config` sheet.
- A row where `Section` equals `Payroll`.
- Config columns named:
  - `Section`
  - `Data load Sheet`
  - `Load cell range`
  - `Column Header`
  - `Load Filter column`

For the current workbook, the `Payroll` config is expected to point to:

- Data sheet: `PayrollData`
- Header range: `B4:S4`
- Data range: `B5:S1531`
- Filter column: `S`

## Railway Backend

The backend is ready for Railway deployment from the `calc-engine/` folder. Railway will use the included `Dockerfile`.

Required environment variable after deployment:

```text
CORS_ORIGINS=https://your-github-username.github.io
```

For early testing only, you can temporarily use:

```text
CORS_ORIGINS=*
```

This MVP has no authentication. Do not treat the preview endpoint as production-secure.
