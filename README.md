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

Hosted backend:

```text
https://hca-calc-engine.onrender.com
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
- Config columns named:
  - `Section`
  - `Setting`
  - `Value`

Required `Model` settings:

- `Last actuals date`
- `Model end date`
- `Financial year end month`

Required `Payroll` settings:

- Data sheet: `PayrollData`
- `Data load Sheet`
- `Cell range`
- `Headers`
- `Filter column`

Example:

```text
Section   Setting                  Value
Model     Last actuals date        31-Mar-26
Model     Model end date           30-Apr-28
Model     Financial year end month 4
Payroll   Data load Sheet          PayrollData
Payroll   Cell range               B5:R1531
Payroll   Headers                  B4:R4
Payroll   Filter column            R
Payroll   Output sheet             HCA_Output
Payroll   Headcount output start cell E4
Payroll   Base salary total output start cell E17
Payroll   Base salary domestic output start cell E30
Payroll   Base salary international output start cell E44
Payroll   Base salary COGS output start cell E57
Payroll   Medical - Domestic       2464
Payroll   Medical - International  2162
Payroll   401k - Domestic          432
Payroll   401k - International     501
Payroll   Other Benefits - Domestic 157
Payroll   Other Benefits - International 20
Payroll   Medical output start cell E70
Payroll   401k output start cell E83
Payroll   Other Benefits output start cell E96
```

Headcount/FTE output is written as a table starting at the configured start cell:

```text
Department | Apr 2026 | May 2026 | ...
```

Base salary output uses:

```text
monthly base salary cost = annual salary for financial year / 12 * FTE
```

Benefits output uses:

```text
monthly benefit cost = full monthly benefit assumption when employee FTE is above 0
```

## Render Backend

The backend is ready for Render deployment from the repo root using `render.yaml`.

The first MVP uses this temporary environment variable:

```text
CORS_ORIGINS=*
```

After the add-in is hosted at a stable URL, restrict it to that origin:

```text
CORS_ORIGINS=https://your-github-username.github.io
```

This MVP has no authentication. Do not treat the preview endpoint as production-secure.

## Microsoft 365 Admin Center Deployment

The production add-in is intended to be hosted on GitHub Pages:

```text
https://xf1-advisory-services.github.io/HCA/taskpane.html
```

Upload this manifest in Microsoft 365 Admin Center:

```text
deploy/m365/hca-production-manifest.xml
```

Recommended first rollout:

1. Deploy to `Just me` first.
2. Test in Windows Desktop Excel.
3. Then assign to a small pilot group.

Microsoft notes that users might need to relaunch Office, and deployed add-ins can take time to appear on the ribbon.
