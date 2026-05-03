# Heavy Calc Assist Handoff

## Current State

Heavy Calc Assist is an Excel Desktop Office.js add-in backed by a FastAPI calc engine.

The current MVP reads Payroll input data from the workbook, sends included rows to the backend, calculates payroll outputs, and writes the results back to the workbook.

It does not yet attempt to replace the full payroll model. Current implemented outputs are:

- Headcount / FTE
- Base salary total
- Base salary domestic
- Base salary international
- Base salary COGS
- Medical benefits
- 401k benefits
- Other benefits

## Deployment

Frontend add-in:

```text
https://xf1-advisory-services.github.io/HCA/taskpane.html
```

Backend:

```text
https://hca-calc-engine.onrender.com
```

Health check:

```text
https://hca-calc-engine.onrender.com/health
```

Microsoft 365 Admin Center manifest:

```text
https://raw.githubusercontent.com/XF1-Advisory-Services/HCA/main/deploy/m365/hca-production-manifest.xml
```

The manifest version is currently:

```text
1.0.2.0
```

If task pane URL, icon URL, display name, permissions, or app metadata changes, bump the manifest version before updating Admin Center.

## Repo Structure

```text
excel-addin/      Office.js task pane add-in
calc-engine/      FastAPI backend
deploy/m365/      Production Microsoft 365 manifest
shared/           Payload notes
README.md         Setup/deploy instructions
HANDOFF.md        Project context for future work
```

## Key Files

Frontend:

```text
excel-addin/src/taskpane/taskpane.js
excel-addin/taskpane.html
excel-addin/src/taskpane/taskpane.css
```

Backend:

```text
calc-engine/app/main.py
calc-engine/app/schemas.py
calc-engine/app/payroll_headcount.py
```

Production manifest:

```text
deploy/m365/hca-production-manifest.xml
```

GitHub Pages workflow:

```text
.github/workflows/deploy-addin.yml
```

Render blueprint:

```text
render.yaml
```

## Workbook Config Contract

The add-in reads the whole used range of the `Config` sheet in one bulk operation. It does not hardcode Config cell addresses.

The Config sheet must have these headers:

```text
Section | Setting | Value
```

Required current settings:

```text
Model   Last actuals date
Model   Model end date
Model   Financial year end month

Payroll Data load Sheet
Payroll Cell range
Payroll Headers
Payroll Filter column
Payroll Output sheet
Payroll Headcount output start cell
Payroll Base salary total output start cell
Payroll Base salary domestic output start cell
Payroll Base salary international output start cell
Payroll Base salary COGS output start cell
Payroll Medical - Domestic
Payroll Medical - International
Payroll 401k - Domestic
Payroll 401k - International
Payroll Other Benefits - Domestic
Payroll Other Benefits - International
Payroll Medical output start cell
Payroll 401k output start cell
Payroll Other Benefits output start cell
```

Example values from the current workbook:

```text
Payroll Data load Sheet                         PayrollData
Payroll Cell range                              B5:R1531
Payroll Headers                                 B4:R4
Payroll Filter column                           R
Payroll Output sheet                            HCA_Output
Payroll Headcount output start cell             E4
Payroll Base salary total output start cell      E17
Payroll Base salary domestic output start cell   E30
Payroll Base salary international output start cell E44
Payroll Base salary COGS output start cell       E57
Payroll Medical output start cell                E70
Payroll 401k output start cell                   E83
Payroll Other Benefits output start cell         E96
```

## PayrollData Contract

Current calculation code assumes this fixed input layout inside the configured Payroll data range:

```text
B  EmployeeID
C  FS_Category
D  Status
E  Department
F  Start Date
G  Termination Date
M:Q Annual base salary assumptions by financial year label
R  Include in LOAD
```

In backend zero-based payload field indexes:

```text
FS_Category      1
Status           2
Department       3
Start Date       4
Termination Date 5
```

The add-in filters rows where the configured filter column equals `1`.

## Calculation Rules

### Timeline

The add-in derives monthly forecast periods from:

```text
calculation start = month after Last actuals date
calculation end   = Model end date
```

Financial year is derived from `Financial year end month`.

If FY end month is `4`, then:

```text
Apr-2026 = FY2026
May-2026 through Apr-2027 = FY2027
```

### Headcount / FTE

For each employee/month:

```text
active_start = max(employee start date, month start date)
active_end   = min(employee end date, month end date)
FTE          = active days / days in month
```

If there is no overlap, FTE is `0`.

Blank termination date means active through model end date.

Output is grouped by Department and rounded to 2 decimals.

### Base Salary

For each employee/month:

```text
monthly base salary = annual salary for period financial year / 12 * FTE
```

Salary field lookup by financial year is precomputed once per period, not repeatedly discovered for every employee row.

Outputs:

- total: all included employees
- domestic: `Status = Domestic`
- international: `Status = International`
- COGS: `FS_Category = COS`, regardless of Status

### Benefits

Benefit assumptions are monthly constants by Status.

For each employee/month:

```text
if FTE > 0:
    benefit = full monthly benefit assumption for employee Status
else:
    benefit = 0
```

Benefits are not prorated by FTE.

Outputs:

- Medical
- 401k
- Other Benefits

Each output is grouped by Department.

## Development Workflow

After changes, run:

```powershell
cd excel-addin
npm run build
npx office-addin-manifest validate ..\deploy\m365\hca-production-manifest.xml
```

Backend syntax check:

```powershell
cd calc-engine
python -m py_compile app\main.py app\schemas.py app\payroll_headcount.py
```

Then commit and push to:

```text
https://github.com/XF1-Advisory-Services/HCA.git
```

GitHub Actions deploys the add-in to GitHub Pages.

Render deploys the backend from the company repo.

## Important Notes

- Do not commit workbook files. `.gitignore` excludes Excel files.
- Do not change the production manifest ID unless intentionally creating a new add-in identity.
- If the Microsoft 365 Admin Center app needs a manifest refresh, update the existing app using the raw manifest URL above.
- If only JavaScript/backend code changes and manifest URLs/metadata do not change, Admin Center usually does not need a manifest update.
- Render free tier may sleep; first backend request can be slow.

## Suggested Next Work

- Add output section titles/formatting in `HCA_Output`.
- Add warnings in task pane for skipped rows or unknown Status values.
- Move fixed PayrollData field indexes to Config if the model layout becomes less stable.
- Add unit tests for payroll calculations.
- Add remaining payroll lines such as bonus, severance, payroll taxes, or allocation logic.
