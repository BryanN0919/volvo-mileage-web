# Volvo Mileage Reimbursement Generator

A fully client-side web page that turns a Volvo `volvo-trips-log.xlsx` export into a completed monthly mileage reimbursement workbook based on **G.G.'s July 2026 Mileage form.xlsx**.

## Use it

1. Keep all files in this folder together.
2. Open `index.html` in Chrome, Edge, or another modern browser.
3. Import the new `volvo-trips-log.xlsx` file.
4. Review the automatically selected trips.
5. Adjust the mileage reimbursement rate if needed.
6. Click **Generate Excel Form**.

No web server, PowerShell, Python, Excel installation, or internet connection is required. Excel processing happens locally in the browser.

## Automatic rules

- Includes trips with `Category = Business`.
- Excludes trips with zero or negative odometer change.
- Calculates miles as `End odometer - Start odometer`.
- Sorts trips newest first, matching the July sample.
- Detects the reimbursement month from the selected trips.
- Uses the July form as an embedded Excel template.
- Populates date, odometer start/end, starting/ending location, mileage, reimbursement, month, period end, and totals.
- Allows individual trips to be unchecked before generation.
- Reimbursement formulas reference the editable rate in the generated Excel form.

## Dynamic form sizing

The July reimbursement workbook originally contains 64 trip lines, but the web generator no longer has that limit. When the form is generated, the trip section is rebuilt to contain exactly the number of selected trips. Totals, signature/approval rows, merged cells, print area, and formulas are moved automatically. This works for months with fewer than 64 trips as well as months with more than 64 trips.

## GitHub Pages

The site can also be hosted as a static GitHub Pages site. Upload all five files together:

- `index.html`
- `styles.css`
- `app.js`
- `template-data.js`
- `jszip.min.js`

Trip-log data is not transmitted anywhere by this application; processing happens in the browser.
