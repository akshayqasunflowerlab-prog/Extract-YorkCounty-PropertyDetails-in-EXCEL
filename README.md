# York County Property Extraction Automation Framework

Reusable Playwright + TypeScript framework that reads 13-digit PINs from an Excel input workbook, searches the York County Property Viewer 2.0, extracts property table data, and creates one output workbook containing a Summary sheet plus one worksheet per unique PIN.

## 1. Requirements

- Node.js 20 or later
- npm
- Access to York County Property Viewer 2.0
- Microsoft Excel or another `.xlsx` editor

## 2. Install

```bash
npm install
npx playwright install chromium
```

## 3. Configure the application URL

Copy `.env.example` to `.env` and set the actual property viewer URL:

```env
BASE_URL=https://your-actual-york-county-property-viewer-url
HEADLESS=false
```

Do not include the placeholder URL. `HEADLESS=false` opens the browser while you are debugging.

## 4. Prepare the input workbook

Create `test-data/PIN_Input.xlsx` with a worksheet named `PINs`:

| PIN |
|---|
| 2600007003700 |
| 2600007003800 |
| 2600007003900 |

- Put `PIN` in cell A1.
- Put one 13-digit PIN per row, starting at A2.
- Format the PIN column as **Text** in Excel. This avoids Excel converting identifiers to scientific notation or changing digits.
- Blank rows are ignored.
- Invalid values are included in the output Summary as `Invalid PIN`.
- Duplicate valid PINs are skipped and marked `Duplicate`.

Optional `.env` settings:

```env
INPUT_FILE=test-data/PIN_Input.xlsx
INPUT_SHEET=PINs
OUTPUT_FILE=YorkCounty-PropertyDetails.xlsx
```

`OUTPUT_FILE` sets the output workbook's base name. A unique 8-character
hexadecimal suffix is added before the `.xlsx` extension on every run, for
example `YorkCounty-PropertyDetails-9f7418a2.xlsx`.

## 5. Run

```bash
npm test
```

For a visible browser:

```bash
npm run test:headed
```

For Playwright Inspector:

```bash
npm run test:debug
```

The generated workbook is saved with a unique filename in the Playwright test output directory and attached to the HTML/Playwright test report. Find the exact location in the terminal output after the run. The HTML report is opened with:

```bash
npm run report
```

## 6. Output workbook

### Summary worksheet

Contains input row, PIN, status, number of extracted records, error/validation message, start time, and completion time.

Statuses:
- `Passed`: extraction and parcel-ID validation passed.
- `Failed`: search, extraction, or parcel-ID validation failed.
- `Invalid PIN`: input value was not exactly 13 digits.
- `Duplicate`: PIN already occurred earlier in the input workbook and was skipped.

### PIN-specific worksheets

Each unique valid PIN gets its own worksheet named with the 13-digit PIN. Data is stored as:

| Section | Field | Value |
|---|---|---|
| Property Information | Property Address | ... |
| Assessed Values | Assessed Total Value | ... |
| Tax Information | Total Taxes | ... |

If extraction fails, the PIN worksheet contains the failure message. The Summary sheet always contains the run status.

## 7. Framework design

- `pages/property-viewer.page.ts`: reusable page object for navigation, search, selection, and table extraction.
- `utils/pin-input.ts`: input workbook reader and 13-digit PIN validation.
- `utils/excel-reporter.ts`: workbook creation, worksheet formatting, and output.
- `tests/property-extraction.spec.ts`: orchestration, duplicate handling, per-PIN error handling, result validation, screenshots, and report attachment.

## 8. Important application-specific notes

1. The page object uses the working locator flow established during debugging:
   - Search textbox named `Search`
   - Suggestion group named `PIDN (13-Digit)`
   - Exact option name equal to the PIN
   - Click the suggestion only once. ArcGIS can detach the selected option during re-render.
2. The extractor reads rows from `arcgis-feature-fields` tables. It captures tabular property fields; non-tabular text, charts, images, and links are not included.
3. The parcel-ID validation assumes the normalized 18-digit displayed parcel ID starts with the 13-digit search PIN, consistent with the supplied example. Confirm this mapping against additional real records before relying on it for production acceptance.
4. The framework processes PINs sequentially in one browser page. This is intentional because the map viewer may retain state between searches.
5. If the application changes its accessible names or DOM structure, update only the page object rather than duplicating locators in the test.
6. The site may throttle or change its search behavior. If the second and subsequent searches do not work in the live application, inspect whether the feature panel needs to be closed or reset before adding a close-panel action; no unverified close button is assumed here.

## 9. Troubleshooting

- **Search suggestion timeout:** Confirm the PIN exists, the page is fully loaded, and the group/option accessible names still match. Check the trace in `test-results`.
- **Property Information heading timeout:** Confirm the suggestion was selected and the feature panel opened. Do not add a second click to the same suggestion.
- **No table rows found:** Inspect the live DOM. The page object currently expects open-shadow-DOM `arcgis-feature-fields` components containing HTML tables.
- **Input worksheet/header error:** Confirm worksheet name `PINs`, header `PIN` in row 1, and values below it.
- **BASE_URL error:** Update `.env` with the actual URL.

## 10. Suggested next enhancement

After confirming repeated searches work on the live viewer, add optional CSV logging, retry policy for transient UI failures, and a schema-change alert for missing expected fields.
