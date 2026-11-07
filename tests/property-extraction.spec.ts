import { test, expect } from '@playwright/test';
import { randomBytes } from 'node:crypto';
import path from 'node:path';
import { PropertyViewerPage, ExtractedPropertyRow } from '../pages/property-viewer.page';
import { readPinsFromExcel } from '../utils/pin-input';
import { ExcelReporter, SummaryRecord, ResultStatus } from '../utils/excel-reporter';

const BASE_URL = process.env.BASE_URL;
const INPUT_FILE = process.env.INPUT_FILE || path.resolve('test-data/PIN_Input.xlsx');
const INPUT_SHEET = process.env.INPUT_SHEET || 'PINs';
const OUTPUT_FILE = process.env.OUTPUT_FILE || 'YorkCounty-PropertyDetails.xlsx';

function createUniqueOutputFileName(fileName: string): string {
  const parsedPath = path.parse(fileName);
  const extension = parsedPath.ext || '.xlsx';
  const uniqueSuffix = randomBytes(4).toString('hex');
  const uniqueName = `${parsedPath.name}-${uniqueSuffix}${extension}`;

  return parsedPath.dir ? path.join(parsedPath.dir, uniqueName) : uniqueName;
}

function normalizeDigits(value: string): string {
  return value.replace(/\D/g, '');
}

/**
 * Validates the returned parcel ID when the PIN format supports a reliable
 * comparison. The existing 13-digit numeric prefix rule is retained for
 * numeric PINs. For alphanumeric PINs, the parcel-ID mapping is not assumed;
 * the presence of the parcel ID field is still checked.
 */
function validateReturnedParcelId(
  pin: string,
  rows: ExtractedPropertyRow[]
): string | null {
  const parcelIdRow = rows.find(row =>
    /18 Digit Parcel ID Number/i.test(row.field)
  );

  if (!parcelIdRow || !parcelIdRow.value?.trim()) {
    return 'Could not validate the result: "18 Digit Parcel ID Number" was missing or empty in extracted data.';
  }

  // Keep the existing prefix validation for numeric PINs only.
  if (/^\d+$/.test(pin)) {
    const displayedDigits = normalizeDigits(parcelIdRow.value);

    if (!displayedDigits.startsWith(pin)) {
      return `PIN mismatch. Requested ${pin}, but the displayed parcel ID was "${parcelIdRow.value}".`;
    }
  }

  // Do not apply a numeric-only comparison to alphanumeric PINs.
  // Confirm the exact alphanumeric PIN-to-parcel-ID mapping before adding
  // a comparison rule for those PINs.

  return null;
}

test('Extract property details for all PINs in the input workbook', async ({ page }, testInfo) => {
  test.setTimeout(10 * 60 * 1000);

  if (!BASE_URL || BASE_URL.includes('REPLACE-WITH')) {
    throw new Error('Set BASE_URL in your .env file to the York County Property Viewer 2.0 URL.');
  }

  const input = await readPinsFromExcel(INPUT_FILE, INPUT_SHEET);
  const reporter = new ExcelReporter();
  const summary: SummaryRecord[] = [];

  // Add every invalid Excel input to the Summary worksheet.
  for (const invalid of input.errors) {
    const now = new Date().toISOString();

    summary.push({
      inputRow: invalid.rowNumber,
      pin: invalid.rawValue,
      status: 'Invalid PIN',
      records: 0,
      message: invalid.message,
      startedAt: now,
      completedAt: now
    });
  }

  // If the worksheet has no usable PIN entries and no validation errors,
  // it is likely empty or contains no data rows.
  if (input.rows.length === 0 && input.errors.length === 0) {
    throw new Error(
      `No PIN entries were found in worksheet "${INPUT_SHEET}" of ${INPUT_FILE}. Check the PIN column and Excel input data.`
    );
  }

  const seenPins = new Set<string>();

  // Do not open the browser when there are no valid PINs.
  // Invalid rows will still be saved in the output workbook below.
  if (input.rows.length > 0) {
    const viewer = new PropertyViewerPage(page);
    await viewer.open(BASE_URL);

    for (const item of input.rows) {
      const startedAt = new Date().toISOString();

      // Treat PINs with different letter casing as duplicates.
      const duplicateKey = item.pin.trim().toUpperCase();

      if (seenPins.has(duplicateKey)) {
        const now = new Date().toISOString();

        summary.push({
          inputRow: item.rowNumber,
          pin: item.pin,
          status: 'Duplicate',
          records: 0,
          message: 'Duplicate PIN in input workbook; skipped to avoid duplicate worksheet names and repeated extraction.',
          startedAt,
          completedAt: now
        });

        continue;
      }

      seenPins.add(duplicateKey);

      await test.step(`Extract PIN ${item.pin}`, async () => {
        let rows: ExtractedPropertyRow[] = [];
        let status: ResultStatus = 'Passed';
        let message = '';

        try {
          await viewer.searchAndSelectPin(item.pin);
          rows = await viewer.extractPropertyRows();

          const validationError = validateReturnedParcelId(item.pin, rows);

          if (validationError) {
            status = 'Failed';
            message = validationError;
          }

          reporter.addPinWorksheet(item.pin, rows, status, message);
        } catch (error) {
          status = 'Failed';
          message = error instanceof Error ? error.message : String(error);

          reporter.addFailureWorksheet(item.pin, message);

          // Attach a screenshot for diagnosis, but continue with the next PIN.
          const screenshot = await page.screenshot({ fullPage: true }).catch(() => undefined);

          if (screenshot) {
            await testInfo.attach(`failure-${item.pin}.png`, {
              body: screenshot,
              contentType: 'image/png'
            });
          }
        }

        const completedAt = new Date().toISOString();

        summary.push({
          inputRow: item.rowNumber,
          pin: item.pin,
          status,
          records: rows.length,
          message,
          startedAt,
          completedAt
        });
      });
    }
  }

  // Save the Summary worksheet, including invalid PIN rows, even when
  // there were no valid PINs to process.
  reporter.addSummary(summary);

  const outputFileName = createUniqueOutputFileName(OUTPUT_FILE);
  const outputPath = testInfo.outputPath(outputFileName);

  await reporter.save(outputPath);

  await testInfo.attach(outputFileName, {
    path: outputPath,
    contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const failedCount = summary.filter(row =>
    row.status === 'Failed' || row.status === 'Invalid PIN'
  ).length;

  const invalidCount = input.errors.length;

  console.log(
    `Extraction complete. ` +
    `Valid PINs: ${input.rows.length}; ` +
    `Invalid PINs: ${invalidCount}; ` +
    `Total summary rows: ${summary.length}; ` +
    `Failures/invalid: ${failedCount}; ` +
    `Output: ${outputPath}`
  );

  // Preserve the report even when some PINs are invalid or extraction fails.
  // The test is marked failed at the end so CI reflects the unsuccessful run.
  expect(
    failedCount,
    `${failedCount} PIN(s) failed validation or extraction. ` +
    `${invalidCount > 0 ? 'Invalid PINs were found. ' : ''}` +
    'Review the attached workbook Summary sheet.'
  ).toBe(0);
});
