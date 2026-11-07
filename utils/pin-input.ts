import ExcelJS from 'exceljs';

export interface PinInputRow {
  rowNumber: number;
  rawValue: string;
  pin: string;
}

export interface PinInputError {
  rowNumber: number;
  rawValue: string;
  message: string;
}

export interface PinInputResult {
  rows: PinInputRow[];
  errors: PinInputError[];
}

/**
 * Converts common Excel cell values to a readable string.
 * Keeps alphanumeric PINs intact and safely handles formula/rich-text cells.
 */
function cellToString(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') {
      return value.text.trim();
    }

    if ('result' in value && value.result !== undefined) {
      return String(value.result).trim();
    }

    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map(part => part.text).join('').trim();
    }

    return '';
  }

  return String(value).trim();
}

/**
 * Validates a PIN without enforcing a 13-digit numeric format.
 *
 * Accepted examples:
 *   47000FI0132A0
 *   1234567890123
 *   ABC123XYZ
 *
 * This is a general input-quality check, not an official York County
 * PIN-format specification.
 */
export function validatePin(pin: unknown): {
  isValid: boolean;
  message: string;
} {
  if (pin === null || pin === undefined) {
    return {
      isValid: false,
      message: 'PIN cannot be null.'
    };
  }

  const value = String(pin).trim();

  if (!value) {
    return {
      isValid: false,
      message: 'PIN cannot be empty.'
    };
  }

  if (!/[a-zA-Z0-9]/.test(value)) {
    return {
      isValid: false,
      message: 'PIN must contain at least one letter or number.'
    };
  }

  if (!/^[a-zA-Z0-9]+$/.test(value)) {
    return {
      isValid: false,
      message: 'PIN contains unsupported special characters. Use letters and numbers only.'
    };
  }

  const placeholders = new Set([
    'test',
    'testing',
    'unknown',
    'null',
    'undefined',
    'na',
    'n/a',
    'none',
    'sample',
    'dummy'
  ]);

  if (placeholders.has(value.toLowerCase())) {
    return {
      isValid: false,
      message: 'Unrealistic or placeholder PIN.'
    };
  }

  if (/^(.)\1+$/.test(value)) {
    return {
      isValid: false,
      message: 'PIN cannot contain only repeated characters.'
    };
  }

  return {
    isValid: true,
    message: ''
  };
}

/**
 * Reads PINs from the configured worksheet and PIN column.
 *
 * Completely blank worksheet rows are ignored. If a row contains data in
 * another column but its PIN cell is blank, that row is reported as invalid.
 */
export async function readPinsFromExcel(
  filePath: string,
  sheetName = 'PINs',
  columnHeader = 'PIN'
): Promise<PinInputResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet(sheetName);

  if (!worksheet) {
    throw new Error(
      `Input worksheet "${sheetName}" was not found in "${filePath}".`
    );
  }

  const headerRow = worksheet.getRow(1);
  let pinColumn = 0;

  headerRow.eachCell((cell, colNumber) => {
    if (
      cellToString(cell.value).toLowerCase() === columnHeader.trim().toLowerCase()
    ) {
      pinColumn = colNumber;
    }
  });

  if (!pinColumn) {
    throw new Error(
      `Column header "${columnHeader}" was not found in row 1 of worksheet "${sheetName}".`
    );
  }

  const rows: PinInputRow[] = [];
  const errors: PinInputError[] = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
    const worksheetRow = worksheet.getRow(rowNumber);

    // Ignore fully blank rows, but validate rows containing data elsewhere
    // when their PIN cell is empty.
    let rowHasData = false;

    worksheetRow.eachCell({ includeEmpty: false }, cell => {
      if (cellToString(cell.value) !== '') {
        rowHasData = true;
      }
    });

    if (!rowHasData) {
      continue;
    }

    const rawValue = cellToString(worksheetRow.getCell(pinColumn).value);
    const validation = validatePin(rawValue);

    if (!validation.isValid) {
      errors.push({
        rowNumber,
        rawValue,
        message: validation.message
      });
      continue;
    }

    // Keep the PIN as text; do not strip letters or force a numeric format.
    const pin = rawValue.trim();

    rows.push({
      rowNumber,
      rawValue,
      pin
    });
  }

  return { rows, errors };
}
