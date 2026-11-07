import ExcelJS from 'exceljs';
import path from 'node:path';
import fs from 'node:fs/promises';
import { ExtractedPropertyRow } from '../pages/property-viewer.page';

export type ResultStatus = 'Passed' | 'Failed' | 'Invalid PIN' | 'Duplicate';

export interface SummaryRecord {
  inputRow: number;
  pin: string;
  status: ResultStatus;
  records: number;
  message: string;
  startedAt: string;
  completedAt: string;
}

export class ExcelReporter {
  private readonly workbook = new ExcelJS.Workbook();

  constructor() {
    this.workbook.creator = 'York County Property Extraction Automation';
    this.workbook.created = new Date();
  }

  addSummary(records: SummaryRecord[]): void {
    const sheet = this.workbook.addWorksheet('Summary');
    sheet.columns = [
      { header: 'Input Row', key: 'inputRow', width: 12 },
      { header: 'PIN', key: 'pin', width: 18 },
      { header: 'Status', key: 'status', width: 16 },
      { header: 'Records Extracted', key: 'records', width: 20 },
      { header: 'Message / Error', key: 'message', width: 65 },
      { header: 'Started At', key: 'startedAt', width: 25 },
      { header: 'Completed At', key: 'completedAt', width: 25 }
    ];
    for (const record of records) sheet.addRow(record);
    this.styleSheet(sheet);
    sheet.autoFilter = { from: 'A1', to: 'G1' };
  }

  addPinWorksheet(pin: string, rows: ExtractedPropertyRow[], status: ResultStatus, message = ''): void {
    // PINs are 13 digits, within Excel's 31-character worksheet-name limit.
    const sheet = this.workbook.addWorksheet(pin);
    sheet.addRow(['Section', 'Field', 'Value']);
    for (const row of rows) sheet.addRow([row.section, row.field, row.value]);
    if (message) {
      sheet.addRow([]);
      sheet.addRow(['Extraction Status', status, message]);
    }
    sheet.columns = [
      { width: 28 },
      { width: 38 },
      { width: 65 }
    ];
    this.styleSheet(sheet);
    sheet.autoFilter = { from: 'A1', to: 'C1' };
  }

  addFailureWorksheet(pin: string, message: string): void {
    this.addPinWorksheet(pin, [], 'Failed', message);
  }

  private styleSheet(sheet: ExcelJS.Worksheet): void {
    const header = sheet.getRow(1);
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
    header.alignment = { vertical: 'middle' };
    header.height = 22;
    sheet.views = [{ state: 'frozen', ySplit: 1 }];
    sheet.eachRow(row => {
      row.alignment = { vertical: 'top', wrapText: true };
    });
  }

  async save(filePath: string): Promise<void> {
    await fs.mkdir(path.dirname(filePath), { recursive: true });
    await this.workbook.xlsx.writeFile(filePath);
  }
}
