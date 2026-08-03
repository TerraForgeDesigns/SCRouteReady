import * as XLSX from 'xlsx';
import type { ActiveWorkbookSession, WorkbookMetadata, WorkbookReadError } from '../types/workbook';
import { hasValidXlsxSignature, validateSelectedFile } from '../utils/fileValidation';

export class WorkbookReaderError extends Error {
  readonly code: WorkbookReadError['code'];

  constructor(code: WorkbookReadError['code'], message: string) {
    super(message);
    this.name = 'WorkbookReaderError';
    this.code = code;
  }
}

export async function readWorkbookSession(file: File): Promise<ActiveWorkbookSession> {
  const validation = validateSelectedFile(file);

  if (!validation.valid) {
    throw new WorkbookReaderError(validation.errorCode ?? 'UNEXPECTED_PARSE_ERROR', validation.message ?? 'The selected file could not be validated.');
  }

  const sourceBuffer = await file.arrayBuffer();

  if (!hasValidXlsxSignature(sourceBuffer)) {
    throw new WorkbookReaderError('INVALID_FILE_SIGNATURE', 'This file does not appear to be a valid .xlsx workbook.');
  }

  let workbook: XLSX.WorkBook;

  try {
    if (typeof XLSX.read !== 'function') {
      throw new Error('SheetJS read function is unavailable.');
    }

    const parseBuffer = sourceBuffer.slice(0);
    workbook = XLSX.read(parseBuffer, {
      cellFormula: false,
      cellHTML: false,
      cellNF: false,
      cellStyles: false,
      bookDeps: false,
      bookFiles: false,
      bookProps: false,
      bookSheets: false,
      bookVBA: false,
      WTF: false,
    });
  } catch {
    throw new WorkbookReaderError('WORKBOOK_PARSE_FAILED', 'An unexpected parsing error occurred. Try another P21 .xlsx export.');
  }

  let metadata: WorkbookMetadata;

  try {
    metadata = getWorkbookMetadata(file, workbook);
  } catch {
    throw new WorkbookReaderError('RESULTS_SHEET_UNREADABLE', 'An unexpected parsing error occurred. Try another P21 .xlsx export.');
  }

  if (metadata.worksheetNames.length === 0) {
    throw new WorkbookReaderError('WORKBOOK_HAS_NO_SHEETS', 'This workbook does not contain any worksheets.');
  }

  if (!metadata.hasResultsWorksheet) {
    throw new WorkbookReaderError('RESULTS_SHEET_MISSING', 'This workbook does not contain the expected Results worksheet. Confirm that you selected the P21 route export.');
  }

  return {
    id: crypto.randomUUID(),
    fileName: file.name,
    workbook,
    metadata,
  };
}

function getWorkbookMetadata(file: File, workbook: XLSX.WorkBook): WorkbookMetadata {
  const worksheetNames = Array.isArray(workbook.SheetNames) ? [...workbook.SheetNames] : [];
  const resultsName = worksheetNames.find((sheetName) => sheetName.toLowerCase() === 'results');
  const resultsSheet = resultsName ? workbook.Sheets[resultsName] : undefined;
  const range = resultsSheet?.['!ref'] ? XLSX.utils.decode_range(resultsSheet['!ref']) : null;
  const headerNames = resultsSheet && range ? readHeaderNames(resultsSheet, range) : [];
  const totalUsedRows = range ? range.e.r - range.s.r + 1 : 0;

  return {
    originalFileName: file.name,
    fileSize: file.size,
    worksheetNames,
    hasResultsWorksheet: Boolean(resultsName),
    resultsRowCount: Math.max(0, totalUsedRows - 1),
    resultsColumnCount: range ? range.e.c - range.s.c + 1 : 0,
    resultsHeaderNames: headerNames,
  };
}

function readHeaderNames(sheet: XLSX.WorkSheet, range: XLSX.Range): string[] {
  const headers: string[] = [];
  const headerRow = range.s.r;

  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: column });
    const value = sheet[cellAddress]?.v;

    if (typeof value === 'string' && value.trim().length > 0) {
      headers.push(value.trim());
    } else if (typeof value === 'number' || typeof value === 'boolean') {
      headers.push(String(value));
    }
  }

  return headers;
}
