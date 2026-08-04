import * as XLSX from 'xlsx';
import type { ActiveWorkbookSession, WorkbookMetadata, WorkbookReadError } from '../types/workbook';
import { hasValidXlsxSignature, validateSelectedFile } from '../utils/fileValidation';
import { normalizeHeader } from '../utils/headerNormalization';
import { normalizeComparableText, sanitizeDisplayText } from '../utils/sanitizeDisplayText';
import { detectReportKind } from './reportDetector';

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

  const detection = detectReportKind(metadata.resultsHeaderNames);

  if (!detection.ok) {
    throw new WorkbookReaderError(
      detection.reason === 'ambiguous' ? 'AMBIGUOUS_REPORT' : 'UNSUPPORTED_REPORT',
      detection.reason === 'ambiguous'
        ? 'This workbook matches more than one supported P21 report type, so SCRouteReady stopped safely.'
        : 'This workbook is not a supported P21 Sales Order or Transfer Staging Report.',
    );
  }

  return {
    id: crypto.randomUUID(),
    fileName: file.name,
    workbook,
    metadata: addReportCounts(metadata, workbook, detection.kind),
    reportKind: detection.kind,
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
    routeCount: 0,
    stopCount: 0,
  };
}

function addReportCounts(metadata: WorkbookMetadata, workbook: XLSX.WorkBook, kind: 'sales' | 'transfer'): WorkbookMetadata {
  const resultsName = workbook.SheetNames.find((sheetName) => sheetName.toLowerCase() === 'results');
  const sheet = resultsName ? workbook.Sheets[resultsName] : undefined;
  const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;

  if (!sheet || !range) {
    return metadata;
  }

  const headerRow = range.s.r;
  const routeColumn = findHeaderColumn(sheet, range, 'Route from Descartes');
  const stopColumn = findHeaderColumn(sheet, range, kind === 'sales' ? 'Stop #' : 'Assigned Stop');

  return {
    ...metadata,
    routeCount: countDistinctValues(sheet, range, routeColumn, headerRow),
    stopCount: countDistinctValues(sheet, range, stopColumn, headerRow),
  };
}

function findHeaderColumn(sheet: XLSX.WorkSheet, range: XLSX.Range, headerName: string): number | null {
  const target = normalizeHeader(headerName);

  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: column });
    const value = sheet[cellAddress]?.v;

    if (normalizeHeader(value) === target) {
      return column;
    }
  }

  return null;
}

function countDistinctValues(sheet: XLSX.WorkSheet, range: XLSX.Range, column: number | null, headerRow: number): number {
  if (column === null) {
    return 0;
  }

  const values = new Set<string>();

  for (let row = headerRow + 1; row <= range.e.r; row += 1) {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: column });
    const normalized = normalizeComparableText(sanitizeDisplayText(sheet[cellAddress]?.w ?? sheet[cellAddress]?.v ?? ''));

    if (normalized) {
      values.add(normalized);
    }
  }

  return values.size;
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
