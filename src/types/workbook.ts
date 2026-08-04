export type ApplicationState = 'idle' | 'dragging' | 'reading' | 'ready' | 'generating' | 'preview' | 'error';

export type FutureApplicationState = ApplicationState | 'printing';

export type ReportKind = 'sales' | 'transfer';

export interface WorkbookMetadata {
  originalFileName: string;
  fileSize: number;
  worksheetNames: string[];
  hasResultsWorksheet: boolean;
  resultsRowCount: number;
  resultsColumnCount: number;
  resultsHeaderNames: string[];
  routeCount: number;
  stopCount: number;
}

export interface ActiveWorkbookSession {
  id: string;
  fileName: string;
  workbook: unknown;
  metadata: WorkbookMetadata;
  reportKind: ReportKind;
}

export type ValidationErrorCode =
  | 'NO_FILE'
  | 'MULTIPLE_FILES'
  | 'WRONG_EXTENSION'
  | 'EMPTY_FILE'
  | 'FILE_TOO_LARGE'
  | 'INVALID_FILE_SIGNATURE';

export interface ValidationResult {
  valid: boolean;
  errorCode?: ValidationErrorCode;
  message?: string;
}

export type WorkbookReadErrorCode =
  | ValidationErrorCode
  | 'WORKBOOK_PARSE_FAILED'
  | 'WORKBOOK_HAS_NO_SHEETS'
  | 'RESULTS_SHEET_MISSING'
  | 'RESULTS_SHEET_UNREADABLE'
  | 'UNSUPPORTED_REPORT'
  | 'AMBIGUOUS_REPORT'
  | 'DUPLICATE_REPORT_KIND'
  | 'UNEXPECTED_PARSE_ERROR';

export interface WorkbookReadError {
  code: WorkbookReadErrorCode;
  message: string;
}
