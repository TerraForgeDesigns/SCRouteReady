export type ApplicationState = 'idle' | 'dragging' | 'reading' | 'ready' | 'generating' | 'preview' | 'error';

export type FutureApplicationState = ApplicationState | 'printing';

export interface WorkbookMetadata {
  originalFileName: string;
  fileSize: number;
  worksheetNames: string[];
  hasResultsWorksheet: boolean;
  resultsRowCount: number;
  resultsColumnCount: number;
  resultsHeaderNames: string[];
}

export interface ActiveWorkbookSession {
  id: string;
  fileName: string;
  workbook: unknown;
  metadata: WorkbookMetadata;
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
  | 'UNEXPECTED_PARSE_ERROR';

export interface WorkbookReadError {
  code: WorkbookReadErrorCode;
  message: string;
}
