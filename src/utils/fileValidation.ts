import type { ValidationResult } from '../types/workbook';

export const MAX_XLSX_FILE_SIZE_BYTES = 20 * 1024 * 1024;

const XLSX_ZIP_SIGNATURES = [
  [0x50, 0x4b, 0x03, 0x04],
  [0x50, 0x4b, 0x05, 0x06],
  [0x50, 0x4b, 0x07, 0x08],
] as const;

export function validateFileSelection(fileList: FileList | File[] | null): ValidationResult {
  if (!fileList || fileList.length === 0) {
    return { valid: false, errorCode: 'NO_FILE', message: 'No file selected. Choose a P21 Staging Report to begin.' };
  }

  if (fileList.length > 2) {
    return { valid: false, errorCode: 'MULTIPLE_FILES', message: 'Select no more than two .xlsx staging reports at a time.' };
  }

  for (const file of Array.from(fileList)) {
    const validation = validateSelectedFile(file);

    if (!validation.valid) {
      return validation;
    }
  }

  return { valid: true };
}

export function validateSelectedFile(file: File): ValidationResult {
  if (file.size === 0) {
    return { valid: false, errorCode: 'EMPTY_FILE', message: 'This file is empty. Choose a valid P21 .xlsx export.' };
  }

  if (file.size > MAX_XLSX_FILE_SIZE_BYTES) {
    return { valid: false, errorCode: 'FILE_TOO_LARGE', message: 'This file exceeds the 20 MB limit. Choose a smaller P21 .xlsx export.' };
  }

  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    return { valid: false, errorCode: 'WRONG_EXTENSION', message: 'Only .xlsx files are supported in this build.' };
  }

  return { valid: true };
}

export function hasValidXlsxSignature(buffer: ArrayBuffer): boolean {
  const bytes = new Uint8Array(buffer.slice(0, 4));

  return XLSX_ZIP_SIGNATURES.some((signature) =>
    signature.every((expectedByte, index) => bytes[index] === expectedByte),
  );
}
