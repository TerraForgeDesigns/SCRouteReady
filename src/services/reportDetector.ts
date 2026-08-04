import type { ReportKind } from '../types/workbook';
import { normalizeHeader } from '../utils/headerNormalization';

export const SALES_REQUIRED_HEADERS = [
  'Route from Descartes',
  'Customer Type',
  'Stop #',
  'Pick Ticket No',
  'Po No',
  'Qty',
  'UOM',
  'Bin',
  'DMAS Part#',
  'P21 Item Id',
  'Item Desc',
  'Ship To Name',
  'Ship To City',
  'Ship To State',
] as const;

export const TRANSFER_REQUIRED_HEADERS = [
  'Route from Descartes',
  'Assigned Stop',
  'Transfer No',
  'Qty',
  'Unit of Measure',
  'Bin Cd',
  'Item Id',
  'Item Desc',
  'To Location Name',
] as const;

export type ReportDetectionResult =
  | { ok: true; kind: ReportKind }
  | { ok: false; reason: 'unsupported' | 'ambiguous' };

export function detectReportKind(headers: readonly string[]): ReportDetectionResult {
  const normalizedHeaders = new Set(headers.map((header) => normalizeHeader(header)).filter(Boolean));
  const matchesSales = hasHeaders(normalizedHeaders, SALES_REQUIRED_HEADERS);
  const matchesTransfer = hasHeaders(normalizedHeaders, TRANSFER_REQUIRED_HEADERS);

  if (matchesSales && matchesTransfer) {
    return { ok: false, reason: 'ambiguous' };
  }

  if (matchesSales) {
    return { ok: true, kind: 'sales' };
  }

  if (matchesTransfer) {
    return { ok: true, kind: 'transfer' };
  }

  return { ok: false, reason: 'unsupported' };
}

function hasHeaders(normalizedHeaders: Set<string>, requiredHeaders: readonly string[]): boolean {
  return requiredHeaders.every((header) => normalizedHeaders.has(normalizeHeader(header)));
}
