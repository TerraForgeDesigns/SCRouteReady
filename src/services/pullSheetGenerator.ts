import * as XLSX from 'xlsx';
import type {
  PalletKeyEntry,
  PalletKeyPage,
  PullSheetColumn,
  PullSheetDocument,
  PullSheetRecord,
  PullSheetRouteSection,
  SourceColumnDefinition,
  SourceField,
  StopBanner,
} from '../types/pullSheet';
import { PullSheetGenerationError } from '../types/pullSheet';
import { compareDesignations, getDesignationLetter, getKeyShadeClass } from '../utils/designationLetters';
import { normalizeHeader } from '../utils/headerNormalization';
import { compareNaturalText, parseSortableNumber } from '../utils/naturalSort';
import { normalizeComparableText, sanitizeDisplayText } from '../utils/sanitizeDisplayText';
import { getStopColorClass } from '../utils/stopColors';

const REQUIRED_COLUMNS: SourceColumnDefinition[] = [
  { field: 'routeFromDescartes', sourceHeader: 'Route from Descartes', displayLabel: 'Route from Descartes', required: true },
  { field: 'customerType', sourceHeader: 'Customer Type', displayLabel: 'Customer Type', required: true },
  { field: 'stopNumber', sourceHeader: 'Stop #', displayLabel: 'Stop #', required: true },
  { field: 'pickTicketNo', sourceHeader: 'Pick Ticket No', displayLabel: 'Pick Ticket No', required: true },
  { field: 'poNo', sourceHeader: 'Po No', displayLabel: 'PO No', required: true },
  { field: 'qty', sourceHeader: 'Qty', displayLabel: 'Qty', required: true },
  { field: 'uom', sourceHeader: 'UOM', displayLabel: 'UOM', required: true },
  { field: 'bin', sourceHeader: 'Bin', displayLabel: 'Bin', required: true },
  { field: 'dmasPart', sourceHeader: 'DMAS Part#', displayLabel: 'DMAS Part #', required: true },
  { field: 'p21ItemId', sourceHeader: 'P21 Item Id', displayLabel: 'P21 Item ID', required: true },
  { field: 'itemDescription', sourceHeader: 'Item Desc', displayLabel: 'Item Description', required: true },
  { field: 'shipToName', sourceHeader: 'Ship To Name', displayLabel: 'Ship To Name', required: true },
  { field: 'shipToCity', sourceHeader: 'Ship To City', displayLabel: 'Ship To City', required: true },
  { field: 'shipToState', sourceHeader: 'Ship To State', displayLabel: 'State', required: true },
];

const BASE_COLUMNS: PullSheetColumn[] = [
  { key: 'stopDisplay', label: 'Stop', className: 'column-stopDisplay', widthPercent: 4.5, align: 'center', wrap: 'nowrap' },
  { key: 'customerType', label: 'Type', className: 'column-customerType', widthPercent: 6.5, align: 'center', wrap: 'nowrap' },
  { key: 'pickTicketNo', label: 'Pick Ticket', className: 'column-pickTicketNo', widthPercent: 11, align: 'center', wrap: 'nowrap' },
  { key: 'poNo', label: 'PO No.', className: 'column-poNo', widthPercent: 10, align: 'center', wrap: 'nowrap' },
  { key: 'qty', label: 'Qty', className: 'column-qty', widthPercent: 4, align: 'right', wrap: 'nowrap' },
  { key: 'uom', label: 'UOM', className: 'column-uom', widthPercent: 5, align: 'center', wrap: 'nowrap' },
  { key: 'bin', label: 'Bin', className: 'column-bin', widthPercent: 6, align: 'center', wrap: 'nowrap' },
  { key: 'dmasPart', label: 'DMAS Part', className: 'column-dmasPart', widthPercent: 11, align: 'left', wrap: 'nowrap' },
  { key: 'p21ItemId', label: 'P21 Item ID', className: 'column-p21ItemId', widthPercent: 11, align: 'left', wrap: 'nowrap' },
  { key: 'itemDescription', label: 'Item Description', className: 'column-itemDescription', widthPercent: 31, align: 'left', wrap: 'description' },
];

const KEY_COLUMN: PullSheetColumn = {
  key: 'keyDesignation',
  label: 'Key',
  className: 'column-keyDesignation',
  widthPercent: 4,
  align: 'center',
  wrap: 'nowrap',
};

const KEYED_COLUMN_WIDTHS = new Map<string, number>([
  ['stopDisplay', 4],
  ['keyDesignation', 4],
  ['customerType', 6],
  ['pickTicketNo', 10.5],
  ['poNo', 9.5],
  ['qty', 4],
  ['uom', 5],
  ['bin', 6],
  ['dmasPart', 10.5],
  ['p21ItemId', 10.5],
  ['itemDescription', 30],
]);

type ColumnIndexMap = Record<SourceField, number>;
type RawRow = string[];

interface StopCustomer {
  key: string;
  displayName: string;
  city: string;
  state: string;
}

interface StopAssignment {
  requiresKey: boolean;
  customerLetters: Map<string, string>;
  customers: StopCustomer[];
}

export function generatePullSheetDocument(workbook: unknown): PullSheetDocument {
  try {
    const typedWorkbook = assertWorkbook(workbook);
    const sheetName = typedWorkbook.SheetNames.find((name) => name.toLowerCase() === 'results');

    if (!sheetName) {
      throw new PullSheetGenerationError('UNEXPECTED_WORKSHEET_STRUCTURE', 'The Results worksheet could not be found.');
    }

    const sheet = typedWorkbook.Sheets[sheetName];
    const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;

    if (!sheet || !range) {
      throw new PullSheetGenerationError('RESULTS_WORKSHEET_EMPTY', 'The Results worksheet is empty.');
    }

    const rawRows = extractRows(sheet, range);
    const [headerRow, ...dataRows] = rawRows;

    if (!headerRow || dataRows.length === 0) {
      throw new PullSheetGenerationError('NO_USABLE_DATA_ROWS', 'The Results worksheet does not contain usable data rows.');
    }

    const columnMap = buildColumnMap(headerRow);
    const records = convertRows(dataRows, columnMap).filter((record) => hasUsableData(record));

    if (records.length === 0) {
      throw new PullSheetGenerationError('NO_USABLE_DATA_ROWS', 'The Results worksheet does not contain usable data rows.');
    }

    return buildDocument(records);
  } catch (error) {
    if (error instanceof PullSheetGenerationError) {
      throw error;
    }

    throw new PullSheetGenerationError('TRANSFORMATION_FAILED', 'The Pull Sheet could not be generated from this worksheet.');
  }
}

function assertWorkbook(workbook: unknown): XLSX.WorkBook {
  if (!workbook || typeof workbook !== 'object') {
    throw new PullSheetGenerationError('UNEXPECTED_WORKSHEET_STRUCTURE', 'The active workbook is unavailable.');
  }

  const candidate = workbook as Partial<XLSX.WorkBook>;

  if (!Array.isArray(candidate.SheetNames) || !candidate.Sheets) {
    throw new PullSheetGenerationError('UNEXPECTED_WORKSHEET_STRUCTURE', 'The active workbook is unavailable.');
  }

  return candidate as XLSX.WorkBook;
}

function extractRows(sheet: XLSX.WorkSheet, range: XLSX.Range): RawRow[] {
  const rows: RawRow[] = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row: string[] = [];

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[cellAddress];
      row.push(sanitizeDisplayText(cell?.w ?? cell?.v ?? ''));
    }

    rows.push(row);
  }

  return rows;
}

function buildColumnMap(headerRow: RawRow): ColumnIndexMap {
  const normalizedHeaderIndexes = new Map<string, number>();

  headerRow.forEach((header, index) => {
    const normalized = normalizeHeader(header);

    if (normalized && !normalizedHeaderIndexes.has(normalized)) {
      normalizedHeaderIndexes.set(normalized, index);
    }
  });

  const missingColumns: string[] = [];
  const map = {} as Partial<ColumnIndexMap>;

  for (const column of REQUIRED_COLUMNS) {
    const index = normalizedHeaderIndexes.get(normalizeHeader(column.sourceHeader));

    if (index === undefined) {
      missingColumns.push(column.displayLabel);
    } else {
      map[column.field] = index;
    }
  }

  if (missingColumns.length > 0) {
    throw new PullSheetGenerationError(
      'REQUIRED_HEADERS_MISSING',
      'This P21 export is missing columns required to create the Pull Sheet.',
      missingColumns,
    );
  }

  return map as ColumnIndexMap;
}

function convertRows(rows: RawRow[], columnMap: ColumnIndexMap): PullSheetRecord[] {
  return rows.map((row, index) => {
    const routeDisplay = getCell(row, columnMap.routeFromDescartes);
    const stopDisplay = getCell(row, columnMap.stopNumber);
    const pickTicketNo = getCell(row, columnMap.pickTicketNo);
    const shipToName = getCell(row, columnMap.shipToName);

    return {
      sourceIndex: index,
      routeKey: normalizeRouteKey(routeDisplay),
      routeDisplay,
      customerType: getCell(row, columnMap.customerType),
      stopKey: normalizeStopKey(stopDisplay),
      stopDisplay,
      stopSortValue: parseSortableNumber(stopDisplay),
      pickTicketNo,
      pickTicketSortValue: parseSortableNumber(pickTicketNo),
      poNo: getCell(row, columnMap.poNo),
      qty: getCell(row, columnMap.qty),
      uom: getCell(row, columnMap.uom),
      bin: getCell(row, columnMap.bin),
      dmasPart: getCell(row, columnMap.dmasPart),
      p21ItemId: getCell(row, columnMap.p21ItemId),
      itemDescription: getCell(row, columnMap.itemDescription),
      shipToName,
      shipToNameKey: normalizeComparableText(shipToName),
      shipToCity: getCell(row, columnMap.shipToCity),
      shipToState: getCell(row, columnMap.shipToState),
      keyDesignation: '',
      keyShadeClass: '',
      stopColorClass: '',
      isFirstRowForStop: false,
      isFirstRowForKeyGroup: false,
      stopBanner: null,
    };
  });
}

function getCell(row: RawRow, index: number): string {
  return sanitizeDisplayText(row[index] ?? '');
}

function hasUsableData(record: PullSheetRecord): boolean {
  return [
    record.routeDisplay,
    record.customerType,
    record.stopDisplay,
    record.pickTicketNo,
    record.poNo,
    record.qty,
    record.uom,
    record.bin,
    record.dmasPart,
    record.p21ItemId,
    record.itemDescription,
    record.shipToName,
    record.shipToCity,
    record.shipToState,
  ].some((value) => value.length > 0);
}

function buildDocument(records: PullSheetRecord[]): PullSheetDocument {
  const routes = groupBy(records, (record) => record.routeKey);
  const routeSections: PullSheetRouteSection[] = [];
  const palletKeyPages: PalletKeyPage[] = [];
  let totalStops = 0;

  for (const routeRecords of routes.values()) {
    const baseSortedRecords = stableSort(routeRecords, compareBasePullSheetRecords);
    const stopKeys = getOrderedUnique(baseSortedRecords.map((record) => record.stopKey));
    const stopColorMap = new Map(stopKeys.map((stopKey, index) => [stopKey, getStopColorClass(index)]));
    const assignments = buildStopAssignments(baseSortedRecords);

    applyStopColorsAndKeys(baseSortedRecords, stopColorMap, assignments);
    const sortedRecords = stableSort(baseSortedRecords, (left, right) => compareFinalPullOrder(left, right, assignments));
    applyRowGroupMetadata(sortedRecords, assignments);

    const routeKey = sortedRecords[0].routeKey;
    const routeDisplay = sortedRecords[0].routeDisplay;
    const routePalletPages = buildPalletKeyPages(routeKey, routeDisplay, sortedRecords, assignments, stopColorMap);
    const hasKeyColumn = routePalletPages.length > 0;
    const columns = hasKeyColumn ? withKeyColumn(BASE_COLUMNS) : BASE_COLUMNS;

    routeSections.push({
      routeKey,
      routeDisplay,
      columns,
      rows: sortedRecords,
      stopCount: stopKeys.length,
      lineCount: sortedRecords.length,
      palletKeyPageCount: routePalletPages.length,
      hasKeyColumn,
    });
    palletKeyPages.push(...routePalletPages);
    totalStops += stopKeys.length;
  }

  return {
    summary: {
      routeCount: routeSections.length,
      stopCount: totalStops,
      productLineCount: records.length,
      palletKeyPageCount: palletKeyPages.length,
    },
    routes: routeSections,
    palletKeyPages,
  };
}

function compareBasePullSheetRecords(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareStopOrder(left, right)
    || compareNormalStopRows(left, right);
}

function compareFinalPullOrder(
  left: PullSheetRecord,
  right: PullSheetRecord,
  assignments: Map<string, StopAssignment>,
): number {
  const stopCompare = compareStopOrder(left, right);

  if (stopCompare !== 0) {
    return stopCompare;
  }

  const assignment = assignments.get(`${left.routeKey}|${left.stopKey}`);

  if (assignment?.requiresKey) {
    return compareKeyedStopRows(left, right);
  }

  return compareNormalStopRows(left, right);
}

function compareStopOrder(left: PullSheetRecord, right: PullSheetRecord): number {
  const stopCompare = compareDescendingNullableNumber(left.stopSortValue, right.stopSortValue)
    || compareNaturalText(right.stopDisplay, left.stopDisplay);

  return stopCompare;
}

function compareKeyedStopRows(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareDesignations(left.keyDesignation, right.keyDesignation)
    || compareNormalStopRows(left, right);
}

function compareNormalStopRows(left: PullSheetRecord, right: PullSheetRecord): number {
  const pickCompare = compareDescendingNullableNumber(left.pickTicketSortValue, right.pickTicketSortValue)
    || compareNaturalText(right.pickTicketNo, left.pickTicketNo);

  if (pickCompare !== 0) {
    return pickCompare;
  }

  return compareNaturalText(left.dmasPart, right.dmasPart)
    || left.sourceIndex - right.sourceIndex;
}

function compareDescendingNullableNumber(left: number | null, right: number | null): number {
  if (left !== null && right !== null && left !== right) {
    return right - left;
  }

  if (left !== null && right === null) {
    return -1;
  }

  if (left === null && right !== null) {
    return 1;
  }

  return 0;
}

function stableSort<T>(items: T[], compare: (left: T, right: T) => number): T[] {
  return items
    .map((item, index) => ({ item, index }))
    .sort((left, right) => compare(left.item, right.item) || left.index - right.index)
    .map(({ item }) => item);
}

function buildStopAssignments(records: PullSheetRecord[]): Map<string, StopAssignment> {
  const stopCustomers = new Map<string, StopCustomer[]>();
  const seenCustomersByStop = new Map<string, Set<string>>();

  for (const record of records) {
    if (!record.shipToNameKey) {
      continue;
    }

    const stopCustomerKey = `${record.routeKey}|${record.stopKey}`;
    const seenCustomers = getOrCreateSet(seenCustomersByStop, stopCustomerKey);

    if (!seenCustomers.has(record.shipToNameKey)) {
      seenCustomers.add(record.shipToNameKey);
      getOrCreateArray(stopCustomers, stopCustomerKey).push({
        key: record.shipToNameKey,
        displayName: record.shipToName,
        city: record.shipToCity,
        state: record.shipToState,
      });
    }
  }

  const assignments = new Map<string, StopAssignment>();

  for (const [stopCustomerKey, customers] of stopCustomers) {
    const customerLetters = new Map<string, string>();
    const requiresKey = customers.length > 1;

    if (requiresKey) {
      customers.forEach((customer, index) => {
        customerLetters.set(customer.key, getDesignationLetter(index));
      });
    }

    assignments.set(stopCustomerKey, { requiresKey, customerLetters, customers });
  }

  return assignments;
}

function applyStopColorsAndKeys(
  records: PullSheetRecord[],
  stopColorMap: Map<string, string>,
  assignments: Map<string, StopAssignment>,
): void {
  for (const record of records) {
    const stopCustomerKey = `${record.routeKey}|${record.stopKey}`;
    const assignment = assignments.get(stopCustomerKey);

    record.stopColorClass = stopColorMap.get(record.stopKey) ?? getStopColorClass(0);
    record.keyDesignation = assignment?.customerLetters.get(record.shipToNameKey) ?? '';
    record.keyShadeClass = record.keyDesignation ? getKeyShadeClass(record.keyDesignation) : '';
  }
}

function applyRowGroupMetadata(records: PullSheetRecord[], assignments: Map<string, StopAssignment>): void {
  const seenStops = new Set<string>();
  const seenKeyGroups = new Set<string>();

  for (const record of records) {
    const stopCustomerKey = `${record.routeKey}|${record.stopKey}`;
    const assignment = assignments.get(stopCustomerKey);
    const keyGroup = `${stopCustomerKey}|${record.keyDesignation}`;

    record.isFirstRowForStop = !seenStops.has(record.stopKey);
    record.isFirstRowForKeyGroup = Boolean(
      assignment?.requiresKey
      && record.keyDesignation
      && !record.isFirstRowForStop
      && !seenKeyGroups.has(keyGroup),
    );
    record.stopBanner = record.isFirstRowForStop ? buildStopBanner(record, assignment) : null;

    seenStops.add(record.stopKey);
    if (record.keyDesignation) {
      seenKeyGroups.add(keyGroup);
    }
  }
}

function buildStopBanner(record: PullSheetRecord, assignment: StopAssignment | undefined): StopBanner {
  const customerCount = assignment?.customers.length ?? (record.shipToNameKey ? 1 : 0);
  const firstCustomer = assignment?.customers[0];
  const stopColorClass = record.stopColorClass;

  if (assignment?.requiresKey) {
    const lastDesignation = getDesignationLetter(customerCount - 1);

    return {
      stopDisplay: record.stopDisplay,
      stopColorClass,
      text: `STOP ${record.stopDisplay} - MULTI-CUSTOMER STOP - USE KEY LETTERS A-${lastDesignation} - ${customerCount} CUSTOMERS`,
      customerCount,
      isMultiCustomer: true,
      firstDesignation: 'A',
      lastDesignation,
    };
  }

  const displayName = firstCustomer?.displayName || record.shipToName || 'Customer not listed';
  const city = firstCustomer?.city || record.shipToCity;
  const state = firstCustomer?.state || record.shipToState;
  const location = [city, state].filter(Boolean).join(', ');
  const text = location
    ? `STOP ${record.stopDisplay} - ${displayName} - ${location}`
    : `STOP ${record.stopDisplay} - ${displayName}`;

  return {
    stopDisplay: record.stopDisplay,
    stopColorClass,
    text,
    customerCount,
    isMultiCustomer: false,
    firstDesignation: '',
    lastDesignation: '',
  };
}

function buildPalletKeyPages(
  routeKey: string,
  routeDisplay: string,
  records: PullSheetRecord[],
  assignments: Map<string, StopAssignment>,
  stopColorMap: Map<string, string>,
): PalletKeyPage[] {
  const pages: PalletKeyPage[] = [];
  const recordsByStop = groupBy(records, (record) => record.stopKey);

  for (const [stopKey, stopRecords] of recordsByStop) {
    const assignment = assignments.get(`${routeKey}|${stopKey}`);

    if (!assignment?.requiresKey) {
      continue;
    }

    const entryMap = new Map<string, PalletKeyEntry>();

    for (const record of stopRecords) {
      const designation = assignment.customerLetters.get(record.shipToNameKey);

      if (!designation) {
        continue;
      }

      if (!entryMap.has(record.shipToNameKey)) {
        entryMap.set(record.shipToNameKey, {
          designation,
          keyShadeClass: getKeyShadeClass(designation),
          shipToName: record.shipToName,
          shipToCity: record.shipToCity,
          shipToState: record.shipToState,
          pickTicketNumbers: [],
          poNumbers: [],
        });
      }

      const entry = entryMap.get(record.shipToNameKey);

      if (entry) {
        pushUnique(entry.pickTicketNumbers, record.pickTicketNo);
        pushUnique(entry.poNumbers, record.poNo);
      }
    }

    pages.push({
      id: `${routeKey}-${stopKey}`,
      routeKey,
      routeDisplay,
      stopKey,
      stopDisplay: stopRecords[0].stopDisplay,
      stopColorClass: stopColorMap.get(stopKey) ?? getStopColorClass(0),
      entries: [...entryMap.values()],
    });
  }

  return pages;
}

function withKeyColumn(columns: PullSheetColumn[]): PullSheetColumn[] {
  const [stopColumn, ...remainingColumns] = columns;
  return [stopColumn, KEY_COLUMN, ...remainingColumns].map((column) => ({
    ...column,
    widthPercent: KEYED_COLUMN_WIDTHS.get(column.key) ?? column.widthPercent,
  }));
}

function normalizeRouteKey(routeDisplay: string): string {
  return normalizeComparableText(routeDisplay) || '(blank route)';
}

function normalizeStopKey(stopDisplay: string): string {
  const numericStop = parseSortableNumber(stopDisplay);
  return numericStop === null ? normalizeComparableText(stopDisplay) : String(numericStop);
}

function groupBy<T>(items: T[], getKey: (item: T) => string): Map<string, T[]> {
  const groups = new Map<string, T[]>();

  for (const item of items) {
    getOrCreateArray(groups, getKey(item)).push(item);
  }

  return groups;
}

function getOrderedUnique(values: string[]): string[] {
  return [...new Set(values)];
}

function getOrCreateArray<T>(map: Map<string, T[]>, key: string): T[] {
  const existing = map.get(key);

  if (existing) {
    return existing;
  }

  const created: T[] = [];
  map.set(key, created);
  return created;
}

function getOrCreateSet(map: Map<string, Set<string>>, key: string): Set<string> {
  const existing = map.get(key);

  if (existing) {
    return existing;
  }

  const created = new Set<string>();
  map.set(key, created);
  return created;
}

function pushUnique(values: string[], value: string): void {
  if (value && !values.includes(value)) {
    values.push(value);
  }
}
