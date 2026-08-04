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
import type { ActiveWorkbookSession, ReportKind } from '../types/workbook';
import { compareDesignations, getDesignationLetter, getKeyShadeClass } from '../utils/designationLetters';
import { normalizeHeader } from '../utils/headerNormalization';
import { compareNaturalText, parseSortableNumber } from '../utils/naturalSort';
import { normalizeComparableText, sanitizeDisplayText } from '../utils/sanitizeDisplayText';
import { getStopColorClass } from '../utils/stopColors';

const SALES_COLUMNS: SourceColumnDefinition[] = [
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

const TRANSFER_COLUMNS: SourceColumnDefinition[] = [
  { field: 'routeFromDescartes', sourceHeader: 'Route from Descartes', displayLabel: 'Route from Descartes', required: true },
  { field: 'assignedStop', sourceHeader: 'Assigned Stop', displayLabel: 'Assigned Stop', required: true },
  { field: 'transferNo', sourceHeader: 'Transfer No', displayLabel: 'Transfer No', required: true },
  { field: 'qty', sourceHeader: 'Qty', displayLabel: 'Qty', required: true },
  { field: 'unitOfMeasure', sourceHeader: 'Unit of Measure', displayLabel: 'Unit of Measure', required: true },
  { field: 'binCd', sourceHeader: 'Bin Cd', displayLabel: 'Bin Cd', required: true },
  { field: 'itemId', sourceHeader: 'Item Id', displayLabel: 'Item Id', required: true },
  { field: 'itemDescription', sourceHeader: 'Item Desc', displayLabel: 'Item Desc', required: true },
  { field: 'toLocationName', sourceHeader: 'To Location Name', displayLabel: 'To Location Name', required: true },
];

const BASE_COLUMNS: PullSheetColumn[] = [
  { key: 'stopDisplay', label: 'Stop', className: 'column-stopDisplay', widthPercent: 4.5, align: 'center', wrap: 'nowrap' },
  { key: 'customerType', label: 'Type', className: 'column-customerType', widthPercent: 6.5, align: 'center', wrap: 'nowrap' },
  { key: 'pickTicketNo', label: 'PT/Xfer #', className: 'column-pickTicketNo', widthPercent: 11, align: 'center', wrap: 'nowrap' },
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

type ColumnIndexMap = Partial<Record<SourceField, number>>;
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

interface ParsedReport {
  kind: ReportKind;
  records: PullSheetRecord[];
}

export function generatePullSheetDocument(workbook: unknown): PullSheetDocument {
  return buildDocument([readReport(workbook, 'sales')]);
}

export function generatePullSheetDocumentFromSessions(sessions: readonly ActiveWorkbookSession[]): PullSheetDocument {
  if (sessions.length === 0) {
    throw new PullSheetGenerationError('NO_USABLE_DATA_ROWS', 'No valid reports are loaded.');
  }

  const reports = sessions.map((session) => readReport(session.workbook, session.reportKind));
  validateRouteCompatibility(reports);
  return buildDocument(reports);
}

function readReport(workbook: unknown, kind: ReportKind): ParsedReport {
  const { headerRow, dataRows } = readResultsRows(workbook);
  const columnMap = buildColumnMap(headerRow, kind === 'sales' ? SALES_COLUMNS : TRANSFER_COLUMNS);
  const records = kind === 'sales'
    ? convertSalesRows(dataRows, columnMap)
    : convertTransferRows(dataRows, columnMap);
  const usableRecords = records.filter((record) => hasUsableData(record));

  if (usableRecords.length === 0) {
    throw new PullSheetGenerationError('NO_USABLE_DATA_ROWS', 'The Results worksheet does not contain usable data rows.');
  }

  return { kind, records: usableRecords };
}

function readResultsRows(workbook: unknown): { headerRow: RawRow; dataRows: RawRow[] } {
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

    return { headerRow, dataRows };
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

function buildColumnMap(headerRow: RawRow, columns: SourceColumnDefinition[]): ColumnIndexMap {
  const normalizedHeaderIndexes = new Map<string, number>();

  headerRow.forEach((header, index) => {
    const normalized = normalizeHeader(header);

    if (normalized && !normalizedHeaderIndexes.has(normalized)) {
      normalizedHeaderIndexes.set(normalized, index);
    }
  });

  const missingColumns: string[] = [];
  const map: ColumnIndexMap = {};

  for (const column of columns) {
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

  return map;
}

function convertSalesRows(rows: RawRow[], columnMap: ColumnIndexMap): PullSheetRecord[] {
  return rows.map((row, index) => {
    const routeDisplay = getRequiredCell(row, columnMap, 'routeFromDescartes');
    const stopDisplay = getRequiredCell(row, columnMap, 'stopNumber');
    const pickTicketNo = getRequiredCell(row, columnMap, 'pickTicketNo');
    const shipToName = getRequiredCell(row, columnMap, 'shipToName');

    return createRecord({
      sourceKind: 'sales',
      sourceIndex: index,
      routeDisplay,
      customerType: getRequiredCell(row, columnMap, 'customerType'),
      stopDisplay,
      documentNumberDisplay: pickTicketNo,
      poNo: getRequiredCell(row, columnMap, 'poNo'),
      qty: getRequiredCell(row, columnMap, 'qty'),
      uom: getRequiredCell(row, columnMap, 'uom'),
      bin: getRequiredCell(row, columnMap, 'bin'),
      dmasPart: getRequiredCell(row, columnMap, 'dmasPart'),
      p21ItemId: getRequiredCell(row, columnMap, 'p21ItemId'),
      itemDescription: getRequiredCell(row, columnMap, 'itemDescription'),
      shipToName,
      shipToCity: getRequiredCell(row, columnMap, 'shipToCity'),
      shipToState: getRequiredCell(row, columnMap, 'shipToState'),
      transferDestinationName: '',
    });
  });
}

function convertTransferRows(rows: RawRow[], columnMap: ColumnIndexMap): PullSheetRecord[] {
  return rows.map((row, index) => {
    const routeDisplay = getRequiredCell(row, columnMap, 'routeFromDescartes');
    const stopDisplay = getRequiredCell(row, columnMap, 'assignedStop');
    const transferNo = getRequiredCell(row, columnMap, 'transferNo');
    const toLocationName = getRequiredCell(row, columnMap, 'toLocationName');

    return createRecord({
      sourceKind: 'transfer',
      sourceIndex: index,
      routeDisplay,
      customerType: 'Xfer',
      stopDisplay,
      documentNumberDisplay: transferNo,
      poNo: '',
      qty: getRequiredCell(row, columnMap, 'qty'),
      uom: getRequiredCell(row, columnMap, 'unitOfMeasure'),
      bin: getRequiredCell(row, columnMap, 'binCd'),
      dmasPart: '',
      p21ItemId: getRequiredCell(row, columnMap, 'itemId'),
      itemDescription: getRequiredCell(row, columnMap, 'itemDescription'),
      shipToName: '',
      shipToCity: '',
      shipToState: '',
      transferDestinationName: toLocationName,
    });
  });
}

function createRecord(values: {
  sourceKind: ReportKind;
  sourceIndex: number;
  routeDisplay: string;
  customerType: string;
  stopDisplay: string;
  documentNumberDisplay: string;
  poNo: string;
  qty: string;
  uom: string;
  bin: string;
  dmasPart: string;
  p21ItemId: string;
  itemDescription: string;
  shipToName: string;
  shipToCity: string;
  shipToState: string;
  transferDestinationName: string;
}): PullSheetRecord {
  const stopSortValue = parseSortableNumber(values.stopDisplay);

  return {
    sourceKind: values.sourceKind,
    sourceIndex: values.sourceIndex,
    originalSourceOrder: values.sourceIndex,
    routeKey: normalizeRouteKey(values.routeDisplay),
    routeDisplay: values.routeDisplay,
    customerType: values.customerType,
    stopKey: normalizeStopKey(values.stopDisplay),
    stopDisplay: values.stopDisplay,
    stopSortValue,
    pickTicketNo: values.documentNumberDisplay,
    pickTicketSortValue: parseSortableNumber(values.documentNumberDisplay),
    documentNumberDisplay: values.documentNumberDisplay,
    documentNumberSortValue: parseSortableNumber(values.documentNumberDisplay),
    poNo: values.poNo,
    qty: values.qty,
    uom: values.uom,
    bin: values.bin,
    dmasPart: values.dmasPart,
    p21ItemId: values.p21ItemId,
    itemDescription: values.itemDescription,
    shipToName: values.shipToName,
    shipToNameKey: normalizeComparableText(values.shipToName),
    shipToCity: values.shipToCity,
    shipToState: values.shipToState,
    transferDestinationName: values.transferDestinationName,
    transferDestinationKey: normalizeComparableText(values.transferDestinationName),
    keyDesignation: '',
    keyShadeClass: '',
    stopColorClass: '',
    isFirstRowForStop: false,
    isFirstRowForKeyGroup: false,
    stopBanner: null,
    sourceDividerLabel: '',
  };
}

function getRequiredCell(row: RawRow, columnMap: ColumnIndexMap, field: SourceField): string {
  const index = columnMap[field];
  return index === undefined ? '' : sanitizeDisplayText(row[index] ?? '');
}

function hasUsableData(record: PullSheetRecord): boolean {
  return [
    record.routeDisplay,
    record.customerType,
    record.stopDisplay,
    record.documentNumberDisplay,
    record.qty,
    record.uom,
    record.bin,
    record.p21ItemId,
    record.itemDescription,
    record.shipToName,
    record.transferDestinationName,
  ].some((value) => value.length > 0);
}

function validateRouteCompatibility(reports: ParsedReport[]): void {
  if (reports.length !== 2) {
    return;
  }

  const salesReport = reports.find((report) => report.kind === 'sales');
  const transferReport = reports.find((report) => report.kind === 'transfer');

  if (!salesReport || !transferReport) {
    return;
  }

  const salesRoutes = getSortedRouteKeys(salesReport.records);
  const transferRoutes = getSortedRouteKeys(transferReport.records);
  const routeSetsMatch = salesRoutes.length === transferRoutes.length
    && salesRoutes.every((routeKey, index) => routeKey === transferRoutes[index]);

  if (!routeSetsMatch) {
    throw new PullSheetGenerationError(
      'ROUTE_MISMATCH',
      'These Sales and Transfer reports do not appear to belong to the same Descartes route.',
    );
  }
}

function buildDocument(reports: ParsedReport[]): PullSheetDocument {
  const records = reports.flatMap((report) => report.records);
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
    const salesLineCount = sortedRecords.filter((record) => record.sourceKind === 'sales').length;
    const transferLineCount = sortedRecords.length - salesLineCount;

    routeSections.push({
      routeKey,
      routeDisplay,
      truckType: getTruckType(salesLineCount, transferLineCount),
      columns,
      rows: sortedRecords,
      stopCount: stopKeys.length,
      lineCount: sortedRecords.length,
      salesLineCount,
      transferLineCount,
      palletKeyPageCount: routePalletPages.length,
      hasKeyColumn,
    });
    palletKeyPages.push(...routePalletPages);
    totalStops += stopKeys.length;
  }

  const salesLineCount = records.filter((record) => record.sourceKind === 'sales').length;
  const transferLineCount = records.length - salesLineCount;

  return {
    summary: {
      reportType: getReportType(salesLineCount, transferLineCount),
      routeCount: routeSections.length,
      stopCount: totalStops,
      productLineCount: records.length,
      salesLineCount,
      transferLineCount,
      palletKeyPageCount: palletKeyPages.length,
    },
    routes: routeSections,
    palletKeyPages,
  };
}

function compareBasePullSheetRecords(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareStopOrder(left, right)
    || compareSourceKind(left, right)
    || compareSourceRows(left, right);
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

  const sourceCompare = compareSourceKind(left, right);

  if (sourceCompare !== 0) {
    return sourceCompare;
  }

  if (left.sourceKind === 'sales') {
    const assignment = assignments.get(`${left.routeKey}|${left.stopKey}`);

    if (assignment?.requiresKey) {
      return compareKeyedStopRows(left, right);
    }
  }

  return compareSourceRows(left, right);
}

function compareStopOrder(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareDescendingNullableNumber(left.stopSortValue, right.stopSortValue)
    || compareNaturalText(right.stopDisplay, left.stopDisplay);
}

function compareSourceKind(left: PullSheetRecord, right: PullSheetRecord): number {
  if (left.sourceKind === right.sourceKind) {
    return 0;
  }

  return left.sourceKind === 'sales' ? -1 : 1;
}

function compareKeyedStopRows(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareDesignations(left.keyDesignation, right.keyDesignation)
    || compareSalesRows(left, right);
}

function compareSourceRows(left: PullSheetRecord, right: PullSheetRecord): number {
  return left.sourceKind === 'transfer'
    ? compareTransferRows(left, right)
    : compareSalesRows(left, right);
}

function compareSalesRows(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareDescendingNullableNumber(left.documentNumberSortValue, right.documentNumberSortValue)
    || compareNaturalText(right.documentNumberDisplay, left.documentNumberDisplay)
    || compareNaturalText(left.dmasPart, right.dmasPart)
    || left.originalSourceOrder - right.originalSourceOrder;
}

function compareTransferRows(left: PullSheetRecord, right: PullSheetRecord): number {
  return compareDescendingNullableNumber(left.documentNumberSortValue, right.documentNumberSortValue)
    || compareNaturalText(right.documentNumberDisplay, left.documentNumberDisplay)
    || compareNaturalText(left.p21ItemId, right.p21ItemId)
    || left.originalSourceOrder - right.originalSourceOrder;
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
    if (record.sourceKind !== 'sales' || !record.shipToNameKey) {
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
    record.keyDesignation = record.sourceKind === 'sales' ? assignment?.customerLetters.get(record.shipToNameKey) ?? '' : '';
    record.keyShadeClass = record.keyDesignation ? getKeyShadeClass(record.keyDesignation) : '';
  }
}

function applyRowGroupMetadata(records: PullSheetRecord[], assignments: Map<string, StopAssignment>): void {
  const seenStops = new Set<string>();
  const seenKeyGroups = new Set<string>();
  const seenSourceGroups = new Set<string>();
  const recordsByStop = groupBy(records, (record) => `${record.routeKey}|${record.stopKey}`);

  for (const record of records) {
    const stopCustomerKey = `${record.routeKey}|${record.stopKey}`;
    const assignment = assignments.get(stopCustomerKey);
    const keyGroup = `${stopCustomerKey}|${record.keyDesignation}`;
    const sourceGroup = `${stopCustomerKey}|${record.sourceKind}`;
    const stopRecords = recordsByStop.get(stopCustomerKey) ?? [];
    const isMixedStop = stopRecords.some((item) => item.sourceKind === 'sales')
      && stopRecords.some((item) => item.sourceKind === 'transfer');

    record.isFirstRowForStop = !seenStops.has(stopCustomerKey);
    record.isFirstRowForKeyGroup = Boolean(
      assignment?.requiresKey
      && record.keyDesignation
      && !record.isFirstRowForStop
      && !seenKeyGroups.has(keyGroup),
    );
    record.stopBanner = record.isFirstRowForStop ? buildStopBanner(record, assignment, stopRecords) : null;
    record.sourceDividerLabel = isMixedStop && !seenSourceGroups.has(sourceGroup)
      ? buildSourceDividerLabel(record.sourceKind, stopRecords)
      : '';

    seenStops.add(stopCustomerKey);
    seenSourceGroups.add(sourceGroup);
    if (record.keyDesignation) {
      seenKeyGroups.add(keyGroup);
    }
  }
}

function buildStopBanner(
  record: PullSheetRecord,
  assignment: StopAssignment | undefined,
  stopRecords: PullSheetRecord[],
): StopBanner {
  const stopColorClass = record.stopColorClass;
  const hasSales = stopRecords.some((item) => item.sourceKind === 'sales');
  const hasTransfer = stopRecords.some((item) => item.sourceKind === 'transfer');
  const customerCount = assignment?.customers.length ?? (record.shipToNameKey ? 1 : 0);

  if (hasSales && hasTransfer && assignment?.requiresKey) {
    const lastDesignation = getDesignationLetter(customerCount - 1);

    return {
      stopDisplay: record.stopDisplay,
      stopColorClass,
      text: `STOP ${record.stopDisplay} - MIXED SALES / XFER - SALES USE KEY LETTERS A-${lastDesignation} - ${customerCount} SALES CUSTOMERS`,
      customerCount,
      isMultiCustomer: true,
      firstDesignation: 'A',
      lastDesignation,
    };
  }

  if (hasSales && hasTransfer) {
    return {
      stopDisplay: record.stopDisplay,
      stopColorClass,
      text: `STOP ${record.stopDisplay} - MIXED SALES / XFER STOP`,
      customerCount,
      isMultiCustomer: false,
      firstDesignation: '',
      lastDesignation: '',
    };
  }

  if (hasTransfer) {
    const destination = getTransferDestinationLabel(stopRecords);
    return {
      stopDisplay: record.stopDisplay,
      stopColorClass,
      text: `STOP ${record.stopDisplay} - XFER TO ${destination}`,
      customerCount: 0,
      isMultiCustomer: false,
      firstDesignation: '',
      lastDesignation: '',
    };
  }

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

  const firstCustomer = assignment?.customers[0];
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

function buildSourceDividerLabel(sourceKind: ReportKind, stopRecords: PullSheetRecord[]): string {
  if (sourceKind === 'sales') {
    return 'SALES ORDERS';
  }

  return `XFER ITEMS - ${getTransferDestinationLabel(stopRecords)}`;
}

function getTransferDestinationLabel(records: PullSheetRecord[]): string {
  const destinations = records
    .filter((record) => record.sourceKind === 'transfer' && record.transferDestinationKey)
    .reduce<Map<string, string>>((map, record) => {
      if (!map.has(record.transferDestinationKey)) {
        map.set(record.transferDestinationKey, record.transferDestinationName);
      }

      return map;
    }, new Map<string, string>());

  if (destinations.size === 0) {
    return 'Destination not listed';
  }

  if (destinations.size > 1) {
    return 'MULTIPLE DESTINATIONS';
  }

  return [...destinations.values()][0];
}

function buildPalletKeyPages(
  routeKey: string,
  routeDisplay: string,
  records: PullSheetRecord[],
  assignments: Map<string, StopAssignment>,
  stopColorMap: Map<string, string>,
): PalletKeyPage[] {
  const pages: PalletKeyPage[] = [];
  const recordsByStop = groupBy(records.filter((record) => record.sourceKind === 'sales'), (record) => record.stopKey);

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
        pushUnique(entry.pickTicketNumbers, record.documentNumberDisplay);
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

function getTruckType(salesLineCount: number, transferLineCount: number): 'Sales' | 'Xfer' | 'Mixed' {
  if (salesLineCount > 0 && transferLineCount > 0) {
    return 'Mixed';
  }

  return transferLineCount > 0 ? 'Xfer' : 'Sales';
}

function getReportType(
  salesLineCount: number,
  transferLineCount: number,
): PullSheetDocument['summary']['reportType'] {
  if (salesLineCount > 0 && transferLineCount > 0) {
    return 'Mixed Sales / Transfer';
  }

  return transferLineCount > 0 ? 'Transfer' : 'Sales Order';
}

function getSortedRouteKeys(records: PullSheetRecord[]): string[] {
  return [...new Set(records.map((record) => record.routeKey))].sort((left, right) => compareNaturalText(left, right));
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
