export type SourceField =
  | 'routeFromDescartes'
  | 'customerType'
  | 'stopNumber'
  | 'pickTicketNo'
  | 'assignedStop'
  | 'transferNo'
  | 'poNo'
  | 'qty'
  | 'uom'
  | 'unitOfMeasure'
  | 'bin'
  | 'binCd'
  | 'dmasPart'
  | 'p21ItemId'
  | 'itemId'
  | 'itemDescription'
  | 'shipToName'
  | 'shipToCity'
  | 'shipToState'
  | 'toLocationName';

export interface SourceColumnDefinition {
  field: SourceField;
  sourceHeader: string;
  displayLabel: string;
  required: boolean;
}

export interface PullSheetRecord {
  sourceKind: 'sales' | 'transfer';
  sourceIndex: number;
  originalSourceOrder: number;
  routeKey: string;
  routeDisplay: string;
  customerType: string;
  stopKey: string;
  stopDisplay: string;
  stopSortValue: number | null;
  pickTicketNo: string;
  pickTicketSortValue: number | null;
  documentNumberDisplay: string;
  documentNumberSortValue: number | null;
  poNo: string;
  qty: string;
  uom: string;
  bin: string;
  dmasPart: string;
  p21ItemId: string;
  itemDescription: string;
  shipToName: string;
  shipToNameKey: string;
  shipToCity: string;
  shipToState: string;
  transferDestinationName: string;
  transferDestinationKey: string;
  keyDesignation: string;
  keyShadeClass: string;
  stopColorClass: string;
  isFirstRowForStop: boolean;
  isFirstRowForKeyGroup: boolean;
  stopBanner: StopBanner | null;
  sourceDividerLabel: string;
}

export interface StopBanner {
  stopDisplay: string;
  stopColorClass: string;
  text: string;
  customerCount: number;
  isMultiCustomer: boolean;
  firstDesignation: string;
  lastDesignation: string;
}

export interface PullSheetColumn {
  key: string;
  label: string;
  className: string;
  widthPercent: number;
  align: 'left' | 'center' | 'right';
  wrap: 'nowrap' | 'description';
}

export interface PullSheetRouteSection {
  routeKey: string;
  routeDisplay: string;
  truckType: 'Sales' | 'Xfer' | 'Mixed';
  columns: PullSheetColumn[];
  rows: PullSheetRecord[];
  stopCount: number;
  lineCount: number;
  salesLineCount: number;
  transferLineCount: number;
  palletKeyPageCount: number;
  hasKeyColumn: boolean;
}

export interface PalletKeyEntry {
  designation: string;
  keyShadeClass: string;
  shipToName: string;
  shipToCity: string;
  shipToState: string;
  pickTicketNumbers: string[];
  poNumbers: string[];
}

export interface PalletKeyPage {
  id: string;
  routeKey: string;
  routeDisplay: string;
  stopKey: string;
  stopDisplay: string;
  stopColorClass: string;
  entries: PalletKeyEntry[];
}

export interface GenerationSummaryData {
  reportType: 'Sales Order' | 'Transfer' | 'Mixed Sales / Transfer';
  routeCount: number;
  stopCount: number;
  productLineCount: number;
  salesLineCount: number;
  transferLineCount: number;
  palletKeyPageCount: number;
}

export interface PullSheetDocument {
  summary: GenerationSummaryData;
  routes: PullSheetRouteSection[];
  palletKeyPages: PalletKeyPage[];
}

export type GenerationErrorCode =
  | 'REQUIRED_HEADERS_MISSING'
  | 'RESULTS_WORKSHEET_EMPTY'
  | 'NO_USABLE_DATA_ROWS'
  | 'UNEXPECTED_WORKSHEET_STRUCTURE'
  | 'ROUTE_MISMATCH'
  | 'TRANSFORMATION_FAILED';

export class PullSheetGenerationError extends Error {
  readonly code: GenerationErrorCode;
  readonly missingColumns: string[];

  constructor(code: GenerationErrorCode, message: string, missingColumns: string[] = []) {
    super(message);
    this.name = 'PullSheetGenerationError';
    this.code = code;
    this.missingColumns = missingColumns;
  }
}
