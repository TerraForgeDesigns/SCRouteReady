import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
import { generatePullSheetDocument } from '../src/services/pullSheetGenerator';
import { compareDesignations, getDesignationLetter, getKeyShadeClass } from '../src/utils/designationLetters';
import { getStopColorClass, STOP_COLOR_TOKENS } from '../src/utils/stopColors';

const HEADERS = [
  '_internalrowindex',
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
  'Order No',
  'Ship To City',
  'Ship To State',
  'Route Code',
  'Extended Desc',
];

function createWorkbook(rows: string[][]): XLSX.WorkBook {
  const worksheet = XLSX.utils.aoa_to_sheet([HEADERS, ...rows]);
  return {
    SheetNames: ['Results'],
    Sheets: {
      Results: worksheet,
    },
  };
}

function createRow(overrides: Partial<Record<string, string>> = {}): string[] {
  const base: Record<string, string> = {
    '_internalrowindex': '1',
    'Route from Descartes': 'Route 100',
    'Customer Type': 'Standard',
    'Stop #': '01',
    'Pick Ticket No': '100',
    'Po No': 'PO-100',
    Qty: '1',
    UOM: 'EA',
    Bin: 'A1',
    'DMAS Part#': 'A2',
    'P21 Item Id': 'P21-1',
    'Item Desc': 'Neutral Item',
    'Ship To Name': 'Customer Alpha',
    'Order No': 'O-1',
    'Ship To City': 'City One',
    'Ship To State': 'SC',
    'Route Code': 'R100',
    'Extended Desc': 'Not displayed',
  };

  const values = { ...base, ...overrides };
  return HEADERS.map((header) => values[header] ?? '');
}

test('one stop with one Ship To Name does not create a key column or Pallet Key page', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ 'Pick Ticket No': '101' }),
    createRow({ 'Pick Ticket No': '100' }),
  ]));

  assert.equal(document.summary.palletKeyPageCount, 0);
  assert.equal(document.routes[0].hasKeyColumn, false);
  assert.equal(document.routes[0].columns.length, 10);
  assert.deepEqual(document.routes[0].columns.map((column) => column.key), [
    'stopDisplay',
    'customerType',
    'pickTicketNo',
    'poNo',
    'qty',
    'uom',
    'bin',
    'dmasPart',
    'p21ItemId',
    'itemDescription',
  ]);
  assert.deepEqual(document.routes[0].columns.map((column) => column.label), [
    'Stop',
    'Type',
    'Pick Ticket',
    'PO No.',
    'Qty',
    'UOM',
    'Bin',
    'DMAS Part',
    'P21 Item ID',
    'Item Description',
  ]);
  assert.deepEqual(document.routes[0].columns.map((column) => column.widthPercent), [
    4.5,
    6.5,
    11,
    10,
    4,
    5,
    6,
    11,
    11,
    31,
  ]);
  assert.equal(document.routes[0].columns.reduce((total, column) => total + column.widthPercent, 0), 100);
  assert.equal(
    Math.max(...document.routes[0].columns.map((column) => column.widthPercent)),
    document.routes[0].columns.find((column) => column.key === 'itemDescription')?.widthPercent,
  );
  assert.equal(document.routes[0].rows.every((row) => row.keyDesignation === ''), true);
  assert.equal(document.routes[0].rows.every((row) => row.keyShadeClass === ''), true);
  assert.equal(Object.prototype.hasOwnProperty.call(document.routes[0].rows[0], 'rowColorClass'), false);
  assert.equal(document.routes[0].rows[0].stopBanner?.text, 'STOP 01 - Customer Alpha - City One, SC');
  assert.equal(document.routes[0].rows[0].stopBanner?.stopColorClass, document.routes[0].rows[0].stopColorClass);
  assert.equal(document.routes[0].rows[0].stopBanner?.isMultiCustomer, false);
  assert.equal(document.routes[0].rows[1].stopBanner, null);
});

test('scattered keyed rows are grouped by designation after assignment', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ 'Pick Ticket No': '105', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'A1' }),
    createRow({ 'Pick Ticket No': '104', 'Ship To Name': 'Customer Beta', 'DMAS Part#': 'B1' }),
    createRow({ 'Pick Ticket No': '103', 'Ship To Name': 'Customer Gamma', 'DMAS Part#': 'C1' }),
    createRow({ 'Pick Ticket No': '102', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'A2' }),
    createRow({ 'Pick Ticket No': '101', 'Ship To Name': 'Customer Beta', 'DMAS Part#': 'B2' }),
  ]));

  assert.equal(document.summary.palletKeyPageCount, 1);
  assert.equal(document.routes[0].hasKeyColumn, true);
  assert.equal(document.routes[0].columns.length, 11);
  assert.deepEqual(document.routes[0].columns.map((column) => column.key).slice(0, 3), [
    'stopDisplay',
    'keyDesignation',
    'customerType',
  ]);
  assert.deepEqual(document.routes[0].columns.map((column) => column.label), [
    'Stop',
    'Key',
    'Type',
    'Pick Ticket',
    'PO No.',
    'Qty',
    'UOM',
    'Bin',
    'DMAS Part',
    'P21 Item ID',
    'Item Description',
  ]);
  assert.deepEqual(document.routes[0].columns.map((column) => column.widthPercent), [
    4,
    4,
    6,
    10.5,
    9.5,
    4,
    5,
    6,
    10.5,
    10.5,
    30,
  ]);
  assert.equal(document.routes[0].columns.reduce((total, column) => total + column.widthPercent, 0), 100);
  assert.equal(document.routes[0].columns[1].key, 'keyDesignation');
  assert.equal(document.routes[0].columns[1].align, 'center');
  assert.equal(
    Math.max(...document.routes[0].columns.map((column) => column.widthPercent)),
    document.routes[0].columns.find((column) => column.key === 'itemDescription')?.widthPercent,
  );
  assert.equal(
    document.routes[0].rows[0].stopBanner?.text,
    'STOP 01 - MULTI-CUSTOMER STOP - USE KEY LETTERS A-C - 3 CUSTOMERS',
  );
  assert.equal(document.routes[0].rows[0].stopBanner?.stopColorClass, document.routes[0].rows[0].stopColorClass);
  assert.equal(document.routes[0].rows[0].stopBanner?.isMultiCustomer, true);
  assert.deepEqual(document.palletKeyPages[0].entries.map((entry) => entry.designation), ['A', 'B', 'C']);
  assert.deepEqual(document.palletKeyPages[0].entries.map((entry) => entry.shipToName), [
    'Customer Alpha',
    'Customer Beta',
    'Customer Gamma',
  ]);
  assert.deepEqual(document.routes[0].rows.map((row) => row.keyDesignation), ['A', 'A', 'B', 'B', 'C']);
  assert.deepEqual(document.routes[0].rows.map((row) => row.keyShadeClass), [
    'key-shade-1',
    'key-shade-1',
    'key-shade-2',
    'key-shade-2',
    'key-shade-3',
  ]);
  assert.deepEqual(document.palletKeyPages[0].entries.map((entry) => entry.keyShadeClass), [
    'key-shade-1',
    'key-shade-2',
    'key-shade-3',
  ]);
  assert.deepEqual(document.routes[0].rows.map((row) => row.shipToName), [
    'Customer Alpha',
    'Customer Alpha',
    'Customer Beta',
    'Customer Beta',
    'Customer Gamma',
  ]);
  assert.deepEqual(document.routes[0].rows.map((row) => row.isFirstRowForKeyGroup), [
    false,
    false,
    true,
    false,
    true,
  ]);
});

test('same customer with multiple Pick Tickets remains one letter', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ 'Pick Ticket No': '103', 'Ship To Name': 'Customer Alpha' }),
    createRow({ 'Pick Ticket No': '102', 'Ship To Name': 'Customer Beta' }),
    createRow({ 'Pick Ticket No': '101', 'Ship To Name': 'Customer Alpha' }),
  ]));

  const alphaRows = document.routes[0].rows.filter((row) => row.shipToNameKey === 'customer alpha');
  assert.deepEqual(alphaRows.map((row) => row.keyDesignation), ['A', 'A']);
});

test('product sorting inside a Key group uses pick ticket, DMAS part, then source order', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ '_internalrowindex': '1', 'Pick Ticket No': '201', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'D-20' }),
    createRow({ '_internalrowindex': '2', 'Pick Ticket No': '200', 'Ship To Name': 'Customer Beta', 'DMAS Part#': 'B-10' }),
    createRow({ '_internalrowindex': '3', 'Pick Ticket No': '200', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'D-2' }),
    createRow({ '_internalrowindex': '4', 'Pick Ticket No': '200', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'D-10' }),
    createRow({ '_internalrowindex': '5', 'Pick Ticket No': '200', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'D-10' }),
  ]));

  const alphaRows = document.routes[0].rows.filter((row) => row.keyDesignation === 'A');
  assert.deepEqual(alphaRows.map((row) => row.pickTicketNo), ['201', '200', '200', '200']);
  assert.deepEqual(alphaRows.map((row) => row.dmasPart), ['D-20', 'D-2', 'D-10', 'D-10']);
  assert.deepEqual(alphaRows.map((row) => row.sourceIndex), [0, 2, 3, 4]);
});

test('same Stop # on two routes is treated independently', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ 'Route from Descartes': 'Route 100', 'Stop #': '01', 'Ship To Name': 'Customer Alpha' }),
    createRow({ 'Route from Descartes': 'Route 100', 'Stop #': '01', 'Ship To Name': 'Customer Beta' }),
    createRow({ 'Route from Descartes': 'Route 200', 'Stop #': '01', 'Ship To Name': 'Customer Gamma' }),
    createRow({ 'Route from Descartes': 'Route 200', 'Stop #': '01', 'Ship To Name': 'Customer Delta' }),
  ]));

  assert.equal(document.summary.routeCount, 2);
  assert.equal(document.summary.palletKeyPageCount, 2);
  assert.deepEqual(
    document.palletKeyPages.find((page) => page.routeDisplay === 'Route 100')?.entries.map((entry) => entry.shipToName),
    ['Customer Alpha', 'Customer Beta'],
  );
  assert.deepEqual(
    document.palletKeyPages.find((page) => page.routeDisplay === 'Route 200')?.entries.map((entry) => entry.shipToName),
    ['Customer Gamma', 'Customer Delta'],
  );
});

test('one route with keyed and normal stops shows key column with blank normal stop rows', () => {
  const document = generatePullSheetDocument(createWorkbook([
    createRow({ 'Stop #': '03', 'Pick Ticket No': '100', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'N-2' }),
    createRow({ 'Stop #': '03', 'Pick Ticket No': '101', 'Ship To Name': 'Customer Alpha', 'DMAS Part#': 'N-1' }),
    createRow({ 'Stop #': '02', 'Pick Ticket No': '202', 'Ship To Name': 'Customer Beta', 'DMAS Part#': 'B-1' }),
    createRow({ 'Stop #': '02', 'Pick Ticket No': '201', 'Ship To Name': 'Customer Gamma', 'DMAS Part#': 'C-1' }),
    createRow({ 'Stop #': '02', 'Pick Ticket No': '200', 'Ship To Name': 'Customer Beta', 'DMAS Part#': 'B-2' }),
  ]));

  assert.equal(document.routes[0].hasKeyColumn, true);
  assert.deepEqual(document.routes[0].rows.map((row) => row.stopDisplay), ['03', '03', '02', '02', '02']);
  assert.deepEqual(document.routes[0].rows.map((row) => row.keyDesignation), ['', '', 'A', 'A', 'B']);
  assert.deepEqual(document.routes[0].rows.map((row) => row.pickTicketNo), ['101', '100', '202', '200', '201']);
  assert.equal(
    document.routes[0].rows[2].stopBanner?.text,
    'STOP 02 - MULTI-CUSTOMER STOP - USE KEY LETTERS A-B - 2 CUSTOMERS',
  );
  assert.equal(document.routes[0].rows[0].stopBanner?.text, 'STOP 03 - Customer Alpha - City One, SC');
});

test('designation letters continue after Z', () => {
  assert.equal(getDesignationLetter(0), 'A');
  assert.equal(getDesignationLetter(25), 'Z');
  assert.equal(getDesignationLetter(26), 'AA');
  assert.equal(getDesignationLetter(27), 'AB');
  assert.equal(getDesignationLetter(28), 'AC');
  assert.equal(compareDesignations('Z', 'AA') < 0, true);
  assert.equal(compareDesignations('AA', 'B') > 0, true);
  assert.equal(getKeyShadeClass('A'), 'key-shade-1');
  assert.equal(getKeyShadeClass('B'), 'key-shade-2');
  assert.equal(getKeyShadeClass('C'), 'key-shade-3');
  assert.equal(getKeyShadeClass('D'), 'key-shade-4');
  assert.equal(getKeyShadeClass('E'), 'key-shade-1');
});

test('multi-customer banner uses final designation beyond Z', () => {
  const rows = Array.from({ length: 29 }, (_, index) => createRow({
    'Pick Ticket No': String(100 + index),
    'Ship To Name': `Customer ${String(index + 1).padStart(2, '0')}`,
  }));

  const document = generatePullSheetDocument(createWorkbook(rows));

  assert.equal(document.summary.palletKeyPageCount, 1);
  assert.equal(document.routes[0].rows[0].stopBanner?.lastDesignation, 'AC');
  assert.equal(
    document.routes[0].rows[0].stopBanner?.text,
    'STOP 01 - MULTI-CUSTOMER STOP - USE KEY LETTERS A-AC - 29 CUSTOMERS',
  );
  assert.deepEqual(document.routes[0].rows.map((row) => row.keyDesignation).slice(24, 28), ['Y', 'Z', 'AA', 'AB']);
});

test('trusted stop color palette supports fourteen stops before repeating', () => {
  assert.equal(STOP_COLOR_TOKENS.length >= 14, true);
  assert.equal(new Set(STOP_COLOR_TOKENS.slice(0, 14).map((token) => token.className)).size, 14);

  const rows = Array.from({ length: 14 }, (_, index) => createRow({
    'Stop #': String(index + 1).padStart(2, '0'),
    'Pick Ticket No': String(100 + index),
    'Ship To Name': `Customer ${String(index + 1).padStart(2, '0')}`,
  }));
  const document = generatePullSheetDocument(createWorkbook(rows));
  const firstRowsByStop = document.routes[0].rows.filter((row) => row.isFirstRowForStop);
  const stopColorClasses = firstRowsByStop.map((row) => row.stopColorClass);

  assert.equal(firstRowsByStop.length, 14);
  assert.equal(firstRowsByStop.every((row) => row.stopBanner?.stopColorClass === row.stopColorClass), true);
  assert.equal(new Set(stopColorClasses).size, 14);
  assert.equal(stopColorClasses.every((className, index) => className === getStopColorClass(index)), true);
  assert.equal(stopColorClasses.every((className, index) => index === 0 || className !== stopColorClasses[index - 1]), true);
});
