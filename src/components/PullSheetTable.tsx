import { Fragment } from 'react';
import type { PullSheetRecord, PullSheetRouteSection } from '../types/pullSheet';

interface PullSheetTableProps {
  route: PullSheetRouteSection;
}

export function PullSheetTable({ route }: PullSheetTableProps) {
  return (
    <table className={`pull-sheet-table ${route.hasKeyColumn ? 'with-key-column' : 'without-key-column'}`}>
      <colgroup>
        {route.columns.map((column) => (
          <col className={column.className} style={{ width: `${column.widthPercent}%` }} key={column.key} />
        ))}
      </colgroup>
      <thead>
        <tr>
          {route.columns.map((column) => (
            <th className={`${column.className} header-cell`} scope="col" key={column.key}>
              {column.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {route.rows.map((row) => (
          <Fragment key={`${row.sourceKind}-${row.sourceIndex}-${row.pickTicketNo}-${row.dmasPart}-${row.p21ItemId}`}>
            {row.stopBanner ? (
              <tr className={`stop-banner-row ${row.stopBanner.stopColorClass}`}>
                <th className="stop-banner-cell" colSpan={route.columns.length} scope="rowgroup">
                  {row.stopBanner.text}
                </th>
              </tr>
            ) : null}
            {row.sourceDividerLabel ? (
              <tr className="source-divider-row">
                <th className="source-divider-cell" colSpan={route.columns.length} scope="rowgroup">
                  {row.sourceDividerLabel}
                </th>
              </tr>
            ) : null}
            <tr className={`${row.stopColorClass} ${row.isFirstRowForStop ? 'first-stop-row' : ''} ${row.isFirstRowForKeyGroup ? 'first-key-group-row' : ''}`}>
              {route.columns.map((column) => (
                <td className={`${column.className} align-${column.align} body-cell wrap-${column.wrap} ${getKeyCellClass(row, column.key)}`} key={column.key}>
                  {getRecordValue(row, column.key)}
                </td>
              ))}
            </tr>
          </Fragment>
        ))}
      </tbody>
    </table>
  );
}

function getKeyCellClass(row: PullSheetRecord, key: string): string {
  return key === 'keyDesignation' ? row.keyShadeClass : '';
}

function getRecordValue(row: PullSheetRecord, key: string): string {
  switch (key) {
    case 'stopDisplay':
      return row.stopDisplay;
    case 'keyDesignation':
      return row.keyDesignation;
    case 'customerType':
      return row.customerType;
    case 'pickTicketNo':
      return row.documentNumberDisplay;
    case 'poNo':
      return row.poNo;
    case 'qty':
      return row.qty;
    case 'uom':
      return row.uom;
    case 'bin':
      return row.bin;
    case 'dmasPart':
      return row.dmasPart;
    case 'p21ItemId':
      return row.p21ItemId;
    case 'itemDescription':
      return row.itemDescription;
    default:
      return '';
  }
}
