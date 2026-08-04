import type { PullSheetRouteSection as PullSheetRouteSectionModel } from '../types/pullSheet';
import { PullSheetTable } from './PullSheetTable';

interface PullSheetRouteSectionProps {
  route: PullSheetRouteSectionModel;
  isFirst: boolean;
}

export function PullSheetRouteSection({ route, isFirst }: PullSheetRouteSectionProps) {
  return (
    <article className={`preview-page pull-sheet-page ${isFirst ? '' : 'new-print-page'}`}>
      <header className="print-page-header">
        <div className="print-brand-block">
          <img src="/assets/SCRouteReady_logo_cropped.png" alt="SCRouteReady" className="print-logo" />
        </div>
        <div className="route-title-block">
          <h3>Truck Pull Sheet</h3>
          <p>Route {route.routeDisplay}</p>
        </div>
        <p className="route-summary">
          {route.truckType} - {route.stopCount} Stops - {route.lineCount} Lines{route.truckType === 'Mixed' ? ` - ${route.salesLineCount} Sales - ${route.transferLineCount} Xfer` : ''} - {route.palletKeyPageCount} Keys
        </p>
      </header>
      <PullSheetTable route={route} />
    </article>
  );
}
