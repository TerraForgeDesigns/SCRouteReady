import { forwardRef } from 'react';
import type { PullSheetDocument } from '../types/pullSheet';
import { GenerationSummary } from './GenerationSummary';
import { PalletKeyPage } from './PalletKeyPage';
import { PullSheetRouteSection } from './PullSheetRouteSection';

interface PullSheetPreviewProps {
  document: PullSheetDocument;
}

export const PullSheetPreview = forwardRef<HTMLHeadingElement, PullSheetPreviewProps>(
  function PullSheetPreview({ document }, ref) {
    return (
      <section className="preview-shell" aria-labelledby="preview-heading">
        <div className="preview-toolbar">
          <div>
            <p className="status-kicker">Preview ready</p>
            <h2 id="preview-heading" ref={ref} tabIndex={-1}>Generated Pull Sheet</h2>
          </div>
        </div>
        <GenerationSummary summary={document.summary} />
        <div className="print-document" aria-label="Generated print document">
          {document.routes.map((route, index) => (
            <PullSheetRouteSection route={route} isFirst={index === 0} key={route.routeKey} />
          ))}
          {document.palletKeyPages.map((page) => (
            <PalletKeyPage page={page} key={page.id} />
          ))}
        </div>
      </section>
    );
  },
);
