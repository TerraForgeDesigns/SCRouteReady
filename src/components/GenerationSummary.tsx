import type { GenerationSummaryData } from '../types/pullSheet';

interface GenerationSummaryProps {
  summary: GenerationSummaryData;
}

export function GenerationSummary({ summary }: GenerationSummaryProps) {
  return (
    <section className="generation-summary" aria-label="Generation summary">
      <div>
        <span>{summary.reportType}</span>
        <strong>Report type</strong>
      </div>
      <div>
        <span>{summary.routeCount}</span>
        <strong>Routes</strong>
      </div>
      <div>
        <span>{summary.stopCount}</span>
        <strong>Stops</strong>
      </div>
      <div>
        <span>{summary.productLineCount}</span>
        <strong>Product lines</strong>
      </div>
      {summary.salesLineCount > 0 ? (
        <div>
          <span>{summary.salesLineCount}</span>
          <strong>Sales lines</strong>
        </div>
      ) : null}
      {summary.transferLineCount > 0 ? (
        <div>
          <span>{summary.transferLineCount}</span>
          <strong>Xfer lines</strong>
        </div>
      ) : null}
      <div>
        <span>{summary.palletKeyPageCount}</span>
        <strong>ABC Pallet Keys</strong>
      </div>
    </section>
  );
}
