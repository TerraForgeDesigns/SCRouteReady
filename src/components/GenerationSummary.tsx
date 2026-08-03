import type { GenerationSummaryData } from '../types/pullSheet';

interface GenerationSummaryProps {
  summary: GenerationSummaryData;
}

export function GenerationSummary({ summary }: GenerationSummaryProps) {
  return (
    <section className="generation-summary" aria-label="Generation summary">
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
      <div>
        <span>{summary.palletKeyPageCount}</span>
        <strong>ABC Pallet Keys</strong>
      </div>
    </section>
  );
}
