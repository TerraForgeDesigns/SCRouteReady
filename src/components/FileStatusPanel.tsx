import type { PullSheetGenerationError } from '../types/pullSheet';
import type { ActiveWorkbookSession, ReportKind, WorkbookReadError } from '../types/workbook';
import { formatFileSize } from '../utils/formatFileSize';

interface FileStatusPanelProps {
  reports: ActiveWorkbookSession[];
  error: WorkbookReadError | null;
  generationError: PullSheetGenerationError | null;
  onRemoveReport: (kind: ReportKind) => void;
}

export function FileStatusPanel({ reports, error, generationError, onRemoveReport }: FileStatusPanelProps) {
  if (error && reports.length === 0) {
    return (
      <section className="status-panel error-panel" aria-labelledby="status-title">
        <p className="status-kicker">Needs attention</p>
        <h2 id="status-title">File could not be read</h2>
        <p>{error.message}</p>
        <p className="muted">Choose another .xlsx export or use Finish &amp; Clear to reset this task.</p>
      </section>
    );
  }

  if (generationError) {
    return (
      <section className="status-panel error-panel" aria-labelledby="generation-error-title">
        <p className="status-kicker">Generation stopped</p>
        <h2 id="generation-error-title">{generationError.message}</h2>
        {generationError.missingColumns.length > 0 ? (
          <>
            <p>Expected missing columns:</p>
            <ul className="missing-column-list">
              {generationError.missingColumns.map((column) => (
                <li key={column}>{column}</li>
              ))}
            </ul>
          </>
        ) : (
          <p>Retry generation or use Finish &amp; Clear to reset this task.</p>
        )}
      </section>
    );
  }

  if (reports.length === 0) {
    return (
      <section className="status-panel" aria-labelledby="status-title">
        <p className="status-kicker">Waiting</p>
        <h2 id="status-title">No staging reports selected</h2>
        <p>Select a Sales Order report, a Transfer report, or both matching reports.</p>
      </section>
    );
  }

  return (
    <section className="status-panel success-panel" aria-labelledby="status-title">
      <p className="status-kicker">Workbook ready</p>
      <h2 id="status-title">{getStatusTitle(reports)}</h2>
      {error ? <p className="inline-error">{error.message}</p> : null}
      <div className="loaded-report-list">
        {reports.map((report) => (
          <article className="loaded-report" key={report.id}>
            <div>
              <strong>{getReportLabel(report.reportKind)}</strong>
              <span>{report.fileName}</span>
            </div>
            <dl className="metadata-grid compact">
              <div>
                <dt>Rows</dt>
                <dd>{report.metadata.resultsRowCount}</dd>
              </div>
              <div>
                <dt>Stops</dt>
                <dd>{report.metadata.stopCount}</dd>
              </div>
              <div>
                <dt>Routes</dt>
                <dd>{report.metadata.routeCount}</dd>
              </div>
              <div>
                <dt>Size</dt>
                <dd>{formatFileSize(report.metadata.fileSize)}</dd>
              </div>
            </dl>
            <button type="button" className="remove-report-button" onClick={() => onRemoveReport(report.reportKind)}>
              Remove {getReportLabel(report.reportKind)}
            </button>
          </article>
        ))}
      </div>
      {reports.length === 1 ? <p className="muted">{getAddOtherMessage(reports[0].reportKind)}</p> : null}
      <ul className="confirmation-list">
        <li>Workbook data is held temporarily in this browser.</li>
        <li>The next stage will generate one combined print document.</li>
      </ul>
    </section>
  );
}

function getStatusTitle(reports: ActiveWorkbookSession[]): string {
  const hasSales = reports.some((report) => report.reportKind === 'sales');
  const hasTransfer = reports.some((report) => report.reportKind === 'transfer');

  if (hasSales && hasTransfer) {
    return 'Mixed Sales / Transfer route ready';
  }

  return hasSales ? 'Sales Order report ready' : 'Transfer report ready';
}

function getReportLabel(kind: ReportKind): string {
  return kind === 'sales' ? 'Sales Report' : 'Transfer Report';
}

function getAddOtherMessage(kind: ReportKind): string {
  return kind === 'sales'
    ? 'You may generate now or add the matching Transfer report.'
    : 'You may generate now or add the matching Sales report.';
}
