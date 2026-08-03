import type { WorkbookMetadata, WorkbookReadError } from '../types/workbook';
import type { PullSheetGenerationError } from '../types/pullSheet';
import { formatFileSize } from '../utils/formatFileSize';

interface FileStatusPanelProps {
  metadata: WorkbookMetadata | null;
  error: WorkbookReadError | null;
  generationError: PullSheetGenerationError | null;
}

export function FileStatusPanel({ metadata, error, generationError }: FileStatusPanelProps) {
  if (error) {
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

  if (!metadata) {
    return (
      <section className="status-panel" aria-labelledby="status-title">
        <p className="status-kicker">Waiting</p>
        <h2 id="status-title">No route file selected</h2>
        <p>Select the P21 Results export to begin.</p>
      </section>
    );
  }

  return (
    <section className="status-panel success-panel" aria-labelledby="status-title">
      <p className="status-kicker">Workbook ready</p>
      <h2 id="status-title">P21 Results export read locally</h2>
      <dl className="metadata-grid">
        <div>
          <dt>File name</dt>
          <dd>{metadata.originalFileName}</dd>
        </div>
        <div>
          <dt>File size</dt>
          <dd>{formatFileSize(metadata.fileSize)}</dd>
        </div>
        <div>
          <dt>Worksheets</dt>
          <dd>{metadata.worksheetNames.length}</dd>
        </div>
        <div>
          <dt>Results rows</dt>
          <dd>{metadata.resultsRowCount}</dd>
        </div>
        <div>
          <dt>Results columns</dt>
          <dd>{metadata.resultsColumnCount}</dd>
        </div>
      </dl>
      <ul className="confirmation-list">
        <li>Workbook read locally in this browser.</li>
        <li>Route data was not transmitted.</li>
        <li>The next stage will generate the temporary print document.</li>
      </ul>
    </section>
  );
}
