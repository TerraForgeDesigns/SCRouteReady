interface ActionBarProps {
  isReady: boolean;
  isGenerating: boolean;
  hasGeneratedDocument: boolean;
  hasActiveSession: boolean;
  onGenerate: () => void;
  onPrint: () => void;
  onClear: () => void;
}

export function ActionBar({
  isReady,
  isGenerating,
  hasGeneratedDocument,
  hasActiveSession,
  onGenerate,
  onPrint,
  onClear,
}: ActionBarProps) {
  const canGenerate = (isReady || hasGeneratedDocument) && !isGenerating;

  return (
    <section className="action-bar" aria-label="Task actions">
      <div>
        <button className="primary-action" type="button" disabled={!canGenerate} onClick={onGenerate}>
          {hasGeneratedDocument ? 'Regenerate Pull Sheet' : isGenerating ? 'Generating...' : 'Generate Pull Sheet'}
        </button>
        <p>{isReady || hasGeneratedDocument ? 'Creates the temporary browser print document.' : 'Available after a valid workbook is ready.'}</p>
      </div>
      <button type="button" disabled={!hasGeneratedDocument || isGenerating} onClick={onPrint}>
        Print
      </button>
      <button type="button" disabled={!hasActiveSession} onClick={onClear}>
        Finish &amp; Clear
      </button>
    </section>
  );
}
