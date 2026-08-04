import { useRef } from 'react';
import type { ApplicationState } from '../types/workbook';

interface FileDropZoneProps {
  appState: ApplicationState;
  onFilesSelected: (files: FileList | File[] | null) => void;
  onDragChange: (isDragging: boolean) => void;
  onInputReady: (input: HTMLInputElement | null) => void;
  inputKey: number;
}

export function FileDropZone({ appState, onFilesSelected, onDragChange, onInputReady, inputKey }: FileDropZoneProps) {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const isReading = appState === 'reading';

  function setInputRef(input: HTMLInputElement | null) {
    inputRef.current = input;
    onInputReady(input);
  }

  return (
    <section
      className={`drop-zone ${appState}`}
      aria-labelledby="drop-zone-title"
      onDragEnter={(event) => {
        event.preventDefault();
        if (!isReading) {
          onDragChange(true);
        }
      }}
      onDragOver={(event) => {
        event.preventDefault();
        if (!isReading) {
          event.dataTransfer.dropEffect = 'copy';
        }
      }}
      onDragLeave={(event) => {
        if (event.currentTarget.contains(event.relatedTarget as Node)) {
          return;
        }
        onDragChange(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        onDragChange(false);
        if (!isReading) {
          onFilesSelected(Array.from(event.dataTransfer.files));
        }
      }}
    >
      <div className="drop-zone-icon" aria-hidden="true">
        <svg viewBox="0 0 48 48">
          <path d="M14 34h20a8 8 0 0 0 1.6-15.8A12 12 0 0 0 12.3 22 6.5 6.5 0 0 0 14 34Z" />
          <path d="M24 31V18m0 0-5 5m5-5 5 5" />
        </svg>
      </div>
      <div>
        <h2 id="drop-zone-title">Drop P21 Staging Report(s) Here</h2>
        <p>Drop a Sales Order report, a Transfer report, or both reports for the same truck.</p>
      </div>
      <label className="file-button">
        <span>Choose Report(s)</span>
        <input
          key={inputKey}
          ref={setInputRef}
          type="file"
          accept=".xlsx"
          multiple
          disabled={isReading}
          aria-label="Choose P21 Staging Report"
          onChange={(event) => onFilesSelected(event.currentTarget.files)}
        />
      </label>
      <p className="file-rule">.xlsx files only - Maximum 20 MB each</p>
      {isReading ? <p className="reading-indicator">Reading workbook locally...</p> : null}
    </section>
  );
}
