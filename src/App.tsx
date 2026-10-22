import type { PullSheetDocument } from './types/pullSheet';
import { useEffect, useRef, useState, type ReactNode } from 'react';
import { ActionBar } from './components/ActionBar';
import { AppHeader } from './components/AppHeader';
import { FileDropZone } from './components/FileDropZone';
import { FileStatusPanel } from './components/FileStatusPanel';
import { PullSheetPreview } from './components/PullSheetPreview';
import { SecurityNotice } from './components/SecurityNotice';
import { StatusMessage } from './components/StatusMessage';
import { WorkflowSteps } from './components/WorkflowSteps';
import { useWorkbookSession } from './hooks/useWorkbookSession';

export default function App({ renderArchive, onGenerated }: { renderArchive?: (document: PullSheetDocument | null) => ReactNode; onGenerated?: (document:PullSheetDocument,markup:string)=>void } = {}) {
  const {
    appState,
    reports,
    error,
    generationError,
    generatedDocument,
    announcement,
    hasActiveSession,
    clearSession,
    generatePullSheet,
    readSelectedFiles,
    removeReport,
    setDragging,
  } = useWorkbookSession();
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const previewHeadingRef = useRef<HTMLHeadingElement | null>(null);
  const printOutputRef = useRef<HTMLDivElement | null>(null);
  const [inputKey, setInputKey] = useState(0);

  useEffect(() => {
    if (generatedDocument) {
      previewHeadingRef.current?.focus();
      previewHeadingRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [generatedDocument]);

  useEffect(() => {
    if (generatedDocument && printOutputRef.current && onGenerated) onGenerated(generatedDocument,printOutputRef.current.outerHTML);
  }, [generatedDocument,onGenerated]);

  function handleClear() {
    clearSession();
    setInputKey((current) => current + 1);
    window.setTimeout(() => fileInputRef.current?.focus(), 0);
  }

  function handlePrint() {
    window.print();
  }

  return (
    <>
      <AppHeader />
      <main className="app-shell">
        <StatusMessage message={announcement} />
        <WorkflowSteps appState={appState} />
        <div className="workspace-grid">
          <div className="primary-column">
            <FileDropZone
              appState={appState}
              inputKey={inputKey}
              onDragChange={setDragging}
              onFilesSelected={(files) => {
                readSelectedFiles(files);
                setInputKey((current) => current + 1);
              }}
              onInputReady={(input) => {
                fileInputRef.current = input;
              }}
            />
            <ActionBar
              isReady={appState === 'ready'}
              isGenerating={appState === 'generating'}
              hasGeneratedDocument={Boolean(generatedDocument)}
              hasActiveSession={hasActiveSession}
              onGenerate={generatePullSheet}
              onPrint={handlePrint}
              onClear={handleClear}
            />
            {renderArchive?.(generatedDocument)}
            {generatedDocument ? <PullSheetPreview document={generatedDocument} printOutputRef={printOutputRef} ref={previewHeadingRef} /> : null}
          </div>
          <div className="secondary-column">
            <SecurityNotice archiveEnabled={!!renderArchive} />
            <FileStatusPanel reports={reports} error={error} generationError={generationError} onRemoveReport={removeReport} />
          </div>
        </div>
      </main>
    </>
  );
}
