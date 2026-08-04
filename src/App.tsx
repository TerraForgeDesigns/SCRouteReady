import { useEffect, useRef, useState } from 'react';
import { ActionBar } from './components/ActionBar';
import { AppHeader } from './components/AppHeader';
import { FileDropZone } from './components/FileDropZone';
import { FileStatusPanel } from './components/FileStatusPanel';
import { PullSheetPreview } from './components/PullSheetPreview';
import { SecurityNotice } from './components/SecurityNotice';
import { StatusMessage } from './components/StatusMessage';
import { WorkflowSteps } from './components/WorkflowSteps';
import { useWorkbookSession } from './hooks/useWorkbookSession';

export default function App() {
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
  const [inputKey, setInputKey] = useState(0);

  useEffect(() => {
    if (generatedDocument) {
      previewHeadingRef.current?.focus();
      previewHeadingRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [generatedDocument]);

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
            {generatedDocument ? <PullSheetPreview document={generatedDocument} ref={previewHeadingRef} /> : null}
          </div>
          <div className="secondary-column">
            <SecurityNotice />
            <FileStatusPanel reports={reports} error={error} generationError={generationError} onRemoveReport={removeReport} />
          </div>
        </div>
      </main>
    </>
  );
}
