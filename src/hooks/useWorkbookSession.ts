import { useCallback, useRef, useState } from 'react';
import { generatePullSheetDocument } from '../services/pullSheetGenerator';
import { readWorkbookSession, WorkbookReaderError } from '../services/workbookReader';
import { PullSheetGenerationError, type PullSheetDocument } from '../types/pullSheet';
import type { ActiveWorkbookSession, ApplicationState, WorkbookMetadata, WorkbookReadError } from '../types/workbook';
import { validateFileSelection } from '../utils/fileValidation';

interface WorkbookSessionState {
  appState: ApplicationState;
  metadata: WorkbookMetadata | null;
  error: WorkbookReadError | null;
  generationError: PullSheetGenerationError | null;
  generatedDocument: PullSheetDocument | null;
  announcement: string;
  hasActiveSession: boolean;
}

export function useWorkbookSession() {
  const activeSessionRef = useRef<ActiveWorkbookSession | null>(null);
  const [state, setState] = useState<WorkbookSessionState>({
    appState: 'idle',
    metadata: null,
    error: null,
    generationError: null,
    generatedDocument: null,
    announcement: 'No route file selected',
    hasActiveSession: false,
  });

  const clearSession = useCallback((announcement = 'Active session cleared') => {
    activeSessionRef.current = null;
    setState({
      appState: 'idle',
      metadata: null,
      error: null,
      generationError: null,
      generatedDocument: null,
      announcement,
      hasActiveSession: false,
    });
  }, []);

  const readSelectedFiles = useCallback(async (files: FileList | File[] | null) => {
    if (state.appState === 'reading') {
      return;
    }

    activeSessionRef.current = null;
    const validation = validateFileSelection(files);

    if (!validation.valid) {
      setState({
        appState: 'error',
        metadata: null,
        error: {
          code: validation.errorCode ?? 'UNEXPECTED_PARSE_ERROR',
          message: validation.message ?? 'The selected file could not be validated.',
        },
        generationError: null,
        generatedDocument: null,
        announcement: validation.message ?? 'The selected file could not be validated.',
        hasActiveSession: true,
      });
      return;
    }

    const file = files?.[0];

    if (!file) {
      setState({
        appState: 'error',
        metadata: null,
        error: { code: 'NO_FILE', message: 'No file selected. Choose a P21 Results export to begin.' },
        generationError: null,
        generatedDocument: null,
        announcement: 'No file selected. Choose a P21 Results export to begin.',
        hasActiveSession: true,
      });
      return;
    }

    setState({
      appState: 'reading',
      metadata: null,
      error: null,
      generationError: null,
      generatedDocument: null,
      announcement: 'Reading workbook locally...',
      hasActiveSession: true,
    });

    try {
      const session = await readWorkbookSession(file);
      activeSessionRef.current = session;
      setState({
        appState: 'ready',
        metadata: session.metadata,
        error: null,
        generationError: null,
        generatedDocument: null,
        announcement: 'Workbook read locally. Generate Pull Sheet is available.',
        hasActiveSession: true,
      });
    } catch (error) {
      activeSessionRef.current = null;
      const workbookError = normalizeWorkbookError(error);
      setState({
        appState: 'error',
        metadata: null,
        error: workbookError,
        generationError: null,
        generatedDocument: null,
        announcement: workbookError.message,
        hasActiveSession: true,
      });
    }
  }, [state.appState]);

  const generatePullSheet = useCallback(async () => {
    if (state.appState === 'generating') {
      return;
    }

    const activeSession = activeSessionRef.current;

    if (!activeSession) {
      return;
    }

    setState((current) => ({
      ...current,
      appState: 'generating',
      generationError: null,
      generatedDocument: null,
      announcement: 'Generating Pull Sheet locally...',
    }));

    await Promise.resolve();

    try {
      const generatedDocument = generatePullSheetDocument(activeSession.workbook);
      setState((current) => ({
        ...current,
        appState: 'preview',
        generationError: null,
        generatedDocument,
        announcement: `Pull Sheet generated. ${generatedDocument.summary.palletKeyPageCount} Pallet Key pages created.`,
      }));
    } catch (error) {
      const generationError = normalizeGenerationError(error);
      setState((current) => ({
        ...current,
        appState: 'ready',
        generationError,
        generatedDocument: null,
        announcement: generationError.message,
      }));
    }
  }, [state.appState]);

  const setDragging = useCallback((isDragging: boolean) => {
    setState((current) => {
      if (current.appState === 'reading') {
        return current;
      }

      return {
        ...current,
        appState: isDragging ? 'dragging' : current.generatedDocument ? 'preview' : current.metadata ? 'ready' : current.error ? 'error' : 'idle',
      };
    });
  }, []);

  return {
    ...state,
    clearSession,
    generatePullSheet,
    readSelectedFiles,
    setDragging,
  };
}

function normalizeWorkbookError(error: unknown): WorkbookReadError {
  if (error instanceof WorkbookReaderError) {
    return { code: error.code, message: error.message };
  }

  return {
    code: 'UNEXPECTED_PARSE_ERROR',
    message: 'An unexpected parsing error occurred. Try another P21 .xlsx export.',
  };
}

function normalizeGenerationError(error: unknown): PullSheetGenerationError {
  if (error instanceof PullSheetGenerationError) {
    return error;
  }

  return new PullSheetGenerationError('TRANSFORMATION_FAILED', 'The Pull Sheet could not be generated from this worksheet.');
}
