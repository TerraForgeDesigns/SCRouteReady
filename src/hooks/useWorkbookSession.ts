import { useCallback, useRef, useState } from 'react';
import { generatePullSheetDocumentFromSessions } from '../services/pullSheetGenerator';
import { readWorkbookSession, WorkbookReaderError } from '../services/workbookReader';
import { PullSheetGenerationError, type PullSheetDocument } from '../types/pullSheet';
import type { ActiveWorkbookSession, ApplicationState, ReportKind, WorkbookReadError } from '../types/workbook';
import { validateFileSelection } from '../utils/fileValidation';

interface WorkbookSessionState {
  appState: ApplicationState;
  reports: ActiveWorkbookSession[];
  error: WorkbookReadError | null;
  generationError: PullSheetGenerationError | null;
  generatedDocument: PullSheetDocument | null;
  announcement: string;
  hasActiveSession: boolean;
}

export function useWorkbookSession() {
  const activeSessionsRef = useRef<ActiveWorkbookSession[]>([]);
  const [state, setState] = useState<WorkbookSessionState>({
    appState: 'idle',
    reports: [],
    error: null,
    generationError: null,
    generatedDocument: null,
    announcement: 'No staging reports selected',
    hasActiveSession: false,
  });

  const clearSession = useCallback((announcement = 'Active session cleared') => {
    activeSessionsRef.current = [];
    setState({
      appState: 'idle',
      reports: [],
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

    const validation = validateFileSelection(files);

    if (!validation.valid) {
      setState({
        appState: 'error',
        reports: activeSessionsRef.current,
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

    const selectedFiles = Array.from(files ?? []);

    if (selectedFiles.length === 0) {
      setState({
        appState: 'error',
        reports: activeSessionsRef.current,
        error: { code: 'NO_FILE', message: 'No file selected. Choose a P21 Results export to begin.' },
        generationError: null,
        generatedDocument: null,
        announcement: 'No file selected. Choose a P21 Staging Report to begin.',
        hasActiveSession: true,
      });
      return;
    }

    setState({
      appState: 'reading',
      reports: activeSessionsRef.current,
      error: null,
      generationError: null,
      generatedDocument: null,
      announcement: selectedFiles.length === 1 ? 'Reading workbook locally...' : 'Reading workbooks locally...',
      hasActiveSession: true,
    });

    try {
      const newSessions = await Promise.all(selectedFiles.map((file) => readWorkbookSession(file)));
      const duplicateKind = findDuplicateKind(newSessions);

      if (duplicateKind) {
        throw new WorkbookReaderError(
          'DUPLICATE_REPORT_KIND',
          duplicateKind === 'sales'
            ? 'Two Sales Order reports were selected. Choose one Sales report and, if needed, one matching Transfer report.'
            : 'Two Transfer reports were selected. Choose one Transfer report and, if needed, one matching Sales report.',
        );
      }

      const mergedSessions = mergeSessionsByKind(activeSessionsRef.current, newSessions);
      activeSessionsRef.current = mergedSessions;
      setState({
        appState: 'ready',
        reports: mergedSessions,
        error: null,
        generationError: null,
        generatedDocument: null,
        announcement: getReadyAnnouncement(mergedSessions),
        hasActiveSession: true,
      });
    } catch (error) {
      const workbookError = normalizeWorkbookError(error);
      setState({
        appState: activeSessionsRef.current.length > 0 ? 'ready' : 'error',
        reports: activeSessionsRef.current,
        error: workbookError,
        generationError: null,
        generatedDocument: null,
        announcement: workbookError.message,
        hasActiveSession: true,
      });
    }
  }, [state.appState]);

  const removeReport = useCallback((kind: ReportKind) => {
    const remainingSessions = activeSessionsRef.current.filter((session) => session.reportKind !== kind);
    activeSessionsRef.current = remainingSessions;
    setState({
      appState: remainingSessions.length > 0 ? 'ready' : 'idle',
      reports: remainingSessions,
      error: null,
      generationError: null,
      generatedDocument: null,
      announcement: remainingSessions.length > 0 ? getReadyAnnouncement(remainingSessions) : 'Report removed. No staging reports selected.',
      hasActiveSession: remainingSessions.length > 0,
    });
  }, []);

  const generatePullSheet = useCallback(async () => {
    if (state.appState === 'generating') {
      return;
    }

    const activeSessions = activeSessionsRef.current;

    if (activeSessions.length === 0) {
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
      const generatedDocument = generatePullSheetDocumentFromSessions(activeSessions);
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
        appState: isDragging ? 'dragging' : current.generatedDocument ? 'preview' : current.reports.length > 0 ? 'ready' : current.error ? 'error' : 'idle',
      };
    });
  }, []);

  return {
    ...state,
    clearSession,
    generatePullSheet,
    readSelectedFiles,
    removeReport,
    setDragging,
  };
}

function findDuplicateKind(sessions: ActiveWorkbookSession[]): ReportKind | null {
  const seen = new Set<ReportKind>();

  for (const session of sessions) {
    if (seen.has(session.reportKind)) {
      return session.reportKind;
    }

    seen.add(session.reportKind);
  }

  return null;
}

function mergeSessionsByKind(
  currentSessions: ActiveWorkbookSession[],
  newSessions: ActiveWorkbookSession[],
): ActiveWorkbookSession[] {
  const sessionMap = new Map<ReportKind, ActiveWorkbookSession>();

  for (const session of currentSessions) {
    sessionMap.set(session.reportKind, session);
  }

  for (const session of newSessions) {
    sessionMap.set(session.reportKind, session);
  }

  return (['sales', 'transfer'] as const)
    .map((kind) => sessionMap.get(kind))
    .filter((session): session is ActiveWorkbookSession => Boolean(session));
}

function getReadyAnnouncement(sessions: ActiveWorkbookSession[]): string {
  const hasSales = sessions.some((session) => session.reportKind === 'sales');
  const hasTransfer = sessions.some((session) => session.reportKind === 'transfer');

  if (hasSales && hasTransfer) {
    return 'Mixed Sales / Transfer route ready. Generate Pull Sheet is available.';
  }

  if (hasSales) {
    return 'Sales Order report ready. You may generate now or add the matching Transfer report.';
  }

  return 'Transfer report ready. You may generate now or add the matching Sales report.';
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
