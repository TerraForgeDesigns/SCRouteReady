import type { ApplicationState } from '../types/workbook';

interface WorkflowStepsProps {
  appState: ApplicationState;
}

const steps = [
  { label: 'Drop File', description: 'Select export' },
  { label: 'Pull Sheet + Keys', description: 'Next build' },
  { label: 'Print', description: 'Future stage' },
];

export function WorkflowSteps({ appState }: WorkflowStepsProps) {
  const hasWorkbook = appState === 'ready' || appState === 'generating' || appState === 'preview';
  const hasPreview = appState === 'preview';

  return (
    <section className="workflow" aria-label="Workflow progress">
      {steps.map((step, index) => {
        const status = getStepStatus(index, hasWorkbook, hasPreview);
        return (
          <div className={`workflow-step ${status}`} key={step.label}>
            <span className="step-icon" aria-hidden="true">
              {status === 'complete' ? (
                <svg viewBox="0 0 24 24"><path d="M20 6 9 17l-5-5" /></svg>
              ) : (
                index + 1
              )}
            </span>
            <span>
              <strong>{step.label}</strong>
              <small>{status === 'active' ? 'Active' : step.description}</small>
            </span>
          </div>
        );
      })}
    </section>
  );
}

function getStepStatus(index: number, hasWorkbook: boolean, hasPreview: boolean): 'active' | 'complete' | 'pending' {
  if (!hasWorkbook) {
    return index === 0 ? 'active' : 'pending';
  }

  if (index === 0) {
    return 'complete';
  }

  if (!hasPreview) {
    return index === 1 ? 'active' : 'pending';
  }

  return index === 1 ? 'complete' : 'active';
}
