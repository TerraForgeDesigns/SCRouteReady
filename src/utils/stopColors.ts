export interface StopColorToken {
  className: string;
}

export const STOP_COLOR_TOKENS: StopColorToken[] = [
  { className: 'stop-color-1' },
  { className: 'stop-color-2' },
  { className: 'stop-color-3' },
  { className: 'stop-color-4' },
  { className: 'stop-color-5' },
  { className: 'stop-color-6' },
  { className: 'stop-color-7' },
  { className: 'stop-color-8' },
  { className: 'stop-color-9' },
  { className: 'stop-color-10' },
  { className: 'stop-color-11' },
  { className: 'stop-color-12' },
  { className: 'stop-color-13' },
  { className: 'stop-color-14' },
];

export function getStopColorClass(index: number): string {
  return STOP_COLOR_TOKENS[index % STOP_COLOR_TOKENS.length].className;
}
