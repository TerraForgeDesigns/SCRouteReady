import { normalizeComparableText } from './sanitizeDisplayText';

export function normalizeHeader(value: unknown): string {
  return normalizeComparableText(value)
    .replace(/\s*#\s*/g, '#')
    .replace(/\s+/g, ' ');
}
