export function sanitizeDisplayText(value: unknown): string {
  if (value === null || value === undefined) {
    return '';
  }

  return [...String(value)].filter((character) => !isUnsafeControlCharacter(character)).join('').trim();
}

export function normalizeComparableText(value: unknown): string {
  return sanitizeDisplayText(value).replace(/\s+/g, ' ').toLowerCase();
}

function isUnsafeControlCharacter(character: string): boolean {
  const codePoint = character.codePointAt(0);

  if (codePoint === undefined) {
    return false;
  }

  return (codePoint <= 8) || codePoint === 11 || codePoint === 12 || (codePoint >= 14 && codePoint <= 31) || codePoint === 127;
}
