const collator = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: 'base',
});

export function compareNaturalText(left: string, right: string): number {
  return collator.compare(left, right);
}

export function parseSortableNumber(value: string): number | null {
  const normalized = value.trim().replace(/,/g, '');

  if (!normalized || !/^-?\d+(\.\d+)?$/.test(normalized)) {
    return null;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}
