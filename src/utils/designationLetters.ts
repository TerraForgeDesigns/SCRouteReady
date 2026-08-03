export function getDesignationLetter(index: number): string {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error('Designation index must be a non-negative integer.');
  }

  let remaining = index + 1;
  let label = '';

  while (remaining > 0) {
    remaining -= 1;
    label = String.fromCharCode(65 + (remaining % 26)) + label;
    remaining = Math.floor(remaining / 26);
  }

  return label;
}

export function designationToNumber(designation: string): number {
  const normalized = designation.trim().toUpperCase();

  if (!/^[A-Z]+$/.test(normalized)) {
    return Number.POSITIVE_INFINITY;
  }

  return [...normalized].reduce((total, character) => (
    total * 26 + character.charCodeAt(0) - 64
  ), 0);
}

export function compareDesignations(left: string, right: string): number {
  return designationToNumber(left) - designationToNumber(right);
}

export function getKeyShadeClass(designation: string): string {
  const designationNumber = designationToNumber(designation);

  if (!Number.isFinite(designationNumber)) {
    return '';
  }

  return `key-shade-${((designationNumber - 1) % 4) + 1}`;
}
