/**
 * Helpers de normalização de áreas.
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export function normalizeAndFormatArea(rawArea: string | undefined | null): string {
  if (!rawArea) return '';
  const trimmed = rawArea.trim();
  if (!trimmed) return '';
  return trimmed
    .toLowerCase()
    .split(/\s+/)
    .map((word) => {
      if (word.length <= 2 && ['de', 'da', 'do', 'em', 'para', 'e'].includes(word)) return word;
      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(' ');
}

export function globalNormalizedMatch(
  val1: string | undefined | null,
  val2: string | undefined | null
): boolean {
  if (!val1 && !val2) return true;
  if (!val1 || !val2) return false;
  return val1.trim().toLowerCase() === val2.trim().toLowerCase();
}
