/**
 * Validação de fontes de logo (URL ou data URL).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export function isValidLogoSource(src: any): boolean {
  if (!src || typeof src !== 'string') return false;
  const trimmed = src.trim();
  if (trimmed.length === 0) return false;
  if (trimmed === 'null' || trimmed === 'undefined') return false;
  if (trimmed.startsWith('data:')) {
    return (
      trimmed.startsWith('data:image/') &&
      trimmed.includes(';base64,') &&
      trimmed.length > 35
    );
  }
  return (
    trimmed.startsWith('http://') ||
    trimmed.startsWith('https://') ||
    trimmed.startsWith('/') ||
    trimmed.startsWith('./')
  );
}
