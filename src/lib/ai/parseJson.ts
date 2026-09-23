/**
 * Extrai e faz parse de JSON a partir de texto livre (respostas do Gemini).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export function extractAndParseJSON(text: string, defaultValue: any = null): any {
  if (!text) return defaultValue;

  let cleaned = text.trim();

  try {
    return JSON.parse(cleaned);
  } catch {
    /* continue */
  }

  if (cleaned.includes('```')) {
    const match = cleaned.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (match && match[1]) {
      const candidateContent = match[1].trim();
      try {
        return JSON.parse(candidateContent);
      } catch {
        cleaned = candidateContent;
      }
    }
  }

  for (let i = 0; i < cleaned.length; i++) {
    const char = cleaned[i];
    if (char === '{' || char === '[') {
      const openChar = char;
      const closeChar = char === '{' ? '}' : ']';
      let depth = 0;
      let closingIndex = -1;

      for (let j = i; j < cleaned.length; j++) {
        if (cleaned[j] === openChar) {
          depth++;
        } else if (cleaned[j] === closeChar) {
          depth--;
          if (depth === 0) {
            closingIndex = j;
            const candidate = cleaned.substring(i, closingIndex + 1);
            try {
              return JSON.parse(candidate);
            } catch {
              /* continue searching */
            }
          }
        }
      }
    }
  }

  const firstSquare = cleaned.indexOf('[');
  const firstCurly = cleaned.indexOf('{');

  let firstChar = -1;
  let lastChar = -1;

  if (firstSquare !== -1 && (firstCurly === -1 || firstSquare < firstCurly)) {
    firstChar = firstSquare;
    lastChar = cleaned.lastIndexOf(']');
  } else if (firstCurly !== -1) {
    firstChar = firstCurly;
    lastChar = cleaned.lastIndexOf('}');
  }

  if (firstChar !== -1 && lastChar !== -1 && lastChar > firstChar) {
    const candidate = cleaned.substring(firstChar, lastChar + 1);
    try {
      return JSON.parse(candidate);
    } catch (substringErr) {
      console.error('Subsegment parsing failed as well:', substringErr);
    }
  }

  console.error('Failed to parse JSON directly/indirectly. Original text:', text);
  return defaultValue;
}
