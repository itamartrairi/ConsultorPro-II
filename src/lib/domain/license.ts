/**
 * Helpers de licença e timestamp de itens.
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

import { LICENSE_FIELDS } from '../constants/license';

export { toJsDate, computeLicenseDaysLeft } from '../licenseDays';

export function pickLicenseFields(item: any): Record<string, any> {
  const out: Record<string, any> = {};
  if (!item) return out;
  for (const k of LICENSE_FIELDS) {
    if (item[k] !== undefined) out[k] = item[k];
  }
  return out;
}

export function extractItemTimestamp(item: any): number {
  if (!item) return 0;

  const dateCandidates = [
    item.updatedAt,
    item.dataModificacao,
    item.dataAtualizacao,
    item.dataDiagnostico,
    item.dataCadastro,
    item.createdAt,
    item.dataCriacao,
    item.dataVencimento,
    item.dataInicio,
    item.timestamp,
  ];

  for (const candidate of dateCandidates) {
    if (!candidate) continue;
    if (typeof candidate === 'number' && !isNaN(candidate) && candidate > 0) {
      return candidate < 10000000000 ? candidate * 1000 : candidate;
    }
    if (typeof candidate === 'object') {
      if (candidate instanceof Date && !isNaN(candidate.getTime())) {
        return candidate.getTime();
      }
      if (typeof candidate.toMillis === 'function') {
        try {
          return candidate.toMillis();
        } catch {
          /* ignore */
        }
      }
      if (typeof candidate.toDate === 'function') {
        try {
          return candidate.toDate().getTime();
        } catch {
          /* ignore */
        }
      }
      if (typeof candidate.seconds === 'number') {
        return candidate.seconds * 1000;
      }
    }
    if (typeof candidate === 'string') {
      const trimmed = candidate.trim();
      if (trimmed) {
        const parsed = Date.parse(trimmed.includes('T') ? trimmed : trimmed + 'T12:00:00');
        if (!isNaN(parsed) && parsed > 0) return parsed;
      }
    }
  }

  let contentWeight = 0;
  if (item.resposta) contentWeight += 1000;
  if (item.observacao && String(item.observacao).trim()) contentWeight += 500;
  if (item.score !== undefined && item.score > 0) contentWeight += 200;
  if (item.cronograma && Array.isArray(item.cronograma) && item.cronograma.length > 0) contentWeight += 5000;
  if (item.dadosConsultoria && (item.dadosConsultoria.codigoSgf || item.dadosConsultoria.consultor)) {
    contentWeight += 2000;
  }

  return contentWeight;
}
