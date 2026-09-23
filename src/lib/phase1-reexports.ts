/**
 * Barrel de reexports da Fase 1.
 */

export type {
  Empresa,
  EmpresaCredenciada,
  Premissa,
  Problema,
  Diagnostico,
  Resposta,
  Solucao,
  TarefaPlanoAcao,
  DadosConsultoria,
  AtividadeCronograma,
} from '../types/domain';

export {
  cleanDigits,
  formatCNPJ,
  isValidCNPJ,
  formatCEP,
  formatCPF,
  isValidCPF,
} from './formatters/br';

export {
  getMdaExpirationStatus,
  parseLocalDate,
  formatFirestoreDate,
  getActivityDateStr,
} from './formatters/dates';

export { sanitizeForFirestore } from './firestore/sanitize';
export { oklchToRgb, captureElementWithHtml2Canvas } from './pdf/capture';

export {
  CONSULTORIA_AREAS,
  AREAS,
  AREAS_ORDER,
  TIPOS_EMPRESA,
  IMPACTO_ORDER,
} from './constants/areas';
export { LICENSE_FIELDS } from './constants/license';
export type { LicenseField } from './constants/license';

export { normalizeAndFormatArea, globalNormalizedMatch } from './domain/areas';
export {
  pickLicenseFields,
  extractItemTimestamp,
  toJsDate,
  computeLicenseDaysLeft,
} from './domain/license';
export {
  deduplicateRespostas,
  loadAllLocalRespostas,
  saveAllLocalRespostas,
  getRespostasForDiagnostico,
  clearRespostasCache,
} from './domain/respostas';

export { isValidLogoSource } from './media/logo';
export { Type } from './ai/schemaTypes';
export { extractAndParseJSON } from './ai/parseJson';
