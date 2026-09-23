/**
 * Campos de licença que só admin / webhook podem alterar.
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export const LICENSE_FIELDS = [
  'status',
  'tipoPlano',
  'diasTeste',
  'validadeLicenca',
  'role',
  'dataCadastro',
  'kiwifyOrderId',
  'dataAtivacaoKiwify',
  'dataCancelamentoKiwify',
] as const;

export type LicenseField = (typeof LICENSE_FIELDS)[number];
