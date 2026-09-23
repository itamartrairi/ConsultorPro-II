/**
 * Constantes de áreas de consultoria, diagnóstico e tipos de empresa.
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

/** Áreas usadas no formulário de dados da consultoria (SGF / Sebrae). */
export const CONSULTORIA_AREAS = [
  'FINANÇAS',
  'MARKETING E VENDAS',
  'PLANEJAMENTO ESTRATÉGICO',
  'GESTÃO DE PESSOAS',
  'PROCESSOS E OPERAÇÕES',
  'INOVAÇÃO E TECNOLOGIA',
  'JURÍDICO',
  'SUSTENTABILIDADE',
  'EMPREENDEDORISMO',
  'ACESSO A CRÉDITO',
  'CRÉDITO',
  'OUTROS',
];

/** Áreas do diagnóstico empresarial (biblioteca de problemas/premissas). */
export const AREAS = [
  { id: 'FIN', nome: 'Financeiro' },
  { id: 'MKT', nome: 'Marketing' },
  { id: 'OPS', nome: 'Operacional' },
  { id: 'RH', nome: 'Recursos Humanos' },
  { id: 'VEN', nome: 'Vendas' },
  { id: 'TEC', nome: 'Tecnologia' },
  { id: 'JUR', nome: 'Jurídico' },
  { id: 'EST', nome: 'Estratégico' },
  { id: 'LOG', nome: 'Logística' },
  { id: 'SAC', nome: 'Atendimento' },
  { id: 'CRE', nome: 'Acesso a Crédito' },
  { id: 'CRD', nome: 'Crédito' },
] as const;

/** Ordem de exibição das áreas em dashboards e relatórios. */
export const AREAS_ORDER = [
  'Estratégico',
  'Financeiro',
  'Acesso a Crédito',
  'Crédito',
  'Marketing',
  'Vendas',
  'Operacional',
  'Recursos Humanos',
  'Tecnologia',
  'Logística',
  'Jurídico',
  'Atendimento',
  'Geral',
];

/** Tipos / segmentos de empresa. */
export const TIPOS_EMPRESA = [
  'Geral',
  'Comércio',
  'Serviços',
  'Indústria',
  'Agronegócio',
  'Alimentos e Bebidas',
];

/** Ordem numérica de impacto para ordenação. */
export const IMPACTO_ORDER: Record<string, number> = {
  Baixo: 1,
  Médio: 2,
  Alto: 3,
};
