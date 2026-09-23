/**
 * Remove duplicatas da Fase 1 do App.tsx e usa imports reais (sem alias _).
 *
 * IMPORTANTE: versao segura — remove so blocos de funcao/const/interface
 * por matching de chaves ou marcadores estreitos. NAO apaga imports
 * (recharts, Button, Modal, cryptoStorage, etc.).
 *
 * Pre-requisito:
 *   node scripts/wire-phase2-app.mjs
 *
 * Uso:
 *   node scripts/remove-phase1-duplicates.mjs
 */
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve } from 'path';

const appPath = resolve('src/App.tsx');
if (!existsSync(appPath)) {
  console.error('src/App.tsx nao encontrado. Rode na raiz do ConsultorPro-II.');
  process.exit(1);
}

let text = readFileSync(appPath, 'utf8');
const originalLen = text.length;

if (!text.includes("from './features/auth/hooks/useAuth'")) {
  console.error('Rode primeiro: node scripts/wire-phase2-app.mjs');
  process.exit(1);
}

function removeUntil(src, startNeedle, endNeedle, label) {
  const start = src.indexOf(startNeedle);
  if (start < 0) {
    console.log(`· ${label}: inicio nao encontrado (ja removido?)`);
    return src;
  }
  const end = src.indexOf(endNeedle, start + startNeedle.length);
  if (end < 0) {
    console.log(`· ${label}: fim nao encontrado — pulando`);
    return src;
  }
  console.log(`✓ ${label}: removidos ~${end - start} chars`);
  return src.slice(0, start) + src.slice(end);
}

function removeBraceBlock(src, startNeedle, label) {
  const start = src.indexOf(startNeedle);
  if (start < 0) {
    console.log(`· ${label}: nao encontrado (ja removido?)`);
    return src;
  }
  let paren = src.indexOf('(', start);
  if (paren < 0 || paren > start + 80) paren = start;
  let pDepth = 0;
  let afterParams = -1;
  for (let i = paren; i < Math.min(src.length, start + 400); i++) {
    if (src[i] === '(') pDepth++;
    else if (src[i] === ')') {
      pDepth--;
      if (pDepth === 0) {
        afterParams = i + 1;
        break;
      }
    }
  }
  let searchFrom = afterParams > 0 ? afterParams : start;
  const arrow = src.indexOf('=>', searchFrom);
  if (arrow > 0 && arrow < searchFrom + 80) searchFrom = arrow;
  const brace = src.indexOf('{', searchFrom);
  if (brace < 0 || brace > start + 500) {
    console.log(`· ${label}: '{' do corpo nao encontrada — pulando`);
    return src;
  }
  let depth = 0;
  for (let j = brace; j < src.length; j++) {
    const c = src[j];
    if (c === '{') depth++;
    else if (c === '}') {
      depth--;
      if (depth === 0) {
        let end = j + 1;
        if (end < src.length && src[end] === ';') end++;
        while (end < src.length && (src[end] === '\n' || src[end] === '\r')) end++;
        console.log(`✓ ${label}: removidos ~${end - start} chars`);
        return src.slice(0, start) + src.slice(end);
      }
    }
  }
  console.log(`· ${label}: chaves nao balanceadas — pulando`);
  return src;
}

function removeOneLiner(src, pattern, label) {
  const re = new RegExp(pattern, 'm');
  const m = src.match(re);
  if (!m) {
    console.log(`· ${label}: nao encontrado (ja removido?)`);
    return src;
  }
  const start = m.index;
  const end = start + m[0].length;
  console.log(`✓ ${label}: removidos ~${end - start} chars`);
  return src.slice(0, start) + src.slice(end);
}

const oldImportBlockStart = '// --- Phase 1 extracted modules ---';
const oldImportBlockEnd = "import { LoginView } from './features/auth/components/LoginView';";

if (text.includes(oldImportBlockStart) && text.includes('as _cleanDigits')) {
  const start = text.indexOf(oldImportBlockStart);
  const endLine = text.indexOf(oldImportBlockEnd, start);
  if (endLine >= 0) {
    const end = endLine + oldImportBlockEnd.length;
    const newImports = `// --- Phase 1 extracted modules (sem duplicata no App) ---\nimport type {\n  Empresa,\n  EmpresaCredenciada,\n  Premissa,\n  Problema,\n  Diagnostico,\n  Resposta,\n  Solucao,\n  TarefaPlanoAcao,\n  DadosConsultoria,\n  AtividadeCronograma,\n} from './types/domain';\n\nexport type {\n  Empresa,\n  EmpresaCredenciada,\n  Premissa,\n  Problema,\n  Diagnostico,\n  Resposta,\n  Solucao,\n  TarefaPlanoAcao,\n  DadosConsultoria,\n  AtividadeCronograma,\n};\n\nimport {\n  cleanDigits,\n  formatCNPJ,\n  isValidCNPJ,\n  formatCEP,\n  formatCPF,\n  isValidCPF,\n} from './lib/formatters/br';\nimport {\n  getMdaExpirationStatus,\n  parseLocalDate,\n  formatFirestoreDate,\n  getActivityDateStr,\n} from './lib/formatters/dates';\nimport { sanitizeForFirestore } from './lib/firestore/sanitize';\nimport { oklchToRgb, captureElementWithHtml2Canvas } from './lib/pdf/capture';\nimport {\n  CONSULTORIA_AREAS,\n  AREAS,\n  AREAS_ORDER,\n  TIPOS_EMPRESA,\n  IMPACTO_ORDER,\n} from './lib/constants/areas';\nimport { LICENSE_FIELDS } from './lib/constants/license';\nimport { normalizeAndFormatArea, globalNormalizedMatch } from './lib/domain/areas';\nimport {\n  pickLicenseFields,\n  extractItemTimestamp,\n  toJsDate,\n  computeLicenseDaysLeft,\n} from './lib/domain/license';\nimport {\n  deduplicateRespostas,\n  loadAllLocalRespostas,\n  saveAllLocalRespostas,\n  getRespostasForDiagnostico,\n} from './lib/domain/respostas';\nimport { isValidLogoSource } from './lib/media/logo';\nimport { Type } from './lib/ai/schemaTypes';\nimport { extractAndParseJSON } from './lib/ai/parseJson';\nimport { useAuth } from './features/auth/hooks/useAuth';\nimport { LoginView } from './features/auth/components/LoginView';\n\nexport {\n  cleanDigits,\n  formatCNPJ,\n  isValidCNPJ,\n  formatCEP,\n  formatCPF,\n  isValidCPF,\n  getMdaExpirationStatus,\n  parseLocalDate,\n  formatFirestoreDate,\n  getActivityDateStr,\n  sanitizeForFirestore,\n  oklchToRgb,\n  captureElementWithHtml2Canvas,\n  CONSULTORIA_AREAS,\n  AREAS,\n  AREAS_ORDER,\n  TIPOS_EMPRESA,\n  IMPACTO_ORDER,\n  LICENSE_FIELDS,\n  normalizeAndFormatArea,\n  globalNormalizedMatch,\n  pickLicenseFields,\n  extractItemTimestamp,\n  toJsDate,\n  computeLicenseDaysLeft,\n  deduplicateRespostas,\n  loadAllLocalRespostas,\n  saveAllLocalRespostas,\n  getRespostasForDiagnostico,\n  isValidLogoSource,\n  Type,\n  extractAndParseJSON,\n};\n`;
    text = text.slice(0, start) + newImports + text.slice(end);
    console.log('✓ Imports realinhados (sem alias _) + reexports');
  }
} else if (text.includes("from './lib/formatters/br'") && !text.includes('as _cleanDigits')) {
  console.log('· Imports ja realinhados');
} else {
  console.log('· Bloco de imports Fase 1 nao encontrado no formato esperado');
}

{
  const re = /export enum Type \{[^}]*\}\n?/;
  const m = text.match(re);
  if (m) {
    text = text.replace(re, '');
    console.log(`✓ enum Type: removidos ~${m[0].length} chars`);
  } else {
    console.log('· enum Type: ja removido');
  }
}

text = removeUntil(text, 'const CONSULTORIA_AREAS = [', '// --- Utilitarios de Validacao e Formatacao', 'CONSULTORIA_AREAS');
text = removeUntil(text, 'const CONSULTORIA_AREAS = [', '// --- Utilitários de Validação e Formatação', 'CONSULTORIA_AREAS utf8');

const braceTargets = [
  ['export function cleanDigits(val?: string | null): string {', 'cleanDigits'],
  ['export function formatCNPJ(val?: string | null): string {', 'formatCNPJ'],
  ['export function isValidCNPJ(val?: string | null): boolean {', 'isValidCNPJ'],
  ['export function getMdaExpirationStatus(', 'getMdaExpirationStatus'],
  ['export function formatCEP(val?: string | null): string {', 'formatCEP'],
  ['export function formatCPF(val?: string | null): string {', 'formatCPF'],
  ['export function isValidCPF(val?: string | null): boolean {', 'isValidCPF'],
  ['function sanitizeForFirestore(obj: any): any {', 'sanitizeForFirestore'],
  ['function oklchToRgb(oklchStr: string): string {', 'oklchToRgb'],
  ['const captureElementWithHtml2Canvas = async (elt: HTMLElement) => {', 'captureElement'],
  ['const isValidLogoSource = (src: any): boolean => {', 'isValidLogoSource'],
  ['export function parseLocalDate(val: any): Date | null {', 'parseLocalDate'],
  ["function formatFirestoreDate(date: any, formatStr: string = 'dd/MM/yyyy', options?: any): string {", 'formatFirestoreDate'],
  ['export const getActivityDateStr = (diagDate: any, index: number): string => {', 'getActivityDateStr'],
  ['export function pickLicenseFields(item: any): Record<string, any> {', 'pickLicenseFields'],
  ['export function extractItemTimestamp(item: any): number {', 'extractItemTimestamp'],
  ['export const deduplicateRespostas = (respostasList: Resposta[]): Resposta[] => {', 'deduplicateRespostas'],
  ['export const loadAllLocalRespostas = (): Resposta[] => {', 'loadAllLocalRespostas'],
  ['export const saveAllLocalRespostas = (newOrUpdated: Resposta[]) => {', 'saveAllLocalRespostas'],
  ['export const getRespostasForDiagnostico = (diagId: string): Resposta[] => {', 'getRespostasForDiagnostico'],
  ['const normalizeAndFormatArea = (', 'normalizeAndFormatArea'],
  ['const globalNormalizedMatch = (', 'globalNormalizedMatch'],
];

for (const [needle, label] of braceTargets) {
  text = removeBraceBlock(text, needle, label);
}

text = removeUntil(
  text,
  'export const extractAndParseJSON = (text: string, defaultValue: any = null): any => {',
  '// A chave do criador NÃO fica mais no código',
  'extractAndParseJSON'
);

text = removeOneLiner(text, String.raw`const LICENSE_FIELDS = \[[^\]]*\];\s*\n?`, 'LICENSE_FIELDS');
if (text.includes('const LICENSE_FIELDS = [')) {
  text = removeUntil(text, 'const LICENSE_FIELDS = [', 'export function pickLicenseFields', 'LICENSE_FIELDS multi');
  if (text.includes('const LICENSE_FIELDS = [')) {
    text = removeUntil(text, 'const LICENSE_FIELDS = [', "import { toJsDate, computeLicenseDaysLeft }", 'LICENSE_FIELDS multi2');
  }
  if (text.includes('const LICENSE_FIELDS = [')) {
    text = removeUntil(text, 'const LICENSE_FIELDS = [', 'export function extractItemTimestamp', 'LICENSE_FIELDS multi3');
  }
}

text = removeOneLiner(text, String.raw`const IMPACTO_ORDER: Record<string, number> = \{[^}]*\};\s*\n?`, 'IMPACTO_ORDER');

{
  const re = /const AREAS = \[[\s\S]*?\];\s*\n?/;
  const m = text.match(re);
  if (m && m[0].length < 2000) {
    text = text.replace(re, '');
    console.log(`✓ AREAS: removidos ~${m[0].length} chars`);
  } else {
    console.log('· AREAS: nao encontrado ou bloco suspeito');
  }
}

{
  let ifaceStart = text.indexOf('interface EmpresaCredenciada {');
  if (ifaceStart < 0) ifaceStart = text.indexOf('export interface EmpresaCredenciada {');
  if (ifaceStart >= 0) {
    let resp = text.indexOf('export interface Resposta {', ifaceStart);
    if (resp < 0) resp = text.indexOf('interface Resposta {', ifaceStart);
    if (resp >= 0) {
      const brace = text.indexOf('{', resp);
      let depth = 0;
      let end = resp;
      for (let j = brace; j < text.length; j++) {
        if (text[j] === '{') depth++;
        else if (text[j] === '}') {
          depth--;
          if (depth === 0) {
            end = j + 1;
            break;
          }
        }
      }
      while (end < text.length && ' \n\r'.includes(text[end])) end++;
      console.log(`✓ interfaces de dominio: removidos ~${end - ifaceStart} chars`);
      text = text.slice(0, ifaceStart) + text.slice(end);
    }
  } else {
    console.log('· interfaces de dominio: ja removidas');
  }
}

text = text.replace(
  /\nimport \{ toJsDate, computeLicenseDaysLeft \} from '\.\/lib\/licenseDays';\nexport \{ toJsDate, computeLicenseDaysLeft \};\n/g,
  '\n'
);

text = removeUntil(
  text,
  'const TIPOS_EMPRESA = [\n  "Geral"',
  '// --- Components ---\n// (We moved Modal and Button to separate files)',
  'TIPOS_EMPRESA + AREAS_ORDER'
);
text = removeUntil(
  text,
  "const TIPOS_EMPRESA = [\n  'Geral'",
  '// --- Components ---\n// (We moved Modal and Button to separate files)',
  'TIPOS_EMPRESA + AREAS_ORDER (alt)'
);

text = text.replace(/let cachedAllRespostasMap: Map<string, Resposta> \| null = null;\s*\n?/g, '');

if (text.includes('| null {') || text.includes('if (!dataValidade) return null;')) {
  const junkStart = text.indexOf('// --- Utilitários de Validação e Formatação (Padrão Brasileiro - CPF e CNPJ) ---');
  let junkEnd = -1;
  const op = text.indexOf('enum OperationType {', junkStart >= 0 ? junkStart : 0);
  const iq = text.indexOf('function isQuotaError(error: unknown): boolean {');
  if (op >= 0 && (junkStart < 0 || op > junkStart)) junkEnd = op;
  else if (iq >= 0) junkEnd = iq;
  if (junkStart >= 0 && junkEnd > junkStart && junkEnd - junkStart < 8000) {
    text = text.slice(0, junkStart) + text.slice(junkEnd);
    console.log('✓ limpeza de fragmentos orfaos de formatters');
  }
}

if (!text.includes('enum OperationType {')) {
  const iq = text.indexOf('function isQuotaError(error: unknown): boolean {');
  if (iq >= 0) {
    const restore = `enum OperationType {\n  CREATE = 'create',\n  UPDATE = 'update',\n  DELETE = 'delete',\n  LIST = 'list',\n  GET = 'get',\n  WRITE = 'write',\n}\n\ninterface FirestoreErrorInfo {\n  error: string;\n  operationType: OperationType;\n  path: string | null;\n  authInfo: {\n    userId: string | undefined;\n    email: string | null | undefined;\n    emailVerified: boolean | undefined;\n    isAnonymous: boolean | undefined;\n    tenantId: string | null | undefined;\n    providerInfo: {\n      providerId: string;\n      displayName: string | null;\n      email: string | null;\n      photoUrl: string | null;\n    }[];\n  }\n}\n\n`;
    text = text.slice(0, iq) + restore + text.slice(iq);
    console.log('✓ enum OperationType + FirestoreErrorInfo restaurados');
  }
}

const critical = [
  "from 'recharts'",
  "from './components/Button'",
  "from './components/Modal'",
  "from './lib/cryptoStorage'",
  "from './lib/formatters/br'",
  'LoginView',
  'playSuccessSound',
  'enum OperationType',
];
let ok = true;
for (const needle of critical) {
  if (!text.includes(needle)) {
    console.error(`✗ CRITICO ausente apos remocao: ${needle}`);
    ok = false;
  }
}
if (!ok) {
  console.error('\nAbortando gravacao — restaure App.tsx com: git checkout -- src/App.tsx');
  process.exit(1);
}
console.log('✓ Imports criticos preservados');

writeFileSync(appPath, text, 'utf8');
const newLen = text.length;
console.log(`\nAntes:  ${originalLen} bytes`);
console.log(`Depois: ${newLen} bytes`);
console.log(`Reducao: ${originalLen - newLen} bytes (${(((originalLen - newLen) / originalLen) * 100).toFixed(1)}%)`);
console.log(`\nPronto. Rode: npm run lint && npm test`);
