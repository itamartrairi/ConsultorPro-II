/**
 * Remove duplicatas da Fase 1 do App.tsx e usa imports reais (sem alias _).
 * Pré-requisito: node scripts/wire-phase2-app.mjs já rodado.
 *
 * Uso:
 *   node scripts/remove-phase1-duplicates.mjs
 */
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve } from 'path';

const appPath = resolve('src/App.tsx');
if (!existsSync(appPath)) {
  console.error('src/App.tsx não encontrado.');
  process.exit(1);
}

let text = readFileSync(appPath, 'utf8');
const originalLen = text.length;

if (!text.includes("from './features/auth/hooks/useAuth'")) {
  console.error('Rode primeiro: node scripts/wire-phase2-app.mjs');
  process.exit(1);
}

function removeBlock(src, startNeedle, endNeedle, label) {
  const start = src.indexOf(startNeedle);
  if (start < 0) {
    console.log(`· ${label}: início não encontrado (já removido?)`);
    return src;
  }
  const end = src.indexOf(endNeedle, start + startNeedle.length);
  if (end < 0) {
    console.log(`· ${label}: fim não encontrado — pulando`);
    return src;
  }
  const removed = end - start;
  console.log(`✓ ${label}: removidos ~${removed} chars`);
  return src.slice(0, start) + src.slice(end);
}

const oldImportBlockStart = '// --- Phase 1 extracted modules ---';
const oldImportBlockEnd = "import { LoginView } from './features/auth/components/LoginView';";

if (text.includes(oldImportBlockStart)) {
  const start = text.indexOf(oldImportBlockStart);
  const endLine = text.indexOf(oldImportBlockEnd, start);
  if (endLine >= 0) {
    const end = endLine + oldImportBlockEnd.length;
    const newImports = `// --- Phase 1 extracted modules (sem duplicata no App) ---\nimport type {\n  Empresa,\n  EmpresaCredenciada,\n  Premissa,\n  Problema,\n  Diagnostico,\n  Resposta,\n  Solucao,\n  TarefaPlanoAcao,\n  DadosConsultoria,\n  AtividadeCronograma,\n} from './types/domain';\n\nexport type {\n  Empresa,\n  EmpresaCredenciada,\n  Premissa,\n  Problema,\n  Diagnostico,\n  Resposta,\n  Solucao,\n  TarefaPlanoAcao,\n  DadosConsultoria,\n  AtividadeCronograma,\n};\n\nimport {\n  cleanDigits,\n  formatCNPJ,\n  isValidCNPJ,\n  formatCEP,\n  formatCPF,\n  isValidCPF,\n} from './lib/formatters/br';\nimport {\n  getMdaExpirationStatus,\n  parseLocalDate,\n  formatFirestoreDate,\n  getActivityDateStr,\n} from './lib/formatters/dates';\nimport { sanitizeForFirestore } from './lib/firestore/sanitize';\nimport { oklchToRgb, captureElementWithHtml2Canvas } from './lib/pdf/capture';\nimport {\n  CONSULTORIA_AREAS,\n  AREAS,\n  AREAS_ORDER,\n  TIPOS_EMPRESA,\n  IMPACTO_ORDER,\n} from './lib/constants/areas';\nimport { LICENSE_FIELDS } from './lib/constants/license';\nimport { normalizeAndFormatArea, globalNormalizedMatch } from './lib/domain/areas';\nimport {\n  pickLicenseFields,\n  extractItemTimestamp,\n  toJsDate,\n  computeLicenseDaysLeft,\n} from './lib/domain/license';\nimport {\n  deduplicateRespostas,\n  loadAllLocalRespostas,\n  saveAllLocalRespostas,\n  getRespostasForDiagnostico,\n} from './lib/domain/respostas';\nimport { isValidLogoSource } from './lib/media/logo';\nimport { Type } from './lib/ai/schemaTypes';\nimport { extractAndParseJSON } from './lib/ai/parseJson';\nimport { useAuth } from './features/auth/hooks/useAuth';\nimport { LoginView } from './features/auth/components/LoginView';\n\nexport {\n  cleanDigits,\n  formatCNPJ,\n  isValidCNPJ,\n  formatCEP,\n  formatCPF,\n  isValidCPF,\n  getMdaExpirationStatus,\n  parseLocalDate,\n  formatFirestoreDate,\n  getActivityDateStr,\n  sanitizeForFirestore,\n  oklchToRgb,\n  captureElementWithHtml2Canvas,\n  CONSULTORIA_AREAS,\n  AREAS,\n  AREAS_ORDER,\n  TIPOS_EMPRESA,\n  IMPACTO_ORDER,\n  LICENSE_FIELDS,\n  normalizeAndFormatArea,\n  globalNormalizedMatch,\n  pickLicenseFields,\n  extractItemTimestamp,\n  toJsDate,\n  computeLicenseDaysLeft,\n  deduplicateRespostas,\n  loadAllLocalRespostas,\n  saveAllLocalRespostas,\n  getRespostasForDiagnostico,\n  isValidLogoSource,\n  Type,\n  extractAndParseJSON,\n};\n`;
    text = text.slice(0, start) + newImports + text.slice(end);
    console.log('✓ Imports realinhados (sem alias _) + reexports');
  }
}

text = removeBlock(text, 'export enum Type {\n  STRING = "STRING"', '\nconst safeLocalStorage', 'enum Type');
text = removeBlock(text, "export enum Type {\n  STRING = 'STRING'", '\nconst safeLocalStorage', 'enum Type (alt)');
text = removeBlock(text, 'const CONSULTORIA_AREAS = [', 'export function cleanDigits', 'CONSULTORIA_AREAS');
text = removeBlock(text, 'export function cleanDigits(val?: string | null): string {', 'function sanitizeForFirestore(obj: any): any {', 'formatters BR');
text = removeBlock(text, 'function sanitizeForFirestore(obj: any): any {', 'function oklchToRgb(oklchStr: string): string {', 'sanitizeForFirestore');
text = removeBlock(text, 'function oklchToRgb(oklchStr: string): string {', 'const isValidLogoSource = (src: any): boolean => {', 'oklchToRgb + captureElement');
text = removeBlock(text, 'const isValidLogoSource = (src: any): boolean => {', 'export function parseLocalDate(val: any): Date | null {', 'isValidLogoSource');
text = removeBlock(text, 'export function parseLocalDate(val: any): Date | null {', 'const LICENSE_FIELDS =', 'parseLocalDate…getActivityDateStr');
text = removeBlock(text, 'const LICENSE_FIELDS =', 'enum OperationType {', 'LICENSE_FIELDS…extractItemTimestamp');
if (text.includes('const LICENSE_FIELDS =')) {
  text = removeBlock(text, 'const LICENSE_FIELDS =', 'interface EmpresaCredenciada {', 'LICENSE_FIELDS… (até EmpresaCredenciada)');
}
text = removeBlock(text, 'interface EmpresaCredenciada {', 'const AREAS = [', 'interfaces de domínio');
text = removeBlock(text, 'export interface EmpresaCredenciada {', 'const AREAS = [', 'interfaces de domínio (export)');
text = removeBlock(text, "const AREAS = [\n  { id: 'FIN'", 'const Card = (', 'AREAS + normalize + IMPACTO + match');
text = removeBlock(text, 'const AREAS = [\n  { id: "FIN"', 'const Card = (', 'AREAS alt');
text = removeBlock(text, 'export const deduplicateRespostas = (respostasList: Resposta[]): Resposta[] => {', 'const DadosConsultoriaView = (', 'deduplicateRespostas');
text = removeBlock(text, 'export const extractAndParseJSON = (text: string, defaultValue: any = null): any => {', '// A chave do criador NÃO fica mais no código', 'extractAndParseJSON');
if (text.includes('export const extractAndParseJSON')) {
  text = removeBlock(text, 'export const extractAndParseJSON = (text: string, defaultValue: any = null): any => {', "import { readCustomGeminiKey", 'extractAndParseJSON (alt)');
}
text = removeBlock(text, 'let cachedAllRespostasMap: Map<string, Resposta> | null = null;', 'export default function App() {', 'local respostas helpers');
if (text.includes('export const loadAllLocalRespostas')) {
  text = removeBlock(text, 'export const loadAllLocalRespostas = (): Resposta[] => {', 'export default function App() {', 'loadAllLocalRespostas…');
}
text = text.replace(/\nimport \{ toJsDate, computeLicenseDaysLeft \} from '\.\/lib\/licenseDays';\nexport \{ toJsDate, computeLicenseDaysLeft \};\n/g, '\n');
if (text.includes('export enum Type {')) {
  text = removeBlock(text, 'export enum Type {', '\nconst ', 'enum Type residual');
}
text = removeBlock(text, 'const TIPOS_EMPRESA = [\n  "Geral"', '// --- Components ---\n// (We moved Modal and Button to separate files)', 'TIPOS_EMPRESA + AREAS_ORDER');
text = removeBlock(text, "const TIPOS_EMPRESA = [\n  'Geral'", '// --- Components ---\n// (We moved Modal and Button to separate files)', 'TIPOS_EMPRESA + AREAS_ORDER (alt)');

writeFileSync(appPath, text, 'utf8');
const newLen = text.length;
console.log(`\nAntes:  ${originalLen} bytes`);
console.log(`Depois: ${newLen} bytes`);
console.log(`Redução: ${originalLen - newLen} bytes (${(((originalLen - newLen) / originalLen) * 100).toFixed(1)}%)`);
console.log(`\nPronto. Rode: npm run lint && npm test`);
