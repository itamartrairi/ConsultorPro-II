/**
 * Remove duplicatas da Fase 1 do App.tsx via AST (ts-morph).
 *
 * Pre-requisito recomendado:
 *   node scripts/wire-phase2-app.mjs
 *
 * Uso:
 *   npx tsx scripts/ast-remove-phase1.ts
 *   npx tsx scripts/ast-remove-phase1.ts --dry-run
 */
import { Project, type SourceFile } from 'ts-morph';
import { resolve } from 'path';

const appPath = resolve('src/App.tsx');
const dryRun = process.argv.includes('--dry-run');

const REMOVE_FUNCTIONS = [
  'cleanDigits',
  'formatCNPJ',
  'isValidCNPJ',
  'formatCEP',
  'formatCPF',
  'isValidCPF',
  'getMdaExpirationStatus',
  'sanitizeForFirestore',
  'oklchToRgb',
  'parseLocalDate',
  'formatFirestoreDate',
  'pickLicenseFields',
  'extractItemTimestamp',
];

const REMOVE_VARIABLES = [
  'CONSULTORIA_AREAS',
  'AREAS',
  'AREAS_ORDER',
  'TIPOS_EMPRESA',
  'IMPACTO_ORDER',
  'LICENSE_FIELDS',
  'captureElementWithHtml2Canvas',
  'isValidLogoSource',
  'getActivityDateStr',
  'normalizeAndFormatArea',
  'globalNormalizedMatch',
  'deduplicateRespostas',
  'loadAllLocalRespostas',
  'saveAllLocalRespostas',
  'getRespostasForDiagnostico',
  'extractAndParseJSON',
  'cachedAllRespostasMap',
];

const REMOVE_ENUMS = ['Type'];

const REMOVE_INTERFACES = [
  'EmpresaCredenciada',
  'Empresa',
  'Premissa',
  'Problema',
  'AtividadeCronograma',
  'DadosConsultoria',
  'Diagnostico',
  'Solucao',
  'TarefaPlanoAcao',
  'Resposta',
];

const CRITICAL_SUBSTRINGS = [
  "from 'recharts'",
  "from './components/Button'",
  "from './components/Modal'",
  "from './lib/cryptoStorage'",
  'playSuccessSound',
  'enum OperationType',
];

function removeVariableByName(sourceFile: SourceFile, name: string): boolean {
  for (const stmt of sourceFile.getVariableStatements()) {
    const decls = stmt.getDeclarations();
    const match = decls.find((d) => d.getName() === name);
    if (!match) continue;
    if (decls.length === 1) stmt.remove();
    else match.remove();
    return true;
  }
  return false;
}

function ensurePhase1Imports(sourceFile: SourceFile): void {
  const PHASE1_MODULES = new Set([
    './types/domain',
    './lib/formatters/br',
    './lib/formatters/dates',
    './lib/firestore/sanitize',
    './lib/pdf/capture',
    './lib/constants/areas',
    './lib/constants/license',
    './lib/domain/areas',
    './lib/domain/license',
    './lib/domain/respostas',
    './lib/media/logo',
    './lib/ai/schemaTypes',
    './lib/ai/parseJson',
    './features/auth/hooks/useAuth',
    './features/auth/components/LoginView',
  ]);

  for (const imp of [...sourceFile.getImportDeclarations()]) {
    if (PHASE1_MODULES.has(imp.getModuleSpecifierValue())) imp.remove();
  }

  for (const exp of [...sourceFile.getExportDeclarations()]) {
    const mod = exp.getModuleSpecifierValue();
    if (mod && PHASE1_MODULES.has(mod)) {
      exp.remove();
      continue;
    }
    const names = exp.getNamedExports().map((n) => n.getName());
    if (
      !mod &&
      names.length > 0 &&
      names.every((n) => n === 'toJsDate' || n === 'computeLicenseDaysLeft')
    ) {
      exp.remove();
      console.log('✓ export residual toJsDate/computeLicenseDaysLeft removido');
    }
  }

  const imports = sourceFile.getImportDeclarations();
  const insertAt =
    imports.length > 0 ? imports[imports.length - 1].getChildIndex() + 1 : 0;

  const block = `// --- Phase 1 extracted modules (AST) ---
import type {
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
} from './types/domain';

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
};

import {
  cleanDigits,
  formatCNPJ,
  isValidCNPJ,
  formatCEP,
  formatCPF,
  isValidCPF,
} from './lib/formatters/br';
import {
  getMdaExpirationStatus,
  parseLocalDate,
  formatFirestoreDate,
  getActivityDateStr,
} from './lib/formatters/dates';
import { sanitizeForFirestore } from './lib/firestore/sanitize';
import { oklchToRgb, captureElementWithHtml2Canvas } from './lib/pdf/capture';
import {
  CONSULTORIA_AREAS,
  AREAS,
  AREAS_ORDER,
  TIPOS_EMPRESA,
  IMPACTO_ORDER,
} from './lib/constants/areas';
import { LICENSE_FIELDS } from './lib/constants/license';
import { normalizeAndFormatArea, globalNormalizedMatch } from './lib/domain/areas';
import {
  pickLicenseFields,
  extractItemTimestamp,
  toJsDate,
  computeLicenseDaysLeft,
} from './lib/domain/license';
import {
  deduplicateRespostas,
  loadAllLocalRespostas,
  saveAllLocalRespostas,
  getRespostasForDiagnostico,
} from './lib/domain/respostas';
import { isValidLogoSource } from './lib/media/logo';
import { Type } from './lib/ai/schemaTypes';
import { extractAndParseJSON } from './lib/ai/parseJson';
import { useAuth } from './features/auth/hooks/useAuth';
import { LoginView } from './features/auth/components/LoginView';

export {
  cleanDigits,
  formatCNPJ,
  isValidCNPJ,
  formatCEP,
  formatCPF,
  isValidCPF,
  getMdaExpirationStatus,
  parseLocalDate,
  formatFirestoreDate,
  getActivityDateStr,
  sanitizeForFirestore,
  oklchToRgb,
  captureElementWithHtml2Canvas,
  CONSULTORIA_AREAS,
  AREAS,
  AREAS_ORDER,
  TIPOS_EMPRESA,
  IMPACTO_ORDER,
  LICENSE_FIELDS,
  normalizeAndFormatArea,
  globalNormalizedMatch,
  pickLicenseFields,
  extractItemTimestamp,
  toJsDate,
  computeLicenseDaysLeft,
  deduplicateRespostas,
  loadAllLocalRespostas,
  saveAllLocalRespostas,
  getRespostasForDiagnostico,
  isValidLogoSource,
  Type,
  extractAndParseJSON,
};
`;

  sourceFile.insertStatements(insertAt, block);
  console.log('✓ Imports + reexports Fase 1 inseridos (limpos)');
}

function removeLicenseDaysReexport(sourceFile: SourceFile): void {
  for (const imp of sourceFile.getImportDeclarations()) {
    if (imp.getModuleSpecifierValue() === './lib/licenseDays') {
      imp.remove();
      console.log('✓ import ./lib/licenseDays removido (agora via domain/license)');
    }
  }
}

function assertCritical(sourceFile: SourceFile): void {
  const text = sourceFile.getFullText();
  const missing: string[] = [];
  for (const s of CRITICAL_SUBSTRINGS) {
    if (!text.includes(s)) missing.push(s);
  }
  if (!sourceFile.getEnum('OperationType')) {
    missing.push('enum OperationType (AST)');
  }
  if (missing.length) {
    console.error('\n✗ CRITICO ausente apos transformacao:');
    for (const m of missing) console.error(`  - ${m}`);
    console.error('\nAbortando. Restaure: git checkout -- src/App.tsx');
    process.exit(1);
  }
  console.log('✓ Imports / simbolos criticos preservados');
}

function main() {
  console.log(`AST remove Fase 1 → ${appPath}${dryRun ? ' (dry-run)' : ''}`);

  const project = new Project({
    skipAddingFilesFromTsConfig: true,
    manipulationSettings: { indentationText: '  ' as any },
    compilerOptions: {
      jsx: 4,
      target: 99,
      module: 99,
      esModuleInterop: true,
      allowSyntheticDefaultImports: true,
    },
  });

  const sourceFile = project.addSourceFileAtPath(appPath);
  const before = sourceFile.getFullText().length;

  for (const name of REMOVE_ENUMS) {
    const e = sourceFile.getEnum(name);
    if (e) {
      e.remove();
      console.log(`✓ enum ${name} removido`);
    } else console.log(`· enum ${name} ja ausente`);
  }

  for (const name of REMOVE_FUNCTIONS) {
    const fn = sourceFile.getFunction(name);
    if (fn) {
      fn.remove();
      console.log(`✓ function ${name} removida`);
    } else console.log(`· function ${name} ja ausente`);
  }

  for (const name of REMOVE_VARIABLES) {
    if (removeVariableByName(sourceFile, name)) {
      console.log(`✓ variable ${name} removida`);
    } else console.log(`· variable ${name} ja ausente`);
  }

  for (const name of REMOVE_INTERFACES) {
    const iface = sourceFile.getInterface(name);
    if (iface) {
      iface.remove();
      console.log(`✓ interface ${name} removida`);
    } else console.log(`· interface ${name} ja ausente`);
  }

  removeLicenseDaysReexport(sourceFile);
  ensurePhase1Imports(sourceFile);
  assertCritical(sourceFile);

  const after = sourceFile.getFullText().length;
  console.log(`\nAntes:  ${before} bytes`);
  console.log(`Depois: ${after} bytes`);
  console.log(
    `Reducao: ${before - after} bytes (${(((before - after) / before) * 100).toFixed(1)}%)`
  );

  if (dryRun) {
    console.log('\nDry-run: nenhuma gravacao.');
    return;
  }

  sourceFile.saveSync();
  console.log(`\nSalvo ${appPath}`);
  console.log('Rode: npm run lint && npm test');
}

main();
