/**
 * Inventário AST do App.tsx (somente leitura).
 *
 * Lista declarações de nível de módulo: functions, const/let, enums, interfaces, types, classes.
 *
 * Uso:
 *   npx tsx scripts/ast-inventory-app.ts
 *   npx tsx scripts/ast-inventory-app.ts --json
 *   npx tsx scripts/ast-inventory-app.ts --exports-only
 */
import { Project, type SourceFile, type Node } from 'ts-morph';
import { resolve } from 'path';

const appPath = resolve('src/App.tsx');
const args = process.argv.slice(2);
const asJson = args.includes('--json');
const exportsOnly = args.includes('--exports-only');

type Kind =
  | 'function'
  | 'variable'
  | 'enum'
  | 'interface'
  | 'type'
  | 'class'
  | 'other';

interface SymbolInfo {
  name: string;
  kind: Kind;
  exported: boolean;
  line: number;
  endLine: number;
  chars: number;
}

function lineOf(node: Node): number {
  return node.getStartLineNumber();
}

function endLineOf(node: Node): number {
  return node.getEndLineNumber();
}

function collect(sourceFile: SourceFile): SymbolInfo[] {
  const out: SymbolInfo[] = [];

  for (const fn of sourceFile.getFunctions()) {
    const name = fn.getName();
    if (!name) continue;
    out.push({
      name,
      kind: 'function',
      exported: fn.isExported() || fn.hasExportKeyword(),
      line: lineOf(fn),
      endLine: endLineOf(fn),
      chars: fn.getFullWidth(),
    });
  }

  for (const stmt of sourceFile.getVariableStatements()) {
    const exported = stmt.isExported() || stmt.hasExportKeyword();
    for (const d of stmt.getDeclarations()) {
      const name = d.getName();
      if (!name || name.startsWith('_')) continue;
      out.push({
        name,
        kind: 'variable',
        exported,
        line: lineOf(stmt),
        endLine: endLineOf(stmt),
        chars: stmt.getFullWidth(),
      });
    }
  }

  for (const e of sourceFile.getEnums()) {
    const name = e.getName();
    if (!name) continue;
    out.push({
      name,
      kind: 'enum',
      exported: e.isExported() || e.hasExportKeyword(),
      line: lineOf(e),
      endLine: endLineOf(e),
      chars: e.getFullWidth(),
    });
  }

  for (const i of sourceFile.getInterfaces()) {
    const name = i.getName();
    if (!name) continue;
    out.push({
      name,
      kind: 'interface',
      exported: i.isExported() || i.hasExportKeyword(),
      line: lineOf(i),
      endLine: endLineOf(i),
      chars: i.getFullWidth(),
    });
  }

  for (const ta of sourceFile.getTypeAliases()) {
    const name = ta.getName();
    if (!name) continue;
    out.push({
      name,
      kind: 'type',
      exported: ta.isExported() || ta.hasExportKeyword(),
      line: lineOf(ta),
      endLine: endLineOf(ta),
      chars: ta.getFullWidth(),
    });
  }

  for (const c of sourceFile.getClasses()) {
    const name = c.getName();
    if (!name) continue;
    out.push({
      name,
      kind: 'class',
      exported: c.isExported() || c.hasExportKeyword(),
      line: lineOf(c),
      endLine: endLineOf(c),
      chars: c.getFullWidth(),
    });
  }

  return out.sort((a, b) => a.line - b.line);
}

const PHASE1_CANDIDATES = new Set([
  'Type',
  'CONSULTORIA_AREAS',
  'AREAS',
  'AREAS_ORDER',
  'TIPOS_EMPRESA',
  'IMPACTO_ORDER',
  'LICENSE_FIELDS',
  'cleanDigits',
  'formatCNPJ',
  'isValidCNPJ',
  'formatCEP',
  'formatCPF',
  'isValidCPF',
  'getMdaExpirationStatus',
  'sanitizeForFirestore',
  'oklchToRgb',
  'captureElementWithHtml2Canvas',
  'isValidLogoSource',
  'parseLocalDate',
  'formatFirestoreDate',
  'getActivityDateStr',
  'pickLicenseFields',
  'extractItemTimestamp',
  'normalizeAndFormatArea',
  'globalNormalizedMatch',
  'deduplicateRespostas',
  'loadAllLocalRespostas',
  'saveAllLocalRespostas',
  'getRespostasForDiagnostico',
  'extractAndParseJSON',
  'Empresa',
  'EmpresaCredenciada',
  'Premissa',
  'Problema',
  'Diagnostico',
  'Resposta',
  'Solucao',
  'TarefaPlanoAcao',
  'DadosConsultoria',
  'AtividadeCronograma',
]);

function main() {
  console.error(`Lendo ${appPath} ...`);
  const project = new Project({
    skipAddingFilesFromTsConfig: true,
    compilerOptions: {
      allowJs: false,
      jsx: 4,
      target: 99,
      module: 99,
    },
  });
  const sourceFile = project.addSourceFileAtPath(appPath);
  let symbols = collect(sourceFile);
  if (exportsOnly) {
    symbols = symbols.filter((s) => s.exported);
  }

  if (asJson) {
    console.log(JSON.stringify(symbols, null, 2));
    return;
  }

  const phase1 = symbols.filter((s) => PHASE1_CANDIDATES.has(s.name));
  const others = symbols.filter((s) => !PHASE1_CANDIDATES.has(s.name));

  console.log('\n=== Candidatos Fase 1 (ainda no App.tsx) ===\n');
  if (phase1.length === 0) {
    console.log('(nenhum — Fase 1 provavelmente já removida)\n');
  } else {
    for (const s of phase1) {
      const exp = s.exported ? 'export' : 'local ';
      console.log(
        `  ${exp}  ${s.kind.padEnd(10)}  L${String(s.line).padStart(5)}-${String(s.endLine).padStart(5)}  ${String(s.chars).padStart(6)}c  ${s.name}`
      );
    }
    console.log(`\n  Total candidatos: ${phase1.length}\n`);
  }

  console.log('=== Outros símbolos de módulo ===\n');
  const byKind = new Map<Kind, number>();
  for (const s of others) {
    byKind.set(s.kind, (byKind.get(s.kind) || 0) + 1);
  }
  for (const [k, n] of [...byKind.entries()].sort()) {
    console.log(`  ${k}: ${n}`);
  }
  console.log(`\n  Total outros: ${others.length}`);
  console.log(`  Total geral: ${symbols.length}`);
  console.log(`\nDica: npx tsx scripts/ast-inventory-app.ts --json > /tmp/app-symbols.json`);
}

main();
