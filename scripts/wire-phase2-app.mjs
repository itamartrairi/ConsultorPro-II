/**
 * Aplica o wiring da Fase 2 no App.tsx (idempotente).
 * Uso (na raiz do repo, branch refactor/phase1-utils):
 *   node scripts/wire-phase2-app.mjs
 */
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { resolve } from 'path';

const appPath = resolve('src/App.tsx');
if (!existsSync(appPath)) {
  console.error('src/App.tsx não encontrado. Rode na raiz do ConsultorPro-II.');
  process.exit(1);
}

let text = readFileSync(appPath, 'utf8');
const original = text;

const phase1Imports = `
// --- Phase 1 extracted modules ---
import type {
  Empresa as _EmpresaType,
  EmpresaCredenciada as _EmpresaCredenciadaType,
  Premissa as _PremissaType,
  Problema as _ProblemaType,
  Diagnostico as _DiagnosticoType,
  Resposta as _RespostaType,
  Solucao as _SolucaoType,
  TarefaPlanoAcao as _TarefaPlanoAcaoType,
  DadosConsultoria as _DadosConsultoriaType,
  AtividadeCronograma as _AtividadeCronogramaType,
} from './types/domain';
import {
  cleanDigits as _cleanDigits,
  formatCNPJ as _formatCNPJ,
  isValidCNPJ as _isValidCNPJ,
  formatCEP as _formatCEP,
  formatCPF as _formatCPF,
  isValidCPF as _isValidCPF,
} from './lib/formatters/br';
import {
  getMdaExpirationStatus as _getMdaExpirationStatus,
  parseLocalDate as _parseLocalDate,
  formatFirestoreDate as _formatFirestoreDate,
  getActivityDateStr as _getActivityDateStr,
} from './lib/formatters/dates';
import { sanitizeForFirestore as _sanitizeForFirestore } from './lib/firestore/sanitize';
import { oklchToRgb as _oklchToRgb, captureElementWithHtml2Canvas as _captureElementWithHtml2Canvas } from './lib/pdf/capture';
import {
  CONSULTORIA_AREAS as _CONSULTORIA_AREAS,
  AREAS as _AREAS,
  AREAS_ORDER as _AREAS_ORDER,
  TIPOS_EMPRESA as _TIPOS_EMPRESA,
  IMPACTO_ORDER as _IMPACTO_ORDER,
} from './lib/constants/areas';
import { LICENSE_FIELDS as _LICENSE_FIELDS } from './lib/constants/license';
import { normalizeAndFormatArea as _normalizeAndFormatArea, globalNormalizedMatch as _globalNormalizedMatch } from './lib/domain/areas';
import {
  pickLicenseFields as _pickLicenseFields,
  extractItemTimestamp as _extractItemTimestamp,
} from './lib/domain/license';
import {
  deduplicateRespostas as _deduplicateRespostas,
  loadAllLocalRespostas as _loadAllLocalRespostas,
  saveAllLocalRespostas as _saveAllLocalRespostas,
  getRespostasForDiagnostico as _getRespostasForDiagnostico,
} from './lib/domain/respostas';
import { isValidLogoSource as _isValidLogoSource } from './lib/media/logo';
import { Type as _Type } from './lib/ai/schemaTypes';
import { extractAndParseJSON as _extractAndParseJSON } from './lib/ai/parseJson';
import { useAuth } from './features/auth/hooks/useAuth';
import { LoginView } from './features/auth/components/LoginView';
`;

if (!text.includes("from './features/auth/hooks/useAuth'")) {
  const marker = "from './lib/cryptoStorage';";
  const idx = text.indexOf(marker);
  if (idx < 0) {
    console.error('Marker cryptoStorage não encontrado.');
    process.exit(1);
  }
  const end = text.indexOf('\n', idx);
  text = text.slice(0, end + 1) + phase1Imports + text.slice(end + 1);
  console.log('✓ Imports Fase 1 + Auth inseridos');
} else {
  console.log('· Imports já presentes');
}

const appFn = 'export default function App() {';
if (text.includes(appFn) && !text.includes('Phase 2: auth from context')) {
  const insert = `export default function App() {
  // Phase 2: auth from context (AuthProvider in main.tsx)
  const {
    user: authUser,
    loading: authLoading,
    logout: authLogout,
    authError: ctxAuthError,
    authSuccess: ctxAuthSuccess,
  } = useAuth();
`;
  text = text.replace(appFn, insert);
  console.log('✓ useAuth() no início de App()');
} else {
  console.log('· useAuth já presente ou App() não encontrado');
}

const loadingMarker = 'const [loading, setLoading] = useState(true);';
if (text.includes(loadingMarker) && !text.includes('Sync Firebase user from AuthProvider')) {
  text = text.replace(
    loadingMarker,
    `const [loading, setLoading] = useState(true);

  // Sync Firebase user from AuthProvider (single source of truth)
  useEffect(() => {
    setUser(authUser);
    setLoading(authLoading);
  }, [authUser, authLoading]);
`
  );
  console.log('✓ Sync user/loading com AuthProvider');
} else {
  console.log('· Sync effect já presente');
}

const loginStart = '  if (!user) {\n    return (\n      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">';
const loginEnd = '\n  return (\n    <div className="flex min-h-screen bg-slate-50 font-sans';
if (text.includes(loginStart) && text.includes('<LoginView />') === false) {
  const start = text.indexOf(loginStart);
  const end = text.indexOf(loginEnd, start);
  if (start >= 0 && end > start) {
    const replacement = `  if (!user) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <LoginView />
      </div>
    );
  }

`;
    text = text.slice(0, start) + replacement + text.slice(end);
    console.log('✓ Tela de login substituída por <LoginView />');
  }
} else if (text.includes('<LoginView />')) {
  console.log('· LoginView já em uso');
} else {
  console.log('· Bloco de login não encontrado (pode já ter sido alterado)');
}

if (text === original) {
  console.log('Nenhuma alteração necessária.');
} else {
  writeFileSync(appPath, text, 'utf8');
  console.log(`Salvo ${appPath} (${text.length} bytes)`);
}
