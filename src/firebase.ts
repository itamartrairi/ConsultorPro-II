import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth } from 'firebase/auth';
import { 
  getFirestore, 
  initializeFirestore, 
  persistentLocalCache, 
  persistentSingleTabManager,
  memoryLocalCache,
  Firestore,
  collection, 
  getDocs, 
  addDoc, 
  doc, 
  getDocFromServer,
  setDoc,
  serverTimestamp
} from 'firebase/firestore';
import firebaseConfig from '../firebase-applet-config.json';

// Safely clean any corrupted or encrypted localStorage keys that break Firebase's internal cache parser
if (typeof window !== 'undefined' && window.localStorage) {
  try {
    const keysToRemove: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && (key.startsWith('firebase:') || key.startsWith('firestore:') || key.includes('firestore') || key.includes('applet_token'))) {
        const val = localStorage.getItem(key);
        if (val && (val.startsWith('enc:') || val === 'undefined' || val === 'null' || !val.startsWith('{') && !val.startsWith('[') && !val.startsWith('"') && !val.startsWith("'"))) {
          keysToRemove.push(key);
        }
      }
    }
    keysToRemove.forEach(k => localStorage.removeItem(k));
  } catch (e) {
    // Ignore storage cleaning errors
  }
}

// Initialize Firebase SDK
const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();
const firestoreDbId = (firebaseConfig as any).firestoreDatabaseId;

function initFirestoreInstance(): Firestore {
  try {
    if (firestoreDbId && firestoreDbId !== '(default)') {
      return initializeFirestore(app, {
        localCache: persistentLocalCache({
          tabManager: persistentSingleTabManager({})
        })
      }, firestoreDbId);
    }
    return initializeFirestore(app, {
      localCache: persistentLocalCache({
        tabManager: persistentSingleTabManager({})
      })
    });
  } catch {
    try {
      if (firestoreDbId && firestoreDbId !== '(default)') {
        return getFirestore(app, firestoreDbId);
      }
      return getFirestore(app);
    } catch {
      return initializeFirestore(app, {
        localCache: memoryLocalCache({})
      }, firestoreDbId && firestoreDbId !== '(default)' ? firestoreDbId : undefined);
    }
  }
}

export const db = initFirestoreInstance();
export const auth = getAuth(app);

// Connection verification helper for UI diagnostics
export async function checkCloudConnection(): Promise<{ ok: boolean; message: string; details?: string }> {
  try {
    const testDocRef = doc(db, 'system_health', 'connection_probe');
    await setDoc(testDocRef, {
      lastProbe: serverTimestamp(),
      app: 'Gestor Consultor Pro',
      probeDate: new Date().toISOString()
    }, { merge: true });
    return { ok: true, message: 'Conexão ativa e permissão de gravação validada no Firestore!' };
  } catch (error: any) {
    console.error("Firestore health probe failed:", error);
    const msg = error?.message || String(error);
    if (msg.includes('permission-denied') || msg.includes('Missing or insufficient permissions')) {
      return {
        ok: false,
        message: 'Permissão negada pelas Regras do Firestore.',
        details: 'As Regras de Segurança (Security Rules) no console do Firebase precisam ser publicadas para permitir leitura e escrita do seu usuário.'
      };
    }
    if (msg.includes('offline') || msg.includes('unavailable') || msg.includes('could not reach')) {
      return {
        ok: false,
        message: 'Modo Offline: Não foi possível alcançar o servidor do Firestore.',
        details: 'Verifique a conexão de rede ou as configurações do projeto no Firebase.'
      };
    }
    return {
      ok: false,
      message: 'Erro ao conectar com o banco de dados.',
      details: msg
    };
  }
}

// Initial probe on app start
async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error) {
      const msg = error.message.toLowerCase();
      if (msg.includes('the client is offline') || msg.includes('unavailable') || msg.includes('could not reach')) {
        console.warn("Firestore is operating in offline mode. Please check connection if live features are needed:", error.message);
      } else {
        console.warn("Firestore initialization status info:", error.message);
      }
    }
  }
}
testConnection();

// Initial seed data for Problems (Problemas)
const INITIAL_PROBLEMAS = [
  { descricao_problemas: "Falta de controle financeiro", area: "Financeiro", impacto: "Alto" },
  { descricao_problemas: "Baixa visibilidade da marca", area: "Marketing", impacto: "Médio" },
  { descricao_problemas: "Processos operacionais ineficientes", area: "Operacional", impacto: "Alto" },
  { descricao_problemas: "Alta rotatividade de funcionários", area: "Recursos Humanos", impacto: "Alto" },
  { descricao_problemas: "Baixa conversão de vendas", area: "Vendas", impacto: "Alto" },
  { descricao_problemas: "Infraestrutura tecnológica obsoleta", area: "Tecnologia", impacto: "Médio" },
  { descricao_problemas: "Insegurança jurídica em contratos", area: "Jurídico", impacto: "Médio" },
  { descricao_problemas: "Falta de planejamento estratégico", area: "Estratégico", impacto: "Alto" },
  { descricao_problemas: "Gargalos na cadeia logística", area: "Logística", impacto: "Médio" },
  { descricao_problemas: "Baixa satisfação do cliente", area: "Atendimento", impacto: "Alto" },
];

// Initial seed data for Premises (Premissas)
// These will be linked to the problems above during seeding
const INITIAL_PREMISSAS_TEMPLATE = [
  { area: "Financeiro", pergunta: "A empresa possui controle diário de entradas e saídas de caixa?", peso: 5 },
  { area: "Financeiro", pergunta: "Existe um planejamento financeiro para os próximos 6 meses?", peso: 4 },
  { area: "Financeiro", pergunta: "As contas pessoais dos sócios estão separadas das contas da empresa?", peso: 5 },
  { area: "Marketing", pergunta: "A empresa possui presença ativa em redes sociais?", peso: 3 },
  { area: "Marketing", pergunta: "Existe um processo claro de captação de novos clientes?", peso: 4 },
  { area: "Operacional", pergunta: "Os processos internos estão documentados e padronizados?", peso: 4 },
  { area: "Operacional", pergunta: "A equipe utiliza o sistema de gestão de forma eficiente?", peso: 3 },
  { area: "Recursos Humanos", pergunta: "Existe um plano de treinamento e capacitação para a equipe?", peso: 3 },
  { area: "Recursos Humanos", pergunta: "As metas individuais e coletivas estão claramente definidas?", peso: 4 },
  { area: "Vendas", pergunta: "A empresa possui um funil de vendas estruturado?", peso: 5 },
  { area: "Tecnologia", pergunta: "Os dados da empresa estão protegidos por backups regulares?", peso: 5 },
  { area: "Jurídico", pergunta: "Os contratos com fornecedores e clientes estão atualizados e revisados?", peso: 3 },
  { area: "Estratégico", pergunta: "A empresa possui uma visão clara de onde quer estar em 5 anos?", peso: 4 },
  { area: "Logística", pergunta: "O processo de entrega/distribuição é monitorado por indicadores?", peso: 3 },
  { area: "Atendimento", pergunta: "Existe um canal formal para reclamações e feedbacks de clientes?", peso: 4 },
];

export async function seedPremissasIfEmpty() {
  // Disabled: do not auto-seed experimental data so user has full control over a clean database.
  return;
}
