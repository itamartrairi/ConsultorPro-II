import { auth } from '../firebase';

async function authHeaders(): Promise<Record<string, string>> {
  const token = await auth.currentUser?.getIdToken();
  return token ? { Authorization: `Bearer ${token}` } : {};
}

export interface DbSyncResult {
  ok: boolean;
  empresas?: any[];
  diagnosticos?: any[];
  respostas?: any[];
  tarefas?: any[];
  credenciadas?: any[];
  error?: string;
}

export async function checkNetlifyDatabaseStatus(): Promise<{ ok: boolean; configured: boolean; message: string; provider?: string }> {
  try {
    const res = await fetch('/api/db/status', { headers: await authHeaders() });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      return {
        ok: false,
        configured: false,
        message: err.message || `Erro no servidor (${res.status})`
      };
    }
    return await res.json();
  } catch (e: any) {
    return {
      ok: false,
      configured: false,
      message: 'Não foi possível conectar ao endpoint do banco de dados Netlify: ' + (e?.message || e)
    };
  }
}

// A identidade (usuário/admin) é verificada no servidor pelo token do Firebase.
export async function syncWithNetlifyDatabase(payload: {
  localEmpresas: any[];
  localDiagnosticos: any[];
  localRespostas: any[];
  localTarefas: any[];
  localCredenciadas: any[];
}): Promise<DbSyncResult> {
  const res = await fetch('/api/db/sync', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...(await authHeaders()) },
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `Erro de sincronização HTTP ${res.status}`);
  }

  return await res.json();
}
