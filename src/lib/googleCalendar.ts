// Integração com o Google Agenda do usuário logado.
// Usa o próprio login Google do Firebase para pedir a permissão da agenda
// (escopo calendar.events) e chama a API do Google Calendar direto do navegador.
import {
  GoogleAuthProvider,
  reauthenticateWithPopup,
  linkWithPopup,
  type User,
} from 'firebase/auth';

export const GOOGLE_CALENDAR_SCOPE = 'https://www.googleapis.com/auth/calendar.events';
const EVENTS_API = 'https://www.googleapis.com/calendar/v3/calendars/primary/events';
const TOKEN_TTL_MS = 55 * 60 * 1000; // o token do Google vale 1h; renovamos um pouco antes

export interface GoogleCalendarToken {
  token: string;
  expiresAt: number;
  email: string;
}

export interface GoogleAgendaEvento {
  id: string;
  titulo: string;
  data: string; // YYYY-MM-DD (data local)
  horaInicio: string; // HH:MM ('' quando é dia inteiro)
  horaFim: string;
  diaInteiro: boolean;
  local?: string;
  link?: string;
  consultorProId?: string; // preenchido quando o evento foi criado pelo app
}

export interface AgendaEventoParaGoogle {
  appId: string;
  titulo: string;
  data: string;
  horaInicio: string;
  horaFim?: string;
  descricao?: string;
  tipo?: string;
  status?: string;
  empresaNome?: string;
}

export class GoogleCalendarAuthError extends Error {
  constructor(message = 'A conexão com o Google Agenda expirou. Conecte novamente.') {
    super(message);
    this.name = 'GoogleCalendarAuthError';
  }
}

const storageKey = (uid: string) => `consultorpro_gcal_${uid}`;

export function getGoogleLoginEmail(user: any): string {
  const google = user?.providerData?.find((p: any) => p?.providerId === 'google.com');
  return google?.email || user?.email || '';
}

export function getStoredGoogleToken(uid?: string | null): GoogleCalendarToken | null {
  if (!uid) return null;
  try {
    const raw = localStorage.getItem(storageKey(uid));
    if (!raw) return null;
    const parsed = JSON.parse(raw) as GoogleCalendarToken;
    if (!parsed?.token || parsed.expiresAt < Date.now() + 60_000) {
      localStorage.removeItem(storageKey(uid));
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

export function clearGoogleToken(uid?: string | null) {
  if (!uid) return;
  try {
    localStorage.removeItem(storageKey(uid));
  } catch {
    /* ignora */
  }
}

/**
 * Abre o popup do Google para a MESMA conta do usuário logado e pede a permissão da agenda.
 * Precisa ser chamada a partir de um clique (o navegador bloqueia popups automáticos).
 */
export async function connectGoogleCalendar(user: User): Promise<GoogleCalendarToken> {
  const provider = new GoogleAuthProvider();
  provider.addScope(GOOGLE_CALENDAR_SCOPE);
  const hasGoogleLogin = user.providerData.some((p) => p.providerId === 'google.com');
  const email = getGoogleLoginEmail(user);
  if (email) provider.setCustomParameters({ login_hint: email });

  let result;
  try {
    result = hasGoogleLogin
      ? await reauthenticateWithPopup(user, provider)
      : await linkWithPopup(user, provider);
  } catch (error: any) {
    const code = error?.code || '';
    if (code === 'auth/user-mismatch') {
      throw new Error(`Escolha a mesma conta Google usada no login (${email}).`);
    }
    if (code === 'auth/credential-already-in-use') {
      throw new Error('Esta conta Google já está vinculada a outro usuário do ConsultorPro.');
    }
    if (code === 'auth/popup-blocked') {
      throw new Error('O navegador bloqueou o popup do Google. Permita popups para este site e tente de novo.');
    }
    if (code === 'auth/popup-closed-by-user' || code === 'auth/cancelled-popup-request') {
      throw new Error('Conexão cancelada.');
    }
    throw new Error(error?.message || 'Não foi possível conectar ao Google Agenda.');
  }

  const credential = GoogleAuthProvider.credentialFromResult(result);
  if (!credential?.accessToken) {
    throw new Error('O Google não devolveu a permissão da agenda. Tente novamente.');
  }
  const googleInfo = result.user.providerData.find((p) => p.providerId === 'google.com');
  const stored: GoogleCalendarToken = {
    token: credential.accessToken,
    expiresAt: Date.now() + TOKEN_TTL_MS,
    email: googleInfo?.email || email,
  };
  try {
    localStorage.setItem(storageKey(user.uid), JSON.stringify(stored));
  } catch {
    /* sem armazenamento: o token vale só nesta tela */
  }
  return stored;
}

async function googleFetch(token: string, url: string, init: RequestInit = {}): Promise<Response> {
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  });
  if (res.status === 401) throw new GoogleCalendarAuthError();
  if (res.status === 403) {
    const body = await res.text();
    if (/insufficient|scope|PERMISSION_DENIED/i.test(body)) {
      throw new GoogleCalendarAuthError(
        'O app não tem permissão na sua agenda. Conecte novamente e marque a permissão do Google Agenda.'
      );
    }
    if (/accessNotConfigured|has not been used|disabled/i.test(body)) {
      throw new Error('A Google Calendar API não está ativada no projeto do Google Cloud deste app.');
    }
    throw new Error('O Google recusou o acesso à agenda (403).');
  }
  return res;
}

const pad = (n: number) => n.toString().padStart(2, '0');
const toDateStr = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const toTimeStr = (d: Date) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;

/** Lista os eventos da agenda principal do usuário entre duas datas. */
export async function listGoogleEvents(
  token: string,
  timeMin: Date,
  timeMax: Date
): Promise<GoogleAgendaEvento[]> {
  const out: GoogleAgendaEvento[] = [];
  let pageToken: string | undefined;
  let pages = 0;
  do {
    const params = new URLSearchParams({
      timeMin: timeMin.toISOString(),
      timeMax: timeMax.toISOString(),
      singleEvents: 'true',
      orderBy: 'startTime',
      maxResults: '250',
    });
    if (pageToken) params.set('pageToken', pageToken);
    const res = await googleFetch(token, `${EVENTS_API}?${params.toString()}`);
    if (!res.ok) throw new Error(`Erro ao ler o Google Agenda (${res.status}).`);
    const json = await res.json();
    for (const item of json.items || []) {
      if (item.status === 'cancelled') continue;
      const base = {
        id: item.id as string,
        titulo: (item.summary as string) || '(Sem título)',
        local: item.location as string | undefined,
        link: item.htmlLink as string | undefined,
        consultorProId: item.extendedProperties?.private?.consultorProId as string | undefined,
      };
      if (item.start?.date) {
        // Evento de dia inteiro: aparece em cada dia do período (a data final é exclusiva)
        const start = new Date(`${item.start.date}T12:00:00`);
        const end = new Date(`${(item.end?.date || item.start.date)}T12:00:00`);
        let d = new Date(start);
        let guard = 0;
        do {
          out.push({ ...base, id: `${base.id}_${toDateStr(d)}`, data: toDateStr(d), horaInicio: '', horaFim: '', diaInteiro: true });
          d.setDate(d.getDate() + 1);
          guard++;
        } while (d < end && guard < 62);
      } else if (item.start?.dateTime) {
        const start = new Date(item.start.dateTime);
        const end = new Date(item.end?.dateTime || item.start.dateTime);
        out.push({
          ...base,
          data: toDateStr(start),
          horaInicio: toTimeStr(start),
          horaFim: toTimeStr(end),
          diaInteiro: false,
        });
      }
    }
    pageToken = json.nextPageToken;
    pages++;
  } while (pageToken && pages < 5);
  return out;
}

function addOneHour(hora: string): string {
  const [h, m] = hora.split(':').map(Number);
  return `${pad(Math.min(h + 1, 23))}:${pad(m || 0)}`;
}

function buildGoogleEventBody(evt: AgendaEventoParaGoogle) {
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'America/Sao_Paulo';
  const inicio = evt.horaInicio || '09:00';
  let fim = evt.horaFim || '';
  if (!fim || fim <= inicio) fim = addOneHour(inicio);

  const prefixo = evt.status === 'Cancelado' ? '[Cancelado] ' : '';
  const titulo = evt.empresaNome ? `${evt.titulo} - ${evt.empresaNome}` : evt.titulo;
  const linhas = [
    evt.descricao?.trim() || '',
    evt.tipo ? `Tipo: ${evt.tipo}` : '',
    evt.empresaNome ? `Cliente: ${evt.empresaNome}` : '',
    evt.status ? `Status: ${evt.status}` : '',
    'Gerado pelo ConsultorPro II',
  ].filter(Boolean);

  return {
    summary: prefixo + titulo,
    description: linhas.join('\n'),
    location: evt.empresaNome || undefined,
    start: { dateTime: `${evt.data}T${inicio}:00`, timeZone },
    end: { dateTime: `${evt.data}T${fim}:00`, timeZone },
    extendedProperties: { private: { consultorProId: evt.appId } },
  };
}

/** Cria (ou atualiza, se já existir) o evento no Google Agenda. Devolve o id e o link do evento. */
export async function upsertGoogleEvent(
  token: string,
  googleEventId: string | null | undefined,
  evt: AgendaEventoParaGoogle
): Promise<{ id: string; link?: string }> {
  const body = JSON.stringify(buildGoogleEventBody(evt));
  if (googleEventId) {
    const res = await googleFetch(token, `${EVENTS_API}/${encodeURIComponent(googleEventId)}`, {
      method: 'PATCH',
      body,
    });
    if (res.ok) {
      const json = await res.json();
      return { id: json.id, link: json.htmlLink };
    }
    // Se o evento foi apagado direto no Google, criamos de novo
    if (res.status !== 404 && res.status !== 410) {
      throw new Error(`Erro ao atualizar o evento no Google Agenda (${res.status}).`);
    }
  }
  const res = await googleFetch(token, EVENTS_API, { method: 'POST', body });
  if (!res.ok) throw new Error(`Erro ao criar o evento no Google Agenda (${res.status}).`);
  const json = await res.json();
  return { id: json.id, link: json.htmlLink };
}

/** Exclui o evento do Google Agenda (ignora se ele já não existir). */
export async function deleteGoogleEvent(token: string, googleEventId: string): Promise<void> {
  const res = await googleFetch(token, `${EVENTS_API}/${encodeURIComponent(googleEventId)}`, {
    method: 'DELETE',
  });
  if (!res.ok && res.status !== 404 && res.status !== 410) {
    throw new Error(`Erro ao excluir o evento no Google Agenda (${res.status}).`);
  }
}
