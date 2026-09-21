import admin from 'firebase-admin';
import { getFirestore, type Firestore } from 'firebase-admin/firestore';
import firebaseConfig from '../../firebase-applet-config.json';

/**
 * Firebase Admin compartilhado pelo server.ts (Docker/dev) e pelas Netlify Functions.
 *
 * - `getAdminApp()` precisa de FIREBASE_SERVICE_ACCOUNT_JSON (acesso ao Firestore,
 *   usado pelo webhook da Kiwify).
 * - `verifyFirebaseUser()` só confere o token de login e funciona mesmo sem a conta de
 *   serviço: basta o ID do projeto (lido do firebase-applet-config.json).
 */

let adminApp: admin.app.App | null = null;
let verifyApp: admin.app.App | null = null;

export function getAdminApp(): admin.app.App | null {
  if (adminApp) return adminApp;
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT_JSON;
  if (!raw) {
    console.error('[Firebase Admin] FIREBASE_SERVICE_ACCOUNT_JSON não configurada — sem acesso ao Firestore.');
    return null;
  }
  try {
    const serviceAccount = JSON.parse(raw);
    const existing = admin.apps.find((a) => a?.name === '[DEFAULT]');
    adminApp = existing ?? admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
    return adminApp;
  } catch (e) {
    console.error('[Firebase Admin] Falha ao inicializar:', e);
    return null;
  }
}

/**
 * O app usa um banco Firestore com nome próprio (firestoreDatabaseId), não o "(default)".
 * Sem isso, o webhook gravaria num banco vazio e nunca encontraria o comprador.
 */
export function getAdminDb(): Firestore | null {
  const app = getAdminApp();
  if (!app) return null;
  const dbId = process.env.FIREBASE_DATABASE_ID || (firebaseConfig as any).firestoreDatabaseId;
  return dbId && dbId !== '(default)' ? getFirestore(app, dbId) : getFirestore(app);
}

function getVerifyApp(): admin.app.App {
  if (verifyApp) return verifyApp;
  const withCreds = getAdminAppSilently();
  if (withCreds) {
    verifyApp = withCreds;
    return verifyApp;
  }
  const projectId = process.env.FIREBASE_PROJECT_ID || (firebaseConfig as any).projectId;
  const existing = admin.apps.find((a) => a?.name === 'token-verify');
  verifyApp = existing ?? admin.initializeApp({ projectId }, 'token-verify');
  return verifyApp;
}

function getAdminAppSilently(): admin.app.App | null {
  if (adminApp) return adminApp;
  if (!process.env.FIREBASE_SERVICE_ACCOUNT_JSON) return null;
  return getAdminApp();
}

export function adminEmails(): string[] {
  return (process.env.ADMIN_EMAILS || 'itamartrairi@gmail.com')
    .split(',')
    .map((e) => e.trim().toLowerCase())
    .filter(Boolean);
}

export interface VerifiedUser {
  uid: string;
  email: string;
  isAdmin: boolean;
}

/** Lê "Authorization: Bearer <idToken>" e confere o login no Firebase. */
export async function verifyFirebaseUser(authorizationHeader: string | null | undefined): Promise<VerifiedUser | null> {
  const header = authorizationHeader || '';
  const idToken = header.startsWith('Bearer ') ? header.slice(7).trim() : '';
  if (!idToken) return null;
  try {
    const decoded = await getVerifyApp().auth().verifyIdToken(idToken);
    const email = (decoded.email || '').toLowerCase();
    return {
      uid: decoded.uid,
      email,
      isAdmin: !!email && decoded.email_verified === true && adminEmails().includes(email),
    };
  } catch (e: any) {
    console.warn('[Firebase Admin] Token de login inválido:', e?.code || e?.message || e);
    return null;
  }
}
