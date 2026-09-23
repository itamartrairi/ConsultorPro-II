/**
 * Hook e contexto de autenticação (Fase 2 da refatoração).
 */

import React, {
  createContext,
  useContext,
  useEffect,
  useMemo,
  useState,
  useCallback,
  type ReactNode,
} from 'react';
import {
  GoogleAuthProvider,
  onAuthStateChanged,
  signInWithPopup,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  sendPasswordResetEmail,
  signOut,
  type User,
} from 'firebase/auth';
import { auth } from '../../../firebase';
import { setStorageUserId } from '../../../lib/cryptoStorage';

export type AuthMode = 'login' | 'register' | 'reset';

export interface AuthContextValue {
  user: User | null;
  loading: boolean;
  authError: string;
  authSuccess: string;
  authMode: AuthMode;
  authEmail: string;
  authPassword: string;
  isSendingResetEmail: boolean;
  setAuthMode: (mode: AuthMode) => void;
  setAuthEmail: (email: string) => void;
  setAuthPassword: (password: string) => void;
  clearMessages: () => void;
  loginWithGoogle: () => Promise<void>;
  loginWithEmail: () => Promise<void>;
  registerWithEmail: () => Promise<void>;
  resetPassword: () => Promise<void>;
  logout: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | null>(null);

export function AuthProvider({ children }: { children: ReactNode }) {
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);
  const [authError, setAuthError] = useState('');
  const [authSuccess, setAuthSuccess] = useState('');
  const [authMode, setAuthMode] = useState<AuthMode>('login');
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [isSendingResetEmail, setIsSendingResetEmail] = useState(false);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setLoading(false);
      try {
        setStorageUserId(u?.uid || null);
      } catch {
        /* ignore */
      }
    });
    return () => unsub();
  }, []);

  const clearMessages = useCallback(() => {
    setAuthError('');
    setAuthSuccess('');
  }, []);

  const loginWithGoogle = useCallback(async () => {
    clearMessages();
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (e: any) {
      console.error(e);
      setAuthError(e?.message || 'Falha ao entrar com Google.');
    }
  }, [clearMessages]);

  const loginWithEmail = useCallback(async () => {
    clearMessages();
    if (!authEmail.trim() || !authPassword) {
      setAuthError('Informe e-mail e senha.');
      return;
    }
    try {
      await signInWithEmailAndPassword(auth, authEmail.trim(), authPassword);
    } catch (e: any) {
      console.error(e);
      const code = e?.code || '';
      if (code.includes('user-not-found') || code.includes('wrong-password') || code.includes('invalid-credential')) {
        setAuthError('E-mail ou senha inválidos.');
      } else {
        setAuthError(e?.message || 'Falha ao entrar.');
      }
    }
  }, [authEmail, authPassword, clearMessages]);

  const registerWithEmail = useCallback(async () => {
    clearMessages();
    if (!authEmail.trim() || !authPassword) {
      setAuthError('Informe e-mail e senha.');
      return;
    }
    if (authPassword.length < 6) {
      setAuthError('A senha deve ter pelo menos 6 caracteres.');
      return;
    }
    try {
      await createUserWithEmailAndPassword(auth, authEmail.trim(), authPassword);
      setAuthSuccess('Conta criada com sucesso!');
    } catch (e: any) {
      console.error(e);
      if (e?.code === 'auth/email-already-in-use') {
        setAuthError('Este e-mail já está em uso.');
      } else {
        setAuthError(e?.message || 'Falha ao criar conta.');
      }
    }
  }, [authEmail, authPassword, clearMessages]);

  const resetPassword = useCallback(async () => {
    clearMessages();
    if (!authEmail.trim()) {
      setAuthError('Informe o e-mail para recuperar a senha.');
      return;
    }
    setIsSendingResetEmail(true);
    try {
      await sendPasswordResetEmail(auth, authEmail.trim());
      setAuthSuccess('E-mail de recuperação enviado. Verifique sua caixa de entrada.');
    } catch (e: any) {
      console.error(e);
      setAuthError(e?.message || 'Falha ao enviar e-mail de recuperação.');
    } finally {
      setIsSendingResetEmail(false);
    }
  }, [authEmail, clearMessages]);

  const logout = useCallback(async () => {
    clearMessages();
    try {
      await signOut(auth);
      try {
        setStorageUserId(null);
      } catch {
        /* ignore */
      }
    } catch (e: any) {
      console.error(e);
      setAuthError(e?.message || 'Falha ao sair.');
    }
  }, [clearMessages]);

  const value = useMemo<AuthContextValue>(
    () => ({
      user,
      loading,
      authError,
      authSuccess,
      authMode,
      authEmail,
      authPassword,
      isSendingResetEmail,
      setAuthMode,
      setAuthEmail,
      setAuthPassword,
      clearMessages,
      loginWithGoogle,
      loginWithEmail,
      registerWithEmail,
      resetPassword,
      logout,
    }),
    [
      user,
      loading,
      authError,
      authSuccess,
      authMode,
      authEmail,
      authPassword,
      isSendingResetEmail,
      clearMessages,
      loginWithGoogle,
      loginWithEmail,
      registerWithEmail,
      resetPassword,
      logout,
    ]
  );

  return React.createElement(AuthContext.Provider, { value }, children);
}

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error('useAuth deve ser usado dentro de <AuthProvider>.');
  }
  return ctx;
}
