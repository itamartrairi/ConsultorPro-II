/**
 * Tela de login / registro / reset de senha (Fase 2).
 * UI desacoplada do App.tsx — usa useAuth().
 */

import React from 'react';
import { Loader2, Mail, Lock, LogIn } from 'lucide-react';
import { useAuth, type AuthMode } from '../hooks/useAuth';
import { Button } from '../../../components/Button';
import { cn } from '../../../lib/utils';

export interface LoginViewProps {
  /** Chamado após login bem-sucedido (opcional; App pode reagir via user). */
  onSuccess?: () => void;
  className?: string;
}

export function LoginView({ onSuccess, className }: LoginViewProps) {
  const {
    user,
    loading,
    authMode,
    authEmail,
    authPassword,
    authError,
    authSuccess,
    isSendingResetEmail,
    setAuthMode,
    setAuthEmail,
    setAuthPassword,
    clearMessages,
    loginWithGoogle,
    loginWithEmail,
    registerWithEmail,
    resetPassword,
  } = useAuth();

  React.useEffect(() => {
    if (user && onSuccess) onSuccess();
  }, [user, onSuccess]);

  if (loading) {
    return (
      <div className={cn('flex items-center justify-center min-h-[40vh]', className)}>
        <Loader2 className="animate-spin text-emerald-600" size={32} />
      </div>
    );
  }

  const switchMode = (mode: AuthMode) => {
    clearMessages();
    setAuthMode(mode);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (authMode === 'login') await loginWithEmail();
    else if (authMode === 'register') await registerWithEmail();
    else await resetPassword();
  };

  const title =
    authMode === 'login'
      ? 'Entrar'
      : authMode === 'register'
        ? 'Criar conta'
        : 'Recuperar senha';

  return (
    <div className={cn('w-full max-w-md mx-auto space-y-6', className)}>
      <div className="text-center space-y-1">
        <h1 className="text-2xl font-bold text-slate-800 tracking-tight">{title}</h1>
        <p className="text-sm text-slate-500">
          {authMode === 'reset'
            ? 'Enviaremos um link para redefinir sua senha.'
            : 'Acesse o Gestor Consultor Pro'}
        </p>
      </div>

      {authError && (
        <div className="rounded-xl border border-red-200 bg-red-50 text-red-800 text-sm px-4 py-3">
          {authError}
        </div>
      )}
      {authSuccess && (
        <div className="rounded-xl border border-emerald-200 bg-emerald-50 text-emerald-800 text-sm px-4 py-3">
          {authSuccess}
        </div>
      )}

      <form onSubmit={handleSubmit} className="space-y-4 bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
        <div>
          <label className="block text-sm font-semibold text-slate-700 mb-1">E-mail</label>
          <div className="relative">
            <Mail className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
            <input
              type="email"
              autoComplete="email"
              className="w-full pl-10 pr-3 py-2.5 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              value={authEmail}
              onChange={(e) => setAuthEmail(e.target.value)}
              placeholder="seu@email.com"
              required
            />
          </div>
        </div>

        {authMode !== 'reset' && (
          <div>
            <label className="block text-sm font-semibold text-slate-700 mb-1">Senha</label>
            <div className="relative">
              <Lock className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
              <input
                type="password"
                autoComplete={authMode === 'register' ? 'new-password' : 'current-password'}
                className="w-full pl-10 pr-3 py-2.5 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
                value={authPassword}
                onChange={(e) => setAuthPassword(e.target.value)}
                placeholder="••••••••"
                required
                minLength={6}
              />
            </div>
          </div>
        )}

        <Button type="submit" className="w-full bg-emerald-600 hover:bg-emerald-700 text-white" disabled={isSendingResetEmail}>
          {isSendingResetEmail ? (
            <>
              <Loader2 className="animate-spin mr-2" size={16} />
              Enviando...
            </>
          ) : (
            <>
              <LogIn className="mr-2" size={16} />
              {title}
            </>
          )}
        </Button>

        {authMode === 'login' && (
          <Button
            type="button"
            variant="outline"
            className="w-full border-slate-200"
            onClick={() => loginWithGoogle()}
          >
            Entrar com Google
          </Button>
        )}
      </form>

      <div className="flex flex-wrap items-center justify-center gap-3 text-sm">
        {authMode !== 'login' && (
          <button type="button" className="text-emerald-700 font-semibold hover:underline" onClick={() => switchMode('login')}>
            Já tenho conta
          </button>
        )}
        {authMode !== 'register' && (
          <button type="button" className="text-emerald-700 font-semibold hover:underline" onClick={() => switchMode('register')}>
            Criar conta
          </button>
        )}
        {authMode !== 'reset' && (
          <button type="button" className="text-slate-500 hover:underline" onClick={() => switchMode('reset')}>
            Esqueci a senha
          </button>
        )}
      </div>
    </div>
  );
}
