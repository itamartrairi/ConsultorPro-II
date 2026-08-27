import React, { useState, useEffect, useCallback } from 'react';
import { Cloud, CloudOff, CheckCircle2, AlertCircle, RefreshCw, Database } from 'lucide-react';
import { checkCloudConnection } from '../firebase';

export interface ConnectionStatusProps {
  className?: string;
  showDetails?: boolean;
  onStatusChange?: (status: { ok: boolean; message: string; details?: string } | null) => void;
}

export const ConnectionStatus: React.FC<ConnectionStatusProps> = ({
  className = '',
  showDetails = true,
  onStatusChange
}) => {
  const [status, setStatus] = useState<{
    checking: boolean;
    ok?: boolean;
    message?: string;
    details?: string;
    checkedAt?: string;
  }>({
    checking: true
  });

  const checkConnection = useCallback(async () => {
    setStatus(prev => ({ ...prev, checking: true }));
    try {
      const res = await checkCloudConnection();
      const newStatus = {
        checking: false,
        ok: res.ok,
        message: res.message,
        details: res.details,
        checkedAt: new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      };
      setStatus(newStatus);
      if (onStatusChange) {
        onStatusChange({ ok: res.ok, message: res.message, details: res.details });
      }
    } catch (e: any) {
      const errStatus = {
        checking: false,
        ok: false,
        message: 'Falha ao testar conexão com a nuvem',
        details: e?.message || String(e),
        checkedAt: new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' })
      };
      setStatus(errStatus);
      if (onStatusChange) {
        onStatusChange({ ok: false, message: errStatus.message, details: errStatus.details });
      }
    }
  }, [onStatusChange]);

  // Executa automaticamente a checagem ao carregar o componente
  useEffect(() => {
    checkConnection();
  }, [checkConnection]);

  return (
    <div className={`rounded-2xl border p-5 transition-all duration-300 ${
      status.checking 
        ? 'bg-slate-50/80 border-slate-200 text-slate-700'
        : status.ok 
          ? 'bg-emerald-50/70 border-emerald-200/80 text-emerald-950 shadow-sm shadow-emerald-100/50'
          : 'bg-amber-50/70 border-amber-200/80 text-amber-950 shadow-sm shadow-amber-100/50'
    } ${className}`}>
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex items-center gap-3.5">
          {/* Indicador visual de nuvem colorida */}
          <div className={`relative p-3 rounded-2xl flex items-center justify-center transition-transform duration-300 shrink-0 ${
            status.checking
              ? 'bg-indigo-100 text-indigo-600 animate-pulse'
              : status.ok
                ? 'bg-emerald-100 text-emerald-600 ring-4 ring-emerald-500/10'
                : 'bg-amber-100 text-amber-600 ring-4 ring-amber-500/10'
          }`}>
            {status.checking ? (
              <Cloud className="w-6 h-6 animate-pulse" />
            ) : status.ok ? (
              <Cloud className="w-6 h-6 text-emerald-600" />
            ) : (
              <CloudOff className="w-6 h-6 text-amber-600" />
            )}
            
            {/* Ponto indicador de status (badge de pulso) */}
            <span className={`absolute -top-1 -right-1 flex h-3 w-3`}>
              <span className={`animate-ping absolute inline-flex h-full w-full rounded-full opacity-75 ${
                status.checking ? 'bg-indigo-400' : status.ok ? 'bg-emerald-400' : 'bg-amber-400'
              }`}></span>
              <span className={`relative inline-flex rounded-full h-3 w-3 ${
                status.checking ? 'bg-indigo-500' : status.ok ? 'bg-emerald-500' : 'bg-amber-500'
              }`}></span>
            </span>
          </div>

          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5">
                Saúde do Banco de Dados
              </h4>
              <span className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full uppercase tracking-wider ${
                status.checking
                  ? 'bg-indigo-100/80 text-indigo-700'
                  : status.ok
                    ? 'bg-emerald-100 text-emerald-800'
                    : 'bg-amber-100 text-amber-800'
              }`}>
                {status.checking ? 'Verificando...' : status.ok ? '✓ Conectado & Operacional' : '⚠ Atenção / Restrito'}
              </span>
            </div>
            <p className="text-xs text-slate-600 mt-0.5">
              {status.checking
                ? 'Testando integridade das regras e permissões do Firestore...'
                : status.ok
                  ? 'Pronto para sincronização em tempo real entre todos os consultores.'
                  : (status.message || 'Verifique as regras de segurança no Console do Firebase.')}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2 self-start sm:self-center shrink-0">
          {status.checkedAt && (
            <span className="text-[10px] text-slate-600 font-mono hidden md:inline-block mr-1">
              Último check: {status.checkedAt}
            </span>
          )}
          <button
            type="button"
            onClick={checkConnection}
            disabled={status.checking}
            className={`text-xs font-bold px-3.5 py-2 rounded-xl flex items-center gap-1.5 transition-all shadow-sm active:scale-95 disabled:opacity-50 ${
              status.ok
                ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                : 'bg-amber-600 hover:bg-amber-700 text-white'
            }`}
          >
            <RefreshCw size={13} className={status.checking ? 'animate-spin' : ''} />
            <span>{status.checking ? 'Verificando...' : 'Reverificar'}</span>
          </button>
        </div>
      </div>

      {showDetails && status.details && !status.checking && (
        <div className="mt-3 pt-3 border-t border-slate-200/60 text-[11px] text-slate-600 flex items-start gap-2">
          {status.ok ? (
            <CheckCircle2 size={14} className="text-emerald-600 shrink-0 mt-0.5" />
          ) : (
            <AlertCircle size={14} className="text-amber-600 shrink-0 mt-0.5" />
          )}
          <div className="leading-relaxed">
            <span className="font-semibold text-slate-700">{status.details}</span>
          </div>
        </div>
      )}
    </div>
  );
};

export default ConnectionStatus;
