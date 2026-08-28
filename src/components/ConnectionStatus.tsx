import React, { useState, useEffect, useCallback } from 'react';
import { Cloud, CloudOff, CheckCircle2, AlertCircle, RefreshCw, Database } from 'lucide-react';
import { checkCloudConnection } from '../firebase';

export interface ConnectionStatusProps {
  className?: string;
  showDetails?: boolean;
  onStatusChange?: (status: { ok: boolean; message: string; details?: string } | null) => void;
}

export const ConnectionStatus: React.FC<ConnectionStatusProps> = ({ className = '', showDetails = false, onStatusChange }) => {
  const [status, setStatus] = useState<{ ok: boolean; message: string; details?: string } | null>(null);
  const [loading, setLoading] = useState(false);

  const check = useCallback(async () => {
    setLoading(true);
    try {
      const res = await checkCloudConnection();
      setStatus(res);
      if (onStatusChange) onStatusChange(res);
    } catch (e: any) {
      const err = { ok: false, message: 'Erro ao verificar conexão', details: e.message };
      setStatus(err);
      if (onStatusChange) onStatusChange(err);
    } finally {
      setLoading(false);
    }
  }, [onStatusChange]);

  useEffect(() => {
    check();
  }, [check]);

  return (
    <div className={`flex flex-col gap-2 ${className}`}>
      <div className="flex items-center gap-2">
        <button onClick={check} disabled={loading} className="p-1 rounded hover:bg-slate-100 transition-colors">
          <RefreshCw size={16} className={loading ? 'animate-spin text-slate-400' : 'text-slate-400'} />
        </button>
        {status ? (
          status.ok ? (
            <div className="flex items-center gap-1.5 text-emerald-600 text-sm font-medium">
              <Cloud size={16} />
              <span>Conectado à Nuvem</span>
            </div>
          ) : (
            <div className="flex items-center gap-1.5 text-rose-500 text-sm font-medium">
              <CloudOff size={16} />
              <span>Desconectado</span>
            </div>
          )
        ) : (
          <div className="text-sm text-slate-500 font-medium">Verificando conexão...</div>
        )}
      </div>
      {showDetails && status && (
        <div className="text-xs text-slate-500 ml-8">
          {status.message}
          {status.details && <span className="block mt-0.5 text-slate-400">{status.details}</span>}
        </div>
      )}
    </div>
  );
};
