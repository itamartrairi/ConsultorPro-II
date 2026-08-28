import React from 'react';
import { 
  Cloud, 
  HardDrive, 
  RefreshCw, 
  CheckCircle2, 
  ArrowUpRight, 
  ArrowDownLeft, 
  ShieldCheck, 
  Layers, 
  Building2, 
  FileText, 
  CheckSquare, 
  Sparkles, 
  X,
  AlertCircle,
  Clock,
  Database
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Button } from './Button';
import { cn } from '../lib/utils';

export interface SyncSummary {
  timestamp: string;
  totalAnalyzed: number;
  uploadedToCloud: number;
  updatedInCloud: number;
  downloadedFromCloud: number;
  updatedInLocal: number;
  identicalOrMerged: number;
  details: {
    empresas: { uploaded: number; downloaded: number; updated: number };
    diagnosticos: { uploaded: number; downloaded: number; updated: number };
    respostas: { uploaded: number; downloaded: number; updated: number };
    tarefas: { uploaded: number; downloaded: number; updated: number };
    biblioteca: { uploaded: number; downloaded: number; updated: number };
    outros: { uploaded: number; downloaded: number; updated: number };
  };
  logs?: string[];
}

interface SmartSyncModalProps {
  isOpen: boolean;
  onClose: () => void;
  isSyncing: boolean;
  onStartSync: () => Promise<void>;
  lastSummary: SyncSummary | null;
  userEmail?: string | null;
  error?: string | null;
}

export const SmartSyncModal: React.FC<SmartSyncModalProps> = ({
  isOpen,
  onClose,
  isSyncing,
  onStartSync,
  lastSummary,
  userEmail,
  error
}) => {
  if (!isOpen) return null;

  const totalCloudPush = lastSummary 
    ? (lastSummary.uploadedToCloud + lastSummary.updatedInCloud) 
    : 0;
  const totalLocalPull = lastSummary 
    ? (lastSummary.downloadedFromCloud + lastSummary.updatedInLocal) 
    : 0;

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
      <motion.div
        initial={{ opacity: 0, scale: 0.95, y: 15 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 15 }}
        transition={{ duration: 0.2 }}
        className="bg-white rounded-3xl shadow-2xl border border-slate-200/80 w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]"
      >
        {/* Header */}
        <div className="px-6 py-5 bg-gradient-to-r from-emerald-700 via-teal-700 to-cyan-800 text-white flex items-center justify-between relative overflow-hidden">
          <div className="absolute right-0 top-0 translate-x-4 -translate-y-4 opacity-10 pointer-events-none">
            <Cloud size={180} />
          </div>

          <div className="flex items-center gap-3.5 relative z-10">
            <div className="w-11 h-11 rounded-2xl bg-white/15 backdrop-blur-md flex items-center justify-center text-white shadow-inner border border-white/20">
              <RefreshCw size={22} className={cn(isSyncing && "animate-spin")} />
            </div>
            <div>
              <h3 className="font-extrabold text-lg text-white tracking-tight flex items-center gap-2">
                Sincronização Inteligente
                <span className="bg-emerald-400/20 text-emerald-200 border border-emerald-300/30 text-[10px] uppercase font-black px-2 py-0.5 rounded-full">
                  Últimos Dados Registrados
                </span>
              </h3>
              <p className="text-xs text-emerald-100/90 font-medium">
                {userEmail ? `Conectado à conta: ${userEmail}` : 'Mesclagem bidirecional entre dispositivo local e nuvem'}
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            disabled={isSyncing}
            className="w-8 h-8 rounded-full bg-white/10 hover:bg-white/20 text-white flex items-center justify-center transition-colors relative z-10 disabled:opacity-50"
          >
            <X size={16} />
          </button>
        </div>

        {/* Body Content */}
        <div className="p-6 overflow-y-auto space-y-5">
          {/* Rule banner */}
          <div className="p-4 bg-slate-50 border border-slate-200/80 rounded-2xl flex items-start gap-3.5">
            <div className="w-8 h-8 rounded-xl bg-teal-100 text-teal-700 flex items-center justify-center shrink-0 mt-0.5">
              <Sparkles size={18} />
            </div>
            <div className="text-xs space-y-1">
              <p className="font-bold text-slate-800">
                Resolução Automática por Data de Atualização (Last-Write-Wins)
              </p>
              <p className="text-slate-600 leading-relaxed">
                O sistema compara os registros locais e da nuvem. O registro modificado mais recentemente sempre tem prioridade e atualiza a outra ponta, garantindo que você nunca perca suas últimas alterações feitas offline ou em outros computadores.
              </p>
            </div>
          </div>

          {error && (
            <div className="p-4 bg-rose-50 border border-rose-200 rounded-2xl flex items-start gap-3 text-xs text-rose-800">
              <AlertCircle size={18} className="text-rose-600 shrink-0 mt-0.5" />
              <div>
                <p className="font-bold">Aviso de Sincronização</p>
                <p>{error}</p>
              </div>
            </div>
          )}

          {/* Sync Progress / Stats */}
          {lastSummary ? (
            <div className="space-y-4">
              <div className="flex items-center justify-between text-xs text-slate-500 pb-1 border-b border-slate-100">
                <span className="flex items-center gap-1.5 font-bold text-slate-700">
                  <Clock size={13} className="text-slate-400" />
                  Última sincronização: {lastSummary.timestamp}
                </span>
                <span className="bg-emerald-100 text-emerald-800 font-bold px-2 py-0.5 rounded-full text-[11px]">
                  {lastSummary.totalAnalyzed} registros verificados
                </span>
              </div>

              {/* 3 Metric Cards */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div className="p-4 bg-emerald-50/70 border border-emerald-200/80 rounded-2xl flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-emerald-800 uppercase tracking-wider">
                      Enviados p/ Nuvem
                    </span>
                    <ArrowUpRight size={16} className="text-emerald-600" />
                  </div>
                  <div className="flex items-baseline gap-1.5">
                    <span className="text-2xl font-black text-emerald-700">
                      {totalCloudPush}
                    </span>
                    <span className="text-[10px] text-emerald-600 font-bold">
                      novos / atualizados
                    </span>
                  </div>
                </div>

                <div className="p-4 bg-cyan-50/70 border border-cyan-200/80 rounded-2xl flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-cyan-800 uppercase tracking-wider">
                      Baixados p/ Local
                    </span>
                    <ArrowDownLeft size={16} className="text-cyan-600" />
                  </div>
                  <div className="flex items-baseline gap-1.5">
                    <span className="text-2xl font-black text-cyan-700">
                      {totalLocalPull}
                    </span>
                    <span className="text-[10px] text-cyan-600 font-bold">
                      novos / atualizados
                    </span>
                  </div>
                </div>

                <div className="p-4 bg-indigo-50/70 border border-indigo-200/80 rounded-2xl flex flex-col justify-between">
                  <div className="flex items-center justify-between mb-2">
                    <span className="text-[11px] font-bold text-indigo-800 uppercase tracking-wider">
                      Já Atualizados
                    </span>
                    <CheckCircle2 size={16} className="text-indigo-600" />
                  </div>
                  <div className="flex items-baseline gap-1.5">
                    <span className="text-2xl font-black text-indigo-700">
                      {lastSummary.identicalOrMerged}
                    </span>
                    <span className="text-[10px] text-indigo-600 font-bold">
                      em conformidade
                    </span>
                  </div>
                </div>
              </div>

              {/* Detailed Breakdown */}
              <div className="p-4 bg-slate-50 border border-slate-200/70 rounded-2xl">
                <p className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-3">
                  Detalhamento dos Módulos Sincronizados
                </p>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs">
                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <Building2 size={16} className="text-teal-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Empresas</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.empresas.uploaded + lastSummary.details.empresas.updated} nuvem • {lastSummary.details.empresas.downloaded} local
                      </p>
                    </div>
                  </div>

                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <FileText size={16} className="text-emerald-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Diagnósticos</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.diagnosticos.uploaded + lastSummary.details.diagnosticos.updated} nuvem • {lastSummary.details.diagnosticos.downloaded} local
                      </p>
                    </div>
                  </div>

                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <CheckSquare size={16} className="text-indigo-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Respostas</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.respostas.uploaded + lastSummary.details.respostas.updated} nuvem • {lastSummary.details.respostas.downloaded} local
                      </p>
                    </div>
                  </div>

                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <Layers size={16} className="text-amber-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Plano &amp; Tarefas</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.tarefas.uploaded + lastSummary.details.tarefas.updated} nuvem • {lastSummary.details.tarefas.downloaded} local
                      </p>
                    </div>
                  </div>

                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <Database size={16} className="text-cyan-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Biblioteca &amp; Base</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.biblioteca.uploaded + lastSummary.details.biblioteca.updated} nuvem • {lastSummary.details.biblioteca.downloaded} local
                      </p>
                    </div>
                  </div>

                  <div className="p-2.5 bg-white border border-slate-200/60 rounded-xl flex items-center gap-2.5">
                    <ShieldCheck size={16} className="text-purple-600 shrink-0" />
                    <div className="min-w-0">
                      <p className="font-bold text-slate-800 truncate">Agenda / Outros</p>
                      <p className="text-[11px] text-slate-500">
                        {lastSummary.details.outros.uploaded + lastSummary.details.outros.updated} nuvem • {lastSummary.details.outros.downloaded} local
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              {/* Status Logs */}
              {lastSummary.logs && lastSummary.logs.length > 0 && (
                <div className="p-4 bg-slate-50 border border-slate-200/70 rounded-2xl">
                  <p className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-3">
                    Histórico de Operações
                  </p>
                  <div className="bg-white border border-slate-200/80 rounded-xl p-3 h-36 overflow-y-auto font-mono text-[10px] sm:text-[11px] leading-relaxed text-slate-600 space-y-1.5 custom-scrollbar">
                    {lastSummary.logs.map((log, i) => {
                      const isLocal = log.includes("baixado") || log.includes("local");
                      const isIdentical = log.includes("conformidade");
                      return (
                        <div key={i} className="flex gap-2 items-start">
                          <span className={cn(
                            "shrink-0 font-bold",
                            isIdentical ? "text-indigo-400" : (isLocal ? "text-cyan-500" : "text-emerald-500")
                          )}>
                            {isIdentical ? "—" : (isLocal ? "↓" : "↑")}
                          </span>
                          <span>{log}</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>
          ) : (
            <div className="text-center py-6 px-4 bg-slate-50 border border-dashed border-slate-200 rounded-2xl">
              <Cloud size={36} className="mx-auto text-teal-600 mb-2.5 opacity-80" />
              <p className="font-bold text-sm text-slate-800">Pronto para Sincronizar</p>
              <p className="text-xs text-slate-500 max-w-md mx-auto mt-1">
                Clique no botão abaixo para iniciar a varredura completa. Seus dados locais e os dados na nuvem serão mesclados com base nas informações mais recentes de cada registro.
              </p>
            </div>
          )}
        </div>

        {/* Footer actions */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200/80 flex items-center justify-between gap-3">
          <Button
            variant="ghost"
            onClick={onClose}
            disabled={isSyncing}
            className="text-xs font-bold text-slate-600 hover:text-slate-800"
          >
            Fechar
          </Button>

          <Button
            onClick={onStartSync}
            disabled={isSyncing}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-5 py-2.5 flex items-center gap-2 shadow-sm rounded-xl"
          >
            <RefreshCw size={15} className={cn(isSyncing && "animate-spin")} />
            <span>{isSyncing ? "Sincronizando Dados Mais Recentes..." : "Sincronizar com a Nuvem Agora"}</span>
          </Button>
        </div>
      </motion.div>
    </div>
  );
};
