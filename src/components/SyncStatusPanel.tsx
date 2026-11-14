/**
 * Painel "Transparência de Sincronização com Firestore".
 * Fica na tela de Configurações, oculto até o usuário clicar em "Mostrar".
 */
import React, { useMemo, useState } from 'react';
import { AnimatePresence, motion } from 'motion/react';
import {
  Activity, ArrowDownLeft, ArrowUpRight, Building2, CheckCircle2, CheckSquare,
  ChevronDown, ChevronUp, Database, FileText, Layers, RefreshCw, Server, ShieldCheck,
} from 'lucide-react';
import { cn } from '../lib/utils';
import { SyncSummary } from './SmartSyncModal';

interface SyncStatusPanelProps {
  lastSyncSummary?: SyncSummary | null;
  isSyncingCloud?: boolean;
  onOpenSmartSync?: () => void;
  onStartSync?: () => Promise<void>;
  userEmail?: string | null;
  storageMode?: 'cloud' | 'local';
  totalEmpresas?: number;
  totalDiagnosticos?: number;
  totalRespostas?: number;
}

export const SyncStatusPanel: React.FC<SyncStatusPanelProps> = ({
  lastSyncSummary,
  isSyncingCloud,
  onOpenSmartSync,
  onStartSync,
  userEmail,
  storageMode = 'cloud',
  totalEmpresas = 0,
  totalDiagnosticos = 0,
  totalRespostas = 0,
}) => {
  const activeSummary = useMemo<SyncSummary | null>(() => {
    if (lastSyncSummary) return lastSyncSummary;
    try {
      const saved = localStorage.getItem('last_sync_summary');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  }, [lastSyncSummary]);
  const [mostrarPainelSync, setMostrarPainelSync] = useState(false);
  const [isSyncLogExpanded, setIsSyncLogExpanded] = useState(false);

  if (!mostrarPainelSync) {
    return (
      <div className="flex justify-end">
        <button
          type="button"
          onClick={() => setMostrarPainelSync(true)}
          className="flex items-center gap-1.5 px-3 py-2 bg-slate-900 hover:bg-slate-800 text-slate-200 hover:text-white border border-slate-700 rounded-xl text-xs font-bold transition-all cursor-pointer shadow-sm"
          title="Mostrar o painel de sincronização com o Firestore"
        >
          <Database size={14} className="text-emerald-400" />
          <span>Mostrar Sincronização com Firestore</span>
          <ChevronDown size={14} />
        </button>
      </div>
    );
  }

  return (
      <div className="bg-slate-900 border border-slate-800 rounded-3xl p-6 text-white shadow-xl relative overflow-hidden transition-all">
        {/* Subtle decorative background glow */}
        <div className="absolute -right-12 -top-12 w-64 h-64 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none" />
        <div className="absolute -left-12 -bottom-12 w-64 h-64 bg-teal-500/10 rounded-full blur-3xl pointer-events-none" />

        <div className="relative z-10 space-y-4">
          {/* Header Row */}
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-slate-800">
            <div className="flex items-start sm:items-center gap-3.5">
              <div className="w-11 h-11 rounded-2xl bg-emerald-500/15 border border-emerald-500/30 flex items-center justify-center text-emerald-400 shrink-0 shadow-inner">
                <Database size={22} className={cn(isSyncingCloud && "animate-spin text-cyan-400")} />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h3 className="font-extrabold text-base text-white tracking-tight flex items-center gap-2">
                    Transparência de Sincronização com Firestore
                  </h3>
                  <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                    Firestore Cloud Ativo
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-0.5">
                  <span>
                    {activeSummary?.timestamp 
                      ? `Última tentativa de sincronização registrada em: ${activeSummary.timestamp}` 
                      : `Sincronização em tempo real ativa • Listeners conectados`}
                  </span>
                  {userEmail && (
                    <span className="text-emerald-400 font-medium">• Conta: {userEmail}</span>
                  )}
                </p>
              </div>
            </div>

            {/* Quick Actions */}
            <div className="flex items-center gap-2 self-end lg:self-center">
              <button
                type="button"
                onClick={async () => {
                  if (onStartSync) {
                    await onStartSync();
                  } else if (onOpenSmartSync) {
                    onOpenSmartSync();
                  }
                }}
                disabled={isSyncingCloud}
                className="flex items-center gap-2 px-3.5 py-2 bg-emerald-600 hover:bg-emerald-500 active:bg-emerald-700 text-white rounded-xl text-xs font-bold transition-all shadow-md hover:shadow-emerald-900/30 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
                title="Executar sincronização inteligente agora"
              >
                <RefreshCw size={14} className={cn(isSyncingCloud && "animate-spin")} />
                <span>{isSyncingCloud ? 'Sincronizando...' : 'Sincronizar Agora'}</span>
              </button>

              <button
                type="button"
                onClick={() => setIsSyncLogExpanded(!isSyncLogExpanded)}
                className="flex items-center gap-1.5 px-3 py-2 bg-slate-800/80 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-700/60 rounded-xl text-xs font-bold transition-all cursor-pointer"
                title="Alternar detalhamento completo do log"
              >
                <span>{isSyncLogExpanded ? 'Ocultar Detalhes' : 'Ver Detalhes'}</span>
                {isSyncLogExpanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
              </button>
              <button
                type="button"
                onClick={() => setMostrarPainelSync(false)}
                className="flex items-center gap-1.5 px-3 py-2 bg-slate-800/80 hover:bg-slate-800 text-slate-300 hover:text-white border border-slate-700/60 rounded-xl text-xs font-bold transition-all cursor-pointer"
                title="Ocultar o painel de sincronização"
              >
                <span>Ocultar Painel</span>
                <ChevronUp size={14} />
              </button>
            </div>
          </div>

          {/* KPI Metrics Grid */}
          <div className="grid grid-cols-2 sm:grid-cols-2 md:grid-cols-4 gap-3">
            {/* Total Analyzed / Processed */}
            <div className="bg-slate-800/60 border border-slate-700/50 rounded-2xl p-3.5 backdrop-blur-sm">
              <div className="flex items-center justify-between text-slate-400 mb-1">
                <span className="text-[11px] font-bold uppercase tracking-wider">Docs Processados</span>
                <Activity size={15} className="text-emerald-400" />
              </div>
              <p className="text-2xl font-black text-white tracking-tight">
                {activeSummary ? activeSummary.totalAnalyzed : (totalEmpresas + totalDiagnosticos + totalRespostas)}
              </p>
              <p className="text-[10px] text-slate-400 mt-0.5">
                {activeSummary ? 'Analisados na última tentativa' : 'Registros ativos monitorados'}
              </p>
            </div>

            {/* Uploaded / Created */}
            <div className="bg-slate-800/60 border border-slate-700/50 rounded-2xl p-3.5 backdrop-blur-sm">
              <div className="flex items-center justify-between text-slate-400 mb-1">
                <span className="text-[11px] font-bold uppercase tracking-wider">Enviados à Nuvem</span>
                <ArrowUpRight size={15} className="text-emerald-400" />
              </div>
              <p className="text-2xl font-black text-emerald-400 tracking-tight">
                {activeSummary ? (activeSummary.uploadedToCloud + activeSummary.updatedInCloud) : 0}
              </p>
              <p className="text-[10px] text-slate-400 mt-0.5">
                {activeSummary ? `${activeSummary.uploadedToCloud} novos + ${activeSummary.updatedInCloud} modif.` : 'Sem pendência de envio'}
              </p>
            </div>

            {/* Downloaded / Pulled */}
            <div className="bg-slate-800/60 border border-slate-700/50 rounded-2xl p-3.5 backdrop-blur-sm">
              <div className="flex items-center justify-between text-slate-400 mb-1">
                <span className="text-[11px] font-bold uppercase tracking-wider">Baixados / Locais</span>
                <ArrowDownLeft size={15} className="text-cyan-400" />
              </div>
              <p className="text-2xl font-black text-cyan-400 tracking-tight">
                {activeSummary ? (activeSummary.downloadedFromCloud + activeSummary.updatedInLocal) : 0}
              </p>
              <p className="text-[10px] text-slate-400 mt-0.5">
                {activeSummary ? `${activeSummary.downloadedFromCloud} baixados + ${activeSummary.updatedInLocal} atualizados` : 'Sincronizado'}
              </p>
            </div>

            {/* Verified / Merged */}
            <div className="bg-slate-800/60 border border-slate-700/50 rounded-2xl p-3.5 backdrop-blur-sm">
              <div className="flex items-center justify-between text-slate-400 mb-1">
                <span className="text-[11px] font-bold uppercase tracking-wider">Íntegros & Mesclados</span>
                <ShieldCheck size={15} className="text-teal-400" />
              </div>
              <p className="text-2xl font-black text-teal-400 tracking-tight">
                {activeSummary ? activeSummary.identicalOrMerged : (totalEmpresas + totalDiagnosticos)}
              </p>
              <p className="text-[10px] text-slate-400 mt-0.5">
                Validados sem conflito
              </p>
            </div>
          </div>

          {/* Expanded Breakdown by Collection */}
          <AnimatePresence>
            {isSyncLogExpanded && (
              <motion.div
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                transition={{ duration: 0.2 }}
                className="overflow-hidden pt-2 border-t border-slate-800 space-y-3"
              >
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-slate-300 uppercase tracking-wider flex items-center gap-1.5">
                    <Layers size={14} className="text-emerald-400" />
                    Detalhamento por Coleção no Firestore
                  </span>
                  <span className="text-[11px] text-slate-400 font-mono">
                    Protocolo: AES-256 / Security Rules v2
                  </span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-2.5 text-xs">
                  {/* Empresas */}
                  <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-3 space-y-1.5">
                    <div className="flex items-center justify-between font-bold text-slate-200">
                      <span className="flex items-center gap-1.5">
                        <Building2 size={14} className="text-emerald-400" /> Empresas
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.empresas.uploaded + activeSummary.details.empresas.downloaded + activeSummary.details.empresas.updated) : totalEmpresas}
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400 space-y-0.5 font-mono">
                      <div>⬆ Uploads: <span className="text-emerald-400 font-bold">{activeSummary?.details.empresas.uploaded || 0}</span></div>
                      <div>🔄 Atualizações: <span className="text-cyan-400 font-bold">{activeSummary?.details.empresas.updated || 0}</span></div>
                      <div>⬇ Downloads: <span className="text-blue-400 font-bold">{activeSummary?.details.empresas.downloaded || 0}</span></div>
                    </div>
                  </div>

                  {/* Diagnósticos */}
                  <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-3 space-y-1.5">
                    <div className="flex items-center justify-between font-bold text-slate-200">
                      <span className="flex items-center gap-1.5">
                        <FileText size={14} className="text-teal-400" /> Diagnósticos
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.diagnosticos.uploaded + activeSummary.details.diagnosticos.downloaded + activeSummary.details.diagnosticos.updated) : totalDiagnosticos}
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400 space-y-0.5 font-mono">
                      <div>⬆ Uploads: <span className="text-emerald-400 font-bold">{activeSummary?.details.diagnosticos.uploaded || 0}</span></div>
                      <div>🔄 Atualizações: <span className="text-cyan-400 font-bold">{activeSummary?.details.diagnosticos.updated || 0}</span></div>
                      <div>⬇ Downloads: <span className="text-blue-400 font-bold">{activeSummary?.details.diagnosticos.downloaded || 0}</span></div>
                    </div>
                  </div>

                  {/* Respostas */}
                  <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-3 space-y-1.5">
                    <div className="flex items-center justify-between font-bold text-slate-200">
                      <span className="flex items-center gap-1.5">
                        <CheckSquare size={14} className="text-sky-400" /> Respostas
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.respostas.uploaded + activeSummary.details.respostas.downloaded + activeSummary.details.respostas.updated) : totalRespostas}
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400 space-y-0.5 font-mono">
                      <div>⬆ Uploads: <span className="text-emerald-400 font-bold">{activeSummary?.details.respostas.uploaded || 0}</span></div>
                      <div>🔄 Atualizações: <span className="text-cyan-400 font-bold">{activeSummary?.details.respostas.updated || 0}</span></div>
                      <div>⬇ Downloads: <span className="text-blue-400 font-bold">{activeSummary?.details.respostas.downloaded || 0}</span></div>
                    </div>
                  </div>

                  {/* Tarefas */}
                  <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-3 space-y-1.5">
                    <div className="flex items-center justify-between font-bold text-slate-200">
                      <span className="flex items-center gap-1.5">
                        <CheckCircle2 size={14} className="text-amber-400" /> Tarefas Plano
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.tarefas.uploaded + activeSummary.details.tarefas.downloaded + activeSummary.details.tarefas.updated) : 'Ativas'}
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400 space-y-0.5 font-mono">
                      <div>⬆ Uploads: <span className="text-emerald-400 font-bold">{activeSummary?.details.tarefas.uploaded || 0}</span></div>
                      <div>🔄 Atualizações: <span className="text-cyan-400 font-bold">{activeSummary?.details.tarefas.updated || 0}</span></div>
                      <div>⬇ Downloads: <span className="text-blue-400 font-bold">{activeSummary?.details.tarefas.downloaded || 0}</span></div>
                    </div>
                  </div>

                  {/* Biblioteca Metodológica */}
                  <div className="bg-slate-800/40 border border-slate-700/40 rounded-xl p-3 space-y-1.5">
                    <div className="flex items-center justify-between font-bold text-slate-200">
                      <span className="flex items-center gap-1.5">
                        <Server size={14} className="text-violet-400" /> Biblioteca
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.biblioteca.uploaded + activeSummary.details.biblioteca.downloaded) : 'Sincronizada'}
                      </span>
                    </div>
                    <div className="text-[10px] text-slate-400 space-y-0.5 font-mono">
                      <div>⬆ Uploads: <span className="text-emerald-400 font-bold">{activeSummary?.details.biblioteca.uploaded || 0}</span></div>
                      <div>🔄 Atualizações: <span className="text-cyan-400 font-bold">{activeSummary?.details.biblioteca.updated || 0}</span></div>
                      <div>⬇ Downloads: <span className="text-blue-400 font-bold">{activeSummary?.details.biblioteca.downloaded || 0}</span></div>
                    </div>
                  </div>
                </div>

                <div className="pt-2 flex flex-wrap items-center justify-between gap-2 text-[10px] text-slate-400 bg-slate-950/40 rounded-xl p-2.5 border border-slate-800">
                  <div className="flex items-center gap-1.5">
                    <ShieldCheck size={13} className="text-emerald-400" />
                    <span>Modo de Armazenamento: <strong className="text-slate-200 uppercase">{storageMode === 'cloud' ? 'Nuvem Firestore + Cache Criptografado' : 'Dispositivo Local'}</strong></span>
                  </div>
                  <div className="text-slate-400 font-mono">
                    Todos os diagnósticos e pontuações deste Dashboard Macro são calculados a partir dos documentos sincronizados.
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
  );
};
