import React, { useState } from 'react';
import { 
  Download, 
  Copy, 
  Check, 
  FileJson, 
  Database, 
  ShieldCheck, 
  Calendar, 
  Building2, 
  ClipboardList, 
  ListTodo, 
  Award, 
  Users, 
  Clock, 
  CheckCircle2, 
  Loader2,
  ChevronDown,
  ChevronUp,
  X
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Button } from './Button';

export interface BackupExportStats {
  totalEmpresas: number;
  totalDiagnosticos: number;
  totalRespostas: number;
  totalTarefasPlano: number;
  totalAgendaEventos: number;
  totalDisc: number;
  totalMaturidade: number;
  totalPremissas: number;
  totalProblemas: number;
  totalSolucoes: number;
  totalCredenciadas: number;
  dataTamanhoKb: number;
  exportedAt: string;
  userEmail?: string;
  source: 'nuvem_e_local' | 'local_only';
}

interface BackupExportModalProps {
  isOpen: boolean;
  onClose: () => void;
  isExporting: boolean;
  stats: BackupExportStats | null;
  fileName: string;
  backupJsonString: string;
  onDownloadFile: () => void;
}

export const BackupExportModal: React.FC<BackupExportModalProps> = ({
  isOpen,
  onClose,
  isExporting,
  stats,
  fileName,
  backupJsonString,
  onDownloadFile
}) => {
  const [copied, setCopied] = useState(false);
  const [showJsonPreview, setShowJsonPreview] = useState(false);

  const handleCopy = async () => {
    if (!backupJsonString) return;
    try {
      if (navigator?.clipboard?.writeText) {
        await navigator.clipboard.writeText(backupJsonString);
      } else {
        const textArea = document.createElement("textarea");
        textArea.value = backupJsonString;
        textArea.style.position = "fixed";
        textArea.style.left = "-999999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }
      setCopied(true);
      setTimeout(() => setCopied(false), 3000);
    } catch (err) {
      console.error("Erro ao copiar JSON:", err);
    }
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          transition={{ duration: 0.2 }}
          className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-3xl w-full overflow-hidden flex flex-col max-h-[92vh]"
        >
          {/* Header */}
          <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between shrink-0 bg-slate-50/50">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-700 flex items-center justify-center">
                <Database size={20} />
              </div>
              <div>
                <h3 className="font-bold text-lg text-slate-800">Backup Completo do Acesso</h3>
                <p className="text-xs text-slate-500">Exportação segura de todas as empresas, diagnósticos e planos de ação</p>
              </div>
            </div>
            <button 
              onClick={onClose} 
              className="p-2 hover:bg-slate-200/60 active:scale-95 rounded-full text-slate-400 hover:text-slate-600 transition-all cursor-pointer"
            >
              <X size={20} />
            </button>
          </div>

          {/* Body */}
          <div className="p-6 overflow-y-auto flex-1 space-y-6">
            {isExporting ? (
              <div className="py-14 flex flex-col items-center justify-center text-center space-y-4">
                <div className="w-16 h-16 rounded-2xl bg-emerald-100 text-emerald-600 flex items-center justify-center animate-pulse">
                  <Database size={32} className="animate-bounce" />
                </div>
                <div className="space-y-1">
                  <h4 className="text-base font-bold text-slate-800">
                    Compilando e Extraindo Seus Dados...
                  </h4>
                  <p className="text-xs text-slate-500 max-w-md mx-auto">
                    Consultando o banco de dados na Nuvem (Firestore) e o armazenamento local para reunir todas as suas empresas, diagnósticos, respostas e planos de ação.
                  </p>
                </div>
                <div className="flex items-center gap-2 text-xs font-semibold text-emerald-600">
                  <Loader2 size={16} className="animate-spin" />
                  <span>Aguarde um instante...</span>
                </div>
              </div>
            ) : stats ? (
              <>
                {/* Header Success Banner */}
                <div className="p-4 bg-emerald-50/90 border border-emerald-200 rounded-2xl flex items-start gap-3.5">
                  <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                    <CheckCircle2 size={22} />
                  </div>
                  <div className="space-y-1 flex-1">
                    <div className="flex items-center justify-between">
                      <h4 className="text-sm font-black text-emerald-950 uppercase tracking-wide">
                        Backup Gerado com Sucesso!
                      </h4>
                      <span className="text-[11px] font-bold text-emerald-800 bg-emerald-100/80 px-2.5 py-0.5 rounded-full">
                        {stats.dataTamanhoKb > 1024 
                          ? `${(stats.dataTamanhoKb / 1024).toFixed(2)} MB` 
                          : `${stats.dataTamanhoKb} KB`}
                      </span>
                    </div>
                    <p className="text-xs text-emerald-800 leading-relaxed">
                      Todos os registros do seu acesso (<strong>{stats.userEmail || 'Acesso Local'}</strong>) foram compilados em um arquivo <strong>.JSON</strong> seguro, pronto para download ou salvamento em nuvem pessoal.
                    </p>
                  </div>
                </div>

                {/* Quick Action Buttons */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <Button
                    onClick={onDownloadFile}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm py-3 px-4 rounded-xl flex items-center justify-center gap-2 shadow-sm transition-all cursor-pointer"
                  >
                    <Download size={18} />
                    <span>Baixar Arquivo .JSON Agora</span>
                  </Button>

                  <Button
                    variant="outline"
                    onClick={handleCopy}
                    className="border-slate-200 hover:bg-slate-100 text-slate-700 font-bold text-sm py-3 px-4 rounded-xl flex items-center justify-center gap-2 cursor-pointer transition-all"
                  >
                    {copied ? (
                      <>
                        <Check size={18} className="text-emerald-600" />
                        <span className="text-emerald-700">Copiado para a Área de Transferência!</span>
                      </>
                    ) : (
                      <>
                        <Copy size={18} className="text-indigo-600" />
                        <span>Copiar Conteúdo (JSON)</span>
                      </>
                    )}
                  </Button>
                </div>

                {/* Detailed Metrics Grid */}
                <div className="space-y-2">
                  <p className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                    Conteúdo Incluído no Backup:
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5">
                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-emerald-100 text-emerald-700 rounded-lg shrink-0">
                        <Building2 size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalEmpresas}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Empresas</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-indigo-100 text-indigo-700 rounded-lg shrink-0">
                        <ClipboardList size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalDiagnosticos}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Diagnósticos</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-teal-100 text-teal-700 rounded-lg shrink-0">
                        <CheckCircle2 size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalRespostas}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Respostas</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-amber-100 text-amber-700 rounded-lg shrink-0">
                        <ListTodo size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalTarefasPlano}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Tarefas / Planos</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-blue-100 text-blue-700 rounded-lg shrink-0">
                        <Calendar size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalAgendaEventos}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Agenda / Eventos</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-purple-100 text-purple-700 rounded-lg shrink-0">
                        <Award size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalDisc + stats.totalMaturidade}</p>
                        <p className="text-[10px] text-slate-500 font-medium">DISC & Maturidade</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-rose-100 text-rose-700 rounded-lg shrink-0">
                        <Users size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">{stats.totalCredenciadas}</p>
                        <p className="text-[10px] text-slate-500 font-medium">Credenciadas</p>
                      </div>
                    </div>

                    <div className="p-3 bg-slate-50 border border-slate-200/80 rounded-xl flex items-center gap-2.5">
                      <div className="p-2 bg-slate-200 text-slate-700 rounded-lg shrink-0">
                        <Database size={16} />
                      </div>
                      <div>
                        <p className="text-xs font-black text-slate-800">
                          {stats.totalPremissas + stats.totalProblemas + stats.totalSolucoes}
                        </p>
                        <p className="text-[10px] text-slate-500 font-medium">Itens Biblioteca</p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* File Info Bar */}
                <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/80 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-600">
                  <div className="flex items-center gap-2">
                    <FileJson size={16} className="text-emerald-600" />
                    <span className="font-mono text-slate-800 font-semibold">{fileName}</span>
                  </div>
                  <div className="flex items-center gap-3 text-[11px] text-slate-400">
                    <span className="flex items-center gap-1">
                      <Clock size={13} />
                      {new Date(stats.exportedAt).toLocaleString('pt-BR')}
                    </span>
                    <span className="flex items-center gap-1">
                      <ShieldCheck size={13} className="text-emerald-600" />
                      JSON Válido
                    </span>
                  </div>
                </div>

                {/* Collapsible JSON Preview */}
                <div className="border border-slate-200 rounded-xl overflow-hidden">
                  <button
                    type="button"
                    onClick={() => setShowJsonPreview(!showJsonPreview)}
                    className="w-full px-4 py-2.5 bg-slate-100 hover:bg-slate-200/80 text-slate-700 font-bold text-xs flex items-center justify-between transition-colors cursor-pointer"
                  >
                    <span>Visualizar Estrutura do Arquivo JSON (Prévia)</span>
                    {showJsonPreview ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                  </button>
                  {showJsonPreview && (
                    <div className="p-3 bg-slate-900 text-emerald-400 text-xs font-mono max-h-60 overflow-y-auto">
                      <pre className="whitespace-pre-wrap break-all">
                        {backupJsonString.slice(0, 4000)}
                        {backupJsonString.length > 4000 && '\n\n... (conteúdo completo compactado no arquivo .JSON)'}
                      </pre>
                    </div>
                  )}
                </div>
              </>
            ) : (
              <div className="py-8 text-center text-slate-500 text-sm">
                Nenhum dado selecionado para backup.
              </div>
            )}
          </div>

          {/* Footer */}
          <div className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end">
            <Button
              variant="outline"
              onClick={onClose}
              className="border-slate-200 text-slate-700 hover:bg-slate-100 px-5"
            >
              Fechar
            </Button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};
