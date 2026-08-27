import React, { useState, useMemo } from 'react';
import { motion } from 'motion/react';
import { Copy, CheckCircle2, Building2, Calendar, Target, HelpCircle, Layers, ArrowRight, X, AlertCircle, Sparkles, Filter } from 'lucide-react';
import { Button } from './Button';
import { Empresa, Diagnostico, Resposta } from '../App';

interface ReplicateDiagnosticoModalProps {
  isOpen: boolean;
  onClose: () => void;
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas?: Resposta[];
  defaultSourceDiagId?: string;
  defaultTargetEmpresaId?: string;
  onReplicate: (
    sourceDiagId: string,
    targetEmpresaId: string,
    options: {
      includeAnswers: boolean;
      includeConsultoria: boolean;
      customName?: string;
      customDate?: string;
    }
  ) => Promise<any>;
}

export const ReplicateDiagnosticoModal: React.FC<ReplicateDiagnosticoModalProps> = ({
  isOpen,
  onClose,
  empresas,
  diagnosticos,
  defaultSourceDiagId,
  defaultTargetEmpresaId,
  onReplicate
}) => {
  if (!isOpen) return null;

  const [selectedSourceDiagId, setSelectedSourceDiagId] = useState<string>(() => {
    return defaultSourceDiagId || (diagnosticos.length > 0 ? diagnosticos[0].id : '');
  });

  const [selectedTargetEmpresaId, setSelectedTargetEmpresaId] = useState<string>(() => {
    if (defaultTargetEmpresaId) return defaultTargetEmpresaId;
    const firstOther = empresas.find(e => {
      const src = diagnosticos.find(d => d.id === (defaultSourceDiagId || diagnosticos[0]?.id));
      return e.id !== src?.empresaId;
    });
    return firstOther ? firstOther.id : (empresas[0]?.id || '');
  });

  const [filterSegment, setFilterSegment] = useState<string>('todos');
  const [includeAnswers, setIncludeAnswers] = useState<boolean>(true);
  const [includeConsultoria, setIncludeConsultoria] = useState<boolean>(false);
  const [customDate, setCustomDate] = useState<string>(() => new Date().toISOString().split('T')[0]);
  const [customName, setCustomName] = useState<string>('');
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);

  const sourceDiag = useMemo(() => {
    return diagnosticos.find(d => d.id === selectedSourceDiagId);
  }, [diagnosticos, selectedSourceDiagId]);

  const sourceEmpresa = useMemo(() => {
    if (!sourceDiag) return null;
    return empresas.find(e => e.id === sourceDiag.empresaId);
  }, [empresas, sourceDiag]);

  const targetEmpresa = useMemo(() => {
    return empresas.find(e => e.id === selectedTargetEmpresaId);
  }, [empresas, selectedTargetEmpresaId]);

  // Available business segments from existing diagnostics
  const availableSegments = useMemo(() => {
    const set = new Set<string>();
    diagnosticos.forEach(d => {
      if (d.tipoEmpresa) set.add(d.tipoEmpresa);
    });
    return Array.from(set).filter(Boolean);
  }, [diagnosticos]);

  // Filter diagnostics by segment if requested
  const filteredDiagnostics = useMemo(() => {
    if (filterSegment === 'todos') return diagnosticos;
    return diagnosticos.filter(d => (d.tipoEmpresa || 'Geral').toLowerCase() === filterSegment.toLowerCase());
  }, [diagnosticos, filterSegment]);

  // Auto-fill suggested name
  const suggestedName = useMemo(() => {
    if (!targetEmpresa) return '';
    const segment = sourceDiag?.tipoEmpresa || targetEmpresa.tipoEmpresa || 'Geral';
    return `${targetEmpresa.nome} - Diagnóstico Replicado (${segment})`;
  }, [targetEmpresa, sourceDiag]);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedSourceDiagId || !selectedTargetEmpresaId) {
      alert("Por favor, selecione o diagnóstico de origem e a empresa de destino.");
      return;
    }
    setIsSubmitting(true);
    try {
      await onReplicate(selectedSourceDiagId, selectedTargetEmpresaId, {
        includeAnswers,
        includeConsultoria,
        customName: customName.trim() || suggestedName,
        customDate
      });
      onClose();
    } catch (err: any) {
      console.error("Erro ao replicar diagnóstico:", err);
      alert("Erro ao replicar diagnóstico: " + (err?.message || "Tente novamente."));
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm overflow-y-auto">
      <motion.div
        initial={{ opacity: 0, scale: 0.95, y: 15 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 15 }}
        className="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-2xl w-full overflow-hidden my-8"
      >
        {/* Header */}
        <div className="px-6 py-5 bg-gradient-to-r from-emerald-600 to-teal-700 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-white/10 rounded-2xl backdrop-blur-md">
              <Copy size={22} className="text-emerald-100" />
            </div>
            <div>
              <h3 className="text-lg font-bold">Replicar Diagnóstico por Atividade</h3>
              <p className="text-xs text-emerald-100/90">
                Aproveite diagnósticos e premissas de empresas do mesmo segmento econômico
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 text-white/80 hover:text-white hover:bg-white/10 rounded-full transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-6">
          {/* Step 1: Source Diagnostic */}
          <div className="space-y-3">
            <div className="flex items-center justify-between">
              <label className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                <span className="flex items-center justify-center w-4 h-4 rounded-full bg-emerald-100 text-emerald-700 text-[10px]">1</span>
                Diagnóstico de Origem (Modelo / Base)
              </label>

              {availableSegments.length > 1 && (
                <div className="flex items-center gap-1.5 text-xs text-slate-500">
                  <Filter size={12} />
                  <span>Filtrar Atividade:</span>
                  <select
                    value={filterSegment}
                    onChange={(e) => setFilterSegment(e.target.value)}
                    className="text-xs py-0.5 px-2 bg-slate-50 border border-slate-200 rounded-lg text-slate-700 focus:outline-none"
                  >
                    <option value="todos">Todas as Atividades</option>
                    {availableSegments.map(seg => (
                      <option key={seg} value={seg}>{seg}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 gap-2 max-h-48 overflow-y-auto p-1 border border-slate-100 rounded-2xl bg-slate-50/50">
              {filteredDiagnostics.map(diag => {
                const emp = empresas.find(e => e.id === diag.empresaId);
                const isSelected = diag.id === selectedSourceDiagId;
                const diagDate = diag.dataDiagnostico?.seconds
                  ? new Date(diag.dataDiagnostico.seconds * 1000).toLocaleDateString('pt-BR')
                  : diag.dataDiagnostico ? new Date(diag.dataDiagnostico).toLocaleDateString('pt-BR') : '';

                return (
                  <div
                    key={diag.id}
                    onClick={() => setSelectedSourceDiagId(diag.id)}
                    className={`p-3 rounded-xl cursor-pointer border transition-all flex items-center justify-between ${
                      isSelected
                        ? 'bg-emerald-50/80 border-emerald-300 ring-2 ring-emerald-500/20 shadow-sm'
                        : 'bg-white border-slate-200/80 hover:border-emerald-200 hover:bg-slate-50/80'
                    }`}
                  >
                    <div className="flex items-center gap-3">
                      <div className={`p-2 rounded-xl ${isSelected ? 'bg-emerald-600 text-white' : 'bg-slate-100 text-slate-500'}`}>
                        <Building2 size={16} />
                      </div>
                      <div>
                        <div className="text-sm font-bold text-slate-800 flex items-center gap-2">
                          {emp?.nome || diag.nomeProjeto || 'Empresa'}
                          <span className="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase bg-slate-100 text-slate-600 border border-slate-200">
                            {diag.tipoEmpresa || emp?.tipoEmpresa || 'Geral'}
                          </span>
                        </div>
                        <div className="text-xs text-slate-500 flex items-center gap-3 mt-0.5">
                          <span>Data: {diagDate || 'Recente'}</span>
                          {diag.areasDiagnostico && diag.areasDiagnostico.length > 0 && (
                            <span>{diag.areasDiagnostico.length} áreas</span>
                          )}
                        </div>
                      </div>
                    </div>

                    {isSelected && (
                      <div className="text-emerald-600 font-bold text-xs flex items-center gap-1">
                        <CheckCircle2 size={16} />
                        <span>Selecionado</span>
                      </div>
                    )}
                  </div>
                );
              })}

              {filteredDiagnostics.length === 0 && (
                <div className="text-center py-6 text-slate-400 text-xs italic">
                  Nenhum diagnóstico encontrado para a atividade selecionada.
                </div>
              )}
            </div>
          </div>

          {/* Step 2: Target Company */}
          <div className="space-y-3">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
              <span className="flex items-center justify-center w-4 h-4 rounded-full bg-emerald-100 text-emerald-700 text-[10px]">2</span>
              Cliente de Destino (Receberá a Cópia Replicada)
            </label>

            <select
              value={selectedTargetEmpresaId}
              onChange={(e) => setSelectedTargetEmpresaId(e.target.value)}
              className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
            >
              <option value="">Selecione a Empresa de Destino...</option>
              {empresas.map(emp => {
                const isSameSegment = sourceDiag?.tipoEmpresa && emp.tipoEmpresa?.toLowerCase() === sourceDiag.tipoEmpresa.toLowerCase();
                return (
                  <option key={emp.id} value={emp.id}>
                    {emp.nome} ({emp.tipoEmpresa || 'Geral'}) {isSameSegment ? '★ Mesma Atividade' : ''}
                  </option>
                );
              })}
            </select>
          </div>

          {/* Step 3: Options & Custom Name */}
          <div className="space-y-4 pt-2 border-t border-slate-100">
            <label className="text-xs font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
              <span className="flex items-center justify-center w-4 h-4 rounded-full bg-emerald-100 text-emerald-700 text-[10px]">3</span>
              Opções de Replicação
            </label>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <label className="flex items-start gap-3 p-3 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={includeAnswers}
                  onChange={(e) => setIncludeAnswers(e.target.checked)}
                  className="mt-1 rounded text-emerald-600 focus:ring-emerald-500 h-4 w-4"
                />
                <div>
                  <div className="text-xs font-bold text-slate-800">Copiar Respostas e Notas</div>
                  <div className="text-[11px] text-slate-500">
                    Importa respostas ("Sim", "Parcial", "Não") e observações já digitadas.
                  </div>
                </div>
              </label>

              <label className="flex items-start gap-3 p-3 rounded-xl border border-slate-200 bg-slate-50/50 hover:bg-slate-50 cursor-pointer">
                <input
                  type="checkbox"
                  checked={includeConsultoria}
                  onChange={(e) => setIncludeConsultoria(e.target.checked)}
                  className="mt-1 rounded text-emerald-600 focus:ring-emerald-500 h-4 w-4"
                />
                <div>
                  <div className="text-xs font-bold text-slate-800">Copiar Dados da Consultoria</div>
                  <div className="text-[11px] text-slate-500">
                    Inclui objetivos, cronograma sugerido e observações gerais.
                  </div>
                </div>
              </label>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
              <div className="sm:col-span-2">
                <label className="block text-xs font-medium text-slate-600 mb-1">Nome do Novo Diagnóstico / Projeto:</label>
                <input
                  type="text"
                  value={customName}
                  onChange={(e) => setCustomName(e.target.value)}
                  placeholder={suggestedName || "Ex: Empresa - Diagnóstico Replicado"}
                  className="w-full p-2.5 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-600 mb-1">Data do Diagnóstico:</label>
                <input
                  type="date"
                  value={customDate}
                  onChange={(e) => setCustomDate(e.target.value)}
                  className="w-full p-2.5 bg-white border border-slate-200 rounded-xl text-xs font-medium text-slate-800 focus:ring-2 focus:ring-emerald-500 focus:outline-none"
                />
              </div>
            </div>
          </div>

          {/* Info Notice */}
          <div className="p-3.5 bg-emerald-50/80 border border-emerald-200/80 rounded-2xl flex items-start gap-3">
            <Sparkles className="text-emerald-600 shrink-0 mt-0.5" size={16} />
            <p className="text-xs text-emerald-900 leading-relaxed">
              O novo diagnóstico será gerado de forma totalmente independente para <strong>{targetEmpresa?.nome || 'o cliente'}</strong>, permitindo edições, ajustes específicos e geração ágil de relatórios sem alterar os dados originais.
            </p>
          </div>

          {/* Footer Actions */}
          <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
            <Button
              type="button"
              variant="outline"
              onClick={onClose}
              disabled={isSubmitting}
              className="px-5 py-2.5 text-xs font-bold text-slate-600 border-slate-200 hover:bg-slate-50"
            >
              Cancelar
            </Button>

            <Button
              type="submit"
              disabled={isSubmitting || !selectedSourceDiagId || !selectedTargetEmpresaId}
              className="px-6 py-2.5 text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white flex items-center gap-2 shadow-lg shadow-emerald-600/20"
            >
              {isSubmitting ? (
                <>
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  <span>Replicando Diagnóstico...</span>
                </>
              ) : (
                <>
                  <Copy size={16} />
                  <span>Confirmar e Replicar</span>
                </>
              )}
            </Button>
          </div>
        </form>
      </motion.div>
    </div>
  );
};
