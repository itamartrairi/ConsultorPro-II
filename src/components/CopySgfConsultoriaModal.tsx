import React, { useState, useMemo } from 'react';
import { Search, Copy, CheckCircle2, FileSpreadsheet, Building2, User, Clock, ShieldCheck, X, AlertCircle } from 'lucide-react';
import { Button } from './Button';
import { Diagnostico, Empresa, DadosConsultoria } from '../App';

interface CopySgfConsultoriaModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentCodigoSgf?: string;
  diagnosticos: Diagnostico[];
  empresas: Empresa[];
  onSelectDadosConsultoria: (dados: DadosConsultoria) => void;
}

export function CopySgfConsultoriaModal({
  isOpen,
  onClose,
  currentCodigoSgf = '',
  diagnosticos,
  empresas,
  onSelectDadosConsultoria
}: CopySgfConsultoriaModalProps) {
  const [searchTerm, setSearchTerm] = useState(currentCodigoSgf);
  const [selectedItem, setSelectedItem] = useState<{ diag: Diagnostico; dados: DadosConsultoria; empresaNome: string } | null>(null);

  // Collect all unique diagnosticos with valid dadosConsultoria
  const consultoriasList = useMemo(() => {
    const list: { diag: Diagnostico; dados: DadosConsultoria; empresaNome: string }[] = [];
    const seen = new Set<string>();

    diagnosticos.forEach(d => {
      if (d.dadosConsultoria && (d.dadosConsultoria.codigoSgf || d.dadosConsultoria.razaoSocial || d.dadosConsultoria.consultor)) {
        const emp = empresas.find(e => e.id === d.empresaId || e.nome === d.nomeEmpresa);
        const empName = emp?.nome || d.nomeEmpresa || d.nomeProjeto || 'Cliente Sem Nome';
        const key = `${d.dadosConsultoria.codigoSgf || ''}_${d.dadosConsultoria.razaoSocial || ''}_${d.dadosConsultoria.consultor || ''}_${empName}`;
        
        if (!seen.has(key)) {
          seen.add(key);
          list.push({
            diag: d,
            dados: d.dadosConsultoria,
            empresaNome: empName
          });
        }
      }
    });

    return list;
  }, [diagnosticos, empresas]);

  // Filter based on search term
  const filteredList = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) return consultoriasList;

    return consultoriasList.filter(item => {
      const sgf = (item.dados.codigoSgf || '').toLowerCase();
      const emp = item.empresaNome.toLowerCase();
      const cred = (item.dados.razaoSocial || '').toLowerCase();
      const consultor = (item.dados.consultor || '').toLowerCase();
      const area = (item.dados.areaConsultoria || '').toLowerCase();
      const tec = (item.dados.tecnicoSebrae || '').toLowerCase();

      return sgf.includes(term) || emp.includes(term) || cred.includes(term) || consultor.includes(term) || area.includes(term) || tec.includes(term);
    });
  }, [consultoriasList, searchTerm]);

  // Exact SGF matches
  const exactSgfMatches = useMemo(() => {
    if (!currentCodigoSgf.trim()) return [];
    const cleanCurrent = currentCodigoSgf.trim().toLowerCase();
    return consultoriasList.filter(item => (item.dados.codigoSgf || '').trim().toLowerCase() === cleanCurrent);
  }, [consultoriasList, currentCodigoSgf]);

  if (!isOpen) return null;

  const handleConfirmCopy = (dados: DadosConsultoria) => {
    onSelectDadosConsultoria(dados);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-100 w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="bg-gradient-to-r from-blue-700 via-indigo-700 to-slate-800 p-6 text-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-xl bg-white/15 flex items-center justify-center backdrop-blur-xs">
              <FileSpreadsheet className="w-6 h-6 text-white" />
            </div>
            <div>
              <h3 className="font-bold text-xl leading-tight">Copiar Dados da Consultoria pelo Código SGF</h3>
              <p className="text-blue-100 text-xs mt-0.5">
                Importe credenciada, consultor, carga horária, técnico Sebrae e escopo contratual
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-white/80 hover:text-white p-2 rounded-lg hover:bg-white/10 transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        {/* Search & Filter Bar */}
        <div className="p-5 bg-slate-50 border-b border-slate-200/80 shrink-0 space-y-3">
          <div className="relative">
            <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Buscar por Código SGF (ex: SGF-2026-0012), Cliente, Credenciada ou Consultor..."
              className="w-full pl-10 pr-4 py-2.5 bg-white border border-slate-200 rounded-xl text-sm font-medium focus:ring-2 focus:ring-blue-500 outline-none shadow-2xs"
            />
            {searchTerm && (
              <button
                onClick={() => setSearchTerm('')}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600 text-xs font-bold px-1.5 py-0.5 rounded bg-slate-100"
              >
                Limpar
              </button>
            )}
          </div>

          {currentCodigoSgf && exactSgfMatches.length > 0 && (
            <div className="flex items-center gap-2 p-2.5 bg-emerald-50 border border-emerald-200 rounded-xl text-xs text-emerald-800">
              <CheckCircle2 size={16} className="text-emerald-600 shrink-0" />
              <span>
                Encontrado(s) <strong>{exactSgfMatches.length}</strong> registro(s) com o mesmo Código SGF digitado (<strong>{currentCodigoSgf}</strong>).
              </span>
            </div>
          )}
        </div>

        {/* List of Available Consultorias */}
        <div className="p-6 overflow-y-auto flex-1 space-y-4">
          {filteredList.length === 0 ? (
            <div className="py-12 text-center">
              <AlertCircle size={40} className="mx-auto text-slate-300 mb-3" />
              <h4 className="text-base font-bold text-slate-700">Nenhuma consultoria encontrada</h4>
              <p className="text-xs text-slate-500 max-w-md mx-auto mt-1">
                Não encontramos registros com o termo buscado. Certifique-se de que os dados da consultoria foram preenchidos em diagnósticos anteriores.
              </p>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              {filteredList.map((item, idx) => {
                const isExactSgf = currentCodigoSgf && item.dados.codigoSgf && item.dados.codigoSgf.trim().toLowerCase() === currentCodigoSgf.trim().toLowerCase();
                const isSelected = selectedItem?.diag.id === item.diag.id;

                return (
                  <div
                    key={`${item.diag.id}-${idx}`}
                    onClick={() => setSelectedItem(item)}
                    className={`p-4 rounded-xl border transition-all cursor-pointer relative flex flex-col justify-between ${
                      isSelected
                        ? 'border-blue-500 bg-blue-50/50 shadow-sm ring-2 ring-blue-500/20'
                        : isExactSgf
                        ? 'border-emerald-300 bg-emerald-50/30 hover:border-emerald-400'
                        : 'border-slate-200 bg-white hover:border-slate-300 hover:shadow-2xs'
                    }`}
                  >
                    <div>
                      <div className="flex items-center justify-between gap-2 mb-2">
                        <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-bold font-mono bg-blue-100 text-blue-800 border border-blue-200">
                          {item.dados.codigoSgf || 'SGF Não informado'}
                        </span>
                        {isExactSgf && (
                          <span className="px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider bg-emerald-100 text-emerald-800 border border-emerald-200">
                            Mesmo Código SGF
                          </span>
                        )}
                      </div>

                      <div className="space-y-1.5 text-xs">
                        <div className="flex items-center gap-1.5 text-slate-800 font-bold text-sm">
                          <Building2 size={15} className="text-slate-400 shrink-0" />
                          <span className="truncate">{item.empresaNome}</span>
                        </div>

                        {item.dados.razaoSocial && (
                          <div className="text-slate-600 flex items-center gap-1.5">
                            <span className="font-semibold text-slate-500">Credenciada:</span>
                            <span className="truncate">{item.dados.razaoSocial}</span>
                          </div>
                        )}

                        {item.dados.consultor && (
                          <div className="text-slate-600 flex items-center gap-1.5">
                            <User size={13} className="text-slate-400 shrink-0" />
                            <span>{item.dados.consultor}</span>
                          </div>
                        )}

                        <div className="grid grid-cols-2 gap-2 pt-2 border-t border-slate-100 mt-2 text-[11px] text-slate-500">
                          <div>
                            <span className="font-semibold block text-slate-400 uppercase text-[9px]">Área</span>
                            <span className="font-medium text-slate-700 truncate block">{item.dados.areaConsultoria || 'Geral'}</span>
                          </div>
                          <div>
                            <span className="font-semibold block text-slate-400 uppercase text-[9px]">Carga Horária</span>
                            <span className="font-bold text-emerald-700">{item.dados.cargaHoraria || 'Não def.'}</span>
                          </div>
                        </div>

                        {item.dados.tecnicoSebrae && (
                          <div className="text-[11px] text-slate-500">
                            <span className="font-semibold text-slate-400">Técnico Sebrae:</span> {item.dados.tecnicoSebrae}
                          </div>
                        )}
                      </div>
                    </div>

                    <div className="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between">
                      <Button
                        type="button"
                        size="sm"
                        variant="outline"
                        onClick={(e) => {
                          e.stopPropagation();
                          handleConfirmCopy(item.dados);
                        }}
                        className="w-full bg-white hover:bg-blue-50 text-blue-700 border-blue-200 font-semibold text-xs py-1.5"
                      >
                        <Copy size={13} className="mr-1.5" /> Copiar Estes Dados
                      </Button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="p-4 bg-slate-50 border-t border-slate-200/80 flex items-center justify-between shrink-0">
          <p className="text-xs text-slate-500">
            {filteredList.length} consultoria(s) disponível(is) para importação de dados
          </p>
          <Button
            type="button"
            variant="outline"
            onClick={onClose}
            className="text-slate-600 border-slate-200 font-medium"
          >
            Fechar
          </Button>
        </div>
      </div>
    </div>
  );
}
