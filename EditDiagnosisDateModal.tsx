import React, { useState, useEffect } from 'react';
import { Calendar, CheckCircle2, Clock, X, AlertCircle } from 'lucide-react';
import { Button } from './Button';
import { format } from 'date-fns';
import { ptBR } from 'date-fns/locale';

interface EditDiagnosisDateModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentDate: any;
  empresaNome?: string;
  projectName?: string;
  onSaveDate: (newDateStr: string) => Promise<void> | void;
}

export function EditDiagnosisDateModal({
  isOpen,
  onClose,
  currentDate,
  empresaNome,
  projectName,
  onSaveDate
}: EditDiagnosisDateModalProps) {
  const [selectedDate, setSelectedDate] = useState<string>('');
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (isOpen) {
      let dateStr = '';
      if (currentDate) {
        if (typeof currentDate === 'string') {
          dateStr = currentDate.split('T')[0];
        } else if (currentDate.seconds) {
          dateStr = new Date(currentDate.seconds * 1000).toISOString().split('T')[0];
        } else if (currentDate instanceof Date && !isNaN(currentDate.getTime())) {
          dateStr = currentDate.toISOString().split('T')[0];
        } else if (currentDate.toDate && typeof currentDate.toDate === 'function') {
          dateStr = currentDate.toDate().toISOString().split('T')[0];
        }
      }
      if (!dateStr || dateStr.length !== 10) {
        dateStr = new Date().toISOString().split('T')[0];
      }
      setSelectedDate(dateStr);
    }
  }, [isOpen, currentDate]);

  if (!isOpen) return null;

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedDate) {
      alert("Por favor, selecione uma data válida.");
      return;
    }
    setIsSaving(true);
    try {
      await onSaveDate(selectedDate);
      onClose();
    } catch (err) {
      console.error("Erro ao salvar data:", err);
    } finally {
      setIsSaving(false);
    }
  };

  const setShortcutDate = (daysAgo: number) => {
    const d = new Date();
    d.setDate(d.getDate() - daysAgo);
    setSelectedDate(d.toISOString().split('T')[0]);
  };

  const getFormattedPreview = () => {
    if (!selectedDate) return '';
    try {
      const parts = selectedDate.split('-');
      if (parts.length === 3) {
        const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
        return format(d, "EEEE, dd 'de' MMMM 'de' yyyy", { locale: ptBR });
      }
    } catch {}
    return selectedDate;
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-100 w-full max-w-md overflow-hidden transform transition-all">
        {/* Header */}
        <div className="bg-gradient-to-r from-emerald-600 to-teal-700 p-5 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-white/15 flex items-center justify-center backdrop-blur-xs">
              <Calendar className="w-5 h-5 text-white" />
            </div>
            <div>
              <h3 className="font-bold text-lg leading-tight">Alterar Data do Diagnóstico</h3>
              <p className="text-emerald-100 text-xs mt-0.5">
                {empresaNome ? `${empresaNome}` : 'Atualizar registro temporal'}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-white/80 hover:text-white p-1.5 rounded-lg hover:bg-white/10 transition-colors"
          >
            <X size={18} />
          </button>
        </div>

        {/* Content */}
        <form onSubmit={handleSave} className="p-6 space-y-5">
          {projectName && (
            <div className="p-3 bg-slate-50 border border-slate-200/70 rounded-xl text-xs text-slate-700">
              <span className="font-semibold text-slate-500 block uppercase text-[10px] mb-0.5">Projeto / Diagnóstico</span>
              <span className="font-medium text-slate-800">{projectName}</span>
            </div>
          )}

          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2">
              Nova Data do Diagnóstico
            </label>
            <div className="relative">
              <input
                type="date"
                required
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-slate-800 font-semibold focus:ring-2 focus:ring-emerald-500 focus:bg-white outline-none transition-all"
              />
            </div>
            {selectedDate && (
              <p className="text-xs text-emerald-700 font-medium capitalize mt-2 flex items-center gap-1.5">
                <CheckCircle2 size={13} className="text-emerald-600 shrink-0" />
                <span>{getFormattedPreview()}</span>
              </p>
            )}
          </div>

          {/* Atalhos Rápidos */}
          <div>
            <span className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1.5">
              Atalhos Rápidos
            </span>
            <div className="flex flex-wrap gap-1.5">
              <button
                type="button"
                onClick={() => setShortcutDate(0)}
                className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-600 transition-colors"
              >
                Hoje
              </button>
              <button
                type="button"
                onClick={() => setShortcutDate(1)}
                className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-600 transition-colors"
              >
                Ontem
              </button>
              <button
                type="button"
                onClick={() => setShortcutDate(7)}
                className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-600 transition-colors"
              >
                Há 1 semana
              </button>
              <button
                type="button"
                onClick={() => setShortcutDate(30)}
                className="px-2.5 py-1 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-emerald-50 hover:text-emerald-700 text-slate-600 transition-colors"
              >
                Há 1 mês
              </button>
            </div>
          </div>

          {/* Footer Actions */}
          <div className="pt-3 border-t border-slate-100 flex items-center justify-end gap-2.5">
            <Button
              type="button"
              variant="outline"
              onClick={onClose}
              disabled={isSaving}
              className="text-slate-600 border-slate-200"
            >
              Cancelar
            </Button>
            <Button
              type="submit"
              disabled={isSaving}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold"
            >
              {isSaving ? "Salvando..." : "Salvar Data"}
            </Button>
          </div>
        </form>
      </div>
    </div>
  );
}
