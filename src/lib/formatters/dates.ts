/**
 * Utilitários de data (Firestore, MDA, atividades do cronograma).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

import { format } from 'date-fns';

export function getMdaExpirationStatus(
  dataValidade?: string | null
): { status: 'vencido' | 'proximo' | 'valido'; diasRestantes: number; label: string } | null {
  if (!dataValidade) return null;
  const validade = new Date(dataValidade + 'T00:00:00');
  if (isNaN(validade.getTime())) return null;
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const diasRestantes = Math.ceil((validade.getTime() - hoje.getTime()) / (1000 * 60 * 60 * 24));
  if (diasRestantes < 0) {
    return { status: 'vencido', diasRestantes, label: `Vencido há ${Math.abs(diasRestantes)} dia(s)` };
  }
  if (diasRestantes <= 30) {
    return { status: 'proximo', diasRestantes, label: `Vence em ${diasRestantes} dia(s)` };
  }
  return { status: 'valido', diasRestantes, label: 'Válido' };
}

export function parseLocalDate(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val.toDate === 'function') {
    const d = val.toDate();
    return isNaN(d.getTime()) ? null : d;
  }
  if (val && typeof val === 'object' && val.seconds) {
    const d = new Date(val.seconds * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      const d = new Date(trimmed + 'T12:00:00');
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(trimmed);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

export function formatFirestoreDate(
  date: any,
  formatStr: string = 'dd/MM/yyyy',
  options?: any
): string {
  if (!date) return '...';
  const d = parseLocalDate(date);
  if (!d) return '...';
  return format(d, formatStr, options);
}

export function getActivityDateStr(diagDate: any, index: number): string {
  let base: Date;
  if (!diagDate) {
    base = new Date();
  } else if (typeof diagDate === 'string') {
    base = new Date(diagDate.includes('T') ? diagDate : diagDate + 'T12:00:00');
  } else if (diagDate.toDate && typeof diagDate.toDate === 'function') {
    base = diagDate.toDate();
  } else if (diagDate.seconds) {
    base = new Date(diagDate.seconds * 1000);
  } else if (diagDate instanceof Date) {
    base = diagDate;
  } else {
    base = new Date(diagDate);
  }

  if (isNaN(base.getTime())) {
    base = new Date();
  }

  const result = new Date(base.getTime());
  result.setDate(result.getDate() + index * 7);
  return result.toISOString().split('T')[0];
}
