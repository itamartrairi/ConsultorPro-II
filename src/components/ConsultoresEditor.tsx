import React, { useState } from 'react';
import { Plus, Trash2, UserCheck, UserX, Users } from 'lucide-react';
import type { ConsultorCredenciada } from '../types/domain';
import { isValidEmail, normalizeEmail, novoConsultorId } from '../lib/consultorAccess';
import { formatCPF, isValidCPF, cleanDigits } from '../lib/formatters/br';

/**
 * Lista de consultores de uma credenciada (dentro do modal de edição).
 * Quem tem e-mail Google informado pode entrar no app e atender pela credenciada.
 */
export function ConsultoresEditor({
  value = [],
  onChange,
  readOnly = false,
}: {
  value?: ConsultorCredenciada[];
  onChange: (lista: ConsultorCredenciada[]) => void;
  readOnly?: boolean;
}) {
  const [nome, setNome] = useState('');
  const [email, setEmail] = useState('');
  const [cpf, setCpf] = useState('');
  const [erro, setErro] = useState('');

  const adicionar = () => {
    setErro('');
    if (!nome.trim()) return setErro('Informe o nome do consultor.');
    if (email.trim() && !isValidEmail(email)) return setErro('E-mail inválido.');
    if (email.trim() && value.some(c => normalizeEmail(c.email) === normalizeEmail(email))) {
      return setErro('Este e-mail já está na lista.');
    }
    if (cpf.trim() && cleanDigits(cpf).length > 0 && !isValidCPF(cpf)) return setErro('CPF inválido.');
    onChange([
      ...value,
      {
        id: novoConsultorId(),
        nome: nome.trim(),
        email: normalizeEmail(email) || undefined,
        cpf: cpf.trim() ? formatCPF(cpf) : undefined,
        ativo: true,
      },
    ]);
    setNome('');
    setEmail('');
    setCpf('');
  };

  const atualizar = (id: string, patch: Partial<ConsultorCredenciada>) =>
    onChange(value.map(c => (c.id === id ? { ...c, ...patch } : c)));

  return (
    <div className="md:col-span-2 border border-slate-200 rounded-xl p-4 bg-slate-50/50 space-y-3">
      <div className="flex items-center gap-2">
        <Users size={16} className="text-emerald-600" />
        <h4 className="text-sm font-bold text-slate-700">Consultores da credenciada</h4>
      </div>
      <p className="text-xs text-slate-500">
        Os consultores aparecem para seleção ao atender um cliente. Quem tiver e-mail Google informado
        pode entrar no app com esse e-mail e trabalhar nos clientes desta credenciada.
      </p>

      {value.length === 0 && (
        <p className="text-xs text-slate-400 italic">Nenhum consultor adicional cadastrado.</p>
      )}

      <ul className="space-y-2">
        {value.map(c => (
          <li
            key={c.id}
            className={`flex flex-wrap items-center gap-2 justify-between bg-white border rounded-lg px-3 py-2 ${
              c.ativo === false ? 'border-slate-200 opacity-60' : 'border-emerald-100'
            }`}
          >
            <div className="min-w-0">
              <p className="text-sm font-semibold text-slate-800 truncate">{c.nome}</p>
              <p className="text-[11px] text-slate-500 truncate">
                {c.email ? `Login Google: ${c.email}` : 'Sem login (apenas seleção)'}
                {c.cpf ? ` · CPF ${c.cpf}` : ''}
              </p>
            </div>
            {!readOnly && (
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => atualizar(c.id, { ativo: c.ativo === false })}
                  className="text-[11px] font-bold px-2 py-1 rounded-md border border-slate-200 hover:bg-slate-50 flex items-center gap-1"
                  title={c.ativo === false ? 'Reativar consultor' : 'Desativar consultor (bloqueia o login)'}
                >
                  {c.ativo === false ? <UserCheck size={12} /> : <UserX size={12} />}
                  {c.ativo === false ? 'Reativar' : 'Desativar'}
                </button>
                <button
                  type="button"
                  onClick={() => onChange(value.filter(x => x.id !== c.id))}
                  className="p-1.5 rounded-md text-rose-500 hover:bg-rose-50"
                  title="Remover consultor"
                >
                  <Trash2 size={14} />
                </button>
              </div>
            )}
          </li>
        ))}
      </ul>

      {!readOnly && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-2 pt-1">
          <input
            type="text"
            placeholder="Nome do consultor"
            className="px-3 py-2 rounded-lg border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
            value={nome}
            onChange={e => setNome(e.target.value)}
          />
          <input
            type="email"
            placeholder="E-mail Google (login)"
            className="px-3 py-2 rounded-lg border border-slate-200 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
            value={email}
            onChange={e => setEmail(e.target.value)}
          />
          <input
            type="text"
            placeholder="CPF (opcional)"
            maxLength={14}
            className="px-3 py-2 rounded-lg border border-slate-200 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
            value={cpf}
            onChange={e => setCpf(formatCPF(e.target.value))}
          />
          <div className="md:col-span-3 flex items-center justify-between gap-2">
            <span className="text-xs text-rose-500 font-medium">{erro}</span>
            <button
              type="button"
              onClick={adicionar}
              className="text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 px-3 py-2 rounded-lg flex items-center gap-1"
            >
              <Plus size={14} /> Adicionar consultor
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
