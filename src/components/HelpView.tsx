import React from 'react';
import { HelpCircle, PlayCircle } from 'lucide-react';

// Base do Vite ('./'): funciona no site do Netlify e no app desktop (Electron).
const BASE = import.meta.env.BASE_URL || './';

const ETAPAS = [
  { n: 1, titulo: 'Acesso', texto: 'Login, credenciada, consultores, chave de IA e logotipos.' },
  { n: 2, titulo: 'Biblioteca', texto: 'Tipos de negócio, áreas, problemas, perguntas e soluções.' },
  { n: 3, titulo: 'Cliente', texto: 'Nova Empresa com busca automática pelo CNPJ.' },
  { n: 4, titulo: 'Atendimento', texto: 'Novo Diagnóstico: projeto, consultor, data e áreas.' },
  { n: 5, titulo: 'Diagnóstico', texto: 'Respostas Sim, Parcial ou Não e Dashboard de Resultados.' },
  { n: 6, titulo: 'Consultoria', texto: 'Dados da consultoria e Relatório de Consultoria.' },
  { n: 7, titulo: 'Plano de ação', texto: 'Tarefas, Kanban, agenda, projetos e resultado.' },
  { n: 8, titulo: 'Relatórios', texto: 'Relatório final, PDFs e Relatório SGF & Áreas.' },
];

/** Aba Help: vídeo tutorial do app, do cadastro do cliente à emissão dos relatórios. */
export function HelpView() {
  return (
    <div className="max-w-5xl mx-auto space-y-8 pb-12">
      <div className="flex items-start gap-3">
        <div className="w-12 h-12 rounded-xl bg-emerald-50 border border-emerald-100 flex items-center justify-center text-emerald-600 shrink-0">
          <HelpCircle size={24} />
        </div>
        <div>
          <h2 className="text-2xl font-black text-slate-800 tracking-tight">Help</h2>
          <p className="text-slate-500 text-sm">
            Vídeo tutorial do Consultor Pro: o passo a passo do cadastro do cliente até a emissão dos relatórios.
          </p>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="flex items-center gap-2 px-5 py-3 border-b border-slate-100">
          <PlayCircle size={18} className="text-emerald-600" />
          <span className="text-sm font-bold text-slate-700">Tutorial completo</span>
          <span className="text-xs text-slate-400 ml-auto">6 min 45 s</span>
        </div>
        <video
          className="w-full aspect-video bg-slate-900"
          controls
          preload="metadata"
          playsInline
          poster={`${BASE}videos/tutorial-consultor-pro.jpg`}
        >
          <source src={`${BASE}videos/tutorial-consultor-pro%20(2).mp4`} type="video/mp4" />
          Seu navegador não conseguiu reproduzir o vídeo.
        </video>
      </div>

      <div>
        <h3 className="text-sm font-black uppercase tracking-wider text-slate-500 mb-3">O que o vídeo mostra</h3>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          {ETAPAS.map(e => (
            <div key={e.n} className="bg-white rounded-xl border border-slate-200 p-4">
              <div className="flex items-center gap-2 mb-1">
                <span className="w-6 h-6 rounded-full bg-emerald-600 text-white text-xs font-bold flex items-center justify-center">
                  {e.n}
                </span>
                <span className="text-sm font-bold text-slate-800">{e.titulo}</span>
              </div>
              <p className="text-xs text-slate-500 leading-relaxed">{e.texto}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
