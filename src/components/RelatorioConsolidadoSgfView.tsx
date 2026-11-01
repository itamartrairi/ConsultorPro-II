import React, { useState, useMemo, useEffect } from 'react';
import {
  FileText,
  Filter,
  Search,
  RotateCcw,
  Download,
  Building2,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  BarChart3,
  Layers,
  Sparkles,
  ChevronDown,
  ChevronUp,
  Tag,
  Briefcase,
  RefreshCw,
  Database,
  Clock
} from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell
} from 'recharts';
import type { Empresa, Diagnostico, Resposta, Premissa, Problema } from '../types/domain';
import {
  montarRelatorioConsolidadoSgf,
  type FiltrosRelatorioConsolidado,
  type RelatorioConsolidadoSgfData
} from '../lib/domain/relatorioConsolidadoSgf';
import { gerarPdfConsolidadoSgf } from '../lib/relatorios/pdfConsolidadoSgf';
import { loadAllLocalRespostas } from '../lib/domain/respostas';
import { db } from '../firebase';
import { collection, query, where, getDocs, ownerFilter } from '../lib/firestoreOwned';
import { playClickSound, playSuccessSound } from '../lib/sound';

interface RelatorioConsolidadoSgfViewProps {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas: Resposta[];
  premissas?: Premissa[];
  problemas?: Problema[];
  user?: any;
  customLogo?: string | null;
  customConsultoraLogo?: string | null;
  logoChoice?: 'sebrae' | 'consultora' | 'none';
  setPdfUrl: (url: string) => void;
  onNavigateToDiagnosis?: (diag: Diagnostico) => void;
}

const CORES_DONUT = ['#059669', '#e11d48', '#d97706', '#94a3b8'];

export const RelatorioConsolidadoSgfView: React.FC<RelatorioConsolidadoSgfViewProps> = ({
  empresas: propEmpresas,
  diagnosticos: propDiagnosticos,
  respostas: propRespostas,
  premissas: propPremissas = [],
  problemas: propProblemas = [],
  user,
  customLogo,
  customConsultoraLogo,
  logoChoice = 'sebrae',
  setPdfUrl,
}) => {
  // Estados para dados completos das tabelas do sistema
  const [empresas, setEmpresas] = useState<Empresa[]>(propEmpresas);
  const [diagnosticos, setDiagnosticos] = useState<Diagnostico[]>(propDiagnosticos);
  const [respostas, setRespostas] = useState<Resposta[]>(propRespostas);
  const [premissas, setPremissas] = useState<Premissa[]>(propPremissas);
  const [problemas, setProblemas] = useState<Problema[]>(propProblemas);
  const [carregandoTabelas, setCarregandoTabelas] = useState<boolean>(true);

  // Estado dos filtros
  const [tipoNegocio, setTipoNegocio] = useState<string>('Todos');
  const [codigoSgf, setCodigoSgf] = useState<string>('Todos');
  const [area, setArea] = useState<string>('Todas');
  const [empresaId, setEmpresaId] = useState<string>('Todas');
  const [classificacao, setClassificacao] = useState<'todos' | 'pontos_fortes' | 'pontos_criticos' | 'respondidas'>('respondidas');
  const [buscaTexto, setBuscaTexto] = useState<string>('');

  const [gerandoPdf, setGerandoPdf] = useState(false);
  const [mostrarGraficos, setMostrarGraficos] = useState(true);
  const [gruposAbertos, setGruposAbertos] = useState<Record<string, boolean>>({});

  // Efeito para carregar e conectar todas as tabelas do sistema
  const carregarTabelasDoSistema = async () => {
    setCarregandoTabelas(true);
    try {
      // 1. Carregar Empresas do LocalStorage ou Cloud
      let listaEmpresas: Empresa[] = propEmpresas.length > 0 ? propEmpresas : [];
      if (listaEmpresas.length === 0) {
        try {
          const savedEmp = localStorage.getItem('local_empresas');
          if (savedEmp) {
            const parsed = JSON.parse(savedEmp);
            if (Array.isArray(parsed) && parsed.length > 0) listaEmpresas = parsed;
          }
        } catch {}
      }
      if (listaEmpresas.length === 0 && user) {
        try {
          const snapE = await getDocs(query(collection(db, 'empresas'), ownerFilter()));
          if (!snapE.empty) {
            listaEmpresas = snapE.docs.map(d => ({ id: d.id, ...d.data() } as Empresa));
          }
        } catch {}
      }
      setEmpresas(listaEmpresas);

      // 2. Carregar Diagnósticos do LocalStorage ou Cloud
      let listaDiags: Diagnostico[] = propDiagnosticos.length > 0 ? propDiagnosticos : [];
      if (listaDiags.length === 0) {
        try {
          const savedDiags = localStorage.getItem('local_diagnosticos');
          if (savedDiags) {
            const parsed = JSON.parse(savedDiags);
            if (Array.isArray(parsed) && parsed.length > 0) listaDiags = parsed;
          }
        } catch {}
      }
      if (listaDiags.length === 0 && user) {
        try {
          const snapD = await getDocs(query(collection(db, 'diagnosticos'), ownerFilter()));
          if (!snapD.empty) {
            listaDiags = snapD.docs.map(d => ({ id: d.id, ...d.data() } as Diagnostico));
          }
        } catch {}
      }
      setDiagnosticos(listaDiags);

      // 3. Carregar TODAS as Respostas de todos os diagnósticos
      const mapaRespostas = new Map<string, Resposta>();

      // A. Respostas do cache local
      try {
        loadAllLocalRespostas().forEach(r => {
          if (r?.id) mapaRespostas.set(r.id, r);
        });
      } catch {}

      try {
        const rawLocal = localStorage.getItem('local_respostas');
        if (rawLocal) {
          const parsed = JSON.parse(rawLocal);
          if (Array.isArray(parsed)) {
            parsed.forEach((r: any) => {
              if (r?.id) mapaRespostas.set(r.id, r);
            });
          }
        }
      } catch {}

      try {
        const rawLocalAll = localStorage.getItem('local_all_respostas');
        if (rawLocalAll) {
          const parsed = JSON.parse(rawLocalAll);
          if (Array.isArray(parsed)) {
            parsed.forEach((r: any) => {
              if (r?.id) mapaRespostas.set(r.id, r);
            });
          }
        }
      } catch {}

      // Respostas passadas nas props
      if (Array.isArray(propRespostas) && propRespostas.length > 0) {
        propRespostas.forEach(r => {
          if (r?.id) mapaRespostas.set(r.id, r);
        });
      }

      // B. Respostas na Nuvem (Firestore)
      if (user) {
        try {
          // Busca geral pelo owner
          const snapResps = await getDocs(query(collection(db, 'respostas'), ownerFilter()));
          snapResps.docs.forEach(docSnap => {
            mapaRespostas.set(docSnap.id, { id: docSnap.id, ...docSnap.data() } as Resposta);
          });
        } catch (e) {
          console.warn('[Relatório SGF] Busca geral de respostas por owner falhou; buscando por diagnósticos:', e);
        }

        // Se ainda tiver diagnósticos sem resposta, buscar em blocos pelos IDs
        const diagIds = listaDiags.map(d => d.id).filter(Boolean);
        const chunkSize = 30;
        for (let i = 0; i < diagIds.length; i += chunkSize) {
          const chunk = diagIds.slice(i, i + chunkSize);
          try {
            const qChunk = query(collection(db, 'respostas'), where('diagnosticoId', 'in', chunk));
            const snapChunk = await getDocs(qChunk);
            snapChunk.docs.forEach(docSnap => {
              mapaRespostas.set(docSnap.id, { id: docSnap.id, ...docSnap.data() } as Resposta);
            });
          } catch (err) {
            // quiet chunk fallback
          }
        }
      }

      setRespostas(Array.from(mapaRespostas.values()));

      // 4. Carregar Premissas e Problemas se vazios
      let listaPremissas = propPremissas;
      if (listaPremissas.length === 0) {
        try {
          const savedP = localStorage.getItem('local_premissas');
          if (savedP) listaPremissas = JSON.parse(savedP);
        } catch {}
      }
      setPremissas(listaPremissas);

      let listaProblemas = propProblemas;
      if (listaProblemas.length === 0) {
        try {
          const savedProb = localStorage.getItem('local_problemas');
          if (savedProb) listaProblemas = JSON.parse(savedProb);
        } catch {}
      }
      setProblemas(listaProblemas);

    } catch (err) {
      console.error('Erro ao conectar tabelas do sistema:', err);
    } finally {
      setCarregandoTabelas(false);
    }
  };

  useEffect(() => {
    carregarTabelasDoSistema();
  }, [propEmpresas, propDiagnosticos, user]);

  const filtrosAtivos: FiltrosRelatorioConsolidado = useMemo(() => ({
    tipoNegocio,
    codigoSgf,
    area,
    empresaId,
    classificacao,
    buscaTexto,
  }), [tipoNegocio, codigoSgf, area, empresaId, classificacao, buscaTexto]);

  // Montagem consolidada com dados completos das tabelas
  const dadosConsolidados: RelatorioConsolidadoSgfData = useMemo(() => {
    return montarRelatorioConsolidadoSgf({
      empresas,
      diagnosticos,
      respostas,
      premissas,
      problemas,
      filtros: filtrosAtivos,
    });
  }, [empresas, diagnosticos, respostas, premissas, problemas, filtrosAtivos]);

  // Limpar filtros
  const handleLimparFiltros = () => {
    playClickSound();
    setTipoNegocio('Todos');
    setCodigoSgf('Todos');
    setArea('Todas');
    setEmpresaId('Todas');
    setClassificacao('respondidas');
    setBuscaTexto('');
  };

  // Gerar e abrir visualizador de PDF
  const handleGerarPdf = () => {
    try {
      setGerandoPdf(true);
      playClickSound();

      if (dadosConsolidados.totalItens === 0) {
        alert('Não há itens para gerar o relatório com os filtros atuais.');
        return;
      }

      const logo = logoChoice === 'consultora' ? customConsultoraLogo : logoChoice === 'sebrae' ? customLogo : null;
      const doc = gerarPdfConsolidadoSgf(dadosConsolidados, {
        consultor: user?.email || undefined,
        logo,
      });

      const url = doc.output('bloburl').toString();
      setPdfUrl(url);
      playSuccessSound();
    } catch (e: any) {
      console.error('Erro ao gerar PDF consolidado SGF:', e);
      alert('Não foi possível gerar o PDF consolidado.\n\nDetalhes: ' + (e?.message || e));
    } finally {
      setGerandoPdf(false);
    }
  };

  // Exportar dados filtrados para CSV
  const handleExportarCsv = () => {
    playClickSound();
    if (dadosConsolidados.itensFiltrados.length === 0) {
      alert('Não há dados filtrados para exportar.');
      return;
    }

    const cabecalho = ['Tipo de Negocio', 'Codigo SGF', 'Area', 'Cliente', 'Pergunta', 'Resposta', 'Avaliacao', 'Observacao'];
    const linhas = dadosConsolidados.itensFiltrados.map(i => [
      `"${i.tipoNegocio.replace(/"/g, '""')}"`,
      `"${i.codigoSgf.replace(/"/g, '""')}"`,
      `"${i.area.replace(/"/g, '""')}"`,
      `"${i.cliente.replace(/"/g, '""')}"`,
      `"${i.pergunta.replace(/"/g, '""')}"`,
      `"${(i.resposta || '').replace(/"/g, '""')}"`,
      `"${i.classificacao.replace(/"/g, '""')}"`,
      `"${(i.observacao || '').replace(/"/g, '""')}"`,
    ]);

    const csvContent = '\uFEFF' + [cabecalho.join(';'), ...linhas.map(l => l.join(';'))].join('\r\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `relatorio_consolidado_sgf_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Alternar abertura de acordeão de grupo
  const toggleGrupo = (chave: string) => {
    setGruposAbertos(prev => ({
      ...prev,
      [chave]: prev[chave] === undefined ? false : !prev[chave]
    }));
  };

  // Dados para o Donut
  const dadosGraficoDonut = useMemo(() => [
    { name: 'Pontos Fortes', value: dadosConsolidados.totalPontosFortes },
    { name: 'Pontos Críticos', value: dadosConsolidados.totalPontosCriticos },
    { name: 'Oportunidades', value: dadosConsolidados.totalOportunidades },
    { name: 'Pendentes', value: dadosConsolidados.totalPendentes },
  ].filter(d => d.value > 0), [dadosConsolidados]);

  return (
    <div className="space-y-6 pb-16">
      {/* Cabeçalho da Página */}
      <div className="bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm">
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <span className="bg-emerald-100 text-emerald-800 text-xs font-black px-2.5 py-0.5 rounded-full uppercase tracking-wider">
                Módulo Gerencial
              </span>
              <span className="text-slate-400 text-xs font-medium">Relatório Consolidado SGF</span>
            </div>
            <h1 className="text-2xl md:text-3xl font-extrabold text-slate-900 tracking-tight">
              Diagnósticos por Tipo de Negócio, SGF e Áreas
            </h1>
            <p className="text-slate-500 text-sm mt-1 max-w-3xl">
              Concatenação analítica dos diagnósticos por segmento, contrato SGF e área diagnosticada, com tabela comparativa de clientes, perguntas, respostas, pontos fortes e pontos críticos.
            </p>
          </div>

          {/* Botões de Ação do Topo */}
          <div className="flex flex-wrap items-center gap-2.5 w-full md:w-auto">
            <button
              onClick={carregarTabelasDoSistema}
              disabled={carregandoTabelas}
              className="flex items-center justify-center gap-1.5 px-3 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold text-xs transition-all cursor-pointer disabled:opacity-50"
              title="Recarregar dados das tabelas do sistema"
            >
              <RefreshCw size={14} className={carregandoTabelas ? "animate-spin text-emerald-600" : "text-slate-600"} />
              <span>{carregandoTabelas ? 'Sincronizando...' : 'Atualizar Dados'}</span>
            </button>

            <button
              onClick={handleExportarCsv}
              disabled={dadosConsolidados.totalItens === 0}
              className="flex items-center justify-center gap-2 px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold text-xs transition-all cursor-pointer disabled:opacity-50"
              title="Exportar planilha CSV dos dados filtrados"
            >
              <FileSpreadsheet size={15} className="text-slate-600" />
              <span>Exportar CSV</span>
            </button>

            <button
              onClick={handleGerarPdf}
              disabled={gerandoPdf || dadosConsolidados.totalItens === 0}
              className="flex-1 md:flex-initial flex items-center justify-center gap-2 px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-sm shadow-md shadow-emerald-200 transition-all cursor-pointer disabled:opacity-60"
            >
              <FileText size={16} />
              <span>{gerandoPdf ? 'Gerando PDF...' : 'Gerar Relatório PDF'}</span>
            </button>
          </div>
        </div>

        {/* Status de Conexão com as Tabelas */}
        <div className="mt-4 pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-2 text-slate-600">
            <Database size={14} className="text-emerald-600" />
            <span className="font-semibold">Tabelas conectadas:</span>
            <span className="bg-slate-100 text-slate-800 font-bold px-2 py-0.5 rounded-md">
              {empresas.length} empresas
            </span>
            <span className="bg-slate-100 text-slate-800 font-bold px-2 py-0.5 rounded-md">
              {diagnosticos.length} diagnósticos
            </span>
            <span className="bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold px-2 py-0.5 rounded-md">
              {respostas.length} respostas sincronizadas
            </span>
          </div>

          {carregandoTabelas && (
            <span className="text-emerald-600 font-bold animate-pulse">
              Carregando dados das tabelas...
            </span>
          )}
        </div>
      </div>

      {/* Barra de Filtros Avançados */}
      <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-sm">
        <div className="flex items-center justify-between gap-2 mb-3 pb-2 border-b border-slate-100">
          <div className="flex items-center gap-2 text-slate-700 font-bold text-sm">
            <Filter size={16} className="text-emerald-600" />
            <span>Filtros do Relatório</span>
          </div>

          <div className="flex items-center gap-3">
            <span className="text-xs text-slate-500 font-medium">
              {dadosConsolidados.totalItens} itens exibidos
            </span>
            <button
              onClick={handleLimparFiltros}
              className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-rose-600 font-semibold transition-colors cursor-pointer"
              title="Restaurar todos os filtros"
            >
              <RotateCcw size={13} />
              <span>Limpar Filtros</span>
            </button>
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3">
          {/* Filtro: Tipo de Negócio */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Tipo de Negócio
            </label>
            <select
              value={tipoNegocio}
              onChange={(e) => setTipoNegocio(e.target.value)}
              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
            >
              <option value="Todos">Todos os tipos ({dadosConsolidados.opcoesTiposNegocio.length})</option>
              {dadosConsolidados.opcoesTiposNegocio.map(t => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </div>

          {/* Filtro: Código SGF */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Código no SGF
            </label>
            <select
              value={codigoSgf}
              onChange={(e) => setCodigoSgf(e.target.value)}
              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
            >
              <option value="Todos">Todos os SGF ({dadosConsolidados.opcoesCodigosSgf.length})</option>
              {dadosConsolidados.opcoesCodigosSgf.map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>
          </div>

          {/* Filtro: Área Diagnosticada */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Área Diagnosticada
            </label>
            <select
              value={area}
              onChange={(e) => setArea(e.target.value)}
              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
            >
              <option value="Todas">Todas as áreas ({dadosConsolidados.opcoesAreas.length})</option>
              {dadosConsolidados.opcoesAreas.map(a => (
                <option key={a} value={a}>{a}</option>
              ))}
            </select>
          </div>

          {/* Filtro: Cliente */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Cliente / Empresa
            </label>
            <select
              value={empresaId}
              onChange={(e) => setEmpresaId(e.target.value)}
              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
            >
              <option value="Todas">Todos os clientes ({dadosConsolidados.opcoesClientes.length})</option>
              {dadosConsolidados.opcoesClientes.map(c => (
                <option key={c.id} value={c.id}>{c.nome}</option>
              ))}
            </select>
          </div>

          {/* Filtro: Classificação */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Classificação
            </label>
            <select
              value={classificacao}
              onChange={(e) => setClassificacao(e.target.value as any)}
              className="w-full px-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-800 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
            >
              <option value="respondidas">Somente Respondidas (Fortes + Críticos)</option>
              <option value="pontos_fortes">Somente Pontos Fortes</option>
              <option value="pontos_criticos">Somente Pontos Críticos / Lacunas</option>
              <option value="todos">Todas (Incluindo Pendentes)</option>
            </select>
          </div>

          {/* Filtro: Busca por texto */}
          <div>
            <label className="block text-[11px] font-bold text-slate-600 uppercase mb-1">
              Busca Rápida
            </label>
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                type="text"
                placeholder="Pergunta, cliente..."
                value={buscaTexto}
                onChange={(e) => setBuscaTexto(e.target.value)}
                className="w-full pl-8 pr-2.5 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs text-slate-800 placeholder-slate-400 focus:bg-white focus:outline-none focus:ring-2 focus:ring-emerald-200"
              />
            </div>
          </div>
        </div>
      </div>

      {/* Cartões de Indicadores (KPIs) */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-white rounded-2xl border border-slate-200/80 p-4 shadow-sm">
          <div className="flex items-center justify-between text-slate-500 text-xs font-bold uppercase mb-1">
            <span>Clientes</span>
            <Building2 size={16} className="text-slate-400" />
          </div>
          <div className="text-2xl font-black text-slate-900">{dadosConsolidados.totalClientes}</div>
          <p className="text-[11px] text-slate-500 mt-0.5">Em {dadosConsolidados.comparativoPorSgf.length} contratos SGF</p>
        </div>

        <div className="bg-white rounded-2xl border border-slate-200/80 p-4 shadow-sm">
          <div className="flex items-center justify-between text-slate-500 text-xs font-bold uppercase mb-1">
            <span>Total Avaliações</span>
            <Layers size={16} className="text-slate-400" />
          </div>
          <div className="text-2xl font-black text-slate-900">{dadosConsolidados.totalItens}</div>
          <p className="text-[11px] text-slate-500 mt-0.5">
            {dadosConsolidados.totalPendentes > 0 ? `${dadosConsolidados.totalPendentes} pendentes` : '100% avaliadas'}
          </p>
        </div>

        <div className="bg-white rounded-2xl border border-emerald-200/80 p-4 shadow-sm bg-gradient-to-br from-white to-emerald-50/40">
          <div className="flex items-center justify-between text-emerald-800 text-xs font-bold uppercase mb-1">
            <span>Pontos Fortes</span>
            <CheckCircle2 size={16} className="text-emerald-600" />
          </div>
          <div className="text-2xl font-black text-emerald-700">
            {dadosConsolidados.totalPontosFortes}
            <span className="text-xs font-semibold text-emerald-600 ml-1.5">
              ({dadosConsolidados.percentualGeralPontosFortes}%)
            </span>
          </div>
          <div className="w-full bg-emerald-100 rounded-full h-1.5 mt-2">
            <div
              className="bg-emerald-600 h-1.5 rounded-full"
              style={{ width: `${dadosConsolidados.percentualGeralPontosFortes}%` }}
            />
          </div>
        </div>

        <div className="bg-white rounded-2xl border border-rose-200/80 p-4 shadow-sm bg-gradient-to-br from-white to-rose-50/40">
          <div className="flex items-center justify-between text-rose-800 text-xs font-bold uppercase mb-1">
            <span>Pontos Críticos</span>
            <AlertCircle size={16} className="text-rose-600" />
          </div>
          <div className="text-2xl font-black text-rose-700">
            {dadosConsolidados.totalPontosCriticos + dadosConsolidados.totalOportunidades}
            <span className="text-xs font-semibold text-rose-600 ml-1.5">
              ({dadosConsolidados.percentualGeralPontosCriticos}%)
            </span>
          </div>
          <div className="w-full bg-rose-100 rounded-full h-1.5 mt-2">
            <div
              className="bg-rose-600 h-1.5 rounded-full"
              style={{ width: `${dadosConsolidados.percentualGeralPontosCriticos}%` }}
            />
          </div>
        </div>
      </div>

      {/* Seção de Gráficos Comparativos */}
      <div className="bg-white rounded-2xl border border-slate-200/80 p-6 shadow-sm">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-2">
            <BarChart3 size={18} className="text-emerald-600" />
            <h2 className="text-lg font-extrabold text-slate-900">Gráficos Comparativos</h2>
          </div>
          <button
            onClick={() => setMostrarGraficos(!mostrarGraficos)}
            className="flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-800 font-semibold cursor-pointer"
          >
            <span>{mostrarGraficos ? 'Ocultar Gráficos' : 'Mostrar Gráficos'}</span>
            {mostrarGraficos ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          </button>
        </div>

        {mostrarGraficos && (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 pt-2">
            {/* Gráfico 1: Comparativo por Área */}
            <div className="lg:col-span-2 bg-slate-50/80 rounded-xl p-4 border border-slate-200/60">
              <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-3">
                Distribuição de Pontos Fortes vs Pontos Críticos por Área
              </h3>
              <div className="h-64">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart
                    data={dadosConsolidados.comparativoPorArea.slice(0, 8)}
                    margin={{ top: 10, right: 10, left: -20, bottom: 20 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                    <XAxis
                      dataKey="area"
                      tick={{ fontSize: 10, fill: '#475569' }}
                      angle={-20}
                      textAnchor="end"
                      height={45}
                    />
                    <YAxis tick={{ fontSize: 10, fill: '#475569' }} />
                    <Tooltip
                      contentStyle={{ backgroundColor: '#ffffff', borderRadius: '12px', border: '1px solid #e2e8f0', fontSize: '12px' }}
                    />
                    <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '6px' }} />
                    <Bar dataKey="pontosFortes" name="Pontos Fortes" fill="#059669" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="pontosCriticos" name="Pontos Críticos" fill="#e11d48" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Gráfico 2: Proporção Geral (Donut) */}
            <div className="bg-slate-50/80 rounded-xl p-4 border border-slate-200/60 flex flex-col items-center justify-center">
              <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 self-start">
                Proporção Geral das Avaliações
              </h3>
              <div className="h-48 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={dadosGraficoDonut}
                      cx="50%"
                      cy="50%"
                      innerRadius={45}
                      outerRadius={70}
                      paddingAngle={3}
                      dataKey="value"
                    >
                      {dadosGraficoDonut.map((_, index) => (
                        <Cell key={`cell-${index}`} fill={CORES_DONUT[index % CORES_DONUT.length]} />
                      ))}
                    </Pie>
                    <Tooltip
                      contentStyle={{ backgroundColor: '#ffffff', borderRadius: '12px', border: '1px solid #e2e8f0', fontSize: '12px' }}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              <div className="flex flex-wrap items-center justify-center gap-3 text-xs mt-2">
                <div className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-600" />
                  <span className="text-slate-600 font-medium">Fortes ({dadosConsolidados.percentualGeralPontosFortes}%)</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-rose-600" />
                  <span className="text-slate-600 font-medium">Críticos ({dadosConsolidados.percentualGeralPontosCriticos}%)</span>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Tabela Concatenada: Tipo de Negócio -> Código SGF -> Área */}
      <div className="space-y-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Layers size={18} className="text-emerald-600" />
            <h2 className="text-lg font-extrabold text-slate-900">
              Tabela Concatenada por Tipo de Negócio, Código SGF e Área
            </h2>
          </div>
          <span className="text-xs text-slate-500 font-medium">
            {dadosConsolidados.grupos.length} segmento(s) listado(s)
          </span>
        </div>

        {carregandoTabelas ? (
          <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center">
            <RefreshCw size={36} className="mx-auto text-emerald-600 animate-spin mb-3" />
            <h3 className="text-base font-bold text-slate-700">Carregando tabelas do sistema...</h3>
            <p className="text-xs text-slate-500 mt-1">Conectando empresas, diagnósticos, SGF e questionários de avaliação.</p>
          </div>
        ) : dadosConsolidados.grupos.length === 0 ? (
          <div className="bg-white rounded-2xl border border-slate-200 p-12 text-center">
            <HelpCircle size={40} className="mx-auto text-slate-300 mb-3" />
            <h3 className="text-base font-bold text-slate-700">Nenhum resultado encontrado</h3>
            <p className="text-xs text-slate-500 mt-1 max-w-md mx-auto">
              Nenhuma avaliação corresponde aos filtros aplicados. Experimente selecionar "Todas (Incluindo Pendentes)" ou clique em Atualizar Dados para recarregar do banco.
            </p>
            <button
              onClick={handleLimparFiltros}
              className="mt-4 px-4 py-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-700 text-xs font-bold rounded-xl transition-all cursor-pointer"
            >
              Limpar Filtros
            </button>
          </div>
        ) : (
          dadosConsolidados.grupos.map((grupoTipo) => {
            const chaveTipo = `tipo_${grupoTipo.tipoNegocio}`;
            const abertoTipo = gruposAbertos[chaveTipo] !== false;

            return (
              <div key={grupoTipo.tipoNegocio} className="bg-white rounded-2xl border border-slate-200/80 shadow-sm overflow-hidden">
                {/* Nível 1: Tipo de Negócio Header */}
                <button
                  onClick={() => toggleGrupo(chaveTipo)}
                  className="w-full flex items-center justify-between px-6 py-4 bg-gradient-to-r from-slate-900 to-slate-800 text-white cursor-pointer hover:bg-slate-850 transition-colors"
                >
                  <div className="flex items-center gap-3">
                    <Briefcase size={18} className="text-emerald-400" />
                    <span className="text-sm font-black tracking-wide uppercase">
                      Tipo de Negócio: {grupoTipo.tipoNegocio}
                    </span>
                    <span className="bg-white/10 text-emerald-300 text-[11px] font-bold px-2 py-0.5 rounded-full">
                      {grupoTipo.totalItens} pergunta(s)
                    </span>
                  </div>

                  <div className="flex items-center gap-4 text-xs font-semibold">
                    <span className="text-emerald-400">{grupoTipo.pontosFortes} Fortes</span>
                    <span className="text-rose-400">{grupoTipo.pontosCriticos} Críticos</span>
                    {abertoTipo ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                  </div>
                </button>

                {abertoTipo && (
                  <div className="p-4 sm:p-6 space-y-5 bg-slate-50/50">
                    {grupoTipo.sgfs.map((grupoSgf) => {
                      const chaveSgf = `sgf_${grupoTipo.tipoNegocio}_${grupoSgf.codigoSgf}`;
                      const abertoSgf = gruposAbertos[chaveSgf] !== false;

                      return (
                        <div key={grupoSgf.codigoSgf} className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                          {/* Nível 2: Código SGF Subheader */}
                          <button
                            onClick={() => toggleGrupo(chaveSgf)}
                            className="w-full flex items-center justify-between px-4 py-3 bg-slate-100 hover:bg-slate-150 text-slate-800 border-b border-slate-200 cursor-pointer transition-colors"
                          >
                            <div className="flex items-center gap-2">
                              <Tag size={15} className="text-blue-600" />
                              <span className="text-xs font-black text-slate-800">
                                Código de Contratação SGF: {grupoSgf.codigoSgf}
                              </span>
                              <span className="text-[10px] text-slate-500 font-bold bg-white px-2 py-0.5 rounded-full border border-slate-200">
                                {grupoSgf.areas.length} área(s)
                              </span>
                            </div>

                            <div className="flex items-center gap-3 text-xs">
                              <span className="text-emerald-700 font-bold">{grupoSgf.pontosFortes} Fortes</span>
                              <span className="text-rose-700 font-bold">{grupoSgf.pontosCriticos} Críticos</span>
                              {abertoSgf ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
                            </div>
                          </button>

                          {abertoSgf && (
                            <div className="p-4 space-y-4">
                              {grupoSgf.areas.map((grupoArea) => (
                                <div key={grupoArea.area} className="border border-slate-200/90 rounded-xl overflow-hidden">
                                  {/* Nível 3: Área Diagnosticada Header */}
                                  <div className="px-4 py-2.5 bg-slate-50 flex items-center justify-between border-b border-slate-200">
                                    <span className="text-xs font-black text-blue-900 uppercase">
                                      Área: {grupoArea.area}
                                    </span>
                                    <div className="flex items-center gap-2 text-[11px] font-bold">
                                      <span className="text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-200">
                                        {grupoArea.pontosFortes} Ponto(s) Forte(s)
                                      </span>
                                      <span className="text-rose-700 bg-rose-50 px-2 py-0.5 rounded-md border border-rose-200">
                                        {grupoArea.pontosCriticos} Ponto(s) Crítico(s)
                                      </span>
                                    </div>
                                  </div>

                                  {/* Tabela de Perguntas / Respostas */}
                                  <div className="overflow-x-auto">
                                    <table className="w-full text-left text-xs">
                                      <thead className="bg-slate-100 text-slate-700 font-bold uppercase text-[10px] tracking-wider border-b border-slate-200">
                                        <tr>
                                          <th className="py-2.5 px-3 w-44">Cliente</th>
                                          <th className="py-2.5 px-3">Pergunta Diagnosticada</th>
                                          <th className="py-2.5 px-3 w-24 text-center">Resposta</th>
                                          <th className="py-2.5 px-3 w-36 text-center">Classificação</th>
                                          <th className="py-2.5 px-3 w-64">Observações</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {grupoArea.itens.map((item) => (
                                          <tr key={item.id} className="hover:bg-slate-50/80 transition-colors">
                                            <td className="py-2.5 px-3 font-bold text-slate-800">
                                              {item.cliente}
                                            </td>
                                            <td className="py-2.5 px-3 text-slate-700">
                                              {item.pergunta}
                                            </td>
                                            <td className="py-2.5 px-3 text-center">
                                              <span className={`inline-block font-extrabold px-2 py-0.5 rounded text-[11px] ${
                                                item.resposta === 'Sim'
                                                  ? 'text-emerald-700 bg-emerald-100/60'
                                                  : item.resposta === 'Não'
                                                  ? 'text-rose-700 bg-rose-100/60'
                                                  : item.resposta === 'Parcial'
                                                  ? 'text-amber-700 bg-amber-100/60'
                                                  : 'text-slate-500 bg-slate-100'
                                              }`}>
                                                {item.resposta || 'Pendente'}
                                              </span>
                                            </td>
                                            <td className="py-2.5 px-3 text-center">
                                              <span className={`inline-flex items-center gap-1 font-bold text-[10px] px-2.5 py-1 rounded-full uppercase tracking-wider ${
                                                item.isPontoForte
                                                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                                                  : item.classificacao === 'Ponto Crítico'
                                                  ? 'bg-rose-100 text-rose-800 border border-rose-200'
                                                  : item.classificacao === 'Oportunidade de Melhoria'
                                                  ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                                  : 'bg-slate-100 text-slate-600 border border-slate-200'
                                              }`}>
                                                {item.isPontoForte ? (
                                                  <CheckCircle2 size={11} className="text-emerald-600" />
                                                ) : item.isPontoCritico ? (
                                                  <AlertCircle size={11} className="text-rose-600" />
                                                ) : (
                                                  <Clock size={11} className="text-slate-400" />
                                                )}
                                                <span>{item.classificacao}</span>
                                              </span>
                                            </td>
                                            <td className="py-2.5 px-3 text-slate-500 text-[11px] italic">
                                              {item.observacao || '—'}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};

export default RelatorioConsolidadoSgfView;
