import React, { useState, useEffect, useMemo } from 'react';
import { collection, query, where, getDocs } from 'firebase/firestore';
import { db } from '../firebase';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '../lib/utils';
import { 
  ResponsiveContainer, 
  LineChart, 
  Line, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  BarChart, 
  Bar, 
  Cell 
} from 'recharts';
import { 
  Loader2, 
  TrendingUp, 
  TrendingDown, 
  Users, 
  Target, 
  Calendar, 
  Sparkles, 
  Download, 
  Search, 
  Building2, 
  Filter, 
  Check, 
  X, 
  Info, 
  AlertTriangle, 
  Trophy, 
  ChevronRight,
  LineChart as LineChartIcon,
  HelpCircle,
  FileSpreadsheet,
  Cloud,
  Database,
  RefreshCw,
  CheckCircle2,
  ArrowUpRight,
  ArrowDownLeft,
  ShieldCheck,
  Layers,
  FileText,
  CheckSquare,
  Clock,
  Activity,
  ChevronDown,
  ChevronUp,
  Server
} from 'lucide-react';
import { getAI, extractAndParseJSON, Type } from '../App';
import { SyncSummary } from './SmartSyncModal';

// --- Custom Interfaces ---
interface Empresa {
  id: string;
  nome: string;
  tipoEmpresa?: string;
  contato?: string;
  cnpj?: string;
  userId?: string;
  criadoEm?: any;
}

interface Diagnostico {
  id: string;
  empresaId: string;
  dataDiagnostico: any;
  ownerId: string;
  nivelMaturidadeAI?: string;
  justificativaMaturidadeAI?: string;
  tipoEmpresa?: string;
  status?: string;
}

interface Resposta {
  id?: string;
  diagnosticoId: string;
  perguntaId: string;
  pergunta: string;
  resposta: 'Sim' | 'Não' | 'Parcial';
  peso: number;
  score: number;
  area: string;
}

interface MacroItem {
  id: string;
  empresaId: string;
  empresaNome: string;
  tipoEmpresa: string;
  dataObj: Date;
  dataStr: string;
  timestamp: number;
  score: number;
  areaBreakdown: { [area: string]: number };
  totalAnswersCount: number;
  yesAnswersCount: number;
  partialAnswersCount: number;
  noAnswersCount: number;
}

// Sparkline component using inline SVG
const Sparkline = ({ data }: { data: number[] }) => {
  if (!data || data.length < 2) {
    return <span className="text-slate-300 text-xs font-mono">-</span>;
  }
  
  const width = 80;
  const height = 24;
  const min = Math.min(...data);
  const max = Math.max(...data);
  const range = max - min === 0 ? 1 : max - min;
  
  const points = data
    .map((val, index) => {
      const x = (index / (data.length - 1)) * width;
      const y = height - ((val - min) / range) * height;
      return `${x},${y}`;
    })
    .join(' ');

  const isUp = data[data.length - 1] >= data[0];

  return (
    <div className="flex items-center gap-2">
      <svg width={width} height={height} className="overflow-visible">
        <polyline
          fill="none"
          stroke={isUp ? "#10b981" : "#f43f5e"}
          strokeWidth="1.8"
          strokeLinecap="round"
          strokeLinejoin="round"
          points={points}
        />
        <circle
          cx={width}
          cy={height - ((data[data.length - 1] - min) / range) * height}
          r="2.5"
          className={isUp ? "fill-emerald-500 stroke-emerald-100" : "fill-rose-500 stroke-rose-100"}
          strokeWidth="1"
        />
      </svg>
    </div>
  );
};

export const MacroDashboardView = ({
  empresas,
  diagnosticos,
  onNavigateToDiagnosis,
  lastSyncSummary,
  isSyncingCloud,
  onOpenSmartSync,
  onStartSync,
  userEmail,
  storageMode = 'cloud'
}: {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  onNavigateToDiagnosis?: (diagnostico: Diagnostico) => void;
  lastSyncSummary?: SyncSummary | null;
  isSyncingCloud?: boolean;
  onOpenSmartSync?: () => void;
  onStartSync?: () => Promise<void>;
  userEmail?: string | null;
  storageMode?: 'cloud' | 'local';
}) => {
  const [loadingMacro, setLoadingMacro] = useState(false);
  const [macroHistory, setMacroHistory] = useState<MacroItem[]>([]);
  const [selectedCompanies, setSelectedCompanies] = useState<string[]>([]);
  const [activeTab, setActiveTab] = useState<'visao-geral' | 'setores' | 'clientes'>('visao-geral');
  const [searchQuery, setSearchQuery] = useState('');
  const [sectorFilter, setSectorFilter] = useState('Todos');
  const [isSyncLogExpanded, setIsSyncLogExpanded] = useState(false);
  
  // Local fallback for lastSyncSummary if not passed from parent
  const activeSummary = useMemo<SyncSummary | null>(() => {
    if (lastSyncSummary) return lastSyncSummary;
    try {
      const saved = localStorage.getItem('last_sync_summary');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  }, [lastSyncSummary]);

  // AI Insights State
  const [aiInsight, setAiInsight] = useState<string | null>(null);
  const [generatingInsight, setGeneratingInsight] = useState(false);

  // Fetch full answers for all diagnostics to calculate exact maturity score history
  useEffect(() => {
    if (diagnosticos.length === 0) {
      setMacroHistory([]);
      return;
    }
    
    const fetchMacroData = async () => {
      setLoadingMacro(true);
      try {
        const chunkSize = 30;
        const chunks = [];
        for (let i = 0; i < diagnosticos.length; i += chunkSize) {
          chunks.push(diagnosticos.slice(i, i + chunkSize));
        }
        
        const allRespsMap = new Map<string, any[]>();
        
        for (const chunk of chunks) {
          const ids = chunk.map(d => d.id);
          const qRes = query(collection(db, 'respostas'), where('diagnosticoId', 'in', ids));
          const snap = await getDocs(qRes);
          snap.docs.forEach(docSnap => {
            const data = docSnap.data();
            const diagId = data.diagnosticoId;
            if (!allRespsMap.has(diagId)) {
              allRespsMap.set(diagId, []);
            }
            allRespsMap.get(diagId)!.push(data);
          });
        }
        
        const items = diagnosticos.map(diag => {
          const resps = allRespsMap.get(diag.id) || [];
          const totalScore = resps.reduce((acc, r) => acc + (r.score || 0), 0);
          const totalPeso = resps.reduce((acc, r) => acc + (2 * (r.peso || 1)), 0);
          const score = totalPeso > 0 ? Math.round((totalScore / totalPeso) * 100) : 0;
          
          const areaScores: { [area: string]: { scoreTotal: number, pesoTotal: number } } = {};
          resps.forEach((r: any) => {
            const areaName = r.area || 'Geral';
            if (!areaScores[areaName]) {
              areaScores[areaName] = { scoreTotal: 0, pesoTotal: 0 };
            }
            areaScores[areaName].scoreTotal += (r.score || 0);
            areaScores[areaName].pesoTotal += (2 * (r.peso || 1));
          });
          
          const areaBreakdown: { [area: string]: number } = {};
          Object.entries(areaScores).forEach(([areaName, vals]) => {
            areaBreakdown[areaName] = vals.pesoTotal > 0 
              ? Math.round((vals.scoreTotal / vals.pesoTotal) * 100)
              : 0;
          });
          
          let dateObj = new Date();
          let dateStr = "";
          let timestamp = 0;
          
          if (diag.dataDiagnostico) {
            if (diag.dataDiagnostico.toDate) {
              dateObj = diag.dataDiagnostico.toDate();
            } else if (typeof diag.dataDiagnostico === 'string') {
              dateObj = new Date(diag.dataDiagnostico);
            } else if (diag.dataDiagnostico.seconds) {
              dateObj = new Date(diag.dataDiagnostico.seconds * 1000);
            }
            dateStr = dateObj.toISOString();
            timestamp = dateObj.getTime();
          }
          
          return {
            id: diag.id,
            empresaId: diag.empresaId,
            empresaNome: empresas.find(e => e.id === diag.empresaId)?.nome || "Não identificada",
            tipoEmpresa: diag.tipoEmpresa || "Geral",
            dataObj: dateObj,
            dataStr: dateStr,
            timestamp,
            score,
            areaBreakdown,
            totalAnswersCount: resps.length,
            yesAnswersCount: resps.filter(r => r.resposta === 'Sim').length,
            partialAnswersCount: resps.filter(r => r.resposta === 'Parcial').length,
            noAnswersCount: resps.filter(r => r.resposta === 'Não').length,
          };
        }).filter(item => item.totalAnswersCount > 0);
        
        items.sort((a, b) => a.timestamp - b.timestamp);
        setMacroHistory(items);
        
        // Auto-select top 5 clients by default for comparison if none selected
        const uniqueClientIds = Array.from(new Set(items.map(i => i.empresaId))).slice(0, 5);
        setSelectedCompanies(uniqueClientIds);
      } catch (err) {
        console.error("Erro ao processar dados macro:", err);
      } finally {
        setLoadingMacro(false);
      }
    };
    
    fetchMacroData();
  }, [diagnosticos, empresas]);

  // List of unique segments in active clients
  const availableSectors = useMemo(() => {
    const sectors = new Set<string>();
    macroHistory.forEach(h => {
      if (h.tipoEmpresa) sectors.add(h.tipoEmpresa);
    });
    return ['Todos', ...Array.from(sectors)];
  }, [macroHistory]);

  // Filter macro history based on query & sector filter
  const filteredHistory = useMemo(() => {
    return macroHistory.filter(h => {
      const matchesSearch = h.empresaNome.toLowerCase().includes(searchQuery.toLowerCase());
      const matchesSector = sectorFilter === 'Todos' || h.tipoEmpresa === sectorFilter;
      return matchesSearch && matchesSector;
    });
  }, [macroHistory, searchQuery, sectorFilter]);

  // Get active diagnosed companies matching filters
  const diagnosedCompanies = useMemo(() => {
    const activeIds = new Set(filteredHistory.map(h => h.empresaId));
    return empresas.filter(emp => activeIds.has(emp.id));
  }, [empresas, filteredHistory]);

  // Construct 12 months data specifically covering the last year (12 months up to now)
  const chartData = useMemo(() => {
    if (macroHistory.length === 0) return [];
    
    const months: { year: number, month: number, label: string, key: string }[] = [];
    const now = new Date();
    // Generate exactly last 12 months
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const label = d.toLocaleDateString('pt-BR', { month: 'short', year: '2-digit' });
      months.push({
        year: d.getFullYear(),
        month: d.getMonth(),
        label: label.charAt(0).toUpperCase() + label.slice(1).replace('.', ''),
        key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
      });
    }
    
    return months.map(m => {
      const monthEnd = new Date(m.year, m.month + 1, 0, 23, 59, 59, 999);
      
      // Calculate LOCF (Last Observation Carried Forward) for each company up to this month
      const activeShares: { [companyId: string]: number } = {};
      const diagsUpToMonth = macroHistory.filter(h => h.timestamp <= monthEnd.getTime());
      
      diagsUpToMonth.forEach(h => {
        // Only apply sector filter if specified, to keep "Média Portfolio" consistent or filtered
        if (sectorFilter === 'Todos' || h.tipoEmpresa === sectorFilter) {
          activeShares[h.empresaId] = h.score;
        }
      });
      
      const activeScores = Object.values(activeShares);
      const averagePortfolioScore = activeScores.length > 0 
        ? Math.round(activeScores.reduce((sum, s) => sum + s, 0) / activeScores.length)
        : null;
        
      const result: any = {
        name: m.label,
        'Média Portfolio': averagePortfolioScore,
      };
      
      // Add individual client values if they are selected
      selectedCompanies.forEach(cId => {
        const compDiags = diagsUpToMonth.filter(h => h.empresaId === cId);
        if (compDiags.length > 0) {
          result[cId] = compDiags[compDiags.length - 1].score;
        }
      });
      
      return result;
    });
  }, [macroHistory, selectedCompanies, sectorFilter]);

  // Sector Data for comparisons
  const sectorData = useMemo(() => {
    const segmentMap = new Map<string, { total: number, count: number }>();
    const latestCompanyDiag = new Map<string, any>();
    
    macroHistory.forEach(h => {
      latestCompanyDiag.set(h.empresaId, h);
    });
    
    latestCompanyDiag.forEach((h) => {
      const sector = h.tipoEmpresa || "Geral";
      if (!segmentMap.has(sector)) {
        segmentMap.set(sector, { total: 0, count: 0 });
      }
      const val = segmentMap.get(sector)!;
      val.total += h.score;
      val.count += 1;
    });
    
    return Array.from(segmentMap.entries()).map(([segment, data]) => ({
      segment,
      Maturidade: Math.round(data.total / data.count),
      Clientes: data.count
    })).sort((a, b) => b.Maturidade - a.Maturidade);
  }, [macroHistory]);

  // Area-specific average maturity scores across portfolio
  const areaProgressData = useMemo(() => {
    const latestCompanyDiag = new Map<string, any>();
    
    macroHistory.forEach(h => {
      // Respect sector filter when aggregating areas
      if (sectorFilter === 'Todos' || h.tipoEmpresa === sectorFilter) {
        latestCompanyDiag.set(h.empresaId, h);
      }
    });
    
    const areaMap = new Map<string, { total: number, count: number }>();
    
    latestCompanyDiag.forEach((h) => {
      if (h.areaBreakdown) {
        Object.entries(h.areaBreakdown).forEach(([area, score]) => {
          if (typeof score !== 'number') return;
          if (!areaMap.has(area)) {
            areaMap.set(area, { total: 0, count: 0 });
          }
          const val = areaMap.get(area)!;
          val.total += score;
          val.count += 1;
        });
      }
    });
    
    return Array.from(areaMap.entries()).map(([area, data]) => ({
      area,
      Maturidade: Math.round(data.total / data.count),
      Clientes: data.count
    })).sort((a, b) => b.Maturidade - a.Maturidade);
  }, [macroHistory, sectorFilter]);

  // Comprehensive Client Summary
  const clientSummary = useMemo(() => {
    const clients: any[] = [];
    
    empresas.forEach(emp => {
      const compDiags = macroHistory.filter(h => h.empresaId === emp.id);
      if (compDiags.length === 0) return;
      
      const firstDiag = compDiags[0];
      const latestDiag = compDiags[compDiags.length - 1];
      const growth = latestDiag.score - firstDiag.score;
      
      // Get all monthly scores over 12 months for sparkline
      const sparklineScores: number[] = [];
      const now = new Date();
      for (let i = 11; i >= 0; i--) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const monthEnd = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999);
        const validDiags = compDiags.filter(h => h.timestamp <= monthEnd.getTime());
        if (validDiags.length > 0) {
          sparklineScores.push(validDiags[validDiags.length - 1].score);
        } else {
          // If no diagnostic prior to this month, push the first diagnostic score as baseline
          sparklineScores.push(firstDiag.score);
        }
      }
      
      clients.push({
        id: emp.id,
        nome: emp.nome,
        segmento: latestDiag.tipoEmpresa || 'Geral',
        diagsCount: compDiags.length,
        firstDate: firstDiag.dataObj.toLocaleDateString('pt-BR'),
        firstScore: firstDiag.score,
        latestDate: latestDiag.dataObj.toLocaleDateString('pt-BR'),
        latestScore: latestDiag.score,
        growth,
        sparklineScores,
        status: growth > 15 ? 'Crescimento Forte' : growth > 2 ? 'Evolução Gradual' : growth < -2 ? 'Alerta queda' : 'Estável',
        diagnosticoOriginal: latestDiag
      });
    });
    
    // Apply filters to table too
    return clients.filter(c => {
      const matchesSearch = c.nome.toLowerCase().includes(searchQuery.toLowerCase());
      const matchesSector = sectorFilter === 'Todos' || c.segmento === sectorFilter;
      return matchesSearch && matchesSector;
    }).sort((a, b) => b.latestScore - a.latestScore);
  }, [empresas, macroHistory, searchQuery, sectorFilter]);

  // Key KPI stats calculations
  const activeAverageScore = useMemo(() => {
    const latestScores = clientSummary.map(c => c.latestScore);
    if (latestScores.length === 0) return 0;
    return Math.round(latestScores.reduce((sum, s) => sum + s, 0) / latestScores.length);
  }, [clientSummary]);

  const totalGrowth = useMemo(() => {
    const growths = clientSummary.map(c => c.growth);
    if (growths.length === 0) return 0;
    return Math.round(growths.reduce((sum, g) => sum + g, 0) / growths.length);
  }, [clientSummary]);

  const strongGrowthCount = useMemo(() => {
    return clientSummary.filter(c => c.growth > 10).length;
  }, [clientSummary]);

  // Generate automated AI insights based on current portfolio status
  const generatePortfolioAIInsight = async () => {
    if (generatingInsight) return;
    setGeneratingInsight(true);
    setAiInsight(null);
    try {
      const ai = getAI();
      if (!ai) {
        throw new Error("Cliente de IA não disponível");
      }

      // Format data for AI
      const formattedClients = clientSummary.map(c => ({
        nome: c.nome,
        segmento: c.segmento,
        notaInicial: `${c.firstScore}% (${c.firstDate})`,
        notaAtual: `${c.latestScore}% (${c.latestDate})`,
        evolucao: `${c.growth >= 0 ? '+' : ''}${c.growth}%`,
        status: c.status
      }));

      const formattedSectors = sectorData.map(s => ({
        segmento: s.segment,
        mediaMaturidade: `${s.Maturidade}%`,
        totalClientes: s.Clientes
      }));

      const formattedAreas = areaProgressData.map(a => ({
        area: a.area,
        mediaMaturidade: `${a.Maturidade}%`
      }));

      const prompt = `Como um consultor empresarial sênior focado em Customer Success e Inteligência Competitiva, analise este relatório consolidado de evolução histórica da maturidade de todos os clientes da nossa consultoria empresarial ao longo do último ano.

      DADOS DO PORTFOLIO DE CLIENTES:
      - Total de Clientes Diagnosticados: ${diagnosedCompanies.length}
      - Média de Maturidade Geral Atual da Carteira: ${activeAverageScore}%
      - Crescimento/Evolução Média por Cliente: ${totalGrowth >= 0 ? '+' : ''}${totalGrowth}%
      
      DIAGNÓSTICO DETALHADO DOS CLIENTES:
      ${JSON.stringify(formattedClients)}

      MÉDIA DE MATURIDADE POR SEGMENTO/SETOR:
      ${JSON.stringify(formattedSectors)}

      MÉDIA DE MATURIDADE POR ÁREA DE CONSULTORIA:
      ${JSON.stringify(formattedAreas)}

      INSTRUÇÕES DE RESPOSTA:
      Gere um relatório analítico sênior estruturado e extremamente focado no negócio. Escreva em Português do Brasil de forma elegante e inspiradora. O relatório deve conter:
      1. **Análise de Trajetória**: Um breve panorama de como a carteira evoluiu no último ano, destacando o crescimento médio e o que isso indica sobre o impacto da consultoria.
      2. **Destaques de Sucesso**: Cite as empresas que mais cresceram e de quais setores elas são, reconhecendo as melhores práticas.
      3. **Pontos Críticos / Gaps**: Identifique as disciplinas/áreas de diagnóstico de menor maturidade geral na carteira e explique as implicações estratégicas disso para os clientes.
      4. **Recomendações para a Consultoria**: Forneça 3 estratégias de intervenção de curto e médio prazo que a nossa consultoria pode aplicar para acelerar a maturidade das empresas nos setores críticos ou áreas fracas.

      Retorne em formato de texto limpo estruturado em markdown clássico, sem cabeçalhos repetidos desnecessários, com tom executivo de consultoria do mais alto nível.`;

      const response = await ai.models.generateContent({
        model: "gemini-3.6-flash",
        contents: prompt
      });

      if (response && response.text) {
        setAiInsight(response.text);
      } else {
        throw new Error("Resposta de IA vazia");
      }
    } catch (err: any) {
      console.error("Erro ao gerar insight de portfólio:", err);
      setAiInsight("Não foi possível gerar a análise automática com Inteligência Artificial no momento. Por favor, certifique-se de que a API Key do Gemini está configurada corretamente nas configurações.");
    } finally {
      setGeneratingInsight(false);
    }
  };

  // Export data as Excel / CSV structure helper
  const exportPortfolioData = () => {
    try {
      let csvContent = "data:text/csv;charset=utf-8,";
      csvContent += "Cliente,Segmento,Diagnosticos,Maturidade Inicial,Maturidade Atual,Evolucao,Status\n";
      
      clientSummary.forEach(c => {
        const row = `"${c.nome}","${c.segmento}",${c.diagsCount},${c.firstScore}%,${c.latestScore}%,${c.growth >= 0 ? '+' : ''}${c.growth}%,"${c.status}"`;
        csvContent += row + "\n";
      });
      
      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", `consolidado_macro_clientes_${new Date().toISOString().split('T')[0]}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } catch (err) {
      console.error("Erro ao exportar dados:", err);
      alert("Não foi possível exportar os dados.");
    }
  };

  if (loadingMacro) {
    return (
      <div className="flex flex-col items-center justify-center py-24 gap-4 bg-white/40 rounded-[2rem] border border-slate-100 backdrop-blur-sm shadow-sm">
        <Loader2 className="animate-spin text-emerald-600" size={48} />
        <p className="text-slate-500 font-medium font-sans">Processando e consolidando histórico macro do portfólio...</p>
      </div>
    );
  }

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-12">
      {/* Title Header */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <div className="flex items-center gap-2">
            <span className="bg-emerald-50 text-emerald-700 px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider border border-emerald-100">Macro</span>
            <span className="text-slate-400 text-xs font-mono">• Últimos 12 Meses</span>
          </div>
          <h2 className="text-3xl font-black tracking-tight text-slate-800 mt-1">Evolução Histórica da Maturidade</h2>
          <p className="text-slate-500 text-sm mt-0.5 font-sans">Estatísticas, análises comparativas e linha do tempo de desempenho da carteira de clientes</p>
        </div>
        
        {/* Navigation & Export Action Header */}
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button
            onClick={exportPortfolioData}
            disabled={clientSummary.length === 0}
            className="flex items-center justify-center gap-2 px-4 py-2.5 bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 hover:border-slate-300 rounded-xl font-bold text-xs transition-all shadow-sm cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed"
          >
            <FileSpreadsheet size={16} className="text-emerald-600" />
            <span>Exportar CSV</span>
          </button>
          
          <div className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab('visao-geral')}
              className={cn(
                "px-4 py-2 rounded-lg text-xs font-bold transition-all border-none cursor-pointer",
                activeTab === 'visao-geral' ? "bg-white text-slate-800 shadow" : "text-slate-500 hover:text-slate-800"
              )}
            >
              Visão Geral
            </button>
            <button
              onClick={() => setActiveTab('setores')}
              className={cn(
                "px-4 py-2 rounded-lg text-xs font-bold transition-all border-none cursor-pointer",
                activeTab === 'setores' ? "bg-white text-slate-800 shadow" : "text-slate-500 hover:text-slate-800"
              )}
            >
              Análise Setorial
            </button>
            <button
              onClick={() => setActiveTab('clientes')}
              className={cn(
                "px-4 py-2 rounded-lg text-xs font-bold transition-all border-none cursor-pointer",
                activeTab === 'clientes' ? "bg-white text-slate-800 shadow" : "text-slate-500 hover:text-slate-800"
              )}
            >
              Clientes Ativos
            </button>
          </div>
        </div>
      </div>

      {/* Detailed Firestore Cloud Sync Status Log */}
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
                {activeSummary ? activeSummary.totalAnalyzed : (empresas.length + diagnosticos.length + macroHistory.length)}
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
                {activeSummary ? activeSummary.identicalOrMerged : (empresas.length + diagnosticos.length)}
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
                        {activeSummary ? (activeSummary.details.empresas.uploaded + activeSummary.details.empresas.downloaded + activeSummary.details.empresas.updated) : empresas.length}
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
                        {activeSummary ? (activeSummary.details.diagnosticos.uploaded + activeSummary.details.diagnosticos.downloaded + activeSummary.details.diagnosticos.updated) : diagnosticos.length}
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
                        <CheckSquare size={14} className="text-indigo-400" /> Respostas
                      </span>
                      <span className="text-slate-400 font-mono text-[11px]">
                        {activeSummary ? (activeSummary.details.respostas.uploaded + activeSummary.details.respostas.downloaded + activeSummary.details.respostas.updated) : macroHistory.reduce((acc, h) => acc + h.totalAnswersCount, 0)}
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

      {diagnosticos.length === 0 ? (
        <div className="bg-white rounded-3xl border border-slate-100 p-16 text-center flex flex-col items-center justify-center shadow-sm">
          <div className="w-16 h-16 rounded-2xl bg-slate-50 flex items-center justify-center text-slate-300 mb-4 border border-slate-100/50">
            <LineChartIcon size={32} />
          </div>
          <h3 className="text-xl font-black text-slate-800 mb-1">Nenhum histórico disponível</h3>
          <p className="text-slate-500 text-sm max-w-md font-sans font-medium leading-relaxed mb-6">
            Para visualizar a evolução histórica dos seus clientes e as estatísticas de maturidade macro, comece realizando diagnósticos para as empresas cadastradas.
          </p>
        </div>
      ) : (
        <>
          {/* Portfolio Metric Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="p-6 bg-white border border-slate-100 flex items-center gap-4 shadow-sm rounded-2xl">
              <div className="w-12 h-12 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center flex-shrink-0 border border-indigo-100/40">
                <Users size={22} />
              </div>
              <div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Carteira Ativa</p>
                <p className="text-2xl font-black text-slate-800 mt-1">{diagnosedCompanies.length}</p>
                <p className="text-xs text-slate-500 mt-0.5 font-sans">de {empresas.length} cadastrados</p>
              </div>
            </div>

            <div className="p-6 bg-white border border-slate-100 flex items-center gap-4 shadow-sm rounded-2xl">
              <div className="w-12 h-12 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center flex-shrink-0 border border-emerald-100/40">
                <Target size={22} />
              </div>
              <div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Maturidade Média</p>
                <p className="text-2xl font-black text-slate-800 mt-1">{activeAverageScore}%</p>
                <p className="text-xs text-slate-500 mt-0.5 font-sans">pontuação ponderada</p>
              </div>
            </div>

            <div className="p-6 bg-white border border-slate-100 flex items-center gap-4 shadow-sm rounded-2xl">
              <div className="w-12 h-12 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center flex-shrink-0 border border-amber-100/40">
                <TrendingUp size={22} />
              </div>
              <div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Evolução Média</p>
                <p className="text-2xl font-black text-amber-600 mt-1">{totalGrowth >= 0 ? `+${totalGrowth}` : totalGrowth}%</p>
                <p className="text-xs text-slate-500 mt-0.5 font-sans">crescimento de notas</p>
              </div>
            </div>

            <div className="p-6 bg-white border border-slate-100 flex items-center gap-4 shadow-sm rounded-2xl">
              <div className="w-12 h-12 rounded-xl bg-violet-50 text-violet-600 flex items-center justify-center flex-shrink-0 border border-violet-100/40">
                <Sparkles size={22} />
              </div>
              <div>
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Crescimento Alto</p>
                <p className="text-2xl font-black text-violet-600 mt-1">{strongGrowthCount}</p>
                <p className="text-xs text-slate-500 mt-0.5 font-sans">clientes evoluídos &gt; 10%</p>
              </div>
            </div>
          </div>

          {/* Filters & Control bar */}
          <div className="bg-slate-50 rounded-2xl p-4 border border-slate-100 flex flex-col sm:flex-row gap-4 items-center justify-between">
            <div className="flex flex-wrap items-center gap-3 w-full sm:w-auto">
              <div className="relative flex-grow sm:flex-grow-0 sm:w-64">
                <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
                  <Search size={16} />
                </span>
                <input
                  type="text"
                  placeholder="Pesquisar por cliente..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full pl-9 pr-4 py-2 bg-white border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500/20 rounded-xl text-xs placeholder:text-slate-400 text-slate-700"
                />
              </div>

              <div className="flex items-center gap-1.5">
                <span className="text-slate-400 text-xs flex items-center gap-1"><Filter size={12} /> Setor:</span>
                <select
                  value={sectorFilter}
                  onChange={(e) => setSectorFilter(e.target.value)}
                  className="bg-white border border-slate-200 text-xs font-semibold text-slate-700 px-3 py-1.5 rounded-xl outline-none focus:border-indigo-500"
                >
                  {availableSectors.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              </div>
            </div>

            {/* Selected stats info helper */}
            <div className="text-right text-slate-400 text-[11px] font-sans">
              Mostrando <strong className="text-slate-600">{diagnosedCompanies.length}</strong> de <strong className="text-slate-600">{empresas.length}</strong> empresas com diagnósticos ativos.
            </div>
          </div>

          {/* TAB 1: VISÃO GERAL */}
          {activeTab === 'visao-geral' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                {/* Main Evolution Line Chart */}
                <div className="p-6 bg-white rounded-3xl border border-slate-100 lg:col-span-3 shadow-sm flex flex-col justify-between space-y-4">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
                    <div>
                      <h3 className="font-bold text-lg text-slate-800">Evolução Mensal de Maturidade</h3>
                      <p className="text-slate-400 text-xs mt-0.5 font-sans">Linha do tempo baseada na última nota de maturidade de cada cliente (LOCF)</p>
                    </div>
                    
                    {/* Color Indicators */}
                    <div className="flex flex-wrap gap-2 text-[10px]">
                      <span className="inline-flex items-center gap-1.5 px-2 py-1 bg-emerald-50 text-emerald-800 rounded-lg font-bold border border-emerald-100">
                        <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                        Média da Carteira
                      </span>
                      {selectedCompanies.map((cId, idx) => {
                        const name = empresas.find(e => e.id === cId)?.nome || "Empresa";
                        const lineColors = ['#6366f1', '#f59e0b', '#3b82f6', '#ec4899', '#8b5cf6'];
                        const color = lineColors[idx % lineColors.length];
                        return (
                          <span 
                            key={cId} 
                            style={{ color, borderColor: color + '25', backgroundColor: color + '05' }} 
                            className="inline-flex items-center gap-1.5 px-2 py-1 border rounded-lg font-bold"
                          >
                            <span style={{ backgroundColor: color }} className="w-2 h-2 rounded-full"></span>
                            {name}
                          </span>
                        );
                      })}
                    </div>
                  </div>

                  <div className="h-80 w-full pt-2">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f8fafc" />
                        <XAxis 
                          dataKey="name" 
                          stroke="#94a3b8" 
                          fontSize={11} 
                          tickLine={false} 
                          axisLine={false} 
                        />
                        <YAxis 
                          stroke="#94a3b8" 
                          fontSize={11} 
                          domain={[0, 100]} 
                          tickLine={false} 
                          axisLine={false} 
                          tickFormatter={(v) => `${v}%`} 
                        />
                        <Tooltip 
                          contentStyle={{ 
                            borderRadius: '1rem', 
                            border: '1px solid #f1f5f9', 
                            boxShadow: '0 10px 25px -5px rgba(0,0,0,0.05)',
                            fontFamily: 'Inter, sans-serif',
                            fontSize: '12px'
                          }}
                          formatter={(value: any, name: string) => {
                            let labelName = name;
                            if (name !== 'Média Portfolio' && name !== 'Média da Carteira') {
                              const comp = empresas.find(e => e.id === name);
                              if (comp) labelName = comp.nome;
                            }
                            return [`${value}%`, labelName];
                          }}
                        />
                        {/* Portfolio Average Line */}
                        <Line 
                          type="monotone" 
                          dataKey="Média Portfolio" 
                          name="Média Portfolio"
                          stroke="#10b981" 
                          strokeWidth={3} 
                          dot={{ r: 4.5, strokeWidth: 0, fill: '#10b981' }} 
                          activeDot={{ r: 6 }} 
                          connectNulls 
                        />
                        
                        {/* Selected Individual Companies Lines */}
                        {selectedCompanies.map((cId, idx) => {
                          const lineColors = ['#6366f1', '#f59e0b', '#3b82f6', '#ec4899', '#8b5cf6'];
                          return (
                            <Line
                              key={cId}
                              type="monotone"
                              dataKey={cId}
                              name={cId}
                              stroke={lineColors[idx % lineColors.length]}
                              strokeWidth={2}
                              dot={{ r: 3 }}
                              connectNulls
                            />
                          );
                        })}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Company Comparison Selector (Right Panel) */}
                <div className="p-6 bg-white rounded-3xl border border-slate-100 shadow-sm flex flex-col justify-between space-y-4">
                  <div>
                    <h3 className="font-bold text-base text-slate-800">Comparar Clientes</h3>
                    <p className="text-slate-400 text-xs mt-0.5 font-sans">Selecione até 5 clientes para plotar individualmente no gráfico</p>
                  </div>

                  <div className="space-y-1.5 flex-grow overflow-y-auto max-h-[17rem] pr-1 scrollbar-thin">
                    {diagnosedCompanies.length === 0 ? (
                      <p className="text-slate-400 text-xs italic p-4 text-center">Nenhuma empresa encontrada com os filtros atuais.</p>
                    ) : (
                      diagnosedCompanies.map(emp => {
                        const isSelected = selectedCompanies.includes(emp.id);
                        const lastDiag = macroHistory.filter(h => h.empresaId === emp.id).pop();
                        return (
                          <label 
                            key={emp.id}
                            className={cn(
                              "flex items-center justify-between p-2 rounded-lg border cursor-pointer transition-all",
                              isSelected 
                                ? "bg-slate-50 border-indigo-200 text-indigo-950 font-semibold"
                                : "bg-white border-slate-100 text-slate-600 hover:bg-slate-50/50"
                            )}
                          >
                            <div className="flex items-center gap-2">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                disabled={!isSelected && selectedCompanies.length >= 5}
                                onChange={() => {
                                  if (isSelected) {
                                    setSelectedCompanies(selectedCompanies.filter(id => id !== emp.id));
                                  } else if (selectedCompanies.length < 5) {
                                    setSelectedCompanies([...selectedCompanies, emp.id]);
                                  }
                                }}
                                className="w-3.5 h-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500/20 cursor-pointer"
                              />
                              <span className="text-xs truncate max-w-32 font-sans">{emp.nome}</span>
                            </div>
                            {lastDiag && (
                              <span className="text-[10px] font-black px-1.5 py-0.5 bg-slate-100 text-slate-600 rounded">
                                {lastDiag.score}%
                              </span>
                            )}
                          </label>
                        );
                      })
                    )}
                  </div>

                  <div className="pt-3 border-t border-slate-100 flex gap-1.5">
                    <button
                      onClick={() => setSelectedCompanies(diagnosedCompanies.map(d => d.id).slice(0, 5))}
                      className="flex-1 py-1.5 text-[10px] font-bold text-slate-600 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg cursor-pointer transition-colors"
                    >
                      Padrão (Top 5)
                    </button>
                    <button
                      onClick={() => setSelectedCompanies([])}
                      className="flex-1 py-1.5 text-[10px] font-bold text-slate-600 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-lg cursor-pointer transition-colors"
                    >
                      Limpar
                    </button>
                  </div>
                </div>
              </div>

              {/* Area Breakdown & Highlights */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div className="p-6 bg-white rounded-3xl border border-slate-100 shadow-sm lg:col-span-2 space-y-4">
                  <div>
                    <h3 className="font-bold text-lg text-slate-800">Maturidade Média por Área de Diagnóstico</h3>
                    <p className="text-slate-400 text-xs mt-0.5 font-sans">Pontuação média ponderada acumulada nas disciplinas avaliadas no portfólio</p>
                  </div>

                  <div className="h-80 w-full pt-2">
                    {areaProgressData.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={areaProgressData} layout="vertical" margin={{ top: 10, right: 20, left: 20, bottom: 5 }}>
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f8fafc" />
                          <XAxis type="number" stroke="#94a3b8" fontSize={11} domain={[0, 100]} tickLine={false} axisLine={false} tickFormatter={(v) => `${v}%`} />
                          <YAxis dataKey="area" type="category" stroke="#64748b" fontSize={11} tickLine={false} axisLine={false} width={130} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '1rem', border: '1px solid #f1f5f9', fontFamily: 'Inter, sans-serif' }}
                            formatter={(value: any) => [`${value}%`, 'Maturidade Média']}
                          />
                          <Bar dataKey="Maturidade" radius={[0, 8, 8, 0]} barSize={15}>
                            {areaProgressData.map((entry, index) => {
                              const colors = ['#6366f1', '#10b981', '#f59e0b', '#3b82f6', '#ec4899', '#8b5cf6', '#14b8a6', '#f43f5e'];
                              return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />;
                            })}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="flex items-center justify-center h-full text-slate-400 text-sm font-sans italic">
                        Sem dados de diagnóstico agregados por área.
                      </div>
                    )}
                  </div>
                </div>

                {/* Highlights and Intervention opportunities */}
                <div className="p-6 bg-white rounded-3xl border border-slate-100 shadow-sm flex flex-col justify-between space-y-4">
                  <div className="space-y-4">
                    <div>
                      <h3 className="font-bold text-base text-slate-800">Análise de Oportunidades</h3>
                      <p className="text-slate-400 text-xs mt-0.5 font-sans">Disciplinas e áreas que demandam atenção na carteira</p>
                    </div>

                    <div className="space-y-3 pt-1">
                      {areaProgressData.length > 0 ? (
                        <>
                          {/* Strongest Area */}
                          <div className="p-3.5 bg-emerald-50/40 rounded-xl border border-emerald-100/50">
                            <div className="flex items-center gap-2">
                              <span className="w-6 h-6 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center">
                                <Trophy size={13} />
                              </span>
                              <span className="text-[9px] font-black tracking-widest text-emerald-700 uppercase">Maior Maturidade</span>
                            </div>
                            <p className="text-sm font-bold text-slate-800 mt-2 font-sans">{areaProgressData[0].area}</p>
                            <div className="flex items-baseline gap-1 mt-0.5">
                              <span className="text-xl font-black text-emerald-600">{areaProgressData[0].Maturidade}%</span>
                              <span className="text-[10px] text-slate-400 font-sans">média de conformidade</span>
                            </div>
                          </div>

                          {/* Weakest Area */}
                          <div className="p-3.5 bg-rose-50/40 rounded-xl border border-rose-100/50">
                            <div className="flex items-center gap-2">
                              <span className="w-6 h-6 rounded-lg bg-rose-100 text-rose-700 flex items-center justify-center">
                                <AlertTriangle size={13} />
                              </span>
                              <span className="text-[9px] font-black tracking-widest text-rose-700 uppercase">Maior Gargalo / Atenção</span>
                            </div>
                            <p className="text-sm font-bold text-slate-800 mt-2 font-sans">{areaProgressData[areaProgressData.length - 1].area}</p>
                            <div className="flex items-baseline gap-1 mt-0.5">
                              <span className="text-xl font-black text-rose-600">{areaProgressData[areaProgressData.length - 1].Maturidade}%</span>
                              <span className="text-[10px] text-slate-400 font-sans">necessita de planos de ação</span>
                            </div>
                          </div>
                        </>
                      ) : (
                        <p className="text-xs text-slate-400 font-sans italic">Dados insuficientes para sugerir oportunidades.</p>
                      )}
                    </div>
                  </div>

                  <div className="text-[11px] text-slate-500 leading-relaxed pt-3 border-t border-slate-100">
                    💡 <strong>Insight Proativo:</strong> A disciplina de <strong className="text-slate-700">{areaProgressData[areaProgressData.length - 1]?.area || 'áreas com menor nota'}</strong> é a que apresenta maior gap médio em seus clientes. Considere sugerir as premissas e soluções pré-cadastradas para essa área.
                  </div>
                </div>
              </div>

              {/* AI Strategic Insights Card */}
              <div className="p-6 bg-gradient-to-br from-indigo-900 via-indigo-950 to-slate-950 rounded-3xl text-white shadow-md relative overflow-hidden border border-indigo-800/20">
                <div className="absolute top-0 right-0 w-80 h-80 bg-indigo-500/10 rounded-full -translate-y-1/3 translate-x-1/3 blur-3xl pointer-events-none"></div>
                <div className="relative z-10 space-y-4">
                  <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                    <div className="flex items-center gap-2.5">
                      <div className="w-10 h-10 rounded-xl bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 flex items-center justify-center">
                        <Sparkles size={20} className="animate-pulse" />
                      </div>
                      <div>
                        <h4 className="font-bold text-base">Relatório Estratégico com Inteligência Artificial</h4>
                        <p className="text-indigo-200 text-xs font-sans">Gere um diagnóstico avançado com recomendações seniores para a sua carteira</p>
                      </div>
                    </div>

                    <button
                      onClick={generatePortfolioAIInsight}
                      disabled={generatingInsight}
                      className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white hover:text-white rounded-xl text-xs font-bold transition-all border-none cursor-pointer flex items-center gap-2 shadow-lg hover:shadow-indigo-500/20 disabled:opacity-50 disabled:cursor-not-allowed"
                    >
                      {generatingInsight ? (
                        <>
                          <Loader2 size={14} className="animate-spin" />
                          <span>Analisando Dados...</span>
                        </>
                      ) : (
                        <>
                          <Sparkles size={14} />
                          <span>Gerar Insights de Consultoria</span>
                        </>
                      )}
                    </button>
                  </div>

                  {aiInsight && (
                    <motion.div
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="p-5 bg-white/5 border border-white/10 rounded-2xl text-xs text-indigo-100 font-sans leading-relaxed overflow-y-auto max-h-[24rem] scrollbar-thin mt-2"
                    >
                      <div className="prose prose-invert prose-xs max-w-none text-slate-100">
                        {aiInsight.split('\n').map((line, idx) => {
                          if (line.startsWith('### ')) {
                            return <h5 key={idx} className="font-bold text-slate-200 text-sm mt-3 mb-1">{line.replace('### ', '')}</h5>;
                          }
                          if (line.startsWith('## ')) {
                            return <h4 key={idx} className="font-extrabold text-white text-base mt-4 mb-2 border-b border-white/10 pb-1">{line.replace('## ', '')}</h4>;
                          }
                          if (line.startsWith('**') && line.endsWith('**')) {
                            return <p key={idx} className="font-bold text-slate-100 mt-2">{line.replace(/\*\*/g, '')}</p>;
                          }
                          return <p key={idx} className="mb-2 last:mb-0 text-slate-300 font-medium">{line}</p>;
                        })}
                      </div>
                    </motion.div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: ANÁLISE SETORIAL */}
          {activeTab === 'setores' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              {/* Sector Maturity Card */}
              <div className="p-6 bg-white rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div>
                  <h3 className="font-bold text-lg text-slate-800">Maturidade Média por Setor</h3>
                  <p className="text-slate-400 text-xs mt-0.5 font-sans">Nível de conformidade médio consolidado das empresas agrupadas por tipo de negócio</p>
                </div>

                <div className="h-80 w-full pt-2">
                  {sectorData.length > 0 ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={sectorData} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f8fafc" />
                        <XAxis dataKey="segment" stroke="#94a3b8" fontSize={11} tickLine={false} axisLine={false} />
                        <YAxis stroke="#94a3b8" fontSize={11} tickLine={false} axisLine={false} domain={[0, 100]} tickFormatter={(v) => `${v}%`} />
                        <Tooltip 
                          contentStyle={{ borderRadius: '1rem', border: '1px solid #f1f5f9', fontFamily: 'Inter, sans-serif' }}
                          formatter={(value: any) => [`${value}%`, 'Maturidade Média']}
                        />
                        <Bar dataKey="Maturidade" radius={[8, 8, 0, 0]} barSize={40}>
                          {sectorData.map((entry, index) => {
                            const colors = ['#10b981', '#6366f1', '#f59e0b', '#3b82f6', '#ec4899'];
                            return <Cell key={`cell-${index}`} fill={colors[index % colors.length]} />;
                          })}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="flex items-center justify-center h-full text-slate-400 text-xs italic">
                      Não há dados setoriais suficientes.
                    </div>
                  )}
                </div>
              </div>

              {/* Customers distribution by Sector */}
              <div className="p-6 bg-white rounded-3xl border border-slate-100 shadow-sm space-y-4">
                <div>
                  <h3 className="font-bold text-lg text-slate-800">Distribuição da Carteira por Segmento</h3>
                  <p className="text-slate-400 text-xs mt-0.5 font-sans">Número absoluto de clientes diagnosticados em cada setor do portfólio</p>
                </div>

                <div className="h-80 w-full pt-2">
                  {sectorData.length > 0 ? (
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={sectorData} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f8fafc" />
                        <XAxis dataKey="segment" stroke="#94a3b8" fontSize={11} tickLine={false} axisLine={false} />
                        <YAxis stroke="#94a3b8" fontSize={11} tickLine={false} axisLine={false} />
                        <Tooltip 
                          contentStyle={{ borderRadius: '1rem', border: '1px solid #f1f5f9', fontFamily: 'Inter, sans-serif' }}
                          formatter={(value: any) => [value, 'Clientes Ativos']}
                        />
                        <Bar dataKey="Clientes" fill="#475569" radius={[8, 8, 0, 0]} barSize={40} />
                      </BarChart>
                    </ResponsiveContainer>
                  ) : (
                    <div className="flex items-center justify-center h-full text-slate-400 text-xs italic">
                      Não há dados setoriais suficientes.
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: CLIENTES ATIVOS COM SPARKLINES */}
          {activeTab === 'clientes' && (
            <div className="p-1 bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
              <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h3 className="font-bold text-lg text-slate-800">Desempenho e Trajetória de Clientes</h3>
                  <p className="text-slate-400 text-xs mt-0.5 font-sans">Tabela de comparação consolidando a evolução entre o primeiro diagnóstico e o estado atual</p>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full text-left text-sm border-collapse text-slate-600">
                  <thead>
                    <tr className="bg-slate-50 text-slate-400 font-bold border-b border-slate-100">
                      <th className="px-6 py-4">Cliente / Empresa</th>
                      <th className="px-6 py-4">Segmento</th>
                      <th className="px-6 py-4 text-center">Nº Diagnósticos</th>
                      <th className="px-6 py-4 text-center">Score Inicial</th>
                      <th className="px-6 py-4 text-center">Tendência (12m)</th>
                      <th className="px-6 py-4 text-center">Score Atual</th>
                      <th className="px-6 py-4 text-center">Evolução Líquida</th>
                      <th className="px-6 py-4">Status de Impacto</th>
                      {onNavigateToDiagnosis && <th className="px-6 py-4 text-right">Ação</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {clientSummary.length === 0 ? (
                      <tr>
                        <td colSpan={onNavigateToDiagnosis ? 9 : 8} className="text-center py-12 text-slate-400 text-sm font-sans italic">
                          Nenhum cliente diagnosticado atende aos critérios dos filtros.
                        </td>
                      </tr>
                    ) : (
                      clientSummary.map(cli => {
                        const isPositive = cli.growth > 0;
                        return (
                          <tr key={cli.id} className="border-b border-slate-50 hover:bg-slate-50/40 transition-colors">
                            {/* Company Name */}
                            <td className="px-6 py-4">
                              <span className="font-bold text-slate-800 block text-sm">{cli.nome}</span>
                              <span className="text-[10px] text-slate-400 block font-sans mt-0.5">Último em: {cli.latestDate}</span>
                            </td>
                            {/* Sector */}
                            <td className="px-6 py-4 text-xs font-semibold uppercase text-slate-500 font-sans">
                              {cli.segmento}
                            </td>
                            {/* Diags count */}
                            <td className="px-6 py-4 text-center text-xs font-bold text-slate-600">
                              {cli.diagsCount}
                            </td>
                            {/* First Score */}
                            <td className="px-6 py-4 text-center">
                              <span className="text-xs font-semibold text-slate-700">{cli.firstScore}%</span>
                              <span className="block text-[9px] text-slate-400 font-sans">{cli.firstDate}</span>
                            </td>
                            {/* Sparkline (Trend) */}
                            <td className="px-6 py-4 text-center flex justify-center mt-1.5">
                              <Sparkline data={cli.sparklineScores} />
                            </td>
                            {/* Latest Score */}
                            <td className="px-6 py-4 text-center">
                              <span className="text-sm font-black text-slate-900">{cli.latestScore}%</span>
                              <span className="block text-[9px] text-slate-400 font-sans">{cli.latestDate}</span>
                            </td>
                            {/* Abs growth */}
                            <td className="px-6 py-4 text-center">
                              <span className={cn(
                                "inline-flex items-center gap-1 text-xs font-black px-2 py-0.5 rounded-full",
                                cli.growth > 15 ? "bg-emerald-100 text-emerald-800" :
                                cli.growth > 0 ? "bg-emerald-50 text-emerald-700" :
                                cli.growth < 0 ? "bg-rose-50 text-rose-700" : "bg-slate-100 text-slate-600"
                              )}>
                                {isPositive ? `+${cli.growth}` : cli.growth}%
                              </span>
                            </td>
                            {/* Impact status label */}
                            <td className="px-6 py-4 text-xs">
                              <span className={cn(
                                "font-bold",
                                cli.status === 'Crescimento Forte' ? "text-emerald-600" :
                                cli.status === 'Evolução Gradual' ? "text-emerald-500" :
                                cli.status === 'Alerta queda' ? "text-rose-500 animate-pulse" : "text-slate-400"
                              )}>
                                {cli.status}
                              </span>
                            </td>
                            {/* Jump to specific diagnosis */}
                            {onNavigateToDiagnosis && (
                              <td className="px-6 py-4 text-right">
                                <button
                                  onClick={() => onNavigateToDiagnosis(cli.diagnosticoOriginal)}
                                  className="p-1 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors cursor-pointer"
                                  title="Ver diagnóstico atual"
                                >
                                  <ChevronRight size={18} />
                                </button>
                              </td>
                            )}
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
};
