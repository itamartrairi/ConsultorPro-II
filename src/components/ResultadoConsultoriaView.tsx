import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  Award,
  CheckCircle2,
  AlertCircle,
  FileText,
  Download,
  Building2,
  Calendar,
  Save,
  Plus,
  Trash2,
  TrendingUp,
  UserCheck,
  Star,
  Layers,
  Sparkles,
  HelpCircle,
  Clock,
  Briefcase,
  ChevronRight,
  ShieldCheck,
  RefreshCw,
  Search,
  Sliders,
  Check,
  ArrowRight,
  Target,
  BarChart3,
  ThumbsUp,
  User,
  Phone,
  Mail,
  FileSpreadsheet,
  Printer
} from 'lucide-react';
import {
  collection,
  query,
  where,
  onSnapshot,
  addDoc,
  updateDoc,
  doc,
  deleteDoc,
  getDocs,
  serverTimestamp
} from 'firebase/firestore';
import { db } from '../firebase';
import { Button } from './Button';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { sebraeLogoBase64 } from '../sebraeLogo';

export interface ItemProblemaSolucao {
  id: string;
  area: string;
  problema: string;
  solucaoImplementada: string;
  status: 'Concluído' | 'Em Andamento' | 'Parcial' | 'Não Iniciado';
  notaEficacia: number; // 0 a 10
  evidenciaResultado: string;
}

export interface AnaliseResultadoNotas {
  eficaciaProblemas: number;       // 0 a 10: Resolução dos problemas diagnosticados
  impactoGestao: number;          // 0 a 10: Impacto na gestão, processos e rotinas
  impactoFinanceiroProd: number;   // 0 a 10: Ganhos financeiros / produtividade
  engajamentoEquipe: number;      // 0 a 10: Engajamento e adesão da equipe do cliente
  satisfacaoConsultoria: number;  // 0 a 10: Satisfação geral com a consultoria e metodologia
  satisfacaoConsultor: number;    // 0 a 10: Satisfação com a postura e atuação do consultor
}

export interface AnaliseResultado {
  id?: string;
  empresaId: string;
  diagnosticoId: string;
  nomeEmpresa: string;
  nomeProjeto: string;
  dataAnalise: string; // YYYY-MM-DD
  consultorNome: string;
  codigoSgf?: string;
  credenciadaNome?: string;
  cargaHoraria?: string;
  notas: AnaliseResultadoNotas;
  mediaGeral: number; // 0 a 10
  classificacao: string;
  problemasSolucoes: ItemProblemaSolucao[];
  parecerImpactoGestao: string;
  feedbackConsultoria: string;
  feedbackConsultor: string;
  ganhosPrincipais: string;
  recomendacoesFuturas: string;
  responsavelClienteNome?: string;
  responsavelClienteCargo?: string;
  ownerId?: string;
  createdAt?: any;
  updatedAt?: any;
}

interface Empresa {
  id: string;
  nome: string;
  razaoSocial?: string;
  nomeFantasia?: string;
  cnpj?: string;
  enderecoComercial?: string;
  telefoneFixo?: string;
  celular?: string;
  email?: string;
  representante?: string;
  cpfRepresentante?: string;
  tipoEmpresa?: string;
  ramoAtividade?: string;
  cafNumero?: string;
}

interface Diagnostico {
  id: string;
  empresaId: string;
  nome?: string;
  nomeProjeto?: string;
  nomeEmpresa?: string;
  dataDiagnostico?: any;
  cargaHoraria?: string;
  dadosConsultoria?: {
    razaoSocial?: string;
    cnpj?: string;
    consultor?: string;
    telefoneFixo?: string;
    celular?: string;
    email?: string;
    areaConsultoria?: string;
    codigoSgf?: string;
    periodoConsultoria?: string;
    tecnicoSebrae?: string;
    objetivo?: string;
    resultadosEsperados?: string;
    solucoesIndicadas?: string;
    cargaHoraria?: string;
  };
}

interface Resposta {
  id: string;
  diagnosticoId: string;
  problema: string;
  pergunta: string;
  area: string;
  resposta?: string;
  observacao?: string;
  score?: number;
}

interface TarefaPlanoAcao {
  id: string;
  diagnosticoId: string;
  empresaId: string;
  problema: string;
  area: string;
  solucaoSugerida: string;
  acoes: string;
  status: 'Pendente' | 'Em Andamento' | 'Concluído';
  responsavel?: string;
}

interface ResultadoConsultoriaViewProps {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas?: Resposta[];
  tarefas?: TarefaPlanoAcao[];
  user?: any;
  selectedEmpresa?: Empresa | null;
  selectedDiagnostico?: Diagnostico | null;
  setView?: (view: any) => void;
  customLogo?: string | null;
  customConsultoraLogo?: string | null;
}

const NOTA_DESCRITORES: Record<number, { label: string; color: string }> = {
  10: { label: 'Excelente / Superou expectativas', color: 'text-emerald-700 bg-emerald-50 border-emerald-200' },
  9: { label: 'Ótimo / Alto impacto alcançado', color: 'text-emerald-600 bg-emerald-50 border-emerald-200' },
  8: { label: 'Muito Bom / Metas atingidas', color: 'text-teal-600 bg-teal-50 border-teal-200' },
  7: { label: 'Bom / Resultados consistentes', color: 'text-blue-600 bg-blue-50 border-blue-200' },
  6: { label: 'Satisfatório / Parcialmente atingido', color: 'text-amber-600 bg-amber-50 border-amber-200' },
  5: { label: 'Regular / Exige melhorias', color: 'text-amber-700 bg-amber-50 border-amber-200' },
  4: { label: 'Insuficiente / Pouca evolução', color: 'text-orange-600 bg-orange-50 border-orange-200' },
  3: { label: 'Baixo / Dificuldades na execução', color: 'text-rose-600 bg-rose-50 border-rose-200' },
  2: { label: 'Muito Baixo / Sem adesão', color: 'text-rose-700 bg-rose-50 border-rose-200' },
  1: { label: 'Crítico / Sem resultados', color: 'text-rose-800 bg-rose-50 border-rose-200' },
  0: { label: 'Não Implementado / Nulo', color: 'text-slate-600 bg-slate-100 border-slate-200' }
};

export const ResultadoConsultoriaView: React.FC<ResultadoConsultoriaViewProps> = ({
  empresas,
  diagnosticos,
  respostas = [],
  tarefas = [],
  user,
  selectedEmpresa: initialEmpresa,
  selectedDiagnostico: initialDiagnostico,
  customLogo,
  customConsultoraLogo
}) => {
  // State for Selection
  const [selectedEmpresaId, setSelectedEmpresaId] = useState<string>(
    initialEmpresa?.id || (empresas.length > 0 ? empresas[0].id : '')
  );
  const [selectedDiagId, setSelectedDiagId] = useState<string>(
    initialDiagnostico?.id || ''
  );

  // Filtered lists
  const currentEmpresa = useMemo(() => {
    return empresas.find(e => e.id === selectedEmpresaId) || null;
  }, [empresas, selectedEmpresaId]);

  const availableDiagnosticos = useMemo(() => {
    if (!selectedEmpresaId) return [];
    return diagnosticos.filter(d => d.empresaId === selectedEmpresaId);
  }, [diagnosticos, selectedEmpresaId]);

  const currentDiagnostico = useMemo(() => {
    return availableDiagnosticos.find(d => d.id === selectedDiagId) || null;
  }, [availableDiagnosticos, selectedDiagId]);

  // Update selectedDiagId if empresa changes
  useEffect(() => {
    if (availableDiagnosticos.length > 0 && (!selectedDiagId || !availableDiagnosticos.some(d => d.id === selectedDiagId))) {
      setSelectedDiagId(availableDiagnosticos[0].id);
    } else if (availableDiagnosticos.length === 0) {
      setSelectedDiagId('');
    }
  }, [availableDiagnosticos, selectedDiagId]);

  // Saved Analyses list from DB / Local
  const [savedAnalyses, setSavedAnalyses] = useState<AnaliseResultado[]>([]);
  const [loadingSaved, setLoadingSaved] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isExportingPdf, setIsExportingPdf] = useState(false);
  const [activeSubTab, setActiveSubTab] = useState<'form' | 'historico' | 'relatorio'>('form');
  const [saveSuccessMessage, setSaveSuccessMessage] = useState<string | null>(null);

  // Form State
  const [dataAnalise, setDataAnalise] = useState<string>(
    new Date().toISOString().split('T')[0]
  );
  const [consultorNome, setConsultorNome] = useState<string>(
    user?.displayName || currentDiagnostico?.dadosConsultoria?.consultor || ''
  );
  const [codigoSgf, setCodigoSgf] = useState<string>(
    currentDiagnostico?.dadosConsultoria?.codigoSgf || ''
  );
  const [credenciadaNome, setCredenciadaNome] = useState<string>(
    currentDiagnostico?.dadosConsultoria?.razaoSocial || ''
  );
  const [cargaHoraria, setCargaHoraria] = useState<string>(
    currentDiagnostico?.dadosConsultoria?.cargaHoraria || currentDiagnostico?.cargaHoraria || '20h'
  );
  const [responsavelClienteNome, setResponsavelClienteNome] = useState<string>(
    currentEmpresa?.representante || ''
  );
  const [responsavelClienteCargo, setResponsavelClienteCargo] = useState<string>('Diretor / Proprietário');

  // Scores (0 to 10)
  const [notas, setNotas] = useState<AnaliseResultadoNotas>({
    eficaciaProblemas: 9,
    impactoGestao: 9,
    impactoFinanceiroProd: 8,
    engajamentoEquipe: 8,
    satisfacaoConsultoria: 10,
    satisfacaoConsultor: 10
  });

  // Problem-Solution Matrix
  const [problemasSolucoes, setProblemasSolucoes] = useState<ItemProblemaSolucao[]>([]);

  // Qualitative Analysis
  const [parecerImpactoGestao, setParecerImpactoGestao] = useState<string>(
    'A consultoria proporcionou a estruturação de rotinas gerenciais claras, implantação de indicadores-chave de desempenho e melhoria substancial na previsibilidade e controle dos processos operacionais e financeiros da empresa.'
  );
  const [feedbackConsultoria, setFeedbackConsultoria] = useState<string>(
    'O cliente manifestou total satisfação com o cronograma executado, a clareza das ferramentas aplicadas e a aplicabilidade prática das soluções.'
  );
  const [feedbackConsultor, setFeedbackConsultor] = useState<string>(
    'O consultor demonstrou pontualidade, excelente didática, domínio técnico aprofundado dos processos e conduziu as reuniões com empatia e assertividade.'
  );
  const [ganhosPrincipais, setGanhosPrincipais] = useState<string>(
    '• Implementação de controle rigoroso de fluxo de caixa e margens.\n• Padronização de rotinas operacionais e redução de retrabalho.\n• Definição de metas e alinhamento da equipe aos objetivos estratégicos.'
  );
  const [recomendacoesFuturas, setRecomendacoesFuturas] = useState<string>(
    'Recomenda-se a realização de reuniões quinzenais de acompanhamento dos indicadores implantados e a continuidade no ciclo de consultoria avançada em Marketing e Vendas.'
  );

  // Active saved analysis ID if editing
  const [currentAnalysisId, setCurrentAnalysisId] = useState<string | null>(null);

  // Load Saved Analyses from Firestore / LocalStorage
  useEffect(() => {
    let unsubscribe = () => {};
    try {
      const q = query(
        collection(db, 'analises_resultado'),
        where('ownerId', '==', user?.uid || 'local_user')
      );
      unsubscribe = onSnapshot(q, (snapshot) => {
        const list: AnaliseResultado[] = [];
        snapshot.forEach((docSnap) => {
          list.push({ id: docSnap.id, ...(docSnap.data() as any) });
        });
        setSavedAnalyses(list);
      }, (err) => {
        console.warn('Firestore snapshot error for analises_resultado, falling back to local:', err);
        loadLocalAnalyses();
      });
    } catch (e) {
      loadLocalAnalyses();
    }

    return () => unsubscribe();
  }, [user]);

  const loadLocalAnalyses = () => {
    try {
      const local = localStorage.getItem('analises_resultado_local');
      if (local) {
        setSavedAnalyses(JSON.parse(local));
      }
    } catch {}
  };

  // Sync initial consultant info when selecting a diagnosis
  useEffect(() => {
    if (currentDiagnostico) {
      if (currentDiagnostico.dadosConsultoria?.consultor) {
        setConsultorNome(currentDiagnostico.dadosConsultoria.consultor);
      }
      if (currentDiagnostico.dadosConsultoria?.codigoSgf) {
        setCodigoSgf(currentDiagnostico.dadosConsultoria.codigoSgf);
      }
      if (currentDiagnostico.dadosConsultoria?.razaoSocial) {
        setCredenciadaNome(currentDiagnostico.dadosConsultoria.razaoSocial);
      }
      if (currentDiagnostico.dadosConsultoria?.cargaHoraria || currentDiagnostico.cargaHoraria) {
        setCargaHoraria(currentDiagnostico.dadosConsultoria?.cargaHoraria || currentDiagnostico.cargaHoraria || '20h');
      }
    }
    if (currentEmpresa) {
      if (currentEmpresa.representante) {
        setResponsavelClienteNome(currentEmpresa.representante);
      }
    }
  }, [currentDiagnostico, currentEmpresa]);

  // Check if there is already a saved analysis for the selected diagnosis
  useEffect(() => {
    if (selectedDiagId && savedAnalyses.length > 0) {
      const found = savedAnalyses.find(a => a.diagnosticoId === selectedDiagId);
      if (found) {
        loadAnalysisIntoForm(found);
      } else {
        // Auto populate matrix from existing diagnosis & action plan
        autoPopulateFromDiagnosis();
      }
    } else if (selectedDiagId) {
      autoPopulateFromDiagnosis();
    }
  }, [selectedDiagId, savedAnalyses]);

  // Auto populate problem-solution matrix from diagnosis answers and tasks
  const autoPopulateFromDiagnosis = () => {
    if (!selectedDiagId) return;

    // 1. Get negative/critical answers from diagnostic
    const diagRespostas = respostas.filter(
      r => r.diagnosticoId === selectedDiagId && (r.resposta === 'Não' || r.resposta === 'Parcial' || (r.score !== undefined && r.score < 100))
    );

    // 2. Get tasks from action plan
    const diagTarefas = tarefas.filter(t => t.diagnosticoId === selectedDiagId);

    const items: ItemProblemaSolucao[] = [];

    if (diagTarefas.length > 0) {
      diagTarefas.forEach((t, idx) => {
        items.push({
          id: 'item-' + (idx + 1) + '-' + Date.now(),
          area: t.area || 'Geral',
          problema: t.problema || 'Gargalo operacional identificado',
          solucaoImplementada: t.solucaoSugerida || t.acoes || 'Solução estruturada no plano de ação',
          status: t.status === 'Concluído' ? 'Concluído' : t.status === 'Em Andamento' ? 'Em Andamento' : 'Parcial',
          notaEficacia: t.status === 'Concluído' ? 10 : 8,
          evidenciaResultado: 'Plano de ação executado com acompanhamento do consultor.'
        });
      });
    } else if (diagRespostas.length > 0) {
      diagRespostas.slice(0, 8).forEach((r, idx) => {
        items.push({
          id: 'item-resp-' + (idx + 1) + '-' + Date.now(),
          area: r.area || 'Gestão',
          problema: r.problema || r.pergunta || 'Deficiência diagnosticada',
          solucaoImplementada: 'Estruturação de procedimentos e boas práticas gerenciais.',
          status: 'Concluído',
          notaEficacia: 9,
          evidenciaResultado: 'Processo padronizado e validado junto à direção da empresa.'
        });
      });
    } else {
      // Default initial templates
      items.push(
        {
          id: 'item-default-1',
          area: 'Financeiro',
          problema: 'Ausência de controle diário de fluxo de caixa e separação das contas pessoais e empresariais.',
          solucaoImplementada: 'Implantação de planilha gerencial de fluxo de caixa diário e política de pró-labore.',
          status: 'Concluído',
          notaEficacia: 10,
          evidenciaResultado: 'Relatório diário de conciliação ativo e controle de saldo projetado.'
        },
        {
          id: 'item-default-2',
          area: 'Operacional',
          problema: 'Processos sem padronização, gerando retrabalho e dependência exclusiva do proprietário.',
          solucaoImplementada: 'Criação de Procedimentos Operacionais Padrão (POPs) para as 3 atividades críticas.',
          status: 'Concluído',
          notaEficacia: 9,
          evidenciaResultado: 'Checklists de execução impressos e seguidos pela equipe.'
        },
        {
          id: 'item-default-3',
          area: 'Vendas & Marketing',
          problema: 'Falta de mensuração de canais de aquisição de clientes e ticket médio.',
          solucaoImplementada: 'Estruturação de funil de vendas simples e registro de contatos comerciais.',
          status: 'Em Andamento',
          notaEficacia: 8,
          evidenciaResultado: 'Aumento de 15% nas conversões no último mês do projeto.'
        }
      );
    }

    setProblemasSolucoes(items);
    setCurrentAnalysisId(null);
  };

  const loadAnalysisIntoForm = (analysis: AnaliseResultado) => {
    setCurrentAnalysisId(analysis.id || null);
    setSelectedEmpresaId(analysis.empresaId);
    setSelectedDiagId(analysis.diagnosticoId);
    setDataAnalise(analysis.dataAnalise || new Date().toISOString().split('T')[0]);
    setConsultorNome(analysis.consultorNome || '');
    setCodigoSgf(analysis.codigoSgf || '');
    setCredenciadaNome(analysis.credenciadaNome || '');
    setCargaHoraria(analysis.cargaHoraria || '20h');
    setResponsavelClienteNome(analysis.responsavelClienteNome || '');
    setResponsavelClienteCargo(analysis.responsavelClienteCargo || 'Diretor / Proprietário');
    if (analysis.notas) {
      setNotas(analysis.notas);
    }
    if (analysis.problemasSolucoes && analysis.problemasSolucoes.length > 0) {
      setProblemasSolucoes(analysis.problemasSolucoes);
    }
    setParecerImpactoGestao(analysis.parecerImpactoGestao || '');
    setFeedbackConsultoria(analysis.feedbackConsultoria || '');
    setFeedbackConsultor(analysis.feedbackConsultor || '');
    setGanhosPrincipais(analysis.ganhosPrincipais || '');
    setRecomendacoesFuturas(analysis.recomendacoesFuturas || '');
  };

  // Calculate Global Average Score (0 to 10)
  const mediaGeralCalculada = useMemo(() => {
    const vals = [
      notas.eficaciaProblemas,
      notas.impactoGestao,
      notas.impactoFinanceiroProd,
      notas.engajamentoEquipe,
      notas.satisfacaoConsultoria,
      notas.satisfacaoConsultor
    ];
    const sum = vals.reduce((acc, v) => acc + (Number(v) || 0), 0);
    return Number((sum / vals.length).toFixed(1));
  }, [notas]);

  // Classification based on score
  const classificacaoGeral = useMemo(() => {
    if (mediaGeralCalculada >= 9.0) return { label: 'Excelente / Alto Impacto', badge: 'bg-emerald-100 text-emerald-800 border-emerald-300', desc: 'A consultoria transformou estruturalmente a gestão da empresa, superando as metas propostas.' };
    if (mediaGeralCalculada >= 7.5) return { label: 'Muito Bom / Metas Atingidas', badge: 'bg-teal-100 text-teal-800 border-teal-300', desc: 'As soluções foram implantadas com sucesso e geraram impacto relevante e sustentável.' };
    if (mediaGeralCalculada >= 6.0) return { label: 'Satisfatório / Impacto Moderado', badge: 'bg-amber-100 text-amber-800 border-amber-300', desc: 'Resultados parciais alcançados; necessita acompanhamento contínuo para consolidação.' };
    return { label: 'Regular / Requer Ajustes', badge: 'bg-rose-100 text-rose-800 border-rose-300', desc: 'Adesão abaixo do esperado; recomenda-se revisão do plano de ação e replanejamento.' };
  }, [mediaGeralCalculada]);

  // Handle Note Change (0 to 10)
  const handleNotaChange = (field: keyof AnaliseResultadoNotas, val: number) => {
    const clamped = Math.max(0, Math.min(10, val));
    setNotas(prev => ({ ...prev, [field]: clamped }));
  };

  // Problem-Solution Matrix Row Helpers
  const handleAddRow = () => {
    const newRow: ItemProblemaSolucao = {
      id: 'row-' + Date.now(),
      area: 'Geral',
      problema: '',
      solucaoImplementada: '',
      status: 'Concluído',
      notaEficacia: 10,
      evidenciaResultado: ''
    };
    setProblemasSolucoes(prev => [...prev, newRow]);
  };

  const handleUpdateRow = (id: string, field: keyof ItemProblemaSolucao, value: any) => {
    setProblemasSolucoes(prev => prev.map(row => {
      if (row.id === id) {
        return { ...row, [field]: value };
      }
      return row;
    }));
  };

  const handleDeleteRow = (id: string) => {
    setProblemasSolucoes(prev => prev.filter(row => row.id !== id));
  };

  // Save Analysis to Firestore & LocalStorage
  const handleSaveAnalysis = async () => {
    if (!selectedEmpresaId) {
      alert('Por favor, selecione um Cliente antes de salvar a análise.');
      return;
    }
    if (!selectedDiagId) {
      alert('Por favor, selecione um Diagnóstico/Projeto antes de salvar.');
      return;
    }

    setIsSaving(true);
    const payload: AnaliseResultado = {
      empresaId: selectedEmpresaId,
      diagnosticoId: selectedDiagId,
      nomeEmpresa: currentEmpresa?.nome || currentDiagnostico?.nomeEmpresa || 'Cliente',
      nomeProjeto: currentDiagnostico?.nomeProjeto || currentDiagnostico?.nome || 'Consultoria Gerencial',
      dataAnalise,
      consultorNome: consultorNome || user?.displayName || 'Consultor',
      codigoSgf,
      credenciadaNome,
      cargaHoraria,
      responsavelClienteNome,
      responsavelClienteCargo,
      notas,
      mediaGeral: mediaGeralCalculada,
      classificacao: classificacaoGeral.label,
      problemasSolucoes,
      parecerImpactoGestao,
      feedbackConsultoria,
      feedbackConsultor,
      ganhosPrincipais,
      recomendacoesFuturas,
      ownerId: user?.uid || 'local_user',
      updatedAt: new Date().toISOString()
    };

    try {
      if (currentAnalysisId) {
        // Update existing
        try {
          await updateDoc(doc(db, 'analises_resultado', currentAnalysisId), {
            ...payload,
            updatedAt: serverTimestamp()
          });
        } catch (e) {
          console.warn('Could not update firestore, saving local backup:', e);
        }
      } else {
        // Create new
        try {
          const docRef = await addDoc(collection(db, 'analises_resultado'), {
            ...payload,
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp()
          });
          setCurrentAnalysisId(docRef.id);
          payload.id = docRef.id;
        } catch (e) {
          console.warn('Could not add to firestore, saving local backup:', e);
          const localId = 'local-' + Date.now();
          payload.id = localId;
          setCurrentAnalysisId(localId);
        }
      }

      // Mirror to local storage
      const existing = [...savedAnalyses];
      const idx = existing.findIndex(a => a.id === currentAnalysisId || (a.diagnosticoId === selectedDiagId && a.empresaId === selectedEmpresaId));
      if (idx >= 0) {
        existing[idx] = { ...payload, id: existing[idx].id || currentAnalysisId || 'local-' + Date.now() };
      } else {
        existing.unshift(payload);
      }
      setSavedAnalyses(existing);
      try {
        localStorage.setItem('analises_resultado_local', JSON.stringify(existing));
      } catch {}

      setSaveSuccessMessage('Análise de Resultado da Consultoria salva com sucesso!');
      setTimeout(() => setSaveSuccessMessage(null), 4000);
    } catch (err: any) {
      alert('Erro ao salvar análise: ' + (err?.message || err));
    } finally {
      setIsSaving(false);
    }
  };

  // Delete an analysis
  const handleDeleteSaved = async (id: string) => {
    if (!confirm('Tem certeza que deseja excluir esta análise de resultado?')) return;
    try {
      try {
        await deleteDoc(doc(db, 'analises_resultado', id));
      } catch {}
      const updated = savedAnalyses.filter(a => a.id !== id);
      setSavedAnalyses(updated);
      try {
        localStorage.setItem('analises_resultado_local', JSON.stringify(updated));
      } catch {}
      if (currentAnalysisId === id) {
        setCurrentAnalysisId(null);
        autoPopulateFromDiagnosis();
      }
    } catch (e: any) {
      alert('Erro ao excluir: ' + (e?.message || e));
    }
  };

  // Export PDF Report with High Resolution and Consultancy Header
  const handleExportPDF = async () => {
    const reportElement = document.getElementById('consultancy-result-pdf-report');
    if (!reportElement) {
      alert('Elemento do relatório não encontrado.');
      return;
    }

    setIsExportingPdf(true);
    try {
      const canvas = await html2canvas(reportElement, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: '#ffffff'
      });

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgWidth = 190;
      const pageHeight = 285;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 10;

      pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;

      while (heightLeft > 0) {
        position = heightLeft - imgHeight + 10;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
      }

      const clientNameClean = (currentEmpresa?.nome || 'Cliente').replace(/[^a-zA-Z0-9]/g, '_');
      pdf.save(`Relatorio_Resultado_Consultoria_${clientNameClean}.pdf`);
    } catch (err: any) {
      console.error('PDF Export Error:', err);
      alert('Erro ao gerar relatório em PDF: ' + (err?.message || err));
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Active logo determination
  const activeLogoSrc = customLogo || customConsultoraLogo || sebraeLogoBase64;

  return (
    <div className="space-y-6 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
      {/* Top Header */}
      <div className="flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-200 pb-5 gap-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-200">
              <Award size={13} className="text-indigo-600" />
              Módulo de Avaliação & Impacto
            </span>
            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-700">
              Escala 0 a 10 Pontos
            </span>
          </div>
          <h1 className="text-2xl sm:text-3xl font-bold text-slate-900 tracking-tight flex items-center gap-3">
            <BarChart3 className="h-7 w-7 text-indigo-600" />
            Análise do Resultado da Consultoria
          </h1>
          <p className="mt-1 text-sm text-slate-500 max-w-3xl">
            Avalie os problemas diagnosticados, a efetividade das soluções implementadas, o impacto real na gestão e a satisfação do cliente e com o consultor.
          </p>
        </div>

        {/* Action Buttons & SubTabs */}
        <div className="flex items-center gap-3 flex-wrap">
          <div className="inline-flex rounded-xl bg-slate-100 p-1 border border-slate-200">
            <button
              onClick={() => setActiveSubTab('form')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition-all flex items-center gap-2 ${
                activeSubTab === 'form'
                  ? 'bg-white text-indigo-700 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Sliders size={14} />
              Avaliação (Formulário)
            </button>
            <button
              onClick={() => setActiveSubTab('relatorio')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition-all flex items-center gap-2 ${
                activeSubTab === 'relatorio'
                  ? 'bg-white text-indigo-700 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileText size={14} />
              Laudo / Relatório PDF
            </button>
            <button
              onClick={() => setActiveSubTab('historico')}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition-all flex items-center gap-2 ${
                activeSubTab === 'historico'
                  ? 'bg-white text-indigo-700 shadow-sm'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Clock size={14} />
              Histórico ({savedAnalyses.length})
            </button>
          </div>

          <Button
            onClick={handleExportPDF}
            disabled={isExportingPdf || !selectedDiagId}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold text-xs px-4 py-2 rounded-xl shadow-sm flex items-center gap-2"
          >
            {isExportingPdf ? (
              <RefreshCw className="h-4 w-4 animate-spin" />
            ) : (
              <Download className="h-4 w-4" />
            )}
            Gerar Laudo PDF
          </Button>
        </div>
      </div>

      {/* Success Notification */}
      {saveSuccessMessage && (
        <motion.div
          initial={{ opacity: 0, y: -10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0 }}
          className="p-4 bg-emerald-50 border border-emerald-200 rounded-xl flex items-center justify-between text-emerald-800 text-sm font-medium shadow-sm"
        >
          <div className="flex items-center gap-2.5">
            <CheckCircle2 className="h-5 w-5 text-emerald-600 flex-shrink-0" />
            <span>{saveSuccessMessage}</span>
          </div>
          <button
            onClick={() => setSaveSuccessMessage(null)}
            className="text-emerald-600 hover:text-emerald-800 text-xs font-bold"
          >
            Fechar
          </button>
        </motion.div>
      )}

      {/* Selectors Bar: Empresa & Diagnóstico */}
      <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          {/* Select Empresa */}
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Building2 size={14} className="text-indigo-600" />
              1. Selecionar Cliente
            </label>
            <select
              value={selectedEmpresaId}
              onChange={(e) => {
                setSelectedEmpresaId(e.target.value);
              }}
              className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none"
            >
              {empresas.length === 0 ? (
                <option value="">Nenhum cliente cadastrado</option>
              ) : (
                empresas.map((emp) => (
                  <option key={emp.id} value={emp.id}>
                    {emp.nome} {emp.cnpj ? `(${emp.cnpj})` : ''}
                  </option>
                ))
              )}
            </select>
          </div>

          {/* Select Diagnóstico */}
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Briefcase size={14} className="text-indigo-600" />
              2. Projeto / Diagnóstico
            </label>
            <select
              value={selectedDiagId}
              onChange={(e) => setSelectedDiagId(e.target.value)}
              disabled={availableDiagnosticos.length === 0}
              className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none disabled:opacity-50"
            >
              {availableDiagnosticos.length === 0 ? (
                <option value="">Nenhum diagnóstico para este cliente</option>
              ) : (
                availableDiagnosticos.map((diag) => (
                  <option key={diag.id} value={diag.id}>
                    {diag.nomeProjeto || diag.nome || `Diagnóstico (${diag.dataDiagnostico ? (typeof diag.dataDiagnostico === 'string' ? diag.dataDiagnostico : 'Data') : 'Recente'})`}
                  </option>
                ))
              )}
            </select>
          </div>

          {/* Data da Análise */}
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Calendar size={14} className="text-indigo-600" />
              Data da Avaliação
            </label>
            <input
              type="date"
              value={dataAnalise}
              onChange={(e) => setDataAnalise(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none"
            />
          </div>

          {/* Carga Horária */}
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1.5 flex items-center gap-1.5">
              <Clock size={14} className="text-indigo-600" />
              Carga Horária
            </label>
            <input
              type="text"
              placeholder="Ex: 20h, 30 horas"
              value={cargaHoraria}
              onChange={(e) => setCargaHoraria(e.target.value)}
              className="w-full bg-slate-50 border border-slate-300 rounded-xl px-3.5 py-2.5 text-sm font-medium text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:outline-none"
            />
          </div>
        </div>

        {/* Quick Diagnostic Metadata Banner */}
        {currentEmpresa && (
          <div className="pt-3 border-t border-slate-100 flex flex-wrap items-center justify-between gap-3 text-xs text-slate-600">
            <div className="flex items-center gap-4 flex-wrap">
              <span className="font-semibold text-slate-900">
                Razão Social: <span className="font-normal text-slate-600">{currentEmpresa.razaoSocial || currentEmpresa.nome}</span>
              </span>
              {currentEmpresa.cnpj && (
                <span>CNPJ: <strong className="text-slate-800">{currentEmpresa.cnpj}</strong></span>
              )}
              {currentEmpresa.representante && (
                <span>Responsável: <strong className="text-slate-800">{currentEmpresa.representante}</strong></span>
              )}
              {codigoSgf && (
                <span className="px-2 py-0.5 bg-indigo-50 text-indigo-700 rounded-md font-mono font-semibold">
                  SGF: {codigoSgf}
                </span>
              )}
            </div>

            <button
              type="button"
              onClick={autoPopulateFromDiagnosis}
              className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 transition-colors"
              title="Recarrega os problemas e tarefas do plano de ação"
            >
              <RefreshCw size={12} />
              Recarregar Dados do Diagnóstico
            </button>
          </div>
        )}
      </div>

      {/* Main View Mode Selector */}
      {activeSubTab === 'form' && (
        <div className="space-y-8">
          {/* 1. SEÇÃO DE NOTAS DE 0 A 10 PONTOS (INDICADORES CHAVE) */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
            <div className="bg-gradient-to-r from-indigo-900 via-indigo-800 to-slate-900 px-6 py-5 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <Star className="h-5 w-5 text-amber-400 fill-amber-400" />
                  <h2 className="text-lg font-bold">1. Sistema de Pontuação dos Resultados (Escala de 0 a 10)</h2>
                </div>
                <p className="text-xs text-indigo-200 mt-1">
                  Atribua notas de 0 (nulo/não atingido) a 10 (excelente/superou expectativas) para cada dimensão crítica da consultoria.
                </p>
              </div>

              {/* Global Score Pill */}
              <div className="bg-white/10 backdrop-blur-md px-5 py-2.5 rounded-2xl border border-white/20 flex items-center gap-3">
                <div className="text-right">
                  <div className="text-[10px] uppercase tracking-widest text-indigo-200 font-bold">Índice Geral</div>
                  <div className="text-2xl font-black text-amber-300">{mediaGeralCalculada.toFixed(1)} <span className="text-xs font-normal text-white/80">/ 10.0</span></div>
                </div>
                <div className="h-9 w-px bg-white/20" />
                <div className="max-w-[150px]">
                  <span className={`inline-block px-2.5 py-0.5 rounded-full text-[11px] font-bold ${classificacaoGeral.badge}`}>
                    {classificacaoGeral.label}
                  </span>
                </div>
              </div>
            </div>

            <div className="p-6">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {/* Dimensão 1: Eficácia nos Problemas */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <Target size={15} className="text-indigo-600" />
                      1. Resolução dos Problemas
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.eficaciaProblemas}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Efetividade na eliminação ou mitigação das dores e gargalos identificados no diagnóstico inicial.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.eficaciaProblemas}
                    onChange={(e) => handleNotaChange('eficaciaProblemas', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Nenhum)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.eficaciaProblemas]?.label}</span>
                    <span>10 (Total)</span>
                  </div>
                </div>

                {/* Dimensão 2: Impacto na Gestão */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <TrendingUp size={15} className="text-indigo-600" />
                      2. Impacto na Gestão
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.impactoGestao}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Evolução estrutural em processos, governança, rotinas de controle e tomadas de decisão.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.impactoGestao}
                    onChange={(e) => handleNotaChange('impactoGestao', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Nulo)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.impactoGestao]?.label}</span>
                    <span>10 (Alto)</span>
                  </div>
                </div>

                {/* Dimensão 3: Impacto Financeiro / Produtividade */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <BarChart3 size={15} className="text-indigo-600" />
                      3. Finanças & Produtividade
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.impactoFinanceiroProd}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Ganhos palpáveis em margens, redução de desperdícios, fluxo de caixa e agilidade operacional.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.impactoFinanceiroProd}
                    onChange={(e) => handleNotaChange('impactoFinanceiroProd', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Sem Ganho)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.impactoFinanceiroProd]?.label}</span>
                    <span>10 (Excelente)</span>
                  </div>
                </div>

                {/* Dimensão 4: Engajamento da Equipe */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <UserCheck size={15} className="text-indigo-600" />
                      4. Adesão & Equipe
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.engajamentoEquipe}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Nível de comprometimento, assiduidade e execução das tarefas pela equipe do cliente.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.engajamentoEquipe}
                    onChange={(e) => handleNotaChange('engajamentoEquipe', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Resistência)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.engajamentoEquipe]?.label}</span>
                    <span>10 (Total)</span>
                  </div>
                </div>

                {/* Dimensão 5: Satisfação com a Consultoria */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <ThumbsUp size={15} className="text-indigo-600" />
                      5. Satisfação com a Consultoria
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.satisfacaoConsultoria}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Avaliação do cliente quanto à metodologia, cronograma, ferramentas e entregáveis.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.satisfacaoConsultoria}
                    onChange={(e) => handleNotaChange('satisfacaoConsultoria', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Insatisfeito)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.satisfacaoConsultoria]?.label}</span>
                    <span>10 (Plena)</span>
                  </div>
                </div>

                {/* Dimensão 6: Satisfação com o Consultor */}
                <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 space-y-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
                      <User size={15} className="text-indigo-600" />
                      6. Satisfação com o Consultor
                    </span>
                    <span className="text-xl font-black text-indigo-700 bg-white px-3 py-0.5 rounded-lg border border-indigo-200">
                      {notas.satisfacaoConsultor}
                    </span>
                  </div>
                  <p className="text-xs text-slate-500">
                    Postura profissional, didática, pontualidade, domínio técnico e relacionamento interpessoal.
                  </p>
                  <input
                    type="range"
                    min={0}
                    max={10}
                    step={1}
                    value={notas.satisfacaoConsultor}
                    onChange={(e) => handleNotaChange('satisfacaoConsultor', parseInt(e.target.value))}
                    className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                  />
                  <div className="flex justify-between text-[10px] font-bold text-slate-400">
                    <span>0 (Insatisfatório)</span>
                    <span className="text-indigo-700 font-semibold">{NOTA_DESCRITORES[notas.satisfacaoConsultor]?.label}</span>
                    <span>10 (Exemplar)</span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* 2. MATRIZ DE PROBLEMAS DIAGNOSTICADOS VS SOLUÇÕES IMPLEMENTADAS */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-100 pb-4">
              <div>
                <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                  <Layers className="h-5 w-5 text-indigo-600" />
                  2. Matriz de Problemas Diagnosticados vs. Soluções Implementadas
                </h2>
                <p className="text-xs text-slate-500 mt-0.5">
                  Confronte cada gargalo encontrado no diagnóstico inicial com a solução prática executada e sua eficácia (0 a 10).
                </p>
              </div>

              <Button
                type="button"
                onClick={handleAddRow}
                variant="outline"
                className="text-xs font-semibold text-indigo-700 border-indigo-300 hover:bg-indigo-50 flex items-center gap-1.5"
              >
                <Plus size={14} /> Adicionar Item
              </Button>
            </div>

            {/* Matrix Table */}
            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-700 font-bold border-b border-slate-200">
                    <th className="py-3 px-3 w-28">Área</th>
                    <th className="py-3 px-3">Problema Diagnosticado</th>
                    <th className="py-3 px-3">Solução Implementada</th>
                    <th className="py-3 px-3 w-36">Status</th>
                    <th className="py-3 px-3 w-24 text-center">Eficácia (0-10)</th>
                    <th className="py-3 px-3">Evidência / Resultado</th>
                    <th className="py-3 px-2 w-10 text-center">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {problemasSolucoes.length === 0 ? (
                    <tr>
                      <td colSpan={7} className="py-8 text-center text-slate-400">
                        Nenhum problema/solução adicionado. Clique no botão acima para adicionar ou recarregar do diagnóstico.
                      </td>
                    </tr>
                  ) : (
                    problemasSolucoes.map((item, idx) => (
                      <tr key={item.id} className="hover:bg-slate-50/70 transition-colors">
                        <td className="py-2.5 px-3 align-top">
                          <input
                            type="text"
                            value={item.area}
                            onChange={(e) => handleUpdateRow(item.id, 'area', e.target.value)}
                            placeholder="Área"
                            className="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-semibold text-slate-800"
                          />
                        </td>
                        <td className="py-2.5 px-3 align-top">
                          <textarea
                            rows={2}
                            value={item.problema}
                            onChange={(e) => handleUpdateRow(item.id, 'problema', e.target.value)}
                            placeholder="Descreva o gargalo identificado..."
                            className="w-full bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 resize-y"
                          />
                        </td>
                        <td className="py-2.5 px-3 align-top">
                          <textarea
                            rows={2}
                            value={item.solucaoImplementada}
                            onChange={(e) => handleUpdateRow(item.id, 'solucaoImplementada', e.target.value)}
                            placeholder="Ações e soluções práticas implantadas..."
                            className="w-full bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 resize-y"
                          />
                        </td>
                        <td className="py-2.5 px-3 align-top">
                          <select
                            value={item.status}
                            onChange={(e) => handleUpdateRow(item.id, 'status', e.target.value)}
                            className="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-medium text-slate-800"
                          >
                            <option value="Concluído">Concluído (100%)</option>
                            <option value="Em Andamento">Em Andamento</option>
                            <option value="Parcial">Parcialmente</option>
                            <option value="Não Iniciado">Não Iniciado</option>
                          </select>
                        </td>
                        <td className="py-2.5 px-3 align-top text-center">
                          <select
                            value={item.notaEficacia}
                            onChange={(e) => handleUpdateRow(item.id, 'notaEficacia', parseInt(e.target.value))}
                            className="w-full bg-white border border-slate-200 rounded-lg px-2 py-1.5 text-xs font-bold text-indigo-700 text-center"
                          >
                            {[10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0].map(n => (
                              <option key={n} value={n}>{n}</option>
                            ))}
                          </select>
                        </td>
                        <td className="py-2.5 px-3 align-top">
                          <textarea
                            rows={2}
                            value={item.evidenciaResultado}
                            onChange={(e) => handleUpdateRow(item.id, 'evidenciaResultado', e.target.value)}
                            placeholder="Ex: Planilha validada, redução de 20% no retrabalho..."
                            className="w-full bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 resize-y"
                          />
                        </td>
                        <td className="py-2.5 px-2 align-top text-center">
                          <button
                            type="button"
                            onClick={() => handleDeleteRow(item.id)}
                            className="p-1.5 text-slate-400 hover:text-rose-600 transition-colors rounded-lg hover:bg-rose-50"
                            title="Excluir linha"
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* 3. ANÁLISE QUALITATIVA, IMPACTO NA GESTÃO E FEEDBACK */}
          <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-6">
            <div className="border-b border-slate-100 pb-3">
              <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                <FileText className="h-5 w-5 text-indigo-600" />
                3. Parecer Técnico, Impacto na Gestão e Satisfação do Cliente
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Registre os depoimentos, a transformação na rotina do cliente e as orientações estratégicas de continuidade.
              </p>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Impacto na Gestão */}
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <TrendingUp size={14} className="text-indigo-600" />
                  Impacto Real na Gestão do Cliente (Antes vs. Depois)
                </label>
                <textarea
                  rows={4}
                  value={parecerImpactoGestao}
                  onChange={(e) => setParecerImpactoGestao(e.target.value)}
                  placeholder="Descreva como a gestão, os processos e a maturidade da liderança evoluíram..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:bg-white focus:outline-none"
                />
              </div>

              {/* Ganhos Principais Conquistados */}
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <CheckCircle2 size={14} className="text-emerald-600" />
                  Principais Ganhos Tangíveis e Intangíveis
                </label>
                <textarea
                  rows={4}
                  value={ganhosPrincipais}
                  onChange={(e) => setGanhosPrincipais(e.target.value)}
                  placeholder="Liste as principais conquistas e números alcançados..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:bg-white focus:outline-none font-mono"
                />
              </div>

              {/* Satisfação com a Consultoria */}
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <ThumbsUp size={14} className="text-indigo-600" />
                  Satisfação do Cliente com a Consultoria & Metodologia
                </label>
                <textarea
                  rows={3}
                  value={feedbackConsultoria}
                  onChange={(e) => setFeedbackConsultoria(e.target.value)}
                  placeholder="Relato do cliente sobre o programa de consultoria, aplicabilidade e entregáveis..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:bg-white focus:outline-none"
                />
              </div>

              {/* Satisfação com o Consultor */}
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <UserCheck size={14} className="text-indigo-600" />
                  Satisfação do Cliente com o Consultor (Atuação e Didática)
                </label>
                <textarea
                  rows={3}
                  value={feedbackConsultor}
                  onChange={(e) => setFeedbackConsultor(e.target.value)}
                  placeholder="Feedback sobre a pontualidade, disponibilidade, didática e clareza do consultor..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:bg-white focus:outline-none"
                />
              </div>

              {/* Recomendações e Próximos Passos */}
              <div className="md:col-span-2">
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                  <Sparkles size={14} className="text-amber-500" />
                  Recomendações Finais e Próximos Ciclos Estratégicos
                </label>
                <textarea
                  rows={3}
                  value={recomendacoesFuturas}
                  onChange={(e) => setRecomendacoesFuturas(e.target.value)}
                  placeholder="Orientações para manutenção dos resultados e sugestão de novos passos..."
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs text-slate-800 focus:ring-2 focus:ring-indigo-500 focus:bg-white focus:outline-none"
                />
              </div>
            </div>

            {/* Dados do Consultor e Responsável */}
            <div className="pt-4 border-t border-slate-100 grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                  Nome do Consultor
                </label>
                <input
                  type="text"
                  value={consultorNome}
                  onChange={(e) => setConsultorNome(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs text-slate-800 font-medium"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                  Empresa Credenciada
                </label>
                <input
                  type="text"
                  value={credenciadaNome}
                  onChange={(e) => setCredenciadaNome(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs text-slate-800 font-medium"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase tracking-wider mb-1">
                  Representante do Cliente
                </label>
                <input
                  type="text"
                  value={responsavelClienteNome}
                  onChange={(e) => setResponsavelClienteNome(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs text-slate-800 font-medium"
                />
              </div>
            </div>
          </div>

          {/* Bottom Action Save Bar */}
          <div className="flex items-center justify-between bg-slate-900 text-white p-5 rounded-2xl shadow-xl flex-wrap gap-4">
            <div className="flex items-center gap-3">
              <div className="h-10 w-10 rounded-xl bg-indigo-600 flex items-center justify-center font-black text-amber-300">
                {mediaGeralCalculada.toFixed(1)}
              </div>
              <div>
                <div className="text-xs text-slate-300">Classificação Geral da Consultoria</div>
                <div className="text-sm font-bold text-white flex items-center gap-2">
                  {classificacaoGeral.label}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <Button
                type="button"
                onClick={() => setActiveSubTab('relatorio')}
                variant="outline"
                className="bg-white/10 text-white border-white/20 hover:bg-white/20 text-xs font-bold px-4 py-2.5 rounded-xl"
              >
                <FileText size={15} /> Pré-visualizar Laudo
              </Button>

              <Button
                type="button"
                onClick={handleSaveAnalysis}
                disabled={isSaving}
                className="bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold px-6 py-2.5 rounded-xl shadow-lg flex items-center gap-2"
              >
                {isSaving ? <RefreshCw className="h-4 w-4 animate-spin" /> : <Save size={15} />}
                {currentAnalysisId ? 'Salvar Alterações' : 'Salvar Avaliação'}
              </Button>
            </div>
          </div>
        </div>
      )}

      {/* HISTÓRICO DE ANÁLISES SALVAS */}
      {activeSubTab === 'historico' && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-6">
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h2 className="text-lg font-bold text-slate-900 flex items-center gap-2">
                <Clock className="h-5 w-5 text-indigo-600" />
                Histórico de Avaliações de Resultado
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Consulte ou recarregue análises de resultados realizadas anteriormente para cada cliente.
              </p>
            </div>
          </div>

          {savedAnalyses.length === 0 ? (
            <div className="text-center py-12 bg-slate-50 rounded-xl border border-dashed border-slate-200">
              <Award className="h-12 w-12 text-slate-300 mx-auto mb-2" />
              <p className="text-sm font-semibold text-slate-600">Nenhuma avaliação salva ainda</p>
              <p className="text-xs text-slate-400 mt-1">Preencha o formulário e clique em "Salvar Avaliação" para criar o primeiro registro.</p>
            </div>
          ) : (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
              {savedAnalyses.map((item) => (
                <div
                  key={item.id}
                  className={`p-5 rounded-2xl border transition-all ${
                    currentAnalysisId === item.id
                      ? 'border-indigo-500 bg-indigo-50/40 ring-2 ring-indigo-500/20'
                      : 'border-slate-200 bg-white hover:border-slate-300 hover:shadow-md'
                  }`}
                >
                  <div className="flex items-start justify-between gap-2 mb-3">
                    <div>
                      <h3 className="font-bold text-sm text-slate-900 line-clamp-1">{item.nomeEmpresa}</h3>
                      <p className="text-xs text-slate-500 line-clamp-1">{item.nomeProjeto || 'Consultoria'}</p>
                    </div>
                    <div className="text-right">
                      <span className="inline-block px-2.5 py-0.5 rounded-full text-xs font-black bg-indigo-100 text-indigo-800">
                        {item.mediaGeral ? item.mediaGeral.toFixed(1) : '9.0'} pts
                      </span>
                    </div>
                  </div>

                  <div className="text-xs text-slate-500 space-y-1 mb-4">
                    <div className="flex items-center gap-1.5">
                      <Calendar size={12} className="text-slate-400" />
                      <span>Data: <strong>{item.dataAnalise || 'Recente'}</strong></span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <User size={12} className="text-slate-400" />
                      <span>Consultor: <strong>{item.consultorNome || 'Não informado'}</strong></span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <Target size={12} className="text-slate-400" />
                      <span>Itens Avaliados: <strong>{item.problemasSolucoes?.length || 0}</strong></span>
                    </div>
                  </div>

                  <div className="flex items-center justify-between pt-3 border-t border-slate-100">
                    <button
                      onClick={() => {
                        loadAnalysisIntoForm(item);
                        setActiveSubTab('form');
                      }}
                      className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1"
                    >
                      Abrir / Editar <ChevronRight size={14} />
                    </button>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => {
                          loadAnalysisIntoForm(item);
                          setActiveSubTab('relatorio');
                        }}
                        className="p-1.5 text-slate-500 hover:text-slate-900 rounded-lg hover:bg-slate-100"
                        title="Ver Relatório PDF"
                      >
                        <FileText size={14} />
                      </button>
                      <button
                        onClick={() => item.id && handleDeleteSaved(item.id)}
                        className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg hover:bg-rose-50"
                        title="Excluir Registro"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* RELATÓRIO EXECUTIVO / LAUDO PDF COM CABEÇALHO COMPLETO */}
      {activeSubTab === 'relatorio' && (
        <div className="space-y-4">
          <div className="bg-slate-100 p-4 rounded-xl flex items-center justify-between border border-slate-200">
            <div className="flex items-center gap-2 text-slate-700 text-xs font-medium">
              <Printer size={16} className="text-indigo-600" />
              <span>Pré-visualização do Laudo Oficial de Resultado da Consultoria. Clique em <strong>"Gerar Laudo PDF"</strong> para baixar.</span>
            </div>
            <Button
              onClick={handleExportPDF}
              disabled={isExportingPdf}
              className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2 rounded-xl flex items-center gap-1.5"
            >
              {isExportingPdf ? <RefreshCw className="h-4 w-4 animate-spin" /> : <Download size={14} />}
              Baixar em PDF
            </Button>
          </div>

          {/* O CONTAINER CAPTURADO PELO HTML2CANVAS */}
          <div
            id="consultancy-result-pdf-report"
            className="bg-white p-8 sm:p-12 rounded-2xl border border-slate-200 shadow-xl max-w-4xl mx-auto text-slate-900 space-y-6"
            style={{ minHeight: '1100px', fontFamily: 'Inter, system-ui, sans-serif' }}
          >
            {/* CABEÇALHO OFICIAL DA CONSULTORIA */}
            <div className="border-b-2 border-slate-800 pb-5">
              <div className="flex items-center justify-between gap-6">
                <div className="flex-1">
                  <div className="text-[11px] font-black text-indigo-900 uppercase tracking-widest mb-1">
                    SISTEMA DE GESTÃO E CONSULTORIA EMPRESARIAL
                  </div>
                  <h1 className="text-xl sm:text-2xl font-black text-slate-900 tracking-tight uppercase">
                    LAUDO DE AVALIAÇÃO DE RESULTADOS DA CONSULTORIA
                  </h1>
                  <p className="text-xs text-slate-500 font-medium">
                    Diagnóstico de Problemas, Soluções Implementadas e Impacto na Gestão
                  </p>
                </div>

                {/* Logotipo */}
                <div className="flex items-center gap-3 flex-shrink-0">
                  {customLogo ? (
                    <img
                      src={customLogo}
                      alt="Logo SEBRAE"
                      className="h-12 max-w-[130px] object-contain"
                    />
                  ) : (!customConsultoraLogo && sebraeLogoBase64 ? (
                    <img
                      src={sebraeLogoBase64}
                      alt="Logo SEBRAE"
                      className="h-12 max-w-[130px] object-contain"
                    />
                  ) : null)}

                  {customConsultoraLogo && (
                    <img
                      src={customConsultoraLogo}
                      alt="Logo Consultora"
                      className="h-12 max-w-[130px] object-contain"
                    />
                  )}
                </div>
              </div>

              {/* Grid com Dados da Consultoria e do Cliente */}
              <div className="mt-5 grid grid-cols-2 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-200 text-xs">
                <div>
                  <div className="text-[10px] uppercase font-bold text-slate-500 tracking-wider">Empresa Cliente</div>
                  <div className="font-bold text-slate-900 text-sm">{currentEmpresa?.nome || currentEmpresa?.razaoSocial || 'Cliente'}</div>
                  {currentEmpresa?.cnpj && <div className="text-slate-600">CNPJ: {currentEmpresa.cnpj}</div>}
                  {currentEmpresa?.representante && <div className="text-slate-600">Representante: {currentEmpresa.representante}</div>}
                  {currentEmpresa?.enderecoComercial && <div className="text-slate-500 text-[11px]">{currentEmpresa.enderecoComercial}</div>}
                </div>

                <div>
                  <div className="text-[10px] uppercase font-bold text-slate-500 tracking-wider">Dados do Atendimento</div>
                  <div className="font-semibold text-slate-900">Projeto: {currentDiagnostico?.nomeProjeto || currentDiagnostico?.nome || 'Consultoria Gerencial'}</div>
                  <div className="text-slate-600">Consultor(a): <strong>{consultorNome || 'Consultor Especialista'}</strong></div>
                  {credenciadaNome && <div className="text-slate-600">Credenciada: {credenciadaNome}</div>}
                  <div className="flex items-center gap-3 text-slate-600 mt-0.5">
                    {codigoSgf && <span>Código SGF: <strong>{codigoSgf}</strong></span>}
                    <span>Carga Horária: <strong>{cargaHoraria}</strong></span>
                    <span>Data: <strong>{dataAnalise}</strong></span>
                  </div>
                </div>
              </div>
            </div>

            {/* PAINEL RESUMO DE NOTAS (0 A 10) */}
            <div className="bg-slate-900 text-white p-5 rounded-xl flex items-center justify-between gap-4">
              <div>
                <div className="text-[10px] uppercase font-bold text-indigo-300 tracking-widest">Desempenho Geral da Consultoria</div>
                <div className="text-2xl font-black text-amber-300 flex items-baseline gap-1.5 mt-0.5">
                  {mediaGeralCalculada.toFixed(1)} <span className="text-xs text-white/70 font-normal">/ 10.0 Pontos</span>
                </div>
                <div className="text-xs font-semibold text-white/90 mt-0.5">
                  Classificação: <span className="text-amber-300 font-bold">{classificacaoGeral.label}</span>
                </div>
              </div>

              <div className="grid grid-cols-3 gap-2 text-[10px] max-w-sm">
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Problemas</div>
                  <div className="text-sm font-black text-white">{notas.eficaciaProblemas}/10</div>
                </div>
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Gestão</div>
                  <div className="text-sm font-black text-white">{notas.impactoGestao}/10</div>
                </div>
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Finanças</div>
                  <div className="text-sm font-black text-white">{notas.impactoFinanceiroProd}/10</div>
                </div>
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Equipe</div>
                  <div className="text-sm font-black text-white">{notas.engajamentoEquipe}/10</div>
                </div>
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Consultoria</div>
                  <div className="text-sm font-black text-white">{notas.satisfacaoConsultoria}/10</div>
                </div>
                <div className="bg-white/10 p-2 rounded-lg text-center">
                  <div className="text-white/70 font-medium">Consultor</div>
                  <div className="text-sm font-black text-white">{notas.satisfacaoConsultor}/10</div>
                </div>
              </div>
            </div>

            {/* TABELA DE PROBLEMAS DIAGNOSTICADOS VS SOLUÇÕES IMPLEMENTADAS */}
            <div className="space-y-2">
              <h3 className="text-xs font-black uppercase text-slate-800 tracking-wider border-b border-slate-200 pb-1 flex items-center justify-between">
                <span>1. Matriz de Problemas Diagnosticados e Soluções Implementadas</span>
                <span className="text-[10px] text-slate-500 font-normal">Total: {problemasSolucoes.length} itens</span>
              </h3>

              <table className="w-full text-left text-[11px] border border-slate-200 border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-800 font-bold border-b border-slate-200">
                    <th className="p-2 border-r border-slate-200 w-24">Área</th>
                    <th className="p-2 border-r border-slate-200">Problema Diagnosticado</th>
                    <th className="p-2 border-r border-slate-200">Solução Prática Implementada</th>
                    <th className="p-2 border-r border-slate-200 w-24">Status</th>
                    <th className="p-2 text-center w-16">Eficácia</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {problemasSolucoes.map((row) => (
                    <tr key={row.id} className="text-slate-800">
                      <td className="p-2 border-r border-slate-200 font-bold text-[10px] bg-slate-50 align-top">
                        {row.area}
                      </td>
                      <td className="p-2 border-r border-slate-200 align-top">
                        {row.problema || '—'}
                      </td>
                      <td className="p-2 border-r border-slate-200 align-top">
                        <div className="font-medium">{row.solucaoImplementada || '—'}</div>
                        {row.evidenciaResultado && (
                          <div className="text-[10px] text-slate-500 italic mt-0.5">
                            Evidência: {row.evidenciaResultado}
                          </div>
                        )}
                      </td>
                      <td className="p-2 border-r border-slate-200 align-top text-[10px]">
                        <span className="font-semibold">{row.status}</span>
                      </td>
                      <td className="p-2 text-center align-top font-bold text-indigo-800">
                        {row.notaEficacia}/10
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* PARECER TÉCNICO E IMPACTO NA GESTÃO */}
            <div className="grid grid-cols-2 gap-4 text-xs">
              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                <div className="text-[10px] font-black uppercase text-slate-700 tracking-wider">
                  2. Impacto na Gestão e Processos
                </div>
                <p className="text-slate-700 leading-relaxed text-[11px]">
                  {parecerImpactoGestao}
                </p>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                <div className="text-[10px] font-black uppercase text-slate-700 tracking-wider">
                  3. Principais Ganhos Conquistados
                </div>
                <p className="text-slate-700 leading-relaxed text-[11px] whitespace-pre-line font-mono">
                  {ganhosPrincipais}
                </p>
              </div>
            </div>

            {/* SATISFAÇÃO COM A CONSULTORIA E COM O CONSULTOR */}
            <div className="grid grid-cols-2 gap-4 text-xs">
              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                <div className="text-[10px] font-black uppercase text-slate-700 tracking-wider flex items-center justify-between">
                  <span>4. Satisfação com a Consultoria</span>
                  <span className="text-indigo-800 font-bold">{notas.satisfacaoConsultoria}/10 pts</span>
                </div>
                <p className="text-slate-700 leading-relaxed text-[11px]">
                  {feedbackConsultoria}
                </p>
              </div>

              <div className="p-3.5 bg-slate-50 border border-slate-200 rounded-xl space-y-1.5">
                <div className="text-[10px] font-black uppercase text-slate-700 tracking-wider flex items-center justify-between">
                  <span>5. Satisfação com o Consultor</span>
                  <span className="text-indigo-800 font-bold">{notas.satisfacaoConsultor}/10 pts</span>
                </div>
                <p className="text-slate-700 leading-relaxed text-[11px]">
                  {feedbackConsultor}
                </p>
              </div>
            </div>

            {/* RECOMENDAÇÕES FINAIS */}
            <div className="p-3.5 bg-indigo-50/50 border border-indigo-200 rounded-xl space-y-1.5 text-xs">
              <div className="text-[10px] font-black uppercase text-indigo-900 tracking-wider">
                6. Recomendações e Próximos Passos
              </div>
              <p className="text-slate-800 leading-relaxed text-[11px]">
                {recomendacoesFuturas}
              </p>
            </div>

            {/* BLOCO DE ASSINATURAS */}
            <div className="pt-8 border-t border-slate-200 grid grid-cols-2 gap-12 text-center text-xs">
              <div>
                <div className="border-b border-slate-400 pb-1 mb-1.5 w-4/5 mx-auto" />
                <div className="font-bold text-slate-900">{consultorNome || 'Consultor Responsável'}</div>
                <div className="text-[10px] text-slate-500">{credenciadaNome || 'Consultoria Empresarial'}</div>
              </div>

              <div>
                <div className="border-b border-slate-400 pb-1 mb-1.5 w-4/5 mx-auto" />
                <div className="font-bold text-slate-900">{responsavelClienteNome || currentEmpresa?.representante || 'Representante Legal'}</div>
                <div className="text-[10px] text-slate-500">{responsavelClienteCargo} — {currentEmpresa?.nome || 'Empresa Cliente'}</div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
