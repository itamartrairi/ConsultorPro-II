import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { db, auth } from '../firebase';
import { collection, query, where, onSnapshot, addDoc, doc, deleteDoc, getDocs } from 'firebase/firestore';
import {
  TrendingUp,
  Award,
  Sparkles,
  ArrowLeft,
  ArrowRight,
  Check,
  Plus,
  Trash2,
  Download,
  RefreshCw,
  Building2,
  Calendar,
  AlertCircle,
  HelpCircle,
  CheckCircle2,
  Info,
  Clock,
  ChevronRight,
  Target,
  FileText,
  Printer
} from 'lucide-react';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';

// Definition of Maturity Assessment from Image and Firestore Schema
interface Competency {
  id: string;
  name: string;
  description: string;
  icon: string;
  question: string;
}

const COMPETENCIES: Competency[] = [
  {
    id: 'disciplina',
    name: '1. DISCIPLINA',
    description: 'Habilidade de manter a constância, seguir rotinas e honrar compromissos de forma firme e consistente.',
    icon: '🎯',
    question: 'Com que frequência você segue rigorosamente os processos definidos e mantém a consistência das rotinas, mesmo diante de distrações?'
  },
  {
    id: 'planejamento',
    name: '2. PLANEJAMENTO',
    description: 'Capacidade de antecipar cenários, estruturar metas e detalhar planos de ação lógicos e claros.',
    icon: '📅',
    question: 'Você possui um planejamento estruturado com metas de curto, médio e longo prazo bem definidas e revisadas periodicamente?'
  },
  {
    id: 'organizacao',
    name: '3. ORGANIZAÇÃO',
    description: 'Estruturação de processos, recursos, arquivos e fluxos de trabalho visando eficiência e clareza.',
    icon: '🗂️',
    question: 'Como você avalia a organização dos processos internos, arquivos e fluxos de trabalho do seu negócio hoje?'
  },
  {
    id: 'gestaoTempo',
    name: '4. GESTÃO DO TEMPO',
    description: 'Priorização inteligente de tarefas, foco no que é estratégico e mitigação de distrações diárias.',
    icon: '⏳',
    question: 'Com que eficácia você prioriza o que é estratégico e gerencia seu tempo para evitar a sensação de "apenas apagar incêndios"?'
  },
  {
    id: 'persistencia',
    name: '5. PERSISTÊNCIA',
    description: 'Determinação para continuar diante de obstáculos, falhas temporárias e cenários adversos de mercado.',
    icon: '🛡️',
    question: 'Diante de crises ou falhas temporárias, qual é a sua capacidade de persistir, replanejar e manter o foco na visão?'
  },
  {
    id: 'lideranca',
    name: '6. LIDERANÇA',
    description: 'Capacidade de influenciar, motivar, engajar e guiar pessoas rumo ao propósito comum e metas da empresa.',
    icon: '👥',
    question: 'Quão bem você delega responsabilidades, engaja sua equipe e desenvolve novos líderes no seu negócio?'
  },
  {
    id: 'tomadaDecisao',
    name: '7. TOMADA DE DECISÃO',
    description: 'Assertividade, coragem e embasamento racional para definir rumos estratégicos sob pressão de tempo.',
    icon: '⚖️',
    question: 'Como você avalia sua velocidade e assertividade para tomar decisões cruciais, equilibrando razão e coragem?'
  },
  {
    id: 'focoResultados',
    name: '8. FOCO EM RESULTADOS',
    description: 'Orientação voltada para a entrega prática, superação de metas e geração de valor tangível.',
    icon: '📈',
    question: 'As suas ações diárias são estritamente guiadas por métricas, indicadores-chave de desempenho (KPIs) e entrega de resultados?'
  },
  {
    id: 'inteligenciaEmocional',
    name: '9. INTELIGÊNCIA EMOCIONAL',
    description: 'Autoconhecimento, empatia e autocontrole para gerir o estresse e conflitos de forma equilibrada.',
    icon: '🧠',
    question: 'Qual é o seu nível de autocontrole, resiliência e empatia para gerenciar crises, pressões e conflitos internos ou externos?'
  },
  {
    id: 'capacidadeExecucao',
    name: '10. CAPACIDADE DE EXECUÇÃO',
    description: 'Atitude proativa ("fazer acontecer"), agilidade operacional e capacidade de transformar ideias em entregas reais.',
    icon: '⚙️',
    question: 'Qual é a sua velocidade e qualidade de execução para tirar planos e projetos do papel e transformá-los em realidade operacional?'
  }
];

// Strategic action suggestions to add directly to Action Plan
const RECOMMENDATIONS_DB: Record<string, string[]> = {
  disciplina: [
    "Criar uma rotina de revisão diária de 15 minutos (Check-in) com os líderes de setor antes de iniciar as atividades.",
    "Instaurar um ritual semanal de acompanhamento operacional rigoroso focado em tarefas pendentes.",
    "Elaborar cronograma pessoal para atividades críticas com alarmes e blocos de tempo fixos na agenda."
  ],
  planejamento: [
    "Definir os objetivos estratégicos da empresa (OKRs) para os próximos 12 meses divididos por trimestre.",
    "Estruturar um plano de contingência documentado para os 3 principais riscos identificados na operação.",
    "Realizar reunião trimestral de alinhamento estratégico com os principais parceiros e equipe chave."
  ],
  organizacao: [
    "Mapear visualmente os fluxos de trabalho atuais e documentar os 5 principais processos internos em POPs.",
    "Organizar um repositório central na nuvem (Drive) com pastas padronizadas para cada setor da empresa.",
    "Implementar ferramenta visual de gestão (como Trello, Asana ou Kanban físico) para as rotinas corporativas."
  ],
  gestaoTempo: [
    "Utilizar a Matriz de Eisenhower semanalmente para separar tarefas Importantes de Urgentes e delegar o operacional.",
    "Estabelecer blocos de tempo intocáveis (Focus Time) de 2 horas diárias dedicadas exclusivamente ao estratégico.",
    "Reduzir o tempo de reuniões internas em 50% aplicando pautas curtas e objetivas pré-definidas."
  ],
  persistencia: [
    "Criar comitê de desvios quinzenal para analisar metas não alcançadas de forma racional sem desmotivação.",
    "Estabelecer indicadores de longo prazo claros para manter a equipe focada na visão macro em tempos difíceis.",
    "Implementar reuniões pós-crise ('Post-Mortem') para extrair aprendizados operacionais das falhas."
  ],
  lideranca: [
    "Instaurar rituais de feedback individual (One-on-One) quinzenais de 30 minutos com cada colaborador direto.",
    "Desenhar plano de metas individual para incentivar a autonomia e engajamento da equipe média.",
    "Mapear competências técnicas e comportamentais necessárias para sucessão interna de cargos chave."
  ],
  tomadaDecisao: [
    "Estruturar relatório mensal consolidado de DRE e Fluxo de Caixa para embasar decisões financeiras em dados.",
    "Definir matriz de alçada de decisão para que gerentes tomem decisões de baixo impacto de forma autônoma.",
    "Adotar checklist formal de análise de risco e custo-benefício antes de contratações ou investimentos maiores."
  ],
  focoResultados: [
    "Definir 3 KPIs principais para cada departamento e criar um painel visual (Dashboard) na sala de reuniões.",
    "Implantar programa simples de premiação ou bônus trimestral atrelado ao cumprimento de metas de faturamento.",
    "Substituir reuniões informais por reuniões rápidas de resultados focadas exclusivamente em desvios de métricas."
  ],
  inteligenciaEmocional: [
    "Realizar avaliação de clima organizacional anônima para compreender os reais pontos de estresse da equipe.",
    "Promover sessões de liderança focadas em comunicação não-violenta (CNV) e inteligência emocional em crises.",
    "Estabelecer momentos de descompressão semanais e canais abertos de escuta para apoio emocional aos colaboradores."
  ],
  capacidadeExecucao: [
    "Adotar rituais diários rápidos de 10 minutos (Daily Scrum) para destravar gargalos operacionais imediatos.",
    "Definir um padrinho/responsável e prazo inegociável para toda nova ideia aprovada antes de iniciar sua execução.",
    "Quebrar grandes projetos em tarefas menores com duração de no máximo 3 dias para acelerar a sensação de progresso."
  ]
};

interface MaturityEvaluation {
  id?: string;
  empresaId: string;
  dataAvaliacao: string;
  ownerId: string;
  pontuacaoTotal: number;
  nivelMaturidade: string;
  scores: {
    disciplina: number;
    planejamento: number;
    organizacao: number;
    gestaoTempo: number;
    persistencia: number;
    lideranca: number;
    tomadaDecisao: number;
    focoResultados: number;
    inteligenciaEmocional: number;
    capacidadeExecucao: number;
  };
  observacoes?: string;
}

interface MaturidadeViewProps {
  empresas: { id: string; nome: string }[];
  setView: (view: string) => void;
}

export default function MaturidadeView({ empresas, setView }: MaturidadeViewProps) {
  const [activeTab, setActiveTab] = useState<'history' | 'new-assessment' | 'evolution'>('history');
  const [selectedEmpresaId, setSelectedEmpresaId] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(true);

  // Firestore Data States
  const [evaluations, setEvaluations] = useState<MaturityEvaluation[]>([]);
  const [selectedEval, setSelectedEval] = useState<MaturityEvaluation | null>(null);
  const [actionPlanTasks, setActionPlanTasks] = useState<any[]>([]);
  const [addedTasksKeys, setAddedTasksKeys] = useState<Record<string, boolean>>({});

  // New Assessment Wizard State
  const [wizardScores, setWizardScores] = useState<Record<string, number>>({
    disciplina: 3,
    planejamento: 3,
    organizacao: 3,
    gestaoTempo: 3,
    persistencia: 3,
    lideranca: 3,
    tomadaDecisao: 3,
    focoResultados: 3,
    inteligenciaEmocional: 3,
    capacidadeExecucao: 3
  });
  const [wizardObservacoes, setWizardObservacoes] = useState<string>('');
  const [savingAssessment, setSavingAssessment] = useState<boolean>(false);

  // Initialize selected company
  useEffect(() => {
    if (empresas.length > 0 && !selectedEmpresaId) {
      setSelectedEmpresaId(empresas[0].id);
    }
  }, [empresas]);

  // Fetch Evaluations from Firestore
  useEffect(() => {
    const user = auth.currentUser;
    if (!user) return;

    setLoading(true);
    const qEvs = query(
      collection(db, 'maturidade_avaliacoes'),
      where('ownerId', '==', user.uid)
    );

    const unsubscribe = onSnapshot(qEvs, (snap) => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() } as MaturityEvaluation));
      // Sort chronologically desc
      data.sort((a, b) => new Date(b.dataAvaliacao).getTime() - new Date(a.dataAvaliacao).getTime());
      setEvaluations(data);
      setLoading(false);

      // Select default latest evaluation
      const filtered = data.filter(e => e.empresaId === selectedEmpresaId);
      if (filtered.length > 0 && !selectedEval) {
        setSelectedEval(filtered[0]);
      }
    }, (error) => {
      console.error("Erro ao carregar diagnósticos de maturidade:", error);
      setLoading(false);
    });

    return () => unsubscribe();
  }, [selectedEmpresaId]);

  // Fetch Action Plan Tasks
  useEffect(() => {
    const user = auth.currentUser;
    if (!user || !selectedEmpresaId) return;

    const qTasks = query(
      collection(db, 'tarefas_plano'),
      where('ownerId', '==', user.uid),
      where('empresaId', '==', selectedEmpresaId)
    );

    const unsubscribe = onSnapshot(qTasks, (snap) => {
      const tasks = snap.docs.map(d => ({ id: d.id, ...d.data() }));
      setActionPlanTasks(tasks);
    }, (error) => {
      console.error("Erro ao carregar plano de ação:", error);
    });

    return () => unsubscribe();
  }, [selectedEmpresaId]);

  // Handle Select Evaluation
  const handleSelectEval = (ev: MaturityEvaluation) => {
    setSelectedEval(ev);
    setActiveTab('history');
  };

  // Level Logic based on total points (10 to 50)
  const getMaturityLevel = (points: number) => {
    if (points >= 10 && points <= 20) {
      return {
        level: "INICIAL",
        range: "10 a 20 pontos",
        color: "text-rose-600 bg-rose-50 border-rose-200",
        badge: "bg-rose-500",
        interpretation: "O empresário necessita de desenvolvimento urgente em diversas áreas. Processos são altamente centralizados, reativos e sem planejamento formal."
      };
    } else if (points >= 21 && points <= 35) {
      return {
        level: "BÁSICO",
        range: "21 a 35 pontos",
        color: "text-amber-600 bg-amber-50 border-amber-200",
        badge: "bg-amber-500",
        interpretation: "Algumas competências já estão em desenvolvimento, mas ainda há áreas críticas que precisam de atenção imediata para estruturar a gestão básica."
      };
    } else if (points >= 36 && points <= 45) {
      return {
        level: "INTERMEDIÁRIO",
        range: "36 a 45 pontos",
        color: "text-blue-600 bg-blue-50 border-blue-200",
        badge: "bg-sky-600",
        interpretation: "O empresário possui uma base de gestão sólida, mas pode aprimorar competências específicas para delegar melhor e acelerar os resultados de mercado."
      };
    } else {
      return {
        level: "AVANÇADO",
        range: "46 a 50 pontos",
        color: "text-emerald-600 bg-emerald-50 border-emerald-200",
        badge: "bg-emerald-600",
        interpretation: "O empresário demonstra excelência na grande maioria das competências chaves, com foco nítido em governança, delegação e melhoria contínua de escala."
      };
    }
  };

  // Add Recommended Action to Action Plan in Firestore
  const handleAddRecommendation = async (compName: string, recommendationText: string, key: string) => {
    const user = auth.currentUser;
    if (!user || !selectedEmpresaId) return;

    try {
      await addDoc(collection(db, 'tarefas_plano'), {
        empresaId: selectedEmpresaId,
        ownerId: user.uid,
        idProblema: `maturidade_${key}`,
        problema: `Maturidade - Baixa pontuação em ${compName}`,
        solucaoSugerida: recommendationText,
        acoes: recommendationText,
        area: "Geral / Competências",
        status: "Pendente",
        prioridade: "Alta",
        dataInicio: new Date().toISOString().split('T')[0],
        responsavel: "Empresário",
        comentarios: "Ação corretiva sugerida automaticamente pelo Diagnóstico de Maturidade das Competências."
      });

      setAddedTasksKeys(prev => ({ ...prev, [key + recommendationText]: true }));
    } catch (err) {
      console.error("Erro ao adicionar recomendação ao plano:", err);
    }
  };

  // Delete Evaluation
  const handleDeleteEval = async (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!window.confirm("Deseja realmente excluir este diagnóstico de maturidade?")) return;

    try {
      await deleteDoc(doc(db, 'maturidade_avaliacoes', id));
      if (selectedEval?.id === id) {
        setSelectedEval(null);
      }
    } catch (err) {
      console.error("Erro ao excluir diagnóstico:", err);
    }
  };

  // Submit New Assessment
  const handleSubmitAssessment = async () => {
    const user = auth.currentUser;
    if (!user || !selectedEmpresaId) {
      alert("Por favor, selecione uma empresa para salvar o diagnóstico.");
      return;
    }

    setSavingAssessment(true);

    const totalPoints = 
      (wizardScores.disciplina || 3) +
      (wizardScores.planejamento || 3) +
      (wizardScores.organizacao || 3) +
      (wizardScores.gestaoTempo || 3) +
      (wizardScores.persistencia || 3) +
      (wizardScores.lideranca || 3) +
      (wizardScores.tomadaDecisao || 3) +
      (wizardScores.focoResultados || 3) +
      (wizardScores.inteligenciaEmocional || 3) +
      (wizardScores.capacidadeExecucao || 3);
    const maturityInfo = getMaturityLevel(totalPoints);

    const newAssessment: MaturityEvaluation = {
      empresaId: selectedEmpresaId,
      dataAvaliacao: new Date().toISOString(),
      ownerId: user.uid,
      pontuacaoTotal: totalPoints,
      nivelMaturidade: maturityInfo.level,
      scores: {
        disciplina: wizardScores.disciplina,
        planejamento: wizardScores.planejamento,
        organizacao: wizardScores.organizacao,
        gestaoTempo: wizardScores.gestaoTempo,
        persistencia: wizardScores.persistencia,
        lideranca: wizardScores.lideranca,
        tomadaDecisao: wizardScores.tomadaDecisao,
        focoResultados: wizardScores.focoResultados,
        inteligenciaEmocional: wizardScores.inteligenciaEmocional,
        capacidadeExecucao: wizardScores.capacidadeExecucao
      },
      observacoes: wizardObservacoes
    };

    try {
      await addDoc(collection(db, 'maturidade_avaliacoes'), newAssessment);
      setSavingAssessment(false);
      // Reset scores to default
      setWizardScores({
        disciplina: 3,
        planejamento: 3,
        organizacao: 3,
        gestaoTempo: 3,
        persistencia: 3,
        lideranca: 3,
        tomadaDecisao: 3,
        focoResultados: 3,
        inteligenciaEmocional: 3,
        capacidadeExecucao: 3
      });
      setWizardObservacoes('');
      setActiveTab('history');
    } catch (err) {
      console.error("Erro ao salvar avaliação de maturidade:", err);
      setSavingAssessment(false);
    }
  };

  // Calculate stats for Action Plan
  const totalTasks = actionPlanTasks.length;
  const completedTasks = actionPlanTasks.filter(t => t.status === 'Concluído' || t.status === 'concluido' || t.status === 'Concluída' || t.status === 'concluida').length;
  const taskProgress = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

  // Filter evaluations for the current selected company
  const companyEvaluations = evaluations.filter(e => e.empresaId === selectedEmpresaId);

  // SVG Radar Decagon Chart Component Math
  const drawMaturityRadar = (scores: MaturityEvaluation['scores']) => {
    const size = 360;
    const center = size / 2;
    const maxRadius = 115;

    // Angles for the 10 vertices
    // index 0 is at 12 o'clock, which is -90 degrees (-pi/2 rad)
    const getCoordinates = (index: number, val: number) => {
      const angle = (index * 2 * Math.PI) / 10 - Math.PI / 2;
      // Map score 0 to 5 onto radius 0 to maxRadius
      const r = (val / 5) * maxRadius;
      return {
        x: center + r * Math.cos(angle),
        y: center + r * Math.sin(angle)
      };
    };

    // Form polygon vertices string
    const polyPoints = COMPETENCIES.map((comp, idx) => {
      const scoreVal = scores[comp.id as keyof MaturityEvaluation['scores']] || 3;
      const coords = getCoordinates(idx, scoreVal);
      return `${coords.x},${coords.y}`;
    }).join(' ');

    return (
      <svg width="100%" height="100%" viewBox={`0 0 ${size} ${size}`} className="max-w-[350px] mx-auto select-none drop-shadow-md">
        {/* Radial Decagon concentric circles/decagons */}
        {[1, 2, 3, 4, 5].map((lvl) => {
          const pointsStr = COMPETENCIES.map((_, idx) => {
            const coords = getCoordinates(idx, lvl);
            return `${coords.x},${coords.y}`;
          }).join(' ');

          return (
            <g key={lvl}>
              {/* Draw Decagon rings */}
              <polygon
                points={pointsStr}
                fill="none"
                stroke={lvl === 5 ? "#475569" : "#e2e8f0"}
                strokeWidth={lvl === 5 ? "1.5" : "1"}
                strokeDasharray={lvl === 5 ? "" : "3 3"}
              />
              {/* Level Markers at center vertical line */}
              <text
                x={center}
                y={center - (lvl / 5) * maxRadius + 11}
                fontSize="8"
                fontWeight="700"
                fill="#94a3b8"
                textAnchor="middle"
              >
                {lvl}
              </text>
            </g>
          );
        })}

        {/* Diagonal axis divider lines */}
        {COMPETENCIES.map((_, idx) => {
          const outerCoords = getCoordinates(idx, 5);
          return (
            <line
              key={idx}
              x1={center}
              y1={center}
              x2={outerCoords.x}
              y2={outerCoords.y}
              stroke="#e2e8f0"
              strokeWidth="1"
            />
          );
        })}

        {/* Highlight quadrants/halves background if needed */}
        {/* Plot actual user score polygon area */}
        <polygon
          points={polyPoints}
          fill="rgba(14, 165, 233, 0.22)"
          stroke="#0284c7"
          strokeWidth="2.5"
          strokeLinejoin="round"
          className="transition-all duration-700 ease-out"
        />

        {/* Center mark */}
        <circle cx={center} cy={center} r="4.5" fill="#0284c7" />

        {/* Vertices value labels and labels on the periphery */}
        {COMPETENCIES.map((comp, idx) => {
          const scoreVal = scores[comp.id as keyof MaturityEvaluation['scores']] || 3;
          const coords = getCoordinates(idx, scoreVal);
          const outerCoords = getCoordinates(idx, 5.8); // Offset outward for labels

          // Calculate correct alignment anchor based on position
          let textAnchor = "middle";
          if (outerCoords.x < center - 30) textAnchor = "end";
          if (outerCoords.x > center + 30) textAnchor = "start";

          let labelYOffset = 3;
          if (idx === 0) labelYOffset = -10; // "Disciplina" at top
          if (idx === 5) labelYOffset = 12; // "Liderança" at bottom

          return (
            <g key={comp.id}>
              {/* Vertex Score point marker dot */}
              <circle
                cx={coords.x}
                cy={coords.y}
                r="4.5"
                fill="#0284c7"
                stroke="white"
                strokeWidth="1.5"
                className="transition-all duration-700"
              />
              {/* Text score value marker */}
              <text
                x={coords.x + (coords.x > center ? 9 : -9)}
                y={coords.y + 3}
                fontSize="8.5"
                fontWeight="800"
                fill="#1e293b"
                textAnchor={coords.x > center ? "start" : "end"}
              >
                {scoreVal.toFixed(1)}
              </text>

              {/* Competency Name around periphery */}
              <text
                x={outerCoords.x}
                y={outerCoords.y + labelYOffset}
                fontSize="8"
                fontWeight="700"
                fill="#475569"
                textAnchor={textAnchor}
              >
                {idx + 1}. {comp.id === 'gestaoTempo' ? 'G. TEMPO' : comp.id === 'inteligenciaEmocional' ? 'INT. EMOCIONAL' : comp.id === 'capacidadeExecucao' ? 'CAP. EXECUÇÃO' : comp.id === 'tomadaDecisao' ? 'DECISÃO' : comp.id === 'focoResultados' ? 'RESULTADOS' : comp.name.split('. ')[1]}
              </text>
            </g>
          );
        })}
      </svg>
    );
  };

  // Export Assessment as PDF
  const handleExportPDF = () => {
    const element = document.getElementById('maturity-report-container');
    if (!element) return;

    const currentEmpresa = empresas.find(e => e.id === selectedEmpresaId);

    const opt = {
      margin: 10,
      filename: `Diagnostico_Maturidade_${currentEmpresa?.nome || 'Empresa'}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2canvas(element, { scale: 2 }).then((canvas) => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgWidth = 190;
      const pageHeight = 295;
      const imgHeight = (canvas.height * imgWidth) / canvas.width;
      let heightLeft = imgHeight;
      let position = 10;

      pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
      heightLeft -= pageHeight;

      while (heightLeft >= 0) {
        position = heightLeft - imgHeight + 10;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;
      }
      pdf.save(opt.filename);
    });
  };

  const currentEmpresaName = empresas.find(e => e.id === selectedEmpresaId)?.nome || "Não selecionada";

  return (
    <div className="space-y-8 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
      {/* Header and Company Selector Row */}
      <div className="flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-100 pb-5 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-900 tracking-tight flex items-center gap-3">
            <TrendingUp className="h-8 w-8 text-sky-600" />
            Maturidade do Empresário
          </h1>
          <p className="mt-1 text-sm text-slate-500">
            Avalie as 10 competências de gestão vitais e acompanhe a evolução conforme o plano de ação avança.
          </p>
        </div>
        
        {/* Selection filters */}
        <div className="flex flex-wrap items-center gap-3">
          <div className="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
            <Building2 className="h-4 w-4 text-slate-400" />
            <select
              value={selectedEmpresaId}
              onChange={(e) => {
                setSelectedEmpresaId(e.target.value);
                setSelectedEval(null); // Reset selection
              }}
              className="bg-transparent border-none text-sm font-semibold text-slate-700 focus:outline-none cursor-pointer"
            >
              {empresas.map((emp) => (
                <option key={emp.id} value={emp.id}>{emp.nome}</option>
              ))}
            </select>
          </div>

          <div className="flex gap-2 bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab('history')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === 'history' ? 'bg-white text-sky-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Diagnóstico Atual
            </button>
            <button
              onClick={() => setActiveTab('evolution')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all ${
                activeTab === 'evolution' ? 'bg-white text-sky-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              Gráfico de Evolução
            </button>
            <button
              onClick={() => setActiveTab('new-assessment')}
              className={`px-3 py-1.5 text-xs font-semibold rounded-lg transition-all flex items-center gap-1 ${
                activeTab === 'new-assessment' ? 'bg-sky-600 text-white shadow-sm' : 'bg-sky-50 text-sky-600 hover:bg-sky-100'
              }`}
            >
              <Plus className="h-3 w-3" /> Novo
            </button>
          </div>
        </div>
      </div>

      {/* Main workspace layout */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        
        {/* LEFT COLUMN: GUIDES & QUICK HISTORY INDEX */}
        <div className="lg:col-span-4 space-y-6">
          
          {/* List of past assessments for the current company */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm">
            <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider mb-4 flex items-center justify-between">
              <span>Histórico da Empresa</span>
              <span className="text-xs bg-slate-100 text-slate-500 font-semibold px-2 py-0.5 rounded-full">
                {companyEvaluations.length}
              </span>
            </h3>
            
            {companyEvaluations.length === 0 ? (
              <div className="text-center py-8">
                <p className="text-xs text-slate-400">Nenhum diagnóstico de competência registrado para esta empresa.</p>
                <button
                  onClick={() => setActiveTab('new-assessment')}
                  className="mt-4 inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-sky-50 text-sky-600 rounded-lg hover:bg-sky-100 transition-all"
                >
                  <Plus className="h-3.5 w-3.5" /> Avaliar Competências
                </button>
              </div>
            ) : (
              <div className="space-y-3 max-h-[300px] overflow-y-auto pr-1">
                {companyEvaluations.map((ev) => {
                  const levelInfo = getMaturityLevel(ev.pontuacaoTotal);
                  const isSelected = selectedEval?.id === ev.id;
                  
                  return (
                    <div
                      key={ev.id}
                      onClick={() => handleSelectEval(ev)}
                      className={`p-3.5 rounded-2xl border transition-all cursor-pointer flex justify-between items-center ${
                        isSelected 
                          ? 'border-sky-200 bg-sky-50/40 shadow-sm' 
                          : 'border-slate-100 bg-slate-50/50 hover:bg-slate-50'
                      }`}
                    >
                      <div className="space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="text-sm font-bold text-slate-800">{ev.pontuacaoTotal} pts</span>
                          <span className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${levelInfo.color.split(' ')[0]} ${levelInfo.color.split(' ')[1]}`}>
                            {ev.nivelMaturidade}
                          </span>
                        </div>
                        <p className="text-xs text-slate-400 flex items-center gap-1">
                          <Calendar className="h-3 w-3 text-slate-300" />
                          {new Date(ev.dataAvaliacao).toLocaleDateString('pt-BR')}
                        </p>
                      </div>
                      
                      <button
                        onClick={(e) => handleDeleteEval(ev.id!, e)}
                        className="p-1.5 text-slate-300 hover:text-rose-500 hover:bg-white rounded-lg transition-all"
                        title="Excluir avaliação"
                      >
                        <Trash2 className="h-3.5 w-3.5" />
                      </button>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* Escala de Avaliação Guide Card (exactly as in the left column of the image) */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
            <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider border-b border-slate-50 pb-2">
              Escala de Avaliação (1 a 5)
            </h3>
            <div className="space-y-3">
              {[
                { lvl: 5, label: "EXCELÊNCIA", desc: "Competência é um ponto forte evidente, servindo de modelo para outros e impulsionando o sucesso do negócio.", color: "bg-emerald-600 text-white" },
                { lvl: 4, label: "AVANÇADO", desc: "Competência bem desenvolvida e aplicada de forma consistente, gerando resultados positivos.", color: "bg-emerald-500 text-white" },
                { lvl: 3, label: "INTERMEDIÁRIO", desc: "Competência demonstrada com regularidade, mas ainda há espaço para melhorias significativas.", color: "bg-amber-400 text-white" },
                { lvl: 2, label: "BÁSICO", desc: "Competência demonstrada de forma inconsistente. Há esforço, mas os resultados ainda são limitados.", color: "bg-amber-500 text-white" },
                { lvl: 1, label: "INICIANTE", desc: "Competência raramente demonstrada. Há grande dificuldade em aplicá-la no dia a dia.", color: "bg-rose-500 text-white" },
              ].map((item) => (
                <div key={item.lvl} className="flex gap-3 text-xs leading-relaxed">
                  <div className={`h-6 w-6 rounded-full flex items-center justify-center shrink-0 font-bold text-xs ${item.color}`}>
                    {item.lvl}
                  </div>
                  <div>
                    <h4 className="font-bold text-slate-700">{item.label}</h4>
                    <p className="text-slate-500 text-[11px]">{item.desc}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Interpretação dos Resultados Score limits card */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
            <h3 className="text-sm font-bold text-slate-800 uppercase tracking-wider border-b border-slate-50 pb-2">
              Faixas de Resultados
            </h3>
            <div className="space-y-2 text-xs">
              <div className="p-2 bg-emerald-50 border-l-4 border-emerald-500 rounded-r-lg">
                <span className="font-bold text-emerald-800">46 a 50 pontos | AVANÇADO</span>
                <p className="text-slate-500 text-[10px] mt-0.5">Excelência ampla em governança e gestão.</p>
              </div>
              <div className="p-2 bg-blue-50 border-l-4 border-blue-500 rounded-r-lg">
                <span className="font-bold text-blue-800">36 a 45 pontos | INTERMEDIÁRIO</span>
                <p className="text-slate-500 text-[10px] mt-0.5">Base sólida com pontos focais de ajuste.</p>
              </div>
              <div className="p-2 bg-amber-50 border-l-4 border-amber-500 rounded-r-lg">
                <span className="font-bold text-amber-800">21 a 35 pontos | BÁSICO</span>
                <p className="text-slate-500 text-[10px] mt-0.5">Gestão reativa e instabilidade comportamental.</p>
              </div>
              <div className="p-2 bg-rose-50 border-l-4 border-rose-500 rounded-r-lg">
                <span className="font-bold text-rose-800">10 a 20 pontos | INICIAL</span>
                <p className="text-slate-500 text-[10px] mt-0.5">Alto risco e urgência severa de reestruturação.</p>
              </div>
            </div>
          </div>

        </div>

        {/* RIGHT COLUMN: INTERACTIVE TABS CONTENT */}
        <div className="lg:col-span-8 space-y-6">

          <AnimatePresence mode="wait">
            
            {/* TAB 1: HISTORY / DETAILED ASSESSMENT PREVIEW */}
            {activeTab === 'history' && (
              <motion.div
                key="history"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                className="space-y-6"
              >
                {!selectedEval ? (
                  <div className="bg-white rounded-[2rem] p-12 border border-slate-100 text-center shadow-sm">
                    <div className="h-16 w-16 bg-slate-50 text-sky-600 rounded-2xl flex items-center justify-center mx-auto mb-6">
                      <TrendingUp className="h-8 w-8" />
                    </div>
                    <h3 className="text-lg font-bold text-slate-800">Nenhum diagnóstico de competência selecionado</h3>
                    <p className="mt-2 text-slate-500 text-sm max-w-sm mx-auto">
                      Selecione um diagnóstico da lista histórica ao lado ou crie uma nova avaliação para este cliente agora.
                    </p>
                    <button
                      onClick={() => setActiveTab('new-assessment')}
                      className="mt-6 inline-flex items-center gap-2 px-5 py-2.5 bg-sky-600 text-white rounded-xl text-sm font-semibold hover:bg-sky-700 shadow-md transition-all"
                    >
                      <Plus className="h-4 w-4" /> Registrar Nova Avaliação
                    </button>
                  </div>
                ) : (
                  <div className="space-y-6">
                    {/* Visual Container representing the exact design components from image */}
                    <div id="maturity-report-container" className="bg-white rounded-[2rem] p-8 border border-slate-100 shadow-sm space-y-8">
                      
                      {/* Internal Header */}
                      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between border-b border-slate-100 pb-5 gap-3">
                        <div>
                          <span className="text-[10px] font-bold text-sky-600 uppercase tracking-widest bg-sky-50 px-2.5 py-1 rounded-full">
                            Diagnóstico de Maturidade Empresarial
                          </span>
                          <h2 className="text-xl font-extrabold text-slate-900 mt-1">
                            Avaliação das Competências de {currentEmpresaName}
                          </h2>
                          <p className="text-xs text-slate-400 mt-1">
                            Realizado em: {new Date(selectedEval.dataAvaliacao).toLocaleDateString('pt-BR')} às {new Date(selectedEval.dataAvaliacao).toLocaleTimeString('pt-BR', {hour: '2-digit', minute:'2-digit'})}
                          </p>
                        </div>
                        
                        <div className="flex gap-2 shrink-0">
                          <button
                            onClick={handleExportPDF}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl transition-all"
                          >
                            <Download className="h-3.5 w-3.5" /> PDF
                          </button>
                        </div>
                      </div>

                      {/* Top Metrics row mirroring the exact visual layout from the image uploaded */}
                      <div className="grid grid-cols-1 md:grid-cols-12 gap-4">
                        
                        {/* total points box */}
                        <div className="md:col-span-3 bg-slate-50 border border-slate-150 rounded-2xl p-4 text-center flex flex-col justify-between relative overflow-hidden">
                          <span className="text-[10px] font-extrabold tracking-wider text-slate-500 uppercase">
                            Pontuação Total
                          </span>
                          <div className="my-2">
                            <span className="text-4xl font-extrabold text-slate-900 tracking-tight">
                              {selectedEval.pontuacaoTotal}
                            </span>
                            <span className="text-xs font-semibold text-slate-400 ml-1">de 50</span>
                          </div>
                          <span className="text-[10px] font-bold text-slate-400">Escala geral</span>
                        </div>

                        {/* maturity level box */}
                        <div className={`md:col-span-4 rounded-2xl border p-4 text-center flex flex-col justify-between ${getMaturityLevel(selectedEval.pontuacaoTotal).color}`}>
                          <span className="text-[10px] font-extrabold tracking-wider uppercase opacity-80">
                            Nível de Maturidade
                          </span>
                          <div className="my-2 font-black text-xl tracking-wider uppercase">
                            {selectedEval.nivelMaturidade}
                          </div>
                          <span className="text-[10px] font-bold opacity-75">
                            ({getMaturityLevel(selectedEval.pontuacaoTotal).range})
                          </span>
                        </div>

                        {/* interpretation text box */}
                        <div className="md:col-span-5 bg-sky-50/20 border border-sky-100 rounded-2xl p-4 flex gap-3">
                          <Target className="h-6 w-6 text-sky-500 shrink-0 mt-0.5" />
                          <div className="space-y-1 text-left">
                            <span className="text-[10px] font-extrabold tracking-wider text-sky-700 uppercase">
                              Interpretação
                            </span>
                            <p className="text-xs text-slate-600 font-medium leading-relaxed">
                              {getMaturityLevel(selectedEval.pontuacaoTotal).interpretation}
                            </p>
                          </div>
                        </div>

                      </div>

                      {/* Main split display: Radar on left, detailed scores list on right */}
                      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-center pt-2">
                        
                        {/* Radar Chart Display */}
                        <div className="lg:col-span-6 flex flex-col items-center justify-center space-y-4">
                          <div className="text-center">
                            <h4 className="text-xs font-extrabold tracking-wider text-slate-500 uppercase mb-1">
                              Radar das Competências
                            </h4>
                            <p className="text-[10px] text-slate-400">Visualização das 10 dimensões avaliadas</p>
                          </div>
                          
                          <div className="w-full max-w-[340px] aspect-square flex items-center justify-center">
                            {drawMaturityRadar(selectedEval.scores)}
                          </div>
                        </div>

                        {/* Exact Individual Competency Scores side breakdown */}
                        <div className="lg:col-span-6 space-y-3">
                          <h4 className="text-xs font-extrabold tracking-wider text-slate-500 uppercase pb-1 border-b border-slate-100">
                            Pontuações Individuais
                          </h4>
                          
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                            {COMPETENCIES.map((comp) => {
                              const scoreVal = selectedEval.scores[comp.id as keyof MaturityEvaluation['scores']] || 3;
                              
                              // Determine color class based on score value (1-2: red, 3: amber, 4-5: green)
                              let textClr = "text-rose-600 bg-rose-50 border-rose-100";
                              if (scoreVal === 3) textClr = "text-amber-600 bg-amber-50 border-amber-100";
                              if (scoreVal >= 4) textClr = "text-emerald-600 bg-emerald-50 border-emerald-100";

                              return (
                                <div key={comp.id} className="flex justify-between items-center p-2 rounded-xl bg-slate-50 border border-slate-100 text-xs">
                                  <span className="font-bold text-slate-700 tracking-tight flex items-center gap-1.5 truncate">
                                    <span>{comp.icon}</span>
                                    <span className="truncate">{comp.name.split('. ')[1]}</span>
                                  </span>
                                  <span className={`px-2 py-0.5 rounded-lg font-extrabold ${textClr}`}>
                                    {scoreVal.toFixed(1)}
                                  </span>
                                </div>
                              );
                            })}
                          </div>

                          {selectedEval.observacoes && (
                            <div className="mt-4 p-4 bg-slate-50 rounded-2xl border border-slate-100 text-xs">
                              <span className="font-bold text-slate-700 block mb-1">Observações do Consultor:</span>
                              <p className="text-slate-500 italic leading-relaxed">{selectedEval.observacoes}</p>
                            </div>
                          )}
                        </div>

                      </div>

                    </div>

                    {/* ACTION PLAN ALIGNMENT & EVOLUTION REPORT (Direct answer to prompt!) */}
                    <div className="bg-white rounded-[2rem] p-8 border border-slate-100 shadow-sm space-y-6">
                      <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                        <CheckCircle2 className="h-6 w-6 text-emerald-500 shrink-0" />
                        <div>
                          <h3 className="text-lg font-bold text-slate-800">
                            Execução do Plano de Ação & Resultados
                          </h3>
                          <p className="text-xs text-slate-400">
                            Como a implantação das tarefas do plano está impulsionando a maturidade corporativa.
                          </p>
                        </div>
                      </div>

                      {/* Summary Stat Rows */}
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                        
                        <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 flex items-center gap-4">
                          <div className="h-12 w-12 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center shrink-0">
                            <Clock className="h-6 w-6" />
                          </div>
                          <div>
                            <span className="text-[10px] font-extrabold text-slate-400 block uppercase">Progresso do Plano</span>
                            <span className="text-xl font-bold text-slate-800">{taskProgress}% concluído</span>
                            <p className="text-[10px] text-slate-400 mt-0.5">{completedTasks} de {totalTasks} tarefas feitas</p>
                          </div>
                        </div>

                        <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 flex items-center gap-4">
                          <div className="h-12 w-12 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center shrink-0">
                            <TrendingUp className="h-6 w-6" />
                          </div>
                          <div>
                            <span className="text-[10px] font-extrabold text-slate-400 block uppercase">Evolução Geral</span>
                            {companyEvaluations.length > 1 ? (
                              <div>
                                <span className="text-xl font-bold text-emerald-600">
                                  +{companyEvaluations[0].pontuacaoTotal - companyEvaluations[companyEvaluations.length - 1].pontuacaoTotal} pontos
                                </span>
                                <p className="text-[10px] text-slate-400 mt-0.5">Desde a primeira avaliação</p>
                              </div>
                            ) : (
                              <div>
                                <span className="text-sm font-bold text-slate-500">Aguardando histórico</span>
                                <p className="text-[10px] text-slate-400 mt-0.5">Faça novos testes periodicamente</p>
                              </div>
                            )}
                          </div>
                        </div>

                        <div className="p-4 bg-sky-600 text-white rounded-2xl flex items-center gap-4">
                          <div className="h-12 w-12 bg-sky-500 text-white rounded-xl flex items-center justify-center shrink-0">
                            <Sparkles className="h-6 w-6" />
                          </div>
                          <div>
                            <span className="text-[10px] font-extrabold text-sky-200 block uppercase">Próximo Alvo</span>
                            <span className="text-sm font-bold">Aprimorar pontos fracos</span>
                            <p className="text-[10px] text-sky-100 mt-0.5">Veja as ações sugeridas abaixo</p>
                          </div>
                        </div>

                      </div>

                      {/* RECOMMENDATIONS ALIGNED TO CRITICAL SCORES (< 4) */}
                      <div className="space-y-4 pt-2">
                        <h4 className="text-xs font-extrabold tracking-wider text-slate-500 uppercase pb-2 border-b border-slate-100 flex items-center justify-between">
                          <span>Ações Corretivas Recomendadas baseadas nos pontos fracos</span>
                          <span className="text-[10px] bg-rose-50 text-rose-600 font-bold px-2 py-0.5 rounded-full lowercase">
                            com pontuação menor que 4.0
                          </span>
                        </h4>

                        <div className="space-y-3">
                          {COMPETENCIES.filter(c => (selectedEval.scores[c.id as keyof MaturityEvaluation['scores']] || 3) < 4).length === 0 ? (
                            <div className="text-center py-4 bg-emerald-50 rounded-2xl border border-emerald-100">
                              <CheckCircle2 className="h-8 w-8 text-emerald-600 mx-auto mb-2" />
                              <p className="text-xs font-bold text-emerald-800">Parabéns! Todas as competências do empresário estão em nível Avançado (4.0+)!</p>
                            </div>
                          ) : (
                            COMPETENCIES.filter(c => (selectedEval.scores[c.id as keyof MaturityEvaluation['scores']] || 3) < 4).map((comp) => {
                              const currentScore = selectedEval.scores[comp.id as keyof MaturityEvaluation['scores']];
                              const recommendations = RECOMMENDATIONS_DB[comp.id] || [];
                              
                              return (
                                <div key={comp.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-3">
                                  <div className="flex justify-between items-center">
                                    <h5 className="text-xs font-extrabold text-slate-800 flex items-center gap-2">
                                      <span>{comp.icon}</span>
                                      <span>{comp.name}</span>
                                      <span className="text-[10px] bg-amber-100 text-amber-700 px-2 py-0.5 rounded-full font-bold">
                                        Pontuação: {currentScore.toFixed(1)}
                                      </span>
                                    </h5>
                                  </div>
                                  
                                  <div className="space-y-2.5">
                                    {recommendations.map((recText, recIdx) => {
                                      const isAdded = addedTasksKeys[comp.id + recText];
                                      return (
                                        <div key={recIdx} className="flex flex-col sm:flex-row sm:items-center sm:justify-between p-3 bg-white rounded-xl border border-slate-150 gap-3 text-xs">
                                          <p className="text-slate-600 leading-relaxed max-w-[85%]">{recText}</p>
                                          <button
                                            disabled={isAdded}
                                            onClick={() => handleAddRecommendation(comp.name.split('. ')[1], recText, comp.id)}
                                            className={`shrink-0 inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-extrabold transition-all ${
                                              isAdded 
                                                ? "bg-emerald-50 text-emerald-600" 
                                                : "bg-sky-50 hover:bg-sky-100 text-sky-600 hover:shadow-sm"
                                            }`}
                                          >
                                            {isAdded ? (
                                              <>
                                                <Check className="h-3.5 w-3.5" /> Adicionado
                                              </>
                                            ) : (
                                              <>
                                                <Plus className="h-3.5 w-3.5" /> Vincular ao Plano
                                              </>
                                            )}
                                          </button>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>
                              );
                            })
                          )}
                        </div>
                      </div>

                    </div>
                  </div>
                )}
              </motion.div>
            )}

            {/* TAB 2: REGISTER NEW ASSESSMENT */}
            {activeTab === 'new-assessment' && (
              <motion.div
                key="new-assessment"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                className="bg-white rounded-[2rem] p-8 border border-slate-100 shadow-sm space-y-8"
              >
                <div>
                  <h2 className="text-2xl font-extrabold text-slate-800">
                    Novo Diagnóstico de Competências
                  </h2>
                  <p className="text-sm text-slate-500 mt-1">
                    Atribua uma nota de 1 a 5 para as competências essenciais de gestão para a empresa <strong className="text-sky-600">{currentEmpresaName}</strong>.
                  </p>
                </div>

                {/* Grid of the 10 dimensions sliders */}
                <div className="space-y-6">
                  {COMPETENCIES.map((comp) => {
                    const currentScore = wizardScores[comp.id] || 3;
                    
                    return (
                      <div key={comp.id} className="p-5 bg-slate-50 rounded-2xl border border-slate-100 space-y-4">
                        <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                          <div className="flex items-start gap-2.5">
                            <span className="text-xl p-1 bg-white rounded-lg shadow-sm">{comp.icon}</span>
                            <div>
                              <h4 className="font-bold text-slate-800 text-sm">{comp.name}</h4>
                              <p className="text-xs text-slate-400">{comp.description}</p>
                            </div>
                          </div>
                          
                          {/* Current Value Display */}
                          <div className="shrink-0 flex items-center gap-1.5 self-start sm:self-center">
                            <span className="text-xs text-slate-400 font-bold uppercase">Nota:</span>
                            <span className={`h-8 w-14 rounded-xl flex items-center justify-center font-black text-sm border shadow-sm ${
                              currentScore <= 2 
                                ? "bg-rose-50 border-rose-200 text-rose-600" 
                                : currentScore === 3
                                  ? "bg-amber-50 border-amber-200 text-amber-600"
                                  : "bg-emerald-50 border-emerald-200 text-emerald-600"
                            }`}>
                              {currentScore.toFixed(1)}
                            </span>
                          </div>
                        </div>

                        {/* Question Text */}
                        <p className="text-xs text-slate-600 leading-relaxed bg-white p-3 rounded-xl border border-slate-100 italic">
                          "{comp.question}"
                        </p>

                        {/* Interactive Sliders/Clickable boxes */}
                        <div className="flex flex-col sm:flex-row sm:items-center gap-4">
                          <input
                            type="range"
                            min="1"
                            max="5"
                            step="1"
                            value={currentScore}
                            onChange={(e) => setWizardScores(prev => ({ ...prev, [comp.id]: parseInt(e.target.value) }))}
                            className="w-full accent-sky-600 h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer"
                          />
                          
                          {/* Colored Buttons alternative selection for rapid tapping */}
                          <div className="flex justify-between w-full sm:w-auto gap-2">
                            {[1, 2, 3, 4, 5].map((lvl) => (
                              <button
                                key={lvl}
                                type="button"
                                onClick={() => setWizardScores(prev => ({ ...prev, [comp.id]: lvl }))}
                                className={`h-8 w-8 rounded-lg font-bold text-xs transition-all border shrink-0 ${
                                  currentScore === lvl
                                    ? "bg-sky-600 border-sky-600 text-white shadow-md scale-105"
                                    : "bg-white hover:bg-slate-100 border-slate-200 text-slate-600"
                                }`}
                              >
                                {lvl}
                              </button>
                            ))}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Additional notes */}
                <div className="space-y-2">
                  <label className="text-sm font-bold text-slate-700 block">Observações do Consultor (Opcional)</label>
                  <textarea
                    rows={3}
                    placeholder="Anote pontos de destaque observados durante o teste de maturidade, histórico de reações, nível de honestidade, etc..."
                    value={wizardObservacoes}
                    onChange={(e) => setWizardObservacoes(e.target.value)}
                    className="w-full p-4 border border-slate-200 rounded-2xl text-sm focus:ring-2 focus:ring-sky-500 focus:border-sky-500 bg-slate-50 focus:bg-white transition-all"
                  />
                </div>

                {/* Bottom Row action buttons */}
                <div className="flex justify-between border-t border-slate-100 pt-6">
                  <button
                    type="button"
                    onClick={() => setActiveTab('history')}
                    className="px-5 py-2.5 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-600 text-sm font-semibold transition-all"
                  >
                    Voltar ao Diagnóstico
                  </button>
                  <button
                    type="button"
                    onClick={handleSubmitAssessment}
                    disabled={savingAssessment}
                    className="px-6 py-2.5 bg-sky-600 hover:bg-sky-700 text-white font-semibold text-sm rounded-xl shadow-md transition-all flex items-center gap-2"
                  >
                    {savingAssessment ? (
                      <>
                        <RefreshCw className="h-4 w-4 animate-spin" /> Salvando...
                      </>
                    ) : (
                      <>
                        <Check className="h-4 w-4" /> Finalizar e Salvar Diagnóstico
                      </>
                    )}
                  </button>
                </div>
              </motion.div>
            )}

            {/* TAB 3: CHRONOLOGICAL EVOLUTION PROGRESS TRACKER (Direct answer to prompt!) */}
            {activeTab === 'evolution' && (
              <motion.div
                key="evolution"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -15 }}
                className="bg-white rounded-[2rem] p-8 border border-slate-100 shadow-sm space-y-8"
              >
                <div>
                  <h2 className="text-2xl font-extrabold text-slate-800">
                    Histórico de Evolução de {currentEmpresaName}
                  </h2>
                  <p className="text-sm text-slate-500 mt-1">
                    Acompanhe o crescimento e evolução das notas do empresário à medida que o plano de ação avança em execução.
                  </p>
                </div>

                {companyEvaluations.length < 2 ? (
                  <div className="p-8 text-center bg-slate-50 rounded-2xl border border-slate-100 space-y-3">
                    <AlertCircle className="h-10 w-10 text-amber-500 mx-auto" />
                    <h4 className="font-bold text-slate-800">Histórico insuficiente</h4>
                    <p className="text-xs text-slate-500 max-w-sm mx-auto">
                      Para traçar um gráfico de evolução cronológico, são necessários pelo menos **2 diagnósticos de competências** registrados.
                    </p>
                    <button
                      onClick={() => setActiveTab('new-assessment')}
                      className="mt-2 inline-flex items-center gap-1 bg-sky-600 text-white px-4 py-2 rounded-xl text-xs font-bold shadow hover:bg-sky-700 transition-all"
                    >
                      <Plus className="h-3.5 w-3.5" /> Registrar Segundo Diagnóstico
                    </button>
                  </div>
                ) : (
                  <div className="space-y-8">
                    
                    {/* Visual Line Graph Representation constructed of pure CSS & HTML for absolute stability */}
                    <div className="p-6 bg-slate-50 rounded-[2rem] border border-slate-100 space-y-4">
                      <h3 className="text-xs font-extrabold tracking-wider text-slate-400 uppercase text-center">
                        Gráfico de Pontuação de Competência ao Longo do Tempo
                      </h3>
                      
                      <div className="h-[250px] flex items-end justify-between px-4 sm:px-12 pt-8 relative border-b border-slate-200">
                        {/* Target Grid lines */}
                        <div className="absolute inset-0 flex flex-col justify-between text-[10px] text-slate-300 font-extrabold pointer-events-none pb-6">
                          <div className="border-b border-slate-100 w-full pt-2">50 pts (Excelente)</div>
                          <div className="border-b border-slate-100 w-full pt-2">40 pts (Avançado)</div>
                          <div className="border-b border-slate-100 w-full pt-2">30 pts (Intermediário)</div>
                          <div className="border-b border-slate-100 w-full pt-2">20 pts (Básico)</div>
                          <div className="border-b border-slate-100 w-full pt-2">10 pts (Inicial)</div>
                        </div>

                        {/* Drawing columns matching chronologically sorted (ascending) list */}
                        {[...companyEvaluations].reverse().map((ev, idx, arr) => {
                          // Height calculation (based on score out of 50)
                          const heightPct = (ev.pontuacaoTotal / 50) * 100;
                          
                          return (
                            <div key={ev.id} className="flex flex-col items-center flex-1 z-10 relative group">
                              {/* Popup display of exact points */}
                              <span className="absolute -top-6 bg-sky-600 text-white text-[10px] font-bold px-2 py-0.5 rounded-md shadow opacity-0 group-hover:opacity-100 transition-opacity">
                                {ev.pontuacaoTotal} pts
                              </span>

                              {/* Styled vertical line/bar representing progress */}
                              <div 
                                style={{ height: `${heightPct * 1.8}px` }} 
                                className="w-4 bg-sky-500 hover:bg-sky-600 rounded-t-lg transition-all shadow-sm flex items-end justify-center"
                              >
                                <span className="text-[9px] font-bold text-white mb-1.5 hidden sm:inline">{ev.pontuacaoTotal}</span>
                              </div>

                              {/* Label on the bottom */}
                              <span className="text-[10px] font-bold text-slate-500 mt-2">
                                #{idx + 1}
                              </span>
                              <span className="text-[8px] text-slate-400">
                                {new Date(ev.dataAvaliacao).toLocaleDateString('pt-BR', {day: 'numeric', month: 'short'})}
                              </span>
                            </div>
                          );
                        })}
                      </div>

                      <p className="text-[10px] text-slate-400 text-center italic">
                        *Passe o mouse por cima das colunas para visualizar as notas pontuais.
                      </p>
                    </div>

                    {/* Correlation with Action Plan Execution (Direct answer to prompt!) */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
                      
                      <div className="p-6 bg-slate-50 rounded-[2rem] border border-slate-100 space-y-4">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="h-5 w-5 text-emerald-500" />
                          <h4 className="font-bold text-slate-800 text-sm">Status do Plano de Ação</h4>
                        </div>
                        <p className="text-xs text-slate-500 leading-relaxed">
                          À medida que o empresário executa as tarefas acordadas, as competências como **Capacidade de Execução**, **Planejamento** e **Organização** tendem a aumentar exponencialmente.
                        </p>
                        
                        <div className="space-y-2">
                          <div className="flex justify-between text-xs font-bold text-slate-600">
                            <span>Progresso Geral</span>
                            <span>{taskProgress}%</span>
                          </div>
                          <div className="w-full bg-slate-200 rounded-full h-2.5">
                            <div 
                              style={{ width: `${taskProgress}%` }} 
                              className="bg-emerald-500 h-2.5 rounded-full transition-all duration-500"
                            />
                          </div>
                          <div className="flex justify-between text-[10px] text-slate-400 font-semibold">
                            <span>{completedTasks} Concluídas</span>
                            <span>{totalTasks - completedTasks} Pendentes</span>
                          </div>
                        </div>
                      </div>

                      <div className="p-6 bg-sky-50/20 border border-sky-100 rounded-[2rem] space-y-3">
                        <div className="flex items-center gap-2">
                          <Sparkles className="h-5 w-5 text-sky-600" />
                          <h4 className="font-bold text-slate-800 text-sm">Análise de Evolução x Execução</h4>
                        </div>
                        
                        <div className="text-xs text-slate-600 space-y-2.5 leading-relaxed">
                          <p>
                            Atualmente, o plano de ação deste cliente está com <strong className="text-sky-600">{taskProgress}% de conclusão</strong>.
                          </p>
                          <p>
                            No histórico avaliado, nota-se que a pontuação de maturidade empresarial partiu de <strong className="text-slate-800">{companyEvaluations[companyEvaluations.length - 1].pontuacaoTotal} pontos</strong> e evoluiu para <strong className="text-sky-600">{companyEvaluations[0].pontuacaoTotal} pontos</strong> no último diagnóstico.
                          </p>
                          <p className="font-bold text-emerald-600 flex items-center gap-1 bg-emerald-50 p-2 rounded-xl border border-emerald-100 mt-2">
                            <TrendingUp className="h-4 w-4 shrink-0" />
                            Crescimento líquido de +{companyEvaluations[0].pontuacaoTotal - companyEvaluations[companyEvaluations.length - 1].pontuacaoTotal} pontos de maturidade do empresário!
                          </p>
                        </div>
                      </div>

                    </div>

                  </div>
                )}
              </motion.div>
            )}

          </AnimatePresence>

        </div>

      </div>

      {/* Bottom informational guidance banner */}
      <div className="bg-slate-50 rounded-3xl p-5 border border-slate-200 flex items-start gap-3 text-xs leading-relaxed max-w-4xl mx-auto">
        <Info className="h-5 w-5 text-sky-500 shrink-0 mt-0.5" />
        <div className="text-slate-500">
          <strong className="text-slate-700">Forma de Aplicação:</strong> O empresário deve ler atenciosamente a descrição de cada critério e atribuir uma nota sincera de 1 a 5, refletindo com total honestidade sobre o seu comportamento atual no dia a dia do negócio. O consultor deve facilitar essa reflexão fazendo perguntas norteadoras cruciais que ajudem o empresário a ponderar seu real desempenho prático e gerir o avanço das competências com base no progresso das tarefas do plano de ação.
        </div>
      </div>
    </div>
  );
}
