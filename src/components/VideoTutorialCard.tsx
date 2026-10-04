import React, { useState, useEffect, useRef } from 'react';
import { 
  Play, 
  Pause, 
  RotateCcw, 
  Volume2, 
  VolumeX, 
  Sparkles, 
  CheckCircle2, 
  Layers, 
  BarChart3, 
  FileSpreadsheet, 
  Bot, 
  ChevronRight, 
  ChevronLeft, 
  Tv, 
  ListOrdered, 
  Lightbulb, 
  Building2, 
  Calendar, 
  Settings, 
  CalendarDays, 
  TrendingUp, 
  CheckSquare, 
  FileText, 
  ShieldCheck, 
  Layout, 
  Home,
  X 
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '../lib/utils';
import { Button } from './Button';

export interface TutorialStep {
  id: string;
  badge: string;
  title: string;
  narration: string;
  subtext: string;
  duration: number; // segundos estimados
  highlights: string[];
  icon: 'home' | 'welcome' | 'company' | 'diagnosis' | 'actionPlan' | 'report' | 'ai' | 'agenda' | 'credenciadas' | 'maturity' | 'resultado' | 'macro' | 'premises' | 'settings' | 'licenses';
  targetView?: string;
  sceneVisual: {
    tag: string;
    headline: string;
    metrics: { label: string; value: string; color: string }[];
    mockDetails: string[];
  };
}

export const TUTORIAL_STEPS: TutorialStep[] = [
  {
    id: 'tab_home',
    badge: 'Aba 1 • Início',
    title: 'Aba Início: Visão Geral e Painel de Controle',
    narration: 'A primeira aba do sistema é a aba Início. Nela você tem a visão geral do seu ambiente de consultoria, o status dos atendimentos e atalhos rápidos para cadastrar clientes, abrir diagnósticos e acessar o macro dashboard. É o seu ponto de partida para navegar com facilidade em todos os recursos do ConsultorPro.',
    subtext: 'Painel inicial com status do sistema, atalhos de navegação rápida e resumo executivo.',
    duration: 20,
    highlights: [
      'Painel de acesso rápido para todos os módulos e ferramentas do sistema',
      'Identificação da Credenciada Sebrae e consultor titular responsável',
      'Atalhos instantâneos para novo cliente, novo diagnóstico ou agenda'
    ],
    icon: 'home',
    targetView: 'home',
    sceneVisual: {
      tag: 'ABA 1 • INÍCIO',
      headline: 'Painel Geral & Ponto de Partida',
      metrics: [
        { label: 'Navegação', value: '12 Abas', color: 'text-amber-400' },
        { label: 'Acesso', value: 'Instantâneo', color: 'text-sky-400' },
        { label: 'Ambiente', value: 'Integrado', color: 'text-emerald-400' }
      ],
      mockDetails: [
        'Visão geral das empresas atendidas e atalhos de acesso rápido',
        'Carimbo com dados da credenciada e consultor titular ativo',
        'Acesso direto para Clientes, Projetos, Macro Dashboard e Agenda'
      ]
    }
  },
  {
    id: 'tab_companies',
    badge: 'Aba 2 • Clientes',
    title: 'Aba Clientes: Cadastro e Busca por CNPJ',
    narration: 'A segunda aba é a aba Clientes, onde você gerencia todas as empresas atendidas. O sistema conta com busca automática de CNPJ pela Receita Federal, preenchendo Razão Social, CNAE, porte e endereço num instante. Nela você também cadastra contatos, faturamento atual e metas de crescimento financeiro.',
    subtext: 'Gestão da carteira de empresas, consulta pública de CNPJ e dados cadastrais.',
    duration: 21,
    highlights: [
      'Preenchimento automático instantâneo com busca de CNPJ na Receita Federal',
      'Histórico centralizado de diagnósticos e ciclos de atendimento por cliente',
      'Registro de contatos estratégicos, porte e metas de faturamento anual'
    ],
    icon: 'company',
    targetView: 'companies',
    sceneVisual: {
      tag: 'ABA 2 • CLIENTES',
      headline: 'Carteira de Empresas & Dados Cadastrais',
      metrics: [
        { label: 'Busca CNPJ', value: 'Receita Federal', color: 'text-sky-400' },
        { label: 'Histórico', value: 'Centralizado', color: 'text-emerald-400' },
        { label: 'Faturamento', value: 'Metas & Real', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Preenchimento automático de Razão Social, CNAE e Endereço comercial',
        'Vínculo direto com a consultoria credenciada e consultor titular',
        'Controle de contatos, porte da empresa e faturamento planejado'
      ]
    }
  },
  {
    id: 'tab_projects',
    badge: 'Aba 3 • Projetos',
    title: 'Aba Projetos: Ciclos e Replicação de Diagnósticos',
    narration: 'A terceira aba é a aba Projetos, o núcleo dos diagnósticos do cliente. Nela você cria novos ciclos de consultoria, acessa o questionário detalhado de premissas com cálculo automático de notas e problemas críticos, ou replica diagnósticos anteriores em um clique para comparar com exatidão a evolução da empresa.',
    subtext: 'Controle de ciclos de atendimento, criação e replicação de diagnósticos.',
    duration: 21,
    highlights: [
      'Criação e acompanhamento de ciclos de consultoria por empresa',
      'Replicação ágil de diagnósticos para análise evolutiva periódica',
      'Cálculo de conformidade de premissas e identificação de problemas prioritários'
    ],
    icon: 'diagnosis',
    targetView: 'projects',
    sceneVisual: {
      tag: 'ABA 3 • PROJETOS',
      headline: 'Ciclos de Diagnóstico & Evolução Contínua',
      metrics: [
        { label: 'Ciclos', value: 'Multiperíodo', color: 'text-sky-400' },
        { label: 'Replicação', value: '1 Clique', color: 'text-emerald-400' },
        { label: 'Avaliação', value: 'Premissas', color: 'text-indigo-400' }
      ],
      mockDetails: [
        'Listagem de diagnósticos com data, status e consultor responsável',
        'Atalho direto para preencher premissas e notas de cada pilar',
        'Replicação de diagnóstico para ciclos de acompanhamento sem retrabalho'
      ]
    }
  },
  {
    id: 'tab_agenda',
    badge: 'Aba 4 • Agenda',
    title: 'Aba Agenda: Visitas Técnicas e Reuniões',
    narration: 'A quarta aba é a Agenda do Consultor. Nela você planeja todas as visitas técnicas na empresa do cliente, reuniões virtuais e sessões de devolutiva. Os compromissos ficam associados ao cliente, assegurando pontualidade, controle da carga horária e cumprimento dos prazos contratuais.',
    subtext: 'Calendário e programação de visitas técnicas e reuniões estratégicas.',
    duration: 20,
    highlights: [
      'Visão organizada de compromissos por dia, semana e mês',
      'Associação de cada agendamento à empresa atendida',
      'Controle de visitas presenciais, reuniões remotas e prazos de entrega'
    ],
    icon: 'agenda',
    targetView: 'agenda',
    sceneVisual: {
      tag: 'ABA 4 • AGENDA',
      headline: 'Programação de Visitas e Reuniões',
      metrics: [
        { label: 'Visitas', value: 'Presenciais', color: 'text-amber-400' },
        { label: 'Sessões', value: 'Remotas', color: 'text-emerald-400' },
        { label: 'Controle', value: 'Por Cliente', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Planejamento de visitas com data, horário e pauta definida',
        'Lembretes de entregas intermediárias e sessões finais',
        'Sincronização com dados do cliente e consultor responsável'
      ]
    }
  },
  {
    id: 'tab_credenciadas',
    badge: 'Aba 5 • Credenciadas',
    title: 'Aba Credenciadas: Consultoria e Consultor Titular',
    narration: 'A quinta aba é a aba Credenciadas. Nela você cadastra as empresas de consultoria credenciadas ao Sebrae e os consultores responsáveis. Esses dados oficiais, com Razão Social, CNPJ e CPF do consultor titular, são carimbados de forma automática nos relatórios em PDF gerados pelo sistema.',
    subtext: 'Configuração dos dados da consultoria credenciada e corpo técnico oficial.',
    duration: 20,
    highlights: [
      'Cadastro de Razão Social, CNPJ e dados de contato da credenciada',
      'Vínculo do consultor titular, CPF e registro profissional',
      'Carimbo e assinatura automática no cabeçalho e rodapé dos PDFs'
    ],
    icon: 'credenciadas',
    targetView: 'credenciadas',
    sceneVisual: {
      tag: 'ABA 5 • CREDENCIADAS',
      headline: 'Dados Oficiais da Consultoria & Carimbo',
      metrics: [
        { label: 'Empresas', value: 'Credenciadas', color: 'text-indigo-400' },
        { label: 'Carimbo PDF', value: 'Automático', color: 'text-emerald-400' },
        { label: 'Padrão', value: 'Sebrae', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Armazenamento dos dados cadastrais da empresa prestadora',
        'Seleção da credenciada ativa para o atendimento corrente',
        'Integração transparente com cabeçalhos de relatórios executivos'
      ]
    }
  },
  {
    id: 'tab_maturity',
    badge: 'Aba 6 • Maturidade',
    title: 'Aba Maturidade: Radares e Competências de Gestão',
    narration: 'A sexta aba é a de Maturidade Empresarial. Ela avalia as dez competências vitais da empresa em áreas como Finanças, Mercado, Processos e Pessoas. Os gráficos de radar dinâmicos e scores percentuais mostram de forma visual e intuitiva os pontos fortes e as prioridades imediatas de melhoria para o empresário.',
    subtext: 'Indicadores gráficos de radar, pontuação percentual e análise por competências.',
    duration: 21,
    highlights: [
      'Gráfico de radar interativo com benchmark de maturidade',
      'Cálculo automático de porcentagem de conformidade por dimensão',
      'Diagnóstico visual de alto impacto para apresentar ao empresário'
    ],
    icon: 'maturity',
    targetView: 'maturity-assessment',
    sceneVisual: {
      tag: 'ABA 6 • MATURIDADE',
      headline: 'Radar de Gestão & Competências Vitais',
      metrics: [
        { label: 'Dimensões', value: '10 Áreas', color: 'text-violet-400' },
        { label: 'Gráfico', value: 'Radar Dinâmico', color: 'text-emerald-400' },
        { label: 'Score Geral', value: 'Percentual', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Pontuação calculada a partir das respostas do diagnóstico',
        'Identificação visual de gargalos em Finanças, Marketing, Operações e RH',
        'Exportação gráfica para apresentação estratégica ao cliente'
      ]
    }
  },
  {
    id: 'tab_resultado',
    badge: 'Aba 7 • Resultado',
    title: 'Aba Resultado: Comparativo Antes vs Depois',
    narration: 'A sétima aba é o Resultado da Consultoria. Nela você compara o diagnóstico inicial com a situação alcançada após as ações de melhoria. O sistema calcula a evolução com notas de zero a dez, aumento de maturidade e metas atingidas, gerando um laudo claro que comprova o valor da consultoria entregue.',
    subtext: 'Demonstrativo antes versus depois para comprovação de valor com laudo executivo.',
    duration: 21,
    highlights: [
      'Comparativo direto entre o status inicial e a situação após a consultoria',
      'Evidência quantitativa e qualitativa do impacto no negócio do cliente',
      'Notas de 0 a 10 e emissão de laudo executivo de comprovação'
    ],
    icon: 'resultado',
    targetView: 'resultado-consultoria',
    sceneVisual: {
      tag: 'ABA 7 • RESULTADO',
      headline: 'Evolução Antes vs. Depois da Consultoria',
      metrics: [
        { label: 'Evolução', value: 'Notas 0 a 10', color: 'text-emerald-400' },
        { label: 'Metas', value: 'Comprovadas', color: 'text-sky-400' },
        { label: 'Laudo', value: 'Executivo', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Gráficos comparativos de evolução entre o primeiro e o último ciclo',
        'Registro de ganhos operacionais, financeiros e mercadológicos',
        'Documento de conclusão com laudo pronto para o cliente'
      ]
    }
  },
  {
    id: 'tab_premises',
    badge: 'Aba 8 • Biblioteca',
    title: 'Aba Biblioteca: Perguntas, Problemas e Soluções',
    narration: 'A oitava aba é a Biblioteca de Perguntas, onde você gerencia toda a base metodológica do sistema. É possível editar perguntas e premissas por área, configurar a árvore de problemas e causas, e cadastrar soluções recomendadas para diferentes ramos de atividade e modelos de negócio.',
    subtext: 'Base de conhecimento customizável com premissas, problemas, soluções e nichos.',
    duration: 21,
    highlights: [
      'Edição de perguntas e premissas por área de gestão',
      'Configuração de problemas típicos e causas estruturais',
      'Cadastro de soluções sugeridas personalizadas por segmento'
    ],
    icon: 'premises',
    targetView: 'premises',
    sceneVisual: {
      tag: 'ABA 8 • BIBLIOTECA',
      headline: 'Base de Premissas, Problemas e Soluções',
      metrics: [
        { label: 'Premissas', value: 'Por Área', color: 'text-indigo-400' },
        { label: 'Problemas', value: 'Árvore de Causas', color: 'text-amber-400' },
        { label: 'Soluções', value: 'Por Segmento', color: 'text-emerald-400' }
      ],
      mockDetails: [
        'Seções organizadas: Problemas, Premissas, Soluções e Segmentos',
        'Inclusão de novos setores econômicos e perguntas específicas',
        'Importação, exportação e restauração dos padrões de fábrica'
      ]
    }
  },
  {
    id: 'tab_macro',
    badge: 'Aba 9 • Dashboard Macro',
    title: 'Aba Dashboard Macro: Inteligência da Carteira',
    narration: 'A nona aba é o Dashboard Macro, que consolida todas as empresas e consultorias em uma visão executiva única. Nela você analisa a média de maturidade de toda a carteira de clientes, a distribuição por setores e os problemas mais frequentes, ideal para relatórios gerenciais e prestação de contas.',
    subtext: 'Visão executiva em tempo real de toda a carteira de empresas atendidas.',
    duration: 20,
    highlights: [
      'Métricas consolidadas de todas as empresas e atendimentos ativos',
      'Distribuição por segmentos de mercado, porte e maturidade média',
      'Identificação de gargalos coletivos para planejamento estratégico'
    ],
    icon: 'macro',
    targetView: 'macro-dashboard',
    sceneVisual: {
      tag: 'ABA 9 • DASHBOARD MACRO',
      headline: 'Inteligência Global da Carteira de Clientes',
      metrics: [
        { label: 'Carteira', value: 'Consolidada', color: 'text-emerald-400' },
        { label: 'Maturidade Média', value: 'Tempo Real', color: 'text-sky-400' },
        { label: 'Setores', value: 'Multissegmento', color: 'text-purple-400' }
      ],
      mockDetails: [
        'Gráficos de barras e pizza com distribuição setorial',
        'Filtros por período, credenciada e consultor responsável',
        'Exportação de indicadores executivos para prestação de contas'
      ]
    }
  },
  {
    id: 'tab_action_plan',
    badge: 'Aba 10 • Plano de Ação',
    title: 'Aba Plano de Ação: Kanban e Cronograma de 34h',
    narration: 'A décima aba é o Plano de Ação. Nela o sistema distribui automaticamente as 8 atividades do cronograma padrão de 34 horas ao clicar em Gerar pelo Diagnóstico. Você controla as tarefas em formato Kanban, define responsáveis e prazos, e anexa fotos e documentos comprobatórios de cada etapa concluída.',
    subtext: 'Quadro Kanban com 8 atividades padrão, 34 horas e fotos de comprovação.',
    duration: 21,
    highlights: [
      'Geração automática das 8 atividades calibradas para 34 horas',
      'Atividades contextualizadas para o produto ou serviço do cliente',
      'Quadro Kanban com status, prazos e anexo de fotos comprobatórias'
    ],
    icon: 'actionPlan',
    targetView: 'kanban',
    sceneVisual: {
      tag: 'ABA 10 • PLANO DE AÇÃO',
      headline: 'Cronograma Kanban de 34h & Evidências',
      metrics: [
        { label: 'Carga Horária', value: '34 Horas', color: 'text-emerald-400' },
        { label: 'Atividades', value: '8 Etapas', color: 'text-sky-400' },
        { label: 'Evidências', value: 'Fotos e Anexos', color: 'text-indigo-400' }
      ],
      mockDetails: [
        'Distribuição automática de datas conforme o período da consultoria',
        'Controle de status: A Fazer, Em Andamento e Concluído',
        'Registro de evidências fotográficas para comprovação ao Sebrae'
      ]
    }
  },
  {
    id: 'tab_report',
    badge: 'Aba 11 • Relatório',
    title: 'Aba Relatório: Redação com IA e Exportação PDF',
    narration: 'A décima primeira aba é o Relatório de Consultoria. Nela a inteligência artificial Gemini auxilia na redação dos objetivos, justificativas e soluções recomendadas. Com apenas um clique você exporta o relatório formal completo em PDF oficial, com tabelas de horas, gráficos de maturidade e campos de assinatura prontos para entrega.',
    subtext: 'Emissão executiva com IA, tabelas de horas, evidências e layout oficial em PDF.',
    duration: 21,
    highlights: [
      'Redação assistida por IA Gemini adaptada ao segmento do cliente',
      'Consolidação de dados do cliente, credenciada e consultor titular',
      'Exportação para PDF oficial completo pronto para assinatura'
    ],
    icon: 'report',
    targetView: 'cronograma',
    sceneVisual: {
      tag: 'ABA 11 • RELATÓRIO',
      headline: 'Relatório Oficial em PDF & Redação IA',
      metrics: [
        { label: 'Formato', value: 'PDF Oficial', color: 'text-rose-400' },
        { label: 'Redação IA', value: 'Gemini Integrado', color: 'text-amber-400' },
        { label: 'Assinaturas', value: 'Formatadas', color: 'text-emerald-400' }
      ],
      mockDetails: [
        'Identificação oficial do cliente, credenciada e consultor',
        'Tabela das 8 atividades com horas e detalhamento de resultados',
        'Sumário executivo e gráficos de maturidade formatados para impressão'
      ]
    }
  },
  {
    id: 'tab_settings',
    badge: 'Aba 12 • Configurações',
    title: 'Aba Configurações: IA, Nuvem e Backups',
    narration: 'A décima segunda aba é a de Configurações. Nela você gerencia sua chave de API do Gemini, alterna entre armazenamento na nuvem com Firestore e salvamento local com criptografia, e pode baixar backups completos dos seus dados em arquivo JSON a qualquer momento, garantindo total segurança e controle.',
    subtext: 'Chave do Gemini, alternância Nuvem/Local e download de backups de segurança.',
    duration: 20,
    highlights: [
      'Configuração da chave de API do Google Gemini para recursos com IA',
      'Exportação e restauração de backups completos em arquivo JSON',
      'Alternância imediata entre Modo Nuvem e Modo Local criptografado'
    ],
    icon: 'settings',
    targetView: 'settings',
    sceneVisual: {
      tag: 'ABA 12 • CONFIGURAÇÕES',
      headline: 'Gestão de IA, Nuvem e Segurança',
      metrics: [
        { label: 'Backup', value: 'JSON Local', color: 'text-emerald-400' },
        { label: 'IA Gemini', value: 'Configurável', color: 'text-sky-400' },
        { label: 'Armazenamento', value: 'Nuvem ou Local', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Guia integrado para obter Chave de API gratuita do Gemini',
        'Sincronização em nuvem e histórico de backups locais',
        'Controle de logos e preferências gerais do sistema'
      ]
    }
  }
];

const renderIcon = (type: TutorialStep['icon']) => {
  switch (type) {
    case 'home':
    case 'welcome': return <Home size={16} className="text-amber-400" />;
    case 'company': return <Building2 size={16} className="text-sky-400" />;
    case 'diagnosis': return <FileSpreadsheet size={16} className="text-indigo-400" />;
    case 'agenda': return <CalendarDays size={16} className="text-amber-400" />;
    case 'credenciadas': return <Building2 size={16} className="text-emerald-400" />;
    case 'maturity': return <TrendingUp size={16} className="text-purple-400" />;
    case 'resultado': return <CheckSquare size={16} className="text-emerald-400" />;
    case 'macro': return <BarChart3 size={16} className="text-sky-400" />;
    case 'premises': return <FileText size={16} className="text-blue-400" />;
    case 'actionPlan': return <Layout size={16} className="text-emerald-400" />;
    case 'report': return <Calendar size={16} className="text-rose-400" />;
    case 'settings': return <Settings size={16} className="text-slate-400" />;
    case 'licenses': return <ShieldCheck size={16} className="text-amber-400" />;
    default: return <Sparkles size={16} className="text-sky-400" />;
  }
};

export const VideoTutorialCard: React.FC<{ 
  setView?: (view: any) => void;
  onClose?: () => void;
  isModal?: boolean;
}> = ({ setView, onClose, isModal = false }) => {
  const [currentStepIndex, setCurrentStepIndex] = useState(0);
  const [isPlaying, setIsPlaying] = useState(false);
  const [isMuted, setIsMuted] = useState(false);
  const [progress, setProgress] = useState(0);
  const [hasVoiceSupport, setHasVoiceSupport] = useState(true);
  const [speechActive, setSpeechActive] = useState(false);

  // Fallback seguro: garante que step nunca seja undefined
  const step = (TUTORIAL_STEPS && TUTORIAL_STEPS[currentStepIndex]) ? TUTORIAL_STEPS[currentStepIndex] : (TUTORIAL_STEPS && TUTORIAL_STEPS[0]);
  const timerRef = useRef<any>(null);
  const keepAliveTimerRef = useRef<any>(null);
  const speechFinishedRef = useRef(false);
  const utteranceRef = useRef<any>(null);

  // Inicializa suporte a voz e cancela ao desmontar
  useEffect(() => {
    try {
      if (typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis) {
        setHasVoiceSupport(true);
        if (typeof window.speechSynthesis.getVoices === 'function') {
          window.speechSynthesis.getVoices();
        }
        const onVoicesChanged = () => {
          try {
            if (typeof window !== 'undefined' && window.speechSynthesis) {
              window.speechSynthesis.getVoices();
            }
          } catch {}
        };
        window.speechSynthesis.onvoiceschanged = onVoicesChanged;
        return () => {
          try {
            if (typeof window !== 'undefined' && window.speechSynthesis) {
              window.speechSynthesis.onvoiceschanged = null;
            }
          } catch {}
        };
      } else {
        setHasVoiceSupport(false);
      }
    } catch {
      setHasVoiceSupport(false);
    }
  }, []);

  const clearKeepAlive = () => {
    if (keepAliveTimerRef.current) {
      clearInterval(keepAliveTimerRef.current);
      keepAliveTimerRef.current = null;
    }
  };

  const stopSpeech = () => {
    clearKeepAlive();
    try {
      if (typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis) {
        window.speechSynthesis.cancel();
      }
    } catch {}
    setSpeechActive(false);
  };

  const playVoiceForStep = (stepItem: TutorialStep | undefined) => {
    stopSpeech();
    speechFinishedRef.current = false;

    if (!stepItem || !stepItem.narration || isMuted || !hasVoiceSupport || typeof window === 'undefined') {
      speechFinishedRef.current = true;
      return;
    }

    try {
      if (typeof SpeechSynthesisUtterance === 'undefined' || !window.speechSynthesis) {
        speechFinishedRef.current = true;
        return;
      }

      const utter = new SpeechSynthesisUtterance(stepItem.narration);
      utter.lang = 'pt-BR';
      utter.rate = 1.0; // ritmo claro e articulado
      utter.pitch = 1.08; // tom amigável e natural

      let voices: SpeechSynthesisVoice[] = [];
      try {
        voices = window.speechSynthesis.getVoices() || [];
      } catch {}

      const femalePtVoice = voices.find(v => 
        (v.lang && (v.lang.includes('pt') || v.lang.includes('PT'))) && 
        (v.name.toLowerCase().includes('maria') || 
         v.name.toLowerCase().includes('francisca') || 
         v.name.toLowerCase().includes('leticia') || 
         v.name.toLowerCase().includes('luciana') || 
         v.name.toLowerCase().includes('female') ||
         v.name.toLowerCase().includes('mulher') ||
         v.name.toLowerCase().includes('zira') ||
         v.name.toLowerCase().includes('google'))
      ) || voices.find(v => v.lang && (v.lang.includes('pt') || v.lang.includes('PT')));

      if (femalePtVoice) {
        utter.voice = femalePtVoice;
      }

      utter.onstart = () => {
        setSpeechActive(true);
        speechFinishedRef.current = false;
      };

      utter.onend = () => {
        setSpeechActive(false);
        speechFinishedRef.current = true;
        clearKeepAlive();
      };

      utter.onerror = () => {
        setSpeechActive(false);
        speechFinishedRef.current = true;
        clearKeepAlive();
      };

      utteranceRef.current = utter;
      window.speechSynthesis.speak(utter);

      clearKeepAlive();
      keepAliveTimerRef.current = setInterval(() => {
        try {
          if (typeof window !== 'undefined' && window.speechSynthesis && window.speechSynthesis.speaking) {
            window.speechSynthesis.pause();
            window.speechSynthesis.resume();
          } else {
            clearKeepAlive();
          }
        } catch {
          clearKeepAlive();
        }
      }, 10000);
    } catch (e) {
      console.warn("TTS não iniciado:", e);
      setSpeechActive(false);
      speechFinishedRef.current = true;
    }
  };

  // Desmonte
  useEffect(() => {
    return () => {
      stopSpeech();
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, []);

  // Mudança de passo
  useEffect(() => {
    setProgress(0);
    if (isPlaying && step) {
      playVoiceForStep(step);
    } else {
      stopSpeech();
    }
  }, [currentStepIndex, isPlaying]);

  // Mudo
  useEffect(() => {
    if (isMuted) {
      stopSpeech();
      speechFinishedRef.current = true;
    } else if (isPlaying && step) {
      playVoiceForStep(step);
    }
  }, [isMuted]);

  // Temporizador de progresso e transição de cena
  useEffect(() => {
    if (!isPlaying || !step) {
      if (timerRef.current) clearInterval(timerRef.current);
      return;
    }

    const intervalTime = 100;
    const totalDurationMs = Math.max(((step.duration || 20) * 1000), 20000);
    const stepIncrement = (intervalTime / totalDurationMs) * 100;

    timerRef.current = setInterval(() => {
      setProgress(prev => {
        let isSpeaking = false;
        try {
          isSpeaking = typeof window !== 'undefined' && 'speechSynthesis' in window && !!window.speechSynthesis?.speaking;
        } catch {}

        const voiceStillBusy = !isMuted && hasVoiceSupport && (isSpeaking || !speechFinishedRef.current);

        if (prev >= 98 && voiceStillBusy) {
          return 98;
        }

        if (prev >= 100) {
          if (voiceStillBusy) {
            return 99;
          }

          if (currentStepIndex < TUTORIAL_STEPS.length - 1) {
            setCurrentStepIndex(currentStepIndex + 1);
            return 0;
          } else {
            setIsPlaying(false);
            stopSpeech();
            return 100;
          }
        }
        return prev + stepIncrement;
      });
    }, intervalTime);

    return () => {
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, [isPlaying, currentStepIndex, step?.duration, isMuted, hasVoiceSupport]);

  const handleTogglePlay = () => {
    if (isPlaying) {
      setIsPlaying(false);
      stopSpeech();
    } else {
      if (progress >= 100 && currentStepIndex === TUTORIAL_STEPS.length - 1) {
        setCurrentStepIndex(0);
        setProgress(0);
      }
      setIsPlaying(true);
    }
  };

  const handleRestart = () => {
    setCurrentStepIndex(0);
    setProgress(0);
    if (isPlaying) {
      playVoiceForStep(TUTORIAL_STEPS[0]);
    }
  };

  const handleSelectStep = (idx: number) => {
    setCurrentStepIndex(idx);
    setProgress(0);
  };

  const handleNextStep = () => {
    if (currentStepIndex < TUTORIAL_STEPS.length - 1) {
      setCurrentStepIndex(currentStepIndex + 1);
    }
  };

  const handlePrevStep = () => {
    if (currentStepIndex > 0) {
      setCurrentStepIndex(currentStepIndex - 1);
    }
  };

  const handleNavigateToTargetView = () => {
    if (setView && step.targetView) {
      setView(step.targetView);
      if (onClose) {
        onClose();
      }
    }
  };

  return (
    <div className={cn(
      "bg-slate-900 text-white rounded-[2.5rem] border border-slate-800 shadow-2xl overflow-hidden relative flex flex-col",
      isModal ? "max-h-[92vh] w-full" : "w-full min-h-[640px] my-2"
    )}>
      {/* Luzes de fundo atmosféricas */}
      <div className="absolute top-0 right-1/4 w-96 h-96 bg-blue-600/20 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute bottom-0 right-0 w-80 h-80 bg-emerald-500/15 rounded-full blur-[100px] pointer-events-none" />

      {/* Barra superior de identificação */}
      <div className="p-5 md:px-8 border-b border-slate-800/80 bg-slate-950/50 backdrop-blur-md flex flex-col sm:flex-row sm:items-center justify-between gap-4 shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-gradient-to-tr from-sky-500 to-blue-600 text-white flex items-center justify-center shadow-lg shadow-sky-500/30 shrink-0">
            <Tv size={20} />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-black uppercase tracking-widest text-sky-400 bg-sky-950/80 px-2.5 py-0.5 rounded-full border border-sky-800/50">
                Tour Guiado por Abas
              </span>
              <span className="text-[10px] font-bold text-slate-400 flex items-center gap-1">
                <Sparkles size={11} className="text-amber-400" /> Locução & Áudio Interativo
              </span>
            </div>
            <h2 className="text-base sm:text-lg font-black tracking-tight text-slate-100 mt-0.5">
              Apresentação Completa de Cada Aba do ConsultorPro II
            </h2>
          </div>
        </div>

        {/* Controles de Áudio, Atalho e Fechar */}
        <div className="flex items-center gap-2 self-end sm:self-auto">
          <button
            type="button"
            onClick={() => setIsMuted(!isMuted)}
            className={cn(
              "px-3 py-1.5 rounded-xl text-xs font-bold flex items-center gap-2 border transition-all cursor-pointer",
              isMuted 
                ? "bg-slate-800 text-slate-400 border-slate-700 hover:bg-slate-700"
                : "bg-blue-600/30 text-sky-300 border-blue-500/50 hover:bg-blue-600/40 shadow-xs"
            )}
            title={isMuted ? "Ativar narração em áudio" : "Silenciar áudio"}
          >
            {isMuted ? <VolumeX size={15} /> : <Volume2 size={15} className={speechActive ? "animate-pulse text-sky-400" : ""} />}
            <span>{isMuted ? "Áudio Mudo" : "Voz Ativa"}</span>
          </button>

          {setView && step?.targetView && (
            <Button
              size="sm"
              onClick={handleNavigateToTargetView}
              className="bg-sky-500 hover:bg-sky-400 text-slate-950 text-xs font-black rounded-xl h-8 px-3.5 border-none shadow-md"
              title={`Acessar diretamente ${step?.badge || ''}`}
            >
              <span>Abrir {(step?.badge || '').replace('Passo ', '').replace('Aba ', '')}</span>
            </Button>
          )}

          {onClose && (
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-400 hover:text-white transition-all cursor-pointer border border-slate-700/60 ml-1"
              title="Fechar Pop-up do Tutorial"
            >
              <X size={18} />
            </button>
          )}
        </div>
      </div>

      {/* Palco do Vídeo / Player Audiovisual */}
      <div className={cn(
        "p-5 md:p-8 grid grid-cols-1 lg:grid-cols-12 gap-6 md:gap-8 items-start flex-1",
        isModal ? "overflow-y-auto" : ""
      )}>
        {/* Lado Esquerdo: Janela Gráfica Simulada do Vídeo */}
        <div className="lg:col-span-7 flex flex-col">
          <div className="relative rounded-3xl overflow-hidden bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950/70 border border-slate-800 p-6 md:p-8 min-h-[380px] flex flex-col justify-between shadow-2xl">
            {/* Linhas de scan decorativas estilo telão moderno */}
            <div className="absolute inset-0 bg-[linear-gradient(to_bottom,transparent_50%,rgba(0,0,0,0.25)_51%)] bg-[length:100%_4px] pointer-events-none opacity-40" />

            {/* Cabeçalho da cena */}
            <div className="relative z-10 flex items-center justify-between gap-2 border-b border-slate-800/80 pb-3">
              <span className="text-[11px] font-black tracking-widest uppercase text-sky-400 bg-sky-950/80 border border-sky-800/60 px-3 py-1 rounded-full flex items-center gap-1.5">
                {renderIcon(step?.icon)}
                {step?.sceneVisual?.tag}
              </span>
              <span className="text-xs font-mono font-bold text-slate-400 flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping" />
                Aba {currentStepIndex + 1} de {TUTORIAL_STEPS.length}
              </span>
            </div>

            {/* Conteúdo central da cena */}
            <div className="relative z-10 py-5">
              <h3 className="text-2xl sm:text-3xl font-black text-white tracking-tight mb-2">
                {step?.sceneVisual?.headline}
              </h3>
              <p className="text-sm text-slate-300 font-medium mb-5">
                {step?.subtext}
              </p>

              {/* Cards de Métricas em Destaque */}
              <div className="grid grid-cols-3 gap-3 mb-5">
                {(step?.sceneVisual?.metrics || []).map((m, i) => (
                  <div key={i} className="bg-slate-900/90 border border-slate-800/90 rounded-2xl p-3 text-center">
                    <span className="text-[10px] uppercase font-bold text-slate-400 block tracking-wider mb-0.5">
                      {m.label}
                    </span>
                    <strong className={cn("text-base sm:text-lg font-black font-mono", m.color)}>
                      {m.value}
                    </strong>
                  </div>
                ))}
              </div>

              {/* Detalhes operacionais da tela */}
              <div className="bg-slate-950/80 border border-slate-800 rounded-2xl p-4 space-y-2">
                {(step?.sceneVisual?.mockDetails || []).map((det, i) => (
                  <div key={i} className="flex items-center gap-2.5 text-xs text-slate-300 font-medium">
                    <CheckCircle2 size={14} className="text-sky-400 shrink-0" />
                    <span>{det}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Legenda Dinâmica na Base do Vídeo (Subtitles com Voz) */}
            <div className="relative z-10 bg-slate-950/95 border border-sky-500/20 rounded-2xl p-4 backdrop-blur-md flex items-start gap-3 shadow-lg mt-2">
              <div className="p-2 bg-blue-600/30 text-sky-400 rounded-xl shrink-0 mt-0.5 border border-blue-500/30">
                <Bot size={18} className={speechActive ? "animate-bounce text-sky-300" : "text-slate-400"} />
              </div>
              <div className="flex-1 min-w-0">
                <div className="flex items-center justify-between gap-2 mb-1">
                  <span className="text-[10px] font-black uppercase text-sky-400 tracking-wider">
                    {speechActive ? "Locução Ativa • Orientações da Aba:" : "Orientações da Aba • Transcrição Completa:"}
                  </span>
                  {speechActive && (
                    <span className="inline-flex items-center gap-1 text-[10px] font-mono font-bold text-emerald-400 bg-emerald-950/70 border border-emerald-800/60 px-2 py-0.5 rounded-full">
                      <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                      narrando
                    </span>
                  )}
                </div>
                <p className="text-xs sm:text-sm text-slate-100 font-semibold leading-relaxed tracking-normal select-text break-words">
                  "{step?.narration}"
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Lado Direito: Navegação por Tópicos, Destaques e Playback */}
        <div className="lg:col-span-5 flex flex-col justify-between space-y-5">
          {/* Seletor de Cenas / Abas com rolagem suave */}
          <div>
            <div className="flex items-center justify-between mb-2.5">
              <span className="text-xs font-black uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
                <ListOrdered size={14} /> Abas do Aplicativo ({TUTORIAL_STEPS.length})
              </span>
              <span className="text-xs font-mono font-bold text-sky-400">
                {Math.round(((currentStepIndex + 1) / TUTORIAL_STEPS.length) * 100)}% Concluído
              </span>
            </div>

            <div className="space-y-1.5 max-h-[260px] overflow-y-auto pr-1">
              {TUTORIAL_STEPS.map((st, idx) => {
                const isActive = idx === currentStepIndex;
                const isPassed = idx < currentStepIndex;
                return (
                  <button
                    key={st.id}
                    type="button"
                    onClick={() => handleSelectStep(idx)}
                    className={cn(
                      "w-full text-left p-2.5 rounded-2xl border transition-all flex items-center justify-between gap-3 cursor-pointer",
                      isActive
                        ? "bg-blue-600 text-white border-blue-500 shadow-lg shadow-blue-600/20"
                        : isPassed
                        ? "bg-slate-950/60 text-slate-300 border-slate-800/80 hover:bg-slate-800/60"
                        : "bg-slate-950/30 text-slate-400 border-slate-800/40 hover:bg-slate-800/40"
                    )}
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <div className={cn(
                        "w-6 h-6 rounded-lg flex items-center justify-center text-[11px] font-black shrink-0",
                        isActive ? "bg-white text-blue-600" : isPassed ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/40" : "bg-slate-800 text-slate-400"
                      )}>
                        {isPassed ? <CheckCircle2 size={13} /> : idx + 1}
                      </div>
                      <div className="truncate">
                        <span className={cn(
                          "text-[9px] font-bold uppercase tracking-wider block leading-none mb-0.5",
                          isActive ? "text-blue-100" : "text-slate-400"
                        )}>
                          {st.badge}
                        </span>
                        <h4 className="text-xs font-bold leading-tight truncate">
                          {st.title}
                        </h4>
                      </div>
                    </div>
                    <span className="text-[10px] font-mono opacity-80 shrink-0">
                      {st.duration}s
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Destaques do Passo Atual */}
          <div className="bg-slate-950/70 border border-slate-800 rounded-2xl p-4">
            <span className="text-[11px] font-black uppercase text-amber-400 tracking-wider flex items-center gap-1.5 mb-2">
              <Lightbulb size={14} /> Destaques da Aba: {step?.title}
            </span>
            <ul className="space-y-1.5">
              {(step?.highlights || []).map((h, i) => (
                <li key={i} className="text-xs text-slate-300 flex items-start gap-2">
                  <span className="w-1.5 h-1.5 rounded-full bg-sky-400 mt-1.5 shrink-0" />
                  <span>{h}</span>
                </li>
              ))}
            </ul>
          </div>

          {/* Barra de Progresso do Capítulo */}
          <div>
            <div className="flex justify-between text-[11px] font-mono font-bold text-slate-400 mb-1.5">
              <span className="truncate pr-2">{step?.badge}: {step?.title}</span>
              <span className="shrink-0">{Math.round(progress)}%</span>
            </div>
            <div className="w-full bg-slate-800 rounded-full h-2 overflow-hidden border border-slate-700/60">
              <div 
                className="h-full bg-gradient-to-r from-sky-400 to-blue-500 rounded-full transition-all duration-100"
                style={{ width: `${progress}%` }}
              />
            </div>
          </div>

          {/* Painel Principal de Playback */}
          <div className="flex items-center justify-between gap-3 pt-2 border-t border-slate-800">
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handlePrevStep}
                disabled={currentStepIndex === 0}
                className="p-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-white transition-all cursor-pointer"
                title="Aba Anterior"
              >
                <ChevronLeft size={18} />
              </button>

              <button
                type="button"
                onClick={handleTogglePlay}
                className={cn(
                  "px-5 py-2.5 rounded-xl font-black text-sm flex items-center gap-2 shadow-lg transition-all cursor-pointer",
                  isPlaying 
                    ? "bg-amber-500 hover:bg-amber-400 text-slate-950 shadow-amber-500/30"
                    : "bg-emerald-500 hover:bg-emerald-400 text-slate-950 shadow-emerald-500/30"
                )}
              >
                {isPlaying ? (
                  <>
                    <Pause size={18} className="fill-current" />
                    <span>Pausar</span>
                  </>
                ) : (
                  <>
                    <Play size={18} className="fill-current" />
                    <span>{progress >= 100 && currentStepIndex === TUTORIAL_STEPS.length - 1 ? "Assistir Novamente" : "Iniciar Apresentação"}</span>
                  </>
                )}
              </button>

              <button
                type="button"
                onClick={handleNextStep}
                disabled={currentStepIndex === TUTORIAL_STEPS.length - 1}
                className="p-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-white transition-all cursor-pointer"
                title="Próxima Aba"
              >
                <ChevronRight size={18} />
              </button>
            </div>

            <button
              type="button"
              onClick={handleRestart}
              className="p-2.5 rounded-xl bg-slate-800/80 hover:bg-slate-700 text-slate-300 hover:text-white transition-all cursor-pointer flex items-center gap-1.5 text-xs font-bold"
              title="Reiniciar do Começo"
            >
              <RotateCcw size={15} />
              <span className="hidden sm:inline">Reiniciar</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
