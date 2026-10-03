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
  icon: 'welcome' | 'company' | 'diagnosis' | 'actionPlan' | 'report' | 'ai' | 'agenda' | 'credenciadas' | 'maturity' | 'resultado' | 'macro' | 'premises' | 'settings' | 'licenses';
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
    id: 'intro',
    badge: 'Visão Geral • Início',
    title: 'Visão Geral & Boas-Vindas',
    narration: 'Olá! Seja muito bem-vindo ao ConsultorPro dois. Este é o seu assistente inteligente completo para consultoria gerencial e estratégica. Agora estruturado em abas dinâmicas, o sistema permite acompanhar clientes, aplicar diagnósticos ágeis, gerar cronogramas Kanban de 34 horas e emitir relatórios executivos em PDF com Inteligência Artificial.',
    subtext: 'A plataforma definitiva para consultores gerenciais e credenciados Sebrae.',
    duration: 22,
    highlights: [
      'Ambiente unificado para gestão de clientes e consultorias',
      'Diagnóstico ágil baseado em premissas e análise de maturidade',
      'Automação de cronograma, planos de ação e relatórios com IA'
    ],
    icon: 'welcome',
    targetView: 'home',
    sceneVisual: {
      tag: 'ABA INÍCIO & VISÃO GERAL',
      headline: 'Plataforma Estratégica para Consultores',
      metrics: [
        { label: 'Produtividade', value: '+300%', color: 'text-emerald-400' },
        { label: 'Tempo de Relatório', value: '< 2 min', color: 'text-sky-400' },
        { label: 'Conformidade', value: '100%', color: 'text-indigo-400' }
      ],
      mockDetails: [
        'Painel inicial limpo com acesso instantâneo a todos os módulos',
        'Atalhos rápidos para Clientes, Plano de Ação, Diagnósticos e Macro',
        'Identificação da Credenciada e consultor responsável em destaque'
      ]
    }
  },
  {
    id: 'companies',
    badge: 'Aba Clientes',
    title: 'Cadastro e Gestão de Clientes',
    narration: 'Na aba Clientes, você cadastra e gerencia todas as empresas atendidas. Com a busca automática por CNPJ via Receita Federal, você preenche a Razão Social, CNAE, porte e endereço em um segundo. Você também define metas de faturamento, faturamento atual, contatos e o consultor responsável.',
    subtext: 'Localize e organize empresas com busca em tempo real e consulta pública de CNPJ.',
    duration: 21,
    highlights: [
      'Preenchimento automático via consulta pública de CNPJ',
      'Histórico de atendimentos e diagnósticos por cliente',
      'Definição de metas de faturamento e segmentos de atuação'
    ],
    icon: 'company',
    targetView: 'companies',
    sceneVisual: {
      tag: 'MÓDULO CLIENTES',
      headline: 'Dossiê Corporativo e Dados Contratuais',
      metrics: [
        { label: 'Busca CNPJ', value: 'Instantânea', color: 'text-sky-400' },
        { label: 'Histórico', value: 'Centralizado', color: 'text-emerald-400' },
        { label: 'Segurança', value: 'Criptografia', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Autopreenchimento de Razão Social, CNAE e Endereço comercial',
        'Vínculo direto com a Empresa Credenciada e Consultor',
        'Controle de contatos, porte e faturamento planejado'
      ]
    }
  },
  {
    id: 'projects',
    badge: 'Aba Projetos',
    title: 'Diagnósticos e Histórico de Consultorias',
    narration: 'A aba Projetos é o coração dos diagnósticos da empresa. Aqui você visualiza todos os diagnósticos abertos e finalizados, cria um novo diagnóstico em segundos ou replica diagnósticos anteriores para acompanhar a evolução do cliente ao longo do tempo.',
    subtext: 'Controle de ciclos de atendimento, criação e replicação de diagnósticos.',
    duration: 20,
    highlights: [
      'Criação de novos ciclos de consultoria para o cliente selecionado',
      'Replicação ágil de diagnósticos para comparativo evolutivo',
      'Histórico completo de pontuações e status de finalização'
    ],
    icon: 'diagnosis',
    targetView: 'projects',
    sceneVisual: {
      tag: 'MÓDULO PROJETOS',
      headline: 'Gestão de Ciclos de Diagnóstico',
      metrics: [
        { label: 'Ciclos', value: 'Ilimitados', color: 'text-sky-400' },
        { label: 'Replicação', value: '1 Clique', color: 'text-emerald-400' },
        { label: 'Status', value: 'Em Andamento', color: 'text-purple-400' }
      ],
      mockDetails: [
        'Listagem de diagnósticos por empresa com data e consultor',
        'Atalho direto para preencher premissas e questionário',
        'Histórico organizado com opção de arquivamento seguro'
      ]
    }
  },
  {
    id: 'agenda',
    badge: 'Aba Agenda',
    title: 'Agenda do Consultor & Compromissos',
    narration: 'Na aba Agenda do Consultor, você planeja todas as visitas técnicas, reuniões de alinhamento e sessões de devolutiva. Organize seus compromissos com data, horário, empresa cliente e lembretes para nunca perder um prazo ou entrega contratual.',
    subtext: 'Calendário e programação de visitas técnicas e reuniões estratégicas.',
    duration: 19,
    highlights: [
      'Visão diária, semanal e mensal de compromissos de consultoria',
      'Vínculo direto entre compromisso e a empresa atendida',
      'Registro de reuniões presenciais, online e entregas de relatórios'
    ],
    icon: 'agenda',
    targetView: 'agenda',
    sceneVisual: {
      tag: 'MÓDULO AGENDA',
      headline: 'Programação de Visitas e Reuniões',
      metrics: [
        { label: 'Visitas', value: 'Agendadas', color: 'text-amber-400' },
        { label: 'Visão', value: 'Semanal / Mês', color: 'text-emerald-400' },
        { label: 'Controle', value: 'Total', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Planejamento de visitas presenciais e sessões remotas',
        'Lembretes de entregas intermediárias e encerramento',
        'Sincronização com dados do cliente e consultor'
      ]
    }
  },
  {
    id: 'credenciadas',
    badge: 'Aba Credenciadas',
    title: 'Gestão de Credenciadas & Consultores',
    narration: 'Na aba Credenciadas, você cadastra as empresas de consultoria credenciadas ao Sebrae, seus dados jurídicos, CNPJ e o consultor responsável com CPF. Todos esses dados são automaticamente carimbados no cabeçalho e rodapé dos relatórios oficiais.',
    subtext: 'Configuração dos dados da consultoria credenciada e corpo técnico.',
    duration: 20,
    highlights: [
      'Cadastro de Razão Social, CNPJ e endereço da Credenciada',
      'Vínculo do Consultor titular, CPF e contatos oficiais',
      'Carimbo e assinatura automática nos relatórios em PDF'
    ],
    icon: 'credenciadas',
    targetView: 'credenciadas',
    sceneVisual: {
      tag: 'MÓDULO CREDENCIADAS',
      headline: 'Dados Institucionais da Consultora',
      metrics: [
        { label: 'Credenciadas', value: 'Cadastradas', color: 'text-indigo-400' },
        { label: 'Carimbo PDF', value: 'Automático', color: 'text-emerald-400' },
        { label: 'Conformidade', value: 'Sebrae', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Armazenamento de dados cadastrais e consultores responsáveis',
        'Seleção da credenciada ativa para o atendimento corrente',
        'Integração transparente com cabeçalhos de relatórios executivos'
      ]
    }
  },
  {
    id: 'maturity',
    badge: 'Aba Maturidade',
    title: 'Maturidade Empresarial & Radares',
    narration: 'A aba Maturidade Empresarial calcula o nível de maturidade da empresa em cada dimensão de gestão: Financeira, Processos, Mercado, Pessoas e Estratégia. Gráficos de radar dinâmicos e scores percentuais mostram de forma visual os pontos fortes e as prioridades imediatas de melhoria.',
    subtext: 'Indicadores gráficos de radar, pontuação percentual e análise por pilares.',
    duration: 21,
    highlights: [
      'Gráfico de radar comparativo com benchmark de maturidade',
      'Cálculo automático de porcentagem de conformidade por pilar',
      'Diagnóstico visual ideal para apresentar ao empresário na primeira sessão'
    ],
    icon: 'maturity',
    targetView: 'maturity-assessment',
    sceneVisual: {
      tag: 'AVALIAÇÃO DE MATURIDADE',
      headline: 'Radar de Gestão e Nível de Eficiência',
      metrics: [
        { label: 'Pilares', value: '5 a 10 Áreas', color: 'text-violet-400' },
        { label: 'Gráfico', value: 'Radar Interativo', color: 'text-emerald-400' },
        { label: 'Score Geral', value: 'Automático', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Pontuação calculada a partir das respostas do questionário',
        'Identificação de gargalos em Finanças, Marketing, Operações e RH',
        'Exportação gráfica para apresentação ao empresário'
      ]
    }
  },
  {
    id: 'resultado',
    badge: 'Aba Resultado',
    title: 'Resultado da Consultoria & Evolução',
    narration: 'Na aba Resultado da Consultoria, você compara o diagnóstico inicial com a situação final após a implantação das melhorias. Comprove numericamente o ganho de maturidade, o aumento de faturamento e a redução de desperdícios conquistados pelo cliente.',
    subtext: 'Demonstrativo antes versus depois para comprovação de valor da consultoria.',
    duration: 21,
    highlights: [
      'Comparativo claro do status inicial versus final da empresa',
      'Evidência quantitativa do impacto da consultoria no negócio',
      'Validação de metas contratuais e satisfação do cliente'
    ],
    icon: 'resultado',
    targetView: 'resultado-consultoria',
    sceneVisual: {
      tag: 'RESULTADO & IMPACTO',
      headline: 'Evolução Antes vs. Depois da Consultoria',
      metrics: [
        { label: 'Evolução Média', value: '+45% Ganho', color: 'text-emerald-400' },
        { label: 'Metas Batidas', value: '100%', color: 'text-sky-400' },
        { label: 'Satisfação', value: 'NPS 98', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Gráficos comparativos de evolução entre o primeiro e último ciclo',
        'Registro de lições aprendidas e recomendações futuras',
        'Relatório resumido de impacto gerencial para a diretoria'
      ]
    }
  },
  {
    id: 'macro',
    badge: 'Aba Dashboard Macro',
    title: 'Dashboard Macro & Análise Global',
    narration: 'O Dashboard Macro consolida os dados de todas as empresas e consultorias em uma visão executiva única. Analise a média de maturidade de toda a sua carteira de clientes, os setores mais atendidos e os problemas mais recorrentes em um só painel.',
    subtext: 'Visão executiva em tempo real de toda a carteira de empresas atendidas.',
    duration: 20,
    highlights: [
      'Métricas consolidadas de todas as consultorias ativas',
      'Distribuição por segmento, porte e região geográfica',
      'Identificação de padrões e gargalos comuns no mercado'
    ],
    icon: 'macro',
    targetView: 'macro-dashboard',
    sceneVisual: {
      tag: 'DASHBOARD MACRO',
      headline: 'Inteligência Global da Carteira de Clientes',
      metrics: [
        { label: 'Empresas Ativas', value: 'Consolidado', color: 'text-emerald-400' },
        { label: 'Maturidade Média', value: 'Tempo Real', color: 'text-sky-400' },
        { label: 'Setores', value: 'Multissegmento', color: 'text-purple-400' }
      ],
      mockDetails: [
        'Gráficos de barras e pizza com distribuição setorial',
        'Filtro por período, credenciada e consultor responsável',
        'Exportação rápida para prestação de contas institucional'
      ]
    }
  },
  {
    id: 'premises',
    badge: 'Aba Biblioteca',
    title: 'Biblioteca de Perguntas, Problemas & Soluções',
    narration: 'Na Biblioteca de Perguntas, você personaliza todo o conhecimento do sistema. Edite as perguntas e premissas do diagnóstico, configure a árvore de problemas e causas-raiz, e defina soluções recomendadas para cada segmento de mercado.',
    subtext: 'Base de conhecimento customizável com premissas, problemas, soluções e áreas.',
    duration: 21,
    highlights: [
      'Edição completa de perguntas por área de gestão',
      'Configuração de problemas padrão e suas causas estruturais',
      'Criação de soluções sugeridas personalizadas por segmento'
    ],
    icon: 'premises',
    targetView: 'premises',
    sceneVisual: {
      tag: 'BIBLIOTECA DE CONHECIMENTO',
      headline: 'Repositório Estratégico de Premissas e Soluções',
      metrics: [
        { label: 'Premissas', value: 'Personalizáveis', color: 'text-indigo-400' },
        { label: 'Problemas', value: 'Mapeados', color: 'text-amber-400' },
        { label: 'Soluções', value: 'Inteligentes', color: 'text-emerald-400' }
      ],
      mockDetails: [
        'Abas organizadas: Problemas, Premissas, Soluções, Áreas e Segmentos',
        'Inclusão de novos setores econômicos e perguntas específicas',
        'Restauração de biblioteca padrão e importação em lote'
      ]
    }
  },
  {
    id: 'action_plan',
    badge: 'Aba Plano de Ação',
    title: 'Plano de Ação Kanban & Cronograma 34h',
    narration: 'Na aba Plano de Ação, o sistema distribui automaticamente as 8 atividades do cronograma padrão de 34 horas através do botão Gerar pelo Diagnóstico. Você controla tarefas em estilo Kanban, atribui responsáveis, datas, status e anexa fotos e evidências comprobatórias de cada etapa.',
    subtext: 'Quadro Kanban com status, prioridades, datas automáticas e evidências.',
    duration: 22,
    highlights: [
      'Geração automática calibrada para 8 atividades e 34 horas',
      'Atividades contextualizadas para o Tipo de Produto ou Serviço do cliente',
      'Quadro Kanban com upload de fotos e comprovações de consultoria'
    ],
    icon: 'actionPlan',
    targetView: 'kanban',
    sceneVisual: {
      tag: 'PLANO DE AÇÃO & KANBAN',
      headline: 'Cronograma Kanban e Atividades Estratégicas',
      metrics: [
        { label: 'Carga Horária', value: '34 Horas', color: 'text-emerald-400' },
        { label: 'Atividades', value: '8 Etapas', color: 'text-sky-400' },
        { label: 'Origem', value: 'Diagnóstico', color: 'text-indigo-400' }
      ],
      mockDetails: [
        'Distribuição automática de datas conforme a consultoria',
        'Controle de status: Pendente, Em Andamento e Concluído',
        'Registro de evidências fotográficas e anexos comprobatórios'
      ]
    }
  },
  {
    id: 'report_ai',
    badge: 'Aba Relatório',
    title: 'Relatório de Consultoria & Exportação PDF',
    narration: 'Na aba Relatório de Consultoria, a Inteligência Artificial Gemini redige e aprimora os objetivos, soluções e resultados esperados. Com apenas um clique, você exporta o relatório formal oficial completo em PDF com layout executivo, pronto para assinatura e entrega ao Sebrae ou cliente.',
    subtext: 'Emissão profissional com brasões, assinaturas, dados técnicos e sumário executivo.',
    duration: 22,
    highlights: [
      'Redação inteligente com IA adaptada ao nicho e produto do cliente',
      'Objetivos, justificativas e planos consolidados em alta qualidade',
      'Exportação para PDF oficial com fotos, tabelas de horas e assinaturas'
    ],
    icon: 'report',
    targetView: 'cronograma',
    sceneVisual: {
      tag: 'RELATÓRIO & EXPORTAÇÃO',
      headline: 'Relatório Executivo Oficial & Exportação PDF',
      metrics: [
        { label: 'Formato', value: 'PDF Oficial', color: 'text-rose-400' },
        { label: 'Redação IA', value: 'Gemini Integrado', color: 'text-amber-400' },
        { label: 'Pronto p/ Entrega', value: 'Sim', color: 'text-emerald-400' }
      ],
      mockDetails: [
        'Identificação completa de cliente, consultor e credenciada',
        'Resumo dos problemas identificados e soluções implantadas',
        'Relatório de 8 atividades com detalhamento de horas e resultados'
      ]
    }
  },
  {
    id: 'settings',
    badge: 'Aba Configurações',
    title: 'Configurações, Backup & Nuvem',
    narration: 'Na aba Configurações, você gerencia a sua Chave da API do Google Gemini, alterna entre salvamento na Nuvem ou Local com criptografia AES-256, e realiza o download de backups de segurança dos seus dados a qualquer momento. Tudo rápido, seguro e sob seu controle!',
    subtext: 'Chave de IA Gemini, gestão de backup em JSON e segurança de dados.',
    duration: 21,
    highlights: [
      'Configuração da Chave da API do Gemini para uso exclusivo ou compartilhado',
      'Exportação e importação de backups completos em arquivo JSON',
      'Alternância imediata entre Modo Nuvem (Firestore) e Modo Local'
    ],
    icon: 'settings',
    targetView: 'settings',
    sceneVisual: {
      tag: 'CONFIGURAÇÕES & SEGURANÇA',
      headline: 'Gestão de Inteligência Artificial e Dados',
      metrics: [
        { label: 'Backup', value: 'JSON Local', color: 'text-emerald-400' },
        { label: 'IA Gemini', value: 'Configurável', color: 'text-sky-400' },
        { label: 'Criptografia', value: 'AES-256', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Modal de tutorial para gerar Chave de API gratuita do Gemini',
        'Sincronização em nuvem e histórico de backups locais',
        'Controle de logo personalizado do Sebrae e da Consultora'
      ]
    }
  }
];

const renderIcon = (type: TutorialStep['icon']) => {
  switch (type) {
    case 'welcome': return <Sparkles size={16} className="text-amber-400" />;
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

  const step = TUTORIAL_STEPS[currentStepIndex];
  const timerRef = useRef<NodeJS.Timeout | null>(null);
  const keepAliveTimerRef = useRef<NodeJS.Timeout | null>(null);
  const speechFinishedRef = useRef(false);
  const utteranceRef = useRef<SpeechSynthesisUtterance | null>(null);

  // Inicializa suporte a voz e cancela ao desmontar
  useEffect(() => {
    if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
      setHasVoiceSupport(true);
      window.speechSynthesis.getVoices();
      const onVoicesChanged = () => {
        window.speechSynthesis.getVoices();
      };
      window.speechSynthesis.onvoiceschanged = onVoicesChanged;
      return () => {
        if (window.speechSynthesis) {
          window.speechSynthesis.onvoiceschanged = null;
        }
      };
    } else {
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
    if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
      window.speechSynthesis.cancel();
      setSpeechActive(false);
    }
  };

  const playVoiceForStep = (stepItem: TutorialStep) => {
    stopSpeech();
    speechFinishedRef.current = false;

    if (isMuted || !hasVoiceSupport || typeof window === 'undefined') {
      speechFinishedRef.current = true;
      return;
    }

    try {
      const utter = new SpeechSynthesisUtterance(stepItem.narration);
      utter.lang = 'pt-BR';
      utter.rate = 1.0; // ritmo claro e articulado
      utter.pitch = 1.08; // tom amigável e natural

      const voices = window.speechSynthesis.getVoices();
      const femalePtVoice = voices.find(v => 
        (v.lang.includes('pt') || v.lang.includes('PT')) && 
        (v.name.toLowerCase().includes('maria') || 
         v.name.toLowerCase().includes('francisca') || 
         v.name.toLowerCase().includes('leticia') || 
         v.name.toLowerCase().includes('luciana') || 
         v.name.toLowerCase().includes('female') ||
         v.name.toLowerCase().includes('mulher') ||
         v.name.toLowerCase().includes('zira') ||
         v.name.toLowerCase().includes('google'))
      ) || voices.find(v => v.lang.includes('pt') || v.lang.includes('PT'));

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
        if (typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis.speaking) {
          window.speechSynthesis.pause();
          window.speechSynthesis.resume();
        } else {
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
    if (isPlaying) {
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
    } else if (isPlaying) {
      playVoiceForStep(step);
    }
  }, [isMuted]);

  // Temporizador de progresso e transição de cena
  useEffect(() => {
    if (!isPlaying) {
      if (timerRef.current) clearInterval(timerRef.current);
      return;
    }

    const intervalTime = 100;
    const totalDurationMs = Math.max(step.duration * 1000, 20000);
    const stepIncrement = (intervalTime / totalDurationMs) * 100;

    timerRef.current = setInterval(() => {
      setProgress(prev => {
        const isSpeaking = typeof window !== 'undefined' && 'speechSynthesis' in window && window.speechSynthesis.speaking;
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
  }, [isPlaying, currentStepIndex, step.duration, isMuted, hasVoiceSupport]);

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
      isModal ? "max-h-[92vh] w-full" : ""
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

          {setView && step.targetView && (
            <Button
              size="sm"
              onClick={handleNavigateToTargetView}
              className="bg-sky-500 hover:bg-sky-400 text-slate-950 text-xs font-black rounded-xl h-8 px-3.5 border-none shadow-md"
              title={`Acessar diretamente ${step.badge}`}
            >
              <span>Abrir {step.badge.replace('Passo ', '').replace('Aba ', '')}</span>
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
      <div className="p-5 md:p-8 grid grid-cols-1 lg:grid-cols-12 gap-6 md:gap-8 items-start overflow-y-auto flex-1">
        {/* Lado Esquerdo: Janela Gráfica Simulada do Vídeo */}
        <div className="lg:col-span-7 flex flex-col">
          <div className="relative rounded-3xl overflow-hidden bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950/70 border border-slate-800 p-6 md:p-8 min-h-[380px] flex flex-col justify-between shadow-2xl">
            {/* Linhas de scan decorativas estilo telão moderno */}
            <div className="absolute inset-0 bg-[linear-gradient(to_bottom,transparent_50%,rgba(0,0,0,0.25)_51%)] bg-[length:100%_4px] pointer-events-none opacity-40" />

            {/* Cabeçalho da cena */}
            <div className="relative z-10 flex items-center justify-between gap-2 border-b border-slate-800/80 pb-3">
              <span className="text-[11px] font-black tracking-widest uppercase text-sky-400 bg-sky-950/80 border border-sky-800/60 px-3 py-1 rounded-full flex items-center gap-1.5">
                {renderIcon(step.icon)}
                {step.sceneVisual.tag}
              </span>
              <span className="text-xs font-mono font-bold text-slate-400 flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping" />
                Aba {currentStepIndex + 1} de {TUTORIAL_STEPS.length}
              </span>
            </div>

            {/* Conteúdo central da cena */}
            <div className="relative z-10 py-5">
              <h3 className="text-2xl sm:text-3xl font-black text-white tracking-tight mb-2">
                {step.sceneVisual.headline}
              </h3>
              <p className="text-sm text-slate-300 font-medium mb-5">
                {step.subtext}
              </p>

              {/* Cards de Métricas em Destaque */}
              <div className="grid grid-cols-3 gap-3 mb-5">
                {step.sceneVisual.metrics.map((m, i) => (
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
                {step.sceneVisual.mockDetails.map((det, i) => (
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
                  "{step.narration}"
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
              <Lightbulb size={14} /> Destaques da Aba: {step.title}
            </span>
            <ul className="space-y-1.5">
              {step.highlights.map((h, i) => (
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
              <span className="truncate pr-2">{step.badge}: {step.title}</span>
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
