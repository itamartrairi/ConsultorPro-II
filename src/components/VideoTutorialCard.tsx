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
  Calendar
} from 'lucide-react';
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
  icon: 'welcome' | 'company' | 'diagnosis' | 'actionPlan' | 'report' | 'ai';
  sceneVisual: {
    tag: string;
    headline: string;
    metrics: { label: string; value: string; color: string }[];
    mockDetails: string[];
  };
}

const TUTORIAL_STEPS: TutorialStep[] = [
  {
    id: 'intro',
    badge: 'Visão Geral & Boas-Vindas',
    title: 'Bem-vindo ao ConsultorPro II',
    narration: 'Olá! Seja muito bem-vindo ao ConsultorPro dois. Este é o seu assistente inteligente e completo para consultoria gerencial e estratégica. Aqui você gerencia empresas, realiza diagnósticos precisos, gera planos de ação estruturados e emite relatórios profissionais de alto padrão com apoio de inteligência artificial.',
    subtext: 'A plataforma definitiva para consultores gerenciais e credenciados Sebrae.',
    duration: 16,
    highlights: [
      'Ambiente unificado para gestão de clientes e consultorias',
      'Diagnóstico ágil baseado em premissas e análise de maturidade',
      'Automação de cronograma, planos de ação e relatórios com IA'
    ],
    icon: 'welcome',
    sceneVisual: {
      tag: 'CONSULTORIA GERENCIAL PRO',
      headline: 'Plataforma Estratégica para Consultores',
      metrics: [
        { label: 'Índice de Produtividade', value: '+300%', color: 'text-emerald-400' },
        { label: 'Tempo de Relatório', value: '< 2 min', color: 'text-sky-400' },
        { label: 'Conformidade', value: '100%', color: 'text-indigo-400' }
      ],
      mockDetails: [
        'Diagnósticos por segmento e tipo de produto/serviço',
        'Cronograma inteligente de 34 horas com 8 atividades padronizadas',
        'Exportação instantânea para PDF formal e relatórios executivos'
      ]
    }
  },
  {
    id: 'companies',
    badge: 'Passo 1 • Clientes',
    title: 'Cadastro e Gestão de Clientes',
    narration: 'O primeiro passo é cadastrar a empresa cliente na aba Clientes. Você pode buscar automaticamente os dados da empresa pelo CNPJ com autopreenchimento de razão social, endereço e atividade econômica. Você também define o consultor responsável, as metas e o segmento de mercado.',
    subtext: 'Localize e organize empresas com busca em tempo real e consulta de CNPJ.',
    duration: 16,
    highlights: [
      'Preenchimento automático via consulta pública de CNPJ',
      'Histórico de atendimentos e diagnósticos por cliente',
      'Definição de metas de faturamento e segmentos de atuação'
    ],
    icon: 'company',
    sceneVisual: {
      tag: 'MÓDULO CLIENTES',
      headline: 'Dossiê Corporativo e Dados Contratuais',
      metrics: [
        { label: 'Busca CNPJ', value: 'Instantânea', color: 'text-sky-400' },
        { label: 'Histórico', value: 'Centralizado', color: 'text-emerald-400' },
        { label: 'Segurança', value: 'Criptografia', color: 'text-amber-400' }
      ],
      mockDetails: [
        'Autopreenchimento de Razão Social, CNAE e Endereço',
        'Vínculo direto com a Empresa Credenciada e Consultor',
        'Controle de contatos, porte e capacidade produtiva'
      ]
    }
  },
  {
    id: 'diagnosis',
    badge: 'Passo 2 • Diagnóstico',
    title: 'Diagnóstico & Tipo de Produto ou Serviço',
    narration: 'No Diagnóstico, você avalia as áreas da empresa respondendo as premissas como Sim, Parcial ou Não. Agora, você conta com o campo exclusivo Tipo de Produto ou Serviço, permitindo escolher segmentos como Pescados, Gastronomia ou Moda, ou digitar o produto específico do cliente.',
    subtext: 'Personalização profunda: a IA agora conhece com precisão o que o cliente vende.',
    duration: 17,
    highlights: [
      'Novo seletor de "Tipo de Produto ou Serviço" com sugestões rápidas',
      'Cálculo em tempo real de maturidade por área de gestão',
      'Identificação automática de lacunas, gargalos e causas-raiz'
    ],
    icon: 'diagnosis',
    sceneVisual: {
      tag: 'DIAGNÓSTICO EMPRESARIAL',
      headline: 'Avaliação de Maturidade & Novo Campo de Segmento',
      metrics: [
        { label: 'Áreas Analisadas', value: 'Até 10 Áreas', color: 'text-indigo-400' },
        { label: 'Segmento Específico', value: 'Ativado', color: 'text-emerald-400' },
        { label: 'Precisão da IA', value: 'Máxima', color: 'text-sky-400' }
      ],
      mockDetails: [
        'Campo: Tipo de Produto ou Serviço (ex: Camarão / Pescados)',
        'Avaliação rápida de conformidade por premissas estruturadas',
        'Gráficos dinâmicos de radar e pontuação por pilar'
      ]
    }
  },
  {
    id: 'action_plan',
    badge: 'Passo 3 • Plano de Ação',
    title: 'Geração do Plano de Ação Unificado',
    narration: 'Com o diagnóstico respondido, basta clicar no botão Gerar pelo Diagnóstico. O sistema processa as lacunas encontradas e distribui as atividades da consultoria de forma inteligente, integrando soluções recomendadas e adaptando o conteúdo ao tipo de produto ou serviço informado.',
    subtext: 'Quadro Kanban com status, prioridades, datas automáticas e evidências.',
    duration: 17,
    highlights: [
      'Botão unificado "Gerar pelo Diagnóstico" em Gestão do Plano e Relatório',
      'Cronograma calibrado para a carga horária da consultoria',
      'Atividades contextualizadas para o produto ou serviço do cliente'
    ],
    icon: 'actionPlan',
    sceneVisual: {
      tag: 'PLANO DE AÇÃO INTELIGENTE',
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
    badge: 'Passo 4 • Relatório com IA',
    title: 'Relatório Executivo e Exportação em PDF',
    narration: 'Na aba Relatório de Consultoria, você revisa os objetivos, as soluções propostas e os resultados esperados redigidos com Inteligência Artificial. Com apenas um clique, você exporta o relatório oficial completo em PDF, pronto para entrega ao cliente ou homologação institucional. Aproveite todo o poder do ConsultorPro dois!',
    subtext: 'Emissão profissional com brasões, assinaturas, dados técnicos e sumário executivo.',
    duration: 18,
    highlights: [
      'Textos redigidos e enriquecidos com Inteligência Artificial Gemini',
      'Objetivos e resultados alinhados ao tipo de produto ou serviço',
      'Exportação instantânea para PDF formal com layout institucional'
    ],
    icon: 'report',
    sceneVisual: {
      tag: 'ENTREGA FINAL',
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
  }
];

export const VideoTutorialCard: React.FC<{ setView?: (view: any) => void }> = ({ setView }) => {
  const [currentStepIndex, setCurrentStepIndex] = useState(0);
  const [isPlaying, setIsPlaying] = useState(false);
  const [isMuted, setIsMuted] = useState(false);
  const [progress, setProgress] = useState(0);
  const [hasVoiceSupport, setHasVoiceSupport] = useState(true);
  const [speechActive, setSpeechActive] = useState(false);

  const step = TUTORIAL_STEPS[currentStepIndex];
  const timerRef = useRef<NodeJS.Timeout | null>(null);
  const utteranceRef = useRef<SpeechSynthesisUtterance | null>(null);

  // Inicializa suporte a voz e cancela ao desmontar
  useEffect(() => {
    if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
      setHasVoiceSupport(true);
    } else {
      setHasVoiceSupport(false);
    }

    return () => {
      stopSpeech();
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, []);

  const stopSpeech = () => {
    if (typeof window !== 'undefined' && 'speechSynthesis' in window) {
      window.speechSynthesis.cancel();
      setSpeechActive(false);
    }
  };

  const playVoiceForStep = (stepItem: TutorialStep) => {
    stopSpeech();
    if (isMuted || !hasVoiceSupport || typeof window === 'undefined') return;

    try {
      const utter = new SpeechSynthesisUtterance(stepItem.narration);
      utter.lang = 'pt-BR';
      utter.rate = 1.02; // cadência natural e fluida
      utter.pitch = 1.12; // tom levemente mais agudo para realçar voz feminina natural

      // Procura voz feminina brasileira em pt-BR (ex: Maria, Francisca, Leticia, Luciana ou Google português)
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

      utter.onstart = () => setSpeechActive(true);
      utter.onend = () => {
        setSpeechActive(false);
      };
      utter.onerror = () => setSpeechActive(false);

      utteranceRef.current = utter;
      window.speechSynthesis.speak(utter);
    } catch (e) {
      console.warn("TTS não iniciado:", e);
      setSpeechActive(false);
    }
  };

  // Efeito quando altera o passo durante a reprodução
  useEffect(() => {
    setProgress(0);
    if (isPlaying) {
      playVoiceForStep(step);
    } else {
      stopSpeech();
    }
  }, [currentStepIndex, isPlaying]);

  // Efeito quando o mudo é alternado
  useEffect(() => {
    if (isMuted) {
      stopSpeech();
    } else if (isPlaying) {
      playVoiceForStep(step);
    }
  }, [isMuted]);

  // Controlador de progresso e transição de cenas
  useEffect(() => {
    if (!isPlaying) {
      if (timerRef.current) clearInterval(timerRef.current);
      return;
    }

    const intervalTime = 100; // atualiza a cada 100ms
    const totalDurationMs = step.duration * 1000;
    const stepIncrement = (intervalTime / totalDurationMs) * 100;

    timerRef.current = setInterval(() => {
      setProgress(prev => {
        if (prev >= 100) {
          // Avança para o próximo passo ou para no final
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
  }, [isPlaying, currentStepIndex, step.duration]);

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

  return (
    <div className="bg-slate-900 text-white rounded-[2.5rem] border border-slate-800 shadow-2xl overflow-hidden relative">
      {/* Luzes de fundo atmosféricas */}
      <div className="absolute top-0 right-1/4 w-96 h-96 bg-blue-600/20 rounded-full blur-[120px] pointer-events-none" />
      <div className="absolute bottom-0 right-0 w-80 h-80 bg-emerald-500/15 rounded-full blur-[100px] pointer-events-none" />

      {/* Barra superior de identificação */}
      <div className="p-6 md:px-8 border-b border-slate-800/80 bg-slate-950/40 backdrop-blur-md flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-2xl bg-gradient-to-tr from-sky-500 to-blue-600 text-white flex items-center justify-center shadow-lg shadow-sky-500/30">
            <Tv size={20} />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <span className="text-[10px] font-black uppercase tracking-widest text-sky-400 bg-sky-950/80 px-2.5 py-0.5 rounded-full border border-sky-800/50">
                Vídeo Tutorial Oficial
              </span>
              <span className="text-[10px] font-bold text-slate-400 flex items-center gap-1">
                <Sparkles size={11} className="text-amber-400" /> Voz Feminina & Áudio Ativo
              </span>
            </div>
            <h2 className="text-base sm:text-lg font-black tracking-tight text-slate-100 mt-0.5">
              Aprenda a Usar o ConsultorPro II em Poucos Minutos
            </h2>
          </div>
        </div>

        {/* Controles de Áudio e Ação Rápida */}
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

          {setView && (
            <Button
              size="sm"
              onClick={() => setView('projects')}
              className="bg-sky-500 hover:bg-sky-400 text-slate-950 text-xs font-black rounded-xl h-8 px-3.5 border-none shadow-md"
            >
              Ir aos Diagnósticos
            </Button>
          )}
        </div>
      </div>

      {/* Palco do Vídeo / Player Audiovisual */}
      <div className="p-6 md:p-8 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
        {/* Lado Esquerdo: Janela Gráfica Simulada do Vídeo */}
        <div className="lg:col-span-7 flex flex-col">
          <div className="relative rounded-3xl overflow-hidden bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950/70 border border-slate-800 p-6 md:p-8 min-h-[360px] flex flex-col justify-between shadow-2xl">
            {/* Linhas de scan decorativas estilo telão moderno */}
            <div className="absolute inset-0 bg-[linear-gradient(to_bottom,transparent_50%,rgba(0,0,0,0.25)_51%)] bg-[length:100%_4px] pointer-events-none opacity-40" />

            {/* Cabeçalho da cena */}
            <div className="relative z-10 flex items-center justify-between gap-2 border-b border-slate-800/80 pb-3">
              <span className="text-[11px] font-black tracking-widest uppercase text-sky-400 bg-sky-950/80 border border-sky-800/60 px-3 py-1 rounded-full">
                {step.sceneVisual.tag}
              </span>
              <span className="text-xs font-mono font-bold text-slate-400 flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-ping" />
                Cena {currentStepIndex + 1} de {TUTORIAL_STEPS.length}
              </span>
            </div>

            {/* Conteúdo central da cena */}
            <div className="relative z-10 py-6">
              <h3 className="text-2xl sm:text-3xl font-black text-white tracking-tight mb-2">
                {step.sceneVisual.headline}
              </h3>
              <p className="text-sm text-slate-300 font-medium mb-6">
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
            <div className="relative z-10 bg-slate-950/90 border border-slate-800/80 rounded-2xl p-3.5 backdrop-blur-md flex items-start gap-3">
              <div className="p-1.5 bg-blue-600/30 text-sky-400 rounded-lg shrink-0 mt-0.5">
                <Bot size={16} className={speechActive ? "animate-bounce" : ""} />
              </div>
              <div className="flex-1">
                <span className="text-[10px] font-black uppercase text-sky-400 block tracking-wider">
                  Locução Feminina Natural • Transcrição:
                </span>
                <p className="text-xs sm:text-sm text-slate-200 font-semibold leading-relaxed mt-0.5">
                  "{step.narration}"
                </p>
              </div>
            </div>
          </div>
        </div>

        {/* Lado Direito: Navegação por Tópicos, Destaques e Playback */}
        <div className="lg:col-span-5 flex flex-col justify-between space-y-6">
          {/* Seletor de Cenas / Capítulos */}
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-xs font-black uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
                <ListOrdered size={14} /> Capítulos do Tutorial
              </span>
              <span className="text-xs font-mono font-bold text-sky-400">
                {Math.round(((currentStepIndex + 1) / TUTORIAL_STEPS.length) * 100)}% Concluído
              </span>
            </div>

            <div className="space-y-2">
              {TUTORIAL_STEPS.map((st, idx) => {
                const isActive = idx === currentStepIndex;
                const isPassed = idx < currentStepIndex;
                return (
                  <button
                    key={st.id}
                    type="button"
                    onClick={() => handleSelectStep(idx)}
                    className={cn(
                      "w-full text-left p-3 rounded-2xl border transition-all flex items-center justify-between gap-3 cursor-pointer",
                      isActive
                        ? "bg-blue-600 text-white border-blue-500 shadow-lg shadow-blue-600/20"
                        : isPassed
                        ? "bg-slate-950/60 text-slate-300 border-slate-800/80 hover:bg-slate-800/60"
                        : "bg-slate-950/30 text-slate-400 border-slate-800/40 hover:bg-slate-800/40"
                    )}
                  >
                    <div className="flex items-center gap-3">
                      <div className={cn(
                        "w-7 h-7 rounded-xl flex items-center justify-center text-xs font-black",
                        isActive ? "bg-white text-blue-600" : isPassed ? "bg-emerald-500/20 text-emerald-400 border border-emerald-500/40" : "bg-slate-800 text-slate-400"
                      )}>
                        {isPassed ? <CheckCircle2 size={14} /> : idx + 1}
                      </div>
                      <div>
                        <span className={cn(
                          "text-[10px] font-bold uppercase tracking-wider block",
                          isActive ? "text-blue-100" : "text-slate-400"
                        )}>
                          {st.badge}
                        </span>
                        <h4 className="text-xs sm:text-sm font-bold leading-tight">
                          {st.title}
                        </h4>
                      </div>
                    </div>
                    <span className="text-[11px] font-mono opacity-80">
                      {st.duration}s
                    </span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Destaques do Passo Atual */}
          <div className="bg-slate-950/70 border border-slate-800 rounded-2xl p-4">
            <span className="text-[11px] font-black uppercase text-amber-400 tracking-wider flex items-center gap-1.5 mb-2.5">
              <Lightbulb size={14} /> Destaques & Recomendações
            </span>
            <ul className="space-y-2">
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
              <span>Progresso do Capítulo: {step.title}</span>
              <span>{Math.round(progress)}%</span>
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
                title="Capítulo Anterior"
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
                    <span>{progress >= 100 && currentStepIndex === TUTORIAL_STEPS.length - 1 ? "Assistir Novamente" : "Iniciar Tutorial"}</span>
                  </>
                )}
              </button>

              <button
                type="button"
                onClick={handleNextStep}
                disabled={currentStepIndex === TUTORIAL_STEPS.length - 1}
                className="p-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 disabled:opacity-30 disabled:cursor-not-allowed text-white transition-all cursor-pointer"
                title="Próximo Capítulo"
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
