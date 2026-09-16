import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { db, auth } from '../firebase';
import { collection, query, where, onSnapshot, addDoc, doc, deleteDoc, updateDoc } from 'firebase/firestore';
import {
  User,
  FileText,
  Sparkles,
  ArrowLeft,
  ArrowRight,
  Check,
  Plus,
  Trash2,
  Download,
  RefreshCw,
  Award,
  Briefcase,
  Calendar,
  Building2,
  ChevronRight,
  Info,
  HelpCircle,
  Printer
} from 'lucide-react';
import html2canvas from 'html2canvas-pro';
import jsPDF from 'jspdf';

// Structure of a DISC Question
interface DiscQuestion {
  id: number;
  scenario: string;
  options: {
    letter: 'D' | 'I' | 'S' | 'C';
    text: string;
    description: string;
  }[];
}

// Model for DISC Evaluation saved in DB
interface DiscEvaluation {
  id?: string;
  empresaId: string;
  nomeAvaliado: string;
  cargo: string;
  dataAvaliacao: string;
  ownerId: string;
  scoreD: number; // raw out of 20
  scoreI: number;
  scoreS: number;
  scoreC: number;
  perfilPrincipal: string;
  relatorio: string;
  sugestoesMelhoria: string;
}

interface DiscEvaluationViewProps {
  empresas: { id: string; nome: string }[];
  setView: (view: string) => void;
}

// 20 High-Quality Situational DISC Questions
const DISC_QUESTIONS: DiscQuestion[] = [
  {
    id: 1,
    scenario: "Em uma reunião de equipe, minha atitude natural é:",
    options: [
      { letter: 'D', text: "Focar direto nos resultados práticos e decisões rápidas.", description: "Direto e focado em metas." },
      { letter: 'I', text: "Estimular a participação, trocar ideias criativas e interagir.", description: "Sociável e focado em relacionamentos." },
      { letter: 'S', text: "Apoiar o grupo, ouvir pacientemente e buscar o consenso.", description: "Colaborativo e ouvinte leal." },
      { letter: 'C', text: "Analisar os fatos objetivos, dados e planejar detalhes com precisão.", description: "Lógico, analítico e rigoroso." }
    ]
  },
  {
    id: 2,
    scenario: "Ao iniciar um novo projeto, meu principal objetivo é:",
    options: [
      { letter: 'D', text: "Conquistar as metas rapidamente e superar todos os desafios.", description: "Orientado a conquistas e desafios." },
      { letter: 'I', text: "Motivar as pessoas e contagiar o time com entusiasmo.", description: "Inspirador e focado no espírito de equipe." },
      { letter: 'S', text: "Garantir que todos estejam confortáveis e seguros com o plano.", description: "Buscando estabilidade e harmonia." },
      { letter: 'C', text: "Assegurar que seguiremos regras, processos e padrões de excelência.", description: "Focado em qualidade e conformidade." }
    ]
  },
  {
    id: 3,
    scenario: "Quando enfrento uma situação de forte pressão ou conflito de interesses:",
    options: [
      { letter: 'D', text: "Sou direto, assertivo e tomo a liderança para resolver logo.", description: "Decidido e voltado a resultados." },
      { letter: 'I', text: "Uso de simpatia, comunicação e persuasão para aliviar o clima.", description: "Conciliador social através do carisma." },
      { letter: 'S', text: "Mantenho a calma, ouço o outro lado e busco restaurar a paz.", description: "Evita atritos e zela pela estabilidade." },
      { letter: 'C', text: "Apoio-me em dados concretos, fatos e atuo de forma racional.", description: "Focado em critérios técnicos objetivos." }
    ]
  },
  {
    id: 4,
    scenario: "As pessoas ao meu redor costumam me descrever principalmente como alguém:",
    options: [
      { letter: 'D', text: "Determinado, focado em resultados rápidos e competitivo.", description: "Competitivo e impaciente." },
      { letter: 'I', text: "Comunicativo, espontâneo, otimista e carismático.", description: "Sociável e persuasivo." },
      { letter: 'S', text: "Confiável, prestativo, paciente e leal.", description: "Calmo e acolhedor." },
      { letter: 'C', text: "Organizado, detalhista, preciso e analítico.", description: "Metódico e perfeccionista." }
    ]
  },
  {
    id: 5,
    scenario: "No trabalho, o que mais me traz energia e motivação é:",
    options: [
      { letter: 'D', text: "Ter poder de decisão, autonomia e vencer metas ambiciosas.", description: "Poder e autonomia de execução." },
      { letter: 'I', text: "Colaborar livremente, criar conexões e receber reconhecimento.", description: "Pertencimento e validação social." },
      { letter: 'S', text: "Contar com estabilidade, ambiente harmônico e processos previsíveis.", description: "Ambiente seguro e cooperativo." },
      { letter: 'C', text: "Atingir a excelência técnica, precisão absoluta e alto padrão de entrega.", description: "Busca constante por qualidade lógica." }
    ]
  },
  {
    id: 6,
    scenario: "Se houver uma mudança brusca ou repentina nas diretrizes estabelecidas:",
    options: [
      { letter: 'D', text: "Ajo rapidamente para redefinir o curso e assumir o controle.", description: "Adaptável pela necessidade de controle." },
      { letter: 'I', text: "Enxergo imediatamente as novas oportunidades criativas de engajamento.", description: "Otimista e aberto à novidade." },
      { letter: 'S', text: "Sinto desconforto inicial e prefiro entender bem os motivos antes de agir.", description: "Prefere transições planejadas." },
      { letter: 'C', text: "Analiso minuciosamente as novas regras e os impactos metodológicos.", description: "Exige clareza de critérios organizados." }
    ]
  },
  {
    id: 7,
    scenario: "Ao tomar decisões cruciais que impactam outras pessoas, meu guia principal é:",
    options: [
      { letter: 'D', text: "Minha intuição prática, assertividade e o senso de urgência.", description: "Tomada de decisão rápida." },
      { letter: 'I', text: "As opiniões, sentimentos e a harmonia das pessoas envolvidas.", description: "Decisões mediadas por empatia social." },
      { letter: 'S', text: "A segurança de longo prazo e evitar transtornos para a equipe.", description: "Prevenção de impactos desestabilizadores." },
      { letter: 'C', text: "Dados históricos estruturados, análises frias e lógica sólida.", description: "Decisões embasadas racionalmente." }
    ]
  },
  {
    id: 8,
    scenario: "Quando recebo feedbacks com críticas severas ao meu trabalho:",
    options: [
      { letter: 'D', text: "Foco nos pontos práticos do problema e parto direto para a correção.", description: "Pragmático e auto-corretivo." },
      { letter: 'I', text: "Sinto no lado emocional e tendo a justificar-me usando de simpatia.", description: "Sensível à desaprovação pessoal." },
      { letter: 'S', text: "Ouço pacientemente de forma receptiva e mudo meu ritmo com cuidado.", description: "Dócil e atencioso com as opiniões." },
      { letter: 'C', text: "Exijo comprovação lógica e dados que sustentem a crítica de forma justa.", description: "Analisa a justiça com base técnica." }
    ]
  },
  {
    id: 9,
    scenario: "Em dinâmicas coletivas, meu papel de destaque costuma ser:",
    options: [
      { letter: 'D', text: "Direcionar e liderar o time focado no resultado do cronograma.", description: "Líder resoluto focado em tarefas." },
      { letter: 'I', text: "Inspirar a colaboração, gerar energia positiva e descontração.", description: "Influenciador e animador do time." },
      { letter: 'S', text: "Zelar pela cooperação e dar suporte a quem estiver com dificuldades.", description: "Apoiador estruturado leal." },
      { letter: 'C', text: "Revisar as tarefas de todos para garantir que a qualidade esteja impecável.", description: "Guardião da excelência técnica." }
    ]
  },
  {
    id: 10,
    scenario: "Identifico que meu principal ponto fraco em momentos de estresse é:",
    options: [
      { letter: 'D', text: "Impaciência extrema ou excesso de rigidez com as pessoas.", description: "Agressividade ou frieza." },
      { letter: 'I', text: "Falta de foco nos detalhes ou dispersão com muitas conversas.", description: "Dificuldade de estruturação métrica." },
      { letter: 'S', text: "Dificuldade de me posicionar assertivamente ou resistir a mudanças.", description: "Hesitação e conformismo passivo." },
      { letter: 'C', text: "Perfeccionismo paralisante ou ser excessivamente crítico com o time.", description: "Hipercriticismo e isolamento analítico." }
    ]
  },
  {
    id: 11,
    scenario: "O ambiente de trabalho ideal para mim deve possuir:",
    options: [
      { letter: 'D', text: "Desafios ambiciosos constantes, autonomia plena e pouca supervisão.", description: "Ambiente dinâmico com liberdade decisória." },
      { letter: 'I', text: "Espaço aberto para conversas, atividades cooperativas e dinamismo.", description: "Ambiente alegre e socialmente ativo." },
      { letter: 'S', text: "Clima amigável de colaboração mútua e previsibilidade de rotinas.", description: "Ambiente acolhedor e com estabilidade." },
      { letter: 'C', text: "Sistemas organizados, regras claras e alta valorização do rigor técnico.", description: "Ambiente ordenado e intelectualmente rico." }
    ]
  },
  {
    id: 12,
    scenario: "Na execução diária de minhas obrigações profissionais:",
    options: [
      { letter: 'D', text: "Finalizo o quanto antes para assumir novas metas audaciosas.", description: "Rápido e focado em quantidade de conquistas." },
      { letter: 'I', text: "Busco fazê-las de forma descontraída e com interação constante.", description: "Criativo e motivado por interações." },
      { letter: 'S', text: "Sigo o cronograma planejado com regularidade, paciência e constância.", description: "Consistente e focado em ritmo regular." },
      { letter: 'C', text: "Asseguro que cada detalhe técnico seja perfeito, custe o tempo que custar.", description: "Qualidade cirúrgica e rigor nos mínimos detalhes." }
    ]
  },
  {
    id: 13,
    scenario: "Se percebo que um colega está enfrentando fortes obstáculos de desempenho:",
    options: [
      { letter: 'D', text: "Dou uma orientação prática, direta e focada na solução do problema.", description: "Mentoria direta e sem rodeios." },
      { letter: 'I', text: "Converso amigavelmente, conto uma história de superação e elevo o ânimo.", description: "Motivador emocional expansivo." },
      { letter: 'S', text: "Coloco-me à disposição com calma para ajudá-lo a fazer a tarefa junto.", description: "Companheiro prestativo e paciente." },
      { letter: 'C', text: "Analiso os gargalos do fluxo de trabalho dele e sugiro melhorias lógicas.", description: "Resolutor de problemas de sistema." }
    ]
  },
  {
    id: 14,
    scenario: "Prefiro ser formalmente elogiado e reconhecido por:",
    options: [
      { letter: 'D', text: "Resultados expressivos tangíveis obtidos e minha determinação fora da curva.", description: "Reconhecimento por liderança e metas." },
      { letter: 'I', text: "Habilidades extraordinárias de comunicação, carisma e conexões.", description: "Reconhecimento por empatia e carisma." },
      { letter: 'S', text: "Minha lealdade extrema, suporte constante e espírito de equipe duradouro.", description: "Reconhecimento por colaboração fiel." },
      { letter: 'C', text: "Minha precisão metodológica, domínio técnico e entrega perfeita.", description: "Reconhecimento por excelência e conhecimento." }
    ]
  },
  {
    id: 15,
    scenario: "Minha postura perante prazos e gestão de tempo é:",
    options: [
      { letter: 'D', text: "Entrego de forma rápida e dinâmica, mesmo ignorando pequenos trâmites.", description: "Urgência pragmática acelerada." },
      { letter: 'I', text: "Posso me dispersar em interações, precisando me esforçar ao final para cumprir.", description: "Dispersão criativa e social." },
      { letter: 'S', text: "Gerencio minhas entregas de forma calma, sem pressões extremas, cumprindo o combinado.", description: "Planejamento linear seguro." },
      { letter: 'C', text: "Sou extremamente minucioso e busco entregar exatamente no padrão técnico com pontualidade.", description: "Cumprimento exato de especificações." }
    ]
  },
  {
    id: 16,
    scenario: "Para que eu seja verdadeiramente persuadido ou convencido de uma ideia:",
    options: [
      { letter: 'D', text: "Preciso ver provas imediatas de retorno prático e vantagens diretas.", description: "Abordagem focada em custo-benefício prático." },
      { letter: 'I', text: "Gosto de entusiasmo, dinamismo e uma narrativa contagiante.", description: "Conexão emocional com a visão sugerida." },
      { letter: 'S', text: "Exijo honestidade, tranquilidade e a certeza de que todos serão apoiados.", description: "Avaliação do impacto nas pessoas e clima." },
      { letter: 'C', text: "Exijo provas empíricas lógicas, dados estruturados e referências confiáveis.", description: "Análise profunda de consistência racional." }
    ]
  },
  {
    id: 17,
    scenario: "Se os resultados planejados começam a desviar severamente das metas:",
    options: [
      { letter: 'D', text: "Fico impaciente, mas de imediato mudo a estratégia de forma enérgica.", description: "Flexibilidade proativa motivada pela urgência." },
      { letter: 'I', text: "Mantenho o otimismo, reúno as pessoas e busco alternativas criativas.", description: "Gestão motivacional de contingência." },
      { letter: 'S', text: "Fico preocupado com o clima do time e busco estabilizar a situação com calma.", description: "Estabilização interna preventiva." },
      { letter: 'C', text: "Busco metodicamente identificar a falha no processo lógico que causou o erro.", description: "Diagnóstico técnico de causa-raiz." }
    ]
  },
  {
    id: 18,
    scenario: "Em reuniões informais ou almoços de trabalho, prefiro falar sobre:",
    options: [
      { letter: 'D', text: "Novas parcerias comerciais, objetivos ambiciosos, metas ou planos de expansão.", description: "Foco comercial direto e competitivo." },
      { letter: 'I', text: "Histórias engraçadas, viagens, cultura, hobbies ou interações cotidianas.", description: "Conversas sociais estimulantes." },
      { letter: 'S', text: "Família, bem-estar pessoal do time, relações sociais ou hobbies calmos.", description: "Conversas focadas em empatia acolhedora." },
      { letter: 'C', text: "Temas de ciência, curiosidades lógicas, tecnologia ou jogos de estratégia.", description: "Conversas ricas em conteúdo e fatos técnicos." }
    ]
  },
  {
    id: 19,
    scenario: "Ao exercer atividades de liderança, meu estilo é caracterizado por:",
    options: [
      { letter: 'D', text: "Ser exigente, delegar visando agilidade de execução e cobrar metas diretas.", description: "Orientado a ação rápida." },
      { letter: 'I', text: "Engajar o time através do entusiasmo, visão inspiradora e carisma.", description: "Líder motivador e integrador." },
      { letter: 'S', text: "Zelar pelo diálogo aberto, bem-estar da equipe e suporte próximo.", description: "Líder servidor e facilitador." },
      { letter: 'C', text: "Assegurar o cumprimento absoluto dos regulamentos e processos de qualidade.", description: "Líder normativo e focado em controle." }
    ]
  },
  {
    id: 20,
    scenario: "Minha relação profunda com regulamentos, processos e diretrizes de trabalho é:",
    options: [
      { letter: 'D', text: "Considero-os úteis, mas podem ser flexibilizados em prol de resultados ágeis.", description: "Foco no resultado em detrimento da formalidade." },
      { letter: 'I', text: "Frequentemente considero-os engessados demais, preferindo a criatividade livre.", description: "Preferencia por caminhos informais ou inovadores." },
      { letter: 'S', text: "Valorizo o respeito às regras porque elas asseguram a harmonia e evitam conflitos.", description: "Regras como garantia de segurança social." },
      { letter: 'C', text: "Acho fundamentais para estruturar a qualidade e evitar falhas operacionais graves.", description: "Conformidade rigorosa à risca." }
    ]
  }
];

// Profile Details & Baseline Reports
const BASELINE_REPORTS: Record<string, {
  perfil: string;
  resumo: string;
  pontosFortes: string[];
  pontosMelhoria: string[];
  sugestoes: string;
  lideranca: string;
}> = {
  'D': {
    perfil: "Dominância (D) - O Direcionador",
    resumo: "Indivíduos com alta Dominância são impulsionados por desafios, metas tangíveis e foco implacável em resultados. São altamente competitivos, objetivos, assertivos e independentes. No ambiente corporativo, exercem liderança natural baseada na urgência e no dinamismo. Não temem riscos e decidem com agilidade, preferindo gerir situações difíceis com pragmatismo.",
    pontosFortes: [
      "Tomada de decisão rápida sob forte pressão comercial",
      "Foco inabalável no atingimento de metas agressivas",
      "Iniciativa pioneira para desbravar novas soluções e produtos",
      "Franqueza, objetividade e comunicação direta sem rodeios",
      "Alta adaptabilidade comercial quando motivado por novos desafios"
    ],
    pontosMelhoria: [
      "Pode parecer impaciente, frio ou autoritário com membros do time",
      "Dificuldade acentuada de escuta ativa e acolhimento das ideias alheias",
      "Tendência a pular processos formais ou regras técnicas em nome da rapidez",
      "Resistência severa a ambientes de controle ou subordinação rígida",
      "Falta de tato interpessoal, podendo gerar atritos de clima organizacional"
    ],
    sugestoes: "Recomenda-se treinar ativamente a escuta empática e reconhecer o tempo de maturação da equipe. Pratique delegar com paciência, explicando o porquê dos processos e não apenas exigindo o resultado final. Desenvolver tato na comunicação evitará desgastes desnecessários nas relações profissionais essenciais.",
    lideranca: "Liderança firme, de alta cobrança por metas, delegando autonomia em troca de agilidade. Funciona muito bem liderando equipes maduras e em ambientes focados em turnarounds ou expansão de mercado."
  },
  'I': {
    perfil: "Influência (I) - O Comunicador",
    resumo: "Indivíduos com alta Influência têm sua força na sociabilidade, entusiasmo e forte carisma interpessoal. São excelentes comunicadores, motivadores, persuasivos e otimistas. Zelas pelo clima alegre do time e nutrem grande prazer em influenciar positivamente as pessoas. Possuem alto poder de engajamento, empatia expressiva e geram sinergias criativas no ambiente de trabalho.",
    pontosFortes: [
      "Excepcional comunicação interpessoal e persuasão empática",
      "Capacidade ímpar de criar networking e novas conexões profissionais",
      "Facilidade em motivar e engajar times inteiros com entusiasmo",
      "Criatividade expressiva e facilidade em propor soluções inovadoras",
      "Excelente flexibilidade interpessoal perante adversidades"
    ],
    pontosMelhoria: [
      "Dificuldade severa em manter o foco e atenção a detalhes analíticos",
      "Tendência a dispersar prazos ou negligenciar tarefas administrativas e burocráticas",
      "Alta necessidade de aprovação e validação social constante dos pares",
      "Impulsividade comercial baseada em empolgação temporária",
      "Dificuldade de dar feedbacks difíceis ou enfrentar conflitos diretos"
    ],
    sugestoes: "Recomenda-se estabelecer metodologias formais de acompanhamento de rotinas (como kanbans ou listas de tarefas estritas) e blocos de tempo livres de distração. Exercite a finalização detalhada de projetos antes de iniciar novos estímulos. Busque equilibrar a emoção carismática com análises de viabilidade lógica fria.",
    lideranca: "Liderança democrática, inspiradora e focada no engajamento coletivo. Lidera com carisma e simpatia, excelente para times de atendimento, vendas, criação e relações institucionais."
  },
  'S': {
    perfil: "Estabilidade (S) - O Planejador",
    resumo: "Indivíduos com alta Estabilidade destacam-se pela calma, lealdade, consistência e paciência admiráveis. São os grandes pacificadores e mantenedores da harmonia coletiva. Preferem ambientes de trabalho seguros, rotinas estruturadas e previsíveis. Ouvem as pessoas de forma exemplar e são consistentes no cumprimento de processos de longo prazo de forma resiliente.",
    pontosFortes: [
      "Excelente ouvinte, com empatia profunda e facilidade em integrar times",
      "Lealdade exemplar ao projeto corporativo e estabilidade emocional",
      "Método de execução linear altamente consistente de médio-longo prazo",
      "Resolução pacífica de conflitos internos com foco em harmonização",
      "Persistência e resiliência em rotinas regulares essenciais de produção"
    ],
    pontosMelhoria: [
      "Resistência instintiva a mudanças abruptas ou falta de planejamento prévio",
      "Extrema dificuldade em se posicionar contra opiniões do grupo ou dizer 'não'",
      "Hesitação e lentidão ao tomar decisões sob pressão que exijam atitude firme",
      "Tendência a reter mágoas ou feedbacks em nome de manter uma harmonia superficial",
      "Dificuldade em cobrar energicamente resultados e metas de colegas"
    ],
    sugestoes: "Recomenda-se desenvolver assertividade na comunicação e no posicionamento pessoal. Lembre-se de que feedbacks honestos e posicionamentos firmes evitam conflitos futuros mais severos. Crie o hábito de planejar contingências rápidas para lidar com cenários corporativos ágeis de forma flexível.",
    lideranca: "Liderança servidora, acolhedora e que apoia de perto os colaboradores. Excelente para liderar equipes técnicas que exigem constância operacional, operações continuadas, suporte e áreas de gestão de pessoas."
  },
  'C': {
    perfil: "Conformidade (C) - O Analista",
    resumo: "Indivíduos com alta Conformidade buscam a precisão, exatidão lógica e altos padrões de excelência. São extremamente focados em regras, processos formais, compliance e análise rigorosa de dados. Agem de forma cautelosa, racional e disciplinada. São os guardiões da qualidade estrutural, garantindo que nada saia fora dos padrões metodológicos estipulados.",
    pontosFortes: [
      "Atenção cirúrgica aos detalhes técnicos e especificações do projeto",
      "Altíssima capacidade analítica para resolução de problemas estruturais",
      "Cumprimento rigoroso de normas de segurança, compliance e procedimentos",
      "Racionalidade fria exemplar em análises lógicas complexas",
      "Busca incansável pela alta qualidade técnica, minimizando riscos"
    ],
    pontosMelhoria: [
      "Perfeccionismo excessivo que pode paralisar entregas (parálise por análise)",
      "Lentidão severa ao tomar decisões devido à busca infinita por dados extras",
      "Alta autocrítica e hipercriticismo técnico exagerado com os colegas",
      "Dificuldade extrema em aceitar erros operacionais saudáveis",
      "Isolamento interpessoal ou comunicação percebida como rígida e fria"
    ],
    sugestoes: "Recomenda-se treinar a aceitação de que o 'feito' pode ser melhor que o 'perfeito não entregue' em ambientes ágeis. Estipule prazos limites improrrogáveis para a coleta de dados e tomada de decisões. Desenvolva habilidades sociais informais, promovendo diálogos leves e focados na empatia.",
    lideranca: "Liderança baseada em fatos lógicos, normas e excelência operacional. Zela de forma exemplar pelas regras de qualidade e conformidade metodológica. Excelente para times de tecnologia, engenharia, compliance e jurídico."
  }
};

export default function DiscAvaliacaoView({ empresas, setView }: DiscEvaluationViewProps) {
  // Navigation & View States
  const [activeTab, setActiveTab] = useState<'assessments' | 'new-wizard'>('assessments');
  const [selectedEvaluation, setSelectedEvaluation] = useState<DiscEvaluation | null>(null);

  // Firestore Data State
  const [evaluations, setEvaluations] = useState<DiscEvaluation[]>([]);
  const [loading, setLoading] = useState(true);

  // Setup Wizard State
  const [wizardStep, setWizardStep] = useState(1); // 1: Setup Candidate, 2: Questionnaire, 3: Completed
  const [empresaId, setEmpresaId] = useState('');
  const [nomeAvaliado, setNomeAvaliado] = useState('');
  const [cargo, setCargo] = useState('');

  // Questionnaire Progress State
  const [currentQuestionIndex, setCurrentQuestionIndex] = useState(0);
  const [answers, setAnswers] = useState<Record<number, 'D' | 'I' | 'S' | 'C'>>({});

  // AI Generation State
  const [generatingAI, setGeneratingAI] = useState(false);
  const [aiError, setAiError] = useState<string | null>(null);

  // Fetch Saved Evaluations
  useEffect(() => {
    const user = auth.currentUser;
    if (!user) return;

    const qEvs = query(
      collection(db, 'disc_avaliacoes'),
      where('ownerId', '==', user.uid)
    );

    const unsubscribe = onSnapshot(qEvs, (snap) => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() } as DiscEvaluation));
      // Sort by date desc
      data.sort((a, b) => new Date(b.dataAvaliacao).getTime() - new Date(a.dataAvaliacao).getTime());
      setEvaluations(data);
      setLoading(false);
    }, (error) => {
      console.error("Erro ao carregar avaliações DISC:", error);
      setLoading(false);
    });

    return () => unsubscribe();
  }, []);

  // Compute Current Score Percentages
  const computePercentages = (selectedAnswers: Record<number, 'D' | 'I' | 'S' | 'C'>) => {
    const counts = { D: 0, I: 0, S: 0, C: 0 };
    Object.values(selectedAnswers).forEach(letter => {
      if (counts[letter] !== undefined) {
        counts[letter]++;
      }
    });

    const total = Object.keys(selectedAnswers).length || 1;
    return {
      scoreD: counts.D,
      scoreI: counts.I,
      scoreS: counts.S,
      scoreC: counts.C,
      percentD: Math.round((counts.D / total) * 100),
      percentI: Math.round((counts.I / total) * 100),
      percentS: Math.round((counts.S / total) * 100),
      percentC: Math.round((counts.C / total) * 100)
    };
  };

  // Determine Predominant Profile Label
  const getProfileLabel = (d: number, i: number, s: number, c: number) => {
    const scores = [
      { key: 'D', value: d, label: 'Dominância (D)' },
      { key: 'I', value: i, label: 'Influência (I)' },
      { key: 'S', value: s, label: 'Estabilidade (S)' },
      { key: 'C', value: c, label: 'Conformidade (C)' }
    ];
    scores.sort((a, b) => b.value - a.value);
    
    // Check if there is a distinct single high, or blended
    if (scores[0].value > scores[1].value + 2) {
      return scores[0].label;
    } else {
      return `Perfil Misto (${scores[0].key} / ${scores[1].key})`;
    }
  };

  // Create & Save Local Baseline Assessment
  const handleSaveAssessment = async () => {
    const user = auth.currentUser;
    if (!user) return;

    const { scoreD, scoreI, scoreS, scoreC, percentD, percentI, percentS, percentC } = computePercentages(answers);
    
    // Determine high factor
    let predominant: 'D' | 'I' | 'S' | 'C' = 'D';
    let maxVal = percentD;
    if (percentI > maxVal) { predominant = 'I'; maxVal = percentI; }
    if (percentS > maxVal) { predominant = 'S'; maxVal = percentS; }
    if (percentC > maxVal) { predominant = 'C'; maxVal = percentC; }

    const baseline = BASELINE_REPORTS[predominant];
    const profileLabel = getProfileLabel(percentD, percentI, percentS, percentC);

    // Form baseline report string
    const baselineRelatorio = `### RESUMO EXECUTIVO DO PERFIL\n${baseline.resumo}\n\n### PRINCIPAIS FORÇAS COMPORTAMENTAIS\n${baseline.pontosFortes.map(f => `- ${f}`).join('\n')}\n\n### ESTILO DE LIDERANÇA\n${baseline.lideranca}`;
    const baselineSugestoes = `### SUGESTÕES DE MELHORIA COMPORTAMENTAL\n${baseline.sugestoes}\n\n### OPORTUNIDADES DE DESENVOLVIMENTO\n${baseline.pontosMelhoria.map(m => `- ${m}`).join('\n')}`;

    const newEvaluation: DiscEvaluation = {
      empresaId,
      nomeAvaliado,
      cargo,
      dataAvaliacao: new Date().toISOString(),
      ownerId: user.uid,
      scoreD: percentD,
      scoreI: percentI,
      scoreS: percentS,
      scoreC: percentC,
      perfilPrincipal: profileLabel,
      relatorio: baselineRelatorio,
      sugestoesMelhoria: baselineSugestoes
    };

    try {
      const docRef = await addDoc(collection(db, 'disc_avaliacoes'), newEvaluation);
      const savedDoc = { id: docRef.id, ...newEvaluation };
      setSelectedEvaluation(savedDoc);
      setWizardStep(3); // Go to completed success tab
    } catch (err) {
      console.error("Erro ao salvar avaliação DISC:", err);
    }
  };

  // Generate Advanced Report using Gemini API Proxy
  const handleRefineWithAI = async (evalId: string) => {
    if (!selectedEvaluation) return;
    setGeneratingAI(true);
    setAiError(null);

    const empresaNome = empresas.find(e => e.id === selectedEvaluation.empresaId)?.nome || "Não identificada";

    const prompt = `Você é um psicólogo organizacional experiente e especialista na metodologia DISC.
Gere um relatório comportamental detalhado e sugestões práticas de melhoria (Plano de Desenvolvimento Individual - PDI) para o profissional ${selectedEvaluation.nomeAvaliado}, que atua como ${selectedEvaluation.cargo || "Profissional"} na empresa "${empresaNome}".

Os resultados do teste DISC dele foram:
- Dominância (D): ${selectedEvaluation.scoreD}% (Foco em desafios, decisões rápidos, assertividade)
- Influência (I): ${selectedEvaluation.scoreI}% (Foco em comunicação, relacionamentos, carisma, otimismo)
- Estabilidade (S): ${selectedEvaluation.scoreS}% (Foco em cooperação, estabilidade, consistência, paciência)
- Conformidade (C): ${selectedEvaluation.scoreC}% (Foco em lógica, regras, precisão, qualidade)

O perfil principal dele é classificado como "${selectedEvaluation.perfilPrincipal}".

Por favor, escreva o relatório em português brasileiro (PT-BR) contendo as seguintes seções claramente divididas com títulos elegantes (usando formatação Markdown h3 como ###):
1. Resumo Executivo do Perfil (Análise rica e elegante sobre a atitude geral dele e dinâmica de tomada de decisão)
2. Principais Forças Comportamentais (Aprofundamento sobre como ele brilha no ambiente de trabalho)
3. Oportunidades de Desenvolvimento e Desafios (Sua faceta de pontos cegos ou fraquezas em ambientes de estresse)
4. Guia de Convivência e Comunicação (Como colegas e gestores podem colaborar melhor e se comunicar com eficácia com ele)
5. Plano de Desenvolvimento Individual (PDI) Sugerido (Passos acionáveis divididos em curto, médio e longo prazo para alavancar a carreira e o equilíbrio dele)

Retorne estritamente o relatório estruturado em formato Markdown de alta qualidade profissional, sem cabeçalhos irrelevantes ou metadados de sistema.`;

    try {
      const response = await fetch("/api/gemini/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "gemini-2.5-flash",
          contents: prompt
        })
      });

      if (!response.ok) {
        throw new Error("Erro na resposta do servidor de IA");
      }

      const data = await response.json();
      if (data.error) {
        throw new Error(data.error);
      }

      if (data.text) {
        // Split text in two parts or keep it fully inside 'relatorio' and suggestions
        const fullAIResponse = data.text;
        
        // Split by some common markers, or just keep it structured
        let aiRelatorio = fullAIResponse;
        let aiSugestoes = "### PLANO DE DESENVOLVIMENTO INDIVIDUAL (PDI)\nConsulte o relatório gerado pela IA acima para o plano completo de desenvolvimento.";

        if (fullAIResponse.includes("### Plano de Desenvolvimento Individual") || fullAIResponse.includes("### PLANO DE DESENVOLVIMENTO INDIVIDUAL")) {
          const parts = fullAIResponse.split(/### Plano de Desenvolvimento Individual|### PLANO DE DESENVOLVIMENTO INDIVIDUAL/i);
          aiRelatorio = parts[0];
          aiSugestoes = "### PLANO DE DESENVOLVIMENTO INDIVIDUAL (PDI)" + parts[1];
        }

        const docRef = doc(db, 'disc_avaliacoes', evalId);
        await updateDoc(docRef, {
          relatorio: aiRelatorio,
          sugestoesMelhoria: aiSugestoes
        });

        setSelectedEvaluation(prev => {
          if (!prev) return null;
          return {
            ...prev,
            relatorio: aiRelatorio,
            sugestoesMelhoria: aiSugestoes
          };
        });
      }
    } catch (err: any) {
      console.error("Erro ao refinar relatório com Gemini:", err);
      setAiError(err?.message || "Não foi possível obter uma resposta do servidor da IA. Verifique as configurações de sua chave de API.");
    } finally {
      setGeneratingAI(false);
    }
  };

  // Handle Deletion
  const handleDeleteEvaluation = async (evalId: string, event: React.MouseEvent) => {
    event.stopPropagation();
    if (!window.confirm("Deseja realmente excluir esta avaliação comportamental?")) return;

    try {
      await deleteDoc(doc(db, 'disc_avaliacoes', evalId));
      if (selectedEvaluation?.id === evalId) {
        setSelectedEvaluation(null);
      }
    } catch (err) {
      console.error("Erro ao excluir avaliação:", err);
    }
  };

  // Reset & Create New Wizard
  const startNewWizard = () => {
    setEmpresaId(empresas[0]?.id || '');
    setNomeAvaliado('');
    setCargo('');
    setWizardStep(1);
    setCurrentQuestionIndex(0);
    setAnswers({});
    setActiveTab('new-wizard');
  };

  // PDF Export using html2canvas & jsPDF
  const exportToPDF = () => {
    const element = document.getElementById('disc-report-container');
    if (!element) return;

    const opt = {
      margin: 10,
      filename: `Relatorio_DISC_${selectedEvaluation?.nomeAvaliado || 'Avaliado'}.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    html2canvas(element, { scale: 2 }).then((canvas) => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgWidth = 190; // Page width in mm minus margins
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

  // Render DISC custom wheel coordinate generation & drawing
  const renderDiscWheelSVG = (dScore: number, iScore: number, sScore: number, cScore: number) => {
    const size = 360;
    const center = size / 2;
    const maxRadius = 130;

    // Convert percentage score to length (0 to maxRadius)
    const lenD = (dScore / 100) * maxRadius;
    const lenI = (iScore / 100) * maxRadius;
    const lenS = (sScore / 100) * maxRadius;
    const lenC = (cScore / 100) * maxRadius;

    // Angles for the quadrants (offset to fit standard coordinate math)
    // D: Top-Left (225 deg)
    const radD = (225 * Math.PI) / 180;
    const xD = center + lenD * Math.cos(radD);
    const yD = center + lenD * Math.sin(radD);

    // I: Top-Right (315 deg)
    const radI = (315 * Math.PI) / 180;
    const xI = center + lenI * Math.cos(radI);
    const yI = center + lenI * Math.sin(radI);

    // S: Bottom-Right (45 deg)
    const radS = (45 * Math.PI) / 180;
    const xS = center + lenS * Math.cos(radS);
    const yS = center + lenS * Math.sin(radS);

    // C: Bottom-Left (135 deg)
    const radC = (135 * Math.PI) / 180;
    const xC = center + lenC * Math.cos(radC);
    const yC = center + lenC * Math.sin(radC);

    return (
      <svg width="100%" height="100%" viewBox={`0 0 ${size} ${size}`} className="max-w-[340px] mx-auto select-none drop-shadow-md">
        {/* Definitions for gradients and drop-shadows */}
        <defs>
          <filter id="shadow" x="-10%" y="-10%" width="120%" height="120%">
            <feDropShadow dx="0" dy="2" stdDeviation="3" floodOpacity="0.1" />
          </filter>
        </defs>

        {/* Circular Ring background */}
        <circle cx={center} cy={center} r={maxRadius} fill="#f8fafc" stroke="#e2e8f0" strokeWidth="2" />
        <circle cx={center} cy={center} r={maxRadius - 35} fill="none" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="4 4" />
        <circle cx={center} cy={center} r={maxRadius - 70} fill="none" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="4 4" />
        <circle cx={center} cy={center} r={maxRadius - 105} fill="none" stroke="#f1f5f9" strokeWidth="1" strokeDasharray="4 4" />

        {/* Quadrant Segments background shading to match DISC wheel visual */}
        {/* D - Top-Left: Red */}
        <path d={`M ${center} ${center} L ${center - maxRadius} ${center} A ${maxRadius} ${maxRadius} 0 0 1 ${center} ${center - maxRadius} Z`} fill="rgba(244, 63, 94, 0.05)" />
        {/* I - Top-Right: Yellow */}
        <path d={`M ${center} ${center} L ${center} ${center - maxRadius} A ${maxRadius} ${maxRadius} 0 0 1 ${center + maxRadius} ${center} Z`} fill="rgba(234, 179, 8, 0.05)" />
        {/* S - Bottom-Right: Green */}
        <path d={`M ${center} ${center} L ${center + maxRadius} ${center} A ${maxRadius} ${maxRadius} 0 0 1 ${center} ${center + maxRadius} Z`} fill="rgba(34, 197, 94, 0.05)" />
        {/* C - Bottom-Left: Blue/Cyan */}
        <path d={`M ${center} ${center} L ${center} ${center + maxRadius} A ${maxRadius} ${maxRadius} 0 0 1 ${center - maxRadius} ${center} Z`} fill="rgba(6, 182, 212, 0.05)" />

        {/* Main Axes dividing the segments */}
        <line x1={center - maxRadius - 10} y1={center} x2={center + maxRadius + 10} y2={center} stroke="#cbd5e1" strokeWidth="1.5" />
        <line x1={center} y1={center - maxRadius - 10} x2={center} y2={center + maxRadius + 10} stroke="#cbd5e1" strokeWidth="1.5" />

        {/* Sub-Diagonal guidelines mapping to center of segments */}
        <line x1={center - maxRadius * 0.8} y1={center - maxRadius * 0.8} x2={center + maxRadius * 0.8} y2={center + maxRadius * 0.8} stroke="#e2e8f0" strokeWidth="1" strokeDasharray="3 3" />
        <line x1={center - maxRadius * 0.8} y1={center + maxRadius * 0.8} x2={center + maxRadius * 0.8} y2={center - maxRadius * 0.8} stroke="#e2e8f0" strokeWidth="1" strokeDasharray="3 3" />

        {/* Plotting the User's Spider Polygon Profile Overlay */}
        <polygon 
          points={`${xD},${yD} ${xI},${yI} ${xS},${yS} ${xC},${yC}`} 
          fill="rgba(14, 165, 233, 0.28)" 
          stroke="rgb(2, 132, 199)" 
          strokeWidth="2.5" 
          strokeLinejoin="round"
          className="transition-all duration-700 ease-out"
        />

        {/* Center coordinate point */}
        <circle cx={center} cy={center} r="4" fill="#0284c7" />

        {/* Quadrant Titles */}
        {/* D (Dominância) */}
        <text x={center - 55} y={center - 55} fontSize="14" fontWeight="700" fill="#f43f5e" textAnchor="middle">D</text>
        <text x={center - 55} y={center - 38} fontSize="8" fontWeight="500" fill="#ef4444" textAnchor="middle">RESULTADOS</text>
        
        {/* I (Influência) */}
        <text x={center + 55} y={center - 55} fontSize="14" fontWeight="700" fill="#eab308" textAnchor="middle">I</text>
        <text x={center + 55} y={center - 38} fontSize="8" fontWeight="500" fill="#ca8a04" textAnchor="middle">ENTUSIASMO</text>

        {/* S (Estabilidade) */}
        <text x={center + 55} y={center + 45} fontSize="14" fontWeight="700" fill="#22c55e" textAnchor="middle">S</text>
        <text x={center + 55} y={center + 60} fontSize="8" fontWeight="500" fill="#16a34a" textAnchor="middle">APOIO</text>

        {/* C (Conformidade) */}
        <text x={center - 55} y={center + 45} fontSize="14" fontWeight="700" fill="#06b6d4" textAnchor="middle">C</text>
        <text x={center - 55} y={center + 60} fontSize="8" fontWeight="500" fill="#0891b2" textAnchor="middle">PRECISÃO</text>

        {/* Axis Labels (On the outer poles) */}
        <rect x={center - 24} y={center - maxRadius - 15} width="48" height="14" rx="4" fill="white" stroke="#e2e8f0" strokeWidth="1" />
        <text x={center} y={center - maxRadius - 5} fontSize="8" fontWeight="700" fill="#334155" textAnchor="middle">AÇÃO</text>

        <rect x={center + maxRadius - 5} y={center - 7} width="70" height="14" rx="4" fill="white" stroke="#e2e8f0" strokeWidth="1" />
        <text x={center + maxRadius + 30} y={center + 3} fontSize="8" fontWeight="700" fill="#334155" textAnchor="middle">COLABORAÇÃO</text>

        <rect x={center - 35} y={center + maxRadius + 2} width="70" height="14" rx="4" fill="white" stroke="#e2e8f0" strokeWidth="1" />
        <text x={center} y={center + maxRadius + 12} fontSize="8" fontWeight="700" fill="#334155" textAnchor="middle">SEGURANÇA</text>

        <rect x={center - maxRadius - 65} y={center - 7} width="60" height="14" rx="4" fill="white" stroke="#e2e8f0" strokeWidth="1" />
        <text x={center - maxRadius - 35} y={center + 3} fontSize="8" fontWeight="700" fill="#334155" textAnchor="middle">DESAFIOS</text>

        {/* Custom plotted score point vertices & markers */}
        <circle cx={xD} cy={yD} r="5" fill="#f43f5e" stroke="white" strokeWidth="1.5" />
        <text x={xD - 10} y={yD - 4} fontSize="9" fontWeight="700" fill="#334155" textAnchor="end">{dScore}%</text>

        <circle cx={xI} cy={yI} r="5" fill="#eab308" stroke="white" strokeWidth="1.5" />
        <text x={xI + 10} y={yI - 4} fontSize="9" fontWeight="700" fill="#334155" textAnchor="start">{iScore}%</text>

        <circle cx={xS} cy={yS} r="5" fill="#22c55e" stroke="white" strokeWidth="1.5" />
        <text x={xS + 10} y={yS + 10} fontSize="9" fontWeight="700" fill="#334155" textAnchor="start">{sScore}%</text>

        <circle cx={xC} cy={yC} r="5" fill="#06b6d4" stroke="white" strokeWidth="1.5" />
        <text x={xC - 10} y={yC + 10} fontSize="9" fontWeight="700" fill="#334155" textAnchor="end">{cScore}%</text>
      </svg>
    );
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
      {/* Header and Menu Tabs */}
      <div className="flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-100 pb-5">
        <div>
          <h1 className="text-3xl font-bold text-slate-900 tracking-tight flex items-center gap-3">
            <Award className="h-8 w-8 text-sky-600" />
            Avaliação Comportamental DISC
          </h1>
          <p className="mt-2 text-sm text-slate-500 max-w-2xl">
            Realize testes de perfil de comportamento para os diretores e fundadores dos seus clientes. Entenda suas reações, competências, estilo de liderança e gere um PDI completo.
          </p>
        </div>
        <div className="mt-4 md:mt-0 flex gap-3">
          <button
            onClick={() => { setActiveTab('assessments'); setSelectedEvaluation(null); }}
            className={`px-4 py-2 text-sm font-semibold rounded-xl transition-all ${
              activeTab === 'assessments' && !selectedEvaluation
                ? 'bg-sky-600 text-white shadow-md'
                : 'bg-white text-slate-600 hover:bg-slate-50 border border-slate-200'
            }`}
          >
            Histórico de Avaliações
          </button>
          <button
            onClick={startNewWizard}
            className={`px-4 py-2 text-sm font-semibold rounded-xl transition-all flex items-center gap-2 ${
              activeTab === 'new-wizard'
                ? 'bg-sky-600 text-white shadow-md'
                : 'bg-sky-50 text-sky-600 hover:bg-sky-100'
            }`}
          >
            <Plus className="h-4 w-4" />
            Nova Avaliação
          </button>
        </div>
      </div>

      {/* Main Container Workspace */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
        {/* Left Side: Navigation history list (Always shown unless active test wizard is running) */}
        {activeTab === 'assessments' && !selectedEvaluation && (
          <div className="col-span-12 space-y-6">
            {loading ? (
              <div className="bg-white rounded-3xl p-12 text-center border border-slate-100 shadow-sm flex flex-col items-center justify-center gap-4">
                <RefreshCw className="h-8 w-8 text-sky-500 animate-spin" />
                <p className="text-slate-500 text-sm">Carregando avaliações comportamentais...</p>
              </div>
            ) : evaluations.length === 0 ? (
              <div className="bg-white rounded-3xl p-16 text-center border border-slate-100 shadow-sm">
                <div className="h-16 w-16 bg-slate-50 text-slate-400 rounded-2xl flex items-center justify-center mx-auto mb-6">
                  <User className="h-8 w-8" />
                </div>
                <h3 className="text-lg font-bold text-slate-800">Nenhuma avaliação cadastrada</h3>
                <p className="mt-2 text-slate-500 text-sm max-w-sm mx-auto">
                  Inicie o mapeamento de perfil comportamental DISC dos colaboradores das suas empresas parceiras hoje mesmo.
                </p>
                <button
                  onClick={startNewWizard}
                  className="mt-6 inline-flex items-center gap-2 px-5 py-2.5 text-sm font-semibold bg-sky-600 text-white rounded-xl hover:bg-sky-700 shadow-sm transition-all"
                >
                  <Plus className="h-4 w-4" />
                  Iniciar Primeira Avaliação
                </button>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {evaluations.map((ev) => {
                  const empresaNome = empresas.find(e => e.id === ev.empresaId)?.nome || "Empresa Cliente";
                  return (
                    <div
                      key={ev.id}
                      onClick={() => setSelectedEvaluation(ev)}
                      className="bg-white rounded-[2rem] p-6 border border-slate-100 shadow-sm hover:shadow-md hover:border-sky-100 transition-all cursor-pointer flex flex-col justify-between group relative overflow-hidden"
                    >
                      <div className="space-y-4">
                        <div className="flex justify-between items-start">
                          <span className="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-sky-50 text-sky-600">
                            {ev.perfilPrincipal}
                          </span>
                          <button
                            onClick={(e) => handleDeleteEvaluation(ev.id!, e)}
                            className="p-1.5 text-slate-300 hover:text-rose-500 hover:bg-rose-50 rounded-lg transition-all"
                            title="Excluir avaliação"
                          >
                            <Trash2 className="h-4 w-4" />
                          </button>
                        </div>

                        <div>
                          <h4 className="text-lg font-bold text-slate-800 group-hover:text-sky-600 transition-colors">
                            {ev.nomeAvaliado}
                          </h4>
                          {ev.cargo && (
                            <p className="text-xs font-semibold text-slate-400 flex items-center gap-1.5 mt-1">
                              <Briefcase className="h-3 w-3 text-slate-300" />
                              {ev.cargo}
                            </p>
                          )}
                          <p className="text-xs text-slate-500 flex items-center gap-1.5 mt-2">
                            <Building2 className="h-3.5 w-3.5 text-slate-400" />
                            {empresaNome}
                          </p>
                        </div>

                        {/* Miniature Spark scores */}
                        <div className="grid grid-cols-4 gap-2 pt-2 border-t border-slate-50 text-center text-xs font-bold">
                          <div className="bg-rose-50/50 rounded-lg py-1 text-rose-600">D: {ev.scoreD}%</div>
                          <div className="bg-amber-50/50 rounded-lg py-1 text-amber-600">I: {ev.scoreI}%</div>
                          <div className="bg-emerald-50/50 rounded-lg py-1 text-emerald-600">S: {ev.scoreS}%</div>
                          <div className="bg-cyan-50/50 rounded-lg py-1 text-cyan-600">C: {ev.scoreC}%</div>
                        </div>
                      </div>

                      <div className="mt-5 pt-4 border-t border-slate-50 flex items-center justify-between text-xs text-slate-400">
                        <span className="flex items-center gap-1">
                          <Calendar className="h-3.5 w-3.5 text-slate-300" />
                          {new Date(ev.dataAvaliacao).toLocaleDateString('pt-BR')}
                        </span>
                        <span className="text-sky-600 font-bold group-hover:translate-x-1 transition-transform inline-flex items-center gap-0.5">
                          Ver Relatório <ChevronRight className="h-3.5 w-3.5" />
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* Wizard Panel: Conducting a test */}
        {activeTab === 'new-wizard' && (
          <div className="col-span-12 lg:col-span-8 bg-white rounded-3xl p-6 sm:p-8 border border-slate-100 shadow-sm">
            {/* Step 1: Mapear Candidato */}
            {wizardStep === 1 && (
              <div className="space-y-6">
                <div>
                  <h3 className="text-xl font-bold text-slate-800">1. Identificação do Avaliado</h3>
                  <p className="text-sm text-slate-500 mt-1">Insira as informações do profissional que responderá o questionário behaviorista.</p>
                </div>

                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold uppercase text-slate-400 mb-2">Empresa Cliente</label>
                    <select
                      value={empresaId}
                      onChange={(e) => setEmpresaId(e.target.value)}
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-sky-500/20 focus:border-sky-600 text-sm text-slate-700 bg-white"
                    >
                      <option value="" disabled>Selecione uma empresa</option>
                      {empresas.map((emp) => (
                        <option key={emp.id} value={emp.id}>{emp.nome}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-slate-400 mb-2">Nome Completo do Avaliado</label>
                    <input
                      type="text"
                      placeholder="Ex: João Silva"
                      value={nomeAvaliado}
                      onChange={(e) => setNomeAvaliado(e.target.value)}
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-sky-500/20 focus:border-sky-600 text-sm text-slate-700"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold uppercase text-slate-400 mb-2">Cargo / Função (Opcional)</label>
                    <input
                      type="text"
                      placeholder="Ex: Diretor de Operações"
                      value={cargo}
                      onChange={(e) => setCargo(e.target.value)}
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-sky-500/20 focus:border-sky-600 text-sm text-slate-700"
                    />
                  </div>
                </div>

                <div className="pt-4 flex justify-between">
                  <button
                    onClick={() => setActiveTab('assessments')}
                    className="px-5 py-2.5 text-sm font-semibold bg-white border border-slate-200 text-slate-600 rounded-xl hover:bg-slate-50 transition-all"
                  >
                    Cancelar
                  </button>
                  <button
                    onClick={() => {
                      if (!empresaId) { alert("Por favor, selecione uma empresa."); return; }
                      if (!nomeAvaliado.trim()) { alert("Por favor, preencha o nome do avaliado."); return; }
                      setWizardStep(2);
                    }}
                    className="px-5 py-2.5 text-sm font-semibold bg-sky-600 text-white rounded-xl hover:bg-sky-700 shadow-sm transition-all inline-flex items-center gap-2"
                  >
                    Ir para o Questionário
                    <ArrowRight className="h-4 w-4" />
                  </button>
                </div>
              </div>
            )}

            {/* Step 2: Questionnaire Wizard */}
            {wizardStep === 2 && (
              <div className="space-y-6">
                <div className="flex justify-between items-center pb-4 border-b border-slate-50">
                  <div>
                    <span className="text-xs font-bold uppercase tracking-wider text-sky-600">Questão {currentQuestionIndex + 1} de {DISC_QUESTIONS.length}</span>
                    <h3 className="text-base font-bold text-slate-800 mt-1">Avaliação em Andamento: {nomeAvaliado}</h3>
                  </div>
                  <div className="w-24 h-2 bg-slate-100 rounded-full overflow-hidden">
                    <div 
                      className="bg-sky-600 h-full transition-all duration-300"
                      style={{ width: `${((currentQuestionIndex + 1) / DISC_QUESTIONS.length) * 100}%` }}
                    />
                  </div>
                </div>

                {/* Scenario Title */}
                <div className="bg-slate-50 rounded-2xl p-5 border border-slate-100/50">
                  <p className="text-base font-semibold text-slate-700 leading-relaxed">
                    {DISC_QUESTIONS[currentQuestionIndex].scenario}
                  </p>
                </div>

                {/* 4 Options */}
                <div className="space-y-3.5">
                  {DISC_QUESTIONS[currentQuestionIndex].options.map((opt) => {
                    const isSelected = answers[DISC_QUESTIONS[currentQuestionIndex].id] === opt.letter;
                    return (
                      <div
                        key={opt.letter}
                        onClick={() => {
                          setAnswers(prev => ({
                            ...prev,
                            [DISC_QUESTIONS[currentQuestionIndex].id]: opt.letter
                          }));
                        }}
                        className={`p-4 rounded-2xl border transition-all cursor-pointer flex items-start gap-4 hover:shadow-sm ${
                          isSelected
                            ? 'bg-sky-50/50 border-sky-200 ring-2 ring-sky-500/10'
                            : 'bg-white border-slate-200 hover:border-slate-300'
                        }`}
                      >
                        <div className={`mt-0.5 h-5 w-5 rounded-full border-2 flex items-center justify-center shrink-0 ${
                          isSelected ? 'border-sky-600 bg-sky-600 text-white' : 'border-slate-300 bg-white'
                        }`}>
                          {isSelected && <Check className="h-3 w-3 stroke-[3]" />}
                        </div>
                        <div>
                          <p className={`text-sm font-semibold leading-relaxed ${isSelected ? 'text-sky-900' : 'text-slate-700'}`}>
                            {opt.text}
                          </p>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Wizard Footers */}
                <div className="pt-4 flex justify-between border-t border-slate-100">
                  <button
                    disabled={currentQuestionIndex === 0}
                    onClick={() => setCurrentQuestionIndex(prev => prev - 1)}
                    className="px-4 py-2.5 text-sm font-semibold border border-slate-200 rounded-xl hover:bg-slate-50 text-slate-600 disabled:opacity-40 disabled:cursor-not-allowed transition-all inline-flex items-center gap-2"
                  >
                    <ArrowLeft className="h-4 w-4" />
                    Anterior
                  </button>

                  {currentQuestionIndex < DISC_QUESTIONS.length - 1 ? (
                    <button
                      disabled={!answers[DISC_QUESTIONS[currentQuestionIndex].id]}
                      onClick={() => setCurrentQuestionIndex(prev => prev + 1)}
                      className="px-5 py-2.5 text-sm font-semibold bg-sky-600 text-white rounded-xl hover:bg-sky-700 disabled:opacity-40 disabled:cursor-not-allowed transition-all inline-flex items-center gap-2"
                    >
                      Próxima
                      <ArrowRight className="h-4 w-4" />
                    </button>
                  ) : (
                    <button
                      disabled={Object.keys(answers).length < DISC_QUESTIONS.length}
                      onClick={handleSaveAssessment}
                      className="px-6 py-2.5 text-sm font-bold bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed transition-all inline-flex items-center gap-2 shadow-md"
                    >
                      <Check className="h-4 w-4" />
                      Finalizar e Salvar
                    </button>
                  )}
                </div>
              </div>
            )}

            {/* Step 3: Success Completed Mapped */}
            {wizardStep === 3 && selectedEvaluation && (
              <div className="text-center py-10 space-y-6">
                <div className="h-16 w-16 bg-emerald-50 text-emerald-600 rounded-full flex items-center justify-center mx-auto shadow-inner">
                  <Check className="h-8 w-8 stroke-[3]" />
                </div>
                <div>
                  <h3 className="text-2xl font-bold text-slate-800">Avaliação Cadastrada com Sucesso!</h3>
                  <p className="text-sm text-slate-500 mt-2 max-w-md mx-auto">
                    Os dados comportamentais básicos para <strong>{selectedEvaluation.nomeAvaliado}</strong> foram computados. Deseja refinar o relatório agora usando Inteligência Artificial?
                  </p>
                </div>

                <div className="flex flex-col sm:flex-row justify-center gap-3 pt-4">
                  <button
                    onClick={() => {
                      setActiveTab('assessments');
                    }}
                    className="px-5 py-2.5 text-sm font-semibold border border-slate-200 text-slate-600 rounded-xl hover:bg-slate-50 transition-all"
                  >
                    Ir para Histórico
                  </button>
                  <button
                    onClick={() => handleRefineWithAI(selectedEvaluation.id!)}
                    disabled={generatingAI}
                    className="px-5 py-2.5 text-sm font-semibold bg-sky-600 hover:bg-sky-700 text-white rounded-xl shadow-md transition-all inline-flex items-center justify-center gap-2 disabled:opacity-50"
                  >
                    {generatingAI ? (
                      <>
                        <RefreshCw className="h-4 w-4 animate-spin" />
                        Gerando Relatório com IA...
                      </>
                    ) : (
                      <>
                        <Sparkles className="h-4 w-4" />
                        Gerar Diagnóstico Avançado com IA
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Info Right Sidebar for Wizard explaining the dimensions */}
        {activeTab === 'new-wizard' && wizardStep === 2 && (
          <div className="col-span-12 lg:col-span-4 bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-6">
            <div>
              <h4 className="text-sm font-bold uppercase tracking-wider text-slate-400">Dimensões DISC</h4>
              <p className="text-xs text-slate-500 mt-1">A metodologia analisa o comportamento focado em quatro quadrantes:</p>
            </div>

            <div className="space-y-4">
              <div className="border-l-4 border-rose-500 pl-3">
                <h5 className="text-sm font-bold text-rose-600">D - Dominância</h5>
                <p className="text-xs text-slate-500 mt-0.5">Determina como a pessoa lida com problemas, barreiras e grandes desafios. Foco em resultados imediatos.</p>
              </div>

              <div className="border-l-4 border-amber-500 pl-3">
                <h5 className="text-sm font-bold text-amber-600">I - Influência</h5>
                <p className="text-xs text-slate-500 mt-0.5">Determina como a pessoa lida com contatos interpessoais, parcerias e comunicação. Foco em otimismo e networking.</p>
              </div>

              <div className="border-l-4 border-emerald-500 pl-3">
                <h5 className="text-sm font-bold text-emerald-600">S - Estabilidade</h5>
                <p className="text-xs text-slate-500 mt-0.5">Determina como a pessoa lida com o ritmo, constância e as mudanças. Foco em previsibilidade e paciência.</p>
              </div>

              <div className="border-l-4 border-cyan-500 pl-3">
                <h5 className="text-sm font-bold text-cyan-600">C - Conformidade</h5>
                <p className="text-xs text-slate-500 mt-0.5">Determina como a pessoa se adéqua a normas, regras de compliance e critérios técnicos. Foco em exatidão lógica.</p>
              </div>
            </div>
          </div>
        )}

        {/* Selected Evaluation Report Detail View */}
        {selectedEvaluation && (
          <div className="col-span-12 space-y-6">
            {/* Action Bar */}
            <div className="bg-slate-50 rounded-2xl p-4 border border-slate-100 flex flex-wrap justify-between items-center gap-3">
              <button
                onClick={() => setSelectedEvaluation(null)}
                className="px-4 py-2 text-sm font-semibold bg-white border border-slate-200 text-slate-600 rounded-xl hover:bg-slate-50 transition-all inline-flex items-center gap-2"
              >
                <ArrowLeft className="h-4 w-4" />
                Voltar ao Histórico
              </button>

              <div className="flex gap-2.5">
                <button
                  onClick={() => handleRefineWithAI(selectedEvaluation.id!)}
                  disabled={generatingAI}
                  className="px-4 py-2 text-sm font-semibold bg-sky-50 text-sky-600 hover:bg-sky-100 rounded-xl transition-all inline-flex items-center gap-2 disabled:opacity-50"
                  title="Enriquece ou reescreve o relatório com análises avançadas da Inteligência Artificial do Gemini"
                >
                  {generatingAI ? (
                    <>
                      <RefreshCw className="h-4 w-4 animate-spin" />
                      Refinando...
                    </>
                  ) : (
                    <>
                      <Sparkles className="h-4 w-4 text-sky-600" />
                      Refinar com IA (Gemini)
                    </>
                  )}
                </button>
                <button
                  onClick={exportToPDF}
                  className="px-4 py-2 text-sm font-semibold bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 rounded-xl transition-all inline-flex items-center gap-2"
                >
                  <Download className="h-4 w-4" />
                  Baixar PDF
                </button>
                <button
                  onClick={() => {
                    const originalTitle = document.title;
                    document.title = `Relatório DISC - ${selectedEvaluation?.nomeAvaliado || 'Avaliado'}`;
                    window.print();
                    setTimeout(() => {
                      document.title = originalTitle;
                    }, 100);
                  }}
                  className="px-4 py-2 text-sm font-semibold bg-white border border-slate-200 text-slate-700 hover:bg-slate-50 rounded-xl transition-all inline-flex items-center gap-2"
                >
                  <Printer className="h-4 w-4" />
                  Imprimir
                </button>
              </div>
            </div>

            {aiError && (
              <div className="bg-rose-50 border border-rose-100 text-rose-700 p-4 rounded-2xl text-sm flex items-center gap-3">
                <Info className="h-5 w-5 shrink-0" />
                <p>{aiError}</p>
              </div>
            )}

            {/* Document Content Box */}
            <div 
              id="disc-report-container" 
              className="bg-white rounded-[2.5rem] border border-slate-100 shadow-md p-6 sm:p-10 md:p-14 space-y-10 print:shadow-none print:border-none print:p-0"
            >
              {/* Report Header */}
              <div className="flex flex-col md:flex-row justify-between items-start border-b border-slate-100 pb-8 gap-6">
                <div>
                  <span className="inline-flex items-center px-3.5 py-1.5 rounded-full text-xs font-bold tracking-wider uppercase bg-sky-50 text-sky-600">
                    {selectedEvaluation.perfilPrincipal}
                  </span>
                  <h2 className="text-3xl font-bold text-slate-900 tracking-tight mt-3">
                    Relatório de Análise Comportamental DISC
                  </h2>
                  <p className="text-slate-500 text-sm mt-1">
                    Avaliação diagnóstica de competências profissionais e perfil comportamental.
                  </p>
                </div>
                <div className="space-y-1.5 text-xs text-slate-500 text-left md:text-right">
                  <p className="font-bold text-slate-800">Avaliado: <span className="font-normal text-slate-600">{selectedEvaluation.nomeAvaliado}</span></p>
                  {selectedEvaluation.cargo && <p className="font-bold text-slate-800">Cargo: <span className="font-normal text-slate-600">{selectedEvaluation.cargo}</span></p>}
                  <p className="font-bold text-slate-800">Empresa: <span className="font-normal text-slate-600">{empresas.find(e => e.id === selectedEvaluation.empresaId)?.nome || "Não identificada"}</span></p>
                  <p className="font-bold text-slate-800">Data: <span className="font-normal text-slate-600">{new Date(selectedEvaluation.dataAvaliacao).toLocaleDateString('pt-BR')}</span></p>
                </div>
              </div>

              {/* Graphic Matrix and Circular Wheel */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-center pt-2">
                {/* SVG Circular Visual Matrix */}
                <div className="lg:col-span-5 flex justify-center">
                  <div className="w-full max-w-[340px]">
                    {renderDiscWheelSVG(
                      selectedEvaluation.scoreD,
                      selectedEvaluation.scoreI,
                      selectedEvaluation.scoreS,
                      selectedEvaluation.scoreC
                    )}
                  </div>
                </div>

                {/* Score breakdown metrics side-by-side */}
                <div className="lg:col-span-7 space-y-4">
                  <h4 className="text-sm font-bold uppercase tracking-wider text-slate-400 mb-2">Composição dos Fatores</h4>
                  
                  {/* Dominance D */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between text-xs font-bold">
                      <span className="text-rose-600">D - Dominância (Resultados / Desafios)</span>
                      <span className="text-slate-700">{selectedEvaluation.scoreD}%</span>
                    </div>
                    <div className="w-full h-2.5 bg-slate-50 rounded-full overflow-hidden border border-slate-100">
                      <div className="bg-rose-500 h-full rounded-full transition-all duration-700" style={{ width: `${selectedEvaluation.scoreD}%` }} />
                    </div>
                  </div>

                  {/* Influence I */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between text-xs font-bold">
                      <span className="text-amber-600">I - Influência (Conexões / Entusiasmo)</span>
                      <span className="text-slate-700">{selectedEvaluation.scoreI}%</span>
                    </div>
                    <div className="w-full h-2.5 bg-slate-50 rounded-full overflow-hidden border border-slate-100">
                      <div className="bg-amber-500 h-full rounded-full transition-all duration-700" style={{ width: `${selectedEvaluation.scoreI}%` }} />
                    </div>
                  </div>

                  {/* Steadiness S */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between text-xs font-bold">
                      <span className="text-emerald-600">S - Estabilidade (Ritmo / Apoio)</span>
                      <span className="text-slate-700">{selectedEvaluation.scoreS}%</span>
                    </div>
                    <div className="w-full h-2.5 bg-slate-50 rounded-full overflow-hidden border border-slate-100">
                      <div className="bg-emerald-500 h-full rounded-full transition-all duration-700" style={{ width: `${selectedEvaluation.scoreS}%` }} />
                    </div>
                  </div>

                  {/* Compliance C */}
                  <div className="space-y-1.5">
                    <div className="flex justify-between text-xs font-bold">
                      <span className="text-cyan-600">C - Conformidade (Racionalidade / Precisão)</span>
                      <span className="text-slate-700">{selectedEvaluation.scoreC}%</span>
                    </div>
                    <div className="w-full h-2.5 bg-slate-50 rounded-full overflow-hidden border border-slate-100">
                      <div className="bg-cyan-500 h-full rounded-full transition-all duration-700" style={{ width: `${selectedEvaluation.scoreC}%` }} />
                    </div>
                  </div>
                </div>
              </div>

              {/* Written Mapped Report Section */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-10 pt-6 border-t border-slate-100">
                {/* Executive Summary & Mapped profile */}
                <div className="space-y-4">
                  <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2 border-b border-slate-50 pb-2">
                    <FileText className="h-5 w-5 text-sky-500" />
                    Diagnóstico Comportamental
                  </h3>
                  <div className="text-sm text-slate-600 leading-relaxed space-y-4 whitespace-pre-wrap">
                    {selectedEvaluation.relatorio.replace(/###/g, '')}
                  </div>
                </div>

                {/* Suggestions, Weakness and PDI Action map */}
                <div className="space-y-4">
                  <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2 border-b border-slate-50 pb-2">
                    <Sparkles className="h-5 w-5 text-sky-500" />
                    PDI e Recomendações
                  </h3>
                  <div className="text-sm text-slate-600 leading-relaxed space-y-4 whitespace-pre-wrap">
                    {selectedEvaluation.sugestoesMelhoria.replace(/###/g, '')}
                  </div>
                </div>
              </div>

              {/* Footer Signature */}
              <div className="pt-8 border-t border-slate-100 text-center text-xs text-slate-400">
                <p>Relatório gerado automaticamente pelo módulo DISC - Consultoria Empresarial Pro.</p>
                <p className="mt-1">© {new Date().getFullYear()} Sebrae / Consultor Parceiro. Todos os direitos reservados.</p>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
