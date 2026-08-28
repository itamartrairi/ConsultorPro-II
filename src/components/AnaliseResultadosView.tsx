import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  Award,
  Trash2,
  Plus,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Clock,
  Star,
  TrendingUp,
  BarChart3,
  FileText,
  Printer,
  Download,
  Save,
  RefreshCw,
  Sparkles,
  Building2,
  User,
  Phone,
  Mail,
  Calendar,
  DollarSign,
  ShieldCheck,
  ChevronRight,
  ArrowLeft,
  ArrowRight,
  ClipboardList,
  Target,
  BookOpen,
  AlertCircle,
  ThumbsUp,
  PieChart,
  HelpCircle,
  Briefcase,
  Check,
  Edit3,
  Layers,
  ArrowUpRight,
  Sliders,
  Send,
  ShieldAlert,
  Percent,
  FileSpreadsheet,
  ImageIcon,
  CheckCheck
} from 'lucide-react';
import { db, auth } from '../firebase';
import { doc, updateDoc, setDoc } from 'firebase/firestore';

export interface ProblemaAvaliado {
  id: string;
  problema: string;
  area: string;
  status: 'Resolvido' | 'Parcialmente Resolvido' | 'Em Andamento' | 'Não Resolvido';
  percentualEficacia: number; // 0 a 100
  impactoGerado: string;
  evidencias: string;
}

export interface SolucaoAvaliada {
  id: string;
  solucao: string;
  area: string;
  statusImplementacao: 'Implementada Integralmente' | 'Implementada Parcialmente' | 'Em Implantação' | 'Não Implementada';
  aderenciaRotina: 'Alta' | 'Média' | 'Baixa';
  resultadoObservado: string;
  dificuldadesEncontradas: string;
}

export interface IndicadorFinanceiroItem {
  id: string;
  nome: string;
  descricao: string;
  status: 'Ativo e Consolidado' | 'Em Implantação' | 'Parcial / Requer Atenção' | 'Não Acompanhado';
  frequencia: 'Diária' | 'Semanal' | 'Mensal' | 'Não Realiza';
  qualidadeRegistro: 'Excelente' | 'Boa' | 'Regular' | 'Incipiente';
  observacao: string;
}

export interface PesquisaSatisfacao {
  npsConsultoria: number; // 0 a 10
  satisfacaoMetodologia: number; // 1 a 5
  satisfacaoConsultorDidatica: number; // 1 a 5
  satisfacaoConsultorPontualidade: number; // 1 a 5
  satisfacaoConsultorAtendimento: number; // 1 a 5
  alcanceExpectativas: 'Superou as Expectativas' | 'Atendeu Plenamente' | 'Atendeu Parcialmente' | 'Não Atendeu';
  depoimentoCliente: string;
  pontosFortesConsultoria: string;
  sugestoesMelhoria: string;
}

export interface EngajamentoCliente {
  comprometimentoLideranca: number; // 1 a 5
  cumprimentoPrazosTarefas: number; // 1 a 5
  disponibilizacaoDadosDocumentos: number; // 1 a 5
  participacaoTreinamentos: number; // 1 a 5
  envolvimentoEquipe: number; // 1 a 5
  observacoesEngajamento: string;
}

export interface TeoriaAdminConsultoria {
  faseCicloMudanca: 'Descongelamento (Sensibilização)' | 'Movimento / Mudança (Implantação de Ferramentas)' | 'Recongelamento (Consolidação da Rotina)';
  maturidadeInicialPercent: number;
  maturidadeFinalPercent: number;
  parecerTecnicoConsultor: string;
  principaisGanhosObtidos: string;
  riscosResiduais: string;
  planoContinuidade30Dias: string;
  planoContinuidade60Dias: string;
  planoContinuidade90Dias: string;
  recomendacoesFinais: string;
}

export interface AnaliseResultadosData {
  id?: string;
  diagnosticoId: string;
  empresaId: string;
  dataAnalise: string;
  dataEncerramento?: string;
  cargaHorariaRealizada: string;
  consultorResponsavel: string;
  consultorDocumento?: string;
  clienteResponsavel: string;
  clienteCargo: string;
  clienteCpfCnpj?: string;
  tecnicoSebrae?: string;
  empresaCredenciada?: string;
  codigoSgf?: string;
  periodoConsultoria?: string;
  areaConsultoria?: string;
  objetivoConsultoria?: string;
  resultadosEsperados?: string;
  tipoRelatorio?: string;
  
  // 5 Dimensões Principais
  problemasAvaliados: ProblemaAvaliado[];
  solucoesAvaliadas: SolucaoAvaliada[];
  engajamentoCliente: EngajamentoCliente;
  indicadoresFinanceiros: IndicadorFinanceiroItem[];
  pesquisaSatisfacao: PesquisaSatisfacao;
  
  // Teoria da Administração e Práticas Avançadas
  teoriaAdmin: TeoriaAdminConsultoria;
  
  // Scores Calculados
  scoreGlobalEficacia?: number;
  updatedAt?: string;
}

const INDICADORES_FINANCEIROS_PADRAO: Omit<IndicadorFinanceiroItem, 'status' | 'frequencia' | 'qualidadeRegistro' | 'observacao'>[] = [
  {
    id: 'fluxo_caixa',
    nome: 'Controle de Fluxo de Caixa Diário',
    descricao: 'Registro rigoroso de entradas e saídas diárias com conciliação bancária e projeção de saldo futuro.'
  },
  {
    id: 'segregacao_contas',
    nome: 'Segregação de Contas PF e PJ (Pró-labore)',
    descricao: 'Separação total entre finanças pessoais dos sócios e da empresa, com pró-labore fixado.'
  },
  {
    id: 'custos_fixos_variaveis',
    nome: 'Controle e Classificação de Custos e Despesas',
    descricao: 'Estrutura detalhada de custos fixos, variáveis, diretos e indiretos apurados por centro de custo/safra.'
  },
  {
    id: 'margem_ponto_equilibrio',
    nome: 'Margem de Contribuição e Ponto de Equilíbrio',
    descricao: 'Cálculo da margem de cada produto/serviço e apuração do faturamento mínimo para cobrir custos.'
  },
  {
    id: 'formacao_preco_venda',
    nome: 'Formação Estratégica do Preço de Venda',
    descricao: 'Precificação técnica baseada no custo unitário real, impostos, comissões e margem de lucro desejada.'
  },
  {
    id: 'contas_pagar_receber',
    nome: 'Gestão de Contas a Pagar e a Receber',
    descricao: 'Agendamento prévio de pagamentos, controle de prazos médios de estocagem, pagamento e recebimento.'
  },
  {
    id: 'capital_giro_reserva',
    nome: 'Capital de Giro e Reserva de Emergência',
    descricao: 'Manutenção de reserva financeira para contingências, entressafras e suporte ao ciclo operacional.'
  },
  {
    id: 'dre_lucro_liquido',
    nome: 'DRE Gerencial e Apuração do Resultado Líquido',
    descricao: 'Fechamento mensal demonstrativo para verificação da lucratividade e rentabilidade real do negócio.'
  }
];

interface Props {
  selectedDiagnostico: any;
  selectedEmpresa: any;
  respostas?: any[];
  solucoes?: any[];
  problemas?: any[];
  tarefasPlano?: any[];
  empresasCredenciadas?: any[];
  customLogo?: string | null;
  customConsultoraLogo?: string | null;
  logoChoice?: 'sebrae' | 'consultora' | 'none' | 'ambos';
  setLogoChoice?: (c: any) => void;
  playSuccessSound: () => void;
  setView: (v: string) => void;
  setSelectedDiagnostico: (d: any) => void;
  setDiagnosticos: React.Dispatch<React.SetStateAction<any[]>>;
  empresas?: any[];
  diagnosticos?: any[];
}


// Helper to deduplicate answers by question/id
const deduplicateRespostas = (respostasList: any[]): any[] => {
  if (!respostasList || !Array.isArray(respostasList)) return [];
  const map = new Map<string, any>();

  for (const resp of respostasList) {
    if (!resp) continue;
    const normQ = (resp.pergunta || '').trim().toLowerCase();
    const key = normQ ? `q:${normQ}` : (resp.premissaId ? `id:${resp.premissaId}` : resp.id);
    if (!key) continue;

    if (!map.has(key)) {
      map.set(key, resp);
    } else {
      const existing = map.get(key)!;
      const isAnswered = (r: any) => r.resposta === 'Sim' || r.resposta === 'Parcial' || r.resposta === 'Não';
      
      const respAnswered = isAnswered(resp);
      const existingAnswered = isAnswered(existing);

      if (respAnswered && !existingAnswered) {
        map.set(key, resp);
      } else if (respAnswered && existingAnswered) {
        if (resp.observacao && !existing.observacao) {
          map.set(key, resp);
        } else if ((resp.score || 0) > (existing.score || 0)) {
          map.set(key, resp);
        }
      }
    }
  }

  return Array.from(map.values());
};

// Helper to extract clean, grouped diagnostic problems for a specific client/diagnostic
export const extractGroupedProblemasForDiag = (diagId: string, allRespostas: any[]): ProblemaAvaliado[] => {
  const diagResps = (allRespostas && allRespostas.length > 0 ? allRespostas : [])
    .filter(r => r.diagnosticoId === diagId);
  
  const cleanResps = deduplicateRespostas(diagResps);
  const negativeResps = cleanResps.filter(r => r.resposta === 'Não' || r.resposta === 'Parcial');

  // Group by unique problem description / title
  const probMap = new Map<string, { problema: string; area: string; hasNao: boolean; count: number }>();
  
  negativeResps.forEach(r => {
    const rawName = (r.problema?.trim() || r.pergunta?.trim() || 'Fragilidade de Gestão Identificada');
    if (!probMap.has(rawName)) {
      probMap.set(rawName, {
        problema: rawName,
        area: r.area || 'Gestão Financeira',
        hasNao: r.resposta === 'Não',
        count: 1
      });
    } else {
      const entry = probMap.get(rawName)!;
      entry.count++;
      if (r.resposta === 'Não') entry.hasNao = true;
    }
  });

  if (probMap.size > 0) {
    return Array.from(probMap.values()).map((p, idx) => ({
      id: `prob_${idx}_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      problema: p.problema,
      area: p.area,
      status: 'Resolvido' as const,
      percentualEficacia: p.hasNao ? 90 : 95,
      impactoGerado: `Processo e controle estruturados durante a consultoria para sanar a fragilidade diagnosticada em "${p.problema}".`,
      evidencias: 'Planilha / rotina gerencial implementada, acompanhada e validada nas sessões de consultoria.'
    }));
  }

  // Fallback defaults if no negative answers
  return [
    {
      id: `prob_def_0`,
      problema: 'Ausência de controle diário e rigoroso de entradas e saídas de caixa',
      area: 'Gestão Financeira',
      status: 'Resolvido',
      percentualEficacia: 92,
      impactoGerado: 'Rotina de fluxo de caixa diário implantada e em operação no dia a dia com conciliação.',
      evidencias: 'Planilha gerencial de caixa e conciliação bancária semanal conferida nas sessões.'
    },
    {
      id: `prob_def_1`,
      problema: 'Mistura de despesas pessoais dos sócios com as da empresa (falta de pró-labore fixo)',
      area: 'Controle e Rotina',
      status: 'Resolvido',
      percentualEficacia: 95,
      impactoGerado: 'Separação total de contas PF e PJ e pró-labore estipulado formalmente.',
      evidencias: 'Contas bancárias segregadas e retiradas limitadas ao pró-labore estipulado.'
    },
    {
      id: `prob_def_2`,
      problema: 'Desconhecimento dos custos operacionais reais e da margem de contribuição',
      area: 'Custos e Resultado',
      status: 'Resolvido',
      percentualEficacia: 90,
      impactoGerado: 'Estrutura detalhada de custos fixos, variáveis e apuração da margem de lucro.',
      evidencias: 'Planilha de custos e ponto de equilíbrio calculados e apresentados.'
    },
    {
      id: `prob_def_3`,
      problema: 'Precificação baseada apenas no mercado sem apuração dos custos unitários',
      area: 'Preço e Mercado',
      status: 'Resolvido',
      percentualEficacia: 88,
      impactoGerado: 'Tabela de preços técnicos e margem mínima estabelecida.',
      evidencias: 'Precificação revista conforme estrutura de custos operacionais apurada.'
    }
  ];
};

export const extractCronogramaSolucoesForDiag = (diag: any): SolucaoAvaliada[] => {
  const cronograma = (diag?.cronograma || []).filter((atv: any) => atv.nome && atv.nome.trim() !== '');
  if (cronograma.length > 0) {
    return cronograma.map((atv: any, idx: number) => ({
      id: `sol_${idx}_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      solucao: atv.solucaoProposta || atv.nome || `Atividade de Consultoria ${idx + 1}`,
      area: atv.area || 'Gestão Financeira',
      statusImplementacao: (atv.status === 'Concluído' ? 'Implementada Integralmente' : 'Implementada Parcialmente') as any,
      aderenciaRotina: 'Alta' as const,
      resultadoObservado: atv.descricao ? `Etapa executada: ${atv.descricao}` : 'Atividade desenvolvida com sucesso e integrada à rotina.',
      dificuldadesEncontradas: 'Nenhuma que comprometesse os resultados esperados.'
    }));
  }

  return [
    {
      id: `sol_def_0`,
      solucao: 'Estruturação da Planilha de Fluxo de Caixa e Treinamento do Lançamento Diário',
      area: 'Gestão Financeira',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Ação 100% orientada e aplicada pelo cliente durante a consultoria.',
      dificuldadesEncontradas: 'Ajuste inicial de hábitos e rotina de anotação de dados.'
    },
    {
      id: `sol_def_1`,
      solucao: 'Definição do Pró-labore e Segregação Rigorosa das Contas PF/PJ',
      area: 'Controle e Rotina',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Separação física e bancária implementada com sucesso.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: `sol_def_2`,
      solucao: 'Mapeamento e Classificação Detalhada dos Custos Fixos e Variáveis',
      area: 'Custos e Resultado',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Classificação completa de custos realizada e assimilada.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: `sol_def_3`,
      solucao: 'Formação Técnica da Tabela de Preço de Venda e Margem de Contribuição',
      area: 'Preço e Mercado',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Margens calculadas e preços ajustados.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: `sol_def_4`,
      solucao: 'Implantação do Fechamento Mensal de Resultados (DRE Gerencial)',
      area: 'Gestão Financeira',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Demonstrativo estruturado para acompanhamento da lucratividade.',
      dificuldadesEncontradas: 'Nenhuma.'
    }
  ];
};

export const AnaliseResultadosView: React.FC<Props> = ({
  selectedDiagnostico,
  selectedEmpresa,
  respostas = [],
  solucoes = [],
  problemas = [],
  tarefasPlano = [],
  empresasCredenciadas = [],
  customLogo,
  customConsultoraLogo,
  logoChoice = 'sebrae',
  setLogoChoice,
  playSuccessSound,
  setView,
  setSelectedDiagnostico,
  setDiagnosticos,
  empresas = [],
  diagnosticos = []
}) => {
  const [activeTab, setActiveTab] = useState<'painel' | 'problemas' | 'solucoes' | 'engajamento' | 'financeiro' | 'satisfacao' | 'teoria' | 'relatorio'>('painel');
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [isAutoFilling, setIsAutoFilling] = useState(false);
  const [localLogoChoice, setLocalLogoChoice] = useState<'sebrae' | 'consultora' | 'ambos' | 'none'>(
    (logoChoice as any) || 'sebrae'
  );
  const reportRef = useRef<HTMLDivElement>(null);

  // Effective Logos resolved from Props or localStorage
  const effectiveSebraeLogo = customLogo || localStorage.getItem('sebrae_custom_logo');
  const effectiveConsultoraLogo = customConsultoraLogo || localStorage.getItem('consultora_custom_logo');

  // Helper to extract credenciada name
  const credenciadaName = useMemo(() => {
    if (selectedDiagnostico?.empresaCredenciadaId) {
      const match = empresasCredenciadas.find(c => c.id === selectedDiagnostico.empresaCredenciadaId);
      if (match) return match.razaoSocial || match.nome;
    }
    if (empresasCredenciadas.length > 0) {
      return empresasCredenciadas[0].razaoSocial || empresasCredenciadas[0].nome;
    }
    return '';
  }, [selectedDiagnostico?.empresaCredenciadaId, empresasCredenciadas]);

  // Initial Form Data State with automatic contract binding
  const [data, setData] = useState<AnaliseResultadosData>(() => {
    const savedInDiag = selectedDiagnostico?.analiseResultados;
    if (savedInDiag && typeof savedInDiag === 'object') {
      return {
        ...savedInDiag,
        // Ensure contract data is updated if present in dadosConsultoria
        codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || savedInDiag.codigoSgf || '',
        periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || savedInDiag.periodoConsultoria || '',
        cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || savedInDiag.cargaHorariaRealizada || selectedDiagnostico?.cargaHoraria || '30 Horas',
        consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || savedInDiag.consultorResponsavel || '',
        tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || savedInDiag.tecnicoSebrae || '',
        areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || savedInDiag.areaConsultoria || 'Gestão Financeira e Produção',
        objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || savedInDiag.objetivoConsultoria || 'Implementar melhorias estratégicas e controles financeiros e produtivos.',
        resultadosEsperados: selectedDiagnostico?.dadosConsultoria?.resultadosEsperados || savedInDiag.resultadosEsperados || '',
        tipoRelatorio: selectedDiagnostico?.dadosConsultoria?.tipoRelatorio || savedInDiag.tipoRelatorio || 'Final',
        clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || savedInDiag.clienteResponsavel || '',
        clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || savedInDiag.clienteCpfCnpj || '',
        empresaCredenciada: savedInDiag.empresaCredenciada || credenciadaName
      };
    }
    
    try {
      const localKey = `analise_resultados_${selectedDiagnostico?.id}`;
      const savedLocal = localStorage.getItem(localKey);
      if (savedLocal) {
        return JSON.parse(savedLocal);
      }
    } catch (e) {
      console.warn("Could not load local analise_resultados", e);
    }

    // Default scaffold with complete contract data bindings
    return {
      diagnosticoId: selectedDiagnostico?.id || '',
      empresaId: selectedEmpresa?.id || selectedDiagnostico?.empresaId || '',
      dataAnalise: new Date().toISOString().split('T')[0],
      dataEncerramento: new Date().toISOString().split('T')[0],
      cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || selectedDiagnostico?.cargaHoraria || '30 Horas',
      consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || '',
      consultorDocumento: '',
      clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || '',
      clienteCargo: 'Proprietário / Gestor',
      clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || '',
      tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || '',
      empresaCredenciada: credenciadaName || '',
      codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || '',
      periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || '',
      areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || 'Gestão Financeira e Produção',
      objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || 'Implementar melhorias estratégicas e práticas nas áreas de Gestão Financeira, Preço e Mercado, Custos e Resultados e Rotinas Operacionais.',
      resultadosEsperados: selectedDiagnostico?.dadosConsultoria?.resultadosEsperados || 'Aumento do controle financeiro, separação de despesas pessoais/empresariais e apuração da margem real de lucro.',
      tipoRelatorio: selectedDiagnostico?.dadosConsultoria?.tipoRelatorio || 'Final',
      
      problemasAvaliados: [],
      solucoesAvaliadas: [],
      engajamentoCliente: {
        comprometimentoLideranca: 5,
        cumprimentoPrazosTarefas: 4,
        disponibilizacaoDadosDocumentos: 5,
        participacaoTreinamentos: 5,
        envolvimentoEquipe: 4,
        observacoesEngajamento: 'Cliente demonstrou alto interesse e compromisso com a transformação gerencial do negócio, comparecendo a todas as sessões agendadas e aplicando as orientações com presteza.'
      },
      indicadoresFinanceiros: INDICADORES_FINANCEIROS_PADRAO.map(ind => ({
        ...ind,
        status: ind.id === 'fluxo_caixa' || ind.id === 'segregacao_contas' ? 'Ativo e Consolidado' : 'Em Implantação',
        frequencia: ind.id === 'fluxo_caixa' ? 'Diária' : 'Mensal',
        qualidadeRegistro: 'Boa',
        observacao: 'Ferramenta estruturada durante a consultoria e assimilada pelo cliente.'
      })),
      pesquisaSatisfacao: {
        npsConsultoria: 10,
        satisfacaoMetodologia: 5,
        satisfacaoConsultorDidatica: 5,
        satisfacaoConsultorPontualidade: 5,
        satisfacaoConsultorAtendimento: 5,
        alcanceExpectativas: 'Superou as Expectativas',
        depoimentoCliente: 'A consultoria foi um divisor de águas na organização das contas e clareza dos números da nossa empresa.',
        pontosFortesConsultoria: 'Praticidade nas ferramentas, paciência e didática do consultor na explicação e foco na realidade do nosso segmento.',
        sugestoesMelhoria: 'Manter acompanhamento de pós-consultoria semestral para reforçar os novos controles.'
      },
      teoriaAdmin: {
        faseCicloMudanca: 'Recongelamento (Consolidação da Rotina)',
        maturidadeInicialPercent: selectedDiagnostico?.percentualGeral || 35,
        maturidadeFinalPercent: 88,
        parecerTecnicoConsultor: 'O cliente concluiu com pleno êxito o programa de consultoria gerencial. Foram sanadas as principais fragilidades financeiras, estabelecida a disciplina de caixa e segregadas as contas pessoais dos custos operacionais.',
        principaisGanhosObtidos: '1. Clareza total da margem real de lucro.\n2. Eliminação de despesas ocultas e juros bancários desnecessários.\n3. Domínio da ferramenta gerencial de caixa e custos.\n4. Autonomia na precificação técnica.',
        riscosResiduais: 'Risco de afrouxamento na disciplina de alimentação diária do fluxo de caixa em períodos de alta intensidade de trabalho no campo.',
        planoContinuidade30Dias: 'Realizar o primeiro fechamento mensal de resultado (DRE) de forma 100% autônoma e conferir a conciliação bancária semanalmente.',
        planoContinuidade60Dias: 'Apurar a margem de contribuição acumulada da safra/período e destinar 20% do lucro líquido apurado para o fundo de reserva.',
        planoContinuidade90Dias: 'Revisar a tabela de preços e custos unitários antes do início do novo ciclo produtivo.',
        recomendacoesFinais: 'Manter a reunião financeira quinzenal de fechamento e não realizar retiradas além do pró-labore estipulado.'
      }
    };
  });

  // Re-sync and isolate data strictly when selected diagnostico changes
  useEffect(() => {
    if (!selectedDiagnostico?.id) return;

    // 1. If saved directly inside the diagnostico object
    if (selectedDiagnostico.analiseResultados && selectedDiagnostico.analiseResultados.problemasAvaliados) {
      setData({
        ...selectedDiagnostico.analiseResultados,
        diagnosticoId: selectedDiagnostico.id,
        empresaId: selectedEmpresa?.id || selectedDiagnostico.empresaId || '',
        codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || selectedDiagnostico.analiseResultados.codigoSgf || '',
        periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || selectedDiagnostico.analiseResultados.periodoConsultoria || '',
        cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || selectedDiagnostico.analiseResultados.cargaHorariaRealizada || selectedDiagnostico?.cargaHoraria || '30 Horas',
        consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || selectedDiagnostico.analiseResultados.consultorResponsavel || '',
        tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || selectedDiagnostico.analiseResultados.tecnicoSebrae || '',
        clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || selectedDiagnostico.analiseResultados.clienteResponsavel || '',
        clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || selectedDiagnostico.analiseResultados.clienteCpfCnpj || '',
        empresaCredenciada: selectedDiagnostico.analiseResultados.empresaCredenciada || credenciadaName || '',
        areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || selectedDiagnostico.analiseResultados.areaConsultoria || 'Gestão Financeira e Produção',
        objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || selectedDiagnostico.analiseResultados.objetivoConsultoria || '',
        resultadosEsperados: selectedDiagnostico?.dadosConsultoria?.resultadosEsperados || selectedDiagnostico.analiseResultados.resultadosEsperados || '',
        tipoRelatorio: selectedDiagnostico?.dadosConsultoria?.tipoRelatorio || selectedDiagnostico.analiseResultados.tipoRelatorio || 'Final'
      });
      return;
    }

    // 2. Check localStorage for this specific diagnostico
    try {
      const localKey = `analise_resultados_${selectedDiagnostico.id}`;
      const savedLocal = localStorage.getItem(localKey);
      if (savedLocal) {
        const parsed = JSON.parse(savedLocal);
        if (parsed && (parsed.diagnosticoId === selectedDiagnostico.id || !parsed.diagnosticoId)) {
          setData({
            ...parsed,
            diagnosticoId: selectedDiagnostico.id,
            clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || parsed.clienteResponsavel,
            clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || parsed.clienteCpfCnpj,
            codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || parsed.codigoSgf,
            consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || parsed.consultorResponsavel,
            tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || parsed.tecnicoSebrae,
            empresaCredenciada: credenciadaName || parsed.empresaCredenciada
          });
          return;
        }
      }
    } catch (e) {
      console.warn("Could not load local analise_resultados for diag", e);
    }

    // 3. Otherwise generate fresh data strictly scoped to this client
    const freshProbs = extractGroupedProblemasForDiag(selectedDiagnostico.id, respostas);
    const freshSols = extractCronogramaSolucoesForDiag(selectedDiagnostico);

    setData({
      diagnosticoId: selectedDiagnostico.id,
      empresaId: selectedEmpresa?.id || selectedDiagnostico.empresaId || '',
      dataAnalise: new Date().toISOString().split('T')[0],
      dataEncerramento: new Date().toISOString().split('T')[0],
      cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || selectedDiagnostico?.cargaHoraria || '30 Horas',
      consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || '',
      consultorDocumento: '',
      clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || '',
      clienteCargo: 'Proprietário / Gestor',
      clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || '',
      tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || '',
      empresaCredenciada: credenciadaName || '',
      codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || '',
      periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || '',
      areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || 'Gestão Financeira e Produção',
      objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || 'Implementar melhorias estratégicas e práticas nas áreas de Gestão Financeira, Preço e Mercado, Custos e Resultados e Rotinas Operacionais.',
      resultadosEsperados: selectedDiagnostico?.dadosConsultoria?.resultadosEsperados || 'Aumento do controle financeiro, separação de despesas pessoais/empresariais e apuração da margem real de lucro.',
      tipoRelatorio: selectedDiagnostico?.dadosConsultoria?.tipoRelatorio || 'Final',
      problemasAvaliados: freshProbs,
      solucoesAvaliadas: freshSols,
      engajamentoCliente: {
        comprometimentoLideranca: 5,
        cumprimentoPrazosTarefas: 4,
        disponibilizacaoDadosDocumentos: 5,
        participacaoTreinamentos: 5,
        envolvimentoEquipe: 4,
        observacoesEngajamento: 'Cliente demonstrou alto interesse e compromisso com a transformação gerencial do negócio, comparecendo a todas as sessões agendadas e aplicando as orientações com presteza.'
      },
      indicadoresFinanceiros: INDICADORES_FINANCEIROS_PADRAO.map(ind => ({
        ...ind,
        status: ind.id === 'fluxo_caixa' || ind.id === 'segregacao_contas' ? 'Ativo e Consolidado' : 'Em Implantação',
        frequencia: ind.id === 'fluxo_caixa' ? 'Diária' : 'Mensal',
        qualidadeRegistro: 'Boa',
        observacao: 'Ferramenta estruturada durante a consultoria e assimilada pelo cliente.'
      })),
      pesquisaSatisfacao: {
        npsConsultoria: 10,
        satisfacaoMetodologia: 5,
        satisfacaoConsultorDidatica: 5,
        satisfacaoConsultorPontualidade: 5,
        satisfacaoConsultorAtendimento: 5,
        alcanceExpectativas: 'Superou as Expectativas',
        depoimentoCliente: 'A consultoria foi um divisor de águas na organização das contas e clareza dos números da nossa empresa.',
        pontosFortesConsultoria: 'Praticidade nas ferramentas, paciência e didática do consultor na explicação e foco na realidade do nosso segmento.',
        sugestoesMelhoria: 'Manter acompanhamento de pós-consultoria semestral para reforçar os novos controles.'
      },
      teoriaAdmin: {
        faseCicloMudanca: 'Recongelamento (Consolidação da Rotina)',
        maturidadeInicialPercent: selectedDiagnostico?.percentualGeral || 35,
        maturidadeFinalPercent: 88,
        parecerTecnicoConsultor: 'O cliente concluiu com pleno êxito o programa de consultoria gerencial. Foram sanadas as principais fragilidades financeiras, estabelecida a disciplina de caixa e segregadas as contas pessoais dos custos operacionais.',
        principaisGanhosObtidos: '1. Clareza total da margem real de lucro.\n2. Eliminação de despesas ocultas e juros bancários desnecessários.\n3. Domínio da ferramenta gerencial de caixa e custos.\n4. Autonomia na precificação técnica.',
        riscosResiduais: 'Risco de afrouxamento na disciplina de alimentação diária do fluxo de caixa em períodos de alta intensidade de trabalho no campo.',
        planoContinuidade30Dias: 'Realizar o primeiro fechamento mensal de resultado (DRE) de forma 100% autônoma e conferir a conciliação bancária semanalmente.',
        planoContinuidade60Dias: 'Apurar a margem de contribuição acumulada da safra/período e destinar 20% do lucro líquido apurado para o fundo de reserva.',
        planoContinuidade90Dias: 'Revisar a tabela de preços e custos unitários antes do início do novo ciclo produtivo.',
        recomendacoesFinais: 'Manter a reunião financeira quinzenal de fechamento e não realizar retiradas além do pró-labore estipulado.'
      }
    });
  }, [selectedDiagnostico?.id, selectedDiagnostico?.dadosConsultoria, credenciadaName]);

  // Sync with Dados da Consultoria explicitly
  const handleSyncContractData = () => {
    if (!selectedDiagnostico) return;
    const dc = selectedDiagnostico.dadosConsultoria || {};
    setData(prev => ({
      ...prev,
      codigoSgf: dc.codigoSgf || prev.codigoSgf,
      periodoConsultoria: dc.periodoConsultoria || prev.periodoConsultoria,
      cargaHorariaRealizada: dc.cargaHoraria || selectedDiagnostico.cargaHoraria || prev.cargaHorariaRealizada,
      consultorResponsavel: dc.consultor || prev.consultorResponsavel,
      tecnicoSebrae: dc.tecnicoSebrae || prev.tecnicoSebrae,
      areaConsultoria: dc.areaConsultoria || prev.areaConsultoria,
      objetivoConsultoria: dc.objetivo || prev.objetivoConsultoria,
      resultadosEsperados: dc.resultadosEsperados || prev.resultadosEsperados,
      tipoRelatorio: dc.tipoRelatorio || prev.tipoRelatorio,
      clienteResponsavel: dc.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico.nomeEmpresa || prev.clienteResponsavel,
      clienteCpfCnpj: dc.cnpj || selectedEmpresa?.cnpj || prev.clienteCpfCnpj,
      empresaCredenciada: credenciadaName || prev.empresaCredenciada
    }));
    playSuccessSound();
    alert('Dados do contrato e da consultoria sincronizados com sucesso!');
  };

  // Extract Diagnostic Problems & Solutions Automatically if empty
  useEffect(() => {
    if (data.problemasAvaliados.length === 0 && selectedDiagnostico?.id) {
      autoPopulateFromDiagnostico();
    }
  }, [selectedDiagnostico?.id, respostas, selectedDiagnostico?.cronograma]);

  const autoPopulateFromDiagnostico = () => {
    if (!selectedDiagnostico?.id) return;
    setIsAutoFilling(true);
    
    // 1. Identify and cleanly group problems from this client's diagnostic answers
    const generatedProblemas = extractGroupedProblemasForDiag(selectedDiagnostico.id, respostas);

    // 2. Identify solutions from cronograma activities
    const generatedSolucoes = extractCronogramaSolucoesForDiag(selectedDiagnostico);

    setData(prev => ({
      ...prev,
      problemasAvaliados: generatedProblemas,
      solucoesAvaliadas: generatedSolucoes
    }));

    playSuccessSound();
    setTimeout(() => setIsAutoFilling(false), 300);
  };

  const handleAddProblema = () => {
    const novo: ProblemaAvaliado = {
      id: `prob_custom_${Date.now()}`,
      problema: 'Novo problema identificado',
      area: data.areaConsultoria || 'Gestão Financeira',
      status: 'Resolvido',
      percentualEficacia: 100,
      impactoGerado: 'Melhoria prática estruturada durante a consultoria.',
      evidencias: 'Rotina e controles validados.'
    };
    setData(prev => ({
      ...prev,
      problemasAvaliados: [...prev.problemasAvaliados, novo]
    }));
  };

  const handleDeleteProblema = (indexToDelete: number) => {
    setData(prev => ({
      ...prev,
      problemasAvaliados: prev.problemasAvaliados.filter((_, idx) => idx !== indexToDelete)
    }));
  };

  const handleAddSolucao = () => {
    const nova: SolucaoAvaliada = {
      id: `sol_custom_${Date.now()}`,
      solucao: 'Nova atividade / ferramenta implantada',
      area: data.areaConsultoria || 'Gestão Financeira',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Ação executada com sucesso e integrada à rotina do cliente.',
      dificuldadesEncontradas: 'Nenhuma.'
    };
    setData(prev => ({
      ...prev,
      solucoesAvaliadas: [...prev.solucoesAvaliadas, nova]
    }));
  };

  const handleDeleteSolucao = (indexToDelete: number) => {
    setData(prev => ({
      ...prev,
      solucoesAvaliadas: prev.solucoesAvaliadas.filter((_, idx) => idx !== indexToDelete)
    }));
  };

  // Mathematical Calculations for Global Score & Efficiency
  const calculatedMetrics = useMemo(() => {
    // 1. Problemas Resolvidos Score (Weight 30%)
    const totalProb = data.problemasAvaliados.length || 1;
    const sumProbEficacia = data.problemasAvaliados.reduce((acc, p) => acc + (p.percentualEficacia || 0), 0);
    const probRate = Math.round(sumProbEficacia / totalProb);

    // 2. Soluções Implementadas Score (Weight 25%)
    const totalSol = data.solucoesAvaliadas.length || 1;
    const solIntegralCount = data.solucoesAvaliadas.filter(s => s.statusImplementacao === 'Implementada Integralmente').length;
    const solParcialCount = data.solucoesAvaliadas.filter(s => s.statusImplementacao === 'Implementada Parcialmente').length;
    const solEmImpCount = data.solucoesAvaliadas.filter(s => s.statusImplementacao === 'Em Implantação').length;
    const solPoints = (solIntegralCount * 100) + (solParcialCount * 65) + (solEmImpCount * 40);
    const solRate = Math.round(solPoints / totalSol);

    // 3. Indicadores Financeiros Score (Weight 20%)
    const totalInd = data.indicadoresFinanceiros.length || 1;
    const indConsolidadosCount = data.indicadoresFinanceiros.filter(i => i.status === 'Ativo e Consolidado').length;
    const indEmImpCount = data.indicadoresFinanceiros.filter(i => i.status === 'Em Implantação').length;
    const indParcialCount = data.indicadoresFinanceiros.filter(i => i.status === 'Parcial / Requer Atenção').length;
    const indPoints = (indConsolidadosCount * 100) + (indEmImpCount * 70) + (indParcialCount * 40);
    const indRate = Math.round(indPoints / totalInd);

    // 4. Engajamento do Cliente Score (Weight 15%)
    const eng = data.engajamentoCliente;
    const engSum = (eng.comprometimentoLideranca + eng.cumprimentoPrazosTarefas + eng.disponibilizacaoDadosDocumentos + eng.participacaoTreinamentos + eng.envolvimentoEquipe);
    const engRate = Math.round((engSum / 25) * 100);

    // 5. Pesquisa de Satisfação NPS (Weight 10%)
    const npsRate = Math.min(100, Math.round((data.pesquisaSatisfacao.npsConsultoria / 10) * 100));

    // Global Weighted Score
    const globalScore = Math.round(
      (probRate * 0.30) +
      (solRate * 0.25) +
      (indRate * 0.20) +
      (engRate * 0.15) +
      (npsRate * 0.10)
    );

    // Classification Category
    let nivelClassificacao = 'Transformação Excelente (Alta Eficácia)';
    let nivelCor = 'text-emerald-600 bg-emerald-50 border-emerald-200';
    if (globalScore < 50) {
      nivelClassificacao = 'Eficácia Baixa (Requer Reforço Urgente)';
      nivelCor = 'text-rose-600 bg-rose-50 border-rose-200';
    } else if (globalScore < 75) {
      nivelClassificacao = 'Eficácia Média / Parcial (Em Evolução)';
      nivelCor = 'text-amber-600 bg-amber-50 border-amber-200';
    } else if (globalScore < 90) {
      nivelClassificacao = 'Eficácia Satisfatória / Muito Boa';
      nivelCor = 'text-blue-600 bg-blue-50 border-blue-200';
    }

    return {
      totalProb,
      probRate,
      resolvidosCount: data.problemasAvaliados.filter(p => p.status === 'Resolvido').length,
      totalSol,
      solRate,
      solIntegralCount,
      totalInd,
      indRate,
      indConsolidadosCount,
      engRate,
      npsRate,
      globalScore,
      nivelClassificacao,
      nivelCor
    };
  }, [data]);

  // Save to Firestore and Local Storage
  const handleSave = async () => {
    setIsSaving(true);
    try {
      const finalData: AnaliseResultadosData = {
        ...data,
        scoreGlobalEficacia: calculatedMetrics.globalScore,
        updatedAt: new Date().toISOString()
      };

      // 1. Save in localStorage for immediate resilience
      const localKey = `analise_resultados_${selectedDiagnostico?.id}`;
      localStorage.setItem(localKey, JSON.stringify(finalData));

      // 2. Save in Cloud Firestore
      const user = auth.currentUser;
      if (user && selectedDiagnostico?.id) {
        // Update document in diagnosticos collection
        await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), {
          analiseResultados: finalData,
          updatedAt: new Date().toISOString()
        });

        // Also save in dedicated collection
        await setDoc(doc(db, 'analise_resultados', selectedDiagnostico.id), {
          ...finalData,
          ownerId: user.uid
        }, { merge: true });
      }

      // Update parent states
      const updatedDiag = {
        ...selectedDiagnostico,
        analiseResultados: finalData
      };
      setSelectedDiagnostico(updatedDiag);
      setDiagnosticos(prev => prev.map(d => d.id === selectedDiagnostico.id ? updatedDiag : d));

      setIsSaving(false);
      setSaveSuccess(true);
      playSuccessSound();
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (error) {
      console.error("Error saving analise_resultados:", error);
      setIsSaving(false);
      alert("Erro ao salvar os dados na nuvem. Os dados foram mantidos localmente.");
    }
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-16">
      {/* Top Header Card */}
      <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-100 shadow-sm print:hidden">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-14 h-14 bg-gradient-to-br from-emerald-500 to-teal-700 text-white rounded-2xl flex items-center justify-center shadow-lg shadow-emerald-200">
              <Award size={30} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                  Avaliação Pós-Consultoria & Relatório Final
                </span>
                {data.codigoSgf && (
                  <span className="px-2.5 py-0.5 rounded-full text-xs font-mono font-bold bg-slate-100 text-slate-700 border border-slate-200">
                    SGF: {data.codigoSgf}
                  </span>
                )}
              </div>
              <h1 className="text-2xl md:text-3xl font-black text-slate-800 tracking-tight mt-1">
                Análise de Resultados da Consultoria
              </h1>
              <p className="text-slate-500 text-sm font-medium">
                {selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || 'Empresa Cliente'} • {data.cargaHorariaRealizada} • {data.areaConsultoria}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            <button
              onClick={handleSyncContractData}
              title="Puxar dados atualizados do cadastro do contrato"
              className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition-all shadow-xs cursor-pointer"
            >
              <RefreshCw size={15} className="text-emerald-600" /> Sincronizar Contrato
            </button>

            <button
              onClick={() => setActiveTab('relatorio')}
              className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-bold transition-all shadow-xs cursor-pointer"
            >
              <Printer size={16} className="text-slate-600" /> Relatório Formal (PDF)
            </button>

            <button
              onClick={handleSave}
              disabled={isSaving}
              className={`inline-flex items-center gap-2 px-6 py-2.5 rounded-xl text-white text-xs font-bold transition-all shadow-sm cursor-pointer ${
                saveSuccess ? 'bg-emerald-700' : 'bg-emerald-600 hover:bg-emerald-700'
              }`}
            >
              {isSaving ? (
                <>
                  <RefreshCw size={16} className="animate-spin" /> Salvando...
                </>
              ) : saveSuccess ? (
                <>
                  <CheckCheck size={16} /> Salvo com Sucesso!
                </>
              ) : (
                <>
                  <Save size={16} /> Salvar Análise
                </>
              )}
            </button>
          </div>
        </div>

        {/* Global Scorecard Banner */}
        <div className="mt-8 pt-6 border-t border-slate-100 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Score de Eficácia</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-emerald-600">{calculatedMetrics.globalScore}%</span>
              <span className="text-xs font-bold text-slate-400">Geral</span>
            </div>
            <span className="text-[10px] font-semibold text-emerald-700 mt-1 truncate">{calculatedMetrics.nivelClassificacao.split(' ')[0]}</span>
          </div>

          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Gargalos Sanados</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-slate-800">{calculatedMetrics.probRate}%</span>
              <span className="text-xs font-bold text-slate-400">Taxa</span>
            </div>
            <span className="text-[10px] font-semibold text-slate-500 mt-1">{calculatedMetrics.resolvidosCount} de {calculatedMetrics.totalProb} Resolvidos</span>
          </div>

          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Ações Implantadas</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-blue-600">{calculatedMetrics.solRate}%</span>
              <span className="text-xs font-bold text-slate-400">Aderência</span>
            </div>
            <span className="text-[10px] font-semibold text-blue-700 mt-1">{calculatedMetrics.solIntegralCount} de {calculatedMetrics.totalSol} Integrais</span>
          </div>

          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Rotinas Financeiras</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-amber-600">{calculatedMetrics.indRate}%</span>
              <span className="text-xs font-bold text-slate-400">Ativas</span>
            </div>
            <span className="text-[10px] font-semibold text-amber-700 mt-1">{calculatedMetrics.indConsolidadosCount} de {calculatedMetrics.totalInd} Controles</span>
          </div>

          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Engajamento Cliente</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-indigo-600">{calculatedMetrics.engRate}%</span>
              <span className="text-xs font-bold text-slate-400">Nota</span>
            </div>
            <span className="text-[10px] font-semibold text-indigo-700 mt-1">Liderança & Prazos</span>
          </div>

          <div className="p-4 rounded-2xl bg-slate-50 border border-slate-100 flex flex-col justify-between">
            <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Satisfação (NPS)</span>
            <div className="flex items-baseline gap-1 mt-2">
              <span className="text-3xl font-black text-purple-600">{data.pesquisaSatisfacao.npsConsultoria}</span>
              <span className="text-xs font-bold text-slate-400">/ 10</span>
            </div>
            <span className="text-[10px] font-semibold text-purple-700 mt-1">{data.pesquisaSatisfacao.npsConsultoria >= 9 ? 'Promotor Ativo' : 'Neutro'}</span>
          </div>
        </div>

        {/* Navigation Tabs */}
        <div className="mt-8 flex items-center gap-1.5 overflow-x-auto pb-1 border-b border-slate-100 text-xs font-bold">
          <button
            onClick={() => setActiveTab('painel')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'painel'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            📊 Painel Executivo & Contrato
          </button>
          <button
            onClick={() => setActiveTab('problemas')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'problemas'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            🎯 1. Problemas Diagnosticados ({data.problemasAvaliados.length})
          </button>
          <button
            onClick={() => setActiveTab('solucoes')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'solucoes'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            💼 2. Implementação de Soluções ({data.solucoesAvaliadas.length})
          </button>
          <button
            onClick={() => setActiveTab('financeiro')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'financeiro'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            💵 3. Índices & Controles Financeiros
          </button>
          <button
            onClick={() => setActiveTab('engajamento')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'engajamento'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            🤝 4. Engajamento do Cliente
          </button>
          <button
            onClick={() => setActiveTab('satisfacao')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'satisfacao'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            ⭐ 5. Pesquisa & NPS
          </button>
          <button
            onClick={() => setActiveTab('teoria')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'teoria'
                ? 'bg-emerald-600 text-white shadow-xs'
                : 'text-slate-600 hover:bg-slate-50 hover:text-slate-900'
            }`}
          >
            📚 6. Parecer Técnico & Continuidade
          </button>
          <button
            onClick={() => setActiveTab('relatorio')}
            className={`px-4 py-2.5 rounded-xl transition-all whitespace-nowrap cursor-pointer ${
              activeTab === 'relatorio'
                ? 'bg-slate-900 text-white shadow-xs'
                : 'text-slate-700 bg-slate-100 hover:bg-slate-200'
            }`}
          >
            📄 7. Relatório Formal com Assinaturas
          </button>
        </div>
      </div>

      {/* TAB 1: PAINEL EXECUTIVO & CONTRATO */}
      {activeTab === 'painel' && (
        <motion.div
          key="painel"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="space-y-6"
        >
          {/* Contract Metadata Connection Card */}
          <div className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div>
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <FileSpreadsheet className="text-emerald-600" size={20} />
                  Dados Contratuais Integrados da Consultoria
                </h3>
                <p className="text-slate-500 text-xs mt-0.5">
                  Estas informações são sincronizadas automaticamente com os <strong>Dados da Consultoria</strong> e alimentam os cabeçalhos e blocos de assinatura.
                </p>
              </div>
              <button
                onClick={handleSyncContractData}
                className="px-3.5 py-1.5 rounded-xl bg-emerald-50 text-emerald-700 hover:bg-emerald-100 text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <RefreshCw size={14} /> Re-sincronizar
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Razão Social / Cliente</label>
                <input
                  type="text"
                  value={data.clienteResponsavel}
                  onChange={(e) => setData({ ...data, clienteResponsavel: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Nome do cliente ou empresa"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">CNPJ / CPF do Cliente</label>
                <input
                  type="text"
                  value={data.clienteCpfCnpj || ''}
                  onChange={(e) => setData({ ...data, clienteCpfCnpj: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="00.000.000/0000-00"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Código SGF / Contrato</label>
                <input
                  type="text"
                  value={data.codigoSgf || ''}
                  onChange={(e) => setData({ ...data, codigoSgf: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Ex: SGF-2026-8842"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Carga Horária</label>
                <input
                  type="text"
                  value={data.cargaHorariaRealizada}
                  onChange={(e) => setData({ ...data, cargaHorariaRealizada: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="30 Horas"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Consultor Técnico Responsável</label>
                <input
                  type="text"
                  value={data.consultorResponsavel}
                  onChange={(e) => setData({ ...data, consultorResponsavel: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Nome do consultor"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Gestor / Técnico Sebrae</label>
                <input
                  type="text"
                  value={data.tecnicoSebrae || ''}
                  onChange={(e) => setData({ ...data, tecnicoSebrae: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Nome do gestor do Sebrae"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Empresa Credenciada</label>
                <input
                  type="text"
                  value={data.empresaCredenciada || ''}
                  onChange={(e) => setData({ ...data, empresaCredenciada: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Razão Social da Credenciada"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Data de Emissão / Encerramento</label>
                <input
                  type="date"
                  value={data.dataAnalise}
                  onChange={(e) => setData({ ...data, dataAnalise: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-slate-800 focus:bg-white focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs pt-2">
              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Objetivo Geral da Consultoria</label>
                <textarea
                  rows={2}
                  value={data.objetivoConsultoria || ''}
                  onChange={(e) => setData({ ...data, objetivoConsultoria: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-medium text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Descreva o objetivo contratual e escopo"
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-500 uppercase tracking-wider text-[10px]">Resultados Esperados pelo Contrato</label>
                <textarea
                  rows={2}
                  value={data.resultadosEsperados || ''}
                  onChange={(e) => setData({ ...data, resultadosEsperados: e.target.value })}
                  className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl font-medium text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Resultados e entregas previstas"
                />
              </div>
            </div>
          </div>

          {/* Theoretical Model: Kurt Lewin's Change Management Banner */}
          <div className="bg-gradient-to-br from-slate-900 to-indigo-950 p-6 md:p-8 rounded-3xl text-white shadow-md space-y-6">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div>
                <span className="text-xs font-extrabold text-emerald-400 uppercase tracking-widest">
                  Fundamentação Teórica da Administração
                </span>
                <h3 className="text-xl font-bold mt-1">Modelo de Mudança Organizacional de Kurt Lewin & PDCA</h3>
                <p className="text-slate-300 text-xs max-w-2xl mt-1">
                  A eficácia da consultoria é mensurada pela transição do estado de fragilidade inicial (descongelamento) para a nova rotina consolidada e sustentável (recongelamento).
                </p>
              </div>

              <div className="bg-white/10 backdrop-blur-md px-4 py-3 rounded-2xl border border-white/10 flex items-center gap-3">
                <div className="text-right">
                  <span className="text-[10px] text-slate-300 uppercase block font-bold">Evolução de Maturidade</span>
                  <span className="text-lg font-black text-emerald-400">
                    {data.teoriaAdmin.maturidadeInicialPercent}% ➔ {data.teoriaAdmin.maturidadeFinalPercent}%
                  </span>
                </div>
                <div className="w-10 h-10 rounded-xl bg-emerald-500/20 flex items-center justify-center text-emerald-400">
                  <TrendingUp size={20} />
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
              <div className="p-4 rounded-2xl bg-white/5 border border-white/10 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-amber-400 uppercase text-[10px]">Fase 1: Descongelamento</span>
                  <span className="text-[10px] text-slate-400">Diagnóstico</span>
                </div>
                <p className="text-slate-300 text-[11px] leading-relaxed">
                  Sensibilização do gestor para os gargalos ocultos, quebra de velhos hábitos e percepção da urgência de controles financeiros.
                </p>
              </div>

              <div className="p-4 rounded-2xl bg-white/5 border border-white/10 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-blue-400 uppercase text-[10px]">Fase 2: Movimento (Ação)</span>
                  <span className="text-[10px] text-slate-400">30 Horas</span>
                </div>
                <p className="text-slate-300 text-[11px] leading-relaxed">
                  Implantação prática das ferramentas de fluxo de caixa, precificação, separação PF/PJ e parametrização de custos com o consultor.
                </p>
              </div>

              <div className="p-4 rounded-2xl bg-white/5 border border-white/10 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-emerald-400 uppercase text-[10px]">Fase 3: Recongelamento</span>
                  <span className="text-[10px] text-slate-400">Rotina Ativa</span>
                </div>
                <p className="text-slate-300 text-[11px] leading-relaxed">
                  Consolidação dos novos padrões operacionais e financeiros como parte inegociável da cultura e rotina permanente da empresa.
                </p>
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* TAB 2: PROBLEMAS DIAGNOSTICADOS */}
      {activeTab === 'problemas' && (
        <motion.div
          key="problemas"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <Target className="text-emerald-600" size={20} />
                  1. Avaliação dos Problemas Diagnosticados
                </h3>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-200">
                  {data.problemasAvaliados.length} problema(s) listado(s)
                </span>
              </div>
              <p className="text-slate-500 text-xs mt-0.5">
                Gargalos e fragilidades gerenciais diagnosticados especificamente para <strong>{selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || 'este cliente'}</strong>.
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={handleAddProblema}
                className="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <Plus size={14} className="text-slate-600" /> Adicionar Problema
              </button>
              <button
                onClick={autoPopulateFromDiagnostico}
                className="px-3.5 py-2 rounded-xl bg-emerald-50 text-emerald-700 hover:bg-emerald-100 text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
                title="Puxar respostas negativas do diagnóstico do cliente e agrupar por tema"
              >
                <RefreshCw size={14} className={isAutoFilling ? 'animate-spin' : ''} /> Recarregar do Diagnóstico
              </button>
            </div>
          </div>

          <div className="space-y-4">
            {data.problemasAvaliados.map((item, index) => (
              <div key={item.id || index} className="p-5 rounded-2xl bg-slate-50 border border-slate-200/80 space-y-4">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="flex items-start gap-3 flex-1">
                    <span className="w-6 h-6 rounded-full bg-emerald-600 text-white font-bold text-xs flex items-center justify-center shrink-0 mt-1">
                      {index + 1}
                    </span>
                    <div className="flex-1 space-y-1">
                      <input
                        type="text"
                        value={item.problema}
                        onChange={(e) => {
                          const updated = [...data.problemasAvaliados];
                          updated[index].problema = e.target.value;
                          setData({ ...data, problemasAvaliados: updated });
                        }}
                        className="w-full p-1.5 bg-white border border-slate-200 rounded-lg font-bold text-slate-800 text-sm focus:border-emerald-500 focus:outline-none"
                        placeholder="Descrição do problema diagnosticado"
                      />
                      <div className="flex items-center gap-2">
                        <span className="text-[11px] font-bold text-slate-500 uppercase">Área:</span>
                        <input
                          type="text"
                          value={item.area}
                          onChange={(e) => {
                            const updated = [...data.problemasAvaliados];
                            updated[index].area = e.target.value;
                            setData({ ...data, problemasAvaliados: updated });
                          }}
                          className="px-2 py-0.5 bg-white border border-slate-200 rounded text-[11px] font-semibold text-slate-600 focus:border-emerald-500 focus:outline-none"
                          placeholder="Gestão Financeira"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-3">
                    <select
                      value={item.status}
                      onChange={(e: any) => {
                        const updated = [...data.problemasAvaliados];
                        updated[index].status = e.target.value;
                        if (e.target.value === 'Resolvido') updated[index].percentualEficacia = 100;
                        if (e.target.value === 'Parcialmente Resolvido') updated[index].percentualEficacia = 70;
                        if (e.target.value === 'Em Andamento') updated[index].percentualEficacia = 40;
                        if (e.target.value === 'Não Resolvido') updated[index].percentualEficacia = 0;
                        setData({ ...data, problemasAvaliados: updated });
                      }}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700"
                    >
                      <option value="Resolvido">✅ Resolvido</option>
                      <option value="Parcialmente Resolvido">⚠️ Parcialmente Resolvido</option>
                      <option value="Em Andamento">⏳ Em Andamento</option>
                      <option value="Não Resolvido">❌ Não Resolvido</option>
                    </select>

                    <div className="flex items-center gap-2 bg-white px-3 py-1 rounded-xl border border-slate-200">
                      <span className="text-xs font-bold text-slate-500">Eficácia:</span>
                      <input
                        type="number"
                        min="0"
                        max="100"
                        value={item.percentualEficacia}
                        onChange={(e) => {
                          const updated = [...data.problemasAvaliados];
                          updated[index].percentualEficacia = Number(e.target.value);
                          setData({ ...data, problemasAvaliados: updated });
                        }}
                        className="w-12 text-center text-xs font-bold text-emerald-600 focus:outline-none"
                      />
                      <span className="text-xs font-bold text-slate-400">%</span>
                    </div>
                    <button
                      type="button"
                      onClick={() => handleDeleteProblema(index)}
                      className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                      title="Excluir este problema da avaliação"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-500 uppercase text-[10px]">Impacto Prático no Negócio</label>
                    <input
                      type="text"
                      value={item.impactoGerado}
                      onChange={(e) => {
                        const updated = [...data.problemasAvaliados];
                        updated[index].impactoGerado = e.target.value;
                        setData({ ...data, problemasAvaliados: updated });
                      }}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-medium text-slate-800"
                      placeholder="Ex: Eliminação de furos no caixa e previsibilidade de pagamentos."
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="font-bold text-slate-500 uppercase text-[10px]">Evidência / Ferramenta Utilizada</label>
                    <input
                      type="text"
                      value={item.evidencias}
                      onChange={(e) => {
                        const updated = [...data.problemasAvaliados];
                        updated[index].evidencias = e.target.value;
                        setData({ ...data, problemasAvaliados: updated });
                      }}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-medium text-slate-800"
                      placeholder="Ex: Planilha de fluxo de caixa preenchida e validada nas sessões."
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}

      {/* TAB 3: IMPLEMENTAÇÃO DE SOLUÇÕES */}
      {activeTab === 'solucoes' && (
        <motion.div
          key="solucoes"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                  <Briefcase className="text-blue-600" size={20} />
                  2. Implementação das Soluções e Ações do Cronograma (30 Horas)
                </h3>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-bold bg-blue-100 text-blue-800 border border-blue-200">
                  {data.solucoesAvaliadas.length} etapa(s)
                </span>
              </div>
              <p className="text-slate-500 text-xs mt-0.5">
                Verificação da execução das etapas estruturadas e sua integração com a rotina do cliente.
              </p>
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button
                onClick={handleAddSolucao}
                className="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
              >
                <Plus size={14} className="text-slate-600" /> Adicionar Solução
              </button>
            </div>
          </div>

          <div className="space-y-4">
            {data.solucoesAvaliadas.map((item, index) => (
              <div key={item.id || index} className="p-5 rounded-2xl bg-slate-50 border border-slate-200/80 space-y-4">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <span className="w-6 h-6 rounded-full bg-blue-600 text-white font-bold text-xs flex items-center justify-center shrink-0 mt-0.5">
                      {index + 1}
                    </span>
                    <div>
                      <h4 className="font-bold text-slate-800 text-sm">{item.solucao}</h4>
                      <span className="text-[11px] font-semibold text-slate-500">Área: {item.area}</span>
                    </div>
                  </div>

                  <div className="flex flex-wrap items-center gap-3">
                    <select
                      value={item.statusImplementacao}
                      onChange={(e: any) => {
                        const updated = [...data.solucoesAvaliadas];
                        updated[index].statusImplementacao = e.target.value;
                        setData({ ...data, solucoesAvaliadas: updated });
                      }}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700"
                    >
                      <option value="Implementada Integralmente">✅ Implementada Integralmente</option>
                      <option value="Implementada Parcialmente">⚠️ Implementada Parcialmente</option>
                      <option value="Em Implantação">⏳ Em Implantação</option>
                      <option value="Não Implementada">❌ Não Implementada</option>
                    </select>

                    <select
                      value={item.aderenciaRotina}
                      onChange={(e: any) => {
                        const updated = [...data.solucoesAvaliadas];
                        updated[index].aderenciaRotina = e.target.value;
                        setData({ ...data, solucoesAvaliadas: updated });
                      }}
                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-slate-700"
                    >
                      <option value="Alta">Aderência: Alta</option>
                      <option value="Média">Aderência: Média</option>
                      <option value="Baixa">Aderência: Baixa</option>
                    </select>
                    <button
                      type="button"
                      onClick={() => handleDeleteSolucao(index)}
                      className="p-1.5 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                      title="Excluir esta atividade da avaliação"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                  <div className="space-y-1">
                    <label className="font-bold text-slate-500 uppercase text-[10px]">Resultado Prático Observado</label>
                    <input
                      type="text"
                      value={item.resultadoObservado}
                      onChange={(e) => {
                        const updated = [...data.solucoesAvaliadas];
                        updated[index].resultadoObservado = e.target.value;
                        setData({ ...data, solucoesAvaliadas: updated });
                      }}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-medium text-slate-800"
                      placeholder="Como a solução transformou o dia a dia?"
                    />
                  </div>

                  <div className="space-y-1">
                    <label className="font-bold text-slate-500 uppercase text-[10px]">Dificuldades Encontradas / Superadas</label>
                    <input
                      type="text"
                      value={item.dificuldadesEncontradas}
                      onChange={(e) => {
                        const updated = [...data.solucoesAvaliadas];
                        updated[index].dificuldadesEncontradas = e.target.value;
                        setData({ ...data, solucoesAvaliadas: updated });
                      }}
                      className="w-full p-2.5 bg-white border border-slate-200 rounded-xl font-medium text-slate-800"
                      placeholder="Desafios técnicos ou culturais durante a implementação."
                    />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}

      {/* TAB 4: ÍNDICES FINANCEIROS */}
      {activeTab === 'financeiro' && (
        <motion.div
          key="financeiro"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <DollarSign className="text-amber-600" size={20} />
                3. Matriz de Acompanhamento das Práticas de Gestão Financeira
              </h3>
              <p className="text-slate-500 text-xs mt-0.5">
                Avaliação da assimilação dos 8 controles fundamentais de gestão financeira empresarial.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {data.indicadoresFinanceiros.map((ind, index) => (
              <div key={ind.id} className="p-5 rounded-2xl bg-slate-50 border border-slate-200 space-y-3">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <h4 className="font-bold text-slate-800 text-sm">{ind.nome}</h4>
                    <p className="text-[11px] text-slate-500 mt-0.5 leading-relaxed">{ind.descricao}</p>
                  </div>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 text-xs">
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Status</label>
                    <select
                      value={ind.status}
                      onChange={(e: any) => {
                        const updated = [...data.indicadoresFinanceiros];
                        updated[index].status = e.target.value;
                        setData({ ...data, indicadoresFinanceiros: updated });
                      }}
                      className="w-full p-2 bg-white border border-slate-200 rounded-xl font-bold text-xs text-slate-700"
                    >
                      <option value="Ativo e Consolidado">Ativo e Consolidado</option>
                      <option value="Em Implantação">Em Implantação</option>
                      <option value="Parcial / Requer Atenção">Parcial / Requer Atenção</option>
                      <option value="Não Acompanhado">Não Acompanhado</option>
                    </select>
                  </div>

                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Frequência</label>
                    <select
                      value={ind.frequencia}
                      onChange={(e: any) => {
                        const updated = [...data.indicadoresFinanceiros];
                        updated[index].frequencia = e.target.value;
                        setData({ ...data, indicadoresFinanceiros: updated });
                      }}
                      className="w-full p-2 bg-white border border-slate-200 rounded-xl font-bold text-xs text-slate-700"
                    >
                      <option value="Diária">Diária</option>
                      <option value="Semanal">Semanal</option>
                      <option value="Mensal">Mensal</option>
                      <option value="Não Realiza">Não Realiza</option>
                    </select>
                  </div>

                  <div className="space-y-1 col-span-2 sm:col-span-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Qualidade</label>
                    <select
                      value={ind.qualidadeRegistro}
                      onChange={(e: any) => {
                        const updated = [...data.indicadoresFinanceiros];
                        updated[index].qualidadeRegistro = e.target.value;
                        setData({ ...data, indicadoresFinanceiros: updated });
                      }}
                      className="w-full p-2 bg-white border border-slate-200 rounded-xl font-bold text-xs text-slate-700"
                    >
                      <option value="Excelente">Excelente</option>
                      <option value="Boa">Boa</option>
                      <option value="Regular">Regular</option>
                      <option value="Incipiente">Incipiente</option>
                    </select>
                  </div>
                </div>

                <div className="space-y-1 text-xs">
                  <label className="text-[10px] font-bold text-slate-400 uppercase">Observações do Consultor</label>
                  <input
                    type="text"
                    value={ind.observacao}
                    onChange={(e) => {
                      const updated = [...data.indicadoresFinanceiros];
                      updated[index].observacao = e.target.value;
                      setData({ ...data, indicadoresFinanceiros: updated });
                    }}
                    className="w-full p-2 bg-white border border-slate-200 rounded-xl font-medium text-slate-800 text-xs"
                    placeholder="Evidência ou recomendação técnica"
                  />
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}

      {/* TAB 5: ENGAJAMENTO DO CLIENTE */}
      {activeTab === 'engajamento' && (
        <motion.div
          key="engajamento"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <UsersIcon className="text-indigo-600" />
                4. Avaliação de Engajamento e Participação do Cliente
              </h3>
              <p className="text-slate-500 text-xs mt-0.5">
                Mensuração do comprometimento da liderança, pontualidade nas sessões e cumprimento das tarefas orientadas.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="space-y-5">
              {[
                { key: 'comprometimentoLideranca', label: 'Comprometimento da Liderança / Gestão' },
                { key: 'cumprimentoPrazosTarefas', label: 'Cumprimento de Prazos e Deveres de Casa' },
                { key: 'disponibilizacaoDadosDocumentos', label: 'Transparência e Entrega de Documentos/Extratos' },
                { key: 'participacaoTreinamentos', label: 'Presença e Pontualidade nos Encontros Técnicos' },
                { key: 'envolvimentoEquipe', label: 'Envolvimento da Equipe Operacional' }
              ].map((item) => (
                <div key={item.key} className="flex items-center justify-between p-3.5 bg-slate-50 rounded-2xl border border-slate-200">
                  <span className="text-xs font-bold text-slate-700">{item.label}</span>
                  <div className="flex items-center gap-1">
                    {[1, 2, 3, 4, 5].map((star) => (
                      <button
                        key={star}
                        type="button"
                        onClick={() => {
                          setData({
                            ...data,
                            engajamentoCliente: {
                              ...data.engajamentoCliente,
                              [item.key]: star
                            }
                          });
                        }}
                        className="p-1 cursor-pointer transition-transform hover:scale-110"
                      >
                        <Star
                          size={18}
                          className={`${
                            star <= (data.engajamentoCliente as any)[item.key]
                              ? 'text-amber-400 fill-amber-400'
                              : 'text-slate-300'
                          }`}
                        />
                      </button>
                    ))}
                    <span className="text-xs font-black text-slate-800 ml-1.5 w-6 text-right">
                      {(data.engajamentoCliente as any)[item.key]}/5
                    </span>
                  </div>
                </div>
              ))}
            </div>

            <div className="space-y-4">
              <div className="space-y-1.5">
                <label className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                  Parecer Qualitativo sobre o Engajamento
                </label>
                <textarea
                  rows={6}
                  value={data.engajamentoCliente.observacoesEngajamento}
                  onChange={(e) => setData({
                    ...data,
                    engajamentoCliente: {
                      ...data.engajamentoCliente,
                      observacoesEngajamento: e.target.value
                    }
                  })}
                  className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium text-slate-800 focus:bg-white focus:border-emerald-500"
                  placeholder="Relate a postura do empresário e da equipe durante a consultoria..."
                />
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* TAB 6: PESQUISA & NPS */}
      {activeTab === 'satisfacao' && (
        <motion.div
          key="satisfacao"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <Star className="text-purple-600" size={20} />
                5. Pesquisa de Satisfação & NPS (Net Promoter Score)
              </h3>
              <p className="text-slate-500 text-xs mt-0.5">
                Registro formal da avaliação do cliente quanto aos serviços, didática do consultor e superação de expectativas.
              </p>
            </div>
          </div>

          <div className="space-y-6">
            {/* NPS 0-10 Selector */}
            <div className="p-6 rounded-3xl bg-purple-50/50 border border-purple-100 text-center space-y-4">
              <span className="text-xs font-extrabold text-purple-900 uppercase tracking-wider">
                Em uma escala de 0 a 10, o quanto você recomendaria esta consultoria a outro empresário?
              </span>
              <div className="flex flex-wrap items-center justify-center gap-2">
                {[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map((score) => (
                  <button
                    key={score}
                    type="button"
                    onClick={() => setData({
                      ...data,
                      pesquisaSatisfacao: {
                        ...data.pesquisaSatisfacao,
                        npsConsultoria: score
                      }
                    })}
                    className={`w-10 h-10 rounded-2xl font-black text-sm transition-all cursor-pointer ${
                      data.pesquisaSatisfacao.npsConsultoria === score
                        ? 'bg-purple-600 text-white scale-110 shadow-md shadow-purple-200'
                        : 'bg-white text-slate-700 border border-slate-200 hover:border-purple-300'
                    }`}
                  >
                    {score}
                  </button>
                ))}
              </div>
              <div className="flex items-center justify-between text-[11px] font-bold text-slate-500 max-w-md mx-auto">
                <span className="text-rose-500">0 = Jamais recomendaria</span>
                <span className="text-emerald-600">10 = Recomendo fortemente</span>
              </div>
            </div>

            {/* Satisfaction fields */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
              <div className="space-y-1">
                <label className="font-bold text-slate-700">Depoimento do Cliente sobre a Transformação</label>
                <textarea
                  rows={3}
                  value={data.pesquisaSatisfacao.depoimentoCliente}
                  onChange={(e) => setData({
                    ...data,
                    pesquisaSatisfacao: {
                      ...data.pesquisaSatisfacao,
                      depoimentoCliente: e.target.value
                    }
                  })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-2xl font-medium text-slate-800"
                  placeholder="Depoimento oficial que constará no relatório final..."
                />
              </div>

              <div className="space-y-1">
                <label className="font-bold text-slate-700">Pontos Fortes da Consultoria</label>
                <textarea
                  rows={3}
                  value={data.pesquisaSatisfacao.pontosFortesConsultoria}
                  onChange={(e) => setData({
                    ...data,
                    pesquisaSatisfacao: {
                      ...data.pesquisaSatisfacao,
                      pontosFortesConsultoria: e.target.value
                    }
                  })}
                  className="w-full p-3 bg-slate-50 border border-slate-200 rounded-2xl font-medium text-slate-800"
                  placeholder="Quais foram os maiores diferenciais do trabalho?"
                />
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* TAB 7: TEORIA DA ADMINISTRAÇÃO & CONTINUIDADE */}
      {activeTab === 'teoria' && (
        <motion.div
          key="teoria"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="bg-white p-6 md:p-8 rounded-3xl border border-slate-100 shadow-sm space-y-6"
        >
          <div className="flex items-center justify-between border-b border-slate-100 pb-4">
            <div>
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <BookOpen className="text-emerald-600" size={20} />
                6. Parecer Técnico Conclusivo & Plano de Continuidade (Pós-Consultoria)
              </h3>
              <p className="text-slate-500 text-xs mt-0.5">
                Diretrizes do consultor para manter a disciplina financeira nos próximos 30, 60 e 90 dias.
              </p>
            </div>
          </div>

          <div className="space-y-5 text-xs">
            <div className="space-y-1.5">
              <label className="font-bold text-slate-800 uppercase text-[11px]">
                Parecer Técnico Conclusivo do Consultor
              </label>
              <textarea
                rows={4}
                value={data.teoriaAdmin.parecerTecnicoConsultor}
                onChange={(e) => setData({
                  ...data,
                  teoriaAdmin: {
                    ...data.teoriaAdmin,
                    parecerTecnicoConsultor: e.target.value
                  }
                })}
                className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl font-medium text-slate-800 focus:bg-white"
                placeholder="Síntese técnica da consultoria e maturidade alcançada..."
              />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div className="p-4 rounded-2xl bg-blue-50/40 border border-blue-100 space-y-1.5">
                <span className="font-extrabold text-blue-900 uppercase text-[10px] block">
                  Metas para os Próximos 30 Dias (Pós-Consultoria)
                </span>
                <textarea
                  rows={4}
                  value={data.teoriaAdmin.planoContinuidade30Dias}
                  onChange={(e) => setData({
                    ...data,
                    teoriaAdmin: {
                      ...data.teoriaAdmin,
                      planoContinuidade30Dias: e.target.value
                    }
                  })}
                  className="w-full p-2.5 bg-white border border-blue-200 rounded-xl font-medium text-slate-800 text-xs"
                />
              </div>

              <div className="p-4 rounded-2xl bg-indigo-50/40 border border-indigo-100 space-y-1.5">
                <span className="font-extrabold text-indigo-900 uppercase text-[10px] block">
                  Metas para os Próximos 60 Dias (Pós-Consultoria)
                </span>
                <textarea
                  rows={4}
                  value={data.teoriaAdmin.planoContinuidade60Dias}
                  onChange={(e) => setData({
                    ...data,
                    teoriaAdmin: {
                      ...data.teoriaAdmin,
                      planoContinuidade60Dias: e.target.value
                    }
                  })}
                  className="w-full p-2.5 bg-white border border-indigo-200 rounded-xl font-medium text-slate-800 text-xs"
                />
              </div>

              <div className="p-4 rounded-2xl bg-emerald-50/40 border border-emerald-100 space-y-1.5">
                <span className="font-extrabold text-emerald-900 uppercase text-[10px] block">
                  Metas para os Próximos 90 Dias (Pós-Consultoria)
                </span>
                <textarea
                  rows={4}
                  value={data.teoriaAdmin.planoContinuidade90Dias}
                  onChange={(e) => setData({
                    ...data,
                    teoriaAdmin: {
                      ...data.teoriaAdmin,
                      planoContinuidade90Dias: e.target.value
                    }
                  })}
                  className="w-full p-2.5 bg-white border border-emerald-200 rounded-xl font-medium text-slate-800 text-xs"
                />
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* TAB 8: RELATÓRIO FORMAL IMPRESSO / PDF COM LOGO DO SEBRAE E ASSINATURAS */}
      {activeTab === 'relatorio' && (
        <motion.div
          key="relatorio"
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="space-y-6"
        >
          {/* Action Bar for Report */}
          <div className="flex flex-wrap items-center justify-between gap-4 bg-white p-5 rounded-2xl border border-slate-100 shadow-xs print:hidden">
            <div className="flex flex-wrap items-center gap-3">
              <span className="text-xs font-bold text-slate-600 flex items-center gap-1.5">
                <ImageIcon size={16} className="text-emerald-600" />
                Exibição de Logotipo no Relatório:
              </span>
              <div className="flex items-center bg-slate-100 p-1 rounded-xl">
                <button
                  type="button"
                  onClick={() => {
                    setLocalLogoChoice('sebrae');
                    if (setLogoChoice) setLogoChoice('sebrae');
                  }}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    localLogoChoice === 'sebrae' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  SEBRAE
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setLocalLogoChoice('consultora');
                    if (setLogoChoice) setLogoChoice('consultora');
                  }}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    localLogoChoice === 'consultora' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Empresa Consultora
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setLocalLogoChoice('ambos');
                    if (setLogoChoice) setLogoChoice('ambos');
                  }}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    localLogoChoice === 'ambos' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Ambos (SEBRAE + Consultora)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setLocalLogoChoice('none');
                    if (setLogoChoice) setLogoChoice('none');
                  }}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    localLogoChoice === 'none' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Nenhum
                </button>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={() => setView('settings')}
                className="text-xs font-bold text-slate-500 hover:text-emerald-700 underline cursor-pointer"
              >
                Gerenciar Logotipos em Configurações
              </button>
              <button
                onClick={handlePrint}
                className="inline-flex items-center gap-2 px-6 py-2.5 text-xs font-bold text-white bg-slate-900 hover:bg-slate-800 rounded-xl shadow-sm transition-all cursor-pointer"
              >
                <Printer size={16} /> Imprimir / Exportar PDF (A4)
              </button>
            </div>
          </div>

          {/* Printable Report Container */}
          <div
            ref={reportRef}
            className="bg-white p-8 md:p-12 rounded-3xl border border-slate-200 shadow-sm print:p-0 print:border-none print:shadow-none max-w-5xl mx-auto font-sans text-slate-800 space-y-8"
          >
            {/* Header with Official Logos and Formal Title */}
            <div className="flex items-center justify-between border-b-2 border-slate-800 pb-6">
              <div className="flex items-center gap-5">
                {/* SEBRAE Logo Rendering */}
                {(localLogoChoice === 'sebrae' || localLogoChoice === 'ambos') && (
                  effectiveSebraeLogo ? (
                    <img
                      src={effectiveSebraeLogo}
                      alt="Logo SEBRAE"
                      className="h-14 max-w-[180px] object-contain"
                      referrerPolicy="no-referrer"
                    />
                  ) : (
                    <div className="px-4 py-2.5 bg-blue-700 rounded-xl text-white font-black text-lg tracking-tighter shadow-xs">
                      SEBRAE
                    </div>
                  )
                )}

                {/* Consultora Logo Rendering */}
                {(localLogoChoice === 'consultora' || localLogoChoice === 'ambos') && effectiveConsultoraLogo && (
                  <div className={`${localLogoChoice === 'ambos' ? 'pl-4 border-l border-slate-300' : ''}`}>
                    <img
                      src={effectiveConsultoraLogo}
                      alt="Logo Consultora"
                      className="h-14 max-w-[180px] object-contain"
                      referrerPolicy="no-referrer"
                    />
                  </div>
                )}

                <div>
                  <h1 className="text-xl font-extrabold text-slate-900 uppercase tracking-tight">
                    Relatório de Análise de Resultados da Consultoria
                  </h1>
                  <p className="text-xs font-semibold text-slate-500 uppercase tracking-wider">
                    Avaliação de Eficácia, Resolução de Problemas e Sustentabilidade Gerencial
                  </p>
                </div>
              </div>

              <div className="text-right text-xs text-slate-500 space-y-0.5">
                <p className="font-bold text-slate-900 text-sm">Código SGF: {data.codigoSgf || '---'}</p>
                <p>Carga Horária: <strong>{data.cargaHorariaRealizada}</strong></p>
                <p>Data de Emissão: {data.dataAnalise}</p>
              </div>
            </div>

            {/* Client & Consulting Contract Details Table */}
            <div className="bg-slate-50 p-5 rounded-2xl border border-slate-200 text-xs space-y-3">
              <h3 className="font-bold text-slate-800 uppercase tracking-wider text-[11px] border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                <Building2 size={14} className="text-emerald-600" /> Dados do Contrato e da Consultoria Gerencial
              </h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Razão Social / Cliente:</span>
                  <span className="font-bold text-slate-800">{data.clienteResponsavel || selectedEmpresa?.nome || 'Não informado'}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">CNPJ / CPF:</span>
                  <span className="font-semibold text-slate-800">{data.clienteCpfCnpj || selectedEmpresa?.cnpj || 'Não informado'}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Consultor Responsável:</span>
                  <span className="font-bold text-slate-800">{data.consultorResponsavel || 'Consultor Técnico'}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Gestor / Técnico Sebrae:</span>
                  <span className="font-semibold text-slate-800">{data.tecnicoSebrae || 'Gestor do Projeto'}</span>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1 border-t border-slate-200/50">
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Empresa Credenciada:</span>
                  <span className="font-semibold text-slate-800">{data.empresaCredenciada || credenciadaName || 'Empresa Credenciada Sebrae'}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Área da Consultoria:</span>
                  <span className="font-semibold text-slate-800">{data.areaConsultoria}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Período de Execução:</span>
                  <span className="font-semibold text-slate-800">{data.periodoConsultoria || '30 Horas de Atendimento'}</span>
                </div>
              </div>

              {data.objetivoConsultoria && (
                <div className="pt-2 border-t border-slate-200/60">
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Objetivo Contratual:</span>
                  <p className="text-slate-700 italic text-[11px] mt-0.5">{data.objetivoConsultoria}</p>
                </div>
              )}
            </div>

            {/* Executive Dashboard Indicators */}
            <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 text-center">
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Score Geral</span>
                <span className="text-2xl font-black text-emerald-600">{calculatedMetrics.globalScore}%</span>
                <span className="text-[9px] text-slate-500 block font-semibold">{calculatedMetrics.nivelClassificacao.split(' ')[0]}</span>
              </div>

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Gargalos Sanados</span>
                <span className="text-2xl font-black text-emerald-600">{calculatedMetrics.probRate}%</span>
                <span className="text-[9px] text-slate-500 block font-semibold">{calculatedMetrics.resolvidosCount}/{calculatedMetrics.totalProb} Sanados</span>
              </div>

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Soluções Implantadas</span>
                <span className="text-2xl font-black text-blue-600">{calculatedMetrics.solRate}%</span>
                <span className="text-[9px] text-slate-500 block font-semibold">{calculatedMetrics.solIntegralCount}/{calculatedMetrics.totalSol} Ações</span>
              </div>

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl">
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Índices Financeiros</span>
                <span className="text-2xl font-black text-amber-600">{calculatedMetrics.indRate}%</span>
                <span className="text-[9px] text-slate-500 block font-semibold">{calculatedMetrics.indConsolidadosCount}/{calculatedMetrics.totalInd} Controles</span>
              </div>

              <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl col-span-2 sm:col-span-1">
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Satisfação NPS</span>
                <span className="text-2xl font-black text-purple-600">{data.pesquisaSatisfacao.npsConsultoria}/10</span>
                <span className="text-[9px] text-slate-500 block font-semibold">Avaliação Excelente</span>
              </div>
            </div>

            {/* Section 1: Problemas Diagnosticados vs. Resolução */}
            <div className="space-y-3">
              <h3 className="font-bold text-slate-900 uppercase text-xs tracking-wider border-b border-slate-200 pb-1 flex items-center gap-1.5">
                <Target size={14} className="text-emerald-600" /> 1. Avaliação dos Problemas Diagnosticados no Cliente
              </h3>
              <div className="overflow-hidden border border-slate-200 rounded-xl">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                    <tr>
                      <th className="p-2.5">Problema Identificado</th>
                      <th className="p-2.5">Área</th>
                      <th className="p-2.5">Status de Resolução</th>
                      <th className="p-2.5 text-center">Eficácia</th>
                      <th className="p-2.5">Impacto / Evidência</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {data.problemasAvaliados.map((item, idx) => (
                      <tr key={item.id || idx} className="hover:bg-slate-50">
                        <td className="p-2.5 font-medium text-slate-800 max-w-xs">{item.problema}</td>
                        <td className="p-2.5 text-slate-500">{item.area}</td>
                        <td className="p-2.5">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            item.status === 'Resolvido' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'
                          }`}>
                            {item.status}
                          </span>
                        </td>
                        <td className="p-2.5 text-center font-bold text-emerald-700">{item.percentualEficacia}%</td>
                        <td className="p-2.5 text-slate-600 text-[11px] max-w-xs">{item.impactoGerado || item.evidencias || '---'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Section 2: Implementação das Soluções */}
            <div className="space-y-3">
              <h3 className="font-bold text-slate-900 uppercase text-xs tracking-wider border-b border-slate-200 pb-1 flex items-center gap-1.5">
                <Briefcase size={14} className="text-blue-600" /> 2. Implementação das Soluções e Ações do Cronograma ({data.cargaHorariaRealizada})
              </h3>
              <div className="overflow-hidden border border-slate-200 rounded-xl">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                    <tr>
                      <th className="p-2.5">Solução / Atividade Orientada</th>
                      <th className="p-2.5">Status de Implantação</th>
                      <th className="p-2.5 text-center">Aderência à Rotina</th>
                      <th className="p-2.5">Resultado Prático Observado</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {data.solucoesAvaliadas.map((item, idx) => (
                      <tr key={item.id || idx} className="hover:bg-slate-50">
                        <td className="p-2.5 font-medium text-slate-800">{item.solucao}</td>
                        <td className="p-2.5">
                          <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                            item.statusImplementacao === 'Implementada Integralmente' ? 'bg-emerald-100 text-emerald-800' : 'bg-blue-100 text-blue-800'
                          }`}>
                            {item.statusImplementacao}
                          </span>
                        </td>
                        <td className="p-2.5 text-center font-semibold text-slate-700">{item.aderenciaRotina}</td>
                        <td className="p-2.5 text-slate-600 text-[11px]">{item.resultadoObservado || 'Ação cumprida com êxito.'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Section 3: Gestão Financeira & Rotinas */}
            <div className="space-y-3">
              <h3 className="font-bold text-slate-900 uppercase text-xs tracking-wider border-b border-slate-200 pb-1 flex items-center gap-1.5">
                <DollarSign size={14} className="text-amber-600" /> 3. Matriz de Acompanhamento das Práticas de Gestão Financeira
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2.5 text-xs">
                {data.indicadoresFinanceiros.map((ind) => (
                  <div key={ind.id} className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between gap-2">
                    <div>
                      <span className="font-bold text-slate-800 block text-[11px]">{ind.nome}</span>
                      <span className="text-[10px] text-slate-500">Frequência: <strong>{ind.frequencia}</strong></span>
                    </div>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                      ind.status === 'Ativo e Consolidado' ? 'bg-emerald-100 text-emerald-800' : 'bg-blue-100 text-blue-800'
                    }`}>
                      {ind.status}
                    </span>
                  </div>
                ))}
              </div>
            </div>

            {/* Section 4: Parecer Técnico e Continuidade */}
            <div className="space-y-3 border-t border-slate-200 pt-4">
              <h3 className="font-bold text-slate-900 uppercase text-xs tracking-wider flex items-center gap-1.5">
                <BookOpen size={14} className="text-purple-600" /> 4. Parecer Técnico do Consultor e Plano de Continuidade
              </h3>
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 text-xs text-slate-800 leading-relaxed italic">
                "{data.teoriaAdmin.parecerTecnicoConsultor}"
              </div>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-3 text-xs">
                <div className="p-3 bg-blue-50/50 rounded-lg border border-blue-100">
                  <span className="font-bold text-blue-900 block text-[10px] uppercase">Próximos 30 Dias</span>
                  <p className="text-slate-700 text-[11px] mt-1">{data.teoriaAdmin.planoContinuidade30Dias}</p>
                </div>
                <div className="p-3 bg-indigo-50/50 rounded-lg border border-indigo-100">
                  <span className="font-bold text-indigo-900 block text-[10px] uppercase">Próximos 60 Dias</span>
                  <p className="text-slate-700 text-[11px] mt-1">{data.teoriaAdmin.planoContinuidade60Dias}</p>
                </div>
                <div className="p-3 bg-emerald-50/50 rounded-lg border border-emerald-100">
                  <span className="font-bold text-emerald-900 block text-[10px] uppercase">Próximos 90 Dias</span>
                  <p className="text-slate-700 text-[11px] mt-1">{data.teoriaAdmin.planoContinuidade90Dias}</p>
                </div>
              </div>
            </div>

            {/* Section 5: Pesquisa de Satisfação e Depoimento do Cliente */}
            {data.pesquisaSatisfacao.depoimentoCliente && (
              <div className="p-4 bg-emerald-50/40 rounded-xl border border-emerald-100 text-xs space-y-1">
                <span className="font-bold text-emerald-900 text-[11px] uppercase tracking-wider block">
                  Declaração de Satisfação do Cliente (NPS: {data.pesquisaSatisfacao.npsConsultoria}/10)
                </span>
                <p className="text-slate-800 italic leading-relaxed">
                  "{data.pesquisaSatisfacao.depoimentoCliente}"
                </p>
              </div>
            )}

            {/* FORMAL SIGNATURE BLOCK AT THE BOTTOM */}
            <div className="pt-12 border-t-2 border-slate-800 space-y-8 break-inside-avoid">
              <div className="text-center text-xs text-slate-500 mb-8">
                Declaro para os devidos fins que os serviços de consultoria gerencial foram prestados a contento, tendo sido transferidos os conhecimentos técnicos e ferramentas necessárias para a gestão do empreendimento.
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-12 text-center text-xs">
                {/* Signature 1: Cliente / Representante Legal */}
                <div className="space-y-2 flex flex-col items-center">
                  <div className="w-72 border-b border-slate-900 pb-1"></div>
                  <span className="font-bold text-slate-900 text-sm block">
                    {data.clienteResponsavel || selectedEmpresa?.nome || 'Representante Legal do Cliente'}
                  </span>
                  <span className="text-slate-500 text-xs block">
                    {data.clienteCargo || 'Proprietário / Gestor Responsável'}
                  </span>
                  <span className="text-slate-400 text-[11px] font-mono block">
                    CNPJ/CPF: {data.clienteCpfCnpj || selectedEmpresa?.cnpj || '__________________'}
                  </span>
                </div>

                {/* Signature 2: Consultor Responsável */}
                <div className="space-y-2 flex flex-col items-center">
                  <div className="w-72 border-b border-slate-900 pb-1"></div>
                  <span className="font-bold text-slate-900 text-sm block">
                    {data.consultorResponsavel || 'Consultor Técnico Responsável'}
                  </span>
                  <span className="text-slate-500 text-xs block">
                    Consultoria Gerencial & Financeira
                  </span>
                  <span className="text-slate-400 text-[11px] block">
                    {data.empresaCredenciada || credenciadaName || 'Empresa Credenciada Sebrae'}
                  </span>
                </div>
              </div>

              <div className="text-center text-[10px] text-slate-400 pt-6">
                Documento gerado eletronicamente através do Sistema Integrado de Consultoria Empresarial e Gestão Estratégica.
              </div>
            </div>
          </div>
        </motion.div>
      )}
    </div>
  );
};

function UsersIcon(props: any) {
  return (
    <svg
      {...props}
      xmlns="http://www.w3.org/2000/svg"
      width="20"
      height="20"
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth="2"
      strokeLinecap="round"
      strokeLinejoin="round"
    >
      <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
      <circle cx="9" cy="7" r="4" />
      <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
      <path d="M16 3.13a4 4 0 0 1 0 7.75" />
    </svg>
  );
}
