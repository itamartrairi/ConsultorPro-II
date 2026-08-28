import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  Award,
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
  Percent
} from 'lucide-react';
import { db, auth } from '../firebase';
import { doc, setDoc, getDoc, serverTimestamp, collection, getDocs } from 'firebase/firestore';

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
  areaConsultoria?: string;
  objetivoConsultoria?: string;
  
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
  logoChoice?: 'sebrae' | 'consultora' | 'none';
  setLogoChoice?: (c: 'sebrae' | 'consultora' | 'none') => void;
  playSuccessSound: () => void;
  setView: (v: string) => void;
  setSelectedDiagnostico: (d: any) => void;
  setDiagnosticos: React.Dispatch<React.SetStateAction<any[]>>;
  empresas?: any[];
  diagnosticos?: any[];
}

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
  const reportRef = useRef<HTMLDivElement>(null);

  // Initial Form Data State
  const [data, setData] = useState<AnaliseResultadosData>(() => {
    // 1. Try to load existing data from diagnostico object or localStorage
    const savedInDiag = selectedDiagnostico?.analiseResultados;
    if (savedInDiag && typeof savedInDiag === 'object') {
      return savedInDiag;
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

    // Default scaffold
    return {
      diagnosticoId: selectedDiagnostico?.id || '',
      empresaId: selectedEmpresa?.id || selectedDiagnostico?.empresaId || '',
      dataAnalise: new Date().toISOString().split('T')[0],
      dataEncerramento: new Date().toISOString().split('T')[0],
      cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || '30 Horas',
      consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || '',
      consultorDocumento: '',
      clienteResponsavel: selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || '',
      clienteCargo: 'Proprietário / Gestor',
      clienteCpfCnpj: selectedEmpresa?.cnpj || selectedDiagnostico?.dadosConsultoria?.cnpj || '',
      tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || '',
      empresaCredenciada: '',
      codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || '',
      areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || 'Gestão Financeira e Produção',
      objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || 'Implementar melhorias estratégicas e práticas nas áreas de Gestão Financeira, Preço e Mercado, Planejamento, Custos e Resultado, Acesso a Crédito, Controle e Rotina e Produção, visando aprimorar a gestão do empreendimento, fortalecer os controles gerenciais e aumentar a eficiência.',
      
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
        pontosFortesConsultoria: 'Praticidade nas planilhas, paciência e didática do consultor na explicação e foco na realidade do nosso segmento.',
        sugestoesMelhoria: 'Manter acompanhamento de pós-consultoria semestral para reforçar os novos controles.'
      },
      teoriaAdmin: {
        faseCicloMudanca: 'Recongelamento (Consolidação da Rotina)',
        maturidadeInicialPercent: selectedDiagnostico?.percentualGeral || 35,
        maturidadeFinalPercent: 88,
        parecerTecnicoConsultor: 'O cliente concluiu com pleno êxito o programa de consultoria gerencial de 30 horas. Foram sanadas as principais fragilidades financeiras, estabelecida a disciplina de caixa e segregadas as contas pessoais dos custos operacionais.',
        principaisGanhosObtidos: '1. Clareza total da margem real de lucro.\n2. Eliminação de despesas ocultas e juros bancários desnecessários.\n3. Domínio da ferramenta gerencial de caixa e custos.\n4. Autonomia na precificação técnica.',
        riscosResiduais: 'Risco de afrouxamento na disciplina de alimentação diária do fluxo de caixa em períodos de alta intensidade de trabalho no campo.',
        planoContinuidade30Dias: 'Realizar o primeiro fechamento mensal de resultado (DRE) de forma 100% autônoma e conferir a conciliação bancária semanalmente.',
        planoContinuidade60Dias: 'Apurar a margem de contribuição acumulada da safra/período e destinar 20% do lucro líquido apurado para o fundo de reserva.',
        planoContinuidade90Dias: 'Revisar a tabela de preços e custos unitários antes do início do novo ciclo produtivo.',
        recomendacoesFinais: 'Manter a reunião financeira quinzenal de fechamento e não realizar retiradas além do pró-labore estipulado.'
      }
    };
  });

  // Re-sync when selected diagnostico changes
  useEffect(() => {
    if (selectedDiagnostico?.analiseResultados) {
      setData(selectedDiagnostico.analiseResultados);
    } else {
      // populate defaults from selected diagnosis
      setData(prev => ({
        ...prev,
        diagnosticoId: selectedDiagnostico?.id || prev.diagnosticoId,
        empresaId: selectedEmpresa?.id || selectedDiagnostico?.empresaId || prev.empresaId,
        clienteResponsavel: selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || prev.clienteResponsavel,
        clienteCpfCnpj: selectedEmpresa?.cnpj || selectedDiagnostico?.dadosConsultoria?.cnpj || prev.clienteCpfCnpj,
        consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || prev.consultorResponsavel,
        tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || prev.tecnicoSebrae,
        codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || prev.codigoSgf,
        cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || prev.cargaHorariaRealizada,
        objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || prev.objetivoConsultoria,
        teoriaAdmin: {
          ...prev.teoriaAdmin,
          maturidadeInicialPercent: selectedDiagnostico?.percentualGeral ?? prev.teoriaAdmin.maturidadeInicialPercent
        }
      }));
    }
  }, [selectedDiagnostico?.id]);

  // Extract Diagnostic Problems Automatically if empty
  useEffect(() => {
    if (data.problemasAvaliados.length === 0 && selectedDiagnostico?.id) {
      autoPopulateFromDiagnostico();
    }
  }, [selectedDiagnostico?.id, respostas]);

  const autoPopulateFromDiagnostico = () => {
    setIsAutoFilling(true);

    // 1. Identify problems from diagnostic answers (answers === 'Não' or 'Parcial')
    const diagRespostas = respostas.filter(r => r.diagnosticoId === selectedDiagnostico?.id);
    const negativeRespostas = diagRespostas.filter(r => r.resposta === 'Não' || r.resposta === 'Parcial');

    const generatedProblemas: ProblemaAvaliado[] = [];

    if (negativeRespostas.length > 0) {
      negativeRespostas.forEach((r, idx) => {
        const probText = r.problema || r.pergunta || `Fragilidade em ${r.area || 'Gestão'}`;
        generatedProblemas.push({
          id: `prob_${idx}_${Date.now()}`,
          problema: probText,
          area: r.area || 'Gestão Financeira',
          status: 'Resolvido',
          percentualEficacia: r.resposta === 'Não' ? 90 : 95,
          impactoGerado: `Processo e controle estruturados durante a consultoria para sanar a falha diagnosticada em "${probText}".`,
          evidencias: 'Planilha / rotina implementada, acompanhada e validada com o consultor.'
        });
      });
    } else {
      // Fallback: use default financial management issues if diagnosis has no direct negative flags
      const defaults = [
        { problema: 'Ausência de controle diário e rigoroso de entradas e saídas de caixa', area: 'Gestão Financeira' },
        { problema: 'Mistura de despesas pessoais com as despesas da empresa (falta de pró-labore)', area: 'Controle e Rotina' },
        { problema: 'Desconhecimento dos custos reais e da margem de contribuição', area: 'Custos e Resultado' },
        { problema: 'Precificação baseada apenas no mercado sem apuração dos custos unitários', area: 'Preço e Mercado' },
        { problema: 'Inexistência de reserva financeira para emergências e entressafras', area: 'Gestão Financeira' }
      ];
      defaults.forEach((item, idx) => {
        generatedProblemas.push({
          id: `prob_def_${idx}`,
          problema: item.problema,
          area: item.area,
          status: 'Resolvido',
          percentualEficacia: 90,
          impactoGerado: 'Implementação de rotina prática e treinamento direto com o gestor.',
          evidencias: 'Registros operacionais e planilhas gerenciais ativas.'
        });
      });
    }

    // 2. Identify solutions from Cronograma (30h activities) or Soluções library
    const cronograma = selectedDiagnostico?.cronograma || [];
    const generatedSolucoes: SolucaoAvaliada[] = [];

    if (cronograma.length > 0) {
      cronograma.forEach((atv: any, idx: number) => {
        if (atv.nome && atv.nome.trim()) {
          generatedSolucoes.push({
            id: `sol_${idx}_${Date.now()}`,
            solucao: `${atv.nome}: ${atv.solucaoProposta || atv.descricao || 'Atividade executada'}`,
            area: atv.premissa || 'Gestão & Finanças',
            statusImplementacao: atv.status === 'Concluído' ? 'Implementada Integralmente' : 'Em Implantação',
            aderenciaRotina: 'Alta',
            resultadoObservado: atv.resultadoEsperado || 'Ação concluída com geração dos relatórios e rotinas acordadas.',
            dificuldadesEncontradas: 'Nenhuma impeditiva; superadas com a orientação técnica.'
          });
        }
      });
    } else {
      const defaultSolucoes = [
        { solucao: 'Implantação da Planilha Gerencial de Fluxo de Caixa e Treinamento do Gestor', area: 'Gestão Financeira' },
        { solucao: 'Formalização de Contas Bancárias Distintas e Fixação de Pró-labore', area: 'Controle e Rotina' },
        { solucao: 'Mapeamento Detalhado da Estrutura de Custos Fixos e Variáveis', area: 'Custos e Resultado' },
        { solucao: 'Estruturação do Simulador de Preço de Venda com Base em Custos Reais', area: 'Preço e Mercado' },
        { solucao: 'Elaboração do Plano de Reserva de Emergência e Formação de Capital de Giro', area: 'Planejamento' }
      ];
      defaultSolucoes.forEach((item, idx) => {
        generatedSolucoes.push({
          id: `sol_def_${idx}`,
          solucao: item.solucao,
          area: item.area,
          statusImplementacao: 'Implementada Integralmente',
          aderenciaRotina: 'Alta',
          resultadoObservado: 'Rotina assimilada pelo cliente e incorporada às decisões semanais.',
          dificuldadesEncontradas: 'Adaptação inicial à disciplina de anotação diária.'
        });
      });
    }

    setData(prev => ({
      ...prev,
      problemasAvaliados: generatedProblemas,
      solucoesAvaliadas: generatedSolucoes
    }));

    setTimeout(() => setIsAutoFilling(false), 400);
  };

  // Calculations for KPI Dashboard and Theory Scores
  const calculatedMetrics = useMemo(() => {
    // 1. Problems Resolution Rate (Weight: 30%)
    const totalProb = data.problemasAvaliados.length;
    let sumProbScore = 0;
    let resolvidosCount = 0;
    let parcialCount = 0;
    let naoResolvidosCount = 0;

    data.problemasAvaliados.forEach(p => {
      if (p.status === 'Resolvido') {
        sumProbScore += p.percentualEficacia || 100;
        resolvidosCount++;
      } else if (p.status === 'Parcialmente Resolvido') {
        sumProbScore += (p.percentualEficacia || 50);
        parcialCount++;
      } else if (p.status === 'Em Andamento') {
        sumProbScore += (p.percentualEficacia || 40);
        parcialCount++;
      } else {
        naoResolvidosCount++;
      }
    });
    const probRate = totalProb > 0 ? Math.round(sumProbScore / totalProb) : 100;

    // 2. Solutions Implementation Rate (Weight: 25%)
    const totalSol = data.solucoesAvaliadas.length;
    let sumSolScore = 0;
    let solIntegralCount = 0;
    data.solucoesAvaliadas.forEach(s => {
      if (s.statusImplementacao === 'Implementada Integralmente') {
        sumSolScore += 100;
        solIntegralCount++;
      } else if (s.statusImplementacao === 'Implementada Parcialmente') {
        sumSolScore += 65;
      } else if (s.statusImplementacao === 'Em Implantação') {
        sumSolScore += 45;
      } else {
        sumSolScore += 0;
      }
    });
    const solRate = totalSol > 0 ? Math.round(sumSolScore / totalSol) : 100;

    // 3. Client Engagement Score (Weight: 15%)
    const eng = data.engajamentoCliente;
    const engSum = (eng.comprometimentoLideranca + eng.cumprimentoPrazosTarefas + eng.disponibilizacaoDadosDocumentos + eng.participacaoTreinamentos + eng.envolvimentoEquipe);
    const engRate = Math.round((engSum / 25) * 100);

    // 4. Financial Indicators Compliance (Weight: 20%)
    const totalInd = data.indicadoresFinanceiros.length;
    let indScoreSum = 0;
    let indConsolidadosCount = 0;
    data.indicadoresFinanceiros.forEach(ind => {
      if (ind.status === 'Ativo e Consolidado') {
        indScoreSum += 100;
        indConsolidadosCount++;
      } else if (ind.status === 'Em Implantação') {
        indScoreSum += 70;
      } else if (ind.status === 'Parcial / Requer Atenção') {
        indScoreSum += 40;
      } else {
        indScoreSum += 0;
      }
    });
    const indRate = totalInd > 0 ? Math.round(indScoreSum / totalInd) : 100;

    // 5. Satisfaction Score (Weight: 10%)
    const sat = data.pesquisaSatisfacao;
    const satAvg = ((sat.satisfacaoMetodologia + sat.satisfacaoConsultorDidatica + sat.satisfacaoConsultorPontualidade + sat.satisfacaoConsultorAtendimento) / 20) * 100;
    const npsPercent = (sat.npsConsultoria / 10) * 100;
    const satRate = Math.round((satAvg * 0.6) + (npsPercent * 0.4));

    // Consolidated Weighted Score (0-100%)
    const globalScore = Math.min(100, Math.max(0, Math.round(
      (probRate * 0.30) +
      (solRate * 0.25) +
      (engRate * 0.15) +
      (indRate * 0.20) +
      (satRate * 0.10)
    )));

    // Qualitative Level
    let nivelClassificacao = 'Excelente (Alto Impacto Transformador)';
    let badgeColor = 'bg-emerald-100 text-emerald-800 border-emerald-300';
    if (globalScore < 50) {
      nivelClassificacao = 'Crítico / Baixa Efetividade';
      badgeColor = 'bg-rose-100 text-rose-800 border-rose-300';
    } else if (globalScore < 75) {
      nivelClassificacao = 'Regular / Requer Continuidade Imediata';
      badgeColor = 'bg-amber-100 text-amber-800 border-amber-300';
    } else if (globalScore < 90) {
      nivelClassificacao = 'Muito Bom / Objetivos Atingidos';
      badgeColor = 'bg-blue-100 text-blue-800 border-blue-300';
    }

    return {
      probRate,
      resolvidosCount,
      parcialCount,
      naoResolvidosCount,
      totalProb,
      solRate,
      solIntegralCount,
      totalSol,
      engRate,
      indRate,
      indConsolidadosCount,
      totalInd,
      satRate,
      globalScore,
      nivelClassificacao,
      badgeColor,
      maturidadeDelta: (data.teoriaAdmin.maturidadeFinalPercent - data.teoriaAdmin.maturidadeInicialPercent)
    };
  }, [data]);

  // Save Analise de Resultados to Firestore and LocalStorage
  const handleSave = async () => {
    setIsSaving(true);
    setSaveSuccess(false);

    try {
      const payload: AnaliseResultadosData = {
        ...data,
        scoreGlobalEficacia: calculatedMetrics.globalScore,
        updatedAt: new Date().toISOString()
      };

      // 1. Save in localStorage
      const localKey = `analise_resultados_${selectedDiagnostico?.id}`;
      localStorage.setItem(localKey, JSON.stringify(payload));

      // Also update local_diagnosticos
      try {
        const savedDiags = localStorage.getItem('local_diagnosticos');
        if (savedDiags) {
          const list = JSON.parse(savedDiags);
          const updatedList = list.map((d: any) => {
            if (d.id === selectedDiagnostico?.id) {
              return { ...d, analiseResultados: payload };
            }
            return d;
          });
          localStorage.setItem('local_diagnosticos', JSON.stringify(updatedList));
        }
      } catch (e) {
        console.warn("Error updating local_diagnosticos with analiseResultados:", e);
      }

      // 2. Update React State
      if (selectedDiagnostico) {
        const updatedDiag = { ...selectedDiagnostico, analiseResultados: payload };
        setSelectedDiagnostico(updatedDiag);
        setDiagnosticos(prev => prev.map(d => d.id === selectedDiagnostico.id ? updatedDiag : d));
      }

      // 3. Save in Firestore if online
      if (auth.currentUser && selectedDiagnostico?.id) {
        try {
          // Save in dedicated collection
          const analiseDocRef = doc(db, 'analise_resultados', selectedDiagnostico.id);
          await setDoc(analiseDocRef, {
            ...payload,
            ownerId: auth.currentUser.uid,
            savedAt: serverTimestamp()
          }, { merge: true });

          // Also update the diagnostico document
          const diagDocRef = doc(db, 'diagnosticos', selectedDiagnostico.id);
          await setDoc(diagDocRef, {
            analiseResultados: payload,
            updatedAt: serverTimestamp()
          }, { merge: true });
        } catch (cloudErr) {
          console.warn("Cloud save warning (saved locally successfully):", cloudErr);
        }
      }

      playSuccessSound();
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3500);
    } catch (err) {
      console.error("Erro ao salvar análise de resultados:", err);
      alert("Houve um problema ao salvar a análise de resultados. Os dados foram mantidos localmente.");
    } finally {
      setIsSaving(false);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-8 pb-16">
      {/* Top Header with Breadcrumbs and Quick Actions */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 bg-white p-6 rounded-2xl border border-slate-100 shadow-sm print:hidden">
        <div className="flex items-center gap-4">
          <button
            onClick={() => setView('dashboard')}
            className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-50 rounded-xl transition-all border border-transparent hover:border-slate-200 cursor-pointer"
            title="Voltar ao Painel"
          >
            <ArrowLeft size={20} />
          </button>
          <div>
            <div className="flex items-center gap-2">
              <span className="inline-flex items-center gap-1.5 bg-emerald-50 text-emerald-700 font-bold px-2.5 py-0.5 rounded-full text-xs border border-emerald-200">
                <Award size={14} /> Análise de Eficácia & Resultados
              </span>
              {selectedDiagnostico?.dadosConsultoria?.codigoSgf && (
                <span className="text-xs bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-mono font-bold">
                  SGF: {selectedDiagnostico.dadosConsultoria.codigoSgf}
                </span>
              )}
            </div>
            <h1 className="text-2xl font-bold text-slate-800 mt-1">
              Avaliação de Resultados da Consultoria
            </h1>
            <p className="text-sm text-slate-500">
              Cliente: <strong className="text-slate-700">{selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || 'Empresa Selecionada'}</strong> • Período: {data.dataAnalise}
            </p>
          </div>
        </div>

        {/* Global Save & Print Actions */}
        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={autoPopulateFromDiagnostico}
            disabled={isAutoFilling}
            className="inline-flex items-center gap-2 px-3.5 py-2 text-xs font-bold text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 rounded-xl transition-all cursor-pointer"
            title="Reimportar problemas do diagnóstico e atividades do cronograma"
          >
            <Sparkles size={15} className={isAutoFilling ? 'animate-spin text-indigo-600' : 'text-indigo-600'} />
            {isAutoFilling ? 'Sincronizando...' : 'Sincronizar Diagnóstico'}
          </button>

          <button
            onClick={handlePrint}
            className="inline-flex items-center gap-2 px-4 py-2 text-xs font-bold text-slate-700 bg-white hover:bg-slate-50 border border-slate-200 rounded-xl shadow-xs transition-all cursor-pointer"
          >
            <Printer size={15} className="text-slate-500" />
            Imprimir Relatório
          </button>

          <button
            onClick={handleSave}
            disabled={isSaving}
            className="inline-flex items-center gap-2 px-5 py-2 text-xs font-bold text-white bg-emerald-600 hover:bg-emerald-700 active:scale-98 rounded-xl shadow-sm hover:shadow transition-all cursor-pointer disabled:opacity-50"
          >
            {isSaving ? (
              <RefreshCw size={15} className="animate-spin" />
            ) : saveSuccess ? (
              <Check size={15} className="text-white" />
            ) : (
              <Save size={15} />
            )}
            {isSaving ? 'Salvando...' : saveSuccess ? 'Salvo com Sucesso!' : 'Salvar Análise'}
          </button>
        </div>
      </div>

      {/* Navigation Sub-Tabs */}
      <div className="flex items-center gap-2 overflow-x-auto pb-2 scrollbar-none border-b border-slate-200/80 print:hidden">
        <button
          onClick={() => setActiveTab('painel')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'painel'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <BarChart3 size={15} />
          Painel Executivo
        </button>

        <button
          onClick={() => setActiveTab('problemas')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'problemas'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <Target size={15} className={activeTab === 'problemas' ? 'text-emerald-400' : 'text-emerald-600'} />
          1. Problemas Diagnosticados ({data.problemasAvaliados.length})
        </button>

        <button
          onClick={() => setActiveTab('solucoes')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'solucoes'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <Briefcase size={15} className={activeTab === 'solucoes' ? 'text-blue-400' : 'text-blue-600'} />
          2. Soluções & Ações ({data.solucoesAvaliadas.length})
        </button>

        <button
          onClick={() => setActiveTab('engajamento')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'engajamento'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <User size={15} className={activeTab === 'engajamento' ? 'text-indigo-400' : 'text-indigo-600'} />
          3. Cumprimento do Cliente
        </button>

        <button
          onClick={() => setActiveTab('financeiro')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'financeiro'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <DollarSign size={15} className={activeTab === 'financeiro' ? 'text-amber-400' : 'text-amber-600'} />
          4. Índices Financeiros ({calculatedMetrics.indConsolidadosCount}/{calculatedMetrics.totalInd})
        </button>

        <button
          onClick={() => setActiveTab('satisfacao')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'satisfacao'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <Star size={15} className={activeTab === 'satisfacao' ? 'text-yellow-400' : 'text-yellow-500'} />
          5. Satisfação (NPS: {data.pesquisaSatisfacao.npsConsultoria}/10)
        </button>

        <button
          onClick={() => setActiveTab('teoria')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'teoria'
              ? 'bg-slate-900 text-white shadow-sm'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-100'
          }`}
        >
          <BookOpen size={15} className={activeTab === 'teoria' ? 'text-purple-400' : 'text-purple-600'} />
          6. Teoria & Continuidade
        </button>

        <button
          onClick={() => setActiveTab('relatorio')}
          className={`flex items-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold whitespace-nowrap transition-all cursor-pointer ${
            activeTab === 'relatorio'
              ? 'bg-emerald-600 text-white shadow-sm'
              : 'bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200'
          }`}
        >
          <FileText size={15} />
          7. Relatório Formal com Assinaturas
        </button>
      </div>

      {/* ========================================================================= */}
      {/* TAB 0: PAINEL EXECUTIVO                                                    */}
      {/* ========================================================================= */}
      {activeTab === 'painel' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="space-y-6"
        >
          {/* Hero Score Card */}
          <div className="bg-gradient-to-br from-slate-900 via-slate-800 to-emerald-950 rounded-3xl p-8 text-white shadow-lg relative overflow-hidden">
            <div className="absolute right-0 top-0 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none" />
            <div className="relative z-10 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
              <div className="lg:col-span-8 space-y-4">
                <div className="inline-flex items-center gap-2 px-3 py-1 bg-emerald-500/20 text-emerald-300 rounded-full text-xs font-semibold border border-emerald-500/30">
                  <ShieldCheck size={14} /> Índice Global de Eficácia da Consultoria
                </div>
                <h2 className="text-3xl font-extrabold tracking-tight">
                  Score de Transformação & Eficácia: <span className="text-emerald-400">{calculatedMetrics.globalScore}%</span>
                </h2>
                <p className="text-slate-300 text-sm max-w-2xl leading-relaxed">
                  Avaliação consolidada integrando a resolução de gargalos diagnosticados, grau de implementação das soluções do cronograma, disciplina de acompanhamento dos índices financeiros e engajamento da liderança.
                </p>
                <div className="flex flex-wrap items-center gap-3 pt-2">
                  <span className={`px-3 py-1 rounded-lg text-xs font-bold border ${calculatedMetrics.badgeColor}`}>
                    {calculatedMetrics.nivelClassificacao}
                  </span>
                  <span className="text-xs text-slate-400">
                    Ciclo de Mudança: <strong className="text-white">{data.teoriaAdmin.faseCicloMudanca}</strong>
                  </span>
                </div>
              </div>

              {/* Radial / Donut / Metric summary */}
              <div className="lg:col-span-4 flex flex-col items-center justify-center p-6 bg-white/5 rounded-2xl border border-white/10 backdrop-blur-xs">
                <div className="relative flex items-center justify-center">
                  <svg className="w-36 h-36 transform -rotate-90">
                    <circle
                      cx="72"
                      cy="72"
                      r="58"
                      stroke="currentColor"
                      strokeWidth="10"
                      className="text-slate-700"
                      fill="transparent"
                    />
                    <circle
                      cx="72"
                      cy="72"
                      r="58"
                      stroke="currentColor"
                      strokeWidth="10"
                      strokeDasharray={364.4}
                      strokeDashoffset={364.4 - (364.4 * calculatedMetrics.globalScore) / 100}
                      strokeLinecap="round"
                      className="text-emerald-400 transition-all duration-1000 ease-out"
                      fill="transparent"
                    />
                  </svg>
                  <div className="absolute flex flex-col items-center">
                    <span className="text-3xl font-black text-white">{calculatedMetrics.globalScore}%</span>
                    <span className="text-[10px] text-slate-400 uppercase tracking-widest font-bold">Eficácia</span>
                  </div>
                </div>
                <div className="text-center mt-3">
                  <p className="text-xs font-bold text-emerald-300">
                    Evolução da Maturidade: +{calculatedMetrics.maturidadeDelta}%
                  </p>
                  <p className="text-[11px] text-slate-400">
                    Inicial: {data.teoriaAdmin.maturidadeInicialPercent}% ➔ Final: {data.teoriaAdmin.maturidadeFinalPercent}%
                  </p>
                </div>
              </div>
            </div>
          </div>

          {/* 5 Pillars Summary Grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
            {/* Pillar 1: Problemas */}
            <div
              onClick={() => setActiveTab('problemas')}
              className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs hover:border-emerald-200 transition-all cursor-pointer group"
            >
              <div className="flex items-center justify-between text-emerald-600 mb-3">
                <div className="p-2.5 bg-emerald-50 rounded-xl group-hover:scale-105 transition-transform">
                  <Target size={20} />
                </div>
                <span className="text-xs font-extrabold bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded-full">
                  {calculatedMetrics.probRate}%
                </span>
              </div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">1. Problemas</h3>
              <p className="text-lg font-black text-slate-800 mt-1">
                {calculatedMetrics.resolvidosCount} de {calculatedMetrics.totalProb}
              </p>
              <p className="text-xs text-slate-500 mt-0.5">Gargalos resolvidos</p>
              <div className="w-full bg-slate-100 h-1.5 rounded-full mt-3 overflow-hidden">
                <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${calculatedMetrics.probRate}%` }} />
              </div>
            </div>

            {/* Pillar 2: Soluções */}
            <div
              onClick={() => setActiveTab('solucoes')}
              className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs hover:border-blue-200 transition-all cursor-pointer group"
            >
              <div className="flex items-center justify-between text-blue-600 mb-3">
                <div className="p-2.5 bg-blue-50 rounded-xl group-hover:scale-105 transition-transform">
                  <Briefcase size={20} />
                </div>
                <span className="text-xs font-extrabold bg-blue-50 text-blue-700 px-2 py-0.5 rounded-full">
                  {calculatedMetrics.solRate}%
                </span>
              </div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">2. Soluções</h3>
              <p className="text-lg font-black text-slate-800 mt-1">
                {calculatedMetrics.solIntegralCount} de {calculatedMetrics.totalSol}
              </p>
              <p className="text-xs text-slate-500 mt-0.5">Ações implantadas</p>
              <div className="w-full bg-slate-100 h-1.5 rounded-full mt-3 overflow-hidden">
                <div className="bg-blue-500 h-full rounded-full" style={{ width: `${calculatedMetrics.solRate}%` }} />
              </div>
            </div>

            {/* Pillar 3: Engajamento */}
            <div
              onClick={() => setActiveTab('engajamento')}
              className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs hover:border-indigo-200 transition-all cursor-pointer group"
            >
              <div className="flex items-center justify-between text-indigo-600 mb-3">
                <div className="p-2.5 bg-indigo-50 rounded-xl group-hover:scale-105 transition-transform">
                  <User size={20} />
                </div>
                <span className="text-xs font-extrabold bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-full">
                  {calculatedMetrics.engRate}%
                </span>
              </div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">3. Cliente</h3>
              <p className="text-lg font-black text-slate-800 mt-1">
                {data.engajamentoCliente.comprometimentoLideranca}/5 Estrelas
              </p>
              <p className="text-xs text-slate-500 mt-0.5">Engajamento liderança</p>
              <div className="w-full bg-slate-100 h-1.5 rounded-full mt-3 overflow-hidden">
                <div className="bg-indigo-500 h-full rounded-full" style={{ width: `${calculatedMetrics.engRate}%` }} />
              </div>
            </div>

            {/* Pillar 4: Gestão Financeira */}
            <div
              onClick={() => setActiveTab('financeiro')}
              className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs hover:border-amber-200 transition-all cursor-pointer group"
            >
              <div className="flex items-center justify-between text-amber-600 mb-3">
                <div className="p-2.5 bg-amber-50 rounded-xl group-hover:scale-105 transition-transform">
                  <DollarSign size={20} />
                </div>
                <span className="text-xs font-extrabold bg-amber-50 text-amber-700 px-2 py-0.5 rounded-full">
                  {calculatedMetrics.indRate}%
                </span>
              </div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">4. Finanças</h3>
              <p className="text-lg font-black text-slate-800 mt-1">
                {calculatedMetrics.indConsolidadosCount}/{calculatedMetrics.totalInd} Rotinas
              </p>
              <p className="text-xs text-slate-500 mt-0.5">Índices consolidados</p>
              <div className="w-full bg-slate-100 h-1.5 rounded-full mt-3 overflow-hidden">
                <div className="bg-amber-500 h-full rounded-full" style={{ width: `${calculatedMetrics.indRate}%` }} />
              </div>
            </div>

            {/* Pillar 5: Satisfação NPS */}
            <div
              onClick={() => setActiveTab('satisfacao')}
              className="bg-white p-5 rounded-2xl border border-slate-100 shadow-xs hover:border-yellow-200 transition-all cursor-pointer group"
            >
              <div className="flex items-center justify-between text-yellow-600 mb-3">
                <div className="p-2.5 bg-yellow-50 rounded-xl group-hover:scale-105 transition-transform">
                  <Star size={20} />
                </div>
                <span className="text-xs font-extrabold bg-yellow-50 text-yellow-800 px-2 py-0.5 rounded-full">
                  NPS {data.pesquisaSatisfacao.npsConsultoria}/10
                </span>
              </div>
              <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">5. Satisfação</h3>
              <p className="text-lg font-black text-slate-800 mt-1">
                {data.pesquisaSatisfacao.npsConsultoria >= 9 ? 'Promotor' : 'Neutro'}
              </p>
              <p className="text-xs text-slate-500 mt-0.5">Nota do consultor: 5/5</p>
              <div className="w-full bg-slate-100 h-1.5 rounded-full mt-3 overflow-hidden">
                <div className="bg-yellow-500 h-full rounded-full" style={{ width: `${(data.pesquisaSatisfacao.npsConsultoria / 10) * 100}%` }} />
              </div>
            </div>
          </div>

          {/* Quick Summary Grid */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Left: Parecer Técnico */}
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-4">
              <div className="flex items-center justify-between">
                <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                  <FileText size={18} className="text-emerald-600" />
                  Parecer Técnico do Consultor
                </h3>
                <button
                  onClick={() => setActiveTab('teoria')}
                  className="text-xs font-bold text-emerald-600 hover:underline flex items-center gap-1 cursor-pointer"
                >
                  Editar Parecer <ChevronRight size={14} />
                </button>
              </div>
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-sm text-slate-700 leading-relaxed italic">
                "{data.teoriaAdmin.parecerTecnicoConsultor || 'Parecer não preenchido.'}"
              </div>
              <div className="space-y-2">
                <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">Principais Ganhos Tangíveis:</h4>
                <div className="bg-emerald-50/50 p-3.5 rounded-xl border border-emerald-100 text-xs text-emerald-900 whitespace-pre-line font-medium leading-relaxed">
                  {data.teoriaAdmin.principaisGanhosObtidos}
                </div>
              </div>
            </div>

            {/* Right: Plano de Continuidade */}
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-4">
              <div className="flex items-center justify-between">
                <h3 className="text-base font-bold text-slate-800 flex items-center gap-2">
                  <Clock size={18} className="text-blue-600" />
                  Governança & Plano de Continuidade (Pós-Consultoria)
                </h3>
                <button
                  onClick={() => setActiveTab('teoria')}
                  className="text-xs font-bold text-blue-600 hover:underline flex items-center gap-1 cursor-pointer"
                >
                  Ver Cronograma <ChevronRight size={14} />
                </button>
              </div>
              <div className="space-y-3 text-xs">
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 flex items-start gap-3">
                  <span className="px-2 py-0.5 bg-blue-100 text-blue-800 rounded font-bold">30 Dias</span>
                  <p className="text-slate-700 flex-1">{data.teoriaAdmin.planoContinuidade30Dias}</p>
                </div>
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 flex items-start gap-3">
                  <span className="px-2 py-0.5 bg-indigo-100 text-indigo-800 rounded font-bold">60 Dias</span>
                  <p className="text-slate-700 flex-1">{data.teoriaAdmin.planoContinuidade60Dias}</p>
                </div>
                <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 flex items-start gap-3">
                  <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded font-bold">90 Dias</span>
                  <p className="text-slate-700 flex-1">{data.teoriaAdmin.planoContinuidade90Dias}</p>
                </div>
              </div>
            </div>
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 1: PROBLEMAS DIAGNOSTICADOS                                            */}
      {/* ========================================================================= */}
      {activeTab === 'problemas' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                <Target size={20} className="text-emerald-600" />
                1. Resolução dos Problemas Diagnosticados
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Avalie o status de resolução para cada problema identificado no diagnóstico inicial do cliente.
              </p>
            </div>
            <button
              onClick={() => {
                const novo: ProblemaAvaliado = {
                  id: `prob_custom_${Date.now()}`,
                  problema: '',
                  area: 'Gestão Financeira',
                  status: 'Resolvido',
                  percentualEficacia: 100,
                  impactoGerado: '',
                  evidencias: ''
                };
                setData(prev => ({ ...prev, problemasAvaliados: [...prev.problemasAvaliados, novo] }));
              }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-bold cursor-pointer transition-all"
            >
              + Adicionar Problema
            </button>
          </div>

          <div className="space-y-4">
            {data.problemasAvaliados.map((item, idx) => (
              <div key={item.id} className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="flex-1">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                      Problema #{idx + 1}
                    </label>
                    <input
                      type="text"
                      value={item.problema}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, problema: val } : p)
                        }));
                      }}
                      placeholder="Descrição do problema diagnosticado..."
                      className="w-full mt-1 px-3 py-1.5 text-sm font-semibold text-slate-800 bg-white border border-slate-200 rounded-lg focus:ring-2 focus:ring-emerald-500"
                    />
                  </div>

                  <div className="w-full md:w-48">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Área</label>
                    <input
                      type="text"
                      value={item.area}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, area: val } : p)
                        }));
                      }}
                      className="w-full mt-1 px-3 py-1.5 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="w-full md:w-48">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Status de Resolução</label>
                    <select
                      value={item.status}
                      onChange={e => {
                        const val = e.target.value as any;
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, status: val } : p)
                        }));
                      }}
                      className="w-full mt-1 px-3 py-1.5 text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg"
                    >
                      <option value="Resolvido">✅ Resolvido</option>
                      <option value="Parcialmente Resolvido">⚠️ Parcialmente Resolvido</option>
                      <option value="Em Andamento">⏳ Em Andamento</option>
                      <option value="Não Resolvido">❌ Não Resolvido</option>
                    </select>
                  </div>
                </div>

                {/* Efficiency slider and evidence notes */}
                <div className="grid grid-cols-1 md:grid-cols-12 gap-3 items-center pt-1 border-t border-slate-200/60">
                  <div className="md:col-span-4 flex items-center gap-3">
                    <label className="text-xs font-bold text-slate-600 whitespace-nowrap">
                      Eficácia: <span className="text-emerald-600">{item.percentualEficacia}%</span>
                    </label>
                    <input
                      type="range"
                      min="0"
                      max="100"
                      step="5"
                      value={item.percentualEficacia}
                      onChange={e => {
                        const val = Number(e.target.value);
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, percentualEficacia: val } : p)
                        }));
                      }}
                      className="w-full accent-emerald-600"
                    />
                  </div>

                  <div className="md:col-span-4">
                    <input
                      type="text"
                      value={item.impactoGerado}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, impactoGerado: val } : p)
                        }));
                      }}
                      placeholder="Impacto gerado na empresa (ex: separação total das contas)..."
                      className="w-full px-3 py-1.5 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="md:col-span-3">
                    <input
                      type="text"
                      value={item.evidencias}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.map(p => p.id === item.id ? { ...p, evidencias: val } : p)
                        }));
                      }}
                      placeholder="Evidências / Comprovações..."
                      className="w-full px-3 py-1.5 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="md:col-span-1 flex justify-end">
                    <button
                      onClick={() => {
                        setData(prev => ({
                          ...prev,
                          problemasAvaliados: prev.problemasAvaliados.filter(p => p.id !== item.id)
                        }));
                      }}
                      className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-lg cursor-pointer transition-all"
                      title="Excluir problema"
                    >
                      <XCircle size={16} />
                    </button>
                  </div>
                </div>
              </div>
            ))}

            {data.problemasAvaliados.length === 0 && (
              <div className="text-center py-8 text-slate-400">
                Nenhum problema cadastrado. Clique em "Sincronizar Diagnóstico" para puxar as perguntas do diagnóstico.
              </div>
            )}
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 2: SOLUÇÕES E AÇÕES PROPOSTAS                                          */}
      {/* ========================================================================= */}
      {activeTab === 'solucoes' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-slate-100 pb-4">
            <div>
              <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                <Briefcase size={20} className="text-blue-600" />
                2. Implementação das Soluções e Ações Propostas
              </h2>
              <p className="text-xs text-slate-500 mt-1">
                Verifique se as soluções sugeridas no cronograma de 30h e no plano de ação foram executadas e incorporadas à rotina.
              </p>
            </div>
            <button
              onClick={() => {
                const nova: SolucaoAvaliada = {
                  id: `sol_custom_${Date.now()}`,
                  solucao: '',
                  area: 'Gestão Financeira',
                  statusImplementacao: 'Implementada Integralmente',
                  aderenciaRotina: 'Alta',
                  resultadoObservado: '',
                  dificuldadesEncontradas: ''
                };
                setData(prev => ({ ...prev, solucoesAvaliadas: [...prev.solucoesAvaliadas, nova] }));
              }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-bold cursor-pointer transition-all"
            >
              + Adicionar Solução
            </button>
          </div>

          <div className="space-y-4">
            {data.solucoesAvaliadas.map((item, idx) => (
              <div key={item.id} className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="flex-1">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">
                      Solução / Ação #{idx + 1}
                    </label>
                    <input
                      type="text"
                      value={item.solucao}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.map(s => s.id === item.id ? { ...s, solucao: val } : s)
                        }));
                      }}
                      placeholder="Descrição da solução indicada..."
                      className="w-full mt-1 px-3 py-1.5 text-sm font-semibold text-slate-800 bg-white border border-slate-200 rounded-lg focus:ring-2 focus:ring-blue-500"
                    />
                  </div>

                  <div className="w-full md:w-56">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Status de Implantação</label>
                    <select
                      value={item.statusImplementacao}
                      onChange={e => {
                        const val = e.target.value as any;
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.map(s => s.id === item.id ? { ...s, statusImplementacao: val } : s)
                        }));
                      }}
                      className="w-full mt-1 px-3 py-1.5 text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg"
                    >
                      <option value="Implementada Integralmente">🟢 Implementada Integralmente</option>
                      <option value="Implementada Parcialmente">🟡 Implementada Parcialmente</option>
                      <option value="Em Implantação">🔵 Em Implantação</option>
                      <option value="Não Implementada">🔴 Não Implementada</option>
                    </select>
                  </div>

                  <div className="w-full md:w-36">
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider">Aderência à Rotina</label>
                    <select
                      value={item.aderenciaRotina}
                      onChange={e => {
                        const val = e.target.value as any;
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.map(s => s.id === item.id ? { ...s, aderenciaRotina: val } : s)
                        }));
                      }}
                      className="w-full mt-1 px-3 py-1.5 text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg"
                    >
                      <option value="Alta">Alta</option>
                      <option value="Média">Média</option>
                      <option value="Baixa">Baixa</option>
                    </select>
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-12 gap-3 pt-1 border-t border-slate-200/60">
                  <div className="md:col-span-6">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Resultado Prático Observado</label>
                    <input
                      type="text"
                      value={item.resultadoObservado}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.map(s => s.id === item.id ? { ...s, resultadoObservado: val } : s)
                        }));
                      }}
                      placeholder="Ex: Planilha de custos preenchida e conferida a cada safra..."
                      className="w-full mt-0.5 px-3 py-1 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="md:col-span-5">
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Dificuldades / Desafios Encontrados</label>
                    <input
                      type="text"
                      value={item.dificuldadesEncontradas}
                      onChange={e => {
                        const val = e.target.value;
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.map(s => s.id === item.id ? { ...s, dificuldadesEncontradas: val } : s)
                        }));
                      }}
                      placeholder="Ex: Falta de tempo inicial para digitação dos dados..."
                      className="w-full mt-0.5 px-3 py-1 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                    />
                  </div>

                  <div className="md:col-span-1 flex items-end justify-end pb-1">
                    <button
                      onClick={() => {
                        setData(prev => ({
                          ...prev,
                          solucoesAvaliadas: prev.solucoesAvaliadas.filter(s => s.id !== item.id)
                        }));
                      }}
                      className="p-1.5 text-rose-500 hover:bg-rose-50 rounded-lg cursor-pointer transition-all"
                      title="Excluir solução"
                    >
                      <XCircle size={16} />
                    </button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 3: CUMPRIMENTO DAS ATIVIDADES PELO CLIENTE (ENGAJAMENTO)              */}
      {/* ========================================================================= */}
      {activeTab === 'engajamento' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="border-b border-slate-100 pb-4">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <User size={20} className="text-indigo-600" />
              3. Cumprimento das Atividades Orientadas e Engajamento do Cliente
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              Avaliação do comportamento, assiduidade, cumprimento dos deveres de casa e comprometimento da liderança no processo de consultoria.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* 1. Comprometimento da Liderança */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-700">1. Comprometimento da Liderança / Proprietário</label>
                <span className="text-xs font-black text-indigo-600">{data.engajamentoCliente.comprometimentoLideranca} / 5</span>
              </div>
              <p className="text-[11px] text-slate-500">Priorização da consultoria, presença ativa nas reuniões e determinação para implementar as mudanças.</p>
              <div className="flex items-center gap-2 pt-1">
                {[1, 2, 3, 4, 5].map(star => (
                  <button
                    key={star}
                    onClick={() => setData(prev => ({
                      ...prev,
                      engajamentoCliente: { ...prev.engajamentoCliente, comprometimentoLideranca: star }
                    }))}
                    className={`p-1.5 rounded-lg transition-all cursor-pointer ${
                      star <= data.engajamentoCliente.comprometimentoLideranca
                        ? 'text-amber-500 bg-amber-50'
                        : 'text-slate-300 hover:text-slate-400'
                    }`}
                  >
                    <Star size={20} className={star <= data.engajamentoCliente.comprometimentoLideranca ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* 2. Cumprimento de Prazos e Tarefas */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-700">2. Cumprimento de Prazos e "Deveres de Casa"</label>
                <span className="text-xs font-black text-indigo-600">{data.engajamentoCliente.cumprimentoPrazosTarefas} / 5</span>
              </div>
              <p className="text-[11px] text-slate-500">Execução das atividades acordadas entre uma sessão e outra (ex: preencher anotações de caixa).</p>
              <div className="flex items-center gap-2 pt-1">
                {[1, 2, 3, 4, 5].map(star => (
                  <button
                    key={star}
                    onClick={() => setData(prev => ({
                      ...prev,
                      engajamentoCliente: { ...prev.engajamentoCliente, cumprimentoPrazosTarefas: star }
                    }))}
                    className={`p-1.5 rounded-lg transition-all cursor-pointer ${
                      star <= data.engajamentoCliente.cumprimentoPrazosTarefas
                        ? 'text-amber-500 bg-amber-50'
                        : 'text-slate-300 hover:text-slate-400'
                    }`}
                  >
                    <Star size={20} className={star <= data.engajamentoCliente.cumprimentoPrazosTarefas ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* 3. Disponibilização de Dados e Documentos */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-700">3. Transparência & Disponibilização de Dados</label>
                <span className="text-xs font-black text-indigo-600">{data.engajamentoCliente.disponibilizacaoDadosDocumentos} / 5</span>
              </div>
              <p className="text-[11px] text-slate-500">Abertura e tempestividade para apresentar extratos, notas fiscais, custos e informações gerenciais.</p>
              <div className="flex items-center gap-2 pt-1">
                {[1, 2, 3, 4, 5].map(star => (
                  <button
                    key={star}
                    onClick={() => setData(prev => ({
                      ...prev,
                      engajamentoCliente: { ...prev.engajamentoCliente, disponibilizacaoDadosDocumentos: star }
                    }))}
                    className={`p-1.5 rounded-lg transition-all cursor-pointer ${
                      star <= data.engajamentoCliente.disponibilizacaoDadosDocumentos
                        ? 'text-amber-500 bg-amber-50'
                        : 'text-slate-300 hover:text-slate-400'
                    }`}
                  >
                    <Star size={20} className={star <= data.engajamentoCliente.disponibilizacaoDadosDocumentos ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* 4. Participação nos Treinamentos */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-slate-700">4. Absorção das Orientações e Capacitações</label>
                <span className="text-xs font-black text-indigo-600">{data.engajamentoCliente.participacaoTreinamentos} / 5</span>
              </div>
              <p className="text-[11px] text-slate-500">Capacidade de assimilação dos conceitos práticos de finanças, precificação e controles.</p>
              <div className="flex items-center gap-2 pt-1">
                {[1, 2, 3, 4, 5].map(star => (
                  <button
                    key={star}
                    onClick={() => setData(prev => ({
                      ...prev,
                      engajamentoCliente: { ...prev.engajamentoCliente, participacaoTreinamentos: star }
                    }))}
                    className={`p-1.5 rounded-lg transition-all cursor-pointer ${
                      star <= data.engajamentoCliente.participacaoTreinamentos
                        ? 'text-amber-500 bg-amber-50'
                        : 'text-slate-300 hover:text-slate-400'
                    }`}
                  >
                    <Star size={20} className={star <= data.engajamentoCliente.participacaoTreinamentos ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-xs font-bold text-slate-700">
              Observações do Consultor sobre a Atitude e Engajamento do Cliente:
            </label>
            <textarea
              rows={3}
              value={data.engajamentoCliente.observacoesEngajamento}
              onChange={e => {
                const val = e.target.value;
                setData(prev => ({
                  ...prev,
                  engajamentoCliente: { ...prev.engajamentoCliente, observacoesEngajamento: val }
                }));
              }}
              placeholder="Descreva a receptividade do produtor/gestor, pontualidade e postura durante a consultoria..."
              className="w-full p-3 text-xs text-slate-800 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500"
            />
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 4: ACOMPANHAMENTO DOS ÍNDICES DE GESTÃO FINANCEIRA                      */}
      {/* ========================================================================= */}
      {activeTab === 'financeiro' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="border-b border-slate-100 pb-4">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <DollarSign size={20} className="text-amber-600" />
              4. Acompanhamento dos Índices & Boas Práticas de Gestão Financeira
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              Avaliação do domínio técnico e implantação das rotinas de controle financeiro ensinadas durante a consultoria.
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {data.indicadoresFinanceiros.map((ind) => (
              <div key={ind.id} className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
                <div className="flex items-start justify-between gap-2">
                  <div>
                    <h3 className="text-sm font-bold text-slate-800">{ind.nome}</h3>
                    <p className="text-[11px] text-slate-500 leading-tight mt-0.5">{ind.descricao}</p>
                  </div>
                  <span className={`px-2 py-0.5 rounded text-[10px] font-bold whitespace-nowrap ${
                    ind.status === 'Ativo e Consolidado'
                      ? 'bg-emerald-100 text-emerald-800 border border-emerald-200'
                      : ind.status === 'Em Implantação'
                      ? 'bg-blue-100 text-blue-800 border border-blue-200'
                      : 'bg-amber-100 text-amber-800 border border-amber-200'
                  }`}>
                    {ind.status}
                  </span>
                </div>

                <div className="grid grid-cols-2 gap-2 pt-1 border-t border-slate-200/60">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Status de Execução</label>
                    <select
                      value={ind.status}
                      onChange={e => {
                        const val = e.target.value as any;
                        setData(prev => ({
                          ...prev,
                          indicadoresFinanceiros: prev.indicadoresFinanceiros.map(i => i.id === ind.id ? { ...i, status: val } : i)
                        }));
                      }}
                      className="w-full mt-0.5 px-2 py-1 text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg"
                    >
                      <option value="Ativo e Consolidado">🟢 Ativo e Consolidado</option>
                      <option value="Em Implantação">🔵 Em Implantação</option>
                      <option value="Parcial / Requer Atenção">🟡 Parcial / Requer Atenção</option>
                      <option value="Não Acompanhado">🔴 Não Acompanhado</option>
                    </select>
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase">Frequência da Rotina</label>
                    <select
                      value={ind.frequencia}
                      onChange={e => {
                        const val = e.target.value as any;
                        setData(prev => ({
                          ...prev,
                          indicadoresFinanceiros: prev.indicadoresFinanceiros.map(i => i.id === ind.id ? { ...i, frequencia: val } : i)
                        }));
                      }}
                      className="w-full mt-0.5 px-2 py-1 text-xs font-bold text-slate-800 bg-white border border-slate-200 rounded-lg"
                    >
                      <option value="Diária">Diária</option>
                      <option value="Semanal">Semanal</option>
                      <option value="Mensal">Mensal</option>
                      <option value="Não Realiza">Não Realiza</option>
                    </select>
                  </div>
                </div>

                <div>
                  <input
                    type="text"
                    value={ind.observacao}
                    onChange={e => {
                      const val = e.target.value;
                      setData(prev => ({
                        ...prev,
                        indicadoresFinanceiros: prev.indicadoresFinanceiros.map(i => i.id === ind.id ? { ...i, observacao: val } : i)
                      }));
                    }}
                    placeholder="Observação / ferramenta utilizada..."
                    className="w-full px-2.5 py-1 text-xs text-slate-700 bg-white border border-slate-200 rounded-lg"
                  />
                </div>
              </div>
            ))}
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 5: PESQUISA DE SATISFAÇÃO COM CONSULTORIA & CONSULTOR                  */}
      {/* ========================================================================= */}
      {activeTab === 'satisfacao' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="border-b border-slate-100 pb-4">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <Star size={20} className="text-yellow-500" />
              5. Pesquisa de Satisfação com a Consultoria e com o Consultor
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              Avaliação formal do cliente sobre a qualidade técnica, pontualidade, didática e atendimento do consultor credenciado.
            </p>
          </div>

          {/* NPS Selector */}
          <div className="p-6 bg-gradient-to-r from-slate-900 to-indigo-950 rounded-2xl text-white space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <h3 className="text-base font-bold flex items-center gap-2 text-yellow-400">
                  <Award size={18} /> Net Promoter Score (NPS da Consultoria)
                </h3>
                <p className="text-xs text-slate-300 mt-0.5">
                  "Em uma escala de 0 a 10, qual a probabilidade de você recomendar esta consultoria do Sebrae para outro empresário?"
                </p>
              </div>
              <span className="text-2xl font-black text-emerald-400 bg-white/10 px-4 py-1 rounded-xl">
                {data.pesquisaSatisfacao.npsConsultoria} / 10
              </span>
            </div>

            <div className="grid grid-cols-11 gap-1 pt-2">
              {[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(score => (
                <button
                  key={score}
                  onClick={() => setData(prev => ({
                    ...prev,
                    pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, npsConsultoria: score }
                  }))}
                  className={`py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                    data.pesquisaSatisfacao.npsConsultoria === score
                      ? score >= 9
                        ? 'bg-emerald-500 text-white ring-2 ring-emerald-300 scale-105'
                        : score >= 7
                        ? 'bg-amber-500 text-white ring-2 ring-amber-300 scale-105'
                        : 'bg-rose-500 text-white ring-2 ring-rose-300 scale-105'
                      : 'bg-white/10 hover:bg-white/20 text-slate-200'
                  }`}
                >
                  {score}
                </button>
              ))}
            </div>
            <div className="flex justify-between text-[10px] text-slate-400 uppercase font-bold pt-1">
              <span>0 = Nada Provável</span>
              <span>10 = Altamente Provável (Promotor)</span>
            </div>
          </div>

          {/* Detailed Criteria */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {/* Metodologia */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex justify-between items-center">
                <label className="text-xs font-bold text-slate-700">Metodologia e Clareza das Ferramentas</label>
                <span className="text-xs font-bold text-indigo-600">{data.pesquisaSatisfacao.satisfacaoMetodologia}/5</span>
              </div>
              <div className="flex items-center gap-1.5">
                {[1, 2, 3, 4, 5].map(s => (
                  <button
                    key={s}
                    onClick={() => setData(prev => ({
                      ...prev,
                      pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, satisfacaoMetodologia: s }
                    }))}
                    className={`p-1 rounded cursor-pointer ${s <= data.pesquisaSatisfacao.satisfacaoMetodologia ? 'text-amber-500' : 'text-slate-300'}`}
                  >
                    <Star size={18} className={s <= data.pesquisaSatisfacao.satisfacaoMetodologia ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* Didática do Consultor */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex justify-between items-center">
                <label className="text-xs font-bold text-slate-700">Didática, Clareza e Domínio do Consultor</label>
                <span className="text-xs font-bold text-indigo-600">{data.pesquisaSatisfacao.satisfacaoConsultorDidatica}/5</span>
              </div>
              <div className="flex items-center gap-1.5">
                {[1, 2, 3, 4, 5].map(s => (
                  <button
                    key={s}
                    onClick={() => setData(prev => ({
                      ...prev,
                      pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, satisfacaoConsultorDidatica: s }
                    }))}
                    className={`p-1 rounded cursor-pointer ${s <= data.pesquisaSatisfacao.satisfacaoConsultorDidatica ? 'text-amber-500' : 'text-slate-300'}`}
                  >
                    <Star size={18} className={s <= data.pesquisaSatisfacao.satisfacaoConsultorDidatica ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* Pontualidade */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex justify-between items-center">
                <label className="text-xs font-bold text-slate-700">Pontualidade e Carga Horária (30h)</label>
                <span className="text-xs font-bold text-indigo-600">{data.pesquisaSatisfacao.satisfacaoConsultorPontualidade}/5</span>
              </div>
              <div className="flex items-center gap-1.5">
                {[1, 2, 3, 4, 5].map(s => (
                  <button
                    key={s}
                    onClick={() => setData(prev => ({
                      ...prev,
                      pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, satisfacaoConsultorPontualidade: s }
                    }))}
                    className={`p-1 rounded cursor-pointer ${s <= data.pesquisaSatisfacao.satisfacaoConsultorPontualidade ? 'text-amber-500' : 'text-slate-300'}`}
                  >
                    <Star size={18} className={s <= data.pesquisaSatisfacao.satisfacaoConsultorPontualidade ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>

            {/* Atendimento Geral */}
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <div className="flex justify-between items-center">
                <label className="text-xs font-bold text-slate-700">Postura Ética e Atendimento</label>
                <span className="text-xs font-bold text-indigo-600">{data.pesquisaSatisfacao.satisfacaoConsultorAtendimento}/5</span>
              </div>
              <div className="flex items-center gap-1.5">
                {[1, 2, 3, 4, 5].map(s => (
                  <button
                    key={s}
                    onClick={() => setData(prev => ({
                      ...prev,
                      pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, satisfacaoConsultorAtendimento: s }
                    }))}
                    className={`p-1 rounded cursor-pointer ${s <= data.pesquisaSatisfacao.satisfacaoConsultorAtendimento ? 'text-amber-500' : 'text-slate-300'}`}
                  >
                    <Star size={18} className={s <= data.pesquisaSatisfacao.satisfacaoConsultorAtendimento ? 'fill-amber-500' : ''} />
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div className="space-y-3">
            <div>
              <label className="text-xs font-bold text-slate-700">Depoimento / Declaração Formal do Cliente:</label>
              <textarea
                rows={2}
                value={data.pesquisaSatisfacao.depoimentoCliente}
                onChange={e => {
                  const val = e.target.value;
                  setData(prev => ({
                    ...prev,
                    pesquisaSatisfacao: { ...prev.pesquisaSatisfacao, depoimentoCliente: val }
                  }));
                }}
                className="w-full mt-1 p-2.5 text-xs text-slate-800 bg-slate-50 border border-slate-200 rounded-xl"
                placeholder="Depoimento do empresário sobre os resultados alcançados..."
              />
            </div>
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 6: TEORIA DA ADMINISTRAÇÃO & PLANO DE CONTINUIDADE                      */}
      {/* ========================================================================= */}
      {activeTab === 'teoria' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-6 rounded-2xl border border-slate-100 shadow-xs space-y-6"
        >
          <div className="border-b border-slate-100 pb-4">
            <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
              <BookOpen size={20} className="text-purple-600" />
              6. Teoria da Administração, Gestão da Mudança & Plano de Continuidade
            </h2>
            <p className="text-xs text-slate-500 mt-1">
              Fundamentação teórica baseada no Modelo de Mudança Organizacional de Kurt Lewin e Ciclo PDCA, garantindo que os resultados se perpetuem após o término da consultoria.
            </p>
          </div>

          {/* Kurt Lewin Change Model Stage */}
          <div className="p-5 bg-purple-50 rounded-2xl border border-purple-100 space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div>
                <h3 className="text-sm font-bold text-purple-950">Estágio do Processo de Mudança (Kurt Lewin / PDCA)</h3>
                <p className="text-xs text-purple-800 mt-0.5">Identifique o grau de consolidação cultural da gestão na empresa:</p>
              </div>
              <select
                value={data.teoriaAdmin.faseCicloMudanca}
                onChange={e => {
                  const val = e.target.value as any;
                  setData(prev => ({
                    ...prev,
                    teoriaAdmin: { ...prev.teoriaAdmin, faseCicloMudanca: val }
                  }));
                }}
                className="px-3 py-1.5 text-xs font-bold text-purple-900 bg-white border border-purple-200 rounded-xl"
              >
                <option value="Descongelamento (Sensibilização)">1. Descongelamento (Sensibilização & Quebra de Maus Hábitos)</option>
                <option value="Movimento / Mudança (Implantação de Ferramentas)">2. Movimento / Mudança (Implantação de Ferramentas)</option>
                <option value="Recongelamento (Consolidação da Rotina)">3. Recongelamento (Consolidação da Nova Rotina Gerencial)</option>
              </select>
            </div>
          </div>

          {/* Parecer Técnico Conclusivo */}
          <div className="space-y-2">
            <label className="text-xs font-bold text-slate-700">
              Parecer Técnico Conclusivo do Consultor:
            </label>
            <textarea
              rows={4}
              value={data.teoriaAdmin.parecerTecnicoConsultor}
              onChange={e => {
                const val = e.target.value;
                setData(prev => ({
                  ...prev,
                  teoriaAdmin: { ...prev.teoriaAdmin, parecerTecnicoConsultor: val }
                }));
              }}
              className="w-full p-3 text-xs text-slate-800 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-purple-500"
              placeholder="Parecer final do consultor sobre os resultados, aprendizado do cliente e sustentabilidade..."
            />
          </div>

          {/* Continuidade 30, 60, 90 dias */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <span className="text-[10px] font-bold text-blue-600 uppercase tracking-widest">Pós-Consultoria (30 Dias)</span>
              <textarea
                rows={3}
                value={data.teoriaAdmin.planoContinuidade30Dias}
                onChange={e => {
                  const val = e.target.value;
                  setData(prev => ({
                    ...prev,
                    teoriaAdmin: { ...prev.teoriaAdmin, planoContinuidade30Dias: val }
                  }));
                }}
                className="w-full p-2 text-xs text-slate-800 bg-white border border-slate-200 rounded-lg"
                placeholder="Ações prioritárias para os primeiros 30 dias..."
              />
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <span className="text-[10px] font-bold text-indigo-600 uppercase tracking-widest">Pós-Consultoria (60 Dias)</span>
              <textarea
                rows={3}
                value={data.teoriaAdmin.planoContinuidade60Dias}
                onChange={e => {
                  const val = e.target.value;
                  setData(prev => ({
                    ...prev,
                    teoriaAdmin: { ...prev.teoriaAdmin, planoContinuidade60Dias: val }
                  }));
                }}
                className="w-full p-2 text-xs text-slate-800 bg-white border border-slate-200 rounded-lg"
                placeholder="Ações prioritárias para 60 dias..."
              />
            </div>

            <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-2">
              <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-widest">Pós-Consultoria (90 Dias)</span>
              <textarea
                rows={3}
                value={data.teoriaAdmin.planoContinuidade90Dias}
                onChange={e => {
                  const val = e.target.value;
                  setData(prev => ({
                    ...prev,
                    teoriaAdmin: { ...prev.teoriaAdmin, planoContinuidade90Dias: val }
                  }));
                }}
                className="w-full p-2 text-xs text-slate-800 bg-white border border-slate-200 rounded-lg"
                placeholder="Ações para 90 dias e revisão anual..."
              />
            </div>
          </div>
        </motion.div>
      )}

      {/* ========================================================================= */}
      {/* TAB 7: RELATÓRIO FORMAL DE ANÁLISE DE RESULTADOS COM ASSINATURAS          */}
      {/* ========================================================================= */}
      {activeTab === 'relatorio' && (
        <motion.div
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          className="space-y-6"
        >
          {/* Action Bar for Report */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-white p-4 rounded-2xl border border-slate-100 shadow-xs print:hidden">
            <div className="flex items-center gap-2">
              <span className="text-xs font-bold text-slate-500">Logo no Relatório:</span>
              <div className="flex items-center bg-slate-100 p-1 rounded-xl">
                <button
                  type="button"
                  onClick={() => setLogoChoice && setLogoChoice('sebrae')}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all ${
                    logoChoice === 'sebrae' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Sebrae
                </button>
                <button
                  type="button"
                  onClick={() => setLogoChoice && setLogoChoice('consultora')}
                  className={`px-3 py-1 text-xs font-bold rounded-lg transition-all ${
                    logoChoice === 'consultora' ? 'bg-white text-emerald-700 shadow-xs' : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  Consultora
                </button>
              </div>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handlePrint}
                className="inline-flex items-center gap-2 px-5 py-2.5 text-xs font-bold text-white bg-slate-900 hover:bg-slate-800 rounded-xl shadow-sm transition-all cursor-pointer"
              >
                <Printer size={16} /> Imprimir / Salvar PDF (A4)
              </button>
            </div>
          </div>

          {/* Printable Report Container */}
          <div
            ref={reportRef}
            className="bg-white p-8 md:p-12 rounded-3xl border border-slate-200 shadow-sm print:p-0 print:border-none print:shadow-none max-w-5xl mx-auto font-sans text-slate-800 space-y-8"
          >
            {/* Header with Logo and Formal Title */}
            <div className="flex items-center justify-between border-b-2 border-slate-800 pb-6">
              <div className="flex items-center gap-4">
                {logoChoice === 'consultora' && customConsultoraLogo ? (
                  <img src={customConsultoraLogo} alt="Logo Consultora" className="h-16 object-contain" referrerPolicy="no-referrer" />
                ) : (
                  <div className="flex items-center gap-2">
                    <div className="w-12 h-12 bg-blue-700 rounded-xl flex items-center justify-center text-white font-black text-xl tracking-tighter">
                      SEBRAE
                    </div>
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
                <p className="font-bold text-slate-800">Código SGF: {data.codigoSgf || '---'}</p>
                <p>Carga Horária: <strong>{data.cargaHorariaRealizada}</strong></p>
                <p>Data de Emissão: {data.dataAnalise}</p>
              </div>
            </div>

            {/* Client & Consulting Contract Details Table */}
            <div className="bg-slate-50 p-5 rounded-2xl border border-slate-200 text-xs space-y-3">
              <h3 className="font-bold text-slate-800 uppercase tracking-wider text-[11px] border-b border-slate-200 pb-1.5 flex items-center gap-1.5">
                <Building2 size={14} className="text-emerald-600" /> Dados do Cliente e da Consultoria Gerencial
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
                  <span className="font-bold text-slate-800">{data.consultorResponsavel || 'Consultor Credenciado'}</span>
                </div>
                <div>
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Técnico Sebrae:</span>
                  <span className="font-semibold text-slate-800">{data.tecnicoSebrae || 'Gestor do Projeto'}</span>
                </div>
              </div>
              {data.objetivoConsultoria && (
                <div className="pt-2 border-t border-slate-200/60">
                  <span className="text-slate-400 uppercase font-bold text-[10px] block">Objetivo da Consultoria:</span>
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
                <span className="text-[10px] font-bold text-slate-400 uppercase block">Problemas Resolvidos</span>
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
                <span className="text-[9px] text-slate-500 block font-semibold">Nota Máxima</span>
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
                <Briefcase size={14} className="text-blue-600" /> 2. Implementação das Soluções e Ações do Cronograma (30 Horas)
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
                    {data.empresaCredenciada || 'Empresa Credenciada Sebrae'}
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
