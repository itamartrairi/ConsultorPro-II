import React, { useState, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  TrendingUp,
  Award,
  CheckCircle2,
  Clock,
  AlertTriangle,
  HelpCircle,
  BarChart3,
  PieChart as PieIcon,
  Layers,
  ArrowRight,
  Filter,
  Plus,
  Trash2,
  Edit3,
  Save,
  Printer,
  FileSpreadsheet,
  Calendar,
  FileText,
  Building2,
  User,
  MapPin,
  FileCheck,
  CheckSquare,
  ShieldCheck,
  Zap,
  ChevronDown,
  ChevronUp,
  Sparkles,
  RefreshCw,
  Info
} from 'lucide-react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip,
  Legend,
  PieChart,
  Pie,
  Cell,
  CartesianGrid,
  ComposedChart,
  Line
} from 'recharts';
import { Button } from './Button';
import { Card } from './Card';
import { db } from '../firebase';
import { doc, updateDoc } from 'firebase/firestore';

// Tipos
export interface ProblemaAnaliseItem {
  id: string;
  problema: string;
  causa: string;
  impacto: string;
  intervencao: string;
  acaoConsultoria: string;
  solucaoImplementada: string;
  resultadoEsperado: string;
  proximoPasso: string;
  indicadorAcompanhamento: string;
  situacao: 'Resolvido' | 'Em implantação' | 'Necessita acompanhamento' | 'Pendente';
  atividadeRelacionada?: string;
  area?: string;
}

export interface IndicadorResultadoItem {
  id: string;
  nome: string;
  unidade: string;
  status: 'Sim' | 'Em implantação' | 'Não';
  valorAtual: string;
  meta: string;
  observacao: string;
  categoria: 'Financeiro' | 'Zootécnico' | 'Operacional' | 'Estratégico';
}

export interface PlanoContinuidadeItem {
  id: string;
  acao: string;
  frequencia: string;
  responsavel: string;
  indicador: string;
  meta: string;
}

export interface AnaliseResultadosState {
  problemas: ProblemaAnaliseItem[];
  indicadores: IndicadorResultadoItem[];
  planoContinuidade: PlanoContinuidadeItem[];
  analiseExecutivaTexto: string;
  conclusaoTexto: string;
  estruturaCustos: { nome: string; valor: number; cor: string }[];
  dadosEconomicos: {
    receitaTotal: number;
    custoTotal: number;
    resultadoCiclo: number;
    producaoKg: number;
    sobrevivenciaPerc: number;
    consumoRacaoKg: number;
    custoPorKg: number;
    precoMedioVenda: number;
  };
}

// 14 Problemas Típicos de Carcinicultura com a Trilha Completa
const DEFAULT_PROBLEMAS_CARCINICULTURA: ProblemaAnaliseItem[] = [
  {
    id: 'p1',
    problema: 'Ausência de registros e conciliação de fluxo de caixa diário/semanal',
    causa: 'O produtor realizava pagamentos e recebia sem anotações padronizadas ou rotina estruturada.',
    impacto: 'Desconhecimento do saldo real, impossibilidade de prever obrigações futuras e risco de inadimplência.',
    intervencao: 'Implantação de planilha gerencial de fluxo de caixa e treinamento prático de conciliação.',
    acaoConsultoria: 'Diagnóstico de fluxos financeiros e estruturação do controle diário de entradas/saídas.',
    solucaoImplementada: 'Planilha de Controle Financeiro Diário e Fluxo de Caixa Projetado.',
    resultadoEsperado: 'Maior previsibilidade financeira, controle rígido de prazos e eliminação de surpresas no caixa.',
    proximoPasso: 'Alimentar os lançamentos financeiros semanalmente às sextas-feiras.',
    indicadorAcompanhamento: 'Fluxo de caixa semanal atualizado (% de conciliação)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 2: Diagnóstico Detalhado e Levantamento de Fluxos',
    area: 'Finanças'
  },
  {
    id: 'p2',
    problema: 'Mistura de despesas pessoais/familiares com os custos do cultivo de camarão',
    causa: 'Uso de conta bancária única e ausência de definição de valor fixo de pró-labore.',
    impacto: 'Distorção do lucro real da atividade, sangria descontrolada do caixa da fazenda e descapitalização.',
    intervencao: 'Orientação para segregação de contas bancárias e fixação de teto de retirada mensal.',
    acaoConsultoria: 'Treinamento de governança financeira e definição de pró-labore.',
    solucaoImplementada: 'Política formal de segregação de contas e abertura de conta exclusiva para o viveiro.',
    resultadoEsperado: 'Finanças da carcinicultura 100% segregadas do orçamento familiar.',
    proximoPasso: 'Realizar apenas uma transferência mensal de pró-labore para a conta pessoal.',
    indicadorAcompanhamento: '% de despesas pessoais pagas pela conta do negócio (Meta: 0%)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 7: Treinamento em Segregação de Finanças e Pró-labore',
    area: 'Finanças'
  },
  {
    id: 'p3',
    problema: 'Desconhecimento do custo de produção real por quilograma (R$/kg)',
    causa: 'Falta de apuração sistemática que considere ração, pós-larvas, energia, mão de obra e insumos.',
    impacto: 'Vendas sem saber se cobrem os custos totais, gerando falsa percepção de rentabilidade.',
    intervencao: 'Estruturação da planilha de apuração de custos fixos, variáveis e custo unitário por lote.',
    acaoConsultoria: 'Levantamento de todos os insumos e custos fixos para formação da planilha de apuração.',
    solucaoImplementada: 'Planilha de Apuração de Custos por Ciclo e Custo Unitário por kg.',
    resultadoEsperado: 'Cálculo exato do custo médio por kg despescado ao término de cada lote.',
    proximoPasso: 'Registrar todas as notas fiscais de insumos e horas trabalhadas por viveiro.',
    indicadorAcompanhamento: 'Custo de produção por kg de camarão (R$/kg)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 3: Estruturação da Planilha de Apuração de Custos',
    area: 'Custos'
  },
  {
    id: 'p4',
    problema: 'Falta de controle rigoroso da conversão alimentar (FCA) e consumo de ração',
    causa: 'Alimentação sem pesagem diária e ausência de ajuste fino baseado em bandejas de alimentação.',
    impacto: 'A ração representa mais de 50% dos custos; desperdício deteriora a qualidade da água e eleva os custos.',
    intervencao: 'Capacitação na leitura diária de bandejas e planilha de controle de FCA por viveiro.',
    acaoConsultoria: 'Elaboração de rotina de biometria semanal e registro de arraçoamento diário.',
    solucaoImplementada: 'Protocolo de Controle de Arraçoamento e Planilha de Acompanhamento de FCA.',
    resultadoEsperado: 'FCA abaixo de 1.35 e redução de até 10% no desperdício de ração.',
    proximoPasso: 'Acompanhar a leitura das bandejas 4 vezes ao dia e registrar no diário de campo.',
    indicadorAcompanhamento: 'Fator de Conversão Alimentar (FCA = kg ração / kg camarão)',
    situacao: 'Em implantação',
    atividadeRelacionada: 'Atividade 5: Treinamento nos Indicadores Econômicos e Zootécnicos',
    area: 'Produção'
  },
  {
    id: 'p5',
    problema: 'Inexistência de cálculo da margem de contribuição por lote/viveiro',
    causa: 'Ausência de ferramentas que subtraiam os custos variáveis diretos da receita de venda.',
    impacto: 'Dificuldade em identificar quais viveiros ou densidades trazem maior ganho líquido real.',
    intervencao: 'Treinamento sobre o conceito de margem de contribuição e aplicação no simulador do ciclo.',
    acaoConsultoria: 'Configuração da planilha para cálculo automático da margem em R$ e % por lote.',
    solucaoImplementada: 'Módulo de Margem de Contribuição por Viveiro e Despesca.',
    resultadoEsperado: 'Visão clara da sobra financeira gerada por cada despesca para cobrir custos fixos.',
    proximoPasso: 'Calcular a margem de contribuição no dia seguinte a cada despesca realizada.',
    indicadorAcompanhamento: 'Margem de contribuição percentual (%) e em R$ por lote',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 5: Treinamento nos Indicadores Econômicos e Zootécnicos',
    area: 'Finanças'
  },
  {
    id: 'p6',
    problema: 'Desconhecimento do ponto de equilíbrio econômico e operacional em kg',
    causa: 'Falta de cruzamento entre custos fixos do empreendimento e margem unitária do camarão.',
    impacto: 'Incerteza sobre qual volume mínimo (kg) precisa ser produzido para não operar no prejuízo.',
    intervencao: 'Desenvolvimento do cálculo de ponto de equilíbrio (Break-even) em quantidade e receita.',
    acaoConsultoria: 'Treinamento e modelagem do ponto de equilíbrio operacional e financeiro.',
    solucaoImplementada: 'Painel de Ponto de Equilíbrio em kg de camarão e faturamento mínimo mensal.',
    resultadoEsperado: 'Saber com precisão quantos kg de camarão devem ser comercializados para cobrir todas as despesas.',
    proximoPasso: 'Revisar o ponto de equilíbrio sempre que houver reajuste no preço da ração ou energia.',
    indicadorAcompanhamento: 'Ponto de equilíbrio (kg de camarão necessários por ciclo)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 5: Treinamento nos Indicadores Econômicos e Zootécnicos',
    area: 'Finanças'
  },
  {
    id: 'p7',
    problema: 'Ausência de reserva de emergência para entressafras, manutenções e riscos sanitários',
    causa: 'Todo o lucro de despescas favoráveis era retirado imediatamente sem retenção preventiva.',
    impacto: 'Vulnerabilidade extrema a quebras de aeradores, chuvas intensas ou oscilações de mercado.',
    intervencao: 'Elaboração de plano financeiro para retenção de 10% do resultado líquido por ciclo.',
    acaoConsultoria: 'Definição da meta de reserva de emergência equivalente a 1 ciclo de custos operacionais.',
    solucaoImplementada: 'Plano de Formação de Fundo de Reserva com conta de investimento dedicada.',
    resultadoEsperado: 'Constituição progressiva de reserva para garantir capital de giro nas entressafras.',
    proximoPasso: 'Destinar 10% da receita líquida de cada despesca para a conta de reserva.',
    indicadorAcompanhamento: 'Saldo acumulado na Reserva de Emergência (R$)',
    situacao: 'Em implantação',
    atividadeRelacionada: 'Atividade 4: Plano de Formação de Reserva de Emergência',
    area: 'Finanças'
  },
  {
    id: 'p8',
    problema: 'Precificação empírica e venda de camarão abaixo do preço mínimo viável',
    causa: 'Preço aceito passivamente conforme cotação de intermediários sem conferência do custo de piso.',
    impacto: 'Risco de fechar lotes com margem negativa em momentos de pressão de compradores.',
    intervencao: 'Criação de simulador de preço de venda ideal e preço de piso por gramatura (tamanho).',
    acaoConsultoria: 'Desenvolvimento de tabela de precificação dinâmica atrelada ao custo por kg e gramatura.',
    solucaoImplementada: 'Tabela de Precificação Baseada em Custos e Margem Alvo.',
    resultadoEsperado: 'Maior poder de negociação e recusa fundamentada de propostas abaixo do custo mínimo.',
    proximoPasso: 'Consultar a tabela de preço mínimo antes de fechar qualquer contrato de despesca.',
    indicadorAcompanhamento: 'Preço médio de venda vs. Preço mínimo viável (R$/kg)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 3: Estruturação da Planilha de Apuração de Custos',
    area: 'Comercial'
  },
  {
    id: 'p9',
    problema: 'Acompanhamento irregular das taxas de sobrevivência e biometrias zootécnicas',
    causa: 'Falta de rotina semanal para pesagem amostral e contagem de biomassa nos viveiros.',
    impacto: 'Dificuldade de planejar o momento ideal de despesca e erro na dosagem de ração diária.',
    intervencao: 'Estruturação de ficha de biometria semanal e treinamento para estimativa de biomassa.',
    acaoConsultoria: 'Padronização da rotina de amostragem zootécnica semanal.',
    solucaoImplementada: 'Ficha de Biometria Semanal e Gráfico de Curva de Crescimento.',
    resultadoEsperado: 'Monitoramento contínuo do peso médio semanal e estimativa precisa de sobrevivência.',
    proximoPasso: 'Realizar biometria com 100 camarões por viveiro todas as terças-feiras.',
    indicadorAcompanhamento: 'Taxa de Sobrevivência (%) e Ganho de Peso Semanal (g/semana)',
    situacao: 'Em implantação',
    atividadeRelacionada: 'Atividade 5: Treinamento nos Indicadores Econômicos e Zootécnicos',
    area: 'Produção'
  },
  {
    id: 'p10',
    problema: 'Inexistência de DRE Gerencial (Demonstrativo do Resultado do Exercício) por ciclo',
    causa: 'O produtor apenas acompanhava o saldo bancário, sem visão estruturada de receitas, custos e lucro líquido.',
    impacto: 'Falta de visão gerencial do negócio e incapacidade de comparar a lucratividade entre safras.',
    intervencao: 'Modelagem do DRE Gerencial simplificado específico para carcinicultura.',
    acaoConsultoria: 'Elaboração e treinamento no DRE Gerencial por ciclo de cultivo.',
    solucaoImplementada: 'DRE Gerencial Automatizado com apuração de Lucro Bruto e Lucro Líquido.',
    resultadoEsperado: 'Emissão de relatório de resultado econômico fechado ao final de cada despesca.',
    proximoPasso: 'Consolidar o DRE no fechamento de cada ciclo e arquivar para histórico comparativo.',
    indicadorAcompanhamento: 'Lucro Líquido do Ciclo (R$) e Margem Líquida (%)',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 8: Apresentação de Resultados e Relatório Final',
    area: 'Finanças'
  },
  {
    id: 'p11',
    problema: 'Dificuldade de acesso a crédito rural orientado e custeio com juros reduzidos',
    causa: 'Falta de demonstrativos contábeis/gerenciais organizados para apresentação às instituições financeiras.',
    impacto: 'Recurso a crédito pessoal de alto custo ou dependência de adiantamento de atravessadores.',
    intervencao: 'Organização do dossiê financeiro do produtor com histórico de safras e DREs estruturados.',
    acaoConsultoria: 'Estruturação de pacote informativo de viabilidade econômico-financeira para bancos.',
    solucaoImplementada: 'Dossiê Financeiro Gerencial para Apresentação a Agentes de Crédito (PRONAF/FCO).',
    resultadoEsperado: 'Condições de pleitear linhas de custeio e investimento com taxas subsidiadas.',
    proximoPasso: 'Apresentar os demonstrativos organizados ao gerente do banco na próxima renovação de limite.',
    indicadorAcompanhamento: 'Acesso a linhas de crédito estruturadas com taxa reduzida',
    situacao: 'Necessita acompanhamento',
    atividadeRelacionada: 'Atividade 8: Apresentação de Resultados e Relatório Final',
    area: 'Estratégia'
  },
  {
    id: 'p12',
    problema: 'Ausência de rotina estruturada de conciliação de contas a pagar e a receber',
    causa: 'Controle de pagamentos feito por comprovantes físicos soltos e memória do produtor.',
    impacto: 'Pagamento de juros e multas por atraso e esquecimento de cobrança de vendas a prazo.',
    intervencao: 'Configuração de calendário financeiro de vencimentos na planilha gerencial.',
    acaoConsultoria: 'Implantação de rotina de controle de contas a pagar e controle de clientes compradores.',
    solucaoImplementada: 'Agenda Financeira de Contas a Pagar e Receber com alertas visuais.',
    resultadoEsperado: 'Pontualidade de 100% nos pagamentos e eliminação total de multas financeiras.',
    proximoPasso: 'Checar a agenda financeira toda segunda-feira de manhã.',
    indicadorAcompanhamento: 'Percentual de pagamentos em dia e saldo de contas a receber',
    situacao: 'Resolvido',
    atividadeRelacionada: 'Atividade 2: Diagnóstico Detalhado e Levantamento de Fluxos',
    area: 'Finanças'
  },
  {
    id: 'p13',
    problema: 'Planejamento financeiro deficiente para compra antecipada de insumos e pós-larvas',
    causa: 'Compra de ração e pós-larvas em pequenas quantidades em cima da hora com preços mais caros.',
    impacto: 'Custos mais altos de frete e insumos, e risco de atraso no povoamento dos viveiros.',
    intervencao: 'Elaboração de cronograma de compras trimestral atrelado ao plano de povoamento.',
    acaoConsultoria: 'Planejamento de fluxo de caixa para aquisição programada de insumos com desconto.',
    solucaoImplementada: 'Calendário de Compras Planejadas e Negociação em Volume.',
    resultadoEsperado: 'Economia de 5 a 8% na aquisição de insumos por negociação antecipada.',
    proximoPasso: 'Programar a compra do próximo lote de ração com 30 dias de antecedência.',
    indicadorAcompanhamento: 'Economia obtida em compras programadas de insumos (R$)',
    situacao: 'Necessita acompanhamento',
    atividadeRelacionada: 'Atividade 6: Elaboração do Plano de Gestão e Continuidade',
    area: 'Operações'
  },
  {
    id: 'p14',
    problema: 'Falta de padronização de procedimentos operacionais e relatórios de despesca',
    causa: 'Cada despesca ocorria de forma diferente sem registro de perdas, quebras ou tempo de gelo.',
    impacto: 'Variação na qualidade final do camarão e falta de histórico comparativo entre tanques.',
    intervencao: 'Elaboração de Procedimento Operacional Padrão (POP) simplificado para despescas.',
    acaoConsultoria: 'Criação de checklist de despesca e ficha técnica de controle de lote.',
    solucaoImplementada: 'Procedimento Operacional Padrão (POP) e Relatório de Encerramento de Lote.',
    resultadoEsperado: 'Padronização da qualidade do camarão entregue e preservação do frescor.',
    proximoPasso: 'Preencher a ficha técnica de despesca no encerramento de todos os lotes.',
    indicadorAcompanhamento: 'Fichas de despesca preenchidas por ciclo (%)',
    situacao: 'Em implantação',
    atividadeRelacionada: 'Atividade 6: Elaboração do Plano de Gestão e Continuidade',
    area: 'Operações'
  }
];

// 14 Indicadores de Resultado
const DEFAULT_INDICADORES_CARCINICULTURA: IndicadorResultadoItem[] = [
  { id: 'ind1', nome: 'Receita Total do Ciclo', unidade: 'R$', status: 'Sim', valorAtual: 'R$ 84.500,00', meta: 'R$ 80.000,00', observacao: 'Faturamento consolidado da despesca dos viveiros 1 e 2.', categoria: 'Financeiro' },
  { id: 'ind2', nome: 'Custo Total do Ciclo', unidade: 'R$', status: 'Sim', valorAtual: 'R$ 56.200,00', meta: '< R$ 58.000,00', observacao: 'Inclui ração, pós-larvas, energia, mão de obra e insumos.', categoria: 'Financeiro' },
  { id: 'ind3', nome: 'Custo por kg de Camarão', unidade: 'R$/kg', status: 'Sim', valorAtual: 'R$ 14,05 / kg', meta: '< R$ 14,50 / kg', observacao: 'Apuração do custo unitário fechado no ciclo.', categoria: 'Financeiro' },
  { id: 'ind4', nome: 'Consumo Total de Ração', unidade: 'kg', status: 'Sim', valorAtual: '5.200 kg', meta: '5.000 kg', observacao: 'Acompanhado através da planilha de arraçoamento.', categoria: 'Zootécnico' },
  { id: 'ind5', nome: 'Fator de Conversão Alimentar (FCA)', unidade: 'kg/kg', status: 'Sim', valorAtual: '1.30', meta: '< 1.35', observacao: 'Melhoria na leitura de bandejas reduziu o desperdício.', categoria: 'Zootécnico' },
  { id: 'ind6', nome: 'Taxa de Sobrevivência', unidade: '%', status: 'Em implantação', valorAtual: '78%', meta: '> 80%', observacao: 'Biometrias semanais implantadas no viveiro 1 e 2.', categoria: 'Zootécnico' },
  { id: 'ind7', nome: 'Margem de Contribuição', unidade: 'R$ / %', status: 'Sim', valorAtual: 'R$ 38.300,00 (45,3%)', meta: '> 40%', observacao: 'Sobra líquida suficiente para cobrir custos fixos e gerar lucro.', categoria: 'Financeiro' },
  { id: 'ind8', nome: 'Ponto de Equilíbrio do Ciclo', unidade: 'kg', status: 'Sim', valorAtual: '2.450 kg', meta: '< 2.600 kg', observacao: 'Volume mínimo a ser comercializado para cobrir custos totais.', categoria: 'Financeiro' },
  { id: 'ind9', nome: 'Resultado / Lucro Líquido do Ciclo', unidade: 'R$', status: 'Sim', valorAtual: 'R$ 28.300,00', meta: '> R$ 22.000,00', observacao: 'Lucro líquido gerado após dedução de todas as despesas.', categoria: 'Financeiro' },
  { id: 'ind10', nome: 'Preço Médio de Venda', unidade: 'R$/kg', status: 'Sim', valorAtual: 'R$ 21,12 / kg', meta: '> R$ 20,00 / kg', observacao: 'Preço médio ponderado pelas diferentes gramaturas comercializadas.', categoria: 'Comercial' as any },
  { id: 'ind11', nome: 'Preço Mínimo de Venda (Piso)', unidade: 'R$/kg', status: 'Sim', valorAtual: 'R$ 15,20 / kg', meta: '< Preço Venda', observacao: 'Preço de piso que garante cobertura dos custos variáveis e fixos.', categoria: 'Financeiro' },
  { id: 'ind12', nome: 'Fluxo de Caixa Estruturado', unidade: 'Status', status: 'Sim', valorAtual: 'Conciliado semanal', meta: '100% em dia', observacao: 'Planilha utilizada regularmente pelo produtor.', categoria: 'Operacional' },
  { id: 'ind13', nome: 'Valor da Reserva de Emergência', unidade: 'R$', status: 'Em implantação', valorAtual: 'R$ 5.500,00', meta: 'R$ 15.000,00', observacao: 'Fundo criado com retenção de 10% da receita líquida da despesca.', categoria: 'Estratégico' },
  { id: 'ind14', nome: 'Segregação de Despesas Pessoais', unidade: '%', status: 'Sim', valorAtual: '0% de mistura', meta: '0% de mistura', observacao: 'Contas bancárias separadas e pró-labore definido.', categoria: 'Estratégico' }
];

// Plano de Continuidade Pós-Consultoria
const DEFAULT_PLANO_CONTINUIDADE: PlanoContinuidadeItem[] = [
  { id: 'pl1', acao: 'Atualizar registros financeiros diários e conciliação semanal', frequencia: 'Semanalmente (toda sexta-feira)', responsavel: 'Produtor / Responsável Financeiro', indicador: '% de lançamentos conciliados no caixa', meta: '100% de conciliação' },
  { id: 'pl2', acao: 'Controlar consumo diário de ração e leitura das bandejas de alimentação', frequencia: 'Diariamente (4x ao dia)', responsavel: 'Tratador / Produtor', indicador: 'Fator de Conversão Alimentar (FCA)', meta: 'FCA ≤ 1.30' },
  { id: 'pl3', acao: 'Analisar apuração de custos totais e custo unitário por kg despescado', frequencia: 'Ao término de cada lote/despesca', responsavel: 'Produtor', indicador: 'Custo de produção por kg (R$/kg)', meta: 'R$ ≤ 14,20/kg' },
  { id: 'pl4', acao: 'Acompanhar a margem de contribuição a cada venda de camarão', frequencia: 'A cada venda/despesca', responsavel: 'Produtor', indicador: 'Margem de contribuição percentual (%)', meta: 'Margem ≥ 42%' },
  { id: 'pl5', acao: 'Avaliar e recalcular o ponto de equilíbrio operacional em kg', frequencia: 'Por ciclo de cultivo', responsavel: 'Produtor', indicador: 'Ponto de equilíbrio (kg)', meta: 'Abaixo de 60% da capacidade' },
  { id: 'pl6', acao: 'Atualizar planejamento financeiro e fluxo de caixa projetado', frequencia: 'Antes de iniciar cada novo ciclo', responsavel: 'Produtor', indicador: 'Plano de despesas e compras projetadas', meta: 'Aprovado antes do povoamento' },
  { id: 'pl7', acao: 'Formar e aportar na Reserva de Emergência para riscos e entressafras', frequencia: 'A cada despesca realizada', responsavel: 'Produtor', indicador: 'Saldo do Fundo de Reserva (R$)', meta: 'Retenção de 10% do lucro líquido' },
  { id: 'pl8', acao: 'Manter a separação rígida entre contas pessoais e empresariais', frequencia: 'Permanentemente / Diariamente', responsavel: 'Produtor / Família', indicador: 'Despesas pessoais na conta do viveiro', meta: 'R$ 0,00 de despesas pessoais' },
  { id: 'pl9', acao: 'Avaliar o resultado econômico e DRE ao final de cada safra', frequencia: 'Ao final de cada ciclo', responsavel: 'Produtor / Consultor', indicador: 'Lucro Líquido do Ciclo e ROI', meta: 'Lucro Líquido > R$ 25.000/safra' }
];

export interface AnaliseResultadosViewProps {
  selectedDiagnostico: any;
  selectedEmpresa: any;
  empresasCredenciadas?: any[];
  dadosConsultoria?: any;
  respostas?: any[];
  cronograma?: any[];
  setView: (v: string) => void;
  playSuccessSound: () => void;
  onSaveConsultoriaState?: (updatedDiag: any) => Promise<void>;
}

export const AnaliseResultadosView: React.FC<AnaliseResultadosViewProps> = ({
  selectedDiagnostico,
  selectedEmpresa,
  empresasCredenciadas = [],
  dadosConsultoria = {} as any,
  respostas = [],
  cronograma = [],
  setView,
  playSuccessSound,
  onSaveConsultoriaState
}) => {
  // Estado principal
  const [problemas, setProblemas] = useState<ProblemaAnaliseItem[]>(() => {
    if (selectedDiagnostico?.analiseResultados?.problemas?.length > 0) {
      return selectedDiagnostico.analiseResultados.problemas;
    }
    return DEFAULT_PROBLEMAS_CARCINICULTURA;
  });

  const [indicadores, setIndicadores] = useState<IndicadorResultadoItem[]>(() => {
    if (selectedDiagnostico?.analiseResultados?.indicadores?.length > 0) {
      return selectedDiagnostico.analiseResultados.indicadores;
    }
    return DEFAULT_INDICADORES_CARCINICULTURA;
  });

  const [planoContinuidade, setPlanoContinuidade] = useState<PlanoContinuidadeItem[]>(() => {
    if (selectedDiagnostico?.analiseResultados?.planoContinuidade?.length > 0) {
      return selectedDiagnostico.analiseResultados.planoContinuidade;
    }
    return DEFAULT_PLANO_CONTINUIDADE;
  });

  const [filtroSituacao, setFiltroSituacao] = useState<string>('Todos');
  const [filtroArea, setFiltroArea] = useState<string>('Todas');
  const [buscaTexto, setBuscaTexto] = useState<string>('');
  const [expandedProblemaId, setExpandedProblemaId] = useState<string | null>('p1');
  const [isEditingAnaliseExecutiva, setIsEditingAnaliseExecutiva] = useState(false);
  const [isEditingConclusao, setIsEditingConclusao] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [activeTabSubSection, setActiveTabSubSection] = useState<'matriz' | 'trilha' | 'indicadores' | 'graficos' | 'executiva' | 'continuidade'>('matriz');

  // Dados do Cliente e Cabeçalho
  const consultoriaData = (dadosConsultoria || {}) as any;
  const clienteNome = selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || 'Produtor Rural de Carcinicultura';
  const clienteDocumento = selectedEmpresa?.cnpj || selectedEmpresa?.cpfRepresentante || selectedEmpresa?.cafNumero || '000.000.000-00';
  const clienteAtividade = selectedEmpresa?.ramoAtividade || selectedEmpresa?.tipoEmpresa || 'Carcinicultura';
  const clienteEndereco = selectedEmpresa?.enderecoComercial || 'Zona Rural / Polo Aquícola';
  const periodoConsultoria = consultoriaData?.periodoConsultoria || 'Fevereiro a Março de 2026';
  const consultorNome = consultoriaData?.consultor || 'Consultor Especialista em Gestão e Aquicultura';
  const tecnicoSebrae = consultoriaData?.tecnicoSebrae || 'Gestor SEBRAE';
  const codigoSgf = consultoriaData?.codigoSgf || 'SGF-89472-CARC';
  const cargaHoraria = consultoriaData?.cargaHoraria || '30 Horas';

  // Cálculos de KPI
  const kpis = useMemo(() => {
    const total = problemas.length;
    const resolvidos = problemas.filter(p => p.situacao === 'Resolvido').length;
    const emImplantacao = problemas.filter(p => p.situacao === 'Em implantação').length;
    const acompanhamento = problemas.filter(p => p.situacao === 'Necessita acompanhamento').length;
    const pendentes = problemas.filter(p => p.situacao === 'Pendente').length;
    const trabalhados = resolvidos + emImplantacao + acompanhamento;
    const percentualCobertura = total > 0 ? Math.round((trabalhados / total) * 100) : 100;
    const percentualResolvidos = total > 0 ? Math.round((resolvidos / total) * 100) : 0;

    const totalIndicadores = indicadores.length;
    const indicadoresImplantados = indicadores.filter(i => i.status === 'Sim').length;
    const indicadoresEmProgresso = indicadores.filter(i => i.status === 'Em implantação').length;

    const totalAtividades = 8;
    const totalSolucoes = 7;

    return {
      total,
      resolvidos,
      emImplantacao,
      acompanhamento,
      pendentes,
      trabalhados,
      percentualCobertura,
      percentualResolvidos,
      totalIndicadores,
      indicadoresImplantados,
      indicadoresEmProgresso,
      totalAtividades,
      totalSolucoes
    };
  }, [problemas, indicadores]);

  // Texto Padrão da Análise Executiva
  const defaultAnaliseExecutivaTexto = useMemo(() => {
    return `ANÁLISE EXECUTIVA DOS RESULTADOS DA CONSULTORIA

1. QUAIS ERAM OS PRINCIPAIS PROBLEMAS ENCONTRADOS?
No diagnóstico inicial da propriedade de Carcinicultura, identificou-se ausência de registros contínuos de fluxo de caixa, mistura constante de finanças pessoais com as contas da atividade aquícola, desconhecimento do custo de produção real por kg de camarão, falta de cálculo de margem de contribuição e ponto de equilíbrio, bem como ausência de fundo de reserva para entressafras e riscos sanitários.

2. QUAIS PROBLEMAS FORAM TRATADOS?
Durante as 30 horas de consultoria gerencial, foram abordados 100% dos 14 pontos críticos mapeados. Foram priorizadas as áreas de Gestão Financeira, Custos e Resultados, Controle de Rotina, Formação de Preço e Planejamento para Sustentabilidade.

3. QUAIS FERRAMENTAS FORAM IMPLANTADAS?
Foram desenvolvidas e implantadas na rotina do produtor:
• Planilha de Fluxo de Caixa Diário e Conciliação Semanal;
• Planilha de Apuração de Custos por Ciclo e Custo Unitário por kg (R$/kg);
• Módulo de Margem de Contribuição por Viveiro e Despesca;
• Painel de Ponto de Equilíbrio em kg de camarão e Faturamento Mínimo;
• Tabela de Precificação Baseada em Custos e Margem Alvo;
• Protocolo de Controle de Arraçoamento e Planilha de FCA;
• Política de Segregação de Contas Bancárias e Pró-labore.

4. QUAIS MELHORIAS FORAM OBSERVADAS?
Houve eliminação imediata da mistura de despesas pessoais na conta da fazenda através da fixação do pró-labore. O produtor passou a registrar semanalmente todas as compras de insumos e vendas, permitindo saber exatamente o custo por kg antes de negociar a despesca, resultando em maior segurança comercial.

5. QUAIS INDICADORES PASSARAM A SER ACOMPANHADOS?
O produtor agora monitora ativamente: Receita total, Custo total, Custo por kg de camarão, Fator de Conversão Alimentar (FCA), Margem de Contribuição (R$ e %), Ponto de Equilíbrio (kg), Preço Médio vs. Mínimo de Venda e saldo da Reserva de Emergência.

6. QUAIS PROBLEMAS AINDA PRECISAM DE ACOMPANHAMENTO?
A consolidação da disciplina de biometria semanal para estimativa zootécnica de biomassa e o fortalecimento contínuo do saldo da reserva de emergência (meta de atingir 1 ciclo de custos operacionais) exigem acompanhamento contínuo nos próximos ciclos.

7. QUAIS SÃO OS PRINCIPAIS RISCOS PARA O PRÓXIMO CICLO?
Oscilação brusca no preço da ração, descontinuidade das anotações de campo por sobrecarga operacional na despesca e eventual retorno à retirada descontrolada de valores além do pró-labore estabelecido.

8. QUAIS AÇÕES DEVEM CONTINUAR APÓS O ENCERRAMENTO DA CONSULTORIA?
Manter a alimentação semanal da planilha financeira, realizar a leitura diária de bandejas de ração, apurar o DRE gerencial ao final de cada safra e destinar pontualmente 10% do resultado líquido para o fundo de reserva.`;
  }, []);

  const [analiseExecutivaTexto, setAnaliseExecutivaTexto] = useState<string>(() => {
    return selectedDiagnostico?.analiseResultados?.analiseExecutivaTexto || defaultAnaliseExecutivaTexto;
  });

  // Conclusão Padrão
  const defaultConclusaoTexto = useMemo(() => {
    return `A consultoria gerencial em carcinicultura proporcionou uma transformação estrutural na gestão do empreendimento. Por meio da implantação de ferramentas práticas, planilhas gerenciais de apuração de custos e treinamento intensivo em indicadores financeiros e produtivos, o produtor adquiriu pleno domínio sobre seus números, fortalecendo sua capacidade de tomada de decisão e negociação no mercado.

Ressalta-se que a consultoria entregou as metodologias, ferramentas e capacitações necessárias. Contudo, a consolidação definitiva dos resultados e a perenidade dos ganhos de lucratividade dependem diretamente do compromisso do produtor em manter a disciplina diária de anotações, conciliação e cumprimento do plano de continuidade pós-consultoria.`;
  }, []);

  const [conclusaoTexto, setConclusaoTexto] = useState<string>(() => {
    return selectedDiagnostico?.analiseResultados?.conclusaoTexto || defaultConclusaoTexto;
  });

  // Dados para Gráficos
  const chartDataProblemasSolucoes = useMemo(() => {
    return [
      { categoria: 'Identificados', quantidade: kpis.total, fill: '#3b82f6' },
      { categoria: 'Trabalhados', quantidade: kpis.trabalhados, fill: '#10b981' },
      { categoria: 'Resolvidos', quantidade: kpis.resolvidos, fill: '#059669' },
      { categoria: 'Em Implantação', quantidade: kpis.emImplantacao, fill: '#f59e0b' },
      { categoria: 'Acompanhamento', quantidade: kpis.acompanhamento, fill: '#8b5cf6' }
    ];
  }, [kpis]);

  const chartDataEstruturaCustos = useMemo(() => {
    return [
      { name: 'Ração para Camarão', value: 52, color: '#059669' },
      { name: 'Pós-Larvas / Povoamento', value: 14, color: '#0284c7' },
      { name: 'Energia Elétrica / Aeradores', value: 12, color: '#f59e0b' },
      { name: 'Mão de Obra e Encargos', value: 10, color: '#6366f1' },
      { name: 'Insumos / Biorremediação', value: 7, color: '#ec4899' },
      { name: 'Outros Custos Operacionais', value: 5, color: '#94a3b8' }
    ];
  }, []);

  const chartDataReceitaCustos = useMemo(() => {
    return [
      { name: 'Ciclo Anterior (Estimado)', Receita: 68000, CustoTotal: 58000, LucroLiquido: 10000 },
      { name: 'Ciclo Atual (Consultoria)', Receita: 84500, CustoTotal: 56200, LucroLiquido: 28300 },
      { name: 'Meta Próximo Ciclo', Receita: 92000, CustoTotal: 57500, LucroLiquido: 34500 }
    ];
  }, []);

  const chartDataIndicadoresProdutivos = useMemo(() => {
    return [
      { periodo: 'Ciclo 1', producaoKg: 3200, fca: 1.55, custoKg: 16.80, sobrevivencia: 65 },
      { periodo: 'Ciclo 2', producaoKg: 3600, fca: 1.42, custoKg: 15.20, sobrevivencia: 72 },
      { periodo: 'Ciclo 3 (Consultoria)', producaoKg: 4000, fca: 1.30, custoKg: 14.05, sobrevivencia: 78 },
      { periodo: 'Ciclo 4 (Meta)', producaoKg: 4400, fca: 1.25, custoKg: 13.50, sobrevivencia: 82 }
    ];
  }, []);

  const chartDataSituacaoPie = useMemo(() => {
    return [
      { name: 'Resolvidos', value: kpis.resolvidos, color: '#10b981' },
      { name: 'Em implantação', value: kpis.emImplantacao, color: '#f59e0b' },
      { name: 'Acompanhamento', value: kpis.acompanhamento, color: '#8b5cf6' },
      { name: 'Pendentes', value: kpis.pendentes, color: '#ef4444' }
    ].filter(item => item.value > 0);
  }, [kpis]);

  // Lista Filtrada de Problemas
  const filteredProblemas = useMemo(() => {
    return problemas.filter(p => {
      const matchSituacao = filtroSituacao === 'Todos' || p.situacao === filtroSituacao;
      const matchArea = filtroArea === 'Todas' || p.area === filtroArea;
      const matchBusca = !buscaTexto.trim() || 
        p.problema.toLowerCase().includes(buscaTexto.toLowerCase()) ||
        p.solucaoImplementada.toLowerCase().includes(buscaTexto.toLowerCase()) ||
        p.indicadorAcompanhamento.toLowerCase().includes(buscaTexto.toLowerCase());
      return matchSituacao && matchArea && matchBusca;
    });
  }, [problemas, filtroSituacao, filtroArea, buscaTexto]);

  // Áreas únicas
  const areasDisponiveis = useMemo(() => {
    const set = new Set(problemas.map(p => p.area).filter(Boolean));
    return ['Todas', ...Array.from(set)];
  }, [problemas]);

  // Salvar Estado
  const handleSalvarAnalise = async () => {
    setIsSaving(true);
    try {
      const payload: AnaliseResultadosState = {
        problemas,
        indicadores,
        planoContinuidade,
        analiseExecutivaTexto,
        conclusaoTexto,
        estruturaCustos: chartDataEstruturaCustos,
        dadosEconomicos: {
          receitaTotal: 84500,
          custoTotal: 56200,
          resultadoCiclo: 28300,
          producaoKg: 4000,
          sobrevivenciaPerc: 78,
          consumoRacaoKg: 5200,
          custoPorKg: 14.05,
          precoMedioVenda: 21.12
        }
      };

      // Salva localmente no storage
      if (selectedDiagnostico?.id) {
        localStorage.setItem(`analise_resultados_${selectedDiagnostico.id}`, JSON.stringify(payload));
      }

      // Atualiza cloud se callback disponível ou via Firestore
      if (onSaveConsultoriaState && selectedDiagnostico) {
        await onSaveConsultoriaState({
          ...selectedDiagnostico,
          analiseResultados: payload
        });
      } else if (selectedDiagnostico?.id) {
        try {
          await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), {
            analiseResultados: payload,
            updatedAt: new Date().toISOString()
          });
        } catch (cloudErr) {
          console.warn("Sync cloud note for analise de resultados:", cloudErr);
        }
      }

      playSuccessSound();
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (e) {
      console.error("Erro ao salvar análise de resultados:", e);
      alert("Erro ao salvar análise de resultados.");
    } finally {
      setIsSaving(false);
    }
  };

  // Função para alterar situação do problema
  const handleUpdateSituacaoProblema = (id: string, novaSituacao: ProblemaAnaliseItem['situacao']) => {
    setProblemas(prev => prev.map(p => p.id === id ? { ...p, situacao: novaSituacao } : p));
  };

  // Função para alterar status do indicador
  const handleUpdateStatusIndicador = (id: string, novoStatus: IndicadorResultadoItem['status']) => {
    setIndicadores(prev => prev.map(ind => ind.id === id ? { ...ind, status: novoStatus } : ind));
  };

  // Imprimir / Exportar
  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="space-y-8 max-w-7xl mx-auto pb-16 print:p-0 print:m-0 print:max-w-full">
      {/* 1. CABEÇALHO DA ABA - PADRÃO EXECUTIVO SEBRAE */}
      <div className="bg-gradient-to-r from-slate-900 via-slate-800 to-emerald-950 text-white rounded-3xl p-6 sm:p-8 shadow-xl border border-slate-700/60 relative overflow-hidden print:bg-white print:text-slate-900 print:border print:border-slate-300 print:shadow-none print:rounded-none">
        {/* Glow de fundo */}
        <div className="absolute top-0 right-0 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl pointer-events-none print:hidden" />
        <div className="absolute bottom-0 left-1/3 w-64 h-64 bg-blue-500/10 rounded-full blur-3xl pointer-events-none print:hidden" />

        <div className="relative z-10">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6 pb-6 border-b border-slate-700/60 print:border-slate-300">
            <div>
              <div className="flex flex-wrap items-center gap-2 mb-2">
                <span className="px-3 py-1 text-[11px] font-black uppercase tracking-wider bg-emerald-500 text-slate-950 rounded-full flex items-center gap-1.5 shadow-sm">
                  <Award size={13} /> Relatório de Resultados SEBRAE
                </span>
                <span className="px-3 py-1 text-[11px] font-bold uppercase tracking-wider bg-slate-800 text-emerald-300 rounded-full border border-emerald-500/30">
                  {clienteAtividade}
                </span>
                <span className="px-3 py-1 text-[11px] font-mono font-bold bg-slate-800/90 text-slate-300 rounded-full border border-slate-700">
                  {codigoSgf}
                </span>
              </div>
              <h1 className="text-2xl sm:text-3xl lg:text-4xl font-black tracking-tight text-white print:text-slate-900">
                ANÁLISE DE RESULTADOS DA CONSULTORIA
              </h1>
              <p className="text-slate-300 text-sm sm:text-base mt-1.5 max-w-3xl leading-relaxed print:text-slate-600">
                Matriz consolidada de problemas diagnosticados, intervenções técnicas realizadas, soluções implantadas, evolução de indicadores zootécnicos/econômicos e plano de continuidade pós-consultoria.
              </p>
            </div>

            <div className="flex flex-wrap items-center gap-3 print:hidden">
              <Button
                variant="outline"
                onClick={handlePrint}
                className="bg-slate-800/80 hover:bg-slate-700 border-slate-600 text-slate-200 text-xs font-bold"
              >
                <Printer size={15} className="mr-1.5 text-emerald-400" /> Imprimir / PDF
              </Button>
              <Button
                onClick={handleSalvarAnalise}
                disabled={isSaving}
                className={`text-xs font-black px-5 shadow-lg transition-all ${
                  saveSuccess 
                    ? 'bg-emerald-500 hover:bg-emerald-600 text-slate-950' 
                    : 'bg-emerald-600 hover:bg-emerald-500 text-white'
                }`}
              >
                {isSaving ? (
                  <>
                    <RefreshCw size={15} className="animate-spin mr-1.5" /> Salvando...
                  </>
                ) : saveSuccess ? (
                  <>
                    <CheckCircle2 size={15} className="mr-1.5" /> Salvo com Sucesso!
                  </>
                ) : (
                  <>
                    <Save size={15} className="mr-1.5" /> Salvar Análise de Resultados
                  </>
                )}
              </Button>
            </div>
          </div>

          {/* Dados do Cliente e da Consultoria */}
          <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 sm:gap-4 mt-6 text-xs">
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">Cliente</span>
              <span className="font-bold text-slate-100 text-sm truncate block print:text-slate-900" title={clienteNome}>{clienteNome}</span>
            </div>
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">CPF / CNPJ</span>
              <span className="font-semibold font-mono text-slate-200 block print:text-slate-900">{clienteDocumento}</span>
            </div>
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">Atividade</span>
              <span className="font-semibold text-emerald-400 block print:text-slate-900">{clienteAtividade}</span>
            </div>
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">Endereço / Polo</span>
              <span className="font-medium text-slate-200 truncate block print:text-slate-900" title={clienteEndereco}>{clienteEndereco}</span>
            </div>
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">Período</span>
              <span className="font-semibold text-slate-200 block print:text-slate-900">{periodoConsultoria}</span>
            </div>
            <div className="p-3.5 bg-slate-800/60 rounded-xl border border-slate-700/50 print:bg-slate-50 print:border-slate-200">
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-1 print:text-slate-500">Carga Horária</span>
              <span className="font-black text-emerald-400 text-sm block print:text-slate-900">{cargaHoraria}</span>
            </div>
          </div>
        </div>
      </div>

      {/* 2. PAINEL RESUMO (CARDS / KPIS NO TOPO) */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3.5">
        <Card className="p-4 bg-white border-slate-200/90 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Problemas</span>
            <AlertTriangle size={16} className="text-amber-500" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-slate-800">{kpis.total}</div>
            <p className="text-[11px] text-slate-500 font-medium">Identificados no Diagnóstico</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-amber-700 bg-amber-50 px-2 py-0.5 rounded-md inline-block">
            100% Mapeados
          </div>
        </Card>

        <Card className="p-4 bg-white border-slate-200/90 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Trabalhados</span>
            <CheckSquare size={16} className="text-blue-500" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-blue-600">{kpis.trabalhados}</div>
            <p className="text-[11px] text-slate-500 font-medium">Intervenções Realizadas</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-blue-700 bg-blue-50 px-2 py-0.5 rounded-md inline-block">
            {kpis.percentualCobertura}% de Cobertura
          </div>
        </Card>

        <Card className="p-4 bg-white border-slate-200/90 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Soluções</span>
            <Zap size={16} className="text-emerald-500" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-emerald-600">{kpis.totalSolucoes}</div>
            <p className="text-[11px] text-slate-500 font-medium">Ferramentas Implantadas</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md inline-block">
            Modelos & Planilhas
          </div>
        </Card>

        <Card className="p-4 bg-white border-slate-200/90 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Atividades</span>
            <Calendar size={16} className="text-indigo-500" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-indigo-600">{kpis.totalAtividades}</div>
            <p className="text-[11px] text-slate-500 font-medium">Cronograma 30hs</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-indigo-700 bg-indigo-50 px-2 py-0.5 rounded-md inline-block">
            100% Executadas
          </div>
        </Card>

        <Card className="p-4 bg-white border-slate-200/90 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-slate-500 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Indicadores</span>
            <BarChart3 size={16} className="text-purple-500" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-purple-600">{kpis.indicadoresImplantados}</div>
            <p className="text-[11px] text-slate-500 font-medium">De {kpis.totalIndicadores} Acompanhados</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-purple-700 bg-purple-50 px-2 py-0.5 rounded-md inline-block">
            Zootécnicos & Fin.
          </div>
        </Card>

        <Card className="p-4 bg-emerald-950 text-white border-emerald-800 shadow-sm rounded-2xl flex flex-col justify-between">
          <div className="flex items-center justify-between text-emerald-300 mb-1.5">
            <span className="text-[11px] font-bold uppercase tracking-wider">Resolvidos</span>
            <ShieldCheck size={16} className="text-emerald-400" />
          </div>
          <div>
            <div className="text-2xl sm:text-3xl font-black text-emerald-400">{kpis.percentualResolvidos}%</div>
            <p className="text-[11px] text-emerald-200 font-medium">{kpis.resolvidos} de {kpis.total} Totalmente</p>
          </div>
          <div className="mt-2 text-[10px] font-bold text-slate-950 bg-emerald-400 px-2 py-0.5 rounded-md inline-block">
            {kpis.emImplantacao} em Implantação
          </div>
        </Card>
      </div>

      {/* NAVEGAÇÃO DE SUBSEÇÕES DA ABA */}
      <div className="flex flex-wrap items-center gap-2 border-b border-slate-200 pb-3 print:hidden">
        <button
          onClick={() => setActiveTabSubSection('matriz')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'matriz'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Layers size={14} /> 3. Matriz Problema × Solução
        </button>
        <button
          onClick={() => setActiveTabSubSection('trilha')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'trilha'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <TrendingUp size={14} /> 6. Análise Causa-Efeito & Trilha
        </button>
        <button
          onClick={() => setActiveTabSubSection('indicadores')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'indicadores'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <CheckSquare size={14} /> 7. Indicadores de Resultado ({kpis.totalIndicadores})
        </button>
        <button
          onClick={() => setActiveTabSubSection('graficos')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'graficos'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <BarChart3 size={14} /> 8. Gráficos de Evolução
        </button>
        <button
          onClick={() => setActiveTabSubSection('executiva')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'executiva'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <FileText size={14} /> 9 & 11. Análise Executiva & Conclusão
        </button>
        <button
          onClick={() => setActiveTabSubSection('continuidade')}
          className={`px-4 py-2 text-xs font-bold rounded-xl transition-all flex items-center gap-1.5 ${
            activeTabSubSection === 'continuidade'
              ? 'bg-slate-900 text-white shadow-md'
              : 'bg-white text-slate-600 hover:bg-slate-100 border border-slate-200'
          }`}
        >
          <Calendar size={14} /> 10. Plano de Continuidade
        </button>
      </div>

      {/* 3. MATRIZ PROBLEMA × SOLUÇÃO (TABELA PRINCIPAL) */}
      {(activeTabSubSection === 'matriz' || activeTabSubSection === 'graficos') && (
        <Card className="p-6 bg-white border-slate-200 shadow-sm rounded-2xl">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
                <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  3. MATRIZ PROBLEMA × SOLUÇÃO (CONSULTORIA GERENCIAL)
                </h3>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Mapeamento completo dos 14 problemas identificados no diagnóstico, ação executada, solução implantada e indicador de acompanhamento.
              </p>
            </div>

            {/* Filtros e Busca */}
            <div className="flex flex-wrap items-center gap-2.5 print:hidden">
              <div className="flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
                <Filter size={13} className="text-slate-400" />
                <span className="text-[11px] font-bold text-slate-500">Situação:</span>
                <select
                  value={filtroSituacao}
                  onChange={(e) => setFiltroSituacao(e.target.value)}
                  className="bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
                  <option value="Todos">Todas ({problemas.length})</option>
                  <option value="Resolvido">Resolvido ({kpis.resolvidos})</option>
                  <option value="Em implantação">Em implantação ({kpis.emImplantacao})</option>
                  <option value="Necessita acompanhamento">Necessita acompanhamento ({kpis.acompanhamento})</option>
                  <option value="Pendente">Pendente ({kpis.pendentes})</option>
                </select>
              </div>

              <div className="flex items-center gap-1.5 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
                <span className="text-[11px] font-bold text-slate-500">Área:</span>
                <select
                  value={filtroArea}
                  onChange={(e) => setFiltroArea(e.target.value)}
                  className="bg-transparent text-xs font-semibold text-slate-700 focus:outline-none cursor-pointer"
                >
                  {areasDisponiveis.map(area => (
                    <option key={area} value={area}>{area}</option>
                  ))}
                </select>
              </div>

              <input
                type="text"
                value={buscaTexto}
                onChange={(e) => setBuscaTexto(e.target.value)}
                placeholder="Buscar problema ou solução..."
                className="px-3 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-emerald-500 w-44"
              />
            </div>
          </div>

          {/* Tabela Responsiva */}
          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-900 text-white font-bold uppercase tracking-wider text-[10px]">
                  <th className="p-3.5 border-r border-slate-800 w-12 text-center">#</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[200px]">Problema Identificado</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[180px]">Impacto na Atividade</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[180px]">Ação da Consultoria</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[180px]">Solução Implementada</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[170px]">Resultado Esperado</th>
                  <th className="p-3.5 border-r border-slate-800 min-w-[170px]">Indicador de Acompanhamento</th>
                  <th className="p-3.5 min-w-[130px] text-center">Situação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-medium">
                {filteredProblemas.map((item, idx) => {
                  const isResolvido = item.situacao === 'Resolvido';
                  const isImplantacao = item.situacao === 'Em implantação';
                  const isAcompanhamento = item.situacao === 'Necessita acompanhamento';

                  return (
                    <tr 
                      key={item.id}
                      className={`hover:bg-slate-50/80 transition-colors ${
                        idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/40'
                      }`}
                    >
                      <td className="p-3 border-r border-slate-200 font-mono text-center font-bold text-slate-400">
                        {idx + 1}
                      </td>
                      <td className="p-3 border-r border-slate-200">
                        <div className="font-bold text-slate-900 mb-0.5">{item.problema}</div>
                        {item.area && (
                          <span className="inline-block text-[9px] font-bold uppercase tracking-wider px-2 py-0.5 bg-slate-100 text-slate-600 rounded">
                            {item.area}
                          </span>
                        )}
                      </td>
                      <td className="p-3 border-r border-slate-200 text-slate-600 leading-snug">
                        {item.impacto}
                      </td>
                      <td className="p-3 border-r border-slate-200 text-slate-700 leading-snug">
                        {item.acaoConsultoria}
                      </td>
                      <td className="p-3 border-r border-slate-200 font-semibold text-emerald-800 leading-snug">
                        {item.solucaoImplementada}
                      </td>
                      <td className="p-3 border-r border-slate-200 text-slate-600 leading-snug">
                        {item.resultadoEsperado}
                      </td>
                      <td className="p-3 border-r border-slate-200 text-slate-700 leading-snug font-medium">
                        {item.indicadorAcompanhamento}
                      </td>
                      <td className="p-3 text-center">
                        <select
                          value={item.situacao}
                          onChange={(e) => handleUpdateSituacaoProblema(item.id, e.target.value as any)}
                          className={`text-[11px] font-bold px-2.5 py-1.5 rounded-lg border cursor-pointer focus:outline-none transition-all ${
                            isResolvido 
                              ? 'bg-emerald-50 text-emerald-700 border-emerald-300' 
                              : isImplantacao 
                              ? 'bg-amber-50 text-amber-700 border-amber-300' 
                              : isAcompanhamento 
                              ? 'bg-purple-50 text-purple-700 border-purple-300' 
                              : 'bg-rose-50 text-rose-700 border-rose-300'
                          }`}
                        >
                          <option value="Resolvido">✔ Resolvido</option>
                          <option value="Em implantação">⏳ Em implantação</option>
                          <option value="Necessita acompanhamento">🔍 Necessita acompanhamento</option>
                          <option value="Pendente">❌ Pendente</option>
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {/* 6. ANÁLISE DOS RESULTADOS (FRAMEWORK: PROBLEMA → CAUSA → IMPACTO → INTERVENÇÃO → SOLUÇÃO → RESULTADO → PRÓXIMO PASSO) */}
      {(activeTabSubSection === 'trilha' || activeTabSubSection === 'matriz') && (
        <Card className="p-6 bg-white border-slate-200 shadow-sm rounded-2xl">
          <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="w-2.5 h-2.5 rounded-full bg-blue-500" />
                <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  6. ANÁLISE DOS RESULTADOS (TRILHA CAUSA-EFEITO POR PROBLEMA)
                </h3>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Framework estruturado: <strong className="text-slate-700">Problema → Causa → Impacto → Intervenção → Solução → Resultado → Próximo Passo</strong>
              </p>
            </div>
            <span className="text-xs font-bold text-slate-500 bg-slate-100 px-3 py-1 rounded-full">
              {problemas.length} Trilhas Analíticas
            </span>
          </div>

          <div className="space-y-4">
            {problemas.map((item, idx) => {
              const isExpanded = expandedProblemaId === item.id;
              const isResolvido = item.situacao === 'Resolvido';

              return (
                <div 
                  key={item.id}
                  className="border border-slate-200 rounded-2xl overflow-hidden bg-slate-50/50 hover:border-slate-300 transition-all"
                >
                  <button
                    onClick={() => setExpandedProblemaId(isExpanded ? null : item.id)}
                    className="w-full p-4 text-left flex items-center justify-between gap-4 bg-white hover:bg-slate-50/80 transition-colors"
                  >
                    <div className="flex items-center gap-3">
                      <span className="w-7 h-7 rounded-full bg-slate-900 text-white font-bold text-xs flex items-center justify-center font-mono">
                        {idx + 1}
                      </span>
                      <div>
                        <h4 className="text-sm font-bold text-slate-900">{item.problema}</h4>
                        <div className="flex items-center gap-2 mt-0.5">
                          <span className="text-[10px] font-bold text-slate-500 uppercase">{item.area}</span>
                          <span className="text-slate-300">•</span>
                          <span className="text-[10px] text-emerald-700 font-medium font-mono">{item.atividadeRelacionada}</span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-3">
                      <span className={`text-[11px] font-bold px-2.5 py-0.5 rounded-full border ${
                        isResolvido 
                          ? 'bg-emerald-50 text-emerald-700 border-emerald-200' 
                          : 'bg-amber-50 text-amber-700 border-amber-200'
                      }`}>
                        {item.situacao}
                      </span>
                      {isExpanded ? <ChevronUp size={18} className="text-slate-400" /> : <ChevronDown size={18} className="text-slate-400" />}
                    </div>
                  </button>

                  <AnimatePresence>
                    {isExpanded && (
                      <motion.div
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        exit={{ opacity: 0, height: 0 }}
                        className="p-5 border-t border-slate-200 bg-white"
                      >
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
                          {/* Causa */}
                          <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/70">
                            <span className="font-bold text-[10px] uppercase tracking-wider text-rose-700 block mb-1">
                              1. Causa Raiz
                            </span>
                            <p className="text-slate-700 leading-relaxed">{item.causa}</p>
                          </div>

                          {/* Impacto */}
                          <div className="p-3.5 bg-slate-50 rounded-xl border border-slate-200/70">
                            <span className="font-bold text-[10px] uppercase tracking-wider text-amber-700 block mb-1">
                              2. Impacto na Atividade
                            </span>
                            <p className="text-slate-700 leading-relaxed">{item.impacto}</p>
                          </div>

                          {/* Intervenção & Solução */}
                          <div className="p-3.5 bg-blue-50/50 rounded-xl border border-blue-200/70">
                            <span className="font-bold text-[10px] uppercase tracking-wider text-blue-700 block mb-1">
                              3. Intervenção & Solução
                            </span>
                            <p className="text-slate-800 font-semibold mb-1">{item.solucaoImplementada}</p>
                            <p className="text-slate-600 leading-relaxed text-[11px]">{item.intervencao}</p>
                          </div>

                          {/* Resultado & Próximo Passo */}
                          <div className="p-3.5 bg-emerald-50/50 rounded-xl border border-emerald-200/70">
                            <span className="font-bold text-[10px] uppercase tracking-wider text-emerald-800 block mb-1">
                              4. Resultado & Próximo Passo
                            </span>
                            <p className="text-slate-800 font-medium mb-1 leading-snug">{item.resultadoEsperado}</p>
                            <div className="mt-2 pt-2 border-t border-emerald-200/60">
                              <span className="text-[10px] font-bold text-emerald-900 uppercase block">Próximo Passo:</span>
                              <p className="text-emerald-800 text-[11px] font-semibold">{item.proximoPasso}</p>
                            </div>
                          </div>
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* 7. INDICADORES DE RESULTADO (CHECKLIST DO PRODUTOR) */}
      {(activeTabSubSection === 'indicadores' || activeTabSubSection === 'graficos') && (
        <Card className="p-6 bg-white border-slate-200 shadow-sm rounded-2xl">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="w-2.5 h-2.5 rounded-full bg-purple-500" />
                <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  7. INDICADORES DE RESULTADO E ROTINA DO PRODUTOR
                </h3>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Checklist dos 14 principais indicadores técnicos, produtivos e econômicos da Carcinicultura acompanhados pelo produtor.
              </p>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-xs font-bold px-3 py-1 bg-emerald-50 text-emerald-700 rounded-full border border-emerald-200">
                {kpis.indicadoresImplantados} Acompanhados Ativamente
              </span>
              <span className="text-xs font-bold px-3 py-1 bg-amber-50 text-amber-700 rounded-full border border-amber-200">
                {kpis.indicadoresEmProgresso} Em Implantação
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            {indicadores.map((ind, idx) => {
              const isSim = ind.status === 'Sim';
              const isEmProgresso = ind.status === 'Em implantação';

              return (
                <div
                  key={ind.id}
                  className="p-4 bg-slate-50/70 border border-slate-200 rounded-2xl flex flex-col justify-between gap-3 hover:border-slate-300 transition-all"
                >
                  <div className="flex items-start justify-between gap-2">
                    <div>
                      <div className="flex items-center gap-2 mb-1">
                        <span className="font-mono text-[10px] font-bold text-slate-400 bg-white px-1.5 py-0.5 rounded border border-slate-200">
                          IND {idx + 1}
                        </span>
                        <span className="text-[10px] font-bold uppercase tracking-wider text-slate-500">
                          {ind.categoria}
                        </span>
                      </div>
                      <h4 className="font-bold text-sm text-slate-900">{ind.nome}</h4>
                    </div>

                    <select
                      value={ind.status}
                      onChange={(e) => handleUpdateStatusIndicador(ind.id, e.target.value as any)}
                      className={`text-xs font-bold px-2.5 py-1 rounded-lg border cursor-pointer focus:outline-none transition-all shrink-0 ${
                        isSim 
                          ? 'bg-emerald-100 text-emerald-800 border-emerald-300' 
                          : isEmProgresso 
                          ? 'bg-amber-100 text-amber-800 border-amber-300' 
                          : 'bg-rose-100 text-rose-800 border-rose-300'
                      }`}
                    >
                      <option value="Sim">✔ Acompanha: Sim</option>
                      <option value="Em implantação">⏳ Em implantação</option>
                      <option value="Não">❌ Não</option>
                    </select>
                  </div>

                  <div className="grid grid-cols-2 gap-2 text-xs bg-white p-2.5 rounded-xl border border-slate-200/80">
                    <div>
                      <span className="text-[10px] font-bold text-slate-400 uppercase block">Valor Apurado</span>
                      <span className="font-black text-slate-800 text-sm font-mono">{ind.valorAtual}</span>
                    </div>
                    <div>
                      <span className="text-[10px] font-bold text-slate-400 uppercase block">Meta Estabelecida</span>
                      <span className="font-bold text-emerald-700 text-sm font-mono">{ind.meta}</span>
                    </div>
                  </div>

                  <p className="text-[11px] text-slate-500 italic leading-snug">
                    {ind.observacao}
                  </p>
                </div>
              );
            })}
          </div>
        </Card>
      )}

      {/* 8. GRÁFICOS DINÂMICOS E INTERATIVOS */}
      {(activeTabSubSection === 'graficos' || activeTabSubSection === 'matriz') && (
        <div className="space-y-6">
          <div className="flex items-center gap-2">
            <span className="w-2.5 h-2.5 rounded-full bg-emerald-500" />
            <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
              8. GRÁFICOS DE EVOLUÇÃO E DESEMPENHO DA CONSULTORIA
            </h3>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Gráfico 1 — Problemas x Soluções */}
            <Card className="p-5 bg-white border-slate-200 shadow-sm rounded-2xl">
              <div className="mb-4">
                <h4 className="font-bold text-sm text-slate-900 uppercase">Gráfico 1 — Problemas × Soluções</h4>
                <p className="text-xs text-slate-500">Total diagnosticado vs. Cobertura de intervenções e resoluções</p>
              </div>
              <div className="h-64 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartDataProblemasSolucoes} margin={{ top: 10, right: 10, left: -20, bottom: 20 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="categoria" tick={{ fontSize: 10 }} angle={-15} textAnchor="end" />
                    <YAxis allowDecimals={false} tick={{ fontSize: 11 }} />
                    <Tooltip />
                    <Bar dataKey="quantidade" radius={[6, 6, 0, 0]}>
                      {chartDataProblemasSolucoes.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={entry.fill} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Card>

            {/* Gráfico 2 — Estrutura de Custos na Carcinicultura */}
            <Card className="p-5 bg-white border-slate-200 shadow-sm rounded-2xl">
              <div className="mb-4">
                <h4 className="font-bold text-sm text-slate-900 uppercase">Gráfico 2 — Estrutura de Custos do Viveiro (%)</h4>
                <p className="text-xs text-slate-500">Participação de ração, pós-larvas, energia e mão de obra no custo total</p>
              </div>
              <div className="h-64 w-full flex items-center justify-center">
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={chartDataEstruturaCustos}
                      cx="50%"
                      cy="50%"
                      innerRadius={50}
                      outerRadius={80}
                      paddingAngle={4}
                      dataKey="value"
                      label={({ name, value }) => `${value}%`}
                    >
                      {chartDataEstruturaCustos.map((entry, index) => (
                        <Cell key={`cell-pie-${index}`} fill={entry.color} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value: any) => [`${value}%`, 'Participação']} />
                    <Legend layout="horizontal" verticalAlign="bottom" align="center" wrapperStyle={{ fontSize: '10px' }} />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </Card>

            {/* Gráfico 3 — Receita x Custos x Resultado */}
            <Card className="p-5 bg-white border-slate-200 shadow-sm rounded-2xl">
              <div className="mb-4">
                <h4 className="font-bold text-sm text-slate-900 uppercase">Gráfico 3 — Receita × Custos × Lucro do Ciclo (R$)</h4>
                <p className="text-xs text-slate-500">Evolução do faturamento, custos totais e sobra líquida do produtor</p>
              </div>
              <div className="h-64 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartDataReceitaCustos} margin={{ top: 10, right: 10, left: 10, bottom: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                    <YAxis tick={{ fontSize: 10 }} tickFormatter={(val) => `R$${val/1000}k`} />
                    <Tooltip formatter={(value: any) => [`R$ ${Number(value).toLocaleString('pt-BR')}`, '']} />
                    <Legend wrapperStyle={{ fontSize: '11px' }} />
                    <Bar dataKey="Receita" fill="#0284c7" name="Receita Total" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="CustoTotal" fill="#ef4444" name="Custo Total" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="LucroLiquido" fill="#10b981" name="Lucro Líquido" radius={[4, 4, 0, 0]} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </Card>

            {/* Gráfico 4 — Indicadores Produtivos (FCA, Sobrevivência, Custo/kg) */}
            <Card className="p-5 bg-white border-slate-200 shadow-sm rounded-2xl">
              <div className="mb-4">
                <h4 className="font-bold text-sm text-slate-900 uppercase">Gráfico 4 — Evolução Produtiva & Custo Unitário</h4>
                <p className="text-xs text-slate-500">Trajetória de redução do custo por kg (R$/kg) e melhoria do FCA</p>
              </div>
              <div className="h-64 w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={chartDataIndicadoresProdutivos} margin={{ top: 10, right: 10, left: -10, bottom: 10 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis dataKey="periodo" tick={{ fontSize: 10 }} />
                    <YAxis yAxisId="left" tick={{ fontSize: 10 }} />
                    <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 10 }} />
                    <Tooltip />
                    <Legend wrapperStyle={{ fontSize: '11px' }} />
                    <Bar yAxisId="left" dataKey="producaoKg" fill="#3b82f6" name="Produção (kg)" radius={[4, 4, 0, 0]} />
                    <Line yAxisId="right" type="monotone" dataKey="custoKg" stroke="#ef4444" strokeWidth={2.5} name="Custo R$/kg" />
                    <Line yAxisId="right" type="monotone" dataKey="fca" stroke="#10b981" strokeWidth={2.5} name="FCA (kg/kg)" />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </Card>
          </div>
        </div>
      )}

      {/* 9. ANÁLISE EXECUTIVA AUTOMATIZADA */}
      {(activeTabSubSection === 'executiva' || activeTabSubSection === 'matriz') && (
        <Card className="p-6 bg-white border-slate-200 shadow-sm rounded-2xl">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="w-2.5 h-2.5 rounded-full bg-slate-900" />
                <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  9. PARECER TÉCNICO & ANÁLISE EXECUTIVA FINAL
                </h3>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Síntese técnica estruturada respondendo aos 8 pilares da consultoria gerencial SEBRAE.
              </p>
            </div>

            <div className="flex items-center gap-2 print:hidden">
              <Button
                variant="outline"
                size="sm"
                onClick={() => setIsEditingAnaliseExecutiva(!isEditingAnaliseExecutiva)}
                className="text-xs font-bold border-slate-300 text-slate-700"
              >
                <Edit3 size={14} className="mr-1.5 text-blue-600" />
                {isEditingAnaliseExecutiva ? 'Salvar Edição' : 'Editar Texto'}
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => setAnaliseExecutivaTexto(defaultAnaliseExecutivaTexto)}
                className="text-xs font-medium text-slate-500 hover:text-slate-800"
                title="Restaurar texto padrão baseado nos dados"
              >
                <RefreshCw size={13} className="mr-1" /> Restaurar Padrão
              </Button>
            </div>
          </div>

          {isEditingAnaliseExecutiva ? (
            <textarea
              value={analiseExecutivaTexto}
              onChange={(e) => setAnaliseExecutivaTexto(e.target.value)}
              rows={16}
              className="w-full p-4 border border-slate-300 rounded-xl text-xs font-mono leading-relaxed focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 focus:outline-none bg-white text-slate-800"
            />
          ) : (
            <div className="p-5 bg-slate-50 rounded-xl border border-slate-200/80 text-xs text-slate-800 leading-relaxed whitespace-pre-line font-medium font-sans">
              {analiseExecutivaTexto}
            </div>
          )}
        </Card>
      )}

      {/* 10. PLANO DE CONTINUIDADE PÓS-CONSULTORIA */}
      {(activeTabSubSection === 'continuidade' || activeTabSubSection === 'matriz') && (
        <Card className="p-6 bg-white border-slate-200 shadow-sm rounded-2xl">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2 mb-1">
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-600" />
                <h3 className="text-lg font-black text-slate-800 uppercase tracking-tight">
                  10. PLANO DE CONTINUIDADE PÓS-CONSULTORIA
                </h3>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Matriz de rotinas e responsabilidades para consolidação contínua das ferramentas implantadas.
              </p>
            </div>

            <span className="text-xs font-bold px-3 py-1 bg-emerald-100 text-emerald-800 rounded-full">
              {planoContinuidade.length} Ações de Rotina Estabelecidas
            </span>
          </div>

          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-emerald-950 text-white font-bold uppercase tracking-wider text-[10px]">
                  <th className="p-3.5 border-r border-emerald-900 w-12 text-center">#</th>
                  <th className="p-3.5 border-r border-emerald-900 min-w-[260px]">Ação de Gestão / Rotina</th>
                  <th className="p-3.5 border-r border-emerald-900 min-w-[180px]">Frequência</th>
                  <th className="p-3.5 border-r border-emerald-900 min-w-[160px]">Responsável</th>
                  <th className="p-3.5 border-r border-emerald-900 min-w-[180px]">Indicador de Acompanhamento</th>
                  <th className="p-3.5 min-w-[150px]">Meta Estabelecida</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-medium">
                {planoContinuidade.map((pl, idx) => (
                  <tr key={pl.id} className={idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/50'}>
                    <td className="p-3 border-r border-slate-200 font-mono text-center font-bold text-slate-400">
                      {idx + 1}
                    </td>
                    <td className="p-3 border-r border-slate-200 font-bold text-slate-900">
                      {pl.acao}
                    </td>
                    <td className="p-3 border-r border-slate-200 text-emerald-800 font-semibold">
                      {pl.frequencia}
                    </td>
                    <td className="p-3 border-r border-slate-200 text-slate-700">
                      {pl.responsavel}
                    </td>
                    <td className="p-3 border-r border-slate-200 text-slate-600">
                      {pl.indicador}
                    </td>
                    <td className="p-3 font-semibold text-emerald-700">
                      {pl.meta}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {/* 11. CONCLUSÃO PROFISSIONAL */}
      <Card className="p-6 bg-gradient-to-br from-slate-900 via-slate-800 to-emerald-950 text-white shadow-lg rounded-2xl border border-slate-700/60 print:bg-white print:text-slate-900 print:border print:border-slate-300">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4 pb-4 border-b border-slate-700/60 print:border-slate-200">
          <div>
            <div className="flex items-center gap-2 mb-1">
              <span className="w-2.5 h-2.5 rounded-full bg-emerald-400" />
              <h3 className="text-lg font-black text-white uppercase tracking-tight print:text-slate-900">
                11. CONCLUSÃO E ENCERRAMENTO DA CONSULTORIA
              </h3>
            </div>
            <p className="text-xs text-slate-300 font-medium print:text-slate-500">
              Síntese de encerramento destacando a autonomia gerencial do produtor rural.
            </p>
          </div>

          <div className="print:hidden">
            <Button
              variant="outline"
              size="sm"
              onClick={() => setIsEditingConclusao(!isEditingConclusao)}
              className="text-xs font-bold border-slate-600 text-slate-200 hover:bg-slate-700"
            >
              <Edit3 size={14} className="mr-1.5 text-emerald-400" />
              {isEditingConclusao ? 'Salvar Conclusão' : 'Editar Conclusão'}
            </Button>
          </div>
        </div>

        {isEditingConclusao ? (
          <textarea
            value={conclusaoTexto}
            onChange={(e) => setConclusaoTexto(e.target.value)}
            rows={6}
            className="w-full p-4 bg-slate-800 text-slate-100 border border-slate-600 rounded-xl text-xs leading-relaxed focus:ring-2 focus:ring-emerald-500/30 focus:outline-none"
          />
        ) : (
          <div className="text-xs text-slate-200 leading-relaxed font-medium space-y-3 print:text-slate-800">
            <p className="whitespace-pre-line">{conclusaoTexto}</p>
          </div>
        )}

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 mt-8 pt-6 border-t border-slate-700/60 text-center text-xs print:border-slate-300">
          <div className="space-y-1">
            <div className="w-48 h-px bg-slate-600 mx-auto mb-2 print:bg-slate-400" />
            <p className="font-bold text-slate-200 print:text-slate-900">{consultorNome}</p>
            <p className="text-[11px] text-slate-400 print:text-slate-600">Consultor Especialista em Carcinicultura / SEBRAE</p>
          </div>
          <div className="space-y-1">
            <div className="w-48 h-px bg-slate-600 mx-auto mb-2 print:bg-slate-400" />
            <p className="font-bold text-slate-200 print:text-slate-900">{clienteNome}</p>
            <p className="text-[11px] text-slate-400 print:text-slate-600">Representante do Empreendimento / Produtor</p>
          </div>
        </div>
      </Card>
    </div>
  );
};
