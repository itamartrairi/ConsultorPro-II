// Constantes e Modelos de Consultoria Estratégica

export const CONSULTORIA_AREAS = [
  "FINANÇAS",
  "MARKETING E VENDAS",
  "PLANEJAMENTO ESTRATÉGICO",
  "GESTÃO DE PESSOAS",
  "PROCESSOS E OPERAÇÕES",
  "INOVAÇÃO E TECNOLOGIA",
  "JURÍDICO",
  "SUSTENTABILIDADE",
  "EMPREENDEDORISMO",
  "ACESSO A CRÉDITO",
  "CRÉDITO",
  "OUTROS"
];

export const AREAS = [
  { id: 'FIN', nome: 'Financeiro' },
  { id: 'MKT', nome: 'Marketing' },
  { id: 'OPS', nome: 'Operacional' },
  { id: 'RH',  nome: 'Recursos Humanos' },
  { id: 'VEN', nome: 'Vendas' },
  { id: 'TEC', nome: 'Tecnologia' },
  { id: 'JUR', nome: 'Jurídico' },
  { id: 'EST', nome: 'Estratégico' },
  { id: 'LOG', nome: 'Logística' },
  { id: 'SAC', nome: 'Atendimento' },
  { id: 'CRE', nome: 'Acesso a Crédito' },
  { id: 'CRD', nome: 'Crédito' },
];

export const IMPACTO_ORDER: Record<string, number> = { 'Baixo': 1, 'Médio': 2, 'Alto': 3 };

export interface AtividadeCronograma {
  id?: string;
  nome: string;
  descricao?: string;
  cargaHoraria: string;
  solucaoProposta: string;
  responsavel: string;
  status: 'Pendente' | 'Em Andamento' | 'Concluída' | 'Cancelada';
  prioridade: 'Alta' | 'Média' | 'Baixa';
  resultadoEsperado?: string;
  dataInicio?: any;
  dataFim?: any;
  ordem?: number;
  observacoes?: string;
}

export interface ModeloRelatorio {
  id: string;
  nome: string;
  categoria: string;
  descricao: string;
  objetivo: string;
  atividades: AtividadeCronograma[];
}

export const PLANO_DE_ACAO_PADRAO: AtividadeCronograma[] = [
  {
    nome: "Entendimento da demanda e diagnóstico inicial",
    descricao: "Alinhamento das expectativas do cliente e levantamento de dados operacionais e financeiros atuais.",
    cargaHoraria: "4h",
    solucaoProposta: "Entendimento da demanda e diagnóstico inicial",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Expectativas alinhadas e diagnóstico operacional/financeiro inicial consolidado."
  },
  {
    nome: "Estrutura do Controle de Custos Operacionais",
    descricao: "Implementação de métodos para registro rigoroso de insumos e despesas operacionais.",
    cargaHoraria: "4h",
    solucaoProposta: "Estruturação de registros e controles de custos operacionais.",
    responsavel: "Consultor/Produtor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Controle de insumos estruturado e planilhas implantadas na rotina do cliente."
  },
  {
    nome: "Cálculo do Custo de Produção e Precificação",
    descricao: "Apuração técnica do custo real unitário e definição de margem de contribuição.",
    cargaHoraria: "4h",
    solucaoProposta: "Análise técnica de custos e apuração do preço de venda ideal.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Custo unitário e margem de contribuição calculados com precisão matemática."
  },
  {
    nome: "Introdução à Gestão Financeira e Fluxo de Caixa",
    descricao: "Capacitação sobre controle financeiro diário e elaboração de DRE simplificado.",
    cargaHoraria: "4h",
    solucaoProposta: "Treinamento sobre fluxo de caixa e demonstrativo de resultado.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Média",
    resultadoEsperado: "Fluxo de caixa implantado e rotina de registros mantida."
  },
  {
    nome: "Planejamento Estratégico e Gestão de Estoque",
    descricao: "Definição de metas de produção/vendas e controle de insumos e matérias-primas.",
    cargaHoraria: "4h",
    solucaoProposta: "Formulação de metas operacionais e otimização da gestão de estoque.",
    responsavel: "Consultor/Produtor",
    status: "Pendente",
    prioridade: "Média",
    resultadoEsperado: "Metas estratégicas e controle de estoque estruturados para os próximos meses."
  },
  {
    nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
    descricao: "Análise e desenvolvimento de sistema de gestão, aplicativo ou planilhas personalizadas para automação dos controles da empresa.",
    cargaHoraria: "8h",
    solucaoProposta: "Desenvolvimento de aplicativo e planilhas integradas de gestão",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Sistema, aplicativo ou planilhas operacionais implantadas no ambiente do cliente."
  },
  {
    nome: "Análise final de atividades da consultoria",
    descricao: "Análise de encerramento e verificação do alcance dos indicadores e resultados previstos no plano de trabalho.",
    cargaHoraria: "2h",
    solucaoProposta: "Análise final e consolidação de resultados da consultoria",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Média",
    resultadoEsperado: "Verificação completa do cumprimento do escopo e metas da consultoria."
  },
  {
    nome: "Relatório final e encerramento",
    descricao: "Consolidação dos resultados e avaliação do impacto da consultoria.",
    cargaHoraria: "4h",
    solucaoProposta: "Elaboração e apresentação do relatório final.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Média",
    resultadoEsperado: "Relatório final consolidado, apresentado e validado pelo produtor/cliente."
  }
];

interface ModeloRelatorio {
  id: string;
  nome: string;
  categoria: string;
  descricao: string;
  objetivo?: string;
  atividades: AtividadeCronograma[];
}


export const MODELOS_RELATORIO: ModeloRelatorio[] = [
  {
    id: 'carcinicultura',
    nome: 'Gestão Financeira & Controle de Custos da Carcinicultura',
    categoria: 'Carcinicultura',
    descricao: 'Modelo de Consultoria Gerencial para Carcinicultura (30hs): Entendimento de demanda, Diagnóstico de fluxos, Planilha de custos, Reserva de emergência, Treinamento de indicadores, Segregação de contas e Relatório final.',
    objetivo: 'Implementar melhorias estratégicas e práticas nas áreas de Gestão Financeira, Preço e Mercado, Planejamento, Custos e Resultado, Acesso a Crédito, Controle e Rotina e Produção, com base no diagnóstico realizado, visando aprimorar a gestão do empreendimento, fortalecer os controles gerenciais, melhorar a tomada de decisão, aumentar a eficiência dos processos produtivos e financeiros e contribuir para a sustentabilidade e melhoria dos resultados',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Alinhamento das expectativas do cliente e levantamento detalhado das necessidades operacionais e financeiras.",
        cargaHoraria: "2h",
        solucaoProposta: "Alinhamento das expectativas e levantamento das necessidades operacionais e financeiras da carcinicultura.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Expectativas alinhadas e necessidades operacionais e financeiras levantadas."
      },
      {
        nome: "Diagnóstico Detalhado e Levantamento de Fluxos",
        descricao: "Reunião inicial para mapeamento dos processos atuais, verificação de registros de caixa e entendimento das falhas de separação entre contas.",
        cargaHoraria: "4h",
        solucaoProposta: "Mapeamento dos processos atuais e verificação das falhas de separação entre contas.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxos atuais mapeados e falhas de controle financeiro identificadas."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Criação e implementação de uma planilha de controle de custos de ração, receitas e cálculo de margem de contribuição.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento e implantação da planilha de controle de custos de ração, receitas e margem de contribuição.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Planilha/ferramenta de gestão operando no cliente para controle de custos e receitas."
      },
      {
        nome: "Planejamento Financeiro e Reserva de Emergência",
        descricao: "Criação de um plano para constituição de reserva financeira visando a sustentabilidade do próximo ciclo.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração de plano financeiro para formação de reserva de emergência e sustentabilidade do próximo ciclo.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Reserva financeira planejada e metas de fluxo de caixa para o próximo ciclo estruturadas."
      },
      {
        nome: "Treinamento de Gestão e Controle de Indicadores",
        descricao: "Capacitação sobre o uso do caderno de campo/planilha e acompanhamento dos indicadores de produtividade.",
        cargaHoraria: "4h",
        solucaoProposta: "Capacitação do produtor no uso do caderno de campo/planilha e acompanhamento dos indicadores.",
        responsavel: "Consultor (capacitação) / Produtor (treinamento)",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Produtor capacitado no registro diário e interpretação dos indicadores de produtividade."
      },
      {
        nome: "Segregação de Contas e Estruturação Financeira",
        descricao: "Orientação para abertura de conta jurídica ou separação formal das finanças pessoais e da carcinicultura.",
        cargaHoraria: "4h",
        solucaoProposta: "Separação formal das contas pessoais e da atividade de carcinicultura com definição de pró-labore.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Finanças da carcinicultura segregadas do caixa familiar com limite de retiradas."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Revisão dos indicadores apurados na ferramenta de gestão implementada ao longo do projeto.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise e revisão de todos os indicadores apurados com as ferramentas implementadas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Indicadores revisados e evolução do desempenho financeiro confirmada."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Entrega do plano de ação contínuo e consolidação do aprendizado do produtor.",
        cargaHoraria: "2h",
        solucaoProposta: "Entrega e apresentação do relatório final da consultoria com plano de continuidade.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final entregue e produtor autônomo na gestão do empreendimento."
      }
    ]
  },
  {
    id: 'financas',
    nome: 'Gestão Financeira & Controle de Custos',
    categoria: 'Finanças',
    descricao: 'Apuração de custos, fluxo de caixa, DRE gerencial, margem de contribuição e precificação.',
    objetivo: 'Implementar melhorias na Gestão Financeira, Estruturação de Custos, Fluxo de Caixa, DRE Gerencial e Formação do Preço de Venda com base no diagnóstico realizado, visando aprimorar a rentabilidade e a sustentabilidade financeira do empreendimento.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Alinhamento das expectativas do cliente e levantamento de dados operacionais e financeiros atuais.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico inicial",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Expectativas alinhadas e diagnóstico financeiro/operacional inicial consolidado."
      },
      {
        nome: "Mapeamento e Classificação de Custos e Despesas",
        descricao: "Categorização detalhada de custos fixos, variáveis, despesas operacionais e tributos.",
        cargaHoraria: "4h",
        solucaoProposta: "Estruturação do plano de contas e classificação de custos.",
        responsavel: "Consultor/Cliente",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Plano de contas gerencial estruturado para controle financeiro."
      },
      {
        nome: "Cálculo da Margem de Contribuição e Ponto de Equilíbrio",
        descricao: "Apuração técnica da margem de contribuição por produto/serviço e determinação do faturamento mínimo.",
        cargaHoraria: "4h",
        solucaoProposta: "Análise de viabilidade financeira e cálculo do ponto de equilíbrio.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ponto de equilíbrio financeiro e margem de contribuição mensurados."
      },
      {
        nome: "Implementação e Rotina do Fluxo de Caixa Diário",
        descricao: "Treinamento e implantação de processo sistemático de controle de entradas e saídas.",
        cargaHoraria: "4h",
        solucaoProposta: "Capacitação e rotina prática de fluxo de caixa.",
        responsavel: "Consultor/Cliente",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxo de caixa projetado e acompanhado diariamente pelo cliente."
      },
      {
        nome: "Elaboração do DRE Gerencial e Formação do Preço de Venda",
        descricao: "Estruturação do Demonstrativo de Resultado e fórmula de precificação lucrativa.",
        cargaHoraria: "4h",
        solucaoProposta: "Implantação do DRE gerencial e modelo de precificação.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DRE gerencial mensal e tabela de preços revisada com margem real."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento e personalização de planilhas financeiras integradas ou aplicativo de gestão de caixa.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e planilhas integradas de gestão financeira.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta/aplicativo financeiro personalizado operando no cliente."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Análise de encerramento e verificação do alcance dos indicadores financeiros.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise final e consolidação dos resultados da consultoria.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório de desempenho financeiro e cumprimento de metas."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Consolidação final das entregas financeiras, apresentação dos resultados e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final da consultoria.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final aprovado e plano de continuidade entregue ao cliente."
      }
    ]
  },
  {
    id: 'operacoes',
    nome: 'Processos Operacionais & Produtividade',
    categoria: 'Operações',
    descricao: 'Mapeamento de processos, eliminação de gargalos, padronização POP e qualidade.',
    objetivo: 'Mapear, padronizar e otimizar os processos operacionais e produtivos com base no diagnóstico realizado, visando a eliminação de desperdícios, redução de gargalos e aumento da produtividade e eficiência da empresa.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Levantamento das etapas do fluxo produtivo/operacional e identificação de gargalos.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico operacional inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Diagnóstico operacional detalhado e gargalos priorizados."
      },
      {
        nome: "Mapeamento do Fluxo de Trabalho (AS-IS)",
        descricao: "Desenho e documentação do fluxo atual de produção e atendimento.",
        cargaHoraria: "4h",
        solucaoProposta: "Mapeamento completo dos processos operacionais vigentes.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxograma operacional atual mapeado e gargalos evidenciados."
      },
      {
        nome: "Redesenho de Processos e Redução de Desperdícios (TO-BE)",
        descricao: "Proposição de melhorias, eliminação de etapas redundantes e tempos mortos.",
        cargaHoraria: "4h",
        solucaoProposta: "Otimização de processos operacionais e redução de perdas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Novo fluxo de trabalho otimizado e aprovado pela gerência."
      },
      {
        nome: "Elaboração de Procedimentos Operacionais Padrão (POPs)",
        descricao: "Criação de manuais práticos e instruções de trabalho para padronização.",
        cargaHoraria: "4h",
        solucaoProposta: "Padronização operacional via POPs e checklists de controle.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Procedimentos operacionais padrão (POPs) redigidos e implantados."
      },
      {
        nome: "Treinamento da Equipe e Indicadores de Produtividade",
        descricao: "Capacitação dos colaboradores nos novos processos e métricas de eficiência.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento operacional e implantação de KPIs de produção.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Equipe treinada e indicadores de qualidade/produtividade ativos."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Criação de ferramenta/aplicativo ou planilhas para acompanhamento em tempo real.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e planilhas de controle operacional.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Sistema/planilha de gestão operacional em pleno funcionamento."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Avaliação do aumento de produtividade e alcance dos padrões previstos.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise de encerramento e validação de KPIs operacionais.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Ganhos de produtividade e eficiência operacional comprovados."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Consolidação do relatório técnico de processos, apresentação de resultados e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final de processos.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final entregue e operacionais padronizados mantidos."
      }
    ]
  },
  {
    id: 'comercial',
    nome: 'Comercialização, Vendas & Marketing Digital',
    categoria: 'Comercial',
    descricao: 'Estratégia de vendas, funil de captação, presença digital e fidelização.',
    objetivo: 'Estruturar o funil de vendas, posicionamento de mercado, estratégias de comunicação e capacitação comercial com base no diagnóstico realizado, visando a atração de clientes e o aumento sustentável do faturamento.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Diagnóstico das vendas atuais, perfil do cliente ideal (ICP) e canais de vendas.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico comercial inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Diagnóstico das metas de vendas e oportunidades de mercado mapeadas."
      },
      {
        nome: "Análise da Concorrência e Posicionamento de Mercado",
        descricao: "Estudo dos diferenciais competitivos e ajuste na proposta de valor do negócio.",
        cargaHoraria: "4h",
        solucaoProposta: "Definição de posicionamento estratégico e proposta de valor.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Proposta de valor diferenciada e alinhada às necessidades do cliente."
      },
      {
        nome: "Estruturação do Funil de Vendas e Processo Comercial",
        descricao: "Definição das etapas de prospecção, qualificação, apresentação e fechamento.",
        cargaHoraria: "4h",
        solucaoProposta: "Desenvolvimento da jornada do cliente e funil comercial.",
        responsavel: "Consultor/Vendedores",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Funil de vendas padronizado com rotina diária de abordagens."
      },
      {
        nome: "Estratégia de Marketing Digital e Redes Sociais",
        descricao: "Planejamento de conteúdos, anúncios e presença nas mídias sociais para atração.",
        cargaHoraria: "4h",
        solucaoProposta: "Plano de comunicação digital e atração de novos clientes.",
        responsavel: "Consultor/Marketing",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Calendário de marketing digital e diretrizes de divulgação ativas."
      },
      {
        nome: "Treinamento de Técnicas de Vendas e Pós-Venda",
        descricao: "Capacitação da equipe em contorno de objeções, fechamento e retenção.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento prático de vendas e retenção de clientes.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Equipe comercial capacitada e taxa de conversão aumentada."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de planilha/aplicativo CRM de controle de leads, propostas e histórico.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de CRM e planilhas de gestão de vendas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta CRM personalizada implantada para controle comercial."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Análise da evolução do faturamento e alcance das metas de vendas.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise final do desempenho comercial e taxas de conversão.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Aumento das vendas mensurado e relatório de métricas concluído."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Apresentação do relatório final de vendas, recomendações futuras e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final comercial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final entregue e rotina comercial consolidada."
      }
    ]
  },
  {
    id: 'tecnologia',
    nome: 'Inovação, Tecnologia & Automação de Processos',
    categoria: 'Inovação',
    descricao: 'Transformação digital, escolha de softwares, automação e treinamento em TI.',
    objetivo: 'Implementar soluções de inovação, automação de processos, integração de ferramentas e capacitação tecnológica com base no diagnóstico realizado, visando a modernização digital e o ganho de eficiência operacional.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Avaliação do nível de maturidade digital e requisitos de tecnologia.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico tecnológico inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Maturidade tecnológica mapeada e plano de inovação priorizado."
      },
      {
        nome: "Mapeamento de Requisitos e Seleção de Tecnologias",
        descricao: "Identificação das soluções tecnológicas, softwares e sistemas adequados.",
        cargaHoraria: "4h",
        solucaoProposta: "Seleção de ferramentas de automação e tecnologia de gestão.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Stack tecnológico selecionado de acordo com custo-benefício."
      },
      {
        nome: "Planejamento da Integração e Automação de Dados",
        descricao: "Desenho da arquitetura de integração entre sistemas e automação de fluxos.",
        cargaHoraria: "4h",
        solucaoProposta: "Automação de processos operacionais e integração de dados.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxos de automação planejados para reduzir digitação manual."
      },
      {
        nome: "Configuração de Parâmetros e Testes do Sistema",
        descricao: "Ajuste de cadastros, parametrização de regras de negócio e validação dos sistemas.",
        cargaHoraria: "4h",
        solucaoProposta: "Configuração técnica e simulação de rotinas digitais.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ambiente tecnológico configurado e aprovado em ambiente de teste."
      },
      {
        nome: "Capacitação da Equipe em Ferramentas Digitais",
        descricao: "Treinamento prático dos usuários para operação fluida das novas plataformas.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento técnico e mudança cultural para o meio digital.",
        responsavel: "Consultor/Usuários",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Usuários autônomos na utilização das soluções tecnológicas."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de aplicativos sob medida, dashboards interativos ou planilhas.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e integração de planilhas/sistemas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Aplicativo/sistema entregue e funcionando perfeitamente."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Análise da redução de tempo de execução e estabilidade do ambiente.",
        cargaHoraria: "2h",
        solucaoProposta: "Validação final das automações e tempo economizado.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Ganho de tempo e precisão dos dados automatizados confirmados."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Apresentação do relatório final de tecnologia, plano de suporte e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final de inovação.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório técnico final aprovado e encerramento concluído."
      }
    ]
  },
  {
    id: 'estrategia',
    nome: 'Planejamento Estratégico & Gestão de Pessoas',
    categoria: 'Estratégia',
    descricao: 'Visão de futuro, metas SWOT, organograma, liderança e KPIs.',
    objetivo: 'Definir diretrizes estratégicas, metas organizacionais, estruturação de papéis e desenvolvimento de liderança com base no diagnóstico realizado, visando o crescimento estruturado e a sustentabilidade do negócio.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Análise da cultura organizacional, liderança e desafios estratégicos.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico estratégico inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Diagnóstico da visão do negócio e direcionamento estratégico."
      },
      {
        nome: "Análise SWOT e Diretrizes Estratégicas",
        descricao: "Mapeamento de Forças, Oportunidades, Fraquezas e Ameaças.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração da Matriz SWOT e visão de crescimento.",
        responsavel: "Consultor/Sócios",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Matriz SWOT consolidada com metas de médio/longo prazo."
      },
      {
        nome: "Estruturação de Organograma e Responsabilidades",
        descricao: "Definição do quadro de funções, atribuições de cargos e matriz de responsabilidade.",
        cargaHoraria: "4h",
        solucaoProposta: "Definição de papéis, organograma e matriz de competências.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Organograma funcional claro e responsabilidades bem delineadas."
      },
      {
        nome: "Desenvolvimento de Lideranças e Comunicação Interna",
        descricao: "Treinamento em gestão de equipes, alinhamento de metas e comunicação.",
        cargaHoraria: "4h",
        solucaoProposta: "Desenvolvimento de competências gerenciais e clima organizacional.",
        responsavel: "Consultor/Gestores",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Gestores alinhados e canais de comunicação interna ativos."
      },
      {
        nome: "Plano de Metas e Acompanhamento de KPIs",
        descricao: "Definição de indicadores-chave de desempenho por setor e acompanhamento.",
        cargaHoraria: "4h",
        solucaoProposta: "Estruturação do painel de controle e acompanhamento de metas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Metas corporativas desdobradas em planos operacionais."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Criação de aplicativo ou planilha de gestão estratégica e controle de KPIs.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo/dashboard estratégico.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Painel estratégico de indicadores implantado no cliente."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Análise do alinhamento estratégico e engajamento da equipe.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise final do atingimento de metas estratégicas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Evolução do engajamento e alcance de diretrizes mensurado."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Apresentação do relatório final de planejamento estratégico e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final estratégico.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final entregue e diretrizes estratégicas vigentes."
      }
    ]
  },
  {
    id: 'agronegocio',
    nome: 'Agronegócio & Gestão de Propriedade Rural',
    categoria: 'Agronegócio',
    descricao: 'Apuração de custos por lote/safra, manejo, estoques de insumos e gestão rural.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnóstico inicial",
        descricao: "Levantamento das características da propriedade rural, cultivos/criações e infraestrutura.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnóstico agropecuário inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Diagnóstico da propriedade rural e pontos críticos de melhoria."
      },
      {
        nome: "Mapeamento dos Custos de Produção por Lote/Safra",
        descricao: "Estruturação da apuração do custo operacional efetivo e custo total por área.",
        cargaHoraria: "4h",
        solucaoProposta: "Análise detalhada do custo de produção agropecuária.",
        responsavel: "Consultor/Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Custo por safra/lote/saco mensurado com precisão."
      },
      {
        nome: "Controle de Estoque de Insumos e Calendário Manejo",
        descricao: "Organização do almoxarifado agrícola, rastreabilidade de aplicação e insumos.",
        cargaHoraria: "4h",
        solucaoProposta: "Implantação de controle de insumos e manejo operacional.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Estoque de insumos controlado e perdas reduzidas."
      },
      {
        nome: "Análise de Viabilidade Financeira e Preço de Comercialização",
        descricao: "Determinação do valor de equilíbrio da produção e estratégias de comercialização.",
        cargaHoraria: "4h",
        solucaoProposta: "Estudo de margens de lucro e comercialização safra.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Preço de venda ideal e rentabilidade da produção calculados."
      },
      {
        nome: "Capacitação no Gerenciamento Prático da Propriedade",
        descricao: "Treinamento do produtor rural e equipe no preenchimento de cadernos de campo.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento de gestão no campo e controles gerenciais.",
        responsavel: "Consultor/Produtor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Produtor autônomo na gestão financeira do agronegócio."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de aplicativo de gestão rural ou planilha agrícola customizada.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo/planilhas de gestão agropecuária.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta de gestão rural implantada na propriedade."
      },
      {
        nome: "Análise final de atividades da consultoria",
        descricao: "Verificação dos resultados do ciclo produtivo e alcance dos indicadores rurais.",
        cargaHoraria: "2h",
        solucaoProposta: "Análise final do desempenho produtivo e financeiro da propriedade.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório de desempenho da safra/lote consolidado."
      },
      {
        nome: "Relatório final e encerramento",
        descricao: "Apresentação do relatório final de gestão agropecuária e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "Elaboração e apresentação do relatório final do agronegócio.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Média",
        resultadoEsperado: "Relatório final entregue com plano de continuidade rural."
      }
    ]
  }
];



export const TIPOS_EMPRESA = [
  "Geral",
  "Comércio",
  "Serviços",
  "Indústria",
  "Agronegócio",
  "Carcinicultura",
  "Tecnologia / Startup",
  "Alimentação & Gastronomia"
];

export const AREAS_ORDER = [
  "Estratégia",
  "Processos",
  "Marketing & Vendas",
  "Financeiro",
  "Recursos Humanos",
  "Tecnologia & IA",
  "Inovação",
  "Gestão da Produção",
  "Qualidade",
  "Sustentabilidade & ESG",
  "Jurídico & Compliance",
  "Acesso a Crédito",
  "Geral"
];
