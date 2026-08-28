const fs = require('fs');

let content = fs.readFileSync('src/components/AnaliseResultadosView.tsx', 'utf8');

// 1. Add Trash2 and Plus to lucide-react imports if missing
if (!content.includes('Trash2,')) {
  content = content.replace(
    "import {\n  Award,",
    "import {\n  Award,\n  Trash2,\n  Plus,"
  );
}

// 2. Add deduplication and extraction helpers before the component
const helperCode = `
// Helper to deduplicate answers by question/id
const deduplicateRespostas = (respostasList: any[]): any[] => {
  if (!respostasList || !Array.isArray(respostasList)) return [];
  const map = new Map<string, any>();

  for (const resp of respostasList) {
    if (!resp) continue;
    const normQ = (resp.pergunta || '').trim().toLowerCase();
    const key = normQ ? \`q:\${normQ}\` : (resp.premissaId ? \`id:\${resp.premissaId}\` : resp.id);
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
      id: \`prob_\${idx}_\${Date.now()}_\${Math.random().toString(36).substring(2, 6)}\`,
      problema: p.problema,
      area: p.area,
      status: 'Resolvido' as const,
      percentualEficacia: p.hasNao ? 90 : 95,
      impactoGerado: \`Processo e controle estruturados durante a consultoria para sanar a fragilidade diagnosticada em "\${p.problema}".\`,
      evidencias: 'Planilha / rotina gerencial implementada, acompanhada e validada nas sessões de consultoria.'
    }));
  }

  // Fallback defaults if no negative answers
  return [
    {
      id: \`prob_def_0\`,
      problema: 'Ausência de controle diário e rigoroso de entradas e saídas de caixa',
      area: 'Gestão Financeira',
      status: 'Resolvido',
      percentualEficacia: 92,
      impactoGerado: 'Rotina de fluxo de caixa diário implantada e em operação no dia a dia com conciliação.',
      evidencias: 'Planilha gerencial de caixa e conciliação bancária semanal conferida nas sessões.'
    },
    {
      id: \`prob_def_1\`,
      problema: 'Mistura de despesas pessoais dos sócios com as da empresa (falta de pró-labore fixo)',
      area: 'Controle e Rotina',
      status: 'Resolvido',
      percentualEficacia: 95,
      impactoGerado: 'Separação total de contas PF e PJ e pró-labore estipulado formalmente.',
      evidencias: 'Contas bancárias segregadas e retiradas limitadas ao pró-labore estipulado.'
    },
    {
      id: \`prob_def_2\`,
      problema: 'Desconhecimento dos custos operacionais reais e da margem de contribuição',
      area: 'Custos e Resultado',
      status: 'Resolvido',
      percentualEficacia: 90,
      impactoGerado: 'Estrutura detalhada de custos fixos, variáveis e apuração da margem de lucro.',
      evidencias: 'Planilha de custos e ponto de equilíbrio calculados e apresentados.'
    },
    {
      id: \`prob_def_3\`,
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
      id: \`sol_\${idx}_\${Date.now()}_\${Math.random().toString(36).substring(2, 6)}\`,
      solucao: atv.solucaoProposta || atv.nome || \`Atividade de Consultoria \${idx + 1}\`,
      area: atv.area || 'Gestão Financeira',
      statusImplementacao: (atv.status === 'Concluído' ? 'Implementada Integralmente' : 'Implementada Parcialmente') as any,
      aderenciaRotina: 'Alta' as const,
      resultadoObservado: atv.descricao ? \`Etapa executada: \${atv.descricao}\` : 'Atividade desenvolvida com sucesso e integrada à rotina.',
      dificuldadesEncontradas: 'Nenhuma que comprometesse os resultados esperados.'
    }));
  }

  return [
    {
      id: \`sol_def_0\`,
      solucao: 'Estruturação da Planilha de Fluxo de Caixa e Treinamento do Lançamento Diário',
      area: 'Gestão Financeira',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Ação 100% orientada e aplicada pelo cliente durante a consultoria.',
      dificuldadesEncontradas: 'Ajuste inicial de hábitos e rotina de anotação de dados.'
    },
    {
      id: \`sol_def_1\`,
      solucao: 'Definição do Pró-labore e Segregação Rigorosa das Contas PF/PJ',
      area: 'Controle e Rotina',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Separação física e bancária implementada com sucesso.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: \`sol_def_2\`,
      solucao: 'Mapeamento e Classificação Detalhada dos Custos Fixos e Variáveis',
      area: 'Custos e Resultado',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Classificação completa de custos realizada e assimilada.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: \`sol_def_3\`,
      solucao: 'Formação Técnica da Tabela de Preço de Venda e Margem de Contribuição',
      area: 'Preço e Mercado',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Margens calculadas e preços ajustados.',
      dificuldadesEncontradas: 'Nenhuma.'
    },
    {
      id: \`sol_def_4\`,
      solucao: 'Implantação do Fechamento Mensal de Resultados (DRE Gerencial)',
      area: 'Gestão Financeira',
      statusImplementacao: 'Implementada Integralmente',
      aderenciaRotina: 'Alta',
      resultadoObservado: 'Demonstrativo estruturado para acompanhamento da lucratividade.',
      dificuldadesEncontradas: 'Nenhuma.'
    }
  ];
};
`;

if (!content.includes('extractGroupedProblemasForDiag')) {
  content = content.replace(
    "export const AnaliseResultadosView: React.FC<Props> = ({",
    helperCode + "\nexport const AnaliseResultadosView: React.FC<Props> = ({"
  );
}

// 3. Update useEffect for selectedDiagnostico and autoPopulateFromDiagnostico
const oldEffectSection = `  // Re-sync when selected diagnostico changes
  useEffect(() => {
    if (selectedDiagnostico?.analiseResultados) {
      setData(prev => ({
        ...selectedDiagnostico.analiseResultados,
        codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || selectedDiagnostico.analiseResultados.codigoSgf || prev.codigoSgf,
        periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || selectedDiagnostico.analiseResultados.periodoConsultoria || prev.periodoConsultoria,
        cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || selectedDiagnostico.analiseResultados.cargaHorariaRealizada || selectedDiagnostico?.cargaHoraria || prev.cargaHorariaRealizada,
        consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || selectedDiagnostico.analiseResultados.consultorResponsavel || prev.consultorResponsavel,
        tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || selectedDiagnostico.analiseResultados.tecnicoSebrae || prev.tecnicoSebrae,
        clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || selectedDiagnostico.analiseResultados.clienteResponsavel || prev.clienteResponsavel,
        clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || selectedDiagnostico.analiseResultados.clienteCpfCnpj || prev.clienteCpfCnpj,
        empresaCredenciada: selectedDiagnostico.analiseResultados.empresaCredenciada || credenciadaName || prev.empresaCredenciada
      }));
    } else {
      setData(prev => ({
        ...prev,
        diagnosticoId: selectedDiagnostico?.id || prev.diagnosticoId,
        empresaId: selectedEmpresa?.id || selectedDiagnostico?.empresaId || prev.empresaId,
        clienteResponsavel: selectedDiagnostico?.dadosConsultoria?.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || prev.clienteResponsavel,
        clienteCpfCnpj: selectedDiagnostico?.dadosConsultoria?.cnpj || selectedEmpresa?.cnpj || prev.clienteCpfCnpj,
        consultorResponsavel: selectedDiagnostico?.dadosConsultoria?.consultor || prev.consultorResponsavel,
        tecnicoSebrae: selectedDiagnostico?.dadosConsultoria?.tecnicoSebrae || prev.tecnicoSebrae,
        codigoSgf: selectedDiagnostico?.dadosConsultoria?.codigoSgf || prev.codigoSgf,
        periodoConsultoria: selectedDiagnostico?.dadosConsultoria?.periodoConsultoria || prev.periodoConsultoria,
        cargaHorariaRealizada: selectedDiagnostico?.dadosConsultoria?.cargaHoraria || selectedDiagnostico?.cargaHoraria || prev.cargaHorariaRealizada,
        areaConsultoria: selectedDiagnostico?.dadosConsultoria?.areaConsultoria || prev.areaConsultoria,
        objetivoConsultoria: selectedDiagnostico?.dadosConsultoria?.objetivo || prev.objetivoConsultoria,
        resultadosEsperados: selectedDiagnostico?.dadosConsultoria?.resultadosEsperados || prev.resultadosEsperados,
        empresaCredenciada: credenciadaName || prev.empresaCredenciada,
        teoriaAdmin: {
          ...prev.teoriaAdmin,
          maturidadeInicialPercent: selectedDiagnostico?.percentualGeral ?? prev.teoriaAdmin.maturidadeInicialPercent
        }
      }));
    }
  }, [selectedDiagnostico?.id, selectedDiagnostico?.dadosConsultoria, credenciadaName]);`;

const newEffectSection = `  // Re-sync and isolate data strictly when selected diagnostico changes
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
      const localKey = \`analise_resultados_\${selectedDiagnostico.id}\`;
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
        principaisGanhosObtidos: '1. Clareza total da margem real de lucro.\\n2. Eliminação de despesas ocultas e juros bancários desnecessários.\\n3. Domínio da ferramenta gerencial de caixa e custos.\\n4. Autonomia na precificação técnica.',
        riscosResiduais: 'Risco de afrouxamento na disciplina de alimentação diária do fluxo de caixa em períodos de alta intensidade de trabalho no campo.',
        planoContinuidade30Dias: 'Realizar o primeiro fechamento mensal de resultado (DRE) de forma 100% autônoma e conferir a conciliação bancária semanalmente.',
        planoContinuidade60Dias: 'Apurar a margem de contribuição acumulada da safra/período e destinar 20% do lucro líquido apurado para o fundo de reserva.',
        planoContinuidade90Dias: 'Revisar a tabela de preços e custos unitários antes do início do novo ciclo produtivo.',
        recomendacoesFinais: 'Manter a reunião financeira quinzenal de fechamento e não realizar retiradas além do pró-labore estipulado.'
      }
    });
  }, [selectedDiagnostico?.id, selectedDiagnostico?.dadosConsultoria, credenciadaName]);`;

if (content.includes(oldEffectSection)) {
  content = content.replace(oldEffectSection, newEffectSection);
  console.log("Updated useEffect in AnaliseResultadosView successfully");
} else {
  console.error("Could not find oldEffectSection in AnaliseResultadosView");
}

// 4. Update autoPopulateFromDiagnostico implementation
const oldAutoPopulate = `  const autoPopulateFromDiagnostico = () => {
    setIsAutoFilling(true);
    
    // 1. Identify problems from diagnostic answers
    const diagRespostas = respostas.filter(r => r.diagnosticoId === selectedDiagnostico?.id);
    const negativeRespostas = diagRespostas.filter(r => r.resposta === 'Não' || r.resposta === 'Parcial');
    const generatedProblemas: ProblemaAvaliado[] = [];

    if (negativeRespostas.length > 0) {
      negativeRespostas.forEach((r, idx) => {
        const probText = r.problema || r.pergunta || \`Fragilidade em \${r.area || 'Gestão'}\`;
        generatedProblemas.push({
          id: \`prob_\${idx}_\${Date.now()}\`,
          problema: probText,
          area: r.area || 'Gestão Financeira',
          status: 'Resolvido',
          percentualEficacia: r.resposta === 'Não' ? 90 : 95,
          impactoGerado: \`Processo e controle estruturados durante a consultoria para sanar a falha diagnosticada em "\${probText}".\`,
          evidencias: 'Planilha / rotina implementada, acompanhada e validada com o consultor.'
        });
      });
    } else {
      const defaults = [
        { problema: 'Ausência de controle diário e rigoroso de entradas e saídas de caixa', area: 'Gestão Financeira' },
        { problema: 'Mistura de despesas pessoais com as despesas da empresa (falta de pró-labore)', area: 'Controle e Rotina' },
        { problema: 'Desconhecimento dos custos reais e da margem de contribuição', area: 'Custos e Resultado' },
        { problema: 'Precificação baseada apenas no mercado sem apuração dos custos unitários', area: 'Preço e Mercado' },
        { problema: 'Inexistência de reserva financeira para emergências e entressafras', area: 'Gestão Financeira' }
      ];
      defaults.forEach((item, idx) => {
        generatedProblemas.push({
          id: \`prob_def_\${idx}\`,
          problema: item.problema,
          area: item.area,
          status: 'Resolvido',
          percentualEficacia: 90,
          impactoGerado: 'Rotina implantada e em operação no dia a dia da empresa.',
          evidencias: 'Registros contábeis e gerenciais conferidos nas sessões técnicas.'
        });
      });
    }

    // 2. Identify solutions from cronograma activities or default 8 financial steps
    const generatedSolucoes: SolucaoAvaliada[] = [];
    const cronograma = selectedDiagnostico?.cronograma || [];

    if (cronograma.length > 0) {
      cronograma.forEach((atv: any, idx: number) => {
        generatedSolucoes.push({
          id: \`sol_\${idx}_\${Date.now()}\`,
          solucao: atv.solucaoProposta || atv.nome || \`Atividade de Consultoria \${idx + 1}\`,
          area: atv.area || 'Gestão Financeira',
          statusImplementacao: atv.status === 'Concluído' ? 'Implementada Integralmente' : 'Implementada Parcialmente',
          aderenciaRotina: 'Alta',
          resultadoObservado: atv.descricao ? \`Etapa executada: \${atv.descricao}\` : 'Atividade desenvolvida com sucesso e integrada à rotina.',
          dificuldadesEncontradas: 'Nenhuma que comprometesse os resultados esperados.'
        });
      });
    } else {
      const defaultCronogramaSolucoes = [
        { solucao: 'Estruturação da Planilha de Fluxo de Caixa e Treinamento do Lançamento Diário', area: 'Gestão Financeira' },
        { solucao: 'Definição do Pró-labore e Segregação Rigorosa das Contas PF/PJ', area: 'Controle e Rotina' },
        { solucao: 'Mapeamento e Classificação Detalhada dos Custos Fixos e Variáveis', area: 'Custos e Resultado' },
        { solucao: 'Formação Técnica da Tabela de Preço de Venda e Margem de Contribuição', area: 'Preço e Mercado' },
        { solucao: 'Implantação do Fechamento Mensal de Resultados (DRE Gerencial)', area: 'Gestão Financeira' }
      ];
      defaultCronogramaSolucoes.forEach((s, idx) => {
        generatedSolucoes.push({
          id: \`sol_def_\${idx}\`,
          solucao: s.solucao,
          area: s.area,
          statusImplementacao: 'Implementada Integralmente',
          aderenciaRotina: 'Alta',
          resultadoObservado: 'Ação 100% orientada e aplicada pelo cliente durante a consultoria.',
          dificuldadesEncontradas: 'Ajuste inicial de hábitos e rotina de anotação de dados.'
        });
      });
    }

    setData(prev => ({
      ...prev,
      problemasAvaliados: generatedProblemas,
      solucoesAvaliadas: generatedSolucoes
    }));

    setTimeout(() => setIsAutoFilling(false), 300);
  };`;

const newAutoPopulate = `  const autoPopulateFromDiagnostico = () => {
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
      id: \`prob_custom_\${Date.now()}\`,
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
      id: \`sol_custom_\${Date.now()}\`,
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
  };`;

if (content.includes(oldAutoPopulate)) {
  content = content.replace(oldAutoPopulate, newAutoPopulate);
  console.log("Updated autoPopulate in AnaliseResultadosView successfully");
} else {
  console.error("Could not find oldAutoPopulate in AnaliseResultadosView");
}

fs.writeFileSync('src/components/AnaliseResultadosView.tsx', content, 'utf8');
console.log("Updated AnaliseResultadosView step 1");
