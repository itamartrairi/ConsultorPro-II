import { montarRelatorioConsolidadoSgf, extrairItensDiagnosticos } from '../src/lib/domain/relatorioConsolidadoSgf';
import type { Empresa, Diagnostico, Resposta } from '../src/types/domain';

const empresasMock: Empresa[] = [
  {
    id: 'emp-1',
    nome: 'Fazenda Boa Esperança',
    tipoEmpresa: 'Agronegócio',
    ownerId: 'owner-1',
    dataCadastro: new Date()
  },
  {
    id: 'emp-2',
    nome: 'Mercadinho Central',
    tipoEmpresa: 'Comércio',
    ownerId: 'owner-1',
    dataCadastro: new Date()
  }
];

const diagnosticosMock: Diagnostico[] = [
  {
    id: 'diag-1',
    empresaId: 'emp-1',
    dataDiagnostico: new Date('2026-05-10'),
    ownerId: 'owner-1',
    dadosConsultoria: {
      codigoSgf: 'SGF-2026-001',
      consultor: 'Consultor Teste'
    }
  },
  {
    id: 'diag-2',
    empresaId: 'emp-2',
    dataDiagnostico: new Date('2026-06-15'),
    ownerId: 'owner-1',
    dadosConsultoria: {
      codigoSgf: 'SGF-2026-002',
      consultor: 'Consultor Teste'
    }
  }
];

const respostasMock: Resposta[] = [
  // emp-1: Agronegócio, SGF-2026-001
  {
    id: 'resp-1',
    diagnosticoId: 'diag-1',
    premissaId: 'p-1',
    idProblema: 'prob-1',
    problema: 'Ausência de controle de caixa',
    pergunta: 'A empresa possui fluxo de caixa diário?',
    peso: 2,
    area: 'Financeiro',
    resposta: 'Sim',
    observacao: 'Fluxo em dia',
    score: 10
  },
  {
    id: 'resp-2',
    diagnosticoId: 'diag-1',
    premissaId: 'p-2',
    idProblema: 'prob-2',
    problema: 'Endividamento excessivo',
    pergunta: 'Possui dívidas em atraso com fornecedores?',
    peso: 3,
    area: 'Financeiro',
    resposta: 'Não', // Pergunta de risco com resposta Não = Ponto Forte
    observacao: '',
    score: 10
  },
  {
    id: 'resp-3',
    diagnosticoId: 'diag-1',
    premissaId: 'p-3',
    idProblema: 'prob-3',
    problema: 'Falta de licença ambiental',
    pergunta: 'Possui licença ambiental ou CAR regularizado?',
    peso: 3,
    area: 'Ambiental e Legal',
    resposta: 'Não', // Obrigação legal com Não = Ponto Crítico
    observacao: 'Precisa regularizar urgente',
    score: 0
  },
  // emp-2: Comércio, SGF-2026-002
  {
    id: 'resp-4',
    diagnosticoId: 'diag-2',
    premissaId: 'p-4',
    idProblema: 'prob-4',
    problema: 'Inadimplência de clientes',
    pergunta: 'A empresa controla a inadimplência dos clientes?',
    peso: 2,
    area: 'Comercial',
    resposta: 'Não', // Falta de controle = Ponto Crítico
    observacao: 'Não há registro',
    score: 0
  }
];

console.log('--- Testando extração e consolidação SGF ---');

const itens = extrairItensDiagnosticos({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasMock
});

if (itens.length !== 4) {
  throw new Error(`Esperado 4 itens, obteve ${itens.length}`);
}

const itemSim = itens.find(i => i.id === 'resp-1')!;
if (itemSim.classificacao !== 'Ponto Forte') {
  throw new Error(`resp-1 deveria ser Ponto Forte, mas foi ${itemSim.classificacao}`);
}

const itemLicenca = itens.find(i => i.id === 'resp-3')!;
if (itemLicenca.classificacao !== 'Ponto Crítico') {
  throw new Error(`resp-3 deveria ser Ponto Crítico, mas foi ${itemLicenca.classificacao}`);
}

console.log('  OK: Classificação semântica de Pontos Fortes e Pontos Críticos');

// Testando relatório consolidado sem filtro
const relatorioGeral = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasMock
});

if (relatorioGeral.totalClientes !== 2) {
  throw new Error(`Total clientes esperado 2, obteve ${relatorioGeral.totalClientes}`);
}
if (relatorioGeral.grupos.length !== 2) {
  throw new Error(`Esperado 2 grupos de tipo de negócio, obteve ${relatorioGeral.grupos.length}`);
}

console.log('  OK: Agrupamento em 3 níveis (Tipo de Negócio -> SGF -> Área)');

// Testando filtro por Tipo de Negócio
const relatorioFiltrado = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasMock,
  filtros: { tipoNegocio: 'Agronegócio' }
});

if (relatorioFiltrado.grupos.length !== 1 || relatorioFiltrado.grupos[0].tipoNegocio !== 'Agronegócio') {
  throw new Error(`Filtro por tipo de negócio falhou`);
}
if (relatorioFiltrado.totalItens !== 3) {
  throw new Error(`Esperado 3 itens no filtro Agronegócio, obteve ${relatorioFiltrado.totalItens}`);
}

console.log('  OK: Filtro por Tipo de Negócio');

// Testando filtro por Código SGF
const relatorioSgf = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasMock,
  filtros: { codigoSgf: 'SGF-2026-002' }
});

if (relatorioSgf.totalItens !== 1 || relatorioSgf.itensFiltrados[0].cliente !== 'Mercadinho Central') {
  throw new Error(`Filtro por Código SGF falhou`);
}

console.log('  OK: Filtro por Código SGF');

// Testando filtro por Classificação (somente Pontos Críticos)
const relatorioCriticos = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasMock,
  filtros: { classificacao: 'pontos_criticos' }
});

if (relatorioCriticos.totalItens !== 2) {
  throw new Error(`Esperado 2 pontos críticos, obteve ${relatorioCriticos.totalItens}`);
}

console.log('  OK: Filtro por Classificação (Pontos Críticos)');

// Testando modificação do diagnóstico: alteração de respostas de Sim para Não
const respostasModificadas: Resposta[] = [
  ...respostasMock,
  // resp-1 era Sim (Ponto Forte). Agora alterada para Não (Ponto Crítico)
  {
    id: 'resp-1',
    diagnosticoId: 'diag-1',
    premissaId: 'p-1',
    idProblema: 'prob-1',
    problema: 'Ausência de controle de caixa',
    pergunta: 'A empresa possui fluxo de caixa diário?',
    peso: 2,
    area: 'Financeiro',
    resposta: 'Não',
    observacao: 'Deixou de controlar o caixa',
    score: 0
  }
];

const relatorioAposModificarResposta = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosMock,
  respostas: respostasModificadas,
  filtros: {}
});

if (relatorioAposModificarResposta.totalPontosFortes !== 1 || relatorioAposModificarResposta.totalPontosCriticos !== 3) {
  throw new Error(`Modificação de resposta no diagnóstico não alterou o relatório corretamente! Fortes: ${relatorioAposModificarResposta.totalPontosFortes}, Críticos: ${relatorioAposModificarResposta.totalPontosCriticos}`);
}
console.log('  OK: Modificação de respostas no diagnóstico altera o relatório em tempo real');

// Testando modificação do diagnóstico: alteração de segmento / tipoEmpresa e código SGF
const diagnosticosModificados: Diagnostico[] = [
  {
    ...diagnosticosMock[0],
    tipoEmpresa: 'Indústria e Transformação',
    dadosConsultoria: {
      codigoSgf: 'SGF-NOVO-999',
      consultor: 'Consultor Alterado'
    }
  },
  diagnosticosMock[1]
];

const relatorioAposModificarDiag = montarRelatorioConsolidadoSgf({
  empresas: empresasMock,
  diagnosticos: diagnosticosModificados,
  respostas: respostasMock,
  filtros: {}
});

const grupoNovo = relatorioAposModificarDiag.grupos.find(g => g.tipoNegocio === 'Indústria e Transformação');
if (!grupoNovo) {
  throw new Error('Alteração de segmento no diagnóstico não refletiu no agrupamento do relatório');
}
const sgfNovo = grupoNovo.sgfs.find(s => s.codigoSgf === 'SGF-NOVO-999');
if (!sgfNovo) {
  throw new Error('Alteração de código SGF no diagnóstico não refletiu no agrupamento do relatório');
}
console.log('  OK: Modificação de segmento e SGF no diagnóstico altera o relatório em tempo real');

console.log('Todos os testes de relatorioConsolidadoSgf passaram com sucesso!');
