/**
 * PDF: "Resumo dos Diagnósticos por Tipo de Negócio".
 * Estrutura: capa com indicadores gerais → visão geral dos segmentos → uma seção por segmento
 * (indicadores, áreas, problemas recorrentes e clientes).
 */
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { RelatorioSegmentos, ResumoSegmento } from '../domain/relatorioSegmentos';
import { faixaConformidade, META_POS_CONSULTORIA } from '../domain/relatorioSegmentos';

const AZUL: [number, number, number] = [0, 90, 160];
const CINZA: [number, number, number] = [71, 85, 105];
const VERMELHO: [number, number, number] = [190, 18, 60];
const AMBAR: [number, number, number] = [180, 83, 9];
const VERDE: [number, number, number] = [5, 150, 105];

const corConformidade = (v: number): [number, number, number] => (v < 40 ? VERMELHO : v < 60 ? AMBAR : v < 80 ? [3, 105, 161] : VERDE);
const dataBR = (d: Date | null) => (d ? d.toLocaleDateString('pt-BR') : '—');

const normArea = (s: string) =>
  s
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
const curto = (s: string, n: number) => (s.length > n ? s.slice(0, n - 1).trimEnd() + '…' : s);

/** Ações de continuidade por área, para sair do patamar pós-consultoria e chegar a 100%. */
const ACAO_AREA: Record<string, string> = {
  capacitacao: 'Manter um calendário anual de capacitações (equipe e gestão), com avaliação da aplicação prática dos conteúdos.',
  organizacional: 'Formalizar processos, papéis e responsabilidades em procedimentos escritos e revisá-los a cada semestre.',
  marketing: 'Executar um plano de marketing com metas mensais, presença digital ativa e pesquisa periódica de satisfação dos clientes.',
  comercializacao: 'Ampliar e diversificar os canais de venda, com metas comerciais, gestão da carteira e contratos formalizados.',
  financeiro: 'Manter fluxo de caixa, DRE e custos atualizados mensalmente, separando finanças pessoais e da empresa e formando reserva.',
  producao: 'Padronizar o processo produtivo com boas práticas, controle de qualidade e registro de produtividade e perdas.',
  embalagens: 'Adequar embalagens e rotulagem à legislação, testar melhorias de apresentação e controlar custos e estoque de embalagens.',
  sustentabilidade: 'Implantar práticas contínuas de gestão de resíduos, uso de água e energia, e buscar certificações ou selos aplicáveis.',
  gestao: 'Realizar reuniões mensais de gestão com acompanhamento de indicadores e do plano de ação.',
};

const ACOES_GERAIS = [
  'Nomear um responsável interno para acompanhar o plano de ação e os indicadores definidos na consultoria.',
  'Fazer reuniões mensais de acompanhamento, registrando resultados, desvios e correções.',
  'Manter em uso as ferramentas de gestão desenvolvidas durante a consultoria, atualizando os dados com regularidade.',
  'Refazer o diagnóstico empresarial em 6 meses para medir a evolução e priorizar as lacunas que restarem.',
  'Buscar capacitações e consultorias complementares nas áreas que ficarem abaixo de 100%.',
];

/** Ações sugeridas após a consultoria para uma área do segmento, a partir das atividades realizadas. */
function acoesPosConsultoria(s: ResumoSegmento, area: string): string[] {
  const out: string[] = [];
  const vistos = new Set<string>();
  const add = (txt?: string) => {
    const v = (txt || '').replace(/\s+/g, ' ').trim();
    if (v && !vistos.has(v.toLowerCase())) {
      vistos.add(v.toLowerCase());
      out.push(v);
    }
  };
  const probs = s.clientes.flatMap((c) => c.problemas).filter((p) => p.area === area);
  const solucoes = Array.from(new Set(probs.map((p) => p.solucao).filter(Boolean) as string[]));
  solucoes.slice(0, 2).forEach((sol) => add(`Consolidar como rotina: ${curto(sol, 110)}`));
  const kpis = probs.map((p) => p.kpis).find(Boolean);
  if (kpis) add(`Acompanhar os indicadores: ${curto(kpis, 120)}`);
  const semSolucao = Array.from(new Set(probs.filter((p) => !p.solucao).map((p) => p.problema)));
  if (semSolucao.length) add(`Elaborar plano específico para: ${curto(semSolucao.slice(0, 2).join('; '), 120)}`);
  add(ACAO_AREA[normArea(area)] || `Revisar as práticas de ${area.toLowerCase()} e corrigir as lacunas restantes do diagnóstico.`);
  return out.slice(0, 4);
}

export function gerarPdfRelatorioSegmentos(
  rel: RelatorioSegmentos,
  opcoes: { consultor?: string; logo?: string | null } = {}
): jsPDF {
  const doc = new jsPDF({ unit: 'mm', format: 'a4' });
  const largura = doc.internal.pageSize.getWidth();
  const M = 15;
  let y = 20;

  const titulo = (texto: string, tamanho = 14, cor = AZUL) => {
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(tamanho);
    doc.setTextColor(...cor);
    doc.text(texto, M, y);
    y += tamanho * 0.5 + 2;
  };
  const texto = (t: string, tamanho = 9, cor = CINZA, negrito = false) => {
    doc.setFont('helvetica', negrito ? 'bold' : 'normal');
    doc.setFontSize(tamanho);
    doc.setTextColor(...cor);
    const linhas = doc.splitTextToSize(t, largura - 2 * M);
    doc.text(linhas, M, y);
    y += linhas.length * tamanho * 0.42 + 2;
  };
  const quebra = (espaco: number) => {
    if (y + espaco > 280) {
      doc.addPage();
      y = 20;
    }
  };
  const fimTabela = () => ((doc as any).lastAutoTable?.finalY || y) + 8;

  const CINZA_CLARO: [number, number, number] = [203, 213, 225];
  const ROT_W = 48;
  const gradePct = (x0: number, w: number, h: number) => {
    doc.setDrawColor(226, 232, 240);
    doc.setLineWidth(0.2);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(6.5);
    doc.setTextColor(148, 163, 184);
    [0, 25, 50, 75, 100].forEach((p) => {
      const x = x0 + (w * p) / 100;
      doc.line(x, y, x, y + h);
      doc.text(`${p}%`, x, y + h + 3.5, { align: 'center' });
    });
  };
  /** Gráfico de barras horizontais (0–100%), colorido pela faixa de conformidade. */
  const graficoBarras = (itens: { rotulo: string; valor: number }[]) => {
    const alt = 5.5;
    const gap = 2;
    const x0 = M + ROT_W;
    const w = largura - 2 * M - ROT_W - 12;
    const h = itens.length * (alt + gap);
    quebra(h + 10);
    gradePct(x0, w, h);
    itens.forEach((it, i) => {
      const yy = y + i * (alt + gap) + gap / 2;
      const cor = corConformidade(it.valor);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(...CINZA);
      doc.text(doc.splitTextToSize(it.rotulo, ROT_W - 3)[0], M, yy + alt * 0.72);
      doc.setFillColor(...cor);
      doc.rect(x0, yy, Math.max(0.6, (w * it.valor) / 100), alt, 'F');
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...cor);
      doc.text(`${it.valor}%`, x0 + (w * it.valor) / 100 + 1.5, yy + alt * 0.72);
    });
    y += h + 8;
  };
  /** Gráfico comparativo: conformidade atual x projetada após a consultoria, com a meta de 90%. */
  const graficoComparativo = (itens: { rotulo: string; atual: number; projetada: number }[]) => {
    const bar = 3.6;
    const gap = 3;
    const rowH = bar * 2 + gap;
    const x0 = M + ROT_W;
    const w = largura - 2 * M - ROT_W - 12;
    const h = itens.length * rowH;
    quebra(h + 16);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(...CINZA);
    doc.setFillColor(...CINZA_CLARO);
    doc.rect(M, y - 2.6, 3, 3, 'F');
    doc.text('Atual (diagnóstico)', M + 4.5, y);
    doc.setFillColor(...VERDE);
    doc.rect(M + 38, y - 2.6, 3, 3, 'F');
    doc.text('Após a consultoria (projeção)', M + 42.5, y);
    doc.setDrawColor(...VERMELHO);
    doc.setLineWidth(0.4);
    doc.setLineDashPattern([1, 1], 0);
    doc.line(M + 92, y - 1, M + 98, y - 1);
    doc.setLineDashPattern([], 0);
    doc.text(`Meta pós-consultoria (${META_POS_CONSULTORIA}%)`, M + 99.5, y);
    y += 4;
    gradePct(x0, w, h);
    itens.forEach((it, i) => {
      const yy = y + i * rowH + gap / 2;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(8);
      doc.setTextColor(...CINZA);
      doc.text(doc.splitTextToSize(it.rotulo, ROT_W - 3)[0], M, yy + bar + 1);
      doc.setFillColor(...CINZA_CLARO);
      doc.rect(x0, yy, Math.max(0.6, (w * it.atual) / 100), bar, 'F');
      doc.setFillColor(...VERDE);
      doc.rect(x0, yy + bar, Math.max(0.6, (w * it.projetada) / 100), bar, 'F');
      doc.setFontSize(7);
      doc.setTextColor(...CINZA);
      doc.text(`${it.atual}%`, x0 + (w * it.atual) / 100 + 1.5, yy + bar * 0.8);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...VERDE);
      doc.text(`${it.projetada}%`, x0 + (w * it.projetada) / 100 + 1.5, yy + bar * 1.8);
    });
    const xm = x0 + (w * META_POS_CONSULTORIA) / 100;
    doc.setDrawColor(...VERMELHO);
    doc.setLineWidth(0.4);
    doc.setLineDashPattern([1, 1], 0);
    doc.line(xm, y, xm, y + h);
    doc.setLineDashPattern([], 0);
    doc.setLineWidth(0.2);
    y += h + 8;
  };

  // ------------------------------------------------------------------ Capa / visão geral
  if (opcoes.logo) {
    try {
      doc.addImage(opcoes.logo, 'PNG', largura - M - 30, 10, 30, 15);
    } catch {
      /* logo opcional */
    }
  }
  titulo('Resumo dos Diagnósticos por Tipo de Negócio', 17);
  if (rel.filtroSegmento) titulo(`Tipo de negócio: ${rel.filtroSegmento}`, 12, [30, 41, 59]);
  texto(
    `Gerado em ${rel.geradoEm.toLocaleString('pt-BR')}${opcoes.consultor ? ` | Consultor: ${opcoes.consultor}` : ''}. ` +
      'Considera o diagnóstico mais recente com respostas de cada cliente. A criticidade segue a mesma regra dos ' +
      'relatórios individuais e do plano de ação (lacunas x peso da pergunta x impacto do problema).',
    9
  );
  y += 2;

  // Cartões de indicadores
  const cartoes = [
    ['Clientes', String(rel.totalClientes)],
    ['Diagnosticados', String(rel.totalDiagnosticados)],
    ['Tipos de negócio', String(rel.segmentos.length)],
    ['Conformidade média', `${rel.conformidadeMedia}%`],
  ];
  const w = (largura - 2 * M - 9) / 4;
  cartoes.forEach(([rotulo, valor], i) => {
    const x = M + i * (w + 3);
    doc.setDrawColor(226, 232, 240);
    doc.setFillColor(248, 250, 252);
    doc.roundedRect(x, y, w, 18, 2, 2, 'FD');
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8);
    doc.setTextColor(...CINZA);
    doc.text(rotulo.toUpperCase(), x + 3, y + 6);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.setTextColor(...AZUL);
    doc.text(valor, x + 3, y + 14);
  });
  y += 26;

  titulo('Visão geral por tipo de negócio', 12);
  autoTable(doc, {
    startY: y,
    theme: 'grid',
    head: [['Tipo de negócio', 'Clientes', 'Diagnost.', 'Conformidade média', 'Faixa', 'Críticos', 'Altos', 'Moderados']],
    body: rel.segmentos.map((s) => [
      s.segmento,
      String(s.totalClientes),
      String(s.clientesDiagnosticados),
      s.clientesDiagnosticados ? `${s.conformidadeMedia}% (${s.conformidadeMin}–${s.conformidadeMax}%)` : '—',
      s.clientesDiagnosticados ? faixaConformidade(s.conformidadeMedia).rotulo : '—',
      String(s.criticos),
      String(s.altos),
      String(s.moderados),
    ]),
    styles: { fontSize: 8, cellPadding: 1.8 },
    headStyles: { fillColor: AZUL },
    columnStyles: {
      1: { halign: 'center' }, 2: { halign: 'center' }, 3: { halign: 'center' }, 4: { halign: 'center' },
      5: { halign: 'center', textColor: VERMELHO, fontStyle: 'bold' },
      6: { halign: 'center', textColor: AMBAR, fontStyle: 'bold' },
      7: { halign: 'center' },
    },
    didParseCell: (d) => {
      if (d.section === 'body' && d.column.index === 3) {
        const s = rel.segmentos[d.row.index];
        if (s?.clientesDiagnosticados) d.cell.styles.textColor = corConformidade(s.conformidadeMedia);
      }
    },
    margin: { left: M, right: M },
  });
  y = fimTabela();

  const segComDiag = rel.segmentos.filter((s) => s.clientesDiagnosticados);
  if (segComDiag.length) {
    quebra(25);
    titulo('Conformidade por tipo de negócio: atual x após a consultoria', 11);
    graficoComparativo(
      segComDiag.map((s) => ({ rotulo: s.segmento, atual: s.conformidadeMedia, projetada: s.conformidadeProjetada ?? s.conformidadeMedia }))
    );
  }

  if (rel.totalDiagnosticados === 0) {
    texto('Nenhum cliente possui diagnóstico respondido até o momento.', 10, CINZA, true);
  }

  // ------------------------------------------------------------------ Seções por segmento
  for (const s of rel.segmentos) {
    doc.addPage();
    y = 20;
    titulo(`Tipo de negócio: ${s.segmento}`, 15);
    texto(
      `${s.totalClientes} cliente(s), ${s.clientesDiagnosticados} com diagnóstico respondido. ` +
        (s.clientesDiagnosticados
          ? `Conformidade média de ${s.conformidadeMedia}% (faixa ${faixaConformidade(s.conformidadeMedia).rotulo.toLowerCase()}), ` +
            `variando de ${s.conformidadeMin}% a ${s.conformidadeMax}%. Foram identificados ${s.criticos} problema(s) crítico(s), ` +
            `${s.altos} alto(s) e ${s.moderados} moderado(s) no total.`
          : 'Ainda não há diagnósticos respondidos neste segmento.'),
      9.5
    );

    if (s.clientesDiagnosticados) {
      // Síntese: pior área e problema mais recorrente
      const piorArea = s.areas[0];
      const topProblema = s.recorrentes[0];
      const sintese: string[] = [];
      if (piorArea) sintese.push(`Área mais frágil: ${piorArea.area} (${piorArea.conformidadeMedia}% de conformidade média).`);
      if (topProblema)
        sintese.push(
          `Problema mais recorrente: "${topProblema.problema}", presente em ${topProblema.clientes} de ${s.clientesDiagnosticados} cliente(s)` +
            `${topProblema.comoCritico ? ` (crítico em ${topProblema.comoCritico})` : ''}.`
        );
      if (sintese.length) texto(sintese.join(' '), 9.5, [30, 41, 59], true);
      y += 2;

      // Áreas
      quebra(30);
      titulo('Conformidade por área', 11);
      autoTable(doc, {
        startY: y,
        theme: 'striped',
        head: [['Área', 'Conformidade média', 'Faixa', 'Clientes avaliados', 'Clientes com lacunas']],
        body: s.areas.map((a) => [a.area, `${a.conformidadeMedia}%`, faixaConformidade(a.conformidadeMedia).rotulo, String(a.clientesAvaliados), String(a.clientesComLacuna)]),
        styles: { fontSize: 8, cellPadding: 1.6 },
        headStyles: { fillColor: AZUL },
        columnStyles: { 1: { halign: 'center', fontStyle: 'bold' }, 2: { halign: 'center' }, 3: { halign: 'center' }, 4: { halign: 'center' } },
        didParseCell: (d) => {
          if (d.section === 'body' && d.column.index === 1) d.cell.styles.textColor = corConformidade(s.areas[d.row.index].conformidadeMedia);
        },
        margin: { left: M, right: M },
      });
      y = fimTabela();

      // Gráficos de conformidade por área: diagnóstico e projeção após a consultoria
      if (s.areas.length) {
        quebra(25);
        titulo('Gráfico: conformidade por área (diagnóstico)', 11);
        graficoBarras(s.areas.map((a) => ({ rotulo: a.area, valor: a.conformidadeMedia })));
        quebra(30);
        titulo('Conformidade projetada após a consultoria', 11);
        texto(
          `Projeção considerando que o cliente implemente as soluções indicadas nas atividades da consultoria para os problemas ` +
            `diagnosticados, com teto de ${META_POS_CONSULTORIA}% de conformidade. Conformidade média do segmento: de ` +
            `${s.conformidadeMedia}% para ${s.conformidadeProjetada ?? s.conformidadeMedia}%.`,
          8.5
        );
        graficoComparativo(
          s.areas.map((a) => ({ rotulo: a.area, atual: a.conformidadeMedia, projetada: a.conformidadeProjetada ?? a.conformidadeMedia }))
        );
      }

      // Problemas recorrentes
      if (s.recorrentes.length) {
        quebra(30);
        titulo('Problemas mais recorrentes no segmento', 11);
        autoTable(doc, {
          startY: y,
          theme: 'striped',
          head: [['#', 'Problema', 'Área', 'Clientes', '% do segmento', 'Como crítico']],
          body: s.recorrentes.map((r, i) => [String(i + 1), r.problema, r.area, String(r.clientes), `${r.percentual}%`, String(r.comoCritico || '—')]),
          styles: { fontSize: 8, cellPadding: 1.6 },
          headStyles: { fillColor: AZUL },
          columnStyles: { 0: { cellWidth: 7, halign: 'center' }, 3: { halign: 'center' }, 4: { halign: 'center' }, 5: { halign: 'center', textColor: VERMELHO, fontStyle: 'bold' } },
          margin: { left: M, right: M },
        });
        y = fimTabela();
      }

      // Clientes
      quebra(30);
      titulo('Clientes (do menor para o maior índice de conformidade)', 11);
      autoTable(doc, {
        startY: y,
        theme: 'grid',
        head: [['Cliente', 'Diagnóstico', 'Conform.', 'Críticos', 'Altos', 'Moder.', 'Principais problemas críticos']],
        body: s.clientes.map((c) => [
          c.cliente,
          dataBR(c.data) + (c.totalDiagnosticos > 1 ? `\n(${c.totalDiagnosticos} diagnósticos)` : ''),
          `${c.conformidade}%`,
          String(c.criticos),
          String(c.altos),
          String(c.moderados),
          c.principaisCriticos.length ? c.principaisCriticos.map((p) => `• ${p}`).join('\n') : '—',
        ]),
        styles: { fontSize: 7.5, cellPadding: 1.5, valign: 'top' },
        headStyles: { fillColor: AZUL },
        columnStyles: {
          0: { cellWidth: 34, fontStyle: 'bold' },
          1: { cellWidth: 22 },
          2: { cellWidth: 20, halign: 'center', fontStyle: 'bold' },
          3: { cellWidth: 13, halign: 'center', textColor: VERMELHO, fontStyle: 'bold' },
          4: { cellWidth: 11, halign: 'center', textColor: AMBAR },
          5: { cellWidth: 12, halign: 'center' },
        },
        didParseCell: (d) => {
          if (d.section === 'body' && d.column.index === 2) d.cell.styles.textColor = corConformidade(s.clientes[d.row.index].conformidade);
        },
        margin: { left: M, right: M },
      });
      y = fimTabela();
    }

    // Problemas diagnosticados: TODOS os problemas de cada cliente, com área, criticidade e evidências
    if (s.clientes.some((c) => c.problemas.length)) {
      quebra(30);
      titulo('Problemas diagnosticados por cliente', 11);
      for (const c of s.clientes) {
        quebra(22);
        texto(
          `${c.cliente} — ${c.problemas.length} problema(s) diagnosticado(s)` +
            (c.problemas.length ? ` (${c.criticos} crítico(s), ${c.altos} alto(s), ${c.moderados} moderado(s))` : ''),
          9.5,
          [30, 41, 59],
          true
        );
        if (!c.problemas.length) {
          texto('Nenhum problema identificado: todas as respostas indicam boas práticas.', 8.5, VERDE);
          y += 2;
          continue;
        }
        autoTable(doc, {
          startY: y,
          theme: 'grid',
          head: [['Criticidade', 'Problema diagnosticado', 'Área', 'Evidências no diagnóstico']],
          body: c.problemas.map((p) => [p.nivel, p.problema, p.area, p.evidencias.length ? p.evidencias.map((e) => `• ${e}`).join('\n') : '—']),
          styles: { fontSize: 7.5, cellPadding: 1.5, valign: 'top' },
          headStyles: { fillColor: [71, 85, 105] },
          columnStyles: {
            0: { cellWidth: 20, halign: 'center', fontStyle: 'bold' },
            1: { cellWidth: 55, fontStyle: 'bold' },
            2: { cellWidth: 28 },
          },
          didParseCell: (d) => {
            if (d.section === 'body' && d.column.index === 0) {
              const n = c.problemas[d.row.index]?.nivel;
              d.cell.styles.textColor = n === 'Crítico' ? VERMELHO : n === 'Alto' ? AMBAR : [3, 105, 161];
            }
          },
          margin: { left: M, right: M },
        });
        y = fimTabela() - 2;
      }
      y += 4;
    }

    if (s.clientesSemDiagnostico.length) {
      quebra(20);
      texto(`Clientes sem diagnóstico respondido: ${s.clientesSemDiagnostico.join(', ')}.`, 8.5, CINZA);
    }
  }

  // ------------------------------------------------------------------ Pós-consultoria (rumo a 100%)
  if (segComDiag.length) {
    doc.addPage();
    y = 20;
    titulo('Sugestões de ações após a consultoria', 15);
    texto(
      `Com a implementação das soluções indicadas nas atividades da consultoria, a projeção é de até ${META_POS_CONSULTORIA}% de ` +
        'conformidade. Para chegar a 100%, recomenda-se dar continuidade às ações abaixo, definidas a partir das atividades ' +
        'realizadas na consultoria e das lacunas que permanecem em cada área.',
      9.5
    );
    y += 2;
    titulo('Ações gerais de continuidade', 11);
    ACOES_GERAIS.forEach((a) => {
      quebra(10);
      texto(`• ${a}`, 9);
    });
    y += 4;
    for (const s of segComDiag) {
      quebra(40);
      titulo(`Tipo de negócio: ${s.segmento}`, 12, [30, 41, 59]);
      const linhas = s.areas
        .filter((a) => (a.conformidadeProjetada ?? a.conformidadeMedia) < 100)
        .map((a) => {
          const proj = a.conformidadeProjetada ?? a.conformidadeMedia;
          return [
            a.area,
            `${a.conformidadeMedia}%`,
            `${proj}%`,
            `${100 - proj} p.p.`,
            acoesPosConsultoria(s, a.area).map((x) => `• ${x}`).join('\n'),
          ];
        });
      if (!linhas.length) {
        texto('Todas as áreas já estão com 100% de conformidade.', 9, VERDE);
        y += 2;
        continue;
      }
      autoTable(doc, {
        startY: y,
        theme: 'grid',
        head: [['Área', 'Atual', 'Após consultoria', 'Falta p/ 100%', 'Ações sugeridas após a consultoria']],
        body: linhas,
        styles: { fontSize: 7.5, cellPadding: 1.5, valign: 'top' },
        headStyles: { fillColor: AZUL },
        columnStyles: {
          0: { cellWidth: 28, fontStyle: 'bold' },
          1: { cellWidth: 14, halign: 'center' },
          2: { cellWidth: 20, halign: 'center', fontStyle: 'bold', textColor: VERDE },
          3: { cellWidth: 17, halign: 'center' },
        },
        margin: { left: M, right: M },
      });
      y = fimTabela();
    }
  }

  // Rodapé com paginação
  const total = doc.getNumberOfPages();
  for (let i = 1; i <= total; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7.5);
    doc.setTextColor(148, 163, 184);
    doc.text(`Resumo dos Diagnósticos por Tipo de Negócio${rel.filtroSegmento ? ` - ${rel.filtroSegmento}` : ''}`, M, 290);
    doc.text(`Página ${i} de ${total}`, largura - M, 290, { align: 'right' });
  }
  return doc;
}
