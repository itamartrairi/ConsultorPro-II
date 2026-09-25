/**
 * PDF: "Resumo dos Diagnósticos por Tipo de Negócio".
 * Estrutura: capa com indicadores gerais → visão geral dos segmentos → uma seção por segmento
 * (indicadores, áreas, problemas recorrentes e clientes).
 */
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { RelatorioSegmentos } from '../domain/relatorioSegmentos';
import { faixaConformidade } from '../domain/relatorioSegmentos';

const AZUL: [number, number, number] = [0, 90, 160];
const CINZA: [number, number, number] = [71, 85, 105];
const VERMELHO: [number, number, number] = [190, 18, 60];
const AMBAR: [number, number, number] = [180, 83, 9];
const VERDE: [number, number, number] = [5, 150, 105];

const corConformidade = (v: number): [number, number, number] => (v < 40 ? VERMELHO : v < 60 ? AMBAR : v < 80 ? [3, 105, 161] : VERDE);
const dataBR = (d: Date | null) => (d ? d.toLocaleDateString('pt-BR') : '—');

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
