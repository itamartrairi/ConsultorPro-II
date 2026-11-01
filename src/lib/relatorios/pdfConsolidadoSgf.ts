/**
 * Gerador de PDF: Relatório Consolidado de Diagnósticos.
 * Concatena os diagnósticos por:
 *  1. Tipo de Negócio
 *  2. Código de Contratação no SGF
 *  3. Área Diagnosticada
 * Inclui:
 *  - Tabela detalhada: Cliente, Pergunta, Resposta, Ponto Forte / Ponto Crítico, Observações
 *  - Gráficos comparativos (vetoriais nativos em alta resolução)
 *  - Cartões de indicadores e filtros aplicados
 */

import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import type { RelatorioConsolidadoSgfData } from '../domain/relatorioConsolidadoSgf';

// Cores corporativas elegantes
const AZUL_NAVY: [number, number, number] = [15, 23, 42]; // slate-900
const AZUL_SEBRAE: [number, number, number] = [0, 90, 160];
const VERDE_SUCESSO: [number, number, number] = [5, 150, 105]; // emerald-600
const VERDE_BG: [number, number, number] = [236, 253, 245]; // emerald-50
const VERMELHO_CRITICO: [number, number, number] = [225, 29, 72]; // rose-600
const VERMELHO_BG: [number, number, number] = [255, 241, 242]; // rose-50
const AMBAR_ALERTA: [number, number, number] = [217, 119, 6]; // amber-600
const AMBAR_BG: [number, number, number] = [254, 243, 199]; // amber-50
const CINZA_TEXTO: [number, number, number] = [71, 85, 105]; // slate-600
const CINZA_BORDA: [number, number, number] = [226, 232, 240]; // slate-200
const CINZA_BG: [number, number, number] = [248, 250, 252]; // slate-50

export interface OpcoesPdfConsolidado {
  consultor?: string;
  logo?: string | null;
}

export function gerarPdfConsolidadoSgf(
  rel: RelatorioConsolidadoSgfData,
  opcoes: OpcoesPdfConsolidado = {}
): jsPDF {
  const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'landscape' });
  const largura = doc.internal.pageSize.getWidth();
  const altura = doc.internal.pageSize.getHeight();
  const M = 14; // margem lateral
  let y = 16;

  const quebra = (espacoNecessario: number) => {
    if (y + espacoNecessario > altura - 18) {
      doc.addPage();
      y = 16;
    }
  };

  const fimTabela = () => ((doc as any).lastAutoTable?.finalY || y) + 6;

  // -------------------------------------------------------------
  // CAPA / CABEÇALHO DO RELATÓRIO
  // -------------------------------------------------------------
  if (opcoes.logo) {
    try {
      doc.addImage(opcoes.logo, 'PNG', largura - M - 35, 10, 35, 16);
    } catch {
      /* logo opcional */
    }
  }

  // Faixa de identificação superior
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(...AZUL_SEBRAE);
  doc.text('CONSULTOR PRO • RELATÓRIO TÉCNICO CONSOLIDADO', M, y);
  y += 5;

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.setTextColor(...AZUL_NAVY);
  doc.text('Diagnósticos por Tipo de Negócio, Código SGF e Área', M, y);
  y += 5;

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(8.5);
  doc.setTextColor(...CINZA_TEXTO);
  const infoEmissao = `Gerado em ${rel.geradoEm.toLocaleDateString('pt-BR')} às ${rel.geradoEm.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}${opcoes.consultor ? ` | Consultor: ${opcoes.consultor}` : ''}`;
  doc.text(infoEmissao, M, y);
  y += 5;

  // Resumo dos filtros aplicados
  const filtrosTexto: string[] = [];
  if (rel.filtrosAplicados.tipoNegocio && rel.filtrosAplicados.tipoNegocio !== 'Todos') {
    filtrosTexto.push(`Tipo de Negócio: ${rel.filtrosAplicados.tipoNegocio}`);
  }
  if (rel.filtrosAplicados.codigoSgf && rel.filtrosAplicados.codigoSgf !== 'Todos') {
    filtrosTexto.push(`Código SGF: ${rel.filtrosAplicados.codigoSgf}`);
  }
  if (rel.filtrosAplicados.area && rel.filtrosAplicados.area !== 'Todas') {
    filtrosTexto.push(`Área: ${rel.filtrosAplicados.area}`);
  }
  if (rel.filtrosAplicados.classificacao && rel.filtrosAplicados.classificacao !== 'todos') {
    filtrosTexto.push(`Filtro: ${rel.filtrosAplicados.classificacao === 'pontos_fortes' ? 'Somente Pontos Fortes' : 'Somente Pontos Críticos'}`);
  }
  if (rel.filtrosAplicados.buscaTexto) {
    filtrosTexto.push(`Busca: "${rel.filtrosAplicados.buscaTexto}"`);
  }

  const textoFiltros = filtrosTexto.length > 0
    ? `Filtros Ativos: ${filtrosTexto.join(' | ')}`
    : 'Filtros Ativos: Todos os clientes, tipos de negócio e contratos SGF';

  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);
  doc.setTextColor(...AZUL_SEBRAE);
  doc.text(textoFiltros, M, y);
  y += 6;

  // -------------------------------------------------------------
  // CARTÕES DE INDICADORES (KPIs)
  // -------------------------------------------------------------
  const kpis = [
    { rotulo: 'CLIENTES ATENDIDOS', valor: String(rel.totalClientes), cor: AZUL_NAVY },
    { rotulo: 'CONTRATOS SGF', valor: String(rel.comparativoPorSgf.length), cor: AZUL_SEBRAE },
    { rotulo: 'TOTAL DE AVALIAÇÕES', valor: String(rel.totalItens), cor: AZUL_NAVY },
    { rotulo: 'PONTOS FORTES', valor: `${rel.totalPontosFortes} (${rel.percentualGeralPontosFortes}%)`, cor: VERDE_SUCESSO },
    { rotulo: 'PONTOS CRÍTICOS / LACUNAS', valor: `${rel.totalPontosCriticos + rel.totalOportunidades} (${rel.percentualGeralPontosCriticos}%)`, cor: VERMELHO_CRITICO },
  ];

  const larguraCartao = (largura - 2 * M - (kpis.length - 1) * 3) / kpis.length;
  const alturaCartao = 16;

  kpis.forEach((kpi, i) => {
    const x = M + i * (larguraCartao + 3);
    doc.setDrawColor(...CINZA_BORDA);
    doc.setFillColor(...CINZA_BG);
    doc.roundedRect(x, y, larguraCartao, alturaCartao, 1.5, 1.5, 'FD');

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(6.5);
    doc.setTextColor(...CINZA_TEXTO);
    doc.text(kpi.rotulo, x + 3, y + 5);

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(...kpi.cor);
    doc.text(kpi.valor, x + 3, y + 12);
  });
  y += alturaCartao + 8;

  // -------------------------------------------------------------
  // SEÇÃO DE GRÁFICOS COMPARATIVOS VETORIAIS
  // -------------------------------------------------------------
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(11);
  doc.setTextColor(...AZUL_NAVY);
  doc.text('Gráficos Comparativos: Distribuição de Pontos Fortes vs Pontos Críticos por Área', M, y);
  y += 5;

  // Tabela / Gráfico de barras comparativas por área
  if (rel.comparativoPorArea.length > 0) {
    const barraMaxLargura = 90;
    const areasParaGrafico = rel.comparativoPorArea.slice(0, 7); // principais áreas

    areasParaGrafico.forEach((area) => {
      quebra(10);
      const percFortes = area.percentualPontosFortes;
      const percCriticos = area.percentualPontosCriticos;

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8);
      doc.setTextColor(...AZUL_NAVY);
      doc.text(area.area, M, y + 3.5);

      const xBarra = M + 55;
      const wFortes = (percFortes / 100) * barraMaxLargura;
      const wCriticos = (percCriticos / 100) * barraMaxLargura;

      // Barra de fundo
      doc.setFillColor(241, 245, 249);
      doc.rect(xBarra, y, barraMaxLargura, 5, 'F');

      // Barra verde (Pontos Fortes)
      if (wFortes > 0) {
        doc.setFillColor(...VERDE_SUCESSO);
        doc.rect(xBarra, y, wFortes, 5, 'F');
      }

      // Barra vermelha (Pontos Críticos)
      if (wCriticos > 0) {
        doc.setFillColor(...VERMELHO_CRITICO);
        doc.rect(xBarra + wFortes, y, wCriticos, 5, 'F');
      }

      // Rótulos numéricos ao lado da barra
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(...VERDE_SUCESSO);
      doc.text(`${percFortes}% Fortes (${area.pontosFortes})`, xBarra + barraMaxLargura + 4, y + 3.5);

      doc.setTextColor(...VERMELHO_CRITICO);
      doc.text(`${percCriticos}% Críticos (${area.pontosCriticos})`, xBarra + barraMaxLargura + 44, y + 3.5);

      doc.setFont('helvetica', 'normal');
      doc.setTextColor(...CINZA_TEXTO);
      doc.text(`Total: ${area.totalPerguntas}`, xBarra + barraMaxLargura + 82, y + 3.5);

      y += 7.5;
    });

    y += 4;
  }

  // -------------------------------------------------------------
  // TABELAS CONCATENADAS POR TIPO DE NEGÓCIO, CÓDIGO SGF E ÁREA
  // -------------------------------------------------------------
  if (rel.grupos.length === 0) {
    quebra(20);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(...CINZA_TEXTO);
    doc.text('Nenhuma avaliação encontrada para os filtros selecionados.', M, y);
  }

  for (const grupoTipo of rel.grupos) {
    quebra(24);

    // Header 1: TIPO DE NEGÓCIO
    doc.setFillColor(...AZUL_SEBRAE);
    doc.roundedRect(M, y, largura - 2 * M, 8, 1, 1, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(10);
    doc.setTextColor(255, 255, 255);
    doc.text(
      `TIPO DE NEGÓCIO: ${grupoTipo.tipoNegocio.toUpperCase()} (${grupoTipo.totalItens} perguntas | ${grupoTipo.pontosFortes} Pontos Fortes | ${grupoTipo.pontosCriticos} Pontos Críticos)`,
      M + 4,
      y + 5.5
    );
    y += 11;

    for (const grupoSgf of grupoTipo.sgfs) {
      quebra(18);

      // Header 2: CÓDIGO DE CONTRATAÇÃO SGF
      doc.setFillColor(241, 245, 249); // slate-100
      doc.setDrawColor(...CINZA_BORDA);
      doc.roundedRect(M, y, largura - 2 * M, 6.5, 1, 1, 'FD');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(8.5);
      doc.setTextColor(...AZUL_NAVY);
      doc.text(
        `CÓDIGO SGF: ${grupoSgf.codigoSgf} — ${grupoSgf.totalItens} avaliações (${grupoSgf.pontosFortes} Fortes / ${grupoSgf.pontosCriticos} Críticos)`,
        M + 3,
        y + 4.5
      );
      y += 9;

      for (const grupoArea of grupoSgf.areas) {
        quebra(18);

        // Header 3: ÁREA DIAGNOSTICADA
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(9);
        doc.setTextColor(...AZUL_SEBRAE);
        doc.text(
          `Área: ${grupoArea.area} (${grupoArea.totalItens} itens: ${grupoArea.pontosFortes} Fortes, ${grupoArea.pontosCriticos} Críticos)`,
          M,
          y
        );
        y += 3;

        // Tabela autoTable detalhada
        autoTable(doc, {
          startY: y,
          theme: 'grid',
          head: [['Cliente', 'Pergunta Diagnosticada', 'Resp.', 'Avaliação', 'Observações / Evidências do Consultor']],
          body: grupoArea.itens.map((it) => [
            it.cliente,
            it.pergunta,
            it.resposta || '—',
            it.classificacao.toUpperCase(),
            it.observacao || '—',
          ]),
          styles: {
            fontSize: 7.5,
            cellPadding: 1.6,
            valign: 'middle',
            textColor: [30, 41, 59],
          },
          headStyles: {
            fillColor: [51, 65, 85], // slate-700
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            fontSize: 7.5,
          },
          columnStyles: {
            0: { cellWidth: 42, fontStyle: 'bold' },
            1: { cellWidth: 105 },
            2: { cellWidth: 16, halign: 'center', fontStyle: 'bold' },
            3: { cellWidth: 32, halign: 'center', fontStyle: 'bold' },
            4: { cellWidth: 74 },
          },
          didParseCell: (d) => {
            if (d.section === 'body') {
              const item = grupoArea.itens[d.row.index];
              if (!item) return;

              // Colorir a coluna de Avaliação
              if (d.column.index === 3) {
                if (item.isPontoForte) {
                  d.cell.styles.fillColor = VERDE_BG;
                  d.cell.styles.textColor = VERDE_SUCESSO;
                } else if (item.classificacao === 'Ponto Crítico') {
                  d.cell.styles.fillColor = VERMELHO_BG;
                  d.cell.styles.textColor = VERMELHO_CRITICO;
                } else {
                  d.cell.styles.fillColor = AMBAR_BG;
                  d.cell.styles.textColor = AMBAR_ALERTA;
                }
              }

              // Colorir coluna Resposta
              if (d.column.index === 2) {
                if (item.resposta === 'Sim') {
                  d.cell.styles.textColor = VERDE_SUCESSO;
                } else if (item.resposta === 'Não') {
                  d.cell.styles.textColor = VERMELHO_CRITICO;
                } else if (item.resposta === 'Parcial') {
                  d.cell.styles.textColor = AMBAR_ALERTA;
                }
              }
            }
          },
          margin: { left: M, right: M },
        });

        y = fimTabela();
      }

      y += 2;
    }

    y += 4;
  }

  // -------------------------------------------------------------
  // RODAPÉ COM PAGINAÇÃO EM TODAS AS PÁGINAS
  // -------------------------------------------------------------
  const totalPaginas = doc.getNumberOfPages();
  for (let i = 1; i <= totalPaginas; i++) {
    doc.setPage(i);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(148, 163, 184); // slate-400

    doc.text(
      'Consultor Pro • Relatório Consolidado por Tipo de Negócio, Código SGF e Área',
      M,
      altura - 7
    );
    doc.text(
      `Página ${i} de ${totalPaginas}`,
      largura - M,
      altura - 7,
      { align: 'right' }
    );
  }

  return doc;
}
