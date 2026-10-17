/**
 * Geração do PLANO DE AÇÃO e dos textos do RELATÓRIO a partir do diagnóstico.
 *
 * Regras:
 *  1. Os problemas são classificados pela criticidade real das respostas (lacunas × peso)
 *     combinada com o impacto cadastrado do problema (Alto / Médio / Baixo).
 *  2. O plano é organizado POR ÁREA diagnosticada: cada área com lacunas vira uma atividade
 *     que trata TODOS os problemas daquela área (críticos primeiro). Nenhum problema fica de
 *     fora; cada um vira uma tarefa no Plano de Ação.
 *  3. Estrutura fixa: diagnóstico → devolutiva → ferramentas → áreas → análise final →
 *     relatórios (máximo de 10 atividades).
 *  4. A IA (opcional) só reescreve os textos de cada atividade; a estrutura é garantida
 *     aqui, então o plano sai correto mesmo sem IA ou se ela falhar.
 */
import type { AtividadeCronograma, Problema, ProblemaTratado, Resposta, Solucao } from '../../types/domain';
import { Type } from '../ai/schemaTypes';

export type NivelCriticidade = 'Crítico' | 'Alto' | 'Moderado';

export interface Lacuna {
  pergunta: string;
  resposta: string;
  observacao?: string;
}

export interface ProblemaAnalisado {
  idProblema: string;
  problema: string;
  area: string;
  impacto: string;
  nivel: NivelCriticidade;
  /** 0–100: quanto maior, mais urgente. */
  pontuacao: number;
  /** Fração das respostas que indicam o problema (0–1), ponderada pelo peso. */
  intensidade: number;
  lacunas: Lacuna[];
  solucao?: Solucao;
}

export interface AnaliseDiagnostico {
  totalRespondidas: number;
  problemas: ProblemaAnalisado[]; // só os que têm lacuna, do mais ao menos crítico
  pontosFortes: { problema: string; area: string }[];
  scoreGeral: number; // 0–100 (100 = sem lacunas)
}

// ---------------------------------------------------------------------------------------
// Estrutura fixa do plano
// ---------------------------------------------------------------------------------------
export const MAX_ATIVIDADES = 10;
export const ATIVIDADE_DIAGNOSTICO = 'Diagnóstico empresarial';
export const ATIVIDADE_DEVOLUTIVA = 'Devolutiva do diagnóstico';
export const ATIVIDADE_FERRAMENTAS = 'Desenvolvimento de ferramentas de gestão (planilhas e aplicativos)';
export const ATIVIDADE_ANALISE_FINAL = 'Análise final de atividades da consultoria';
export const ATIVIDADE_RELATORIO_FINAL = 'Preparo de relatórios e apresentação dos resultados';
/** Atividades fixas: 3 no início + 2 no fim. Sobram até 5 para os problemas críticos. */
export const MAX_ATIVIDADES_PROBLEMA = MAX_ATIVIDADES - 5;

export type PapelAtividade = 'diagnostico' | 'devolutiva' | 'ferramentas' | 'problema' | 'analise' | 'relatorio';

const norm = (s: unknown) =>
  String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();

/**
 * Remove repetições de perguntas e respostas do diagnóstico que possam estar anexadas
 * no texto de ações recomendadas (ex.: "... Evidências: \"Pergunta?\" -> Resposta").
 */
export function limparEvidenciasAcoes(texto?: string): string {
  if (!texto) return '';
  return texto
    .replace(/\s*\.?\s*Evidências\s*:[\s\S]*$/i, '.')
    .replace(/\s*\.?\s*Evidencias\s*:[\s\S]*$/i, '.')
    .trim()
    .replace(/\.+$/, '.');
}

/** Limpa e formata descrições de atividades para serem objetivas, limpas e sem numeração repetitiva */
export function formatarDescricaoObjetiva(texto?: string): string {
  if (!texto) return '';
  const linhas = texto
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  const linhasLimpas = linhas.map(l => {
    // Remove numerações repetidas como "1. 1.", "2. 1.", "1.", "1 -", etc. no início
    return l.replace(/^(\d+[\.\-\)]\s*)+/g, '').trim();
  }).filter(Boolean);

  return linhasLimpas.join('\n');
}

/** Identifica o papel de uma atividade pelo nome (aceita os nomes antigos do app). */
export function papelDaAtividade(nome?: string): PapelAtividade {
  const n = norm(nome);
  if (n.includes('devolutiva')) return 'devolutiva';
  if (n.includes('analise final')) return 'analise';
  if (
    n.includes('relatorio final') ||
    n.includes('preparo de relatorio') ||
    n.includes('apresentacao dos resultados') ||
    n.includes('encerramento')
  ) return 'relatorio';
  if (
    n.includes('ferramentas de gestao') ||
    n.includes('desenvolvimento de sistema') ||
    n.includes('desenvolvimento de ferramenta') ||
    (n.includes('desenvolvimento') && (n.includes('planilha') || n.includes('aplicativo')))
  ) return 'ferramentas';
  if (
    n === 'diagnostico' ||
    n.startsWith('diagnostico empresarial') ||
    n.includes('diagnostico inicial') ||
    n.includes('entendimento da demanda')
  ) return 'diagnostico';
  return 'problema';
}

export const isAtividadeDiagnostico = (nome?: string) => papelDaAtividade(nome) === 'diagnostico';
export const isAtividadeRelatorioFinal = (nome?: string) => papelDaAtividade(nome) === 'relatorio';


// ---------------------------------------------------------------------------------------
// Semântica das perguntas
// ---------------------------------------------------------------------------------------
/**
 * Expressões regulares que identificam quando uma pergunta ou o próprio tema do problema
 * investiga uma LIMITAÇÃO, RISCO, DEPENDÊNCIA, FALHA ou GARGALO.
 * Nesses casos:
 *  - "Sim" = O problema/risco ocorre -> É LACUNA / GARGALO (gera ação corretiva).
 *  - "Parcial" = O problema/risco ocorre parcialmente -> É LACUNA MODERADA (gera ação corretiva).
 *  - "Não" = O problema/risco NÃO ocorre (o cliente diversifica, não tem dívida, não tem rejeição, etc.)
 *            -> É PONTO FORTE / CONFORMIDADE (NÃO gera ação corretiva).
 */
const PADRAO_PERGUNTA_NEGATIVA = new RegExp(
  [
    '\\b(depende|dependencia|comercializa|vende|atende|atua|concentra|opera|limita|limitad|restrito) (apenas|somente|exclusivamente|unicamente|em poucos|em um|de um|de poucos)',
    '\\b(apenas|somente|exclusivamente|unicamente) (em |para |por |no |na |)(feiras|festivais|eventos|feira|um cliente|poucos clientes|um comprador|poucos compradores|um fornecedor|poucos fornecedores|uma regiao|um canal|balcao|intermediario|atravessador)',
    '\\b(concentra|concentracao) (em poucos|em um|as vendas|os clientes|de clientes|de vendas|em feiras|em festivais|em canais)',
    '\\b(canal unico|comprador unico|cliente unico|depende de atravessador|depende de intermediario)',
    '\\bja (sofreu|perdeu|atrasou)',
    '\\bja (teve|enfrentou) (prejuizo|perda|atraso|rejeicao|multa|processo|dificuldade|problema|inadimplencia)',
    '\\bfoi (rejeitad|negad|recusad|autuad|multad|notificad|embargad|protestad|negativad)',
    '\\b(tentou|tenta) (acessar|obter) (credito|financiamento|emprestimo) (e foi|sendo) (rejeitad|negad|recusad)',
    '\\b(sofre|enfrenta|apresenta) (com |)(perda|prejuizo|atraso|dificuldade|problema|queda|falta|reclama|mortalidade|inadimpl|gargalo|retrabalho)',
    '\\b(ha|existe|existem|ocorre|ocorrem) (frequentes? |constantes? |)(atraso|perda|desperdicio|reclama|inadimpl|divida|pendencia|falta|mortalidade|retrabalho|devolu|acidente|gargalo|conflito)',
    '\\b(possui|tem) (dividas?|pendencias?|inadimplencia|restricao|restricoes|multas?|dificuldade|problemas?|nome negativado|gargalo|prejuizo)',
    '\\besta (endividad|inadimplent|negativad|irregular)',
    '\\b(mistura|misturam|utiliza|usa) (contas?|financas?|despesas?|gastos?) (pessoais|da familia|de casa)',
    '\\b(retiradas?|saques?) (descontrolad|sem controle|sem criterio|livres?)',
    '\\bpaga (despesas?|contas?) (pessoais|de casa) com (o |)(caixa|dinheiro|recursos) (da empresa|do negocio)',
    '\\b(vende|opera|trabalha) (sem nota|sem emissao|na informalidade|sem alvara|sem licenca|sem registro)',
    '\\bnao (possui|tem|realiza|faz|utiliza|adota|separa|calcula|controla|emite|registra|acompanha)',
    '(^|\\s)(falta|ausencia|inexistencia|sem) (de |)(controle|planejamento|registro|separacao|fluxo de caixa|precificacao)',
    // Preceitos ambientais: práticas destrutivas, degradadoras ou consumo de biomassa predatória
    '\\b(uso de queimadas|faz uso de queimadas|utiliza queimadas|realiza queimadas|queimada|queimadas|fogo no plantio|queima a ceu aberto)\\b',
    '\\b(ha uso de (lenha|madeira)|uso de (lenha|madeira)|utiliza (lenha|madeira)|consome (lenha|madeira)|queima de (lenha|madeira)|fonte poluidora)\\b',
    '\\b(poluicao|polui|contaminacao|contamina|descarte inadequado|descarte irregular|lixo a ceu aberto|esgoto a ceu aberto|efluentes sem tratamento)\\b',
    '\\b(desmatamento|desmata|degradacao do solo|erosao acentuada|assoreamento)\\b',
    '\\b(agrotoxico(s)? (sem|proibid|nociv)|defensivo(s)? sem receituario|aplicacao sem epi)\\b',
    '\\b(infracao ambiental|crime ambiental|multa ambiental|autuacao ambiental|embargo ambiental|notificacao ambiental)\\b',
    // Preceitos trabalhistas e de segurança:
    '\\b(acidente(s)? de trabalho|trabalho (irregular|infantil|analogo|sem registro)|autuac(ao|oes) trabalhista|multa trabalhista)\\b',
  ].join('|'),
  'i'
);

/**
 * Títulos ou descrições de temas/problemas que representam uma vulnerabilidade cadastrada.
 * Mantido apenas como referência/compatibilidade, NÃO deve inverter perguntas de boas práticas.
 */
const PADRAO_PROBLEMA_NEGATIVO = new RegExp(
  [
    '\\b(dependencia de cliente|canal unico|endividamento excessivo|inadimplencia cronica|mistura de contas)\\b',
  ].join('|'),
  'i'
);

export function isPerguntaNegativa(pergunta?: string, _problema?: string): boolean {
  const pNorm = norm(pergunta);
  if (!pNorm) return false;
  return PADRAO_PERGUNTA_NEGATIVA.test(pNorm);
}

/**
 * Perguntas sobre interesse, intenção, projeto futuro ou oportunidade voluntária.
 * Pela teoria administrativa e econômica, a resposta "Não" representa decisão estratégica do gestor
 * (foco no modelo de negócio atual, ausência de demanda pelo insumo/parceria), e NÃO uma lacuna/falha.
 */
const PADRAO_PERGUNTA_INTERESSE_OPCAO = new RegExp(
  '\\b(tem interesse em|teria interesse em|deseja|pretende|planeja|tem intencao de|ha interesse em)\\b',
  'i'
);

export function isPerguntaInteresseOuOpcao(pergunta?: string): boolean {
  const pNorm = norm(pergunta);
  if (!pNorm) return false;
  return PADRAO_PERGUNTA_INTERESSE_OPCAO.test(pNorm);
}

/**
 * Observações que indicam AÇÃO CORRETIVA, PRESCRIÇÃO DE CONSULTORIA ou PROBLEMA REAL.
 * Se o consultor orientou elaborar, capacitar, implantar, treinar, corrigir, adequar, etc.,
 * ou descreveu uma carência/problema, isso representa INEQUIVOCAMENTE uma LACUNA A TRATAR.
 */
const PADRAO_OBS_MANDA_CORRIGIR = new RegExp(
  [
    // Verbos de ação no infinitivo, imperativo ou presente do subjuntivo de consultoria para sanar deficiência
    '\\b(elaborar?|elabore[ms]?|implantar?|implante[ms]?|implementar?|implemente[ms]?|capacitar?|capacite[ms]?|treinar?|treine[ms]?|qualificar?|qualifique[ms]?|desenvolver?|desenvolva[ms]?|padronizar?|padronize[ms]?|formalizar?|formalize[ms]?|estruturar?|estruture[ms]?|organizar?|organize[ms]?|instituir?|institua[ms]?|reestruturar?|reestruture[ms]?|readequar?|readeque[ms]?|ajustar?|ajuste[ms]?|corrigir?|corrija[ms]?|sanar?|sane[ms]?|sanear?|revisar?|revise[ms]?|providenciar?|providencie[ms]?|definir?|defina[ms]?|estabelecer?|estabeleca[ms]?|adotar?|adote[ms]?|introduzir?|introduza[ms]?|criar?|crie[ms]?|planejar?|planeje[ms]?|fazer?|faca[ms]?|renegociar?|renegocie[ms]?|quitar?|quite[ms]?)\\b',
    // Substantivos e locuções de ação corretiva recomendada
    '\\b(elaboracao|implantacao|implementacao|capacitacao|treinamento|qualificacao|padronizacao|formalizacao|estruturacao|reestruturacao|organizacao|readequacao|ajuste|correcao|revisao|saneamento|adocao)\\s+(de|da|do|para|urgente|prioritari)',
    // Expressões prescritivas e mandatórias de consultoria
    '\\b(manda|mando|orienta|recomenda|sugere|determina)\\s+(a |para |)(corrigir|ajustar|elaborar|implantar|capacitar|adequar|sanar|fazer|criar|padronizar)',
    '\\bcorrigir\\s+(o |a |os |as |)(problema|falha|gargalo|deficiencia|situacao|erro|pendencia|processo|defeito|inadimplencia)',
    '\\bdeve(m)?\\s+(ser\\s+)?(corrigid|ajustad|elaborad|implantad|implementad|capacitad|criad|estruturad|padronizad|feit|adotad|revisad|controlad)',
    '\\bprecisa(m)?\\s+(de\\s+)?(corrigir|ajustar|elaborar|implantar|implementar|capacitar|criar|desenvolver|padronizar|estruturar|adotar|organizar|fazer|melhorar|controlar)',
    '\\bnecessita(m)?\\s+(de\\s+)?(correcao|ajuste|elaboracao|implantacao|capacitacao|padronizacao|estrutura|organizacao|melhoria|controle)',
    '\\bnecessari[oa]\\s+(corrigir|ajustar|elaborar|implantar|implementar|capacitar|criar|desenvolver|padronizar|estruturar|adotar|fazer|melhorar|controlar)',
    '\\b(recomenda-se|sugere-se|orienta-se|indica-se)\\s+(a |o |de |)(elaborar|implantar|implementar|capacitar|corrigir|ajustar|adotar|criar|padronizar|estruturar|treinar)',
    '\\b(importante|urgente|prioritari[oa])\\s+(elaborar|implantar|implementar|capacitar|corrigir|ajustar|criar|padronizar|adotar)',
    '\\b(plano\\s+de\\s+acao|acao\\s+corretiva|medida\\s+corretiva|oportunidade\\s+de\\s+melhoria)\\b',
    '\\ba\\s+ser\\s+(elaborad|implantad|implementad|capacitad|corrigid|criad|padronizad|feit|adotad)',
    '\\ba\\s+(elaborar|implantar|implementar|capacitar|corrigir|estruturar|padronizar|ajustar)\\b',
    // Problemas e vulnerabilidades expressas (evitando falsos positivos em "não tem necessidade/problema")
    '\\bnao\\s+(possui|tem|faz|realiza|adota|utiliza|controla|separa|emite|registra|planeja|acompanha|funciona|consegue|aplica)\\s+(nenhum|qualquer|controle|registro|planejamento|processo|rotina|separacao|padrao|documento|fluxo|gestao|ferramenta|nada|recurso|alvara|procedimento|politica|estrategia)\\b',
    '\\b(falta|ausencia|inexistencia|sem)\\s+(de\\s+)?(controle|planejamento|registro|separacao|fluxo|rotina|organizacao|padrao|criterio|processo|documento|gestao|previsao)\\b',
    '\\b(problema(s)?\\s+(grave|relevante|serio|real|critico)|e\\s+um\\s+problema|representa\\s+um\\s+problema|ha\\s+problema|tem\\s+problema)\\b',
    '\\b(gargalo(s)?|deficiencia(s)?|vulnerabilidade(s)?|ponto(s)?\\s+fraco(s)?|debilidade(s)?|desorganizad[oa]|descontrolad[oa]|inadequado|inseguro|vulneravel)\\b',
    '\\b(prejuizo(s)?|desperdicio(s)?|desvio(s)?|perda(s)?\\s+de\\s+(mercadoria|produto|dinheiro|producao|cliente)|retrabalho(s)?\\s+constante|reclamac(ao|oes)\\s+frequente)\\b',
    '\\b(inadimplent|endividament|dividas?\\s+em\\s+atraso|nome\\s+negativad|restricao\\s+cadastral|protesto)\\b',
    '\\b(mistura|misturam)\\s+(contas?|financas?|despesas?)\\s+(pessoais|da\\s+familia)\\b',
    '\\b(retirada(s)?|saque(s)?)\\s+(descontrolad|sem\\s+controle|sem\\s+criterio|livre)\\b',
    '\\b(apenas\\s+no\\s+papel|nao\\s+funciona\\s+na\\s+pratica|nao\\s+e\\s+feito|nao\\s+e\\s+realizado|nao\\s+e\\s+aplicado|informalidade\\s+total|totalmente\\s+informal)\\b',
    '\\bcliente\\s+nao\\s+controla\\b',
  ].join('|'),
  'i'
);

/**
 * Observações / justificativas que comprovam que a prática NÃO SE APLICA ou que a situação
 * JÁ ESTÁ ATENDIDA/REGULARIZADA (ponto forte real ou dispensa legal).
 * NUNCA deve capturar palavras soltas como "boas práticas" ou "conforme", que frequentemente
 * aparecem em sugestões corretivas (ex: "capacitar em boas práticas de fabricação").
 */
const PADRAO_OBS_SEM_PROBLEMA = new RegExp(
  [
    '\\bnao\\s+(se\\s+aplica|e\\s+aplicavel|aplicavel|inaplicavel|se\\s+enquadra|ha\\s+necessidade|existe\\s+necessidade|tem\\s+necessidade|vejo\\s+necessidade|ha\\s+demanda|e\\s+necessari|precisa|necessita|ha\\s+problema|existe\\s+problema|tem\\s+problema|e\\s+um\\s+problema|representa\\s+(um\\s+)?problema|afeta|impacta|ha\\s+impacto|tem\\s+impacto|e\\s+relevante|tem\\s+relevancia|e\\s+prioridade|e\\s+gargalo|e\\s+deficiencia|ha\\s+deficiencia|tem\\s+deficiencia|e\\s+critico)\\b',
    '\\b(desnecessari[oa]|dispensavel|dispensad[oa]s?|isento|isenta|isentos|isentas|irrelevante|nao\\s+aplicavel|inaplicavel|sem\\s+necessidade|sem\\s+impacto|sem\\s+problema|sem\\s+restricao|sem\\s+divida|sem\\s+debito|sem\\s+pendencia|sem\\s+reclamac|sem\\s+atraso|sem\\s+gargalo|sem\\s+perda)\\b',
    '(^|\\s)(n\\/a|n\\.a\\.)(\\s|$|[.,;])',
    '\\b(atende|atendido|atendida|atendem)\\s+(plenamente|adequadamente|bem|satisfatoriamente|de\\s+outra\\s+forma|integralmente)\\b',
    '\\b(funciona|funcionam)\\s+(bem|adequadamente|perfeitamente|satisfatoriamente|redondo|direitinho)\\b',
    '\\b(de|por)\\s+outra\\s+forma\\b|\\bforma\\s+alternativa\\b|\\balternativ(a|o)s?\\s+(adequad|eficaz)',
    '\\b(terceirizad[oa]|terceirizada)\\b|\\b(contador|contabilidade|escritorio|empresa\\s+parceira|cooperativa|associacao|assessoria)\\s+(faz|cuida|realiza|presta|responde|atende|acompanha|gerencia)',
    '\\b(situacao|processo|controle|documentacao|fluxo|rotina)\\s+(esta\\s+|)(regular|adequad[oa]|ok|em\\s+dia|resolvid[oa]|satisfatori[oa]|conforme)\\b',
    '\\bja\\s+(esta\\s+|foi\\s+|)(resolvid[oa]|regularizad[oa]|solucionad[oa]|em\\s+dia|adequad[oa]|implantad[oa]|implementad[oa]|organizado|estruturado)\\b',
    '\\b(proprio|propria|pelo)\\s+(dono|proprietari|gestor|titular|socio)\\s+(faz|cuida|controla|realiza|acompanha|gerencia)',
    '\\b(gestao|controle)\\s+(familiar|simplificad[oa]|diret[oa]|informal)(\\s+(diret[oa]|atende|funciona|suficiente|eficaz))?\\b',
    '\\bmei\\b[^\n.]*\\b(dispensad|isent|nao\\s+exige|nao\\s+obrigad|simplificad)',
    '\\bporte\\s+(pequeno|reduzido|mei|micro|familiar)',
    '\\b(vendas?|vende|comercializa|contato|atendimento)\\s+(diret[oa]|balcao|loja\\s+fisica|pessoal|whatsapp|proprio|propria)',
    '\\b(a\\s+vista|pix|dinheiro)\\b',
    '\\b(tudo\\s+(regular|ok|conforme|em\\s+dia)|totalmente\\s+(conforme|regular|em\\s+dia)|100%\\s+conforme)\\b',
    '\\b(esta|estao)\\s+(regular|regularizad[oa]|em\\s+dia|conforme|adequado|satisfatorio|ok)\\b',
    '\\b(ponto(s)?\\s+forte(s)?|ponto(s)?\\s+positivo(s)?|vantagem\\s+competitiva|atuacao\\s+exemplar|caso\\s+de\\s+sucesso)\\b',
    '\\b(ja\\s+(adota|segue|cumpre|aplica|possui)|empresa\\s+(adota|segue|cumpre|aplica))\\s+(as\\s+|todas\\s+as\\s+)?boas\\s+praticas\\b',
    '\\b(fortalece|fortalecendo|fortalecimento|potencializa|impulsiona|consolida|consolidando|consolidado|consolidada)\\b',
    '\\b(diversifica|diversificando|diversificacao|diversificado|diversificada|multiplos\\s+canais|varios\\s+canais|outros\\s+canais)\\b',
    '\\b(excelente|otimo|otima|otimos|otimas|muito\\s+bom|muito\\s+boa|muito\\s+positivo|muito\\s+positiva)\\b',
    '\\b(parceria(s)?\\s+(ativa|ativas|consolidada|consolidadas|existente|existentes|com\\s+sucesso)|acesso\\s+a\\s+mercados?|venda\\s+institucional|pnae|paa)\\b',
    '\\bnao\\s+(utilizacao|uso|realizacao|adocao|pratica)\\s+de\\s+(queimadas?|fogo|lenha|madeira|agrotoxicos?|defensivos?|poluentes?)\\b',
    '\\b(maturidade\\s+(ambiental|empresarial|de\\s+gestao|produtiva)|preservacao\\s+(do\\s+solo|ambiental|das\\s+nascentes|da\\s+agua|dos\\s+recursos|da\\s+vegetacao))\\b',
    '\\b(evidencia|demonstra|comprova|indica|revela)\\s+(a\\s+|o\\s+|)(maturidade|ponto\\s+forte|conformidade|preservacao|boas\\s+praticas|compromisso|gestao|sustentabilidade)\\b',
    '\\b(sustentabilidade|sustentavel|pratica\\s+sustentavel|conservacao\\s+do\\s+solo|agroecolog(ia|ico|ica))\\b',
    '\\b(conformidade\\s+ambiental|adequacao\\s+ambiental|sem\\s+impacto\\s+ambiental|baixo\\s+impacto|livre\\s+de\\s+(queimadas?|agrotoxicos?|lenha|poluicao))\\b',
  ].join('|'),
  'i'
);

/** Observações que indicam que o controle/prática já existe (só vale para perguntas de controle). */
const PADRAO_OBS_JA_FAZ = new RegExp(
  '\\bja (possui|tem|faz|realiza|atende|controla|utiliza|usa|adota|registra|emite|implanta|implantou|implementou|conta com|participa|atua|mantem|desenvolve)',
  'i'
);

/** Expressões em perguntas de risco/falha indicando que o risco não existe ou está resolvido. */
const PADRAO_OBS_RISCO_RESOLVIDO = new RegExp(
  '\\b(sem (dividas?|debitos?|pendencias?|restricoes?|multas?|atrasos?|reclamacoes?|problemas?|queimadas?|fogo|lenha|madeira|agrotoxicos?|poluicao)|nao (tem|possui|ha|existe|faz|utiliza|usa|realiza) (dividas?|debitos?|pendencias?|restricoes?|queimadas?|queima|lenha|madeira|agrotoxicos?)|ja (quitou|quitad|pagou|pago|regularizad|resolvid)|em dia|tudo regular|tudo ok|situacao regular|diversifica|diversificacao|ponto forte|fortalece|nao depende|outros canais|parcerias?|evidencia a maturidade|preservacao do solo|sustentabilidade)\\b',
  'i'
);

export function observacaoIndicaProblema(observacao?: string): boolean {
  const o = norm(observacao);
  if (!o) return false;
  return PADRAO_OBS_MANDA_CORRIGIR.test(o);
}

export function observacaoIndicaSemProblema(observacao?: string, perguntaNegativa = false): boolean {
  const o = norm(observacao);
  if (!o) return false;
  // Se a observação prescreve ação corretiva ou aponta problema, NUNCA pode ser considerada "sem problema"
  if (PADRAO_OBS_MANDA_CORRIGIR.test(o)) {
    return false;
  }
  if (perguntaNegativa) {
    return PADRAO_OBS_RISCO_RESOLVIDO.test(o) || PADRAO_OBS_SEM_PROBLEMA.test(o);
  }
  if (PADRAO_OBS_SEM_PROBLEMA.test(o)) return true;
  return PADRAO_OBS_JA_FAZ.test(o);
}

// Vereditos de IA mantidos em memória para que relatório, plano e painéis usem o mesmo critério.
const vereditosIA = new Map<string, { sig: string; temProblema: boolean }>();
const assinaturaResposta = (r: { resposta?: string; observacao?: string; pergunta?: string }) =>
  `${norm(r.resposta)}|${norm(r.observacao)}|${norm(r.pergunta)}`;

/** 2 = lacuna total, 1 = parcial, 0 = sem lacuna, null = não respondida. */
export function grauLacuna(
  r: { resposta?: string; pergunta?: string; problema?: string; id?: string; observacao?: string },
  overrideIA?: Map<string, boolean>
): number | null {
  const v = r.resposta;
  if (v !== 'Sim' && v !== 'Não' && v !== 'Parcial') return null;

  const negativa = isPerguntaNegativa(r.pergunta, r.problema);
  const interesseOpcao = isPerguntaInteresseOuOpcao(r.pergunta);
  const mandaCorrigir = observacaoIndicaProblema(r.observacao);
  const temSemProblema = !mandaCorrigir && observacaoIndicaSemProblema(r.observacao, negativa);

  // 1. Preceitos administrativos e econômicos: "Não" em oportunidade ou parceria voluntária é decisão estratégica
  if (interesseOpcao && v === 'Não') {
    return 0; // Decisão do gestor de foco no modelo produtivo atual; não constitui deficiência
  }

  // 2. Preceitos ambientais de práticas nocivas ou consumo de biomassa (queimadas, uso de lenha/madeira, poluição):
  // Se o produtor respondeu "Não", ele NÃO pratica o ato danoso, o que é conformidade ambiental plena.
  const isPraticaNocivaAmbiental = /\b(queimada|queimadas|fogo|lenha|madeira|poluicao|descarte inadequado|agrotoxico)\b/i.test(r.pergunta || '');
  if (isPraticaNocivaAmbiental && v === 'Não') {
    return 0; // Ausência de queimada/lenha/poluição é ponto forte ambiental absoluto
  }

  // 3. A OBSERVAÇÃO DO CONSULTOR TEM SOBERANIA NA ANÁLISE:
  // Se a observação manda corrigir, elaborar, capacitar, implantar, estruturar ou descreve deficiência/problema real:
  if (mandaCorrigir) {
    // É lacuna a tratar: 1 se for Parcial, 2 se for Não ou Sim onde a prática requer correção/elaboração
    return v === 'Parcial' ? 1 : 2;
  }

  // Se a observação expressa inequivocamente ausência de problema (dispensa legal, MEI isento, já resolvido/atendido):
  if (temSemProblema) {
    return 0;
  }

  // 4. Veredito de IA em memória (apenas se resposta/observação não mudaram)
  if (r.id) {
    const cache = vereditosIA.get(r.id);
    if (cache && cache.sig === assinaturaResposta(r)) {
      if (!cache.temProblema) return 0;
      return v === 'Parcial' ? 1 : 2;
    }
  }

  // 5. Respeita o veredito da IA se houver
  if (overrideIA) {
    if (r.id && overrideIA.has(r.id)) {
      const temProb = overrideIA.get(r.id)!;
      if (!temProb) return 0;
      return v === 'Parcial' ? 1 : 2;
    }
    const chaveQ = norm(r.pergunta);
    if (chaveQ && overrideIA.has(chaveQ)) {
      const temProb = overrideIA.get(chaveQ)!;
      if (!temProb) return 0;
      return v === 'Parcial' ? 1 : 2;
    }
  }

  // 6. Resposta Parcial sem observação conclusiva indica oportunidade de melhoria/lacuna parcial
  if (v === 'Parcial') return 1;

  // 7. Sem observação conclusiva: segue a semântica da pergunta
  if (negativa) {
    // Pergunta de risco/falha/limitação: "Sim" é o problema (2). "Não" significa que NÃO há problema (0).
    return v === 'Sim' ? 2 : 0;
  }
  // Pergunta de controle/boa prática: "Não" é a falta de controle (2). "Sim" significa que tem (0).
  return v === 'Não' ? 2 : 0;
}

export interface StatusRespostaInfo {
  tipo: 'pendente' | 'conforme' | 'conforme_obs' | 'conforme_contexto' | 'parcial' | 'lacuna';
  label: string;
  detalhe: string;
  isConforme: boolean;
}

/**
 * Classifica a resposta e gera metadados explicativos para a interface e relatórios,
 * garantindo total transparência quando observações ou o contexto da pergunta
 * justificam a conformidade sem gerar lacuna.
 */
export function classificarStatusResposta(
  r: { resposta?: string; pergunta?: string; problema?: string; observacao?: string; id?: string },
  overrideIA?: Map<string, boolean>
): StatusRespostaInfo {
  if (!r.resposta) {
    return {
      tipo: 'pendente',
      label: 'Pendente',
      detalhe: 'Aguardando preenchimento no diagnóstico',
      isConforme: false,
    };
  }
  const lac = grauLacuna(r, overrideIA);
  const negativa = isPerguntaNegativa(r.pergunta, r.problema);
  const mandaCorrigir = observacaoIndicaProblema(r.observacao);
  const temObsOk = !mandaCorrigir && observacaoIndicaSemProblema(r.observacao, negativa);

  if (lac === 0) {
    if (negativa && r.resposta === 'Não') {
      return {
        tipo: 'conforme_contexto',
        label: 'Ponto Forte (Sem Risco)',
        detalhe: 'Resposta Não em questão ambiental, regulatória ou de risco confirma ausência de vulnerabilidade na empresa',
        isConforme: true,
      };
    }
    if (isPerguntaInteresseOuOpcao(r.pergunta) && r.resposta === 'Não') {
      return {
        tipo: 'conforme_contexto',
        label: 'Conforme (Decisão Estratégica)',
        detalhe: 'Resposta Não em oportunidade ou parceria voluntária reflete alinhamento ao modelo de negócio e foco produtivo atual',
        isConforme: true,
      };
    }
    if (!negativa && r.resposta === 'Sim') {
      return {
        tipo: 'conforme',
        label: 'Atendido / Conforme',
        detalhe: temObsOk
          ? 'Prática de gestão plenamente atendida e reforçada pela observação do consultor'
          : 'Prática de gestão plenamente atendida',
        isConforme: true,
      };
    }
    if (temObsOk) {
      return {
        tipo: 'conforme_obs',
        label: 'Conforme (Justificado)',
        detalhe: 'Observação do consultor indica ausência de problema, dispensa ou controle existente',
        isConforme: true,
      };
    }
    return {
      tipo: 'conforme',
      label: 'Atendido / Conforme',
      detalhe: 'Prática de gestão plenamente atendida',
      isConforme: true,
    };
  }

  if (lac === 1) {
    return {
      tipo: 'parcial',
      label: 'Atendimento Parcial',
      detalhe: mandaCorrigir
        ? 'Observação do consultor prescreveu ação de melhoria ou identificou vulnerabilidade moderada'
        : 'Prática parcialmente estruturada; oportunidade de melhoria',
      isConforme: false,
    };
  }

  return {
    tipo: 'lacuna',
    label: mandaCorrigir
      ? 'Lacuna a Tratar (Ação Recomendada)'
      : 'Lacuna a Tratar',
    detalhe: mandaCorrigir
      ? 'Observação do consultor prescreve ações corretivas (elaboração, capacitação, implantação ou ajuste) para sanar a deficiência'
      : 'Oportunidade de melhoria que gerará ação no Plano de Ação',
    isConforme: false,
  };
}

function pesoImpacto(impacto?: string): number {
  const n = norm(impacto);
  if (n.startsWith('alt') || n.startsWith('critic')) return 3;
  if (n.startsWith('baix')) return 1;
  return 2; // Médio ou não informado
}

const ORDEM_NIVEL: Record<NivelCriticidade, number> = { Crítico: 3, Alto: 2, Moderado: 1 };

// ---------------------------------------------------------------------------------------
// Análise
// ---------------------------------------------------------------------------------------
export function analisarDiagnostico(
  respostas: Resposta[],
  problemas: Problema[] = [],
  solucoes: Solucao[] = [],
  overrideIA?: Map<string, boolean>
): AnaliseDiagnostico {
  const grupos = new Map<string, Resposta[]>();
  let totalRespondidas = 0;
  let somaLacuna = 0;
  let somaMax = 0;

  for (const r of respostas || []) {
    if (!r) continue;
    const g = grauLacuna(r, overrideIA);
    if (g === null) continue;
    totalRespondidas++;
    const peso = Number(r.peso) > 0 ? Number(r.peso) : 1;
    somaLacuna += g * peso;
    somaMax += 2 * peso;
    const chave = r.idProblema || norm(r.problema) || 'sem_problema';
    if (!grupos.has(chave)) grupos.set(chave, []);
    grupos.get(chave)!.push(r);
  }

  const analisados: ProblemaAnalisado[] = [];
  const pontosFortes: AnaliseDiagnostico['pontosFortes'] = [];

  for (const [chave, lista] of grupos) {
    const ref = lista[0];
    const cadastro =
      problemas.find((p) => p.id === ref.idProblema) ||
      problemas.find((p) => norm(p.descricao_problemas) === norm(ref.problema));
    const nome = cadastro?.descricao_problemas || ref.problema || 'Problema identificado';
    const area = cadastro?.area || ref.area || 'Geral';
    const impacto = cadastro?.impacto || 'Médio';

    let lac = 0;
    let max = 0;
    const lacunas: Lacuna[] = [];
    for (const r of lista) {
      const g = grauLacuna(r, overrideIA)!;
      const peso = Number(r.peso) > 0 ? Number(r.peso) : 1;
      lac += g * peso;
      max += 2 * peso;
      if (g > 0) lacunas.push({ pergunta: r.pergunta, resposta: r.resposta || '', observacao: r.observacao || undefined });
    }
    const intensidade = max > 0 ? lac / max : 0;

    if (intensidade === 0) {
      pontosFortes.push({ problema: nome, area });
      continue;
    }

    const imp = pesoImpacto(impacto);
    const temLacunaTotal = lista.some((r) => grauLacuna(r, overrideIA) === 2);
    let nivel: NivelCriticidade = 'Moderado';
    if ((imp >= 3 && temLacunaTotal) || (intensidade >= 0.75 && imp >= 2)) nivel = 'Crítico';
    else if ((intensidade >= 0.5 && imp >= 2) || imp >= 3) nivel = 'Alto';

    const solucao =
      solucoes.find((s) => s.idProblema && s.idProblema === (cadastro?.id || ref.idProblema)) ||
      solucoes.find((s) => norm(s.problema) === norm(nome));

    analisados.push({
      idProblema: cadastro?.id || ref.idProblema || chave,
      problema: nome,
      area,
      impacto,
      nivel,
      pontuacao: Math.round(intensidade * 100 * (0.5 + imp / 6)),
      intensidade,
      lacunas,
      solucao,
    });
  }

  analisados.sort(
    (a, b) => ORDEM_NIVEL[b.nivel] - ORDEM_NIVEL[a.nivel] || b.pontuacao - a.pontuacao || a.problema.localeCompare(b.problema)
  );

  return {
    totalRespondidas,
    problemas: analisados,
    pontosFortes,
    scoreGeral: somaMax > 0 ? Math.round((1 - somaLacuna / somaMax) * 100) : 0,
  };
}

// ---------------------------------------------------------------------------------------
// Plano de ação (atividades do cronograma)
// ---------------------------------------------------------------------------------------
const PRIORIDADE: Record<NivelCriticidade, AtividadeCronograma['prioridade']> = {
  Crítico: 'Alta',
  Alto: 'Alta',
  Moderado: 'Média',
};

/** Ferramentas de gestão sugeridas conforme a área dos problemas. */
const FERRAMENTAS_POR_AREA: [RegExp, string][] = [
  [/financ|caixa|contab|tribut|fiscal|credito|divida|inadimpl/, 'planilha de fluxo de caixa e DRE gerencial'],
  [/custo|preco|precifica|margem/, 'planilha de custos e formação de preço de venda'],
  [/comerc|venda|cliente|mercado|marketing|atendimento/, 'controle de vendas, clientes e metas comerciais'],
  [/estoque|compra|suprimento|insumo|almoxarif/, 'controle de estoque e compras'],
  [/produc|operac|process|qualidade|manejo|logist/, 'controle de produção e indicadores operacionais'],
  [/pessoa|rh|equipe|gestao de pessoas|colaborador/, 'controle de rotinas, escalas e desempenho da equipe'],
  [/ambient|legal|juridic|licen|document|regulariz|sanitar/, 'checklist de conformidade e agenda de vencimentos'],
  [/estrateg|planejamento|gestao|governanca|inovac/, 'painel de indicadores e plano de metas'],
];

export function ferramentasSugeridas(analise?: AnaliseDiagnostico): string[] {
  const out: string[] = [];
  for (const p of analise?.problemas || []) {
    const alvo = norm(`${p.area} ${p.problema}`);
    for (const [re, ferramenta] of FERRAMENTAS_POR_AREA) {
      if (re.test(alvo) && !out.includes(ferramenta)) out.push(ferramenta);
    }
  }
  if (out.length === 0) out.push('planilha de fluxo de caixa e DRE gerencial', 'painel de indicadores e plano de metas');
  return out.slice(0, 4);
}

function listaTexto(itens: string[]): string {
  if (itens.length <= 1) return itens[0] || '';
  return `${itens.slice(0, -1).join(', ')} e ${itens[itens.length - 1]}`;
}

export function atividadeDiagnostico(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const areas = Array.from(new Set((analise?.problemas || []).map((p) => p.area))).slice(0, 5);
  return {
    nome: ATIVIDADE_DIAGNOSTICO,
    descricao:
      'Entendimento da demanda, aplicação do questionário de diagnóstico, levantamento de dados e documentos e ' +
      `análise da situação atual da empresa${areas.length ? ` (áreas: ${areas.join(', ')})` : ''}.`,
    cargaHoraria: '4h',
    solucaoProposta: 'Diagnóstico empresarial com levantamento das lacunas de gestão',
    resultadoEsperado: 'Diagnóstico aplicado, com 100% das perguntas respondidas e problemas mapeados por área.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeDevolutiva(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const criticos = (analise?.problemas || []).filter((p) => p.nivel === 'Crítico').map((p) => p.problema);
  const prioritarios = criticos.length ? criticos : (analise?.problemas || []).slice(0, 3).map((p) => p.problema);
  return {
    nome: ATIVIDADE_DEVOLUTIVA,
    descricao:
      'Apresentação dos resultados do diagnóstico ao empresário, validação dos problemas identificados, ' +
      'priorização conjunta e pactuação do plano de ação' +
      (prioritarios.length ? `, com foco em: ${prioritarios.slice(0, 5).join('; ')}.` : '.'),
    cargaHoraria: '2h',
    solucaoProposta: 'Devolutiva do diagnóstico e pactuação do plano de ação com o cliente',
    resultadoEsperado: 'Problemas críticos validados pelo empresário e plano de ação aprovado.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeFerramentas(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const ferramentas = ferramentasSugeridas(analise);
  return {
    nome: ATIVIDADE_FERRAMENTAS,
    descricao:
      `Desenvolvimento e implantação de ferramentas de gestão sob medida: ${listaTexto(ferramentas)}. ` +
      'Inclui parametrização com os dados da empresa e treinamento do responsável pelo uso.',
    cargaHoraria: '8h',
    solucaoProposta: 'Planilhas e aplicativos de gestão personalizados para os problemas identificados',
    resultadoEsperado: `Ferramentas implantadas e em uso na rotina da empresa (${ferramentas.length} controle(s) ativo(s)).`,
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeAnaliseFinal(): AtividadeCronograma {
  return {
    nome: ATIVIDADE_ANALISE_FINAL,
    descricao:
      'Análise crítica de todas as atividades executadas: verificação das entregas, mensuração dos indicadores ' +
      'em relação às metas do plano e registro das recomendações para a continuidade após a consultoria.',
    cargaHoraria: '2h',
    solucaoProposta: 'Análise final e consolidação dos resultados da consultoria',
    resultadoEsperado: 'Checklist de entregas concluído e indicadores de cada atividade apurados em relação às metas.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

export function atividadeRelatorioFinal(): AtividadeCronograma {
  return {
    nome: ATIVIDADE_RELATORIO_FINAL,
    descricao:
      'Consolidação das atividades executadas e dos indicadores alcançados, elaboração do relatório técnico da ' +
      'consultoria e apresentação dos resultados e recomendações ao empresário.',
    cargaHoraria: '4h',
    solucaoProposta: 'Elaboração dos relatórios e apresentação dos resultados',
    resultadoEsperado: 'Relatório final entregue e resultados apresentados e validados com o cliente.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

const PESO_NIVEL: Record<NivelCriticidade, number> = { Crítico: 3, Alto: 2, Moderado: 1 };

/** Converte um problema analisado no item que será tratado (e virará tarefa). */
export function problemaTratado(p: ProblemaAnalisado): ProblemaTratado {
  const s = p.solucao;
  return {
    idProblema: p.idProblema,
    problema: p.problema,
    area: p.area,
    nivel: p.nivel,
    acao: s?.solucao_recomendada || `Corrigir: ${p.problema}`,
    solucao: s?.solucao_recomendada || `Plano de correção para ${p.problema.toLowerCase()}`,
    passos:
      limparEvidenciasAcoes(s?.acoes_sugeridas) ||
      'Levantar as causas com o empresário, definir as correções necessárias e implantar o controle na rotina.',
    resultadoEsperado:
      [s?.resultado_esperado, s?.kpis_sugeridos ? `Indicadores: ${s.kpis_sugeridos}` : ''].filter(Boolean).join(' — ') ||
      `Problema "${p.problema}" solucionado, com controle implantado e acompanhado mensalmente.`,
    responsavel: s?.responsavel_sugerido || 'Consultor/Cliente',
  };
}

export interface GrupoArea {
  area: string;
  problemas: ProblemaAnalisado[];
  nivelMax: NivelCriticidade;
  peso: number;
}

/** Agrupa os problemas por área, da área mais crítica para a menos crítica. */
export function agruparPorArea(analise: AnaliseDiagnostico): GrupoArea[] {
  const mapa = new Map<string, ProblemaAnalisado[]>();
  for (const p of analise.problemas) {
    const area = (p.area || 'Geral').trim() || 'Geral';
    if (!mapa.has(area)) mapa.set(area, []);
    mapa.get(area)!.push(p);
  }
  const grupos: GrupoArea[] = Array.from(mapa, ([area, lista]) => {
    const problemas = [...lista].sort((a, b) => ORDEM_NIVEL[b.nivel] - ORDEM_NIVEL[a.nivel] || b.pontuacao - a.pontuacao);
    return {
      area,
      problemas,
      nivelMax: problemas[0].nivel,
      peso: problemas.reduce((s, p) => s + PESO_NIVEL[p.nivel], 0),
    };
  });
  const conta = (g: GrupoArea, n: NivelCriticidade) => g.problemas.filter((p) => p.nivel === n).length;
  return grupos.sort(
    (a, b) =>
      conta(b, 'Crítico') - conta(a, 'Crítico') ||
      conta(b, 'Alto') - conta(a, 'Alto') ||
      b.peso - a.peso ||
      b.problemas[0].pontuacao - a.problemas[0].pontuacao ||
      a.area.localeCompare(b.area)
  );
}

function resumoNiveis(itens: { nivel: NivelCriticidade }[]): string {
  const c = itens.filter((p) => p.nivel === 'Crítico').length;
  const a = itens.filter((p) => p.nivel === 'Alto').length;
  const m = itens.filter((p) => p.nivel === 'Moderado').length;
  return [c && `${c} crítico${c > 1 ? 's' : ''}`, a && `${a} alto${a > 1 ? 's' : ''}`, m && `${m} moderado${m > 1 ? 's' : ''}`]
    .filter(Boolean)
    .join(', ');
}

/** (Re)monta os textos de uma atividade de área a partir dos problemas tratados. */
export function textosDaAtividadeDeArea(area: string, itens: ProblemaTratado[]): Pick<AtividadeCronograma, 'nome' | 'descricao' | 'solucaoProposta' | 'resultadoEsperado' | 'responsavel'> {
  const plural = itens.length > 1;
  const descLinhas = itens
    .map(p => formatarDescricaoObjetiva(p.passos))
    .filter(Boolean);

  // Considera os problemas reais a serem solucionados no título da atividade
  const nomesProblemas = itens.map(p => p.problema.trim()).filter(Boolean);
  let tituloProblemas = '';
  if (nomesProblemas.length === 1) {
    tituloProblemas = nomesProblemas[0];
  } else if (nomesProblemas.length === 2) {
    tituloProblemas = `${nomesProblemas[0]} e ${nomesProblemas[1]}`;
  } else if (nomesProblemas.length > 2) {
    tituloProblemas = `${nomesProblemas.slice(0, 2).join(', ')} e outros (+${nomesProblemas.length - 2})`;
  } else {
    tituloProblemas = `${itens.length} problema${plural ? 's' : ''}`;
  }

  const tituloFinal = area && area !== 'Geral' && area !== 'Demais áreas'
    ? `${area}: Resolução de ${tituloProblemas} (${resumoNiveis(itens)})`
    : `Resolução de ${tituloProblemas} (${resumoNiveis(itens)})`;

  return {
    nome: tituloFinal,
    descricao: descLinhas.join('\n'),
    solucaoProposta: itens.map((p) => `[${p.nivel}] ${p.solucao}`).join('; '),
    resultadoEsperado: itens.map((p) => p.resultadoEsperado).join(' | '),
    responsavel: itens.every((p) => p.responsavel === itens[0].responsavel) ? itens[0].responsavel : 'Consultor/Cliente',
  };
}

function atividadeDaArea(area: string, problemas: ProblemaAnalisado[]): AtividadeCronograma {
  const itens = problemas.map(problemaTratado);
  const nivelMax = problemas.reduce<NivelCriticidade>((m, p) => (ORDEM_NIVEL[p.nivel] > ORDEM_NIVEL[m] ? p.nivel : m), 'Moderado');
  return {
    ...textosDaAtividadeDeArea(area, itens),
    cargaHoraria: '4h',
    idProblema: itens[0]?.idProblema,
    status: 'Pendente',
    prioridade: PRIORIDADE[nivelMax],
    area,
    criticidade: nivelMax,
    problemasTratados: itens,
  };
}

/**
 * Atividades de área. Cabem `limite` atividades; se houver mais áreas, as MENOS críticas são
 * reunidas numa atividade "Demais áreas", que ainda trata cada problema individualmente.
 */
export function atividadesPorArea(analise: AnaliseDiagnostico, limite: number): AtividadeCronograma[] {
  const grupos = agruparPorArea(analise);
  const max = Math.max(1, limite);
  if (grupos.length <= max) return grupos.map((g) => atividadeDaArea(g.area, g.problemas));
  const principais = grupos.slice(0, max - 1).map((g) => atividadeDaArea(g.area, g.problemas));
  const resto = grupos.slice(max - 1);
  const rotulo = resto.length <= 3 ? resto.map((g) => g.area).join(', ') : 'Demais áreas';
  return [...principais, atividadeDaArea(rotulo, resto.flatMap((g) => g.problemas))];
}

function atividadeComplementar(extras: { nome: string; nivel?: string; problemasTratados?: ProblemaTratado[] }[]): AtividadeCronograma {
  const tratados = extras.flatMap((e) => e.problemasTratados || []);
  if (tratados.length) {
    const area = Array.from(new Set(tratados.map((t) => t.area))).slice(0, 3).join(', ') || 'Demais áreas';
    const nivelMax = tratados.some((t) => t.nivel === 'Crítico') ? 'Crítico' : tratados.some((t) => t.nivel === 'Alto') ? 'Alto' : 'Moderado';
    return {
      ...textosDaAtividadeDeArea(area, tratados),
      cargaHoraria: '4h',
      idProblema: tratados[0].idProblema,
      status: 'Pendente',
      prioridade: PRIORIDADE[nivelMax],
      area,
      criticidade: nivelMax,
      problemasTratados: tratados,
    };
  }
  return {
    nome: 'Ações complementares para os demais problemas identificados',
    descricao: `Orientação e plano de correção para: ${extras.map((p) => (p.nivel ? `${p.nome} (${p.nivel.toLowerCase()})` : p.nome)).join('; ')}.`,
    cargaHoraria: '4h',
    solucaoProposta: 'Orientações técnicas e plano de correção para os problemas complementares',
    resultadoEsperado: 'Empresário orientado e com responsável e prazo definidos para cada problema complementar.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: extras.some((p) => p.nivel && p.nivel !== 'Moderado') ? 'Alta' : 'Média',
  };
}

function atividadeSemLacunas(): AtividadeCronograma {
  return {
    nome: 'Fortalecimento dos controles e indicadores de gestão',
    descricao: 'O diagnóstico não apontou lacunas relevantes: revisão dos controles existentes, definição de metas e rotina de acompanhamento de indicadores.',
    cargaHoraria: '4h',
    solucaoProposta: 'Rotina de acompanhamento de indicadores e metas',
    resultadoEsperado: 'Metas definidas e indicadores acompanhados mensalmente pelo empresário.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

/** Peso de uma atividade de área na divisão das horas (críticos pesam mais). */
function pesoDaAtividade(a: AtividadeCronograma): number {
  const itens = a.problemasTratados || [];
  if (!itens.length) return 1;
  return itens.reduce((s, p) => s + PESO_NIVEL[p.nivel], 0);
}

/** Horas (número) de uma atividade. */
export function horasDe(a?: Pick<AtividadeCronograma, 'cargaHoraria'>): number {
  const n = parseInt(String(a?.cargaHoraria || '').replace(/\D/g, ''), 10);
  return Number.isFinite(n) ? n : 0;
}

export function parseCargaHoraria(total?: string | number, padrao = 34): number {
  const n = typeof total === 'number' ? total : parseInt(String(total || '').replace(/\D/g, ''), 10);
  return Number.isFinite(n) && n > 0 ? n : padrao;
}

/** Quantas atividades de problema cabem na carga horária (mínimo 2h cada, máximo 6). */
export function capacidadeAtividades(cargaHorariaTotal: number): number {
  const minimoFixo = 12; // diagnóstico 2h + devolutiva 2h + ferramentas 4h + análise final 2h + relatório 2h
  return Math.max(1, Math.min(MAX_ATIVIDADES_PROBLEMA, Math.floor((cargaHorariaTotal - minimoFixo) / 2)));
}

/**
 * Distribui a carga horária total entre as atividades, somando EXATAMENTE o total.
 * Padrão: diagnóstico 4h, devolutiva 2h, ferramentas 8h, análise final 2h (fixa), relatório 4h
 * e 4h por problema. Com pouca carga, reduz até os mínimos (2h, 2h, 4h, 2h, 2h e 2h por problema).
 */
export function distribuirCargaHoraria(atividades: AtividadeCronograma[], cargaHorariaTotal: number): AtividadeCronograma[] {
  const lista = atividades.map((a) => ({ ...a }));
  const papeis = lista.map((a) => papelDaAtividade(a.nome));
  const idx = (p: PapelAtividade) => papeis.indexOf(p);
  const probs = papeis.map((p, i) => (p === 'problema' ? i : -1)).filter((i) => i >= 0);
  const n = probs.length;

  const padrao: Record<string, number> = { diagnostico: 4, devolutiva: 2, ferramentas: 8, analise: 2, relatorio: 4 };
  const minimo: Record<string, number> = { diagnostico: 2, devolutiva: 2, ferramentas: 4, analise: 2, relatorio: 2 };
  const fixos = (['diagnostico', 'devolutiva', 'ferramentas', 'analise', 'relatorio'] as const).filter((p) => idx(p) >= 0);
  const horas = new Map<number, number>();

  const somaPadrao = fixos.reduce((s, p) => s + padrao[p], 0);
  const somaMin = fixos.reduce((s, p) => s + minimo[p], 0);
  let total = Math.max(cargaHorariaTotal, somaMin + 2 * n);

  if (total >= somaPadrao + 2 * n) {
    fixos.forEach((p) => horas.set(idx(p), padrao[p]));
    let resto = total - somaPadrao;
    if (n > 0) {
      // Cada área recebe no mínimo 2h; o restante é dividido pelo peso dos problemas
      // (crítico 3, alto 2, moderado 1), até 8h por área — áreas mais críticas recebem mais.
      probs.forEach((i) => horas.set(i, 2));
      resto -= 2 * n;
      const pesos = probs.map((i) => pesoDaAtividade(lista[i]));
      const somaPesos = pesos.reduce((a, b) => a + b, 0) || 1;
      const alvoTotal = Math.min(resto, 6 * n);
      const extras = probs.map((_, k) => Math.min(6, (alvoTotal * pesos[k]) / somaPesos));
      const inteiros = extras.map(Math.floor);
      probs.forEach((i, k) => horas.set(i, horas.get(i)! + inteiros[k]));
      resto -= inteiros.reduce((a, b) => a + b, 0);
      let sobraAlvo = alvoTotal - inteiros.reduce((a, b) => a + b, 0);
      const ordem = probs.map((_, k) => k).sort((x, y) => extras[y] - inteiros[y] - (extras[x] - inteiros[x]) || pesos[y] - pesos[x]);
      for (const k of ordem) {
        if (sobraAlvo <= 0 || resto <= 0) break;
        if (horas.get(probs[k])! < 8) { horas.set(probs[k], horas.get(probs[k])! + 1); resto--; sobraAlvo--; }
      }
    }
    // Sobra de horas: ferramentas (até 12h), depois problemas (até 8h), depois diagnóstico e relatório.
    const alvos: [number, number][] = [
      [idx('ferramentas'), 12],
      ...probs.map((i) => [i, 8] as [number, number]),
      [idx('diagnostico'), 8],
      [idx('relatorio'), 8],
      [idx('devolutiva'), 4],
      // A análise final fica sempre com 2h.
    ].filter(([i]) => i >= 0) as [number, number][];
    let guarda = 0;
    while (resto > 0 && guarda++ < 1000) {
      let aplicou = false;
      for (const [i, max] of alvos) {
        if (resto <= 0) break;
        if ((horas.get(i) || 0) < max) { horas.set(i, (horas.get(i) || 0) + 1); resto--; aplicou = true; }
      }
      if (!aplicou) { const alvo = idx('ferramentas') >= 0 ? idx('ferramentas') : probs[0] ?? 0; horas.set(alvo, (horas.get(alvo) || 0) + resto); resto = 0; }
    }
  } else {
    probs.forEach((i) => horas.set(i, 2));
    fixos.forEach((p) => horas.set(idx(p), minimo[p]));
    let resto = total - 2 * n - somaMin;
    for (const p of ['ferramentas', 'diagnostico', 'relatorio', 'devolutiva'] as const) {
      const i = idx(p);
      if (i < 0) continue;
      const add = Math.min(resto, padrao[p] - minimo[p]);
      horas.set(i, horas.get(i)! + add);
      resto -= add;
    }
  }

  lista.forEach((a, i) => { if (horas.has(i)) a.cargaHoraria = `${horas.get(i)}h`; });
  return lista;
}

/**
 * Normaliza QUALQUER lista (gerada, replicada, de modelo ou editada) para a estrutura:
 *   1. Diagnóstico · 2. Devolutiva · 3. Ferramentas de gestão · 4..8 problemas ·
 *   penúltima: Análise final (2h) · última: Relatórios.
 * Mantém no máximo 10 atividades (excedentes viram "Ações complementares") e ajusta as horas.
 */
export function estruturarPlano(
  lista: AtividadeCronograma[],
  cargaHorariaTotal: number | string | undefined,
  analise?: AnaliseDiagnostico
): AtividadeCronograma[] {
  const total = parseCargaHoraria(cargaHorariaTotal);
  const itens = (lista || []).filter((a) => a && (a.nome || a.solucaoProposta));
  const pega = (p: PapelAtividade) => itens.find((a) => papelDaAtividade(a.nome) === p);

  const diag = pega('diagnostico');
  const devol = pega('devolutiva');
  const ferr = pega('ferramentas');
  const analiseFinal = pega('analise');
  const rel = pega('relatorio');
  let meio = itens.filter((a) => papelDaAtividade(a.nome) === 'problema');

  const limite = Math.min(MAX_ATIVIDADES_PROBLEMA, capacidadeAtividades(total));
  if (meio.length > limite) {
    const extras = meio.slice(limite - 1);
    meio = [...meio.slice(0, limite - 1), atividadeComplementar(extras.map((e) => ({ nome: e.nome, nivel: e.criticidade, problemasTratados: e.problemasTratados })))];
  }
  if (meio.length === 0) meio = [atividadeSemLacunas()];

  const estruturado: AtividadeCronograma[] = [
    diag ? { ...diag, nome: ATIVIDADE_DIAGNOSTICO } : atividadeDiagnostico(analise),
    devol ? { ...devol, nome: ATIVIDADE_DEVOLUTIVA } : atividadeDevolutiva(analise),
    ferr ? { ...ferr, nome: ATIVIDADE_FERRAMENTAS } : atividadeFerramentas(analise),
    ...meio,
    analiseFinal ? { ...analiseFinal, nome: ATIVIDADE_ANALISE_FINAL } : atividadeAnaliseFinal(),
    rel ? { ...rel, nome: ATIVIDADE_RELATORIO_FINAL } : atividadeRelatorioFinal(),
  ];
  return distribuirCargaHoraria(estruturado, total).map((a, i) => ({ ...a, ordem: i }));
}

/**
 * Plano completo a partir da análise do diagnóstico: uma atividade por ÁREA com lacunas
 * (áreas mais críticas primeiro), cada uma tratando todos os problemas daquela área.
 */
export function montarAtividades(analise: AnaliseDiagnostico, cargaHorariaTotal: number | string | undefined): AtividadeCronograma[] {
  const total = parseCargaHoraria(cargaHorariaTotal);
  const meio = atividadesPorArea(analise, capacidadeAtividades(total));
  return estruturarPlano(
    [atividadeDiagnostico(analise), atividadeDevolutiva(analise), atividadeFerramentas(analise), ...meio, atividadeAnaliseFinal(), atividadeRelatorioFinal()],
    total,
    analise
  );
}

/** @deprecated use estruturarPlano. Mantido para compatibilidade. */
export function garantirPrimeiraEUltima(lista: AtividadeCronograma[], analise?: AnaliseDiagnostico): AtividadeCronograma[] {
  const total = (lista || []).reduce((s, a) => s + horasDe(a), 0);
  return estruturarPlano(lista, total || undefined, analise);
}

// ---------------------------------------------------------------------------------------
// IA: só melhora os TEXTOS (a estrutura, a ordem e as horas são garantidas pelo código)
// ---------------------------------------------------------------------------------------
export interface GeradorIA {
  models: { generateContent: (args: { model: string; contents: any; config?: any }) => Promise<{ text?: string }> };
}

/**
 * Avalia as respostas do diagnóstico com IA (Gemini) para determinar com precisão semântica
 * se cada resposta dada representa um PROBLEMA/GARGALO REAL que exige ação corretiva (true)
 * ou se é uma SITUAÇÃO POSITIVA/DIVERSIFICAÇÃO/BOA PRÁTICA onde NÃO se deve gerar ação corretiva (false).
 */
export async function avaliarRespostasComIA(
  ai: GeradorIA | null | undefined,
  respostas: Resposta[]
): Promise<Map<string, boolean>> {
  const mapa = new Map<string, boolean>();
  if (!ai || !Array.isArray(respostas) || respostas.length === 0) return mapa;

  const validas = respostas.filter(r => r && (r.resposta === 'Sim' || r.resposta === 'Não' || r.resposta === 'Parcial'));
  if (validas.length === 0) return mapa;

  const itensTodos = validas.map((r, idx) => ({
    id: r.id || `item_${idx}`,
    pergunta: r.pergunta || '',
    problema: r.problema || '',
    resposta: r.resposta,
    observacao: r.observacao || ''
  }));
  const sigPorId = new Map<string, string>(validas.map((r, idx) => [r.id || `item_${idx}`, assinaturaResposta(r)]));

  const montarPrompt = (itens: typeof itensTodos) => `Como consultor empresarial sênior do SEBRAE especialista em diagnósticos de gestão e comércio:
Analise cada par de pergunta, problema e resposta dada pelo empresário. Para CADA item, determine se a resposta indica a EXISTÊNCIA DE UM PROBLEMA OU GARGALO REAL que necessita de ação corretiva ("temProblema": true) ou se representa AUSÊNCIA DE PROBLEMA (situação positiva, conformidade ou diversificação) onde NÃO deve ser gerada ação corretiva ("temProblema": false).

REGRA MAIS IMPORTANTE: uma resposta negativa ("Não") NEM SEMPRE significa deficiência ou requisito não atendido. Avalie o CONTEXTO da pergunta (o tema, o porte e a atividade da empresa) e, principalmente, a OBSERVAÇÃO/SUGESTÃO registrada pelo consultor, que tem prioridade sobre a resposta literal. Exemplos de "Não" que NÃO são problema: item que não se aplica à empresa, prática desnecessária para o porte (ex.: MEI dispensado), necessidade já suprida de outra forma (terceirizada, contador, informal mas eficaz), ou observação dizendo que está resolvido/sem impacto. Se a observação descrever um problema real mesmo com resposta aparentemente positiva, marque true. Na dúvida, use o texto da observação para decidir.

DIRETRIZES SEMÂNTICAS OBRIGATÓRIAS DE GESTÃO E COMÉRCIO:
1. PERGUNTAS SOBRE RISCO, LIMITAÇÃO, DEPENDÊNCIA OU OCORRÊNCIA DE FALHAS:
   Exemplos: "O produtor comercializa apenas em feiras/festivais locais?", "Depende de poucos compradores?", "Já teve crédito rejeitado?", "Possui dívidas em atraso?", "Mistura finanças pessoais?", "Sofre perdas na produção?":
   - Resposta "Não": NÃO HÁ PROBLEMA (o produtor NÃO opera apenas em feiras, não possui dívidas, etc. Ele diversifica e tem maturidade). -> "temProblema": false. NÃO gerar ação corretiva!
   - Resposta "Sim": HÁ PROBLEMA REAL (dependência comercial, inadimplência, perdas, etc.). -> "temProblema": true.
   - Resposta "Parcial": HÁ GARGALO PARCIAL. -> "temProblema": true.

2. PERGUNTAS SOBRE CONTROLES, PLANEJAMENTO E BOAS PRÁTICAS:
   Exemplos: "Possui controle financeiro?", "Realiza precificação?", "Possui alvará sanitário?", "A maior parte da produção é vendida beneficiada?":
   - Resposta "Sim": NÃO HÁ PROBLEMA (aplica a boa prática). -> "temProblema": false. NÃO gerar ação corretiva!
   - Resposta "Não": HÁ PROBLEMA (falta de controle essencial). -> "temProblema": true.
   - Resposta "Parcial": HÁ GARGALO PARCIAL (oportunidade de ampliação/melhoria). -> "temProblema": true.

3. AÇÕES CORRETIVAS E OBSERVAÇÕES PRESCRITIVAS DO CONSULTOR:
   Se a observação do consultor prescrever ou orientar CORRIGIR, ELABORAR, IMPLANTAR, CAPACITAR, TREINAR, ESTRUTURAR, PADRONIZAR ou AJUSTAR (ex.: "Elaborar procedimentos padronizados de higiene e capacitar a equipe produtiva em Boas Práticas de Fabricação"):
   -> "temProblema": true (HÁ PROBLEMA / LACUNA QUE EXIGE AÇÃO CORRETIVA). NUNCA marque false quando o consultor prescrever ações de correção, melhoria ou capacitação!

4. PRECEITOS AMBIENTAIS, ESG E SUSTENTABILIDADE:
   Exemplos: "O produtor faz uso de queimadas no plantio ou na colheita?", "Há uso de lenha/madeira na produção?", "Gera poluição/descarte inadequado de efluentes ou embalagens?":
   - Resposta "Não": NÃO HÁ PROBLEMA (o produtor NÃO comete a prática degradadora, preservando o solo, os recursos florestais e o meio ambiente). -> "temProblema": false. NÃO gerar ação corretiva!
   - Resposta "Sim": HÁ PROBLEMA REAL (dano/risco ambiental que exige controle, cessação de queimadas ou transição energética). -> "temProblema": true.

5. DECISÕES ESTRATÉGICAS E PARCERIAS VOLUNTÁRIAS:
   Exemplos: "O produtor tem interesse em fazer parceria com a Secretaria de Meio Ambiente para receber madeira de poda?", "Deseja participar de projeto voluntário?":
   - Resposta "Não": NÃO HÁ PROBLEMA (decisão estratégica de foco no processo produtivo atual; o produtor não é obrigado a adotar insumos ou parcerias fora de sua demanda). -> "temProblema": false. NÃO gerar ação corretiva!
   - Resposta "Sim": OPORTUNIDADE DE PARCERIA / INOVAÇÃO. -> "temProblema": false se já atende ou true apenas se necessitar de estruturação acordada.

ITENS A AVALIAR:
${JSON.stringify(itens)}

Retorne APENAS um array JSON:
[
  { "id": "id_do_item", "temProblema": false, "motivo": "Não vende apenas em feiras; diversifica canais." }
]`;

  const TAMANHO_LOTE = 25;
  for (let i = 0; i < itensTodos.length; i += TAMANHO_LOTE) {
    const lote = itensTodos.slice(i, i + TAMANHO_LOTE);
    try {
      const resp = await ai.models.generateContent({
        model: 'gemini-3.6-flash',
        contents: montarPrompt(lote),
        config: {
          responseMimeType: 'application/json',
        }
      });

      const bruto = String(resp?.text || '').replace(/```json|```/g, '').trim();
      const parsed: any[] = JSON.parse(bruto.slice(bruto.indexOf('['), bruto.lastIndexOf(']') + 1));
      if (Array.isArray(parsed)) {
        for (const item of parsed) {
          if (item && item.id !== undefined && item.temProblema !== undefined) {
            const id = String(item.id);
            const tem = Boolean(item.temProblema);
            mapa.set(id, tem);
            const sig = sigPorId.get(id);
            if (sig) vereditosIA.set(id, { sig, temProblema: tem });
          }
        }
      }
    } catch (err) {
      console.warn('[planoAcao] Avaliação semântica por IA falhou em um lote; usando regras semânticas locais:', err);
    }
  }
  return mapa;
}

const ID_FERRAMENTAS = '__ferramentas__';

export async function enriquecerComIA(
  ai: GeradorIA | null | undefined,
  atividades: AtividadeCronograma[],
  analise: AnaliseDiagnostico,
  contexto: { tipoEmpresa?: string; nomeEmpresa?: string; tipoProdutoServico?: string } = {}
): Promise<{ atividades: AtividadeCronograma[]; usouIA: boolean; erro?: string }> {
  const tratadosNoPlano = new Set(atividades.flatMap((a) => (a.problemasTratados || []).map((t) => t.idProblema)));
  const alvo = analise.problemas.filter((p) => tratadosNoPlano.has(p.idProblema));
  if (!ai || alvo.length === 0) return { atividades, usouIA: false };

  const dados = alvo.map((p) => ({
    idProblema: p.idProblema,
    problema: p.problema,
    area: p.area,
    criticidade: p.nivel,
    respostas: p.lacunas.slice(0, 5),
    solucaoCadastrada: p.solucao
      ? { solucao: p.solucao.solucao_recomendada, acoes: p.solucao.acoes_sugeridas, kpis: p.solucao.kpis_sugeridos }
      : undefined,
  }));
  const ferr = atividades.find((a) => papelDaAtividade(a.nome) === 'ferramentas');

  const prompt = `Você é um consultor empresarial sênior (padrão SEBRAE). Com base no diagnóstico${
    contexto.nomeEmpresa ? ` da empresa "${contexto.nomeEmpresa}"` : ''
  }${contexto.tipoEmpresa ? ` (segmento: ${contexto.tipoEmpresa})` : ''}${
    contexto.tipoProdutoServico ? ` e considerando especificamente seu produto/serviço: "${contexto.tipoProdutoServico}"` : ''
  }, escreva a solução de CADA problema abaixo com total aplicabilidade prática ao seu produto/serviço e modelo operacional.

PARA CADA PROBLEMA (um item por idProblema, mantendo o mesmo idProblema):
- "acao": título curto e profissional da ação (até 80 caracteres), começando por um verbo no infinitivo.
- "solucao": a solução em uma frase.
- "passos": passo a passo prático de execução e ações recomendadas (2 a 4 passos diretos), SEM repetir a pergunta e a resposta do diagnóstico.
- "resultadoEsperado": resultado com pelo menos um indicador mensurável (número, % ou prazo).
- "responsavel": "Consultor", "Cliente" ou "Consultor/Cliente".
- Problemas CRÍTICOS exigem ações mais completas e prazos mais curtos. Se houver solução cadastrada, use-a como base.

INCLUA TAMBÉM um item com "idProblema": "${ID_FERRAMENTAS}" descrevendo, em "passos", as planilhas e/ou aplicativos de
gestão que serão desenvolvidos para apoiar a solução desses problemas (quais controles, o que cada um mede e quem vai usar).

REGRAS: português do Brasil, linguagem objetiva e técnica, sem inventar dados da empresa.

PROBLEMAS (do mais crítico ao menos crítico):
${JSON.stringify(dados)}
${ferr ? `\nFERRAMENTAS PREVISTAS (${horasDe(ferr)}h): ${ferr.descricao}` : ''}`;

  try {
    const resp = await ai.models.generateContent({
      model: 'gemini-3.6-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              idProblema: { type: Type.STRING },
              acao: { type: Type.STRING },
              solucao: { type: Type.STRING },
              passos: { type: Type.STRING },
              resultadoEsperado: { type: Type.STRING },
              responsavel: { type: Type.STRING },
            },
            required: ['idProblema', 'acao', 'solucao', 'passos', 'resultadoEsperado'],
          },
        },
      },
    });
    const bruto = String(resp?.text || '').replace(/```json|```/g, '').trim();
    const itens: any[] = JSON.parse(bruto.slice(bruto.indexOf('['), bruto.lastIndexOf(']') + 1));
    const porId = new Map(itens.filter((i) => i && i.idProblema).map((i) => [String(i.idProblema), i]));
    const txt = (v: any, atual: string, max = 600) => (typeof v === 'string' && v.trim() ? v.trim().slice(0, max) : atual);

    let aplicados = 0;
    const novas = atividades.map((a) => {
      if (papelDaAtividade(a.nome) === 'ferramentas') {
        const i = porId.get(ID_FERRAMENTAS);
        if (!i) return a;
        aplicados++;
        return { ...a, descricao: txt(i.passos, a.descricao), solucaoProposta: txt(i.solucao, a.solucaoProposta), resultadoEsperado: txt(i.resultadoEsperado, a.resultadoEsperado || '') };
      }
      if (!a.problemasTratados?.length) return a;
      const tratados = a.problemasTratados.map((t) => {
        const i = porId.get(t.idProblema);
        if (!i) return t;
        aplicados++;
        return {
          ...t,
          acao: txt(i.acao, t.acao, 120),
          solucao: txt(i.solucao, t.solucao, 300),
          passos: txt(i.passos, t.passos, 700),
          resultadoEsperado: txt(i.resultadoEsperado, t.resultadoEsperado, 400),
          responsavel: txt(i.responsavel, t.responsavel, 40),
        };
      });
      // O nome da atividade de área segue o padrão fixo; só os textos dos problemas mudam.
      const area = a.area || tratados[0]?.area || 'Área';
      return { ...a, ...textosDaAtividadeDeArea(area, tratados), nome: a.nome, problemasTratados: tratados };
    });
    return { atividades: novas, usouIA: aplicados > 0 };
  } catch (e: any) {
    return { atividades, usouIA: false, erro: e?.message || String(e) };
  }
}

// ---------------------------------------------------------------------------------------
// Tarefas do Plano de Ação: UMA POR PROBLEMA nas atividades de área
// ---------------------------------------------------------------------------------------
export interface TarefaGerada {
  problema: string;
  area?: string;
  idProblema: string;
  solucaoSugerida: string;
  acoes: string;
  prioridade: 'Alta' | 'Média' | 'Baixa';
  responsavel: string;
  cargaHoraria?: string;
  resultadoEsperado: string;
  dataInicio?: string;
  dataFim?: string;
  /** Atividade do cronograma a que a tarefa pertence. */
  atividade: string;
  atividadeOrdem: number;
  criticidade?: NivelCriticidade;
}

/** Converte as atividades do cronograma nas tarefas do Plano de Ação. */
export function tarefasDasAtividades(atividades: AtividadeCronograma[]): TarefaGerada[] {
  const out: TarefaGerada[] = [];
  atividades.forEach((atv, atividadeOrdem) => {
    const itens = atv.problemasTratados || [];
    if (!itens.length) {
      out.push({
        problema: atv.nome,
        area: atv.area,
        idProblema: atv.idProblema || '',
        solucaoSugerida: atv.solucaoProposta || '',
        acoes: atv.descricao || '',
        prioridade: (atv.prioridade as any) || 'Média',
        responsavel: atv.responsavel || 'Consultor',
        cargaHoraria: atv.cargaHoraria,
        resultadoEsperado: atv.resultadoEsperado || '',
        dataInicio: atv.dataInicio,
        dataFim: atv.dataFim,
        atividade: atv.nome,
        atividadeOrdem,
      });
      return;
    }
    // Horas da atividade repartidas entre os problemas pelo peso da criticidade, em blocos de
    // meia hora, somando EXATAMENTE as horas da atividade.
    const horas = horasDe(atv);
    const soma = itens.reduce((s, p) => s + PESO_NIVEL[p.nivel], 0) || 1;
    const blocos = horas * 2;
    const ideais = itens.map((p) => (blocos * PESO_NIVEL[p.nivel]) / soma);
    const partes = ideais.map(Math.floor);
    let faltam = blocos - partes.reduce((a, b) => a + b, 0);
    ideais
      .map((v, k) => [v - partes[k], k] as const)
      .sort((x, y) => y[0] - x[0])
      .forEach(([, k]) => { if (faltam > 0) { partes[k]++; faltam--; } });
    itens.forEach((p, k) => {
      const h = partes[k] / 2;
      out.push({
        problema: p.problema,
        area: p.area,
        idProblema: p.idProblema,
        solucaoSugerida: `${p.acao} — ${p.solucao}`,
        acoes: p.passos,
        prioridade: p.nivel === 'Moderado' ? 'Média' : 'Alta',
        responsavel: p.responsavel,
        cargaHoraria: horas ? `${String(h).replace('.', ',')}h` : undefined,
        resultadoEsperado: p.resultadoEsperado,
        dataInicio: atv.dataInicio,
        dataFim: atv.dataFim,
        atividade: atv.nome,
        atividadeOrdem,
        criticidade: p.nivel,
      });
    });
  });
  return out;
}

// ---------------------------------------------------------------------------------------
// Relatório
// ---------------------------------------------------------------------------------------
export function montarTextosRelatorio(
  analise: AnaliseDiagnostico,
  atividades: AtividadeCronograma[],
  contexto: { nomeEmpresa?: string; tipoEmpresa?: string; tipoProdutoServico?: string } = {}
): { objetivo: string; solucoesIndicadas: string; resultadosEsperados: string } {
  const criticos = analise.problemas.filter((p) => p.nivel === 'Crítico');
  const grupos = agruparPorArea(analise);
  const areas = grupos.map((g) => g.area);
  const empresa = contexto.nomeEmpresa ? ` da empresa ${contexto.nomeEmpresa}` : '';
  const prodServ = contexto.tipoProdutoServico ? ` no segmento de ${contexto.tipoEmpresa || 'atuação'}, com foco em ${contexto.tipoProdutoServico}` : '';
  const totalHoras = atividades.reduce((s, a) => s + horasDe(a), 0);
  const tratados = new Map(atividades.flatMap((a) => (a.problemasTratados || []).map((t) => [t.idProblema, t] as const)));

  const objetivo =
    analise.problemas.length === 0
      ? `Consolidar as boas práticas de gestão identificadas no diagnóstico${empresa}${prodServ} e estruturar controles e indicadores que garantam a continuidade dos resultados.`
      : `Solucionar os ${analise.problemas.length} problema${analise.problemas.length > 1 ? 's' : ''} identificado${analise.problemas.length > 1 ? 's' : ''} no diagnóstico${empresa}${prodServ}` +
        `${criticos.length ? ` (${criticos.length} crítico${criticos.length > 1 ? 's' : ''})` : ''}` +
        `${areas.length ? `, distribuídos em ${areas.length} área${areas.length > 1 ? 's' : ''} (${listaTexto(areas.slice(0, 6))})` : ''}, ` +
        `por meio de ${atividades.length} atividades (${totalHoras}h), com implantação de ferramentas de gestão e acompanhamento de indicadores. ` +
        `Índice de conformidade no diagnóstico: ${analise.scoreGeral}%.`;

  const porArea = grupos
    .map((g) => {
      const linhas = g.problemas.map((p) => {
        const t = tratados.get(p.idProblema);
        if (t?.acao) {
          return `• [${p.nivel}] ${p.problema}\n  Ação Proposta: ${t.acao}`;
        }
        return `• [${p.nivel}] ${p.problema}`;
      });
      return `${g.area.toUpperCase()} (${resumoNiveis(g.problemas)}):\n${linhas.join('\n\n')}`;
    })
    .join('\n\n');
  const acoesTxt = listarAcoesTexto(atividades);
  const fortesTxt = analise.pontosFortes.slice(0, 10).map((p) => `• ${p.problema} (${p.area})`).join('\n');

  const solucoesIndicadas = [
    `PROBLEMAS IDENTIFICADOS:\n${porArea || '• Nenhuma lacuna relevante identificada.'}`,
    `SOLUÇÕES / AÇÕES PROPOSTAS:\n${acoesTxt}`,
    fortesTxt ? `PONTOS FORTES:\n${fortesTxt}` : '',
  ]
    .filter(Boolean)
    .join('\n\n');

  return { objetivo, solucoesIndicadas, resultadosEsperados: listarResultadosTexto(atividades) };
}

// ---------------------------------------------------------------------------------------
// Sincronia Plano de Ação ↔ Relatório
// ---------------------------------------------------------------------------------------
const TITULO_PROBLEMAS = 'PROBLEMAS IDENTIFICADOS:';
const TITULO_ACOES = 'SOLUÇÕES / AÇÕES PROPOSTAS:';
const TITULO_FORTES = 'PONTOS FORTES:';

/** Lista numerada das atividades (seção "SOLUÇÕES / AÇÕES PROPOSTAS" do relatório). */
export function listarAcoesTexto(atividades: AtividadeCronograma[]): string {
  return atividades
    .map((a, i) => `${i + 1}. ${a.nome} (${a.cargaHoraria})${!a.problemasTratados?.length && a.solucaoProposta && a.solucaoProposta !== a.nome ? ` - ${a.solucaoProposta}` : ''}`)
    .join('\n');
}

/** Resultados esperados do relatório a partir das atividades / problemas tratados. */
export function listarResultadosTexto(atividades: AtividadeCronograma[]): string {
  const resultados: string[] = [];
  for (const a of atividades) {
    if (a.problemasTratados?.length) a.problemasTratados.forEach((t) => resultados.push(`• [${t.nivel}] ${t.problema}: ${t.resultadoEsperado}`));
    else if (a.resultadoEsperado?.trim()) resultados.push(`• ${a.resultadoEsperado.trim().startsWith('•') ? a.resultadoEsperado.trim().slice(1).trim() : a.resultadoEsperado.trim()}`);
  }
  return resultados.join('\n');
}

type DadosTexto = { objetivo?: string; solucoesIndicadas?: string; resultadosEsperados?: string; cargaHoraria?: string };

/**
 * Atualiza os textos do relatório (ações propostas, resultados esperados, carga horária e a contagem
 * de atividades do objetivo) para refletir o cronograma/plano atual. Preserva as seções
 * "PROBLEMAS IDENTIFICADOS" e "PONTOS FORTES" já existentes.
 */
export function sincronizarRelatorioComCronograma<T extends DadosTexto>(dados: T | undefined, atividades: AtividadeCronograma[]): T {
  const base = (dados || {}) as T;
  const validas = (atividades || []).filter((a) => a?.nome?.trim());
  if (validas.length === 0) return base;
  const totalHoras = validas.reduce((s, a) => s + horasDe(a), 0);

  const atual = base.solucoesIndicadas || '';
  const iProb = atual.indexOf(TITULO_PROBLEMAS);
  const iAcoes = atual.indexOf(TITULO_ACOES);
  const iFortes = atual.indexOf(TITULO_FORTES);
  const blocoProblemas = iProb >= 0 ? atual.slice(iProb, iAcoes > iProb ? iAcoes : (iFortes > iProb ? iFortes : atual.length)).trim() : '';
  const blocoFortes = iFortes >= 0 ? atual.slice(iFortes).trim() : '';
  const solucoesIndicadas = [blocoProblemas, `${TITULO_ACOES}\n${listarAcoesTexto(validas)}`, blocoFortes].filter(Boolean).join('\n\n');

  const resultados = listarResultadosTexto(validas);
  const objetivo = (base.objetivo || '').replace(
    /por meio de \d+ atividades? \(\d+(?:[.,]\d+)?h\)/,
    `por meio de ${validas.length} atividade${validas.length > 1 ? 's' : ''} (${String(totalHoras).replace('.', ',')}h)`
  );

  return {
    ...base,
    objetivo: objetivo || base.objetivo,
    solucoesIndicadas,
    resultadosEsperados: resultados || base.resultadosEsperados,
    cargaHoraria: totalHoras > 0 ? `${String(totalHoras).replace('.', ',')}h` : base.cargaHoraria,
  } as T;
}

/** Campos da tarefa do Plano de Ação usados na sincronia (compatível com TarefaPlanoAcao). */
export interface TarefaSync {
  idProblema?: string;
  problema?: string;
  area?: string;
  solucaoSugerida?: string;
  acoes?: string;
  prioridade?: string;
  responsavel?: string;
  cargaHoraria?: string;
  resultadoEsperado?: string;
  dataInicio?: any;
  dataFim?: any;
  atividade?: string;
  atividadeOrdem?: number;
  criticidade?: string;
}

const dataStr = (d: any): string | undefined => {
  if (!d) return undefined;
  if (typeof d === 'string') return d;
  if (d instanceof Date) return d.toISOString().split('T')[0];
  if (typeof d?.toDate === 'function') return d.toDate().toISOString().split('T')[0];
  return undefined;
};

/** Índice (no array original) da atividade a que a tarefa pertence; -1 se não achar. */
function indiceDaAtividade(atividades: AtividadeCronograma[], t: TarefaSync): number {
  const idxValidos = atividades.map((a, i) => (a?.nome?.trim() ? i : -1)).filter((i) => i >= 0);
  if (typeof t.atividadeOrdem === 'number' && idxValidos[t.atividadeOrdem] !== undefined) {
    const i = idxValidos[t.atividadeOrdem];
    if (!t.atividade || norm(atividades[i].nome) === norm(t.atividade)) return i;
  }
  if (t.atividade) {
    const i = atividades.findIndex((a) => norm(a?.nome) === norm(t.atividade));
    if (i >= 0) return i;
  }
  if (t.idProblema) {
    const i = atividades.findIndex((a) => a.idProblema === t.idProblema || a.problemasTratados?.some((p) => p.idProblema === t.idProblema));
    if (i >= 0) return i;
  }
  if (t.problema) {
    return atividades.findIndex((a) => norm(a?.nome) === norm(t.problema) || a.problemasTratados?.some((p) => norm(p.problema) === norm(t.problema)));
  }
  return -1;
}

function indiceDoProblema(a: AtividadeCronograma, t: TarefaSync, original?: TarefaSync): number {
  const lista = a.problemasTratados || [];
  const alvoId = original?.idProblema || t.idProblema;
  const alvoNome = norm(original?.problema || t.problema);
  let i = alvoId ? lista.findIndex((p) => p.idProblema === alvoId) : -1;
  if (i < 0) i = lista.findIndex((p) => norm(p.problema) === alvoNome);
  return i;
}

/** Aplica no cronograma a edição de uma tarefa do Plano de Ação. `original` = tarefa antes da edição. */
export function aplicarTarefaNoCronograma(atividades: AtividadeCronograma[], tarefa: TarefaSync, original?: TarefaSync): AtividadeCronograma[] {
  const lista = atividades.map((a) => ({ ...a }));
  const i = indiceDaAtividade(lista, original || tarefa);
  if (i < 0) return lista;
  const atv = lista[i];
  const prio = tarefa.prioridade === 'Alta' || tarefa.prioridade === 'Baixa' || tarefa.prioridade === 'Média' ? tarefa.prioridade : atv.prioridade;
  const ini = dataStr(tarefa.dataInicio);
  const fim = dataStr(tarefa.dataFim);

  if (atv.problemasTratados?.length) {
    const k = indiceDoProblema(atv, tarefa, original);
    if (k >= 0) {
      const p = { ...atv.problemasTratados[k] };
      if (tarefa.problema) p.problema = tarefa.problema;
      if (tarefa.solucaoSugerida !== undefined) {
        const sep = tarefa.solucaoSugerida.indexOf(' — ');
        if (sep >= 0) { p.acao = tarefa.solucaoSugerida.slice(0, sep); p.solucao = tarefa.solucaoSugerida.slice(sep + 3); }
        else p.solucao = tarefa.solucaoSugerida;
      }
      if (tarefa.acoes !== undefined) p.passos = tarefa.acoes;
      if (tarefa.resultadoEsperado !== undefined) p.resultadoEsperado = tarefa.resultadoEsperado;
      if (tarefa.responsavel) p.responsavel = tarefa.responsavel;
      atv.problemasTratados = atv.problemasTratados.map((x, n) => (n === k ? p : x));
    }
  } else {
    if (tarefa.problema) atv.nome = tarefa.problema;
    if (tarefa.solucaoSugerida !== undefined) atv.solucaoProposta = tarefa.solucaoSugerida;
    if (tarefa.acoes !== undefined) atv.descricao = tarefa.acoes;
    if (tarefa.resultadoEsperado !== undefined) atv.resultadoEsperado = tarefa.resultadoEsperado;
    if (tarefa.responsavel) atv.responsavel = tarefa.responsavel;
    if (tarefa.cargaHoraria) atv.cargaHoraria = tarefa.cargaHoraria;
  }
  atv.prioridade = prio as AtividadeCronograma['prioridade'];
  if (ini) atv.dataInicio = ini;
  if (fim) atv.dataFim = fim;
  return lista;
}

/** Remove do cronograma a atividade (ou o problema tratado) correspondente à tarefa excluída. */
export function removerTarefaDoCronograma(atividades: AtividadeCronograma[], tarefa: TarefaSync): AtividadeCronograma[] {
  const lista = atividades.map((a) => ({ ...a }));
  const i = indiceDaAtividade(lista, tarefa);
  if (i < 0) return lista;
  const atv = lista[i];
  if (atv.problemasTratados && atv.problemasTratados.length > 1) {
    const k = indiceDoProblema(atv, tarefa);
    if (k >= 0) {
      atv.problemasTratados = atv.problemasTratados.filter((_, n) => n !== k);
      return lista;
    }
  }
  return lista.filter((_, n) => n !== i);
}

/** Inclui no cronograma a tarefa criada no Plano de Ação (nova atividade antes do encerramento). */
export function adicionarTarefaAoCronograma(atividades: AtividadeCronograma[], tarefa: TarefaSync): AtividadeCronograma[] {
  const lista = atividades.map((a) => ({ ...a }));
  if (indiceDaAtividade(lista, tarefa) >= 0) return lista; // já representada
  const nova: AtividadeCronograma = {
    nome: tarefa.problema || 'Nova atividade',
    descricao: tarefa.acoes || '',
    cargaHoraria: tarefa.cargaHoraria || '2h',
    solucaoProposta: tarefa.solucaoSugerida || tarefa.problema || '',
    resultadoEsperado: tarefa.resultadoEsperado || '',
    responsavel: tarefa.responsavel || 'Consultor',
    idProblema: tarefa.idProblema || undefined,
    status: 'Pendente',
    prioridade: (tarefa.prioridade as AtividadeCronograma['prioridade']) || 'Média',
    dataInicio: dataStr(tarefa.dataInicio),
    dataFim: dataStr(tarefa.dataFim),
    area: tarefa.area,
  };
  const fim = lista.findIndex((a) => { const p = papelDaAtividade(a.nome); return p === 'analise' || p === 'relatorio'; });
  if (fim >= 0) lista.splice(fim, 0, nova); else lista.push(nova);
  return lista;
}
