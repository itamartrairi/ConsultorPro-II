import os

os.makedirs('src/constants', exist_ok=True)

with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Extract consultoriaConstants.ts
with open('src/constants/consultoriaConstants.ts', 'w', encoding='utf-8') as out:
    out.write('// Constantes e Modelos de Consultoria Estratégica\n\n')
    
    out.write('export const CONSULTORIA_AREAS = [\n')
    for l in lines[190:203]:
        out.write(l)
    out.write('];\n\n')
    
    out.write('export const AREAS = [\n')
    for l in lines[840:853]:
        out.write(l)
    out.write('];\n\n')
    
    out.write('export const IMPACTO_ORDER: Record<string, number> = { \'Baixo\': 1, \'Médio\': 2, \'Alto\': 3 };\n\n')
    
    out.write('''export interface AtividadeCronograma {
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

''')
    out.write('export const PLANO_DE_ACAO_PADRAO: AtividadeCronograma[] = [\n')
    for l in lines[1136:1226]:
        out.write(l)
    out.write('\n\nexport const MODELOS_RELATORIO: ModeloRelatorio[] = [\n')
    for l in lines[1228:1852]:
        out.write(l)
    out.write('\n\nexport const TIPOS_EMPRESA = [\n')
    out.write('  "Geral",\n  "Comércio",\n  "Serviços",\n  "Indústria",\n  "Agronegócio",\n  "Carcinicultura",\n  "Tecnologia / Startup",\n  "Alimentação & Gastronomia"\n];\n\n')
    out.write('export const AREAS_ORDER = [\n')
    out.write('  "Estratégia",\n  "Processos",\n  "Marketing & Vendas",\n  "Financeiro",\n  "Recursos Humanos",\n  "Tecnologia & IA",\n  "Inovação",\n  "Gestão da Produção",\n  "Qualidade",\n  "Sustentabilidade & ESG",\n  "Jurídico & Compliance",\n  "Acesso a Crédito",\n  "Geral"\n];\n')

print('src/constants/consultoriaConstants.ts generated successfully!')
