// Utilitário de Temas e Cores Corporativas para o Sistema e Relatórios PDF

export interface ColorPreset {
  id: string;
  name: string;
  hex: string;
  rgb: [number, number, number];
  description: string;
}

export const COLOR_PRESETS: ColorPreset[] = [
  {
    id: 'emerald',
    name: 'Verde SEBRAE (Padrão)',
    hex: '#059669',
    rgb: [5, 150, 105],
    description: 'Verde esmeralda oficial SEBRAE / Consultoria'
  },
  {
    id: 'royal-blue',
    name: 'Azul Corporativo',
    hex: '#0284c7',
    rgb: [2, 132, 199],
    description: 'Azul executivo moderno e confiável'
  },
  {
    id: 'indigo',
    name: 'Índigo Estratégico',
    hex: '#4f46e5',
    rgb: [79, 70, 229],
    description: 'Tom sofisticado e tecnológico'
  },
  {
    id: 'teal',
    name: 'Teal Oceano',
    hex: '#0d9488',
    rgb: [13, 148, 136],
    description: 'Elegância equilibrada para finanças e gestão'
  },
  {
    id: 'ruby',
    name: 'Ruby & Vinho',
    hex: '#e11d48',
    rgb: [225, 29, 72],
    description: 'Alto impacto para diagnósticos e auditorias'
  },
  {
    id: 'amber',
    name: 'Âmbar & Ouro',
    hex: '#d97706',
    rgb: [217, 119, 6],
    description: 'Tom premium de alta energia e liderança'
  },
  {
    id: 'slate',
    name: 'Grafite Executivo',
    hex: '#334155',
    rgb: [51, 65, 85],
    description: 'Minimalismo clássico e sóbrio'
  }
];

export const DEFAULT_CORPORATE_COLOR = '#059669';

export function hexToRgb(hex: string): [number, number, number] {
  let clean = hex.replace('#', '').trim();
  if (clean.length === 3) {
    clean = clean.split('').map(c => c + c).join('');
  }
  const num = parseInt(clean, 16);
  if (isNaN(num)) return [5, 150, 105];
  const r = (num >> 16) & 255;
  const g = (num >> 8) & 255;
  const b = num & 255;
  return [r, g, b];
}

export function getCorporateColor(): string {
  if (typeof window === 'undefined') return DEFAULT_CORPORATE_COLOR;
  return localStorage.getItem('corporate_primary_color') || DEFAULT_CORPORATE_COLOR;
}

export function getCorporateRgb(): [number, number, number] {
  const hex = getCorporateColor();
  return hexToRgb(hex);
}

export function setCorporateColor(hex: string): void {
  if (typeof window === 'undefined') return;
  const validHex = hex && hex.startsWith('#') ? hex : DEFAULT_CORPORATE_COLOR;
  localStorage.setItem('corporate_primary_color', validHex);
  applyCorporateTheme(validHex);
}

export function applyCorporateTheme(customHex?: string): void {
  if (typeof document === 'undefined') return;
  const hex = customHex || getCorporateColor();
  const [r, g, b] = hexToRgb(hex);

  document.documentElement.style.setProperty('--corporate-primary', hex);
  document.documentElement.style.setProperty('--corporate-primary-rgb', `${r}, ${g}, ${b}`);
  document.documentElement.style.setProperty('--corporate-primary-light', `rgba(${r}, ${g}, ${b}, 0.12)`);
  document.documentElement.style.setProperty('--corporate-primary-dark', `rgb(${Math.max(0, r - 30)}, ${Math.max(0, g - 30)}, ${Math.max(0, b - 30)})`);
}
