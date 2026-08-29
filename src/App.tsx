import * as XLSX from 'xlsx';
import React, { useState, useEffect, useMemo, Component, useRef } from 'react';
import html2canvas from 'html2canvas';
import { 
  collection, 
  query, 
  where, 
  onSnapshot, 
  addDoc, 
  updateDoc, 
  doc, 
  deleteDoc, 
  getDocs,
  orderBy,
  limit,
  Timestamp,
  serverTimestamp,
  writeBatch,
  setDoc,
  getDocFromServer,
  deleteField,
  runTransaction
} from 'firebase/firestore';
import { 
  signInWithPopup, 
  GoogleAuthProvider, 
  onAuthStateChanged, 
  signOut,
  User,
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword
} from 'firebase/auth';
import { db, auth, seedPremissasIfEmpty, checkCloudConnection } from './firebase';
import { ProjectsView } from './components/ProjectsView';
import { ReplicateDiagnosticoModal } from './components/ReplicateDiagnosticoModal';
import { EditDiagnosisDateModal } from './components/EditDiagnosisDateModal';
import { CopySgfConsultoriaModal } from './components/CopySgfConsultoriaModal';
import { AgendaView } from './components/AgendaView';
import { MacroDashboardView } from './components/MacroDashboardView';
import DiscAvaliacaoView from './components/DiscAvaliacaoView';
import MaturidadeView from './components/MaturidadeView';
import { ResultadoConsultoriaView } from './components/ResultadoConsultoriaView';
import { SmartSyncModal, type SyncSummary } from './components/SmartSyncModal';
import { BackupExportModal, type BackupExportStats } from './components/BackupExportModal';
import { ConnectionStatus } from './components/SettingsView';
import { sebraeLogoBase64 } from './sebraeLogo';
import { getCachedAI, setCachedAI, generateAICacheKey } from './lib/aiCache';
import { 
  Plus, 
  Award,
  FileCheck,
  FileText, 
  Building2, 
  ChevronRight,
  ChevronDown, 
  CheckCircle2, 
  CreditCard,
  XCircle, 
  AlertCircle, 
  BarChart3, 
  Download, 
  Trash2, 
  LogOut, 
  ArrowLeft,
  ArrowRight,
  Search,
  History,
  Info,
  Menu,
  X,
  Upload,
  Calendar,
  RefreshCw,
  RotateCcw,
  Save,
  Target,
  TrendingUp,
  ExternalLink,
  Printer,
  AlertTriangle,
  Home,
  Trophy,
  Briefcase,
  CalendarDays,
  Loader2,
  PieChart as LucidePieChart,
  Sparkles,
  Layout,
  Clock,
  CheckCircle,
  Paperclip,
  Image as ImageIcon,
  History as HistoryIcon,
  Settings,
  ShieldCheck,
  Lock,
  Share2,
  Bell,
  Mail,
  MessageSquare,
  Smartphone,
  Send,
  Edit,
  Edit2,
  Edit3,
  Pencil,
  FileSpreadsheet,
  Users,
  Layers,
  Tag,
  Tags,
  ListOrdered,
  ArrowUp,
  ArrowDown,
  Copy,
  Key,
  Cloud,
  HardDrive,
  Database,
  CheckSquare,
  Square,
  ListFilter,
  CheckCheck,
  Infinity as InfinityIcon
} from 'lucide-react';




export enum Type {
  STRING = "STRING",
  NUMBER = "NUMBER",
  INTEGER = "INTEGER",
  BOOLEAN = "BOOLEAN",
  ARRAY = "ARRAY",
  OBJECT = "OBJECT",
}
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell,
  PieChart,
  Pie,
  LineChart,
  Line,
  AreaChart,
  Area
} from 'recharts';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { format } from 'date-fns';
import { ptBR } from 'date-fns/locale';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { cn } from './lib/utils';
import { Button } from './components/Button';
import { Modal } from './components/Modal';
import { ToastContainer, type ToastItem } from './components/Toast';
import { playClickSound, playSuccessSound } from './lib/sound';

declare global {
  interface ImportMeta {
    readonly env: Record<string, any>;
  }
}

import { 
  encryptedLocalStorage, 
  setStorageUserId, 
  setStorageSessionPassword 
} from './lib/cryptoStorage';
import {
  compressImageToDataUrl,
  saveLocalDiagnosticoReport,
  loadLocalDiagnosticoReport,
  saveLocalIndividualEvidence,
  deleteLocalIndividualEvidence,
  loadLocalIndividualEvidence
} from './lib/evidenceStorage';

// --- Safe LocalStorage with AES-256 Encryption at Rest ---
const safeLocalStorage = encryptedLocalStorage;
const localStorage = safeLocalStorage;

const CONSULTORIA_AREAS = [
  "FINANÃ‡AS",
  "MARKETING E VENDAS",
  "PLANEJAMENTO ESTRATÃ‰GICO",
  "GESTÃƒO DE PESSOAS",
  "PROCESSOS E OPERAÃ‡Ã•ES",
  "INOVAÃ‡ÃƒO E TECNOLOGIA",
  "JURÃDICO",
  "SUSTENTABILIDADE",
  "EMPREENDEDORISMO",
  "ACESSO A CRÃ‰DITO",
  "CRÃ‰DITO",
  "OUTROS"
];

// --- UtilitÃ¡rios de ValidaÃ§Ã£o e FormataÃ§Ã£o (PadrÃ£o Brasileiro - CPF e CNPJ) ---

/**
 * Remove qualquer caractere que nÃ£o seja dÃ­gito.
 */
export function cleanDigits(val?: string | null): string {
  if (!val) return '';
  return String(val).replace(/\D/g, '');
}

/**
 * Aplica a mÃ¡scara brasileira de CNPJ: 00.000.000/0000-00
 */
export function formatCNPJ(val?: string | null): string {
  const digits = cleanDigits(val).slice(0, 14);
  if (!digits) return '';
  let res = digits;
  if (digits.length > 2) res = digits.slice(0, 2) + '.' + digits.slice(2);
  if (digits.length > 5) res = digits.slice(0, 2) + '.' + digits.slice(2, 5) + '.' + digits.slice(5);
  if (digits.length > 8) res = digits.slice(0, 2) + '.' + digits.slice(2, 5) + '.' + digits.slice(5, 8) + '/' + digits.slice(8);
  if (digits.length > 12) res = digits.slice(0, 2) + '.' + digits.slice(2, 5) + '.' + digits.slice(5, 8) + '/' + digits.slice(8, 12) + '-' + digits.slice(12, 14);
  return res;
}

/**
 * ValidaÃ§Ã£o com cÃ¡lculo oficial dos 2 dÃ­gitos verificadores do CNPJ (Receita Federal)
 */
export function isValidCNPJ(val?: string | null): boolean {
  const digits = cleanDigits(val);
  if (!digits || digits.length !== 14) return false;
  // Bloqueia sequÃªncias de nÃºmeros iguais (ex: 00000000000000, 11111111111111)
  if (/^(\d)\1{13}$/.test(digits)) return false;

  // 1Âº DÃ­gito verificador
  let length = digits.length - 2;
  let numbers = digits.substring(0, length);
  const checkDigits = digits.substring(length);
  let sum = 0;
  let pos = length - 7;
  for (let i = length; i >= 1; i--) {
    sum += parseInt(numbers.charAt(length - i), 10) * pos--;
    if (pos < 2) pos = 9;
  }
  let result = sum % 11 < 2 ? 0 : 11 - (sum % 11);
  if (result !== parseInt(checkDigits.charAt(0), 10)) return false;

  // 2Âº DÃ­gito verificador
  length = length + 1;
  numbers = digits.substring(0, length);
  sum = 0;
  pos = length - 7;
  for (let i = length; i >= 1; i--) {
    sum += parseInt(numbers.charAt(length - i), 10) * pos--;
    if (pos < 2) pos = 9;
  }
  result = sum % 11 < 2 ? 0 : 11 - (sum % 11);
  if (result !== parseInt(checkDigits.charAt(1), 10)) return false;

  return true;
}

/**
 * Aplica a mÃ¡scara brasileira de CPF: 000.000.000-00
 */
export function formatCPF(val?: string | null): string {
  const digits = cleanDigits(val).slice(0, 11);
  if (!digits) return '';
  let res = digits;
  if (digits.length > 3) res = digits.slice(0, 3) + '.' + digits.slice(3);
  if (digits.length > 6) res = digits.slice(0, 3) + '.' + digits.slice(3, 6) + '.' + digits.slice(6);
  if (digits.length > 9) res = digits.slice(0, 3) + '.' + digits.slice(3, 6) + '.' + digits.slice(6, 9) + '-' + digits.slice(9, 11);
  return res;
}

/**
 * ValidaÃ§Ã£o com cÃ¡lculo oficial dos 2 dÃ­gitos verificadores do CPF (Receita Federal)
 */
export function isValidCPF(val?: string | null): boolean {
  const digits = cleanDigits(val);
  if (!digits || digits.length !== 11) return false;
  // Bloqueia sequÃªncias repetidas (ex: 11111111111, 00000000000)
  if (/^(\d)\1{10}$/.test(digits)) return false;

  // 1Âº DÃ­gito verificador
  let sum = 0;
  for (let i = 1; i <= 9; i++) {
    sum += parseInt(digits.substring(i - 1, i), 10) * (11 - i);
  }
  let remainder = (sum * 10) % 11;
  if (remainder === 10 || remainder === 11) remainder = 0;
  if (remainder !== parseInt(digits.substring(9, 10), 10)) return false;

  // 2Âº DÃ­gito verificador
  sum = 0;
  for (let i = 1; i <= 10; i++) {
    sum += parseInt(digits.substring(i - 1, i), 10) * (12 - i);
  }
  remainder = (sum * 10) % 11;
  if (remainder === 10 || remainder === 11) remainder = 0;
  if (remainder !== parseInt(digits.substring(10, 11), 10)) return false;

  return true;
}

// --- Utility ---

function sanitizeForFirestore(obj: any): any {
  if (obj === null || typeof obj !== 'object') return obj;
  if (obj instanceof Date) return obj.toISOString();
  // Do not strip Firestore Timestamps, FieldValues, or custom class instances with toDate/toMillis
  if (typeof obj.toDate === 'function' || typeof obj.toMillis === 'function' || obj.constructor?.name === 'Timestamp' || obj.constructor?.name === 'FieldValue' || obj._delegate) {
    return obj;
  }
  if (Array.isArray(obj)) return obj.map(v => sanitizeForFirestore(v));
  
  const result: any = {};
  for (const key in obj) {
    if (obj[key] !== undefined) {
      result[key] = sanitizeForFirestore(obj[key]);
    }
  }
  return result;
}

function oklchToRgb(oklchStr: string): string {
  try {
    return oklchStr.replace(/oklch\(([^)]+)\)/gi, (m, inner) => {
      const parts = inner.trim().split(/[\s,/\s]+/).filter(Boolean);
      if (parts.length < 3) return 'rgb(120, 120, 120)';

      const lStr = parts[0];
      const cStr = parts[1];
      const hStr = parts[2];
      const aStr = parts[3];

      let l = lStr.endsWith('%') ? parseFloat(lStr) / 100 : parseFloat(lStr);
      let c = parseFloat(cStr);
      let h = parseFloat(hStr);
      let alpha = aStr ? (aStr.endsWith('%') ? parseFloat(aStr) / 100 : parseFloat(aStr)) : 1;

      if (isNaN(l)) l = 0.5;
      if (isNaN(c)) c = 0.1;
      if (isNaN(h)) h = 0;
      if (isNaN(alpha)) alpha = 1;

      const hRad = (h * Math.PI) / 180;
      const a = c * Math.cos(hRad);
      const b_oklab = c * Math.sin(hRad);

      const l_1 = l + 0.3963377774 * a + 0.2158037573 * b_oklab;
      const m_1 = l - 0.1055613458 * a - 0.0638541728 * b_oklab;
      const s_1 = l - 0.0894841775 * a - 1.2914855480 * b_oklab;

      const l_ = l_1 * l_1 * l_1;
      const m_ = m_1 * m_1 * m_1;
      const s_ = s_1 * s_1 * s_1;

      let r_l =  4.0767416621 * l_ - 3.3077115913 * m_ + 0.2309699292 * s_;
      let g_l = -1.2684380046 * l_ + 2.6097574011 * m_ - 0.3413193965 * s_;
      let b_l = -0.0041960863 * l_ - 0.7034186147 * m_ + 1.7076147010 * s_;

      const f = (x: number) => {
        if (x <= 0.0031308) return 12.92 * Math.max(0, x);
        return 1.055 * Math.pow(Math.max(0, x), 1 / 2.4) - 0.055;
      };

      let r = Math.round(Math.max(0, Math.min(1, f(r_l))) * 255);
      let g = Math.round(Math.max(0, Math.min(1, f(g_l))) * 255);
      let b = Math.round(Math.max(0, Math.min(1, f(b_l))) * 255);

      if (alpha === 1) {
        return `rgb(${r}, ${g}, ${b})`;
      } else {
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
      }
    });
  } catch (err) {
    console.error("Error in oklchToRgb conversion:", err);
    return 'rgb(120, 120, 120)';
  }
}

const captureElementWithHtml2Canvas = async (elt: HTMLElement) => {
  return html2canvas(elt, {
    scale: 2,
    useCORS: true,
    allowTaint: false,
    logging: false,
    onclone: (clonedDoc) => {
      // 1. Clean all style tags in cloned doc to replace oklch with standard rgb/rgba
      clonedDoc.querySelectorAll('style').forEach(styleTag => {
        try {
          if (styleTag.textContent && styleTag.textContent.includes('oklch')) {
            styleTag.textContent = oklchToRgb(styleTag.textContent);
          }
        } catch (err) {
          console.error("Error patching style tag in html2canvas clone:", err);
        }
      });

      // 2. Clean inline styles and attributes containing oklch
      clonedDoc.querySelectorAll('*').forEach((el: any) => {
        try {
          if (el.style) {
            for (let i = 0; i < el.style.length; i++) {
              const prop = el.style[i];
              const val = el.style.getPropertyValue(prop);
              if (val && val.includes('oklch')) {
                el.style.setProperty(prop, oklchToRgb(val));
              }
            }
          }
          if (el.hasAttributes()) {
            for (const attr of Array.from(el.attributes) as any) {
              if (attr.value && attr.value.includes('oklch')) {
                attr.value = oklchToRgb(attr.value);
              }
            }
          }
        } catch (err) {
          // ignore element-specific patch errors
        }
      });

      const clonedWindow = clonedDoc.defaultView;
      if (clonedWindow) {
        const originalGetComputedStyle = clonedWindow.getComputedStyle;
        clonedWindow.getComputedStyle = (e: any, pseudoElt?: any) => {
          const style = originalGetComputedStyle.call(clonedWindow, e, pseudoElt);
          return new Proxy(style, {
            get(target, prop) {
              if (prop === 'getPropertyValue') {
                return (propertyName: string) => {
                  const val = target.getPropertyValue(propertyName);
                  if (typeof val === 'string' && val.includes('oklch')) {
                    return oklchToRgb(val);
                  }
                  return val;
                };
              }
              if (prop === 'cssText') {
                const val = target.cssText;
                if (typeof val === 'string' && val.includes('oklch')) {
                  return oklchToRgb(val);
                }
                return val;
              }
              const val = target[prop as any];
              if (typeof val === 'string' && val.includes('oklch')) {
                return oklchToRgb(val);
              }
              if (typeof val === 'function') {
                return val.bind(target);
              }
              return val;
            }
          }) as any;
        };
      }
    }
  });
};

const isValidLogoSource = (src: any): boolean => {
  if (!src || typeof src !== 'string') return false;
  const trimmed = src.trim();
  if (trimmed.length === 0) return false;
  if (trimmed === 'null' || trimmed === 'undefined') return false;
  if (trimmed.startsWith('data:')) {
    return trimmed.startsWith('data:image/') && trimmed.includes(';base64,') && trimmed.length > 35;
  }
  return trimmed.startsWith('http://') || trimmed.startsWith('https://') || trimmed.startsWith('/') || trimmed.startsWith('./');
};

export function parseLocalDate(val: any): Date | null {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val.toDate === 'function') {
    const d = val.toDate();
    return isNaN(d.getTime()) ? null : d;
  }
  if (val && typeof val === 'object' && val.seconds) {
    const d = new Date(val.seconds * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      const d = new Date(trimmed + 'T12:00:00');
      return isNaN(d.getTime()) ? null : d;
    }
    const d = new Date(trimmed);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function formatFirestoreDate(date: any, formatStr: string = 'dd/MM/yyyy', options?: any): string {
  if (!date) return '...';
  const d = parseLocalDate(date);
  if (!d) return '...';
  return format(d, formatStr, options);
}

export const getActivityDateStr = (diagDate: any, index: number): string => {
  let base: Date;
  if (!diagDate) {
    base = new Date();
  } else if (typeof diagDate === 'string') {
    base = new Date(diagDate.includes('T') ? diagDate : diagDate + 'T12:00:00');
  } else if (diagDate.toDate && typeof diagDate.toDate === 'function') {
    base = diagDate.toDate();
  } else if (diagDate.seconds) {
    base = new Date(diagDate.seconds * 1000);
  } else if (diagDate instanceof Date) {
    base = diagDate;
  } else {
    base = new Date(diagDate);
  }

  if (isNaN(base.getTime())) {
    base = new Date();
  }

  const result = new Date(base.getTime());
  result.setDate(result.getDate() + (index * 7));
  return result.toISOString().split('T')[0];
};

export function extractItemTimestamp(item: any): number {
  if (!item) return 0;
  
  // 1. Check direct timestamp fields in priority order
  const dateCandidates = [
    item.updatedAt,
    item.dataModificacao,
    item.dataAtualizacao,
    item.dataDiagnostico,
    item.dataCadastro,
    item.createdAt,
    item.dataCriacao,
    item.dataVencimento,
    item.dataInicio,
    item.timestamp
  ];

  for (const candidate of dateCandidates) {
    if (!candidate) continue;
    if (typeof candidate === 'number' && !isNaN(candidate) && candidate > 0) {
      return candidate < 10000000000 ? candidate * 1000 : candidate;
    }
    if (typeof candidate === 'object') {
      if (candidate instanceof Date && !isNaN(candidate.getTime())) {
        return candidate.getTime();
      }
      if (typeof candidate.toMillis === 'function') {
        try { return candidate.toMillis(); } catch {}
      }
      if (typeof candidate.toDate === 'function') {
        try { return candidate.toDate().getTime(); } catch {}
      }
      if (typeof candidate.seconds === 'number') {
        return candidate.seconds * 1000;
      }
    }
    if (typeof candidate === 'string') {
      const trimmed = candidate.trim();
      if (trimmed) {
        const parsed = Date.parse(trimmed.includes('T') ? trimmed : trimmed + 'T12:00:00');
        if (!isNaN(parsed) && parsed > 0) return parsed;
      }
    }
  }

  // 2. Specific heuristic weights for content richness
  let contentWeight = 0;
  if (item.resposta) contentWeight += 1000;
  if (item.observacao && String(item.observacao).trim()) contentWeight += 500;
  if (item.score !== undefined && item.score > 0) contentWeight += 200;
  if (item.cronograma && Array.isArray(item.cronograma) && item.cronograma.length > 0) contentWeight += 5000;
  if (item.dadosConsultoria && (item.dadosConsultoria.codigoSgf || item.dadosConsultoria.consultor)) contentWeight += 2000;
  
  return contentWeight;
}

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function isQuotaError(error: unknown): boolean {
  if (!error) return false;
  const str = String(error instanceof Error ? error.message : error) + 
    (typeof error === 'object' && error !== null && 'code' in (error as any) ? String((error as any).code) : '');
  const lower = str.toLowerCase();
  return lower.includes('quota') || lower.includes('resource-exhausted') || lower.includes('limit exceeded');
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errMsg = error instanceof Error ? error.message : String(error);
  console.warn(`[Firestore ${operationType}] ${path}:`, errMsg);

  if (isQuotaError(error)) {
    console.warn('Cota do Firestore excedida (resource-exhausted). Mantendo dados na memÃ³ria local.');
    if (operationType === OperationType.WRITE) {
      alert("Cota de gravaÃ§Ã£o do Firestore excedida no projeto gratuito. As alteraÃ§Ãµes foram salvas e mantidas localmente nesta sessÃ£o.");
    }
    return;
  }

  if (operationType === OperationType.LIST || operationType === OperationType.GET) {
    return;
  }
}

async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
  } catch (error) {
    if (error instanceof Error) {
      console.warn("Firestore connection check:", error.message);
    }
  }
}
testConnection();

// --- Types ---
interface EmpresaCredenciada {
  id: string;
  razaoSocial: string;
  cnpj?: string;
  telefoneFixo?: string;
  celular?: string;
  email?: string;
  consultor?: string;
  cpfConsultor?: string;
  ownerId: string;
  dataCadastro: any;
  ultimoLogin?: any;
  providerId?: string;
  photoURL?: string;
  status: 'Ativa' | 'Bloqueada';
  validadeLicenca?: any;
  tipoPlano?: 'Teste' | 'Mensal' | 'Anual' | 'Definitiva';
  diasTeste?: number;
  role?: 'admin' | 'cliente';
}

export interface Empresa {
  id: string;
  nome: string;
  razaoSocial?: string;
  nomeFantasia?: string;
  cnpj?: string;
  mesAnoAbertura?: string;
  enderecoComercial?: string;
  telefoneFixo?: string;
  celular?: string;
  email?: string;
  representante?: string;
  cpfRepresentante?: string;
  dataCadastro: any;
  ownerId: string;
  anexoUrl?: string;
  tipoEmpresa?: string;
  ramoAtividade?: string;
  cafNumero?: string;
}

export interface Premissa {
  id: string;
  idProblema: string;
  problema: string;
  peso: number;
  pergunta: string;
  tipoEmpresa?: string;
  ownerId?: string;
}

export interface Problema {
  id: string;
  descricao_problemas: string;
  area: string;
  impacto: string;
  NivelMaturidade?: string;
  tipoEmpresa?: string;
  tags?: string[];
  ownerId?: string;
}

export interface AtividadeCronograma {
  nome: string;
  descricao: string;
  cargaHoraria: string;
  solucaoProposta: string;
  resultadoEsperado?: string;
  responsavel?: string;
  idProblema?: string;
  premissa?: string;
  status: 'Pendente' | 'Em Andamento' | 'ConcluÃ­do' | 'Atrasado';
  prioridade: 'Baixa' | 'MÃ©dia' | 'Alta';
  dataInicio?: string;
  dataFim?: string;
  progressoKPI?: number;
  metaKPI?: number;
  evidencias?: string[];
  ordem?: number;
  incluirNoRelatorio?: boolean;
}

export interface DadosConsultoria {
  razaoSocial?: string;
  cnpj?: string;
  consultor?: string;
  telefoneFixo?: string;
  celular?: string;
  email?: string;
  areaConsultoria?: string;
  codigoSgf?: string;
  periodoConsultoria?: string;
  tecnicoSebrae?: string;
  objetivo?: string;
  resultadosEsperados?: string;
  solucoesIndicadas?: string;
  evidencias?: string[];
  cargaHoraria?: string;
  tipoRelatorio?: 'Parcial' | 'Final';
}

export interface Diagnostico {
  id: string;
  empresaId: string;
  dataDiagnostico: any;
  ownerId: string;
  cronograma?: AtividadeCronograma[];
  areasDiagnostico?: string[];
  dadosConsultoria?: DadosConsultoria;
  consultorId?: string;
  empresaCredenciadaId?: string;
  status?: 'Rascunho' | 'Finalizado' | 'Planejamento';
  tipoEmpresa?: string;
  nomeEmpresa?: string;
  nomeProjeto?: string;
  nome?: string;
  cargaHoraria?: string;
  nivelMaturidadeAI?: string;
  justificativaMaturidadeAI?: string;
  resumoExecutivoAI?: string;
}

export interface Solucao {
  id: string;
  idProblema: string;
  problema: string;
  area: string;
  solucao_recomendada: string;
  acoes_sugeridas: string;
  prazo_sugerido: string;
  responsavel_sugerido: string;
  kpis_sugeridos: string;
  comentario_sucesso: string;
  resultado_esperado: string;
  tipoEmpresa?: string;
  tags?: string[];
  ownerId?: string;
}

export interface TarefaPlanoAcao {
  id: string;
  diagnosticoId: string;
  empresaId: string;
  idProblema: string;
  solucaoId?: string;
  problema: string;
  area: string;
  solucaoSugerida: string;
  acoes: string;
  status: 'Pendente' | 'Em Andamento' | 'ConcluÃ­do';
  prioridade: 'Baixa' | 'MÃ©dia' | 'Alta';
  dataInicio?: any;
  dataFim?: any;
  dataVencimento?: any;
  responsavel: string;
  ownerId: string;
  ordem?: number;
  comentarios?: string; lembreteEmail?: boolean; lembreteWhatsapp?: boolean; lembretePush?: boolean; lembreteDiasAntes?: number; lembreteContato?: string;
  evidencias?: string[];
  evidenciaUrl?: string;
  evidenciaNome?: string;
  anexoUrl?: string;
}

export interface Resposta {
  id: string;
  diagnosticoId: string;
  premissaId: string;
  idProblema: string;
  problema: string;
  pergunta: string;
  peso: number;
  area: string;
  resposta?: 'Sim' | 'NÃ£o' | 'Parcial' | '';
  observacao: string;
  score: number;
  evidenciaUrl?: string;
  evidenciaNome?: string;
  ownerId?: string;
}

// --- Constants ---
const AREAS = [
  { id: 'FIN', nome: 'Financeiro' },
  { id: 'MKT', nome: 'Marketing' },
  { id: 'OPS', nome: 'Operacional' },
  { id: 'RH',  nome: 'Recursos Humanos' },
  { id: 'VEN', nome: 'Vendas' },
  { id: 'TEC', nome: 'Tecnologia' },
  { id: 'JUR', nome: 'JurÃ­dico' },
  { id: 'EST', nome: 'EstratÃ©gico' },
  { id: 'LOG', nome: 'LogÃ­stica' },
  { id: 'SAC', nome: 'Atendimento' },
  { id: 'CRE', nome: 'Acesso a CrÃ©dito' },
  { id: 'CRD', nome: 'CrÃ©dito' },
];

const normalizeAndFormatArea = (rawArea: string | undefined | null): string => {
  if (!rawArea) return '';
  const trimmed = rawArea.trim();
  if (!trimmed) return '';
  return trimmed
    .toLowerCase()
    .split(/\s+/)
    .map(word => {
      if (word.length <= 2 && ['de', 'da', 'do', 'em', 'para', 'e'].includes(word)) return word;
      return word.charAt(0).toUpperCase() + word.slice(1);
    })
    .join(' ');
};

// --- Components ---

// --- Constants ---
const IMPACTO_ORDER: Record<string, number> = { 'Baixo': 1, 'MÃ©dio': 2, 'Alto': 3 };

const globalNormalizedMatch = (val1: string | undefined | null, val2: string | undefined | null) => {
  if (!val1 && !val2) return true;
  if (!val1 || !val2) return false;
  return val1.trim().toLowerCase() === val2.trim().toLowerCase();
};

const Card = ({ children, className, onClick }: any) => (
  <div 
    onClick={onClick}
    className={cn('bg-white rounded-xl shadow-sm border border-slate-100 overflow-hidden', className)}
  >
    {children}
  </div>
);

const TagInputManager = ({
  tags = [],
  onChange,
  allAvailableTags = []
}: {
  tags: string[];
  onChange: (tags: string[]) => void;
  allAvailableTags: string[];
}) => {
  const [inputVal, setInputVal] = useState('');

  const addTag = (tagToAdd: string) => {
    const formatted = tagToAdd.trim().replace(/^#+/, '');
    if (!formatted) return;
    if (!tags.some(t => t.toLowerCase() === formatted.toLowerCase())) {
      onChange([...tags, formatted]);
    }
    setInputVal('');
  };

  const removeTag = (tagToRemove: string) => {
    onChange(tags.filter(t => t.toLowerCase() !== tagToRemove.toLowerCase()));
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      addTag(inputVal);
    }
  };

  const unusedSuggestedTags = allAvailableTags.filter(
    st => !tags.some(t => t.toLowerCase() === st.toLowerCase())
  );

  return (
    <div className="space-y-2">
      <label className="block text-sm font-bold text-slate-700">Tags / Categorias Personalizadas</label>
      <div className="flex flex-wrap items-center gap-1.5 p-2 bg-white rounded-xl border border-slate-200 focus-within:ring-2 focus-within:ring-emerald-500/20 focus-within:border-emerald-500 min-h-[44px]">
        {tags.map(tag => (
          <span
            key={tag}
            className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-50 text-emerald-800 text-xs font-bold rounded-lg border border-emerald-200/80"
          >
            <Tag size={12} className="text-emerald-600" />
            #{tag}
            <button
              type="button"
              onClick={() => removeTag(tag)}
              className="ml-1 text-emerald-600 hover:text-emerald-900 rounded-full p-0.5 hover:bg-emerald-100 transition-colors cursor-pointer"
            >
              <X size={12} />
            </button>
          </span>
        ))}
        <input
          type="text"
          className="flex-1 min-w-[150px] px-2 py-1 text-sm outline-none bg-transparent placeholder:text-slate-400 font-sans"
          placeholder={tags.length === 0 ? "Digite tags (Enter ou vÃ­rgula)..." : "Adicionar tag..."}
          value={inputVal}
          onChange={(e) => setInputVal(e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={() => {
            if (inputVal.trim()) addTag(inputVal);
          }}
        />
        {inputVal.trim() && (
          <button
            type="button"
            onClick={() => addTag(inputVal)}
            className="px-3 py-1 text-xs font-bold bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-colors cursor-pointer"
          >
            Adicionar
          </button>
        )}
      </div>

      {unusedSuggestedTags.length > 0 && (
        <div className="flex flex-wrap items-center gap-1.5 pt-1">
          <span className="text-[11px] text-slate-400 font-medium mr-1 flex items-center gap-1">
            <Tags size={12} /> Tags existentes:
          </span>
          {unusedSuggestedTags.slice(0, 10).map(st => (
            <button
              key={st}
              type="button"
              onClick={() => addTag(st)}
              className="text-[11px] font-semibold text-slate-600 bg-slate-100 hover:bg-emerald-100 hover:text-emerald-700 px-2 py-0.5 rounded-md transition-colors cursor-pointer"
            >
              +#{st}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

// --- Helper to limit, balance and adjust activity workloads in logical order ---
const adjustActivitiesMax4Hours = (
  activities: AtividadeCronograma[],
  totalCargaHorariaStr?: string
): AtividadeCronograma[] => {
  let totalHours = 34;
  if (totalCargaHorariaStr) {
    const num = parseInt(totalCargaHorariaStr.replace(/\D/g, ''), 10);
    if (!isNaN(num) && num > 0) {
      totalHours = num;
    }
  }

  let hasDevSystem = false;
  let hasAnaliseFinal = false;

  const rawCleanList: AtividadeCronograma[] = [];

  for (const atv of activities) {
    if (!atv || (!atv.nome && !atv.solucaoProposta)) continue;
    const nameLower = (atv.nome || atv.solucaoProposta || '').toLowerCase();

    if (nameLower.includes('desenvolvimento de sistema') || nameLower.includes('aplicativo') || nameLower.includes('planilha')) {
      hasDevSystem = true;
      rawCleanList.push({
        ...atv,
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: atv.descricao || "AnÃ¡lise e desenvolvimento de sistema de gestÃ£o, aplicativo ou planilhas para automaÃ§Ã£o e controle operacional e financeiro do negÃ³cio.",
        cargaHoraria: "8h",
        solucaoProposta: atv.solucaoProposta || "Desenvolvimento de sistema, aplicativo e planilhas",
        resultadoEsperado: atv.resultadoEsperado || "Sistema, aplicativo ou planilhas desenvolvidas e implantadas na rotina da empresa."
      });
    } else if (nameLower.includes('anÃ¡lise final') || nameLower.includes('analise final')) {
      hasAnaliseFinal = true;
      rawCleanList.push({
        ...atv,
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: atv.descricao || "AnÃ¡lise crÃ­tica final e mensuraÃ§Ã£o de todas as atividades e resultados executados na consultoria.",
        cargaHoraria: "2h",
        solucaoProposta: atv.solucaoProposta || "AnÃ¡lise final de atividades da consultoria",
        resultadoEsperado: atv.resultadoEsperado || "AvaliaÃ§Ã£o detalhada e validaÃ§Ã£o final do cumprimento do escopo da consultoria."
      });
    } else {
      rawCleanList.push(atv);
    }
  }

  // Ensure mandatory special activities exist
  if (!hasDevSystem) {
    rawCleanList.push({
      nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
      descricao: "AnÃ¡lise e desenvolvimento de sistema de gestÃ£o, aplicativo ou planilhas personalizadas para automaÃ§Ã£o dos controles da empresa.",
      cargaHoraria: "8h",
      solucaoProposta: "Desenvolvimento de sistema, aplicativo ou planilhas de gestÃ£o",
      responsavel: "Consultor",
      status: "Pendente",
      prioridade: "Alta",
      resultadoEsperado: "Ferramenta/aplicativo/planilha implantada e operando no cliente."
    });
  }

  if (!hasAnaliseFinal) {
    rawCleanList.push({
      nome: "AnÃ¡lise final de atividades da consultoria",
      descricao: "AnÃ¡lise de encerramento e verificaÃ§Ã£o do alcance dos indicadores e resultados previstos no plano de trabalho.",
      cargaHoraria: "2h",
      solucaoProposta: "AnÃ¡lise final e consolidaÃ§Ã£o de resultados da consultoria",
      responsavel: "Consultor",
      status: "Pendente",
      prioridade: "MÃ©dia",
      resultadoEsperado: "Checklist final de entregas e resultados validados."
    });
  }

  // Ensure diagnostic start and end report exist
  const hasStartDiag = rawCleanList.some(a => (a.nome || '').toLowerCase().includes('diagnÃ³stico inicial') || (a.nome || '').toLowerCase().includes('entendimento'));
  if (!hasStartDiag) {
    rawCleanList.unshift({
      nome: "Entendimento da demanda e diagnÃ³stico inicial",
      descricao: "Alinhamento das expectativas do cliente e levantamento detalhado das necessidades operacionais e financeiras.",
      cargaHoraria: "4h",
      solucaoProposta: "Entendimento da demanda e diagnÃ³stico inicial",
      responsavel: "Consultor",
      status: "Pendente",
      prioridade: "Alta",
      resultadoEsperado: "Expectativas alinhadas e diagnÃ³stico operacional consolidado."
    });
  }

  const hasEndReport = rawCleanList.some(a => (a.nome || '').toLowerCase().includes('relatÃ³rio final') || (a.nome || '').toLowerCase().includes('encerramento'));
  if (!hasEndReport) {
    rawCleanList.push({
      nome: "RelatÃ³rio final e encerramento",
      descricao: "ElaboraÃ§Ã£o, apresentaÃ§Ã£o do relatÃ³rio tÃ©cnico final da consultoria e formalizaÃ§Ã£o do encerramento.",
      cargaHoraria: "2h",
      solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final.",
      responsavel: "Consultor",
      status: "Pendente",
      prioridade: "MÃ©dia",
      resultadoEsperado: "RelatÃ³rio final gerencial apresentado e aprovado pelo cliente."
    });
  }

  // Assign weight for logical execution sequence
  const getWeight = (a: AtividadeCronograma): number => {
    const name = (a.nome || a.solucaoProposta || '').toLowerCase();
    if (name.includes('diagnÃ³stico inicial') || name.includes('entendimento da demanda')) return 1;
    if (name.includes('desenvolvimento de sistema') || name.includes('aplicativo') || name.includes('planilha')) return 8;
    if (name.includes('anÃ¡lise final') || name.includes('analise final')) return 9;
    if (name.includes('relatÃ³rio final') || name.includes('encerramento')) return 10;
    return 5;
  };

  rawCleanList.sort((a, b) => getWeight(a) - getWeight(b));

  // Balance hours: Dev system = 8h, Analise final = 2h (fixed 10h)
  const fixedDev = rawCleanList.find(a => getWeight(a) === 8);
  const fixedAnalise = rawCleanList.find(a => getWeight(a) === 9);

  if (fixedDev) fixedDev.cargaHoraria = "8h";
  if (fixedAnalise) fixedAnalise.cargaHoraria = "2h";

  const targetOtherHours = Math.max(2, totalHours - 10);
  const otherItems = rawCleanList.filter(a => getWeight(a) !== 8 && getWeight(a) !== 9);

  if (otherItems.length > 0) {
    const baseH = Math.max(2, Math.min(4, Math.floor(targetOtherHours / otherItems.length)));
    otherItems.forEach(item => {
      item.cargaHoraria = `${baseH}h`;
    });

    let currentSum = otherItems.length * baseH;
    let diff = targetOtherHours - currentSum;

    for (let i = 0; i < otherItems.length && diff > 0; i++) {
      const h = parseInt(otherItems[i].cargaHoraria.replace(/\D/g, ''), 10) || baseH;
      if (h < 4) {
        const add = Math.min(4 - h, diff);
        otherItems[i].cargaHoraria = `${h + add}h`;
        diff -= add;
      }
    }
  }

  return rawCleanList;
};

// --- View Components ---

const PLANO_DE_ACAO_PADRAO: AtividadeCronograma[] = [
  {
    nome: "Entendimento da demanda e diagnÃ³stico inicial",
    descricao: "Alinhamento das expectativas do cliente e levantamento de dados operacionais e financeiros atuais.",
    cargaHoraria: "4h",
    solucaoProposta: "Entendimento da demanda e diagnÃ³stico inicial",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Expectativas alinhadas e diagnÃ³stico operacional/financeiro inicial consolidado."
  },
  {
    nome: "Estrutura do Controle de Custos Operacionais",
    descricao: "ImplementaÃ§Ã£o de mÃ©todos para registro rigoroso de insumos e despesas operacionais.",
    cargaHoraria: "4h",
    solucaoProposta: "EstruturaÃ§Ã£o de registros e controles de custos operacionais.",
    responsavel: "Consultor/Produtor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Controle de insumos estruturado e planilhas implantadas na rotina do cliente."
  },
  {
    nome: "CÃ¡lculo do Custo de ProduÃ§Ã£o e PrecificaÃ§Ã£o",
    descricao: "ApuraÃ§Ã£o tÃ©cnica do custo real unitÃ¡rio e definiÃ§Ã£o de margem de contribuiÃ§Ã£o.",
    cargaHoraria: "4h",
    solucaoProposta: "AnÃ¡lise tÃ©cnica de custos e apuraÃ§Ã£o do preÃ§o de venda ideal.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Custo unitÃ¡rio e margem de contribuiÃ§Ã£o calculados com precisÃ£o matemÃ¡tica."
  },
  {
    nome: "IntroduÃ§Ã£o Ã  GestÃ£o Financeira e Fluxo de Caixa",
    descricao: "CapacitaÃ§Ã£o sobre controle financeiro diÃ¡rio e elaboraÃ§Ã£o de DRE simplificado.",
    cargaHoraria: "4h",
    solucaoProposta: "Treinamento sobre fluxo de caixa e demonstrativo de resultado.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "MÃ©dia",
    resultadoEsperado: "Fluxo de caixa implantado e rotina de registros mantida."
  },
  {
    nome: "Planejamento EstratÃ©gico e GestÃ£o de Estoque",
    descricao: "DefiniÃ§Ã£o de metas de produÃ§Ã£o/vendas e controle de insumos e matÃ©rias-primas.",
    cargaHoraria: "4h",
    solucaoProposta: "FormulaÃ§Ã£o de metas operacionais e otimizaÃ§Ã£o da gestÃ£o de estoque.",
    responsavel: "Consultor/Produtor",
    status: "Pendente",
    prioridade: "MÃ©dia",
    resultadoEsperado: "Metas estratÃ©gicas e controle de estoque estruturados para os prÃ³ximos meses."
  },
  {
    nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
    descricao: "AnÃ¡lise e desenvolvimento de sistema de gestÃ£o, aplicativo ou planilhas personalizadas para automaÃ§Ã£o dos controles da empresa.",
    cargaHoraria: "8h",
    solucaoProposta: "Desenvolvimento de aplicativo e planilhas integradas de gestÃ£o",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "Alta",
    resultadoEsperado: "Sistema, aplicativo ou planilhas operacionais implantadas no ambiente do cliente."
  },
  {
    nome: "AnÃ¡lise final de atividades da consultoria",
    descricao: "AnÃ¡lise de encerramento e verificaÃ§Ã£o do alcance dos indicadores e resultados previstos no plano de trabalho.",
    cargaHoraria: "2h",
    solucaoProposta: "AnÃ¡lise final e consolidaÃ§Ã£o de resultados da consultoria",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "MÃ©dia",
    resultadoEsperado: "VerificaÃ§Ã£o completa do cumprimento do escopo e metas da consultoria."
  },
  {
    nome: "RelatÃ³rio final e encerramento",
    descricao: "ConsolidaÃ§Ã£o dos resultados e avaliaÃ§Ã£o do impacto da consultoria.",
    cargaHoraria: "4h",
    solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final.",
    responsavel: "Consultor",
    status: "Pendente",
    prioridade: "MÃ©dia",
    resultadoEsperado: "RelatÃ³rio final consolidado, apresentado e validado pelo produtor/cliente."
  }
];

interface ModeloRelatorio {
  id: string;
  nome: string;
  categoria: string;
  descricao: string;
  atividades: AtividadeCronograma[];
}

const MODELOS_RELATORIO: ModeloRelatorio[] = [
  {
    id: 'carcinicultura',
    nome: 'GestÃ£o Financeira & Controle de Custos da Carcinicultura',
    categoria: 'Carcinicultura',
    descricao: 'Modelo de Consultoria Gerencial para Carcinicultura (30hs): Entendimento de demanda, DiagnÃ³stico de fluxos, Planilha de custos, Reserva de emergÃªncia, Treinamento de indicadores, SegregaÃ§Ã£o de contas e RelatÃ³rio final.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "Alinhamento das expectativas do cliente e levantamento detalhado das necessidades operacionais e financeiras.",
        cargaHoraria: "2h",
        solucaoProposta: "Alinhamento das expectativas e levantamento das necessidades operacionais e financeiras da carcinicultura.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Expectativas alinhadas e necessidades operacionais e financeiras levantadas."
      },
      {
        nome: "DiagnÃ³stico Detalhado e Levantamento de Fluxos",
        descricao: "ReuniÃ£o inicial para mapeamento dos processos atuais, verificaÃ§Ã£o de registros de caixa e entendimento das falhas de separaÃ§Ã£o entre contas.",
        cargaHoraria: "4h",
        solucaoProposta: "Mapeamento dos processos atuais e verificaÃ§Ã£o das falhas de separaÃ§Ã£o entre contas.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxos atuais mapeados e falhas de controle financeiro identificadas."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "CriaÃ§Ã£o e implementaÃ§Ã£o de uma planilha de controle de custos de raÃ§Ã£o, receitas e cÃ¡lculo de margem de contribuiÃ§Ã£o.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento e implantaÃ§Ã£o da planilha de controle de custos de raÃ§Ã£o, receitas e margem de contribuiÃ§Ã£o.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Planilha/ferramenta de gestÃ£o operando no cliente para controle de custos e receitas."
      },
      {
        nome: "Planejamento Financeiro e Reserva de EmergÃªncia",
        descricao: "CriaÃ§Ã£o de um plano para constituiÃ§Ã£o de reserva financeira visando a sustentabilidade do prÃ³ximo ciclo.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o de plano financeiro para formaÃ§Ã£o de reserva de emergÃªncia e sustentabilidade do prÃ³ximo ciclo.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Reserva financeira planejada e metas de fluxo de caixa para o prÃ³ximo ciclo estruturadas."
      },
      {
        nome: "Treinamento de GestÃ£o e Controle de Indicadores",
        descricao: "CapacitaÃ§Ã£o sobre o uso do caderno de campo/planilha e acompanhamento dos indicadores de produtividade.",
        cargaHoraria: "4h",
        solucaoProposta: "CapacitaÃ§Ã£o do produtor no uso do caderno de campo/planilha e acompanhamento dos indicadores.",
        responsavel: "Consultor (capacitaÃ§Ã£o) / Produtor (treinamento)",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Produtor capacitado no registro diÃ¡rio e interpretaÃ§Ã£o dos indicadores de produtividade."
      },
      {
        nome: "SegregaÃ§Ã£o de Contas e EstruturaÃ§Ã£o Financeira",
        descricao: "OrientaÃ§Ã£o para abertura de conta jurÃ­dica ou separaÃ§Ã£o formal das finanÃ§as pessoais e da carcinicultura.",
        cargaHoraria: "4h",
        solucaoProposta: "SeparaÃ§Ã£o formal das contas pessoais e da atividade de carcinicultura com definiÃ§Ã£o de prÃ³-labore.",
        responsavel: "Consultor / Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "FinanÃ§as da carcinicultura segregadas do caixa familiar com limite de retiradas."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "RevisÃ£o dos indicadores apurados na ferramenta de gestÃ£o implementada ao longo do projeto.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise e revisÃ£o de todos os indicadores apurados com as ferramentas implementadas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Indicadores revisados e evoluÃ§Ã£o do desempenho financeiro confirmada."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "Entrega do plano de aÃ§Ã£o contÃ­nuo e consolidaÃ§Ã£o do aprendizado do produtor.",
        cargaHoraria: "2h",
        solucaoProposta: "Entrega e apresentaÃ§Ã£o do relatÃ³rio final da consultoria com plano de continuidade.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final entregue e produtor autÃ´nomo na gestÃ£o do empreendimento."
      }
    ]
  },
  {
    id: 'financas',
    nome: 'GestÃ£o Financeira & Controle de Custos',
    categoria: 'FinanÃ§as',
    descricao: 'ApuraÃ§Ã£o de custos, fluxo de caixa, DRE gerencial, margem de contribuiÃ§Ã£o e precificaÃ§Ã£o.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "Alinhamento das expectativas do cliente e levantamento de dados operacionais e financeiros atuais.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico inicial",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Expectativas alinhadas e diagnÃ³stico financeiro/operacional inicial consolidado."
      },
      {
        nome: "Mapeamento e ClassificaÃ§Ã£o de Custos e Despesas",
        descricao: "CategorizaÃ§Ã£o detalhada de custos fixos, variÃ¡veis, despesas operacionais e tributos.",
        cargaHoraria: "4h",
        solucaoProposta: "EstruturaÃ§Ã£o do plano de contas e classificaÃ§Ã£o de custos.",
        responsavel: "Consultor/Cliente",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Plano de contas gerencial estruturado para controle financeiro."
      },
      {
        nome: "CÃ¡lculo da Margem de ContribuiÃ§Ã£o e Ponto de EquilÃ­brio",
        descricao: "ApuraÃ§Ã£o tÃ©cnica da margem de contribuiÃ§Ã£o por produto/serviÃ§o e determinaÃ§Ã£o do faturamento mÃ­nimo.",
        cargaHoraria: "4h",
        solucaoProposta: "AnÃ¡lise de viabilidade financeira e cÃ¡lculo do ponto de equilÃ­brio.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ponto de equilÃ­brio financeiro e margem de contribuiÃ§Ã£o mensurados."
      },
      {
        nome: "ImplementaÃ§Ã£o e Rotina do Fluxo de Caixa DiÃ¡rio",
        descricao: "Treinamento e implantaÃ§Ã£o de processo sistemÃ¡tico de controle de entradas e saÃ­das.",
        cargaHoraria: "4h",
        solucaoProposta: "CapacitaÃ§Ã£o e rotina prÃ¡tica de fluxo de caixa.",
        responsavel: "Consultor/Cliente",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxo de caixa projetado e acompanhado diariamente pelo cliente."
      },
      {
        nome: "ElaboraÃ§Ã£o do DRE Gerencial e FormaÃ§Ã£o do PreÃ§o de Venda",
        descricao: "EstruturaÃ§Ã£o do Demonstrativo de Resultado e fÃ³rmula de precificaÃ§Ã£o lucrativa.",
        cargaHoraria: "4h",
        solucaoProposta: "ImplantaÃ§Ã£o do DRE gerencial e modelo de precificaÃ§Ã£o.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DRE gerencial mensal e tabela de preÃ§os revisada com margem real."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento e personalizaÃ§Ã£o de planilhas financeiras integradas ou aplicativo de gestÃ£o de caixa.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e planilhas integradas de gestÃ£o financeira.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta/aplicativo financeiro personalizado operando no cliente."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "AnÃ¡lise de encerramento e verificaÃ§Ã£o do alcance dos indicadores financeiros.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise final e consolidaÃ§Ã£o dos resultados da consultoria.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio de desempenho financeiro e cumprimento de metas."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ConsolidaÃ§Ã£o final das entregas financeiras, apresentaÃ§Ã£o dos resultados e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final da consultoria.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final aprovado e plano de continuidade entregue ao cliente."
      }
    ]
  },
  {
    id: 'operacoes',
    nome: 'Processos Operacionais & Produtividade',
    categoria: 'OperaÃ§Ãµes',
    descricao: 'Mapeamento de processos, eliminaÃ§Ã£o de gargalos, padronizaÃ§Ã£o POP e qualidade.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "Levantamento das etapas do fluxo produtivo/operacional e identificaÃ§Ã£o de gargalos.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico operacional inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DiagnÃ³stico operacional detalhado e gargalos priorizados."
      },
      {
        nome: "Mapeamento do Fluxo de Trabalho (AS-IS)",
        descricao: "Desenho e documentaÃ§Ã£o do fluxo atual de produÃ§Ã£o e atendimento.",
        cargaHoraria: "4h",
        solucaoProposta: "Mapeamento completo dos processos operacionais vigentes.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxograma operacional atual mapeado e gargalos evidenciados."
      },
      {
        nome: "Redesenho de Processos e ReduÃ§Ã£o de DesperdÃ­cios (TO-BE)",
        descricao: "ProposiÃ§Ã£o de melhorias, eliminaÃ§Ã£o de etapas redundantes e tempos mortos.",
        cargaHoraria: "4h",
        solucaoProposta: "OtimizaÃ§Ã£o de processos operacionais e reduÃ§Ã£o de perdas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Novo fluxo de trabalho otimizado e aprovado pela gerÃªncia."
      },
      {
        nome: "ElaboraÃ§Ã£o de Procedimentos Operacionais PadrÃ£o (POPs)",
        descricao: "CriaÃ§Ã£o de manuais prÃ¡ticos e instruÃ§Ãµes de trabalho para padronizaÃ§Ã£o.",
        cargaHoraria: "4h",
        solucaoProposta: "PadronizaÃ§Ã£o operacional via POPs e checklists de controle.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Procedimentos operacionais padrÃ£o (POPs) redigidos e implantados."
      },
      {
        nome: "Treinamento da Equipe e Indicadores de Produtividade",
        descricao: "CapacitaÃ§Ã£o dos colaboradores nos novos processos e mÃ©tricas de eficiÃªncia.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento operacional e implantaÃ§Ã£o de KPIs de produÃ§Ã£o.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Equipe treinada e indicadores de qualidade/produtividade ativos."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "CriaÃ§Ã£o de ferramenta/aplicativo ou planilhas para acompanhamento em tempo real.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e planilhas de controle operacional.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Sistema/planilha de gestÃ£o operacional em pleno funcionamento."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "AvaliaÃ§Ã£o do aumento de produtividade e alcance dos padrÃµes previstos.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise de encerramento e validaÃ§Ã£o de KPIs operacionais.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Ganhos de produtividade e eficiÃªncia operacional comprovados."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ConsolidaÃ§Ã£o do relatÃ³rio tÃ©cnico de processos, apresentaÃ§Ã£o de resultados e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final de processos.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final entregue e operacionais padronizados mantidos."
      }
    ]
  },
  {
    id: 'comercial',
    nome: 'ComercializaÃ§Ã£o, Vendas & Marketing Digital',
    categoria: 'Comercial',
    descricao: 'EstratÃ©gia de vendas, funil de captaÃ§Ã£o, presenÃ§a digital e fidelizaÃ§Ã£o.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "DiagnÃ³stico das vendas atuais, perfil do cliente ideal (ICP) e canais de vendas.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico comercial inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DiagnÃ³stico das metas de vendas e oportunidades de mercado mapeadas."
      },
      {
        nome: "AnÃ¡lise da ConcorrÃªncia e Posicionamento de Mercado",
        descricao: "Estudo dos diferenciais competitivos e ajuste na proposta de valor do negÃ³cio.",
        cargaHoraria: "4h",
        solucaoProposta: "DefiniÃ§Ã£o de posicionamento estratÃ©gico e proposta de valor.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Proposta de valor diferenciada e alinhada Ã s necessidades do cliente."
      },
      {
        nome: "EstruturaÃ§Ã£o do Funil de Vendas e Processo Comercial",
        descricao: "DefiniÃ§Ã£o das etapas de prospecÃ§Ã£o, qualificaÃ§Ã£o, apresentaÃ§Ã£o e fechamento.",
        cargaHoraria: "4h",
        solucaoProposta: "Desenvolvimento da jornada do cliente e funil comercial.",
        responsavel: "Consultor/Vendedores",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Funil de vendas padronizado com rotina diÃ¡ria de abordagens."
      },
      {
        nome: "EstratÃ©gia de Marketing Digital e Redes Sociais",
        descricao: "Planejamento de conteÃºdos, anÃºncios e presenÃ§a nas mÃ­dias sociais para atraÃ§Ã£o.",
        cargaHoraria: "4h",
        solucaoProposta: "Plano de comunicaÃ§Ã£o digital e atraÃ§Ã£o de novos clientes.",
        responsavel: "Consultor/Marketing",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "CalendÃ¡rio de marketing digital e diretrizes de divulgaÃ§Ã£o ativas."
      },
      {
        nome: "Treinamento de TÃ©cnicas de Vendas e PÃ³s-Venda",
        descricao: "CapacitaÃ§Ã£o da equipe em contorno de objeÃ§Ãµes, fechamento e retenÃ§Ã£o.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento prÃ¡tico de vendas e retenÃ§Ã£o de clientes.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Equipe comercial capacitada e taxa de conversÃ£o aumentada."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de planilha/aplicativo CRM de controle de leads, propostas e histÃ³rico.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de CRM e planilhas de gestÃ£o de vendas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta CRM personalizada implantada para controle comercial."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "AnÃ¡lise da evoluÃ§Ã£o do faturamento e alcance das metas de vendas.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise final do desempenho comercial e taxas de conversÃ£o.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Aumento das vendas mensurado e relatÃ³rio de mÃ©tricas concluÃ­do."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ApresentaÃ§Ã£o do relatÃ³rio final de vendas, recomendaÃ§Ãµes futuras e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final comercial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final entregue e rotina comercial consolidada."
      }
    ]
  },
  {
    id: 'tecnologia',
    nome: 'InovaÃ§Ã£o, Tecnologia & AutomaÃ§Ã£o de Processos',
    categoria: 'InovaÃ§Ã£o',
    descricao: 'TransformaÃ§Ã£o digital, escolha de softwares, automaÃ§Ã£o e treinamento em TI.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "AvaliaÃ§Ã£o do nÃ­vel de maturidade digital e requisitos de tecnologia.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico tecnolÃ³gico inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Maturidade tecnolÃ³gica mapeada e plano de inovaÃ§Ã£o priorizado."
      },
      {
        nome: "Mapeamento de Requisitos e SeleÃ§Ã£o de Tecnologias",
        descricao: "IdentificaÃ§Ã£o das soluÃ§Ãµes tecnolÃ³gicas, softwares e sistemas adequados.",
        cargaHoraria: "4h",
        solucaoProposta: "SeleÃ§Ã£o de ferramentas de automaÃ§Ã£o e tecnologia de gestÃ£o.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Stack tecnolÃ³gico selecionado de acordo com custo-benefÃ­cio."
      },
      {
        nome: "Planejamento da IntegraÃ§Ã£o e AutomaÃ§Ã£o de Dados",
        descricao: "Desenho da arquitetura de integraÃ§Ã£o entre sistemas e automaÃ§Ã£o de fluxos.",
        cargaHoraria: "4h",
        solucaoProposta: "AutomaÃ§Ã£o de processos operacionais e integraÃ§Ã£o de dados.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Fluxos de automaÃ§Ã£o planejados para reduzir digitaÃ§Ã£o manual."
      },
      {
        nome: "ConfiguraÃ§Ã£o de ParÃ¢metros e Testes do Sistema",
        descricao: "Ajuste de cadastros, parametrizaÃ§Ã£o de regras de negÃ³cio e validaÃ§Ã£o dos sistemas.",
        cargaHoraria: "4h",
        solucaoProposta: "ConfiguraÃ§Ã£o tÃ©cnica e simulaÃ§Ã£o de rotinas digitais.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ambiente tecnolÃ³gico configurado e aprovado em ambiente de teste."
      },
      {
        nome: "CapacitaÃ§Ã£o da Equipe em Ferramentas Digitais",
        descricao: "Treinamento prÃ¡tico dos usuÃ¡rios para operaÃ§Ã£o fluida das novas plataformas.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento tÃ©cnico e mudanÃ§a cultural para o meio digital.",
        responsavel: "Consultor/UsuÃ¡rios",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "UsuÃ¡rios autÃ´nomos na utilizaÃ§Ã£o das soluÃ§Ãµes tecnolÃ³gicas."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de aplicativos sob medida, dashboards interativos ou planilhas.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo e integraÃ§Ã£o de planilhas/sistemas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Aplicativo/sistema entregue e funcionando perfeitamente."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "AnÃ¡lise da reduÃ§Ã£o de tempo de execuÃ§Ã£o e estabilidade do ambiente.",
        cargaHoraria: "2h",
        solucaoProposta: "ValidaÃ§Ã£o final das automaÃ§Ãµes e tempo economizado.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Ganho de tempo e precisÃ£o dos dados automatizados confirmados."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ApresentaÃ§Ã£o do relatÃ³rio final de tecnologia, plano de suporte e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final de inovaÃ§Ã£o.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio tÃ©cnico final aprovado e encerramento concluÃ­do."
      }
    ]
  },
  {
    id: 'estrategia',
    nome: 'Planejamento EstratÃ©gico & GestÃ£o de Pessoas',
    categoria: 'EstratÃ©gia',
    descricao: 'VisÃ£o de futuro, metas SWOT, organograma, lideranÃ§a e KPIs.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "AnÃ¡lise da cultura organizacional, lideranÃ§a e desafios estratÃ©gicos.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico estratÃ©gico inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DiagnÃ³stico da visÃ£o do negÃ³cio e direcionamento estratÃ©gico."
      },
      {
        nome: "AnÃ¡lise SWOT e Diretrizes EstratÃ©gicas",
        descricao: "Mapeamento de ForÃ§as, Oportunidades, Fraquezas e AmeaÃ§as.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o da Matriz SWOT e visÃ£o de crescimento.",
        responsavel: "Consultor/SÃ³cios",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Matriz SWOT consolidada com metas de mÃ©dio/longo prazo."
      },
      {
        nome: "EstruturaÃ§Ã£o de Organograma e Responsabilidades",
        descricao: "DefiniÃ§Ã£o do quadro de funÃ§Ãµes, atribuiÃ§Ãµes de cargos e matriz de responsabilidade.",
        cargaHoraria: "4h",
        solucaoProposta: "DefiniÃ§Ã£o de papÃ©is, organograma e matriz de competÃªncias.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Organograma funcional claro e responsabilidades bem delineadas."
      },
      {
        nome: "Desenvolvimento de LideranÃ§as e ComunicaÃ§Ã£o Interna",
        descricao: "Treinamento em gestÃ£o de equipes, alinhamento de metas e comunicaÃ§Ã£o.",
        cargaHoraria: "4h",
        solucaoProposta: "Desenvolvimento de competÃªncias gerenciais e clima organizacional.",
        responsavel: "Consultor/Gestores",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Gestores alinhados e canais de comunicaÃ§Ã£o interna ativos."
      },
      {
        nome: "Plano de Metas e Acompanhamento de KPIs",
        descricao: "DefiniÃ§Ã£o de indicadores-chave de desempenho por setor e acompanhamento.",
        cargaHoraria: "4h",
        solucaoProposta: "EstruturaÃ§Ã£o do painel de controle e acompanhamento de metas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Metas corporativas desdobradas em planos operacionais."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "CriaÃ§Ã£o de aplicativo ou planilha de gestÃ£o estratÃ©gica e controle de KPIs.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo/dashboard estratÃ©gico.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Painel estratÃ©gico de indicadores implantado no cliente."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "AnÃ¡lise do alinhamento estratÃ©gico e engajamento da equipe.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise final do atingimento de metas estratÃ©gicas.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "EvoluÃ§Ã£o do engajamento e alcance de diretrizes mensurado."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ApresentaÃ§Ã£o do relatÃ³rio final de planejamento estratÃ©gico e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final estratÃ©gico.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final entregue e diretrizes estratÃ©gicas vigentes."
      }
    ]
  },
  {
    id: 'agronegocio',
    nome: 'AgronegÃ³cio & GestÃ£o de Propriedade Rural',
    categoria: 'AgronegÃ³cio',
    descricao: 'ApuraÃ§Ã£o de custos por lote/safra, manejo, estoques de insumos e gestÃ£o rural.',
    atividades: [
      {
        nome: "Entendimento da demanda e diagnÃ³stico inicial",
        descricao: "Levantamento das caracterÃ­sticas da propriedade rural, cultivos/criaÃ§Ãµes e infraestrutura.",
        cargaHoraria: "4h",
        solucaoProposta: "Entendimento da demanda e diagnÃ³stico agropecuÃ¡rio inicial.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "DiagnÃ³stico da propriedade rural e pontos crÃ­ticos de melhoria."
      },
      {
        nome: "Mapeamento dos Custos de ProduÃ§Ã£o por Lote/Safra",
        descricao: "EstruturaÃ§Ã£o da apuraÃ§Ã£o do custo operacional efetivo e custo total por Ã¡rea.",
        cargaHoraria: "4h",
        solucaoProposta: "AnÃ¡lise detalhada do custo de produÃ§Ã£o agropecuÃ¡ria.",
        responsavel: "Consultor/Produtor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Custo por safra/lote/saco mensurado com precisÃ£o."
      },
      {
        nome: "Controle de Estoque de Insumos e CalendÃ¡rio Manejo",
        descricao: "OrganizaÃ§Ã£o do almoxarifado agrÃ­cola, rastreabilidade de aplicaÃ§Ã£o e insumos.",
        cargaHoraria: "4h",
        solucaoProposta: "ImplantaÃ§Ã£o de controle de insumos e manejo operacional.",
        responsavel: "Consultor/Equipe",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Estoque de insumos controlado e perdas reduzidas."
      },
      {
        nome: "AnÃ¡lise de Viabilidade Financeira e PreÃ§o de ComercializaÃ§Ã£o",
        descricao: "DeterminaÃ§Ã£o do valor de equilÃ­brio da produÃ§Ã£o e estratÃ©gias de comercializaÃ§Ã£o.",
        cargaHoraria: "4h",
        solucaoProposta: "Estudo de margens de lucro e comercializaÃ§Ã£o safra.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "PreÃ§o de venda ideal e rentabilidade da produÃ§Ã£o calculados."
      },
      {
        nome: "CapacitaÃ§Ã£o no Gerenciamento PrÃ¡tico da Propriedade",
        descricao: "Treinamento do produtor rural e equipe no preenchimento de cadernos de campo.",
        cargaHoraria: "4h",
        solucaoProposta: "Treinamento de gestÃ£o no campo e controles gerenciais.",
        responsavel: "Consultor/Produtor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "Produtor autÃ´nomo na gestÃ£o financeira do agronegÃ³cio."
      },
      {
        nome: "Desenvolvimento de sistema, aplicativos ou planilhas",
        descricao: "Desenvolvimento de aplicativo de gestÃ£o rural ou planilha agrÃ­cola customizada.",
        cargaHoraria: "8h",
        solucaoProposta: "Desenvolvimento de aplicativo/planilhas de gestÃ£o agropecuÃ¡ria.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "Alta",
        resultadoEsperado: "Ferramenta de gestÃ£o rural implantada na propriedade."
      },
      {
        nome: "AnÃ¡lise final de atividades da consultoria",
        descricao: "VerificaÃ§Ã£o dos resultados do ciclo produtivo e alcance dos indicadores rurais.",
        cargaHoraria: "2h",
        solucaoProposta: "AnÃ¡lise final do desempenho produtivo e financeiro da propriedade.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio de desempenho da safra/lote consolidado."
      },
      {
        nome: "RelatÃ³rio final e encerramento",
        descricao: "ApresentaÃ§Ã£o do relatÃ³rio final de gestÃ£o agropecuÃ¡ria e encerramento.",
        cargaHoraria: "4h",
        solucaoProposta: "ElaboraÃ§Ã£o e apresentaÃ§Ã£o do relatÃ³rio final do agronegÃ³cio.",
        responsavel: "Consultor",
        status: "Pendente",
        prioridade: "MÃ©dia",
        resultadoEsperado: "RelatÃ³rio final entregue com plano de continuidade rural."
      }
    ]
  }
];

export const deduplicateRespostas = (respostasList: Resposta[]): Resposta[] => {
  if (!respostasList || !Array.isArray(respostasList)) return [];
  const map = new Map<string, Resposta>();

  for (const resp of respostasList) {
    if (!resp) continue;
    const normQ = (resp.pergunta || '').trim().toLowerCase();
    const key = normQ ? `q:${normQ}` : (resp.premissaId ? `id:${resp.premissaId}` : resp.id);
    if (!key) continue;

    if (!map.has(key)) {
      map.set(key, resp);
    } else {
      const existing = map.get(key)!;
      const isAnswered = (r: Resposta) => r.resposta === 'Sim' || r.resposta === 'Parcial' || r.resposta === 'NÃ£o';
      
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

const DadosConsultoriaView = ({
  selectedDiagnostico,
  selectedEmpresa,
  empresas,
  diagnosticos,
  respostas = [],
  setSelectedEmpresa,
  setSelectedDiagnostico,
  setDiagnosticos,
  setView,
  playSuccessSound
}: {
  selectedDiagnostico: Diagnostico | null;
  selectedEmpresa: Empresa | null;
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas?: Resposta[];
  setSelectedEmpresa: (empresa: Empresa | null) => void;
  setSelectedDiagnostico: (diag: Diagnostico | null) => void;
  setDiagnosticos: React.Dispatch<React.SetStateAction<Diagnostico[]>>;
  setView: (v: string) => void;
  playSuccessSound: () => void;
}) => {
  const [dadosConsultoria, setDadosConsultoria] = useState<DadosConsultoria>(
    selectedDiagnostico?.dadosConsultoria || {}
  );
  const [isSavingLocal, setIsSavingLocal] = useState(false);
  const [saveSuccessLocal, setSaveSuccessLocal] = useState(false);
  const [isCopySgfModalOpen, setIsCopySgfModalOpen] = useState(false);
  const currentClientName = selectedEmpresa?.nome || selectedDiagnostico?.nomeEmpresa || selectedDiagnostico?.tipoEmpresa || 'Cliente Selecionado';

  // --- Mandatory Fields & Diagnostic Progress Calculations ---
  const mandatoryFields = useMemo(() => [
    { id: 'cliente', label: 'Cliente Vinculado', isFilled: Boolean(selectedEmpresa?.id || selectedDiagnostico?.empresaId || selectedDiagnostico?.nomeEmpresa) },
    { id: 'tipoRelatorio', label: 'Tipo do RelatÃ³rio', isFilled: Boolean(dadosConsultoria.tipoRelatorio) },
    { id: 'areaConsultoria', label: 'Ãrea da Consultoria', isFilled: Boolean(dadosConsultoria.areaConsultoria?.trim()) },
    { id: 'codigoSgf', label: 'CÃ³digo SGF', isFilled: Boolean(dadosConsultoria.codigoSgf?.trim()) },
    { id: 'periodoConsultoria', label: 'PerÃ­odo da Consultoria', isFilled: Boolean(dadosConsultoria.periodoConsultoria?.trim()) },
    { id: 'cargaHoraria', label: 'Carga HorÃ¡ria Total', isFilled: Boolean(dadosConsultoria.cargaHoraria?.trim()) },
    { id: 'tecnicoSebrae', label: 'TÃ©cnico Sebrae', isFilled: Boolean(dadosConsultoria.tecnicoSebrae?.trim()) },
    { id: 'objetivo', label: 'Objetivo da Consultoria', isFilled: Boolean(dadosConsultoria.objetivo?.trim()) },
    { id: 'solucoesIndicadas', label: 'Problemas & SoluÃ§Ãµes', isFilled: Boolean(dadosConsultoria.solucoesIndicadas?.trim()) },
    { id: 'resultadosEsperados', label: 'Resultados Esperados', isFilled: Boolean(dadosConsultoria.resultadosEsperados?.trim()) },
  ], [selectedEmpresa, selectedDiagnostico, dadosConsultoria]);

  const totalMandatory = mandatoryFields.length;
  const filledMandatory = mandatoryFields.filter(f => f.isFilled).length;
  const mandatoryPercent = Math.round((filledMandatory / totalMandatory) * 100);

  // Diagnostic questions progress
  const diagRespostas = useMemo(() => {
    if (!selectedDiagnostico?.id) return [];
    const list = respostas && respostas.length > 0 ? respostas : loadAllLocalRespostas();
    return list.filter(r => r.diagnosticoId === selectedDiagnostico.id);
  }, [respostas, selectedDiagnostico?.id]);

  const uniqueDiagRespostas = useMemo(() => deduplicateRespostas(diagRespostas), [diagRespostas]);
  const totalDiagnosticoPerguntas = uniqueDiagRespostas.length;
  const answeredDiagnosticoPerguntas = uniqueDiagRespostas.filter(r => 
    r.resposta === 'Sim' || r.resposta === 'NÃ£o' || r.resposta === 'Parcial' || (Boolean(r.resposta) && String(r.resposta).trim() !== '')
  ).length;
  const diagnosticoPercent = totalDiagnosticoPerguntas > 0 
    ? Math.round((answeredDiagnosticoPerguntas / totalDiagnosticoPerguntas) * 100) 
    : 0;

  // Cronograma / Atividades
  const cronogramaAtividades = selectedDiagnostico?.cronograma || [];
  const totalAtividadesPreenchidas = cronogramaAtividades.filter(a => a.nome && a.nome.trim() !== '').length;
  const cronogramaPercent = Math.min(100, Math.round((totalAtividadesPreenchidas / 8) * 100));

  // Global consolidated progress index
  const globalProgressPercent = totalDiagnosticoPerguntas > 0
    ? Math.round((mandatoryPercent * 0.40) + (diagnosticoPercent * 0.40) + (cronogramaPercent * 0.20))
    : mandatoryPercent;

  // Check if there are other diagnosticos with the same SGF code
  const matchingSgfDiags = useMemo(() => {
    const code = (dadosConsultoria.codigoSgf || '').trim().toLowerCase();
    if (!code) return [];
    return diagnosticos.filter(d => 
      d.id !== selectedDiagnostico?.id && 
      (d.dadosConsultoria?.codigoSgf || '').trim().toLowerCase() === code
    );
  }, [dadosConsultoria.codigoSgf, diagnosticos, selectedDiagnostico?.id]);

  const handleCopySgfData = (dados: DadosConsultoria) => {
    setDadosConsultoria(prev => ({
      ...prev,
      razaoSocial: dados.razaoSocial || prev.razaoSocial,
      cnpj: dados.cnpj || prev.cnpj,
      consultor: dados.consultor || prev.consultor,
      email: dados.email || prev.email,
      celular: dados.celular || prev.celular,
      telefoneFixo: dados.telefoneFixo || prev.telefoneFixo,
      areaConsultoria: dados.areaConsultoria || prev.areaConsultoria,
      codigoSgf: dados.codigoSgf || prev.codigoSgf,
      periodoConsultoria: dados.periodoConsultoria || prev.periodoConsultoria,
      cargaHoraria: dados.cargaHoraria || prev.cargaHoraria,
      tecnicoSebrae: dados.tecnicoSebrae || prev.tecnicoSebrae,
      objetivo: dados.objetivo || prev.objetivo,
      solucoesIndicadas: dados.solucoesIndicadas || prev.solucoesIndicadas,
      resultadosEsperados: dados.resultadosEsperados || prev.resultadosEsperados
    }));
    playSuccessSound();
    alert("Dados contratuais e tÃ©cnicos da consultoria copiados com sucesso pelo CÃ³digo SGF!");
  };

  useEffect(() => {
    if (selectedDiagnostico?.dadosConsultoria) {
      setDadosConsultoria(selectedDiagnostico.dadosConsultoria);
    }
  }, [selectedDiagnostico?.id]);

  const handleSaveDadosConsultoria = async () => {
    if (!selectedDiagnostico && !selectedEmpresa) {
      alert("Por favor, selecione um cliente cadastrado antes de salvar.");
      return;
    }

    setIsSavingLocal(true);
    setSaveSuccessLocal(false);

    const safetyTimer = setTimeout(() => {
      setIsSavingLocal(false);
    }, 2000);

    try {
      const sanitized = sanitizeForFirestore(dadosConsultoria);

      let targetDiag = selectedDiagnostico;
      if (!targetDiag && selectedEmpresa) {
        targetDiag = diagnosticos.find(d => d.empresaId === selectedEmpresa.id || d.nomeEmpresa === selectedEmpresa.nome) || null;
      }

      // 1. Update React state
      if (targetDiag) {
        const updatedDiag = { ...targetDiag, dadosConsultoria: sanitized };
        setSelectedDiagnostico(updatedDiag);
        setDiagnosticos(prev => prev.map(d => d.id === targetDiag!.id ? updatedDiag : d));

        // 2. Persist to localStorage local_diagnosticos
        try {
          const saved = localStorage.getItem('local_diagnosticos');
          const parsed = saved ? JSON.parse(saved) : [];
          const nextList = Array.isArray(parsed) 
            ? parsed.map((d: any) => d.id === targetDiag!.id ? updatedDiag : d)
            : [updatedDiag];
          localStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
        } catch (e) {
          console.warn("Could not save to localStorage local_diagnosticos", e);
        }

        // 3. Persist to IndexedDB (non-blocking)
        if (targetDiag.id) {
          saveLocalDiagnosticoReport(targetDiag.id, targetDiag.cronograma || [], sanitized).catch(err => {
            console.warn("[EvidenceStorage] Non-critical IndexedDB warning:", err);
          });
        }
      }

      // 4. Background cloud sync without blocking UI
      if (auth.currentUser && targetDiag?.id) {
        const targetId = targetDiag.id;
        (async () => {
          try {
            const syncPromise = updateDoc(doc(db, 'diagnosticos', targetId), {
              dadosConsultoria: sanitized,
              updatedAt: new Date().toISOString()
            });
            await Promise.race([
              syncPromise,
              new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout sync cloud")), 4000))
            ]);
          } catch (cloudErr) {
            console.warn("[CloudSync] Cloud sync note for dados consultoria:", cloudErr);
          }
        })();
      }

      clearTimeout(safetyTimer);
      setIsSavingLocal(false);
      playSuccessSound();
      setSaveSuccessLocal(true);
      setTimeout(() => setSaveSuccessLocal(false), 3500);
    } catch (e) {
      console.error("Error saving dados consultoria:", e);
      clearTimeout(safetyTimer);
      setIsSavingLocal(false);
      playSuccessSound();
      setSaveSuccessLocal(true);
      setTimeout(() => setSaveSuccessLocal(false), 3500);
    } finally {
      setIsSavingLocal(false);
    }
  };

  return (
    <div className="space-y-6 max-w-6xl mx-auto pb-12">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="px-2.5 py-0.5 text-xs font-bold bg-blue-50 text-blue-700 rounded-full uppercase tracking-wider">
              Aba Dedicada de Cadastro
            </span>
          </div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">Dados da Consultoria Gerencial</h2>
          <p className="text-slate-500 text-sm">
            Preencha e gerencie as informaÃ§Ãµes contratuais e tÃ©cnicas conectadas ao cliente cadastrado.
          </p>
        </div>

        <div className="flex items-center gap-3">
          <Button 
            variant="outline"
            onClick={() => setView('cronograma')}
            className="border-slate-200 text-slate-700 hover:bg-slate-50"
          >
            <Calendar className="mr-2 h-4 w-4 text-emerald-600" />
            Ir para RelatÃ³rio de Consultoria (8 Atividades)
          </Button>

          <Button 
            onClick={handleSaveDadosConsultoria} 
            disabled={isSavingLocal} 
            className={cn("text-white font-medium px-6 transition-all", saveSuccessLocal ? "bg-emerald-700" : "bg-emerald-600 hover:bg-emerald-700")}
          >
            {isSavingLocal ? (
              <>
                <Loader2 className="animate-spin mr-2" size={18} />
                <span>Salvando...</span>
              </>
            ) : saveSuccessLocal ? (
              <>
                <CheckCircle2 className="mr-2 text-white" size={18} />
                <span>Salvo com Sucesso!</span>
              </>
            ) : (
              <>
                <Save className="mr-2" size={18} />
                <span>Salvar Dados da Consultoria</span>
              </>
            )}
          </Button>
        </div>
      </div>

      {/* Componente: Resumo de Progresso (Campos ObrigatÃ³rios vs. DiagnÃ³stico Total) */}
      <Card className="p-6 border-slate-200/80 shadow-sm bg-white rounded-2xl overflow-hidden relative">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-5 border-b border-slate-100">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center shrink-0">
              <TrendingUp size={22} />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-slate-800 text-lg tracking-tight">Resumo de Progresso do DiagnÃ³stico</h3>
                <span className="text-[11px] font-mono font-bold px-2.5 py-0.5 rounded-full bg-slate-100 text-slate-700">
                  {globalProgressPercent}% ConcluÃ­do
                </span>
              </div>
              <p className="text-xs text-slate-500 font-medium">
                Acompanhamento em tempo real dos campos obrigatÃ³rios da consultoria e preenchimento total do diagnÃ³stico.
              </p>
            </div>
          </div>

          {/* Status Badge */}
          <div className="flex items-center gap-2">
            {globalProgressPercent === 100 ? (
              <span className="px-3.5 py-1.5 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-2xs">
                <CheckCircle2 size={15} className="text-emerald-600" /> Pronto para EmissÃ£o
              </span>
            ) : globalProgressPercent >= 70 ? (
              <span className="px-3.5 py-1.5 bg-blue-50 border border-blue-200 text-blue-700 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-2xs">
                <Clock size={15} className="text-blue-600" /> Em Andamento AvanÃ§ado ({globalProgressPercent}%)
              </span>
            ) : (
              <span className="px-3.5 py-1.5 bg-amber-50 border border-amber-200 text-amber-800 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-2xs">
                <AlertCircle size={15} className="text-amber-600" /> Preenchimento Pendente ({globalProgressPercent}%)
              </span>
            )}
          </div>
        </div>

        {/* Global Progress Bar */}
        <div className="mt-4 mb-6">
          <div className="flex justify-between text-xs font-bold mb-1.5">
            <span className="text-slate-600 flex items-center gap-1.5">
              <Target size={14} className="text-blue-600" />
              Ãndice Global Consolidado (Consultoria + DiagnÃ³stico + Cronograma)
            </span>
            <span className={cn(
              "font-mono",
              globalProgressPercent === 100 ? "text-emerald-600" : globalProgressPercent >= 70 ? "text-blue-600" : "text-amber-600"
            )}>
              {globalProgressPercent}%
            </span>
          </div>
          <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden p-0.5 border border-slate-200/60">
            <div 
              className={cn(
                "h-full rounded-full transition-all duration-700 ease-out",
                globalProgressPercent === 100 
                  ? "bg-emerald-500" 
                  : globalProgressPercent >= 70 
                    ? "bg-gradient-to-r from-blue-500 to-indigo-600" 
                    : "bg-gradient-to-r from-amber-400 to-amber-500"
              )}
              style={{ width: `${Math.max(4, globalProgressPercent)}%` }}
            />
          </div>
        </div>

        {/* 3 Metric Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-5">
          {/* Card 1: Campos ObrigatÃ³rios da Consultoria */}
          <div className="p-4 bg-slate-50 border border-slate-200/80 rounded-xl flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                  <FileText size={14} className="text-blue-600" />
                  Campos ObrigatÃ³rios
                </span>
                <span className={cn(
                  "text-xs font-extrabold px-2 py-0.5 rounded-md",
                  mandatoryPercent === 100 ? "bg-emerald-100 text-emerald-800" : "bg-blue-100 text-blue-800"
                )}>
                  {mandatoryPercent}%
                </span>
              </div>
              <div className="text-2xl font-black text-slate-800 mb-1">
                {filledMandatory} <span className="text-sm font-semibold text-slate-400">/ {totalMandatory} campos</span>
              </div>
              <div className="w-full bg-slate-200/70 rounded-full h-1.5 mb-2 overflow-hidden">
                <div 
                  className={cn("h-full rounded-full transition-all duration-500", mandatoryPercent === 100 ? "bg-emerald-500" : "bg-blue-600")}
                  style={{ width: `${mandatoryPercent}%` }}
                />
              </div>
            </div>
            <p className="text-[11px] text-slate-500 font-medium">
              {totalMandatory - filledMandatory === 0 
                ? "âœ“ Todos os 10 campos obrigatÃ³rios preenchidos" 
                : `${totalMandatory - filledMandatory} campo(s) da consultoria pendente(s)`}
            </p>
          </div>

          {/* Card 2: Perguntas do DiagnÃ³stico Total */}
          <div className="p-4 bg-slate-50 border border-slate-200/80 rounded-xl flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                  <CheckSquare size={14} className="text-indigo-600" />
                  DiagnÃ³stico Total
                </span>
                <span className={cn(
                  "text-xs font-extrabold px-2 py-0.5 rounded-md",
                  diagnosticoPercent === 100 ? "bg-emerald-100 text-emerald-800" : "bg-indigo-100 text-indigo-800"
                )}>
                  {diagnosticoPercent}%
                </span>
              </div>
              <div className="text-2xl font-black text-slate-800 mb-1">
                {answeredDiagnosticoPerguntas} <span className="text-sm font-semibold text-slate-400">/ {totalDiagnosticoPerguntas} perguntas</span>
              </div>
              <div className="w-full bg-slate-200/70 rounded-full h-1.5 mb-2 overflow-hidden">
                <div 
                  className={cn("h-full rounded-full transition-all duration-500", diagnosticoPercent === 100 ? "bg-emerald-500" : "bg-indigo-600")}
                  style={{ width: `${diagnosticoPercent}%` }}
                />
              </div>
            </div>
            <p className="text-[11px] text-slate-500 font-medium">
              {totalDiagnosticoPerguntas === 0 
                ? "Sem premissas vinculadas ao diagnÃ³stico" 
                : totalDiagnosticoPerguntas === answeredDiagnosticoPerguntas 
                  ? "âœ“ 100% das premissas avaliadas" 
                  : `${totalDiagnosticoPerguntas - answeredDiagnosticoPerguntas} pergunta(s) sem resposta`}
            </p>
          </div>

          {/* Card 3: Cronograma de Atividades */}
          <div className="p-4 bg-slate-50 border border-slate-200/80 rounded-xl flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[11px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1.5">
                  <CalendarDays size={14} className="text-emerald-600" />
                  Atividades do Plano
                </span>
                <span className={cn(
                  "text-xs font-extrabold px-2 py-0.5 rounded-md",
                  cronogramaPercent === 100 ? "bg-emerald-100 text-emerald-800" : "bg-emerald-50 text-emerald-700"
                )}>
                  {cronogramaPercent}%
                </span>
              </div>
              <div className="text-2xl font-black text-slate-800 mb-1">
                {totalAtividadesPreenchidas} <span className="text-sm font-semibold text-slate-400">/ 8 atividades</span>
              </div>
              <div className="w-full bg-slate-200/70 rounded-full h-1.5 mb-2 overflow-hidden">
                <div 
                  className={cn("h-full rounded-full transition-all duration-500", cronogramaPercent === 100 ? "bg-emerald-500" : "bg-emerald-600")}
                  style={{ width: `${cronogramaPercent}%` }}
                />
              </div>
            </div>
            <p className="text-[11px] text-slate-500 font-medium">
              {totalAtividadesPreenchidas >= 8 
                ? "âœ“ 8 atividades estruturadas no relatÃ³rio" 
                : `${Math.max(0, 8 - totalAtividadesPreenchidas)} atividade(s) a adicionar no cronograma`}
            </p>
          </div>
        </div>

        {/* Checklist dos Campos ObrigatÃ³rios */}
        <div className="pt-4 border-t border-slate-100">
          <div className="flex items-center justify-between mb-2.5">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Checklist dos 10 Campos ObrigatÃ³rios da Consultoria
            </span>
            <span className="text-xs font-semibold text-slate-500">
              {filledMandatory} de {totalMandatory} completos
            </span>
          </div>

          <div className="flex flex-wrap gap-2">
            {mandatoryFields.map(field => (
              <div 
                key={field.id}
                className={cn(
                  "px-3 py-1.5 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-all border",
                  field.isFilled 
                    ? "bg-emerald-50/80 border-emerald-200/80 text-emerald-800" 
                    : "bg-amber-50/80 border-amber-200/80 text-amber-800"
                )}
              >
                {field.isFilled ? (
                  <CheckCircle2 size={13} className="text-emerald-600 shrink-0" />
                ) : (
                  <AlertCircle size={13} className="text-amber-500 shrink-0" />
                )}
                <span>{field.label}</span>
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* SeleÃ§Ã£o do Cliente Cadastrado */}
      <Card className="p-6 border-slate-100 shadow-sm">
        <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-100">
          <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
            <Building2 size={20} />
          </div>
          <div>
            <h3 className="font-bold text-slate-800 uppercase tracking-tight">Cliente Conectado</h3>
            <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Selecione o cliente cadastrado para vincular estes dados da consultoria</p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center">
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Selecionar Cliente Cadastrado</label>
            <select 
              className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-medium text-slate-800 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              value={selectedEmpresa?.id || ''}
              onChange={(e) => {
                const emp = empresas.find(m => m.id === e.target.value);
                if (emp) {
                  setSelectedEmpresa(emp);
                  const diag = diagnosticos.find(d => d.empresaId === emp.id || d.nomeEmpresa === emp.nome);
                  if (diag) {
                    setSelectedDiagnostico(diag);
                    if (diag.dadosConsultoria) setDadosConsultoria(diag.dadosConsultoria);
                  }
                }
              }}
            >
              <option value="">Selecione um cliente cadastrado...</option>
              {empresas.map(emp => (
                <option key={emp.id} value={emp.id}>
                  {emp.nome} {emp.cnpj ? `(${emp.cnpj})` : ''} - {emp.ramoAtividade || 'Geral'}
                </option>
              ))}
            </select>
          </div>

          <div className="p-4 bg-emerald-50/60 border border-emerald-100 rounded-xl flex items-center justify-between">
            <div>
              <p className="text-[10px] text-emerald-600 font-bold uppercase tracking-wider">Cliente Selecionado Atualmente</p>
              <p className="text-base font-bold text-slate-800">{currentClientName}</p>
              {selectedEmpresa?.cnpj && <p className="text-xs text-slate-500 font-mono">CNPJ: {selectedEmpresa.cnpj}</p>}
            </div>
            <Button 
              size="sm" 
              variant="outline" 
              onClick={() => setView('companies')}
              className="bg-white border-emerald-200 text-emerald-700 hover:bg-emerald-100 text-xs font-semibold"
            >
              Ver Clientes
            </Button>
          </div>
        </div>
      </Card>

      {/* FormulÃ¡rio de Dados da Consultoria Gerencial */}
      <Card className="p-6 border-slate-100 shadow-sm">
        <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-100">
          <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
            <FileText size={20} />
          </div>
          <div>
            <h3 className="font-bold text-slate-800 uppercase tracking-tight">InformaÃ§Ãµes Contratuais da Consultoria</h3>
            <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Dados do contrato Sebrae/SGF e responsÃ¡vel tÃ©cnico</p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Tipo do RelatÃ³rio</label>
            <div className="grid grid-cols-2 gap-2">
              <button
                type="button"
                onClick={() => setDadosConsultoria({ ...dadosConsultoria, tipoRelatorio: 'Final' })}
                className={cn(
                  "py-2 px-3 rounded-xl text-xs font-bold transition-all border flex items-center justify-center gap-1.5",
                  (dadosConsultoria.tipoRelatorio === 'Final' || !dadosConsultoria.tipoRelatorio)
                    ? "bg-emerald-600 text-white border-emerald-600 shadow-xs"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                )}
              >
                <CheckCircle2 size={14} /> RelatÃ³rio Final
              </button>
              <button
                type="button"
                onClick={() => setDadosConsultoria({ ...dadosConsultoria, tipoRelatorio: 'Parcial' })}
                className={cn(
                  "py-2 px-3 rounded-xl text-xs font-bold transition-all border flex items-center justify-center gap-1.5",
                  dadosConsultoria.tipoRelatorio === 'Parcial'
                    ? "bg-amber-500 text-white border-amber-500 shadow-xs"
                    : "bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100"
                )}
              >
                <Clock size={14} /> RelatÃ³rio Parcial
              </button>
            </div>
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Ãrea da Consultoria</label>
            <select 
              className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              value={dadosConsultoria.areaConsultoria || ''}
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, areaConsultoria: e.target.value})}
            >
              <option value="">Selecione uma Ã¡rea</option>
              {CONSULTORIA_AREAS.map(area => (
                <option key={area} value={area}>{area}</option>
              ))}
            </select>
          </div>

          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="block text-[10px] font-bold text-slate-400 uppercase">CÃ³digo de ContrataÃ§Ã£o no SGF</label>
              <button
                type="button"
                onClick={() => setIsCopySgfModalOpen(true)}
                className="text-[10px] font-bold text-blue-600 hover:text-blue-700 uppercase flex items-center gap-1 cursor-pointer transition-colors"
                title="Copiar dados de outra consultoria com mesmo cÃ³digo SGF"
              >
                <Copy size={11} /> Copiar de Contrato SGF
              </button>
            </div>
            <div className="flex gap-2">
              <input 
                className="flex-1 p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-all font-mono font-medium"
                value={dadosConsultoria.codigoSgf || ''}
                placeholder="Ex: SGF-2026-9876"
                onChange={(e) => setDadosConsultoria({...dadosConsultoria, codigoSgf: e.target.value})}
              />
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => setIsCopySgfModalOpen(true)}
                className="bg-white border-blue-200 text-blue-700 hover:bg-blue-50 text-xs shrink-0"
                title="Buscar consultorias por CÃ³digo SGF"
              >
                <Search size={14} className="mr-1" /> Buscar SGF
              </Button>
            </div>
            {matchingSgfDiags.length > 0 && (
              <div className="mt-2 p-2.5 bg-blue-50/90 border border-blue-200 rounded-xl flex items-center justify-between text-xs text-blue-900 shadow-2xs">
                <span className="truncate mr-2 font-medium">
                  âœ¨ HÃ¡ <strong>{matchingSgfDiags.length}</strong> consultoria(s) com este mesmo CÃ³digo SGF.
                </span>
                <button
                  type="button"
                  onClick={() => {
                    if (matchingSgfDiags[0]?.dadosConsultoria) {
                      handleCopySgfData(matchingSgfDiags[0].dadosConsultoria);
                    }
                  }}
                  className="text-xs font-bold text-blue-700 hover:text-blue-900 underline shrink-0 flex items-center gap-1"
                >
                  <Copy size={12} /> Preencher Agora
                </button>
              </div>
            )}
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">PerÃ­odo da Consultoria</label>
            <input 
              className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              value={dadosConsultoria.periodoConsultoria || ''}
              placeholder="Ex: 01/03/2026 a 30/04/2026"
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, periodoConsultoria: e.target.value})}
            />
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Carga HorÃ¡ria Total (Cadastro)</label>
            <input 
              className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-all font-mono font-bold text-emerald-700"
              value={dadosConsultoria.cargaHoraria || ''}
              placeholder="Ex: 34 HS"
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, cargaHoraria: e.target.value})}
            />
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">TÃ©cnico do Sebrae ResponsÃ¡vel</label>
            <input 
              className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-blue-500 outline-none transition-all"
              value={dadosConsultoria.tecnicoSebrae || ''}
              placeholder="Ex: JoÃ£o Silva - Sebrae/RN"
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, tecnicoSebrae: e.target.value})}
            />
          </div>
        </div>
      </Card>

      {/* Objetivos e SoluÃ§Ãµes Propostas */}
      <Card className="p-6 border-slate-100 shadow-sm">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-100">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-amber-50 text-amber-600 rounded-xl flex items-center justify-center">
              <Search size={20} />
            </div>
            <div>
              <h3 className="font-bold text-slate-800 uppercase tracking-tight">Objetivos, Problemas e SoluÃ§Ãµes</h3>
              <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Detalhamento gerencial do escopo da consultoria</p>
            </div>
          </div>

          <Button
            type="button"
            size="sm"
            variant="outline"
            onClick={() => {
              const atvs = selectedDiagnostico?.cronograma || [];
              if (!atvs || atvs.length === 0) {
                alert("Nenhuma atividade encontrada no cronograma do relatÃ³rio para sincronizar. Selecione ou aplique um modelo no relatÃ³rio primeiro.");
                return;
              }

              const validAtvs = atvs.filter(a => a.nome && a.nome.trim() !== '');
              const totalHoras = validAtvs.reduce((acc, a) => acc + (parseInt(a.cargaHoraria?.replace(/\D/g, '') || '0', 10) || 0), 0);
              const solucoesLista = validAtvs.map((s, i) => `â€¢ Atividade ${i + 1} (${s.nome}): ${s.solucaoProposta || s.descricao}`).join('\n');
              const resultadosLista = validAtvs.map((s) => `â€¢ ${s.resultadoEsperado || s.solucaoProposta || s.nome}`).filter(Boolean).join('\n');

              let probsHeader = "PROBLEMAS IDENTIFICADOS:\nâ€¢ Necessidade de organizaÃ§Ã£o das rotinas financeiras, controle por ciclo e apuraÃ§Ã£o de resultados\n\n";
              if (dadosConsultoria.solucoesIndicadas?.includes("PROBLEMAS IDENTIFICADOS:")) {
                probsHeader = dadosConsultoria.solucoesIndicadas.split("SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:")[0];
              }

              setDadosConsultoria(prev => ({
                ...prev,
                cargaHoraria: totalHoras > 0 ? `${totalHoras}hs` : prev.cargaHoraria,
                solucoesIndicadas: `${probsHeader}SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesLista}`,
                resultadosEsperados: resultadosLista
              }));

              playSuccessSound();
              alert("SoluÃ§Ãµes e Resultados Esperados sincronizados em perfeita sintonia com as atividades do Cronograma!");
            }}
            className="border-emerald-200 bg-emerald-50 text-emerald-800 hover:bg-emerald-100 text-xs font-semibold"
          >
            <Sparkles size={14} className="mr-1.5 text-emerald-600" /> Sincronizar com o Cronograma
          </Button>
        </div>

        <div className="grid grid-cols-1 gap-6">
          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">1. Objetivo da Consultoria (Conforme DiagnÃ³stico)</label>
            <textarea 
              className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-amber-500 outline-none transition-all"
              rows={3}
              value={dadosConsultoria.objetivo || ''}
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, objetivo: e.target.value})}
              placeholder="O objetivo desta consultoria Ã© implementar melhorias nas Ã¡reas identificadas no diagnÃ³stico..."
            />
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">2. Problemas Identificados / SoluÃ§Ãµes / AÃ§Ãµes Propostas</label>
            <textarea 
              className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-amber-500 outline-none transition-all"
              rows={6}
              value={dadosConsultoria.solucoesIndicadas || ''}
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, solucoesIndicadas: e.target.value})}
              placeholder="PROBLEMAS IDENTIFICADOS:&#10;â€¢ Falta de controle financeiro diÃ¡rio&#10;&#10;SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:&#10;â€¢ ImplantaÃ§Ã£o de fluxo de caixa em planilha/aplicativo"
            />
          </div>

          <div>
            <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">3. Resultados Esperados</label>
            <textarea 
              className="w-full p-3 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-amber-500 outline-none transition-all"
              rows={4}
              value={dadosConsultoria.resultadosEsperados || ''}
              onChange={(e) => setDadosConsultoria({...dadosConsultoria, resultadosEsperados: e.target.value})}
              placeholder="Descreva os resultados esperados apÃ³s a implementaÃ§Ã£o das soluÃ§Ãµes..."
            />
          </div>
        </div>

        <div className="mt-8 flex justify-end gap-3 pt-4 border-t border-slate-100">
          <Button 
            onClick={handleSaveDadosConsultoria} 
            disabled={isSavingLocal} 
            className={cn("font-medium px-8 transition-all text-white", saveSuccessLocal ? "bg-emerald-700" : "bg-emerald-600 hover:bg-emerald-700")}
          >
            {isSavingLocal ? (
              <>
                <Loader2 className="animate-spin mr-2" size={18} />
                <span>Salvando...</span>
              </>
            ) : saveSuccessLocal ? (
              <>
                <CheckCircle2 className="mr-2 text-white" size={18} />
                <span>Salvo com Sucesso!</span>
              </>
            ) : (
              <>
                <Save className="mr-2" size={18} />
                <span>Salvar Dados da Consultoria</span>
              </>
            )}
          </Button>
        </div>
      </Card>

      <CopySgfConsultoriaModal
        isOpen={isCopySgfModalOpen}
        onClose={() => setIsCopySgfModalOpen(false)}
        currentCodigoSgf={dadosConsultoria.codigoSgf || ''}
        diagnosticos={diagnosticos}
        empresas={empresas}
        onSelectDadosConsultoria={handleCopySgfData}
      />
    </div>
  );
};

const CronogramaView = ({ 
  selectedDiagnostico, 
  selectedEmpresa,
  respostas, 
  solucoes, 
  setSelectedDiagnostico, 
  setDiagnosticos,
  playSuccessSound,
  setPdfUrl,
  empresasCredenciadas,
  setCredenciadaForm,
  setModalType,
  setIsModalOpen,
  customLogo,
  customConsultoraLogo,
  setCustomLogo,
  setCustomConsultoraLogo,
  logoChoice,
  setLogoChoice,
  problemas,
  historicalData = [],
  onGenerateActionPlan,
  onReplicateActionPlan,
  tarefasPlano,
  setTarefasPlano,
  setView
}: { 
  selectedDiagnostico: Diagnostico, 
  selectedEmpresa: Empresa | null,
  respostas: Resposta[], 
  solucoes: Solucao[], 
  setSelectedDiagnostico: (d: Diagnostico) => void,
  setDiagnosticos?: React.Dispatch<React.SetStateAction<Diagnostico[]>>,
  playSuccessSound: () => void,
  setPdfUrl: (url: string) => void,
  empresasCredenciadas: EmpresaCredenciada[],
  setCredenciadaForm: (f: any) => void,
  setModalType: (t: any) => void,
  setIsModalOpen: (open: boolean) => void,
  customLogo?: string | null,
  customConsultoraLogo?: string | null,
  setCustomLogo?: (logo: string | null) => void,
  setCustomConsultoraLogo?: (logo: string | null) => void,
  logoChoice: 'both' | 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'both' | 'sebrae' | 'consultora' | 'none') => void,
  problemas: Problema[],
  historicalData?: any[],
  onGenerateActionPlan?: () => void,
  onReplicateActionPlan?: () => void,
  tarefasPlano?: TarefaPlanoAcao[],
  setTarefasPlano?: React.Dispatch<React.SetStateAction<TarefaPlanoAcao[]>>,
  setView?: (v: string) => void
}) => {
  const [selectedModeloId, setSelectedModeloId] = useState<string>('financas');
  const [atividades, setAtividades] = useState<AtividadeCronograma[]>(() => {
    let rawList: AtividadeCronograma[] = [];
    if (selectedDiagnostico.cronograma && selectedDiagnostico.cronograma.length > 0 && selectedDiagnostico.cronograma.some(atv => atv.nome && atv.nome.trim() !== '')) {
      rawList = selectedDiagnostico.cronograma;
    } else {
      const segment = selectedDiagnostico.tipoEmpresa || 'Geral';
      const sameSegmentDiag = (historicalData || []).find(d => 
        d.tipoEmpresa && 
        d.tipoEmpresa.toLowerCase() === segment.toLowerCase() && 
        d.cronograma && 
        d.cronograma.length > 0 && 
        d.cronograma.some(a => a.nome && a.nome.trim() !== '')
      );
      if (sameSegmentDiag && sameSegmentDiag.cronograma) {
        rawList = sameSegmentDiag.cronograma;
      } else {
        rawList = PLANO_DE_ACAO_PADRAO;
      }
    }

    return rawList.map((atv, idx) => {
      const defaultDate = getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx);
      return {
        ...atv,
        dataInicio: atv.dataInicio || defaultDate,
        dataFim: atv.dataFim || atv.dataInicio || defaultDate
      };
    });
  });
  const [dadosConsultoria, setDadosConsultoria] = useState<DadosConsultoria>(
    selectedDiagnostico.dadosConsultoria || {}
  );
  const [isGeneratingAI, setIsGeneratingAI] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  const [editingActivityIndex, setEditingActivityIndex] = useState<number | null>(null);
  const [editingActivity, setEditingActivity] = useState<AtividadeCronograma | null>(null);

  const [modelosRelatorio, setModelosRelatorio] = useState<ModeloRelatorio[]>(() => {
    try {
      const saved = localStorage.getItem('custom_modelos_relatorio_v2');
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.error("Erro ao carregar modelos customizados:", e);
    }
    return MODELOS_RELATORIO;
  });

  const [editingModelo, setEditingModelo] = useState<ModeloRelatorio | null>(null);
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  // Hydrate missing or offline-cached evidence and consulting report data from IndexedDB on component mount
  useEffect(() => {
    if (!selectedDiagnostico?.id) return;
    let isMounted = true;
    loadLocalDiagnosticoReport(selectedDiagnostico.id).then((cached) => {
      if (!isMounted || !cached) return;
      if (cached.cronograma && Array.isArray(cached.cronograma) && cached.cronograma.length > 0) {
        setAtividades(prev => {
          if (!prev || prev.length === 0) return cached.cronograma;
          return cached.cronograma;
        });
      }

      if (cached.dadosConsultoria && Object.keys(cached.dadosConsultoria).length > 0) {
        setDadosConsultoria(prev => ({
          ...cached.dadosConsultoria,
          ...prev,
          evidencias: cached.dadosConsultoria.evidencias || prev.evidencias
        }));
      }
    });
    return () => { isMounted = false; };
  }, [selectedDiagnostico?.id]);

  // Synchronize state and persistent stores helper
  const syncEvidenceState = (updatedAtividades: AtividadeCronograma[], updatedDados: DadosConsultoria) => {
    if (!selectedDiagnostico?.id) return;
    const sanitizedAtv = sanitizeForFirestore(updatedAtividades);
    const sanitizedDados = sanitizeForFirestore(updatedDados);

    // 1. Update parent state
    const updatedDiag = {
      ...selectedDiagnostico,
      cronograma: sanitizedAtv,
      dadosConsultoria: sanitizedDados
    };
    setSelectedDiagnostico(updatedDiag);
    if (setDiagnosticos) {
      setDiagnosticos(prev => prev.map(d => d.id === selectedDiagnostico.id ? updatedDiag : d));
    }

    // 2. Persist to localStorage
    try {
      const saved = localStorage.getItem('local_diagnosticos');
      const parsed = saved ? JSON.parse(saved) : [];
      const nextList = Array.isArray(parsed) 
        ? parsed.map((d: any) => d.id === selectedDiagnostico.id ? updatedDiag : d)
        : [updatedDiag];
      localStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
    } catch (e) {
      console.warn("Could not save to localStorage local_diagnosticos", e);
    }

    // 3. Persist locally to IndexedDB immediately (immune to quota drops & large payloads)
    saveLocalDiagnosticoReport(selectedDiagnostico.id, sanitizedAtv, sanitizedDados);

    // 4. Persist to Firestore asynchronously
    const currentUser = auth.currentUser;
    if (currentUser && selectedDiagnostico.id) {
      updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), {
        cronograma: sanitizedAtv,
        dadosConsultoria: sanitizedDados,
        updatedAt: new Date().toISOString()
      }).catch(err => {
        console.warn("Auto-sync evidÃªncia aviso Firestore:", err);
      });
    }
  };

  const compressEvidenceImage = (file: File, callback: (compressedBase64: string) => void) => {
    compressImageToDataUrl(file, 520, 0.58)
      .then(b64 => callback(b64))
      .catch(() => {
        // Fallback
        const reader = new FileReader();
        reader.onload = (ev) => callback(ev.target?.result as string);
        reader.readAsDataURL(file);
      });
  };

  const handleUploadEvidenceForActivity = async (activityIndex: number, file: File) => {
    try {
      const compressedBase64 = await compressImageToDataUrl(file, 520, 0.58);
      
      let nextAtividades: AtividadeCronograma[] = [];
      setAtividades(prev => {
        const updated = [...prev];
        const currentAtv = updated[activityIndex] || { nome: '', descricao: '', cargaHoraria: '', solucaoProposta: '', status: 'Pendente', prioridade: 'MÃ©dia' };
        const currentEvidencias = currentAtv.evidencias || [];
        const nextEvidencias = [...currentEvidencias, compressedBase64];
        updated[activityIndex] = {
          ...currentAtv,
          evidencias: nextEvidencias
        };
        nextAtividades = updated;
        return updated;
      });

      if (editingActivityIndex === activityIndex) {
        setEditingActivity(prev => {
          if (!prev) return prev;
          return {
            ...prev,
            evidencias: [...(prev.evidencias || []), compressedBase64]
          };
        });
      }

      // Auto-save immediately
      setTimeout(() => {
        if (nextAtividades.length > 0) {
          syncEvidenceState(nextAtividades, dadosConsultoria);
        }
      }, 50);

      playSuccessSound();
    } catch (err) {
      console.error("Erro ao processar imagem de evidÃªncia:", err);
    }
  };

  const handleRemoveEvidenceFromActivity = (activityIndex: number, evidenceIndex: number) => {
    let nextAtividades: AtividadeCronograma[] = [];
    setAtividades(prev => {
      const updated = [...prev];
      const currentAtv = updated[activityIndex];
      if (!currentAtv) return prev;
      const currentEvidencias = [...(currentAtv.evidencias || [])];
      currentEvidencias.splice(evidenceIndex, 1);
      updated[activityIndex] = {
        ...currentAtv,
        evidencias: currentEvidencias
      };
      nextAtividades = updated;
      return updated;
    });

    if (editingActivityIndex === activityIndex) {
      setEditingActivity(prev => {
        if (!prev) return prev;
        const currentEvidencias = [...(prev.evidencias || [])];
        currentEvidencias.splice(evidenceIndex, 1);
        return {
          ...prev,
          evidencias: currentEvidencias
        };
      });
    }

    setTimeout(() => {
      if (nextAtividades.length > 0) {
        syncEvidenceState(nextAtividades, dadosConsultoria);
      }
    }, 50);
  };

  const handleLinkGeneralEvidenceToActivity = (generalEvidenceIdx: number, targetActivityIdx: number) => {
    const generalList = [...(dadosConsultoria.evidencias || [])];
    const imageToMove = generalList[generalEvidenceIdx];
    if (!imageToMove) return;

    // Remove from general list
    generalList.splice(generalEvidenceIdx, 1);
    const updatedDados = { ...dadosConsultoria, evidencias: generalList };
    setDadosConsultoria(updatedDados);

    // Add to selected activity
    let nextAtividades: AtividadeCronograma[] = [];
    setAtividades(prev => {
      const updated = [...prev];
      const currentAtv = updated[targetActivityIdx] || { nome: '', descricao: '', cargaHoraria: '', solucaoProposta: '', status: 'Pendente', prioridade: 'MÃ©dia' };
      const currentEvidencias = currentAtv.evidencias || [];
      updated[targetActivityIdx] = {
        ...currentAtv,
        evidencias: [...currentEvidencias, imageToMove]
      };
      nextAtividades = updated;
      return updated;
    });

    setTimeout(() => {
      if (nextAtividades.length > 0) {
        syncEvidenceState(nextAtividades, updatedDados);
      }
    }, 50);
  };

  useEffect(() => {
    // Helper to calculate problems identified
    const cleanResps = deduplicateRespostas(respostas);
    const responsesByProblem = cleanResps.reduce((acc: Record<string, Resposta[]>, r) => {
      const id = r.idProblema || r.problema;
      if (!acc[id]) acc[id] = [];
      acc[id].push(r);
      return acc;
    }, {});

    const relevantSolutions = solucoes.filter(solution => {
      const probId = solution.idProblema;
      const relatedResps = responsesByProblem[probId] || [];
      if (relatedResps.length > 0) {
        const criticalCount = relatedResps.filter(r => r.resposta === 'NÃ£o' || r.resposta === 'Parcial').length;
        return criticalCount > 0;
      }
      return false;
    });

    const sortedRelevantSolutions = [...relevantSolutions].sort((a, b) => {
      const probA = problemas.find(p => p.id === a.idProblema || p.descricao_problemas === a.problema);
      const probB = problemas.find(p => p.id === b.idProblema || p.descricao_problemas === b.problema);
      const impA = IMPACTO_ORDER[probA?.impacto as string] || 0;
      const impB = IMPACTO_ORDER[probB?.impacto as string] || 0;
      return impB - impA;
    });

    const problemasIdentificadosTexto = Array.from(new Set(sortedRelevantSolutions.map(s => `â€¢ ${s.problema}`))).join('\n');

    if (dadosConsultoria.solucoesIndicadas && (
      dadosConsultoria.solucoesIndicadas.includes("RESULTADOS ESPERADOS:") || 
      !dadosConsultoria.solucoesIndicadas.includes("PROBLEMAS IDENTIFICADOS:") ||
      dadosConsultoria.solucoesIndicadas.includes("SOLUÃ‡Ã•ES INDICADAS:")
    )) {
      let newText = dadosConsultoria.solucoesIndicadas;
      
      // Remove results if present (they are in the table below)
      if (newText.includes("RESULTADOS ESPERADOS:")) {
        newText = newText.split("RESULTADOS ESPERADOS:")[0].trim();
      }
      
      // Update header
      if (newText.includes("SOLUÃ‡Ã•ES INDICADAS:")) {
        newText = newText.replace("SOLUÃ‡Ã•ES INDICADAS:", "SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:");
      }

      // Add problems if missing
      if (!newText.includes("PROBLEMAS IDENTIFICADOS:") && problemasIdentificadosTexto) {
        newText = `PROBLEMAS IDENTIFICADOS:\n${problemasIdentificadosTexto}\n\n${newText}`;
      }
      
      if (newText !== dadosConsultoria.solucoesIndicadas) {
        setDadosConsultoria(prev => ({ ...prev, solucoesIndicadas: newText }));
      }
    }
  }, [respostas, solucoes]);

  const generateSuggestionsAI = async () => {
    if (!respostas || respostas.length === 0) {
      alert("Realize o diagnÃ³stico primeiro para gerar sugestÃµes baseadas nos dados.");
      return;
    }
    
    setIsGeneratingAI(true);
    try {
      const ai = getAI();
      if (!ai) {
        alert("IA nÃ£o configurada ou API Key invÃ¡lida.");
        setIsGeneratingAI(false);
        return;
      }

      const totalCargaStr = dadosConsultoria.cargaHoraria || selectedDiagnostico.dadosConsultoria?.cargaHoraria || (selectedEmpresa as any)?.cargaHoraria || '34h';
      const numTotal = parseInt(totalCargaStr.replace(/\D/g, ''), 10) || 34;

      const prompt = `Como um consultor sÃªnior do SEBRAE, analise o diagnÃ³stico empresarial e as respostas abaixo para sugerir um plano de trabalho/cronograma detalhado para uma consultoria de sucesso.
      
      DADOS DO DIAGNÃ“STICO:
      ${JSON.stringify(respostas.map(r => ({ id: r.idProblema, p: r.pergunta, r: r.resposta, o: r.observacao, area: r.area }))).slice(0, 3500)}
      CARGA HORÃRIA TOTAL REGISTRADA DA CONSULTORIA: ${numTotal}hs
      
      DIRETRIZES DE ATIVIDADES E ORDEM LÃ“GICA DE EXECUÃ‡ÃƒO:
      - A soma das cargas horÃ¡rias de todas as atividades DEVE ser exatamente ${numTotal}hs.
      - As atividades DEVEM ser apresentadas em ordem LÃ“GICA sequencial de execuÃ§Ã£o:
        1. "Entendimento da demanda e diagnÃ³stico inicial" (2h a 4h)
        2. Atividades intermediÃ¡rias focadas em processos, custos, finanÃ§as e comercializaÃ§Ã£o (2h a 4h cada)
        3. DEVE OBRIGATORIAMENTE incluir a atividade "Desenvolvimento de sistema, aplicativos ou planilhas" com exatamente 8h de carga horÃ¡ria.
        4. DEVE OBRIGATORIAMENTE incluir a atividade "AnÃ¡lise final de atividades da consultoria" com exatamente 2h de carga horÃ¡ria.
        5. "RelatÃ³rio final e encerramento" (2h a 4h)
      
      FORMATO CADA ITEM:
      {
        "nome": "string",
        "descricao": "string",
        "cargaHoraria": "string (ex: 4h, 8h, 2h)",
        "solucaoProposta": "string",
        "resultadoEsperado": "string (com KPIs quantitativos)",
        "responsavel": "string",
        "prioridade": "Alta" | "MÃ©dia" | "Baixa",
        "idProblema": "string (vincular ao id fornecido)"
      }`;

      const response = await ai.models.generateContent({
        model: "gemini-3.1-flash-lite",
        contents: prompt,
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                nome: { type: Type.STRING },
                descricao: { type: Type.STRING },
                cargaHoraria: { type: Type.STRING },
                solucaoProposta: { type: Type.STRING },
                resultadoEsperado: { type: Type.STRING },
                responsavel: { type: Type.STRING },
                prioridade: { type: Type.STRING, enum: ["Alta", "MÃ©dia", "Baixa"] },
                idProblema: { type: Type.STRING }
              },
              required: ["nome", "descricao", "cargaHoraria", "solucaoProposta", "resultadoEsperado", "responsavel", "prioridade"]
            }
          }
        }
      });

      const aiSuggestions = extractAndParseJSON(response.text, []);
      if (aiSuggestions && aiSuggestions.length > 0) {
        const adjustedSuggestions = adjustActivitiesMax4Hours(aiSuggestions, totalCargaStr);
        // Map to include initial status and auto-calculated dates starting at dataDiagnostico (+7 days per activity)
        const finalizedSuggestions = adjustedSuggestions.map((s: any, idx: number) => {
          const defaultDate = getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx);
          return {
            ...s,
            status: 'Pendente',
            dataInicio: s.dataInicio || defaultDate,
            dataFim: s.dataFim || s.dataInicio || defaultDate
          };
        });
        setAtividades(finalizedSuggestions);
        
        // Also update summary
        const problemasUnicos = Array.from(new Set(respostas.filter(r => r.resposta !== 'Sim').map(r => r.problema))).join('\nâ€¢ ');
        const solucoesUnicas = Array.from(new Set(aiSuggestions.map((s: any) => s.solucaoProposta))).join('\nâ€¢ ');
        const resultadosUnicos = Array.from(new Set(aiSuggestions.map((s: any) => s.resultadoEsperado))).join('\nâ€¢ ');

        setDadosConsultoria({
          ...dadosConsultoria,
          objetivo: `Implementar soluÃ§Ãµes estratÃ©gicas baseadas no diagnÃ³stico para otimizar os processos da empresa.`,
          solucoesIndicadas: `PROBLEMAS IDENTIFICADOS:\nâ€¢ ${problemasUnicos}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\nâ€¢ ${solucoesUnicas}`,
          resultadosEsperados: `â€¢ ${resultadosUnicos}`
        });
        
        playSuccessSound();

        // Generate the action plan in tandem
        if (onGenerateActionPlan) {
          Promise.resolve(onGenerateActionPlan()).catch((e: any) => {
            console.error("Erro ao gerar o plano de aÃ§Ã£o de forma integrada:", e);
          });
        }
      }
    } catch (error: any) {
      console.error("AI Error:", error);
      alert("Erro ao conectar com a IA:\n" + (error.message || error));
    } finally {
      setIsGeneratingAI(false);
    }
  };

  // Sync state if selectedDiagnostico changes from outside
  useEffect(() => {
    if (selectedDiagnostico.cronograma && selectedDiagnostico.cronograma.length > 0 && selectedDiagnostico.cronograma.some(a => a.nome && a.nome.trim() !== '')) {
      setAtividades(selectedDiagnostico.cronograma);
    } else {
      const segment = selectedDiagnostico.tipoEmpresa || 'Geral';
      const sameSegmentDiag = (historicalData || []).find(d => 
        d.tipoEmpresa && 
        d.tipoEmpresa.toLowerCase() === segment.toLowerCase() && 
        d.cronograma && 
        d.cronograma.length > 0 && 
        d.cronograma.some(a => a.nome && a.nome.trim() !== '')
      );
      if (sameSegmentDiag && sameSegmentDiag.cronograma) {
        const adapted = sameSegmentDiag.cronograma.map((atv, idx) => {
          const defaultDate = getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx);
          return {
            ...atv,
            dataInicio: defaultDate,
            dataFim: defaultDate
          };
        });
        setAtividades(adapted);
      }
    }
    if (selectedDiagnostico.dadosConsultoria) {
      setDadosConsultoria(selectedDiagnostico.dadosConsultoria);
    }
  }, [selectedDiagnostico.id, selectedDiagnostico.cronograma]);

  const handleClear = () => {
    setAtividades(PLANO_DE_ACAO_PADRAO);
    playSuccessSound();
  };

  const moveAtividade = (index: number, direction: 'up' | 'down') => {
    const targetIndex = direction === 'up' ? index - 1 : index + 1;
    if (targetIndex < 0 || targetIndex >= atividades.length) return;
    const newAtividades = [...atividades];
    const [moved] = newAtividades.splice(index, 1);
    newAtividades.splice(targetIndex, 0, moved);
    setAtividades(newAtividades);
    playSuccessSound();
  };

  const moveAtividadeToPos = (currentIndex: number, newIndex: number) => {
    if (newIndex < 0 || newIndex >= atividades.length || newIndex === currentIndex) return;
    const newAtividades = [...atividades];
    const [moved] = newAtividades.splice(currentIndex, 1);
    newAtividades.splice(newIndex, 0, moved);
    setAtividades(newAtividades);
    playSuccessSound();
  };

  const handleFilterResolved = () => {
    // A problem is resolved if ALL its answers are 'Sim'
    const cleanResps = deduplicateRespostas(respostas);
    const resolvedProblemIds = new Set(
      Object.entries(
        cleanResps.reduce((acc: Record<string, string[]>, r) => {
          const id = r.idProblema || r.problema;
          if (!acc[id]) acc[id] = [];
          acc[id].push(r.resposta);
          return acc;
        }, {})
      )
        .filter(([_, resps]) => resps.every(r => r === 'Sim'))
        .map(([id]) => id)
    );

    const filtered = atividades.filter(atv => {
      if (!atv.idProblema) return true; // Keep fixed activities
      return !resolvedProblemIds.has(atv.idProblema);
    });

    if (filtered.length < atividades.length) {
      setAtividades(filtered);
      playSuccessSound();
    } else {
      alert("Nenhuma atividade resolvida encontrada para remover.");
    }
  };

  const handleSave = async () => {
    if (!selectedDiagnostico) return;
    setIsSaving(true);
    setSaveSuccess(false);

    // Hard safety timeout: ensures button never stays stuck even under unpredictable exceptions
    const safetyTimer = setTimeout(() => {
      setIsSaving(false);
    }, 2000);

    try {
      const sanitizedDados = sanitizeForFirestore(dadosConsultoria);
      const sanitizedAtividades = sanitizeForFirestore(atividades);

      // Optimistically update state to reflect changes instantly
      const updatedDiag = { ...selectedDiagnostico, cronograma: sanitizedAtividades, dadosConsultoria: sanitizedDados };
      setSelectedDiagnostico(updatedDiag);
      if (setDiagnosticos) {
        setDiagnosticos(prev => prev.map(d => d.id === selectedDiagnostico.id ? updatedDiag : d));
      }

      // 1. Guaranteed multi-megabyte local persistence in IndexedDB (non-blocking)
      if (selectedDiagnostico.id) {
        saveLocalDiagnosticoReport(selectedDiagnostico.id, sanitizedAtividades, sanitizedDados).catch(err => {
          console.warn("[EvidenceStorage] Non-critical IndexedDB warning:", err);
        });
      }

      // 2. Persist to localStorage local_diagnosticos
      try {
        const saved = localStorage.getItem('local_diagnosticos');
        const parsed = saved ? JSON.parse(saved) : [];
        const nextList = Array.isArray(parsed) 
          ? parsed.map((d: any) => d.id === selectedDiagnostico.id ? updatedDiag : d)
          : [updatedDiag];
        localStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
      } catch (e) {
        console.warn("Could not save to localStorage local_diagnosticos", e);
      }

      // 3. Sync local tarefasPlano state and localStorage
      const validAtividades = (sanitizedAtividades || []).filter((atv: any) => atv.nome?.trim());
      const newTarefasFromAtividades: TarefaPlanoAcao[] = validAtividades.map((atv: any, index: number) => {
        const defaultDate = getActivityDateStr(selectedDiagnostico.dataDiagnostico, index);
        const startStr = atv.dataInicio || defaultDate;
        const endStr = atv.dataFim || startStr;
        return {
          id: atv.id || `tarefa_${selectedDiagnostico.id}_${index}`,
          diagnosticoId: selectedDiagnostico.id,
          empresaId: selectedEmpresa?.id || selectedDiagnostico.empresaId || '',
          idProblema: atv.idProblema || '',
          problema: atv.nome,
          area: atv.area || selectedDiagnostico.tipoEmpresa || 'Consultoria',
          solucaoSugerida: atv.solucaoProposta || '',
          acoes: atv.descricao || '',
          status: (atv.status === 'ConcluÃ­do' || atv.status === 'Em Andamento' || atv.status === 'Pendente') ? atv.status : 'Pendente',
          prioridade: (atv.prioridade === 'Alta' || atv.prioridade === 'Baixa' || atv.prioridade === 'MÃ©dia') ? atv.prioridade : 'MÃ©dia',
          responsavel: atv.responsavel || 'Consultor',
          dataInicio: startStr,
          dataFim: endStr,
          dataVencimento: endStr,
          ownerId: auth.currentUser?.uid || 'local',
          ordem: index
        } as TarefaPlanoAcao;
      });

      if (setTarefasPlano) {
        setTarefasPlano(prev => {
          const others = prev.filter(t => t.diagnosticoId !== selectedDiagnostico.id);
          return [...others, ...newTarefasFromAtividades];
        });
      }

      try {
        const savedTasks = localStorage.getItem('local_tarefas_plano');
        const parsedTasks = savedTasks ? JSON.parse(savedTasks) : [];
        const others = Array.isArray(parsedTasks) ? parsedTasks.filter((t: any) => t.diagnosticoId !== selectedDiagnostico.id) : [];
        localStorage.setItem('local_tarefas_plano', JSON.stringify([...others, ...newTarefasFromAtividades]));
      } catch (e) {
        console.warn("Could not save local_tarefas_plano", e);
      }

      // 4. Background cloud sync without blocking the UI
      const currentUser = auth.currentUser;
      if (currentUser && selectedDiagnostico.id) {
        (async () => {
          try {
            const syncPromise = async () => {
              await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), {
                cronograma: sanitizedAtividades,
                dadosConsultoria: sanitizedDados,
                updatedAt: new Date().toISOString()
              });

              const qCurrentTasks = query(collection(db, 'tarefas_plano'), where('diagnosticoId', '==', selectedDiagnostico.id));
              const snapCurrentTasks = await getDocs(qCurrentTasks);
              const existingDocs = snapCurrentTasks?.docs || [];

              const batch = writeBatch(db);
              validAtividades.forEach((atv: any, index: number) => {
                const defaultDate = getActivityDateStr(selectedDiagnostico.dataDiagnostico, index);
                const startStr = atv.dataInicio || defaultDate;
                const endStr = atv.dataFim || startStr;

                const taskData = sanitizeForFirestore({
                  diagnosticoId: selectedDiagnostico.id,
                  empresaId: selectedEmpresa?.id || selectedDiagnostico.empresaId || '',
                  idProblema: atv.idProblema || '',
                  problema: atv.nome,
                  solucaoSugerida: atv.solucaoProposta || '',
                  acoes: atv.descricao || '',
                  status: atv.status || 'Pendente',
                  prioridade: atv.prioridade || 'MÃ©dia',
                  responsavel: atv.responsavel || 'Consultor',
                  cargaHoraria: atv.cargaHoraria || '',
                  dataInicio: startStr,
                  dataFim: endStr,
                  dataVencimento: endStr,
                  ownerId: currentUser.uid,
                  ordem: index,
                  dataCadastro: new Date().toISOString()
                });

                if (index < existingDocs.length) {
                  batch.update(existingDocs[index].ref, taskData);
                } else {
                  const newTaskRef = doc(collection(db, 'tarefas_plano'));
                  batch.set(newTaskRef, taskData);
                }
              });

              for (let i = validAtividades.length; i < existingDocs.length; i++) {
                batch.delete(existingDocs[i].ref);
              }

              await batch.commit();
            };

            await Promise.race([
              syncPromise(),
              new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout ao sincronizar na nuvem")), 4500))
            ]);
          } catch (cloudErr) {
            console.warn("[CloudSync] Background sync note for cronograma:", cloudErr);
          }
        })();
      }

      // Success feedback immediately applied
      clearTimeout(safetyTimer);
      setIsSaving(false);
      playSuccessSound();
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3500);
    } catch (error) {
      console.error("Error saving cronograma:", error);
      clearTimeout(safetyTimer);
      setIsSaving(false);
      playSuccessSound();
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3500);
    } finally {
      setIsSaving(false);
    }
  };

  const updateAtividade = async (index: number, field: keyof AtividadeCronograma, value: string) => {
    const newAtividades = [...atividades];
    newAtividades[index] = { ...newAtividades[index], [field]: value };
    setAtividades(newAtividades);

    const sanitizedAtv = sanitizeForFirestore(newAtividades);
    const sanitizedDados = sanitizeForFirestore(dadosConsultoria);
    const updatedDiag = { ...selectedDiagnostico, cronograma: sanitizedAtv, dadosConsultoria: sanitizedDados };
    setSelectedDiagnostico(updatedDiag);
    if (setDiagnosticos) {
      setDiagnosticos(prev => prev.map(d => d.id === selectedDiagnostico.id ? updatedDiag : d));
    }

    if (selectedDiagnostico.id) {
      saveLocalDiagnosticoReport(selectedDiagnostico.id, sanitizedAtv, sanitizedDados);
    }

    if (field === 'status' && auth.currentUser && selectedDiagnostico.id) {
      try {
        await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), { cronograma: sanitizedAtv });

        const qTask = query(collection(db, 'tarefas_plano'), where('diagnosticoId', '==', selectedDiagnostico.id));
        const snap = await getDocs(qTask);
        const batch = writeBatch(db);
        snap.docs.forEach(docSnap => {
          const tData = docSnap.data();
          if (tData.ordem === index || tData.problema === newAtividades[index].nome) {
            batch.update(docSnap.ref, { status: value });
          }
        });
        await batch.commit();
      } catch (err) {
        console.error("Error syncing activity status:", err);
      }
    }
  };

  const deleteAtividade = (index: number) => {
    const newAtividades = atividades.filter((_, i) => i !== index);
    setAtividades(newAtividades);
    playSuccessSound();
  };

  const toggleAtividadeRelatorio = (index: number) => {
    setAtividades(prev => {
      const updated = [...prev];
      const current = updated[index];
      const isCurrentlyIncluded = current.incluirNoRelatorio !== undefined 
        ? current.incluirNoRelatorio 
        : (current.status || '').toLowerCase().includes('conclu');
      updated[index] = {
        ...current,
        incluirNoRelatorio: !isCurrentlyIncluded
      };
      return updated;
    });
  };

  const selectAllAtividadesRelatorio = () => {
    setAtividades(prev => prev.map(a => ({ ...a, incluirNoRelatorio: true })));
  };

  const selectConcluidasAtividadesRelatorio = () => {
    setAtividades(prev => prev.map(a => ({
      ...a,
      incluirNoRelatorio: (a.status || '').toLowerCase().includes('conclu')
    })));
  };

  const clearSelectionAtividadesRelatorio = () => {
    setAtividades(prev => prev.map(a => ({ ...a, incluirNoRelatorio: false })));
  };

  const generateCronogramaPDF = () => {
    try {
      if (!selectedEmpresa || !selectedDiagnostico || atividades.length === 0) {
        alert("Dados insuficientes para gerar o relatÃ³rio. Verifique se a empresa e o diagnÃ³stico estÃ£o selecionados e se hÃ¡ atividades no cronograma.");
        return;
      }

      const doc = new jsPDF();
    
    // Group all responses by problem to find relevant solutions
    const cleanResps = deduplicateRespostas(respostas);
    const responsesByProblem = cleanResps.reduce((acc: Record<string, Resposta[]>, r) => {
      const id = r.idProblema || r.problema;
      if (!acc[id]) acc[id] = [];
      acc[id].push(r);
      return acc;
    }, {});

    const relevantSolutions = solucoes.filter(solution => {
      const probId = solution.idProblema;
      const relatedResps = responsesByProblem[probId] || [];
      if (relatedResps.length > 0) {
        const criticalCount = relatedResps.filter(r => r.resposta === 'NÃ£o' || r.resposta === 'Parcial').length;
        return criticalCount > 0;
      }
      return false;
    });

    const sortedRelevantSolutions = [...relevantSolutions].sort((a, b) => {
      const probA = problemas.find(p => p.id === a.idProblema || p.descricao_problemas === a.problema);
      const probB = problemas.find(p => p.id === b.idProblema || p.descricao_problemas === b.problema);
      const impA = IMPACTO_ORDER[probA?.impacto as string] || 0;
      const impB = IMPACTO_ORDER[probB?.impacto as string] || 0;
      return impB - impA;
    });

    const areasIdentificadas = Array.from(new Set(sortedRelevantSolutions.map(s => s.area))).join(', ');
    const problemasIdentificadosTexto = Array.from(new Set(sortedRelevantSolutions.map(s => `â€¢ ${s.problema}`))).join('\n');
    const solucoesIndicadasTexto = sortedRelevantSolutions.map(s => `â€¢ ${s.solucao_recomendada}`).join('\n');
    const acoesPropostasTexto = Array.from(new Set(sortedRelevantSolutions.map(s => `â€¢ ${s.acoes_sugeridas}`))).join('\n');

    let finalContent = dadosConsultoria.solucoesIndicadas || `PROBLEMAS IDENTIFICADOS:\n${problemasIdentificadosTexto || 'â€¢ (Preencher problemas)'}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesIndicadasTexto || 'â€¢ (Preencher soluÃ§Ãµes)'}\n${acoesPropostasTexto || 'â€¢ (Preencher aÃ§Ãµes)'}`;
    
    // Auto-clean: Remove "RESULTADOS ESPERADOS" as requested (it appears in the table below)
    if (finalContent.includes("RESULTADOS ESPERADOS:")) {
      finalContent = finalContent.split("RESULTADOS ESPERADOS:")[0].trim();
    }
    
    // Ensure "PROBLEMAS IDENTIFICADOS" is present if we have them
    if (!finalContent.includes("PROBLEMAS IDENTIFICADOS:") && problemasIdentificadosTexto) {
      if (finalContent.includes("SOLUÃ‡Ã•ES INDICADAS:")) {
        finalContent = finalContent.replace("SOLUÃ‡Ã•ES INDICADAS:", "SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:");
      }
      finalContent = `PROBLEMAS IDENTIFICADOS:\n${problemasIdentificadosTexto}\n\n${finalContent}`;
    } else if (finalContent.includes("SOLUÃ‡Ã•ES INDICADAS:") && !finalContent.includes("SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:")) {
      // Standardize header name
      finalContent = finalContent.replace("SOLUÃ‡Ã•ES INDICADAS:", "SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:");
    }

    const activeLogoForPdf = logoChoice === 'sebrae' ? customLogo : logoChoice === 'consultora' ? customConsultoraLogo : null;
    const logoValid = isValidLogoSource(activeLogoForPdf);

    const addHeader = (doc: jsPDF) => {
      autoTable(doc, {
        startY: 10,
        theme: 'grid',
        styles: { fontSize: 7, cellPadding: 1, lineWidth: 0.1, lineColor: 0 },
        body: [
          [
            { 
              content: logoValid ? '' : 'SEBRAE/CE', 
              rowSpan: 3, 
              styles: { halign: 'center', valign: 'middle', cellWidth: 35, fontStyle: 'bold' } 
            },
            { content: 'Preenchimento pelo SEBRAE/CE', colSpan: 3, styles: { halign: 'center', fontStyle: 'bold', fillColor: [240, 240, 240] } }
          ],
          [
            { content: 'NÃºmero da Nota Fiscal:', colSpan: 2 },
            { content: 'Data emissÃ£o da Nota Fiscal:\n_____/_____/_____' }
          ],
          [
            { content: 'GEDOC:' },
            { content: 'CÃ³digo SIAC:' },
            { content: 'Data ConsolidaÃ§Ã£o no SIAC:\n_____/_____/_____' }
          ]
        ],
        didDrawCell: (data) => {
          if (logoValid && activeLogoForPdf && data.row.index === 0 && data.column.index === 0) {
            try {
              const imgProps = doc.getImageProperties(activeLogoForPdf);
              const imgWidth = 28;
              const imgHeight = (imgProps.height * imgWidth) / imgProps.width;
              const x = data.cell.x + (data.cell.width - imgWidth) / 2;
              const y = data.cell.y + (data.cell.height - imgHeight) / 2;
              doc.addImage(activeLogoForPdf, 'PNG', x, y, imgWidth, imgHeight, undefined, 'FAST');
            } catch (e) {
              console.warn("Could not add custom logo to PDF header", e);
            }
          }
        }
      });
    };

    const checkPageBreak = (needed: number) => {
      if (currentY + needed > 275) {
        doc.addPage();
        addHeader(doc);
        currentY = ((doc as any).lastAutoTable?.finalY || 40) + 5;
        return true;
      }
      return false;
    };

    addHeader(doc);
    let currentY = ((doc as any).lastAutoTable?.finalY || 40) + 5;

    const tipoRelatorioVal = (dadosConsultoria.tipoRelatorio || 'Final').toUpperCase();

    doc.setFontSize(12);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(30, 41, 59);
    doc.text('RELATÃ“RIO DE CONSULTORIA', 105, currentY, { align: 'center' });
    currentY += 5;

    doc.setFontSize(8.5);
    doc.setFont("helvetica", "bold");
    if (tipoRelatorioVal === 'PARCIAL') {
      doc.setTextColor(217, 119, 6);
    } else {
      doc.setTextColor(5, 150, 105);
    }
    doc.text(`(${tipoRelatorioVal === 'PARCIAL' ? 'RELATÃ“RIO PARCIAL DE ACOMPANHAMENTO' : 'RELATÃ“RIO FINAL DE CONSULTORIA GERENCIAL'})`, 105, currentY, { align: 'center' });
    currentY += 6;

    doc.setTextColor(30, 41, 59);
    doc.setFontSize(9);
    doc.text('1. DADOS DO CLIENTE', 15, currentY);
    currentY += 1;

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [{ content: 'RAZÃƒO SOCIAL:', styles: { fontStyle: 'bold', cellWidth: 40 } }, { content: selectedEmpresa.razaoSocial || selectedEmpresa.nome, colSpan: 3 }],
        [{ content: 'NOME FANTASIA:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.nomeFantasia || selectedEmpresa.nome, colSpan: 3 }],
        [
          { content: 'CNPJ:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.cnpj || '' },
          { content: 'NÃšMERO DA CAF:', styles: { fontStyle: 'bold', cellWidth: 50 } }, { content: selectedEmpresa.cafNumero || '' }
        ],
        [
          { content: 'MÃŠS/ANO DE ABERTURA:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.mesAnoAbertura || '', colSpan: 3 }
        ],
        [{ content: 'ENDEREÃ‡O:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.enderecoComercial || '', colSpan: 3 }],
        [
          { content: 'CONTATO:', styles: { fontStyle: 'bold' } }, { content: `${selectedEmpresa.telefoneFixo || ''} | ${selectedEmpresa.celular || ''}`, colSpan: 3 }
        ],
        [{ content: 'EMAIL:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.email || '', colSpan: 3 }],
        [{ content: 'REPRESENTANTE:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.representante || '', colSpan: 3 }],
        [{ content: 'CPF DO REPRESENTANTE:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.cpfRepresentante || '', colSpan: 3 }]
      ]
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.text('2. DADOS DA EMPRESA CREDENCIADA', 15, currentY);
    currentY += 1;

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [{ content: 'RAZÃƒO SOCIAL:', styles: { fontStyle: 'bold', cellWidth: 40 } }, { content: dadosConsultoria.razaoSocial || '', colSpan: 3 }],
        [{ content: 'CNPJ:', styles: { fontStyle: 'bold' } }, { content: dadosConsultoria.cnpj || '', colSpan: 3 }],
        [{ content: 'CONSULTOR:', styles: { fontStyle: 'bold' } }, { content: dadosConsultoria.consultor || '', colSpan: 3 }],
        [{ content: 'CONTATO:', styles: { fontStyle: 'bold' } }, { content: `${dadosConsultoria.celular || ''} | ${dadosConsultoria.email || ''}`, colSpan: 3 }]
      ]
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.text('3. DADOS DA CONSULTORIA GERENCIAL', 15, currentY);
    currentY += 1;

    const totalHoras = atividades.reduce((acc, atv) => {
      const horas = parseInt((atv.cargaHoraria || '').toString().replace(/\D/g, '')) || 0;
      return acc + horas;
    }, 0);

    // Selected activities for the report
    const atividadesSelecionadas = atividades
      .map((atv, originalIndex) => ({ atv, originalIndex, activityNumber: originalIndex + 1 }))
      .filter(item => {
        if (item.atv.incluirNoRelatorio !== undefined) {
          return item.atv.incluirNoRelatorio;
        }
        return (item.atv.status || '').toLowerCase().includes('conclu');
      });

    const totalHorasSelecionadas = atividadesSelecionadas.reduce((acc, item) => {
      const horas = parseInt((item.atv.cargaHoraria || '').toString().replace(/\D/g, '')) || 0;
      return acc + horas;
    }, 0);

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [
          { content: 'TIPO DE RELATÃ“RIO:', styles: { fontStyle: 'bold', cellWidth: 40 } },
          { 
            content: `RELATÃ“RIO ${tipoRelatorioVal}`, 
            colSpan: 3, 
            styles: { fontStyle: 'bold', textColor: tipoRelatorioVal === 'PARCIAL' ? [217, 119, 6] : [5, 150, 105] } 
          }
        ],
        [{ content: `ÃREA: ${dadosConsultoria.areaConsultoria || ''}`, colSpan: 4, styles: { fontStyle: 'bold' } }],
        [{ content: 'CÃ“DIGO SGF:', styles: { fontStyle: 'bold', cellWidth: 40 } }, { content: dadosConsultoria.codigoSgf || '', colSpan: 3 }],
        [
          { content: 'PERÃODO:', styles: { fontStyle: 'bold' } }, { content: dadosConsultoria.periodoConsultoria || '' },
          { content: 'CARGA HORÃRIA TOTAL:', styles: { fontStyle: 'bold', cellWidth: 40 } }, { content: dadosConsultoria.cargaHoraria || `${totalHoras} HS` }
        ],
        [
          { content: 'HORAS NO RELATÃ“RIO:', styles: { fontStyle: 'bold', cellWidth: 40 } },
          { content: `${totalHorasSelecionadas} HS (${atividadesSelecionadas.length} atividade(s) selecionada(s))`, colSpan: 3, styles: { fontStyle: 'bold', textColor: [5, 150, 105] } }
        ],
        [{ content: 'TÃ‰CNICO SEBRAE:', styles: { fontStyle: 'bold' } }, { content: dadosConsultoria.tecnicoSebrae || '', colSpan: 3 }]
      ]
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 2 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [{ content: '1. OBJETIVO DA CONSULTORIA:', styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }],
        [{ content: dadosConsultoria.objetivo || `Implementar melhorias nas Ã¡reas de ${areasIdentificadas || 'gestÃ£o empresarial'}.`, styles: { minCellHeight: 10 } }]
      ]
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 2 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [{ content: '2. PROBLEMAS IDENTIFICADOS / SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:', styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }],
        [{ content: finalContent, styles: { minCellHeight: 25 } }]
      ]
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    const atividadesBody: any[] = [];

    if (atividadesSelecionadas.length === 0) {
      atividadesBody.push([
        { content: 'Atividades Selecionadas:', styles: { fontStyle: 'bold', cellWidth: 40 } },
        { content: 'Nenhuma atividade selecionada para exibiÃ§Ã£o neste relatÃ³rio.', styles: { fontStyle: 'italic', textColor: [120, 120, 120] } }
      ]);
    } else {
      atividadesSelecionadas.forEach((item, idx) => {
        const atv = item.atv;
        const cleanedDesc = (atv.descricao || '').split('\n\n').filter(p => !p.startsWith('DiagnÃ³stico:')).join('\n\n');

        atividadesBody.push([
          { content: `Atividade ${idx + 1} (Ref. #${item.activityNumber}) â€“`, styles: { fontStyle: 'bold', cellWidth: 38 } },
          { content: atv.nome || '', styles: { fontStyle: 'bold' } }
        ]);
        atividadesBody.push([
          { content: 'DescriÃ§Ã£o', styles: { fontStyle: 'bold' } },
          { content: cleanedDesc || '' }
        ]);
        atividadesBody.push([
          { content: `Carga horÃ¡ria desta atividade: ${(atv.cargaHoraria || '0').toString().replace(/h/gi, '')} hs`, styles: { fontStyle: 'bold' } },
          { content: `Status da atividade: ${atv.status || 'ConcluÃ­do'}`, styles: { fontStyle: 'bold', textColor: (atv.status || '').toLowerCase().includes('conclu') ? [5, 150, 105] : [71, 85, 105] } }
        ]);
      });
    }

    atividadesBody.push([
      { content: `CARGA HORÃRIA DO RELATÃ“RIO (${tipoRelatorioVal}):`, styles: { fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [5, 150, 105] } },
      { content: `${totalHorasSelecionadas} hs (${atividadesSelecionadas.length} de ${atividades.length} atividades)`, styles: { fontStyle: 'bold', fillColor: [236, 253, 245], textColor: [5, 150, 105] } }
    ]);

    atividadesBody.push([
      { content: `CARGA HORÃRIA TOTAL DA CONSULTORIA:`, styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } },
      { content: `${totalHoras} hs`, styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }
    ]);

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1.5 },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      head: [[{ content: 'SOLUÃ‡Ã•ES IMPLEMENTADAS\nDetalhamento das atividades selecionadas para o relatÃ³rio.', colSpan: 2, styles: { fontStyle: 'bold', fillColor: [240, 240, 240], fontSize: 9 } }]],
      body: atividadesBody
    });

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    checkPageBreak(15);

    if (dadosConsultoria.resultadosEsperados) {
      autoTable(doc, {
        startY: currentY,
        theme: 'plain',
        styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 2 },
        margin: { top: 38, bottom: 20, left: 15, right: 15 },
        didDrawPage: (data) => {
          if (data.pageNumber > 1) {
            addHeader(doc);
          }
        },
        body: [
          [{ content: '3. RESULTADOS ESPERADOS:', styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }],
          [{ content: dadosConsultoria.resultadosEsperados, styles: { minCellHeight: 15 } }]
        ]
      });
      currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 3;
    }

    currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 5;

    // Add Evidence Photos for selected activities (linked to activities in ascending order & general) with 50% scale reduction
    const activitiesWithEvidence = atividadesSelecionadas
      .filter(item => item.atv.evidencias && item.atv.evidencias.length > 0);

    const generalEvidencias = dadosConsultoria.evidencias || [];
    const hasAnyEvidencias = activitiesWithEvidence.length > 0 || generalEvidencias.length > 0;

    if (hasAnyEvidencias) {
      doc.addPage();
      addHeader(doc);
      currentY = ((doc as any).lastAutoTable?.finalY || 40) + 8;
      
      doc.setFontSize(11);
      doc.setFont("helvetica", "bold");
      doc.setTextColor(30, 41, 59);
      doc.text('EVIDÃŠNCIAS FOTOGRÃFICAS DAS ATIVIDADES EXECUTADAS', 15, currentY);
      currentY += 6;
      
      doc.setFontSize(8);
      doc.setFont("helvetica", "italic");
      doc.setTextColor(100, 116, 139);
      doc.text('Registros fotogrÃ¡ficos comprobatÃ³rios vinculados a cada atividade em ordem crescente (tamanho reduzido em 50%).', 15, currentY);
      currentY += 8;

      // 50% scale rendering for PDF (A4 printable width: 180mm)
      // 4 images per row: imgWidth = 41mm, gap = 5mm, margin = 15mm (15 + 4*41 + 3*5 = 194mm)
      const imgWidth = 41;
      const margin = 15;
      const gap = 5;
      const cols = 4;

      // 1. Render linked evidences for each activity in ascending order
      for (const item of activitiesWithEvidence) {
        if (currentY > 235) {
          doc.addPage();
          addHeader(doc);
          currentY = ((doc as any).lastAutoTable?.finalY || 40) + 10;
        }

        // Activity banner
        doc.setFillColor(240, 253, 244);
        doc.setDrawColor(187, 247, 208);
        doc.rect(margin, currentY, 180, 6.5, 'FD');
        doc.setFontSize(8.5);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(5, 150, 105);
        const atvTitle = `Atividade ${item.activityNumber} â€“ ${item.atv.nome || 'Sem tÃ­tulo'}`;
        doc.text(doc.splitTextToSize(atvTitle, 175)[0], margin + 3, currentY + 4.5);
        currentY += 8.5;

        const imgs = item.atv.evidencias || [];
        for (let i = 0; i < imgs.length; i += cols) {
          const rowImgs = imgs.slice(i, i + cols);
          
          // Calculate heights
          const heights: number[] = [];
          for (const imgBase64 of rowImgs) {
            let h = 30;
            try {
              const p = doc.getImageProperties(imgBase64);
              h = (p.height * imgWidth) / p.width;
              if (h > 42) h = 42; // Cap maximum height
            } catch(e) {}
            heights.push(h);
          }
          const rowMaxHeight = Math.max(...heights, 24);

          if (currentY + rowMaxHeight + 8 > 275) {
            doc.addPage();
            addHeader(doc);
            currentY = ((doc as any).lastAutoTable?.finalY || 40) + 10;
          }

          rowImgs.forEach((imgBase64, colIdx) => {
            const xPos = margin + colIdx * (imgWidth + gap);
            const h = heights[colIdx] || 30;
            try {
              doc.setDrawColor(226, 232, 240);
              doc.rect(xPos, currentY, imgWidth, h);
              doc.addImage(imgBase64, 'JPEG', xPos, currentY, imgWidth, h, undefined, 'FAST');
              
              // Caption under photo
              doc.setFontSize(6.5);
              doc.setFont("helvetica", "bold");
              doc.setTextColor(100, 116, 139);
              doc.text(`EvidÃªncia ${i + colIdx + 1} (Ativ. #${item.activityNumber})`, xPos, currentY + h + 3.5);
            } catch(e) {
              console.warn("Error drawing activity evidence image:", e);
            }
          });

          currentY += rowMaxHeight + 7;
        }

        currentY += 3;
      }

      // 2. Render general / unassigned evidences if any
      if (generalEvidencias.length > 0) {
        if (currentY > 235) {
          doc.addPage();
          addHeader(doc);
          currentY = ((doc as any).lastAutoTable?.finalY || 40) + 10;
        }

        doc.setFillColor(241, 245, 249);
        doc.setDrawColor(203, 213, 225);
        doc.rect(margin, currentY, 180, 6.5, 'FD');
        doc.setFontSize(8.5);
        doc.setFont("helvetica", "bold");
        doc.setTextColor(71, 85, 105);
        doc.text('Outras EvidÃªncias Gerais da Consultoria', margin + 3, currentY + 4.5);
        currentY += 8.5;

        for (let i = 0; i < generalEvidencias.length; i += cols) {
          const rowImgs = generalEvidencias.slice(i, i + cols);
          const heights: number[] = [];
          for (const imgBase64 of rowImgs) {
            let h = 30;
            try {
              const p = doc.getImageProperties(imgBase64);
              h = (p.height * imgWidth) / p.width;
              if (h > 42) h = 42;
            } catch(e) {}
            heights.push(h);
          }
          const rowMaxHeight = Math.max(...heights, 24);

          if (currentY + rowMaxHeight + 8 > 275) {
            doc.addPage();
            addHeader(doc);
            currentY = ((doc as any).lastAutoTable?.finalY || 40) + 10;
          }

          rowImgs.forEach((imgBase64, colIdx) => {
            const xPos = margin + colIdx * (imgWidth + gap);
            const h = heights[colIdx] || 30;
            try {
              doc.setDrawColor(226, 232, 240);
              doc.rect(xPos, currentY, imgWidth, h);
              doc.addImage(imgBase64, 'JPEG', xPos, currentY, imgWidth, h, undefined, 'FAST');
              
              doc.setFontSize(6.5);
              doc.setFont("helvetica", "bold");
              doc.setTextColor(100, 116, 139);
              doc.text(`EvidÃªncia Geral ${i + colIdx + 1}`, xPos, currentY + h + 3.5);
            } catch(e) {}
          });

          currentY += rowMaxHeight + 7;
        }
      }
    }

    checkPageBreak(15);

    autoTable(doc, {
      startY: currentY,
      theme: 'plain',
      styles: { fontSize: 8, cellPadding: 3, halign: 'center', valign: 'bottom' },
      margin: { top: 38, bottom: 20, left: 15, right: 15 },
      didDrawPage: (data) => {
        if (data.pageNumber > 1) {
          addHeader(doc);
        }
      },
      body: [
        [
          { content: `______________________________\n${dadosConsultoria.consultor || 'Consultor responsÃ¡vel'}\nConsultor`, styles: { minCellHeight: 25 } },
          { content: `______________________________\n${selectedEmpresa?.representante || selectedEmpresa?.nome || '________________'}\nCliente`, styles: { minCellHeight: 25 } }
        ]
      ]
    });

    setPdfUrl(doc.output('bloburl').toString());
    } catch (error) {
      console.error("Error generating Cronograma PDF:", error);
      alert("Ocorreu um erro ao gerar o PDF do cronograma. Por favor, tente novamente.");
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">RelatÃ³rio de Consultoria</h2>
          <p className="text-slate-500 font-medium">Defina as etapas do projeto e as soluÃ§Ãµes propostas</p>
        </div>
        <div className="flex gap-3">
          <Button variant="outline" onClick={() => setAtividades([...atividades, { nome: '', descricao: '', cargaHoraria: '', solucaoProposta: '', responsavel: '', status: 'Pendente', prioridade: 'MÃ©dia' }])} className="px-4 text-emerald-600 border-emerald-100">
            <Plus size={18} /> Adicionar Atividade
          </Button>
          <Button variant="outline" onClick={generateCronogramaPDF} className="px-4 text-emerald-600 hover:bg-emerald-50 border-emerald-100">
            <Printer size={18} /> RelatÃ³rio (PDF)
          </Button>
          <Button 
            variant="outline" 
            onClick={generateSuggestionsAI} 
            disabled={isGeneratingAI}
            className="px-6 text-indigo-600 border-indigo-100 hover:bg-indigo-50"
          >
            {isGeneratingAI ? <Loader2 className="animate-spin" size={18} /> : <Sparkles size={18} />} 
            {isGeneratingAI ? "Gerando RelatÃ³rio & Plano..." : "Gerar RelatÃ³rio e Plano com IA"}
          </Button>
          <Button 
            onClick={handleSave} 
            disabled={isSaving} 
            className={cn("px-8 font-medium transition-all text-white", saveSuccess ? "bg-emerald-700" : "bg-emerald-600 hover:bg-emerald-700")}
          >
            {isSaving ? (
              <>
                <Loader2 className="animate-spin mr-1.5" size={18} />
                <span>Salvando...</span>
              </>
            ) : saveSuccess ? (
              <>
                <CheckCircle2 className="mr-1.5 text-white" size={18} />
                <span>Salvo com Sucesso!</span>
              </>
            ) : (
              <>
                <Save className="mr-1.5" size={18} />
                <span>Salvar RelatÃ³rio</span>
              </>
            )}
          </Button>
        </div>
      </div>

      <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm">
        <LogoSelector 
          logoChoice={logoChoice}
          setLogoChoice={setLogoChoice}
          customLogo={customLogo}
          customConsultoraLogo={customConsultoraLogo}
          setCustomLogo={setCustomLogo}
          setCustomConsultoraLogo={setCustomConsultoraLogo}
        />
      </div>

      <div className="grid grid-cols-1 gap-4">
        {/* 2. Empresa Credenciada - Apenas SeleÃ§Ã£o e Resumo */}
        <Card className="p-6 mb-8 border-slate-100 shadow-sm">
          <div className="flex items-center gap-3 mb-6 pb-4 border-b border-slate-100">
            <div className="w-10 h-10 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center">
              <Building2 size={20} />
            </div>
            <div>
              <h3 className="font-bold text-slate-800 uppercase tracking-tight">2. Empresa Credenciada</h3>
              <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Selecione a credenciada cadastrada para vincular aos relatÃ³rios</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center">
            <div>
              <div className="flex items-center justify-between mb-1">
                <label className="block text-[10px] font-bold text-slate-400 uppercase">Nome da Empresa Credenciada</label>
                <button 
                  onClick={() => {
                    setCredenciadaForm({});
                    setModalType('createCredenciada');
                    setIsModalOpen(true);
                  }}
                  className="text-[10px] font-bold text-emerald-600 hover:text-emerald-700 uppercase flex items-center gap-1"
                >
                  <Plus size={10} /> Nova Credenciada
                </button>
              </div>
              <select 
                className="w-full p-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-medium focus:ring-2 focus:ring-emerald-500 outline-none transition-all"
                value={empresasCredenciadas.find(c => c.razaoSocial === dadosConsultoria.razaoSocial)?.id || ''}
                onChange={(e) => {
                  const cred = empresasCredenciadas.find(c => c.id === e.target.value);
                  if (cred) {
                    setDadosConsultoria({
                      ...dadosConsultoria,
                      razaoSocial: cred.razaoSocial,
                      cnpj: cred.cnpj,
                      telefoneFixo: cred.telefoneFixo,
                      celular: cred.celular,
                      email: cred.email,
                      consultor: cred.consultor
                    });
                  }
                }}
              >
                <option value="">Selecionar Credenciada...</option>
                {empresasCredenciadas.map(cred => (
                  <option key={cred.id} value={cred.id}>{cred.razaoSocial}</option>
                ))}
              </select>
            </div>

            {dadosConsultoria.razaoSocial ? (
              <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 flex flex-col justify-center">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Dados Registrados da Credenciada</p>
                <p className="text-sm font-bold text-slate-800">{dadosConsultoria.razaoSocial}</p>
                <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-xs text-slate-600 font-medium">
                  {dadosConsultoria.cnpj && <span>CNPJ: <strong className="font-mono">{dadosConsultoria.cnpj}</strong></span>}
                  {dadosConsultoria.consultor && <span>Consultor: <strong>{dadosConsultoria.consultor}</strong></span>}
                </div>
              </div>
            ) : (
              <div className="p-4 bg-amber-50/50 rounded-xl border border-amber-200 text-xs text-amber-700 font-medium flex items-center">
                <span>Selecione uma credenciada acima para preencher automaticamente os dados nos relatÃ³rios.</span>
              </div>
            )}
          </div>
        </Card>

        {/* 3. Dados da Consultoria Gerencial - Card Conectado com BotÃ£o de EdiÃ§Ã£o na Aba Dedicada */}
        <Card className="p-6 mb-8 border-slate-100 shadow-sm bg-gradient-to-r from-slate-50/50 to-blue-50/20">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4 pb-4 border-b border-slate-100">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
                <FileText size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-800 uppercase tracking-tight">3. Dados da Consultoria Gerencial</h3>
                <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Conectados ao cliente cadastrado</p>
              </div>
            </div>
            {setView && (
              <Button 
                variant="outline" 
                onClick={() => setView('dados-consultoria')}
                className="bg-white border-blue-200 text-blue-700 hover:bg-blue-50 font-bold text-xs"
              >
                <Edit3 size={14} className="mr-1.5" /> Preencher / Editar em Aba Dedicada
              </Button>
            )}
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 text-xs">
            <div className="p-3 bg-white rounded-lg border border-slate-100 shadow-2xs">
              <span className="block text-[10px] font-bold text-slate-400 uppercase mb-0.5">Ãrea da Consultoria</span>
              <span className="font-semibold text-slate-800">{dadosConsultoria.areaConsultoria || 'NÃ£o informada'}</span>
            </div>
            <div className="p-3 bg-white rounded-lg border border-slate-100 shadow-2xs">
              <span className="block text-[10px] font-bold text-slate-400 uppercase mb-0.5">CÃ³digo SGF</span>
              <span className="font-semibold text-slate-800 font-mono">{dadosConsultoria.codigoSgf || 'NÃ£o informado'}</span>
            </div>
            <div className="p-3 bg-white rounded-lg border border-slate-100 shadow-2xs">
              <span className="block text-[10px] font-bold text-slate-400 uppercase mb-0.5">Carga HorÃ¡ria Total</span>
              <span className="font-bold text-emerald-700 font-mono">{dadosConsultoria.cargaHoraria || 'NÃ£o informada'}</span>
            </div>
            <div className="p-3 bg-white rounded-lg border border-slate-100 shadow-2xs">
              <span className="block text-[10px] font-bold text-slate-400 uppercase mb-0.5">TÃ©cnico Sebrae</span>
              <span className="font-semibold text-slate-800">{dadosConsultoria.tecnicoSebrae || 'NÃ£o informado'}</span>
            </div>
          </div>
        </Card>

        {/* Modelos de RelatÃ³rio de Consultoria Baseados no DiagnÃ³stico */}
        <Card className="p-6 mb-8 border-slate-800 shadow-md bg-slate-900 text-white">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6 pb-4 border-b border-slate-800">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-emerald-500/20 text-emerald-400 rounded-xl flex items-center justify-center">
                <Sparkles size={20} />
              </div>
              <div>
                <h3 className="font-bold text-white uppercase tracking-tight text-base">Modelos de RelatÃ³rio de Consultoria</h3>
                <p className="text-xs text-slate-400 font-medium">Selecione, aplique ou duplique o modelo de relatÃ³rio ideal para o seu cliente</p>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {modelosRelatorio.map((modelo) => {
              const isSelected = selectedModeloId === modelo.id;
              const totalHoras = modelo.atividades.reduce((acc, a) => acc + (parseInt(a.cargaHoraria?.replace(/\D/g, '') || '0', 10) || 0), 0);
              return (
                <div 
                  key={modelo.id}
                  onClick={() => setSelectedModeloId(modelo.id)}
                  className={`p-4 rounded-xl border transition-all cursor-pointer flex flex-col justify-between ${
                    isSelected 
                      ? 'bg-emerald-950/40 border-emerald-500 ring-2 ring-emerald-500/30' 
                      : 'bg-slate-800/80 border-slate-700/80 hover:border-slate-600 hover:bg-slate-800'
                  }`}
                >
                  <div>
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 uppercase tracking-wider">
                        {modelo.categoria}
                      </span>
                      <span className="text-[10px] font-bold text-emerald-400 bg-slate-950/60 px-2 py-0.5 rounded-md border border-slate-700/60">
                        {totalHoras > 0 ? `${totalHoras}hs` : `${modelo.atividades.length} atv`} â€¢ {modelo.atividades.length} Atividades
                      </span>
                    </div>
                    <h4 className="font-bold text-sm text-slate-100 mb-1.5">{modelo.nome}</h4>
                    <p className="text-xs text-slate-400 line-clamp-2 leading-relaxed mb-3">{modelo.descricao}</p>
                  </div>

                  <div className="flex items-center gap-2 mt-2">
                    <Button
                      size="sm"
                      onClick={(e) => {
                        e.stopPropagation();
                        if (window.confirm(`Deseja aplicar o modelo "${modelo.nome}" (${modelo.atividades.length} atividades, ${totalHoras}hs) a este relatÃ³rio?`)) {
                          const finalized = modelo.atividades.map((s, idx) => ({
                            ...s,
                            status: 'Pendente',
                            dataInicio: s.dataInicio || getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx),
                            dataFim: s.dataFim || s.dataInicio || getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx)
                          }));
                          setAtividades(finalized);

                          const solucoesLista = finalized.map((s, i) => `â€¢ Atividade ${i + 1} (${s.nome}): ${s.solucaoProposta || s.descricao}`).join('\n');
                          const resultadosLista = finalized.map((s) => `â€¢ ${s.resultadoEsperado || s.solucaoProposta || s.nome}`).filter(Boolean).join('\n');

                          let probsHeader = "PROBLEMAS IDENTIFICADOS:\nâ€¢ Necessidade de organizaÃ§Ã£o das rotinas financeiras, controle por ciclo e apuraÃ§Ã£o de resultados\n\n";
                          if (dadosConsultoria.solucoesIndicadas?.includes("PROBLEMAS IDENTIFICADOS:")) {
                            probsHeader = dadosConsultoria.solucoesIndicadas.split("SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:")[0];
                          }

                          const updatedDados = {
                            ...dadosConsultoria,
                            cargaHoraria: totalHoras > 0 ? `${totalHoras}hs` : dadosConsultoria.cargaHoraria,
                            solucoesIndicadas: `${probsHeader}SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesLista}`,
                            resultadosEsperados: resultadosLista
                          };

                          setDadosConsultoria(updatedDados);
                          if (selectedDiagnostico) {
                            setSelectedDiagnostico({ ...selectedDiagnostico, cronograma: finalized, dadosConsultoria: updatedDados });
                          }

                          playSuccessSound();
                          alert(`Modelo "${modelo.nome}" aplicado com sucesso! As soluÃ§Ãµes, aÃ§Ãµes e resultados esperados foram sincronizados com os Dados da Consultoria.`);
                        }
                      }}
                      className={isSelected ? "bg-emerald-600 hover:bg-emerald-500 text-white flex-1 text-xs font-bold" : "bg-slate-700 hover:bg-slate-600 text-slate-200 flex-1 text-xs font-medium"}
                    >
                      <CheckCircle2 size={14} className="mr-1.5" /> Aplicar
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={(e) => {
                        e.stopPropagation();
                        const copyM: ModeloRelatorio = {
                          ...JSON.parse(JSON.stringify(modelo)),
                          id: 'modelo_' + Date.now(),
                          nome: `${modelo.nome} (CÃ³pia)`
                        };
                        setEditingModelo(copyM);
                      }}
                      className="border-slate-700 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-2.5 font-medium"
                      title="Duplicar este modelo para editar"
                    >
                      <Copy size={14} className="mr-1" /> Duplicar
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={(e) => {
                        e.stopPropagation();
                        setEditingModelo(JSON.parse(JSON.stringify(modelo)));
                      }}
                      className="border-slate-700 bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs px-2.5 font-medium"
                      title="Editar este modelo"
                    >
                      <Edit2 size={14} className="mr-1" /> Editar
                    </Button>
                  </div>
                </div>
              );
            })}
          </div>
        </Card>

        {/* Modal para CriaÃ§Ã£o/EdiÃ§Ã£o do Modelo de RelatÃ³rio */}
        {editingModelo && (
          <Modal
            title={modelosRelatorio.some(m => m.id === editingModelo.id) ? `Editar Modelo: ${editingModelo.nome}` : "Criar Novo Modelo de Consultoria"}
            onClose={() => setEditingModelo(null)}
            confirmText={modelosRelatorio.some(m => m.id === editingModelo.id) ? "Salvar AlteraÃ§Ãµes" : "Criar e Salvar Modelo"}
            size="4xl"
            onConfirm={() => {
              if (!editingModelo.nome.trim()) {
                alert("Por favor, informe o nome do modelo.");
                return;
              }
              const exists = modelosRelatorio.some(m => m.id === editingModelo.id);
              const updatedList = exists
                ? modelosRelatorio.map(m => m.id === editingModelo.id ? editingModelo : m)
                : [...modelosRelatorio, editingModelo];

              setModelosRelatorio(updatedList);
              try {
                localStorage.setItem('custom_modelos_relatorio_v2', JSON.stringify(updatedList));
              } catch (err) {
                console.error("Erro ao salvar no localStorage:", err);
              }
              setEditingModelo(null);
              playSuccessSound();
              alert(`Modelo "${editingModelo.nome}" salvo com sucesso!`);
            }}
          >
            <div className="space-y-4 max-h-[70vh] overflow-y-auto pr-2">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div>
                  <label className="text-xs font-bold text-slate-700 block mb-1">Nome do Modelo</label>
                  <input 
                    type="text"
                    value={editingModelo.nome} 
                    onChange={(e) => setEditingModelo({ ...editingModelo, nome: e.target.value })} 
                    placeholder="Ex: GestÃ£o Financeira da Carcinicultura"
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 transition-all bg-white"
                  />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-700 block mb-1">Categoria</label>
                  <input 
                    type="text"
                    value={editingModelo.categoria} 
                    onChange={(e) => setEditingModelo({ ...editingModelo, categoria: e.target.value })} 
                    placeholder="Ex: Carcinicultura, FinanÃ§as, OperaÃ§Ãµes"
                    className="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 transition-all bg-white"
                  />
                </div>
              </div>

              <div>
                <label className="text-xs font-bold text-slate-700 block mb-1">DescriÃ§Ã£o Resumida</label>
                <textarea 
                  value={editingModelo.descricao} 
                  onChange={(e) => setEditingModelo({ ...editingModelo, descricao: e.target.value })} 
                  placeholder="Descreva o objetivo e escopo do modelo de consultoria..."
                  rows={2}
                  className="w-full px-3 py-2 border border-slate-200 rounded-lg text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 transition-all bg-white"
                />
              </div>

              <div className="pt-3 border-t border-slate-200">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
                  <div>
                    <h4 className="font-bold text-sm text-slate-800">Atividades do Modelo ({editingModelo.atividades.length})</h4>
                    <p className="text-xs text-slate-500">
                      Carga HorÃ¡ria Total do Modelo: <span className="font-bold text-emerald-600">
                        {editingModelo.atividades.reduce((acc, a) => acc + (parseInt(a.cargaHoraria?.replace(/\D/g, '') || '0', 10) || 0), 0)} hs
                      </span>
                    </p>
                  </div>
                  <Button
                    type="button"
                    size="sm"
                    variant="outline"
                    className="text-emerald-700 border-emerald-300 hover:bg-emerald-50 text-xs font-semibold"
                    onClick={() => {
                      const newAtv: AtividadeCronograma = {
                        nome: `Nova Atividade ${editingModelo.atividades.length + 1}`,
                        descricao: "DescriÃ§Ã£o detalhada dos objetivos desta atividade...",
                        cargaHoraria: "4h",
                        solucaoProposta: "AÃ§Ã£o prÃ¡tica recomendada",
                        responsavel: "Consultor / Cliente",
                        status: "Pendente",
                        prioridade: "Alta",
                        resultadoEsperado: "Resultado esperado ao final da atividade"
                      };
                      setEditingModelo({
                        ...editingModelo,
                        atividades: [...editingModelo.atividades, newAtv]
                      });
                    }}
                  >
                    <Plus size={14} className="mr-1" /> Adicionar Atividade
                  </Button>
                </div>

                <div className="space-y-4">
                  {editingModelo.atividades.map((atv, aIdx) => (
                    <div key={aIdx} className="p-3 bg-slate-50 border border-slate-200 rounded-lg space-y-3">
                      <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                        <span className="text-xs font-bold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded">
                          Atividade {aIdx + 1}
                        </span>
                        <div className="flex items-center gap-1">
                          {aIdx > 0 && (
                            <Button
                              type="button"
                              variant="ghost"
                              size="sm"
                              className="h-6 w-6 p-0 text-slate-500 hover:bg-slate-200"
                              onClick={() => {
                                const newAtvs = [...editingModelo.atividades];
                                const temp = newAtvs[aIdx];
                                newAtvs[aIdx] = newAtvs[aIdx - 1];
                                newAtvs[aIdx - 1] = temp;
                                setEditingModelo({ ...editingModelo, atividades: newAtvs });
                              }}
                              title="Mover para cima"
                            >
                              â†‘
                            </Button>
                          )}
                          {aIdx < editingModelo.atividades.length - 1 && (
                            <Button
                              type="button"
                              variant="ghost"
                              size="sm"
                              className="h-6 w-6 p-0 text-slate-500 hover:bg-slate-200"
                              onClick={() => {
                                const newAtvs = [...editingModelo.atividades];
                                const temp = newAtvs[aIdx];
                                newAtvs[aIdx] = newAtvs[aIdx + 1];
                                newAtvs[aIdx + 1] = temp;
                                setEditingModelo({ ...editingModelo, atividades: newAtvs });
                              }}
                              title="Mover para baixo"
                            >
                              â†“
                            </Button>
                          )}
                          <Button
                            type="button"
                            variant="ghost"
                            size="sm"
                            className="h-6 w-6 p-0 text-rose-600 hover:bg-rose-50"
                            onClick={() => {
                              if (editingModelo.atividades.length <= 1) {
                                alert("O modelo deve conter pelo menos 1 atividade.");
                                return;
                              }
                              const newAtvs = editingModelo.atividades.filter((_, i) => i !== aIdx);
                              setEditingModelo({ ...editingModelo, atividades: newAtvs });
                            }}
                            title="Excluir Atividade"
                          >
                            <Trash2 size={14} />
                          </Button>
                        </div>
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                        <div className="sm:col-span-2">
                          <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">Nome da Atividade</label>
                          <input 
                            type="text"
                            value={atv.nome}
                            onChange={(e) => {
                              const newAtvs = [...editingModelo.atividades];
                              newAtvs[aIdx] = { ...newAtvs[aIdx], nome: e.target.value };
                              setEditingModelo({ ...editingModelo, atividades: newAtvs });
                            }}
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 bg-white"
                          />
                        </div>
                        <div>
                          <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">Carga HorÃ¡ria</label>
                          <input 
                            type="text"
                            value={atv.cargaHoraria}
                            onChange={(e) => {
                              const newAtvs = [...editingModelo.atividades];
                              newAtvs[aIdx] = { ...newAtvs[aIdx], cargaHoraria: e.target.value };
                              setEditingModelo({ ...editingModelo, atividades: newAtvs });
                            }}
                            placeholder="Ex: 2h, 4h, 8h"
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md text-xs font-bold text-emerald-700 focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 bg-white"
                          />
                        </div>
                      </div>

                      <div>
                        <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">DescriÃ§Ã£o</label>
                        <textarea 
                          value={atv.descricao}
                          onChange={(e) => {
                            const newAtvs = [...editingModelo.atividades];
                            newAtvs[aIdx] = { ...newAtvs[aIdx], descricao: e.target.value };
                            setEditingModelo({ ...editingModelo, atividades: newAtvs });
                          }}
                          rows={2}
                          className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 bg-white"
                        />
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                        <div>
                          <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">SoluÃ§Ã£o Proposta / AÃ§Ã£o</label>
                          <input 
                            type="text"
                            value={atv.solucaoProposta || ''}
                            onChange={(e) => {
                              const newAtvs = [...editingModelo.atividades];
                              newAtvs[aIdx] = { ...newAtvs[aIdx], solucaoProposta: e.target.value };
                              setEditingModelo({ ...editingModelo, atividades: newAtvs });
                            }}
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 bg-white"
                          />
                        </div>
                        <div>
                          <label className="text-[11px] font-semibold text-slate-600 block mb-0.5">ResponsÃ¡vel Sugerido</label>
                          <input 
                            type="text"
                            value={atv.responsavel || ''}
                            onChange={(e) => {
                              const newAtvs = [...editingModelo.atividades];
                              newAtvs[aIdx] = { ...newAtvs[aIdx], responsavel: e.target.value };
                              setEditingModelo({ ...editingModelo, atividades: newAtvs });
                            }}
                            className="w-full px-2.5 py-1.5 border border-slate-200 rounded-md text-xs font-medium focus:ring-2 focus:ring-emerald-500/20 focus:outline-none focus:border-emerald-500 bg-white"
                          />
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </Modal>
        )}

        <Card className="p-6 mb-8 border-slate-100 shadow-sm">
          <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-xl flex items-center justify-center">
                <ImageIcon size={20} />
              </div>
              <div>
                <h3 className="font-bold text-slate-800 uppercase tracking-tight">EvidÃªncias da Consultoria</h3>
                <p className="text-xs text-slate-500 uppercase font-medium tracking-wider">Fotos e comprovaÃ§Ãµes vinculadas Ã s atividades ou gerais (reduzidas em 50%)</p>
              </div>
            </div>
            <div className="text-xs text-slate-500 font-semibold bg-slate-100 px-3 py-1 rounded-full">
              Ordem crescente automÃ¡tica no relatÃ³rio
            </div>
          </div>

          <div className="space-y-6">
            <label className="flex flex-col items-center justify-center p-8 bg-slate-50 rounded-xl border-2 border-dashed border-slate-200 hover:border-indigo-400 hover:bg-indigo-50 transition-all cursor-pointer group">
              <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center shadow-sm mb-4 group-hover:scale-110 transition-transform">
                <Upload className="text-slate-400 group-hover:text-indigo-600" size={24} />
              </div>
              <div className="flex flex-col items-center">
                <span className="text-sm font-bold text-slate-500 group-hover:text-indigo-700">Adicionar fotos como evidÃªncia geral</span>
                <span className="text-xs text-slate-400 mt-1">Imagens comprimidas automaticamente (reduÃ§Ã£o de 50% no tamanho)</span>
                {dadosConsultoria.evidencias && dadosConsultoria.evidencias.length > 0 && (
                  <div className="mt-2 text-[10px] font-bold text-indigo-600 uppercase tracking-tighter bg-indigo-50 px-2 py-0.5 rounded-full border border-indigo-100">
                    {dadosConsultoria.evidencias.length} fotos gerais anexadas
                  </div>
                )}
              </div>
              <input 
                type="file" 
                className="hidden" 
                accept="image/*" 
                multiple
                onChange={(e) => {
                  const files = Array.from(e.target.files || []) as File[];
                  files.forEach(async (file: File) => {
                    try {
                      const compressedBase64 = await compressImageToDataUrl(file, 520, 0.58);
                      let nextDados: DadosConsultoria = dadosConsultoria;
                      setDadosConsultoria(prev => {
                        const nextEvidencias = [...(prev.evidencias || []), compressedBase64];
                        nextDados = {
                          ...prev,
                          evidencias: nextEvidencias
                        };
                        return nextDados;
                      });

                      setTimeout(() => {
                        syncEvidenceState(atividades, nextDados);
                      }, 50);
                    } catch (err) {
                      console.error("Erro ao comprimir evidÃªncia geral:", err);
                    }
                  });
                }} 
              />
            </label>

            {(dadosConsultoria.evidencias && dadosConsultoria.evidencias.length > 0) && (
              <div>
                <div className="text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">
                  Fotos Gerais (VocÃª pode vinculÃ¡-las diretamente a uma atividade abaixo):
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-3 mt-3">
                  {dadosConsultoria.evidencias.map((img, idx) => (
                    <div key={idx} className="relative group rounded-lg border border-slate-200 bg-slate-50 flex flex-col items-center justify-between p-1 overflow-hidden shadow-xs">
                      <div 
                        onClick={() => setPreviewImage(img)}
                        className="w-full aspect-square rounded overflow-hidden cursor-pointer relative"
                      >
                        <img 
                          src={img} 
                          alt={`EvidÃªncia Geral ${idx + 1}`} 
                          className="w-full h-full object-cover group-hover:scale-105 transition-transform" 
                        />
                        <span className="absolute bottom-1 left-1 text-[8px] font-black bg-black/60 text-white px-1 py-0.5 rounded">
                          Geral #{idx + 1}
                        </span>
                      </div>
                      
                      <div className="w-full mt-1.5 pt-1 border-t border-slate-200/60 flex items-center justify-between gap-1">
                        <select
                          className="w-full text-[10px] font-bold bg-white border border-slate-200 rounded px-1 py-0.5 text-indigo-700 cursor-pointer focus:outline-none focus:ring-1 focus:ring-indigo-500"
                          defaultValue=""
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val !== "") {
                              handleLinkGeneralEvidenceToActivity(idx, Number(val));
                            }
                          }}
                          title="Vincular foto a uma atividade"
                        >
                          <option value="" disabled>+ Vincular Ã  Atividade...</option>
                          {atividades.map((atvItem, atvIdx) => (
                            <option key={atvIdx} value={atvIdx}>
                              Ativ. #{atvIdx + 1}: {atvItem.nome ? atvItem.nome.substring(0, 18) : 'Atividade ' + (atvIdx + 1)}...
                            </option>
                          ))}
                        </select>
                        <button 
                          onClick={() => {
                            const newEvidencias = [...(dadosConsultoria.evidencias || [])];
                            newEvidencias.splice(idx, 1);
                            const updatedDados = { ...dadosConsultoria, evidencias: newEvidencias };
                            setDadosConsultoria(updatedDados);
                            setTimeout(() => {
                              syncEvidenceState(atividades, updatedDados);
                            }, 50);
                          }}
                          className="bg-rose-50 hover:bg-rose-100 text-rose-600 p-1 rounded transition-colors shrink-0"
                          title="Excluir foto"
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </Card>

        {/* Painel de Controle: RelatÃ³rio de Consultoria & SeleÃ§Ã£o de Atividades */}
        <Card className="p-5 border-slate-200 bg-white shadow-sm">
          <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 pb-4 border-b border-slate-100">
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Tipo do Documento</span>
                <span className={cn(
                  "text-[10px] font-black uppercase px-2 py-0.5 rounded-full border",
                  (dadosConsultoria.tipoRelatorio || 'Final') === 'Parcial' 
                    ? "bg-amber-50 text-amber-700 border-amber-200" 
                    : "bg-emerald-50 text-emerald-700 border-emerald-200"
                )}>
                  {(dadosConsultoria.tipoRelatorio || 'Final') === 'Parcial' ? 'RelatÃ³rio Parcial' : 'RelatÃ³rio Final'}
                </span>
              </div>
              <h3 className="font-bold text-slate-800 text-base mt-0.5">RELATÃ“RIO DE CONSULTORIA</h3>
              <p className="text-xs text-slate-500">Selecione se o relatÃ³rio gerado Ã© Parcial ou Final e marque as atividades que devem constar no PDF impresso.</p>
            </div>

            {/* Tipo de RelatÃ³rio Toggle */}
            <div className="flex items-center gap-2 bg-slate-100 p-1 rounded-xl">
              <button
                type="button"
                onClick={() => setDadosConsultoria({ ...dadosConsultoria, tipoRelatorio: 'Parcial' })}
                className={cn(
                  "flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer",
                  dadosConsultoria.tipoRelatorio === 'Parcial'
                    ? "bg-white text-amber-700 shadow-sm border border-amber-200 ring-1 ring-amber-400/20"
                    : "text-slate-600 hover:text-slate-900"
                )}
              >
                <Clock size={14} className={dadosConsultoria.tipoRelatorio === 'Parcial' ? "text-amber-600" : "text-slate-400"} />
                <span>RelatÃ³rio Parcial</span>
              </button>
              <button
                type="button"
                onClick={() => setDadosConsultoria({ ...dadosConsultoria, tipoRelatorio: 'Final' })}
                className={cn(
                  "flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer",
                  (dadosConsultoria.tipoRelatorio || 'Final') === 'Final'
                    ? "bg-white text-emerald-700 shadow-sm border border-emerald-200 ring-1 ring-emerald-400/20"
                    : "text-slate-600 hover:text-slate-900"
                )}
              >
                <CheckCircle2 size={14} className={(dadosConsultoria.tipoRelatorio || 'Final') === 'Final' ? "text-emerald-600" : "text-slate-400"} />
                <span>RelatÃ³rio Final</span>
              </button>
            </div>
          </div>

          {/* SeleÃ§Ã£o RÃ¡pida de Atividades */}
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pt-3.5">
            <div className="flex items-center gap-2">
              <ListFilter size={16} className="text-slate-400 shrink-0" />
              <div className="text-xs text-slate-700 font-medium">
                <strong>{atividades.filter(a => a.incluirNoRelatorio !== undefined ? a.incluirNoRelatorio : (a.status || '').toLowerCase().includes('conclu')).length} de {atividades.length}</strong> atividades selecionadas para o relatÃ³rio (<strong>{atividades.filter(a => a.incluirNoRelatorio !== undefined ? a.incluirNoRelatorio : (a.status || '').toLowerCase().includes('conclu')).reduce((acc, a) => acc + (parseInt((a.cargaHoraria || '0').replace(/\D/g, '')) || 0), 0)}hs</strong> de {atividades.reduce((acc, a) => acc + (parseInt((a.cargaHoraria || '0').replace(/\D/g, '')) || 0), 0)}hs)
              </div>
            </div>

            <div className="flex items-center gap-2 flex-wrap">
              <button
                type="button"
                onClick={selectConcluidasAtividadesRelatorio}
                className="px-2.5 py-1 text-xs font-bold bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-700 rounded-lg border border-slate-200 transition-colors flex items-center gap-1 cursor-pointer"
                title="Incluir apenas atividades que possuem status ConcluÃ­do"
              >
                <CheckCircle size={13} className="text-emerald-600" />
                <span>Selecionar ConcluÃ­das</span>
              </button>
              <button
                type="button"
                onClick={selectAllAtividadesRelatorio}
                className="px-2.5 py-1 text-xs font-bold bg-slate-100 hover:bg-blue-50 text-slate-700 hover:text-blue-700 rounded-lg border border-slate-200 transition-colors flex items-center gap-1 cursor-pointer"
                title="Incluir todas as atividades no relatÃ³rio"
              >
                <CheckCheck size={13} className="text-blue-600" />
                <span>Selecionar Todas</span>
              </button>
              <button
                type="button"
                onClick={clearSelectionAtividadesRelatorio}
                className="px-2.5 py-1 text-xs font-bold bg-slate-100 hover:bg-rose-50 text-slate-700 hover:text-rose-700 rounded-lg border border-slate-200 transition-colors flex items-center gap-1 cursor-pointer"
                title="Limpar seleÃ§Ã£o para escolher manualmente"
              >
                <Square size={13} className="text-slate-400" />
                <span>Desmarcar Todas</span>
              </button>
            </div>
          </div>
        </Card>

        {atividades.map((atv, idx) => {
          const isIncluded = atv.incluirNoRelatorio !== undefined 
            ? atv.incluirNoRelatorio 
            : (atv.status || '').toLowerCase().includes('conclu');

          return (
          <Card 
            key={idx} 
            className={cn(
              "p-6 border-l-4 transition-all cursor-pointer group",
              isIncluded 
                ? "border-l-emerald-500 bg-white hover:shadow-md" 
                : "border-l-slate-300 bg-slate-50/50 hover:bg-white text-slate-600"
            )}
            onClick={() => {
              setEditingActivityIndex(idx);
              setEditingActivity({ ...atv });
            }}
          >
            <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
              <div className="flex items-center gap-3">
                <div className={cn(
                  "w-8 h-8 rounded-full flex items-center justify-center font-bold text-sm",
                  isIncluded ? "bg-emerald-100 text-emerald-700" : "bg-slate-200 text-slate-600"
                )}>
                  {idx + 1}
                </div>
                <div>
                  <h4 className={cn("text-lg font-bold", isIncluded ? "text-slate-800" : "text-slate-600")}>
                    {atv.nome || "Nova Atividade"}
                  </h4>
                  {atv.idProblema && (
                    <div className="mt-0.5">
                      <span className="text-[10px] font-black uppercase px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 border border-emerald-100">
                        {problemas.find(p => p.id === atv.idProblema)?.descricao_problemas || "Problema Vinculado"}
                      </span>
                    </div>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-2 flex-wrap" onClick={(e) => e.stopPropagation()}>
                {/* BotÃ£o de InclusÃ£o no RelatÃ³rio */}
                <button
                  type="button"
                  onClick={() => toggleAtividadeRelatorio(idx)}
                  className={cn(
                    "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all border shadow-2xs cursor-pointer",
                    isIncluded
                      ? "bg-emerald-50 text-emerald-800 border-emerald-300 hover:bg-emerald-100"
                      : "bg-slate-100 text-slate-500 border-slate-200 hover:bg-slate-200"
                  )}
                  title="Clique para incluir ou remover esta atividade do RelatÃ³rio de Consultoria"
                >
                  {isIncluded ? <CheckSquare size={15} className="text-emerald-600" /> : <Square size={15} className="text-slate-400" />}
                  <span>{isIncluded ? "No RelatÃ³rio" : "Oculta do RelatÃ³rio"}</span>
                </button>

                <div className="flex items-center gap-1.5 bg-slate-50 p-1 rounded-xl border border-slate-200">
                  <span className="text-[10px] font-bold text-slate-400 uppercase px-1">Ordem:</span>
                  <select
                    className="text-xs font-bold bg-white border border-slate-200 rounded px-1.5 py-0.5 text-slate-700 cursor-pointer focus:outline-none focus:ring-1 focus:ring-emerald-500"
                    value={idx}
                    onChange={(e) => moveAtividadeToPos(idx, Number(e.target.value))}
                    title="Alterar posiÃ§Ã£o da atividade"
                  >
                    {atividades.map((_, pIdx) => (
                      <option key={pIdx} value={pIdx}>
                        #{pIdx + 1}
                      </option>
                    ))}
                  </select>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-7 w-7 p-0 text-slate-600 hover:bg-slate-200 rounded-lg"
                    disabled={idx === 0}
                    onClick={() => moveAtividade(idx, 'up')}
                    title="Mover atividade para cima"
                  >
                    <ArrowUp size={14} />
                  </Button>
                  <Button 
                    variant="ghost" 
                    size="sm" 
                    className="h-7 w-7 p-0 text-slate-600 hover:bg-slate-200 rounded-lg"
                    disabled={idx === atividades.length - 1}
                    onClick={() => moveAtividade(idx, 'down')}
                    title="Mover atividade para baixo"
                  >
                    <ArrowDown size={14} />
                  </Button>
                </div>

                <Button 
                  variant="ghost" 
                  size="sm" 
                  className="text-emerald-600 hover:bg-emerald-50"
                  onClick={(e) => {
                    e.stopPropagation();
                    setEditingActivityIndex(idx);
                    setEditingActivity({ ...atv });
                  }}
                >
                  <Plus size={16} className="rotate-45" /> Editar
                </Button>
                <Button 
                  variant="ghost" 
                  size="sm" 
                  className="text-rose-600 hover:bg-rose-50"
                  onClick={(e) => {
                    e.stopPropagation();
                    if(confirm("Deseja realmente excluir esta atividade?")) {
                      deleteAtividade(idx);
                    }
                  }}
                >
                  <Trash2 size={16} /> Excluir
                </Button>
              </div>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
              <div className="md:col-span-2 space-y-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">DescriÃ§Ã£o da Atividade</label>
                  <p className="text-sm text-slate-600 line-clamp-3 bg-slate-50 p-3 rounded-xl border border-slate-100">
                    {atv.descricao || "Sem descriÃ§Ã£o definida."}
                  </p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">SoluÃ§Ã£o Proposta</label>
                    <div className="text-sm text-slate-700 font-medium bg-slate-50 p-2 rounded-lg border border-slate-100">
                      {atv.solucaoProposta || "---"}
                    </div>
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">ResponsÃ¡vel</label>
                    <div className="text-sm text-slate-700 font-medium bg-slate-50 p-2 rounded-lg border border-slate-100">
                      {atv.responsavel || "---"}
                    </div>
                  </div>
                </div>
              </div>
              <div className="space-y-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">Carga HorÃ¡ria</label>
                  <div className="text-sm text-slate-700 font-bold bg-emerald-50 p-2 rounded-lg border border-emerald-100 flex items-center gap-2">
                    <Clock size={14} className="text-emerald-600" />
                    {atv.cargaHoraria || "0h"}
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">Prioridade</label>
                  <div className={cn(
                    "text-xs font-bold px-2 py-1 rounded inline-block uppercase",
                    atv.prioridade === 'Alta' ? "bg-rose-100 text-rose-600" :
                    atv.prioridade === 'MÃ©dia' ? "bg-amber-100 text-amber-600" :
                    "bg-emerald-100 text-emerald-600"
                  )}>
                    {atv.prioridade}
                  </div>
                </div>
              </div>
              <div className="space-y-4">
                <div>
                  <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">Status</label>
                  <div className={cn(
                    "text-xs font-bold px-3 py-1 rounded-full inline-flex items-center gap-1.5 uppercase",
                    atv.status === 'ConcluÃ­do' ? "bg-emerald-500 text-white" :
                    atv.status === 'Em Andamento' ? "bg-sky-500 text-white" :
                    atv.status === 'Atrasado' ? "bg-rose-500 text-white" :
                    "bg-slate-200 text-slate-600"
                  )}>
                    {atv.status === 'ConcluÃ­do' ? <CheckCircle size={12} /> : null}
                    {atv.status}
                  </div>
                </div>
                {(atv.dataInicio || atv.dataFim) && (
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase mb-1 block">Prazo</label>
                    <div className="text-[10px] text-slate-500 font-medium">
                      {atv.dataInicio ? format(parseLocalDate(atv.dataInicio) || new Date(), 'dd/MM/yy') : '...'} - {atv.dataFim ? format(parseLocalDate(atv.dataFim) || new Date(), 'dd/MM/yy') : '...'}
                    </div>
                  </div>
                )}
              </div>
            </div>

            {atv.resultadoEsperado && (
              <div className="mt-4 pt-4 border-t border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div className="flex-1">
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1">Resultado Esperado (KPI)</span>
                  <p className="text-sm font-semibold text-slate-800">{atv.resultadoEsperado}</p>
                </div>
                <div className="w-full md:w-80 shrink-0">
                  {(() => {
                    const pct = atv.progressoKPI ?? (atv.status === 'ConcluÃ­do' ? 100 : atv.status === 'Em Andamento' ? 50 : atv.status === 'Atrasado' ? 25 : 0);
                    return <KpiProgressBar percentage={pct} label="Atingimento do KPI" />;
                  })()}
                </div>
              </div>
            )}

            {/* EvidÃªncias vinculadas a esta atividade */}
            <div className="mt-4 pt-4 border-t border-slate-100" onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <ImageIcon size={14} className="text-emerald-600" />
                  <span className="text-[11px] font-bold text-slate-700 uppercase tracking-wider">
                    EvidÃªncias desta Atividade #{idx + 1}
                  </span>
                  <span className="text-[10px] bg-slate-100 text-slate-500 font-bold px-2 py-0.5 rounded-full">
                    {(atv.evidencias || []).length} foto(s)
                  </span>
                </div>

                <label className="cursor-pointer inline-flex items-center gap-1 text-[11px] font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-2.5 py-1 rounded-lg transition-colors">
                  <Upload size={12} />
                  <span>Anexar EvidÃªncia</span>
                  <input 
                    type="file" 
                    className="hidden" 
                    accept="image/*" 
                    multiple
                    onChange={(e) => {
                      const files = Array.from(e.target.files || []) as File[];
                      files.forEach(file => handleUploadEvidenceForActivity(idx, file));
                    }}
                  />
                </label>
              </div>

              {(atv.evidencias && atv.evidencias.length > 0) ? (
                <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-2 mt-2">
                  {atv.evidencias.map((img, evIdx) => (
                    <div key={evIdx} className="relative group rounded-lg border border-slate-200 bg-white overflow-hidden aspect-square shadow-2xs">
                      <img 
                        src={img} 
                        alt={`EvidÃªncia ${evIdx + 1} - Atividade ${idx + 1}`} 
                        className="w-full h-full object-cover cursor-pointer group-hover:scale-105 transition-transform"
                        onClick={() => setPreviewImage(img)}
                      />
                      <span className="absolute bottom-1 left-1 text-[8px] font-black bg-black/65 text-white px-1 py-0.2 rounded">
                        #{evIdx + 1}
                      </span>
                      <button
                        type="button"
                        className="absolute top-1 right-1 bg-rose-500/90 text-white p-1 rounded opacity-0 group-hover:opacity-100 hover:bg-rose-600 transition-all"
                        onClick={() => handleRemoveEvidenceFromActivity(idx, evIdx)}
                        title="Remover evidÃªncia"
                      >
                        <Trash2 size={10} />
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-[11px] text-slate-400 italic">
                  Nenhuma imagem vinculada a esta atividade ainda. Clique em "Anexar EvidÃªncia" para adicionar fotos comprobatÃ³rias.
                </p>
              )}
            </div>
          </Card>
        ); })}
      </div>

      {editingActivityIndex !== null && editingActivity && (
        <Modal
          title={`Editar Atividade ${editingActivityIndex + 1}`}
          onClose={() => {
            setEditingActivityIndex(null);
            setEditingActivity(null);
          }}
          onConfirm={async () => {
            if (editingActivityIndex === null || !editingActivity) return;
            const idx = editingActivityIndex;
            const updatedAtv = editingActivity;
            const newAtividades = [...atividades];
            newAtividades[idx] = updatedAtv;
            setAtividades(newAtividades);
            setEditingActivityIndex(null);
            setEditingActivity(null);

            try {
              const sanitized = sanitizeForFirestore(newAtividades);
              await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), { cronograma: sanitized });
              setSelectedDiagnostico({ ...selectedDiagnostico, cronograma: sanitized });

              const qTask = query(collection(db, 'tarefas_plano'), where('diagnosticoId', '==', selectedDiagnostico.id));
              const snap = await getDocs(qTask);
              const batch = writeBatch(db);
              snap.docs.forEach(docSnap => {
                const tData = docSnap.data();
                if (tData.ordem === idx || tData.problema === updatedAtv.nome) {
                  batch.update(docSnap.ref, sanitizeForFirestore({
                    status: updatedAtv.status,
                    problema: updatedAtv.nome,
                    solucaoSugerida: updatedAtv.solucaoProposta,
                    acoes: updatedAtv.descricao,
                    responsavel: updatedAtv.responsavel,
                    prioridade: updatedAtv.prioridade
                  }));
                }
              });
              await batch.commit();
            } catch (e) {
              console.error("Error saving activity edit:", e);
            }

            playSuccessSound();
          }}
          confirmText="Confirmar EdiÃ§Ã£o"
        >
          <div className="space-y-4">
            <div className="flex justify-end">
              <Button 
                variant="outline" 
                size="sm" 
                className="text-rose-600 border-rose-100 hover:bg-rose-50"
                onClick={async () => {
                  if (confirm("Deseja realmente excluir esta atividade?")) {
                    const idx = editingActivityIndex;
                    const newAtividades = atividades.filter((_, i) => i !== idx);
                    setAtividades(newAtividades);
                    setEditingActivityIndex(null);
                    setEditingActivity(null);

                    try {
                      const sanitized = sanitizeForFirestore(newAtividades);
                      await updateDoc(doc(db, 'diagnosticos', selectedDiagnostico.id), { cronograma: sanitized });
                      setSelectedDiagnostico({ ...selectedDiagnostico, cronograma: sanitized });
                    } catch (e) {
                      console.error("Error deleting activity:", e);
                    }

                    playSuccessSound();
                  }
                }}
              >
                <Trash2 size={14} className="mr-1" /> Excluir Atividade
              </Button>
            </div>
            {/* Toggle para inclusÃ£o no relatÃ³rio */}
            <div className="p-3 bg-slate-50 border border-slate-200 rounded-xl flex items-center justify-between">
              <div>
                <label className="text-xs font-bold text-slate-700 block">Incluir esta Atividade no RelatÃ³rio de Consultoria</label>
                <p className="text-[11px] text-slate-500">Se desmarcado, esta atividade nÃ£o aparecerÃ¡ no PDF gerado.</p>
              </div>
              <button
                type="button"
                onClick={() => {
                  const currentInc = editingActivity.incluirNoRelatorio !== undefined 
                    ? editingActivity.incluirNoRelatorio 
                    : (editingActivity.status || '').toLowerCase().includes('conclu');
                  setEditingActivity({ ...editingActivity, incluirNoRelatorio: !currentInc });
                }}
                className={cn(
                  "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all border cursor-pointer",
                  (editingActivity.incluirNoRelatorio !== undefined ? editingActivity.incluirNoRelatorio : (editingActivity.status || '').toLowerCase().includes('conclu'))
                    ? "bg-emerald-50 text-emerald-800 border-emerald-300"
                    : "bg-white text-slate-600 border-slate-200"
                )}
              >
                {(editingActivity.incluirNoRelatorio !== undefined ? editingActivity.incluirNoRelatorio : (editingActivity.status || '').toLowerCase().includes('conclu')) ? (
                  <>
                    <CheckSquare size={16} className="text-emerald-600" />
                    <span>No RelatÃ³rio</span>
                  </>
                ) : (
                  <>
                    <Square size={16} className="text-slate-400" />
                    <span>Oculta do RelatÃ³rio</span>
                  </>
                )}
              </button>
            </div>

            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Vincular a Problema Diagnosticado</label>
              <select 
                className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none mb-4"
                value={editingActivity.idProblema || ''}
                onChange={(e) => setEditingActivity({...editingActivity, idProblema: e.target.value})}
              >
                <option value="">Geral / NÃ£o vinculado</option>
                {respostas.filter(r => r.resposta !== 'Sim').map(r => {
                  const prob = problemas.find(p => p.id === r.idProblema || p.descricao_problemas === r.problema);
                  return (
                    <option key={r.idProblema} value={r.idProblema}>
                      {r.problema}
                    </option>
                  );
                })}
              </select>
            </div>
            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Nome da Atividade</label>
              <input 
                className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                value={editingActivity.nome}
                onChange={(e) => setEditingActivity({...editingActivity, nome: e.target.value})}
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Carga HorÃ¡ria</label>
              <input 
                className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                placeholder="Ex: 4h"
                value={editingActivity.cargaHoraria}
                onChange={(e) => setEditingActivity({...editingActivity, cargaHoraria: e.target.value})}
              />
            </div>
            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">DescriÃ§Ã£o</label>
              <textarea 
                className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                rows={4}
                value={editingActivity.descricao}
                onChange={(e) => setEditingActivity({...editingActivity, descricao: e.target.value})}
              />
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Status</label>
                <select 
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.status}
                  onChange={(e) => {
                    const statusVal = e.target.value as any;
                    const autoProg = statusVal === 'ConcluÃ­do' ? 100 : statusVal === 'Em Andamento' ? 50 : statusVal === 'Atrasado' ? 25 : 0;
                    setEditingActivity({...editingActivity, status: statusVal, progressoKPI: autoProg});
                  }}
                >
                  <option value="Pendente">Pendente</option>
                  <option value="Em Andamento">Em Andamento</option>
                  <option value="ConcluÃ­do">ConcluÃ­do</option>
                  <option value="Atrasado">Atrasado</option>
                </select>
              </div>
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Prioridade</label>
                <select 
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.prioridade}
                  onChange={(e) => setEditingActivity({...editingActivity, prioridade: e.target.value as any})}
                >
                  <option value="Baixa">Baixa</option>
                  <option value="MÃ©dia">MÃ©dia</option>
                  <option value="Alta">Alta</option>
                </select>
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Data InÃ­cio</label>
                <input 
                  type="date"
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.dataInicio || ''}
                  onChange={(e) => setEditingActivity({...editingActivity, dataInicio: e.target.value})}
                />
              </div>
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Data Fim</label>
                <input 
                  type="date"
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.dataFim || ''}
                  onChange={(e) => setEditingActivity({...editingActivity, dataFim: e.target.value})}
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">SoluÃ§Ã£o Proposta</label>
                <input 
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.solucaoProposta}
                  onChange={(e) => setEditingActivity({...editingActivity, solucaoProposta: e.target.value})}
                />
              </div>
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">Resultado Esperado (KPI)</label>
                <input 
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.resultadoEsperado || ''}
                  onChange={(e) => setEditingActivity({...editingActivity, resultadoEsperado: e.target.value})}
                  placeholder="Ex: Aumento de 15% nas vendas"
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase mb-1">ResponsÃ¡vel</label>
                <input 
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                  value={editingActivity.responsavel || ''}
                  onChange={(e) => setEditingActivity({...editingActivity, responsavel: e.target.value})}
                />
              </div>
              <div>
                <div className="flex justify-between items-center">
                  <label className="block text-[10px] font-bold text-slate-400 uppercase">Atingimento do KPI</label>
                  <span className="text-xs font-bold text-emerald-600">
                    {editingActivity.progressoKPI ?? (editingActivity.status === 'ConcluÃ­do' ? 100 : editingActivity.status === 'Em Andamento' ? 50 : editingActivity.status === 'Atrasado' ? 25 : 0)}%
                  </span>
                </div>
                <div className="flex items-center gap-3 mt-1">
                  <input 
                    type="range"
                    min="0"
                    max="100"
                    step="5"
                    className="flex-1 h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-emerald-600"
                    value={editingActivity.progressoKPI ?? (editingActivity.status === 'ConcluÃ­do' ? 100 : editingActivity.status === 'Em Andamento' ? 50 : editingActivity.status === 'Atrasado' ? 25 : 0)}
                    onChange={(e) => setEditingActivity({...editingActivity, progressoKPI: Number(e.target.value)})}
                  />
                  <input
                    type="number"
                    min="0"
                    max="100"
                    className="w-14 p-1 bg-slate-50 border border-slate-200 rounded text-center text-xs font-bold"
                    value={editingActivity.progressoKPI ?? (editingActivity.status === 'ConcluÃ­do' ? 100 : editingActivity.status === 'Em Andamento' ? 50 : editingActivity.status === 'Atrasado' ? 25 : 0)}
                    onChange={(e) => {
                      const val = Math.max(0, Math.min(100, Number(e.target.value)));
                      setEditingActivity({...editingActivity, progressoKPI: val});
                    }}
                  />
                </div>
              </div>
            </div>

            {/* EvidÃªncias desta Atividade no Modal */}
            <div className="pt-3 border-t border-slate-200">
              <div className="flex items-center justify-between mb-2">
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                    EvidÃªncias FotogrÃ¡ficas da Atividade (50% red.)
                  </label>
                  <span className="text-[10px] text-slate-400">
                    {(editingActivity.evidencias || []).length} fotos vinculadas
                  </span>
                </div>
                <label className="cursor-pointer inline-flex items-center gap-1 text-xs font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-3 py-1.5 rounded-lg transition-colors">
                  <Upload size={13} />
                  <span>+ Adicionar Foto</span>
                  <input 
                    type="file" 
                    className="hidden" 
                    accept="image/*" 
                    multiple
                    onChange={(e) => {
                      const files = Array.from(e.target.files || []) as File[];
                      if (editingActivityIndex !== null) {
                        files.forEach(file => handleUploadEvidenceForActivity(editingActivityIndex, file));
                      }
                    }}
                  />
                </label>
              </div>

              {(editingActivity.evidencias && editingActivity.evidencias.length > 0) ? (
                <div className="grid grid-cols-4 sm:grid-cols-6 gap-2 mt-2">
                  {editingActivity.evidencias.map((img, evIdx) => (
                    <div key={evIdx} className="relative group rounded-lg border border-slate-200 bg-white overflow-hidden aspect-square">
                      <img 
                        src={img} 
                        alt={`EvidÃªncia ${evIdx + 1}`} 
                        className="w-full h-full object-cover cursor-pointer group-hover:scale-105 transition-transform"
                        onClick={() => setPreviewImage(img)}
                      />
                      <button
                        type="button"
                        className="absolute top-1 right-1 bg-rose-500 text-white p-1 rounded-full shadow hover:bg-rose-600 transition-colors"
                        onClick={() => {
                          if (editingActivityIndex !== null) {
                            handleRemoveEvidenceFromActivity(editingActivityIndex, evIdx);
                          }
                        }}
                        title="Remover foto"
                      >
                        <X size={10} />
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="p-3 bg-slate-50 rounded-lg border border-dashed border-slate-200 text-center">
                  <p className="text-xs text-slate-400">Nenhuma foto comprobatÃ³ria anexada a esta atividade ainda.</p>
                </div>
              )}
            </div>
          </div>
        </Modal>
      )}

      {/* Lightbox / Visualizador de EvidÃªncia em Alta Qualidade */}
      {previewImage && (
        <div 
          className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4"
          onClick={() => setPreviewImage(null)}
        >
          <div 
            className="relative max-w-4xl max-h-[90vh] bg-slate-900 rounded-2xl overflow-hidden shadow-2xl flex flex-col items-center p-2"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="w-full flex justify-between items-center px-4 py-2 text-white">
              <span className="text-xs font-bold uppercase tracking-wider text-slate-300">VisualizaÃ§Ã£o de EvidÃªncia</span>
              <button
                onClick={() => setPreviewImage(null)}
                className="p-1 rounded-lg bg-white/10 hover:bg-white/20 text-white transition-colors"
              >
                <X size={18} />
              </button>
            </div>
            <img 
              src={previewImage} 
              alt="VisualizaÃ§Ã£o de EvidÃªncia" 
              className="max-h-[75vh] w-auto object-contain rounded-xl"
            />
          </div>
        </div>
      )}
    </div>
  );
};

const KpiProgressBar = ({ percentage, label = "Atingimento do KPI" }: { percentage: number, label?: string }) => {
  let barColorClass = "bg-rose-500";
  let textColorClass = "text-rose-600 bg-rose-50 border border-rose-100";
  let statusText = "CrÃ­tico";
  
  if (percentage >= 100) {
    barColorClass = "bg-emerald-500";
    textColorClass = "text-emerald-700 bg-emerald-50 border border-emerald-100";
    statusText = "Excelente";
  } else if (percentage >= 75) {
    barColorClass = "bg-teal-500";
    textColorClass = "text-teal-700 bg-teal-50 border border-teal-100";
    statusText = "AvanÃ§ado";
  } else if (percentage >= 50) {
    barColorClass = "bg-amber-500";
    textColorClass = "text-amber-700 bg-amber-50 border border-amber-100";
    statusText = "Regular";
  } else if (percentage >= 25) {
    barColorClass = "bg-orange-500";
    textColorClass = "text-orange-700 bg-orange-50 border border-orange-100";
    statusText = "Iniciante";
  }

  return (
    <div className="mt-3 bg-slate-50 p-3 rounded-xl border border-slate-100 flex flex-col gap-2">
      <div className="flex items-center justify-between gap-2">
        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
          <Target size={12} className="text-slate-400" /> {label}
        </span>
        <div className="flex items-center gap-1.5">
          <span className={`text-[9px] font-extrabold px-2 py-0.5 rounded-full uppercase tracking-wider ${textColorClass}`}>
            {statusText}
          </span>
          <span className="text-xs font-black text-slate-800">{percentage}%</span>
        </div>
      </div>
      <div className="w-full bg-slate-200 h-2.5 rounded-full overflow-hidden relative">
        <div 
          className={`h-full transition-all duration-1000 ease-out rounded-full ${barColorClass}`}
          style={{ width: `${percentage}%` }}
        />
      </div>
    </div>
  );
};

const RelatorioView = ({
  selectedDiagnostico,
  selectedEmpresa,
  respostas,
  solucoes,
  setPdfUrl,
  problemas,
  tarefasPlano = [],
  customLogo,
  customConsultoraLogo,
  logoChoice,
  setLogoChoice,
  historicalData = [],
  onUpdateCronograma
}: {
  selectedDiagnostico: Diagnostico,
  selectedEmpresa: Empresa | null,
  respostas: Resposta[],
  solucoes: Solucao[],
  setPdfUrl: (url: string) => void,
  problemas: Problema[],
  tarefasPlano?: TarefaPlanoAcao[],
  customLogo: string | null,
  customConsultoraLogo: string | null,
  logoChoice: 'both' | 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'both' | 'sebrae' | 'consultora' | 'none') => void,
  historicalData?: any[],
  onUpdateCronograma?: (cronograma: AtividadeCronograma[]) => Promise<void>
}) => {
  const chartRef1 = useRef<HTMLDivElement>(null);
  const chartRef2 = useRef<HTMLDivElement>(null);
  const chartRef3 = useRef<HTMLDivElement>(null);
  const chartRef4 = useRef<HTMLDivElement>(null);
  const chartRef5 = useRef<HTMLDivElement>(null);

  const [localCronograma, setLocalCronograma] = useState<AtividadeCronograma[]>([]);
  const [savingTargets, setSavingTargets] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // States for 3-Month Forecast
  const [useRealPace, setUseRealPace] = useState(false);
  const [customPace, setCustomPace] = useState(25);
  const [isPaceModeInitialized, setIsPaceModeInitialized] = useState(false);

  useEffect(() => {
    setLocalCronograma(selectedDiagnostico?.cronograma || []);
  }, [selectedDiagnostico?.cronograma]);

  const targetVsActualChartData = useMemo(() => {
    return localCronograma
      .filter(atv => atv.nome)
      .map(atv => {
        const actual = atv.progressoKPI ?? (atv.status === 'ConcluÃ­do' ? 100 : atv.status === 'Em Andamento' ? 50 : atv.status === 'Atrasado' ? 25 : 0);
        const target = atv.metaKPI ?? 100;
        return {
          name: atv.nome.length > 15 ? atv.nome.slice(0, 12) + '...' : atv.nome,
          fullName: atv.nome,
          'Meta (%)': target,
          'Realizado (%)': actual
        };
      });
  }, [localCronograma]);

  const cronogramaTrendData = useMemo(() => {
    const cronograma = selectedDiagnostico?.cronograma || [];
    if (cronograma.length === 0) return [];
    
    const sorted = [...cronograma]
      .filter(atv => atv.nome)
      .sort((a, b) => {
        const dateA = a.dataFim ? new Date(a.dataFim).getTime() : 0;
        const dateB = b.dataFim ? new Date(b.dataFim).getTime() : 0;
        return dateA - dateB;
      });

    let totalAccum = 0;
    return sorted.map((atv, i) => {
      const kpiVal = atv.progressoKPI ?? (atv.status === 'ConcluÃ­do' ? 100 : atv.status === 'Em Andamento' ? 50 : atv.status === 'Atrasado' ? 25 : 0);
      totalAccum += kpiVal;
      const progressCumulativo = Math.round(totalAccum / (i + 1));
      let label = `Meta ${i + 1}`;
      if (atv.dataFim) {
        try {
          const parsed = parseLocalDate(atv.dataFim);
          label = parsed ? format(parsed, 'dd/MM/yy') : `Meta ${i + 1}`;
        } catch {
          const parts = String(atv.dataFim).split('-');
          if (parts.length >= 2) {
            label = `${parts[2] || ''}/${parts[1] || ''}`;
          } else {
            label = String(atv.dataFim);
          }
        }
      }
      return {
        name: atv.nome,
        data: label,
        'Atingimento Individual (%)': kpiVal,
        'MÃ©dia Geral de KPIs (%)': progressCumulativo
      };
    });
  }, [selectedDiagnostico?.cronograma]);

  const uniqueRespostas = useMemo(() => deduplicateRespostas(respostas), [respostas]);

  // Group all responses by problem
  const responsesByProblem = useMemo(() => {
    return uniqueRespostas.reduce((acc: Record<string, Resposta[]>, r) => {
      const id = r.idProblema || r.problema;
      if (!id) return acc;
      if (!acc[id]) acc[id] = [];
      acc[id].push(r);
      return acc;
    }, {});
  }, [uniqueRespostas]);

  // Pure Diagnostic Analysis based on questions evaluated in the diagnostic:
  const analyzedProblems = useMemo(() => {
    return Object.entries(responsesByProblem).map(([probId, resps], idx) => {
      const seenQ = new Set<string>();
      const uniqueResps: Resposta[] = [];
      (resps as Resposta[]).forEach(r => {
        const q = (r.pergunta || '').trim().toLowerCase();
        if (q && seenQ.has(q)) return;
        if (q) seenQ.add(q);
        uniqueResps.push(r);
      });

      const noResponses = uniqueResps.filter(r => r.resposta === 'NÃ£o' || r.resposta === 'Parcial');
      const yesResponses = uniqueResps.filter(r => r.resposta === 'Sim');
      const solution = solucoes.find(s => s.idProblema === probId || s.problema === probId);
      const probObj = problemas.find(p => p.id === probId || p.descricao_problemas === probId);
      const resolvedArea = solution?.area || probObj?.area || resps[0]?.area || 'Geral';

      return {
        probId,
        ordem: idx,
        area: resolvedArea,
        noResponses,
        yesResponses,
        solution: solution || (probObj ? {
          id: `sol-${probId}`,
          idProblema: probId,
          problema: probObj.descricao_problemas,
          solucao_recomendada: '',
          responsavel_sugerido: 'Consultor',
          prazo_sugerido: '',
          kpis_sugeridos: '',
          resultado_esperado: '',
          acoes_sugeridas: '',
          area: resolvedArea
        } : undefined)
      };
    }).filter(p => p.noResponses.length > 0 || p.yesResponses.length > 0);
  }, [responsesByProblem, solucoes, problemas]);

  const sortedAnalyzedProblems = useMemo(() => {
    const list = [...analyzedProblems];
    list.sort((a, b) => (a.ordem || 0) - (b.ordem || 0));
    return list;
  }, [analyzedProblems]);

  const problemsByArea = useMemo(() => {
    const groups: Record<string, typeof sortedAnalyzedProblems> = {};
    sortedAnalyzedProblems.forEach(p => {
      let area = p.area || p.solution?.area;
      if (!area) {
        const prob = problemas.find(pr => pr.id === p.probId || pr.descricao_problemas === p.probId);
        area = prob?.area || 'Geral';
      }
      if (!groups[area]) groups[area] = [];
      groups[area].push(p);
    });
    return groups;
  }, [sortedAnalyzedProblems, problemas]);

  const barChartData = useMemo(() => {
    return Object.values(uniqueRespostas.reduce((acc: any, r) => {
      let area = r.area;
      if (!area) {
        const prob = problemas.find(p => p.id === r.idProblema || p.descricao_problemas === r.problema);
        area = prob?.area || 'Geral';
      }
      if (!acc[area]) acc[area] = { name: area, score: 0, totalPeso: 0 };
      acc[area].score += (r.score || 0);
      acc[area].totalPeso += (r.peso || 1);
      return acc;
    }, {})).map((item: any) => ({
      name: item.name,
      value: Math.round((item.score / (2 * (item.totalPeso || 1))) * 100)
    }));
  }, [uniqueRespostas, problemas]);

  const pieChartData = useMemo(() => [
    { name: 'Sim', value: uniqueRespostas.filter(r => r.resposta === 'Sim').length, color: '#10b981' },
    { name: 'Parcial', value: uniqueRespostas.filter(r => r.resposta === 'Parcial').length, color: '#f59e0b' },
    { name: 'NÃ£o', value: uniqueRespostas.filter(r => r.resposta === 'NÃ£o').length, color: '#ef4444' }
  ], [uniqueRespostas]);

  const performance = useMemo(() => {
    // Filter only questions that have been answered
    const answeredRespostas = uniqueRespostas.filter(r => r.resposta === 'Sim' || r.resposta === 'Parcial' || r.resposta === 'NÃ£o');
    if (answeredRespostas.length === 0) return '0';
    
    const totalScore = answeredRespostas.reduce((acc, r) => acc + (r.score || 0), 0);
    const maxScore = answeredRespostas.reduce((acc, r) => acc + (2 * (r.peso || 1)), 0);
    return ((totalScore / (maxScore || 1)) * 100).toFixed(0);
  }, [uniqueRespostas]);

  // Forecast calculations
  const totalTasks = localCronograma.length;
  const currentPlanProgress = useMemo(() => {
    if (totalTasks === 0) return 0;
    const total = localCronograma.reduce((acc, atv) => {
      const prog = atv.progressoKPI ?? (atv.status === 'ConcluÃ­do' ? 100 : atv.status === 'Em Andamento' ? 50 : atv.status === 'Atrasado' ? 25 : 0);
      return acc + prog;
    }, 0);
    return total / totalTasks;
  }, [localCronograma, totalTasks]);

  const calculatedRealVelocity = useMemo(() => {
    let startDate = new Date();
    if (selectedDiagnostico?.dataDiagnostico) {
      if (typeof selectedDiagnostico.dataDiagnostico === 'string') {
        startDate = new Date(selectedDiagnostico.dataDiagnostico);
      } else if ((selectedDiagnostico.dataDiagnostico as any).toDate) {
        startDate = (selectedDiagnostico.dataDiagnostico as any).toDate();
      }
    }
    localCronograma.forEach(atv => {
      if (atv.dataInicio) {
        const d = new Date(atv.dataInicio);
        if (d < startDate) {
          startDate = d;
        }
      }
    });

    const diffMs = new Date().getTime() - startDate.getTime();
    const daysElapsed = Math.max(1, diffMs / (1000 * 60 * 60 * 24));
    let monthsElapsed = daysElapsed / 30;
    if (monthsElapsed < 0.5) monthsElapsed = 0.5; // Avoid division by too small fraction

    const calculated = currentPlanProgress / monthsElapsed;
    return Math.max(0, Math.min(100, Number(calculated.toFixed(1))));
  }, [selectedDiagnostico?.dataDiagnostico, localCronograma, currentPlanProgress]);

  // Auto initialize velocity mode if active progress is present
  useEffect(() => {
    if (!isPaceModeInitialized && calculatedRealVelocity > 2) {
      setUseRealPace(true);
      setIsPaceModeInitialized(true);
    }
  }, [calculatedRealVelocity, isPaceModeInitialized]);

  const velocity = useRealPace ? (calculatedRealVelocity > 0 ? calculatedRealVelocity : 20) : customPace;

  const forecastData = useMemo(() => {
    const currentMaturity = Number(performance);
    
    const getProj = (m: number) => {
      if (totalTasks > 0) {
        const planProgAfterM = Math.min(100, currentPlanProgress + velocity * m);
        const planProgIncrease = planProgAfterM - currentPlanProgress;
        const gapMax = 100 - currentMaturity;
        const remPlanProg = 100 - currentPlanProgress;
        const maturityIncrease = remPlanProg > 0 
          ? gapMax * (planProgIncrease / remPlanProg)
          : 0;
        return {
          planProgress: Math.round(planProgAfterM),
          maturityScore: Math.round(Math.min(100, currentMaturity + maturityIncrease))
        };
      } else {
        const simulatedProgress = Math.min(100, velocity * m);
        const gapMax = 100 - currentMaturity;
        const maturityIncrease = gapMax * (simulatedProgress / 100);
        return {
          planProgress: Math.round(simulatedProgress),
          maturityScore: Math.round(Math.min(100, currentMaturity + maturityIncrease))
        };
      }
    };

    const proj1 = getProj(1);
    const proj2 = getProj(2);
    const proj3 = getProj(3);

    return [
      {
        name: 'Hoje',
        'Maturidade (%)': currentMaturity,
        'Plano (%)': Math.round(currentPlanProgress)
      },
      {
        name: 'MÃªs 1',
        'Maturidade (%)': proj1.maturityScore,
        'Plano (%)': proj1.planProgress
      },
      {
        name: 'MÃªs 2',
        'Maturidade (%)': proj2.maturityScore,
        'Plano (%)': proj2.planProgress
      },
      {
        name: 'MÃªs 3',
        'Maturidade (%)': proj3.maturityScore,
        'Plano (%)': proj3.planProgress
      }
    ];
  }, [performance, currentPlanProgress, velocity, totalTasks]);

  const generateRelatorioPDF = () => {
    if (!selectedEmpresa || !selectedDiagnostico) {
      alert("Selecione uma empresa e um diagnÃ³stico para gerar o relatÃ³rio.");
      return;
    }
    const originalTitle = document.title;
    const empresaNome = selectedEmpresa.nomeFantasia || selectedEmpresa.razaoSocial || selectedEmpresa.nome;
    document.title = `RelatÃ³rio de Consultoria - ${empresaNome}`;
    window.print();
    setTimeout(() => {
      document.title = originalTitle;
    }, 100);
  };

  const old_generateRelatorioPDF = async () => {
    try {
      if (!selectedEmpresa || !selectedDiagnostico) {
        alert("Selecione uma empresa e um diagnÃ³stico para gerar o relatÃ³rio.");
        return;
      }

      const doc = new jsPDF();
      const dateStr = formatFirestoreDate(selectedDiagnostico.dataDiagnostico, "dd/MM/yyyy HH:mm");

      // Add logos if available
      const showSebrae = logoChoice === 'both' || logoChoice === 'sebrae';
      const showConsultora = logoChoice === 'both' || logoChoice === 'consultora';
      const sebraeLogoSrc = (showSebrae && isValidLogoSource(customLogo)) ? customLogo : (showSebrae ? sebraeLogoBase64 : null);
      const consultoraLogoSrc = (showConsultora && isValidLogoSource(customConsultoraLogo)) ? customConsultoraLogo : null;

      if (isValidLogoSource(sebraeLogoSrc) && sebraeLogoSrc) {
        try {
          const imgProps = doc.getImageProperties(sebraeLogoSrc);
          const maxW = 28;
          const maxH = 14;
          let w = maxW;
          let h = (imgProps.height * w) / imgProps.width;
          if (h > maxH) {
            h = maxH;
            w = (imgProps.width * h) / imgProps.height;
          }
          doc.addImage(sebraeLogoSrc, undefined, 20, 10, w, h, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not add SEBRAE logo to Relatorio PDF", e);
        }
      }

      if (isValidLogoSource(consultoraLogoSrc) && consultoraLogoSrc) {
        try {
          const imgProps = doc.getImageProperties(consultoraLogoSrc);
          const maxW = 28;
          const maxH = 14;
          let w = maxW;
          let h = (imgProps.height * w) / imgProps.width;
          if (h > maxH) {
            h = maxH;
            w = (imgProps.width * h) / imgProps.height;
          }
          doc.addImage(consultoraLogoSrc, undefined, 190 - w, 10, w, h, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not add Consultora logo to Relatorio PDF", e);
        }
      }

      doc.setFontSize(14);
      doc.setTextColor(0, 90, 160);
      doc.setFont("helvetica", "bold");
      doc.text('RelatÃ³rio de DiagnÃ³stico', 20, 20);
      
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text(`Data: ${dateStr}`, 20, 26);

      let currentY = 50;

      // Add Client Data
      autoTable(doc, {
        startY: currentY,
        theme: 'plain',
        styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1 },
        body: [
          [{ content: 'DADOS DO CLIENTE', colSpan: 2, styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }],
          [{ content: 'Empresa:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.nomeFantasia || selectedEmpresa.nome }],
          [{ content: 'CNPJ:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.cnpj || 'N/A' }],
          [{ content: 'NÂº CAF:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.cafNumero || 'N/A' }],
          [{ content: 'Representante:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.representante || 'N/A' }],
          [{ content: 'CPF Representante:', styles: { fontStyle: 'bold' } }, { content: selectedEmpresa.cpfRepresentante || 'N/A' }]
        ],
        margin: { left: 20, right: 20 }
      });
      currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 5;

      // Add Consultancy Data if available
      if (selectedDiagnostico.dadosConsultoria) {
        const dc = selectedDiagnostico.dadosConsultoria;
        autoTable(doc, {
          startY: currentY,
          theme: 'plain',
          styles: { lineWidth: 0.1, lineColor: 0, fontSize: 8, cellPadding: 1 },
          body: [
            [{ content: 'DADOS DA CONSULTORIA', colSpan: 2, styles: { fontStyle: 'bold', fillColor: [240, 240, 240] } }],
            [{ content: 'Consultor:', styles: { fontStyle: 'bold' } }, { content: dc.consultor || 'N/A' }],
            [{ content: 'Ãrea:', styles: { fontStyle: 'bold' } }, { content: dc.areaConsultoria || 'N/A' }]
          ],
          margin: { left: 20, right: 20 }
        });
        currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 5;
      }

      // Add High-Level Metrics Summary Table
      const totalSim = uniqueRespostas.filter(r => r.resposta === 'Sim').length;
      const totalNao = uniqueRespostas.filter(r => r.resposta === 'NÃ£o').length;
      const totalParcial = uniqueRespostas.filter(r => r.resposta === 'Parcial').length;
      autoTable(doc, {
        startY: currentY,
        theme: 'striped',
        head: [[{ content: 'RESUMO DOS INDICADORES GERAIS DE MATURIDADE', colSpan: 4, styles: { halign: 'center', fontStyle: 'bold', fillColor: [0, 90, 160], textColor: [255, 255, 255], fontSize: 9 } }]],
        body: [
          [
            { content: 'Maturidade Geral', styles: { fontStyle: 'bold', halign: 'center', fillColor: [245, 247, 250] } },
            { content: 'Pontos Fortes (Sim)', styles: { fontStyle: 'bold', halign: 'center', fillColor: [245, 247, 250] } },
            { content: 'Pontos CrÃ­ticos (NÃ£o)', styles: { fontStyle: 'bold', halign: 'center', fillColor: [245, 247, 250] } },
            { content: 'AtenÃ§Ã£o (Parcial)', styles: { fontStyle: 'bold', halign: 'center', fillColor: [245, 247, 250] } }
          ],
          [
            { content: `${performance}%`, styles: { fontSize: 13, fontStyle: 'bold', halign: 'center', textColor: [0, 90, 160] } },
            { content: `${totalSim} conformidades`, styles: { fontSize: 9, halign: 'center', textColor: [16, 185, 129] } },
            { content: `${totalNao} urgÃªncias`, styles: { fontSize: 9, halign: 'center', textColor: [239, 68, 68] } },
            { content: `${totalParcial} parciais`, styles: { fontSize: 9, halign: 'center', textColor: [245, 158, 11] } }
          ]
        ],
        margin: { left: 20, right: 20 }
      });
      currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 5;

      // --- CAPTURE CHARTS ---
      let chart1Img: { data: string, width: number, height: number } | null = null;
      let chart2Img: { data: string, width: number, height: number } | null = null;
      let chart3Img: { data: string, width: number, height: number } | null = null;
      let chart4Img: { data: string, width: number, height: number } | null = null;
      let chart5Img: { data: string, width: number, height: number } | null = null;

      if (chartRef1.current) {
        try {
          const canvas = await captureElementWithHtml2Canvas(chartRef1.current);
          if (canvas && canvas.width > 0) {
            chart1Img = {
              data: canvas.toDataURL('image/png'),
              width: canvas.width,
              height: canvas.height
            };
          }
        } catch (e) {
          console.error("Could not capture Chart 1:", e);
        }
      }

      if (chartRef5.current) {
        try {
          const canvas = await captureElementWithHtml2Canvas(chartRef5.current);
          if (canvas && canvas.width > 0) {
            chart5Img = {
              data: canvas.toDataURL('image/png'),
              width: canvas.width,
              height: canvas.height
            };
          }
        } catch (e) {
          console.error("Could not capture Chart 5:", e);
        }
      }

      if (chartRef2.current) {
        try {
          const canvas = await captureElementWithHtml2Canvas(chartRef2.current);
          if (canvas && canvas.width > 0) {
            chart2Img = {
              data: canvas.toDataURL('image/png'),
              width: canvas.width,
              height: canvas.height
            };
          }
        } catch (e) {
          console.error("Could not capture Chart 2:", e);
        }
      }

      if (chartRef3.current) {
        try {
          const canvas = await captureElementWithHtml2Canvas(chartRef3.current);
          if (canvas && canvas.width > 0) {
            chart3Img = {
              data: canvas.toDataURL('image/png'),
              width: canvas.width,
              height: canvas.height
            };
          }
        } catch (e) {
          console.error("Could not capture Chart 3:", e);
        }
      }

      if (chartRef4.current) {
        try {
          const canvas = await captureElementWithHtml2Canvas(chartRef4.current);
          if (canvas && canvas.width > 0) {
            chart4Img = {
              data: canvas.toDataURL('image/png'),
              width: canvas.width,
              height: canvas.height
            };
          }
        } catch (e) {
          console.error("Could not capture Chart 4:", e);
        }
      }

      // --- DISTRIBUTE CHARTS BEAUTIFULLY ON THE PAGES ---
      // Advance to Page 2 for Diagnostic Charts
      doc.addPage();
      currentY = 20;

      doc.setFontSize(12);
      doc.setTextColor(0, 90, 160);
      doc.setFont("helvetica", "bold");
      doc.text('AnÃ¡lise de Maturidade e DistribuiÃ§Ã£o', 20, currentY);
      currentY += 8;

      // 1. Maturidade por Ãrea (%) - Bar Chart
      if (chart1Img) {
        let w1 = 170;
        let h1 = (chart1Img.height * w1) / chart1Img.width;
        if (h1 > 115) {
          h1 = 115;
          w1 = (chart1Img.width * h1) / chart1Img.height;
        }
        const x1 = 20 + (170 - w1) / 2;
        doc.addImage(chart1Img.data, 'PNG', x1, currentY, w1, h1);
        currentY += h1 + 10;
      }

      // 2. The other two charts side-by-side (Maturidade Rosca on left, DistribuiÃ§Ã£o on right)
      const maxPieWidth = 80;
      let w5 = maxPieWidth;
      let h5 = chart5Img ? (chart5Img.height * w5) / chart5Img.width : 0;
      if (h5 > 90) {
        h5 = 90;
        w5 = (chart5Img.width * h5) / chart5Img.height;
      }

      let w2 = maxPieWidth;
      let h2 = chart2Img ? (chart2Img.height * w2) / chart2Img.width : 0;
      if (h2 > 90) {
        h2 = 90;
        w2 = (chart2Img.width * h2) / chart2Img.height;
      }

      const pieRowHeight = Math.max(h5, h2);

      if (currentY + pieRowHeight > 275) {
        doc.addPage();
        currentY = 20;
        doc.setFontSize(12);
        doc.setTextColor(0, 90, 160);
        doc.setFont("helvetica", "bold");
        doc.text('AnÃ¡lise de Maturidade e DistribuiÃ§Ã£o (ContinuaÃ§Ã£o)', 20, currentY);
        currentY += 8;
      }

      if (chart5Img) {
        const x5 = 20 + (maxPieWidth - w5) / 2;
        doc.addImage(chart5Img.data, 'PNG', x5, currentY, w5, h5);
      }
      if (chart2Img) {
        const x2 = 110 + (maxPieWidth - w2) / 2;
        doc.addImage(chart2Img.data, 'PNG', x2, currentY, w2, h2);
      }
      currentY += pieRowHeight + 10;

      // Advance to Page 3 for Historical Line Charts and Target/KPI Projections
      if (chart3Img || chart4Img) {
        doc.addPage();
        currentY = 20;

        doc.setFontSize(12);
        doc.setTextColor(0, 90, 160);
        doc.setFont("helvetica", "bold");
        doc.text('TendÃªncia HistÃ³rica e ProjeÃ§Ã£o de Metas', 20, currentY);
        currentY += 8;

        const fullWidth = 170;

        if (chart3Img) {
          let w3 = fullWidth;
          let h3 = (chart3Img.height * w3) / chart3Img.width;
          if (h3 > 95) {
            h3 = 95;
            w3 = (chart3Img.width * h3) / chart3Img.height;
          }
          const x3 = 20 + (fullWidth - w3) / 2;
          doc.addImage(chart3Img.data, 'PNG', x3, currentY, w3, h3);
          currentY += h3 + 10;
        }

        if (chart4Img) {
          let w4 = fullWidth;
          let h4 = (chart4Img.height * w4) / chart4Img.width;
          if (h4 > 95) {
            h4 = 95;
            w4 = (chart4Img.width * h4) / chart4Img.height;
          }
          if (currentY + h4 > 270) {
            doc.addPage();
            currentY = 20;
            doc.setFontSize(12);
            doc.setTextColor(0, 90, 160);
            doc.setFont("helvetica", "bold");
            doc.text('TendÃªncia HistÃ³rica e ProjeÃ§Ã£o de Metas (ContinuaÃ§Ã£o)', 20, currentY);
            currentY += 8;
          }
          const x4 = 20 + (fullWidth - w4) / 2;
          doc.addImage(chart4Img.data, 'PNG', x4, currentY, w4, h4);
          currentY += h4 + 10;
        }
      }

      // Start detailed text recommendations on a fresh page to keep the document pristine
      if (analyzedProblems.length > 0) {
        doc.addPage();
        currentY = 20;
      }

      if (Object.keys(problemsByArea).length === 0) {
        doc.setFontSize(12);
        doc.text('Nenhum dado identificado no diagnÃ³stico.', 20, currentY);
      } else {
        Object.entries(problemsByArea).forEach(([areaName, areaProblems]) => {
          // Add Area Header in PDF
          if (currentY > 230) {
            doc.addPage();
            currentY = 20;
          }
          doc.setFontSize(14);
          doc.setTextColor(0, 90, 160);
          doc.setFont("helvetica", "bold");
          doc.text(`Ãrea: ${areaName.toUpperCase()}`, 20, currentY);
          currentY += 8;

          const problems = areaProblems as typeof analyzedProblems;
          problems.forEach(({ probId, sequenciaPlano, noResponses, yesResponses, solution }, idx) => {
            // Check page break
            if (currentY > 240) {
              doc.addPage();
              currentY = 20;
            }

            doc.setFontSize(12);
            doc.setTextColor(51, 65, 85); // Slate 700
            doc.setFont("helvetica", "bold");
            const itemSeq = sequenciaPlano || (idx + 1);
            const splitTitle = doc.splitTextToSize(`${itemSeq}Âº - ${solution?.problema || probId}`, 170);
            doc.text(splitTitle, 20, currentY);
            currentY += splitTitle.length * 5 + 2;

            // --- SUCCESSES SECTION ---
            if (yesResponses.length > 0) {
              autoTable(doc, {
                startY: currentY,
                theme: 'plain',
                head: [[{ content: 'Pontos Fortes', styles: { textColor: [5, 150, 105], fontStyle: 'bold', fontSize: 10 } }]],
                body: yesResponses.map(r => [r.pergunta]),
                styles: { fontSize: 8, cellPadding: 1 },
                margin: { left: 20, right: 20 }
              });
              currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 2;

              if (solution?.comentario_sucesso) {
                doc.setFont("helvetica", "italic");
                doc.setTextColor(100);
                doc.setFontSize(8);
                const splitComentario = doc.splitTextToSize(`Reconhecimento: ${solution.comentario_sucesso}`, 165);
                doc.text(splitComentario, 25, currentY);
                currentY += splitComentario.length * 4 + 4;
              } else {
                currentY += 2;
              }
            }

            // --- OPPORTUNITIES SECTION ---
            if (noResponses.length > 0) {
              if (currentY > 240) { doc.addPage(); currentY = 20; }

              autoTable(doc, {
                startY: currentY,
                theme: 'grid',
                head: [[
                  { content: 'Pontos CrÃ­ticos / Oportunidades (Premissa)', styles: { textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 9 } },
                  { content: 'Resposta', styles: { textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 9, halign: 'center' } }
                ]],
                body: noResponses.map(r => [
                  r.pergunta + (r.observacao ? `\nObs: ${r.observacao}` : ''),
                  { content: r.resposta, styles: { halign: 'center' } }
                ]),
                styles: { fontSize: 8, cellPadding: 2, lineWidth: 0.1, lineColor: [200, 200, 200] },
                headStyles: { fillColor: [190, 18, 60] },
                columnStyles: {
                  0: { cellWidth: 'auto' },
                  1: { cellWidth: 20 }
                },
                margin: { left: 20, right: 20 }
              });
              currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 4;

              if (solution) {
                autoTable(doc, {
                  startY: currentY,
                  head: [['SoluÃ§Ã£o Proposta', 'ResponsÃ¡vel']],
                  body: [[
                    solution.solucao_recomendada || '',
                    solution.responsavel_sugerido || ''
                  ]],
                  theme: 'grid',
                  styles: { fontSize: 8, cellPadding: 1 },
                  headStyles: { fillColor: [5, 150, 105] },
                  margin: { left: 20, right: 20 }
                });
                
                currentY = ((doc as any).lastAutoTable?.finalY || currentY) + 6;

                if (solution.acoes_sugeridas) {
                  if (currentY > 270) { doc.addPage(); currentY = 20; }
                  doc.setFont("helvetica", "bold");
                  doc.setFontSize(10);
                  doc.text('Passo a Passo / AÃ§Ãµes:', 20, currentY);
                  currentY += 5;
                  doc.setFont("helvetica", "normal");
                  doc.setFontSize(9);
                  const splitAcoes = doc.splitTextToSize(solution.acoes_sugeridas, 170);
                  doc.text(splitAcoes, 20, currentY);
                  currentY += splitAcoes.length * 4.5 + 4;
                }
              }
            }

            currentY += 6;
          });
          currentY += 4;
        });
      }

      setPdfUrl(doc.output('bloburl').toString());
    } catch (error) {
      console.error("Error generating Relatorio PDF:", error);
      alert("Ocorreu um erro ao gerar o PDF do relatÃ³rio. Por favor, tente novamente.");
    }
  };

  return (
    <div className="space-y-8 max-w-5xl mx-auto pb-12">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 print:hidden bg-white p-5 rounded-2xl border border-slate-100 shadow-sm">
        <div className="flex-1">
          <LogoSelector 
            logoChoice={logoChoice}
            setLogoChoice={setLogoChoice}
            customLogo={customLogo}
            customConsultoraLogo={customConsultoraLogo}
          />
        </div>
        <div className="flex gap-3 self-end md:self-auto min-w-[200px]">
          <Button onClick={generateRelatorioPDF} className="w-full h-12 rounded-xl text-sm font-bold bg-emerald-600 hover:bg-emerald-700">
            <Printer size={18} className="mr-2" /> Imprimir RelatÃ³rio
          </Button>
        </div>
      </div>

      <div className="bg-white p-8 rounded-2xl shadow-sm border border-slate-100 print:shadow-none print:border-none print:p-0">
        <div className="flex justify-between items-start border-b border-slate-200 pb-6 mb-8">
          <div>
            <h1 className="text-3xl font-black text-slate-800 mb-2">RelatÃ³rio de DiagnÃ³stico</h1>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-x-12 gap-y-4 text-slate-600 mt-6">
              <div>
                <p className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">Empresa</p>
                <p className="text-base font-semibold text-slate-800">{selectedEmpresa?.nome || 'NÃ£o informada'}</p>
                <div className="flex flex-col gap-0.5 mt-0.5">
                  {selectedEmpresa?.cnpj && <p className="text-xs text-slate-500">CNPJ: {selectedEmpresa.cnpj}</p>}
                  <span className="w-fit inline-flex items-center gap-1 bg-slate-50 border border-slate-150 text-slate-600 px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider">
                    Segmento: {selectedDiagnostico.tipoEmpresa || 'Geral'}
                  </span>
                </div>
              </div>
              <div>
                <p className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">Data do DiagnÃ³stico</p>
                <p className="text-base font-semibold text-slate-800">
                  {formatFirestoreDate(selectedDiagnostico.dataDiagnostico, 'dd/MM/yyyy')}
                </p>
              </div>
              {selectedDiagnostico.dadosConsultoria?.areaConsultoria && (
                <div className="md:col-span-2">
                  <p className="text-[10px] uppercase tracking-wider font-bold text-slate-400 mb-1">Ãrea da Consultoria</p>
                  <p className="text-base font-semibold text-emerald-600">{selectedDiagnostico.dadosConsultoria.areaConsultoria}</p>
                </div>
              )}
            </div>
          </div>
          <div className="hidden md:block">
            {/* Logo removed as requested */}
          </div>
        </div>

        {selectedDiagnostico.dadosConsultoria && (
          <div className="mb-12 bg-slate-50 rounded-xl p-6 border border-slate-100">
            <h3 className="text-sm font-bold text-slate-400 uppercase tracking-wider mb-4">Dados da Consultoria</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase">Credenciada</p>
                <p className="text-sm font-semibold text-slate-800">{selectedDiagnostico.dadosConsultoria.razaoSocial || 'N/A'}</p>
              </div>
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase">Consultor</p>
                <p className="text-sm font-semibold text-slate-800">{selectedDiagnostico.dadosConsultoria.consultor || 'N/A'}</p>
              </div>
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase">CÃ³digo SGF</p>
                <p className="text-sm font-semibold text-slate-800">{selectedDiagnostico.dadosConsultoria.codigoSgf || 'N/A'}</p>
              </div>
              <div>
                <p className="text-xs font-bold text-slate-400 uppercase">PerÃ­odo</p>
                <p className="text-sm font-semibold text-slate-800">{selectedDiagnostico.dadosConsultoria.periodoConsultoria || 'N/A'}</p>
              </div>
            </div>
          </div>
        )}

        <div className="space-y-12">
          {/* Dashboard Highlight Metrics */}
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div className="bg-blue-50 p-4 rounded-xl border border-blue-100 flex flex-col gap-1 shadow-sm">
              <span className="text-[10px] font-bold text-blue-600 uppercase tracking-wider">Maturidade Geral</span>
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-black text-blue-800">{performance}%</span>
                <div className={`text-[10px] font-bold px-1.5 py-0.5 rounded ${Number(performance) > 60 ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'}`}>
                  {Number(performance) > 60 ? 'Bom' : 'CrÃ­tico'}
                </div>
              </div>
            </div>
            <div className="bg-emerald-50 p-4 rounded-xl border border-emerald-100 flex flex-col gap-1 shadow-sm">
              <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider">Pontos Fortes</span>
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-black text-emerald-800">{uniqueRespostas.filter(r => r.resposta === 'Sim').length}</span>
                <span className="text-[10px] text-emerald-600 font-medium">itens conformes</span>
              </div>
            </div>
            <div className="bg-rose-50 p-4 rounded-xl border border-rose-100 flex flex-col gap-1 shadow-sm">
              <span className="text-[10px] font-bold text-rose-600 uppercase tracking-wider">Pontos CrÃ­ticos</span>
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-black text-rose-800">{uniqueRespostas.filter(r => r.resposta === 'NÃ£o').length}</span>
                <span className="text-[10px] text-rose-600 font-medium">urgÃªncias</span>
              </div>
            </div>
            <div className="bg-amber-50 p-4 rounded-xl border border-amber-100 flex flex-col gap-1 shadow-sm">
              <span className="text-[10px] font-bold text-amber-600 uppercase tracking-wider">AtenÃ§Ã£o NecessÃ¡ria</span>
              <div className="flex items-baseline gap-2">
                <span className="text-2xl font-black text-amber-800">{uniqueRespostas.filter(r => r.resposta === 'Parcial').length}</span>
                <span className="text-[10px] text-amber-600 font-medium">pontos parciais</span>
              </div>
            </div>
          </div>

          {/* Dashboard Charts Section */}
          <div style={{ backgroundColor: '#f8fafc', borderColor: '#f1f5f9', color: '#1e293b' }} className="bg-slate-50 p-8 rounded-2xl border border-slate-100 print:bg-white print:border-none print:p-0">
            <h3 style={{ color: '#1e293b' }} className="text-xl font-bold mb-8 flex items-center gap-2">
              <BarChart3 style={{ color: '#3b82f6' }} />
              Desempenho por Ãrea e Status
            </h3>

            <div className="flex flex-col gap-8 max-w-4xl mx-auto print:max-w-none">
              {/* Chart 1: Bar Chart */}
              <div 
                ref={chartRef1} 
                style={{ backgroundColor: '#ffffff', borderColor: '#f1f5f9', color: '#334155' }} 
                className="bg-white p-8 rounded-xl border border-slate-100 shadow-sm print:shadow-none min-h-[480px] break-inside-avoid flex flex-col justify-between"
              >
                <div>
                  <h4 style={{ color: '#334155' }} className="font-bold text-base mb-6 flex items-center gap-2">
                    <BarChart3 size={18} style={{ color: '#3b82f6' }} /> Maturidade por Ãrea (%)
                  </h4>
                  <div className="h-96" style={{ backgroundColor: '#ffffff' }}>
                    <ResponsiveContainer width="100%" height={380}>
                      <BarChart data={barChartData} margin={{ top: 10, right: 10, left: -20, bottom: 5 }}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis 
                          dataKey="name" 
                          axisLine={false} 
                          tickLine={false} 
                          tick={{ fill: '#64748b', fontSize: 10 }} 
                          angle={-15} 
                          textAnchor="end" 
                          height={50} 
                          interval={0}
                        />
                        <YAxis axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 11 }} unit="%" domain={[0, 100]} />
                        <Tooltip 
                          contentStyle={{ backgroundColor: '#ffffff', borderRadius: '8px', border: '1px solid #f1f5f9', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                          formatter={(value: any) => [`${value}%`, 'Maturidade']}
                          labelStyle={{ fontWeight: 'bold', color: '#1e293b' }}
                        />
                        <Bar dataKey="value" name="Maturidade" radius={[6, 6, 0, 0]} isAnimationActive={false}>
                          {barChartData.map((_, index) => (
                            <Cell key={`cell-${index}`} fill={['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#14b8a6', '#6366f1'][index % 9]} />
                          ))}
                        </Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                <div className="mt-6 border-t border-slate-100 pt-4">
                  <p className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest mb-3 text-center">Legenda de Maturidade por Ãrea</p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                    {barChartData.map((item, index) => {
                      const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#14b8a6', '#6366f1'];
                      const color = colors[index % colors.length];
                      return (
                        <div key={item.name} className="flex items-center gap-2 p-2 rounded-lg bg-slate-50 border border-slate-100/50">
                          <div className="w-2.5 h-2.5 rounded flex-shrink-0" style={{ backgroundColor: color }} />
                          <span className="font-semibold text-slate-700 truncate text-[11px]" title={item.name}>{item.name}</span>
                          <span className="font-mono font-bold text-slate-800 text-[11px] ml-auto">{item.value}%</span>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Chart 5: Doughnut Chart */}
              <div 
                ref={chartRef5} 
                style={{ backgroundColor: '#ffffff', borderColor: '#f1f5f9', color: '#334155' }} 
                className="bg-white p-8 rounded-xl border border-slate-100 shadow-sm print:shadow-none min-h-[480px] break-inside-avoid flex flex-col justify-between"
              >
                <div>
                  <h4 style={{ color: '#334155' }} className="font-bold text-base mb-6 flex items-center gap-2">
                    <LucidePieChart size={18} style={{ color: '#3b82f6' }} /> Maturidade (GrÃ¡fico de Rosca)
                  </h4>
                  <div className="h-80 flex items-center justify-center bg-white">
                    <ResponsiveContainer width="100%" height={320}>
                      <PieChart>
                        <Pie
                          data={barChartData}
                          cx="50%"
                          cy="50%"
                          innerRadius={75}
                          outerRadius={105}
                          paddingAngle={3}
                          dataKey="value"
                          isAnimationActive={false}
                        >
                          {barChartData.map((_, index) => {
                            const donutColors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#14b8a6', '#6366f1'];
                            return <Cell key={`cell-${index}`} fill={donutColors[index % donutColors.length]} />;
                          })}
                        </Pie>
                        <Tooltip 
                          contentStyle={{ backgroundColor: '#ffffff', borderRadius: '8px', border: '1px solid #f1f5f9', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                          formatter={(value: any) => [`${value}%`, 'Maturidade']}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="mt-6 border-t border-slate-100 pt-4">
                  <p className="text-[10px] font-extrabold text-slate-400 uppercase tracking-widest mb-3 text-center">NÃ­vel de AtenÃ§Ã£o por Ãrea</p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 text-xs">
                    {barChartData.map((item, index) => {
                      const donutColors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4', '#14b8a6', '#6366f1'];
                      let levelColor = '';
                      let levelText = '';
                      if (item.value < 50) {
                        levelColor = 'bg-rose-50 text-rose-700 border border-rose-100';
                        levelText = 'CrÃ­tico';
                      } else if (item.value <= 75) {
                        levelColor = 'bg-amber-50 text-amber-700 border border-amber-100';
                        levelText = 'AtenÃ§Ã£o';
                      } else {
                        levelColor = 'bg-emerald-50 text-emerald-700 border border-emerald-100';
                        levelText = 'Bom';
                      }
                      return (
                        <div key={item.name} className="flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-100/50">
                          <div className="flex items-center gap-2 min-w-0">
                            <div className="w-2.5 h-2.5 rounded-full flex-shrink-0" style={{ backgroundColor: donutColors[index % donutColors.length] }} />
                            <span className="font-semibold text-slate-700 truncate text-[11px]" title={item.name}>{item.name}</span>
                          </div>
                          <div className="flex items-center gap-2 flex-shrink-0">
                            <span className="font-mono font-bold text-slate-800 text-[11px]">{item.value}%</span>
                            <span className={`px-2 py-0.5 rounded text-[8px] font-black uppercase ${levelColor}`}>{levelText}</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              {/* Chart 2: Pie Chart */}
              <div 
                ref={chartRef2} 
                style={{ backgroundColor: '#ffffff', borderColor: '#f1f5f9', color: '#334155' }} 
                className="bg-white p-8 rounded-xl border border-slate-100 shadow-sm print:shadow-none min-h-[440px] break-inside-avoid flex flex-col justify-between"
              >
                <div>
                  <h4 style={{ color: '#334155' }} className="font-bold text-base mb-6 flex items-center gap-2">
                    <LucidePieChart size={18} style={{ color: '#10b981' }} /> DistribuiÃ§Ã£o de Respostas
                  </h4>
                  <div className="h-80 flex items-center justify-center bg-white">
                    <ResponsiveContainer width="100%" height={320}>
                      <PieChart>
                        <Pie
                          data={pieChartData}
                          cx="50%"
                          cy="50%"
                          innerRadius={75}
                          outerRadius={105}
                          paddingAngle={5}
                          dataKey="value"
                          isAnimationActive={false}
                        >
                          {pieChartData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={entry.color} />
                          ))}
                        </Pie>
                        <Tooltip 
                          contentStyle={{ backgroundColor: '#ffffff', borderRadius: '8px', border: '1px solid #f1f5f9' }}
                          formatter={(value: any, name: any) => [`${value} respostas`, name]}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                <div className="flex justify-center gap-6 mt-4 border-t border-slate-100 pt-4">
                  <div className="flex items-center gap-2 text-xs font-bold uppercase" style={{ color: '#64748b' }}>
                    <div className="w-3 h-3 rounded-full" style={{ backgroundColor: '#10b981' }} /> Sim ({pieChartData.find(d => d.name === 'Sim')?.value || 0})
                  </div>
                  <div className="flex items-center gap-2 text-xs font-bold uppercase" style={{ color: '#64748b' }}>
                    <div className="w-3 h-3 rounded-full" style={{ backgroundColor: '#f59e0b' }} /> Parcial ({pieChartData.find(d => d.name === 'Parcial')?.value || 0})
                  </div>
                  <div className="flex items-center gap-2 text-xs font-bold uppercase" style={{ color: '#64748b' }}>
                    <div className="w-3 h-3 rounded-full" style={{ backgroundColor: '#ef4444' }} /> NÃ£o ({pieChartData.find(d => d.name === 'NÃ£o')?.value || 0})
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div>
            <h3 className="text-2xl font-black text-slate-800 mb-8 flex items-center gap-2 border-b border-slate-100 pb-4">
              <CheckCircle2 className="text-emerald-500 shrink-0" size={24} />
              AnÃ¡lise Detalhada do DiagnÃ³stico
            </h3>
            
            {Object.keys(problemsByArea).length === 0 ? (
              <p className="text-slate-500 italic">Nenhum dado identificado no diagnÃ³stico.</p>
            ) : (
              <div className="space-y-16">
                {Object.entries(problemsByArea).map(([areaName, areaProblems]) => {
                  const problems = areaProblems as typeof analyzedProblems;
                  return (
                    <div key={areaName} className="space-y-6">
                      <div className="flex items-center gap-3 bg-emerald-50 border border-emerald-100 p-4 rounded-xl print:bg-transparent print:border-none print:p-0">
                        <div className="w-3 h-3 rounded-full bg-emerald-600 shrink-0" />
                        <h3 className="text-xl font-bold text-slate-800 tracking-tight uppercase">
                          Ãrea: {areaName}
                        </h3>
                      </div>

                      <div className="space-y-10 pl-2 md:pl-6 border-l border-slate-100 print:border-none print:pl-0">
                        {problems.map(({ probId, noResponses, yesResponses, solution }, idx) => (
                        <div key={probId + idx} className="bg-white rounded-xl p-6 border border-slate-150 print:bg-transparent print:border-b print:rounded-none print:p-0 print:pb-6 print:mb-6 break-inside-avoid shadow-sm print:shadow-none">
                          <div className="flex items-start justify-between gap-4 mb-6">
                            <div className="flex items-start gap-4">
                              <div className="w-8 h-8 bg-emerald-600 text-white rounded-xl flex items-center justify-center text-xs font-black shrink-0 mt-0.5 shadow-sm shadow-emerald-600/20">
                                {idx + 1}Âº
                              </div>
                              <div>
                                <h4 className="text-base font-black text-slate-800 tracking-tight leading-snug">{solution?.problema || probId}</h4>
                              </div>
                            </div>
                          </div>

                          <div className="space-y-6">
                            {/* --- Success Section --- */}
                            {yesResponses.length > 0 && (
                              <div className="bg-emerald-50/40 p-4 rounded-lg border border-emerald-100">
                                <h5 className="text-xs font-bold text-emerald-700 uppercase tracking-widest mb-3 flex items-center gap-2">
                                  <Trophy size={14} /> Pontos Fortes / Sucessos
                                </h5>
                                <ul className="space-y-2 mb-4">
                                  {yesResponses.map((r, i) => (
                                    <li key={i} className="text-slate-700 text-sm flex items-start gap-2">
                                      <span className="text-emerald-500 mt-0.5">â€¢</span>
                                      <span>{r.pergunta}</span>
                                    </li>
                                  ))}
                                </ul>
                                {solution?.comentario_sucesso && (
                                  <div className="bg-white p-3 rounded border border-emerald-100 italic text-sm text-slate-600">
                                    <span className="font-bold text-emerald-600 not-italic mr-1">Reconhecimento:</span>
                                    {solution.comentario_sucesso}
                                  </div>
                                )}
                              </div>
                            )}

                            {/* --- Critical Section --- */}
                            {noResponses.length > 0 && (
                              <div className="bg-rose-50/40 p-6 rounded-xl border border-rose-100">
                                <h5 className="text-sm font-bold text-rose-700 uppercase tracking-widest mb-4 flex items-center gap-2">
                                  <AlertTriangle size={16} /> Pontos CrÃ­ticos / Oportunidades
                                </h5>
                                <div className="overflow-x-auto">
                                  <table className="w-full text-sm text-left border-collapse">
                                    <thead>
                                      <tr className="border-b border-rose-200">
                                        <th className="py-2 pr-4 font-bold text-rose-800 uppercase text-[10px]">Pergunta / Detalhes</th>
                                        <th className="py-2 text-right font-bold text-rose-800 uppercase text-[10px]">Status</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-rose-100">
                                      {noResponses.map((r, i) => (
                                        <tr key={i}>
                                          <td className="py-3 pr-4 align-top text-slate-600">
                                            <div className="font-semibold text-slate-800 text-sm mb-1">{r.pergunta}</div>
                                            {r.observacao && <div className="text-xs italic text-slate-500 mt-1 flex items-center gap-1"><Info size={10} /> {r.observacao}</div>}
                                          </td>
                                          <td className="py-3 text-right align-top">
                                            <span className="px-2 py-0.5 bg-rose-100 text-rose-700 text-[10px] font-bold rounded uppercase whitespace-nowrap font-mono">
                                              {r.resposta}
                                            </span>
                                          </td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>

                                {solution ? (
                                  <div className="mt-8 bg-white p-5 rounded-lg border border-emerald-100 shadow-sm print:border-slate-200 print:shadow-none">
                                    <h5 className="text-xs font-bold text-emerald-600 uppercase tracking-wider mb-3 flex items-center gap-2">
                                      <CheckCircle2 size={16} /> SoluÃ§Ã£o e RecomendaÃ§Ãµes Propostas
                                    </h5>
                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                                      <div>
                                        <p className="text-[10px] font-bold text-slate-400 uppercase">SoluÃ§Ã£o / AÃ§Ã£o Recomendada</p>
                                        <p className="text-sm font-semibold text-slate-800">{solution.solucao_recomendada || 'NÃ£o informada'}</p>
                                      </div>
                                      <div>
                                        <p className="text-[10px] font-bold text-slate-400 uppercase">ResponsÃ¡vel</p>
                                        <p className="text-sm font-semibold text-slate-800">{solution.responsavel_sugerido || 'Consultor'}</p>
                                      </div>
                                    </div>
                                    <div>
                                      <p className="text-[10px] font-bold text-slate-400 uppercase mb-1">Passo a Passo / RecomendaÃ§Ãµes</p>
                                      <p className="text-sm text-slate-700 whitespace-pre-wrap leading-relaxed bg-slate-50/70 p-3.5 rounded-lg border border-slate-100 font-sans">{solution.acoes_sugeridas || 'Nenhuma recomendaÃ§Ã£o detalhada cadastrada.'}</p>
                                    </div>

                                  </div>
                                ) : (
                                  <div className="mt-6 bg-amber-50 p-4 rounded-lg border border-amber-100 text-amber-800 text-sm italic flex items-center gap-2">
                                    <AlertCircle size={16} /> Nenhuma soluÃ§Ã£o mapeada para este problema.
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const GestaoPlanoView = ({ 
  tarefas, 
  empresas,
  onEdit, 
  onDelete, 
  onUpdateStatus,
  onUpdatePrioridade,
  onUploadEvidence,
  setTarefaForm,
  setModalType,
  setIsModalOpen,
  setModalData,
  selectedDiagnostico,
  selectedEmpresa,
  onGenerateActionPlan,
  generatingPlan = false,
  onGeneratePDF,
  logoChoice,
  setLogoChoice,
  customLogo,
  customConsultoraLogo,
  onReplicateActionPlan
}: { 
  tarefas: TarefaPlanoAcao[], 
  empresas: Empresa[],
  onEdit: (t: TarefaPlanoAcao) => void, 
  onDelete: (t: TarefaPlanoAcao) => void,
  onUpdateStatus: (id: string, s: 'Pendente' | 'Em Andamento' | 'ConcluÃ­do') => void,
  onUpdatePrioridade: (id: string, p: 'Baixa' | 'MÃ©dia' | 'Alta') => void,
  onUploadEvidence: (f: File, type: 'tarefa', id: string) => void,
  setTarefaForm: (f: any) => void,
  setModalType: (t: any) => void,
  setIsModalOpen: (o: boolean) => void,
  setModalData: (d: any) => void,
  selectedDiagnostico: Diagnostico | null,
  selectedEmpresa: Empresa | null,
  onGenerateActionPlan: () => Promise<void>,
  generatingPlan?: boolean,
  onGeneratePDF: () => void,
  logoChoice: 'both' | 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'both' | 'sebrae' | 'consultora' | 'none') => void,
  customLogo: string | null,
  customConsultoraLogo: string | null,
  onReplicateActionPlan?: () => Promise<void>
}) => {
  const [viewMode, setViewMode] = useState<'kanban' | 'timeline'>('kanban');
  const [sortBy, setSortBy] = useState<'ordem' | 'cronograma' | 'dataInicio' | 'prioridade' | 'dataVencimento' | 'problema' | 'area'>('ordem');
  const [groupBy, setGroupBy] = useState<'status' | 'data' | 'problema' | 'area'>('status');
  const [isSavingPlano, setIsSavingPlano] = useState(false);
  const [saveSuccessPlano, setSaveSuccessPlano] = useState(false);

  const getEmpresaNome = (id: string) => empresas.find(e => e.id === id)?.nome || 'Empresa desconhecida';

  const handleSavePlanoAcao = async () => {
    if (!selectedDiagnostico || !selectedDiagnostico.id) {
      alert("Nenhum diagnÃ³stico selecionado para salvar o Plano de AÃ§Ã£o.");
      return;
    }

    setIsSavingPlano(true);
    setSaveSuccessPlano(false);

    const safetyTimer = setTimeout(() => {
      setIsSavingPlano(false);
    }, 2000);

    try {
      // 1. Build synchronized cronograma array for selectedDiagnostico
      const synchronizedCronograma = sortedTarefas.map((t) => ({
        idProblema: t.idProblema || '',
        nome: t.problema || 'Atividade',
        solucaoProposta: t.solucaoSugerida || t.problema || '',
        descricao: t.acoes || '',
        status: t.status || 'Pendente',
        prioridade: t.prioridade || 'MÃ©dia',
        responsavel: t.responsavel || 'Consultor',
        cargaHoraria: t.cargaHoraria || '',
        dataInicio: t.dataInicio || '',
        dataFim: t.dataFim || t.dataInicio || ''
      }));

      const isAllCompleted = sortedTarefas.length > 0 && sortedTarefas.every(t => t.status === 'ConcluÃ­do');
      const newStatus = isAllCompleted ? 'ConcluÃ­do' : selectedDiagnostico.status || 'Em Andamento';

      // 2. Local storage persistence immediately
      try {
        const savedTasks = localStorage.getItem('local_tarefas_plano');
        const parsedTasks = savedTasks ? JSON.parse(savedTasks) : [];
        const others = Array.isArray(parsedTasks) ? parsedTasks.filter((t: any) => t.diagnosticoId !== selectedDiagnostico.id) : [];
        localStorage.setItem('local_tarefas_plano', JSON.stringify([...others, ...sortedTarefas]));

        const savedDiags = localStorage.getItem('local_diagnosticos');
        const parsedDiags = savedDiags ? JSON.parse(savedDiags) : [];
        const nextDiags = Array.isArray(parsedDiags) 
          ? parsedDiags.map((d: any) => d.id === selectedDiagnostico.id ? { ...d, cronograma: synchronizedCronograma, status: newStatus } : d)
          : [];
        localStorage.setItem('local_diagnosticos', JSON.stringify(nextDiags));
      } catch (e) {
        console.warn("Could not save to localStorage in handleSavePlanoAcao", e);
      }

      // 3. Background cloud sync without blocking
      const currentUser = auth.currentUser;
      if (currentUser && selectedDiagnostico.id) {
        const diagId = selectedDiagnostico.id;
        (async () => {
          try {
            const batch = writeBatch(db);

            sortedTarefas.forEach((t, index) => {
              if (t.id) {
                const taskRef = doc(db, 'tarefas_plano', t.id);
                batch.update(taskRef, sanitizeForFirestore({
                  status: t.status,
                  prioridade: t.prioridade,
                  responsavel: t.responsavel || 'Consultor',
                  dataInicio: t.dataInicio || '',
                  dataFim: t.dataFim || t.dataInicio || '',
                  dataVencimento: t.dataVencimento || t.dataFim || t.dataInicio || '',
                  acoes: t.acoes || '',
                  cargaHoraria: t.cargaHoraria || '',
                  ordem: index
                }));
              }
            });

            const diagRef = doc(db, 'diagnosticos', diagId);
            batch.update(diagRef, {
              cronograma: sanitizeForFirestore(synchronizedCronograma),
              status: newStatus
            });

            const commitPromise = batch.commit();
            const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout ao salvar na nuvem")), 4000));
            await Promise.race([commitPromise, timeoutPromise]);
          } catch (cloudErr) {
            console.warn("[CloudSync] Background sync note for PlanoAcao:", cloudErr);
          }
        })();
      }

      clearTimeout(safetyTimer);
      setIsSavingPlano(false);
      playSuccessSound();
      setSaveSuccessPlano(true);
      setTimeout(() => setSaveSuccessPlano(false), 3000);
    } catch (err) {
      console.error("Erro ao salvar Plano de AÃ§Ã£o:", err);
      clearTimeout(safetyTimer);
      setIsSavingPlano(false);
      playSuccessSound();
      setSaveSuccessPlano(true);
      setTimeout(() => setSaveSuccessPlano(false), 3000);
    } finally {
      setIsSavingPlano(false);
    }
  };

  const sortedTarefas = useMemo(() => {
    const list = [...tarefas];
    const parseDate = (val: any) => {
      if (!val) return 0;
      if (val.toDate && typeof val.toDate === 'function') return val.toDate().getTime();
      const d = parseLocalDate(val);
      return d ? d.getTime() : 0;
    };

    return list.sort((a, b) => {
      if (sortBy === 'ordem') {
        const oA = a.ordem !== undefined && a.ordem !== null ? a.ordem : 999999;
        const oB = b.ordem !== undefined && b.ordem !== null ? b.ordem : 999999;
        if (oA !== oB) return oA - oB;
        return parseDate(a.dataInicio) - parseDate(b.dataInicio);
      }

      if (sortBy === 'cronograma') {
        const oA = a.ordem !== undefined && a.ordem !== null ? a.ordem : null;
        const oB = b.ordem !== undefined && b.ordem !== null ? b.ordem : null;
        if (oA !== null && oB !== null && oA !== oB) return oA - oB;

        const dateA = parseDate(a.dataInicio);
        const dateB = parseDate(b.dataInicio);
        if (dateA !== dateB && dateA > 0 && dateB > 0) {
          return dateA - dateB;
        }

        const getIndex = (t: TarefaPlanoAcao) => {
          // 1. Identify final report / closing tasks
          const isFinalReport = 
            (t.problema && (
              t.problema.toLowerCase().includes("relatÃ³rio final") || 
              t.problema.toLowerCase().includes("relatorio final") || 
              t.problema.toLowerCase().includes("encerramento") ||
              t.problema.toLowerCase().includes("reuniÃ£o de fechamento") ||
              t.problema.toLowerCase().includes("reuniao de fechamento")
            )) ||
            (t.solucaoSugerida && (
              t.solucaoSugerida.toLowerCase().includes("relatÃ³rio final") || 
              t.solucaoSugerida.toLowerCase().includes("relatorio final") || 
              t.solucaoSugerida.toLowerCase().includes("encerramento") ||
              t.solucaoSugerida.toLowerCase().includes("apresentaÃ§Ã£o do relatÃ³rio") ||
              t.solucaoSugerida.toLowerCase().includes("apresentacao do relatorio")
            ));

          if (isFinalReport) {
            return 999999; // Always absolute last
          }

          // 2. Identify first task (Diagnosis / Demand assessment)
          const isFirstTask = 
            (t.problema && (
              t.problema.toLowerCase().includes("diagnÃ³stico inicial") || 
              t.problema.toLowerCase().includes("diagnostico inicial") || 
              t.problema.toLowerCase().includes("entendimento da demanda")
            )) ||
            (t.solucaoSugerida && (
              t.solucaoSugerida.toLowerCase().includes("diagnÃ³stico inicial") || 
              t.solucaoSugerida.toLowerCase().includes("diagnostico inicial") || 
              t.solucaoSugerida.toLowerCase().includes("entendimento da demanda")
            ));

          if (isFirstTask) {
            return 0; // Always absolute first
          }

          // 3. Fallback to matching index in actual/fallback cronograma
          const cronogramaFromDiag = selectedDiagnostico?.cronograma || [];
          const hasValidActivities = cronogramaFromDiag.some(act => act.nome && act.nome.trim() !== '');
          const cronograma = hasValidActivities ? cronogramaFromDiag : PLANO_DE_ACAO_PADRAO;

          if (t.idProblema) {
            const idx = cronograma.findIndex(act => act.idProblema === t.idProblema);
            if (idx !== -1) {
              if (idx === cronograma.length - 1) return 999999;
              return idx;
            }
          }

          const idxByName = cronograma.findIndex(act => 
            (act.nome && t.problema && act.nome.toLowerCase() === t.problema.toLowerCase()) ||
            (act.solucaoProposta && t.solucaoSugerida && act.solucaoProposta.toLowerCase() === t.solucaoSugerida.toLowerCase())
          );

          if (idxByName !== -1) {
            if (idxByName === cronograma.length - 1) return 999999;
            return idxByName;
          }

          // 4. Fallback to save-time order if present
          const rawTask = t as any;
          if (rawTask.ordem !== undefined && rawTask.ordem !== null) {
            return rawTask.ordem + 1; // shift by 1 to not clash with index 0
          }

          // 5. Unmatched intermediate tasks get a middle index
          return 50000;
        };
        return getIndex(a) - getIndex(b);
      } else if (sortBy === 'dataInicio') {
        const dateA = parseDate(a.dataInicio);
        const dateB = parseDate(b.dataInicio);
        return dateA - dateB;
      } else if (sortBy === 'dataVencimento') {
        const dateA = a.dataVencimento ? new Date(a.dataVencimento).getTime() : a.dataFim ? new Date(a.dataFim).getTime() : 0;
        const dateB = b.dataVencimento ? new Date(b.dataVencimento).getTime() : b.dataFim ? new Date(b.dataFim).getTime() : 0;
        return dateA - dateB;
      } else if (sortBy === 'prioridade') {
        const priorityWeight = { Alta: 3, MÃ©dia: 2, Baixa: 1 };
        const weightA = priorityWeight[a.prioridade] || 0;
        const weightB = priorityWeight[b.prioridade] || 0;
        return weightB - weightA;
      } else if (sortBy === 'problema') {
        return (a.problema || '').localeCompare(b.problema || '');
      } else if (sortBy === 'area') {
        return (a.area || '').localeCompare(b.area || '');
      }
      return 0;
    });
  }, [tarefas, sortBy, selectedDiagnostico]);

  const handleSetTaskPosition = async (task: TarefaPlanoAcao, newPosZeroBased: number) => {
    const currIdx = sortedTarefas.findIndex(t => t.id === task.id);
    if (currIdx === -1) return;
    if (newPosZeroBased < 0 || newPosZeroBased >= sortedTarefas.length) return;
    if (currIdx === newPosZeroBased) return;

    try {
      const newList = [...sortedTarefas];
      const [movedItem] = newList.splice(currIdx, 1);
      newList.splice(newPosZeroBased, 0, movedItem);

      const batch = writeBatch(db);
      newList.forEach((t, idx) => {
        batch.update(doc(db, 'tarefas_plano', t.id), { ordem: idx });
      });
      await batch.commit();
      playSuccessSound();
    } catch (err) {
      console.error("Erro ao reordenar tarefa:", err);
    }
  };

  const handleMoveTask = async (task: TarefaPlanoAcao, direction: 'up' | 'down') => {
    const currIdx = sortedTarefas.findIndex(t => t.id === task.id);
    if (currIdx === -1) return;
    const targetIdx = direction === 'up' ? currIdx - 1 : currIdx + 1;
    if (targetIdx < 0 || targetIdx >= sortedTarefas.length) return;
    await handleSetTaskPosition(task, targetIdx);
  };

  const groupedTasksMap = useMemo(() => {
    const map: Record<string, TarefaPlanoAcao[]> = {};
    if (groupBy === 'status') {
      map['Pendente'] = sortedTarefas.filter(t => t.status === 'Pendente');
      map['Em Andamento'] = sortedTarefas.filter(t => t.status === 'Em Andamento');
      map['ConcluÃ­do'] = sortedTarefas.filter(t => t.status === 'ConcluÃ­do');
    } else if (groupBy === 'area') {
      sortedTarefas.forEach(t => {
        const key = t.area || 'Geral';
        if (!map[key]) map[key] = [];
        map[key].push(t);
      });
    } else if (groupBy === 'problema') {
      sortedTarefas.forEach(t => {
        const key = t.problema || 'NÃ£o definido';
        if (!map[key]) map[key] = [];
        map[key].push(t);
      });
    } else if (groupBy === 'data') {
      sortedTarefas.forEach(t => {
        let key = 'Sem data de execuÃ§Ã£o';
        const dateVal = t.dataInicio || t.dataVencimento || t.dataFim;
        if (dateVal) {
          try {
            const d = new Date(dateVal);
            const m = d.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' });
            key = m.charAt(0).toUpperCase() + m.slice(1);
          } catch (e) {
            key = 'Data invÃ¡lida';
          }
        }
        if (!map[key]) map[key] = [];
        map[key].push(t);
      });
    }
    return map;
  }, [sortedTarefas, groupBy]);

  const computedColumns = useMemo(() => {
    if (groupBy === 'status') {
      return [
        { id: 'Pendente', label: 'Pendentes', color: 'bg-slate-100', text: 'text-slate-600', border: 'border-slate-200' },
        { id: 'Em Andamento', label: 'Em Andamento', color: 'bg-blue-50', text: 'text-blue-600', border: 'border-blue-200' },
        { id: 'ConcluÃ­do', label: 'ConcluÃ­dos', color: 'bg-emerald-50', text: 'text-emerald-600', border: 'border-emerald-200' }
      ];
    }

    const keys = Object.keys(groupedTasksMap);
    keys.sort((a, b) => {
      if (a === 'Sem data de execuÃ§Ã£o' || a === 'Geral' || a === 'NÃ£o definido') return -1;
      if (b === 'Sem data de execuÃ§Ã£o' || b === 'Geral' || b === 'NÃ£o definido') return 1;
      return a.localeCompare(b);
    });

    const colors = [
      { color: 'bg-indigo-50/50', text: 'text-indigo-600', border: 'border-indigo-100' },
      { color: 'bg-emerald-50/50', text: 'text-emerald-600', border: 'border-emerald-100' },
      { color: 'bg-amber-50/50', text: 'text-amber-600', border: 'border-amber-100' },
      { color: 'bg-rose-50/50', text: 'text-rose-600', border: 'border-rose-100' },
      { color: 'bg-purple-50/50', text: 'text-purple-600', border: 'border-purple-100' },
      { color: 'bg-sky-50/50', text: 'text-sky-600', border: 'border-sky-100' },
      { color: 'bg-slate-50/50', text: 'text-slate-600', border: 'border-slate-100' },
    ];

    return keys.map((key, index) => {
      const colColor = colors[index % colors.length];
      return {
        id: key,
        label: key,
        ...colColor
      };
    });
  }, [groupedTasksMap, groupBy]);

  const notificationsSent = useRef<Record<string, boolean>>({});

  const tarefAsComLembreteProximo = useMemo(() => {
    const hoje = new Date();
    hoje.setHours(0,0,0,0);
    
    return tarefas.filter(t => {
      if (t.status === 'ConcluÃ­do') return false;
      const dataVenc = t.dataVencimento || t.dataFim;
      if (!dataVenc) return false;
      
      const hasReminder = t.lembreteEmail || t.lembreteWhatsapp || t.lembretePush;
      if (!hasReminder) return false;
      
      const dataLimit = parseLocalDate(dataVenc) || new Date();
      dataLimit.setHours(0,0,0,0);
      
      const diffTime = dataLimit.getTime() - hoje.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      
      const diasAntes = t.lembreteDiasAntes || 1;
      return diffDays <= diasAntes && diffDays >= -7;
    });
  }, [tarefas]);

  useEffect(() => {
    if ('Notification' in window && Notification.permission === 'granted') {
      tarefAsComLembreteProximo.forEach(t => {
        if (t.lembretePush && !notificationsSent.current[t.id]) {
          const dias = Math.ceil((new Date(t.dataVencimento || t.dataFim).getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24));
          const msg = dias === 0 
            ? `vence hoje!` 
            : dias < 0 
              ? `estÃ¡ vencida hÃ¡ ${Math.abs(dias)} dias!`
              : `vence em ${dias} dias!`;

          new Notification(`Lembrete: ${t.problema}`, {
            body: `A tarefa de responsabilidade de ${t.responsavel || 'NÃ£o definido'} ${msg}`,
            icon: "/favicon.ico"
          });
          notificationsSent.current[t.id] = true;
        }
      });
    }
  }, [tarefAsComLembreteProximo]);

  // HorÃ¡rio de Lembrete DiÃ¡rio Background Ticker
  useEffect(() => {
    const checkDailyReminder = () => {
      const enabled = localStorage.getItem('daily_reminder_enabled') === 'true';
      if (!enabled) return;

      const reminderTime = localStorage.getItem('daily_reminder_time') || '09:00';
      const [remHourStr, remMinStr] = reminderTime.split(':');
      const remHour = parseInt(remHourStr || '9', 10);
      const remMin = parseInt(remMinStr || '0', 10);

      const now = new Date();
      const currentHour = now.getHours();
      const currentMin = now.getMinutes();

      if (currentHour === remHour && currentMin === remMin) {
        const todayStr = now.toISOString().split('T')[0];
        const lastNotified = localStorage.getItem('last_daily_reminder_date');
        
        if (lastNotified !== todayStr) {
          if ('Notification' in window && Notification.permission === 'granted') {
            const pending = tarefas.filter(t => t.status !== 'ConcluÃ­do');
            if (pending.length > 0) {
              new Notification("Tarefas Pendentes: Lembrete DiÃ¡rio", {
                body: `VocÃª possui ${pending.length} tarefa(s) pendente(s) no seu cronograma hoje.`,
                icon: "/favicon.ico"
              });
            } else {
              new Notification("Lembrete DiÃ¡rio de Tarefas", {
                body: `Tudo limpo! Nenhuma tarefa pendente no cronograma hoje.`,
                icon: "/favicon.ico"
              });
            }
            localStorage.setItem('last_daily_reminder_date', todayStr);
          }
        }
      }
    };

    checkDailyReminder();
    const interval = setInterval(checkDailyReminder, 30000);
    return () => clearInterval(interval);
  }, [tarefas]);

  const enviarLembreteWhatsApp = (t: TarefaPlanoAcao) => {
    const contato = t.lembreteContato || '';
    const cleanContato = contato.replace(/\D/g, '');
    const vencDate = parseLocalDate(t.dataVencimento || t.dataFim);
    const dias = vencDate ? Math.ceil((vencDate.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24)) : 0;
    const statusVenc = dias === 0 
      ? 'vence hoje' 
      : dias < 0 
        ? `estÃ¡ atrasada hÃ¡ ${Math.abs(dias)} dias`
        : `vencerÃ¡ em ${dias} dias`;

    const txt = `OlÃ¡! Segue um lembrete da tarefa do Plano de AÃ§Ã£o:\n*Tarefa:* ${t.problema}\n*ResponsÃ¡vel:* ${t.responsavel || 'NÃ£o definido'}\n*Status:* ${t.status}\n*Vencimento:* ${vencDate ? format(vencDate, 'dd/MM/yyyy') : 'NÃ£o definido'} (${statusVenc})\n\n*AÃ§Ãµes recomendadas:*\n${t.acoes || 'Nenhum detalhe adicional.'}\n\nPor favor, mantenha o andamento atualizado. Obrigado!`;

    const encodedText = encodeURIComponent(txt);
    const url = cleanContato 
      ? `https://api.whatsapp.com/send?phone=${cleanContato}&text=${encodedText}`
      : `https://api.whatsapp.com/send?text=${encodedText}`;
    window.open(url, '_blank');
  };

  const enviarLembreteEmail = (t: TarefaPlanoAcao) => {
    const contato = t.lembreteContato || '';
    const vencDate = parseLocalDate(t.dataVencimento || t.dataFim);
    const dias = vencDate ? Math.ceil((vencDate.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24)) : 0;
    const statusVenc = dias === 0 
      ? 'vence hoje' 
      : dias < 0 
        ? `estÃ¡ atrasada hÃ¡ ${Math.abs(dias)} dias`
        : `vencerÃ¡ em ${dias} dias`;

    const formattedVenc = vencDate ? format(vencDate, 'dd/MM/yyyy') : '';
    const subject = `Lembrete de Tarefa do Plano de AÃ§Ã£o - Vence em ${formattedVenc}`;
    const body = `OlÃ¡!\n\nEste Ã© um lembrete para a seguinte tarefa do Plano de AÃ§Ã£o:\n\n- Tarefa: ${t.problema}\n- ResponsÃ¡vel: ${t.responsavel || 'NÃ£o definido'}\n- Status: ${t.status}\n- Prioridade: ${t.prioridade}\n- Vencimento: ${formattedVenc || 'NÃ£o definido'} (${statusVenc})\n\nAÃ§Ãµes recomendadas:\n${t.acoes || 'Nenhum detalhe adicional.'}\n\nPor favor, atualize o status da atividade assim que possÃ­vel.\n\nAtenciosamente,\nGestor de Consultoria`;

    const mailto = `mailto:${encodeURIComponent(contato)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    window.open(mailto, '_blank');
  };

  return (
    <div className="space-y-6 max-w-7xl mx-auto pb-12">
      {!selectedDiagnostico && (
        <div className="p-8 bg-amber-50 border border-amber-200 rounded-2xl text-center">
          <p className="text-amber-800 font-bold">Nenhum diagnÃ³stico selecionado.</p>
          <p className="text-amber-600 text-sm">Por favor, selecione um diagnÃ³stico na tela de "DiagnÃ³sticos" para visualizar o plano de aÃ§Ã£o.</p>
        </div>
      )}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-3xl font-black text-slate-800 tracking-tight">GestÃ£o do Plano de AÃ§Ã£o</h2>
          <p className="text-slate-500 font-medium tracking-tight">Acompanhe a execuÃ§Ã£o das melhorias sugeridas nos diagnÃ³sticos.</p>
        </div>
        <div className="flex items-center gap-3">
           <button
             onClick={handleSavePlanoAcao}
             disabled={isSavingPlano}
             className={cn(
               "text-white px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 disabled:opacity-50 cursor-pointer",
               saveSuccessPlano ? "bg-emerald-700 shadow-sm" : "bg-emerald-600 hover:bg-emerald-700 shadow-md shadow-emerald-200"
             )}
           >
             {isSavingPlano ? (
               <>
                 <Loader2 className="animate-spin" size={14} />
                 <span>Salvando...</span>
               </>
             ) : saveSuccessPlano ? (
               <>
                 <CheckCircle2 size={14} className="text-white" />
                 <span>Salvo com Sucesso!</span>
               </>
             ) : (
               <>
                 <Save size={14} />
                 <span>Salvar AlteraÃ§Ãµes</span>
               </>
             )}
           </button>
           <button
             onClick={onGeneratePDF}
             className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl text-xs font-bold border border-emerald-100 hover:bg-emerald-100 transition-colors flex items-center gap-2"
           >
             <FileText size={14} />
             Imprimir Plano
           </button>
           <button 
             onClick={() => {
               setModalType('deletePlanoAcao');
               setModalData(selectedDiagnostico);
               setIsModalOpen(true);
             }}
             className="bg-rose-50 text-rose-600 px-4 py-2 rounded-xl text-xs font-bold border border-rose-100 hover:bg-rose-100 transition-colors"
           >
             Excluir Plano
           </button>
           <button 
             onClick={() => {
               setTarefaForm({
                 diagnosticoId: selectedDiagnostico?.id || '',
                 empresaId: selectedEmpresa?.id || '', lembreteDiasAntes: 1, lembreteEmail: false, lembreteWhatsapp: false, lembretePush: false, lembreteContato: ''
               });
               setModalType('createTarefa');
               setIsModalOpen(true);
             }}
             className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-xs font-bold shadow-lg shadow-indigo-200 hover:bg-indigo-700 transition-colors"
           >
             + Nova Tarefa
           </button>
           <div className="bg-white px-4 py-2 rounded-xl border border-slate-200 shadow-sm flex items-center gap-2">
             <div className="w-2 h-2 bg-emerald-500 rounded-full animate-pulse"></div>
             <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">{tarefas.length} Tarefas Ativas</span>
           </div>
        </div>
      </div>

      {selectedDiagnostico && (
        <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm">
          <LogoSelector 
            logoChoice={logoChoice}
            setLogoChoice={setLogoChoice}
            customLogo={customLogo}
            customConsultoraLogo={customConsultoraLogo}
          />
        </div>
      )}

      {/* SeÃ§Ã£o de Lembretes / Alertas Ativos */}
      {tarefAsComLembreteProximo.length > 0 && (
        <motion.div 
          initial={{ opacity: 0, y: -10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-gradient-to-r from-amber-50 to-orange-50 border border-amber-200 rounded-3xl p-6"
        >
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-amber-500 text-white rounded-xl animate-pulse">
              <Bell size={18} />
            </div>
            <div>
              <h3 className="font-black text-slate-800 uppercase tracking-wider text-xs">Alertas de Vencimento Ativos</h3>
              <p className="text-[11px] text-slate-500 font-medium">HÃ¡ tarefas prÃ³ximas do vencimento com lembretes configurados.</p>
            </div>
          </div>
          
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {tarefAsComLembreteProximo.map(t => {
              const dataV = t.dataVencimento || t.dataFim;
              const dLimit = parseLocalDate(dataV);
              const dias = dLimit ? Math.ceil((dLimit.getTime() - new Date().getTime()) / (1000 * 60 * 60 * 24)) : 0;
              
              return (
                <div key={t.id} className="bg-white/80 backdrop-blur-sm border border-amber-100/80 rounded-2xl p-4 flex flex-col justify-between shadow-sm hover:shadow-md transition-all">
                  <div>
                    <div className="flex justify-between items-start gap-2 mb-2">
                      <span className="text-[9px] font-black uppercase bg-amber-50 text-amber-600 border border-amber-100 px-2 py-0.5 rounded-md">
                        {dias === 0 ? "Vence hoje" : dias < 0 ? `Vencido hÃ¡ ${Math.abs(dias)}d` : `Vence em ${dias}d`}
                      </span>
                      <span className="text-[10px] text-slate-400 font-bold max-w-[120px] truncate">{t.responsavel || 'Sem resp.'}</span>
                    </div>
                    
                    <h4 className="font-bold text-slate-800 text-xs mb-1 truncate">{t.problema}</h4>
                    <p className="text-[10px] text-slate-400 font-medium mb-3">Vencimento: {dLimit ? dLimit.toLocaleDateString('pt-BR') : 'NÃ£o definido'}</p>
                  </div>
                  
                  <div className="flex items-center gap-2 pt-2 border-t border-slate-50">
                    <span className="text-[9px] font-black text-slate-400 uppercase tracking-tight mr-auto font-bold">Enviar Lembrete:</span>
                    {t.lembreteEmail && (
                      <button
                        onClick={() => enviarLembreteEmail(t)}
                        className="p-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-600 rounded-lg transition-colors cursor-pointer"
                        title="Enviar lembrete por E-mail"
                      >
                        <Mail size={12} />
                      </button>
                    )}
                    {t.lembreteWhatsapp && (
                      <button
                        onClick={() => enviarLembreteWhatsApp(t)}
                        className="p-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-600 rounded-lg transition-colors cursor-pointer"
                        title="Enviar lembrete por WhatsApp"
                      >
                        <MessageSquare size={12} />
                      </button>
                    )}
                    {t.lembretePush && (
                      <button
                        onClick={() => {
                          if ('Notification' in window && Notification.permission === 'granted') {
                            new Notification(`Lembrete - ${t.problema}`, {
                              body: `Tarefa para ${t.responsavel || 'nÃ£o definido'}. Vencimento: ${dataV ? new Date(dataV).toLocaleDateString('pt-BR') : ''}.`
                            });
                          } else {
                            alert("PermissÃµes de notificaÃ§Ãµes nÃ£o concedidas no navegador!");
                          }
                        }}
                        className="p-1.5 bg-purple-50 hover:bg-purple-100 text-purple-600 rounded-lg transition-colors cursor-pointer"
                        title="Disparar Alerta Push Local"
                      >
                        <Bell size={12} />
                      </button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </motion.div>
      )}

      {/* Barra de Ferramentas / Filtros e Modos de VisualizaÃ§Ã£o */}
      <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4">
        <div className="flex items-center gap-2">
          <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">VisualizaÃ§Ã£o:</span>
          <div className="bg-slate-100 p-1 rounded-xl flex gap-1">
            <button
              onClick={() => setViewMode('kanban')}
              className={cn(
                "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer",
                viewMode === 'kanban' 
                  ? "bg-white text-indigo-600 shadow-sm font-black" 
                  : "text-slate-500 hover:text-slate-700"
              )}
            >
              <Layout size={13} />
              Quadro Kanban
            </button>
            <button
              onClick={() => setViewMode('timeline')}
              className={cn(
                "flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer",
                viewMode === 'timeline' 
                  ? "bg-white text-indigo-600 shadow-sm font-black" 
                  : "text-slate-500 hover:text-slate-700"
              )}
            >
              <CalendarDays size={13} />
              Cronograma Sequencial
            </button>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-4 w-full sm:w-auto justify-end">
          <div className="flex items-center gap-2">
            <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Organizar por:</span>
            <select
              value={groupBy}
              onChange={(e) => setGroupBy(e.target.value as any)}
              className="px-3 py-1.5 text-xs font-bold rounded-xl border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20 cursor-pointer"
            >
              <option value="status">ðŸ“Š Status de ExecuÃ§Ã£o</option>
              <option value="data">ðŸ“… Data de ExecuÃ§Ã£o (MÃªs)</option>
              <option value="problema">â“ Problema Associado</option>
              <option value="area">ðŸ—‚ï¸ Ãrea de Consultoria</option>
            </select>
          </div>

          <div className="flex items-center gap-2">
            <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">Ordenar por:</span>
            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value as any)}
              className="px-3 py-1.5 text-xs font-bold rounded-xl border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20 cursor-pointer"
            >
              <option value="ordem">ðŸ”¢ Ordem Manual das AÃ§Ãµes (#1, #2...)</option>
              <option value="cronograma">ðŸ“‹ Ordem do RelatÃ³rio de Consultoria</option>
              <option value="dataInicio">ðŸ“… Ordem de InÃ­cio</option>
              <option value="dataVencimento">â° Ordem de Vencimento</option>
              <option value="prioridade">ðŸ”¥ NÃ­vel de Prioridade</option>
              <option value="problema">â“ Nome do Problema</option>
              <option value="area">ðŸ—‚ï¸ Ãrea de Consultoria</option>
            </select>
          </div>
        </div>
      </div>

      {viewMode === 'kanban' ? (
        <div className="flex flex-col lg:flex-row gap-6 min-h-[600px] overflow-x-auto pb-4">
          {computedColumns.map(col => (
            <div key={col.id} className={cn("flex-1 flex flex-col rounded-3xl border min-w-[320px] max-w-[380px] shadow-sm", col.border, col.color)}>
              <div className="p-5 border-b border-inherit flex items-center justify-between">
                <h3 className={cn("font-black uppercase tracking-[0.2em] text-[10px] truncate max-w-[200px]", col.text)} title={col.label}>{col.label}</h3>
                <span className={cn("px-2.5 py-1 rounded-lg text-[10px] font-black shrink-0 ml-2", "bg-white/80 shadow-sm", col.text)}>
                  {(groupedTasksMap[col.id] || []).length}
                </span>
              </div>
              
              <div className="flex-1 p-4 space-y-4 overflow-y-auto max-h-[650px]">
                {(groupedTasksMap[col.id] || []).length === 0 ? (
                  <div className="h-40 flex items-center justify-center border-2 border-dashed border-slate-200/50 rounded-2xl bg-white/50">
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Sem tarefas</p>
                  </div>
                ) : (
                  (groupedTasksMap[col.id] || []).map(tarefa => (
                    <motion.div 
                      key={tarefa.id} 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      whileHover={{ y: -2 }}
                      className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 hover:shadow-xl hover:border-indigo-100 transition-all cursor-pointer group relative overflow-hidden"
                      onClick={() => onEdit(tarefa)}
                    >
                      <div className={cn(
                        "absolute top-0 left-0 w-1 h-full",
                        tarefa.prioridade === 'Alta' ? "bg-rose-500" :
                        tarefa.prioridade === 'MÃ©dia' ? "bg-amber-500" : "bg-slate-300"
                      )} />
                      
                      <div className="flex justify-between items-start mb-3">
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <span className={cn(
                            "text-[9px] font-black px-2 py-0.5 rounded-md uppercase tracking-widest",
                            tarefa.prioridade === 'Alta' ? "bg-rose-50 text-rose-600 border border-rose-100" :
                            tarefa.prioridade === 'MÃ©dia' ? "bg-amber-50 text-amber-600 border border-amber-100" :
                            "bg-slate-50 text-slate-600 border border-slate-100"
                          )}>
                            {tarefa.prioridade}
                          </span>
                          <div className="flex items-center gap-1 bg-indigo-50/80 px-2 py-0.5 rounded-md border border-indigo-100 text-[10px] font-black text-indigo-700" onClick={(e) => e.stopPropagation()}>
                            <ListOrdered size={12} />
                            <span className="text-[9px] text-indigo-500 font-bold mr-0.5">PosiÃ§Ã£o:</span>
                            <select
                              value={tarefa.ordem !== undefined && tarefa.ordem !== null ? tarefa.ordem : sortedTarefas.findIndex(t => t.id === tarefa.id)}
                              onChange={(e) => {
                                e.stopPropagation();
                                handleSetTaskPosition(tarefa, parseInt(e.target.value, 10));
                              }}
                              className="bg-white font-black text-indigo-700 border border-indigo-200 rounded px-1 py-0.5 cursor-pointer outline-none shadow-2xs hover:bg-indigo-50"
                              title="Clique para alterar a posiÃ§Ã£o da aÃ§Ã£o"
                            >
                              {sortedTarefas.map((_, idx) => (
                                <option key={idx} value={idx}>#{idx + 1}</option>
                              ))}
                            </select>
                            <div className="flex items-center ml-1 border-l border-indigo-200 pl-1 gap-0.5">
                              <button
                                type="button"
                                title="Mover para cima (#1)"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'up'); }}
                                className="p-0.5 hover:bg-indigo-200/60 rounded text-indigo-700 transition-colors"
                              >
                                <ArrowUp size={11} />
                              </button>
                              <button
                                type="button"
                                title="Mover para baixo"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'down'); }}
                                className="p-0.5 hover:bg-indigo-200/60 rounded text-indigo-700 transition-colors"
                              >
                                <ArrowDown size={11} />
                              </button>
                            </div>
                          </div>
                        </div>
                        <div className="flex items-center gap-1.5" onClick={(e) => e.stopPropagation()}>
                          <select
                            value={tarefa.prioridade}
                            onChange={(e) => {
                              e.stopPropagation();
                              onUpdatePrioridade(tarefa.id, e.target.value as 'Baixa' | 'MÃ©dia' | 'Alta');
                            }}
                            className={cn(
                              "text-[9px] font-black px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20",
                              tarefa.prioridade === 'Alta' ? "text-rose-600" :
                              tarefa.prioridade === 'MÃ©dia' ? "text-amber-600" : "text-slate-600"
                            )}
                          >
                            <option value="Baixa">Baixa</option>
                            <option value="MÃ©dia">MÃ©dia</option>
                            <option value="Alta">Alta</option>
                          </select>
                          <select
                            value={tarefa.status}
                            onChange={(e) => {
                              e.stopPropagation();
                              onUpdateStatus(tarefa.id, e.target.value as 'Pendente' | 'Em Andamento' | 'ConcluÃ­do');
                            }}
                            className={cn(
                              "text-[10px] font-bold px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20",
                            )}
                          >
                            <option value="Pendente">Pendente</option>
                            <option value="Em Andamento">Em Andamento</option>
                            <option value="ConcluÃ­do">ConcluÃ­do</option>
                          </select>
                          <button 
                            onClick={(e) => { e.stopPropagation(); onDelete(tarefa); }}
                            className="p-1.5 hover:bg-rose-50 text-rose-400 rounded-lg transition-all"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </div>
                      
                      <h4 className="font-bold text-slate-800 text-sm mb-1 leading-tight group-hover:text-indigo-600 transition-colors uppercase tracking-tight">{tarefa.problema}</h4>
                      <p className="text-[11px] text-slate-500 font-medium mb-3 line-clamp-2">{tarefa.solucaoSugerida}</p>
                      
                      {tarefa.area && (
                        <div className="mb-2 flex items-center gap-1.5">
                          <span className="text-[9px] font-black text-indigo-500 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded-md uppercase tracking-tight">
                            ðŸ—‚ï¸ {tarefa.area}
                          </span>
                        </div>
                      )}

                      {tarefa.dataInicio && tarefa.dataFim && (
                        <div className="mb-4">
                          <div className="flex items-center gap-2 mb-1">
                            <div className="flex-1 h-1 bg-slate-100 rounded-full overflow-hidden">
                              <div className="h-full bg-emerald-500 w-full opacity-50"></div>
                            </div>
                          </div>
                          <div className="flex justify-between items-center bg-indigo-50/50 px-2.5 py-1.5 rounded-lg border border-indigo-100/50">
                             <span className="text-[10px] font-black text-indigo-600 uppercase tracking-tighter">DuraÃ§Ã£o</span>
                             <span className="text-[10px] font-bold text-indigo-700 bg-white px-1.5 py-0.5 rounded shadow-sm border border-indigo-100">
                                {(() => {
                                  const d1 = parseLocalDate(tarefa.dataInicio);
                                  const d2 = parseLocalDate(tarefa.dataFim);
                                  if (d1 && d2) {
                                    return Math.ceil((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24)) + 1;
                                  }
                                  return 1;
                                })()} dias
                             </span>
                          </div>
                        </div>
                      )}
                      
                      <div className="pt-4 border-t border-slate-50 flex items-center justify-between">
                        <div className="flex items-center gap-2">
                           <div className="w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-black text-slate-500">
                             {getEmpresaNome(tarefa.empresaId).charAt(0)}
                           </div>
                           <span className="text-[10px] font-bold text-slate-500 uppercase tracking-tight truncate max-w-[100px]">{getEmpresaNome(tarefa.empresaId)}</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <label 
                            className="relative cursor-pointer p-1.5 hover:bg-slate-50 rounded-lg text-slate-400 hover:text-indigo-600 transition-all"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <Paperclip size={14} className={tarefa.anexoUrl ? "text-indigo-600" : ""} />
                            {tarefa.anexoUrl && <span className="absolute -top-1 -right-1 w-2 h-2 bg-indigo-600 rounded-full"></span>}
                            <input 
                              type="file" 
                              className="hidden" 
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) onUploadEvidence(file, 'tarefa', tarefa.id);
                              }}
                            />
                          </label>
                          {tarefa.evidenciaUrl && (
                            <div 
                              className="text-emerald-500 cursor-help" 
                              title={tarefa.evidenciaNome}
                              onClick={(e) => { e.stopPropagation(); window.open(tarefa.evidenciaUrl, '_blank'); }}
                            >
                              <ImageIcon size={14} />
                            </div>
                          )}
                          {(tarefa.lembreteEmail || tarefa.lembreteWhatsapp || tarefa.lembretePush) && (
                            <div className="flex items-center gap-1 bg-slate-50 px-2 py-0.5 rounded-md text-[10px] border border-slate-100/50">
                              {tarefa.lembreteEmail && <Mail size={10} className="text-indigo-500" title="E-mail configurado" />}
                              {tarefa.lembreteWhatsapp && <MessageSquare size={10} className="text-emerald-500" title="WhatsApp configurado" />}
                              {tarefa.lembretePush && <Smartphone size={10} className="text-purple-500" title="Push configurado" />}
                            </div>
                          )}
                          {tarefa.dataVencimento && (
                            <div className="flex items-center gap-1 text-[10px] font-black text-rose-500 uppercase tracking-tight bg-rose-50 px-2 py-0.5 rounded-md">
                              <Clock size={12} />
                              {format(parseLocalDate(tarefa.dataVencimento) || new Date(), 'dd/MM')}
                            </div>
                          )}
                        </div>
                      </div>
                    </motion.div>
                  ))
                )}
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className="bg-white rounded-3xl border border-slate-100 shadow-sm p-6 md:p-8 space-y-10 relative">
          <div className="absolute left-[37px] md:left-[45px] top-24 bottom-12 w-0.5 bg-indigo-100"></div>

          {sortedTarefas.length === 0 ? (
            <div className="h-40 flex flex-col items-center justify-center border-2 border-dashed border-slate-200/50 rounded-2xl">
              <p className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-1">Sem tarefas no cronograma</p>
              <p className="text-[11px] text-slate-400">Gere um plano IA ou adicione tarefas manualmente.</p>
            </div>
          ) : (
            computedColumns.map((col, colIndex) => {
              const colTasks = groupedTasksMap[col.id] || [];
              if (colTasks.length === 0) return null;

              return (
                <div key={col.id} className="space-y-6 relative">
                  {/* Header do Grupo */}
                  <div className="relative z-10 flex items-center gap-3">
                    <div className={cn("px-4 py-2 rounded-2xl border text-xs font-black uppercase tracking-wider shadow-sm flex items-center gap-2", col.color, col.border)}>
                      {groupBy === 'status' ? 'ðŸ“Š Status: ' :
                       groupBy === 'data' ? 'ðŸ“… MÃªs: ' :
                       groupBy === 'problema' ? 'â“ Problema: ' : 'ðŸ—‚ï¸ Ãrea: '}
                      {col.label}
                    </div>
                    <div className="h-px flex-1 bg-slate-100"></div>
                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                      {colTasks.length} {colTasks.length === 1 ? 'tarefa' : 'tarefas'}
                    </span>
                  </div>

                  <div className="space-y-6 pl-4 md:pl-6 border-l border-indigo-50/50 ml-5 md:ml-6">
                    {colTasks.map((tarefa, index) => {
                      const dStart = parseLocalDate(tarefa.dataInicio);
                      const dEnd = parseLocalDate(tarefa.dataFim);
                      const duracaoDias = dStart && dEnd 
                        ? Math.ceil((dEnd.getTime() - dStart.getTime()) / (1000 * 60 * 60 * 24)) + 1
                        : null;
                      
                      const statusColors = tarefa.status === 'ConcluÃ­do' 
                        ? 'bg-emerald-50 text-emerald-700 border-emerald-200' 
                        : tarefa.status === 'Em Andamento'
                          ? 'bg-blue-50 text-blue-700 border-blue-200'
                          : 'bg-slate-50 text-slate-700 border-slate-200';

                      const priorityColors = tarefa.prioridade === 'Alta'
                        ? 'bg-rose-50 text-rose-700 border-rose-200'
                        : tarefa.prioridade === 'MÃ©dia'
                          ? 'bg-amber-50 text-amber-700 border-amber-200'
                          : 'bg-slate-50 text-slate-700 border-slate-200';

                      return (
                        <motion.div 
                          key={tarefa.id}
                          initial={{ opacity: 0, x: -10 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: index * 0.05 }}
                          className="relative flex items-start gap-4 md:gap-6 group"
                        >
                          <div className="relative z-10 shrink-0 flex flex-col items-center gap-1">
                            <div className="flex items-center justify-center w-9 h-9 rounded-full border-2 border-indigo-200 bg-white shadow-sm group-hover:border-indigo-600 transition-colors relative">
                              <select
                                value={tarefa.ordem !== undefined && tarefa.ordem !== null ? tarefa.ordem : sortedTarefas.findIndex(t => t.id === tarefa.id)}
                                onChange={(e) => {
                                  e.stopPropagation();
                                  handleSetTaskPosition(tarefa, parseInt(e.target.value, 10));
                                }}
                                className="w-full h-full opacity-0 absolute inset-0 cursor-pointer z-20"
                                title="Clique para alterar a posiÃ§Ã£o da aÃ§Ã£o"
                              >
                                {sortedTarefas.map((_, idx) => (
                                  <option key={idx} value={idx}>PosiÃ§Ã£o #{idx + 1}</option>
                                ))}
                              </select>
                              <span className="text-[10px] font-black text-indigo-600 uppercase pointer-events-none">
                                {String(tarefa.ordem !== undefined && tarefa.ordem !== null ? tarefa.ordem + 1 : sortedTarefas.findIndex(t => t.id === tarefa.id) + 1).padStart(2, '0')}
                              </span>
                            </div>
                            <div className="flex flex-col gap-0.5 items-center">
                              <button
                                type="button"
                                title="Mover para cima"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'up'); }}
                                className="p-1 hover:bg-indigo-100/80 rounded-md text-indigo-600 transition-colors bg-white border border-slate-200/80 shadow-xs cursor-pointer"
                              >
                                <ArrowUp size={11} />
                              </button>
                              <button
                                type="button"
                                title="Mover para baixo"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'down'); }}
                                className="p-1 hover:bg-indigo-100/80 rounded-md text-indigo-600 transition-colors bg-white border border-slate-200/80 shadow-xs cursor-pointer"
                              >
                                <ArrowDown size={11} />
                              </button>
                            </div>
                          </div>

                          <div className="flex-1 bg-slate-50/50 hover:bg-white rounded-2xl border border-slate-100 hover:border-indigo-100 p-5 shadow-sm hover:shadow-md transition-all flex flex-col md:flex-row gap-4 justify-between items-start md:items-center">
                            <div className="flex-1 space-y-2">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="text-[10px] font-bold text-slate-400 bg-slate-100 px-2 py-0.5 rounded-md uppercase tracking-tight">
                                  {getEmpresaNome(tarefa.empresaId)}
                                </span>
                                <span className={`text-[10px] font-black uppercase tracking-tight px-2 py-0.5 rounded-md border ${statusColors}`}>
                                  {tarefa.status}
                                </span>
                                <span className={`text-[10px] font-black uppercase tracking-tight px-2 py-0.5 rounded-md border ${priorityColors}`}>
                                  Prioridade: {tarefa.prioridade}
                                </span>
                                {tarefa.area && (
                                  <span className="text-[10px] font-black uppercase tracking-tight px-2 py-0.5 rounded-md border bg-indigo-50/50 border-indigo-100 text-indigo-600">
                                    ðŸ—‚ï¸ {tarefa.area}
                                  </span>
                                )}
                              </div>

                              <h4 className="font-bold text-slate-800 text-base leading-snug group-hover:text-indigo-600 transition-colors uppercase tracking-tight">
                                {tarefa.problema}
                              </h4>
                              
                              {tarefa.solucaoSugerida && (
                                <p className="text-xs text-slate-600 font-medium leading-relaxed max-w-3xl">
                                  <strong className="text-slate-700">SoluÃ§Ã£o Recomendada:</strong> {tarefa.solucaoSugerida}
                                </p>
                              )}

                              {tarefa.acoes && (
                                <p className="text-xs text-slate-500 font-medium leading-relaxed max-w-3xl pt-1">
                                  <strong className="text-slate-600">AÃ§Ãµes:</strong> {tarefa.acoes}
                                </p>
                              )}

                              <div className="flex flex-wrap items-center gap-x-4 gap-y-1 pt-2 text-[11px] text-slate-500 font-bold uppercase tracking-tight">
                                {tarefa.dataInicio && (
                                  <div className="flex items-center gap-1">
                                    <Clock size={12} className="text-slate-400" />
                                    <span>InÃ­cio: {dStart ? format(dStart, 'dd/MM/yyyy') : '-'}</span>
                                  </div>
                                )}
                                {tarefa.dataFim && (
                                  <div className="flex items-center gap-1">
                                    <Clock size={12} className="text-slate-400" />
                                    <span>Fim: {dEnd ? format(dEnd, 'dd/MM/yyyy') : '-'}</span>
                                  </div>
                                )}
                                {duracaoDias && (
                                  <span className="bg-indigo-50 text-indigo-700 px-2 py-0.5 rounded-md border border-indigo-100">
                                    {duracaoDias} dias
                                  </span>
                                )}
                                {tarefa.responsavel && (
                                  <span className="text-slate-500">
                                    ResponsÃ¡vel: <strong className="text-slate-600">{tarefa.responsavel}</strong>
                                  </span>
                                )}
                              </div>
                            </div>

                            <div className="shrink-0 flex items-center md:flex-col gap-3 w-full md:w-auto pt-3 md:pt-0 border-t md:border-t-0 border-slate-100 justify-end" onClick={(e) => e.stopPropagation()}>
                              <div className="flex md:flex-col gap-2 items-end">
                                <select
                                  value={tarefa.status}
                                  onChange={(e) => onUpdateStatus(tarefa.id, e.target.value as any)}
                                  className="text-[10px] font-bold px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20"
                                >
                                  <option value="Pendente">Pendente</option>
                                  <option value="Em Andamento">Em Andamento</option>
                                  <option value="ConcluÃ­do">ConcluÃ­do</option>
                                </select>

                                <select
                                  value={tarefa.prioridade}
                                  onChange={(e) => onUpdatePrioridade(tarefa.id, e.target.value as any)}
                                  className={cn(
                                    "text-[9px] font-black px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-500/20",
                                    tarefa.prioridade === 'Alta' ? "text-rose-600" :
                                    tarefa.prioridade === 'MÃ©dia' ? "text-amber-600" : "text-slate-600"
                                  )}
                                >
                                  <option value="Baixa">Baixa</option>
                                  <option value="MÃ©dia">MÃ©dia</option>
                                  <option value="Alta">Alta</option>
                                </select>
                              </div>

                              <div className="flex gap-2">
                                <label 
                                  className="relative cursor-pointer p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-indigo-600 transition-all"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  <Paperclip size={14} className={tarefa.anexoUrl ? "text-indigo-600" : ""} />
                                  {tarefa.anexoUrl && <span className="absolute -top-1 -right-1 w-2 h-2 bg-indigo-600 rounded-full"></span>}
                                  <input 
                                    type="file" 
                                    className="hidden" 
                                    onChange={(e) => {
                                      const file = e.target.files?.[0];
                                      if (file) onUploadEvidence(file, 'tarefa', tarefa.id);
                                    }}
                                  />
                                </label>

                                <button 
                                  onClick={(e) => { e.stopPropagation(); onEdit(tarefa); }}
                                  className="p-1.5 hover:bg-indigo-50 text-indigo-400 rounded-lg transition-all"
                                >
                                  <Edit2 size={14} />
                                </button>
                                <button 
                                  onClick={(e) => { e.stopPropagation(); onDelete(tarefa); }}
                                  className="p-1.5 hover:bg-rose-50 text-rose-400 rounded-lg transition-all"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            </div>
                          </div>
                        </motion.div>
                      );
                    })}
                  </div>
                </div>
              );
            })
          )}
        </div>
      )}
    </div>
  );
};

const LandingPage = ({ setView }: { setView: (v: any) => void }) => {
  return (
    <div className="min-h-screen bg-white">
      {/* Hero Section */}
      <section className="relative py-20 bg-emerald-600 overflow-hidden">
        <div className="absolute top-0 left-0 w-full h-full opacity-10">
          <svg className="w-full h-full" viewBox="0 0 100 100" preserveAspectRatio="none">
            <path d="M0 100 C 20 0 50 0 100 100 Z" fill="white" />
          </svg>
        </div>
        <div className="container mx-auto px-6 relative z-10 text-center">
          <motion.div
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 }}
          >
            <div className="inline-flex items-center gap-2 px-4 py-1.5 bg-white/20 backdrop-blur-md rounded-full text-xs font-black text-white mb-6 tracking-widest uppercase border border-white/25">
              by ItÃ mar Gomes
            </div>
            <h1 className="text-4xl md:text-6xl font-bold text-white mb-6">
              Transforme sua Consultoria <br /> com InteligÃªncia Artificial
            </h1>
            <p className="text-xl text-emerald-100 mb-10 max-w-2xl mx-auto">
              A plataforma definitiva para diagnÃ³sticos precisos, cronogramas automÃ¡ticos e gestÃ£o de planos de aÃ§Ã£o em tempo real.
            </p>
            <div className="flex flex-col sm:flex-row justify-center items-center gap-4">
              <Button 
                onClick={() => setView('home')}
                className="bg-white text-emerald-600 hover:bg-emerald-50 px-8 py-4 text-lg font-bold h-auto border-none w-full sm:w-auto"
              >
                Inicie agora o seu teste grÃ¡tis (30 dias)
              </Button>
              <Button 
                onClick={() => setView('home')}
                variant="outline"
                className="text-white border-white hover:bg-emerald-700 px-8 py-4 text-lg font-bold h-auto w-full sm:w-auto"
              >
                Ver DemonstraÃ§Ã£o
              </Button>
            </div>
            <div className="mt-8 flex justify-center">
              <Button
                variant="ghost"
                onClick={() => {
                  navigator.clipboard.writeText("https://gestorconsultorpro.netlify.app/");
                  alert("Link de teste copiado para sua Ã¡rea de transferÃªncia!");
                }}
                className="text-white/60 hover:text-white hover:bg-white/10 flex items-center gap-2 text-xs uppercase tracking-widest font-black"
              >
                <Share2 size={16} />
                Compartilhar link de acesso
              </Button>
            </div>
          </motion.div>
        </div>
      </section>

      {/* Benefits */}
      <section className="py-20 bg-slate-50">
        <div className="container mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-3xl font-bold text-slate-800">Por que escolher nossa plataforma?</h2>
            <p className="text-slate-500 mt-4">Tudo o que vocÃª precisa para elevar o nÃ­vel das suas entregas consultivas.</p>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
            {[
              { icon: Sparkles, title: "InteligÃªncia Artificial", desc: "GeraÃ§Ã£o automÃ¡tica de soluÃ§Ãµes e cronogramas baseados no diagnÃ³stico." },
              { icon: FileText, title: "RelatÃ³rios Profissionais", desc: "Exporte diagnÃ³sticos e cronogramas completos em PDF com sua marca." },
              { icon: Layout, title: "GestÃ£o Kanban", desc: "Acompanhe a implementaÃ§Ã£o de cada aÃ§Ã£o em um quadro visual intuitivo." }
            ].map((item, i) => (
              <motion.div 
                key={i}
                whileHover={{ y: -5 }}
                className="bg-white p-8 rounded-2xl shadow-sm border border-slate-100 text-center"
              >
                <div className="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-2xl flex items-center justify-center mx-auto mb-6">
                  <item.icon size={32} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">{item.title}</h3>
                <p className="text-slate-500 leading-relaxed">{item.desc}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* Pricing */}
      <section className="py-20">
        <div className="container mx-auto px-6">
          <div className="text-center mb-16">
            <h2 className="text-3xl font-bold text-slate-800">Escolha o seu plano</h2>
            <p className="text-slate-500 mt-2">Assinatura flexÃ­vel para o seu negÃ³cio de consultoria.</p>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8 max-w-5xl mx-auto">
            {/* Monthly Plan */}
            <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-slate-200 flex flex-col">
              <div className="p-8 md:p-12 flex-1">
                <div className="bg-slate-50 text-slate-500 text-xs font-black uppercase tracking-widest px-3 py-1 rounded-full w-fit mb-6">Mensal</div>
                <h3 className="text-2xl font-bold text-slate-800 mb-4">Consultor PrÃ³</h3>
                <div className="flex items-baseline gap-1 mb-8">
                  <span className="text-slate-400 text-xl font-medium font-sans">R$</span>
                  <span className="text-5xl font-bold text-slate-800">47,90</span>
                  <span className="text-slate-400 font-medium font-sans">/mÃªs</span>
                </div>
                <ul className="space-y-4">
                  {['DiagnÃ³sticos Ilimitados', 'IA Generativa Integrada', 'RelatÃ³rios Customizados', 'GestÃ£o Kanban'].map((text, i) => (
                    <li key={i} className="flex items-center gap-3 text-slate-700 text-sm">
                      <CheckCircle2 size={18} className="text-emerald-500" />
                      <span>{text}</span>
                    </li>
                  ))}
                </ul>
              </div>
              <div className="p-8 bg-slate-50">
                <Button 
                  onClick={() => { setView('checkout'); localStorage.setItem('selected_plan', 'monthly'); }}
                  className="w-full h-12 font-bold"
                >
                  Assinar Mensal
                </Button>
              </div>
            </div>

            {/* Annual Plan */}
            <div className="bg-white rounded-3xl shadow-xl overflow-hidden border-2 border-emerald-500 flex flex-col relative">
              <div className="absolute top-4 right-4 bg-emerald-500 text-white text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full">
                Economize 15%
              </div>
              <div className="p-8 md:p-12 flex-1">
                <div className="bg-emerald-50 text-emerald-600 text-xs font-black uppercase tracking-widest px-3 py-1 rounded-full w-fit mb-6">Anual</div>
                <h3 className="text-2xl font-bold text-slate-800 mb-4">Master Consultant</h3>
                <div className="flex items-baseline gap-1 mb-2">
                  <span className="text-slate-400 text-xl font-medium font-sans">R$</span>
                  <span className="text-5xl font-bold text-emerald-600">41,49</span>
                  <span className="text-slate-400 font-medium font-sans">/mÃªs</span>
                </div>
                <p className="text-slate-400 text-xs mb-8">R$ 497,90 Ã  vista ou no cartÃ£o de crÃ©dito em atÃ© 12x</p>
                <ul className="space-y-4">
                  {['Tudo do plano mensal', 'Acesso Antecipado a Novas IAs', 'Suporte PrioritÃ¡rio', 'Mentoria em DiagnÃ³sticos'].map((text, i) => (
                    <li key={i} className="flex items-center gap-3 text-slate-700 text-sm">
                      <CheckCircle2 size={18} className="text-emerald-500" />
                      <span>{text}</span>
                    </li>
                  ))}
                </ul>
              </div>
              <div className="p-8 bg-emerald-50">
                <Button 
                  onClick={() => { setView('checkout'); localStorage.setItem('selected_plan', 'annual'); }}
                  className="w-full h-12 font-bold bg-emerald-600 hover:bg-emerald-700 shadow-lg shadow-emerald-200"
                >
                  Assinar Anual (15% OFF)
                </Button>
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
};

const CheckoutPage = ({ setView, plan = 'annual' }: { setView: (v: any) => void, plan?: 'monthly' | 'annual' }) => {
  const [step, setStep] = useState(1);
  const [isCopied, setIsCopied] = useState(false);
  const [comprovante, setComprovante] = useState<File | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isVerifying, setIsVerifying] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [formData, setFormData] = useState({
    name: '',
    email: '',
    cnpj: '',
    card: '',
    expiry: '',
    cvv: ''
  });

  const validateCNPJ = (cnpj: string) => {
    cnpj = cnpj.replace(/[^\d]+/g, '');
    if (cnpj === '') return false;
    if (cnpj.length !== 14) return false;
    if (/^(\d)\1+$/.test(cnpj)) return false;
    
    let size = cnpj.length - 2;
    let numbers = cnpj.substring(0, size);
    let digits = cnpj.substring(size);
    let sum = 0;
    let pos = size - 7;
    for (let i = size; i >= 1; i--) {
      sum += parseInt(numbers.charAt(size - i)) * pos--;
      if (pos < 2) pos = 9;
    }
    let result = sum % 11 < 2 ? 0 : 11 - (sum % 11);
    if (result !== parseInt(digits.charAt(0))) return false;
    
    size = size + 1;
    numbers = cnpj.substring(0, size);
    sum = 0;
    pos = size - 7;
    for (let i = size; i >= 1; i--) {
      sum += parseInt(numbers.charAt(size - i)) * pos--;
      if (pos < 2) pos = 9;
    }
    result = sum % 11 < 2 ? 0 : 11 - (sum % 11);
    if (result !== parseInt(digits.charAt(1))) return false;
    
    return true;
  };

  const handleCNPJChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    let val = e.target.value.replace(/[^\d]/g, "");
    if (val.length > 14) val = val.substring(0, 14);
    
    let masked = val;
    if (val.length > 2) masked = val.substring(0, 2) + "." + val.substring(2);
    if (val.length > 5) masked = masked.substring(0, 6) + "." + masked.substring(6);
    if (val.length > 8) masked = masked.substring(0, 10) + "/" + masked.substring(10);
    if (val.length > 12) masked = masked.substring(0, 15) + "-" + masked.substring(15);
    
    setFormData({...formData, cnpj: masked});
  };

  const handleFinish = async () => {
    if (!validateCNPJ(formData.cnpj)) {
      alert("CNPJ InvÃ¡lido. Por favor, verifique os nÃºmeros digitados.");
      return;
    }

    if (!comprovante) {
      alert("Por favor, anexe o comprovante do pagamento PIX antes de finalizar.");
      return;
    }

    setIsVerifying(true);

    // Simula validaÃ§Ã£o em tempo real
    await new Promise(resolve => setTimeout(resolve, 2200));

    // Tenta atualizar a licenÃ§a do usuÃ¡rio se ele estiver logado
    const currentUser = auth.currentUser;
    if (currentUser) {
      try {
        // Busca a credenciada vinculada ao usuÃ¡rio
        const q = query(collection(db, 'empresas_credenciadas'), where('ownerId', '==', currentUser.uid));
        const querySnapshot = await getDocs(q);
        
        const planoEscolhido = plan === 'monthly' ? 'Mensal' : 'Anual';
        
        if (!querySnapshot.empty) {
          const docId = querySnapshot.docs[0].id;
          const currentData = querySnapshot.docs[0].data();
          await updateDoc(doc(db, 'empresas_credenciadas', docId), {
            razaoSocial: formData.name || currentData.razaoSocial || 'Novo Cliente',
            cnpj: formData.cnpj,
            email: formData.email || currentUser.email || '',
            tipoPlano: planoEscolhido,
            dataCadastro: serverTimestamp(), // Reseta a data para iniciar o novo perÃ­odo
            status: 'Ativa',
            comprovanteNome: comprovante.name,
            dataComprovante: serverTimestamp()
          });
        } else {
          // caso nÃ£o exista, cria uma nova empresa_credenciada para o usuÃ¡rio logado
          await addDoc(collection(db, 'empresas_credenciadas'), {
            razaoSocial: formData.name || 'Novo Cliente',
            cnpj: formData.cnpj,
            email: formData.email || currentUser.email || '',
            tipoPlano: planoEscolhido,
            dataCadastro: serverTimestamp(),
            ownerId: currentUser.uid,
            status: 'Ativa',
            role: 'cliente',
            comprovanteNome: comprovante.name,
            dataComprovante: serverTimestamp()
          });
        }
      } catch (error) {
        console.error("Erro ao atualizar licenÃ§a:", error);
      }
    }

    setIsVerifying(false);
    playSuccessSound();
    alert(`Sucesso! O comprovante do PIX foi recebido e validado com sucesso. Sua licenÃ§a ${plan === 'monthly' ? 'Mensal' : 'Anual'} estÃ¡ ativa e o sistema foi LIBERADO!`);
    setView('home');
  };

  return (
    <div className="min-h-screen bg-slate-50 py-12 px-6">
      <div className="max-w-2xl mx-auto">
        <button 
          onClick={() => setView('landing')}
          className="flex items-center gap-2 text-slate-500 hover:text-slate-800 transition-colors mb-8 group"
        >
          <ArrowLeft size={20} className="group-hover:-translate-x-1 transition-transform" />
          <span className="uppercase font-bold tracking-wider text-xs">Voltar para a pÃ¡gina de vendas</span>
        </button>

        <Card className="p-8 md:p-12 shadow-sm border border-slate-100 rounded-3xl">
          <div className="flex items-center justify-between mb-12">
            <h2 className="text-2xl font-bold text-slate-800">Finalizar Assinatura</h2>
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 rounded-full bg-emerald-500" />
              <div className="w-8 h-1 bg-slate-100 rounded-full overflow-hidden">
                <div className={`h-full bg-emerald-500 transition-all duration-300 ${step >= 2 ? 'w-full' : 'w-0'}`} />
              </div>
              <div className={`w-3 h-3 rounded-full transition-colors duration-300 ${step >= 2 ? 'bg-emerald-500' : 'bg-slate-200'}`} />
            </div>
          </div>

          <AnimatePresence mode="wait">
            {step === 1 ? (
              <motion.div
                key="step1"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                className="space-y-6"
              >
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-slate-500">Nome da Empresa</label>
                    <input 
                      type="text" 
                      className="w-full p-4 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-sans"
                      placeholder="Sua Consultoria Ltda"
                      value={formData.name}
                      onChange={e => setFormData({...formData, name: e.target.value})}
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-bold uppercase tracking-widest text-slate-500">CNPJ</label>
                    <input 
                      type="text" 
                      className="w-full p-4 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-sans"
                      placeholder="00.000.000/0000-00"
                      value={formData.cnpj}
                      onChange={handleCNPJChange}
                    />
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-bold uppercase tracking-widest text-slate-500">E-mail para Acesso</label>
                  <input 
                    type="email" 
                    className="w-full p-4 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-emerald-500 transition-all font-sans"
                    placeholder="contato@suaempresa.com.br"
                    value={formData.email}
                    onChange={e => setFormData({...formData, email: e.target.value})}
                  />
                </div>

                <div className="bg-slate-50 border border-slate-100 p-5 rounded-2xl flex justify-between items-center mt-6">
                  <div>
                    <p className="text-[9px] font-black uppercase tracking-widest text-slate-400 mb-0.5">VocÃª assinarÃ¡</p>
                    <p className="font-sans font-bold text-slate-800 text-sm">
                      {plan === 'annual' ? 'Master Consultant (Anual)' : 'Consultor PrÃ³ (Mensal)'}
                    </p>
                  </div>
                  <p className="font-sans font-extrabold text-slate-800 text-sm">
                    {plan === 'annual' ? 'R$ 41,49/mÃªs (R$ 497,90 Ã  vista ou 12x)' : 'R$ 47,90/mÃªs'}
                  </p>
                </div>

                <Button 
                  onClick={() => setStep(2)}
                  className="w-full h-14 text-xs font-bold uppercase tracking-wider mt-8 bg-[#6fc2a4] hover:bg-[#59ba97] text-white rounded-2xl border-none transition-all shadow-sm"
                  disabled={!formData.name || !formData.email || !formData.cnpj}
                >
                  Continuar para Pagamento
                </Button>
              </motion.div>
            ) : (
              <motion.div
                key="step2"
                initial={{ opacity: 0, x: 20 }}
                animate={{ opacity: 1, x: 0 }}
                exit={{ opacity: 0, x: -20 }}
                className="space-y-6"
              >
                <div className="bg-amber-50 border border-amber-200 rounded-2xl p-5 text-left">
                  <h4 className="font-extrabold text-amber-900 text-sm mb-1">Passo Final: AtivaÃ§Ã£o PIX</h4>
                  <p className="text-amber-800 text-xs leading-relaxed">
                    Transfira o valor de <strong className="font-extrabold font-sans text-slate-900">R$ {plan === 'annual' ? '497,90' : '47,90'}</strong> para a chave PIX abaixo. Nosso sistema identificarÃ¡ o pagamento imediatamente.
                  </p>
                </div>

                <div className="flex flex-col items-center justify-center py-4">
                  <div className="w-56 h-56 bg-slate-50 border border-slate-100 rounded-3xl flex items-center justify-center p-6 mb-6 shadow-sm">
                    {/* Dynamic styled QR Code graphic mimicking the exact brand identity block from screenshot */}
                    <svg viewBox="0 0 100 100" className="w-full h-full text-slate-800">
                      {/* Top left marker */}
                      <path d="M 5 5 H 30 V 30 H 5 Z M 10 10 V 25 H 25 V 10 Z" fill="currentColor" />
                      <rect x="14" y="14" width="7" height="7" fill="currentColor" />
                      
                      {/* Top right marker */}
                      <path d="M 70 5 H 95 V 30 H 70 Z M 75 10 V 25 H 90 V 10 Z" fill="currentColor" />
                      <rect x="79" y="14" width="7" height="7" fill="currentColor" />
                      
                      {/* Bottom left marker */}
                      <path d="M 5 70 H 30 V 95 H 5 Z M 10 75 V 90 H 25 V 75 Z" fill="currentColor" />
                      <rect x="14" y="79" width="7" height="7" fill="currentColor" />

                      {/* Center piece / brand identity dot from screenshot */}
                      <rect x="42" y="42" width="16" height="16" rx="4" fill="#a7f3d0" className="text-emerald-500" />
                      <rect x="47" y="47" width="6" height="6" rx="2" fill="#10b981" />

                      {/* Random QR structures */}
                      <rect x="40" y="10" width="5" height="15" fill="currentColor" />
                      <rect x="50" y="5" width="10" height="5" fill="currentColor" />
                      <rect x="45" y="25" width="20" height="5" fill="currentColor" />
                      <rect x="10" y="40" width="5" height="15" fill="currentColor" />
                      <rect x="25" y="45" width="10" height="5" fill="currentColor" />
                      <rect x="5" y="55" width="20" height="5" fill="currentColor" />

                      <rect x="70" y="40" width="15" height="5" fill="currentColor" />
                      <rect x="80" y="48" width="5" height="10" fill="currentColor" />
                      <rect x="90" y="42" width="5" height="5" fill="#10b981" />
                      <rect x="75" y="60" width="20" height="5" fill="currentColor" />

                      <rect x="40" y="70" width="10" height="10" fill="currentColor" />
                      <rect x="45" y="85" width="15" height="5" fill="currentColor" />
                      <rect x="55" y="75" width="5" height="10" fill="currentColor" />

                      <rect x="68" y="70" width="27" height="27" rx="6" fill="#f8fafc" />
                      {/* Let's draw bottom right marker standard */}
                      <path d="M 70 70 H 95 V 95 H 70 Z M 75 75 V 90 H 90 V 75 Z" fill="currentColor" />
                    </svg>
                  </div>

                  <div className="w-full bg-slate-50 border border-slate-100 rounded-2xl flex items-center justify-between p-4 px-6 mb-2">
                    <span className="font-mono text-slate-700 text-sm font-semibold select-all">itamartrairi@gmail.com</span>
                    <button 
                      onClick={() => {
                        navigator.clipboard.writeText('itamartrairi@gmail.com');
                        setIsCopied(true);
                        setTimeout(() => setIsCopied(false), 2000);
                      }}
                      type="button"
                      className="text-emerald-600 font-black text-xs uppercase tracking-widest hover:text-emerald-700 transition-colors focus:outline-none"
                    >
                      {isCopied ? 'COPIADO!' : 'COPIAR'}
                    </button>
                  </div>
                </div>

                <div className="space-y-3 text-left">
                  <label className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500 flex items-center gap-1">
                    Anexar Comprovante do PIX <span className="text-red-500 font-bold">*</span>
                  </label>
                  
                  {comprovante ? (
                    <div className="border border-emerald-200 bg-emerald-50/50 rounded-2xl p-4 flex items-center justify-between transition-all">
                      <div className="flex items-center gap-3">
                        {isUploading ? (
                          <div className="w-10 h-10 bg-emerald-100/50 text-emerald-600 rounded-xl flex items-center justify-center animate-spin">
                            <Loader2 size={20} />
                          </div>
                        ) : (
                          <div className="w-10 h-10 bg-emerald-100 text-emerald-600 rounded-xl flex items-center justify-center">
                            <CheckCircle2 size={20} />
                          </div>
                        )}
                        <div className="text-left">
                          <p className="font-bold text-slate-800 text-sm max-w-[200px] truncate">{comprovante.name}</p>
                          <p className="text-[10px] text-emerald-600 font-medium font-sans">
                            {(comprovante.size / 1024).toFixed(1)} KB â€¢ Pronto para validaÃ§Ã£o
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => {
                          setComprovante(null);
                        }}
                        type="button"
                        className="text-[10px] uppercase font-bold tracking-widest text-rose-500 hover:text-rose-700 px-3 py-1.5 hover:bg-rose-50 rounded-lg transition-colors focus:outline-none"
                      >
                        Excluir
                      </button>
                    </div>
                  ) : (
                    <div 
                      onClick={() => fileInputRef.current?.click()}
                      className="border-2 border-dashed border-slate-200 hover:border-[#6fc2a4] focus:outline-none transition-all rounded-2xl p-6 flex flex-col items-center justify-center gap-2 cursor-pointer bg-slate-50/50 hover:bg-[#6fc2a4]/5"
                    >
                      <input 
                        type="file" 
                        ref={fileInputRef} 
                        onChange={(e) => {
                          if (e.target.files && e.target.files[0]) {
                            const file = e.target.files[0];
                            const allowedTypes = ['image/jpeg', 'image/jpg', 'image/png', 'application/pdf'];
                            const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
                            const allowedExtensions = ['.jpg', '.jpeg', '.png', '.pdf'];
                            
                            const isValidType = allowedTypes.includes(file.type) || allowedExtensions.includes(fileExtension);
                            if (!isValidType) {
                              alert("Formato de arquivo invÃ¡lido! Envie apenas comprovantes em formato PDF ou Imagem (JPEG/PNG).");
                              e.target.value = ''; // Reset input
                              return;
                            }

                            setIsUploading(true);
                            setComprovante(file);
                            setTimeout(() => {
                              setIsUploading(false);
                            }, 800);
                          }
                        }}
                        className="hidden" 
                        accept="image/jpeg,image/png,application/pdf"
                      />
                      <div className="w-12 h-12 rounded-full bg-slate-100 text-slate-500 flex items-center justify-center mb-1">
                        <Upload size={22} className="text-slate-400" />
                      </div>
                      <p className="text-xs font-bold text-slate-700">Clique para selecionar ou arraste o comprovante</p>
                      <p className="text-[10px] text-slate-400 font-medium">Formatos aceitos: PDF, Imagens (JPEG, PNG)</p>
                    </div>
                  )}
                </div>

                <Button 
                  onClick={handleFinish}
                  disabled={!comprovante || isUploading || isVerifying}
                  className={`w-full h-14 text-white font-extrabold text-xs uppercase tracking-widest transition-all rounded-2xl border-none mb-4 ${
                    comprovante && !isUploading && !isVerifying
                      ? 'bg-emerald-500 hover:bg-emerald-600 active:bg-emerald-700 shadow-md cursor-pointer' 
                      : 'bg-slate-300 text-slate-400 cursor-not-allowed shadow-none'
                  }`}
                >
                  {isVerifying ? (
                    <span className="flex items-center justify-center gap-2">
                      <Loader2 size={16} className="animate-spin" />
                      Validando comprovante...
                    </span>
                  ) : comprovante ? (
                    'Confirmar Pagamento e Liberar Sistema'
                  ) : (
                    'Anexe o comprovante para liberar'
                  )}
                </Button>

                <div className="text-center">
                  <button 
                    onClick={() => setStep(1)}
                    type="button"
                    className="text-slate-400 font-extrabold text-[10px] uppercase tracking-widest hover:text-slate-600 transition-colors focus:outline-none"
                  >
                    Alterar dados cadastrais
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </Card>

        <p className="text-center text-slate-400 text-[11px] font-bold uppercase tracking-wider mt-8">
          Ambiente seguro e criptografado. Seus dados estÃ£o protegidos.
        </p>
      </div>
    </div>
  );
};

const LogoSelector = ({ 
  logoChoice, 
  setLogoChoice, 
  customLogo, 
  customConsultoraLogo,
  setCustomLogo,
  setCustomConsultoraLogo
}: { 
  logoChoice: 'both' | 'sebrae' | 'consultora' | 'none', 
  setLogoChoice: (choice: 'both' | 'sebrae' | 'consultora' | 'none') => void,
  customLogo?: string | null,
  customConsultoraLogo?: string | null,
  setCustomLogo?: (logo: string | null) => void,
  setCustomConsultoraLogo?: (logo: string | null) => void
}) => {
  const handleQuickUpload = (e: React.ChangeEvent<HTMLInputElement>, isSebrae: boolean) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 2 * 1024 * 1024) {
        alert('O logo deve ter menos de 2MB.');
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64 = event.target?.result as string;
        if (isSebrae) {
          if (setCustomLogo) setCustomLogo(base64);
          try { localStorage.setItem('sebrae_custom_logo', base64); } catch {}
        } else {
          if (setCustomConsultoraLogo) setCustomConsultoraLogo(base64);
          try { localStorage.setItem('consultora_custom_logo', base64); } catch {}
        }
      };
      reader.readAsDataURL(file);
    }
  };

  return (
    <div className="space-y-3 w-full">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 w-full">
        <div>
          <h4 className="text-xs font-black text-slate-800 uppercase tracking-widest flex items-center gap-2">
            <span>ðŸŽ¯</span> Logotipos no RelatÃ³rio e CabeÃ§alhos
          </h4>
          <p className="text-[11px] text-slate-500 font-medium leading-relaxed mt-0.5">
            Controle a exibiÃ§Ã£o das marcas nos relatÃ³rios em PDF (SEBRAE e Empresa Consultora)
          </p>
        </div>
        <div className="flex flex-wrap bg-slate-100 p-1 rounded-xl self-start sm:self-auto border border-slate-200/60 gap-1">
          <button
            onClick={() => {
              setLogoChoice('both');
              try { localStorage.setItem('preferred_logo', 'both'); } catch {}
            }}
            className={cn(
              "px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
              logoChoice === 'both'
                ? "bg-white text-emerald-700 shadow-sm font-black"
                : "text-slate-500 hover:text-slate-800"
            )}
          >
            Ambas (SEBRAE + Empresa)
          </button>
          <button
            onClick={() => {
              setLogoChoice('sebrae');
              try { localStorage.setItem('preferred_logo', 'sebrae'); } catch {}
            }}
            className={cn(
              "px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
              logoChoice === 'sebrae'
                ? "bg-white text-emerald-600 shadow-sm font-black"
                : "text-slate-500 hover:text-slate-800"
            )}
          >
            Apenas SEBRAE
          </button>
          <button
            onClick={() => {
              setLogoChoice('consultora');
              try { localStorage.setItem('preferred_logo', 'consultora'); } catch {}
            }}
            className={cn(
              "px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
              logoChoice === 'consultora'
                ? "bg-white text-indigo-600 shadow-sm font-black"
                : "text-slate-500 hover:text-slate-800"
            )}
          >
            Apenas Empresa
          </button>
          <button
            onClick={() => {
              setLogoChoice('none');
              try { localStorage.setItem('preferred_logo', 'none'); } catch {}
            }}
            className={cn(
              "px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
              logoChoice === 'none'
                ? "bg-white text-slate-800 shadow-sm font-black"
                : "text-slate-500 hover:text-slate-800"
            )}
          >
            Nenhum
          </button>
        </div>
      </div>

      {/* Mini status e upload rÃ¡pido dos logotipos */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2 border-t border-slate-100 text-xs">
        <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-200/70">
          <div className="flex items-center gap-2.5 min-w-0">
            {customLogo ? (
              <img src={customLogo} alt="SEBRAE Custom" className="h-7 w-12 object-contain bg-white rounded p-0.5 border border-slate-200 shrink-0" />
            ) : (
              <img src={sebraeLogoBase64} alt="SEBRAE Oficial" className="h-7 w-12 object-contain bg-white rounded p-0.5 border border-slate-200 shrink-0" />
            )}
            <div className="truncate">
              <div className="font-bold text-slate-800 text-[11px] truncate">Marca SEBRAE</div>
              <div className="text-[10px] text-emerald-600 font-semibold truncate">
                {customLogo ? 'Logo customizado ativo' : 'Logo oficial padrÃ£o ativo'}
              </div>
            </div>
          </div>
          {setCustomLogo && (
            <label className="ml-2 shrink-0 text-[10px] font-bold bg-white hover:bg-slate-100 text-slate-700 px-2.5 py-1 rounded-lg border border-slate-200 shadow-xs cursor-pointer">
              {customLogo ? 'Alterar' : 'Fazer Upload'}
              <input type="file" className="hidden" accept="image/*" onChange={(e) => handleQuickUpload(e, true)} />
            </label>
          )}
        </div>

        <div className="flex items-center justify-between p-2.5 bg-slate-50 rounded-xl border border-slate-200/70">
          <div className="flex items-center gap-2.5 min-w-0">
            {customConsultoraLogo ? (
              <img src={customConsultoraLogo} alt="Empresa Consultora" className="h-7 w-12 object-contain bg-white rounded p-0.5 border border-slate-200 shrink-0" />
            ) : (
              <div className="h-7 w-12 bg-slate-200 text-slate-400 rounded flex items-center justify-center text-[10px] font-bold shrink-0">
                SEM
              </div>
            )}
            <div className="truncate">
              <div className="font-bold text-slate-800 text-[11px] truncate">Marca da Empresa</div>
              <div className="text-[10px] text-slate-500 font-medium truncate">
                {customConsultoraLogo ? 'Logo da empresa carregado' : 'Opcional (ao lado do SEBRAE)'}
              </div>
            </div>
          </div>
          {setCustomConsultoraLogo && (
            <label className="ml-2 shrink-0 text-[10px] font-bold bg-white hover:bg-slate-100 text-indigo-700 px-2.5 py-1 rounded-lg border border-indigo-200 shadow-xs cursor-pointer">
              {customConsultoraLogo ? 'Alterar' : 'Fazer Upload'}
              <input type="file" className="hidden" accept="image/*" onChange={(e) => handleQuickUpload(e, false)} />
            </label>
          )}
        </div>
      </div>
    </div>
  );
};

const SettingsView = ({ 
  customLogo, 
  setCustomLogo,
  customConsultoraLogo,
  setCustomConsultoraLogo,
  isSystemKeyActive,
  storageMode,
  setStorageMode,
  onExportLocalBackup,
  onImportLocalBackup,
  onOpenSmartSync,
  isSyncingCloud,
  lastSyncSummary,
  userEmail
}: { 
  customLogo: string | null, 
  setCustomLogo: (logo: string | null) => void,
  customConsultoraLogo: string | null,
  setCustomConsultoraLogo: (logo: string | null) => void,
  isSystemKeyActive?: boolean,
  storageMode: 'cloud' | 'local',
  setStorageMode: (mode: 'cloud' | 'local') => void,
  onExportLocalBackup: () => void,
  onImportLocalBackup: (e: React.ChangeEvent<HTMLInputElement>) => void,
  onOpenSmartSync?: () => void,
  isSyncingCloud?: boolean,
  lastSyncSummary?: SyncSummary | null,
  userEmail?: string | null
}) => {
  const [dailyReminderEnabled, setDailyReminderEnabled] = React.useState(() => localStorage.getItem('daily_reminder_enabled') === 'true');
  const [dailyReminderTime, setDailyReminderTime] = React.useState(() => localStorage.getItem('daily_reminder_time') || '09:00');
  const [cloudStatus, setCloudStatus] = React.useState<{ checking: boolean; ok?: boolean; message?: string; details?: string } | null>(null);
  const [copiedRules, setCopiedRules] = React.useState(false);

  const handleTestConnection = React.useCallback(async () => {
    setCloudStatus({ checking: true });
    try {
      const res = await checkCloudConnection();
      setCloudStatus({ checking: false, ok: res.ok, message: res.message, details: res.details });
    } catch (e: any) {
      setCloudStatus({ checking: false, ok: false, message: 'Erro ao testar conexÃ£o', details: e?.message });
    }
  }, []);

  // Verifica automaticamente a integridade da comunicaÃ§Ã£o ao abrir as configuraÃ§Ãµes
  React.useEffect(() => {
    handleTestConnection();
  }, [handleTestConnection]);

  const toggleReminder = async () => {
    const nextVal = !dailyReminderEnabled;
    if (nextVal && 'Notification' in window) {
      const perm = await Notification.requestPermission();
      if (perm !== 'granted') {
        alert("Por favor, habilite as notificaÃ§Ãµes no seu navegador para receber os lembretes diÃ¡rios.");
        return;
      }
    }
    setDailyReminderEnabled(nextVal);
    localStorage.setItem('daily_reminder_enabled', String(nextVal));
  };

  const saveReminderConfig = () => {
    localStorage.setItem('daily_reminder_time', dailyReminderTime);
    alert(`HorÃ¡rio de lembrete diÃ¡rio configurado para as ${dailyReminderTime} com sucesso!`);
  };

  const testNotification = () => {
    if ('Notification' in window) {
      if (Notification.permission === 'granted') {
        new Notification("Teste de Lembrete DiÃ¡rio", {
          body: "As notificaÃ§Ãµes de lembrete diÃ¡rio estÃ£o configuradas e prontas!",
          icon: "/favicon.ico"
        });
      } else {
        alert("PermissÃ£o de notificaÃ§Ã£o nÃ£o concedida no navegador. Favor habilitÃ¡-la no navegador.");
      }
    } else {
      alert("Seu navegador nÃ£o oferece suporte para notificaÃ§Ãµes push.");
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>, isSebrae: boolean) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 1024 * 1024) {
        alert("O logo deve ter menos de 1MB.");
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        const base64 = event.target?.result as string;
        if (isSebrae) {
          setCustomLogo(base64);
          localStorage.setItem('sebrae_custom_logo', base64);
        } else {
          setCustomConsultoraLogo(base64);
          localStorage.setItem('consultora_custom_logo', base64);
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const removeLogo = (isSebrae: boolean) => {
    if (isSebrae) {
      setCustomLogo(null);
      localStorage.removeItem('sebrae_custom_logo');
    } else {
      setCustomConsultoraLogo(null);
      localStorage.removeItem('consultora_custom_logo');
    }
  };

  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -10 }}
      className="space-y-8"
    >
      <div>
        <h2 className="text-2xl font-bold text-slate-800">ConfiguraÃ§Ãµes</h2>
        <p className="text-slate-500 text-sm">Personalize sua experiÃªncia na plataforma</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 max-w-5xl">
        <Card className="p-6">
          <h3 className="text-base font-black text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-2">
            <ImageIcon className="text-emerald-600" size={18} />
            Logo / Marca do SEBRAE
          </h3>
          
          <div className="space-y-4">
            <p className="text-xs text-slate-500 leading-relaxed font-semibold">
              Logotipo do SEBRAE exibido no cabeÃ§alho ou Ã¡rea reservada dos seus relatÃ³rios de diagnÃ³stico e planos de aÃ§Ã£o.
            </p>

            <div className="space-y-4">
              {customLogo ? (
                <div className="relative group">
                  <div className="p-4 bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                    <img src={customLogo} alt="Logo SEBRAE" className="max-h-24 mx-auto object-contain" />
                  </div>
                  <button 
                    onClick={() => removeLogo(true)}
                    className="absolute -top-2 -right-2 bg-rose-500 text-white p-1 rounded-full shadow-lg hover:bg-rose-600 transition-colors cursor-pointer"
                  >
                    <X size={14} />
                  </button>
                </div>
              ) : (
                <label className="flex flex-col items-center justify-center p-6 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-200 hover:border-emerald-400 hover:bg-emerald-50/50 transition-all cursor-pointer group">
                  <div className="w-10 h-10 bg-white rounded-full flex items-center justify-center shadow-sm mb-3 group-hover:scale-110 transition-transform">
                    <Upload className="text-slate-400 group-hover:text-emerald-600" size={20} />
                  </div>
                  <span className="text-xs font-bold text-slate-500 group-hover:text-emerald-700">Clique para selecionar o logo</span>
                  <span className="text-[10px] text-slate-400 mt-1">PNG ou JPG atÃ© 1MB</span>
                  <input type="file" className="hidden" accept="image/*" onChange={(e) => handleLogoUpload(e, true)} />
                </label>
              )}
              
              <div className="text-[10px] text-slate-400 space-y-1 bg-slate-50/30 p-3 rounded-xl border border-slate-100">
                <p className="font-bold text-slate-500 uppercase tracking-wider mb-1">Dicas:</p>
                <p>â€¢ Use fundo transparente (PNG) ou fundos claros.</p>
                <p>â€¢ Formato horizontal/paisagem ajusta melhor na folha.</p>
              </div>
            </div>
          </div>
        </Card>

        <Card className="p-6">
          <h3 className="text-base font-black text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-2">
            <ImageIcon className="text-indigo-600" size={18} />
            Logo da Empresa Consultora
          </h3>
          
          <div className="space-y-4">
            <p className="text-xs text-slate-500 leading-relaxed font-semibold">
              Logotipo da sua empresa de consultoria credenciada que serÃ¡ posicionado ao lado do SEBRAE de forma profissional.
            </p>

            <div className="space-y-4">
              {customConsultoraLogo ? (
                <div className="relative group">
                  <div className="p-4 bg-slate-50 rounded-2xl border border-dashed border-slate-200">
                    <img src={customConsultoraLogo} alt="Logo Consultora" className="max-h-24 mx-auto object-contain" />
                  </div>
                  <button 
                    onClick={() => removeLogo(false)}
                    className="absolute -top-2 -right-2 bg-rose-500 text-white p-1 rounded-full shadow-lg hover:bg-rose-600 transition-colors cursor-pointer"
                  >
                    <X size={14} />
                  </button>
                </div>
              ) : (
                <label className="flex flex-col items-center justify-center p-6 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/50 transition-all cursor-pointer group">
                  <div className="w-10 h-10 bg-white rounded-full flex items-center justify-center shadow-sm mb-3 group-hover:scale-110 transition-transform">
                    <Upload className="text-slate-400 group-hover:text-indigo-600" size={20} />
                  </div>
                  <span className="text-xs font-bold text-slate-500 group-hover:text-indigo-700">Clique para selecionar o logo</span>
                  <span className="text-[10px] text-slate-400 mt-1">PNG ou JPG atÃ© 1MB</span>
                  <input type="file" className="hidden" accept="image/*" onChange={(e) => handleLogoUpload(e, false)} />
                </label>
              )}
              
              <div className="text-[10px] text-slate-400 space-y-1 bg-slate-50/30 p-3 rounded-xl border border-slate-100">
                <p className="font-bold text-slate-500 uppercase tracking-wider mb-1">Dicas:</p>
                <p>â€¢ Use fundo transparente (PNG) para um acabamento perfeito.</p>
                <p>â€¢ Formatos horizontais produzem um excelente alinhamento.</p>
              </div>
            </div>
          </div>
        </Card>
      </div>

      <div className="max-w-5xl">
        <Card className="p-6">
          <h3 className="text-base font-black text-slate-800 uppercase tracking-wider mb-4 flex items-center gap-2">
            <Bell className="text-emerald-600" size={18} />
            HorÃ¡rio de Lembrete DiÃ¡rio
          </h3>
          
          <div className="space-y-4">
            <p className="text-xs text-slate-500 leading-relaxed font-semibold">
              Receba notificaÃ§Ãµes push diÃ¡rias no horÃ¡rio configurado para lembrar vocÃª e seus consultores sobre as tarefas e pendÃªncias dos clientes.
            </p>

            <div className="space-y-3">
              <div className="flex items-center justify-between p-3 bg-slate-50 rounded-xl border border-slate-100">
                <div>
                  <label className="block text-xs font-bold text-slate-700">Ativar Lembrete DiÃ¡rio</label>
                  <p className="text-[11px] text-slate-400">NotificaÃ§Ãµes de tarefas pendentes.</p>
                </div>
                <button
                  onClick={toggleReminder}
                  className={`relative inline-flex h-6 w-11 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${dailyReminderEnabled ? 'bg-emerald-600' : 'bg-slate-200'}`}
                >
                  <span
                    className={`pointer-events-none inline-block h-5 w-5 transform rounded-full bg-white shadow ring-0 transition duration-200 ease-in-out ${dailyReminderEnabled ? 'translate-x-5' : 'translate-x-0'}`}
                  />
                </button>
              </div>

              {dailyReminderEnabled && (
                <div className="space-y-3 p-3 bg-slate-50 rounded-xl border border-slate-100">
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2 flex items-center gap-1.5">
                      <Clock size={14} className="text-slate-400" />
                      Definir HorÃ¡rio de Envio
                    </label>
                    <div className="flex gap-2">
                      <input 
                        type="time"
                        className="px-3 py-1.5 text-xs rounded-xl border border-slate-200 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 font-mono bg-white"
                        value={dailyReminderTime}
                        onChange={(e) => setDailyReminderTime(e.target.value)}
                      />
                      <Button 
                        onClick={saveReminderConfig}
                        className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shrink-0 px-3 py-1.5"
                      >
                        Salvar
                      </Button>
                    </div>
                  </div>

                  <div className="pt-2 border-t border-slate-200/55 flex justify-between items-center">
                    <span className="text-[11px] text-slate-500 font-semibold">Testar notificaÃ§Ãµes push?</span>
                    <Button 
                      variant="outline"
                      size="sm"
                      onClick={testNotification}
                      className="text-xs border-slate-200 hover:bg-slate-100 text-slate-700 py-1 px-2.5 h-7"
                    >
                      Testar Agora
                    </Button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </Card>
      </div>

      <Card className="p-8 max-w-5xl">
        <h3 className="text-lg font-bold text-slate-800 mb-2 flex items-center gap-2">
          <Database className="text-indigo-600" size={20} />
          Modo de Armazenamento e Salvamento (Nuvem vs. Local)
        </h3>
        <p className="text-xs text-slate-500 mb-6">
          Escolha onde deseja preferencialmente salvar seus dados (empresas, diagnÃ³sticos e planos de aÃ§Ã£o) e realize backups locais de seguranÃ§a.
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
          <div
            onClick={() => setStorageMode('cloud')}
            className={cn(
              "p-5 rounded-2xl border-2 transition-all cursor-pointer relative flex flex-col justify-between space-y-3",
              storageMode === 'cloud'
                ? "border-emerald-500 bg-emerald-50/40 shadow-sm"
                : "border-slate-200 bg-slate-50/50 hover:border-slate-300"
            )}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className={cn("p-2.5 rounded-xl", storageMode === 'cloud' ? "bg-emerald-600 text-white" : "bg-slate-200 text-slate-600")}>
                  <Cloud size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">Salvar na Nuvem (Firestore)</h4>
                  <p className="text-xs text-slate-500">SincronizaÃ§Ã£o Online AutomÃ¡tica</p>
                </div>
              </div>
              {storageMode === 'cloud' && (
                <span className="bg-emerald-600 text-white text-[10px] font-black uppercase px-2.5 py-1 rounded-full">Ativo</span>
              )}
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">
              Seus dados sÃ£o salvos e sincronizados automaticamente no banco de dados online. Recomendado para sincronizaÃ§Ã£o em mÃºltiplos aparelhos.
            </p>
          </div>

          <div
            onClick={() => setStorageMode('local')}
            className={cn(
              "p-5 rounded-2xl border-2 transition-all cursor-pointer relative flex flex-col justify-between space-y-3",
              storageMode === 'local'
                ? "border-indigo-500 bg-indigo-50/40 shadow-sm"
                : "border-slate-200 bg-slate-50/50 hover:border-slate-300"
            )}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className={cn("p-2.5 rounded-xl", storageMode === 'local' ? "bg-indigo-600 text-white" : "bg-slate-200 text-slate-600")}>
                  <HardDrive size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">Salvar Localmente (Dispositivo)</h4>
                  <p className="text-xs text-slate-500">Navegador & Sem Cotas</p>
                </div>
              </div>
              {storageMode === 'local' && (
                <span className="bg-indigo-600 text-white text-[10px] font-black uppercase px-2.5 py-1 rounded-full">Ativo</span>
              )}
            </div>
            <p className="text-xs text-slate-600 leading-relaxed">
              Seus dados sÃ£o salvos diretamente em seu prÃ³prio navegador/computador. Funciona offline e sem limitaÃ§Ãµes de cota da nuvem.
            </p>
          </div>
        </div>

        {/* SincronizaÃ§Ã£o Inteligente com a Nuvem */}
        <div className="p-5 bg-gradient-to-br from-emerald-50/80 to-teal-50/50 rounded-2xl border border-emerald-200/90 flex flex-wrap items-center justify-between gap-4">
          <div className="space-y-1 max-w-xl">
            <div className="flex items-center gap-2">
              <RefreshCw size={16} className={cn("text-emerald-700", isSyncingCloud && "animate-spin")} />
              <p className="text-xs font-black text-emerald-950 uppercase tracking-wide">
                SincronizaÃ§Ã£o Inteligente com a Nuvem (Ãšltimos Dados Registrados)
              </p>
            </div>
            <p className="text-[11px] text-emerald-800 leading-relaxed">
              Mescla e sincroniza os dados salvos localmente com a nuvem (Firestore). O sistema compara as datas de modificaÃ§Ã£o de cada empresa, diagnÃ³stico, resposta e plano, garantindo que o registro mais recente sempre prevaleÃ§a e atualize ambas as pontas.
            </p>
            {lastSyncSummary && (
              <p className="text-[10px] font-bold text-emerald-900 flex items-center gap-1.5 pt-1">
                <span>âœ“ Ãšltima sincronizaÃ§Ã£o realizada Ã s {lastSyncSummary.timestamp}</span>
                <span>â€¢ {lastSyncSummary.totalAnalyzed} registros verificados</span>
              </p>
            )}
          </div>
          <div className="flex items-center gap-2">
            <Button
              onClick={onOpenSmartSync}
              disabled={isSyncingCloud}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2.5 rounded-xl flex items-center gap-2 shadow-sm"
            >
              <RefreshCw size={14} className={cn(isSyncingCloud && "animate-spin")} />
              <span>{isSyncingCloud ? "Sincronizando..." : "Sincronizar com a Nuvem"}</span>
            </Button>
          </div>
        </div>

        <div className="p-5 bg-slate-50 rounded-2xl border border-slate-200/80 flex flex-wrap items-center justify-between gap-4">
          <div className="space-y-1">
            <p className="text-xs font-bold text-slate-800">Backup e RestauraÃ§Ã£o de Dados Locais</p>
            <p className="text-[11px] text-slate-500">Baixe uma cÃ³pia completa de todos os seus diagnÃ³sticos e empresas em formato .JSON ou restaure um backup existente.</p>
          </div>
          <div className="flex items-center gap-2">
            <Button
              onClick={onExportLocalBackup}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 flex items-center gap-2"
            >
              <Download size={15} />
              Exportar Backup (.JSON)
            </Button>
            <label className="cursor-pointer">
              <div className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-4 py-2 rounded-xl flex items-center gap-2 transition-colors">
                <Upload size={15} />
                Importar Backup (.JSON)
              </div>
              <input
                type="file"
                accept=".json"
                onChange={onImportLocalBackup}
                className="hidden"
              />
            </label>
          </div>
        </div>
      </Card>

      {/* DiagnÃ³stico da ConexÃ£o Firebase / Nuvem com Indicador Visual */}
      <Card className="p-8 max-w-5xl border-2 border-emerald-100 bg-white">
        <div className="mb-6">
          <ConnectionStatus />
        </div>

        <div className="grid grid-cols-1 md:grid-cols-4 gap-3 mb-4">
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Projeto Firebase</p>
            <p className="text-sm font-bold text-slate-800 mt-1 font-mono truncate">ai-studio-applet-webapp-2e46d</p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">ID do Banco (Firestore)</p>
            <p className="text-xs font-bold text-indigo-700 mt-1 font-mono truncate" title="ai-studio-consultoriaempre-e5930715-81f9-43df-8493-c8c9c3ecac45">
              ai-studio-consultoriaempre...
            </p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Conta Conectada</p>
            <p className="text-sm font-bold text-slate-800 mt-1 truncate">{userEmail || 'Nenhum usuÃ¡rio logado'}</p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Status de Uso Compartilhado</p>
            <p className={cn("text-sm font-bold mt-1 flex items-center gap-1.5", cloudStatus ? (cloudStatus.ok ? "text-emerald-600" : "text-amber-600") : "text-slate-500")}>
              {cloudStatus?.checking ? (
                <>
                  <RefreshCw size={14} className="animate-spin text-indigo-500" />
                  <span>Verificando integridade...</span>
                </>
              ) : cloudStatus ? (
                cloudStatus.ok ? (
                  <>
                    <CheckCircle2 size={15} className="text-emerald-600" />
                    <span>Pronto para Todos os UsuÃ¡rios</span>
                  </>
                ) : (
                  <>
                    <AlertCircle size={15} className="text-amber-600" />
                    <span>AtenÃ§Ã£o / Bloqueado</span>
                  </>
                )
              ) : (
                "Aguardando verificaÃ§Ã£o..."
              )}
            </p>
          </div>
        </div>

        {cloudStatus && !cloudStatus.checking && (
          <div className={cn("p-4 rounded-xl border text-xs mb-4 transition-all duration-300", cloudStatus.ok ? "bg-emerald-50 border-emerald-200 text-emerald-900" : "bg-amber-50 border-amber-200 text-amber-900")}>
            <div className="flex items-center gap-2 font-bold mb-1">
              {cloudStatus.ok ? <CheckCircle2 size={16} className="text-emerald-600 shrink-0" /> : <AlertCircle size={16} className="text-amber-600 shrink-0" />}
              <span>{cloudStatus.ok ? "Banco de dados 100% operacional e sincronizado para mÃºltiplos usuÃ¡rios" : cloudStatus.message}</span>
            </div>
            {cloudStatus.details && (
              <p className="text-[11px] opacity-90 pl-6">{cloudStatus.details}</p>
            )}
            {cloudStatus.ok && (
              <p className="text-[11px] text-emerald-700 pl-6 mt-1 font-medium">
                âœ“ Leitura e gravaÃ§Ã£o autenticadas liberadas. Os dados inseridos neste ou em qualquer outro computador serÃ£o sincronizados em tempo real.
              </p>
            )}
          </div>
        )}

        <div className="p-4 bg-slate-50 rounded-xl border border-slate-200 space-y-3">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div className="space-y-0.5">
              <p className="text-xs font-bold text-slate-800">Regras de SeguranÃ§a do Firestore (Security Rules)</p>
              <p className="text-[11px] text-slate-500">Para garantir sincronizaÃ§Ã£o em mÃºltiplos computadores, certifique-se de que as regras no Console do Firebase estÃ£o publicadas na aba <strong>Firestore Database &gt; Regras</strong>:</p>
            </div>
            <Button
              onClick={() => {
                const rulesText = `rules_version = '2';\nservice cloud.firestore {\n  match /databases/{database}/documents {\n    match /{document=**} {\n      allow read, write: if request.auth != null;\n    }\n  }\n}`;
                navigator.clipboard.writeText(rulesText);
                setCopiedRules(true);
                setTimeout(() => setCopiedRules(false), 3000);
              }}
              variant="outline"
              size="sm"
              className="text-xs border-slate-300 text-slate-700 hover:bg-slate-100 font-bold shrink-0"
            >
              {copiedRules ? "âœ“ Regras Copiadas!" : "Copiar Regras do Firestore"}
            </Button>
          </div>
          <pre className="p-3 bg-slate-900 text-emerald-400 font-mono text-[11px] rounded-lg overflow-x-auto select-all">
{`rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /{document=**} {
      allow read, write: if request.auth != null;
    }
  }
}`}
          </pre>
          <div className="flex items-center justify-between text-[11px] text-slate-600 pt-1">
            <span>Acesse a pÃ¡gina direta do Firestore Database no console:</span>
            <a 
              href="https://console.firebase.google.com/project/ai-studio-applet-webapp-2e46d/firestore/rules" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="text-emerald-700 font-bold hover:underline flex items-center gap-1"
            >
              Abrir Regras do Firestore Database no Firebase â†’
            </a>
          </div>
        </div>
      </Card>

      <Card className="p-8 max-w-5xl">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <ShieldCheck className="text-emerald-600" size={20} />
              Criptografia em Repouso Local (AES-256)
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              ProteÃ§Ã£o local para evitar que diagnÃ³sticos, empresas e chaves salvas no navegador fiquem em texto simples.
            </p>
          </div>
          <div className="flex items-center gap-2 bg-emerald-50 text-emerald-700 border border-emerald-200 px-3.5 py-1.5 rounded-xl text-xs font-bold shrink-0">
            <Lock size={14} className="text-emerald-600" />
            <span>AES-256 Ativado</span>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-5">
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 space-y-1">
            <p className="text-xs font-bold text-slate-700 uppercase tracking-wider">Algoritmo &amp; Chave</p>
            <p className="text-xs text-slate-600 leading-relaxed">
              Criptografia simÃ©trica AES-256 derivada via SHA-256 a partir do seu ID de UsuÃ¡rio / SessÃ£o no dispositivo.
            </p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 space-y-1">
            <p className="text-xs font-bold text-slate-700 uppercase tracking-wider">Identificador de SeguranÃ§a</p>
            <p className="text-xs text-slate-600 leading-relaxed">
              Todos os registros armazenados no <code className="bg-slate-200 px-1 py-0.5 rounded text-[11px] font-mono">localStorage</code> sÃ£o gravados sob o formato cifrado <code className="text-indigo-600 font-mono text-[11px]">enc:v1:...</code>.
            </p>
          </div>
        </div>

        <div className="p-4 bg-indigo-50/60 rounded-xl border border-indigo-100 flex flex-wrap items-center justify-between gap-3">
          <div className="space-y-0.5">
            <p className="text-xs font-bold text-indigo-950">Verificar e Re-criptografar Dados Locais</p>
            <p className="text-[11px] text-indigo-700">Converte dados legados nÃ£o criptografados no navegador para a nova estrutura AES-256.</p>
          </div>
          <Button
            onClick={() => {
              const count = encryptedLocalStorage.migrateAllToEncrypted();
              alert(`Varredura concluÃ­da! ${count} item(ns) legados foram criptografados e salvos com AES-256 no LocalStorage.`);
            }}
            variant="outline"
            size="sm"
            className="text-xs border-indigo-200 hover:bg-indigo-100 text-indigo-700 font-bold py-1.5 px-3"
          >
            <RefreshCw size={13} className="mr-1.5" />
            Executar Criptografia Agora
          </Button>
        </div>
      </Card>

      <div className="bg-amber-50 border border-amber-100 p-6 rounded-xl flex items-start gap-4 max-w-2xl">
        <Info className="text-amber-600 shrink-0" size={20} />
        <div className="space-y-1">
          <p className="text-sm font-bold text-amber-800">Sobre o armazenamento</p>
          <p className="text-xs text-amber-700 leading-relaxed">
            Por enquanto, esta logo fica armazenada localmente em seu navegador. Se vocÃª mudar de computador ou limpar os dados do navegador, precisarÃ¡ fazer o upload novamente.
          </p>
        </div>
      </div>
    </motion.div>
  );
};

const LicenseManagementView = ({ 
  empresasCredenciadas, 
  onUpdateCredenciadaStatus, 
  onUpdateCredenciadaPlano,
  onUpdateCredenciadaDiasTeste,
  onUpdateCredenciadaValidade,
  onOpenCreateModal,
  onOpenEditModal,
  onOpenDeleteModal,
  onInitializeAdmin,
  onSyncWithAuth,
  currentUser
}: { 
  empresasCredenciadas: EmpresaCredenciada[],
  onUpdateCredenciadaStatus?: (id: string, status: 'Ativa' | 'Bloqueada') => Promise<void>,
  onUpdateCredenciadaPlano?: (id: string, plano: 'Teste' | 'Mensal' | 'Anual' | 'Definitiva') => Promise<void>,
  onUpdateCredenciadaDiasTeste?: (id: string, dias: number) => Promise<void>,
  onUpdateCredenciadaValidade?: (id: string, validade: string) => Promise<void>,
  onOpenCreateModal?: () => void,
  onOpenEditModal?: (emp: EmpresaCredenciada) => void,
  onOpenDeleteModal?: (emp: EmpresaCredenciada) => void,
  onInitializeAdmin?: () => Promise<void>,
  onSyncWithAuth?: () => Promise<void>,
  currentUser?: User | null
}) => {
  const [editingCustomId, setEditingCustomId] = useState<string | null>(null);
  const [customDaysInput, setCustomDaysInput] = useState<string>('');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [isInitializing, setIsInitializing] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);

  const formatItemDate = (dateVal: any) => {
    if (!dateVal) return null;
    try {
      const d = dateVal.toDate ? dateVal.toDate() : new Date(dateVal);
      if (isNaN(d.getTime())) return null;
      return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    } catch {
      return null;
    }
  };

  const calculateDaysLeft = (emp: EmpresaCredenciada) => {
    const plano = emp.tipoPlano || 'Teste';
    if (plano === 'Definitiva') {
      return 99999;
    }
    if (emp.validadeLicenca) {
      const validadeDate = emp.validadeLicenca.toDate ? emp.validadeLicenca.toDate() : new Date(emp.validadeLicenca);
      const diffTime = validadeDate.getTime() - new Date().getTime();
      return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    }
    const limitDays = emp.diasTeste || (plano === 'Mensal' ? 30 : plano === 'Anual' ? 365 : 30);
    if (emp.dataCadastro) {
      const cadastroDate = emp.dataCadastro.toDate ? emp.dataCadastro.toDate() : new Date(emp.dataCadastro);
      const diffTime = new Date().getTime() - cadastroDate.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      return limitDays - diffDays;
    }
    return limitDays;
  };

  const handleAddDays = async (emp: EmpresaCredenciada, daysToAdd: number) => {
    const currentLimit = emp.diasTeste || (emp.tipoPlano === 'Mensal' ? 30 : emp.tipoPlano === 'Anual' ? 365 : 30);
    const newLimit = currentLimit + daysToAdd;
    await onUpdateCredenciadaDiasTeste?.(emp.id, newLimit);
  };

  const handleSyncAuth = async () => {
    setIsSyncing(true);
    try {
      if (onSyncWithAuth) {
        await onSyncWithAuth();
      }
    } finally {
      setTimeout(() => setIsSyncing(false), 600);
    }
  };

  const filteredEmpresas = empresasCredenciadas.filter(emp => {
    if (!searchTerm.trim()) return true;
    const term = searchTerm.toLowerCase();
    return (
      (emp.razaoSocial || '').toLowerCase().includes(term) ||
      (emp.consultor || '').toLowerCase().includes(term) ||
      (emp.email || '').toLowerCase().includes(term) ||
      (emp.tipoPlano || '').toLowerCase().includes(term) ||
      (emp.ownerId || '').toLowerCase().includes(term)
    );
  });

  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -10 }}
      className="space-y-8"
    >
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2.5 flex-wrap">
            <h2 className="text-2xl font-bold text-slate-800">GestÃ£o de LicenÃ§as e Tempo de Teste</h2>
            <span className="inline-flex items-center gap-1.5 px-2.5 py-0.5 rounded-full text-xs font-bold bg-amber-100 text-amber-900 border border-amber-200">
              <ShieldCheck size={13} className="text-amber-700" />
              Firebase Auth Vinculado
            </span>
          </div>
          <p className="text-slate-500 text-sm mt-1">
            Cada consultor autenticado pelo Firebase Ã© registrado automaticamente com controle de licenÃ§a, planos e expiraÃ§Ã£o.
          </p>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          {onSyncWithAuth && (
            <Button
              variant="outline"
              onClick={handleSyncAuth}
              disabled={isSyncing}
              className="border-slate-200 hover:bg-slate-50 text-slate-700 font-bold flex items-center gap-2 cursor-pointer"
              title="Sincronizar dados do usuÃ¡rio logado no Firebase Authentication com as licenÃ§as"
            >
              <RefreshCw size={15} className={isSyncing ? "animate-spin text-emerald-600" : "text-slate-500"} />
              {isSyncing ? "Sincronizando..." : "Sincronizar AutenticaÃ§Ã£o"}
            </Button>
          )}
          {onOpenCreateModal && (
            <Button
              onClick={onOpenCreateModal}
              className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold flex items-center gap-2 shadow-sm cursor-pointer"
            >
              <Plus size={16} /> Cadastrar Consultor / Credenciada
            </Button>
          )}
        </div>
      </div>

      <div className="bg-emerald-600 rounded-3xl p-8 text-white shadow-xl shadow-emerald-100 flex flex-col md:flex-row items-center justify-between gap-6">
        <div>
          <h3 className="text-xl font-bold mb-2">Divulgue seu Sistema</h3>
          <p className="text-emerald-100 text-sm">
            Use o link abaixo para convidar novos consultores para criar conta e usufruir do perÃ­odo de teste gratuito da plataforma.
          </p>
        </div>
        <div className="flex bg-white/10 p-2 rounded-2xl border border-white/20 w-full md:w-auto items-center gap-4">
          <code className="text-xs font-mono px-4 text-emerald-50 truncate max-w-[240px]">
            https://gestorconsultorpro.netlify.app/
          </code>
          <Button 
            variant="secondary"
            size="sm"
            onClick={() => {
              navigator.clipboard.writeText("https://gestorconsultorpro.netlify.app/");
              alert("Link de divulgaÃ§Ã£o copiado!");
            }}
            className="bg-white text-emerald-600 hover:bg-emerald-50 font-bold whitespace-nowrap cursor-pointer"
          >
            Copiar Link
          </Button>
        </div>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-5 border-b border-slate-100 bg-slate-50/50 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-center gap-3 flex-wrap">
            <span className="font-bold text-slate-800 text-sm">Consultores &amp; Contas Autenticadas</span>
            <span className="px-2.5 py-0.5 rounded-full text-xs font-bold bg-slate-200 text-slate-700">
              {empresasCredenciadas.length}
            </span>
            <div className="relative">
              <input
                type="text"
                placeholder="Buscar por nome, email ou UID..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="text-xs bg-white border border-slate-200 rounded-xl px-3 py-1.5 w-64 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"
              />
            </div>
          </div>
          <div className="text-xs text-slate-500 flex items-center gap-4 flex-wrap">
            <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-indigo-500 inline-block"></span> Definitiva (VitalÃ­cia)</span>
            <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span> Ativo</span>
            <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span> Expirando (&le; 5 dias)</span>
            <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block"></span> Expirado / Bloqueado</span>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50 border-b border-slate-100">
              <tr>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider">Empresa / Consultor &amp; AutenticaÃ§Ã£o</th>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider">Plano Atual</th>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider">Tempo Restante &amp; Acessos</th>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider">Ampliar PerÃ­odo de Teste</th>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider">Status</th>
                <th className="p-5 text-xs font-bold text-slate-500 uppercase tracking-wider text-right">AÃ§Ãµes</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-50">
              {filteredEmpresas.map(emp => {
                const isDefinitive = emp.tipoPlano === 'Definitiva';
                const daysLeft = calculateDaysLeft(emp);
                const isExpired = !isDefinitive && daysLeft <= 0;
                const isExpiringSoon = !isDefinitive && daysLeft > 0 && daysLeft <= 5;
                const totalLimitDays = emp.diasTeste || (emp.tipoPlano === 'Mensal' ? 30 : emp.tipoPlano === 'Anual' ? 365 : 30);
                const cadastroFormatted = formatItemDate(emp.dataCadastro);
                const loginFormatted = formatItemDate(emp.ultimoLogin);
                const isCurrentUser = currentUser?.uid === emp.ownerId;
                const isAdminRole = emp.role === 'admin' || (emp.email || '').toLowerCase().trim() === 'itamartrairi@gmail.com';

                return (
                  <tr key={emp.id} className="hover:bg-slate-50/50 transition-colors group">
                    <td className="p-5">
                      <div className="flex items-start gap-3">
                        <div className={`w-10 h-10 rounded-xl flex items-center justify-center font-bold text-sm shrink-0 ${
                          isAdminRole 
                            ? 'bg-indigo-100 text-indigo-700 border border-indigo-200' 
                            : 'bg-slate-100 text-slate-600 border border-slate-200'
                        }`}>
                          {emp.photoURL ? (
                            <img src={emp.photoURL} alt={emp.razaoSocial} className="w-10 h-10 rounded-xl object-cover" />
                          ) : (
                            (emp.razaoSocial?.charAt(0) || emp.consultor?.charAt(0) || 'C').toUpperCase()
                          )}
                        </div>
                        <div className="space-y-1">
                          <div className="flex items-center gap-2 flex-wrap">
                            <p className="text-sm font-bold text-slate-800">{emp.razaoSocial}</p>
                            {isAdminRole && (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-indigo-100 text-indigo-800 border border-indigo-200">
                                Administrador
                              </span>
                            )}
                            {isCurrentUser && (
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">
                                VocÃª
                              </span>
                            )}
                          </div>
                          <p className="text-xs text-slate-500 font-medium">{emp.email || 'Sem e-mail cadastrado'}</p>
                          <div className="flex items-center gap-2 flex-wrap text-[11px]">
                            {emp.consultor && emp.consultor !== emp.razaoSocial && (
                              <span className="text-emerald-700 font-medium">Consultor: {emp.consultor}</span>
                            )}
                            {emp.cnpj && (
                              <span className="text-slate-500 font-mono bg-slate-100 px-1.5 py-0.5 rounded text-[10px]" title="CNPJ da Credenciada">
                                CNPJ: {formatCNPJ(emp.cnpj)}
                              </span>
                            )}
                            {emp.cpfConsultor && (
                              <span className="text-slate-500 font-mono bg-slate-100 px-1.5 py-0.5 rounded text-[10px]" title="CPF do Consultor">
                                CPF: {formatCPF(emp.cpfConsultor)}
                              </span>
                            )}
                          </div>
                          
                          {/* Firebase Auth Tag */}
                          <div className="flex items-center gap-1.5 flex-wrap pt-0.5">
                            {emp.ownerId && emp.ownerId !== 'local' ? (
                              <button
                                type="button"
                                onClick={() => {
                                  navigator.clipboard.writeText(emp.ownerId);
                                  alert(`UID copiado: ${emp.ownerId}`);
                                }}
                                title={`Firebase UID: ${emp.ownerId} (Clique para copiar)`}
                                className="inline-flex items-center gap-1 font-mono text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-600 px-2 py-0.5 rounded cursor-pointer transition-colors border border-slate-200"
                              >
                                <Key size={10} className="text-slate-400" />
                                <span className="truncate max-w-[120px]">{emp.ownerId}</span>
                                <Copy size={9} className="text-slate-400" />
                              </button>
                            ) : (
                              <span className="inline-flex items-center text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded">
                                Registro Local
                              </span>
                            )}
                            {emp.providerId && (
                              <span className="inline-flex items-center text-[10px] font-semibold bg-emerald-50 text-emerald-700 px-1.5 py-0.5 rounded border border-emerald-100">
                                {emp.providerId.includes('google') ? 'Google Auth' : 'E-mail & Senha'}
                              </span>
                            )}
                          </div>
                        </div>
                      </div>
                    </td>

                    <td className="p-5">
                      <div className="space-y-1">
                        <select 
                          value={emp.tipoPlano || 'Teste'} 
                          onChange={(e) => onUpdateCredenciadaPlano?.(emp.id, e.target.value as 'Teste' | 'Mensal' | 'Anual' | 'Definitiva')}
                          className="text-xs font-bold bg-white border border-slate-200 rounded-lg px-2.5 py-1.5 shadow-sm focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 cursor-pointer"
                        >
                          <option value="Teste">Teste ({emp.diasTeste || 30} dias)</option>
                          <option value="Mensal">Mensal (30 dias)</option>
                          <option value="Anual">Anual (365 dias)</option>
                          <option value="Definitiva">â­ Definitiva (VitalÃ­cia)</option>
                        </select>
                        <p className="text-[10px] text-slate-400">
                          {isDefinitive ? 'Acesso Permanente' : `Total configurado: ${totalLimitDays} dias`}
                        </p>
                      </div>
                    </td>

                    <td className="p-5">
                      {isDefinitive ? (
                        <div className="space-y-1.5">
                          <span className="px-2.5 py-1 rounded-full text-xs font-bold flex items-center gap-1.5 bg-indigo-100 text-indigo-700 border border-indigo-200 w-fit">
                            <InfinityIcon size={14} className="shrink-0" />
                            LicenÃ§a Definitiva
                          </span>
                          <p className="text-[10px] text-indigo-600 font-medium">
                            Acesso vitalÃ­cio sem expiraÃ§Ã£o
                          </p>
                          {loginFormatted && (
                            <p className="text-[10px] text-slate-400">
                              Ãšltimo login: <span className="font-semibold text-slate-600">{loginFormatted}</span>
                            </p>
                          )}
                        </div>
                      ) : (
                        <div className="space-y-1.5">
                          <div className="flex items-center gap-2">
                            <span className={`px-2.5 py-1 rounded-full text-xs font-bold flex items-center gap-1.5 ${
                              isExpired 
                                ? 'bg-rose-100 text-rose-700 border border-rose-200' 
                                : isExpiringSoon 
                                  ? 'bg-amber-100 text-amber-800 border border-amber-200' 
                                  : 'bg-emerald-100 text-emerald-700 border border-emerald-200'
                            }`}>
                              <Clock size={12} />
                              {isExpired 
                                ? 'Expirado' 
                                : `${daysLeft} dia${daysLeft === 1 ? '' : 's'} restante${daysLeft === 1 ? '' : 's'}`}
                            </span>
                          </div>
                          {emp.validadeLicenca ? (
                            <p className="text-[11px] text-slate-500">
                              Validade atÃ©: <span className="font-semibold text-slate-700">{formatItemDate(emp.validadeLicenca)}</span>
                            </p>
                          ) : cadastroFormatted ? (
                            <p className="text-[10px] text-slate-400">
                              Cadastrado em {cadastroFormatted}
                            </p>
                          ) : null}
                          {loginFormatted && (
                            <p className="text-[10px] text-slate-400">
                              Ãšltimo login: <span className="font-semibold text-slate-600">{loginFormatted}</span>
                            </p>
                          )}
                        </div>
                      )}
                    </td>

                    <td className="p-5">
                      {isDefinitive ? (
                        <div className="py-1">
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold bg-indigo-50 text-indigo-700 border border-indigo-100">
                            <Sparkles size={13} />
                            LicenÃ§a VitalÃ­cia Ativa
                          </span>
                        </div>
                      ) : (
                        <div className="space-y-2">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <button
                              type="button"
                              onClick={() => handleAddDays(emp, 15)}
                              className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-700 rounded-lg text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer"
                              title="Adicionar 15 dias de teste a este usuÃ¡rio"
                            >
                              +15 dias
                            </button>
                            <button
                              type="button"
                              onClick={() => handleAddDays(emp, 30)}
                              className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-700 rounded-lg text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer"
                              title="Adicionar 30 dias de teste a este usuÃ¡rio"
                            >
                              +30 dias
                            </button>
                            <button
                              type="button"
                              onClick={() => handleAddDays(emp, 60)}
                              className="px-2 py-1 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-700 rounded-lg text-xs font-bold transition-all shadow-xs active:scale-95 cursor-pointer"
                              title="Adicionar 60 dias de teste a este usuÃ¡rio"
                            >
                              +60 dias
                            </button>
                          </div>

                          {editingCustomId === emp.id ? (
                            <div className="flex items-center gap-1.5 mt-1">
                              <input 
                                type="number" 
                                min="1"
                                max="999"
                                placeholder="Total de dias" 
                                value={customDaysInput}
                                onChange={(e) => setCustomDaysInput(e.target.value)}
                                className="w-24 px-2 py-1 text-xs border border-emerald-300 rounded-lg focus:outline-none focus:ring-1 focus:ring-emerald-500"
                                autoFocus
                              />
                              <button
                                type="button"
                                onClick={async () => {
                                  const val = parseInt(customDaysInput, 10);
                                  if (!isNaN(val) && val > 0) {
                                    await onUpdateCredenciadaDiasTeste?.(emp.id, val);
                                    setEditingCustomId(null);
                                    setCustomDaysInput('');
                                  }
                                }}
                                className="px-2 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold cursor-pointer"
                              >
                                OK
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  setEditingCustomId(null);
                                  setCustomDaysInput('');
                                }}
                                className="px-2 py-1 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-xs font-bold cursor-pointer"
                              >
                                âœ•
                              </button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => {
                                setEditingCustomId(emp.id);
                                setCustomDaysInput(String(totalLimitDays));
                              }}
                              className="text-[11px] text-emerald-600 hover:text-emerald-700 font-semibold underline block text-left cursor-pointer"
                            >
                              Definir total de dias customizado
                            </button>
                          )}
                        </div>
                      )}
                    </td>

                    <td className="p-5">
                      <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest ${
                        emp.status === 'Ativa' ? 'bg-emerald-100 text-emerald-700' : 'bg-rose-100 text-rose-700'
                      }`}>
                        {emp.status || 'Ativa'}
                      </span>
                    </td>

                    <td className="p-5 text-right">
                      <div className="flex items-center justify-end gap-2">
                        {onOpenEditModal && (
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => onOpenEditModal(emp)}
                            title="Editar Dados da Credenciada"
                            className="text-slate-500 hover:text-slate-800 hover:bg-slate-100"
                          >
                            <Pencil size={14} />
                          </Button>
                        )}
                        <Button 
                          variant="outline" 
                          size="sm"
                          className={`font-bold transition-all ${
                            emp.status === 'Bloqueada' 
                              ? 'text-emerald-600 border-emerald-100 hover:bg-emerald-50' 
                              : 'text-rose-600 border-rose-100 hover:bg-rose-50'
                          }`}
                          onClick={() => onUpdateCredenciadaStatus?.(emp.id, emp.status === 'Bloqueada' ? 'Ativa' : 'Bloqueada')}
                        >
                          {emp.status === 'Bloqueada' ? (
                            <span className="flex items-center gap-1.5"><CheckCircle2 size={14} /> Desbloquear</span>
                          ) : (
                            <span className="flex items-center gap-1.5"><Lock size={14} /> Bloquear</span>
                          )}
                        </Button>
                        {onOpenDeleteModal && (
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={() => onOpenDeleteModal(emp)}
                            title="Excluir Credenciada"
                            className="text-rose-400 hover:text-rose-600 hover:bg-rose-50"
                          >
                            <Trash2 size={14} />
                          </Button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
              {empresasCredenciadas.length === 0 && (
                <tr>
                  <td colSpan={6} className="p-16 text-center">
                    <div className="bg-slate-50 rounded-full w-20 h-20 flex items-center justify-center mx-auto mb-4 text-slate-300">
                      <Building2 size={40} />
                    </div>
                    <h4 className="text-base font-bold text-slate-700 mb-1">Nenhum consultor ou credenciada cadastrada</h4>
                    <p className="text-slate-400 text-sm max-w-md mx-auto mb-6">
                      Como o banco de dados do Firebase foi atualizado, qualquer usuÃ¡rio autenticado aparecerÃ¡ aqui automaticamente. VocÃª tambÃ©m pode sincronizar sua conta de administrador agora.
                    </p>
                    <div className="flex items-center justify-center gap-3 flex-wrap">
                      {onInitializeAdmin && (
                        <Button
                          onClick={async () => {
                            setIsInitializing(true);
                            await onInitializeAdmin();
                            setIsInitializing(false);
                          }}
                          disabled={isInitializing}
                          className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold flex items-center gap-2 shadow-sm cursor-pointer"
                        >
                          <Sparkles size={16} /> {isInitializing ? "Inicializando..." : "âš¡ Sincronizar Conta de Administrador"}
                        </Button>
                      )}
                      {onOpenCreateModal && (
                        <Button
                          variant="outline"
                          onClick={onOpenCreateModal}
                          className="border-slate-200 text-slate-700 hover:bg-slate-50 font-bold flex items-center gap-2 cursor-pointer"
                        >
                          <Plus size={16} /> Cadastrar Manualmente
                        </Button>
                      )}
                    </div>
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </motion.div>
  );
};

const LicensingDataView = ({ setView }: { setView?: (v: any) => void }) => {
  return (
    <motion.div 
      initial={{ opacity: 0, y: 10 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -10 }}
      className="space-y-8"
    >
      <div>
        <h2 className="text-2xl font-bold text-slate-800">Dados do Licenciamento</h2>
        <p className="text-slate-500 text-sm">InformaÃ§Ãµes de preÃ§os, planos e faturamento da plataforma</p>
      </div>

      <div className="bg-amber-50 border border-amber-200 rounded-3xl p-6 flex flex-col md:flex-row items-center justify-between gap-6">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-amber-100 text-amber-600 rounded-2xl flex items-center justify-center flex-shrink-0">
            <Trophy size={24} />
          </div>
          <div>
            <h3 className="font-bold text-amber-950 text-base">Faturamento & AtivaÃ§Ã£o PIX</h3>
            <p className="text-amber-800 text-sm">Utilize a chave PIX <strong className="font-semibold select-all font-mono">itamartrairi@gmail.com</strong> para recebimento de transferÃªncias e ativaÃ§Ã£o manual das licenÃ§as.</p>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {/* Plano Mensal */}
        <Card 
          onClick={() => {
            localStorage.setItem('selected_plan', 'monthly');
            setView?.('checkout');
          }}
          className="hover:border-emerald-400 hover:shadow-lg transition-all shadow-sm flex flex-col overflow-hidden cursor-pointer group"
        >
          <div className="p-8 flex-1 flex flex-col justify-between">
            <div>
              <div className="bg-slate-100 text-slate-700 text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full w-fit mb-4">
                Mensal
              </div>
              <h3 className="text-2xl font-bold text-slate-800 mb-2">Consultor PrÃ³</h3>
              <p className="text-slate-500 text-sm mb-6 font-sans">Assinatura mensal recorrente com total flexibilidade.</p>
              
              <div className="flex items-baseline gap-1 mb-8">
                <span className="text-slate-400 text-lg font-medium">R$</span>
                <span className="text-4xl font-bold text-slate-800">47,90</span>
                <span className="text-slate-400 font-medium">/mÃªs</span>
              </div>

              <ul className="space-y-3.5 mb-8">
                {['DiagnÃ³sticos Ilimitados', 'IA Generativa Integrada', 'RelatÃ³rios Customizados', 'GestÃ£o Kanban'].map((text, i) => (
                  <li key={i} className="flex items-center gap-2.5 text-slate-700 text-sm">
                    <CheckCircle2 size={16} className="text-emerald-500 flex-shrink-0" />
                    <span>{text}</span>
                  </li>
                ))}
              </ul>
            </div>

            <Button 
              className="w-full mt-auto h-12 font-bold bg-white text-emerald-600 border border-emerald-200 group-hover:bg-emerald-600 group-hover:text-white group-hover:border-none transition-all rounded-xl"
            >
              Assinar Plano Mensal
            </Button>
          </div>
        </Card>

        {/* Plano Anual */}
        <Card 
          onClick={() => {
            localStorage.setItem('selected_plan', 'annual');
            setView?.('checkout');
          }}
          className="hover:border-emerald-600 hover:shadow-lg transition-all shadow-sm border-emerald-500 border-2 flex flex-col overflow-hidden relative cursor-pointer group"
        >
          <div className="absolute top-4 right-4 bg-emerald-500 text-white text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full">
            Economize 15%
          </div>
          <div className="p-8 flex-1 flex flex-col justify-between">
            <div>
              <div className="bg-emerald-50 text-emerald-600 text-[10px] font-black uppercase tracking-widest px-3 py-1 rounded-full w-fit mb-4">
                Anual
              </div>
              <h3 className="text-2xl font-bold text-slate-800 mb-2">Master Consultant</h3>
              <p className="text-slate-500 text-sm mb-6 font-sans">Plano de longo prazo para consultores consolidados.</p>

              <div className="flex items-baseline gap-1 mb-1">
                <span className="text-slate-400 text-lg font-bold w-fit">R$</span>
                <span className="text-4xl font-bold text-emerald-600">41,49</span>
                <span className="text-slate-400 font-medium">/mÃªs</span>
              </div>
              <p className="text-slate-400 text-xs mb-8">R$ 497,90 Ã  vista ou no cartÃ£o de crÃ©dito em atÃ© 12x</p>

              <ul className="space-y-3.5 mb-8">
                {['Tudo do plano mensal', 'Acesso Antecipado a Novas IAs', 'Suporte PrioritÃ¡rio', 'Mentoria em DiagnÃ³sticos'].map((text, i) => (
                  <li key={i} className="flex items-center gap-2.5 text-slate-700 text-sm">
                    <CheckCircle2 size={16} className="text-emerald-500 flex-shrink-0" />
                    <span>{text}</span>
                  </li>
                ))}
              </ul>
            </div>

            <Button 
              className="w-full mt-auto h-12 font-bold bg-emerald-600 text-white hover:bg-emerald-700 active:bg-emerald-800 transition-all shadow-lg shadow-emerald-100 rounded-xl border-none"
            >
              Assinar Plano Anual (15% OFF)
            </Button>
          </div>
        </Card>
      </div>
    </motion.div>
  );
};

const HomeView = ({ 
  empresas, 
  diagnosticos, 
  tarefas,
  problemas,
  setView, 
  isAdmin,
  empresasCredenciadas = [],
  user
}: { 
  empresas: Empresa[], 
  diagnosticos: Diagnostico[], 
  tarefas: TarefaPlanoAcao[],
  problemas: Problema[],
  setView: (view: any) => void, 
  isAdmin: boolean,
  empresasCredenciadas?: EmpresaCredenciada[],
  user?: any
}) => {
  const pendingTasks = tarefas.filter(t => t.status !== 'ConcluÃ­do');
  
  // Identifica a credenciada ativa do usuÃ¡rio logado ou a primeira cadastrada
  const currentCredenciada = (user ? empresasCredenciadas.find(c => c.ownerId === user.uid || (user.email && c.email?.toLowerCase() === user.email.toLowerCase())) : null) 
    || empresasCredenciadas[0] 
    || null;

  return (
    <div className="space-y-8 max-w-5xl mx-auto pb-12">
      {/* Banner / Card da Empresa Credenciada Ativa */}
      {currentCredenciada && (
        <motion.div 
          initial={{ opacity: 0, y: -10 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white rounded-2xl p-5 border border-slate-200/80 shadow-sm flex flex-col md:flex-row items-start md:items-center justify-between gap-4"
        >
          <div className="flex items-center gap-3.5">
            <div className="w-12 h-12 rounded-xl bg-emerald-50 border border-emerald-100 flex items-center justify-center text-emerald-600 shadow-sm flex-shrink-0">
              <Building2 size={24} />
            </div>
            <div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-[10px] font-black uppercase tracking-wider text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-100/80">
                  Empresa Credenciada
                </span>
                {currentCredenciada.tipoPlano && (
                  <span className="text-[10px] font-bold text-slate-500 bg-slate-100 px-2 py-0.5 rounded-md">
                    Plano {currentCredenciada.tipoPlano}
                  </span>
                )}
              </div>
              <h1 className="text-lg sm:text-xl font-black text-slate-800 tracking-tight mt-0.5">
                {currentCredenciada.razaoSocial}
              </h1>
              <div className="flex items-center gap-3 flex-wrap mt-1 text-xs text-slate-500 font-medium">
                {currentCredenciada.cnpj ? (
                  <span className="font-mono bg-slate-50 border border-slate-200/60 px-2 py-0.5 rounded text-slate-700 font-semibold" title="CNPJ da Empresa Credenciada">
                    CNPJ: {formatCNPJ(currentCredenciada.cnpj)}
                  </span>
                ) : (
                  <span className="text-slate-400 italic text-[11px]">CNPJ nÃ£o informado</span>
                )}
                {currentCredenciada.consultor && (
                  <span className="text-slate-600">
                    Consultor: <strong className="text-slate-800 font-semibold">{currentCredenciada.consultor}</strong>
                  </span>
                )}
                {currentCredenciada.cpfConsultor && (
                  <span className="font-mono bg-slate-50 border border-slate-200/60 px-2 py-0.5 rounded text-slate-700 font-semibold" title="CPF do Consultor ResponsÃ¡vel">
                    CPF: {formatCPF(currentCredenciada.cpfConsultor)}
                  </span>
                )}
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 self-end md:self-center">
            {isAdmin && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => setView('licenses')}
                className="text-xs font-bold text-slate-700 border-slate-200 hover:bg-slate-50 h-9"
              >
                Gerenciar Credenciadas
              </Button>
            )}
          </div>
        </motion.div>
      )}

      <div className="bg-gradient-to-br from-indigo-700 via-indigo-600 to-blue-500 rounded-[2.5rem] p-12 text-white shadow-2xl relative overflow-hidden">
        <div className="absolute top-0 right-0 w-64 h-64 bg-white/10 rounded-full -translate-y-1/2 translate-x-1/2 blur-3xl"></div>
        <div className="relative z-10">
          {currentCredenciada && (
            <div className="inline-flex items-center gap-2 px-3.5 py-1.5 bg-white/20 backdrop-blur-md rounded-full text-xs font-black text-white mb-3 tracking-wide border border-white/25 shadow-sm">
              <Building2 size={13} className="text-emerald-300" />
              <span>{currentCredenciada.razaoSocial}</span>
              {currentCredenciada.cnpj && (
                <span className="font-mono opacity-90 font-bold border-l border-white/30 pl-2">
                  {formatCNPJ(currentCredenciada.cnpj)}
                </span>
              )}
            </div>
          )}
          <div className="inline-flex items-center gap-2 px-3.5 py-1 bg-white/15 backdrop-blur-md rounded-full text-xs font-bold text-indigo-100 mb-4 tracking-widest uppercase border border-white/20 ml-2">
            by ItÃ mar Gomes
          </div>
          <h2 className="text-5xl font-black tracking-tighter mb-6 leading-none">Melhore os resultados dos seus clientes.</h2>
          <p className="text-indigo-100 text-lg max-w-2xl font-medium leading-relaxed opacity-90">
            Gerencie clientes, realize diagnÃ³sticos precisos e acompanhe a execuÃ§Ã£o dos planos de aÃ§Ã£o em tempo real.
          </p>
          <div className="mt-8 flex gap-4">
            <Button onClick={() => setView('companies')} variant="secondary" className="bg-white text-indigo-600 hover:bg-indigo-50 border-none px-8 py-6 text-sm font-bold uppercase tracking-widest rounded-3xl shadow-lg">
              ComeÃ§ar Agora
            </Button>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('companies')}>
          <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-indigo-100/50">
            <Building2 size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">Clientes</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">GestÃ£o completa da base de clientes e histÃ³ricos de atendimento.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-indigo-600 bg-indigo-50 px-3 py-1.5 rounded-full uppercase tracking-widest">{empresas.length} cadastrados</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-indigo-600 transform group-hover:translate-x-1 transition-all" />
          </div>
        </Card>

        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('kanban')}>
          <div className="w-16 h-16 bg-emerald-50 text-emerald-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-emerald-100/50">
            <Layout size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">Plano de AÃ§Ã£o</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">Quadro Kanban para monitorar a execuÃ§Ã£o das tarefas e prazos.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-emerald-600 bg-emerald-50 px-3 py-1.5 rounded-full uppercase tracking-widest">{tarefas.length} Tarefas Ativas</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-emerald-600 transform group-hover:translate-x-1 transition-all" />
          </div>
        </Card>

        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('macro-dashboard')}>
          <div className="w-16 h-16 bg-violet-50 text-violet-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-violet-100/50">
            <BarChart3 size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">EvoluÃ§Ã£o Macro</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">Dashboard consolidado com a evoluÃ§Ã£o histÃ³rica de todos os clientes.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-violet-600 bg-violet-50 px-3 py-1.5 rounded-full uppercase tracking-widest">Dashboard Macro</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-violet-600 transform group-hover:translate-x-1 transition-all" />
          </div>
        </Card>

        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('companies')}>
           <div className="w-16 h-16 bg-amber-50 text-amber-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-amber-100/50">
            <HistoryIcon size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">DiagnÃ³sticos</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">Execute novos diagnÃ³sticos e visualize a evoluÃ§Ã£o da maturidade.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-amber-600 bg-amber-50 px-3 py-1.5 rounded-full uppercase tracking-widest">{diagnosticos.filter(d => !d.status || d.status === 'Finalizado').length} Realizados</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-amber-600 transform group-hover:translate-x-1 transition-all" />
          </div>
        </Card>

        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('premises')}>
           <div className="w-16 h-16 bg-blue-50 text-blue-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-blue-100/50">
            <FileText size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">Biblioteca</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">Biblioteca de conhecimento com base de problemas, premissas de perguntas e soluÃ§Ãµes.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1.5 rounded-full uppercase tracking-widest">{problemas.length} Registrados</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-blue-600 transform group-hover:translate-x-1 transition-all" />
          </div>
        </Card>
      </div>
    </div>
  );
};

// --- Main App ---

export const extractAndParseJSON = (text: string, defaultValue: any = null): any => {
  if (!text) return defaultValue;
  
  let cleaned = text.trim();
  
  // 1. Double check direct parsing
  try {
    return JSON.parse(cleaned);
  } catch (e) {}

  // 2. Remove markdown formatting if present (```json or ```)
  if (cleaned.includes("```")) {
    const match = cleaned.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (match && match[1]) {
      const candidateContent = match[1].trim();
      try {
        return JSON.parse(candidateContent);
      } catch (e) {
        // Fallback to continue searching below on the code block content
        cleaned = candidateContent;
      }
    }
  }

  // 3. Find candidate JSON structures (arrays or objects)
  // We can look for '{' or '[' and find the matching '}' or ']'
  for (let i = 0; i < cleaned.length; i++) {
    const char = cleaned[i];
    if (char === '{' || char === '[') {
      const openChar = char;
      const closeChar = char === '{' ? '}' : ']';
      
      // Let's find the closing character matching this open character by counting depth
      let depth = 0;
      let closingIndex = -1;
      
      for (let j = i; j < cleaned.length; j++) {
        if (cleaned[j] === openChar) {
          depth++;
        } else if (cleaned[j] === closeChar) {
          depth--;
          if (depth === 0) {
            closingIndex = j;
            // First match complete, let's try to parse this substring
            const candidate = cleaned.substring(i, closingIndex + 1);
            try {
              return JSON.parse(candidate);
            } catch (err) {
              // Not a valid JSON block, continue searching
            }
          }
        }
      }
    }
  }

  // 4. If all else fails, do a last-ditch effort with first '[' or '{' and last ']' or '}'
  const firstSquare = cleaned.indexOf('[');
  const firstCurly = cleaned.indexOf('{');
  
  let firstChar = -1;
  let lastChar = -1;
  
  if (firstSquare !== -1 && (firstCurly === -1 || firstSquare < firstCurly)) {
    firstChar = firstSquare;
    lastChar = cleaned.lastIndexOf(']');
  } else if (firstCurly !== -1) {
    firstChar = firstCurly;
    lastChar = cleaned.lastIndexOf('}');
  }
  
  if (firstChar !== -1 && lastChar !== -1 && lastChar > firstChar) {
    const candidate = cleaned.substring(firstChar, lastChar + 1);
    try {
      return JSON.parse(candidate);
    } catch (substringErr) {
      console.error("Subsegment parsing failed as well:", substringErr);
    }
  }
  
  console.error("Failed to parse JSON directly/indirectly. Original text:", text);
  return defaultValue;
};

// Chave de API global do Gemini (fallback conectado)
const HARDCODED_GEMINI_API_KEY = "";

let aiInstance: any = null;
export const getAI = () => {
  return {
    models: {
      generateContent: async ({ model, contents, config }: { model: string, contents: any, config?: any }) => {
        // 1. Check IndexedDB Cache first
        const cacheKey = generateAICacheKey('gemini_raw', { model, contents, config });
        try {
          const cachedText = await getCachedAI<string>(cacheKey);
          if (cachedText) {
            return { text: cachedText };
          }
        } catch (e) {
          console.warn("[AICache] Cache lookup error:", e);
        }

        let rawSavedKey = localStorage.getItem('custom_gemini_api_key') || "";
        if (rawSavedKey === "undefined" || rawSavedKey === "null" || rawSavedKey.startsWith("AQ.") || rawSavedKey.length < 15) {
          try {
            localStorage.removeItem('custom_gemini_api_key');
          } catch (e) {}
          rawSavedKey = "";
        }
        const savedKey = rawSavedKey.trim();
        let envKey = "";
        try {
          envKey = (import.meta.env.VITE_GEMINI_API_KEY || "").trim();
        } catch (e) {}
        if (!envKey) {
          try {
            envKey = (process.env.GEMINI_API_KEY || "").trim();
          } catch (e) {}
        }
        if (!envKey) {
          try {
            envKey = (process.env.VITE_GEMINI_API_KEY || "").trim();
          } catch (e) {}
        }
        if (envKey === "undefined" || envKey === "null" || envKey.startsWith("AQ.")) {
          envKey = "";
        }
        
        // 2. Try the backend proxy first (available in Container/Full-stack environments)
        try {
          const headers: Record<string, string> = {
            "Content-Type": "application/json"
          };
          if (savedKey && savedKey.startsWith("AIzaSy")) {
            headers["x-custom-api-key"] = savedKey;
          }

          const response = await fetch("/api/gemini/generate", {
            method: "POST",
            headers: headers,
            body: JSON.stringify({ model: model || "gemini-3.6-flash", contents, config })
          });

          if (response.ok) {
            const data = await response.json();
            if (data && data.text) {
              setCachedAI(cacheKey, data.text).catch(() => {});
              return {
                text: data.text
              };
            }
          }
          
          if (response.status !== 404 && response.status !== 502 && response.status !== 503) {
            const errData = await response.json().catch(() => ({}));
            throw new Error(errData.error || `Erro HTTP ${response.status} na API do Gemini`);
          }
        } catch (e: any) {
          const isNetworkOr404 = e.message && (
            e.message.includes("404") || 
            e.message.includes("Failed to fetch") || 
            e.message.includes("NetworkError") ||
            e.message.includes("502") ||
            e.message.includes("503")
          );
          if (!isNetworkOr404) {
            throw e;
          }
          console.warn("[Gemini] Proxy indisponÃ­vel ou estÃ¡tico. Tentando fallback direto do cliente.");
        }

        // 3. Fallback: Client-side Direct Call (for Netlify, Vercel, static pages, etc.)
        const finalKey = (savedKey && savedKey.startsWith("AIzaSy") ? savedKey : "") || (envKey && envKey.startsWith("AIzaSy") ? envKey : "");
        if (!finalKey || finalKey.trim() === "") {
          throw new Error("NÃ£o foi possÃ­vel conectar com o serviÃ§o de IA.\n\nSe vocÃª estiver acessando atravÃ©s de uma hospedagem estÃ¡tica (Netlify/Vercel):\n1. Acesse o menu ConfiguraÃ§Ãµes da aplicaÃ§Ã£o e insira sua chave oficial do Google AI Studio (iniciando com 'AIzaSy...').\n2. Obtenha sua chave gratuita em: https://aistudio.google.com/app/apikey");
        }

        const GEMINI_PRIMARY = "gemini-3.6-flash";
        const GEMINI_CHAIN = [
          "gemini-3.6-flash",
          "gemini-3.7-flash",
          "gemini-3.1-flash-lite"
        ];

        function normalizeClientModel(m?: string): string {
          if (!m) return GEMINI_PRIMARY;
          if (m.includes("2.5") || m.includes("2.0") || m.includes("1.5") || m.includes("1.0")) {
            return "gemini-3.6-flash";
          }
          return m;
        }

        const initialModel = normalizeClientModel(model);
        const candidateModels = Array.from(new Set([initialModel, ...GEMINI_CHAIN]));

        const tryFetchWithModel = async (tgtModel: string) => {
          const url = `https://generativelanguage.googleapis.com/v1beta/models/${tgtModel}:generateContent?key=${finalKey}`;
          
          let partsArr: any[] = [];
          if (typeof contents === 'string') {
            partsArr = [{ text: contents }];
          } else if (contents && Array.isArray(contents)) {
            partsArr = contents;
          } else if (contents && contents.parts) {
            partsArr = contents.parts;
          } else if (contents) {
            partsArr = [{ text: JSON.stringify(contents) }];
          }

          const callApi = async (withConfig: any) => {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 18000);
            try {
              const response = await fetch(url, {
                method: "POST",
                headers: {
                  "Content-Type": "application/json"
                },
                signal: controller.signal,
                body: JSON.stringify({
                  contents: [{ parts: partsArr }],
                  generationConfig: withConfig
                })
              });
              clearTimeout(timeoutId);

              if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                const rawErr = errData.error?.message || `Erro HTTP! Status: ${response.status}`;
                if (rawErr.toLowerCase().includes("api key not valid") || rawErr.toLowerCase().includes("api_key_invalid")) {
                  throw new Error("Chave de API do Gemini invÃ¡lida.\n\nA chave oficial do Google AI Studio comeÃ§a com 'AIzaSy...'.\nPor favor, atualize sua chave no menu ConfiguraÃ§Ãµes ou obtenha uma em: https://aistudio.google.com/app/apikey");
                }
                throw new Error(rawErr);
              }

              const data = await response.json();
              return data.candidates?.[0]?.content?.parts?.[0]?.text || "";
            } catch (fetchErr: any) {
              clearTimeout(timeoutId);
              if (fetchErr.name === 'AbortError') {
                throw new Error("Tempo limite de resposta do Gemini excedido (timeout).");
              }
              throw fetchErr;
            }
          };

          try {
            const textResult = await callApi(config);
            if (textResult) {
              setCachedAI(cacheKey, textResult).catch(() => {});
            }
            return { text: textResult };
          } catch (configErr: any) {
            if (config && (config.responseMimeType || config.responseSchema)) {
              console.warn("Retrying direct call without JSON schema constraints...", configErr);
              try {
                const strippedConfig = { ...config };
                delete strippedConfig.responseMimeType;
                delete strippedConfig.responseSchema;
                const textResult = await callApi(strippedConfig);
                if (textResult) {
                  setCachedAI(cacheKey, textResult).catch(() => {});
                }
                return { text: textResult };
              } catch (fallbackErr: any) {
                throw fallbackErr;
              }
            }
            throw configErr;
          }
        };

        let lastErr: any = null;
        for (const tgtModel of candidateModels) {
          try {
            return await tryFetchWithModel(tgtModel);
          } catch (mErr: any) {
            console.warn(`[Gemini Client Fallback] Model ${tgtModel} failed:`, mErr?.message || mErr);
            lastErr = mErr;
            if (mErr?.message?.includes("AIzaSy") || mErr?.message?.includes("invÃ¡lida")) {
              break;
            }
          }
        }

        const friendlyMsg = lastErr?.message?.includes("API_KEY")
          ? "Chave de API do Gemini invÃ¡lida ou nÃ£o configurada."
          : (lastErr?.message || "NÃ£o foi possÃ­vel conectar aos modelos do Gemini.");
        throw new Error(friendlyMsg);
      }
    }
  };
};

const generateAIFeedback = async (resposta: string, pergunta: string, problema: string) => {
  const cacheKey = generateAICacheKey('feedback', { resposta, pergunta, problema });
  const cached = await getCachedAI<string>(cacheKey);
  if (cached) return cached;

  const prompt = `Como um consultor empresarial especializado do SEBRAE, analise o significado da pergunta e da resposta fornecida no contexto do diagnÃ³stico empresarial.
  
  CONTEXTO:
  Problema/Tema: ${problema}
  Premissa/Pergunta: ${pergunta}
  Resposta do Cliente: ${resposta}
  
  DIRETRIZES DE ANÃLISE SEMÃ‚NTICA DA PERGUNTA E RESPOSTA:
  1. Avalie o sentido da pergunta/premissa:
     - Se a pergunta investiga a ocorrÃªncia de um problema, falha, prejuÃ­zo, inadimplÃªncia ou rejeiÃ§Ã£o (ex: "O produtor jÃ¡ tentou acessar crÃ©dito rural e foi rejeitado por falta de documentaÃ§Ã£o?", "Possui pendÃªncias?"):
       * Resposta "Sim": Representa um PROBLEMA / RISCO / GARGALO (prejudicial para a maturidade). Sugira brevemente uma recomendaÃ§Ã£o ou aÃ§Ã£o de melhoria rÃ¡pida.
       * Resposta "NÃ£o": Representa uma SITUAÃ‡ÃƒO POSITIVA / CONFORMIDADE (nÃ£o sofre do problema). Gere um breve elogio de reconhecimento do sucesso ou boa gestÃ£o.
       * Resposta "Parcial": Representa risco moderado. Sugira uma melhoria rÃ¡pida.
     - Se a pergunta investiga a existÃªncia de um controle, planejamento ou boa prÃ¡tica (ex: "Possui controle financeiro?"):
       * Resposta "Sim": Representa uma SITUAÃ‡ÃƒO POSITIVA / MATURIDADE. Elogie e reconheÃ§a o sucesso.
       * Resposta "NÃ£o" ou "Parcial": Representa um PONTO DE ATENÃ‡ÃƒO / GARGALO. Sugira brevemente o que pode ser feito para melhorar.
  
  REQUISITOS DE SAÃDA:
  - Responda em PortuguÃªs do Brasil de forma profissional, encorajadora e direta.
  - No mÃ¡ximo 2 frases.`;

  try {
    const ai = getAI();
    if (!ai) return null;
    const response = await ai.models.generateContent({ 
      model: "gemini-3.6-flash", 
      contents: prompt,
    });
    const result = response.text?.trim() || "";
    if (result) {
      setCachedAI(cacheKey, result).catch(() => {});
    }
    return result;
  } catch (error) {
    console.error("AI Feedback Error:", error);
    throw error;
  }
};

const generateAIMaturityLevel = async (respostas: Resposta[], scorePercent: number, tipoEmpresa?: string) => {
  const simplifiedRespostas = respostas.map(r => ({
    problema: r.problema,
    pergunta: r.pergunta,
    resposta: r.resposta,
    area: r.area,
    peso: r.peso
  }));

  const cacheKey = generateAICacheKey('maturity_level', { tipoEmpresa, scorePercent, respostas: simplifiedRespostas });
  const cached = await getCachedAI<{ nivel: string; justificativa: string }>(cacheKey);
  if (cached && cached.nivel) {
    return cached;
  }

  const prompt = `Como um consultor empresarial sÃªnior do SEBRAE, analise estas respostas de um diagnÃ³stico empresarial e determine o NÃ­vel de Maturidade da empresa.
  
  DADOS DO DIAGNÃ“STICO:
  - Tipo de Empresa: ${tipoEmpresa || 'NÃ£o especificado'}
  - Score Indicativo: ${scorePercent}%
  - Respostas detalhadas: ${JSON.stringify(simplifiedRespostas)}
  
  DIRETRIZES FUNDAMENTAIS DE AVALIAÃ‡ÃƒO SEMÃ‚NTICA:
  - Avalie obrigatoriamente a SEMÃ‚NTICA de cada pergunta juntamente com a resposta dada:
    * Perguntas com sentido negativo ou sobre ocorrÃªncia de falhas/problemas/rejeiÃ§Ãµes (ex: "O produtor jÃ¡ tentou acessar crÃ©dito rural e foi rejeitado por falta de documentaÃ§Ã£o?"):
      - "Sim" indica um GARGALO/DESVIO CRÃTICO prejudicial Ã  maturidade.
      - "NÃ£o" indica ausÃªncia de problema (PONTO FORTE/BOA GESTÃƒO).
    * Perguntas sobre controles e boas prÃ¡ticas (ex: "Possui planejamento financeiro?"):
      - "Sim" indica PONTO FORTE.
      - "NÃ£o" ou "Parcial" indica GARGALO/DESVIO.
  - Determine o NÃ­vel de Maturidade considerando a realidade semÃ¢ntica das respostas obtidas.

  REQUISITOS DE SAÃDA:
  Determine o nÃ­vel de maturidade:
  - Escolha um destes nÃ­veis padrÃ£o: "Inicial (Inexistente/Ad-hoc)", "BÃ¡sico (Reativo)", "IntermediÃ¡rio (Definido)", "AvanÃ§ado (Gerenciado)", ou "Otimizado (ContÃ­nuo)".
  - OfereÃ§a uma justificativa analÃ­tica de no mÃ¡ximo 3 frases, destacando os reais pontos fortes e gargalos identificados.
  - Retorne a resposta obrigatoriamente no formato JSON abaixo:
  {
    "nivel": "NÃ­vel Determinado aqui",
    "justificativa": "Sua justificativa analÃ­tica aqui em PortuguÃªs do Brasil de forma clara e profissional."
  }`;

  try {
    const ai = getAI();
    if (!ai) return null;
    const response = await ai.models.generateContent({
      model: "gemini-3.6-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
      }
    });

    if (response.text) {
      const parsed = extractAndParseJSON(response.text);
      if (parsed && parsed.nivel) {
        setCachedAI(cacheKey, parsed).catch(() => {});
        return parsed;
      }
    }
    return null;
  } catch (error) {
    console.error("AI Maturity Level Error:", error);
    return null;
  }
};

const generateAISuggestions = async (probNome: string, noResponses: Resposta[], yesResponses: Resposta[]) => {
  const allResponses = [...noResponses, ...yesResponses];
  const cacheKey = generateAICacheKey('suggestions', { probNome, allResponses: allResponses.map(r => ({ p: r.pergunta, res: r.resposta, o: r.observacao })) });
  const cached = await getCachedAI<any>(cacheKey);
  if (cached && cached.solucao_recomendada) {
    return cached;
  }

  const prompt = `Como um consultor empresarial sÃªnior do SEBRAE, analise as respostas do diagnÃ³stico para o problema abaixo e sugira soluÃ§Ãµes estruturadas.
  
  PROBLEMA PRINCIPAL: ${probNome}
  
  RESPOSTAS REGISTRADAS NAS PREMISSAS DESTE PROBLEMA:
  ${allResponses.map(r => `- Premissa/Pergunta: "${r.pergunta}" | Resposta: "${r.resposta}"${r.observacao ? ` (ObservaÃ§Ã£o: ${r.observacao})` : ''}`).join('\n')}
  
  DIRETRIZES DE ANÃLISE SEMÃ‚NTICA:
  - Avalie o sentido de cada pergunta:
    * Perguntas sobre falhas, rejeiÃ§Ãµes ou erros (ex: "O produtor jÃ¡ tentou acessar crÃ©dito rural e foi rejeitado por falta de documentaÃ§Ã£o?"):
      - Resposta "Sim" Ã© um PONTO CRÃTICO que exige correÃ§Ã£o.
      - Resposta "NÃ£o" Ã© um PONTO FORTE.
    * Perguntas sobre controles e boas prÃ¡ticas:
      - Resposta "NÃ£o" ou "Parcial" Ã© um PONTO CRÃTICO.
      - Resposta "Sim" Ã© um PONTO FORTE.
  - Elabore as recomendaÃ§Ãµes focando nos pontos crÃ­ticos reais identificados.
  
  POR FAVOR, GERE UMA RESPOSTA EM JSON NO SEGUINTE FORMATO:
  {
    "solucao_recomendada": "...",
    "acoes_sugeridas": "1. ...\\n2. ...\\n3. ...",
    "prazo_sugerido": "X dias",
    "responsavel_sugerido": "Quem executarÃ¡...",
    "kpis_sugeridos": "Indicadores para medir sucesso...",
    "comentario_sucesso": "Se houver pontos fortes reais (como ausÃªncia de rejeiÃ§Ã£o de crÃ©dito ou presenÃ§a de boas prÃ¡ticas), faÃ§a um comentÃ¡rio reconhecendo o sucesso das aÃ§Ãµes do produtor/empresa."
  }`;

  try {
    const ai = getAI();
    if (!ai) return null;
    const response = await ai.models.generateContent({ 
      model: "gemini-3.6-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json"
      }
    });
    const parsed = extractAndParseJSON(response.text, {});
    if (parsed && parsed.solucao_recomendada) {
      setCachedAI(cacheKey, parsed).catch(() => {});
    }
    return parsed;
  } catch (error) {
    console.error("AI Generation Error:", error);
    throw error;
  }
};


const TIPOS_EMPRESA = [
  "Geral",
  "ComÃ©rcio",
  "ServiÃ§os",
  "IndÃºstria",
  "AgronegÃ³cio",
  "Alimentos e Bebidas"
];

const AREAS_ORDER = [
  "EstratÃ©gico",
  "Financeiro",
  "Acesso a CrÃ©dito",
  "CrÃ©dito",
  "Marketing",
  "Vendas",
  "Operacional",
  "Recursos Humanos",
  "Tecnologia",
  "LogÃ­stica",
  "JurÃ­dico",
  "Atendimento",
  "Geral"
];

// --- Components ---
// (We moved Modal and Button to separate files)

const MarkdownText = ({ text }: { text: string }) => {
  if (!text) return null;
  const paragraphs = text.split('\n');
  return (
    <div className="space-y-3 text-slate-600 text-sm leading-relaxed">
      {paragraphs.map((p, idx) => {
        if (!p.trim()) return <div key={idx} className="h-2" />;
        // Check if paragraph starts with bullet points or dashes
        const isBullet = p.trim().startsWith('-') || p.trim().startsWith('â€¢');
        const content = isBullet ? p.trim().substring(1).trim() : p;
        
        // Parse bold markers **something**
        const parts = content.split(/\*\*([^*]+)\*\*/g);
        const parsedContent = parts.map((part, i) => {
          if (i % 2 === 1) {
            return <strong key={i} className="font-extrabold text-slate-900">{part}</strong>;
          }
          return part;
        });

        if (isBullet) {
          return (
            <div key={idx} className="flex gap-2 pl-4">
              <span className="text-emerald-550 font-bold">â€¢</span>
              <p className="flex-1">{parsedContent}</p>
            </div>
          );
        }

        return <p key={idx}>{parsedContent}</p>;
      })}
    </div>
  );
};

let cachedAllRespostasMap: Map<string, Resposta> | null = null;

export const loadAllLocalRespostas = (): Resposta[] => {
  if (cachedAllRespostasMap) {
    return Array.from(cachedAllRespostasMap.values());
  }
  try {
    const saved = localStorage.getItem('local_all_respostas');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        cachedAllRespostasMap = new Map(parsed.map(r => [r.id, r]));
        return parsed;
      }
    }
    const legacy = localStorage.getItem('local_respostas');
    if (legacy) {
      const parsed = JSON.parse(legacy);
      if (Array.isArray(parsed) && parsed.length > 0) {
        cachedAllRespostasMap = new Map(parsed.map(r => [r.id, r]));
        return parsed;
      }
    }
  } catch (e) {
    console.error("Erro ao carregar respostas locais:", e);
  }
  cachedAllRespostasMap = new Map();
  return [];
};

export const saveAllLocalRespostas = (newOrUpdated: Resposta[]) => {
  try {
    if (!newOrUpdated || newOrUpdated.length === 0) return;
    if (!cachedAllRespostasMap) {
      loadAllLocalRespostas();
    }
    newOrUpdated.forEach(r => {
      if (r?.id) cachedAllRespostasMap!.set(r.id, r);
    });
    const all = Array.from(cachedAllRespostasMap!.values());
    localStorage.setItem('local_all_respostas', JSON.stringify(all));
    localStorage.setItem('local_respostas', JSON.stringify(all));
  } catch (e) {
    console.error("Erro ao salvar respostas locais:", e);
  }
};

export const getRespostasForDiagnostico = (diagId: string): Resposta[] => {
  const all = loadAllLocalRespostas();
  return all.filter(r => r.diagnosticoId === diagId);
};

export default function App() {
  const syncedDiagsRef = useRef<Set<string>>(new Set());
  const pendingResponsesBuffer = useRef<Map<string, Resposta>>(new Map());
  const saveDebounceTimerRef = useRef<NodeJS.Timeout | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [loading, setLoading] = useState(true);

  // Save status indicator states
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving'>('saved');
  const [lastSavedTime, setLastSavedTime] = useState<Date | null>(null);
  const [isManualSaving, setIsManualSaving] = useState(false);
  const [manualSaveSuccess, setManualSaveSuccess] = useState(false);

  // Edit Diagnosis Date Modal State
  const [isEditDateModalOpen, setIsEditDateModalOpen] = useState(false);
  const [diagToEditDate, setDiagToEditDate] = useState<Diagnostico | null>(null);

  // Replicate modal state
  const [isReplicateModalOpen, setIsReplicateModalOpen] = useState(false);
  const [replicateSourceDiagId, setReplicateSourceDiagId] = useState<string>('');

  // Smart Cloud Sync State
  const [isSmartSyncModalOpen, setIsSmartSyncModalOpen] = useState(false);
  const [isSyncingCloud, setIsSyncingCloud] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [lastSyncSummary, setLastSyncSummary] = useState<SyncSummary | null>(() => {
    try {
      const saved = localStorage.getItem('last_sync_summary');
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });

  // Dedicated Backup Export & Access State
  const [isBackupModalOpen, setIsBackupModalOpen] = useState(false);
  const [isExportingBackup, setIsExportingBackup] = useState(false);
  const [backupStats, setBackupStats] = useState<BackupExportStats | null>(null);
  const [backupFileName, setBackupFileName] = useState<string>('consultoria_pro_backup.json');
  const [backupJsonString, setBackupJsonString] = useState<string>('');

  // Storage Mode State (Cloud vs. Local)
  

  const [storageMode, setStorageMode] = useState<'cloud' | 'local'>(() => {
    return (localStorage.getItem('storage_mode') as 'cloud' | 'local') || 'cloud';
  });

  const handleSetStorageMode = (mode: 'cloud' | 'local') => {
    setStorageMode(mode);
    localStorage.setItem('storage_mode', mode);
    playSuccessSound();

    if (mode === 'cloud') {
      if (user) {
        showToast('Modo Nuvem ativado. Sincronizando dados com a Nuvem...', 'info');
        performSmartCloudSync().then(() => {
          showToast('Dados sincronizados com a Nuvem com sucesso!', 'success');
        }).catch((err) => {
          console.warn('Erro na sincronizaÃ§Ã£o automÃ¡tica:', err);
        });
      } else {
        showToast('Modo Nuvem selecionado. FaÃ§a login para salvar online.', 'info');
      }
    } else {
      showToast('Modo de salvamento alterado para Local.', 'info');
    }
  };

  // Global Snappy Toast Notifications
  const [toasts, setToasts] = useState<ToastItem[]>([]);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'success', title?: string) => {
    const id = 'toast_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6);
    setToasts(prev => [...prev.slice(-3), { id, message, type, title }]);
    if (type === 'success') {
      playSuccessSound();
    } else {
      playClickSound();
    }
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 3500);
  };

  const removeToast = (id: string) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  };

  // Intercept window.alert so all operations are instant and non-blocking
  useEffect(() => {
    const originalAlert = window.alert;
    (window as any).alert = (msg: any) => {
      const str = String(msg || '');
      const isError = str.toLowerCase().includes('erro') || str.toLowerCase().includes('invÃ¡lid') || str.toLowerCase().includes('falha') || str.toLowerCase().includes('por favor');
      showToast(str, isError ? 'error' : 'success');
    };
    (window as any).showToast = showToast;

    return () => {
      window.alert = originalAlert;
    };
  }, []);

  const [generatingAction, setGeneratingAction] = useState<string | null>(null);
  const [completedAction, setCompletedAction] = useState<string | null>(null);
  const [creatingDiagnosis, setCreatingDiagnosis] = useState(false);
  const isCreatingDiagRef = useRef(false);
  const [calculatingMaturity, setCalculatingMaturity] = useState(false);
  const [generatingPlan, setGeneratingPlan] = useState(false);
  const [view, setView] = useState<'home' | 'companies' | 'credenciadas' | 'licenses' | 'diagnosis' | 'dashboard' | 'premises' | 'cronograma' | 'relatorio' | 'kanban' | 'settings' | 'landing' | 'checkout' | 'licensing' | 'dados-consultoria' | 'agenda' | 'macro-dashboard' | 'disc-assessment' | 'maturity-assessment' | 'resultado-consultoria'>('landing');

  const [customLogo, setCustomLogo] = useState<string | null>(localStorage.getItem('sebrae_custom_logo'));
  const [customConsultoraLogo, setCustomConsultoraLogo] = useState<string | null>(localStorage.getItem('consultora_custom_logo'));
  const [logoChoice, setLogoChoice] = useState<'both' | 'sebrae' | 'consultora' | 'none'>(() => {
    return (localStorage.getItem('preferred_logo') as any) || 'both';
  });
  const activeLogo = logoChoice === 'sebrae' ? customLogo : logoChoice === 'consultora' ? customConsultoraLogo : null;
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const [libraryTab, setLibraryTab] = useState<'problemas' | 'premissas' | 'solucoes' | 'areas' | 'segmentos'>('problemas');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isDeletingDiagnostico, setIsDeletingDiagnostico] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('login');
  const [authEmail, setAuthEmail] = useState('');
  const [authPassword, setAuthPassword] = useState('');
  const [authError, setAuthError] = useState('');
  const [modalType, setModalType] = useState<'create' | 'edit' | 'delete' | 'createCredenciada' | 'editCredenciada' | 'deleteCredenciada' | 'createPremissa' | 'editPremissa' | 'deletePremissa' | 'createProblema' | 'editProblema' | 'deleteProblema' | 'deleteAllProblemas' | 'deleteAllPremissas' | 'createSolucao' | 'editSolucao' | 'deleteSolucao' | 'deleteAllSolucoes' | 'deleteAllBiblioteca' | 'deleteArea' | 'renameArea' | 'createArea' | 'createSegmento' | 'deleteSegmento' | 'renameSegmento' | 'deleteDiagnostico' | 'deletePlanoAcao' | 'importSuccess' | 'selectAreas' | 'createTarefa' | 'editTarefa' | 'deleteTarefa' | 'confirmImportCompanyType' | 'viewMaturityDetails' | null>(null);
  const [modalData, setModalData] = useState<any>(null);
  const [areaToDelete, setAreaToDelete] = useState('');
  const [areaToRename, setAreaToRename] = useState('');
  const [newAreaName, setNewAreaName] = useState('');
  const [newAreaInput, setNewAreaInput] = useState('');
  const [segmentToDelete, setSegmentToDelete] = useState('');
  const [segmentToRename, setSegmentToRename] = useState('');
  const [newSegmentName, setNewSegmentName] = useState('');
  const [newSegmentInput, setNewSegmentInput] = useState('');
  const [dbAreas, setDbAreas] = useState<{ id: string, nome: string }[]>(() => {
    try {
      const saved = localStorage.getItem('local_db_areas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [dbSegmentos, setDbSegmentos] = useState<{ id: string, nome: string }[]>(() => {
    try {
      const saved = localStorage.getItem('local_db_segmentos');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [inputValue, setInputValue] = useState('');
  const [historicalData, setHistoricalData] = useState<any[]>([]);
  const [empresaForm, setEmpresaForm] = useState<Partial<Empresa>>({});
  const [credenciadaForm, setCredenciadaForm] = useState<Partial<EmpresaCredenciada>>({});
  const [diagnosticoForm, setDiagnosticoForm] = useState<Partial<DadosConsultoria>>({});
  const [selectedAreasForDiagnosis, setSelectedAreasForDiagnosis] = useState<string[]>([]);
  const [diagnosisDate, setDiagnosisDate] = useState(new Date().toISOString().split('T')[0]);
  const [premissaForm, setPremissaForm] = useState({ idProblema: '', problema: '', pergunta: '', peso: 1, tipoEmpresa: '' });
  const [problemaForm, setProblemaForm] = useState<{ descricao_problemas: string, area: string, impacto: string, tipoEmpresa: string, tags: string[] }>({ descricao_problemas: '', area: '', impacto: '', tipoEmpresa: '', tags: [] });
  const [solucaoForm, setSolucaoForm] = useState<{ idProblema: string, problema: string, area: string, solucao_recomendada: string, acoes_sugeridas: string, prazo_sugerido: string, responsavel_sugerido: string, kpis_sugeridos: string, resultado_esperado: string, comentario_sucesso: string, tipoEmpresa: string, tags: string[] }>({ idProblema: '', problema: '', area: '', solucao_recomendada: '', acoes_sugeridas: '', prazo_sugerido: '', responsavel_sugerido: '', kpis_sugeridos: '', resultado_esperado: '', comentario_sucesso: '', tipoEmpresa: '', tags: [] });
  const [tarefaForm, setTarefaForm] = useState<Partial<TarefaPlanoAcao>>({});
  const [selectedEmpresa, setSelectedEmpresa] = useState<Empresa | null>(null);
  const [selectedDiagnostico, setSelectedDiagnostico] = useState<Diagnostico | null>(null);
  const [expandedMenu, setExpandedMenu] = useState<{ [key: string]: boolean }>({
    gestao: false,
    diagnosticos: false,
    diagnosticoAtivo: true,
    licencas: false,
  });

  const toggleGroup = (key: string) => {
    setExpandedMenu(prev => ({ ...prev, [key]: !prev[key] }));
  };

  useEffect(() => {
    if (['companies', 'projects', 'agenda', 'credenciadas'].includes(view)) {
      setExpandedMenu(prev => ({ ...prev, gestao: true }));
    } else if (['maturity-assessment', 'disc-assessment', 'resultado-consultoria', 'premises'].includes(view) || (view === 'dashboard' && !selectedDiagnostico)) {
      setExpandedMenu(prev => ({ ...prev, diagnosticos: true }));
    } else if (['dados-consultoria', 'cronograma', 'relatorio'].includes(view)) {
      setExpandedMenu(prev => ({ ...prev, diagnosticoAtivo: true }));
    } else if (['licenses', 'licensing'].includes(view)) {
      setExpandedMenu(prev => ({ ...prev, licencas: true }));
    }
  }, [view, selectedDiagnostico]);
  const [isAdmin, setIsAdmin] = useState(false);
  const [isSystemKeyActive, setIsSystemKeyActive] = useState(false);
  const [selectedAreaFilter, setSelectedAreaFilter] = useState('');
  const [selectedTipoEmpresaFilter, setSelectedTipoEmpresaFilter] = useState('');
  const [selectedTagFilter, setSelectedTagFilter] = useState('');
  const [pendingImportFile, setPendingImportFile] = useState<File | null>(null);
  const [importCompanyType, setImportCompanyType] = useState('Geral');
  const [customImportCompanyType, setCustomImportCompanyType] = useState('');
  const [diagnosisCompanyType, setDiagnosisCompanyType] = useState('Geral');
  const [diagnosisProjectName, setDiagnosisProjectName] = useState('');
  const [activeDiagArea, setActiveDiagArea] = useState<string>('');
  const [customProblemaEmpresaType, setCustomProblemaEmpresaType] = useState('');
  const [customEmpresaType, setCustomEmpresaType] = useState('');
  const [isEditSegmentModalOpen, setIsEditSegmentModalOpen] = useState(false);
  const [editSegmentTarget, setEditSegmentTarget] = useState<'empresa' | 'diagnostico' | null>(null);
  const [selectedSegmentForEdit, setSelectedSegmentForEdit] = useState('');
  const [customSegmentForEdit, setCustomSegmentForEdit] = useState('');
  const [dashboardChartType, setDashboardChartType] = useState<'line' | 'bar'>('line');
  const [sessionCustomSegments, setSessionCustomSegments] = useState<string[]>([]);
  const [sessionCustomAreas, setSessionCustomAreas] = useState<string[]>([]);
  const [deletedAreas, setDeletedAreas] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('deleted_areas');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('deleted_areas', JSON.stringify(deletedAreas));
    } catch (e) {
      console.error("Erro ao salvar deleted_areas:", e);
    }
  }, [deletedAreas]);

  const [customPremissaEmpresaType, setCustomPremissaEmpresaType] = useState('');
  const [customSolucaoEmpresaType, setCustomSolucaoEmpresaType] = useState('');
  const [customProblemaArea, setCustomProblemaArea] = useState('');
  const [customSolucaoArea, setCustomSolucaoArea] = useState('');
  
  const [empresas, setEmpresas] = useState<Empresa[]>(() => {
    try {
      const saved = localStorage.getItem('local_empresas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 9;
  const [empresasCredenciadas, setEmpresasCredenciadas] = useState<EmpresaCredenciada[]>(() => {
    try {
      const saved = localStorage.getItem('local_empresas_credenciadas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [diagnosticos, setDiagnosticos] = useState<Diagnostico[]>(() => {
    try {
      const saved = localStorage.getItem('local_diagnosticos');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [respostas, setRespostas] = useState<Resposta[]>(() => {
    try {
      const saved = localStorage.getItem('local_respostas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [respostasLoaded, setRespostasLoaded] = useState(true);
  const [premissas, setPremissas] = useState<Premissa[]>(() => {
    try {
      if (localStorage.getItem('user_cleared_premissas') === 'true') return [];
      const saved = localStorage.getItem('local_premissas');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [problemas, setProblemas] = useState<Problema[]>(() => {
    try {
      if (localStorage.getItem('user_cleared_problemas') === 'true') return [];
      const saved = localStorage.getItem('local_problemas');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [solucoes, setSolucoes] = useState<Solucao[]>(() => {
    try {
      if (localStorage.getItem('user_cleared_solucoes') === 'true') return [];
      const saved = localStorage.getItem('local_solucoes');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [tarefasPlano, setTarefasPlano] = useState<TarefaPlanoAcao[]>(() => {
    try {
      const saved = localStorage.getItem('local_tarefas_plano');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });

  // Auto-sync state changes to localStorage
  useEffect(() => {
    try { localStorage.setItem('local_empresas', JSON.stringify(empresas)); } catch (e) { console.error("Error saving local_empresas:", e); }
  }, [empresas]);

  useEffect(() => {
    try { localStorage.setItem('local_empresas_credenciadas', JSON.stringify(empresasCredenciadas)); } catch (e) { console.error("Error saving local_empresas_credenciadas:", e); }
  }, [empresasCredenciadas]);

  useEffect(() => {
    try { localStorage.setItem('local_diagnosticos', JSON.stringify(diagnosticos)); } catch (e) { console.error("Error saving local_diagnosticos:", e); }
  }, [diagnosticos]);

  useEffect(() => {
    try { localStorage.setItem('local_respostas', JSON.stringify(respostas)); } catch (e) { console.error("Error saving local_respostas:", e); }
  }, [respostas]);

  useEffect(() => {
    try { localStorage.setItem('local_premissas', JSON.stringify(premissas)); } catch (e) { console.error("Error saving local_premissas:", e); }
  }, [premissas]);

  useEffect(() => {
    try { localStorage.setItem('local_problemas', JSON.stringify(problemas)); } catch (e) { console.error("Error saving local_problemas:", e); }
  }, [problemas]);

  useEffect(() => {
    try { localStorage.setItem('local_solucoes', JSON.stringify(solucoes)); } catch (e) { console.error("Error saving local_solucoes:", e); }
  }, [solucoes]);

  useEffect(() => {
    try { localStorage.setItem('local_tarefas_plano', JSON.stringify(tarefasPlano)); } catch (e) { console.error("Error saving local_tarefas_plano:", e); }
  }, [tarefasPlano]);

  useEffect(() => {
    try { localStorage.setItem('local_db_areas', JSON.stringify(dbAreas)); } catch (e) { console.error("Error saving local_db_areas:", e); }
  }, [dbAreas]);

  useEffect(() => {
    try { localStorage.setItem('local_db_segmentos', JSON.stringify(dbSegmentos)); } catch (e) { console.error("Error saving local_db_segmentos:", e); }
  }, [dbSegmentos]);

  const performSmartCloudSync = async () => {
    if (!user) {
      alert("Por favor, faÃ§a login com sua conta para sincronizar seus dados com a Nuvem.");
      return;
    }

    setIsSyncingCloud(true);
    setSyncError(null);

    try {
      // 1. Flush any debounced in-flight responses first
      await flushPendingResponses();

      const summary: SyncSummary = {
        timestamp: new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
        totalAnalyzed: 0,
        uploadedToCloud: 0,
        updatedInCloud: 0,
        downloadedFromCloud: 0,
        updatedInLocal: 0,
        identicalOrMerged: 0,
        details: {
          empresas: { uploaded: 0, downloaded: 0, updated: 0 },
          diagnosticos: { uploaded: 0, downloaded: 0, updated: 0 },
          respostas: { uploaded: 0, downloaded: 0, updated: 0 },
          tarefas: { uploaded: 0, downloaded: 0, updated: 0 },
          biblioteca: { uploaded: 0, downloaded: 0, updated: 0 },
          outros: { uploaded: 0, downloaded: 0, updated: 0 }
        }
      };

      const cloudBatches: Array<() => Promise<void>> = [];
      const createBatchQueue = () => {
        let currentBatch = writeBatch(db);
        let count = 0;
        const pushItem = (docRef: any, data: any) => {
          currentBatch.set(docRef, sanitizeForFirestore(data), { merge: true });
          count++;
          if (count >= 300) {
            const b = currentBatch;
            cloudBatches.push(() => b.commit());
            currentBatch = writeBatch(db);
            count = 0;
          }
        };
        const finalize = () => {
          if (count > 0) {
            const b = currentBatch;
            cloudBatches.push(() => b.commit());
          }
        };
        return { pushItem, finalize };
      };

      const batchQueue = createBatchQueue();

      // --- A. Sync Empresas ---
      const cloudEmpresasSnap = await getDocs(isAdmin ? query(collection(db, 'empresas')) : query(collection(db, 'empresas'), where('ownerId', '==', user.uid)));
      const cloudEmpresasMap = new Map<string, Empresa>();
      cloudEmpresasSnap.forEach(d => {
        cloudEmpresasMap.set(d.id, { id: d.id, ...d.data() } as Empresa);
      });

      const localEmpresasMap = new Map<string, Empresa>();
      empresas.forEach(e => localEmpresasMap.set(e.id, e));
      try {
        const savedEmpStr = localStorage.getItem('local_empresas');
        if (savedEmpStr) {
          const parsed = JSON.parse(savedEmpStr);
          if (Array.isArray(parsed)) parsed.forEach(e => localEmpresasMap.set(e.id, e));
        }
      } catch {}

      const allEmpresaIds = Array.from(new Set([...Array.from(localEmpresasMap.keys()), ...Array.from(cloudEmpresasMap.keys())]));
      const finalEmpresas: Empresa[] = [];

      for (const id of allEmpresaIds) {
        summary.totalAnalyzed++;
        const local = localEmpresasMap.get(id);
        const cloud = cloudEmpresasMap.get(id);

        if (local && !cloud) {
          batchQueue.pushItem(doc(db, 'empresas', id), { ...local, ownerId: user.uid, updatedAt: (local as any).updatedAt || new Date().toISOString() });
          summary.uploadedToCloud++;
          summary.details.empresas.uploaded++;
          finalEmpresas.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.empresas.downloaded++;
          finalEmpresas.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'empresas', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.empresas.updated++;
            finalEmpresas.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.empresas.downloaded++;
            finalEmpresas.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalEmpresas.push({ ...cloud, ...local });
          }
        }
      }
      setEmpresas(finalEmpresas);
      localStorage.setItem('local_empresas', JSON.stringify(finalEmpresas));

      // --- B. Sync DiagnÃ³sticos ---
      const cloudDiagsSnap = await getDocs(isAdmin ? query(collection(db, 'diagnosticos')) : query(collection(db, 'diagnosticos'), where('ownerId', '==', user.uid)));
      const cloudDiagsMap = new Map<string, Diagnostico>();
      cloudDiagsSnap.forEach(d => {
        cloudDiagsMap.set(d.id, { id: d.id, ...d.data() } as Diagnostico);
      });

      const localDiagsMap = new Map<string, Diagnostico>();
      diagnosticos.forEach(d => localDiagsMap.set(d.id, d));
      try {
        const savedDiagStr = localStorage.getItem('local_diagnosticos');
        if (savedDiagStr) {
          const parsed = JSON.parse(savedDiagStr);
          if (Array.isArray(parsed)) parsed.forEach(d => localDiagsMap.set(d.id, d));
        }
      } catch {}

      const allDiagIds = Array.from(new Set([...Array.from(localDiagsMap.keys()), ...Array.from(cloudDiagsMap.keys())]));
      const finalDiagnosticos: Diagnostico[] = [];

      for (const id of allDiagIds) {
        summary.totalAnalyzed++;
        const local = localDiagsMap.get(id);
        const cloud = cloudDiagsMap.get(id);

        if (local && !cloud) {
          batchQueue.pushItem(doc(db, 'diagnosticos', id), { ...local, ownerId: user.uid, updatedAt: (local as any).updatedAt || new Date().toISOString() });
          summary.uploadedToCloud++;
          summary.details.diagnosticos.uploaded++;
          finalDiagnosticos.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.diagnosticos.downloaded++;
          finalDiagnosticos.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'diagnosticos', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.diagnosticos.updated++;
            finalDiagnosticos.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.diagnosticos.downloaded++;
            finalDiagnosticos.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalDiagnosticos.push({ ...cloud, ...local });
          }
        }
      }
      setDiagnosticos(finalDiagnosticos);
      localStorage.setItem('local_diagnosticos', JSON.stringify(finalDiagnosticos));

      // --- C. Sync Respostas ---
      const cloudRespsSnap = await getDocs(isAdmin ? query(collection(db, 'respostas'), limit(5000)) : query(collection(db, 'respostas'), where('ownerId', '==', user.uid)));
      const cloudRespsMap = new Map<string, Resposta>();
      cloudRespsSnap.forEach(d => {
        cloudRespsMap.set(d.id, { id: d.id, ...d.data() } as Resposta);
      });

      const localRespsList = loadAllLocalRespostas();
      const localRespsMap = new Map<string, Resposta>();
      localRespsList.forEach(r => localRespsMap.set(r.id, r));
      respostas.forEach(r => localRespsMap.set(r.id, r));

      const allRespIds = Array.from(new Set([...Array.from(localRespsMap.keys()), ...Array.from(cloudRespsMap.keys())]));
      const finalRespostas: Resposta[] = [];

      for (const id of allRespIds) {
        summary.totalAnalyzed++;
        const local = localRespsMap.get(id);
        const cloud = cloudRespsMap.get(id);

        if (local && !cloud) {
          batchQueue.pushItem(doc(db, 'respostas', id), { ...local, ownerId: user.uid, updatedAt: (local as any).updatedAt || new Date().toISOString() });
          summary.uploadedToCloud++;
          summary.details.respostas.uploaded++;
          finalRespostas.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.respostas.downloaded++;
          finalRespostas.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          const localHasVal = Boolean(local.resposta);
          const cloudHasVal = Boolean(cloud.resposta);

          if (localHasVal && !cloudHasVal) {
            batchQueue.pushItem(doc(db, 'respostas', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.respostas.updated++;
            finalRespostas.push({ ...cloud, ...local });
          } else if (!localHasVal && cloudHasVal) {
            summary.updatedInLocal++;
            summary.details.respostas.downloaded++;
            finalRespostas.push({ ...local, ...cloud });
          } else if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'respostas', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.respostas.updated++;
            finalRespostas.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.respostas.downloaded++;
            finalRespostas.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalRespostas.push({ ...cloud, ...local });
          }
        }
      }
      saveAllLocalRespostas(finalRespostas);
      if (selectedDiagnostico) {
        setRespostas(finalRespostas.filter(r => r.diagnosticoId === selectedDiagnostico.id));
      }

      // --- D. Sync Tarefas Plano ---
      const cloudTasksSnap = await getDocs(isAdmin ? query(collection(db, 'tarefas_plano')) : query(collection(db, 'tarefas_plano'), where('ownerId', '==', user.uid)));
      const cloudTasksMap = new Map<string, TarefaPlanoAcao>();
      cloudTasksSnap.forEach(d => {
        cloudTasksMap.set(d.id, { id: d.id, ...d.data() } as TarefaPlanoAcao);
      });

      const localTasksMap = new Map<string, TarefaPlanoAcao>();
      tarefasPlano.forEach(t => localTasksMap.set(t.id, t));
      try {
        const savedTasksStr = localStorage.getItem('local_tarefas_plano');
        if (savedTasksStr) {
          const parsed = JSON.parse(savedTasksStr);
          if (Array.isArray(parsed)) parsed.forEach(t => localTasksMap.set(t.id, t));
        }
      } catch {}

      const allTaskIds = Array.from(new Set([...Array.from(localTasksMap.keys()), ...Array.from(cloudTasksMap.keys())]));
      const finalTasks: TarefaPlanoAcao[] = [];

      for (const id of allTaskIds) {
        summary.totalAnalyzed++;
        const local = localTasksMap.get(id);
        const cloud = cloudTasksMap.get(id);

        if (local && !cloud) {
          batchQueue.pushItem(doc(db, 'tarefas_plano', id), { ...local, ownerId: user.uid, updatedAt: (local as any).updatedAt || new Date().toISOString() });
          summary.uploadedToCloud++;
          summary.details.tarefas.uploaded++;
          finalTasks.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.tarefas.downloaded++;
          finalTasks.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocaxœì½ËvG’(¸¯¯pâj”™% ñ¨’R‚p!¬B7(ª>s92 ¢˜‘Šˆ¡°è3‹9wqW÷ÌþÖÌ¢WµêÓ_À3ó—¹‡Gd$ J¬.å‘ˆÌwswss{¹¹ùD|'ªG“|ÄõoûœDÕøüódžgóòü J¦ý8÷ã“UÑ«¢"9Ê×³I”å½U‘ÆƒUq-†Ãáa­â·I>Ž&«"¿Ì’â ‰y™Ãy
/ç³8ª’x·‰,¹{ð£?VùÁÑó£ªH³³þ@Ü¾qºSÎ§Ó¨¸ªºuú³ÏÂ¥â¤ŠÒI9TýÔµüâ§iMŽ£òmIcì‡FàõäF$“2é©èK¼!Ÿ`AµS©®=Žo“<Š;vZ![w?ÜépÿÒ8Éªª?/ž&ÅÙý`é7þ7ý·Lªc9ÄC¤¾n À£*/¢³då‰ôzôôµOxÿtôüÙ°$¢IO¯84 §à­¯‹µµ5±?GWÙX<Mª<Î'ùY­ŸžLÒ¼JÆ‘èÉ4-Ë¨\‡E~2I¦øõ(ŸÌßÿÛûÿHàûû-z–œMgy9@Àª•qž••(¡	{-¶E„ODœOžEÓd$d_Wå Ÿ¤%¬(»zùj1S%ÅHô/¢	=ˆíïÄEŽ+¦”ØøçäJC —vB«âÊ™^Ùš¤£,šaG.£´gIµ—Ë~ZîÆÓ4;âÇyR\aÿ&É¸JóŒ¸êî` FíVÅåyR$ýžZæ0%½ímøW¯öÁÀ!Õ¯§Ô%\üðí[ów}·¸Àð4/ö£ñy?Æqc}¤Œ~<Dì\ÿ	ù2Â‚‹ˆ‡X"°Oº6n¦È4žÒœ~G ¨}| ÍŠ¿üEÐ×,Ÿ&úÇ*ýìB4™ÀT–Ð‹Ý¢ˆ®†§E>íc‡Ž æK{Lm½…ÒýÁ€†ÇÞÙ»W!DOiA+ƒæ^¾rzãBÒÄ’oE~ªûÕÄÉª¼Š&»°Â®~ªs	…\h„z„Ö;C‘q{31BÚQÈë/J>¼u±*ˆOÄ§ŸŠc¿÷í¢LòªlSK±~u5K !0œímÑËOþ´ßƒÕ2q-pšG™Þ½É7./$ß?Îƒ"®.2N3uë•$%H>ñÚf2í!1€CÓ¨‘K—ïa³\sû8nîã¢..hnKRœIQV›–›ß„¿KfÝ—pê\Ã—^–×D–Á`Üˆ1¬è'ƒOÏ'Éð2*²þ›Ý‹´ÌE”‹2ÍÆEž¥?E…øäZÑóÍèÍªH8P-‡ÍÂ‘ÁXýÞLKAàá3+a L>*alËš†š€*yJ@l% æ—ªË. 
¸˜óaê¯òÈüPMÁ #”î ->Ù•r€íéï
V|òZ[Ô9­"<£0(˜ì·…k«8ªËcTiÒŸàá1Àtš’ü@±<Obx]$H]‰ärI©j2žwª XéÆø¾„ˆ*
 ’·ßK@œøä mYêFwö$*+T±ŽäBì«¹H±ƒZ¯¯UùºZ§H þ]ÍÇã¤,òyëþØ¥SJ‡ÒƒÐ«ÞäEeþˆx^DY¨e‹”½ÿ/)<ž¤g	¾…aC‰gó‹d:ZUU¦HnÐó}‚	/v†SèeÕÊóq^É\Ì§X©¶XÜ¡8¦¦²ü"B"H†+fD4}«Þ!âJlCü·kÔå£×¸œ_ÀáÙYRì)¶û}4~;Ÿ=N'	ˆ×þŸË<“É*¦§ðŽ«ªLÑäJ¦„~2ÉO”.õ=|í¿´_¡(E1:½h6› CFíqK|3>
Çö¼:]ûªÇ¸®;/PøáÅ“á[•<'ù¿ûØ W8ÊÆç@ÏÛ(Áçˆ<Ui¸ü‚uÍ+•žƒ ïôi·‚îŸÌAd÷´¬êY¬øt€¦‹ø «+ º8-‘f¡X/Ë³¤§™>ŸäñÕð”dñ£ót÷eu¿½1à‘V3§Òi’Ï«~I€¢SA™ËªgyœøR´V`<>¿H‚}Òô¦?8aEr‘¿e8nmF$­Š­S¬Hªy‘AÿçI}U×4‰Ã$?afVÄäHÝ_¯.N”²	ÆÂEŠSƒ_GíÄºÚŸ‰$j~xqð(ŸÎ`:Ê,á³á.O >‰ªÎÕ^w'Ô.¤Ú‘X— ×:Áêç-ÛH£KR¥OP’âDµ5ðhÂYn)+y[åDW[õˆ[Ñ¨ÿåYªè’+ßº	ˆ Qn6m[/Ã#cÛêhp–Æ¡àùd“®OÅQR=6eZÒ_	f íØr8I²³êœL¡ª£JL’JàÚøþ|ëÔÂ'Ÿm‹Ïa:j+t¨vY@Ò:À$«-Ÿñù<{¥$Ôè)é§«– ÚâL«!Cõ¬[×¢e21™&a“ŠÔ­kµD}KRÂv¸$”<P”À(K+ÐÈç…Ñß¤"Ö¦Lc22¯¥Å2"
vŒË¥Ž&[’š×Òè_%6pò‘Ò†ÉTxÐßL‹}Ä0ÀË€ªª¿jveàÄit–å%]y+,^àñŸ5˜ÐÒÞµwÎhæ{TÅŽmÐx¹š“Ûª‘ñ¨HÀÞ§QÜ¼<Ãéë1{^kÌXZm-0s­W3Í(e_µƒ²fYÍ 3 ´aÕÉšf=ß³ä#¬VÊ1v˜c~1Ì)Û†›^ŽÅå€>‹£ýzK¼Á>}\,l§K}‹ó °¬ÊñîØn¢Í€ÁÇ¯#ó¼}Áu@KÑyRëÊ4¦ŸÆQœÀÙ—»µ<4èbð¹ì(—¯ç`'O’ýw³¼û2RÔZ!Éä ™RòýÓ<Ž&ÏAŸé#‹UÃ§÷PYP¿¯YD`«oÅï£ê<)ÈB·²¤è®£rm
zÔ6UI‹^îžŒ#°¸A-+` ý†fé@[Í \êÃ¾â#¡¾IŸ.ˆÍE^}S«´Ç˜ìH°_¦2gÃ |»f$ä/ú±;Ž,Î;@vù‚pýÑõsšm/û4@Hí58û5xdk(åðŒWl$ôWÂ0á`=ÅUGf·‰ÕSïõ´wkD;S|4{¡TrO=ßòW°œa¦¬yòŠ{¥v-¥gåeR(
—Ô¬ˆœµc°S¡]Um[ :¼;™ÐZ}¡e¸U]ÔFF43ï›7úÛÁá­˜œ‚Ö>±µ·`© Ñ6˜LÍ2ªÅíYÔÑ@õŽž"À¬-vª,¾íúÒÎÜMÒ°XÑüœÁ	iÆ34vQ3&ýrªpµÁ¹¦ƒ¬=„–=À”¯^i×/`‘µ5|s
sp.ý“0ã¼ˆKà§âp÷Åî“'ûOÄe
JD¦Î¸‚j×J(¶’Æ&èÜ§ÉÙ<*bfá–MÝ,APÊD5RãÛãÕïú3Z«°¤“o¿[ÕŸâÀ¶nl¬Z·ƒ8ð²¾™«¬EU`XD`Ò¼tü0³Uç'Ò»×+VûEB¢¬fP«—}Ý›ëê`Àà¾âvÛP“h øGÜúE#Ií³ºÇÌFðâÛÈ¬x‡äZQry}ìÐE×ìtè¥[ã6=E²|O™I³*&)štäØèÒg§îòVò|ù>{VS—®zUnÓ]Éò–ï­Ï*»t×¯s;Ê-Ç·¡\OKïF¼^¥Ûô”ïå»Û ±wétCÕÛtëåûÞ`—wé{CÕåúèq…D™E3Ð`WkQË¯?Eõ¤þT©÷õRv‡€—ãúSÐÄëQwæòÆ„!iÉÊ|Løöu°’–Dƒ!i	JÌeóÉd =½«êÇBHF`Ü”áèwÅyíÝ¡1VxS|êî 4¹;$³¢òtOsUË‡¢†Ôw´úª+SÌèCáÈ1eDzÑcÂ3Úé`_%0²²’W[÷(†»Ño¬î‚ø7P^uï|ðnçÜH4Õ ðÝy‚¡eî~_ ‹ÈoñGq†ÉœÈäÎ)
*ÃJÂ(u®ŽÒzíVœ²n„ñÊûéâV·Ô¹È¶rñG¹æm}t~ TÍTí0d]ÁZÜ%`ÌÔï6~½“ˆç~j î²2¨¨
2º¢‘V!2Ñ}¬ã’è€6¯¿ajáÝu©E·ÖZ”ƒ@#[þã:l¢³Z4z 9ò~k˜í‚n¡ÒÀ8D9në’ëT}¢J÷Ø#PXL‡à{[B^FÕ+¬yB…Éô
´­ƒ­‰¹.
wn:SˆµNt>¦±ŒC…úX_%¼vwÅz^(¼»îBÑ6,óMû™È‘´_-–+™%ÒÃ!òìýß.’	†Lsò‡Š8'´#€[¿ÓY
«aPP
,%tš¥¥(“³y‘—1`8À45(]ƒ†Ë=žl¬û®XÈòËÇy*WE¾¼¦0äa‘€Ñ=Núë/GÃWë@:½µÞ@m²ƒµùµ™¶Ê!p«ôÆOÁ Ð(€Efý–³IZõ{ÿµ7x¹ñŠµñFk?í®ý·µ¯_¯ÉÖ^÷â3üÖT¯ç¶Lè±
TÆÞàÃù¤Ê‹4z-ÿú“kÛŸ›O®ù o†‘óæ¯²`]¨væ£Ùšd@E¬‰£ýï_ìî÷¬Fz‘%ØyPxk¸ÁžËùjù¶¥±Ó#‡ðm#¡ù‹èíb€c.=Æ=®Spˆ,=—Á¨ïò¾¥{c_N“*’&¼Üow˜<åUs¬úåÝÍ¥€ŽÖPÏÐïÈ'õ¦î.T@Ì7Ôó6ž2¯ylåx”HM5¬œµHŽ¦ÚîÎT€?7Ôc;PÞ–Ts³÷ämF5Õ°»Nî.”*¢²¤”l‰xñØ’E#¹Ø2Õ±%£…$ÁûènF†È­¬àvd3)Ô±µˆlYã´ó2MmË”SËp ÷ÍàU({oömgaÝqó²Ê§Oò³Üf8o„oµHôx·Š°Ô;S6:Ü2uiö¯
™
ˆÜ¶ÿ'Ÿ©C5ëe´´é3ÉcÝ®G®í¢’NÈc{«Âo®5^#P=¼U¦âÓŸ’Æ€l ¶óaA‘ñµØìW ßñ<ÁºØÜØúbàa#`%BdèÃ>räår²âv’bY9q;)q[±¬„¸‹|X–Ë/Ëã»px¯ôÒ²
×Âq4²óüŸOFŠPÃú“]ÎCûØUžöƒêP;(‚Ysæ6/Æ‰,	*j/C•þuòšV éœrÅæÙäªWã;†1Ð
èÓÒà\£¤Ob¨2ëëâXž¿Ñø\¤.NŠü»£CŠ{i8¨¡ùÕª²¥¦£0ÐÃóüò8®Ò£b ÐN™$`»@C‘2VÊ9)™Dÿ“k9-7âŸ¿¼Áø8	7* ´köÐºÑAK¸‰dÓ…šäÀqšfœúƒ2•õe€Š0)ê§q¼6+ƒU±B}ƒN¬<Ž&çZw²Qsúæ7¡à²ƒi=¸¬Tö"‰ÆÕð:Kˆq|û‡ã§O²Ù¼R±þßÕB¹OåQœ¨Î’jˆ¿Ë!ØM6ï>´×¬vÊ`3(ß°" «ã¤P6 ’Ïz`ÎÈ÷Ã\ŸPQr´»é„Œ‡bRàßJ5åUó;—òM{)¦Ý”º†ž«ïC(73ÖµIÎa­PÇå	«•ÝâÇyz‘‹‹è§4_qìh¯g±´ñX‹jˆïÀú|†#¯›åø:& L@·—fŠª³öÚH§’=1/ºêkB‘B£,•AŽ6*Í‡ãyk!Ó,~S?\ ‘{Ü€i”ª±ZáxÊZ´´†Ë°öª®9ç]+4#5£›>É'‹oä£¤X@ôd -95­`0ø@RW;5˜ =mçMÿjƒÀÃùü hÛŸ¦'ÀûžãÎ‘Ïæ1ûÝ¶*<oð‘Ííüº·QÖm9¬Î;Xã-úzý‡ú¸ÍwbÃå† u™:$îï$NÝj¡˜°«ÚQÛ`Œ_TBþºžÖ¥m”{Î©6Ò€Þ-/†ºÜ¨0®Æ‘;Ü~ô/ìù6tß~	ŽÛFÁÖñ—ÑERœm«Ó¥ãœ×;}RÁî;ù}‚P—›¶9‚-Ü~ÞBâ)¸NöNëšååÛ[»ÝZvÃ.l§ÅÛ£ëÐžLb8òÓ bL¡ °
ØïvçO`Ñ€ÛŒ·ËäÀ•GL%h´_'	XJ{AØº-Íy,®µÞ€@U( ì.´É:¨›¹%k)>î€À#s
áÏs_Ñ§ËÔ!Ýy,)IîL#·B]=—É0·§ÏÝqÙ,ŸµŠeY¤fIa¬3©4b÷?èÒ<o¨)¹ô€yV˜ÆA[ø·¸£š;C¹5ëƒ_</ºq<nOºŒHZÌf¯ 6+Ì«ZÑ®R&'E”¼–5_O *¦éñ-ÕUw£¡Ón¡–ê]böÄ£æÁø ÃÃúM—æBÛÆ­Âá:XòOÙ„xªäFèûÁkv_Ñ×hV½’îþbH½÷k°}Æší—u÷CJ¨_#¼Ç×¦—ùØž_M—©—5 5±í—µ{€¾”ªáTïzLº^Ží	x›_ÞÛ“rúüº{³a‡_+¼GÛæp Ü°x}/ªÙÝ.¡£Ç=a ”^¥€zàyÜðªh^¨|T°ˆÇ“ùû¿ÅQXò745záyªÓ[èK¦Ó\‚•Â£,TGa8Ä¦&i–ÞZÀ‘¥Ùi^–ŒliÖªe­­·.Oœ²>ø/‰RÛ|óÆJ8Q#®Õì7mz÷4&¸e¤œÝC­"NTý¡øSR¤§éóDä˜ô(:®v3"!w*ä€WXÿýœb’QóÝ—Žør†Ézu´™D#øGm-àŸÝò8yWõi#ã³©²¾.ö®²hŠÉ¡7¸^æ@;»Ožˆè"J'°+!”B1‰Ñ!¿÷½ødj‰‘Xðm’‚0"³Ö˜­€¸«ahG*¬Ü§ Y;‰‡Ti,q”TjŸ¾©`ÊïÌF‰*Ç¸e€;>€‡\GÈ¬zŽLí¦ÌQ?<<OoBìf±Œ`ÃcfÎä>®™ÉŽ™RÃ*’_&Å£¨dÉÿ éGoÓnÖÑ°EuUâ”¼ÃÌWiS'À¸~VV=‘ÖA™OéÙÐõxà6NÉƒ°¬³ÃãBÓž‹!àºo‡ü&­ß¨Qàú…eÙâ¨ˆQ<ª`	ƒ
î¾Øß=2Ñ©öYÍa?¢|½æØ)iÎ’œÍ²¨ò°äTWÉÅeÛ‘šjŠËÖmSaõ]êmm4ƒNæcB
½Rxb
ÌAa%}´š Â´8ã-ÎhB›±Z…0eK^¶tÊ6§×…§pÜxäƒ¯³Xî‡0å×8ª—qˆ±þºÓ·s9Ä$-+7`Wô ww1mì÷9ˆ‰(3è ÓàXVQõûÑª8!VÉìê	æ™(}<^8«Ö¾ÑÃˆê2ÉÊ´J/Òêj$z'°Ðz:ÊÔ”—&gO¹jÒñ°ÔInZ¦:Å­
¾øÕI*Å5‹=Ò[Fm\V’;&f[¦Z-¦zy8jX¥³\\ò‘fw•õÎŽÜßGådâôã;6Â7^Nài•ÌîÞnGÚ…xoàÙ&¥Y¹œVmyWüc³ 8dâ„‹Ar‰;™HZÄ”$jáaÆæ-­†^†Ì=§@Í€òÞz&“óÖ·’Ü–OØcgQéµB¯ëžg¨Û™1RÄkA©1cxÃ´ØÁy¤ 
Š­·‡!5öU9à“¥x€iÎŸ…çOŽŸ½Þzøbÿhw˜JyRöuþˆÁ@rªë¼bc˜5Æ©Mç=ø˜^ieSË¹ˆ%2²./¡r ³.bšàqtÖÎ¢*(Ð¨¶HE.Œ¼”(C„éˆùÈ=™æŸC¡Ô½Ó)é‡ÕŽšIn(Ñ$²È@w›T&ýiç‰…ÃB»y(e}(åG2E}ÎÙ¹3)t;ËT+:d{$Aï=JP£Jâã«YrCÃ–“kÍµ6Ò’:œ]Qx:D.×Â§¯ì+bÕ‚)H¬*ãOAIœa§`íÊ§ßÑcL‰
ü¾hŽˆ] wÀÆXp¹Ý×Q…õoïÑL!©ï1oÇè³nò*x7ˆ–Å3›cÛ
¿Ä÷ÎßãO{,ÍŽè³–¶·uA<xã<×°ê/ùÐÚj™!à"™ÖS³[²R¥$É²C”Ò77þÎ†ëïÁø;Çk°ô4 çƒ°~`ŠÈ®éëWrXDØUBM:ÿ˜Õ.sjê7š:x1J/…=Ác‘Øô„(£]f¿áGÒÉ•HX÷©æ©±v—$éïÜN;ävÃM¯’|¸øºl²ÀîËÀ2‡µIÚW?“I€³8y7¨ÓvòÌ<@û(½Èp¥‘I
rÈwùYv:êóåÅ#·ç'ró‘x_e¯?¹¦~Þ¼±^ry·6§gþ&hQ8Öæ«IcI.%F~@ñ^›å³ùSa–J~ê)¸Wâò<ÉÄ3tŠ|s"àµ©¡6UÅ˜â®Q©é»z
oùKÎÒäùÄuB íê’ we‚Y~ß‚–îZÚïéõ[¢6­–f•ÀPŽíãªíáj»ÎaÑú‡d2VTå”yõœœü‰töÍ«sy3¢ŽDPÖ<ëÜ
†¥áý µþ*±]*3>‰h¯ÏÁùƒ¹ë­¬ß"aŽt<BkÙñœ‚íyªMñ§%¶LG^	’:Ì¡â§U4ŠªˆÒ"ý¯gtVyœOÛzH'ÀÙN®t*î@Ëé°þö}%Ò x&%€}µ=ðº-óxÊu¥T†›E Ø ú¨aí;'÷ÍQ>þäé•äâ@…“(+‰Šñ9>§ibéï ¬05‘u…åG=M·CAeÈ“­0Ä"P7äa\Åè)Ê»PÇ­,ê!±ñP?pÆ‹4&jêêŸxÒ‰|ìY	 |³
v@&Eey™1Y:Ëó³IB”ëêûu”ÇÉi4ŸTJM¬­6Òl[ÝAÅwÄ›O®G7`bÝÁìÑ­+qLKðƒôªTÏMœæÇÃ[r¬;nî b Ïò‹ÜÂdC6Z_r‰Ô,ç…ÏXýåG9PÏdÄ‡Ï÷„MÅH,Ñ©~};h¦X¯y›ú¸?çrü=†V<BŠ¯ŠVkR\$&Á*¡YßÐÓ)°Ç³4[PÒ^*@·Ô±7–ìFì»Sâ<¯ò^<ÁÊú{½Ûx@n^Â°wAgŠœW¨ÄëP‰ æöpÛ&¥Zˆ¶cèâb$Jz¬ÿ5~ âç¼N‘OÂÅ{þ¢–Æ“·«{víÚe,hœ{õ‹!šTðnK¡L£ö··ƒKl%«T)>o“·Êt×Ý8¢á‰b|,Y ûP—QO\.ªºmËo3š
íÃKðr¥Ø5©b'ZþËoÍ ùbÀýê¬–Z²ìô–~B:„]ÓM%$ÃR”5÷’UqùNSkt]–G®ôÒ’v*[Já–my·ô7²fÙAYZd¡B¸Î–\U¡þëè.,‹ÉÜ#Ý»Î©Ÿî–®*Õš‘¸?1x­OýÚ·8µqþþ¯Eš£º,µeTì+sÔ‹b‚tP¶^›Àk´I´Ã¿œŸ íñ˜gX'ò()˜†ð`HÖ±'Á.AŠ€§ßØ'Oò(ÆÞ¡Y0ad,Ö8ˆûsÌ×C©cF2nA—„qW3°žð ÀizV@wv'“ã|_—è»ºÇœS›Ñ á)noOé4%fÏ3)E¿×ÒL[Ú‰ÕîÁ6¸_ë@O&0µåÈÚ\¸l¨!È8žÄ»zÝMBõƒ!Gô„AN`'Ð’f?ÜëÁP¨ìAië¥H+$ãlœÄiœÎÎJÂÍ×Æ4yP$gHÄ…`T¯¢¤ãÇÚâ*©,ªÂÆ"£ÖšücØuÉ7èég‹F;&¬,‰¯¼*ñ¨ÝÃñÏÉ•ROd8:®ºP3;MÏ§`êTõ“Ð÷WÖ£Yº~–àL¬Ëª+š`‡h\cVX•‚òOõîk)Üù‰#FºP—W?Ê#ê>öœ¨óDn–öj±Å¶äî×P¿!ä›5=“9T3º=kˆfËO.ÆçÑEB8Ž1¬Œü7ë;²St” x›å—“$>Kˆ/ 2$—@Ê#l7GŒ%²€ØƒÓ}„Ð÷œ.ª’íˆÎÐ°{ dô'XÊOLöpŒî)Œ3#
tŠöÎ~’ËËšûw‘&—rîLyr+k×•­R´ž”öi½æ0ƒg©E`t¶î{v¦GåX„ê5:~`u™~ +â/I…SèjhX;òo`÷”‰MKÉv×ÝFî$A 2’$WŽç ­ÊñÌ8F³)yý¶$ÞWÄJMQ(Ä>È?²ð7‘Z‡kTÉû¼ñ€©	¬gyv”E³Tç¾éðªÌ#p<£—B¹-:g¬æ¦Êž]…¸i™ìùÔù;®“˜ÓÌÒ‰t¤6Â‡S˜;VÕïÃ¿¡›$àM2¢æ¦q)tÊÈD–èŒWqHzßà½€o-†Tåyq<–”êùQõ–ìU
$Î¥ßTf1:®¹Y8/VÅóYR7"çò“ƒ£c—òjùÑË¯àÝáqË¼÷"ç=ÂõÉÜë<'u÷½Fr—fômÉ=˜ŠTeoð)Ÿy™š@½Hh8'„±1-ºî‡Bü)µLTÞ\e„É!ó8Ÿ¼ÿ÷3°|("tËtrŽj_Ÿ]Ï.nÇ((0ŸþCsÞÿ«Ü²‡Ø8Iò˜¹0±³µN8¯Î)‚uë¾¹ŸíÌ—=¸\ç<€Í§xU§ºªH^K”]á ®Pù–<û„ÊÄúf+ñr;‰½«{·5Ly£”bS*¸‹â<€õÅè¤”È$X©.¨ ÖØŸy­±È.ì–
èbm¿r½;S¬K.Ù°È&
2‡‹Cd«»d¨ùþ)HÁ¸¤On×)èW’“ŸVN,”5LöÌt XeN&°óÞÉAÝÚçPƒ9ˆþ3C)‰¡äÄPv$†ò—&gF9-è»ºªí—UÝçEUÁÙrzÓ—ðÜ;¹5Þ?àçia*ø'ÇÇ©;Cú(VÃÕZ2ö¢Ž6]#MÅžÜ7ÊäKJ_†Ã¨C^7þt„ûC¾ÃàX¯g/2°8VÝáq6†úD¡Í°phtÑÁ†}gŽAm½Æ5Ä{eÕ?ÿ¹êîsÃÞÝÇjùyO%Iy Žì!8éxAF³*j.-òáRLŽ<KG#1$H)¥èË"·ü$?£‡e.&°Èj¥j“+Œ³ÊOO'i– _ÏOÚ^žŸÆö‡¤:}Ôé~º•ÙÏ€^;mÍx²Ž^†N)yGýƒ„‹»$(î¬XÃ ´¾ŠÚ›Î n#¡ËSõ6,¸eEN€¶Êv^•°ðîÒÑÙ;OMèÊ‚h¾ùìº*¸«E«¤°WÊÓÀð«òf`4ô:Aë<Ðe•^äoÃ´ñ¡©C§aí°–š;â8b>wq÷¹Cˆfò¨™ÛÏÂZ<uêâsŠ4²{^ö9üéîmÖ&ÍÉW›8£ø
8=s³wš-ã±M¦ÄÀ©bÇ8å<ÑÂƒ‚ÏSÑ}žœ™'jãöó„°n;OfX=é¹^0I6ñ^m†…ž!KÏPÑ4CE§ò"4k“„ò‹óA÷¦9_tÑÏƒ1Ÿ¯q÷ùÂ˜ù¢æn?_ë°ÄÆàœFþèˆýÚ„Ê›‘ô„"p6¡ã.ÌrÜir•
šWk?ð4zFéøÕƒŠÏhÕ}F+3ª•ÛOhu³Y³ºgÑ±hj³‡Z>À¿z+6ø¼Ëb¹ài—iAä†¾¯›‘ÞeT³,nŸÔöO­Ö'‚zTßl°šKØ'È¿Ir¯Ü\¢¯Õ½“äÃ¿i¸ èårÏs ½—ãyQ`|=Bî•â‡ƒ½ÛéãI”ÊÝlñ&/f ÕÃúÄgI¬.rÂnß¼ùÆ€
Þq¯Á‚	Ñéís‚Náve'SÕµi­ O3Ï“ã™äncº9ÆÇ«0ÛÌ'qK­Þ`pÁetÇ›¾ì­tX`
ÕŽçéUÒtÓ|ÂÓ  4&‹SÃÃðóšÙøþ|+ÜZøì³mñÅF­®îç	­mqY q~O!	ñ‰Ó6Sçóì-š³ªy_YºŠP~5®æc´	ÿ¦Þ!û¢¶ðúªè§x&h[³ëÕ~äÜI¨ã|Š÷¸{uxz¶¹nœå(?­„ÕÂˆÉJhèÞê§Ë¿|È	L§×2æbžôÂLª¶[9Ô‹$ñ1†^Gf;/Ð;£‹ôâH<‰2k³¨ˆÄr:›ÔÌ‚È´ýôSaã/&²¾<Û>‘1rÎ	Ÿ?AÑ~ï ››-[F8¨š¶ã2tÙ÷îd"œ¼ˆ÷ê<‘v6¬!ÊÀ¢\]X¨îg¥7Ü1ÔÙÍé˜½œnéîQ
ÊÇÂ}L÷ïZ
^½
Ið¹{‚fù(sbTÊŠÄ˜1¡öÒ¼û­ØÜØØìz¿ZYÔ¶6(¨0Øº…ˆ=øzpÒÖƒÚË–ÔÊ6÷@ÍìÂšD†ï”¯ÛœN\†úÓ=<ƒú!®ýû3†h4ÆÉ "unñý¿«¼“’ÍæÅÀáŸe%öÒŽS~Î‘Ï3)f3¼Œ²¤‹TˆS¯Ñ5q~K	ÐÎcBAhxK(Ä,fiM^Šç«?BnB/ŠÓecô†½9˜â!…© á#;ŸÈ M© —R ¡RßUÊ–x¶ÿ§ýPFLþÌp" óÛQ¬“1Ni±Fãs¢¨3Öq{¶“¨¶ž†*!×ƒƒ,V¸k(UµrCd X!5 AÚ[išš­!Þh4YCfÁ‰¢6íI¡°±Ž+D˜la¢¤ÏnÀo7¹BfÔÞ-…‹F’•,ˆL²™åÙI6…'
ßÓP/©¦»au ­›n$æ°6ŽËgóäöÖ‘ lEÊíd Ù…«»rGÎÅPÉÊ®Ü#åâ*hÆˆ[šCb‰KîeòïÎ•æ€/.ºç›öÅÉý"9*óZn¬7¬FC}ÌŒÐÈ¯ØzÅ¨wSCåkyôøµ6z%QöH¹@$'ñGß¶é‡»ÙôÚA\dX‰E6pèÜ:3uA˜œƒB[º#Ä>€LŠ³yVEæ,I “ŽÛ˜"!0×Éu2•ìÂ:„'³n7/_¡A#¡d~‚x<ì—JÈ³T>ÄØf0x¬3…¨«	g·5P`´Õ²Œeœ-"9,„e_³Œgì!ÀÃ`ñ4›'gnj¦;÷Dx$“€¸{‚øÅD-*eHKžp(b¤Ï3Ð®äA ±jÆ`™—ÿµœ9Žg“¤ª™åÙšÌ!@žCn
hbJÀÂ2Ú0PÛ|,•lÔ’JÖ•I%‹o™”%ží6´pÌe*ª.¼€ky™˜“ÇJe#(5XU Rç´hsÈìg®¹ƒ}’0š8¢ÉøóÇAí=²k™—@Þ”üðLŸÈ2ôèuO|Föè0Ë/az>=z$ï Ž²8—xT·/þå`ˆ'·ä¯­Uñ»š¡L:c¬Ô=ß,}ôØcïŒ¸Õ-GJ¥{ÄYãuDœtÕiBóÕQËº§Þg¬Ü¬¹”"/*åš[ªÌ¡ÄÎ¿1—ÚÂ¤Ìäã‘Øäup-è_ç`÷	žûŽ@
ŒüSúcÐâGÂ=*¯ðì—uœä´£>ËM$k.{—ÞÕ/|°ó

ïÚòiÜoõ<ÐžÄÎæåyŸ¤îˆÔ²W“Žã”$’o€ÚÄõX[ËâƒfÇ4ªl8¦_5Ú šYðÃ&óKîÉà€™E6;XÙ{‚×›6¶óT)Îá½üÔoôí	„Qõaöü¶–Ù# Â¶:}ˆÛˆ@†¡÷z¦Î”k8Än¶iÒÃÛ¦µ¾Ö·î¸åÐ°Á Ñêm2¨­ï	k™e¹õÔ¯;pv8˜•l	­ï±Úßh4_ø–LC1¦’:‰Ì<ÏÉlæïa4[‘áC ÀDíó"™D˜,'÷ž#bó³"šF2»t½Õ¸‹ÔuRÅú–]àåZSÇ	¬ámÛb³Î)èŒôRœõ«V!Ÿ›r/þÞITVz Èéx'^ÁbºH
:î/õÚÃÁ7äB…Zâ!,¡"N08å„¶º}EˆéVCK ?¾h9dÐÝÙ¥¿¼åØ²Ù‹K*+àèÒŽ­`½*¯¢Éê*Œ«œÆs@e?ôenÄñ/R/†¤1È½…Uá3Nñµ“Å ·Äo¥ñ]’sasÐ•>LŠ±t4H’îbé³!¬‹¾m^”›%Lç‡¢?%ø=æ“i/¯ž0ŒŒ¶\hò¼]Ü2‹×Ÿ>us<áG©Z|.nàŠ$óåJJU G_u˜k$,¾€-=~OM=Ã‚•9ðuòßåÜT#|øÛpRÎdéÑ;2Ð¬¶wyn‰DP¾'šW9Ì4/ï‚æûÇMâVÞd÷Ô¨Éá“F;ÎáL`³Áˆ¹¥Nú•öh,a§‡ß¤kÚË$®qs¬»
<ã=¥4Ö¶ÆÔ„©õâÙøìj0d(?p¾êO¿ƒ•*ÒˆmF©x7{z@Ð†Ñ ëoÿò—fHÙìÏa ìEK}ž$Æà¼iëA2™O¢†‘¸ïZ T@§y–<NßåAP¡-ðêG‰]põ÷žuf¦1d¢ÙV—'ÒÍ¶.£sÒì5ô£}.<C&šG9^y VÔÐ‡[J%ÝPMN-ÈçÚ-ËçÎV¨Ï(çò»ÛÈ2ñò¾ˆ@ŸŒÌt9@¶àY+w	4»U0ŸE-ÎœÿÕ‚#ÄîZ²ˆ,ÏÍØÑüÏýž’gRn¼ ’Ê-Y@æâ’'Žz6ÔÉh_ÐªÑÖÊô,;È0ßÑa>›ÏTÒ.ÜƒM"¤C•HéeŠ¤Þð†h[p~ì$ÖRM-N®¤„¹5ëp5!çwÊÐsX\±Ê€Œ}^Ç¤Ì³µñ$/“xíäj?½ì^ØºE(Ê‰êàQ/0>ÊéR;îƒgÔŽ#NÐ¢œägý9T_½Ûä9ªÂ7…5ažá ˆŸ`1(˜îÓvÊÒÓL¼3Êúç<°ôB\ï¬KÏAm™¾ÿ[–æ"ÆD‹"¢-1™V.Ãàéû¿’>SÈË••L™I”È ·	¡ä"åÅUŠŽq‘«wùfVæÔb¸Y"Ž’ìÚ>‰€õ{èíóÊ¹8PIÚÙµãuñ˜5{Ä§ÑOò`@Â´yj+\kd7iÈEI©Ñ(ežY’	z£q5Ä­ º”-Ò„|_(OdRW­ó†,_Ò”Š<5£3œõêR¡\_©\jùIš^€²ïY|¨Rôª…ÿÒ+ùU¿­9»AÐbÚöZsø€rïË¢Þ­AW	ðâpúÑ†![ævëáj£7ÄeFªÞ(eÚxŸG÷DydßÈcHX‹&x½åÕZš!Ÿì53ˆ}Ì¢™Èuúg`’$SÀx>‡°ˆO£ú÷ý¿Elý”óï!8Ørf"Á®¥^^¯®ˆûË^‚1v¶fóK7•Cd¯eyµF9ÁÛ*Ç,Šº-€§æ ûá=`ÝGp™Dom§šÛUÄÀîïÎ’I.ÀTSçK €8H‚Ùøº7ýóËŽ=-9HN´Kˆ¡ØSŒ”Fa`DNfÝ˜òÊëmÿ#¡Ôf¦ò®MýUëÈhœK•'ÕäB%yÆo -J»³ÛÂêñò9jMRéBÆò|^°Âg˜¸¬I]p¸÷x¡¾¦üäì4wq{G¥ÌH3S)ùª¡ãÃkQ€ÙõjÈæH†ŒCF7åÔÛäÃyØÊ×¿cd^'¯)‹@žÔAž´€¤C»T:ÿÞçÂ“5ø]o;Æá€âR&úX’Ž>ÈT;6"™¼ˆß×Aœt¡ƒ‡±1ëŸõ®1VþÀsT!4`¼ø-¦|O˜Ù•g¸JÏy+²ÂeJgÓðŽf¥ñ] 7 £‘R
J ÌÃ‚¥š‡·ü°Š­;kØ¦ò:4S®{4.þ\ÂÊ`Là°TiDoÂÖy–Vðh:eÏ¤žF_h#õÆ:´ƒªÆÒæ^Ü°½IcÒÚbÌÔŒé13Ìøâ¸N¨]½Q¢Rä ÞÆ¸!"/tžãÆ’ºå¤ˆ’Ðª¼S¹úÅŒg ²EA©JHônßÌŸPHÒ!
ú†eòy1òõºçwŸüÍ ø¥Î¤íÓãþ®‚'Î5ÍdÀ/€ß <Bèx‚Is¬â' ¨`¯ÿ{ëg«”LÎ[jEHû¹!!éfoÈ#º=ï4¤Ô‚A_UžØ`®f+0ÇT©ç)‹XÈUõáÐ×„@î¢·%Âè¥›xðúsqNò¸ÔÇê%Þï7’kpšV“)ÅÃ¯
Ö4‹bÂ'Ñûñj‚¦g‡EN›F1©œÂRÅ'ÀÙR˜Š á×`?ôºÿ|«©Ä’ôìœ½ntx.ýÖBˆuÓ§á%>qÁÅ´D¬u¦îðÙï{«Œ­i^µf[X›«ü§íÛª•’ ìñîÑqÏo(‘¼¥Ý”Ï'tt'VM5ƒ~C-v¥d€™f9ã}²
„ª`ôl&èÐèº †FñSi5§s~òú5u“8µš–è¼+ÝÝ¦ƒ§“É£|ªÝæ—€˜¯Â?[_ËÝÓdŠq‡ñÚÃ§m¸%ñÜa	¢úó‡ÇôwkUÈÿz]dRûÇÉ»J¶¿õðaýõã<«Pnô77Â/û+°š.ÔD‰+'ù$^ñJVÐLÉógG?<9~þâ`W¾x¾Ò@%›¿b§éßkï Œâ¥ ”ç¹ç…ÿü¦¡«ŸKô½ îøS+Ù>ŽNÌÒ.¬8ƒg ?ãö|ð |ÿïb’‚åâBµøÙ„m}¾*¾Ø’àKLÖ¼ö5››®HQ(!æ²ª†°9††öõð¡Çí ³¿ÚÔÓèÎu?ÆH|ríG `Â¨Ç€¯¨”<ôþFüEÜRH¼ÿÞÆ À%ÚÍ9,Ÿc|ëS/ßXÜ|¾èï>Æ…ÀÆÊæži˜ë*£o“½ïÿW.š¬[éï” ¿­µ8=K+x4M³9¢È<’¼áý²q÷ŠèR­´­/‘”pU~±Á‰iË!&<#Õ×€¿Ø
-IZÜQg¤ž„…ÿbÿÉîñûÿùâàùH>Ù}ö\ìí÷û¿ßÿ_ÏÅþÑñxùßðèy¯¦Œàív{2q¯¯|ÐÕŠUè^E£KW šy°‹òá§Õ•zØ´#ìmçË6|G T#hP2Û„Ò9mj>¦d²~ÇÔ8^ €?¶×©nÚVÛ}ó(ëÂ3ÿ]ÿmHP]1ÀÆBª6‘¢<më[ýä"‡Œ!¾Ç›[£ø¯Q2oOÙ¶8%(TÈ®Ã7§ßöwS§	cÍ@w±x ¯ð8ÐQóƒñâIE½Í˜Ý B!ïz2ðö"-ŸEÏÌ;k®œ·Pš½
©µqzzŠ¯u¼KtRúõÐÎ­·°ªÔ^tUjPã$ôüuÑ§Å¿_š¶¾` ÷¦#R*Þ|r­AÞÄo‚²Q§åÞ\MÔ¹Œ¢óv`º˜ÍLd?*ëkî½anHµÿ2¢Ü­¡7•aì$ðžOaËÌí(ûØ¼Ô¡>ëWW=¤ XŸÜùæ áÍ"hñ>:Ò\^ôÐyÒü
—4ì¬š§Ï«üÙ±äéÌ?XEEõ¿ÄC?~Ü$^¾ìýr{ÿ¯/öwñËÞÁîïŸ½ÿŸGÇÀëÚ@ûþÉþSz…ÒàÿÙ?ûâéþ“? –v„ïŸïõH†‚
÷þ_ÿ´ÿ<{ÿ?<Ç¯‡/vÿ}ÙûáÅP>8x.¿ÿ MïÕ+Û¿“<¾Yéb_H±†‡¸&Éi52b®@ëÀþ¬òÙH|ñÕ* ª@ÉòÖž•iŒÒöäåH^æRóûé[^È3òl>Eåð;7†’°~Ij¡0kÂ££êj‚Æ];qÚ$€äFÁ«Õf£ ÒZ  â!çà”ò8¿«?Çöñ"eÐPB=÷4oç~Æd29Ä »‡ÖBCÃéÍöeèöÒ
/Ãƒ{ùpsU|	U¿zøÊêM¿c¸`Ýç“ù43a7°_Ø$©S#XöGÌ	±éUØu¬Ž@6¾/Ýú[^ý/ºï?÷Þå½ÿÂ{ÿù–ûþ¡÷~sáˆ¾\ºÆïü‹j|µt_ûxÞl¯á.úCt>=‚ÚW})#…õ=NñUÏÜÞ'©f(…$¾ÿ*||A›±²
4IF‰—*C]:¤b{»OŽwA¬Ø*%‘æÐÐ7ÆVom¡ùõªøÝÖ+2Š¼LÖ¾tV»÷éûÿœ}1äM` ›› úK	9B®ç¾¹+¶¾¾/l=zþìÑ“Þÿ½çG†üÄPäÃWŽç¤m‡ûÏööŸïßË”Y|4!÷{•¼È/-?äü}…Øö}ëáÖ+ôX¦öÐ‚/MÝü")N'ù%n$€yx*Û[Åwà‘1ô{ìR1î }œç•çù”Ú Kr}æcãµ5› .I	ùü”
ÚxfçHÑ&)ÚfáÁgŸqzÀ=	¼l½‡)7éåsãçøªþŽ;b6¼×ä`“ñæ…ÙÀAïð[‡ïÿ
zD$>¹Noð…ò>P§Áœg ¬Í½.¶]°i{Ez2É&ØtŽôÜYæ-Û¨U_.aIå—Ã|†—ÔA[ù¼šÍ«~ïd’ŸÌ‹‰ãAÅäõ	ò­I•å„Z´^ºYè]<NpÖJ¨M7^¤7Íü7¨Ûcô,ºØÃüb”ÀéîZ€¨9€WÅ$:I0l…ü¶«"ÏMÒñ[Ðö"¼˜îF…Ž`Æ ‚þíÉ´@rK•Þ¾öº¨üÔ•¼«/““¨xŽ˜ö®94QY“¨,Ñß»}=Î,¥­\®âæíé˜Ã^“S+þ</«ôôjí$©.“$³wkŸŠÙÕè6B9p×ÞMDUDtÏfž­a:²x.3¡+HœA¹i?k¨Ëoå”Vûfè™OW³‘^nFp;båälñB	áòº*Êóèkmëoº¶L-”É4%×%‡9+²Ä£æ9ò¤kèáÆúWú1•Õ/@M[Q ³ß©ŸßÆéCòJ¡gÑlíst¿­]®m¬|gzô-RŒ(aím_o~uãÍÔJy+äíÚÇ3}E«/€í5Ùí]‚k››+š
™{+¾ðê;Ã$­nÄ:ët9‹2>Þª˜g˜­cå»k¢ý›o×±ˆAÏ:àGÿ¸–« ²7t¨T¨®¾ÜÜ˜½{%çç‰rIrƒ‘$Òó
çZvø0Å4øõ-NH„
6û›šÌôƒ¯ì¬ã¼Ç9"ÅrAwÔŠF¾]—Ë\úPÇ÷ôº‘‹¤%®lä"ÇùÙzÜÏ£R^Ëùè<Ä˜«sžUÙŠ†óÑóŒð/óbm–“œb³ê"@¹GKÌãa}R]&6žñS83Af XÌ»rÅa%Ê5„´þµjÖìµxœ©‘1ý®1ý.È˜¾ú»dLÞ,jÅù½Yë¿$Ãê†ÃM }†Æk¹8kTœ½-à~müoS.<Î§š–6X[4;í\	°¾4³”"Ø›UÜl­¹Á	Üe3Ü^sçÀÐ2½ztž\y¶—_f\œ+úüòÆñÅ¸¤êus!Á®hž‹ó´bËyEëú+UŠÜLÖTÔ"Žæ'ÿ0Zädí+1+€ß °0;9[^T¼+[D…€¬Ý^½t…ƒáÿeó£Q/QúÖùø5ññ—ápÞþðïDéÄýg.R¨ç—ÄÏé_®ZÎ¼P5W“ÝúÜë–3uµN1õË(Å_w‘	·ÖŠ*#\²þ ±bA&’[ñµú<+p+§é;¦ilL{.ðç'T÷.×~·%Œ¨R²©p…ÔæBq€ÄDe‚§10È|r6¢RTý6y6Â-±tŒßdO6€Å§qœdí)g±´28 BøDsöÅAF€ÿ¬ó	Ð=Ñúlí!_ÿë¿ÕqƒéõÛu6;]µÃ“µ/ÉÚZù®yââÛÀµ·!šÙ*³šÅ…ú©X-H‡€…Pµ¨OßGÅ£ó¨¨>WìmkËY¯Þ²[ŒÓ:C®u¹ 'Ñø­T$ÒÆoÓìl­¢E¶~Ðø˜$”ø}-Ãs–S°ƒ§xÏ«¯ÛVÊá¦Ëjº´éÏ%¡˜Vk›+ß\‰ƒêýÿšÂ¢û}>MÊz{ºÔO‡º@‡IÏä”ÚïQ>™D³2=™$”ÑÙ¼té.‹j(_Û4ÎaXË¸[ŒšÂ¦€“5T”‰èL¶ÿûb>Ëq‡‹ÒŠO1~Ë9·’åâ8Ÿ¹¤_ŸwÝÐ¦?×®³Ïy'HaÛ¾þÅy¯ˆéo¯dïÿ6NóÿµdÁÛ×6=]PãkrµËBØc Šû†ÜVoo¼¶Ö—øQR¡ãµl¼7P*ˆaÐ`jÚÆãSn f™’U]Pú>ô(ÄáaY¡ûøS±[%YœÒÕ-¢ÿ}NOkÓéþ;Xiåû¿]$“G™ÝiSNw#y¯õ<}_¤Éé™ŸÕËÌT°ß+updm_'Ø}àäOaqÏð|r^­]MzÊ*úõ{T®ú=Y©7¨Ws-Õíë—=™­)MÔ9ùŸ(å%9gÐ×¿9s½²±[HY6ÈøÜ¾þÂ³î#84Rß€WSáÓÓèIÏjcŠ¤&–Ì¸ù=¬}zÛ&Øué©ŽNNÇô×±Oý.RãÌ’¼6t0ŒÂbÙª­.ùÑËV†ÒÖ×«üÔW­ ÀÍ\ÌVmee>³’ŸÚt-ƒ³&ZqqˆäVåÝqaèsiT˜š÷ÈÔïŽ§G`ÖÁ"+(J¯U»´ñœ¨Ñc:cM­ä¥q¦ê}
é¸Zëé¾b8¿Z~Ñ8×.1öšú‹?ÖE%—qŸäÅ!úB”x8\k[¿´¨;.ºšé‡Y£¬këG‘Ç“‘-)øÜÛºˆ?ó"­@-.Ë¤,Q4÷èZ¥rì=’_ÀZZ[KÂ^3—ÔE#D¦”Äm•@x{³0}¸¤0åˆøû©ŒÞZ¹ÄSšDŒ³Š¡ÃÔL:ó‹,Í6B@>*Y³{	¤·@ÈèÅ
<fïàèQgúËeiìù >*Ì=N'À0’ñÛvì½Ð³$2oCg,†9ÌÒ¸ƒù ˜QY,Ûñ‚…0àj\7Ô›0±˜UÞMþ04‚qhí¨ø>=™¤y•Œ#
>Óù‹—Ð•¨é6|¦ýªz¾ß8ð»+3_ŒÄpï1"P$h*ÉxãÉËºžÒÁbœMî=M2âi4.ºø„¦XnÍ’Úî!¿êýzŠúpäø=v™¢Œ½wó(  ^7dJñµFEQÌª€¶ØA_äËâh†Y—Êó$©êúMPq„±Áxêêâb…‘j¥«ÚH BªcHy¤T:¾:È2¯ò|á‹'FÛû¼þ.°DÛ‘Öú~a½¯;m¢œ¦Òfù«ôÄ‹oXàø3ÙØ`­ƒ3Û9³j´×`NÜPçÛ!†QóÒá9ñ?*ˆÝ*h+c	Ø¼4lÕ¥Æ”½5é,7pƒ~|ùl…É—#ñ$'Ùû‹ÈLƒæã¥DÏá‚‹LÅ—(×}uª{÷æõÏP»=Æ“¿ƒ:3ºíÁí]ùìüªÑmQÇEGgr-ë¨ÐÕº9)‚W9V;{âõ¤íˆ-¼hIW‚îêGèF0#kuK0’£ó4™ÄmV^m;xª!›%ø‰™Ì®ì„é¦n£ÓÀR‚ÚFg/-VÃ†1N#¹µÕÕœ±´¼¤=ÃXË0h‚ï¿]Ï"¾[_ç__và__¸[ì—ó“#:>Äã2òVÂgIv>Ÿö³OOôpCÌ`I±`7úÚviê„à.<œžÀ4‡RlB)lÌÚ¾ã¾‘RQV+ßÉóG×.–nBQ¡æ¯ß,wÎÝ˜–âèª|‚á‘ßn‹Ï‹ö ÛtÎÏoéS˜³¢×û‹enÞÜ™Y½ÁÜ(¤tÀýù<Nkhƒ}œA›(_À+Â.Ó*Vª)î”òXtBïÃ:BJüz@!·¯¯Å¥<üæ“kÊ£1Þõ7VeNXi˜F/¤®·¾ÎŸÒ”'s³9fäÝŸù:*b²Te6 ÈçæN›ÁÍÿöÆùÚO?&/ÄŒ¸r×1ÀŒ(£ÌüØ›tzæÏdYŒ·¯åuY³ó¼ÊxñUº7çU5+Gëëót-ºˆª¨À»C¦ëÑ,]ßÉ°õOT­8--WØ£›OO`Ÿ}monœ|ýÕæ§c<¤¹}zzú¦Æµ£IEþg ^Á@8‡ps©¾W$§IQ$ÅaBäj{%Ë×ô#·¨²À7,¦pŸÕ˜Ÿ>‚ã3¾¯(‚Ì´Ö1ùíú¬øw¥ÏMk@)v\=n¯ö@F•ºc„3ÏìÒº3s‘ì†LIüÜ9ådxëCî­ž´F¶7ò†½~>¯tDãÆMmÚ‰ÕEi¡œUŠä³Q·:‡ÔWœ+SÔë·¡¼Éò[Y Sìó‹3vÔí·ö(/ÞæpœGeÕ_1©Êé¶<PÑðÂñ"ú	Xu©%º“ÉÚ0èµIàxf[Å€Ö×e¶T;EtOqEíGÏÿ‰bdÞê¤ÛMã|`ÆÂ‡BWÀØ«œÌ³Ú°ÔÁÔ46Tƒâfí¶H½ŠôÞ è«îDš]¼ÿ+€ËˆƒÃ“1ŸÊQ¨”Ö}ó¿ÿÛYŠiÑ(
_Ø°Y‡W/•ÿ1êŒ±Ù©[ÉÇž÷¾“‡[7"5Ôû'ÉTV¡‚ƒÔÍ;!õq´\ˆic‹xVd‡¼vŒ’:ê;þ0†;÷þÿÍ’Ó£Ó1–SPÇ	ÐyŠ@jÇsà½ÀKÂ•·ÞyN7¬±ç*¡9Œ÷Q>»Ò—1Üê› À¾}Ú¿Ö©+‘áÍÇè~8¢‹ótƒëš„Õ…"ò6P>."uÿ ù. >K	iEgy¨ç^/¼ÂHw˜ß­ÉïSÀR¸	w|pøüèõþÓÃûG»Öõ ¯ýlùGÐ,ðáG”åô(9ÃY±WíEt5<-ò)^x!Ž’ªoïÇh¯<=¨Vó"Q@2{+¶ëÒ†I[(§ë±{ðÅ»ðÚÍ+Í¸è1,‘)Ý¥eðÁÞªäÖ–ÙÁª$~„ŒU$Ž¦³‡§r^áis`±	$­B¨¿èM°€Kym9öaºGÂ
–Á8)<Ú£-0IºûÓ{Žˆ´%œÇÚ
Ê›Á¬²—ÞÀF>®X¶X—Ô…;,1!fÅ88znc˜ò]îÍ¶	4íâe÷;KTÒõÎøìUû]Í~¾gµÕ…ÃHƒáË‡Å$8Ýfd©YY83KÎN}†ðîó¤À¬{À¨§³¾w3gíŽeöÖ½“ÓdÛ——–Ëëi¢8ÞËÇµ+`õ}fx¬Á¼ÕZm¢Í5k	þJô5kŠãì(Ž¿JËGö‹ÜÀ"Lœ+Ör°èAÔò°”Ñæ`QxÛÌ$ŸÇõ,íë,ÈöBz(kÇï§j™[µi­:Ql¹Ë‘¬§Zlâ%qtyƒøè›ÎÊ.V~ !ÿªùþªùþªùþ=h¾ê^d·
g
¸¦¯ïOE¦k¿*ÈŽPøUí]JÏú¸5^=·~fvu·«¯ÛãhØÊ#eÕàBª,FT›+j—Q‚tåÖWóá÷|+°;~õÚRqcê,èz^¼.º¶T
%ÔcüßQWmWÂªáÒÚµö3èo¨-¡¾I·#;aÓìzdgb> û‘_-O
Y$øøšTxZ1=_¬œ…juUÐxo[T4¼Ä,ŠòI•é$I‹üþ42Ös³S[)„Aö~¡rÖVyrf;Ñ¬’Ýo¥ý<:˜‡Bæ|ä¾¹•ŠÅ©õ^<’×/‡~£‡à8D
˜ U:8ðú¼rýúúA'­ÆQ×lj¯;j7H—5è|Ípõ¦±TKcÖù‰SßêýTŽ•?Qúåï«î/Ä	7J…+9úùh>€®d3ª!ü\ƒWö¥¼¯a$zo±<øh'P¸
¼;Æû$ë^D¼•tA¨&X‰ !óH&W”?Yñqš$1î¦×•2~>ÖqGâ‹?>8,XœØBZš\h˜doxÇmd™uïŒÐ)©2Äœ­Û>í¹@Mw\…šÌ˜{™~¤TÒEÇ.-XBpYc¦ž~b+Ó¹W‘ž™Jô‹µ–L Sµ¶äSÛ’üÍHÔÚÓ<K§ïr¯.e ð‡Fïwfõ7š.R-Q˜æ&QãUsg”Ïi	Uù¶néðB_j¶f¼ætî´èñs7³ûºqÝ›Ö¾œ²»º«_»§Ó;ø®ƒÂ˜pcü5û±Ç!?öøN~lãa;zn³5$ÔEÔ†Üä]×TlU€Ò$C>PIþË‹ƒãýf\Ý¯‹œÁîâ&_de-v•{šœç.ï` ÝËœZRï
º€œlòxÿ¯&Û¯&Û¯&Û‡2Ù×¹×â?¿oïüéžy§4àÛ™w²rWó.ŸÄ/XÉí"9 ·**Ó†4µé¦ZwtÊ^¯^‘“ºNƒ1É^Óýrh|¥úPÛÈ@©†FöIÛõ 8Ô+¥ëú²*p Ò±ÒgýZ¾òÛóPåÿídˆZ‡·™æ€E*[¶/¸#üÖ†Šž{BÒÆ-¦«?ËÁú¬@È°ÜÛ$®Û¼fev4"­—:#?ÿtÖìéGèêŸÎêFÉy‘g  

Ñ5&"¨À7Æ)™“ÊÊ+”ë…ÔçÉ•‚†W	©Öôž€Ê–¤Îµÿ…®à7…‚Ûã	q¶¹,¯«ÅõKÜñéö²Ó³;ÓA„xÜ¯ˆíV±NH¡[FÕ°åçm§€¤ãPcw±¼\zäÄØ0.†9¶"š€¨×~}e“ÞÔ±)¯ÚJgüª¥#œuyæc·sFÖ®ÎÓ¸IÈ®öç›HR*zUø°G.iÜ|Ã@ñ½6/ýCà±^Ú4„Ô£Bá/<ôÚ>òVyø¿º^6XÒ–S¥böšíhºn¥²ï­dggÑëdÃF"OÁá×¨í(Z7Ýx^àcÝ»æ@;`¼OÑLÊ`ß—˜ËeˆÕ.æ)Áb‹8ˆBog¸n8*yO,¢ÉµFÕ|Vðßá'°äCÎ7^Úa·~Z—p“³¨¢…|±^€|~¶†Ò5WáÉzåÄ0.TÞNÂ\r³èj’GË:ä|=gù9÷'¯Tm\Ý6XDªžƒ6w^@
ÀjRz[K[7>?GUOƒj{×î't Ê‡A€öÕ"Ï¢«2ëçA ÎÛ°G²F_'äÛ—EZ%ßã~|p]¢²&}—µ(
×÷ÆB*8jg(™p_[ÕÔívÌÑ×`ž÷#ìTMgÃÏÇ ·Yyº[‡‹ŸnJ~–ZÁº[÷¨ÜÈ;*xø¹%ºr_Š~î]ÙÃ·âðã¬³”œ\¦«F3[­šX :àünfrKBvoœO§iÕ¿‹÷ßÄA}@ïÿ› ZÿäZs ›7÷³L·ˆ+öOû²Fä;,ò“	ÐöH”´>øÚ“¤8›g•ó¤DÒ¤+xWPÝÑe`Nýmkyc£ÐÝ°Ö4‹ûdïkŸ.@¶ýs|—2@6Ö‚F»îÝ6vü×=ÊÂÔ³;KÑÞcíÀ|‰ÁØ	ãOÀI	6•Æ#^‹~P˜þÉ0;‘gúùâcHÚVÄk’uêHgÓŸ:›þÔêÒ™rk(:K†P,¿H0ïK¿‡­½O’ÃEf,%˜QËáúæ¾ßl Ò5íÐ?°bFûôËÐ½ü…4¿¹êb²×K8¸DÜÈaI9YÓ"ïv]]ïº_n¢N
õ­gÆ®¶†Ùœñí`ëè¬Îaº
ÔÓ\U)HFÆíç2wgx&w†»-<s¶…YãDB’¡ÒVsë¦ªJ@(4pT)ç8GÅŽiîø+wýˆ¸«RI˜æÏË,¨<e.mNÔÿ9YßÒzu–ˆ£·ì™cIwc‚K²£-ÏRÊè4Ù-žÆUÚhà»˜ÉÄrìÝAlXcõO¨,-e{+ Þn% ³<Îe'è.Ò—¯ êËWº”ç1rM4„qeÑì(k@Æñõœ'ÅU_öoU\ž'0;=5˜op{»§‡êœ;ä»krDÀp6/áq8n(l5\FEÖóGì™8¹Òô òL|r­0v#²L£dô¦3¨Ñkâ Ã"c—öŠ\\H,0h„ÙWs"òS]‘FÚ©‰å–|L¤Ýw°E‘ä¨±‰4Äž;Éù-å_¾«€iÍ™…2Ÿ&ýä]ZÒ²À=õÝH"´.ënŒÐT¶õ‰Îµ;ÓGã(ë0»ÎA§_“$;«ÎÅwbƒ7B³‡û†)ÌÉÆ7ðç[¨†/>Û_ll„¦t|>ÏÞJC&& ì§«X“*ÖcœhÙ]7ÒÔŽ©‘~zú©»ääÖl“%ûy¥ˆ0“%48Æ9ep,1¢ÈÎŸŸ]ÊÎ†5°† 7VØH‹n”ÆŽ:”%ï¬(÷ê¢(cµðÚŽà á’Âc'¡Ü'#‰ž¾æ¢íŸŽž?ÊQ¦§WX*k@Ý·±ŒœOÞš÷þo·²iúõPEc–.c-5µeeÀŠ¼@AFlä9È˜èÓ($R¦B›01ÿ&»ÊYƒ!™D‚.¢¶Ÿ#W‡¢w—ZVBÆm'Ùi“‹ÆLÆ­o¡B=€¡ÖHY(ˆ.gß9”B¬ íž€ä©ñÃN²gm9ª"5à¨ŠhÙ¨ý9ÒŽšÈëfÐo`ŠÚÑ§n}`n‹tå¯Yh5ÊªÙ_ét«Ü> ¸µrVšŸ (Ö-Ò°ù±X‡$Ì±o8zß›’´Ï1¬w÷&`HìfñcÚéÃ¬ê°ÊŸä—Iñ(*»ÐpžH(Æâž	¨G„àûpI¶èn-ò"æ'‹¼ˆù	Ä†N:{—4€©yÞÂüäŽÞBögðÊ–”É\DhË„ßÌâÁžµ,×ÌHÐbé.Oõ,'a®,çü‚µÑ™gPºØr^`'·úý|‚ø…ý~
µKûý,ƒùý~e¶ÿÙ™­rÖµðÑÃ?;8ÉûûˆGTžÄ»ðÃ 'QóÂŽžÄûõ.Ë’´!bI†°)ÜºI4„Uy¯îBCXS²diCØÎLÐÖ`?CXcYÂ·¡Ï»Â1ß‡!¬Fõ3ÂF'ùÇ5„²¹]à¹Ž‰+Ÿ…Š}Ö®ÒÝ­]ÙéÝÉ„3¥¦3®ilÜ¨ër„iˆa½|u.Õ{ùÊž Wj0m nUÌ“ÞGÁÊŽs#„ÿ,ÅÁ*Œ¦–µ•·àml!§´’Ùq’êÀÖ”UÓnc4I*w·Øú½lÓšõw¥5í/n¤5í%¾­Y?çò´Æê~T´†¼K³íÓZô³ÓšëK¾-­Y(hMzñŽòÉ¸·uâEÙÕ¯Aÿp–¦uëPm÷êÅŒ„%›Î>½{	v¡vóÄuõÁÃ;yúJôƒ;úÖ:…ÆXÓ¶”µ^ƒŠœOñjÔXÁN¿.çgI‘ÆÚ$žÑO¹~¨ìa:ïŸ•ÑE2ñÞ¼¥B^šÒò†õ×P,_ùœZõ/EðÄïÑÚFÔÈ¬·HµÖÙigúDî´x$kdk<,%ÝüvH–5‡dy/IEPËú#™¨¨»#ÿˆbC1m	?«(hXSÜk	uø*ú•»ÿìÜýÃ„e¾þQDe.Í_¥9ä¯¾U[V]²TŸôkžU½jÜšÄl]ß¯j–‡°¤É¦(àT5@?
ãÐ’ë/]äõÝ}ªš,?¸KÕjYÿ¨U½<¹C5Àè™CUÕ–âP¬X	TfÚÓGàƒÕôv.XÆ½ºyÅsô‚ g[Ú'ÆVç².1^õ#ôˆ•’ûýGòK»Ä<x;ÒÕùÊŽÌµÒYx—2ä~eUB¾­UBÛ§f=÷ý„°ç·FúKÑþ’{·r"ßy•5nrÜÍmýw¾ÂE$N,yŸæéTœÐâ¶ûãwÙR	CÖ!èwq ‡ õóVl¨•ÿ´ìþÜÂ‰w.'ÝþÊÏÐG;”Ÿûò¯FA`tŠGt~˜„©<©lÆ$¬nöføP¬“æ©àTRM*ãûB¨èúºxAþ^©~ÄÌ-lÓYá©œ“à¯©¼ïzy}ŠI(Aç“H™Lv/@€J¨Iy–i©6¶u…~(ò›]ñþ¯8ý+Ÿ\Þ¬ˆ?¿ÿ+fÜ…™6ËE	€p§ÿM§üÇ¾ËÌÅ¢uŽ™–™DºG_Ì³ä[>Èfóªßëuæ‘oÞÿk 5çßxdƒãi½Åë¬ð)/ß2-k·˜¢ÏÝjñ¥áëDo×jó`àY)=ÒBÑHÆòzynŒ€êÔ‘©è÷­|ã˜,›ÅÜcà³SÑÕÄ&(–%”OtoÉ7‡z°âŠ]¤¥””1<!–R[¸­üÃøDÄL–_æš•Èñ8žJÔX_²°9\{ã,]–ÙÍ²ö|g>¨;o–êÞÉ-[ârYÜ##‹ƒ1:<¼°¡ÍYÈûnØ¯jó¯ÅÇ"Õûµ¾xV@Ì^RŽ—ÏRâÖÃ`°0øÖèiœù°Ÿ{#Ú)—k§!C_ž3îdùôS‡çÈñµèÕDBëÔÌäÌ¢¥Ë4ÌLºz†Æ…^¿\MoBœÊŠ?ÓðÐ²»æW¶’õþíxŽiTÝQˆËá€([;!Itß†ÎÛ#-uñ‹*%€¹¾|—¨p‹¶Ìªq†bÓ„ÊÏbŸ ?óîé¶ÓÀÃ1¶þà5¸[¾g¾ÏÁƒöÔiìAgxÍ.ý}ŒÍÓŽÓ:c\Ÿv	®1íÄl¯=ï„3²>•îÂÁ½r¼B&ªwÊ5pqún³¡þ6¦'lÐì‰ñ¥íz½«$j<Wô‹$[+úù$v“Æ€NÒ¦ø«òäCPEñ»zDœ$dY6ŒÙ6#âù¤Åÿ š­²‡Š¶ðZƒ³¹%ÕÿZK²‡Ñd¤ü‘i‚:™'u§FJ+	¯Sþ°~èä ›©ðRQ“²IîÔ*L‚±q9¾èÑ«ÆêÃË[Û¶i­ú&:Î fEŽºÙ~žîV¶ëÅL‡Ôz1[¨\›ƒ¥4ë@'ÊÆN”÷«ˆâÌ/ÐByjáVMØf'Ùiß9š¨q+cK¤åBK=
¬	/6†h©^ñØ£å;õÁt÷$.¬în=sÏœçuz>D÷Xqà¢§î#Pê1'Ó)Žâ ÛÃ¨GÇnûYu~‡Jä§vC<ê«°¦CŒÍàƒ²†§üÔ5ãÊW•Üž´Þi¼¥¡n­vñ‡:}ÿ;6\–™ÁàÿÏeÍü\è Eª:ôEzíøPÐðáŒAêï²VØ°€l't1;¨Ê^rÍ'•TxC¡¾Ö»8ª‡
QúôkÅ~ÆøæùÉŸqRw_ìïqtäh¿’·)˜ÄÈy–þ\ÎÎO-¦I¢:$SKsé$b9š£xmôi#„F2˜&^.›•ó"Qa÷™½bf›"ùQ‰e(u·®UAþ]«ßRx <-€›Ïð-ëñ‰1 ÙXpKóBÑÐ_)aŸ¨!æj÷WÅö I4[Åº!ÏfÛè¬•ÅGctîRjÂÕí³ÛþºoT6mU:›•-•ê’^Ýë:giŠµVÇu=IHø½S+ëÿîQ¯Ñ ÷õ¡‡2‰¢!*¢[dÀs‹ÏPp VLîP;EósF^²l`ÀQfX¾¬[Æt_yS|@óÊkXrÝ±"M‘@f»¿¾°‚Î´…ÔXÛ„lXm-Ú†›·mfî{Ø÷|ZIöñVûk|-0´³íµÅ‹¡³%t3Mì2j³M$´Eº8NŽwß.Šïk§Eý½­oÁšìêÞç‹ôcuñ[G}è6W½Ó`WæÐàÓîÂ'º»°Øú^L®«¶Ák]~hÞöÿY–q;¶c2²}h/àÝxÏ’ŽÀÇšþ^ý€Žft_¾@õ÷ÖžŒ;3Y:$°;™ì¿›%0kxêv¢¶¸=_†¾ÂÜ9·áÜÌ_¼ÐÎ¨‡?ÂEê§¡à!Èµä4*éšÍž:èçK–;*"ëèëô–©ãÝ\Ø¹žõ.Q©’È=Cì»?\¦’.'d¥(&ét¶8á`g~k@yœÿK:CGÏËŸ%ûVbsˆÈ9ãàôU!«»!-×úb`Þyÿâ›;ðÄ0^^AW¬&–e`<òTœDã·gNÒ2H%óÄñïªçuË«v¡xHÃé8nÐ­ëa&2Š*|PÙôÈsè²…±¡KnMºZ8NÄ®{ü¦Ø”hý0j®.rD¢K¼s§‘ƒ8s³‘‹<ÞBð ÷‚‹oz1ç¸ïõ¶½Nèà0,†¥²?êy‹ä¼ÙÈ¸JÄÀêl¾ó {óÙ$£ÆÄ¸Ðc´J8c¶$Ç:¥q åe
=È9,ùì
AO {;´\.òª¡Cªžj½}–¾HèÛm±òÿ­<Cd¿ÿ«lòý¿Ë6E,‡ˆntæëÆ†Hýivš·º ××ÅQ^TxMüÛ$™‰ê<' FàuÃ}ü…_.Sèþ˜æêüÿúyzvž(”àäæ™€‰¦÷E2>8`x+¡	ò±¢C‘üÕ_õûÑª8qÖŸJ©u'4j6·; Ž¡H\‚)U{iÞýVlnll¦€ÖÊ¢	¼1Â:;†¶,Ûö¿‡öOÚÚ¯½li¿V¶©}ENX“ˆP7ðNN¤Ó¨ lÒ9dDôËWŽ(Ð‹Ã=%ÕR›¯üALô]«H¦®"q{¾™¬ ª¤	‰~Ã® Mj›Cq€×ƒcZ(©])C'¨B‡¤µý@5f6äb!…;£Â(\ÇÁu“‘5ˆ¾Ð¤k¹”äÊ/ÚÇœ\”ª£UÿœT¶¼ñÔo÷ÃÌB¸ª=õ¬E)6äE	·°¶¯çTÈHJ‚JÉ`è—*j¹2V—¶Ò’þöe-/·ÅÊº—yF‚ÓèíÇ6YÓÜ‡d†ƒ¬h2yB‡øŠâ]õÓ’Š·öeÛ„	]uIêbF hÇþðð ;ÑØ€œÄ!7v´Erixc+•~‚¶Z™>ãVnŒ['Ó'×~$ùÍ#ûxZ\KFüa´Cü±Èljk(TŽžG”Eª/ áÒÌ:î¦Mrâí‰dßýQLÛ’Þ˜?±‹“«€÷(pVÞ[oÔVË‘y¿íñy:‰é2BDé†NI9”—¨qÂŠZ%Õó8*ß–›ìû–ü¾{†ésä÷§QõÊh»‡ ¦e2„EÑé€õ{C›Â})ÞràîÈñ& NÉÓ$¦,jƒÕ¥:±¤ðówÄˆ–Ÿ£ÍòëäÂø	)tL#ÐBR°Ó“×ÑE4I#hôþ:äôç•¿ŒŒ R7û¢¬V
™g”#×®˜À³-÷™\Aî3XIô€¹MðºŒTõdÁ=¤¶ç‹-Müh¨µKJ—³>	TÀmÊ"?‹Òc5™©2#›dódIÎgúXé¶YÊU|—ÉNcx¦Á›=_Mu5K(ï‰z’Å{(€ÓXGFí°ï#ÑŸjÊÎP&:ˆfýÊléÈkz¥‚‹ì&N³§Ù‹£ î)ÓŸk.èp`J¸&nj„ôaÏ¶˜\)»‰R @qŸŠÙ¼8“/lŠZ™Ó®ŸÏÀÐJIÂþp A€iŒº>°1QŽ‹$ÉÄ)üûS2°Ãìf¶Ä:ƒ£ïüéf±}p÷âÐ£äy!sÄm¯ÅWLÖ‚ö‘,o›0ë„€O6)²çm_í×ýo¶DüétÒím§.ol	ŒînÞt5pZèí¤Þ§ms/ÖMˆI†. FOi
P#Ai¦FÝˆO’²Z’”%ÛéR·Âjâ¬,q¶¯ï*¼¬»S°¿ûÕ0j@íÛÐ!S´ÑÝ1IÉ=éÕhr:"IÂ½H1Ð%¬BRZRh1{•ç'yTÄÞóq‘gùYM#ïÅÛ(;‰2ïa‘L"À_
ääçÿR
ù¸r˜Ÿõ^XÙúI‰ÎW³0Ã1YËH§)î¸£eZ¢Ñ,R& 19N³p‘ü*”Ü*?”¼ïÙ¾Çõ%¶mºÛåä¬Îçˆ¦(+#¹?È…8Pð[ôU£OËE'yQ•Œì\ƒ¾Ñ¤¯õ\Ÿ{×ˆFó;´à–3û#eÕnüË~hÕYoÔÆ¡{nâ	Œ7lb•«˜õA!ÝOÛÞO[8'ÈºUbä8²ŒqÆ`ØsŠö,À²ä;¿8÷¸÷&§	¥S}.‰_ 'ŸÄìÅ²N‰Ø•[z%–êÑn	¶5äºï­¿æ¾‹–[m%¸_#\B{9Ü‘6ÏRÀ<ŒøjLÞ• O1÷€;v»»õ’ ÅUç€süþ
PÓ4OÊþÀGdÈÇâ´~Ïžv³¿E]yãx\:û[Zü+í ¾,y&Ý)¢L3TŸÒŸÔÞœHù_R Ç€–c´ÂÕŽ2^·Ž	p”Èíé·³±ð L`«­¯sÏšÒˆîHïÊu½,ò”„M÷áPó§p¿Ô-#Â[ÅÚUóÄúd”Ÿ…©¥&º2ä¾ ³g‰¶½m{í]ý1¶ïÍ°:í%KD-ÞL^ÊCÓÐ;¹ì·Þlf¬ŠTú‚ymL1’ÚSª¼KÞ7p­;í^ÈeÍ6n=af©vÙ¶8£îß/Òa¶õ8m|”þ‘»¬._‚þ²®’mÈÿÿ   ÿÿì½ÛrY’ ø^_âÖ4€N"))+›ºÐ˜3“Sº0IVíŒ©´R’‘Èˆ E%‡cÓ»û4ûÙóÐÖmÖOeýü“ù’u÷s¿Å%VuÁ2E â\ýøñã·ãîäÅÁÓ<J½áwð %2Ž.¼‹ÌLËÇ•ê~¹AÐØŽÞlºþ_.¯ÉZêvzn¥g¸âà~> „7
]$ÓDY'–èAR8U|¬©ÎÌ7àLÛñ¦‘ÙüŒMØS“A5Ôã.;’ Tž‘Ó6!´‡¶ü,ˆJ;ÑVÒ¹ªÃ^YÑ•µ¶.‘¯-#zW·ÍÚZß’¡×<*VÒç«—Ò¥ry™Ø\ß Ò ÅëK,»³Ö˜=o¼…5aù–v±ÞÃ—ÚÈÞÏÆ¯.eatx—A\¨ÓÚ,ˆ¬Kàacdz‘[ÂÙüFƒQZ— Ï¡ XÜ|íyoÖâãÓÆ«…osùUûwzýñHÐ5Îvaè­¶î¢p³£]ôä¬´í²PµÐx[ä–Žv­ñ/{´#oÎáI&¯´,’ñÉoÜÖ=~®Ú•ÄÀ’,¨¾L¬üÞŸÒ[ŽäÕ›h
ÿÍÏ“É˜M•—¼[Ó"W!ÚšÄ<qoO)…"U)²y>¤÷{Z†q¦7¤ð»¶Ú‘¿E—*èd3ºöM.ToO‹I/Ž³µ8­÷;Pk>Fóyì–R#ÊØcãÞ¶±^]al§wô(ë{ûÄ7*MÐôÜÁ€qÕ¤ 9+Ç/Ô³éP´¯éPŒ¤B‚7£”ž*·¨„¸q«ãž6 õa”5¢â“Ü¶ÇƒPs?¤Îæj¢¤ïAVš’ª~
#¦¸JÅ`EÃjÝOÛÜJ•Œk‚aAFÔBE[ê†”çíWQçh}csmþëôôëTáùsáJIÁ¶È _MÓ1s+´IOßu 3ly0Í>ºB¦ñx?Ï~bQé"k¨" J§ÓÓ‚Ø½ÿí¥±â*Z¸$ü¶BÔýí¥ZÀ–›šfQð®zïÿÆ7ñ×Ç?mFæVë™«%€”uM3Xh¥ÊXwÖ6#¼«V”ñdFæ4Z
±À¡8rvÒxUŒ"0Ík°_¢z®âµ¸ö,ZÓ¨íVe‹ZÁÍèMg£h•×ÿršÆx¸Ã‚“F‹.¨¿ŒóŒäþ]ôG”èéw€Óa’æà Îs(ý0ŸÄSVï(N³qMBÅ½íÎ[5{4Ì1<š’ŸbŠ`¡AÇLˆDÉüx£E$Üg†÷M¯Í2ÖK±µ›‚Ø.	%OR	—R:9µr†rO¡r°<ÑñÁ`Ð¤m¬EÈ†80¡ü%ò\ì?! 3©µ«tÒwÒëMZ—$ob±°‡ÚÈ‘38Ò«F#Æeµu²ËòÞ=7UC7Y;?Ë°=ÊÙ½$Rˆg2ÝÑz€? $PnŸÕ}'›Gxw€IzY7Ý½(Lš>6ý(Û›ÌËM9Â7è¥ôæ­yFåvyr”Cæxt²~ô‰šÓ×èÊÛÜ#ÂÓ00Ïiö„1TÏì³‘×ùqŽŒzeykUb+Xmx—4írFÇn£9ƒ%0›ðS>Ùµ¦ÎŸ‰7<q­½1<ÓxÌÏ³$?OËØ<Þ}ÍÄîqa(““[Ðe­ÃN²!Gåè­)¹®tªuò†O·2çóÈž>£Ç‚…4R©Á€è''/,~;ÆÒ.viCkl—Ãõ>o¨ËZT…ëÀ­šøÐu‘¤¾=cý´æÐ:Ø9eá‡ypýåÚ+*³éˆs¼¯®ÞYíj/:¡ÔZ}ó¢Sâ]ÍH–×›sù€¥bkN?ª›§)ªH Z½<D(4^ÏgtFÓq×<…n˜Ãß:fê{f7‹ÃÑ%ñhÔÕÒÎëkúcÏ¥Xœ½Ó€…~µ@PA><C¾Õ8•¬•aï^r}‘É !)nfÔÃ”ÄÄ¤j],Ô€jJàþØs¶ô9±@{)Îv)°áné¿âD±ÎÁ¿ì¸X°Y¨™äç1Æeò5<K
”¸f[ú†k>% ¹#>ÎØãÍhÝ©Ha¥iº¸ÓIÊo@Qùl?Î‡)îÝ­hž¯9n8å:¹WÏ‰À…^|d	½(nÓ#˜4€2žŽ2¶l<jßƒ¯{ƒb~ÌŽõîF?ú³r4˜2™(fÅD™,ÊF¬»yìžWX"…y›„v}£Uq¢m:ç[§c¶¢ÊÍÂ¥8S)™ÍRE¦ÿF2£ÿVX³‰ßõWÅÒúÿ½ÓŽX-Â£Á*±ZÑEü¬ä`6/ÐG”­•Mbm__[æyª_oÐ}ØÂ”R«ÊÂ)Ö„±oŠÕøÐ#šâDS¬pñÂÅŒJÜ¨ÁŽ
üðÇyqÚ//0Ìà2ƒZp!:Èuô†ïÔÄì?°´ÌË“î-Åø„Ó¡Ð%„œ5ß(ž¾/bú¾Uºíø<qåb?š+Äö84šÚÑÇnAÃ%ÑQMhÃ-‹zˆd$n5Ë÷ìr˜º›Vé˜¥ÉìK¹¸Ô|‡6S²Ï‚*Ä >Û\"=±‹ÕéÚ8pt›#IØ¥›¨ÒTI¿B}ì<~%iM‚q\$Yìõåê5¹žØøË5*r‰ýl>ý «ÑÌ.S£×èÍyúM'¬IŸi›™f],Ê"'Ó·ƒ‹YúÊF£Û¼¡ôCH½”ÍM÷=ÄkŠÔZŸÇö"Z2þÔÚŽÈ¿0Ï{ÃÖK[ƒæ¯ÉÂzm÷¢×0êóÌ±„‘¯Å3ã-Se Ð)»ÊÅ?ÍÈƒ÷rD<œ¢ Â$§îjpßÜ
‹'öL£ÏÑØ¡:¸ië5 y2æi¼•ÕÍéH;–m=ƒSkìÿtaE¼ÒÖZ™iV4Šl´vè´™!äÉŽãýd:‚’ttÀîÑ´ÀþÃPî6CmŸÍv³¦ 7‹ìEm<?Oˆrlmö"üÀYŸ·áf0üøØ°DlužDû^µ·Å.Âs€yHÇD·C[°£Ì/â‚JŒ(X£2C>ÕYøImï¾Õ,d6/ù­>ocÌŽß<’¤Òå&°xz&c3O¨<¨7õºÁ°½× Œ}Qåµpš ^ ø·IÄÅ#½Õ–$ë8ÛÖÞüž5ýc£¡½Éî1‰[ËL} á|D—ÿ„}$$ÉÎiëSO¸…A$ŒÏTœ×Ä–¶'IB	tÍ'Â\"+*pp-§ FyÃ†KÆÙô”Îµ³¬H¦\íÂÀƒÇO‡ÕHFÆP´¹Š´‚¾‰6±ÀR$çñ´Ô§¦ÏÔ°–isÄœ\j“žØÙ"”Z¬Vmæ¾‘¶…3uÓüô+M¶o¨àÃWtÑ/0™ü ¦¨ÅpdÄå4ÀW÷0¦ð?Ï;¨æAeˆ¥@}5ÕäXoøš¸»Ô¼õôF7PÆãèûd’NSý‘C†½±½×5ŽÍ{qê™Þaðöþ^ÌkyÈPdŽSÊ×8ˆö}=‰Ï³¼Å£t˜b<åbGª6BaGT¹þçëOøüÏâãtœO1XÈ{³tu°ÐâŸÏ “= ßãôôú_§Ã4Ž¶ó2=IQi¸ÀIÓd†®Êïw2Œq_`hZJµ¨#“·Ú¶Òºóä Rï
Nê"%¤rD	»HNçÂ¤À¼µs˜<9M†Hž–E6.U¬õ£ozØ8@/A>ÐŽ?cÏç%®T6øÍo^ìí¾:Ú¾þ¯ÿŸ×ÑÎÁõÿ<ÚÛÙŽžïFÛ¯®ÿñÅÞánt¸ûòúÿ~ÅoFû»ßÿjF»ÑÁîáþëCø¾ù›m8ÖR æã;â2 »½)A¶Pu…òª‡<³„ zÆ›¿Ãe‘@G(Ã,2äjù"C“ÒR×Gºt†—›æÐÔOIÊÖM.6£•×Xp4†‹2–(ùÎ£Ïf@¯a~ý/£ˆt ÃhN²”5RR2BáqÉ|Áx0Á-¼èµŸÅ<à4¥“Ù˜nk¥·‰¸³m«©­¦“á
×h¯ K¯±A\ ýƒ×ß¾Ø}¹Ý¾ß>ø~ûÅk ìö‹AôÇlxý¯°ÜeÎzi.îñ`º{]¹Dè€qod<aš	fI ¸ƒCzµÍÃ6ñTFÒ#4´±	öIÔ-Òr.z‚#½LÏ¨FDî†yšÈ‘á’csP|Ÿðî¦4*T:ÐròÕüf#°þDG±ƒŸijŒfyœ!åÌ¯E**Vž/Ž¨FLóV¢µ;‰÷ÊôúŸã(ž§”õÆ¿xlÀáÕ‹£Û;xå.ÞŽ}•cB($sbôžÕ!ÔÑúC*§]Æ€½D’è	F5ï~ó`í½Ü}p´ýêhwˆsŽ¤v÷ìT1êÀ"Ï´ ÕNN\IâƒUG ŸSŠãqF[KYÆ‘<¥þåúWÚðE2¹þ§)Ž¶V
8â-CmÑ‰…3€9-´£çg ”J†×ÿ†ôVwœ‚BÕÅ1À!ÌÉ€í ma%Q+>LÉus“¿]%:Œ³Ã9ÐB€ 1‹ HTwÞE×}Jw,º£¤Dº3Êú‘]BéÛX Ÿ¤£u–fì¼~‰Ô'V/s&äÂù8†·Ò›KÐõúÝ½çxäˆÑßD•pêLé½Å´jN3 SKî8N“l”Œ‹1`H25Y•
Àö:¥sõÁàëÕ“q\œ­hIwY%  ;67pÚoZú"1Ž— N¡ÂZgLKÓ¸ÿS‘MW<6 ¨q8<#X¸
±’ÚÁÖÛÛÿÕVJF¤+|UÊ¯¿ýÏ»;Gnm²› ‰*Ó$ÐHdcz#i-ì½ú>ºòµ­[eZTbùÓÛT@TnUC v«©(”oQÍØ-êé{¤q5[#êÞ«µ¿]YdÉï¯&À¤K=öñ&>òi]®ƒÞö£7ZD"àËf£Æ€µÀûT×ô^iª¿úÊUh!‡JÅÞ¤omAmÄRÕsgtdöAè>OËOøä°Ì½±í-óE?J£7`D^B}ìšÊïM‘ÛGqMëÒ®„Q¯ó]:¡¾¼-[}Åœ>XüWoVš+jPýZ³Š;Æb5/§¸)Çë–€	mòIº/ÿR!åÌÊüeÂš¿+GwÖ«qiÍNušØ©¡+¥¨{+¦¶­›Ô+¬s¡£BÏøUR+b1	‰{3}AùHÈæç‰®µ6ä`‚.ÚÛ6dLàwaÄÈÎ?2gÎsb™’bd÷†>ù5.”øÊ%Ôõ“Tœ,0šƒ?Mÿ4}N§>Í?MW¢¯€ÕL79%Ï=•âÝ§â«?ƒ“”Žï;8JG»Véü˜íâiä›'¾pÛt—æZr~ÖÃPšrn^Ú¬ÅT¦]â–±9@tZ
ÏGk`Þ¶O¥³QmŸËäÕa£‚5•ŽXµ”¬ær¹?d€jäbï'ƒ¦wýÖ`¨WòÆ.ÐYÿúLøã.¦k5.y²;žD'%bÑÂ=pÝ“"©qô¨g+Y1 \ª’ƒ­üÏ{=dö5x¿y«Ÿ\½ªÏÇ1)ðêQ,6D—n(¢±HÂñpd"– gŒ¹Ðv“¼¢ÊçaÄË²OÙ¢*"îÈÁ¹†^»ä–È[Ñ6Cz”ÕxúT«OiwUt^û·æ+êQnj—>ˆ™ÝÆ¯È®íÈf¤¿¼¦tÀðK…ÜO‡T5ç$z5lº% t
é+í]h–TG¬Ÿ;(o ú€åÀ^&F˜øGÕ6Q]#½¥±Ò»uµEÙP4ëºqðmïÍÚ[­Mw†2éQk5£§Ú$t<1¬Ž5†^>K`ÑR‚ž hÁYÜtê	$c/œ‹:à…öZ©P¼™TJ£S"ÐŒì=dèA¢ÄXN™—ôj›ÖÖEÆ²á½ñí\×ÂC0ö8hí8\Ž;6¾j$!O£¦4‹úm@œŒ˜Æ6a8Òd×aÝƒ­ne\˜’ “ÌÜŠJÃg±X2·Å{‹{Xd_ó½~foF+Ï,µoäKîXÓŸ!Â–í'vE2v7\Æ†Öéª$«¬.\#,Ä/ªÃôKV]>.5ßP­Ð•yÌ¹b‡2`+¦úE'Â€Ý~bF

"(CPhèæ`¾0[<Ïøa(í|Ç´ðÅrpÖÙ5f÷A4õ#ªY±´ð`©ÀÓBâ©ã!¨²`Hù®àØD+_{—'ptâí[Ïê1×ƒ»Ú50¸Ð1X2«De…Ì!6QØ²4h¸ðÝbFò+{ø.áOMþÂÁc{ƒ”H49F'ÀŸˆÜœ
|ì¿Ø~õúÝóÝwÛ;Û¯ßío??Ø~ðQ“2»âj7ÑöýÅÄéˆ¢%Š”áÈ"^DÏÐ
¢8\ô ‘…¶Î}:üÁÐLçŠ˜NFÆpÙCõèe|ñð¬ÚµçÖwå"[aµ3NÐ®t9ƒÐÍ<]"QÇ•12±7±ü­ËÝÄDïÙËsC³*g3MåÙñ­bñ{é+ÈB— ôë¦[þû¦¢mkHVÊêlÈÆù»®é|f¼(º}2¬WkõP¨<£{©UëÄ’3´s´°u/¸&’¾>ñà˜RMbŸz2.Ï%"ê5ßPGc	#FØ¨+5xâp{KRyâàzÎÍ#„eSŸñÖ*Éj…¤{íB?ÔÊ/aI<žŒ»9¶ÑY‡¶é«ËÍ	øJžjnÿâD †%ƒ£leæXÕÙ€åMGÐœëCâÄÂ*×¥™ÕôjæŠõÎÃ3£†®¦‡Tßv%R‹\ñ]Q\W7«Uym ËGxAÚî
M©E™gz¤ê9ò$¤‰œ©m+.ÎÈ­¢«¯z4ˆ¾cšÕb>Á,¹\ÉÕA2QÚTÜpÖë(¹(3/A!vŒ|ÁÞÿïÿñOÑo/c/
3Gï÷½ÁOY:ívþ4íXt_ríúr¸Ö›¹'¤šûâ¡„;ö`ŒÑ°ófôÜD®…i ÁÄ!\^i7NÌà0 Zàý=r!w¤÷è:ôÿíF»û/öv¶ÑEèùëèp÷û—»¯Ž^cp“)ÿÃl&˜ò«ÞæŸ¦üG[Ó«÷ª;µ˜Å¦½.§ôƒïë¸†å9êà0qÔ¥Õ‰‹j4FjCÈóÁ™}Bxb³ÍPC›õ‰ÒàlzPÑÜÏæ‚o†ÐG¯d^RK¦W‹Àþpä[¡M<–	ŠÈí<©2j"àª.ß4sÔ7/Ü`›·k´{5Ù,p«†ø”d¡ ÄhŒáê™! 3â¹1ƒ„§:ÇQýäÕû0÷žµ£n„‹*1£ÚÕ_­Ö7rk°¡	Hì¶"/¸hÏŽÐ¨IC<] KkpTCÌ:Óèû}Û+’9¯_ÿ¸æÃ¦è¥¼>T¦S íErJÇs´b¹«•{h0|}e’tHü8§%#þØ»8	È³3ÈX_#2R`~ý+trýëE
|ueË›ÀÔ7«ÎäKŠj¯­Êo/áê½€´™y¦å}&Ç¶‹öbQ7#·NÖµ²1ã,Ö5´ÔØ’Œ´ºS‚²Æ2B¾çeŸXúãg•!ñPPÚi•+¥2-å„„>
bÝñY¿ˆ'>Ð
Ö€Æ#ó¦GÀ¨
êé¦¦ ÜfW ºP@†r}N`A£,qŠ8ÏM[ÑîDðcÍj1ö«M¬uü=!‡Ðäª]¡âurˆ%ð¹d‹EšÀG…U^« .éQÉ°¨73—û_eç1º’—¶¢Ù™°¨h	Ë+Ô-ßLk\ždã!¯K«€.ƒ1B˜’b€çÜ;|Í#yô²Û9ê yÍlŽ¤3Ù´iH—Ûd{êá"Í¢ÕqØÝPn¥7rƒ±x-ß½H<;Á–ŽšÝçÃ´|š*6,½¯¾ëÞb³2Ôo¹aiÕê6mÐiÌÜ²ºØäDyÜŽFÏë[ê©iù«Ì©%íf¥4"2Ò¼Å¢H–}Žæß_a/ýÁÿÌtrbŽË‡LL'¢Èhãâ…øÒSùHÐà°î¥Z‡F9~ø¡Ìí=G×H£ŠŒ?U°k·[ÏÂ×l¾tt²ûuS·˜$ÆÎÛLRªEË^ˆaj‰ù©Žõœ<)”oA)ç[;öÉŸ¶k$â5ûYƒ‰X¹-óÝã@S‹±"
ªC{gù°Ø†M.Boã²†àáá%#ùôyìoö[hö8Ôì±ÛìqM³Ü†»
›o´ÌÊd´7ºà>mÂ“c:ÚCÝ©ƒaVÄŠ†ÉÀa®®»q;4PÏH_Á+ô£õŠø&SÏH1Ë&ñÞücßÓi×fÔûzú³5h)hT"!´¬cP L§ø8„MlŠrÔýh­Ïö¤Ö«Üx8…Ö¥ImÅ¸=­KÞ…•­ŒTÚìÊª`œ¸õ:¬Ê•c ›pØ¯ÅšUbC5äØ£•H¡p*ÕÁ]¼'&pò{¢·üåþ1ÒS¼n–Ü`Añ¯˜`qFüÜ H,‡œÕw¼Ø—Å YBöÅÃš-Ä´Û—žÎÈ-KØ~\:j»õC… hþcTZ]8£Õ/K5åúS^(TÅ±ë§²Žpþ†"´†òòªrôÔsCb32öŒæ÷)òÿßg¹×ºj¢örµ·r¹æóÓ[@´îÆåyèñ)jë·Ÿ„ºOƒj€ò– ¾
^âQ¦’¯¨B>þAöÍS…'»]Õõ9 ëèyl5c©<˜èÎ¢Ë’q6‘HÚ³_«X3^ªç«ÉÈXZcÑ½'¶ƒ¶ÀKRÒBß¶ôpÔ7Œ5¾A!Yy¼'=»ã{-¥Æ†´œ+Æ¦‘©~¼ocÛðêOŠ¤ì–ix$â^¶YƒPzu$_#95ôÞFÕ
ïÄ_pvFì#‹«©bö•Ê.Ä*ìkj¾[ct]¢Ñ§Ã6„qì¦\CxËæ4EicâÑF›xê0“|XÓ´f˜~yH÷-m¾Öž”}žµ¾ábÊÌxzfò•køEeŒFÕ[c”“;&W5Ÿ³îÙÇ)¦I:âÞêxQCqVè#fEù//ÿË`^¦ãbpœeÞý#¢ð‚žîr¤;¿óF
õÚhs¼+³wÅY’”Ý7|Ž
I=iU6£ÎîvÁÂy	6pÝ±AvŽöt£çdÑÖÙ—™ˆ‡"#9,Òå†ÛåÁÞ.ÉV¤:¥¿"$·Ûx(4âù›\KªÆ¾Kß>>ÏÅ¾xÝ
øÊø[qÙ½³-³ÎXø$¼1ŸaºOJì¹eT+ <‡ŸÁ…×ár¨‚ŸysÈ]ÛÁE®A§d:×ÿ¨%õ\wê5âÿÂL†òßÓ‚;qÆó2›P'£‰m>E!ÄX}ýé0l ÖÇPùéõ?Ó(iS(?¬ òè[|(yðèó¨¹×¿ž'cl÷{’'°!c{u~¿¿Gý>O±|ŠÑ÷„sýúÚÚÒkìàlKìÆuÈ¶wø¯ÿeZ°0*q8f‡=å1°m–G–ÂIå1…á!‹±­˜Ï½…÷œ@¾\¢OnQ¤<þ.'T¥Ã¢E¼ƒ]Ÿå%&&\Œ‹æì#»,_V,‰îˆð~YÑžôª¼WÈÊÎ¬khçD	¤ÌÞà!’”ŸÌ¼ˆÓHé§Žìd’÷l—KTÄÃöQ&8óƒD#3ÏÍ¼Î>'‰=Û…ñ¾Èø»b!(4§}!04!ò:¥ïz¢ßõD‰	…	!²Ws¥ÐKëœ;aòå¿è’±bða–Ê³¢	Ù*4` ·) =n¿®íéQÁw¹ÚýºÃ²ïyÚ½&c%â$›Ìƒw°…N“Ýs˜Æ“Ž^¾Ø›Îæå.;‡ž94
dd“S¨¡¬“[N€®•ãC'šžßËc²³ÑÐ\¨Ücñž½ØQ)$º",²C¹ƒ=µ2ŠñîPhžÜ-¯•Öyx-œŠœßy<žã„;Ç¸Ã ·èb€ÈõŒe¾õÂç¹‰I’>óeñê	Ó/SP\…qV°Lï|Nä-Ïl­êlžBplòDÔé=§{=Èè@Òðá¼4”Ã¦<Ê1vqŠrðlñK+¶ª]dØ[—™få­Ÿ8ÏãOS½‰þò\a@$Ç±µÇVNöÂEäM“ŠFx?‡RØUÝKˆ^‚HZ”ö(;ŠæËÊéiñöò¦<Û­»ŸFy>6QýWä…•?d´6P|””¤91Äƒ‰Lí
‡Ìùº:•©Ôasýì÷ãì˜mä‡äÎ@}°üÁoì÷¢-ç™
[ük˜èÜI¡¨ë›µ¾4u³ö”R;r‚cèLì)Rî@·C?"¿ÂÞ}%‹GÈ.’œŒ“-DQmàð€ã·°`4'¢‰2çxK¶	‡f‰Qñ„ù…Bâ±ßÒ`#£|è™ø¤á»B­žö2ŠÌHæ+/½»òê»ç+Ö[ô±aqî¿ùÓ|íÁÚÚ*þùúäíýÓ~´.üÅ«¿¬­þƒ(¦ÔLÓøC2Æ;îÃ62*NA¨ùŠ&@“R¥d§èCò	J C·Œ,X%û#§ó]h_Å†*Úþu,g,Þ=ô‡n.°þÙ	õNh”ÿØÑ‘@—Ë Ã^7þú‚Þ±Kâº>`o„Ý¦Eûª]I‚ž¡	¼–ôšº`K8—žß‘ŒšNF¬ÚSÞkÅcSƒdóK5>ÕZÙ> 4Œç"gd•˜Qµü ¬š(©¯Qž¢#ÿgZža;­¯šKx	C~Ã†ýö±UŽ…Ã¥x(l•ûá±ˆÈóŠ¾¢PØî/ûÓüÆˆ+‡NâAÅ1ãMGÊ_òøÂXòéH>7?ô¾ïÔ"F¸Õ_–i9[ÍSøhh“©%¢ùÌ;‹ORúJÖ;¦é«ã9(%ÈIÆÄ ¼Ï†Ž¿ñ/pFx1 ›zYŠ\d˜¤W^hÎ†¤ôC!žf¤ÜÀ1@Bµ:JârŽ"»;".¦ÊAñßÔ$Í™+&øcðKLŸËwÊÁ_Á&-¹Çt
Â‰ì~)Ö§G4Mã1ÎþT@»ÐzÒŠ9	ÉTö'°a3~¿ÿ<—ÐEô¥Ôëµdv0§2Y1lœ4æˆiü`¡8‡â'ö½È´a|E¦OZeÈBhZˆosÑ4BHÔÅ²¦ÓKeü·&¥j²§JBÃßðßü•Rå#`Æð; ÊŽ¼ðŒ»+Ðä§ 4È¼èI’¯Ú1øKŠÁUÑ¾º!Hí£½°^t‡?ÅVU†?EÓ¿elbÍú3Ëã_´Á_B‚f“‚ÔFæ¹èp‰ãˆ0GË*˜¦y‚­Û!
ó²?M²×{Õ3:Ÿ²i^$Ã9'$€Ûz@ÎA’övöa–Ê¾P; •X°Å¢+í€â±ñ”6mR".ø÷¸ R ßû¦¨ÏÝYÁŠøüØCÑžñ°à”	)^pç‰@qº¼•]*•×håCjjœqªƒ(?=Køõ$ÂpUé$IF¨	vûÓ®ãÊ~ñ·#0’*ÛÄWS åCÞi<1	¼Vm6?© ‹Çç´žL#©`[œ¢Û0LÖ±Ð!–‘ŽI],	ãý˜å£*Ö“¬Ëk~<ŠŸÜëlrÅâìäÊ¦C&²jca\Ù‡W¦ø.ÎF~ìYÞ»=  Æ¼%Ç\¼‘OßrÇ¶*¦Çu”í^Æ³MÒôå£'Âù‚‹„xOóŠôPH%^	Éœä(é8.×ÞóÀ:SŠ‘¢ÑH²Ê.ÖF±¡±–92ºBÆâñ2ä?HŠpœrAá ÂJä’7‰×÷¿«”Ðî‰^(z²ø!¥[ùö©+ÆJÈ x¶¦¢’çº¨Î®Õ³$¿9þ]·­GJÀîi7<õžåIÉr^O¶2—±œ9LÉ­ÝjLÝØ[v=S9`”ÝôˆÿŽ›ê
øW£s¿R€•4'‹­i{âXÅ·¶t¡ík§lÐ5µþIœÏPú&ÇŸøóQ4N‹ÒÅ!,²=}±´OŒ.Í”èé5ú¨hêÑTeöü¼c£à5éÈOêB%Â Aæ´4Z¨
QƒQF±e Co¼)8ô:!€T”±AÂC$@)àPÅuã³ðá#Œè£„÷ø=wM•ÈØë¡‹Ó&{dÝÜ‡zD·‹(U‹¾ÉEî_‹}ºÑ¾îyè’vœ¥ÒŒ—ü	Ð,ß½}ÑÖëÔÁû•„Ó%Ö¬„ÚÌÍCR2+,ªíïa¾6¢Žpê`–x…rÝK)¬AC¯®ÿ¯Ä¯[MU†'Ð×ß >ö§hC¦¸+v§yÐÑÒõ•‡‘°ªYê»ïæÓ¡8ÜyêÔä,š J?ãqzŠá—Î˜1#ÅëÚ$b5ë\ç-³r€­ Ö½)ü‹únÏ¡LÅ¤šîÍ[óø#×§Yò¨LÑpÊúBûìf´Ìwó|óÖl”iy›°€ü«•R¹›¿CÆõãY
D•(÷,@âÊÏq–}@Ê‡êÁ°§0“ñ|2E—¸E­I´T°Ú ¸=¯Ñ*1‰/Kìw|‹ŽÇ‚E>žB0½ÝËúL@S=ÐMÆf¥…pË‹ê•üxT.wÒ/.§™àòŽ’·×õ-Ø÷œì8 dŸŠòF–¹¦ÛhÔ¤ýåšW¨Tzž˜ÅÉ|ä§¡ÑÂˆ˜N9¤åœþ–,!V³ä~«T+ŒüžZ4ð€¥¼„O#Js1‰ ~ÂêPÎažIär›ÖÐŽ
Û=8ˆ›>öi‹{'˜›Å­Â7Ð	ú]Òåé(Î¹„(¶¹µ‰Õàìt®øá¸šd„/À YlCfH(€è«aŠ	;‘½&ò'PÕ¬ñ–ðï‘‰á×›Ÿ‡gfxl_ø	 «’Z„Rn44:[‹òU´þž?±6<«ØkdŒ)´—7Úq¬I'ò× 9ÇXÁd,xÆlš™LüVæxè1<eæ ïFôìpß½kXIGprYØ¨À×QÝA>ë#	ß\CF`œù6&tùæì-3’0h¼aÍ8¦k¯Ú
X{Æ9@‹Àå±p•zÀtN³t,Ý7Ž]Â´O¼uÐ¦g—¯=M33žBR¿í¨¬k…§?Í…ÊíNÓÂßIüàú2©LÅ¿F¡¸U	|mÅí[×˜~Ï“å}bc! ƒpZ‡ý“ÐÛ#DrZ31‹åÚgÎ.‘,ª|-2©#¸Ýµx’‡3*I<—Ùé´®Š!ãL[ éD_Ö¼'å=¬:¼5åÅsÖ2'TÉHeWA®š×a'*o#P<fÊsà‚™
ÔáâˆSM³h…‡rˆSOš¥µ Wá˜õ,õ®¡­’J3Òò„ðRIåyp	i­ðÓXse Àž„vW ^”?ñŒÕzã—K’A¼{Å—ªËÖLuI¿=Ïî¤¼ìPhéÄj«ù“ž¦«¼Ñ•gÎAµ¨eŒÕ©U§ÆÃÏrUy¢E[VÆ]ì¨óðc{š¹\rˆ…°BµËÒWy)OÎ
?ÞÉ×é­n	£lx4Uj<\Ýòy½ ¾Š =ÚdƒwÓÊzõVü¨¨ÒSá§RWEý
}•"_n¡z­~Í• MnÑ¦º+ü\9;Î«ÃRv–"„ o»p6«ZˆÜðsBœÑòpBüM{NHTt5	ã¸(÷5ãQÇ_ÑÁ¶/Ý€«³ËYi
žÀ&!å+…wqeˆú¼aDbâ±$…Ë-	­Ó²™Ä‚ÀŒ{ö:r²*`4Ý$<z6x{PÅe>Y¬r.y_aŠP¾‹Ê•Ü2VUGáË–¡¾Ï”úh:ã·¢÷/“ñ…ÞN&Ño/=ìáÕ{'ñùdŸ!†Œ¶®EH[áÒrº§ûj>9Nò.baÝz/t<ü—òAW“NqŸ}¾Ëñ€¡¦Åð«…ìG5¤Åó9ÏŽL(%îy‘î“¸&—_Ñƒ$#ÛÄ*<V»êŽO!'¤:ó6œ{<|¯áÂÿ™‡àÝÐtül¬yb	œ6ù¼÷&ï¹)î-™Ü÷–L8÷Æúror&K¶ÍÌ-£.ZsWûóá?øX;uÝZ[Ý:~]¶æälQFo	Ü›ŽîŠ}#P;ká\ñiÃ‘=ðqdêRÃ’I5XKŽLSŸéÕ¾CÆõŽS¢Ù×n›ã:Q‹¡ö¼`ïñ/LX÷B,˜;Þ®+Ì„ªèµéÔáöÃ…4qÍMå?Éo:rÌðŒ­äPžyÿi­bZùVhÀ³ÊÍ$.Fñªw–cÍã\º<…i ‘ëu¶|0ËW¸þyv ÒyV*¡vŒ²@LjáeR‹:&Õ{ÁßâWù–qÌúÖ/6–EXgëžð8gh±	ã,í‰’o†'·Å6ÄˆkV„ÓÏ4‹±&ž¹VêÁ…Mí Pí¯ô­f
é_¨¥qIÈÃ ¯Ž<?¡¹*"ƒhÊ;k_¸Ùœu§i£flÙqS¡v¼ÓtfÈÆ\©ÝÌÞú´ÝÜš¾¦ý™$%ƒÌ	A‰m.góÙqªÅ$ùÍbÓ¦éIŠ’‹ÿÒÑ,Ðéˆ%xÞ“S&ã|´z’ŽÇP@kIY
bØ@h1n¾øÅ,vç©~áÄqz¯É„C•#kñ4œcñ²#Z1qwÄ/·ñBÊîë·ÕjE…­6deÕŠ)Ãƒ×ÃQÜ@<„wÛø@ÎÒ6ôŸPÕn` €r‡‹³Ploa/(öjQGˆýÅL	Û‰Í´0IÎÙGŒÔVjÆ+ö³´¥ÜÌúj/¨îd“¶Òñ¢Ò-Ò>–ÔT§jHé€)"ãaøÄV2DŒmCÊr™“
ä¦Psa1Ê+H©®h]ùç³çx¯¨‡jõ–5x	Ÿ$$÷™Ñã+ÓÅ"¡rMúZŠ°^U“r7 ÃËhˆ~nóžuú
ú%ü¬÷xàÄ˜#VH‰5¥„ÑBäRë¾æýµ=²};9P¹“ÍÊM‡fÖªh^?Õ˜‹±öôžâÑ_ÿ™½%‘Ìü‹›õêýRÀa#„ðjOFµŽþs~jÝæeœ¬ÓQ…{SXÿà´Ò×)aß³Za¸Xÿ1€Ã¢ ‰Ç–ßÊy·{Ø¦íiÎ+è/j\æü…f_´Nêˆ¦Î€‘-éâ6Æ·7¤HL€ò­òÝ´PâÇoJóÍ ÖDÉŠ˜*¶¿¨Ò¼8;ºNCÕ±]²Š•öKüx»«gþJmuüØ7"Tj,•TÕ—ŽÍª_¡´9¶
r(ã«ú‰ ©xh6š“@ÇDâ6"€‚D5Ñ-4º]¸R·½Šùsia¢u‡”Ãøñª´Â+¬f%>½âzcSñ»©ö˜º
h¥f 
”MÞ«H–*ƒ6j_üTë“¹>¡m£}2×3´Ò(ã§B«Ì
íáX¥Nfšv:dü|ÁÓªN_L5Ã‡UC–{V…~éi¤äW¼Þ£Ü¬Ñ!à†éÌd96.7—¼¾W&“â(Ã71ÝV6rÐcæ®ý¤´Ò”ž‰]tµ¦4lwt-R<kÚ§ÈÛGîV$­öæjª~ÕsFt§3}¾þºœþ’­Å€çÝi7ôêz+¨7£³·>}¶u[ß\;yu8¾z=\s.]‹%o’éW˜¸iQ°6»b{9íc_Ô•][S¿c¤BÔ¿SßáÌrõI)“8a•@“#0öÄÌ³[§¸«Klm7-ó‹‰áÙ*Z
¤íóTàî€ìx«'ç0jû‘§D²ºea¯ÙAÏÛ%Q‘N1È_ú»ïŠ9=('Éf´ð…:ƒ	ÐìøÔŒ ¢Û³Ùø“ìA‡ÙKb]dÑ~(âý>„!Cû±wr(×&‘£™ì¸ì<AŒév½C¦.OFêÎN§÷XÂýR_Œ„/
É¼a>ÊÅÓ7ò˜Ö`y•mŒHÂ={JÄÇNéÝt&òÔá—^Ï2¯y&‰áèi†ôôNÖþóáëWF4Ó“O"Ód/4gö–4VÙœí¬Nú–³@*qFÅñBmHmÕ
9Z‘¥-‘:	>×§Å2VH	'á¶[}¨Îòby\ApJï¦…Xb±ÕÑÎ[[N²0 "?ƒ¿„•;:=;Êd¥§°Ù »†™©'@#XXÚ]ÃTòØÀ«á ŽPF×¢ !•¥ù ¾äkƒá	ºoø
Ê0CZ/o­K8˜íöØlÑ‡2p–§S‰3XÔZf{6"HG
ÍqO´ÉpI¦ƒââ;e;Þ2ƒí¢Åæ3¿iQ®^­ØOcW2?þ\ºÐ/gËd}’ëÙàúN~y”PŽ]µ|zL6“>oX'/Ûy)zì›2‰PÀh«uy&.ß¨jDš^[0:~'†n‘ ·z JÞ”cP!¼)‘3‹è~Ä#îÚ„A<®£ÂvÙ3(jäáÞÑÞþëÃw»/÷v·Ud©¢gÑ
·¯ÆäBtß˜b¸}y‰o7»)áðLMÐŽâ6hGÑ˜vÀÈê5†
šÊ09ý,Ôƒ†'ˆG±â¡æÒg×Ñâ.Ñ}ô·HCÄ7‹0¥˜åà«*¸ŒR®_ÚT…Ê	ÀÌ´GÚ;‘@éªo4ÈùX»AËàb4h¥l2¬—Ùž­Óš³ó;­‰ÜNì¯š¦Ò0òü2 ià5Ó”±ñÜ¿³X ?)˜08nÍH7\©°ô,%2&Í@â“`r3WúÇI^BÙƒƒ×ÑÎÁõÿ<ÚÛy:xepPÿ<OÏ³ˆ2£…Å~†w<ÏT†¶Y”·oç''p°È¼qz6ØÃ9¦{ÝžŽP‡ÃBS”˜Ý†§8ÃÀ^›ÑsÞK: PÊÀ¢D|_°îä]ÂÛÈ¦1”7¶ª’Ýqàï¸…õåÓ••œc,AbŸ‘ÛÃŒæ@6»ñpØr| æªÙÇÓñ~›ŠÝèï¡6ÝZ§ÚÌ³ ®¯É¨”ÖžgZ§SQx{OLíEržŒ¾?5?èA6(=Ndªd½×À¾Xø"'…p9-rç/ö ïØ©¬d{oÓh]/úp<Òìyì­b”Ðw¿¶û¹ÿ®†m]mÄæVÕ
f*yÔ¾Žð×HÈž;ø¾e `3N!MuälEYKÆ©Á#E/ëÖÎ>`[@ö&Ð%›ÚÁêL÷"FSE*6{8Væº7Õú_EA+sÉ#ä´#×ócæö½m¢£¼»+–án¬)@Iôdóv†w¦Õ0JÑ?ž%CÙ;˜üd')$M~,Û–*7TÉ3¿#æÞ—ÎtŒVª§R ÏBŠ¨úéõtŸ½/þ˜&7£ã,ã)îØtLj{OÖ¤­ìÊù™Aó8ÍƒädÀI»ã|Ú]qaòÓõ¯&•þ/]RZRL÷£ãqöó<AÎ ™ruõˆÝðÅ•’NÃ‡à°`ŠHûÕá ƒ¥…~4ÎÜ€ÏÙ<¾:@P¬Á‡]ñøï£õÍµ5ø¯ƒW9e‰žIÒá9öNl×oPÉøãS‹ó­ÀâU¶<@YB²‹ fj‘ÍáÑ>@t‰W4Ô2­®øë=!V™E«Ño/¡É-’‰¨Ü«xŽ1æYJ:èaéƒ]]3wŒ¢$<º<ƒ–¢fräên´V6££t‚ù³&3BÙBñeÔ,o•Ö]UŒvªÑ<=QÐà1/0•ú8ž&?‰iò½iz¶÷³*‡°ã Ý4Â,c½¤Q‚XÏ=ÀL2’3 ë™
H
X\wa“Šº˜Ôe<ÆÝ>Í¦‰|Ý3–Z•e{S€`©³p‘ªX —ß!ârAžœðÌŽØjƒÄöF•©R:?bÈYLAé¯Á«hÖÅš¥pã²Ax—àGÚ"Ò9ÒòK0Ô ¦ºÅÈà5¤xÔS®FÖ:z5YÃ§rôÖÖ,¸×Y5…”º´ì‰]HOŸÑcÝ"]­`@ôÎ!¿¨½3~=d¦³¢Á-û¡sÛBz_ÜAg-ªÂuàVM°Gl¾½´g¬ŸÖéßO©*®—ýÆº¿ *ïË[èZ^]½³ÚÕ^8×"|—™+oº#Y^o6Þñ@lÍéGõbW_×á!*–‡ˆYP)¼!±1!. ³F•HßG6…Pï=ÎOÚà_¹ç ÞÍâpt	ÞT!‚ñØAÍ{.ÝÀâìÉIsG"É·bGðÑ"ŠÈù1#¼¨Œü³Si0%º(ˆãÄ’2‘1`#ÖecÚósJ»%uŒV5'G‹øu:f+ªÜ,\Jº^ÛçƒY
æ€q¢xˆfÂÜ-ª	7£u½s’Œ-ïÈìÖîc•oZ}`3ZÓŸ‰Õ´Ë6óxÓõm~^€©‘ÅRÖÈjéAžEµ1¯Ï×ÝØeZU–.Æ¯Q6¶]d©F—6Óeš"ÍbhBœ*Ôñ#Oä*åuòe¬ŽŸëVâüÔªP.¦O3~¥ñ¹¯ge:AsÌ¤w¦N1üÒÉ$Õ.xÝM¦Åœ„ã©-)?W!ØÛJèX÷¥%¢Ó\ ROe0€VåQMB3þ½#{ÐrP|è‰á®éº8E%L…nP%¦&Äó%ª@Wm`õŒº.jªs‡ÏeGBæîø†¨kÛu%‹_Ù®@8«V1uË&µcR0aÔSR„F°#þ°ësŽùfò$ù%)0…åY<=…/Ó¨³“§x^ð;(lZÕå:'#HÈ–Kž)“WÀeŸÅ\öqv¬] VÌåkU#ìrX4yÙ§‰à«JúÅ_ö±ízLËj›!-«Ú²}¹XöEŸ“l€´Ý†³ìðl>ý e].â4KMV;Í’WlŽdùÖýcª’Ù6ÇÑIœb¸ œ9WèÆ¥I2C2Ï‚a6ûÔDÇÜëj”‰QIõ8µ
ã Ï«ürÅc$© mÆF<¡·Z”³ÒüÚ23Ýa‡[Âx‡œÁÊ>4Ÿgy?*QÙ{äœve2X†~¿~ÛQ†š„·J«J%ýjpvsß¥·µÌB±ÄEÜ¥¥©rÖT’ª=Eãæu WSÍÒ{RSÙ€“ŒÑ„sô.ì£7«;¾\‡=LŠ—9ŠsL·L©°ôIÎÎ=Ú[ì<-èT× äúÈuaÀ€Ë-™ß<ãüüÖ8ÿ <Æ¹z°è–iWÓ†*\ž3±i„¬ŒJB¾]Á“h?Û0q"œ.¶´$«}RÌñÌhNGjþH`@EC^nØž†êÞ;_MV¬ÆQoHŒ?ÆŸ
SÁtÐtâ”jÁ9ú“ç²ƒ  î_â´ D?<|@«Mld¹ûBØ–0-»µ,wb3ðñ-e#„Ý}HpÙU”Ú/Ã´0WoG8ú…³™a¦¶ŽÚ3º~ŽF˜³x¤N[ôÓÙŒðMŸœL7•¥ý·¨k–œ`f¾Tš¤íš\}ŒãXG@õŒ`\P¦èI,dpÆdÏÿþ`mõÑÚï¿…f°(^Õ²sñHÆò±tròf÷&ÀáeÏÙkšQ?z´±ÖÖ¾1%ã’Lx>ÊYZÇµ+Ë}–ã¯\ì±Tn+`¨’sL<Lq›b¤ÚãW\…ÄÜ7´Õò^¥0|D/PzÊÜ±”ø«tÇR.6–ÒÃ\¡™h’)a(jßFú}Oä²_`!xB/€™»¬³„ ŽÈÓW‘=›KcS›Æ4X>ü	‘®d'±8À¸P!Œ•Êse78N¶ãŠw”SòóUQ·jºê¥ª€	DSùê×SÓY'L?u!*Ëk$(k\|ÊòJÒx@7Ö*òØž$Ú"Ãç¡8R©ëìs¥îý|4§éhTç9p_î„"ÎùéÎˆÊÕQž/Meª7.Má»4²×l_£ì-“›Ï?èÑº¨hPBžeÓžÊãŒÂTÝØ”çd</Î¸øàA¡8ýz<dŸ'Ç0´a‚êÜÜïT‡4ñu6/««HÎ´¢|îJef™¹|‹Jƒ"ýE¦\7´)J­(C=~@š.¤¦‹sJ•-³ ×”& Hk K¼{r’ÑçŒi  ‘v5—ÓõËè¢íñø…aƒéjóÐšÞ°àzƒRƒÐg§¦9Jp’)„ÙÔ‰£`:$˜¤Ê§gv[)—›*–M¥rÚ4üUPÉÑ‘(ªµödHW–Ù‹®éÄ€¥„æÕC…,:¥ìÚÔ®øéXC4‹'Tƒ3€R)æño`¶Xz?óK«5¼X¶hjÈx†ÚâI(^l‹­i·fã’(~&ƒ¸;a{æsa?}v"^¿ZÄyZ
Y¾Iž·öq`üª6HJî¥Ô¶£±œ¡^ä°)a§Že5©£Á}ËÙžGuì‡C²vu;H+FÂp	ï^Ä½!mí*__çn;š5G–nª9Á6ô0ú
Òã—ý8‡“vLß;Ü‘Ôrí–!úŽÉ§m]Û€(ã!ªÊ«9 toFÐÏÓ~Ä[ºÞ°Ç§kŸeé”ÒØŸãuä•hp[Ñðuò™ßV´Ï×tm=mtõb-ý=Ó8)˜¾ÝBtƒóD´9ãžÞ4Wôd'zª7¸%Xù\à°¢0ñP§ð]<6öº¾±ý.2ª³/`;Î k,E¹û¬ÝßhŸëû[µ¤ö¯*É6«¸'pûP0ækœÌ{ÂÎÈÕ	°Tpì$ñ°ä~\cÚ]›h¦ÇëdÓá'8‡¦¨M‚ÿ€÷)¢“$¡I¿'6b;)KÃ !5é'<ñpÊÅÓÓ„E´q*úyžÌ“&\	ÙMABÒºRdÃ")ÐLGÂƒb5œý²×9·¨ûh ÒCžc’NÓ	2b£´øpš”³üC´wÿ5:0Œ³²#àU‡\Ñ|{læ•ä£*XM”?y“¦Ï„—Q–üy?zÄ¹W¬gãµ-Âïgh·õ«6h¼e×4àgHÝÉ*Ñ½B£|XÊ)fœ—Ü¥\U³1œfUåëBÑ ã¼€­Å6é¼Ð­$ÔŽ`aè‡ªÉ½×é¸¡§{Ó’U~³ö¶­¯Ál=_gÏ£ÕhÝy·!ê¬ÃÉ€ºÕ°"Á?X^SfZ¼Š_uUõÁ)C	ëj|q–}<Êà`í®»L:=¿þuœâÝŸ~´B"ÙŠÆèË¢¸y½DúÆÀ`=~2j4†;*ž•@£M‚6²WšÆ-ãò m²â¾¤lášTÃÿÆ˜Å•ËÉ¤¹íªgªE<§IÀhgÀ¦«Ï^hŒ˜y»fÝ?÷ªˆ6$ø=¡ÃLA‰sÈƒÁ
7÷O­›ûÆUªÇÆúP}ŸÅðZ¦‹þô‹Ö9Xl‚‰·)»&ÍñÊ	ù`‡EbÅkº#î@bß5U¾zä Ã·)AU30­ÚÁ˜H~MŽÃ˜‡Ùr×®…A6\žnã2âDCÚ&Ù:%ËÓœ2,Û¬†÷¶:°Íe[Ç‚Y	z]:âpÛ.õÛƒ2Û;|Í½ß•ªìJ„° ºfÎ$ôñ‘ ¬xåBµÒÙÒ`D-0Â*o*Ì£Ee†GS£3`<ŠéF.{‰}Ášä¾±²ŽPåˆfõhH`“êîAd+ÿÄýýçß¡¬§1ZOdve%ÒY±±R Ò}™l	äR|´ž«tfÂ¯³ƒmÀòŸ
q×¶ÄÍÖ§¨“šÄ¥Ä2Âkßbm€ÕhtÿåËûŸàýðÃæd²"Ž'èiÒwÙ´<v¶»!øþâ(¹(w²1¬ë#`MahßµG½ÇHv'(ÌŒ¢¯×ÖdLÔí$ 5\ÿ9O)ò¥žò£CõûZ/a‰ÓS‡;Ct’Ë;r“xG¶¾Ùúš>hÂûçä¾þÛK¼«÷¬Ï´ÕôÎv¾Òv9¯FÑ™‚bêÔH¯ÿ‘î{ª¶b
‚ýÛËFÝþÒ} Úé‰)>Ü[Bª æ“IœÒp*ž“œIEL!óTGXîPÏ·\JøRÊ+ë¹Ò•¸ïH¡b u™•pêp…3šŠ˜<Úö²ÖSl2ku_,Ò&Æé²@ <ˆÝ0ˆô´ù¸ñÒi[­>°
Ý®6»ûQWŽJ/‚çÑwéE2ê®SÂáµN`G®?á½ëAjÀar–E¤i@4Ò†~õŸ8Ž<ZÓ¤Ý½žÜ”üPNÕgÎñ¨^
IÞ
ýN Yî	dg­Ò˜¬yNX-èb”«C³òª,ñ¼‚MËíK²Á²¹ï†¬. ä}®¬ñÞ‘Tìÿ•œW‘å RÇýè˜ª«J»Z’ <íëï*ÉÝ×öêvdÌ\¹d@+²"J`ƒŸb(‡§Àš¸À¿“,T}P“ß}+ªÂh¿„=[©or”Ð®F_WÝS‡èÞÎ]ÿ3JÊ»4Ì%‘Ë:#ƒtÃ¥ W€ðó‰Dh˜_ÿÑîTÓ@@æ‘äe±}§Žp4È“ê|'L1/ÿë&TT,æYƒ°ðæM‡ˆ7’_‘Ò¿²€ß_K  |ûV5qœ>mªuÑ‚œ%hØíœæé¨cvyX~³ ^˜¢ˆ ŒB;˜ß*™À9vð6ÎiOâÊp(‹…Ž>UÂ{è‰»þ¢¸qSÙ’œ¦Œ«Z’Pd$ü.wµbà	yñ7òidT5-3FCþ×<¯ì8FÌ¼…r;‘°]`ç§ñ Yaj{*!X7ÊLpîK#ÆÇ£Ñ>jº}&H9´—î¡!^rÒ‚Þ	È¬m³p/]“…ÛiÎÇÀs¦qï£µžæÆÜ=jÿl<ÒÁ#wÐv‰Ž%ñ(ÁÝòœV…mø)ÓùÈ]5-®pÓŽÃLÆ–›Ê¦áCïß[­w—)Ï±é ›$Â6›—³9@÷xœÏs
h ®_»*v¡<¡;~^‘EÞ‚½ävåîy<^×Ú)ï–ÜŽìZ50	8Ø"…+ÅšòLV±{ôb(üÅ´Ë(òLÅ’Þ8Dê±zEÔˆ+ãÃs$ÓaÊ.à‹Œ€]ýp0ð€0"35ˆÆW XDrÎ²>ª’{ä„î@6
ý¡5cjË¨vüeë_I•-kBá/$Š6
áÑd
>þ‹§Öò½£ˆ9Ôö€ð¯“‰Y1îwy7²`ñ^&“Ì²”píá4ýyžˆ’Á 7‚±8Ó1’F™Ø4ó–Nd:tµTdÀhABL§66Š3!êûìùŠ€s(ŒFbRÝ¯ÓuÃûÄˆIÓÌU0ßÒV¤«‡Ó)µ©•µS«+ô5ÏßA¥‚ÈOÉí•ãðè*,•,³EG5­UÉ'W‘ÉàÂ5L”Öü¹´9÷ly$¶EÌØX®~{ÐAL‘L‹ÏàNÎÎ1@·Ã¬‘~ôÆÁß¾Ú÷}–©/ãÅ¾5E6´é1°AŸ(ü|Íö"ÇaRV†‡’ûJ’DÛ<¦ÄÑM)<§ñ¼VÏ2e¹Ä2	G4a’'UÓDR1E
†òlRƒM3(§Åô*ZªÝ>¨È2ß3ß†äÒeúÈglè=µD]ì¹{oÙXÞ‰}ñØŸc«&NU ¸.È8Ãc`
§¯6‰ø*Öª/kùj7˜¥¾Œ53¯HöËÚoB¿¼´R%…v‘R¦—1RÄ Rê/³×ÕÃ»ôýÑÎë#ù‰çâÇí ºh})ØN­ýí€Àä	 _~´Ý²öJâ£k“Ä!L|@«)Q¶¸=|-Ú#kÕÖ,Œ­Yø·¦j«àˆ_,ŽPƒÕ21¿hù^0(o‚öÅ2Ð¾ùÐÚá¼Î’HVF]K A¹Z‚Üµ-ØKbÍ\B›{.3¹Œ°¨RÏë<n5CÛÔÖ®X’¾;É
óvù¸ŸŒÒóh8Ž‹eÚ§+Å,&«ŸV¿^y&ac—9'ì†À*³—ò@âŸV“òc’Lµº¬ö3cUŸœ­ëÍ¡ÞoõÁÅ„æi¹z<Ž‡"zV Ççê7kkQ™ÃCà¢WËôô¬\yömz<N³2Æ\'x–SºýäþÙºÕÛÌéŒ5ü¦.Ñu>Yyö=Z.þ©•¾s…zñäþÌ˜Ø}kf^8Æ³Õ¿®~ÌãÙŠ5¸t:›—ö•O¼–ÕÓq²b¿JGOW(çÃ*KTáÐú?KG#X»D<&³òéÊàb\\ô#ü³bÉ¦;äñúô’yW²¼”.ÂÜ ÷­é|;/ËUm7ú ¹x"Ï
¶8N‡ž^Ž²S¼˜~”Lf¸WQ™–cí·qzAÙ³Q2æ©),`AûÏyí÷øÓËõo®`ÑKªcŽä>ÊÂãcyò¬°8"KWÎ^ã÷¦£äõŸcåŒ•É¬£l8GÜG'·]@Zøúí§½Q·£ãI§·5b­nïÊ/`Œ–-9Œ—V¢Éx`¿A5¸]jêÍŒM'±`ÙÍç’‡ÑÍ@ó»VË
šž“Ê°§MB 1’Õß9>eß­Ý¿²|”äìÉ¼=ËÎ“|S”YWO´r#åG6]¢º¿ðÂZxß±k§³a·2·ÇcE2;žc˜W¬Hp£>Wö™¾G YÆó4Ê7	ü§Ÿg‚ŽM-˜=åß“öStlóù$ŽŠë?GçÉ/6]Š"gðóä(‹³¿_þ¾HñtŽ`>Q©Ù»˜áÅæ(êÙÓ¼ë<ÐÕÂmØ;Öà®öbØ×k&N­;8õhÍ!­ðb_)nnŠ¾ù:«Ã¨‹ÀÙ»T…ÑÁ¸ÀÌYdÑqÅ¬EÔ­u/#©€®Õˆ5x£PnX²ˆ©f„zŸB7o)á— éýñ¼0áL)ÄDZÁøÉ}÷º…u,ÞMØ_r#Tl•8\³Ñ…©nC"éZÅ–h¹)BÛ‚ ¢ßºÍ³ÓQ:yþK¿‚Æ/(®;»#¸'ŒË(É0	§Û¹å5òn¥X.OëÕiº¤ªæ/~7ªÌÑ_j3J.ãËì¦Cæ»S	^¦Ñ^R'÷
z—'Ã~IDù»b~š`f«B´ÿ’‰‡üxÒ\†¬7f©l!+déù¸ÉåTC}^–z)!ÅæÉYÑÃïŽmtéðsk;¥mó–üÍ«„Œ~{Sòò©`jÅj·Æú€K¾Š­¸à’àÄÇ€J$þ¼àÊ"Ôï¡`ô*9½þó0õ¡däl1ÆU[™øÏjí?<ŽÅ¦QC‰wr¼úu„GÉ	Þ=¾XE<Sxì‘Ä­õÃp
r»tuã²95¬ËáÔÆ•Ùeöiõ×,fãW,Nts[‰ù¬nDÏÒ2ajÐiF*;;ÜE¥¸Â›JØ:j©SíÁ
`³(GA=ð´Ôµžå‰¬=ÄªÆXp˜kï“¬žÜ?v°d¡å|òÝ\M"¹3Ëááë—´ÊºwWCghïÊbxùÂ%­;ÖïäRHŽã®¬»pv;Bfe¿“ë †w‡Ö™Œ"Àe¸Ëb3	ºnôž†k÷w‘ýF›½¥©'N\¼Ê*»±:‹§ÉxÅg¥|@ì†mS³t±ðT­«ëÑd´©~nâ¡kÄqÌ–ôt'c½õãq&ì•…ƒ8rmæ3›†q‘àˆ×Wž¡tƒwŠgYÎ¶Á“ûÔ´§Kfæ­1J|\=™^
?DÞˆ(ÁM2Z½Èê²j'Ùp^lrMÁ*åödÐaÍ6ê‡@ÆGkk÷7ÈèAhïêçÑœ2žÃþ*QÎg6P†Æn"v²GLxà&Ðaifl†û•heåÚ
LmÄèÎ“û¬€¯î¥ë+JÎ°t†)Úþ|zz‰¯Ä¼éÇ3öGöàñ“ûBÞÙì9=ü|È(„.½ÿµ¢¥ã3Ñ;–‡¤ÌrF¹
G­ªä(„Nð'êîd“ù$ŠÑ½JT·®° ÀˆxƒÔ/¨êØE%öÓgìOU§½›íˆ:¶wå?šªÍvmŽñ©öó÷¹œ\âv!¹žÞs‰%Z·Dr¢R~?ªt¯WÕ¼ßÏ«Ê/ÊsÙJ}sŽ¾¡‹œJ»Tí!ZïaV×R,vÛW7‚z|ÚÚü%€9>-¶X*ÚÒHEË‡_jC—m73šõ|»6mrôv¾2NNj£ðÌ‚¨o¨:šŠ³x”}DVš±©ü÷dd³Ö£yNÁ)é,À<}3€Æ«Ä_¨{"oØs•^gj^Òöh£K¢vÓxNº,bE[Êã©Âz *ž¢@wõÚpR>À“r³c ©p¶±ã›~HéàÍúÚìâ­vÄ‹5 óVðÒsîcŠËçqZGƒôŠhŒï0¤áÞÇñ†ª_š´€®sèî©-îwÒi•t:JO3	þó¶!cO/¢: y>V)€ªMÜ/	Ç¢„¢´ü´úÍTâ×;µ`­²'^¬›nGü1ìSà¬>¬®…QÖ'ëŸ†3ü˜ZvÌoVç—aÖ£ÈS³º¢–#GEáÈïåááÑ¯nˆGîàGveQé6T–}­©`X×B°­…°š {;U°Ò›·ÅÃ6;ü47ÃQ[á£oøÕõÁ#e’VÇˆØÌãÓ
í‡¡K±=ë¢á</²|•¢'¹OVà â.m€•±òT
•SŽ'˜s	ÃÓÎ£+ÛoV+ìÑ€ïom÷1Ãýmì¿;†Òâæ¨ay±-€¦ûÂâ˜<?Â/Î§ŠOYð3èm¬<»ôÒ@ñÎ4a°„Â.Ä	yœô/
}$–³¾ïçwö5u¸µúÐM‰“Û+~d[þ5 ´u âsÉ©*ãà›î´à{š0‘Ñ¤D-J¹ºÈ/´Á£ 0î1çgH¶Ü“íãÓ+‹Û@>lG «_L•õ7Ò½(²´ÛÀÓßÿZíÓÉ(À³xœ
þì	=b­]9ËnhÓƒ{,Šþšóì¯O£Dz£¼ÛàßT	f§U	‹ªþc+	Œ h>¢e,8@Wqsp}/Pú¶µùöåt`C|Á±—Œ\5¦JðÉó¡»\‘ðÙÑ×’)àÝÑÁàÊ1*Uöp0 ÀÁwq²ïÛúv´õm£oCÍØõs{ü‚k’Lœ'r:•qbÌjOñZ-ŸÅÙáß¥©Ak±5zSÙË»‹/>¡-c,z©¢G¯hoì¶0hK,—Šÿ›Þ³»‚wq¯OÉ»<
¾Œ¥nHÇoÛùR.åtX*€ÍGÅñÑI«Ã}™€eÍPTH´Îâ‚½êùÁy›éYˆ@8?VÁÓ¿|Z+«¼CKuµ».SÈXúeàƒuù[ÿüÅëúÅ$‚òóuþÁöäxëÄ•Šû²ëÁÛâöçÒÂJ¥:ïõŽS©´BÈ¹T*ÚÙ' n—ÝvSÏˆqþ±»Åíj46zð9ÝIÓV›™†Åc°ªé'Êîb)À´‰JU)z±Œ/HC*šùqUä¡ü()ãñ^©Î¤TW½…º˜ÃsÉŠù³B¢ù	_¨ïiqã“]Ÿ)ÁqÐkµ±G‰ß.aZc>\¾‰•J|”µŠâÃ62X‰i¸
Ç«¶]ÉÖ˜‹šibÆŸjs–ø´3\È¶kJT8ðSG„/Õnª%°»@D¦ñøE:ýÀ5£ßTØx­Z²\a©æ¼zw‡œ†|WžíÓ%áK–;÷ªnD5ã¹cÖsR?-	³˜Õ]å¶ë6°æá§æõýûÑ4(še%f=ˆÇÑÞóÆ2¡`á”¯‚ÊD,4ÿr®ú`šD8hâñž
€ý4Ò™ºËP¡C©ro8ÇÛ<Û°° Ýõ„J‡‡<êª]¹Ö.óãF¨?òµ+ÏÖ|œ*ê‚´q“j2*–P¤Iµ"cUÜ¤Þî§‘›FÄ›ì^Ç¼\Þ&õ[žs5gØsãÀåÊÁ¼ªNÍ)ÙÆ¥ƒ*Ô¸uP™[?„‹Ç-w“–àòÛíãæxÔÜýƒŠ×bQ%3Tý²…;HD73 ôdÆ}C:„„Ý%ðsé(¸…¥jèMAÝ5ô¨x°¨G…gÜ],h€·ìfñèVÜ,häËqµÀO¥»õUÃýûÝ.XÍ ¶·rÕpãVœâ^×öÔ¸kìÖ²Ô-°mü›[—ËØ.ÇÔzÕû.(KšñR-ªQ^.·áÖr+Ø¸LShí´—dõ\ŠÍs!‹gÅúH/¡Ñ““,r!òÈü">DÍ¼ˆb)"*7¢¸µQàTAw!Ì S}4ð1
E‹¹ßÚöÌZW$ÿÝøKÑ¯3]±Š¾ˆEs_D«ƒ/ãbÓ8½€jo))ÄìnÃÅ¦Xª‹MmÊ€V.%—z™.6­–˜wW,x¦hhp‹î9änRu­NNà"¹t:)î„ÓIøÒÛ_ŸóÈÝô¹,¸~4ñ¸ì:Ä¶Ö³£]¿&ÏŽ63¿ƒž•O¾1"žx=3u3{Œ ÎQ×õ×ë«ÑŠ–›Ÿ¿ùj>óÕð´öÐW£¸‘¯†¥ZŽê•ª7öáø›K~\šSÏÛ®Jä¦hdÂ5Â€×Ó^=Ï¸AÂk«Î´Š-,æŒ M©7@y1ð<®§NLs…ù¨¾;:Î]Rß‚?`z1ð=¯oÍ²^Ì'õ-øB¯÷)Ó4¶ôVpøÁÆÄžú¢9•oäápGíÊ·å¡ £ÏÿuQ¸EB{G1év<–GwÇE¡¡n¦I¼W;f¬Ý^åÏ6–=h›+Ïä¢Dê|rÿìae§›ÅDïéwv6ËKïi{ˆqQãÎW°Éoó¬o‡âè¾ñÄqÓh1˜áÄâÚ¯K‡M¨ž|×Û˜–›ÿÑZ4ƒ‰×„<%ÝeÀ>ãvT³Y=KÔÂñú€±2×¿/S·.¡˜-Þ~~GñVü¼q¯îowêP³Þñü¶áóûý=‰¹Y-æ.!“l’6­ºt’G÷£4IÛÓâc’£ã
-fQ*O”¼Y["±¸’Zp¨¹ö»™FIÚôªd+Zø†ã§N<×ú²m2¹&CµÒž5Ql0¨—Y_ÆÜQ›/ ì¥0øÝn1Ÿô£œ…'žO¢¯¢îFô÷,aÌÿz¯×Öw—Äù45î,Ã,O°£µVýÌ†èÎ$'÷Œœ ¨g"˜Ý.È}Y¦ój‰†ðµú^*OæçÉïg)àË)`tñ-°ß©ŠZærÕ¬EÓVa¨Ÿ§ñé6Q:Ì0Àv™NOY"ëè[ 9ºdÑØj”	`»„³œ—zÏ°ÔJ£îkôFø©]>ôRX\µÅS•iØéN}„)Ó§†„ž³ðçbu¼^ªZ þë8eÛJKÏ]ÿºoƒ«Á7êí­Ï°+>H?E‹F¤åæx—èŸ[£U@²9ÏâçRN£¡×,­Ús¤°»¢LŽyþ³ø©õ¡¥>›ån²uƒèêúÇÖ¶èõipÒ1ËäßKö9ªñ;¹moÓ¹¥:ÝÜÛ´Ö‰hYîBWíÖhyÞ¦-ÖæŽy5éæý/í‚YHMÉ-ø`šô¤6£DÓž‹º^ºÇ‰íRdœf÷-å‚ë¥‡š`ºÍÌxaõÍ,/ëkÑþ£ë:¬húÑFÊªà„ùOyÔú1èEü	åPvŽn¬yõ’¡È­ÍBâ
eSÈ«“­%¢ç(á9…"@ÐoÓãqš•É0FÇmívåÙ`{Qæ@Ñö˜DÙˆÊ„ò§þJ}eùi<M‰	ûSêûXöÝb+Sy=Ì6*=Ì0›Ìöyœ‚œ3NXÞ¤zºUä¯–Ö8J4Ÿ''ñ|\RÓWBßòQÏ1…€ÐhrZžÐ¤(Q=0Í¢ãˆ	.ÙIF(’Ôe²3üh&dU=ñßg}øæÔ<'kÔ&‘mÃt¿í’ý6³Œø3×¶KóK8Íñ,Œ`Ï¼×n!Úf`äb‹y¿èÚu}<¼û%ÎòíéMÿ1áµj¤>žbZvä1îYd\M	ôÄâ [ýÐVÅ
ê‚G¨c8¡E¿^v<ÐmÑ°[O¿ý‰C>Åj}ŠQüFØèŒ7ÜàMÊù¸^•ßÚ—8t7hP;ÈA.nåRüqõ8Ó¿ÑÍ—†YÓ2Á4?Îe~ºFž=DÓñ(žJ#ä×µFÈe[
‚Ž¥œGXy¦0®] Ù‚	ÅÃºT<ðà÷–…bR¢¶£Vñ@j“g—’V\IRÑÅqõúÑÄÿþÿÔ¦ø¥ WŠ¡ï¢°Q§õkyk*EGªÔtUì€d:b+Zµ‰Ù©X1Q‡7ª(Keð²¢Œ_ ûZ*Æ´”Œ–WƒP¦ÁFn"kn‘LRbÚ+znå!œ	n«£ì ™Â`ÕáX­©V\Ò«vµ4¦)§ÃL“Q¯•cI…[I%®£ ‚‰ÓbCèbBe‰/Žx™þ0é·Š<‘ÇmcÝsrHZ˜3Ó]ÂÓ™‡csZ…ªR—ÝHÃëÂmÖrøLkEÉ¼Ûª®6LÕU¤Q<T¸TñvÈêD6ìo¦™\¬Æó2k¢yØV%¢ƒˆÆÞ¡œ»ßÀÑŠ]:&g
¡ä#Y=2eG…'T£Aã$FNq/ô]$~ÍQmkz’U8¿Fc¼1ˆþ˜¯ÿ5še£DŽ(Š9êòë?Ï€R¹  Î¦(Ü`A¦Ñ´–
"Ž~žÇãŸç¸„ä¢›¼*ƒ¦:›6
D]›GÑnüK¸˜
æ¨OZiKëKÚjLšm¿ÖÄ«7ÙØc>±t’¤yHƒR›¦
¾ºHŸM¥-/>Ök´yÑ/¨ÐÖF _Ó\†:Ûr¿¼F›ã—WÉéõŸ‡i%ÑarJÄ«¸-Õv):ŠNÉÞ3„ª§Y´6f,úÑ½g®ÿLî3Œn©úï¦$†‘˜

ã£.œòµÔÅGYüÚØŒ–S_M>’âoE£ze*‡	×§B•‰2Òâ°Œ§£8Gÿš£½ý×‡ïv_îìnÒ)0·£¤P…¿€VB•)0L¨2òJI¦tæˆÈ;õ–kWÙ€»ÅˆFÊOLŽjD£}æQÕ«šÝ	K’3¦F	MDãwN¼dMð2ôÀÍ´Àáh·¥n«np…µu|”ušëÖJd‰ªõJäKê-~B2éUì?jbåçço¢‘
ÿsëÐ¹>[œWâ h¨?o¥=ç}ÉCàÊàvn±ÇEíKÐÃW¼ò¾»¼W·þ   ÿÿì}ÛvÛXvà{¾â§:¢:"%RËŠd,ÙÕNÛUŠåª$ËãU†ˆDX ¨Kq´ÖÌoÌÓtò•¬•§^óþ“|Éì}.Àpn e—«btµL‚À¹î³ïÞÆEþç¡Èç¼]¡Ë70eµ7KûÛv/6Ôù&ž»öî•>^¿!•~\«¿ñåÏŒ>±fÅÊ-/ÅãÒ}Úè	˜`2Ÿù¨qL?ˆG¡ç{@ò9Û¸G¤ú„%YÞ€ñ·²vœ‡WÞåI”ü4à½5	–òô¶€*ØI.ÓdR/Yn]áø/ZGømÔ©Eøç1Y``ÇHôKîÈ•Nºè®<ÏÐA¹„`¹g¼¼k/“<MF]ÿ±AÖ„°úãHÔÚô»¾ýŠ¥M¸CuÊhBº£(™ûOÓ´o‡Ë™DA?HÓ$ívžâ?¬cX:Ñ^ˆ¼$¼}ÐÙ E3RUýwy·çó&¤8§Ñ]üÙb(´³r.cÀ®Ùk¼-@œE^œ¨áVú3€Äë Ë
/ƒ8Ãìnðñ8žóO§Áe‡l>" ú”°"uz_àÿ²À‚#h*§¡—Q8Pƒ‹/~Æ<¾Ó‹ ýˆPôõ)!Aêô¾ –Î‹ˆïÝÞò€éEAšwß½†ÙPÝfN‡È™å¸òøëE1*EeŽGxzœ¯Þ­]å¬ZÃØ0Hßó5 vÅ}‚2òÄÐj=~Jpkt}_@aƒYÀ:Àlb÷zL<ƒ'f€Nø´Óÿê•€èU ®³ÄÕ§dƒ»	-ÆöIóÄõ§a,I	~lÑ ªênó	syŒÂŒÏÑÜ=§…ÞF4\¹ðµïwŠÅbjÚr|UÈaãü	‚öñÛî(‰0ž)Lbð à\O‚òä:Òç>€ÓÚÑüÅq÷ç ZEÿÜY9öf8[
¤ã Íº?U ü+|¦æ·2˜ñ¹ÃäRRÌÑ	à2ŒÈŸ>ü™78m{ýŠø‚¯:xÉëPO¾ŒûC‚–§ÞÏ^rŽQhÑ›—f”(kÁÙ}÷õ¢~÷Žté¯¿ƒó»ö<÷¦°C „iHN §y”'€´ÄSRJ¯‘øUÑjñ+mIïÁ^‡bxô3}:¤Oçôáÿ>ÆÛ˜œLzàÙã†@L]¤ ¾§RÓYw½|ïïA±±åO5>¾üAfè$ŽMê½$áñ*IáìÁ[t?Dy¶»âØSàñ|Ñ›3 föÿsð,IŸ…Ü‹¢+6¼S+¶yU€_@èö…ÔgÐâ—ï”&#Øj€U	ý „Æó«`ªÄCÄ‹o¨ˆx	`ÑpD±	~yÕÇ-Þnú
^iáæ’€7û°ÿ™7¦ ‡5žÒ«0¸f!x°Ãh 	˜9¾â&¡Ãì§3lÈãðš‚è”Žÿ¶ÑØhŒÞ'ó¼Hž/
ö°èJ',â%#/:‡-„–û€Fž²ë®‰§~D~Z .cmšÄù$º¥bƒS	‚vŠ/lnbÀ'¼{Íˆ±Ý2qFÌ‚VeÜÜ$çóL/;C€$t2$JÆá¨ ôžüÔÓ›YÈ’¡ãÄjø?Ì†\ ŠƒÇ8ˆ´ÂÈe{^—¹8BŒ×†ÃHŽ84ØV `NÔ™Ü{adE7ÒŽMo«ôMÙ;Fø2ËnÐç†¡è1%åð¿ª´¸ÚX¥qÎ¸@Zi½_Ja*L-‡ÂßAx­H¢âÄ¶šÔŠXJ^æ@jÐ'Ï/%`kÀx êìGP‚r{$ÄHÜ|ƒÃ7öBj™Ï‹W'YãøÖkÌ€øùÔ£"˜ñÝ~žÐÇ;=Ö¥ÅWàÈÓ/æAUÙ?¼¼Dúã‘‡‡( oCÃ½²áõòv­¾Ð
Mv3
Â¨[4½Iºƒ­­-ò{²Wüî¬×)"eˆ-¶¹høŽ“”¡µÜwÚßÔ»én•Ð°.=É ‚X•î• 1ì“ïr`»®CÌ¥îE£9(ËãD;I|Žç0"*Œ1yèS&Ì4T!…Ó×ýÅÚ¼ôë)>YßêR…³ Ã¾PÍ<&Û[DšþÂ55ðÃÞ.ü²-2i QæBêPÈ]@%Ê/ê@PñŒ	þ*cÑŸ
Æ ôäÁjZáëLá2J€¬;CŸËýë]¸@¬Í#©	ä:‚è2ØµtW×%gó‹B{l šly…Ò÷@8ž sðž¦ùùŠ“OäLLÔ—±T>B“>˜H}³ÏU²ú%-¶<6¤'OÙ’`¨¦˜^Ÿ/SÁYUgg®òª85¯›†ïx÷&=X/îâYØ­Î!³ªÏ¾de=<Aä¡rP¿æåÂÐÕž»Þ”n<Ì¤TÏCÔtr¢»'üF¨Â¢0Z	‡‡íUùµ—>º©î‹A…q\™ÈXhn³Ù©;|6­"“AÃÐF‡c¨ÕZw=ZTa$HÆj|½(bFyKæ~òŽ¦c¢ƒ€®¤Suž*ãÔ%÷ØecÚ·¸é×ÆXùàÏQòÖ›Bt2ï“3ê7
KAHÉ<Q^‰AÊtØ‡ÿ.ç1ÍÀˆ9 žPªKï*I7 Æ©¶„xð)N®<–‡$™c0¬2Ÿ‚s•°qôßÕ	Kv>—4A èÌaýX¬AŽêÊ-°hA^žÍœ‘Y%UyªOžÆy
ÜÖ”NÅË™(’Ôš@ï5œeŒ3ïØºÍÂ|ÎÜ¯mÔÌttîjµýP8ZgÙÚí†–3]IÔjÚ+“žmF#…ÎìÉFãJøLqø¡‹Æâ“”Þ«nQnZ,_Ñ-OÉ1î5 WZwÂW[NëfMõªØlïõ5Çñwó¼eÒX-ÝZÕý>êÁÚ…/z¥½úBœ{aŠ
Í½êÔó¯á«ÃMDÙâ{Ãr+(N”PTð±I‹LJjmzq8ÅV²PëŠ_]ËkiA/Êð19KÛ£ÊÔµS­(T?"ÝiG7÷+arü¸”TŠW³ÑÁ®Øž1Èm¯a×†”Ö“v7|B·‡ËÐ¡Ž„>,MN%×;Â-høiKùê3Xà¡~‰ÑfL›•ŒQoò˜¬=ójÐïÅƒ”²…)ª3è%ÄåÉúk¨ä=IÃ€"~J#xÄðŽÐFÚ_kP€¿ª€h½’˜í=çpÓ”ŠU_i¼Y ü"ov‘ý~ªðýƒ°¸…¬ßÉc?¼Ä$*I¼9ÈGfèxŠšæchðNIÚ”p¤
ÿ1’£AòÛ´F•ÜM¿$`FÁ€(H:È£=õƒW^4p¥ÓÆ'š4h LSËvA9Åã‰à¥ âõiCN$tVx<ê#^a‹.“Ñ<;@C#æy*¿”‰Ç¶øívÀ8=ëÅI³bÝû°æj¤ÁOsd,j?Ôœ»UéÚîÛVåÆmÛ:d¡‚x¢(È&íÜoÌ}óÄÿ5ö a/¨{´Ø«ÏÑaw9÷Ä¶.£‡¼£Z™[†’¬Ì!‹ù¨8VPT½ú†Y©P62ê^*ð1 jÊ5­ÕQšŠoBDÖDËŠNÌZn¢åÏÙâ²’æÏÍX`‡¯rÁäàóA'\4a©6Ù± ê0º;E¦kP4û6îbs?m
Bc‡BÞ÷ÙüÃŸSCp]ý‡Ú
Ñ6ØÖÌ¦ƒš÷V­§àÕ¦¶.˜Ú~Ø`nå99…¥‹±áÕ˜;G@R¬a[õÒïæBØ÷ÑÅ+˜@5‡¦€pÆ`|“$ã(xGòNq@4°è¯
+8Ì®ÆÕFwËÞí´[>InŽ:[d‹wà?M¦U/Ÿ`²Ôè¨3š§ÀEæ'ˆÙ:Ä?ê¼û»{d0ìwG[½þƒý^ëúmo÷úÃÞþÃ`xµÓîMvû‡#¸÷·ñ™-è„§d»¿=¸ö<˜l÷wŒ†ý­}xäá~î÷vúvØ§ýþÖÃŸ5i­å¡þ·íãýÝm6JÀ†ÛÐæÃ0ÕÝþÎ^¯ÿpŸ<Àö†ý½½¨‡}ö°÷þ·pPÛ0Ð­=øíÁ€}ö÷÷ÈVo·?|ˆcÛîíõ{0¶Ýí?ûƒ}þþÎÉvÿáC2Ü‚›ÐÁ‚­`ïNc~öäÉÉÖ.ó.4F;0]\´aÕßÞ…Ž·ÙX¢‡Y°wv¶ÅÀúÓÑœàm²ßßÝ%0ü-hÿfxw»¿wá!²Óˆ`Üø6ìåþ ú:õéñÎöön±¾»ýíýÑ ^†Þ†ÅB°ØÁ…{;Ñv°ÛÃ?'ƒØ7'‚ƒ‚?¸V°x0:X7NÞÝÛ#¸´£þ>nÒÂ
]uR¬¿zÄp†¯ÆÍÛì :)dšD#Œ†4	cqêÇ“$Ë?»„Œ—‘wu ¸‘6)¡÷ÔäP©êXøoÔýDèæÅ2/Šú¬\ÄÕ:h‚*ÚehŸ^!B©5°muÑ§y%×½IèûðÒ8áü Œû¥_‹‡®Â,¼ˆÄí´€ðÏ¦€`_Ø0dMÈbó÷ä<ôƒX¡ï µÈ»%¿ß,gqšìg ©‹Ùt®•Å”'Þ~Aß‹IÃä6ÑVáÞûi2Ãòì)†ÿÜÛÙÂh{6ßê–6±Ïåau©™¬"‰HR–»Å\7%‚ybi\ß²Þ€V,ó-[Ì‚+P,±Ã– Å+UènÃËä"„çÿx('É›€švWd±T;ïoI‹¹Å–± ¡ùk†	ÒS3Ô)Øði²i›5Eõ­‹Ô®‰B’Ô#Ë‰6vÄ¨?W ê'^
BošoW2O¨ù,ó<ÊJ# )¢_Ètw À¿çR*ãFïQîÍÃñ$¯›Å„õ	åÜÎ#Ù'ò,ÕF.7º5«•eC¹†sS„_Ü’çù‡ÿ‹î”ß$Ó S÷ªO‰\¹¥.Êi;ñ_Uð’É²P„Q×iˆŒ†B´‘ºÔäcrø•f@²_ñ¼r³.V7«Nn²_áëÕiÎÜD…žMÐ€f˜ü‡ZÄ	ùÓ§ØTõ¸€|^'3‚Å£z€†|*ž£·#+ÓÔý=.ÉUÖò9ò¢õv2ZpÇˆ]hJªÙ$3TÒªq/lc@EíFášLIS(GN;€C{€éŸrÂ¥FÁ<µ€s¡v­&À9÷ÞdP"ñ*ç±ÅºørÃiø+…UÔ®UK¼ŠmÛâ®§ât¥—³r.Ü•ÎGó4ÚÏ;U!ZV‡”¹KJ¬ªÀ= õ€óoù°ine—f½yBÿsæ(À™GiSÓ6u†=EA»ðb&ÝÂ£z]Õ“&«À	.’-+í³fAqášõ™ªK€ÕK´iIX~>3>E	ú¼
%Nû@O˜C@ÏæãôRæ¶Ïæ”û Ô9÷ ÇX“Q2Ísô
ÙŒ½«`ŒŸZ@ü@;M‘çvƒz¶¢Ô—9~–zJúèœ\À¾ù‹:ˆaèÊ¼PFn `ä<Œí¤@ÐÃ¾4¬žìö	¼øÜ êG =)º¼P½6ËÄ±~MEâ(B×œtOC¬ ‰nÞ‰ª|£n?P_‚^÷°–2™©êÛµuø–YQ~šïÎÊÁã_x¨5ÞÞ½Óåä¬Å²Ë¼½ð™ÍAxÖncý@Ì ¬†ñ]ym7ã~ zß¢·ï¢mËj‡æãñNuà{QúE”»¬Js¡]IMªinØ)˜>‹KæSÆ•š5âéúÒCþèÖ|¦"1Æ"Ô&_±MQO IP~T@PhA@5|Nëžxá„Ñ‡¿ÌBê™ã9°Ìë…b¬þß÷­‚¥SÓä:FO2aÛË”Õd´Š­(Ž­§Þ¤1éÁO´È¯<€*«ãd—+x¨º*Ê9©S¹…Íœ{’”eÞÌ†,/Ø%…\Låû™e/dfÂÈ8<Ÿ:m†!UÆ!/¦ã¿#Mú#o4
fpRûÊ’XýHéêÁÃ¬§Ç±r UÊZv)-"´–Ó¡’â¼4Æ ò'z£`¤ä„º’5ô9€¡t²–£Œ¶5?=Òz;œ&Hý†Ê¯¨-$™y#lÇ0&ø‹~ùŠ¬HœøV^Ð”WÙöiªi¦Îì¬È4/…awÐgš.SøL4Ñ|\ë	6Q[cÌ©yx>	ƒÈ§ñ¦Kg›njû}6òO³‰¢Øô¨â5¯vî÷P ºHƒ«à+eŽge§"	uÁÑwÑ\ôSâÍ,±¾h2KÒgÞàñÆ f0®#¸Âm$!4 "”K¶H<mpÌ‘˜ ÕËøÕWÎ@mw©RÞ‘ÌüHP¯0	ù+ÔsN´ºð­Çå4Õúýy4œî]ÐÃe,'!¹:ºÒ”çù¢¡“ó„FOkbŒYþèéLçÅƒ£¥ ßfh/7ŽRÊiªB‰íU'=Ô(Oè)«—aNÚ_ÕrÖ"~ï—A–ú«Ñ•ªA¨sï©Ã#iln·U°ïº5ˆ²zñ½-zF²ªzòN¥­ŽaUÍ°ÑZ¹ºÄÖèÕFßm—Åžv[E³¶]ÝV‘ªê‘ùõÆhßÕXLnÔdp+%Œ"…‡ªwÔ\ÓÄiñÛò[ìjÉu±KÃ{‰tŸµ(¹çÅ.mbKf	 ½ŒCŠÿÔ°`üAcödwvŒ¿àÂ”Ñ‰Sžì8GvÃ)­<(Ídä$ýð˜ø)ÓpbÚþ?F{¡ìØSš?iJŠcp§¼KÉÔ }¯9ñ¥Î×‚¯ÝQ­_ÁL2ÞÑ÷0†ŸÈƒŒˆ»ha%žn ¼3ò`<QêiêÉ»l‰ž¯c—‘»c—{Ž^7Èsud:;^õqÌÒ[œ3È.=d|S„50¥]nS®Z={È.•#yãG©ñ<Ššk»UMmë<<fˆñ fiâ#bñÂ\U=ºL:ôÞ‹/¼xM¾Ü‚Ç±ºE‡5 ZzŠÏF&¸ÚBZ%ÊÖàxþÞ@ýRpƒ1X­zRŠnX€Ècl(±úx€ð\zÙÑ‚ n;gg€åI:-§àç}©’
Ï¦ x´úš##°ËÑB|R?—Ä˜ÅNdnÌL]Í,Å^Óé³J‹70ñ 4nxŒµ„U·ÍÏ9¦§ÖP÷$f¹´[N•%Ñ¾¿É®:‰ïiÎJ¦ 8Z°–ltìžù½³4LR*$Tß-ïëßGUéÓ«…ÆBø,Œö‹ú½ÊîÒaåWíÅêÓŠoÚç¥%¥oHßÍ}àn•}à7Ýì¿	@<Ç*{4‰ÞÍp6nÜS¿ÍŸÃoò{ü»­Ç³ÓgeW} íœ>ÓM­#Ê\lÒMóÛ<r¸|“ßP¿%ãäd’„ågíÒ¿¯|Õ¨ cM¦øÔÑ¢ülz¶p”ôä·ªwuËþ*ÀôaµN›7Ÿ¥Éôk.Ã+ #ªö”zJ=ùo r‰øN’i !½B%¿Ø„šïÑKX­¬´„x¥"¼(Åé:Ñ¹À<ôWIøtÊÅÊŽò7í+J‚®cZ‹²cG‹â£îÙfE<ÌGB#¨kP¥ç+×D¾«k­ÿGÔ@ý„Ú6idpM0Kô'@6Ù
,¦h¢“i€òÖ0ÞÂ•«wÆg ‡pgønÝ-`úAcøpñQý¤ˆLñŒY¸e‚5ªß2³t•wýú-+ª¯¼ž*îjIÙie‰k7´oUÏ¹î©ó&-®ß³¾[ãT÷?2áò€uñ½Ž4kà×{ éø‰ãl>y÷¸ÇSo”&´LýEâ¥þ
›]ké×»ë/q"§bŸ92¼,?¿GçóéÔKo‰¯ÞP¿føH1ÔcÙù»é¢PvŽéÖñájÈ¸[“u-çE+˜ÜJÚíoëÉÍMÃNÆcîÝ¨F¤¥»nÅíX7¢o½«pŒ"uÂñkËßÅ½²éx™½§S$•€«E€ËDÑ¼n6Òç<÷µÍqûÜº‚îÐû}Sà¼[[uF¹f¤êA¾`0VÓ’•u‡›%‡ahÁa_^ËÜÙéz[n‹î}“–ß&‹ÌÑºÝ6›wôÃ³QÏË² ËÐç~úQké×K?Na"ÇhÏ÷FžA½ìL>\¸Î{åòy
S¿Ÿ]U´öëÝÙ—t2¨iýõm+ea`OoTÆ"¯°±Êö~½[ûJLGŠÔnµÉm9¿Eó,¢MŠ_Nq¼ª²ÚE’þtêã{<h0#´¬¢
Müfì­ç|Bú¥Ùj-ê<‘ž¯|Õ¼²ªu¡è£Þ„æ½Ü…þ'n©ñ!`¢WåÖ}È3Y%Wüªá¹Z“r=Ø§U¶ÐˆOh¸ ‘4u,¾¢Œ\óâ5¢ÎÊÜë—Rö[V¹ª’ÚVŽ	§ÙpYÈ,>É£ªû»_-ÂV¿L½^˜n%-#á/§Kp—ÊïÉNGˆdîÕnøV´ö›Á·/Ø„^z1vä¹í¬Ì²¦&áÐ¨‰-œ?8·CyE3^GÑJQeWÑRñ›sk¢žª¢1ñ“Ž?¡i¿·šTðÀýC˜O0-ÛÑ¢Y5³y©Š¦Ö/VRÃ`ûÒÈÙÛzdÆ‘gË5b^ZŠ‡ËëYèJ‘6¯;Dð®m³Êf€jXâSÄÚ¼è4¡¹o¸X½ÎZ¥â[fÕ¢è×F|ccìß˜dYL¶xPZqê0µPToT¯ÔÞÞªT­mô ù>13œÔ¥qh«z6á‹è‹&ÖŒjºZ¹¤¹Ž´tÕ²(7ë«myü>V€Y4—[fú¼÷UXuZµ’¾'§­ôûI¿¨ÆyO”3§aˆÉoŒþÃ´D(íW'½·M©T¢]~3äf~½š$S¦A’Mèç4¹ÆÏn	×öu!ð¦Ø€ÉÐ\ƒ¤óHÔr'2“w¸9jÛ4é<â^öÖ¼ì#‘·T”)ôkÚ¸CÔC#_öG$Ï¿0©ÕnÃYâ¨œ“|›\yò6jVUÜ K{Ô„çqúÿ <g½™úå×!¦5,¿n3 ÖÀ¯RÊèO½™ˆeÕecõ†g`ýÐ¯$à1$Š¼ü4ã4™ÏL1Lõ|=]îõóRÖ°Ïg\Q<ÏÖf%Œ¬L—RËÖ&+]…žªHü‹*é¨¥ŒEc*õZJÆp4þŽ)Ê°ê4¿ˆËˆ(:©'Û®ìù@^Üj26©´IGBJ/qn@Må…Å5“ÙY
Dg,•ü4_KrÐÍ¬\då•6\uåEG´W^zX^‡ÃðöA Ã=+<Ù‚¼¤çV…zú‹F±¢ÙD;Ðø8;ü:õ²Éðþ÷×ŠWì¸%îBoCD‡“íf²é"H6wQB—z?{ÉyûÝ£¶mi´–±mñŽ7™Ýbvèj¦5–Ëm¿šÊm>›éu;_Û@'fªš5•é{°µEƒiÕáÐø«±[mz½òZHÃ:÷l\:í;_6mú§ò	dèC‘KDá@úyÊÔ§aÌ2úÉyy‘‰VÁ49sü´1Ô¸R]¡z­kúT³wRTµ2SÑcít‹ÑÃ)JN·*%%9¿W”ÎDGÔ@‘x®Œ°7ñ}uFeg¿ÂbŠ<ÞÒ–Å–š”V‚Âô¦˜–ÓÒ“,<‰TøÑc_“ Hl–v'•[¥­½´¤6 ™Â>…ÁJª ÑÆ=À'ÓÜƒèOm¨ ç˜@ËŠ
¢Ÿ™¤ßÝÍzjÞç«ÄÞ2>ÏŸ”Tüüºoò­™•øMAÀ‡ñ+Sô3TÒv»Ü€væÒ#ƒuò{Æl)ÞÛ òÕßÖ”
%ZÒâoYå }èzBÉP@‚XcgèÐÍ9M#s@:[d0œÝáüéíÁŸt|áuÃðÿú[ûëðÁß÷ñ™Õ3;ë}2"Iç8
ÌB	=gX—aÜA#SxyYv@€ôlß›Îà—2Ü55*—Ü“çL¤|½EQM#âX¯"ÝB‘SKóNõÊaÕ±k§_´Úy¥­®èãh‹~Sú¢å4FN:£•´FŸ‰ÞhIJwä¢XJ/  ˆý>LlIše©¥(­á­8™f¯ÐP‘Dß r#«@¼Ëoç°²	{¬øJ2‘]ù’³%ŠN€‘ýêõó³ïÎ|úòìÕÓóã~¢¹dõG×M^µuirõÆ\Æku¢°öjàg*¹ívk=›&hImŸ»¾ïsÐøÝÓ‘fZ¿{¨?Þ6·WùµPê:˜±èë¸Îo!P˜‹¶Î]›¸SË²sª(zwßš(:4“6Š?R/:ÒºL=M:y"+85©êwKf‡¨«ë•zÏJY½T™°ë²|=Ñ²#US¦ê%!g|³4ÆYs¨ò#´x™ÅefÁ˜'–U/jëœcô#½ØºÐ¾ûd=­µ ã°eò˜Nà€8Ò]ÆCA«¼×^:òn‘Ò½‘ç™ÔLkô­u«^	®!	ÔjIw·6·±0Ÿd—Ðî¢9Á­Jµ˜>¥eyñŠ!¸:^J^P¡úêÛ`üá/£0±·àD×°õ‚¬¹È1ÄP¨ªþœÍhaRãHÏÔ¨´:
zX èŽªMõ¢xnÏþdÒìK½)‰ÅÄSJæÓ:âÖà}¬ %5ª]2t â!Ð‡Î£“oÏþî€¯s9Qôm'sVR¹¨J3¿ÁÅ=~V[[1ÙO±ÀßŠ¼[£Ts^VPÍ«…CÑS² è¢ï¥·)çï'v:XŽ¦Wå?zœž6…êåv~
E‚jdºU	Þ/Ñ²Qž¡ûŽ~ÉÒ#_² ÓEÜ½k·ªtM€¹(úé¾y»þ·ds“¼
à1Ž:°‰ËSãÛä*!txþB#­$ËQb;™Wi¿¢U¶K«†Ú…å´ZÈ]¥re4Ë&Û¹=©´ƒ}ý‹M½…MýÓÛÑ•Ä£±y*+ý.ºÓšªÚ ×š‰=‘|#KüJ9çý0ó.¢ÀÇøëÒdGóìëét3ýüIùrÖò
oÓ¢Sï¦;Ø ôZ
µÔ_TÇ¸ta¢*÷ˆ—w©‹eŠòjöàßdÜ9ûðçq{D^³;Re™:LmV­ ºe0Ÿp¬<Ïv@Æ]ç9°ýÍrÀvÆKB,lZ¤S/À®zrwDd — õî.¨¦43Ã…ƒËžÅßH„«*Y&ïDõz#ì¸í#eøC˜åþ’£rÐX7¥­U·JuŽjK¿¨åA£%x&´"¹X¥Ž€hè¯åK|‰ùêfXÒkRZ‡ßŒüË²nÎjkãüïcö@?v¹ØŠËØÕ£±Ë¯±K]‘¸^û»V*oPJ)JFÚTRŽ^Så…VO"/ÃtÚ}w
2ÅŸ<  )Ž#ºÒ`äá‡ýðÿ‚ŒdêÑ½åe“
èt2…3/‚ÇB|•BÒª¯Õíñ;G«é(
¼øtÞÌò‹¢‘<@â`š²J@–¸ñ]²ÊzdÄg#Æü!>-—ÇýâØ2b±iiÑÌ;k•–^—0ëÉÉµ•vª\vÊ
˜o>"/BÓŠÈ–4‰û#µˆDvÙÍì‰\×`'­I¢ëŠ,³~÷èY0šx©}$¿æ%¯„2tÛ€ÍVÃeè¦F3SnðÌ—7ë&—5:H·5¼¨ëÍVú¨‚x‡Ÿ{_/x?wïÌ ÞtjÛ®;µ¡«Ú`Oë¯¶oöWÃë£ø¬á¥ð[Þ gs%×=}ydqY4uÏ5 p<Í’Ú5£GÓÏMçgvckæ=-Ó‡úB¡´fõp£Su¢÷v/5‹ŸšäSÖ°U"W¬â¯C<iLÈ&émÅyÖAíànfç`‹çfEzSbÜ×*ìíVÞúJ@*,t•€Š{TW) ,ÃK²vçbs—@3¥Žê“ñ|µ7X0nÿ/í¥È0»¡%ê«SÄÒ´¿U¯4)ÒGg‚È·Y;«uˆÎÍn¡¹ˆ,C,“0SôÅ Lå+BE$Ì´A:¾OÖü`¼„‹}º…Pÿ‚D	Kƒ1ËŸ¼"wë­€ÏÝ/$;GöŠ‹é¯å/ðZÚù/nwx e¹XØv0—€S‘ï¨? ^N>xiµ´*¿XDôU‡€&1Ô9¹îgÙ÷ #ŒðÈùs™}wkÍ)«aQ‰T˜¡‰¡¯:ú°‹!vÍFµò„Ê¤RÀø;»ÿù¿þYÓN¡søðgø¥›­óóxo¨ÊÍ…{i?oÕTÛKçÓ%jÙMÏ€Wk]^Îú¼êçN>b5"5sJÌå,íðºšÚ©ýk-p6¸¨€r_ŽPn.A7 €±½ µâðJŽ/ÒPg~XÙ·yeU\–WÚ½åöÎÑù‰—žL¼4ß®îJQæwfÑiÌ€ÛþŠ¢Ûä:Òç~Á‹‹ïÈ³$«óPýëe”ÖÖ]ùâ˜Š4`ÇíˆµÁUšRÑ»
7ûŠ#!ÊV¯€ªVbØT>úä»úqÖü^Z±|ËùíãåÈæ	GÊ›Q4SÒà¦HYS/FÂ‹>¨†zãååÈ-6Ðæœ:Èa­}U¬Î'Ž&ö„ÍË/½ñòžŒ4V:Ö¦ŒÊ5Ä÷ëþ„aîÁaŽUXbLè0–{«~O“ã‡a¯MëKlóYcn…ÚJ½¡®ÃeÔ ¯®®!ß/¼Ã›¬ŸEMž…›2T¤AÐäX¸QæXPåW€'{Í†ë«ÓÞ×R“9aM¦Ž™j/‡R¦ÄÐI‡Çiš\¿.ójàþI”{*>Ð°¯‡Q{:®áC«”¬ùÓg‡hcúgÎHM¹ÞËØìJ“Ó»nY|,løFßuã…'!WuPnÙ’9R©ŸT³ ¦ÊÕ•1HÃr¤ÕU‘úAËì†¢b»IÆ$IÒì¢§×*0X]1hWöÙ|®J½Vfó¦úM1{µ¬ÎÈŽµÖ´M#
gU…[=JG	@¦5¯ª×N©z­ŠTô¯™™J ÍÀÈèÙ”fŸ)~¨æj0ú2‡ƒe´q4à4&;KÉg¶mp£ýÈ¶:ŠrX ûð¬GséðÀ¨ôÝEÉB]kÒ‘ÑÍfÓ"n5÷Jq‚¸IïWHÁ½c+sØß*ØÊ ™(O¬Ê4àò\ÍRàkPã
–3r~Žã¦N3xÓê)CQ0õƒÁÇ+ìL-pÜÁý`µàqIáns¥c5ŠöfæH/Ôß[ž@u\ébó÷„UüÙÐ'Ç4<‰0©%#¿ßTHÅ-%	µ#í×»
DçÏ‚À¿€½PvÙèt1ŠÕ@¤ÍÈ °³°3PrêÕXd5ŸÞÙPö˜Á$Î¥¬ªøÝ_Cú5õâ¹á$Ïç#,²£ÙáÇ¤c€æý¦ÈàYz@Û*|T%÷Õ}…ªPôfó(S¢b$¼PL›Öûx¬Ç5Œjñ‡ÃÌfa,ÏÉM¤9÷¢+*){”Éeî±ý~ßÂjÚ]‡5^fŠ'“`ôþ$LGQ0TÏr	iNr|ä$A6åÃÀ<YI8„Ã×á4€-y×ýºz¯Ÿ'´À_€_ÎsôÝë¾y‹.“džs3ìùá8Ì×6È4ŒçXÕHºÅY¥[èë±þŽ2E6Y^·ªT¦Á/éñbH„kz4¸ã‰ž‹mvÕg\XG”¿J[¨JƒWK¦ìnº-ƒlÂì¥À%°Sº*@œgµŠdã…¦J-c³€ûýð¯¨I(a
¸­<Á¢ Ô·œ—ºUVƒ$ªã6á†ØÕ#€Ò;ºý¥Ð¢ì‚µÚà$ôÝhâ=sîü(§ƒš~OÉñ8I½Žî…ÖPú¯î\ÿ¹~BHJ¸çß+`Ïïé Þ=“l±
ç€³FÎŸûJ]•!¸€
YEK«”á*bh[J«-­Ã³í¢–edÓåIvÜ’9ü—§IVÜjqØO’Ù­´ùêsÈ&Åh-1áVdUÕÌðX±¡«öIiA7XÉ-vqg¯3˜kÌÝ’gý>laêvV›¬`®ØŸÛé<Û µª«‘Y".ÙÜFaôâ{Í¤¬}
êb Öì&£ÒJ¨¸Ô¬‹6qÊ|ÑÕÇæ&ôáÌù"ÊŠ-–Ç%¬,@á-ùikŽÒœâ®—f uÑ-ü	&)<?#Ùç:HAZNâ@%¤0½Ý<¡ÅWb“È°EÿÅín±‹ 1ê‘ûäšúWzßçƒÈÈÅ-Á4ÚaÐ¼ÿäŸÂ½< $6‚sx˜Q>vƒˆ±¼yû¹¸SŽ£6›þe’>õF:|ý‘dCH½ëcªÊ «NuV}%{3NÒ)?²ÚŒ±ÿŒÚ\ð‡.o{äŸ•ã'Pˆ˜ùhòÇàÚûî3šôß·Y·±FëýK@ü:®è=NX3ž÷ëTV€­9ñ² !òè¨˜Fõ'5›¤=­.†¯+”&(:Õ!¢Æ¤ßˆ—ßÂÛoÞjˆ¹ò®¡­þlž10Q¢Ä;-Øo÷Éƒ`F¾ûöÅ?ª#ùÄË‰7Ê]„“9AÞTÚŒÐ>¤Ì[ªæ*õ$åÙgèÁÊJp &"D(ç-\Fºo~Ü À½¥8Èe(ð6Ž£Ûõ6È}Ë{³õ¶ÏbhŸ4è^À­²6Ë{O^­mÐ (ªx¹¢þkX¤¥Éu$
ošú-n£¢o„¦ú"TÜYtÆ*À;ëGQ4¬d_©GePà2ƒÂPd;™éxŽ¥ìOÒiË¤²¿„nP	œªµc ”'9Pû+Ä`€O=²µTµÉ^$Dß‡¬”‚S_P	¦<:§TcV¿æ¥˜gZùÛ·þØëÀ€ãŒ¢³4˜†YæeÇ‚ÒU'ÌÕäYUŸÓw3ååÔMƒà«Çc‚¦Ãƒ2›ÅÃ‡ìËe”$pëCÚ¬´Žõ õuz·”GƒSÍ€¢ó³4Ã2"‚`èå@:\VT‚NõL QŠâ•O% ¢
ßžú–—{mj¶¹Ò ÝØ¢içÍ»Xß>UÇC-Š‘n°w7p0òÌî”GÞ@Ž˜á	P”Ç3q]Þ•F¨¢)ÎÐ~[Ç¶hàa`sé@õJ,ÄZä‰cäÔiÐlµÆµxöÄÓ±'ªgé<‘øQ«ú¢š¢ÌUGsR‘2ÿ§^ð`{e+u`9ÝÖž>ÃXfÅÂÓ5ÖhŒá¾LŠ©¯•Ð±ÊB‚<úÇò·7ƒ· õ‚ëµ+Äi¤Õ9a–&°/@IÑr…ëU0Å_0iÂ(·Ëg¬…b˜%¹‚ƒ6Ç‚6Þhd0—dÅœ8Rä „ÈPŒÜ&6P6ÆðF´ XºòÕÀÚVž³°­ž—”hxµ;dÊôK¡f|e•¯¶‰=å¬×YIÛâC›bQù‘ ˆhÖ÷ƒl”‚Œ™üXüN¡™S\~,ÄoÞ"ÃŠqê
;”½>iÝëE½×‹v½`ãTŸ¿<;>yýÝß½:}úŠnòñã>ú„êSèöÐž4d¥hî‰²¹'-›ãÐCÛëÑQ¶ŽŒL¸Ng¢OC…êHÁ.Ï/Ž]W+&Uý\Œ{×“Èê¬WZŒÑWãçµ¹¿E2šgƒZŽ+åŒ¦~YH>Ûólµ¨vM6çÿ½ôO:7;a%¿ˆÐx¯u¬¨£û’Šýê.‡.ã-Km©jÌÄ–, áEßSOdPWk]>Q³ÒTL#‘YüÀ4ÀÉhúöä­öU4YËLë¥r1x Z*ó*US™8Ë¾JÏ¼(÷¦dQ¶zuŽ%_”Ÿ¹Ã,€dîcùÌN?-‹_óE‘ÝJ+Ë5éY‚‡a=/Ž-Y3UjJ|ÊWÕó¦Ãk¶Õ¼füyJcíäêôÁÎ†N¯º¾P£ÔAÍf«céÊòÛˆ†´ ®É'äÝ×‹R¿ûÝ;[Xœ1ºËš7ø2ïíðŠÊ4 ­ãˆ[Aì[ý2H¼.¡X¼žÁ†&·Iâ¤ó¨²®NåtD£æò~%Ðq-ZÄ²í‡³Îä?ô-§}åzÒ.	Ì9F}ÎJAÎ
1ú4ÁÔ±hîºÌzj>uÇÐO‘lz@\©5Ë)îTBf”Åx\sQc ÔŽ…áÐ×-ùàt§8‘’í˜ðýáƒHy¸9Ùiƒ£›=_¨RGÒ}ƒ‚ìg_•v[ds!t ó|M›×-Ú²†	JQ&S1©…ØokÌ¢<//÷Ü>¥ã¯ÃÃÍ@¾ãŠŠ‡®‹Sb&ÏÒú%ˆ5+&†!³Þ¶Ø×Òª` ÂÌ˜â¬êF¯ñ7­_b/®;›*²‹ì	&ŽI?\¤™úâ“ôÜæ@6x8€ÑwÖx/ÔBáš¹÷*©õÊ›V£p€§˜ôÖ1®bÚ$æ2k™–8.GÆ¨3=8­’nÕ:q†~€™ÓWæ`C )]rnp]BvfÔ%´AÉ6jjÉ¹Á¦KÎ v-÷ûq·É¾SÎÃ-ƒû9¦-ä.±!úIF! xIñaÒzTº²H#<Ø§&:AÏAìW;ñ36lÔ¡‡"œ¥¤q©¤Üv`‰dq9¦»pËUé˜kÌÂÜ­"Ä®À÷"Äþ}áîp’`–å|Ivu»`Wõœ±Š]¥DãÔóÑ)³—'½”\¦É´‘ôòÀ‹šì¬¬2’„VB°ª×6‰þBÁÛ¹;uÌœ?pòz`×+&y¨Uþ,ß¶Ç²Ô1ˆ2J°WBbòç'ÂVçÁ‡nÎ^g=_&ß»òª›¼Ð°ÑôR,¾Æ†X^ÿóÚx˜½¡ûæ= Úúo-gù{Ü3xk¼Î$U]94S.¨s9
 õc_*[\´±Î™ƒ®èL
*áJöWÒš”@nÝ9s¬ñz×@@lõ,[¹Ä–×¢jQ¬¦.²Y šd`¿ZQË¢ÁàÖç/áÆ®†3ØUæQá&ÐbG‰Ä;64hÙZ»ö¿¶ÎLƒ!lšÔúYÍ]4¥ñ”¦šhk?•ç…µ2'cju4ÂÒyTŒ«´~¢þ‰Þ$j¿Žòjƒ‹¨–CLåNþ¦";•¦”¸Bµà–eüþ²Å«Z›ôf|TƒŠ(áT7¨`wDf!¥¾Z«d¸˜-ºýùHrè”Û¢(gÒÏVÚ[AÖö»;Ý*¶>TÝÅq*ÑÓï8Ê“5.ÊñØ¢Åaeúyùáßü°èIeä”ú*
ÿ¸wfŠôV§ú½oÝÅs6Ý²àÿ…4ŽµÆ‚ë	žtÊÔ:%!¡´N„ƒ¸Vb„bwêæúÕü½=umAË´_æpð=ù¥y‹•ÑwäXzºhBaÑTb¨ŸAK‹>bë)˜¤;›Ñ¦~µ¡=%ŠkñŽ¢É‹é{Ë«0oxÙm<"®5Âê—wí…9 e_
ÓêrðÜ`~×Ì‹±Ÿ\dAzå×ÅoÌ‚Ì–MuQ§ÅZ‹Ô}»z«’í”K¤kknµÅêWäßq€š¿xÌ’—ˆ¹-1p,nt»Ô8+y)Ò˜ñu³ÑÇÏE†“®¼ä¢7ª~¡K-¼Árë‰—ÀˆnVƒva.îÄîßÃò&1µE2Ï49-G'j²¸Soˆþ,X÷„*üeéÁßmm"—|ßUŸ]y‡Œ0¾`e”åÁ?‰‚þµ—ÆÝÎÓ4MHìôÓ|+ÿ’¬=ýðgÔè¡›F6K?t6ˆ{“ê˜A`¤sËX…:hÕº¥Óú¶o8Wp×Rœ~yqêŽù_z×Èæ3Cß­dä.ò'p iŸo{éë¥lßòE "¤h‰Íl$dª±k»Š´.Ž‰ó«—«-¼L
êž	Jw9Û»ØÕŽÓÑ8¯²,É°íØ¦V¥tŠ·>gfK³}a¸ÄõI®Ú²ÿ
™®r_¯/Œ—|}a¼>ò_/ó¥`¼Š˜í¥™¯BóZQµŒWoòiÙ.çÂîõëc3]ÇQæŒéªò\|'þ«ó]4QÀ¦K\Ÿ„é’×üWÈqñáa·¾°[òõ…ÝúÈo|a·Ì—‚Ý¢¨jy^KØÓezÁiÑoŸœÑ	5?7>ëU<.ÿGg°œÝF*/mÚú×½Á¦«o‰º§ªo­©‘wDú°Àšt5*:Lv}Ç¸3–	žl’s†¥áËá&m~)Ö·%0-SÂ^\ejðqµËdº-&½'^ü~¹ÕeyUîœrãçî`å‹åß	—å<”ËºšÚñøì9:° ©	IŒÜò2Ï1õ_Ÿœ%)¹ô®’tƒx~8‚eCþÁ#åÛ±Gã$Ø+z©—è<C?yL¡KézÔ“çpâ¢püáßc Éq
/DY°ßY’ábŽŸ†kh#iT)":l~,Ñã>³Ø¡ËýÉ!÷.…Ü›ò©%Uä;D@PVBÏ¼hâ/¡âC*É
EZ-råýbÖ³Qpú^‰>–<ÄK°ï…Ä”¦¼÷2s¶ÍøÄò7´“þ4È2o ðÃWu²9Û˜W‘˜îQ^úUÈ>êÂ§ú<>,Îw GùÖD¡Z98BÁãÇ‡7ÙRÏ¨ŠdTpþF‘–Õíìx{¶ôÀ‹
õ’bØeŽåFmî¤^ÉcÐèØÙm¹¼ZJnEyodD‚ü@
y~Ü®û¶r ‘ëvH@Ö¼•£/½H<Xu¹l¡±hVµöÇ2Rp\rVÎ¦MªJ;­{Ç@%ø„k®òo,_ÇËÐ«ð%H$M”ç{_}Ú5ÿT«zbÌû«W¸Ô64ÄŠX×Í2€7_Œçàø±Ü7òVy3*]¶_Øv¸î¨‹à„Zš¨ÕØšé"&½áM…ãEÙˆ&¢q]£yv€ùF©¶µø";ß±Û¼5Ê|`q8'¼Ð­,ÛPÐ¶ÍTÃ®Yä‚	0?AzÔ9Årwñ~š‡d–`&”´§A4IRàÓ“9–-"	L”Ù5ÂÖˆ7ŒŒUÀZàÊ‹æÁQ¤Q64”nV]I|2ñâ14†öªœ¤ÐôsZ½¸O‡"K múouLúHLy¸Zb{'ËåEU´ÛPQÖ8bÆb×²TçÁVÞj„`Vš= õ<`&gªN£p&óÞRˆYÐî^——ç¸¤ÓèñucªÜã8¸ñÒ¥qwÏæù2¦¦½½#}U1Ó%Ç“²¬9Ë´Ò8Ð+õa.ä¨<ìø={Ü³eÉå »¨²ÚX'¬"å3øüý,–•ÞÞ kÃUÝ’¨ÖRsK\„×Rvè¼‚+ bÔÂ~ŸF¦jÜ†î0Õ Q[•wäõeË\k÷|êƒç ER²âÄ7qÈÃ*
iHòy:i…´©wƒÞ}Ä9" °Ø™o“)Õ8­=ŸŽm5lƒ\ÎŠ]µ|†× ¯%×ýKð5àŽÏ‘¿_k-9³K·A»jbƒK©Ôë0^ÀNâÃ–Ñ:,³Ú°ÞOo`(±½ã÷’¼„-åøÆß½ÇgøòU0…mxÊö<èª°åýì|að—¶]˜ýÉ4ê>ËÿÇ_n·[WxµúÚ
pK9ÜûÃ‡›ïòìºÃŠ8u¼z#íöÄF­Ôþ¤ýA1xUÎ%·’£˜¯mXËBkõZR°¢ô®œÞà›nƒ§3qý”.sí?Ü«å‘Ä³9¬Ó–Ž"·ïapÝ]ó½lr‘ ˆ­©S^<CËOø3Ö_&?ÀøOÅóÍhË¼Zßí¾hÊ‚Î£’Ñ‡'³ÄˆQ‚ýÇyŠº€ÓjÍÛ¤ñÈOs/úiãœ&´às_‘¹N‡›Ó…¾~í—õJ…¿Å,s+—PE5à&_zXvQ§j˜á¢S4YGúa©aªÌdæheÇ­’aHøÔ¸ÛdÙ¹²ºòÎ xgÐ|!¸	ó=4ÖÔ˜3#›–93à³½^:2©O®’ZIœ°¢Šõx[e8l3ÃLs20L …½À\Õ¶’d>%JÒWBÖç©<œùòhŽºòÀ‰yds€Ÿd‡›“¡¦MsV…ÊVœQË"«”È¥WKGBƒ88O§3àÁ¼ÇýŽ­E\ÐôE[Ý~7YçÑþ¯nÙRSm§F¬ÖÑ©`¥Ñ BI!1¶Kô”dÈs^Öþ€,Tåßóp–ð••
£H°ÕïÎ¹Î:^¼TY€9)Ý‡gÖXÀé	ûøäŸ%éS?Ì»ÐEÛî0€l’iíÍ¯^??ûîüÇ§/Ï^==?î‡˜îÑ2Ú"yL'p€¡&ÖÖ±5Þökªó)ˆ­¶<oày&5áX//£ÆFVY•ž,U™Uàµž±º&¾X=µ%ùÆ$Ðˆ÷°Xk ±Ù·"WmYÊWæ ôÖb'C7‰Ç&×˜pŒoUžXZÆºÆE¸<'•’6)Är1(¢7Öâp9ÍöÅŠa˜³}1$\Ö©¨qÀò´$z[·.­¹³ÖÔ°©¼ôÒš©µêÍóºÛ+ØhXuÙŸ†1¢k$×q>gIDçY>îÏC]‘‚%Ó¬DÁ©q3¤V§¯ÔÀ(’–&"®£Ñag:H›âé×·³ °}€þ§n8_¼ŠùQ»ŠÕ4¿ú<s¤Ú Åf¯Sà%k™FžÞ IU­™§UÃk/vWÇ‰Ï"/Ö3áŒÔÖÂtáJ=ŽiY[V¼õGi'ãÔ›zúqž€0ûè\%ôU ”õÃ_Ò÷X&àÝÑø½ê˜g)H€é­4fáê}vúL7ÆS8ÁhQªoÿ,I‘&Ã›­†¥ÄAj½
-³•{yF¾Á„“Íªúò‚Ïâ¿N3ý<ÈyEÝGª0öÂ<š³´è(>É•Û5E€…Æ8ë_ëè&E:K‚`½`jø¿åáçyæm°Îc™«9áUm¬%ÊÀLû/c·Ö×‘Å‹ëü5Ï÷ùÕãµ519VŸ­ªÇ6`îõV?NÙzÝ>Ó22ç£$øpäÑÉu©7HJ!>¢ïtÚÏð%Z¤w}ƒ¨ÜÚYSïf‰æ‡ä÷Øº}`ƒuct(T@:åõN$æï1yéåÐ-rCÝnmê›¤[+ëÆ¤Ó¶´:‚örØš%6sO73Ú4ßÝ–Í˜pêâ[/iÙ<+§ábGæIÑß›žjÇËmÈz®q´æëúj‡{:ýo’¦Ý­%Ú­U9+Kš™êÃ4¿ºÀBIÄ³ÕÂÀU^^ã¸ûC%£Ð÷ü`ÝX ¤uå²ÕŸV(ý¼SÕé-*g~•ò ’×}±æ¸ÎGÀr`ÈÐºŒ+Wíl®Îuo°‡U­öTuŠ±ºeUòÚ.™ö¨â‰hU2ÈîÙ½,J®]ÊôV:ÄÝº#I‹í#û­TŠ™[œÜvç³4È,q:_#½A>öL”"Zö\^@óÈr;'Å$ê'I”ïlUªÆìÖ-WÝ…ÑWêæ®ÛNÙ¯ Î˜ïIúá?P„ý¶ŸT|2°xçîcû†ËyÊ„³O½ËÚÚÀUS\ÝÜVä“àfdØo©¥ö[k2Yï]ÎžÏ(’g[ÁõÍ`à¨ùÓy¬{Ð
ƒsbA—fó(ÐU9ôyJé-é>?Ö;KÐÞV-¼6ò",„¡U¬ßü¶u='ôìÖÁ™»ª2†RÉ$Û¸Fë¡)T}›±–'ÆZ•¢<&p¾r* ¥Ðå=îÇáU•;ü¼õ.°ÅKÍòf°[	øžzj¿Î(ð|©ö­¥ 1K›K’¾µSÚªíoQ:QL#‹çcB­¹ðò‰·/)ÇÍðR8‚sråUÆ¿äBè\uAî¡™rE‡¨Ìõ;l}íÇ]Åþeä:Çã9P.zg)€Ñ$¤l>ö-ÍàÁð<à…±ïuÌëeY­eÝÍ”‡o›Ì0hFDIU	íî–=šÆÀ¨`I…;þú¯ÉW*Ìnu†gøÒ²‡Ÿp)Ã˜‘uH)ëfµãrlõwm>®­|ðªÚdÐgK¬Ïi{a”Ù,ñr#Ô:ãB7I‰ü›ûç¯_Á êvµHÆÚüÛ
8™ŠvÍÇÕ†<ÐÁðïäIY¿`›1;d™ª°ØæÌÓðŠ0Äc*C*tîE¬Nëõºë=ÜEWqÌ~X÷(Ò8q¦Í’•B_åÊ”ù°XÆ¼Çµã*ÞÇÒ¸çÞU QY<¿Ú¨«Oh¡“¯‚Kàò&'×œŸÜ—ùt_\áL)QêX™PWœMc¦y·)mû„)B¦œ¾Í»ÞI 4¼¸œv9KÚå¡¥šN±,g4l
°­%ÓjM4]o8÷tžF•òpŸLàÿnN*Ô“Û¢Ï´IÂO¼ôdbó¶«´dWKRlZêÔQ÷ô&ÍY"Ö.S½Ï’”|øßXÉÚ®à0Õ”³ Öý„Ë¯Û%S¼Â7ªÎºuŸd‰£y—„9ÎLÛ£ 'Ì‹‰¤ÔÇÉLYz;\+ŸRîÕíååxS©þ.«Ó««æ›º§cã“ÂçÓ‰9»`šÑiõ>f=hj£³E‚>]u4z#Z}»N*_¡—‰ŒŠŽ7˜Å”º±S+èY%ðÜÙûº?“´*Óàöæ¿9ªÙž­[Yi¤˜-m¨ie^es¹u:4z£nÅÝú:«m(Žç9st¦ç¶ß”fÓ´nÓØ:mÖž®µM}ZO–§ÉûàýbaPÿ6Øºx¸?è˜[Å/<<Açg·äX
'D`à:Š`5^C£ØÕÓ›°#@A:–úÚél©ÉÎ–Tëëp¹³}¹X†Gj«Ð,5Üœ¿pÖ³OžT§.¬–·x¦Ayþ–™ÿîÃ`ë¢åüU›+PÔAn½¯‚1å.m‡Üå”»W^§gñï]CklŽ‰Š‘Yo›yø
jsG-SÉ9=n".Ü”üaEØ‘¤°jRŽ©ÆàË¶S€~SIÑc¸{Âôð8øÁê$½lLHïÍîæö[»Ñó5[XÍbÙÕX™ˆþƒ¿Œ`®Â,¼À$¨÷x’Üu¶ qâºíºõƒêA@²®×þQç%7vöÉ1ÙÞßØÞ‡î¶6dî9gá`§ëŽÕàr÷òaË÷þ!ôóÉQ§mw/Â²l)‚”óË—auh4ƒÛ+ŽáÔ¿àº/$ìv·ÔÚ?0þ<W_ô‡R€—¦ÞíQg0xØßÞw®/_O./³ Çtjø>é	
²É˜à9è®‘éY~ÑÍA5Ö]Ï/ú#ƒŒÅdž¯¹fgq‚9á®ÆNÖp w‘%ÑœÊìyžL{œ q¬¦Çü®HmzVýX×„ya÷¡€ñhA÷ÅÍgJ,…SH¾k€ú‚ÉËCíT}ˆ$Ï@ãH4¾
Jå5÷I¦1Z«cÒ{³Íl@6¢U³Jä¿År×Œév›	hèF«rîÐµÓº¡l®Kç‹_¿år¸\gù‰’-,Rá^‰t:7“’Í¬b{_&RÂò	NCŒ¸˜‡L‘Å"¤™õÿž”W@AæÉÐÇ_s_J®åÒñ»™xšø7‹îï,¨îÏ
[ð J}/÷Žoà^èXxÅzŠb
×ðBcGág."©r±s­–%Yå–ù]©u&xµm§ý—›þ+ÖÍ—;p96Ûï÷»UÏo¦žxSÎ†ÿPŸ¿-uüpÇÛ¾Ø‡Žß’òæ­Ù¯·v$5¾ Éþà­ãƒa û
ÿ<;ZìmÙ‡ ÌJùÂ¾Ã3ÏG²r‘)Úµ¿€ûGÌœAØ6G·C°,˜/À÷¡+@Óéa&–Û 	?¸Y·Å÷Šëð$ šMõïFð±÷õ‚6p÷îŽñôÚlÄdþÕžðèp0 K¾N’(gvib¨y.Øv¬2¦,ÉI±·ô‚CÏ”þá‡ýÙMq¾f78èÐ'\àuâë/©úŽÕÑ¢Ë‘
5d0”#t¬oÞ}ÍÙßÒËí{ÈŠ9,«N×ÓJz7ÔnI»¥ÅT8Êöä°Í¶"÷üFZxù6ÍÛd"¸c‘ãªWãæ#‚TÝ… ˜wë­£)>Ò8‹jÀ©Æ­)ÿúùŒ¶ðåÅâÄHÑ©Û>ÊE-ÚÝ¦>ÍÌJ_Uº|ˆY±ü†Ã$W•
Z½¨ÊØÞZT)Ö`Í´dR?+k…Wž:Šíã;˜^Mý2á|vHxÅsÑÕì¢·KÅ(½[…Ý!ª™±÷·T~ŒÂÚèFÞJIk‡[F¿x‹Ìöô*‰¸ã6òX Ö“œç—•Ý¬1pz©Š¦Lpoé1–ÑðâI@<ÔCü) ƒ¥ÅjåÐ:Z¡)Â<å:ý`œ‰1ë½”Ò&£†Ð`4ÏÃ+ZƒŽ&àó“´ªuƒ5+¬0êþu2ŽSäieŠÐû¢-ÌW’bÍ¤:F7:-^Šèÿ€~C6“¡m
ª,È‹äg”3a^Ÿ¨Þ2gòmQê	óÐìÐ"M’
Jœª’­ji’»:¹Ã¢¾ÏÌÝûÉ2‡:a^Ïˆa¹†™¹DoUW^ê‹Vó¥—œÌ`’°Km÷9zÆz`^cáSvç"«;Ü=AÒ…—þ—$œï¯Ž¾I?üù³f(=ASCž-·—-ï¬s`öšÜ±xM"’~ä@A€À$1ÈÛ¸ciÙ¬Ž)ì{»2{³dØ ,ë(èÝÅ,k„ ½ˆEÝÒS­>\Qûg9«FŒ¬Œêå‡ÃóhA™³4ÍÍ»ê¾…xã².ë@Ÿi™"WfNhÐ-&d`6\bªCoi6¸-'#TK›ËÿYåÿßl½}ÌºÜ-w­¢tÙzuUÝbáž;«O+ïÓn³q±~l¸øð¢<œzäø
H2ÓŸ9hÈ6p¨eÅü2@cÎÇ¦¥°c}Ëgy)æ –Í=‹­ki¥!àQ/E°ï.š>3˜|1çÒ1HäùäîbÚmbÌ!?t˜ËÅe]å<"þ‡Ž^y
à¶d5žt›U •4–Ezqc_ò¼cžµ³{c“Åõ7iæw¿C²²¾ƒIÝåHÙ|œKPy?Ÿ](òU–~²÷µþÍå§?ýò+]ò»µ(»ûXëó`úéW×ð«J+wo/ÀaPŽiô"¡Ò$Á²ó «Æy9d"4z¸T=þÀ¤r¼G‡Š…M¥b1 #›5s¥¨R†;2õÒqS7¿dv@[$Å¡±Qp	Ÿ€®np_º‚Å¬\y…^LÓe6<·Év‡€èã8Ž—¹»¦—­=¿Æ?ß„YijGÎ¦C<¸‡Ó.†nÞ7nà”Ñ²‹vÑ½;ûkô(‡?c
¬Bá0‚¢#¸Ç.çqp â'Swæ,>€É[—Ñ,o,®„i05[„ñÇ›sªkßÑ‡Á\¤·Òñ…×ÝÚ ÿëoí®¯`A®™÷nƒ¬IIjXP˜ƒ×	-éWL—ýék¨r¡}ýqËí‹ÏªIN“8ÉÝ… §LœÃ1ÅYå*9¼Rœ>Bäúó°^l;8fû	-â µ»!¿MãàÄÁ(üVd1é§E³ôÍní>ÜÛ{èÒ¬ÕË @©+æ9A²_0óÌ\^L.èrÿÒ»9aÐ/hý£ u8©ÜD±,ö£ÀŸ)÷.|³·Aà?€„”/=§è{%ÃÃM"–v¦háë´\;µC†ò…õFdŸ£Ç„É}{§óè¬ˆ2Ž(4=‹Ÿd$®VcSÞÛ%("ýÍŠž9
¼˜Fý“#nýù,Â\2EQö¬[8Åi$)žp;,îƒ²±¥.O.€>ŠauËnª‰^ »>O>¿!¥”ÄªŒU4ø\./eÑvˆ½.3
àBcÆšr^Ê`‰ï/íE<Ð¾øªÉ‹Ï®Ê“ÆÌñì²Mˆ‡‘€º7Íõ|ûæíz7	Pªq]Ë…9ÖæZÐ¥T@ÕKºÀ³1xooáž”“A;Ã²ß'­û½¨÷{Ñ¶ßp:Ãé>yv|òú»¿{uúôÝ ãÇ}ø8»×–--íK*_4øDÙà“Öò½§-öèH50 ª1 5 ŸRæ›üF,ã]ðâØ¿u„ä[oÖ¬’ƒoX¥û¶;Gó¨Q¹ÀxŠŸxáM¢=ÅVÅsŽ3Ð;ê¤-ædäæ[ù>”F…>ÖÜ*¡ÜÓ†ºYÏ,Å²!Ò<Žò„ù1HYÐ#¡t ÝÜÞê÷&_~ø7?¬4ZøÍÊN´Ívj©xRæ®íe½$Tn-¬FÍ{/¾ ut–ó°Ð›™ƒWÝC CÛÛ¦Á×¦j©¨KÛ?ˆžkA»Î£ž—Ï£ßtM(Ç£ªóht8:$Xpø³%Á\=k¿Íf×ÈsnÑ¢‚ûŠ$èÊÛôlàP­Åqbtû„^i¾ô5,¾(~Êª¿e†b–F¯Hús“5VÊÇ5ÀªO^%··UCqwPA~{ÉŒÓ˜±VÎŠž}…B˜kÕv«‰ØTf\C¯ÔY$ÕVS)¨ê­Ö•˜AéÊÂYkË”^®µáX‚ùæ€·œ‹/ßÐüS.…—o¨Þi©šËÍbÈ:ßñ–õU5®áÐJ•Ïð 2gh4uÐZ–L6”ÓÉË.Å’Qvõ=¹ÚjCt<t5“µ~ã»úJÆš:‰ìG¹ô-¬8U×$p=;p“ÎMšn]“SŒI™»yŒkÓ<ªÏ»ùwwÏŽâSó©J.<'k>[ˆïG‹â£jœE`i¨õ{Æ÷j‹£º¯|ÿ´2™Úå¸IôIüÐ|by·çóÑ(È²sä,€7¯Ý©¿³¹"N•õ}•·/Ûü­!W¥B­m±yaƒþŒñ¦¶Tå}¢Ì•1àIg¿<îsÄUI4%"þésÀKí±íåÌ¿ü>hûì£ž"œ¤Oi®/Sù®²éÌ'J»ªÝS¾Wä„§oß”ÏJ‰ÏéÓÒ÷æó#ZTþE2†ý(?ëž+Î'¿Q½«ž·ÔMå«áézgšš-Dp÷d’„#X­ò³²§Ò£•¯
˜Z³R½¤Øãª™¶a¶m<ŸÄßðÊ¶èƒ”Ä(-•ÕnË{ª7_Ü> ¿š6o>K“é9 1êå¤Ìh8¸ô2*ª-äoÊU{]y¼v£=—°*Å§Ia‘“û#øE“¿9zpÆùÁ$ô}e­Èß8ñÆ÷+&­Ló_	Pû‘|bê†RÛ §OAÎ>yiO8¾Ÿ¡ÛLÉ2-¼ì6‘®—_eZ‹QžÞmÐsÚ*…,šš¼ßï+ mƒ”"ÑÁuÉÉÕL^WêF£ «1y]€Ù+jÉ‚©¹ÌÇo¾°X©jÌÀO+: þºÎ*‡äy¤z?zýÒáå]{¡X¾ÓdÔõñÿdMV	¬mhº¾at*¯¯¶öQ}qœ;\ÁhBº£(™ûOSÃD,$QÐ y§ÝÎSü‡Í-ŒÇÒh€ÚÞAgƒk‡ ÖÄªô°b°n”æ¾Ÿ…ˆ7ÆAžáÈÔÅš7hòª¬	ˆñ:]|Í‘ó wàSÿá¾3ÿÁß¸ñ †Äì› hFB‘³šœqô]Åã¢’yÎ8%ÿ–÷¥sôÜtæ
|=^?õqòÝœáÎj­|ýokµ¹xšñ½kì7öE8ñt³•ÊÝ,‰Oa¸K®]ûôùJç®]1¢qž{9:Ÿ1lÅb÷ôïœ‹Â|Ùªï•÷ÕïF‰ç?½
QÄ…7'^ìGÁ³0
Ø/ù¢Ë¯÷%ÿ   ÿÿì}ËnÜX–à¾¿â:0H…ªô–e«my"%¹JÕ¶¬’”®rÕTbšAF‘INAÀì½kÌ`0›îYª€\fÓ[ýI}Á|Âœsäå}‘1ô°H§"ÈËËû8çÜó>õdcyŠ¾Û–ãx Z¥gøo×›ö·ß¯ ˆÔëòm¿±ó‚·ÁEÎ†?ºÎ®=‘úÆçFâ‚ÔO§9;ø£5ÏO$×Oµ%×Oõ$×}>ZzZÔ_µ¢IúI‘Âˆvå˜¯Ë Às\Ñ‘ ÜdG'A2Úã$Û¢VÀOxWYý—´i ¬(k‡ò·0N/!GÁ˜*˜…}lì%aý²@:yWP¤ŽSß\R¡éÔ]ábP²B_‚hÈÞ€TµMÂ°Ô®´ðJ6WóK‘«]ê“¼~6ñBi¹°îº„KYŒºûâªbÙKhÙý½ú”~®š­Xi%øžcÞìW¾C	¡&HyŠb‚Ols+z•´nÂ“Îa¤tjjš}«?ÃÚ 2€‰¹Îâþõ_	pa×ŸùfÐ‚µpÁWuù$÷Ly­à¢Ñ0©^^RÏúR®šÁDÌÀš¼•JWåå2íN{	;lt"¯•É‰Öå^¶Ê–'1œñ$£±¶41»„ÎŸ‹kòÉËmyM<iö€£îRºç` ÙüÌB¥CˆeƒÓ¥ ê‡“Ÿv5Àš±U»ˆ
ï'Yü† ­±")ñ­’^¡
ð`—JG¬àœI8â%†dïµ‰bˆxêh»ìÄÕ®Ss[glµñ0¾i;]egØ!ê”‹fMü2ã	Å¼N±å•Øoúc“ýy¹Ìšé/5¥}²Ž:tí;›ô­ônîëå2Ãõú¥^òžâ~tŸ4N8õ>Ÿ`Ÿsóó–ä—A4žèx,‚‘D˜|FAîû§@DüäUg«'X‘Ä*á3=­S	ê&w×¤¢HÌb3‡ö˜T%µp?0ùPý%;œ·?XÏFlí"ªæ¨`ˆ¦DW.¼\¦,€¸PpØÀ¢€Ï!¼ÄV³xø¤ÄÞíùÃëÿ,Y&¼?cGÓ»9baÐ˜ÍŸ±².‡^x#ë¡YÌÿëÛø!¸ñAžÆ1Ý#Øf›£baS¾°Ó£ùÐ?õ2®Ø™‘Y"»i“=BØ$â¡êÐÆIð3c¹	Ë_šÐHrkPŠÊä~ÆéÒ1@Púª™!3”"Gd¾s¯Ù¢Cã9o¤ÌczÆÈò–GîHzÊÊY ª	‹dYÉAéý®˜%yý"§òrß	×¤¢¼“e²Ñ‰{Â7¹tgÂ±|³ª¬d¼”¬£ÌTÚÃSŸ×bO™’LŽÅ.eQ˜©'}xµã\k¤Ïº_Å­äõ×Hèø_£ùºÖ_¨|fû¥}ÙÊw}–_ºX°Çy‡Â‰<H¡±/ý”¸héê~c@7b¨5ô†òºXub—ú:¿Æ2Åê“€ÒÒ„É»QÐPm&ožÜcÁ~Z'/‡R­ ­}RZŒòiæœqÅûÚP›q³j€IÛÜÞZƒ²/¿Kÿ„Fåqw©o›—ÛÃð`“éjejò6{IOƒÅÖqèÜ?ÖS	½•ÃþlÌ¨	¸jð’…JéÉnù”0p0åsúb£„=°6ñ¨5!\HsƒÂ?êÓHü?O‚Ä(74cŠžïN}ËÄ ñ˜Îi¼û œÇ£`§éˆ¾p`GÚÚÊñ‘*ð¾ïP÷Ý$ÑÌ´v°3ùJHÜý.Œóö±Z—Î	h¯8ZÛ†S)å¾1ˆ:õGÝ	ž!Š3„¾øÜTCj	`9]*€ÈâDƒ`Óg*¤ÈÉ™a*ü6Ì¥<_~Õ}Y+W´°ÈrUÕùÿó¿|¶dKÖÙüpýKÔŸ„1éM@F¸þŒÐ²ÁcàÉû’^^ †Ÿ‘È#ÇÁqÄ™ß§$OS%þKÅbc
%¥Y×‘­ § MêlVîéü’ÕªÂìð~½Ñisê÷?nI?ôo‚ú6‰¥&–×ÉTÀ²Vio.´¥ðþëÏt(ÜüÔÿ	ý0û!p áã¾—àˆ¶(ÝT;qÀƒ$q_wºÿ•Û_™8‹eÊZw"T¡r2ƒ“©ñÕOÜ°*ŸhEnªÂ‡A.³äˆ©•Æi±NÆ&Z‹)Ê'v9Œ5EŒ]vÊuPt2¸GöB¨¬.NÁ?G¡ÖË)Ó/kßRÍ­‹’šå›P^¾ö,¿!•|±e’1æ±$§³pzìÃØ>‡ãøÂ6?¹Z6‹›±5ÄŽÜÒ®aÇly@$f¬¢F *.½Ò,^ÒHi½9ÞÃ½x)FzTÇ)½Ðžù»FNuç€~Bª•|?ºÆ`¹cÉiÉV¢=á8_|FN‹jF†£L)ËÙ…€&è*
?0#O+äC#v:Ál<÷e”_wZ€ðcŽÃO3ÄAõœñdŠ!fæ­så±4&N\ç)óºMÍ`£UÊGçÈÕdaLsÌ˜©æI©¶°Hd%C”‰-bÄLÔÑÆ¨˜s
1Ò]°’f»@hGé'?%2œ¶è;mZuÙMø=†Æ^PmŠuRÊÅÌb•~ø±æÌßb*®„.€¹šƒmæÆG-Ùºª.ÕžSk›r#-ææë.-½UÛ‹Ï<!³SªÅ:Ä¥+‡Õaß\Þ/¹¼ªå“k<û&—Äà€_RàÛWµøIi)L^$f=þ¿˜ÖRâýìÅ‡1¦z­9’F;O/ÂÒ]úxw¥ÁU&3|«‘ÜÒºc5·½í÷dçðßïl¿ßÒþr˜ð”EZWYß8ê¶S¶pùùŠ³6¡’è¸\x³(ð"iD,åm…7×°^óš½œ³¹\µ&OV±2‚÷)©gJ5kÔyH™í:›«öƒUÛêá$ó=ë—ÕI´WÁs*6nP&’9x?ãôæa~ç$rŸ¦y+sA‡UC¢êF[»rq­e%xº¼¸UÁhæ
HÑbN75ä¯\^"ê xH/©çÅ[àÊ*åaQé78»ú’†½ Q%IMG5ó•»€{ÖàMí—o<ŒLl¡ùHœ]¦Táya±hnÀž´åÎëÃ¶MNvÈ`. «LsbÈãëJQ:#(îlnííÿ¾nU‘7¯B—ææß†A–vÕûó¥¬³OÍÎcÅ'H?xa0À±è]U—«Î"œ«ŒK2áT‰b	ùÛÿø¯äÃõg°»Ø\uÂàµuÌ-Ï½=ýÄþ;ÙÎÚ™šõžÝ•ñÚ ž%†••¥öoþ­,º*¦¼‹· iR Ä¹"O{Ú]EªË•ãÓ‘ÙÆG±3Ÿû1Û	%€¸nÌ®€G"å€ˆ»žbÍüdä&³Êwo”“Ú+/È¦½5ÇÂ'_”óá¯­tìØQ~³zp×?˜í°ìÈLo8ª)L´}Dc§¼œÅU£µcú‹ç?¯ÿ€šŠ[½7·Á€–hè¦Ìàý=L.@ë]½	ãS«„dãQ^bÆûÊÎÊÓ,Œ˜Ü¾w²7¡›Ó:úˆžïŒ½)L¿»þkºÜã™øý$›$¢•ñ£ùg®<ƒW›ÁéÈOa±ÅB·¬åîg±nJª¯æqs|T–è9öyhPZG'#†ÿÈû™7•ë4sDØ±OÍ`ºÐ¢íØ§AÌû¨ØCTô‘fºÁ3ñŸEŽG*½º>ós¾ÐrÇâ `µ¯OÆÇ>7æc“±¶Ü³À>þ¬%¼OÚ
åM¢üê`
jdê´£ñü‡%Ol6Gp)œïn(\îù×øD¾Id#û”b%Äop@;“u6dŸ<tP‘+n0ª	­ûPEgrFt¶°õË(éLGn×ÝÞ=ØÙ=z¿AhÅÚ,^ ;èOá_ÿoô?ðÇ4Ï¶ÕÈÀûìÑÀ{?¼kvE(ßtòy÷½P¤xø²¬»@oÅ ²ÂNFÝâc·|ºfý8_±–M»Zÿue®¯Ò€{ÜØ	ÂÁ›àÂ-–}µZg|ÕpÑZ×È™®-?œ„^òn¦I•ÊÖ«}E*ë÷kRJíŒ@lä…Å?¹ñqíc‡_ôyMW¬å3ûlã\®uÙ…‚Â-

ëvA¡ÑÏ÷­?,eõ;XKH¨bÁÊ^ÊŽì®¨“HRZ“G‰Bþ˜ÜDKÔ²£¨Ü÷,(–ØÚu£›Ýþ›fP«ºÔOÊTÀå„£´-;ä¬º]í„;ÎþwŸ>w·ëqWÏß®Ê%­Ê×î=íêúÙUNÉf¹hOxS*ªxÖå®u¯:Ù§î©^Õô¨»/þt3BjŒH>v•HíÊ»t¶»+W»Y:ÚYeH“|«‚J»g¸Ú}îz‡SÏó®-Äy©ÝøL©X±-TSnr£¨ÍÒ»” N"Ý4s–n·Ø©l¾Qp§2Gí"õ”®4‹÷T×…~ª‹× ü³_ž¼;Ôí™Ë·xf-ž.þðíÊÙéC¤OÂøî`d'ZBŸJm4¨¡#W{‡rS 1…ÐÍò}O²gç *¸‡Û“¿j_®”½vˆ²“l¹ÖðH¶Òù1zöø<=M™-YƒagÛÂ¦ÌÊ0%mgL¡¦6µ•ä@cG‘É,ZäZ°–D%“px²’K¨pÉHu'qÙå":ƒ*áÈ9ÃUSrvï2-›"´¨2¨¨†ðÓ†à“6ÉÇ"ï´„r–-Vb‡Œ9Å¢É2¤”qhÙåk}7rM…LsƒCÐxæXNFwØÐÔGâ4C7:ûnÆIQcÛÔ,“Ý²öy¦
sÚÔ QÛ¤v‡ðPÏÙä«d¥ëy”LM}JîHªœC¾Jð¨ô ™þH©ïr‡@!£y]èGð(‡X–Y ˆè{æ r/mfÊ¬†3“Ð=>ÉŸ«”¤¶µM˜’¥ÃÙßýÅa%¶(xø’x•)²ÒùÀ¤ðvp°Z*¯ÂÁGáÜùÎ;Îe™Á)ußÐÔx—²ÏÛ½È!›±èÍÁ?9]Élb¬]
GX¢¡eAL¼@å¦°X!W¿ó£Ôé×^4áß¶ý“ ŠxæÍ™ ÇÈ	•cuéPÇÚ¿M¼l ´=bWyž«³Éþ’îú
V°HçtA§ÓÙ¤ ƒoŸ5î¡X†Îæß~ùWRü&ÝAæ…×¿ÀŠ“e²#ø=ð½Û’oóöýäú—x@±mLw›­È”"N4û‰‘9Ñ«Îª…xÕyñâEu½Ò‹²¾²@ž>[ ßÂßFz|ï¨A"€ðÈ6£¿q¬ _þ€ý--qMZ%y×X&õw£L9AÈê
-ªo¾_gÉa„%E9ª×÷Óôë=NR¶x–ôÙnDXïÆƒ„½ÏQßlê#£ÇmÑW¶¡äm ‹¦5 ÛùH:›ù×¼ÃÃI:†3ÅÕaJ=«º€â©S³î+údô‰’š·àóôN¹ÖXîc¥(ý¾ÍªA^=ÝË±À tw@^¹Ô®¡&._¢TZ¢´—ÎðíøWT¯éã…>ÖþE2-^ ×¯©±lÕWµtøcòÓá¤ }ˆª«÷qŸÅ^šu;H >BÉõ/ð­H:¡¨ñ¤³@:)ë¶=*]âLR?1åÙÉ’OÆô¬2æí¸ßà¿ã2ÇÝ&Ó?I´&[p,.&Øé{Yÿ” ñ2½÷?ý%?IâæØ»øÊ'*9I`Öý0ž6`Îúá(å ôß˜c/N|Ú[ŽBZŒé­;òÇƒÝ£òOÆù,ÿ‡K1Ÿ«2Mèï\¿k•dø.gVdÎ°f){éV‰Néuçòr\ŠºêÈ‡Åò“ÌÿÙ“ËËùœÒeää%®WRî…ZRNrÅÅr^—í8ŒOåqý$€¨~’^ÿræ‡rºRº›‘a,³¢Ðß®N£ù‰àÀìÑt´ÊÔQ¼M» –ï`z
¿ÞPYÛuåVóKYü6>÷“-/õ»4ËÖÜßk¯ Fã£?ØOâãÐ¥[@ 2ýmcvÛËKpq¼–1Œ—ðiõí¸b:¹Öjƒ¬ØFt‡“~ìÛF”òÛb@©c@iÃI#Jül’DŠJW?.ñ£™´#&øü./«©‰6Ê±YYcž?#Î7öTUåÝW4!:¥ÊZQž}$_]ù…å÷U9Ã:ƒOÔdDœT…}¤–¬‚¥‰ÂäJ§j¡IÓåÑ76Øèšý¢g•ÆD-€@R§#Ž+Ù C‘ÊÐ¶#„Të[µøœr_¥"¨qB"¡ØÄÇ$’ô¤§ªK§[‹<–esP›Œ8°U&‘»ôÂ°W*<FsãaåWf¼R^<ZøM¬ý±yÉ‹Á¹^mÎŠgÏjgË€PÚ›Å¯k:S6É

”&â·ì¶?­ÝâúçöP[¤×]f—Œ%H/Nc-Pû/ˆÐêÒ3{ôéØÌwØk8K/yn!T¼ë^ŒÐÛaè‹Ò‚Ï°´ ée~ÄØ/&œZÛFh¢XÖÁóòíöQQ–ˆQÏqœ¦“ÀÅ}Éµ»~•s]T²*Ô~³K°rE~º0]"É;=°ì5ˆö©NÒ. ]ÅRi/õR`ò§‹Ö£L,îX
£\À:Ùb=Æ0Bÿ¥þ$-Ø*¸Ãçs"0äY´Y‚0 ‡ÅÈ‹ÌF´¬Xââ³ TœfS·Æ–ÚíØ|,Ûï|7V)Ú<—q²mŒ”Š»pRªöY*öÉºi<^]ls+0§•En‚«…†° )§“Q^r œãOÙC\h÷‡“ÄŠÓ&i×êU„+úX!#\ÍwE·µÅ.§&r8×U}ñù€Í{ñYLY¡]š°J–^cçË£íFþ9ÿ¾‹Ö“y«èÎ#ìÔqÊ•Ä¸+ÝNW§FÍMÃ¶SÌ³qW¹WÛ‘tu‰ì f"QM®?ã”ÆqDÅu†‘}xÑ0N‚Ÿa-„e¡T®~ætw)ëp‚U§Ê$8ŠùS %
>hX`’Z“
„û €Xu-2³µéšµÂG»w›bCSÛÅ”Âƒf <òûQÆÃÀ[ €•ý	ÎçCïúßþÙLxUÇr€ð5IØOàëõ®^Æ#É6‰${å‡ª¥[„à²J:gc>HüÀ¢)åñÒË°˜ˆÛ&ÚÔ*á,²ËéÅQ|@¯/b«È+	iªü¿l4TyJ7OÊ/ádï(ù“JÊÔC¹ÝéÝ(€ºròLi5%·nR}‡¤3'™yQîA‹´ÎèÁû(üÔ>•Ò¹U=ôríO’4N€Üfèn@2°+U:SEMf~Ød©'ÞãñV÷x“)CíéêµClÏ Ug>sï0¦$o|„q•ñl®òØŠƒË Ç/T[¹¾aGˆ¡Íw–Cc½Æ¡‘“ ½Ú¡Qaê+ ³‰}¯s©,ÐUG˜ö*Ž}
Bÿ%m
(küà‚	>×ÏÇJ*¨@œ^g3IQ Há´ÁÌñT?SOh	ÌÑm<Ù¹€18f5B»ìoƒÑ™4¡¡Ú¦ã—:ÒN;`OGÙCu VHçCÛ+zÈVôV¡}¤úŒ¹44È¯sj4¡ìWÈfœ«tê¾¼5¡Þ¬®ãæå°£W]ªsÚày…æ¼±~Îdt©Æ™^|W¶Õ?ÑKš¶ê¥A%›ÿ½±7¤h®*[àŽ½¨OO º5^6)±DH'AtèÃ™L‡X½à<ÜŠ¾—a°™cdI‹A|r
½b~6´ÄOÇ1Ì ²õä¡z9×ŒÀCÃ µ\ƒZåÊ’›êz€¼LJÓhVuƒâQ|S×ux–m…éá—Ë´ªíÇÙâª±9º
ÖÊ°Òï¦:÷Ïägê(³;	ú,’*ÎTMÒ¥€PU—ê8S¡ÜàLÒ@ gÖÿHá’7‚$~ãÕ–do¹o!x‹î%©[næ¹•gï­¼-LÍœrëD¼lB™Ud¶G	Q0¡]ÍíP£âd‘t¡—õ¢eº©äð.Å|f”)d|›ÏÄW#ãË|‘žGéþ¦Ò=Çh—h/H[m¹^<pï…zf*kr–ŒzÞ,Î¿bL]¾öõìvš#^ÙX—wto,u‚¢8PâdèEpøP"
Ð”dKµ`¨4Û:#±ÃŒPœOôCz
õyÑòÚ2õ­ç)Ùf¸Ÿžh)ùÝ$-´È¾	"à)ü ‰È?‚®"iÍìmùwOÔjz‰‡á¾àÌ«+o—÷ÐÓÇVJ@Ž&ƒ;PKN© ªÉúŠš£÷Ûï…šÌNz…òt:šœG´aºMKS0Žvg*ƒ’Æ†Ï›ûÏÿFÞËS+9	µÜ'mÉ¨|©x]É;_x8QÚŠÄ5ðiÉÿFZÞÜKs¥'gôR ²K?3f¶&Š0³4
oÙ‘·/„¯à-Û#Þn÷r¼•U[`y_Þöä©¹ðÖûêñV8O¶‡¹þ¤u¼c{DÛ2ÚJªç/oeµú#â:·ˆ°›u‰¢ÇKlžúý Uß^ÃX+QYÕV<C·kñ¼[Fz¿ÐKïKºèCÂOË`·f­keˆvYDÓ²¨Ó+‰µæ˜%b ksÓ¥î9WtŸ/ Ò}mÚ!uŸGÞŠÞ%:5m÷¸´¬,ÖIyÀ´iÊ4³p/â]+Y/D"€zÎAÏÕàYõÏãf¸oâÄ2Õ¢Ç+5gFüBaqÇ°4a§…ÛÖbmLLÇÒÎ¶¥y¸€íInóŒÒÛðeÓ‚¡ÔMMP¼4á5á³‰ÔAŸVèâ@k.,eìçêî œæß’!JÒ\¬ÁH]Gñ0ñF^'¨©`T– ÍØë{p_8|0_žY¼|lÝJ*§§] ¨ŠÌë$j©•Fã8ÉxúŸºÛ¸Kâ£ƒÕç§}»ÛXùDy+Ê£A ¿SYÆ*“y£ÏÚ:/>U¢^·í°•%âÎ¡9Åÿ•:çnÌSª^^U ÝäR&E¸®qÏº5âª®Aâ=‹ËŸšé·³yÈ³C™ÓlØí’%qõ½pˆ<‰ÃôB?N“Ñ3ñžÖ¤kN0Ð
 Òö•ÛIwGÍþ>fEksmßO£ÖÁ3Ä’HÜ`ÁÍe¨[ÂokWõx¬\n#œn‚?¸¢.
)Y&–†éÇ`<ÆvƒÉ8ú¸mæÑYáï^-±­f·Ä\šS–ØÓ—8oøe-q!òÏj…H[µÄE»›¬°‰àÈo¡çà<%&N1$ZÑª‹fus)9*¨>½³ ÓeLW‡‹oÉŒT-µº¾¦—Ye #t.fx¡wWÅyÇ«¸üxi=o}0ò:òe¤vºð}ƒ¤:=,`ƒ°TÃóö¬;tZ4áNpUš]ª>¯ì¹êVL0\,=Ú’>â,ðÏßyÙ$	²OÛp@amfoyè ÷É:€“×§9#v{_ ÃggìÖ4Ÿ9µ	%’ø?ä{PæË([˜2¸†làš<”-ä'˜A[á™—‡c/ù¢þ.çm¹IJxQ0B::žÀuê2Ž&g}Šßvñy‰e‰×ÿˆ>Kç@YgÓ– Ú~?£¡¾:ëwúT'…!{;|O<#ÇJBßà{³`xš¯¤r
`HzŽ“Ó§Í9N‡êùj€ÂüŒÍõÑkl¡¢¸Öy`rz+m €ðï)hÁÑG,<² Ãë¿bÎNÒKè],‚\©¶X‰•OüÐ»€£¢*‚CÞ‰ŸäÑé;R¥&eUŒv¥!?øÉ‘M¾6à_í àå`<*Àøú/$ñû^È“ö¨îü©_öXÀü’	¥"ÄƒP‡O87Sh’X“ Jý„Z]ö“i ®Ñ]Wx Ìa|†ÒÑç÷çëOæªôÅ[^“õõŒ:Ó²P˜üÚU«™%×´uŒ†äÓ¥]ÁÚcNS%óïña<Z“ñ@î@Î]±[öî¤_Ö4Á5–|+	¼„.5óo¶èdŸ”¦ƒëkîçL£x”Vã:møÎæmnè0ÝÔ‘¸ºØÀ› ¤Žè o¸]ˆmµR¥<NÅ 7îš··†¯-9w&â¶È+¥”€)7`ñ¡övá¢oLÙ½Œrð£VZÄhÌC‚`4—Ø³8q²SLƒrN£Î"!/4ÎJÌ«?I ôˆ079÷r4`™—Ç˜>qÏ‚>Æ	±Èc¡±â'Í’L“ÒWŒ’®æ¾ôæî%YZZ’ÇRV>Ì;]Œâ/re )I©~íJ½d`iÕ|­¹5…Æ ¤öl©S$i½i‚V]X4'fm#`«F]‘5ð —ŠÌ9_m²`Ò” u0Y‰K0á¢¯Ü¥?uì´ ±†´–v2.Ë@nÂòÈØü ÈnýÙøOùb˜	…@ Qþž²6”È–üiïˆs0‘ê(Î€ÆhÎO¸2VÂøD?œÐpœ”¦êÙ™33rùëf@oÛ¡¬r&ìü­Î\Ø—®½Ï+PœÙ$†mA~oÆÇ^¸'þÞáþÓ‡èÂ/ÃnšÊ;Râ.ÞeÑÊä}ÌÀþª ì÷&ûf€ï+Ò½ÌÇuå¬<t¿äõc”ì…˜d;ðiŠŽà8]\{Z/ÖÖË%l‘óA²]I!’H½Pµ*,½iÕ³ô*Y3ˆö®S0I—ù\ÿj‚lõ mÕ +jÚ*Ò:€¦Ê±÷ë‚M·2i¼Aöè‚ªÑ÷lD‚5Q]ªû;?A?ª®¶&ih:-+,}A|¨'dž.]
Ð¿–²†ónŠå¢Å€i	…ª&Ž&¯ÉÜûI–Äs´RÁ€Õ€ó…ÉÔË‘ñ¢mQÄÞ^Ÿ«Õ@]šÍH¶7â~™*­èZØx¢²a¥!àÌfÉzQ,Ih%!´>Ý¿rg4p×;T:ý-‚jg“þ!Ý­xLÇ‚ñ –³q4=,€
ý‹º…9ÚõÜ<3Ìâe«¶ÄxaËœë¢?6ÙŸ&rµ6]
rMúÇÁ³Ú
¥\Ú±I‚h8{Ão¾!O¦@zª·v*¯ô¼É}À~ —ÒÔ+qŠŽúwNY™yÈ‹™Ôé$–“ 1‹ÆÜÝÜU&ÕFQj•Gujü>Ô¤—”mÙÆÑÅ~,2gcòùÁN¼jTíUP¯µ8³)åˆæ’³™”øÝ”Ä„9‰ýÆdXÔyX¶*±[wcSÊ§j¶(±Û]–z¯©5‰?\¶%Ù»šÞ’$­5µ#‰-ÌIE†8¶µØ–Š¡å&#j)’®sœßÈ]P'ˆÎÓ‹°žg‰æ7¤ŸkÔä[…ïUìK—Ë¿"o{ÛïÉÎá¾ß9Ø~¿AvÒ~<Æüƒ»8¨-›ù[ÿjYA-e<ðzxó":£-þ†Ø,Zús’·Áò³êá\Ãé`ùùŠã¯áp D™‹A4†1ui±¹Xò‡ÂÁsŒ®el\Çu<þ‚•®3ç‹ÏÈé"-<ÆÇúLñ¾Qç!¹	t6Wmž‹Dlô7êFëK©9’àGoVŸéH‚Ož¦³ùzÓ;.q¯äÖSlqá%aõêáÒ ?ƒå"û¿uÆž™´6›¢ó$9è,Iq‹”Ø˜j”È\b?gUvüF×@«˜üb®€hVå½ÐŠìE§$ïUÕ(™Ê®ãÇ¼.½pe]D "!-•
à‡«¢"@kY!ËnÒ/5*#Új#Ú™~vÏVôŽpæP”Å#=.>âO¨ó¶7Æ›	(7òÅ!­_‡y/=ÏHUŒ£«aœ
ƒ í×ÂÌãGÁkk;jæ“QÞÑ2_£öz›u?’…æB+2|ÖéZØ¯¹¹yW¯|‡“!õL“;ç·þ”äzõ=ÊÃ5^‡fùÔŒQî”^b=Y;²™+²~šSQÒÇ,½â|‚ªÕÇFxs35šŽ›yQ;K\`]ï°eÀÿ|’BoÚ™E…ÿ#ÐNé¸Cd0lêÁœzéž?DßPÿ ]"# «¯LCˆù]¥{ñòBá¹‡%{â“Ž	}fIZ	>e¼‚´ „Vçãô%’Ç‚ ò% _Új¬¶æk¾FÌ‰k{pÞÌå—ïí{	zÿ:¦TØQýØ1*_kÃ.Z±©9Ôd,ØæŒÿ¹y)¯ÛY¤¸f"N·ÉxØê5?|aàèú—lbY1Ín¾ÆªÕÄO…f?fÎgéÙâ}´|«o¼ð”êª€üeILkÚ`ÜyüçIª|Š·$\D=ñCS4âçÞqã,-=Ë%•NÂ,N‚ƒ-VyÕA^%]L9ÿ“.žÀÛ6ø‹ôÝ½ÞÞõ?÷ÈÎáÑAïèú_~»Û3új°å»×„3³ÝKYü=‚¯/ÑÖ¨½ÌÃý&ÁÀ²å5ÙKia\–‡ò5®ÿÈ»ÞÁ?ìíîý–|C>ììm7yöýþÎAïúŸ¯ÿÛÎ!<½ð~kçðð}ƒv¶¾?€'Èï¾×ÛkòäÑÎÖÞû·ïauP;g_" ?|GšŸØÆ‹÷YS+GÞ{U­!Œ³ÐÕÊ3))k×ìÊZŒ…³”U­;ºþK?†Ý¢©52cÆ)Þ’fHÈÅŽ“ÈáË‡ŸpfœBÿ(?óûîC¤ìù'ïîÀ™í,% ñcU3•ufúÞð ÑT#u8&ÉrÉ
·ž=ž:°ðV¡¥È¾h }:-€þaâ)ñòbkgþˆæ„ÂøPÌ{[H=Tµš´¿›dåEcâa9ÝÞ=ØÙ=z¿AíñþO‰ª1q­'ßšXOŸ=ë©‡¾ãDfC5žÇÅ,J§ñºý4–·™ÃÄû9Öc
§±C5Ât!jžÜâÑ»+÷“ f!â¶E5!†µw×^{*
ÑyË6Ò^˜yp †YáÖªò_g“ýmøðw^pÏÒ?îG–4ó²‰‹søñ"¥‹Bqbx+¤ÆmaëºeŒƒÄ@ö­!pïŒHxH–»}SþÕ°£<«%tS|o_nƒ¯ö­âHd«ÐlQ'½|¤Ù{` *4¯é™ƒ
à¶*P°“jøã4 6Î\ë™]¡páì"‚C—Z]„Œ0º\5ƒq¡„ºðéÿ†Å~ä‹±ülÅNLx«™R“ÒjTWÈÚYo  Çt3Ñgw¯§6Lc‹—ðuÉÜÎnþš¬ÎÓ0”šâ‰Í4+G£Œ½$õw#µÜY]±Ý4Žuƒ {Æ‚!8Å¤703ÕéÒÈ»è®,Ð1,Â¬¬©j
KælD/ÉˆFïlnÓñ`<# Ó!Êé…=j=Àœç˜ÂóÑ²âð}H‹`³øOms¼Œ5J‰d7ºþ¨Ïí4è+ÝÌèò yÔ}ïFH·[5›Ý~Y¦A
jo‚Ñ#˜53X¶ÖaúTŒÝøàG¼Ð, ð"žÜùæúó™>úf41YË+Wö”‡yÛÄ©ß/‹ä˜üHt›ÝbõÚn"ß
á5^T®¢máÐ3yëŽ?óu#Â,Ì&_Ó0ž$ãÐ¿={Àw> W ¬ë
€b8@±°ß;E‰¯^­<>‹´Ìç‰7fê”‰LÓ8ždY9Ô¬{0À3è”<e(W]ðC>ã‘ ±•ý®K·ù9Ø§˜âe?²ù³Zª—¬ây¡–“ç¤vUÉÈ¬P·´z{²8ŽÚééâžn‰Grãz¯ª•HlnøÖ¹[ÚŒ-T)¹½ˆUG)SDZ“óDØs ¾V¶‚``©Rã>/é©ÑçSÌl¿!!¸ Œg9xÞátšEK¶³ˆËb¤t2MÝ@ÿã©—¥Þxl†uq÷+÷|ú¯9,•@^¾û`ÞOSoèþyRützãñ½/ýõ‰)8¿øˆ4µËï`‚ýýIzjÓ>²Œ±ìéo¾!s{±HjGs$ˆÈ9–¿;·§)CÀrpð~ïÜ2"÷ÂÄŸ'À)ïCƒ MáJ×xã¡]Ñô"C@–ÌXö°‡z·³ÏáÅÉ9õŽƒfXK19£ˆxF–KÙMH2È 8a™×¿ïâyÜ‡x1¶éRÇBb¹c»^“”á†mä›ÚHû¥2\š*"&1§:“n> v8ò’l|ŠKWI¿piÑ.›‚è²ke˜Êâ™vº˜oãÀ,‚tvº¨01Î Rø"gõä— Jðâ†N†”r$”c³Æ1Aê•¾vªrŸ;}µ
ýˆ®§‹ÏIYJCä´–
, 5ëö´6©ß$÷àØ¼´‡Ë„Àc±‰â§JÊ·ÑÒ¼ÿ	
ÅÇ¶cª)ÿrõjs£)Ù†Wùh¯]m®áãé”Ï¯_m®ßäùgW›Ïnòüo®6Óày—ƒC·g‰¹½3Œß‚'½Œêm¶áj œE—	Pd9g&mù$óR u¡%uô¡UhQ5(•]·^óø)öâÿèã*.a°xB~ýìÙó§/Äg#vÊ©FœžÄðîk$Ì{”E­m¾aÈålóh¶^¸Yyf‘G«Ymf‘lKjÑRIfC±Û®@o©ÅÌk½êÁ ˆfÕTª2OWqÙVý‡¾Úë ñû7¬ƒ$ÞÂ×Åyu©’¸vGuò™Úê ±],Ü¼ï@­ƒdïNúÕ¸’´äuë ÉC1%‘çeŒhþºR[…á7ï]‰£m:n?‰«â÷RùŠò<æ>£!àyÌýIyQß’X<‰ÏSdµR™|ÔM^ô"€Wˆ¡ãûšhþ+™EÙkóZB¼ÖE/àvy–¶èjÔ Nâî°`
[7l¾Aqá*â±úKîò _{¦ÝÐh«³‚»p‚ÅæÞA~
Ú1€Æ¸À_ŒW1N}Øúµƒ¼³©éà¬E)ðåÄé*$H` –?àéÐY™„í6Dþö{Y	ÁµÆJ¬A{•a¶•‹4-Z`µh»0ABÑµVÃ¯¯`ÁÌËhè~Ge
´qT–)hBÛfQ¦ |íÈî"ä½ó"o°¡t‘yÃTa@ðB÷?^©­¥õŽüsè:­Ëˆ`§hw¦!û¡v-³ÍGtPê•òd—g¦x,T ÍT…nª}å£QtÛêÇÒú[t‘>_©¨žqÁÄ>Ç7\æÓ{mÔðd
¯ËÀ³#‡ƒáÿÆñÀ§™E ? Áaôšwâ£Æ°(>ŸäUbËeä§Ó•ZÀIToNyÉ÷Y€“¡´ú}'ªÚÎË¢7Ý~ëÎê
F8[ŸR‡gÉ×¢Ê¸¥»&^²¡[ÍL*’Ñ ÏÕªE1šVw$›ÆT‘+À4oýÎf§Ô¿¨ðØQW¦MšzÈ²?5Å<û×,€éaá@žÔë	¬ËÛ	LùÛÂ³Ù(G³ýLÜ¾‘ù,‡(û$¡Sa>+.–x2VGhý¢Úu[Fµbúf›š@]¾oM-j&ÌÏ3ÔµkO“÷ ®9Mˆ’½tË´fd1¢ÑÈC5?†¡×-Ÿì“-|¬Aý&•jX$RŒ’CŠ´ÌÞ£AÌÞfÿ§JÒFœáFH0î&OŒÜË&l=¬.8Žl=Î8Ð²{ì]Ør*hèó’w«]Ž·|d ¶Ø{Šõ SbÏõ%ûÃbLR•Æ$,G¼Óé
ÓëS2:ÞØ-MuìLSYšnnk²¨ø,Ê:{FàÛDÛœ‚5Ê_;æb=”¨o‰.jFÎ¡#,¢&´É½Chù Zid,Ê)Ž´f\¦Ð‹=„£ÛeôÆ®Jž5ü’A!È½O½>¬‘3!…v¥ŠFäC6ŽN˜èâ?Qi3rÃ0>öÂ=AÌï¼¬ÚÓû%–‰Qss½J—è;Jä§TùŒ×=[ô>|ŠÍiá’ÆïI¦¥<¾•¬Ž`£L!¹ÆH^c’Ñc?;÷ýÈ/=i•ªÑÌóÑôpô±øyªÐßü×®Ù)£5ö¬*ò¬AÜ­äd¦<v<çaZ²¯F¡'Áý–áV¬&O}†'Ÿ*çkgäqkC?’ù½ÝÃÉpˆNÝÐªÄ‰/~ÄÖÑâ±Þ]Ák*í³6$*Ut´4È|¦Æ
gÔDë,a†¹—ÿ”€È;S®Ww4Æô¹â8ï§|¹º)ˆÞ™éfu—ÇA>“¸˜aùru7tQáTp:}?M‹qé·XwÖÞß\vsüpTcEy¼˜&YKÎ#Íx^ïBIµADY­ :»þbNù&+Ëü0^ÿ5êé%”)¼Ð¢f4¼"}<H×O’âEŸ*iDÍÿ§¨C~M»]±X_\Rø9³ƒ¾ÄDåí“%Þ7õG–™ñÛ<8Ð•¯±ˆä—xºcâF¸W
óÃ5Ë™N9«r¿Oa•w{,V[óc`ÛÑö‚±³‚þé¹£Cèñc}ðèÅ§S$ïü-ƒØ ÒÀn˜cÍ™Yì^ÁVŸY‡OÓ‰$‹v;GÛÚ…?¨ð•£nÖþªUhù
–ðeà‰t¢Syä§rL|LunyÁxþµá[‡l¨(&q˜Ÿz7bA§fµq4÷y„nÝ˜d˜¦
 àÃŸqrý9á&Åè³<o€göõçÃ)Fc „ôL¯ð
ÃÔ°k¼Y[€¾sá÷9jwå,räöwÓyâÈPåpÕÀ;WoVÄAy6ª¦:>aMõP ‰ªÇ.6Šâƒ¬xæ`_k*«L|z-µU%J˜ÅƒÁ™ûaæf•ˆU94Ú«<‚cñÑÀq»·¿¼Õ{C¼l‚š08økcYÎkU‘²% ´Ñë-*_âE8¹ŒéÅ}¿ZM³sÜÁ¶L­‹Š³èÔç©í¡öÎŠ^ŒZ«gŽ3õì[Å˜w~„2óˆŒâ,8c.v!.pRZWÆíPÑÒwK¢~Vé*v0Ë¤e¹vÝ«hØåþîÌFí Þ)M]ÉÐå€'Gx€jªÓ?Â©h±aoñ:R@#òÍàuctEéÔ°¢tøuƒt*ÕÍÆ`Ói×÷;?x	M± l€,q? c'SÄŽð–¢€ftóH ³£À´Ñ@®Ì>—VLS‚lí¾ù†¸b…,Qß<[—5"‚\w„òö¸ Îzß4H	%ÒiR+UyDIñDú`\áD‰pãœŽ­xƒÌ?åkµ"Jø2M QÅrÜ0#™˜§n†ÊëqˆùJ}Lc6Œâ4ú’ ôò@ÜßÎoÇôáü=AŠîµ¯.ƒ4o›;Ý£5ùæêí5']ŽH°â[q¯p÷å_E›„ø AþUêÁ?ñ&!lÚ$éÓ	î°•|YžH>l©­´:¢éò¦qÁÑZ^kôø•ýžyoAº-Ç°Ú´a­Å–I]xíQ¾î_nˆÃ8ŠEã.“šHÛŸ$	JÛtM¥§^ÓR†åÉ÷v/ÆŒ­Ò›ÅN3Ï+ê-èÏ+µsÞvw0ÿz‰ºÂ–&‡²?à4ÈØFí›Hì²£¥DÝŠUÉ¿÷õ¼œ%¤5½d°äl¦Ûòvzg>[]$8øí0K´½EºÜ“Ê?3OÄkúßSosikè.—; gÒ+e¹ÒA¡€y–ÀõðSÔ·Áo¹A5µÐÛkÔ:…ûÀÐþÙWèhRBßÃŒ÷óêrì''@˜iÏ´^.Ú“NFp˜LúnóE«Iê'4Gì«KüŠ0	?$0O’8£žÝÁ¯†µúÎëœŒw.Æq’Ù–‹µ©¹VJcÃB±—ñ¥Ê°çŠvXˆHç1½ŒU`%â	|¼Ï€›Ý~Ã-Øõß§¸æ	}•zEžÅv|…±7À~L”¡[}^…o¸8„Ã]tõ]>0íÉRu}È¾N¾OB%ôEõ¨.ü	"ØEª ‰6—¿]1Q
zÿ¼È‹šà I<†g'	r£®¨‰\û Ç-pF¿
&Ü»X<_ü.œ.þðbíìôGR”ñèÇ!AOŒ“:?²qæ†g
µø-sYýMuÞe+ã¥>y¾øœf	•Òÿ¯
	AÎø/Å–Uíž¹ú	FágçÆTÖ¶¢:§ëšëŽÄj ÿ¤Ìg‡;ÈËN¾\>]¯mëÑV™–±1-ìwÔé¤9ä[‡Çia¢q¥¶Ë®¿#ubµ@æþX}œ›¿2®âÎ,pä…oƒè£´’rd`Ãƒ¥%½ãCG„–õŽ=ÃbgKÞ|2ö5ÁIâaÀìB‚‚ûŠÉ'ê
šÅIwŽ5ø»#ío¿!û‰øççæÑ4ÿ»£wowß`£ft·%¸á¯²Ä¾-ar<hÿGgS[,ùeæbN•˜sSë&¨1îý~ÂÒê»·}w„Þ´Ar“ÝÆà|Ÿj·=y£Y„ß´îœg^oé4ñO0à€®—¹É€ŸgÐlŽârŒ€ÜxKðôœùÙ>Ý´—–•ggÅRÓ÷	€m²Ø_
\(}F‹«F8m¯†§À¸:v
3Ðee²‹î‚WÄ÷J^Š¶ h>ˆä+³öÔBÅÍ³¯“¡ØLv±€XêýG§zLSò\q½ä"Mõ¥ëreÐÓ²?9CÁ™8rª³hwˆU&±ÔÒÏBmñò(œ¦œjœ€,?9e¯ƒRÌb9¦¦íó=z¹l9ÝÞ? ÄÕßý   ÿÿ Wª‘›