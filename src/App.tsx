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
import { getCachedAI, setCachedAI, generateAICacheKey } from './lib/aiCache';
import { 
  Plus, 
  Award,
  FileText, 
  Building2, 
  ChevronRight, 
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
                    ? "bg-gradient-to-r from-blue-500 to-sky-600" 
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
                  <CheckSquare size={14} className="text-sky-600" />
                  DiagnÃ³stico Total
                </span>
                <span className={cn(
                  "text-xs font-extrabold px-2 py-0.5 rounded-md",
                  diagnosticoPercent === 100 ? "bg-emerald-100 text-emerald-800" : "bg-sky-100 text-sky-800"
                )}>
                  {diagnosticoPercent}%
                </span>
              </div>
              <div className="text-2xl font-black text-slate-800 mb-1">
                {answeredDiagnosticoPerguntas} <span className="text-sm font-semibold text-slate-400">/ {totalDiagnosticoPerguntas} perguntas</span>
              </div>
              <div className="w-full bg-slate-200/70 rounded-full h-1.5 mb-2 overflow-hidden">
                <div 
                  className={cn("h-full rounded-full transition-all duration-500", diagnosticoPercent === 100 ? "bg-emerald-500" : "bg-sky-600")}
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
          <div className="w-10 h-10 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center">
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
  logoChoice: 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'sebrae' | 'consultora' | 'none') => void,
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
            className="px-6 text-sky-600 border-sky-100 hover:bg-sky-50"
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
              <div className="w-10 h-10 bg-sky-50 text-sky-600 rounded-xl flex items-center justify-center">
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
            <label className="flex flex-col items-center justify-center p-8 bg-slate-50 rounded-xl border-2 border-dashed border-slate-200 hover:border-sky-400 hover:bg-sky-50 transition-all cursor-pointer group">
              <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center shadow-sm mb-4 group-hover:scale-110 transition-transform">
                <Upload className="text-slate-400 group-hover:text-sky-600" size={24} />
              </div>
              <div className="flex flex-col items-center">
                <span className="text-sm font-bold text-slate-500 group-hover:text-sky-700">Adicionar fotos como evidÃªncia geral</span>
                <span className="text-xs text-slate-400 mt-1">Imagens comprimidas automaticamente (reduÃ§Ã£o de 50% no tamanho)</span>
                {dadosConsultoria.evidencias && dadosConsultoria.evidencias.length > 0 && (
                  <div className="mt-2 text-[10px] font-bold text-sky-600 uppercase tracking-tighter bg-sky-50 px-2 py-0.5 rounded-full border border-sky-100">
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
                          className="w-full text-[10px] font-bold bg-white border border-slate-200 rounded px-1 py-0.5 text-sky-700 cursor-pointer focus:outline-none focus:ring-1 focus:ring-sky-500"
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
  logoChoice: 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'sebrae' | 'consultora' | 'none') => void,
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

      // Add logo if available
      const activeLogoForPdf = logoChoice === 'sebrae' ? customLogo : logoChoice === 'consultora' ? customConsultoraLogo : null;
      const logoValid = isValidLogoSource(activeLogoForPdf);
      if (logoValid && activeLogoForPdf) {
        try {
          const imgProps = doc.getImageProperties(activeLogoForPdf);
          const logoWidth = 25;
          const logoHeight = (imgProps.height * logoWidth) / imgProps.width;
          doc.addImage(activeLogoForPdf, 'PNG', 190 - logoWidth, 10, logoWidth, logoHeight, undefined, 'FAST');
        } catch (e) {
          console.warn("Could not add custom logo to Relatorio PDF", e);
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
  logoChoice: 'sebrae' | 'consultora' | 'none',
  setLogoChoice: (choice: 'sebrae' | 'consultora' | 'none') => void,
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
      { color: 'bg-sky-50/50', text: 'text-sky-600', border: 'border-sky-100' },
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
             className="bg-sky-600 text-white px-4 py-2 rounded-xl text-xs font-bold shadow-lg shadow-sky-200 hover:bg-sky-700 transition-colors"
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
                        className="p-1.5 bg-sky-50 hover:bg-sky-100 text-sky-600 rounded-lg transition-colors cursor-pointer"
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
                  ? "bg-white text-sky-600 shadow-sm font-black" 
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
                  ? "bg-white text-sky-600 shadow-sm font-black" 
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
              className="px-3 py-1.5 text-xs font-bold rounded-xl border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20 cursor-pointer"
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
              className="px-3 py-1.5 text-xs font-bold rounded-xl border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20 cursor-pointer"
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
                      className="bg-white p-5 rounded-2xl shadow-sm border border-slate-100 hover:shadow-xl hover:border-sky-100 transition-all cursor-pointer group relative overflow-hidden"
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
                          <div className="flex items-center gap-1 bg-sky-50/80 px-2 py-0.5 rounded-md border border-sky-100 text-[10px] font-black text-sky-700" onClick={(e) => e.stopPropagation()}>
                            <ListOrdered size={12} />
                            <span className="text-[9px] text-sky-500 font-bold mr-0.5">PosiÃ§Ã£o:</span>
                            <select
                              value={tarefa.ordem !== undefined && tarefa.ordem !== null ? tarefa.ordem : sortedTarefas.findIndex(t => t.id === tarefa.id)}
                              onChange={(e) => {
                                e.stopPropagation();
                                handleSetTaskPosition(tarefa, parseInt(e.target.value, 10));
                              }}
                              className="bg-white font-black text-sky-700 border border-sky-200 rounded px-1 py-0.5 cursor-pointer outline-none shadow-2xs hover:bg-sky-50"
                              title="Clique para alterar a posiÃ§Ã£o da aÃ§Ã£o"
                            >
                              {sortedTarefas.map((_, idx) => (
                                <option key={idx} value={idx}>#{idx + 1}</option>
                              ))}
                            </select>
                            <div className="flex items-center ml-1 border-l border-sky-200 pl-1 gap-0.5">
                              <button
                                type="button"
                                title="Mover para cima (#1)"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'up'); }}
                                className="p-0.5 hover:bg-sky-200/60 rounded text-sky-700 transition-colors"
                              >
                                <ArrowUp size={11} />
                              </button>
                              <button
                                type="button"
                                title="Mover para baixo"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'down'); }}
                                className="p-0.5 hover:bg-sky-200/60 rounded text-sky-700 transition-colors"
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
                              "text-[9px] font-black px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20",
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
                              "text-[10px] font-bold px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20",
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
                      
                      <h4 className="font-bold text-slate-800 text-sm mb-1 leading-tight group-hover:text-sky-600 transition-colors uppercase tracking-tight">{tarefa.problema}</h4>
                      <p className="text-[11px] text-slate-500 font-medium mb-3 line-clamp-2">{tarefa.solucaoSugerida}</p>
                      
                      {tarefa.area && (
                        <div className="mb-2 flex items-center gap-1.5">
                          <span className="text-[9px] font-black text-sky-500 bg-sky-50 border border-sky-100 px-2 py-0.5 rounded-md uppercase tracking-tight">
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
                          <div className="flex justify-between items-center bg-sky-50/50 px-2.5 py-1.5 rounded-lg border border-sky-100/50">
                             <span className="text-[10px] font-black text-sky-600 uppercase tracking-tighter">DuraÃ§Ã£o</span>
                             <span className="text-[10px] font-bold text-sky-700 bg-white px-1.5 py-0.5 rounded shadow-sm border border-sky-100">
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
                            className="relative cursor-pointer p-1.5 hover:bg-slate-50 rounded-lg text-slate-400 hover:text-sky-600 transition-all"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <Paperclip size={14} className={tarefa.anexoUrl ? "text-sky-600" : ""} />
                            {tarefa.anexoUrl && <span className="absolute -top-1 -right-1 w-2 h-2 bg-sky-600 rounded-full"></span>}
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
                              {tarefa.lembreteEmail && <Mail size={10} className="text-sky-500" title="E-mail configurado" />}
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
          <div className="absolute left-[37px] md:left-[45px] top-24 bottom-12 w-0.5 bg-sky-100"></div>

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

                  <div className="space-y-6 pl-4 md:pl-6 border-l border-sky-50/50 ml-5 md:ml-6">
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
                            <div className="flex items-center justify-center w-9 h-9 rounded-full border-2 border-sky-200 bg-white shadow-sm group-hover:border-sky-600 transition-colors relative">
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
                              <span className="text-[10px] font-black text-sky-600 uppercase pointer-events-none">
                                {String(tarefa.ordem !== undefined && tarefa.ordem !== null ? tarefa.ordem + 1 : sortedTarefas.findIndex(t => t.id === tarefa.id) + 1).padStart(2, '0')}
                              </span>
                            </div>
                            <div className="flex flex-col gap-0.5 items-center">
                              <button
                                type="button"
                                title="Mover para cima"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'up'); }}
                                className="p-1 hover:bg-sky-100/80 rounded-md text-sky-600 transition-colors bg-white border border-slate-200/80 shadow-xs cursor-pointer"
                              >
                                <ArrowUp size={11} />
                              </button>
                              <button
                                type="button"
                                title="Mover para baixo"
                                onClick={(e) => { e.stopPropagation(); handleMoveTask(tarefa, 'down'); }}
                                className="p-1 hover:bg-sky-100/80 rounded-md text-sky-600 transition-colors bg-white border border-slate-200/80 shadow-xs cursor-pointer"
                              >
                                <ArrowDown size={11} />
                              </button>
                            </div>
                          </div>

                          <div className="flex-1 bg-slate-50/50 hover:bg-white rounded-2xl border border-slate-100 hover:border-sky-100 p-5 shadow-sm hover:shadow-md transition-all flex flex-col md:flex-row gap-4 justify-between items-start md:items-center">
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
                                  <span className="text-[10px] font-black uppercase tracking-tight px-2 py-0.5 rounded-md border bg-sky-50/50 border-sky-100 text-sky-600">
                                    ðŸ—‚ï¸ {tarefa.area}
                                  </span>
                                )}
                              </div>

                              <h4 className="font-bold text-slate-800 text-base leading-snug group-hover:text-sky-600 transition-colors uppercase tracking-tight">
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
                                  <span className="bg-sky-50 text-sky-700 px-2 py-0.5 rounded-md border border-sky-100">
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
                                  className="text-[10px] font-bold px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20"
                                >
                                  <option value="Pendente">Pendente</option>
                                  <option value="Em Andamento">Em Andamento</option>
                                  <option value="ConcluÃ­do">ConcluÃ­do</option>
                                </select>

                                <select
                                  value={tarefa.prioridade}
                                  onChange={(e) => onUpdatePrioridade(tarefa.id, e.target.value as any)}
                                  className={cn(
                                    "text-[9px] font-black px-2 py-1 rounded-lg border border-slate-200 bg-white focus:outline-none focus:ring-2 focus:ring-sky-500/20",
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
                                  className="relative cursor-pointer p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-sky-600 transition-all"
                                  onClick={(e) => e.stopPropagation()}
                                >
                                  <Paperclip size={14} className={tarefa.anexoUrl ? "text-sky-600" : ""} />
                                  {tarefa.anexoUrl && <span className="absolute -top-1 -right-1 w-2 h-2 bg-sky-600 rounded-full"></span>}
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
                                  className="p-1.5 hover:bg-sky-50 text-sky-400 rounded-lg transition-all"
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
  customConsultoraLogo 
}: { 
  logoChoice: 'sebrae' | 'consultora' | 'none', 
  setLogoChoice: (choice: 'sebrae' | 'consultora' | 'none') => void,
  customLogo: string | null,
  customConsultoraLogo: string | null
}) => {
  return (
    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 w-full">
      <div>
        <h4 className="text-xs font-black text-slate-700 uppercase tracking-widest flex items-center gap-2">
          <span>ðŸŽ¯</span> Logotipo no RelatÃ³rio
        </h4>
        <p className="text-[10px] text-slate-400 font-bold leading-relaxed uppercase tracking-wider mt-0.5">
          Escolha qual marca exibir nos cabeÃ§alhos dos PDFs gerados neste painel
        </p>
      </div>
      <div className="flex bg-slate-100 p-1 rounded-xl self-start sm:self-auto border border-slate-200/50">
        <button
          onClick={() => {
            setLogoChoice('sebrae');
            localStorage.setItem('preferred_logo', 'sebrae');
          }}
          className={cn(
            "px-4 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
            logoChoice === 'sebrae'
              ? "bg-white text-emerald-600 shadow-sm"
              : "text-slate-500 hover:text-slate-800"
          )}
        >
          SEBRAE {customLogo ? 'âœ…' : 'âš ï¸'}
        </button>
        <button
          onClick={() => {
            setLogoChoice('consultora');
            localStorage.setItem('preferred_logo', 'consultora');
          }}
          className={cn(
            "px-4 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
            logoChoice === 'consultora'
              ? "bg-white text-sky-600 shadow-sm"
              : "text-slate-500 hover:text-slate-800"
          )}
        >
          Consultora {customConsultoraLogo ? 'âœ…' : 'âš ï¸'}
        </button>
        <button
          onClick={() => {
            setLogoChoice('none');
            localStorage.setItem('preferred_logo', 'none');
          }}
          className={cn(
            "px-4 py-1.5 rounded-lg text-xs font-black transition-all cursor-pointer border-0 select-none uppercase tracking-wider",
            logoChoice === 'none'
              ? "bg-white text-slate-800 shadow-sm"
              : "text-slate-500 hover:text-slate-800"
          )}
        >
          Nenhum
        </button>
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
            <ImageIcon className="text-sky-600" size={18} />
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
                <label className="flex flex-col items-center justify-center p-6 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-200 hover:border-sky-400 hover:bg-sky-50/50 transition-all cursor-pointer group">
                  <div className="w-10 h-10 bg-white rounded-full flex items-center justify-center shadow-sm mb-3 group-hover:scale-110 transition-transform">
                    <Upload className="text-slate-400 group-hover:text-sky-600" size={20} />
                  </div>
                  <span className="text-xs font-bold text-slate-500 group-hover:text-sky-700">Clique para selecionar o logo</span>
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
          <Database className="text-sky-600" size={20} />
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
                ? "border-sky-500 bg-sky-50/40 shadow-sm"
                : "border-slate-200 bg-slate-50/50 hover:border-slate-300"
            )}
          >
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className={cn("p-2.5 rounded-xl", storageMode === 'local' ? "bg-sky-600 text-white" : "bg-slate-200 text-slate-600")}>
                  <HardDrive size={20} />
                </div>
                <div>
                  <h4 className="font-bold text-sm text-slate-800">Salvar Localmente (Dispositivo)</h4>
                  <p className="text-xs text-slate-500">Navegador & Sem Cotas</p>
                </div>
              </div>
              {storageMode === 'local' && (
                <span className="bg-sky-600 text-white text-[10px] font-black uppercase px-2.5 py-1 rounded-full">Ativo</span>
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
              <div className="bg-sky-600 hover:bg-sky-700 text-white font-bold text-xs px-4 py-2 rounded-xl flex items-center gap-2 transition-colors">
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
            <p className="text-xs font-bold text-sky-700 mt-1 font-mono truncate" title="ai-studio-consultoriaempre-e5930715-81f9-43df-8493-c8c9c3ecac45">
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
                  <RefreshCw size={14} className="animate-spin text-sky-500" />
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
              Todos os registros armazenados no <code className="bg-slate-200 px-1 py-0.5 rounded text-[11px] font-mono">localStorage</code> sÃ£o gravados sob o formato cifrado <code className="text-sky-600 font-mono text-[11px]">enc:v1:...</code>.
            </p>
          </div>
        </div>

        <div className="p-4 bg-sky-50/60 rounded-xl border border-sky-100 flex flex-wrap items-center justify-between gap-3">
          <div className="space-y-0.5">
            <p className="text-xs font-bold text-sky-950">Verificar e Re-criptografar Dados Locais</p>
            <p className="text-[11px] text-sky-700">Converte dados legados nÃ£o criptografados no navegador para a nova estrutura AES-256.</p>
          </div>
          <Button
            onClick={() => {
              const count = encryptedLocalStorage.migrateAllToEncrypted();
              alert(`Varredura concluÃ­da! ${count} item(ns) legados foram criptografados e salvos com AES-256 no LocalStorage.`);
            }}
            variant="outline"
            size="sm"
            className="text-xs border-sky-200 hover:bg-sky-100 text-sky-700 font-bold py-1.5 px-3"
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
            <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-sky-500 inline-block"></span> Definitiva (VitalÃ­cia)</span>
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
                            ? 'bg-sky-100 text-sky-700 border border-sky-200' 
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
                              <span className="px-2 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-sky-100 text-sky-800 border border-sky-200">
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
                          <span className="px-2.5 py-1 rounded-full text-xs font-bold flex items-center gap-1.5 bg-sky-100 text-sky-700 border border-sky-200 w-fit">
                            <InfinityIcon size={14} className="shrink-0" />
                            LicenÃ§a Definitiva
                          </span>
                          <p className="text-[10px] text-sky-600 font-medium">
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
                          <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-semibold bg-sky-50 text-sky-700 border border-sky-100">
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

      <div className="bg-gradient-to-br from-sky-700 via-sky-600 to-blue-500 rounded-[2.5rem] p-12 text-white shadow-2xl relative overflow-hidden">
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
          <div className="inline-flex items-center gap-2 px-3.5 py-1 bg-white/15 backdrop-blur-md rounded-full text-xs font-bold text-sky-100 mb-4 tracking-widest uppercase border border-white/20 ml-2">
            by ItÃ mar Gomes
          </div>
          <h2 className="text-5xl font-black tracking-tighter mb-6 leading-none">Melhore os resultados dos seus clientes.</h2>
          <p className="text-sky-100 text-lg max-w-2xl font-medium leading-relaxed opacity-90">
            Gerencie clientes, realize diagnÃ³sticos precisos e acompanhe a execuÃ§Ã£o dos planos de aÃ§Ã£o em tempo real.
          </p>
          <div className="mt-8 flex gap-4">
            <Button onClick={() => setView('companies')} variant="secondary" className="bg-white text-sky-600 hover:bg-sky-50 border-none px-8 py-6 text-sm font-bold uppercase tracking-widest rounded-3xl shadow-lg">
              ComeÃ§ar Agora
            </Button>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
        <Card className="p-8 hover:shadow-2xl transition-all cursor-pointer border-slate-100 group rounded-[2rem] bg-white relative overflow-hidden" onClick={() => setView('companies')}>
          <div className="w-16 h-16 bg-sky-50 text-sky-600 rounded-3xl flex items-center justify-center mb-8 group-hover:scale-110 transition-transform duration-500 shadow-sm border border-sky-100/50">
            <Building2 size={32} />
          </div>
          <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight uppercase">Clientes</h3>
          <p className="text-slate-500 text-sm mb-8 font-medium leading-relaxed">GestÃ£o completa da base de clientes e histÃ³ricos de atendimento.</p>
          <div className="pt-6 border-t border-slate-50 flex justify-between items-center">
            <span className="text-[10px] font-black text-sky-600 bg-sky-50 px-3 py-1.5 rounded-full uppercase tracking-widest">{empresas.length} cadastrados</span>
            <ChevronRight size={20} className="text-slate-300 group-hover:text-sky-600 transform group-hover:translate-x-1 transition-all" />
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
  const cacheKey = generateAICacheKey('feedback_v2', { resposta, pergunta, problema });
  const cached = await getCachedAI<string>(cacheKey);
  if (cached) return cached;

  const prompt = `Como um consultor empresarial especializado do SEBRAE, analise o significado da pergunta e da resposta fornecida no contexto do diagnÃ³stico empresarial e gere uma observaÃ§Ã£o prÃ¡tica, ultra-resumida e objetiva (com exatamente 50% da extensÃ£o habitual).

CONTEXTO:
Problema/Tema: ${problema}
Premissa/Pergunta: ${pergunta}
Resposta do Cliente: ${resposta}

DIRETRIZES DE ANÃLISE SEMÃ‚NTICA DA PERGUNTA E RESPOSTA:
1. Avalie o sentido da pergunta/premissa:
   - Se a pergunta investiga a ocorrÃªncia de um problema, falha, prejuÃ­zo, inadimplÃªncia ou rejeiÃ§Ã£o (ex: "O produtor jÃ¡ tentou acessar crÃ©dito rural e foi rejeitado por falta de documentaÃ§Ã£o?", "Possui pendÃªncias?"):
     * Resposta "Sim": Representa um PROBLEMA / RISCO / GARGALO. Sugira em 1 frase curta a recomendaÃ§Ã£o corretiva direta.
     * Resposta "NÃ£o": Representa uma SITUAÃ‡ÃƒO POSITIVA / CONFORMIDADE. Gere uma breve frase de reconhecimento do sucesso ou conformidade.
     * Resposta "Parcial": Representa risco moderado. Sugira em 1 frase curta o ajuste prioritÃ¡rio.
   - Se a pergunta investiga a existÃªncia de um controle, planejamento ou boa prÃ¡tica (ex: "Possui controle financeiro?"):
     * Resposta "Sim": Representa uma SITUAÃ‡ÃƒO POSITIVA / MATURIDADE. Elogie e reconheÃ§a de forma sucinta.
     * Resposta "NÃ£o" ou "Parcial": Representa um PONTO DE ATENÃ‡ÃƒO / GARGALO. Sugira diretamente o que deve ser implantado.

REQUISITOS DE SAÃDA OBRIGATÃ“RIOS (RESUMO DE 50%):
- Formato ultra-conciso e direto (reduzido a 50% do tamanho habitual).
- No mÃ¡ximo 1 Ãºnica frase curta e objetiva (mÃ¡ximo de 12 a 18 palavras).
- Sem introduÃ§Ãµes ou saudaÃ§Ãµes, vÃ¡ direto ao ponto prÃ¡tico.
- Responda em PortuguÃªs do Brasil de forma profissional e encorajadora.`;

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
  const [view, setView] = useState<'home' | 'companies' | 'credenciadas' | 'licenses' | 'diagnosis' | 'dashboard' | 'premises' | 'cronograma' | 'relatorio' | 'kanban' | 'settings' | 'landing' | 'checkout' | 'licensing' | 'dados-consultoria' | 'agenda' | 'macro-dashboard' | 'disc-assessment' | 'maturity-assessment' | 'projects' | 'plan' | 'resultado-consultoria'>('landing');
  const [customLogo, setCustomLogo] = useState<string | null>(localStorage.getItem('sebrae_custom_logo'));
  const [customConsultoraLogo, setCustomConsultoraLogo] = useState<string | null>(localStorage.getItem('consultora_custom_logo'));
  const [logoChoice, setLogoChoice] = useState<'sebrae' | 'consultora' | 'none'>(() => {
    return (localStorage.getItem('preferred_logo') as 'sebrae' | 'consultora' | 'none') || 'sebrae';
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
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'tarefas_plano', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.tarefas.updated++;
            finalTasks.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.tarefas.downloaded++;
            finalTasks.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalTasks.push({ ...cloud, ...local });
          }
        }
      }
      setTarefasPlano(finalTasks);
      localStorage.setItem('local_tarefas_plano', JSON.stringify(finalTasks));

      // --- E. Sync Metodologia/Biblioteca (Premissas, Problemas, SoluÃ§Ãµes, Ãreas, Segmentos) ---
      const syncBiblioItem = async (colName: string, localList: any[], setter: (val: any) => void, storageKey: string) => {
        try {
          const cloudSnap = await getDocs(isAdmin ? query(collection(db, colName)) : query(collection(db, colName), where('ownerId', '==', user.uid)));
          const cMap = new Map<string, any>();
          cloudSnap.forEach(d => cMap.set(d.id, { id: d.id, ...d.data() }));

          const lMap = new Map<string, any>();
          localList.forEach(item => lMap.set(item.id || item.nome || item, item));

          const allKeys = Array.from(new Set([...Array.from(lMap.keys()), ...Array.from(cMap.keys())]));
          const merged: any[] = [];

          for (const k of allKeys) {
            summary.totalAnalyzed++;
            const l = lMap.get(k);
            const c = cMap.get(k);
            const docId = l?.id || c?.id || String(k);

            if (l && !c) {
              batchQueue.pushItem(doc(db, colName, docId), { ...(typeof l === 'object' ? l : { nome: l }), ownerId: user.uid, updatedAt: new Date().toISOString() });
              summary.uploadedToCloud++;
              summary.details.biblioteca.uploaded++;
              merged.push(l);
            } else if (!l && c) {
              summary.downloadedFromCloud++;
              summary.details.biblioteca.downloaded++;
              merged.push(c);
            } else if (l && c) {
              summary.identicalOrMerged++;
              merged.push({ ...c, ...l });
            }
          }
          setter(merged);
          localStorage.setItem(storageKey, JSON.stringify(merged));
        } catch (e) {
          console.warn(`Aviso ao sincronizar ${colName}:`, e);
        }
      };

      await syncBiblioItem('premissas', premissas, setPremissas, 'local_premissas');
      await syncBiblioItem('problemas', problemas, setProblemas, 'local_problemas');
      await syncBiblioItem('solucoes', solucoes, setSolucoes, 'local_solucoes');
      await syncBiblioItem('areas', dbAreas, setDbAreas, 'local_db_areas');
      await syncBiblioItem('segmentos', dbSegmentos, setDbSegmentos, 'local_db_segmentos');

      // Finalize and commit all queued Firestore batches
      batchQueue.finalize();
      for (const commitTask of cloudBatches) {
        await commitTask();
      }

      setLastSyncSummary(summary);
      localStorage.setItem('last_sync_summary', JSON.stringify(summary));
      playSuccessSound();
    } catch (err: any) {
      console.error("Erro durante a sincronizaÃ§Ã£o inteligente com a Nuvem:", err);
      setSyncError(err?.message || "Ocorreu um erro ao sincronizar com a Nuvem. Tente novamente.");
    } finally {
      setIsSyncingCloud(false);
    }
  };

  const triggerDownloadBackupFile = (jsonString: string, fileName: string) => {
    try {
      const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.setAttribute('download', fileName);
      anchor.download = fileName;
      anchor.style.display = 'none';
      document.body.appendChild(anchor);
      anchor.click();

      setTimeout(() => {
        try {
          if (anchor.parentNode) {
            anchor.parentNode.removeChild(anchor);
          }
          URL.revokeObjectURL(url);
        } catch (e) {}
      }, 2000);
      return true;
    } catch (err) {
      console.warn("Blob download fallback:", err);
      try {
        const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(jsonString);
        const anchor = document.createElement('a');
        anchor.href = dataUri;
        anchor.setAttribute('download', fileName);
        anchor.download = fileName;
        anchor.style.display = 'none';
        document.body.appendChild(anchor);
        anchor.click();
        setTimeout(() => {
          if (anchor.parentNode) anchor.parentNode.removeChild(anchor);
        }, 2000);
        return true;
      } catch (e2) {
        console.error("Erro no download direto:", e2);
        return false;
      }
    }
  };

  const uploadBackupToCloud = async (data: any, uid: string) => {
    const syncCollection = async (items: any[], colName: string) => {
      if (!items || !Array.isArray(items) || items.length === 0) return;
      for (let i = 0; i < items.length; i += 300) {
        const batch = writeBatch(db);
        const chunk = items.slice(i, i + 300);
        for (const item of chunk) {
          const itemId = item.id || doc(collection(db, colName)).id;
          const ref = doc(db, colName, itemId);
          batch.set(ref, sanitizeForFirestore({ ...item, ownerId: uid }), { merge: true });
        }
        await batch.commit();
      }
    };

    if (data.empresas) await syncCollection(data.empresas, 'empresas');
    if (data.diagnosticos) await syncCollection(data.diagnosticos, 'diagnosticos');
    if (data.respostas) await syncCollection(data.respostas, 'respostas');
    if (data.tarefasPlano) await syncCollection(data.tarefasPlano, 'tarefas_plano');
    if (data.empresasCredenciadas) await syncCollection(data.empresasCredenciadas, 'empresas_credenciadas');
    if (data.premissas) await syncCollection(data.premissas, 'premissas');
    if (data.problemas) await syncCollection(data.problemas, 'problemas');
    if (data.solucoes) await syncCollection(data.solucoes, 'solucoes');
    if (data.dbAreas) await syncCollection(data.dbAreas, 'areas');
    if (data.dbSegmentos) await syncCollection(data.dbSegmentos, 'segmentos');
    if (data.agendaEventos || data.agenda_eventos) await syncCollection(data.agendaEventos || data.agenda_eventos, 'agenda_eventos');
    if (data.discAvaliacoes || data.disc_avaliacoes) await syncCollection(data.discAvaliacoes || data.disc_avaliacoes, 'disc_avaliacoes');
    if (data.maturidadeAvaliacoes || data.maturidade_avaliacoes) await syncCollection(data.maturidadeAvaliacoes || data.maturidade_avaliacoes, 'maturidade_avaliacoes');
  };

  const handleExportLocalBackup = async () => {
    setIsBackupModalOpen(true);
    setIsExportingBackup(true);

    try {
      // 1. Gather all collections from in-memory state and local cache first (instant)
      let exportEmpresas: Empresa[] = [...empresas];
      let exportDiagnosticos: Diagnosticoxœì½MwäF’ x¯_áÉÑ*"ªÈH’RªT‘¢8T’©bw~°Hªúíäæf‚„2R‹‡~s˜7‡9ÍÛûÔÎ¡ßêÔ¯Aþ±53ÿ2w8’)eu‹¯J ÜÍÝÍÍÍíËÍ_½[âÕp8ŒÓè,ËË*çåëÇ¿ô7I*‘ü8Ë‹ê8*’Ó¨<˜DY>ò‰vÆQþJ©X© ³$‹£½‹$«òr$¢ìJÖÝMËñÎE4I£qž,(û<ªæEGqÒµÆ“"‰“lœFq%÷¦³")#öÒ(‘ŸJ^! ï H¦iY"0ýÓ€˜éoÁzùÉ$™Êzò'«§¾êå“¹%ýb3Pª/!”žìITêù–OÁrGÉÙ”æÈ”5o ¼ªðð¡ØŠo£ê<) ÙåeR”â´È§b’£‰˜&Ó¼¸ãh|žˆÓ´(+Uqœge%í¨j[ðÅ;“É3|{˜”3 Ä¨ì;u¦ÑÌ|ƒ:Yr)žG³¯ÊªH³³U¡¿}mëñV†§y±½ébëkq-ÒSÑ/¶‡i<p Ë¤êðzUƒÇâÆÀ*L‰[²¨£Š²Ê‹è,-qM&'Ñø*Vú®Õom²ø–|:’@†gIµ_%Ó~Þ¾‘0ß$rµõÌ0õ—Á°Ft3³¨(“Zø‡£—/†ôäTyÌj ¸¢ˆ®†iIÿöeíÁ ´ì¦ülAÜ¨_7@,Õø\ôèÒCÖæP<MðË)ÌÁ¹OòyÓ1Î‹¸i&vwž=Û{&.Óê\D	b\AµGk%”ÊbQ¥Ó$ŸW¢ŒN“³yTÄ¿±}Ÿ—I1¨!A«Z["*¯²±øêxõëþŒÖ*,éä«ã¯W5äç8°ÍGëë«fMxY"†·"¶•éÃ"'ýWì»³UçéÝ‚ë÷ùä€ÖToû}z¡>öuo¬«ƒƒûšMæ!Q`€š'@ˆ[f7ž¶YÝmñ~ìóÉ$Wižõã“UÑÓl³4ÀŠ_—ÀN’~/¿Ì’b?îÁ×­-ø/NÒpKlPë"î\^;t‘ïw=!M½tkÜ¦§ÈA–ï©á;Øè$¦Us}}½SŸºËwXíçË÷YU|3CI z½*·é®dyË÷Ög•]ºë×¹å–ãÛPn9~I§+ñz•nÓßiT-ßÝ©‘Î–îtCÕÛt}Üò}×ŒéÍ˜ÉÝúÞPu¹¾»€‹»I”Y4	vµörøUý-Š'õ·J¼¯{wx9®¿I¼þeg¾ßàvz¥•Ùù`‡ò6>¶ùöA¬ÙÍÇe_ïDƒ!I	j›Ëæ“	 âªzXÉlwe8úÝAq^{whŒÞÂŸº;(ÍBîÉ¬èE <YÇ“\ÕòŸ~*¨ß¨õUW®XÌä!P9Ê‡R"™î!ÿ¤¬¾ZÕ!Á¾J`¤'$¤'$¯¶îQ#6uc¯nLuA‰GBþFmGU8¹áU÷ÎïvF%åxÔâúªà»ó41§æM‹Èoñ¡¸ÃˆÜµbV2Y	)
*ÃJÂ(u®ŽÒzíVœ²n„ñÊûéâV·Ô¹È¶rñ!Œ\óµ>:· ¨«ª†¬+xÃ½YÜµÁ˜!¨ç6~½“ˆg~j n²2¨¨
2º¢‘V!2Ñ}¬ã’è€6¯¿ajáÝu©E·ÖZ”@#[>†qVÑY-=‹y¿q 8D9f¢·uÉµ ª>Q¥{ì,¦Cð»­?!+£êÖ¼ÇN¡Àdz…m«·ÁÖ‰™.
7n:ã0p¢ó1eZ(ÔÇú*áµ»ïQ¬7á…Â»ë.Ý`ÃB1¿´‰I{EQ·€å“dxYåÅü"™Š4h–gïÿv‘L,ÐiNöPç­+ó &ŸÎRX½ð
~
 XJh4KKQ&gó"/G+«Â40€ijP²—[<ÙXö]°å—OóD®ŠlyH »Q•ôÃ*ß?zyD„OEJ÷8é?|5¾~¤Ó[ë†å$…w EmüÁƒ‹ºÊp«ôG€ŠÛ@¡Q:=ŠÔz–³IZõ{ÿ¹7xµþšµñGk?í¬ý—õµ?¼Y“­½éÄïðÐ¦z=·-`BOÓIò"š‚È$ÞâËù¤Ê‹4z#ÿæ“kÛŸ›O®ù o†ß—yöö±‹WY0&T;óÑlÍ?± ‚"ÖÄÑÞ7‡;{=+‘^$E	zÞ®³÷r¾’x§5âÚ–ÆNÂ#´„þå/¢·3NÊ2—ã—‹ç¸˜diø‰eQÞå%|M÷Æ~œ&U„«ÏíB•WÑDËƒ#_x$ÙYu¾ê—çrÎ($£5Ô3ô;òI½©†ë…
lóõ<ÇS`Ïk[9w¤¦vŸµìMµ]ÏT€?7Ôc(Ï%Õ\Ãøž<gTSëur½Pª|ˆÊ’R²%âÄcKäbËTÄ–Œ’ï£ëŒ‘[YAwd3)Ô±µˆlYã´ó2MmË”SËp ý…fðò‘7>B[Æ¼²åÆó²Ê§Ïò³Üg8o„_õ–èñnäçÀKµgJÊ:ééUß2u©ö¯ŠM³k ò}þY9±T/£w›>Ûy¬ÙÀµÈ9þ·Òõ¿©Ý	yloUøÍ5ÖÊêÞHŒ@õ°«L¢¨LJþñ0Äv>,òy÷q;úf’Ÿô_)H¯a‡râ¡ØXßü|àaUÓ(²')_9ûår{ÅívŠe÷‰Ûí·Ý#–Ý!î²?,Ëå—åñ]8¼Wzé½
×Âq4²óüOFŠPÃò“]ÎCûÚžö‚âP;‚Ysæ6/Æ‰,	"j/C‘þMò†V ÉœrÅæÙäªWã;†1Ð
èÓÒà\c¿”˜Š,‡.×’±Œ‡Åq‘ža˜Äø\¤.NŠü»ç—Æ>öBwÕ[Ë‘4¿ZA¶²õÕÑ|Œ’â-xÛÃóüò8®Ò+¡‘ž2I@w†"¥¬”s2ˆþ'×rZnÄ?~3xÂy)Á¢£B@½fµT`¹QQPàU¨´2_ò¢¿zO."Ùt¡&•¢¤(ºáÔ”¨´/T¬€JÚN¡ûeÔ±r+9Åè"Â)Y¬Šêtbåi49P»“®èáýFÓƒäŽç ãM’ý)ö¤r…—-ho$“h\Ÿ@¡³„ÇW<~þl?›Í«=XzðâëµÇIˆ§0¡P=Á ª³¤âs¹=½Iö -ðå@E¨‰—µS›AyÌŠÀ^'…Ò‘|éE_Sü>Ì‰êL¼CŸ¼›n=Å¤À+Õ4–WÍoƒÊ‡û9ªör›v#Pèj0-ú÷ÊMAuu’sX+Ôñ=ID;Åóô"ÑOi¾âèÑ^Ïb©ã± Õ÷Àú|
ã÷ÃYŽŸÓhÐAD…˜Hªï’«Òk#JöÄ¬è1±-ýã£H¡Q–&ôÆF¥ùp<«1UåB²g_.€È-TÓHÒ–Žßj…ãY(©"¸44X.€µu}Ì9ßZ¡™mPƒ0r±é“|³ùf´€ôˆÞ, ¤wNGØ¾ÔÕN&hOÎ¹z4ó¦ŸÚ ðp>Å¾2ýiz³ ¼o9–Ýr^šÎ:!æ5{n[ž5X-ç­]N°C;ä».
©r¦•`t‚nËauntœ¿hëõ_*ÙI|-Ö]n»£.S‡ÄíÄ©[5vµê+S5À¿¨6ùkfHí0JÎÎœ‘äþàhy¹0ÔåFí„q5ŽÜiàö£7¬×º¯¿Çm£`ëÀøˆËè"©GÎ¶ÕéÒqÎë¾©`÷y¹0Ôå¦Íkœ7§…ÛÏ[h{
®Ó€¾ÓºfyùöÖn·–Ýp£…Ûiñöè2û¯ƒ#_9"Æ
 «¡€=·_ÌþßŒÛŒçeràB¡ü"‘ Q|3ž$@b1kÁõ2ÝJÌðèjëT…Àî‚@£º™["P·pÔò–ƒ?Ï|DŸ.S‡t{äia¯w¦‘[¡ÎÀ¿Ì)	ÓÝ›å»ÖmY©Yr3>yC"lËF¬àÞA±Ò®7Hó¾} ¦XäÒ6’vÛ -üÛÜÍ¡‡ÌšõÁ/Ý8·']F$5fã+¨ÍÊó©V´ë„”ÉI%odÍ7¨
¨[ª«®£¡Ón¡–ê]b|âQó`|Ðáaý¦Ks!÷†1«p¸–üS6!ž*Ã‚ç3kú<ÎŸ`¸Dß^³~E_¢YõJºþÅxï×`~Æší—uý!!Ô¯öñµÉe>æó«É2õ²ÆPÛ¶ý²ÖèïR5œj_ Ç¤ëå˜O0ÀÛüòžO6È=êóëúfÃ¿VØGÛfp Ü°x}/ªÙÜ.¡£Å=eš‹<KÒøHP ÑÏâ„WEó"zÿ/ïÿw.`'ó÷‹£ðÎßÐÔh…·½ÐfsÕò4+…GY¨ŽÂpˆ?uÝq-àÈÒì4¯KF¶4k%“2q9i­uyâ”õÁoxI”Úæ›+ÜéáàZÍ~“Ó#èÓ˜ ËH»‡¸µŠ8Qõ‡âÏI‘ž¦?Ì‘‹S
xÊ‡+oF$¤§Bx…õÿ4Í¢É„3adÔÜ›àÒÿ"CÎÄ‹×âqlôßÇÜµ€ÿì”ÇÉUŸSåáC±{•EÓtL½Áõ2ÚÙyöLDQ:‰€]	ièìãÙ$Fƒüî7âw°§–‰¿&)læ°¡¤°E•ãªˆ;†6¤ÂÊ}’µ:/ÀÝ0ÔÈQR)?	üRÁ”_G‰*Çè2@àa{¤BF`9ru@!ë¹±dR©‰SáyÚ	±“Å2‚{Œ˜9“~ ]'2“3¥†Uþ,¿LŠ'Q™Xê¤½Kgè¬£a‹ê<ªÄ% (ùq6IÇiS'À¸~VV½‘ÚA™OéÙÐõxà6.¶¶ä0;MCz.†€ë¾òcMZ¿Q£ØÇõË5lEŒÛ£ò KTpçpoçÈD§FØg5‡ýh˜Á8Ì±+RÒœ%¹§)rx=©´{è‹êª}qÙv¤„¦šƒâ²5EÛTXý–òc[ÍÃ “ù°ÌJ±}‡F³˜	Júh52„iqÆ[œ‘B«±Z„0eK^¶tÊ6§×…§pÜxäƒ¯³XúC˜p‡J«ÕË8ÄXÿÜ‚i‡Û¹b’–•°«	z€ÞÝ*)úßä°MD™AÇz°°ŠªßVÅ	±’hHû[òˆEÒÇã…³jí›ÃFT—IV¦Uz‘VW#Ñ;…ÖÓQæ ¦¼R4¹j'rÕLÓª„UÅ«Š[|ñ«“TŠj{¤]Fm\V’ûSÀ	s™j±˜Bèåá¨a•Îr}pÉGšõ*kÏŽÜ÷£r
2qúñáŽG#—xZ%³;ƒ·îH»ï<sRš•‹ÀiÕ–wÅ?† 1€C&N¸$ßq'I‹˜’D-¼ápX§¥ÕÐÇºç¨)PÞWOer¾úZ’Ûò	{í,*½Vès}Àóe;3fÆH”hÒ×¥ÆŒáaÓbç‘‚*(¶Þ†ÔØWå€O–âlÏ½o“£¢ðüÉñþÁË£7{Ï÷Žv†©ÜOÊ¾Î1HN5`Wl³Æ8µé¼ÓkÍ ôÊ÷9YN¬‹‘–Ç®4f]Å$Áãè¬EUP QlÙùfä¥D"Lg«‘¯Ü“iþ9”
#Â¦S’«m5“\QÂ–T‘î6‰Lú%“Î;þÞ´›‡RÖ‡R~$CQÔçœ;“›nç=Õn²o{´#‚Ü{” D•ÄÇW³äŠ†-'×(ªjm¤%u8»¢&ðtˆ\®…O_ØWÄª7¦ ±ªŒ?0Xg—:îË·_Óëa“ ?ƒš#bèð…q\ný:ª°~öÎÍ’úÎ&¦ámû/½qÖUžàÑÀ²XŒÏiÓë+üß;3|¿í±6Û¢ÏZ‚¯ª ¼qÞkXõŒ|hmµÌp‘¨Ôs!Y©ƒR‰¤Ù!Jé—g
Ãˆõï`ü£‰5hz ÷gÝÔ–hëTDvÝ@ H_¿’Ã"rÀ®ªhÒùC`V»Ì©©ß¨ê€’ùT¥—Âžà±HìZB”Ò‰&³ßpŠ£Féä…9œj–«wI’þÚí´Cn7\õ*É†‹ŸË&ì¾,³ã°6Iúêgðs&?ê´üj 	Š}”Ö	d¸RÉ$9d;‰ü×,;u„Ùòâ‘ÛÎ¶óˆÜ|$ÞÆWÙ›O®©Ÿ7o­•ÛQszæ³7ÇÚ|5Il¡Km£@? x¯ÍòÙ|U	`ZîŸ†z
n•¸<O21Íãh"òÌ‰€Ï¦†rªŠ1Å]£P»ïÞé)|uö_²p–ÏÊK ‚ë„@ÚÕ%AÒz–<1R@µ¥wû]=°~« @ôÀ¦ÕÒ¬RÐøÊ±}\µ=\m—9,Zÿ˜LfÀŠª\` ÷9ùiì›Wçá<Žmt$‚²æYãV0”È=ï;¨õOP‰y©lÌø|$¾#_Ÿƒós×ZÉöP1éx‚Ú ²ã9;ÛóD› ãOKl™Ž¼$u:™CÅN«hU¥EúŸÏè¬ò8Ÿ;¶µ>9O€³\	•¬)Æò:¬¿u_I t(Þ„I	 ó&é8ÔöÀë¶Ìã)×9”R
l`h£†µïœÜ7Gøø÷‘§W’‹34 NJ ¬$*Æçøž¦‰íHx`…©‰¬,?èiºú*Cžl…‰ º!ƒôzà
FLùPÞ…:neQ‰‡ú3^¤1QPçP?âI':ð±=d%€òÍ*Ø†=)*ËË¼ˆéÌÒYžŸM¢\WàØ¯£<NN£ù¤RbbmÅ°‘ncÛ˜ö`]Qñmñö“kçÕè†Xwð{´OëJÓÂü ½*Õs§yàñð–ë¶›;€À‹ü"·0ÙÔ—\"5Ëyá3VD?EùQÔ3ñásŸ°	¡‰%:Õ¯¡oÛÃÍë5oS÷ç\ŽÇÐŠ'HñU‘ÃjMŠ‹¤À$X%40ë;  z:öx–fJª%<’¬†±d7b¿çy•wø+ëßõnã¹y	ÃÞ™)r>¡¯C%˜ÛE·MJµmÇÐÿÄÅH•ô6Xÿø?[çuŠ|.Þ‹ð‰ZORtW÷ìÚµËX28°º3¨2Âü”Àm6Ö¾¥P&Qûîmçà[É*UŠÏÛæ3x¯Ò]{tÿáˆ"„'Šñ±dæCy\F½q¹¨ê¶-¿Åh*ä‡—àåJ1°k»Šœèý_þjÉ; î?PgµÔ’e§·ô’!ìšn*!–¢¬A¸—¬ŠËwšú[£cè²<r¥—–ÔSÙR
·lË»¥Êšeei‘…
á:CXrU…ú¯£»°<.¦ÿß¼ŠVê±¨aU©Ô”T~BžíÀ~hF±0ç 6ÎßÿµHs—¥´¬‚J‚}¥`ŽÚaQLÂ¶x£H`5ê$Úà_ÎOÐˆv‚xÌ3¬‹GŽy”TCx±
$ëè“ — EÀÛÇöÍ³<Š1ÇwhTË‡5öãþóõPê˜‘Œ[Ð%a\ÅÕôƒg< pšžÐÉä8ßÓ%ú®ì1çÔf$høBŠFŠîí)¦Äìây&÷@ÑÃóZšiM;±rÂ=è÷«¨óÉ¦¶Y›—í5Çƒ6oßbn“P}gÈA =a(ÇI´„¤Ù÷z0*{PZàz)Ò
É8'qçÃ†³³’póµ1MÆÉq!Õ«(iø±:‡¸J*‹ª°²È¨µ¶ÿ1ìºä´ô³E£V	–ÄW^AÏ§bç`_ücr¥Ä†«†9MÏ§ êTõ“Ð÷WF³ôáY‚3ñPV]Ñ;Då³Âªì”ª?p?ËÍ{Ht1ÒEŠ|~*¹úyTQ÷±çD'ÒYÚ«ÅïÛ’;c\Cý„mÖôLæPÎèö¬!š-?AjP¸ZŸG	á8Æ°Bl0òCÜ¬íÈNÑQ|à]–_N’ø,!¾€È\)4°Ö©\0–Èb÷O÷Bß3º¨J¶#:CÃÎ¾ÑŸ )?K0ÙpÀ1š§0ÎŒLP¸¡S´pþó“\^&ÐÜ¿‹4¹”sgÊ“YY›®¼h•¢õl ÔOë5‡Ì8K-Ã@cúžéQ9¡úŒ†ß#˜@]¦hÅŠøKRáºÖ¶¼Ã˜Å=ebÓ»d»éî#w’ I’‚+Çs€Våxf£ÙÔ~ý[ï«
b¥¦(b>Û`á)n"µ×¨’÷yãXÏòì(‹f%ˆÎ}ÓáU™G2`xF+…2[tÎXÌL•=3º
qÓ{²gSçß¸LbN3#H'Ò‘ÚNaæXU¿ÿ½ðýÎD2øA“Œ¨™i\E
22‘%ãU’ö{œ¡ð­·!Uy[\yåÏ…xPïoÄˆª·x°W)8—vS™aÄÈ¸2ï}^/gIAÜˆŒËÏöŽ]Ês¨å/¿‚w‡Ç-óÞˆœ÷×'s¯óœÔÍä.ÕèÛ’{0©ÊÞàS>³25-€z‘Ð:pNcczëº
ñ§Ô2ÑoÒ“IšƒDarÈ<Î'ïÿõ4ŠÈÙ2œ£Ø×·±RÂFÑET >ý†æ¼ÿgé²‡Ø8Iò˜¹0±³µN8¯Î)‚uë¾¹ŸíÌ—=¸\ç<€ÍçI¼ªH^K”]á ®Pø–<ý„ÊîÇúf+ñr;‰½«[·5Ly£”bS*¸‹â<€µƒbô	%2	Ö@Šª(5ög^k,²»¥ºXÛ¯]AïÎëRKA6,²‰‚ÌáâÙê.j¾
RA0.é“Ûu
ú•$Â$Á§•“eS„=3]'V™Óƒ	ì¼wrP·ö9Ô`¢ÿŒÄPJb(91”‰¡ü¥‰Á™QNú®®b{ÃeU÷yQ•Gp¶‡œÞô%<÷Nn÷øyZ˜~ÇÉñqêÎ>ŠÕpµ–Œ½¨£MWãHS±'÷2ùQ§Ò—á0ê×ƒ?áþPg‡ï08–Ä«Ã™¼!Çª;<ÎÆPÿ‚(´î.:¸Ñ°ïÒQ¨­Õ¸¦€xŸ¬øç¿×›ºûÞ°w÷µZ~Þ[IR ƒ#{N^Ñ¬ŠšI‹l¸“#ÇÒÑH	RB)Ú²È,?ÉÏ(Ça™K…	4ÒZ©Úä
ã¬òÓÓIš%hÁ—Çó“v£—g§±Ä„ý¡Ý >êt¿jI5‰V;­Íx{}*Rû)ô¾Ý%ÁíÎnk„ÖWQ{Ó Âm$tyªvSáÀ‚.+2 ´UæyU›…w—ŽÎÞyjBWx¢¹óÙ5UpS‹1H`'®”¥€áOeÍÀh4è!t‚Öy Ë°Uz‘¿ÓÆOx„¦†µÃZjîˆã<ˆùÜÅÝç!šÉ£fn?skñÔÉé*)ÒÈú:½ìsøèú6k“æä¿«Mœü
œž¹ÆÙŠ;Í–±Ø†¦
SbàT±cœržÈð àóTtŸ'gæ‰Ú¸ý<!¬ÛÎ“VOZ®L’M¼W›!ÂG¡gHÁÒ3T4ÍPÑi†¼ÍÚ$áþÅù {Óœ/ºèçÁ˜Ï×¸û|aÌ|Qs·Ÿ/„õXbcpN#t¶ýÚ„Ê›‘ô„"p6¡ã.ÌrÜir•šW«?ð4zFéøÕƒŠÏhÕ}F+3ª•ÛOhu³YÓºgÑÑhj³‡R¾ÀõVlñ}—9ÄrÁÓ.Ò‚H‡¾/›‘ÜeD³,nŸÔü§VêA9ªo¬æR æà„ýo’œEã+7W†èkqï$9Åðo.z¹ôy¤õr</
Œ¯GÈ½R|·¿{;ùo<‰RéÍoóbR=¬O|—Äê"'ìöÍÛÇTðŽ{fLˆN__t
·+;I˜ª®M‹€h@xâ˜YžË$7CÕÍQ>^‡Ùf>¡ˆ3Xjõƒ.£;ÞÜðeo( ƒÀS¨vÔ8O®’ª›æœ¡1Yœ†Ÿ§ÐÌúcøç+áÖÂw¿ÛŸ¯×êê~žÐÚÙ—ç7’Ÿ8}a3u>ÏÞ¡:«Ú÷•¥«ØµáWãâ`>F˜ð`êý²/Ê…×W@¦8Å3A‹Øš]¯öOÎ„:Î§x»W‡×¨g›ëÆYŽòÓJÈP-ŒX¬„†î­~ºüËg1œ@Á$pzý!c.æI/Ì¤jë°•C&1lcèudÜñxÞ]¤GâY”aX›˜EE$þ˜ÓÙ¤fDªí§Ÿ
1‘õåÙö‰Œ‘sNøüŠö{ç ÙÜlá¨Ø2ÂAÕ´—¡È¾w&áä@¼Wç‰Ô³aQà2èÂBu÷8+½á†¡ÎfNGõXhåtKwRP6ncºÓRðêÝPH‚ÏÝ+Øhv8'F¥L HŒjÍ·ßŠõõuÁ®÷«•Eik}€[×£±ß@NÚzPûØÒƒZÙæ¨¹‘]X“ÈðòuÓ‰ËPÿtÏ` >Dˆ†«ÿþŒ!q2¨H[|ÿ¯*ï¤d³y1pø_ÀfgY‰½´ã”ŸsäóL‚Ù/£,é"âÔktFMœ_Århç1¡ 4¼%b–Æ³´*/ÅóÕß?#3¡Åé²1ú†½1û˜â!…©€ÍGv>‘AšR@.å„6J}W	[âÅÞŸ÷¡Œ˜ì™%àD@æ·£X&cœRcÆçDQg¬ãöl'Qm.<	UB
®Y¬pÖpWÕ{"» Å
©¤½•¦©ÙâF“5dì(JÓÞ.ÞG¬!Ån"loa[IžÝ€ßnû
©Q»·Ü\4’ìÎ‚È4 ›YždS˜q¢ð=õ’jº»VÒºéFbkã¸laq6O®ao	ÚpÁÆ`Q¤Ü'Õ.\ÝE;r.†BVvå)çPAƒ0JÜÒ»H\²pÇ(“—xw¨4|qÑ=ßä'Côar:Têµt¬7¬FC}ÌŒÐÈ¯ØzÅ¨½©¡òµ<züZ½Æ’({¢Ì°%'ñGß¶é‡»Ù´m?.÷3¬Ä"8tnF™:ƒ Ì?ÍA¡.Ýb@&ÅÙ<«"s–$IÇmL‘¨ëd:Hzá6Â“Y·G†‹ˆ›W¯Ñ‚ ‘Ç”PR?Á<öK%äY*bl3<Õ™‹BÔÕ‹³Û(PÚê/YÆ2Î‚Â²ŸYÆ3öàa°xšÍ375ÓŽ{"<’I@ÜŒ=Aüb¢•2¤%O8	1Ò—HWò ÐX5c°ÌÀËÿZÎ@Ç³IRÔÌòlMæ Ë!W4„†mJÀÂ2Z1Pn>–J6jI%ëÊ¤’Å¯l—%ží6´pÌŸd*ª.¼€Ky™˜“ÇJe#¸k4°*ª@¥ÎiÑæÙÏ\uû$a4qD“ñçOƒÚzd×2/¼)5Øá™<79dèÕ›žøé£Ã,¿„éùèÑ+yq”Å¹Ä£º}ù³/C<¹%Ÿ6WÅïkŠ2uèŒ±R÷|³´Ñc½3âV¶5•îg×qÒU§	ÍWG5.ëžzŸ±r³æRŠ¼¨”Cjn©2‡ÛCüA\Ä\j“2“¯Gbƒ×Áµ<¢ÿ:»OðÜw»ÀÈ?¥?)~$Ü£ò
Ï~YÇØIF;ê³t"YuÙ»ô®~áƒWxÉåÓèoõ,ÐÞŽ;œÍËó>íº#Pï½št£l ‘|Ô6 ®EÀêZ>¨4Ó¨²á˜~E”LÈA5²àÿ'ó+nÉà€™FV;XÙ{‚×›6¶ó\	Îa_
þÕoôùÂ¨ú0¾¿­e|ä)lsÐéCt#†¾kL)×pˆÝ¸iÒÃnÓZ_ë.†;º­ž“A¹¼{$¬f–åÖ^P¿îÀñp0-ÙZßSbµ-¾Qi~ð5™†bL$u™y*ž“ÙÌ÷a4k‘áC ÀDíû"™D˜,'÷Þ#bó³"šF2»t½Õ¸‹”uRÅú–]àåRÇ	¬ám[b£Î)èŒôSœõ«V!Ÿ›r¯öÞITVz Èéx'^ÃbºH
:î/õÚ£Ác2¡B-ñ–P'œrB®®„Ag?CbºÕÐ¨Ä‡-‡º»tà—·[œ½Øp°t¡"±†.mØ
Ö«ò*š¡ Â¸JÀi<Tö£ñH_æFñ"õbHƒô-¬
Ÿq2ˆ(,¸)~+•ï’ŒƒF¨ÔðARŒ¥¡AJ”tKŸá¡èÛæ%@é,ñ`:ŠþÔÆï1Ï˜T{yõ„adär¡Éó¼.è2‹>îæxÂ?%jña¸,¸+Òž/WRª=úªÃ\"añléñ{jê<¨Ì€¯“§ø&ç¦áÃß†“rþ#«H‹Þ‘Ù1 ÍÊ½ËsK$‚ò­èì8Ñ¼Êa yy¯€Ä0ßÿ+:‰[yoÝS£&‡Om;‡3Í#æZ„:iWÚ¥±„~“®j/“¸ÆÍ±z@î*ðŒ÷”"ÐXÛþS¦Öwzˆgã³«Á¡|ßeøª?uþZªLH#¶¥:àÝìIhAFƒ¬ýË_š!e³ïÃ@Ø‡–ú<IŒÀùÒÖƒd2ŸD#q¿µ@©€ Nó,yšþ˜A…
´À«%vÁÕ¿{Ú™™ÆŠf­.O¤›l/\Fç¤ÙkèGû\x†L,4r¼ò@¬¨¡·”Jº šœZÏµ[–Ï­PŸQÎå½ÛÈ2ñò¾ˆ@ŸŒÌt9@\ð¬•»š‡Í*˜OŠ¢gÎÿÕGˆÝµdYž››mGó/<Gô-%Ï¤ÜxA!•k²€ÌÅ%Oõl¨“‘¾ U#­•éY¶Ÿa¾£ƒ|6Ÿ©¤]èƒM"¤•HÉeŠ¤Üð†¨[p~ì$ÖRM-N®¤6s«ÖájBÎï$”¡÷°¸b•ûü“2ÏÖÆ“¼Lâµ“«5
üô²{aëB¡('ªƒG½`Àø*§KíD²FœQ;Îv‚å$?ë¯È™ ú ÈíØ&ÏQ¾)¬qóÿ‚ø	†ƒ é1m§,=M@Å;£¬Î{ÑI/ÁµÎºÄñÄ–éû¿ei.bL´("r‰É´rþÞ¾ÿ+É3…¼üQiÉ”ùè‰D‰p›J.ÒHQÎP|W¥h¹ú–Ïafe~P@]!öÖˆ›%â(ÉÎ¡í“XÏ°·žŸWÎÅ)€Jâg×Ž×Å#`ÖøˆO£ŸäÁ2€„ióÔV¸ÔÈnÒ‹’R£QÊ<³$´þFãjˆ® º”-Ò„l_¸ŸÈ¤®Zæ/X¾¤)=ynFg8ëÕ%¤B¹¾*¹Ôò“&4½ eß³ø@¥èUÿKŸäOýµfìA‹iËkÍáÊ¼/‹zw4´r]%À‹ÃéG†l™Û­‡«MŒÞ—M¨zã.ÓÆû<º'Êk ûFCÄZ4Áë-¯ÖÒùd¯™AìaÍD®ÓïHvLãùPÀ">.è¿ïÿ%bë§œGxÁyÄ–3	v-ÍðòÚxmlpEÜ/Xö”±³5›_º©"{-Ë«5Ê	Þ6P9F`QÔm<5Ùïë>‚Ë$zg;ÕÜØŽj$n Ð¿;K&¹ UT/€
à 	fãëÞôÏ¿wìêƒö‰öb(vâ#¥q30[NjÝ˜òÊëmÿ-¡Ôf¦òŽMýUëÈhœK•'ÕäB¥ýŒß@[&”0vg·…Õã%äs”š¤Ð…Œåå¼".2`…Ï0#pY“ºà`÷éByMÙ!ÉØiîâöŽJ™‘:g¦RòUCÇ†×" ³ëÕ‘Í‘‡ŒfÊ©çäÃyØÊ×¿m^d^'¯)‹@žÔAž´€¤C;T:ÿÆçÂ›5x®·ãp@p)},IGïgHªƒm‘LVÄ ˆoê Nº‚ÐÁÃØ¿˜õÏZ×+àªšG0^ü¾'	ÌìÊ\¥çÀ¼Yá2¥³ixG³’ø.€€ÒH©%PæaÁRÌÃ[~
XÅÖœ5ly)×=*ß—°2˜x,UZç#Ñ›0…už¥¼šNÙ;iã„·ÑçZI½±m *¤²´ñ¹7lïŸÒ˜¤¶35czÌ6füpÜ'
Ô®^Š(Q)òïbtÃ–—?9ÏÑ±¤nc9)¢¤´*ïTÄÂ@®~1cˆlQ£Fª½ÛwógÜ$éýÂ²Gù¼ùzÝó»Oöæ?üRgÒöiÈ1WÁ	çšæ÷3àÀ‡Œm!t¼@Á¤9Vñ Ô°þ_»ÏV)™œ·&ÔŠöösCBÒÌÞ0'tzßiH©ƒ¾ª<±Á\ÍvÃS¥ž',2`!SÕ‡C_¹‰Þ–£—nâÁëÏÅ9]ÈãR«x¿ßH®aÀiZML¦¿2(XÓ,n>‰.ðÇ«	šž99b:R9…¥Šo€³¥0Â¯Á~èuÿÙfS‰?&éÙ9môºÑá¹|õ[a š>/ñ.Æ %êd­k0u/¾í­26´¦yÕšmaUl¬¯òGÛ·U»K°§;GÇ=g¼¡Dòv”Ö'þ$ŸOèè.N¬šjí zgØ‘;Ì4Ëï“U T£ g3A‡F
bh?•Vs:ç'¯_S7‰S«i‰Æ»Òõ6ÅŒð4Lžäí6¾ Ä|ùþ³ùé=M¦w¯=Z_wª‘Ã-‰ÑÃDõg$ŽéßÍU!ÿ×{ê"“Ú?N~¬dû›Õ??Í³
÷þÆzøcVÓE‚’( qå$ŸÄ+^É
š€)yùâè»gÇ/÷wÄÁáË•*Ùøý;Mÿ½†ý„Q¼„ò<÷¼ðŸß4tõ3‰¾CàŽ?¥±ÚÛÇÑI‚YÚå+ÎàÈÏèÞ…^”ïÿULRÐ\\¨?Ð£ÍÏVÅç›|‰Éš×þÀæ¦+RJˆ¹¬ª¡lŽá†¡ýaøÈãvÐñ¯6uFÆtºóVÝ1Ÿ\û˜0ê)à+*%}¿´o	¹I¼ÿÞÆ Àw´›sX>ÇøÕ-¦>¾µ¸ùl=Ðß=ŒŒ•Í	<¹¥a®W|T[Œ¾MöF¼ÿ_e¸4J²ni¤¿s>€ü6×âô,­àÕ4Íæˆ"ó
Hò†wô‹ÆIÜ-¢KµÒ6¿@RÂUùù:'¦M‡˜ðŒT_þ|3´`$5êíÎluf7Ãs°ð÷ží¿ÿŸ‡û/GâàÙÎ‹—bw¸ß{ÿ__Š½£ãCøøß¿Ýò²WFðv»]™¸×>èjÅ*t¯¢‰¥)PM„<ØEùðÓêJ½lò{î|Ù†o¨„JeJ¦›P:"§í@Í§”L@VÃß˜ÇÀ ðÇö&ÕMÛjÛ¢o€²°.¼ó_Ñõß†Õl,$j)ÊÓÑ¶¾•Ï`_ä1Ä÷xcs´¾ÿÃk”Ì—ÁÀ¶í N	
²ëðËé·}nê4¡1Ðc¬è.ô^:j~`Ð"^|#©¨·³@(ä]OÞÀCãÂ^¤å‹è…ùfÕÕóJ³O!±6NOOñ³Žw‰NJ¿ê¹õvZ5€Ú®Jjœ¤“¾ÿPôé@ñoÅæ?›Ÿ0€{ÃÙR*Þ~r­AÞÄoƒ{£NÊ­	¸š¨sEçmÃt7›™È~Ö×Ü{ÃÜjÿcD¹[C_(+ÃØIà;ŸÂ–™ÛVú±ù¨C}^]õ‚j`}jpç›„/‹ IÄûèHsyÑCäIõ+\Òt²³jž>¯òcdÇ’§3û`Õÿ9Xüø9pÿ‘xõª÷Ÿ0Èíý?îíàÝýo_¼ÿŸGÇÀëÚ@ûæÙÞsú„»Áÿ³w$öÄó½gD)í¿¼<:Þ9êÑrt "ÜûþóÞ3|Þñþ<Ù‰?wþýØýîp(_ì¿”¿Žö¿“MïõkÛ¿“<¾ÙÝÅ~Ûâš$§ÕÈlsjö±Êg#ñù—« ¨!{Âh{vïLcÜm`¿ÉË\jv?}ËYF^Ì§(~íÆPÖï°“Z(L›pçè¨ºšà…q×N¤V	`¹RðzµY)¨´¨xÅ98%<ŽÄïëï±}¼H$T‡PÏ=ÉÛ¹Ÿ1™L 1Àî¡µÐÐpzÃC³}º½tÂÇðà^=ÚX_@Õ/½¶rÓï.X7Æùd>ÍLGÄuì6IâÔH –ýs@Blx6A«#ï·þ¦WÿóGî÷Ï¼ï_zß?÷¾¶é~ä}ßX8¢/–®ñ{¿Æú¢_.ÝÆ|<o´×pýŸž@í…«¾”‘Âú§øªgnï“T3”›$~ÿ2||AÎXYš$¥ÄK•¡.R±½gÇ;°­Ø*%‘æÐÐ7ÆVon¢žù‡UñûÍ×¤y™¬}á¬vîó÷ÿ8ûbÈÀ 66 ôr„\Ï}sWlýá¾°õäå‹'Ï¾{ÿ?v_.òGØþ€"½v,'­h;Ø{±»÷âxï^¦$Èâ£	™ß«ä0¿´üó÷MÜÄ6±ï›6_£ÅÒ0µG|iêæIq:É/Ñ‘ êá	ˆlïßWFÑï±KÅ¸ôižWžåSjÃ&X’é3«­q€¸$wÈ—§TÐÆ3;GŠ6èHÑƒ/~÷;Nè“ÀËÐz˜r•^¾7vŽ/ëß¸!fÝûLö 6o->¾ÁoM¼ÿ+È‘øä:½ÁÊú@už²:÷C±éÈ‚MîiÉ$`Ã9Ò0g™¯ÌQ«~\Â’Ê/‡ù/©ƒ¶òy5›WýÞÉ$?™Ç‚‚É›ä;“*Ë	µh½t³ÐŽ]<N0ÖJ(§›/ÒN3ßáucŒ^%?Î"´a>O²ù*¥ûd/^ËÐºpó«kñê]r¥mð¯GBÝº'n¾V>¶3ŒîÊAœ,æêiéÞð/Dvêà¹•l±ò7Üû]ågg“äÛ"Ÿãi>ëÛd¼q˜ÀÍ>Ks„£QPš%ú­O“Ý,8·ôª'ÏÜ¤‰Ê{˜ÅdªC˜Â8Â_NzÕ×VÇ#L¹íÒQŽV{âÅj¼êMÕukQY&e‰çz”®¬{¯d$eçkÆ;˜F6}cRï,x4Ý×Ê²äðê4Ò0H
¥ð»ÍN‚­òóbwCy˜>úE[Jb¿Ñþr·¸‹€5ü›ÐÉ†ü×|­¼ˆ.è°î%÷£Û{þà£“ÃÑÈ³*òì	4øNÜ¨X0ì5÷ÕÉÔ:COÛºöO®©ü¥¼|3““¨x‰¬Ó»·Ô„YN€>Ñ³u=ÎìÖ±r¹vŠÑ§f(±ÂšäÕâ,š­}&f?®}.fWðKù`Ö~œ ª"º*7ÏÖ0£`<—É¾Ðš+Îa¬0‰@º¢Ø®°-VNÎÖ˜ÀBêÎÚå94/Ês ¾ËµÉ™þ¥Ëìe$V¨šÖÅ9Š
#ú‘~GYƒ+
Î@"çkõøÎ™(aWÛºÞ\¿ñP¶ÂFM?ÑŒûšl³DûÚÆÆúŠž~¸éŠ7‚Ï½úµ>nÄCÓÓÈšõoÅ¡µy^¾¾&º»ùê!–“•¾z(IŸ<Ò5<>H»Ô)J¶9ÆëiáŸótIÆ©X’¨vÀ#¼ø¦÷JR"°¾C)N/øˆ çãdíjmcåk3ëî*	¬¶gõus.à³	ÿû9,ðÓ«µ“¤ºL` Á% d¦)Ïž|³Ý¼BVLw¾f«ÏhÃ‚œ¦ÙÚåÚú
¯ì’îÆ—œtýURžÃþ Xjjè¬8%Q{éJ0P¿§AÎ3Ì¸¤7´þ¦Û@7mGÝn^+K<£Ü)3Õi›ì;e›õjccöãk60˜;É~<dÁÔoâÔ¯™É'ª9¡ó½êŸ5ãËòúJý¥ŽúGãëx1ÜÈyrž\À¦{ˆÂµW]Mô>XŸc¹ìe![YÕ‹XÕZ‘WÒßK¬Ê>yý|Ø6ßœë(|¨¼Éê²ü%:÷¹qZ…Ú¶œé {ÏNÉ¥Ô–i‚@Õ>ìvƒm•+§éI,ðšø
úµNæZøç'ä —k¿ß´D¤¨Â#…3ð”Hä]ôkŸ:9Q)ªþ#6y6Bkz:Æ_²'@£°LªÑyƒlöÝ”‹´ñp@4¡küµdLƒ³ÉxSBÿ¢áDœË…0[û‚3èP•;tÐêr¦×¯~‰¥Îñ?ÍãÉÍ¬\=¶
7ü&*žœGEõ™6o|Rwˆn¾|ðÁ4ZipRÊ¿	w­¢ˆ%¶˜¿\Gò£<ÐkYžfÇÏð¶Ç{éÈ
}¹Ãvã(JL+\’'Wb¿zÿ¿¦°~¾Í§IbüvÔ£}Î¢®Ö6„^õ›ÆìèrÉD‹Û®¥ÖõÖõ).Ãù@{ÓÖÊ~öþoã4_q?JÑlëÚê]”€ÛáäµtÝ¡+¡Ñµë~áG!Õ×§¥‡]‡y”Th¾(ÃC}â†á/r© -¶´)?ð{ß²f¬v¿Œ‘>å°¾)Òä´1˜B5”oñŒÊÿÎÅ§b§•?¥{VŒPºµ"•ö)l]~ãQ}‰ßÌa3‘lÞøß5&e„Â­­¢¬™|l¥Vºñ@ú¢[û -RÃƒ<@ÃJ•w¤±Á,1FSçÞ–É’xšXG4à`‡Jx`Å°ÒèPf¨%¡jÜyN.3‹u!^nE[†~Ëê«‡z¹·2ã"¡›¾›.°Ëó™#/À³eÒêÊy·pŽðh)ŽÀcö9™ 1úA_Ô›î¤ŽC¶Ã%PªþK­¨Ë¨ˆ—’±»ô¤j|ûéhñ«þbLæ<¿;úa|öP›„1ß“Ø: (lM^Ma w´J÷Ó4`üŒ>©±/ui`¡½û¶ûê†ø4$è‰kã7éÉ$Í«d‘ãMçnë´—*GÀ¢q±TÕÜ†ÍÖŒgíÜf\äkvNî°¬š,ÔÖ6-ÿºØæZŒÒÍ*¢UnÚÀºýø÷]y®)/¤<®ßÍm_ïj‹çˆð¥(ìåµº¹£nñªï¥ùìüÊì£ÏÐÛñþ_"ÚCÇÀ2íSûBÌÖ¹émÍkÅè%çi2‰‰Ö¨Zø2=¬Ó¡Úz	­ã 
Àn&eS+À
jÌ`ñ5J›FGC‰Hi*‡…¾¬ÅãòF…•|fPgpŒ!Ø·ƒ‡¤N·“2Ç£f(Ï“¤
Êk‚{bšú¦èí³¥é­Ö¦YÙÕyÞÚöÝð¶ä;C—˜¤zå…“µxÌZi¬²ìˆ™§w‰¡òL¡wcã®c§AY‡õc²•n¿Æl,r,`uóÈÚ`\Õ-ÇÚ2ö¹c	»~Àv¼¥b­‡˜"óÿ¿H²óù4àñ›çþ¬ì®|cø:°KSß|‚»Ø½6=	€i6W®Ì•ÖÅ1Çs´u¸fË²ZùZFú\»X
¸¬ÂÍ_¿·ßè!ú„µGWå³ä´ÂX±ÏÄ¶è¶)¢nC*:¾O‡ô˜b
0Ñ¦wó¶Æ'‰êÞà]e}pÁA×Ú¡ù<_Û°¾O”šœA«­rGðŠðç˜V®…&Ñ‹NŒtBï£:BIüz@!‡[××âRâ¾ýäšN¬L£ûë«òô
¬4<°‚W?Õ[ÈßÒ”1°Ùsßl‹Ï¾x	y;¹*ƒ—ë}¶n²Çnþ·¾œŒu†ZŸ¼Ú+wÇïê«!
‡Eú…ïÂH§gþL–ÅxëZ&¦žçUþÝá3Œjz{^U³rôðá<]‹.¢**0Kçôa4KngØú'ªhó€–+ìÑÍ§'°‚Ïˆ¾¶6ÖOþðåÆ§c‡Ü:==}[ãÇÑ¤"Ó ¯mA—RÀŸU?Øú[ñ€ÉiRIqÃžÒL–¯éWnQ‡àSØŸÅg}ñ_º¾ø:&¿z8ë þÇÒç¦5 ”oª®î
«½ZLƒ{gå¹WÐòV'@H½icr#jšU=èõËyµHÏ;ŠÒBIŸUÔ¦å1ÈÌOÃ#“Áél2¡Eò¶Y “Ùñ•Ûêž8wž_çQYõWLR0ÊKŠ¦ö‡ÑO¨À©ô©t¤\‹ÆóÖ;k[‡/ö£ÊHH7Ud?yòâàÈ/ý.˜?Y·›ÆùÀŒ……’­Ú¤Éæ]mXêvjªAqócY$È^EÚÞ+úª{‘fïÿ
àòb?Co~"æS9
õ¨m|.â÷;K1Y EáÔˆ‹‘uðqu˜¨LQgŒÍNÝJ>ö¼ïí˜<xÚ
¼©¡Þ?KÎ€ °
¤nÜ	©O åBLó[”×ªß”>AgÓÇìÜûÿ7K"LDF1ð§ Ž$²n=G µèKøDð:.esÞS.sö^¥ƒñ>ÉgW:/Ã­Î±‹­ömÿZ'‰@†4£açˆRÔë-Ö5	«Ô2Û9‘ÊôGV¡ˆÏRBZÑç)êYÎÊ“ëó[,xæB,…Ýãýƒ—GoöžîíØ ]øìç¥;‚f?¡|"Gò2{›x§(¢«!ÞÎÚ×·ëÙ›h Úë?ˆ¢Õ¼HýÌ^z‚íº´a"ßåtU~xWK¹œ=†%2¥¬Õì«ZAnmy·Jâ'ÈÁXEâhúœ.q*çžÓ5Ycš@Ò*t€ú‹Þ4 ¸±”×–Ó`¦{$ìÆÃrÅ#…G
“¤Û±/1‘Æ(€H[ÂIíÝVPæàv°Ê>zù¸bçê`]>Q©mY
 <²ôÒA1å»ÜPeSUØÅËnR’¨\Õqé¯ÛoEò3+©Õ¨®öA_óã&ÁÙè6#KÍÊÂ™Yrvê3„·Œ%žoF=õ½;0j·±¯îí&¯¼L&‚b¼!·vÙŠÎŽw­ÌûIÍkmš'ø”è„æŠã¨ÛÁ“UyA0õ‹àíà‰“Ì¼á´“DíÄS]ài'Õiy]ñêßúQ§ö¶ÎU‘ÊÚñÁû‡¢`Ï­Ú¤V’¥À,aHÖS½mêKpÙæˆ¯wvé.DùWÉ÷WÉ÷WÉ÷ïAòU'LÝJ(œ)àš¾¾?™Î2ÿ* ;›Â¯bïRrÖÇ-ñê¹õs ©[T|Ùž.‰´+„Uƒ)²˜­Ú\³Œ¤+o³¸’¿QKÝö«×–Š²bA×O w‘µ¥P(¡ \hî5áªíJX4\Zz£Ö~ùå±%Ä7ivd‘¦Í¦G'úÍü·{È"ÁÇ×$bøÃÓ‚Yèýbá,T««€Æ{Û"¢aºð(.pÈ'ET¦“$-òûÐLèAwÌÍNm¥Ù÷…ÂY[åÂ™íD³HvoH¼•@öóÈ`
™ò‰ûåV"§Ö{±H^¼¦ú}ï20A+t0àõyåÀev¤F]²©}î(Ý ]Ö ó5ÃÅ›ÆR-Yã'N}«õSVþL‰ŽB¶¯º½'ÜîÎ±ÐvÈGód%ƒœQá÷`´¸²efÄ‘èa\Ä2Î¡ž@á*ðíon¨[ñþ3’A¡š`%‚†Ì#™\Q¦"Å/Äi’ÄèM¯eüœˆcŽÄ!{$4¼Êóv[HK“ãë­á]jèF—iuN·³Hè”ºù³uÛ§](Ë×ÁÉ€Ô˜{™~¤T’EÇ.-XBp¿ÿQÖÓoleuÃ£SÑ\óH•è‰µ¦¯ztÛ²÷=Ê–ä3#çÖG§®õ#à/ÜïÌê
n4]$Zâfšg˜rKDYŒIÝÏèøó¢òmÍÒá…¾Ôl7ÌxÍèÜiÑãßÝ,ÎîçÆuOHlZûrÊîj®~ãžØê`»r£Â©[ƒvìqÈŽ=¾“ÛX˜ÇŽœÛ¬	uå“!7vwàª ¡I†| 4ü§Ãýã½f\Ý¯‰œÁîb&_¤e-6•{’œg.ï  ÝÉœZï
Ê»GF6yNùW•íW•íW•íC©lŽéÜ€k±Ÿß·‚wötO½SðíÔ;Y¹«z—OâCVr«†HÈ­ŠÂ´!M­º©Ö™²×«WäÊ¤®Ó L²Ï”É= ¯Tj»(ÕÐÝÀêCg7×{¥d]¿CüÊóZ¥c%Ïúµê÷ž‡ÛQkð6ÓÐHeËö7„ßZQÑsC¨QÚ¸Euõg9XŸ)–C`›Äµ@c›×´ÌŽJ¤µòOgdçŸÎš-ýMýÓY])9/ò/ž¦]£"‚è |cœ’:©´LÐB¹\H}ž\)h˜´Wµ¦}#ÁO»ýEÞK¹¸©°Êï§G¢cò†z¼s¾v]¾m¿ žâq¼Œ¥[Å:!…îóPÃ–gà¶œ’ŽCÝEóré‘cÃ¸<>RäØŠh¢>ûõ•NnxS3Æ¦¼.h-ñ«–ŽpÖå©wŽÞÎY3¸:Oã*!»¨ÇŸo"I) èUáÃ¹¤qó˜â¾6ïPhøúY"HBêQ¡ðÚmy«<ü_eˆ–¬Ý‚³ÏÌ£éš•Ê¾·’Ï¢×ÉG"?áî×¨y­™n</
u.ÐšíT¤xÍ)(’Iì»cs¹±ÚÅ<%XlQè­ñ×G%ï‰E4™Ö¨šÏ
ö;ü,ùñ@†—vØì†­K¸ÉØFTÑB¾X/@¾@?›Ciš«ä]ÚdÄ0&ÞNÂLr³èj’GËä|9gù9÷'¯m\Ù6XDŠžƒ6s^@
Àjz[K[76?GUoƒbûÖn't ÚÛlk ÙE·,‹®È¬ß:_ÃÉ}=lK\i•|ƒýø$`ºDaMÚ.kQ®í…Tp(ÔÎP2á¾¶ª©Ûí˜#¯Á<ïEØ©šÌ†ƒÜfqäÉn,þuêðo©¬»uÂ¼£€‡÷ äQWîKÐÃ¿{öðÏ[qøç¬³”8É÷Vd¶Z#4±@.tÀynfÒ%!»7Î§Ó´êßÅúoâ > õÿm=üäZs ›·÷ã	¦[Ä•N2¥mY#²‹¨K
õ-hƒ¯½IŠ³yV9oJ$MÊ ½*X€ê¶.sê»¬æâèª%`­i÷Iß7Ú>]5dûçØ.e€l¬7m6ºwÛØö?÷¾ÅÚ=ëÁX*ˆöChoØhlìfê
ÆN{NâH°©4ñZôƒ’Àô#Ãz"ÏôûÅÇ´®ˆéDeŽÓß:NjeHiL9‚5%C(–_$˜À¥ßÃÖÞŒ'I„á"3–’-©åp}sß†o6 išvè¿×ã´OO†îåÒüÆª‹Éž½l»ƒû@ADGKÇ‚˜Y·ëâzW¹!ˆ:)Ô]ÏŒ]7¸†Ùœqw°5tVç0]Êi®¨$#cös™€ëžIÏpaÝÂ3Ç-Ì'’•\Í­NU=”À¦ÐÀQå>Ç9*vLsÇ_¹ëGÄ]•HÒÀ4^fÙ@å)3is¢þ÷Éú–vèÕY^ ŽÞ²#dŒ%Ý	.ÉRŒt¶<K)£Ód¤xWi£5€ïb&Ë°wûñvàZµºç*KMÙãÞ
¨ç­d–Ç¹ì]Òó
o˜{e.Ÿô,Fî¡‰†0®,š}eM Ã8¾þó¤¸êËþ­ŠËóf§§f oêÚÚêé¡ÜÂ5È!v×*åˆ>€ál^ÂëpÜPXk¸ŒŠ¬ÿöOØ3qr¥éAä™øäZaìFd9¨FÉè-LgP¢×ÄA‡EÆùŠ\\H,0ht]¤š‘ŸêŠ4âàÍëÒ%i÷lQ¤„Ä#jTâ7Ê‰=s’ó,SL_~«€iÍ™…2Ÿ&ýäÇ´¤e=õÛìD¨]ÖÍ¡©lQ ëœkw¦ÆQÖavƒ>N¿&IvV‹¯Å:oÄ¹ìs.ûjøáw[âóõõÐ”ŽÏçÙ;iâbÈÄôýtkRÅºq¬ƒ!»ëFªÚ15ÒAN?u—Ü…¼A›mÒdßê1©Á¨À1Î)ïi)ïÕ³Á§Ág—²³a	¬!À6»EH6JcGÊ’íVîÕÅ­ŒÕ~Àk;€„kH
CŽž„û>)IôößÚþáèå‹¡ezzåÕ¨²
Ô}‹ËìóÉ æ½ÿÛ­tš~=TÑÇ˜¥ËXïšÒÀ²{ÀÊ.Hc9°ÙÏa_€‰>ÒI"÷ôÀBhÛCÌ¿I¯rÖ`hOl¶]D¹_"WÇ{÷î¼hyXm2n›8ÉN‹˜\4f"0ºZ¼…Š—g«µDÊBAt9ûÍ© Ú`­÷vžQ?ì´÷, -GT¤Q5åŸ#é¨‰¼ný¦¨}zàÖÐêÖ¸HÇQþÆ¨…V¢Œ š}J§³h\åöEHE×ÊYi_½h¤a=òcÑÍ¥í8zßš’Ž´ï1¬½{P$v²ø)yúð«:¬ògùeR<‰ÊÄ.4ôÂ	ÅXÜSÕÆá‹ø>\Ú[t·Yó“EVÄübC'­ˆK*ÀŠÔ<ka~rGk¡û3XeKJe."ÔåÂ_fñàƒ§-Ë53´Xºï§z–Œ‘0Wšs~‹µÑ˜gPºXs^ '·Úý|‚ø…í~
µKÛý,ƒÙý~e¶ÿÞ™­2ÖµðÑÃ?;ÉúûˆGT–Ä»ðÃ %QóÂŽ–Äûµ.Ë’´"bIŠ°)Üê$*Âª¼Ww¡"¬)Ù²´"lg&¨k°‡"¬±¬áÛÐçÝa‡˜ïCV£úa#“üÇU„²¹^à¹ŽŠ+ß…Š}Ú®ÒÝµ]ÙéÉ„3¥¦3®ilÌ¨ër6ÓÃzõú\ª÷êµ=®Ô Ú@Ýª˜'½‚•çFÿ³«0šZÖVÞ‚·±…œÒJfsÄIª[S›¬šv» ‰HR¹ëbë;Pô²]LkÖÜ•Ö´½¸‘Ö´•ø´fíœËÓ«ûQÑð.Ú§µèg§5×–|[Z³PÐš´âå“9pokÄ‹²«_ƒ>þÃišÖ¬›j»U(f$,Ùt¶éÝK°µ›'®©^ÞÉÒW* ÜÐ§°Ö)4Æª¶¥¬õDä|Š—ëÅºvúM9?KðÊçRC‰~ÊõK¥Óyÿ¬Œ.’‰÷åÝ,5òÒ”–·ñ¾j¤ùÊ÷Ô:ˆ)‚'®xÚ6¢Ff}¸EªµÎH;Ó r§Å"Y#[ca)éN¿°A²¬$Ë{1H*‚ZÖÉ¶Šº9ò×mã?â¶¡˜6m	?ëVÐ°¦¸ÕêðUô+wÿÙ¹û‡	Ë4|ý£ˆÊ\š¿Ju È_l«¶¬ºd©>é×,«zÕ¸5‰3Øº¾]Õ,5aI’MQÀ¨j€~Ê¡%×_2ºÈ%ê»ÛT5Y~p“ª•²þ£ZTõòäÕ £gUU#XŠC±ÛJ 2“ž>¬¦·{4Á2îÕÍ*fxœ#8ÛÒ61¶:—5‰ñª¡E¬”Üïß’_Ú$æ1ÀÛYÄ®ÆWvd®•ÎÂ^Êù•U	Eø¶V	íØ>5ë¹øÂ–ßé/EûKúneD¾ó*ktrÜÍlýw¾ÂE$N,yŸæéDœÐâ¶þñ»¸TÂuú]è!ÈFü¼jå?-ÞŸ[ñïÎå¤Ù_Ùú¨‡òs_þÕ(ŒNñè‚ÎƒI˜ÊÓ‘ÊfLÂêfk†Åi˜Ú!ÐN%Õ¤2¾-„Š>|(ÉÞ+Å˜™E€b:+<•s’ ü5•÷]/¯cB1©å è|É1“ÉÎÈâP	5)Ï2-ÕÆ–nÃ¢PÂ%B~»#Þÿ§å“kƒÃ›ñýû¿bÆ]ø/³Àf¹(zúßvÊì›Ì\,Zã˜i™íH÷h‹y‘\bËûÙl^õ{½Î<òíû ¦áàüÛ l0<#­·X>åå[¦emSô¹S-¾Ô!|èáZm<+¥GZ(ÉX^/ÏŒ:2ý½•o“f³˜{|Ö`*z šØÄAÅ²„ò‰N£âÙæPV<C±‹´”;eoˆ¥Ôn+ÿ06QD 1“å—¹f%r<Ž†§µÖ—,l×Þ8K—ev3Yû
¾3Ô7Ku÷ä–-q¹,n†‘‡Ù‹ƒ1:<¼°¡ÍYÈúnØ¯jó±×â¾£‘j?A­/žP³›”ã¥Â³ÔvëŽa0X|kô4Î|Ø‚Ï­‘í”ËµÓ!ˆ/Ï7²|ú©?‡ÃsäØZôj¢Më±jærfQÒef&]½CåB¯_.¦7!NåÅ‡§4ü´ìºÂüÊvg½}ÞcCUwÜÄåp`+k;¡<h¾·GZ<êbUB 3}ù&Q*`m™Uc1Ä¦	•‹m‚üÌ»7p¤ÛNÇØúƒ×àn9øœù>ÒS§±á5S¸ö÷1rTO;N{èŒq}Ú%¸Æ´³Ýö¼ÎÈúTº÷Êð
™¨Ü)×ÀÅè»Í†ú·1=aƒdOŒ/m—ë]!QÛã¹ _$ÈÜZÐÏ'±›4d’6Á_•'‚*Š¿Õë â$!Ë²aÌ¶)/'-öÕlM•8ŒP´…×(œÍ-©þ×Z’•8Œ&%å»ŒT”ùH=©5RZIxò‡µc@'ÝT…WŠš”Nr§VaŒŽËñE¯^/P&P^^“Ø²Më€¬hÕWÑq10+rÄÍöót·Ò-X/f:|¤Ö‹ÙBáÚ8–’¬(;QÞ¯ Š3¿@
å]¨…[5a›0\4f§}çh¢Æ­Œ-‘šw-õ(°&¼Ø¢¥zÅc–ïÔ“Ýi»°²»µÌ½pÞw–éMøÝcÅ%|€‹–º@¨ÇœL§8Šýl£½íg•ù*‘µâQÖX…5bl^p ”U<å_]B1¦|UÉíIë}ÆZêÖj{¨Ó÷¿cÅe™®ñ_ÚÌÏ…¤:¡C_¤×Ž­ÎÀQÉø(¤þ]Vë }ðwBµƒªì&§Ñ|RI7šáK½‹£zø¡%O¿Qì'à0Æ//O¾ÇIÝ9ÜÛ9âèÈÑ~%oS0‰‘ó,ý¸œŸZ0LÓŽêÐžZšK'ËÑ·×Ö@Ÿ6Bh$	YÑiráå²Y9/v¿ŸÙ+f¶(’…ˆP†R×u­
òßìÒXý•Âåi|Ý|>€»¬CÊ'Æ€dGt`Á-ÍE@C|¥6ÓøD1WÞ_Ûƒ$Ñ¬ë†<}˜¹ÑY+‹Æè,ôRjÂÕí³Ûþº;*›\•Ž³²ÅQ©.éÕ½®s–¦XkuìQ×“„„¿;…A°²þsÇ`ˆz†¸ í[Ô‡NÊh$Š†¨ˆn‘/->CÁX1¹{H@íÍÏ`xÉ²G˜aù²nÐ}å}LñÍ+¯-D`ÉuÇŠ4Ew}ai©±æ„lXm-mÃÍÛ€63÷ö=ŸV’}¼•¯†væ^[¼:kâÐ@7ÕÄ.£6ÝDB[$‹ãä8q÷í[ñ}yZÔ¿·µí/X“]Íû|‘~¬&~k¨oÝfªwìÊlÚ]øDwvû`Aß‹©Ñ5Õ6X­ËÍ[Âö¿#Ë2nÇvLF¶m¼ïYÒøáXÓß«Ð‘ŒîË¨þ½µ%ãÎL–	ìL&{?Î˜5<u;Q.nÏ–¡¯0wÎm87ó‡ÚÕáðGø Hý4¼„}-9Jºfóƒ§ºÃù’åŽŠÈ:ú:½eêx7v®gm…KTª$òßÌûnÅ—©¤Ë	Y©ŠI:-N8Ø™ßZPçÿ”ÎÐÐóªÇgÉÇ¾ƒU‡Ø"rä88}dÈêndËµ>…˜wÃ¿øæ|1Œ—WÐ«‰e<'ÑøÝY“´RÉ<qü;êÄy]óª](Ãp:ö¥Y#tëz˜‰„”¢
_TC6½ûòºládlè’[“®Ž±ãžÿ€)6%Z Œš«‹œ-Ñ%Þ¹ÓÈAœ¹ÙÈEo!xû‡@ÁÅ7½˜sÜ÷zÛ‹^'tpÃRÙõ¼ErÞì	d\%b`u6ßy€ÆÎÝùl’ŽQbb\è)Z%œ1.	Å±öCihy™ArK>»ÂFÐèÞ-—‹¼‡jhÀ¨gZk‡¥/újKl„ì+/Yçïÿ*›|ÿ¯²MËa#"ùº±!Ršæ­&À‡ÅQ^TxMü»$™‰ê<' FàuÃ}|Â—)t¦¹:ÿÿð<=;OJpróLÀDÓ÷"0¼•ÐÙXÑ ÈþzˆŸúýhUœ8ëO¥Ô Žºƒ“G5›ÛmØŽ¡H\‚*Uûh¾ýVl¬¯¯&€ÖÊ¢
¼>Â:;†¶,Ûö¿öOÚÚ¯}li¿V¶©}ENX“ˆP7ðNN¤Ó¨ lÒ9dDô«õ×ÎV ‡{þKª¥6^ùý˜è»V‘T]Eâö|3iATI2ýº]AšÔ6†b¯Ç´PRºRŠNP„íFÔöÕ˜qþÈ;ÄBwF„Q¸[×m¬Aô7Mº–{@I®ü¢}ÌÉE©º0Zõû¤²å% ~»fÂUåÈ©g-B±!/J¸…µ}9§BFRTJCOª¨åÊX]fØJKú·/ky¹-–Ö½Ì3œFo?¶iÌšæ>´g8ÈŠ&“gDqˆ¯(ÞQ–T¼µ/Û&LèªKP5@;ú‡‡Ý‰nÀ– ä$¹±Ë8 -’IÃkX¨ô´ÕÊô·rcÜ:™>¹ð#Éonœ}±§ÅõÎˆF:Ä‡E&(`S›C¡rô<¡ä(R|.Í¬aánÒ$71!nQŸ‘HöíAÐÅ´-yáyðˆ]œ\¬G³òÞz£¶ZŽÌûmÏÓIL—"H6tJÊ¡¼B‰VÔ*‰žÇQù®Ü`¿7åï3LŸ#?ª×FÚ= i0-“!,Šþ+¼¨ßÚî³Hð–wGŽ7Y pJž&ù0eQ¬.Õ‰%5€Ÿ¿#fkù9:Ñ,¿I.Œð—BÇ4)$==y]D“4Ò‰Fï¯CN^ûËÈl@êf_”‚ÕJ!õŒräÚx·é¾“+È}+‰^0³	þ….#U=Yp©íùbMÿ4ÔÚ%¥ËiŸ* 6e‘‹Òc5©©2#›dó¤IÎgúXi¶YJU|—íFñLƒ76{¶šêj–PÞô$‹÷pNcµÍ~Dª-(ÛC™è,°5ëOÆ¥#¯é•&¾e7Ypšm8ÍVpW©þ\
pA‡SÂ5Ñ©’‡=Ýbr¥ô
$J¥hÀî ‚ƒøTÌæÅ™ü`SÔÊœvý|ŠVJ;ìwû¨Æ(ëå¸H’LœÂJv˜ÝÔ–Xgpô?Ý4–b¡î^z”</¤Ž¸íµØŠI[Ð6’åu¦àÉÃ&%EÖà¼¡£®â‹ýºÿÍšˆ?.Cº½îãôÀå-ÑÝÕ›®
N½1Oê}ê6÷¢Ý„ˆ‘öÐÄè	Mj$(ÍÔ¨ñIRVëB’²d;]êVXCœ•%Îöõ]…—uw
ö½_ó¡Ôî†©¢Êè¶IJîí^*§³%I¸)šà«”–ZÌÞGåùI±÷~\äY~VDÓÈûð.ÊN¢Ì{Y$“ð—9ùyçÿ…”@>®æg­voù¤Dã+l³0Ã1iËH§)zÜQ3-Qi)Û 19N³p‘ÁþU¨}«üPû}§Èö]¦¨/á¶é®—“±:ŸW°5EYIÿ ßÄ‚ß¡­mªX.:É‹ªddç*ô*}]©çòÜ!^7 ÕïÐ‚[NíŽP”U»ò/û¡Egí©C÷Ü
Ä‡	Œ6±ÊUÌú î'‹mï‚'­OdHÝ*1rœ YÆ8c0ì9E{‰`YrÏ/Î=ú>àÅäïÔ"¡dªÀ$ñô$`“ø€½XÖ(ñ»rK«ÄR=ZÂ,Á\C>¡ûÖŠðgn»h)±ÙV‚Û5Â%´•ÃYp`ó,ÌÃØ€ÿ Ää]	òs¸c·Þ­Wd˜ )®:œãï×€r˜¦yRö>"C6§õ{¶´¸°›í-êÊÇâÒÙÞÒb_i ð÷’Òœ"Ê4Cñ)ýIùþãDîÿ%pLh9F-\y”ñºpL€#DlOo<ÏÆÂp,0­¶z¼Î=KJ7°ug@zW®ée‘¥$llº—‡š=…Û¥névkSÍ3k“Qv&–šèÊù‚Îž%Z÷¶íq´wµÇØ¾7ÃêäK–ˆZìL^ÊBÓÐ;é@ö[;›«â
•¾`^+SŒ$‚ú”*ï’÷\ëN»rYµëBÏ˜ZªM¶-Æ¨û·‹t˜í_í#N¥}ä.«ËßAYSÉ/£Èë§{–z­Ã?Á>NÌ8\y×73Ý:n•h4÷›‚ÎvŒfãö3½®hÉýôÊK/i£I‚ûá.„¸³RX'2 *ñbD“²V%$šra¾ƒdºœl*\ð÷àì"žºªc¯Ó³"	CÕ9mAxËš©,§Úžbf5 aø3«›òæ¶Î$›çV2½Ž³»ÌÜzÀ?£×W<ZfÒUçÛ§²ÎÌôJµ³q~KL0ŸbÓœ7Çò}ç%Ì”å´Šy¿ÔBîx>-m)·&‡çQÕH‹¬6·$	Ù¤Gð²31h»È¢þ£&ƒ8-Ç÷A » §‘¼Fî>÷ª5oòñmçÙÇÂrú-üzþqKà–ŠN{»vô7Í6Q¸ÛÖ®[ªÍ´²Ð6ÑxZämíø/»µ£l®ðI.¯´*“ÉéoêÐq®ìHbÃ”ÜÒ|™x)Ô¹?k·ŒÍÑ‘ÁÿæÉôŽÌ®ÆKÕ¬ë‘kÑ
}Kb‘ÔOOYƒ"U)óy1¦ïûì†qi7¤ô{¾ÙQ}Å*hd$®µS)Õ;Yy™ðá$ÏÑŠóØûþjÍ'è>êeÆ”ÂÉ2öØù„§m¼O7xv­uŒ([¶‰_ULÑœÁ0ˆ©›IAs¶_hgãXôq,
cP`¬ÑÓÞ-j0îœêxÀ:„©>œ²NV¼IRøþxPÊaìg@ÔùÜ®@ÔÔá;èJ™ê3è1åU*‡+Œªy[ÜÖ”Œs‚iA$EíD‰m{B*ðõw¢w¼±9Z_‡ÿõü8•‡Cx¿«C))1Ø69àÛy:ÞÜ
0éí›4†‡Y~Y?’E“ƒ"ÿ^f¥^WuB•^oÀ’Ø½ýäÚ™Ê	q#Ö„ƒ÷ÃDVýO®íÙmØZfÁ»¼}ü›ÐÀ_ž|?î¶ó7WYïsxdeËxgÖFÏª•U4‘;¦BOpS9ÿÒx[Œ20:àühžkù¬ÏŸ}-Ö·Ýn…È
ŽÄ«ÞfÑªÞÿgi„›;L8Y´è€úó¨x'Yî§âÏ¨ÑÛ§@Ù8IÊp˜Œç”œO£LÖ;NÆY>É$TÜßé½¶£GÈo€Gw@ò}D,vÜÑIÃ%²0ÝX H„Òñ>réÚ-ã}ÔKÛñ)èUPg¡IÊØ‘)¥“•sŒ{–”Ë‡]`cg=†D>„(VÈ„j´¨(sÉj‡€ÆŒÕ®5<²‘¿“}œƒôIÞÅcáwµS gcOo:õ§Õ·É6–¦è¹«iÔéº+ÚùèøÁ¤í±ÞÈþ5±BÜ“éì„ÎÖò‘Æòò·º?Éç“ÏHMÂ#/ï¤{0…ËÓ'Î¤çûà¼™¾Â(¥W¯Ý=ªðËS 
_ £“÷c•¸9ý7A*qNO#ÀêN³¯¤@õµ¿7ª:š£ €QYÁN•Hå
¶¾ÎšwÕz'O£;C^2à‚s>Ó43çÏô7êâZbz¦?á6?Î’âlžU‘»½‡šéÕSÇ¡¹œŒÄ‚¾„+ÉÇ•£¯þMñ(u¥OêLŸîÝœÂ£|û5½Ö"¤s•tˆ¾î¤S…õsÍY:Ã‰À&}ÌcíÆé¨£ú@êKˆ¶ð"t[†úu"YÏ™?½ƒ½3™~Xå ç_¼Ð^]YGg˜STuûÍƒË>ôJ©7ûîA§Ô»=¹¿Öj‡äµCrÎé¡}²Õ5E-ˆ¶O1
&ë…œÎè:N"Z€»ÐïðÅ¯57õ·™Ûã±Î0¢8î³kçùœþiPçX\~cÈÂ¸Z`h Ÿ£ÜêìJÞÌÈoÏ•½È ˆ‘”„·oÔÒÝ4ÌÄåj},ÔkäþiP[Ò$yâ¥ÞÛ »»ÍŸ†zDµ®FùIyK°P3)."ÌË<KJÔ¸fÛCú…kžÅ	`.ÁŒ3ùz$6j)­4ãQ:íAùM(jÞDÅ8Åµ»-6àýz-Œ‡¼Hïåw"(¥_yJ/ªÛô
¨Œ²8—Ó¦²ö}öÅ`XÎOä¶Þß\¿¯Íu¦J¦VXq4Q©‹Êó0OGÜ*«BXÊÙ­:PõŽ6ªío½žÅ–›5—RL¥bvK•9F6ÃŸ-ÕŒð7ÿTŽ1A!Íÿokpôl8 ÛÔjË[‰„EÉál^bŒ¨œ+ŸÅú>Ÿ[yÊ7ðØ#9±§dUe:Åiì»RC;=,C]i¢+UÔé¢N­´±€:Zè£9?þ¹§Ãú‚¤¥3Ø	×ªƒ™Ç`úN¦f'ÐÊ(O:·<ófcmKh
Ö|eeúUÓ÷µµmGI]/“¹%ì@@£k}\/è„$ÖL¬B‡nÔË¢"‰õ©fó]³gÓZ³˜Î~/—º¯ÐnÆCùwKb=ûê&Dzã[dkSÈá·š&á—îbJ³%Ã5ùç§4Pù+ÉjÒ˜ÇÅ°ÅÁª™½.Çþý:•Æ~>ÏÞq3šÛdÍÕtz+™¾B×Ä©rmKw#à¬EeædúK¼¸¥o|2úÇ ¬}¹—õ¹ñØC<¦HÐVUn/â%“«¥ýˆê‡ôñ¼u|…ñ5°xM™>0è«x ^B¯/r'Åf|ÿWÜ31ß2UÊ£"ú~L¾¾5=Réì,Õ.9{VCÅæ¶x<e²gê}áôÆOÕ¡\[/a-’¹˜OóÖTw‡cüX¾gôv­	ˆÿÙ<šÁŒ¤Kgh•®YÅh¶pÞ,	òtÛñA’ÅP’¶X=Ì
ÞÍjsìÐþÑm53¸>Yäï(vá…eBÔcÞ^„°×‹¤™ÿ  ÿÿì½Ùv#W’ ø®¯ðà¨@‰—XRÉX8IIìŒ…I25Ý'2&Â	w’® à»ƒ‹ØœÓ5ýÐOý8 ê‡<]çÔSžúþIÉ˜ÙÝ7w	†"³„#÷»Úµk×¶k¦á?…–
ƒ­Î“hßëö¶ØExî0è˜èvhvôƒùe\R‰„‚5*3¤áS-‘…ŸÔæñ^ã[Íb@æÓŠßêó6Æìø‹ÑÃÇ’Tú£Ü6OÏdlæõ‚§¡õ¦yBw¶÷”±/ê¼NRÀ ÿ&‰¸x¤Ï´%É:Îöƒµ7¿cCÿØhho²{LâÖ2Sh8Ñå á@ÉI²sÚúÔÅžÁ Æg*Îâ
KÛ“$!Èºæa.‘8¸–S £:aÃ¥Ã||BçÚi^¦c®öGaàÁ#Œ§Ãj¤‰1m®"­ o¢ml°éY<®ô©é35¬ešÅ1§Ú$'v¶¥ÖÄëU›…/G¤íBáLÝ4¿ÁýJ“Í;*8Åð]4ÅŒAC&?€)j1ñH9 ðÕ=Œ1üÏóªyPb)P_G59Ö¾&î.5o½½QÄ”á0ú.eãÌDäaolîvcóAœy¦·E¼¹·3ÅZc2™ã„ò5ö£=`_ã³¼XŒâ$dO¹œÆ‘ªPØUnþróo)Ÿÿi|”3àÉ"yo–®–Züó	t²ä{˜ÜüËxÅÑfQeÇ*ç8i`MÐUùÃVŽ1îK-BK©51y«u`+­;O"õ®áÔ¡.²$%•#JØez2Å&%æ­rÀéIV2Dò´,²	p©be1úº‡ôRäíøã0öbZáJåý/¾x³¿»óúpóæ¿ßü·7ÑÖþÍÿ8ÜÝÚŒ¶w¢Í×7ÿôr÷`':Øyuóÿ¾f7¢½ýïþ5¢hç`ïÍ|_ÿbŽµ@ƒùøNÆ¸ÈnC¯J-”F]¡¼ê!Ï,!ˆž…ñúp¸ì!ˆâ…#˜EŽ\-_dhRZê‘.âå¦)4õcšq€uÓ‹õháL¦ÀpQÆÀ
%ßiãÙè5(nþW’‘. CÍqž±F*JæC(<¬˜/ÆÁ ¸½öò²œfœæI6šÙà6zëˆ;KÑ¦šÚÂA6ZÀ.pö²äðÄÚÛóÍËW›ÑrôÝæþw›/ß `7_ö£òÁÍ¿ÀZü°Ãœõ²BÜãÁt/öºr‰ÐãÞ¤ÈxÂ4SÌ’( Öw‡ôjš‡mã©Œ¤GhicìÓ¨[fÕTôGz•QˆÜŠ,•#Ã%Çæ ø>áÝiT¨t åä«Ýÿb-°þDG±ƒŸhjŒfy”#å,n~A**Vž/Ž¨FLóV¢µ;‰÷Êøæ/qŽ2Êzã_<6àðêÅÑËÍ­?½voËŠ¾Jˆ1)!’91zÏêêhý!•…S‹.cÀ^"Iô£š÷ú_<ìG»¯öÞìn¾>ÜYâ\ ©Ýy;Uƒ:°È3.iµÓãW’ø`ÕÀç¤âx˜ÓÖR–q$OD©€¹ù…6|™ŽnþyŒã‡­•ŽExËP[tbaà CF­Ãèù¥ÒÁÍ¿"½€Õf ¤Pu1dð_Ê#³o2`[H[XIÔŠ2rÝ\ço—ˆ†âü`
´`HÌ¢ÈÕ7dÑuŸÒ‹n’VHw’|1²K(}ä“v´Î²œÝ£×¯úÄêeÁ„\8‡ðVzs©º^¿»»GŽý² ªÄ€3PçbH,¦Usš™ZrÇqÖåI:,ûBÆ€1 ÉÔdU* Ûë„Îý¥‡ý'KÇÃ¸<]Ð’î²J  vloà´_·ôEb¯@œB…#´O˜–¦±üc™<6 ¨q08%X¸
±ŠÚÁÖú›ûû›ÿÙVJF¤+}UÊo¾ù;[‡nm²› ‰ª²4ÐHdcz#i-îï¾þ.ºöµ­[ef¨Äò§ÏRQy¦µgšŠBùª›a†zúi]ÍÖˆº÷jío×–Yñû«é0iƒ
D=¼‰|ZW k¿‚·‹Ñ[-"ðå@³QcÀZà}ªkz¯4Õ_}å*´C¥bo³w¶ –°TõÜ™}ºÏ²êŸT…7¶½e¾XŒ2ÇèQTP»¦ò»cäöQ\Óº´kaÔë|›(†/oËV_1§?ÄÕ›UæŠZT¿–À¬¢ÇŽqC…XÍË)®Ëñº%`Bë|’îË@*¤œY¹¿LXówíèîÂz5.­Ù)£NR;5t­õ`ÁÔ¶uÓf…u!tTèy¿JRjE,&!qo¦/(é÷ùô,ÕµÖ†¬óOpÂE»›†Œ	ü.ŒÙâG¦ì ÂÃyJ,SzAŒ,ðþÑÀ'¿Æ¥_¹º~’j“F³ÿçñŸÇÛtêÑüóx!ú
XÍþy“òÜS)Þ}*¾ù3x1Iéø¾…£ô x4±k•ÎÙ.‘F¾yä·ÝBxe®Åç×h=¥)çæ¥=ÁZœ~mÚ%n›DÇ•ð|´&æmûT:Õö¹ÜB^6*P XSéˆõTH•Ãjn!—û}¨F.ö~2hz×oôz%ßhìÕ'§Â÷vºVã’'»ãÉItÂQ"–3x¡®{rB$5®‚õl%+”ËTr°r&ÿó^™}Þoßéç#W¯êsçqLJ¼z‹Ñ¥Šh,’pAüÜ¤Ì€%èc.µÝ$¯¨òyñ²ìS6ATEÄMœkéµKn‰¼m34©GYçÏµú”vWEçµk¾¢å¦véƒÁ˜ÙMüŠìÚ–lFúËkºA|¿TXÁ½lðQ¥Qãqî@¢WÃ¦[@G ¾ÒÞ…¶`IuÄú¹ƒò¢Xìeb„‰ßqTýhÕ5Òz+½[÷H[”Ñ å@±®ßöÞ®¼ÓúXwg(“že±V3z®MBÇÃZàxQcèåÓ-#ø)ŠœÅÍÆž@2öÂù·¨^ho"Š7Ó‚JitJš‘½‡=H”Ë)ó’¾Õ@m³ÁÚºÈX6¼7#¾ëZ˜0ÓhsÖŽÃå¸cãë¡Fò4jK³¨ßÄÉˆil†CMfqÖ=ØêVÆ…©Øb0ÉÌ­¨2|«¾%s[¼·¸çõHö5ßëgöz´ðèÔR;ðÆA¾äˆý"lÕ×~â`$CawÃelh¾Qa¡J²ÊêÂ5ÂBü¢:L¿dÕÐåãJóÕ
]›·Àœ+x(¶bª_´p"Øí'f¤  ‚2…†næ³ÅvÎCiÇà;¦¤…/çƒ³Î®1»¢©QÍºˆ¥¥Kž–OA…”%CÊ÷%Ç&ZùRàØû"…£oßÚxÖŒ¹ÜmÕ®Á¥ŽÁby-*+d!°‰Â~”¥q@óÀ…ï”’_ùhØÃ÷)jêlìðóØ¤¤@¢É1:þDäæTàs0`ïåæë7ï·wÞonm¾y¿·¹½¿ù&à£&evÅÕ®õ£ÍýÅÄéˆ¢%Š”áÈ"^DNÑ
¢8\ô ‘…¶Î}<¸ŒNah¦sEL§‰1\öP=z_<ú«ví¹-ºr‘­°Ú¦hWºÀœAèæ@ž.‘¨ãÊ™Ø›Xþ3–»‹ˆÞ³—ç0†fUÎ'šÊ³â['"Äâ÷ÒW….è×M7ü÷MEÛÖ¬”Õù€óv/\ÓùÌþxQtûdX¯Öê¡PyF÷RëÖ‰%;fhçh!aë^pM$}}æÁ1¥šÄ">õd\IDÔk¾¥ŽÆFŒ°QW.ðÄáöæ¤òÄÁõœ›GË¶>ã3«$ë’îµýP(·¼„%ñx2îæØF,dØ¦¯.7'à+yª¹ý‹–Ž²•™cUg–7a@s®‰«X\—fVÓ«™(Ös8Nº˜fR}Û•H5,jpÅwMq]YÜ®Víµ¼Hð‚!µÝšRËªÈõH3Õsä!HH95RÛV\œ‘[EW_+õ¸}3Å4«åt„Yr¹þ’ªýt¢´©¸á¬×azQå^‚Bìù‚}øßÿõŸ£/¯b/
3Gï½þy6îvþ<îXt_r³õåp+¬7sOH5öÅC	wìÁ£açM²í	"W†Âƒ´Ð`â®®µ'æp -p‡þ.¹‹;Òtúÿv¢ý½—»[›è"´ý&:ØùîÕÎëÃ7ÜÇdÊÿ4™¦üº·þç1ÿÑÖôúƒêN- f¹n¯‹À)ýà{Ò×°<G&Žº´>qQƒÆHm9p¾18“ OOlÖ¡jh½ù Qœu*šûÙ\ðõúè•Ì+@jÉ4"àjØ¿.ƒ|Ë#´‰Ç2A¹§uFM\Ýå›vŽúæ…l³ævv¯&ŸnÕ° ?ýŠ,”1\=3`N<—!fáTçX ª³Ÿ¼ú"Ã½§Cí¨Ûá¢JÌ¨wõW«õµÜ,Cè~
»­È.dÆ³#4jÒOo¥8ª!f“iôÃžíÉœ×oþ
\óaSôR^*Ó)Ðö2=¡ã9Z0ÈÜõÂ4~ˆ¾2I:$þ8Ç#þØ»8	È³3ÈX_#2R`qóttóËE|ueËëÀÔ·«ÎäKŠj¯­Ê—WÎ‰pýA@ÚÌ<3ã}&Ç¶‹öbQ7#·N6µ²1ã,Ö5´ÔØœŒ´ºS‚²Æ2B¾UŸYúãµ!ñPPÚi•+¥6-å„„>
bÝñY¿ˆ'>Ð
Ö€Æ#ó¦GÀ¨
êéº¦ ÜfW ºP@†r}N`A£,qŠ8Ïu[ÑîDðcÍj1ö«Mlæø{B¡ÉÕ"ºBÅ›äKàsÈ‹4
«¼VA]Ò£’aQob4.÷¿ÎÏbt+$!/mE³12%`QÑ6–#V2¨[¾™Ö¸<ÉÆC^—V]c06„0%Å Ï¹{ð†GòèõK U·sØAóšÙIg²-.hÝ¦!]n“í©‡·i­æÐˆÃî†r+½•Ìˆ%ÀkùîEúãÙ	¶4iw?žÓòiªÙ°ô¾þ®û›•¡þŒ–V­iÓÆÌ-«ë€½á@Ž•Çm’l7+¶ÔRÓòW›SKÚÍ*iDd¤yƒE‘¬y8šc|„½òÿ3ÓÉ‰y8._t2e0ˆ"£{ˆ{âsOå#AƒÀºWjZåøá‡:0w¶÷]#j2þÔÁ®Üî=ßl `ó¥£“ÝhšºÅ$1vÞf’2-Zö­¦1?Ó±ž“'…ò3ÐCÊù6ûäOÛ•ˆxÍ~Ö`$VnÃ|÷4ÐÔíØFÕ¡‚Œ½³|XlÃ&¡7qYCððð’Š‘ü=}žú›ýš=
5{ä6{ÔÐ,w€á.Aãæ['-³2í&Ü§MxrŒ“]Ô:fEÌ h˜¼æÒª·Cõ„ô¼Âb´ZÃdêÙ )fÙ(¾À›ì{6îÚŒú¢†žþlšE
•HÈ-ë˜TÃ Ó)>a›¢õb´²Èö¤Ö«Üx8…Ö¥MmÅ¸=­KÞ…•­ŒTÚìªº`œ¸õ:¬Êµc ›pØ¯ÅšUbC5äØ£•H¡p*õÁ]¼'&pò{¢·üíþ1ÒS¼n–Þ`Añï˜`qFüÜ H,œÕw¼ØçÅ YBöÅÃšÝŠ#˜m?^y:7"·ÌaûqéhÖ]¨*ü hAóŸj ÒêÂ­~Yª)×ŸZðB¡*Ž]?k•u„ó7¡5”—W•Ã §žë‘±g4¿O‘ÿ—ø>Ë½Ö}Ôµ—«½•Ë5ŸŸÞ¢u7®ÎBç˜ˆÿKQ[¿¹ê>ªvÊ[‚ø*x‰G™J¾¢
ùøÙ7O"œìtUT×ç`€¬£ç±ÕŒ¥>òp`0 ;cˆ.KÆÙD"jÏ~­cÍx}¨jœ¯&#ciUŒE÷žØvØj o,II}ÛBÐÃdÑ0Öø…d1ñxOzvÆöZJi9WŒM!#Sóx-ÞÆ·áÕŸ/HÙ-Óp"â^Í8²¡ôšH¾Frè½ª5Þ‰!¾à.ìŒØGWSÇ&ì)•]ˆUØÓÔ|÷Æ.èºD£O‡mãØ]¹†ðæÍ=hŠÒÖÄ=¢6ñÔa&ù°¦é6˜ajPøå!Ý·´ýZ{RöyÖúŽ‹)3ãqè™ÉWZ¬á¯*c´Ê¨>3FY0ùÌäªös¶Ñ=?cš¤Cî­Ž÷5g…Î1+ÊzyðŸúÓ*–ý£<ÿøè‡ýƒôdp—CÝù7Rª×Fk˜ã}•¿/OÓ´ê¾åsTHêI«²uv.°ÎK°«Ž²s¸«=G“x€†´Îæ°ÊE<Éá6]®¹]îïí’lEªSú+BrÛ°'€B	‡È9lrx,©û.}ûø<OøâõLÀWÆßFˆ«ÈîM™ýpÂÂ'áˆéÓ}RbÏ£Z	àyx{¸ðy\øw.*ø™0Üep6¸È5è´€LçæŸ 5°¤žëN½æCüŸ˜ÉPÞ¡Áâ»Zp'ÂxZå#
òd4±É§(„˜«¯¢?†Äú*?»ù’á)…òÁ
Š¨È¾ÅÀ‡’Þ8š{óËY:Äv¿#y2¶Wç{»Ôïv†å3Œ¾'œ“øèWWVþƒ^cg[a‡4®F¸°¼Ãtó¿Æ%Ó ‡cvØSÛfyd)œTS®¹XŒmÅ¼xÜzÏ	ÔX”Kôo©Â-j€”ÇßfÃ”ªtX´ˆ÷°ëó¢ÂÄý‹ayÁü}dwŸ…ãËË9ÑÝ}Þ//g'½*ï²²ëZg›ˆ"´ƒÙ<DÒƒò“¾™·q)"½âÔ‘Lòží|©á¾Šx8;A”	Îü ÑˆÀÄs3¯³ÇIbGÏva¼/sþ®¼ÚÓ¾Ú?yÒw=Q‚ïz¢Ä„ÒÆ„Ùk¸Rè¥uÎÇ0ùòß?tÉXÙÿ8Éd‹yÙ†l•}0ÐÛ€ÇÏ¾®³Ó¢‚ïµût‡dß%ò´s1H‡JÄI×™o
¤;g0gß¾z¹;žL«v½phÈ6È4¦}¦PCY'-7ú2œ ]+Ç‡*N4=¾—Çdg£¡¹P¹§â={±¥RHtE:Yd‹r{
jeãÝ¡Ð<…[^+­óðZ89¿³x8Å	w:Oq‡nÑÅ6 ‘ê	Ë|ë…5Îs#’¤‹Ì{”Å«s$L¿LAq†yÉ2½ó9‘·<³µª³Ax
Á±É=R§ûô@œ"ìu?§IÃ‡³ÊP›ò(?ÆØÅ5(ÊÁ³Á/­Øªvu’ao]fš•·~â¢ˆ/;¦züå¹Â€HŽc+O­"œì…‹È›&5ð~>f°ëº	–½´(íQv(Í—•ÓÓò%ìåuy¶[w?ò|l¢û® È+Àhm x’V¤91Äƒ‰Lí
Ìùº>•©ÔasýìwÃüˆmä‡äÎ@}°üÁoì÷¢ç™
[üûLtî¤PÔõÍZ_šºY{J©9Á1t&ö)w Û€¡‘_aï¾Å#äGIŽ‡é†¢¨6pxÀñ[Z0S
ÑD™s¼À%Û„C³Ä¨xÂüB!ñØoi°‘Q>ôL|ÒðÀ]H¡VO{Ef$ó•Œ—Þ]xýíö‚õ}ìcXœå·ž®<\YYÂ?OŽß-Ÿ,FáÂÿw¼ôóÊÒïE1¤fšÆïÓ!ÞqÇ¶˜°‘QélBÍytt€&e KÉNÑÇôJ C·Œ,X%ûÓù.´À¯bCmÿ:–3ïúC7Xÿü˜z'4Ê~ìèH ËeÐa¯›Nÿ }AïØ%q]±7BîGÓ†¢}Õ®$AÏÐ^KzC]÷±%œKÏïHFM§É>«öœwÅZñÄÔ YÀüJO5‚V¶ã¹ÈYB%fT­>*«æCJêk”§èˆåÿ•U§XÃNë«¦À^Âß²a¿{jÃc!Çp)
[%Ç~x,âƒò¼â‡¯(¶ûÅþ4¿‡1âÚ¡“xPqÌxÛ‘ò—<¾0–|šÈçÆcã‡^Â÷ZÄ·úË*«¦C«y
m2µDÔá!ŸygñqF_)ÂzÇ4}uœ#¥9É˜¸”÷ÙÐñ7þÎ/`ó@/+‘‹“ôŠ¢ÀMÙÒì‚þb(Ä“œ”8¨SªV“4®¦(²»#âbªÿMMÒœI1A°b‚?¿¤Áñ™|§lülÑ’{Lg œÈé—R`}zDÓ4ãìO´K­'­˜Ó™Leâ6ã'ðûOS	]QD_Jí±^K†` y*˜ÃÆÉbŽ˜ÆŠs ~baßÐË\vÉQæú U†,„¦¥ø6M#„D½$–5Î¸X*ûã¿5)•P“=íPþ†ÿæ¯„*3†ß=Pvä…'Ü]&? DæEO’|ÕŽÁ_R&¨ŠöÕAjíE€õ¢;ü)~¤°‚¬2ü)ó˜þí(ckÖ7˜Iÿ¬­þ4›< †€¨0ú0-D¯€KG„9ZVÁ4Í#lÝQ˜—ýi’½Þ«ö˜Ñâ„Mó"L9!ÜÑ3 r’´·³“Lö…Ú©Àšø€-]i¼§´iÓ
	pÉ¿Ç%ùÞ7E}†ìÎºVÄçÇŠöŒ‡%§LHiôZ€k8OŠÓ%à­ìR©4¸Fƒ`(RSÃœSDùñiÊ¯'†«JÇiš &ØíO»Ž+ûÅgÜŽÀHþ‰l_–x§ñÈ$ðZµÉô¤‚<žÑzÖ0¤‚AnqŒnGÀ0YcÄBXF:&u±$Œ÷</’:Ö“¬ËkžõÄOîu6¹bqvHrec‹!Ùµ±0®ìã¹Æ•)¾‹³‘ç=Ë{·çÔ˜7ä˜Ë·òé;îØVÇ´à¸óÝäU<Y'M_‘<Î\$Ä{š×¤×€B*ñR"$sn£¤ã¸\»Ûu¦#E£‘d‰]¬b!Bc-sdt…ŒÅ5âeÈá8;å$‚Â÷@„•È%o4¯ïW+¡=½PôdñCJ·òísWŒ•@ñ2lME%Ïu1*PŸ]«+fI~sü»&n[”€ÝÓnxê=Ê’’å(¼žlU!c9s˜’[!ÿ.ºÕ™2º±)6ìz¦rÀ(»îÿ75>8Ôð¯Fç~¥ +iN[ÓöÄ[±ŠïléBÛ×NÙ kjýÓ¸œ¢ô;NÏ‡—ü†y³²rq‹l&	ÑKûÄèÒD‰ž^£Ê€æ MUv`ÏÏ;16
^£Ÿ%~R*…2§e ÑBU(€ˆ‚Œ2ŠÍzãmÁ¡×	¤¦Œš rH‘ *®Ÿ…GaD%¼Çï¹kªDÆ^]œÖÙ#ëæ>üèÐ#º]D©ÊXôM.r?|"öéÚbô¤ç¡cHÚq–J3j\ò'@³|÷öE[¯SïWN—X<^±j37IÉ¬°¨¶¿‡ùÚˆ:Â©ƒYâ5Êu¯¤°½¾ùW¼¿j5Už@_ƒøØ;œ¢	˜à®ÙæBGKÿÕWFÂªf©ï¾Žâpç©OP“S²h‚(ý‡Ù	†_:eÆŒ¯k“xˆÕ¬s·@ÌÊ>¶‚ZSü!ô¦ð/ê»=‡2“jº·ï\,âs¯O³äQ54˜¢á”õ…ö6Øõh˜+îæùöÙ(Óò6aùW+7¤6s7‹ŒëùiD•(÷4@âÊÏažDÊ‡êÁ°Ç0“át4F—¸E­I´T°Ú ¸]¯Ñ*1Š/€Kíw|‹‡‚E>žB0½ÝËúL@S=ÐMÆf¥…pË‹—ê•üxT.wÒ/.§™àòŽ’·W/Z°ï9Ùq4$ È>å,9rM7Ñ¨ÕÏJúË5¯P¨ô45‹“ùÈNC£;€-1rHÊ9ý-,)XB¬fÉüV©Vù=µháKyŸF”æbü„Õ¡œÃ"¼Èå6­¡¶{p7{ìÓ>w1"6‹[…o cô»¤Ëã$.¸„(¶¹µ‰Õàìt®øá¸žd„/À YlCfH(è«aŠ	;‘½&ò'PÕ¬ñŽðï‘‰á×›Ÿ‡§fxl_ø	 «’Z„Rn44:[‹òU´úž?³6<«ÙkdŒ)µW´Úq¬I'òW?=ÃXÁd,xÁlš™LüVæxè1<gæ ïFôìpß½kXIGprYØ¨À×QÝA>]D¾¸†ŒÀ8õmLèòíé;f$aÐxËšqL%Ö^µ°öŒs€7€%Êcá:õ€éœféXºo»„iŸxç MÏ._{šfº`:?<…¤~ÛQYÖ
Oš•Û¦.…¿“øÁõeR™ŠBq«øÚŠÛw®1ý4.·Ó„ò>±±ÐÁ8­C‰~ŠIèí"9­™˜Årí1g—HU>‰™ÔÆÜîZ<ÉÃƒ‡$žËìŒtZW
å€…q¦-€tbQÖ¼'å¬:¼5åÅ6k™“ªd$‹²« WÍë°•·(3å9pÁLêpqÄ©¦Y´BƒC9Ä©'ÍRZPˆ«pÌz–z×È0«’J3Òò„ðRIåyp	i­ðÓZse À®„vW ^”?ñŒÕzã—K’A¼{Í—ªËÖLuI¿=Ïî¤¼ìPhéÄj«ù“ž¦«¼Ñ•gÎAµ¨eŒÕ©Õ¤ÆÃÏ|Uy¢E[VÆ]ì¨óðc{š¹\rˆ…°BµóÒWy)O“Î
?ÞÉ7é­î	I>ø4Uj<\Ýòy½ ¾Š ¬³Á»ie½z+~TÔé©ðS««¢~…¾J‘/·P³Ö
?ŽæJ&·h[Ý~®çÕa)H;KBP‚·]¸A›U/D®ù9!Îhy8!þfvNHTt5	Ã¸¬ö4ãQÇ_ÑÁ¶/Ý«³ËYi
žÀ&!å+…wqeˆú¼nÃ ‰6ÄÄcIK—[Z§y3=ˆ÷ìuäd'T<Àhº=Hxôl4ðö ŠË|²Xå\òžÂ¡|•k¹e¬ªŽ8Â—C}Ÿ+õ=ÐtÆ'nD^¥ÃS
½Ž¢/¯<ìáõ'ñéh!†Œ¶ªEH[áÒ
º§ûz::J‹.baÝj/t<üÆKù ‚«I§¸Ï>ßåxÀPÓbøÕB.Fdåö”gG&”÷¼H÷I\“Ë¯èA’‘mâ«]uÇ‹§RÀyÎ=¾×páÿÄCðnh:~6Ö¼1Î›ÜnÅ½É{nŠ{KG÷Ç½¥#Î½±~ƒÜ›œ€‡ÉÒ‚m³sË¨‹ÖÁÕþ|ô{k§®[k«ÛÄ²ËÖœœÝ–Ñ›÷¦£»bßÔÎZ8W|fáÈú82u©ÎaÉ¤lFŽLSŸéÕ~=†Œë¦D³¯Ý7;Æu¢C!ìyÁÞãŸ=˜°î…X0-v¼]W˜	UÑkÓ©Ãí‡·ÒÄµc4•ÿ$¿=ê4È]0Ã3¶’CyæýÓZÇ´ò­Ð‚g=›I\ŒâU?[ŽµˆÏ¹u~
Ó "7ëlù`æ¯pýyv Ò’yV*¡Ùe
€˜ÔÒË¤–MLª÷‚¿Å¯ò-â˜õ;¬¿ÚXnÃ:[÷„çÀ9C‹mgiO”|3<¹/¶ F\³"œ~¦YŒýñÌêP.¬k€j5 o5ƒPHÿB-íŒKBxuäydø	ÍU¤@SÞYûÂ]Èæì¬;m5ccÈæˆ›
µã¦0C6æòH³AÐÍì­OÛÍ­ék:°ÐŸHR2Èœ”Øær6Ÿç ^L’ßÜèð 6ýiœg()±ø/=ÍŽX‚g8á=9¥`rÉÒq6B­%Md)‰a¡Å¸ùâ³ØEœçú…Çé½I$UŽ¬ÅÓpŽÅËŽhÅÄÝ¿ÜÆ)»¯ßV«¶Ú•U+z¨P^GqñÞmÞÂŸ Êp–¶¥ÿ„ªv”;¸µ8ÅvoíÅ^ßÖbïv¦„Í[‰Í´0·’œósŒ4«ÔŒWì'ÙŒr3ëkvAu+Í*ßVºEÚÇ’šêT)0e@d<ŸØJ†ˆ±iHY.ócR¡€Üj.,Fy)Õ•­kÿ|vïõPÃ¡~Á²/á“„ä>3z|mú¯X$4P®M_sÃ«jRîdxÑÏmÞ³ .BOA¿„ÿõœsÄ
)1 ¦”°#úo@ˆCjpÝ×¼¿¶G²a§1*w²Y¹íÐÌZ5Íë§s1‚bËž>P<ú›¿² ÷IÉÌ¿¸Y¯?Ì6Bh¯veTëè¿0èçÖm^öÁÉ:Õ¸7…õN+‹:%\ôl ™0\¬†ŠÿÀaQÐÄcËoåŒ¼Û=lÓö4çôok\æü…f_´Nêˆ¶Î€‘-éímŒ·0tÞßH<"m0Ê·ÊŸ§…?~Sšo&JVÄT¹°ýE•æÅÙÑM:ª~Û%«Xk¿Ä7±»zæ¯4‹ ŽûF„€Jƒ¥’ªúÒ±Yõk”ƒ6ÇVCe|U?"•ALÃF{è˜HÜÆBP¨6ºå€F×¢wPêÎ®bþCº5ÑúŒ”Ãøñª´Â+¬f%>½âzcSñ»«ö˜º
h¥f`6(›¼W‘,U³¨}ñS¯Oæú„Yè“¹ža&2~j´Ê\¡0;ëÔÉL³1›?¿âiÕ¤/¦šáÃª…!Ë=«B¿ô4Rò+^o†QnÖè ðÃÈtf²œ?—›‹KÞNß­ÒQy˜c‹ë˜n+
9è)s
×~RZiJÏÄ.ºZSê÷û¶;º)ž5m„Säí#w+’V{sµU¿î9H#ºÓ™>_‚OÉÖbÀóî´›ŠzM½•Ô›ÑÙ;Ÿ>Ûº­o®¼:œE_=­8—®Å’·Éô+LÜ´(X›]±½œ-b_Ô•][S¿c¤BÔ¿SßáÌrõ–i%“8a•>@“#0öÄÌ³Û¤¸kJlm7-ó‹‰áÙ*Z
¤Í³Tàî€ì•x«§à0kû‘§D²ºea¯ÙAÏÛIÒ¨ÌÆä/û™ÝwÅœ”“d=Z øBþhv|bFÐ	Q‰ÍÉdxi;BÐaöR DY´ŠxC¿`ÈÐ¾ADìÊµI¤Åh¦;.?Kcº¤Eï‘©+ÒDÝÙéôžJ¸_é‹‘ðE!™7ÌG¹xúFSÃ,‚¡²Ix`O‰øØ1½ODž:üÒëYæ5Ï$1=Íž¾×ÉÚ<xóºÏˆfv|)2MöBsVaoYAc•ÀÙÎúè¤o>¤gÔ,/4ëi£­[!G+2·%R'Á§Z"ã´˜Ç
)á$¼@¢ÓÙÖGª³<¥XWÓ»q)V§¼Ýêhç½-ŽN'Y‘ŸÁ€_ÂÊ@Žžæ²ÒóˆØÀl€]ÃÌÔ ‹,,í®a*yjàˆÕpG(£kY‚ÊÒ|P_rµÁðÝ·|e˜!­—wÖ%Ìv{d¶èC8K²±Ä,j-³=¤#…æx Úd¸$ÓAqñ²o˜Áö	Ñbó™ß´(W¯Qì§±+™Î]è—³e²>Éõlp‹N~y”PŽ]µ|zL6“EÞ2°N^¶óJô¸hÊP$BCÚ§­Öå™¸|£jiz³’€äè½ºEÜši€(yW"p„A…ð¦FÎ,£åˆGÜµ	ƒxÜD„í²gPÔÉÃƒÃÝ½7ïw^ííïlªÈReÏ¢n_­É…è¾5ÅpûòÞn~WÂá™š å}ÐŽ²5í€‘5k4•azòI¨OrÄCÍe‘5ÞD?ÊÏ‰~è£¿G"¾YÄ€)Å,_-PÁU”qýÒºz,TN f¦=ÒÞ‰J×‹Fƒœµ´.FƒVÊ&£AÁz™íÙ*1­9;¿“ÑšÈíÄþ* i*#Ï/š^3MÏýÛ1‹ò“‚	ƒéÖŒtÃµ:KÀR"cÒ$>)&7s¥ÿx˜”ÝßmíßüÃÝ­7¨C€ÇQõOÓì,(3ZXìgxÇã¡ñLeøg“Eyûfz|‹Ì§'`ƒí1˜bº×Íq‚:š¢Âì6<ÅöZ¶Ux/é€@)Ë
ñM|5Âº“w	o@ ›ÆPÞØªJvÇ¿åÖ—OWVrŒ±‰}FBl`LlvãÁ`1*hdð€Xô©æ"žŽË³Tì®EÿµéÖ:ÕfžðpuEF¥´ö<Ó:ðˆÂ›»bj/Ó³tèômø1¨ùAo²~åqú S%ë¸ö­ÏÂ9Éè¸(„Ëi‘#8±}ÇŽÍ`%›»ëFëzÑ›à‘fÏbo£„¾ûµÝwÀýw5lëj#6·ªV¨4SÉ£ö5Á_‰=	vð}Ã Àz”N!muälEYKÆ©Á#E/ëÖÎ>`g€ì] K6µƒõ™îDŒ¦ŠTl$öp6®Íuoª!õ¿Š‚Öæ’GÈiG¡ç5ÆÌí»›DGyw×,ÃÝP-R€’èÉæíïL!ªa”¢<K†²w0ùÉNRHšüX¶Un ’g~KÌ½/3œé­TO¥@ž…QõòÍx½/ÈÒóõè(ÏyŠ;6“Ú>ƒ5i+»rE~fÐ<Ns?=îsÒã®Ãy\Œ».L~¼ù%Â¤Ò¿à¥KJCŠéÅèh˜ÿ4M‘sHf\]°^IÜW)éôA1|¦ˆ´_Z1PVêGCà|À¸ÍîäñÕ
€b>ìŠGÀ®®­¯¬À¼Ê)KôL’Ï±wbûk¸~ƒ
HÆŸZœ¿hÝè–¯j°…äÊR’]1S‹lö¢‹H¼¢¡–iuÅ_±ª<ZŠ¾¼‚&7H&¢r¯sà9¶†˜g)í ‡¥
tuÍÜ1Š’ðèòZŠšÉ‘«Gh¸ÑZY³æÏMHeÅ—Q³¼ÕZwU1Ú©FóôD@ƒÇ´ÄTêÃxœþ(R¤É÷¦éÙÞÏªÂŽtÝ\³ŒõRF	b=÷ 3ÉHÎ€¬g* )P`qÝ…M*êbR—áwû8§òuÏX>jU6z˜ïŽ‚•ÎÂEz¨b-t€^~‹ˆ?ÈEzÌ3?:V`«nt3ØMjS¥uþˆ!g1¥¿¯¢Y'hZ”ÂËá]‚?Ò‘Î‘–_‚¡0Õ-îDú§¨!Å£žr5²ÖÑ«É>•£·¶fÁ½Îª)¤Ô¥eOìBzú‚ëvéj¢wžy¼°øíDíðë!ínØÛ
Ô{â:kQn·j‚m8bóí]Ø¢=cý´æHÿ~BUq½ì7ÖýQyOÞB×*ðêêÕ®öÂ¹á»Ì\{ûÐÉüz³ñŽò`kN?ê»þº¡P³<DÌ‚
Lá‰	q™5ªDú>²)„zïyt~bÐÿÊ=õnnG—`àM"OÔücÏ¥Xœ½s"Y#inâH$ùVì>º"rzÄ/*#çìTL….
â8±¤LdØˆuGã˜výœƒÒnIG£UÍÉÑ"~ŽÙŠ*7	—’®×öù`–B‡9`œ(¢™0wƒjÂãõhU¯Ãœ$cË;2?‚µ;ÃXåëV¤XVôgb5í²í<Þt}›Ÿ`jd±”Í²Zz†ƒgQCmÌ[äënì2­*Kã×(Û®²Ô£Ë,ÓeÚ"ÍíÐ&„8u¨ãGžÈUÊëäËX?×'¬Äù©U \LŸfüZãsßLªl„æ˜IïLbø!e£Q
«]¥ðº›ŽË)	Çc)ZR~®R°·!•Ð[%°8îJKD§¹@¥žÊ` ­Ê¢š„fü{Gö å øÐ#Ã]ÓuqŠJ˜
Ý JLMˆçK2T®ÚÀêu;\8ÔTçŸ!Ë&Bæîø†¨kÛu%‹_Ù®@8«V1uË)&µ–CR0aÔR„F°#þ´ës†ùfŠ4ý9-1…åi<>/ã¨³Udx^ð;(lZÕå:'#HÈ–Kž-)“WÀeŸ[Š¹ìãìX»@£˜Ë×ªAØå°h)ò²OÁW•ô‹¿ìcÛõ˜–Õ6CZ*VµeåbÙA|N²ÒvÎ²ƒÓéø#”ty§Yj²Þi–¼b$Ë÷îÛR•Ì¶9ÞˆˆŽãÃàÌÈ¹B7.}LÓ	’yvòÉes¯«Q&þE%ÕãÔR(Œƒ6<¯ò{À‘¤‚¶Oñ„ÞhQÁJókËÌt‡nãr{Ðüq|–‹Q…Ê6Ø#g´+Óþ‚0ôûõÛŽ2Ô$¼uZU*éWƒ³3˜ûn(¸­eŠ%^(âN’U¦ÊYSIªöü›oÔ^M5KïuJMeN2FÎÑ{k½YÝñåÚ8ìaR¼Ìa\`ºeJ…¥§H¢pvîÑ>£ÁÎÓ‚NuJ®\¸Ü“ùÍ3ÎOoóÂcœknžv5m¨Âåév&6UQEÈ·#xmòg}&#Ž„ÓÅ†–ƒáoµ OŠ9žÍéHÍ	¨iÈË›ÀÓPÝ{ç«ÍŠ58êà‰áy|Yš‚¦ƒ¦—ø L†\ÊéxÐŸŒ8/  pÿ
§%ÃÃ´ZÇFæ»/„m	Ó²[ËòYl>¾¹l„°»	.;ŠRûe˜ÌÕÀ[£ÆŽ~álf˜©­£ö”®Ÿ£‡ÑŸ&Ã<NÔi‹~:ë¾Y$'Óu¥EéDÿ%êÀš¥Ç˜™/“&iû€&Wã8ÖÑP}+'—”)z„1™Æóÿy¸²ôxåß@3X/ŠjY„¹x$ãOÅP:¹x³»#àpómöšf´=^[YŒVú¿6%ãŠLx>ÊYZÇµ+Ë}Và¯Bì±Ln+`¨Ò3L<Èpëb¤Úã×\…ÄÜ7´ÕŠ^­0|H/QzÊÝ±Tø«rÇRÝn,•‡¹B3Ñ&&$SÂP8Ô¾‰ôûžÈe¿ÄBð:ƒ^ 3wXg)‘gQDöl.MelÓbùð$DºrÄvà ?àB…,0ÖR(Ï•Ýà8ÙŽ+ßSNÉO?VEÝêéª—ª&Må«ßDLMg0ýÔ…¨¼h ¬1pñ)/jIã>ÝtX«Èãì$Ñ>Å‘J]gŸ+uï§£9mG :ÛÀ}a¸?Š8cä§;	•k¢<¿6•©ß¸4…o³tÈÞ°}²÷Ln>ý ïDwè¢¢AyyæM{(3
SucSžãá´<å.àû,€…âôëAðÝN`hƒÕ¹…ß©/hâë|ZÕW‘œiM!#øÜµÊÄ3sù•úeö³L¹nhS”ZQ†z0ü€4]HCg”*[fAo(M@‘Ö –yçø8 ÏÓ@A"íj$.§ë—ÑD›ÃáKÃÓÕæ¡5½ÖgÁ1ô¥¡7ÌO0Ls6–à$S:³™GÁtH0I•OÏì¶R.·U,›Jå¬møª ’£#QTkíÉ®,³]Ó‰?J	Í«‡.
YtJÙµ©]ñÓ±†hO*¨8f ¥RÌãß*Àl±ô~â	–Öhx±lÑÔñµ+Å“P¼Ø[ÓnÍÆ%PüLqwÂöÌçÂ~úìD¼~µˆ?ò´²|<oíãÀøUo*”ÜK©mGc8C¼(`SÂNMRe5©£Á}ËÙžGuì‡²vu;H+a¸„w/ã’Þ'H[»Ê××¹ÃŽfÍ‘¥›iNà@ð€=ÈFÄ¾Æ„ôøe/.à¤Ò÷wäµ\»eˆ¾#òi[Õ6 ÊxˆªòjÝëôó|m1â­?_]Œ°Çç+ŸäÙ˜ÒØŸáuä•hpÑðuò™ßF´
ÏWtm=mtõb-ý#Ó8)˜¾ÝBtƒóD´9ãžÞ4Wôîd'z®7¸!Xù\à°¢0ñP§ð]<6öº¾±ý.2ª³>/`;Î k,E¹û¬ÝßjŸëû[µ¤ö¯*É6«¸'pÿP0ækœÌ»ÂÎÈ¥°Tpìí§ñ â~\cÚ]•h¦ÇëdãÁ%œCcÔ&ÁÀû”Ñqš&hÒï‰8›”¥ašôžø8åâñIÊ"Ú$œ…Š~š¦Ó´WBvS´®Ù°H
4ÓÑ†ð°Xg¿ìµBÎ-ê>^¨ôçeãl„ŒX’•—Çiuž£Ýå7èÀ0ÌÏdàU\Ñ|læµä£jXM”?y“¦Ï„—Q–üùbô˜s®XÏÇk[„ßÏÐnêWmÐxË®iÀÏº“U¢{…Fù°
”SÌ¸¨¸K¹ªÔ/'C8Í:KÊ×…¢AÇE	[Š­Óy¡[I¨Á "Â>Ô;U“{¯ÓqCOwÇ«üvåÝb´ºG°õ|•=–¢UçÝš¨³
'êVÃŠÿL`yM™Yù:~ÝUÕû'%¬«ñåi~~˜ÃÁÚ] w™l|vóË0Ã»?‹Ñ‰d /‹âäõéƒõøÉ¨ÑîD@¨xV6	RØÊ^i·ŒËƒ¶ÉŠû²…kRÿc×.g$[æ¶ëž©ñœ&£qxœ›®>{¡1bæí†yDtcüÌ«"Z“à÷„3%Îm +lÜÜ?±nîW©žëCõ}2Ãh™.úÓ/^Xç`±	&Þfìš4Ç+'äƒ‰#¬é&ÜÄ¾kª|ÍÈA†oS‚ªf`Zµƒ1‘üš‡1³å®]'‚l¸><ÝÆe"ÄC‰†´M²uF–§S8eX¶Yïmuà,—mf-èuéˆÃm³Òoÿõ«|÷à÷~Wª²kÂèš-8·lÐÇG‚²&0àµÕZgK#<‚µÀ«`¼©1Z=É¦(FgÀ8‰éF.û ‰}Éšä¾µ²ŽPåˆfõhH`›êîAd+ÿÄý½íoQÖÓ˜-‡'2;È2$©ôAVl¬ˆt_&ÛG¹m‡çª™ðëì`ð£üÇFÜ5¤-q³õ9ê¤Fq%±ŒðÚw„X`•$Ë¯^-_Â'úþûõÑhAOÐ3Ò¤oóqu ìlwMðüÅazQmåCX×ÇÀš<ÆÐ¾+{O‘$ìŒP˜I¢'++²&êvöSnþZdùROùÑ¡ú‹Z¯`‰³‡;t’+:r“xG¶ºÙêŠ>hÂ‡mr_ÿòŠïúëóácm5} ³oC§´]Î«Qt¦ ˜:5Ò›¢ûžª­˜‚`yÕªÛAºïT;=1ÅGkrKHÄt4Š‹K§âqyžL*b
™ç:Âr‡z.¸RÂ—zT^YÏ•®Ä}G
©«¼‚S‡«(œÑÔÄìàÑ>°—•žb“Y«£øâ6mb.Âc€ØƒHO›/¶Õê«Ðíj³[ŽºrTz|<¾Í.Ò¤»J	‡W:¹ú(„÷¬©u‡Éi‘¦ÑHúõà8òxE“vw<¹)ù¡œªÏœ?âQ½’¼ú ²<È"ÎZ¥0Yó‚°ZÐÅ¨P‡fíUYây›VØ—dƒ7dßY] (¹²Æ{GR±ÿ×r^e^€H/FG„TÝXUêôØÕ’}tài÷HWKîžØ«Û‘1så’­ÈË(…~‚¡þkâÿN.°PõAM~÷A¬¨
£ýV0öl¥¾ÉQB»}]uOM {;{tó•”wi˜K"—uF)èšKA;¯á§#‰ Ñ ¸ùW¢Ý™¦¾€ÌcÉËb5úNâh'ÕùN˜bQýçu¨¨XÌÓ4aáíÛo$¿"¥~d¿¿‘  
øîjâ(O.×ÕºhA
NS4ìvNŠ,é˜]T—CÀSô€Qhó;%8‡ÀÞÆ9)âQ\{ d±ÐÑ§Jx=q×_7nj[’Ó”qU+jáŒdßå®V<¡1/þV>Œª¦eÆhÈÿšçµ€Çˆ™·Pa'¶ââ$þ$+LmO%dëF™	®À}iÄø8Iö@PÓí3AÊ¡½tñ’“ôN@fm“…{éš,ÜLs:ž3‹{|­­ô47ÖàîQûgí±¹ƒ6+t,‰“wË6­
Û,ðS¦ó‘»j\Þüà¦‡™<Œ-%6•	LÃ‡Þ¿·fÞ]¦<Çþ¥?t.9F7I„m>­&S€îÑ0?šÐ@\¿vTìByJwü¼"‹¼Ÿ¼âvåîY<\ÕÚ)ï–ÜŽìZk550	8Ø"…+ÅšòLV±{ôb(üÅ´Ë(òLÅ’Þ8Dê±zMÔˆkãÃs$ÓAÆ.à‹Œ€]ýp0ð€0"3µˆÆW XDrÎ²>ª’{ä˜î@¶
ý¡5cjË¨vüeë_I•-kBá/$Š¶
áÑf
>þ‹§Öò½£ˆ9ÔÙá_'³bÜïòndÉâ/¼JG¹e)áÚÃqöÓ4%ƒAocqgC$2±i æ-Ètèj©È€Ñ‚„˜NmlgBÔ÷ÙóçPH1©nŒ×éºá}bÄ†¤i*˜oe+ÒˆÕÃéTÚÔªÆ©5•úZäŽï ‚ÊAä§äöÊqxt–J–Ùˆ¢£šÖªˆä“­«Èdpá&Jkþ\Úœ{¶<Û"fl¬–¾Ùï ¦LÇe†gp'gç ÛaÖ‚ÈbôÖÁßEµïUX¦E/ö)z°¡½ÌŽ€º¤ðóÛ‹tiUJbìkIm‹˜G·¥ðœÆóZ=Ë”åËp$Ñ„IžTMIÅe(^È³I6qÈ œÓ«h©Vtû "Ë|CN|’K{H”é[ Ÿ±¡÷Ôu±äî½ecMx'öÅcŽ­†8Uàº ã)œ¾Ú\$â«XX«EYËP»Å,õel˜yí@Â°ŸïÐ¾ýòÒJ•ÚEJ™^Æ@JAƒHi¨{¼Ì^Wï²èv¦XÉoH<?îÕEësÁvjíoh&O øõ7Á,£›×>PI|tm’8äƒ‰h5%Ê–÷‡¯åìÈZ·5Kck–þ­©Ú*9â—·G¨Ájž˜_ÎŽù^0(ï‚öå<Ð¾ýÐfÃy%‘¬Œº–@‚r+´¹kS°—Äš¹„6)ö\fraQ¥™ÖyÜz†v1R[»fIÝéHV˜´ËÇý,ÉÎ¢Á0.K”iŸ/”“x.].=Yx!ac—9¦ì†À³—ò@â—KGiuž¦c­.«ýÂXÕg§«zs¨÷[zx1¡y\-ãÁÇˆž•èñ¹ôõÊJTð¸è¥*;9­^|“³¼J1×	ž¦ƒŒîF?[>]µz›8±†CÃÔ%z¡NG/¾CËÅÿ$5£2pÀw®P/Ÿ-OŒ‰-[3óÂé$ž,­Eøué¼ˆ'Öà²ñdZÙW>ñZTÏ†é‚ý*Kž/PÎ‡%–¨Â) õš%	,†]"ÒIõ|¡1,/#ü³`ÉÇ[äñúüŠyW²¼”.ÂÜ ËÖt¾™VUŽª6Œ›}€\<‘g[fƒÏ¯’ü|ŒÓÓÑ—á:ª²j£ý&Î.0({ž¤CžšÂ´¿ÍkG¸ÇŸ_­~}Cˆ^Qs$Ël(·ËTä¥=À„,]{ßÛŽ’×ßÆ:ó+“Y“|0EÜG'·@ZøúÍånÒíèxÒémôX«Û»vÆËC£eKã¥•h3ÞØoPnWšzC3cÓI,Xvó¹äat3Ðü®Õ²‚¦ç¤rì)d“@Œtéw@AŽNØ÷Ç+ËáW^$iÁž¬ÁÛÓü,-ÖE™UõD+÷8R~äã¥ªûKï ¬…÷»v:v+ss8T$³ã9†yÅš7êsmŸÙøá{e8ÍŠ¨Êq“Àúy&è(ÐÔ’ÙSþ--a?EG9ŸŽâ¨¼ùkt–þlÓ¥(²q?Ï‹¸<]øû„ð÷e†§któ‰âHÍÞÅ/6GQÏžæUXç®nÃÞ±wµÃž¬˜8µêàÔã‡ Ì„{JqsW´ðÍ×YF]žÈÞ} 
,ŒÆ[ÌœA×ÌZAÝZ÷*ò˜
èZ¸Pƒ7ªå‹˜jF¸¡ñ	Ôxû.~ÞNKÎ”rALd&?[vŸ¡{PXÇâÝmðkn„š­ ²‡k¶Z¢ð"5mA$ý@«Ù3nŠÐ¶` ¨…€(Ä7†nóìt”NžÿÒ¯ ñŠ«Îîî	ã²J2LÂévîy¼[)–Ë3óê´ÝNRUó7¿›Tæè_k3H.ã×ÙMÌw§¼L«½¤Nîô¾H90ü “ˆ"ò÷åô$ÅÌV¥h%þ9ùñ¤¹Yo>N2ÙB^ÊÒÓa’Ë{¨†ú¼,õRB†Í“³þm¿Ïl£K‡Ÿ{Ûé,õph›ÏÈß¼NÉè·;&/Ÿ¦ÆÁQ¬vo¬_ ¸ä«87‚ œœ¸ñø É‡?-¸òõ{(½NOnþ:È|(Y9[ŒqÕVæþ³^ûÇ#ñ…iÔPâ-=‰ð(9Æ»ÇKègê<’¸µ~NAn—®n\6g£†u5Ûø±0¹€¡L.—rÍb>L¸bq¤Ü˜ÛJÌgi-:?Íª”©AÇ9©ììpµrào*ewRè¨¥Nµ€Í¢õÀãJ×z>’'²ö«c1Àa®½O²z¶|ä`É­–CòÉŸçrhÉg³¾~NË¡¬{Ÿãjèíç²^¾pN«ÁŽõÏr)$Çñ¹¬»pv;Bfeÿ,×Aï3Zd2Ê —á.‹Í$èºÑ®ýÃ?Dömö–¦ž84pñ*KìÆÒ$§ÃŸ•ò!±¶MÍbXÐÅ>ÂPµ^.­F£d]ý\#Å#×ˆã˜-éé0>J‡zëGÃ\Ø+/JqäÚL' 6â2Å¯.¼@éïOò‚mƒgËÔ´§Kfæm0Jœ/O/…!
¯E”à&M–.²º¬Úq>˜–ë\S°D¹=Ù#tXD³ú!ññÊÊò=í]ý<šS†SØ_¥#ÊùÌÊÐØMÅNö)ÜÄÂ:,mÀ,Op¿ò-,¼@[@‰©Ýy¶Ì
øê^¹¾¢äK×x`˜¢íéåó+|x-æM?^°?²ÏˆŸ-39xg³çôðÓ!£6¸ôþ÷Š–ŽÏÄ¬Øé40?$e–3"Èu8jU%G!tb€?Qw+MGQŒ†¸¸ìÕ¢ºuå€ F,À¤~AUÇ~,*±Ÿ~¼`ê:íÝmG4É°½kÿÑTo¶›õàžh?zÏ‘«IÀÀ%n’ëé—X¢uÛq@$'*å÷£J÷zuÍûý¼êü¢<—­ÔçvÎÑwt‘Si—ê=D›=ÌšZ
ƒÅnûúNPOZA›¿0Ç'åKE[©hùð+mè²íc¦€C“žo·Ó† MŽÞÎ×ÆIÃIm4žYõµ UGóBy'ù9²ÒŒMå¿G‰ÍZ'Ó‚‚SÒY€yú& ŒW‰	¾P÷DÞ$°çj½ÎÔ¼¤3ìÑF—Dí¦ñœtYÄš¶”ÇR…Õ@U<Eîêµá¤|ˆ'å*fG@SálcÇ7ýÒÁÛÕ•ÉÅ;íˆk@ç­:à¥çÜy†ËçqVGƒôŠhŒï0¤á.ãxCÕ¯LZ@×9ôÔ÷;é´ƒJùñR¿ß7Lì‰…€ÓÏ‘Ç*´ã%aW”ƒ8”U—K_¯à%~­°ó
6ÐÛrâÅªépÄÃžêãÒJY}’¸þii*Ã©_ÇÌfMf=Š95i*j¹pÔŽüþÞè‹õñ˜ü°®-*F€¾²¯»ZèèµõVd‰cç	Vzû®¦xØZ‡Ÿö8j+¼qô­¾´Ú¬ŒÑê ›yxR£÷0´(¶O]4˜e^,Qâ´ðI	DÜ™°2V>J¡òaÊñ³-aÀaÔy|m{Ìj…=º/ãý½í>f²¿ý÷™a‡ô}¸;jXþk·ÀÓqáö˜<?Â/N§ŠOMð53å­-¼¸òÒ@îN¶a­„ª.ÄyÜó/J}$–›¾ïç·v5uøµúÐM‰“Ûk~d[®ý ´µŸâsÅ©*ã|à›î®àxÚ°Ñ¨BýIµ´È/ôÀ‹ 0ë1ÛgHª|“êã“k‹[ClG +^Leõ×Ò±(²ôÚÀÍ/?Qût”x¶ SÃŸ=qGl£•kgÙ=zpEÑÿAs¾ãëÓeQ§Ùï6ø›Áì´¥á¶
ƒßê#ü™hÙÐ%Ü\Ó”~mm¾}9X_8öÒÄUhJŸ$ºÅõ	Ÿw-~Ñí«£:a' üx§ï!û¾ý¨oG[Ó8‹Kô}(»~n_mMÓñ§i‰œNm„³ÃS¼PË'BvøwgiÐÚBlÞÔcöüná‹OhcËèŠ^ªèÑ(Ú{VÌJ,çŠWþÛÞ|³»ƒj÷öÊ]Ÿzw~|KÝ’Žß¶ó¥œËép;x¨Ð5ÄÀ8*‚§ˆKZèË,k†âAj¤­—ìUÏÎÛLÏjD„Àùc<ýË§E°²Ê;´TW¸»àâÑ„Œ¥Ÿ>X×¾õÏß¼–_L"(?ÏQÛlOŽ·I\©¹)»¼'n®,©U§ó^?3SC`*µö9—ZE;ûÔí²Û®sê™1Î?v«x¶­Í|NŸ™Ñƒƒi–9‚ecPjè'ÄÖb.À´TˆJI)z±Ì.HCÊŽ†ùq%äP>I«xxŠ×¨s)þ4UŸAQÌá9gÖÄüˆ Y!ÞÐü„¯4÷t{³“]Ÿ©¿qÐ-kÍb‰¿E*¶´Áp$>* |û”ø(;Å„meªÓdŽ%Voµ’­1ë5ÓÆ€%>õ†,ñ™Íd!Ûn(QcÚÀO¾R»©‘Àî\ ÇÃ—Ùø#×‰~]cYàµÉršðjÜròC\x±Gƒ¯X¾Üë¦5Œç3³›Óš¬wø™‘0ûˆYÓõm»n;~^//GßS  h’W˜é F»ÛŒeDÂ)G;’‰X8ÿ…\õÁÔˆpÐÄÃ]ôúy¤2M#–- *‡ÒãÞqŽ÷yÎrzÀÂu7*òhªvíÚ¹Ì ùÈ×®9[óiqª¨KÑ
ÄmªÉHZ‘6ÕÊœUqy»ŸVMo³{=^òryÛÔŸñœk8Ã>3üXN|Áëê4œ’³8sP…‡*sïÇ‚pî¸§ƒáóÅ¤98{àÇvø¸;µwü âXTËÕ¿œÁ$¢ÛPz4á^!‚†]AÂŽø¹r	ÜuÂÒµô£ îZúR<¼­/…gÜ­+h€÷ì`ñø^,häóq²ÀO­£õÕÀýû.XÍ ¶Ïä¤ášÅ­ØÄ½®í£qvØ¬e¥Û^gñ hmh/c;#ëuï8ŸÌiÆsµ¥þMù·Ü‡CË½`ã< Óž“½s.ÖÎ[Ù:kÖGú…Ìœd‘ó7(æ¯â=ÔÎ(–"¢r Šgv 
œ*è(„Y_ê’ÞE¡q!Ç[Û’Ùè„ä¿%úu¦+¶C9ƒbÙÞÑêà×q®iR@µ7—´bv÷á\SÎÕ¹¦1MÀLÎ$-—zžÎ53-1ï®¼å™¢¡Á=:æ£IÔÌäjrè.Kw“ò³p7	_wûûsù<}@®ÊNm|®º±môé˜­Æß‡OÇ,sþ}:jã›|mÄ7ñúd8ŠföA\ –ëï×Kc&*n~~óÒ|~óÒð´öïÐK£¼“—†¥T.›Õ©wöÞøÍ™?.ÍiŒÞmW%rS¶2ÞA¿›i¯žUÜ áU'ZÅlåŒ •m©7yÙ÷<n¦Ns…ù¨¹;à9Î]ÒÜ‚?<zÙ÷=onÍ©^öÍ'Í-ø­—}÷)Ó1Îè§àðƒ­‰=#ôe{*ßÊ·á3µ(ß—o‚Œ5ÿïÔ9á	ígŠI÷ã›0<ú|œZjeÚDwµ#ÄÚíÕNñôQkÙÓ‰­¹¶ðB.J´¯NÁgË§j;õX+FzO¿³sW^yOÛë@\‹ßqÞ¬‚M~“çx;G÷'Ž›F‹¸'×{]9lBýä¸ÞÖì°ÜüW¢	L¼!À)i-–·£†ÍêY¢\®÷+sóð2MëŠÓâíçwcÅÏ+wñzy³Ó„šÍ.ç÷Ÿ?ìíJÌÍ1÷V2¹ÀÖ°!iÓªÛBpQq?J“´9.ÏÓ]Vh1ËJù íÚiÄ•Ô‚C-´ßí4JÒª W%ëW9ƒW8~šÄs­/ÛSh2ÔLÚ³6Šõ*¯âá«ø‚»hó€½” ƒßí–ÓÑbT°`ÄÓQôUÔ]‹þ1Â¥Œù_íõ£•ÖÝ¥q1N“Öýr)v´2S?“:2ÉÉ½ ÷  êiŸf·Ë²,Ëô`^@-Ñ¾ÒÜK­ÉÉü<ûÃ$|9Œ.¿ö»ò!U@QƒAÃ\®ÛÕ¢ØÙ*èôvŸŒaeƒ#ÿoVÙø„¥­Ž¾’ƒáJÆqí‹­Fqÿ7+8Ëy©WñK-´ê¾Ao„ŸÆåCÿ„Û«öï¢xªs»Û©0Â£`úÜPÃsž\¬Ž×?µUÀel[i	á¹Ó_÷m¿ß§q-Fðz{ç3éŠÒOÑ¢G¹€9~%ú§ÅÖ˜)ÙÃ[¹ÍâçJN£¥¿,­Þg¤°Ï3,™ÿü<gñÓè=K}¶1ÊÝeëÑÕõŒmlÑëÍà$_–©¾çìmÔàqrß~¦3XäæêVtw?ÓF÷¡y9
]Ï¶Fóó3am>3 aÜ\¢ÿk;_–RSrÞ—&=iÌÑ6_çm.ÝãÄv&2N³eK¹àúDéA&˜îÅ‚@;3ÞÃ@}3§ËêJtŠÿèº+Î…~´‘²*8aþSµ~z_¢ÊÎÑµ¯^2­µ]\¡l
y•`j•²BôLRžA(ý&;fy•â`DÜn^l¶—U”mi”O¨Œ([ê/ÔW^œÄãìç˜°?£¾dß3l…`‚!¯oÙZ­oæŽÙ<‹3s†)Ë’ÔLw¢šlÕÒš çA‚ævzO‡5}-ôÝ UñX 9€&qRÐòä€&e…êqÅ@LpÉ$¡èQû tVéÖàÜL¿ªzâ¾ÇúðÍ©}Öh–´µ-“ûÎ–Ú·eÄŸ§v¶¤¾!°„“úÏÂöÂ{Áà"l†AÎµ˜å‹.\7Ç¾Ã[_â,ß'hú	¯U#ÍðÓ²%¯†qÏ"ãRJ 'û×ê‡¶*VPW;BÃ	-úõ²ãnË–ÝzúméI,@ò&~ÜèMŒrà×ÂFg¼yè†mRnÇÍªü™½ˆCÇp‹µƒäâ™œ‰Ï—¾†3ýkÝ|i˜5-Lûã\f£kåÙC¤1&ÙøD!Ÿ4!çm):–raá…Â¸vv&«RñÀÞ[ŠQ…ÚŽFÅ©M^\IZq-IEóÁ5ëDÿû¿þó,Å¯‘¸V}€­:m^Ë{S):âP­¦«f¤ã„­hÝ&f§bÍDÞ¨¦,m”ÀËš2~ì‰TŒi	-¯¡Lƒ=Œ""*Ü.DŽÜ2eÄ´×ô<“‡p&¸­óýtƒU‡c½¦ZqI¯g«¥1MõfšŒz39–Ô¸•Ôâ:ú	*øˆ8-æ1Tƒþ!&T–øÕÏB#Ó†ã"ý6P‘'ï¸o¬Û&‡¤ÛásfúœðÇtæáÄœ…nBuéÊî¤áõÞÖ9|¦µ¢ÔÝ³ª®ÖLÕU¤Q<T¸ÔñvÈêDÖìï¦],ÅÓ*o£y4«JDþ(¼C9w¿c&véˆœ)„’dõhÀ”5žP­ÒÓ9Å%¼Êw‘ú5GQ´©éI"TáüñrD?ú!ÜüK4É“TŽ¨ˆÊ)êŠ›¿N€–R¹  ÎÇ(Ü`A¦Ñ´–
"Ž~šÆÃŸ¦¸„ä¢›÷½*ƒ¶:›Yˆº6âÜø—ðv*˜_A}2“¶¤µ¾dVI{¢í×šxõ&[ {Ì!–Ò¬iPê`ÓVÁ×cà“©´Ù•Çfu6–ûUÙ¢{y5s*l[°ýõµØ”b¿¼NOnþ:Èò(Ò"Xå}©³+ÑéXtJ6žT=É ¯E”2{`¹%è1sóWr™)atsÕy·%+Œ¤pÀÔPEà”o¤(>jâ×Àæ´œúj:@ð‘ŸP|/ZÔ«X(P9L¸µäªMˆ‘•U<Nâ}jw÷Þ¼ßyµ·¿s°ÙÏÆÀÐ&i©
{­„úR`˜P_¦ä‰’Ìçþ=Ìùwê-Ö¨²!v‹%Ê7LŽ*¡Q%ŸxTÍêe7ÂÜ†äŒ©UâÑøg§ÿ³öwºß–š_Ol‚ûRúÎªòmqauæ8.›ôÔ3«Œ%’6«Œ¯4zwË0z0I¯ÿñŒÑ&^pîýÎZ)ì?µÆœk¯Åq-Žˆ–Úò™tå¼/Iþ¯>ç{¼­5`Z÷šWÞwWÚíßÔöŸ‰ÚžsuRs_ÃŽY5sýz¶ŠŽò¾ŽÛ¶êþ¦ÀÇÏß‘_â×áß
ø…äÏ>±ÿŽ¡’gtµÅµçÔè5	L$™NÔ/i’ŽYœÄp2TSvIãr<ˆºY²±`Ê‹ðß­GÍ*;‹;H'ßóŸ¦)Ôëh¸T—«`%¹4Sj½”]xt†5ð/‰«ü5èƒ<ŒBüÙˆ®ðÇ@ô]GëÑ@¹ä¢sò´Dwd…ÁzÏø‰ÏãLLr;tüÿh1ê1õý@Tgúí-B¿¢Gm®Q‘28ºƒa>MvŠÂ¼]‡àÌ‡i?-Š¼è.ìàÖ1€N´—!/	µ×#ÙŒÖÅæßÉ0¾<˜0üÄÝåâeåP¨35Ž6cÀ®Y5ÞVBìãqîÇUé5 ÄaZV)¡Ä«t\b7øº9žòoÛéq6ÎmîQä€>%®hÎ]`áÇ¿.²àfF•í,.	üè’ˆ×¯wt”÷ˆ	²¯O‰	Z§óÂ ]<Œ’ørnÄ¦7L‹ªûáfCZÍŠ†Äá¸ÚøË+9’¢J¼|ñ`8>Ü	»Ô¬fÆ±`IœPìŒ¿}™2ˆÅ‹{D4«ÇO‰nN×óBº!6X¦¬Œ6gÔ[À03ˆybèrOÞü%Vˆ·p'Œ³§Ô„w§”tm,æw›Œ²±†$
ý4Ü¢ëRæjó	s{x‘²’LåhÜžRB·]N–žõý	,¦ Uã31‡ó'ZÃ/»ƒ|ˆ·—²|\‡<€8ç§iy~>N‹ÝÐ©óü9ü‹ãîOµdÿÜ5yOp¶„¤'iZv2ü–éC‡Õ¥Žf|îo`"'0¹‚] H§hò?Î†Ñ7¿DéN;Çz3ÈÛ4ôÒá ŒYÆu!ÁUÇEüsœà³á:›W’•ˆ”(kÁÞýðå•ýô:êÒ÷>ÀþíìVñV¤ƒ¬È¢-èi:¬r Z¢”Àk ÞzºBý½Ñ–VÖ:Ã£ïT:£Òþ?Oð1†"Óª>ÇÜˆÊÒ¤ÀÃl„(5št{ª _ßu¹°ê•ÅÇ«:C§qlZïêÿ=~Ô›öÔ¢õiØ®å¶'ä‰“É[kD-cìÿçôÛ¼ø6ã>]±à
M©Í¾D¿4¢åËÈCÈÒ——‚î¨	GXjÀUü †Ž§géÈK‡Ö£x|YKŠ¢8*šˆšàæ×ÇÁ‰‘=nQ¸é3¨â\¹Y0rÔìÃú—ñ	!6ä5Ü¥gYzÎ.ÜÁ
£Q€-szÅAÏ^²W{Ø?@4Ñ”ŽŸ:NÓÁÇ|ZÉ ù"1»K‰è„É:†ù ÀBË} #»@ìºQê=ò›ÐpQ>®N‡—$6Äc’ è§øÁæ&¼Å»ŒÛUa"p0b”}qy9:˜a0Ù	"dD“‰†ùI6G=ÓKí\L2ô'fÑÿ¬dBð
HOa#R&‘Eb{^¦Ç•ØBŒ×†ÍH[Gll+õ0'þ8í{aîb²mÅF—æùæíïó2›nÚç††‚X´AÇ‡þ£Å»Uç„¤Fë}%…!ª0iT…×A|5$Q±cgšÖŠ %Oç	(µÚv$`kÀx éìGªP¹½(Ã{·Õ"Ç»ø$ÎÈ&_É›“´8¾žÅˆ×Û1‰`µuûUNÅ6ZëR’Øòô£~P&ûdÇÇx>Áxôá!	ÀÇÐð’j¸§[­ðe€V(´Í Í†]ÙôrÔ]]YY‰þ1z"ÿY{Ô³©H‹+—Y6ü¶“U­;õ7Š/º+
zZI†ÁÌsO!ÄZ?zSÛužaäôx8˜¢!Xž$B²“³“)Œˆ„1&ïÁùT
3)†Ù(ãù}1/ýÜÆ’öR+9ö‚ŽûB5³=\‰Ä é×ÔÀ‹'áÍC(€ˆ:bc!wøÌ½X¨W¡ §Lþc	"ŸÇ õôÁ‘Záp&d8æp¬·Æ¾"«õ[’]´ÁX1šZÌm‰¢·¡~¬¥k[—\N¤ö¸æÐdàÞG‚ƒã`>RPŸüøDÎ¤îôãé*½E(ÄCÝE5û\%KØ¯i±õ±áy²Ã@‚3ÅôúL’³2g{Î¨*ö‡åoãxŠgã¥Ó%€—pè¶ÆFç‰é¡¯YYŸm!ñð¹£Ÿó´`èXÏn”3)ÙQ‡\÷–5tñ„„JúAH£•pxxxáKöà÷ÉG×Ô¯Å ²ñØ˜ä% šÛlÙ®ž®UätÕ1´Ñpjr²ÚNGW&.€ÉX/¯RLˆ·dhžä(Àvå¦³ðéª1ÎP(ÇlL_78å[c4ÞE8ä”¼Ã&ŒÎ§ýh<F4E4­2”WÆ U9ºçÃÇÓ1EJ`‡9ž	œTÇñY^,Ž“¶$ŠáÛ8?‹YÔ‘|Š!¿0ÿÇtbÌYÎÆÑÿ`@v0Õ4A èL~ìfA…êÊ-°hA^ZNð8‹&é07å©~´3®
à¶F4•¸b¢HRShýÖp–EzÇ0ó‹-¡Û2«¦lÀ}k¡&uÛ ä¨f­‡ÇÑB¸ÉZ-fºš¨åÚI{›.ÍhežèFcã²ŒÜ	|ÓOÄ7-˜—mQv-–û´äE´‰kÈUØ.÷~Ë©mÖôC¥ÉönÃ,;¿™V] (§´B°²ý>ì«ÙÒÝhÏÄAœ¨ÐÜBÔ3§æ™¿E¯ž-#É¿Ë­8q†9‘‚û>Zô£Äj3g#l¥œÀimøµ,Ï5€©ËbzL¶ÆÔƒS5ª÷x†>šíÜüÚ¸Ç·‹:AµÛi‡èê„Ø“Ú+mO¬ût³¥vˆnÇ'ôáÚmŽÐµÐú{y¹d[s½‹¸¡?³ž|ö®pS¿Â»eL›•Ÿ Þd#ê|ã©A¿‰?IbT;—ÐK†à)ûTònYJ„ŸÎ~WxGh£èwœàk ÌÀàu%0¶#zÎáÁMM/i|(‰þÐ[Æºy¼þÅ ÜBz×úØŸcÈ”|¼9ÈÏ¯˜¡c5Í›Ðàµ÷hóâ‘ïâG†w8œƒ¡ºœ@k¤ävý’€¤§€Diñ|y”tÉ_ð,N]iÚXÂ=á€é` Ùn*NÎMQï Äˆx}j¨Õ:‘áû­°DÇù`Z®£¡£:©*ÌØ
lQ|ÁÏ³¥qîæG£µ‚ æB£Hš"ca½°œ»}Á€f]·	@åÆÝ´tÈB¥ãSOú5måöxcíOÔø÷±~QgØKrŸ_=±çØbu9÷Ä–®¤M¾àƒÌ%#IÌ!MÝö0eçÚð+ÉFF=.=BM\SÇ&i>¾		™F–ÿ   ÿÿì}ërIvæ«dÃ=Kp† 	€”(š¤‚¢.#OKM‹êžqh£"PjT@aª Šl˜ö+ìÏý³Ûö‡1¿&ü|?Á>Âž“™U•™•· µº­š
(Tå=OžëwÌ•Ø	´9ÐÄÈŸ³Áe	:ÕŸ«‘¿._ç‚É—?öpÑ„¦FÐ±“ü0ÐÐÝjp­M‹¢ã[¯¯¬—ûiUXÈ5v(ä}—Íï~L-auêÊZÎVØÖÄÎAÍûF­¥áõ!¦¶˜Þ¾[anÅ>y¡ç=bÍS˜;G@Ò¬[j0é·ó\Ø‘÷1Å+ØB?›¦XáŒÁx–$Ã8ü·ä­fƒÖ¢W,¾.¬à0»Ê…î—½× h·|”\5vÈéîÂ\Õ`6BhÔø¨ÑŸ§ÀEÎN‘²5Èà¨ñ¢ÛmïÝ#n»»×ßiµïï·Ú;÷Ño»×jw[xû·îÕn»{o´×~ÐíÃ]x¼ÝÃgv N|žê^»×¹ê¶ïßõÚ{÷ûÝöÎ><ò ?t÷[»íû»ìÓ~{çÁk±©ÓÛ=Ùßë±V‚ ÖíA™îCW÷Ú»÷Zíûä>–×mß»·°ÎÖÞÇ_à6ªÝ¹¿Ýï°OÝöþ=²ÓÚkw`Ûz­{íÎ=hÛ^ï·Ývgš¿¿{Úk?x@º;p*¸O°¬Ý«ÍO=:ÝÙcmÞƒÂHgº‹ƒÖma£Ú½=¨¸Ç>À=ÈÚÜÙíå7¾¿ãO[sŠ·É~{o@ów ü·›áÝ^{îÂCdºC»ñm˜ËýÔÓõjë““Ý^o¯ß½vo¿ß—a„{0X¸,vqbáÞnÜkwöZøç´sëÆ¦bç`B°QðÇ
W´Æ[‡wïÝ#8´ýö>NÒ=\+tÔI1þúÃ¾Vo³è¥©J´ÃÑh=‹]?%Ù¬ò³O°xy ²Ã]úm=«'¡ö(r¨c¬ù·ê¾Obtób8Ë„’>'W#pµš I¢í3+Dèil›*ºc7/ãäCkðÒ8áÙAñK¿]EYtç·/ÐÂ?˜‚}aÍ5!‹í_“óh^À}¥ÅÁùõvÙ{ŒÓd?c ˆ*fÓ¾Jƒ)v<ºÅù^t:·¶Š ÿ~&SLÃžb ñ­ÝŒ³gý•§´ªˆ}.6«IÍd’$"lHQî>Ìûº-˜‡ –NÔ)kuH1aÅ0ß°Á,¸Í{L	Z¼¢¾tÞá4¼H."xþ·a€r’8	¨ùaw…FCEP±óþ†Ì`0wØ0Ë$×üUÃ)žÔu
nXÞM61u1RôñÞ¦mE¤QNt±#Vý¹†T?
RzÓYOÂœÐóYö~”yE@Ó2D
_Ètw À¿çRê°Lz?‹†£™jË­O(ç6ŽEŸÈ³Ô¹\©Ö’*LIÊšòöMf|qCžÏîþ/ºS>KÆa¦¯Õ€,ÝÒ§àtíø¯$ºd³,aÔê"’a«m=]ùþA,ƒ#ûðE8™K7U±ºšcòp›íx‰¯×ƒšÝ¿Ž=3š 0Lð‡íZ¤	œøÓ§Ù²‘õ¸@|^'S‚©¢Z@†T<GoG–”©ù=.ÉUÖ†ã³Ä›
u²mš^ÇJ](˜
Õl’)*iõ4‹§±±¢z­ðP2¤Åa°i÷ï©#‚,UÒãéœs°«²˜ wÎ¸÷&cÀh€Î‰WÙEbŠ/·t‘†¿òPXýAíšœÐ5Ÿ¶îzÊVœ)ÑrVö…»±ÒþžFûyC¢EuH‰ZRRUs îQÎ+¼åƒª¹•]†ñæðýçÌQ8„=?é§É„š¶©«0Ì)
Ú…3iÕ›ºš¨§8HB´¬0Ï†Å«fc’‡ s•aI¾ï4Ÿ’3®BIÓ>Æ¢§NÌ¿ EÏúã¿èsÀ¶ÏlÁ³%>€=ãÎ!°Â1Ð¤ŸŒ§óº„lO‚«pˆŸj,÷ß‚ ö8E†ÛoÉ³á,–<G÷Y~¹ÓCvÈgÁWÑc€XÚ­Å‚2°p`T']!°nØ—Š½“Ýî2Q€'™ëÈÆ=bÂ‚ªLT·šŽÕkKG9:æ¤ù8ÂL‘èàèÒ4šæ5%è¯°†±YÓnÌ··Ìˆò­´xw^œ<ò…IQ³íí;ö¦t$h†]äêsŸ‘éÄfã40¦ g‰ÐÛQw2Ö³¢÷=Vto­+Ú5¬îÕ|2œÃ® ®{ß”å,ë .Œ#i€”öáƒ½ð¿ÌØ+>˜+6¬Åñä= é&‡Ö|ª;_¬É¦m^b÷«Bž)?óQõ@WþºæóƒîQ]ÃÐ¿ûë4¢~‘Y8œcÀü](ÅjÿÝù·/5ÌœöÐxœ|˜ Y¾ÂzË¤ÏdgQhO³	É`ÌƒŸh2_±2“ãe‘+¸'U	åçTNamO¯ì“Y‘âs^I#ÿ]ùnê˜‹‚“°rÏÇ^3aA¨1øàÅTû—Ql@=
úýp
Û´ý§,™è)=<xtõØc/J»Q§£e—ÖB×Õrê!ÔMü>H';þ(@'tX)9¥d5ÄBŒ[(}«Åà¢žá§cC6·Ãq‚«¨]Ñtã±¸ô£Å‚$Ó ñì½Ñ_†ÄO^é…}Aû¸ì¿¶O!¥}8:»"S¸²ƒÛGÀ*]4¦p•,Dû^Uq5QIc…Ò<<Ea< a¦KÃKW•üÞÙ<+€Ó¬£(0KÎòzŸþ E¡‹4¼
¿Ò‚:k+ÍQ§v¾qLçÇ$˜†˜G}Q]È#¤ÍœÀCAÆ`,Gx…Àl$!4"óÖ@š¶øã€îÇeÜé¥=P¬6ÉKªvkßÔL ü­œ×A]¸ÔãpZÎTõ~I<*¾ö>dŽÑ²
•ˆœJ®ŒYã9@4®E¨ä<¡AÓ†Ðb=žš\¢xL´ë[èå6Qj™ÓùJ©‘“Xž/©©A‰›Ô<aŠŸ”/KŸŒ¡½ºáT}5ï—±•æå J]#ô{ú¨H’Û¬ã»éŒ”/>·EÍx¬êž¼Õ)ÿòÖ‡®êz¨²ª/ËCìZ­Ô]gpYÈi³VkÝÑ­ ªŸ‘ßÑOŒñ]Ã†EL£*ûƒS)P!*T?£öd—6N‹WX—ßbWM®‹]Þ+GùT‚ò°Ë•8/vñ,m\˜#îs=<o‡öi`ÁøƒVÐdvŒ¿àÃ”ÑŽSžìd†ìF=Ò³ò À—ÉÈiz÷Ä{Êœ˜±¾œ£µPvì	…M“…fÜjïÒcªƒ.×üð¥>×ùA¼qKU~3ÉxÇA€¡›øÄ,ÌH~Í0,Ó5,ïŒÜïO”†¤ñ>SbæëØeåîØåÍkâ9DGf2ïà¥NŒ'8o±?ì »Ì+ãYÍT¬)ãpÛ jÍì!»tþã•;œ¤Næq\ªmÊjZ•ë<<a„ñVô($ãd€„%ˆfºÑ%ÖÐû`rL6ô¨å:Žé,¬ ÝÐH|Ö×j]gµ$Û@ãù{ýKá5†^ÕªI+ºaÆ¡€±¡¸ˆõÛ„¯ð2ÈŽü}<ç¶gl0x¤Çe¦âY[HÂA4¶£aËäÔåh‘Ò?—L¼väÌ
H-Jcö£×´Gf0éüÄ„Â-±’0µ¶ý9OTjÃéžL„vÍ®2ììõuvÕN|G¡*™àhÁ€+YëØ=û{gi”¤THß-ï›ßG=é“«…ÆBø4ŠCö‹þ=iv)4XùÕøF1úô…â›ñyaHéÂw{8[eøÍÔûg!ˆç˜JbçQ`³aåžþmþ"»‰ïñï®Ï?-«¢äã¤dçñSS×*4¢„`nÚßæÃå›ü†þ­8&§£$ÂeQ~6ý7ÂãÒWƒ
8ÖdŒO-ÊÏ¶gÿÈ@|K¾köW!¢†)3Vo>M“ñ9&V†;WpŒèÊÓê)ÍÇ…‡ï(‡†£×Cˆ¢Ç/¡ç{ÌVÇ(+-!^é^”âL•˜ü æW{ð™³“-ÄoÆW´º‰i-òŒ-Š¦g«ð‰&y˜)r ©@ž¯ñ®©4ýQ ýzÃ¤•Áµ­q¢?±ÉV`1ó"j0™–U^{×[áÚÑ;ã=0¯pïõ]{u×XÛPï
Ã‡‹ú']‹(Adgß¬¾zËÎÒIïÔ[NR/½žjî²ÇÒ+7ŒoÉûÜôÔyõ,Vï9ßU¸ Ýý|pÀº‚¶4+àç»¡OhûŠílßykœãqÐOš‹þ"	ÒÁ
“­”ôóõØ‘Çy?>sbÙìüfÒ?ŸÇAzL¼|CÿV”á# ÅPw9dGÄï&¢‹BÙ9¢¬ãÃr¤O~W‘M%ÏŠRÓJZ­o›#
IÃvÆCîÚ¨'¤¥¯®äpljÑËà*¢HpúÁð7q®\ú]fï™	Eê_9ëo‰ÍÓe`!mþÀó±8nŸÛÔœ;ô¾EŸÁ¸ïVWQŽÙï¨z´Õ6de¢ájŽašax Ž£ïlŽLµ-7EkŸ¤å§É!s4sÝnÉ[ãù1ˆ²~+È²0ËÐá~…óC)éç{~<†Žœ =?èõ²÷ñáÃu®•#˜ÍSèúzfUSÚÏwf_ÐÎ ¦õç7­Ð”yŒ!=­~‚¼ÂÄjËûùNí«¼;B€v­I®Ëù­(š×`]RürŠãU•Õ>’ô§S¯q£AÐ>°Š*,/âco=ç2ï(ÃTIç©ð¼ôÕðÊªÖ…¢µÃf¹ýO~ÞPãCÈD/éÖ:ä™LŠ¿Í~rT®Ñ¤¬FúÔ
Â*ñ	5’¦‰Å×d«^<5ÔY	¹~)€Þ²„U¢­
NApY°,>Éƒ©Û»/ç^S/So¦kIgËHøËéü¥ò5ÙÂé‰¬Å½ÚÞæ¥ýbèí7¬C/‚	lvä¹Ý¬Ì²¦¦Ü;¡’
;wR¨üà]åi4ÅXxM)Er]MIÅoÞ¥åiT5…å?™øŠ†øÓ¤‚î÷Ñl„hlG‹j²Ìê¥Ë•ª^,“ †-Àò…–³·ÍÄŒÏ2–9j"-D8š›ËÓX˜2V¯[Æð®“l³„f@jÞ)RmžýsœPÈ-Ö³Q©…ô–Yµ(ùu>H±1öï{ÄVÖcÈ
#N¦šDà•dñRúçÞŽ”¬¶RƒàûÄÌpB•Ö¦­êÙ„/¢/Z>fTÓUË%Í·¥¥«–C¹©Ž¶ãñuŒ ³h.7Ìô¹öQXµ[J&ß‚“3&øý$ž„sM'?¦aˆÉ/ìü‡ná¡g¿.'éÚ&EJ@»üdˆÅü|5I6€A’èç4ù€ŸýpÖöMñï¶Ø€Q×žz¤qœ§p'"“w¸=êË´äis/ûS]åì#§4ÏRè×Œq–¨‡
LöG<žâ£Ö8g1ˆ£"ð%y™\â4FÕÜ`Â<ª®çaþÁõœµ:d<8(¿vÍ°üÚc‹Ø°~µRF{LóXVK3„4‹hG	ùÇhàø+Ã4™Om1L*XFË<£^ˆáÂÔõ³Ê×äÌs•)…‘•X)
RXôU:
-]nñÖ¨#{E¥+j
%k8Çeu
.âÓ"JNTŒmiÎ;âàîÊ0lBF“†@”BžÙÜBšÊsj&Ó³¡éÓ~-ÉAWpr‘Ò+u¸jéEO²W^fX^‡Ãðæ!'‡÷œëÉä%<·êÚà¡§?éÂ¨!VT‹¨·4>Î¿NƒlÔ]ÿü:éŠû?Ô.¤ñ.Bt8êU1¦‹ ÙxØ8¦]ü$ç	Ìc|ŒZÏQ¨×¶xGsšLoZ†Yc@nû2ŽÛ|:Ó>êv¾v-Ulœ˜˜©.0hÖ–ïþÎ¦Õ‡Cã¯ÖjØzåµ†éíY»LÚw>lFø§ò	×Ê0‡"—„Âãèç¨OÃ˜eô“f ­ÁÖ9{ü´5ÔXJª _›†:õìU­Åo*jTv°-œ˜"ÓDwGÊ$Éù½"c&:¢†Ô¹2ÂÞÆ÷©ŒÊî¾ÄbæðÝÂ”vó)µ(£…Ø¦ÇLËäÒ“(<åPø1ˆ& P>YÆ™ÔN•1åÒ’Ú€dó…+©ò2¾è>™`¢?µ¡‚œ"€Ö ,‡~f’~Ev·ë¨yŸ{Ëú<RPñó;pÔ=Ã“oÃ®Äÿh
ÞŒŸ™r ¡’¶Ùä´³`’él’_3fë,LñÞÛôR*”dÉH¿E•ƒñ¡#J~‹b R¤»[p]ŸS™ÒØ!îôštwáOëüI‡A(ÿ¯½³¿	ÿü}ŸÙÕ=³»Ù0ƒáA:ÇV 
%Ôœa:†aLÑåå$Ì²GÏã)ür@º{¶BÅL{pÎH ë-r©
±ŽjòèŠÝê-´ÍR©k£B_ŒÚy¥®®èãh‹~Qú¢å4F^:£•´FŸ‰ÞhINwä£XJ/ 9Ûm˜Ø"‚4Ë …(­á­I2Ê.¡9ˆ ¿Aå:v¬Â\¾œÃÈ&ì±â+}ÈvìŠ—ˆ–˜WŒìW¯ŸŸ}{þÇ'/Î^=9?iG“~<„™úè¦ÍC—*g¡æÓ^§…³V?#Uä7Ûµõ@~š %µ}þú¾ÏAã·¦-Í´~wS¼i®¯ò«¡Ôõ0'x<âÐ×qß"'a>Ú:mân™#ËÍx¨¢<Î»uk¢hÓlÚ(þˆšq"¢é˜Z,y"*88õ{%³CôIõJ½§”M/Õââwö|†oa>´ÜDÕ–}I¾¢ãMo–¦8KR>B—Y\f9°¬~Pk7ä<¢!p(èÅÖ„òý;ÃÎS¥·€%“‡´ÄóÜe<”Êëx¤ÃpÖ, )ýyž	ÅÔ&ßF71ùªHpI@I$ÑÝÛÙîa>>Á.aÈ×E1ÁŠœúËiY^<]ŽN’×°¨P}õ2Þýµ%î¼Î5,½8Ö|äbÉR¥>ç2ZØÔ8Â3Ê)¥öÃfsG—€F¾(ëO¦²iö…Ú´‡¥ÄcJæc•pè¾Öœ%Ê©]2tÂ@ˆp>4ŽO_žýÝç²£äÛ}Ì9Ê…,Íü÷ä©2¶yg?Å {p|+òn•Í³2qêLæ(<ŠŠœ’…p.‚ô¦!`þ~b§ƒåÎtYnañ£'iÑ(t/×óS( >N©‘éF'x¼DÍB9.„æ;úe–ùº˜.âö]½Q¥cÌEQOóÍÛÍ¿%ÛÛäUñpÜÈƒMüXž/“«„ÐæÝý•Fþ:,O‰ít^¥ÉäM®]Z5ôö(Ì&gÔBîiµ+{ 9°ÙÎÝ Òöõ/6õ6õOoG×•É;–,VæYô?kdµÁÌh>&n ù
JüJ˜óƒ(.âp€ñ×¥ÉŽâì›Ïé*üüiùrÆò
oÓ¤ãàºÙÙ"ôZ
§¿yQàÐE‰.×#^.Ú¥Ï”™§?Ð³_°þmÆ³»‡Ñ$ â˜Ý¢²(³È¨kj[¶‚š†Á.<|6‹cå~Ö[DÑ¤é]#_l¿Yn±ñ”K-6#ÑQ“°K÷ð'4@A:f	Ðìî‚jJ;3\8¸Üs¸â[aY%ËäXÍ7Â¶Û>ž¿²ÙÝ_S`T*£â§´uêVé¶ž¡ZÃÁÒ/4šb€#¡àbR¼¡Ñ`³H_2˜¯f†)½†=ëð›•YÖ-Ð[mmíÿ:zçgÇ-;i»jP4vyÐ5véÓ«‰¿•Ty2QJ‘2Ò¥’òôš*/´z‚y¥ãæ»Ç Sü)€4Æ–`DWNÃYt÷owÿf$$ã€Î-Ok˜HK “I?š1<á«t%àYõµ~£=|çi5íÇa0y<¯¢ü¢h”žšEâašrJ@Ž¸ñ2i<22`-FüM—ÇýâØ0b¦iaÐì3ë”–^…—ÐëÑé‡\TÚ•¹ì”e/ß>&ßD þ¦’È–4‰ó-uˆDnÙÍî‰¬j0ŠV=¢UÅ¦ÎÙ¼=~öGAên‰Ã¯¹FÊ«\Ú³P³Õhº©QdÊ-Ž|y½isY£ôp[Ã‹º®Q´Êh€*ˆwø¹õõ‚×sûÎ¾À«Nm=Õ©]Õ:÷Œþjûv5¼>ŠÏ^¿5à°7—qò¡eNœ_ê¹æ¡ Ž§šO[1zTýÜL.qv7¶*îi	:ÈJN7ÚU¯óÞí¥æðS|Ê*Ö#)rÅ)þz$ÈÚ„La’ÞHÎ³
hw37[<7­(Ò«ã¾Qaï¶Òðj¨ÐÇHžÂ¹®hCqê*ÅÎ†ð’lÜúØœò+'3¥Žê“ñ|J+,·ÿ—v‰Rd˜^ÓüôrÄ±…4íï¨™&ÅMz|–òuÆÎi¢}sAèc>"‹ÁCà%Ìu±¦ó¡EÊ´EƒÙ„ä\ìÓ\@ú$NÆtöè¹Ý¬µøü(ðÂcç¨Á^ñ1ãµœã^K;_àÅí¯´,[¯ æð8Ç;ªç€——O ^F-­Î/	½ìP=MÎ@¾³ÆYö¤=Àãzäü¹È¾û•æ·)«cR‰4?bº^G}ÕÓ€]Œ°h6RÒjRã3lïþ×?ý‹¡œBçp÷#üÒÌ6ù~\©òsá^ÚÏ[sjêí¥šýéµì§gÀ«¶®/o}^ê¾·˜rHuDçœ’óÙK;¼®æ€¶´¿Rgƒ‹(ër„òs	º†8	âo€G•^ÉÉE™ÌB+û6¯ì/°ŠËòJ³·ÜÜyú ?
ÒÓQÎzò¬iaÖ83‹fcÜöW”Ü&&aú|Pðâùwä‡Èê<ÒÿºA¥M_¾¸¥"•µã·ÅêÐ*Cªè=›½äHÈŸò;«W U+1l:}aåûúq*~/µX¾åüöñòdórGÊë~<RRá¦HYã`‚/ú Zò——'{´LØ@}ê!‡ÕöUq:ŸxF,Ø<LØ./¼ÌÆË5iœ*´­U•3j ˆï«þ„Ñ,€Íœ;zÈk-œ  7ÀöfýÎ=AlŽ–¹¶¬[,±Õg­Ø
Ê)ßN®®Ãa4$ —G×‚·Àïð"Õ½hÀY¸> ]‚cáZ‹± ÃW€'[Õ‚ÕÑ©ïki@NXE“ib¦ªÆÆ¡”&éáð$M“ß„—39pÿû$ž:>Ð°o^£n8®áC«”¨ù3£CÔ1ý3g¤ª‰Üìelw¥ÉÂqÄš­Zæ6|«o‹¾ðÂ“+È(	×,É©¤Ä')– ÃtXMñPƒÔ-[*Š8Ð÷kF`W½ê1&H’v=³&PCÁTÅ [ÙçVðù*õj™Í«ê7MoìÕ²8+;V[Ó6Žé:“nj”ŽvÙÆ\V¯=¦ê5™¨˜_·23R MÇÊè¹”fŸ)}±¬›¾ÄpÐŒš!ŽzÀƒÆDg)qÏÖntoÙZ[Qôožsk.øqÃõá€Qr…P@_Âã¢‘ÜlW-âNs¯'èA›ñ~…´®ZÙÃþV¡VÉD»cu¦ŸçKo¬Â2*gåü<ÛMfð¦ÓS†’`êƒKìŒ8îá~°Zð¸ pw¹†Ò¶ZE{;sdê×†¨+]lÿš°Œ? È	O"LjÉÈ¯·5RqMIBïˆEë®Â¼ò§a8¸€¹ÐVY©tÑŸè‘‘ÖN×NGË©Ë±Èz>½±¥­1ƒNœ¨ªø}°ç×8˜Ìƒ;y>ïc’Ã?$ËjÞ¯€,ž¥´¬ÂGUp_Ý×ø§æê€ÖtgZRl’„šnÓ|Í´ÆBQþy3³i4ûä'Òœñ•”Êä2÷Øv»íàånÂ/ÓÅÓQØ¥ý8ìê{¹„´F»F9>rš ›r÷è'K	‡ëðu4aJÞ5¿–ïµg	Mðâ—óúî5ß¼EŽQ2O¹é¶Ñ0šml‘q4™cV#ádn¡¯Çæ;Ê¹dyÓ¨ZH™z¼ Û‹®é1ÐŽGf.¶bØÕïñÜ:¢ýU˜Bž¦ìoº-ƒl¢ìENK`¦LY€8'Îrÿ/ˆÆC–ZÆf÷{÷o¨I(×p[³“‚PßržêV×X‘Ûm£°g& ¥wtýJW‹¶
Vj…“0Wcˆ÷¬ô¹‘Ó 9$Ðô{JN†I4ôë>×–þ«;×n?{#È¼ðEZðxÃ´Øýd‹ÎŸ©êÓåý| UQYb
¨lU”´Jö­"¤€–¥5¦ÑŒ:d•+ã0'(F²]–Ìá¿YšdÅ­{ìð4™Þ¸W4=_Í¤hª#Ü¹Y&5û”ìæºIÒZÍ-–q‡-ÜÛSÇ¾´&nÁ‡~ïÖ0o{«JV°@K6çzzÎ:Fg£zÁiX”lî:UÌ"»bFÖ­{ºÔó†8l†­eP#déÙ#XÎ|Ñ4/ŽímÒiÃžä‘Ulð0%.É`daÞ?ÏaÌQ‚@ÄÒV]|ú!F¿FI
BÎx.¼ó!LABN&¡N0aººù$‚_å“DŽ€%(ê/n7‹Y„£oy·MžQ¸_áýoDF.nBg›A±"ÂÁ£|
çò ˆXöáaFy×-’·åÍÛchäâVÛ¥7íË$}ôG´ùæ-ÉšN¨ú‚Ž:Õo8u”ìÍI’Žù«ø‘åcœžR;þÐäeobü³¶ý„ ) ±rÖý.¼ò¾½@“öûð&kVÆh³}M&Nè=vØÐž÷›T>€©9²WäÑQÑù'=kdh=Íž7ßK(t0¯ÔDˆ*~“¿üÞ~óÖp’kïZÊjOç[&Z’xk\ö½6ù]NÉ·/¿ùBubd6
f$èÏ€E„9B~4ß´¡uh;ºæjô$åˆ3tceår "BÝŠÐö;wi¾ùã]ÜÙ[Jƒ|šoc;šÍ`‹\Ð·‚7;oÛ,€"÷¤aónm‘é¬õèÕÆz¢Ê–+êS°‰‰Q‚ÜÄCáMuC¿ÅiÔÔ«IÉ…Å´ÂXÖwo(*ºâŠ‰AµD@Ë,JÂád¦Ã9¦¯/¼^A-•í%ôÚÅ©;¶„fÉNû+¤`1,•<²±Ô•É^Ïs	|e¥t9µóS‚)ŒÎ£1Õ’©÷Ï‚±¥µ¿½¼ûW`œŽ3ŽÏÒpeYä'Ý‘Üa®¯ôâH~Î\Í´?+»^)„]C;4”l±/—q’ÀFT›´-•Ž9 €ÍMtw´[ƒŸš!%çgi2„aDÁÈË°¹,¬¨°Fø©g[ê¢(^ùäk \H€¨„Â§Gòr®måÀ4KÐ‰-ŠöžÌ|ÕéÓUÌ)Ô¢hé{w³%öìV»å-Ç3–âÁÉØ¸®à*ˆbÔ@QX3´ÙªÔ:,‹L.ý ¤^K…X‰,F„Kƒbåz×XØ“ÀÄžèž¥ýÄÃrXò‹úÖœÜ"è“^ˆù?ý€ïÛ,[É¨Ëé7öôÆ2kžŽÑÐ2FCœÿaÒLŒ:V–s€¶UÄÖ?¿½é¼…Uo8Ða¼ör1A#aô@N˜¦	Ìœ¤h­Âq†*ã/”ÐŸ¹eˆ3VBÑÌò¸‚6Ç$6A¿o0—ä¤œØRä r‘¡h¹Kl l2´áM^Pié«…µ•žs°­—´dxµ[dÊÌC¡g|E•¶=ålª¬¤kpO Ì|Pù– „hÚ„Y?3ùcñ;]ýÈœâðcj ~£óÖ(o§)™CYë£Úµ^¨µ^Ô«6võù‹³“Ó×ßþñÛWŸ¼¢“|ò°Í}Bõ)tzhM†c¥(î‘¶¸G5‹ã«‡–×¢­¬!Y™p“ÎÄ=…êÈœ]*žGÁdìº^1©«çbØú0ŠàX¶J+1zÞ|»¶÷wHF±5¨µXJa4”Éãà³[«F†kZ°ówO&×ºÜ2~£ÁÞèL¡& Ëç%!”ú©n†>Áâ#Kí
£ªôÄPñïi#è3´.Î¬5Ó¨cdkÖIÜ€­îQ´YÈlã¥s+¸¯*û(É›©ËrÒÓ žc²„­–*Â1ÀEñ™[DþcKfÃgwô©™øÑ•T®Q«Ë@º*Ž‘,9Ñ©i=ÅKö¶ið<mŠ§Ì`žÒø:1c}°±eÃewj”:P<`vÎ0®lvÓ0 5³Ñy÷õ¢”ÇoõÎ
gèrb¯mo—gQ¦Ah•h˜ÁA~+œœ„~"^”P*®¢ÖP@›d’4Ž¥qõJa`:47÷‹8Š•Ç´/<ö:“ÑÀ±ÛWÎ!íZæÉyTrrJI8wéŠ1CSg:8s/ÐM6¸Ðûíé+†z
€éñO«¦PÜ•Âd´	x|ñ§1øi×Áp˜sˆüpºãìHÉvLAø¾ûg)·G»uhta³sc„ju$Í7(ÈnQqöU©awE3B:Ì+Ú¼fQ–34¸PŠ2™úˆIÕ(ˆhÄ~Waåyyùãù”Î¾Wƒ÷N$/0&oRõÊk–@Ã §­ž°°î¤²`‡–1Å™ì:oð1U¯|.½VL5ˆ"÷r&ŽI?\¤òOÂsÛÑàáÙ v¾³Â+tA	«âíIpze\M­Vx¬¯8ôÚq­lÅÔã²%h—4n†ŒœÎtãÔÚR*ñ^ý°þEN_‹»†‹¦ltÉEø­ëre+kö¾*¡uJ¶Ñ?Îomúà±kt¸³·Ûì;å<ü@üçÈª»ÁFèG@âÅ‡Më!UåFx€"t*¨œžKÏCìÈ¯zâ¢4l©«‡œ¥¤‹üÒI8íÀ{Èùå	qá‡Oé‰/æà
nWbWà{qÅþ}áîpš ²òlIvµW°«fÎXÇ®Òà¡aÐ#³5KZ)¹L“±‘´faWÙYQe$*Áª2\]àü…†·óIpê‰–ßñòz`×+&¨UÎýÆvÀé˜D	Ô+!“<ÌoäÎ³&—Þt;bsÙ|ïÊK5y¡a£ê9¤|ƒ±¼þñÝ<ÌÞÐ|óž	uý·–³ü=l¿é¼u7Þd’’GÎ$Í”ê \žˆºíKe‹6Ö[c"rÐ’Î¤8%|ý•´&eð_u‡=¾x³i9@\9,k¹Ä–×B¶(ÊpE.@õØ—³h94ÜÚ#Òü%ÜÀØUq»Š&ýy\ø‚åd±¡%âtL­[û¯Œ3Ó`ä6Í-jý,Œæ>šŒÒxJá%êÚOEÃyaC•áeL•[“[:Šv•ÖO4Ð?
¢ëDï×Q^uhÕrä]¹Õ-[b©(mðm®ZðC_B¼®´QkÊ[Õ‘D	 5^ †ÝÉÑ„´új£Ñb6ènöç#uÈ[ ÓN‹&UœM?+•·‚¬ŸÛï:b¼t­xú,Ošê/þç»=ýNâY²ÁE9[CÌ¡8L#°L=/îþ}5éŒœB]E²ÿÊlÑÝzxßuë.ž³îïøO¤±ð|¬6¬XOpøS¦Ö)zÖåá ~;…¥¡ÔºyE³¢uOŸ¤SêË>¾'Ð¿tVcdÌy¦›.ŠÐXt€”XrfÐt¢Çl<s&éÖe´Q¯:-tÃ ø&ì(Š¼¨Ù[^…y#Èn&}â›L½‚A4²<Â´š|yn1¿kæÅØN.²0½
€ëâ7¦aæ‚CÓ]ÔiQ)‘ºoË·Ú ÙŽ¹Dº±á—OL½²pö,œ„¨ù›`IÞ·%Ž	n–jGÎJ^æÐ%G|Ü‡¬uáÉóÕ¤)y¾¢·d¿Ð¥ÚÎ>oÁrã‰—×‚É«Ym¥°±·¸û`sÈ‹D8‹d>óÒäÔl]žw”ý{Äz#ôgÁ\g Tá/K7þv‹ô@Š\ò}_}¶ôéc|Ál(Ë/ÿ$Û‚tÒl<IÓ„LXýcå_‚™¢Çw?¢FÝ4²ùXBøá ±EüS—Èm‘Î-×bé ™ê–jLíwê¾áµ%¿–âôË‹ŸîˆùÒú€l>3ôÝFî3	„Aû¬1x»Ó]/eû/º5!EKLf„Ia×ö4P.ž`ùòåk/@ýÑŸL—·½‹]õ8ƒó*CF†é¨Ç6ÕJŸS¼õ93[y0Û†+¿>	Ã¥ûÏé*{ð…ñúÂx‰×Æë#¿ñ…ñ²_Æ«ˆÙ^šù*4¯’ªµ`¼Šx“OËvy'sW¯ÍtÄa:cL—Ìsñ™øïÎwQ €/LW~}¦KóŸ!ÇÅ›ÿ…ÝúÂn‰×vë#¿ñ…Ý²_v‹’ªåy­Üž.ZÐN‹~ûäŒV¨ù¹ñYÐñX8üÁòv‘_Ú´;|hu¶}}Kô…xez«mLƒ‹06‡*ÐrT u˜lË¸3†þN¶É9£Òðåp›¿ë[s1-“¶>¿J8ð¡BÚÅcº.%]/¾^nuY^•;gGÜÁø¹ÖWñbø;Ñ²ÜB€rY³qJ¡OÎž£û³ŽšˆL[€F^FÃ9BÿµÉY’’Ëà*I·H0ˆú0lÈ?¤|{Ð8	ö
_½ÔKtž¡Ÿ<BèÒó=êÉsØqq4¼û	H€ä$…¡,Øn,Ép1ÇOÃ5Ô‘4ä6?–è±ÁÃ.vû²>9díRÈÚdO-¬"Ü’0†e%²ð4ˆG	*>¤‚¬PÀj‘«à‡QÏúáE4Jò±ä&^‚}/$¦4å¹·—é³«ÇÿsÒ ¿¡•´Ça–Ã?|ÕƒÍ¹Ú¼ŠÄ´Fyég!ûè“šq|XœoGŒòUD!%¡àñãÝël	©§/Ó‘¬ÃÀ_IÌ²º¯bcÏ–n8Ë!%s(Ä+–tc¹ÆšÊ–’vtäÊ¼=”Ë«¦Vdx&Òd}`Qç'õª¯+òYl¬±êeIeyé›$€!³Zsb	9>–©úzãx³T5u ¤rj×ŽI–½þ	\çÇÒY>G—¥&–½+?
) Þ øêÓŽù§ÕsWÞÇ˜¥Â'°¥ – Z5¿ }|2Þ‚$kÏÄµð1¤*ël=B÷	ÔBð®Z
úS«mUXˆQ«»K!o<¢%+QC4~«?ÏW”jU‹/¢“»ÍK£L&˜ÌOèC³ÆV±uiØ5ƒ~8&'L1•]H‚?Ï#2M0JÔã0%)ðãÉÓ‘:J·ì†¸Â6HÐrŒ¾j5à*ˆçáQeI£hIË¬»’Éé(˜¡0´s€ôä%mn‘°=£™‰Û´)¢¤Y§þZÛ¤¢wDhÃÕ Ü•,‡ª)·¢ŠT8_ÆJ+(FÝ~på{w˜u¢Ôîôä,@iMsÂ|o)Â¬\÷Ï¹Ë±,i7Z|Ü˜Êöd^éÒ´;šLç³eLLK{Åæìa¶KŒeè8Ë”RÙÐ+ï¡/ä¨Üìø={Ø~³ãÀl0]T)el–mò)|þn¿Joo‘œÂä*¹%5Mµ¥ãš´¯¥ôÿ %Ð~…W°ˆQÛú]Û2m[ª÷¢TJÎk¾.ÈåË¦°64îù8†Ïa	 Ä‰®Ò2	©Hì³t>¡™ÐÆÁ5r÷‘æäƒÅÌ¼LÆT³´ñ|<tå§µ6r9o'v)¸…"²>´LµWYG°}þx“÷µÅfv™&hOØàPjõ7Œ×‚e'ðaË¨–mï'×Ð”IMÞBðr+l)7þî'žÑËWá¦á	›ó°©£–ë™ùÂ°/L{nÞ'ã¸Õù,'þ?Ýl×®ðª-ôÕà–rXûÃ‡Ûëîóì¦ÇˆxU¼:P‘ñG7€Qí?Ð4^‡­ä—ZqÙº
Ú¬Ó;IÃä)vEƒk,¸[žÇïÓs™kùáž‚„g»«ž-†ï÷Qø¡¹1²ÑEKlCmñ-<Ñ˜d™|íœ?_í1«àn·@ò¢Ðãï“þÝi2IÓ	êŸÌRÔ<–sÛ&e“H@þ<â?Ï¡ã„funkê4ãp{œ Ð×V~Ù”2ù-®`È˜+X9lÈ€j²þVùÒÃ²
õTC$‹FQ¤Jô£	¤AŒ˜É4èÓŽ;[$Ã|ð©ý`¯Ê²sMµôN§x§S}!¼Žf5j¨Œ©#—ØðÙ—–t “«@(å;¬ÈV=ÁTYv›Ì(3ìH¡Gß &µä¨ö}ƒÌ§%IæŒÇf<ÊÃQ×‚‹G±èÊ- ;æU˜Í¡ƒ$;ÜueÚÑ74*$tFqF/KäèQ9fž;BC¾qžŒ§ÀƒÛØ¶qÁP-µ—û;^gãÿú§©YR4¡ÚNƒ4(ç1©`…Ö BI#1ÖtJ-2ä9Ï]@º4ï³hšð‘ YŽ`§w>u¼xJ²±'ý›g×XÀîÌùxçŸ&é“A4kB=mº½ ±IÆÊ›_½~~öíùŸ¼8{õäü¤!¬ã Ìh‰ä!íÀ†”8KÇÒxÙ¯©Î§ "®ò¼€ç™P„g^x¼¬Q5æTz2H$2•Ö«ŠL­ˆ/NlA¾±	4y"{Ì©«©ÙË“¶LÙ+r æ­À´XI×OâqÉ56cá[µ;–¦«V¸Ÿç„”Ñ6…˜æR‡zkÎ¯–ST/–ôÂŽêÅˆp™Bá€âé t[Ž¶.­¹sæÎp©¼ÌÒš-Ù´îÃó¦Û+Øh˜]y0Ž&H®u‹ ù0	Óç,tž…éÃö<lšWdÎ’F¢àÔ¸Ò¨Ó×j`4`…¥‰ˆëhLÔÃûœ¢Méôë›iÔ>DŠÇ~4?qP›šÑ´¿ú<ó<gƒ‘š½N—TEž\Ã‘é8Z;O«_¯*^Ì®‰ŸÆÁÄÌ„³£V	Ç…7èéqB£Åê²âµ[ØO“I2Lƒq`nç)S“zV‰-}ÂÉz÷×4¢Í=–	xw4~¯Úæi
`z#´9wé>{üÔÔÆÇ°ƒÑ¢¤Nÿ4IñL†7k5KKƒôzšNkÌ2ò%«Ù	ÌPñ9ŸNþ8Í	ˆèçáŒgÎ=Ö…#°æ“èÏsŽâ“˜¡Ýì7×gíK`ƒþ¨€­ÔVBÀÿ}ž¾p’gÞ›<fYÆ~×md´ý±’(C =m‚ŒÝÚÜ´D/nò×‚Á€ßÐ=®Œ‰5™±6É¬œË€¾«¥~œôô¦y¦ébÎûIòæˆ­óOo‘”®0øˆ>Òi;Ã—h2ÞÍ-¢s_gŒƒë%Šï’_cèöUt6­uÐ¦Pé¬?S+˜¿‡äE0ƒj‘j6•®o“¦ÔVV-´ŽFtNÛ1NhÊ›ÁÔ,1™›¼y¦žÑ¢ùìÖ,>_^U¼’šÅ³eåÕ| ìÈ<i*àsÓÒÍ˜f¸-èæG@'.ïxPÉjxÏ¤_ãERx]PWÉfV¦.³å©*~+ùs%G¥…†u,ªºxy.ãæ÷aœô£A07­‰HjÇK.›å–Á
)žweÞBÚó«¤¼ë‹1ÇqnË¡qpÐa\9;gut>´:÷0{Õ=]>bÌb)K^½’i%OD§’AôÍneqòÁ'¯ÜÐ.6´ë×Ð]Aº,¶Çî!Z)Ó3ÖØ¹õögiYbw¾ÆóùØ³<åÐ²ûòŠG–Ûc;i:¡î$=Q¾ue/òÈ³§Z®š¡—òãnºvÙÏxœ1ÞÓôî/(Â~ÓO$ŸLÒÅ¹uLÅp9Oy pö©gÙ˜X6Å©æ¶7‚›‘a¾…|;ŽoµIµv%ŸHvší\®o:OÍŸÁÈãôØ«„VXœó:äçÒtg!ºê!‡>OéyKšÏOÌÎ´¶U¬õƒÓa\«wvS;o“ÉzzãáÌ]DTYã¨ÄL‘u\£MŽÐtÈ}›1g'ÆZ•¢<&ð¾òJô¤Ñå=lO¢«0.gþäyíY°`‹—ª;äMg¯Øð=ô~q„·ŽzV€J/].IŽõmìÒŽ2¿EŠÄ¼Ùd>$Ôš/ñáö%m»]Šú°O®©ýK„¹ÀUd¹Æ´#Ú-‰Î\¿ËÆ×½ÝuìO‘.®q2œÃÉEwà4… ?Š¨›†½¤HŒ^ÁA4ûx9FkYw3íæë‘)ÍäQRòA»·ãŽ¦ñ8`tkIG;þÇÿ _é(»ÓžÑKÇZ|Â¤Œ^FÖ!¥¬›ÓJŒÃ±ÓÞsù¸Öò9ÀK¶É ÏV>>ÃYÅ™Ë/B­3>a t’´Ä¿:>qùÑ—k$2ÎâoýFÀÛÈT”kß®.â†'vÊ±ù]¶=á²š”Äbw˜3OÅ+Â©©0¹±|¬Ôë®õ`]ÅåPõ(28q¦Í’•B_åÒ”¸W*cŸce»æïc
Üóà*ÌÑY<¿ÚRÕ§Ö	tœ“¯ÂKàòF§8?¹/òé¾XâLé¡Ôp2¡¾4›ÆLójSZö)ÿR„L[9}—w½—@iyq9ír<´Ë]Gª3“bYD.¬
°µ%S9÷™)_7
œ÷LžVòpŸŒàÿ~N*Ô“Û¡ÏtIÂ‚ôtbsÏWZr«%)5-uê¨Ç{röçpµÉTïÓ$%wÿŒ«Ý
[î8ÇÂê¢Ÿpùµç‘uÁ3yÓ ê¬©êüKÅWÊÍqö³=g„y1‘”ú8ÙO@c‡£ãäSÊ½Æä¼<ín*äÙeùxMY{SØ5Þ)|þ!í˜·¦ü—VïVƒ!:$¨ÓwPûý7y©o7‰ôjY	,£ƒ¢â-f1¥nìÔ
zf	|#·îºÖg’v.©ÜÞü›#ÅöìœJ©¢·´ ª•y•ÉåÖI¨ÐêºE·››,‡5’8Žgæé0L÷/L)¾)ô¦jÝ¦°qÚVž–Ú¶:»?›¥Éûðýb¡Q¿éì\<Øï4ì¥â‹Á`>Bçg?,"0pÅ!°¯¡P¬êÉuØ8AŽ<Ú¸ÒÙP’Ý÷RUÇár·w¹:šG”Q¨¦®ö?wÖswžÈ]Ï­Ž·8¢ ØÿûËôïA¸sQ³ÿºDÇÕ(ò×‚Wár—®Mî³Ëý3¬Ó½‡ô÷¶¢5¶ÇÄåŠ‘i«Ç<|sj{W/S‰˜×1nJþPv)LõÌg¾ô¼ôUN%E=àî	èáqð.ÕIÙ˜Ö›½íÞ[·Ñq™¬fÑŒìj¨…¢ÿ`Ç/c…«(‹.Dõ’ë£ÆN·=¿zP=DÖ7ðzpÔxA:Ý­Ý}rBzû[½}¨ng«Cö÷áž7
Û]G°­:—{—j¾÷ûh05êV÷M4l
SŠKÊûåË(Ž4šÁïÏpêŸpÜu»]jìÀ2þ<G?¯¥€ Mƒ›£F§ó ÝÛ÷®/_O./³p†pjø>iå'È6c6€ç ?øF¦g³›˜†h–ê€l°êZƒ¢>ÒÉHÀ¡˜Ìg¾è,^kD¸«¡×ƒ
.²$žS™}6KÆ-~ qªf¦ü¾HozÏ­ú1U+Â|n÷¡ãxAçÅÏg*
¯|_¼uÀÄá¡v*N¾-‡$G ñ<¾
Zå5÷	¦1š“cÔzÓc6 ×¡¥X
…ã¿Æp+FXÝ* hæ;£ÊVÎu™|ñÕkQ0Q>›Ë·—ŸlaiH…µ:™tn6%›]Åö|r™qË =Ž0àb1E‹fÖÿ5)¯F@‚ì¡5¿‚3v2`ó¥ä;9áèøUƒŒBÜMü›C÷w…T÷ç\[ð s•‚Yp´xã±îsÏLOIìAá¾Ehì(üÌEd8ª|ì\E©eêU±d~W(	^uKçðþbÑ/13]Qlx¹—g±ív»){~3õÄ›²7üµ;ü¶PñƒÝ w±¿%äÍ[»?^oÝDª|ÿ,$÷ƒ7žFX²¯€ðÏ³£Å½w€Y)_Ø÷xaðX9™‘)Ús¿€+÷wˆœAØÕG¿M°ì2_v¯}…®°ò˜N‘Xn¶`IÂëMW|o~ž†pfS}Á»>|l}½ Ü¾»e<ý‚Ûî3ÄƒuntRÉ×IÏ¢©[šèé†Fžçl;æ	R–ä´˜[zÁ¦gJ¶þá‡ýéuq¾v¦×8èh@¸ÀëÅ×_Rõ9l«£E“jÈ`$'×±¾y÷5gK/·wì!'åpŒ:OçÑs¸­9í–´[:üH£aßEÛn+òÇ”éÜË·jþë‘‘ÈŽEe¯Æíc‚TÍE~bÞnÖŽ¦øHí,²þ6„\¶yKù×Ï§µ…/wž.o):u»[¹¨’E·ÛÔ§éYé«Jç71+f‚ßðèäªRA­uˆí¡E±¶kž %“úY9Q,4^yú(¶ïh`¼JÀ+øìxÅsQmÕô¢µGÅ(³[…€ZÝ!ddìý£‡°ö[:‘7hmwÇêïÙž\%1w<À¢1@½‚óü²²›3† v/UÑ” ÷¶–ž``2
I€zˆ?…´±4)­ZG31ÅØƒC¡ÓÆ¹	¢Þ65„†ýù,º¢¹æ( ß ±I«F7X»Â
£î_'Cà¸z`–&q¦	½/ÊÒ­y	k*ä+ºŽÑiñ2îÐoèÑfÓ#Ô… ÊÂY~F9æõ‰ê-;’o”NˆC³K“1	*¨|ãÈJ6ÙÒ$»z¹Ã¢¾ÏÎÝ*e´Ã<o£r3sIÞdW^ê‹¦øÒN†e°IX–¡vû}MFtÃ¼†Ž…O¹‹œîpkZIAúßj!a¦ëèYz÷ã%¢fÁRz„¦ <¶—wÖ‹9°{Mî:¼&‘H¿gp‚  1ÈÛøSiÑ¬Žö­=‘½Y2l†µ¶n¬b–3BÐÄBµôÈY†%µ6cY‡‘U‚V½¸ûw±€&”9K“!œ¹YâVÝ×o|ÆÀgè35!rEæ„ÝÒnÂFfÃ'¦Ê3ô–¢Áíx¡jÚìþ§Ìÿ¿Ùyû9tù[îjEé²ñjêªÅÄ=·NŸV^§Ûfãc'üØëâîÇ³h“+` ÊLæKC´éT‡^Vl‘ÎO³hlÍù8‹ié\‹Qòé¬Hg• –í{[×ÒK–†€ãŠz‘' ûöÂƒ ™‘ÁÄ‹9—A"Ÿ¾ç.¦ÍºKŒ9ä¢ã‡‰rù¸ì¢ë¡Øcâáèé•§YÜTcàI{,Ó¬ ±,àyòû‚çó\Pöîµ]LÎ¯ß,„žßþ
­°õ=Lê>[Êåã\:€Š³pøùÌBWYúÉ®kü«ÃO?ýð#]ò»J”Ý:Æú<úÑµüªSÁŠÕ»pX$”½H¨4I0½<Èª“Y9d"4Z8T=~ß¦r\£CÅÂ¥Rq˜Ð‘õš¹RÈ'Ã-é0šP7¿dz@:;[$Å¦±qx	Ÿà\Ýâ¾t“-8¹ga
—Yñpì‘^ƒ€è=Ãv-.Ør·U/[7¾ÆN®£¬4µ#gÓ ÜÃnC5ï+7°ËhÙE»è½Ýû»û[t+G?`
ÌBáÑ‚ -Xc•óIë–È 83o`ða™¼õiÍòÆbÙ ÜIÃ±Ý"Œ?^ŸS]üŽ>ÜðbyVþ¤Ã‹ ¹³Eÿ×ÞÙÛ\Á‚¬˜õn‹l 5,(ÌÃë„¦ô+:ŒÃþû­ôÔy´Ð=þ8åîÁgÙ$ÇÉ$™ù9
œ2q/L(Í*GÉã•b÷ñ"ß7˜‡õ¢çá˜=HhXP{[âÛ4.ß…ßŠG,&ýqQì}s±;{îÝ{àS¬ÓË  ©+âœäA²_(óÊ\^L.6èrÿ2¸ì{QÐ/dý£uØ©\%±,ö£ Ÿ)÷.|so‹À°Vp¥\é9µ@ïº3nç$bigŠ¾NËe±Ó;dh_Ø¬DöyzLØÜ·wÇgE”ùst@¡ð,ƒ$#9›Þð^ ¸k‰ô·+z8sõOŽ@ºÌ§1bÉIÙ³fág¤8àvT|ðdcIM.€>ŠQ˜5Ëjd`W!¬ÝŸß åY–Q«¢ŠŸ‹©âmØëQ (ä¼€àˆï/íó2x }ñÕ€‹Ï.éI+r<»\âaäVÝ›êx¾}óv³“$Õ:®åÀœ±L
¨Úc !xÓy÷LcËzÕ®÷B­÷¢n½ÑxŠÝ}þâìäôõ·üöÕã'¯è<lÃoÀÙ%8¶lhi]ÆPù¢ÀGÚÕ.Ï=-±E[jXºBâ.e¾ÉoòaÜ¢^lû·ž+9ÃÚ“5•Ð9ø„IÕ×9:˜GE‹ÊÆ]ü(ˆ®ã.v*Î˜sœå¼£NÚyŸ¬Ü|-ß‡Ò¨ÃGÅA
åW ÐÍzêðh(†‰æI<K˜ƒ€*€	¥évo§Aü‹|q÷ïƒH*´ð›h=Šm(P¼¹«·¬—„Î­…å¨yL.hå<,ÌfAæà¥z`h{]|#TSE]Úþé‚hù&´k—ëyyýªkBÙ]žG«Ãi×|`Á×ŸsuÔ~—Í®‚snmÑB¢}ºö6Ý[ØTErœ	º}B­/}“/æ?eòo™%™¥Õ+’þ\e»Rú¸Ê²*—'Ï’ÛÚ‘Öa~·#/A~{IÄiÄDTÒYñÄ³¯PóÍÚî4ÛÒŒÎ+=Š¤ÞÊ`K%ßª‰$‘V¿Lœµ±Lêe¥ÏÌ×¤»ã|ùšâOù$^¾¦z§¥r.W“!›|ÇkæCÖå¸†M+d>ÃÊœ¡AB2äA«™2Ù’ZÌ$/û$KFÙuˆÙÖP¢ãilÊ™lôß3g26äId?Š©oaÄi°¨>'ïÞ›´oB¿pêªœ¦MZ\ìê6Vºy¤ö»úwwÏŽù§ê3B–\xNüV}¶ßÅG];‹<ÁBSÕ{Ö÷”ÁÑÝ×¾ÿXêŒrCûN}?TŸ˜ÆÁÍù¼ß³ì9àÍ•;ê;Û+ÒTqS¯+½}Yæ/¸jju“Íç6èÏ˜nSU®“d®LO‹uöÓÓ>OZ•Äó~"±àŸ>ºTŸêÐZÎ—ß¥1-Ÿ}4Ÿ§i8 gî@<Ä»Ú„O”V¥ÜÓ¾W`ÂÓ7ŠoÚgàsú´ð½ú|Ÿ&•ÿ&Â|”ŸMÏû&ßïVßáîé(‰úÐúò³¶íßJ_5sœk±JufÌe³iÅŒZy>™<ã™fÑ'(™ ôRfŸ-ïéÞ|r}½øjZ½ù4MÆç@V¨×‘“ÈbxdTt:Zˆß´£öZz\¹QÿÔ^õ¦ ƒ0ÈÉúà¢È_Üùë8šÌFÑ` ÍÝø?ŒŸF“õŠ-+ŸÁ¯ò¥ö;‚=7?’Z‡8ý|—úÇwStc)Y¸£EÝLú¤Ì®2£g–ÞXmÂsZ*]Y*¼ÝnkVÛ)E”‚šÀÂõLWS¨Æ 0S˜®&¬Ù+jY‚©ùj€ß¹I—ó~z(uè€6MV24èÎ³05ûµ›‡¯àCåÃ÷8é7øÿ‹-²!Šè[††nnY¼ÕÑ6>jNVsK€+èH³'óÁ“ÔÒQ¶’8l‡@¼Ófã	þÃúM†Bkà<$´¼ƒÆ)
66A¯ÕéEóÆ†¦VÚ[ø~!Ý†³[¦OÆS­¸r&¯Êš€0YBÇŠ¯yr4|j?Ø÷æ?ø;?ÄRCå0{Â™‘Pâ¬?Î8ù–éxžY|Æ8-ÿ6kûè¹mÏkê>zµdòd€oÎí”³ZÍ6ÿVÉ•ÂÓŒïÝ`¿±/(RåOWSBéÜ¿’ÉcèÎ’oÕú¼T¹oUìÐ8Ÿ3tcÔŠÄî™ß9…ù–Éï•÷õïÆI0xr¡È	oŽ‚É ŸFqÈ~±È…ÀZ~ý´²ª8eÙë–ãø3@µ¤wøw[MgŸ–UÐtÒ‡Í÷ø©™ü\äÇá>g·>‘zås#ÇQfËœüUÏóã¦–äzã-¹ÞøI®g¼µô´ðµò‘ÖP('C
“?'Ç`-¢rŸãˆŽs"ÁUøÉä2JÇÏÇÓ$R«ÜþªŒþ!-ChK’Ú ü-´3HÉëhJ¾¹½j¤aå²À6qVP¤N²PŸ¬Q¡éÔ;\ábP²BÛþdÈj@ªÚœÌãXzNš¨’õU_)òÀSµÈ*Æ$… ?›±0\˜‡FÂö,ÆÍ
ûâªFY%4KµÞ­jJ9·õFL	>çˆc}Ôà3”jç¨Lî÷•©oe©â"õÀÐéÆJ€eEMs†ËêÏ06°Ù‚Ë²®’þÝàÂî~ä“AÈÂPÎ
Iá)òPÙwŠFC§z9¤žîv
EÉ{`Sg©Ë•r¸M‹«TÂ›*‘¯¤öÈÁ‚z"¤–)ÓäeçÁA2ŸÑØW
Än¡3f«+~¼Î¶»ù›z4ê¾TõD£ ›_QŒKé<bß¬Múñ|fÍÊÂÚ±µr=¾ÏÒDãÇÛ«!’’Ð(ù£`ŸÂÚ¥ÒK §ŽxÊF Y½&Q7žÚÚ&;q+÷©_ŸSÖë¬­Ú·¬`WUÙifˆ:Éâ3÷&S<¡˜(>y›Ï7ýrÌþ9ÜfU+ÕÁòæe²‚tìÇôšy]_Öá6Û+êýEuy‰sŠóÑüªÞâ„C±ZæWXæÆæ¦Œ2šLçÕ}œ"‰Ðùp‚ÜG@DÂô¨ñ³E—˜©$*áÓ½]¥Ômí§&%PŠ¦Í|õvNUâµ÷‹ˆï=·?³ê[ÑÄRÌZ»ÏkDÔŠcž²C*JtåÆá6eò%‡< 
ø|…Kì ›ÅÃ7öîe8¼û+°|d›ðòlŒeLtusÅÀ 1!›¿cd]Îƒø
Zv‚
d1ÿ34ñ/¹àÆKyÛôÓ119Û,îæI9°X‚pe!”O½¾€kvFddÚäy–%äD&‰¨:‡IýÀXnÂðD‹„iaÊP™Ü²$kÿdf)ý·f†ô«9"ý/Ÿ5[t®m2ç”~,ÏjùÂ	o¹#ÃªªÃ"™–¬àŒ ”þS1Kòú‹dœäáþI¸&uË[Y&øLø&›îŒ­pL§¬*+/%ê(sf*ë'ñ(ä¹Ñ3¦$c£%´…™úªï vœk­£Œí3îÇÀWq+¹Aý5ÎuüÑ|EW+¬|fû¥†¢•»(Ê*-¿t°`Ž‹·r§î(ƒ‡Cá«ÀEwÏÒ¬q‡V–‡‡ÞP£NlQç‡˜68B}PZ
`ü|õ#T›‰“'–X²ŸÆÎ‹¡M;Hk¿’C¾#ôœ3®ø{¥©õ¸Y5àcÝÜ^·Fš€ÃGi^Ò(9î.u¯~:·—Nl2­™
¦fb/é©£±ØZÏõTBaÅ0<3ª[\¼déŸ"½Ù”O	#ŸÓ×Òî!°kÓ€Zâ­|I«0(ü¡Ú4üó<JÃòCÅ˜R¥ûIúi  ÄcùuNãÏ2®Æ:;…ú…/v¤­5W9¾âZÞŸûª{4b43uWXvZ#Ÿ´‰›â8ï`¨yâ¬ë±íÍÖu¯S_Ôœ…ãˆÎ%˜$3\}É‡pà^uª%€a¬8‘=äúý7¦Šˆ”Œý€ß JOßßðô+]9±„9O•-¹Íÿû?ÿëGCp~>Rãïïþ2éÏã„œÌA4¸ûÉE'ßJê,¨,Â}‹‡Lr]ÄQ2û”ÒU´Fi8ÄŒ­ø0]†’®¬iÀn§HŠÇÎéá-ÝlW³ª'ûÛ½Wë…ý÷§QÚÃUv¼IPñÜÜ>€<ªRs©$…úï~¤Í@™~fáŸÐý²Ãá‚=Î»´ŽhcËJÞø°ˆ»¸ÓùwN¿¿ŠVõ¬ªÔ4å¢B.S›;jøÁ¨sbKK+iS=¾â˜ªÅ¤Å`Ñ#S‰E¤ù –á\ŒH-f‘©P=ÑÎà™ó‘²ô4%Û@tŒ„š¶F¦_Æ²„Ô7•¤(
H4ƒ}P*ïî?™WL€.Z8FœÁc;WûH.’kSÿÄ¤Õ,\Æô dr53f‚ãx0T?ê+jUiÇlQ•Ò,õpç]º#ªÚ*4p{@›[ôFˆLÐhTŠ6¼µµÁð‹¤Ñ ¢=áøÐÚ#£2©æ(S²µrÞE! ƒ9zˆÂÆYù¨	3`¦º÷Å-ß³~ðÒˆáUoã VN{²ïbd™~êlp’Zü¢œáÜd‘>iFd*ÉyŠÖY “ìÎÎ¡Ó¥‘RüæxRâŠÒ±E,‘—Ž:š=´#Ý¸”ÿ?   ÿÿì}ÛnÜH–àû|E”°°RÝºÛ®²5.y³$¹J=¶¤VÊîj5T&%±Ì$³I¦$— `ßû¶Àb±/ûR³‹n ~`^õ'ý%{NÜFÉL¦.¶åR&Æåœç~+iV´kqƒv”®qòS"ÃiKy§M«®zg_S±×5›`´ª-³X¥ª9ó·˜+¡`.ª`›¹1ÕPKæƒŽî{K•æÔÈ¦ÝÁ ‹ù…z6…+K¯@ÕöâsOˆê”jq_qéÚal˜Ä%—÷Kã@®®k¹âZÏ¾	=ÃÁ%1øÝôöö•@å}RX
“óˆY}ÿ•ŸOk9ñ~õâ^Œ€^—üG£ÆsÏ.ÃÒ]jxwÁ¿5&3|]"¹…uÇ¢o»Ûûd§÷ç÷;‡ÛûJö]îùž9¨TgÃU]ãGù¾vª®¼Xu–Ôòë_*¥â¸"ÄRòÏVürk&¯ÛK*›KF—„É*>F0>¹V¦P4¦0%¯ÜÜæšý<-í°Äs¥ÃlËåËªÚkÐ9õSid®B‡Þ¯8=†p˜]9	N¹Ó‚•§ Ãª!Qå¢­‰]•¸Þ²Ê›EYÝÍ¯Wp–RÑh¡RÌ¹¦†À%•Þ*’†Òñ"QzÞÐ¼®­bs~ƒS«/ZØ	UÒÐtX3O¸¬gØÔNùÆÃ„À Gâ„åZÐüàÙÑÈöŠårƒô@­v^ªm"±CÜr[e"Cæ\WRÐÁïÜæÖÞÁŸ* ¸uóúÑèdfi6üíà4ÈÒŽ~¡çõ™Ù=,ÿé/8–rWÕÅ¸ªóöJípAü›(5+!ÿø?ÿ“|¸ùì.ïV¢·F5›s“Ù®'ŸØÿ&»Ñy;S³Þ³¢ »]d³à±°ºº¼Êþ­À¿Õ%Wò¡wù–4MûSƒ,WdF¯Cµ«ˆt±V{:4[Ñã(vfPŸ³PˆëÆì
x$JÖu¸OåñÕÌÃ@Eþ—h«|÷F1¼öbl¥·J,Ìqòe1ýúêœ;ŠoÖì:G²ŠYà‡4…†¶gìtƒ—.b@¡¹a´v@ö<çÍ 8Saq«ûæ6˜ÎõÜÂT¼¿‹‰hm©7aœ`Ú”¬ 3<ÂÀKìÀxÿXXu‚¹- cÛ÷NöÆt[ZGÑó±´3…æw7OWº<ßý±ŸdãäQÒ>%ËÕçÐÛZ3ú),³XâÖÁ´Øýì`ÕMÝïB¹Õ<ŽÊÇ>>«#âT0Ä@5>sÐ¢²^fŽX9ö©C—¸QÜû4ˆžcoˆŽ8ÊL7xîþ3Ï6àHŠW×¡Á p~ÎZîX|¬&óIòØgÚ°<ö1Ù_‹ý7ÑãÏZõø¤­PÞ$^¯¦ æ¥N;ÉÇXÂùøÄfsìón›¶I7¾Æ§ð4Ñ‰ìSˆ÷SP¾Á¡ìL¸Ù4ð}døŸ2"WìŸ|ª	•ûPEarFD¶°õK•i¤ÝÞ=ÜÙ=Úß ´
l/’tŽðoþëú#šÌ‹gÌjd­}þh­5låíYk×íN¾×ä	y÷½Pdgø¼Lµ9oÅ u²IZü„*Ÿ®V?–kÕ²¶Ô]Áê‹´Æãu‚ð&¸tË^_ Á ½_/\®ÖjçŸ§êjËÇ¡—<ZñcÒŽ²•j_7Êúý’´M;C
yE.äF¦>œ}ìð3=éZµ|"cŸmœÂµ.›áïQ˜½ðÔ.äâºâ¾õOùöÄÆÕ’ªø¬¢#±ƒíº+r$uÖäQ``“KgvZvêTûž‰r‚Y»Žo³s{;xÓ^u'¸ÑIÿ]n3ZÛ¢ÍšÛ9N8Ð¼q÷Yá%w»>rõ<äªœÈª¼ãnÑ7®®g\å”l6ˆæÁ~ÓÒOÍN:Ã9üàT/¸g6zUÓî¾xÀÍ©0¢xÅU"µ(ïÒ=î®œãfãgMB¬$ížÞz÷ÒYa¥ž¯\[â†óR»a“J)_[¥Údª`ÊÂ»´¸J¢Ü4ÄXn·o©•3ž*æR›cé"ÁT®4ÃÔ×Edê‹× *³_œ¼;2Ó„)EY<­–Î–~üzõüì'‚‘Ë'a|w0~‰Œ’RžJ¡l8¨¡üÖ{‡¤QÏ˜ÙgºìÛ…ÔÙy‡
¾áöd®š—+®¢ì$ÛD®Kx	$[ë|‚H:{]9i˜-‡‚agÛ¿&ÌÊ¢ÒÎ˜ÂLmj©AÆŽÂ’Y¨
¯–„$pxR’KœpIGu'1…T„A•Xäœƒáª)Uº…w™”ŒMT TCìiCä‘"Eæ±H:-¡œe‹µ8cGNhFR)$ZqyJßDS!ÍLqÏËÉèô™øHœ$Ôgª³o:NŠZÔ&f™ìæ³‡Ì3UXÎ&ÚÖ³;„‡z^$_$+]Ïadbèhê2r‡@RåûñE‚G¥›ÇäGJ}G;
a•UšÁ£ bYf ¢ï™ƒÈ½´™³ÎLB÷èD>W)	(mk/‡³¿û!‹ÃJ>lQ&ðð%ñ*#d¥òIáíà`µT^…ƒÂ¹ów*œ«02ƒSé¾¡©ñ.e3hwz‘C6c˜‚r:‘ÙÄ0Xº.Ž°rBË‚˜xÎMaþB¯~çG©Ò¯ÝhÌ¿mû'AñÜ›7A‘*ÆÛÒ  Ž%!¾O¼l £³=êV{žkn“ý%§«XX"]hÐÎÜ&ý|ý¼qù2Ìmþã÷ÿAòß¤ó!È¼ðæwXq²BvÃ`¿ž£w[PñíaÞŸÜü(ö±él³™PÄ‰ÆÃc?12Ã úvnÍÂ|;÷òåËêê¡—äéê"yö|‘|_éñ½£ˆ Â#ÛŒàÆ±|ùfô·´Ä5i•PÈ®7°ZKêïF™v‚,’µUZâÞ|¿Î’^„%E9ªÛ÷ÓôË=NR¶x–tÙnDXïÆƒ„½ÏQmlâ#£Ëm—ÑW¶¡äm K˜5 Ûr$s›ò«ì°7NGp¦¸:l@©gUÐv Ð<uj–µ½dµ˜Œ>QJó|žÞ‰!w"ÀË},àT¾oó‡jP·œ²…eK`º; ßµð­¡B-_¢TY¢´—ÎñíøW•éã…>VâE2-^P.+ScÙª®jéðÇä§Þ¸ ÝC
Õ)÷q_Å^šuæ@|„’›ßá[?’tLQã«¹E2—²nç G­KÜqê'¦\9YòÉ˜È‚•$fÀ¼÷;üw¼Hæ¹Ûdú³BkÒùEÇâb’œ¾—õÏ/ÓËpÿãÐ_ö“$N`Þø‡½ˆ¯z¢’“fÝãñ`æ\Þ#ü ¥„þ› ³âÅ‰O{ëÀQHK£"Ý å@þr¸{´³HþÕ8Ÿ•ÿt%æsý¯¦	ý“ëw­ÉßÕ£âÜÃúÈVb¯Ü*Ð©r98Y%KSWù°X~’ù¿zjÕ7ŸSš¼º›ºÄõ*½½Ô+½)®X¢†Íëâ¡½‡áI9nþF‚ÕOÒ›ßÏýP­W¨7Æê'ýí”i#?˜]ZŽ:Š·iÀÒáÌ8áw£Á*«c»ŽÚja9‹ßÆ~²å¥~‡fÊšÿçÒ+€Ñøè’ø8ô‡é€¬ü¶»íÉÊX#¯e£e|Z;®€˜ŽÔZmUÛˆzq8îÇ¾mD)¿-”:”62¢ÄÏÆI¤©tËÇ%~´#“ÖWÄ”œßÉj—%ÑF;6++¾ógÄùÆžÊ¡ª¸ûúƒ&Dg£Ô9C+Ê³â««¾°ø¾*gXgœ©š‰“ª°RY’–4Q˜Z€T¯ÿ¨aº:úÆ›²f?/™Ç£1ÑƒA‹ äÇ©ÃˆãJèA”ê°í!ÕúÖRdN±â®V¨Ó8!‘OÐ
lâcIºÊSÕ…Ì­µ‹2ÈŠÔÆCl•‰à®¼0ìêÑüvX¥À•Ý®ÛŽÖc+Bl^ñm®W›3ÛÙ3ÓÙÒ6Áfñë˜Î”M²Š¥‰¸Ã-»í¯T!néé2ÞöP[T®…Ì.#+^¨gÆZ ö_¡µåçö¸Ó‘™ï°—VV^òÂB¨x×Ý¡#$¶§¡/*þ=ÇŠ¤›ùcK¼˜pjm¡‰bYÏ‹©ÛGEY"F=GqšŽ÷¥–\6ìúµä:¨dÅ¡ýf`åšüta:D";=°5ˆöi™¤7\&;º†±£Ê^–+t©Ÿn,X2±¸c)ŒrËW‹õ"Ã	ü—úã4g«àŸÌ‰ÀHgÐf	Â C/B0Òj_‰?ŒÏlPqšULÝUj·cð±h¿ÓðQÜXs`¤hóBÅÉ¶1R©ìÂI¥g¡'gèB¦Aòxir±Í­À\©Zq\(œ0„H9%xåÈp6Š>e38q¡Ý?'Pœ6a¨t­NMPM¸¢å2ÂõBGt[[ìbqj"s]ÕÏjØ¼ŸªÔç9]š°J–¾ÄÎGÛ‰üþ}­'VÑGØéãT9*…q×º,®Nš›„m§>˜15â®r®¶#}ê2ÙÌD¢šÜü†SÅ×FöáE§qü
k!Ô(‹…*òðSÒÝE¤¬§c¬U$ÁQÈŸ-Ñð¡„&É 5©@¸
h€eÐ÷×"#0[[Y³–ûà”îÝ¦ØÐÔv1¡ðP2 ùý(ãÓÀ[$€•ýÎçžwó‘ìÆ¼êâ"9Dø'ì'ðõå®^Ä£2“mIöŠUK%¶ÁtÎÆ|ø€ESÊã¥W`1·M´©UÂYe‡Ó‹£ø^_$ùV!WRªüm4T{J7__ÂÉ*ÞÑó;&•”©‡b»; Ó»Q uåä™ÒjJnÝ¤úI§$™²Vö EZ	gô`?
?µOD•ÌmÕD½\ûã$ ·º¬ÄJƒAÎTQ“™v Yú‰÷x¼Õ=ÞFJÃP{âùÒ!¶g ‹ª3Ÿ¹÷GS’7>Â¸Êx¶WqlùÁeÐãª-©oØbhsÃåÐxZãÐ$ˆCoéÐ¨0õåÙÄ¾7w¥-Ðõœ0íU7å)ýƒ$J¥) ¬1ðƒK&lø\?(© 6 pzÍ$E"…Ó“ÁSýL=A¢%0G·ñdçÆàp˜Ôíe`GÈ¤	Õ6¿ÒYÎµöt”Ý0Ôj…t>´m±¢=¶¢·
íCÝgÌ¥¡±@~˜S£	e¿B6c©.Ð©ûêfÔ„z³ºŽ›g´ÃŽ^u©ÎiƒšóÆú9“	Ð¥gzñ]ÕVÿUI±TÒV½2è¯Tó¿7òN)šëÊV ¸c/êÓ€îÃF—,Ò‰%úpfÓ!Vo8·¢ïUlJŒ,h1ˆOÎ WÌïÃ†–øé(†ù `ÀC¶ž<T/KÍ<t –kPë1©,Y´©®ÈË¤4fUgÀ0ø'Åçu]‡gÙV˜~µ2.sVµý([ZS V¢«`­+½/è4Õ¹ÿF~¥Ž‚0»“ ÏÀ"‰ñ¨âLÕ8Q
uuig*”œIäÌú?)\ñFPÄo¼Ú’ì­ö-oÑ½"u«Íœ"·öì½•·…©™S®|ˆ—)Ó ŠÌöÐ2á1
&´«Ò5ÌOEšãqQ/Z¤;‘NïRÌgF™\Æ·ùL|12¾
ñ÷WÀ÷éy”î§•î9F»D{AÚjËõâ{/Ô3SY“#°`Ôófqþåcêðµ¯g·+9âu²£{c©óÅ'§^‡%Â¡ MÅ1AµT†ªd»Cg$v˜±Šó‰~HO¡>/3^[¦¾ãÜ#%›ÀwèÓ-%?ŒAÒB‹ì› žÂ’x‘|ð#Øé*’ÖÌÞ&Ÿ¸{¢VÓK<gÖX])x;ÙC[4®<¶ªPr4ÜZrBPMÖWÔíoï÷„šÊNz…òd:É#Ú0Ý¦¥ÉG»3•AIcÃçÍü·'ûêÔ
NgÂGMú¤-•/¯+xgá{cx¥­HüQŸüo”å•^r˜+=9§—]ú™i0³51PÜ€™¥éDxËŽ¼!|Í oùØñv»+ñVUmå€å}xÛU§æÂ[ï‹Ç[áh<1Úö¤þ¤u¼c{DÛ"Ú*ªçÏoUµú#â:7°›u‰¢Çëhžùý Uß^ÃX+QYÕV<A·kñ¼[Fú ×K(ºèžá'e°[3‚Ö52D»Ê£iYÔéµÂZsÌ1€µ¹éB÷œÇË»—¨u_›v(ÝËÈ[Ñ»B§&ížW)+K†uR0mš0ÍÆ,Ü‹x×ZÖ‹‘ žsÐ=øŸGEVEýó¸îÛƒ8ñƒCµèñJMÁ™Ñ ¿˜[Ü1¬MØinÃ¶¤µ˜G`Ó±4…³me.`û*H·yÆåmø²IÁPé¦&(^™ÇðšðÙDê O+tq 5–2ös}w Nóo©¥©+T0R—ÄQ|šxC/,©`t– ÍØí{ð@8|0_žY<9¶N%•+§] ¨šÌë$j¹•†£8ÉxúŸºÛ¸Kâ£ƒÕç§}»ÛXùDq+Ê£A 
¿YÆ*“y£ÏÚS2Zz¦E½
n1<µÀV–ˆ»€VäÿWèœ»	0L¡\yU¹s“K™áºÎ=ëÖiˆ«>¸‰÷,.z¦ß¹ÍÏeN³a·KÄÕ}áy!$†é…~œ$ '1¢gâ=­I×œ`P* HÖWn'Ý=ûûˆ•«•Ú¾ŸF­ƒ¹ì½9‘¸Á‚+e¨[ÂokWõx,)·ÀN7Á\S…”¬KÃôc0a»Áx}Ü6óè¬ðw¯–XˆV³[b.ÍiKì•—X6ü¼–8ùgµÆB¤­Zâ¼Ý4+l"8ê[hÆ9xOI…‰S‰VJÕE³º¹”t‡îyÆé
¦«ÃÅ·dF*ƒ–Z}º^.³Ê Fè\ÌðBï®‰óŽ×o7øñÒJÞåÁ¨ëÈ—‘fØéÀ÷’fèô°H‚ÂR/Ø³îÐiÑ„;ÁuaztY¨ú¼¶ç>¨[1Áp±ðhKúˆóÀ¿xçeã$È>mÃ„µ™½=æ¡Ü'ë N^ŸæŒØí~†Ÿ±[/ùÜ¨©M(‘Äÿ!‡ÀØƒ"_F¡ØÂ”ØÀud×Õ¡H´Pkœ4bm…g^õF^ò1Dýäm¹IJxQ0D::ÃÍÕeMÎú¿-ìâ‹=Ê¯ÿ}–.€²¹M#XhûýŒ†ú–Y¿³geR²·Ã÷Ä3r¬$ô½¾7NÏ2ÃâTN©œãäìYsŽ“Ã¡~¾ Pž±R½Îfú'å×:LNo… þ-8úh‚åGvÄÂàôæï˜³“tz‹ Wª€-Vbå?ô.á¨¨ŠàPwâutå©Ò“¢*Æ»ÊŸ	üäÈ¦¿4à?ì àI0æ`|ó7’ø}/äI{twþÔ/z,`~E„R‘?äÁ¨CŠÇœ›É5I¬I¥~B-ˆ.ûÉ¤GPÇè®+<Pæ1>Ck…èåý…zç“¹*}þ–×dŽúú
Fi
Y(Œ¼vÝêAfÉ5m£!ùt¡AG°ö˜ÓTËü{M|O©ƒñh v æ®ÆØ-{wÊ/kšàK¾•^B—šù7[t²_¦ƒëkîçL£x´V£:møÎÊ6S:L7u$®.6ð&©#:Ènb[­‡Ti@& À©øôÆ}CóöÖpàµ%çÎDÜùV+%`JÀXÜ+½]¸èSv¯` <FÄ¨…•1óƒ Ígö,ŽN	ˆ8Ù¦A¹ Qg†‘ˆg%æÕ'	zÄ˜›š{9°ÌË#L¸gAã„ØÔÎ±ÐXþ“fI¦Ié+FIWó@ysçŠ,//«c)KóNç£ø‹\[hJRZ¾v­_2°´z¾ViM¡1©=[êIZ§MÐZÍ‰YÛØªQ×Gd<Æ¥"sÎgD›,˜4”£&kq	&ü@ôU»´ág;-h\BZK;—Õ 7aydd~ d·~Œlü³\3!° Êï)kC‰,aÉß‘öi1©ŽâxaŒæü„+c%Œ_•O''4'…©z¶ƒÅBæÌŒœ|Ýèm;”UÍ„-ßêÌ…}åÚ{Ya€âÌ&1lò{§a|ì…{"áÿàî?}ˆ.ü"1ì¦©¼#%îâ]] JÞGì¯s²Á~o²oø¾&+9®kgå¡ûu ¯£d/Ä$Û)€OSt| çÀÙÒú³z±¶ž”°EÎUÈv%…4J"õBÕª°\ôVªžU®’5ƒhï:EÓ˜t˜ÀÍ¿£&ÈVÐV°¢ ­ ­hª{ï@°.Ø¤q[ “Æd.¨ýx_ÀF$XÕ¥:?ø	ZøQuµ5N³xHÓiYaé3âC=!{ð„pér€þ5p°5˜wS,-LK(T5Ùp4yMæ÷ÇYÏÓJAVÎ&SK,GÆ‹¶E={{}®¶êÊl6@²E¸ðËTiE×ÂÆMÉ††€3›%ëE±$¡”„Ð–ûvý>ÈÑÀ]ïPëô{Õ¹Mú‡t¶â!0uÆwtb€ZÎÆÑô° *ô/êæi×óÌ0‹—­fØã…-%×El²?MäêÒt)ÈÍmÒ?žÕV(åÊŽM
D#ÀÙ>yB¾š é©zÜÚ©ºÒ&÷û\HOP¯Ä):êß9eeæ!/2dÐS¤#XNtÆÌswsW™TE©UÕ©ñûP“^R´eGû±ÈœÉç;ñªQµWC½Öâ,Ì¦”#šKÎfRâw§2(‰7sûÉ°¨ó°jUb·îÆ¦$§j¶(±Û–z¯©5‰?\´%Ù»šÜ’¤¬5µ#‰ÍÍIy†8¶µØ–ò¡I“µ)×9ÎïÔ.¨ÄÜ³Ë°žgIÉ74<ÝP~®Sg¯5¾W³/]­ü¼ínï“ÞŸßïnïo´0ÿàî jËfþÖXÑPK¼Þ¼„ÎhKß›E«üœâm°ò|•z8×p:Xy±êx‡ÅëB8 Q"ÂÇOÔŸÅæo`I
§Î1ú•±A×q;48VúÍ\,='gK´êô¹æwS˜â0·¹fsX$bŸèû[^Á’ÿ~ÊÍªÀ2*`É Ò´a6ozÇ%å¼y”•ÎVg.ò£ÇQõÐ!é±QÞ{Ê¸!ŠÖæNÊ¬Ûü2Â¯“?“ŒÒ ¹ILÚç,ÃŽÁÙˆXÌ%Í:#Ù-ÁžwºHd¯º
ÉTg?æp)‚+!Áh© ?\5I Z«Xþs“~©Q
ÑVÑÎå³{¶ w„-=Qt¹È÷ˆ9s¨Ÿ¶7Æ›µGÚóâ–‚¯Ã§ž‚g”‚ÅÑ€•+N…íökáÛñ£a´µµè©Èîh)×hƒ½ÞfÈÃbÌƒ¹ÐâË9KµA:Nk~~ÁÕ+Ao|JÐÔÎù­Ÿ©ÂÐß£=\ãuh¡PßA-ÅNé%Ö“µ#›eÒ(á§9ý$xÄ2).Ô!¥¥RØo.`¦ÖNÓA³ ê]g‰¬+à¶xžOJ”-@;3žðŸb¤®©tˆÆM<˜3/ÝóOÑÔ?DïÇê·¦!DýŽÒ½x²&¸t¦dOâO|Òñ ¡Ï,++Á§ŒW€|ê|œ¾DqNÄC½àK[ô6âÂBÍ×ˆù1ÉlNšyÙyñÞ— £¯i
5õ£äZvÑŠ…L³È‘ &KÁÖP²üçæ•ºn×d‰âš‰8Ý&Ëa+Íüð€£›ß³qˆÄ4»9«?JLü˜yžåç‚ë)%U}ã…gT!„/KbZ¸ƒËã¿Ž°NÅó»%™#çê‰%m"~îÎrÏ³„Qé8Ìâ$˜1Àb‘oçKI—RÎù¤K'°Æ¶þÌ@|w¯»wóoÝÞ"Ùévnþû÷»]£+û”á;gñZ‚mÆVáz9‹ß#ððòM ÁÙË<Üi,›]“¥TÖÈ¥eyH®qýGÞuÿeçhwï{ò„|ØÙÛnòìþÁÎa÷æßnþ×Nž>8ÜßÚéõötp¸³õþž ?¼×ÝkòäÑÎÖÞþÛ}XÔÂÙ—ß‘æ§´ñâ}VÄªÁ…÷^kˆÒÌ²êL
JÙu»RCcáeEéŽnþÖ€I·hd˜qŠ·¤²°ãr¸êáç^ŸgÐ3JÁ¯ü¾ûø(ºô©û:ðDÊ:KmgüX•JEƒ™²7<BJŠ:\’X¹$ƒÛÌ.Ï	˜»¡ÐcŸ)h>›4ÿ<ö‚”x²~Ú¹?¤iž0äSÙÖOU@-%íoZP¬¼hÌ…!Œ¡Û»‡;»GûÔÄîÿÂÐw‘ ís¡Ñð­‰AôùC1ˆzè®ù0Na6TãœÏ¢p?µŸÀêV#C˜x¿Æå0Ç	ìP0‡žú6ô.„Èƒ$ˆYÔ÷£õFª Ä’PröîæoƒÀkO!:oÙþÙ3½Ã$Üº@]z¢ó›Ûd>ü\Â³ôûÑ‡eÜÌ¼lìâ¾,ŒHérPl8 N
)p[øÀºnÄ 1„}kÖ;CÒŽ‘¥`ßT5ìH&§„nòï­cÊ­pÞn)Ž;¶*ËÑ«·Aší#èúQey½œý‡ƒµUE‚=Qo§µ\JI«kvjÂ%³‹=iEx~0DèjÍÀuÈ‚«sÿ†Wí¤Ë°ò|ÕA: ÉŒGa•¤k‹d}‘<m ÙÆt÷Ðe'u¯§¦Hc‹×ìuÁjÎnþ‘¬-ÐÀ‘šÒ‡ÍÂªÆŒ¼$õw#½hÛ"Y[µÚÝäŽŽuƒ {x‚A3ù¤70—ÔÙòÐ»ì¬.Ò1,Á¬¬É
jÊBæüA/éƒ†–ŸÛÜ¦ãÁDÀ'¦ÜAÏ{Ô€YÊ1é8f¢…Àáû)-[mÌ»?±éð>òÍ(’Ýèæw 7S]Ð»¹™å2¨ÈÞD·jÉ»ý¼,|ÈÞÃG 3Ý`°`­Cô©ƒ»ñÁx©žY€ÞgD0¹ßÌÍoç~øèVQÏæ¬®Y}HØSæõk”~?/2›ãð#¡5ÝLhóuC8»*|+ÄÖxQ»Š&‚ž/8còÖ'~æ—m³Ðê›èzM­þhœŒBÿöÔúßù \\¬ZëóáX$û|aŸs‹E¯_­<2ó„É‰7bJ
”}LÓ8gY9ô¬{0ô2èTB/U(×=æC>ã¡ ™UÝ¤·™3Ø'ŸâU?²¹ŸZêŠ¬á!y©zçDvMË•¬Q·´®z²4ŠÚéÙÒKž‰ÇXãz¯é5Bl^óÖ¹[ÚŒØäjióaEKZHëdžz×ÊDöÔêÎË2›%š|†Ùæ7Ô6ƒ‚qË,‡Í;œh®³(¾v–pAŒ4ŽÁ¤é±{â9ó²ÔÌP.î~¡€.§_ë%/£¼«wÌûiêú½¿ŽaAªŸ®Qw4º7àï¥Ÿ¢>1Ìç‘:ö2c9L°0NÏlúE–Å•=ýä	™ß‹E¢ñ8š'AD.°$Ý…=u –hƒ÷{^µ þ:îø i
W:Ž	íŠ¦ü8dÉü5‰ûxx”wæbà6¼ó8Y$gÞqÒ¬g)&LáÈè’ay¹1‰@î8NX6´ÄïûÇ€xwÂ .Œ-dº<çŠõ°Ü±]¯IÊpÃ6ä¦6RØ~î„—¦Šˆ)li™†)7 	ë½$áÒUÒ/\šF´Ë¦ºêXY¥¢`V:]Ì·q`–¬=eF:¯ú0Ê Rø"gõ$—Ð_%ö¿œÒ?RŽ„Àaƒ,`Ö(&H½Ò×N5€4ýÛé«UÐGt=[zAŠò"§µ|_­¬ÛÐ Ø$}“¬/Àb;ðÒ..‚ÅÀ‰Ÿ*ÉÞFEeÿÒŠm¯t‹üÕÚõæ<²­®2ç—_¿Þ\ÇÇÓ	Ÿz½ùtšçŸ_o>Ÿæùo®7¿ið¼ËÁ¡É³„ÅÞ®oÁ“^Fu5Ûp5 ž¢ÃD'²"ÙH[vG9Bªÿ³.´¢™n2Ô ­(•òçöa—lü¿Ä^üŸ}\¿e¬ÄÉŸ?ñì¥øLB]Ä9•†“Þ}ÄµâcÇ±¨¯Í79•mnÈÖÓ•Iù¬šÕHI¯”-•F6$òºíJð–šÈ¼æ:¡ÎŠhvK­:òd•mUxXlª½¿?e="ñ–@.–!Î…zDâÚÕ#’3µÕ#b:"ž·y="Þ^ÈÞò«q="eÉëÖ#R‡bJæÎË	Ñ<r…¶¹ÂoÞ»RCÛtÜZW×?¤tî™ÜyX|F£´eXüIyQß’¸üx_¤È$–n(éþ]ðQ7ewÞ‹HÄ™_Y$†ŽïkÂ÷Ï¬tqd‘•5}xÍ‰n4Àíò2,1Ñ)Qƒ:	´kÀ‚)¾Ü°ùe…«P‡Çê ¹Ët|é•v‡@£­N	î™{ù5(h7Ä¨;ž1ÔÄ@8Mô1`ë×ðÎ&¦ƒ³­„ÁçC'«T €^†€§%gå
j´ÛyÔïeE7Ô+°íU$h0„ÙV$x,Ð´x€üõâî5pE×Z¿¼Â3/PB÷;*PGe¹€&´måŠ×Ž¼Ó]„¼w^älh]dÞiª1 x	¡ûÇŸ®õÖÊzGþtÖeD°S´5Ó‡ýÐ»VÙæ#:(ýJq²+3S<æ*fªÇ\7Õ¾òÑ¨ºmõcaý-ºHŸ/†R\§œ&ÁÄ¾À7¨ÅèMzžáuxvÔ /üß(ø4ôqâ4ä‹^óN|ÔæEàY­µXÎ}2]©œD…ï†à$K¯Ïœ%Îï38QÕ¶,OÞtû­;[V0ÂÙúŒº7+þUf­²›€aâ»¹ÕÈ¤"d"ÕRl¢iu‡ªiL¹ÌÇÖŸÛœ+ô/*-Îé+Ó&Mí±dMMq@¦éšð1=,9¸‘Àº¼Í‘À”œ¸-|0›$B˜ígâöTæ3ùQ~IA§Ü|–_,ðd¬žÏÓËB×mÕòé›mjuù¾5µ¨™0_&”k×ž¦îA]sš2-Gyá–	hÍ$ÈbD£±…zŽB¯ZÆØ'[øXƒ:J:Õ°H¤¤•Òn5Òn›=sœ*Ir!AÂ(™»¸›ÙzXo‰vœ‘žE—Ø»°åTÐÐV»o1ø¨ l±÷4ëA¦ÄžëKö½|LRUÆ$,G¼ÓÉ
Ä—§dt¼±[šêØ™&²4Mok²¨ø,Ê:{êÞÛD[IAó¢aƒ/s±XIÔ·ÄýæçÐÔ‹(„ij¤w­íCË€ŒDmÍ…ÖŒË4ˆz±‡mTb»ŠÞØUÁ³†_2("™ ÷©×‡5zb&"×®TÑ9dãè„É€.þW:mFŽã4Œ½pOóÁ;/ëŸuFô~ebÔÜ\Ì„Ò%úŽù)$ãåÈF½ŸbsÚc¸TâwòlÐJÚÝJVG°Q¦ \cÜ®13è±Ÿ]ø~dŠŽžÄ´JÕhæù”ôpô±øyªÐßü×®Ù)£5Þ¬*Ú¬A¬-³d¦<v<ç¡Yª¯F®%Áý–!V¬`N}†'Ÿ*‚åË[ÈXµS?’ùÝÝÞøôÝ¹¡U^$?þ„ÿ¬£Å%b½»ÖtÚgmHtªèhi96øLåÇ¨‰ÖY_&ÿ,+5ÈÎ´ëÕ0ç­x –ý/Ww£¤)wfºYÝåÇQ gç3,^®î†.*œŠN§ï§i>®ò-Öµ7G96—Ý?ÕXÝ/¦©Ó‚óH3žŒ;WRmCQó*ˆÎo~1ü Se~œÞü=êé&”)¼Ð¢f4¼&}<HÇO’âEŸ*iDÍÿ×hŽü‘v»<dñ½¸¤ðsaºƒ¾ÀDåíA’Þ7õ‡A1Áâ×2ÐœsQ	„ß<!G9ó"Ü(Dt¨1™ÅèK§àÇu·ïSXÖÝ.x--ò1ðéhlÁ YAðÊÙ¡ôø1„>xˆâ³¦7¿§ƒÁÐü˜sT£9ëŠÝØê!ëð`º3¢‘ÂvbçhI»ôžqÔ©Ú?G%
­-ÁòÓ¢Ä+Q(ç‡1§A:qªÂÄÇä?ç–Œ§]žtÈtŠJ=yÆMÅpNÌ,–ÆÑÜÃýºtcZ`š €{FÉÍoˆ2)ÆšÉÌ >žÐ7¿%ÎH‰d0¤'x]€×Ø£6€½Ä‰µè;—~Ÿ£vGÍ·Hþå`7] þ¬UÇQ¸sE±ð]ETg#ðÇ¡XªãÖTëô ¨ÚUìr#¯ÈêXÆðµ¦jÊÄ•×RRU¢„Y¨ŠYb?ÌüÂl ±JB£½À#8æŸ8nwV¶ºoˆ—QïM`,Juí€¡.@¶€6z½E¥I< ¨ËfK”>ÑœõÛ¡Õ4ÿÆmlËÔ:¨&‹Î|ž¦ÑRï,·Å¸ õ
™oæ8SHß*®¼ó#”‡dgÁ9s¥qi“ÂŠ2>‡J|”²[Òìð©¬Úh§LÚ”)ðêvÀWGÀ÷kg¶hØNhÒzH-<9Â TPÓûY°ME‹{‹×2ÅÓºëÆâ4ŠÆ©a#3DãðêãTª›ÁŽÓ®½í?x	M¥ l	€,q? cgS
ÄŽ0–¢}fï3}ÄÙ!`Ò¨Wîž++¦é?¶vOžWLå1êƒgë²Fä‹Õ®‘ÞÿóÀ™îi†´¡2Mj¥¦Nƒ¨!%n¨<WØPc"Ü8_£AÞ ÃOñZ­ˆ¢¾LPT±÷#œ¤%æ‘›¡Úzb.R•Fqš}EDzu(îoËÛ1}X¾'HÑöÛ« •m¥sm>Z“n¹}É—#,„ø–ßËÃ}õWÞ&á†<h ¿*=ø'Þ8„M'}:ÁÝ¶ã£R/«‘ÃVÚ*«#š®lý¤ÕµFÏ^uÁÑ¿™÷¤Ûbq«MÖZl•TÐ…/=Ê×ýŸ‹qG±hÜÁr`Jcûã$Ci›®É ðÔkZ À°<ro÷bÌÆª¼Yì4ó°¢^¾ð°Ò;çmw¯—©! lax(ûNƒŒmô.±‰Â.;Z*Ô-_ù¸¯_àå,Ù¬é%Ø€%a3ÝV·³çûlu‘àà·^–”öèbO:ÿÌ<F¯Aè¿§^åÊÖÐ].v@=Ë”Wªr]…æYrÖÞ§¨oƒßbƒjjQn_¢Ð)Ün€öÏ¾BGãúö2ÞÏ·W#?9ÂL{¦íðrÞX˜¬7Â]`2é¸Í/ä­Æ©ŸÐü¯ß^áW„Iø¡€y’Ä	-ðì~5¬Õw^ÿãx´s9Š“Ì¶\¬MÍµÒŠ½Œ/•üÁžËÛaÅP Çô2jUˆ'ðñ>nvûÿ·`×ÿ”âš'ôUúuÛñEÆÞ û1Q.„nýy¾áâ)î¢«ïäÀJO.’â¨ëCöÕhpò>	µÝs.¸ô$ˆ`7–¨j€¦Ò\ùzÕPòC«¶ýë/U‚$ñž'Èº¢#¤öAOàŒ,~L¸w¹t±ôõ%Ö±ÿñåúùÙO$/ÎÑC‚®'!<t~dãÌ‰Ì9–â´ÌÅNÊoªó.[A.ýÉ‹¥4¨’ÚMHj6%†¬j÷Ì5MÈ0ÚFø€¼0¦©¶•Ê9{ZòÁ‘XMã‚”©ÿe¸CY4òÕÊÙÓÚVžÒ*Óâ4¦…ýŽº›ä¡…\b›ãñX˜D\«[Á2ç/ÇHB-’ùŸK¢ó×ÆUÜ¹„Ž¼ðm}TVR L`x°´¤{œ`èáÐ ²î±gXìïl‰¿›OÆž‰&8I<l€]èQp_A1ùD]>³8éÌ³?r/¤ƒí7ä ñÏÿbî§ù4Êÿpôîíîl´ÃÌí¶D6üU@–Ø·eL‚íÿâ¬Q`j‹UÙ£Ì\ˆÀ©snjÝD4Æ½?HXÊ|÷¶ïÑk6H¦ÙmìÎ÷‰vÛS7šEÒñMëÌ{æð–Ïÿèz™›øyÍæ).ÇhÆçÀ[†§çÍÏöqè¦½´¬´8;+–šV¡O l“Å.àÑ¥ À¹Ògx¹´fô…+íÕé0®ŽÂtY™ìbð¼“q¼ªC¢%Ð™â¿ð•Yf¡âæÙ×ÉDl&»˜è_¬uz„£S?¦)y®8Š^q
‘&ÈúÒu¹6èéÎØÉPp&Žœê<ÙsÄ@È*“U–ÊÏ\mñê(œ¦–jœ€,?9e¯‘ƒRÌb9¦¦íå½Z¶œî
ïPâúŸþ?   ÿÿ ô?ý«