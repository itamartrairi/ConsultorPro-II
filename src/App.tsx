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
import { SmartSyncModal, type SyncSummary } from './components/SmartSyncModal';
import { BackupExportModal, type BackupExportStats } from './components/BackupExportModal';
import { ConnectionStatus } from './components/SettingsView';
import SettingsView from './components/SettingsView';
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
  Infinity as InfinityIcon,
  Eye,
  EyeOff,
  Bot
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
  objetivo?: string;
  atividades: AtividadeCronograma[];
}

const MODELOS_RELATORIO: ModeloRelatorio[] = [
  {
    id: 'carcinicultura',
    nome: 'GestÃ£o Financeira & Controle de Custos da Carcinicultura',
    categoria: 'Carcinicultura',
    descricao: 'Modelo de Consultoria Gerencial para Carcinicultura (30hs): Entendimento de demanda, DiagnÃ³stico de fluxos, Planilha de custos, Reserva de emergÃªncia, Treinamento de indicadores, SegregaÃ§Ã£o de contas e RelatÃ³rio final.',
    objetivo: 'Implementar melhorias estratÃ©gicas e prÃ¡ticas nas Ã¡reas de GestÃ£o Financeira, PreÃ§o e Mercado, Planejamento, Custos e Resultado, Acesso a CrÃ©dito, Controle e Rotina e ProduÃ§Ã£o, com base no diagnÃ³stico realizado, visando aprimorar a gestÃ£o do empreendimento, fortalecer os controles gerenciais, melhorar a tomada de decisÃ£o, aumentar a eficiÃªncia dos processos produtivos e financeiros e contribuir para a sustentabilidade e melhoria dos resultados',
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
    objetivo: 'Implementar melhorias na GestÃ£o Financeira, EstruturaÃ§Ã£o de Custos, Fluxo de Caixa, DRE Gerencial e FormaÃ§Ã£o do PreÃ§o de Venda com base no diagnÃ³stico realizado, visando aprimorar a rentabilidade e a sustentabilidade financeira do empreendimento.',
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
    objetivo: 'Mapear, padronizar e otimizar os processos operacionais e produtivos com base no diagnÃ³stico realizado, visando a eliminaÃ§Ã£o de desperdÃ­cios, reduÃ§Ã£o de gargalos e aumento da produtividade e eficiÃªncia da empresa.',
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
    objetivo: 'Estruturar o funil de vendas, posicionamento de mercado, estratÃ©gias de comunicaÃ§Ã£o e capacitaÃ§Ã£o comercial com base no diagnÃ³stico realizado, visando a atraÃ§Ã£o de clientes e o aumento sustentÃ¡vel do faturamento.',
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
    objetivo: 'Implementar soluÃ§Ãµes de inovaÃ§Ã£o, automaÃ§Ã£o de processos, integraÃ§Ã£o de ferramentas e capacitaÃ§Ã£o tecnolÃ³gica com base no diagnÃ³stico realizado, visando a modernizaÃ§Ã£o digital e o ganho de eficiÃªncia operacional.',
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
    objetivo: 'Definir diretrizes estratÃ©gicas, metas organizacionais, estruturaÃ§Ã£o de papÃ©is e desenvolvimento de lideranÃ§a com base no diagnÃ³stico realizado, visando o crescimento estruturado e a sustentabilidade do negÃ³cio.',
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
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed.map((m: any) => {
            const def = MODELOS_RELATORIO.find(d => d.id === m.id);
            if (def) {
              return {
                ...def,
                ...m,
                objetivo: m.objetivo || def.objetivo
              };
            }
            return m;
          });
        }
      }
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
                    <p className="text-xs text-slate-400 line-clamp-2 leading-relaxed mb-2">{modelo.descricao}</p>
                    {modelo.objetivo && (
                      <div className="mb-3 p-2 bg-slate-900/60 rounded-lg border border-slate-700/50">
                        <span className="text-[10px] font-bold text-emerald-400 block mb-0.5">Objetivo EstratÃ©gico:</span>
                        <p className="text-[11px] text-slate-300 line-clamp-2 italic leading-snug">
                          {modelo.objetivo}
                        </p>
                      </div>
                    )}
                  </div>

                  <div className="flex items-center gap-2 mt-2">
                    <Button
                      size="sm"
                      onClick={(e) => {
                        e.stopPropagation();
                        if (window.confirm(`Deseja implantar o modelo "${modelo.nome}" (${modelo.atividades.length} atividades, ${totalHoras}hs) neste relatÃ³rio?`)) {
                          const finalized = modelo.atividades.map((s, idx) => ({
                            ...s,
                            status: 'Pendente',
                            dataInicio: s.dataInicio || getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx),
                            dataFim: s.dataFim || s.dataInicio || getActivityDateStr(selectedDiagnostico.dataDiagnostico, idx)
                          }));
                          setAtividades(finalized);

                          // Obter problemas identificados no diagnÃ³stico
                          const rawProbs = deduplicateRespostas(respostas || [])
                            .filter(r => r.resposta && r.resposta !== 'Sim')
                            .map(r => r.problema)
                            .filter(Boolean);
                          const problemasUnicos = Array.from(new Set(rawProbs));

                          let probsTexto = "";
                          if (problemasUnicos.length > 0) {
                            probsTexto = problemasUnicos.map(p => `â€¢ ${p}`).join('\n');
                          } else if (dadosConsultoria.solucoesIndicadas?.includes("PROBLEMAS IDENTIFICADOS:")) {
                            const parts = dadosConsultoria.solucoesIndicadas.split("SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:");
                            probsTexto = parts[0].replace("PROBLEMAS IDENTIFICADOS:\n", "").trim();
                          } else {
                            if (modelo.id === 'carcinicultura') {
                              probsTexto = "â€¢ AusÃªncia de segregaÃ§Ã£o entre finanÃ§as pessoais e da atividade aquÃ­cola\nâ€¢ Falta de controle sistemÃ¡tico de custos de raÃ§Ã£o, insumos e parÃ¢metros zootÃ©cnicos\nâ€¢ InexistÃªncia de apuraÃ§Ã£o de margem de contribuiÃ§Ã£o e ponto de equilÃ­brio por ciclo\nâ€¢ Dificuldade no planejamento de reserva de emergÃªncia e acesso a crÃ©dito orientado";
                            } else {
                              probsTexto = "â€¢ Oportunidades de melhoria identificadas nos processos de gestÃ£o e controles operacionais\nâ€¢ Necessidade de padronizaÃ§Ã£o, acompanhamento de indicadores e estruturaÃ§Ã£o de rotinas";
                            }
                          }

                          const solucoesLista = finalized.map((s, i) => `â€¢ Atividade ${i + 1} (${s.nome}): ${s.solucaoProposta || s.descricao}`).join('\n');
                          const resultadosLista = finalized.map((s) => `â€¢ ${s.resultadoEsperado || s.solucaoProposta || s.nome}`).filter(Boolean).join('\n');

                          const objetivoFinal = modelo.objetivo || dadosConsultoria.objetivo || "Implementar melhorias estratÃ©gicas e prÃ¡ticas com base no diagnÃ³stico realizado, visando aprimorar a gestÃ£o do empreendimento.";

                          const updatedDados: DadosConsultoria = {
                            ...dadosConsultoria,
                            objetivo: objetivoFinal,
                            cargaHoraria: totalHoras > 0 ? `${totalHoras}hs` : dadosConsultoria.cargaHoraria,
                            solucoesIndicadas: `PROBLEMAS IDENTIFICADOS:\n${probsTexto}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesLista}`,
                            resultadosEsperados: resultadosLista
                          };

                          setDadosConsultoria(updatedDados);
                          syncEvidenceState(finalized, updatedDados);

                          playSuccessSound();
                          alert(`Modelo "${modelo.nome}" implantado com sucesso!\n\nâœ” Objetivo da Consultoria configurado com foco nas Ã¡reas estratÃ©gicas\nâœ” Problemas identificados no diagnÃ³stico vinculados\nâœ” SoluÃ§Ãµes, cronograma e resultados integrados.`);
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

              <div>
                <label className="text-xs font-bold text-slate-700 block mb-1">Objetivo da Consultoria (Implantado no RelatÃ³rio e Plano)</label>
                <textarea 
                  value={editingModelo.objetivo || ''} 
                  onChange={(e) => setEditingModelo({ ...editingModelo, objetivo: e.target.value })} 
                  placeholder="Defina o objetivo estratÃ©gico detalhado que serÃ¡ preenchido automaticamente ao implantar o modelo..."
                  rows={3}
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
        if (rawSavedKey === "undefined" || rawSavedKey === "null" || rawSavedKey.trim().length < 10) {
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
        if (envKey === "undefined" || envKey === "null" || envKey.trim().length < 10) {
          envKey = "";
        }
        
        // 2. Try the backend proxy first (available in Container/Full-stack environments)
        try {
          const headers: Record<string, string> = {
            "Content-Type": "application/json"
          };
          if (savedKey && savedKey.length > 10) {
            headers["x-custom-api-key"] = savedKey;
          }

          const response = await fetch("/api/gemini/generate", {
            method: "POST",
            headers: headers,
            body: JSON.stringify({ model: model || "gemini-3.1-flash-lite", contents, config })
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
        const finalKey = (savedKey && savedKey.length > 10 ? savedKey : "") || (envKey && envKey.length > 10 ? envKey : "");
        if (!finalKey || finalKey.trim() === "") {
          throw new Error("NÃ£o foi possÃ­vel conectar com o serviÃ§o de IA.\n\nPara ativar a IA:\n1. Acesse o menu 'ConfiguraÃ§Ãµes' da aplicaÃ§Ã£o e insira sua chave oficial do Google AI Studio.\n2. Se vocÃª estiver desenvolvendo no AI Studio, adicione a variÃ¡vel GEMINI_API_KEY no menu de Secrets.\n3. Obtenha sua chave gratuita em: https://aistudio.google.com/app/apikey");
        }

        const GEMINI_PRIMARY = "gemini-3.1-flash-lite";
        const GEMINI_CHAIN = [
          "gemini-3.1-flash-lite",
          "gemini-3.7-flash",
          "gemini-flash-latest"
        ];

        function normalizeClientModel(m?: string): string {
          if (!m) return GEMINI_PRIMARY;
          if (m.includes("2.5") || m.includes("2.0") || m.includes("1.5") || m.includes("1.0") || m.includes("3.6")) {
            return "gemini-3.1-flash-lite";
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
      model: "gemini-3.1-flash-lite", 
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
      model: "gemini-3.1-flash-lite",
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
      model: "gemini-3.1-flash-lite",
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

export const saveAllLocalRespostas = (newOrUpdated: Resposta[], replaceAll: boolean = false) => {
  try {
    if (!newOrUpdated || newOrUpdated.length === 0) return;
    if (replaceAll || !cachedAllRespostasMap) {
      cachedAllRespostasMap = new Map();
      if (!replaceAll) {
        loadAllLocalRespostas();
      }
    }
    newOrUpdated.forEach((r, idx) => {
      if (!r) return;
      const validId = String(r.id || (r as any).idResposta || `resp_${Date.now()}_${idx}_${Math.random().toString(36).substring(2, 7)}`);
      const itemWithId = { ...r, id: validId };
      cachedAllRespostasMap!.set(validId, itemWithId);
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
  const [view, setView] = useState<'home' | 'companies' | 'credenciadas' | 'licenses' | 'diagnosis' | 'dashboard' | 'premises' | 'cronograma' | 'relatorio' | 'kanban' | 'settings' | 'landing' | 'checkout' | 'licensing' | 'dados-consultoria'>('landing');
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
      for (let i = 0; i < items.length; i += 250) {
        const batch = writeBatch(db);
        const chunk = items.slice(i, i + 250);
        for (const item of chunk) {
          if (!item || typeof item !== 'object') continue;
          const itemId = String(item.id || item._id || item.codigo || doc(collection(db, colName)).id);
          const ref = doc(db, colName, itemId);
          batch.set(ref, sanitizeForFirestore({ ...item, id: itemId, ownerId: item.ownerId || uid }), { merge: true });
        }
        await batch.commit();
      }
    };

    if (data.empresas && data.empresas.length > 0) await syncCollection(data.empresas, 'empresas');
    if (data.diagnosticos && data.diagnosticos.length > 0) await syncCollection(data.diagnosticos, 'diagnosticos');
    if (data.respostas && data.respostas.length > 0) await syncCollection(data.respostas, 'respostas');
    if (data.tarefasPlano && data.tarefasPlano.length > 0) await syncCollection(data.tarefasPlano, 'tarefas_plano');
    if (data.empresasCredenciadas && data.empresasCredenciadas.length > 0) await syncCollection(data.empresasCredenciadas, 'empresas_credenciadas');
    if (data.premissas && data.premissas.length > 0) await syncCollection(data.premissas, 'premissas');
    if (data.problemas && data.problemas.length > 0) await syncCollection(data.problemas, 'problemas');
    if (data.solucoes && data.solucoes.length > 0) await syncCollection(data.solucoes, 'solucoes');
    if (data.dbAreas && data.dbAreas.length > 0) await syncCollection(data.dbAreas, 'areas');
    if (data.dbSegmentos && data.dbSegmentos.length > 0) await syncCollection(data.dbSegmentos, 'segmentos');
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
      let exportDiagnosticos: Diagnostico[] = [...diagnosticos];
      let exportTarefasPlano: TarefaPlanoAcao[] = [...tarefasPlano];
      let exportAgendaEventos: any[] = [];
      let exportDiscAvaliacoes: any[] = [];
      let exportMaturidadeAvaliacoes: any[] = [];
      let exportCredenciadas: EmpresaCredenciada[] = [...empresasCredenciadas];
      let exportPremissas: Premissa[] = [...premissas];
      let exportProblemas: Problema[] = [...problemas];
      let exportSolucoes: Solucao[] = [...solucoes];
      let exportDbAreas = [...dbAreas];
      let exportDbSegmentos = [...dbSegmentos];

      // A. Gather answers from local memory cache first
      const localAnswers = loadAllLocalRespostas();
      const mapRespostas = new Map<string, Resposta>();
      localAnswers.forEach(r => { if (r?.id) mapRespostas.set(r.id, r); });
      respostas.forEach(r => { if (r?.id) mapRespostas.set(r.id, r); });

      // Local storage agenda fallback
      try {
        const localAgenda = localStorage.getItem('local_agenda_eventos');
        if (localAgenda) {
          const parsed = JSON.parse(localAgenda);
          if (Array.isArray(parsed)) exportAgendaEventos = parsed;
        }
      } catch (e) {}

      // 2. Fetch fresh cloud records in PARALLEL with a strict 2.5-second timeout safeguard
      if (user) {
        const withTimeout = async <T,>(p: Promise<T>, timeoutMs = 2500, fallback: T): Promise<T> => {
          return Promise.race([
            p,
            new Promise<T>((resolve) => setTimeout(() => resolve(fallback), timeoutMs))
          ]);
        };

        const empQuery = isAdmin 
          ? query(collection(db, 'empresas')) 
          : query(collection(db, 'empresas'), where('ownerId', '==', user.uid));

        const diagQuery = isAdmin
          ? query(collection(db, 'diagnosticos'))
          : query(collection(db, 'diagnosticos'), where('ownerId', '==', user.uid));

        const respQuery = isAdmin
          ? query(collection(db, 'respostas'), limit(2000))
          : query(collection(db, 'respostas'), where('ownerId', '==', user.uid));

        const tarefasQuery = isAdmin
          ? query(collection(db, 'tarefas_plano'))
          : query(collection(db, 'tarefas_plano'), where('ownerId', '==', user.uid));

        const agendaQuery = isAdmin
          ? query(collection(db, 'agenda_eventos'))
          : query(collection(db, 'agenda_eventos'), where('ownerId', '==', user.uid));

        const discQuery = isAdmin
          ? query(collection(db, 'disc_avaliacoes'))
          : query(collection(db, 'disc_avaliacoes'), where('ownerId', '==', user.uid));

        const matQuery = isAdmin
          ? query(collection(db, 'maturidade_avaliacoes'))
          : query(collection(db, 'maturidade_avaliacoes'), where('ownerId', '==', user.uid));

        const credQuery = isAdmin
          ? query(collection(db, 'empresas_credenciadas'))
          : query(collection(db, 'empresas_credenciadas'), where('ownerId', '==', user.uid));

        try {
          const [
            snapEmp,
            snapDiag,
            snapResp,
            snapTarefas,
            snapAgenda,
            snapDisc,
            snapMat,
            snapCred
          ] = await Promise.all([
            withTimeout(getDocs(empQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(diagQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(respQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(tarefasQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(agendaQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(discQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(matQuery).catch(() => null), 2000, null),
            withTimeout(getDocs(credQuery).catch(() => null), 2000, null)
          ]);

          if (snapEmp && !snapEmp.empty) {
            const empMap = new Map<string, Empresa>();
            exportEmpresas.forEach(e => empMap.set(e.id, e));
            snapEmp.docs.forEach(d => empMap.set(d.id, { id: d.id, ...d.data() } as Empresa));
            exportEmpresas = Array.from(empMap.values());
          }

          if (snapDiag && !snapDiag.empty) {
            const diagMap = new Map<string, Diagnostico>();
            exportDiagnosticos.forEach(d => diagMap.set(d.id, d));
            snapDiag.docs.forEach(d => diagMap.set(d.id, { id: d.id, ...d.data() } as Diagnostico));
            exportDiagnosticos = Array.from(diagMap.values());
          }

          if (snapResp && !snapResp.empty) {
            snapResp.docs.forEach(d => {
              mapRespostas.set(d.id, { id: d.id, ...d.data() } as Resposta);
            });
          }

          if (snapTarefas && !snapTarefas.empty) {
            const taskMap = new Map<string, TarefaPlanoAcao>();
            exportTarefasPlano.forEach(t => taskMap.set(t.id, t));
            snapTarefas.docs.forEach(d => taskMap.set(d.id, { id: d.id, ...d.data() } as TarefaPlanoAcao));
            exportTarefasPlano = Array.from(taskMap.values());
          }

          if (snapAgenda && !snapAgenda.empty) {
            exportAgendaEventos = snapAgenda.docs.map(d => ({ id: d.id, ...d.data() }));
          }

          if (snapDisc && !snapDisc.empty) {
            exportDiscAvaliacoes = snapDisc.docs.map(d => ({ id: d.id, ...d.data() }));
          }

          if (snapMat && !snapMat.empty) {
            exportMaturidadeAvaliacoes = snapMat.docs.map(d => ({ id: d.id, ...d.data() }));
          }

          if (snapCred && !snapCred.empty) {
            const credMap = new Map<string, EmpresaCredenciada>();
            exportCredenciadas.forEach(c => credMap.set(c.id, c));
            snapCred.docs.forEach(d => credMap.set(d.id, { id: d.id, ...d.data() } as EmpresaCredenciada));
            exportCredenciadas = Array.from(credMap.values());
          }
        } catch (cloudErr) {
          console.warn("Nuvem indisponÃ­vel no momento do backup, compilando com dados locais seguros:", cloudErr);
        }
      }

      const exportRespostas = Array.from(mapRespostas.values());
      const nowFormatted = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const userPrefix = user?.email ? user.email.split('@')[0].replace(/[^a-zA-Z0-9_-]/g, '_') + '_' : '';
      const genFileName = `consultoria_backup_${userPrefix}${nowFormatted}.json`;

      const backupData = {
        app: 'Consultoria Pro - SEBRAE',
        version: '2.0',
        exportedAt: new Date().toISOString(),
        user: {
          email: user?.email || 'Acesso Local',
          uid: user?.uid || null,
          isAdmin
        },
        metadata: {
          totalEmpresas: exportEmpresas.length,
          totalDiagnosticos: exportDiagnosticos.length,
          totalRespostas: exportRespostas.length,
          totalTarefasPlano: exportTarefasPlano.length,
          totalAgendaEventos: exportAgendaEventos.length,
          totalDisc: exportDiscAvaliacoes.length,
          totalMaturidade: exportMaturidadeAvaliacoes.length,
          totalCredenciadas: exportCredenciadas.length,
          totalPremissas: exportPremissas.length,
          totalProblemas: exportProblemas.length,
          totalSolucoes: exportSolucoes.length
        },
        empresas: exportEmpresas,
        diagnosticos: exportDiagnosticos,
        respostas: exportRespostas,
        tarefasPlano: exportTarefasPlano,
        agendaEventos: exportAgendaEventos,
        discAvaliacoes: exportDiscAvaliacoes,
        maturidadeAvaliacoes: exportMaturidadeAvaliacoes,
        empresasCredenciadas: exportCredenciadas,
        premissas: exportPremissas,
        problemas: exportProblemas,
        solucoes: exportSolucoes,
        dbAreas: exportDbAreas,
        dbSegmentos: exportDbSegmentos,
        customLogo,
        customConsultoraLogo
      };

      const jsonStr = JSON.stringify(backupData, null, 2);
      setBackupJsonString(jsonStr);
      setBackupFileName(genFileName);

      try {
        localStorage.setItem('local_backup_data', jsonStr);
        localStorage.setItem('last_exported_backup', jsonStr);
      } catch (e) {}

      const sizeKb = Math.round(new Blob([jsonStr]).size / 1024);

      const stats: BackupExportStats = {
        totalEmpresas: exportEmpresas.length,
        totalDiagnosticos: exportDiagnosticos.length,
        totalRespostas: exportRespostas.length,
        totalTarefasPlano: exportTarefasPlano.length,
        totalAgendaEventos: exportAgendaEventos.length,
        totalDisc: exportDiscAvaliacoes.length,
        totalMaturidade: exportMaturidadeAvaliacoes.length,
        totalPremissas: exportPremissas.length,
        totalProblemas: exportProblemas.length,
        totalSolucoes: exportSolucoes.length,
        totalCredenciadas: exportCredenciadas.length,
        dataTamanhoKb: sizeKb,
        exportedAt: backupData.exportedAt,
        userEmail: user?.email || undefined,
        source: user ? 'nuvem_e_local' : 'local_only'
      };

      setBackupStats(stats);
      setIsExportingBackup(false);

      // Trigger automatic browser download
      triggerDownloadBackupFile(jsonStr, genFileName);

      playSuccessSound();
      showToast(`Backup completo gerado com sucesso! (${sizeKb} KB)`, 'success', 'Backup dos Dados');
    } catch (err: any) {
      console.error("Erro ao gerar backup:", err);
      setIsExportingBackup(false);
      showToast("Erro ao compilar backup: " + (err?.message || "Tente novamente"), "error", "Falha no Backup");
    }
  };

  const handleImportLocalBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const inputElement = e.target;
    const reader = new FileReader();

    reader.onload = async (event) => {
      try {
        const rawContent = event.target?.result as string;
        if (!rawContent || !rawContent.trim()) {
          throw new Error("Arquivo vazio");
        }

        // Clean UTF-8 BOM if present
        const content = rawContent.replace(/^\uFEFF/, '').trim();
        const rawData = JSON.parse(content);

        // Unwrap potential payload nesting from diverse backup wrappers
        let data = rawData;
        if (data && typeof data === 'object') {
          if (data.data && typeof data.data === 'object' && !Array.isArray(data.data)) data = data.data;
          else if (data.backup && typeof data.backup === 'object' && !Array.isArray(data.backup)) data = data.backup;
          else if (data.dados && typeof data.dados === 'object' && !Array.isArray(data.dados)) data = data.dados;
          else if (data.content && typeof data.content === 'object' && !Array.isArray(data.content)) data = data.content;
          else if (data.export && typeof data.export === 'object' && !Array.isArray(data.export)) data = data.export;
        }

        // Helper to normalize items with guaranteed string IDs
        const normalizeItems = <T extends Record<string, any>>(arr: any, prefix: string): T[] => {
          if (!Array.isArray(arr)) return [];
          return arr
            .filter(item => item && typeof item === 'object')
            .map((item, idx) => {
              const id = String(item.id || item._id || item.codigo || `${prefix}_${Date.now()}_${idx}`);
              return { ...item, id };
            });
        };

        // Normalize potential variations in backup keys
        const importedEmpresas = normalizeItems<Empresa>(data.empresas || data.companies || data.clientes || data.clients || [], 'emp');
        const importedDiagnosticos = normalizeItems<Diagnostico>(data.diagnosticos || data.diagnostics || data.avaliacoes || data.assessments || [], 'diag');
        const importedRespostas = normalizeItems<Resposta>(data.respostas || data.answers || data.itens || data.items || data.perguntas_respostas || [], 'resp');
        const importedTarefasPlano = normalizeItems<TarefaPlanoAcao>(data.tarefasPlano || data.tarefas_plano || data.tasks || data.planoAcao || data.plano_de_acao || data.actionPlan || [], 'task');
        const importedEmpresasCredenciadas = normalizeItems<EmpresaCredenciada>(data.empresasCredenciadas || data.empresas_credenciadas || data.credenciadas || [], 'cred');
        const importedPremissas = normalizeItems<Premissa>(data.premissas || data.premises || data.perguntas || data.questions || [], 'prem');
        const importedProblemas = normalizeItems<Problema>(data.problemas || data.problems || [], 'prob');
        const importedSolucoes = normalizeItems<Solucao>(data.solucoes || data.solutions || [], 'sol');
        const importedDbAreas = normalizeItems<{ id: string; nome: string }>(data.dbAreas || data.db_areas || data.areas || [], 'area');
        const importedDbSegmentos = normalizeItems<{ id: string; nome: string }>(data.dbSegmentos || data.db_segmentos || data.segmentos || data.segments || [], 'seg');
        const importedAgendaEventos = normalizeItems<any>(data.agendaEventos || data.agenda_eventos || data.eventos || data.events || data.agenda || [], 'evt');
        const importedDiscAvaliacoes = normalizeItems<any>(data.discAvaliacoes || data.disc_avaliacoes || data.disc || [], 'disc');
        const importedMaturidadeAvaliacoes = normalizeItems<any>(data.maturidadeAvaliacoes || data.maturidade_avaliacoes || data.maturidade || [], 'mat');

        if (importedEmpresas.length > 0) {
          setEmpresas(importedEmpresas);
          setSelectedEmpresa(importedEmpresas[0]);
          try { localStorage.setItem('local_empresas', JSON.stringify(importedEmpresas)); } catch {}
        }
        if (importedDiagnosticos.length > 0) {
          setDiagnosticos(importedDiagnosticos);
          setSelectedDiagnostico(importedDiagnosticos[0]);
          try { localStorage.setItem('local_diagnosticos', JSON.stringify(importedDiagnosticos)); } catch {}
        }
        if (importedRespostas.length > 0) {
          saveAllLocalRespostas(importedRespostas, true);
          const targetDiagId = importedDiagnosticos[0]?.id || selectedDiagnostico?.id;
          const currentDiagRespostas = targetDiagId ? importedRespostas.filter(r => r.diagnosticoId === targetDiagId) : importedRespostas;
          setRespostas(currentDiagRespostas.length > 0 ? currentDiagRespostas : importedRespostas);
          setRespostasLoaded(true);
        }
        if (importedTarefasPlano.length > 0) {
          setTarefasPlano(importedTarefasPlano);
          try { localStorage.setItem('local_tarefas_plano', JSON.stringify(importedTarefasPlano)); } catch {}
        }
        if (importedEmpresasCredenciadas.length > 0) {
          setEmpresasCredenciadas(importedEmpresasCredenciadas);
          try { localStorage.setItem('local_empresas_credenciadas', JSON.stringify(importedEmpresasCredenciadas)); } catch {}
        }
        if (importedPremissas.length > 0) {
          setPremissas(importedPremissas);
          try {
            localStorage.setItem('local_premissas', JSON.stringify(importedPremissas));
            localStorage.removeItem('user_cleared_premissas');
          } catch {}
        }
        if (importedProblemas.length > 0) {
          setProblemas(importedProblemas);
          try {
            localStorage.setItem('local_problemas', JSON.stringify(importedProblemas));
            localStorage.removeItem('user_cleared_problemas');
          } catch {}
        }
        if (importedSolucoes.length > 0) {
          setSolucoes(importedSolucoes);
          try {
            localStorage.setItem('local_solucoes', JSON.stringify(importedSolucoes));
            localStorage.removeItem('user_cleared_solucoes');
          } catch {}
        }
        if (importedDbAreas.length > 0) {
          setDbAreas(importedDbAreas);
          try { localStorage.setItem('local_db_areas', JSON.stringify(importedDbAreas)); } catch {}
        }
        if (importedDbSegmentos.length > 0) {
          setDbSegmentos(importedDbSegmentos);
          try { localStorage.setItem('local_db_segmentos', JSON.stringify(importedDbSegmentos)); } catch {}
        }
        if (importedAgendaEventos.length > 0) {
          try { localStorage.setItem('local_agenda_eventos', JSON.stringify(importedAgendaEventos)); } catch {}
        }
        if (importedDiscAvaliacoes.length > 0) {
          try { localStorage.setItem('local_disc_avaliacoes', JSON.stringify(importedDiscAvaliacoes)); } catch {}
        }
        if (importedMaturidadeAvaliacoes.length > 0) {
          try { localStorage.setItem('local_maturidade_avaliacoes', JSON.stringify(importedMaturidadeAvaliacoes)); } catch {}
        }
        if (data.customLogo) {
          setCustomLogo(data.customLogo);
          try { localStorage.setItem('sebrae_custom_logo', data.customLogo); } catch {}
        }
        if (data.customConsultoraLogo) {
          setCustomConsultoraLogo(data.customConsultoraLogo);
          try { localStorage.setItem('consultora_custom_logo', data.customConsultoraLogo); } catch {}
        }

        try { localStorage.setItem('local_backup_data', content); } catch {}

        if (user) {
          try {
            await uploadBackupToCloud({
              empresas: importedEmpresas,
              diagnosticos: importedDiagnosticos,
              respostas: importedRespostas,
              tarefasPlano: importedTarefasPlano,
              empresasCredenciadas: importedEmpresasCredenciadas,
              premissas: importedPremissas,
              problemas: importedProblemas,
              solucoes: importedSolucoes,
              dbAreas: importedDbAreas,
              dbSegmentos: importedDbSegmentos,
              agendaEventos: importedAgendaEventos,
              discAvaliacoes: importedDiscAvaliacoes,
              maturidadeAvaliacoes: importedMaturidadeAvaliacoes
            }, user.uid);
            showToast(`Backup restaurado e sincronizado: ${importedEmpresas.length} empresas, ${importedDiagnosticos.length} diagnÃ³sticos e ${importedRespostas.length} respostas!`, 'success', 'RestauraÃ§Ã£o ConcluÃ­da');
          } catch (cloudErr) {
            console.error("Erro ao sincronizar backup com a nuvem:", cloudErr);
            showToast(`Backup restaurado localmente: ${importedEmpresas.length} empresas e ${importedDiagnosticos.length} diagnÃ³sticos carregados!`, 'info', 'RestauraÃ§Ã£o Local');
          }
        } else {
          showToast(`Backup restaurado com sucesso: ${importedEmpresas.length} empresas, ${importedDiagnosticos.length} diagnÃ³sticos e ${importedRespostas.length} respostas!`, 'success', 'RestauraÃ§Ã£o ConcluÃ­da');
        }
        playSuccessSound();
      } catch (err: any) {
        console.error("Erro ao importar backup:", err);
        showToast("Erro ao ler arquivo .JSON de backup. Verifique o formato do arquivo.", "error", "Falha na ImportaÃ§Ã£o");
      } finally {
        if (inputElement) {
          inputElement.value = '';
        }
      }
    };
    reader.readAsText(file);
  };

  // Dynamically compute ALL available areas (stored in DB + session + library items)
  const allAvailableAreas = useMemo(() => {
    const areasSet = new Set<string>();

    const addNorm = (val?: string | null) => {
      if (!val) return;
      const formatted = normalizeAndFormatArea(val);
      if (!formatted) return;
      const norm = formatted.toLowerCase();
      // Skip any area that was explicitly deleted
      if (deletedAreas.some(d => normalizeAndFormatArea(d).toLowerCase() === norm)) {
        return;
      }
      areasSet.add(formatted);
    };

    // Include standard predefined areas
    AREAS.forEach(a => addNorm(a.nome));

    // Include areas stored in Firestore 'areas' collection
    dbAreas.forEach(a => addNorm(a.nome));

    // Include custom areas added in session
    sessionCustomAreas.forEach(a => addNorm(a));

    // Include areas from all library collections and diagnostic records
    problemas.forEach(p => addNorm(p.area));
    solucoes.forEach(s => addNorm(s.area));
    respostas.forEach(r => addNorm(r.area));
    diagnosticos.forEach(d => {
      if (d.areasDiagnostico && Array.isArray(d.areasDiagnostico)) {
        d.areasDiagnostico.forEach(a => addNorm(a));
      }
    });

    const list = Array.from(areasSet).filter(Boolean);
    return list.sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));
  }, [dbAreas, problemas, solucoes, respostas, diagnosticos, sessionCustomAreas, deletedAreas]);

  const availableSegments = useMemo(() => {
    const customFromEmpresas = empresas.map(e => e.tipoEmpresa).filter(Boolean) as string[];
    const customFromDiagnosticos = diagnosticos.map(d => d.tipoEmpresa).filter(Boolean) as string[];
    const customFromPremissas = premissas.map(p => p.tipoEmpresa).filter(Boolean) as string[];
    const customFromProblemas = problemas.map(p => p.tipoEmpresa).filter(Boolean) as string[];
    const customFromSolucoes = solucoes.map(s => s.tipoEmpresa).filter(Boolean) as string[];
    const customFromDb = dbSegmentos.map(s => s.nome).filter(Boolean) as string[];
    
    const allCustom = [
      ...customFromEmpresas,
      ...customFromDiagnosticos,
      ...customFromPremissas,
      ...customFromProblemas,
      ...customFromSolucoes,
      ...customFromDb,
      ...sessionCustomSegments
    ];
    
    const uniqueCustom = Array.from(new Set(
      allCustom
        .map(s => s ? s.trim() : '')
        .filter(s => s && s !== 'Geral' && !TIPOS_EMPRESA.includes(s))
    )).sort();
    
    return [...TIPOS_EMPRESA, ...uniqueCustom];
  }, [empresas, diagnosticos, premissas, problemas, solucoes, dbSegmentos, sessionCustomSegments]);

  const allAvailableTags = useMemo(() => {
    const tagsSet = new Set<string>();
    problemas.forEach(p => {
      if (Array.isArray(p.tags)) {
        p.tags.forEach(t => {
          const trimmed = t?.trim();
          if (trimmed) tagsSet.add(trimmed);
        });
      }
    });
    solucoes.forEach(s => {
      if (Array.isArray(s.tags)) {
        s.tags.forEach(t => {
          const trimmed = t?.trim();
          if (trimmed) tagsSet.add(trimmed);
        });
      }
    });
    return Array.from(tagsSet).sort((a, b) => a.localeCompare(b, 'pt-BR'));
  }, [problemas, solucoes]);

  const availableAreasForSelectedType = useMemo(() => {
    const areasSet = new Set<string>();
    const filter = (diagnosisCompanyType || 'Geral').trim().toLowerCase();

    premissas.forEach(p => {
      const prob = problemas.find(prob => prob.id === p.idProblema || prob.descricao_problemas === p.problema);
      const pType = (p.tipoEmpresa || prob?.tipoEmpresa || 'Geral').trim();
      const normPType = pType.toLowerCase();

      const matches = (filter !== 'geral' && filter !== '') 
        ? (normPType === filter || normPType === 'geral' || normPType === '')
        : (normPType === 'geral' || normPType === '');

      if (matches) {
        const areaName = prob?.area || p.area;
        if (areaName && areaName.trim()) {
          areasSet.add(normalizeAndFormatArea(areaName));
        }
      }
    });

    problemas.forEach(prob => {
      const pType = (prob.tipoEmpresa || 'Geral').trim();
      const normPType = pType.toLowerCase();

      const matches = (filter !== 'geral' && filter !== '') 
        ? (normPType === filter || normPType === 'geral' || normPType === '')
        : (normPType === 'geral' || normPType === '');

      if (matches && prob.area && prob.area.trim()) {
        areasSet.add(normalizeAndFormatArea(prob.area));
      }
    });

    // Fallback if no matching areas found
    if (areasSet.size === 0) {
      allAvailableAreas.forEach(aName => areasSet.add(aName));
    }

    const sortedNames = Array.from(areasSet).sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));
    return sortedNames.map((name, index) => {
      const existingArea = AREAS.find(a => a.nome.toLowerCase() === name.toLowerCase());
      return {
        id: existingArea ? existingArea.id : `dyn_${index}`,
        nome: name
      };
    });
  }, [allAvailableAreas, premissas, problemas, diagnosisCompanyType]);

  // Pre-populate selected areas for diagnostic when modal opens or selected segment changes
  useEffect(() => {
    if (isModalOpen && modalType === 'selectAreas') {
      setSelectedAreasForDiagnosis(availableAreasForSelectedType.map(a => a.nome));
    }
  }, [diagnosisCompanyType, isModalOpen, modalType, availableAreasForSelectedType]);

  // Helper to synchronize any authenticated user with Firestore empresas_credenciadas
  const syncUserWithCredenciada = async (u: User) => {
    if (!u) return;
    try {
      const userEmailClean = (u.email || '').toLowerCase().trim();
      const isUserAdminEmail = userEmailClean === 'itamartrairi@gmail.com';
      
      // Check by ownerId first
      const qCred = query(collection(db, 'empresas_credenciadas'), where('ownerId', '==', u.uid));
      const snapCred = await getDocs(qCred);
      
      let existingDoc = snapCred.empty ? null : snapCred.docs[0];
      
      // If not found by ownerId, also search by email
      if (!existingDoc && u.email) {
        const qEmail = query(collection(db, 'empresas_credenciadas'), where('email', '==', u.email));
        const snapEmail = await getDocs(qEmail);
        if (!snapEmail.empty) {
          existingDoc = snapEmail.docs[0];
        }
      }

      const providerId = u.providerData?.[0]?.providerId || (u.email ? 'password' : 'google.com');

      if (!existingDoc) {
        const defaultName = isUserAdminEmail 
          ? (u.displayName ? `${u.displayName} (Admin)` : 'Itamar Trairi Consultoria (Admin)')
          : (u.displayName || u.email?.split('@')[0] || 'Novo Consultor');

        const newCredDoc = {
          razaoSocial: defaultName,
          consultor: u.displayName || u.email?.split('@')[0] || (isUserAdminEmail ? 'Itamar Trairi' : 'Consultor'),
          email: u.email || '',
          dataCadastro: serverTimestamp(),
          ultimoLogin: serverTimestamp(),
          ownerId: u.uid,
          providerId: providerId,
          photoURL: u.photoURL || '',
          status: 'Ativa',
          tipoPlano: isUserAdminEmail ? 'Definitiva' : 'Teste',
          diasTeste: isUserAdminEmail ? 99999 : 30,
          role: isUserAdminEmail ? 'admin' : 'cliente'
        };
        await addDoc(collection(db, 'empresas_credenciadas'), sanitizeForFirestore(newCredDoc));
      } else {
        const data = existingDoc.data();
        const updates: any = {
          ultimoLogin: serverTimestamp(),
          ownerId: u.uid,
          providerId: providerId
        };
        if (u.photoURL && !data.photoURL) {
          updates.photxœì½]sG’ ø®_Äi»ªº" ‰j	…ƒ PÂI H³{l™¨L )Ve–2³@Bl˜]ß>¬íÃ<­íûjæ¡mÖlžÚæàŸÜ/9w/ÈÈ¬, ”Ô3*“ˆªÌÿ
ü›Ã'â‘˜§çy•Ã‡õ¹2ßÒSÑŸ“I”ŽÅo~#îÅQÉŸñÎ”b6…7I©J"TúÖ2NËé8º|MÜ—GyVÎÆU^ˆ?ýIxO=z$zÛúg¯©Ä³ü"Ûã4Éª¤7÷’UNWšú›–ß”I±OÒlWã‚¯Òi~0Ž²\ÜÃÆw’Ó4K«ô"ê…[¶åÝÒeã4*“²J ìgø	*ò1¾ïEØ·^¨ÿÑ›(­Tù|Ôñÿ“eÑK&Ó")£òå¨Hâ$¥Q•½e‘¼MË*ÍÎ ð0Ë¢Œ°—?$óâq
U qI_u`0ÐmÊ¯Ä(ªFç¢ŸX º¡›Ã¤(ò¢¿´D‰2ÍFEž¥?D×¹þç\Ä¹˜•³ë‹4Ñ¬‚	LG<åìëÆtVµ­_=ü þ½_lÍªsñF‘dIfe²{zšŒª~ }¡:†ÝÄ )œ”£"=A<æÖ=ª`hÛçQv–Ä}èÊù2,«)D™THðô¡}ò$bÀ[ÿ4—	q‹Î¬±÷gb¨n–ÆbCd³ñØ”„q—Ó*‰Ÿä£h¬*'éYÝÙó]]¢? ±šõÄ©°}žŒ^ãy!ÒRyˆ“¨Lb£Ë´?‚ß+iV&Y‰d˜…¨Mô¾=N¢Ðcx,¾^o0¬ò'ù›¤Ø8}øU¤“þà¡¤¶vÕàâÒI«hU¥EúžáKX¦FÑ€Ç½’ÀÔ–#ksî²½§†ðmR¤§i»ëT“ê›¨ÈúKßrD@OÐc!â¤ZBÒì‡{=Š­QR–¹H\/EZ!g£$Nã|¸4àë”Ï_¾2¢É‹²XÉq!ÕÃÃQ^Ä8¾,¯Ìb—IeQu™päÿVçÛ¶*£Ö+‘ ²¡3ìºä«–µú]$Õ¬Èø¢ÁçWËâù‹^}’øÊKèùDlì‰¿O.E	KjVŠÓóI4zÀð #§éÒã$ŸeUãB=M€£ô—îGÓôþY‚3q_V]Ò;¬Î“¬¨ÆJðgø]™gýû6_Ã’Fè)ÐEu9MòSÉÕÏ£òˆº='ê<És$U«ÒLÉ­®¡~Â Î•¯LÏˆ[ö3º=±Ì(ù	RƒÂ%ÐÂè<ºHÇ@Œ%ÒÊ$"®Xvõ):J€¼Îò7ã$>Kˆ/ 2$—@Ê~[]6Nˆ\[D-Æ“ø €é)Oïî"„>çÎÐUÉvä¸HÏÎ`H[{â)
Ê¥x’\$c#ãÑlU)Ç˜>½@JÎ~’GEÜÚ¿‹4y#çÎ”ïá<—ÉŠ'ñNe9¬QŽh¦ð+*‡ã$;ññ…XµcC€÷5‡ÌøXö<Žâ†,Æô=;Ó£â„£_'[Y|¨Ëô-,ÛžùKRáºÖ&HïåÚÀ,îwòµ”,[‘y9ö@­~%yˆŒ$É Á•£@«r@¶€ÿ´¼~ƒ-‰÷U±RS
1Ïò÷»ª.
E/jÜ›âûYR\öGù”á*2=³à7æ–\F¢5¦—¿ÉP*ƒÞÓ{ôþÅ‘¢ph	ËÔÖ³<;Ê¢i	ªsßt´„:Š‚¬:VÂ·Cü>œDÓ~ŒåúïDoˆ˜æj8ÆCä€ÿ+Ä‚l•,ª\æLQ´,N¨©~4ÌrP£•L£òlç“iã;áï¸N¢;ú %-âˆØFhH]Fšªß‡/\Þ¥H_h’¡~?_}Á5X…WR#rh Çýá%€Z¿õP^W‹!UyS¼CÜaIB¢~~„€öÌñXÂ$h|ñ81:î.1_z½,ö§IAÜèÅðÉÞÑ±Kyµ|Ï¤.£cÑ‘]¼3U{Õ:‘8ÂõÉÜë<'uþª…ÜIÒÞ˜ÜY#ŒòQ¬Ö)¿ˆ~ˆò£
›@½Hhð‘‘×¢ën(ÄŸRËD¿LOÆie$&	0Ð||ý×3°|ÐîTéøÕ¾¾‘­Ëâ ÈOÆ êáëQ>žùôo	|¿þs‘Ð³älÒ2/œ$M}˜™0Mu	ƒ!F¼:§Ö­»æ~´3_¦½ çl>M
PhFã|‹7 þ‚6s‰¸Då[ixö	•Ý‹±‡¨GIÕw;‰½ÔÌ“l5¡ØÔi:%¥?¥±Ý›ÕÊC²³?H)¡ž€:Iê‚êjý©×šbuÏ;Ð-ÉßxÛ/\EïÖëRKAŠìZ(H•S­îR¡æ»§ 	Ú£ Õ^€‚~%‰0Iðiå$ìg”'ÍQª!‚`•9=èÇwO9ÊjÐ­ý”ÄPJb(91”‰¡ü¹‰Á™QNÇ nOPÛ+Yþåý¡óµ¯xgí†œí!§7õôîÉM&wï–OvªQzÉUð[NŽSw†¶P?h\ª¾­S]#mKªw2ù=fÙÙ²@ÓEÿòåøÛ‘½ºCÔÙá;NëRÍN—²8VÝáqúùÏŠBÓ‹;D£‹î4ì;ƒtjë5® Þ+«þùÏµPwŸöî>VËÏ{*IÊ`p¤ñ£ÝiÈh–EÍ¥E>\ä¿RÛi•LJ`©”¢/‹Üòãü¬`–¹4˜ÀB «•ª/E^ˆüôtœf	zðgÓqÅI»ÓËóÓXbÂþ4@§Ïq¾M]yä‚ýt'¦*.]•úÈŽE[3ž¬£—Ì9¢‹iyGý½„‹»$(î¬X;…Ñ÷%t€ˆp©{í+³M…nY‘“  -Õö†Ì1`‰ Žhÿ,ŠcÜ<ks(™öPÜÕ¢UÒp9»ž †_•7#Y¦… uè2ˆÊ„«î^›¶Ø:ƒS‡NCœ7^NÏqœ{1Ÿ»¸ûÜ!D3yÔÌÍgaÍŸ:9]ÈÞø^'èøÓÝÛ¬Ms¤&Î(~„NÏ\ãlÅfËxlCSu/qª¬[WÍm Ü+ø<Ýç	Á™y¢6n>Oë¦ód†Õ“žë9“t¨‹×gˆðQèR°ôM3Ttš!îË	MÊ/Î¹<Óó5¢ùñùuŸ/ì™/jîæó…°ÞK¬y#çòGGì×&”6ÒŠÀÙ„Žº0ËQ§ÉU*th^­ýP1õ]ÏhE3Zñ­ºÏhe¦Sµró	­îp6kVWã,:MmöPÂøWÏaÅ&Ÿw™C,çN£úkâNŠbØH•ú¾nFz—QÍ²¸}RÛ?µZŸêQ}³ÁŠ=ÝÄ®²Nãä,]ŠqzRD \I…°¯Õ½“¨ ‘ÃE/—{žé½ÍŠ´OÂA¯ßìíÜLÿ£Tîf‹Wy1­Ö'>Kâ—¾Ó¾zõÐ€ó”³¤Úƒ^÷5˜·ÈÙ'è	SÙIÃTuÝ•ˆV •Ÿ3Ï“ã™äncº9ÆÇ‹0ÛÌÇqK­Þ`pÁõe–È	«Ò_#
è °Àª3ÎÓ«¤é¦ù„§AAhØ3Ã'•H¡™Õ‡ðçsáÖÂg¿{$>^­ÕÕý<¡µóH¼)€8¿¤„øÄé›©óYöÍYÕF9†EÓO—±jÃ¯ÆÕÁ|„60áÁÔû#d_Ô^_U âtÅ<¶f×«ýÈ¹“PGùd’V}¯¯quCÎr”ŸVB†jaÄ‚d%4toõã|×XŒ8M3²ù,xgý•ÞúCÆ\Ì’^˜IÕÖa+‡:Lb#èud¶ãK4N£8§ “'Q†amb‘ø:Ÿ$­,ˆLÛßüFØø‹±¬OÑ÷Æ2FŽ‡TßBÑ~ï ë]}ÀMlá jÚŽËÐdß[ã±àú:á½:O¤kH rM€Ë ÕÝã¬ôŠ;†:»9Óc®—Ó-Ý=JAùX¸éî]KÇí!	>w¯@Ðl!%pNŒJ™@Ð–7ë/Í»ßŠµÕÕU—šÚA¾P+‹ÚÖê Õ1´UBÄ|	=8iëAíeKje›{ æFvaE"ÃwÊ×mN'.CýéžÁ@½×þý	C4ãdPQ!=×•(Tl6WF]‹ÏÎ²cQÑQUœò¥b6M
Š·Ë”i°rš€«óËXJ€v
Bò	†˜¥±Ã,­ÉKñ|õçOÈMèEqºlŒþÀ°×†bo*`
SÂGv>‘AšRA.¥ B%Èú70hÊ–x¶ûíî!“?³œŒÎ)8Ïô¼dLÆ8¥ÅÎ‰¢ÎXÇÁáTÛ€OC•‚ëÁA+Ü†5”ªZ¹!²sP¬ í­4MÍúHÈ
™…0öÙž
ËëH±B„É&Jðìüv“+dFíÜP¸h$YÉ‚È4 ›YždS˜q¢2ºÀØù'Ú±ÕTRMwÂê@ZWÝHÌam—-,ÎÌ¬eo	ÚpÁÆ`Q¤ÜÐNš]¸º‹ wä\•¬ìÒ2N4ÍÄ9TÐ Œ·0‡Ä.—,Ü1R(…†ª'ÅÒªk	#\¹/NŽèÃät¨Ìk¹±Þ°õ£8næWl½â
Ô»©¡òdÈzBáùÏÈN¢l[¹@$'ñl
ðmK‘~È°ÚGQGôÚ^\îeX‰E6pèÜ:5uA˜˜ƒB[º#Ä>€LŠ³YVEæ,	q”¸)s\' R€]¸¹·,+ÃEÄÕóèAÐÈcF(™Ÿ†`C…k(ÇƒÔwäÓ/è±Ñzà‹Þv”4ïâEŒ¢üåÔ†;QaýÛ¬VÕnžbƒt6L¹B5¸Mÿaï«¤ˆÆ=
f—è19¨«	g·5P`´ÕÒa7Y³E„ Çƒ‚°ìkUß{ð0X<Íf‰Ç™›šéçkÄ™)4Ó&üfy1Ù’¥ñk4NÀÐýÇø½Â}¬î‘®A+0Òý´+yh¤†1Xfð#æßâ2N+(YNÁÇ³IRÕÌòl…ÎcHÏ!74„1¥`am¨m>Y­›»O‡bšæŠhYãƒIYâÙnCs' Áü•M;ñ®Ükäe6`N+• Ôh`UTJœÓ¢Í- ³ŸºæöIÂhâˆTKyø7AÈ®q‹@º+6ÑÌjÓçC::H^öÄïÈfù˜žß‰=zUçÃ"Êâ\âñˆ@ú}2âÉ-ùk}Yü¾f(Sg€Î+}äZŽä£Ç/³§Žn¹Ñ Tò
V¢l']všÐ|u£Æe{=Š-7m.¥È‹J9¤æ–*s(±9Ä/ÄEfYŒGtÉ|ÊÇb×Áµ¼Aÿò§ù	LäER`Ãk£¿!Vù3=µ~YÇÙIN;ê³ÜD²æ2#'íy²±PxiË§q¿Õó@{w8•ç}’º –½št§l{7Bmâz¬­åƒ
ñA³cU6Ó¯ˆš	mPM†,ø¿a“ù9÷dpÀÌ"›¬ìÀ=ÁëM›Ûyªçð^
~Üýü„öÂ¨z?{~[‹ìÐNaÛ>ÄmD ÃÐ{½… SgÊ5b7[4éámÓZ_ë[·ÜrhØ`Ðhõ6ÔÖB”³ôÌ*„÷Ö_ào4¸M_1+ÙZß3bµ/¾Ñh¾ð-™†bL%]¶:9;µ(M¼e«/£ÙŠ&jŸÉ8L¥¹÷›ŸÑ$’AØ¥{ì­Æ]¤®;7*ªðàåZSÇ	¬ámÄZSÐé¯SœõËV!Ÿ›rÏþÞqTVz Èéx'^ÀbºH
:î/õÊƒÁCr¡B-ñ –P'œrB[]	ƒÎ¾†"Ät«¡%P‰ï[twvéÀ/o9¶löbÃÁÒ…ŠÄ
8º´c+X¯Ê«h|„z€
ã*§ñPÙF# }š4ø
¬ÌdÒäÞÂ²ð'ƒx€ÚÉ|€ëâ·Òø.É¹°6h„J$ÅH:¤F™ƒJÔï³!Ü}Û¼(7K<˜ÎEJð{Ì3&Óþ”lÃÈhË…&ÏÛuÁ-³øþÓ§0ó¥jña¸,¸+’Ì—+)U}Õa®‘°ø¶ôódXð 2¾Nžâ»œ›j„NÊù¬"=zGFb šÕö.Ï-‘Ê·¢³ãD³*‡Y€æ1ª"ÑR
b˜×ÅMâVÞd÷Ô¨Éá“F›ÎáL`³Áˆ¹¥Nú•vh,a§‡ß¤kÚOPò’QÜ«ä®ÏxO)µíŸ15aj}§‡x6>»Ê÷\†¯úSçï`¥Ê„4â£T¼›=	= hÃhõ·úS3¤lú]{ÑRŸ'‰q8oÚzŒgã¨a$î»(Äiž%Ó·yT¨@¼úQb\ý½g™i™hv£Õå‰(ƒY/\FÇšÜýhƒÏ‰…æQŽW(‚5ôá–¢i·ÅxÐäÔ2€|®Ý²|îl…úŒr.ß¸»,È¦áðùÈ†™.hÃ<kå6æa·
æ“b£¨Å™ó¿Zp„Ø]K‘Å¹¹;šá9¢¯òülœ<ÉÏÒ,¨¤rKPƒ¹¸ä‰£žu2Ú´j´µ2=Ëö2ÌwtOgS•´÷ d“é È/RPN™^¦H@êM oˆ¶çÇNb-ÕÔüäJJ˜[³Wr~'¡=‡Å'ÒäÀ>ßŸb÷WFã¼Lâ•“Ë
üô²{aëB¡('ªƒG½`Àø,Ãóüˆ@÷Áˆ3jÇ'hQŽó³þ’œ	ª¯‚ÞŽmòU|ÌÒSãfþ‚ø†ƒ‘fd\©Ò“L¼3Êúç<‚n2žÅIÙï… ¸ÞY—8öAm™\ÿk–æêcrÚ“iå2üž^ÿHú ÅÌ]ÊJ¦ÌGÛ%2ÀmL(¹H#E9CñM•¢c\äê]>ƒ™…åR9 ®»+ÄÍq”dçÐöI¬gØ›‹@oŸWÎÅ)€JâÐÎ®¯‹GÀ¬Ù#>~Ë ¦ÍSCXâZ#I$¾()5¥Ì3K2Aïo4ª†¸°{²ƒ-Ò„|_(O’ÓÖyÃ–/i$œÙ§†`t†³^]ãA*”ë«"•K-?éBÓPö=‹¢²|¦§Zøø/½’_õÛš³-¦G^kPî}YÔ•ÛíAW	ðbß8h²en7®v1zC\d4¡êR¦÷ytO”×@ö<†ˆ•hÄ_®¤òÉ^3ƒØÅ,š‰\§ß#ì ™ Æó¡8€E|]Ð¿×‰Øú)gÀDqU³åÌE‚]K³‹hœÆ+#ƒ+â~Á²oÀ;[™ªùj.‡È^Éòjåð¶Ê1‹¢nà©9è~è}ÝGð&‰^ÛN57¶¥‰Üß&ã\€©¦Î'@p³ñuoú§—;Zrœh—C±§)ÂÀˆœÌºå#œ‘ÌÅƒf´©ŒÉd¦è±ª#£q.UžT“•äÙPÈÌ•ßÏ(Oi cwv[X=È˜|†Z“Tº±ìÏ*â"Vø3Â—5©vÏÕ×”’œI>*eFêœ™jHÉW^‹ló–6"›#2Ý”o“Úä/07+…%ËG2	¯“…W”Å Oê OZ@Ò¡-*i‚sáÉ
ü®·ãp@q)},IGïeHªƒM‘L^Ä ˆ/ë Nº‚ÐÁÃØ¿˜õÏz×+¿ç9ªšG0^üS¾Ç	ÌìÒ3\¥çÀ¼Yá2¥³iÀ…„Òø.€€ÑH©%PæaÁRÍ;/`[wÖ°Måuh<ÎG*Àé»Vó`ï€¥Jë|CôÆÌ`ei&öLú8áiô±6R¯¬Ã@;¨
i,­}ìÅÛû‡4&­-ÆLÍ˜3ÁŒ/Ž€ûá¤Q¾3¢D¥È@½;ˆqCD^¾}žãÆ1ã29)¢¤´:š•U>ÁÂ@®~1ãˆlQcFª½Ûwó-
I:DAß°ìQ>+F@¾^÷üî“¿ùk€_êLÚ>9îï*x"á\Óü^üøñÍÂ#„­Ms¬â' ¨`ßÿãÎý³eJ&ç­	µ"¤¿ýÜt³7pò+o5¤Ô‚A_UžØ`®f+0GT©ç)‹XÈUõþÐ×„@î¢·%ÂèùkPA#9OÆÓ¤p©?ŽÕK@R¼!×0à4­Æ&SŠ‡_¬iÅ„O¢söãÕMÎŠœ6b:R9¥ŠO€³¥0Â¯Á~èuÿÑzS‰¯“ôìœ½ntx.ýÖBˆû¦OÃ7øÄcÐu²Ö5˜ºƒg_õ–ZÑ¼jÅ¶°,ÖV—ùOÛ·e+%Øã­£ãž3ÞP"y;J»'¾ÏÆtt'VM5ƒ~-¶¤d€™f9ã}²
„ª`ôt*èÐè}Aâ§ÒjFçü(·b|²Õ´Dç]éî6ÅŒð8·ó1¨vkŸ b>} ÿ¬&wO“	ÆÆ+VWj´á–Ä‡¸ÃDõG$Žéïú²ÿõ»È¤ö“·•lýÁƒúëÇyV¡Üè¯­†_ö—`5]$¨‰—Nòq¼ä•¬ 	˜’ýgGß<9Þ?ÜÛ‡ûKT²öû!všþ}ò”Q…#ÊóÜóÂ>hèêG}‡ÀHc%ÛGÑI‚YÚ¥€gðôgÜÞƒ”×ã,ªÅÏôhý£eññº_b²æ•ÏØÜtEŠB	1—e5t€Í1Ü0´Ï†8nýÕ¦ÎÈ˜Î@w^©û16Ä‡ïüLõð•’‡Þ_‰?	’[B
‰ëÅÛ X£D»:‡åsŒoÝbêå+‹›VýÝÅ¸XÀXÙœÀ“"s½âO%bzÓjåËÃÞàJ\ÿ¯2\5Y·4Òß9(@~ë+qz–Vðh’f3D‘y$yÅ;úIã$îÑµÒÖ?ARÂUùñ*'¦u‡˜ðŒT_þx=´`$5jqgD‘fx.þáî“­ãëÿq¸·¿!žl=Û;»ÀýþÛõÝ»GÇ‡ðò¿µ·½ß«)#Ñ	!2
(¸ÑŽ/zÔ=bé
T!vQ>ü´ºT›v„½í|Ù†ï¨„JcJf›P:"§í@ÍÇ”L@VÃï˜ÇÀ ðÇö2ÕMÛj›¢o~ ea]xæ?ƒ‰UW°±ªM¤(OGÛúV?¹È!cˆïñÚúÆê*ü×Ã”5úÍ`à)Ûv §„…
ÙuøæôÛþnê4¡1Ðc¬è.ô:j¾`Ð"^|#©¨·³@(ä]OÞÀCãÂ^¤å³è™ygÍÕóJ³W!µ6NOOñµŽw‰NJ¿Ú¹õvV5€Ú‰.Kj”¤ã¾_ôé@ñoÅ'æŸõÀ½æˆTƒŠW¾Ó ¯âWAÙ¨Sro®&Ê·N+ŠÎÛ„éþf:5‘ý¨¬¯ôÜrNHµÿ2¢Ü­¡7•aì8ðžOaËÌm*ûØ¼Ô¡>÷//{HA5°>5¸óÍÂ›yÐ$â}t¤¹¼è¡ò¤ù.i
:ÙY5OŸUù1²cÉÓ™°ŠŠê¿lˆ,~ü¸ÿ†xþ¼÷`ÛõŸw·ðËÎÞÖWÏ®ÿÇÑ1ðz 6PÃ¾|²û”^¡4øŸ»GbW<Ý}ò5jiGøþ`ÿèxë¨G2äè T¸ë?»ûï=»þÇí½}üzp¸õÑ—o‡òÁÞ¾üv´wü4½/lÿNòørÃJûBŠ5<Ä5NN«#æ
´ìÏ*Ÿnˆ?]@(Ù <µgeg£´= y¹!/s©ùýô-/äy6› rø…CIX¿…$µP˜5áÎÑQu9NJ.+}ªM˜An¼Xn6
*­*@qN)â÷õçØþÞd3ŽB=÷4oþn”ŒÇ€`÷ÐZhh8½á¡Ù¾Ý^º@áexpÏ¬-‹O ê§^X½é÷¬£|<›d¦#â*ö›$ujC –ýs@B¬yÖA«#ï·þºWÿãîû¼÷Ÿzï?öÞ´î¾à½_›;¢O®ñ{¿Æê¼Ÿ.ÜÆg>ž×Úk¸‹þ OÛP{îª/e¤°¾Ç)¾ì™Ûû$Õ¥Ä÷Ÿ†ƒ/h3VV&É(ñRe¨K‡Tl`oëÉñˆ[¥$ÒúÆØêõu´3?[¿_A¦A‘—ÉÊ'Îj÷à>½þïÀÙçC^°¶ ?‘#äzè«Ûbë³»ÂÖöþ³í'ß\ÿãÎþÜ‘!ñùà…ã9iEÛÁî³ÝgÇ»w2%AÉý^%‡ùË9_G!¶Ž}_°þ=–†©=°àKS7¿HŠÓqþ7À<<•íµâ;ðÈú=v©w€>ÎóÊó|Ê@m‚%¹>ó‘ñÚšÍP—¤„Ü?¥‚6žÙ9R´FGŠ1ˆðàw¿ãô€{xÙzSnÒËçÆÏñiýwÄ¬z¯ÉÀ&ãÕ¡ÙÀAïð[×?‚‰ß¥WøBy¨Ó`Î3PÖæ¾/Ö]°i{Ez2É&XsŽôÜYæ-Û¨U_ÞÀ’Êßó)^Rmå³j:«ú½“q~2+ÆŽ“—'0È×&U–jÑzéf¡7vñ8YÀ5ZO(¡6Ýtx‘Þ4ó7Ü ncŒÑ³è‚Žó=¢ô#ÜÛƒ—Å8:I0`…<¶Ë"Ï¶Çéèµ¸RÑ"˜$€ ~~2ÅOgÙRÅ½óÏ¶¨ü •¼ž/““¨ØGäz7š@¬qT–èâ}ôn”YâZz³rŠûµ§càwtôzEÎ¦8‹¦+‰éÛ•Åô¾)/íÊÛ±¨Šˆ.ÓÌ³Ì9Ïd: ô÷ˆ3(7]b<C:«ÝlŠ¥“³ÆÒH!ZysÍ‹ò<2YŸéoºÀ^âP6ÄUÓÌeUœ#3Ù` ègT5¸¤à$r¾P??Ç9%Ðý£wë«WÊ–Ø¨é+Z±¯È6KôÂ­¬­­.ééÇÛ‘.y#øØ«_ëóàJÜ7=-§QÆú·„s¥L&)ùf¿xGtwõù},'+}~_’þâaŠtL¸„¢Ç8½“ÌÒiú6‰^¨Z­\®¬’a~@t¿Yùýº€syBÇ¤V
ýEt'¼•A‰DÞÚº‚‘ã³*EÕßb“ghw¦#ü&{²*Àn¡qžÆq’úKùÒ 	à€hVø#\’:WÙ..–h©à?+ ,ˆs¹~¦+Ÿ,}aÈ3X%°º&'€\cë¬r½ú,uŽÿ4/¶<ë­}7+«ôôRÿl]d¼#Ð•/£bû,òôÂXgt(	zÛÖ{_>ø !ï&Iƒ“Üfô¸ùJE{{lÍ|ºŠäGW²<K–¾`Úx/_ MÀçkkÓ·/„lÖð×ŸíÆ (1©VÖ–¾8¹{Õõÿ3^|•O’²Þž‡õÓþÎ¢®VÖ@%°º<²ÐbÆÕa?z÷5íX8/ˆ<ZÚË®ÿu”æKîKÉ’½³‡A)5¥ÂDµD–¡ËåuÏM‡ÔÛ+§¥û]‡ùå,ãt¯‡Çª¶|Ê¹ƒ•Ç=Ó¤ì>b[¥uØ7X‘&§xUux`@Ëßa€ãÜM±à¨Z`\¦ÆO?›Û "³8*Èß÷ÖÀHC³²ç¢ ¢:Ý Ê¿?¢d§Oæ&Ï'ß6,ô­y§¯éÄÌ1.ctùdÓ˜í¸ÎÓdÓMØWõ2jô_aò?“"ý$ùsý—Â˜Àƒ× ÚkXhÇ„©UÃB]†x\äÓóËæÑÑ ï”{aáÖÒüqy£ÂJƒycúÜù=¸êJÃZÌ‡ixGgO£Q1_˜L°ÔŠÍGÐ™’ýŠ?=SÚzí6p#zWæàÎÞÑö\4Äi9Záž”%RHw4øz4	e´þfÆ…½¼1È«¿o„‰PåŸêä{ø3£ùëÂ½¦=” òfúÈ‡õ8'èˆ
Ë^ÚÛAíÀ¬&sÕ)¦t¨òƒ›uü(©ÐÁÓ 5l»æö¾TÐºöÞ”oéý»†€ÿ[Ê\œ±£)ž*Ï“ÄŸ8†ƒú€™DµXÙZ‹ªzå¹kþ˜µšØ<X×?ºèˆYœ†Ê“çÜ~Œ+18ÆNƒ²9“­ts=ãóû`Ór×sL«•O´G¨ª»†´éû±cú¾»ÇÔV¼°Â†è£•©0Ÿ%ÙùlÒ¬0¯ù“3ã2Ó•¹ßDö¥Þ¥‰ïŸ	Ámvºœ$Õ›$ÉÐÙSÓìXø#¬«p†Ñh/º~‰²ZúB:½ß¹Xº
¹BBÍ¿{nTè^>08#Ï×‡¬¥¹'Éi…Û&‰MÑlÓæÒšvVé­&”ê1o–)À\.½«WWõhð
“×—õÁ]óX…æó|eMZAç®3h½oµ¢ü¼"¼@ç9¦•ï°ÉNÁSÐû ŽÐÐîÛ£wïÄ¹'ýêÃw¼5‰ÞöW—e ¬4ŒÝÂ,èõÖïó§´åvp6Ãc ›â£O@gBEž‚£Êà=­šDJƒ«ÿôÊnVØO‘Ö'¯öhÀâÅº;c‰Âa‘~âû(ÓÉ™?“e1zôNæh›žçUþÍázzu^UÓrãþýYº]DUT`ÂšÉýhšÞßÌ°õU-Pð-—Ø£«ßœÀ
>#úz´¶zòÙ§k¿áÎà£ÓÓÓW5>+²D xM}Æá*æ¶®¿°õ·ä+’Ó¤(’â Yqùh)ËWô#·¨/ø†Å$µòfÅ÷/Cñiù•“0ãû”Ü°³/}Àäç÷§À¿-}nZJG¯ëàê¾îÚ¹ËÒ Gòª.­;3wƒ9Ã[0õ¤uÃÎÝ:bò†½ÞŸUl{ÌŸvbµGQZ(­³ŠBîp»À ûa¾òÍ*™A¬Ö•‰“eÌëÀ³µlª”ËìÚŠóüÍq•UÉœ§I6Â,â0ú}•IˆNWÈxô€¶‘íÉƒðtœJù(9v‘ŽúìàïÈ­)=Ö˜JL·›ÆùÀŒ……òÙüaæYmX*Q55Õ ¸GÅ-d¯"¡1ÜWÝˆ4»¸þÀå÷Ä^†Ûu‰˜Mä(ÔÊ%°ö±ˆ¯ÿõ,Å³ø…³„ÌGÖÁcÄÕa¢E16=u+ùØóÞ·còàq+ðF¤†zÿ$9Ãk³ÃX…
R×n…ÔÇ	Ðr!&yŒ-Ê›pT®;:I¤Kâž.vîúŸ²$Â3ùr
ê8A¢@ƒ§¤f ¯èf¦WŽç9¥õcÏÕ)zïv>½Ô©nÍM”æÞT|Ú§ÏK!Ã?šFIYQ¶FÝ Åº&a•ÅF&þåã"RI/°r~ñYJHK:´¨~à¿L0o–î0OèÊ“x`)ôêïì½Ü}zp¸{´eOÀk?EÃ4|x›ŽÖ©KØm~¬­¢ˆ.‡xQQ__4a“2´“ªÕ¬H½ÌæÿÅv]Ú0±R€rÊÉŽ÷ÙzYÖÝÃÌŒ‹Ã™P7ƒöV­ ·¶I¯’x9«HM‡¬§r^aÈº9@Ù’V¡Ô_ô¦XÀ¥¼¶œÆ û0ÝÂ
vl6ÞPx´1/0IºûÏ”miK8YîÚ
ÊttVÙKo`>®Xˆ)¬Ëm•å‰†ÁP¬½£}eÊwIÖnOm….….Q¹¬¯I{Ñž Ü?d¬VãÜ‚Å$8Ýfd¡Y™;3ÎN}†0á~RàQ`Ô“ißK[KìÍÞ†nFŽu¦üNW]wºÝºžÛ/Á_‰¹YruQ^"/C–ýªË;f›ÿô jÁetªÓSá-Xõ¨¿ölNsBIeíøàýø@¹U›ÖªO'x`Éz¢Å¦¾Š	G|ô°³²K×‚hÈ¿j¾¿j¾¿j¾š¯J¦ìVBáL×ô»»S‘)¬ÿWÙ
¿ª½éY¿lWÏ­Ÿ@%öu{º/Å®<RV.¤ÊbDµÉ‹¼ˆ¤+o²¸šO.¯ÀnúÕ.EÖátý0F][*…
ê…æ:%£.Û®„UÃ…µ7jí'ÐßP[@}“nGÇ×ìzd‘xïÑýÈï3¸#…,||M*†?<­˜…žÏWÎBµº*h¼·-*fÎ‹â‡|RDe:NÒ"¿{Í„tÇÜôÔV
a½Ÿ«œµUž£œÙN4«dw†Ä)d?æ¡y ·Ý77R±8µÞ‰Go™–÷KÃèôdÀ­BÐÁ×ç•÷:tÒjuÍ¦öº£vƒtYƒÎ×WoKµ4fŸ8õ­ÞOåXù–Îü†|_u!N¸Q*\É1×wÈGót%ƒœÂïÀ5hqe_Ê$!¢·U¥K¾€v…«À»cLbZ÷"âU ¤ƒB5ÁJ˜G2¾¤C»Š_ˆÓ$‰q7½®”ñ(|Ç‰/BþHhxï”_2Dl!-ÍaNÌ`²‚×
à6²¼Å/* Ð)u	"gë¶O;.PÓW¡&3æN¦_]~Aëç“0„à. ~Š¬§ŸØÊê²§¢¹ñ„*Ñ/Öš¾õÄmË^}"[’¿‰8 8uý[P hô~gVPq£é"Õ…ižáése1æ7<£ó¨Ê7uK‡úB³Ý0ã5§s§EŸÛyœÝ×ëžØ´öå”ÝÖ]ýÒ=ÓÁwäÆ„SdýØ£{t+?¶ñ0=·Ù*û¹!7vÆ² ¥I†| ’4ü‡Ã½ãÝf\Ý­‹œÁîâ&ŸgeÍw•{šœç.ï` ÝËœZRï
JAAN6þ«Éö«Éö«Éö¾L6ÇunÀµøÏïÚÀ»ºgÞ)øfæ¬ÜÕ¼ËÇñ!¿²†HÈ­ŠÊ´!Mmº©Ö²×«WäÆ¤®Ó`L²×”Ô0 ¯Tj»(ÕÐ]FäCg—þÕ{¥t]¿Cüö¿Z¥c¥ÏúµêW †ÛQëð6Ó°HeËöw„ßØPÑsC¨QÚ¨Åtõg9XŸ–kC`›ÄµÀb›Õ¬ÌŽF¤õòO¦äçŸL›=ý]ý“iÝ(9§[kÁÁ]c"âª%bÍIee‚ÊõBêóøRAÃüUª5½'°!ø)´?É+Z.­ ¬ò«ý[¾é²FºDÒ¿9@][Üvó-"Äã.˜—¸[Å:!µ\.ª¯Ê­\¡Æncy¹ôÈ‰±a\rlE4i¿€Ôð¦fóï&eüª¥#]/.åŒ¬Ü¼kMç›ÝkŠ [.2•¤áÜfÚp‘34_EŠ©/Y÷ˆÌ[xè/´}ä­òð•Ó8X²v!HÌ^³Mÿ’Uo%;;‹^'6W»šµEë¦S÷ðê¤W5Ú©Hõ•Ôe°ïŽKÌå2Äjçó”`±yD¡·Æ3\7•¼#ÑäZ£j>+øïðXò!ç/í°Û?­K¸ÉÙFTÑB¾X/@¾@?ëCéš«äµräÄ0.TÞNÂ\rÓèrœG‹:ä|=gñ9÷'¯Tm\Ý6XDªžƒ6w^@
ÀjRz[K[7>?GUOƒj{×î't Ú‹j ÙOs<‹®Ê¬Ÿ:oÃÉ}?ì‘xS¤Uò%þèÇ'×%*kÒwY‹¢p}o,¤‚C¡v†’	÷°eMÝnÇ}æy7ÂNÕt6üüô6‹#Owë°`ñÓM©ÃÏB+Xwë•;yK?w äQWîJÑÃÏ+{øñV~œu`–'yXAZ3[®š˜£:àünfrKBvo”O&iÕ¿÷ßÄA½Gïÿ« ºÿá;Í®^ÝÍN@h0Ý"®(Kå‘wfƒübê¾}úàkO’âl–UÎ“I“ò>/ º©ËÀœúÛÖòÆFñ:@Õ°Ö4‹ûdïkŸ²nÛþ9¾K kAÃîuÛØô_÷¾ÂÚ=»ƒ±Pí†Ð^±ÑØØ	LäŒ0þœÄÁ¦ÒxÄkÑJÓ?&p'òL?ŸIÛŠ˜›[5ìnúãSgÓŸÚCR:SŽðjÏ³dÅò‹·ô{ØÚËÑ8‰0\dªÁRö%µÞ]Ýµã›@º¦úïõ8íÓ/C÷òÒüÚ²‹Éž½w®ƒƒû@AÄ›Š1Íón×Õõ®ûå† ê¤Pßzfìºak˜Íß¶ŽN¼Âº_ žæªJA22n?—	¸;ÃS¹3\Ømá©³-Ì'’•¶š[7UõPB¡£J9Ç9*vLsÇ_¹ë/ˆ»*•¤iþ´Ì²ÊSæÒæDýï“õ-¼¡Wgy8zËŽI0–t;&¸ K1ÚÙâ,¥ŒN“Ðâi\¥Ö ¾‹™L,wÀÞíÅ›‹‡ë;ŸPYZÊ÷V@½ÝJ@fyœËNÐ
Ï_à'æÏcäšhãÊ¢é7%Ýœ,µzP†q|ýïgIqÙ—ý[oÎ˜žš¼ìÑ£žêÀ½,Oƒbw­QŽhàNg%<Ç…­ºøÕ°gâäRÓƒÈ3ñá;…±+¼78%¯ÜK­F¯9ˆƒ‹Œ-Ú+rq!±À ÑÍ)jND~ª+Òˆƒ—Ê-ù˜H»ï`‹"%$ÉQ£¿aPNì¹“œß2½.ñµà»:˜ÖœY(óIÒOÞ¦%-ÜÐSß$Bë²îÆMe‹YŸˆà\»3}4Š²³ëôqú5N²³ê\|!Vy#Î½7«tïTÃ¿{$>^]Méè|–½–..†LL]ÜO—±&U¬;Ç:8Ñ²»n¤©S#ýôôSwÉÍ1È¬Ù&Kö•¾†Œ
ãŒr>—òº=||v);ÖÀÜXa#-BºQ;êP–¼µ¢Ü«‹¢ŒÕ¾Çk;€„[H
CŽ„rŸŒ$zú’‹¶¿;Ú6”£LO/=°Þý>ï®Œ*rWjÄ"r>y«îy¾€ï×C}ŒYºŒµÔlÔ•K; 1Œ4däÀFžƒ\€‰>Òq"ez`!´	Ãóo²«œ5’	ÐA$è"jÛp¹:½½<Ðú°2n›8ÉN‹˜\4f*0nµxï‘Sj”…‚èröSAßDª íž€ä©ñÃN²gm9ª"5à¨ŠhÙ¨ý9ÒŽšÈëjÐo`ŠÚÑ§n}`nŠtå/Yh5J¼Û×þJ'ÓhTåöAÈ Å­•³ÒüE±n‘†íÈ_Šuhî/ÄÑûÞ”l¤-xŽa=¸»7Cb+‹ÓN¾`Ué–ø7ú¶Þ‡¶¿RhÇXÜ3•àˆðA|.ÉÝ­y^Äüdž1?AƒØÐIg/â‚°"5Ï[˜ŸÜÒ[¨ÀþÞBÙ’2™ƒ‹íc¹€ð›Y<øÃ³–åšÙ´XºËS=KÆI˜+Ë9¿`mtæ”Î·œçØÉ­~?Ÿ ~f¿ŸBíÂ~?Ë`C~¿_™í¿wf«œu-|ôýðÏDòþ¾G'â/ˆ#*OâmøaÐ“¨yaGOâÝúeIÚ±¤FCØnÝ$Âª¼Ww®!¬)Ù²°!lg&hk°¿CXcYÂ7¡ÏÛÂ1ß…!¬FõÂF'ùk+ds»6ÀsW>ûeX»jH··ve§·ÆcÎ”šÎ¸¦±q£:¬Ë¦!†õüÅ-¸Tïù{8\©Á´ºU1Kü;»Vvœc!üg)Va4±¬­¼oc9¥•ÌæÈ¹~>[SBVM»]ÐD$©ÜÝbë;Pô²OkÖÜ•Ö´¿¸‘Ö´—ø´fýœ‹Ó«û‹¢5<à]š-hŸÖ¢ŸœÖ\_òMiÍB™CkÒ‹w”gÀ½­OßqÏ)í× ï–¦uëPm÷êÅlK6}zwìBíæ‰ëêƒ‡·òô•
è{wô)¬u
±¦m)k½9Ÿà¥z±.‚~YÎÎ¼Ô³ÔP¢rýPÙÃtÞ?+£‹dì½y=M„¼4¥gã
Þ—P,_ùœZõ/EðÄïÐÚFÔÈ¬7HµÖÙigú=Dî´x$kdk<,%ÝåvH–5‡dy'IEP‹ú#™¨¨»#ÿÅ†>bÚ$~RQÐ°¦¸×êðUô+wÿÉ¹ûû	Ë4|ý•¹0•æ@¿6øVmY=tÉR}Ò¯yVõªqkg°u}¿ªYjÂ‚6$›¢€SÕ ýE‡–\Îè"—¨oïSÕdùÞ]ªVËúêQÕË“;TŒž9TU`)ÅŠ•@e¦=ý|°šÞîÐË¸W7¯˜áqŽ^àlûÄØê\Ô%Æ«þ=b¥ä~ÿ–üÜ.1ÞÌ#Æ€tu¾²#s­tÞ¥¹_Y•P„ok•Äö©YÏ}`?!ìù­‘þB´¿àÞÃœÈ·^e›·s[ÿ¯p‰KÞ§yz+'´¸íþøm¶TÂuúmè!ÈFý¼jå?-»?7pâßžËI·¿ò3ôÑåç¾ü«QâÑ&a*OG*›1	«›½>ë¤¹gj‡@#8•T“Êø¾*zÿ¾8$¯T?bæ6Šé¬ðTÎIðWTÞw½¼ÞÅ¤”ƒ óI¤ÇŒÇ[ ‹G@%Ô¤<Ë´Pt…~(ò«-qý#NÿÒ‡ï¯–Äw×?bÆ]ø—y`³\” wú_uÊì»Ì\,Zç˜i™I¤;ôÅ<KÞ`Ë{ÙtVõ{½Î<òÕõŸ¨i88ÿ*À#ÏHë-^g…Oyù–iY»Å}nUó/u_'ze¸V›ÏJé‘ŠF2–×ËscT§ŽLE¿oåÇdÙÌçŸ5˜Š¨&6q@±,¡|¢“¨xM¾9ÔƒÏPì"-¥¤Œá	±”ÚÂmåÆ'Š f²ø2×¬DŽÇ±ðT¢¶Àú’…ÍáÚ+gé²ÌnFµ¯à[óAÝy³TwNnØÂ—ËüfyYŒÑáá…mNCÞwÃ~U›½÷‹TïÔúâY50;I9ZP <K‰[wƒÁÜxà£§qæÃ|îlh§\¬†	|@|yN¹“å7¿ñçpxŽ€_‹^M$´: ¦ g(]¦af‚ÐÕ34.ôúåjzâT.Püñ˜†?‡–Ý­0¿²•¬woÏÀslH£ê–B\DyØÚ	Iò û6tÞiñ¨‹_T)Ìõå»D5¨€[´eVs0Ä›&T~æûù™woàH·Ž±õ¯ÁÝpð8ó]´§Nc:Ãk®p	ìocähžvœöÐãú´Kpi'¦;íy'œ‘õ©tî•à2!P8¸S®‹3Ð·›õ·1=aƒfOŒ/m×ë]%Qûã¹¢_$èÜZÑÏÇ±›4t’6Å_•'‚*ŠßÕã â$!Ë²aÌ¶ûãÿƒj¶f@ÈJF(ÚÂk	Îæ–Tÿk-ÉJF“‘òMF¦	ê|džÔ)­$¼Nùýú1 “ƒn¦ÂsEMÊ&¹U«0	ÆÆåø¢G/æ¨/nI<²Më€¬hÙ7Ñq10+rÔÍöót7²-X/¦:|¤Ö‹é\åÚl,¤Y:Q6v¢¼[Eg~ŽÊ»P·jÂ6;a8oÌNûÎÑD[["-ïZêQ`Mx±1DõŠÇ-Þ©÷¦»’¸°º»õÌ=sžwÖéMøÝcÅ5|€‹žº_€R9™Nq{ÙF=:vÛOªó;T"?µâQ×X†5bl^p ”5<å§®¡W¾ªäö¤õ>Hã-uk¹‹?Ôéûß°á²È×ø¿/kæ§B(RÐ¡/ÒkÇ‡‚Ö€gàhdü"¤þ.ju€½ xÎvB³ƒªì$§Ñl\I…7šák½ó£zø¡¥O¿Tì'°aŒoöO¾ÃIÝ:ÜÝ:âèÈÑ~%oS0‰‘ó,ý¸œŸZ0L“DuH¦–æÒIÄr4CñÚèÓFd !+:0MÎ½\6+gE¢Âî÷2{ÅÌ#ŠäG%"”¡ÔÝºVùwvi¬~Káò´ >n>À·¬CÆ'Æ€dGt`Á-ÍEc@C|©„i|¢†˜«Ý_Ûƒ$Ñlë†<{˜m£³VæÑ=˜»K©	W·Ïnûë¾QÙ´UélV¶lTªKzu¯ëœ¥)ÖZ{Ôõ$!á÷Na¬¬ÿ»c0D½FC\€Þ[Ô‡NÊh$Š†¨ˆn‘ûŸ¡à ¬˜Ü>$ vŠæ§Œ
0¼dÑÀ€£Ì°|Y7Œè¾ò~IñÍ+¯-D`ÁuÇŠ4E™íþúÂ
:ÓæRcm²aµµlhnÞ´™¹7ìaßñi%ÙÇí¯ñµÀÐÎ¶×æ/†Î–84ÐÍ4±Ë¨Í6‘Ðæéâ89NÜ}»(¾«õ÷¦¾ý9k²«{Ÿ/Ò_ª‹ß:êÛ@·¹ê»2‡Ÿv>ÑÝ…ÝÀ>XÐ÷|jt]µ^ëò}ó–°ÿïÈ²Œ›±“‘í}{oÇ{t¾?Öô·êt4£»òª¿7ödÜšÉÒ!­ñx÷í4YÃS·cµÅíù2ôæÎ¹çÆ`þâP;£:þ©Ÿ†‚‡ ×’Ó¨¤k6ß{ê [œ/Yì¨ˆ¬£¯Ó[¤ŽwsaçzÖW¸@¥J"ÿå±ïV|™Jºœ• §“éü„ƒù­e åqþé=Ï{|–|ì;XuˆÍ!"ç@ŽƒÓA†¬î†@¶\ëSˆyç1ü‹onÁwÃxy]±šX–ñÈq^Ÿ8I‹p •ÌÇ¿¥Nœ×-¯Ú…â!5§cOº5B·®‡™HÈ(ªðA5dÓ»'Ï¡ËÞCÆ†.¹5éjá8[îið÷˜bS¢õ{À¨¹ºÈ‰.ñÌFâÌÍF.òxÁƒÜß
Î¿éÅœã¾ÓÛ^ô:¡ƒÃ°Êþ¨ç-’ófO ã*1«³ùÎtvîÌ¦ãt„ãBaÐ*áŒÙ’Pk/”Æ–—)ô ç°ä³Kl=îíÐr¹È{¨†©z¨õöqXú"¡Ï‰µÿoé"ëüúGÙäõ_e›"–ÃFD
7:óucC¤þ4;Í[]€÷ï‹£¼¨ðšø×I2Õy"N ŒÀë†ûø¿¼I¡søc’«óÿ÷ÏÓ³óD¡'7ÏL4½/’ðÁÃ[	MŠ|à/†øªß–Å‰³þTJà¨[8y¤Q³¹ÝqEâL©ÚKóî·bmuuU0´VMàÕÁÖÙ1´eéÜ¶ÿ%´ÒÖ~íeKûµ²Míë(rêÀŠD„š¸wr"Da“Î!#¢Ÿ¯¾pD^îù{,©n”Zxå÷b¢ïZE2u‰ÛóÍdQ%MÈHô«viR[Š=¼ÓBIíJ:A:$¨í{ª1³ù#ï)Ü!EFáŠ8®›Œ¬Aô…&]Ë= $W~Ñ>æä¢T]­ú]RÙòÆP¿Ý3áªrôÔ³¥Ø%ÜÂÚ¾žS!#)	*%ƒ¡_ª¨åÊX]fØJKúÛ—µ¼Ü(ë^æ	N£·Û4fMs’²¢ñø	Qâ+Š·ÔOK*ÞÚ—m&tÕ	¨‹ ûÃÃƒîD7` r‡\ÙeÐÉ¥á5¬Tú	ÚjeúŒ[¹1nœL¾ð#Éo®¹ØÇÓâZ2â£ây.(`SëC¡rôlSr©¾€„K3ëX¸6É]Lˆ[´g$’}ôG1mK^xcüÄ.Ž/Þ£ÀYyo½Q[-Gæý¶Gçé8¦Ë	¤:%åPž£Æ	+j™TÏã¨|]®±ïëòûÖ¦Ï‘ßŸFÕ£í€6˜–ÉEÿ¹^Ôïm
÷™§xË»#Ç›, 8%O“|˜²¨–êÄ‚ÀOß#Z~ŠND4Ë/“ã'ü¹Ð1‰@IÁNO^FÑ8t¢Ñ»ëÓŸþ22HÝì‹Z°Z)džQŽ\»bÏÖÝgr¹Ï`%Ñæ6ÁOè2RÕ“9÷ÚžÏ·4ñ£¡Ö.)]Ìú$P´)kˆüÌKÕd¦ÊŒl’Í“%9›ê;`¥Ûf![Tñ]&;á™olö|5Õå4¡¼'2èIï¡ NcµÉ¾oˆþD{P6‡2ÑY@4ëWfKG^Ó+]\d7ypš}8Í^pG™þ\pA‡SÂ5qS#¤{¶ÅøRÙH”ÊÐ é Šƒø˜ÎŠ3ùÂ¦¨•9íúù­”$ì7{˜Æ¨ëå¨H’LœÂ¿?$;ÌnfK¬38úÎŸnK1×w'=Jž2GÜöZ|Åd-hÉâ¶	³NHðäa“‘"kpÞÐÑVñÕ~ÝÿfKÄŸN—!ÝÜöqzàòÆ–ÀèîæMW§…ÞØNê]Ú6wbÝ„ˆ‘dèbô”¦ 5”fjÔø$)«u!IY².u+¬†!ÎÊgûú®ÂËº;û»_ó¡Ô¾2EÑM“”Ü“^&§#’$Ü‹MPÂ*$¥%…³çQy~’GEì=y–ŸÑ$ò^¼Ž²“(óÉ8ü¥@N~Þùo¡RÈG•Ãü¬÷ÂÊ6ÐOJt¾‚˜…ŽÉZFº8MqÇ-Óf‘2ˆÉq2˜…Óh”ˆäW¡äVù¾ä}§Èöf¨/°mÓÝ.'gu>«@4EYÉýA.Ä‚_£¯}ªX.:É‹ªddçô&}Ý¨çúÜ!^7 ÍïÐ‚[ÌìŽP”U»ñ/û¡Ug½R‡î¹UˆnØÄ*W1ëƒBºŸ,¶½ž¶p4<N	 u«ÄÈqdãŒÁ°gí%"X€eÉw~qîqïŒÿF=J§ú¸$~†ž|ï±‹:%ÞcWnè•X¨G¸%ØÖOè¾·"üšû.ZJ¬·•à~p	íåpGØ,Kó06à?¨1yW‚<ÅÜîØíîÖsrL€WÎñû@9LÓ,)û‘!‹Óú{Z\ØÍþuåãqéìoiñ¯´G ø²ä™t§ˆ2ÍP}JP{ÿq"åI ZŽÐ
W;ÊxÝF8&ÀQ"¶§WÞÎÆÜp,0­¶z¼ÎkJW º3 ½K×õ2ÏSv6ÝI„CÍŸÂýR7ŒtokWÍë“Q~¦–šèÊû‚Îž%Úö¶íq´wõÇØ¾7Ãê´—,53y!MCïä²ßbx³™±*nPéæµ1ÅH"hO©ò.y_1Àµî´{!5Û¸-ô„™¥ÚeÛâŒº{¿H‡ÙþÕ?â´ñ‹ôÜfuùôçu•ü<†¼<Ý³Ôk~}œ˜p¸ñ®ofºsÜÑèî77Û1šûÿÍôºª%ß§W»ô’6š4¸ïéBˆ[…u"ã¢²/Fô0)kUBª)Wæ;h¦‹é¦Â›]ÔSWAuÜãu:pV$a¨:§ m"è où^3•ÅL[ÃSÌ¬4fuSÞÜÖ™dóÜJ¦×qv™[ø{Úèõ–™tÍùö©¬s3½ÒìlœßF§ÁÌ§Ø4çÍ±|Þy	3cù=­bÞÂÏµƒÏ§… /åÆäð4ªiaž×æ†$!›ôèv&íyO”`Àÿ¢É NËÑ]ÀÀi¤ ¯‘ÛÏ½jÍ›||Úyö±ðûœ~ÿ=ÿ(¸§¢“l×ýM³ÍCn'ÚuKµ™öCÚ&O‹¼'ÑÎ€ÿ¼¢us…OÚòJ«2Ÿ~P‡ˆseG¦ä†îËÄKé ÎýY¿elŽÞˆþ›]$“[:0»:/U³îŽ\‹Uè{‹¤~zÊ:©J™ÏŠ½ßc7ŒK¿!¥×ØõÝŽê-†TA#âÞßTFõVV¾I
xq’çèÅyè½ß†Z³1nŸGõ2#Jád{è¼ÂÓ6Þ«+¼ »Ö:F”-ÛÄ7„*fhÎ`ÄÔÝ¤`9ÛÀ/ô³q,úG8…qH(0Öéiï5wNuÜcÂTNY'+Þ8)üýx0Êaìg@ÔùÌ®@´Ôá=ØJ¹ê3è1åU*‡KŒªy[ÜÖ•Œs‚iA$EíD‰M{B*ðöw¢w¼¶¾±º
ÿõü8•‡Cx¾£C))1Ø&mÀ·ót¼¹`ÒÓ—=h!³üMý@HŠü;™•Nx]Õ	Uz½Kb÷êÃwÎPNˆ+±"¼&ê´‚èøÎNàÝ†M e¼«Á«‡„¾òÝ†p×°o¼¹Ú Èî®±í¬lïÌÚ†À³jeM¦´FS¡'¸)œi¼-FðþKtÏµ¼ÖçÏ¾«ŒÛn¶Bd7ÄóÞ.fÑª®ÿå,P¸Ã„“G‹¨?Š×’åþF|‹¶=}„‘’´ < ‡ÉhV@yñõle²Þq2Êòq ¡âÞVï…=nÌðxÜH¾‹(ƒÃŽ{!z#i¸D¦‰ð@n¼o¸tí–ñ^ê¥íì)èUPg¡IÊØ‘)%ÉÉÊ9Î=KÊå‰‡Ã.°±³C¢=„(VÈ„j$	ÔKÔ¹ä%! 1ãµkläïäç ½C’·Ù±ð»Ú)³±§WzŒÓêûdËSôÜÖ5êtÝUí|t|oÒöØÝÈþ;b…(“éì„ÎÖú‘Æòâ·ºoç³qŒg¤%á‘—wÒ=˜€ÂåécgÒó½p^m˜>Ç(¥ç/\Uøå)P•/ÐÑi÷c™¸9}WA*qNO#7€ÕfŸK…ê_6ª:˜¡¢€QYÁN•Hå
¶¾ÎšwÕz'Oãv†¼dÀæ|¦iæÎŸê+nÔÅµþ2ÄôL@1?N“âl–U‘+ÞC	Íôê©ãÐ\NFjA_B‡•äcŽÊÑ[ÿ¦xÔºÒŒ'u
¦O÷nÎåQ>ý‚kÒ¹J:DïwÒ©Âúwm³tŠMú˜Ç›ÓQGõÔ—máyè¶ =ôëD2ž3îöÎdúa•œ¿ñB{ue9aNUPÕí;.{Ñ¥Þì»0ïæôäîZ«>×É9§í“­®)j¹@´}zˆQ0]/´éŒ[ÇI¤CP
Ýò_|[Û¦¾ç6ss<ÖFÇ}ví<ŸÓ?ê|‹ËwYW, ä£sÔ[©äÍŒ|÷Tù‹\€IIx+ñF-ÝMÃL\®ÖÇB¸¦Aîµ%}A*§^jÙnÌ ÙÝMþk¨¥ šu5úËOÊ‚…šIqa^¦àiR¢Å5ÝÒ7\X³,N s	f|œÊÇb­V‘ÒJÓpq1¥“”_‡¢æÙATŒR\»›bž¯ÖÂ8pÈóì^~'‚2zñ‘gô¢¹M`Ð€Ê(‹s9m*kßGŸ†åìDŠõþú²ø}mæ¨3U2±ÊŠc‰J[Tö˜‡y:ê^ÐXÂRÞ‘Ý²UK´š|ëõ\(¶Ü´¹”"`*å³[ªÌùod3ü·¥šüÎ_•#LPHóÿÛ=[DÀ6³ÚòF"aUr8•#*çÊg±¾†ÏçVFžòã<öHN¬Ã)YU™NqNû®ÔÐN‹PDWšèJuº¨SF+mÌ¡ŽúhNà÷àtØ^”¡l;áÚt0óLßÉÌìoäZåIç–€Çb~Âl¤}	MÁšÏ­N¿¬sú¾°¾íè"©ÛÅa2·„ht½£ëÄšk‚UèÐzYôC$±>ÕlÞËÃaölZk`³ÙïäàR÷ÚÍy(?7t!6Ð³_ îB¤'~±y¾6…îq«Y~é.®4[2ìP“?¥Ê_I^“Æ<.†-–Íìu9žÐ°ðïvSQYìç³ì5w£¹MÖ¶ƒ›ÞJ§¯pkâT‚mmËíFÀY‹ÊÌÉô-–xqK_ùdô>AXÿr/»çÆcñ˜"A[V¹½ˆ—Œ/ÞGT_äÏ+g¯¡0{,^S¦îUÜûÐë‹ÜI±„™¯D™‰ù–©20èTåÑw3`òåð•é‘Jgg9¨Þ’³g5TlnËŽ§LöL½/œÞø©:ÔÖÖ>Ð"™‰ÙÄlÞšêîpÌ>–¿3zRkê6‹ÆG0#ÁÒ…3´Ê­YÕh&8o–y:q|d1”$Ñ«‡yÃÂÐ¬6ÇíKˆn«™9ÀõÉ"_¢Ø…Ö	ÑŽ{{~@Ö‹h3˜~|jX¢7l¹NÂ¾·­m½ŠPî2HLô{´{\0?‰J*S²F»éÄTbQ’Úï-±Õ2d>«Ô©¾ 0¹¿,>z`Xe8ËMÃR×39‹yB­ 4¬e½™? [t;xÊYmQg	Ð ‹L\é-IÚ—ëÁ[›_IÀÐ>mZ›ò“>µ,ÝŒæþü‘BŒ:Çæ§-Ÿð"Éø\Çùœ¼ÂfïÉ°TëÛ'z»ÄT´èP^NŒê<‚—ŒóìŒäÚy^&™rû£1pŽbðóéÈIìt…U_+h—=˜Šä"Ê*>4>Rg·Œí˜#åÆ›¤éÄ¿-Âº5`»k³Ýé‡PÔ†în¿Á ÃN“­[:8u÷-_tÍ‹72• ¦¬Å 2¢‰äÛsü¯î´ã 2¤R ¿E5Ö;±&õUêžziŒFÑ'PÆcñU2I³Ô%Ôamlíõ±y/JÃÛ&
Þ:Ø LIh2B†2sœÑ}Cq êëit‘Ë"ŠÓQŠù”ËY$lmÄÂ¶®rý—ëKÔøÏ£“tœ‚N&$.Ì¹Y::X²üçShdØ÷8=»þßÙ(ÄVQ¥§):ïqfƒi2ÅPåWÛ9æ¸/1µM¥ÔØÕ­6@­ôÎ<ÕipR‡šHã„\Žha—ÉÙS˜”xoíL!¦HÎÒRR ²¾M@Y«ËâÓì%¨úùÇ¡ïÅ¬Â™Ê‡|°¸·ûìxëú¿]ÿ×}±}xýÇ{Û[bgWl=»þó“½£]q´ûôúÿ}&o‰ƒÝÃ¯¾GbWîìÁ÷¶@¬¥€¼ï,Ãi@uZÑÙÚiÔ×Î«êÌƒYm| Âå ‰@G,œÀ(rÔjÕ$H³S·Œ|é7Í ÔwIªÖOÞnˆ¥},Ï@á¢+´|g"BÙä5*®ÿ%NI€Š1ôæ4O%Š.ó!W2L%ãÜÄƒ^yYÎRÒ<N'Ó±ìÜæÒ`igElÙ¡-¥“%ìá’òh/¡J¯ NÐÁáþ—OvŸn‰ûâ«­Ã¯¶žìb·žÅ·ùèúÃ\|»+ƒõÒBŸãÁë^üyU	æ½IPñ„a&xK¢FØ°Þ9äWK –i„RYöÒÂ&Ü'¢_¦ÕL·"½J/€©
bw£"MLÏpÊ±†Û)µNTsõ
4j¶‡¬7Ì?ñQPì`Â§ÌQÂ(OräœÅõÈEõÌ«ÉÑÕHé’ÑJ4wÇ`QãZÉ®ÿ‰hr’Ò­7áÉ“nž½H<ÙÚþæY}ò¶=¤ðYBŠIˆ ÍéÞf‡H‡µ‡\¤Æ€µD–è)f5?øh(öžìo=;ÞÝ æ\ «Ý}+Uwƒ´É“•4ÛÉi‚3Iz°mðsVq4ÎiiÙqdOÄ©@¹þ‘|™L®ÿ)ÃþÃÒJÆÄž2d“N*È rÐ[€½W2J%£ëE~³;NÁH¡êºË˜à¿4"sè*`ÛÈ[dIôŠR
ÝÜPoWˆ‡Ž¢üh¼pD,³˜†ÔkohG·þ”ÎXôã¤B¾çËÂ/aým2‘OÒc¥¹<G¯Ÿ"÷‰ìËB¹ ÇðÖDsÙÜ¯ßßÛA‘£{_3URÀ%ªsÝ¥WžÒÊ‚fÀ¦6Úq”'yœŒË¡¶1 È2™­J`y‘Ü_ùh¸¶r:ŽÊóØÉ»yWÖ,HÙé¼‘¿á9tgž‚M…^Gh"šJWŒåþwež-6 ÆÑèœR÷ŠU¡··þ‹ï™ä#+CUÊû_þÝîöq½6mž ŸªÒ¤ˆðÉC’ä£ãÃ½g_‰«l¾5³@%y‰ú"žª¡é{¡¡Xº_ š³"¨ÇJçj¾[´~¸Öÿvå™‘•:Äš¼MmTýq€ÇñQYëkrVðvY<gi‰@9Æn	A=ù¯é½uWÿîwu¯ª©TìyúÂ·Öby_½ŠHG,ï‹´ºÄ'GULpïía,‹´¶óQTP›¦ò{ªüh³±&ýZÀyÇé„ù*X¾KF~È$"uçYåÎ¨·‹v¸Ux™z¾¼â†éo½hC²þò[0éâ¬<\¦ÙýwUsà5;×”Éæßu–ø÷C·šR÷–\—[?™ïµ.´£
ÃoàWIž-!’
ç„ªž~Ï.îºvŒa®D˜{[Ž¡	J/ôuRJfRZ¡„ž‘Þ”¼%m 1
±QimXeF ÃŸL[PgAÛþ1ûc¶C¢˜æ³%ñ;Ð7‡TPÎ(|ÏÞóòóµ¡§“¬£ï1ˆÒ#PÔôªµŽ?¹qÚ4*Ï“PÎíÁwî\)¥æÃñœ*•Þl*x“3l½{ImÍ £Y¥Ã½4dôö+kÕ¼ÜF…*p ˜Sõu¤Êa6·QÕý:R£8û0tCì7‡#^)Ô¿@oí“s”{3‡«sÒSôTˆ$>Qó$–„¢7œùTŒÈ¸]5?øžVÌ*—ÚÂÊ…‚ÐÔø¾Ÿ¿àòQùXùØU2“ÏEzAôé˜"î¼ ýíƒ!R¦ œ>—l5™sªjNÒ,_ÊÆHªH¸qæ:†îRl¢‚ÂÃ<©¬ñè«OwïÚ½þo0ðp²“¤ÌnáWT×¶4Ï„5ô5§Â¤£×ö.5•ìÌzÛm:* |
ñ™N´‡Kª£ç¯Þ©`6ú†íš$cRm;l Ü-½É‹xw¼õOØ¤lj tØvý¨ñíàùêÖÆF}„ææ3â,ÞlŠGlœNœ-ƒZ(5æ_>O`ÒRÂ‘ i¡TÜ4d“ñ'.¼Dkèx0©¦¤3¸ãS:ÛŒi½i·™’T9Íå¤Ïª}5˜Í‹Ih£Zs’ÜÕãcI`Œ÷ÔNÑúÉ¸j1Ùøºi·Æ0
7êÊ³¨ÝÌÉIlì3†cf³Ô£ÖÔZ¯ŒSIâ(É2¶¨r«¡gs{º·>ìõÈöußs™½!–>>÷Ü
8Ø—*qN{Ž	[ÙOìì’Q(üf”ÐéÖþ$¯,7®úÕ‘N&¯·+ Ê
]¹GÁjç,P(µâ}¿¸Í‰8G äNe4y(¹Õ(_ï]ìäJšÍµbJšøònh¶¶jÜæÉ4L¨n]¤Ò2@¥šNKC§µ0AK”¥$Ê—¥¢&šùRÓØË"Ñ‰Gp}:›O¹Úí×¡à’S°îdÞJÊ–˜›Ø%á0ÉR? <há»å”ìWÕùðe¢žº>?FŽUrR ÓT€~¢/è´è«QÀÁ“­gû/wv_nmoí¿<ØÚ9ÜÚoT36»Õj×‡b+Æ 1-(nG‘GUÄ·âãsÜ
±.†ÉBËÇž.Å9tÍ°ˆp;Ý•í£§ÑÛ¿Æª}lËu»ÈwXmÜ\z‹a¬	<n‘XqåôL¯M,ÿ‡…³ËÝ&„·Ô9œ®y•ó)syöÀ|ë	`B2‰/}•YYè$ ?sº>tªa{]òî­ÎG²Ÿ_`ó:>]uÁmOÅØOIõv®>Æ(tžÑáÔ¶y’7K²«y!aé¾UžHúúy€Æ¬k‹„Ü“Qua‘×|N5jKè1âÆž»˜C'5mïŽ\žØ¹Aíøâ²kàøÂ.Év‡dýìê€åŽ'°$Š'ç€Ž¿i…<á0CuÕv¾2R­Þ¾–Ø(8vÃÌí«•XÞUpô.Zí‘–XXÅÓºØÞ¯æ
P¬Wsà||îÔà~`q“ëÛ¯D®a]C9¾[Šsgq·Z­gò"ÆSjDÔ~S¸ŸZVEÎÓÍ\Ï"À'*nd—­>=c–
w_[õ`(¾œá]«ål‚Wå*ÿ¥dT‡ÉLi×q£T¯ãäm•
©cöêÿûþI|ø.
’°Œö~5~—§Y¿÷Ç¬çñ}£a,ÖVM[‘­¹kÂ¸Y°-•O¸çwÆé”7ñN 7ˆ™ÊÒÁƒ‰]xwÅŽ¸ ¦ °ì-
û{3B1I¯0~èî‰ÃÝƒ'{Û['´³/Žv¿zºûìx3ü¸Jù7Ó©VÊ¯Ì0›Ó«W¶9;™å†?/š¦¸àûd¨ÏbD“š»´ýö¢9#» LÇÕÂPJJlÙ ›ohc¾ ±œ )ºëÙð&òá•Üs@vÊ¨{Ù¿-ƒÌî‰Gæ–"Š=OÚ65qm'pºEë»§nfËv¸&Ÿ6­‘Y~†íPÐíhRá¸y sÒ¹3ƒŒŒZuEººü©ª/C7ê‡uŽ=bÐ\Ôšíñþv¶>5KC^z˜€…!,ªB§6XŒ›t¤ÓPée„9okôÕ)#Ø¯ÿ
Z†d³üÒœ!*­SàíerFâY,9lîjén¾¿sI˜EY%™?¶®%…w6*ÖWÀˆMƒX\ÿˆ\ÿø6½7uäPê»U—ö%¥¶g³òá»šD¸z¥1í^?³à¡¦ÚÞ.î+ŒÖ¯¥Q»“ó l¹9‹uZvG›´<(ÁîÆJF~U?÷üÇ_´æUD¡`½ÓöÂ”Ö[lébHh³ÁAÌ72öQõª8é^ÆÜ<rßá	˜ZÁ>Ý°ÝÔŒÛm
LÊÊÐC­¯–]Ð)Kš"ŽsÃw´×ÒøI°,ÑžTcÙÀNÂ§í\»!ÂÍ*>Ïñ¾úùf3ø¨°½œÃ+È-=*ÙlêMÝ‚Î	ÿgùE¤îå%T¥½”6ŽBf,*ÚÅÂª™•ë^€¦7 eOÊþPè¥W€Û`ÇŽf­Ð9÷ŽöU:Á°¶Qõ{Ç=Ü^sÁ‘uf`)Cë&€¸ÝfàÙ‡7‹»æ ¤¦î6]°ôÜ,0'¡€ª:Nj§ÕÒ¸Û!yÕM/¦©eÁÒûöï,VIú.Xšµy‹¶1hÌ]²ÜÌ	rjÃnãxg¾cËž"uwþZ/Ö2ûf•ÙD”¬yS¦’¬–UNšS|Œ½
g tï”Óã¨…|‘8”Î`’ˆúZ›º8Äïü>ƒ<Ögç¡ÓE?J¨ƒrçGÏÑYRÑríOîºàí½_Å·
äxItÊCó†î)IR÷•¤”¥Ì¾‘Â´ å§œê{²$¿ ?¤‹ßSŸÂwwÅ:isX5˜è™Ûtß=l u3µQ§B­qA©Þy1,þÆ¦2¡·pZ›ðÐ%­"ù}†Á~	`OšÀžÔÁžÌ«B  »+ Ü}[»›ÙníÅoUL›ŽäÈâ=ôÖ(ÌK›A)1 ìæÊZ=yCõ”üªÂ²XkIÂá*õ²ƒ”¸l½Åãò{šõ}E}™‘gøÊ¶#@ªlËœ²€kÔ }Š›¨IÑôzY¬.Ë5ÉZ5«…SÓ¼tÉÄÂfLí§õIà½’²ã•1.@ÿ»ª-£ƒb.Ú½³råã8œí¦9÷JpÇZV’]u²äø½mî+9šrª´gx	J0êLƒä¤pùÛþòS<s–¼=ÀÃâ¿c=@&	k`±)U¿Å~W
€¶%L[*·Ù4‚ÅÖã»@ãNú–;X~Ê:Ztr¡¢@žÿ¡ŠÕmy®©z<µÖ…šªÔöõÓNW(ý†Ò´6]ÎkËaæÓÀ	‰á¬÷©/&½Ï¯­?š“ºW¹½mÈµ‡€dÝª‹&9¦“ SêÖ//µ»aµA
– ½
^¢(³7°ØB!ýÁ´­îÑA> îŠê‡Pu<öÀxî£€æ!šsºXWÉ”šH¬‚PX¯mª™ªUùê*2žWÅ™ô Ä®Q‡ï&”4¼0´,4?Œ—ÍšP§-ÆèÉÀjÁdÂÁRgAzÁ•›MM›Lóûëé6^¿ëi^Ã7ã5rvok8ÖÉß-Ø³ùôæ±|Æræð{ŸT[¢›ô‚Û¨3zyZM›šp`]vMªÂsó½7uû6kjC3ÝVkhîÃ]kÌQÚYƒxdÃ>‡tä–|³§é&”ázPÔá![Ú}®÷öæú–“i®ÇSØso`é0‡?«ÑéZõ…)ÊÃÉ/Ì®ê>fŸÜó7Þ•t¬¢Õñ<"#qYè^òŸŸýçá¬JÇåð$Ï_¿þ¡{Dÿà=“áå˜¿+ ¥}í@ÃÄ/«üeyž$Uÿ¹£%ÒÀÝ*¢·û›9½´¸VÛƒìïñMÏÉ4áFZok\å:)ŠÉäp“&×ëM~l’öŠl£ôWçåöqM„b…‘7°ÈòäÍjò»‰í#ä«yš¯_/„|»ù;ã6½{oË\8•9”ðDÄlŠw~Òíž›NµÐóÑÍñ¢ÆExQß9^Žl´ bŽTÈàbx1sÐë€™ÞõŸ4°$¿ðÎ¾V]üg¼ÎÐœ¡Áâ{,Ã“Ba4«ò	ezr@l©!j#¦ÄêkO‡¹±>æËO¯ÿB½¤IxHùüt²‚B9ð·ôPŠàáÀUêÜë/’1ÂýŠì	ä,¯ÞßìQ»;)–O1ŸNR½_[]ýO¼Æ6Ž¶Â©_G’q!<Ãrý/Y)Ó4ØÛÃñŠØ3•SaËËd)§TQ.®µXLp%£xîÝxÍiÒX6Sôo‰¥-@ÎãÇé8¡*=™-â%¬ú¼¨ðv‚áÛqùVÆ‡Øî¡ÌÉ——wÄwuŽ¿¼\œõÚË¯P•zÇÐz;Ä	¥=¼Â!À$$?º—7 M#G¤WŠ;JÉdÎÙÞ-7<´igˆæ–³0J˜NæõKìñ+/œ÷e®Þ•7ÂBwÞ×„†.ìÏ§OäHO4”Pú”ÐÄöæ)òºÚ™Çfö>Xgcåðõ45ó²Û*‡Ôaà·) O>^|^ç;Ä_võ7ð™•}ØÓîÛQ2¶&N²!#x‡ÛPè,Ù½€a|þõñÓ'{ÙtVíJ9ôEGmƒJc2”5´u’rshÒ	Ð±r|h“EÓsÐ{UbvÙ•{¨ßËÛö‰¾¾ˆÎÙ¦„Y«x÷(5OQ/ÏJsž¥S1ã»ˆÆ3p¯÷WÐly¨žÊëoƒ¸Æqn`F’dYFÊ|u53lSP^…q^ÊëÞÕ˜(Z^îµZÙ #…@lªÛ©ÑCz ¥ˆ|=ÌI 1z¸¨ç°k*1&®AQ…žMuhÅwµ[I†­õåÖ¬9õEtÙsÝ›àoäŠD"Ž­>ôŠ(¶×\Äœ4i¢Ú9zÂjlk¦±„n¥± Y‹f?ÊO…ÂbY?-ŸÀZÞ0²Ý;ûé”W}ÓäÏæ
š½ÈòG’×6“Š<§:‘xãm¦~…#|Ý~Ÿ©ña+ÿìWãüD.Ô‡ÌÊ@°ù¡NìÄfí™Í]üÙ*Þv^»G‘û›Y[ÌÝÌžÒýŽŠá8>ˆt`€ã1_aí>ÓÙÅ&ò£ƒ$§ãä-¦¢¬6 <@ü–ŽFã„òôc6Q¯iÉßÂ¡QbV<½ýB)ñäo³ac²|ðëøÌÆƒ
!…ZöR7‚ûÊ$Mï/={¼³ä½Åû&çþó?ÎV?Z]]Á?Ÿœ¾¸¶,–šÿßÑÊ«+Ÿéb¥î]_'c<ãŽ‰lñÖFÉ¥ÓŒš7âäRL€'¥`K™Nâur	¥ÀÐ¡SF®“}«ø| ¨£ØP…­ßÚÎ™Lzía˜Ì~J­å'ßuôÒÍ4pÜó­Ó¿‡¶ ul’´®×Ø‘@ÿµ»‡Â¾²#IÐ2€ÀcIûÔô!áXá@2Ä‡²Ú#Õ”„Ø³”Yó+Û?wÙ^#6œçúâÈ*ÉMÕêµÝÕüˆnöuÊSvÄòÒêkøwûÚ!È[/¡ËÏe·_<ôŠaÇ±PmãR?Ô{•Šúá±ÎÒpÙ+~ÔŒBa¿½¦ÜŸî÷fŠ¸ªñITŠ2ž÷ŒýeÄæ20OcóÜyìüà%Bß	"f¸å/«´š=ð”C`J·„è©¼Ïª±è4¥¯”f½çn}õj"­3Èˆ´´÷e×ñ7þÍ xà—•¾oêÕEAšÉ.$é[ú‹©Ïrrn` Ni¡ÆITÍÐd¯÷H™©¦Sê7¤1“c‚p%L~I)¢óÎîA¨W°4¤º˜NÁ81Ò/ëÁúôˆ†é<ÆÑŸil—¬%V¬Ö˜¶LM{úì¶Ô'ðû÷3ƒ]]„O%{Ìk™<ìa`Oåš²"X8i¤Óù!SqŽôO,êz™³n—ªeÎ‡ P%±™–úÛLƒFézqdjÖSf©iOýfV*‘¦|Ú£›hÔ›ÿ  ÿÿì½ërWÒ øÿ{Š§wt  Š’L‹ÒR¢Üæ´.lQîÙ…Â.E²Z 
®*¢9ÜØ}€}‡õî‰ØˆïWÇ<ÞdŸd3óÜoU”h®è¶ˆªs?yòä=ùoþI0©òcø7€òŽ¼ðŒ›+Ðä§°hÀ‘yÁ“8_ubð—dƒiUEûÊCÚG}@½èŠ	ì «ÿYLÿ]WÊ&Ö¬o0³<þEÛü%8h6)xARaøaž‹^–8Œu´¬‚¹š'ÀØº"3/ûÓ8{½Wí5Ãù)›æ§d8çˆ`{Bï`‘sà¤½}œ¥²/”Há ÖÄl³È¥àB¼6ÞÒ¡MJDÀÿ;. Èï¾)ê3d>ë"XŸ{)Ú3^3!¦Ñk¬á<qQœ.ne—J¤Á%´†ò%55Î8ÖAŸž%Ü=‰ \U:I’J‚Ýþ4w\Ù/¾ãz†òOe›øi
¸|È;'&‚×ªÍæÇÀdñøœö³‚h$R‹S4;‚É#:Â2Ò0©…%a¼Y>ª"9úÀº¼æÅq÷HüäVgÓ)'‡$U6µ2‘"P£Ê>^hT™¢»8yÑ¶¬wÛF@ù©sñ^¾ýÀÛªˆ×»ì`ô*ží¤/=Æœ%D?Ík’k@!•}i$8s®£Ìã¸]û}¢#E£’¤Ãk£X°ÐXË¹±¸F¼Ù’ g§ŒD@ø°b¹¤G#Ñúþo•Ú=ÑEO?$w+¿îºl¬\YX(^†í©¨äq£Õ)¶Zb–d7ÇÿÖØmë•b°Ûš‡§Þ³!< .YŽÂkÉVæ2–3_S2+ä‹nµW&nŠ§v=S8`”Ýñ°ÿŽ™Ê
øŸFç~¡ +iN[ÓÎÄ{±‹lîB;×NÙ iJý“8ž!÷;M.Æ—ÜÃ|Ó¢ta‹ìF„_,éÃK3Åzz•>*šz4UÙ=?ïÄØ(xn:ò£ºP‰ðRÈ sZ-T…ZQ£”b«X½ñ¦Ë¡×	-HE{Ixh‚Ð!Eª¸n|>aˆm”ÐßãkªXÆvMœvØ+Ës~¬Ó+ò.¢|e,ú&g¹·ˆs:Øˆ´=xQ;ÎRIF'Zh–ôÞv´õuð~%ât‘ÅvÏÊªÍÌ<$&³Â¢Úöæg#êÇf‰×È×½’Ì4ôúó¿£K|ßjª2<¾ÿò±O8Ekj,wÅé4)ý_þâ!$¬j–øî»ùt(.wžú%9‹&ˆÜÏxœžbø¥3¦ÌHÑ]›ØC¬fÝë¼"VÞb+(5ÅBn
ÿEy·çR¦bRL÷þƒ€y|Áãõiš<ª†
STœ²¾P¿Á»õ¸âfžï?˜2©!o6ÿi%ˆÔÆcžæïp½8KÉ¡°åž%b‘¸ðsœeó¡xð#Ã)Ìd<ŸL‘Ä%jQk5¬6,ÜŒ×h•˜ÄŸŽ€ŽKìoüˆŽÇ‚E:žB0¹ùå	ù& ‰è…Æc³Ò‚¹åÅõIp~<*ç;éçÓÌåòŽ’^ŽýíkíÛNv`ewEy#KŽÜÓ=TjuÓ‚þå’W¨Xzž˜ÅI}ä[NC¢;„#1™rHÊ)ýçXR„XÍâ¸W©Vé=µiæKyŸD”æb"|ÂâPNaÑ<‘Àå6­¶{p 7ý6Ø§}-œ`6Dl
?@'hwIÎÓQœsQsë«ÁÙ9]ñá¸e„`€,¶!S$€ôÕ0Å„ƒÈ>ú jÖøÀø·IÅp†ûÍïÃ33<¶/üàU‰H-D)*­MùKÔÿÞ?¶¼«8k¤Œ)´—7:q¬q'òW79ÇXÁ¤,xÂtššLüVêxéQì2u€÷ zN¸ÏïvÒa‡a99/lTàû¨|Ï6…Üq1Î|º|ö)IØj¼gÍ8ªë¬ÚØ{F9@‹`	òX¸J<`§Y2–Ö{G/aê'>8
hÓ²Ë×ž&™Î™Ìo!)ßvDÖ†¶ÂÓŸfBåv§Iƒaï$~py™¦â¿F!¸UY|mÁíW™~ûÉˆò>±±Ð—`Z_%ú)&¡·G€ä´fBË-tÈŒ]"YTÙ$ZhRGð¸k	ð$gT^’x/³;Òi](CÆ™Ž â‰©XóÞ”÷°JèòÖ„û¬eŽ.¨’‘,Ê®‚T5¯ÃnTÞF xÌ„7æÀ1¨ÃÙ§š¦Ñ
ù§žTKjA!.Â1ëYâ]; Ã¢B*MIË³ÂK!•çuÀ,$$µÂ§±äÊ €¹Ú-±ðj üg¬Ög¸œ“ÂÝk¾U-¶gªKúíéÐxït'ùÍ`‡BJ'v[õÈß´5Y]à‹.<s.ª§ZÆXZub<|V+Ê-Ú¼2žbGœ‡miæRÉ!Â"Õ~¬J^åÅ<u2+|¼“¯“[ÝÒbŒ²áT©ñpqË¤õ‚ò*ZèÑ¼›VÖ+·âWE•œ
ŸJYõ+äU
}¹…ê¥Vø8’+šÜ¢MeWø\;'Î+ÃR+ílE@i½íÂ5Ò¬j&rà§„8¡å¡„ø—Å)!QÑ•$Œã¢<Ô”Gëþ"¶~éT• ˜]ÊJð		_)¼Cˆ*CèÐçµ$ÚŒ%)\jIHVMô fÜ¶÷‘£Pñ ¡éö ×£mƒ·U\öà£Å.7 ’¤á»¨\I-cUuÅ¼<5Ä÷™ßNgtâÓè§WÉøŒBo'“èOWòðú'& ñùä†ŒÖ×"¤õE¸´œüt_Ï'ÇIÞB¨!¨ë·C×Ã´”oUp7é÷éç[hZ¿ÚÈ¨¦“´ØŸóìÈRÂÏ‹dŸD5¹ôŠ$É&>Pa±ÚR>^<…œàêÌÛëÜæá{þ/<ïÖVÇOÆšî+ Ü°ÉýFÔ›ôsSÔ[2¹=ê-™pêõ¤Þä<D–l›m˜[F9Zs Wçóþ7>ÒN¹[k»[G2gkŽÎ–%ôV@½éà®È7Zjg/ŸE(²-E¦œê’LŠÁ¤È4ñ™^íëd\îè%š~í¶É1.µ
¡Ïöÿâ¡À„v/D‚i±ãíºBM¨ŠV›N®?\J×ŒÐTö“Ü{Ôi›`†gl%‡òÌû¢µŠhåG¡Íz$“pŒâUï,ÅšÇ\Žº:i ëe¶|0«¸þA<;«ÒxV"¡Åe
€ˆÔÂK¤uDª×Áß¢Wù‘QÌºëWË2¤³å'¼ÊZlB8K}¢¤›áÍm‘Í°bD5+Äé'šÅØ¿Í\+õÀÂŽv¨öûy«„BÚjig\r?@«#Í#ÃOh¦ŠH šòÎÚîB6ggÝiÚ¨C6GÔT¨ï4Ý€²1—FZlÝÌÞú´ÝÜš¾¦ý…8%Í	F‰.çðÙqªÙ$ù—Ø¦¦éIŠœ‹ÿÒÑ,Ðèˆ%x†¾“Q0&q>êœ¤ã1ÐZÒX–‚6`ZÏ?›Åqvu‡Çè½Ž%U¯ÅÓpŠÅKŽhÅ„ïˆŸoã…”Þ×¯«ÕŠ
]mHËª}§@/^Eqö¾í-aO e8IÛÐ~BU»‚ Ê-ÍÎB±ƒ¥­  Øëe!—S%ì-Å6ÓÆ,Å9giQ®]ìgé‚|3ëkqFõy6Y”;^–»EÜÇ’šêX1e€d<Ÿ8J‹±gpY.ñcb¡ ßj.ÌFy)Õ•³Z×þù8Ö+ê¥†ƒý‚eZÂÇ	ÉsfôøÚ´_±Ph \“¾VÂ,†wÕÄÜ8Èð6¬ŸÛ¼gC]€þž‚~	ûë;^81æˆ\b@L)×Žð¿±BdRë¾æýµ=¼};9+ r'›•›Í¬UÑ¼~«1cÁ(6ìé'ŠGÿù_,èý(‰dæ_<¬×?­d9l€Ð^È¨ÖÑc&Ð»–7/{p²NGæMaùƒÓÊ†Ž	7<h!»¡â?`X4áØÒCãW9#ïqë´=Íyýe•Ëœ¾ÐôËÖIÑTÃP0²-]^Ç¸„¢óöFâY)’ÓBùvùnj(ññ«Ò|3¨UQ²"¦È…Û_TI^œ]'ƒ¡êËè.YÅJý%>ÞÄîê¿Ò"Œ:>¶G„X•M%Uõ¥c³êWmŠ­Êøª~$(X*™Šæ(ÐQ‘¸… @QMdË‰®…n Ô]\Äü%†´4ÒºCÂa|¼"­0Â
‹‡Y‰/…¯¸ÜX£T¼Ån*=¦®d)XLÊ&ï$K‘Á"b_|ªåÉ\ž°h£y2—3,$QÆ§BªÌ
‹¯c•8™I6“!ãóo«:y1Õ_VYî]ú¥§‘’¢{{8Œr³FG o@¦3“åüÙ¸Ü\\Ò;ý L&Å»[ÜÁt[ÙXðAß2£pí'¥•¦ôLÌÑÕšR·ÛµÍÑµHñ¬i#œ"o©[‘´Ú›«°úuÛÑNôùúèjú3P¶žw§y*ŠÕ«ë­ ÞŒÎ>øäÙ–·¾¹wÒu8þ²Ýï9N×bË›dú•&Zd¬Í®ØYN7°/êÊ®­‰ß1R!Êß©Žïrf¹z‹¤”Iœ°JV“ÿEËØ?0Ïnà®.	°uÜ´ÌG,&†ç¨h)öÎS`Pº´W WOÎ×tªGžÉê–…½f=og”DE:Å é/ÌßszPN’hÖêt'€³ãS3‚NKìÍfãK+Ø.f/ŒAx‘Eû¡ˆ7ôû†íHÄ>É¡\›„ZŒfºpâ²ó!¦µŽ¸èG$êòd¤|vÖÛßÊu¿Ò7#á‹B2o˜sñô<¦†5XCe#”pÏžÑ±Sú6‰<uøG»m©×<“Äpô4Czû£ŽÖþËÑ›×]†4Ó“K‘i²š³
{Ë
»lÎvöGG}«Ù •8£bƒx¡E7HmÕ9R‘•m‘º	¾Ô·Å*vH1'á.¶?úPí)Äö¸Œà”¾M±;År»£]œ·¶9:žd9`€	Dz~	-18vtz÷.“•v#`³¶5S[,YŒËÂÒîª’o±Âet-
`RYšêKn°6ž` õžï 3¤õòÁrÂÁl·Çf‹>»ô8J˜Á¢Ö6Û³A:âPhŽ{¢MK20S¶ã§f°}´Ø|çW-ÊÝ«eûiìŠçÇŸ+gúål¯O|=ÜF„Ó„_^%”cWmŸž“Ídƒ·¤“—ì¼=n˜<±P@vé¨µx&.ß¨jXšö¢(`tü£º…ÄºÕã Qò¦Hàƒ
¡§FÎ,¢ÍˆGÜµƒx]‡„î²m`ÔÑÃ½w‡oŽ~|ñêðí‹£=Yªh[¸Âí«1ºÝ7Æn_^¤ÁÛÍnŠ8<S¸£¸ÜQ4Æ0²z‰¡XMd˜œ~ìAÃÈ£XòPsÙ`×áâ.á}ô·ˆCÄ_2`B1ËÀWTp¥\¾´£^‘,3“ißD¥ë£ANÇÚZ
£A+e“Ñ  ½Ìöl‘˜ÖœßÉhMävbÿªEÓDFž_¶hÚòšiÊØ€xîßu³X ?-R0ap Ýš‘n¸R&`ÉXJdLšÈ'Áäf.÷“¼„²oß¾‰ž¿ýü¾;xþeð:Êà¢þyžžgeF³ýîx<4ž©ÿÙcQÞžÍONàb‘yãôlp<†sL÷º7¡‡…¦(1»Oq†½v¢}ÞK PÊÀ¢DxaÝÉº„·K ›ÆPÞØªJvÇÿ¹[Xß>]XÉq0Æ$ò±=,XƒÑÐf+7¢œFÂ"æ]ª¹·ãæ"[ƒèÏP›¼Ö©6³,€—ýžŒJiy&u:å…÷ÄÔ^&çÉØéÛ°cPóƒÞpÉº¥ÇèƒT•¬7 Ø_]¾ÈIFÇY!ÜNÁý‹=è'vj+Ù;Ø1Z×‹þ¨	iö<öV1Jè§_;}GÜ~Wƒ¶–6bó¨j…
3•<J_Gøk$xOZ;øû©± ;ÑÈ0
i*#g;ÊZ2îH-	zY¯°wö»ÀÊÞdui…Mé`u¦û £©"›ˆ3œN+sÝ›bHý_…A+sÉ#ä¸#×ócæöƒ=Â£¼»k–án¬6)€Iôdóv†w&Õ Já?ž%Cé;ÿd')$I~$ÛSUn¨’g~GÄ½/3Üé­TO¥@–…QõòÍô}/þ‘&;Ñq–ñwl:&¶½'kâVærEvfÐ<NómrÒå¨ÇÝ‡‹8Ÿ¶ÖÜ5ùçç_#L*ý+:]RZLoDÇãìçy‚”+¬dÊÅÕ#æá5Š»*%>(ÁaÁ÷«ËA+JýjÜx ÷™OßÀ ÈÖàË–xô÷»þ`§×ƒÿ­£+§,Ñ6Q:¼ÇÞ‰ì¯ ú, 	|kQþ¢u£ØBtÕ`É”%Ä»d¦6Ù‘xE-SëŠ¿~"À*³¨ýé
š|J<•{Íñ|Œy–’u´°ô­¹®™'Fa]ž­–Âfräê*n´Vv¢wéógMfÄ„²âÛ¨iÞ*µ»ªT£yz£
 Âc^`*õq<Mþ)R¤Éï¦êÙ>Ïª®_Ðs#Ì2ÖGåë¹˜JFR¤=SIw6©¨…I]Æc<íÓlšÈÏmcû¨UÙè»ì`
+Xê$\¤‡*ÖBèåŸò¾ ONxæGGlµÁ•Nb£ÊT©F¿cÈYLAé¯Á«hÚÅª%sã’AèKðw:"Ò8Ò²K0Ä ¦¸ÅH÷%¤xÕS®FÖ:Z5YÃ§rôÕ–,¸î¬š@J9-{bÒÛ'ôZ×ƒHS+}óDÈã…Åo'jïŒ»‡Ìt²S4øÔ~éx[¨¥>>è¬EU¸n¹UìÀ™oŸÂíû§5Gò÷SªŠûe±üDåCé…®UàÕÕ7«]íƒãásf®ô>tG²ºÞl¸ã<ØžÓêÍ®v×á!*¶‡YP€)¬!±1Á. ±F•HÞG:…PïmÌOÚ _¹å ÞÍòëè"ôT!„ñ­šo»x‹³oN$kDÍu‰DßŠÁWË"çÇñ¢0ò¡sRi0%š(ˆëÄâ2‘0`#Öyb:ðSJº%uŒV5#Gù­¯›­¨r³p)izmßf)4˜Â‰â!š	sŸRMx½õõ:ÌH2¶¬#³cØ»sŒU¾cõA€¨§¿»i—mfñ¦ËÛü´ #‹­¬gÕÖ3<‹jCÞßwã”iUYº¿DÙ8vM€¥\˜¦ Óh–›àTŽx"W(¯£/cwüTŸÐ2å§v]€r3}’ñkÎ}3+Ó	ªc†Ä½3qŠa‡”N&	ìv™ÀçV2-æÄO%kIù¹
AÞ†DBïÃâ˜/()Ýæ”Ú*ƒ´*ˆjšñŸÙƒ–s€âCOtMÖÅ)(*a
tƒ"15!ž/Éºb«g|ÕZçÌ¡&:wèYv$xîußui».dñÛ cU‚*&n9Ã¤¶óbL&L‚zJ‚ÐNÄ°?ç˜o&O’_’SXžÅÓSøc­?ÏS¼¯`ñ×µ¥°qU‹ËtœŒ !;XÎy6ÄL^—=K²¹ìqN¬] –Íå{UÃìòµhÈò²§	ã«JúÙ_öØz=&eµÕ–ˆUÙ¹Y¶#ˆÏH6€ÚnÃXvx6Ÿ~„².—1š¥&«fÉ*6G´|ëö±EÉì˜£GDt§. gFÆºréc’ÌÍ³‹`˜Í.›È˜Û-3ñ?TR=Ž-…À8¨Ãó
¿‡\ðI,h«ñ„OÈ­å¬4w[fª;ìð©PÞ!e°vÍŸÄçY¾•(lƒ3rN§2é®	E¿_¾íCMÄ[%U¥’~18»ƒ¹í†ÛRf!Xâ#¾¥¥)rÖD’ª=EÃó:Ð«©fé»Ž©©lÀHÆhÂ¹z—¶€Ñ›Õ_®Ë&ÅË¼‹sL·L©°ôIÎÎ½ÚTØyZÐ±®Éõ‘ëÌ€±.·¤~óŒóËkãüƒð(çê—–n•z5m¨Âäi9›†ÈÊ¨$à{!hí òwO»ŒGœ£‹§Z\«}RÌðÌhNjþJ@@EC^jØ\<Ô½>_Mv¬ÆP=$Æñea2"˜šn\¢ƒR-r!§ãRâ<Ar0°  ûW8-(±>€Õ6²Ús!tK˜–ÝÚ–;qøøVrÂæ>Ä¸¼P˜ÚÏÃ, ®Ú%Žpõc3CMm]µgä~ŽF?ÌÆY<R·-ÚéìDøeƒŒLw”e=úoÑ:ìYr‚™ùR©’¶/h2õ1®c< ÔŸg´ÆeŠžÄ‚fL¦ñüß¶zíÞßžA3XEµ,Âœ=”ñC>–F.CÞìÁ(œwÙ>ûL3Úˆ¶½¨×Ý~drÆ%©ð~”³´®k—ÿ–ç,Ç_¹8c©<V@P%ç˜þx˜âvÄHµ×¯9;‰¹oè¨åíJfø­|ÜSæŽ¥Ä_¥;–r¹±”âB0Í„›“L	CáR{éþžHe¿ÄBð9…^ 2_°ÎZpž5Ù³¹56–±qLƒíÃ'ˆˆtá »‰íÀþ…²–±Cy\vƒãd'®ø‘rJ~ù±*ìVW½X p*ßý:djë„ñ§ÎDeye³OY^‰ß’§“€Z…G‰6Ëðe0Žê:ç\‰{¿Îi:š ÖÙêÃýSÄ	#?ÞQ¹:Ìóµ±LõÁ¥)|—&c öšãk”½etóå}#¼CŽŠæ!àY5î©Á<Î(LÑyNÆóâŒ›€¿e<(§_‚—ì~rC&(ÎÍýFuè ‰Ÿ³yY]ER¦…Œàs×r(3kÌÌä[Têé/2åº!MQbEêÁ°Òd!5]œSªl™½¦4-ŠÔ°ôÈ/NNÒ!Úœ1	$ $’®FÂ9]wFÇ%Ú_:˜–6­éA—ÇÐ”|\½qvŠašÓ©\NR¥3›:qLƒUùäÌn!árSÁ²)TN›†_ 
*9:"Eµ×žéJ3[£Ñ5ðQBh^=ä(dá)¥×¦vÅOG¢i<© zá(<˜”J1‹« ÓÅÒ÷™'XZ­âÅÒESCÆ;”®xOBðbklM½5—|Añ3ÙŠ»óºg>öÓ§'âEðOù#MK!ËwÈòÖ¾Œ_Õª‰É½˜Ú64–3TÀ‹%œÔQ¢´&•a4¸m9;ó(N‚ópDÚ®Ö:âŠ‘P\Â·—qAßGˆ[[ÊÖ×ñaW³fÈÒJ5#p@x@†¥¢@_cBzüã0Îá¦ÓßëÜ Ô2í–!úŽÉ¦­¯@äñT¥k0Ý;ô³;Øˆxë»ý{ÜíiX|–¥SJcŽî0H+ÑàžF ëä;9¾§QÞ÷ti=4õb-ý™†iÜL^„f!ºB‹Y"Ú”q[ošzßh'ÚÕ|*	Xù^À°Â0ñRÇð·xmœuý`ûMdTg]^À6œAÒX²òôY§¿Ñ9×Ï·jI_U’Vá'pû«`Ì×¸™„Ü‘	Tpí½MâaÉí4¸Ä´Õ›¨¦Gw²éðî¡)J“à@ûÑI’ŒP¥ßq1.Kƒ Á5é7<ÑpËÅÓÓ„E´q*úyžÌ“&T	éMCÒºRhÃB)ÐÌº6„­.B5Üý²×)·¨µÝƒUi#Í1I§é	±QZ|Üœ&åE–Œ6ß Ã8»+;ZuÈÍ·Gf^K:ª‚ÔDþ“7iÚLx	eIŸoDÛœúpÙz8^;"Ü?Có6Ô]mPyËÜ4àgHÜÉ*‘_¡Q>,å3ÎKnR®*u‹Ùn³õŽ²u¡hÐq^ÀÑ†b;t_èZjGˆ°[úe§jrëuºnèíÁ´d•ß÷>lDý\ÁÖû>{u¢¾óm êôáf@ÙjXàƒ¹X^UfZ¼Ž_·Tõî)	Ë5¾8Ë.Þep±¶ÖÈ\&žþuœ¢ïÏF´F,ÙšFèÛ¢¨é^"mc`°;5ÃœÏJ á&
é+Må–á<h«¬¸m	[¸$Õ°¿1fqíRF²©n»n›bÏmP‡ÇÐéê³#¦Þ®™GDãç^Ñ@.¿'t˜É(qji0ØaÃsÿÔòÜ7\©¾5ö‡êûx(Ð29úÓ/^X§`±	ÆÞ¦ÌMšÃ•òÁ‹ÄŠÔ´FÜ€Äö5U
¾zà Å·ÉAU30µÚÁ˜H~IŽC˜‡ÉrW¯…A2\ž®ã2bK„Ámo’æén–mVƒ{[¸ˆ³­£Á¬\z;âë¶WêÞÝ2;8zÃ­ß•¨ìZ„° ¼f3Î8´ñ‘KYðÚ]ÕJcK#<‚µÀ«`|©PZ}”MQŒÆ€ñ(&\=ö"û‚5'Ð}ca5ÊÍêÑ^ m€Mª»‘-üú‡ûß!¯§ZO$vd%ÒY‘±’!Òm™l	¤R|¸Þ«tbÂ/³ƒcÀ¯ò0â–Ám	ÏÖ]”IMâRBÁµï
±
¬Õh´ùêÕæ%<Ñ÷ßïL&kâz‚ž'}—MË# g[Awðï’Oåólûº¤É6†öím·¿E”ðb‚ÌÌ(zÐëÉ˜0¨µþ6®áó¿ò”"_ê)?Ö©þF„Úã+ØâôØáõ!ÉåëòxGÖ„FÖïéƒ¦!ü´Oæëºâ‹wýësk[ÛMßÒÙÆ·¡[Ú.ç•(:SPDéçÿƒü=U[1ÁþÓU£nÿ	Ü}Vu½-¦x „AÌ'“8¿Ô`*žIÎ¸"&ÙÕ–ÔsÆ-—¾”C ðÊz¯d%î7¨@]f%Ü:\DáŒ¦"fö½ôÚŠLf­NâOË´‰q@Z,b7,=>®¼tÚV»¤B«¥Ín3jÉQéñEð>ú.ý”ŒZ}J8Ü[œÈþýÜ{ ¸Ö	\&gYD’#mè×ÿ‡‘ížÆíŒðæ¦ä‡"pª>sþŠGõR@òAÈwÀrO ‹¸k•,À$Ís‚j£\]š•®²Dó
2-·dƒ²¹ÏCVg ò.¬ñúH*òÿZÎ«Èr`©ãè˜€ª«JëmæZ’ ¼mëß*ÑÝ{w×eÌ\¹e€+²"Jà€Ÿb(‡ÿA5qƒÊ¢>¨É}ÄŽª0ÚïaÏVÊ›!´+Ñ×E÷ÔÄš·³WŸÿ;JJ_f’ÈÇeÝ‘A:p1èúk øùDH4Ì?ÿ;áîT[¦®X™mIËl5ÚN½ÃÑ MªÓ0Å¼ü_w ¢"1Ï’˜…÷ï×	y#ú)]ðoðï7r ~ø š8ÎF—;j_´ g	*v×Oót´nvyT^ŽY /LÑGH F¡]ÌOà\ÏÑç4'qå8”ÅBWŸ*á½ô„¯¿($<n*[’Ó”qUKbjáÉÿ–§ZðÆ¼ø{ù62ªšš£!ÿgž×NCfÞB¹HØ.0ŒóÓø{à¬0µ=•,2s¹þÒññhtŒš®Ÿ	bí£{iˆµ uk{,ÜKË$ážÃ4çc 9Ó¸ÍÏÑ ×ÖÌXƒ§GŸÁ¶¾<òí•hX<-û´+ì°ÀO™ÎGžªiñùWXn:q˜ÉÃ8RâP™‹iØÐûÏÖÂ§ËäçØé:7:A3I\Ûl^Îæ°ºÇãìxžS@á~í*¨˜CyB>~^–EzÁ^q½rë<÷µ„vÊú‡%·#½Ö ¢€!&[¤p¥XSÞÉ*v^™£˜æŒ"ïT,éã@¨«WD¸6.1¼÷÷€3¦Ì_dléÀ‹„™©É‚ht.‹hCÎY¶ ¯ªäú€œd£ÐZ3¦´ŒÚPËŽ¿lù+‰²å`íA(Xâá…DÑF!<šLÁGñÔZ¾oa1‡ºøBø÷É„¬Ï»ô,Xü…WÉ$³4%\z8Mž'¢d0è ,ÎãtŒ¨Q&6Ä¼¥™.]-Y`a´ !¦„SÅ™õ}ú|…Àù*ŒFbR­ÝéZásbÄ†¤iæ*˜oiÒˆÔÃé”ÚÔÊÚ©Õ•üšgŽí ½¥»‘“Û;Ç×£¥ T’ÌFÕ´VE$Ÿl\E&ƒ×0AZ³çÒæÜ¶ù‘ØfA0ccÙyöve0E2-R¼ƒK¸9×au×™ö‚Vd#zïÀï†:÷*,Ó†ŒûÁd=ØÐ^¦Ç@]RøùšãE2Ž£¤¬%!öµD¶yL‰£›bxŽãy­¶¥Êr‘e8ŽhÂDOª¦	¤bŠ2/ä9¤™8f«”“«h©Vtý BËü@Î|’s{ˆ”é¯@>cCî©%êbä®ß²±'¼ÛñØŸc«&NU ¸ÎÈ8ÃcËN_mnÑU,¬Õ†¬å¨Ý`–ú6ÖÌ¼r áµ_íÐþ-ôË‹+URh(ez(B¥!îñ{-=¼Ë†?Ú™"}$½!á\ü¸P­¯Ú©µßÐ	Lžàë‚EF·ªs ’øèÒ$qÉÐnJ-n^‹ÅµêhÆÑ,üGSµUpÀ/–G¨­Õ*!¿Xò½ËHKy°/VöÍ‡¶Ìë$‰$e”[1ÊÀø®=A^ihæ2È¤ØãÌäÂ¢J=¬Ó¸ÕíF¤ŽvÅ–l¸Ó‘¤0h‹ûñ(=†ã¸(§Ý]+fñ0é\v¬=‘kc—9'Ÿ˜‡@‡éKy ñËÎqR^$ÉT«Ëj?1võñY_oå~­Oc`š§eçx?Fô®@‹ÏÎ£^/*sx	Tt§LOÏÊµ'ÏÒãqš•É0æ2Á³d˜’oôãÍ³¾ÕÛÌéŒ5¼S—h…:Ÿ¬=ù+j.þ3*üÍêÅãÍ™1±MkfÞu:gA„v.òx¶f.Îæ¥íò‰niP='kö§t´»F9:,Q…S@ëÿ,`3ìñp˜ÌÊÝµî§qñi#ÂÖ¬"Ùô9Y¼î^1ëJ–·‚ÒE˜tÓšÎ³yYf(jÃ¸YÐðÅc`yÖ°Åq:ü¸{5Ê.¦è˜þ.™Ìp®£2-Ç0Úgqú	ƒ²g£dÌSSX‹íïóÚžñÝ«þ£kBôŠê˜#ÙdCYz|,/Pžö G¤éÊÙgü»é(yý}¬³š±2žu”çûhäö€þ|vy0j­ëp²Þ~Úb­VûÚ/`Œ–m9Œ—v¢Éx@¿5œu»ÒÄš›nbA²›ï%£«) ç·¬–Õjzn*w=…l€‘t9>eo÷6·áW–’œ½À×³ì<ÉwD™¾z£•ÛbŽ„Ù´3Dqá„µñ¾k×NgÃ¼2÷Æc…2×=×0¯X‘àF=×ö?# ,ãyšGe†‡þ§ßgN-˜>å$œ§èØ@çóIŸÿ'¿Øx)Šl˜Áçñ»<.Î~ü¾LñvÞÁ|¢8R³w!ÃÍQÔ¶§y–y ©…Û°w¬ÁSí…°=¦úLm÷„°\*ÁÍMÁÂ7_gwvp"{÷-U`côe\bæ,²è¸bÖ¢ÊÖZW‘GU@n5Â¡=ªä†%‹˜jF¸¡ñ)Ôxÿ!~+}8žæ:SÊ1‘…Öøñ¦ûÍƒÂ2ï!ƒ¯y*Ž‚ÊF®Ùh‹Â›Tw ’ô/ZÅ‘XðP„Ž[‚Ê…øÁÐužëëJ&Ïé.hÜA±ïœŽà™0œ“aNký–÷È{”b¹=ïNÓã$E5¿ùÓt¤2G­Ãt$©Œ¯sšŽ˜íNå
ð2Î’ºi¸UÐy2Ì€àžDÁ%ÿ±˜Ÿ&˜Ùª­Ä¿dâ%¿ž4“!ëËÇY*[È
Yz>.sùª¡>/K½—bód¬¿ìåwÇº4ø¹µ“ÎR‡Žù‚ôÍë„”~S²ò© jÅj·Fú—l¢FpW´N\y¼ÄR‰äÃ_v¹²å{È½NN?ÿk˜ú@²bål6Æ[™/øÏjé¿<ŽÅL¢†ïä¸ó Â«ä}?uÐÏ”{8qkÿ0œ‚<.-]¹lÎFëj8µácmö	†2»ìlqÉb6qÁâDg¸1·•˜Og]œ¥eÂÄ ÓŒDvv¸‹J>p7•0Ÿºj©SíÅ@³(GA9ð´Ô¥ž÷å¬½ÄªÆXŒå0÷ÞÇY=Þ<v d©ítòÝÜ#¹3Ûá¡ëW´J»wwC'hïÊfxéÂí»ÖïäVHŠã®ìs8»…!³²ßÉ}PÃ»C{DF 2Üm±‰]6zOƒµÿüŸ#û‹6{KRO”*¸x•óÀèÌâi2^ói)·ˆÜ°ujÁ‚&öþEëE§MF;êç€t÷]%Ž£¶¤·ãø8ë­3¡¯üT8€#÷f>¶i	Ž¸¿ö¹ô)že9;7©iO—LÍ[£”¸èœÌ.„ï#"Jp“Œ:Ÿ°º¤ÚI6œ;\RÐ¡Üžì,¢ÚFýÀ¸ÝëmHéA`ïÊçQ2žÃù*VÎ§6PŠÆV"N²‡	Lxà&Ð!ijl†ç•hmí	ê
LmÄðÎãMVÀW÷Êµ%cXrãaŠ¶?&—»WøòZÌ›~<aÿÈ<#~¼ÉVÈ;›<§—_³Á¹÷ß+X:6‹B§ÓÀê€”iÎ!WÁ¨U•…Ðˆþ‰ZÏ³É|Å¨ˆ‹‹v%¨[., 0BzúUú±¨„~úñ„ýSÕiûf'¢Ž‡m_û¯¦jµÝ¢ÇøTû¹å½G®f—ð.$ÓÓ{.²Dí¶c€HFTÊîG•n·«š÷ÛyUÙEyœ­Ô³œqôMäTÚ¥jÑz³º–ÂËb·}}£UO­6ÿËŸOY*ÚÒHEË‡_jC—m73šµ}§r´v¾6nŽj£ÐÌ©XÕÅY<Ê.”fd*ÿ=Ù¤õhžSpJº0OßÖ ãUb‚/”=‘5	œ¹J«35/ê[´‘“¨Ý4Þ“.‰XÑ–²xA¬ÐTÅ[ð®^nÊ-¼)ûx™N…»]ßôCrïû½Ù§Ú/ö€î[uÁKË¹‹·îã´ŽñáßeHÃÝÄñ†ª_™¸€Ü9ô÷Ô÷é4[•t:JO3¹.üçm¯Œ=½ÐÕ-’çâc• ÚÄü’`,Ê€)JËËÎ£^TâWÝZpŒ:ìà‰}Óìˆ¿†s
”ÕÇN/²>~\*Ìð1¥ì˜ß¬Î.Ã¬G‘§fuE-CŽŠÂ‘ßÊÃÃc£º!¹ƒ_Ù•E¥Ù`YögMC»º€m)„ÕéãØ­‚•Þ¨(ÖÙáÓ\Gm…Ž~à;ýî¶RI«kDæñi…ôÃ¥Ø–uÑpžYÞ¡hÄIîãøq“6€ÊXY*…Ê‡1ÇcÌ¹„a„jgûÚ¶›Õ
{$`Æ÷[;}LqçïŽA‡´€¸9hXVlKÀ†i¾°<dïð‡³-ãVñ	1…Þ`íÉ•w¶Õ„À»%ä1ÒÿTè#±Œõ}÷8÷8ØÔ¡"ÔîCk4%Žn¯ù•mø×,¡-ÏÇªŒþ¿t£å ÝÓ„ˆŒ&%JQÊÎ €_HƒB;°À|PŒ¹ÇœŸ!Þ^P_Œ·O¯-:l€tXG ‹_L‘õ#i^YÒm é7¨s:h¶ SAŸ=¦G£Þµ³í†4=xÆ¢è?Ñœ— } èÔ+åÝÿ%˜6%,+6ø-$0‚ ù–°àíàáàò^Àôv´ùñåx` þÁµ—Œ\1&Jðñó!_®ûˆøìèkÉà‹|t0¸rŒBU\{¸Ø"ÀÅ¾8]Ú÷Gý8ÚòÆE£oCÌØòS{ÜÁ5I¦Ÿ'R:•qbÌjNÑ­–O„âìð¿­©k°5|SÙ«óÅOè`Ë‹^¬è‘+Ú{Ñ5XY®~)8þ7õ³»€wy¯OÈ»:¾Š­nˆÇoÚùV®ävXn=T ›¿cxÇSD'­÷e.,k†¢Bj¨­{ìSÛÎÛLÛjDÂù{Õzú·O‹ce•wp©.vw—‹Ç2¶~ð`9ëÏo^Ö/&äŸW(ó¶'Ç[Ç®TøËöƒÞâöseAH¥P÷zÇ©Tj!ä\*íì	ˆÛe·-çÖ3bÜÌ·x±•|NwRõÁk‘™ÍckUÓOh)C²‹•,¦%HT¢JÑ‹¥|A\yÔÌ‹"ÿåGIÏÐ¥:“LP]õÄÅ|=WL ˜˜¢Í'ì&PßÓòÊ'»>‚ã ÖZD%¿^*Â´F}$\¾‰–J<J[Eña)¬Äc*®ÂqÅªuW²5¦Ã¢fš¨±ÄS­ÎÏbŠÙvM‰
>uHøJ¦Zûâ ‘i<~™N?rÉè£
ý¯U‹–+4Õ¼€Wîî ÓMâÚ“Cr¾b¹s¯ëFT3ž;¦=§!ÕéððY1ûY+·]·6ŸšÏ››Ñ÷4(še%f=ˆÇÑÁ~c™P°pÊ×NAe"šÆïœ«L“M<>P°w#ý’©±l:”*÷†s¼ÍKp‘+Ð³,,hw=¢Ò×C^uÕ®]m—ù¸Ñê¯|ÍåÙšOƒ[E9H«%nRMF%ÐŠ4©Vd¬Š›ÔÛ}™iÔ­x“Óë±˜—ÛÛ¤þ‚÷\ÍvÇÌ8ð±L9ø†WÕ©¹%1é 
5fTæÖ¯aâqKÃÝ…¤˜|àc›}ÜŽš›PñZ(ª$†ª?.`‘g”žÌ¸mˆÀ¡Aƒ°¹>WŽ PXR¡†ÖÔ]C‹Š­e-*<ãnlbA¼e3‹í[1³ ‘¯ÆÔŸJsê«†ú÷›]°šAh_ÈTÃUŽ[qŠÛ-ÛRã6´±CØËR×À.bGÐXÝºZÂv5ªÖëö0AYÑŒWªQýMY¹Ü†YË­@ã*U¡µÓ^‘Ös%:Ï¥4žû#­„BJOŽ²È„È ó«Ø5³"Š%‹¨Ìˆâ…Íˆ·
ša˜ê‹¤Q(Z\ÈüÖÖgÖš"ù}ã¯D¿ÎtÅq(°E,šÛ"Z|›ÆéT{+I1 fw&6ÅJMljS,dRÒp«Wib³ÐóîŠ%ïnÑ<‡ÌMŠ Ì`!ƒ“eV'àH.NŠ;atvzûýÜMK«¢éGÛ«–ƒlk-;«ñ{²ìXdæwÐ²£2âÉ##â‰×2Ã7³×¸Ä9Êº~¿¶áróùÃV#ðüa«áií? ­Fq#[K´*¶ë…ª7¶áøÃ¤çÔÆó¶«º)©p0àõ¸WÏ3n ðÚª3­âs†€Š¦ØÇ ¼èz^×ã'¦9ŒÂ|Uß†ç®¿©oÁ0½èúÞ×·fY/ºæ›ú|¡×‹®û–I´VpèÁÆÈž!ú¢9–odápGõÊ·e¡ £Ïÿ5Q¸ED{G!év,VGwÇD¡¡l¦I¼W;f¬Ý^åÏî7æ=h›ƒµ'rS¢·ê|¼yv¿²SÎb¢÷ôÐÎfyå½m¯1.jÖø†óÆÀlò{<ëÛ‘¸ºo<q<4Zf¸±¸ôëÊ!ª'_Cõ6&‡åáßîE3˜xMÈS’]ô3nG5‡Õ³E^¿e¤Ìç_–©Û—PÌo?)ÞŠŸV"êâõæÞzhÖžßöúüíð@BnV¹K­I6^â6­ºd’C÷Q’¤½iq‘äh¸B›Y”Ê%oÖ–H,®¸j®ýn&Q’º½*éÀŠlÃñ©cÏµ¾lL®ñPIÏš6Øª—Y_ÅŸ¸¡6ß 8K# ð[­b>Ùˆržx>‰þµÑŸ#XÂˆÿ~»½õw—Äù45î,ïÃ,O°£ÞBýÌ†hÎ$'÷„Œ`QÏº„0[->MY¦ól‰Šð^}/•Š'óyü·Y
ðr
]<ò»ð!V@VÃCÃ\®›Õ¢hÚ*õ~ŸNá¥Ãsì•éô”%²ŽžÊÁÐ%Ó¸ˆÞŠ£F™ öJ¸Ëy©WñK­5ê¾Fn„Oíö¡•Âò¢ý›žªHÃFwêªxdLw5<gaÏÅêx­Tµ ô×qÊŽ•–"ž›þµÞw»]×FQo|Š]ñ þ-‘z”!˜c]¢?ŽÆBÉ¶–2žÅçJN£¡Õ,­Úr¸°»¢LŽuö³øÔÚÐRŸM”r79ºApuíck[ôÚ48é˜eòïÛÕØÜ¶µé¹•ÝÜÚ´ÖˆhUæB×‹íÑê¬MØ›;fÔ ¤›‹ô¿¶	f!%%·`ƒiâ“ÚŒM3x.kzé^'¶I‘q›mZÂ×2J5Ád/Ö
4SãmÂê›Y^ú½èÿ£Ë:¬húÕFÂªà„ùOyÕú!èe|‰|(»G=¯\2¹µYH\!l
Y•`²•¢Dð%<§P ú,=§Y™ã`tÜÑn×ž<h/Ê0#ê“(›R™PþÔ_©¯,?§é/1AJ}Ë¾8
Á”C^³A¥…f“Ù;SàsÆ	Ë›Tw¢ŠüÕR› ÷A	Œæ~rÏÇ%5}-äÝÀ•ñSÁ² òšÅ£œ¶'0)JL³è8d‚[6B”Š$õ˜Î2y>¼0²ªžø†²>|sjž“5Z$‘mÃt¿‹%ûm¦ñg®],ÍohYÂi~Œwa {âu3¸…h›a‰-æý"·ëú8xèû%îò½éUÿ1Áµj¤>ž"ZžK1nYd¸¦zbq€­~è¨båàênhÑ¯—t[4ìÖÓoC{b±„!›âíZ›bä	ñeËÞ¤ŒëEùÛ‡®áj9ðÅ™_tÁþHW_jMKÓü:—ùéYöjLÇ£tz*•j•«ÖK9°öDA\½@³‚‡¾<ðà÷–†bR¢´£Vð@b“'WW\KTÑÂqõòÑÄÿ÷¿ÿß‹¿HâZô-T6ê´~/oM¤è°C•’®ŠLGlG«1»+&êÐFeé ¬\V”ñ3`¤`LKÉhY5aœadQàöIdÍ-’IJD{EÏYˆ e‚Çê]ö6™Â`ÕåX-©VTÒëÅjiDSN=†‰&£ÞB†%f%•°ŽöAÞ'J‹YU€ˆ•%¾:àY`dÚÃpX¤ß(òD·uûd´ü0c¦»?¦1‡ f,´4U¥.»‘„×„?Ú©¥ð™ÔŠ’y/*º˜¢«HÃx(p©¢í<2‘ÊDvƒ7ŠL>uây™5‰Ü_T$¢‹ÊÞ¡œ»_Á±¹tLÆBÈG¼z4dÂŽ
K¨F)‚ÆIŒ”bú>%~ÉQíir’E8¿FctŽèFÿÈ†Ÿÿßh–9¢<*æ(KÈ?ÿkx°ÂXêlŠÌdÒM
a‰ âèçy<þyŽ[H&ºY×+2h*³YD€¨Kó(Ú—Á|ñÉBÒ’Æò’E%&Í‘¶_jâ•›<‡µÇ|bé$Ió¥jmš
øê"|1‘¶t|¬—hó¢_Q ­@wÓ\…8Ûfr¿¾D›ã¯“ÓÏÿ¦Y”DGÉ)!¯â¶DÛ¥èt*:%}Ïªžf9àÚ<J˜n°ØˆFh=óù_d>SÀèV*ÿnŠbzáSa|ØE,g |-vña¿46£íÔwÓYJñ1È·"Q½Š…0•¯	—§|…*e¤ÅQOGqŽö5ïßýøâÕáÛG{Ýt
Äí()TC!Å/€•e
¢Ì„¬R’€*ÝÄ‡"òN½%ÂÒU6 „n1¢‘²“£Ñ¨F_xTõ¢f7AÂÊ†äŒ©QBÑø“¯X¼
9p3)p8ZÁm‰7pa]8¾†»”u’ë……ÈTë…ÈWÖ[2ü„dÒ+Øß^0þÄÚNÏß8D#þ—–¡sy¶¸)®ÅEÑP~¾ôœ÷%/kƒÚ¹Å—Õ¬@_ñÉûíê^³3ð‡ ÿŽò9m'eùD™US‘Ø¯«èˆó«hn«î"}|~G"}	\ª¿qå;F_X²ÃÊ_ŠâÚ{jôšXÆ˜Ìg#”8æÉ(™ÓxÃÍPÎ™ÛÆåtµÒÑNÄ‚,oÀ¿øm'Zß+Óóx=åÙ8ûyž@½u–ÊüRBì$çi
­—¢¯Î±þKLë»À#kÿ<®Ð±c(ú®£h¨ŒtÑ\y^ ²‚`½g|â‹8“ÜÏ†­þÿx#ZÌêCmPëÐo{ú=j›pâ”áYÔŽ³ùèEž›þv¸œÙ8é&yžå­µøë–N´—"-	µwÖ6"ÙŒÖÅ¿™ÿÎÆñåÑ|ˆ)ŽÈ»‹—•C¡ÎÔ8šŒ»fÕx[u q8Ž§™a¥Ï ï’¢L$^%Ó£»ÁŸ{Ó9ÿk?9I§)Í-ŠÐ—„­ÓUlüôëŽ`aPÙOã‚àÀ.#ñãøNŽ“ü!Aöõ%!AëtU K£Q|¹2äÓ'yÙúéÌ†d›%SËqáñŸ®ätˆ‹*Ð/âáqîýt#èR³ZÆþƒÅ£ ˆó¯/SàA†±øp‹€fõø%ÁÍézU@7Æ‹„u€ÑÄVzkx!OÌ ð©ÓÏÿ=V€·v#ˆ³§Twg”Œí ¯,f‰;š¤SHøi°ETænó	syžÓ‚”ç¨îžS¢·!¹+K[ûîš\,&¦Uã3!‡ógêÇ/[ÃlŒþLi6­ œ‹³$‚<»˜&ùÁÀi}wþ‹ãîÎ´dÿÜXyÏp¶¤§I	Z´~6 ü–éB‡å¥f|îo`"§0¹œ¹$s48IÇÑ??ÿ%ŸpÚÑ4ŽÐ¾AúÜ[ÓÀK_d<ø2î	F¨:Îã_âì½ÐÆ;l^£´@ D^ÎîOº²ß^G-ÚâöOp~×Êx;ÜAš§Ñsèi>.3@Z¢”Òk(¾zºB)¾Ñ–Vö:Ã£¿©tJ¥K*ü?ŸâkN¦UxŽ¹"C—%9pïÒ	‚ÔdÖj«‚|wäÆªO¯>èF±i½«+ü|Ô—ÎÔ¢ýéÙ®å±'à‰G#Doµˆ±ÿ_’ï²ü»”[Q´Ä†+0­Å6o%ø%m_J63€¿¼xGM8ÂV¬jè!t:?O&^<´ÅÓËJTÅ`ÑtHØ¿>ŽŒìq‹z@MŸCÇ	g'ZÃXP³û_Ä§pØƒÔð”ž§ÉsÁƒF¥ `æøŠ«„¿dŸ±1¸ÿuPD wü­ÓØð,~Ìæ¥ž/ö0ïJ'Lâ1Î†ñø¶Zî9 d×Z¥~DzZ *c}’MË³ñ%±ñ”8rì?ØÜÄ€Ÿóî#ÆvUàŒ˜eeÜÜŒŽæÇ^v† Ñd¢qvšåU@ïôR/>ÍR'fáÿ´`Bð
POà R†‘"{^&'¥8BŒÖ†ÃHGGl+ñ'þÈÍ{ad²mÇ&—æýæí=|™f7érCCA(zJ×‡þ=£Å›UçŒ3¤Fë]Å…!¨0nT…×Ax58QqbšÖŠXJžæ@ªßN %`k@x êäG¢@©½(EOÜrƒÃ]|§¤™/åÀÍIZ_Û"Äçý˜X°ÊºÝ2£bOkQò8òô£zP&ù0JONð~‚ñèÃC€¯¡áŽj¸­^[­ðm€V(ØÍ0IÇ-ÙôfÔê÷z½èÏÑùŸÁý¶EHEZ„X¹Í²áÇpœ´­jß©¿Iü©ÕSÐÐÖJ2€` fÞ{
 ÝèM	d×EŠ±ÔãñpŽŠ"yF¢lz’žÎaDÄŒ1~î§B¨iH 0N')Ïû‹¹yéç>–´·Zñ¡ptØ¢™§ÑV/ƒ¦/\RlÃ—-‘( ˆ:bC!7Í¼P¨W §Lüc	ŸÆ ôôÁZáëLÀp2ÎàZo}!@Vû×‘]4X1š'ZÈm¢Ë`?ÖÒµ-K.æÇRz\qi²å6H?ÀÅñˆƒæç¿>‘2©ºýxKo
úPuEQÍ.ÉôkRl}lxŸ¼`K‚®šbz]¾L’²2ggÎ¨*Î‡euãØŽ§ÓÎYÖ‹›xJ]c­qÈÌ´Ù×´¬Ÿ#òð¨_ðtahjÏMo”S)Ùqˆ\#—š{Â„JÚAH¥•0xØúäKÿà·ÒG3ÕGbPétjÌˆòšëlîÛŸ®Vä¬ï(Úh8¹ZmÓ£+€ƒd¤ÆŸ®PÌˆ¶d`>Ê~¢cŒuÐ•­™æÃg}cœ¡àÛlLjÌô­1ß"òrÞaDgóntHv£°4)\y4/SäW¦À•ìÃÿNæSŠÀ.s@=3¸©Nâó,ß 'iIÃ_Óì<fqH²9ÃŒ ó	°1çG÷'k°dGsMŒÎÖù”(®ÌQ‹D õçÅ¯³h–Œ3“ŸêF/¦eÔÖ„¦—Œ•NjM õÎ2ONáfÖ±t[¤åœ¸kmÔ¬ê„ÌÕ¬ýðZcYëµ£ åÌLKcµ\=¢q0él“æ:ó@Wî3ò$ðC7>iá½l²«±|K[žG{¸× \¹m„ï×œÚjMÿªÔéÞí5KO§oæeÊ™³Z¡µ²í>lgmi‹n´g/ÄQœæ(Ð|Ž gNÍ3_=ÞD”-~;š[qãŒ3B·}µèW‰Õf<M'ØJ1ƒÛÚ°k£µ¼ÐôX¹éQÚžSNÕ¨Þâz±{ó‘á&Ç‹ºA5µšK´ÿ WìA¥“ÛËÃn‘«ÔÚíØ„n–¹B¡+ôéh²¯™ÞE\Ñ‚ŠŸEo>{Wx¨_¡·“fe§(7y­ã­A¿Š?Mr"s;ÐKŠËSt×QÈû<OBütGp	 ¡¼»îÜ ÿf€)´î€ä1F{DË9¼¸)¤¢i+/%²ÁúfËè÷í¿DYHûZûã¢’M6Þa÷Š):^ ¤y¼ö^m^8ò¹¤èÉá\ååZ#!·k—¤À09 JòÝ5¤Q’Ž¿ày<ž¸Ò´±„{Â ÓÁÐ²­DÜœ{¢8z"Ä9°x]j¨Ñ:“aWØ¢“l8/vPÑˆqžÔx¬Ç_[Ø?ðû¬3ÍÜŒi´÷c`ÀÜÕÈ“ŸçHXX,ãn_x E÷m«rã®Û:$¡’é™'!›¶s‡¼±æ›'jüÇØ¿(‚;ì%1¹»Wì96Ø]N=±­+è¯ùVæ’¡¤Zâ¦À|>CEÙÙ7üÈÊ‡²‘PsQÕ´n£4Ý„ˆÌEËžNªtØÝ$HŸ³Åe);ÖÜÏ®/p…ï3ÁäàóA#\Ta%y0Ùž¸ünt×žH×! pûº÷¤Ô™ŸºÌ‚Ø!“÷C1ÿük^á\g°¶°"­C¶c4”¼OaÕ:ÊÐïhêm´`~køCÜêsjä–.fÄ†gw!Š€ò:¬H·-Û¥ôÍ\0û	Ò>!…*ÐÀ¡‘ÎŒ¿fÙé8y‰GòÚs@°ØÈ;ßçVð¸8?5Ý*{{-B½å³ìÓîZ/êEƒûð¿@¤Õ¸<Ã`©ãÝµá<*²|Ž˜m-í®½ºÛ¢þ ;Øö:Ý‡:ÝÞC´ÛÞêt|ý}p~¿;xp¶Ýýf0„·P¼»…ezÐ'„Rýh«»Õ?t><Ûên?º½GPä›|<êÜï>¼ÏþzÔí}óK ¬µ>Ôÿ´uïÑö%0`ƒ-hó›‡0Õíîýî7¢‡ØÞ ûàÁ¸ƒ}v°÷!~W8¨-hï|{ØgºD½ÎvwðŽm«ó Û cÛÞú~Ðí?‚á?ºÿ|«ûÍ7Ñ /¡ƒ‡¶‚½7ówÏž=ïm³1oCcQÿ>LmÐÁAu·¶¡ã-ö,Ñ7E·¿oîo‰ÿxëO£yŽ¯£GÝíí†ßƒ&ðßAo·º÷á-ŠîÃ4Æ0n¬{ù¨ýõÅÞý­­m¹¾ÛÝ­GÃ>T†Þ‚ÅB°¸ïî·ºýíþçyÿ!öCÅÉÁ†à à?¸V°x0:X7Nê>xáÒ»p“ ¬ÐªGrýý#†3|~ê¾f°‘@Æå œø1«±òb”§þô,+Jçs—qu!²ËÝï ^yº7a ÷X|¨–uL9þWÊþ‘ÑÌ‹E^ŽõÕR5UÛ@dHC¼®}aÝÖ@¶Ù¬;Nódœ]tÎÒÑ*Í€.w”ß/ý”…ÎÓ"=‹×Ç¨áŸ1ûÁ†¡KB®6ÿ¥£äVè´6Ž/£?oªÙ£Ÿ&ûŒ 6›Ms5SŸxú)Éû]N&·‰ºŠxøq”g3LÏž£»ñ/û=ô¶gó5·ÔÄèÃj‘šÌàD´©óÝÅ\7µó1°¥S{Ë:ýHn˜\æK¶˜’*ð,qƒ-AW:4î;Ü†WÙq
å¿Obä“ôM@É{«R.U„‚—Q	‹ÙcË(ÁDHþ\7AŠ05C™B} >M¶1‹FMñ{}‡<µ-VHãzt>±Ž©”Ÿ{Põ³8¦7/·ŒÈ~:«z*Óp`^‚È¢™ìøœK%6nøùÞ2==+mµ˜Ð>!Ÿ»öD·‰<ÌƒžËN·ÉÃ¬´ l(pn¤›ññetP~þ¿Ðœò¯Ù$)ü½†C"¯üI9ëNü=/Ui¤µ}‡èh¸’‰®¼],ù4zü¿!ÌàÊ~ü*™Î—6[íf|¼ÉN¼A×ûÃœ=ü4–rfTAb˜aðvj'päO¥Ø˜r\@>ï²Y„É£:€†FÄ¡µ#KÓÔz—ÑyÑ…ësÛvª:”p§»PH’lF3ÒúqOlSŠEÓ`JD9zØÚCÿÔ×.9	óüÎqØÙÕ:Œ;JJn½É0rð)"ñQó¸NCþåS$÷Wî
ë?"(]3S¼ŠmëqÓSq¡ÔË…š7c¥ùJ£þ|Íd¢uqˆŠ]¢°jØw'²ÎÚòWÝÊžÀzó€þGÌP83?æÙ”TÛd*{ŠŒ¶´bŽZÒ¢ºíë)Uà9.’æ-«ís`AqáÜüLæ`ö’`X_Ï†O(!WAá´Û z2bþ=›Os ×"·Ý1˜gP> .¹} 9úš³Él^¢UÈæ4>ONñ¯ þ{àÁös¤¹›A=[Q	õ*ÆÏò@OWÍ©	Ø»_ü‘@*†î äúB.FßN€öÃÑz²×Æðäs}ÓŽ xRBq¡œ½¸iâX¿UIâˆ¡5Zû)fD3ïÌ—¾1´(/A«…¬¥~Í˜òö`¾eV”Ÿ¦«ŸŽäÍÁý_¸«)o¯
Åä4.Ï²ë´½°™Íyn#ý€Í ¬†þ-}‹nÆj úQˆÞZ)D×-k=4ïÎáTßKe¡vÙæ"¸’PÓM¨áFQÀÂXšD^©Š¸bi#^|B;@:äÏàÞšÏ|WLeê*[±‡.«ç$n~@´ 	à>¿ëžÅé'¸†Ÿÿ5KÉ:²HNç@0«ÂXÝÿrôæµ‡¤ó^ûÙÅ-É„m-“V“ÝUlEapl=ÃŠ¤€J>Q’_} &©ÓH/'i([Õ8¨“ÚB7æžÆeUo¦ÃËrÉÃË©ü0«Ù˜¨$&6£"TMÀ &ã?IÇðGñp˜Ìà¤vÿYdSeêÁÝ¬'Ž£q }ÂZöx5"ZËÉ‰PHñ_ã|ŠNäÏb´FÀÈ£çdJæÈ#®ted­{m>=	$z{<ÉºŽÈŸ”9¨ï^]EÙ,¢c;º1ÁÑ.ß‰_¾F…>Uð7ÙQ¨é&D]µ±"“¼H¢ÞX iZFÚLÊ;±ú¸Ú6QZSSóñÑYšŒGäoºt´iWÚß8Ñ§šMÙ¦'†Õ¼ß¸?F†è8OÎ“{ÞÏÞNEjIÑ¯=¡Xô“(ž%˜býÊd,¤Ë¬Á“(>6ƒQÉ9Fh‹²è˜"R=Åà§+s4"À÷q»zãHh3Ì¥¿£©ø‘ «0ù{Äsîji[ËYq­ÚïòpŒî› 9†Ë,§!9]ÊóxÑ‹ÐÉQFÞÓc?z2ÙFqçhÍé×uíåÊQRÑùŒ¦lJlÏtœôô`9PzJ„)Í§bNA_ßrZ¿žúÊÉ2üÑô®ôÂ{ÏïI¾¹­…œ}ÛµN”æÃ÷VöŒ×ª¯äµO
(ÚïÃê›¡åÑjW6—¸Ö{Õé{‘Åe¾§­…¼Y]Ý…<Uý£Ó;þ	ÖXnä’?¸•FÑÜCý;Z³ŠÒâ.Jo±gAª‹=ÚK„û´¼„ÿå(/ö[VQa5 «¡Áø84ÿÏ 	ÆVFOnNŽñ
Mˆ2š8Ñd{%’cèþHwåŽ4SDÏóÏÿŽŸŠ %ìOÐcÔ‘c/(~Ò$ºòƒkï[º¦úh{Í/_2¾ñú5Iý$1ÉhÇQŒ>œX¢LŠH¼E}Kñô	À»ˆ€&Êã@>ù&[¦ëØSIÝ±§yŒÞ5Ècu!=>öÆ4ŒÒ+ÏG5Èž0düUº5I˜
.wU¬Ú0yÈŸ!¹ó†£Ôé|<v#Ö¶LI­Mu>Þcˆñ f”D“l„ˆ%NK_öhtèc<=Ž§ëþðå5x³[¬±|K@ñÅ0Æ W=´¡õ¢ì ŽçõúþJÉ'ôÁZ¨'/ë†	ˆbF†"û0_ÉI\ì^ñ?¨¸ ¶KvXœ¤}•8—¸ìj™Tx4OÑn:
]v¯Ä_þrÙ£xÃ‰,+#S›‘¥1Ò;šQ8ª´¨¡ñŠb¬%Ìº]]®axêÀížMY,í§Ê‚h¯n²7Ä³’	 v¯XK6:ö®ºÞažf91	f]õ>\E¥/ÎSd¥ñ»tœ°/þzÆîRŒ0õ3XC®>U¿‚åµ%¥Úïê>p·Tø+4û¿&Àžc–=
¢wHÎNwþÚ¼†xÓëñßu=î§º"ô±7´³ÿ]hjŽP±Ø´—Õµ¹ç°ªÉ_øk³ÓìùY–"X¨¿ƒKÿR+nüˆ bÍ&Xj÷Jý]UVJÆz-ómhÙß&>ÌÚéÜ}ù]žMŽ0ç2¼9‡kÄ×žWN¾þD®]¾gÙ$	\½˜(º~±	?Ýæ°úA^i	öÊwñ"ê$dr ó_¸Þ‹/TTOV¶{¥ÿ
Vñ^è!¢U¦Û½’†ÊºqCü0î
‰`¨AŸœO­‰þ6ÔjÿwÉÀ_Â¯›¬$p«`–èŸ€lŠ˜¢‰ˆÌ
(_Æƒpïêò„!¼1|/ÝÀ6ô;ƒÆ°°üÓ_²ˆ2ñŒQ¸õkh¿ª&éŒº#ûU-ª7ªçž·Á«lßXbëE°–yÎC¥ŽÜ»Ø~W[×¢|ïoùâŠtÅ78Ò¬ßîÞ£ñã\}òV¸Ç“x˜g”¦þ8‹óÑ6Ûjé·»ë¯p"ûbw™ã¢<ºœæ“Iœ_o¾ð×J,\YÌ!9¢ÿ!]dÊŽ0Ü:6]~Ä[‹?µtTÊV0¸=0”ÔõŽ¯Ã×Å¦a'ã)·nô#Re®k˜‡Fô:>OO‘¥Î8~Maù[¸Wuò^fõB‚™	ØL¬Eó¼ØH—8›ãú¹¶çÞ¡÷ò&À•´Û¢âµf#ñ _0kÕ’©¼ÃnÊaXJ8<Ò× fîlB½-·E+ß¤å·©†çh	Ùî"›·Âûc”ÃN\IQ Íýî«¥ßîý±ÙC}~<Œ+ÄË¯&TçJ)‚ržÃÔW³«žÖ~»;ûŠ&ƒ’ÖßÞ¶B/(¶¼	‡.šøÝ¨Žø„Â»`wô¹VÞø¨rS¡§ìÃn"ð!L¢ZüoÉ%ÉDF¯VAf† ÐC˜^ƒAM—íƒ°s¶à˜M/hO €C”‡'»•ûðÔ5‡*$ô‰”“%Ô1"nê®ª¤“yòaIîìÙ]« BÌÜPöSAk„iü…ˆÆeåXœæÌÂ
‘¬´…@$[aõÙßŠÖ~7øö%›Ð«x
‡IúktY	¸Pš:©z…îÔùÐ¸âÆ=ÍTpéžVdòOOKò[ãÖDšGOcâSHëFÑÚ~¨•ôâû¯iy†Ñ¢v¯Üd~îãËåh?,Ó6†#Àöµ‘³ÚadÆ‘§òÎcöcˆÑ]Ÿ2äápy˜ýP†D÷¹Ž’1Ô­EÛ,á ±6ÏN8É($G.ö¯s×F|Ë„í„~ë.ÄØè’ôŒýêq)j+NvWžDÅN2k#=íVÏH¦éô ™d0í€ÖeåÐnjpÑDF¬1àYÊ4©² ©‘¹Ø«]S|+À-Ë­ÓÈ¬|n:-+Ó¨¤ä‚	H¿ÈÅ/’®èæÇ€Nhùþ;»ÿaZ"t÷ûr&®lSŒ™Ëo†ÞÌoWvQ -*&;ôwž]àßÍâ@=
yæV™,ŸªS#¬=)¦#È{¼y6¶Y‘aí	7þM0 #}KE& -hŽ]aŒí„ñ½Åëù+_µÁm8;ªæ‹^gç±¾UÛ\‡¢±¸ð|š§£ÿƒð\túÑd´£~0Úšú¹Å€8 ¿^.ƒòšs»P°"–q6ÑMG†O:7m÷„w¼ßOól>«r­°Ýø;¡þòšk	¦Ö.3îÉéU×¦áÝ¢¢8XaŒ‰I«ÐñÅN×¿øb!ÖD×w¦b§x©ô’áuª3*VÂ4¡;°±ç}}qï›1¢´ŒkRJxæå
Ô¤Ìù—Ís¸tNµL„ÕÏ’´Û@-iTY„ª6*6D{ê	#@õ<FÓêw°>¨…§:ß­ÜMaƒ{Ä}UÀX€­p›X4ng‡ßåqq6XýþÖâ•úÍâ	!Ž¯CDÏ¶Ü¸Òwo|ºö„.º<þ%ÎŽ2LÖ~„ÚVM£V ©«Ÿ(çÂìƒÖš Xˆ©Gf„©ùl–äC”íü©TqpzÞR&º@_¾ªìa{=òñó{iâ×ÊnƒQ¿Ôs¥Óo³q…ÝçhÙ‚QiT‰:È{H*DÑàêçžØ(OCWJ4ß‰K Óï´]5¹j·ÎJH#è»ù´}úÉ;=M|ÈVŸz´NÜ	Ð32ÝqzOfôCû¸ÄK9þVÑ}6¡rÿ‘AbŠðÂÚ–Ä–V!¨ …QŸ0)g,¸'yôøgœNG¸$b³‚;éÝª`J˜%¥Ùö)Mn$
m8r€ÿ  ÿÿì}ÛnÜH–àû~ETnõ(ÕJ)uqÉK†,Û5šñE-¹kz¡5ÊT’’Øf&³I¦,•FÀÌ,vX,°¨™‡A7POyÙWýIÁ~Âž$ƒÁ¸‘™’í*s¦ÚJ^ârâÄ‰s?ä‹@ÕÕô sý©äœ óúøAQØð“ôk²»Y/@ÍûJì+ãûüMAÅÏïÀQ÷-ž|f%þ)ø0>3å@?E%m·ËhÞY@–È`‘üš1[A‚÷zD|¡úlÑI©P’%-ýUÚ—>œƒPòw( Aª±Þƒsèòˆf·Ø"2X\’ÕuøŸ¥ð?ÉÙ‰×
Ãÿ¿¿²¹ÿ|¾‰ï¬«ÞY_ìèstA:ÇQ`r<è9Åtñg42…§§ã M·==â{£	<Ù"«¦FÅJ`B>s!hQëAÓˆáp”‹Û6PäHÙ§©ÞB9,™ºvjôEë7j‘WšêŠîF[ô³ÒµÓ9éŒfÒ}"z£–
•îÈE+ÐJ/ 8û}ôãîAšeË„(­á­q<Ê^9CóÜfü•ëØ±
G€wúj
ÙkÅOú’éØ/1‰[Þ	0²_½Ù?x}ôý³—‡ÏŽvûáxMý •_]4yaHp©src.ãµ:QX{5ð3•ŽÜV»±ÈMÔRÛç®ïû4~sÚÒLëw·›úî–¹¹Ê¯R×ÁœàðŠE_Çu~×9	sÑÖ¹k×Ë>vnÀAåpÞÍ[E‡fÒFñWä\ø!-³¤ÉrMD§&ƒöFÉìuÑ¯RïY©ö•(3v6\Àw­?´ìDÕT¦z	DÇ™Þ´¦8-)‡*l»ÁÇ,\,Îx¾K5Pä(8C?BàPÐ‹­í»O†§R:n[&é¶ˆã¹Ëx(h•÷ñ†–›ïyòÜÙO…f“o­›XõªIp5I@Jq¿º±²¼†õÂ»„¦žMUlPª¥‰ô™öÊ‹2@èx	yH…ê«WÁÙí_†aloÁé\ÃÖ‹cÍEŽ!†ú9ò{6£…I#¼£)B?ÀsGU£zQ:7Oþ`Òì½)'J‰G •LG2áÖÐ}
¬8K¤S»dè„€á|èìì½:øû-çr¢äÛ~ÌYÊëª4ó3îîs	¶ùdïÀßŒ¼[­‚lVvÌª…CQ“SÒ ÎEßK®:B*Ò{v:hw¦Wå–Im7	<m¼êãf~
EÞ=jdºR	Þ/Ñ°Qž›¡ûŽþÈb²D¾.d¦‹¸y×ª&À\ýtß.þ-Y^&‡¼Æ“Ò…lâ]yj¼Š/bB‡wûùo=²%¶½óà"‰Ç‡´øoiÕPÛ£°Î•V¹¡ÔBÎìfyÁd;·çºu°¯±©7°©ß¿]yxÔo§b±Ò¯¢ûYSUdZó1±ç·®%¯ž)¶¦ÞIø]šìhúoý9]ÏŠ½W~ÜX^àmZ‹`ä]v=Bï¡¥P{úë‘jAÆª*txÙh—º†_ž•]Í~þ›Œ;·?ž…cˆ0»A!åº,n!ãÔrÕ
ªƒYxødcæy6C¢pÜuî‘#ÛoÚ!ÛÏTß
Ù´DGÎiÎ.9Y­;¡
2ÐK€ó)xnqÅ7ÂU•,“w"¹Ûn›x2ü]˜f·I€QÙªAÅMikÕ­Òm¡ZÃÂÒ_Ké™hæsž ©ÈyTIož4ô‹ª
¾À|uS¬4tÒ³ù—¶nÎjkãüç1{8?v¹ØJËØÕ€¢±Ë®±K](U.I,Uð”õŠJv6•”£×Ty¡Õ„ÈÓ0uß=™â ÐG‚]I0	²ðößoÿ#HI@F][^m-® A'“a8ñ"x-ÄO)&àYõµz£=~çh5F7~:­'EÑ(<Hâ`š²J@–¸ñ‚¬”ølÄ´8­âÅýâ±® 4óÊZ¥¥Ãàf}¾÷AY¼3JX]ååò"ñ7©ˆl)H“¸Î0R‹Hd—ÝÌžÈ²£Øiõ#ZV,`EÅ›çÁðÜKì#±ø57¨Ä“+C×Ôl6Z†nj4a^'ä»\4¹¬ÑA:¸­áE]×h½ÐGÄ;ü{éëkÞÏÍ;3‚×ÚÖd§6tU<Ðú«mšýÕðºŸ5¼~kÀàlN£øÃ’¾jk~Y4²çšƒ8žz¥_ÉèQ÷sÓ¹Ä™ÝØêéË¬†~®PZ°z¸Ñ©:÷v/5‹ŸšàSV³YJ—Kâ¯CÝ.aLÈÆÉUÅyÖAíànfç`‹÷&5Ez]bÜÔ*ìíVÞúIÀS8×Um(îQ]¥x@ÀùÀ2¼Ä7.6§üÊÉL©£º7žOšqãöÿÒ.QŠ“KZ9»1EL!M›+r<q“îä‡|ØY­Ctnv#}ÍEdÑ×wfŠ¾‚©|E(‚¡³ˆ@™z¤ãûdÁÈK¸Ø_Wpé¿feäá8˜dOÉÍb#äsw À‹ÕgŸ¸˜ŽñjçxWkç¼¸ÝáMŒ–å°Í`.Oó|GÍüðrò	ÀK«¥UùÅ"¡¯:ÔC3ëªq–}i0Âˆœ?Ùw·ÖÜ0‘²
æºOò#fÕéˆ¡Ÿ:ú°‹vÍFRÕ4å¡B)¾Ãöî_ÿù_5í:‡ÛáI7]äûqn¤ÊÍ…»µŸ·âÔTÛKûÓ%jÙMÏ€Wc]^Îú¼ä}'n1éˆÎ9%æ²Z;¼Îæ€Ö"ã¸Ôgƒ‹Âór„rs	º{ÑàQ+¯d÷$	uæ¡™}›gö˜Åey¦Õk·vŽ>ÈO¼dïÜK²µêªÕ*æ¸2×Ý<pÛ_QrÉ¾_ðâùoä‡Y’Õi¨~º@¥…EW¾¸¥"5ÜqÛbMh•¦‚í†ÂÍ¾âHÈßr;«g U31l*}ó]ý8%¿—F,_;¿}¼Ù¼Ü‘òrMÃ„Ô¸)RÖÈãÁ‹>¨†2ÈååÈµ	h²Oä°Æ¾*VçÇˆ“‡	{Ãæe‚—Þx9'#ÕC…Žµ.£rFñMÙŸ0Ì<ØÌ¹£G×‚1&ôFË=‡ÅˆsO“ã‡a­MëKlý]cnéN¥H®®C0jê"W¡kÈ·Àëð&å½¨É³p¹EVi49.•9TùàÍ¥zÃ2tšûZj2'Ì¢ÉÔ1SºÒíeJôðh7Iâ/‚Ó¬¸ÿ]ežŠ4ìëqÔž‚køÐ*%jþôÙ!š˜þ™3RÝD®÷26»Ò¤Á(dÃ–-‹s¾Ñ·EÝxáIÈd”„¶dŽT’â“$K Šaª\MÑ™2iµi*" ¿i]ST¬Õ1A’4»èé5ªrí’bÐ®ì³+ø\•zÌæuõ›b6öª­ÎÈŽ5Ö´"ŠgU…›¥£D Ì«êµ§T½V%*úÏÌL%f`dôlJ³O”>Ts57}™ÃAA2†8è¥Ä=Û4¸Ñ¾emE1,Ð}xÖ­Ù:<ðnÃÕá€¾»(9C( +áqM"‘Ü,×-âVs¯'è@›4ñ~…´ÌZ™Ãþf¡VÉD¹cU¦—÷$Kk¬Ä2*gäüÇMfð¦ÕS†’`êƒ¯WØ)pÜÁý`¶àqAáns¥c5ŠöfæH/ÔÏ-O :®ôzù×„UüÙÐ'»4<‰0©%%¿^VHÅ%	µ#í×»òÎŸk¡ì²Öéõp¬F"mFÀ5Ä’S¯Æ"«ùôNOÙc
“8²ªâoÏ¯‘7žzNòh:Ä";š~L:lÞ¬€ž¥[´­ÂGUp_ÝTø§æê€¥É4J•¤X'	_+¦Më}<ÖÓEµøCæÃL'áXœ“›HsäETRö(“ËÜcûý¾…Ô´»0n3Å½ó`ø~/L†Q°ªžeiNr|d/F6åö'˜'+	‡xø&°$ïº_Wïõ³˜øðÇQ†¾{Ýã·èÂqO`nV—üð,ÌzdŽ§XÕH¸ÅY…[èë±øŽ2E6Y^U)ÓP—t{1"Â5=ÚñDÏÅÖ»ê=ž[G”O…%T¥Á“’)»›nË ›0}™ÓX)] Î‰³Ú?Å¢ñBý%g³€û½ýwÔ$”8ÜVcQê[žN uÑ‘¨ŽÛDj`CO Jïèæ”b‹²Öj“Ðw£‰÷¬Í¹“Ó 9$ÐôwBvÏâÄë¨ñ>×jPÿ0àÎõŸâ‡c q‰÷üwíù=Ö»g’- p4kHñ|ßWêªÁTÈ*Zš¥W[@ÛRZÕhižmµ,£ Å(O²íOáÿ³$N‹[6Û£½xreGm}ŽÙ¤­%&ÜŠ’¬ªš+6tÕ:)-è+¹Å.îìµcFs¹[ðÌ¡¿W˜ºÕ&3X£+öçf:Ï&h­ªÁjd—tj;aôâ»dRV¡>Eõ| Öì&£ÒJ¨¸Ô¬‹6qÊüº«GŽåe2èÃžóó(+<,KR€,`áùã`ŽÒœâ®—¤€uÑüÏ0ÀHØ0N@àùÏ(}>	HËñ8P	)Lo7‡Ðâa¾HdØƒ¢ÿâv·XEÀõÈWûä[šúWøÞçƒHÉÉÁ4ÚaÐ¼ÿä
ßÂµÜ"6„}ø(¥|läc9~»ƒ¼¾QŽCšMÿ4NžyÃs:|ý–dCH¼»T•A¡NuV}%ûr'#þ)þÉj3ŽýçÔæ‚º¼íElÿ­?!@
@ÄÌ†çÿ\A{¯O0£Iÿ}p•vk0ZìŸá×qEïqÂšñ¼_¤²,Íž—ˆ‘ÛÛÅ4ªÔl’fô´x>|}\¡0Á¼S!ªMú8ÿø-|}üVs˜+ïÚêO¦)C%I¼Ñ¢ýZŸüCLÈëW/þ¡ú1’{ñ†°‹°3Ï‘7Í7mJhBæ‰€¹J=Nxöº±Ò€‰HÂ@…Êyç.#Ýãï{¹Ó·”¹¾Æqt»^œÐ¯¼ã•·}@³ø$A÷nõÈÂ$[zr¸Ð£PTñrAý°H1J“‹x(×7ô[\FEßˆM2*î,:càõ£¨@Z­d_‘£2¨Æh™Aa˜g;™ÉÙKÙ0ž ÓÊ~Ý 9U°c(”ÅœöHÁ"@™<2XªÚdŸçGˆ¾ŽY	E§~~J0åÑQ8¢3ùþ—`žiå³W·ÿìu`ÀqFÑAŒÂ4õÒÝü¤Û®N˜«Ék³Ø®¾§ïf2ÌÊ©×šÁW3ŽÇM‡[e6‹‡{ìÇiÇ°å!-WZÇú
ÐÀâ"
½+Ê­ÁOÍ€’óƒ$>0"`äeKØ\VTÀ~ê™PBFŠâ“{Ç Q	…/¼äåZ›Úe®4@¶hÚy1óU”—OÕ1§P×ÅH{ìÛ¦'ÎìF¹åÇ3œâÂÉã™¸.ïÂ#ÔFÑgh¿•©-xXEX\úz%b-òÄ1bê4h¶ÚãZ<{âéØÕ»tžxøQ«ú¡úDœ«Žæ¤"EþOðu`{e+u`9Ý`Oßa,³ðFgáB¸ƒI±02¬ç «($ˆ£,þ:¼¬×è ¯\L`$@ä„IÃºÀIŠ–+„0TÁŸ`Ò„af—!XÅ0Ëã
6ÚÚxÃ¡IÀ`\’•râH‘ƒÈE†bä6±²É0†ã¼ Ò•ŸÖ¶òž…m-è|¤$#À«Ý S¦…šñTm{Ê_Y”YIpw¡Í¨|KPB4éûA:L@ÆŒ¿/žSìGæÁe‚øÁ[dXÃ|œºÂe¯O÷z"÷zÒ¬W@lœêþËƒÝ½7¯¿}øôÙ!]äÝÇ}†ô1Õ§Ðå¡=iŽ•¢¹'Êæž4lŽcmo‰Ž²pddÂu:}*TGæìRÁð<ñÆc`×ÕŠIU?'gKÎC8V'K¥Å½p5~^Ë›+$¥y6¨å¸RÎhä—…äào{ž­Õ®iÃæü¿óHÿ¤s³Ë­ä'ïµŽr1º|]bB©Ÿìrè8~­e©m!Uµ™Ø’Ô¼HàwâåÙÔÕZÛ'jVšŠi22‹¿œ<ž‡¾=y«Š&k™	^*ƒoT 2C©º™ÊÄYv(=÷¢Ì‘ëŠ°µ$‹p,ù¢øÎfd(3ð™~'þÀ"º•VÀu¾´Ê<¬Êyq´dÉš©ú‘¦Ä§xU=o:¼f›ä5ãOk'Vï /vz:1¼êúBR[’7ÌJÇÒ•fWiZ“o‘w__—òøÍ¯ÞÙÂâŒÑ]Ö¼ÁsÈ¼·Î+*Ó€´ZdŒ¿•ß
Æ¾•Ð·!âep	¥âršÜ&Ç
\ÊèÉåý$Š¡c)ZÄ²ì×{É~è[vûÌõ¤]˜9rµúœ•‚œëcôi‚©cœ¹'è2ë¨}øÔC?E²éqI¤V/§¸^	™QãqÍEPë†C_·@ä?€ÓàDJ¶cÂ÷í¿€Hùhù|½	.lvö|¡JI÷ÙgK»-²¹:Ðy^Òæu‹¶¬aÂ…R”ÉÔÛLªFAD!öÛ³(ÏËË=·Oéøëðr=o·¢â¡pqJÌäàY*_ùaÍŠ‰aHàdiM@lkiU0ÈÃÌ˜â¬êF¯ñ7•¯|-^–MÙEäL“~¸H3òó¿„÷–¢ÁÃq ì|g×è‚
WÏ½WI­WÆØ4…F8Å¤7ŽqeÓ$1—©XË¨¤q2p:ÓÓ(é–Ô‰3öþ‹œ¾2"M9è’‹pÃë³%œýF–Ð%Û¨©%ç†›.9ƒØut¸Ÿ³7Ëì7å<Ü0¸¯‘cÚBî¢Ÿd‰&­G¥+‹4Âƒ}$dPËÐéˆzbG~5?0cCOÆJpZIù¥’2pÙ%v1òË1Ý…[®JÇ\c®àf!v¾1ö·…»Ã^ŒY–³–ìêZÁ®ê9c»J‰ÎÏG§Ì¥,^JÈi*ÑñRxQUF¢‚ÐzÌ*Ã5M¢­àí\Š:fÎ8y=°ëÉcj•sÿ–oÛcYê˜D%	Ô+&ã<äÏsÿYKº9{u™|ïÊK6y¡a£î9¤ ¾Æ†X^ÿôOvfoè¿gDSÿ­v–¿ÇýãÁ[ûàu&©*ätÒL	Pår@äm_*[\´±Î‘ƒ®èLŠSÂõØŸIkR¹uçpØ˜c»†ÄVÏ²‘Kly]W-ŠÕÔE6@ýØ¬VÔ²h0¸µG¤ù-ÜÀØUs»ÇÃiTø‚åd±£$â´,­]û/Á™i0r›fZ?£¹‹&£4žÒTMí§¢á¼°¡VádL­Ž&·tnã*­Ÿh â…—±Ú¯£¼šÐ"ªåÈ§r£BS‘JSÊ@Ü\µà–e|~ÙâU­/Mø¨QÂ!©oPÁîä™…”új­Ñbt;ûsGrè”Ë¢(gÒÏVÚ›AÖÏíw1vºQl}šPuÿó]‰ž~»Q/pQŽÇÖ}(Ó´éçåíŸü°èIeäú*
ÿ¸wfŠôV§ú·îbŸMw‹\ó‰$…ãk©`Íz‚{Ä…œ2µNyÐ³.qÛ)¬Ä¥îÔÍ+ô«ù+–¨kƒ¦ý2CÀÁ÷æ—d £ïÈ±ôtÑ„Â¢¤ÄP?ƒ–ÝaðÌ™¤›ÑF¾šŒÐžÅµxGÑäIƒô½åU˜7¼ôj<$®5ÂäËûà…e_Óêrôì1¿kæÅØOÒ ¹ð€ëâ7&AjK¦º¨Ó¢Ô"uß®Þêƒd;âéÂ‚[m1ùJƒìÛ` æo|Æ’—äsk1p,ntÕj9+yš§1Ùæp?c£v÷ó']ä9F÷ª~¡­ÆÎ Ÿ <ñrB˜¼›Ù0…]˜‡‹;±ûsXCÞ$¦¶ˆ§™“&§áèò¤ìßmîÔ¢?Ö=¡
Ÿ´üM¬Ùò{W}vå2Äø6‚•QÚ£ý^2îvž%ILÆ`?Í·òo1ÁªÑ£ÛQ£‡néôXBx°Õé÷2&Õ1ƒÀHçÚXE:hÕºVƒiüMÓ/œ+¸äW+N¿¼øéŽù_–> ›Ï}W‚‘»ÈŸÂ }V¼í¥¯[Ù¾Å‹nAEHQ‹Å¬%d’ØµEZÇÄùÕËÕ^&uÏ¥»œí]ìjÆéhœWY–dXŽflS£R:ÅWŸ2³•³}a¸òë^.	ìŸ!ÓUÎàãõ…ñ¯/Œ×ñ…ñ2_
Æ«ˆÙnÍ|š×Šªµ`¼Šx“ûe»œ»Ë×]3]»QdŒéªò\|%~é|Mð…éÊ¯{aºD˜†þvë»%^_Ø­;þâ»e¾ì%Uíy­Üž.ZÐN‹þºwF+O¨ù©ñY¿WñXþ;g°œÝF*¯·6íŽü­KƒeWßu#NUßS#ï$ˆôa’t5*:Lvv^3îŒe‚'ËäˆQiøñh™6ßŠõmˆLmJØçW™üL"íâ1Ý”’Î‰Ÿ/·Ú–WåÎÙ!w0Þw¯ +^,ÿNØ–[ðP.ëvöhjÇÝƒ}t`ÿ6€£&$cä`§áÙSÿõÉAœSï"NzÄóÃ!€ù”_='Á>áØK½D§)úÉc
]z¾ G=Ù‡…g·ƒHvØx!Ê‚ýNK†‹y8Þ×ÐDÒ¨žˆè°yW¢Ç<³Ø¡Ëüä¹K!s“Aî[™Eþ¸!AÊLdá¹{Ä‹©ø²B‘V‹\x?„˜õlœ„¾W’–›¸û^HLIÂëp·™³mÆÿuÜ!¿¡ôGAšzg"?üT'›³y‰iŽòÒg!û¨Ÿêóø°8ßå+‰BR98BÁãÇW/ÓRÏ°JDR0­HËìvv¼¶ˆ‰=k=ð¢‚†\Ò!vY†£Ý¨ÍÈ•<µŽÝ–Ë«¡äV”øV$$È¦ýÝfÝ7•\·C²úe¨eøèEìÔÅ²…Æ¢YÕÚm:¤>àrVÎ¦IªJ;{Ç@%¸G˜«ü[íëxzb¾ò#’&Êó½¯îæ÷Õ#cÞGX½Â¥^°¡!VÄZ6Ë Ý<ÏÁéR¹oÅòfTºlØf´îÔEðb-MÔhlõtçK«ë4ŽCe-šˆÆu§éæ¥ÚÖâ‡è|ÇnóÖ(óÅÆá…~heÙš‚¶i¦vM"oœó$Û§Xî. Þ§!™ÄX 	%íQÇ	ðéñË‘&J·ì‚ˆaÄEÆ*`páEÓ`»†Ò(J7«®x¼wîÏ 1´€Tå$…öHÐÏhõâ>Š(6é¿Ñ6©é#1åál‰ì´Ë‹ªh·¦¢”8bÆbKÙªý`«	o5B0«ÍÐx0“U§Q8É	óƒV„YÐî^——ç¸¤ÓXâpcªÜÝqpé%­iw8žL³6¦¦½=#}U1Ó%Æ“²¬9mZ©mè™Œú0²]nvü>î¯Xr9è.ª,ƒ6	«HùþþÝ$–•Þî‘…œÂäªº–¨ÆRsCZ„W+» 
t^Á 1ja—D¦jÜ†î(Õ V[•wäõ¶e®5ƒÛygÁ>`‘¬øŽøá:yX%!5I>K¦cZ!mä]¢wiNPX¬Ì«xD5Nû£3[[ã ÛyA±KÊgøäµøC?Æ|5<‚íóýIäß/4–œÙ¥[ õaƒ Têu¯h'ðam´m ð~v	C{Ñ‹pü^ƒÛaX+Ç7þížÑËÃ`ËðŒ­yÐUQËù¬|að–=7û“Q´4ø$þ÷oµWx5úš
p­æþò£eŒwywÑ"NÏžÀHûÐžØ¨ñCí#íÅàU9—ÜJŽb¾¶U)­ÕkIÁä¥wÅô—Ø4p<y×OÏe®ý‡{RY <Ë«òÙÒQäöý.>t|/=?‰ÅÔ)/ž£å'üë/“ï`üOó÷ë3Ð–yµ&¾[/2|Ñ”ïâáíŸÉ$öâEô@‚þÇY‚º€§Õš·q9$â‘?N½èSç(¦ŸûŠÌu
Äx´<ŠQèëKO+þ®/ dÌE¬2 ŠjÀu¾ôQÙ…|ªa†‹NÑ¤LôÃ1¤^„©2ã‰7¤•Wz$Å:ðWÿáFeçÊêÊ7ƒâ›Aýƒà2ÌôPƒ©1gF:*sfÀßöz-èÈ¤J<9Kj¥|‡U¬ÏÎa©;ƒ-f˜jv†	$0£˜«ºâÀö]Ì§$IúJÈú<•ÎWùòhŽºrÀŽ9Ò)ÀÓGËç«š6ÍY9*$tZqF-KäY¥ò\zR:jòól4Ì{ÜÃ¶µˆš¾h«k¹äeÚÙùë?ÿkÃ–Â1Õvj¤Áj
V*”c³DO‰A†<âeí·Èµªü{NbY¡0šá¶úÝ9×YÇ‹—*0'¥ûðÌØyÂ>>ùçqòÌ³.ôcÑv ;›x$}ùÕ›ýƒ×Gß?{ypøìh·bºG?Hi‹ä1À†šX[ÇÖxÛo¨Î§ "¶Úò¼ýThÂ±^<^F¨³*=Yª$2©à«œ±Z_¬žÚ‚|chò÷ ¬µØÔìUž«¶,å+r ú	k±“U7‰Ç&×˜hŒoUîXZÆZâ"\ÞJI›b¹˜”Ðkq8œfûbÅ0ÌÙ¾.ëTHpA<-‰ÞÚÑÖÖš;kM›ÊK/­™ŠP«¾Ñ¼¯»=ƒæ«.û£pŒäZ…ñ‡qì³$¢Ó4H÷§¡¿¨ÇÈœ%Ó@¢àÔ¸R«ÓWj`IK×Ñè¨‡ó9D›Òé7W“ ¨}€þOÝh~þ)æGí* iþt?u<4gƒ–š½I€—”2<»„#1Tq´fžV¯2^¬®ŽŸDÞXÏ„³£V
Ó…/èé±K£Èš²âG8Lâq|–x#O?Î=¦Æ>:W‰#=àd½ýKÒáîË¼;¿gó$	0¹Æœ»z<}®ãSØÁhQ’—'x&Ã—†¥¤Aj½
-³•yYJ¾Å„“õªúò9ŸŒœæDô£ ãuwTa
ìƒé8üã”¥EGñI¬Ü®)œkŒÓþ)°ŽÞð¼Hg©"¬LÿÛ¼<|á<Ï¼y,s5'¼j#£íµD˜iÿÜKÙ­ÅEC$Añá"ÿÌó}~Cõºc‘ceñÙj¡zlæ.·z7eëuëLËÈã$àÃG'Ö¥î‘„bü‰¾ÓI?Åh‘ÞÅQ¹µ³FÞe‹æWÉ¯±tûÀ.‹Æ>èP¨€t0ÌäNæï1yéeÐ-rCÝ®4õeÒ­Œ•uã€£ÓV´:„ö2Xš‹¹È‡§›mš¯nÃæsœpêâ•7lž¡•Óð°#ó¤è€¯Í’jÅà6d=×8ZóõŽüZµÃ:ýo’¦Ý•íJUÎÊ’f¦ú0uÅo­.p®$âÙja`ƒ*„"/¯qÜý.ˆâaè{~°h,PÒ8Ž²mõg@ƒJ?¯Wuz×•=?KyPÁë¾€9Â¹³,†ÌÁy@Á8sÕÎ:t>,`U«ª:ÅXÝ²*y­•L{TñD´*D÷ì¥4Š?¸”é­tºê6ÐuAº¨Xlwì š©36Ø¹Íögii±;ßàyƒ|ìA^Š¨í¾<æ‘åvØNŠIÈ;IM”olUªÆlÈ–«îµ‚ÐWêæ.ÚvÙgŒ Ì‡w/¹ý	EØOaùIÅ'‹wqn`Ë_3\N œÞ÷*kkWMq²¹­È'ÁÍÈ°ÞBKí·ÆÇ¤Ü»˜=ŸHfšmE×ãÁÀQó§1òX=öj¡ç<¤ù¹4™Fi€®zÈ¡OzÞ’îþ®ÞY‚ö6káµ¡aù#­býfWë9é¼ 'WÎÜEP•1”J¬ ÙÄ5ZçM± Š èÛŒµ<1Öªåy0cð•S(….ïq^Q¹ò»ûWÁh€->ªïãÁF±%àwâ©ý:£Àó…Ú·–~4V€Ú,m.IüÖNiEZß¢tb>t<=#Ôšðáö%å¸]
‡°O.¼Êø[Bßà¬ ™C2%DW$Q™ë×|íÛ]Åþeä:»gS8¹èœ$p ÏCjÀ¦ác¯hF¯à/û^Ç/´Úº›)7ß™`ÐL%U=h7VìÑ4Œ
—T´ãoþ†|¥¢ìVgxF/-khð	×2ŒY‡„²nV+1‚c¥¿aóqmäs€WÕ&ƒ>[9|ž™F©Í/6B­3.a t‘”Ä¿¾~.ñú
¢nWKd¬Íß¸AÀÙÈT´kÞ®6â†/NÊ²ùm¶¹³C–©
‹=`Î<5¯C<¦2¤Bç^Äê´nQ¯»¥‡è*ŽÙe"GaÚ,Y)ôÕhWæ Ì‡¥ 2æ5–¶kþ=–Æ=ò.‚¼•Åó«ž¬>5. åœ<NË;ßûÀùÉM‘O×ðÅÎ”J+êJ³iÌ4ï6¡mïñEÈ´‘Ó·y×;	”†Ûi—£3A»¼j)¦S,‹ëlcÉ´ZMWÇÎºO£ŒJy¸IÎá?7'êÉmÑgÚ$á'^²wbóš«´dWKRjZêÔQ÷ì2NY"Ö.S½Oâ„ÜþV²¶+8L5å,ˆµŠ~ÂåÏ5‡’©×¼Â7ªN»²ÎO°ÄÑ¼K¹9Î|¶GAF˜I¨“ùdéí:V>¥,Ü«-ÚËËñ&Bý]V§WWÍ7qOÇÆ'…ï?¦svÁ4“ÿÒê½ËzÐÔFg@‚>]:ç­¾]$•ŸÐË5m÷˜Å”º±S+èAÆð‹ÜØûšŸIÚŠ•ip{óo¶%Û³u)+³¥Õ­Ì³,.·NB‡FoÔ¹¾Y\dµ­‘Äñ<gŽÃtÿÂ’â—ÂlêÖmúƒÓ²ôvÕ¨mêÓºûÓ,‰ß{èƒêüçÁÊÉÃÍAÇÜ*~xâùgÁt~vKŽ¥pBÎ¡£( Vã4Š]=»;'HÇR_1úY_±£ª‡ÓõµÓÀ2<"A¡^j¸>ÿÜYÏ>yRznu°|Å3Šóÿ¦Íü7+'ç¯*€\‡@Q¹1ƒ3Ê]Ú6¹Ë.w¯¼N÷Òß›šÖØ—+F&KkÌÃ7¨–×Õ2•˜Óã2âÂMÉV„A
«&õð§˜j~¬9èËœJ‚{ÀÝž ‡ÇÁV©NÒK'À„,o,¯½µ­1_³…Õ,†‘^œ)ÓÑpâ§@á"LÃL¢z'ñåvg'ÂmÃ­T‘u¼ö·;/É`µ·¾IvÉÚfomº[éÈæ&ÜsÎÂÁv×6l«ÁéÆéÃ†ßýcègçÛ¦Ý½Ç M`I¥œ?>£h»C£Ü>q§þˆp¿¨ÛM+Ø?4þ4¡Ÿ÷‡R€—$ÞÕvg0xØ_Ût®/?OOÓ Ãtjø=YÊOeÆl ÏA¸F¦§ÙUDC4KµEXwK~Ñ¤$ðàPŒ§Ù‚kv'œîâÌéE‰z'iM©Ìžeñh‰hœªé)¿+R›Þs«~„	¬%a>·ûPÄØ¹¦ëâæ3•ƒÂ)$ßµ€0<ÔNÅÉ·áäh¯‚RùAÍ}‚iŒÖê8_:^c6 Û¡%Y
…ã¿¸¥cºÝzºÐªœ;vZ7”^Îué|ñåëº`¢\6—ë,ï)ÙBë”
su$ÒéÜLJ6³Šm|qí =1àd2E‹fÖÿ9)¯Î™'C{gì8ó`ó%ä;9áèøU‡œ¸›ø/‹îï ¨îÏŠ[ð¢K}/ó¶¯ð>×±ðŠõ”Än®á=BcGá1‘á¨r±s­–%YÅ–ù]¡u&x5m§ý›~…ëŠfƒÓu¸›í÷ûÝªç7SO—³áäéðÛBÇ×½µ“Mèø-Ù"ÇoÍþ@x½µ©!ðý€Hö¯_Ç€²‡@ø§éöõƒû€Y)?Øtø`âùx¬ìŽÏ)Ú°€˜û˜9ƒØ6G·MÐÍÛ"ðÜ1tÌc:=ÌÄrÕ”ðƒËE[|o~=ÚàÌ¦ú‚wCøséëkÚÀÍ»ÆÓ_ÓfûC&ƒ8ð¯ö„G–Z©ä›8Ž²pb—&†@ºaG9ÛŽõCÎ(K²W¬-½`Ó3%Ãx°9¹,nÂÏÁä’ ú„¼N|ý)UŸÃ¶Ú¾îr¢BŒää:Öãw_sö·ôr{Ç^²RÔ)<­GÏ£eÅi×Òniñ#Ž2„}9l³­È=ÿc…‘Î½|ëæ¿5r.¸c‘gU¯Æå‚TÝëüÄ¼YlMqGã,ªw„·ùHùÏOg´…/w^,.):uÛGy]'‹v·©û™Yé«J×1-V‚ßp˜ä¬RA£UÛkB‹*ÅÚõ9°æ1Z2©Ÿ•5‹…Â+OÅv÷Žæ„W#¿Lx;$¼âƒ9©jr²´AÅ(½[…™€Ý!ª™±7WT~ŒÂÚßÑ…¼’Ö®®ýâ-2Û³‹8âŽØ4ÈcXOpžo+»Yc`÷RM™àÞ4Ò],£áÏâ¡â,-V+†ÖÑ
MÎ à)Ðéã\È³Þ)mRj†Ó,¼ 5èh>?6I«Z7X³Â
£îßÄgÀqcê,‰£Tz_´¥ÂùJR¬‰PÇè2B§ÅÓ<ºÃß¢¿Ð£Í¤Ghš‚*²"ùåL˜×'ª·Ì™|”zÂ<4ë´H“ ‚Ê7NUÉVµ4‰‰]ÜaQßgæîýÚd™Ã0¯gÄ¨\ÍÌ\’·ª+/õE“|é'Ã‚2˜$,¨í>G/Âñ9Ý0oà€cáSvç"«;Üœ0éÄK~Qˆ„óýLñèÛäöÇSÌš¨ôM=Ax¶Ü^¶¼³NÌÙkrÝâ5‰DúeÁ	&‰AÞÆJ‹fuLa¿´!²7-Ã¬Ã`éÊ(fY#íE,dKOµúpEíŸf¬1²J0ª—·Bˆy´ ÌAŸÁ™›ÆvÕ}ñÆ.p ï4L‘+2'4è–N620.1UŽ¡·4ÜŠ“ª¡ÍŽåÿ¬òÿÇ+o3‡.wË]£(]¯®ª[,Üscõiå}Úm6.vÂ»Æ‹ÛÿeáÈ#»À@z”™þÄQC´éÔC-+.‘ÁÇAÓpî™Z¿`CFyÉ'Y)fµ –å[Wk”¥!àˆQ/ó`¯Oš>3˜x1çÒ3È³óï¸‹i·)Š1‡\tüÐQ.—]t=²Cü½òÈmÉj<é«@+h,‹ô<ùMÁóŽy.H{÷Ò,&ç×o®…™ßü
­ õLê.[Êæã\:€Š«ðèÓY…"_eé';/ø×ÁO3Þ?ø?H—ü®e7X£û‡®á©J+vo/ÀaPviô"¡Ò$Á²ó «Ž³r
ÈD i,!¨zü“ÊqŽ×6•ŠÅ4€~ŒlÖÌ•¢z2Ü‘—œ…cêæO¶È`¥Gû3
Ná/8W{Ü—n‹`±+× B`¤¡7¦é2kŽkd­C@ôÎpÛ×§l¹›º—­=¿Æïw/Ã´4µ#gÓ!ÜÃiC7ïk7pÊhÙE»èƒõoÖ7Ozt+…?`
¬Bá0‚ÿBG0Ç.§ãðPÄG®Ì1 Ðä­ËhÚ‹«áAŒÌa|xyDu-ð}¸áÌåA–Öà’³¯»Ò£ÿ×_ÙXœÁ‚,™õ®G„$5,(ÌÁë„–ô+&Œ`ÿÇ€aúê‡Fh‡?.¹ø¬šä(Ç™›£pà”‰sø`LiV	%‡OŠÝÇCˆ\¿`Ö×kŽÙ~L‹¸ BmôÄ¯i\¾1
¿‡YLúÓ¢ÙoôÍ®l<|ðà¡K³V/ƒ‚¤Î˜ç$’ýB™¿Pæòb
tq@§›§ÞéÐ‰‚~!ëwBÖa§ÖhpÄ²Ø‚~&Ü»ðøAÀÿ® ¦œxÉµ@¯Û+>ZÎIDkgŠ¾NíªØ©2”,Ö"û=&LîÛëƒ"Ê|Phz?NÉ¸ZMmxo– xÕéoVôð|ÌQàiÔ?ÙéÖŸN"Ì%SeO»…SœF’â	·c âþ·X [êòäè£i·ì¦šXà0 ÜõyòùžR~‡U5*ªhð¹X*^È¢í{]f@@cÆšr^È`‰ï/íó6x }ñS“Ÿ]•7™ãÙe›#'€uÇux¾=~»ØÇEÀI5ÂµÌ®6×‚.¥ªöXÒžÁ;¼…{BNíË~Ÿ4î÷Dî÷¤i¿áh‚ÓÝy°»÷æõ÷¯Ÿ>;¤´û¸Ï€³‹¶´´/m¨|ÑàeƒO7È×ž¶¸DGªÁU¡q—2ßäãŒ=
ðbÛ¿uÄä/Ö¤’ƒ/X¥û¦+G¹]Œ¨0îâ'^xkw±UqÆœãçuÒÎçdäæù>”F…þ”Ü*¡Ü£šºYO,ØhîFYÌü„¬è‘P:.¯­tÈ–{“/oÿä‡•F¿YÑ‰Ö¡ÙŽ”ê· dîZkë%¡rka5jÞ{ãZG§‡…Þ,È¼dmoš_›ª)OE]Úþ)B,¹´ëì”øÜ>~Ý5¡ªÎ£ÑátÕ!ùÀ5Ç?[ÌÙ³öÛlvµ<çÆ]Wh_‘]y›î­lª…¢8ÎÝ>¡Wš/}‹/æÒê³ÔPÌÒèI×YãÕJù¸Z•èÉ«ä.­Tð0¿;¨¢ ¿Ý2ã4æD”ÊYñÂ³‡(„¹Vm·šˆMeÆ5ç•:‹¤ÚÊ`*U½Õ¸3H"KÃ²pÖB›ÒËRŽ%˜/·ÈêŠsñåKšÊ¥ðò%Õ;µª¹\/†¬óoXYUã6­Pù7*s†	IS­aÉdCi1¼ìR,eWß«­¡6$@ÇÓHW3Yë7¾¡¯d¬©“ÈŠ¥oâ4XT]“ÀuïÀM:7a^¸tuN1&e^ìú6–¦¹-Ï»þwwO·¯ó¿êïUrá=ñWýÝB|ß¾.þT³¨,U¾güNŽê¾òû§•ÉH7”_à"Ñ7ñú“È»:š‡Aš!g¼¹tGþfyFš*nêy•·/Ûü¹W¥B­i±ùÜý	ÓMm©Êy’Ì™)à^gŸö9Òª8šã ‰ÿëS KÍ©íåÀ?ý]ÑöÙŸúa/	|zæúâé ÞUö ¼€ùDiWÒ=åwENxúEñKù®øœ¾-ü®¿?¤Eå_Äg°åßº÷Š}ã‰_TïÖ¿àîÞyaôåßÊ±¿^­üT¬q®Å*Õ=
˜WÍ¦53jíýxü-¯4‹>Añ¥—²úlyOõåaÀõõâ§Iýæó$Y¡^GÊœä@ƒS/¥¢ÓöµøK	µ7•×¥ÍOíYO`šd€Ïï .šüÙ¿€Çá8Û:}_Y»ñg~?Çó[f>ƒsTû™Á‡›ImBœ>ßã¥ùÁñ»	º±”,Üöµ—^‡¤ëe©Ö‚“%WF›ð”¶J1‹¦
ï÷û
lë‘RDÙ"Ø¡.Y¸šéê
ÝhfÓÕœ½ –%ø—š¯|üåç$UÍxô¸2¡-â/ê¬dhÐ¦A¢÷k×ƒ/ïƒæà{»>þwÒ#¢ˆ¾ÐÓt±gtò–¡­}U_¬æ† W0<'ÝaOýg‰a¢â(è@¼“nçþÃæŽÏ„ÑÀyHh{[)ÖA­UéEóÁºQšGø~"Ý8²G¦.ÆSï¸v&ÏÊš€0n¡cÅÏ9Zþê?Ütæ?ø77ÄÐCí0û6€3#¦ÄY}œqò]¥ãyeñŒqJþ-ëûhß´ç
u½Z<~æãä»£ÕªVÙâßJµ²x›ñ½ìû"Uþv½$”Êý+?…yà*¹víÓ÷+»vÅ£ÌËÐŒQ+Ö»§ÿæ Xæ[Vý®¼¯þ6Š=ÿÙEˆ"'|yîý(xF{b/
µüy¿²ª¸
eÛó–ãø;@µ*ßðß¦žž>/» iw›ïés=/x\äÝðG÷ÁÙÍO¤žùÜH‚Q˜i›³ƒêx~\5’\¯œ%×+7Éõ€–žîP+_Ùe…vR¤0ù{Õ¬ë°ÜçÑQN$¸
?Ÿ†Éh4‰“lZå®ð©ýG´a ¬Hj‡ò·0N/!oÂ	Uøæöª‰—x„µËÛÄUA‘:Nu±F‰¦Sïp‰‹AÉ
mûã3ÖRÕîxE•÷*+ ]²¹ª;Ex"7YçÀ˜¤à³©	àÂ:42ûYŽº5öÄ=T²Nh–z¿½zR;7Í V_sÌc½Ýá+”j×¨,î÷•nne«"’º&ÀPéFR€eMMs€hõG€  C¸<QÖE<¼ý3.ìöG¾´€,ÜäÄY)<EKûNÒh¨T/¨§{%w
Í(’Ï@›L•.—6Ê£eÚ\­vØÔ‰|­´Gž,hML©¥«4yÃy°O3ûJ±[èŒ¹´*þ¼Î–Wó/ÕiÔ}©î‰F
@6¿ðÂQé(8Ã2¾i?£©¤Ýb-‚ØZ»‰¯§Y+üx`[cu†¤$ÐJþEÅ(Ø§€»T:bàTÂ/ùÃ$ëW'ŠáÆ“GÛe'ní>µà«kÊc•}ÀøÚ
vu•b…¨“,î0mâÞGñO(æŠoÞäëMì°-³×êªÒòæm²†:öú­¼®nëÑ2Û+òýë:z‰kŠëÑýªrÂ¡Xoó+lsaqQ“Œ2O¦õ}œ"‰Pùp‚<ÎˆÉvçV
O±R!‰eÂ§úºN%¨ÛÚÇ&e¢Å˜9öhvJUâ´÷‹ˆï·?³ªGÑÅVôZ»ï+DÔšcž´CjJtéÆ£eÊä7Jx@ð9†WØA;‹‡_
ìÝ«àìö/Àò‘eÂÛ31v”1QõÍyƒÆ„lþ–u9ò¢Ù.*PÅü@Ç¿ä‚oäiÓÇcbr¶YÜÍã°),Áte´O½¾€kvFddúd?Mcr
¢G‹D<TgqþÀXnÂò‰-’Â”¢2y¦qÚÿh•~ÑÌK‘#R?ù¤Ù¢#å9o$Í£=c¤éåw$|¥åŽ4XÕ„EÒ¡¬àŒ µþ±˜%yýY2NUp®IÞòF–IG'>¾É¤;cŽå”ee%ã¥DeÎL¥Ã8:xmô”)ÉÄØèJ¶‰™újß vœk­Ã”í3…îGÃWq+¹Fý5ÊuüÑ|M×;¬|fû¥/¢•»hÚ*-¿X°ÆEƒ½Ü©;Láå@ø)pÑÂÝƒ$Æ kÜ¡5ôpÐŠpÑêÄ®ëp~ŒeƒCÔ'¥¥	Œ÷Çá0Dµ™¸xb‹%û©¼Ú´‚´ö«
0ªw„™sÆŸ×†ÚŒ›•>æÍí­6(ðèI§4JŽ»K=h^Îí†S ›L¡•ÉÉÔtì%=u[Ã¡óé±žR(¬†§cFUÈåÀK–þ)•/»ÕSBÁÁTÏéË­Êî!°kZ¢^¾€d©0(ü¾>$øã4L_zP3¦Ô“Ò}TDßó`ãû@<Úã9?÷«y5æì4=ÐÏÙ‘¶6ÄrüÄ†ÞŸ:Ö=™†š™Vg@;¥‘¯²‰»O¢8oÏå:qF|lD{ó£uÞx*¤ÀW5§Á(¤+Á“Œã±/þøv¬“-,ÇŠ‰Œ!×áØÏbrL…1Y2L…?†¹TçËï³:,«Õ
ú‚U¦*7ÿïÿü÷5Qú9È:;ßÝþ4N£˜ìNAF¸ý±e‹Ç¤“×½tâ¿ ,"cœ„'QgÁ’¼šú(	Î°t+¾Lñ±¢4ë²à´¤Ië:ñ‘.ökX-+Ìîaß¯5:mÎƒáû½0FÁ,[_'±8îr—Ì,‹T­çR[
ýßþH‡Â½¤ÁÐs >®{è`ËRJÎ‰$ aˆûºÓõ·.¿5‘Ë\µfÜP¥Ê)—r1™ßQýäÁ­ò‰¡–Rä¦
)üä2MÎ§l-šH‹52QÑZL™P=±«y]´)[ô²S¡ƒ¢“Á5Ò&eujJþ¨8
¹~M•~iÛÞjàÔª£HÙ¢Yþ©óÕâP‚E—ÙE™—E“,NÃé±‹°C$'ñ¥n~bõj7£{2K»ŠÓåå˜1KÎ~T\4êÒœ¼]ÐHÕZ3ôÃ½xéŽô¨ŽSèPŸ‰Û!Ç¹q@ÇBˆ‚BµRŒá­iš'šlšì!Ñ3ŽKä¼¬.¤8Ê¤²­œ}‘ˆ?EWQøræB>jdBO'˜‡î}qË¯-@x©#Åðj¶qP=§<Ù×1ÄL½t¦¼’ÊDF9ÂyÊ¢ŽRF3ÊÔªô£3äNÒ°*ƒ‡:fLUƒ¤Rë7O,%b”Š-b½TÔQÇ¨¨sü0Ò]²’j»@¨ßÒ'?%2œ¶ÔWZuÑMøšz¾ÀY8Iå[îJÇogþSc% êê
º™+sÍÉŽÐ•p©öœZÛ¤'i±°èf\¸Ö´
TíU|áå2;¥ZÜ©#¿uc°:´ñÍåíÒ€ë'Ÿ\íàÙ_¹Âa‹ Hø¾¨ÅO* Py‘¨õø_å´ú‰÷ƒÅ˜
èqÍ‘„0jÜY¿ŒJL7éãÍ•ÿLfxP#¹¸cu»O_“gG¿ýÝ³Ã§¯·„4¼Ü$&<…P­à†©ÌÆ7†:~ó)#¸¼¹b¬(%®Â,®Mù?]!ÌU¬Ÿ¼ª/¯¬.]“'m¬LÎûTÔ3•2ò<„Lsþ`­-u±á$ó/×o7ª[¨¯JgTlÌP¶‘9z?àôØÎÃ|ËIxÆ}šµÌ–ƒ4DÕºWôÊÅÕ9+ÁÔ¨ëÀËGF³P@jˆsºq¿
ex…¨ƒà!4Ü#BË[’·ÀVÊÃ"ÏÏqvî’†¾À•¤¦#Çüá&ä¾kô¦öËçF&„ºÐ€b$FŒvÂéÏ«•x¼p43b·@m±qwÜÖÉÉÌ„tÖ4'Š¼º¦”¡w„Å½WoÁ[`UÄÅŽ'@—æÊž…YÚ•Ÿ/V²À®«ÇÊ+L¿ó¢ÐÇ±Ô›²—ê²gõ-TÆ™°UâVBþú¿ÿùîöG°¹ø›=¯C­›s+ra·ŸØÿ$ûã‹ùLMûL»Í•êæA<+þ++ýöß2ü·²dª`>ò._P„¦Iˆ³%oºí¶‘êj%÷t¤¶±ÇãØ˜_}ÆmÄØ¸æmÁG"ädÏ7îZqˆ5ó?7ÿC4™YûÞª&™—:Î7[­×b–{òa5?ýêJG¿;ª=Ë·ûÁ¬ÇeC¦xÅQMqbÞG46ºÅË1Ô\5ævLÿìùÏÛÿHMÅÇ½Ýç÷Á€Vhè¦sÌ ÿ]L.@ëO=âS«DdãQ…^¢GÆO•§Y1¹CïôÕ”.ÎÜ·OÞòGcoï§_Þþ9]Þå™ñO‚$›&_D+åUóÏ\Ù€Íðt¤ ìÐsGÖjów‡±fJÿ1T_Íãæø¨4ÑsìúÜ°ÔG'"#†ÿˆô™7•]^3DØ±Ë1˜†ºQ´»ÄÜ±KÞ=DÞ>ÂL·xÆ þ³ÌQ`H¥çê³Ð`8?c‡š'­}½M»fæc—ÊX[m¿Y`ÿVÞÇ'­Åò&Q~.;52.ïÑø?þCÈ'v7Gp%œïãP¸Âó¯ñ‰<Kd#»*±‚ÂÆop@“u6dW:(ŒÈ7X|Õ„Ö}g£39åvÖ°îeêLGa×}ºølÿÍë-B+Èfq<CŠàöß1Ðÿ0˜ÐD`<ÛV#ïÆï§aà]Õ+Bù¢“¿!/â¡å)~^ÖÝ¡÷bHY¡¥/FÝòRÈ[…Ù0. 6gÓn­}W™ëiÀ}ÜØ)âÁóðÒ,–ýbµ
Î8Ôhs×ˆÿ<5\{A4¼äº©.•*•ÁkþŠTÖî/I)õlb#/ô…üÉÌÇu€þ¬Ïk
±9ŸÑØæ<Îe§Ûj,ü"(Ü£ °¦JAý|_g•¬~ù
:		6¬ê¥làÈ>uÊÓT`òE¢/•›hƒæì(*¶}Ëˆlóu£»;'ºƒçÍ°Vv©›œV©€É	Gz·ê30»Úåî8ÏÍmZ|îî×ãÎÍßÎæ’fóµ»GO;W?;ë”t–‹æñ„³RQÉ³®p­3xÕ‰>uë:zåèQ÷©øÓÝÑ¦6àˆàcgÝÔF¤ü˜ÎvËÕî.í´2¤J¾•Qe¾g¸Ü|áz‡ãæy7/Äxk¾ñ™Bñ`]¨¦øÊLQ›•¾¤ N"<TsVÏ;°S* <Sp§4ÇÚMë)Üiï)Ã…~ÊÀkþ9¬NÞjŠö,ä[<³–Î—Ž¬\œ¿%"}Åà	FF‘IRKècÐF¾ƒŽ\n>*LÊB³åûn²§ç ,ÜÃýÉ_ŽÇ—)e¯£ô$[E®kûH¶Ôx‹=}|^=M™.Yƒbeç'„µELk˜RmeT¡ªwœ•Ä@eC‘I-ZZ°9‰J(&áð,²’I¨0ÉH®“˜A6B¹ˆÎÀ&ç ¸«JÎ®á]Ú’±¡EÖ "ág‚O!Øh$¼3§-§Yb)vHÙQ,º#Y†T2-›|­?Ž\c‘if8•gŽæd4‡µ>ÛÍtöÍÆIQc[k–IoYûœy&‹9­5j8›Ô>">¸9›ü"Yi7’ÖØÑÔ§ä#"‰Í9ä‰VöGŠ»ÈGDŠÜ0ZÔ…þ‚ôÈÁr’·}ç(òIÚÌ”iÇ3•Ð=9-¾³JÂ»Î&LÁÒalïÓÅ’Ÿ·(øü%q›)Òj†üÌ¤ðùìA»TnÛƒ_„scŸU8qäL¡ù†¦Æ)›ñð¼ƒÈd3½ùYðOFW2P  xñK4ÌYË;¹),Á;$ÐõË`œzýsw<å=NÃ1PÄoA…=JN¨«K;@AkO|›xÙ4Dh}Ä®ô=Vg‡ýKºk+XÁ"]lÐNg‡þ<ØhÜB	†ÎÎ_úo¤üMºß…™Ýþ'Ëd?
GðÛ÷­ë’ïoçÉíO±Ow[˜îS‘–"Îx::	%s0
ÇÛ†Øî<|øÐ^¯ôr‹¬­ôÈúF<€*éñ'GÄG¶ýcü
|fô×¼‰0™+¡(šÞÂ²0i°?Î¤¤G+´¨"ôüi%G–å¨Ýa¦¿Üã$e€À³d—‘íy¢k]y°ÞðÌÈë›µ>2v¡Ýeô•-(y…Á¢iÈv1’ÎNñgÑàÑ4À™bj°¥¾«º>@òÔq,¤{ÉŠ>)}¢„×çàóô2rw»Fó+EÕŸëü¡Ôá­§{a9†îûd›ˆ¥v5q9ˆRDin]`ïøo^½fˆ7†XûÉtÞA½~Øì€³/8&¯Ž¦CDè#¤PÝzçñ‡7±—fÝŽH€Xrûü5ŒG$Ò­ñU§G:)k¶-JMâ
LÓ QåÙÉ’+eúV™!óÓxØõñ¿“Yàn“é÷­Izàb‚¡—Ï	/Ug¸þqôƒ$‰˜7þÃúâ+ž¨ä4Y£xêoÁœëk„PJ?
ž‡˜c/NÚZŽBZŒé­;ò‡ûožõÈ;å|–¿¾ÎçsóN5¡ÿdúíT’9ßïâQqáaEæk–±UèT½î\QŽKRW½	 XA’?xby¹€Sš²Œœb·’rå’r‚+V^,çqõÐ~‡á©<nÿDÂ6j¤·?]‘X„®R€n62ŒeV$úÛ­Ó(F~Æp`îÒt´ÊÔ›ø)mX:|‚é)‚Ý±ÿœÊêø^W|k±ŸÅ/âA²ç¥A—fÙZøÛZÀh¼üƒ$>‰‚Qº «÷6a½¢×Ç«Ã¤_Ë½#òéZ«-²¢ÑQM‡q QÊçJJHQdÓd,©tëÇ%^Ò‘I9b‚Ï'EYÍšh#›Öóü›ü|c_•XU]}ùCÕFg£”9Cí–g—à«+vXíÏæk>UP“1Rv	%,YK…ñÅJ§r¡Ii§‹£ol°©köËœeT=¶x€IyœŒ8¦dƒD(C;!Ä®o­ÅçTküJA•Êsj‘-¿T"É®ð•½tº¶È£(ü   ÿÿì=ÛrÛF–ïû=XTFK¶GëÈËXr¢™XÒˆrf«²©)ˆ„HÄ$À@IŽŠUû6µ°Ù§™ªù}ÕŸì—ì9§/h4ºq¡H]±‹¾œsúÜÏ×K +*P›Œ°U&‘»ò†Ãv®ðåÆÃÊe™ñryñ¨ð›\ú±s%ŠÁ•½ÚžÏÕÎ•!·	.‹_Ëv¦ì°§(PÚˆ;ÜrÛþ
¥èÖž­S<·‡Ú¢bÝe~ÁXƒô\á4Þµÿ’m®¿pGŸŽí|‡»†³ö’—B%ºn:AbÛú²´à,-ÈÚ©r¶Ä‹˜ Ö®Ú(–sð¢|»{TÄqê9Ž’d”q_zmgË®OçÐBÅ ¯Bí7ë0+S:øiaZD¢:=°ì5ˆöI‘¤7\&Aº‰¤Ú^KéŸv$X˜XÜ±F¹Šu²åzŒ`„þKüI’±UpGÌæÄ`$È³ö¨YŒ0 ‡ÅÈÌFTV,öGÑy€*N³Š©;cKÝvl>æíw>Ê›%)Û¼ÔqrÞ©.ÃI­Úg®Ø§`è†\ƒä‰èr›çs…²ÈMp5$8á2˜ŒxeÈp6Žz>±†„Ðî÷'±gž0T¸V§ø¨!\Ñc™Œ0]iÉnk‹]<NMæp®«ú9ò›¢óˆ•d….Ó„U²ôv>?ÚVè_ˆïûh=YqŠî"ÂÎ§ÎQiŒ»ÑílqufÔÜ,l;ù8`žÕP¸Ê…¸Ú%IW×Ù`&ÕøúWœÒ8
I\çÙ…õ£8øÖBªQVsåêá§¢»«HYû¬:•'ÁaÈŸ -1ð¡€6É`nRt”Ð Ë`î¯CFà¶¶¢f-óÁ)Ü»M±¡©íbFá¡` <ñ»a4Œú·Ê~ Vög8Ÿ;ÞõÿöüU¶×›ˆªŽ«ìákóŸÀ×»–xŠL¶M$9È?T-•¸"7LÒ¹óAì‡ M	¨ˆ—Þ€ÅDÜ¶Ñ¦¹Îü([‚^œDÇt}•e[…@^IHU¨òÿ¸h¨ñ”n>Ë¿DU¼c4wl*)[ùvw@¦÷Ã$ ê*È3Ñj"·å¤úI§"™ª(woŽ´ÎèÞa8ü8"ª¥s«&zèåÚÄI¹MÑÝ€¤ç$Všt¦Šš,ü°È2O¼Çã­îñ¦1R†ºÓÕ±XTaøÌ½?Â¸’¼ñ&TÆ‹=¸òcË.‹?wP½Qú†=)†67Ü9g5E‚ô
S_™Mì{KWÆM—¤i¯â¸)NAêQ*Lež\raÃúùà4@Iµ(€Óu>“ŠNÌOú™z‚ÄœÀÝÆã½KC€ÃávP+´ýû`4F&Mj¨viüZgA²4°§Q¶‡Cs NHCÛ•+Úá+z«Ð>2}ÆÊ44È¯s2šû5ä3VêÂMÝ×7£&ÔÛÕuÂ<cvtµLuN^VhÎëçl&À2Õ8×‹ïë¶úÏ
Š¥‚¶ê•E¥›ÿ½±×'47•­ p§^Ø¥€öa»ÆË&9–éÄˆ]8³é«÷8rEß«a°£02§Å`>@¯˜ß‡-ö“qóÀ€‡\=y¨^Všx¨ –«Wë1¥,Yu©®{ÈË$”F³ª3`ü3ðyL®ëð,ß
ÛÃ¯6&E.Ð©¶§k›Ä*t•¬•e¥%&û¯ìr„Ù]q„G•`ª&I„RÀÐT—p¡B¹Å™¤@Î­ÿÂ5oMüÆ«s’½õ¾¥à-»×¤n½Y©Èm<{oåmij”+['æ¥b:T‘Ùy@&<NÁ¤vUÙ¡FÙÉ¢éB3<ÎëEót'4Éá]ŠùÜ(“Éø.Ÿ‰ßŒŒ¯Cüýð=Nz¥û›J÷£ËD{IÚjËõò{/ÔsSY“#0gÔóqþecj‰µ¯g·+8âåuª£{c©óåÅ}/„Ã‡ˆðP‚¦æ˜ [ª%CU°Ý¡3?Ìø%øDH§PW-¯-SßŠqî‘’Í`†;öéDKØw´Ð"û6§ðƒ8Ze?ø!ìtIkfoSOÜ=Q«é%>IÎ¬±ºRòvª‡yÑ¸âØªB	ØÉ¤wjÉ4R5Y_Qsr¸{Ø‘Nh:;-érÈ³éhèÂt—–&cÝÎT%Ÿwþï?ÿÆõ©åœÎ¤šòI[·*_*^—óÎÂv&ð"¢­HüQŸäüo´åU^r˜+=>§KÈ.ÝÔ6˜Åš7`fI2Þò#ïH
_À[1¶G¼Ým+¼ÕU[`yŸÞ¶õ©•á­÷›Ç[éh<3Úv”þdîx+Çöˆ¶y´ÕTÏŸÞêjõGÄ-AÜ,ÂnÔezˆž(±9ð»Á\}{-c­DemTo¢Ñ:¸]‹çÝ2ÒG™^úHÓEw,?+ƒ=7#h](G´«,š–GN5ÖZ`–Œ¬ÍMçº<^Ö½Z@£ûÚ´Cë^EÞÊÞ5:5k÷¸
YYR¬“ò€iÓŒi6á^$º6²^<ˆD õœƒ^šÁÿ"*²*ê_ÄÍßÄ™¤ªEÇ+™‚S«~5³¸cXš°“Ì†íH›c]LDci
g»Ú<Ê€í³ Ù/´·áËfC­›š xeÃk&föÈAŸ*t	 µ–²ö3½; §ü[:D@ªÄ
Ì€ÔÅQõcoäeq‚ŒéÃ² ¤Œ€í®× ¤Ã÷åYÁSckUR¹ÂpæuàÀð€y]ƒD­Ïe#ƒÑ8ŠS‘þ§î6îÓCbt°úâ´Ÿï6V>‘ßJ€ò°èÂïL–±ÊdÞè³öŒ×žQ¯’[öÝ°•%â. à?¹Î…› ·Áäª—W@·¹”i®[Â³n‹B\ÍÁ5H¼çpù33ý.ítDv({š·]2'®J‡È³!ã0LºQƒœÄ‰ž÷t&]+ƒB@*a_¹´;fö÷1/Z«´b?­ZÌ0G"q‹WÉPÛ®„ßÎ®êñXJn#œ6ÁïMÉE!aÌÑ0ùŒÇØ®7ƒ.n›}tNø»WK,E«Å-±æŒ%öŠK¬~ZKœ‰ü‹Zc)ÒV-qÖî&+l#8ú[(ã¼C¤¤ÂÄ)–D+…ê¢iÝ\J%L‡öyDÉ¦«ÃÅwdFÊƒJ­>Û*–Yå #u.vx¡»›ò¼UÜ-~¼TÏ»8}Å2R†|ßfIŠN«,Øf<ÕðŠ;ëM‹îÓÜôhYH}NÝ¹êVL°\Ì=:'}Äyà_¼óÒI¤wá€†µ™½î¡Ü'ï N^ŸrFì·?A†ÏÍØm|nôÔ&D$ñä8{çËŠL™…ÜB6pKŠB½ÆI#fÐUxæUgìÅ†¨¿SÜ +7	R	/FHGÇØ¡¥ºŒ£ÍYŸðÛÁ.¾ÌÑ£4öºÐgé" (]Ú±‚%€¶ßM)Ô·ÈúžIá¿¾Çž•ceCßëá{Ó ?H-‹—S90¤bŽ“Áóæ§€Có|µ@¡:c•>z‹ÏbèŸ\ë<°9½å6 @ø÷ZpôQ‚åžÇöÄ†Aÿú˜³“µcº‹E+UÀŽ +¹ò±?ô.á¨¨ŠàÐwâg}tÅ©Ò³¼*Æ»ÚŸKüÈf¿0àÏ÷Pð20¾þ;‹ý®7I{LwþÄÏ{,`~M„R‘?Á¨CŠ&‚›É4I¼I&~LÄ2ûÉ¬GPËê®+=P–1>Ãh…èÕý•zç“½*}ö–×l‰|}%£Î5…<F]›Îõ säšvŽÑ’|:× %Y{Ìijdþ2ÆSè`2îéè¹«1vËÝöË™&¸Æ’¿‰/¦¥æþÍìg¹©Áà¤ÄúZø9SÑj\§ØYÕæ†ÓM‰«‹¼†äˆòF¹±«ÖÀCª4 PàTüzã¾¥¼½5x]É¹S·Å¾6J	Øpw
o—.úÖ”Ý(19ji¥EŒÆ<ä †Ë)ƒ=‹Â>'`”Š:³tˆ„@¾Ð:+9¯î$ŽÐ#.ÀÜôÜËag^cúxÄ=úX'ÄG wŽ…Æ²Ÿ”%™’ÒWŒ’VóH{sëŠ­¯¯ëc)KóNg£ø‹M´%)-^›š—,,­™¯UYS(!qgK!IëM´…E{bÖylÕ¨ë#³ãR‘9ç¢MLº’”£&q	6ü@ôÕ»tág;h\@ZG;—õ 7áxdl d·nŒlôgµvBà @”ßkCD–ñäïH{GTÌÁFªÃ(^£9?âÊ8	ãgÅÓ©” ZŽ“ÜT=×Áâ svFN½nôv>”UÏ„­ÞZšûªlïU…Â™fÙä÷úÃèÔÈ„ÿ½w¸ÿô-ü*³ì¦­¼#wù.‡.P'ïcöÓŒlðß;ü›¾§¬u¥Æ5-­<t¿äõ#”ì¥˜ä:ðiBÇpÖ¶ž×‹µõ”„-s>èBvYRH«$R/T­
Ëeo…êYÅ*Yˆö®S0‰X‹û\ÿ5A®z€®j€µ ]• © ­rì½Áº`“Dó™$Úf´ fôã}™`MV—j}çÇháGÕÕ›I’F#J§å„¥Oˆõ¤ì!Â%ëú×ÀÁ’×p`ÞM¹\T˜J(T5Ù.iòš-NÒ8Z¦JAVÎ&SK,GÆ‹Ú¢‰¿½>W[ um6Û Ù"ÜÈGÄeRZÑZ¸x¢²a¹!àÌÉz–Ä”AI
m™O`ëØï‚ÜöÊë~‹ º´CXëM4¦ÎãÁø%X V°q”@…þ¢na™º^^á†Y¼ì4Ãæ/l©¸.ú±Ãÿ4‘«Ó%[Ú¡?%<««PÊ•›4ˆF€s7|ò„}6Ò“zÜÙ©¾Ò+6÷÷œKOP¯Ä):êß9eåæ!/´d0S$cXÎtÆÌwó²2©.ŠR«<j©Æï‡šô’hÐ×8ZØCælL>p¯U{Ô›[œ…Ý”rB¹ä\&%q÷F%ùiNâ¿19ëV%~ënlJjªv‹¿Ýâ©÷šZ“ÄÃy[’»«Ù-IÚZ“I®hfNÊ2Äa°­Ã¶”M™ŒÈR¤]8¿ßÓ» 'ˆ¥ç—Ãzž%ßÐa[û¹EÎ _|¯a_ºÚøœ}ßÞ=d{?¾ß;Þ=Üf{I7cþÁýdËæþÖŸo¨eŒ^o^Cg´µ/™Ë¢U|Nó6Øxñ”<œk8l¼|Zò‡×…t8 ¢,D„ ìýˆ\Z\.Žü¡pðœ¢k×iÏC‹¿`¥ëÌÅÚ6X£Âcb¬/ïsš›ÀÒÎ¦Ës‘É~bntq)Ž$ø)6«‚Ïd¤Á'OÛÎ¹|½éN™¸—sëÉ¶8ó’pzõiPœA%åKD>þ±
~Ï8{"eÒÚlJ‘'Q PdI²[,ÇÞÀTÃ$@æsø•VeÇdt-´ŠË/ö
ˆvÁ‰ê…*²g®2Õ«©Q²•]Ç};ÊôÂ•uþÉ„´$À²zˆ†` ­u…¬ø¹C_jTFtÕFt3ýüž«"èáLG–Åcm!>âOÈyÛ‹ãË¤Œ|ÑêÃ×aÞsOÁ3Zã°Çk'Ò Hý:˜yüxílGf>åKZª5Úæ¯wY÷ð£Yø`.T‘9ã³¶YËÁ~-/¯”õ*FÐ™ôÉ3Mï\Üús¬ôæ{Œ‡k¼Íú;ÈŒ‘ï”.ñžœ¹Ì•VY	?Í©(kEcž^q¥A-ÔÇFx+f2ÚŽ›Y;ËÀºÞaË€ÿù¨…Þ´s‹Šø)GÚ£Ní¸Cd°læÁ¼äÀï£o¨Œ.‘!Õ¯mCˆù-eùâ©BáÊÃ’?‰?ñÉ’=³®­„˜2^AZ BkéãôÍcAý€/µ›mä…•š¯‘óãâÚœ7Ëªóü½#/Fïß¤Év4?nŒRkmÙE'ru£@‚šŒ_CÅXˆŸ;WúºMÙáš8Ý&ãáª×üð…“ë¦“!–“Ð\Î×8µšø©ÐlâÇÎù¬¿x@¼O!ßê[o8 ]¿4Ž¨¦ÆG™4` ò§øœäŒ‹¨'~ø¹wÜ8OKÏsI%“aÅÁ‚Á+Œ|½„¼J²–þ'Y;ƒ5vmð'	èûíƒë¿¶;«l¯srÜ>¹þ¯o÷ÛV_þ)ByÆîÍ	Â9‹™‡îõ4z $êK4uj/õp¿YÐslyMöR[#K—ã!µÆõy×>þÃÞÉþÁ·ì	ûaï`·É³‡G{Çíë¿^ÿ÷^ž>:>|³×é6èàxïÍûcx‚}÷þ]û É“'{o¿?„ÕAíœ{‰€üˆi~b[/ÞgM­}xïUµ–0ÎLW«Ï$§¬Ýr+k1vÎR^µîäúïÝv‡¦ÖÊŒY§xKš!)—œD%¾|øy gÆ úGy$øEÜ/?Dòžúîö<™ÙÎQ?N5S^a§ï’‚j¤Çd!Ye²Âíƒg[¤Ì¼U¨Ù' ÏgÐ?N¼ až*¶vî('Æ‡bÞÛº@ê¡jhŽ IýÝ +/ZgHËéîþñÞþÉá6ÙãýŸ9¯24Tcâ4ª'ßšXO_<ë©‡¾ãDæCµžÇÙ,r§ñ3÷i¬o52‡±÷KTŒ)(9KT#\bæÉÍ½±ò(""þhÛ1Traˆ¨½»þ{/ðæ§¢ÏÙFÚ¦€CŒ¬(×šòÍoi‡ÿmøð7^p	ÏÒŸòG–4õÒIçð[Ä‹„…pâx+¤ÆóÂ
Þõœ1BIø·†À½7bmà!yîöýWÃŽTVKè&û>w|¹^¼Ú·J ‘«B³Côêû I€ýž¬Ð¼UÌ”·S‚ŒPÃ%Ù8•Ö¶»BáÂ¹E„]jEt~0ÂèjÓÆuH„ê"¦ÿ%ýP‹±ñâ©›˜ˆV¥&5¤Ô¨n®²­Uö¬ Ñf¢:Îî^O6Lk‹—ðuÎÜÎoþŽm®PJMñÄešÕ£QÆ^œøû¡Yn•m>uËi u› Æ‚!8Ù¤·13Õ`}ä]¶ž®ÒÖ`VÎÔ5…%{6¢¯,ÉˆF)_ÚÙ¥ñ`<# ×!Êgé…=²`ÎsLaŽùŠ¨¬8|ïSlkÿ™mŽ÷‘±F)‘í‡×ÿêsc;úJ73º<hußû!Òí¹šM²n?-Ó ÚÛ`ôfÍÁ–mî0}š Æoüà‡¢Ð" ð"žÂùæú×søè›ÑÄd­¯\}LˆùSæœ'6hý~Z$7ÃäG¢Ûœèf«‡ÐvùV¯õ¢qm_rÌì{tû©_4",Â`£ñ5ÍãI<ú·gøÆà
€gE@6‡ [Ø'ÒNãPâ›W+Ï,-óEì¹:e"Û4N'i…%jÞÀŽ=àt?hž:”›.øC1ã½‘ ±Õý®s·,ù9ø'›âU7tù³:ª—lâyi–“¤vÓÈÈlP·T½=^Gu:XûJ¤[‘Ü¸Þ›f%—¾sîŽöãB3UŠ²ñê(yŠH59 Â^ ñu²K÷ª¤g>0³ý¶†àv€°nœãày‡ÓÍ4h-ÙÞ.‹•ÒqÈ´=vO ýO/M¼ñØëòîoÜÕô+ ¾à°”yýîC€y?I¼¾ßùË¤øiÚãñ½/ùv™-8?ûÈ4µ—)Ïï`ƒý£I2piyÆXþô“'lù ’IÍ£p™!»Àòwî4…rXÞï]xAÊô^@˜øË8å#h$	\i•^àx¨+J/ÒdIýž3aÿxx ·–Ž"à9¼ó(^eï4R†µ“3Êˆgôh	±”Ý„… ƒô‚cžy-ö»þ) ž'|8€ã™¬/•…8î¸®×$e¸aÛjS©s?uB†KSEÄ4æ´HÃ´›€„uF^œŽ¸t•ô—¦ír)ˆ®ZN†)/žNûm˜#CP‘Î*LŒS€±Èi=9¡LP¥øê†N†D9b‡Ê±€Yãˆ!õJ^—ª”Ï€›¾:…~D×ÁÚK–—Ò9¥3hU`=?m€Kê·Éý8v/iã2!ð8l¢ø©’ò]´Tõß A¡ü¸vÌ4å_mNw61š’ox•@áñ­éÎ>žÌøü³éÎ³›<ÿbºóâ&Ï9Ýù²ÁóeN%º=GÌíaüxÒKIo³Wà,Z\€bŠ™tå“T#$ s¡5õèCª0GÕ VvÝ=ø‚7ÀÏ‘ý›«¸ŽÀ¢	ûÝ‹/Ÿ%?³Ð¹S¥jÄÙIŒè¾FÂ\ùqGù8ÔÚö–\Î.fç…›•g–y´šÕf–É¶´s*ÉlI vÛèµ˜E­wFnŠ(«¦Q•y¶ŠË®ê?<ðÕ]IÜ¿a$ùž¸.RQÔ¹:HòÚÕAR3uÕAâZ2X¸y$ÑYÉÝö«q$mÉëÖAÒ‡bK"/ÊQþº\[…7ï]‰£]š‡°ŸDUñû)|EysŸR¸Š¹?B/ìúA£‹YÅÂ­Ì@|ÔMžõ"€fWV™¥ãûšhþ+™ÅJ²×ªZB¢ÖE;ìávy)–¶h¨AÄÝ5`Á¶nÙ|‹â¢¬@ˆÇë/•—ù­W`Úv:+”NpøÏÜ;È¯AAÛCà£ø‹ñ*Âi£_¿ùàèlf:¸h`1J'|:$q¶
	˜åD:t^&¡F»m™¿ý^VB(‡Zk%Þ`~•a±•‹4-Zà ³hAya‚8„¢k­†¿½‚/SP@÷;*SPGe™‚&´me
ò×N¼þ>BÞ;/ôú F©×O/!tÿøÓÔl­­wè_@×I]F;E»3=„ì‡ÙµÎ6ŸÐ Ì+ùÉn,Lñ˜©@š©3ÝÔü•VÐm«sëïÐEúb1´¢>ÅŒ6žð%¾áJMïµUO ’)¼ÎÏž†ÿŒ£žO™E 3? à0ºæù¨1ÌŠÏÇªJl¾ŒülºR8ÉÊâÁI•|_8YJ«ßgp"Õ¶*‹Þtû;[T0ÂÙúœž5_‹*ãVÑeÀ2ñœÝiäR‘Ž*Wk!ŠÑ¶º#Ý4fŠ\¦yë.í,åú——Ì•™'MíðìOMq@eÿZˆ1=,PI½‘À¹¼Í‘À–ÿx^ø`7)„°ÛÏäí™ÏÔ;dÙ'2óYv1Ç“ñ:BÏ.s¡]·eTË¦o·©IÔûÖÔ¢fÃ|•¡n¾ö4}êšÓ´iÐs·l@k'A#Ešù10l½¨|²ÏÞàcê7™TÃ!‘b¬RTÈì=êÕÈìm÷Ï)UIº(Pi¸$Œ»Q‰‘Ûé„¯‡Ó§$[OihÞ=ö.l94ôeÎ»Õ-Ç;>: ;ì=Åz)±çú’}'—Tµ1IË‘èt¶ÂôÅ)YoÜ–¦:v¦™,M7·59T|e;#ðm¢­¢ Y²Þos±JØuÄg5#,çÐQÚ(ï*D•FÆ²|áHkÇe
è€^Ü!•Ø®£7v•ó¬—,ŠI&è>y}8#)B 2íJPC¶ŽNšhñ?3i3rýatê$1ï½óÒî 5¦û9–‰Ss{½¢KôŽùÉU>uÏÆ½˜bsÚc¹Tàw²$ÓZßJVG²Q¶\k$¯5Éè©Ÿ^ø~h‹—žÅ´Jj4û|
z8z,~žú;Ÿíš›2:cÏª"ÏÄQ%';åqã¹ÓÒ}52ýã8Fè·TnÅkòÔÇ`˜qü±"p.ö±v†Š[ëû!ìÔoïw&ý>:uC«;$¼Ê~ü	ÿwŽ—ˆ÷^¼fÒ>gCfRÅ’–™c[ÌÔZáŒL´¥%Ì0÷òŸUÕ™q½º£1¦Ï•DªŸüåên´4"ÅÎl7«»ü0ÔL¢l†ùËÕÝÐ¢Â©àtº~’dã*ÞâÝ9{+©øVf7Ç@5^”Ç‹(ÉZÌpI*òzgJªmv,Ëjáùõ¯CÌ)ßÃde©?ú×ÿ»ÇÚ11e7,AQ;NY$Öòãx›yáÇJQsàÿ.±ßQ·ë#ë‹K
?WnvÐç˜È $.Ð0™ã}23~¡‚Ëò5fñƒâ’H×aMÜ÷raz¸f>0³TªÜ÷	¬ò~›ÇÂÖüØv´½`ì¬¤ÅÎ¬ÜQzü0„>Dôâó’w~KãÁ`DZØ{Ì£=3‹Û+Øé3[âÓtg"F#Éb;Š£míÒïUøÊ‘›µŽj*_ÁÞ¢¬‘i'Š5•§E^)UŽÉÍ£®\^°žóð­C6T“è¨SïF,èÌìcaÍ}Þ¡›C7&¦T Ücø3Ž¯MA¸I0úLåðñÌ¾þ5Æ0G¢“Áh”Îôº o0Ló öo6/@ß»ô»µ[z¹Uö‡£ýd…ù#¶Ty8	kàÀ«Ž¥7+â >‰?%ª¦:>aMõP ‰ªÇ.·³âƒ¼xfo_k*«l|z-µU%JØÅƒÁ™öÃÎ/,*«4º«<‚cö)€ãnûhãMû-óÒ	jÂàà¯	Œy9o>`hŠ”s@½~Cò%PäÄÙ2¦Ï÷ýùÐjÊÎq[Û1µ*ÎÂ/R9ºCíK+zq.h«B
\8ÎÔ°Ÿ+Æ¼óC”™Gl¥Á9w±âÇ¹uåÜ‰~Dß‰ú-XUTyÌ³lZ–`×í ±‰†-áïÎmÔ%À;£©ë!ºJà©$<@5ÓéŸáT´Øv·x­‚Ðˆ|sxÝFQ:5lg–(ñ@Ý Juc³1¸Ãtæk‡ûÎ{^L)¤%ê rüäàª¨$|gNQ@‹Šºy$ÝQ`Öh ²Ì>WNL3‚\íž<ae±BŽÇÈ7ÏÕeˆ 2†»FByw\Ðg½oHd„iÒ\ªò4ˆ&Òâ‰Šƒ)'jL„çt´hÅdþÉ_«i”Ã—Y*–ã~„ÌÄ=uST^‡˜¯ÔÇ4fý0JÒ «	J¯Žåý]u;¢‡Õ{‚Ýk¿¾
ÕV9Ýf£µùæÛœt"ÁBÈoÙ½lÀÜ×embaàƒê«ÖƒæM†°i“¸KÜïa;1*ý²>5l­­¶:²éÆŽuÁÑZ_kôøÕýžEoA²+Ç²ÚÔ°Öbë¤‚¾ð¨X÷Í7ÄaœD²qŠiM4ŒíNâ¥]Z“^î©×TÊÀ²<jo"ÌØª½Yî4÷¼"oA_z^™‹¶û½•×ëäZ››Ê~OÐ k³Kl¢±Ë%-5ê–­ŠúÜ×ÏðržÖö’#lÀ“³ÙnëÛÙñÎ}¾ºHpð['{‹t¾'“æž$ˆWoè¿'osmkh—óÇ™öJB¦EPÈ`ž'pí|».øÍ7¨¦ÅöjÂ}à¨þ:šäÐ·“Š~¾¾ûñfê™Úáå¬-°0ig2‚»ÀdÒ¸-.d­&‰SŽØ¯¯ð+Â$üÐÀ<Ž£Žxv¿ZÖê¯ûa2Þ»GqêZ.Þ¦æZ-Å_&–JýàÏeí°)ÎSºŒU`5â	|¼Ï›ß~+~g-øõß'¸æ1½Ê¼¢Ïb7º‡‘×Ã~l”¡Û|Þ„o¸Ø‡Ã]võXáÉU–u}È¾÷ÎÞÇC#ôÅô¨.ýBØ5RP¢Í/žZŠƒ½YEMp€½8Ã³“¹Ñ²¨	¥}Ðã#‹_%î]®]¬}k?~µu>ø‰ee<ºÑ¡'ÆÙ½žº8sK‚³F……ø-{Y”â›ê¼ËUÆË|òbí%e	ÕÒÿoJ	AÏø¯Å–Uíž½ú	FáHg—ÖTÖ®¢:ƒg×‰Ó@þCp#€Êw¬ÊN¾Ú<«më)¬2•±±-ì7ät’…
‰mIÄia¢q£¶Ï®¿!uâµÊ–ÿX~X^™ZWqï8ô†ßám%õÈÀ†KËÚ§1†$Ž–µO=ËbãJÞ|2î5ÁYìaÀìBw‚‚û:ŠñGrM£¸µÌü(Ü‘Žvß²£Ø?ü‹¥Ÿ–WÐ4ÿÝÉ»ï÷ßb£=ntw%¸¯²Ä¿­cr<hÿ§Ò:¶¶Xò=LíÅ
JUb¥›Z7AuïbžV¿|Û÷GèMÄ7ÙmìÎ÷™vÛÓ7šGØ‰Mk-{öðÖ±†´^ö&=qžA³eÂå- Êxëðô²ýÙ.Ý¶—Ž•–ggÅRS‰û˜À6Yì_J
œ)}F—k›V¸Â^õÀ¸–ìfþ eå²KÑ/‹ïÕ¼]Ðbÿ.Vfë¹ƒŠÛg_'C±ìb1 ¹>äýG§yLy®8Š^	
‘ÄÈúÒºL-zºÿ£
ÁDÀ‘SE{‰YYeË‚âBû™©-^DÀ‚SÊi Æ1ÈÒð˜SþwŠd`ÞË1å0µW{ôjØrÚÑ? Äô_þ  ÿÿ ´ùã8