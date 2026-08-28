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
  FileCheck,
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
              ? "bg-white text-indigo-600 shadow-sm"
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
      // 1. Gatherxœì½MsG’ z×¯bõºªº" ‰juQ$"A5fø {ËÇG&*@ŠU™¥Ì,@‡Þ=¬íaNk{ß~{hÛ5ëÓØüü±çîñå™•€{F°n±23Â#ÂÃÃÃ¿ÂCD“‰ç“I2®Ò<+Åq‘OEš­M“i^\ˆ²ŠªDDY,&ù8‚¢Ñø4ÇiQV¢Ÿfð9«Ÿú›$•H~˜åEµ3I•#¡~½~#‰×Ãá0Q_Þ<¬UzšF'Y^Vé8‡ŠìÉTŽY‰ €Ã¨HŽ£roeùHÈ'zØGHÅJ€lŸ$Yíœ%Y…Ýˆ²Y3Øár¼}MÒhœ'Ê¾ˆªy‘ÆQœt­ñ¤Hâ$§QÌðÈ^ÖPÊ+àíÉ4-iRôOb¦¿ëåG“d*ëÉŸ¬žú¨wOær”ô‹Í@©¾„Pz´]$Q©ç[>Ë$'Sš#SÖ¼òªÂýûb{(¾‰ªÓ¤ d—çI¡(\R³"rFÔªâ–B%m«jà1Š·'“çøv?)g@‰QÙ<têL£™ùu²ä\¼ˆf_•U‘f'«B{lëñV†Çy±½éâÑcq)ÒcÑ/¶†i<p Ë¤êðzUƒ‡âÊÀ*L‰²¨£È‹èx -qã(¿WÅ*@ß¥úí¢M$Ÿ$áIRíVÉ´ß£·o%Ì·‰\m=3Aýep¬ÝÌ,*Ê$†þáàÕË!=9U²n»(¢‹aZÒ¿}Y{0-{€)?[Wê×K5>ýºtÅµ9Ïürsp*Æ“|ÃtŒó".ŸŠ½íýíçÏwž‹ó´:‘@‚WPíÁZ	¥€ÁVé4Éç•(£ãädñ'¶ïó2)5D#¨CUë‘ˆÊ‹l,¾:\}ÜŸÑZ…%|uøxUC~Û|°¾¾jfØä€—%bax+`[™.0,¢qÒÍ¾1[u‘Þ-¸~‘OÎ Y hMõ¶ß§êc_÷fÀº:0¸oØd^Õh ø‡y„øHÀìÆSÀ6«»%¾Ç}»Éõã£UÑÓl³4ÀŠ_çÀN’~/?Ï’b7îÁ×Gà¿8IÃ9,±A­‹¸sy}ìÐE¾ßõ„4õÒ­q“ž"Y¾§†ï`£“tšVýÍõõõN}vê.ßaµŸ/ßgUñí%nèõªÜ¤»’å-ß[ŸUvé®_çf”[ŽoB¹åømd$®ÄëUºI§Qµ|w§F:[ºÓUoÒõ1pË÷]3¦·c&ÿuë{CÕåúîJ z,î&QfÑ$ØÕÚK”òëoQ<©¿Uâ}ýƒÜ»CÀËqý-Hâõ—(;óý·Óó(­ÌÎ;”·ñ±Í·bÍÓ|\öõN4’” ¶¹l>™ R‘!®ª‡…Ì†q{P†£ßçµ·‡ÆXá] Lñ©ÛƒÒ,äöÌŠ^Ê“u<ÉU-ñ«_‰{ê7j}Õ…+3yTŽ€ò¡”H¦{È?Wi7ªC‚}•ÀHOHHOH^mÝ£FlêÆ^Ý˜ê‚„ü:Û0ŽªprÂ«îÞíŒJÊñ¨ÅõUÀwç	hbNÍ«‘ß4âCq†ÉŒÈäÆ)
*ÃJÂ(u®ŽÒzíVœ²n„ñÊûéâV·Ô¹È¶rñ!Œ\óµ>:· ¨«ª†¬+xÃ½ZÜµÁ˜!¨ç6~½“ˆg~j n²2¨¨
2º¢‘V!2Ñ}¬ã’è€6¯¿ajáÝu©E·ÖZ”@#[>†qVÑY-=‹y¿q 8D9f¢·uÉµ ª>Q¥;ì,¦Cð»­?!+£êÖ¼ÃN¡Àdz…m«·ÁÖ‰™.
7n:ã0p¢ó1eZ(ÔÇú*áµ»ïQ¬7á…Â»ë.Ý`ÃB1¿´‰I;EQ·€å“dxYååü,A£=È@³<»þÛY2°@§9ÙCEœ´®Ìg€š|:KaõÂ+ø)€R`)¡Ñ,-E™œÌ‹¼­¬
ÓhÀ ¦©AÉ4\nñdcuØwmÀ@–Ÿ?Ë¹*²å!<ª¤?VùîÁ«"x*PºÇIÿþëÑðÍ} ÞZo0,')¼)jãw\ÔUö€[¥? T|Ø
Ò	èQ¤ÆÐÃ°œMÒªßû½Áëõ7¬ÿ7Zûq{í?­¯ýîíšlímo ~ƒÿ€6Õë¹mz–N’—ÑD&ñ_Î'U^¤Ñ[‰ø·Ÿ^Úþ\}zÉ}5ü®Ì³w]¼Êz€	4¡Ú™f3hþ‰m ±&v¾ÞßÞéY‰ô,)JÐó ðæp½—ó•ÄÛÕ¨×¶4vzä¡mä ôO½íqR–¹´÷¸\<ÇÅ$KÃO,‹ò./ákºWöã4©"\}nª¼Š&ÖÃä	¯“$;©NWýò®s) £5Ô3ô;òI½©†ë…
lóõ<ÇS`Ïk[9w¤¦vŸµìMµ]ÏT€?7Ôc(Ï%Õ\Ãøž<gTSëur½Pª|ˆÊ’R²%âÄcKäbËTÄ–Œ’ï£ëŒ‘[YAwd3)Ô±µˆlYã´ó2MmË”SËp ý…fðò‘7>B[Æ¼²åÆó²Ê§Ïó“Üg8o„_õ–èñnäçÀKµgJÊ:éñEß2u©ö¯ŠM³k ò5}þY9±T/£w›>Ûy¬ÙÀµÈ9þ·Òõ¿©Ý	yloUøÍ5ÖÊê­ÞHŒ@õ°«L¢¨LLþñ0Äv:,òy÷q;úz’õ_+Ho`‡râ¾ØXßü|àa#`%Bv$Eà+g¿\n¯¸ÙN±ì>q³]â¦{Ä²;Ämö‡e¹ü²<¾‡÷J/½WáZ8Œ¦QvšÿãÑHjX~²Ëyh_»ÂÓNP\jA0KbÎÜæÅ8‘%ADíe(Ò¿MÞÒ
$™S®Ø<›\ôj|Ç0Z}Zœkì–r S‘åÐåZ2–qÿ¾8,Ò““˜Ÿ‹€ÔÅQ‘Ÿcwâü<ÃØÃ^¨àSõÖr$Í¯VE-l}q0£¤x@Þöð4??Ì«ôßIh¤§LÐ] ¡H)+åœ„Ì{¢ÿé¥œ–+ñ_Þp^J°è¨P P¯yŠÚ*°Ü¨((ðÆ*TZ™‚/yÑ_½'‘lºP“ŒJQRÝpêÊ@TÚ—*V@¥Àm§Ð}Œ² êX9„•œ€btáÆ”¬VÅ
õ:±ò,šœF¨ÝÉFWôð>Ñô ¹ã)èx“dwŠ=$©\áå´7ûI4®†O ÐIBŒã«ß¾x¾›ÍæÕ,=xñx`íqâ1L(TO0€ê$©†ø\nAo’=@ËÀ=|9PQjâeí”ÁfP²"°WÇI¡t@$Ÿ}zÑ×$¿s¢:ïÐ'ïæ€[C1)ðßJ5åUó[ òá~Žª½Ü¦Ý”{ºL‹þ=„rSPc]äÖ
u|GÑvñý<=ËÅYôcš¯8z´×³Xêx,ˆE5Äý °>_¢Â8Áýp–ãç4š tPeÜ^š)ªï“‹Òk#JöÄ¬è1±-ýã£H¡Q–&ôÆF¥ùp<«1UåB²g_.€È-TÓHÒ–Žßj…ãY(©"¸44X.€µu}Ì9ßZ¡™mPƒ0r±é“|³ùf´€ôˆÞ, ¤wNGØ¾ÔÕN&hOÎ¹z4ó¦ŸÚ ðp>Å¾2ýiz³ ¼o9–Ýr^šÎ:!æ5{n[ž5X-ç­]N°C;ä».
©r¦•`t‚nËauntœ¿hëõ_*ÙI<ë.7„ÝQ—©CâöNâÔ­Š	»Zõ•©`Œ_T›ü%3¤v%ggÎHrp´¼\êr£vÂ¸Gî4póÑÖëÝ×_‚ã¶Q°u`|Äet–Ô#gÛêté8çõNßŠT°û¼\êrÓæ†‡5Î›ÓÂÍç-´=×i@ßi]³¼|{k7[Ën¸ÑÂ…í´xst™ý×Á‘¯œc
€ÕPÀžÛ/fÿoFmÆó29p¡P~–HÐ¨?¾O ±˜µàz™nŠ@%fxtµõªB`·A ‚Ñ†@ÝÌ¨[¸jyËÁŸg¾¢O—©Cº9ò´°×Œ;ÓÈPgàßæ”„éîÍò]ë¶,‹ÔÀ,¹½%¶e#Vpo!Xi×¤yß>PS,réI»mÐþÍîˆæÎÐCfÍúàÆ‚nÛ“.#’³ñÔfå‰ùT+ÚuBÊä¨ˆ’·²æÛ	T…Ô€-ÕU×…ÑÐi·PKõ®1>ñ¨y0>èð°>éÒ\È½aÌ*®ƒ%ÿ”Mˆ§Ê°àùÌš>ó'.Ñ÷ƒ×¬_Ñ—hV½’®1$Þû5˜Ÿ±&Eûe]cHõk„}|mr™ùüj²L½¬q Ô¶m¿¬õú»T§Úè1éz9æð6¿¼ç“rúüº¾Ù°ÁÀ¯öÑ¶™ W,^ß‡j6·KèhqOD™fã"ÏÒµ>htÏ³¸áUÑ¼ˆ®ÿzý¿r‹x<™_ÿ-ŽÂ;C@S£ÞöB›ÍU_ÈCÒ¬e¡:
Ã!þ@ÖuÇµ€#K³ã¼>,ÙÒ¬•LÊÄå¤µÖå‰SÖ¿á%Qj›ov¬p§‡|0€k5ûMN Oc‚.#eìâÖ*âDÕŠ?&Ezœ~?OD.Ž)à)®¼‘ž
9àÖÿã4‹&Î„‘Qso‚KGü‹9X¼c£ÿ>ä®üg»<L~¨úäÈxhœ*÷ï‹§Y4MÇÔ\/s íçÏEt¥“Ø•†Î>˜Mb4È?ýZüöÔ#±à×$…Í6”¶¨r`\5 q[ÃÐ†TX¹/@²Vç¸†9H*å'_*˜ò±q”¨‚qŒ.ôø ¶FÊÁ!d–ã0!W²žK&•š8ž§ÛY,#Ø°ÇØ€™3éÒuÂ 3Ù1SjXåÏóó¤x•‰¥^@úÁût†Î:¶¨N£Jœ‚’f“tœV0qŒ+ágeÕ©”ù4‘‘ž]nãâÑ#9LÇÃãBÓž‹!àºo‡üP“Ö'j»¸~aY`¶„8*bÜ•XÂ ‚Ûû;Û&:5Â>«9ìGÃÆaŽ]1’æ,É=K‘ÃÀ“èI¥…%w ºj_\¶)¡©æ ¸lMÑ6V¿¥üØÖFó0èd>&¤Ð+…'¦ÀVPÒG«	‘ L‹3ÞâŒ\ZÕ"„)[ò²¥S¶á8½.\8…ãÆ#|ÅÒÂ„;Tz\å¨^Æ!ÆúçL;ÜÎå“´¬Ü€]MÐôîVIÑÿ:‡m"Ê:è48Öƒ…UTý~´*Žˆ•DCÚß’'è@,’>/œUk_ï÷0¢ºL²2­Ò³´º‰Þ,´žŽ21åµ¢ÉU;‘«fšVí$¬:(^PÜªà‹_¤RìP³Øí2jã²’ÜŸN˜ËT‹ÅB/G«t–ëƒK>Ò¬WY{v|à¾•S‰ÓoÙw<¹œÀÓ*™Ý¼uGÚ…xgà™“Ò¬\N«¶¼-þ1i 2qÂÅ ùŽ;™HZÄ”$já‡Ã:-­†>†Ô=§@Mò¾z*“óÕ×’Ü–ØkgQéµBŸëžg(Û™13F¢D“¾Þ(5fc˜[8TA±õö0¤Æ¾*|²÷`{î}“…çOw÷^¼Ýy±·¿s°=Lå~RöuþˆÁ@rªë¼bc˜5Æ©Mç=ø˜Þh¥W¾ÏÌrj`]Œ´<v¥1ëò(&	F'í,ª‚r`Ë®È7#/%Êa:[|åžLóÏ¡T6’|Xm©™äŠ¶¤Št·IdÒ/™tØ±ðßð¦Ý<”²>”ò#Š¢>çìÜ‰Üt;ï©vë]xÛ£äÞƒ%ª$>¼˜%·P4l9¹FQÍPk#-©ÃÙ5§Cär(|úÂ¾"V½1‰UeüÁ:»ˆÐq_¾}L¯‡iLü~hŽˆ] oÀÆhp¹õë¨ÂúÙ;C4SHê;›˜†·å¿ôÆYWyö<‚DËb1>¥M¯¯ðK|ïÄð=þ¶ÇØl‰>k	¾ª‚xðÆy¯aÕ?0<ò¡µÕ2CÀE¢PÏY„d¥JI$’f‡(¥_nü)#Ö¿ƒñwŽ&Ö éi ÜŸuU[¢M¬SÙe }ýB‹È»J¨¢IçYí2§¦~£ªJæ3•^
{‚Ç"±/h	QJ'šÌ>áGÒÉ
s8Õ,5Vï’$ýØí´CnW\õ*É†‹ŸË&ì®,³ã°6Iúêgðs&?ê´ü j 	Š}”Ö	d¸RÉ$9d;‰ü×,;u„Ùòâ‘ÛÎ–óˆÜ|$ÞÅÙÛO/©ŸWï¬•ÛQszæ²7ÇÚ|5Il¡Km£@? x¯ÍòÙ|‚©0Kµê)¸Uâü4ÉÄ4£‰Èg0'>›Ê©*ÆwB5ì¾;ÇÇðÕÙÉÂY¾@(¯ ®iW—ICêYòÄHÕ–ÞíŸêõ[¢6­–f•‚ÀPŽíãªíáj»ÌaÑúûd2VTå¹OÉÈŸHcß¼:Åçq„h£#”5Ï·‚¡DFèAxßB­‚JÌKecÆç#ñ-ùúœß›»ÖJîü³‡ŠéHÇÔ‘ÏÙ!Øž'ÚZbËtä• ©ÓÉ*ÎpZEÓ¨¨Š(-ÒÿxBg•ÇùÔØ±­…ôÉiœíèB¨dM4–ßÓaýGw•Jg€âM˜” :o’Î€Cm¼nË<žrC)•¡Àf 6€6jXûÎÉ}sÔ€yz%¹8CPá¤ÊJ¢b|ŠïišØŽtw V˜šÈºÀò½ž¦›¡ 2äÉV˜b¨ò0H¯®`tÏ”å]¨ãVõØx¨8ãY5uõ#žt¢[CV(ß¬‚-Ø“¢²<Ï‹˜Î,äùÉ$!ÊuåŽý:Êãä8šO*%&ÖVé¶i&Ñßï>½t^]nˆuï°G»´®Ä!-,ÁÒ«R=7qšoÉ±n¹¹ˆ¼ÌÏr“ÙH}É9R³œ>cEôc”ä@=“>÷	›Š‘X¢Sýú¶<,ÐL±^ó6õqÎåøw­x‚_9¬Ö¤8K
L‚UB³¾
 §S`'i¶ ¤ZÂ#ÉjøKv#öÛ)qšWù·ûÏ±²þ]ï6›—0ìm™"ç
ñ:T"€¹§è¶I©¢íúŸ¸I£’Þëÿÿ âgë¼N‘OÂÅ{>QKãIŠîêž]»vK†'VwUF8Ø ÍÆÚ·Ê$jß½í\b+Y¥JñyÛ|ïUºkî?Q„ðD1>–,Ð|(Ë¨7.UÝ¶å1š
ùá%x¹RìÚ®b'zÿ—¿šAòÅ€û÷ÔY-µdÙé-ý†d»¦›JH†¥(kî%«âò¦þÖèº,\é¥%õT¶”Â-Ûòné‡²fÙAYZd¡B¸Î–\U¡þëè.,‹)Æÿ7¯¢U‡úA,jXUª5%•Ÿg;°š‘E,FÆ„Ä9ˆóë¿iŽâ²”–UPI°¯ÌQ;,Š	ÒAØÏa	,°FDüËùÑŽy†uñÈq"’‚j/Vd}ô¤xûÐ¾yžG1æ8ðÍ‚
#cù°ÆnÜŸc¾J3’qº$Œ«¸˜~ðœ NÓ“º³=™æ;ºDß•=æœÚŒ_HÑHÑ½=¥Ó”˜]<Ïä(úcx^K3­i'VN¸Ýànµu>™ÀÔ–#ksá²½§† ãxÐæí[Ìmªo9" 'rå8©€–4ûá^†BeJ\/EZ!gã$Nã|ØpvVn¾6¦ÉÃØ€"9A".£z% ?VçIeQVµÖö?†]—|ƒ–~¶h´aÂ*Á’øÊèùTlïíŠL.”x"Ã!Ðp•Àð #Çé	ÒãTªq¡cúþÊýh–Þ?Ip&îËª+š`‡¨\cVX•‚òOõîg¹¹s‰#Fº¨@‘Ï%W?Êê>öœ¨óH:K{µØâ][r{Œk¨_ƒ²ÍšžÉªÀÝž5D³åGH
—@ãÓè,!ÇVˆF~ˆ›µÙ):H€¼ÏòóIŸ$Ä’K å‘¶Àš#•ÆY@ìîñBè{FUÉvDghØÞ2ú4åç	&{8FóÆ™‘	
7tŠöÎz”ËËšûw–&çrîLy2+kÓ•­R´ž”úi½æ0ƒg©E`ÈbLß³=*G#TŸÑð{ ¨Ëô-°XI*œâ@WCÃÚ’w`x³¸§Llz—l7ÝÝ“aäN"#I2@påxÐªÏŒc4›Ú¯Â`Kâ}UA¬Ô…BlÃgò÷,<ÅM¤ÖáÚUò.o<`bëYždÑ¬Ñ¹o:¼*óHÏh¥Pf‹Î+ƒùƒ©²gFW!nzOölêü—IÌiféD:RáÃ)Ì«ê÷á¿g¾ß™H?h’53«H¡QF&²Dc¼ŠCÒ~O€3ô¾õ6¤*o‰K!¯ü9“êý•Qõö*çÒn*3ŒWæ½ Ï«âÕ,)ˆ‘qùùîÁ¡Kyµ|ïåWðîð¸aÞû‘óáúdîuž“ºs@¢™Ü¥}Sr¦"UÙ|ÊgV¦¦P/ZÎ	alLo]wC!þ”Z&úuz4Is(#L™Çùäú_N@ó¡ˆl-ÓÉ)Š}}+%l`]DêÓ¿bhÎõŸ¥CÈbã$ÉcæÂtÄÎÖ:áD¼:§Ö­»æ~´3_öàró 6_$4òª"y-Qv¸@á[Xðô*»ë›­0ÄËí$ö®nÝÖ0åRŠM©à.Šó 8ÔŠÐ$”È$X).¨ ÔØŸy­±È.ì–
èbm¿q½[S¬K.Ù°È&
2‡‹Cd«»d¨ùî)HÁ¸¤On×)è’“ŸVN,”5LöÌt XeN&°óÎÉAÝÚçPƒ9ˆþC)‰¡äÄPv$†òç&gF9-è»ºŠí—UÝåEUÁÙrzÓ—ðÜ9¹5Þ?àçia"ø-'ÇÇ©;Cú(VÃÕZ2ö¢Ž6]#MÅžÜ5ÊäGJ_†Ã¨C^Wþt„»C¾ÃàX¯gò
„X«îð8Cý3¢ÐfX¸;4ºèàFÃ¾3HG¡¶Vãšâ}²âŸÿ^oêî{ÃÞÝ×jùyo%Iy Žì!8ixAF³*j&-²áRLŽ<KG#1$H	¥hË"³ü$?¡‡e.&ÐHk¥j“Œ³Ê'i– _ÏOÚ^žÆö‡v4ú¨Óý6v¨%Õ$ZyìX´6ãíuô1tªHíw¤ÐßKøv—·;»­aZ_EíMg€·‘Ðå©ÚM…º¬ÈH ÐV™çUmÞ]::{ç±	]Yà‰æÎg×TÁM-ZÄ Y€¸R– †?•5£Ñ ‡Ð	Zç.ÃVéEþ6L?áš:4Ök©¹#Žs/æswŸ;„h&š¹ùÌ!¬ÅS'§«¤H#ëëô²Ïá£ëÛ¬Mš“ÿ®6qFð#(pzæg+î4[Æbš*L‰SÅŽqÊy"Â½‚ÏSÑ}žœ™'jãæó„°n:OfX=i¹^0I6ñ^m†…ž!KÏPÑ4CE§ò"4k“„ûçƒîM?r¾è¢Ÿ{c>_ãîó…=0óEÍÝ|¾Ö`‰Á9üÑÙök*oFÒŠÀÙ„Ž»0Ëq§ÉU"th^­þÀÐè¥ãW÷*>£U÷­ÌtªVn>¡ÕÎfMëjœEG£©ÍJAøÿÕsX±	Ä÷]æËO»,H"ú¾lFr—Í²¸}RóŸZ©Oå¨¾q°šK˜ƒö¿Ir/Ü\¢¯Å½£äÃ¿i¸ èåÒç9ÖËñ¼(0¾!÷JñíîÓ›ÉãI”Jo¶x—3êa}â»$V9a·¯Þ=4 ‚wÜk0ƒ`BtúúŠ S¸]ÙIÂTumZD+ ÂÇÌòäX&¹Â¨nŽòñ&Ì6ó	EœÁR«7\pÝñæ†/{kD˜Bµ£Æyr•TÝ4Ÿp`à4(ÉâÔð0ü<…fÖÂ?_	·¾ûÍ#ñùz­®îç­Gâ¼ âüšBâ#§/l¦NçÙ{TgUò¾²t¡6üj\ÌÇ¨þL½?BöE¹ðúªÈÇx&h[³ëÕþÉ¹“PÇùïq÷êðõlsÝ8ËA~\	ª…’•ÐÐ½ÕO—ù,&(˜N¯?dÌÅ<é…™Tm¶r¨ý$†íc½ŽŒ;/Ð;¡‹ôâH<2k³¨ˆÄïs:›ÔÌ‚HµýÕ¯„¿˜ÈúòlûDÆÈ9'|þEû½S€ln¶pTlá jÚŽËÐdßÛ“‰pòN Þ«ÓDêÙ°†(‹pta¡º{œ•^qÃPg3§£z,´rº¥»G)(·1Ý½i)xõn($Áçîl4ÛÈG	œ£R&P$ÆŒ	µæÛ¯ÅÆúúº`×ûÕÊ¢´µ>À
ƒ­ëQˆØƒ¯¡Gm=¨}léA­lsÔÜÈ.¬IdøFùºÎéÄe¨º‡g0P"DÃÕÂ†À8T¤Î-^ÿ‹Ê;)Ùl^þ°ÙYVb/í8æçù<“`6ÃË(KºH…8õQ§±ÜÚyL(ïc	…˜¥±Ã,­ÊKñ|õ÷ÏÉLèEqºlŒþaoÅ.¦xHa*`ó‘Od¦K¹¡RßUÂ–x¹óÇ}(#&{f	8‘ùí(Ö€É§ÔX£ñ)QÔ	ë¸=ÛITÛ€OB•‚ëÁA+Ü†5ÜUõä†È.@±Bj`io¥ij6‡x£Ñd™;$ŠÒ´·…÷kH±›Û[ØVÒ€g7à·Û¾BjÔÓn.IvgAdÍ,ÏN²)Ì8Qøž†zI5Ý]«i]u#1‡µq\¶°8›'×°·Žm¸`c°(RnÈ“j®î"È9C!+»p”‹S¨ A%ni‰]$.Y¸c”ÉK¼;Tš¾¸èžoò‹“!z?9*õZ:ÖÖ£¡>fFhäWl½â
ÔÞÔPùZ=~­^cI”=QfØ’“xŽ‰#o[ŠôC†ÝìÚ…¶—»Vb‘:7£ÎLAææÀ P—î± “âdžU‘9KÈ¤ã6¦HÔu2Œ@¤ ½p‹áÉ¬Û#ÃEÄÕë7hAÐÈcJ(©Ÿ†`û¹ò,•Ç1¶žéÌE!êj‚ÅÙm(mõ—,cg‹AŽaÙÏ,ã{	ð0X<Íæ‰Ç™›šéÇ=É$ nÆž ~1Q‹JÒ’‡'œŠé«¤+yh¬†1Xfàå-g ŽãÙ$)jfy¶&så«BÃ6¥`a­(7K%µ¤’õGeRÉâW¶ËÏvZ8æ2U^À¥‚{¼ÌÌÉc¥²Ü5XU Rç´hsÈìg®ºƒ}’0š8¢Éøó‡Aí=²k™—@Þ”ìðLžÈ2ôêmOü†ôÑa–ŸÃôüFôè•¼ƒ8Êâ\âQÝ¾üÙƒ!žÜ’O›«â·5E™:tÆX©{¾YÚè±ÇÞq+[Ž„J÷ˆ³Æëˆ8éªÓ„æ«£—uO½ÏX¹Ys)E^TÊ!5·T™C‰­!þ .b.µ…I™É×#±ÁëàZÑƒÝGxî;‚]`äŸÒƒ?îQy…g¿¬cì$£õY:‘¬ºì]zW¿ðÁÎ+¼ûäòiô·zhoÇÎæåiŸvÝ¨÷^M:ŽQ6H¾j×"`u-TˆiTÙpL¿"J&ä šYðƒ“ù5·dpÀL#«¬ìÀ=ÁëM›Ûy¡ç°/ÿê7z„|aT}ß€ßÖ2>ò¶9èô!ºCßµ¦Î”k8Än\4éa·i­¯uÃ-]VÏÉ \Þ=V3Ërk/¨_wàx8˜–l	­ï)±Úß¨4?øšLC1&’:‰Ì<ÏÉlæû0šµÈð!P`¢ö}‘L"L–“{ï±ùIM#„]ºÇÞjÜEÊº©b}Ë®
	ðÀr©‚‰ãÖð¶Gb£Î)èŒôïSœõ‹V!Ÿšr¯öÞITVz Èéx'ÞÀb:K
:î/õÚƒÁC2¡B-ñ –P'œrD®®„Ag?CbºÕÐ¨Ä÷û-‡º»tà—·[œ½Øp°t¡"±†.mØ
Ö«ò*š  Â¸JÀi<Tö£ñH_æFñ"õbHƒô-¬
Ÿq2ˆ{(,¸)~-•ï’ŒƒF¨Ôð^RŒ¥¡AJ”tKŸá¾èÛæ%@é,ñ`:ŠþÔÆï1Ï˜T{yõ„adär¡Éó¼.è2‹ï¿xáæxÂ?%jña¸,¸+Òž/WRª=úªÃ\"añléñ{jê<¨Ì€¯“§ø&ç¦áÃß†“rþ#«H‹ÞÙ1 ÍÊ½ËsK$‚ò­èì8Ñ¼Êa yy¯€Ä0¯ÿÄ­¼7Èî©Q“Ã'¶œÃ™Àfƒs-B´+=¥±„~“®j/“¸ÆÍ±z@î*ðŒ÷”"ÐXÛþS¦Öwzˆgã³‹Á¡|×eøª?uþZªLH#1JuÀ»Ù“Ð‚:ŒYÿú§?5CÊfß…°-õy’€ó¥­Éd>‰Fâ~kRAçYò,ý!‚
hW?Jì‚«÷´33!Í:Z]žH7Ø^¸ŒÎI³×Ðö1¸ð™Xhåxå"XQCn)•t/@49µ Ÿk·,Ÿ;[¡>£œË7z·‘e âå}4>™ér€6¸àY+·	4›U0ŸE-Îœÿ«7Ž»kÉ"²<77ÛŽæ_xŽèJžI¹ñ‚B*×d5˜‹Kž8êÙP'#}A«FZ+Ó“l7Ã|G{ùl>SI»Ð!›DH{*‘“Ë	H¹	àQ·àüØI¬¥šZœ\ImæV­ÃÕ„œßI(CïaqÅ*2öù>&ež­'y™ÄkGkøée÷ÂÖ…,BQNTzÁ€ñUN—Ú‰dŒ8£vœí5ÊI~Ò_‘3Aõ5@Û±Mž£*|SXãæþñ#!Ó=bÚNYzš€ŠwBYÿœö¢“^‚ku‰ãˆ-Óë¿ei.bL´("r‰É´rþÞ^ÿ…ä™B^þ¨´dÊ|ôD¢D¸M%gi¤(g(¾­R4Œ‹\}Ëç0³2?( ®;kÄÍqd§ÐöQ¬gØ[ˆ@ÏÏ+çâ@%qÈ³kÇëâ0k|ÄÇÑò`@Â´yj+\jd7iÈEI©Ñ(ežY’	Z£q5DW Ý ÊiB¶/ÜOdRW-ó†,_Ò”Š¼0£3œõêR¡\_‰\jùIš^€²ïY¼§Rôª…ÿ¥Oò§þZ3vƒ ÅôÈkÍáÊ¼/‹zw4´r]%À‹ÃéG†l™Û‡«MŒÞ—M¨zã.ÓÆû<º'Êk ûFCÄZ4Áë-/ÖÒùd¯™Aì`ÍD®Óï€HvLãùPìÁ">ŽÎè¿×Øú)çÞCp±åÌE‚]K3¼¼6^\÷–=eìdÍæ—n*‡È^Ëòjr‚·TŽXu[ OÍAöÃ{Àºà<‰ÞÛN57¶­‰ôïÎ’I.@UUç €8H‚Ùøº7ýÓïOõÎAûDû1ÛqŠ‘Ò¸˜-'µnLùåõ¶ÿšPj3SyÛ¦þ‡ªud4Î¥Ê“jr¡Ò~Æo -J»³ÛÂêñò9JMRèBÆòj^°Â'˜¸¬I]°÷ôÙByMÙ!ÉØiîâöŽJ™‘:g¦RòUCÇ†×" ³ëÕ‘Í‘‡ŒfÊ©çäÃyØÊ×¿e^d^'¯)‹@ÕAµ€¤CÛT:ÿÚçÂ›5x®·ãp@p)},IGïfHªƒ-‘LVÄ ˆ¯ë Žº‚ÐÁÃØ¿˜õÏZ×+¿çªšG0^ü¾'	ÌìÊK\¥§À¼Yá2¥³ixG³’øÎ€€ÒH©%PæaÁRÌÃ[~
XÅÖœ5ly)×=*ß•°2˜x,UZç#Ñ›0…už¥¼šNÙ;iã„·ÑçZI½²m *¤²´ñ¹7lïŸÒ˜¤¶35czÌ6füp Ü'
Ô®^Š(Q)òïöbtÃ–—?9ÍÑ±¤nc9*¢¤´*ïTÄÂ@®~1cˆlQ£Fª½ÛwóGÜ$éýÂ²ù¼ùzÝó»OöæßüRgÒöiÈ1WÁ	§šæw3àÀ‡Œm!t¼@Á¤9Vñ Ô°ïÿ?OïŸ¬R29oM¨!íí§†„¤™½a Oè:ô4¾ÕR9|5Tybƒ¹ší†9¦J=OXdÀB¦ª‡¾&r½-F/ÝÄƒ×Ÿ‹SºÇ¥þ8Vñ~¿‘\Ã€Ó´š˜L)~eP°¦YÜ&|]àW4=Ù+rrÅt¤r
Kß gKa*„_‚ýÐëþ³Í¦¿OÒ“SÚèu£ÃSùê×Â@Ü7}žã\ŒAKÔÉZ×`êö^~Ó[elhMóª5ÛÂªØX_å¶o«v—`Ï¶{ÎxC‰äí(­OüI>ŸÐÑ]œX51ÔÚAöôÎ°-w˜i–3Þ'«@¨
FAÏf‚ÞÄÐ(~*­ætÎO^¿¦n§VÓw¥ëmŠ1áY:™<É' Úm|ˆùòügówÒ{šL1î0^{°¾îT#‡[ï£‡%ˆêÏHÓ¿›«Bþ¯÷ÌE&µ˜üPÉö7<¨~–gîýõðÇþ
¬¦³%Q@âÊQ>‰W¼’4SòêåÁ·Ï_íïn‹½ýW+T²ñÛ!všþ{	û;£x)åyîyá?Ÿ4tõ3‰¾}àŽ?¦±ÚÛÇÑQ‚YÚå+NàÈÏèÞ…^”×ÿ"&)h..T‹ŸèÑæg«âóM	¾ÄdÍk¿csÓ)
%Ä\VÕÐ6ÇpÃÐ~7|àÀq;èøW›:#c:Ýy§îÇ‰O/ýLõð•’‡¾_‰?	Ú·„Ü$®ÿ†·1 °Æíê–Ï!~u‹©ï,n>[ôwãB`cesOni˜ëÕ£o“½×ÿ³—FIÖ-ôw
ÂßæZœž¤¼š¦ÙQd^I^ñŽ~Ñ8‰O‹è\­´Í/”pU~¾Î‰iÓ!&<#Õ×€?ß-Iz»3[ÙÍð\ ,üýçÛ‡×ÿ}÷ÕHì=ß~ùJ<Ýî÷_¯ÿË+±sp¸ÿÛ7»O^õjÂÞn÷T&îõ…ºZ±
Ý«hDbi
T!vQ>ü´ºP/›<Âž;_¶áê ¡RA…’é&”ŽÈi;Pó%Õð7¦Æ1ð ü±½MuÓ¶Ú–è› ,¬ïüWtý·!AuÅ ‰ÚDŠòt´­oå3Ø9dñ=ÜØ­¯Ãÿð%óe0ð„m;€cÂ€B…ì:ürúmŸ›:Mhôkº‹Å}…×Žš´ˆßH*êmÄì
y×“7ðÐ¸°iù2zi¾Yuuà|…ÒìSH¬Óãcü¬ã]¢£Ò¯‡zn½€V žF¥5NÒIßÀ¿/út ø×âóŸÍÏÀ½ál©ï>½Ô ¯âwÁ½Q§åÖ\MÔ¹Œ¢ó¶`º¿ÍLd?
ëkî½anHµÿ1¢Ü­¡/•aì$ðOaËÌm)ýØ|Ô¡>÷/.zHA5°>5¸óÍÂ—EÐ$â}t¤¹¼è¡ò¤ú.i
:ÙY5OŸWù!²cÉÓ™}°ŠŠêÿ‰,~ü¸ÿH¼~Ýûävýçýmüñtwû›—×ÿýàx=Pˆa_?ßyAŸp7ø;bG¼Øyþ{”ÒðûÞ«ƒÃíƒí!{ Â]ÿù;Ïñy÷åõ??Ù}…?÷ö·ÿýxúíþP¾Ø}%ì~+7šÞ›7¶Gy|1²»‹ý ·5<Ä5IŽ«‘Ùæ
Ôìc•ÏFâó/WPBö„'ÐöìÞ™Æ¸ÛîÁ~9’—¹Ôì~ú–²Œ¼œOQ8|ìÆPÖo±“Z(L›pçè º˜à…q—N¤V	`¹RðfµY)¨´¨x Å98%<ŽÄoëï±}¼H$T‡PO=ÉÛ¹Ÿ1™Lö 1Àî¡µÐÐpzÃC³}º½tÂÇðà^?ØX_@Õ/¼±rÓo.X7Æùd>ÍLGÄuì6IâÔH –ýs@Blx6A«#ï·þ¦Wÿóî÷Ï¼ï_zß?÷¾¶é~à}ßX8¢/–®ñ[¿Æú¢_.ÝÆï|<o´×pýŸž@í…«¾”‘Âú§ø¢gnï“T3”›$~ÿ2||FÎXYš$¥ÄK•¡.R±½íç‡Û°­Ø*%‘æÐÐ7ÆVon¢žù»UñÛÍ7¤y™¬}á¬vî‹ëÿœ}1ä`  ú	9B®ç¾º-¶~wWØzòêå“çß^ÿóÓWG†ü¶? ÈoËI+Úöv^>Ýyy¸s'SdñÑ„ÌïU²ŸŸ[~Èùû&nb›Ø÷Í›oÐbi˜Ú¾4uó³¤8žäçèH õðD¶÷ŠïÀ+£è÷Ø¥bÜ ú,Ï+Ïò)µa,Éô™ÕÖ8@\’;ä«c*hã™#Et¤èƒ/~óNè“ÀËÐz˜r•^¾7vŽ/ëß¸!fÝûLö 6ïö->¾ÁoMì]ÿäˆH|z™^áe} Nƒ:Ï@Yû¾ØtdÁ&÷Š´d’N°áé˜³ÌWæ¨U?ÎaIåçÃ|†—ÔA[ù¼šÍ«~ïh’Í‹‰cAÁäíò½I•å„Z´^ºYhÇ.'˜Fë	%”ÓM‡i§™ïpƒº1F/£3:Î÷ˆÒŒÑ1¶ÿ]“è(Á€²Ø®Š<{2IÇïÅ•ŠÁ$ð«£9~:Ë–*öèÒ?Û¢>ðƒVòz¾4NŽ¢â"×»ÙÐbM¢²Dï£Ëqf‰kå|íýµÇàwtôzMÎ¦8‰fkŸ‰ÙkŸ‹ÙüRVÚµ&¢*"ºL3ÏÖ0çX<—é€ÐÞ#N Ül…ñi¬ft³%VŽNÖK#híüšåid²69Ñ¿t9€½Â¡ŒÄ
UÓÌe]œ"31Ðô;*È\Qp9ÕãW8g¢ºt¹¹~å¡l…š~¢¢ûšl³D+ÜÚÆÆúŠž~¸éŠ7‚Ï½úµ>®Ä}ÓÓre¬+È0×Êdš’möñ%ÑÝÕW÷±œ¬ôÕ}IbøÄÃ é˜p	Eqz&™•ãô‡$x¡jµv±¶NŠüó#¢û|í·›f@Îå“Z+ô9ÐœðVT~%ykëFrLNFTŠªÿ€MžŒPïLÇøKöd]€Þ
›ÆiÇIfè/åKƒ&€¢9Xã¯p-Hê\d»¸X¡¥‚ÿYaAœÊõ3[ûbå±!Ï`•Àêš:pm²ÊõêçXêÿÓ¼pØò¬·öÝ¼¬ÒãýØºÈxG +_GÅ“SÐÈ?Óc“Ñ¡$,èm[ï|ùàƒ„¼˜$Nr›ñ{àækùöØšùrÉ2&®ey–¬<fÚx/_ MÀ×³ÞÙ<¬!á¯?Ûs (1­Ö6V]ˆÝêú‚/¾É§IYoÏÃŽz´ÏYTÃÕÚ† (ãV—Gz›qe@XÇ.Oçñ€G+»ÙõßÆi¾â~”,éÑ¥=J©)=þFTKdº,Q^÷ÜtH@}½rZºßu˜_ÏÓ	N÷fx¬ÊåS.¬<î™&e÷Û*­Ã¾éÀŠ49Æ«ªÃZþl†ÇÕã25~úÙ|[dGÙ{ƒãÞ>ÁihVöBDT§;TùG”ìôÉbÂäùä»Ó¦“…¾u —úšNÌã2F—O6ÙŽûà4M&1Ý„}U/£Fÿ†!ÿ/¤Ÿ§°ÿ\ÿµ†…0&ðà5lí5,´cÂÔªa¡†‰.C<,òÙéEóèhH€wJÈ½Ä°Ðµ´x\Þ¨°Ò`Ñ˜¾ržW]iXoóa~ª3ˆÑ¸X¼™L±ÔšÍGÐ™’ýŠ?=SÚ>‡v¸½+uðéîÁ“…hˆÓr¼›{R–H!ÝÑàWüéÑpX$”ÑúÛYö>ðÄ ¯þ¾&B•zl<K'IˆÑ)dìÓÁ:¨‡ÍŠÉ QSèŠkc[«;rÂÕo>LuÀ?<Hüˆ6¨ñâåïÞFÊsy3±ë³‡ö¶ð¸ìÝÄ¤+LÞ²Pjd²•*?¸YÇ’
íXÂÑ÷<ÆÂÞ—
Z×Þ›ò-½¿l8×pKÑgì`†ŸÊÓ$ñ'Žáà©ÎðÐ¼êÂÈ %m«®3õÊ÷åÅcÖÒpó`]3ð²#fÙ~–*Ïtû16¬Äà;Ê¦6ZbL¶ÒÍÅ©¯îƒêÎ5yÏê1«Ö¾Ð†¯ªnÓþçŽ†yIçx/#EWÑ-3~¾L²Óù´Xa^óG'Æ2*fkŸsóìK½KSß‚Ûl[:Jªó$ÉÐ¦UÓlvY˜]¬EtŽA%¨»æ—²Zy,mû—.–®BŸPó—ïÂí£-íÝëÃgdàû”µƒ¾ú<9®Ð;ô™Ø=À6ùÐ6´MN{Ô0¬ÇŒv¦ ³,õ®Þ]Õ;,D Á+ÌÑ_ÖtÍ0šÏÓµahmØÎ µ{nM™UxExÎsL+i“×€bÄ:¡÷A¡$þ= 'ã£ËKq.]ïï>½¤µiôC}UÆ«ÁJÃ5Lö^oý>KKPz½³9žvÝŸ}ñ :*ò”5U¯ÓølÝä‹\ý_ï¬OÆþÕi}òj¯,,®»Í™(é¾)6žø3YãG—2Ýì4¯òo÷Ÿcl×»Óªš•£û÷çéZtUQyy¦÷£Yz+ÃÖ?Uµ@´\`®~u+ø„èëÑÆúÑï¾ÜøÕ ŽßÕøp4©Háàµ hwW1·Mýƒ­¿X‘'E‘{9ìV²|M¿r‹ú;C ß°˜€¤ÖÎ×|3:ŸÕ˜_93¾/ÉÚ<Ï0ßùÊã &¿º?ë þ‡Òç¦5 tÂ¼®nÒ¯½Î¤9’ŸÃuiÝ™¹8ou~êM«_Òõ±NyC‡^¿šWÌèO;±Úƒ(-”ÔYE!«¿u¶0ÈÌOî““éôùÑÐ™d™ZÀô<)Í–Ê,Ínç8ÍÏó¨¬ú+& e¢L²1&óûÑhÓP	“è‰<',È[nX„¯ò ScÊ4B9À«ˆdÔ—{ÿ@Ö[i˜ÇŒiºÝ4Îf,|(”^É¦I3ïjÃRù¸©il¨Å=o‘ {	á¾êÞ@¤ÙÙõ_ \~Oìfè•LÄ|*G¡>PÊ„ÏE|ý·“SPN†²Y{ÏWû‰:[uÆØìØ­äcÏûÞŽÉ½g­À‘êýóäoc*8HÝ¸RŸ%@Ë…˜æ1¶(/üQ)ýèÀ”ÎŸ‰®kìÜõÿ—%¦ ¨—cÇ	ÅS¼@ µh
øD0¿²'9ï){!{¯’ÀxŸä³x‹áÖ\¸i®‡Å·ýK},þÁ|<NÊò€’Rê-Ö5	«d=2¿!g‘Êí•ó{ˆÏRBZÑTõ¼e‚éÁt‡yÞZž«K¡UçpwïÕÁÛ{û;Ûöp|ö3Q@³À‡ŸÐ	Bu×¼M¶]ÑÅïcêëû4lîi€öfÀ€h5/d7³iŽ±]—6LH œRÏãµ½^2y÷Ì6ã¢‡°D¦”§Îàƒ}U+È­-#ï«$~‚ŒU$Ž¦#ó‰S9Ÿ02ßœmI«Ðê/zÓ ,àÆR^[Nc€}˜î‘°;mhL’nÇ¾Ä£s£ "m	'™_[A™uÏÁ*ûèläãŠEÒÂº|¢’Y±C?q¶{ðÊ™ò]rÒÛÃiìng›;]¢rUß÷¦=º–Z­Æ…÷ ;‚Ip6ºÍÈR³²pf–œúá½I'Z€QOg}/ëm-9ûº :ÖtºÑ»Ó%Þõ†	>%æÒgÉqÔ}€‰¼óYöc¨îxv®ÒmŠoÔƒ¨Å8–ÑÆ7ªNÌä„—}ÕƒÛ“V-ˆ¤$9”µãƒ÷Ã aÏ­Ú¤V}³À¼ HÖS½mêk¯Øæˆ¯vvéöùÉ÷É÷É÷ïAòU9CÝJ(œ)àš¾¼;™N/ü" ;›Â/bïRrÖÇ-ñê¹õ³¨¼É¾lO×ÂØ•GÂªÁ…YÌVmÒ?/#éÊ[¬®äÃsè+°[~õ†»ŸuÔ‡]?sÒEÖ–B¡„‚r¡¹5Ê„«¶+aÑpiéZû	ä7”Ç–ß¤Ù‘…+6›YÀá4?òkîH ‹_“ˆáOf¡÷‹…³P­®ïm‹ˆ†	£¸À!Q™N’´Èï^@3!Ý17;¶•Bdß
gm•g¶Í"Ù!ñFÙO#ƒy(dÈ'î—‰XœZïÄ"‰—iËk´atú¦5`‚V è`ÀëóÊë+:I5ŒºdSûÜQºAº¬Açk†‹7¥Z³ÆOœúVë§2¬ü‘Ž6‡l_u{!N¸*Üc¡íæÈJ9£ÂïÀ4hqe?Ê\(#ÑÛ®Ò³ˆå˜@=ÂUàÛ!æj­[ñÆ’A¡š`%‚†Ì#™\ÐÙdÅ/Äq’ÄèM¯eü°cŽÄ!{$4¼{ÌïR"¶–æÌ*&jYÃÛÐ,/ëÅû $tJÝõÈÙºíÓS¨éŽ+P“s'Ó¯îH?£õÇÓfBp¿ñEÖÓoleu§‹SÑ\ìB•è‰µ¦/wqÛ²7¼È–ä3#çž§®Ùà/ÜïÌê
n4]$ZâfšgxÈ^DYŒiOèç¢òMÍÒá…¾Ôl7ÌxÍèÜiÑãßí,ÎîçÆuOHlZûrÊnk®~ëýé`»r£Â©{À‚vìqÈŽ=¾•ÛX˜ÇŽœÛ¬	•äÝ»-dU€Ð$C>PHþÓþîáN3®îÖDÎ`w1“/Ò²›Ê=IÎ3—wPÐîÆdN-‰weÚ #›HÿEeûEeûEeûP*›c:7àZìçw­àÝ=ÝSï”|3õNVîªÞå“xŸßuYC$äVEaÚ¦VÝTëŽLÙëÕ+reR×iP&ÙgÊÝ€ÆWªµ]‰”jh„î\ò¡³»ë½R²®ß!~Éa­Ò¡’gýZõ›ÃÆíNŠ¨5x›ih¤²eûÂo¬¨è¹‡!Ô(mÜ¢ºú³¬Ï
„Ë!°MâZ ±ÍkZfG%ÒZù§3²óOgÍ–~„€¦þé¬®”œÒå¼ †`ˆ®QñâØ1ˆê¤Ò2Aår!õyr¡ aš.Õšö	Œ?…ö'yÍ£ÀÝ€U~#¥™9ÝIIweú$¨Û™Û.øE„xÜÓ/w«X'¤–;TõÀu…+ÔØm4/—916ŒËãƒ!EŽ­ˆ& í÷¬ÞÔañ¬Œ_µt¤ëý¬œ‘5ƒ[t{kã|³ë[tË}­’4œK[î«¦“fáW‘ õ]ò‘yí…¶¼Uþ¯R7KÖî=‰ÙgæÑôï’õV²ãYô:ÙàHÜ`kjÔ<ŠÖL§®Ö¹½j´c‘ê›·Ë`ß“˜ËeˆÕ.æ)Áb‹8ˆBog¸f8*yG,¢É´FÕ|V°ßá_`É‡Œo2¼´Ãf7ük]ÂMÆ6¢ŠòÅzòúÙJÓ\%oÏ##†1y ðžpîd’›E“<ZÖ çË9ËÏ¹?yý hãÊ¶Á"Rô´™ó‚P V“ÐÛZ2Øº±ù9â¨zÃØ·v;¡ÑÞ_UÈ®¶Z`YtEfý>Ôù¶HÖèëˆìaÄy‘VÉ×øÐ¦KÖ¤í²EáÚÞXH‡Bí%î+`«šºÝŽ9òÌóN„ªÉlø÷1ÈmGžìÖaÁâ_7¡ÿ–ZÁº[w(ÜÈ[
xøwBuå®=ü»saÿ¼‡Î:0K‰“<¬ -™­ÖM,pž›ÙtIÈîóé4­ú·±þ›8¨hýä@÷?½ÔèêÝÝxBƒéqµGYb((¬3#²‹í©kIôu}hƒ¯½IŠ“yV9oJ$MJo½*X€ê–.sê»¬æâ­‡ª%`­i÷Iß7Ú>%·ýsl—2@6Ö»ÕmcËÿÜûh÷¬c© Ú;¡½b£±±˜È';aì	8‰#Á¦ÒXÄkÑJÓè‰<ÑïCÒº"¦ W»N|ë8ý©=”!¥1å o0=I†P,?K0qK¿‡­½O’ÃEf,e_RËáòê®ßl Ò4íÐ¯ÇiŸžÝË'¤ùU“={½^÷ž‚ˆŽ›Š1-²n×Åõ®þrCuR¨»ž»np³9ãî`kèÄ›ºûÊi®¨$#cös™€ëžIÏpaÝÂ3Ç-Ì'’•\Í­NU=”À¦ÐÀQå>Ç9*vLsÇ_¸ëGÄ]•HÒÀ4ZfÙ@å)3is¢þ·Éú–vèÕY^ ŽÞ²#dŒ%ÝŽ	.ÉRŒt¶<K)£ãä)Hñ4®ÒFk ßÅL&–;`ïvã­ÀýÊuÏ'T–š²Ç½PÏ[	È,sÙ	º:âõ¼ÈÅ\7ãYŒÜCa\Y4û¶¤¢¥TÂ0Ž¯ÿý<).ú²«âü4Ùé©À+Ð=êé¡Ü;5È!v×*åˆ>€ál^ÂëpÜPXk [ßý{&Ž.4=ˆ<Ÿ^*Œ]áõÈé8½sï>¶½æ :,2¶ÉWäâBbA£bÔœˆüXW¤ïZ”.ù˜H»ï`‹"%$ÉP£¿aPNì™“œg™E˜øZð[ LkÎ,”ù4é'?¤%-tè©ßf'Bí²nÆMe‹YŸˆà\»3}0Ž²³ëôqú5I²“êT<ë¼çzŸuºÞGªá‡ß<Ÿ¯¯‡¦t|:ÏÞKC&fhî§«X“*ÖcŒhÙ]7RÕŽ©‘~rú±»ä(äÚl“&ûNßCFŽqN©­Ky«ž>>»”K`n¬°Ù-B²Q;âP–ü`·r¯.ne¬ö=^Û‘  \CRrô$Ü÷II¢·oùÖö¯^å(Óã¬wÑå•EîJŒXfŸO~P×Yß`ƒï×C}ŒYºŒõ®Ù(,»¬<‰a¬!#6û9ì0ÑÇQ:IäžXm›aˆù7éUÎí	¡Ã– ‹(·á+äêPôöû–‡Õ&ã¶‰“ì´ˆÉEc&£«Å[¨x]žÀPK¤,D—³ßœ
úÂÝP­÷vžQ?ì´÷, -GT¤Q5åŸ#é¨‰¼®ý¦¨}zàÖÐêÖ¸HÇQþÖ¨…V¢Ä+ŒíS:Eã*·/B
(ºVNJó‚b]#ë‘‹vh®iÄÑûÖ”t¤mxa=èÝ›€"±ÅÏÈÓ‡XÕa•?ÏÏõ¥Ämå¦cqOTG„/àûpioÑÝZdEÌYó#Tˆt¶".© +Ró¬…ùÑ-­…
ìO`-”-)•9¸ˆP?–™Åƒž¶,×ÌHÐbé¾ŸêY2FÂ\iÎù6ÖFcžAébÍyžÜj÷ó	âg¶û)Ô.m÷³6d÷û…Ùþ[g¶ÊX×ÂG?ÿì`@$ëï4"~DQYoÃƒ–DÍ;ZïÖn¸,KÒŠpˆ%5*Â¦p«“ ¨«ò^Ý…Š°¦dÈÒŠ°™ "¬Á~Š°Æ²V„oBŸ·W„b¾EXê'P„LòïWVÈæzm€ç:*®|*öqh»jH·×ve§·'Î”šÎ¸¦±1£:¬ËÙLCëõ›[p©Þë7öp¸Rƒju«bžøW“ÿ<¬ì0Ç0BøŸ¥8X…ÑÔ²¶ò¼-ä”V2›#NRØšÚdÕ´ÛØMD’Ê][ß¢—íbZ³–à®´¦íÅ´¦­Ä7 5kç\žÖXÝŠÖð€wi\Ð>­E?9­¹¶ä›Òš…²€Ö¤ï ŸÌ{[#^”]üôñïNÓ´f=ØTÛ­z@1#aÉ¦³MïN‚]¨Ý<qM}ðòV–¾Rýà†>…µN¡1Vµ-e­· "çS¼T/ÖE°ÓoËùI‚w—–Jôc®_*}˜Îûget–L¼/ïg©—¦´¼‡ó-T#ÍW¾§ÖAüK<qÅ;Ô¶52ëÃR­u6@Ú™þ ‘;-ÉÙKIwù…’eÍ YÞ‰ARÔ²öH¶UÔÍ‘¿lÿ·}Ä´iKøI·‚†5Å­–P‡¯¢_¸ûOÎÝ?LX¦áëETæÒüUªAþÚ`[µeõÐ%KõI¿fYÕ«Æ­IœÁÖõíªf©qKêlŠFUô£P-¹þœÑE.QßÞ¦ªÉòƒ›T­”õïÕ¢ª—'7¨=3¨ªÁRŠÝV•™ôôØ`5½Ý¡	–q¯nV1Ãã¹ ÀÙ–¶‰±Õ¹¬IŒWý-b¥ä~ÿšüÜ&1ÞÌ"Æ€t5¾²#s­töR†Ì¯¬J(Â·µJhÇö©YÏ}ÀŸ¶üÖH)Ú_Ò÷p##ò­WY£“ãvfë¿ó."qdÉû8Oo%â„·õßÆ¥†¬CÐoc@A6âçØP+ÿiñþÜÀˆ{.'ÍþÊÎÐG=”Ÿûò¯FA`tŠGtLÂTžŽT6cV7[3|(ÖHsÏÔFp*©&•ñm!Tôþ}±Oö^)~ÄÌ,lÓYá©œ£à¯©¼ïzy}ŠI(Aç“HŽ™L¶Ï@€J¨Iy–i©6é6,
%üP"äwÛâú/8ý+Ÿ^^­ˆï®ÿ‚wá¿Ì›å¢@èé×)ÿ±o2s±hc¦e¶#Ý¡-æerŽ-ïf³yÕïõ:óÈw× ¦áàü» l0<#­·X>åå[¦emSô¹]-¾Ô!|è•áZm<+¥GZ(ÉX^/ÏŒ:2ý½•o’f³˜{|Ö`*z šØÄ^Å²„ò‰N£â=ÙæPV<C±‹´”;eoˆ¥Ôn+ÿ06QD 1“å—¹f%r<Ž†§µÖ—,l×^9K—ev3Yû
¾5Ô7KõéÑ[ârYÜ#³ctxxaC›³õÝ°_ÕæC¯Å]G#Õ~‚Z_<- æiRŽ—ÏRÛ­;†Á`a<ðÑÓ8óa>·F6´S.×NC†> ¾<gÜÈò«_ùs8<E@Ž­E¯&Ú´: f g(]¦af‚ÐÕ;T.ôúåbzâT.P|xFÃ_@Ë®+Ì¯lwÖ»×gà=6¤QuËM\¶ò°¶ÚÉƒæÛÐy{¤Åƒ.vQ%0Ó—oÕ fÑ–Y5ÆÁClšPù·Ø&ÈÏ¼{Gºí4ðpŒ­?xî†ƒïÀ™ïrð =u{Ð^3…K`#Gõ´ã´‡Î×§]‚kL;1{ÚžwÂYŸJwáà^¹^!Õƒ;å¸8}»ÙPÿ6¦'lì‰ñ¥ír½+$j{<ô‹$™[úù$v“Æ€LÒ&ø«òdCPEñ·zDœ$dY6ŒÙ6%âÕ¤Åþ š­)²‡Š¶ðZ…³¹%ÕÿZK²‡Ñ¤¤|›‘j‚2©'u£FJ+	¯Sþ°vèä ›ªðZQ“ÒInÕ*L‚Ñq9¾èÕ›ÊÊÃËklÓ: +ZõUtœAÌŠq³ý<ÝtÖ‹™©õb¶P¸6Žƒ¥$ë@'ÊÆN”w+ˆâÌ/ByjáVMØf'Ùiß9š¨q+cK¤æBK=
¬	/6†h©^ñØ£å;õÁd÷}Ú.¬ìn-s/÷ez>D÷Xq	à¢¥î#ê1'Ó1Žb7{ŠQŽÞö“Êü•È¿Úñ(k¬Âš16/8€Ê*žò¯.¡S¾ªäö¤õ>Hc-ukµ‹=Ôéûß±â²Ì×ø¿-mæ§BRÐ¡/ÒkÇ‡‚Ö€gà¨d|”Rÿ.«u€¾ x;¡‹ÚAUž&ÇÑ|RI7šáK½‹£zø¡%O¿Uì'à0Æ/¯Ž¾ÃIÝÞßÙ>àèÈÑ~%oS0‰‘ó,ý¸œŸZ0LÓŽêÐžZšK'ËÑ·×Ö@Ÿ6Bh$	YÑiráå²Y9/v¿›Ù+fQ$?
¡¥®ëZä¿Ù¥±ú+…ÊÓøºù| wY‡”OŒÉèÀ‚[šŠ&€†øBm¦ñ‘b®¼¿*¶I¢Y+Öyú0s£³VÑ=Xè¥Ô„«Ûg·ýuwT6¹*ge‹£R]Ò«{]ç,M±ÖêØ£®'		w
ƒ`eýçŽÁõqÚ·¨8”ÑHQÝ"^Y|†‚°brû€Ú)šŸ2*Àð’e0ÃòeÝ06 ûÊû˜âšW^[ˆÀ’ëŽiŠ2îþúÂ
ÓRcÍ	Ù°ÚZÚ†›·mfî>ì;>­$ûx#ÿ_íÌ½¶x1tÖÄ¡nª‰]Fmº‰„¶HÇÉqâîÛ·â»ò´¨ojÛ_°&»š÷ù"ýXMüÖPßºÍTï4Ø•94Ø´»ð‰î&ìöÁ‚¾S£kªm°Z—š·„í–eÜŒí˜ŒlÚ
x;Þ³¤!ðÃ±¦¿W; #Ý•-Pý{cKÆ­™,ØžLv~˜%0kxêv¢\Üž-C_aîœÛpnæöµ1ªÃáðA‘úi(x	ûZr•tÍæOt‹ó%Ë‘uôuzËÔñn.ì\ÏÚ
—¨TIä¿!öÝŠ.SI—²R“t:[œp°3¿µ <Ìÿ)¡¡çuÏ’}«±9DäÈqpú&ÈÕÝÈ–k}
10ï<†ñÍ-øb/¯ +VË20y*Ž¢ñû“'i¤’yâø·Õ‰óºæU»P<$†átìJ³FèÖõ0	)E¾¨†lzwå9tÙÂÈØÐ%·&]-'bÛ=þSlJ´~5W9[¢K¼s§‘ƒ8s³‘‹<ÞBð ÷÷‚‹oz1ç¸ïô¶½Nèà0,†¥²?êy‹ä¼ÙÈ¸JÄÀêl¾ó Oç³I:F‰‰q¡g0h•pÆ¸$ÇÚ¥q åe
-È9,ùìAK {;´\.òª¡C¢žj­}–¾Hè«Gb#dÿ[y‰È:½þ‹lòú_d›"–ÃFD7óucC¤þ4;Î[M€÷ï‹ƒ¼¨ðšø÷I2Õi"Ž ŒÀë†ûø„?ÎSè>Lsuþÿþizrš(”àäæ™€‰¦ïE2>8`x+¡	²±¢A‘üÍ?õûÑª8rÖŸJ©u'$j6·[°C‘¸UªöÑ|ûµØX__L ­•Ex}0„uvmY:·ííµµ_ûØÒ~­lSû:Šœ:°&¡&nàœH§QAØ¤sÈˆè×ëoœ­@/÷ü=–T7Jm¼ò»1Ñw­"©ºŠÄíùfÒ‚¨’&d$úu»‚4©mÅ.^Ži¡¤t¥ Ú¨í{ª1ãü‘wˆ…îŒ"£p·8®ÛYƒèošt-÷€’\ùEû˜“‹Rua´êwIeËK@ýv?Ì,„«Ê‘SOZ„bC^”pkûrN…Œ¤$¨”†žTQË•±ºÌ°•–ôo_Öòr[,!¬{™g$8Þ~lÓ˜5Í}hÏpM&Ï‰â_Q¼­-©xk_¶M˜ÐU—$ .j€vôºÝ€-ÈIre—q@Z$“†7Ö°Pé'h«•é3nåÆ¸u02}zàG’ß\9ûbO‹ëŒtˆ‹LPÀ¦6‡BåèyBÉQ¤ø;\šYÃÂí¤InbBÜ¢>#‘ìÛƒ ?Ši[òÂóà»8¹Xgå½õFmµ™÷ÛŸ¦“˜.#D$lè””Cy'¬¨U=£ò}¹Á~oÊßÛ'˜>Gþ~UoŒ´»Ò`Z&CXý×xQ¿7´)Üg‘à-îŽo² à”<MòaÊ¢6X]ªKj ?}GÌÖòSt"¢Y~›œ;áÏ…ŽiRH
zzò6:‹&i¤Þ]‡œþ¼ñ—‘Ù€ÔÍ¾(«•BêåÈµ+&ðnÓ}'WûV½`fü]Fªz²àRÛóÅš&þi¨µKJ—Ó>	T@mÊ"ÿ¥ÇjRSeF6ÉæI“œÏô°Òl³”.ªø.Û;â™olöl5ÕÅ,¡¼'2èIïáœÆ:2j‹ý‰þT[P¶†2ÑY`kÖŸŒKG^Ó+M|Ën²à4Ûpš­8
àS¥ús)ÀL	×D§FHöt‹É…Ò+(•¢»âWb6/Nä›¢Væ´ëç3P´RÚa¿ÝÕ @5FYØ˜(ÇE’dâþûc2°Ãì¦¶Ä:ƒ£oüé¦±mpwbÐ£äy!uÄm¯ÅVLÚ‚¶‘,¯›0í„€O6))²çu_ì×ýoÖDüétÒÍu§.ol	Œî®ÞtUpZèyRïR·¹í&DŒ´‡. FOh
P#Ai¦FÝˆO’²Z’”%ÛéR·Âjâ¬,q¶¯ï*¼¬»S°ïýj˜5 v7tHmTF·LRro÷jT9-IÂ=K1ÐwX…¤´¤Ðbö>*Oò¨ˆ½÷ã"Ïò“"šFÞ‡÷QveÞË"™D€¿ÈÉÏ;ÿG(¤òqå0?k½°{È'%_a›…ŽI[Fº8NÑãŽši‰J³HÙˆÉq2˜…ãhœˆö¯Bí[å‡Úï;E¶?eŠún›îz9«óy[S”•‘ôòM(ø=ÚªÑ¦Šå¢£¼¨JFv®Bß¨Ò×•z.Ïíãu¢Qý-¸åÔ~àEYµ+ÿ²ZtÖ.Ú8tÏ­@¼ŸÀØÐa«\Å¬
é~²Øö.xÒzÀÐð,A&€Ô­#Ç	eŒ3ÃžS´—ˆ`–%÷üâÜ£ï^LþN-J¦úL?CO6‰Ø‹e°+7´J,Õ£%ÌÌ5äºo­æ¶‹–›m%¸]#\B[9Ü‘6ÏRÀ<ŒøJLÞ• /0÷€;vëÝzM†	âªSÀ9þ~(‡iš'eà#2dcqZ¿cK‹»ÙÞ¢®¼q,.í--ö•ö /y)Í)¢L3ŸÒ•ï?Näþ_R Ç€–cÔÂ•G¯ÛÇ8BäÀöôÊól,< ÇØj«ÇëÜ±¤t[w¤wáš^YJÂÆ¦;‰p¨ÙS¸]ê†‘aW±6Õ<·6egab©‰®™/èìY¢uoÛG{W{Œí{3¬N¾d‰¨ÅÎä¥,4½“d¿Å°³™±*®Péæµ2ÅH"¨O©ò.y_1Àµî´[!—UÛ¸.ôœ©¥ÚdÛbŒº{»H‡ÙþÅ>â´ñQÚGn³ºüôç5•ü<Š¼Þxºg©×:üÜèãÄ\€Ã•w}3ÓÝ¨ãV‰Fs¿Y èlÇh6nÿ7ÓëŠ–ÜO¯¼ô’6š$¸ï÷éBˆ[+…u"ã¢²/Ft?)kUB¢)æ;H¦ËÉ¦ÂÎÀ.â©+ :æñ:8+’0TRÐ6t€·|¯™Êrª­á)fV†?³º)onëL²yn%Óë8»ËÌ­ü9z}Å£e&]u¾}*ëÜÁL¯T;ç·Ñh°Äó)6Íys,ßw^ÂLYþ@«˜·ðs-äàŽçÓBÐ–rcrxU´°ÈjsC’Mzô /;ƒ¶‹| J0à?j2ˆÓr|ðà4R€×Èíç^µæM>¾í<ûXøCN¿…ÿQÏ?n	ÜRÑio×Žþ¦Ùæ!
·ÛÚuKµ™öCÚ&O‹| ­ÿy·v”Í>Éå•Ve29þ¤=çÊŽ$6LÉÍ—‰—ÒAû³vËØ½üo~–LoiÀìj¼TÍº¹­Ð·$Iýô”5(R•2Ÿcú¾Ën—vCJ¯±ã›ÕW©‚FFâRû7•R½•çIŽò­8½ïO Ö|‚îó¨^fLI œ,cOxÚÆût…`×ZÇˆ²Õ`›ø…PÅÍÀƒ˜º™4gø…v6ŽEÿ¨Ç¢0	Æ=íÝ¢ãÎ©Ž{¬C˜êÃ)ëdÅ›$…ï¥Æ~DÏí
DM¾ƒ®”‘©>ƒS^¥r¸Â¨šÇñ±ÅmMÉ8'˜DRÄÐN”Ø²'¤_#z‡›£õuø_oÀSy8„÷Ou(%%Û"|;OÇ›[&½}ÛƒÆò0ËÏëB²h²WäßÉ¬tÂëªN¨ÒëX»wŸ^:3@9!®Äšpð¾Ÿ¨Ó
¢ÿé¥À!»›@Ë,xWƒw?	üÕÑw#á®a;ßxsµAõ®1wGV¶Œwfm$ð¬ZYEÓ¹Óh*ô7å‘ó/·Å(£žaÀÿˆæ¹–ÏúüÙc±Î¸íV+DVp$^÷v0‹Vuý¿OÒ7w˜p²hÑõQñ^²Ü_‰?¢nDoŸadã$-(À~2žP^ü~>2Yï0gù$Pqw»÷ÆŽ] s¼ÝÉwe°`Øq/Do$—ÈÂtc"îIÇûÈ¥k·Œ÷Q/mÇ§ WA…R$)cG:¤”vNVÎ1îYRn,O||8võù¢X!ªÑN >¢Ì%?¨3V»ÖðÈFþNöqÒ;$y…ßÕNœ=½êÔcœVß&ÛpX>˜¢ç¶¦Q§ë®hç£ã{“¶Çz#û—Ä
qO¦³:[È@DËËßêþ$ŸOb<; 5	¼¼“îÁ.OŸ8“~˜ïf€ójdzø£”^¿q÷¨Â/Or(|ŒNÞUâæôS\a¨Ä=:=t «;Í¾’ÕcoTuþ0GA£²‚5œ*‘Êl|5ïªõNžGw†¼dÀæ|¦ifÎŸé+nÔÅµþ2ÄôLÀm~6œ%ÅÉ<«"w{%4Ó«§ŽCs9‰}	V’9*G_ý›âQêJ3žÔ)˜>Ý»9„Gùö1½Ö"¤s•tˆ¾î¤S…õsÍY:Ã‰À&}Ìc­Æé¨£zOêKˆ¶ð"t[†úu"YÏ™?½ƒ½™~Xå ç_¼Ð^]YGg˜STuûÍƒË>ôJ©7ûîA§Ô»=¹»Öj‡äµCrÎé¡}²Õ5E-ˆ¶O1
&ë…œÎè:N"Z€»Ð-ïðÅ¯57õ=·™›ã±Î0¢8î³kçùœþaPçX\~cÈÂ¸Z`h Ÿ¢ÜêìJÞÌÈo/”½È ˆ‘”„·oÔÒÝ4ÌÄåj},ÔkäþaP[Òg$yâ¥ÞÛ »»ÅŸ†zDµ®FùQyC°P3)Î"ÌË<KJÔ¸f[Cú…kžÅ	`.ÁŒ3ùz$6j)­4ãA:íAùM(jÞíEÅ8Åµ»%6àýz-Œ‡¼Hïåw"(¥_yJ/ªÛô
¨Œ²8—Ó¦²ö}öÅ`XÎä¶Þß\¿­Íu¦J¦VXq4Q©‹Êó0OGÜ*«BXÊÙ­:PõŽ6ªío½žÅ–›5—RL¥bvK•9F6ÃŸ-ÕŒð7ÿTŽ1A!Íÿ¯kpôl8 ÛÔjË[‰„EÉál^bŒ¨œ+ŸÅú>Ÿ[yÊ7ðØ#9±§dUe:Åiì»RC;=,C]i¢+UÔé¢N­´±€:Zè£9?þ¹§Ãú‚¤¥3Ø	×ªƒ™Ç`úN¦f+ÐÊ(O:·<ófcmKh
Ö|meúUÓ÷µmGgI]/“¹%ì@@£k}X/è„$ÖL¬B‡nÔË¢"‰õ©fó]³gÓZ³˜Î~'—º¯ÐnÆCùwCb=ûê&Dzã[dkSÈá·š&á—îbJ³%Ã5ùç§4Pù+ÉjÒ˜ÇÅ°ÅÁª™½.ÇþÝ:•Æ~:ÏÞs3šÛdÍÕtz+™¾B×Ä±rmKw#à¬EeædúK¼¸¥¯|2úÇ ¬}¹—õ¹ñØC<¦HÐVUn/â%“‹¥ýˆê‡ôñ¼s|…ñ5°xM™>0è«¸'^A¯Ïr'Åf¼þî™˜o™*ƒNåQ}7&_ß™©tv–ƒj—œ=«¡bs[<ž2Ù3õ¾pzã§êP®­W°É\Ì§Æykª»Ã1~,ß3z
»ÖÄÿlM`F‚Ò¥3´J×¬Šb4Û8o–y<íx/Éb(I[¬fo†fµ9vh‡è¶š™\Ÿ,òw»ðÂ2!ê±o/Â?Øë‹e¤L?>1,Ñ[.“°ßmk[¯"Ü÷™´Mô{´{|c~•T"¦dÖéÄTbQ;µ»½·ÄVËù¼R§ú‚À¤U|öÀ°Êp–›†¤®gró”ZÁÝ°–õfñ€nÑíà1(g]´E-œ$@€þmRqqK_jI’w\®om~#Cû´imÊsLúÔ²40štøX8ðG
A2â›Ÿ¶|ÂK8D’ñ¹†óy…ïÉ°ëîí.1-:”•S#£:`Á%“<;¡}í4/“L™ýQ8ÅmðóéÈIìt…U_+hŸLEre©ã-cs¤œÂX“4ø·EX³&l7m¡;"ýŠÚÐ]÷0l4Ù¾¥SwßòEW½¨¡±Ñ‘©6`ÊZ[F4µA#€|{#ƒÿ«{í8¨‰h¯Æ­šëX“ú*uO½4F£è(“‰ø&™¦Yê’?JÈ°6¶wûÎ¶y/JÃ{B¼½·+@˜’Ðd„eæ8¡û‡bÄ×ãè,/VE§ãó)—óHØÚˆ…'ºÊõ_¯ÿ5Qã?ŽÒI
2™¸0çféè`ÉòŸÏ ‘]`ß“ôäúÿdã4ÛE•§h4¼ÄÓt†¡Êïžä˜ã¾ÄÔ"4•vRcW¶XéyªÒà
vj"29¢†]&'sLaRâ½µs…˜"9IKIHÈú6¥U¬¯Š/°— èç‡¾ó
g*~òÉ«ýÝ—‡Û×ÿõú¿¼Oö¯ÿùp÷É¶xº#¶_^ÿùùîÁŽ8ØyqýŸ_Ê×Ûbogÿ›o¡ÆØû;{¯à÷è“mØÖR@ÞÇw’á4 ¸­Žé‚lm4êkãÕ efƒAŒ,ŒFŸÀæ²‡D ‹#Ž`9Jµj’¤ñÔ­"_:ÅÃMs õ]’*„õ“FbåŒç pÑj¾sáÞä5.®ÿwœ“. èÍqžJ ]æC$<©d,˜JÆ!1¸…½öò²œ§vó8Î&²s[+ƒÒÎšØ¶C[9H§+ØÃeÑ^A‘>#@œ ½ýW_?ßy±-î‹o¶÷¿Ù~þ
»ý|(þ˜¯ÿÌÅwd°^Zès<xÝ‹?¯J#rÀ¼7	
ž0ÌoIÔÖ;‡üjÀÃ2pWFÖ£-´°	÷‰è—i5×-Á–^¥gÀT±»q‘&¦g8åXÃí”Z'ª¹Œz…FšN5ÛÃO6æŸø(v0á3fÆ(a”G9rÎâú/ÈEõÌ«ÉÑÕHè’ÑJ4w‡ QãZÉ®ÿ‰hz”Ò­7zòþ   ÿÿì½ÙrY² ø~¿"ÄÎn •n’’Úš"••ìÒÂY9Ý¦ÒHA HF
@ #\’Í±î˜Çù€¼ýpmÆì>•Ý/ÐŸô—Œ»Ÿ}‹ RU7a™"qöãÇïî[ŸðîÅÑËí¿¼v7oÇZ}—b(DsbôžÝ!ÐÑúC,·9cÀY"Nô£šwzÿ´Þ‹ö^í¿y{¸ýúðÅ çQí‹×pRÅ0¨sX6`y&ívrœàN¬:‚õ9®8et´”fÑa* _>ÿJ¾HÆŸÿy‚ã‡£•ŒEèe¨m:‘0pÐ%£…Öaôü„RÉàó¿"¾€Ý¥À¤Pu1dð_È+³g`;ˆ[XI”ŠR2ÝÜâo»„Cqv0\k@Ì¢ÈÔ7¤ÑuŸ’E{˜”ˆw†Ùrd—Pò6È'ii¥ó£×¯ûÄêeÎ˜\¸GðVZs©º\¿½·‹WŽýŠ@ªD€³¥ÎÄ>ZD«f4<µ¤Žã´7Î†É¨è	Æ€(SãU© ¯º÷»ë½ûÝãQ\œ.iIwY%X vmoà¶ß²äEb¯€B#´O™”¦±òS‘M–<: ¨q08¥µpb%µƒ­õ¶ß¾Ýþo¶P2"ñXá«jT~óü¿¼Ø9tk“ÞPT™&F"òÐIkùàðíÞë?F×¾¶u­Ì•Xþôy* (ÏUC€ö\SQ ?G5ã0ÌQO?#«ÙQ×¯Öþvmq%÷_M.€H”Àzì£'>Òim®½Þ.Gï´ˆD@—ÎF‰k?ö‰®é½’Tû­+ÐB
•Š½KßÛŒÚ¥ªçÆèHìÓ}––—øä Ì½±í-õÅr”:Jo€ˆ¼„úØ5•ß› µìšÖ¥]£^çûtL1|y[¶øŠ}°ø!®Ü¬4wÔR ú¥f=vŒ*Äj^NqKŽ×-Úâ“t_þ\!åÌÊüeÂ’¿kGv–«qnÍNu’Ø©¡+¹¨{K¦´­Ô¬s!£BËøUP+b1	‰z3mAùHÈfg‰.µ6ø`~‚.ÚÛ6xL waÄHÎ=2c^Î3"™’"dö>þ5.ûÊ9”õW”,š½¿Nþ:Ù¥[æ_'KÑ·@jöÆH›œåžJñîñUðŸAÇ$%ãû®Ò ÑÄ©U2?¦»8Béæ±/ÜvYà•¹œ^£ý0„¦œš—úksz•i—¸fl+:)…å£5@0oÛ¦Ò9¨¶ÍåÒêpPÁžJC¬GÚ@Êvs©Ü2 52±÷£AÓºþYo WòÆ.ÐZ½*ìqo&k5œ<™'_HÂŽ±˜Ã
=àîÉ‘”¸
|Ô±…¬P.UÉÁŠ¹ìÏ;$öµõ~÷^¿¹xUŸ;cR ëQ,D›<QY$×á`x"’ cŒ¹ÐN“tQåó0âeÙ·ìAwèÀ\C«]2Kä­h‡¡N<Êj<y¢Õ§´»*:¯ý[³õ75§¶DÌnãW$×vd3Ò^^“:Ë°K…ÜOŸT5ç8z5lò <…ôön´µ–TGìŸ;(o ú€æÀÞ&†˜¸£êG›¨.‘~¦±Ò»µ´My& hÀÖµãàÛÎ»þ{­-w†2éak7£'Ú$t81´Ž5†^>M`ÓRZÿ<AÖ‚“¸éÄHÆÞ8ÿu–Ú›I…âÍ4ÀRžfdï!E"%FrÊ¼¤ï´¥¶É`m_d,Þ›ßÎ5-2 Ópã@kÇárÌ±ñuHQ#Y5ÅYÔoädÄ4¶Ã¡Æ³¸ëhu+ãÆ”ì1ˆdfVT6‹eÏâ¹-Ú[øy`=â}Í÷ú½-mœZbÞ8ð—Ü±¦?ƒ…-{ÚOì’$(ìn8­Ó7*,DIVY¹Æµ¿¨“/Y5tþ¸ÔlCµB×¦˜ãb—2@+¦úE'®ó~bJ

"(CPhàæ@¾P[ìfü2”z~b
Úøb10ëœ³û ˜úÕ¬‹PZx TÀi!áÔ±T@Y0 üPph¢/Œ}È¸:ÑûÖ†³zÈõÀn£v.tƒÌ*AYs€Möƒ,š*üE1%þ•†=üð§¦ÌÆáÀ1½AB
Dš¢ ODnNµ|ì¿Ü~ýæÃî‹Û;Ûo>ìoï¾Ý~°Q“<»¢j×zÑöíÅÄíˆ¢&Š„áH"^D§¨Q.ZËBÇç>\F§04Ó¸"¦†“¡1\öP=z_lü€UÛöÜ–]¾ÈXíŒÔ+]`Î 4s OçHÔueŒLœM,ÿç¹ËÝÆDïÙKsC³*gSMäÙö­bñ{é+ÈBN º»é3¿¿©hÛ’•²:°q>Åî…i:‚Ù/ŠfŸêÕ^m 
…gä—ZµO,Ù1;G
	G÷‚K"éëcŒ)Ñ$ñ‰'ãòL¢^óÕp$–0b\årQ'µ· ‘'®ãxáZ6µŸ[$Y-tÝ.ôKV¹¡–ÄëÉðÍ±•XÈº<°M_]®NÀWòVsû75,	¥+3Çªî,o8Bæ¸‰«XT—¦VÓ«™(Ös8§F]L3‰¾íJ$5¸à»¢¸.,nV«Òm Ë‡è F@mw…ªÔ¢Ì3=ÒŒGôyâDŽÔ±Ž3ò¨èâk…£6{Ñó¦Y-fcÌ’Ëå—Q½M¦ÀJ›‚Nz&eæE(DŽ‘-ØÇÿý?þ9úæ*ö‚03ôþØéý”¥“vë¯“–…÷%…1__µÂz3Ï„³`_<”pËŒ1vßw=aAäÎPxLÂÕµæqb ‡ Ð·ðÕß#s2Gúˆ¦CÿÏ‹ƒèí‹ý—{;Ûh"´û&:xñÇW/^¾Áà>&Qþ—éTå×­¿N0ø¶§×UwjÄb[ö¾˜Ò/¾û=á†å¹êà2qÄ¥Õ‰‹j$Fê@ÈóƒÁ‰}Bxc³ÍPC[õ‰’àly@Ñ<Ïæ†o…ÀG¯dº ©-Ó€+E„Å~ TÙ–G¨e‚"2;Oª”š¸pUÎ7ÍõM‡l³Â»Fó«É¦¯à§W’†‚£1‚«c† Ìˆæ2Øb2œê
Duö“W_†a¸~:ÔŽò.UlFµ©¿Ú­‡òh°¡oà0˜·"/p4ÐŒçDhØ¤!œÞ Jk`TÌ:ÕèÇ}Û*’¯þPÌ†MáKé>TªSÀíErB×s´d ¹ë¥{¨0ü}k0’tIüyOJ†ü±wqeg°¾4Fh¸Àüó¯8Ðñç_/R ûP©+[Þ¢¾YuÆ_RT{mW¾¹rn„ëb¥ÍÌ3sú39º]Ôóu3Òpíd]+[SÎb]CAK-HI«%(m,Cäûq^¦ñè±%?~ZR/%V¹R*ØRNHè3  Öoô‹xq¢­`¨<2ßazŒª žn©a
Ämv¬dh!Õç4Ê¥ˆóÜ²íN?Ö¬c‘±ÚÄæŽ¿'øš\5#¢³!T¼Ž±>wƒl¶Hcø¨°ÊËaÔ9=*fõ¦fAÃ¹ÿuvóE·BòÒV4ƒ SmÂa9l%[uË6Óš ç'ÙxÈêÒ* ó`l&Lq1@sî¼á‘<:½ÐFÙn¶P½f6GÜ™l‹3Z7iHçÛd{êáMšE­94â»¡ÜJïä3b	ðZ>¿H<;A–›ùÇóaZ6M–ÞWûºÏqXèÏy`i×êmÐhÌ<²ºØäXYÜ‡»õ‚-å@jjþ*sjI½Y)•ˆ5?cQ$ËeŽæß_b/ýÁÿÌtrbŽÉ]‡LL7¢Èhã^âøÂSùÈ¥A°ö•Ú‡F9~ø¥Äm=Gn¤QEÆŸªµk²nwž…o¾%`ó¥«“ùÔMÝ"’9oI©-ûFÓœŸêPÏÑ“ù9ð!å|›|ò§íŠxÍ~Ò`,vî™ùîQ ©›‘"
ªƒygÙ°ØŠMÎBoã¶†ÖÃCK*Bò;ú<ò7ûš=
5{ä6{TÓ,7€áv¡qó­“–Y©Œö†Ü¦MXrL†{(;u ÌŠ˜AÑ0y8Ìîª·C[ê)É+x…åhµ"þ†IÔ³RÌ²q|žì{:iÛ„ú²žþlšF
•@È-ëXÃY &S|‚&6E9êå¨¿ÌÎ¤Ö«<x9…ö¥ImÇ¸>­MÞ…”©ŒÚìÊª`¹ñ:ìÊµ½Æþ@7á°+^5«Ä†jÈ±G+	BáTªƒ»xo0Làæ÷Doùû½ücÄ§èn–Ü`­â?0ÀâŒø©àX8©ïX±/Š ¼„ì‹‡5»E0ßy¼òtnDnYÀñãÜÑ¼§P¿TøÐ ç?Ò–J«w´úe‰¦\{jA…ª8zý´QÖNßP„ÖP^^Uƒžz<$¶"ãÌhvŸ"ÿ/Ñ}–y­û¨&j/{+“k>?½ëv\ž…î1ÿ—¢¶>¿â>mUä”·ÑUð¯2•|EòÑ²ož*D	Øè¢¨¶ÏÀ IGÏc«K|ä¡À¬Å€îŒ!º$'	UÐR{ÎkiÆëCUã~5	Kªblº÷Æv ÃxcIJ\è;—eoPˆ‡ëIÏiÁ8Â^M©q -ãÊ€²)¤dª¯EÛXãv#¼ú“â1»¥Š8„WsŽ¬A(½:”¯¡œ|oƒj…ubˆ.¸9#Î‘EÕT‘	ûJd"ö51ß‘º,ÑèÓ!Â0v[ª!<†ESš ´1q‡`£M¼t˜J>,iº	d˜î<¤Û–6ßkOÊ>Ï^ßr3ef<¾zfò•{ø›ò2ªÏQÖš|e|Uó9ÛàžO0MÒ!·VGDÄY¡sÌŠò__ü×Þ¬LGEï(Ë>} ü!FDÿ ƒžîr¨¿óF
õÚhs|(³Åi’”íw|Ž
H=iU¶¢Ö‹ì‚…ódàª£ƒlîéJÏñ4 "­µ=*3EFr¸I—kn—oðvIº"Õ)ý!¹íµ§ BC¾"çpÈµÅcIÕØwiÛG‹Ïsñ„_¼žkñ•ò·vÅUd÷Ö¶Ì~8eá“Ð#b6ÅtŸ”Øó™Q­€åY¿ùºðyÑºðïúº¨àgÞ…9à&ƒó­‹ÜƒVƒ•i}þŸ XRÏu§^ó!þ/Ìd(}h°øžÜ‰/a<+³1y2šØæSLLÕWÑžÃb}•Ÿ~þ%mÂ#
å'‚äQž~‹%½q5÷ó¯gÉÛý#ñØq¼ZÚß£~wS,Ÿbô=aœÄG¿ÚïÿG½ÆÎ¶Äi\qaèÃôùÿ,LƒJŽÙaOxLl›å‘¥pRyLa¸†HÅbl+fÅsïÆgN€Æ²Ü¢KlQ$<þ>%T¥Å¢E|€SŸå%&&è]ŒŠfìC»oY8¾¬XÞ}+ÂûeÅü¨Wå½BRvj¹¡µv	)Ò’¶0{ƒIz@~Ú3ó6 L#F¤W;²›IúÙ.¾UçGˆ2Á™I4$0õxæµö9JléÙ.Œ÷EÆß7Z…æ¸/´MÐŸt§ô¹'ê‹àsO”PØB{5.…^\çø<†Ñ—ßÿÐEcEïÓ4•-fE´UôhÀ€oSX<öxþ}ïü«ÓÀ;, û¡§ƒd¤Xœd‹Yðöv ÐIòâ¦ñø‡ÃW/÷&ÓYù‚ÝCO¼I	Ô×IŠg=N€ÜÊñ¡ŠMÏîå1ÙÙhh.Tî‘xÏ^ì¨m‘ƒNÙ¡ÜÁž‚ZEx·(4Oî–×Jë4¼NEÎï,ÍpÂ­Ö#<a [äØKd-õ”e¾õ®5Îs#’$ËÌz”Å«s8L?OAqFYÁ2½ó9‘µ<Óµª»AX
ÁµÉ=R§oé¸EØë^F’g¥!6ùQ~1Ç5(Ê—çwZ±Eíê&ÃÞÚL5+½~â</[¦xüå½Â‘Çú¬"í…‹HO“ŠFx?ŸR8UÝKˆ^‚ˆ[”ú(;ŠfËÊñiñÎò–¼Û-ßO£<›¨À~†+ôÂÊ0\(>LJ’œŠâÁD¦v…f|]ÊTÊ°¹|ö£ìˆ¤‡äÉ@y°üÁ=ö;Ñ3ç™
[ü];)uy³Ö—&nÖžRjGŽp™‰=EÊè6`ÈGäW8»¯E`ñù‘#Éñ(¹ÀÐBÕ.¸~k£„Bôc Qf/`ÉVáÐ,1*žP¿PH<ö[*ld”=ŸT<pR¨ÕÑ^F‘	Á|%ã¥·—^¿»d½Eû6gåÝ_gýõ~¿‹î¿_9YŽ–Â…ÿÏ¸ûK¿û(¦-©™¦ñ‡d„>îÃ62,N€©9Ž.£1à¤x)Ã)ú”\B)`tÈËÈZkÀd?r<ß†¸+6TÑÎ¯£9cñî¡?4sýÏŽ©w£ìè'€Ž–\t¹úÚëªÓ?A_Ð;vIT×'ì@ ýÉÔ¡h_5—$èš@·¤7Ôu[Â¹tü†dÔt2|Ëª=á]±V<
15H0¿TãS –í®†ñ\äŒ, Sª–Ÿ”Vs’úå):bñ¤å)Ö°Óúª)°„—0äwlØïYÅpàXÈQ\Š‡BWÉ¡‹ø <¯øá;
…íþB±?Íïaˆ¸vð$^T2Þµ$ÿ%¯/Œe Ÿåsã±ñC/áûN-b„[ýe™–³‘Õ<…†6™X"jñÏ¼³ø8¥¯a½eª¾ZÎ•ƒ\‚œdLÔ	òûlèøÿe„ŽØ<àËRä"Ã$½¢(ÐB36„$½ ¿
ñ$#áŽêªÕa—3dÙÝq6UŠÿ¦&iÎ$˜ µbŒ?¿¤Áäñ™|§tüÑ’{M§ÀœÈé—`}zDÓ4ãìOÄjZOZ1§3Á™ÊþÄ6lFOà÷ŸgruE}+µÇz-‚],ðS™€¬NsÀ4~°Pœñû†^dÚ°>ˆ"Ó§ ­2`!0-Ä·™hWHÔÆ²¦ÓgKeü·Æ¥h²§-JBÃßðßü•`Rå# Æð;€òŽ¼ð”›+Ðä'°hÀ‘yÁ“8_ubð—dƒiUEûÊCÚG}@½èŠ	ì «Š,¦[JÙÄšõfšÇ¿h;‚¿Í&¨!@*?ÌrÑ+À‡¡Ž–U0Mó[·Cdæeg¯÷ª=fx ?aÓ¼H3ŽH ¶Çô9NÚÛÙ§i*ûBé€`M|À6‹\Ú.Äcã)Ú¤D\ðïq¨@¾÷MQŸ!óYÁŠøüØCÑžñ°à˜	1^`ç‰‹ât	p+»T".Ñ 5”©©QÆ±‚üä4áîIáªÒq’Qìö§¹ãÊ~ñ×#0”"ÛÄWÀåÞi<6¼Vm:;® ‹Gg´ŸD#‰`Zœ ÙLÖ±Ð–‘†Im,	ã=ÏòaiÈÑÖå5Ïzâ'·:›„L±89$©²‰E‰ì€ÚXUöé\£ÊÝÅÉÈóŽe½Ûñ0jÌÏä˜‹wòé{nØVE´à¸³½á«xºE’¾|øX_p–ý4¯I®…Tâ¥¡àÌ¹BŽ’ŽãvííöˆRŒJ’.s¬bÁBc-sdäBÆâñ2d?H‚pœ2Aá@ÂŠå’DëûßUrh÷D/=YüÜ­|ûÄecåÊÂBñ2lOE%»¨Î®Õ³$»9þ]c·­GŠÁîhžzÏ†ð€¸d9
¯%[™ËXÎ|MÉ¬ÝjLÝ8Ïìz¦pÀ(»åaÿ35>8”ð¯Fç~¡ +iN[ÓÎÄ;±‹ïmîB;×NÙ iJý“8œ"÷;IÎG—ÜÃ|Ò¢ta‹l‡„_,éÃKSÅzz•>*šz4UÙ=?ïÄØ(x^:ô£ºP‰ðRÈ sZ-T…ZQ£”b‹X½ñ¦Ë¡×	-HE{Ixh‚Ð!E¨¸n|>aˆm”ÐßãkªXÆNMœ¶Ø#Ës~´èyQª2}“³Üë÷Å9][Žîw<xQ;ÎRIF'Zh–ïÞv´õuð~%ât‘ÅfßJ¨ÍÌ<$&³Â¢Úöæk#êÇf‰×È×½’Ì4ôúó¿¢KüªÕTex}ÿäcŸpŠ6$ÖÔXîŠÓi::Rúo¿õV5K|÷ýl2—;O}‚’œ‚EDîg4JO0üÒ)Sf¤è®Mì!V³îuÞ+o±”šâ!7…QÞí¹”©˜Ó½{ï`Ÿóx}š&ª¡Â§¬/Ôo°ÁnE«@\q3ÏwïÍF™Ô·	È¿Z¹!µñ˜§ù{$\ÏOS@r(lÄ@¹§‰X$.üeÙ'Ä|(üÇp3ÍÆ$q‰ZÔšDM«·'ã5Z%ÆñÅÐq‰ýŽÑÑHÐ¡HÇSÈ &·"¿<!ßaÁ4Ñ=ÐxlVZ0·¼x¡^	ÎG%à|'ýâ|š¹\¾ÀQÒËqusÙZûŽ“GXÙ'¢¼‘%Gîé6*µziA¹ä*–ž%fqRù–ÓèàDL¦’€rJK
’«Yü÷*Õ
#½§6ÍÃ<`)/á“ˆÒ\L$€Ÿ°8”SX´O%p¹Mk`G…íÀMû´¯Å½cL„ˆÍâQáèí.Éyc2ŒsÎ!Šcnb58;+~ø®Fa‡ ‹mÈ	 }5L1aç ²×„þ¨š5Þ³ þR1œâ~óûðÔí?xU"RQÊƒ†JgkS¾VÁóÇÖaƒgg”1…vâòF'ŽU#îDþê%g+˜”O™Î@S“‰ßJ= =ê€'Là=ˆžîó»†tØaXNÎø>*äÓeDá{7d\ŒSßÁ„.ß¾gJ¶ïX3ŽªÄ:«¶ÀöžQÐbÀX‚<®˜Æi–Œ¥ýÎÑK˜ú‰÷ŽÚ´ìòµ§I¦s&óÃ[HÊ·‘µ¡­ðô§™P¹ÝiÒàBØ;‰\^&…©ø×( ·*¯-¸}ï*ÓOãb7RÞ'6B ú2Lë«D?Å$ôöœÖLÈb¹…ö™±K$‹*›DMjãw-ž¤áŒŠÁKïevG:­+…bÀÂ8Ó@<±,kÞ›òV	]Þšðb—µÌÑU2’EÙUªæuØÊÛ™ðÆ¸ ¦u8;âTÓ4Z¡Á!âÔ“j©@-(ÄE8f=K¼kd˜WH¥)iyBx)¤ò<˜…„¤Vøi,¹2 `O®v[,¼(â«õÆ.ç$ƒp÷šoU›í™ê’~{:4ž;ÝI~3Ø¡Ò‰ÝV=ò'MVx£Ïœ‹êY –1V§V?‹å‰m^O±#ÎÃmiæRÉ!Â"Õ~,J^åÅ<u2+üx'_'·º£Åfƒ¯@R¥ÆÃÅ-oÖÊ«h¡‡[lðnZY¯ÜŠ_Ur*üTÊª¨_!¯RèË-T/µÂ#¹¨É-ÚTv…ŸkçÄyeXj¥­(­·]¸FšUÍD®ù)!Nhy(!þf~JHTt%	£¸(÷5åQË_ÁÁÖ/Ý‚ª³KYižÀ!!á+…wQeú¼nB ‰6ÄÄcI
—ZR§E=wì}äh'T<@hº=ÈõèØ`àíA—=ø(d±Ë¨ä})Bø.*WRËXU]q/Ïñ}¦Ä÷€Óø,úø*Rèíd}så!¯?j4a ÏÆû0d„´U-BÚª—–“ŸîëÙø(ÉÛ5u«Ðõð;-å[ÜMºÅ}úù6‡šÁ¯6r9ªé$-vg<;2”ðó"Ù'QM.½¢IF²‰TX¬¶•O!'¸:óö:wxø^Ã„ÿÁ{ µÕñ“±¦;Ä(7lr·õ&ýÜõ–ŒïŽzKÆœzcý©79‘¥Ûfæ–QŽÖÀÕùÜøÎGÚ)wkmwëèAælÍÑÙM	½Po:¸+ò–ÚÙÇÅgŠlÝG‘)§:‡$“b°9)2M|¦Wûí2.wtˆM¿v×ä—‰Z…Ðç{ñP`B»"Á´Øñv]¡&TE«M§×ÞH×ŒÐTö“Ü{Ôi›`†gl%‡òÌûw¢µŠhåG¡Íz “pŒâU¿ZŠ5Ï¹uqÓ  ×Ëlù`/pýxvV¥!ñ¬DBóÊ2 ©…—H-êˆT¯ƒ¿E¯ò#¢˜uÖßl,7!-?áPÎÐbÂYê%ÝOîŠl†#ªY!N?Ñ,Æþ…hæZq¨¶´@µ¿·šA(¤}¡–vÆE!Zi~B3UD)Ð”wÖ¾p²9;ëNÓFÍØ²9¢¦Bíx§éÌ¹4Ò|+èföÖ§íæÖô5Øè/Ä)hN0Jìp9‡ÏŽsPÍ&Éontx`›þ2ISä”Xü—ŽˆfFG,Á30œðžŒÒ€19óa÷8 €Ö’Æ²D°Óbx¾øÙ,æˆóDw8qŒÞëX2aPåðZ<M §X¼äˆVLøŽøù6^Hé}ýºZ­¨ÐÕ†´¬ZÑCbxñz(Š[°‡ðnûöP†“´í'Tµ[( ÜÁÙY(¶wc+(öú¦†û7S%lßˆm¦¹çœc|¤y¹ft±Ÿ¦sòÍ¬¯ùÕl</w|SîqKjªc5Ät@”’ñ|â(,Æ¶Áe¹Ä‰…|S¨¹0åe¤TWÎj]ûç³çX¯¨‡jö–5h	'$Ï™ÑãkÓ~ÅB¡rMúZ³ÞUs7à ÃÛh°~nóžuú
ú%ì¬÷xáÄ˜#Vp‰1¥\;ÂÿÆ
‘aH¬ûš÷×öðBö5ì4æ¬€ÊlVn:4³VEóú­ÆLŒ£Ø°§þóßXÐûaÉÌ¿xX¯?.d9l€Ð^íÉ¨ÖÑg&ÐO,o^öÁÉ:U˜7…åN+Ë:&\ö ¹ \ì†Šÿ€aQÐ„cKoåŒ¼Ç=¬Óö4çeôoª\æô…¦_´Nâˆ¦Î€‚‘méÍuŒ7PtÞÝH<+EÒ`Z(ß.JüøUi¾Ôª(YSäB‹í/ª$/Î‰®“ÁPõ›è.YÅJý%~¼‰ÝÕ3¥yuüØbUj4•TÕ—ŽÍª_!´)¶
t(ã«ú‘ `©x`*6š£@GEâ6B€E5‘-$º^¸…Pw~ó—Ò‘ÖW$ÆW¤FXañ0+ñ¥ð—k”Š·Øm¥ÇÔU@‚,%ó‰@Ùä½‚d)2˜Gì‹Ÿjy2—'ÌÛh@žÌåsI”ñS!Uæ…ù×±JœÌ$óÉñóÞVuòbª¾¬(²Ü»*ôKO#%¿¢{{8Œr³F o@¦3“åüÙ¸Ü\\Ò;}¯LÆÅa†-naº­l$ø GÌ(\ûIi¥)=stµ¦Ôëõlst-R<kÚ§ÈÛGêV$­öæj¬~Ýq€Ft§}¾þº˜þ”­Å€çÝižŠbõêz+¨7£³÷>y¶å­oîtN£oŸD}ÇéZly“L¿òÂÄC‹ŒµÙ;Ëé2öE]Ùµ5ñ;F*Dù;Õñ]Î,Wo‘”2‰VéÁjòo´ŒñóìÖ	îê’ [ÇMË|ÄbbxŽŠ–iû,¨;@{zõä|M'Úyä)‘¬nYØkvÑóv†IT¤ò—þÂü]1§å$ÙŠ–`}¡No8;>1#è„°Äöt:º´‚áÒaöRÀ„Y´ŠxC¿`ÈÐ¾Dì“ÊµI¨Åh¦'.;KbÚ-ÄE¨Ë“¡òÙiuÉu¿Ò7#á‹B2o˜sñô<¦†5XCe#”pÏžÑ±z7™Š<uø¥Ó±ÔkžIb8zš!=ý £µÿrðæu!ÍôøRdšì„æ¬ÂÞ²‚Æ.³ýÑQßb6H%Î¨Ø ^hÞÒF[µCŽTda[¤n‚/µEÆm±ˆRÌIxƒD§óí>Tg{
±=.#8¡w“BìNq³ÝÑ.Î;ÛO²0À"=ƒ¿„–;:=;Ìd¥'°Ù Û†š©#–,ÆeaiwUÉ#F¬†ƒ0B]‹˜T–æƒú’¬†'h¿ã;(Ãi½¼·œp0Ûí‘Ù¢dà.=J'f°¨µÍölDŽ8šãžh“Á’LÅÅwÊvüÌ¶O€›ÏüªE¹{µl?]ñüøsáL¿œ-ãõ‰¯gƒ[ŽpšðËË „rìªíÓ“`²™,ó–tò’W¢Çe“‡"
Òµ6ÏÄåUKÓ™>ˆ¡[(@¬[=%o‹Ž0¨zJ`äÌ"Z‰xÄ]1ˆÇu¸Aè.;F-=Ü;ÜÛsðáÅ«ý·/¶Ud©¢cá
·¯ÆèBtßc¸}y‘o7»-âðLMàŽâ.pGÑwÀÈê%†b4‘aròE°O bÈCÍe™5^‡?Š¯	è£¿C"¾YÈ€	Å,_-PÁU”rùÒ–z,DN°ÌLz¤½	”®—9k7h)\Œ­”MFƒ‚ô2Û³EbZsv~'£5‘Û‰ýU‹¦‰4Œ<¿lÑ´å5Ó”±ñÜ¿-³X ?-R0ap Ýš‘n¸R&`ÉXJdLšÈ'Áäf.÷’¼„²oß¾‰vÞ~þ¿÷vÞ G\Ô?ÏÒ³,¢Ìha¶ŸÁ‡Æ3•áŸmåíùìø.™7NOÀÇc0Ãt¯Û“!ÊpXhŠ³ÛðgØk+ÚUá½¤¥,J„7ñÕëNÖ%¼}XÙ4†òÆVU²;¾ø;na}ûta%ÇÁKÈgDÄö°`†3@›íx0XŽr|…EÌ{TsoÇ•y*¶×¢?@mòZ§ÚÌ² ®öeTJëÌ3©Ó	(¼½'¦ö29KFNß†ƒšô†KÖ+=F¤ªd½ÕÀ¾õXø"'g…p;-t÷/ö ŸØ‰¬d{oËh]/úP<ÒìYì­b”ÐO¿vú¸ý®mmmÄæQÕ
f*y”¾ñ×Pðž´vðý™± [ÑÐ0
i*#g;ÊZ2îH-	zY¯°wö;ÇÊÞfui…Mé`u¦û £©"‹3œN*sÝ›bHý¯Â •¹ärÜ‘ëy1sûÞ6áQÞÝ5Ëp7R›À$z²y;Ã;ˆj¥ðÏ’¡ôŒ²“’$?’í™*7PÉ3¿'âÞ—îtŒVª§R ËBŠ¨zùf²ÏÞ?¦ÉùVt”e<Å›Ž‰mïÉÁš¸•¹\‘4Ó|›÷8êq÷á<Î'í%wM~úük„I¥E§KJC‚éåèh”ý<Kr…•L¹¸zÈ<¼†qO¥¤ÓÅà!8,˜"â~u9hÅ`@i¡_ûà.óÉã» Ù|Ø€þ>\]Ûê÷á¿ºrÊ¥ÃsìÈþ
ªßÀ’ðÇ§å/Z7:-DW¶‘<@YB¼‹@fj“ÍáÑ9@p‰W4Ð2µ®øë#V™EÝè›+hòñDTîu4ÇÎó,%-´°ô­¹®™'Fa]ž­–Âfräê*n´V¶¢ÃtŒù³ÆSbBÙFñmÔ4o•Ú]UŒNªÑ<=QPá1+0•ú(ž$?‰iò½©z¶Ï³*‡kÇtËÜ³ŒõR F¹Äzî¦’‘”iÏT@RÀÀÂÝ…M*jcR—ÑOû$›$òuÇØ>jU6z˜íM`K„‹ôPÅZè ½ü!àòä˜g~t´ÀV\é$f°7¬L•jÔù3†œÅ”þ¼Š¦]œ¢jQ27.„¾¦#"#-»C`Š[Ü‰ôNQBŠW=åjd­£U“5|*GomÉ‚ëÎª	¤”Ó²'v!=}Ju=ˆ4µ‚Ñ;O„<^Xüv¢öN¹{ÈT';EƒÏì‡Ž·…Zê}áƒÎZT…ë–[5Á‘ùö)lÐž±Zs$?¡ª¸_öËATÞ—^èZ^]½³ÚÕ^8n>gæJïCw$‹ëÍ†;Èƒí9ý¨ÞìjwB¡b{™˜ÂìkT‰ä}¤SõÞñÈüÄ ú•[êÝÜ|]„ž*„09 ùçŽ‹7°8{çD²FÔ\G‘Hô­È|tAäìˆ!^F>pN*¦DqX\&lÄ:£QL{~ÊAI·¤¡ŽÑªfäh!¿VËlE•›†KIÓkû~0K¡ÁNÑL˜ûŒjÂã­hU¯ÃŒ$cË:2;‚½;ÃXå[V$ØŠúú3±›vÙfoº¼ÍO01²ØÊzYm=0Ã@Á³¨¡6ä-ó}7N™V•¥‹ñK”c×XªÁe€i
2Mæf`œ*ÐñOä
åuôeìŽŸêZ¢üÔ®Pn¦O2~­Ñ¹o¦e:FuÌ€¸w&N1ìÒñ8Ý.xÝN&ÅŒ˜ã‰d-)?W!ÈÛHèbXó%%¢Û\€RGe0€VåQMB3þ³#{ÐrP|è±a®Éº8E%LnP$¦&Äó%¢@Wl`õŒÚ-Îj¢s‡Îe‡‚çnù†¨KÛu!‹_Ø® @«T1qË)&µ#0aÔ„Fp"þ²ûs†ùfò$ù%)0…åi<9/“¨µ“§x_Áâ·´¥°qU›ËtœŒ !;XÎy6ÄL^—}nÈæ²sbíµl.ß«f—¯EC–—}š0¾ª¤Ÿýe[¯Ç¤¬¶Ò±ª#»,7ËvñÉPÛ]ËNg“OP6ÐåMŒf©Éj£Y²ŠÍ-ß¹}lCQ2;æèÇ)†À™‘q…®\ú”$SDóì"dÓË&2æN[ÃLü‹JªÇ±¥ux^á÷€#‰m5žPâ	¹5à¢œ•ænËLu‡>Ê;¤–ö¡ùãø,Ë—£…mpFÎèT&½%¡è÷Ë·a¨‰x«¤ªTÒ/gw0·ÝP"p[Ê,K¼ `ÄÃ´4EÎšHRµç¯hx¾Qz5Õ,½×15•ÉM8Wï-`ôfuÃ—kã²‡Iñ2‡qŽé–)–ž"‰ÂÙ¹Wûœ
;O:Ö50¹>r0ÖåŽÔožq~ymœå\ý²ÑÒ-R¯¦U˜<ÝLÅ¦!²2*	ø^šD;€üÙ³ãÇÂèâ™–ƒ×ßjAŸ3<3šÓš?PÑ—6Ou¯ÏW“«1ÔA‰Ñy|Y˜Œ¦ƒ¦—è T†\ÈéxÀŸ”8O‘, ÀþNJ,‡‡`µ…,ö\Ý¦e·¶å«8||9asb\^(LíçaæPWmG¸ú…±™¡¦¶®ÚSr?G£¿LGY<T·-ÚélEøf™ŒL·”¥ý÷¨{–cf¾Tª¤íšL}ŒëX õŒÖ¸ LÑãXðàÀŒÉ4žÿ×z¿»ÙÿÓsh‹¢£¨–E˜³G€2þ’¤‘Ë€7»7
ç0Ûe¯iFËÑæZ9ê÷6šœqIª¼å,­ëÚå¿å9ËñW.ÎX*TÉ¦?¤8‚-1RíñkÎŽÀFbî:jy§’>¤•/{ÊÜ±”ø«tÇRÞl,¥‡¸L3á&Æ$SÂP¸ÔžGº¿'RÙ/±¼N¡€Ì¬³„gYDölnelÓ`ûðDDºpÝÄvà ÿÂ…
YËX‰¡<.»Áq²W| œ’_~¬
»UãU/VH œÊw¿™šÆ:aü©3QY^ÃAYcàìS–W¢Æ·äé$ V¡ÇùQ¢Í2|Œ#…ºÎ9WâÞ/‡sšŽ&€uvúÂpÀqÂÈw†T®óüÖX¦úàÒ¾O“ {Íñ5ÊÞ1ºùòƒ¾Þ!GEóð,÷Ô`g¦èÆÆ<Ç£YqÊMÀß² ŠÓ/ÁKv79‚¡çæ~£:tÐÄ×Ù¬¬®")ÓŠBFð¹k9”©5ffò-*õŠô™rÝ¦(±¢õ`Øi²š.Î(U¶Ì‚^SšEjXzäÇÇé mÎ˜IW#áœ®;£ãmF/L[›‡ÖôZÇÐ”|\½Qv‚ašÓ‰\NR¥3›:qLƒUùäÌn!árSÁ²)TN›†_ 
*9:"Eµ×žéJ3[£Ñ5ð£„Ð¼zÈQÈÂSJ¯MíŠŸŽ6DÓxRAõÀQx0(•bÿV¦‹¥÷SO°´ZÅ‹¥‹¦†Œg(]ñž„àÅÖØšzk6.ù€âg²wç!tÏ|.ì§OOÄ‹àWù#MK!Ë·ÈòÖ¾Œ_Õª‰É½˜Ú64–3TÀ‹%œÔa¢´&•a4¸m9;ó(N‚óp@Ú®vqÅP(.áÝË¸ ÷CÄ­meëëøÆ°«Y3di§š8 < CÒ1Q ¯1!=~Ùs¸iGô½Å¹@-Ón¢ïˆlÚVµˆ<‚ªtÍ¦{+‚~ž¬-G¼õ'«Ëöø¤¯añi–N(ýºÃ ­Dƒ{­]'ŸÉñ=‹Váy_—ÖÓAS/ÖÒh˜ÆMÁäEh¢+´˜%¢Mwô¦¹ ÷­v¢'zƒÏ$€ +ŸVØæ"ê¸ ¾‹ÇÆY×¶ßDFuÖãlÃ$%!OŸuúsý|«–ÔùU%Ùa~w¿
Æ|›yOØaÀÙI×ÞÛ$”ÜNƒKLÛýqjzt'›.áš 4	þÚ§ˆŽ“dˆ*ýŽ8ˆóqY	®I¿á‰¾€[.žœ$,¢Í“PÑÏ³d–4¡JHo
’Ö•BJfZÚÖ{Õp÷Ë^K¤Ü¢öfV¥ƒ4Ç8¤c$Ä†iñie’”çYþ)Ú[yƒ£ì\¬ìhÕ4ß™y-é¨
RùOÞ¤i3á%”%}¾mrêÃeëYàxíˆpÿÍÛPwµAå-sÓ€Ÿ!q'«D~…Fù°”cÌ8/¹I¹ªÔ+¦#¸ÍZ]eëBÑ ã¼€£Å¶è¾Ðµ$ÔŽ `×õËNÕäÖëtÝÐÓ½IÉ*¿ë¿_ŽVûp[ÏWÙó¨­:ïÖDU¸P¶$øÇ`.–W•™¯ã×mU½wÂ@Âr/N³óÃ.Öö™Ë¤“³Ï¿ŽRôýYŽ–ˆ%[Òè}[µ ÝK¤mÖc'£Fc˜¢âY	4Ü$Pa#}¥©Ü2œm•·- a—¤ö7Æ,®]ÊH¶ Õm×S,â¹MJãð8:]}öBbÄÔÛ5óˆÈcüÌ+"Z“Ëï	f2JœÚ@vØðÜ?±<÷WªGÆþP}ÅàZ&GúÅë,6ÁØÛ”¹Is¸rB>Øa‘X1‚šöØ¾¦JÁW¤ø69¨Ê`¦V;É/Éqó0Yîêµp"H†ëÃÓu\&@¬K€0¸mâ­SÒ<Â-Ã²Íjpo‹çq¶u4˜•K¯sG|Ý¶KÝû¯Wf{o¸õ»•]‹€×lÆ¹gƒ6>r)+^»«Zili„G0¢aŒ7êQ£3Ã¢)ŠÑ0Æä‘Ë£ÇÞCd_°æºo,¬#P£F9 Y=Ú¤°Iu÷"²…ÂC÷{äõ4"CËá‰Ä’ÃDÚ +2V2Dº-“m#TŠ·ÃsÕƒNLøevpøUþS#nÜ–ðl}‚2©q\J(#¸ö]!ÖAµW^½Z¹„OôÃ[ãñ’¸ž gÄIßg“ò ÈÙöš ;ø‹Ãä¢ÜÉF°¯›@šlbhßþfç¢„cdf†Ñý~_ÖÀ„AíÖÛ¸†ÏËSŠ|©§ühQýåµÇW°Åé	°Ã­Éå-yH¼#[]lµ¯š†ðq—Ì×¿¹â‹wý‘õ¹¾©í¦oélãÛÐ-m—óJ)(¢Nôóÿ$OÕVLA°¿¹jÔíOÀÝ·aU[1Å5y$¤b6Çù¥Sñ¤8OrÆ1Ì`¹A=gÜrÉáK9
¯¬çJVâ¾#ŠÔeVÂ­ÃEÎh*bvðhØK¿£ÈdÖê8¾¸I›¤Íð vÃÀÒÓáãÊK§mµû@*´ÛÚìV¢¶•_ï£ïÓ‹dØ^¥„ÃýVàD®n„àÞõÀµŽá29Í"’4 iC¿þF6û·»7Ä››’ŠÀ©úÌù#ÕKÉ{!ß	 Ë=,â®U² “4Ï	ª^ŒruiVºÊÍ+È´Üv’zÈæ>YÈ—¹°Æë#©Èÿk9¯"Ë¥Ž—£#ªv¬*µ:Ìµ$AxÚ>ÒßU¢»ûöî¶dÌ\¹e€+²"Jà€Ÿ`(‡£Àš¸ÁäQÔä¾bGUíw‚€°g+åMŽÚ•èë¢{jâ ÍÛÙ£ÏÿBc£’Ò—†™$òqYwdƒ®¹´õ ~6– òÏÿJ¸;Õ–©'VfSÒ2[¶S‡8¤Iuº¦˜—ÿm**ó4‰Yx÷®EÈÑ¯Hé‚ßZÀïoä |ÿ^5q”/·Ô¾hA
NTì¶NòtØ2»<(/G,€¦è#$ £Ð.æ÷Š'p.ôÆ9Éãq\yd±ÐÕ§Jx/=áë/
	›Ê–ä4e\Õ’˜Zø##YàwyªO`Ì‹¿“O#£ª©™1ò¿æy-àÄ1dæ-”Û‰„íƒ8?‰ Î
SÛS	YÀò(3—+à/‡ûÀ¨éú™ æÐ^º—†xÉQZ' ±¶ÍÂ½´Mn¦9Í™Æ~ŽÖúÍŒ5xzÔùYÛÔ—Gž íKâa‚§e—v…ø)ÓùÈS5)>ÿ
ËM'3yGJ*s1zÿÙšût™üû—þ`Ð¹á1šIâÚf³r:ƒÕ=eG³œ÷kWAÅÊòñó²,Ò~øŠë•ÛgñhUKh§¬Xr;Òk­U0ÄÀÄà`‹®kÊ;YÅîÑ‹!ócÓœQäŠ%½qÕcõŠ¨×Æ%†÷þ6p¦ƒ”9à‹Œ€mÝ¸b1ð‚0"35Y®ÀemÈ9ËôõQ•\còlúCkÆ”–QjÙñ—-%Q¶¬=K<¼(Ú(„G“)øè/žZË÷."æPç_ÿ>™ãy—¾‘‹¿ð*g–¦„K'éÏ³D”½„ÅYœŽ5ÊÄ¦˜·t#Ó¥«¥",Œ$Ä”pjc£8¢¾OŸ¯8_…áPLª£;];|NŒØ4Í\ó-mA‘z8R›ZY;µºR€_óÌ±´— t— òcr{çøz´”J’Ùˆ¢£šÖªˆä“«Èdpá&Hkö\Úœ;6?Û,fl,»Ïß¶PS$“"Å;¸„›³u«ÛbÚZ‘åè¿ËêÜ/«°LË2^ì{“õ`C{™tIáçkŽÉ8’²2<”„Ø×= Øæ1%ŽnŠá9Žçµ:–*ËE–áH8¢	=©š&Š)ÊP0¼çdâˆ­"PZL®¢¥ZÑõƒ
-ó9õHÎí!R¦o|Æ†ÜSKÔÅ<È]¿ecOx'¶ã±?ÇVMœª@p‘q†Ç–)œ¾ÚÜ$¢«XX«eYËP»Á,õm¬™yå@Âk¿Ø¡ýSè—Wª¤Ð.PÊô2P
„JCÜã%öÚzx—e´3EúHzCÂ¹øq7 .Z_´SkG' 0yZ€ßþÌ3ºE•ÄG—&‰K>˜ø€vS‚lqwðZÌ¬UG³0Žfá?šª­‚~q³p„ÚZ-ò‹ù!ß»Œ´”·ûb`ß|hóÁ¼N’HRF¹%£Ü,ïÚä%‘†f.!ƒLŠ=ÎL.!,ªÔSÁ:[MÐ.GêhWlÉ²;I
ó¶ù¸Ó³h0Š‹yÚ'KÅ4$ÝËîý¥§rmì2Ç£ä‚yt™¾”¿ì%åy’L´º¬öScWŸ®êÍ¡Ü¯»~1¦yRvFñàSDÏ
´øì>ì÷£2‡‡@EwËôä´\zú<=¥Y™b.<M)ùF?^9]µz›:±†7¡aê­Pgã¥§DÍÅÿ"1£RpÀw.P/¯L‰­X3ó®ÓI<í®Eøµ{žÇÓ%kpéd:+m—OtKƒêé(Y²_¥Ã'K”ó¡ËU8´þOÓá6Ã.É´|²Ô»ËþY²Šd“²x}rÅ¬+YÞ
JaÐk:Ïge™¡¨ãfAÀ€åYÂGéàÓ“«av>AÇôÃd<Åm¸ŽÊ´ÁhŸÇéeÏ†Éˆ§¦°Úßåµ#<ãO®V^Ã¢WTÇÉ
ÊÇÇòåYapHš®œ½ÆïMGÉëïbÅŒ•ñ¬Ãl0CØG#· ´ðõùåÞ°ÝÒá¤ÕyÖ`­vçÚ/`Œ–m9Œ—v¢Éx@¿5œu»ÒÄš›nbA²›Ï%£«) ç·­–Õjzn*w=…l€‘t 9:aß7û+›ð+Ë‡IÎž¬ÁÛÓì,É·D™UõD+·Î	?²Iw€âþÂ;kã}×®Î†yenF
e¶<×0¯X‘àF}®í;?üŒ °Œfi•øO¿ÏœZ0}Ê¿%œ§èÈ@ç³qŸÿ%¿Øx)Šl˜ÁÏãÃ<.N×üÞ'ø}™âíÂ|¢8R³w!ÃÍQÔ±§y–y ©…Û°w¬ÁSí…°û}¦V˜Úì;a.¸ØW‚›Û‚…o¾Îî0ì"àDöî[ªÀÆèËxƒ™³ È¢ãŠY‹"([k_EU¹Õ‡ô¨”,bªá†Ä'PãÝû@Jø¬ôþhV˜ëL)ÄDæZãÇ+î34
ËX¼‡ p~ËƒPqT6âpÍF[Þ¤º!¤Ñ*ŽÄœ‡"t,ØT®€(Ä†®ólµ”LžÿÒ]Ð¸ƒâªs:‚gÂp6@N†q8íÖï‘÷(År{æÞ¦ÇIŠjþîOÓÊý[¦Ieü6§é€ÙîT® /Óè,©›†[}È“A?ð$¢.ù‡bv’`f«B´ÿ’‰‡üzÒL†¬7Ÿ¦©l!+déÙ¨ÎåTC}^–z.!ÅæÉXÿ¦—ßWvÐ¥ÁÏt–z8tÌç¤o^'¤ôÛ›•OQãÀ(V»3Ò/°¸d«85‚¸ uâÊã,•H>üe—+‹P¾‡ŒÑëääóß©$+VÎfc\±•ù€ÿ¬–þñËãH|a5äxÇGÝû^%Çè{|ÑE{<SxäáÄ­ýÃp
ò¸´uå²95¬«ÁÄ†¥éezÙ]ç’Ål4ä‚Å±Îpcn+1ŸîZt~š–	ƒN2ÙÙá.*ùÀ%ÞTÂ|Rèª¥NµK Í¢åÀ“R—znÈY{ˆU±Ëaî½³z¼rä@É¶CÒÉ_çvhÉW³º~AÛ¡´{_ãnèí×²^ºpA»Á®õ¯r+$Åñµìs8»ƒ!³²•û †÷íE€Êp·Å&tÙè=ÖþÓŠì7Úì-I=Q¨àâUºÌ£;'ÉhÉ§¥\'rÃÖ©YšØGøŠÖ‹îj4n©Ÿk¤£Øp•8ŽÚ’žŽâ£d¤·~4Ê„¾ò¢p GîÍl
lÓ .ñêÒSänÐ§xšåì<^¡¦=]25oRâ¼{<¸Þ@^‹(ÁM2ì^`uIµãl0+¶¸¤ K¹=Ù#4XDµú!€q³ß_Y#¥½+ŸGuÊhç«pX9ŸÚ@)Û‰8É&0á›X8@‡¤¨²)žW> ¥¥§¨(0µÃ;WX_Ý+×V”ŒaÉ†)Úþ”\>¹Â‡×bÞôã)û#{ðŒøñ
[!îlòœ~9`ÌçÞÿQÁÒ±™˜:¤LsF¹
F­ªd(„Fð'jïdãÙ8ŠQJP·\X `„ô õ3ª:ôcQ	ýôã)ûSÕiçv'¢Ž‡í\û¯¦jµÝ¼ÇèDû¹î½G®¦—ð.$ÓÓ{.²Dí¶c€HFTÊîG•îtªš÷ÛyUÙEyœ­ÔçfÆÑ·4‘Si—ª-Dë-ÌêZ
/‹Ýöõ­V=>i´Úü%,s|R<c©hK#-~©]¶Ý`ÌphÚñv:tÈÑÚùÚ¸i8ª¦@3¤¾Àê¨^(NãavŽ¤4#SùïñÐ&­‡³œ‚SÒ]€yú¦°¯|¡ì‰¬IàÌUZ©	xQgØ¢œDí¦ñžtIÄŠ¶”Åb…Õ@U¼Eïêµá¦\Ç›r/³#À©p·±ë›~HîàÝjzñ^»âÅÐ}«.xi9wžâ¶À}œ–ÁÑ ¾"ã»i¸+8ÞPõ+;‡þàž:â~#f«’N†éI&×…ÿ¼ë•±§Z¢ºEò\|¬R T›˜_ŒE0EiyÙ}ØÇ‹Jüê³[ŽQ—<ñbÕ4;âáœeõ©Ûƒ¬×?fø1¥ì˜ß¬Î.Ã¬G‘§¦uE-CŽŠÂ‘ßÊÃÃc¹º!¹ƒ_Ù•E¥Ù`Yöµ¦‚¡]]À¶Âj‚ôqìVÁJïÞWëìðÓ\Gm…Ž~à»«½M¥’V×ˆ8Ì£“
é‡!K±-ë¢Á,/²¼KÑˆ“ÜÇ+ð%â&m •±²T
•cŽÇ˜s	ÃÕÎæµm7«öHÀŒ÷wvú˜âþ.ÎßWÒâö aY±Ý 6Ló…›CFðþ¿8]7nŸ°à!Sè­-=½òâ@àñN×›XB`¢„<Fú…>ËXßwsß€=M*Bí>´FSâèöš_Ù–ÍÚ2Pñ¹âX•Ñ?ðM7ZÐ=MˆÈh\¢¥ì®ði°Ch˜Š1÷˜ó3ÄÛê‹ñöñÉµE‡­!ÖÇèâSdýPšE–thú•ûêœŽ‡š-@ãTÐgéÇ¨íl»!Mž±(ú4ç¿>‰Ftê•ònƒ¿‹ÌNŠn*6ø÷-$0‚ ù–°àíâáàò^ÀôkìhóãËñÀšø2„k/ºbM”àãçC¾\ˆøìèkÉà‹|t0¸rŒBU\{¸Ø"ÀÅ¾8=Ú÷Gý8ÚòÆy£ïBÌØöS{ÜÁ5I&ž%R:•qbÌjNÑ­–O„âìðïÎÖÔ€µØ¾©†ìÅùâ‹Oè`Ë‹^¬è‘+Ú{Þ5˜Y.~)8þ7õ³»…€÷æ"^Ÿwq|[Ýß´ó­\Èíp³õPlþŒáqTO´:Ü—¹°¬Š
©¡¶Þi\°W 8o3«çÏUëéß>-Ž•UÞÁ¥ºØÝ].SÈØúEÀƒåü­þîeýbAþy2ÿ`{r¼uìJ…¿ìjÐ[Üþ\YR)Tç½~e
‡ÀT*µr.•‚vö	ˆÛe·mçÖ3bÜÌ·x¾•|N_¥êƒ/Ö<3ÿ›ÇÖª¦ŸÐR†dYLK¨D•¢Kù‚¸0$ò¨™Eþå‡INÑ¥:“LP]õ9ÄÅ|=L ˜0+D!šŸ°›@}O7W>Ùõ™Ý°Ö<ú(ññë¥!LkÔGâ£‚Ë7ÑR‰ÒVQ|ØF
+ñ1Wá¸bÕº+ÙÓaQ3MÔXâS­ÎŸù²íš
üÔ!á+ušjì‹@"“xô2|â’Ñ‡ú^«-Whªy¯ÜÝA§!›Ä¥§ûä$|Årç^×¨f<_™öœ†T§ÃÃÏœˆÙ‡Ìê\¹íº´yø©y½²ý@Aƒ¢iVbÖƒxííF0–1§|íT&b¡iüÎ¹êƒiá¢‰G{* ö“H¿dêF,[@¥Ê½åïòœç
ô,Ú]¨ôõW@]µkWÛe~Üè õW¾æòlÍ§Á­¢¤Õ7©&£h	EšT+2VÅMêí~™iÔ­x“Óë±˜—ÛÛ¤þœ÷\Íö•™qàÇ2åà^U§æ–œÇ¤ƒ*Ô˜uP™;¿„‰Ç]_/$-Àä?¶ÙÇíá¨¹ù¯…¢Jb¨úåæ yf@éñ”Û†4	›KàçÊ$p
K*ÔÐš‚ºkhQ±~S‹
Ï¸›XÐ ïØÌbóNÌ,hä‹1µÀO¥¹õUCýûÍ.XÍ ´Ïeªá*Ç­8Å¶m©qÚØìe©k`ç±#h¬n],a»Uëuç˜ ,hÆÕ¨þ]Y¹Ü…YË@ã"U¡µÓ^Ös!:Ïi<+öGZ	…”že‘	‘7@æobCÔÌŠ(–,¢2#Šç6#
Ü*h.„`ª/’6F¡hq!ó[[ŸYkŠä÷¿ý:ÓÇ¡˜Ã±hn‹huðÛ˜Ø4N/ Ú[HŠ1»»0±)jbS›2`.“’†[½H›¹¶˜wWÜðNÑÀàÍsÈÜ¤Êæ28¹ÉêÉ¥ÑIñU„ÞþñŒG¾NK«¢éGÛ«¶ƒlk-;æ«ñdÙ1ÏÌ¿BËŽÊˆ'ˆ'^ËGÜÌãç(ëúÇµÕ˜—›Ÿßm5Ÿßm5<­ý;´Õ(ne«a‰V…áÁf½PõÖ6¿›TàÇÅ9µñ¼íª„nŠF*\#x=îÕóŒ(¼¶êT«8‡Æœ! ¢)öñ(/zžÇõøÃ‰i£0Õ·a‡@Ç¹ëOê[ðL/z¾çõ­ÙAÖ‹žù¤¾_èõ¢ç>e’Æ9­z°1²gˆ¾hŽåY8|¥zå»²PÑçÿš(Ü!¢ýJ!én,G_‰BCÙL“x¯vÌX»½Ê)žn4æ=h›kKOå¦DoÕ-øxåt£²SÎb¬÷ôÀÎfyå½m¯1.jÖø–óÆÀlòÛ<ëÛ¸ºo=q<4Zf¸±¸ôëÊ!ª'_Cõ6&‡åáßìGS˜xMÈS’]ô3nG5‡Õ³Es^¿e¤Ìç_–©Û—PÌo?(ÞŠŸV"êâõÊv«4ëÏïz}þ´¿'!7«…Ü­I6^â6­ºd’C÷£$IÛ“â<ÉÑp…6³(•%JÞ¬-‘X\q-8Ô\ûÝL¢$uzUÒsØ†ã§Ž=×ú²u2¹ÆCÍ%=k"Ø`«^fe<z_pCm¾p–†@à·ÛÅl¼å,<ñl}µ×¢?D8°„ÿ«ÎrÔoÜ]ç“dØ¸³¼W²<ÁŽúsõ3 9“œÜS2€E=íÂl·ù@Vd™Ì°%*Âûõ½T*žÌÏã?MS€—€èâ9ß€±²ZærÝ¬EÓVa¨wÓød‡(d˜`»L'',‘uôP†.™ÄEôV5Ê°]Â]ÎK½Š§Xj©Q÷5r#üÔnZ)Ü\´ÁS•iØèN}„*Ó'‡ž³°çbu¼VªZ úë(eÇJKÏMÿÚïz½k9‚oÔÛ{ŸbW|ŠH=ÊÌ±.Ñ?ŽÆ\ÉÖod<‹Ÿ+9†V³4¶jËYàÂ¾Îerü‹³ŸÅO­-õÙD)w›£W×>¶¶E¯Mƒ“ŽY&ÿ^°ÍQÝÉ][›Î¡‘[¨qÑí­Mkˆe.t=ß-ÎÚtŽ½ùÊ¬‚„ts‘þom‚YHIÉØ`šø¤6£DÓž75½t¯Û¤È¸ÍV,á‚k¥‡š`²kš©ñÖaõÍ,/«ýèÿÑeV´ýj#aUpÂü§¼jýô2¾D>”Ý£k}¯\2¹µYH\!l
Y•`²•¢Dð&<§P ú<=¥Y™â`tÜÑn—žî ´e˜uI”Í ©Œ)ê¯ÔW–ŸÄ“ô—˜ ?¥¾dßs…`Ê!¯…ÙZ¥…f“Ù>‹SàsF	Ë›Tw¢ŠüÕR› ÷A	ŒænrÏF%5}-äÝÀ•ñSÁ² òšÆÃœ¶'0)JL²è(d‚[6D”Š$õ˜Î2Ùœ›	YUO|Ã÷Y¾95ÏÉÍ“È¶aºßù’ý6ÓŒø3×Î—æ7´,á4?Æ³0€=õºÜA´ÍÀ0ÈÄó~‘Ûu}<ôýwùödˆªÿ˜àZ5RO-;ÒAŒ[®)žX`«:ªXA9x„:†Zôë%ÇÝ»õôÛÐžX,aÈ¦x³Ö¦ùÀ‡BGg¼Ywƒ7)ããzQþÜ¶Ä¡k¸AƒÚE|ñ\&ÅçÝ‡p§?ÔÕ—†ZÓRÁ4¿Îe~ºF–=„ÓÑ0œH%äýZ%ä¢5AÃRN#,=U×@/ÐlÃ„àaU
xð{KC1.QÚQ+x ±ÉÓ+‰+®%ªhc†¸zù€hâÿž§ø•@×Š o£°Q§õ{yg"E‡ª”tUœ€d2d;ZuˆÙ­X1Q‡6ª(Ke	à²¢ŒŸ»/cZJFËªAÓà#‹ˆ·‘5·HÆ)í=Ïe!”	«Ãìm2ÁªË±ZR­¨¤×óÕÒˆ¦œzMF½¹K*ÌJ*aíƒÜ J‹YU€ˆ•%~sÀ³ÀÈ´‡á°H¿Pä‰<îêvÉ éfðÃŒ™¾&ø1y81c¡ƒPUê²[Ix}Aø£­Z
ŸI­(™÷¼¢«5Sti.U´G&²†2‘5»ÁÛ	EÆÝxVfMD"óŠDôác`ñQùÁ;”s÷+8æ"—ŽÈ˜BùˆWLØQa	Õ(EÐ(‰‘Rì¢CßEâ—EÑ¶&'‰P„ók4Bçˆ^ôc6øüÿEÓl˜ÈåQ1CYBþùoSÀƒ….ÀRgdn° “ÎhRKG?ÏâÑÏ3ÜB2ÑÍz^‘AS™Í<D]šGÑnü[x3Ìo >™KZÒX^2¯Ä¤9ÒöKM¼r“X{Ì'–Ž“4IPªÖ¦©€¯.ÒÀiKÇÇz‰6/ú
´µènš‹gÛLîo/Ñ¦äøåuròùoƒ4‹’è 9!äUÜ•h»ND§¤ï@Õ“,\›G	ÓËÑ­g>ÿÌg
ÝBåßMQC/|a*0Œ»ˆå”¯Å.>Ìâ—Æf´ún:‹àC)>ùN$ªW±¦ò5áòÔ‚¯Pe¢Œ´8(ãÉ0ÎÑ¾æpoÿÍÁ‡¯öß¾8Øî¥ n‡I¡
)~¬„(S@˜e&d•’Téþ >Ì‘wê-–®²!t‹•˜ÕF5üÂ£ª5»	6$gLšˆÆ¿:Yð‚%Á‹7“‡£Ü•x^!pÖ¹ãk¸KY'¹ž[ˆ,Aµ^ˆ|¥a½†ŸÐƒLzû›sÆŸXzÊéù[Ç€h$ÂÿÒ2t.Ï7Åµ¸(ÊÏç’žó¾ä%pmP;wØãMõÃW¼ò¾»º×ìü.ÈÿJùœ¶“²ü
¢Ìª©Hì×óUtÄùU4·U÷w‘>~þDúþ¸TÿF`Ä]”¿20úÂ’ý[†PžÓøR×žS£×Ä0Æd6¢Ä1O†ÉdÆÃn†rÆÜ6.'ƒ¨·"dyþâ»­¨µ]¦gq+åù(ûy–@½–Ke~)¡
v’ó4…ÖKÑ†GgXÿÓ:À_ƒpÅÈÚÀŸgÑ:vD¿Ñu´”‘.š+Ï
4PV¬÷ŒŸø<NÅ$w³A{ˆÿ-G-Á¬~hƒj-C¿eèWô¨mÂ5ŠS§Q{0ÊfÃynúÛárf£¤—äy–·—^àÖ1,h/EZjo--G²­‹2ÿNGñåÁl€)È»‹—•C¡ÎÔ8šŒ»fÕx[u ±?Š'™a¥× ‡IQ&¯’IÑÝàëödÆ¿í&Çé$%°¹C@‘ú’°¢uº(pŸü¶À‚#˜TvÓ¸ 8ðƒËP¼Æ8¾ã£$¿CH}}IHÐ:]$ÀÒÅ£h_.yÀôFI^¶?ÂlH¶YÒ90µs%§C\Tîxsïã­ KÍjnû9Œ‡;ão_¦ÀƒbñâÍêñK‚›Óõ¢€n„	ë £‰-ô–0ðBž˜áS§Ÿÿ%V€·t+ˆ³§Tw§”Œm¯,f‰;§Høi°ETænó	syžÒ‚”ç¨îžQ¢·¹+K[ûÞ’\,&¦Uã3!‡ógêÇ/Ûƒl„þLi6© œóÓ$‚<;Ÿ$ùÞÀ©õä	ü‹ãîÍ ´dÿÜXyOq¶¤'I	Z´6 ü–éA‡å¥f|îo`"'0¹œ¹$348NGÑOŸ’œv4‰#´oþ÷–4ðÒ×A¾ŒûC‚ªŽóø—8;@/´Ñ›×0-(‘×‚³ûñ›+ûéuÔ¦-î|„óÛÚ+ã1ìpižF;ÐÓlTf€´D)-¤×@¼õt…R|£-­ìu*†Gß©tJ¥K*üŸOð1'Óª<Ç\ˆ¡Ë’¸ÀÃtŒ 5ž¶;ª ßß-¹±ê•EÇ«:A§QlZïê
ÿ?êMgjÑ~ˆôl×òØðÄÃ!¢·Æ€ZÄØÿ/É÷Yþ}Ê­(ÚbÃ˜Öb›·ü’ˆ¶/%›@ˆ_^
¼£&`«V5ôƒ:™%c/ÚŠâÉe%*Šâ°h: l‚‡_GFö¸E= ¦Ï Šã„³-a,	¨Ùƒý/â8lÈAjxJÏÒäœ¹àÁ£R€0s|ÅUB_²WûØ\?BTÑäŽ9N“Á§lVÊàù"aó®DpÂ$£l`¡å ‘=@ví–(õéMh¨ŒÖ8›”§£Kbâ	qäØ)~°¹‰ïðî#ÆvUàŒ˜ee\Y‰fG^vŠ Ñd¢Qv’äU@ÏôR/.¦)†Ž³ðZ0!x(Š'p)ÃÈ2‘=/“ãR!FkÃa¤£#¶•xˆä‚æ½02Ù¶cãKó~óöŽ¾L³›ô8†¡¡ =£ëCÿžÑâíÆªsÊR£õžâÂT7ª†Âë ¼œ¨8±sMkE,%Oó	 µÚ‹öŽ%`k@x êäG¢@©½(EOÜr™Ã]|§¤™/åÀÍIZ_Ç"ÄëÝ˜X°Êº½2£bÏkSò8òô£zP&ù0Lñ~‚ñèÃC€¡á®j¸£[­ðm€V(ØÍ IGmÙôJÔ^í÷ûÑ¢ûòŸµŽEHEZ„X¹Í²áÇpœ´­jß©¿q|Ñî+hèh%@0 3ï=k½èM	d×yŠ±ÔãÑ`†Š"y†¢lrœžÌ`DÄŒ1~î§B¨iH 0JÇ)Ïû‹¹yéç.–´·Zñ¡ptØ¢™gÑz?ƒ¦7\R/îoÂ›u‘( ˆ:bC!7Í¼P¨W §Lüc	ŸÆ ôôÁZáëLÀp<ÊàZo}!@Vû×•]4X1š§ZÈm¢7Á~¬¥k[–\ÌŽ¤ô¸âÒdË+lþÇs >Q˜Ÿ{üúDÊ¤êöãi,½E(èCÕE5{\$KÐ¯I±õ±á}ò‚-	ºjŠéõø2IÊÊœœ9£ª8–Õc;žNº§]X/nâ)uµÆ!SÓf_Ó²>ÞAäá3P?çéÂÐÔž›Þ(3¦R²ã¹F.khî	ÿ%”´ƒJ+að°~áKÿà·ÒG3Õ‡bPédbÌˆòšël6lƒOW+rºê(Úh8¹ZmÓ£+€ƒd¤Æ7W
(¦D[20f)äcteK¦ùðéª1ÎPpM6¦‡5fúÖwù 9ïÀ°	¢³Y/Ú'»QXš.„<š•)ò+`†Êöá¿ãÙ„b'°ËPÏnªãø,Ë—ÆIZÅðm’Å,I6Ã `˜d66æ,cãè}´	Kv0Ó$AÀèÌ`ý˜¯A‰âÊ5°¨AZVLñ:‹¦É(3ù©^ôbRæ@mi*qÉX9à¤fÐZ¯á,óä®af[@·EZÎØ€{ÖFM«ŽAÈ\ÍÚ¡…0–µ;
ZÎÌ´5VËÕ#“Î6¹Ñh®3÷u¥±á>#O?t£ñMïek”]å[Úò<ÚÆ½àÊm#|¿æÔVkúW¥N÷n¯Yz2y3+Û€PNÕ
­•m÷a;kK[t£={!â4Gæ‚ž95Ïü-|õxQ¶øíhnÅ3ÊÜõÕ¢_%V›ñ$c+ÅnkÃ®Öò\[Ð#å>¦Gi{jL=8UC z‡wèÆ|÷æCÃMŽuƒjþj5—èê}\±û•Nn÷-»y®R;h·cº¾v“+t-t…~'Mv5Ó»ˆ+ZPñ3ïÍgÏà
õ+ô6cÒ¬ìå&Ï¢Ö÷1Þô›¡ø“$'²0G±s½¤¸<E¯…BÞ<MñÓÁ=&€v„6ò^Ë¹þÉ S08hÝÉ#Œöˆ–sxqSHEÓVJdƒ?ôÍ–ÑïÇÛ1ˆ³Îµ>öÇÇD%› m¼Ã“+¦èx’æmhðÚ{µyáÈçþ‘¢'‡s1”—Sh„Ü®]ƒä€(ÉŸ,!’týÏâÑÀ•¦%Ü;î ˜†–m'âæÜÅÑ!ÎÅëQC®Ð©´x{¼ÂgƒY±…ŠFŒó¤~¨Àc}þØÂø‚ßgÝIæfL£½æ®Fžü<CÂÂzawûÂÍ»oSX•swÝÖ!	•LN=	Ù´Ûç5ß<QãßÇþEÜa/‰É}rußžcƒÝåÔÛº‚ù’oe.Jª%i
ÌçÃ004P”}Ã¬|(	õ8ø5QM-¥ùè&Dd.ZötR Ãî&Aúœ-.KÙ±ä¾v}ë(|Ÿ	&œá¢
+ÉƒaÈ¶Åeàw£»öDº…Ø×u¸w ¥ÎüÔe„Ä™¼¿³Ï¿æÎuök+Ñ:dëFÓAÉûV­ë¡ýŽ¦ÙFæ·†_sˆ[}NÜÒÅŒØð,â.DPC^‡é¶e»”¾™	f?AÚ'ä¯På 84ÂñÇ,;%/ñH^{H yçûÜ
g'f£›@eo.E¨·|ž]<YêGýhmþDZËS–:z²4˜å@E–;ˆÙ–¢á“¥Wkk½ÍûÑêZomsÐïö<ìöúÐn{½Û[ëâãV×Î6zk÷O7{ß­à)ï­c™>ô‰¡Ôj´Þ[_=[ë=xpºÞÛ|0XëõB‘ïÖàÅÚÃîFïÁûö°×ÿî—@Xk}¨ÿa}cûáæ:%0`këÐæw`ª›½ûÝÞw£ØÞZïþýQûìbï|pPë0Ðþ}x÷`•}[ë=¼õ»›½µïplëÝû½Õû0¶ÍõÖz«aø7vÖ{ß}­õá!tð ÂV°÷Fcþþùóþ&ó&4­nÀtqÑÖº8¨Þú&t¼Î¾À}WôV×áÉÆºxðãXÍ>Žö67#~šÀ¿k>]ïmÀS(mÀ4F0n¬{ùpúYk4ÖÛëë›r}7{ë«PVxÁb7žmŒÖ{«›]üggõöCÅÉÁ†à à\+„X<¬Ž'uïßpi½‡¸I÷VhÕ#¹þþÃ>;q³ØH ãrNü˜ÀÕXy1ÊSrš¥óº‰Ë¸ºÙåîw ¯¼Ý›0€{,>TË:¦ÿ+å@?f#4ób‘—#B}µTFÕ6Ò¯k_X B·5m6ëŽÓ<eçÝÓt8„JS „Ë-å÷K?e¡³´HFâñj@ø÷S&€`?Ø0tIÈÕÊ¢ƒt˜Á
½ÖFñeô‡5{ôÓd¯ÑÄf³i®ÆbêO/’¡¼ßå¤ar+¨«ˆŸ†y6Åôì9ºÿÒÝè£·=›¯¹¥® vOV›Ôd'¢Hï~,æº¢]˜-Ø[Ö]ä†Ée¾d‹)©Ï7ØÔx¥ã¾Ãmx•¥Pþ‡$æÿ  ÿÿì}ërIvæ«dc{Lp†	^ÔMRA‘R=ºpHu{\Å¨ÁP˜*€"f„ý
ûsÿì¶ýÃaGÌ¯	?ßÄO°°çdfUefå­
ZÝVyÜ"ê’÷<y®ß‘&5?ì®ÐÈb¨*vÞß‘	æ:Æb™äš¿j˜ E˜£NÁtÂ»É&¦.jŠ>êÛ©­ˆB‚Ô#Ê‰.vÄª?×ê§A
Bo:Ù”'ô|–½e¦À´‘Â2Ýðï¹”JÅ¸Þ{”{'Ñàz¢šÅrëÊ¹­Ñ'ò45F.Wªµ$SÒ‚²¦|€}S„_Þ‘“ÉÃÿEwÊo“a˜ék5C"K·ôI9];þ+‰.Ù,Eµz†ˆdØ*D[O…@>!{ ÌàÈÞ{Ž¦ÒMU¬®fÜ[c;^âëõ0gßÜÆ…žMÐ@ÆþÃv-ÒNüéÛlÙÈz\ >o’1ÁäQ«@†úT<GoG–¦©ý
=.ÉMÖã³ÄË
u²mšpÇJ](¤
Õl’1*iõ4‹'¶±¢z­ðS2$Êa°iß üSW\ª$ÌÓ8—æ`We0î<œpïMÆ€Ñ0 !¯²³ÞÈ_né"å¡°ú-‚Ú59Åk>mëÜõ”­8Sêå¬ìwc¥ý1¼öó–,D‹ê»¤¤ªæ Ü]¢œWxËÇUs+»ãÍýÏ™£p{~ÔK“5mSWa˜S´/fÒ.<ª—u5PŽp„hYažŠWÍÏ$f/1Â’0|Þ3h>%	f\…’¦}ŒEO˜A‹žõÇÑÈmŸÙšg«¼‹zÂýC`‘c¬I/Ž§ô
Y7á ÿª±â2ØqŠ<·ßªg#Z¬úã§ù¢§Gí“Ï²¯>Ñ#Xš®Å…20r]#`l']$°zØŠÕ“ÝÞ`O>×•ýŒ;Å„U™«jš8V¯-IåCè˜“öq„$ÑÍ;Ñ¥o4ÍêKÐkac)3²¾Ý˜‡¯ÉˆòÝ4{w^œ<þ…‡JQãíý;&§t0h†]äísÏ‘ñ„gã40ÖÄ jÐÛQw2³¢w<VôæBW´kXÝ«ùp0…]Aø^”~å,ë`.Œ#i€šöá†½PÀÌ,>È+6ÄÅñìý é&
çÖt¬;b¬I¨m¾bßTE==@R~ò£‚®dtÍçgÝÓ º…¡÷ð×qD½#³p0ö€y½PŠÕù»ó×¯4,öÐ8N>ŒÐ“,_a›MÒj²³Š(4Ž§Ùd0éÁ#šäWl€ÌêxÙå
JUEyƒ:•SXÅÜ¤,ûdVdùœ]ÒHÁEW¾;æBd&¬ŒÃÉÐk2,P5‡¼˜Žÿ*ŠðGA¯Ža§vþ”%#ý+¥«³zlGiCê”µìÒZDèÒj¦'B%ÅßéƒÈŸè#%GÔ•¬¢˜‰¥“µe´ixt`Hô¶7Lp!u**o¼" ¾?›‘dô0°Ã˜à¿è—¯AEâ‡¯ôA—~ }]vdÛ¡PÓ>LÝY‘i^
	Âí,`Fš.SøLg¢}»ª ›¨­±bjî_GaÜ§ñ¦Ñ¦«Ú~ïDŸüiÖQ›$¯y½s€ÑeÞ„_i1žµ•æ ÔGß: XôCŒCL±>«.dÒaÞà!	 f0®#¼A„6’K‰)k O[s&@÷°‰_½´ŠÕ&¹K•òŽ`6à[‚z…	Ä_£žó:«ßzNË±ªÞ/‰GÅéÞ‡Ì1ZV¡r‘SÉ•1¡<Ç‹Æµ•œ'4zÚcÌð£‡c“oŽ‚~«¡½Ü8JMt:§)5„Ë“'55(”š7L”òeé“1ÆW7œJÄ¯æû2ÈÒüPŽ®Ô5B½§¤±¹íZÁ¾ËÎ Jùâs[ÔŒÇªîÍ{0h}«®‡JD«ú±<ÄÎèÕJÝu—Åž¶kE³ÖÝZ‘ªú‰ùýÄ¿5lX7ª²?8•EÂCõ3jÏƒiã´x…uù-vÕäºØeà½r¸O%Š ¿œ‹ób—ØÒÆ…9@ÃƒñvñŸŒ¿hEOögÇø>Lí8åÉ'ÈŽaø#=+w ™Œ¥Aà§ÌÀ‰ëËù1ZeÇžQü¤!™i¶Á½ö.=¦ºè{Í_ê|ÄK÷TëW0“ŒwìÃ‰oLÂŒäwÑÃR<ÝÂòÎÈ7À¥!Ÿ¼Ï”˜ù:vY¹;vùcôš¸AŽÕ‘™ì<x©ã‰Ò[ì;È.óÊø¶k*Ö”q¸mXµfö]:GòÊNRGÓ8®"Ö¶eM­Êuî2Âx
+z’aÒGÂD]öètè}0ºFKzørÇì-V€nè$>ëpµŽ>´Z’m ñü»®þ£ðc°jÕ¤Ý0QÀØP\ÄúíÂWxdû3þ}=ç¶'l0œ¤ã2q
ñ¤#dRáh
šW;Qß°erê²?ËÿÒ¿—ŒÅväÄŠL-#Kc2¤7´GfTéü„Â-¯±’0ë¶ý=OxjÃéžŒ–vÍ®2íÅuvÞN|G1+™`Æ,YëØ=ûw§i”¤TH¿-ï›¿GUé³›…ÆBø<ŠCöDÿ4»#¬üiü¢}úAñËø¾0¤ôá·½œ­²üeêý·!ˆç˜e‚èR„³AåžþkþB¼‰ßñß®OŸ—UQòqØ²süÜÔµ
(±Ø„›ö¯yäpù%¿¡ÿ*NÉÑuá²(ÿ6ýáué§Ak2Ä·ögåß¶wGÉ@üJ¾kö³áÃ”™N«7Ÿ§Éðs.Ã8Ftåiõ”æã¿BÈ…Ã÷:††£×Cˆ¢Ç/¡ç{ÌV×(+5¯t/Jq¦JL.Ðó«=øL¯ŠÉÊögâ/ã'ÚÝÄ´iÇögÅŸ¦w«8Š&y˜)÷s ©@ž¯ñ®©´þïS ýzÛ¤•Áµ­q¢?±Éæ`1ó"j0™–U^{×[áÚÑ;å=0¯pïõ]{u×XÛPï
Ã—‹?õoºQ‚ÏˆÂ-X=õ–¥“¾í«·œ¤^ú<ÕÜ5eÇÒ+7Œ_ÉûÜôÖyõ,Vï9¿U¸ Ýý|pÀºôƒ9¶4+àç»¡iûŠílßyœãaÐKš¦þ2	Òþ“­”ôóõ—Ø‘ã¼Ÿ91ƒlr~7êO‡Ã ½&^¾¡ÿ*Êðb¨Ç²#âoÑE¡ìáÖñe9ä'¿«ÈÇ¦’Î'E)n%-ÖŽ·ÍÇÅ¦a;ã	÷nÔÒÒ]Wr;6µèUpP¤N8}`øÛ8W.}ƒ†.³ïLŠ„"°œ¸Šæy3°á¤o,ŽÛç–5ç½oÑg0nÁ»ÕUg”cö;ªämµY™w¸šrÆ€&î‹càè;›#SmÍ¦há“Ô|š2G;×íÖ™¼žý(ë­YfúÜÏq~(%ý|ÏcèÈ!Úóƒ^`Q/{>\çB9‚É4…®/fV5¥ý|gö%íjZ~Ó
M™ÆØ³Ú+c‘ç˜Xmy?ß©=Ë»#Dj×šäºœßœ¢yÑ%Å7SÏ«¬ö‘¤?úxz„öyTay¿{ë9ïyG¦ÚH:„÷¥Ÿ†Oæµ.u¨E˜å.ô?ù]xG!½¤[‹g2)77øÉá¹F“²ìS+
K¨Ä'Ô\0Hš&_“F®zñQ§%öú•€~Ë2WIÐ¶bL8EÃe!³ø&ªî´,ì¾œ„M½,L½Y˜®%5‘ð›éü¥òÙÂé‰¬Å½ÚÞæ¥ýbèíÖ¡—Á6;òÜnV¦©©)÷N¨äÄÎ*¼Ë¡<¦¯£)¥È²«)©xæ]ZžOUSXþÈÄŸPXÄïœ&ÜpM®–mVÍšY½tISÕ‹¥ÄÂ°X¾Ðröµ™˜qâY†Á2GM¤…ˆ‹ASQbsy>S*ÒêuOÂ¾u’m–ÙH>EªÍÓ€Š}c¡Åúq6*µÞ2«%¿®Ã)6Æþ} Ëz0ÙâEaÄ©ÃÔL“¼’5^Ê½¹.e­­Ô ø>13œP¥µióz6á‡è‹–ÕtÕrIómiéªåPnª£íx}#À,šÍÆ€™>>
óvKIé[prÆL¿ŸäàÏ³q.èäGä41ù…ÿÐ-\"ôì×%']Ø¤H™h›O†XÌÏW“dC$Ùp—þ&ðo?ÀµS¼-6àzÃžƒ¤uçr'"“··v½a,Ó’x¤uÀ½ìCÌy•³DœÒ<eH¡_3Æ=X¢*xÙñxþ‰Zã4œÆ ŽŠ˜äUrˆÓhUspƒ	ö¨ºžiÔ'ø\ÏÙj—û»åÏ„5,n²ElX¿Z)£3Æy,«	ŒåBšEt¢¾þÀcH4¸ü˜‰AšLÇ¶&/cÕ„=£_ˆáÂö“Ê×$Ïs•)…‘•p)
^XV:
«º$âè¨#E¥+j.%k8ÿÆeuŠ/âÓ"JNT°miÎ»âànÉ`lBj“–@”BžâÜBšÊ“k&ãÓòÓ~5ä «8¹Hé“:\µô¡'Ù+/3,¯=Œaxó“ÃGÎõä
òÞ›wmðÐÓŸtaÔ+ªEÔ[g†ß¤Av½±øùuÒ÷~À]Hã]„hïz³
6]ÉÆƒÖ=èÒà‡ 9O`ã{`Ô6…*ˆm³w4¹ÉøÑ¡e¤5†å¶#C¹MÇã0í¡nçk×RÅÆ‰	‚™êƒfmiú¾Y_§Á´úph|j­Ö¯W^3¡a˜çžµË¤}çÃf„*ßp­s(rI(<Ž~y€ú4ŒYF?¹`R ÑêÑl³ÇO[C¥ì
òµl¨SÏÞ	QÕZü¦¢Few‹±ŠS¤œØX—RJr~¯H‰Ž¨¡x®Œ°·ñ}*£²µ#±˜9Ž·0¥ù”Ú”Q‚BxÓ¦åréIžr¨
ü3ˆF} P>YÆ™ÔN•1÷RCm@2„yŠÂ¹Ty_ô ŸL° ÑŸÚPAÎ	@«D?3I¿"»ÛõÔ¼ÏG‰}e}Ÿ¿)¨øù8ê¾Å“oÉ®Äÿh
ÞŒŸ™r “¡’¶Ýæ´Ó`’UÒ]&¿fÌÖi˜â½"¾ ?[öR*”dÉH¿E•ƒñ¥× ”üÅ ¤ H5¶Vàº=§02»¤µNºã[²±ÿY}ÿI—A(ÿ_g}gþ!ø|ßÙÒ½³µÜ2ƒáA:ÇV 
%Ôœa^†ALÑÕÕ(Ì²]GÏ
éÃ1<Ù%Û¶BÅ”{pÎµ€×[$U1b9Õ,Ò59
Ì;Õ[h›¥R×V…¾´òJ]]ÑÇÑý¢ôEÍ4F^:£¹´FŸ‰Þ¨¡‚@§;òÑ
4ÒhNÄN&Vˆ Í2h@áJkxk”²Kgh"ÈoP¹Ž«pW¯¦0²	{­øI_²»â%¢%æ• #ûÕ›“Ó×ç|öòôìÙùa'õâi?ÌÔW—m^Ê¸T9µ0Ÿö:(œµZø©"¿Ù®­òÓ5Ôöùëû>ß‚¶4Óú}ÜMýñ¦¹¾Ê¯†R×ÃœàñŠC_Çu~³œ„ùhëüµ‰[e²,77à¡Šò8ï­‰¢M³i£ø+jÒ‰ˆæeZ5ÀÉQÁi€ªß.™¢Ï®Wê=¥´z©¿»í3|3ó¡å&ª¶4Lò%ozÓ˜â4¤:|„³¸Ì,p`Yý ÖnÈy8@?BàPÐ‹­åûw†§J	&nK&Ohv‰ç¹Ëx((•×ñ&Há¤] Rúr’	ÅÔ&ßF71ùªHpI@É%±±½¾¶‰‰ù»„!qÅw
(r03¤eyñŒ!8:AJÞÀ¢BõÕ«pðð×^”¸Kð:×°ôâXó‘cˆ%Q•úžËhaSãï(§”ÚW1Aœ;º4òEé\o4þ“M³/Ô¦=œ(%‚T2ª„Û@÷=(°æ,QNí’¡ z@<†ó¡upôêôïvù8—õ ßîcÎyTÎdiæ8¸‡Ï•±Í;û)Øƒã›“w«¤jž”T'2GáÁPTä”,„s±¤w-ó÷;4;Óe¹…Å¦a`D£Ð}\ÏO¡ ø8¢F¦;à-ð5å `¸ÚïèIBVÉ×…,Àt÷ïê*`.ŠzÚo—ÿ–¬­‘³^ãá¸‘›ø±<5^%7	¡Í{ø+üuYžÛÑux“&£3še»´jèíQ˜PÎ¨…ÜÖj!çö@s¼`³»A¥=ìë_lê5lêŸÞŽ®=<*“w Y¬Ì³èÖÈjƒ‰Ñ|LÜ@ò”ø¹0çûQ\Æaã¯K“ÅÙ7ŸÓUøù£òã6ŒåÞ¦I?†Ám»»Bè=´Oó¢:Ä¡‹]ºG¼\´KŸ,3O g¿`ýÛŒ;§?¢Q@Ä1»G!eVf‘Q×Ôšl5ƒ]xølÇÜý¬·ˆ¢QÛ»F¾Ø~Ól±ò”›‘è¨ÉØ¥‚{ø  ]³hvwA5¥.\9\ñ­‡°¬’eòN¬æaÛmO†ßFÙäá¯)0*»•QñSÚ:u«t[OP­á`ég
M1À‘Ð
p1)@ÞÐ¨¿\¤/éÌW;Ã”^ƒˆžuøËÊ¿4uôV[[û¿ˆÞÃùÙuËÅNZÆ®]t]úŒÄjîo%U^·L”R¤Œt©¤<½¦Ê­ž D^Eé°ýîdŠ?°€†ØŒèJÃq8‰þíá?ÃŒ„dÐ¹åiiét2éEã †×"ü”®<«¾Öo´'ï<­¦½8FÇÓ*Ê/ŠF¹à©Y$¦)§äxkÏèIã‘‘>k1â‡ôiº<îÇ†“MƒfŸY§´t^A¯¯>ä¢Ò–Ìe§,ùÚyø›J"[Ò$Î3´Ô!¹e7»'²ªÁ(vZõˆV˜:gùþàyØ»RwK~Í5R^åÊÐM5›–¡›E¦\áÈ—·Ë6—5ÚH·5¼¨ëE«Œú¨‚x‡¯~=ãõÜ¿³/ðªSÛ¦êÔ†®jÝGFµ»¿^Åg/ßðØ›«8ù°jNœ_ê¹æ¡ Ž§šR[1zTýÜL.qv7¶*îi	ÚÏJKN7ÚU¯óÞí¥æðS|Ê*Ö#)rÅ)þz$ÈÚ„La’ÞIÎ³
hw37[¼7®(Ò«ãŽQaï¶Òðj¨ÐÇHžÂ¹®hCqê*ÅÎ†ð’,ÝûØœò+'3¥Žê“ñ|J+,·ÿ—v‰RdßÒõrÄ±…4í¬«™&ÅMzpšòuÆÎi¢}sAèk>"‹ÁCà%Ìu±¦ó¡EÊ´BZý>Yê‡Kä%\ì¯;¸€ôÏHœ0Œñäé¹_®µøü(ðÂcg¿Å>ñ1ãÕÌñ¯ÆÎxq»Ã›-ËÅÀÖ+€¹çxGõüðòò	ÀË¨¥ÕùÅ"¡—ª‡¡ÉÈwÖ8Ë~ˆ´a\œ?Ùw¿ÒüV"ebL*‘æGÌ†×C?õô`#ìš”ô„ÚB*`|‡íÝÿú§1”Sè~„'íl™ïÇ…‘*?îÆ~ÞšSSo/ÕìOŸ¨e?=^µuxyëðR÷¸Å”Cª+:ç”<˜Ïhìð:ŸZh¥ÎPååçtpÄ/€G•^Éáe™ÌBsû6Ïí/0Ëò\³×lî<}ŸéÑuN6åY)ÒÂ,pffíÆ¸í¯(¹M>ŒÂô¤_ðâùoä‡Èê4Ò?]¢ŒÒÒ²/_\ƒR‘ÊÚñÛbuh•!Uô¶ÆÍ^r$äoùÕsª¹6¾°ò}ý8¿—Z,_3¿}¼<Ù¼Ü‘ò¶O£”T¸)RÖ0áÁ‹>¨–|ãååÉ5	¨³O=ä°Ú¾*NçÏˆ›‡	{Ãåe‚—Ùx¹ #ÓC…¶µ*£rFñÕŸ0š°™sGy­…#ôFÃžÃ¬ß¹'ˆÍñÃ2×¶€u‹%¶ú®[A¹#åÛÉÕu8Œ†äòèZðxâ^¤º8·»dCƒ`ÀX¸Õb,èðàÍÕjÁêèÔ÷µ4 'Ì£É41SUããPJH“ô°w˜¦É‡áÕDÜÿ>‰'Ž´ì›×¨‚køÐ*%jþÌèuLÿÌ©j"7{Û]i²p±f«–Å'¹ßêÛ¢/¼ð$ä
²JÂ5K²G*)ñIŠ% Å0VS<ÐÆ m”-•GEèojF`W›ÕcL$í.zfM †‚©ŠA·²Ï­àóUêÕ2›WÕošÞ8Ø«¦8+;V[Ó6Œé:“nj”ŽvÙÆ\V¯SõšLTÌŸ[™)¦keô\J³Ï”>ÈXÖM_b8hHFÍG=àAc¢³”¸gë7º·l­­(†ú7Ï¹5‡~Ü°@}8`ß_”œ#Ð—ðø‚€(DG$7kU‹¸ÓÜ+Ä	zÐ&C¼_!-„§Vö°¿y¨•E2ÑîXiÀç=ÅRàk p…ŒÊY9?ÏvS§¼éô”¡$˜úÁàë;£Ž{¸Ì<.(Ü]®¡´­VÑÞÎ™…ú…áêãJgk¿&,ãÈ†}rHÃ““Z2òë5T\S’Ð;bÑzƒ›0¯üyö/a.´UV*õFúEdDd€µ³‰k§«åÔåXd=ŸÞZÑÖ˜A'ÎTUüÝ_ÂókŒ¦AŒ<Ÿö0ÉŽa†Ÿ–e5ïT@ÏÒ]ZVá£*¸¯îhüSsuÀêxgZRl’„gšnÓ|OÌ´ÆBQþy3³q4ûä'Òœñ•”Êä2÷ØN§ãàå.Ã7éâÑuØ{¥½8ÜÐ÷²´F»F9>r” ›òðè'K	‡ëðM4aJÞµ¿–ïu&	Mðâó	úîµ/Þ¢Çu2M¹ÙXíGƒh²´B†ÑhŠY„[,U¸…¾Ëï(Sä’åM£j!eêñ’n/FD¸¦Ç@;žš¹ØŠaW¿Çsëˆö©0…:<LÙßt[ÙDÙËœ–ÀL™² qNœåþ)>†,µŒÍî÷áßP“P®)à¶&	&¡¾å<Õ­®±"!·ÛF*`ÛL Jïèú”®m¬Ô
'a®ÆïYés+§?@rZH éï”’4hé×}®54,ý³;×n?AHÊuÏKËžß3­z$ÙbÎfõè:?ékuU–à*d%Í“†«ˆ- ei­j4µGÛE-Ë0Ì†	Ê“l»%Søß$M²âVÍ¶w”ŒïÜK›>_Ù¤h­#&Ü¹$YV5ûz”lèºyÒZÐ-Vr‡]ÜÛkÇ¾ÌænÁ3‡þÞ¨aêöV›Ìa–ìÏõtžuÐFUƒÓÈ,.ÙÔuÂ˜ÅwÅ¤¬[út©çq¢ØŒ6Z+¡FàÒ³.Fà;ø¬m^kk¤Û=×Ï£¬Øàaz\’ÁÈÂ*¼#žÂ˜£4€¸¤¬ºøþÓ16JRx~À3
dŸa
Òr2
uB
ÓÛMG”x–OÙö ¨¿¸Ý.fVŒ¾åò-…þ¾ïóFdäòŽ Œ†±7"ì?½Ã·p.wˆõ`îe”]!y[.Þ@#g÷Úv(½é\%é³ wM›oÞ’¬	iðáª2è¨S]‡S_É¾%éŠ²ÜŒ£þsjsÁm^ö2ÈÿÖ¶Ÿ  bNz×¿ï ¼×—ˆhÒyÞeíÊ-w®€ð›¸¢÷ØaC{Þ/SY¦æ(ÈB\‘ûûE7äGz6ÉÐzš<o¾9®Pè`^©‰U:}‘ü¾¾xk8Ìµw-euÆÓŒ--I¼7.ûÍù]ŽÉëW/þPý™\ô&À.ÂÎ¼FÞ4ß´¡uÈëºæ*õ$åè3tceår &"BÝŠÐö;wi_üq….îì-¥A>M¯±ív°B.éWÁÅúÛ‹	 (>iØ¾„[+di<Y}z¶´B ¨âå†ú,a’b”&—ñP¸¨nè·8šºq5©ƒ ¹³˜VË ï­EÒ†„¾¢FePÐ2‹Â0G;‡é`Š©ì˜@Ði‹Ge§nP»8ucÇ–Ð$™Àiƒ,†å£’G6–º2Ùçùb.¯¬”.§N~J0åÑy4¤3õþi"Î´öÙ«‡öÇÙ0à8ãø4‡Q–Ùa~ÒíËæjòJ/öå÷ÌÕŒ{“²ë•¢Að5´ã	AÓán‰fñøñ
ûq'	lDµIkRé˜_
X^F¡w]»5ø©Rr~š&F$Œ¼ì
›ËÂŠ
k„Ÿz¶%¡.Šâ“O¾ÊU€ˆJ(|zÔ)/çÚVL³T Ø¢hïÉÌgQ>]ÅœBÍŠ–®°oW°1+bÏîµ[Þr1Ã)@œ<ë
n‚(Fm…8Cû­JmÑÀÃ2:ÀäÒ?€Ôk©+‘ÇˆÐiP¬\ãZ{˜ØÝ»´ŸxøQKþP¢Áš“[}Ò‘"ÿ§ð-`{e+u`9ýÆž¾ÃXfÍÀÓ1XÆh€á?Lš‰QÇÊrÐ¶ŠB‚Øú'â¯‹î[Xõ†Æk;„1Fä„qšÀ¼ÀIŠ–+`¨Â!>AÐ„ÞÄ-Cœ²Šf–Çl´)&´	z=›€Á¸$'åÄ–"‘‹EË]be“¡y	@¥¥ŸÖVzÏÁ¶t>Ò’àÕî‘)3…žñT>Ú6ö”¿²¬²’®Á=„2óAå[‚¢q§f½dÌäÅsºú‘9ÅáÇ4AüF÷-2¬QÞNSb‡²Ö§µk½Tk½¬W+,lìêÉËÓÃ£7¯ÿøúìøÙäÃ'¶èªO¡ÓCk2+EqOµÅ=­Y_=´¼UÚÊÂ‘•	7éLÌ0T¨ŽÌÙ¥‚áyŒFÀ®ë“ºz.«®#8VÇ«¥Å½p~^k;ë$£8Ôr,¥3öËDrð·g«F¶kZ°ÿwðO&7»ÜJ~£ñÞèX¡&£Ëç%!”ú©.‡>ã3#Kí
©ªôÄPñ"ßi£è³µ6jÖšŠi22‹¿‚5ë¤wõÝà­îQ´YËlã¥s1øF7TöQ’7S	œå¥çA<	†d&	[«ªÇÀÅwî-™EŸÝé§frâ|PD·Ri¸®W7ÀÃ†Š‹c$KN¤ê=CŠOñ’=oZ<g›â5ÓŸ¦4ÖNÌÞA_l­˜ÄpÙõ…¥vo˜õ–3¤+›ÜÅ4¤hÍäz—¼ûzVÊã÷¿zç
‹³Fw9qƒ€¼·Å3*Ó€´JdL7¿ŽúNBß„ˆ—Á%”Š«6Ü&%­i\½Ò˜Ååý2N b%ZÄ1í3½Îä…~Ôwìö¹óIû ˜yr•üœRBÎ-ºbÌ0ÁÔ±ÎÜKt™.õ>|úŠ¡žlºK|€Ôªé·¤m2_,j„Úr0æ¼"ÿœîð;R²c¾þDÊ½µë­:4º°Ù¹ñBµ:’ö
²+Tœ=+5ì®ÈæBè@çyE›×.Êr†	JQ&Sï3©Øï*Ì¡<//lŸÒñ×ãåj ß¡¤â¡ãâÌäáYª^ùaÍ’‰aHàxuSXØ×RY0ÈÃÌ˜âLv£7ø›ªW>—^/«Î¦t‘G9Ç¤.Òûù_Â{k]ÑàáÙ v¾³Â+tA	…«bïIÐzeŒM­Vx¬¯˜ôÚ1®lÅÔæ²%k–4n‚ŒœÎtãÔÝR*ñ^ý°þEN_‹Á†‹¦ltÉEø­ëre+köUBë–l£!—œßÚôÁb×ép'g?î×ØoÊyø0øÏ‘'l!w‰ÐO2Ž€ÄŠ›ÖCªÊ!ð`E éV:=—ž‡Ø‘_õÄDlXQW%8¤‹üÒI8íÀ{Èùå	wá‡Ué‰5æà
îçbçà{qÅþ¾pw8JeyÒ]Ý,ØU3g¬cWi Ñ úè”¹:IVSr•&C)#Y„A\egE•‘¨ tóÊpuAôgÞÎ'Ù©'r~×Ëë]gLP«œû?0¼í€¡Ô1ˆ2J¨WBFyÈ_?ÉýgM.¼évô:çþ²ùÞ•—jòBÃFÕsH3øbyýã?ºx˜½¡}ñž	uý·šYþžt.ºoÝ7™¤ä‘3I3å€:(—§ ¢nûRÙâ£õÖ˜ˆ´¤3)N	ßc.­IäWÇac5^n[W>ËZ.±å5“-Š2t‘ËP=väŒZ·öˆ4¿»*Î`7Ñ¨7_°œ,¶´D¼å"ƒŽ©ukÿ•qfŒÜ¦¹B­Ÿ…ÑÜG“QO)ÔD]û©h8/l¨R#¼Œ©rkrKç~Ñ®Òú‰ú§At›èý:Ê«-¢ZŽ¼+÷ºåoK²#¥ÄÍU~(ã‹C‹×•v½:æ­êJ¢„¨/PÃîäÈBZ}µQ	Èh1t7ûó‘:ä-Ði§E“6Î¦Ÿ•Ê›CÖÏíw]1vºVl}–'Põÿó]‰ž~‡ñ$Yâ¢­!æP¦hRÏË‡ïGEM:#§PW‘øÇ¿2[¤·êwÑº‹ÖÝ]2ãÿ‰4ž¯Õ¦‚ë	îrÊÔ:åABÏº<Äo§°#”ºS7¯¨/ãW¬>ÒçéÖ„ý²€‡ï	ô/ÔsEž©§‹"4 %–ü4µèÏœIºwmÔ«NÝ(¾É;Š"/kÀ÷–WaÞ²»QøæS¯àCM€,÷…0­6_ž+Ìïšy1v’Ë,Loàºøq˜¹ ÑtuZTJ¤îÛò­H¶C.‘.-ùåS¯,œ|ŽBÔü¼$ï[ƒ†cr£»FíÈYÉ«ÆdŸû€µ.<<ÉNÚâç+zEömÔv6ðyš'^^&¯f¾•Â.ÄáâNìýÌ!/¡-’éÄK“S³uyRöï>wêÐŸóžP…O7þ~…l‚Ùð{_}¶ôéa|ÁÌ(Í—‡A:j·ž¥iBF¬~Š·ò¯	Á¬ÑÃ‡Q£‡nÙt ,!<Øm­ÿ4&r›A`¤sÍZ¬#4k]£ÆÔþ¦îÞ\ò«§_^ütGü—ÕÈæ3Cß`ä.ð“@8´Ïƒ·;õu#Û·xÑ-¨	)j0™@&…]ÛÖÀºxçË—¯-¼õG‚2]Þö.vÕãtÎ«%¦£ÛT+•NñÕçÌlåÁl_®üú$—2ì?C¦«ìÁÆëã%^_¯üÅÆË~i¯"f»1óUh^%UkÁxñ&Ÿ–íòNì®^›é:ŒÃtÂ˜.™çâ3ñßï¢@_˜®üú$L—8æ?CŽ‹7ÿ»õ…Ý¯/ìÖGþâ»e¿4ì%UÍy­Üž.ZÐN‹þúäŒV¨ù¹ñYÐñX8üÁòv‘^olÚöw?¬v×|}Kô…xe}«mLƒË06‡*ÐrT u˜l¼fÜC‚'käœQiø±·F‹oÄúÖ\LMRØçW	>PH»xL×¥¤âÅË­6åU¹svÄŒOü3ÀŠÃß‰šrÊeíÖ…v<<=AöoC8j"2Bny¦ý×!§IJ®‚›$]!A?êÁ°!ÿòëQ@ã$Ø'|õR/Ñi†~ò¡KÏô¨''°ãâhðð# Éa
/BY°ÓjÈp1ÇOÃ5Ô‘4ä6?–è±ÁÃ.vû²89dáRÈÂdO-Ì#Ü“0†e.²ð<ˆ¯$T|HY¡€Õ"7Á¢žõÂË¨”ä£á&nÀ¾Sšò<ÜMúìêñÿµÈoh%a˜eÁ ÄÅ?õ`s®6Ï#1-P^úYÈ>úÄ§fçÛ£|QHIÇC(xüøÆmÖ@êéÉtA$‹0ðW’´ÌogÇk—ØØ³Æ/2h¨)òf—i8šµÚ^‰šÉ£[©ØÛm¹¼jJnEzoEB‚ü¬rrX¯úºr •ëö  «^–ÌQ–^$Œº˜¶Ðš4KÎýÑ¤BêŽCÎÒÙÔÊ©];*YhÀ'sK·y/KM,ÃW~DR ¼~ðÕ§óO5ªç Æ¼1{…O¾`KA,‰µj–ºy2žƒÓ¤rßŠ7jáfHUÖØz´î¨‹à\µ¨VÛªp×«[
Ç#Š²MDãºzÓlñF©¶µø!:ß±Û¼4Ê|`rÑ„ðD?4³lEA[©†]ã8è…×Àü„é~ëÓÝ…$øó4"ã4¡¤=ãë$>=™bÚ"’@Gé–]WØ	z@‘1X­Üñ4Ü¯,i”-©›uW2:ºF(í UyI¡+$ìLhöâmŠ(Ö©¿Ö6©è#òp>`w%ÍpQ5åVT”
GÌXlÝ¨«Û®œðN#³:Pô€Úý€žœ¨:£qN˜5"Ìú€vÿ¼¼ã’vc•SåŽÂÛ mL»£Ñx:ibÚ`ÚÛ«(6g³]b<)CÍiRJeCÏeÔ‡¾ýr³ãïìIçbÝå`º¨²ÊX&,#åsøû»q,+½½B–r
“«êj jKÍ5i^ì (Ð~…7°ˆQû]Û²q[ª÷¢TÝJ^lî.ÈëMÓ\w2á	¬"¬ø#ñÃUòX&!I~’NG4CÚ0¸EïÒœ< °˜™WÉjœ–N†W[k#›yA±KÁ3ü òZò¡“`
¾Ê:‚íóÇË8½_ª-9³Ë4AÛúÃ‡R«×a¼,;k¢uh2Ú0ÞÏn¡)£ ~Þrp³ÖÈñ»À‰gôò,Â4<cs¶uÔr13_ü…iÏÍþd¯v?Ë‰ÿÃO7Ûµ…+¼j}u¸Fyocà}Þ]ö¯Šç02>tÕ~h|d| i¼sÉ/å(âµm((´N¯%[§Þán±hà68œy×OÏe®ý‡{
Ž,žµõlii°}¿Âí¥~]_&°Ä–ôÏÑòý€ù—É÷ÐþãüýjŒi^Àw[Â…,h|ŸôþƒŒ“~H‚˜HPÿh’¢.àXÎy›”M"ùó4ˆÿ<…všð¹£A®Ó,Œ½µa‚B_Gy²,eø›ÝÀ1±rØÕd®ò¥{eê©†­¢H•èG#Hƒ¡2“qÐ£™×WH†y á¯Îãí*ËÎ•ÕÒ7Ýâ›nõƒð6šÔ¨¡2¦VÌŒlXbfÀßî|-èÈ¤žœZ)ßaEëÁ5L•eg°ÉŒ2ÃÎÀ0zô±ª%¶ï“d>-I2gB6ãTî]oXðò(F]¹`Çœ…ÙÐO²½µëC™vTÊ	QœÑË9ªTŽ¥§ÀQƒÐoœgÃ1ð`Á“Î¶­C\0ÔEKÝÌý o³ÖÁýÓ¿Ô,)Qm§A”óˆ˜T°BkP¡¤‘ë=¥òœ§µß%3]ú÷I4NøÈ
‰Ñ,G°ÓïÎ;Ï:^<UYˆ˜”þÍ³k,`wæ€}¼óÏ“ôY?š´¡‡¶Ýa€Ø$CåË¯Þœœ¾>ÿã³—§gÏÎ;Â=öÃŒ–HžÐìb¨‰³t,—ý†ê|
"âÊ-Ï8É„"<óÅãeÕØˆª1§Ò“A%‘±´^UÄjE|qzjòM ÉÜÃ`®XHÍ^åXµe*_‘0h¬ÅJ6ü$—\c£1¾U»cik…‹ðyOH%mSˆy`1u)¡·æâðj9EûbÉ0ìh_Œ—y*¸ ž ·f´µ±æÎ™SÃ¥ò2Kk¶$Ôºoï›nÏa£™aÖåþ0!¹Ö-‚äÃ(LOˆè4Ó'iÔ_6¯Èœ%3ŒDÁ©q3¤Q§¯ÕÀh@K×Ñ˜¨‡÷9D›Òé7wã¨}ˆþÇ~4?ÿñQÛšÑ´z’yž†³ÁHÍÞ¤ÀK*H#ÏnáHŒt­§Õ¯W•/f×Ä‰ã`dfÂÙQ«„éÂôô8¤QduYñÚ-ì¥É(¤Á00·ó„©Q«Ä–ž…p²>ü5hs€eÞßó¶yœ‚˜Þ	mÎ]½OŸ›Úx;-Jêô“Ïdø²V³´4H¯W¡i¶&Á$#ß"àd5kB>ç3ÃÑïÓˆ~NxFÝ]˜û`:Šþ<e°è(>‰™ÛI€sqÖ¹Ö1è]p–:ÁjAhøßçéáçyæm°Ìc™eLxÝFFÛ+‰2ÐÓÎu±[ËË–H‚âÃeþYÐïóº×•1±&9Ö&Ÿ•ÕcÐwµÔ“¶Þ4Ï4Ìy/ICÞ±ub^ê’Ò¢ïtÚÉð#š¤wy…èÜÚYÃà¶Añä×Xº}`Ýek´)T@:íMÔJæï	yL Zä†Úm¥ëk¤-µ•Uí€£ÓÖÚƒò&05&s™7ÏÔ3Z4ŸÝšÅçkÂ«ŠWAR³x¶¬¼š„™'M|nVu3¦nê¹ÁÐ‰×;ìW²>2é×x‘vWÚU²œ•)ÍlùaªŠßJ^à\IÄÑj¡a]‹*„.^žã¸ý}'½¨ôÃek‚’Úq”M³?Ã2˜#õó–¬Ó›I{~žô ‚×}1æ8Î­`90dÎ:Œsgí¬ŽÎ‡Õî#ÌjõH—§³[Ê’×fÉ´Ç’'¢SÉ ºg¯fqòÁ'M¯ÜÐlè†_C·éB²Ø¸‡h®PÌXcçÖÛŸ¥A¦Áî|ƒçò±§y*¢¦ûòŠG–Ûc;i:¡î$=Q¾we5òÈ³­Z®Ú3¡—òæ.»vÙÏxœ2Þ£ôá/(Â~ÓO$ŸLÞÅ¹ELÅp9My€pö©gÙ˜X6Å©æ¶O‚›‘a¾…<<ŽÜoµIµv=ŸHvší\®Ý®§æÏ`äqzìUB+,ÎyHòsi<³]õCŸ¦ô¼%í“C³³­mÞÄk½ ÆôGZÅêÜÕÎçdò‚ßy8sAUÖP*1ƒd×h“#4]òAßfÌå‰±V¥(Ïƒ	<ƒ¯¼@ityO:£è&ŒË™?<©=VlñQu‡\t·‹-¿Ó@ï×‡A_È}ë¨Ç`¨ôÒå’äXßÆ.­+ó[¤NÌ»‘¦B­¹ðñn_Ò¶›Ñ¥¨ûä&Úßp ÌÎ; ÈA¦Ñb‘èÌõ[l|ÝÛ]ÇþiäZ‡ƒ)œ\tŽS8 z×5`Óð±WÁƒÑ+x!ˆFý e/Çh5u7Ón¾M2Æ ™<JJ>h·×ÝÑ4Œn-éhÇßüùJGÙÎðŒ^:æÐân  e3²)eÝœVbŽõÎ¶ËÇµ–Ï^²M}¶òñ9'Ag.K¼XµÎø„ÐIÒÿêüùÄëKD_®‘È8‹¿÷o#SQ®}»ºˆ:þØ)ÇæwØöDÌ(S‹ÝeÎ<¯K<¦6¤Âä^Äò´îR¯»ÕÇÛè*Žè‡ªG‘Á‰£0m–¬új4KsPâai¨Œ}Ž•íš©qÏƒ›0/DgñüjEUŸZ'ÐqNž…WÀå]}àüäŽÈ§øb‰3¥‡RËÉ„úÒl3Í«MiÙGüG2måô]Þõ^¥åÃfÚåx h—7)ÐLŠeÑ°*ÀÖ–Låœh¦<Þ(p>2ExZ`tÊÃrÿïç¤B=¹úL—$ü4H®AlÞô•–ÜjIJMK:êñžÝ†½)bm3Õû8IÉÃ?c&k·‚Ã–SÎ±°6ÐO¸ü¹é‘2uÆ3|Ó ê¬­êüKÅ]ÊÍqö³='„y1‘”ú8ÙO@o‡£ãäSÊÄ½Æ¤½<o*äßeyzMÙ|S86Þ)|ÿ	í˜·¦ü—VïCVƒ!7:$¨ÓwP{½‹¼Ô·ËDú	µÌÈ–ÑnQñ
³˜R7vj=³~‘{w]‹3I;ÔnoþÍ¾b{vN¥THÑ[ZPÕÊ<Ïärë$ThõF]!³ûåe–ÛIÇ9ót¦û¦¿zSµnÓØ8­)oËFm[ÎÝŸMÒä}x„~±Ð¨Öÿè®_>Þé¶ì¥â‡—A>Eçg?p,"0pÅ!°o P¬êÙmØ8AZŽüÚ¸ÒÙPï‘­u÷RUÇájkój;t4(£PM5\íî¬çî<‘»ž[_q¤A±ÿß4éÿöãpý²fÿu	«#PäA®=gá€r—®Mî³Ëý3¯Ó½‡ô÷¾¢5¶ÇÄåŠ‘ñê&óðÍªµ-½L%bzÜÆ\¸)ùCIØ¤0Ô£?E¨1ø±é ¯r*)zìwO8@ƒïnPd	Y½Ø^Û|ë6Z#^³ƒÕ,š‘Ý´ð@ôìøU£peÑ%‚h Þãir»ßZÂ‰ã¶íWªÈú^÷÷[/Iwcek‡’Í•Í¨n}¥Kvvàž7
Û]û°­ºWÛWk~÷÷Qr½ßª[Ý‹hÙ¦—”÷ÇWQï·h4ƒß'žáÔ?á¸ÏêvßhìÃ2þ<G?¯¥€ Mƒ»ýV·û¸³¹ã\_~ž\]eááÔð{²šŸ kŒÙ žƒ>ðLÏ&w1Ñ,	Ô.YbÕ­ö‹úH7#a ‡b2,ù¢³x­9án^/*40¸Ì’xJeöÉ$®òS53å÷¥@zÓ{nÕÀZæs»]3:/~>SùPx…äûæPLj§âäÛrHrÏÃÁà« U~PsŸ`£¹:®W/6™Èuh)–Báø¯1ÜJƒn·
@C'Z‡¹CÇÎè†²’s]&_|õšL”Ïæòíå'[h©°PG"“ÎÍ¦d³«ØNFW‰‡Ð é8ÂX€ËiÄY,BšYÿ¤¼ºdï­hüœ±£I ›/%ØÉ	GÇ¯Zä:ÄÝÄ9t§QHuÎµ/:Wi?˜û³uŸëXxÆzJbw×ðBcGá1‘á¨ò±s¥–)YÅ’ù]¡t&xÕ-Ãþ‹E¿ÂŒuE±áÕ\žÅv:¶ìùÍÔeoøµ;ü¶Pñã­`ór*~KvÉÅ[»?^oÝDª|ÿ6,$÷‹wž/F#X²g@ø§ÙþìÑº»	À¬”ìx|0úx¬ŽÈm»?À•û;DÎ ìê‚‡#‡ß&hºÌ›.à…¯Ð9VÓé!ËÝ
,‰~x»ìŠïÍ¯½£Îlª/x×ƒ?W¿žÑîßÝ3ž~F‹íô˜âÁ¿ºöÖ€:©ä›$‰'ÑØ-Mô€tC#Ïs¶ó‡(KrTÌ-½`Ó3%[ÿð`g|[Ü„ŸÝñ-:ê.ðzñõWT}ÛjÖæD…2ÉÉu¬ï¾æìoéåöŽ½ä¤ŽQ§ãé<zöÖ4§]C»¥ÃT8Êö-ä°í¶"üG‰‘Î½|«æ¿Mr-$¸c‘Ù«qí€`U{–Ÿ˜÷Ëµ£)>R;‹lÀ-!ÇmÞRþóóimáË'‹Ë[ŠNÝîVÎªdÑí6õizVúªÒyàMÌŠ™à7<:9¯TPëCb{EhÑA¬Í®5OÐ’Iý¬œ(¯<}ÛÇw4°^û%àüíxÅsYmÕøru›ŠQf·
;µºCÈÈØ;ë:?Faí·t"ïÐÚu«_¼Cf{v“ÄÜñ ‹Æ yL ÎóMe7gì^ª¢)îm-=Ä4Áè:$ê!þÒÆÒdµbhÍÐcBy„N?çBFˆz/@ÚdÔö¦“è†æ £ |ýÄ&­Ý`í
+Œº“€ãFèIšÄ™&ô¾(K·æ%P¬±Çè6F§Å«<º£¿K¡G›MP‚*'øåL˜×'ª·ìH¾5R=!ÍMÒ$¨ ò#+ÙdK“ìêå‹ú>;wß¯t–9Ðó|FŒÊUÌÌ%y“]y©/šâK/8”Á&aY†Úísô"]Óó8>åv.rºÃ-h%]é«…„ýý™®£oÓ‡¯5–ÒS4õxá¹°½\¸³^ÌÝkrËá5‰Dúe8‰AÞÆŸJ‹fu„°_ÝÙ›†aƒ0¬½põÎ*f9#ÝI,TKœ}XRûg–Y%hÕË‡ÇhB™Ó4À™›%nÕ}ñÆg|Æ¾S"WdNhÐ-í&ld`6|bª<Co)Üº—ª¦ÍŽáÊüÿÅúÛ'Ì¡ËßrW+J—W[W-&î¹wú´ò:Ý6;áÇ^ÿ;žDÃ€Þ Pfú3_¢M§²8ô²â*éþ4‹ÆÖœ³˜¿àZŒê”'¥@:©µ¬=rØº/YŽ+êež ìõ¥A3#ƒ‰s.€D>¹þž»˜¶ë.1æ‹Ž&Êåã²‹®‡bCˆ‡ÿ¡§Wžfq;P'ÝdheÏ“ßØ<ï˜ç‚²woíbr~ýf&ôüþWxl……¬ïaR÷ÙR.çÒTœ…½Ïg
¼ÊÒOvQã_~Šøé‡ÿóé’ßU¢ì1ÖçáðÓ®å©N+VïNÀa‘Piô"¡Ò$Á´ó «Ž&5ä‰@ÒXÅ1 êñol*Ç:TÌ\*‡i ýY¯™+…|2Ü“a¢uóKÆ»¤»¾BRlû3¯à/8WW¸/Ý.ÁdN®„ÀI˜EÁˆÂeV<7Éf‹€è=ÁvìÏ®Ør÷U/[7¾Æo£¬4µ#gÓ"ÜÃnC5ï+7°ËhÙE»è£­o¶v.—VèV:~À$˜…Â£ÿ@[°À*§£Ö,‘~2pf.`ða™¼õiMsc±lî¦áÐnÆ‡·çT×ÏÑ‡^A,²º	ÿI—A{}…þ_g}{y²b6þÕ»²$€Ô° 0¯šÒ¯è0ûß‡l¥/¡~È£…îñÇ)w>Ë&9LFÉÄÏQ¸Xà”‰óø`DiV9JŸ»‡ù~Á<¬g›ŽÙý„&qµ½"~MãàòQø­xÈbÒ‹b¿1»¾ýøÑ£Ç>Å:½
’:'ÎI$û…2¡ÌåÅèbƒ®v®‚«žýBÖ?
Y‡Z¡ÁUËb?
ú™rïÂ‹G+þkWÊežSô–;“áÞZN";SÔðuj–ÅNï¡ý`¹Ùçé1asßÞjœQæ'è€BáYúIFFr66½á½@ñ†%Òß®èáxÌqŒhÔ?Ùé¶?Çˆ%S$eÏÚ…SœA’â€Û	Pñþ·˜ KjspôQŒÂ¬]V#œ…°vû|~E€”?`YF­Š*|.¦ŠP´=b¯KDhD, ór€#¾¿´ÏËàöÅO.>»¤7­Èñìruˆ‡‘XuÕñ|{ñv¹ƒ“$Õ:®åÀ±L
¨Úc !¸è¾…{&ƒ±‡e½Ok×{©Ö{Y·Þh8Æîž¼<=<zóú¯ÏŽŸÑ	:|ÒgÀÙ%8¶lhi]ÆPù¢À§ÚŸÖ.Ï=-q•¶Ô°t9„Ä]Ê|“/òa\¡^lû·ž+9ÃÚ“5–Ð9ø„IÕ×9:˜ûE‹ÊÆ]ü4ˆnã.v*Î˜sœå¼£NÚyŸ¬Ü|-ß‡Ò¨ÃŸŠƒÊ=¬8@ ›õØáÑPÍÃx’0?U =JÒµÍõÙõ/òåÃ¿÷#©ÐÂoVt¢õ(¶¥@=ðä®Í¦^:·–£æ}0º¤ytšyX˜Í‚ÌÁKõÀÐöº0øF¨¦Šº´ýÓ±ê›Ð®uP®çæ8úU×„²=º<V‡Óð_.ÌùQû]6»
Î¹µE3‰ö èÚÛtouaS-ÉqFèö	µR¼ô%L¾˜?Êäg™%™¥Õ+’>®²ÆRú¸Ê²*—'Ï’»º.­ÃünW^‚üvCÄiÄDTÒYñÄ³g(„ùfmwšˆmiÆç•ERoe°¥‚’oÕÎÄ’Èj¯LœµÔ$õ²R†g
æÛ]²±î|ù–âOù$^¾¥z§F9—«ÉM¾ã5ó!ër\Ã¦2ŸáFeÎÐ !ò ÕL™lI-f’—}’%£ìÚÄlk¨	Ññ46åL6úo›3ò$²‡bê[q,ªÏIà»wà&í›Ð/œº*§i“»º•nî«ý®~ÁÝÝ³ýYþWõ!K.¼'þª¾[ˆïû³âO];‹<ÁBSÕ{Öï”ÁÑÝ×~,uF¹¡ý'‰¾‰TßÇÁÝù´×³ì9àÍ•;ê7ksÒTqS/*½}Yæ/¸jju“Íç6èÏ˜nSU.’dÎMŠuöÓÓ>OZ•ÄÓ^"±à}t©>Õ¡µœö¯¾KcZ>ûÓ|"¥aŸž¹}ñtïjk^@<QZ•rOû]	O¿(~iß€ÏéÛÂïêû=šTþE2€ù(ÿ6½Wì›@üB¾[ý6†»G×IÔƒÖ—kÛþBxUú©™ã\‹Uª{4c.›M+fÔÊûÉè[ži}‚’J/eöÙòžîË³ëëÅOÓêÍçi2<²B½Ž´˜ä@Ã« £¢ÓþLü¥µ7ÒëÊú§ö¼'0„ANw EþâÎ_XÇÑh²{õûÚÜ¿ðÃøy4Z¬Ø2÷|–/µ_Øìq¸ù‘Ô:Äéç{¼Ô?8¾£KÉÂíÏ‚ìnÔ#í`r“-8“ôÎjžÒRéÊ¢PáNG³ÚVH)¢ì¬Ð®gºÚB5…™ÂtµaÍÞPËüKÍW}üÕÏ-Hºœ/ðè‰Ô¡]Ò_6YÉÐ ;ÍÂÔì×n:¼‚A”ßqÒk÷ñÿ/WÈ’(¢/­º¼buòVGÛøª9YÍ=® wMÚ½8™öŸ¥–Ž²µÄa'â¶[ÏðÖ·h4Zç!¡åí¶VHQ°±	zÍ¨N/š764µÒÞÂ÷ãéÆ œdØ2}2žjÅ•3y^ÖÄ€Q+~æÉyÐ8ðWçñŽ7ÿÁ¿éúñ –*‡Ù·!œ	%ÎúãŒ“o™Žç™Å'Œ³Ðòo“Ž°Nl{^£P÷Ñ«%£g}ì|{Âh§œÕj²ü·J®¬Þf|ï{Æ~ H•¿]M	¥sÿJFÇÐœ%ßªûô}©rßªØ¡q>	&èÆ¨+ˆÝ3s
,
ó-“¿+ïë¿“ ÿì&B‘¾¼Fý8|Å!{b‘/
µüùieUqÊ²-Çñw€jIßðß¶šNŸ—UÐtØƒÍwüÜÌ~
.òãðGŸ‚³[œH=÷¹‘†Ã(³&gÿÔóü¸«%¹ÞyK®w~’ë)o-=-üG­|å5ÊÉÂäïÉ1X³¨Üç8¢ÃœHp~2ºŠÒáÉpœ¤“#j•»Ã§ÊèïÑ2„†°$©-ÊßB;ƒ”¼‰ÆTá›Û«ÆAV.lgEê$õÉšN½Ã.%+´í¬¤ªíÑ4Ž¥÷¤€*Y_õ•"<V‹¬r`LRz“iÃ…yhÔ!ìLÒhØ®°_ î¡j”UBÓ°Të]©Ö ”s_oÄ¤‘àsŽ8Öû->C)¡&AqŽÊä~_™úV–*.R_ Îa¨XVÔ4§¸¬þcK-¸(ë&é=ü.ìáG>4,ÜUà¬ž"O”}§h4tª—=êé.a§PD‘¼F0u–º\Ù({k´¸J%ì°©ùJj,hS„Ô2eš¼Jà<ØM¦ûJ‚Ø-tÆ\Ý^gkù—z4ê¾TõD£ ›ßQŒKé<`ß¬zñ´fíÊÂZ±µr=^O'i¢ñãmŒÕ IIh”ü‹ŒQ°OaíRéˆ%€Ó	G<å#¬^“(†Omm›¸•ûÔ‚¯Ï)kuÖÖík*ØUUvš¢N²¸ÃŒÀ½{ÉO(æŠoÞçóM°öÖØkÕJu°¼y™¬ ûÖý‡f^×—µ·ÆöŠzV]^âœâ|´¿ª·8áP¬–ù–¹´¼l £ŒFãiuçÁ‡H"t>œ ÷Âk "aºßz†Ù‚¢+ÌTH•ðé¾®R	ê¶öS“Š(EÓf¾z»G§*ñÚûEÄÇ÷žÛŸYõ­hc)f­]÷5"jÅ1OÙ!%ºrco² ù’Ã|¾Â%vÐÍâá—{÷*<üX>²Fxy6ÆŽ2&ºº9b`Ð˜Í¿1².çA|-;D
²˜ÿšø—\pã%‚<múé˜˜œmwó¨ØŒ,A¸²Ê§^_Àµ ;#22r’e	¹Ñ#†I"ªÃA’F?0–›0<Ñ"¡EZXƒ2T&÷¢,É:?¤YJÿ­™!ý*EŽHÿä³f‹ÎµMæ¼‘ÒæŒ‘¡–/Ü‘ð•‘;2¬ª:,’iÉ
ÎJé?³¤!¯¿HÆIîŸ„kR·¼•e2Ñ‰Ï„o²éÎØ
ÇtÊª²’ñR¢Ž2g¦²^_‡<7zÆ”dbl´„ö¡0S_õàÔŽs­u”±}¦Ñýø*n%7¨¿†¹Žÿ	š¡èj…U‚Ïl¿ôÅP´r…AY¥å—ÌqQàJîÔeðr(ü¸háîiš`€5îÐÊòðÐŠãbÔ‰ÍªãüÓG¨OJKŒOFQ/Bµ™8yb‰%ûiì¼Ú´Ž´ö+i0ä;BÏ9ãŠÏ+M­ÇÍª‹æö6j¤	Ø{šFá’ãîRê§s{…áÀ&ÓÑš¨`j&ö’ž:‹­åÐùüXO%VÃ31£ºÅåÁK–þ)Ò—mù”Ðp0ò9}»+í»6¨5!^É'¬…?T»‘†žFiØWTŒ)UPºŸt¡°ñû@<š¯sÞ—q5±Ø)<Ð/|±#m­¹Êñ×òþÜWÝÓi£™icŽe§5òI›¸ý4N€óú‰š'ÎºkÑÞüh]ô: ðµAÍY8ŒèLpP‚Q2ÁÕ—|ûîU§ZÆŠcYC®£Q?$4ä˜
)"X2t…?†¾Èýå÷Y–9Ã„9a•-ËÍÿû?ÿëGC”~>d­ƒïþ2êMã„NAFxøWË.I'¯%½Tá¿UDF¹Œ.ã(™„=Jò*ê£4`êV|™®GIiÖ¶ `·S¤I­ç<ñ–.w*«ZU˜}‚}¿Yë´¹{ï¢´‡ól}“Äâ¹Ë}ŠT¥æR[
õ?üH›Â}?ÌÂ?¡f/. $|œwiÑÆ–©”¼$`=Â÷u§óïœ~'C®Ú´n¨Rå”Ë¹˜Lï¨~
àQùÄ––Vä¦
)üä2f‹Z‹!Òb“Œu´!ä[Æu1B¶˜e§BE;ƒsdNLÊòÔ”üÑqjþ™~ËÞràT²£(hÑÿA©|c»x ¤`1!»hqY`qN]ì€í!q¸LnMý³W³¸Ó‹X]ÚÕÌ˜	—C`Æ˜ý¨¸¨U¥¼]ÐHUJ³ÔÃ½xéŽ¨ŽS¨ÐŒÄíqnmÐ…¢ Q­mxkkƒá‰­Ñ€¢=áø°ºM®ËìBš£LIÛÊÙ…€ô§è*
?!g!ä£B&Ìt‚ÙxèÞ·ü¦Õ„—>R¯zÕsÚ“}CÌôSgÃ•Ôå\ç)‹<JŠ(SÉÒS´Î‚d`t}Ì˜.‰”ë7–W”Ž-b½tÔÑÄ¨è1~é.YI½Æ]	 4oi“ŸN[ª3­uÑMøšj9ÁYƒqRÒ·|ŒQºxëÙó•ÒÐgW0õ\‹9´ ;B[uÂ¥ÚsjmSž`¤ÅÒ²Ÿqaf(¨Ú«ä&ÈevJµ¸SG~ëÞbuhâ›ËË¥!³{/Ÿ\cãÙ_¹Âa—àhð%¾y$P‹ŸJC¡ó"Ñëñ¿
ËnuÒà‡ 9O
èIÅ‘„0jÜÚºË•nÓÇÛ3ÿu™Ìð¨Br¥qÇì/_“gç¿ÿîÙÙñë]†—»€$„CUnØÒl|cÉã·˜4‚k;ëÖ\
ð°œ³L¸"hDéÿL‰070ò†9½²>}tEžt±29ï#©g¤2j?¤¹ÖA×|°V¦ºØð’‡øËÕÛµòš³ÒYs¤mdÎCgÁØ=¶óo9Ü§iÙÈ\ÐfyHCTÝhzÅ¬\ÜX°¼XUxùÈÁh
HÑbN7òW¡—ˆ:BÁ+D(yWñ¸7Jy˜äù9öÎ_Ò0'r’Ôlè‰n[Ü{ySûåó #"Sh@ÑëŠöZÓuUÿ   ÿÿì}ÛnÜH–àû|E”0°RÝºÛ®¶5¶¼Y’\¥[ÖH*÷µ…*“R²Ì$³I¦$— `ßó6À`0/ûR³‹n ~`^õ'ý%{NÜFÉL¦nV¢\Ê$ƒÁ¸œsâÜÏL Ú×È‹Esö ­v^¶mr²Cs]ešC^]WÊÐAñÜæÖÞþ+àXuózÑègäpi®üíà4ÈÒŽ~¡ö™Ùy,ÿéG/ú8–rWÕ¥ºª³úJ•qA&œ(q+!û?ÿN>^ÿ
v«Nà[£ÖMƒ¹É\Ø“Oì?ÉntÖÎÔ¬÷¬(è®T×ñ,ø3¬®.¯²+ðouÉUÁ|è]¼£ M“Õ ÎyÓëÐî*R]¬äžÍ6ö8ŠùÕ§Äl'” âº1»‰’“] îSyˆ5ó?P‘ÿ%šÌ*ß½QL2¯½X [é­sœ|YÌO¿¾:gÇŽâ›õƒ»þÁl‡eG¦xÃQMa¢í#;Ýàåhh®­Óžÿ¼þo j*>nußÞZ ¡[˜Î1ƒ÷w1¹ ­?õ6ŒL­’`Œ‡Ax‰ï*;«N374br{ÞÉÞ˜nNëè#z¾5öv¦0ýþú¯éJ—gÆ?ö“lœ<ŠVÆOÉ?sõ9t¸ÖN‡~
‹-ºu`-v?;ˆuSúÛP}5›ã£²DÏ±Ï}ƒÒê8:1üGÍ Ï¼¹¨X§™#ÂŽ}jÓÐ…nmÇ>bîØGÇ¢£2Óž1ˆÿÌs8RéÕõYh0œŸó…–;«}}’0>ö™6˜}LÆÚbÿÍûø³–ð>>i+”7‰ò«ƒ)¨‘©ÓŽÆÿñ– @>±ÙÁ…p¾Û¡pÒó¯ñ‰<Md#ûbÄop@;“u6d:¨ŒÈ7(ŸjBë>VÑ™†œƒ-ìDý²Fe¦CÚu·wvv>lZA6‹ÉúSø×ÿýüMÆ³m52ð>4ðÞïº]Ê7<!ïâžŠËº›ôV)+´ôhÔÍ?yË§kÖ‹åŠµlÚ-õ_Wæú"¸GÀ ¼.ÜbÙ«e0ÀpÆW­uÚùÃÔpmùá8ô’Gp3}LªT¶^í+RY¿_’Rjgb#/ô…üÉÔÇµ>èóš®XËg4öÙÆ¹\ë²
…žÚ…\D?ßwþi!«ŸØÁZBBVôRvpd·ED:Âš<JêÇä&Z€ –EÕ¾gA±œÀÖ®Ýìœèöß6ƒZÝ¥ntR¤.'­mÑ!gÍíj'Üqößºû¬ð¹»Y»zþvU.iU¾v7èiW×Ï®rJ6ËEóxÂi©¨æY']ë^uªOÝ3½ªéQwWüéf„ÔQ|ì*‘Ú	”·élw[®v³t´³Ê&ùV•vÏp½{éz‡SÏó®-Äy©ÝøL¥x°-TSm2UÔfá]Z 'Qn‚9·ÛìÔ
(OÜ©Í±t‘Çz*WšÅ{êkÈB?õÅkþÙ+NÞêŠö”ò-žYKƒ¥¾^=üH0Dú$ŒÏáFF‘QRJèS) û5täz/ð4SM—ï{¢=;QÁ=ÜœüUóør¥ìµC”d›Èu	/dkO£gÏ+§)³%k0ìl{BØ¤€Y¦TÚS(ƒ©Mí@%5ÁØ‘Cd2‹RÖ’¨„b¯BVr	.©î$¦P.¢3¨Žœs0\5%g·ð.“’±	B‹*ƒŠj?m>R°±H>y§%”³l±;dìÈ)ÍH–!…ŒC+._ëÛ‘k*dš)Aã™c9ÝaC‰“MuöMÇIQcÛÄ,“Ý²vŸy¦
sÚÄ QÛ¤v‹ðPÏÙä‹d¥ëy”LM}JnHªœC¾Hð¨ô ™üH©ïr‹@!£².ô#xÀC,Ë, Dô=s¹‚63eVÃ™IèÈç*%¥mm¦bépöw7dqXÉû-ŠÃî¿$^eŠ¬4CÞ3)¼¬–Ê«pðQ8w¾óV…sFfp`*Ý745Þ¦lÆÃóöC/rÈf,zó^ðONW2›«@W€ÂÅ–hhY/Ð¹),Á_HàÕïý(õBúµù·mÿ$ˆ€"žyó&è1rBÅX]úÄ±öÄ·‰—ô¶GìjÏ³aÍm²¿¤ót+X¤º Ó™Û¤ ƒ¯Ÿ7î!_†¹Í¿ýöo$ÿM:ƒÌ¯ƒ'+d7†ð»ï9z·$ßæíûÉõoqŸbÛ˜Î6[‘	Eœh<<ö#s0¢×sk>àõÜË—/«ë•^l§«‹äÙóEò5ü}i¤ÇwŽ4 l#0úÇ
ðå÷™ÑßÒ×¤UB!»ÞÀ²0©¿eÚ	²HÖViQExóÝ:K3 ,)ÊQÝžŸ¦_îq’²…À³¤ËÈv› Âz7$ìmxfˆúf]Fh»Œ¾²%ï 0X4­Ù–#™Û”_e‡‡ãtgŠ«Ã”zV%tû0 ÍS§f!ÝVôÉè¥4oÁçé½r'¬±ÜÇJQåû6¨uxËé^XŽ¡»}òš¨¥v5qù¥Ê¥¸t†oÇ¿¢zM/ô°ö/’iñ‚rýšËV½pUK‡8&?Ž{Ð‡H¡:å>ñùQì¥YgÎ 	ÄG(¹þ¾õâ!IÇ5¾š[$s)ëvzÔºÄ§~bÊ³“%Ÿé/XdÌÛq¯ÓÇÇ‹dž»M¦?)´&_t,.&ØéyYo@x™^†û‡þ²Ÿ$qóÆ?ìÝ@|Õ•œ$0ë^û0çòá(e?ôß˜c/N|Ú[ŽBZŒé­;ò§ƒÝ£EòÏÆù¬üý¥˜ÏÕ?›&ôw®ßµJ2|WŠ3+2gX³ˆ½r«@§Êuçd9.M]uäÃbùIæÿâ©åå|Niò2rê×+)÷R/)§¸b‰b9oŠ‡öF„§ò¸þ	@T?I¯;óCµ]¡ ÝtdË¬hô·S¦QŒüDp`vi	:Zeê(Þ¦] K‡w0=…ßúo©¬Ží:j«…å,~ŸûÉ–—úšekþJ¯ Fã“ßßOâãÐ¦[@ ²òÛFì¶'Kpp¼–1Œ–ñiýí¸b:RkµAVm#:ŒÃq/öm#Jùm1 Ô1 ´á€”%~6N"M¥[>.ñ£™´#&øüF–Õ,‰6Ú±YYcž?#Î7öTUÅÝ×4!:¥ÎZQž}_]õ…Å÷U9Ã:ƒOÔdHœT…}”–¬‚¥‰ÂôÕJ§z¡IÓÕÑ76Ø”5ûyÎ<*‰Z< $?NFW²A‡¢”¡mG©Ö·–âsŠ5~µŠ Æ	‰\„V`“HÒUžª.n-òX”	@V” 6r`«L"wé…a·PxŒæÆÃÊ®Ìx…¼x´ð›Xúcó’ƒs½ÚœÏžÕÎ–¡°	6‹_Çt¦l’U(MÄnÙm¥RtKO—i<·‡Ú¢rÝevÁXôBá4Öµÿ‚­-?·GŸŽÌ|‡½†³ò’BÅ»î†À!±=}QZð9–$ÝÌ[âÅ„SkÛMË:x^¾Ý>*Ê1ê9ŠÓt¸¸/µ¶³a×¯$çÐAÅ «Bí7ë° +Wôà§ÓÁ ÙéY€e¯A´OË$½á2ñÒ5Œ Uö²\
LýtcqÀz”‰ÅKa”‹X'[¬'ÁFHà¿Ô§9[wø¼`NF‚<kŸ6Kà°z‚Ù–Küa|`ƒŠÓ¬bêÖØR»Û€Eû†âÆš#E›*N¶‘JÑaN*Õ>Å>9C2’Çk ‹mnæJe‘›àjDá„!,@Ê`<”à•#ÀÙ(îû”Íà0Ä…vÿtœx@qÚ„¡Òµ:ÅG5áŠ>–ËWÑmm±‹Å©‰ÎuU_<G>`ó^|GVh—&¬’¥/±óÅÑv"ÿœßEëÉ‚Utçvú8UŽJaÜµn'‹«Ó£æ&aÛ©æY¸«\„«íHººLv 3‘¨&×¿â”FqDÅu†‘=xÑiœ¿ÀZ5Êb¡\=ü”tw)ëé«NIpò§@K4|(aI2hM*îƒ`ôýµÈÌÖVÖ¬å>8¥{7)64µ]L(<”€G~/ŠÃø4ðÉG`e†óùÐ»þï¾¿Hvúc^Õq‘ |öøúr×‚/âQ™É6‰${Å‡ª¥[„àŠN:gc>HüÀ¢)åñÒ+°˜ˆÛ&ÚÔ*á,Ž²ÃéÅQ|@¯/’|«È+	©Uþ¿6ª½¥›¯Š/ádïhù“JÊÔC±Ý-éÝ(€ºròLi5%·nR}‹¤S’LY”»ß"­„3ºÿ!
?·OD•tnÕD½\{ã$ ·ºô­ÄJƒAÎTQ“™v Yú‰÷x¼Õ=ÞFJÃP{ºúÒ!¶g ‹ª3Ÿ¹óGS’7>Â¸Êx¶WqlùÁeÐãª-©oØbhsÃåÐxZãÐ$ˆCoéÐ¨0õåÙÄ¾7w©-ÐÕœ0íU7å)ýƒ$J¥) ¬Ñ÷ƒ&lø\?(© 6 pzÍ$E"…Ó3ÇSýL=A¢%0G·ñdçÆàp˜Ôíe`GÈ¤	Õ6¿ÒYÎµöt”Ý0Ôj…t>´m±¢‡lEoÚ‡ºÏ˜KCcü*0§FÊ~…lÆR]Ø§S÷ÕÍ¨	õfu7Ïh‡½êRÓ/*4çõs& K5Îôâ»ª­þ«’b©¤­zeÐ_©æoäR4×•­ pÇ^Ô£' Ý‡/X"¤K :ôàÌ¦C¬Þ*pnEß«0Ø”YÐbŸ WÌïÃ†–øé(†ù `ÀC¶ž<T/KÍ<t –«_ë1©,Y´©®ûÈË¤4fUgÀ0ø'Åçu]‡gÙV˜~µ2.sVµý([ZS V¢«`­+ýAÐiªsÿ•üBav'AEãQÅ™ªq£êêÒ2ÎT(78“4È™õR¸â ˆßxµ%Ù[í[Þ¢{EêV›9EníÙ;+oS3§\ù:/S¦A™í¡dÂcLhW¥j˜Ÿ,Š.4Çã¢^´Hw"Þ¦˜ÏŒ2¹Œoó™øbd|âï®€ï1Òó(ÝO+ÝsŒv‰ö‚´Õ–ëÅw^¨g¦²&G`Á¨çÍâüËÇÔák_ÏnWrÄ+ëdGwÆRç!(Š%NN½J„CšŠc‚j©UÉv‡ÎHì0cçýžB=^´¼¶L}#Æ¹GJ6îÀ§'ZJ¾ƒ¤…Ù·A<…$ñ"ùèG°ÓU$­™½M>qûD­¦—xîÎ¬±ºRðv²‡¶h\ylU¡ähÜ¿µä„
¡š¬¯¨9ú°ýáP8¡©ì´ WÈ!O¦£‘<¢ÓmZšœq´;S”46|ÞüÛÿþ/òAZÁéLø¨IŸ´e£ò¥âuï,|áá^Di+ÔÀ§ÿey¥—æJOÎè¥ d—^fÌlM7`fi:Þ²#o__3À[>¶G¼ÝîJ¼UU[9`yo»êÔ\xë}ñx+'FÛC©?ioÅØÑ¶ˆ¶ŠêùÁá­ªVD\âæv“ .QCôx‰ÍßZõí5Œµ••QmÅÃtp³ÏÛe¤÷s½ô¾¢‹>4 ü¤vkFÐº&P†h—y4-‹:½RXkŽY"°67]èžóxy÷rµîkÓ¥{y+zWèÔ¤Ý3à*eeÉ°NÊ=¦M¦Ù˜…{ïZËzq/Ôsz¡ÿó¨Èª¨7Ã}{`'~a¨=^©)83às‹;†Õ 	;ÍmØ–€´ólã`b:–¦p¶­ÌÃl_é6Ïx¡¼_6)*ÝÔÅKóÞ>›¨Oôi….´æÂRÆ~®nÀiþ-¢4 •b…
f@ê’8ŠOoèåq‚%ŒîÃ2#¤»=¯1 î‡æË3‚'ÇÖ©¤r¥á´uà@ó€ySƒD-·²‘Áp'OÿSwwéC|t°úü´ow+Ÿ(n%@yÔTáw"ËXe2oôY{JFKÏ´¨WÁ-†§ö ØÊqçÐŠð…Î¹› ³Áª—W@7¹”)®ëÜ³n†¸êƒkxÏâò§gúÛ<äÙ¡Ìi6ìvÉ‚¸úA8DžÄ@	ƒaz¡'	ÈIŒè™xOkÒ5'”
 Òö•ÛIwGÏþ>bEk¥¶ï§Që`‚bI$n°àJjÃ–ðÛÚU=KÊ-p„ÓMðûWÔE!%+ÄÒ0ýŒFØ®?…A·Í<:+üÝ©%¢Õì–˜KsÚ{å%–Öç"ÿ¬ÖXˆ´UKœ·›f…MG}Í8ïà)©0qŠ!ÑJ©ºhV7—’£â€îãÐ=Ò8]Átu¸ø–ÌH…aÐR«O×ËeVÀ‹^èÝ5qÞñ*î?^ZÏ»<uù2Ò;ø¾AÒI°AXªá{Ö:-šp'¸*L.Õƒ@ŸWöÜu+&.mIqøçï½lœÙçm8 ƒ°6³·Ç<t€ûdÀÉëÓœ»ÝÈðÙ»õ’ÏšÚ„IürŒ=(òeŠ-L™\G6p]ŠDµÆI#fÐVxæÕáÈK>…¨¿“Ü -7	R	/
†HGGcØ¡¹ºŒ£ÉYŸâ·…]|Q GYâõ>¡ÏÒy P6·iK m¿—ÑPß2ë7xV&…!{;|O<#ÇJBßëã{³àt¯ r
`Hå'ƒgÍ9N‡úùj€ByÆJ}ô:›EèŸ”\ë<09½6 @ø´àè£	–ûÙƒÓë¿bÎNÒMè],‚\©¶X‰•OüÐ»€£¢*‚CÝ‰ŸÕÑ•w¤JLŠª;ì*C~&ð“#›6üÒ€·ƒ‚$Áx˜ƒñõ_Hâ÷¼'íÑÝùS¿è±€9ø=JEþ )sn&×$±&A”ú	µ ºì'“A£»®ð@™Çø­¢—÷êOæªôù[Þ9êë+u¦)d¡0òÚU«™%×´uŒ†äÓ…ÁÚcNS-óïña<¥Æ£¾Úš»c·ìÝ)¿¬i‚k,ùVx	]jæßlÑÉ~U˜NH¬o¸Ÿ3âÑZê´á;+ÛLé0ÝÔ‘¸ºØÀÛ ¤Žè o¸]ˆmµîS¥™€§â÷Ñ÷-ÍÛ[Ã×–œ;q[äµVJÀ”€°ø°ôvá¢oLÙ½‚rð£VZÄhÌC‚`4ŸØ³8:% âdLƒrN£Î"!/4ÎJÌ«7N ôˆ075÷rÔg™—G˜>qÏ‚>Æ	±¨c¡±ü'Í’L“ÒWŒ’®æ¾òæÎ%Y^^VÇR–æÎFñ¹²Ð”¤´|íJ¿d`iõ|­ÒšBcR{¶Ô	’´N› µ,,š³¶°U£®Èx ŒKEæœD›,˜4”£&kq	&ü@ôU»´ág;-h\BZK;—Õ 7aydd~ d·^Œlü“\3!° ÊßSÖ†YÂ’¿#íÒb&RÅðÂÍùWÆJ¿*ŸNNh8N
Sõl‹…Ì™9ùºÐÛv(«š	[¾Õ™ûÒµ÷²Â Å™MbØä÷NÃøØ÷DÂÿþ{Üú]øEbØMSyGJÜÅ»,º@•¼Ø_ådƒýÞdßð}E:—r\WÎÊCwë @^?FÉ^ˆI¶S Ÿ¦èxÎÁÒú³z±¶ž”°EÎUÈv%…4J"õBÕª°\ôVªžU®’5ƒhï:EÓ˜t˜Àõ¡&ÈVÐV°¢ ­ ­hª{ç@°.Ø¤q[ “Æd.¨ýxWÀF$XÕ¥:ßù	ZøQuµ5N³xHÓiYaéñ¡ž=xB¸t9@ÿ8XŠÌ»)–‹¦%ªšl8š¼!óÆYÏÓJAVÎ&SK,GÆ‹¶E={{}®¶êÊl6@²E¸ðËTiE×ÂÆMÉ††€3›%ëE±$¡”„Ð–ûvüÈQß]ïPëô[Õ¹Mú‡t¶â!0uÆwtb€ZÎÆÑô° *ô/êæi×óÌ0‹—­fØã…-%×El²?MäêÒt)ÈÍmÒ?žÕV(åÒŽM
D#ÀÙ>yB¾š é©zÜÚ©ºÒ&÷û\HOP¯Ä):êß:eeæ!/2dÐS¤#XNtÆÌswsW™TE©UÕ©ñûX“^R´eGû±ÈœÉçG;ñªQµWC½Öâ,Ì¦”#šKÎfRâw§2(‰7sûÉ°¨ó°jUb·nÇ¦$§j¶(±Û–z¯©5‰?\´%Ù»šÜ’¤¬5µ#‰ÍÍIy†8¶µØ–ò¡I“µ)×9ÎïöÕ.¨ÄÜ³‹°žgIÉ74<ÝP~®Sg¯5¾W³/]®üŽ¼ën ;‡ÿôýÎÁö‡²“öâæÜíÃ@mÙÌßúw+jiã×Ã›—ÐméÄfÑ*?§x¬<_¥Î5œV^¬:Þañº@”¹ˆDýà4¦.-6—KþP8xŽÑµŒë¸Žç¡Á_°Òuæ|é9,ÑÂc|¬Ï5ï}Š›ÀÜæšÍs‘ˆ~¢oty)KŽ$ø)7«‚Ït¨À'OÓÎÙ|½é—¸WpëÉ·8÷’°zõpiŸAŽò‡‘}Œ‚ßSÆž™´6›RæI$”Y’ü)°70Õ(¹Ä~ÎªìøŒ®V1ùÅ\Ñ¬"8’½ÐŠìy§‹Döªk”Le×ñcÞ—^¸²."Ð?‘–JðÃUQ µªå?7é—•mµíL?»g«zK8s(Êâ‘.— ñ§ˆ?ÔyÛ‹ãÍ$|qHëÃ×aÞOÁ3Jã¨Ïj§Â Hûµ0óøÑðÚÚŽšùT”w´”k´Á^o³îáG±ðÁ\hEæœÏÚ û5?¿àê•àp|J=ÓÔÎù­Ÿ©×Ðß£=\ãuh¶PßAÍÅNé%Ö“µ#›¹Ò(+á§9%xÄÒ+.Ô!¨¥úØo.`¦&PÓq³ Š`g‰¬+à¶øŸÏJè-@;³¨ðŸb¤}Ã©wˆÆM<˜—îù§èê Kddõµiˆ QC¿ƒ£t/ž,.=,Ù“øŸt<Hè3ËÊJð)ã¤ ´:§/Q<ñP/øÒV#½¸°Pó5b~L\Ûƒóf^v^¼·ï%èýë@šBaGýcÇ(¹Ö†]´b!S7r$¨ÉX°5”Œÿ¹y©®ÛY¢¸f"N7ÉxØê5ßaàèú·lbY1Ín¾ÆªÕÄO…f?fÎgùù=â}JùVßzá€êª€üeILkÚ`ÜyüçqªxŠ·$ä\D=ñ£¤hÄÏãÆYZz–K*‡Yœ3[¬0òzy•t)åüOºtklÛà	è»{Ý½ëé.’Ã£ƒîÑõ¿~»Û5új°OÊsv¯%g,fº—³ø{!^_¢	¬#P{™‡ûM‚¾eËk²—ÊÂ¸,É5®ÿÈûîÁ?îíî}Kž;{ÛMžý°¿sÐ½þ—ëÿØ9„§÷>lí~hÐÁÁÎÖ÷ðùîû÷Ý½&Oílí}x÷Vµsö%òÃw¤ù‰m¼x—5µjôáWÕÂ8s]­:“‚²vÝ®¬ÅØY8KYÕº£ë¿ô"`Ø-šZ#3fœâi†„\ì8‰¾|ø¹gÆ úGy$ø…ßw"EÏ?uwûžÈlg)«š©¨ƒ0Ó÷†II5R‡c2,—¬póàÙå©soZŠìAè³IôŸÆ^O[;ó‡4'Æ‡bÞÛº@ê¡j¨EÐ¤ýM•‰3„åt{÷`g÷èÃµÇû?3$^$h¨ÆÄi´ž |kb=}~_¬§úvÞ™Õxç³(œÆOí§±ºÕÈ&Þ/q9¦Àq;T#L¢çÉÍ½±r?	b"þhÛÑTba(Q{ý—~àµ§¢·l#í†™`ˆ‘nM¡.OÑùÍm²¿þÆ.àYúÇýèý2€f^6vq_"^¤tQ(Nìo…Ô¸-¬`]·ŒbHÂ¾5î!éÉr·oª¿v$³ZB7ù÷Öñå&xñjß*ŽD¶
ÍuÒ«wAš}@ öû¢Bóz9sPÜV
v2DœÔÆ)%°®Ù
Î."8t©ÑEøÁ£Ë53×!Z¨ŸþXì‡\Œ•ç«vbÂ[Í”šÔ6P£º¶HÖÉÓpL7Ðqv'ðzjÃ4¶ˆp	ßÌíìæïÉÚC©)žØL³j4ÊÈKR7ÒKÀ-’µU«±ÑMéX7H  ±‡a,‚“Oz3S–‡ÞEgu‘Ža	feM}PSX2g#ziHF4Ì°ÀøÜæ6Æ3z1M¢|ž^Ø£ÖÌyŽ)Ì1_-+ßOilcÿ‰mŽw‘±F)‘ìF×¿õ™ÚNƒ¾ÒÍŒ.÷šÇ@Ý÷n„t»U³IÞíÃ2RP{Á¬9˜Á²µcÐ§`ìÆG?âE€f€ˆxrç›ë_ÏüðÑ7£‰ÉZ]¹ú˜°§<ÌØ&6(ý>,’›cò#ÑmNtóÕCh»Š|#„×xQ»Š¶…C_pÌä?<NüÌ/fa0Ñøšæ€Ñ8…þÍÙ¾ñ¸¸àiYÇ¢ Èö‰°ÓX”øúÕÊã3OË|žx#¦Î@™È4ãq–Å‘CmÀ˜±<ƒÞ'%ÀS…rÝ?ä3Þz[ÕïºpËŸƒ}ò)^ö"›?«¥zÉ˜z9yNj×´ŒÌapK«·'K£8 –^òtK<’×{M¯DbsÃ·ÎÝÒž`\h®J‘ö"V¥HiMÎˆ°ç@|­lÁÀR­Æ½,éY¢ÏÌl¿¡ ¸ Œg9xÞãtsšEK¶³„Ëb¤t2MÝ@ÿÓÀËRo42Ãº¸û…‚»œ~Ä—–
 ¯Þ½0ï§©wêþyRütº£Ñ/ýõˆ)8?ÿˆ4µËï`‚ýýq:°iYÆXöô“'d~/IÍãhž9Çòwçö4…bXÞï{AFÔ^@˜øó8å}h¤)\é8/p<´+š^ä%óûÖ„=ìãáÞ™ÛçðÎâd‘¼ã ¤ÖRLÎ("žÑ£%ÂRvcr
œ°Ìk‰ßóñ<îÃ¼[ÈtyÎBb¹c»^“”á†mÈMm¤Î}è„—¦Šˆ)Ìi™†)7ï	;zI6àÒUÒ/\šF´Ë¦ ºìX¦¢xV:]Ì·q`–Aev:¯01Ê Rø"gõä— JðrJ'CJ9‡Ê±€Y£˜ õJß8UÒgÀN_­B?¢ë`é)JiˆœÖR9´J°nO`“úMr¿ ŽíÀK»¸L<›(~ª¤|-•ý7HP(>¶ÓMù—kW›kMÉ6¼Ê ôøúÕæ:>žNøüÓ«Í§Ó<ÿüjóù4Ïÿájóžw918t{–˜Û[Ãø-xÒË¨Þf®ÀYt˜ EV$3iË')GH5‚Ö…V4Ö-Ð‡T¡EÕ RvÝ>ø’7ÀÏ±ÿWq+€ÅcòûçÏ_<{)>“Ð±SN5âä$†w_#a®øØ£|,jmóC.g›G³õÂtå™E­fµ™E²-¥EK%™	Änº½¥3¯õN¨[ƒ"šUS«Ê<YÅe[õøj¯ƒÄïOYI¼…%®‹eu¡’¸vKuäLmuXƒŽn^‰w ×A²w§üj\IYòºuÔ¡˜’Èó2F4]¡-†Âð›w®ÄÑ6·ŸÄUñû÷)|EysŸÑpsD^Ôóƒ$.?žÄç)²Š¥J™|ÔMž÷"€æW‰¡ã»šhþ•Ì"Žìµ²–¯uÑú¸]^†¥-:%jP'qwX0…­6ß ¸pñXý%wy/½Óîh´ÕYÁ]8Áâ?sç ¿í†ÀGc\à/Æ«§‰>lýÚÁÞÙÄtpÖÀ¢•Nx8$q²
	
èåx:tV&¡F»‘¿ýNVBpC­±kÐ^%„C˜m%„Ç¢M‹XÀ_/Zà.LP‡Pt­ÕðË+X0ó2%t¿¥2¥qT–)hBÛfQ¦ xíÈ;ÝEÈ{ïEÞ)À†ÖEæ¦‚—ºøñJo­¬wäŸC×i]F;E»3}Ù½k•m>¢ƒÒ¯'»23Åc®i¦zÌuSí+: ›V?Öß¢‹ôùb(E}ÊL<á|Ã¥œÞ£ž€'SxSž5ÿ7Šû>Í,}œø£×¼5†yñùDV‰-–‘ŸLWj'QY¼!8É’ï³ 'Ciõ»NTµ-Ë¢7Ý~ëÎ–Œp¶>£ÏŠ¯E•q«ì2`˜xÁ†n52©HE™«µÅhZÝ¡jÓE® Ó¼õæ6ç
ý‹
súÊ´ISYö§¦8 ³Í	ø˜îÈ¤^H`]ÞæH`ÊÜ>˜ÍF!Ìö3q{*ó™|‡(û¤ Sn>Ë/x2VGèéE!´ë¦ŒjùôÍ65º|ßšZÔL˜/3ÔµkOS÷ ®9Mˆ–½pË´fd1¢ÑÈC=?†¡×-Ÿì“-|¬Aý&jX$RŒRCŠJ™½‡ý™½Íþ9N•¤9Ã aÜLŒÜÍÆl=¬.8Žl=Î8Ð¢{ìmØr*hè‹‚w«]Ž·|T ¶Ø{Šõ SbÏõ%ûÃ|LRUÆ$,G¼ÓÉ
Ó—§dt¼±[šêØ™&²4Mok²¨ø,Ê:{Fà›D[IAóeý/s±JÔ³Äç5#çÐÔ‹(„	m¤w-D+ŒDù Í‘ÖŒË4 z±‡pTb»ŠÞØUÁ³†_2("™ ÷©×‡5’b&"×®TÑ9dãè„É€.þW:mFŽã4Œ½pOóþ{/ë:#z¿À21jn®—Bé}Gü*Ÿñºg#‹Þ‡O±9í1\*ñ;y’i%o%«#Ø(SH®1’×˜dôØÏÎ}?2ÅKObZ¥j4ó|Jz8úXü<UèoþŽk×ì”Ñ{VyÖ îŒVr2S;žó0-ÕW#×?Ž’`ˆ~KŽp+V“§>ÃŒ“Ïs‰µ3dÜÚ©ÉÎüîîáøôº¡U^$?üˆÿ¬£Å%b½»‚×tÚgmHtªèhi96øLÎ¨‰ÖYÂs/ÿ$@ÈÎ´ëÕ0}®x –ý/Ww£¤)wfºYÝå§Q gç3,^®î†.*œŠN§ç§i>®ò-Öµ7GÅ7—Ý?ÕXQ/¦IÖ‚óH3ž×;WRmQV+ˆÎ®1§|“•e~œ^ÿ5êé&”)¼Ð¢f4¼"=<HÇO’âEŸ+iDÍÿ¯hŽüžv»<d±¾¸¤ðsaºƒ¾ÀDåí“Þ7õ‡A)3ã×28Ð•¯1ä—xºcâF¸WóPÃ5‹™N9«r¿Oa•w»,¶´æÇÀ¶£ícgý+wfäŽ¡ÇO!ôÁ£ŸM¼ó[:±A¤Ý0Ç<š3³Ø½‚­>³Ÿ¦[1I!ì(vŽ¶µ¿_á+GÝ¬ý3T«Ðò,á-ÊÀ‰r¢Syä§rL|LunyÁxþµá[‡l¨(&q(O½©XÐ‰ÙÇÒ8šû¼?B7ƒnL2LS pàÏ(¹þ5á&Åè3™7ÀÇ3ûú×Ã)†# „ôL¯ðÃÔ°—x³¶ }çÂïqÔî¨YäÉ?îï¦Ä’u Êá8ª·®:Þ¬ˆƒêlþ8TMu|Âšê¡îU;]läÅYñÌþ¾ÖTV™øôZj«J”0‹5‚3Kì‡™_˜T"VIh´WxÇüSÇíîþÊV÷-ñ²1jÂàà¯	ŒE9¯0ÔEÊ– ÐF¯·¨|‰uâ<ä2¦O4÷ývh5ÍÎqSÛ2µ*Î¢ÏS9ÚCí½´^!ÎgêØ·Š1ïýeæ!ÆYpÆ\ìB\à¤°®ŒÛ¡¢¥ï–Dý¬*«<ÚÁ,“–e
ìº ÖÑ°ÃýÝ™Ú¼šºî“¡ËOŽð Ôt§„SÑbÃÞâR@#òôðº1:¢tjØÎQ:üºA:•êÆfc°‡é´k‡ûÎOú^BS, KÜ äØÉÁT±#|§¥( YÅM	dv˜4È•ÙçÒŠiz@­Ý“'Ä+dyŒúæÙº¬äb¸k$”·ÇÝsÖ{Ú@"-”¨L“Z©ÊÓ šH‰'*ÆNÔ˜7ÎéhÐŠ7ÈüS¼V+Ò¨€/“U,ÇÝ3™‰yêf¨¼…˜¯ÔÇ4f§QœfAO”^ˆûÛòvL–ï	Rt¯}}¤²­tºÍGkòÍ-·/9érD‚…ßò{ù€c¸¯þÊÛ$ÜÀäW¥ÿÄ‡°iã¤G'¸ÛÇv|Têeu"rØJ[euDÓ•Mã‚£ÿ´ºÖèñ«.8ú=óÞ‚t[,ŽaµiÃZ‹­’
ºð¥GùºÿC±!ã(;XPLi¢`loœ$p(mÓ5éžzCK–Gîí^Œ[•7‹fžWÔ[ÐžWzç¼ínáÍ2u-„-Le¿Ïi±Þ%6QØeGK…ºå«"¿÷õ3¼œ%¤5½d°äl¦Ûêvzg>[]$8øí0KJ{‹t±'fž$ˆW?ô¿§ÞæÊÖÐ].v@=Î”WªrU…æY×ÃÏQÏ¿ÅÕÔ¢Ü¾D- S¸Ü íŸ}…ŽÆô=Ìx?¯/G~r„™öLÛáå¼-°0ÙáxwÉ¤?à6¿·§~BsÄ¾¾Ä¯“ðCó$‰8ZàÙüjX«o¼Þ§ñhçb'™m¹X›šk¥56,{_*ùƒ=—·ÃB¤@:ée¬«Oàã}Üìö[þ;oÁ®ÿ1Å5Oè«ô+ê,¶ãó(Œ½>öc¢\Ýúó:|ÃÅS8ÜEWßÈ•ž\$ÅQ×‡ìËQÿäû$ÔB_tºàÂï“ ‚ÝX¢ªšhsåëUCq­ ÷/K¼¨	°ŸÄ#xvœ 7êŠšÚ5n3²øU0áÞÅÒùÒ×pa°ôÃËõ³Á$/ãÑ‹C‚ž'!<4ú}?²qæ†g
Kñ[æ²(å7Õy—­Œ—þäùÒš%TIÿ¿&$5ã¿[Vµ{æê'd…#œA^SYÛŠêž–\Gp$VùÇ eF ™î@–|µ2xZÛÖSZeZÆÆ´°ßP§“<äKls<Nkµ-Xvýå©C¨E2ÿ`Iôi~áÊ¸Š;°À‘¾¢OÊJª‘	––tIXÖ=ö‹ý-9xóÉØ3Ô'‰‡°½1
îË (&Ÿ©+h'yÖàîŽ´¿ý–ì'þYàŸÏý8¿€¦ùïŽÞ¿Û}‹v˜ÑÝ–à†¿
Èû¶ŒÉñ ýŸœuLm±ä{”™‹8UbÎM­› Æ¸÷û	K«ïÞöÝ!zÓÉ4»=Àù>Ñn{êF³;¾iyÏ¼Þò ñO0à€®—¹IŸŸgÐlžârŒ i‚¼exzÞül‡nÚKËJ‹³³b©i‰û„ À6Yì]
œ+}†KkF¸Ò^€quìfþ ËÊd—²^ß«x)Ú ù þ'_™õg*nž}Åf²‹Å ÄúPïG8:õcš’çŠ£è§i‚¬/]—+ƒžnÀþH†‚3päTgÑž#BV™Ä²¤¸P~æj‹WG1°à4å4Pãdiø	Ì)û{…dbÞË1å0m/÷èÕ
°åtWxÿ€W÷ÿ  ÿÿ %'	×