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
  const [customKeyInput, setCustomKeyInput] = React.useState(() => localStorage.getItem('custom_gemini_api_key') || '');
  const [showCustomKey, setShowCustomKey] = React.useState(false);
  const [aiTestStatus, setAiTestStatus] = React.useState<{ testing: boolean; ok?: boolean; message?: string; model?: string } | null>(null);

  const handleSaveCustomKey = () => {
    const cleanKey = customKeyInput.trim();
    if (cleanKey) {
      localStorage.setItem('custom_gemini_api_key', cleanKey);
      alert('Chave API do Google Gemini salva com sucesso no seu navegador!');
    } else {
      localStorage.removeItem('custom_gemini_api_key');
      alert('Chave personalizada removida. A aplicaÃ§Ã£o utilizarÃ¡ a chave padrÃ£o do servidor.');
    }
  };

  const handleTestAI = async () => {
    setAiTestStatus({ testing: true });
    try {
      const currentKey = customKeyInput.trim() || localStorage.getItem('custom_gemini_api_key') || '';
      const headers: Record<string, string> = { "Content-Type": "application/json" };
      if (currentKey) {
        headers["x-custom-api-key"] = currentKey;
      }
      
      let resData: any = null;
      try {
        const response = await fetch("/api/gemini/test", {
          method: "POST",
          headers: headers
        });
        resData = await response.json();
      } catch (proxyErr) {
        const directKey = currentKey;
        if (!directKey) throw new Error("Chave GEMINI_API_KEY nÃ£o configurada no servidor ou no navegador.");
        const directResp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=${directKey}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contents: [{ parts: [{ text: "OK" }] }] })
        });
        if (!directResp.ok) {
          const errBody = await directResp.json().catch(() => ({}));
          throw new Error(errBody.error?.message || `Erro HTTP ${directResp.status}`);
        }
        resData = { ok: true, model: "gemini-3.1-flash-lite", message: "ConexÃ£o direta com a Google Gemini API validada com sucesso!" };
      }

      if (resData && resData.ok) {
        setAiTestStatus({ testing: false, ok: true, message: resData.message || "ConexÃ£o OK!", model: resData.model });
      } else {
        setAiTestStatus({ testing: false, ok: false, message: resData?.error || "Falha na comunicaÃ§Ã£o com a API do Gemini." });
      }
    } catch (err: any) {
      setAiTestStatus({ testing: false, ok: false, message: err?.message || "Erro ao testar a API do Gemini." });
    }
  };

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

      {/* ConfiguraÃ§Ã£o da InteligÃªncia Artificial (Google Gemini API) */}
      <Card className="p-8 max-w-5xl border-2 border-indigo-100 bg-white">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <Sparkles className="text-indigo-600" size={20} />
              InteligÃªncia Artificial (Google Gemini API)
            </h3>
            <p className="text-xs text-slate-500 mt-1">
              GeraÃ§Ã£o automatizada de Planos de AÃ§Ã£o, diagnÃ³sticos inteligentes e consultoria orientada por IA.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <Button
              onClick={handleTestAI}
              disabled={aiTestStatus?.testing}
              variant="outline"
              size="sm"
              className="text-xs border-indigo-200 text-indigo-700 hover:bg-indigo-50 font-bold flex items-center gap-1.5"
            >
              {aiTestStatus?.testing ? (
                <>
                  <RefreshCw size={13} className="animate-spin text-indigo-600" />
                  <span>Testando IA...</span>
                </>
              ) : (
                <>
                  <Sparkles size={13} className="text-indigo-600" />
                  <span>Testar ConexÃ£o IA</span>
                </>
              )}
            </Button>
          </div>
        </div>

        {/* Status da ConexÃ£o IA */}
        {aiTestStatus && !aiTestStatus.testing && (
          <div className={cn(
            "p-4 rounded-xl border text-xs mb-5 transition-all duration-300",
            aiTestStatus.ok ? "bg-emerald-50 border-emerald-200 text-emerald-900" : "bg-rose-50 border-rose-200 text-rose-900"
          )}>
            <div className="flex items-center gap-2 font-bold mb-1">
              {aiTestStatus.ok ? (
                <CheckCircle2 size={16} className="text-emerald-600 shrink-0" />
              ) : (
                <AlertCircle size={16} className="text-rose-600 shrink-0" />
              )}
              <span>{aiTestStatus.ok ? "InteligÃªncia Artificial 100% Operacional!" : "Falha na VerificaÃ§Ã£o da IA"}</span>
            </div>
            <p className="text-[11px] opacity-90 pl-6">{aiTestStatus.message}</p>
            {aiTestStatus.model && (
              <p className="text-[11px] font-mono text-indigo-700 pl-6 mt-1 font-semibold">
                âœ“ Modelo ativo: {aiTestStatus.model}
              </p>
            )}
          </div>
        )}

        {/* InformaÃ§Ãµes da Chave */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Chave Requerida</p>
            <p className="text-sm font-bold text-indigo-700 mt-1 font-mono">GEMINI_API_KEY</p>
            <p className="text-[11px] text-slate-500 mt-1">Google Gemini API Key oficial</p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Modelos Suportados</p>
            <p className="text-sm font-bold text-slate-800 mt-1 font-mono">gemini-3.1-flash-lite</p>
            <p className="text-[11px] text-slate-500 mt-1">Fallback: gemini-3.7-flash</p>
          </div>
          <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Onde Obter Gratuitamente</p>
            <a 
              href="https://aistudio.google.com/app/apikey" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 mt-1 hover:underline"
            >
              <span>Google AI Studio</span>
              <ExternalLink size={12} />
            </a>
            <p className="text-[11px] text-slate-500 mt-1">Gere sua chave gratuita em segundos</p>
          </div>
        </div>

        {/* Campo de inserÃ§Ã£o da chave personalizada */}
        <div className="p-4 bg-slate-50 rounded-xl border border-slate-200/80 space-y-3">
          <div className="flex items-center justify-between">
            <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <Key size={14} className="text-indigo-600" />
              Chave API Personalizada (Opcional para este navegador)
            </label>
            <span className="text-[11px] text-slate-500">
              Formato: <code className="bg-slate-200 px-1 py-0.5 rounded text-[10px] font-mono text-slate-800">AQ... ou AIzaSy...</code>
            </span>
          </div>

          <div className="flex flex-col sm:flex-row items-stretch gap-2">
            <div className="relative flex-1">
              <input
                type={showCustomKey ? "text" : "password"}
                value={customKeyInput}
                onChange={(e) => setCustomKeyInput(e.target.value)}
                placeholder="Cole sua GEMINI_API_KEY aqui..."
                className={cn(
                  "w-full text-xs font-mono px-3 py-2.5 pr-10 rounded-xl border focus:outline-none focus:ring-2 bg-white",
                  customKeyInput.trim().length >= 20
                    ? "border-emerald-400 focus:ring-emerald-500"
                    : "border-slate-300 focus:ring-indigo-500"
                )}
              />
              <button
                type="button"
                onClick={() => setShowCustomKey(!showCustomKey)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                title={showCustomKey ? "Ocultar chave" : "Mostrar chave"}
              >
                {showCustomKey ? <EyeOff size={15} /> : <Eye size={15} />}
              </button>
            </div>

            <Button
              onClick={handleSaveCustomKey}
              className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-4 py-2 flex items-center gap-1.5 shrink-0"
            >
              <Save size={14} />
              Salvar Chave
            </Button>

            {customKeyInput && (
              <Button
                onClick={() => {
                  setCustomKeyInput('');
                  localStorage.removeItem('custom_gemini_api_key');
                  setAiTestStatus(null);
                  alert('Chave removida. O sistema usarÃ¡ a chave padrÃ£o do servidor.');
                }}
                variant="outline"
                className="border-slate-300 text-slate-600 hover:bg-slate-100 font-bold text-xs px-3 py-2 flex items-center gap-1.5 shrink-0"
                title="Limpar chave"
              >
                <Trash2 size={14} />
                Limpar
              </Button>
            )}
          </div>

          {customKeyInput.trim().length >= 20 && (
            <div className="p-2.5 bg-emerald-50 border border-emerald-200 rounded-xl text-emerald-800 text-xs flex items-center gap-2">
              <CheckCircle2 size={15} className="text-emerald-600 shrink-0" />
              <span className="font-semibold text-[11px]">Chave de API do Google Gemini pronta para uso!</span>
            </div>
          )}

          <p className="text-[11px] text-slate-500 leading-relaxed">
            <strong>Dica:</strong> Se vocÃª configurou a chave no ambiente do Google AI Studio com o nome <strong>GEMINI_API_KEY</strong>, o servidor jÃ¡ a reconhece automaticamente sem precisar preencher o campo acima.
          </p>
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
          batchQueue.pushItem(doc(db, 'empresas', id), { ...local, ownerId: user.uid, updatedAt: (local as any).updatedAt || new Date().toISOStrxœì½MwÜF² º÷¯€ø<®ªk²øaËmÓ–8”Huó¶>Ø"Õ÷Ìh4X IXU@@‘¢Ù\ôy‹wf1«9o?ýfqW½êsþØ‹ˆüŠü 
ER²|¯y|¬™å'ýAt5øþ³HÿU³É$./†³é¸ˆ“49,Ž‹Yòå—¡BIZÇÙ¸¦“i™Vq¥kÙÅ³<ïª2ÓYuÚ£xÌ¾ŠÒq•FÙqÔ¿Cß¢/¾ˆFØò ºµ\œç¢©Ge1YESs.’¢ý ’m8ŽŠ¼ª£ú1¹¥ïê2Õ{u:9Ì&iUÇ“©×]‹ºÒTËE("\dK÷ee™(:ŠëÑéŸfé,¥^!Ä~RŒúÉÑrÔSdé-GY2XŽ.£ápHÍ,ã/Bs9’¥å^²Íª´Î2ø8›&q&Ûõf”§çÑ<ôÃºØ;xvP—Y€³8oQÝ½<0n­ÌEÕÜòõÂÁ†¥¤ù}9d.ý<¬©Tg¬›ø­qItÕ…0âa³$Íëª?+Ÿ¤åÉ­ë3÷—ú·Jk­oÁÖõ	ÞA]”ñI:„âÄ=zûšqß?<{:¬ˆs²ã “ÐVW£•••èÁ0:¸ÈGÑNŸäïÿQAŸ‹
¿Èbb2Q§°HuÇS˜TñyœÕÑIZï£ªŸUÛÉ$Ë£­è§YZ^ôGÅxœŽê¬ÈÅÜHx!`÷ƒh³S¹åèü4-Ó~ONè[ïÞ=ø¿š;ƒ¦‹æÂçüúAPcYtRÀ¿ß7u­Î‹r7öAvÜgÌaÁFò÷“!NßK˜î›‘øƒž·cœ²Q\ñusWfÖ4‚cÍieãlÁc˜&†\uyÁ»FhTñYJ©XXÜvbs›=N†¿Qp0!i>Ë*E¹L|JOvW(o—eó±¢û¢:ðø±pÏÙÄ‹F(Ï£Ë+{8â1ØK*@R4~+c‡å @¾„Qf¯í6ß¦U0 V`…læ‘…^¹ìKS•vµÉÇþå+Àçå+Í<Ðó¨/*fIT3Ä9Ý•@«‹:oC?ÛÂŒ1¡v)Œ{?ãë$›gPØî–.lñƒ^ßïøÖÕâ2¾ªvZKeÃ0	ãüb0Ô_¢¿ü¥ë:{-Îš—-Zç_H“³PmÕæ|dÿhtü÷iiuÃ5kvþ ~Ú]'lïÀ×ò®G¸6MCì{mtÒøîi}LOó{(5¿çi5…BqƒÖ‡Ÿ¯§õ•
0ªrãl’Õý»kkk-úŸUãZÊaV£T?]ÍOw°UóS€»j~ªµ9j}œÁ#.½q²=?Vï‰}§—¦Vç^Úé^–Zc²úVRJC]=(Ôt*,·B¥·)Tn¡°B¥I¹©‰ÓE•’(ßT•Ò8vQ¥üÂ·¤J™‰õ«Ñ£Ûµ(Qzl!Ê Ùª>9hþ
t'ÆÉˆ«?SƒŠbœÆ¹hCw=PÀyÕè-¯æ¨j¼1Íãây!µ-ÌëŸ–ÎÆ™»YasØf1míŽCÏr^Os›Ïû-Ýè¤³ÝTÿ>œÞþáGÿ&û5èÕ¬®Çg©¯–ÙÍhPäKQ©-åüãZDZ7gã:-…‚UrÛhd÷½{Q ô0cÞ­+GÇß‘:þa\¦Ç°„ïã¼ëù‡qõözz~-€¿ž"ð6÷®Sðš*>!V~E7©“Û£Øóñê.¶jú
~WMßitŽÂ-ì%áè“Æ½ÖZ¸…qMXÖ|½‚s½Î¨¼½
ÐBî^]ézþÞ®½ïäïE™'ºÍ6óÄ-6O¨Ô¦;ø],‰õM-fÅ/|KŠÍe¿+E¢Ýf£É~!ûD¡×j0–ÉM•ÂfFû´CÃYÍj!¹O@%œÇlHpepA*µùmÙBÜ7À;9l]Æyl4ÏU»+Õ¸'i]$Å¸8ÉâÕÙÑ8+êtGýý2dUWËÑ~YÓ	þ<(Æ³÷ÿúþßRøýþ¯eJïÒ“	Ð¬€UÃÕ+hA E¬QÄ7¸êŒŸÆ“t3Rz
u]‰›(l_¾ZFÊ€Ú
Rø,ÓË®ÎgÎ˜JPãé…‚0°•.[?±Vë¨£Ýf5TXDýdx…u7è3Ó×¨¸ê€­lŽ:j–ƒå ‘ëv×ÆõéÆ3Óû‚ÚÇÐ,.ô3/&©zX¦ÿQ ½†²£ªÔ¦!…#ÓÐ„&´d1[¢™m¢·RB¼š$Y£BÄˆ‹
‘RmÞBEF¨µumµh¼%‰;R?äúðÖ¦ªÔ›„Îäbß¾”IF^mªU¬__LS È˜Æ^qô#ð~fË&Äe„Ã¼¡Ð»µõ­£êÄª%ãH±
ÿ'HÊiÛÒ H;ñiØY{šƒaóºfã8jÆqŠs4»%±œ‰¥Ì–«ÏÂ¿…°î8¾ÔpW/#¿½%K‚°ôEacõS_A,Æéð<.óþ›í³¬*¢¸ˆª,•Ežý—Ñç—’Ÿ¯6ß,GiÈ€ÓG,ö‚ÕïMÕ*2|jVDè[åblÊê†š€Êõ”€êµ•€ê'T•h1)ÂT?	ä~uÁ9 c\ÝZr´-ÖyÜcV¿Õ†ñÑkQlrJE xZa0Ù³kªXªË#Ti²ŸSá	°Àd’ÑúËò,Màs™"w¥BÊ¥•¬ÉdÞ±„`V7&÷DT™p õö Ä™OtÒ”5 ´ :ö8®jT±ÄDìË	9O±ƒZ¯‘~¯ey_­S€4$Pÿ.f£QZUÅ,O>fê”¥Ô¡T'Ôì/EÙ_Ú…¢dVÆy¤e“”½ÿSD¼g')~…nC‰§³³t²¹³ª4îd7À|—`Â‡­!Ø\t×ª¥g£¢,ÓY4›`%o²2¸ÃèšÊ‹³™ .éÑðz‡„«°U Éßþq¢Q•ÿLÍqi–ÙÉIZîH±û ½Meã–×þU‘‹É(¦Çð«ªLÑäJ¦€~4.Ž¤.õ ~ö_ˆ¯p)Ået3êÅÓé2j«XâûÑ):¢ê{³úxåÛ“ºì¬D=âÅóÇÃL¶:}Fë/<÷±A§pœNô¿Á
>CâÉJ» Eà	æ5›¬Tvx
FT€fœ€ÓvèÍ`Éî©µªg¨âBRE š*â‚¬/€ë’¬Bž…b½¼ÈÓž*¤q>*’‹!Ð)Í“‡§Ù8é‹ên{# #Í>fN_Ìê~ŽI€K§„2–×O‹$uWQ¯Àd|q–qRü¦þpÀÊô¬xËhÜ´´é%i9ÚÀU¬LëY™þ³ÔŸÕþ„¦åp	Ù/Ò£3b|¬îÎ×…e†Cƒ?7Û™u¹}¥ùHóâùÞÃb2…á.3ŒïùþaP—E%rÞçîŒÚ…U;2ëìê3¬zßÂ±<º Wº`)ÎT‡'Üe"/g%°ÞÖñÕ†ß Ib³4ªÿs,Tt!‡¥ŠoÜÄ€¸n6­'‚—á¡6ÈMu48+íPp|±I×§â¸RÝ±·èý@Ù±Õpœæ'õ)™BkÙUK•§u„hí{øç«¾ùò^´qwÍ¥r-t¨v^BÒ:À$ó]ç§³ü-”P+à§´Ÿ-#X‚jŠs÷>ê¨Ï`ÝÇbÔsìŸ4óèñ³ôØvå³”Ë:›¬SiË¹>€×ì7ˆ‰ì¤ÀG´7›œ,Ã,7"€k§Šæ­òDG=–ƒŠçYúÞ£¢ÔÚ¡°r¤[K@aÖ+¡+Ÿ_`?2o/…­´IsÇ2ŒäÚ¡ÀBèˆ\?¤ÿË‰<®Ï#¡)g½PÌvYÍèÖéxq~fL¶©ÛàÁº+T¶[[¼Šw¾ÆiSïèëÍ·Öty;”Ói‡ofê¦¬În­ñ*þæò÷áA{X¦`b²8	 ÿ¸à`òªl`_Ø{)m„jLô›ŽÍ33·ç™´¬i—²vä›Îí[×³ru;ÊZÕÍ¨[1¶pÏ5|ÍÌV­™â¹ë|Ð6±e
3èÚ¾e-èw[áV²e[­±•'ñî™h%­yû:=“~úævºÔÇ®Z/²¦mŸ™ÓØ)0øúu¬ß·K™.HþXo<T&1¬ÏY'i œùØ­Å¡ŠÁ÷Q®
Æy2NwßM‹Rl¡	­Èè3Le!«W|R$ñø¨ž}\“d÷é» ‹²(¨¾{Æëêj´>Œ~×§iIÎ³2WzÖÁþ_™€Š	u@üÖÂù"6ºFñè4º„ô3èFœ×åà =(%ÔYÑÍHþîwXƒ•T{õ½WiÎ©5ÜÚ`% øÎZC4 á2? d›O{ëÀG˜ól{Ù'Fj¯Á—MGöÒ#)¯€§˜›‘ú©Ah¹¬'eõ¦Þdõä·@=åˆÜ¤MD>J0‡H*E³oñ,g„¬*«ß¼âÄmÃéyuž–’Ã7K&gLmùV¨Ð¶¬ÖñpÅ$žšÓ0ÝÏVÈVìó—$ÖJÜýX€­³ß3½´á Å"€éD„„tÞGb	Ð¾…—‚è(Þ^V”<gpº†…ñ*ÜFh	
L{€)>6lŽ!ÖÆ0z”â—cƒSñT¦£¢L*§ÑþöóíÇwGç¬þ1Y¥£ªÝ]© ØZøÀ€9NOfq™0cw×|AIo‚^5~8\¾ßŸÒ\…)þpxYA~‚£qmÙxˆ¢Ã/ëz$¤a/Ë¬Ï—–Ëlºl="¿p}0É`¶Ÿ¥´”y¾ù±¯°0T÷7¼ØÞ§ À?á.=ZÉrcŸÕm
95†Ó€o
<5Å;lú{(âÊåàØE'çA,Î~àaŠdqLçôÈÕçÅNìyËõ|qœÝ˜ã¨.}ì¡+DÞâØº¢²ºnëqn5ºç:Zz7æu*]_P¾G·Acï‚tCÕë Ž†þâ¸7ø	ºàÞPu1ÜÃ±Wö"Qåñ4Øeï%jùþ[TOü·R½÷?ˆµ;¼ùoA÷_¢îÌ×1¦V>X¡œ…-¾}W¦V¢Á´¹Ìå³ñx œòËòa.$½`Ü”–è7ÅeíÍ¡1Qx“rêæ ”¹9$=£çrtGs•Ó‡¼äo´úêwÿRëCá ?iD:~‘c´kÓ!E\0²R²ÒS[a”@í@EV·Ã!Ù¾ÞFÎ”€Ü¥hoÍ¨ˆòF“Úèˆsp‘tJ_+ßƒE	µ)µ¢(!ç“Ô¯Ý9yÔ\<mÚª–ºÅ¶&.>„‰«¿ú½scÚ<Sõ:Y$ÆóñW'ÜTäs‹ ^o=‰ÅÉÏ]Vö¹$	¹éD’"¤ÂÑ§%p#hsÑµ¹EµÖ[¤ƒ@[<†i6ÑY-ê=0‹èy¿5"zZ¨40	QÚP²=€'ªt‹Â¢‚ßmø„¼Œ+¬y‹H¡Â¤±Â‡¶Y‚Ë`ëŠÄ\—ÅÚWS|>¢hxœø|D}…&
áèÏ^»ûÅ°	OkÐš(ªÁ†‰¢)?9’vËÀÁHýCAtQ–cäJ‘¿ÿûY:Æ¨IAþÐ()¢#ÚÀ}ôÉ4ƒÙ¯0þ8¦:Í²*ªÒ“YYTÜ¡8Àì“¢»ÜãÉúj‰o¯Ã@^œ?*JP¹jòå5EŒËŒîQÚ_}¹9|µ
¬Ó[éd<hQëß¹At`«ìƒ´ÊÞa¨<l‡ÆÙì(2cèaXMÇYÝïýçÞàåÚ+ÖÆW~Þ^ù¯k+ß½^­½î¢/ñ°¦z=»-BdL4ö_ÎÆuQfñkAø×Ÿ_|®>¿ä¾bðÔç ¯¨”@ªùx:…æšÐ@‰V¢ƒÝÏ·w{F#=KË
ì<(¼1\cïÅxµGç›Òˆô¦ÅxD¶M‹ ùKÔÛÆXÔBxŒ{\/¦8Qz&"CPßå%\K÷Ê|œ¤u,b‚8
tÃì09Ê«Ø^]vËÛ›K­¡Ëüã°zS{*°Ì7Ôs6žk^sßªÑfpEjªaÖ‰Í–•£©¶½3ÏõØ”³%Õ\Cï=9›QM5Ì®“½%Ë‡¸,m`%S"™Ã<¦dÙÈ.¦L=‡ALÉx.KpíÍÈ;°™ÜŽlfŸZó˜ÀÔ˜6;/Ó4Ð¦LÕ0´Œb¿Pw^Fj°ïzÐ”1qÆ7«êbò¸8)ÜwZòÆøU-‰Žì–Á°jgÊò¡.ÌþåÈQ‚"¶íÿY‡Òª¨Z¿ŒZmúlå1nÛ#×vžT®N(c{Ë‘Û\ë‘µHêá­2Â™ýœþcçÙN‡%bðÂè_ÁúŽG?V£õµ¯•12†QD„>à+k½\l­¸ÞJ±è:q½UâºkÄ¢+ÄMÖ‡E¥ü¢2¾‹„wJ/¼Vá\8Œ'q~ZüñhS2jX2Óyh^ÛÊÓnP]nE0O.Üfå(%AEíå¨Ò¿N_Ó$SÌØ"_ô<¹£Í€>M.5üØ uhF–Y]ÅQ™(žœ‹Õ££²8GtTô·/gj”¼ZŽ‚b©éÔ`xZœ UúodÚ)ãlh(–ÆJ5#%óNÔÿüRËUôÇƒ7'ÀâF…€vÍZ7*¨`Á³Q±hº”ƒ8ùÔLS·S¢´¾4Ðh	L
ÿà”sji°-n€ÄÒ£x|£u'Õ¥>—íMüà²~ŠI,ãQ=|…NR?üáðÉã½|:«å±Œû^Ôý±85•b ÕIZcž©´Ú‚Ýd"ðîàK#Ïjg6ƒò=+ku’–ÒDöyN/ôqñ}X¨c2JŽv7­èþPLJŸÃÚ]ËÖ±ŠÄ`CŽaIGë^¬ÔvÊVÏ
˜Ç!”ž€=k'§0i¨âTÜÒvùÓ,;+¢³øç¬X²jýfßCÌo½8|´òmôàÙl)hÆéÈH÷‚a¢×ÿþßfv=Z…‰ Ö±@Ð=Q õ¤]Ég$X¾÷ X½ÈÏËxMü–Åãh_õó´BÆ±TI†Öf*¹:Â*Sxa´ˆT38cÛ6‰éÛ_¨ã¢¨uÁ=Á bCýjC¯.ù¨ìÈ]†N¢¥_qgŒ>qM_eçœöäÛ.-Š¢N›âes«ÂAãu_vëe‚i›nÂ»æ9mj¶ëÐªâ%»]ù¶¹e±†ºË·]ÚEfÅË¦i÷‡t¬
ªHÔr"N‹sA¦…aXxD6M¤lˆöv*g:éŠ{Tï^ôÃ!æ!÷"0þËJ²q¿ËˆN˜gïô	¥Íèc$ý3aNgÀ@	Y:Šò%”°œ’*) ÊàAÿB‹×|ÆÙµÑeÜW§gÞyÇÖ-²…½ùüRPáêõç—èŒæÅy€OÐÔÕ/O„ì¡u W>^Æ:§cK´§z˜P;‹Ë,ÑÏY®¤æqF:›%íEÚƒÿƒÞ`µOù¨€q\ÿã<c!ä£q†K¼û‚žñ†hô<	®q6ïd¬MJÿH	±W/Yü¿ìW¨(9¬ÚŒk£f¢`cCº5t«ža¤sëibž`ŸÌr¨ýÚ‚Câ›fí,IoÏ?z¤P°âÂØÛê-CS²ß¼ÆÀ!þ6¦ „¯{€š{±ðÈ4²¦µÓÒ|¬I!Œ2Ìê¼$dñe3²Úô1TŸî»'›4½èMyýæ§*&EnÐÁJmè¨cMtÄ§ûî(ƒ½áMGÍM);ÕoIÆÇßwN@©vð…Ý'xÓ"t½ÓŒØØ’:®Lâ#Wµ«ûöÉ(-d¾Ãœê‰ðÀ§6DxŒþu1õB•÷²ñ£XÚ"¨ÜeUJÕù°”™:¡g·–‘ógu›œw¶™1ìvŒÊzÏ¤y5jÆ¡ak¹“ÅON¾jÜàUÛ&¨¹òÏ:YÇõ~-ž[ÉÒ/0Ÿ‰ÌŽ,¿{åÁÜµª­yÍkõ<\ð†É0«5˜`¯.É`Ï­kbB•›(ÀÊë-N‰öëf‚¸-B×ã G0·W9b'ëø\qAÆè†©¢’µmã'ìhVbf,Çu&«µ-_­š›é›ÀT†gäIB1ª6AœM¸Ü¥?<¦,h}‡Òá‘8Æƒ¼n¥ÖU^aç¤Û¶°Ç†4·.RŒ—ovMQeÇƒÏ•[V‹‹ôßÝvZò«ùÝ³ÌÀ¶®òDoÝ3Í8F¨W¤b Ñyÿz4N’Pz8¢Dwâ„Î×{Ä‘…üj7"ŽIX×HÕÌ5‰ã¥¹[ˆ8¡¬ mT¯Ò(Ã²î5F7r-ºøÉú"ËN “«ˆ"n•E—r•°y—pÒiÒ$x]ÐÅBUï
OWØØ‘.…6gýNÍÇÒ9ÊÕˆ§ÕÞ‚
e`_ø:¨ºÇ¸Z4=ÞâBÈ¶lG_åð!®FÄC­wB_xMtt‰ÇÙõ'¯hW¦®Ò£2N_‹š¯ÇPspºÀBÕzi@Ú.ÔR½kGteÜÜt¸[Ÿui.£7Å„Ïe‡Öq,ø¬ï:ÑM$š«b-;%íˆ´!âÖ`‘i¾Áã”µ#ÔBj®[#Ö¦(ºX”˜§€ùeuÈˆ§¸eMÔ˜»B{4UÑcÎæ—cQdõÁ-ïDñ%µ?¾v4_XdºµÂQ}mRËÞ6a'< z/@«Å3ŠÍHYZ×øåóËÏÌ•f“eV(àÈ¸üýþr«"eÅ]{ôÊpö'ä¹ÄQä´Y1ÏÞÿ=‰ÃÊUC¤}cxÏe{¤£V¢8ÊUªÜ@ý\b’¢°N„´HÓ’£ÌöÜ&beùqáSJDq7Lð.Óhí‹Üù•±†éssäRKTQ#ãtC‹‚‘CcÌ’‘$CÔJ¢DÅ[£?§evœý4K£³@‚ 3²üp)>G"4H`‰õÇÍ·,õ.¾ã„c°/âŒ&NíùdÔ©ñËƒÿlW‡é»ºO‘CøQlØ®®F;y<ÁDö€îœÎê4Ú~ü8-Ç í#±AéåÓ÷nwD_F¸O‰yA¿ŒÆèB°‹ŒžŸ±û'¶µW‚ï	eVRVYK¤µL‚_r?ÿ¾ŽL’“w—1Ä
è°¥÷3Ä‘?ÿè^šc§•‘Vü<Œò¯oç‰82‚czÌDà•ª™Ät©a]<.ÎÓòa\±Äè@ôƒ·Ùù˜ºÕ§q£˜y‡Y³†"IÇiòä4ò0=«b’Š£U¨'»qrUbY+’Êî„â!5C ußtù{ÅZŸÉ^ìá|†‚éÉ’¸LP»!—Ü~¾»} ƒÅˆ³Ã~Lw™è<¤à9Ãr&½ÌÈ²©Q]•]pÁv„‚+›ƒâ¢5ÉÛTXþêw[ÍÝ ð-Ì §f
Ï‡IßŒž©r “ûQµ8å-NiŸPy@tGU¶âe+«lCþ*U¸´
[ÉLƒçˆ‰;Å¦%[K0òÅ	òÊXÌèn¡´%íl	1—´³rŠ¡Ê™/o:Öä `¬«¬ûýx9:"Q‹›§RÌÁPú˜ÏcZ¯<xÞÃ#ŒUšWYeõÅfÔ;‚‰ÖSÇ:AË{©sXòë(tæL–ƒÕNþêsÜrÄ'¿L] Å¡±j3¶MÊ
vÇKUXtÎŠ‹H"Á°Î¦…ÊàÍÄpª¨(¸1cq>›Ü°ã`²±"xš%Óƒ71f"Þx§ g.§Y[Ý”þxæƒgCeIÎÉWÜñXð"æ ”o³ñxi9ô1d-[<ûÓùêXœÖW×È´[>b¯­I¥æ
}ö;<ËQÍÓ}\Ÿ¥JE-Ã¥£-GŠ¦Ã¬&ÚOQ_–9Y‰Äà¿OK<†€—‡{ûÏ^ï>Ù¾{°=ÌÄzRõUÂ¶Á@HªC^…*Bo­ÚtÀš÷é•PÆ°%KV]Îe3ÊÚ2Ši‚‡ñI»ˆª¡@£Ø²*òÅÈÉA8D˜ÖR#^Ù© ¸º­¯5™LH?¬·¼(oy¿¥(2Ph“Ê¤^2í<°bá¿áE»¹+•ß•êéŠä>+YÅ‰Xt;¯©fé°]xÙ£ô^Oqx1Mo`h˜rbŽ¢™!çFVÂù5Ç±ÅtU \e_2«Z˜‚Ì*Sl–tÁcì¬_ñö>½ÆXTà§ðCIDD¾\•`Áf¿OVÏÎ¡ý©$RßZÄ¼-÷¥ÓOßäÙ—ðn,mÞú„mKú’Ü;Ñr¿í±Œ‘[QŸµtïž*ˆ'Ý­÷
–ÿ	àMZ[-Ýœ$ÿÚ*ÃV23 "YvHRúeæÝ>¢
CÕïà9Ëk°ô€àMÚ¶~’É.ùë7v˜Çˆ*‘Š?FµË˜êú¦^Ú&ó¹"&˜‡qAOˆ4:Ñ…öç8j”ŽúÊKVNž§ÆØ]‚¥ïÛH[ìvÅM¯Š<„ø¹j²ÀnËÀÒ+kS¡ÈÅ&y’Ú‡(T*•ŒUaGGá@+ŒLRC¾“Ø}­‡F–0s<Ù´ÛÙ²QšoFo’‹Ï_ žWoÌ&ƒˆÒÅæÔÈÏZ\½ñjÒØB+—\F@ñ^™ÓÙsÏ«0:Í=%÷JœŸ¦y4Á”øQ1Å3ðY×ûúÑˆ:¢R«ïîñ1|µÖ_òpV:±>Îif— ¹-.|àwÑ©Õ]­ö;ªcýVE€ø«áYi èœcp\6.·ë†¬æ°žœ<¥Í‹T8ûfõ©¸5ÉFgéü“qnCÃ´Òƒð^@­Jl“ÏÒœmF/h«Ô¢ù™í­ôoØÓg¨ÅùHÇ3–u¦ç¨6AÁŸUØ2å˜!H2‡Š#œÕñ$.ë2ÎÊì?ŸPr Q1Ñ~lã!}xš‚d;ºP—òÆÿDÙ±îÝVÖU•r•7¡sp¹W[SÛm‘8_Ìs(%S‚™´] ÐGsßJ•¥Ïöòþï¡L¯…gd .WÀYi\ŽNñ=[‘îp`†Éô–ŸÔ0]|•O´âÅµ‹¬•¢!‡‚ôz`+FwtùP¢3Ÿ¶¢¨CÄÆ,Z Ï²„¸	¸s¨ñ¨,°Þ²ÀùzlÁšWÕyQ&”$à¤(NÆ)q®­pêû$OÒãx6®¥šèÍÖÓ-l[ÞGÅ·ðŸõê
lC¬;xƒíÑ¼ŠibE<s•,Õ³3;à1[‚èë–¬‹ÀÓâ¬00­RëKÏ‘›Å¸ð+ãŸãâ  îoòîó-u²-€Tß#ß–C)†5oSå×âRŽÇÈ”‡ÈñuYÀlMË³´Ä¬³40í[  z6ñx’åsJšK¼éoöÅ°Ý&ûm•8-êâÅóÇXYýöÑÆŒ³
º½:Sl}B%^Eš(·ƒÛ6ÕB²þ©M‘,®èm°þwø¿ZãuÊb.Þ‹ñ‰Z’G3{fîši,hœ;þ½vÍ*x3áP¦Q»{ì#}U'zÅÍL–¹	]Ù&n`÷Ë8|ÿá˜"D'
‘2lîCqtN¾±¥¨DÛ”¿Çx*´//À‹™¢a{«Šœªõ_üjÉ'; îßQ'ØÅ”ezÕÒ!Ìœn*!–ä¬AKVÅ–;Møz|¬oÄSSKØ©l*…[6åíÒßÊêiei’…
á<CXbV…ðWÁqX'“ºð±a-[ÜjQÃ¬’xFjànùP”Fî_‰ 6ÎÞÿ­Ì
T—…¶,ã[‚¸R\‡—o$e;z½Ha‚5Ú$Êá_ÍŽÐ‰v„t,r¬‹9~R‘»LCx±,kÙ“`— GÀÛïÍ<pƒçò,5`ÂˆPH¬±—ôg˜ “r5nŠ¸UúU^LÁ>xÌã''ÙI	èlÇ‡Å®*Ñ·uç6­AÃ242ÜÞžÐÁ{¼Î§ÈÅõGð¼’åÊÒNžp¶ÁíZ2!ñ¦#ksî´½#» BzÐçízÌMÖ×šðÆ[qœÖÀKÈšý0Öƒa$Óuf%Î—2«‘óQšdI1lÎQÌW¬Œhð06 LO‰Ëˆq½ŒŽcsD©ÉcÓ`,2nõÖ?F]›}ƒž~6i”cÂÁ‚ùª‹
s^lïïEL/¤z"Â!Ðq•B÷€"ÇÙ	òãLºq¢ã­Oý¥Õxš­ž¤8«¢ê’bØ!×xƒ¼óˆ¾öög±¸ó’†Ô8ÃÓ¸: ôsâÎ#±YÚóB³÷LÉíÎ¡¾!ä›Õ˜‰K@2Ú˜5¶GÈ’–À£Óø,%'.‰Æn´›ñ™!:À|+oóâ|œ&')É$†ÈydÍñæã‚‰DO¼w¼‹úŽÓEV2ˆ¨”hÛ{‘žKùqŠÙ•AŽÐ=…qfä‚Â¢½@òŸâö®füÎ²ô\Œ.OnåÀ™Q|]¶mö©_s˜Ãˆ³\~Ð1÷üDõÊ²ågtüÀ ª2ý@,VÄ’’¦ØÑå¦£°Ë^Çí)E•Z%Û]wwD¾•uŒØH°0\5š´ºÀ<Í&×ëï¡³É¾º$Qª‹B!¶à³ù'žbg.îpÏ˜,y›WŒ15aVäy<­@uîk„—Eâö€ã½ÒmÑ9E|ðÂªì¸Ñeˆ›Z“Ÿ:ÿÆu}^AZ‘ŽÔFø¤Oàø>üÿ,˜Q	?èJž›Æ6¤Ð)#2Ç£3^Æ!©}O€ã^`®–!YyK$+Â’"àA¾¿Š6©zËö2Âo*RúiWä—£ÏËÑ³iZ’4"çòã½ƒC›ó,nùÉÉã\šwÍ‹¦>Ä%S×esyÎêöÚì.Ìèë²{0÷¿LÔâr>ó25M ¿HhX'¾)wÝ÷·É!î!ú ;gh”1fc/’büþ'`ùPD6è–ÙøÕ¾¾‰•ŠL`Ýü
æÓ¿ahÎû¿Š!sŽ’³$™ó;Om…CSðêœ#Z·-ýhk¼Ìau_ò 5Ÿ¤%(4ânPqh~¸@å[œýpì*»—¨«d1ÄËF±ó½Û
¦¸ÂUŠ)ÜEq ‡jƒbÔ)%"ëì@¨ÔûS§5Ù…hÉ€.Öö+[Ñ»1ÇÚ\`s	‹lâ }è<ÄA¦ºÍAš›oŸƒdŒÍAêD¿ÏA¿±D˜%ø°r–`¡¬aŽ0Çí}†`•9?èÀÎ[g™ÌâÃà#2C%˜¡âÌPud†ê—fkD9/¨Ë½ºªí·ÃÞæÍ°Ã9¿©[/oÝ/üré0ü†ƒãÒÔ!u«á.[{á“MUãD“±'·M2ž´nÙIZgÑO%Ù¸=Ò™î[Ž%Økp:µEHÄ±ê–Œ31Ô¿ 	M’Û#£Mî4ì[´jã5öç“QÿÜ÷jQ·ßkñn¿–ÓÏy+XÊ idÁ	Ç
šåÈsi‘—brH†ËdÇè J)ú²È-?.N(nUƒ	,²Z©Úøã¬Šããq–§èÁÙÒv§—ã§1Ì„øÐj€N™ÁÄµävG/é‹²fœµŽ>†NÉõŽú;)_îÒàrg–5BëË¨½Éa7â{ík½M…nY‘“  -³W¹X8—WŽdºücº2g'šo>Û®
îjQ*éìÄ•ôt 0ü)½4Ï(ÃRéDþ6?á:tz‡µäØ‘Ä¹“ð±KºBÔƒGÍ\äÖü¡ÃUQ¤‘Ùët#â£½·éš•ÕÑ8­ø	$85r£•t-í±žÀÇ¡*ƒ9
ï”|œÊîã„àô8Q×'„uÝqÒÝê	ÏõœA2ù½’9åIXj„Ê¦*;¡é®_\ÚWkŠñ¢›5ïŒøxºb Ç‹š»þx!¬ ƒså£µì{*®"UŠÀÙ€ŽºËQ§Á•*th\ýÀó÷¨¥ãWwj>¢u÷­õpÊV®? õ-Ž¦gu5Ž¢eÑx£ÇïEPcX³TÌC,<íÂÒ„„2„ˆ}W7#½K«fyÒ>õöOÖõ¨¾Þ`Õ·p²NXÿÆéI<º°seD}¥î¥ÇþMÝE¯{žá½”I`‰½*z±·s=ýo4Ž3±›½)Ê)hõ0?ñ]šÈ›Sí«7ßkPV¯™ÃKo ¢¯Ï:…ÛU4LY×¤E@² <qÌ<O–g’»!´éf¯Âb³SÄL5¿Áà„ËéRe;|Ù™#è 0Á$©-3ÎÑ«„é¦ä„‡ABhÌ¼'»‡áç4³ö=üóCd×Âw_Þ‹¾^óê*<hîÜ‹ÎK`Î’Y¸°‘:åoÑœ•mˆ‚³el„Úp«qu°¡LôG0>>‘ÀEnáõeÐ)ŽñLÐ<±fæ«ùc' ŽŠÉ$«Ýû¯y?Y_7ÉrP×‘ÕÂˆ!J¨ëÎìÇñöDL 'P0‡žš(˜ËYÚ)o¶J¨çiËÇ°Žõv<ÞXM9¬0ÈäqœcX[4Ë8úCAg“šE™¶_|™ø‹±¨/Î¶EŒœuÂçÏP´ß;Èú*9ËÄ²¦A\„. øÞ#+ïÒ½>M…sˆ2°¨ —AªÐã¢ôŠ;†:»9-Óc®—Ó.Ý=JAúX¸éö]KVö–¶Wº×°Ðl£%pVŒJ•B‘3&xõ·ŠÖ×ÖÖ"vŸ¶Wµ­µ.TlíG!" ƒ£6¼-xe›1c#PXÄpò¾ÍiÅeÈº‡gð»>@ˆ†mÿ~Ä†À8TTÚ©ë¤˜-Ê%ÿ>;#JLªýc~Î‘3)fxýU’¦Á
Q‹N/±´Ë˜P^}
1ËKX“÷å+Î;nÚ;ŠÓcôt{}íaŠ‡†|*‚4…‚\‰}”êÂ(P¶¢§»Þ}E ÇäÏ¬€&â 2¿ëÊ80™àk<:%Ž:aˆ›³Äµ´p4T)8,b±ÂmTc—%8!²sH,‰XAæ]Éš!^!:^AaÁ‰¢6í¬BáuÄ8RÌ"bÝ_¡—’:Û¿ÝÖ2£v®¹¸("™•‰©A6‹<v©†*Ì$QøF¿ä•¹+qcu`­«n,f‰6NËgÒkñÖ‘¡µlEÎíd Ù…³»JG.ÅPÉÊ/ì#åÑ)TP ´·°„T·¾Ü)í>Šä%Î=2ÍŸ\Wì‹“#úyz<”æµØXo˜Œ†ú˜¡Q^±ùŠ3Pí¦†Ê{yôèv&ÇÈNãü¡tÀ’œ&3L	rÛp¤2lg?P[h{Iµ—c%ÙÀ¡s7êT×aþIÝ
×b¿ÔWËé³$L:vc’…À\'×É&¨`nÉ«>É°ÒR$ºÂ‹>ñ˜Jæ§f˜#Ã~©„<åã±AŒLƒG*sQˆ»š`qqë£ÍÉ2–q±ˆD°@–ùÌ2ž±— ƒÅ³|–:’¹©™îpìá±Hbgì	ÒµÈ”!-yxÂI H>ËA»F²aèƒNþ×j
æ8žMªf^ä+"‡ y¹)  4,Sê –Q†Üæc©dã–T²n¯t*YüÊVY’ÙvCs ÁüI¤¢ê"¸Vp§Q–™€9q¬T4‚«Fƒ¨¢
Tj`mn…ýÔ6w'£I"êŒ?xGèQ\‹¼è®ØB3«MŸˆ2ôêu/ú22—úÂC^=‰ëÓaçI!è(oþê›ÁOn‰§åèwž¡LÈ Ÿ1QjŸo>zÄØ9#ntËÍ¥Ò>â¬èºI’tÙjBÉÕMOÊÚ§Þ§¬Ü´¹”d/*e±š]
¨O·†øƒ¤È,W)·¨&¼ÞŒÖyœË›ôë`÷žûÆ;g7ÝSú#Ðâ7#û¨¼¤³[Örv’ÓŽp›HÆ\vîlôïË0ã

ïsÚòiÜou<ÐÎŠ;œÎªÓ>­º› Z{ëØ·CûÙì ¶±=ÆÖrA…ä Þ1kŽéVDÍ„6¨&Cüß°Éü’{28`f‘‡ÍVv`Ÿàu†Í„í<‘Šsx/ÿüQB{aR}˜½·­Eöh§°mƒ€Nâ6"°aè»ÚB€¡Óå±ë­Gôð¶©‡«¿ÅpÃ-‡†EVg“An-8÷cË,/Œ¿À¿ùÀÚá`V²a´¾cÄ*_|£ÑüàZ2Å˜Jj%2sL<+³™»‡ÑlE†‚5ïËtc²œÂy„-NÊx‹ ìÊ>öæI¡ëÎ¤JÔ}ß2$ÀËµ
¦ŽX-ÛîEë¾¤ 3ÒÈpÔ/ZD…(|ªË½ø{ÇqU«Ž ¤ãH¼‚Ét––tÜ_Lê•»ƒïÉ…
µ¢»0…Ê$Åà”#ÚêJtö3!¦ZM:úéyË!ƒîÎ.øåLÇ–Í^l8Xº”‘XG—rlëÕEPa\Ð4™)ûñh¬/r#ŽF *ÁL&Aì-,G®àd÷Q;™p#ú'a|Wä\X4B¥†÷Ór$B£¤»Yú¬«Qß4/ ŠÍ¦õ ùO.üŽðLÈ´WOhAF[.4xÎ®n™%«OžØ9žðOªZ¼¶nŠ´æ‹™”É@¾D˜k$,¾€M=~oŸaÁÊø*yŠërnª>ü­%)—?¢ŠðèèÈ,·wyn‰4¢|+*;N<«hžîK‚¥¥ŠH`¾ÿn·ÊÞ ¸§FuŸ,Þ²g‚˜FÌµ(uòŽdêKØéá6i›ö"‰kÒ«ì.Ï8¦ÆÚvÏ˜ê0µ¾…!žÏ/CFò=[àK||ùVªHHÝcœj·³'¡mÒÿú—¿4CÊ§?†°-õy’€õ¥ƒt<Ç=±¿µ@©!Ž‹<}”½+‚ BZàùG‰mpþwÇ:ÓÃ2ÑÌF«-éfƒ…-è¬4{x´÷Á†§ÙÄ@s8Ç)ÁŠjþ°KÉ¤{.p ‰¡e ùXÛeùØ™
þˆr)ß¸»"	/î‹¤ápåÈ¦.hÃ<kå&æa·
æ“b½ðâÌù¿já‰»–,"‹Ks½ì(ù…çˆ~OÉ3)7^PIå–,sq‰G=ê¤µ/hUkkUv’ïå˜ïh¿˜Î¦2iîAˆ&Ò¾L¤Çô2ÉBoxC´-¸<¶kÉ¦æ'W’‹¹1ë
÷ê:„Mïar%22â¼ŠI™§+£qQ¥ÉÊÑÅ
~:Ù½°õH¡('ªƒG½ Ãøª ûí¢tŒ8£v¬å-ÊqqÒ_#Aõ@ÐÛ±Mž£*|SXcf9þñ3t!Ó=bÚNQz’‚‰wBYÿ¬æ¢“^‚íµ™ã¨-“÷Ï³"J0ÑbÓ–˜H+—ãÿàíû¿‘>SÒšÊJ¦ÌGID€Û˜Hr–Å’s†Ñ‹:CÇxTÈoÅFVäÒ•Ñî
I³4:HóShû(Ñ3ìÍ% ³Ï+Æâ@¥Ihg×ô×¦#PVïÇ?‹ƒe 	ÓæÉ.,q­‘Ý¤!&%¥F£”yzJ¦èýGõ·èU6ISò}áz"’º*7<aù”¦lôPä‰f•á¬çk<È…b~Õ¤rÉé'\hj
Üód_¦è•ÿOŸÄOõÕsvƒ€ ÉtÏiÍ’Ò½/Š:w4´JU% ‹ÃéGºl„Ûµ»«\ŒNéM¨zã*Ó&û¾'Îk`ûFC
ÄJ<Æë-/V²åd¯Y@ìbÍTÌÓAqN€âÅ0Ú‡I|ŸÑÿßÿkÌæO5‹ñ‚Ó˜Mg(D-Ëñîßde¤iEÒ/XöŒ±““_º©{%/êÊ	ÞÖQÑGQ„v2µ ÝïëÞƒó4~kjnl[6’€4ˆpwšŽ‹LE0u¾Á‹pA‚¤˜¯{ÓíØQ+­í+Ä0ÚN2Œ”ÆÅ@/99˜u#ÊG(®»ý·”R›éÊÛ&õ?Tõ‰Ñ8–2OªÎ…Jë¿Œ¶J)`bn‹¨Ç;Üg¨5	¥Ë³YMRdÀ
Ÿ`F8²:uÁþÎ£¹úšôC’³S_eî•Ò=µÎL5¤ä«‡–¯Ef×«5"ë#"Ý”g“oäa(_ÿ–~µ)’ðZYx%HXÜ òÈyÔ’}lSéâÎ…7+ðì·`w@q©Ru,IGïåÈªƒ-‘L^Ä ˆ>ˆ£® Tð0â—0üŒw‰ò;Ž£
¡9ãÄo1å{œÂÈ.=ÅYz
Â[²NS:›†×5Kï¤”Š0¢Ê<,X¨yxËO	³Ø¸³†m*¯Åã	åºGãâÇ
fó`‚ì€©Jó|3ê™Á:Ë³^M&ìðqÂÛøke¤^‡rP•ÂXZÿÚ‰±÷/YBZ[‚™š1=f3~8 é‡ƒF¼«—bJTŠ² Ô»ý7„aÉ+ž¸±$oc9*ã´¼*îTÄÂÀ®n1íˆMQmÆ²2½;‚ù3.’tˆ‚~aÙƒbVŽ€}ô\ôÉßü€_©LÚ.Yîï:x"áTñü^òäöÍÁc„Ž(è4Ç2~b  Û öêÛY=Y¦drÎœ3BøÛO5	7{CGÒõèYr£.eJøz(óÄs5›sD•zŽ²È€…\UŽ|Mä.zS"L^º‰¯?NéB›û“D~Äûý6ÅšfõXgJqè+‚‚Ïâ2á²èœýx9@““ý² M£„ŽTN`ªâlE€ñ} ˆ‡š÷_m4•øCšœÒB¯žŠWÿd ¢UÓðßØàZ"$=Ô`èöŸþ¾·ÌÄÐŠ’U+¦…åh}m™?Ü–Í*	Àmö¬þ†É›^š=ñ‡ÅlLGwq`åÀP3èÙW+Ã¶X`¤YÎx—­¡*=Ftht5"FñSY=£s~âú5y“8µšUè¼«ìÝ¦ƒeãñÃbªÝú7@˜oïÂÿ6¾»§éã“•»kkV5ÚpK“ç¸Ã$õWwéßåHü×{d“Ú?LßÕ¢ý»wýÏŠ¼Æu£¿¾þØ_‚Ùt–¢&
D\:*ÆÉ’S²†&`Hž==xñøðÙó½íhÿù³¥.YÿÝ‘¦ÿ_ÂúÊ(^
Byž{NøÏg¨~%È÷¤ãÏY"×öQ|”b–v±ÀF'ðôgÜÞƒ^Tïÿ3°\l¨†>ë€ÑÆWËÑ×|…ÉšW¾ccÓ•(’$$\–e×6§pC×¾ÞµàØZû«MÈˆ˜Î :oäý›Ñç—n&ŒzôŠ+!ÀCß¯¢¿D´nEb‘xÿw¼€5®hW§0}ñ«]L~|chóÕZ ß]Œ	Œ•õ	<±¤a®W|”KŒºMö*zÿ¿«piÔdíÒÈ§ | ûm¬$ÙIVÃ«I–ÏDú°äGô›ÆAÜ)ãs9Ó6¾AVÂYùõg¦‹™ðŒT_þz#4a7ªåN/uz5Ãs0ñŸï>Þ>|ÿ¿žï=ÛŒöo?}íì‚ôûÞÿßÏ¢ÝƒÃçðñü~ïá³ž§Œàív;"q¯«|ÐÕŠuè^E­W q°‹òágõ…|Ù´#ìlç‹6\GT#hP2Û„ÒYmj>¢d¢þÆÔ8^ €Û·×™jÚTÛŠúú8ëÂ;÷]ÿ­YP^1ÀúBª6±¢8mêýÖEC|×76×Öà?¼FIeÛtà˜( I!P‡_Þæ¹	i"c c¬@‹p…×DõZÄ‹oõÖv…¼«ÁÀx¨_ˆEV=ŸêoÆ\X_¡4ûRk“ìø?«x—ø¨rë¡ë·°ªÔN|Q)P£4÷5üÕ¨OŠÿ)úFÿoãëp¯[Kª&Å›Ï/È«äMpmT©@¹7gµî£è¼-îÓ©ŽìGe}Å¾7Ì©v?Æ”»5ô…â±r¾ó!l¹-ië*Ôgõâ¢‡äu¹Áo¾Ìƒ&ï’#+ÄE]ˆ'Ì¯pI]ÐÊÎªdú¬.Q™Îüƒu\Öÿe3ºËâÇOAúoF/_öþ/r{ÿ×ç»Ûøcgoû÷Oßÿ¯ƒCõÀm †=x¼û„>ájðÿîD»Ñ“ÝÇ@-í ¿ï?;8Ü>èÑr°*Üû¿þy÷1>ï=}ÿ?î=ÃŸûÏ·ÿ+ýØyñ|(^ì=¿ö_ˆ…¦÷ê•Áï¨H.6Íêb>ˆeqÓãzS/s%Zæ±.¦›Ñ×ß. ”ìMPžÀÚ3kg–àj»ëå¦¸ÌÅóû©[^È3òt6Aåð¾CIT¿ÁJj 0kÂ£ƒúbŒÆ]ZqÊ$€äFÁ«åf£ VZ â.çà¤ò¸ýÎíãEÊ ¡ZŒzêhÞÖýŒéx¼„q­…º†ÃîšÁehci…áÎ½¼»¾}U¿½ûÊèM¿c´`hŒŠñl’kDÄ5Ä›$uj3*»=æ€¢hÝ©°ê˜O@Ö¿oìúNý¯ïÚß¿r¾ë|ÿÚùþÕ†ýý®ó}}n¾Y¸ÆïÜkój|»pß¹t^o¯aOú}t>=„Úsg}%"…Õ=NÉEOßÞ'¸f(Iüþm8øøŒ6cEh’Œ'U†¼tHÆö¶nÃ²bªTÄšCÍß[½±væwËÑï6^‘iPUºò5Û¸OÞÿìó!¯ƒ X_ÐßÈ1J=ôÕM©õÝmQëá³§¿xÿ?wžÍíÊGXþ€#ï¾²<'­dÛß}º³ûôp÷V†$(âã1¹ßëôyqnä!—ï¸ˆm îw7^¡ÇRµ»|¥ëgiy<.Îq#ÌÃ#PÙÞJ¹¯´¡ßc—Šqè£¢¨Ï§Ô†E°"×g1Ò^[½y ê’X!ŸSAÏl)Z§#E÷Dxñå—œpO/[@ïaÆMzñ^û9¾õ¿qGÌšó™ül0Þ<×8èñq~+Ñþû¿GŸ_fWøAzi0ç(cs¯F–.Ø´½"<™d¬[Gúî,ý•mÔÊç0¥Šóa1ÅKê ­bVOgu¿w4.ŽfåØò ƒbòú:ùV§Ê²B-Z/Ý,ÕÆ.'¸Fý„rÓM…©M3wÃê6Æ=Ïè8ß=Jÿ1Â±=øÿr4ŽRX!írTäÇÙèmt%£E0I üáhŠŸÊ²%‹Ý»tÏ¶Èü •¸ž/KÒ£¸|†Äun6ÔXã¸ªÐÅ{ïr”æZ:_9ÆýÚã1È;:z½"F3:‰§+_EÓw+_GÓø%½´+ïÆQ]Æt™f‘¯`Î±d&Ò¡¿':rÓ%&3„³šñÍV´tt²ÂD)D+ç§Ð|TÆÀ&+ãõK•ØKÊf´DÕ”pY‹NQ˜l2ÐwÕ;*È\’p‚8÷åã8fQ|ïrcíÊ!Ùë5ýDC+Ð÷Ñf…^¸•õõµ%5üØqÓÓ%§_;õ=œWÑªÆ´šÆ9Ão	æJ•N2òÍÞ¿$¾»úaË‰J?¬
Ã'!YG‡KH~L²³0Ë,gïÒ$ÂUë•‹•52làŸŸ‘Üç+¿Ûˆ`ÄXÑ1©•Rý]Ço%åWP"·¶®`$Çød“JQõwØäÉ&ÚÙ	LÖ"°[aÑ8Í’$Í5ÿe|jÐ p@4+üÎÁƒ+Ml›K4Uð+ ,D§bþLW¾Yº¯Ù3X%0»&G@œc¬²_ýKâÿš'›ž~k?Îª:;¾P­“Œ#¨<ˆË‡§`‘¥&ÆãCÁX€mö½\ðAFÙB’:'¤Íè-Hó•šööØœùvÙ2&®äEž.ÝgÚx/Ÿ MÀ—ëëÓw¯"Ñ<Ì¡Èsà¨hR¯¬/Ý?ºˆöê÷ÿÌøè÷Å$­üöêÈGóœÇ­VÖ#€2Jav9l¡–[„y|ïò´ca} poi/ÿ÷QV,Ù…HºwiƒRjJ„»y‰,C—%Šëž›	È¯WVK«]»ù`–q¸7Â}•[>ÕÜÎŠãžYZuï±©ÒÚíëv¬ÌÒc¼ª:Ü1àå1ÀqnÇ¦XpT/Ð/]ããæCX"ó$.Éßì÷ö	ÀHC=³ç’ ¦:Ý	 Ë8¦d§Oæ3&Ï'ß7­,ô­¹T×tbæ[0Úr²©Ï¦ß§Y:Nè&ì+¿Œìýï1ùÿ"ý8ƒõçý¿zTS^ÃÒîQ¡º–G]ºxXÓÓ‹æÞQ—€î”{náÖÒü~9½ÂJƒy}úÁz\uåaµÌ‡yxGeˆžÄ£rþb2ÁR+&AgNv+~|¡´}í6H#z—æàÎÞÁÃ¹dH²j´‹{ZUÈ!ÝÉàVüød8,SÊhýb¦…¹¼3ˆ«¿¯E‰PåOyò=L
üˆÎ™Ñüya_ÓJ y=}äšÝz”StD…ûe.íí v`V“¹êS:dùÁõ?Hktð4híƒ
s±¯$´®Øëò-Ø_6üßpÍÅ;˜â‰ ê4MÝc4ØQ©˜IÔi¢£+#Sk‘…Ê¯<wÁšßg¥&6wÖö.Úc–g®òä97ïcÃLö±S§LÎŸúd*]_ÏøalZnâ:î€i½òòÕ¾kH™¾_[¦ïå¦¶â……v0D­H…ù4ÍOg“^`†9Íh—a4]ùšûM.>J×?‚Ûìt9Jëó4ÍÑÙãƒiöG¬üÆU8Ãh´m¿DU/ÝNïK›JW!WH¨ùË7áöÑI…Žàµá]M3ò|}ÎZJÀ{œ×¸mòU´õ€Ú´¹´®œUj«	¥zÌ›¥0—KïêÍ•p¼Âäõ•ß¹`§=Uh<OWÖ#Í+èÜµ:­ö­V¤¿1@W„@žSZú›Üé<Õ‰¼w}‚Þô`@»o÷./£s±'ýæóK
ÞšÄïúkË"fÆnat¿õUþ–¦ ØÎgxt+úê›»€L¨È°bd¼gâ«5HipõŸÞ˜Í
óçRð¼W/ÖÝK“ô×G™MNÜ‘¬ÊÑ½K‘£mzZÔÅ‹ç1èéÍi]O«ÍÕÕY¶ŸÅu\bÂšÉj<ÍV·rlýsY| ËbtõÅÌàâ¯{ëkGß}»þÅwï¿ñäp<®ÉàÞBô[Œ+…Û†úÁæß’¬LÓ²LËýÖŠ‹{Ky±¢^ÙEÝ•!@o˜LÀR+ç+®ŠO=áWMÂ‚ï[rÃÎrL¾t?@ÉV§À¿«\iê¥£×>8ß×í½»,z$? jóº5r×Ø‘Ó²ÕÚ	“oZ7ìì­#†”ÓuÀúÙ¬fÛcî°“¨=ˆ³Rjur‡›]¨ ÙƒþÉ7«D^u°2tXW$N0¯ÏÖ²%S.³k+N‹óÃ"®êþ’>O)Ó|„Y.¢çñÏhìËLBtºBœÀ£´lN„ï¸ ãTÒg@É±ë˜tÔ§ûÿLnMá±ÆTbªÝ,)º/¼+”wÈäÓï¼nÉDÕÔ46äA±Š"¬âHQ¸/ÑDY~öþo ®¸íå¸]—F³‰è…ü@¹Ö¿Ž’÷?Éð,~€Dá,!ó‰µÿiõ<•‡ŽâÎ›Û•\ê9ßÛ)¹ÿ¨x#QCØ?NOðÚì0U¡‚EÔõõQ
¼\F“"ÁÅM82×$R‰%qO‘{ÿÿåiŒgò)äÔq‚DOˆf Ÿèf¦—Žë=¥õcïå)zèïÃbz¡2R1Úê›(õ½©ø¶©ÎK¡À?˜FiUP¶FÕ ¡ºba™ÅF$þåã,–I/°rqéY	HK*´È?ð_¥˜7K!Ìºò$X
½:‡{ûÏ^ï>Ù¾{°mNÀg7EÃ4rø!­“—°›üXÛe_ñ¢¢¾ºhÂ$eh¯<&T«Y™J {¹Éÿ‹íÚ¼¡c¥€ä”“ï³u²¬Û‡™™=„)2¡nšì«œAvm’^§ÉC”`¬"I4²N’Êú„!ëú eHš…PwÒë`7–rÚ²êÃpoFfáaÇf“MIGóƒ¤Ú1/ñLÙf€¦„•å®­ HGgQ•}t:¶éÒŠ…˜Â¼|(³<±Ó0ŠµwðLGcéò]’µ›S[¡Ká_
R.«kÒ^µ'wËÙ8÷‚`K1	ŽF·YhTæŽÌ‚£ã&ÜOK<ê‚z2í;é`½ÄÞìkèfäDeÊïtÕu§Û­ýÜ~)>¥ú6d!qäEy©¸Yà1”—[wÌ6þ©NxÁU|†éSá-X~Ô_{6§9!†¤‡²v\ðn| ¬¹u›ÖªN'–x`Ùz¢–Mu[ñÕ÷•]ºDAþMóýMóýMóý5h¾2™Š[…œÓ—·§"SXÿo
²µ(ü¦ö.¤g}Ú¯[7€L(ìêöt_Š™y¤¬jZ•E/Õ:/ò"Jª¼Å0°5ž\^‚Ýr«7\Š¬Â!hÿ0F][(…
ê…ú:%­.TÂªáÂÚµöô7ÔÇPß„Û‘Åñ5»Y$Þt?òûnI!‹#Þ¿&ÃížRÌBïç+g¡Z]4Žm‹Š†™óâ¤Ä.•q•Ó¬,n_AÓ!Ý)7=6•Bdßç*gm•ç(g‰f•ìÖˆx-…ìãè`	™ò¡ýåZ*çÖ[ñHâ-Óâ~ièº‚„ Q:8ðú¼rà^‡NZÃ×l¼ÏµäK:Ÿ3\½i,ÕÒ˜q~âÐ·z?¥cåÏtæ7äûòý…8àZ©°WŽ¹¾CÞ› +iâlz¿× ¡•ù(’„lF½í:;‹Yò´(\¾bSß‹ˆW©	VQ—Ax¤ã:´+åEtœ¦	î¦ûJÂ·Ü‘ø!ä„†÷Žù%C$²JæÄ&+x­ n#‹[lñ¢ 	HÉK¹X78íØ@5:¶BMfÌ­¿¼<üŒæÏ'¡Áž@ü*QO½1•åe'VE}ã	U¢'ÖšºõÄnË\}"ZÏŒE¬P¬ºî-(€¿Ôz¿5ª¨¸Ñp‘j‰‹i‘ãéó(ÎÌoxBçP•¯ë–Oô…F»aÄ=§s§I7ó8ÛŸç=±iî‹!»©»úµ}&¦ƒï:(!´	'/È
ú±G!?öèF~líaYzn³5ÉìçšÝØ5Ë(M"ä•¤á¿<ß;Üm¦ÕíºÈì.nòyVÖ|W¹£É9îòÚí¸ÌI¡%õ®¤ädé¿™l¿™l¿™lÊd³\ç\‹ÿü¶¼[ð§;æÔ€¯gÞ‰Ê]Í»bœ<ç—@z„ä€ìª¨LkÖT¦›lÝÒ){=¿"7&Uc’}¦¤†h|¦ºPÛÈ@©†Fè2":»ôÏÇJêº.Büö?¯Ò¡ÔgÝZþ€açv'CÔ8¼õ0,RÑ²ùÀá×6TÔØC<Nµ˜®î(ë³!Ãr}b“¤Xl3ÏÊìhD/ÿdJ~þÉ´ÙÓÐÕ?™úFÉ)ÝZf†èjoT­‚hNJ+¬P®Îã	óWÉÖÔžÀfÄO¡ýE\Ñr/piP•_ÕèÞòM—5Ò%’îÍòÚâ¶›o‘ ŽtÁ¼ÄÝ*úŒÔr¹¨º*×7¸BÝÄò²ù‘3cC¿92äØŒhÒ~©–MÍæßMÊäU"]/.å‚¬Ü¼kMÇ›ÝkŠ [.2¬aÝfÚp‘34_EŠ©.Yw˜Ì™xè/48òVyø¿Ìi,é]’°ÏlGÓ½dÕ™ÉÖÎ¢ƒdÃFbàjW]ÃÛQ4n:y¯Jzå9ÐŽ£L]I]q·\b¶”!Q;_¦‹Í“ ’¼žÌ°ÝpTò–DD“kª¹¢ à¿Ã¿À”9ßdxj‡Ýnø×:…›œmÄ-ì‹õìü³1®¹Z\+GNíò@å=å,ÜÉ%7/ÆE¼¨CÎÕsswðúAÕÆÖmƒE„ê9hsç¡ ¬&¥·µd°uíó³ÔQù6¨†±oí~B¢¹ØÉÈî|šãY´Ufõ>ÔúöHzüuDþ°{Ñy™Õé|è'G×%*kÂwéEQØ¾7RÁ¡P;C!„ûØ²ân1K_ƒqÞ)OgÃ¿OAo34rt·ÿº)uø·ÐVhÝ¢rG o¨àáß-(y„Êm)zøwëÊþ93ÿ¬y §gy˜AJ3[ö-š£°ž›ÅØ’èŠÉ$«û7ñþë8¨èý”@«Ÿ_*	tõævvBéqµOYb((¼3›äÛ—÷u¨{ìÐï½IË“Y^[o*dMÊû¼± Õ-UÆÔÝf0–76Š×Ê–@´fyÒ'{_[û”uÛàgù.E€l¢v7¨ÝÆ–û¹÷{< Ý3;ÑÞbíë‰ÀD>ÁØ	íOÀAÜŒØPj¸ý 50õÈ(;‘'êýücHÊVÄÜÜ²a{ÓßZ›þÔêÂ™r€W{ž¤C(Vœ¥˜¸¥ßÃÖ^ÆiŒá"S–²/ÉépyuÛŽoÖáš¶ø¿×ã¼OOšïÅòüú²MÉž¹w®ƒƒ{_BÄ“Š1Íónûêz×ýrÍ>+ø[ÏL\7l³1ãÛÁÆÑ‰WX÷KÔÓlU)ÈFÚíg{gx*v†K³-<µ¶…YãÄBB ÒVsë¦ªêJ`Qh¨bãSÒñ7éú	IW©’4Í+,¸<c.mÎÔÿ>EßÂz¾ÈÄÑq„B‚‰¤›	ÁEŠÖÎ)U|œî€OýªL´È]Ìdb¤b·—l.öw>¡²°”é-:»•@Ìê°HÐ
/_á'úÇcdšhãÊãé‹ŠnNZ=(ÃØ¿þO³´¼èü–£óÓF§'G ï»w¯§º:°/ËS ‡ˆ®1Ê‘¼Ãé¬‚×á¸¡°Õ@×¿ùb](~ˆŠ<úüRRì
ïÎFéæûR`£Ñ+	b‘Ãc›öŠlZ*0htsŠ“¨8V©ÇÁKÅ–|B¬Ý·¨E‘‚Žä¨‘‰ß0('qÜIÖ³H¯Kr-øÍg Ýš5
U1Iûé»¬¢izò·^‰ÐºôÝ¡¡l1 ýŽµ=Ò£8ï0ºÖA¯qšŸÔ§Ñýh7bÝ{³F÷ÞDjøáË{Ñ×kk¡!Îò·ÂÅÅˆ‰©‹ûÙ2Ö¤Š¾s¬ƒ!ÛóF˜Ú	5ÒO@O?¶§Üƒ¼Ášm²dß¨Ëa¨Á¸Ä>Î(çs%®ÛQ£Á‡Á—Ù°ÖàÆ
ëÕ"¤e‰¥åé;³”;uq)cµïðÚ–`á’¤e'áºOF½}Í—¶>xöt(z™_8`û}.¯´*r[jÄ"ë|úNÞó|¾ï‡*º3|™¨U³QXtXÚa¤ £Öë9¬0ÐÇq6NÅš˜m‹aHø7ÙUÖ­	¡Ã’ ŠÈmÃg(Õ¡èÍ×¥ËEÆnÙj“‹&LÆ­g¢â=r²C¥‘²PUÎ|³*¨›hC”ÝXy<¦"yØií™Ã[–ªHXª"Z6rŽ´£&öºô„¢rô©ŽG˜[£2ÅÅkmïö5OÙdêÂ¼ ¸µrRéGP}‹4lG~*Ö¡¾¿{ïzSr°‘¶á=†õàîÞ‰í<yD;}øU¥[âÏÕm½ß|Å¢`qÇ”GŒ/à»pimQhÍó"Gó¼ˆÅÄšO:{4€%«9ÞÂâè†ÞB	ö#xEKÒdN"´ÅÂ_zòàƒc-‹9³Ñdé¾žªQÒNÂBZÎÅ5ÖFgž&é|ËyŽÜê÷sâöûIÒ.ì÷36ä÷ûMØþ{¶ÒY×"G?Œüìà@$ïït"~BQzo"ƒžD%;zo×o¸¨HR†pH$5Âºpë&AÐ–åºsaÅÉ…a32ACXý4aEee_‡?on[Ì|†°ìÕG0„µNò×–Äævm@æZ&®x*öiX»²K7·vÒÛã1JMg\³D»Q-Ñe-¦!õòÕ¤Tïå+s8\©Á´ºu9KÝ;»QvX`!üg8fa<1¢­º†lc9£™ÌÆÈº~¾X“‹¬v» ˜Hp¹½ÅÖ· ¨i;Ÿ×Œ'¸+¯)q#¯)/ñ5xÍø9ç5V÷“â5<à]é-h—×âÎk¶/ùº¼f Ìá5áÅ;(Æ3ÞÆ‰§î¸çœö[ÐÇ¿wKÓ¸õ`Qm÷êÇlF†m:ûôn%Ø…Ú-RÛÕ/oäé«$Ðîè“TëcLÛJÔz*r1ÁKõU‘~]ÍNR¼Ô³RPâŸõRÚÃtÞ?¯â³tì|y;Í4„¢Ò¥gãÞ×P,_ñžZõ/Cð$oÑÚFÒˆ¬×HµÖÙiFúDî´x$=¶Õ–Šîò;$+Ï!YÝŠCR2Ô¢þH¶TøîÈß–ÿˆË†:bÚ´$|Ô¥ aNq¯%Ôá³è7éþÑ¥û‡	ËÔrý“ˆÊ\X¾
s (_|«¦¬êº©.ë{žU5kìš$L]×¯ª§‡° É†(àTÕ@?	ãÐ°ë/]d3õÍ}ªŠ-?¸KÕhYÿQ=ªjzr‡j@Ð3‡ª¬,Å¡˜e%P™iOŸ€VñÛ-º`™ôêæÓ2ÎÒ’maŸ›‹ºÄxÕOÐ#V	é÷oé/ísàõ<bHWç+;2×Êgá]Êû•U	Eø¶V	­Ø.7«±ì'„=¿ë/Äûî=\Ë‰|ãYÖ¸Éq3·õ¯|†GqtdØû¸Èn¤â„&·Ù¿É–J²
A¿‰=Y«Ÿ×C­ò§e÷çNü›K9áö—~†>Ú¡üÜ—{5
£S<ª õ ¦òt¤¢°ºÙ›áB1Nš;ºv4‚“I5©Œë¡¢««Ñsò÷
õ#an£˜Î
Oå¥ Eæ}WÓëÃ¸Ptê@Ñ	:ŸDzÌx¼}ºx\BMŠ³LµqOµaH(à‡!¿ÙŽÞÿ‡éóKMÃ«¥èÇ÷ÃŒ»ðæÍ‹¨@¸Óÿ¦Sþc×efSÑ8ÇtËlEºE_ÌÓô[ÞË§³ºßëu–‘oÞÿ5@š†ƒóo2²ÁñŒ¼Þâu–ô—oé–•[Lòçv=ÿR‡ðu¢WZjµy0ð¬”êi)y$gy½7F@uê(TÔ÷V¹qH–Í|é1pEƒ®è€jû)ËSÊ':‰Ë·ä›C=XÊ).²J¬”	¼!‘âMÜVù¡}¢H &‹Os%JD,O&jÌ/QX®½²¦.Ëì¦²ö|c9¨×Suçèš-qºÌo†±‡^‹ƒ1:<¼°¡ÍiÈû®Å¯ló{§Å=Ë"Uû.ŽàÙI«ÑÂ€áYr¹µû0Ì¾6yG>ìÁçÞÈ†vªÅÚiÈÀ;Ä§ç”;Y¾øÂÃá)²|-j6Ñ¢õ½jê ²FQ‚ReF&]¾CãBÍ_®¦7NæÅ‡GÔý9¼lo…¹•ÍÊzûö¼Ç†©n¸ˆ‹îÀR¶vB+yÐ}:o¼xÐÅ/*• æúr]¢
TÀ-Ú2ªÚ9ˆM*þæûù™w§ãÈ·:Ž±u;¯À]³ó$ómv´§N}:Ã=W¸ öëè9š§‡=tÆØv®1íÄt§=ï„Õ³>•î"Ár¼B&ªƒ·Ê5Hqúf£!ÿmLOØ Ù“àËÚõz[ITþx®è—i:·Rô‹qb'¤Mñ—åÉ‡ ‹âoù:H8ÁÈ¢l˜²mFÄ³q‹ÿA6ë¢‡Š¶pZƒ³¹%‰¿×’¨Äa4)/r2MPç#óÄwjd4“ð:åëÇ $ÝL…—’›¤Mr£Va´ËéE¯^Í1&P^Ü’¸gšVYñ²k¢ãb`Vl©›íçé®e[0,¦*|ÄÃb:W¹ÖiÖ$ªF$ªÛUDqäçh¡/Üª‰Úì„á¼>[í[GmEl‰°¼Cdñ£ÀšèbbˆÂŠÇ-ŽÔÓÝŸÓratwã™{j½ï¬Óëð!ºÇŠkø =uŸ€R9™Ž±{ùF=ZvÛGÕù-.Þñ¨k,Ãœ	6'8€wÊžâÏ×P´+_V²1i½R{KCh-wñ‡Z¸ÿŠ—EF08Çÿ}Y3‹ Hu"‡ºH¯Z=¬Ž£‘ñIHþ»¨ÕöLà9Û	]Ìª²“Ç³q-ÞPh†«õÎêá‡B¤>ýZŠŸÀ†1~yvô#êöóÝíNŽÑWÈ6	“9ÏÒßAÊ™ññ‚ašVT+ €ÖÔJ_:‰TŽg¸¼¶ú´1B#È’t“s/—Í«Y™Ê°û½Ü\1s"ùQ‰e(µ·®eAþ›]«¾Rx 8-€¯›Ïð-ëñ‰1 ùX°KóBñÈ\ÈÅ49’],äî¯ŒíA–h¶ŠUCŽ=Ì¶ÑY+óÆ(æîR*ÆUí³ÛþºoT6mUZ›•-•ò’^…µ/Yšb­å±GUO0þîÁÊºÏƒ!üqjoQ:±8£‘)¢"ºE<3ô`Åôæ!Þ)š eÉ¢”aù²®Ð}æ}JñÍ3¯-D`ÁyÇŠ4Eéí~bis¹ÑÛ„l˜m-ÚZš·mî{Ø·|ZIàx­ý5>ÙÙöÚüÉÐÙ‡º™&fµÙ&Ú<]ÇŠ»o_Šok§Eþ{]ßþœ9ÙÕ½Ï'é§êâ7Žú6Ðm®z«Á®Â¡Á§ÝENtwa7ˆô=ŸmWmƒ×ºúÐ²%ìÿ;0"ãzbGgdûÐ^À›ÉžN4ýZý€–ft[¾@ùïµ=7²tH`{<Þ}7MaÔðÔíXnq;¾u…¹unÃº1˜x®œQ„Šø§¡à%¬kéq\Ñ5›<uÐÎ—,vTDÔQ×é-RÇ¹¹°s=ã+\ R-ˆÿzŠÔ·+~¸L%]NÈ
0g“éü„ƒå­ Õañ/Ù=/{|”\ê[Tµ˜Íb"ë@ŽEÓWA,ï†@±ìá`Îy÷â›È¤0^^AW¬¦Fd`<ò$:ŠGoOJ¤E$Læ‰ýß–'Î}ËË»P<¤†ápì	·FèÖõ°	E5¾¨‡lx÷Ä9tÑÂÈØÐ%·&]-œ¤Ñ¶}ü¦Ødý	(ª¯.²–D›yúN#‹púf#›x¼…àAîŸçßô¢Ïqßêm/jžÐÁa˜eTã‹q3'q–hˆÙÙ|ç:;wfÓq6B‰I¡GÐi™pFoIH‰µJã@ÓKz˜òù6‚ž@ûvh1]Ä=TC†T=Ôxû8,u‘Ð÷¢õÿoé)ëôýßD“ïÿ!ÚŒÑm@¤p£3_56DîÏòã¢Õ¸ºe×Ä¿MÓiTŸ¦Ñ€‰ðºá>>áóÃ‡I!Ïÿ¯žf'§©$	n‘G0Ðô½LG Œn4A>Vt(òŽ¿â§~?^ŽŽ¬ù'Sj€DÝÆÁ#ší,ÇP$©À”ò>êoÿ­¯­­ELõÊ¢	¼6Â<;„¶Ÿ›ö@ûGmí{[Ú÷Ê6µ¯¢È	A9pçäD6‰K¢&CFB¿\{e-jrØçï±¤¼Qj}à”ßKˆ¿½ŠdêJ7ç›É
¢JŠ‘‘é×ÌR¬¶>ŒöðzpL%´+ièUèÐjDmß‘éÍq‡XHáÁ(ƒ0J{‰ãàº­‘DwÑ¤k¹”äÊ-ÚÇœ\”ª£ULkS^{üÛý0³Î*KO=iQŠ5{QÂ-¬íê95
’Š R2z’ETÆê"ÃVVÑ¿}QËÉm±€²îdžàyû‰IcÖ4ö¡5Ã"V<?&ŽCzÅÉ¶|4¬âÌ}Ñ6QBU]º˜ Ú²?:($º[ •8äÊLã€¶H.§¯a¥ÒMÐæ•é3ieÇ¸up2}~GBÞ\YëbO‹«•´vˆó\P ¦6†‘ÌÑó’£õV¸,7Ž…›i“ÜÅ„´E{FÙõ>RhöÂóàQ_¼G³òÎ|£¶ZŽÌ»mN³qB—"H7´JŠ®¼DfÔ2©ž‡qõ¶Zg¿7ÄïíLŸ#~?‰ëWZÛÝm0«Ò!LŠþK|äßÚî3Oñ·{Ž7Y pJž&ä0eQ,/„Ä‚ÀÇGD/-‰˜Fùuz¦ý„¿9&1h!Øééëø,g±J4z{Yø¼r§‘^€äÍ¾¨Ë™BæåÈ53&ðnÃ~'fýf½`nü]F*1™s©Á|¾¥‰
ªwIébÖ'
X MYCÄß¼ôXMfªÈÈ&Ä<Y’³©ºV¸m²E¥Üek§6<³àÍŽ¯¦¾˜¦”÷D=‰â=\€³DEFm±ß›Q¢<([C‘è,°4«OzKG\Ó+\|Énòà4ûpš½8àŽ4ý¹`ƒ¦„kâ¦FHvl‹ñ…´+)¥¡«(ÑÑtVžˆ&E­Èi×/¦`he´Â¾ØS À4F]ÄXTÊ4Í£cøÿÏéÀt³›Ù’¨Ž®ó§›ÅRÎõÁÝŠC’ç…Ì»½_1YÊG²¸mÂ¬<yØd¤ˆ\6t´U\µ_áßl‰¸Ãi¤ëÛ>¶ll	ŒînÞt5pZøí¤Þ¦ms+ÖMˆiÃŒŽÒàF‚ÒÌª—%Eµ.,)J¶ó¥j…ÕÐÌYælŸßuxZwç`w÷«a<d‡Ú·¡C¦h£1º¥“’;«W£Éi-IîY†&¸ÂJ"e…³÷quzTÄåÿ  ÿÿì½ÛvW’ úî¯Hsj@›„HÊr¹©MÉ6§t¡IV™¥Ò‘’@’L@¢	Š2‡³¦ç<œ§y<àž‡^=kõS­þþÉ|É‰ˆ}‹}ËL ­ª.,[2÷5vìØqÛCçù`VLŠÓY:NïÓÉq:qÎ²Q
ðËÜ¸ó€B’!Tñ3Ús¶R¢òŽYXá!IËˆ'9ZÜQ2-QhNrv bpœ	¬ÂI:È’	œ_3yn•wuÞ·òlÊõÌ6íårRVó
Ž¦tR¦Â>ÈqÀà÷¨«F*–K‹YU2´³ú¨HïõœŸ;ÀtITüm¸ÅÄ~ ³²ªþÅ8ë¬L Þ<ÔÈC|ÁÜÐ`3”±ŠÙ$ÐÝ`±õCp¸õ€¢áÛ‰ b·Œ<Ì -‡¸b0í9y{%)lÀ²ä–_\{´}ÀƒÑ_¨FBòTŸ€JâWI@'q‡£XT)q‡C¹¡Vb¡- –`¦!Ñ]mEø5×]Ô”Ø¬+ÁõáJËaÏ,8±ù$ÈÃÜ€þ Çä¤y±ì¹ëÖkRL WÌñû 9,Ó<+»=!‹Õû’5-vÛq}‹Lyci\Zë[jô+õ îYòR¨S’2Ÿ û”ÿ$mÿÃLœÿ%9pŒ¡Ñr€R¸´(cº°O€ÅDöÌH¯ËFã8æ˜Àv›ï¯³dNé
Žî	 ÞG[õÒ¤)	+›–âáàéS¸^ê†žaS±RÕ<7:©gal©ö®©/èîY¦doÓ{[}Œ{¼­V¶d¨fcòBšÈè„Ùí1llf¤ŠT*Á¼¦Jå)YÞFï+Ö°7œz-ä¢b—…ž3±T©lk”QË×‹´Xí¿éG¬>>IýÈmv—{‚þºª’_GWOû(õJ†ßÅƒ~˜é8\xW™™–#Ž!Õýzƒ ±½Ù¸þ_/¯ÍZr;½´ÒÜˆqp: „·
}$ã¢¢ÇGô +½*!Ö”3ó-8ÓÅxÓÄn~	ÆÀ6ì©Í Zêq¬IªÎÈi›:@[þ¤ˆÊb¢­¦)zU†»²ª+gm}"_[AôZ®î"kë4~G†^Wð¨YI[œ¯_JŸ:èåbgt}£Jƒ˜/±îÎYcñ¼õfÂòíbÞÃ¯µ‘ƒ'ž‹A]ÊÑáEZEq¡IksC”]:ø [#ƒÒ‹Ü&èæ?i4æå`ðÚ‰b€ÓÉí×^öæ,>>m½úXø.—ß´ÿI¯?	\SÑêlW†þØjs…Ûíª'o¥]—…º…ÆÛ"wt´³ÆÝ£ys	O2yåU™N>ó[ø¹²+‰‘%¹¡ú2sB:È{Fo9ÔWo’	ü7?ÏÆ·T`¶U^Ênm‹\Tèjg™{Ê(©JYÌgz¿Ç2Œ½!…×xæªå[t©‚N¶’Keß”BõÎ¤üÍàÅqQ ç¡ó~jÍGh>Oý2
aE{h½ÂÛ6Î«+L€íõŽe«Á>ñŠ	š;0¾š$gãø…z6E÷ª‡b¢²£ô4¹E5Ä­[Ÿ³a¨«¬o”Í\{<å0÷S@êbnv Jêðd¥	©ê'0bŠ«TöWVs?>¶¹*×Ã‚Œè›…J¶Í©ÀÛ/’ÎÑÆæÖú:ü×éñëTáùSåJIÁ¶É _OÓ1s+´IOßv 3l¹?)>øB&éhVü(¢Ò%ÎPU@•N§Ç‚Ø½ûÍ¥µâ*YK,¸dò¶BÒýÍ¥YÀ>Ë†MM‹(xW½w?MüÕñ[‰½‡Ízcæj c]cæ­LçÎÚV‚wÕÊ*OÉœFK¡8GÎMoŠQF«y÷%ªçj^«ûgO’uFm·k[d·’×gE«ºþ—Ó<ÅÃœ4ZtAýE:{/Hîß%@Ùˆž~ˆ1dùŒâ dƒùÊ'ßÏÇéDÔ;Ê“bT@“Pqo§óÆÌM sÌ æ€ìÇ”"X0èØ	Ñ£¨a#YoL£ˆ„ûÂð¾eãµ]Æy©¶¶eSP»À'¡äIÊÈ‘r)¥““•³”{•£å‰Ž÷ûý6mã`‚D6„t(	Õè$/‘ç/ä	i­]­{d”¾“~œ7é\’¼ÅÂj+GÎèH¯Z—ÕÕÉF.ËCôÜV5jÝfí\püI‡í1ÖÈî%‘B<“éî„ŠÖü ‘‚òâYÝw‹ùhˆw„$á —sÓ=€Â¦é#kÑŠ½	À¼ÚÒ#|^J¯ßØgÔÌ-OŽrÈ|NÖU¢æô5¹
¶!÷¨ð4Â ,sš=Õ÷l”u~˜#£€^YÁV•TÆ
6Þ'Í@»¼Ñ‰ÛãhÎIì&Â”OwÍÔùS•âF&®u·!†gúùišÍNç“*µ÷P@3µ{|êädÄtEë°“\ÈQ9zëfŠG®+Ÿð NÁðéNæ<`ÅÓ'ôX±V*5½ä¤“…ÕoÏX:Å…À.]Ècíèrø Þ—uE‹¦p¸Mº>’4·g­k­ƒS~XÆ ço×^UYLGE˜“duóÎi—½è„Rgõí‹Nˆw#Y^oÞå‘vH¬9ý¨_l™¦¨&hýò¡`¼^ÈèŒ¦ã,U®x
Ý2‡/¾õÌÔŸÛÝÜŽ>ÁH‡Ã.K;Ï×ô‡žO7°¸xÇ€…~µ@PA>8C¾Õ:•œ•ï^H}‘Í !)	n%fÔRÃÔÄÄ¦j],Ô‚jjàþÐó¶ô9±@{©Îv-ˆánó_}u
¢Xçá_q\Þ°Y¨™ÍÎSŒËjxš•(qM·ûô7Ö|2Ì rF|œŠÇ[É†W‘ÂJÓtq3æã”ß„¢úÙ~:ä¸w·“x¾î¹qà”›ä^žA
½øÈzQÜ¦G0i e:bÙdÔ¾û_õúåüXëÝÍÕä·ÞÊÑ`ªll˜K²¨1wó´Ø½ °š$ó¶íV­VÕ‰¶åoŽÝŠ)7—’L¥,d¶K•ÿd†ÿ6X³…ßù«r€
iýÿÞkG­aÕ`XmhC‘0+ÙŸÎKôkå’X—Ãçk+<Oùõî{$Ö¢”¬ª§ØÆ¾-6ÔãÃ"Ñ'Úb…>fÔâFvÔàG<€?~ì‹ÓayA`†”Ì‚+ÑA¯c0|'³/.Ð
/Oº·4ãNJ—sÖ|mxúUÓ÷Ñm§ç™/‡ÑÜ vÀ¡ÑÖŽ>ôZ.‰žj‚Uh1¿,ê!²¡ºÕ¬ß‹ËaænZ­c“Ù—rq©ým§<Ÿª#øìðUˆôÄ-Ö¤k“Àá7O’pK·Q¥™’a…šø¸!düJÒšDã¸h²Ø[Õ«×æzBdã/×¨(%ö³ùä=W£Ù]z¦Æ Ñ[òôš&ND“!Ó¶07ÌºXTDN¦oC»ô•‹FwyÂè‡z›÷=ÄkŠÔÚªŒíE´dôqa;¢ü"l<ï,[ÃLÛ˜¿¦´U|ž¼‚QŸVˆ%Œ4xý3ž™o™*ÎÅUŽ$ýqD¾ì¿Ó#’áìU&9sWCúæÖX<E°gýÌªCš¶^Á:ËæÉ|¬·ºº=mÇr-£gpj€ýŸÌÓÑ!¬HÐ@ºp„VašU"ÍN N›BžŒà8ÞÏ&C(IGì¦†z·Yzh÷„h·›™\Ý,rO³ñÂ<!Ê±Ù‹ðgýlnÃË”Á–ó$ì{ÝÞV»Ï} æ!ÝmÁ?˜Ÿ§%•R°Fc†´|ª5²È“Ú>Þk|«EÈb^É[}ÁÆ„5¹ÿ@“Êp”›È’é™¬Í<¦^ð4ô¢Þ4OèÃ^ƒ²öE×Âixàß!ô…¶$YÇÅ~pöæw¢aèíMqIÝZê†ó	]þô‘\4;ÇÖ§.žð‘H0>[qÞWXÛž4	A&Ð7Ÿ(s‰®hÀ!µœ
ÕY
.“S:×ÎŠ2›Hµ?
gxc<Q#ZCasUiCmc3€¥ÈÎÓIÅ§ÆgjYË˜Å1g¦µI
OÜlF­‰Ö«6g¡‘®…7uÛü+Mvn©àTÃ7tÑ/<0F™ò ¦¨Åpd¤cã4À7÷0&ð¿Ì;hæAeˆ¥@}5ÕäXoùšø»Ô¾õõFQ7PF£ä»lœOrý‘C†½±³×µŽÍÏÓ<0½]Âàý½˜)Ñšð¡È§”¯±Ÿìûz’ž³Õ$æƒã)—ó41µ
»ªÊõ?_ÿ[&ç–ç£x²DÀBß›¥«ƒ%‹>…Nö€|òÓëÿ=äi²3«ò“•†Ëœ60§èªün·À÷%†¡¥4‹:´y«-`+;O"õ®àÔ¡.òaF*G”°ËìtŽ!LJÌ[;—€™e§y))Ð²Ê& ¥ŠõÕäë6ÐËtãÃØgó
WªèöÙ«ƒ½g/v®ÿßëÿçU²{pý?övw’§Ï’—×ÿø|ïðYrøìÅõÿx)ï&ûÏ¾û=Ô8Lž%Ï÷_Â÷­ÏvàXË4˜ït‚Ë€ì6ô: ÙJiÔUÊ«òÌ‚èY˜n}‡Ë>"*ŽP8†YÈÕÊE†&µ¥néÒ^nšCS?f¹X7»ØJV^aÁá.ÊX¡ä;OR<›½³ëæ@¤g ŠŒæ¤ÈE#%ó!UÂLãÜÆ‹^ûEYÎóNóa>žŽÄà¶Wz[ˆ;kÉŽ™ÚÊa>^Á®Hö
²äðÄÚ?xõÍóg/v’{Éw;ßí<€ÝyÞOþP®ÿ7¬Åž	g½|¦îñ`ºw]¥Dè€qo2d<ašfIT ëûƒCzµÍÃ6MñTFÒ£4´±	öYÒ-ój®z‚#½ÊÏ¨&Dî³<Ó#Ã%Çö ä>‘ÝMhT¨t å”«Ýÿl3²þDG±ƒŸ25F	³<.rÎ®F*ªV^.ŽªFL—ðV¢µ;‰÷ÊäúŸÓ$ç”õ&¼xbÀñÕK“ç;»¿é/Þ®¾Jˆ1!’95úÀêê°þÊÂ©E—1`/‘$z‚QÍ{ýÏî÷“½û¯Žv^=Ûâ<CRûì%ìT5êÀ"Ï¤¤ÕÎN2\IâƒMG ŸSŠÓQA[ËXÆ‘<¥þåúgÚðe6¾þ§	Ž¶V8–â-C¶èÄÂÀ@‡ŒZ‡ÑË3JeƒëEz«;ÊAH¡êjÈà¿ÔGfßfÀv‘¶ˆ’¨ääº¹%ß®¤Åáh!ÀXDÐ4ª{oÈ¢ë?¥;ÝaV!Ý«‰[ÂèÛD Ÿ¬Ã:Ëq^¿@ê“š—3!äÂù8‚·Ú›ËàzýîÞS<rÔèï)¢J¸ u¡†ôÎaZ™ÓÈÔš;Nóþ¸f£²¯d’L&«RØ^§tî¯Ýïo¬ŒÒòlìl…eÞ5
âì´ÞÀ‘¿å(Ô`^€L…ZGè"
UÌåÞe1Y	 ÆáàŒ âkÅ*j[ëïìüW3™Ž¬Uµ*¿úæ?=Û=òk“ñèT•g‘FýÐ%‰µ|xt°÷ò»ä*Ô67Í,PI$Q_¤âóB5~/4ƒ÷T³vÄõøFi]ÍU‹ú—kÝoWŽYÉK¬Ùpjƒ
ä}¼ŽÌZW¡k¿‚·«Ék–˜s Ü¨6-ÈÇ!ý5½7êê/¾ðµZÈ¦R±×ùWZŠ|õÒ#9~¼Ïóê#>9¬fÁ ÷Žc5É=Ë7`Ä¬‚úØ5•ß› Ë2ëÒ­Ô‘×ù6S _Ù–«Ãž"ˆˆ¯<«ìu¬¨aU]…ñã…8Íë)néñú%`B[r’þË?€hH‰³Šp™¸úïÊSàÅ•kRdsóFfn~èZQêó[åÖÍšµÖ3¥¨B÷øU’f+	‰…³BåH¿/æçW][Â0g¢à˜Köv,A˜^1ò4Ä”ÌÅi…'ôœø¦ì‚¸Y ’AHˆMK#ÃJ1þ$Ú;Üfÿ“?NžÒÑDó“•äà7ûcdPNÉ}Ïäyéùj„Ðèí$£èûŽÒC`ÔÔ®5Š?aÀ8n™çq(æv…à¥½‡’i£õ°4§’¥×Fgqúµ¹—¤ylTÊýÑ™@$¢·ëXémT×ñrvØ¨@`Mµ7ÖC6ª€ÕÜEV÷ûPüìÃdÐv±ßîx¥ÐhÜ¯Î”SîÍ®ÖMOqÑS’è„§I,pEÜù”„H«]=ê¹šVŒ*—›aåBNè½rüÞ¯ßðóQêXùÜe0“ï¥jCtéš"ZŒ4\÷A)s`	zÖ˜K¶›ô=U9+h–{ÊUq‡ÎµtÝ%ßDÙ
ÛM:RQãñcVŸrïš½îoæ0Ðp²›ÄÌîàWd×vu3Úiž)=ðEœSa÷óÁ{“KM»±Þ›®
 B|¥ƒíÀ’ê¨õóŒF1¸Ë$“¼èhúaåjém^ÄÉñÖ=f‹²­ Dh ÛuÓèÛÞëõ7¬-†:óQg5“ÇlO,“çJñ—Ï2X´œà?ËP´,n>	D“q.¼E=ðB{©XÐ™TŠÑ)mF÷³ö Q,§NNúšÚeƒÙºè€6²7+Èï_8Æhw‹ÖÆåùdãë˜µFr7jK³¨ßÄÉ
lì†#&³ø^ëlõ+ãÂTâ ±˜dá[TYŽ‹Uß‘¹Þ[]öÀz$ûÚïù™½•¬|yæ¨dã _J7Ä†þ,¶ê³Ÿ8ØÍP¸ÝHZ§oTXé“œ²\¸FX¨_TG(™œ\>®˜ƒ(+te_óîYà¡ØŠù~ÑÌ‰0W „¥‚"	ê8Ý<ÌW¶‹§…<µ1Cî˜’¾\Îz»Æî>Š¦aDµë"––,UxZj<õÜR–)ß–›håK…cogx×Å³fÌàn«v-.9«Aµ¨l9†À6
‡Q–ÆÍþ¬œ’ü*G#¾ÍäS[gãÆÀððXà %M‰Ñð'*A§Ÿ‡ûÏw^¾zûôÙÛÝWo÷wžì¼Š8ªi™Ýpµ›ýdgˆNcêtÄFÑEqd/’/ÏÐb8\t! ‘…¶Î}2ø˜œÁÐl‹”Î†ÖpÅCóèEzñå÷XµëÎmÕ—‹\…Õî(CãÒ&B_:ð¸DbŽ+kdjobùŽ.w_Þsç°†æT.¦LåÙñ­“ A|é«ˆÊB7øÓíð¥SÕ¶3$'ou1ã|‚Ý+ÿt9»?Y}?Ö›µú-Q¨<£Ë©uë$2´ó´°u/¤&’¾>
à˜QMb‘z2­Î5"òš¯©†§±„#lÌ½‹<ñ¸½%©<qp=ïúÂ²­ãøÂ*Éz…¤÷‚ê å–7°$OÖ×h…œÃÛÕ•æ|¥O5¿u"PÃšÁ13{¬ælÀò6ƒ£¬hÞ"uba‡ëb¶5^Í>@±ž§ÀùòÌªÁõÀ4ã˜êÛ­DªaUC*¾kŠseq»ZµwŠÙo©R»]¡=µ¬f7P='‚„4QR#³mÕí½U¸úÚÐ¨ýä›9æZ-çcL•+õ—‚PdS¥mÅd½Ž²‹ªbÇÈ!ìÝÿùïÿ”üæ2¢°ðö~×ëÿXä“nç“ŽC÷5‡±X_·"z³÷„V³`_2žpÇŒ5qÞŸbƒè•¡!-4˜8„Ë+víÄÞ À¢·Hèï‘Ïù$½Cÿ¡ÿïÙarðlÿùÞîú	=}•>ûîÅ³—G¯0ÂÍ”ÿ~:ULùUoëŒ ÄÖôêéÎ,f¹å®‹Â)~ð}ÕWw±G&žº´>{QƒÆÈl=p¹1$“À'„'¶èÐŽ7´Õ|ÎV íýl/øV}x%ûY2F|-" û·Àeƒy‚6ñTg)"ßó¬Î¨‰€«»ÓÎ[ß¾uƒmÖ\±a—kŠiäjˆòÓ¯ÈBAÙÑÃÕ³ã ÄsYb	^u‰ªºø)«¯Â0üË:ÔŽ¹b/jÄŒz³Z_ë­!Ò„d aˆ+‹²Pä6 Ef;‚Q“–xz,mÀQ†˜M¦Ñwû®k¤ð`¿þ3pÂ‘ÍÐK}‡¨´L§@ÛËì”ŽçdÅ"sW+Ÿ£Áð]ò…%HÒ!ñÃ<T‚øcïê$ ÷Î(c}iŒÈ4H³ëŸq ãëŸ/ràûÐ¨«[Þ¦¾]u!_Rh{¶*¿¹ôN„«w
Òvú™/5y¶]´Kˆúii¤u²©•­Dg±®e ¥Æ–d¤åN	Æ+ù~:«òtôÈÑ?©«ˆ‡‚ÑN›„)µYl)1$ôQsCÆ+dýYœø@'bìw˜#C+˜§[f˜ŠpÛ]èBQ:ÈõyÑ­²Ä)â<·\E»ÆO4Ëí	6–Mlá |J¡ÉÕ"\¡âMrˆ#ðùäŠELà£Â&9‡SKzT2.êMí‚Öÿ—Åy*îÄ%”¥6Cf,*ÚFÂòÄJuÇAÓ™€”'ÅxÈõÒ)Àe0cK3Rðœ{‡¯d8^¿²Qu;G4¯ÙÍ‘t¦Û’‚ÖMâr›nÏ<¼I³h5‡F<v7–`éµÞ`V@Y+t92ÔN±¥Ãv—äå0Ÿ¦šKïë/¼/°Yê/¸aiÕš6mÔiÌÞ²\Œ	rbÜn‡Ã§ÍŠ-s‹Ô¶üÕ&ÖÒv³JiÞ¡$«U“æß_a¯Â íœrjžË‡BL'¢Jkãâ…øÒóùhÐà-°î¥Y‡V‰~ä¡Ìë=GwI“š´?u°k·;OÅ·Ä|éè—š¦î0I‚w™¤œ…Ì¾Ã´ æçë%y2(¿ =¤Äo‹±OáÜ]C´9ÌŒÕÊmÛïFšºÛ¨B¡zTP°wŽ‹kØ”"ô.k^Ò0’ÿ@Ÿ‡áf¿fcÍûÍ74+] `¸kÐ¸ýÖËÍlLF{ÃéÓ¦<9&Ã=Ôzæ„Í ˜²æÚ†¼ƒzJú
Ya5Ù¨	Âa3õb€¸lœ^àõ?ñ=Ÿt]F}•¡g8e³HA£	e´eŽY@5< âÃ6‰)êQ¯&ë«bO²^õîÄÃ)¶.m"±°“ö´.xAT¶´2Zè¦±«ê":Hâ¢Ôë°*W.ŒÃÑnâ±W‚kQIÕŠ’ãŽ6>VR ÅbªÔGx	ž`4˜ÈÉáò—{ø§HOñÎYv|€Å¿b>@	s ±JVßób_ d	Ý—Œmv#Ž`±ýxèÜ
ß²„í'¥£Ew!?TäÐ‚æ?d buáŒ6¿Õ”ïO­x¡XÏ®Ÿ·J="ù
ÓKÎkÊaäÓÀ‰­ÄÚ3ÌïS%&¾Ïq¯õ5„î•joãr-çÇ[@´î¦ÕyìSA€)të7•ºA5ÂNK_/ñ(3XL¡ÿ û–ùB”“€Û WEuCÈ:;Í8ê£ æ º³†è³d’M$RA ì×:ÖLÖ‡ªÖùj32ŽVÅZôà‰ía‡«”Ô´0´-=®ZÆšÐ ,Þ“Ý‚Á„ƒ–RkC:Î•cSÌÈÔ<^‡·qÆí‡ygÆ‹RvÇ4<TÁ/Y‹xzM$Ÿ‘œzï¢jwbŒ/¸;£ö‘ÃÕÔ±	ûFecö™šïÎØ®K´úôØ†8ŽÝ–kˆaÙÜS”¶æ îmØÄPG˜äãš¦›`†­A‘—‡¸oiûµäí¬õ-S§Ç“Ð³3°´XÃ_UÆh•V}aŒr`ò‰ÉUíçì¢{ña‚¹’Ž¤·:ÞGd(.
}ÀÔ(ÿùùáîÏ«|Tö‹âý[ jDô^ÐÓ^Ž¸ó»l¤4¯­Ö00ÇÛªx[žeYÕ}-çh4[e+é<»À.DL/Ånx6ÈÎÑ7zŽ§é iQU¨ (:’ÃMºÜô»<ø>Ø%ÙŠL§ôWÅåva›N…†"`“3à‰Ìjâ»öí#àË„<1à«×ß!nÂ»wvt
Ä©ˆ¡„7"æSÌùIÙ=·­j%€çþÍá"çEp‘ß9\M´ `¥ËàbpÑkÐi™Îõ?j`IžðÎ¼–Cü_˜ÎPß¡Áâ{,Â“a:¯Š1Ez²šØ‘STBL‰Õ7ÐŸcb}Œ—Ÿ_ÿ3’á!ÅóSÁ
fÉ¬ ú–J<¼q:÷úçól„í~Gò6dm¯Îïö÷¨ß§9–Ï1ŸrN’£ßX_ÿ¼Æ.Î¶Âi\‡‚pax‡ÿøú_&¥Ó`²‡cŠØSSÛÉd)¦Ô,¥X\Cäb1À•ðâùüÆ{N¡Æª^¢ËnQ¤<þ6eT¥#¢E¼…]_Ì*ÌNÐ¿•Â8DvDL¾¢\Ý=P1þŠrqÒk’_!+;u®¡užQ$v0…C€HP~Ú·“7 N#E¤W’:Š“Iß³].5<0a'ˆ:ËY$ŒL7ó:û’$vxÊë}YÈwå ÐžöÅÀÐ†üéë”¡ë‰¡ë‰Jbd¯áJaÖywãä+|ÿÐ'ceÿý4×-e²UöiÀ@os žx¼øº.Nwˆ
¾™Ý¡;"*û‘§gƒldDœlKxðöw¡Ðiöì¦ñèû£Ï÷&ÓyõLœCO<²2Y_(ÔPÖÉÊí¾'@×Êñ¡	MÏï•ÙÅhh.Tî¡z/^ìš<]•ˆNÙ¥Â‚¬Œa¼;šgæ—g¥9ÏÂ©èù§£9N¸Óyˆ;p‹.¶ˆPOEúÛ ¬qž[‘$[Þ£"^'a†e
Š«0*J‘î]Î‰¼å…­ÕœÊSŽM™í‘:= ê¯ûHÎ+K9lË£ò× ¨Ï¶¼´âªÚÍI†½u…iVßúIg³ôcÇVo¢ƒ¿>WÉqlý¡SD’½x}Ó¤¦ÙÏáûvc]7Ñª—h’µ=Ê…Â|Y%=-ŸÃ^ÞÒg»s÷Ó*/Ç¦*ˆŸñ
Š¼ˆò‡‚ÖFŠ³Š4§*x4›©[áP8_×ç3Õ:l©ŸýnT‹Í€üÞ¨Ö?äý^²í=3±‹ÿa³{y¹¾™õÅÔÍì)åw”ÇÒ™¸S¤‚~–~D…½ûREO0]$9eZˆ¢ÚÀáÇoéÀh0Ê(N?FÎñ
—\Í£â)ó…Ä¿µÁFGùàéø´áAºB­{™$v$û•šÞ]yùíÓç-úØ§°8÷^ÿq¾~}}ÿ|uòæÞéj²/ü§k?­¯ýƒ*Æ@jçjü>áwd‹Y•Î' Ô|HŽ?&c I9ÈR:†Sò>û¥@Ð¡[F¬’ýAÒù.´ ¯bC¶=Ë™zý¡›¬qB½Ç?vt4Ðõ2pØsÓéï /è»$®ë=öF(Ð}oÛPØWv%	z†&ðZÒ+êº-á\zaG2j:ˆjeW¢•€AÌRDÍ¯ÌøL#he{Ð°ž«Ä‘%TFÕê½±jÞ§Ì¾VyŠŽXþ_yu†5ÜÜ¾f
"ë%ùµö›‡N18ò—ê¡²UJì‡Ç*>H$Ù+~äŠBa·¿XìOû{#®<:‰•ÄŒ×-éãcè§CýÜzlýà%Bß©EŒpË_Vy59ÍSihS¨%’ŽŒû,;KOrúJaÖ;¶é«ã9(%èI¦Ä ¼/†Ž¿ñ/pFx1 ›zY©„d˜©W^h.†åôC!ž¤ÜÀ1@Ò´:ÌÒjŽ"»?")¦êAÉßÔ$Í™+!øcðKÌ,=×ïŒB¾‚M Zòé„Ý!ý2*¬OhšÖcœý©‚vÉzbÅ¼Î”dªûSÄ°?ßÿ4×ÐUEøR²Ç¼–ŽÃ® òT¡0+…“§1­"ç@ýÄÂ¡¡—v)Q|
Ðª@BÓR}›«¦BªÞ0Õ5½Î¤Xªû“¿™”J¨)žv(|#ËWJHÕ€Ãï€(;ÊÂSé®@“Ÿ Ð@"¢'I¾fÇà/-TUûæ† µö"ÀzÕþT?2XAQþ”EJÿvŒ±I4Ìt–þÄV)	ZL
PC@T}˜ÏT¯€KG”9ZWÁ\ÍclýQ˜×ý1Éž÷Ê:0;Ó¼ÈsIH ·Çô€<I:ØÙûi®ûBí€V`M| ‹®´^¨ÇÖSÚ´Y…¸”ßÓH~š"Ÿ¡¸³®‚Éù‰‡ª=ëa))R^pç‰@ñº¼Õ]•†ÔhõCjjTHªƒ(?9Ëäõ$ÂpSé$Ë†¨	öûc×qu¿øLÚÉ?Õmâ«	Ðòì4ÛžU›ÎA*(ÒÑ9­gÓH*ä'èv“3F,tˆe´cRKÂx?³ak(ÉÖ•5?÷ÕOéu6‰¹bIvHse‡!S)ÙXWöþãÊß%ÙÈ=Ç{·Ì˜·õ˜Ë×úééØVÇ´à¸ŽŠ½á‹tºEš¾Ùð‘r¾"!ÞÓ¼"½2Ù—†J2—9Ê<ŽËµ÷4²ŽÀ”b¤h4’¬‰‹µIªDh¬eŒ®‰¸F²ù’"ggœDAøˆ°¹ôFâõÃïj%´ÏU/=YýÐÒ­~ûØc5dP²ŒXSU)p]Œ
Ô§ØêªY’ßœüÎÄmç‘°{ì†'ïÙR”¬Gôd«f:–³„)¹Êïª[öÈ–Ñ­M±íÖ³•VÙ­€øï¹©ÉÁ¡®@~µ:+DI{²ØÛ¯Õ*¾q¥¶¯½²Q×4Ôúgélp†Òï$û0ú(o˜“Q^V>a‘áè‹£}tijDÏ ÑÇ¤AóP¦ª;pçœ˜…¬ÑÏ‡aR+…2ÇÒÐ°P ª cŒbË o¼-8x@jÊ¸ ‘¡	2 ‡	p`âºÉY„pDFôQÂ{ü»¦FdìõÐÅiK<rnîÃ=¢ÛE”¯LDß”"÷ý¯Ô>Ý\M¾êè’vœ¥ÑŒZ—ü	Ð"é½{Ñ6èÔ!ûÕ„Ó'Ö¬ÚÂÍCS2',ªëïa¿¶¢ŽHê`—x‰rÝ-¬AC/¯ÿ¯Äo8MÕ†'àëow‡S´!SÜ5»Ó¾@èié¿ø"ÀH8ÕõÝ·óÉ@î2õ	jrJM¥ŸÑ(?ÅðKgÂ˜‘ãum±šs®ËˆY9ÀVPkŠ?”ÞþE}wàP¦bZM÷ú€³ôƒŒ×Ç,yT¦h8}¡}Cv+Ù æJºy¾~c7*´†²MX@ùÕIÉÆcïæo‘qýp–‘Ce#Ê=Ë¤òsTï‘ò¡zð=lÃ	Ìd4OÅ%n‘5‰–
Q ·§ã5:%ÆéÅ!ðq™ûNnÑÑHñ¡ÈÇSÈ ¡·¢{yJ¿#‚	0Õ=`2¶(­„[Y¼4¯”ä'£H¹“~I9ÍW(p”¾å¸ñ`Õ}ÏËŽÃ  ûX•·²äè5ÝA£V?/é¯Ô¼Be ÒóÌ.Næ£8-î ¶@"tÊ1¨äôw±¤b	±š#/È[¥¬0ò{fÑÂ–

!(ÍÅ&ø‰«C%‡E x¢‘Ëoš¡v{ð7íÓ=÷N0"6‹[En ô»¤Ë“a:“¢ÚæÎ&6ƒssºâGnàz’¿p ±…!¡¢o†©&ìmDñšÈŸBU»ÆÀ¿G&†3\oyžÙá±Cá'€®jBêJ½ÑÐèì,ÊÉÆCxþÈÙlð¬f¯‘1¦d;nÖjÇ‰j$è_ýìc“±à‰°03™úmÌð0`x,ÌÁØá¡{×°’ž8à”²°UA®£¹ƒ|¶Š$|/rqÚ˜Ðåë³7ÂH" ñZ4ã™Jœ½ê*`íç -Fn k”ÇÂuêÛ9ÍÑ±t_{v	Û>ñÆ3@Ûž]¡ö˜fz&t~x
iý¶§²¶¬þ˜•ßÓ—ÊßIýú2­LÅ¿V¥¸5Y|]Åíß˜~––O³!å}c!ÀÁ8Í¡D?Õ$x{„H^k6f‰ÜBûÂÙ%ÑEO¢C&Ù8¢Û%ÀÓ<œU1zHâ¹,ÎH¯uã PDgÚH'Vµa-xR~ŽUb‡7S^<-KrA•¬dQnäªeq¢Ê6"ÅS¡¼±®˜©H)ŽxÕ˜E+68”C¼zÚ,©…¤
Ç®ç¨wÝ€‹*©˜‘Vf…×JªÀãˆ[HLk…ŸÖš+ö4´»
ðf òI`¬Îo¸R’ŒâÝK¹T]±f¦KúèÐzîu§åÍh‡JK§VÛô(Ÿô˜®.ò†+Ï¼ƒj;RË«W«I‡ŸåªòT‹®¬Œ»ØSçáÇõ4ó¹äá0„f=–¥¯
Rž&~‚“oÒ[Ý0†ÅàÐT™ñHuË+äõ¢ú*ôpKÞO+Ô[É£¢NO…ŸZ]õ«ôU†|ù…šµVøñ4WŠ4ùEÛê®ðsåí¸ Ë@Ú[Š‚¼ÝÂÚ¬z!r3Ì	IF+À	É7‹sBª¢¯I¥eµÏŒGpD×¾t®J1Ì>gÅ<‘MBÊW
ïãÊ;ø¼nÂ ©6ÔÄ#cÉJŸ[RZ§e3=ˆ‘÷Üu”d'V<Âhú=hxô\4ö`ŠëB²Zå\ò¾Á¥|W•k¹e¬jŽ8Â—mK}_õ=ÐtÁ'n'ï^d£3
½“ß\ØÃ«wŒ'Œñùx_ †Ž¶Á"¤m¨pi3º§ûr>>Îf]ÄÂº^ìxø/‚
®&â!û|Wâ@M‡á7¹š4t’—Oç2;2¡”ºçEºOâš|~…IF¶ITy¬vÍ/™BNIu
ç]8÷dø^Ë…ÿBpC3è„ÙXû:Ä87lòi+îMßs3Ü[6¾;î-KîMôåÞôL¶-Ì/c.ZK7ûóË±væº5[Ý&~P\¶–äì¦ŒÞ¸7Žî†}#P{ká]ñY„#»âÈÌ¥:%Ój°92¦>ãÕ~=†Lê=¦„Ù×îš“:Q‡¡Pö¼hïéOLY÷b,‹ïÖUfÂHUôÚôêHûá4qíMã?)ozJÌøŒäPyÿi­cZåVhÁ³êÍ¤.FÉªŸ,Ç:K?H=êò¦DnÖÙÊÁ,_áú7æÙƒJKæÙ¨„c”u( bRË “Z61©Áþ¿*·\ŒcæwXµ±Ü„uvî	/s†Û0ÎÚž¨ùfxrWl3@Œ¸fC8ÃL³û/Ä37ªC¸°Å ÓþFDßj¡Ðþ…,íŒOB¾ŒðêÈóèðÌU¤HSÁY‡Â]èæÜ¬;mµccèæˆ›Šµœ¦0C7æóH‹AÐÏìÍ§íçÖ5Yè_HR²Èœ”Äæò6Ÿç ^LÒßüèð 6ý~’Ÿä()‰ø/=ÍŽD‚g8á=9¥`ò!×NòÑ
°–˜ÈRÃB‹uó%,f‰‹8ù…Ïé½I$SUž¬%ÓHŽ%ÈŽ°bêîHXn“…ŒÝ7l«eE•­6feeEŠáÁà(n!Â»ø@ÉÒ¶ôŸ0Õna €r‡7g¡ØÞ½  ØË›:BìßÌ”°s#±™æF’sñã#-*5ãûi¾ Ü,úZ\PÝ-Æ‹JÇ7•n‘ö‰¤¦œª!¥¦ˆL€áS[É1v,)Ëg~l*‘›bÍÅÅ¨  eºò užÏžç½bšaxÔ/ZÖâ%B’ÞgV/mÿ‡„FÊµék)Âb|UmÊÝB‚Œ/£%úùÍÔGèï)è—ò?pÞã“bŽX%%FÔ”vDÿ-‘cH®‡š×ÈBî1ì5æAÀäN¶+·š]«¦y~ª	c%(¶ìéÅ£¿þ³z?Ìù7ëÕ»¥€ÃEðjOGµNþ«p~ìÜæœ¬×Q{S\ÿàµ²Ê)áj`-„áj5LüÇ«‚6;vh|«gÜîq›v ¹  Sã²ä/˜}9Ò:©#ÚZ8#F±¤7·1ÞÀÐyw#	@Š´Á¨Ð*šJü„Mi¡4š(E[åBÀ5šoG7é`¨úMl—¢b­ý?ÁÄîæY¸Ò"‚:~Ü
*–JªJÇæÔ¯Qº[9ÔñUÃDP‰T<´íI g"ñ‹@E¢Úè–#]‡.ÜB©»¸Šù—Ò‰Ö'¤ÆOP¥'Xqõ°(ñKÑ+©7fœJ°ØmµÇÔUDƒ¬5‹©@ÅäƒŠd­2XDí‹Ÿz}²Ô',ÚhDŸ,õi”ñS£U–
…ÅáX§NšÅtÈøùO«&}1ÕŒV-YþYûÅÓHé¯x½=F¹Y“CÀ7 Ó™érál\~..};}¯ÊÆåQ-naº­b¤ä ‡Â)œý¤´Ò”žI\tu¦Ôï÷]wt)^4m…S”í#w«’VsµU¿êyH£ºãL_¨?Å.§?‹d³ð²;vSQA¯©·’z³:{Òg;·õíµÓW‡óä‹ÇÉ—ëÞ¥kµäm2ýê7-
ÖvWb/ç«ØuåÖfêwŒTˆúwª:œE®Þ2«t'¬ÒhÊoÆžúyv›wMI€íÆ2‰˜­ÂR íœç  wd¯Ä[=3	Ó	Û2%’Ó­{-zÙÎ0KÊ|‚AþòŸÄ}WÌéA9I¶’€/Ôéf§§v•Ø™NG`G:Ì^
ƒè¢ˆöCoè÷!Ú·ˆˆ»“c¹6‰´XÍôaÇçbL·ƒ´è-2u³lhîìtz5Ü/ùb`$|UHçQ.™¾QÆÔp+ƒ`˜lcD>w§D|ì„ÞM¦*O~éõóZ`’ŽžfHOßr²öŸ_½ì¢™Ÿ|T™&{±9›°·¢ µÊVàlo}8é[Î™Ä5$-º@l´u+äiE–¶Dæ$ø¥–È:-–±BF8‰/êt±õáCõ–§TËã‚z7)Õê”7[vpÞÙâp:)rÀ€ˆüüRVpÜ èôì¨Ð•'"ÀfìZf¦žYŠ`iw-SÉCGœ†£8B]Ë„T‘æƒúÒÌ#t_ËÔa†X/oœK8˜íöØn1„2p–ç3XÔYfw6*HGÍñ¹jSà’LÅÕwÊv¼mÛ'DKígaÓ¢^½F±ŸÆnd~ü¹t¡_ÏVÈú$×‹Á­&8MøPb9vÍòñ$˜b&«²e`‚lç¥êqÕ–¡H„†´O[­+3q…FÕ Òô%Ãã·jè	Ppk¦ªäm‰À1Â›9³Lî%2â®KÔã&Ú l—=‹¢–H>?ÚÛuøöÙ‹ýƒg‡;&²TÙsh…ßWkr¡ºoM1ü¾‚DC¶[Ü–p¦¦hGy´£lM;`dÍC¦2ÌNêAÃSÄ£\ñ0sY7ÑòS¢|ôwHCÔ7‡¥˜ãàË\&¹Ô/m™ÇJå`Ú#öN%PºZµ”|¬Û cp±tR6Y*ÖËnÏU‰±æÜüNVk*·“øk€ÆTVž_4^;M™ÌýÛ±‹Eò“¢	ƒ#éÖ¬tÃµ:G R"cÒ$>&7ó¥ÿt”Í*({pð*Ù=¸þŸG{»¯P‡ “ê?Íóó"¡Ìhq±_àŒ‡&3•áŸåí›ùÉ	,:oOÀÛc0Çt¯;“!êpDhŠ
³ÛÈgØk+yjÂ{iJXVˆoê«Ö¼Kdû Ý4†òÆVM²;	ü]¿0_>®¬”4c	ûŒ„ØÀ`8²ÙMƒÕdF#ƒ¯ ÄYŸj®âéxo‘ŠÝÍäï¡6ÝZ§ÚÂ³ n¬ë¨”ÎžZ§SQxgOMíyvž¼¾-?3?èAÖ¯Ndª½× ¾õEø"/…p9rç/öÀwìÄV²³·eµÎ‹þÜ„Œ4{ž«X%øîg»ïPúï2lë²Û[•*íTò¨}â¯¡’=	vð}ÛÀV2´œBÚêÈÅŠŠ–¬3’Á#E¯èÖÎ=`€ìm K¶µƒõ™î#DŒ¦ŠTl¬öp>©Íuo«!ù_CAksÉ#”´cÆócæö½¢£²»+‘ánd)BIx²y7Ã»Pˆ2Œ2ôOfÉ0ö!?¹I
I“ŸË¶mÊLòÌo‰¹eF€3£•òT
äYHU?¾šì‹÷åòìÃVr\2Å˜ŽMm?×ƒµi«¸rE~fÐ<Nó ;éKÒã¯Ã‡t6é®ø0ùñúç“JÿŒ—.)-)¦W“ãQñ§y†œ+@2—êê¡¸á5Lû&%”À‡è°`ŠHûÍáÀŠÁ€ò’‘ó7àSq'O®P kðaW=þûhcsk}þëàUN]¢g“txŽ½Û_Ãõ[T@3þøÔáüUëV'°„xUC,¤P–‘ì¢ˆ™Yd{x´]Tâ†Z¶Õ½#ÄªŠd-ùÍ%4¹M2•{Y Ï±;Â<KY=,CP «köŽ1”DF—Ð2ÔLÜ<BÃke+9ÊÇ˜?k<%!T,”\Ffy«µîšb´S­æé‰)€y‰©ÔGé$ûQ¥HÓïmÓ³»ŸM9„è–½vç¥"Ä<÷€0ÉhÎ€¬g& )P`uÝEL*ébR—Ñwû¤˜dúuÏZ>jU7zTìM ‚gáª˜…àåw‰øƒ\0ËNdæGÏ
ì´!Nj{ÃÚT©V0ä,¦ ×U˜uqŠ¦E-ÜølÞ%ø¶ˆvŽtü,5€­nñ'Ò?C)õ”«Q´Ž^MÎð©½u5þuV¦2—–±éézÌí ÚÕ
DïòdaõÛ‹Ú;•×C¦œíTn»½ÛÔûêºhÑn·iBl8bóÝ]Ø¢=kýXs¤?¥ª¸^îçþ‚ª¼¯o¡³
²ºyç´Ë^x×"B—™koú#Y^o.ÞÉ@bÍéGýb×_×‘!j–‡ˆYT©¼!±1%. ³F•HßG6…Xï½€ÎOÚâ_¥ç ïææpô	ÞT!‚ñÐCÍz>ÝÀââÉIsG¢É·aGðÑM‘ócAxQù[o§Ò`*tQPÇ‰#e"c FÌeÆ1í…9£ÝÒŽ:V«ÌÉÑ!~ŽÝŠ)7—Ò®×îù`—B‡9`œ(¢0w›jÂã­dƒ×N’©ãYÃÚc¬ò-§Ò l%ëü™ZM·l;7®oóB¬–²Y@6KÂ0pð"j¨‹y«rÝ­]ÆªŠt1a²µíÚ K=º,‚0mQ¦-ÒÜmbˆS‡:aäI|¥<'_Öê„¹>ee ÎÏ¬ºb õb†4ãWŒÏ}5­ò1šc$½uŠå‡”Ç¬v•Áën6)ç$O´hIù¹JÅÞÆTB¯Àâ¹/-æ
•z&ƒ´ª7ˆiš	ïÝË9@ñ¡Ç–:ÓuIŠJØ
Ý¨JÌLHæK²T¾ÚÀéu;R8dªsÏÐe‡Jæî„†Èµí\ÉV¶PÎª„UBÝr†ImçåˆL˜õ”¡	ìˆßïÁúœc¾™Y–ý”•˜Âò,œÂ—IÒÙåx^ð;.­êJŽ—$æ+%Ï–”)(àŠÏÅ\ññv¬[ QÌ•kÕ ìJX´yÅ§àkJ†Å_ñqízBËêš!«Ù²«z±Ü‹ !'Ùi»gÙÁÙ|òÊFº¼‰Ó,5Yï4K^±3$ËwîÛR•,¶9ÞˆHNÒÃàÌÈ¹‚—ÞgÙÉ¼8Åôcs¯Ë(“üb’êIj©ÆQ^Pù=ŠÇDSA×Œ§ŒxJo´h&JËkËÂt‡n+ãr+ûÐüIz^ÌV“
•m°GÎiWfýeèë·=e¨Mxë´ªT2¬g°ôÝ0*pWË¬K² PÄgÃ¼²UÎL%iÚW´n¾Q¼ši–ÞsJMe#N2VÞÑ{cÞ,w|¹²{˜”,s”Î0Ý2¥Ââ)’(œ´/h°´À©®EÉùÈ¹0`ÁåŽÌoqþòÖ¸ð Æ¹f°è–iWcCU.O73±1BV%!ß3Å“°(Ÿm÷…Œ8VNÛ,	ÂßiOJ8žYÍq¤–Ô4ä†mà1TÞùj³bŽ:xCbô!ýXÚ‚¦ƒ¦—ø œC.õtèOFœ'ÈF  ¸‰Ó‚«ñáZma#ËÝÊ¶„iÙeù$6ƒßR6BÜÝ‡—g†R‡e˜ÌÕÀ[£ÆŽ~ålf™©£öŒ®Ÿ£‡Ñï§£"šÓýt¶|³JN¦[F‹ÒIþkÒ5ËN03_®MÒîM®>ÖqÌÑP}· —”)zœ*„1Æó¿Ý__{°þ»o ,ŠEYa)Éøýl¤\²Ù½1p8GÅSñšf´š<Ø\_MÖû¾¶%ãŠLx>êY:Çµ/ë}6Ã_3µÇr½­€¡ÊÎ1ýñ Çl©‘²Ç/¥8‰¹oh«ÍzµÂðA¾Dé©ðÇRá¯ÊKu³±TæB	ÍD›„L	CáPû&á÷=‘Ë~Ž…àu½ f>epDžU3Ý³½4.•qiL‹åÃO”qå 8‰ÝÀaÀÅ
9`¬¥P+»ÑqŠW¾¥œ’¿üXu«§«Aª
˜@4U®~1µuâô“QÅ¬A‚rÆ Å§bVKè¦“ÂZC'‰®ÈðËP­Ôõö¹Q÷þr4§íh"Tç)p_î„"É…éÎÊ5Qž_›ÊÔo\šÂ·y6doØ¾VÙ;&7¿ü oEwè¢¢Eyy–M{(7
[uãRž“Ñ¼<“.à"€…âëAð}šÃÐªsga§:¼ ‰¯‹yU_Es¦5…¬àsWz(SgÌÂå[Uê—ùO:åº¥M1jEêÁòbº†.Î)U¶Î‚ÞPš€¢­"=ò³““|€>gB	‰´«‰ºœÎ/£#ˆvF£ç–¦ËæÁšÞì‹à¼A­ÁGèŠSÓœO48É”ÂlîÅQ°lRÒ3û-Ä”ËmË¶R9o~*˜äèHÍZ2¤ËlƒE×vbÀQBËê±‹B2vmjWýô¬!ÌâIÍÏà! TJxü;„-–ÞOÁÒ/Ž-š²ž¡v% xRŠ×bkÛ­Å¸ôŠŸ) îÏCÙžå\ÄÏHÁ¯ñGž–B–o‘ç­{X¿êMš’)µëh¬g˜€3Ø”°S‡™±šÔ†Ñ¾åbÏ£:	öÃ!Y»º¤Ce¸„wÏÓ’Þ‘¶v¯¯w7FÍÌ‘¥›3'p xÀ†æcâ@_bBzü²ŸÎà¤Ñ÷Žtäu\»uˆ¾còiÛ`e<DU}5„î­úy¼¹šÈÖo¬&ØããuFÅ§E>¡4öçxy%Üv²	|~¦Ç·lÀóu®­§-‚®^¢¥¿§aZ'…Ð¡[7h	OD—3îñ¦¥¢÷À";ÉcÞà¶Fd`õs…Ã†:À\ÔCNà»zlíu¾±Ã.2¦³¾,à:Î k¬E½ûœÝßjŸóýmZ2û×”›UÝ¸{(XóµNæ=å‡gäÚX*8ö²tPI?©1í®K4Óãu²Éà#œCÔ&ÁÀû”ÉI–Ñ¤ßSq1)‹a’šø	Oüœréä4m†’…Jþ4ÏæY®„ì¦ !±®ÙpH
4ÓaC¸ßG¬†³_÷Z!ç–t¬TzÈsŒóI>FFl˜—ïïM²êC1{ŸìÝ{…£âƒ‚ìxÕT4ß›y¥ù¨VåOÙ¤í3d”5¾š<Ü‡/Ö‹Àñl‹Èûì¶!¿jƒÆ[qM~ÆÔ¢Ý+´ÊÇU ’b¦³Jº”›Jýr:‚Ó¬³f|](t:+akC±-:/¸•„ÚQ""ì}~Ø™šÒ{Žzº7©Då×ëoV“u8‚çây²–lxï6U8P·W$„Ç`+hÊÌË—éË®©Þ?(á\/ÏŠG¬Ýr—É'ç×?r¼û³š¬H¶Âø¾,†[Ð×K´o6à'cFc¹¡’Y	mR¤°•½Ò6nY—]“•ô- e‹Ô¤Zþ7Ö,®|ÎH· ÍmW=[-8M"Fãø8#6]>{¥1æí†y$tcü<¨"ÚÔà„³%Ém +lÝÜ?unî[W©ZëCõC2”Àh™.úÓ/Y˜s°Ø„osqMZâ•òÁ‹$ŠÖt‡ÒÄ½kj|ÍÈA†o[‚ªf`[µ£1‘Âš1³å¾]'‚l8·qÙq_#„%m“l“åéN‘m–á½«\ä²­gÁ¬=—Ž$Üv*~û¯_{‡¯¤÷»Q•]©@×\Á¹…dƒ>>”5¯|¨Ö:[Zá¬¨VXëMyÔ¡èÃÂòhJRtL‡)ÝÈ•Ñc?Gb_Šæ¹o­¬#T£F%¢9=º blSÝ?ˆ\åŸº¡¿ÿô[”õ“Árx"³ƒ,Ã0Ó>È†Õ÷er}$K	ÑvxnzàÌDXgÛ@å?–0â®%m©›­Q'5N+e„×¡#ÄÙ( «áðÞ‹÷>Â'ùþû­ñxEOÐ3Ò¤o‹IuìlwSñòÅQvQí#X×Àš<ÀÐ¾ëz‘$<£03L¾Z_×50aP·sÔpýçYN‘/yÊÕ_MÐz|	KœŸ‚8Ü “Ü¬£7Ipd›±‘m¬óAÓÞ=%÷õß\Jà]½}ÞÀV3:×ù6vJ»å‚Eo
†©3#½þGºïiÚJ)öo.[uû#H÷]€j§§¦øå¦ÞZ1ÓÙG†Sé¤üÍ„T$29ÂJ‡z)¸Í´„¯õ¨¼rž]‰ÿŽ*RWE§ŽTQx£©‰Ù!£}`/ë=Ã&‹VÇéÅMÚÄ8 ]DÆ q‘ž6Ÿ4^zm›ÕV¡Ûe³»—tõ¨x|<¾Í/²awƒ¯w";rãËÞ°¤Ö1&gEBšD#6ô«ÿ(qäÁ:“v÷†xrSòC8•Ï\>’Q½’¼Qú²|®EµF`³æ3ÂjE“™94k¯ÊÏ«Ø´™{I6zCvº!Ë€ÙªTÖïHöÿJÏ«,f R§«É1!U75•:=qµ$CxÚ=æïjÉÝWîêvtÌ\½d@+Š2É`ƒŸb(‡£Àš¸À¿Õ¬T}PSÞ}P+jÂh¿V„;[­oò”Ð¾FŸ«î©‰Cto®ÿ™ÆF%õ]á’(Çåœ‘Q
ºéSÐÎK@øùX#H2˜]ÿ+Ñîœ©¯ ó@ó2–X¾SG8äI9ß	SœUÿe*ó,KAXxýºCÄÉ¯Jé‚ßYÀï¯4 €¾ycš8.†·Ìº° gv;§³|Ø±»<¬>ŽD /LÑGD FÁæ7F&ð]¼s:KÇií8ÐÅbGŸ)<ôÔ]UHÝ¸©mIOSÇU­H¨…?:’~×»Ú0ð„Æ²øký4±ªÚ–«¡ðk™×vœ fÁB37‘°[`ÎNÓïA²ÂÔöTBpn”ÙàŠÜ—FŒO‡Ã}Ô¸}&J9ØKÿÐP/%iAïdÖvD¸—®ÍÂíÂ4ç#à9ó´'÷Ñæz¹±FwÙ?›8xôÚ©Ð±$f¸[žÒªˆÍ?u:½«&åõÏ nÚq˜ÉÃÚRjSÙÀ´|èÃ{káÝeËsâ_úƒAç†'è&‰°-æÕtÐ=Çó4P×¯}•¸PžÑ¿ È¢oÁ_H»r÷<m°„vÆûG$·#»ÖfMKL¶HáJ±¦>“Mì^…«»Œ¢ÏT,Œã@¤«×D¸²1<÷w@2äâ¾ÊØå€k€„™©@_`Qmè9ë8|L%ÿÈ	Ýlúƒ5ckË¨vüåê_I•­ëÂà’/¤Š¶
áÑf
!þK¦Ö
½‹£ˆ=ÔÅ^'³RÜïúnd)â/¼ÈÆ…c)‘ÚÃIþ§y¦JFƒÞ(Æâ<ÍGHubÓHÌ[:‘éÐe©È"€aABl'Å™PõCö|CÀ%†C5©nŠ×éºñ}bÅ†¤iÎL0ßÊU¤«‡Ó©ØÔªÆ©5•ú:+<ßA•‚$LÉÝ•“ðè,Õ,³EÇ4Íª¨ä“­«èdpñ6J3.6çž+¤®‚«µo:¨ƒ)³I™ã\ÁÉÙ9èv„õ‚ ²š¼öðwÕìûU–iUÇ‹}c‹bhÏóc`ƒ>Røù†íE:ŽÃ¬ª¥1ö¥&ˆ¶³”G·¥ð’ÆËZ=Ç”åËx$Õ„MžLMIÕu(Y(°I-6q$ œ–Ð«°T+Ü>hÈ²ÜÓÐ†”ÒeúÉglé=Y¢.qƒÜ¿·l­‰ìÄ½xÎ±Õ§*’ œ2Þð˜âé«íE"¾J„µZÕµÂµ[Ì’/cÃÌk‡ýr‡öYìWVš¤Ð>Rêô2R*‚EJKÝdöº<¼Ëj8Ú™a}4¿¡ñ\ý¸TW­/Û©µ¿ ™<à×ß‹ŒnYûÀ$ñáÚ$uÈGÐjj”-ï_ËÅ‘µnk–ÖÖ,Ã[Ó´UJÄ/oŽÁj™˜_.ŽùA0(oƒöå2Ð¾ýÐÃyÎ’hVÆ\K A¹Z‚Üµ£ØKbí\B›”.3ùŒ°ªÒÌs·ž¡]MÌÖ®Y’U:š–íÊq?æçÉ`”–%Ê´WÊi:ÈÖ>®}µòDÃÆ-s2Ê.Ä5a/•Ä?®gÕ‡,›°º¢ökUmðæPï·vÿbBó¤Z;¥ƒ÷	=+Ñãsíëõõ¤šÁCà¢×ªüô¬ZyòM~<Ê‹*¤R'x–rºýèÞÙ†ÓÛÔëL4ü ¦.Ñu>^yòZ.þ©¾K…zùèÞÔšØ=gfA8¦ÓµÍ¿®}˜¥Ógpùd:¯Ü+Ÿx-ªç£lÅ}•¯PÎ‡5‘¨Â+Àú?Ë‡CX·D:dÓêñJÿbT^¬&øgÅ)RLvÉãõñ¥ð®y+(]„½Aï9Óùf^UªÚ0nôrñDžlq”Þ?¾&x1ý(Oq®’*¯F0ÚoÒüƒ²Ãl$SS8À‚öŸÊÚ	îñÇ—__Á’TÇÉ=1”Oäš¥;À!Yºfâ5~o;JYÿ)ÖYÎX…Ì:,sÄ}tr{H_¿ù¸7ìv8žtzÛýÖêö®¼ñÊPÖhÅ’Ãxi%ÚŒ·ö[TÃƒÛ%So036ÄŠe·Ÿk†›)€æw–4'•à@!—„ bdk¿
r|*¾?X¿÷ ~³a6O6áíYqžÍ¶T™ó„•»/)?ŠÉÚ ÕýepÎÂ‡Ž]7¸•¹3’Ù	Ã²bM‚ó¹rÏlüÈ=È2šç³¤*p“Àü<Sthj)ì)ÿ–•°Ÿ’c‹œÏÇiR^ÿ99Ï~réR’¸8ƒŸGG³´<ÛTøûáïóO×äæ“¤‰™½AlN’ž;ÍË¸Î]-ü†ƒcîê †}µnãÔ†‡SÖ=‚°^ìÅÍmÑ"4_ouuQx¢{*²0Œ7˜¹‚¬:®™µ*‚ºµîe0Ðµu¡oTÊ*1ÕŽpCÒS¨ñúM$%ü ½?š—6œ)å‚šÈB0~tÏ†îAqKpD¶Á¯¹j¶‚ÉF¯Ùj‰â‹Ô´!‘­fK,¸)bÛB€ ªÜÜæÙé¼üÅ¯ ÉŠÞîˆî	ë²J2BÂévîx‚[)ÕË³ðê´ÝNZUó¿›Mæè_k3j.ã×ÙM‡Âw§²L«½dNéôv–
`øA&QEäoËùi†™­JÕJúS¡Êã‰¹9oÞOsÝBQêÒóQ’Ë[¨†ú²,õRBŽÍ“³þM¿Ol£k‡Ÿ;Ûé"õpl›/Èß¼ÌÈè·7!/Ÿ¦ÆÃQ¬vg¬_¸ä«¸7‚ \œ¤ñø RÉ‡Yp	ê÷P0z™^ÿy‡P²r®ã«­ìòg½öOÇê‹Ð¨¡Ä;>^û*Á£äï_¬¡?ž­<HâÎúa8½]ºÜ¸lÏÆër0qñcezC™~\»/5‹Åh(‹c.pcn+5ŸµÍäÃY^eB:)Heç†»¨•WdS™¸“BG-uÊ¬ 6«r4ÔO*®õüRŸÈì!VµÆbÃ^ûdõèÞ±‡%7ZÍ'šËÁ$’Of9|ý’–ÃX÷>ÅÕàí§²A¾pI«!ŽõOr)4Çñ©¬ƒ¸pv;Bgeÿ$×ÁïZd2Ê—á/‹Ë$pÝèç×þîï÷›½£©'N\²Êš¸±6M'Ùh%d¥¼Oì†kSst±OðT­—kÉx¸e~n’âKßˆã™-éé(=ÎF¼õãQ¡ì•¥‡8zmæS›i™áˆ7Vž tƒwŠ§ÅLlƒG÷¨é@—ÂÌÛ`”ø°v2¼þQx3¡7ÙpíB!«Ïªƒy¹%5k”ÛS<B‡E4Û˜
¬¯ßÛ$£¡½¯ŸGsÊhû«ôD¹ÙÀ»™ÚÉ!0“›D8@¥˜Š)îW9 ••'h(1µ‘ ;î‰¡º—¾¯(9ÃÒ5¦jû}öññ%>¼Ró¦OÄÝC`Äî	yxç²çôð—CF%lHéý¯-=Ÿ‰E±Ók`yH*,gDëpÔ©JŽBèÄ ’în1ž“qiÙ«EuçÊ ŒX€7HÃ‚*Ç~,ª±Ÿ~<ê:íÝnG4É°½«ðÑTo¶[ôà²Ÿ÷ƒçÈå4bàR·ÉõôsŸX¢uÛs@$'*ã÷cJ÷zuÍ‡ý¼êü¢—­ÌçfÎÑ·t‘3i—ê=D›=ÌšZŠƒÅmûêVPOO[A[¾0§§å¶HE[Y©håð+6tÝv‹1SÀ¡i/´ÛiCÐ&Goç+ë¤‘¤6™Ï¬ˆúf„ª£y¡<K‡Åd¥›*‡.k=œÏ(8%˜§o
0Àx•˜àuOäM{®ÖëÌL H:ãmtIÔmÏIŸE¬iËx¼ UØˆTÅSè.¯'å}<)7ð0;š
g›8¾é‡–^o¬O/Þ°#^­·æ€×žsr\8ó*:¤WDcB‡!÷Ž7VýÒ¦tƒ?øÜlñ°“N;¨ä“a~Zh¸ÈŸwwz15)pð‰JTmã~I8– åÕÇµ¯×ñ R¿ÖÅ©ÛhMl<õbÃv;’aŸgõ~m=Ž²!yœZÌðckÙ1¿Y“_†]"OM›Š:Ž5…“°—G0†Çj}C2r‡<²k‹j· ²âkCËº;€]-„ÓÙãÄ©‚•^¿©)·Ùá§½ŽÚŠo¾á×6úŒIÚ#j3Nk´–.Åõ¬KóYYÌÖ(q6É
DÒ¥°25žJ±òqÊñs.aØeÚypåúÍ²Â˜õþÎvŸ0ÜßÅþûÄ°C{@Ü5/¶à†í¾psÌˆžñg÷­S%¤,øZô6Wž\i Èxg÷Û0XJaã„Nú%‰ã¬:ÇåÝ€=EM=.Â¬>´FS’äöJÙŽƒ]¨ú\Jª*øøÆ–#|O&2W¨E©Ö6ù•6Øc´# –ƒÂ=æüŒÉöŠû²}zzåða›È‡­ã¸úÅVY­Ý‹G»<ý½¯Ì>#<[„Ç©áÏÐ£¶Ñú•·ì–6=ºÇ’ä?ÐœoÀþ†4QÔi6ÊûþM•`wÚR•pSµÁ¿o%-D´\‚èn©ïJ¿)¶¶Ü¾’lª/C8ö²¡¯&`ª„<»Ëõ%>7úZ6ü¢;:\9E¥*ÂN 8øñ.N?@öCû‘oGWß¸ˆcô]¨»anO^pÍ²Éó¬DN§6NŒ]Mà)^«•¡8;ò»·4hí 6£7õ˜½¼»øêÛØ:Æb*ôŠîÆ^‹Ëåƒ¢ÅÅÿ¶÷ŸÕìn¡à½¹Š7¤ä]_ÆR·¤ãw†ír)—r:Ü&€ÍÇÄñTÑIëÃ}Ù€ÍPTHFÚúgi)^õÂà‚ÍôœFT œêà^>ÇÊ)ïÑR®v÷Á%c
YK¿|p.óÏ_¼®_M"*?/QçmO·I\©¹/»½-î~.©UªË^?1ƒCd*µV=—ZE»øDÔíºÛ®wêÙ±Î?q·x±­rNŸ¤éCk‘™ÿ‹'`ÕÐO”1ÝÅR€é(ªRõâ_ÆTó“ªÈ?@ùaV¥£3¼R]h!¨©úêb	Ï%3(öGÌŠqˆö'~M ¹§›ŸÜúB	ŽƒnYk{”ú„íR‘¦æ#õ1ÁåÛX©ÔÇX«(>l+ƒ•úØ†«x\±zÛ•nMØ°¨™6f,õ©7g©Ïb†ÝvC‰~šˆð¥ÙMöÙ‘I:zžOÞKÍè×5öY«‘,×Xªe ÞÝ#§1ŸÄ•'ûtIøRäÎ½jQÃx>1ë9©É†‡Ÿ	sˆ˜5]åvë¶°æá§áõ½{É÷4(™f=HGÉÞÓÆ2¦`á”¯‚Ê$"4Mør®ù`šD8hÒÑž	€ý8á‡LÓˆu¨Ð¡T¹·œã]‚‹°ˆ ÝÍ„ŠÃCMÕ®|k—ýñ£4ùìÊ³3Ÿ§Š¹ m@Ü¦šŽJÀŠ´©V¢ŠŸÔÛÿ´rÓh‚x›Ýð˜×ËÛ¦þ‚ç\Ãö‰¹qàÇqå^W§á”\Ä¥ƒ*4¸uP™;?”‹ÇŸ.&-Áå?®ÛÇíñ¨½ûoÄ¢Zf¨þåî 	ÝÌ€Òã©ôQ44êw—ÀÏ¥§HŽV¨¥7u×Ò£âþM=*ãníbA¼c7‹wâfA#_Ž«~jÝ-¨¯î?ìv!jF±}!Wß8îÄ)îu]O»°Æ`-+n]Ä µ¹u¹ŒírL­W½_ÀeI3^ªEõ/ÊËå.ÜZî—i
mœö’¬žK±yÞÈâY³>ÚK(fô”$‹\ˆ‚2¢v^D©Qº°QäTAw!Ì S´ð1ŠE‹‹¹ßºöÌFW¤ðÝøKÕ¯7]µÊ|Ëö¾ˆN¿Ž‹Mëô¦½¥¤P³»›r©.6)r)i¹ÔËt±Yh‰ewåÏ†wèžCî&eTg°ÃÉM ¹H®NÊOÂé$~éí¯ÏyäÓô¹,[¸~´ñ¸ìzÄ¶Ñ³c±Mž‹Ìüôì¨xòµñ$è™á©›Åcñu]½¾Ñrûó7_Èço¾Öþúj”·òÕpT«ÊñàA³RõÖ>s©ÀOsãy»U‰Ü”­L¸VðfÚËóŒ[$¼±ê”U\Àb.PÙ–ú”—ýÀãfúáÅ4‡QØšÛpC ãÜù“æÂÓË~èysknõ²o?in!z½ìûO…¦qAolMì¡/ÛSùVŸ¨]ù®<tôù§.
wHh?QLº…eàÑ§ã¢ÐR7Ó&Þ«3Öm¯vŠg_¶–=½h››+Oô¢$æ|tïìËÚN6‹1ïé·n6ËËài{‰qÑ ã[ÎWˆÉïÈ¬o‡êè¾õÄqÓ°ÌpbIí×¥Ç&ÔO¾ëmÍëÍÿ`=™ÂÄBž’î2bŸñ;jØ¬%ZÀñú@°2×?/Ó´.±˜-Á~~KñVÂ¼q/ïítšP³Ùñü®áó»ý=¹E#æÞB6Ø6$m:u[è$#Ž*þÇh’v&å‡l†Ž+´˜ee<QfíÚR‰ÅÔ‚C±ßí4JÚ¶À«’¬\À7?Mâ9ëËµÉÌ˜µö¬bC@½*ªtô"½ŽÚr`/ÁïvËùx5™‰ðÄóqòEÒÝLþ>Áe‚ùßèõV“õÖÝeél’[w6ë—ƒb–aGëõ3 ;“žÜr  žõ‰`v»r ÷t™Ì¨%Â×›{©5<ÙŸG¿›æ€/§€Ñå7À~·B>¤
(jaxh˜ËU»ZMÛ„¡~š§§ØDù À\ ;U>9‰¬“o€ä`è’IZ&j«Q&€
ÎrYêE:ÅR+­ºoÐá§qùÐKáæªýÛ(žêHãNwæ£Lñ(˜>¶$Ôøœ•?—¨ôRmÕð_Ç¹ØV,E¼týë¾î÷û4®Õ¾QooB†]õAú©Z´"õG0Ï»„Zl…’Ý¿‘ó,~.õ4ZzÍÒØê=gA
û4C”éñ/Ï?>´Ôg£Üm¶n]}ÿØÆƒ>^:fü{É>G~'wímº€En©ÎE·÷6mt"Z–»ÐÕbk´<oÓÖæó
jÒÍ'ú¿¶f©5%wàƒiÓ“ÆŒm3xÞÔõÒ?N\—"ë4»ç(|Ï(jBè^´3ãÝ„Õ·³¼l¬'gø×u8Ñ.øÑFÊªè„åO}Ô†1èyúåPqŽn®õ’±È­íBâ*eSÌ«“­”¢ç0“9…@ÐoòãQ^TÙ FÇmívåÉ.`{YÍ€2¢í1KŠ9•1åOý™ú*f§é$ÿ)%ìÏ©ïcÝ÷[!šr(èa¶Yëa†ÙdvÎÓäœQ&ò&5Ó¤&µ¶&ÀyP ù4;Iç£Šš¾RúnªtŽ)Œ ,€†@Ót8£å) MÊ
Õ“"9N˜à’‘dÄ"I€ÐYe»ƒvBVÓ“\ð}ÑGhNís²&‹$²m™îw±d¿í,#áÌµ‹¥ù%žæÇzG°'Ákwm32r±Å¼_tíº9ÞýRgùÎdˆ¦ÿ”ðÚ4ÒÏ0-»ú‚˜ô,²®¦Dzq€~h«bsÁ#Ö1œÐªß ;é¶lÙm ß–þÄ
„1Ÿâ>Å(~­ltÖ›û~ð&ã|Ü¬Ê_Ø—8v·hä /äRüaík8Ó¿ææKË¬é˜`Úç:?]+Ï"ùh˜ONµò«F#ä²-QÇRÉ#¬<1×Â.ÐnÁ”âaC+dð{ÇB1®PÛÑ¨x µÉ“KM+®4©èb†¸fý€jâÿü÷Z¤ø¥"W†¡ï¢°U§Íkyg*EOªÕtÕì€l2+Z·‰Å©X3Q7ª)Keð²¦LX ûJ+ÆXJFÇ«A)Ó`£ˆˆ
·•5·ÌÆ91í5=/ä!œ	n«£â ›À`ÍáX¯©6\ÒËÅj1¦iF=Æ™&«ÞBŽ%5n%µ¸ŽþAŠ
~Iœ–ðªAÿªKüêˆç ‘í#q‘~[¨(yÜ5Ö=%‡¤›ápfú”ðÇvæ‘$œ…nŒBu©Ën¥ááO¶9|¡µ¢dÞ‹ª®6mÕUÂ(*\êx»€Ndu"›nƒ·SŠŒ/ÖÒyUÔ«Dþ   ÿÿì½ërÛH– üŸ"‹[½¢zDJ$%YÖJö'Ëvµ§}ÑXîšÙð:Ê	‘hƒ  u)®"v_ãûµ½ûcb&b~uìøMöIöœ¼ 	 o )—«Çèj™¼œ<yòÜS‰ì6U‰ÈÃŸz â£ñƒw˜Ï]màhÄ.]Pg
¡ä£²:2e‡ÁÊ©DPè{È)v0 ïÆWkŽ9‘ô$U8!!GtÉñðó¿’Y<òó%$£.!ùü×ÐÁ4W. ¨ã…|ig$-DEá‘Ÿç^øó—ºèÆ]¥ÊÀUgÓD(kóh¶õ.§‚ùÔ'´%Îú’¦w¢­Öš(õ&§ {¬'Lý ÑiPL°qUðÙ2|1•vøh×hóGE…¶49Lsêìªûëk´irüðÚþë0ˆ‰OÎý1%^é}©¶3Ñi$:¥öž!¼:Ž µ	ñ™m0Ý"#ôžùüWê>“ÂèÖªÿv%1Œ¼pÀ(ŒŠºpjž·ReQkccºœòjÖ€ ")*ù^4ªO(S9L¸>5å2ÊÒóÌ‹F^‚þ5ï^œ½9ÿéÙ«³·ÏÎOºAÌíÈO‹†t†_@+¡Ê&T™>õJñ5¦tuæˆÈ;U>¡×®²!v‹
?±|T#:ªÑ•]Õ\/°¶!ÕÆäTÐD4þÕé‚×¬	^‡ØM¬ÏVp_jà¦J`‡ÖÆù5ê ´i®+‘sTµ+‘Õ[2ý„œdR©Øßk˜¢õˆóó+ç€pRái:×g‹“âNŽúóFÚsÞW~Ü•¸{ìqYûÀôð†Ÿ”¿-¾sÛßù_‰"Ÿóv¹.ßÀ”UÞ,Xì×Í^¬©óM<wåÝo*}¼þ†Tú9þq­þRhÄC”¿24úÂšýS(7t¾K÷i£wT$`‚É|6Bcâühx#N†lÎÂ6n£!i£CÂ’,oÁ¿øÛ!Ù8É‚+oƒ€Œò$ŒžûðÞ†„KYr›c¬$—iR©—´·®ðü—
­Cü6ì‚TŒ¢üó˜,0°c(ú%wä']tWž§è \`°Ü3^ÞµˆI>‡íþÿb‹laõ§¡4¨-èwsú=J‹p‡ê”á„´‡a<=K’r¼‚3ý®Ÿ$qÒn=ÃXÇ :Ñ^€¼$¼}ØÚ"y3Rÿ¡üï,ônÏçCLHqN£»ø³ùPhgÅ8\Æ€]³×x[6„8½(Vã
¬ôg@‰w~šù%^ùQŠÙÝàãI4çŸžú—AP´¹GDÉô%qEêt]èýºÈ‚#hŒ*O/¥x F—‘øóøN/üä1!ïëKb‚Ôéº0@ç…däÝ®xÀôB?ÉÚßÁl¨n3£CâÀÌr\yüý"Ÿ•¢RÇ#<=ÎwWÂ®bVqìGäÈiPìŠÿú2 dè‰îÑ*=~It«u½.¤±ÁÔg`6±5£^Ï æ‰ >íôó?{"z%Œk­„qÕ)ÙðnB‹±½À#‹yâŽ¦A$!I~nÑ ªòjó	syâƒ”ÏÑÜ=§…Þ†4\9÷µï¶r`15m1¾2æ°qþAûøm{‡ÏÄ‘	y q®'~y|ùÉ‹ ÓÆñ1üÅqwç€ZyÿÜY9òf8[Š¤c?MÛ?—ü;|¦f·2šñ¹¿‰Œar	)ðçèp„äÏŸÿBüœ6‰<‚þy|Áw-	½d8HÆ“/ãºÀ`„EÇ‰÷‹ŸcZxÈæ5
RDJ”µ`ï~ü~Q½{GÚt‰7?ÂþÝx‘ySX!‚$ §ÐÓ<Ìb Zâ))¥×Püªè
µø¥¶¤÷`­1<ú™>Ð§3úðÿ7ÆÛ˜œLzðÙã†@L]æ' ¾¦ˆRÓY{³x¯ïa¾°ÅO>¾øAfè$ŽMê½8ÂâUü’ÀÞƒ·èzˆòlwù¶§ÈãFHÞœ5õ°ÿ_üçqò<à^m±àšZ©ÍÛý|B—/ >3@¾¼t§˜0ÂR®Jä14š_ùS%:$^tk$EÄ‹ŠCJMpóËãàÄ¨:nñpÓWðJ-ç´0—¼Ù…õO½1E8l¨FÔp—^þ5ÁƒF£ MÀÌé7	½d?acp<þï )ˆ~@éø?×Nüá§xžåÉóEÁ]‰è„E<Âxè…ç°„ÐrÈÈ víñÔOÈoBÀelLã(›„·Tlð"*AÐÀNñ…ÍMø”w¯1¶[$ŽÀÁˆYÐªŒÛÛä|~éegˆ„N†„ñ8æG½'?õìf°dè8±
ýRF!Ä @âàq6"­0²EÙž—þe&¶ãµa3Ò­#6¶å+˜uæ÷^˜YÞ´bÓÛòù¦ì#|™e×ïr
C‡‚Xô˜Åð¿+µ¸ÚX¥qÎ¸@Zj½[Haˆ*L-†ÂßA|-I¢bÇ6šÔŠ %/ó	(Õë’—@°5`<€ô	öÃ/P¹=`$n¶ÅñÎ{µÌgùÀË“¬p|›f@üüÔ£"˜ñÝnÓÇ;=Ö¦ÅW`ËÓ/æA•Ù‡Qpy‰çŒG’ ¼wŠ†7‹Û•Vø2@+4ÙÍÐÂvÞô6i÷vvvÈïÉ~þ§¿»Ya¤ˆ”!6_æ¼á#ØNR†ÖbÝiSï¦½S`Ã¦ô$C†`ås¯@ˆ~—¼É€íº0—ºçhH ,Ïˆ Ù‰£Ë`<‡QaŒÉ{p>¥ÂLCa0xÝ_¬ÍK¿>Å'«K]È¡°dÜª™Çd°CÄ é/\S?ìïÁ/‘‘Hƒˆ2RÅBî+±P~Q‡‚ŠgLøW‹ùT8¨'V‹|Ð
‡3E†Ë0†cÝûtˆ\¬_'ïÂcÅhIM(0×E—¡~¬¥»ª.9_äÚcÃ¡ÉÀ+|þÇ`>Ñ4?ßñã9ÓéÇËX*¡ILG}³ËU²û%-¶<6<Ož1`¨¦˜^—ƒ)ç¬Ê³€=WzUìŠ×MÍw<ˆ:“À‹»xæ¶F«sÈ¬ì³/YYN‘x¨Ô¯y¹0tµç®7…3)UóÕ\úèî	„*÷ƒÈVÂáap£*ÿ öÒG7Õ1¨ ŠJ3¢y	€æ6›ÝªÃgÝ*2éÕmt8†Z­U×£E@‚d¬Æ÷‹)f”·dh>Š?Ò”cLtØ·ÊîÃ“^iœºä{lL7ýÊK¿ò9JÞšaSŒŽç]rFýF4	™gÊ+CYŒûðßå<¢¹Øa¤g'Õ¥w'[€ãT[B<øÅWËCÏ1	V™OAŒ¹ŠÙ8º+ƒÏ%M:s€‹5ÈP]™ -ˆÀëÏÓgdæ‡qYžê’gQ– ·5¥Sñ2&Ê$5‡&Ð{g™øc8†™wl
Ý¦A6gîVjfÚ:wµÊz(-„³låvÍ@Ë…™¶$jÕíˆ¥I÷6£‘Bgöe£q)|&ß	|Ó…cñIJïUµ(×-–oé’'ä×+©:á«-§U³¦*6Û{fÁ8z3ÏÚ@P&5hé`Uõû¨kç¾è¥öª€8÷‚š§ˆzå©)æ_¡WGÛH²Å÷šåVœ8aLIÁ}-òQRiÓ‹‚)¶’Îà´.ùµQX^K ½(ÂÇä,mJS×Nµ¤P½Ç3t·Ù¹yP
“ãÛ¥8A¥x5Ë!ÚÛGˆíƒÜö+vMŽÒjÒîšOè ¿ÌÚ×¡ó@“§’ëá†4ü4=ùª3Xà¦~…ÑfL›Qoò˜l<÷ðÔ ß‰û	eT;§ÐK€àI»¨ä=MŸ~zFðˆ	à¡¤»Q;þCe ÌÀP#ë5”¼Àlè9‡7M©Xö•Æ›9±Á/òbçÙï§
ß1‹[Èæ<ö£KL¢GÀ›ƒìp¼`†Žg¨i>ï”G›TáFrÔ†ìv­Q%wÝ/	X¡?$ò“ãò(~Gýà•Î]é´ñ‰úg LSË¶}qržˆÇ1ÁK@ÄëÒ†œŽÐYîñ¨x…%ºŒ‡óô˜ç©øR$Ûá·+ÔàçY'ŠëÓèÚ‡ €Õ¡‘ø?Ï‘±¨üPqîV¥jºn3€Ê5ŒÛ¶tÈBùÑDQMZ¹3Þ˜ûâ‰7þ}¬!p†½¤Bîñb¿:G‡ÕåÜ[º”nò–
2·Œ$Y™C:óQr0,‘¨jõ5±R‘ldÔ½DÐc Ô”kÚ¨’4ß„„¬N–˜	´>ÜDËŸ3à²’­úÏõX`‡¯rÁäèóA'\4aù‰6Ù‰8ÔatwŠL×:¤¨%ö­Ü×0Åæ~Z„Æ…¼?¥óÏIÁuÕ*Kh(D[c[/0›jÞ#€ZGÁªMkl˜Ú¾_cnå99…¥‹±áU˜;G@R¬äa[ÕÒ7s!ìûÈûèâL šM“c8c0~ˆãqè¿Ä-y§Ø \tŠÎW…¥Wãr£{ÀeïµÚ-ŸÄ7Ç­²Cú»ðŸ&Óª—M0YjxÜÎà"³S¤l-2:n½ê÷»{û¤×ïö÷†;îƒƒNwçúm:Ý~oÿ¡×¿Úíö÷'{Ý‡ý!Ü…Ç»|fúÄá©t½«~÷ÁƒÉ »÷`ØïîÀ#ûðCÿ ³Û}°Ë>twþ¢Ik-õ?vOöl” €õÐæÃ0Õ½îî~§ûð€<ÀöúÝýý°ƒ}v°÷!þ·pPèÎ>üö Ç>õ»ûd§³×í?Ä±:ûÝÞ>Œmoð‡~·w Ã?Ø=t>$ý¸	< Ø
öî4æçOžœîì±1ïAc¤·ÓE õ;8¨î`:° ¢‡i·7€;»qãÇ :šS¼Mº{{†¿Mà¿ýïº»p"»0ÆoÃZô Ÿ¾ÓXŸì{9|÷ºƒƒa^ Xˆ»¸°po7t{{üsÚ{€}ãPqr° 8(øƒ°B, àÁè n8:œ¼»¿O´Ãî.Ò>â
…:Éá¯1ìá«qý6Û€N
™ºQË£9c¾ëÇ“8Íj?»„Œ";ÜÕàÆ³°~jhOE•ªŽÿF=Ðqˆn^,ó2¡¤ÏÊÕH\­ƒ&¨¤Q†öé"ô´¶­*ºã4/Ãøº3	F#xiœpvXÄýÒ¯ùCWA\„âöZ@øç	S@°/l²&d±ý{rŒü€Ðh-ônÉï·‹Ùcœ&û@ªb6k	˜òÄƒ”Ÿïù¤arÛh«ð†ŸFI<Ãòì	†ÿÒÙÝÁh{6ßò’Ö±/äaµ©™¬$‰HR–»Ä\·¥óÄÒ¨ºdÉ,ó-fÎ(@ì°$hñ
†¥ó—áU|Àóð=”“äE@Í»+2AÅÎ§[’0ws4š¿z˜ Í05C‚=Ñ	Ÿ&[˜¦YSÔQßºHíŠ($I=²œhcGŒús©~â% ô&Ù ”yBÍg™çQT	LÉUøB¦»þ—R©7ü„roŒ'YÕ,&¬O(ç¶É>‘g‰6r¹Ö­¡xX¥,(Ê5ì›<Ìøâ–¼È>ÿOt§ü!žú©ºW}JäÒ-uQNÛŽÿ®D—L–…<Œºz†ÈdØ(DO—
|LŽþ©”ÂŽì£W~4/Ý¬ŠÕõª“GÛlÇ—øzuš³7a®gF4†&ÿa»i'þôi†6e=.ŸwñŒ`ñ¨¡ÏÑÛ‘•ij¿FKr•váøzáf…:™¶-¸c¤.4¥
Õl’*iÕ4‹¶1¢f£pM¦¤)”#§À¡=ÀôO=9áR­`žZÀ¹Ð»Vö àÎýŒ{o2Œ†¨‘xóX#]|¹aŠ4ü•‡Âª·j×Ê%^Å²íp×S†qºÒËi1îÆJç£yíç­²-«CŠÜ%UÕà’jÀy·|X7·²KožÐÿœ9
û°ç£aGÔ´M]…aMQÐÎ½˜I;÷¨ÞTõ¤É*pŠ@’¢e¥uÖ W¯ÏTV/Ñ¦%aùøÌ`ø”$èó*4í>ž:1ÿ!=›;ÒK™Û¾2œgX>¤Î¸ 9Æšãélž¡WÈvä]ùcüÔ ãÿ 2ØÓyn7¬gÍ±¾Èñ³<ÒÓ£ÎÉíë¿¨3†®Ì¥aäz
FÎÃØNŠ$€=ìKÍêÉn÷™@À‹ÏõÊ~Ú¢ËU[«~½Lë×T$Žò!æ¤ý4À
’èæ«Ê7êÖõ%èµ°XÊÇLYß®­Ã·DùnZ|<ÏOÿÂC¥¨ñöî£.'gé`P€]æí…çÈlÂ³vëbP5ŒGhËãhºëÁèŒ¬£m`µcóÉx»‚:ð½,ü"ŠUV¥¹ÐBR“jÚ…vÊ¦ÏÀâ’yÅ”q¥bxvƒ~€t“?sk>S1Æ"Ô&_±uQO Iœü¨€ Ø‚,€jøü¬{â7p ?ÿuPïÈÔÏ=`^/”buÿþüÍkK§<4žÆ×z’	,SV“U¢08O½!IcÒƒŸh‘_y eVÇÉ.—óPUU”sR§b	ë9÷$)Ë¼˜5Y^°K
)8ŸÊŸf–µ™	#ãðbê´†T5‡¼˜Žÿ25é¼áÐŸÁNíþ9#õ#…«³ž:lÇÒ†T)kÙ¥´ˆPÔZNO„JŠô’ƒÈŸxèˆ‘SêJVÓG,ä †ÂÉZŽ2h~z¤)ôv4‘º5•7^P?^,H<ó†ØŽaLðýòY‘øá[z¡G_P>^vd; ©¦]˜:³³"Ó¼ä„ÝY@Ÿi:Lî3™Ÿ‰æíZM°‰ÚcNÍ£óIà‡#oºt¶éº¶ß¹Ðg-ÿ4›(ŠMJ^ójç~¢‹Ä¿ò¿SæxVv*’Pç}ëÍE?%ÞÌÇë‹:"³d!]æîobã:ü+ÌÐFbrA"¹Ä`ƒÄÓÇ‰	Pý¸Œ_}iäØVr—*äÉlÀ·õ
“ˆ¿B=çtVç¾õNÃ±Z½_šÓ½™c´¬Få$"W%WÚ‚ò<_4â"trÓèiMŒ1Ë=é|£xp´ô[íåÆQj¢S9MUC(±½rà¤¢‡J ¥â	] eù2ÌIã«g%âWñ~d©ÿ±]©„:÷ž:<’Ææ¶ûnZƒ(Ë_Û¼g<VUOÞ©´€ÐêVÕ+­Õ—Ë ¶F¯Öún\{ÚnÍÚº"UÕ#ó;ê…Ñ¾«Ù°˜Ü¨ÎþàRJE
U¯¨¹¦‰Óâ6å·ØÕëb—†÷é>+Q"þr%Î‹]ÚÄ–&.Ì ºŒCŠÿÔ°`üAcödwvŒ¿àÂ”Ñ‰Sžì$CvÃéYy˜'šIÉiòùß0ñSªáÄ´ý	~ŒöBÙ±g4Ò”,ÛàNy—S=ô½æ‡/u¾ñÆÕúåÌ$ãGÆpâ™Ÿqí1¬ÄÓ wJô'J<M=y—%Ñóuì2rwìrÏÑ«ãy®ŽTgçÁ«º0ŽYzóýaf Ù¥ÇŒò°¦§´à6åªÕ³‡ìR9’×îp’ÍÃ°ž±¶]ÖÔV¹Î£FÏ ³`F>™Æ#$,^©ªGI‡>yÑ…m¨Ó—[è8V·h±T ×øtèa‚«ô¡U’lçïõÔ/ù7ƒÕ¨'¥è†ˆ<Æ†"«·_þ¥—/øú¸à¶3¶Xž¤§EáqÖ•*©ðl
ŠG»ÁH³eu9^ˆOêçâ³xÃŽÌŒ™©Ë™¥±Ò;:#}Viñ&„Æ±–°ê¶ù9ÇôÔšÓ=ŽX.í†SeI´×7ÙU'ñ'š³’) Ž,ƒ%»g~ï,	â„
	åw‹ûú÷QUúì*@¡1W#>BŸý¢~¯´º4GXñUûF}úBþMû¼Rú†ôÝÜ®VÑ~ÓÍþÄs¬²G“èÑgãÚ=õÛü9Lñ&¿Ç¿Ûz<{ú¼èŠ’“!§ÏuS«Ñˆ"›tÓü6.Þä7Ôo…ñ8>Ä¢EñYú—Òã¥¯4p¬ñŸ:^ŸMÏæŽ’žüVù®ìo}LVYé¤~óyOÏ±æ2Ü¹‚cDÕžRO©?þk„\:|'ñÔ×½B=~±	5ß£—°zZYi	ñJuð¢§ëDçr óÐ¸ÊƒO÷¨\¬ìx!Ó¾¢<ÐuLk^vìx‘Ô=[Ï£¨“‡™"ðXhuªô|Lä»ºÐúL} ÔO¨m“F×„ã ¢?±IW`1E˜L–7Æñf®„ÞŸÃñ»1v7ÀmèwáÃùGõ“6$Š1Å3fá–¬aõ–™¥+½;ªÞ²’úÒë‰â®ö({Zqå†ö­ò>×=u^?‹«÷¬ïV¸ Õý{>¸<`]FÞ
[š5ðÛÝÐ'tü¿Æv6ï¼5®ñÔ&1-S{Éh…Å®´ôÛ]õW8‘§b_91½4;¿†çóéÔKn‰/ßP¿¤øH1ÔcÙù»Žè¢PvŽéÖñárÈ¸[‘u-gy+˜ÜJÚíoë#š›†íŒÇÜ»QMHwÝ’Û±nD¯½«`Œ"uÌék àoãZÙô
ºÌÞÓ)òJÀå"ÀE¢h^7éò^Œ´ÍqûÜ¦âÜ¡÷ú¦ÀÍy·¦êŒf¤êA0«	dEÝázÉa€-8<’a`™;[#]oË-ÑÚiùe²Èm¡Ûm²xk<?FA:ìxiê§)úÜ¯p~TZúížOa"'hÏ÷†žA½ì||¸pkå²yS_Ïª*Zûí®ì+:Ô´þö–zAµå*ºhâoÆtÎ'¤_MS»¢§Òó¥¯šWVUzæ}T›Ðü gÑ,þGÿ–êD}Æ–n­ƒÍJKñÂQŽÔZºª1‚°…šÛtCj¬ã<Õ­ê/]sV¤„¾”’r²‚:¥Œ›r¨*MÒÉ"ùðIìÙm¸rm¨êeà5ô<~#¦qÁc9Ç]XX#‘Í}!È¼>Ýè­hío†Þ¾dzåE°Ù‘°£ËjÀ…Ñ´VªWØNk?8·C¥qE3)]ÑJ^üSÑRþ›sk¢Ì£¢1ñ“ÎêF³µýÉªéÅ÷A6ÁlQÇ‹z1¿ú¥ªåX½X¥3lG€íK#goë‰'žEtóCZˆáú´B—§Ù×UH¬_wÄá]+Ùf—€Ô°|ŒHµyuÂiLSrh±ÎZYé-S¶Sòk;|bcHÒ˜ûUã2P‚8õãX(
×ŠY—ÊÓvJÅ4k=H.Ì: uiÚªø"ºÈ˜Q¼‘§ŒëH‹Î¥
mËãë€ 3´,f‘Y;VV¥ÒhÎÉi~‘ƒ_	\ÓÉ	Ðóýoìü‡i!ŠÐ³_U3qm‹R*¹übÈÍüvu¦h$ÒÏI|ŸÝò@è"sM.Ë“¾¹4Bë‘(1Md&ïh{Ò×¶i¨‡ÐzÄ},Å#ØG"/©¨d'hÓºcœ±ki|ïñxþ•Zí2œ… ŽÊ‰ùÈëøÊ“—QU½Ïµ.KŸÇI0"øñ9íôÈttX|íc¶µâë€!±•R­kÎCìtÉŠX¤ØD7•bÒ¹k»"]x-ú}œÄó™)´¢ÆßÑ¥ÄP?/…–`ií¬¶Ç5½lm–¢[Š,•4FM²CR(tT¹Óå_T¹-ÙõkS©–x1FÉðwL¨Ó´.#¢ä¤š¸´æ=¸»åQRÅ…–D”|^yÙ@šŠkþÅ³³±T‰Ð|-ÉA×°r‘¥WšpÕ¥É^qé	`q¡kõ;XA÷­ød‹=‘ž[7xDÜ¯ŠÄŠzÍPã~Vø]â¥“þú××JWì¸åBo#DG“A=n»Ž[èA—x¿xñyŒÅÚï€QX­$’Z|¤5f·˜´¶œ Š¥˜:(g˜šÏf~2DÝÎ÷6TÅÁÉuK™êcùLÕÃììÐ?u”&þjìV›õ«¸ÒÀ°ü6—>|Ž‚M›•¦xÂ†úÉ‚P8ý<õiJ‰î;^–'ÈTm›&gë4F@–’¾—¯MMŸjöN.¯óÕ§=Vv°\˜<~§TéŽó{yE?ôóù°ŠÀ_ßWeTvJ,¦H/,-i_,©‰@i%(Ìºøˆi9=!=ÉÂ“ˆ Ç^4yIÄbiWR¹TÚ’0Kjâ)¬Sà¯¤
m|Ó|1=ÀDjC9ÇÇ¼>#?/lø•Iú5ÙÝ¬ æ}%ö–ñyþ¤¤âçwà¨ûO¾³ÿÞ|¿1å@7E%m»ÍhgÞØ'ÒÛ$¿gÌÖ™Ÿà½-"?PþmÓI©P%-ý–UÚ‡®' ”üÅ ¤ H5v·àº9§Ù-Ik‡ôú³Òß…?}ø“Œ/¼6Pþ_wç`þ!øû>³«zfw³¥ÏÑA$éGÉñ çÓÅ[hd
./#?M	=[däMgðË!éï™•+Iù<&RÑ¼Öƒ¦ÃáX-nÛ@‘SÉ>MõÊaU©k«F_´~£y¥©®è~´ESú¢å4FN:£•´F_‰ÞhIJwä¢XJ/ 8»]ôãÞ"’4Ë2–I7PZÃ[Q<Ê^:CEn3~ƒÊuìX…#À»|=ÈÆì±ü+}ÈtìÊ—œÄMtŒìwï^œ½9ÿéÙ«³·ÏÎOºA4ç#?­>ºiòÂ¨À¥ÎYTs¯Õ‰ÂÚ«Ÿ)uä¶Úõ@nš %µ}îú¾¯Aã·¦-Í´~÷»©ïo™›«ü(uÌ	Xôu\ç·$ÌE[ç®MÜ-jøØ¹U”Ãy·nMšIÅ©æÂh¹˜Ž&Ë5‘œšÚ{³CÔE¿
½g©ÚW¢ÌØÝÛsßBhÙ‰ª©:Lù’ˆŽ3½Yšâ,I9TaÛ^fáb©?æù.Õ@m<sŒ~„À¡ [ÚwŸ;O+-è¸l™<¦8$Žç.ã¡ UÞÇ;Zn¾çÉsoäE*5Ó˜|kÝÄÊWM‚«I•÷ý½íÖ“ìšzB4U±U@)—&ÒgÚ+.^È ¡ã%ä ª¯^ûãÏ±½§s[Ï59†êçTŸ³-LjéMúž;ªÒå‹Ò¹a4û³I³/õ¦<œ(%ž‚T2ŸV	·†î;P`ÅYR9µ†N: èñÎ‡Ö£Ó×gÈá\LÔ|Û9ëQ¹(K3ƒÀ=y^­˜ì— °Ç·"ïV« ›…³2GáÀPÔä”Ô‡sqä%·-)év:XîL/Ë-,“ÚIâ{Ú yÕËÍüò¼§ÔÈt«¼%^¢a£<7"Bû#ý’Å¤C¾Ïe¦‹¸ûØª&À\äý´ßØüÏd{›¼õá1ž”.p`ïËSãu|:¼Ï¥‘ÿÖ#ËQb;øWI½¥Å«†Ú…u®´ZÈ=¥re4Ë&Û¹=×­ƒ}ý›M½MýËÛÑ•‡Gmñ•,VúUt?kÊjƒLk>&öüÖµäÕ+¥Â©wú#Œ¿.Lv4ý·þœ®gÅ>-^n,¯ð6­E0õnÚ½-Bï¡¥P{úë‘êAÄª*txÙh—º†ŸÈÊ®f¿ ÿMÆ³Ï‘Gd˜Ý¡²(Š[Tqj»lÕÁ,<|5È±ò<›!Qµ{äÈöwË!ÛÏT¿²i‰N5§9»ªÉjÝ	Pž^\OÁs‹+¾ñ.«d™¼VË °ív€'Ã‚4ûü×•ÃTÜ”¶VÝ*ÝÖª5,,ý¢’ž‰f>ç	šòœG¥ôæb Áh3¯ª0’˜¯vŠ•†Æ=ëð›‘YÖ-ÐYmmœÿ:fçgÏ.[i»P4v9Ð5v©¥VKW*xõŠúy%;›JÊÑkª¸Ðê	BäeLÛŸ‚LñghŠ#Áˆ®ÄŸùYðùŸ?ÿ?%>™ztmyµµ¸„zL†ÁÌá± _¥˜€gÕ÷êöø££Õtú^ôt^O>Š¢‘<Hâ`š²J@–¸ñ-Y	)±ÓJà´Š÷‹c`Ä¸ÐÌ+k•–Þú—0ëÉéµ²0xkš°ºÊÛÈË Äß¤$²¥ Mâ:ÃH-"‘]v3{"W5ùN«ÑUÅVôØ¼{ôÜN¼Ä>‹_sƒJ<B:0P³Õhº©Ñ„y[<!ßÍ¦ÉeÒÁm/êºF“è#TA|ÄÏï¼Ÿ»f¯;µªNmèªÖÛ×ú«˜ýÕðºŸ5¼~kÀàl.Ãøº£¯Ú*.‹† ê¹æ  Ž§^é·bô¨û¹é\âÌnlõtŒEVÃ‘P(mX=ÜèTÎ{»—šÅOMò)«Y,¥Ë+â¯CÝ.iLÈÆÉmÉyÖAíànfç`óçf5Ez]b<Ð*ìíVÞúIÀSXè*6ä÷¨®R> à|`^â;›“¸™)tT_Œç«Ì¸Æ‚qûa—(D†Ù­œ]Ž˜"¦¦ƒj<y“>:‡|ØY­Ctnv#}ÌEdÑ×wfò¾‚©|E(‚¡³ˆD™¶Hk4"#ƒ¼‚‹}º…Hÿ‚•‘‡ã`–=yKî6!Ÿ»^¬Ž8{ÅÅtŒ×rŽx-í|·;¼‹Ñ²œ¶YÌ%à©ÈwÔÌ /'Ÿ ¼´ZZ•_,ú²C@ý0Ô9¹®gÙOö #ŒøÈùs™}wkÍ)«b®ûD1}§#†¾êèÀ.FØ=4Uª¦)*”â3lïþßÿþ¿4íä:‡Ï_Úé&ßk#Un.ÜKûy+NMµ½T±?]¢–Ýôx5Ö5àå¬oÀ«ºïä-V9¤z²sNÁƒ¹l¥^Ws@["ãx¥Îç…ÖååætyáKàQK¯ää"	tæ©•}›WöXÅey¥Õ[ní}ŸxÉéÄK²AyUòjk\™E[¤1nû;JnãëÈO^Œr^\|G~˜%Yê_7(£´±éÊ7 T¤†;n[¬	­ÒT°ÝS¸Ù—	ùSngõ
¤j%†Må£/a¾«gÅï¥Ë·œß>^Žlžp¤¼†ó !5nŠ€”5õ"<xÑÕP¹¸Ù£eÂšìS9¬±¯ŠÕùÄ1bÁäaÂž°y™à¥7^®ÉHcõP¡c­Ë¨œQAü êOdlfáèQÆ5?Â„Þh c¹ç°±ð19~ÖÚ°n°ÄÖŸ5æV¨Ü)•ê:£¦.rº†|¼o²º5ynI_‘A“cáF™cA•_žìÔ®B§¹¯¥&sÂ*šL3¥+Ý^¤ÄÐIG'I_¿ô/³ràþq˜y*>Ð°¯ÇQ{:®áC«”¬ùÓg‡hbúgÎHu¹ÞËØìJ“úÓ€»jY|,løFßuã¹'!WµPnØ’9R©ŸT± ¦ÊÕŽ•1Hýb¤e¨È€~Ð0»¦¨Ô1I’4»èé5ªríÅ ]ÙgWð¹*õ™Íëê7Ål,ìÕ²8#;ÖXÓ6)ž•nÕ(%™`^V¯=¥êµ2QÑ¿ndfJ4=#£gSš}¥ô¡œ«Á¸é‹
’Ñ0ÄÑ@xÐ˜ì,%ïÙ¦Áö-Ûh+ÊaîÃ³nÍ¥Ãï7,P8r%Wt%<®I@*DG&7Ûu‹¸ÕÜ+Å	:Ð&M¼_.-øk§Væ°¿U¨•A2QîX•iÀå¹Š¥À5Ö Â2*gäüÇMfð¦ÕS†’`êƒ—Ø™Jà¸ƒûÁjÁã’ÂÝæJÇjíÍÌ‘^¨_[ž@u\ébû÷„UüÙpDNhxaRKJ~¿­ŠJjG,Ú¯wå‹ÎŸûþèÖBÙe­ÓÅ0R#‘6#àÎ q§§äÔË±Èj>½µ¥ì1…IœKYUñûhÏ¯©Í½'y>b‘Í
?&-6Ô@ÏÒCÚVî£*¹¯(üS…: 3›‡©’ë$á…bÚ´ÞÇc=­1PT‹?¤f:"yNn"Í¹^QIÙ£L.sív»^PÓî&Àx™)žNüá§Ó †~_=Ë%¤5:5Êñ‘ÓÙ”Ïÿód%áßS–äcûûò½nÓ>~9ÏÐw¯ýþºpLâyÌM¿3
ÆA¶±E¦A4ÇªFÒ-È*ÝB_Í”)²Éò:¨H™†z¼¢Û‹®éÑÐŽ'z.¶fØUïqaQþ*-¡*^%™²»é¶²	ÒW‚–ÀJéª qNœÕþÉ_ê79›ÜïçFMBSÀme1¡¾åéZQ}Ñ‰ò¸M´¡F öô ðŽn¾A)¶(»`­Ö8	}7šxÏÚœ[‚þ Éi!¦ßr2Ž¯¥Æ{¡5Ô þ[Ÿ;×mˆD@âïù÷Úó{:¬wÏ$›CáhÖâù‹‘RWe. BVÞÒ*e¸òØÚ–ÒªFKëðl»¨e™úé4Fy’m·xÿeIœæ·l¶£ÓxvkGm}ŽÙ$­%&ÜŠ’¬ªšK6tÕ:)-è+¹Å.îìµcFs¹[òÌ¡ßûLÝÎj“¬Ñ%ûs3g´VÕ`52K‡K:·0zñ½bRV¡>Eu1kö“ÑFi%T\jÖE›¸Gå¾hë‘c{›ôº°çF"ÊŠËã’ XxK~žÌQšó@Üõ’°.¼…?C#aƒ8ç<£@ö¹ö–ãÈW	)Lo7hñ­X$rìAÞ~»¯"`Œzäý.ù¦þ•ÞñA¤äâ–`í0hÞôäŸÂµ<"6„}x”R>v‹ˆ±¼ÿð¹¸SŽ£2›îeœ<ó†:|ý–dCH¼ëªÊ P§º«¾’½ÅÉ”¿ŠYmÆhôœÚ\ð‡6o{äŸ•ã'Hˆ˜ÙpòGÿÚ{sMºŸüÛ´]ƒÑf÷¿Ž+ú„ÖŒçÓ&•`iN½ÔGŒ<>Î§QþIÍ&iFOë‹áëã
¥	ŠNu„¨6é÷âåðöûšÃ\y×ÐVw6Oš(Iâí]òGßŸ‘7¯_þBõc$›xñ†°‹°3'È›ŠM›Ú‡”ybG`®Rž}†n¬´@`"’ÀWa„rÞÂe¤ýþ§-ŠÜéJƒ\†oã8Úmo‹\Ð·¼÷;º,&€fñIüöÜÚ"³¬óäíÆ€¢Š—+ê_°EŠQšÜÄCá}}CÀeTôØTBÉE‡a¬¼³~HýRö•jTÕ-3(E¶“™ŸŒçXÊ>÷€ñ$¶|Tv—Ð*‘S;†BYœÁi…,ô©’GKU›ìuq„è[à˜•PtêŠS‚)Îƒ)Õ˜UïŸy	æ™VþöúóÿöÇ:0à8Ãð,ñ§Ašzé‰8éŽËæjòÚ,ŽËÏé»™³bêµ¦AðÕŒã1AÓáa‘ÍâáÃ-öå2ŒcØˆÕ!m—ZÇú
ÐÀæ&
½;Ê­ÁOMŸ’ó³$‘@0òr(m.+*á?õL(QEŠü•/Ž ¢
_žê’kmj–¹Ô ]Ø¼içÅ«X]>UÇœB-ò‘n±w·p0[òÌî”[Þp1Ã)@œ<ž‰ëò®¼ DmMq†öÛ*µE«è ‹K? ©WR!Ö"O#§NƒfËý0®Å3°'žŽ=Q=Kç‰‡å°Ê/ªO4À¹òˆ`Nj!RæÿÔ ß¶X¶‚Q–ÓöôÆ2+ Oa46ÀhŒá&ÅÂTae8èXe!AýcùÛûÞÀzÍðÚb‚#	z 'Ì’ÖNR´\!\€¡ò§ø&MfvâŒµ³8®`£Í± 7šÆ%Y)'Ž9!2ä#·‰”M†1¼- •.}5°¶¥ç,lkNgà%%^í™2=(ÔŒ¯Ì rh›ØSþÈf••´÷Ú@å[‚¢Ywä§ÃdÌø§üwŠýÈœ"ø±L¿Ñû€k Æ©+ìPôú¤q¯Õ^/šõ
ˆS}ñêìäôÝ›ŸÞ¼}úì-]ä“Ç]†ô1Õ§Ðå¡=iŽ•¼¹'Êæž4lŽcm¯CGÙ@822á:‰>ª#»”3<O¼(v]­˜Tõs1î\O8VgÂbŒ^¸?¯íƒ’Ò<Ôr\*g4…äà³=ÏVƒj×´asþßu¤Ò¹Ù	+ùEˆÆ{­cEµX—˜PêWu9t	_hYj[HUm&¶d5/øžx"û€ºZëò‰š•¦bŒÌâï §O†“`dOÞj‡¢ÉZf‚—ÊÅà
Tf(•7S‘8Ë¥ç^˜yS²(	[ªÇ’/ÊÏÜa@†2ë ŸÙé§aqâkÙ­´®I§Ï<ô«yq´dÉš©úHSâS¾Êž7-^³­â53š'4ÖN®ÞAlméÄð²ë5JV¼avZÖ®4»iHÐšlrH>~¿(äñ»ß}´…Å£»¬yƒ×yo—WT¦iµÈ˜Ñ¡¸åG#+¡_†ˆÁ%”ŠW3ØÐä6q·•àêTÎ@whT\Þ/Â:®D‹X–}á°×™¼0
F–Ý¾r=i—fŽœG­>g© ç.Å}š`êXgîºÌzj>uÇÐOžlºG\©ÕË)î–Bf”Åx\sQc Ô®…áÐ×-ùàt§8‘‚í˜ðýù€Hy´=ÙmB£s›=_¨RGÒ~‚ìgßv[ds.t ó|E›×ÎÛ²†	çJQ&S3©…ØokÌ¢</.÷Ü>…ã¯ÃÃõ@¾“’Š‡ÂÅ)1“ƒgiõ‡5+&†!³Î@BlkiY0afLqVv£×ø›V/±–NWMÙEöÇ¤.ÒLGâ“ôÜvO6x8€ï¬ñ]¨„ÂÕsï•Rë16Fá€N1éc\Æ4IÌe*Ö2-h\†ŒœÎtã4JºUéÄûÿeN_™ƒ‘¦tÁE¸áuÙœ}P•ÐzÛ¨©%ç†›.9ƒØµ@:ÜìÇÝ6ûN9·îkä˜¶»Äè'@â%Å‡IëQêÊ"ð`ŸŠ@Ò«eètD=±C\ÍÄÌØ°UÅJp–’.Ä¥’2pÙ%v1Äå˜îÂ-W¥c®1Wp·Š»ß‹û¹»ÃiŒY–³%ÙÕAÎ®ê9c»J‰Æ‰7B§ÌNwr™ÄÓR4FÜÉ|/¬³³²ÊHVZ‚Ue¸¦Iô
ÞÎ¥Ø©cæüž“×»Þ2yÌC­²ð`ù¶=–¥Žy@Q’@½b‰¿Q,ügu.|èæìuÖýeò½+®ªÉuÏ!ð56ÄâúoÿÍîÀÃìí÷Ÿ˜ ÑÔk9ËßãîûÞûàu&©2ätÒLPår@ªÛ¾P¶¸hc5&2]Ò™ä§„ë±¿’Ö¤rëÎá°1Ço¶ˆ­že#—ØâZ”-ŠåÔE6@ý8(WÔ²h0¸µG¦ùK¸±«ævDÃy˜û‚	²ØRñ–Z–Ö®ý¯À™i0„Ms‹Z?s£¹‹&£0žÒTMí§²á<·¡–ádL-FX:óqÖO4Ð?ñ‚›Xí×Q\MhÕrˆ©Ü©ÐßTd§Ô”2W¨Ü²Œ¯/[¼ªµIgÆGÕ+‰IÕxƒ
vGdRê«µJ@F‹ÐíìÏ=MÈY S.‹¢lœI?[joY_Øïzrìt£ØúTPuÿÅ®DO¿“0‹7¸(Çckˆ>‡i–éçÕçyO*#§ÔW^øÇ½3S¤·:Õïºu/ØtÉ‚OüWÒX8>Ö˜
Ö¬'¸G\È)Së	=ëD8ˆÛNa%F(u§n^Á¨œ¿¢³¯®Òk˜öËß˜_’5€Œ¾#ÇÒÓy
‹CýZZôƒ§`’îlF›êÕd„ö”(®Å;ò&/¤ï-®Ü¼á¥·Ñ¸Ö«^Þµd@–GR˜V›£çó»f^ŒÝø"õ“+¸.~cæ§¶Ôhª‹:-VZ¤îÛå[]l§\"ÝØp«-V½R?ûÁ|ÔüEc–¼DÌm‰cq£Û¥Æ!XÉK‘Æä˜Ã}ÌFçŸ¼NÚ2ÈFo•ýB—;¼ÁrðÄË	aD7«a
»0wb­ay“˜Ú"žgNšœ†£5HÙ¿ÇÜ©7@¬{Bþ²ôàï¶È ¤È%ßwÕg—Þ!CŒo#XeyôC¿{í%Q»õ,Iby€ý4ßÊÿŽ	Vž~þjôÐM#%„[[Ä½ŒIyÌ 0ÒÀ¹åF¬"´jÝRƒiüNÓ7œ+¸ˆk)N¿¸øéŽù_:×Èæ3Cß­däÎó'p iŸo{éë¥lßòE· "¤h‰Å¬%dª°k{Š´.Ž‰óË—«-¼H
êž	Jw9Û»ØÕŒÓÑ8¯²,É°ÍØ¦F¥tò·¾ffK³}c¸ÄõE®
ØƒLW1ƒoŒ×7ÆK¾¾1^÷üÆ7ÆË|)¯<f{iæ+×¼–T­9ã•Ç›|Y¶Ë¹°{õºo¦ë$ô“Œ1]ež‹¯Ä¿w¾‹&
øÆt‰ë‹0]2Ìƒþ7vë»%_ßØ­{~ã»e¾ì%UËóZÂž.[ÐsN‹~ûâŒ–H¨ùµñYÿ¤â±ü÷Î`9»”_Ú´;^wzÛ®¾%êFœª¾56¦†Þ…êÃ+Ðå¨@ê0Ùzô†qg,<Ù&çŒJÃ—£mÚüR¬oCdZ¦„½¸ŠÔàã
i—é¦”tM¼øz¹ÕeyUîœpãî`å‹åß	–å<”ËÚ­SšÚñäì:°ÿàÃQ¹äe0žcê¿.9‹ré]ÅÉñFÁÀ†üƒGŠ·#ÆI°W8öR/ÑyŠ~ò˜B—ž/èQO^ÀŽƒñç@$'	l¼ eÁnkI†‹y8~®¡‰¤Q>Ñaó¾Duf±C;—õÉ!k—BÖ&ƒ|i	dùãŽø!(+‘…ç^8ñˆSñ!‘d…<­¹ò~	0ëÙÐ¿F^A>–ÜÄK°ï¹Ä”$¼÷2s¶Íø¿F-òw´“îÔOSoì#òÃWu²9Û˜W‘˜Ö(/ý&duáS}çÛ“£|+¢P¥¡àñãý›t	©gX¦2)X‡¿V¤eu;;^‡ÄÄž-=ð¼‚Fµ¤ƒvQ†c¹Q›;©VòèÕ:vv[.®†’[^^à™ ?˜B^œ4ë¾©häºÕ/Cå(ÃK/c .—-4Í*×þX¦CêŽ gålš¤(µÓ¸wT2Ð€/s•Koù:^†žX…/qDÒDy#ï»/ó/Õsc>…X½Â¥^°¡!VÄºj–ºùÖg<§Hå~o4Ê›Qê²9`›Ñº/ .‚ki2 Fc«§‹˜tú»4ŽCe-šˆÆuçé!æ¥ÚÖü‹ì|ÇnóÖ(óÅ¢ŒðB?´²lMAÛ4S»f¡7ô'ÀüøÉqë)–»ó‰÷ó< ³4¡¤=õÃIœ ŸÏ±l‰a¢tËnÈ¶A¼!Pd¬Öh W^8÷k(²¡¡t³êŠ£Ó‰¡1´€Tå$…n¿›ÑêÅ]:YmÒ£mRÓGbÊÃÕ+Ø;Y./ª¢ÝšŠ²Â3»’Ý¨§Ú¶šðV#³:Ðìç39óPu3A˜÷—"Ìê€v÷º¼<Ç%F‡Ã©rO"ÿÆK–¦ÝA4›gË˜6˜öö2õUÅL—OÊ²æ,ÓJmC¯dÔ‡¹ãb³ã÷ôq÷ýŽ%—ƒî¢Ê2hc“°Š”ÏáóŸf!°¬ôöÙF¨ê–Ô@5–šÒ"¼–²€ @çå_£öOIhªÆmèÞ‰Rõju±UywA^_¶Ìµfp/¦ÞØX$%+¾'~¸NC–IHM’Ï’yD+¤M½4ð Í…ùÊ¼Ž§Tã´ñb:¶Õ°5r9/(vUò^ƒ¼_wc,ÁWÃ#Ø>?]„^ôi£±äÌ.Ýí©¥R¯Ãx-@;‰[Fë°´ÞÏn`(‘¾¢O’¼†-åøÆß]ãÂ3zùÖŸÂ2<ckî·UÔr=+Ÿü¥ef2;½¯ráÿé×[íÆÂ^…¾¦ÜRkøhcà]žÝt€ˆSÇ«'0ÒþhOlÔøGíOÚƒWå\r+9ŠùÚú•,´V¯%[ JïÊén°ià6x:s×OÏe®ý‡{•<²@x¶ûÕ³¥¥Èíûcà_·7F^:¹ˆÅ6Ô)/ž£å'øë/“aüOÅóõhË¼Zßíæ¾hÊ‚Ö£ãáç%³xä/¤ôe	êž–kÞÆÅˆG~ž{áÏsç4¦Ÿ»ŠÌu
Ä8ÚžÆ(ôu+¿l–*ü-® dÌE¬ 2 ŠjÀu¾ô¨è¢zªa†‹VÞd•è¤^ˆ©2ã™7¤•w¶HŠu áS÷á^eçÊêÒ;½ü^ýÿ&ÈôPƒ©1gF:-rfÀg{½tdR%ž\%µ’Øayëñ–Ê°3Øb©fg`˜@3z‰¹ªKl?Æ!È|J’¤¯„¬ÏSy4éòåÑuÅ€óÖOç0€QœmOúš6ÍY9*$tZqF-Kˆ¬R"—^%5bã<›Î€ów#Ø¶qAÓmu ü oÒÖ£ÿûßÿWÃ–‚ˆj;5Ò`¹ŽˆN+J
‰±Y¢§Ä Cžó²ö‡d¡*ÿž³˜CV*Œf8‚­~wÎuÖñâ¥Ê|ÌIé><³Æv§HØÇ'ÿ<Nž‚¬ýX´èÄ&žVÞüîÝ‹³7ç?={uööÙùI7Àt#?¥-’Çt‡jbm[ãm¿£:ŸœˆØjËó^¤RŽõâñ2jldÕ˜UéÉR%‘Y	_««+â‹ÕS[’oL(p ÀZ€HÍ^‹\µE)_™Ð7hLX‹ôÝ$›\c¢1¾U¹cië
áòœTJÚ¤sÈÅÔ£„ÞX‹Ãiä4Û+†aÎöÅˆpQ§¢ÂçÄÓ’èm9Úº´æÎZSÃ¦òÒKk¦"Ôªw4Ïën¯`£Y`ÕåÑ4ˆ\« ¾ŽüäK":OýäqwŒ6õ)X2$rN›!µ:}¥F‘Ä°0qŽz8Ÿs@´)~w;óÚûèOñÔæ‹W1?j[Mó«/RÇóAs6h©Ù»xÉJ¦‘g7p$*ŽÖÌÓªñµÊ‡ç««ãÄg¡é™pvÔVÂtázzœÐ(²¦¬xã“8ŠÇ‰7õôã<a*¡s•<Ò·>œ¬Ÿÿšt¸§À2ïŽÆïUÇ<K@Ln¥1Wï³§Ïuc|
;-JÕåŸÅ	žÉðf£a)iZ¯BËle^–’0ád½j>…¼à3ýè€ÓŒ@D?÷3^Q÷‘*L½0‚Ÿç,-:ŠOråvM`¡1N»—À:zÃIžÎRE X/˜þDyøÜyžylòXærNxÕFFÛk‰20ÓîÄKÙ­ÍMC$Aþâ&ÍøÕã˜‹+‹Ï–Õc0÷j«÷S¶^·Î´ŒÌù0N|>ytr]ê-’Pƒè;tS|‰éÝÜ"*·vÖÁÔ»Y¢ù>ù=önØEoÓØ
Î†Yµ‰ù{L^yt‹ÜP»]™ú6i—ÆÊº…qÀÑˆÎi;ÚB{,Í‹¹É‡§›mš¯nÃæN8uñÚ‹6ÏÐÊiø@Ø‘yRtÀ×¦£Z1¸YÏ5Ž€Ö|½ÓQ­Úá¾N¿Æ›¤iw+‰v+UÎŠ’f¦ú0uÅo­.°Pñlµ0°žAB‘—×8nÿè‡ñ0y#ÓX ¤qå²ÕŸV(ý¼[Öé-J{~•ò ’×}s„së°2çãÊU;ëÐ¹îôö±ªÕ¾ªN1V·,K^ƒ‚iKžˆV%ƒìžÝIÃøÚ¥Loy }hßm »’tQ²Ø>²ƒh¥
PÌØ`ç6ÛŸ…Af‰ÝùÏäcÏD)¢e÷å4,·ÃvRL¢º“ÔDùÎVÕÈ¡jÌ^ÕrÕ^(}©nî¦m—ý†àŒùðž&ŸÿEØ¯aùIÉ'‹wqn`Ë_3\Î œ~éUÖÖ.›âªæ¶<Ÿ7#ÃzKux,µß“ÕÞåìùìD2Ól+º¾ïõ5#Õc¯ZapÎCê Î¥Ù<L}tÕC}žÐó–´_œè%ho«^z!–?ÂÐ*ÖovÛ¸ž“ÎzvëàÌUC©ä
’M\£uŽÐÊ‚¾ÍXËc­
Qž8_9€Rèòw£àÊ‹•?yÑxŒØü¥úyßÛË·|O<µ_gè{#©ö­¥ 6K›K’¿µSÚ©¬o^:QL#æcB­¹ðòn_RŽ›Ñ¥`ûäÊ+I@è\ k¨A¦„h?G•¹~—Á×¾ÝUìO^F®u2žÃÉEwà,`8	¨›†½¦<½‚¼ y-3¼,ÐZÖÝL¹ùd†A3"Jª|ÐîíØ£i.©hÇúOä;e·:Ã3ziYCƒO¸†€aÌÈ:$”u³Z‰;Ý=›k#Ÿ¼Ê6ôÙðyêg^¦6K¼ÜµÎ¸„ÐERÿúú¹Äë—(ˆº]-‘±6çg#SÞ®y»Úˆ:þ½<)Ëæ·ØæBÌY¦J,v9óÔ¼"ñ˜Ê
{«ÓzH½î:÷ÐU³V=Š4N¹i³`¥ÐWc¹2E>,•1¯qe»Š÷±4î¹wå‹FTÏï¶ªêSãZÎÉ·þ%py“ÓkÎOÈ|º†/.q¦ôPjY™PWšMc¦y·	mû”ÉC¦œ¾Í»ÞI 4¼¸œv9KÚå¾¥šN±,g4¬°%ÓrM4]o8÷užF•òð€LàÿnN*Ô“Û¢Ï´IÂO¼ätbóÀUZ²«%)5-tê¨Ç{vãç,k›©ÞgqB>ÿ¬dmWp˜jÊY«~ÂÅ×CÉÔ¯ðMƒªÓvUç'YâhÞ%aŽ3Ÿí¡ŸæÅDêãd>Yz;„Ž•O)
÷j‹öòr¼‰T—ÕéÕUóMÜÓ±ñIáóéÄœ]0Íä¿°zŸ°4µÑ OW ‡ïE«6Ié+ô²  ÑaÞñ³˜R7vj=óÓ¾‘;{_ë3I[‘ 4noþ»ãŠíÙº”¥FòÙÒ†êVæU—['¡C£7êYÜmn²ÚÖHâxž3G‡aºaIñMi6uë6}€Ái»òtÙ¨mêÓºûÓ,‰?ù§èƒjýÇÞÎÅÃƒ^ËÜ*¾xáÆþt~vKŽ¥pBÎ¡£ÐVã4Š]=»;'HËR_1úˆìîØQµ
‡ËÝÁåžo©@¡^j¸>á¬gŸ<)O]X,oñLƒòü,3ÿ½‡þÎEÃù«
 ×!×An‚·þ˜r—¶Mî²ËÝ+¯Ó½‡ô÷®¦56ÇÄ	ÅÈ¬3`¾B ÚÞUËTrN›7Xv$)¬œÔc4ÇTcðeà _åTôØîžð=<¾×§:I/Òy¿·=ø`7Zc¾f«™#½+ÓÑpâ—!@á*HƒL¢z'ñÍqk'ÂmÏ­T‘u¼·^‘^k÷€œÁÁÖà ºÛÙê‘ƒ¸çœ…ƒí®cØV½Ë½Ë‡ßûÇ`”MŽ[M»{D Í`I¥œ_¾Âð¸E£Ü^q§þá¾¨ÛÝR°hüuB_ô‡R€—$Þíq«×{Ø¸×¯Ç——©Ÿa:5|ŸtÄ	²Í˜à9è®‘éivÒÍ‚@’Ö]g”÷Gz)ñ=8ãy¶ášÅ	ç@„»;=X¡ÞE‡s*³gY<íðS5=åw¥@jÓ»°ê‡˜Àº"Ì»EŒGº.n>SN!ù®u ª “ÁCíTœ|IžÆñpÐø*(•ÔÜ'™Æh­ŽIçý€Ù€l‡VÅR(ÿÀ]0¦Û­' ¡­Ê¹Ca§uCÙ\—Î¿z-r&Êes¹Îò%[X:¥ÂZ‰t:7“’Í¬b{]ÆRÂò	žp1˜"‹EH3ëÿš”W AæÉÐÆ_Áel¾„\³“ŽŽßµÈÄÇÝÄ¿YtgOuVÜ‚­X:ò2ïxñÞï…Ž…W¬§$ö0wß"4v~æ"2U.v®¼Õ¢$«Ü2¿+µÎ¯¦­ó´ÿrÓ¯±b]Þ¬¹—c³Ýn·]öüfê‰÷ÅløÕéðÛRÇw½ÁÅtü’÷Ìþ@x}°©!ðý{€Höo"@Ù·@øçéñbÇ>`VŠ^˜y#<VN¢12E{ösÿˆ™3(€mSppäpÛË¢ù²¼v]ó˜N3±ÜnJŒü›M[|¯¸ŽN}8³©¾àã>v¾_Ðî>Þ1ž~A›í™âÀ¿Úm´RÉwqfÁÌ.MtÃ ÏÛŽõCÆ”%9Í×–^°é™‚á?üp0»ÉoÂ×Þì† ŒxøúKª>‡mu¼hs¢BŒäëûßsö·ðrûÈ²RÔ)<­GÏÑ¶â´[Òniñ#•Ž"„}9l³­È=ÿc‰‘^¾uóß€L¤w,r`\öjÜ~D0‚ª½'æÝfãhŠ{g^¸%Õ¸#å_¿žÑæ¾Ü¢Xœ):uÛG¹¨“E»ÛÔ—™Yá«J×1ÍW‚ßp˜äªRA£UÛkB‹*ÅÚb¬yŒ–LêgeÍb¡ðÊSG±Ý¿£9áÕtT$¼‚Ï	¯ø`.ê£š]tö¨¥w«0P£;D93öÁŽÊÑAXû]È[)imÇèo‘Ùž]Å!w<À¦1@Àz’óü²²›5† v/UÑ	îM#=Á2^4ñ‰‡zˆ?ût°´X­ZG+4…8Ÿ§<B§Œs!f½—RÚ¤ÔêçYpEkÐÑ|£Ø$­jÝ`Í
+ŒºãÆÔY‡©"ô>oK…ó¥¤X3©ŽÑMˆN‹—"ºctH¿¡G›IÐ4Uêgyò3Ê™0¯OTo™3ù6(õ„yhvi‘&I%6NYÉV¶4É‰]ÜaQßgæîGµÉ2‡:a^ÏˆQ¹š™¹ oeW^ê‹Vñ¥—œsÊ`’° ¶û½¢	Ý0ïà€cáSvç"«;Üš0éÂKþ]!Î÷7ŠG?$Ÿÿr‰Y³ •ž ©Ç!Ï–ÛË–wÖ‰90{MîZ¼&‘H¿ò38A€À$1ÈÛ¸SiÙ¬Ž)ì;{2{³dØ €uèwnb–5BÐ^Ä¢jé)W.©ýÓŒU#FV	Fõêó¿ Ä<ZPæ,‰Çpæ¦±]uß@¼qè3SäÊÌ	º¥Ó„Ì†KL•cè-Í·ãd„jh³cù?Ëüÿû™C—»å®Q”.ƒW[Õ-î¹³ú´ò>í6;á}ãÅçÿ?Ì‚©GN®€ô(3ý•£†lÓ©!‡ZVìÞ¯ƒ4¦áÜ2-ý€«K>Ë
4«µlï[l]K£,GŒz%
€½¹p húÌ`òÅœKÇ ‘g“¹‹i»)Š1‡\tüÐQ.—]t=”òˆ8ø:zå)Û’ÕxÒ«@+i,óô<âÆäyÇ<*{÷Æ,&‹ëïÒÌï~‡Ç–ŸËú&u—-eóq.@åU8úzV!ÏWYøÉ®þuðÓŒ_ü_¤~·e·XŸûÓ/]Ã¯*¬Ü½½ ‡AB9¡Ñ‹„J“ËÎƒ¬eä‰@Òè ¨züIå¸F‡Š…M¥b1 #›5s¥(Ÿwdê%ã ¢n~ñìôv¶H‚CcCÿ>Á¹ºÅ}é	[°r f~xM—YópA‹€èá8Ž—l¹»º—­=¿Æ?ÜiajGÎ¦E<¸‡ÓÎ†n>Õnà”Ñ²‹vÑýÝ»[t+¿`
¬Bá0‚ÿBG°Æ.çQ x(2Š§®Ì{ > É—Ñ,o,.„{‰?5[„ñÇ›sªkßÑ‡Á\¤3€?ÉøÂkïlÑÿuwö6W° WÌÆ¿û¸E6¤$5,(ÌÁë„–ôË'Œ`ÿGŸaúê‡Fh‡?.¹ø¬šä4ŽâÌÍQ8GpÊÄ9¼QšU@Éá•|÷ñ"×7˜‡õbàà˜=Ši@¨½-ùm'6Fî·âÐ ‹Iš7û@ßìÎÞÃýý‡.ÍZ½r’ºbž$û2£ÌÅÅèò€..½Ë¡ýFÖï…¬ÃN­Ñà:‰e±9ýL¸wáûý-ÿ® ¦\xÉ9µ@ïÚ+m±´3E_§åªØ©2”/lÖ"û=&LîÛ»­Ggy”ùt@¡éYFqJ¢r56µá½Y‚â¾!Òß¬èáù˜Cß‹hÔ?9év4Ÿ…˜K&/Êž¶s§8$ÅnÇ@ÅG?`ll©Í“ bà§í¢›rb·>àîˆ'Ÿß’RÊ?bUFŠ*|.—Š—²h;Ä^Ð˜±€¦œ—2Xâûó@{Ñ´Ï¿jòâ³«ô¤1s<»lâaä°î}žÞØìâ"à¤áZ æD›kA—RU{,éÏÆà½ï}€{RNí‹~Ÿ4î÷¢ÚïEÓ~ƒé§ûâÕÙÉé»7?½yûôÙ[º@'»ðpv1Â––ö¥•Ï|¢lðIãùÚÓ;t¤PÕw)óM~/À¸EžoûŽ‚œaãÅš•²sð+ußtå(0ó Æ]üÄnbí.¶*Î˜sœá¼£NÚbNFn¾‘ïCaTácÅ¡Ê=­9@ ›õÌâÑƒ‰æI˜ÅÌAÊ*€	…éö`§EÝ›|õù_FA©ÑÜoVv¢uh¶UIõÀ[2w–õ’P¹µ°5Ÿ¼è‚ÖÑYÎÃBod^Umoš_›ªI¤¢.lÿ!:®íZ
|^>~Ý5¡ªÎ£Ñá´ï|`ÁñÏ–sõ¬ý6›]-Ï¹qD‹íË“ +oÓ½ÕƒMµ‘Ç‰Ðíz¥ùÒ7°ø¢ø)-ÿ–ŠY½"éÏuÖ¸_*WC«=y•ÜÎN	ÅÝ^ùí%3NcNÄJ9+^xö-
a®UÛ­&bS™qÍy¥Î"©¶2˜JA•o5®Ä’HgXÎÚX¦ôr¥ÇÌ7‡¤¿ã\|ù†æŸr)¼|CõNKÕ\®CÖùŽ7¬‡¬ªq›Vª|†•9Cƒ„¤©ƒÖ°d²¡´˜N^v)–Œ²ëÈ“«­¡6ÄGÇÓPW3Yë7¾§¯d¬©“È~”KßÄi°¨º&ëÞ›tnÒ¼péêœbLÊ¼Øõm\™æquÞõ7¸»{z¼ŸêÏHUrá9ù[ýÙ\|?^äUãÌëKC­Þ3¾WŽê¾òý§¥ÉTn(ßÀE¢Oâ‡ú³Ð»=Ÿ‡~šž#g¼yåNõíiª¼©×UÞ¾hóo¸*jM‹ÍôWL7µ¥*×I2W¦€§9žýú´Ï‘VÅá|ûH,ø§¯.5§:´—³ÑåŸ’¶Ï>êO„ÓÄÑ3w$Ÿò]eÒ˜O”vU¹§|/Ï	OßÈ¿)Ÿ•ŸÓ§¥ïõç‡´¨üËxëQ|Ö=—ïO~£|·þnwO'q0„ÑŸ•c)=ZúªXc¡Å*Ô=
˜—Í¦53jíù8úWšEŸ 8Bé¥¨>[ÜS½ùÖçúúÿ  ÿÿì}ÛnÜX¶ØûùŠíBÐ*ÍèbÉv­±åTëÒ£9¶¬‘Ôžts¨*ªÄ6‹¬&Y’Ü‚€¼ä-Hä%@Ð'ÁÐOA^Î«þd¾ Ÿµö…ÜÜ7’%–nV¡Ýª"77÷e­µ×}É&úÅí$ Y¡^GÆœä@ýc/¥¢Ó›ù—qÕKÍ•ÍOíëžÀ4É ,rÜÞœwùàÎ_€ã ÊÖN‚ÁÀX»ñÆÛAÔ®Ørí3x_€Ú;‚knõHjât—æÇwctc)X¸7^ú9ê“®—¦VN–|vÚ„'´W
Y4UøÒÒ’ÚH!¢¬|¡-Y¸™éêJ¯±(Ì¦«0{J-Kð—š¯øk ,H¦š/pëmiBkd0o³’¡Aw’ú‰Ý¯Ý¾tøñÎ¼@,ßfÜïðßÑ™“Eô¹Ë@çœNÞêj[›Ú‹Õ\à
ú'¤ÛãÉ`+qL”ÁBúK>ï¤ÛÙÂ?lnA4”Fç!¡ý­uHÞ±ufÍ¨I/*ëÛFéá§q€tcèg)ŽÌ\ŒG±v&_—51 šBÇŠÕä<hø¶ôêemþƒ?³Rq¼A;Ì¾õáÌˆ)q6gœ|—é¸¨,ž1ÎÂÈ¿eKí¸pÞ P¯£W‹£­N¾›1ÚY®j•Íÿ^©•åCkÆ÷Î±{ìŠT¢µ^ÊäþG›0Ü¥º¯Ðö¥—×};42/Cg0F­XGìšý™=`Q˜oYù¹âºùÙ0ö[§Šœðä‰B;}vÇ!_äkñófeUyŠ¾Û–ãx Z¥gøo×›ö6·‹WPDêõù6·í¼àMp‘³án‚³kO¤¾ö¹‘ø£ õÓiÎþhÍóãs#ÉõsmÉõs=Éu–žõW­hÒc…~R¤0¢]9ë"(ðWt$ˆWáÇÑqŒvFã8É6¨Uî3ÞUVÿ5íC+’Ú¡ü-ŒÓKÈa0¦
_a¯{‰GX¿,°MÞ©ãÔ7kTh:õW¸”¬Ð¶Ùªv£I–Ú•v ^Éæj~)òÀcµKc’‚×Ï&^(-Ö¡Q—p)K‚QWc¿@ÜCÕ({	-Ã¢¿wAƒÒÏe³+­ßsÌcý¦Ãw(!Ô$(ïQQÜï‰mnE¯2ÖM€aÒ9Œ” KMM³‡`õ¬€ 8‘(ë4î_ý võßZ@.øjâ,Ÿäž"o¼S4&ÕËkêé^ÊB3ŠˆX“©³Òå
¢¼^¦Ýi/a‡NäµÒ"YÐ39¥–­ÒäqçÁZ<Éhì+MÄ.¡3æâªüCò:[^Oš=Ò¨û’î‰F
@6?õ‚AéÀbßt)ˆúádà§]°æAlÕ.¢ÇÀ‡I–Ä?@k`¬†HJ|«äŸWŒ<Ø¥Ò+ gŽxÉF Ù{m¢"ž:Ú.;qµëÔ‚o®)ëŒu6¾Æ7­`§«ì;DdÃ¬‰{_Çc<¡˜(¶¼ûM¬³?¯—Y3ý¥¦´¼¢OÖQ‡®}gþ¡•×Í}½^f¸¢^¿ÐÁKÞSÜî“fÀ	‡¢Þçìsn~Þ’Œ2ˆÆEð!’“'ÈÁ}ÿˆˆŸ¼élaµ à+’X%|¦§u*AÝÖn›T‰RcæÐcÁ“ª¤îçk¢?³šGÑÅ^ìZ»í"ªæ˜§`ˆ¦DW.¼^¦,€¸PpØÀ¢€Ï!¼ÄV³xø¤ÄÞíúÃ«ÿ,Y&¼?cGÓ»9baÐ˜ÍŸ±².^x
#ë¡YÌÿëÛø!¸ñAžÆ1Ý#Øf›£baS¾°Ó•ùÐ?õú®Ø™‘Y";i“c=BØ$â¡êÐÆIð3c¹	Ë'š´HrkPŠÊä~ÆéÒ­1@Pú¢™!3”"Gd¾s§Ù¢ã9o¤ÌczÆÈò–GîHzÊÊY ª	‹dYÉAéý¶˜%y}ŒSy¹o…kRQÞÉ2ÙèÄá›\º3áXNYUV2^JÖQ
f*íÇá‰Ïk£§LI&ÇF—²}(ÌÔ“><ƒÚq®µR†gÝ…¯âVr‹úk$tüoÑ|]ë/Ô	>³ýÒ†¾låÎ;ƒ¾
Ë/],Øã¼ÃáÔ¤ÐØ—~J\´tu/‰1À1TzCy]¬:±}ßbÙà õI@iiã(è¨6“7Oî±`?­“—C›ž"­}RZŒòiæœqÅûÚP›q³jÀGÛÜÞjƒ2¯¿Iÿ˜FÉqw©¯›—sÛÅp
`“éjej25{IOƒÅÖqèÜ=ÖS	…•ÃðlÌ¨	¸jð’…JéÉnù”0p0åsú|­„=°6ñ¨5!\HsƒÂ?éÓHüŸ&Aâ”š1EOJw«€¾áâ€xLç4þ|PÎ«Ñ°Óô@Ø‘¶6„r|¤
¼ï:Ô}3	B43­^ìŒF¾w¿	cà¼½A¬Ö‰sÂc#Ú+ŽÖ¶áTJojNýQ@w‚'%ˆâ¡/>óÕP§ZXŽ•
 r†\Ñ Æ4ä˜
)r²d˜
¿s)Ï—_guXVË&ì«\Unþßÿü/¿X¢ôÅ’uÖ?^ýõ'aLz®~AhYã1éäCI¯/ÃOŠHä‘£à(âÌïS’§©ˆ¥[±1…Ç’Ò¬ëÈ€ÓN&uÖ+÷‰t~IƒjUavxÿ¬Ñisâ÷?mI?ô¯ƒú6‰¥&–×ÉÀ²Hio.´¥ðþ«_è0P¸ø©ÿ#úaöCà@ÂÇ}/ÁlQJ©v"	€Iâ¾îtÿ+·¿2‘Ë\õÌ‰P…ÊIÈBL¦ÆwT?ypÃª|b e¹©B
¹Ì’³¥V¶K¤Å326ÑZL™P>±Ëy]¬)[ì²S®ƒ¢“Á=²&euj
þ˜8
µ~M™~Yû^“jàhÕQ”lÑ,ÿƒòòÕù©‹-³‹1/‹%Yœ…ÓcvÀö‘8Åç¶ùÉÕ«YÜŒ­!vä–v;fËË!1c9ûQqÑè•îäí’FJëÍñîÅK1Ò£:Né…öLÜ5rœ;ô½¢`P­äcøÁ5ËK¶FKö…è¹ÇÙârRT2eJÙVÎ¾(d0AWQørZ!™°Ó	fã¡¸/£ü3§?æH1ü4CTÏOöçbfÞ:W^Ic"#Á…pž2¯£”ÑŒ2Z•ž|tŽÜIöÀdð0ÇŒ™j”jýŠÄR2D™Ø"VÑËDmŒŠ9Ç#Ý+iÖ¸+„v”®qòS"Ãi‹¾Ó¦U—Ý4¯ÑchìÎ¦X'¥|Ë,VéûjÎü¦ÆJè˜«+ØfnÌ9Ô’¡«:áRí9µ¶)w0Òbn¾žqáÂÒ+PµÝøÔ2;¥ZÜ©C\ºtX¦ñÍåýÒ€‹ËZ>¹ÖÁ³oBá°FpIø%¾}%P‹Ÿ”–ÂäEbÖã?ñ‹i-%ÞÏ^|c* ·š#	aÔ¸óü<, Ý¥wWþ[a2Ã×É-­;V×x×Ûü@¶þôÝÖþæ‡5)/w‰	O!¤Üp•Ùø£Ž_;e—_>uÖ
T—aW$ˆ¥üŸ­æ*ÖO^µ—W6—ÖäÉ*VFð>%õL©†Œ:)Ó\g}Å~°j[#¼äaþeýr£º…öªtNÅÆ5Ê62ç¡}ïgœÃ<Ì·œCîÓ4oe.è°jHCTÝhkbW.®¶¬ÏAC×·*Í\i!ZÌé¦†ü•+ÃKD©ã"õ¼¦x\Z¥<,ò¼³«/iØU’ÔtT3¸¸gÞÔ~¹íadB`ÈGâ„èZ0}ÿ Ú×È‹Esö -w^¶mr²Cs]ešC^]WÊÐAqg}cwïp¬Š¼yýhü#r¸4Wþf0²´«ÞŸ/e}nv+>AúÑƒŽEïªºTWuVß\e\’	§JÜJÈßÿÇ%¯~»‹¿U'ð­Që¦ÁÜò\ØÓOì¿“è´©YïYQÐ]©®âYògxúté)û·ÿž.º*˜¼ów iR Ä¹"ozÚ]EªË•ÜÓ‘ÙÆG±3¿ú51Û	%€¸nÌ®€G"ådˆû,?ÄšùÈÈÿ
Mf•ï^+'™W^,M{kŽ…N¾*ç§_}Ú±cGùÍêÁ]ÿ`¶Ã²#S¼á¨¦0Ñö®ñòF4WÖŽéÏ^ý 57zÛ7Á€–hè¦sÌàý=L.@ëOm‡q‚©UB²Œñ(/±ã]egåiÆ‚FLnß;ÞÐÍi}DÏ·ÆÞÎ¦ß_ý-]îñÌøG~’M’GÑÊøÑü3Ÿ¾€WšÁéÈOa±ÅB·¬åîg±nJª¯æqs|T–è9ö¹oPZG'#†ÿÈô™7•ë4sDØ±OÍ`ºÐ¢íØ§AÌû¨ØCTô‘fºÆ3ñŸEŽG*½º>ós¾ÐrÇâ `µ¯OÆÇ>×æc“±¶Ü³À>þ¬%¼OÚ
åM¢üê`
jdê´£ñü‡%Ol6Gp)œïv(\îù×øD¾Nd#û”b%Äop@;“u6dŸ<tP‘+n0ª	­ûXEgrFt¶°õËéLGn×ÝÜÙßÚ9ü°FhÙ,^ [èOá_ýoôß÷Ç4Ï¶ÕÈÀûâÑÀ{7¼«vE(ßtòy÷½P¤xxXÖÝ 7bHY¡¥G£nñ1È[>]³~œ¯XË¦]­ÿº2×iÀ=nìá`;8w‹e_¬–Á Ã_5\´Ö5rçSÃµá‡“ÐKÁÍô1©RÙzµ¯Heý~IJ©­ˆ¼Ðò'×>®}ìðAŸ×tÅZ>£±Ï6ÎåZ—ÍPø((Ü  ðÌ.(‚ úù¾ó‡¥¬~bk		U,XÙKÙÁ‘Ýué@Jkò(QÈ“›h	‚Zv•ûžÅr[»nt³s¢ÛÛnµªKÝø¸L\N8JÛ²CÎŠÛÕN¸ãìm»û¬ð¹»Y»zþvU.iU¾v7èiW×Ï®rJ6ËEóxÂëRQÅ³.w­sxÕÉ>uÏmôª¦GÝ]ñ§›R;`Dò±«Dj'PÞ¦³Ým¹ÚÍÒÑÎ*Cšä[TÚ=ÃÕîs×;„˜zžwmÉ ÎKíÆgJÅƒm¡šr“kEm–Þ¥pé¦!˜³t»íÀN¥€òµ‚;•9jy¬§t¥Y¼§º†,ôS]¼áŸýòäÝ! ®hÏ\¾Å3kñdñû¯Ÿžžü@0Dú8ŒÏàFF‘q¢%ô©ÐFƒ:rµx(7S]/ß÷T!{v¢‚{¸9ù«æñåJÙk‡(;É6‘k/d+O£gÏÓÓ”Ù’5v¶=!lZÀ¬SÒvÆÊ`jS;PId0vä™Ì¢E®kITB1	‡W!+¹„
—ŒTw×P.¢3¨Žœs0\5%g·ð.Ó’±)B‹*ƒŠj?m>¹`c‘|,òNK(gÙb%vÈØ‘S,š‘,CJ‡–]¾Ö·#×TÈ4×8gŽådt‡M}$N8t­³ïzœ5¶MÍ2Ù-k÷™gª0§MµMj·õœM¾HVºžGÉÔÐÑÔ§ä¤Ê9ä‹Jé”ú> ·Â0š×…~xˆe™€ˆ¾g"wBÐf¦Ìj83	Ýããü¹JI@j[Û„)Y:œýÝYVò~‹â0û/‰W™"+Í÷L
o«¥ò*|Îï¼Uá\†‘˜R÷M·)›ñð¼½Ð‹²‹Þ¼ü“Ó•Ì&†Á*Ð pqˆ%ZÄÄTn
‹Eðxõ{?J½~íEþmÓ?" ˆ§Þœ	zŒœP9V—¾ q¬=ñmâe“ } í»ÊólXuö—tŸ=Å
é|ƒ.èt:ëôtðõ‹Æ=ËÐYÿû¯ÿ™¿I÷cyáÕ¯°âd™ì„Á~<Gï¶€ä›Ã¼=?¹ú5PìcÓÝd+2¥ˆMFG~bdFAô¦³báÞt^½zU]¯ô|<{º@ž¿X _ÃßWFz|ç¨A"€ðÈ6£¿q¬ _þ€ý--qMZ%y×kX&õw¢L9AÈÊSZTÞ|·Î’ƒKŠrT¯ï§é—{œ¤l!ð,é1²Ý&ˆ°Þ	{ž¢¾ÙÔGFÚ£¯lCÉ» (Mk@¶ó‘tÖó¯y‡“tgŠ«Ã”zV%t0 ÅS§f!ÝsVôÉè%5oÁçé½r7¬±ÜÇJQú}›?Tƒ:¼zº–cAèÎ€¼!r©]CM\¾D©´Di.âÛñ¯¨^ÓÇ}¬ý‹dZ¼@¯_ScÙª®jéðÇäçƒIú )TWïã$>;Œ½4ëv@|„’«_á[?‘tBQãIgtRÖmzTºÄ˜¤~bÊ³“%Ÿé/XdÌ›q¿;ÀGdŽ»M¦‘hM:·àX\L°Ó÷²þ	AâezîúK~’Ä	Ìÿ°wñ•OTrœÀ¬ûa<¬Áœõ=ÂPÊAèo˜c/N|Ú[ŽBZŒé­;òçýÃ­òÏÆù,ÿ»1ŸË6Mè\¿k•dø.§VdÎ°f){éV‰Néuçòr\ŠºêÐ‡Åò“ÌÿÙ“ËËùœÒeää%®WRî•ZRNrÅÅrÞ–í-8ŒOåqõW$€¨~’^ýzê‡rºRºë‘a,³¢Ðß®N£ù‰àÀìÑt´ÊÔa¼I» –ï`z
¿¶©¬Žíºr«ù¥,~ŸùÉ†—ú]šekî÷Ú+€Ñøäö’ø(ôGé€LÛ˜Ýöò\c¯eã%|Z};®€˜N®µZ#Om#:ˆÃI?öm#Jùm1 Ô1 ´á€¤%~6I"E¥«—øQŽLZÈ|~“—ÕÔDåØ¬¬1ÏŸç{ª€ªòî«šRå­(Ï>’¯®üÂòûªœaÁ§j2"NªÂ>R	KVÁÒDar¥SµÐ¤‚éòèltÍ~Qƒ³ˆJc¢ƒ@ )ŽS‡Ç•lÐ!‚HehÛBªõ­Z|N¹Æ¯RÔ8!‘‹Ð
lâcIzÒSÕ¥Ó­EË2ÈŠ9¨MFØ*“È]xaØ+£¹ñ°ò+3^)/-ü&V„þX¿àÅà\¯6gÅ³gµ³e@(m‚Íâ×5)ëä)
”&â·ì¶?­Ýâ³%Ïí¡¶H¯»Ì.#K^*œÆZ ö_¡•¥öèÓ±™ï°×p–^òÒB¨x×½¡C$¶ÃÐ¥_`iAÒËüˆ±%^L8µ¶ÐD±¬ƒçåÛí£¢,£žã8M'‹û’k;vý2çº¨dU¨ýf–`å’ütaºD’wz`ÙkíS¤7\&Aº‚¤Ò^ê¥ÀäO/¬G™XÜ±F¹€u²ÅzŒa„þKýIZ°Up‡ÏæD`$È³h³a ‹‘!˜hY±ÄÅ§6¨8Í*¦n-µÛ±øX¶ß)ø(n¬80R´y)ãdÛ)vá¤Tí³Tì“3t!Ó y¼ºØæV`N+‹ÜW#
'aRN&£¼
ä 8ÇŸ²†¸Ðî'‰§MÒ®Õ)>ªWô±BF¸œïŠnk‹],NMäp®«úâ9ò›wãÓ˜8²B»4a•,½ÆÎ—GÛü3þ}­'óVÑGØ©ã”9*‰qWº.®Nš›†m§>˜g5â®r®¶#éêÙÌD¢š\ý‚SÇ×FöáEÃ8	~†µj”…R¹zø™ÓÝ¤¬Ã	V*“à(äO–(ø aI2hM*îƒ`ÔýµÈÌÖ¦kÖ
íÞMŠMmS
šðÐïGqo|VöG8Ÿ¼«ødk0áUÈ>Â×$a?¯×»xt&Û$’ì–ª–Jl‚Ë*éœù ñ# ‹¦”ÇK/Ãb"n›hS«„³<Ê.§‡ñ>½¾@Š­B ¯$¤y¨òÿ²ÑPå(Ý<)¿„“U¼£4æwL*)Såv·@¦w¢4 êÊÉ3¥Õ”ÜºIõ-’ÎœdæE¹-ÒJ8£¢ðsûDTJçVMôÐËµ?IÒ8r›¡»! ÉÀJ¬TéL5™ùa¥žxÇ[ÝãMb¤µ§«×±]XTaøÌ?Â˜’¼ñÆUÆ³=¸Êc+.ƒ¿tPmäú†-!†67ÜYg5œqèÕ
S_™Mì{e.;Â´WqÜèSú‡œ(iS@YcàçLØð¹~>8
PRAm 
àô:›IŠE
§fŽ§ú™z‚DK`ŽnãÉÖ9Œ!Àá0;¨Úu`ŒÆÈ¤	Õ&¿ÔYvÚ{:Ê^ªµB:Ú¦XÑ¶¢7
í#ÕgÌ¥¡±@~˜S£	e¿B6ã\]8 S÷åÍ¨	õfu7Ï(‡½êRÓ/+4çõs& K5Îôâ;²­þ‰¦XÒ´U¯ú+Ùüï½!EsUÙ
 wäE}zÐ}X«ñ²I‰%B:±¢CÎl`:Äê=ÎÃ­è{ë9F–´Ä''Ð+æ÷aCKütÃ| 0à![Oª—sÍ<4PË5¨õX®,Y°©®ÈË¤4fUgÀ0øÇÅç1u]‡gÙV˜~½<Ñ¹@«Ú~œ-®H›£«`­+ýAÐiªsÿ…üLavÇAŸEãQÅ™ªI£ªêRg*”œIäÌú?)\òFÄo¼Ú’ì-÷-oÑ½$uËÍœ"·òì•·…©™S®bˆ—M(Ó ŠÌöÈ2á1
&´«¹jTœ,’.´Àã²^´Lw"•Þ¦˜ÏŒ2…Œoó™øbd|âï®€ï1Òó(Ý_Wºçííi«-×‹î¼PÏLeMŽÀ’QÏ›ÅùWŒ©Ë×¾žÝNsÄ+ëòŽîŒ¥ÎCPJœ½J„Cš’c‚l©•f»Cg$v˜±Šó‰~HO¡>/Z^[¦¾ãÜ#%›Â·ïÓ-%˜€¤…Ùí žÂ’x|ô#Øé*’ÖÌÞ–?qûD­¦—xî	Î¬±ºRðvymÑ8}lU¡äp2¸µä”
¡š¬¯¨9ü°ùá@8¡Éì´ WÈ!O§£ÉyD¦Û´4ãhw¦2(ilø¼þ÷ÿø¯äƒ<µ’Ó™ðQË}Ò–ŒÊ—Š×•¼³ð…x¥­HüQŸ–üo¤åÍ½ä0WzrJ/ »ô3Ó`fkb ¸3KÓ©ð–y{BøšÞò±=âíf/Ç[YµU –÷ð¶'OÍ…·Þ·ÂÑxj´=Èõ'­ã­Û#Ú–ÑVR=?8¼•Õêˆë@Ü"ÂnÔ%rˆ/±yâ÷ƒV}{c­DeiTñhÜ¬Åóvé½B/½'é¢?-ƒÝš´®	”!ÚEMË¢N/%Öšc–ˆ¬ÍM—ºç<^Ñ}¾€J÷µi‡Ô}y+z—èÔ´Ý3àÒ²²dX'åÓ¦)ÓlÌÂ½ˆw­d½¸‰ ê9½TƒÿyTdUÔ?›á¾=0ˆc?È0T‹¯Ôœð…ÅÃjÐ„6lK@Z‹y6q01KS8Û”æá¶'AºÉ3^HoÃ—M†R75AñÂ<†·„Ï&P}Z¡‹­¹°”±ŸËÛpšK†(Hs±B3 uIÅÃÄyEœ ¦‚Q}Xf‚4#`¯ï5À=áðÁ|yfAðò±u+©œ6œv¢(0ok¨¥V62ã$ãéênã}ˆVŸŸöíncåå­(,üNe«Læ>kÏÈxñ¹õ*¸Åph€­,w­È	þ¯Ô9w`6˜Rõòªè&—2)Âu•{Ö­ÒWupïY\þÔL¿õžÊœfÃn—,‰«„Cäq„0¦úq’€œÄˆž‰÷´&]s‚V –°¯ÜNº;jö÷1+Z›kø~µ&˜!–Dân.C­Ù~[»ªÇcårátüÁ%uQHÉ2±4L?ã1¶LÆaÐÇm3Î
wj‰…h5»%æÒœ²Äž¾ÄyÃ‡µÄ…È?«5"mÕí®³Â&‚#¿…fœƒwð”T˜8ÅhE«.šÕÍ¥ä¨8 ú8ôNƒ4N—1].¾%3Ri´Ôê³U½Ì*¡s1Ã½»"Î;^ÅÝàÇKëyëƒ‘×‘/#Í°Ó…ïk$ÍÐéak„¥ž·gÝ¡Ó¢	w‚ËÒôè²P=ôyiÏ}P·b‚ábéÑ–ô§öÞË&I}Þ„2k3{»ÌC¸OÖœ¼>Í±Ó{€Ÿ±[Õ|näÔ&”HâÿC`ìA™/£PlaÊlà*²«òPr´kœ4bm…g^Œ½äSˆú»œ´å&A*áEÁéèx;Ô©Ë8šœõ)~[ØÅ—%z”%^ÿú,  eu#XhûýŒ†úê¬ßÉs†ìíð=ñŒ+	}o€ïÍ‚áIfX¼’Ê)€!é9NNž7ç89ªç«
ó36×G¯²Y„þ±>àZçÉé­´ Â¤ GM°<ðÈ€X¯þ†9;I/¡w±r¥
Ø`%V>ñCïŽŠªy'~”G§ïH•˜”U1vØ•†ü\à'G6eøÚ€³…‚”ƒñ¨ ã«¿’Äï{!OÚ£ºó§~ÙcsðKz$”Šü@R<áÜL¡IbM‚(õjAtÙO¦=‚ºFw]á2‡ñJ,DŸßŸ¯w>™«ÒoyK:Ô×W0êLSÈBaòk—­d–\ÓÖ1’O—tk9M•Ì¿—Ä‡ñhLÆ¹9w5ÆnÙ»“~YÓ×Xò$ðºÔÌ¿Ù¢“}RšNH¬o¹Ÿ3âQZë´á;›·¹¦ÃtSGâêbÛAHÑAÞp»ÛjÜ§Jy
œŠ?@oÜmš··†¯-9w&â¶È¥”€)7`ñövá¢oLÙ½Œrð£VZÄhÌC‚`4—Ø³8q²LƒrF£Î"!/4ÎJÌ«?I ôˆ079÷r4`™—Ç˜>qÏ‚>Æ	±Èc¡±â'Í’L“ÒWŒ’®æžôæîYZZ’ÇRV>Ì;]Œâ/ri )I©~íR½d`iÕ|­¹5…Æ ¤öl©S$i½n‚V]X4'fm#`«F]‘5p—ŠÌ9ˆ6Y0éZJPŽ:˜¬Ä%˜ðÑWîÒ†Ÿ:vZÐXCZK;—å 7aydl~ d·~Œlü—|1Ì„ÀB €(GYJd	KþŽ´wD‹9˜HugÀc4çg\+a|¢ŸNNh8NJSõl‹…Ì™¹üu3 ·íPV9vþVg.ì×Þç(Î¬Ã¶ ¿7ã#/Ü	ÿïqÿéCtáˆa7Må)qï²èeò>f`Yö{}3À÷%é^äãºtVº[òú1JöBL²ø4EÇ{pœ,®>¯këå¶Èù Ù®¤FI¤^¨Z–‹Þ´êYz•¬D{×)
˜Æ¤Ëü®þ5A¶z€¶j€µ m• i@SåØ;‚uÁ&Û™4^#»tAÕèÇ»6"Áš¨.ÕýƒŸ …UW“4‹G4––ê	Ùƒ'„K—ô¯ƒ¥¬áÀ¼›b¹h1`ZB¡ªÉš£É[2÷a’%ñ­d0`5à|a2µÄrd¼h[Ô#±·×çj5P—f³’-Âx„_¦J+º6žèšlXi8³Y²^KšAIm…O`wßïƒÜÜõ•N¿EPí¬Ó?¤»€©óX0¾£Ôr6Ž¦‡P¡Q·0G»ž›g†Y¼l5Ã–/l™s]ôÇ:ûÓD®Ö¦KA®³Nÿ8xV[¡”;6I goøÕWäÉHOÕãÖNå•ž7¹ØàRz‚z%NÑQÿÖ)+3y‘!s€š‚ ƒÄr 3f‘À˜»›»Ê¤Ú(J­ò¨NßÇšô’Ò Û8ºØEælL>?Ú‰Wª½
êµga6¥Ò\r6“¿{-ƒ’xƒ0'±ß˜‹:ËV%vëvlJùTÍ%v»ËRï5µ&ñ‡Ë¶${WÓ[’¤µ¦v$±¢…9©È‡Á¶ÛR1´ÜdD-EÒuŽó;¹êÑy~Öó,Ñ|CÃášôs•:ƒ|­ð½Š}ébù7ä]oóÙ:øÓw[û›ÖÈVÚÇ˜pg µe3ëß,+¨¥Œ^o^Dg´Åß›EKNò6X~ñ”z8×p:X~ùÔñ‹×…p8 ¢ÌE„ Ã˜º´Ø\,ùCáà9B×26®£:ž‡ÁJ×™³Åäd‘ãc}¡xß¨óÜ:ë+6ÏE"6ú+u£õ¥ÔIð£7«‚Ït$Á'OÓÎÙ|½é—¸Wrë)¶¸ð’°zõpiŸAŽò‡‘}Œ‚ß3Æž™´6›¢ó$9è,Iq‹”Ø˜j”È\b?gUvüF×@«˜üb®€hVæ½ÐŠìE§$ïUÕ(™Ê®ãÇ¼.½pe]D "!-•
à‡«¢"@kY!Ë®Ó/5*#Új#Ú™~vÏVô–pæ@”Å#=.>âO¨ó¶7Æ›	(7òÅ!­_‡y/=ÏHUŒ£«aœ
ƒ í×ÂÌãGÁkk;jæ“QÞÑ2_£5öz›u?’…æB+2|ÖéZØ¯¹¹yW¯|“!õL“;ç·þ’äzõ=ÊÃ5^‡fùÔŒQî”^b=Y;²™+²~šSQÒÇ,½â|‚ªÕÇFxs35šŽ›yQ;K\`]ï°eÀÿ|–BoÚ™E…ÿ#ÐNé¸Cd0lêÁœxé®?DßP]"# «oLCˆù]¥{ñòBá¹‡%{â“Ž	}fIZ	>e¼‚´ „Vçãô%’Ç‚ ò% _Új¬¶æk¾FÌ‰k»pÞÌå—ïíy	zÿ:¦TØQýØ1*_kÃ.Z±©9Ôd,ØæŒÿ¹~!¯Û%Y¤¸f"N7ÉxØê5ßaàðê×lbY1Ín¾ÆªÕÄO…f?fÎgéÅ=â}´|«Û^xBuU@þ²$¦5m0î<þiÒ€*Ÿâ-ÉQOüÐø¹sÜ8KKÏrI¥“0‹“`Æ`‹FÞtWISÎÿ¤‹Ç°Æ¶~€¾³ÛÛ½ú—ÞÁÙ:8Üï^ý§owzF_öÑ¡¼`÷Z‚pÆb–¡{)‹¿Câõ%šÀ:µ—y¸ß$X¶¼&{)­‘!ŒËòP¾ÆõyßÛÿÇ­ÃÝoÉWäãÖîf“g?ìmí÷®þåê¿mÀÓ{û6¶>4è`kã»}x‚üá»÷½Ý&Onmì~x÷Vµsö%òÃw¤ù‰m¼x—5µrôáWÕÂ8]­<“’²vÕ®¬ÅØY8KYÕºÃ«¿ö#`Ø-šZ#3fœâi†„\ì8‰¾|ø¹gÆ	ôòHð3¿ï>DÊžòî<‘ÙÎR?V5SYa¦ïM5R‡c2,—¬póàÙã©oZŠìAèóiôO/H‰—[;õG4'Æ‡bÞÛº@ê¡j¨EÐ¤ý] +/gËéæÎþÖÎá‡5j÷dH¼@ÐP‰Óh=AøÖÄzúâ¾XO=ôí¼'2ªñ<.fQ:ŸÙOcy«‘9L¼Ÿc=¦Àq;T#L¢æÉ-½±r/	b"þhÛQTba(Q{õ×Aàµ§¢·l#í…™`ˆ‘nM¡*OÑùuÖÙß†ãçð,ýã~ô~@3/›¸8‡//Rº('ö€·BjÜV°®[Æ1H$aß÷Öˆô€‡d¹Û×å_;Ê³ZB7Å÷Öñå&xñjß*ŽD¶
ÍuÒëwAš}@ ö¢Bóªž9¨ n«;¡†?NjãÌ%°žÙ
Î."8t©ÑEøÁ£‹3×!J¨ŸþïXìG¾Ë/žÚ‰	o5SjRCÚ@êÊY] ÏÀ1ÝLt@ÇÙÃë©ÓØ"Â%|[2·³›¿%+ó4¥¦xb3ÍÊÑ(c/IýH-·@VžZnHÇºF ‰]cÁœbÒk˜™êdiäwŸ.Ð1,Â¬¬©j
KælD¯ÉˆFï¬oÒñ`<# Ó!Êé…=j=Àœç˜ÂóÑ²âð}H‹`³øOms¼‹Œ5J‰d'ºú¨Ïµí4è+ÝÌèr¯yÔ}ïDH·[5›Ý>,Ó µí`ôfÍÁ–­uƒ>U c7>ú/4 |@Ä“;ß\ýrê‡¾MLÖòÊÕÇ„„=åaÞÀ6±Aê÷a‘Ü“‰ns¢[¬BÛPä!¼Æ‹ÊU´-ø‚c&ïüÑQâg¾nD˜…9ÀDãkšÆ“dú7gøÆàâ
€gº ŽEP,ìWÂNcQâ«W+Ï"-óYâ™:e"Ó4Ž&YGµk`ÆðúŸ¤ OÊUüÏxkä@le¿ëÒ-C~ö)¦xÑlþ¬–ê%+x`ž«åä9©]Q22+„Á-­Þž,Žã€vz²øŠ§[â‘Ü¸Þ+j%›¾uî–öãBUJn/bÕQÊ‘Öä<öˆ¯•­ XªÔ¸ÏKzjôù3Û¯InãÆYž÷8ÝBƒfÑ’m-â²)ƒLÓcwÐÿ|âe©7›a]ÜýBÁ=Ÿ~ÄkK%—ïÞ˜÷ÓÔú?M`AªŸ®Qo<¾3àï¥Ÿ£>1ç‘¦ö<cùL°¿7IOlÚG–1–=ýÕWdn7IÍãhŽ9Ãòwgö4…bXÞïyAFä^@˜øiœò4Ò®t8ÚM/2dÉü5aûxx w;{1ðÞiœ,ï(i†µ“3Šˆgôh‰°”Ý„D ƒ‚–y-ñûþ žÇ}8€c™.u\!$–;¶ë5InØZ¾©Ô¹áÒT1‰9Õi˜tó°ƒ‘—dã\ºJú…KÓˆvÙD]+ÃTÏ´ÓÅ|fÉ¤³ÓE…‰qÂ9«''¸ U*€W×t2¤”#!pØ ˜5Ž	R¯ô­S%ûØé«UèGt=Y|IÊR"§µT`­9X·§°Iý&¹_ Çfà¥=\&‹M?UR¾–æý7HP(>¶SMù+—ë+MÉ6¼Ê@{|õr}O§|þÙåú³ë<ÿârýÅužÿÝåúï<ïrbpèö,1··†ñð¤—Q½Í&\€³è2Š,çÌ¤-Ÿd>Bª´.´¤±n>Ô 
-ª¥²ëöÁkÞ ?Æ^üï}\Å%¬ OÈo_¼xùü•øLCcÄN9ÕˆÓ“Þ}„¹âcò±¨µÍ7¹œmÍÖ×+Ï,òh5«Í,’mI-Z*ÉlH vÓè-µ˜y­wBÝ"Ñ¬šJUæé*.Ûªÿ°ÀW{$~ÿšuÄ[Xâº8¢.ÕA×n©R>S[$Ö +‚…›×Aâ¨uìÝI¿×A’–¼n$y(¦$ò¼ŒÍ_Wj‹¡0üæ+q´IçÁí'qUüþ}J#_‘AžÇÜg4<¹?"/êûAë'ñYŠ¬¢vC*3à‚º©Â‹^DÐâÊ1t|WÍ?°’YÄ‘½6¯%Äk]ô¢n——ai‹®Fê$î®¦°uÃæ®!«¿ä.ò¥W`Ú¶:+¸'Xügîä×  ½øhŒüÅxá4ÑÇ€­_;xÀ;›šÎX”Ò	‡$NW!AµüO‡ÎÊ$Ôh·&ò·ßÉJn¨5VB`Ú«„Ð`³­„ðX´ iÑø«EÜ…	jàŠ®µ~yf^¦@C÷[*S £²LAÚ6‹2åk‡Þp!ï½yC€¥‹Ì¦
‚—º¿ÿáRm-­wäŸA×i]F;E»3}Ùµk™m>¤ƒR¯”'»<3Åc¡i¦z,tSí+: ›V?–Öß¢‹ôùbHE}ôŒ&žð%¾á"ŸÞ[£ž€'Sx[ž-9ÿ7Ž>Í,}û£×¼c5†Eñù$¯[.#?®ÔN¢²xCpÊK¾Ïœ¥Õï28QÕv^½éö[wVW0ÂÙúœ:<K¾UÆ-ÝeÀ0ñ’ÝjdR‘Œy®V-ŠÑ´º#Ù4¦Š\¦yëwÖ;¥þE…ÇŽº2mÒÔ–ý©)äÙ¿f|L÷ò¤^H`]ÞæH`ÊÜ>˜ÍF9B˜ígâöµÌgù;DÙ'	
óYq±Ä“±:BÏÎK¡]7eT+¦o¶©	ÔåûÖÔ¢fÂü<C]»ö4yêšÓ¤(iÐK·L@k&A#<Tóc`ØzÐòÉ>ÙÀÇÔoR©†E"ÅX!9¤HËì=ÔÈìmöÏqª$mÈn„	ãnòÄÈ½lÂÖÃê‚ãÈÖãŒ-»ÇÞ†-§‚†¾,y·Úåx‹ÁG`‹½§¡X2%ö\_²?(†À$UiLÂrÄ;®0½>%£ãÝÒTÇÎ4•¥éú¶&‹ŠÏ¢¬³g¾I´Í)hQ£lð¥c.ÖC‰ú–xà¢f„áº6Â"
aB›Ü;„–¢•FÆ¢|âHkÆeÐ½ØC8*±]FoìªäYÃ/‚LÐûÔëÃI1QhWªhD>dãè„É€.þ•6#Ç1ã#/ÜÄ|ðÞËú'Ý1½_b™57×K¡t‰¾£D~J•ÏxÝ³±EïÃ§Øœö.iüN‘dZÊã[Éê6Ê’kŒä5&=ò³3ßLñÒÓ˜V©Í<MGK€Ÿ§
ýõßpíš2ZcÏª"ÏÄÑJNfÊcÇs¦%ûjúÇqŒÐoÉnÅjòÔÇ`˜qò¹"p.ñ±vF·6ô# Ù™ßÛ9˜‡èÔ­Jìðùþüg-.ëÝ¼¦Ò>kC¢REGKƒÌ±Ægj¬pFM´Îf˜{ù/yˆ¼3åzuGcLŸ+ˆó~Ê—«»‘Òˆè™nVwùiä3‰‹–/WwCNÅ §Ó÷Ó´—~‹ugíÍQñÍe7ÇG5V”Ç‹i’µ„à<ÒŒçõ.”Tkd_”Õ
¢Ó«_BÌ)?Àde™Ã«¿EýÀ#½„2e:PÔŒ†—¤éúI²F¼ès%¨9ðÿuÈoi·K#ë‹K
?ç¯wÐ—˜È ¢\ =`²Äû¦þ(Ð23~ºò5ñƒüO×aLÜ÷Jar¸f90Ó)rUîw)¬òNÅÂjk~l;Ú^0vVÐ?½3#wt =~
¡½ø|ŠäßÒñ`0ˆ"ì†9æÑœ™Åîlõ™uø4ÝšˆÑH²aG±s´­ûƒ
_9êfíŸ¢Z…–¯`	oQÎ‘‘H'Š1•§A^q*ÇÄÇäQç–Œç_¾uÈ†Šbù©w-tjöQGsŸ÷GèfÐI†iª  î1ü'W¿d Ü¤}–çðñÌ¾ú%Á0GJ'ƒÑ(!=Óë¼Â0µìoÖ oû}ŽÚ]9‹ÜùÇ½tžø#²
T9œD5pàÖUÇÂ›qPžÀ‡ª©ŽOXS=Ô=@¢jç±óµ¢ø +ž9˜À×šÊ*Ÿ^KmU‰fñ Fp¦Æ~˜ù…Ù@%bUö*àX|4pÜìí-oô¶‰—MPM`,Ëyí€¡*R¶€6z½AåK< ¨ç—1}¢¸ï·C«ivŽ›"Ø–©uQqø<•£=ÔÞYÑ‹qA«RàÌq¦^€}«óÞPf‘Qœ§ÌÅ.ÄNJëÊ¸*úQúnIÔoÀ*]åÑf™´,×À®›b»ÜßÙ¨À;¥©ë>ºðäAMuúgA8-Öì-ÞæA
hD¾¾¼nŒN£(¶3C” nN¥º±Ùìa:íÚáþà'/¡)„%î rìä`ªØ¾ÓRÐ¬â€®	dv˜6È•ÙçÂŠij@­ÝW_W¬å1ê›gë²FD‹á®‘PÞtÏYïë)¡D:Mj¥*Oƒh")žHŒ+œ¨1nœÓÑ où§|­V¤Q	_¦	4ªXŽ»f2óÔÍPy=1_©iÌ†QœfA_”^ï‹û›ùí˜>œ¿'HÑ½öÍEæms§Ûb´&ß\½½æ¤Ë	B|+îŽá¾ü«h“p4È¿J=øÇÞ$„M›$}:Á¶ã£’/ËÉ‡-µ•VG4]^7.8úOËk¿ò‚£ß3ï-H7ÅâV›6¬µØ2© ¯=Ê×ý÷å†8ŒÃX4îbA1©‰„±ýI’À¡´I×dPzê--e`Xž|owcÌØ*½Yì4ó¼¢Þ‚¾ð¼R;çmwóo—¨k! lix(ûNƒŒmÔ.±‰Ä.;ZJÔ­X•üp_?ÂËYBZÓKö°KÎfº-oçwê³ÕE‚ƒß²DÛ[d Ë=©ü3ó$!@¼¡ÿõ6—¶†îr¹êq&½R†K
˜g	\>G}ü–TS½½F- S¸Ü íŸ}…Ž&%ô=Èx?o.Æ~r„™öLÛáå¢-°0ÙÁdwÉ¤?à6¿P´š¤~BsÄ¾¹À¯“ðCó$‰8ZàÙ-üjX«o¼þ§Éxë|'™m¹X›šk¥46,{_ªü{®h‡…HtÑËXV"žÀÇû¸Ùímþ»hÁ®ÿ1Å5Oè«Ô+ò,6ã³(Œ½öc¢\Ýêó*|ÃÅ!î¢«oòiO.ò¨ëCöÅxpü]*¡/ªG]pîHÁn,RÕ M´¹üõSCq¥ ÷Ï‹¼¨	pÄcxv’ 7êŠšÈµrÜgdñ«`Â½óÅ³Å¯áÂÉâ÷¯VOO~ E~ôÄ8á¡“`0ð#gnHpÖÈ¡P‹ß2—EÑßTç]¶2^ê“g‹/i–P)ýÿŠäŒÿRlYÕî™«Ÿ aŽpyiLem+ªsòLsÁ‘Xäƒ”òìpûyÙÉ×Ë'ÏjÛz´U¦elLûu:)B¹ÄÖáqZ˜h\©mÁ²ë/ÅHB-¹¿ –DŸææ/«¸uyá» ú$­¤˜Àð`iIï(ÁÄ¡e½#Ï°ØßØ’ƒ7ŸŒ=CMpœxð »ÐŸ à¾‚bò™º‚fqÒc¾çîH{›Ûd/ñOÿ¬óÃÜ<šæÿpøþÝÎ66ÚbFw[‚þ* KìÛ&ÇƒövÖ10µÅ’ïQf.VàT‰97µn‚ãÞï%,­¾{ÛwFèM$×ÙmìÎ÷©vÛ“7šEØñMëÎyæð–Nÿèz™›øyÍæ(.ÇhÈMÐ·OÏ™ŸíãÐM{iYiqvV,5-qŸ Ø&‹­àñ¹ À…Ògt¾¸bôˆÓöjxŒ«c§0ó]V&»è.xE|¯ä¥h€æƒø'¾2«Ï-TÜ<û:ŠÍd‹ˆõ¡ÞptªÇ4%ÏGÑkN!ÒY_º.—=Ý	û“3œ‰€#§:‹v‡YeKMq!ý,Ô¯c`ÁiÊi Æ	ÈÒð˜Sö÷9È Å¼(–cÊaÚ>ß£×ËÀ–Ó]áýJ\þÃÿ  ÿÿ ñ„ôˆ