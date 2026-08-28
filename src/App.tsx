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
        const directResp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.7-flash:generateContent?key=${directKey}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ contents: [{ parts: [{ text: "OK" }] }] })
        });
        if (!directResp.ok) {
          const errBody = await directResp.json().catch(() => ({}));
          throw new Error(errBody.error?.message || `Erro HTTP ${directResp.status}`);
        }
        resData = { ok: true, model: "gemini-3.7-flash", message: "ConexÃ£o direta com a Google Gemini API validada com sucesso!" };
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
            <p className="text-sm font-bold text-slate-800 mt-1 font-mono">gemini-3.7-flash</p>
            <p className="text-[11px] text-slate-500 mt-1">Fallback: gemini-3.1-flash-lite</p>
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
            body: JSON.stringify({ model: model || "gemini-3.7-flash", contents, config })
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

        const GEMINI_PRIMARY = "gemini-3.7-flash";
        const GEMINI_CHAIN = [
          "gemini-3.7-flash",
          "gemini-3.1-flash-lite",
          "gemini-flash-latest"
        ];

        function normalizeClientModel(m?: string): string {
          if (!m) return GEMINI_PRIMARY;
          if (m.includes("2.5") || m.includes("2.0") || m.includes("1.5") || m.includes("1.0") || m.includes("3.6")) {
            return "gemini-3.7-flash";
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
  const [isSmartSyncModalOpen, setIsSmartSyncModalOpen] = useState(true);
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
          summary.uploadedToCloxœì½MwÜF² º÷¯€ø<®ªk²øaËmÓ–8”Huó¶>Ø"Õ÷Ìh4X IXU@@‘¢Ù\ôy‹wf1«9o?ýfqW½êsþØ‹ˆüŠü 
ER²|¯y|¬9K¾üòûÏ"ýWÍ&“¸¼&igãj˜N¦eZÅÕp6q’:Å³<ïª2ÓYuÚ£x<0…®¢t\¥QvõïÐ·è‹/¢Ñ¸˜%ƒè2Ôrqž‹¦•Åä!ìˆ¢©9IÑ~É6GE^ÕQý˜ŠÜ‹Òwuê½:f“´ªãÉÔë¿®E]iªå".²¥û²²LÅõèôO³t–R¯b?)Fýäh9ê)²ô–£,,G—Ñp8¤f–ñ¡¹ÉÒr/ÙŒfUZg|œM“¸N“íz3ÊÓóhúƒa]ì<;¨Ë,?é¢+U3&²î^·Væ¢jnùÀÀ…zá`ÃÆRÒü¾2—~ÖTª3ÖMüÖŒ¸$ºêBñ0ŽY’æuÕŸ•OÒòäÖˆõ™ûKý[¥µ‚Ö·`ëúï .Êø$Bqâ¿½}Í¸ïŸž=VÄ9Ùñ…€Ih««ÑÊÊJô`\ä£h'‹Oò÷ÿ¨ ÏE…_d11™¨SX¤:Èã)Lªø<Îêè$­wŠQÕÏªíd’åÑVôÓ,-/ú£b<NGuVäbn$¼°{ƒA´Ù©Ürt~š–i¿''ô­wïü_ÍÁ@ÓÆEó	a‰ó	~ý ¨±,:)àßï›ºVç†ÇE¹Nû ;î3æ°`#ùûÉ§ï%L÷ÍHü†AO†ÀÛ1NÙ(®xƒº¹+3kÁ…±æ´²q¶à1LC®º¼à]#4ªø,¥‚Ô,,n;±¹Í'Ãß(8˜4ŸÆe•¢\&>¥'»Ž+”·Ë2†ùXÑ¿}QxHüX¸çlâE#”çÑå•=ñ˜@ì% )?†•±Ãr  _Â(³×v›oÓ‹ª?+°B6óÈB¯\ö¥©ÊF»Úäcÿòàóò•fèyÔ³$*ŽâœîJ ÕE·¡‹ŸmaÆ˜P»ÆÆ½Ÿñu’Í3(lwK¶øA¯ïw|ëŠjq_U;­¥²a˜„q~1ê/Ñ_þÒu5ë•Xv‹ê‘5/[´8>Î¿&g¡ÚªÍùÈþÐèùïÓÒê†kÖìüAü´»N<ØÞ®å]pmš‡Ø÷Úè¤ñ9ÜÒú,˜žæ÷Pj~ÏÓj
…â­?_Oë+`TåÆÙ$«ûw×ÖÖZô?«Æµ”?Â6¬F©~ºšŸî`«æ§ wÕüTksÔ>û8ƒG\zãd{<~¬Þ%úN/M­Î½´Ò½,µÆdõ­¤þ”†ºzP¨é*TXn!…JnS¨ÜBa…J“rS§‹*%Q¾©*¥qì¢Jù…oI•2ëW£G¶kQ¢ôØþB”A²U}rÐüèNŒ“ÿW¦Å8sÑ†îz óªÑ[^ÍQÕxcšÇÅóBj[˜×?-3w³Âæ°ÍbÚÚ‡ž-ä¼žæ6Ÿ÷[ºÑIg»©ÿ'|8½ýÃþM4ökÐ«Y]ÏR_-³›Ñ È–¢R+ZÊùÇµˆ´n3<ÎÆuZ
«ä¶ÑÈî{÷¢ èaÆ¼[WŽŽ¿#uüÃ¸La	ßÇyÖóãêíõôüZ =Eàmî]§à5U|B4¬üŠnR'·G±çãÕ]lÕôü®š¾Óè…ÿZØKÂÑ'{­µpãš°¬;øz9æ:{Qx{ …Ü½ºÒõü½]{ßÉß‹ 2Ot›mæ‰[(lžP©Mwð»X(ë›Z(Í.Š_ø–,›Ë~5VŠD»ÍF!’ýBö‰B¯Õ:aþ
,“›*…ÍŒöi)††³šÕB6rŸ€J8ÙþàÊà‚TjóÛ²…¸o€wrØºŒòØ
hž«vWªqOÒºHŠqq’Å«²£qVÔé(Žúûe:Éª*®–£ý²8§üyPŒgïÿõý¿¥ðûý_Ë”Þ¥' Y«†«VÐ‚ ŠX£ˆopÕ?'éf¤ôêº7QØ¾|µŒ”µ¤ðY<¦—\Ï
œ1• ÆÓa`+]¶~b­:×QG%ºÍj¨*°ˆúÉð
ënÐg¦¯QqÕ[ÙuÔ,ËA"×í®ë!Òg4¦÷	µ/ Y\ ég^LRõ°Lÿ¢ zeGU©MC
+F¦¡	MhÉb¶2D3Û(Do¥>„x5I²F…ˆ"¥Ú¼„ŠŒPj/ëÚjÑxKw¤~Èõá­MU©7	ÉÅ¾})“Œ¼,ÚT«X¿¾˜¦@1Œ½âèGàýÌ–1LˆË‡y3B¡wkë[GÕ‰TKÆ‘bþ	N”Ó¶¥A‘vâÓ°³ö4ÃæuÍÆqÔŒã<ç.hvKb9K™7,WŸ…aÝp|©á®^F~{K–aé‹ÂÆê§¾‚XŒÓáy\æý7ÛgYUDqUY>*‹<û9.£Ï/%?_m¾YŽÒ§'ŽXì«ß›ªUdøÔ¬ˆÐ¶>ÊÅØ”Õ5•ë)Õk+ÕO¨*;(Ðb6*R„©~Èý !ê‚s Æ¸º´äh[¬ó¸Ç¬~«ã£×¢Ø<ä”Š@ð´Â a²g×T±T—G¨Òd?§ Ã`É$£õ—åYšÀç2EîJ…”K+Y“É¼c	Á¬nLîˆ¨2á@ëíˆ3Ÿè¤)k@i? tìq\Õ¨bˆ‰Ø—ržbµ^#ý^Ëò¾Z§ iH þ]ÌF£´ªŠYž(|ÌÔ)K©C©N¨Ù_Š²¿´ÿDÉ¬ŒóHË&({ÿ§ˆ2x=ÎNRü
Ý†Oggéds	fUiÜ-Èn€ù.Á„[C°¹*è®UKÏFEY¦³h6ÁJÞdep‡Ñ!5•g12A:\Ò=¢áõ	Wa«@’¿ýãD£*ÿ™šãÒ2,³““´Ü‘b÷A<z;›>ÊÆ),¯ý«"+’QLáWU™¢É•Lýh\I]êüì¿4_áRŠËèfÔ‹§Ó1dÔW±Ä÷£StDÕ÷fõñÊ·=&uØY‰zÄ‹ç‡#˜luúŒÖ_xîcƒNá8èƒ|†Ä“•vAŠÀÌk6Y©ìðŒ¨ Í8 §íÐ?šÁ’ÝSkUÏPÅ…¤Š 4UÄY_ ×%Y…<Åzy‘§=UHã|T$C Sš'O³qÒÕÝöF@Gš}Ìœ;¿˜Õýþ“ —N	e
2,¯ŸIê®¢^!Èøâ,â¤øMýá€•éYñ–Ð¸iiÓKÒr´!-ªX™Ö³2üg©?«ý	MËá²_¤GfÄøXÝ¯!7./Ê‡n¶3ër/ú2JóæÅó½‡Åd
Ã	\fßóý-Â .‹Jä¼ÏÝµ«vdÖØÕgXõ¾…cytA®t*ÀRœ©6O¸ËD^ÎJ`½­â«¿’ÄfiTÿçY¨èBKß¸	ˆqÝ#mZO/ÃCm›êhpVÚ¡àø,b“®OÅq¥ºcoÐû²c«á8ÍOêS2…Ö²«–*1Në'ÐÚ÷ðÏV-|óå½hãîšKåZè(Pí¼„¤u€Iæ»ÎOgù[(% VÀOi?[F°Õçî}4ÔQŸÁºÅ¨çØ?iæÑãfé°í:Ëg)—u6Y§Ò–s} ¯ÙoÙIho69Y†Y(nD
 ×NÍ[å‰Ž"z,=3Î³ô½GE©µCaåH·0–€Â¬WBW>!¾À~dÞ^
[i“æŽeÉ/´C…Ð¹~Hÿ—y\ŸGBSÎz¡˜í>²šÑ­2ÒñâüÌ˜lS·Áƒ;u;V¨l·¶xï|Ó¦ÞÑ×š=þn­éòv(§ÓßÌÔMY;œÝZãUüÍåïÃƒö°LÁÄeq@þqÁÁäUÙÀ¾±÷RÚÕ˜è7›gfnÏ3iY;Ò.eíÈ7Û1¶®gåêv”µª›Q/:¶blážkøš™!¬Z3)Äs×ù mbËfÐµ}ËZÐï:·Â­dË8¶Zc+OâÝ3ÑJZóöuz&ýôÍít©]µ^dM5Ú>3;¦±S`ðõëX¿o—2] ü±Þx¨LbXŸ³$NÒ 8ó±#Z‹Cƒï¢\:ódœî¾›¥ØBZ‘Ñg˜ÊBV¯øþ¤Hâñ3P=û¸&ÉîÓw
eQP}÷Œ×ÕÕh}ý>®OÓ’œ)fe®"ô¬ƒý¿2ê€ø­…óEltâÑi
t	ègÐ8¯ÊÁzPJ8¨³¢›‘ü%Üï°+©öê{¯ÒœSk¸µÁJ ðµ†h Âe~ È6ŸöÖ0çÙö²OŒÔ^ƒ/šŽì¥GR^! O;07#õSƒÐr?XOÊêM½1ÈêÉozÊ¹I›ˆ|”`‘TŠf9Þâ)XÎYUV¿yÅˆÛ†Óóê<-%‡n–LÎ˜Úò­P¡mY­ãáŠI<5§aºŸ­­Øç#.I¬•¸û3° [g%¾gziÃA‹E Ò‰	é¼Ä }-.Ñ!Q¼=.,¼¬(yÎàtãU¸Ð˜ö S|l Ø.C¬aô(Å/Ç0§2â©LGE™T O£ýíçÛï>ŽÎ3Xýc²JG5T»»RA)°µð€sœžÌâ2aÆ*î®ù$‚’Þ½jüp¸|¿?¥¹
S:ýáðþ²‚ü;FãÚ²ñE‡^ÖõHHÃ^–1XŸ/-—ÙtÙzD~7àú`’Ál?Ki)ó|òc_a3`¨î+nx±½OA€Â]z´’åÆ>«Ûrj§/ÞxjŠwØô÷PÄ•ËÁ±ŠNÎƒX.œýÀÃ%Èâ˜Îé‘ªÎ‹Øó–ëùâ8»1ÇP]<úØCWˆ¼Å±uEetÝ:×ãÜjtÎu´ônÌëTº¾ |/ŽnƒÆÞé†ª×AýÅqoðtÁ½¡êb¸‡c¯ìE¢Êã)h°ËÞKÔòý·¨žøo¥zïkwx5òß‚&î¿DÝ™¯7:bL­|°B9[|û*®L­Dƒ!i	r™Ëgãñ@8å—åÃ\HzÁ¸9(-ÑoŠËÚ›Cc¢ð6&åÔÍA)rsHzFÏåè:Žæ*§xÉßhõÕîþ¥Ö‡ÂA~Òˆtý"Çh×¦CŠ¸
`d'¤d'¤§¶Â(ÛŠ¬n‡C0²}¼œ4( ¹;K1
ÐÞšPå&#>´Ñçà"é”8¾V¾‹(jSj!EQBÎ'©_»sò¨¹xÚ´U-u'.ŠmM\|Wõ{çÆ´y¦êu²&HŒçã¯N¸©.Èç6¼Þ.z‹“Ÿ»¬ìsIrÓ‰$EH…£OKàGÐæ¢ks‹j­;·H"¶xÓ:l¢³ZÔ{`Ñó~kDô´Pi`¢µ¡d{ %NTé1…E#¿Ûð	y%VXó‘B…Ic…m³—ÁÖ‰¹.&Šµ¯¦ø|DÑð8ñùˆú2
MÂÑŸ%¼v÷5Šaž(Ö¡5QTƒEÿR~&r$í–ƒ‘&ú‡‚è¢,ÇÈ•"ÿ÷³tŒQ“‚ü¡QRDG´#€ûè“i³^aüp
L%tšeUT¥'³²¨0¸C7p€Ù'Ew¹Ç“õÕß^‡€¼8T” rÕäËkŠ–)Ý£´¿úrsøjX§·ÒÈxÐ¢Ö¿sƒèÀVÙi•½ÃP7xØ³1ØQdÆÐÃ°šŽ³ºßûÏ½ÁËµW¬ÿ¯ü¼½ò_×V¾{½"Z{ÝD_â?`Mõzv[ „É˜"hì¾œë¢Ìâ×‚ð¯?¿4ø\}~É;}5Äà©7ÎA_Q(.T3òñt
Í?4 ­D»žoïöŒFz––ØyPxc¸ÆÞ‹ñjÎ7¥éM‹ñˆl›Aÿò—¨·±¨…ð÷¸^Lq<¢ôLD† ¾ËK¸–î•ù8IëXÄqè†Ùar”W±½ºì–·7—:ZC=–ùÇaõ¦ö.T`™o¨çl<Ö¼æ¾U£ÍàŠÔTÃ¬›-+GSm{g* Ÿê±(gKª¹†Þ{r6£šj˜]'{J–qYÚÀJ¦D2‡yLÉ²‘]L™zƒ˜’ñ\–à8Ú›‘!v`3+¸ÙÌ
>µæ1©1mv^¦i M™ªahÄ~¡î¼ŒÔ`ßõ¡)câ,Œ;nVÕÅäqqR¸ï´äñ«ZÙ-ƒaÕÎ”	ä7B]˜ýË‘	¢DlÛÿ³¥UQµ~µÚôÙÊcÜ¶G®í<©\PÆö–#·¹Ö#j!‘0ÕÃ[e2„3û9ý#ÆÎ³K:Äà…Ñ¿‚õ~¬Fëk_*cd£ ˆ}8ÀWÖz¹ØZq½•bÑuâz«Äu×ˆEWˆ›¬‹JùEe|	ï”^x­Â¹pOâü´øãÑ¦dÔ°þd¦óÐ¼¶•§Ý ºÜŠ`ž&\¸ÍÊQ*J‚ŠÚËQ¥¾¦H:§˜±E>¾èyrGš}š\jø±AêÐŒ,³ºŠ£2Q<9«GGeqŽè¨èo-^ÎÔ(yµÅRÓ©%Àð´8?,@ªôßÈ(´SÆ)Ø.ÐP,•jFJæ¨ÿù¥–«èo0>N€Å
	 íš´nTPÁ‚g£bÑt)9pò©™¦n§4Di}i Ñ˜þÁ)ç0ÔÒ`9Z"Ü ‰¥Gñø4FëN4ªJ}
.Û›øÁeý“XÆ£zø
¤$8~øÃá“Ç{ùtVËc÷½¨ûcqj*Å ª“´Æ<Siµ5»ÉDàÝÁ—&FžÕÎlå{VÖê$-¥ˆìóœ^èã>âû°PÇ8d”ínZÑý¡˜”2>‡µ»–­c‰Á†Ã’ŽÖ½X©í ”;¬&ž0C(={Ö6NNaÒPÄ©¸¥íò§YvVDgñÏY±dÔú'Ì¾‡˜ß0zqøhåÛèÁ³'Ø4.RÐŒÓ‘‘îÃD®ÿý¿Íí>z´
¬c {¢ êI»’ÎH°|ï°z‘Ÿ—ñ4šø-‹ÇÑ4¾ êçi…Œ/b©’­ÍTru„U¦ðÂh©8f$pÆ¶mÓ·/¾PÇDQë8‚{‚AÄ†úÕ†^]òQÙ‘?º DK¿âÎ}âš¾ÊÎ9íÉ·]ZE6ÅËæV…ƒÆë$¾ìÖËÓ6;Ý„wÍ-*þrÚÔl×¡UÅKv»òmsËbu–o»´+Š:ÍŠ—MÓîéXT¨åDœ4ç‚(LÃ°ðˆlšHÙííTÎtÒ÷¨Þ½è‡CÌC:îE`ü—•dã~?–0?ÎÞéJ›Ñ!ÆHúgÂœÎ€²,tÿäK(a9%UR@•Áƒþ5„¯ùŒ³k£Ë¸¯NÏ¼óŽ­Zdzóù¥ ÂÕëÏ/Ñ5Ì‹óþ Ÿ ©«7^žÙCë@®|¼ŒuNÇ–hOõ0¡v—Y,¢Ÿ³\I1Ì+âŒt6JÛ‹´ÿ½ÁjŸòQã¸þÇyÆBÈGã—x÷=ã!6Ñèy\!âlÞ9ÈX›”þ‘ b¯^²ø?Ø=®*PQ&rXµ;+ÖFÍDÁ:Ç†tk2èV=ÃHçÖÓÄ<Á>™åPûµ‡0Ä7Í:ÛY’ÞžôH¡`Å…±·Õ[†¦d¿yCümL:_÷5÷b7à‘idMk§¥ùX“B&!d˜ÕyIÈâËfdµ1èc¨>ÝwO6izÑ›40òúÍO3TLŠÜ ƒ•ÚÐQÇšèˆO÷ÝPzÃ›*Žš›Rvªß’Œ¿ïœ€Ríà»Oð¦E(èz§±±%u\™ÄG®jW÷í“QZ6È|†9ÕáOmˆðýë cê3„*ïeãF±´EP¹Ëª”«óa)3uBÏn-#çÏê69ïl37bØí•õžIójÔŒCÃÖr#&‹Ÿœ
|Õ¸Á«·MP+råŸu²Žë'üZ<·’¥_`>™Y~÷Êƒ¹kU![óš×êy¸à	“aVk0Á^\’Áž[×Ä„*7Q€•	Ö[œí×Íq[„"®Ç5@Ž`n¯þrÄNÖñ¹ <â‚4ŒÑSE%k$ÚÆO>ØÑ¬ÄÌXŽëLVk[¾Z57Ó7€©=ÎÈ’„0bTl‚8šp¹KxLYÐú¥Ã#pŒyÝJ­ª¼ ÃÎI·laŽin]¤/ß
ìš¢ÊŽŸ+·¬é¿»3ì´.äWó»g™m]å‰ÞºgšqŒP®HÅ"@£óþõhœ?$¡ôpD‰îÄ	¯÷ˆ#ùÕnD“°®‘8ª™kÇKs·qBYÚ¨2^¥P†eÝk ŒnäZtñ“õ-D–@&WEÜ*‹.å*`ó2.á.¤Ó4¤Iðº ‹…ª.Þž®°±;þ"]
mÎúš¥s”«O«½ÊÀ¾ðuPuqµhz¼Å…mÙŽ¾ÊáC\ˆ‡Zï„¾ðšèè³êO^Ñ®L]¥Geœ¾5_¡*æàt-„ªôÒ€´]¨¥z×Žè(Ê¸¹3.èp·>ëÒ\( FoŠ.þŸË­'â Y ñYßu¢›H4WÅZvJÚi!CÄ­Á"Ó|ƒÇ)kG¨…Ô\·F8*¬MQt!°(1OóËêOqËš¨1w…öhª¢Çœ%Ì/Ç¢Èëƒ[Þ‰âJj|íh¾°Ètk…£úÚ¤–½mÂNx:ô^€&V‹g›‘²´®	ðËç—ž™+Í&Ë¬PÀ‘q%øûý?äVEÊŠ»öè•áì;N(Ès‰£Èi²b4ž½ÿ{‡•«†HûÆðžËöHG­Dq”«T¹(ú¹Ä$1Dai‘¦%G1˜í'¸LÄÊòãÂ§”ˆân<˜à]¦ÑÚ!¹ó+cÓçæÈ¥–¨¢FÆè6†#‡Æ˜%#I†¨•D‰Š·FNËì8ûi–Ff)@g>dùáR |(ŽDh Àë›oYê],|Ç	Ç`_ÄLœÚóÉ¨S!â?2–ÿÙ®ÓwuŸ"‡ð£Ø°]]v.òx‚‰ìÜ9Õi´ýøqZ6ŽAÚGb7‚ÒË§	îÝî<ˆ¾ŒpŸó‚~3Ð…`==?c÷Ol+j¯ß0Ê¬¤¬²4–8Hk˜¿ä~þ}™$&	î.cˆÐaKïgˆ#~þÑ3¼4ÇN'*#­øyå_ßÎqd1Æô˜‰À+U'2ˆéRÃºx\œ§åÃ¸b‰Ñèo³)ò1u;ªOã::G1ó³g5E’ŽÓ:åÉiäazVÅ$G«POvãäªÄ²V$•Ý	ÅCj,†@ë¾éò÷Šµ>“½ØÃù3Ó“%q™ v!C.*¸ý|wû@‹g9†ý˜î2ÑyHÁs†åLz™eS£º*»à‚íW6ÅEk’·©°ü-Ôï¶6š»Aá[˜NÍž	“¾=Så2"@&÷£jqÊ[œÒ>¡ò€èŽªlÅËVVÙ†üUªpi¶’™ÏwŠMK¶–`ä‹>ä•±˜ÑÿÜBiKÚÙb,.ig'äC”3_Þt¬ÉAÁ.X&VY÷ûñrtD¢$7O¥˜ƒ ô1ŸÇ´^yð¼‡G«4¯²:;Ëê‹Í¨w­§Žu‚–÷Rç°ä×QèÌ™,«üÕç¸åˆO~™º@ŠC%bÔfl›”ìŽ—ª°èD6‚aM•)À%š‰áTQQ.p'bÆâ }06¹a#<ÆÁdcEð4K¦7obÌD¼5ð,NAÏ\N³¶º)ýñÌÏ†Ê “$œ’¯¸ã±àEÌ('ÞfãñÒrècÈZ¶
xö§óÕ±8­¯®‘i·|Ä^[“JÍúìwx–£š§û¸>K-”Š2Z†1JG[8ŽL‡YM´Ÿ¢¾,r²‰ÁŸ–x/÷öŸ¼Þ}²ÿ|÷`{˜‰õ¤ê«„mƒT†¼
U„ÞZµé€5ïÓ+% Œ5`K –¬8(ºœËf”µeÓã“vUCF=°eUä‹‘“ƒpˆ0­¥F¼²SApu[_k2™~XoyQÞò~KQd Ð&•I½dÚy`ÅÂÃ‹vsW*¿+Õ'ÒÉ}V²Š±èv^SÍÒ`»ð²G+"è½*žâðbšÞÀÐ0åÄE3CÎ¬"„ójc‹éª ¸Ê¾dVµ0™U¦Ø,é‚ÆØX¿âí}z±¨ÀOá‡’ˆˆ}¹0*Á‚+Ì~Ÿ,¬žCûSI¤¾µˆ)x[îK§Ÿ¾É³/áÜ XÚ8¼õ	Û–ô%¹w¢åÛc#·¢>kéÞ=UOº[ï,ÿÀ›.´¶Zº8I&þµU†­dfAD²ì¤ôËÌº}D†«ßÁs.–%Ö`é) Á›´mý8 :%“]60ò×oì0U":Œj—1ÕõM¼´MæsEL0	â‚žit¢í3ÎqÔ(õ•—¬(œ<O±»Kß·‘¶ØíŠ›^yñsÕdÝ–¥WÖ¦8B‘‹+Lò$µQ¨T*ªÂŽ"ŽÂ;W™¤ ‡|'±ûZ:,aæx²i·³e=¢4ßŒÞ$9ž¿@<¯Þ˜M¥‹Í©‘ÿž-´¸8zãÕ¤±…V.¹Œÿ€â½2-¦³1æžWatš{Jî•8?Móh‚)ñ£bŠgà³®!÷õ£tD¥VßÝãcøj­¿äá¬tb}œ'ÒÌ.r[\øÀï¢S«»ZíwTÇú­Š ñVÃ³Ò@Ð8Çà¸l0\n×9YÍa'<9yJ›©pöÍêSqk(’Î Óù'ãÜ
††i¥á½€Zÿ•Ø&Ÿ9¤9ÛŒ^ÐV©Eó;3Û[éß°§ÏP‹ó‘ Žg,ëLÏQm‚‚?«°eÊ1Cd: G8«ãI\Öeœ•Ù>¡ä@£b¢ýØÆCúð4Évt¡.
äÿ‰²cÝ»­¬«*å*oBçàr¯¶¦¶Ú"q¾˜çPJ¦3i»@ æ¾•*KŸíåýßC™^)ÎÈ \8®€³Ò¸â{&¶"ÝáÀ“é+,?©aºù*#žhÅ‹kY+ECéõÀVŒîèò¡Dg>mEQ‡ˆY´@2že	qpçP=âQY:a½5d%€óõ,Ø‚5)®ªó¢L(IÀIQœŒSâ\[àÔ÷Iž¤Çñl\K5Ñ›1¬§[Ø¶¼!Šoá>ëÕØ†Xwð1Ú£yÒÄŠxæ*Yªgg*vÀc¶Ñ×-;Y	€§ÅYa`Z'¤Ö—ž#7‹qá#VÆ?ÇÅAÜ3ÞäÝç[ê:e3Z ©¾G¾-‡
4RkÞ¦Ê¯Å¥ÿŽ‘)‘ãë²€Ùš–gi‰Yg+h`Ú·@ôlâñ$Ëç”4—xÓÞì‹a»MöÛ*qZÔÅ‹ç±²úí£)ft{t¦Øú„J¼Š4	Pn·m2ª…d;üS›"Y\ÑÛ`ýïð*~µÆë”Å8\¼ãµ$föÌÜ5ÓXÐ8wü{íšTðf:Ã¡L£v÷ØGúªNôŠ›™,sº²MÜÀ.î—qøþÃ1EˆN"eØÝ‡âèœ|cKQ‰¶)ñTh_^€3EÃöV8Uë¿øÕ’Ov Ü¿£N°‹)Ëôª7¤C˜9ÝTB,ÉYƒ0–¬Š-wšðõøXßˆ§¦–°SÙT
·lÊÛ¥¿”ÕÓÊÒ$Ây†°Ä¬
á¯‚ã°<N&uácÃ,Z¶¸Ô¢†Y%ðŒÔÀÝò¡(Ü¿;)@mœ½ÿ[™¨.mYÆ·q¥¸/;ÞHÊvôz‘Âk´I”Ã¿š¡íéXäXsü¤"w˜†ðbXÖ²'Á.AŽ€·ß›7xàÏå;YjÀ„¡Xc/éÏ0A&åjÜqª$ô«¼˜‚}ð˜ÇON²“ÐÙ‹]U¢oë3ÎmZƒ†/dhd¸½=¡ƒ÷xO‘‹50êày%Ë•¥=álƒÛµdB ãMGÖæÜi{GvA„ô ÏÛõ˜›¬¯/4;"à'Œ·ã8­—5ûa¬ÃH¦ëÌJœ/eV#ç£4É’bØœ£˜¯XÑàal@™ž —ãz% ?Ææˆ.R“Ç¦ÁXdÜê­Œº6û=ýlÒ(Ç„1‚óUæ¼ØÞß‹þ˜^HõD„C ã*…îEŽ³äÇ	˜:uãD=Æ[ŸúK«ñ4[=Iq$VEÕ%Å°C4®ñyç%|íìÏbqç;$©q†§qu@è#æÄGb³´ç…fï™’Û#œC}BÈ7«1—€d´1kl+Ž$-F§ñYJ4N0\ŒÝh7ã;2Ct€ùVÞæÅù8MNR’H!%óÈ›ãÍÆ‰,žxïx!ô§‹¬dQ)Ñ¶÷"<–òã³+ƒ¡{
ãÌÈ…:E{ä?=*Äí]Íøeé¹;]žÜÊ3£øºl=Ú*ìS¿æ0‡g¹ü (b4îù‰ê•eÊÏèø=€TeúX¬ˆ;%%M±£ËMGa—½ŽÚSŠ*µJ¶»îîˆ(|+ë±‘``¸j4huy*0šM®×ßCg+’}uI¢T…BlÁgòO,<ÅÎ\Üáž1Yò6¯cjÃ¬ÈòxZêÜ×/‹ÄíÇ3z)¤Û¢sŠøà…TÙq£Ë7µ&;>uþë$ú¼>‚´"©ðIŸÀñ}øÿY0£~Ð9”<7mH¡SFdŽGg¼ŒCRûž Ç½À\-C²ò–HV„%EÀƒ|mRõ–ìe
,.„ßT¤ôÓ:®È/GŸ—£gÓ´$iDÎåÇ{‡6çYÜò““)Æ¹4ïšM}ˆK¦®ËæòœÕí´Ù]˜Ñ×e÷`î™¨Åå|æejš ~‘Ð<°N|Sîºïo“CÜ!5BôAv4Î
Ð(cÌÆ^$Åøý?NÀò¡ˆlÐ-³ñ)ª}}+™À(ºùÌ§ÃÐœ÷Bæ%gI3æ#vžÚ
‡ ¦àÕ9G0´n[ú)ÐÖx™Ãê¾äj>IKPhÄÝ âÐü;pÊ·8ûáØ'Tv/QWÉbˆ—$bç{·Lq…«S2¸‹â<@ÕÅ¨RJDÖÙP$¨5ö§Nk,²Ñ’]¬íW¶¢wcŽµ¹Àæ ÙÄAúÐyˆƒLu›ƒ47ß>É ›ƒÔ‰~Ÿƒ~c‰0Kðaå,ÁBYÃaŽÛûÁ*s~Ð·Î2˜Å:‡ÁGd†J0CÅ™¡êÈÕ/ÍÖˆr^P—{uUÛn‡½Í›a†3r~S·^Þ:»5^øå&Òa*øÇ¥©=Bê(VÃ]¶"öÂ'›ªÆ‰&cOn›d<iÝ²“´Î¢ŸJ²q{¤3Ý·K°× átj‹ˆcÕ-gb¨Aš$·GF›ÜiØ·:iÔÆkì Î'£þ¹ïÕ¢n¿×âÝ~-§ŸóV°”@ÓÈ‚Ž4Ë‘çÒ".Åä—ÉŽÑ,”Rôe‘[~\œP:ÜªXdµRµñÆYÇÇã,OÑƒ/²¤íN/ÇOc˜	ñ¡Õ >29‚‰jÉíŽ^ÓeÍ8k}*’ëôwR¾Ü¥ÁåÎ,k„Ö—Q{“)ÂnÄ÷Ú×z›
;Ü²"'@[f;¯r±p.¯ÉtùÇ:teÎN4ß|¶]ÜÕ¢TÒØ‰+éé `øSz300$hžP†¥Ò‰üm6~Â#4tè4ôkÉ±#‰s'ác—t;„¨š¹þÈ!¬ùC'†«¢H#³×é$FÄG{oÓ4+«£7pZñ#HpjäG+é4ZÚc*<CUsÞ)ù8•ÝÇ	Áéq¢6®?Nëºã¤»Õžë9ƒdò)z#$s6Ê’°Ô•M#Tv!'BÓ$\¿¸´¯ÖãE7kÞññu/Ä@5wýñBX@$6ç4ÊGkÙ÷T\Eª³u–£Nƒ+UèÐ¸ûçïQ#JÇ¯îÔ|Dëî#Zëá”­\@ë[MÏêjEË¢ñFß‹ Æ°f¨.˜7†X.xÚ…¥		eú®nFz—VÍò¤}êíŸ­/
êQ}½Áªoádœ°þÓ“xtaçÊˆúJÝ;J1ü›ºŠ^!ö<Â{)“ÀzUôboçzúßhgb7;zS”SÐêa~â»4‘7§"ÚWo¾× ¬^'2‡—3Þ@D_Ÿt
·«:i˜²®I‹€d@xâ˜yž,Ï$wChÓÍ2>^…Åf1¦ˆ3˜j~ƒÁ	—Ó¥Êvø²3G$ÐA`‚IR[fœ£W	ÓMÉ	ƒ„Ð˜yOvÃÏ3hfí{øç‡È®…ï¾¼}½æÕUxÑÜ¹—Àœ($!9²pa#u:Ëß¢9+ÛgËØµáVãê`1B˜è`||"‹ÜÂëË
 Sã™ ybÍÌWó'ÆN@“IV»÷_ó~²¾n’å 8®#ª…B”P×Ùãí‰˜@N `=5ÿP0—³´RÞ<l•PÏÓ–`ëíx¼±šrXaÉã8Ç°¶h—qô‡‚Î&5‹ 2m¿ø"2ñcQ_œm‹9ë„ÏŸ¡h¿w
õUr–‰-"dMƒ¸]@ñ½=GVÞ	¤{}š
;æe`Q.ƒ."T¡ÇEéwuvsZ¦Ç\/§]º{”‚ô±pÓí»–¬ì-m!	®t¯a¡ÙF9Jà¬•*…"	fLð>êoÿ­¯­­Eì>m¯,j[k\¨0ØÚBD Gmx[0ðÊ6c ÇF °"ˆá:å}›ÓŠËÿtÏàw|€Ûþýˆ!q"¨¨´S×I1[”Kþ|vF”˜TûÇüœ#gRÌðú=ª$Mƒ:£^$bh—1¡ 4¼ú0b–%–°4&ïËWœwÜ´ÿv§-Æèèöú0ÚÃ,>ùTi
¹ú(Õ…Q lEOwÿ¼ûŠ@ÉŸYMÄd~×•q`2Á),ÖxtJuÂ7g;‰kháh¨Rp>XÄb…Û¨Æ.KpBdçX5°‚Ì»’!<4C¼Bt¼‚Â‚EmÚY…Âëˆq¤˜EÄº¿B/%t¶~»­+dFí\sqQD2+SƒlyìRU˜I¢ð$~É+sWâ<ÆêÀZWÝXÌmœ–-"Î¤Öâ­#Ck)Ø,ŠœÚÉ@³gw”Ž\Š¡’•_ØGÊ£S¨ @h#na	©n}¹SÚ}ÉKœ{.dš>¹®Ø'Gôóôx(Ík±±Þ0,õ13B£¼bóg ÚM•÷òèÑíLŽ‘ÆùCé€%9Mf˜8ä¶áH7dØÎ~ ¶Ðö’j/ÇJ,²CçnÔ©®3Âü“º®#Ä~©¯–ÓgI™tìÆ$¹N®“MP)À.Ü’W}’a¥¥Ht…}â1#”ÌOÍ0G†ýR	yÊÇcƒ™Tæ¢w5ÁââÖF›ÿ’e,ãb!ˆþ` ,ó™e<c/‹gù,u$sS3ÝáØ'Âc‘ÄÎØ¤/&j‘)CZòð„“@‘ }–ƒv%dÃÐ#œü¯ÕÌq<›$TÍ¼ÈWDòrS@AhX¦Ô,,£¹ÍÇRÉÆ-©dÝ^éT²ø•­²$³í†æ ‚ù“HEÕEp­àN£,3sâX©hWQE¨ÔÀ:-ÚÜ
û©mî NF“DÔþ4ðŽÐ£¸y	Ð]±…fV›>79dèÕë^ôed.õ…‡½z×§Ã2Î“BÐQÞüÕ7ƒ!žÜOËÑï<C™>c¢Ô>ß,|ôˆ±sFÜè–›J¥}ÄYÑu“$é²Õ„’«›ž”µO½OY¹is)É^TÊb5»&PŸnñI‘Y®RoQMx½­ó:8—7éÿÖÁî#<÷wÎnº§ôG ÅoFöQyIg·¬åì$§á,6‘Œ¹ìÜÙèß—aÆÞç´åÓ¸ßêx w8U§}Zu7@µö*Ö±o‡ö³Ù7@mb{Œ­å‚
ÉA½c×&Ó­ˆš	mPM†,ø¿a“ù%÷dpÀÌ"›¬ìÀ>Áë›	Ûy"çð^
þù¢„öÂ¤ú0{n[‹ìÐNaÛ>ÄmD`ÃÐwµ… C§Ë5b×[4èámSW‹á†[Š¬Î&ƒÜZpîÇ0–Y^óµÃÁ¬dÃh}ÇˆU¾øF£1øÁµdŠ1•ÔJdæ˜xVf3w£ÙŠ!jÞ—é8Æd9…ó	[œ”ñ$AØ•}ìÍ“.B×H•¨û¾eH€–kL'°Z¶Ý‹Ö}IAg¤ÿá¨_´ˆ
QøT—{ð÷ŽãªVAIÇ‘x“é,-é¸¿˜Ô+wß“jEwa
•IŠÁ)G´Õ•2èìg(BLµšuôÓó–CÝ]*ðË™Ž-›½Øp°t)#±Ž.åØ
Ö«‹:  Ã¸* i2RöãÑX_äF@T‚™LƒØ[XŽ\ÁÉ î£v2àFôOÂø®È¹°>h„Jï§åH8„FIw³ôYV£¾i^ ›%LëAòŸ\øá™i/®žÐ‚Œ¶\hðœ]Ü2KVŸ<±s<áŸTµx7lÜ iÍ3)“}‰0×HX|›züÞ?Ã‚•9ðUò×åÜT#|ø[KR.DáÑ;Ð+YnïòÜiDùVTvœxV0
Ð<Ý—KK‘À|ÿÜ$n•½AqOê>Y¼eÎ1Œ˜kQêäÉÔ—°ÓÃmÒ6íE×¤9VØ]žqL)µíž1Õaj}C<Ÿ_†Œä{¶À—øøò¬T‘&ºÇ8ÕogOBÚ0
¤ÿõ/i†”OaZêó$16 ëKéx6ŽzbkRCyú({WA…
´ÀóÛàüïŽu¦‡1d¢™V[&ÒÍ[ÐYiöðhïƒO³‰æpŽS8‚Õüa—’I÷\à@CË ò±¶Ëò±3üåR¾qwE^ÜHÃáÊ‘M=\Ð†-xÖÊMÍÃnÌ'ÅzáÅ™óÕÂw-YD—æzÙQòÏýž’gRn¼ ’Ê-Y æâ'Žz&ÔIk_ÐªÖÖªì$ßË1ßÑ~1MeÒ.ÜƒM"¤}™Hée’„Þð†h[pyl%Ö’MÍO®$scÖîÕu›ÞÃäJddÄy“2OWFã¢J“•£‹
üt²{aë‘(BQNTzA‡ñUA÷ÛE1è>qFíXË	Z”ãâ¤¿$F‚ê+€ ·c›<GUø¦°Æ.ÌrüâgèB
¦{Ä´¢ô$ï„²þY/ÌE'½Û;k3Ç3P[&ïÿžgE”`¢Å(¦-1‘V.ÇÿÁÛ÷#}¦¤;4••L™
’ˆ ·1‘ä,‹%ç£u†Žñ¨ßŠŒ¬È
¤+£Ý’fitæ§ÐöQ¢gØ›K@gŸWŒÅ1€J“ÐÎ®é¯MG ¬Þ#>ŽË ¦Í“]XâZ#»ICLJJF)óô”LÑûê!nÐªl’¦äûÂõD$uU:oxÂò)MÙè¡ÈÍ0*ÃYÏ×xÅüªIå’ÓO¸ÐÔ¸çÉ¾LÑ+'>þŸ>‰Ÿê«çìA“éžÓš%¤{_uîhh•ªJ@‡Ó6tÙ·kwW¹..Ò›PõÆU¦Mö9|Oœ×Àö2†ˆ•xŒ×[^¬d9ÊÉ^³€ØÅ,š©˜§?‚ â  Å‹a´“ø8>£ÿ¿ÿ×˜ÍŸjã=§1›Î,P$ˆZ–ãÝ¿ÉÊHÓŠ¤_°ì9c'+&¿tS9$öJ^Ô+”¼­£¢ ¢ídjºÞÖ½çiüÖ ÕÜØ¶l$iáþî4˜Š`ê|ƒá‚I1_÷¦?þÚ±£VZ'ÚWˆa´d)‹^rr0ëF”P\wûo)¥6Ó•·Mê¨ê£q,ežT•Ö3~m•RÀÄÝQw¸ÏPkJ
–g³š¤È€>ÁŒp euê‚ýGsõ5é‡$g§¾ÊÜ9*¥{j™jHÉW-^‹Ì®WkDÖG2D2º)'Î&ÞÈÃ>P¾þ-ýjS$áµ²ðJ"°¸ä‘ò¨$úØ¦ÒÅœoVàÙo;Áî€âR¥êX’
ŽÞË‘U[&"™¼ˆ|G]A¨àaÄ/aøïåwGBsÆ‰ßbÊ÷8…‘]zŠ³ô„·d+œ¦t6¯k–ßH0)aD	”yX°Póð–Ÿf±qgÛT^‹ÇÊuÆÅÌæÁÙS•æùfÔ3ƒu–g5¼šLØ;áã„·ñ×ÊH½2å *…±´þµ7bï_²„´¶35czÌfüp Ò
xW/Å”¨e¨wû	nÃ’W<<-pcIÞÆrTÆixUÜ©ˆ…]ÝbÚ3›¢Ú Œe%dzwóg\$éýÂ²Å¬û:è¹è“¿ù ¿R™´]²ÜßuðDÂ©âù½äÈ!í›‚Ç/PÐiŽeüÄ @·ìÕÿ¶³z²LÉäœ9!g„ð·Ÿjnö†Ž<¤ëÑ³äF]Ê”ðõPæ‰æj6æˆ*õe‘¹ª>ùšÈ]ô¦D˜¼t^Ò…<6÷'‰üˆ÷ûmŠ94Íê±Î”âÐW+žÅeÂeÑ9ûñr€&'ûeA›F	©œÀTÅ7 Ù2Š ãû@5ï¿Úh*ñ‡4;9¥…^5:<¯þÉ@D«§á9¾±Á%´DHz¨ÁÐí?ý}o™‰¡%«VLËÑúÚ24¸-›U€=Ú>8ìYý%’7½4{â‹Ù˜ŽîâÀÊ¡fÐ²¯V†m±2ÀH³œñ.[BU0
z:èÐèjDâ§²zFçüÄõkò&qj5«ÐyWÙ»M	#<ÊÆã‡ÅT»õo€0ßÞ…ÿm|'vOÓ	Æ&+w×Ö¬j´á–&Ïq‡%Hê¯î
Ó¿Ë‘ø¯÷È&&µ˜¾«EûwïúŸyëF}-ü±¿³é,EMˆ¸tTŒ“%§dMÀ<{zðâñá³ç{ÛÑþógK\²þ»!"Mÿ¿„õ”Q¼„ò<÷œðŸÏPýJï9HÇŸ³D®í£ø(Å,íbNàèÏ¸½¼¨Þÿ#g`¹ØP}Ö£¯–£¯7ø
“5¯|ÇÆ¦+Q$IH¸,Ë®lNá†®}7¼kÁ±´öW›1tÞÈû16£Ï/ÝLõèWB€‡¾_E‰hÝŠÄ"ñþïx k\Ñ®NaúâW»˜üøÆÐæ«µ ¾»+ëxbIÃ\¯ø(—u›ìUôþWáÒ¨ÉÚ¥‘ÿNAù öÛXI²“¬†W“,Ÿ!‰ô+`É+Žè7ƒ¸SÆçr¦m|ƒ¬„³òë5ÎL3á©¾üõFhÂnTË^êôj†çaâ?ß}¼}øþ=ß{¶í?Þ~ú,ÚÙé÷ÿ¼ÿ¿ŸE»‡ÏáãÿøýÞÃg=OÁÛívDâ^Wù «ëÐ½ŠZ%®@9â`åÃÏêù²iGØÙÎm¸Ž: ¨0FÐ d¶	¥#²ÚÔ|DÉD5ü©q4¼  ·o¯3Õ´©¶õõpÖ…wî+ºþ[³ ¼b€õ…TmbEq:ÚÔ7ú¬‹2†ø®ol®­Áx’þ28Ê¶éÀ1Q@’B ¿,¼ÍsÒDÆ ÆX3€.à
¯ˆê´ˆß.ê­'ì
yWƒ7ðP¿‹¬z?ÕßŒ¹:°¾Biö)¤Ö&Ùñ1~Vñ.ñQåÖC;×o'`U¨ø¢R Fi6îkø«QŸÿSôþßÆ×à^·–TMŠ7Ÿ_*WÉ›àÚ¨RroÎ&jÝFÑy[0Ü/¦SÙÊúŠ}o˜Rí~Œ)wkèÅcå;|çCØ2r[Ò>ÖU¨ÏêÅE9Èërƒ=Þ |™MÞ%GVˆ‹ºO˜_á’º •UÉôY]¢82ùë¸¬ÿËft—ÅŸ‚ôßŒ^¾ìý_äöþ¯Ïw·ñÇÎÞöïŸ¾ÿ_‡ ëÛ@{ðx÷	}ÂÕàÿÝ=ˆv£'»ÿ€ZÚ~ßvp¸}Ð£5ä`T¸÷ýóîc|Þ{úþ>Ü{†?÷ŸoÿWú±óâùP¼Ø{&~ì¾MïÕ+ƒßQ‘\lšÕÅ|Ëâ§Çõ¦^æJ´Ìc]L7£¯¿]@5(Ù› <µgÖÎ,ÁÕvÖËMq™‹ç÷S·¼gäél‚Êá};†’¨~ƒ•Ô@aÖ„=FõÅ/Œ»´â •I #È‚WËÍFA­´@ Å](ÎÁIåq3úÿÛÇ‹”ACµõÔÑ¼­ûÓñxâZu‡7Ü5ƒËÐÆÒ
Ã{yw}9úª~{÷•Ñ›~ÇhÁÐãÙ$×ˆ0ˆkˆ6IêÔfTv{ÌEÑºSaÔ1Ÿ€¬ßØõ7œú_ßµ¿å|ÿÖùþµóý«ûû]çûúÜ}³pß¹5ÖæÕøvá6¾sé¼Þ^Ãžôûè|zµçÎúJD
«{œ’‹ž¾½OpÍP,’øýÛpðñmÆŠ*Ð$%NªyéŒìm?>Ü†eÅT©ˆ5‡š¿1¶zcíÌï–£ßm¼"Ó ,ªtåk¶;pŸ¼ÿ ÙçC^°¾ ¿c”zè«›Rë»Û¢ÖÃgO>~ñþî<›Û3”°üGÞ}eyNZÉ¶¿ûtg÷éáî­IPÄÇcr¿×éóâÜÈC.ß7pÛ@Ü7în¼B¥jwøJ×-ÎÒòx\œãF˜‡G ²½•r^iC¿Ç.ãÐGEQ;žO¨‹`E®Ïb¤½¶zó Ô%±B>;¦‚&žÙ:R´NGŠî1ˆðâË/9?àž^¶€ÞÃŒ›ôâ½ös|ëãŽ˜5ç3ùØ`¼y®7pÐãã:üV¢ý÷="Ž>¿Ì®ðƒô>Ò`Î3PÆæ^6,]°i{Ex2É&X·ŽôÜYú+Û¨•?ÎaJçÃbŠ—ÔA[Å¬žÎê~ïh\ÍÊ±åAÅäõtò­N•e…Z´^ºYª]<Npú	%ä¦›
/R›fî†ÔmŒ1zŸÑq¾{”þc„c{ðÿåh¥°BÛå¨ÈŽ³ÑÛèJF‹`’ øÃÑ?•eK»wéžm‘øA+q=_–¤Gqù‰ëÜl¨±ÆqU¡‹÷Þå(7Ìµt¾rŒûµÇcwtôzEŒftOW¾Š¦ïV¾Ž¦ðKziWÞ£ºŒé2Í"_ÁœcÉL¤BOtå¦KLfg5ã›­héèd…‰4RˆVÎO¡ù¨:MVÆ'ê—*°—8”Íh‰ª)á²¢0Ùd ïªwT5¸$áqîËÇpÌ¢
øþÞåÆÚ•C²%Ökú‰†V ï+¢Í
½p+ëëkKjø±ã¦§KN¾vê{8®¢Ui5s†ß
Ì•*dä›½I|wõÃ*–•~X,†O<B²Ž—ü˜dga–Y:ÎÞ¥I„ªÖ++kdØÀ??#¹ÏW~·Áˆ±<¢cR+¥ú!:ºŽÞJÊ¯ D*nm]ÁHŽñÉ&•¢êï°É“M´;³þ˜¬E`·Â¢qš%IškþËøÔ à€hVø+œ‚;WšØ6-–hªàÿV@YˆNÅü™®|³t_³g°J`vMŽ€8Ç6Xe¿ú9–:Åÿ5O6=ýÖ~œUuv|¡['GPy—OÁ"ÿJMŒÆ‡‚± Û6ì-z¹àƒŒ²„$uNH›Ñ[æ+5íí±9óí²eL\É‹<]ºÏµñ^$>›$€/××§ï^E¢y˜C‘;ÿçÀQÑ¤^Y_ºtíÕïÿ7˜ñÑï‹IZùí9Ô‘æ9=Z­¬G e”ÂìrØB-3¶óøÞåhÇÂú@2àÞÒ^þþï£¬X²?
‘tïÒ¥Ô”w!òY†.K×=7_¯¬–V»vóÁ,ãpo„û*·|ª¹Ç=³´êÞcS¥µÛ×íX™¥ÇxUu¸cÀË?b€ãÜŽM±à¨^ _ºÆÇÍ‡°DæI\’¿7Øïí,€‘†zfÏ%ALuº@–ÿpLÉNŸÌgLžO¾;oZYè[;r©®éÄÌ1¶`´ådSŸM¿N³tœÐMØW~Ùûßcòÿ!EúqëÏûõ¨¦¼†¥Ý£B;%t-
%ºtñ°,¦§Í½£.Ý)!÷ÝÂ­¥ùýrz…•óúôƒõ<¸êÊÃj™óðŽÊ,=‰GåüÅd‚¥VL>‚ÎœìVüøBiûÚmFô.ÍÁ½ƒ‡sÉdÕh÷´ªCº“Á­øñÉpX¦”ÑúÅ4LsxfW_‹¡ÊŸòä{˜ø3£ùóÂ¾¦=” òzúÈ5»õ(§èˆ
÷Ë\ÚÛAíÀ¬&sÕ)¦tÈòƒë!~ÖèàiÐÚæb_Ih]±×å[°¿lø¿áš‹#v0ÅAÕišºÇh°£R0“¨ÓEG-VF¦Ö"•_yî‚5¿ÏJMlî¬í]´Ç,Î]åÉsnÞÇ†™ìc§N™œ?ôÉTº¾žñÃ*Ø´ÜÄuÜÓzååª}×2}¿¶LßË;LmÅ)ì`ˆ>Z‘
óišŸÎ&½Àsš?:Ñ.Ãhºò5÷›\|”&®&·Ùér”Öçiš£³ÇÓìXø#Œ«p†Ñh/Ú~‰ª^º/œÞ—6•®B®Pó—oÂí£“
ÁkÃ»šfäùúœµ”€!÷8=®qÛä«h+êµisi]9«ÔVJõ˜7K`.—ÞÕ›+á(
4x…Éë+¿sÁN{«Ðxž®¬GšWÐ¹kuZí[­Hc€®/€<§´ô6¹Ó)xªyïú½+èÀ€vßî]^FçbOúÍç—¼5‰ßõ×–E Ì4ŒÝÂ,è~ë«ü-MA±œÏðèVôÕ7w™P‘'`ÅÈ2xÏÄWk:‘Òàê?½1›æÏ¤þày¯,^¬»3–8&é7®2›œ¸#Y•£{—"GÛô´¨‹ÏcÐÓ›ÓºžV›««³l%>‹ë¸Ä„5“Õxš­nåØúç²(ø@–Äèê‹#˜Á'Ä_÷Ö×Ž¾ûvý‹îÞ;>>~ãÉáx\“%À½… è3¶W
·õƒÍ¿%X™§e™–û¬÷–òbE½²‹º+C€Þ0™€¥VÎW\ÿ2ŸzÂ¯š„ß·ä†å˜|é~€’?¬N;€W¹ÒÔJG¯}p¾¯Û{!vYôH~@Õæukä®±#§e«µ&ß´nØÙ[G)§ë€õ³YÍ¶ÇÜa'Q{g¥Ô:ë8ä7»P
²ý“oV‰¼ê`eè°®Hœ,
`^ž­eK¦\f×Vœç‡E\Õý%}>žR4¦ù³\DÏãŸÑØ—™„èt…8!€GhÙœ<ßqAÇ©¤Ï€’c×1é¨O÷ÿ™ÜšÂc©ÄT»YRt_xW(ïÉ¦ßyÝ’‰ª©ilÈƒb7DXÅ‘¢p_¢7ˆ²üìýß \q'ÚËq».fÑùr	¬%ïÿ~’áYü ‰ÂYBækÿÒêy*Å)6=¶+¹Ôs¾·SrÿQ+ðF¢†°œžàµÙaªB‹¨ë7"ê£x¹Œ&E‚-Š›pd®;:I¤Kâž."÷þÿËÓÏäS8È1¨ã‰ž /Ì >ÑÌL/-Ö{JëÇÞËSôÐß‡ÅôBe¤b´Õ7Qê{SñmÿR—B0Òª: lªACuÅÂ2‹HüÊÇY,“^`åâÒ³–Th‘à¿J1o–B˜'tåI<°zu÷öŸ¼Þ}²ÿ|÷`Ûœ€ÏnŠ†häðC:Z'/a7ù±¶Ë2¾âEE}uÑ„IÊÐ^xL>¨V³2•@ör“ÿÛµyCÇJÉ)';ÞgëdY·33)zSdB	Ü4=ØW9ƒìÚ"$½N“‡(ÁXE’h*d$•õ	CÖõÊ&4- î¤×Àn,å´e5Ô‡áÞŒÌÂÃŽÍ&›’Ž&æIµc^â™²Í !M	+Ë][A‘ŽÎ¢*ûètlÓ¥1…yùPfyb§a0kïà™ŽÆÒå»$k7§¶B—Â¿¤\V×¤½jOî2–³qîÁ–bn#²Ð¨Ì™GÇ!L¸Ÿ–xÔõdÚwÒÁz‰½Ù×ÐÍÈ‰Ê”ßéªëN·[û¹ýR|JõmÈBâÈ‹òRq²Àc(/?¶î˜m
üSð‚ÿªøÿ$Ò¦8Â[°ü¨¿ölNsBIeí¸àÝø@Xsë6­UN,ñÀ<²õD-›ê>(¶8â«ï;+»t-ˆ‚ü›æû›æû›æûkÐ|e2·

8§/oOE¦°þßdkQøMí]HÏú´5^5¶n: ™PØÕíé¾3óHYÕ´*‹^ªu^äE” Uy‹a`k><¹¼»åVo¸Y…CÐþaŒ.º¶P
ÔõuJZ!\6¨„UÃ…µ7jí#èo¨- ¾	·#‹ãkv=²H¼è~ä÷Ü’BG¼M*†Û=¥˜…ÞÏWÎBµº*hÛ3çÅI‰]>*ã*§YYÜ¾‚¦CºSnzl*…(È¾ÏUÎÚ*ÏQÎÍ*Ù­ñZ
ÙÇÑÁ2äCûËµT,Î­·â‘Ä[¦ÅýÒÐ;uA£tpàõyåÀ½´†¯ÙxŸ;j7È—t>g¸zÓXª¥1ãüÄ¡oõ~JÇÊŸéÌoÈ÷åûqÀµRa¯s}‡¼7@WÒÄÙô~®AC+óQ$	ÙŒzÛuv³äh'P¸
|;Ä$¦¾¯ $"R¬4¢.ƒðHÇthWÊ‹è8MÜM÷•2…o¹#ñCÈ	ïóK†H,d•>Ì‰LVðZÜF·ØâE ’— r±npÚ±jtl…šÌ˜[~yyøÍ?žOB3‚=øU(¢žzc*ËËN¬ŠúÆªDO¬5uë‰Ý–¹úD´$ž‹X XuÝ[P ©õ~kTPq£á"ÕÓ"ÇÓçQœ'˜ßð„Î7. *_×-žèvÃˆ{NçN“ÿnæq¶?7Î{"bÓÜCvSwõkûLLßuPBhN^ôcB~ìÑüØÚÃ<²ôÜfk(’ÙÏ5»±k4–#PšDÈ*IÃy¾w¸ÛL«Ûu‘3Ø]Üäó¬¬ù®rG“sÜå´Ûq™“BKê]I)(ÈÉ&Ò3Ù~3Ù~3Ù>”Éf¹Î5¸ÿùmx·àOwÌ;©_Ï¼•»šwÅ8yÎ/ôÉÙUQ™Ö¬©L7Ùº¥Söz~EnLª:Æ$ûLIÐøLu¡¶‘RÐeD.tvéŸ•Ôu]„øí^¥C©Ïºµü+ ÃÎíN†¨qxëaX¤¢eó;Â¯m¨¨±‡.xœ6j1]ÝQÖgB†åúÄ&I-°Øfž•ÙÑˆ4^þÉ”üü“i³§! «2õ’SºµÌÑÕ&"Þ¨Z!Ñœ”V&X¡\/$œÇæ¯’­©=ÍˆŸBû‹¸¢å^àÒ
 *¿ªÑ½å›.k¤K$Ý›äµÅm7ß"Aé‚y‰»Uô©årQuU®op…»‰åeó#gÆ†~9r0dÈ±Ñ¤ýR-›š!Ì¿›”É«Dº^\ÊY3¸y×š6Ž7»×A·\d*XÃºÍ´á"g:i¾ŠR]²î0™3ñÐ_hpä­òð™Ó8XÒ»$aŸÙŽ¦{Éª3“­EÉ†ÄÀÕ®º†·£hÜtò^•ôÊs G™º’º
ân¹Äl)C¢v¾L	›'A$y=™a»á¨ä-‰ˆ&×UsEAÀ‡)r¾ÈðÔ»Ýð¯u
79Ûˆ+ZØëØøgc(\sµ¸VŽœÚåÊ{ÊY¸“Kn_Œ‹xQ‡œ«ç,>æîàõƒª­Û‹ÕsÐæÎ(BXMJokÉ`ëÚçg©£òmPcßÚý„Ds±“Ýù4Ç³h«Ìê}¨õ5ì‘ôøëˆüa÷¢ó2«ÓøÐOŽ®KTÖ„ïÒ‹¢°}o,¤‚C¡v†B÷%°eÅÝ6b–¾ã¼#RžÎ†Ÿ‚Þfhäèn&,þuSêðo¡¬ÐºEåŽ@ÞPÁÃ¿[Pò•ÛRôðïÖ•=üsfþYó@O%Îò0ƒ”f¶ì1Z4G/ `=7‹±%!Ð“IV÷oâý×qPÐûÿ&(V?¿TèêÍíì„:Ó-âjŸ²ÄPPyg6É/¶/ïëP÷Ø¡Þ{“–'³¼¶ÞTÈš”÷y9bª[ªŒ©»Í`,ol¯”-hÍò¤Oö¾¶ö)ë¶ÁÏò]Š ÙD-4ìnP»-÷sï÷x@»gv0
¢½ÅÚ+Ö;‰|‚±ÚŸ€ƒ¸±¡Ôq/úAj`ê‘Qw"OÔûùÇ”­ˆ¹¹eÃö¦?¾µ6ý©=Ô!…3å ¯ö<I‡P¬8K1qK¿‡­½ÓÃE¦
,e_’Óáòê¶ß¬Â5mñ¯ÇyŸž4ß‹'äùõe›’=sï\÷¾„ˆ9&!bšçÝöÕõ®ûåš!|Vð·ž™¸nØfcÆ·ƒ£¯°î—¨§ÙªR´ÛÏöÎðTì—f[xjm³Æ‰…„@¥­æÖMUÕ•À¢Ð QÅ:Ç%*"¦¤ãoÒõ’®R%išWX6pyÆ\Úœ©ÿ}Š¾…7ô|‘ˆ£7â…I7‚Š­-.Rªø8Ý-žúU™h»˜ÉÄHÄn/Ù
\<ìï|Bea);Ò[uv+˜Õa! ;^¾ÂNô=,ŽÇÈ>4ÑÆ•ÇÓÝœ,´zP†±ýŸfiyÑø-Gç§)ŒNOŽ Þvï^Ouu`_–§@]c”#x†ÓY¯ÃqCa«®~ó'Ä,:ºPüyôù¥¤ØÞœÒÍ7ö¥ÀF£WÄ"‡!Æ6íÙ´T`Ðèæ9&Qq¬*Rƒ—Š-ù„X»oQ‹"%ÉQ#¿aPNâ¸“¬g‘^—äZð›Ï º5kªb’öÓwYEÓ7ôäo½¡ué»1BCÙb@úk{¤FqÞat­ƒ>^ã4?©O£ûÑoÄº÷fî½‰ÕðÃ—÷¢¯×ÖBC::åo…‹‹S÷³e¬I}çX'B¶ç0µj¤Ÿ€ž~lO¹9yƒ5ÛdÉ¾Q—ÃPƒq‰}œQÎçJ\·£Fƒƒ+.²a¬!ÀÖ«EH7ÊKÊÓwf)wêâRÆjßáµ-ÀÂ-$I!ËNÂuŸŒ$zûš/mÿ|ðìéPô2;¾pÀ:÷û\^iUä¶ÔˆEÖùô¼çù|ßUt)fø2Q«f£6°è°´ÃHAF	¬×sX` ãlœŠ5=0ÚÃðo²«¬9ZB‡%A‘Û†ÏPªCÑ›¯J–‹ŒÝ&²Õ"&M˜
Œ[-ÎDÅ{äd†J#e¡ ªœùfUP7Ñ†*(»'°òxLEò°ÓÚ3‡·,U‘°TE´läþiGMìu5è7EåèS7Ž>0·Fe6Š‹×Ú,4%Þíkž²É4Õ…y2@qkå¤Ò (úiØŽüT¬C}!öÞõ¦ä`#mÃ{ëÁÝ½1Ûyòˆvúð«J·ÄŸ«Ûz¿7øŠE;ÁâŽ	(Ž_4ÀwáÒÚ¢ÐšçE,Žæy‹#4ˆ5Ÿtö".h KVs¼…ÅÑ½…ìGðŠ–¤ÉœDh‹	„¿ôäÁÇZsf3¢ÉÒ}=U£¤„…´œ‹k,¬Î<MÒù–ó;¹Õïç2Ä/ì÷“¤]ØïglÈï÷›°ý÷.l¥³®EŽ~ùÙÁHÞßèDü„$¢ô$ÞD=‰Jvô$Þ®ßpQ‘¤áHj4„uáÖM‚ !,Ë;uçÂŠ“- Âfd‚†°ûiÂŠÊÊ¾ÞÜ¶˜ù6aÙ«`kä?®!,‰ÍíÚ€ÌµL\ñ.TìÓ°ve—nní
¤·Çc.”šÎ¸f‰v£Z¢ËZLCëå«H©ÞËWæp¸Rƒiuër–ºwvÿ2¢ì°À0BøÏpÌÂxbD[uÙÆ&rF3™‘u/ü|±&Y9ì&vA1‘àr{‹­oAQÓv>¯OpW^SþâF^S^âkðšñs.Îk¬î'ÅkxÀ»Ò[Ð.¯Å×l_òuyÍ@™ÃkÂ‹wPŒg ½OÝqÏ9í· ï–¦qëÁ¢ÚîÕŽÙŒÛtöéÝJ°µ[¤¶«^ÞÈÓWI ÜÑ'©Ö)4Æ˜¶•¨õTäb‚—ê%ª"ýºš¤x©g¥ Ä?ê¥´‡é¼^ÅgéØùòvšiE¥KÏÆ5¼¯¡Y¾â=µê_†àI*Þ¢µ¤Y®‘j­³ÒŒôˆÜiñHzl«=,ÝåvHVžC²º‡¤d¨Eý‘l©ðÝ‘¿-ÿ—uÄ´iIø¨KAÃœâ^K¨ÃgÑoÒý£K÷–©åú'•¹°|æ@P¾6øVMYÕu!R]Ö÷<«jÖØ5I2˜º®_UO5aA’QÀ©ª~Æ¡a×_2ºÈfê›ûT[~p—ªÑ²þ£zTÕôäÕ€ gUY#XŠC1ËJ 2Óž>¬â·[tÁ2éÕÍ+¦eœ¥$ÛÂ>16;u‰ñªŸ G¬ÒïßÒ_Ú%æÀëyÄ®ÎWvd®•ÏÂ»”!÷+«Šðm­Z±]nVcØO{~=Ö_ˆ÷Ü{¸–ùÆ³¬q“ãfnë_ùâèÈ°÷q‘ÝHÅ	Mn³?~“-•0d‚~z²V?¯%†ZåOËîÏ5œø7—rÂí/ý}´Cù¹/÷jF§xTAëA'LåéHE3:au³7Ã…bœ4wtíh'“jR×BEWW£çäïêGÂÜ" F1žÊ9JþŠÌû®¦×‡q¡èÔ¢t>‰ô˜ñxûtñ¸„šg™jãžjÃPÀ%B~³½ÿÿÒç—š†WKÑïÿ†wáÿÌ›Q€p§ÿM§üÇ®ËÌ¦¢qŽé–ÙŠt‹¾˜§é9¶¼—Ogu¿×ë,#ß¼ÿk€4çßddƒãy½Åë,é).ßÒ-+·˜äÏízþ¥áëD¯´Ôjó`àY)ÕÓRòHÎòz9nŒ€êÔQ¨¨ï­rã,›ùÒcàŠ]ÑÕ$&öS(–§”Ot—oÉ7‡z°”R\d•X)xC"Å›¸­òCûD‘ $LŸæJ”ˆþXžLÔ˜_¢°>\{eM]–ÙM/dí3øÆrP!¯§êÎÑ5[ât™ßc½ctxxaC›Ó÷]‹_Ùæ÷N‹{–Eªö	<\+À³“V£…Â³ärk÷a0˜|mò4Ž|ØƒÏ½‘íT‹µÓ!wˆOÏ)w²|ñ…;†ÃSdùZÔl¢Eë{ÔÔd¢¥Ê4ŒLº|‡Æ…š¿\Mo"œÌŠ¨ûsxÙÞ
s+›•õöíx)RÝpÝ¥<lí„Vò û6tÞyñ ‹_T*ÌõåºD¨€[´eTµs0$›TüÍ÷	ò3ïNÇ‘o;u<cëv^»fç;HæÛì<hOút†{®pì×Ñs4O;{èŒ±?ì\cÚ‰éN{Þ	«g}*ÝE‚;åx…LT	n•kâôÍFCþÛ˜ž°A³'Á—µëõ¶’¨üñ\Ñ/Ótn¥èãÄN:I›â/Ë“AÅßòup‚‘EÙ0eÛŒˆgãÿƒlÖ3 D%#má´gsK¯%Q‰Ãh2R^ädš ÎGæ‰ïÔÈh&áuÊÖHº™
/%7I›äF­Â h—Ó‹^½šcL >¼¸%qÏ4­²âe×DÇÄÀ¬ØR7ÛÏÓ]Ë¶`XLUøˆ‡Åt®r­7Ò¬HTHT·«ˆâÈÏÑB9
^¸UµÙ	Ãy}¶Ú·Ž&*ÚŠØay‡ÈâG5ÑÅÄ-„=Z©¦»?§åÂèîÆ3÷ÔzßY§×áCt×ð.zê>¥s2c/öòŒz´ì¶ªó[\"þ¼âQ×X†9lNp ï”1<ÅŸ¯¡hW¾¬dcÒz¤ö–†ÐZîâµpÿ.‹Œ`pŽÿû²f>9@‘êDu‘^;=$´zXG#ã“$üwQ«ì˜Às¶º˜Te'=ŽgãZ(¼¡ÐWëÕÃ…H}úµ?cüòìèGÔíç»Ûœ%£¯m&	rž¥¿ƒ”3ããÃ4­¨V@ ­©•¾t©Ïpymôic„F6%è&ç^.›W³2•a÷{¹¹bæEò£ÊPjo]Ë‚ü7»4V}¥ð@qZ _7Ÿà[Ö!ãc@ò:°`—æ…â1!¹‹ir$»XÈÝ_Ûƒ,Ñl«†{˜m£³VæQÌÝ¥TŒ«Úg·ýuß¨lÚª´6+[6*å%½
k_²4ÅZËcªž`$üÝ)‚•uŸ;Cø5âÔÞ¢:tbqF#S4DEt‹xfè
ÀŠéÍC¼S43*@Ë’E(Ãòe]36 ûÌû”âšg^[ˆÀ‚óŽiŠÒÛýþÄ
:Óær£·	Ù0ÛZ6´µ4oÚ,Üö°où´’ÀñZûk|.0²³íµù“¡³%t3MÌ4j³M´yº8Žwß¾ßÖN‹ü÷º¾ý9s²«{ŸOÒOÕÅoõm Û\õVƒ]…CƒO»‹œèîÂn,è{>7Ú®Ú¯uõ¡eKØÿw`DÆõÄŽÎÈö¡½€7“=:?œhúµú-Íè¶|òßk{2n,déÀöx¼ûnšÂ¨á©Û±Üâv|ê
sëÜ†uc0ÿð\9£:þñOCÁKX×Òã¸¢k6?xê œ/Yì¨ˆ¨£®Ó[¤ŽssaçzÆW¸@¥Zÿõ©oWüp™Jºœ`4Î&Óù	;Ë[# ªÃâ_²):z^öø(¹Ô·¨j1›ÅDÖ‹¦¯‚YÞbÙÃ)$ÀœóîÅ77;Ha¼¼‚®XMÈÀxäItÞž”8H‹H ™Ìû¿-Oœû–—w¡xHÃáØnÐ­ëa!2Šj|QÙðî‰sè¢…±¡KnMºZ8I£mû4øL±)ÈúPT_]d-‰6óôFáôÍF6ñxÁƒÜ?
Î¿éEŸã¾ÕÛ^Ô<¡ƒÃ0Êþ¨Æ-ãfN ã,Ñ³³ùÎtvîÌ¦ãl„“B Ó2áŒÞ’k/”Æ¦—.ô 0åól=öíÐbºˆ{¨†©z¨ñöqXê"¡îEë!ÿßÒS$Öéû¿‰&ßÿC´%¢Û(€HáFg¾jlˆÜŸåÇE«pu5:(Ê¯‰›¦Ó¨>M£# áuÃ}|Âç ‡“Bžÿ_=ÍNNSIÜ"` é{™Ž@Ý*h‚|¬èPä5ÄOý~¼YóO¦Ô ‰ºƒG5Û-XŽ¡HR)å}Ôßþ)Z_[[‹˜ê•Exm0„yvm>7í?€öÚÚ÷>¶´ï•mj_E‘+‚ràÎÉ‰l—DM:‡Œ„~¹öÊZ
Ôä°ÏßcIy£ÔúÀ)¿—{ÉÔ•,nÎ7“D•##Ó¯™¤Xm}íáõà˜JhWÒÐ	ªÐ¡ÕˆÚ¾#Ó›?â±Â‚Qa”öÇÁu[#=ˆî¢I×r(É•[´9¹(UF«þ˜Ö¦¼öø·ûaf!œU–žzÒ¢kö¢„[XÛÕsj$A¥d0ô$‹©ŒÕE†­¬¢û¢–“ÛbeÝÉ<#À)òö“Æ¬iìCk†E¬x<~L‡ôŠ“mùhXÅ™û¢m¢„ªº u1# ´e8tPHt¶  +qÈ•™Æm‘\N_ÃJ¥› Í+ÓgÒÊŽqëàdúü2 „¼¹²ÖÅ>žW+#>híæ¹ @Lm#™£ç!%Gê¬pYn7Ó&¹‹	i‹öŒ ²ë|¤Ð6ì…7æÁ#¢8¾xgåùFmµ™wÛfã„.#D"nh•]y‰'Ì¨eR=ãêmµÎ~oˆßÛ'˜>Gü~×¯´¶»Ú`V¥C˜ý—øÈ¿7´)Ügžâ-:n÷o² à”<MÈaÊ¢6X^‰-€ˆ^Z>1òëôLû	)rLbÐB2°ÓÓ×ñY<Îb•hôö²ðyåN#½ É›}Q–3…Ì3Ê‘kfLàÝ†ýNÌ ûÌ$zÁÜ&øºŒTb2çRƒù|KÿTï’ÒÅ¬O°@›²†ˆ¿yé±šÌT‘‘Mˆy²$gSu¬pÛ,d‹J¹ËÖNmxfÁ›_M}1M)ï‰zÅ{¸ g‰ŠŒÚb¿7£þDyP¶†"ÑY`iVŸô–Ž¸¦W¸ ø’ÝäÁiöá4{q$Àiús-ÀL	×ÄM>ìØãiW SJCVP¢/¢é¬<LŠZ‘Ó®_LÁÐÊh…}±§@€iŒº>ˆ±¨•išGÇðÿŸÓéf7³%Q]çO7‹¥œëƒ»‡%Ï™#v{-¾b²”dqÛ„Y'$xò°ÉH5¸lèh«¸j¿Â¿Ùq‡ÓH×·},lÙØÝÝ¼éjà´ðÛI½MÛæV¬›3Ò:‡¥)À¥™U#.KŠj]XR”lçKÕ
«¡™³6ÌÙ>¿ëð´îÎÁîîWÃxÈµoC‡LÑFctK'%wV¯F“ÓZ’Ü³Mp…•DÊ*
-fïãêô¨ˆËÄy?*‹¼8)ãIì|øÿ  ÿÿì½ÛvW’ úî¯Hsj@›„HÊr¹©MÉ6§t¡IV™¥Ò‘’@’L@¢	Š2‡³¦ç<œ§y<àž‡^=kõS­þþÉ|É‰ˆ}‹}ËL ­ª.,[2÷5vìØqÛïÓÉq:qÎ²Q
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
²äðÄÚ?xõÍóg/v’{Éw;ßí<€ÝyÞOþP®ÿ7¬Åž	g½|¦îñ`ºw]¥Dè€qo2d<ašfIT ëûƒCzµÍÃ6MñTFÒ£4´±	öYÒ-ój®z‚#½ÊÏ¨&Dî³<Ó#Ã%Çö ä>‘ÝMhT¨t å”«Ýÿl3²þDG±ƒŸ25F	³<.rÎ®F*ªV^.ŽªFL—ðV¢µ;‰÷ÊäúŸÓ$ç”õ&¼xbÀñÕK“ç;»¿é/Þ®¾Jˆ1!’95úÀêê°þÊÂ©E—1`/‘$z‚QÍ{ýÏî÷“½û¯Žv^=Ûâ<CRûì%ìT5êÀ"Ï¤¤ÕÎN2\IâƒMG ŸSŠÓQA[ËXÆ‘<¥þåúgÚðe6¾þ§	Ž¶V8–â-C¶èÄÂÀ@‡ŒZ‡ÑË3JeƒëEz«;ÊAH¡êjÈà¿ÔGfßfÀv‘¶ˆ’¨ääº¹%ß®¤Åáh!ÀXDÐ4ª{oÈ¢ë?¥;ÝaV!Ý«‰[ÂèÛD Ÿ¬Ã:Ëq^¿@ê“š—3!äÂù8‚·Ú›ËàzýîÞS<rÔèï)¢J¸ u¡†ôÎaZ™ÓÈÔš;Nóþ¸f£²¯d’L&«RØ^§tî¯Ýïµv2JË³–tWT ˆcÓz§ý–£/Rãxâ*¡õt*´D0{?–Åd%`€‡ƒ3‚…¯«¨l­¿sp°ó_\¥dBê±2TÕªüê›ÿôl÷È¯Mv QUžEI\ÌCo$ÖòáÑÁÞËï’«PÛÜ*³@%‘?}‘
ˆÊÕP¨½ÐTÊ/PÍÚÔã{¤u5W#êß«u¿]9d%ï¯fÀ¤*=öñ&>òi]…®ý
Þ®&¯YD"àËf£Æ@´ ‡T×ôÞhª¿øÂWh!‡JÅ^ço\Am(RÕKgtdöAè>Ï«øä°šcÛ;æ‹Õ$÷ŒÞ€³
êc×T~o‚Ü>Šk¬K·F^çÛ|L1|e[®úJ8}ˆø!¾Þ¬²WÔ1 †µv;Æâ4¯§¸¥Çë—€	mÉIú/ÿ R!åÌ*Âeâš¿+Ow×«IiÍMuš¹©¡k¥¨ÏWlm[7kVXÏ”Ž
=oàWIJ­DÄ$$îÍö•#ý¾˜Ÿg\kmÉÁœ‚.ÙÛ±dLàwaÄÈÎ?2Îsb™²bd÷O!ù5-ø*%Ôõ“Tœ,0šý?Nþ8yJ§>Í?NV’/€Õì‘79%Ï=“â=¤â«‘?£“ŒŽï[8JGS»Öèü„íâiä›Ç¡pÛ-t—öZJ~ÖÃRšJn^ÛœÅé×¦]’–±9@tR)ÏGg‘`Þ®O¥·Q]ŸË]äÕa£‚5ÕŽXÙ@ªVs¹Üï@5r±“AÛ»~»?à•B£qt6¾:Sþ¸7ÓµZ—<ÅO	H¢ž±\À=rÝS"­qUô¨ç*Y1 \n’ƒ•ùŸ÷zÈì3x¿~ÃÏG©^ås—qLJ¼z”ªÑ¥Šh,ÒpAüÜ¤Ì%èYc.ÙnÒWTå<¬xYî);DTEÄz8×Òk—Üe+l34©GEÇY}J»k¢óº¿™¯h@¹É.}83»ƒ_‘]ÛÕÍhy¦ôÀñK…ÜÏïM5ç$z3lº% t
ñ•.´Kª£ÖÏT0}Ärà.“ LòŽ£é‡M”k¤·y'½[÷˜-Ê¶j€r X×M£o{¯×ß°>¶üê¤gDYœÕL³Ip<±¬ž5†^>Ë`Ñr‚ÿ,CÑB²¸ù$HÆ]¸ðõÀí-@¤bñfZP)F§T Ý{ÌÐƒDI°œ:/ékj—fë¢cÙÈÞ¬øv¾káP £=ÞZ7—çŽ¯c†M@ÈÓ¨-Í¢~['+¦±KŽ˜Ìâ;¬°Õ¯ŒS‰Äb’…[Qeù,V}GævxouÏë‘ìk¿çgöV²òå™£vƒ|)=ú³DØªÏ~â`W4Cáv#elh¾Qa¥JrÊráa¡~Q¡_rjpù¸b¾¡¬Ð•}Ì»b‡2`+¦úE'Â@Ü~F

"¨CP0tó0_™-žò0Ôv¹cJZør98ëí»û(š†Õ®‹XZ°Tái©ñÔó4HY
¤|[Jl¢•/Ž½eptâí[Ïš17€»­Úµ0¸ä¬YÔ¢²AæÛ(FY4\ø³rJò«xø6“OmþÂÃc{ƒ”H4%FgÀŸ¨Üœ|ì?ßyùêíÓgowvw^½Ýßyz°ó*â£¦evÃÕnö“!ú‹©ÓEK)Ã‘E¼H¾<C+ˆápÑ{€DÚ>8÷ÉàcrC³+Rj8ZÃÍ£éÅ—ßcÕ®;·U_.rV»£íJ˜3ÝèÀã‰9®¬‘©½‰åX8°ÜmÜ@xÏAžÃšS¹˜2•gÄ·NDHÄï¥¯" ]à×M·Ã÷MUÛÎœ”ÕÅ@Œó	v¯\ÓåìþdQtûXoÖêK4B¡òŒî¥Ö­“Hv,ÐÎÓBÂÖ½šHúú(€cF5‰EBêÉ´:×ˆÈk¾¦žÆFŒ°1W.ðÄãö–¤òÄÁõ¼›GË¶>ã«$ë’þµ~¨”[^Â’x<Yws\£rl3TWšð•>ÕüþÕ‰@kÇØÊì±š³ËÛŽ2 y×‡Ô‰…U®‹™Õx5û ÅzžçË3«×ÓŒcªo·©†U©ø®)Î•ÅíjÕ^(fC¼ FHív…¦Ô²š<ÒL@õœÒDIÌ¶UgôVáêkC£ô“oæ˜fµœ1K®Ô_
BuMA”¶7’õ:Ê.ª"HPˆ#_°wÿç¿ÿSò›Ë4ˆÂÂÑû]¯ÿc‘Oº?N:Ý×Æb}yÜŠèÍÞZÍ‚}ÉPÂw0ÖhÄy3|¢W†Âƒ´Ð`â.¯Ø{x Ü"¡¿Gî"äŽô]‡þ¿g‡ÉÁ³ýç{»;è"ôôUrøì»Ï^½Âà>6SþûéT1åW½­?N0ø[Ó«w¦;³
˜å–».
§øÁ÷U_]Ã
up˜xêÒúÄE#³!ôÀåÆLŸžØ¢C;ÔÐVóAb48[T´÷³½à[1ôá•ì+@fÉðµˆ ìß—A¾å	ÚÄS ˆÜÎ³:£&®îòM;G}ûÂ¶Ys»†Ý«)¦‘[5"ÀO¿"%FWÏXÏe‰$dxÕ%¨êâ§¬¾
ÃðïéP;ævA¼¨3ê]ýÍj}­·†Èz„!n+ÊB‘‹€™	ìFMZâé°´Gb6™Fßí»^‘ÂyýúÏÀe6C/õõ¡Ò2m/³S:ž“‹Ì]­|ŽÃwÉ– I‡ÄótR	â½«“€<;£ŒõU¤1"Ó Î®ÆŽ¯¾ÈïC£®ny˜úvÕ…|IQíÙªüæÒ;®Þ)HÛ™g¼ÏäÙvÑ^,!êg¤‘ÖÉ¦V¶aœÅº––[’‘–;%k¬ äûé¬ÊÓÑ#Gü¤6¤"
F;mr¥Ô&°¥œÐgDAÌ¯õKdqâ`h<²ßazŒª`žn™a*Âmw¢dè ×ç´Ê§ˆóÜrí^?Ñ,‹±'ØX6±…ãï)9„&W/ˆp1„Š7É!ŽÀç/+1
›¼NA.éQÉ¸¨7µZ—û_ç©º’P–v¢ÙX™°¨h	Ë+ÔßLgRžã!¯K§ —ÁŒ-!ÌH1Àsî¾’‘<zýÈFÕíuÐ¼f7GÒ™nK
Z7iˆËmº=óð&Í¢ÕñØÝXn¥×zƒY±d­Ð½Èp<;Å–ÛÝ—Ãt|šj6,½¯¿ë¾Àf¨¿à†¥UkÚ´Q§1{Ërp0È‰ñ¸Ÿ6+¶ÌRÛòW›SKÛÍ*mD¤y[D‘¬Ve8š|„½
ÿ³ÓÉ©yx._t
e0ˆ*£ˆâKOå£AƒÀº—fZåø‘‡:0w®÷]#Mj2þÔÁ®Üî<ßb ó¥£SÜhšºÃ$	vÞe’r-ûFÓ‚˜Ÿs¬—äÉ üôr¾-Æ>…ÓvU¼æ0k0V+·m¿{iêfl£Š‚êQAÁÞ9>,®aSŠÐ;¸¬1xxIÃHþ}†›ýš=Ž5{ì7{ÜÐ¬t€á®Aãö[/-³1í/¤O›òä˜÷Pwêa˜1ƒ¢aÊp˜k~Üê)é+d…Õd£&þ†ÍÔ‹RÌ²qz7ÿÄ÷|ÒuõU†žálÌ"j$”–9fÕð  tŠcØ$¦¨G½š¬¯Š=ÉzÕ»§Øº´	ÂÂVLÚÓºtà]QÙÒÊh ›Á®ªæ ‰‹R¯Ãª\¹0º‰‡]	Z¬E%1T+@Ž;ÚøXI§RÜ%x‚Ñ`"' zË_îáŸ"=ÅëfÙðÿŠù g$Ì€Är(Y}Ï‹}Y€’%t_2¬Ù8‚Åöãe s+rË¶Ÿ”ŽÝ…üP‘@šÿŠÕ…3ÚürTS¾?µâ…bU<»~Þ*ëˆäo(Bk,/¯)‡AO7$¶kÏ0¿O•ÿ—ø>Ç½ÖÔµWª½ËµœoÑº›Vç±sLÅÿ¥¨­ß|Tê>Õ;,A|¼Ä£Ì$_1…Büƒî[¦
QNn\Õ9 ëxì4ã¨˜èÎ¢Ï’I6‘H:°_ëX3YªZç«ÍÈ8ZkÑƒ'¶‡® KRÓÂÐ¶Pôp¸jkBƒB²8xOvÆZJ­é8WFŒM1#SóxÞÆ·á5œ/JÙÓðPÅ!¼\pd-Bé5‘|Frè½‹ª5Þ‰1¾à6ìŒÚGWSÇ&ì•]ŒUØgj¾;c¸.ÑêÓcâ8v[®!>†esLQÚšƒ¸C´ao@a’kšn‚¶E^â¾¥í×:²/°Ö·\LOBÏN¾ÒbU£UFõ…1ÊÉ'&WµŸ³‹îÅ‡	¦I:’Þêx‘¡¸(ô³¢üçç‡ÿ¹?¯òQÙ?.Š÷o~¨Ñ?xAOw9âÎï²‘Ò¼¶ZÃÀo«âmy–eU÷µœ£AÒ@Z•­¤óì»á¼¸áÙ ;G{Üè9ž¦4¤uvFU¡â¡èH7érÓïòàû`—d+2Ò_’Û…m:Jˆ|€MÎ€'’ª‰ïÚ·€/sñÄ€¯^/|cüm„¸‰ìÞÙÑÙ§"|Þˆ˜O1Ý'%öÜ¶ª• žû7‡‹œÁE~çp94ÁÏ‚€9”.ƒ‹ÁE¯A§d:×ÿ¨%y®;óZña&C}‡‹ï±àN„é¼*ÆäÉjbGNQ	1%Vß@:ˆõ1T~~ýÏ4JZ„‡ÊO+˜%³è[
|(yððÆeÔÜëŸÏ³¶ûÉØµ½:¿Ûß£~ŸæX>Çè{Ê9IŽ~c}ý?ò»8Û
;¤q
Â…màþãë™”"LƒIŽÙaOeLl[ä‘¥pR³”Âp‘‹ÅØVÂ‹çóï9…«z‰þ-3¸EòøÛ|”Q•Žˆñv}1«01AÿbT^àÙ=áøŠrIt÷@…÷+ÊÅI¯É{…¬ìÔ¹†ÖyJD‘@ÚÁì"@ùißÎÛ€8‘^Iê(N&}Ïv¹ÔðÀD<\œ êga0"0ÜÌëìK’ØáÙ.¬÷e!ß•7‚B{ÚCò§¯S†®'r „®'jL(]Lˆ‘½†+…AZçÝyŒ“¯ðýCŸŒ•ý÷Ó\·X”mÈVÙ§½Íxâñâëº8Ý!*øvfv„îˆ€ì{Džž]²‘q²-áÁÛß…B§Ù³s˜Æ£ï^<ß›LçÕ3q=ñhÈ6È4f}¡PCY'+·û:œ ]+Ç‡&N4=¾WÆd£¡¹P¹‡ê½x±kRHtU:]d—r
²2†ñîPhž™_ž•æ<<§¢çwžŽæ8áNç!î0À-ºØ r@=™oƒ°ÆynaD’lUxŠxuž„–)(®Â¨(E¦w9'ò–¶Vs6(O!86e¢Gêô€¨SD¼ît 1|8¯,å°-ÊcL\\ƒ¢<ÛòÒŠ«j7'öÖ¦Y}ë'ÍÒ[½‰þú\@$Ç±õ‡NIöâEôM“šFd?‡ïsØuÝDK¨^¢HZÔö(7
óe•ô´|{yKŸíÎÝO«¼›ª ~Æ+(ò"Ê
Z)>Ì*ÒœªâÑD¦n…Cá|]ŸÊTë°¥~ö»Qq,6òCzg >Xÿ7ö{É¶÷Ì„-þ‡uLtî¥PäúfÖS7³§”ÚQKgâN‘rúXúýöîKX<Á@~t‘äd”]`h!Šj‡¿¥£Á(£ýHT8Ç+\rM84KŒŠ§Ì/OüÖåƒgâÓ†éB
µzìe’Ø‘ìW:^zwåå·OWœ·ècŸÂâÜ{ýÇùúýõõ5üóÕÉ›{§«ÉJ¼ðÿ®ý´¾öª©¦ñûl„wÜ1†-&lT:Ÿ€Pó!9þ˜Œ&å KéNÉûì#”A‡n9°JöIç»Ð‚¼ŠUØþõ,g"Þ=ô‡n.°þÅ	õNhTÿØÑÑ@×ËÀaÏM§¿ƒ¾ wì’¸®÷Ø¡@÷½mCa_Ù•$èšÀkI¯¨ë>¶„sé…É¨élx ª=–]‰V13H0¿2ã3 •í=BÃz®rF–PIU«÷ÆªyŸ’úZå):bùåÕÖpÓúš)ˆ„—0ä×bØo:ÅpàXÈ3\ª‡ÊV)±«ø ‘<¯ø‘+
…Ýþb±?íïqŒ¸òè$T3^w´ü¥/Œe Ÿõsë±õƒ—}§1Â-YåÕ|ä4Oá£¡M¡–H:2ä³ì,=Éé+EXïØ¦¯Žwä ” '™w‚ò¾:þÆ¿ÀáÅ lèe¥r‘a’^Ux¡¹B–_Ð_…xZrÇ uJÓê0K«9Šìþˆ¤˜ª%S“4gRL¬„àÁ/i0³ô\¿36ù
6jÉ?¦sNt‡ôË¨@°>=¢iZqö§
Ú%ë‰ó:S’©îO=Ãü~ÿÓ\CWáKÉóZ:»ÈS…Â¬6NžJÄ´~ˆPœõ‡†^lØ¥DYð)@«YMKõm®šF©zÃT×ô:“b©îOþfR*¡¦xÚ¡$4òü-_)!U?f¿z ì(O¥»M~@‰,ˆž$ùšƒ¿´LPUí›‚Ô>Ú‹ ëUwøSýÈ`EeøS)ýÛ1Æ&Ñlh0ÓYú[ü¥$h1)x@Qôa>S½.IQæh]Ó4A°õ;Da^÷Ç${Þ+{,èÀìTLó"Ì%!ÜÓ3 ò$é`gï§¹îµZ9€5ñX,ºÒx¡[OiÓfàR~OK ú}hŠ|†âÎº
V$ç'ªö¬‡¥¤LHix-À5œ'ÅëðVwiTR£A0Ô©©Q!©¢üä,“×“ÃM¥“,¢&Øï]ÇÕýâ3iG$ÿT·‰¯&@Ë²ÓtlxVm:?© HGç´ž5L#©`[œ Û0LÎ±Ð!–ÑŽI],	ãýPÌ†u¬¡$XWÖüpÜ?T?¥×Ù$æŠ%Ù!Í•M†Ledc\ÙûŒ+3|—d#?ôïÝ^@0cÞÖc._ë§o¤c[Ó‚ã:*ö†/ÒéiúfÃGÊùBŠ„xOóŠôPÈ$^*É\ä(é8.×ÞÓÈ:SŠ‘¢ÑH²&.Ö&©¡±–=2ºB&âÉ2ä?HŠpœqAá{ ÂFäÒ7‰×¿«•Ð>W½PôdõCK·úíc_ŒÕ@É2bMU¥Àu1*PŸ]««fI~sò;·GFÀî±ž¼gKy@R²EÐ“­šéXÎ¦äV(¿«nÙ#[F·6Å¶[ÏVXe·â¿ç¦&‡ºùÕê<¬%íÉbklO¼V«øÆ•.Ø¾öÊF]ÓPëŸ¥³ÁJ¿“ìÃè£¼a>LFyYù8„Ev†C¢/ŽöIÐ¥©=ƒF“ÍC=šªîÀ_pbb²F?†I]¬D:ÈË@ÃBU€¨‚Œ1Š-¼ñ¶ààub ©)ã‚D†&È€R$À‰ë&gÂAÑG	ïñîš‘±×C§-ñÈ¹¹?:ôˆnQª2}SŠÜ÷¿Rûts5ùª cHÚq–F3j]ò'@‹|÷îEÛ S‡ìWNŸX<Xwj7MÉœ°¨®¿‡ýÚŠ:"©ƒ]â%Êu/´°½¼þW¼¿á4Už€¯¿E|ÜNÑ†L-p×ìNû¡§¥ÿâ‹ #áTsÔwßÎ'u¸ËÔ'¨É)E4A”~F£üÃ/	cFŽ×µI<ÄjÎ¹.[ få [A­)þPzSøõÝC™Ši5Ýë7>ÎÒ2^³äQ54˜¢áTô…ö1Ø­d˜+éæùúÝ¨ÐÊ6aåW'7$½›¿EÆõÃYD•(÷,S@’ÊÏQQ¼GÊ‡êÁ÷°'0“Ñ|<A—¸EÖ$Z*Dm ÜžŽ×è”§‡ÀÇeî;¹EG#Å‡"O!„ÞŠîå)ýŽ&ÀTô€ÉØ¢´neñÒ¼R’ŸŒJ åNú%å4\¡ÀQú–ãÆƒUö=/;C€ìcUÞÊ’£×tZý¼¤¿Ró
•JÏ3»8™Bà´4ºØ‰Ð)Ç4 ’ÓßÅ’Š%ÄjŽ¼ o•²ÂÈï™EX*(@„4¢4›à'®•à‰F.¿i†vTØíÁCÜüa´O÷XÜ;ÁDˆØ,n¹NÐï’.oL†éLJˆj›;›ØÎMçŠ¹ëIFüÂPÄ6†„ˆ¾¦š°·Åk"
UíoD ÿ™Îp½åyxf‡Ç…Ÿ ºª	©C(õFC£³³(_$áù#g³Á³š½FÆ˜’í¸Y«'ª‘t¢õ³sŒLÆ‚'ÂfÀÌdê·1ÀÃ€9à±07b`‡‡î]ÃJzâ0€SÊÂV¹ŽæòÙ*’ð½È5dÆYhcB—¯ÏÞ#‰€ÆkÑŒg*qöª«p€µœ´¹¬Q×©lç4GÇÒ}íÙ%lûÄÏ m{v…Úcšé™Ðùá)¤õÛžÊÚ²Vúc.T~wL\*'õCêË´2ÿZ”âÖ$ðu·o|cúYZ>Í†”÷IŒ… 3à4‡ýT“àí"y­Ù˜%ríg—D5>‰™dãˆnw– OópVÅè!‰ç²8#½ÖƒB9aœi XÕ†µàIù9V‰ÞLyñT´,ÉU²’E¹U«–uÄ‰*ÛˆO…òÆ¸b¦"u¤8âUc­ØàPñêi³T¤’*»ž£Þu2,ª¤bFZ™^+©#n!1­~Zk®,ØÓÐî*À›Ê'±:o¼áJI2Šw/åRuÅš™.éw Cë¹×–7£*-ZmÓ£|ÒcººÈ®<óªíH-k¬^­&5~–«ÊS-º²2îbO‡×ÓÌç’c,„ÃšõX–¾*HyštVø	N¾IouGÀƒO@SeÆ#Õ-¯×‹ê«ÐÃ-1x?­lPo%Š:=~juUÔ¯ÒWòåjÖZáÇÓ\)Òäm«»ÂÏ•·ã‚:,io)bJðv7h³ê…ÈÍ0'$­ '$ß,Î	©Š¾&a”–Õ>3uÂE\ûÒ-¸*Å0ûœSðD6	)_)¼CŒ+Cìàóº	ƒ¤ÚPŒ%+}nIi–Íô DfÜs×Q’Xñ£é÷ áÑsÑ Øƒ)®{qÈj•[pÉûS”ò]U®å–±ª9â_¶-õ}aÔ÷@ÓŸ¸¼{‘Î(ôv6N~s`¯Þ1ž0BÆçã}:BÚ‹¶¡Â¥ÍèžîËùø8›ukë6z±ãáo¼T*¸štŠ‡ìó]‰5†ß,äjÒÐI^>ËìÈ„Rêžé>‰kòù$Ù&9På±Ú5w¼d
9%Õ)œwáÜ“á{-þ_xÁÍ fcíëKàÜ°É§­¸7}ÏÍpoÙøî¸·l,¹7Ño”{Ó0Y,Ø¶X0¿Œ¹h-ÜìÏ/ÿ!ÄÚ™ëÖlu›øAqÙZ’³›2zKàÞ8ºö@í­…wÅgŽì~ˆ#3—ê<–L«ÁäÈ˜úŒWûõ2©wô˜f_»kvLêD†BÙó¢½§?80eÝ‹±`,v¼[W™	#UÑkÓ«#í‡7ÒÄµc4ÿ¤¼=ê5(]0ã3v’Cæý7¦µŽi•[¡Ïz¨7“º%«~²ë,ý õ¨ËS˜F¹Yg+³|…ëß˜g*-™g£ZŒQÖ¡ ˆI-ƒLjÙÄ¤/ø;üªÜr1Ž™ßaýÕÆrÖÙ¹'¼ÎZlÃ8k{¢æ›áÉ]±Í 1âšá3Íjì¿ÏÜ¨àÂ; Lû}«„Bû²´3>	ù2Â«#Ï£ÃO0WEd"Mg
w¡›s³î´mÔŽ¡›#n*ÖNpš~ÀÝ˜Ï#-A?³7Ÿ¶Ÿ[3Ôtd¡!IÉ"sJP›ËÛ|nœƒz1Ió£ÃƒØôûI~’£¤$â¿ôT4t:	žAà„÷ä”‚É‡t6\;ÉG#(ÀZb"KI-ÖÍ—°˜%.â<æN<§÷&‘L9Ty²–L 9– ;ÂŠ©»#a¹M2vß°­–U¶Ú˜••=2(†o€£¸…xïvnàO e$KÛÒÂT»… ÊÞXœ…b{7ö‚€b/oê±3SÂÎÄfZ˜IÎÅŒ´¨ÔŒWì§ù‚r³èkqAu·/*ßTºEÚ'’šrª†”˜2 2†Om%KÄØ±¤,Ÿù±©PDnŠ5£‚‚”éÊƒÖUx>{ž÷Šyh†áQ¿hY‹—IBzŸY=¾´ýW)×¦¯¥‹ñUµ)w	2¾Œ–èç7XP¡¿§ _ÊÿÀyNŠ9b•”QSjØý· DŽ!5¸j>\; ¹Ç°×˜“;Ù®Üvhv­šæù©&\Œ• Ø²§wþúÏ"èý0Ktæ_Ü¬Wï–!XÀ«=Õ:ù¯Âú±s›W|p²^G5îMqýƒ×Ê*§„«´†«Õ0ñ#8¬
ÚxìØ¡ñ­žQp»ÇmÚæ‚‚þMË’¿`öåHë¤ŽhkáŒÅ’ÞÜÆxCçÝ$ )Ò B«üiZ(ñ6¥…fÐh¢El•;\Ôh^¼Ý¤ƒ¡ê7±]ŠŠµöKü»›gáJ‹êøqoD(¨4X*©j(›S¿F9èrl5äPÇWA%RY4ðÐ6l´'ž‰Äo,F ‰j£[Žhtºp¥îâ*æ_bH7&ZŸr?A•Vœ`ÅÕÃ¢Ä/E¯¤Þ˜q*Áb·ÕSW²Ö,¦“*’µÊ`µ/~êõÉRŸ°h£}²Ô3,¤QÆOVY*‡c:Yh6Ó!ãçW<­šôÅT3~Xµ0dùgUìO#¥¿âõöhåfMß0<€Ng¦Ë…³qù¹¸ôíô½*—G¶¸…é¶Š‘’ƒ
§pö“ÒJSz&qÑÕ™R¿ßwÝÑY¤xÑ´NQ¶Ü­JZÌÕTýªç!êŽ3}¡þºœþ,’ÍbÀËîØME½¦ÞJêÍêìMHŸíÜÖ·×N_Î“/'_®{—®Õ’·Éô«LÜ´(XÛ]‰½œ¯b_Ô•[›©ß1R!êß©Nèp¹zË¬ÒIœ°J )¿{êæÙmRÜ5%v¶Ë|$bb¶
K´sžƒ€
Ü½oõÌ$L'l?Ê”HN·"ìµ8èe;Ã,)ó	ùË÷]1§å$ÙJV ¾P§?šžÚtbTbg:}t‚!è0{)P¢‹"ÚE¼¡ß‡0dhß""îNŽåÚ$Òb5Ó‡Wœgˆ1ÝÒ¢·ÈÔÍ²¡¹³Óé=Ôp¿ä‹‘ðU!7,D¹dúFSÃ¬‚a²IøÜñ±z7™ª<uø¥×sÌkIb8zš!=}ËÉÚ:|õ²/ˆf~òQešìÅælÂÞŠ‚Ö*[³½õá¤o9dgÔ,,´è±ÑÖ­§YÚ™“à—Z"ë´XÆ
á$¾@ªÓÅÖ‡Õ[žR-/NèÝ¤T«SÞluØÁyg‹Ãé¤ÈB ò3ðKYHÀqƒ Ó³£BWzœˆ ˜°k™™z
d)‚E¤ÝµL%-qŽâet-KREšêK/0ŒL0Ð}-WP‡b½¼q.á`¶Ûc»ÅÊÀYzœO4Î`Qg™ÝÙ¨ i,4ÇçªMK:0Wß)Ûñ¶lŸ-µŸ…M‹zõÅ~»‘ùñçÒ…~=[!ë“\/·šà4áWP@‰åØ5ËÇ“`Š™¬Ê–u
²—ªÇU[†"
Ò>mµ®ÌÄUƒHÓ[”ßª¡;$@Á­™¨’·%ÇToJ`äÌ2¹—Èˆ».aP›hƒ²]ö,ŠZ"yøühoÿÕáÛg/öžî˜ÈReÏ¡~_­É…ê¾5Åðû
Ùnq[Â˜š¢å]ÐŽ²5í€‘5k˜Ê0;ýE¨Or	ÄÃÌeU4ÞD?ÊO‰~ðÑß!Qßb ”bŽƒ/Tp™äR¿´e+•€YhØ;•@éjÕjPò±nƒŽÁÅjÐIÙd5¨X/»=W%Æšsó;Y­©ÜNâ¯SiXy~Ðxí4eb@2÷oÇ.ÉOF@Š&Ž¤[³Ò×ê}€H‰ŒI3ød˜ÜÌ—þÓQ6« ìÁÁ«d÷àúíí¾B<N
8¨ÿ4ÏÏ‹„2£ÅÅ~w2šÌT†vD”·oæ''p°è¼q<lÁÓ½îL†¨Ã¡)*Ìn#Sœa`¯­ä©	ï¥(e`Y!¾©¯VXwò.‘ítÓÊ[5Éî$ðwýÂ|ù¸²RÒ`Œ%Hì3bwX ƒáÈf7V“¾g}ª¹Š§ã½E*v7“¿‡ÚtkjÏx¸±®£R:{^hNeDá=5µçÙy6òú¶üÌü 7Y¿
8}©Rô\ƒøÖá‹¼dtRÂåtÈœ¿Øß±;XÉÎÞ–Õ:/ú#p2Òìy¬b•à»Ÿí¾Cé¿Ë°­ËFloUV¨´SÉ£öuˆ¿†Jö$ØÁ÷m [ÉÐr
i«#+*Z²ÎH½¢WX;÷€] ²·.AØÖÖgº1š*R±±ÚÃù¤6×½­†ä­Í%4PÒŽÏkŒ™Û÷vˆŽÊî®D†»‘Y¤%áÉæÝïB!Ê0ÊÐ?™%ÃØ;„üä&)$M~,Û¶)70É3¿%æ>”ÎtŒVÊS)g!ETýøj²/Þ—È³[ÉqQÈwb:6µý\Ö¦­âÊù™Aó8Íƒì¤/I¿ÒÙ¤»âÃäÇëŸL*ý3^º¤´0¤˜^MŽGÅŸær® É\ª«‡â†×0í›”t|P¢Ã‚)"í7‡+ÊK~4DÎÜ€OÅ<¹:@P¬Á‡]õøï£Í­õuø¯ƒW9u‰žMÒá9öNl×oQÍøãS‡óW­[ÀâU±2@YF²‹"ff‘íáÑ>@tQ‰WjÙVWüõŽ«*’µä7—Ðä6ÉDTîe<Çîó,eô°A®®Ù;ÆP]^@ËP3=ró7¬•­ä(cþ¬ñ”„P±Pr™å­ÖºkŠÑNµš§'¦ <æ%¦R¥“ìG•"M¿·MÏî~6åv [öBØeœ—Š4jóÜÂ$£9²ž™€¤@Õu1©¤‹I]F#Üí“b’é×=kù¨UÝèQ±7Vœ…Kx¨b:€—ß%ârÁ,;‘™=+°Ó†4:©ìkS¥Zu~À³˜‚2\CVaÖÅ)šµpã³Ax—àÚ"Ú9ÒñK°Ô ¶ºÅŸHÿ5¤xÔS®FÑ:z59Ã§rôÖÕ,ø×Y™BÊ\ZÄ.¤§Oè1·ƒhW+½DÈ“…Õo/jïT^™r¶S5¸í>ôn[Pï«;è¢ES¸	Ü¦	±áˆÍwwa‹ö¬õcÍ‘þý”ªâz¹oœûªò¾¾…Î*ÈêæÓ.{á]‹]f®½}èdy½¹x'yˆ5§õ‹]]G†P¨Y"fQ¦ò†ÄÆ”¸€ÌU"}Ùb½÷:?5h‹•žƒ¼››ÃÑ'xS…ÆC5èùt‹‹w^$k$ÍM‰&ß†ÁG7QDÎáEeäo½Jƒ©ÐEA'Ž”‰Œ1—qÇ´æŒvK;êX­2'G‡øu:v+¦Ü4^J»^»çƒ]
æ€q¢xˆvÂÜmª	·’^G8I¦ŽwdqkwŽ±Ê·œ>H°•¬ógj5Ý²í<Þ¸¾-Ì5²ZÊfÙ,=ÃÀÁ‹¨¡.æ­Êu·v«*ÒÅ„5ÊÖ¶kƒ,õè²Â´E™¶Hs3´‰!Nê„‘'ñ•òœ|Y«æú”•8?³êŠÔ‹ÒŒ_1>÷Õ´ÊÇhŽô.Ô)–R>g°ÚU¯»Ù¤œ“p<Ñ¢%åç*{S	½6‹ç¾`´Dtš+Tê™ÐªÞ ¦Ih&¼wt,ç Å‡[èL×%9(*a+t£*13!™/ÉRúj§g|ÔíHá©Î=>C—*™»"×¶s%KXÙn@9«V	uË&µ—#R0aÔSR„&°#~¿ësŽùffYöSVb
Ë³tr
_&Igw–ãyÀï0P¸´ª+u:^F˜¬”<[R¦ €+>7sÅÇÛ±nF1W®Uƒ°+aÑRäŸ6‚¯)ÅÇµë	-«k†tT¬fË®êÅr/‚„œd#¤í.œegóÉ{(éò&N³Ôd½Ó,yÅÎ,ß¹lKU²Øæx#"9Is€3#ç
n\zŸeS$óâ ÓmtÌ½.£Lò‹Iª'©¥RGmxAå÷@*M]3ž2â)½5Ð¢™(-¯-Óv¸­ŒwÈ¬ìCó'éy1[M*T¶Á9§]™õW”¡?¬ßö”¡6á­ÓªRÉ°\œÁÒwÃ¨À]-³R,É‚@ŸóÊV93•¤i/\ÑºùFðj¦YzÏ)5•8ÉXMxGï=`x³ÜñåÊ:ìaR²ÌQ:ÃtË”
‹§H¢pvþÑ¾ Á.Ð§º%ç#çÂ€—;2¿ÆùË[ãÂƒçšÁF [¦]U¹<ÝÌÄÆY•T„|ÏOÂ6 |¶Ý2âX9]l³$§>)áxf5Ç‘Z>RPÓP¶ÇP=xç«ÍŠ58êà‰Ñ‡ôci"˜šN\âƒr¹ÔÓ	 ?qž ; àþ%NJ¬Æ‡hµ…,w_(Û¦ew–å“Ør|KÙqw\žJ–a0WoG8ú•³™e¦vŽÚ3º~ŽF¿ŸŽŠthN[ôÓÙJðÍ*9™n-J'ù¯IÖ,;ÁÌ|¹6I»4¹úXÇ1G@õÝ‚`\R¦èqªdpÆtÏÿv}íÁúï¾f°(^eY„¥x$ã÷³‘vrÈf÷ÆÀáOÅkšÑjò`s}5Yï?øÚ–Œ+2àù¨gé×¾ü­÷ÙÍÔËõ¶†*;ÇôÇƒG°¥FÊ¿”â,$æ¾¡­6ëÕ
ÃGù¥§ÂK…¿*,ÕÍÆR˜%4mB2%…Cí›„ß÷D.û9‚×9ô˜ùLt–ÀyVÍ@tÏöÒ¸TÆ¥1-–?QBÄ•ƒâ$v„+ä€±–B®ìFÇ)v\ù–rJþòc5Ô­ž®©*`ÑT¹úMÄÔvÖ‰ÓO.D³	ÊƒŸŠY-i< ›N
ky\œ$º"Ã/Cq´R×ÛçFÝûËÑœ¶£‰P§À}a¸?Š$c¦;C*×Dy~m*S¿qi
ßæÙ½aûZeï˜Üüòƒ¾Ý¡‹Šå!äY6íi <Þ(lÕKyNFóòLº€ˆ Š3¬ÁCöivCd¨Î…êð‚&¾.æU}Í™Ö²‚Ï]é¡L1—oU©_æ?é”ë–6Å¨u¨ËˆéBº8§TÙ:zCiŠ¶ˆôÈÏNNòúœ	$$Ò®&êr:¿ŒŽ Úž[6˜.›kz³/‚cðµ¡7*N1Ls>Ñà$S:³¹GÁvH°IUHÏì·S.·U,ÛJå¼møª`’£#Q4kÈn,³]Û‰?F	-«Ç.
9tÊØµ©]õÓ³†0‹'4<ƒ‡0€R)áñï¶Xz?Kk4¼8¶hjÈz†Ú•€âI)^\‹­m·ãÒ(~¦€¸?e{–s?Cv"Y¿:ÄyZ
Y¾Ež·îq`ýª7hJ¤Ô®£±œa^Ì`SÂNfÆjRFCú–‹=ê$Ø‡díêvV•áÞ=OKz?DÚÚ5¾¾ÞÝq43G–nÎœÀàz˜‰}‰	éñË~:ƒ“vDß;Ò‘ÔqíÖ!úŽÉ§mƒm@”ñUõÕº·èçñæj"[¼±š`×Ÿù„ÒØŸãuä•hpÛÉ&ðuú™ßv²Ï×¹¶ž¶ºz‰–þž†iB_„n!Ü %<]Î¸Ç›–ŠÞ‹ì$yƒÛAÕÏê sQ9-€ïê±µ×ùÆ»È˜Îú²€ë8ƒ¬±!ôîsv«}Î÷·iÉì_SRlVuOàî¡`Í×:™÷”œ‘kc`©àØ;ÈÒA%ý4¤Æ´»>.ÑL×É&ƒpMP›ÿïS&'Y6D“~OmÄÅ¤,†AJjâ'<ñpÊ¥“ÓLD´J*ùÓ<›gm¸²›‚„Äº2dÃ!)ÐL‡á~±Î~Ýk…œ[Ò}°Pé!Ï1Î'ù±a^¾¿7ÉªÅì}²wï:0ŒŠ
²CàURÑ|wlæ•æ£jXM”?e“¶ÏDQÖüùjò@r¾X/Ç³-"ïg°Û†üªoÅ5øSwŠJt¯Ð*WJŠ™Î*éRn*õËéN³Îšñu¡hÐé¬„­Å¶è¼àVjG1ˆˆ°÷ùagjJïu:nèéÞ¤•_¯¿YM6ÖávžoˆçÉZ²á½ÛTu6àd@Ýj\‘ƒ¬ )3/_¦/»¦zÿT „s5¾<+>p°vWÈ]&Ÿœ_ÿ<ÊñîÏj²B"Ù
ãø²nA_/Ñ¾10Ø€ŸŒåN„Jf%`´I‘ÂVöJÛ¸e]tMVÒ·€”-R“jùßX³¸ò9#Ý‚6·]õlµHà4‰ããŒØtùì•ÆH˜·æ‘Ðñó ŠhSƒ?:Ì”$·<¬°usÿÔ¹¹o]¥zh­ÕÉP? eºèO¿daÎÁbB¼ÍÅ5i‰W^È7,’(FXÓJ÷®©1ð5#¾m	ª6˜mÕŽÆD
kr<Æ<Î–ûv-œ²á|xÜÆe#Ä}–´M²uN–§38eD¶Y†÷®:p‘Ë¶ž³ô\:’pÛ©øí¿~Uì¾’ÞïFUv¥BX ]sç’úøhPÖ¼ò¡Zëli…G°¢Xa¬75æQ‡¢Ë£)IÑ0¦t#WFý‰})šSä¾µ²ŽP•ˆæôèˆ°Muÿ r•ê†þþÓoQÖcLËá‰Ì²ÃLû 6VDÜ—Éõ‘@.%DÛá¹é3aly”ÿXÂˆ»–´¥n¶>FÔ8­4–^‡Žg£ ¬†Ã{/^ÜûŸäûï·Æãu<AÏH“¾-&Õ!°³ÝMÅwÈGÙEµ[Œ`] kò Cû®?è=D’ðlŒÂÌ0ùj}]×À„AÝÎARÃõŸg9E¾ä)?:T5Aëñ%,q~
âpg€Nr³ŽÞ$Á‘mlÆF¶±ÎMCx÷”Ü×s)wõNôyÿ[Íè\çÛØ)í–j½)¦ÎŒôúé¾§i+¥ Ø¿¹lÕí Ýwªžšâ—›zKhÄ|<NgN¥“òC6R‘PÈ<æ+ê¥à6Ó¾ÖC òÊynt%þ;R¨XH]œ:REá¦&f‡Œö½¬÷›,Z§7iã€tE ÄmDzÚ|ÒxéµmVX…n—Íî^ÒÕ£âñEð<ú6¿È†ÝJ8¼Þ‰ìÈ/cxÀzZÇp˜œ	iØÐ¯þ£Ä‘ëLÚÝâÉMÉUàT>sùHFõ2HòFéw"Èò¹BuÖ]€ÍšÏ«]LfæÐ¬½*K<¯bÓfî%ÙèÙYè†, f«RY¼#iØÿ+=¯²˜H®&Ç„TÝÔTêôÄÕ’}tài÷˜¿«%w_¹«ÛÑ1sõ’­(Ê$ƒ~Š¡þkâÿV/°RõAMy÷A­¨	£ýZ1îlµ¾ÉSBû}®º§&Ñ½]<ºþg•Ôwi„K¢—sFF)è¦OA;/áçc É`vý¯D»s¦¾‚ÌÍËXb5úNáh'å|'LqVý—-¨hXÌ³,aáõëo$¿*¥~Wd¿¿Ò  
øæiâ¸~Ü2ëÂ‚œehØíœÎòaÇîò°ú8¼0E;˜ß™À;vñ6Îé,§µà@‹}¦DðÐSwýU!uã¦¶%=MWµ"¡þèHø]ïjÃÀËâ¯õÓÄªj[f¬†Â¯e^Øq‚˜ÍÜDÂnA:;M¿É
SÛS	]À¹Qfƒ+r_1>÷APãö™(å`/ýCC½”¤½YÛá^º6·Óœ€çÌÓžÜG›ë=æÆÝ=fÿl>ààÑ;h§BÇ’t˜ányJ«"6üÔé|ô®š”×?¸iÇa&kK©MeÓò¡ï­…w—-Ï‰éž ›$Â¶˜WÓ9@÷xTÏgÐ@]¿öTâByFwü‚"‹¾?|!íÊÝót´ÁÚï‘ÜŽìZ›5,50	8Ø"…+ÅšúL6±{x1~¬bì2Š>S±d0Ž‘z¬^5âÊ:ÄðÜßÉt‹ø*#`—_ ®Vd¦6 a|‚Eµ¡ç¬[àð1•ü; 't²UèÖŒ­-£6Øñ—«%U¶¬;ƒK2¼*Ú*„G›)„ø/™Z+ô.Ž"öPDxlÌJq¿ë»‘¥ˆ¿ð"Ž¥Dj'ùŸæ™*z£‹ó4!iÔ‰M#1oéD¦C—¥"‹ †	±5œllgBÕÙó—PÕ¤º)^§ëÆ÷‰’¦93Á|+W‘F¬N§bS«§ÖT
èë¬ð|]T>’0%wWNÂ£k°T³ÌVÓ4«¢’O¶®¢“ÁÅkØ(Íü¹Øœ{®<’º"fl¬Ö¾9è ¦Ì&eŽgp'gç ÛÖ‚ÈjòÚÃßU³ïWMX¦U/ö-zˆ¡=ÏúHáç¶é8³ª6<”ÆØ—š< ÚÎRJÝ–ÂK/kõS–O,ã‘pT6y25m$USÔ¡`d¡À&µØÄ‘€"pZB¯ÂR­pû !ËrCNCRJ{H”é[$Ÿ±¥÷d‰ºÄrÿÞ²µ&²÷âq8ÇVCœªHp.ÈxÃ`Š§¯¶‰ø*ÖjU×
Ôn1K¾Œ3¯HöËÚg±_AZi’BûH©ÓËXH©b)-uOÙëòð.«áhg†õÑü†ÆsõãnP]µ¾l§Öþ‚v@dò€_,2ºeí“Ä‡k“Ô!M|@«©Q¶¼;|-GÖº­YZ[³oMÓV)¿¼Y8B«eb~¹8æÁH ¼Ú—Ë@ûöC[ç9K¢Ys-åVh	r×Žb/‰5´s	YlR¸Ìä3ÂªJ3ÌyÜz†v51[»fIVýéhVX´+Çýh˜Ÿ'ƒQZ–(Ó>^)§é [û¸öÕÊ·ÌÉ(»7Ö„½Tÿ¸vœU²lÂêŠÚO¬U}t¶Á›C½ßÚý‹Í“jíx”Þ'ô¬DÏµ¯××“j‹^«òÓ³jåÉ7ùñ(/ªlJàY6Èénô£{gNoS¯3Ñðh˜ºD/ÔùxåÉwh¹ø_¤f4ø.êå£{Skb÷œ™átšN×6üºöa–NWœÁå“é¼r¯|âµ4¨ž²÷U>|¼B9ÖD¢
¯ ëÿ,a1Üé`M«Ç+ý‹Qy±šàŸ§H1Ù%×Ç—Â»Rä­ tö½çLç›yU¨jÃ¸YÐÈÅ#yV°ÅQ>xÿørX|˜àÅô£l<Åe¸Jª¼Áh¿IóÊ^³‘LMá Ú*k'¸Ç_n|}CH^P{$÷ÄPn<>‘hV”î ‡déš‰×ø½í(eý§Xg9c2ë°Ì÷ÑÉí -|ýæãÞ°ÛáxÒém÷X«Û»òÆ+CX£Kã¥•h3ÞØoQn—L½ÁÌØt+–Ý~®yn¦ šßuZ6ÐœT>€…\ˆ‘­ý(Èñ©øþ`ýÞøUÌ†ÙL<Ù„·gÅy6ÛRe6ÌVî¾ p¤ü(&kT÷—ÁA8:vÝt6âVæÎhdHf'pËŠ5	nÌçÊ=³ñ#÷ ËhžÏ’ªÀMÿñóLÑQ ©¥°§ü[VÂ~JŽ-r>§Iyýçä<ûÉ¥KIââ~ÍÒòlSáïW„¿Ïs<]“#˜O’&fö>f±9Izî4/ã:tµðŽ5º«ƒöÕºSN=X÷ÂBx±o7·E‹Ð|½ÕÔEá‰î=ªÈÂp0Þ`æ"²ê¸fÖªêÖº—IÀT@×jÔ…¼Q(7¨DÄT;Â=HO¡Æë7‘”ðK€ôþh^Úp¦”j"ÁøÑ=ÿºÅu,ÁMÙ¿æF¨Ù
&q¼f«%Š/RÓ†PD2´š-±à¦ˆm‚Z¨Brcp›g§ctòò¿‚&/(nx»#º'¬Ë(É	§Û¹ã5
n¥T/ÏÂ«Óv;iUÍ_ün:4™£­Ít¨¹Œ_g7
ßZÈ2­ö’9i¤WÐÛY6(€á™DA¿-ç§f¶*U+éO…z('æ2ä¼y?ÍuE©KÏGH.o¡:èË²Ô;H	96OÎú7=ü>±®~îl§‹ÔÃ±m¾ ó2#£ßÞ„¼|j˜G±Ú±~à’¯âBÜpIp’Æã€J%þeÁU$¨ßCÁèevzýçABÉÈ¹bŒ¯¶²ÈŸõÚ?yx«/B£†ïøxí«’¼{|±†þx¶>ð8 ‰;ë‡áôvérã²=3¬ËÁÄÅ•éeúqí¾Ô,£¡T,Ž¹À¹­Ô|Ö6“gy•	5è¤ •î¢V\‘MeâN
µÔ){°Ø¬ÊÑ@P<©¸ÖóK}"³‡XÕ‹{íC’Õ£{Ç–Üh94Ÿüi.“H>™åðõKZcÝûWƒ3´ŸÊbùÂ%­†8Ö?É¥ÐÇ§²âÂÙì•ý“\3¼Oh-É(#\†¿,.“Àu£Ÿ3\û»¿KÜ7löŽ¦ž84pÉ*kâÆÚ4d£••ò>±®MÍaXÐÅ>ÁPµ^®m$ãá–ù¹I6Š/}#Žg¶¤§£ô8ñÖG…²W^”âèµ™OAl¤e†#ÞXy‚ÒÞ)ž3±Ý£¦]
3oƒQâÃÚÉðPøKDáÍ„ÜdÃµ…¬>«vRæå–Ô¬QnOñÑlc~(d|°¾~o“Œ„ö¾~Í)£9ì¯ÒåBfchìfj'„ÀLná =–6b(¦¸_å€VVž - ÄÔF‚î<º'
„ê^ú¾¢äK×x`˜ªí÷ÙÇÇ—øðJÍ›~<t?º' äáËžÓÃ_•°!¥÷¿V´ô|&ÅN¯å!©°œA®ÃQ§*9
¡üIº»Åx>NR4Ä¥e¯Õ+" 0bÞ ªû±¨Æ~úñDü©ë´w»Ñ$Ãö®ÂGS½ÙnÑƒctÊ~Þž#—ÓˆKÝ.$×ÓÏ}b‰ÖmÏ‘œ¨Œß)ÝëÕ5öóªó‹
\¶2Ÿ›9GßÒEÎ¤]ª÷mö0kj)·í«[A==mmùÀœž–Û"me¥¢•Ã¯ØÐuÛ-ÆL‡¦½Ðn§A›½¯¬“F’Úd
<³"ê›ªŽæ…ò,•lªü=º¬õp>£à”t`ž¾)À ãUb‚/Ô=‘7	ì¹Z¯33 éŒ{´Ñ%Q·i<'}±¦-ãñ‚Ta#ROQ »¼6œ”÷ñ¤ÜÀÃìh*œmâø¦Z:x½±>½xÃŽxµtÞš^{Î}ÈqYà<Î«èh^	†4Ü{8ÞXõK›Ðuþàs³ÅÃN:í ’O†ùi¡á"Þ5dÜéÅ@Ô¤ÀÁ'*EPµû%áXR€P”W×¾^ÇƒJýZ§l£5±ñÔ‹ÛíH>†}
œÕûµõ8Ê†äqþii0Ã­eÇüfM~v=Š<5m*ê8rÔNÂ^Á«õÉÈòÈ®-ªÝF€ÊŠ¯,ëZì vµNd§
Vzý¦¦xÜf‡Ÿöf8j+¾qø†_Ûè?0&isŒ¨Í<:­Ñ~Xº×³.Ìge1[£hÄÙ,$+HI—6ÀÊÔx*ÅÊÇ)Ç#Ì¹„a”içÁ•ë7Ë
4`Öû;Û}ÂpûïÃíq{Ôp¼Øn€¶ûÂÍ1#z~Ä_œÝ·N•²àkaÐÛ\yr¤ ãÝoÃ`)…]Œ
8é_”|$Ž³~è—wö5õ¸³úÐMI’Û+yd;þ tu ês)©ªààwZŽð=m˜Èd\¡¥ZÛäWÚ`ÑŽ XJ÷˜ó3&Û+îKÈöéé•Ã‡m"¶Ž#àê[eýµv/Jí6ðô÷¾2ût<Œðl§†?{BÚFëWÞ²[ÚôèK’ÿ@s¾ûÒhDQ§Ù(ï7ø7U‚ÝiKUÂMÕÿ¾•V´Ñr	 k¸9¤¾(ý¦ØÚrûJ:°©¾áØË†¾š€©Bò|ì.×—HøÜèkÙð‹îè`på•ª{8ààÇ»8ý ÙíG¾]}ã"ŽÑw¡fì†¹=yÁ5Ë&?Ì³9Ú81v5§x­VN„âìÈïÞÒ4 µƒØŒÞÔcöòîâ«Olcë‹AªÐ+º{Q,J,—ŠÿÛÞV³»…‚÷æ*Þ’wy|KÝ’Žß¶Ë¥\Êép3x˜ 6?`xÇSE'­÷eV4CQ!iëŸ¥¥xÕ‚6ÓsQp~¨ƒgxùX+§¼GK¹ÚÝ—Œ)d-ý2ðÁ¹üÍ?ñº~5‰¨ü¼D´==Þ&q¥æ¾ìFô¶¸û¹t0¤V©.{ýÄ‘©ÔZ!ô\jíâQ·ën»Þ©gÄ:ÿÄÝâÅj´6zÈ9}’¦	¬Efþ,ž€UC?1PÆtK¦£H4ªJÕ‹c|AZSy4ÌOª"ÿ å‡Y•ŽÎðJu¡… ¦ê¨‹%<—Ì Ø0+Æ!ÚŸø5æžnn|rë%8ºe­EìQê¶KEB˜6˜ÔÇ—oc¥Rc­¢ø°­Vêc®âqÅêmWº5aÃ¢fÚ˜±Ô§Þœ¥>‹.tÛ%jøi"Â—f75Øg@D&éèy>y/5£_×Ød­F²\c©–‚zwœÆ|WžìÓ%áK‘;÷ªiDãùÄ¬ç4¤&~$Ì!bÖt•Û­ÛÂš‡Ÿ†×÷î%ßSÐ dZT˜õ %{OË˜‚…S¾v
*“ˆÐ4áË¹æƒiá IG{& öã„2M#Ö- B‡RåÞrŽwy.rÀ"‚v7*}4U»ò­]öÇÐ|ä³+ÏÎ|Zœ*æ‚´q›j:*K(Ò¦ZYˆ*~RoÿÓÊM£	âmvoÀc^/o›úžsgØ'æÆÇ•C.x]†Sr—ªÐàÖAeîüXP.wt0|º˜´—ü¸n·Ç£öîT¼‹j™¡ú—¸ƒ$t3J§Ò7DÑÐ¨CHÜ]?—ž"A:P8Z¡–ÞÔ]KŠû7õ¨Œ»µ‹ðŽÝ,Ü‰›|9®ø©u· ¾¸ÿ°Û…¨Åö…\5|ã¸§¸×u=5îÂ;€µ¬¸v?‚ÖæÖå2¶Ë1µ^õ~”%Íx©Õ¿(/—»pk¹l\¦)´qÚK²z.Åæy#‹gÍúh/¡˜ÑS’,r!
ÈüU|ˆÚy¥ZD4nDéÂnD‘SÝ…0LýAÒÂÇ(-.æ~ëÚ3]‘Âwã/U¿ÞtÕv(ðE,Ûû":ü:.6­Ó˜ö–’b@Íî.\lÊ¥ºØ4¦XÈ¥¤åR/ÓÅf¡%–Ý•7<SÜ¡{¹›”QÁB'7Nä"¹v:)?	§“ø¥·¿>ç‘OÓä²láúÑÆwà²ëÛFÏŽÅjü5yv,2óOÐ³£6âÉ×VÄ“ g†§nÄ3Ôuýõúj,DËíÏß|5"Ÿ¿ùjZûwè«QÞÊWÃQ­*ÇƒÍJÕ[ûpüÍ¥?>ÍiŒçíV%rS¶2áZaÀ›i/Ï3n‘ðÆªSVq‹¹ @e[êP^ö›é‡ÓFa?jnÃŽsçOš[L/û¡çÍ­¹AÖË¾ý¤¹…Pèõ²ï?šÆ½<~°5±„¾lOå[y8|¢vå»òPÐÑçÿº(Ü!¡ýD1én<–GŸŽ‹BKÝL›x¯nÌX·½Ú)ž}ÙZöô¢mn®<Ñ‹’˜SðÑ½³/k;Ø,Æ¼§ßºÙ,/ƒ§íU$ÆEŒo9o\!&¿#³¾ª£ûÖÇMÃb0Ã‰%µ_—›P?ù®·5;¬7ÿƒõd
oyJºËˆ}Æï¨a³–hÇëÁÊ\ÿ¼LÓºÄb¶ûù-Å[	óJÄ]¼¼·ÓiBÍfÇó»†Ïïö÷4æ˜{#Ù\`kØ´éÔm¡“Œ8ªø£IÚ™”²:®Ðb–•ñD™µkK%7RuÆ~·Ó(iÛ¯J6°rßpü4‰ç¬/×&3c2ÔBÚ³6Šõª¨ÒÑ‹ôB:jË€½4¿Û-çãÕd&ÂÏÇÉIw3ùû–	æ£×[MÖ[w—¥³I6lÝÙ¬_ŠY†­/ÔÏt€îLzrOÈI €zÖ'‚ÙíÊÜÓez0/ –h_oî¥Öðdýnš¾œF—ß ûÝ
ù* ¨…á¡a.WíjQ4m†úižžN`åƒsìTùäT$²N¾’ƒ¡K&i™¨­F™ v*8Ëe©éK­´ê¾Ao„ŸÆåC/…›«öo£xªs ;Ý™2Å£`úØ’PãsVþ\¢NÐKµUÀçb[±ñÒõ¯ûºßïÓ¸VøF½½	vÕé§jÑŠÔcÁ<ïþi±5
HvÿFÎ³ø¹ÔÓhé5Kc«÷œ)ìÓQ¦Ç¿<ÿYü4úÐRŸmŒr·ÙºQtõýc[ú4xé˜uòï%û5øÜµ·é¹¥:ÝÞÛ´Ñ‰hYîBW‹­Ñò¼MX›OÌ+¨EH7ŸèÿÚ.˜¥Ö”Ü¦MO3J´ÍàyS×Kÿ8q]Š¬Óìž£\ð=£x¨	¡{q ÐÎŒw?VßÎò²±žœá?\×áD»àG)«¢–?õQÆ çéG”CÅ9º¹ÔKÆ"·¶‰«”M1¯L¶RVˆžÃLæJ A¿ÉGyQeƒ4·E´Û•'»€íe5Êˆ¶Ç,)æ@TÆ”?õgê«˜¦“ü§”°?§¾ußl…hÊ¡ ‡Ùf­‡f“Ù9OssF™È›ÔLw’šüÕÚš çA‚æÓì$*júJé»A>ªÒ9¦0° J MÓáŒ–§ 4)+TLŠä8b‚K6D’‹$u Bg•í>Ø	YMOrÁ÷E¡9µÏÉš,’È¶eºßÅ’ý¶³Œ„3×.–æ7–xšëYÁž¯ÜA´ÍÈ0ÈÅó~Ñµëæ8xx÷Kå;“!šþSÂkÓHs4<Ã´ìêbÒ³ÈºšéIÄvú¡­ŠÌXÇpB«~ƒìx¤Û²e·~[ú+Æ|Š4ú£øµ²ÑYoîûÁ›Œóq³*a_âØ1Ü¢Avƒ\¼Kñ‡µ¯áLÿš›/-³¦c‚iœëüt­<{ˆ4æ£a>9ÕFÈ¯Ë¶DK%°òÄ`\»@»SŠ‡­xÁïÅ¸BmG£âÔ&O.5­¸Ò¤¢‹âšõª‰ÿóßÿi‘â—ŠH\†¾‹FÀV6¯å©=q¨VÓU³²ÉP¬hÝ&§bÍD=Þ¨¦,m”ÀËš2aì+­c)¯¥Lƒ=Œ""*Ü.TÖÜ2çÄ´×ô¼‡p&¸­ŽŠƒlƒ5‡c½¦ÚpI/«Å˜¦õgš¬z9–Ô¸•Ôâ:ú)*ø%qZÂc¨ýcL¨.ñ«#žƒF¶?ŒÄEúm¡¢Läq×X÷”’n†?Â™éSÂÛ™Gbpº1
Õ¥.»•†7„?Ùjäð…ÖŠ’y/ªºÚ´UW	£x¨p©ãí:‘MÔ‰lºÞN)2¾XKçUÑF%òeK•Èÿ  ÿÿì½ëvÛH’ üŸ"‹[½¢zDJ$%YÖJöÊ²]­i_4–»föøó)C$D¢, ”¥âèœÝ×Ø_Û;?öÌœ3¿úü&û$_D^€7”ËÕcLKÄ%/‘‘‘qåð§ˆøhüàæsW8±K—Ô™B(ù¨¬N†LÙað„r*úrŠè»ñÕš#BN$=	AÎ_HˆÁ]òc<üüodü|D	Iç¨KH>ÿut0Í• ê8Bá_dÚIQQAxäç¹þ<Ç%¤.ºqW©2pÕÙ4Q ÊÚ<šíF½„Ë©`~õI#m‰³¾¤©ÆÄh«µ&J½É)Àë‰S?Htl\|¶L_L¥>Ú5ÚüÕ_Q¡-@Ó\‡:»*äþúmZ€ÿxå?ÿuÄÄ'þ˜¯ô¾TÛ™è4R{Ï>Ç	ÐÚ„øÌ6˜n‘zÏ|þ+uŸIatkÕ»’F^8`FE]85ï[©‹Š²¨µ±1]Ny5k@P‘•€|/Õ…'”©&\ŸšreéEæE#/Aÿš·gç¯/~zöòüÍ³‹“ns;òÓ¢!áÐJ¨2†	U¦O½R|)]Ä‡9"òN•oèµ«l@ˆÝbD£ÂO,ÕˆŽjô…GeW5×$¬mHµ194uºà5k‚×¡vÓë³Ü—¸©Ø!„µq~:(mšëÆJäUíJä…Dõ–L?!'™T*ö÷æŸh=âüüÊ9 œTø_Z‡ÎõÙâ¤¸…£þ¼‘öœ÷•w%nç{\Ö>°=¼á‘òÙâ;·=ðM‘ÿ•(ò9o—ëòLYåË‚Å~ÕìÃš:ßÄsW¾ý¦ÒÇëoH¥Ÿã×ê/…F<Dù+C£/¬Ù_1…rCçKñºtŸ6zGE&˜Ìg#Ô8&þÈ†7òàdÈæ,lã6’v0:$,ÉòüŸ’“,¸ö6È(OÂøç¹ßmH¸”%·9VÁJr™&•zIÛpë¿ÀÿR¡uˆ¿†]ŠQ´ÿ<&ìŠ~É9$ÃÂIÝ•ç):(,÷Œ—÷ÉÄ$ŸÆÃöÿÿr‹laõ§¡4¨-èwsú=J‹p‡ê”á„´‡a<=K’r¼‚3ý®Ÿ$qÒn=Ãÿ°Žt¢½ yIøú°µEòf¤.þSù¿³Ð»½˜1!ÅîâïæC¡ãpvÍ>ãmÙâ<ô¢X(°ÒÇ€oý4ó)J¼ô£³»ÁŸ'ÑœÿõÔ¿
¢€¢Í="J> /‰+R§ëBXøè×EAcTyx)Å5ºŒÄcÌã;½ô“{Ä„¼¯/‰	R§ëÂ ’‘w»6âÓý$kx³¡ºÍŒˆ3Ëqåñ÷‹|:TŠJ1ðô8ß}X	»ŠY5Æ±a#o¤A±kþôE 2ÈÐîÑ*=~It«u½.¤±ÁÔg`6±5£^Ï æ‰ >íôóÿõ
DôJ×Z	ãªS²áÝ„c;Ã#‹yâŽ¦A$!I~nÑ ªòjó	syâƒ”ÏÑÜ=§…Þ†4\9÷µï¶r`15m1¾2æ°qþAûøm{‡ÏÄ‘	y q>MüòøSä'g#@§ãcøÇÝjåýsgåÈ›ál)’Žý04mÿ\Bðïð.t˜ÝÊhÆçþ&2†É%,¤ÀŸ£ÀU’?þñopÚ$òú7äñßµ$ô’áP O¾ŒëC‚'Þ/^|Qhá!›×(H)QÖ‚½ûáûEõîiÓ%Þü ûwã,ó¦°B I@N¡§y˜Å@´Ä[RJ¯¡xªè
µø¥¶¤ï`­1<ú7}; ogôåÿ6ÆÛ˜œLúðÙã†@L]æ' ¾¦ˆRÓY{³x‘¯ïa¾°Å£
_<:‰c“z/Žð‡xOØ{ð]Qží.ßöy¼ÑÉ›3¢¦öÿ‹ÿ<NžÜ‹¢-¼@S+µy“£ŸOèòÔgÈÒ—‚î&CXjÀU‰ü †Fókª¤C‡Ä‹n¤ˆx1PÑ`H©	n~yœUÇ-¾nú>©á’æ’€/»°þ©7¦‡ÕˆîÒëÀÿÄBð`…Ñ(@0szÅMBG/Ø£slŽÇá4Ñ?P:þ¯µÆ†ø1žgyò|Q°‡EW":a0zá,!´Ü2rÄ®½!Þú	ùMh¸Œie“ð–Š^D%Ø)~°¹‰Ÿòî5#Æv‹Ä81Z•q{›\Ì/1½ì’ÐÉ0Ãü( ÷ä·žÝÌ–'V¡ÿAÊ(„àH¼îÃF¤F¶(ÛóÂ¿ÊÄb¼6lFºuÄ¦Á¶|s¢Î\àÞs Ë»‘Vlz[>ß”½c„/³ìú]NaèP‹Óã£þw¥W«4ÎHK­w)Q…I£ÅPø7ˆ¯%ITìØÆC“Z äe>¥z]rv$[ÆHŸ`?üuÛ#Fâf[ï¼±PË|–¼<É
Ç·YaÄã§ÁŒßv³˜¾öØéµ6-¾[žþ0ªÌ>Œ‚«+<Ÿ`<òðàmh¸S4¼YÜ®´Â—Z¡Én†~¶ó¦·I»·³³C~Oöóú»›FŠHbóeÎ>‚í$eh-Öö7õnÚ;6lJo2„`V>÷
„èwÉëØ®OæR÷Âá	”å$;qtŒç0"*Œ1yÎ§T˜i¨B ¦¯û‹µyéÏ§øfu©9ö‚ŒûB5ó˜vˆ4}Â55ð`žDF""Ê\H¹h¬ÄBùC
*Þ1á_i,ZäSá ž<X-òA+Î®ÂŽugìÓ!r±~¼Œ£y$5¡À\G]†ú±–îªºät~™k‡&¯ðAúO€9øHÓü|ÇOäLL§/c©|…&}0QôË.WÉRì—´ØòØð<yÆ@‚¡šbz]¦œ³*Ïö\éS±?*^75ßñ êL: /îâ™Û­Î!³²Ï¾de=:Eâ¡rPÿÄË…¡«=w½)Üx˜I©š‡¨îäÒGwOøG¡r?ˆÜh%7ªòj/}tS=ƒ
¢¨4#: hn³Ù­:|Ö­"“^ÍÐF‡c¨ÕZu=Z”q$HÆj|¿(bFyK†æ£øM9ÆD]q«ì><é•Æ©Kî±ÇÆt`qÓ¯Œ±ôŒà/PòÖ›bt<ï’sê7
 	à@HÈ<P^‰@ÊbtØ‡ÿ]Í#š;æ@zfpR]y×q²8Nµ%Äƒ¿¢øÚcyHâ9&ÃŠ ó)ˆ1×1G÷Ce ²‹¹¤	Agðc±ª+´À¢xýy:ÃãŒÌü0.ËS]ò,Êà¶¦t*^ÆD9¤æÐz¯á,Ç0óŽM¡Û4ÈælÀÝÊBÍLÛ@ç®VY…£…p–­Ü®h¹0Ó–D­º±´1éÞ¦a4RèÌ¾l4.…Ïä;oºp,þ’Ò{U-Êu‹åºä	9ÁµäJªNøjËiÕ¬©†ŠÍö^…Y0Ž^Ï³6”IZ:XUý>ªÁÚ¹/z©½* .¼ A…æ)¢^yjŠùWèÕÑ6’lñ»f¹'NSRpßG‹|”TÚô¢`Š­¤38­K~m–Ÿ$€^ácr–¶G¥©k§ZR¨ÞãºÛìÜ<(…ÉñíRœ R¼šåíí#ÄöAnû•»&Gi5iwÍ'tÐ_æíëŽÐ‡y ÉSÉõŽpC~šž|Õ,pS¿Äh3¦ÍŠÇ¨7yL6ž{xjÐßŒÄý„²…	ªSè%@ð¤ÝTòž&O	?=#xÄðŽÐFÒÝ¨ ÿ©2 f`¨‘õJ^b¶GôœÃƒ›¦T,ûJãÍœØày±óì÷S…ï¿„Å-dóNûÑ&Q‰#àÍAv8^0CÇ3Ô4Ÿ@ƒwÊ£M‰Gªð #9jCv;ƒÖ¨’»î—¬ÀÐŸ ùÉqy¿£~ñÚç€®tÚøFý„3 ¦ƒ©eÛ¾89OÄë‰à% âuiCNGè,÷xÔG¼Â]ÅÃyzˆ†FÌóTü(íðÛê€øyÖ‰âzÅ4ºö!`uh$þÏsd,**ÎÝªô@M×mPùã¶-²P~4Qd“Vîœ7æ¾xâ‹ÿëGœa/¨{¼Ø¯ÎÑau9÷Ä–.¥›¼¥‚Ì-#IVæNÅ|”K$ªZ}CM¬T$u/ô5åš6ª$MÅ7!!«“eE'f­7Ñòç¸¬dG«þ¸lãðU.˜ýq>è„‹&,?Ñ¦!;‡:ŒîN‘éZ‡µÄ¾õ€û¦ØÜOëÂ‚ÐØ¡÷§tþù/‰!¸®ú ²„†B´5¶õ³é æ=¨uœ¡:Ð´Æ¶Q€©½áû5æVž“SXº˜^…¹Óq´!eÀJ¶U)}=Â¾¼.^Á ªÙ49†3ã‡8‡þÜ’wŠ¢ÁE§è|UXÁQz=.7º\ö^‹ ÝòI|sÜÚ!;¤¿ÿÓdZõ²	&K[Ãy\dvŠ”­EFÇ­—ý~woŸôúÝþÞp§Ó}pÐéî<@¿íA§Ûïàí?ôú×»Ýþþd¯û°?„»ðzw€ïì@Ÿø"¼Õ#ƒî wÝï>x0t÷ûÝxåaô:»Ý»ì¯ƒîÎÃ_4i­å¡þçÁîÉÁÞ€°þ Ú|ø ¦º×ÝÝïtØ^¿»¿v°Ïö>Ä'p5€îìÃ³=öW¿{°Ov:{ÝþCÛ ³ßííÃØöèw{0üƒÝÓA÷áCÒß›ÐÁ‚­`ïNc~þäÉéÎó4Fz»0]Z¿ƒƒêö ãû@ô0íöpgw nüø àOGsŠ·ÉAwoÀðw 	üo?Å»ƒî.Ü…—È.L#„qã×°–=è§ï4Ög'»ƒÁ^ß½îà`ØƒÂ ¢Å..,ÜÛÝÞ^ÿ9í=À¾q¨89Xüƒ°B, àÁè n8:œ|»¿O´Ãî.Ò>â
…:Éá¯1ìáëqý6Û€N
™ºQË£9c¾ëÇ“8Íj]BÆ‹‘îê pãYX?	5´§"‡JUÇŠÀ£èÇ8D7/–y™PÒgåj$®ÖATÒ†(Cûô
zZÛVÝqšWaü©3	F#øhœpvXÄýÒŸùK×A\†âö%Z@øß¦€`?Ø0dMÈbû÷ä"ù— ¡×ÐZèÝ’ßo³Ç8Mö@ªb6k	˜òÄƒ”Ÿïù¤arÛh«ð†GI<Ãòì	†ÿÒÙÝÁh{6ßò’Ö±gò°ÚÔLV’D¤)ËÝGb®ÛÒybiT]²Nä–ƒù–3ç
 vX´xÃÒy‡Ëð2¾àý?øÊIò" æ‡Ý•™ƒŠ bçã-É ˜;Œ9šÍ_=Lf˜š¡NÁžè„O“-LÓ¬)ê¨o]¤vE’¤YN´±#Fý¹‚T?ñz“lPÊ<¡æ³Ìó(*€¦dˆ*|!ÓÝ ÿ‘K©TŒ~D¹7Æ“¬jÖ'”s[dŸÈóD¹\ëÖP<¬R”åì›<Ìøò–œeŸÿ7ºSþOýTÝ«>%ré–º(§mÇW¢K&ËBF]=Cd2l¢§K…@>&GÿTJaGöÑK?š—nVÅêzÕÉ£m¶ãK|½:ÍÙƒ›0×3£	Ã“ÿ°]‹4ú6C›²ˆÏÛxF°xTÈÐˆ
ƒèíÈÊ4µ_¡Ç%¹N»p|½p³BLÛ€Ü1RšR…j6É•´jšÅÛHQ³Q¸&SÒÊ‘ÓàÐ`ú§žœp©V0O-à\êƒ]+{€	p~Æ½7FÃ T†H¼Šy,†‘.¾Ü0EþÊCaÕ[µkå¯bÙv¸ë)Ã8]éå´˜wc¥óÑ¼öóVYˆ–Õ!Eî’‚ªêpI5à¼Æ[>¬›[Ù¥7OèÁ…}ØóÑ0‰#jÚ¦®Â°¦(hç^Ì¤{ToªzÒd8E IÑ²Ò:k Š€«×g*ƒ «—hÓ’°ü
|f0|Jôy
švHO˜ÿ†žÍÇé¥Ìm_Î3,RgÜ?cM†ñt6ÏÐ+d;ò®ý1þÕ ãÿ 2ØÓyn7¬gÍ±¾Èñ³<ÒÓ£ÎÉíëOÔ™@CWæ…rß:ŒgÂ(–‹ÛhX	â·mÁ ö¢‡1£ù +Ù=‰©bwzµuì~ŸÉ¼Î]¯ì² ß”•íV¯3ž;Ï¥jìåh
úyÎ¼Ë$HHL(Œ	H6L"¸
†hªÎ‰7«ú¼
¶öë…ò8LeòöQ~Œâi?°’&º»ÇÆœ6Goü+8V&§yÓR¥>¤‹ÕAla: 0ûìtQ¿ä\ÒÚÔl=Î£¾
½F¾b\–Ù‡²¥f[©c´l›ùõpzñÁ‚ÔmáLÂH*'."Ýî6?¨`çÂt(Z–âÌæaê[œ	bŒÞZðZærî³¿V”f{ú`9ò<øúÈóÌŽ9¤"iŽÓËòÊ+Ðá¢óeŽ]zx	Ö)sŸ>k’K¶$S–¤ŠñÙúîÒé	ðšó™
W…ãMþêêuR3y¨4¤˜Šl»jø³žxÁêÃÏÔ£9õÇs@Sæ©FO×îß_¼~¥Ã”+ý4þ¡÷§À¬Á2¥pf1ˆÂà<õÆ_ÑÂÜò ÊÒÉ–žË=Uõ±s"¶b	ëy2%Íˆy1µ´C¡¹Ê§ò§™e-dÀ¸ÍÏ¦N‹aH/¥qâÂ‹Ùå®‚P“²ÌýìÔîŸÓ8R¿R¸gñÔS‡íXÚ*»”VLŠZËévQ±ø^aâ‡'F b$ä”ºÖtˆ9è¨Œ#šG4Å¦1"R·f¦Â+`I%ŽÏ¼!&£ÀÐCøci™Ì8ÏPú G?P¾^v>= éáåY'™Œ™¶4—úí>úìð%f |š·k5).jXyp.&ŽhŒøÒâë:çâ¼µœñl¢¨êxTŠtQäx¨Ä¸Lükÿ;e^ve§"q|.-·ÑúSâ›ç¥r_‘Y‚Ÿ.‹àð‰7Žîí_cVEà†.iS —m,ÞàL'1ª‡ËÄÂ”ö@Žm%ÇB— ™úø– žœñW¨ÔÎê<Ái8V«÷âQ”q!sŒ–Õ¨œDäªäª­s¹à9Þ¡“‹˜f<Ðä`9ß§3?#Oh ê×Ãñ¹C5«««aÏØ^9ØYÑC%èYñ†.ø¹|æ¤ËW³¥¯ø¾ŒÖ?,GD«¡Î—©i¦ñôíFú›ÖÀçòÅ×6ïUÕ›w*Í½ ´:î\5ÃJzõã2ˆ­çµ¾› —Å‹·E 7…n£èrõÂÈüŽza´ßj6,&$«³?¸”E‘BºÕ+j®]kâ´x‡Mù-v5äºØ¥á½DŠÞJäˆ™^‰ób—Z´°pa– íõð`|RÌ¶†ã/3ž»³cü¦ŒNœòd'²c²LÏÊÃ<9TJN“ÏÿŽÉÚR'¦íOðc´ÊŽ=£9Ï¦d¡ØwÊ»ô˜êa¼?|iÀ„8ˆ7îhxsÎL2ÞqäaÜ5¾‘ù)wÑ†ÊÊ²Ý z§äAx¢pÃ´ÌK¢çëØeäîØåžW[Çòü:©Î6‹Wua3kçûÃÌ ²K?ä¡ˆ9NiÁmÊ/­gÙ¥
þ¨Ýá$5š‡a=Ët{ÓÈu0Âx˜3òÉ4!añ‚LUñ½HöÑ‹.½hC]rÀBÇ±"M‹5 ½†Ä§C“Òí ß»’dkh<ÿ®§þÈ¿Á¸ÉF=)E7,æ16‘X½=@øò¯¼ôxÁÿ ¯n;c{€å6{Z;Bg]©úÏ€¢xµŒ4[FP—ã…øKý^aæ}Ø‘™1›|9<0{Kg¤Ï/¾Àd¡Ð¸á5ÖÒó8™šßsL)¯9Ýãˆå¿o8U–ø~}“]u¢yf™àxÁ²Î²Ñ±{æïÎ“ N¨Pþ¶¸¯ÿU¥Ï®s5âó ôÙõw¥Õ¥yýŠŸÚ/rèÓò_Ú÷%Ò/¤ßæ>pµŠ>ð—nö?ø žceLšøòœf%×î©¿æïaZFù;þÛÖãùÓçEW”|œì<}®›ZFù¥›æ¯y´ñ%¿¡þ*ŒÇñé$-Š¿µ !½^ú©QAÇOñ­ãEñ·éÝÜ¹Ù“¿*ßÕý)ÿ*+Ôo>OâéÖI‡;×pŒ¨ÚSê)õÇK‡ï$žúš£×Aˆ¢Ç/6¡æ{ôVO++-!^©^”âtèœd`úWyðé^•/ä_ÚO”ºŽiÍK/ò?uïÖsŸêäa¦<A]ƒ*=_ù®®ôø8¦~ê74Î&×„ã ¢?±IW`1E˜L–7Æñf®„Þ9ŸÃñ»1v7ÀmèwáËùŸê7mHcZvÌœ/XÃê-3KWúvT½e%õ¥ÏÅ]íQö´âÊíWå}®{ë¢~WïY¿­pªû÷|pyÀºŒ¼¶4kà·»¡OèøílÞyk\ã©7Lâ$“ËØKF+,v¥¥ßîª¿Ä‰<óøÊ‰yè¥Ô™ðb>zÉ-0ñåê¯ÊÞ‘ÈŽÈ¿uD…²Ü{qŸÆ8ºÈòV° ”´Ú;ÞÖG4ŸÛ¹G«š.ö¥PÝˆ^y×ÁEê˜Ó× ÀßÆµ²ét™}§S$äÕ»Ë…»‹äî¼Ö6Òå/œ´ÍqûÜ¦âÜ¡÷ú¦ÀÍy·¦êŒf¤êA0«	dE­ðz™p€->’a`™;[#]oË-ÑÚiùe²Èm¡Ûm²xk<?FA:ìxiê§)ÆÉ¬p~TZúížOa"'hÏ÷†žA½ì||¸pkå²yS_Ïª*Zûí®ìK:Ô´þö–zAµå*ºhâoÆtÁ'¤_MS»¢§Òû¥ŸšOVUzæ}T›Ð<Ð³ƒhÿ£Ku¢>ãK·ÖÁf¥¥˜^a‡(Gúj-]Õ„FÁ	ØBÍmº¡?µ†ÖqŠŠtõ‹—›:/Ò¸_I‰tY¬R–\9¼œ&ÖeÑ·ø&Ðî¶\H¹ž[õ2ðz¿Ó¸Œà±œˆã.,¬‘Èæ¾Hd^ŸnôV´ö7Co_°	½ô"ØìÈ
ØÑe5àÂhZ+¯-l§µÎíPi\ÑŒAJW´’ìU´”?snM”fU4&é¬n4ÃâŸ¬š^ÜpÿdÌðv¼¨à¬_ªú«Õ‹U'ÄÆpØ¾4röµž˜qây!mÿ«7)6hUK./¡«jZ¿îˆÂ·V²ÍŠ¤©a9T‘jóŠ¢Ó˜¦Ñ1Ðb5œµ²6Ò[¦l§ä×vø ÅÆ¤1_³:/mþ¢qêÇ±P¯ /•”ì”
àÖz\2˜u@êÒ8´U.ðCt‘0£x#O×‘$KÚ–××fhYÌ"³v(¬:­Juàœ“Óþ"¿(ì¹¦““°¡çûßØùÓB¡g¿ªÎéÚ¥TÔvùÅ›ùíê.LII:=¤'ñ'üÛ-wÛ.2×ä²<é›Ë™´‰²ðDfòŽ¶'}m›†&­GÜù×ÇòY‚}$ò’Šê#yRE­;¶Á»–zûç_ù¨Õ.Ãyâ¨œL“¼Š¯=y5PÕû\ë2(Õñyœ#‚ÿ >§™Ž‹Ÿ}ÌXü0$Öà¯RÊèN½™±Ó%c¥‹f`Ý`TŠIç®íŠÿµè÷qÏg¦ÐŠjgÏð¶1ÅDŠ–¸ÚWÔá³µYŠn)²8TR5ÉèJ¡ÐQÕ;Ÿ¨ò—Z*bÔ¦R-ËdŒ’áß˜3P§i\FDÉI5owiÍ{2pwËyÝ¤*)-‰(ù¼Zº4ÖéŒgç	:c©z¨ùZ’ƒ®7`å"KŸ4áªK:’½âÒÀâ:B×ê·°‚î[ñÉ{"½·*nðˆ¸_1ˆõ&š¡Æý¬ðÛÄK'ýõ¯¯•®Ø_pË'„4ÞFˆŽ&ƒzÞê<v/·Ñƒ.ñ~ñâ‹Ö1¼Fm`i´’—zñÖI™Ýb¢érò)–ñê œðj>›ùÉu;ßÛP'×fªŒå3Uü{°³CcüÔQšøÔØíÝ+þ,¤ýó?‹qéÃç(Ø9ÒÄ6ÌÐGH„Âáèç‘Ø¨OÃPJtßñ²<©­:hÛ49sX§1²T¨¡|mjúT³wR°§2­LÞcew‹ÑÁ…É«WôwJÕ)9¿—WáDÿ8_‘«ü5ñ}UFe÷ ÄbŠ”àÒ’öÅ’š”V‚ÂL©˜–ÓÒ“,<‰züÓ¢‘&/‰X,íJ*—J[ÆiIm@<…u
ü•T¢oz€/¦XƒèOm¨ çø˜×gäçÅH¿2I¿&»›õÔ¼Ï¡Ä¾2¾Ïß”Tüüu?àÉ·aVâß›‚€ã7¦è¦¨¤m·¹íÜû¤Cz›ä÷ŒÙ:÷¼·EäÊÏ6”
YÒÒoYå }éÓ„’? €©ÆîœC74»Å!ií^vCú»ðOgþIÆ—^(ÿ_wç`þCðù¾³«zgw³¥ÏÑA$éGÉñ çK<Œ[hd
®®"?M	=[däMgðäô÷LÊÕû¤|)…i^ŸEÓˆáp¬¤n È©äH¥zå°ªÔµU£/Z¿Q‹¼ÒTWt?Ú¢¿)}Ñr#'ÑJZ£¯Do´¤‚@¥;rÑ
,¥PœˆÝ.úqoIšeË¤(­á­(že/¡"·¿Aå:v¬Âà]½šdcöZþ“¾d:våKNâ&:Fö»·gç¯/~zöòüÍ³‹“nÃùÈO«¯nš¼0*p©sÕÆ\Æku¢°öjàgJ¹­vc=›&hImŸ»¾ïkÐø­iK3­ßýnêû[ææ*¿J]s‚Ã+}×ù-	sÑÖ¹kw‹º[vnÀAåpÞ­[E‡fÒFñWª¥õZâ©£ÉrMd§&ƒö^Áìu¡¾BïYªÐ—(3v÷ö\À·ÐZv¢jª{P¾$¢ãLo–¦8KRUØvƒY¸Xêy¾K5PäÂ£!p(èÅÖ†öÝ'ÃÎÓJ:n[&é‰ã¹Ëx(h•÷ñÖKÆ~ÖÎóä¹7r–JÍ4&ßZ7±òU“àj’@%Å}og{€5þ$»„¦MUlPÊåÄô™öŠ‹2@èx	yH…ê«Wþøó_‡AloÁé\ÃÖûRq§cMWô¢úžÍhaRãHïTN!huèw°nœ;ªÒå‹Ò¹a4û³I³/õ¦<œ(%ž‚T2ŸV	·†î;P`ÅYR9µ†N: èñÎ‡Ö£ÓWçÈá\LÔ|Û9ëQ¹(K3ƒÀ=y^­˜ì— °Ç·"ïV«úœÅX³2GáÀPÔä”Ô‡sqä%·-)év:XîL/Ë-,“ÚIâ{Ú yÕÇÍüò¼§ÔÈt«¼%^¢a£<7"Bûý‘Å¤C¾Ïe¦‹¸ûÐª&À\äý´ß½ßü¯d{›¼ñá5ž”.p`ïËSãU|:¼Ï¥‘ÿÖ#ËQb;ø×I½¡»«†Ú…5¶´ZÈ=¥re4Ë&Û¹=×­ƒ}ý›M½MýËÛÑ•‡Gmñ•,VúUt?kÊjƒLk>&öüÖµäÕ+¥Â©wú#Œ¿.Lv4ý·þœ®gÅ>->n,¯ñ6­E0õnÚ½-Bï¡¥P{úë‘êAÄ‰3Ì´«Æ•²²«Ù/À“qçüó_ÆAäfw(¤,ŠâUœÚ.[Au`0_r¬<ÏfHDmç9²ýÝrÈvÎ3Õ/…lZ¢SÍiÎ®j²ZwB¤§— õî.¨¦43Ã¹ƒË¾Åßx—U²LÞ	«eØv;À“áAš}þkŒÊa*nJ[«n•nëÕ–~QIÏD3ŸóMyÎ£Rzs1Ð`´™WUIÌW;ÅJCã€žuøËÈ¿,ëè¬¶6Î³‡ó³g—‹­´Œ](»è»ÔRË}•ìl*)G¯©âB«'‘WA2mx
2ÅŸ=@ )Ž#ºæg+ÖM|2õèÚòjkq	õ:™ƒ™Âk~J1ÏªïÕíñG«é0ô½èé¼ž|E#!x*ÄÁ4e•€,/pâ
²<R2b#Æü!#ZÅ‹ûÅ10F£˜H@3¯¬UZªUqÞ-sÙ	E†š¼@üMJ"[
Ò$®3ŒÔ"Ùe7³'rUƒ‘ï´ú]U,`EÍ»GÏýáÄKì#±ø57¨Ä#”¡5[–¡›M˜·ÅòÝlš\Öè ÜÖð¢®k4‰^0BÄü»óý‚÷s§¬V^\u§¶AÕ©]ÕzûZµ³¿^÷â³†—ÂoxœÍUêè«¶ŠË¢!¨z®9( ã©Wúµ–úÖ¹Ä™ÝØêé‹¬†#¡PÚ°z¸Ñ©:÷v/5‹ŸšäSV³YÊ¦WÄ_‡º]Ò˜)Œ“Û’ó¬ƒÚÁÝÌÎÁæïÍjŠôºÄx UØÛ­4¼*ô1’€§°ÐUmÈïQ]¥|@ÀùÀ2¼Äw.6'q	2Sè¨¾ÏW™qãöÿÂ.Qˆ³Z9»1EL!M;Õxò&}t.ù&°³Z‡èÜìFúš‹È¢¯!î$Ìä}1SùŠPCg‰2m‘ÖhD6Fþy	ûë. ýVFŽƒYöä¹Ûl„|îx±:âìÓ1^Ë9^àµ´ó^Üîð6FËrØf0—€§"ßQ3 ¼œ|ðÒjiU~±HèËõÃPçäºjœe?AÚŒ0â#çÏeöÝ­57L¤¬Bˆ¹îqÄôŽú©£/ »a÷ÐlT©š¦| U(ÅwØÞýÿãÿhÚÉuŸÿOÚé&ßk#Un.ÜKûy+NMµ½T±?]¢–Ýôx5Ö5àå¬oÀ«ºïä-V9¤z²sNÁƒ¹l¥^Ws@["ãx¥Îç…ÖååætyáàQK¯ää2	tæ©•}›WöXÅey¥Õ[ní}ŸxÉéÄK²AyUòjk\™E[¤1nû;JnãO‘Ÿœr^\üF~˜%Yê§”QÚØtå‹P*RÃ·-Ö„Vi*Øî)ÜìKŽ„ü-·³zRµÃ¦òÑ—0ßÕ³â÷Òˆå[Îo/G6O8RÞÃy7E@Êšz¼èƒj(ƒ\\ŽìÑ2aMö©ƒÖØWÅê|â±`ò0aoØ¼LðÒ/×d¤±z¨Ð±ÖeTÎ¨ ~Põ'26³pô(ãšaBo4€±ÜsXŒXx‚˜?km
X7Xbëïs+Tî”Ê€u‚QS¹]C¾^„7YÝ‹š<7‡¤¯Hƒ É±p£Ì± Ê¯ ovêW¡ÓÜ×R“9aM¦Ž™Ò•n/Rbè¤‡£“$‰?½ð¯²ràþq˜y*>Ð°¯ÇQ{:®áC«”¬ùÓg‡hbúgÎHu¹ÞËØìJ“úÓ€»jY|,løFßuã¹'!WµPnØ’9R©ŸT± ¦ÊÕŽ•1Hýb¤e¨È€~Ð0»¦¨Ô1I’4»èé5ªríÅ ]ÙgWð¹*õ™Íëê7Ål,ìÕ²8#;ÖXÓ6)ž•nÕ(%™`^V¯=¥êµ2QÑndfJ4=#£gSš}¥ô¡œ«Á¸é‹
’Ñ0ÄÑ@xÐ˜ì,%ïÙ¦Áö-Ûh+ÊaîÃ³nÍ¥Ãï7,P8r%Wt%<®I@*DG&7Ûu‹¸ÕÜ+Å	:Ð&M¼_.-øk§Væ°¿U¨•A2QîX•iÀå½Š¥À5Ö Â2*gäüÇMfð¦ÕS†’`êƒ¯—Ø™Jà¸ƒûÁjÁã’ÂÝæJÇjíÍÌ‘^¨_[ž@u\ébû÷„UüÙpDNhxaRKJ~¿­ŠJjG,Ú¯wí‹ÎŸûþèÖBÙe­ÓÅ0R#‘6#àÎ q§§äÔË±Èj>½µ¥ì1…I\HYUñ÷hÏ¯©Í½'y1b‘Í
?&-6Ô@ÏÒCÚVî£*¹¯(üS…: 3›‡©’ë$á…bÚ´ÞÇc=­1PT‹?¤f:"yNn"Í…^SIÙ£L.sív»^PÓî&Àx™)žNüáÇÓ †~_=Ë%¤5:5Êñ‘ÓÙ”Ïÿód%áßS–äCûûò½nÓ>þ¸ÈÐw¯ýî=ºpLâyÌM¿3
ÆA¶±E¦A4ÇªFÒ-È*ÝB_Í”)²Éò:¨H™†z¼¤Û‹®éÑÐŽ'z.¶fØUïqaQ>•–P•¯’LÙÝt[ÙéKAK`¥tU€8'ÎjÿäÈÆõ—œÍî÷óÿEMBSÀme1¡¾åéZQ}Ñ‰ò¸M´¡F öô ðŽn¾A)¶(»`­Ö8	}7šxÏÚœ[‚þ Éi!¦¿r2Ž¯¥Æ{¡5Ô þŸ;×mˆD@âïùïÚó{:¬wÏ$›CáhÖâùÙH©«2P!+oi•2\ylmKiU£¥ux¶]Ô²Lýt£<É¶[<‡ÿeIœæ·l¶£ÓxvkGm}ŽÙ$­%&ÜŠ’¬ªšK6tÕ:)-è+¹Å.îìµcFs¹[òÌ¡¿ûLÝÎj“¬Ñ%ûs3g´VÕ`52K‡K:·0zñ½bRV¡>Eu1kö“ÑFi%T\jÖE›¸Gå¾hë‘c{›ôº°çF"ÊŠËã’ XxK~žÌQšó@Üõ’°.¼…†>FÂqÏ/xFìóÉO@ZŽ#_%¤0½Ý<
 Å7b‘È1°yÿùív¾Š€1ê‘÷»äšúWú~Ä‘’Ë[‚i4´Ã y#üÑ“[|×òˆØöáQJùØ-"Æòîý#äâN9ŽÊlºWqòÌNèðõ[’!ñ>PU…:ÕuXõ•ìË(N¦üSü“ÕfŒFÏ©Í´yÛ›Ø ÿ[9~B€€ˆ™'ôo¡½×—˜Ñ¤ûÑ¿MÛ5mv¯€ðë¸¢8aÍx>nRY–æÔK}ÄÈãã|åGj6I3zZ\_W(MPtª#DµI¿¿‡¯ß½×æÊ»†¶º³yÊÐDIï´h?è’?úþŒ¼~õâ¿ª#ÙÄËˆ7Ì€]„9AÞTlÚ”Ð>¤Ì;* s•zœðì3tc¥: ‘¾
#”ó.#íw?mQäNßSä2øÇÑn{[ä’~å½Ûyße14‹Oâ·/áÖÙ˜e'o6¶h U¼\Sÿ‚,RŒÒä&
ïêú=.£¢oÄ¦*Jî,:càõ£¨@ê—²¯T£2¨Æh™Aa(²Ìüd<ÇRö¹Œ'é´å£²»„nP‰œ*Ø1ÊâNûk¤`! O•<2XªÚdŸ‹#DßÇ¬„¢SWœLytL©Æ¬zÿÜK0Ï´òÙ«Ïÿìu`Àq†áyâOƒ4õÒqÒ—'ÌÕäµY—ßÓw3fÅÔkMƒà«Çc‚¦ÃÃ"›ÅÃ‡[ìÇUÇ°«CÚ.µŽõ ÍMzw”[ƒŸš>%ççI<0"`äåPÚ\VTÂ~ê™P¢Šù'_
,@D%¾<Õ%/ÖÚÔ,s©º°yÓÎ‹)V±º|ªŽ9…Zä#Ýbßná`¶ä™Ý)·¼á8b†S<€8y<×å]{AˆÚ(šâí·Uj‹VÑ—þ¤^I…X‹<qŒœ:š-÷Ã¸ÏÀžx:öDõ.'~”Ã*¨>Ñ çÊ#‚9©…H™ÿS|Ø^`Ù
FXN7ØÓwË¬ <…ÑØ £1.„;˜S…•á c•…yôå_ïzïë5:ÀkOˆ	Œ$èœ0KbX8IÑr…p†ÊŸâLš0Ìì2Ä9k!fq\ÁF›cAo84	ŒK²RN)rBdÈGn(›cx'Z *]úi`mKïYØÖœÎÀGJ2¼Ú2ezP¨_™AåÐ6±§ü•Í*+iî	´)€Ê·%D³îÈO‡	È˜ñOùsŠýÈœ"ø±L¿Ñ{k Æ©+ìPôú¤q¯—Õ^/›õ
ˆS={y~rúöõO¯ß<}ö†.òÉã.Cú˜êSèòÐž4ÇJÞÜesO6Ç±‡¶×¡£l ™pÎDŸ†
Õ‘‚]Êž'^»®VLªú¹w>M8VgÂbŒ^¸?¯íƒ’Ò<Ôr\*g4…äào{ž­Õ®iÃæü¿ëHÿ¤s³VòË÷ZÇŠj1:±.1¡Ô¯êrè8¾Ð²Ô¶ªÚLlÉj^$ð;ñDöuµÖå5+MÅ4™ÅßNž'ÁÈž¼ÕE“µÌ/•‹Á¨ÌP*o¦"q–JÏ½0ó¦dQ¶:UŽ%_”ß¹Ã,€eÖ>³ÓOÃâÄŸ8Pd·Ò¸&>KðÐ¯æÅÑ’%k¦ê#M‰Où*{Þ´xÍ¶Š×ÌhžÐX;¹z}±µ¥ÃË®/Ô(uXñ†ÙiYCºÒì6¤!-@k²É!ùðý¢Çï~÷ÁgŒî²æ^Cæ½]^Q™¤Õ"cF‡â–¬„~"^—P*^Í`C“ÛÄQÜzT‚«S9Ý¡Qqy¿cè¸-bYö…Ã^gòÂ(YvûÊõ¤]˜9rµúœ¥‚œ»côi‚©cœ¹—è2ë]ª}øÔC?y²éqI¤V/§¸[
™QãqÍEP»†C_·@æ?€Ó^âD
¶cÂ÷çÿ	"åÑöd·	Îmvö|¡JIû
²[Tœ}ShØm‘Í¹ÐÎóm^;oË&œ+E™L}Ì¤jDb¿­1‹ò¼¸ÜsûŽ¿/×ùNJ*
§ÄLž¥ÕKÖ¬˜†Î:	±5®¥eÁ@„™1ÅYÙ^ãoZ½ÄZ:½\u6UdÙL“~¸H3‰¿¤÷¶{²ÁÃq ì|g×èB%®ž{¯”Z¯ˆ±i4
ŒpŠIoãÊ0¦Ib.S±–iAã2dàt¦§QÒ­J'ÎØø/súÊlˆ4Å .Â¯Ì®àìƒª„Ö+ØFM-97ÜtÉÄ®Òá®`?î¶ÙoÊy¸%`p_#Ç´…Ü%6@?É0 /)>LZRWi„ûT’^-C§#ê9ˆâj&~`Æ†­*öP‚³”t!.•”Ë,±ƒŒ!.Çtn¹*sY¸‚»U„Øø^ÄØÈÝNcÌ²œ-É®rvUÏ«ØUH4N¼:ev²¸“«$ž–¢1âNæ{a•UF²‚Ðz¬*Ã5M¢¿Pðv.ÅN3ç÷œ¼Øõ†Écj•…ÿË·í±,uÌ¢ˆ’ê“H„übá?«s‰àC7g¯³î/“ï]qUM^hØ¨{)€¯±!×?ÿ³Ý‡ÙÚï>2¢©ÿÖr–¿ÇÝw½÷öÁëLReÈé¤™ Êå(€T·}¡lqÑÆ:kLdº¤3ÉO	×c%­IäÖÃacŽ5Þl[=ËF.±Åµ([Ë©‹l€ú1pP®¨eÑ`pkLó—pcWÍì:ˆ†ó0÷d±¥$â-´,­]û_3Ó`›æµ~æFsMFa<¥©&šÚOeÃynC-ÂÉ˜Z°tçã*¬Ÿh â7±Ú¯£¸šÐ"ªåS¹S¡¿©ÈN©)e ®P-¸e__¶xUk“ÎŒªW%’ªñìŽÈ,¤ÔWk•€Œ3 ÛÙŸ{š³@§\EÙ8“~¶ÔÞ
²¾°ßõäØéF±õ©( ê.þ‹]‰ž~'aopQŽÇÖ}(Ó,ÓÏËÏÿ:
òžTFN©¯¼ð{g¦Houªßuë.ÎØtÉ‚OüWÒX8¾Ö˜
Ö¬'¸G\È)Së	=ëD8ˆÛNa%F(u§n^Á¨œ¿¢³¯®Òk˜öËß˜_’5€Œ¾#ÇÒÓy
‹CýZZôƒ§`’îlF›êÕd„ö”(®Å;ò&/¤ï-®Ü¼á¥·Ñ¸Ö«^Þ'/È€,¤0­6GÏ-æwÍ¼»ñeê'×p]üÆÌOm©ÑTuZ¬´HÝ·Ë·º ÙN¹Dº±áV[¬z¥~öƒù¨ù‹Æ,y‰˜ÛÇâF·KC°’W"É1‡û˜Î?9NÚ2ÈFo•ýB—;¼ÁrðÄË	aD7«a
»0wb­ay“˜Ú"žgNšœ†£5HÙ¹So€þ,X÷„*|²ôàï¶È ¤È%¿wÕg—¾!CŒo#XeyôC¿ûÉK¢vëY’Ä$ò ûi¾•‰	Vž~þjôÐM#%„‡­-â^Æ¤<fiàÜr#V‘Zµn©Á4þ¦éÎ\Äµ§_\ütÇü/OÈæ3Cß­däÎó'p iŸo{éë¥lßòE· "¤h‰Å¬%dª°k{Š´.Ž‰óË—«-¼H
êž	Jw9Û»ØÕŒÓÑ8¯²,É°ÍØ¦F¥tò¯¾ffK³}c¸ÄõE®
ØƒLW1ƒoŒ×7ÆK¾¾1^÷üÅ7ÆË|)¯<f{iæ+×¼–T­9ã•Ç›|Y¶Ë¹°{õºo¦ë$ô“Œ1]ež‹¯Ät¾‹&
øÆt‰ë‹0]2Ìƒþ7vë»%_ßØ­{þâ»e¾ì%UËóZÂž.[ÐsN‹þúâŒ–H¨ùµñYÿ¤â±ü÷Î`9»”^_Ú´;~êô¶]}KÔ8U}klL½K?Ô‡V ËQÔa²õè5ãÎX&x²M.•†GÛ´ù¥Xß†È´L	{q©ÁÇÒ.ÓM)éšxñõr«ËòªÜ9;àÆgî`å‹åß	–å<”ËÚ­SšÚñäüØðá¨	H„Üò*Ï1õ_—œÇ	¹ò®ãd‹x£``CþÁ#Å×‘Gã$Ø'{©—è<E?yL¡KÏô¨'g°ãÂ`üùß" ÉI/@Y°ÛZ’ábŽ_†kh"i”ODtØ¼/Ñc‚‡YìÐÎe}rÈÚ¥µÉ _ZYEþ¸#~ÊJdá¹N<âÅT|H$Y!O«E®½_Ìz6ô/ƒ‘W%7ñì{.1%	¯Ã½Ìœm3þÿ¢ù;ÚIwê§©7öùá§:ÙœmÌ«HLk”—~²ºð©>‹óíÉQ¾Q¨RŽ‡PðøñþMº„Ô3,Ó™¬ÃÀ_+Ò²º¯CbbÏ–x^A£ZÒA»(Ã±Ü¨ÍT+yôj;»-WCÉ-//ðƒLHL!g'Íºo*¹n‡dõËP9ÊðÑ‹Ø¨ËeE³Êµ?–éú€#ÈY9›&©Jí4î•4àÂ\åßÒ[¾Ž—¡'VáK‘4QÞÈûîËÂüKAõÄ˜!V¯p©lhˆ±®še€n¾ñÏÁéR¹äòf”ºlØf´î¨‹àÄZš¨ÑØêé"&þ.M…ãEY‹&¢q]ÃyzˆùF©¶5ÿ!;ß±Û¼5Ê|`q(#¼Ð­,[SÐ6ÍTÃ®Yèý	0?~rÜzŠåî|âý<È,ÆM(iOýp'À§Çs,[Db˜(Ý²2†mo«€5ÀµÎýãJ£lh(Ý¬ºâètâEchí U9I¡[Äïf´zq—E–@›ôßh›Ôô‘˜òpµÄ
öN–Ë‹ªh·¦¢¬pÄŒÅ®d7ê©öƒ­&¼ÕÁ¬4{@ãyÀLÎ=T†ÁLæý¥³: Ý½./ÏqI§ÑápcªÜ“È¿ñ’¥iwÍæÙ2¦¦½½
B}U1Ó%Ç“²¬9Ë´RÛÐ+õa.ä¸Øìø;}Ü}·cÉå »¨²ÚØ$¬"åsøûO³XVz{‹l
#TuKj KÍi^KÙ@P óò¯‰Qû§$4Uã6tïD©zµºØª¼» ¯/[æZ3¸³©7öÏ ‹¤dÅ÷Ä×iÈÃ2	©IòY2h…´©wƒÞ¤9" 0_™Wñ”jœ6Î¦c[[ã —ó‚bW%Ÿá'×âOÝKðÕð¶ÏO—¡}Üh,9³K·@{êÃA©Ôë0^ÐNâÃ–Ñ:,m€÷³Jä…/‚è£$/‡aK9¾ño×¸ðŒ^¾ñ§°ÏØšûmµ\ÏÊçiÙ…ÙŸLÃNï«\øúõV»±p…Wc¡¯© ·”ƒÀÚ_>ÚÆx—w7 âÔñê	Œ´í‰?Ô>Ò>P^•sÉ­ä(ækëW²ÐZ½–l(½+§7¸Á¦ÛàéÌE\?=—¹öîUòÈáÙîWÏ––"·ïÿ©½1òÒÉe(¶¡Nyñ-?Á/X™üã*Þ¯Ï@[æÕšøn7ÏðES´ý?ÿ™Å#Ÿx!= ÿ(KPð´\ó6.†D<òóÜžÃ8§1-øÜUd®S ÆÑö4F¡¯[y²Yªð·¸1±lÈ€*ª×ùÒ£¢‹ê©†.Zy“U¢D z!¦ÊŒgÞVvÜÙ")Ö„¿º÷ê,;WV—¾éåßôêø7AÖ ‡L93Òi‘3þ¶×kAG&UâÉUR+‰–W±O`©;ƒ-fjv†	$0£˜«ºäÀöc‚Ì§$IúJÈú<•G“¾!_ÍQWlØ1oütÅéÑö¤¯iÓœ•C¡ò@B§gÔ²„È*%réUÒQƒÐ 6Î³éx0ïq7‚mk4}ÑVÂò&m=úÿãÿ4l)ˆ¨¶S#–ëˆèT°ÒhP¡¤›%zJ2ä/kHªòïY0‹9d¥Âh†#Øêwç\g/^ªÌÇœ”îÃ3k,`wŠ„}|òÏãäÙ(ÈÚÐEÛî0@lâiåËïÞž¿¾øéÙËó7Ï.Nº¦{ù)m‘<¦8ÄPkëØoû-ÕùäDÄV[ž7p–JM8Ö‹ÇË¨±‘UcV¥'K•Df%|­f¬®ˆ/VOmI¾1	4¢À=€ km v 5{%rÕ¥|e@ß 1a-vÒw“xlr‰ÆøVåŽ¥e¬+\„Ë{R)i“BÌ!Szc-§‘Ól_¬†9Û#ÂEŠ
œOK¢·åhëÒš;kM›ÊK/­™ŠP«¾Ñ¼¯»½‚fU—GÓ Br­B‚øSä'g,‰è<õ“ÇÝy0ÚÔc¤`É4È95n†Ôêô•EÃÂDÄu4:êá|ÎÑ¦túííÌjï£?ÅS7š/>Åü¨m4ÍŸž¥ŽçƒælÐR³·	ð’•L#ÏnàHT­™§Uãk•ÏWWÇ‰ÏB/Ò3áì¨­„éÂôô8¡QdMYñÆ#&qoêéÇy
ÂT4Bç*y¤o|8Y?ÿ5	èpOeÞß«Žy–€˜ÜJc®ÞçOŸëÆøv0Z”ªË?‹<“áËFÃRÒ µ^…–ÙÊ¼,%?`ÂÉzÕ}
yÁgúÑ? §ˆ~ág¼¢î#U˜û`?ÏYZtŸäÊíš"ÀBcœv¯€uô†“<¥Š@°^05ü?ˆòð¹ó<ó6Øä±ÌåœðªŒ¶?Öe`¦Ý‰—²[››†H‚üÃMþ™7ñª×+019VŸ-ªÇ6`îÕVï§l½ni™‹aœø|8òèäºÔ[$¡¢ïtÒMñ#Z¤ws‹¨ÜÚYSïf‰æûä÷Øº}`½Mct(T@:fÕN$æï1yéeÐ-rCíveêÛ¤]+ëÆG#:§íhtíe°4K,æ&žnf´i¾º›8áÔÅ+/nØ<C+§áaGæIÑ_›ŽjÅà6d=×8ZóõNGµj‡û:ýo’¦Ý­$Ú­T9+Jš™êÃÔ¿µºÀBIÄ³ÕÂÀzUE^^ã¸ý£ÆÃ`äüMc’Æq”ËV4X¡ôónY§·(íùUÊƒJ^÷9ÌÎ­GÀr`ÈœŒ+Wí¬CçS§·U­öUuŠ±ºeYòL{XòD´*d÷ìNÆŸ\Êô–ÚÇöÝº+I%‹í#;ˆVª ÅÌvn³ýYd–Øoñ¼A>ö\”"Zv_^BóÈr;l'Å$ª;IM”ïlUªÆìU-Wí…‚Ð—êænÚvÙoÎ™ïiòùßQ„ý–Ÿ”|2°xçÖ±ü5Ãå<áÂé—^emmà²)®jnËóIp32¬·T‡ÇRû­ñ1Yí]ÎžÏN$3Í¶¢ë»^ÏQó§1òX=öj¡ç<¤â\šÍÃÔGW=äÐç	=oIûìDï,A{[µðÚÐ±ü†V±~³ÛÆõœt^Ð³[gî<¨ÊJ%Wlâ­s„¦XPFômÆZžkUˆò<˜À1øÊ© ”B—÷¸×~X¬üÉYãU0`óê;ä]o/ßð;ñÔ~¡ï¤Ú·–~4V€Ú,m.IüÖNi§²¾yéD14š	µæÂÇS<D¸}I9nF—‚!ì“k¯4þ%¡opU€¬¡™¢ýITæú]_ûvW±?y¹ÖÉx'Ý³€á$ l>öŠfð`ô
^ð‚häµÌð²@kYw3åæÍˆ(©òA»·c¦q8`T¸¤¢ÿå¿ïT”ÝêÏè¥e>áR„1#ëPÖÍj%Fpìt÷l>®|ð*ÛdÐgKÀç©ŸyA˜Ú,ñr#Ô:ãBIIüëëç¯_¢ êvµDÆÚüœLy»æíj#è`ø÷ò¤,›ßv`›1;d™*±Ø=æÌSóŠ0Äc*C*tîE¬Në!õºë<ÜCWqÌ~Xõ(Ò8qä¦Í‚•B_åÊù°TÆ¼Æ•í*¾ÇÒ¸Þµ/QY<¿ÛªªOh9'ßøWÀåMN?q~ò@æÓ5|q‰3¥‡RËÊ„ºÒl3Í»MhÛ§üG2mäômÞõN¥áÃå´ËáXÒ.÷-%ÐtŠe9£a]€m,™–k¢éêx£À¹¯‹ð4:À¨”‡dÿïæ¤B=¹-úL›$üÄKN' 6\¥%»Z’RÓB§Žz¼g7þpÎ±¶™ê}'äóÿÄJÖv‡©¦œ±úè'\ü8”L]ð
ß4¨:mWu~’%Žæ]æ8óÙúa^L$¡>Næ¥·CèXù”¢p¯¶h//Ç›HõwY^]5ßÄ=Ÿ¾ÿ˜NÌÙÓLþ«÷	ëAS	útêpøN´ú~“”~B/æo1‹)uc§VÐs?á¹³÷µ>“´	JÓàöæ¿;®Øž­KYj$Ÿ-m¨ne^eq¹u:4z£n‘ÅÝæ&«m$Žç9st¦û–¿”fS·nÓœ¶+o—Ú¦>­»?Í’ø£Š~±0¨Öîí\><èµÌ­â‡—Þhì?Açg·äX
'D`à:
}`5ÞB£ØÕ³›!°#p‚´,õµÓ¨ÈîŽU«p¸Ú\íù–á‘
ê¥†ëóÎzöÉ“òÔ…ÕÁòÏ4(ÏÿÁ2óß{èï\6œ¿ª ryäÆ xã)wiÛä.»Ü½ò:Ý{HïjZcsLœPŒÌ:æá+ªí]µL%çô¸	¹pSð‡%aG’ÂÊI=FsL5?NúUN%A=àî	OÐÃãà{}ª“ôÒ0!w{Ûƒ÷v£5æk¶°šù0Òë±2=ýNü*(\ip‰I4Pïñ$¾9ní áD¸í¹õƒêA ²®×£ãÖKÒëoí28Ø@w;[=rp ÷œ³p°ÝuÛªwµwõ°áwÿŒ²Éq«iw/‚²,)¢”óÇWA·h4ƒÛ'ŽáÔ¿"Üu»[
ö¿Nè‹þP
ð’Ä»=nõz»ƒ×àúâóøê*õ3L§†ß“Ž8A¶³<}à™žf·!Ñ,Ô!Ù`ÝuFy¤—ßƒC1žg®ÙYœpD¸ë±Ó‹è]¦q8§2{–ÅÓ?Ð8UÓS~W
¤6½«~ˆ	¬+Â¼°ûPÄx´ ëâæ3%@á’ïZ 
0<ÔNÅÉ·áäh¯‚RùAÍ}’iŒÖê˜tÞ˜ÈvhU,…Òñß Ü•cºÝzºÐªœ;vZ7”-Áué|ñ«×"g¢\6—ë,¿P²…¥S*¬Õ‘H§s3)ÙÌ*¶³è*–â–O€ô4ÀX€ËyÀY,BšYÿ×¤¼š 	2O†ö4þÎØ(ó`ó%ä;9áèø]‹L|ÜMü—E÷wøT÷gÅ-xÑŠ¥#/óŽïð^èXxÅzJbs×ð-BcGá1‘á¨r±så­%Yå–ù]©u&x5m§ý—›~…ëòfý«]¸›ív»í²ç7SO¼+fÃT§ÃoK?Üõ—Ðñ{rHÞ½7ûáõÞN¤†À÷ï"Ù_¼u|1ˆ eß áŸ§Ç‹ýû€Y)>8pø`æðX9‰ÆÈíÙ?@Ìý#fÎ  ¶MÁÁ‘Ãm,‹æË"ðÚ1tÌc:=ÌÄr»(1òo6mñ½â::õáÌ¦ú‚Cø³óý‚6p÷áŽñôÚlwÈdþÕžðèh( •J¾ã0fvib¤y!Øv¬2¦,Éi¾¶ô‚MÏ”ÿáÁÁì&¿	?{³t0"\àuâë¯¨ú¶Õñ¢Í‰
5d0’#t¬ï>|ÏÙßÂËí{ÉJ9,P§ð´=GÛŠÓnI»¥ÅT8Šö]ä°Í¶"÷ü%FZxùÖÍ2‘
Ü±ÈqÙ«qûÁªöBœ˜w›£)îiœy5à–TãVŒ”ÿüzF›ûr‹bqb¤èÔmå¢NínS_ff…¯*]>Ä4_	~Ãa’«J>Tel¯	-ªk‹	°æ1Z2©Ÿ•5‹…Â+OÅvÿŽæ„WÓQ‘ð
þvHxÅsYÕì²³GÅ(½[…™€Ý!Ê™±vT~ŒÂÚèBÞJIkû;F¿x‹Ìöì:¹ã6òX Ö“œç—•Ý¬1°{©Š¦Hpoé	–Ñð¢‰O<ÔCüÙ§ƒ¥ÅjåÐ:Z¡)Äø<å:ý`œ‰0ë½”Ò&¥†P8Ï‚kZƒŽ&àÅ&iUëkVXaÔýÛx7¦È’8L¡÷y[*œ/%ÅšIuŒnBtZ¼Ñ£Cú=ÚLz„¦)¨R?Ë“ŸQÎ„y}¢zËœÉ·A©'ÌC³K‹4I*(±qÊJ¶²¥INìêä‹ú>3w?ªM–9Ð	ózFŒÊÕÌÌy+»òR_´Š/½äd˜S“„e µÝçèEMè†yŸ²;YÝáÖ„I—^ò
‘p¾¿Q<ú!ùü—+Ìš¨ôM=Ax¶Ü^¶¼³NÌÙkr×â5‰Dú¥ŸÁ	&‰AÞÆJËfuLaßÙ“Ù›%Ã¬C¿sk³¬‚ö"UKO¹úpIíŸf¬1²J0ª—Ÿÿ!æÑ‚2çI<†37íªûâ\à@ßi˜"WfNhÐ-&ld`6\bªCoi6¸'#TC›ËÿYæÿßí¼ÌºÜ-w¢t¼Úªn±pÏÕ§•÷i·Ù¸Ø	ï/>ÿ¯0¦9¹Ò£ÌôWŽ²M§†jY±Cz¿Ò˜†s?È´ô6d¬.ù,+Ò¬Ô²½o±u-²41ê¥( öúÒ é3ƒÉs.ƒDžM~ä.¦í¦(ÆrÑñCG¹\\vÑõPÈ#âàèè•§@nKVcàI¬­¤±ÌÓóˆ’çó\¨ìÝ³˜,®¿[H3¿û[~.ë;˜Ô]¶”ÍÇ¹p •WáèëY…<_eá'».ø×ÁO3~yð.øÝJ”Ý:`}áO¿<tOU*X¹{{ƒ„rB£	•&	–Y5ÊÈ) ¤ÑAPõø“Êq›JÅb@?F6kæJQ>îÈÔKÆADÝüâÙ!éíl‘‡Æþý+øÎÕ-îKwH°Ø‚•k !0óÓÀ‹hºÌš‡ã€ZDïÇq¼¸ò`ËÝÕ½líù5þéä&HS;r6-âÁ=œvÞ0tó±v§Œ–]´‹îï>Ø=¸ÜØ¢[é"ø‹P`
‡üw:‚5v9À@‘Q<õpeÞðMÞ»ŒfycqÙ ÜKü©Ù"Œo.¨®ž£7¼‚¹<Hg ÿ$ãK¯½³Eÿ¯»³·¹‚¹b6þÝ‡-²!%©aAa^'´¤_>aû?úÓ7P?ä0B;üqÉíÀgÕ$§qgnŽÂ9‚S&ÎáƒˆÒ¬JŸä»‡¹~Á<¬ÇìQL‹¸ BímÉ_Ó88±1r¿‡YLúÓ¼Ùúfwöîï?tiÖêe“Ôóœˆ Ùo”ùe..¦@—tupå](è7²~/dvj×I,‹ýÈégÂ½ßíoøà
bÊ¥—\Pô®½’áÑ¶ K;S4ðuZ®ŠÚ!CùÁf-²ÏÑcÂä¾½ÛztžG™Ÿ¡
MÏ2ŠS•«±©ïÍ÷‘þfEÏÇú^D£þÉ1H·£ù,Ä\2yQö´;Åi$)žp;*>údcKmž\ }?mÝ”¼ñwG<ùü–”Rþ«2jTTÑàs¹T¼”EÛ!öºÈ(€€ÆŒ4å¼”9ÀßŸÚ‹6x }þS“Ÿ]¥7™ãÙe›#'€uïêð|ÿîýf	 $Õ×0'Ú\º”
¨ÚcIx6ï]ï=Ü“r2hgXôû¤q¿—Õ~/›öLg8Ý³—ç'§o_ÿôúÍÓgoè<îÂ3àìb„--íK*Ÿ7øDÙà“Æòµ§-vèH58 ª1 5 ïRæ›üN€q‹<ßöï19ÃÆ‹5+eçàVê¾éÊQ`ç#* Œ»ø‰ÜÄÚ]lUœ1ç8ÃyG´ÅœŒÜ|#ß‡Â¨ÂŸ7†R(÷´æ nÖ3‹GC6$š'a3?)« z$¤Ûƒ9toòåç¥Fs¿YÙ‰Ö¡ÙV%ÕoAÊÜ5XÖKBåÖÂjÔ|ô¢KZGg9½Y9xU=0´½i|mª&‘Šº°ýS„è¸´k=*ðyù<úu×„b<ª:F‡Ó¾CòÇ?[ÌÕ³öÛlvµ<çÆ-J´/O‚®¼M÷V6ÕF^'B·Oè•æKßÀâ‹âQZ~–ŠY½"éã:kÜ/•«¡Už¼Jng§„‡ân¯Œ‚üö’§1'b¥œ/<û…0×ªíV±©Ì¸æ¼Rg‘T[L¥ Ê·WbI¤3,
gm,Sz¹Ò†c	æ›CÒßq.¾|CóO¹^¾¡z§¥j.×‹!ë|ÇÖCVÕ¸†M+U>ÃÊœ¡ABÒÔAkX2ÙPZL'/»KFÙuäÉÕÖPâ£ãi¨«™¬õßÓW2ÖÔIdåÒ· q,ª®Iàºwà&›4/\º:§“2/v}W¦y\wýîîž/Ä_õw¤*¹ðžü«þn.¾/ò?UãÌëKC­Þ3~WŽê¾òû§¥ÉTn(¿ÀE¢oâõ7f¡w{1ý4½@ÎxóÊê7Û+ÒTyS¯«¼}ÑæßqU*Ôš›6è¯˜njKU®“d®LOs<[’öýÿ   ÿÿì}ÛnÜX¶ØûùŠíBÐ*ÍèbÉv­±åTëÒ£9¶¬‘Ôžts¨*ªÄ6‹¬&Y’Ü‚€¼ä-Hä%@Ð'ÁÐOA^Î«þd¾ Ÿµö…ÜÜ7’%–nV¡Ýª"77÷e­µ×}µHûjÒª8œôc‰ÿvèRsªCß²78þ.	iÿì«ýDØHü=sòé _5¾Aj€ùDé«”kÆçòœðô‰ü—±­”øœ¶–~ëíû´¨ü»xûQ|·µËñÆ“Ÿ(_ÕŸáêÆIôaôÅwãØßIMK?{,´X…ºÇ°æe³©fFÕÚÇÑ·¼Ò,úÅJ/EõÙâšéÉ}ŸëëåGýâv€¬P¯#cNr ‹þ±—RÑéÍ…üË¸j‡¥æÊ…æ§öuO`šd9nï Î»|pç/Àqek'Á``¬ÝøÀãí jWl¹ö¼/@íÁ5·z$µ	qº¿ÇKóƒã»1º±,Ü›/ýõI×ËNS«'K>;mÂÚ+…,š*|iiÉ m¤QÖ¾Ð–,ÜÌtu¥×Xf
ÓÕ˜=¥–%øKÍWü5$SÍ¸õ¶4¡52˜·YÉÐ ;IýÄî×n_:üxg^ –o3îwøïhÌÉ"úÜ‚e óN'ouµ­MíÅj.	pýÒí‡ñd°•8&Ê`!ý%ˆwÒílá6· J£óÐþÖ:$ïØ:³fÔ¤ƒõm£tðÓ8@º1ô³Gf.Æ£¿X;“¯Ëš€M¡cÅÇjr´|[zõ²6ÿÁŸY©Çƒ8Þ fßúpfÄ”8›3N¾Ët\TÏgaäß²%	v\8oP¨×Ñ«ÅÑÖ 'ßÍí,WµÊæ¯ÔÊò¡5ã{çØ=öE*ÑZ/	erÿŠ£M˜îRÝWhûÒËë¾Š™—¡3£V¬#vÍþÌ°(Ì·¬ü\qÝül{ƒ­Ó ENxòÄ‹¡¿„>»ã/rµøy³²ª¼EßmËq¼P­Ò3ü·ëM{›ÛÅ+("õú€|›Ûv^ð&¸ÈÙðG7ÁÙµ'R_ûÜHüQúé4g´æùñ¹‘äú¹¶äú¹žäºÇGKO‹ú«V4é±B?)RÑ®ƒuxŽ+:D‚«ðãè8HF;£qœdÔ*÷ï*«ÿšö!„IíPþÆé%ä0S…¯°W½Ä#¬_Ø&ï
ŠÔqê›‹5*4z‡+\JVhÛ†ìHU»Ñ$KíJ; ¯ds5¿yà±Ú¥Î1IÁëg/”–ëÐ¨K¸”%Á¨«±_ î¡j”½„–aÑß» ¿Aéç²ÙŠ•V‚ï9æ±~Óá;”j”÷¨(î÷Ä6·¢WHë&À0éFJ€¥¦¦ÙC°ú	Ö@œH”u÷¯þF€»ú…o- |5q–OrO‘·
Þ)“êå5õt/åN¡EÄ¬ÉÔYérQ^/Óî´—°ÃF'òZi‘,è™œRËViò8†ó`-žd4ö•&
b—ÐsqUþ!y-¯Š'ÍiÔ}I÷D£ ›ŸzAˆ tà±ŒoºDýp2ðÓ®Xó ¶jÑcàÃ$Kbƒ 50VC$%¾UòÏ+FžìRéˆ€3	G¼ä#ì½6QOm—¸ÚujÁ7×”uÆ:ßã›V°ÓUv†¢N²ˆaÖÄ½¯ã1žPÌ[^Šý¦?ÖÙŸ×Ë¬™þRSZ^Ñ'ë¨C×¾³NÿÐÊëæ¾^/3\Q¯_èà%ï)îG÷I3à„CQïó	ö97?oIFDã‰ŽÇ"øI„É‡äà¾DÄOÞt¶°ZPpŒ•
I¬>ÓÓ:• nk·M*ŠD)†1sè±`IUR÷óˆ5ÑŸYÍ£èb/v­]ƒöUsÌS0DS¢+^/S@\(8làQÀç^b«Y<|RbïvýáÕÿ–,ÞŸ‹±£Œ‰éÝœG±0hLÈæÏXY—/<…‘õP‚,æÿõmü‹Üx Oã˜n‰l³ŒÍQ±°)_X‚éÊ|èŸz}×ìŒÌÈ,‘4É1ˆ!lñPuèã$ø™±Ü„åÍZ$¹5(Eer?HãtéÖ (}ÑÌJ‘#2ß¹ÓlÑqÈœ7Ræ1=cdyË#w$=eåŽ,PÕ„E²¬äŒ ô~[Ì’¼>HÆ©¼Ü·Â5©(ïd™ltâŽðM.Ýƒp,§¬*+/%ë(3•öãðÄçµÑS¦$“c£KÙ>fêIžAí8×Z)Ã3ƒîÇÂWq+¹Eý5:þ·h>†®õêŸÙ~iC_¶rçA_…å—.ìqÞá‚pêRhìK?%.Zºº—Ä`ªG½¡¼.VØ…¾Îo±lp€ú$ ´4ñNôT›É›'÷X°ŸÖÉË¡MO‘Ö>)-FùŠ4sÎ¸â}m¨Í¸Y5à£mnoµA™€×ß$L£ä¸»Ô×ÍË¹íb8°Étµ25™š½¤§ŽÁbë8tîë©„ÂÊax6fÔ\5xÉÂ?¥ôd·|J8˜ò9}¾VÂX›xÔš.ˆ$‹¹AáŸôi$þO“ ñÊÍ˜¢'¥»U@ßð ñ@<¦‡s>(çÕhØiz ìH[B9>RÞwê¾™!š™V¯vF#_	‰»ß„1pÞÞ VëÄ9á±íGkÛp*¥À75§þ( ;Á“Dq†ÐŸùƒj¨S-,ÇJ9C®ƒhcrL…9Y2L…ß†¹”çË¯³:,«å
ö‚U®*7ÿïþ—_,QúbÉ:ë¯~ú“0&½	ÈW¿ ´¬ñ˜tò¡¤×‚—ˆá§ E$òÈQpqæ÷)ÉÓÔG‰?ÄÒ­Ø˜ÂcIiÖudÀi'H“:ë•ûÄG:¿¤Aµª0»¼Öè´9ñûŸ6‚¤ú×A}›ÄRËëd`Y¤´7ÚRxÿÕ/t(ÜüÔÿý0û!p áã¾—àˆ¶(¥T;‘Àƒ$q_wºÿ•Û_™ÈŠe®zæD¨Bå$d!&Sã;ªŸ<¸aU>1Ð2ŠÜT!…ƒ\fÉÙR+[‹%Òâ›h-¦L(ŸØå¼.Ö”-vÙ)×AÑÉàÙ“²:5ÿ@L…Z¿¦L¿¬}¯I5p´ê(J¶h–ÿAyùê‹ü†T‚Å–ÙÅ˜—Å’,ÎÂé±;`ûHŽâsÛüäêÕ,nÆÖ;rK»†³åå˜±Šœý¨¸hôJwòvI#¥õæx÷â¥éQ§ôB{&î9Îú^
Q0¨Vò1üàƒåŽ%[£%{ˆBô\„ãlñ9)ªŽ2¥l+g_2˜ «(üÀ9­LØé³ñPÜ—Qþ™Ó„s¤~š!ªçŒ'ûs13o+¯¤1‘‘àB8O™×QÊhF­JO>:Gî${`2x˜cÆL5HJµ~Eb)¢Ll«èe¢Ž6FÅœã‡‘î‚•4kÜ• B;J×8ù)‘á´EßiÓªËnšÀ×è14ögS¬“R¾e«ôý5gþSc%tÌÕl37æjÉŽÐUp©öœZÛ”;i17_Ï¸paé¨Ún|ê	™R-îÔ!.]:¬Óøæò~i@ÈÅe-Ÿ\ëàÙ7¡pX#¸$ü’ß¾¨ÅOJKaò"1ëñŸøÅ´–ïg/>ˆ1Ð[Í‘„0jÜy~îÒÇ»+ÿ­0™ákä–Ö«k¼ëm~ [únkóÃš”†—»€Ä„§Ò
n¸ÊlüÎQÇ¯2‚Ë/Ÿ:k*‰‡Ë…0‹‚+’FÄRþÏVsë'¯ÚË+›ËGkòd+#xŸ’z¦TCF‡”i®³¾b?Xµ­ÎÞ@ò0ÿ²~¹QÝB{U:§bãe™óÐ¾÷3Naæ[N‚!÷iš·2tX5¤!ªn´5±+W[V‚ç ¡ëÀ‹[Œf®€´-ætSCþÊ•á%¢‚‡Ôñ‘z^S¼.­RyÞÆÙÕ—4ì†*Ij:ª™?ÜÜ³oj¿Üö02!°…ä#qBt-˜¾Pm„kä…Å¢¹{
Ð–;¯Û69Ù!ƒ¹€®2Í‰!¯®+eèŒ ¸³¾±»÷Ç
¸VEÞ¼~4þ9\š+3YÚUïÏ—²À>7;Ÿ ýè…Á Ç¢wU]ª«:«o®2.É„S%n%äïÿã¿’W¿À€ÝÅßªøÖ¨uÓ`ny.ìé'ößÉNtÚÎÔ¬÷¬(è®T×ñ,ù3<}ºô”ý[†O]ÌGÞù;
Ð4)Pâ\‘7½í®"ÕåJîéÈlc£Ø™_ýš˜í„@\7fWÀ#‘r²Ä}–bÍüdä…&³Êw¯•“Ì+/È¦½5ÇÂ'_•óÓ¯>íØ±£üfõà®0ÛaÙ‘)ÞpTS˜hûˆÆN×xy#Š«FkÇôƒç?¯þ€šŠ½í›`@K4tÓ9fðþ& õ§¶Ã8ÁÔ*!YÆx„—Øñ®²³ò4cA#&·ïïNèæ´Ž>¢ç[cog
Óï¯þ–.÷xfü#?É&É£heühþ™O_@‡+Íàtä§°Øb¡[Ör÷³ƒX7¥¿ÕWó¸9>*KôûÜ7(­Ž£“ÃäúÌ›‹Ê€uš9"ìØ§f0]èFÑvìÓ æŽ}Tì!*úH3]ãƒøÏ"G#•^]Ÿ…ƒÀù9_h¹cqP°Ú×§	ãcŸëó±ÉX[î¿Y`ÖÞÇ'm…ò&Q~u052uÚÑø?þÃÈ'6›#¸Îw;.÷ük|"_'²‘}J±‚â78 É:›²O:(È7˜?Õ„Ö}¬¢39#:[Ø‰úet¦#·ënîìoí~X#´‚l/-ô§ð¯þ7úïûcšŒgÛjdà}ñhà½ÞU»"”o:ùŠ¼‹û^(R<<,ënÐ1¤¬ÐÒ£Q·øä-Ÿ®Y?ÎW¬eÓ®Ö]™ë‹4à7vŒp°œ»Å²/VË`€áŒ¯.Zë¹ó‡©áÚðÃIè%àfú˜T©l½ÚW¤²~¿$¥ÔÖÄF^èù“k×>vø Ïkºb-ŸÑØgçr­Ëf(|nPPxf
Aý|ßùÃRV?±ƒµ„„*¬ì¥ìàÈn‹:‰t ¥5y”(äÉM´A-;ŠÊ}Ï‚b9­]7ºÙ9Ñím7ƒZÕ¥n|\¦.'¥mÙ!gÅíj'Üqö¶Ý}VøÜÝ¬Ç]=»*—´*_»ô´«ëgW9%›å¢y<áu©¨âY—»Ö9¼êdŸºç6zUÓ£î®øÓÍ©0"ùØU"µ(oÓÙî¶\íféhg•!Mò­
*ížáj÷¹ëBL=Ï»¶dç¥vã3¥âÁ¶PM¹Éµ¢6KïR8‰tÓÌYºÝv`§R@ùZÁÊµ‹<ÖSºÒ,ÞS]Cú©.^ƒðÏ~yòîPW´g.ßâ™µx²øý×OOO~ "}Ægp#£È8ÑúT
h£A¹Ú<”›)„®—ï{ª=;QÁ=ÜœüUóør¥ìµC”d›Èµ†—@²•Î§ˆÑ³ÇçéiÊlÉ;Ûž6-`V†)i;c
e0µ©¨$2;rˆLfÑ"×‚µ$*¡˜„Ã«•\B…KFª;‰kÈF(ÑT	GÎ9®š’³[x—iÉØ¡E•AE5„Ÿ6Ÿ\°±H>y§%”³l±;dìÈ)ÍH–!¥ŒCË._ëÛ‘k*dšk‚Æ3Çr2ºÃ†¦>§	ºÖÙw=NŠÛ¦f™ì–µûÌ3U˜Ó¦Ú&µ[„‡zÎ&_$+]Ï£djèhêSr‹@RåòE‚G¥ÈôGJ}[
aÍëB?‚G	<Ä²Ì@Dß3‘;!h3Sf5œ™„îñqþ\¥$ µ­mÂ”,Îþî†,+y¿Eq˜Àý—Ä«L‘•fÈ{&…·ƒƒÕRy>
çÎwÞªp.ÃÈL©û†¦ÆÛ”ÍxxÞ^èEÙŒEoÞþÉéJfÃ`è
P¸8Ä-bâ*7…Å"ø	¼ú½¥^H¿ö¢	ÿ¶éPÄSoÎ=FN¨«K_€‚8Öžø6ñ²I€>Ðöˆ]åy6¬Î:ûKºÏžb‹t¾At:uú:øúEãŠeè¬ÿý×ÿLŠß¤û1È¼ðêWXq²LvÂ`¿ž£w[@òÍaÞžŸ\ý(ö±én²™RÄ‰&£#?12£ zÓY±ðo:¯^½ª®Wz¾Fž=] Ï_,¯áï+#=¾sÔ @xdÑß8V€/ÀŒþ––¸&­Š¼ë5,“ú;Q¦œ då)-ªo¾[gÉA„%E9ª×÷ÓôË=NR¶x–ôÙnDXïÆƒ„½ÏQßlê#£ÇmÑW¶¡ä] ‹¦5 ÛùH:ëù×¼ÃƒI:†3ÅÕaJ=«º€â©S³î9+údô‰’š·àóô^¹ÖXîc¥(ý¾ÍªA^=ÝË±À tg@Þ¹Ô®¡&._¢TZ¢´—NñíøWT¯éã…>ÖþE2-^ ×¯©±lÕWµtøcòóÁ¤ }€ª«÷qŸÆ^šu;H >BÉÕ¯ð­H:¡¨ñ¤³@:)ë¶=*]âLR?1åÙÉ’ÏÆô¬2æÍ¸ßà¿£2ÇÝ&Ó¿H´&[p,.&Øé{Yÿ„ ñ2½÷?ý%?IâæØ»øÊ'*9N`Öý0žÖ`Îúá(å ô·Ì±'>í­G!-ÆŠtƒÖùóþÎáÖùgã|–ÿÝ…˜Ïå?›&ô®ßµJ2|—ŠS+2gX³ˆ½t«D§ôºsy9.E]uèÃbùIæÿìÉåå|NiŠ2rò×+)÷J-)'¹b‰b9oË‡öF„§ò¸ú+	@T?I¯~=õC¹]© ÝõÈ0–YQèoW§QŒüDp`öh	:Zeê0Þ¤] K‡w0=…ß‹ÛTVÇv]¹ÕüR¿‹ÏüdÃKý.Í²5÷{íÀh|ò{I|ú£t@¦¿mÌn{y	®1Ž×2†ñ>­¾W@L'×Z­‘§¶Äá¤û¶¥ü¶PêPÚp@Òˆ?›$‘¢ÒÕKü(G&-äˆ	>¿ÉËjj¢rlVÖ˜çÏˆó=U@Uy÷ÕMˆÎF©r†V”gÉWW~aù}UÎ°ÎàS5'Ua©„%«`i¢0¹Ò©ZhRÁtyô6ºf¿¨ÁYD¥1ÑƒA‹ Ç©ÃˆãJ6èA¤2´í!ÕúV->§\ãW©jœÈEh6ñ1‰$=é©êÒéÖ"e™ dÅÔ&#l•Iä.¼0ì•
ÑÜxXùÀ•¯”~+B¬_ðbp®W›³âÙ³ÚÙ2 ”6ÁfñëšÎ”uòJq‡[vÛŸVŠnñÙçöP[¤×]f—Œ%H/Nc-Pû/ˆÐÊÒ{ôéØÌwØk8K/yi!T¼ë^ŒÐ!Ûaè‹Ò‚/°´ ée~ÄØ/&œZÛFh¢XÖÁóòíöQQ–ˆQÏqœ¦“ÀÅ}Éµ»~™s]T²*Ô~³K°rI~º0]"É;=°ì5ˆö©NÒ. ]ÁRi/õR`ò§‹Ö£L,îX
£\À:Ùb=Æ0Bÿ¥þ$-Ø*¸Ãçs"0äY´Y‚0 ‡ÅÈ‹ÌF´¬XââÓ TœfS·Æ–ÚíØ|,Ûï|7V)Ú¼”q²mŒ”Š»pRªöY*öÉºi<^]ls+0§•En‚«…†° )'“Q^r œãOÙC\h÷‡“ÄŠÓ&i×êU„+úX!#\ÎwE·µÅ.§&r8×U}ñù€Í»ñiLY¡]š°J–^cçË£íFþÿ¾ƒÖ“y«èÎ#ìÔqÊ•Ä¸+ÝNW§FÍMÃ¶SÌ³qW¹WÛ‘tu‰lf"QM®~Á)ãˆŠë#ûð¢aœ?ÃZ5ÊB©\=üÌéîRÖá«N•Ipò§@K|Ð°À$´&÷A°êþZdfkÓ5k…Žvï&Å†¦¶‹)…Í xè÷£8Œ‡·@>+û#œÏÞÕ¿ü²5˜ðªŽdák’°ŸÀ×ë]¼ŒG:“mIvËUK%¶Áe•tÎÆ|ø€ESÊã¥—a1·M´©UÂYe—Ó‹ÃxŸ^_ ÅV!WÒ<TùÙh¨ò”nž”_ÂÉ*ÞQó;&•”©‡r»[ Ó;Q uåä™ÒjJnÝ¤úIgN2ó¢Üƒi%œÑƒQø¹}"*¥s«&zèåÚŸ$iœ ¹ÍÐÝ€d`%V
ªt¦ŠšÌü°ÈRO¼Çã­îñ&1R
†ÚÓÕk‡Ø®,ªÎ0|æÎaLIÞøã*ãÙ\å±—A_:¨6r}Ã–C›î,‡Æ³‡FN‚8ôj‡F…©¯€Ì&ö½Î…²@—aÚ«8nô)ýCN”´) ¬1ðƒs&lø\?(© 6 pzÍ$E"…Ó3ÇSýL=A¢%0G·ñdëÆàp˜Ôí:°¿FcdÒ„†j“Ž_ê,H;í€=e/ÕZ!mS¬è[Ñ…ö‘ê3æÒÐX ¿
Ì©Ñ„²_!›q®.Ð©ûòfÔ„z³ºŽ›g”ÃŽ^u©Îiƒ—šóÆú9“	Ð¥gzñÙVÿDS,iÚª×ý•lþ÷ÆÞ¢¹ªl€;ò¢>=è>¬ÕxÙ¤Ä!XÑ¡g60bõžçáVô½ƒõ#KZâ“èóû°¡%~:Ža> ð­'ÕË¹f¨åÔz,W–,ØT×äeRšF³ª3`ücâó˜º®Ã³l+L¿^žè\ Um?ÎW$ˆÍÑU°V†•þ è4Õ¹ÿB~¦Ž‚0»ã ÏÀ"‰ñ¨âLÕ$Q
Uu©Ž€3ÊÎ$rfýŸ.y#Hâ7^mIö–û‚·è^’ºåfN‘[yöÎÊÛÂÔÌ)W±NÄË&”é@PEf{ä™ðÚÕÜ5*NIZàqY/Z¦;‘JoSÌgF™BÆ·ùL|12¾ñwWÀ÷éy”î¯+ÝsŒv‰ö‚´Õ–ëÅw^¨g¦²&G`É¨çÍâü+ÆÔåk_Ïn§9â•uyGwÆRç!(Š%N†^‡%Â¡ MÉ1A¶T†J³Ý¡3;ÌØÅùD?¤§PŸ-¯-Sßˆqî‘’Ma†Û÷é‰–’?L@ÒB‹ìvOáI¼@>úìtIkfoËŸ¸}¢VÓK<÷gÖX])x»¼‡¶hœ>¶ªPr8Ü‚ZrJPMÖWÔ~Øüp œÐdvZÐ+ä§ÓÑä<¢ÓmZš‚q´;S”46|^ÿûüWòAžZÉéLø¨å>iKFåKÅëJÞYøÂƒ	¼ˆÒV$þ¨OKþ7Òòæ^r˜+=9¥—]ú™i0³51PÜ€™¥éTxËŽ¼=!|Í oùØñv³—ã­¬Ú* Ë{xÛ“§æÂ[ï‹Ç[áh<5Úäú“ÖñVŒímËh+©žÞÊjõGÄu na7ê9D—Ø<ñûA«¾½†±V¢²4ªx4†nÖây»Œô^¡—Þ“tÑ„Ÿ–ÁnÍZ×Êí¢ˆ¦eQ§—kÍ1KÄ Öæ¦KÝs¯è>_@¥ûÚ´Cê>¼½KtjÚîpiYY2¬“riÓ”i6fá^Ä»V²^Ü‹D õœƒ^ªÁÿ<*²*êŸÇÍpßÄ±dªEWj
ÎŒø…ÂâŽa5hÂN¶% ­Å<›8˜˜Ž¥)œmJópÛ“ Ýä/¤·áË¦C©›š xaÃ[Âg¨ƒ>­ÐÅÖ\XÊØÏåí8Í¿%C”¤¹X!ƒº$Žâaâ¼"NPSÁ¨>,3Aš°×÷àžpø`¾<³ xùØº•TNN»@Q˜·5HÔR+ŒÆq’ñô?u·q‡>ÄG«ÏOûv·±ò‰òV”Gƒ@~§²ŒU&óFŸµgd¼ø\‰zÜb8´ÀV–ˆ;ƒVäÿWêœ»	0L©zyUt“K™áºÊ=ëViˆ«:¸‰÷,.j¦ßÎúÏeN³a·K–ÄÕÂ!ò8BHÓý8I@NbDÏÄ{Z“®9Á@+ HKØWn'Ý5ûû˜­Íµ|?ZÌK"qƒ7—¡Öl	¿­]Õã±r¹Žpº	þà’º(¤d™X¦Ÿ‚ñÛ&ã0èã¶™Gg…¿;µÄB´šÝsiNYbO_â¼áÃZâBäŸÕ‘¶j‰‹v×YaÁ‘ßB3ÎÁ;xJ*LœbH´¢UÍêæRrTP}z§A§Ë˜®ß’©4ZjõÙª^f•ŒÐ¹˜á…Þ]ç¯ânðã¥õ¼õÁÈëÈ—‘fØéÂ÷5’fèô°@‚5ÂRÏÛ³îÐiÑ„;ÁeiztY¨ú¼´ç>¨[1Áp±ôhKúˆÓÀ?{ïe“$È>oÂ„µ™½]æ¡Ü'ë N^ŸæŒØé=@†ÏÎØ­j>7rjJ$ñÈ!0ö Ì—Q(¶0e6pÙÀUy(9ZÈ5N1ƒ¶Â3¯Æ^ò)Dý]ÎÚr“ •ð¢`„tt<êÔeMÎú¿-ìâË=Ê¯ÿ	}–Î€²Îº,´ý~FC}uÖïä¹N
CövøžxFŽ•„¾7À÷fÁð$3,^IåÀô''Ï›sœÕóÕ …ù›ë£WÙ,BÿXp­óÀäôVÚ  á?RÐ‚£&Xxd@,†WÃœ¤—Ð»X¹Rl	°+Ÿø¡wGEU‡¼?Ê£Ów¤JLÊª;ìJC~.ð“#›2|mÀ¿ÙBÁ?ÊÁxT€ñÕ_Iâ÷½'íQÝùS¿ì±€9ø%=JEþˆ )žpn¦Ð$±&A”ú	µ ºì'ÓA]£»®ð@™Ãø¥¢ÏïÏ×;ŸÌUé‹·¼%êë+u¦)d¡0ùµËV2K®iëÉ§Kº‚µÇœ¦JæßKâÃx´&ãÜœ»c·ìÝI¿¬i‚k,ùFx	]jæßlÑÉ>)M'$Ö·ÜÏ™Fñ(­ÆuÚðÍÛ\Óaº©#qu±í ¤Žè o¸]ˆmµîS¥<NÅ 7î6ÍÛ[Ã×–œ;q[äRJÀ”€°ø@{»pÑ7¦ì^Æ@9xŒˆQ+-b4æ!A0šËìY	ˆ8Ù	¦A9£Qg†‘ˆg%æÕŸ$	zÄ˜›œ{9°ÌËcL¸gAã„ØäÎ±ÐXñ“fI¦Ié+FIWsOzs÷‚,--Éc)+æ.Fñ¹´Ð”¤T¿v©^2°´j¾ÖÜšBcR{¶Ô)’´^7A«.,š³¶°U£®È¸ŒKEæœD›,˜t-%(GLVâLøè+wiÃO;-h¬!­¥ŒËò ›°<26? ²[?F6þK¾fB`!@”¿£¬%²„%GÚ;¢ÅL¤:Š3à…1šó3®Œ•0>ÑO''4'¥©z¶ƒÅBæÌŒ\þºÐÛv(«œ	;«3ö…kïó
gÖ‰a[ß†ñ‘îŠ„ÿƒ÷¸ÿô!ºðÄ°›¦òŽ”¸‹wYt2y3°¿,Èû½Î¾àû’t/òq]:+Ý­ƒ yý%{!&ÙN|š¢ã=8NWŸ×‹µõr	[ä|…lWRH£$R/T­
ËEoZõ,½JÖ¢½ëLcÒeþWÿŠš [=@[5ÀŠZ€¶J€´ ©rìÁº`“ÆmL¯‘]º jôã]‘`MT—êþÁOÐÂª«IšÅ#šNË
Kˆõ„ìÁÂ¥Kú×ÀÁRÖp`ÞM±\´0-¡PÕdÍÑä-™û0É’xŽV
2°p¾0™Zb92^´-ê‘ØÛësµ¨K³YÉáF<Â/S¥]OtM6¬4œÙ,Y/Š%	Í $„¶Â'°»ï÷AîŒîz‡J§ß"¨vÖéÒÝˆGÀÔy,ßÑ‰j9GÓÃ¨Ð¿¨[˜£]ÏÍ3Ã,^¶šaKŒ¶Ì¹.úcýi"WkÓ¥ ×Y§<«­PÊ…›$ˆF€³7üê+òd
¤§êqk§òJÏ›Üìp)=A½§è¨ë”•™‡¼È9@MAŽAb9Ð³H`ÌÝÍ]eRm¥VyT§ÆïcMzIiÐ†m]ìÇ"s6&ŸíÄ«FÕ^õZ‹³0›Ri.9›I‰ß½–AI¼A˜“ØoL†E‡e«»u;6¥|ªf‹»Ýe©÷šZ“øÃe[’½«é-IÒZS;’XÑÂœTdˆÃ`[‹m©Zn2¢–"é:ÇùÜu‚è<?ëy–h¾¡ápMú¹JA¾Vø^Å¾t±üò®·ùlüé»­ýÍkd+íÇcÌ?¸3€Ú²™¿õo–ÔRÆ¯‡7/¢3ÚâïˆÍ¢¥?'y,¿xJ=œk8,¿|êx‡ÅëB8 Qæ"B‚aL]Zl.–ü¡pð¡k×QÏCƒ¿`¥ëÌÙâr²Hñ±¾P¼oÔyHnõ›ç"ý•ºÑúRjŽ$øÑ›UÁg:’à“§içl¾ÞôŽKÜ+¹õ[\xIX½z¸4ÈÏ GùC‡ÈÇ>FÁïcO„LZ›MÑy’t–¤¸EJìL5Jd.1‡Ÿ³*;~£k UL~1W@4«ó^hEö¢Ó’÷ªj”Le×ñcÞ—^¸²."Ð?‘–JðÃUQ µ¬å?×é—•mµíL?»g«zK8s Êâ‘— ñ§Œ?ÔyÛ‹ãÍ”ùâÖ‡¯Ã¼—ž‚g¤*ÆÑ€Õ0N…Aökaæñ£àµµ5óÉ(ïh™¯Ñ{½Íº‡ÉÂs¡™>kt-ì×ÜÜ¼«W>‚ƒÉz¦Éó[Ir½†úåá¯C³…üjÆ(wJ/±ž¬ÙÌ•FY	?Í©(éÆc–^q¾AÕêc#¼¹€™š@MÇÍ¼(‚%.°®€wØ2à>K¡· íÌ¢ÂŠ‘h§tÜ!26õ`N¼t×¢o¨¿.‘Õ7¦!Dü.ŽÒ½xy¡ðÜÃ’=‰?ñIÇƒ„>³$­Ÿ2^AZ B«óqúÉcAù€/m5VÛˆó5_#æÇÄµ]8oæòÎË÷ö¼½HS*ì¨~ì•¯µa­XÈÔ	j2lsÆ‚ÿ\¿×í’,R\3§›d<lõšï¿0pxõk6	±¬˜€f7_cÕjâ§B³‰3ç³ôâñ>Z¾Õm/<¡º* YÓš6wÿ4iÀ@•Oñ–ä‚‹¨'~hŠFüÜ9nœ¥¥g¹¤ÒI˜ÅI0c°Å
#o:È«¤‹)çÒÅcXcÛ?H@ßÙíí^ýKï`lî÷¯þÓ·;=£¯ûèP^°{-A8c1ËÐ½”Åß!ñúM`ÚË<Üo,[^“½”ÖÈÆey(_ãú¼ïíÿãÖáÎî·ä+òqkw³É³ö¶ö{Wÿrõß¶àé½ý[t°¿µñÝ><AþðÝûÞn“'·6v?¼û «ƒÚ9ûùá;ÒüÄ6^¼ËšZ9úðÎ«jaœ…®VžIIY»jWÖbì,œ¥¬jÝáÕ_û0ìM­‘3Nñ†4CB.vœD_>üÜƒ3ãúGy$ø™ßw"eÏ?ywžÈlg)«š©¬ƒ0Ó÷†‰¦©Ã1H–KV¸yðìñÔ…·
-Eö ôù´ ú§‰¤ÄË‹­ú#š
ãC1ïm] õP5Ô"hÒþ®•‰3„åtsgkçðÃµÇû?2$^ h¨ÆÄi´ž |kb=}q_¬§úvÞ™Õx³(ÆÏì§±¼ÕÈ&ÞÏ±Sà8ª¦QóäÞ†X¹—1´í(ª	±0”¨½¿úë ðÚSQˆÎ[¶‘öÂÌƒ0ÄÈ
·¦P•§èü:ëìoÃ‡¿ñ‚sx–þq?z¿ ™—M\œÃ—ˆ)]Š{À[!5n+X×-c„$’°o{kDzÀC²Üíëò¯†åY-¡›â{ëør¼xµoG"[…f‹:éõ» Í>  ûQ¡yUÏT ·U‚ŒPÃ§µqæXÏì
…gºÔŠè"ü`„ÑÅŠŒë%Ô…Oÿw,ö#_ŒåOíÄ„·š)5©!m Fue¬.gà˜n&: ãìŽáõÔ†ilá¾-™ÛÙÍß’•y†RS<±™fåh”±—¤þN¤–€[ +O­ÆF7¤c]#€Ä.†±`N1é5ÌLu²4òÎ»OèaVÖÔ5…%s6¢W†dD£ŒwÖ7éx0žÐ‹i‚å‹ôÂµ`ÎsLaŽùŠhYqø>¤E°Yü§¶9ÞEÆ¥D²]ý
ÔçÚvô•nft¹×<ê¾w"¤Û­šMŠn–i‚Úv0z³æ`ËÖ:ŒAŸ*€±ýˆš > âÉo®~9õÃGßŒ&&kyåêcBÂžò0o`›Ø õû°HnÉD·9Ñ-V¡í(ò^ãEå*Ú|Á1“wþè(ñ3_7"ÌÂ`¢ñ5ÍãI2ý›³|ãpqÀ3]PÇ¢ (ö+a§±(ñÕ«•Çg‘–ù,ñÆL2‘iG“,‹#‡Ú€50cxýOR€§åª~Èg¼5ò ¶²ßué–!?ûS¼èG6VKõ’<0ÏÕròœÔ®(™
Âà–VoOÇq@;=Y|ÅÓ-ñHn\ïµ‰Íß:wK{‚q¡…*%·±ê(eŠHkrž€{Ä×ÊV,UjÜç%=5ú|‚™í×$7„qã,Ï{œn¡A³hÉ¶qYŒ”ŽA¦é±;è>ñ²ÔÍ°.î~¡àžO¿â5‡¥ÈËwïÌûiêýƒŸ&° ÕÀO×¨7ßð÷ÒÏQŸ˜‚ó‹HS{ž±ü&Øß›¤'6í#ËËžþê+2·‹¤æq4G‚ˆœaù»3{šB1,ï÷Î¼ #r/ Lü4Nyi
WºŽÀíŠ¦²dþÀš°‡}<<Ð»½xï4NÈ‰w„4ÃZŠÉEÄ3z´DXÊnB"A† Á	Ë¼–ø}ÿÏã>À‹±…L—:®ËÛõš¤7l-ßÔFêÜ‡NÈpiªˆ˜Äœê4LºyHØÁÈK²ñ	.]%ýÂ¥iD»l
¢‹®•a*‹gÚéb¾³dÒÙé¢ÂÄ8Há‹œÕ“\
€*À«k:RÊ‘8lPŽÌÇ©WúÖ©È}ìôÕ*ô#ºž,¾$e)‘ÓZ*°€Ö¬ÛÓØ¤~“Ü/€c3ðÒ.Å&ŠŸ*)ßFKóþ$(ÛŽ©¦ü‹•ËõŒ¦d^å =¾z¹¾Š§S>ÿìrýÙužq¹þâ:Ïÿîrýwžw918t{–˜Û[ÃøxÒË¨Þf®ÀYt™ E–sfÒ–O2!ÕZZÒX·@jP…UƒRÙuûà5o€c/þ÷>®âV ‹'ä·/^¼|þJ|¦¡1b§œjÄéIï¾FÂ\ñ±GùXÔÚæ†\Î6fë…ë•gy´šÕfÉ¶¤-•d6$»é
ô–ZÌ¼Ö;¡nŠhVM¥*ót—mÕXà«½¿Í:Hâ-,q]œGQ—ê ‰k·T)Ÿ©­kÐÁÂÍë ñÔ:Höî¤_ë IK^·’<Sy^Æˆæ¯+µÅP~óÎ•8Ú¤óàö“¸*~ÿ>¥‘¯È Ïcî3žÇÜ‘õý ‰õÇ“ø,EVQ»!•pÁGÝTáE/"hqe:¾«‰æXÉ,âÈ^›×âµ.zÑ ·ËË°´EW£uw×€SØºaóŠWÕ_r—ùÒ+0íŒ€F[Ü…,þ3wòkPÐ^ˆ|4Æþb¼ŠpšècÀÖ¯<àMMg,Jé„‡C§« Zþ€§Cgej´[ùÛïd%7Ô+!°íUBh0„ÙVBx,ZÐ´hüÕ¢îÂ5pE×Z¿¼‚3/S ¡û-•)ÐÆQY¦ 	m›E™‚òµCo¸ƒ÷Þ‹¼!À†ÒEæS…ÁKÝßÿp©¶–Ö;òÏ ë´.#‚¢Ý™>„ì‡ÚµÌ6ÒA©WÊ“]ž™â±P4S=º©ö•FÐM«KëoÑEú|1¤¢>zÆOøßp‘Oï­QOÀ“)¼-Ï–†ÿÇŸf>Žý€‡ÑkÞ±Ã¢ø|’W‰-—‘ŸNWj'QY¼!8å%ßgN†Òêwœ¨j;/‹Þtû­;«+ál}Nž%_‹*ã–î2`˜xÉ†n52©HFƒ<W«ÅhZÝ‘lSE® Ó¼õ;ëRÿ¢ÂcG]™6iêËþÔòì_³@>¦û…yR¯G$°.os$0å?nÌf£!Ìö3qûZæ³ü¢ì“„N…ù¬¸XâÉX¡gç¥Ð®›2ªÓ7ÛÔêò}kjQ3a~ž¡®]{š¼uÍiÒ@”4è¥[& 5“ ‹Fªù10l½hùdŸlàcê7©TÃ"‘b¬R¤eöjdö6ûç8U’6
ä7B‚„q7ybä^6aëauÁqdëqÆ–ÝcoÃ–SAC_–¼[ír¼Åà#°ÅÞÓP¬™{®/ÙC`’ª4&a9âNW˜^Ÿ’ÑñÆniªcgšÊÒt}[“EÅgQÖÙ3ß$Úæ´¨Q6øÒ1ë¡D}K<pQ3Âp]a…0¡MîBËÑJ#cQ>Hq¤5ã2è€^ì!•Ø.£7vUò¬á—ŠA&è}êõa¤˜	(´+U4"²qtÂd@ÿ‰J›‘ã†ñ‘î
b>xïeý“î˜Þ/±LŒš›ë¥PºDßQ"?¥Êg¼îÙØ¢÷áSlN{—4~§H2-åñ­due
É5Fò“ŒùÙ™ïG¦xéiL«Tfž¦‡£%ÀÏS…þúo¸vÍN­±gU‘gâÎh%'3å±ã9Ó’}5
ýã8	Fè·ä·b5yêc0Ì8ù\8—øX;#[úìÌïíL†Ctê†V%vHx|ÿþ³Ž—ˆõî
^SiŸµ!Q©¢£¥AæXã35V8£&Zg	3Ì½ü—¼ DÞ™r½º£1¦ÏÄy?åËÕÝHiDôÎL7«»ü4ò™ÄÅË—«»¡‹
§b€ÓéûiZŒK¿Åº³öæ¨øæ²›ã‡£+ÊãÅ4ÉZBpiÆózJª5²/ÊjÑéÕ/!æ”`²²ÌƒáÕß¢~à‘^B™²À(jFÃKÒÇ‰tý$Y#^ô¹’FÔøˆ:ä·´Û¥‹õÅ%…Ÿó×;èKLdQ.Ð0Yâ}Sh™¿Îƒ]ù‹øA~‰§ë0&n„{¥09\³˜é”¹*÷»Vy§Çbaµ5?¶m/;+èŸÞ™‘;:€?…Ð^|>EòÎoéx0Ä‘vÃóhÎÌb÷
¶úÌ:|šnMÄh$Y„°£Ø9ÚÖÎýA…¯u³öOQ­BËW°„·(çÈH¤Å˜ÊÓ ¯8•câcò¨sËÆó¯ß:dCE1‰ƒüÔ»:5û¨£¹Ïû#t3èÆ$Ã4U  ÷þŒ“«_2nRŒ>Ëóøxf_ý’`˜#¥“Áh”žéu^a˜Ú v7kÐ·Îý>Gí®œEnüãÞN:OüYªN¢8pëªcáÍŠ8(ÏFàCÕTÇ'¬©ê QµóØùZQ|ÏLàkMe•‰O¯¥¶ªD	³xP#8Sc?ÌüÂl ±*‡F{•Gp,>8nöö–7zÛÄË&¨	ƒƒ¿&0–å¼vÀP)[@½Þ ò%PÔ‰ó€Ë˜>QÜ÷Û¡Õ4;ÇMlËÔº¨8‹N|žÊÑjï¬èÅ¸ Õ
)pæ8S/À¾UŒyïG(3È(Î‚Sæbâ'¥ueÜý(}·$ê7`•®òh³LZ–k`×Í ±Š†]îïÎlÔàÒÔuŸ]xr„È ¦:ý³ œŠköoó 4"_ß ^7F§Q”NÛ™!J‡?P7H§RÝØlö0vípð“—ÐÂÆ È÷ 9vr0Õ@ìßi)
hVq@×2;
LäÊìsaÅ45 ÈÖî«¯ˆ+VÈòõÍ³uY#"ÈÅp×H(oºç¬÷u‰”P"&µR•§A4‘O¤ÆNÔ˜7ÎéhÐŠ7ÈüS¾V+Ò¨„/ÓU,ÇÝ3™‰yêf¨¼‡˜¯ÔÇ4fÃ(N³ /	J¯÷ÅýÍüvLÎß¤è^ûæ"Hó¶¹Óm1Z“o®Þ^sÒåˆ!¾÷ŠÇp_þU´I¸ä_¥ücoÂ¦M’>àÎ ÛñQÉ—å‰äÃ–ÚJ«#š.¯ý§åµF_yÁÑï™÷¤›bq«MÖZl™TÐ…×åëþûrCÆa,w± ˜ÔDÂØþ$IàPÚ¤k2(=õ––20,O¾·»1fl•Þ,všy^QoA_x^©ó¶;ƒù·KÔµ¶4	<”ý§AÆ6j—ØDb—-%êV¬Jþ¸¯áå,!­é%{Ø€%g3Ý–·óÀ;õÙê"ÁÁoY¢í-2ÐåžTþ™y’ ^ƒÐÿŽz›K[Cw¹Üõ8“^)CÈ¥
Ì³®Ÿ£¾~Ëª©…Þ^£Ð)Ün€öÏ¾BG“úd¼Ÿ7c?9ÂL{¦íðrÑX˜ì`2‚»ÀdÒp›_(ZMR?¡9bß\àW„Iø!y’Ä	-ðì~5¬Õ7^ÿÓd¼u>Ž“Ì¶\¬MÍµRŠ½Œ/Uþƒ=W´ÃB¤@:èe¬+Oàã}Üìö6ÿ]´`×ÿ˜âš'ôUêy›ñYÆÞ û1Q.„nõy¾áâwÑÕ7ùÀ´'HyÔõ!ûb<8þ.	•ÐÕ£.8÷$ˆ`7©j€&Ú\þú©¡8ˆRÐûçE^Ô8Hâ1<;IuEMäÚ9n3²øU0áÞùâÙâ×pádñûW«§'?¢ŒG?	zb‡ðÐI0ø‘37$8käP¨Åo™Ë¢èoªó.[/õÉ³Å—4K¨”þEHrÆ)¶¬j÷ÌÕOÈ0
G8ƒ¼4¦²¶Õ9y¦¹ŽàH¬òAÊŒ yv¸ý¼ìäëå“gµm=Ú*Ó26¦…ý†:!‡\bëð8-L4®Ô¶`Ùõ—b¤N¡ÈÜ_ K¢Osó—ÆUÜ:‡Ž¼ð]}’VRŽL`x°´¤w”`HâˆÐÀ²Þ‘gXìolÉÁ›OÆž¡&8N<x€]èOPp_A1ùL]A³8éÎ±ßsw¤½Ím²—ø§ÖùanMó8|ÿngm1£»-Á%öm	“ãAû?;ë˜ÚbÉ÷(3+pªÄœ›Z7Aqï÷–Vß½í;#ô¦’ëì6ö çûT»íÉÍ"ìø¦uç<ó
xK'‰Œt½ÌMü<ƒfs—c´ ä&èÀ[‚§çÌÏöqè¦½´¬´8;+–š–¸O l“ÅÖðø\PàBé3:_\1zÄi{5<ÆÕ±S˜ùƒ.+“]t¼"¾WòR´@óAü_™Õç*nž}Åf²‹Å ÄúPïG8:Õcš’çŠ£è5§i‚¬/]—Kƒžî„ýÉ
ÎDÀ‘SE»C„¬2‰¥¦¸~j‹×‡1°à4å4Pãdiø	Ì)û{‰dbÞË1å0mŸïÑëe`Ëé®ðþ%.ÿáÿ  ÿÿ æöo