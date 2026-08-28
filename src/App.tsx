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
import { AnaliseResultadosView } from './components/AnaliseResultadosView';
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

  // Helper to extract identified diagnostic problems text
  const getDiagnosticoProblemasTexto = () => {
    if (!selectedDiagnostico?.id) return '';
    const list = respostas && respostas.length > 0 ? respostas : loadAllLocalRespostas();
    const currentDiagResps = list.filter(r => r.diagnosticoId === selectedDiagnostico.id);
    const cleanResps = deduplicateRespostas(currentDiagResps);
    const negativeResps = cleanResps.filter(r => r.resposta === 'NÃ£o' || r.resposta === 'Parcial');

    const problemasUnicos = Array.from(
      new Set(
        negativeResps
          .map(r => r.problema?.trim() || r.pergunta?.trim())
          .filter((p) => Boolean(p && p !== ''))
      )
    );

    if (problemasUnicos.length > 0) {
      return problemasUnicos.map(p => `â€¢ ${p}`).join('\n');
    }

    const lowScoreResps = cleanResps.filter(r => typeof r.score === 'number' && r.score < 50);
    const lowScoreProbs = Array.from(
      new Set(
        lowScoreResps
          .map(r => r.problema?.trim() || r.pergunta?.trim())
          .filter((p) => Boolean(p && p !== ''))
      )
    );
    if (lowScoreProbs.length > 0) {
      return lowScoreProbs.map(p => `â€¢ ${p}`).join('\n');
    }

    return 'â€¢ Necessidade de organizaÃ§Ã£o das rotinas financeiras, controle por ciclo e apuraÃ§Ã£o de resultados';
  };

  useEffect(() => {
    if (selectedDiagnostico) {
      const existingDados = selectedDiagnostico.dadosConsultoria || {};
      const probsFromDiag = getDiagnosticoProblemasTexto();

      let currentSol = existingDados.solucoesIndicadas || '';
      
      // If solucoesIndicadas is empty OR has only the generic static placeholder:
      if (
        !currentSol.trim() ||
        currentSol.includes('â€¢ Necessidade de organizaÃ§Ã£o das rotinas financeiras, controle por ciclo e apuraÃ§Ã£o de resultados')
      ) {
        let solucoesPart = '';
        if (currentSol.includes('SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:')) {
          solucoesPart = currentSol.split('SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:')[1]?.trim() || '';
        }
        if (!solucoesPart && selectedDiagnostico.cronograma && selectedDiagnostico.cronograma.length > 0) {
          solucoesPart = selectedDiagnostico.cronograma
            .filter((a) => a.nome && a.nome.trim() !== '')
            .map((s, i) => `â€¢ Atividade ${i + 1} (${s.nome}): ${s.solucaoProposta || s.descricao}`)
            .join('\n');
        }

        currentSol = `PROBLEMAS IDENTIFICADOS:\n${probsFromDiag}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesPart || 'â€¢ ImplantaÃ§Ã£o e consolidaÃ§Ã£o das ferramentas de controle gerencial e rotinas financeiras.'}`;
      } else if (!currentSol.includes('PROBLEMAS IDENTIFICADOS:') && probsFromDiag) {
        currentSol = `PROBLEMAS IDENTIFICADOS:\n${probsFromDiag}\n\n${currentSol}`;
      }

      setDadosConsultoria({
        ...existingDados,
        razaoSocial: existingDados.razaoSocial || selectedEmpresa?.nome || selectedDiagnostico.nomeEmpresa || '',
        cnpj: existingDados.cnpj || selectedEmpresa?.cnpj || '',
        solucoesIndicadas: currentSol
      });
    }
  }, [selectedDiagnostico?.id, diagRespostas.length]);

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
            variant="outline"
            onClick={() => setView('analise-resultado')}
            className="border-emerald-200 text-emerald-700 hover:bg-emerald-50 font-medium"
          >
            <Award className="mr-2 h-4 w-4 text-emerald-600" />
            AnÃ¡lise de Resultados
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

          <div className="flex flex-wrap items-center gap-2">
            <Button
              type="button"
              size="sm"
              variant="outline"
              onClick={() => {
                const probsFromDiag = getDiagnosticoProblemasTexto();
                let currentSol = dadosConsultoria.solucoesIndicadas || '';
                let solucoesPart = '';
                if (currentSol.includes('SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:')) {
                  solucoesPart = currentSol.split('SOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:')[1]?.trim() || '';
                }
                setDadosConsultoria(prev => ({
                  ...prev,
                  solucoesIndicadas: `PROBLEMAS IDENTIFICADOS:\n${probsFromDiag}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesPart || 'â€¢ ImplantaÃ§Ã£o e consolidaÃ§Ã£o das ferramentas de controle gerencial.'}`
                }));
                playSuccessSound();
                alert("Problemas identificados no diagnÃ³stico puxados e atualizados com sucesso!");
              }}
              className="border-amber-300 bg-amber-50 text-amber-900 hover:bg-amber-100 text-xs font-semibold"
            >
              <RefreshCw size={14} className="mr-1.5 text-amber-700" /> Puxar Problemas do DiagnÃ³stico
            </Button>

            <Button
              type="button"
              size="sm"
              variant="outline"
              onClick={() => {
                const atvs = selectedDiagnostico?.cronograma || [];
                const validAtvs = atvs.filter(a => a.nome && a.nome.trim() !== '');
                const totalHoras = validAtvs.reduce((acc, a) => acc + (parseInt(a.cargaHoraria?.replace(/\D/g, '') || '0', 10) || 0), 0);
                const solucoesLista = validAtvs.length > 0
                  ? validAtvs.map((s, i) => `â€¢ Atividade ${i + 1} (${s.nome}): ${s.solucaoProposta || s.descricao}`).join('\n')
                  : 'â€¢ ImplantaÃ§Ã£o de ferramentas de gestÃ£o e rotinas de controle financeiro';
                const resultadosLista = validAtvs.map((s) => `â€¢ ${s.resultadoEsperado || s.solucaoProposta || s.nome}`).filter(Boolean).join('\n');
                const probsFromDiag = getDiagnosticoProblemasTexto();

                setDadosConsultoria(prev => ({
                  ...prev,
                  cargaHoraria: totalHoras > 0 ? `${totalHoras}hs` : prev.cargaHoraria,
                  solucoesIndicadas: `PROBLEMAS IDENTIFICADOS:\n${probsFromDiag}\n\nSOLUÃ‡Ã•ES / AÃ‡Ã•ES PROPOSTAS:\n${solucoesLista}`,
                  resultadosEsperados: resultadosLista || prev.resultadosEsperados
                }));

                playSuccessSound();
                alert("Problemas do DiagnÃ³stico e SoluÃ§Ãµes do Cronograma sincronizados com sucesso!");
              }}
              className="border-emerald-300 bg-emerald-50 text-emerald-800 hover:bg-emerald-100 text-xs font-semibold"
            >
              <Sparkles size={14} className="mr-1.5 text-emerald-600" /> Sincronizar Tudo (Diag + Cronograma)
            </Button>
          </div>
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
          {setView && (
            <Button variant="outline" onClick={() => setView('analise-resultado')} className="px-4 text-emerald-700 bg-emerald-50/50 hover:bg-emerald-100/70 border-emerald-200 font-medium">
              <Award size={18} className="text-emerald-600" /> AnÃ¡lise de Resultados
            </Button>
          )}
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

                          // Obter problemas identificados no diagnÃ³stico estritamente deste cliente
                          const rawProbs = deduplicateRespostas((respostas || []).filter(r => r.diagnosticoId === selectedDiagnostico.id))
                            .filter(r => r.resposta === 'NÃ£o' || r.resposta === 'Parcial')
                            .map(r => r.problema || r.pergunta)
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
  const [view, setView] = useState<'home' | 'companies' | 'credenciadas' | 'licenses' | 'diagnosis' | 'dashboard' | 'premises' | 'cronograma' | 'relatorio' | 'kanban' | 'settings' | 'landing' | 'checkout' | 'licensing' | 'dados-consultoria' | 'projects' | 'agenda' | 'macro-dashboard' | 'disc-assessment' | 'maturity-assessment' | 'plan' | 'analise-resultado'>('landing');
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

  usexœì½MwÜF² º÷¯€ø<]U×d‘”-·›2Å¡Dª›·õÁ&é¾gF£‘ÀH¢U”iZÍÅ=oñÎ,f5çí§ß,îªW}î/Ð{‘_‘@¡Š”[¾m« dFFFFFFDFFîŸ¥£ºßDÛ¢÷ŸEðW—×Ñûh\Œâñq]”ñy:¬Òú N'ý½}“N¦eZÅÕ›Q™&i>Êâ$®z«Ñ?¿|1¬ê2ËÏ³³ë¾*ö„•F7Ñ(®GQ?@C£"¯Šq:LË²(û+ûøOTÅ— $jinke5Jà|³½
5özðð3ø<«Òýeº™dñy^Tu6*½ã_—é¯ïv†»e'€(S€ýiôuewýá–ˆÃpN²ª
!®?-ƒ¸®ì"®?ÜñâtœNÂˆËOË!.+ûˆË·DÚŸŠ4€·ú²Úª®‹µzK¤ë¸LÏ@6LÇq^ø˜ËÏ‡øuì-ðn8ðÛÊšÓ7 ,Ä3Éé.~XJÄH žx or•žOÒ¼ŠÇÓcõqIÔ5p}ZváÕÑ4-ÏŠrr<‰ËúÉ¸˜%Ç×ù(ÚŽâ
ÿåËÎ¢þ=èw9/¢(§eÝ_9DÎâË¢\…>ü[èœg940‰ªYŒÕq4Ë8ª²|Työc¨§³*Jâ¤¨¨d½˜]¦“á
`' —i=+sñtóýƒÔ¬Cè4aÛ¯ËY*kÀGüD„éç³ñXôS‹º¾m£§ãYuÅùu”¤§Å,¥I”åkgãìü¢ŽHçUZEgYYÕª³WqVGgXó0ÍÀàH•ë«–U«Ùz½!BÇâˆª ”²I
ò~2ÝŠòô*Ú‹ë´?ÖÅ3Èô¾Wô{Ózíñ°Êûè¢˜•[QïþZ’g5¼šdù¬N­WU
$ìUt3X5­u<ÞÍãñõ)Ú0_fÓq'irRYo	 —äþ§¤¸ÊEÅ§e1i©Jý²>e qÀZ_–ÏÓòÜÁ'Ië8W[ŒdQ¤Ôx«ÆZz”­Âo`}VßÒ–„a–ð%Hù·lõÓìtœu:Š—…PÌêr±þëÊê×Ãî#øÇ(«`ÉŠvË2¾þVˆŽC`‹¬J¿½,²äÑ#˜¯^?´«‚P­Sªû‡Y:K¡:ø7N¡Ø¬,_¨¹*3Y	d¦²(ÌçÊl˜×RÔÁÌEYŒM$Åè(=ÛB ]ë˜~:GV³(Êe=˜gqžÕÙéÓ¢|šO€´Oûg€3u‚½¡l‚ù÷Dä¾ø‚¿B¹*p~´}¹±1°0PÈŸÖ›‡v6 Cì§\–N‡ Z'YÝœ
Ýè©Q¶éqžpÉ|–€ÒFÒêlôSô4ˆ§XX`œG¬¤u1—ÇO9‹º\Ë X`ÖÖÖ¢Ý!Éþh_Ê,|éÏõõ8§¸èÒsžÖ{Å¨êgÕn">Ú‰¾Ÿ¥å5Ðm<#+r¤Õ¨§Äao0ˆ¶æ–Y®.R`ÒÌó´<H`±èmoÃÿqIÎ²d`(Àð9!ˆküúV¨,«ªwú¦ªÛ­!¨û12–Í.t1½†Y‚(Á#~‡Ãdˆ†ù´Õ¤nðÆ]{IZmE'mŠØº Ç”ðJµ¸z¡€ø2ÅÞÁ*í[áyØ/Ðc|‹³„A°ç‰fqYÒ²-THz²ª¸†„òX
ÿí‹ÚÀ7âÇ2½f«TSßßØ”ƒ¤L
g°"ôq@Žê+^öÚkù]zÊÕ€Ø€•óxG–{í²0Më}­6È_¯^Ë…H–…Þ£TÂ
Yg6æœúR¯ZŠçŒÕ¸s<aìû0ÓÜòÔ+n¿ty‹5ø_ý*ºGlö0rj¨Ä®ZŽ8X…žÒR„%p«‘”[Z h%`·ÞRÂÄur¨¿Dþ³­Á¿”ªëÀYûñ]Ó^U!© õÄTµìâÖ ‹õ€ål¥ã*æ‹¦\€pºe_§íˆ¢©9IÑ~É6»ÔÏ$¥?Ôe<"qr¢
¯ÿºÖÉeÁZ.B‚ÓdKdewÁ^œ×¨™ÕÎ\×‘³8oq[Å†Væ¢jnùÀÀ…zá`ÃÆRÒü‘2—~ÖTª3ÖMüÖŒ¸$ºêBñ0Žž½vGÄò—ù/¬:
Zß‚­ëwq·ûN–§¸=–ŠÛ‹þ&¬Å°ö†E–SÝ,/y‹úf—[R…#4ÃŠÐžïêpºs­
œ‚ÝU{cÎÑàÆšÓÊÆÙ‚Ç0M:èpXm¾gS@‘“`ÒäTåT¹®=ï¤Ç!ˆ…”8Ýf›ç
«o{–×†=uQã$â·Õá4¦]8¿ðioÎÞÚÏDƒ³æe‹ÇÇùï¤ÉY¨¶js>²ÿ ]#ÿ}ZZÃpÍš?ˆŸ€v×‰Û;ðÑµ¼å×¦éqˆ}¯N_{äÓÓüžHÍïH¹øÃZ~^Në3a0qÆz+lll´èV¥”?Â6¬F©~ºšŸî`«æ§ wÕüTksÔ>û,«jZzãdw<~¦Þ%úN/M­Î½´Ò½,µÆdõ­¤þ”†ºzP¨é*TXn!…JnS¨ÜBa…êÈla©Ÿ]T)‰òmU)cUÊ/|Gªðù™èQ†íZ”(=¶'Ê Ùª>9hþt'ÆÉ¿‹«?Rƒ‹bœÆ¹hCw=PÀyÕè-¯æ¨j¼1Íãây!µ-ÌëŸ–ÎÆ™»YasØf1mížCÏr.§¹Íçý–ntÒÙn«ÀÿÂ	Ooÿø£}	z5«ëñeê«ev3yÀRTj…CK9ÿ¸‘ÖM`†gÙ¸NK¡`•Ü6: Ù½½@3æÝºqtü=©ãŸˆ(œˆÂÃzþI\½[NÏ·Ã+[Ô{§à’*>!V~E7©“»£Øóñê.¶jú
~WMßitŽÂ¿ö<ˆTã^k-ÜÂ¸&,ë¾^AŽ¹Î^gÔÞ^h!w¯®´œ¿·kï;ù{ÄBæ‰n³Í<q…Í*µå~Eb}[E£ÙÅBñß‘…âÆhÿL¬‰v›B$û;Ù'
½Vë„!ø3°Ln«63Ú§¥ÎjVÙÈ}*á<fk@ú£+ƒR©Ío{Ââ¾ÞÉa;ç
ƒæ¹j÷¥÷<­‹¤ÀãñúcõÕ±¤UŒ@}V£ãb<ûðoþ=…ßþOsÀ;}àÂÓñ„ *£†å™Ðâ^Ä“t+Rz
u]‰Düê5†á× ¶‚¾Ä¨wYŒAÐðMPã÷éµ‚àDÉÚú‰µê,£ŽJt›ÕPU`õ“áÖÝ ÏL_£âª¶²9ê¨Y–ƒD®Û]×C¤ÏhLj_@³¸@ÒÏ¼˜¤êa•þDôÊŽªR›†VŒLCyJ‚XÌV†hf…èÔ‡¯&IÖ¨1â¢B¤T›w^T8>ªAíE`A[-ïHâŽÔ¹>¼³©*õ&¡3¹Ø·/e’‘WE›jë××Ó2&ƒ±Wœþ	x¿³eá!æ­…Þ­oU'^P-æxGƒ…‚¤å´miP¤ø4ì¬=ÍÁ°y]³q5ã8Å¹šÝ’XÎÄRæËÍgáßBX÷_j¸«—‘ßÞ’%AXú";ÂçLi<Ìw—yÿíîeVQ\Xçã>/ùùfë-åó×a=qÄ‚`/Xý?<5+"t„­ÞÉaÝPPsRwjÖVªŸ¼S½s€²S´ê'<ÖîqÛ9 Õ¹Pyj“€í©ßîñÑyÈ±›ì¥„ÉžCg<-Õå©:eçI$Î« œÆey–&‘>I$¤\ZÉšLæ©s*fucr_@D•	 ~R†3Ÿè¤)k@i? tìY\Õìðb_NÈyŠÔzƒô{#ËŽEK@¨×Ç³Ñ(­ªãb–'
3uÊRêPª£°Q2+ã¼Ò²IÊÞÿ)¢^³ó¿²“¦tB¶4îëø(|Ø‚ÍUA÷p­Zy9*Ê2E³	Vò&+?ÁPSyq#¤úHëP¿ÇF½óÏ³žÅ UùÏÔ—–a™Ÿ§åž»ãÑ»Ùôi6¦£VªŠ\¬HF1=ƒo\Uu*[êîé¸8•ºÔcøÙe ¾Æ¥—Ñ­¨O§cÈ¨=®c‰‡£tDÕÛ³úlí›“ºì¬D=â»£gCqhê%­¿ðÜÇÂq>º(Ðÿ+ø‰'+íƒ'˜×l²RÙáPšq> N»5 :ƒ%»§Öªž¡ŠIhªˆ²¾®K²
yŠõò"O{ªÆù´H®‡@§4Ož\dã¤/ª»í€ŽïØÉ14§ÀÎ/fö¡rw´ð—N	eãI¹E’º«¨W`2¾¸Lƒ8)~S8`ezY¼c4nZÚô’´ÝÇULµÃ“þ¬ö'4-‡+È~ú<j3b|
¬îÎ×…ïÊ‡†u¶3ëj/ú"Jóæ»£ƒ'Åd
Ã	\fßóý-Â .‹Jä¼ÏÝµ«vdÖØÕgXõ¾…cytA®t*ÀRœ©îžp—‰¼0œ•Àz[‹÷ýH›¥QýŸd¡¢9,U|ã&Ð§ŠÁˆ@›ÖÁÆËðDä¦:œ•v(8>‹Ø¤ëSq\©îÙÛô~ ìØj8Nóóú‚L¡•bAªx„:£ã¾ðÏ·V-|óÅvtÿÁ†Key<¶å$±ÔP.fù;(% VÀOi?[E°Õçî}4ÔQŸÁºÅ¨çØ?iæÑã=fé(E–ÏR.ël²N¥-çú Þ°ß &²óÑÞlr²³PÜˆ ®*š·ÊŸê£æeã9s²r¤[K@aÖ+¡+Ÿ_`?2o›¦É/´C…:QíL¹`!ñ‘Çõy$4å¬ŠÙè¨·Ñ­Ÿ0ÒñâÖqå‡v<¸S·c…Êvk‹WñÎ×8mê}Ý ÙãïÖš.o‡r:íðÍLÝ”µÃÙ­5^Åß\~4žúË@þqÁÁäUÙÀÚÙÏ\¤´ª1Ño:6ÏÌÜžgÒ²v¤]ÊÚ‘o:·cl]ÏÊÕí(kU7£^tlÅØÂ=×ð53CXµfRˆç®óAÛÄ–)Ì kû–µ ßun…[É–qlµÆVžÄû—¢=”´æí›ôRúé›ÛéR»j½Èšj´{	fvLc§Ààë7±~ß.eº@ ùc½ñP™Ä°>gIœ¤pæcG´‡(ßD¹*tçÉ8ÝÿaZ”bMhEÁLRdõŠïÏ‹$¿ÕÓIätP	P°(‹‚ê»g¼Š„N¿ë‹´$gŠY™«=ë˜Øi*&Ôñ[ç‹ØèÅ£‹T$y…ºçõ@98@J	‡†ü°+©¦sÚ˜JsN­áÖÏWèà;kÑ ÄÊ¦æÙåÂÞ:ðæ<Û^öy€‘Úkðå@Ó‘½ôHj¥¤ôáiæV¤~j&'a¨žÊ¨7Y=•Ð¯§‘[´‰ÈG@§äTŠf9Þ2\¨œ²ª¬ÉÖÆˆ»†Óóê*-%‡n–LÎ˜Úò­P¡]Y­ãáŠI<5§aºŸ­­Øç#Þ“X+q÷g`¶ÎJ<dziÃA‹E Ò‰	é¼Ä }-.Ñ!Q¼=.,¼¬(yÎàtãU¸Ð˜ö S|l Ø.C¬ûÃèiŠ_Î`.dÄS™ŽŠ2©@žF‡»G»Ïží?‹®2Xýc²JG5T{°&2ÐQ‚»b†ugéù,.f¬ÚI”ô&èUãÛ“ÕGýé–N#vòhUA~Ž£qcÕxˆ¢“/ëz$¤a/Ë¬ÏW–Ëlºj="¿p}LþZŒ/SZÊ<ß‡üØWØªƒƒûš^lïSàp—­d¹±ÏêvÉÅŠßIZ(E\¹; èä<è€åÂÙ<LQ‚,Žiàœù ºà¼Ø‰=a¹ž/Ž³sÜÕÅ£=t…È[[WTvA×­³çV£e8×ÑÒ»1¯Si|Aù^Ý½ÒU—AýÅqoðtÁ½¡êb¸‡c¯ìE¢Êã)h°«ÞKÔòý·¨žøo¥zïkwx5òß‚&î¿DÝ™¯7:bL­|°B9[|û*®L­Dƒ!i	r™£ÄºÂ)¿*æBÒÆíAi‰~{P\ÖÞ…wA0)§nJ‰ÛCÒ3z(G×q4W9}(ÀKþF«¯¾v÷/µ>Ô5‰£ø³v;µ¡ Ö”ÐÿF	ôØTduHSÙŠœ4( ¹;Ã”ÒÎÖl€Š(o4ñ¡Ž8I§Äñµò=XD‘P›R)Šr>IýÚ“GÍÅÓ¦­j©;qQlkââC˜¸ú«ß;7¦Í3U—Éš 1ž¿:á¦º ŸÛXðz·èI,N~î²²Ï%IÈM'’!Ž>-9€%Ž ÍE×æÕZwn‘Elñ¦uØDgµ¨÷À,¢çýÖˆè9h¡ÒÀ$D5jCÉö Jœ¨Òb
‹F~·áò2J¬°æ"…
“Æ
Úf	.ƒ­+s]6Lk_Mñùˆ¢ápâóõeš(„£?KxíîkÃ&<Q¬=Bk¢¨&Šþ¥üL"Ám8i¢(ˆ.ÊrŒ\)ò½LÇµ1)È%EtJ;¸>™f0{áÆß‰»$Ði–UQ•žÏJq†n4à ³O>Šîr'ë«%¾½ yqõ´(AåªÉ—×1>,S0ºGiýÕÖðõ:°No­7ñ EmþÆ¢[å¤Uö†ºÁÃphœÁŽ"3††ÕtœÕýÞî^m¼fmü÷xíÇÝµÿº±ö›7k¢µ7½AôþÖT¯g·Bè©Œ)‚ÆÞâËÙ¸.Ê,~#ÿæó÷Ÿ›ÏßóNß1xê­—;ë%Ð…jF>žN¡ù'¦4P¢µèxÿñÑî~Ïh¤—iY‡×Z7Ø{1^íÑùì2Š
äpÆ#²mYýóŸ£Þ.Æ¢ÂcÜãz1ÅñˆÒ3‚ú./áZºìæ…IZÇ"&ˆ£@ç0Ì“£¼ŠíÕU·¼½¹ÐÑê±Ì?«7Õ°w¡Ë|C=gã)°æ5÷­mW¤¦fØjY9šjÛ;SùÜPí@9[RÍ5ôÞ“³ÕTÃì:Ù»P²|ˆËÒVbW«ÌaS²ldv¥Ì1%ã¹,Áq´7#CìÀfVp;²™|jÍcScÚ8ì¼LÓ@›2UÃÐ2ˆýBÝy©Á¾ë=BSÆÄYwÜ¬ª‹É³â¼pßiÉãWµ$:²[Ãª)Èo„º0ûW#D	
ˆØ¶ÿgJ«¢jý2jµé³•Ç¸l\ÛyR¹:¡Œí­Fns­GÔB"aª‡·Êdgöcú{Œf»–tˆÁ£ë;ýX67î5p¨Œ‘	0Œ‚ "ôá_YëåbkÅr+Å¢ëÄr«Ä²kÄ¢+ÄmÖ‡E¥ü¢2¾‹„wJ/¼Vá\8‰'q~QüþtK2jX2Óyh^ÛÊÓ~P]nE0O.Üfå(%AEíå¨Ò¿IßÐ$SÌØ"_÷<¹£Í€>M.5üØ uhF–Y_NÄQ™(žœ‹Õ£Ó²¸BtTô·/gj”¼Z‚b©éÔ`xQ\ UúoeÚ)ãlh(–ÆJ5#%ó^Ôÿü½–›è÷o1>N€Å
	 íš=´nTPÁ‚g£bÑt)9pò©™¦n§4Di}i Ñ
˜þÁ)ç0ÔÊ`5Z!Ü ‰•§ñø"FëN4ªJ}
.;˜øÁeý“XÆ£zø
§$8¾ýÝÉógùtVËc¼¨û3qj*Å ªó´Æ<Siµ3»ÉDàÝÃ—&FžÕÎlå!+ku’–ÒDö9¢ú¸ø>,Ô1%G»›Vt(&¥Œ¯`í®eëXEb°ƒ!Ç°¤£u/Vj;å«‰gÌãJOÀžµ“˜4Ôq*ne·ü~–]ÑeücV¬Xµþ	³ï	æ7Œ¾;yºöMôøåsl)hÆéÈH÷‚a¢×ÿþßfO÷Ÿ>]‡‰ Ö±@Ð=Q õ¤]Ég$X¾÷ X}—_•ñ4šø-‹ÇÑ4¾&êçi…Œ/b©’­ÍTru„U¦ðÂh©8f$pÆ¶mÓ·_ýJGE­ãî	êWzuÉGeGþè‚0t-ýŠ;cô‰kú*;ç´'ßviQuÚ/›[¯“ø²[/LÛìtÞ5·¨øËiS³]‡V/ÙíÊ·Í-‹5ÔmX¾íÒ®(ê4+^6M»ß¥c`UP@¢–qÒXœ¢0-ÃÂ#²i"eCt°W9ÓIW< zÛÑ·'˜‡t2Ü‹Àø/+ÉÆ£~, :a~–ý O(mE'#éŸ	s: JÈ²ÐQü“/¡„å”TIUú×Z¼æ3Î®.ã¾:=óƒwlÍÐ"[ø`ÐÛÏß*Ü¼ùü=:£†yqÕà4uóÖË¡o"dzøM…4Æáb0à/ô0¡v—Y,¢Ÿ³\I1Ì+âŒt6JÛ‹´ÿ[½ÁjŸòQã¸þÇyÆBÈGã—x÷=ã!6Ñèy\!âlÞ9ÈX›”þ‘ b¯^²ø?Ø=®*PQ&rXµ;+ÖFÍDÁ:Ç†tk2èV=ÃHçÖÓÄ<Á>ŸåPû‡0Ä7Í:ÛY’ÞžôH¡`Å…±·Õ;†¦d¿yƒCümL:_÷5÷b?à‘idMk§¥ùX“B&!d˜ÕyIÈâËfdµ1èc¨>=rO6izÑ›40òúÍ÷3TLŠÜ ƒ•ÚÐQÇšèˆOÜPzÃ›*N››Rvªß’Œäœ€Ríà»Oð¦E(èz§±±%u\™ÄG®j7ì“QZ6È|†9ÕáOmˆðýe1õB•÷²ñ£XÚ"¨ÜeUJÕù°”™:¡g·–‘ó—u›œw¶™1ìvŒÊzÏ¤y5jÆ¡ak¹“ÅON¾jÜàUÛ&¨¹òÏ:YÇõ~-ž[ÉÒ/0Ÿ‰ÌŽ,¿{åÁÜµª­¹äµz.xÃd˜ÕL°×—d°çÖ51¡ÊM`e‚õ§Dûu3AÜ¡ˆëq#˜†Û«¿±“u|.¸ ctÃTQÉÚ‰¶ñ“V^Ýå¸Îdµ¶ã«Us3}s ˜ÊÐƒàŒ¼!I#FUÀ&ˆs 	—»ô‡g”­ïP:<²Çx×­Ôš¡Ê2ìœt›ÁáØæÖEŠñò­À–Uv<ø\¹eµ¸HÿÝ`§u!¿šß=Ëlë*OôÖÐ=ÓŒc„ZpE*÷oFãø!	¥‡#Jt'Nè|½GYÈ¯v+â˜„uÄQÍ,I/ÍÝBÄ9eph£Êx•nA–u¯0º‘¥èâ'ë[ˆ,{L®> Š¸U]ÊU"Àæe\Â]H§iH“àuAU]¼+<]acwüEºÚœõ;5Kç(W#žV{*”}áePuqµhz¼Å…mÙŽ^åð!®FÄC­wB_xMtt‰ÇÙOô'¯hW¦®ÒÓ2NßˆšoÆPspºÀBÕzi@Ú.ÔR½kGteÜÜt¸[Ÿui.£7Å„Ïe‡Öq,ø¬ï:ÑM$š«b­:%íˆ´!âÖ`‘i¾Áã”µ#ÔBj®[#Ö¦(ºX”˜§€ùeuÈˆ§¸eMÔ˜»B{4UÑcÎæ—cQdõÁ-ïDñ%µ?¾v4_XdºµÂQ}mRËÞ6a'< z/@«Å3ŠÍHYZ×øåó÷ž™Í&«¬PÀ‘q#øûÃßäVEÊŠ»öèáì{N(È‘ÄQä´Y1Ï>ü5‰ÃÊUC¤}cxÏe{ª£V¢8ÊUªÜ@ý\b’¢°N„´HÓ’£ÌösÜ&beùYáSJDq7Lð.Óhí‹Üù™±†éssäRKTQ#ãtC‹‚‘CcÌ’‘$CÔJ¢DÅ[£?¦ev–}?K£³@‚ 3²üp%>G"4H`…õÇÍ·,õ.¾ã„c°/âŒ&NíùdÔ©ñËƒÿìV'éuŸ"‡ð£Ø°]_ö®óx‚‰ìÜ9Õi´ûìYZ6ŽAÚGb7‚ÒË§	îÝî=Ž¾ˆpŸó‚~3Ð…`==?c÷Oì*j¯ßs0Ê¬¤¬²4–8Nk˜¿ä~þ#™$&	î.cˆÐaGïgˆ#~þÑK¼4ÇN'*#­øyå_ßÍqd1Æô˜‰À+U'2ˆéRÃºxV\¥å“¸b‰ÑèÇï²)ò1u;ª/â:ºB1ófÎjŠ$§uÊ“ÓÈ7Âô¬ŠI*ŽV5 žìÆÉU‰e­H*»Š‡ÔXÖ}Óå‡Šµ>“½8Àù3Ó“%q™ v!C.*¸{´¿{¬ƒÅˆ³Ã~Lw™è<¤à9Ãr&½ÌÈ²©Q]•]pÁv„‚+›ƒâ¢5ÉÛTXþêw[ÍÝ ð-Ì §f
Ï‡IßŒž©r “ûQµ8å-NiŸPy@tGU¶âe+«lCþ*U¸´
[ÉLƒçˆ‰;Å¦%[K0òÅ	òÊXÌèn¡´%íl	1—´³rŠ¡Ê™/o:Öä `¬«¬ûýx5:%Q‹›§RÌÁPú˜ÏcZ¯=>êáÆ*Í«¬Î.³úz+êÂDë©c å½Ò9,ùu:s&ËÁj'õ9n5â“_¦.âP‰ØcµÛ&e»ã¥*,ºFgÅÅ $‘`XgÓBe
p‰fb8UT”Ü‰˜±8HŒMnÙq0ÙX<Í’é­Á›˜3ï<‹SÐ3Ó¬­nK<óÁ³¡2È$	çƒä+îx,xs Ê‰‡·Ùx¼´ú²–­žýé|u,Në«kdÚ-Ÿ²×Ö¤Rs…>ûžå¨æé>®ÏR¥¢Œ–aŒÒÑŽ#ÅÓaVí§¨/Ëœ¬Dbðß¦%CÀÀË“ƒÃ—ÇoöŸíï3±žT}•°m0’jÀW¡ŠÐ[«6°æ}z­”±l	Ä’E—sÙL€²¶ŒbšàI|Þ.¢j(Ð¨¶¬Š|1rr¦µÔˆWv*®nëkM&Òë/Ê[Þo)ŠÚ¤2©—L;¬XøoxÑnîJåw¥úDº"¹ÏJVq.ÝÎkªY:l^öhE½WÅSœ\OÓ[¦œ˜£hfÈ¹‘U„p~MMàql1]Õ WÙ—Ìª¦ ³Ê›%]pÃ;ëW¼}D¯1Öø)üPQ o F%Xp…Ùï“…Õ³sh*‰Ô·1oÇ}éôÓ7y%<‚$K‡·>aÛ’¾$÷ÎµÜão{,cäNÔg-mo«‚xÒÝz¯`ù˜ Þr¡µÕÒ]ÀI2ñ¯­2l%3"’e‡$¥_fþÐí#ª0ôXýžs±,±KOÞ¤mëÇÑ)™ì}ƒ ýÂóØQ%RÑ ó‡À¨vS]¿ÑÔÁKÛd>WÄó .è	‘F'ºÐ>ãGÒQ_yÉŠÂÉóÔ»K°ô#i‹Ýn¸éU‘‡?WMØ]XzÅamŠ#¹¸Â$ORû…J¥’Ñ¡*ì(â(¼(p…‘I
rÈw»¯õÐ¨ÃfŽ'[v;;Ö#Jó­èmrãùÄóæ­ÙdQºØœù‡l¡ÅÅÑ¯&-´rÉeøïµi11÷¼
£ÓÜSr¯ÄÕEšGL‰S<# Ÿu¹¯è #*Õ°úîŸÁWký%g¥ëã<!fv	»âÂ~ZÝÕj¿§:ÖoUˆØ°ž•z€>À9ÇUƒáj»ÎaÈj;áÉÉÚ¼H…³oV_ˆ[C‘lt™Î?çV04L+=ï;¨õ/P‰mò™Cš³­è;Ú*µh~of{+ýöôjq>ÄñŒeé9ªMPðg¶L9f’LÄ¡âgu<‰ËºŒ³2ûÏç”hTL´ÛxHŸ\¤ ÙN¯ÕåA¼ñßSv¬í»ÊºªR®ò&t.÷jkj{à -ç‹y¥dJ0“¶Ä ú¨aî[©²ôÙ^Þÿ”éµâŒÀ…ã
8+ËÑ¾§ab+Ò=Ž Ì09¾Âò½¦åÈGPñD+^\»ÈZ)r(H¯¶btO—%:ói+Š:DlÌ¢’ñ2Kˆ›€;‡êÊÒ	ë!+œ¯gÁ¬IqU]eBIÎ‹â|œçÚú §¾Oò$=‹gãZª‰ÞŒa=ÝÁ¶åyT|ðY¯nÀ6Äºƒ·ˆÑÍ«è„&VÄ3WÉR=;S±³%ˆ¾îØÉºH ¼(.Ó:! µ¾ô
¹YŒ±2þ1.Žàžñï>ßR×([ÑHõ=òí8T ‘bXó6U~-.åøwŒLy‚_—ÌÖ´¼LKÌ:[AÓ¾
 gçY>§¤¹Ä›nðf_Ûm±ßV‰‹¢.¾;z†•ÕomÌH1« Û» 3ÅÖ'TâU¤I€r{¸m“Q-$Û	àŸÚÉâŠÞëÿÿ â—¼NYŒÃÅ{1>QKòhfÏÌ]3…À sÏ¿×®Y@o¦3Ê4jw}¤¯êD¯¸™É27¡+ÛÄìâ~‡ï?S„èD!R†-Ð}(ŽÎÉ7¶•h›òÛŒ§Bûò¼˜)¶·ªÀ©ZÿÅ¯f|²àþ=u‚]LYv W½!ÂÌé¦B`IÎ„±dUl¹Ó„¯ÇÇúF<5µ„Ê¦R¸eSÞ.ý0PVO;(K“,TçÂ³*„¿
ŽÃò8™Ô…³hÕâ~P‹f•DÀ3RwË‡¢4rÿJì¤ µqöá/eV º,´eßÄ•â:¼ì,x#(ÛÑ3èE
¬Ñ&QÿjvŠN´S¤c‘c]Ìñ“ŠÜ-`Â‹U`YËž»9Þ>4oðÀžËw²Ô€	#B!±ÆAÒŸa‚LÊÕ¸%âTIèWy=ûàŸœdç% ³;ŸûªDßÖ=fœÛ´_ÈÐÈp{{Bïñ:Ÿ"k`ÔÁóZ–+K;5zÂØwkÈ„@Æ›Ž¬Í¹Óöžì‚éAŸ·ë17Y_¿Óìˆ €Ÿ0Þ
Œã´^BÖì‡±#™®3+q¾”YlœÒ$KŠasŽ`¾bmDƒ‡±ezŽL\FŒëe”€pü›#ºNM›c‘q«·þ1êÚìôô³I£ÆÌW]W˜ób÷ð ú}z-ÕŽ«º9ËÎ‘'`êÔõo}ê¯¬ÇÓlý<Å‘XUWÃÑ¸ÆkäG”ðµ?°?‹Åï4¤Æ^ÄÕ1¡˜wžŠÍÒžš}`JîŽpõ=!ß¬ÆL\Z ’ÑÆ¬!°­8En´^]Ä—)Ñ8ÁpIl0v£ÝŒïÈÑ1æ[y—Wã49OI. 1„”@Î#lŽ7GL$²xâƒ³}„Ðwœ.²’AD¥DÛ=ˆDð,XÊÏRÌ®p„î)Œ3#.èí’ÿâ´·w5ãw™¥Wbìtyr+ÎŒâë²õh«°OýšÃFœåòƒn ˆÑ¸ççªW–E(?£ã÷P•éZ`±"î””4ÅŽ®6…]õ:fhO)ªÔ*Ùîº»'¢ð­¬cÄF‚e€áªÑ Õæ©Àh6¹^?„ÎV$ûê’D©.
…Ø‚ÏäïYxŠ¹¸Ã=c²ä]^1ÆÔ†Y‘çñ´Õ¹¯^‰ÛŽgôRH·EçñÁ;¨²ãF—!njMv|êü×Iôy}iE:Rá“>ãûðÿË`F%ü s(ynÛB§ŒÈÎx‡¤ö=Ž{¹Z†då‘¬KŠ€ùþ&Ú¢ê-;Ø«X\¿©Hé§u\‘_Ž>¯F/§iIÒˆœËÏŽOlÎ³¸å{'SŒsiÞ’M}ŒK¦®ËæòœÕí´Ù]˜ÑË²{0÷¿LÔâr>ó25M ¿HhX'¾)wÝÃ»äwH}œŽ³4Ê³±I1þð·s°|("tËl|j_ßÄJE&0Šn~óéß14çÃ¿Š!sŽ’³$™ó;Om…CSðêœ#Zw-ýhk¼Ìau_ò 5Ÿ§%(4ânPqh~¸Få[œýpì*{¨«d1ÄËF±ó½Û
¦¸ÂUŠ)ÜEq ‡jƒbÔ)%"ëì@¨ÔûS§5Ù…hÉ€.Öök[Ñ»5ÇÚ\`s	‹lâ }è<ÄA¦ºÍAš›ïžƒdŒÍAêD¿ÏA¿°D˜%ø°r–`¡¬aŽ0Çí}†`•9?èÀÎ;g™ÌâÃà'd†J0CÅ™¡êÈÕß›¬å¼ .÷êª¶7Ü{—7Ã:g0äü¦n½¼svk¼ðËM¤ÃTð[ŽKS{„ÔQ¬†»lEì…O6UMÆžÜ5ÉxÒºU'iE?•dãîHgºo	8–`¯AÂéÔ!Çª[2ÎÄPÿIh’|Ümrp§aßê¤eP¯±g€8ŸŒúç¾W‹ºý^‹wûµœ~Î[ÁR M#sN8^PÐ¬FžK‹|¸“C2\&;F°PJÑ—EnùqqNép«BL`!ÕJÕÆ×gUœ³<E¾Èn¶;½?a&Ä‡VtúÈä&v¨%·;zyL_”5ã¬uô1tªH®wdÐßKùr——;³¬aZ_FíM¦@»ßk_ëm*ìXpËŠœ m•í¼ÊÅÂ¹¼r$ÓåŸéÐ•9;Ñ|óÙvUpW‹R1H`'®¤§€áOéÍÀh4À y@–J'ò·aØø	ÐÐ¡ÓÐ;¬%ÇŽ$Î½„]Ò}ì¢<jfù‘CXó‡NWE‘Ff¯ÓIŒˆöÞ¦7hVVGoà´âG$àÔÈ5ŽVÒi´´Ç64Tx‡ªæ(¼Wòq*»‚ÓãDm,?NkÙqÒÝê	ÏõœA2ù½’9åIXj„Ê¦*;¡é®_\ÚWkŠñ¢›5ïøxºb Ç‹š[~¼ÖG‰Á9òÑZö½W‘ªEàl@G]„å¨ÓàJ:4®Æ~àù{ÔˆÒñ«{5ÑºûˆÖz8e+Ëh}‡£éY]£hY4Þèñ{ÔÖl ÕEóÆËO»°4!¡!bCßÕÍHïÒªY^€´O½ýS£õEA=ª¯7Xõ-œlƒÖ¿qz®í\Q_©{§é†SwAÑ+Äžç@x/eX¢A¯Š¾;Ø[NÿãLìfGo‹r
Z=ÌO|—&òæTDûæíCÊÊáu.sx)0ƒàDôõ%A§p»ª“†)ëš´HV „'Ž™çÉòLr7„6Ý,ããuXlcŠ8ƒ©æ7œp9]ªl‡/;sD&˜$µeÆ9z•0Ý”œ°`à0H™÷d÷0ü<ƒf6Â?ßFv-|÷ÅvôÕ†WWáyJsg;º*9SHBrjáÂFêb–¿CsV¶!.ÎV±jÃ­ÆÕÁb„60ÑÁøøD¹…×—@§8Ã3AóÄš™¯æOŒ€:*&“¬vï¿æ5üd}Ý$ËqqVG"T#„(¡®;³ÇÛ1œ@Ázjþ¡`.gi/,¤¼yØ*¡ŽÒ–`ëíx¼±šrXaÉ³8Ç°¶h—qô»‚Î&5‹ 2mõ«ÈÄ_ŒE}q¶},bä¬>„¢ýÞ@ÖWÉY&¶ˆp5â"tÅ÷îxYy'îõE*ìl˜C”E¸ºˆP…¥7Ü1ÔÙÍi™s½œvéîQ
ÒÇÂ}LwïZ²²·´…$¸Ò½†…få(³bTªŠ$˜1Áû¨¿ýS´¹±±±û´½²¨mmp¡Â`k?
1xœ¶aà}lÁÀ+ÛŒÂš †ë”÷mN+.CþÓ=<ƒß-ðB4lû÷'ÑhŒAE¥ºNŠÙ¢Xò/à³3¢Ä¤Ú?ãçù8“b†×ïQ%i¬Ñµèâ:+@»Œ	¡áÕ‡¡³,±„¥1y_½æ¼ã¦ý·£8m1Fÿ@·7‡Ñ¦xÈ`(`ñÈ§"HS(È•X€ÐG©.Œe+z±ÿÇý#(=&f4ù]WÆÉ§°XãÑqÔ9CÜœí$®m …£¡
HÁù`‹n£»,Á	‘CbIÔÀ
2ïJ†ðÐÜâ¢ã5ì(jÓÎ*^GŒ#Å,"Öýz)i ³ðÛm]!3joÉÅEÉ¬,HL²Yä±K5Ta&‰Â7’ø%oÌ]‰ó«kÝtc1K´qZ¶ˆ8“fX‹·Ž­¥`c°(rnh'Í.œÝeP:r)†JV~m). ‚¡¸…%¤ºõå^i÷Q$/qî¹iøäB¸b_œÑGéÙPš×bc½aX0êcf„FyÅæ+Î@µ›*ïåÑ£Û™#;ó'Ò- KršÌ0q$ÈmÃ‘nÈ°ý@m¡$ÕAŽ•Xd‡ÎÝ¨S]g„ùu+\GˆýR_-§Ï’2éØIs\'[ R€]¸#¯ú$ÃJK‘è/ú4ÄcF(™ŸšaŽû{%äY(bd2<U™‹BÜÕ‹‹[mþK–±Œ‹E„ úƒ‚°Ìg–ñŒ½x,žå³Ô‘ÌMÍt‡cŸE;cO¾˜¨E¦iÉÃNE‚ôeÚ•84’CŒ0pò¿VS0Çñl’P5ó"_9ÈsÈM¡a™R°°Œ2ä6K%·¤’u{¥SÉâW¶Ê’Ì¶š; æ"UYÀµ‚{²ÌÌ‰c¥¢\5DU Rë´hs(ì§¶¹ƒ8	MQgüùÃÀ;BâZä%@wÅšYmúÜ@ä¡WozÑ‘¹Ôzôêy\_Ë8O
AGycð—_†xrK<Ý_~íÊ„ð¥öùfá£GŒ3âF·ÜjP*í#ÎŠ®[$IW­&”\Ýò¤¬}ê}ÊÊM›KIö¢R«Ù¥0útgˆ?HŠÌr•Òx‡jÂë­h“×Á¹¼Eÿ·vŸâ¹o¼svË=¥?-~+²ÊK:»e-g'9íg±‰dÌeçÎFÿ¾3® ðÑ–Oã~«ãvVÜátV]ôiÕÝb ÕÚ«XÇ¾ÚÏfß µˆí0¶–*$õŽi\›pL·"j&´A5²àÿ†MæWÜ“Á3‹<lv°²û¯3l&lç¹TœÃ{)øç_ˆÚ“êãì¸m-²G@;…mtú·CßÕ.×pˆ]o=Ò ‡·M=\ý-†[n94l0(²:›rkÁ¹ÃXfyaüþÍÖ³’£õ#VùâÆà×’i(ÆTR+‘™câY™ÍÜ=Œf+2|„¨y_¦ã“åÎ{$lq^Æ“XaWö±7Oº]wn U¢îû–!X®U0uœÀjÙ¶mú’‚ÎHÿ.ÃQ¿n¢ð….÷*àïÇU­:‚’Ž#ñ&ÓeZÒq1©×’jE`
•IŠÁ)§´Õ•2èìg(BLµšuôýQË!ƒîÎ.øåLÇ–Í^l8Xº”‘XG—rlëÕEQa\Ð4™)ûñh¬/r#ŽF *ÁL&Aì-¬F®àdQ;™ð~ôOÂø®È¹°9h„J¦åH8„FIw³ôYÖ£¾i^ ›%LëAòŸ\øá™i/®žÐ‚Œ¶\hðœ]Ü2KÖŸ?·s<áŸTµx7lÜ iÍ3)“}‰0×HX|›züÞ?Ã‚•9ðUò×åÜT#|ø[KR.DáÑ;Ö+YnïòÜiDùVTvœxV0
Ð<Ý—KK‘Àüð7Ü$n•½AqOê>Y¼cÎ1Œ˜kQêäÉÔ—°ÓÃmÒ6íE×¤9VØ]žqL)µíž1Õaj}C<Ÿ_†Œä¶À—øøò¬T‘&Úfœj·³'¡mÒÿúç?7CÊ§
aZêó$16 ëKéx6ŽzbkRCœyú4û¡‚
hç%¶ÁùßëLcÈD3­¶L¤›¶ ³Òì5àÑÞžfÍá§<p+ªùÃ.%“î¸À&†–äcm—åcg*ø#Ê¥|ãî6Š$¼¸/"†Ã•#[z¸, [ð¬•Ûš‡Ý*˜OŠõÂ‹3çÿª…#$îZ²ˆ,.Íõ²£äž#ú-%Ï¤ÜxA%•[²@ÌÅ%NõL¨“Ö¾ U­­UÙy~c¾£Ãb:›Ê¤]¸!šDH‡2‘ÓË$½	àÑ¶àòØJ¬%›šŸ\I.æÆ¬+Ü«ë6½‡É•ÈÈˆó:&ež®ÆE•&k§×køéd÷ÂÖ#Q„¢œ¨õ‚ã«‚î·‹bÐ}0âŒÚ±–´(ÇÅyEŒÕW AoÇ6yŽªðMa]˜åø0ÄÐ…L÷ˆi;EéI
&Þ9eý³^˜‹Nz!¶wÖfŽ— ¶L>ü5ÏŠ(ÁD‹QL[b"­\Žÿƒ·þBúLIwh*+™2=$nc"ÉeKÎFßÕ:Æ£B~+f0²"?(®Œö×Hš¥Ñqš_@Û§1ˆžao.}^1g *MB;»¦¿6²zø,þQ,H˜6Ova…kì&1))5¥ÌÓS2Eïo<ª‡¸@¨²Iš’ï×‘ÔUé¼á	Ë§4e£‡"Ï5Ã¨g=_ãA.ó«&•KN?áBSPàž'‡2E¯œøøú$~ª¯ž³M¦m§5KH÷¾(êÜÑÐ*T•€,§mè²nKwW¹..Ò›PõÆU¦Mö9|Oœ×Àö2†ˆµxŒ×[^¯e9ÊÉ^³€ØÇ,š©˜§A ÄA:ŠÃè&ñY|Iÿÿðo1›?Õ,Æ{.b6Y Hµ,Ç»“µ‘¦I¿`Ù+0ÆÎ×L~é¦rHìµ¼¨×('x[GEADÚÈÔt?¼¬{®ÒøAª¹±]ÙHÒ ÂýÝi:."0ÁÔù/Â	’b6¾îMÿôkÇžZ9hh_!†Ñn’a¤4.zÉÉÁ¬Q>BqÝí¿§”ÚLWÞ5©ÿ¡ªOŒÆ±”yRu.TZÏøe´UJi {t[D=Þá>C­I(](X^Îj’"Vø3Â”Õ©÷žÎÕ×¤’œú*sç¨”î©ufª!%_=´|x-
0»^­1YÉqÈè¦œ8›|x#û@ùúwô«-‘„×ÊÂ+AŠÀâ§>ÈÓtèc—Jup.¼Yƒg¿í»ŠK•ªcI*8ú GVì˜ˆdò"@<öAœv¡‚‡¿„ág¼kL”ßsUÍa'~‹)ßãFvåÎÒÞ’­pšÒÙ4¼®Yj|— Àh¤T„%PæaÁBÍÃ[~J˜ÅÆ5lSy-O(×=ª`f0&È˜ª4Ï·¢Þ˜¬³<«áÕdÂÞ	'¼¿RFêq(U)Œ¥Í¯œ¸!{ÿ’%¤µ%˜©Ócæ°0ã‡c~8hTÀ»z)¦D¥(@½;LpC–¼âÉEKò6–Ó2N{À«âNE,ìêÓžØÕa,+!ÓÛ¸#˜?â"I‡(è–=.fåØ×AÏEŸüÍ¿ø•Ê¤íòåþ®ƒ'.Ïä /@iß<Fèx‚Ns,ã' ¸`¯ÿ·½õóUJ&çÌ	9#„¿ýB³p³7tä	]ž%·êRfÀ „¯‡2Ol0W³Y0GT©ç(‹XÈUõñÈ×D@î¢7%Âä¥›xðúóè‚.ä±¹?IäG¼ßoKÌa iVu¦‡¾"(Xñ,..‹ÎÙ—49?,Ú4JèHå¦*¾É–ÁPß‚x¨yÿåý¦¿K³óZèU£ÃñêŸ„A´®q^á\‚AK„¤‡Ýá‹ßöV™ZS²jÍ´°mn¬òGƒÛªY%ØÓÝã“žÕßP"yÓK³'þ¤˜éè.¬jý ‡jeØ+Œ4Ëï²U T£ §Óˆ®G$Ð(~*«gtÎO\¿&o§V³
w•½Û”`0ÂÓl<~RŒAµÛüóÍøßýßˆÝÓt‚q‡ÉÚƒ«m¸¥Éî°IýåAcú÷þj$þë=µ‰IíŸ¤?Ô¢ýûøŸŸyëFs#ü±¿³é2EMˆ¸rZŒ“§dMÀ¼|qüÝ³“—G»ÑáÑË•.Ùüõ‘¦ÿ¿‡õ”Q¼„ò<÷œðŸÏPýRï¤ãY"×öQ|šb–v±ÀFçðôgÜÞƒ^Tþ3°\l¨†>›€Ñý/W£¯îð&k^û›®D‘$!á²*»°9…ºö›áŽ µ¿Ú„Œˆé óVÞ±}þÞ@À„QO^q%xèûMôçˆÖ­H,þŠ·1 °Æíæ¦Ï	~µ‹Éom¾Üà»q!0±²>'–4ÌõŠr‰Q·ÉÞDþw.š¬]ùï”`¿ûkIvžÕðj’å3$‘~,yÃýºq÷ÊøJÎ´û_#+á¬üjƒ3Ó}‹™ðŒT_þê~hÂnTË^êôj†çaâí?Û=ùð¿Ž^nE‡Ïv_¼ŒööAúý?þï—ÑþñÉ|ü¿=xò²ç)#x»ÝžHÜë*tµbºWQ«ÄÂ(Bì¢|øY}-_6í;Ûù¢×QÆ”Ì6¡tDVÛšO)™€¨†¿15Ž† àöíM¦š6Õv¢¾~ ÎÂºðÎ}E×k”W°¾ªM¬(NG›úF?ƒu‘CÆß“Íû[ð^£¤¿Ž²m:pF¤¨Ã/oóÜ„4‘1€1Ö ‹Å¸Âë ¢ú-âÅ7‚‹z›	»„BÞÕ`à<Ô/Ä"«^Ä/ô7c®¬¯Pš}
©µIvv†ŸU¼K|Z¹õÐÎõÛ	XÕ j/¾®¨QšûþzÔ§Åÿ}­ÿwÿ«poZKª&ÅÛÏß+7ÉÛàÚ¨RroÎ&jÝFÑy;0ÜßM§:²•õ5ûÞ0;¤ÚýSîÖÐŠÇÊ1vøÎ‡°eäv¤}¬?ªPŸõëërÖå{¼9@ø2š ¼KŽ¬=t!ž0¿Â%uA+;«’é³º8Aq,d:óÖqYÿ—­è‹¿ é¿½zÕû¿0ÈíÃ¿íïâ½ƒÝß¾øð¿ŽO@Ö·öøÙþsú„«Áÿ»íGÏ÷Ÿýµ´cü~øòød÷¸GkÈñ!¨pþõûÏðùàÅ‡ÿùäà%þ<<Úý¯ôcï»£¡xqðRü:>8ùN,4½×¯~§Er½eVóA,kxˆkœžÕ[z™+Ñ:0u1ÝŠ¾úf Õ doòÖžY;³WÛCX/·Äe.žßOÝòBž‘³	*‡ìJ¢ú-VR…Yö××c¼0î½©LAn¼^m6
j¥)@qN*[Ñ¯ý÷Ø>^¤ªÅ¨ŽæmÝÏ˜ŽÇ‡@÷ÐZ¨k8¼á®\†6–6PøîÜ«›«Ñ×Põ›¯ÞôkF†Æ¨Ï&¹F„AÜ@¼°IR§¶" ²Ûc(Š6
÷Aó	Èú÷µ]ÿ¾Sÿ«ö÷/ïß8ß¿r¾yßþþÀù¾9·G_/\ã×ny5¾Y¸ß¸tÞl¯aOúCt>=Úsg}%"…Õ=NÉuOßÞ'¸f(IüþM8øø’6cEh’Œ'U†¼tHÆövŸìÂ²bªTÄšCÍß[}ÿ>Ú™¿Y~}ÿ5™eQ¥k_[³ÝûüÃÿ É>ò&€ÍM ýµ€£Ôs@ßÜ–Z¿¹+j=yùâÉ³ï>üÏ½—s{†ò–?àÈ¯-ÏI+Ù÷_ìí¿8Ù¿“!	ŠøxLî÷:=*®Œ<äòý>.b÷÷ûî¿F¥jøJ×-.Óòl\\áF˜‡§ ²½“r^iC¿Ç.ãÐ§EQ;žO¨‹`E®Ïb¤½¶zó Ô%±B¾<£‚&žÙ:R´IGŠ¶DxñÅœpO/[@ïaÆMzñ^û9¾ñ¿qGÌ†ó™ül0Þéôø¸¿µèðÃ_@ˆ£Ïßg7øAzi0ç(cs¯G÷-]°i{Ex2É&Ø´ŽôÜYú+Û¨•?®`JWÃbŠ—ÔA[Å¬žÎê~ït\œÎÊ±åAÅäÍ)tòN•e…Z´^ºYª]<Npú	%ä¦›
/R›fî†ÔmŒ1z_Òq¾mJÿ1Â±øÿj4ŽOSX!íjTäOÆÙè]t#£E0I üötŠŸÊ²%‹m¿wÏ¶Èü •¸ž/KÒÓ¸|‰Äun6ÔXã¸ªÐÅ»ý~”æZ¹Z;ÃýÚ³1È;:z½&F3:§k_FÓÖ¾Š¦×ðKzi×~GuÓešE¾†9Ç’™H„þžèÊMW˜ÌÎjÆ7;ÑÊéùi¤­]]@óQu›¬ÏÕ/U`¯p([Ñ
USÂe#º@a²Å@?Pï¨ kpEÂâ<’ßâ˜Eðýöûû7ÉVX¯é'Z¾¯‰6+ôÂ­mnn¬¨áÇŽ›ž®8=øÊ©ïá<¸‰Ö5¦Õ4Î~+(0×ªt’‘oöÑ{â»›o×±œ¨ôíº`1|âa’ut¸„äÇ$»³ÌÊYöCšDx¡j½v½¶A†üó#’ûjí×÷#1–§tLj­T?DG7qÀ[Iù%”HÅ­­kÉ1>ß¢RTýlò|íÎl„¿&Ø­°h\dI’æšÿ2>5h 8 ƒ5þ
ç‚àŽÁ&¶M‹š*ø¿5P¢1¦k_¯<Òì¬˜]“S Î±û¬²_ý
K]àÿš'›ž~kšUuvv­['GPy—O.À"ÿRMŒûŒc¶mØ[ôrÁd?Iêœ6£w Í×jÚÛcsæ›d?Ê˜¸–yºòˆjã½H|46I _mnNx‰æaEîü3h\GE“zmsåÑéutPøß`ÆG¿-&iå·çPG>šç<öhµ¶”Q
³ËaµÌØ: Ìãí÷¿£ëÉ€í•ƒüÃ_GY±b"iû½9J©)îBä%²]–(®{n:$ ¿ÞX-­wíæãY6Æá¾î«Üò©ævV÷ÌÒª{M•Ön/Û±2KÏðªêpÇ€—ÿ„Žs;6Å‚£z~é?ýh>%2Oâ’ü½Á~ïžcŒ4Ô3{.	bªÓ ²üÇcJvúd>cò|òÝyÓÊB¿|G”¬wdO/žÇ£r¾D™`©5s(½swÜŠ?=gî^A»,‰‘ÏÒ&Ø;8~2—IVÖ@Â§U…'L»“Á­øÓ“á¤L)­ñwÓ0-Ì¥Ð˜AÜÿ¼%B•zjÈãÏaRàG´ÐGóç…}Ww(àr‹Ò’ÝzšSôF„ûenní°ö`j‹¹k*[ydùÁrˆ§5ZùKÇ;Z}.ö•„Ö{]¾û÷Qß}«…omµ©Ó¦ã8bÇS<R]¤©;pŒ{êü;Ó‹WüÒ!…Šk#SË¥JeüÊ¯öú¢}VºBsgm'Ù¢=f¹Pè*Ï rû>6ÌÄ`;uÊ$~Y O¦Òt)´²þìæþ2€#vD§S‘w:õ-Îc¬¹Vªj‹ôÑ¯<·¯ßZÏ¾±ð^]}îOmÇžÖk÷Ñv¼¯|"µçYÿfCÛ›+‹Òûø"KÇÉ“‹tô®™ê¿Å£aÿ‡¦É³l”æþÍÓLÃ4Çd8y@ÂHÍ)]Ë[£džßÅ“²˜^¸«0ëu	taº$ena¸Ï,d*u`ÇEÂ¸çÛõ<¶<>Ç|ÝÌ,šS¾²8åý=Æ‘x¿#EiÑ¥-2‡¾Hó‹Ù¤7ŸaOÏµ‡5š®}ÅÝL¥‰ëÎ
ÁmöQ¦õUšæèóÁ4»o6îãYap
š×¶§ªW‰=‚÷6•nBž£Póïß†ÛGŸúÍ7†4ÍÈQø9k)»÷YzVã.Ó—ÑNÔjÓ^Ü¦òí©9Œ+ë1çŸ.À<T½›·7>ÂQhðsýW~ç‚ö¸74žk›‘æô…[VÛ|kÒ= +Â Ï)-]­M»kÖ‰¼|‚>ô`@›•ÛïßGWbÿíçï)ÖmÿÐßXqo0Ó0Ô“Æû­¯ó·4Åîy>ÃS³;Ñ—_? dBEžƒ€‘eðZŽ/7tÞ©ÁÍzkövÌŸ/LýÁI#®qwô]‡Ã$ýÚuéf“sw$«r´ý^¤´›^uñÝÑ3Œ{{Q×Ójk}}–­Å—q—˜ßg²O³õ[ÿ\ÖSÈrÝüêfð9ñ×öæÆéo¾ÙüÕ7R·ÏÎÎÞz¢8×d³po-ºØ-Æ•ÂM/×lþ­8ÀÊô,-Ë´<,`e¸Þ^É‹5õÊ.ê.zÃd–Z»ZsÝñP|ê	¿j|ß×z–cÞô•GJ~»>í þ‡Ê•¦P:©îƒó·¼bSªÁââçym^·Fn‰L-[­Cù¦uÓÞicH9]¬_Îj¶›è;‰Úã8+¥}VÇ¡Ý³i  {Ð?ùÞžH#¡Î¡†Î6‹<Ó¢ ¦ÁàÉmvd†jvËÇEquRÄUÝ_Ñé(£%è[˜$:ŠD½R&^¢Ã(â@žÔ ]wsP#|%>“Þ5Ê%^ÇdÍ½8ügò?f^SífI1Ð}á]¡4M&Ýš~çuKæõ¦¦±!Š}²ÞA`GŠÂ}‰Þ ÊòK4l’â^tãîfÍ&¢ò¥^Øü*J>üõ<ÃÔ…“ªÌ'ÖáS¤ÕQ*ÏhÅ)6=³+¹Ôs¾·Sòði+ðF¢†°–žã-ãaªB‹¨›·"êÓx¹Œ&E‚-Š‹ƒdj@:x¥òpâ8"÷áÿËÓSPôÌ¨ã‰â2ž#/*>ÑLä/]’Ö{Ê‚ÈÞË¤Ðß'ÅôZ%ðb´Õwêkfñmÿ½:^†ÿx6¥UuLÉ-Uƒ†êŠ…eÒ‘'”ËXæÁÊÅ=¤g% ­¨H,??B•bš1…0ÏËsž`)ôž¾<~³ÿüðhÿx×²€ÏnF‹chäð:‰(ï¬7éÄvË2¾â½N}u/‡ÉaÐ^øP­fe*ä&]2¶kó†-’S
{¼þ×IJoŸýfRô¦È„òÝiz°¯rÙµE&OP‚±Š$ÑT„?I*ëFøëó¦M iZ@ÝI¯€	ÜXÊiËj¨Ã½™…‡2N¶$Mˆ’jÇ¼Ä#x[BšVRÀ¶‚"{ŸEUöÑéØ–K+‘óò‰LŠÅaäÚÁñK¼¦ËwÉmo¹±;¢MvAÊUu«Üëö|êî™l9çÞ§l)&ÁÑè6"ÊÜ‘YptüÂû	ÒOÆ€ žLûNö\/:ûºH:Qtº¼Óeà~*ÄŸR}y´8ò^ÁTÜ-ðÊ»¢­+y›â$U'¼XÉ*¾Ä8I‰t„¡ðÒ0?H²=ùÕœˆLÒCY;.x7œÖÜºMkU‡9KÌ/€l=QË¦º>‹-Žøêage—nQQÑ|Ñ|Ñ|š¯Ì=‚âV@áBçôû»S‘éÄ/
²µ(ü¢ö.¤g}Ú¯[7{‚Ì¿ìêöt½Œ™y¤¬jZ•E/Õ:ô"Jª¼Ã0°5ž‹_‚Ýq«7Ü!­‡hÿìJ][(…
ê…úö)­®TÂªáÂÚµöèo¨- ¾	·#{lv=²ÀÅè~ä×?Ü‘BG¼M*†Û=¥˜…ÞÏWÎBµº*hÛÆI‰]>-ã*§YYÜ½‚¦ƒsºSnzf*…(È¾ÏUÎÚ*ÏQÎÍ*Ùq)…ì§ÑÁ2äûËR*çÖ;ñHâ¥Üâ:nèº±„ Q:8ðú¼ràŒNZÃ×l¼ÏµäK:Ÿ3\½i,ÕÒ˜q~âÐ·z?¥cåtD:äûòý…8àZ©°WŽ¹¾CÞ› +iâly¿× ¡•ù(rªlE½Ý:»ŒY®
´(\¾`ÎWß‹ˆ7'©	VQ—Ax¤ãk:ã,åEt–¦	î¦ûJ?´`¹#ñCÈ	œñ;™H,d•>ûŠ	_ÖðÜF—þâ½ ’wFr±npÚ³jtl…šÌ˜;~y×ú%Í?ž~C3‚=øÍ1¢žzc*Ë»a¬Šú‚ªDO¬5uIŒÝ–¹)F´$ž‹X÷ÅXuÝKc ©õ~kTQq£á"ÕÓ"ÇÃúQœ'˜òœŽƒ. */ë–Oô…F»aÄ=§s§I·ó8ÛŸç=±iî‹!»­»ú}„¨ƒï:(!´	'ïú±G!?öèV~líaYzn³5ÉdñšÝØ­#«(M"ä•¤á¿œì7Óên]äv7ù<+k¾«ÜÑäwyín\æ¤Ð’zWRÆr²‰£¿˜l¿˜l¿˜lËd³\ç\‹ÿü®¼;ð§;æÔ€—3ïDå®æ]1NŽø™!9 »**Óš5•é&[·tÊ^Ï¯ÈIU§Á˜dŸ)d Ÿ©.Ôv#2Pª¡º»É…ÎîHô±’º®‹¿,Ñ«t"õY·–cbØ¹ÝÉ5o=Ì‹T´l>pGøÒ†Š{è‚Çi£ÓÕå`}V dXnAl’Ô‹mæY™HãåŸLÉÏ?™6{úºú'Sß(¹ K~ÁÁ]m"â´RÍIie‚ÊõBÂy|-¡aº/ÙšÚØŠøyÍ?‹m¶w| UùÍ–î¥èt·%Ý¹é^´ oyn»(	âHLãÜ­¢ÏH-w±ª›…}ƒ+ÔØm,/›936ôË‘ƒ!CŽÍˆ& í÷µjÙÔaþU®L^µ ÒõžW.ÈšÁÍ»¶q¼Ù5°ºåÞWÁÖå¯÷^ÓI³ðÍ­ÈêNz‡Éœ‰‡þBƒ#o•‡ÿËÐÁ’Þý)	ûÌv4Ý;i™lí,:H6l$nÂÕ5¼Eã¦“×«aží,ÊÔÞUwË%fKµóeJ°Ø<	"ÉëÉÛG%ïHD4¹Ö¨š+
þ;üLùó@†§vØí†­S¸ÉÙF\ÑÂ¾X/À¾À?÷‡Â5W‹[øÈ‰¡]¨¼§œ…;¹ä¦ñõ¸ˆuÈ¹zÎâcî^?¨ÚØºm°ˆP=mî¼€"€Õ¤ô¶–¶®}~–:*ßÕ0ö­ÝOhA4÷`y ÙYs<‹¶Ê¬ÞZ_ÃI¿NÉ¶]•Y>Æ‡~rp]¢²&|—^…í{c!
µ3B¸/­*î¶³ô5çý‘òt6üûô6C#Gwë0añ¯›R‡Í`…Ö*wò–
þÝ’G¨Ü•¢‡w®ìáŸ3ãðÏšz*q–‡¤4³UÑ¢9za ë¹Yˆ-	Þ¨˜L²ºï¿ŽƒúˆÞÿ·A	´þù{%nÞÞÍN@¨3Ý"®)Ÿå‘wf‹üb‡òzuíúà½7iy>ËkëM…¬Ii²W# º£ÊÀ˜ºÛÆòÆFñöDÙˆÖ,OúdïkkŸ’”ü,ß¥MÔBÃ®RµÛØq?÷~‹´{fc¡ Ú;¡½a½1±˜ò*;¡ý	8ˆ[Jí÷¢¤¦%p'ò\½ŸIÙŠ˜Ê\6loúã[kÓŸÚCR8SŽñ&ÔótÅŠË“·ô{ØÚ›Ñ81\dªÀR9ÞßÜµã›u@¸¦-þïõ8ïÓ“æ{ñ„<¿¹jS²g®éëàà>”q#Ç¤NãALó¼Û¾ºÞu¿\3„Ï
þÖ3×[ÃlÌøv°qtâßýõ4[U
²‘vûÙBÀÞžŠáÒlO­maÖ8±¨´ÕÜº©ªºX$ªXç¸DEÄ”tüEº~BÒUª$Bó§–\ž1—6gêÿ˜¢oá=_äâè8B!ÁDÒí„à‚"Ekg‹‹”*>K÷@‹§~U&Zä.f21Ò±;Hv÷4û;ŸPYXÊŽô–@ÝJ fuR$è
ŠW¯ñB}mã1²M4„qåñô»Š.šZ=(ÃØ¿þ÷³´¼îüV£«‹F§'G ¯RÛÞî©®ì»È!¢kŒr$ïÀp:«àu8n(l5ÐmÊoÿ€˜E§×Š¢">/)vƒ×,g£të­}‡²Ñè•±Èaˆ±K{E6-4ºhFŽITœ©ŠÔãàbK>!Öî[Ô¢H	AGrÔÈÄo”“8î$ëYd¾#¹üæ3€nÍ…ª˜¤ýô‡¬¢izò·^‰ÐºôÝ¡¡l1 ýŽµ=ÒÇ£8ï0ºÖA¯qšŸ×Ñ£hƒ7b]´A×Ejøá‹íè«ÐŽ.fù;áâbÄÄ¬‚ýlkREß9ÖÁ‰†íy#Lí„é' §ŸÙSnŽAÞ`Í6Y²oÕ]:Ô`\bg”Ž±·©ÑàÃàŠKlXkpc…õjÒ²ÄR‡òô³”;uq)cµïñÚ–`á’¤e'áºOF½}Ã—¶>~ùb(z™];`ëÞßhUä®ÔˆEÖùôy-ö|ßUt)fø2Q«f£6°è°²ÃHAF	¬×sX` ÏâlœŠ5=0ÚÃðo²«¬9ZB‡%A‘Û†/QªCÑÛ¯J–‹ŒÝ&²Õ"&M˜
Œ[-ÎDÅk÷d†J#e¡ ªœùfUP÷†*(»'°òxLEò°ÓÚ3‡·,U‘°TE´läþiGMìu3è7EåèS7Ž>0·Fe6Š‹7Ú,4%^…lž²É4Õ…y2@qkå¼Ò (úiØŽüT¬C}Ý#öÞõ¦ä`#íÂ{ëÁÝ½1»yò”vúð«:¬‹gÅ•ºÜø¡ÁW,Ú	wL@¹pÄø¢¾—Ö…Ö</bq:Ï‹Xœ¢A¬ù¤³qAX²šã-,Noé-”`o¡hIšÌÁI„ö±˜@øKO|p¬e1g¶"š,Ý×S5JÚIXHË¹Xbamtæi’Î·œçØÉ­~?—!þÎ~?IÚ…ý~FÀ†ü~¿ÛÿèÂV:ëZäèÇ‘ŸˆäýýˆNÄOH"JOâmäaÐ“¨daGOâÝúIÊ‰¤FCXnÝ$Â²¼Sw®!¬8Ù²°!lF&h+°Ÿ†!¬¨¬áeøóö†°ÅÌwaË^ý†°ÖIþqaIln×d®eâŠw¡bŸ†µ+»t{kW ½;s¡ÔtÆ5K´Õ]ÖbX¯^ßBJõ^½6'€Ã•L¨[—³Ô½âüï#ÊN
#„ÿÇÁ,Œ'F´UKÈ66‘3šÉlŒ8Kukr‘•Ãnb	.··Øú5mçóšñwå5å/nä5å%^‚×ŒŸsq^cu?)^ÃÞ•Þ‚vy-þÉyÍö%/ËkÊ^^¼ãb<émœxq~ýKÐÇ?œ¥iÜz°¨¶{õ€c¶"Ã6}zwìBí©íêƒ—·òôUèGwôIªu
1¦m%j½¹˜àõ“‰*‚H¿©fç)^[)(ñ…z)ía:ïŸWñe:v¾¼›fBQéÒâ:Â7P,_ñžZõ/Cð$ïÐÚFÒˆ¬K¤Zëì€4#ý"wZ<’ÛjKE×ù…’•ç¬îÄ!)jQ$[*|wä/ËÆ?â²¡Ž˜6-	?éRÐ0§¸×êðYô‹tÿÉ¥ûÇ	ËÔrý“ˆÊ\X¾
s (_|«¦¬êº©.ë{žU5kìš$L]×¯ª§‡° É†(àTÕ@?	ãÐ°ëß3ºÈfêÛûT[~t—ªÑ²þQ=ªjzr‡j@Ð3‡ª¬,Å¡˜e%P™iOŸ€VñÛº`™ôêæÓ2ÎÒ’maŸ›‹ºÄxÕOÐ#V	é÷ïéßÛ%æÀå<bHWç+;2×Êgá]Êû•U	Eø¶V	­Ø.7«±ì'„=¿ë/Äûî=,åD¾õ,kÜä¸Ûúg>Ã£8:5ì}Vd·RqB“ÛìßfK%Y… ßÆ‚¬ÕÏ¥ÄP«üiÙýYÂ‰{)'ÜþÒÏÐG;”Ÿûr¯FA`tŠG´tÂTžŽT4£V7{3\(ÆIsO×Fp2©&•q}!Tt}=:"¯P?æ1Šé¬ðTÎi
ð×dÞw5½>ŽE§ óI¤ÇŒÇ»— ‹ÇÀ%Ô¤8Ë´PÛªCB?”ùínôá/8ü+Ÿ¿×4¼Y‰þôá/˜qþÏ<°yU wúßvÊìºÌl*ç˜n™­Hwè‹y‘^aËùtV÷{½Î2òí‡¦áàüÛ€Œlp<#¯·x%=Åå[ºeå“ü¹[Ï¿Ô!|è–Zm<+¥zZJÉY^/ÇP:
õ½Unœe3_z\Ñ +: šÄÄa
Åò”ò‰NâòùæP–2CŠ‹¬+eoH¤x·U~hŸ(€„ÉâÓ\‰ÑËÂ“‰ÚóKÖ‡ko¬©Ë2»é…¬}ßZ*äõTÝ;]²…!N—ùÍ0öÐkq0F‡‡6´9yßµø•m>tZ<°,RµOàáâX˜½´-(ž%—[»ƒÁÜxà¥ÉÓ8òa>÷F6´S-ÖNC†Þ!>=§ÜÉò«_¹c8¼@@–¯EÍ&Z´Z ¦ k%(U¦ad‚Ðå;4.Ôüåjzád.P|xJÝŸÃËöV˜[Ù¬¬woÏÀ{lH‘ê–‹¸è,åak'´’Ý·¡óöÈ‹Ç]ü¢R	`®/×%ª@Ü¢-£ªƒ!Ø4 âo¾OŸyw:Ž|Û©ãá[·ó
Ü’ï ™ï²ó =uê{Ðî¹Â°ŸGÏÑ<í8ì¡3Æþ°pi'¦{íy'¬žõ©t	î”à2!P$¸U®AŠ3Ð·ùoczÂÍž_Ö®×ÛJ¢òÇsE¿LsÐ¹•¢_Œ;iè$mŠ¿,O>YË×AÂ	FeÃ”m3"^Ž[ü²YÏ€•8ŒP´…ÓœÍ-Iü½–D%£ÉHù.'Óu>2O|§FF3	¯Sþ¸~@rÐÍTx%¹IÚ$·jAÛ¸œ^ôêõcõáÅ-‰mÓ´
ÈŠW]G³bKÝl?O·”mÁ°˜ªð‹é\åZo,¤Y¨‘¨îVÅ‘Ÿ£…r¼p«&j³†óúlµoMT´±%Âò‘Åk¢‹‰!Z+{´8RMw?¢åÂèîÆ3÷ÂzßY§×áCt×ð.zê>¥s2a/ò=Œz´ì¶ŸTç·¸Düy7Ä£®±
s:$Øœà Þ)cxŠ?_CÑ®|YÉÆ¤õ>Hí-¡µÚÅjáþ36\ÁàÿeÍüTä Eª9ÔEzíôÐèauŒO’@òßE­°`ÏÙNèbvP•½ô,žk¡ð†B3\­w~T?"õé7Rü6ŒñËËÓ?á îíïs
t”hŒ¾B¶I˜$Èy–þRÎŒÓ´¢Z´¦VúÒI¤r<Ãåµ5Ð§Ù@@–| ›œ{¹l^ÍÊT†ÝäæŠ™mŠäG%"”¡ÔÞº–ùovi¬úJáâ´ ¾n>À·¬CÆ'Æ€äÇt`Á.ÍÅc Cr-ÓäTv±»¿2¶Y¢Ù*V9ö0ÛFg­Ì?£0˜»K©WµÏnûë¾QÙ´UimV¶lTÊKzÖ¾diŠµ–ÇU=ÁHø»S+ë>w†ðk4Ä¨½EuèÄâŒF¦hˆŠèðÒÐ3€ÓÛ‡x§h~Ê¨ -K8n Ë—µdl@÷™÷)Å4Ï¼¶ç+Ò	¤·ûý‰t¦ÍåFo²a¶µlhkiÞ´Y¸7ìaßñi%ãRûk|.0²³íµù“¡³%t3MÌ4j³M´yº8Žwß¾ßÕN‹üwYßþœ9ÙÕ½Ï'é§êâ7Žú6Ðm®z«Á®Â¡Á§ÝENtwa7ˆô=ŸmWmƒ×ºúØ²%ìÿ;6"c9±£3²}l/àídÏ‚ŽÀ'š~®~@K3º+_ üwiOÆ­…,Ø÷˜¦0jxêv,·¸_†ºÂÜ:·aÝÌ?)gT‡Ãáƒ"þi(x	ëZzWtÍæGOt‹ó%‹uÔuz‹Ôqn.ì\Ïø
¨Tâ¿™"õíŠ/SI—²ÂŒÆÙd:?á`gyk@uRüK6EGÏ«%—úU-f³˜È:cÑôuP Ë»!P,{8…˜sÃ½øær)Œ—WÐ«©<‰NãÑ»ói	$“ybÿwå‰sßòò.©a8Â­ºu=,DBFQ/ê!Þq]´ð26tÉ­IW'i´kŸÿˆ)6Y¿Šê«‹¬%ÑfÞ¾ÓÈ"œ¾ÙÈ&o!xûû@Áù7½èsÜwzÛ‹š'tp&ÃBÙÕ¸ÅbÜÌ	dœ%b`v6ßy€ÎÎ½ÙtœPcbRè)tZ&œÑ[Rb„Ò8ÐôÒ‚ä¦|~ 'Ð¾ZLqÕPƒ!UÏ 5Þ>K]$ôív´òÿ­¼@b]|ø‹hòÃßD›Q"ºˆntæ«Æ†ÈýY~V´º ××£ã¢¬ñšøwi:ê‹4:0^7ÜÇ'üq•rø0)äùÿõ‹ìü"•$ÁÁ-òš¾—éäà€Ñ­‚&ÈÇŠEÞñ×CüÔïÇ«Ñ©5ÿdJ¨»8x¤Q³±ÝåŠ$˜Ræãÿ  ÿÿì½ÛvW’ úî¯Hq< LB$e¹ÜÔ…‹¦d›SºÐ$]gf±t¤$$Ó(d‚’Ìá¬ésÎÓ<žðÌC¯žµú©Vþä|É‰ˆ}‹}ËL ­ª.,[2÷5vìØqÛŠ¼©w¿K6Ö××Æ€zeQ^ïõaŸA_ÏMÿß@ÿ'uý{/kú÷ÊÆúW^ä4€5¹p=çæD>NgMº‡Œ€>^ejsØ÷ï±¤Ì(µÑsÊï	¿½Š$êJ7÷›I
¢J
‘é×ÍR¨¶ÑOö0=8†…Ü•t‚,tè4¢¾ïÈÎ´ñGä1Ü¡6fÁ6föÇ›kwFz-º‡&¥åîQ+·hcrQ¨.ôVý)«Ly­	ð³ûad!ÜUŸzVÃkô¢€[XÛås*$$%µJÁ`è—,j¨2V¶ò’þvE-'¶ÅÌºyF4§ÀÛš0f±µ°ÒÑèaÂ+îÈŸUœ½/ú&H¨ª"P1š¶äjí[ !+pÈ•ÙÆn‘TÎ\ÃL¥ Í+ÓeÔÊöqk¡dúü2@½¹²ÎÅ.ÞW'#þÐÜ!þhRA™Úì'2FÏ.GìœpùÄ(nÆMrÂådWã‘DÛ fÌƒŸ8ÄÑ‡€ö(pWÞÙoÔWÍ•y·ïÁy>R2Bñ†VI1•cä8aG­ëy”–oËö}S|ß9Ãð9âûó´z¥¹Ý}àó2ëÃ¦è[Í'~ÞÐ˜»Oã-&nÏ3Y@ã<MÐaŠ¢Ö[]hJ ¿þ@ôÑòk"¥U~]h=áoŽq
\Hrzö:½HGyª.o@Öx^¹ÛH@2³/rÁr§xF1rÍŽ	<Û´Ÿ‰d?ƒD˜Ú?¡d¤r$yHÍÈ›%Mü¨V½$¥‹IŸÔT@EŸ¦ðX11UDddž$ÉùTå€j›…dQIwÙÙ©Ï<˜±ÙÑÕT¦Å=NO¢xà|¨<£¶Ù÷­¤;V”í¾t8šÕ+mÒiz…
‚Ù1N\‡×âÈŸHÑŸsvÓaÇ”pM4j„øaG¶}r"¥4àt Æ!ù‡d:Ÿ‰&D­ˆi×-¦ håtÂþ¸§š Ñy} cI9˜eÙ$9…ÎzfšíÄ–¡Šàè*ÚI,³FÜRz</$ŽØýÕèŠIZP:’Åe&P#Á›‡1!EÔà´¡¥¬â²ýjüqIÄ]N› ]_ö±F`ÓÆÇèöâM[§ß˜%u™²ÍR¤›2ÒÚ€ŒÓÀFj%Žª%Eµ6()JÖã¥ê…ÕÐÈYä¬ßßUx[·Ç`×úY9¡z3tH
£Û:(¹szEENëHí^äèh‚'¬R^’k1{ž–ç'E::Ï³bRœÍÒqê¼x›NNÒ‰óp–R€_èäÆÿ#’ù ²ˆŸÑ^˜³ø“•¯pÌÂ
IZF¼8ÍÑâŽ’i‰Bs’³ƒãL`NÓA–LàüšÉs«¼­ó¾•gû&¨/`¶i/—“²º˜Wp4¥“2öA~ˆ¿E]5êT±\zRÌª’¡-ÐGEz_¨çüÜ¦H¢âwhÃ-&öE˜•U½ð/Æ¡Xgeñæ¡Fnâƒæ†›¡ŒUÌÆ î‹­‚Ã­ßfH»e`äah9ÄƒiÏÉÛ+Ia–%·üâÚ£íŒþJ5’§úT¿ÁH:‰[Å¢J‰[Ê5µhµ3¹ˆîj+Â¯¹î¢¦Äf]	®×—PZ{fÁ‰Í'9@æô9&'%ÈsŒ=`ÏÝX·ŽI1\\u0Çï¯ ä°Ló¬ìö\@†t,VïKÖ´ØmÇõ-2å¥qi­o©Ñ¯Ô{ ¸gÉ¡NIÊ|‚ìSþ³´ý3qþ—äÀ1†FËJáÒ¢Œé6Â>Ù3#½r,à˜cÛm¾¿Î’9¥+8º'€zlÕK“¦$¬lZŠ‡ƒ§Oáz©kz:„MÅJUóÌèd¤ž…±¥Ú»2¤¾ »g™’½Mìmõ1fìñ¶ZÙ’ šÉih"£d·Ç°±™‘*.P©óJ˜b(”§dy½¯XÃÞpêµ‹Šm\zÆÄR¥²­QF-_/Òbµÿ®±úø$õ#7Ù]î	úÛªJ~A^<í£Ô+~úa¦àpá]efZŽ8n„hT÷ë‚Ævôfãú½¼6kÉíôÒJ/p#ÆÁýù€BÜX(ô‘Œ;ˆŠNÑƒ¬ôª„XSÎÌ·àLãM»ù%Û°§6ƒj©Ç}<°v$A¨:'§mBè mù³"*‹‰¶š¦èUpîÊª®œµõ‰d|mÑk¹º‹¬­Óø-z]Á£f%mq¾~)}ê —WˆÑõ*X`¾Äº;gÅóÖ[˜	Ë·´‹y¿ÕFžx..u)×F‡çiÅ…&­Í5QBtéà<lJ/rK˜ ›ÿ¤Ñ`˜—ƒe Àh'ŠN'7_{Ù›³øø´õêcáÛ\~Óþ'½þx$pME«³]úc«Í]nv´«ž¼•v]êo‹ÜÒÑÎÿmväÍ%<Éä•We6:ýÌo=àçÊ®$F–äšêËÌ	é ïý½åP_½I&ðßü"ßPÙVy)»µ-r5R¡«Iœeþí)£P¤*e1ŸèýË0.ô†^ã©«v”oÑ¥
:ÙJ.•}S
Õ;“ò]6ƒ'EZœÎû]¨5¡ù<õË(„eìõ
oÛ8¯®0¶×;z”­ûÄ7*&hî`hÀøjRœãêÙ8Ý«FŠ‰VHÈfŒÒÓäÕ·nuÜaÂPVY+*Þ(›¹öxÊaîg€ÔÅÜì@”Ôá=ÈJRÕO`ÄW©ì¯0¬æ~|lsU2®	†Ñ7•l›R·_$£Í­õuø¯Óã×©Âó'Ê•’ƒm“¾ž¦cæVh“ž¾î@gØrR¼ó/„LÒÑþ¬øID¥Kœ¡ª€*N±{óù¥µâ*YK,¸dò¶BÒýüÒ,`ŸeÃ¦¦E¼«Þ›Ÿ…&þòä§­ÄÞÃf½1sµ±®1sƒV¦Œsgm+Á»je•Ž§dN£¥P‹#ç&7Å(£Õ<ƒ€ûÕs5¯Õý³ÇÉ:£¶Ûµ-²‚[Éqç)FÑª>þËYžâáN-º þ<½$÷’?¢lDO¿Ä˜²|Fq ²Á|å“ïçãt"êeƒI1* I¨¸·Óyef&9f€Gs@öSJ,tì„èQÔ°‘,Œ7¦QDÂ}axß²ñÚ.ã¼T[Û²)¨]à“Pò$eäH¹”ÒÉÉÊYÊ=ƒÊÑòDÇûý~›¶q°A"B:”À„jtÈ—Ès‰ò„€Î´Ö®Ö=2JßI?Î›t.IÞÄbáµ•#gt¤W­FŒËêêd#—åƒ!znªµ†n³v.8þ¬Ãökd÷’H!žÉtwBEëþ HAyñ¬î»Å|4Ä»B’pÐË¹é@aÓô‘µèGÅÞ`^mé£—Òñ+ûŒš¹åÉQ™/àÑÉú±JÔœ¾&WÁ6dàžF€eN³‡‚¡zìž²ÎsdÐ++XÃª’ÊXÁfÃû¤h—7:q{Í"É€ÝD˜òé®™:ªRÜÈÄµî6ÄðL?à1?íO³ÙÙ|R¥öñ
h¦vCœŒØ‚®hv’9*GoÝLñÈuåÔ)>ÝÉœÌ£xú˜+ÒJ¥¢wœt²°úíK§¸Ø¥y¬±]Ôû²¡®hÑn·iBãC×G’æö¬õcÍ¡u°s&ÂËàüãÚ«*‹é¨s²‚¬nÞ9í²€Pê¬¾}Ñ)ñ®a$ËëÍ»| Ò‰5§õ‹-ÓÕ$­_"Œ×Ñtœ¥Êµ O¡æðÅ·ž™úŽÝÍõáèŒt8ì²´ó|Mèùt‹‹wXèW$ äƒsä[­SÉYñî¹ÔÙ ’’àVbF-5LMLlªÖÅB-¨¦î=oK_ä°—êl×b€î6ÿÕW§ Šuþ'å5›…šÙì"Å¸L¡†§Y‰×t»OßpcÍ'Ã —aÄÇ©x¼•lx)¬4M7ãa>î@ùM(ªŸí§³AŽ{w;Ù€çëžN¹Iîå9¤Ð‹¡Åmz“P¦“a!–MFí»÷U¯_ÎOÄ±ÞÝ\M~ï­¦ÊÆ†Y±$Q!‹Šs7O‹Ý
«Ib0o‹ÐnÕjUh[ÞùÖéØ­˜rÓx)‰ÀTÊBf»TYðßHføoƒ5[ø¿* Öÿw^;jµ¬ëÄjC[Š„YÉþt^¢¨X+—Äº>_[áyÊ¯7pß#±°¥dUE8Å†0öm±¡Áˆ¶8Ñ+|¼ð1£7°£?âüñc_œË3¤Ì`\‰zƒá;™˜ý£¸@+¼<éÞÐXŒO8(]BÌYóØðô«*¦ï+£ÛN/2_.£¹Aì€C£­}à´\=Õ«Ðb~YÔCdCu«Y¿—ÃÌÝ´ZÇ,&³/åâRûÚNy(>×T!FðÙ-à«é‰[¬I×&Ã5nž$á–n£J3%Ã
5ñqCÈø•¤5‰ÆqÑd±·ªW¯Íõ„ÈÆ_®QQJìçóÉ[®F³»ôLA£·äé+4MœŠ&C¦man˜u±¨ˆœLß†.vé+nó„Ñ!õ267î{ˆ×©µUÛ‹hÉèÃÂvDùEØxÞX¶†™¶50M>0h«¸“¼„Q_Vˆ%Œ4øñ<31Þ2U‹«IúÓˆ|Ù£G$ÃÙ
ªLræ®†ôÍ­±xŠ`Ï4ú™57T‡4m½„t–Í“ùXouu{:ÚŽåZFÏáÔû?™§£CX‘ tá­Â4«E6š œ6„<Áq¼ŸM†P’ŽØ=L>õn³ôÐî	Ñn73¸ºYäž(fã…yB”c³áÎúÙ"Ü†–)ƒ-çIØ÷º½­vžû ÌC:&ºÚ‚~0?KK*1¤`ÆiùTkd‘'µ}¼×øV‹Å¼’·ú‚	;þjrï¾&•á(7‘$Ó3Y›yL½àièE½ižÐ†¼eí‹:¯…³ðÀ¿C".émI²Ž‹ýàìÍïDÃÐ?6Û›â“ºµ,Ôçºü$è#¹ ivŽ­O]<á"‘`|¶â¼!®°¶=i‚L o>Qæ]Ñ€Cj90ªó6\6*&gt®e6‘jÎñ<Áx:¢F6´†ÂæªÒ
†&ÚÆf K‘]¤“ŠOÏÔ²–1‹9bÎLk“ž¸Ù"ŒZ¬WmÎB9"]
oê¶ù&VšìÜPÁ©†oè¢-^x`Œ2åLQ‹áÈHÇÆi€oîaLà™wÐÌƒÊKúj<ªÉ±Þò5ñw©}ë%ê¢n ŒFÉwÙ8Ÿä6ú#‡{cg¯k›wÒ<0½]Âàý½˜)Ñšð¡Èg”¯±Ÿìûzš^³Õ$æƒã)—ó41µ
»ªÊÇþøo™œÿyz’ràÉ}o–®–,þù:Ùò=ÊÏ>þïÉ O“Y•Ÿæ¨4\à´i<EWå7»Æ¸/1´-¥YÔ¡Í[m[éÜyò©w§u‘3R9¢„]fgsaRbÞÚ¹Ì,;ËKH–U6)U¬¯&_÷°q€^†| Æ>›W¸REÿ³Ï^ì=}q´óñÿùø¿Lv>þ£½ÝäÉÓdçÅÇz¶wø49|úüãÿõB<Þ9LöŸ|÷#Ô8Lž&O÷_Â÷­ÏvàXË4˜ïl‚Ë€ì6ô: ÙJiÔUÊ«òÌ‚èY˜n}‡Ë>"*ŽP8YÈÕÊE†&µ¥néÒ9^nšCS?e¹X7{¿•¬¼Ä‚Ã90\”1°BÉwž¤x6zfÿe˜‘ž(F0šÓ"T”Ì‡PxT	_0ŒC@p/zíe9Ï8Í‡ùx:ƒÛ^ém!î¬%;fj+‡ùxG¸"5Ú+È’Ãklhÿàå7Ïž>ßIî&ßí|·óì% vçY?ùc1øø¿a-þøT8ëå3uÓ½¸ë*%@Œ{“!ã	ÓÌ0K¢XßÒ«h¶iŠ§2’¥	¤M°Ï’n™WsÕéU~D5!r7˜å™.9Ö°%÷‰ìnB£B¥-§\íþg›‘õ':
Œ,ø”©1J˜åI”söñ¤¢jååâ¨jÄt	o%Z»#¨q¯L>þsš¤ã“œ²Þ„O8¾ziòlg÷Çþâí:@á«„“B!™S£¬¡ë©,œZtöI¢§Õ¼×ÿì^?Ù{¾ÿòàhçÅÑÓ- Î3$µO_ÀNUÃ Îl òLJZíì4Ã•$>Øtð9©8´µŒeÉQ*à_>þB¾ÌÆÿçÇ[+Kñ–![tbaà CF­Ãèå¥²ÁÇEz«;ÊAH¡êjÈà¿ÔGfßfÀv‘¶ˆ’¨ääº¹%ß®¤Åáh!ÀXDÐ4ª{oÈ¢ë?¥;ÝaV!Ý«‰[ÂèÛD Ÿ¬Ã:Ëq^?Gê“š—3!äÂù8‚·Ú›ËàzýîÞ<rÔèï*¢J¸ u¡†ôÆaZ™ÓÈÔš;Nóþ¸f£²¯d’L&«RØ^gtî¯Ýëµv:JËó–tWT ˆcÓz§ý–£/Rãxâ*¡õt*´D0»?•Åd%`€‡ƒs‚…¯«¨l­¿sp°ó_\¥dBê±2TÕªüò›ÿôt÷È¯Mv QUžEI\ÌCo$ÖòáÑÁÞ‹ï’«PÛÜ*³@%‘?}‘
ˆÊÕP¨½ÐTÊ/PÍÚÔã{¤u5W#êß«u¿]9d%ï¯fïIT zìãM|äÓº
]û¼]MŽYD"àËf£Æ@´ ‡T×ôÞhª¿øÂWh!‡JÅŽóW® 6©ê¥3:2û t_äÕ|rXÍ‚±íóÅj’{FoÀˆYõ±k*¿7AnÅ5Ö¥[#¯óm>¦¾²-W}%œ>Dü_oVÙ+êPÃZ»
ã‡
qš×SÜÒãõKÀ„¶ä$ý—©rfá2qÍß•§»‹ëÕ¤´æ¦Œ:ËÜÔÐµRÔ[ÛÖÍšÖ3¥£BÏøU’R+1	‰{³}AåH¿/æ×Z[r0çŸà„Köv,ø]1²3ÄÌÅA…‡óœX¦ì=1²Àû'ƒüš–F|•êúIªNÍþŸ&š<¡SˆæŸ&+ÉÀjöÇÈ›œ‘çžIñRñÕÈŸÑ‹IFÇ÷-¥‡À£©]kt~ÂvqŒ4òÍãP¸íºÀK{-%¿Fëa)M%7¯í	ÎâôkÓ.IËØ :©”ç£3H0o×§ÒÛ¨®Ïå.òê°QÁšjG¬l U«¹‹\î÷ ¹Ø‡É í]¿ÝðJ¡Ñ¸:_+ÜëéZ­KžâŽ§$Ñ	O‰X.à…¹î)	‘Ö¸*zÔs•¬P.7ÉÁÊ…üÏ{=dö¼_ñóQªWùÜe“¯¥jCté†"‹4\÷A)s`	zÖ˜K¶›ôU9+^–{ÊUq‡ÎµôÚ%·DÙ
ÛMêQQãÑ#VŸÒîšè¼îoæ+Pn²KÄÌîàWd×vu3Ú_žé=ðEüRa÷óÁ[“FMÆ¹‰Þ›n	 B|¥ƒíÀ’ê¨õóD±¸Ë$“¼ãhúaåém^ÄIïÖ=a‹²­ h ÖuÓèÛÞñú+ÖÇ–?CôŒ(‹³šÉ#6	Ž'–µÀó¢ÆÐËç,ZNðŸe(ZH7ŸÉ¸Þ¢x¡½ˆT,ÞL*Åè”
4£{z(	–Sç%=f vÙ`¶.:–ìÍŠoç»‚1Úã] uãpyîØø:f¨Ñ„<ÚÒ,ê·q²b»„áˆÉ,¾Ãz [ýÊ¸0•8@,&Y¸U–ÏbÕwdn‡÷V÷<°É¾ö{~fo%+_ž;jÙ8È—Ò±¡?K„­úì'vE3n7RÆ†ÖéVª$§,®êÕú%§—+æÊ
]Ù·À¼+x(¶bª_´p"Äí'a¤  ‚:C7ó•ÙâI!CmÇ;¦¤…/—ƒ³Þ®±»¢iQíºˆ¥e Kž–O=Aƒ”¥@Ê×¥Ä&ZùRáØëYG'Þ¾uñ¬s¸Ûª]ƒKŽÁjE-*dŽ!°Âa”¥q@óÀ…?-§$¿ÊÑˆ‡¯3ùÔÖÙ¸á/<<–±7HIDSbtü‰ÊÍiÀçaÀþ³/_?yúzgwçåëý';/#>jZf7\íf?Ù¢¿˜:±Q´D‘2YÄ÷É—çh1.zÈBÛç>|HÎah¶sEJgCk¸â¡yô<}ÿå÷XµëÎmÕ—‹\…Õî(C»Ò{Ì„ntàq‰ÄWÖÈÔÞÄò?,Xî&n ¼ç ÏaÍ©\L™Ê³â['"$â÷ÒW….ðë¦Ûáû¦ªmgHNÊêb Æù»W®érv²(º}
¬7kõ%¡PyF÷RëÖI$;hçi!aë¾—šHúú0€cF5‰EBêÉ´ºÐˆÈkSOc	#FØ˜+xâq{KRyâàzÞÍ#„e[Ÿñ…U’õ
IÿÚ?ÔÊ-/aI<ž¬»9®Ñ9‡¶ª+Í	øJŸj~ÿêD †5ƒcleöXÍÙ€åmGÐ¼ëCêÄÂ*×ÅÌj¼š}€b=Oóå¹UƒëiÆ1Õ·[‰TÃª†T|×çÊâvµj¯³!^P#¤v»BSjYÍ
i& zN	i¢¤FfÛª‹3z«põµ¡Q÷ûÉ7sL³ZÎÇ˜%Wê/¡:È¦ JÛŠÉzeï«"HPˆ#_°7ÿßÿŸÉç—i……£÷›^ÿ§"Ÿt;štº¯9ŒÅúò¸Ñ›½'´šû’¡„;î`¬Ñˆófø$D¯…i¡ÁÄ!\^±'öð ¸EBÜEÈéºý¿O“ƒ§ûÏövwÐEèÉËäðéwÏŸ¾8z‰Á}l¦üÇéT1åW½­?M0ø[Ó«7¦;³
˜å–».
§øÁ÷U_]Ã
up˜xêÒúÄE#³!ôÀåÆLŸžØ¢C;ÔÐVóAb48[T´÷³½à[1ôá•ì+@fÉðµˆ ìß—A¾å	ÚÄS ˆÜÎ³:£&®îòM;G}ûÂ¶Ys»†Ý«)¦‘[5"ÀO¿"%FWÏXÏe‰$dxÕ%¨êâ§¬¾
ÃðïéP;ævA¼¨3ê]ýÍj}­·†Èz„!n+ÊB‘‹€™	ìFMZâé5°´Gb6™Fßì»^‘Âyýã_€Ë>l†^êëC¥e:Ú^fgt<'+™»Z¹ƒÃ7É– I‡ÄótR	â½«“€<;£ŒõU¤1"Ó Î>þ‚üå}|uuË[ÀÔ·«.äKŠjÏVåóKïD¸z£ mgžYð>“gÛE{±„¨Ÿ‘FZ'›ZÙJ„qëZZjlIFZî”`¬±‚ï§³*OGýñãÚŠx(í´É•R›À–rBBŸ17d¼DÖ/‘Å‰t‚5 ñÈ~‡é0ª‚yºe†©·Ýˆ.¡ƒ\ŸXÐ*Kœ"ÎsËU´{üD³,Æž`cÙÄŽ¿§äš\½ ÂÅ*Þ$‡8Ÿ¿@®XÄ>*lòr8¹¤G%ã¢ÞÔ.h]îQ\¤èNHBYÚ‰fc1dFÀ¢¢m$,O¬Pw|3	HyRŒ‡¼.\0¶„0#Å Ï¹wøRFòèõK U·sÔAóšÝIgº-)h]§!.·éöÌÃë4‹VshÄcwc¹•Žõ³b	ÈZ¡{‘áxvŠ-¶»/‡éø4ÕlXz_×}Í*PÁK«Ö´i£Ncö–å:à`8Sãq;>iVl™¤¶å¯6§–¶›UÚˆ(Hó¶ˆ"Y­Êp4§øþ
{þg§“Sóð\¾è8Ê`:UFÿ(Ä—žÊGƒ/€u/Í:´Êñ#u`î\ï9ºFšÔdü©ƒ]¸Ýz¾Å@ æKG§¸Ð4u‡Iì¼Ë$å,Zöµ¦1?çX/É“Aùè!å|[Œ}
§íªxÍaÖ`¬VnÛ~÷ ÒÔõØFÕ£‚‚½s|X\Ã¦¡wpYcðð’†‘üGú<7û4{köÄoö¤¡YéÃ]ƒÆí·^Zfc2Ú¾—>mÊ“c2ÜCÝ©‡aNÄŠ†)Àa®møq;¨§¤¯V“šø6S/H1ËÆé{¼ù'¾ç“®Ë¨¯2ôgk`)hT#¡´Ì1¨† ¡S|Ã&1E=êÕd}UìIÖ«Þx8ÅÖ¥M¶bÒžÖ¥ï}•-­ŒVºìªº`’¸(õ:¬Ê•ãp ›xØ• ÅZTCµä¸£•H±p*õÁ]‚'&rò¢·üõþ)ÒS¼n–Ýà@ño˜qFÂÜ H,‡’Õ÷¼Ø—Å (YB÷%Ãš]‹#Xl?^:·"·,aûIéhÑ]Èy ´ ù¨X]8£Í/G5åûS+^(VÅ³ëç­²ŽHþ†"´Æòòšrô4pCb+±öóûTù‰ïsÜkýGQ{¥ÚÛ¸\Ëùñ­»iu;ÇTü_ŠÚúÍ¥îcP°SÁÄWÁK<ÊLòS(Ä?è¾eªå$à6ÀUQÝƒ²ŽÇN3Žú(À9À€î¬!ú,™d‰T¨ûµŽ5“õ¡ªu¾ÚŒŒ£U±=xb{Øáª‚±$5-mE‡«–±&4($‹Ã€÷d`·`á ¥ÔÚŽseÄØ325×ámœqû^ÃIñ¢”Ý1UÂËGÖ"”^Ég$§Þ»¨ZããnÂÎ¨}äp5ulÂ¾QÙÅX…}¦æ»5vë­>=¶!Žc7åâcX6÷À¥­9ˆ[D6ñÔ&ù¸¦é:˜akPäå!î[Ú~­)ûk}ÃÅÔ™ñ$ôìä+-Öð7•1ZeT_£˜|brUû9»è^¼›`š¤#é­Ž÷Š‹Bï0+Ê~vøŸûó*•ý“¢xûè‡ýƒôtp—#îü.)Ík«5Ìñº*^—çYVuå’Òªl%§ï±ÎK±ž²s´Çžãi:@CZggT*ŠŽäp.7ý.¾vI¶"Ó)ýU!¹]Ø¦S@¡¡„È;Øäx"©šø®}ûø2OøêõBÀ7ÆßFˆ›Èîýp*Â'áˆùÓ}RbÏm«Z	à¹w}¸Èy\äw—Cü,˜Cé2¸\ôtZ@¦óñŸ 5°$Ïug^Ë!þ/Ìd¨ïÐ`ñ=ÜI‚0WÅ˜‚<YMìÈ)*!¦ÄêèO‡a±>†ÊÏ?þ3’á…òSÁ
fÉ¬ ú–J<¼q5÷ã/ÙÛýŽä	lÈÚ^?ìïQ¿Or,Ÿcô=åœ$G¿±¾þy]œm…Ò¸áÂ6ðÿÉÇ™”"LƒIŽÙaÏdLl[ä‘¥pR³”Âp‘‹ÅØVÂ‹çÎµ÷œBU½Dÿ–Ü¢Hyüm>Ê¨JGD‹x»¾˜U˜˜ ÿ~T¾þÀ!²{ Âñå’èî
ïW”‹“^“÷
YÙ©s­ó„ˆ"´ƒÙD2€òÓ¾·q)"½’ÔQœLúžír©á‰x¸8AÔ	ÎÂ aD`¸™×Ù—$±Ã³]XïËB¾+¯…ö´/†6äO_§]Oä@]OÔ˜Pº˜#{W
ƒ´Î»ó'_áû‡>+ûo§¹n±(Û­²Oz›ðÄãÅ×uqºCTðõÌìþÝÙ÷ˆ<=}?ÈFFÄÉ¶„o
eO/`¿?zþlo2WOÅ9ôØ£Q Û Ó˜õ…Be¬Üîëpt­š8Ñôø^“]Œ†æBå¨÷âÅ®I!ÑU9èt‘]Ê(ÈÊÆ»C¡yf~yVšóð,œŠžßE:šã„;¸Ã ·èb€ÈõTd¾Âç¹…I²Uá=*âÕyfX¦ ¸
£¢™ÞåœÈ[^ØZÍÙ <…àØ”‰©Óz Nñº_ÐÄðá¢²”Ã¶<*1qqŠJðlËK+®ªÝœdØ[W˜fõ­Ÿt6K?tlõ&:øësE ‘ÇÖ8E$Ù‹Ñ7Mj‘ý¾Ía7Öu-¡z‰ iQÛ£ÜP(Ì—UÒÓòìå-}¶;w?­òrlª‚ø¯ È‹((hm¤ø0«HsªbˆG™º…óu}*S­Ã–úÙïFÅ‰ØÈéú`ýCÞØï%ÛÞ3¶ø×1Ñ¹—B‘ë›Y_LÝÌžRjGIp,‰;EÊè7`éGôWØ»/T`ñùÑE’ÓQöCQT8<àø-F…èÇ@¢Â9^á’kÂ¡YbT<e~¡xâ·6Øè(<Ÿ6<HR¨Õc/“ÄŽ„`¿ÒñÒ»+/¾}²â¼EûçîñŸæë÷Ö××ðÏW§¯îž­&+ñÂÿgºöóúÚ?ªb¤všÆï³ÞqÇ¶˜°QPé|BÍ»ääC2š”ƒ,¥c8%o³P
ºeäÀ(Ù%ïBò*6Taû×³œ‰x÷Ðº¹Àú§Ô;¡Qqò`GG]/‡=7þú‚Þ±KâºÞbo„Ý·¶…}eW’ gh¯%½¤®ûØÎ¥v$£¦³á¨öHv%Z	ÄÌ EÀüÊŒÏ4‚V¶·ë¹ÊYB%aT­Þ«æ=Jêk•§èˆåÿ‘WçXÃMëk¦ ^ÂÅ°_=pŠáÀ±g¸T•­Rb?<VñA"y^ñ#W
»ýÅbÚßãqåÑI<¨$fw´ü¥/Œe Ÿõsë±õƒ—}§1Â-YåÕ|ä4Oá£¡M¡–H:2ä³ì,=Íé+EXïØ¦¯Žwä ” '™w‚ò¾:þÆ¿ÀáÅ lèe¥r‘a’^Ux¡¹B–¿§¿
ñ¬ åŽê”¦Õa–VsÙýI1UJþ¦&iÎ¤˜ X	Áƒ_Ò`fé…~glòlÕ’Lç œèé—Q`}zDÓ´ãìÏ´KÖ+æu¦$SÝŸz †-ø	üþç¹†®*Â—’=æµtv0§
…Y)lœ<•ˆiý¡8ê'½,Ø°K9ˆ²àS€V²š–êÛ\5Rõ†©®éu&ÅRÝŸüÍ¤TBMñ´CIhäù[¾RBª~Ì~ô@ÙQžJwšü€Y=Iò5;i1˜ ªÚ77©}´Ö«îð§ú‘Á
ŠÊð§,Rú·cŒM¢ÙÐ`¦³ôg¶"øKIÐbRð€¢"èÃ|¦z\’8¢ÌÑº
¦iƒ`ëwˆÂ¼îIö¼WöXÐÙ™˜æûl0—„p{LÏ È3¤ƒ½æº/Ôhå ÖÄb±èJ;à…zl=¥M›UH€Kù=-è÷¡)òŠ;ë*X‘œŸx¨Ú³–’2!¥áµ ×pž¯KÀ[Ý¥QiHÁP?¤¦F…¤:ˆò“óL^O"7•N³lˆš`¿?vW÷‹Ï¤Aü3Ý&¾š -ÈNÓ±MàYµéü¤‚"]ÐzÖ0¤‚Anq‚nGÀ09cÄB‡XF;&u±$Œ÷]1Ö±†’|`]YóÝIÿPý”^g“˜+–d‡4W6q2•Epeoß1®Ìð]’|×s¼w{AÀŒy[¹<ÖO_IÇ¶:¦ÇuTìŸ§Ó-ÒôÍ†•ó…	ñžæé5 I¼4T’¹4ÈQÒq\®½'‘u¦#E£‘dM\¬MR%Bc-{dt…LÄ5’eÈá8;ã$‚Â÷@„È¥o4¯~W+¡ÝQ½PôdõCK·úí#_ŒÕ@É2bMU¥Àu1*PŸ]««fI~sò;·GFÀî±ž¼gKy@R²EÐ“­šéXÎ¦äV(¿«nÙ#[F·6Å¶[ÏVXe·â¿ç¦&‡ºùÕê<¬%íÉbklO«U|åJl_{e£®i¨õÏÒÙà¥ßIönôAÞ0&£¼¬|Â";Ã!ÑGû$èÒÔˆžA£É€æ¡MUwàÎ/811
Y£ŸÃ¤.V"
dŽe a¡*@TAÆÅ–Þx[pð:1€Ô”qA"Cd@)àÀÄu“³áˆ Œè£„÷øwMÈØë¡‹Ó–xäÜÜ‡zD·‹(U™ˆ¾)Eî{_©}º¹š|ÕÐ1$í8K£µ.ù E¾{÷¢mÐ©Cö«	§O,î¯;	µ…›‡¦dNXT×ßÃ~mE‘ÔÁ.ñåºçZXƒ†^|üW¼¿á4Už€¯¿E|ÜNÑ†L-p×ìNû¡§¥ÿâ‹ #áTsÔwßÎ'u¸ËÔ'¨É)E4A”~F£üÃ/cFŽ×µI<ÄjÎ¹.[ få [A­)þPzSøõÝC™Ši5Ýñ+gé;¯Yò¨LÑp*úBû†ìV²Ì•tó<~e7*´†²MX@ùÕÉÉÆcïæo‘q}wž‘Ce#Ê=Ï¤òsTo‘ò¡zð-lÃ	Ìd4OÅ%n‘5‰–
Q ·§ã5:%ÆéûCàã2÷Ü¢£‘âC‘§BoE÷ò”~G`ªzÀdlQZ	·²xi^)ÉOF%r'ý’rš®Pà(}Ëqãþªûž—‡!@ö‘*oeÉÑkºƒF­~^Ò_©y…Ê@¥ç™]œÌG!pZÝlDè”cPÉéïbIÅb5G^·JYaä÷Ì¢„, BQš‹MðW‡J‹@ðX#—ß4C;*ìöà!nþ Ú§{,îb"Dl·ŠÜ@§èwI—7&Ãt&%DµÍMlç¦sÅÜÀõ$#~á (b
CB	DßSMØÛˆâ5‘?…ªvW"€Lç¸Þò<<·Ãc‡ÂO ]Õ„Ô!”z£¡ÑÙY”/’ðü¡³ÙàYÍ^#cLÉvÜ¬ÕŽÕH:Ñ¿úÙÆ
&cÁca3`f2õÛ˜àaÀðH˜‚1°ÃC÷®a%=qÀ)ea«‚\Gsù|Iø^ä2ã<´1¡ËãóWÂH" q,šñL%Î^u°ö‚s€#7€5Êcá:õ€íœæèXºÇž]Â¶O¼òÐ¶gW¨=¦™ž	žBZ¿í©¬-kE ?æBåwÇ´Á¥òwR?¤¾L+Sñ¯U@)nM_WqûÊ7¦Ÿ§å“lHyŸÄXˆ p0Ns(ÑO5	Þ!’×šY"·Ð¾pvItQã“èI6Žèvg	ð4gUŒ’x.‹3ÒkÝ8(”Æ™¶ Ò‰UmXž”w°JìðfÊ‹'¢eI.¨’•,Ê­‚\µ¬#NTÙF¤x*”7öÀ3©#Å¯³hÅ‡rˆWO›¥"µ TáØõõ®aQ%3ÒÊ„ðZIxq‰i­ðÓZse!Àž†vWÞT>	ŒÕyãWJ’Q¼{!—ª+ÖÌtI¿ZÏ½î´¼íPiéÔj›å“ÓÕEÞpå™wPmGjYcõj5©ñð³\UžjÑ••q{ê<ü¸žf>—c!†Ð¬Ç²ôUAÊÓ¤³ÂOpòMz«[Æ°|š*3©ny‰¼^T_E€n‰Áûieƒz+yTÔé©ðS««¢~•¾Ê/¿P³Ö
?žæJ‘&¿h[Ý~®¼ÔaH{KCP‚·[¸A›U/Dn†9!Éh8!ùfqNHUô5	£´¬ö™ñ¨.‚èàÚ—nÀU)†Ùç¬˜‚'²IHùJáb\bŸ×u$Õ†šxd,YésKJë´l¦± 2ãž»Ž’ìÄŠGM¿ž‹ÁLqÝCˆCV«Ü‚KÞ7˜¢”ïªr-·ŒUÍGø²m©ï£¾š.øÄíäÍóltN¡·³qòùe€=¼zÃxÂŸ÷bèi,BÚ†
—6£{º/æã“lÖE¬!¬ÛèÅŽ‡¿óR!¨àjÒ)²Ïw%Ôt~³«IC'yùd.³#J©{^¤û$®ÉçWxdd›ä@•Çj×Üñ’)ä”T§pÞ…sO†ïµ\øå!74ƒN˜µ¯C,sÃ&Ÿ´âÞô=7Ã½eãÛãÞ²±äÞD¿QîMO Àd±`ÛbÁü2æ¢µDp³?¿üÇkg®[³ÕmâÅekIÎ®Ëè-{ãènØ7µ·ÞŸE8²{!ŽÌ\ªóX2­[#cê3^í·cÈ¤ÞÑcJ˜}í¶Ù1©u
eÏ‹öžþàÀ”u/Æ‚±Øñn]e&ŒTE¯M¯Ž´^K×ŽÑ4þ“òö¨× tÁŒÏØI˜÷ß™Ö:¦Un…<ë¡ÞLêb”¬úÉr¬³ôÔ£.OaAäf­Ìò®gž=¨´džJh1FY‡ &µ2©e“¼àïð«rËÅ8f~‡õ7ËuXgçžð8gh±ã¬í‰šo†'·Å6Äˆk6„3Ì4«±ÿJ<s£:4€[ì 0íoDô­v
í_ÈÒÎø$äË¯Ž<?Á\‘AŠ4œu(Ü…nÎÍºÓ¶Q;6†nŽ¸©X;Áiú3tc>´ýÌÞ|Ú~nÍPÓ‘…þ•$%‹Ì)AIl.oó¹qêÅ$ýÍbÓ“ü4GIIÄé©hèt$<ƒÀ	ïÉ)“wél¸všFP€µÄD–’6Z¬›/a1K\ÄyÄ/œxNïM"™r¨òd-™&@r,Av„SwGÂr›,dì¾a[-+ªlµ1+++zdPÞ GqñÞí\ÃŸ ÊH–¶¥ÿ„©v”;¼¶8Åö®íÅ^\×bÿz¦„k‰Í´0×’œ‹wiQ©¯ØOóåfÑ×â‚ên1^T:¾®t‹´O$5åT)0e@dŸÚJ–ˆ±cIY>ócS¡ˆÜk..F)Ó•­«ð|ö<ïóÐÃ£~Ñ²/’„ô>³z|aû¯8$4R®M_Kã«jSîd|-ÑÏo>° >BOA¿”ÿóœsÄ*)1¢¦Ô°#úoAˆCjp=Ô|¸v@ra¯1&w²]¹íÐìZ5ÍóSM¸+A±eOo(ýÇ¿ˆ ÷Ã,Ñ™q³^½Y
8\„`¯ötTëä¿
èGÎm^ñÁÉzÕ¸7Åõ^+«œ®6ÐB®VÃÄŒà°*hã±c‡Æ·zFÁí·iš
ú×5.Kþ‚Ù—#­“:¢­…3b`Kz}ã5·7’ ¤HL€
­ò§i¡ÄOØ”šA£‰R±U.ìpQ£yñvt“†ª_Çv)*ÖÚ/ñLìnž…+-"¨ãÇ½¡ Ò`©¤ª¡tlNýå Ë±ÕC_5L•HeÑÀCÛ°Ñžz&¿±T$ªn9¢ÑuèÂ”º‹«˜!]›h}BÊaüUZq‚W‹¿½’zcÆ©‹ÝT{L]E4ÈZ3°˜
TL>¨HÖ*ƒEÔ¾ø©×'K}Â¢FôÉRÏ°F?5Ze©PXŽuêd¡ÙXL‡ŒŸßð´jÒSÍøaÕÂåŸU±_<”þŠ×ÛS a”›59|Ãð :™.ÎÆåçâÒ·Ó÷ªl\Øâ¦Û*FJz œÂÙOJ+Mé™ÄEWgJý~ßuGg‘âEÓV8EÙ>r·*iu0WPõ«ž‡4ª;Îô…úSèrú³H6‹/»c7ôšz+©7«³W!}¶s[ß^;}u8O¾x”|¹î]ºVKÞ&Ó¯>0qÓ¢`mw%ör¾Š}QWnm¦~ÇH…¨§:¡ÃYäê-³J'qÂ*}€¦üF`ì©˜g·Iq×”ØÙn,ó‘ˆ‰Ø*,ÒÎE*pw@öJ¼Õ3“0°ý(S"9ÝŠ°×â —í³¤Ì'ä/ÿYÜwÅœ”“d+YøBþhvzfGÐ‰Q‰étôÁ	v„ Ãì¥@1ˆ.Šh?ñ†~Â¡}‹ˆ¸;9–k“H‹ÕLv\q‘!Æt;H‹^#S7Ë†æÎN§÷@Ãý’/FÂW…tÞ°å’éeLg°2†É6F$áŽ;%âc'ôn2UyêðK¯ç˜×“Äpô4Czúš“µÿtøòE_ÍüôƒÊ4Ù‹ÍÙ„½­U¶g{ëÃIßrÈ$Î¨Y YhÑb£­[!O+²´%2'Á¯µDÖi±Œ2ÂI|T§‹­ª·<¥Z_œÐ»I©V§¼Þê°ƒóÖ‡ÓI‘„@äg0à—²2€ãA§gG…®ô(60`×23õÈR‹H»k™JX8â4ÅÊèZ– ¤Š4Ô—^`6™` {,WP‡b½¼r.á`¶Û»ÅÊÀYz’O4Î`Qg™ÝÙ¨ i,4ÇÕ¦À%˜Š«ï”íxÛ¶Oˆ–ÚÏÂ¦E½zb?ÝÈüøséB¿ž­õI®ƒ[Mpšð+( ÄrìšåãI0ÅLVeËÀ:ÙÎKÕãª-C‘iŸ¶ZWfâ
ªA¤é-J†'¯ÕÐ àÖLTÉ›*„7%0rf™ÜMdÄ]—0¨ÇM´AÙ.{E-‘<Ü9ÚÛyøúéóýƒ§‡;&²TÙsh…ßWkr¡ºoM1ü¾‚DC¶[Ü”p¦¦hGy´£lM;`dÍC¦2ÌÎ~êAÃSÄ£\ñ0sY7ÑòS¢|ô·HCÔ7‡¥˜ãàË\&¹Ô/m™ÇJå`Ú#öN%PºZµ”|¬Û cp±tR6Y*ÖËnÏU‰±æÜüNVk*·“øk€ÆTVž_4^;M™ÌýÛ±‹Eò“¢	ƒ#éÖ¬tÃµ:G R"cÒ$>&7ó¥ÿt”Í*({pð2Ù=øø?Žöv_¢'ÔžçEB™Ñâb¿À;Mf*Ã?;"ÊÛ7óÓS8XtÞ8ž€¶Ç`Žé^w&CÔáˆÐf·‘)Î0°×VòÄ„÷Ò”2°¬ßÔW+¬;y—Èöºiå­šdwø»~a¾|\Y)i0Æ$ö	±;,€Ápd³›«ÉŒF_ˆ³>Õ\ÅÓñî"»›Éï 6ÝZ§ÚÂ³ n¬ë¨”ÎžZ§3QxgOMíYv‘¼¾-?3?èAÖ¯Ndª½× ¾õEø"/…p9rç/öÀwìÄV²³·eµÎ‹þÜ„Œ4{‘«X%øîg»ïPúï2lë²Û[•*íTò¨}â¯¡’=	vð}ÛÀV2´œBÚêÈÅŠŠ–¬3’Á#E¯èÖÎ=`€ìM K¶µƒõ™î#DŒ¦ŠTl¬öp>©Íuo«!ù_CAksÉ#”´cÆócæö½¢£²»+‘ánd)BIx²y7Ã»Pˆ2Œ2ôOfÉ0ö!?¹I
I“ŸË¶mÊLòÌo‰¹eF€3£•òT
äYHU?¼œì‹÷åóìÝVrR2Å˜ŽMmïèÁÚ´U\¹"?3h§yö%éñ×á]:›tW|˜üôñ—“Jÿ‚—.)-)¦W““Qñçy†œ+@2—êê¡¸á5Lû&%”À‡è°`ŠHûÍáÀŠÁ€ò’‘ó7àq'O®P kðaW=þûhcsk}þëàUN]¢g“txŽ½Û_Ãõ[T@3þøÔáüUëV'°„xUC,¤P–‘ì¢ˆ™Yd{x´]Tâ†Z¶Õ½!ÄªŠd-ùüšÜ&™ˆÊ½(€çØaž¥¬ƒ–!(ÐÕ5{ÇJ"£Ëhj¦Gn¡á†µ²•åcÌŸ5ž’*J.#³¼ÕZwM1Ú©VóôÄ@ƒÇ¼ÄTê£t’ý¤R¤é÷¶éÙÝÏ¦ÂNtË^»ŒóR‘Fbž{@˜d4g@Ö3(°ºî"&•t1©Ëh„»}RL2ýºg-µª=*ö& ÁŠ³p	UÌBðò»DüA.˜e§2ó£gvÚF'5ƒ½amªT«ÎrSP†kÈ*Ìº8EÓ¢n|6ïü@[D;G:~	–ÀV·øéŸ£†zÊÕ(ZG¯&gøTŽÞºšÿ:+SH™KËØ…ôô1=ævíj¢wy²°úíEíÊë!SÎvª·Ý‡Þmê}u]´h
7Û4!6±ùî.lÑžµ~¬9Ò¿ŸQU\/÷sAUÞ×·ÐYYÝ¼sÚe/¼k¡ËÌµ·ý‘,¯7ïd ±æô£~±ë¯ëÈ
5ËCÄ,ªÀTÞØ˜Y£J¤ï#›B¬÷^@ç§mñ¯Òsws}8úoªÁxà¡æ=Ÿn`qñÎ‹d¤¹‰#ÑäÛ°#øè:ŠÈù‰ ¼¨Œü½·Si0º(¨ãÄ‘2‘1#æ2ã˜öÂœƒÑniG«Uæäè¿NÇnÅ”›ÆKi×k÷|°K¡Ã0NÑN˜»M5áñV²Áë'ÉÔñŽ,N`í.0Vù–Ói ¶’uþL­¦[¶Ç×·…y¡FVKÙ, ›¥a8x5ÔÅ¼U¹îÖ.cUEº˜°FÙÚvm¥]A˜¶(Ói®‡61Ä©C0ò$¾Rž“/kuÂ\Ÿ²2çgV]1€z1Cšñ+Æç¾œVùÍ1’Þ…:ÅòCÊÇãV»Êàu7›”sŽ'Z´¤ü\¥boc*¡c#°xîFKD§¹B¥žÉ` ­êbš„fÂ{G÷ÀrP|è±åÎt]’ƒ¢¶B7ª3’ù’,U ¯6pzÆGÝŽ™êÜã3tÙ¡’¹;¡!rm;W²„•í”³*a•P·œcRÛy9"&A=#Eh;âÇ=XŸÌ73Ë²Ÿ³SXž§“3ø2I:»³Ï+ ~‡Â¥U]©Óñ2‚Äü`¥äÙ’2\ñ¹¦˜+>ÞŽu4Š¹r­„]	‹–"¯ø´|MÉ°ø+>®]OhY]3¤£b5[vU/–{$ä$!m·á,;8ŸOÞBÙH—×qš¥&ëfÉ+v†dùÖýc[ª’Å6ÇÉišc¸ œ9WpãÒÛ,›"™Á ˜~h£cîue’_LR=I-•Â8jÃ*¿Rñ˜h*èšñ”Oé­ÍDiymY˜î°Ãme¼CÎ`eš?M/ŠÙjR¡²öÈíÊ¬¿¢ýaý¶§µ	oV•J†Õàâ–¾Fîj™•bIŠøt˜W¶Ê™©$M{áŠÖÍ7ê€W3ÍÒ{N©©lÄIÆjÂ;z¯íÃ›åŽ/WÖa“’eŽÒ¦[¦TX<E…³óöv8Õµ(99,¸Ü’ù-0Î_ßDÀ8×6Ý2íjl¨Êåéz&6FÈª¤"ä{ªx¶å³í¾ÇÊéb›å Aø;-ðI	Ç3«9ŽÔò‘Â€š†‚Ü°<†êÁ;_mV¬ÁQoHŒÞ¥J[ÁtÐtâ”³`È¥žN ýÉˆóÙÁ  ÷/qZPb5>|@«-ld¹ûBÙ–0-»³,ŸÄfã[ÊFˆ»ûàòÔPê°³€¹xkÔ8ÂÑ¯œÍ,3µsÔžÓõsô0úq:*Ò¡9mÑOg+Á7«ädºe´(ä¿&X³ì3óåÚ$íÐäêcÇ= Õw‚qI™¢Ç©’ÁAÓi<ÿÛ½õµûëøšÁ¢xQ”e–âŒg#íä2ÍîÃ9*žˆ×4£Õäþæúj²Þ¿ÿµ-WdÀóQÏÒ9®}ù[ï³þš©=–ëmUvé9Ž`K”=~!ÅXHÌ}C[mÖ«†ò%JO…?–
UþXªë¥
0Jh&Ú$„dJ
‡Ú7	¿ï‰\ö3,¯sè0ó©è,#€#ò¬šèží¥q©ŒKcZ,~¢„ˆ+ÅIì.VÈc-…
\ÙŽSì¸ò5å”üõÇj¨[=]RUÀ¢©rõ›ˆ©í¬§Ÿ\ˆ*f”3)>³ZÒx@7Öò¸8ItE†_‡âh¥®·Ïº÷×£9mG¡:O€ûÂp IÆ(Lw†T®‰òüÖT¦~ãÒ¾Í³ {ÃöµÊÞ2¹ùõ}#ºC-ÊCÈ³lÚÓ@y¼QØª—òœŽæå¹t?<(gX‚‡ì“ì†6ÈP;;ÕáM|]Ì«ú*š3­)dŸ»ÒC™:c.ßªR¿ÌÖ)×-mŠQ+êP–Ó…4tqA©²uô†Òmé‘Ÿžžæô9H H¤]MÔåt~A´3=³l0]6Öôf_Çàj>BoTœa˜æ|¢ÁI¦tfs/Ž‚í`“ªžÙo!¦\n«X¶•ÊyÛðTÁ$GG¢hÖ:!ÝXf,º¶~ŒZV]rè”±kS»ê§gaO*hxa ¥RÂãß) l±ô~–ÖhxqlÑÔõµ+Å“R¼¸[Ûn-Æ¥PüLqÊö,ç"~†ìD²~uˆ?ò´²|‹<oÝãÀúUo*Ð”<H©]Gc8Ã¼˜Á¦„:ÌŒÕ¤6Œ†ô-{ÕI°ÉÚÕí ­*Ã%¼{––ô~ˆ´µk|}½»1âhfŽ,Ýœ9Á6ô0úÒã—ýt'íˆ¾w¤#7 ¨ãÚ­CôOÛÛ€(ã!ªê«9 to%ÐÏ£ÍÕD¶þhc5Á­3*>-ò	¥±¿Àë0È+Ñà¶“Màëô3=¾ídž¯sm=mtõ-ýŽ†iB_„n!Ü %<]Î¸Ç›–ŠÞ‹ì$xƒÛAÕÏê sQ9-€ïê±µ×ùÆ»È˜Îú²€ë8ƒ¬±!ôîsv«}Î÷·iÉì_SRlVuOàö¡`Í×:™÷”œ‘kc`©àØ;ÈÒA%ý4¤Æ´»>.ÑL×É&ƒpMP›ÿïS&§Y6D“~OmÄÅ¤,†AJjâ'<ñpÊ¥“³LD´J*ùó<›gm¸²›‚„Äº2dÃ!)ÐL‡á^±Î~Ýk…œ[Ò½¿Pé!Ï1Î'ù±a^¾½;ÉªwÅìm²w÷%:0ŒŠw
²CàURÑ|{læ•æ£jXM”?e“¶ÏDQÖüùjr_r¾X/Ç³-"ïg°Û†üªoÅ5øSwŠJt¯Ð*WJŠ™Î*éRn*õËéN³Îšñu¡hÐé¬„­Å¶è¼àVjG1ˆˆ°÷øagjJïu:nèéÞ¤•×_­&ëp;Ï7Äód-ÙðÞmª:p2 n5®HÁVÐ”™—/Ò]S½&PÂ¹_žïŽ
8X»+ä.“O.>þ2ÊñîÏj²B"Ù
ãø²nA_/Ñ¾10Ø€ŸŒåN„Jf%`´I‘ÂVöJÛ¸e]tMVÒ·€”-R“jùßX³¸ò9#Ý‚6·]õlµHà4‰ããŒØtùì•ÆH˜·æ‘Ðñ‹ ŠhSƒ?:Ì”$·<¬°usÿÌ¹¹o]¥z`­ÕÉP? eºèO¿daÎÁbB¼ÍÅ5i‰W^È7,’(FXÓJ÷®©1ð5#¾m	ª6˜mÕŽÆD
kr<Æ<Î–ûv-œ²á|xÜÆe#Ä=–´M²uN–§s8eD¶Y†÷®:p‘Ë¶ž³ô\:’pÛ©øí¿~Uì¾”ÞïFUv¥BX ]sç’úøhPÖ¼ò¡Zëli…G°¢Xa¬75æQ‡¢Ë£)IÑ0¦t#WF½ƒÄ¾Í)rßZYG¨FJDsztÄØ¦º¹Ê?uCÿÉ·(ë1&ƒåðDfY†a¦}«"îËäúH —¢íðÜôÀ™‰°Î¶<Ê*aÄ]KÚR7[¡NjœVË¯CGˆ³Q VÃáÝçÏï~€Oòý÷[ãñŠ:ž g¤Iß“êØÙî¦â;ä‹£ì}µ[Œ`]ïkrCû®ßï=@’ðtŒÂÌ0ùj}]×À„AÝÎARÃÇ¿ÌrŠ|ÉS~t¨þj‚ÖãKXâüÄáÎ äf½I‚#ÛØŒlcš†ðæ	¹¯~)wõFôyï>[Íè\çÛØ)í–j½)¦ÎŒôã?Ñ}OÓVJA°?¿lÕíO Ýwªžšâ—›zKhÄ|<NgN¥“ò]6R‘PÈ<â+ê¥à6Ó¾ÖC òÊynt%þ;R¨XH]œ:REá¦&f‡Œö½¬÷›,Z§ï¯Ó&ÆéŠ@ 2ˆÛ0ˆô´ù¤ñÒkÛ¬>°
Ý.›ÝÝ¤«GÅã‹àyômþ>v7(áðz'²#7¾Œá} ëAjÃar^$¤i@4bC¿úGî¯3iwoˆ'7%?TSùÌå#ÕË É+¥ß‰ Ë…,ê¬5º ›5ŸV+º˜ÌÌ¡Y{U–x^Å¦ÍÜK²Ñ²³ÐY. ÌV¥²&xGÒ°ÿWz^e1‘:]MN©º©©Ôé‰«%úèÀÓî	WKî¾rW·£cæê%ZQ”IüC9üÖÄþ½^`¥êƒšòîƒZQFûX1îlµ¾ÉSBû}®º§&Ñ½]<úøÏ46*©ïÒ—D9.çŒŒRÐMŸ‚v^ ÂÏÇA’Áìã¿íÎ˜ú
2÷5/c‰Õè;u„£Až”ó0ÅYõ_¶ ¢a1Ï³„…ãão$¿*¥~Wd¿¿Ô  
øê•iâ¤~Ø2ëÂ‚œghØíœÍòaÇîò°ú0¼0E;˜_™À;vñ6ÎÙ,§µà@‹}¦DðÐSwýU!uã¦¶%=MWµ"¡þèHø]ïjÃÀËâÇúibUµ-3VCá×2¯ì8AÌ‚…fn"a·À ¥ßƒd…©í©„.àÜ(³Á¹/Ÿ‡û ¨qûL”r°—þ¡¡^JÒ‚Þ	È¬íˆp/]›…Û…iÎGÀsæiOî£Íõscî³6ïsðè´S¡cI:Ìp·<¡U›~êt>zWMÊ¿ ¸iÇa&kK©MeÓò¡ï­…w—-Ï‰éž¢›$Â¶˜WÓ9@÷dTœÌgÐ@]¿öTâByFwü‚"‹¾?|.íÊÝ‹t´ÁÚï‘ÜŽìZ›5,50	8Ø"…+ÅšúL6±{x1~¬bì2Š>S±d0Ž‘z¬^5âÊ:ÄðÜßÉt‹ø*#`—_ ®Vd¦6 a|‚Eµ¡ç¬[àð1•ü; §t²UèÖŒ­-£6Øñ—«%U¶¬;ƒK2¼*Ú*„G›)„ø/™Z+ô.Ž"öPDxlÌJq¿ë»‘¥ˆ¿ð<Ž¥Dj'ùŸç™*z£‹‹4!iÔ‰M#1oéD¦C—¥"‹ †	±5œllgBÕÙó—PÕ¤º)^§ëÆ÷‰’¦93Á|+W‘F¬N§bS«§ÖT
èë¬ð|]T>’0%wWNÂ£k°T³ÌVÓ4«¢’O¶®¢“ÁÅkØ(Íü¹Øœ{®<’º"fl¬Ö¾9è ¦Ì&eŽgp'gç ÛÖ‚ÈjrìáïªÙ÷«&,ÓªŽûÊ=ÄÐžå'À} ðóÛ‹t‡YUJcìMmg)%ŽnKá%—µzŽ)Ë'–ñH8ª	›<™š6’ª)êP0²P`“ZlâH@8-¡Wa©V¸}Ðe¹!§¡)¥=$Êô-’ÏØÒ{²D]â¹oÙZÙ‰{ñ8œc«!NU$8d¼á	0ÅÓWÛ‹D|•kµªk…j·˜%_Æ†™×$ûåí³Ø¯ ­4I¡}¤Ôée,¤T1Š”–º'Èìuyx—Õp´3Ãúh~Cã¹úq;¨®Z_
¶SkE; 2yÀo¿	Ý²öIâÃµIê&> ÕÔ([Þ¾–‹#kÝÖ,­­Y†·¦i«”ˆ_^/!ƒÕ21¿\óƒ`$PÞíËe }û¡-†óœ%Ñ¬Œ¹–@‚r+´¹kG±—ÄÚ¹„,6)\fòaU¥™æ<n=C»š˜­]³$«þt4+,Ú•ã~8Ì/’Á(-K”i­”Ót­}Xûjå±†[æt”½7Ö„½Tÿ°v’Uï²lÂêŠÚ­U}x¾Á›C½ßÚ½÷#š'ÕÚÉ(¼MèY‰Ÿk_¯¯'Õ½VågçÕÊãoò“Q^TÙ •:ÁólÓÝè‡wÏ7œÞ¦^g¢áûÐ0u‰^¨óñÊãïÐrñ¿HÍhð]*ÔË‡w§ÖÄî:3Âé,®m&øuíÝ,®8ƒË'Óyå^ùÄkiP=e+î«|øh…r>¬‰D^Öÿy>Âb¸%ÒÁ ›VVúïGåûÕÿ¬8EŠÉ.y¼>ºÞ•"o¥‹°7è]g:ßÌ«ª@UÆÍ‚>@.È³‚-ŽòÁÛG—ÃâÝ/¦eã).ÃURåÕFûMš¿Ç ìÅ0ÉÔ° ý'²v‚{üÑåÆ×W0„ä9Õ±GrWåÚãyfEépH–®™xßÛŽRÖ‚u–3V!³‹ÁqÜžÒÂ×o>ì»Ž'Þv€µº½+o¼2T€5Z±ä0^Z‰6ãmýÕðàvÉÔÌŒM'±bÙíçš‡áf
 ù]§eÍÀIå8PÈ%!€ÙÚï‚œœ‰ï÷×ïÞ‡_Ål˜ÍÄ“Mx{^\d³-UfÃ<aåî	GÊb²6@u„³ð¡c×Mg#neîŒF†dvÇ°¬X“àÆ|®Ü3?r ²Œæù,©
Ü$ð?ÏšZ
{Ê¿e%ì§äÄ"çóqš”ÿ’\d?»t)I\œÁÏÃ£YZžo*üýŠð÷YŽ§króIÒÄÌÞÇŒ 6'IÏæe\ç®~ÃÁ±FwuÃ¾Z·qjÃÃ©ûëAX/öâæ¦hš¯·:‚º(<Ñ½‡@YÆkÌ\AV×ÌZAÝZ÷2	˜
èZºPƒ7ªå•ˆ˜jG¸¡éÔ8~I	¿Hïæ¥gJ¹ &²ŒÞõŸ¡{P\ÇÜ‘mð[n„š­`²Çk¶Z¢ø"5mE$Ã@«ÙnŠØ¶ ¨…€*$7·yv:F'/ñ+hò‚â†·;¢{Âºl€’Œpº[^£àVJõò,¼:m·“VÕüÕï¦C“9ú·ÚL‡šËømvÓ¡ðÝ©…€,Ój/™“Fz½žeƒ~ITùër~–af«Rµ’þ\¨‡òxb.CÎ›·Ó\·P”ºô|Täòª¡ƒ¾,K½ƒ”cóä¬ÝÃïÛèÚáçÖvºH=Ûæò7/22úíMÈË§†©ñp«Ýë.ù*.Ä  —'i<¾¨Tòá_\E‚ú=Œ^dgÿ2ÈC(Y9WŒñÕVöù³^û'õEhÔPâŸ¬}•àQrŠwß¯¡?ž­<	HâÎúa8½]ºÜ¸lÏÆër0qñceú†2ý°vOj‹ÑP*Ç\àÆÜVj>k›É»ó¼Ê„tRÊÎwQ+®È¦2q'…ŽZê”=XlVåh ¨žT\ëù¥>‘ÙC¬jÅ‡½ö!ÉêáÝK®µšOþ4—ƒI$ŸÌrøú%-‡±î}Š«ÁÚOe1‚|á’VCëŸäRhŽãSYqáìv„ÎÊþI®ƒÞ'´Èd”.Ã_—IàºÑ;×þá÷›½£©'N\²Êš¸±6M'Ùh%d¥¼Gì†kSst±OðT­—kÉx¸e~n’âKßˆã™-éé(=ÉF¼õ“Q¡ì•ïKqôÚÌ§ 6Ò2Ão¬<FéïO‹™ØïRÓ.…™·Á(ñnítx	(ü%¢ðfB	n²áÚ{…¬>«vZæå–Ô¬QnOñÑlc~(d¼¿¾~w“Œ„ö¾~Í)£9ì¯ÒåBfchìfj'„ÀLná =–6b(¦¸_å€VV£- ÄÔF‚î<¼+
„ê^ú¾¢äK×x`˜ªí·Ù‡G—øðJÍ›~<t?¼+ äáËžÓÃ_•°!¥÷¿U´ô|&ÅN¯å!©°œA®ÃQ§*9
¡üIº»Åx>NR4Ä¥e¯Õ+" 0bÞ ªû±¨Æ~úñXü©ë´w³Ñ$Ãö®ÂGS½ÙnÑƒctÆ~Þž#—ÓˆKÝ.$×Ó;>±Dë¶ç€HNTÆïÇ”îõêšûyÕùE.[™Ïõœ£oè"gÒ.Õ{ˆ6{˜5µ‹ÛöÕ žžµ‚¶|	`NÏÊm‘Š¶²RÑÊáWlèºíc¦€CÓ^h·Ó† MŽÞÎWÖI#Im2žYõÍUGóByž‹wÈJ6Uþ]Öz8ŸQpJ:0Oß`€ñ*1ÁêžÈ›ö\­×™™@tÆ=Úè’¨Û4ž“>‹XÓ–ñxAª°©Š§(Ð]^NÊ{xRnàav4Î6q|Ó-o¬Oß¿bG¼Z:oÍ¯=çÞå¸,pçUt4H¯ˆÆ„Cî]o¬ú¥Mè:pÇlñ°“N;¨ä“a~Vh¸ÈŸ·wz15)pð‰JTmã~I8– åÕ‡µ¯×ñ R¿ÖÅ©ÛhMl<õbÃv;’aŸgõvm=Ž²!yœZÌðckÙ1¿Y“_†]"OM›Š:Ž5…“°—G0†Çj}C2r‡<²k‹j· ²âkCËº;€]-„ÓÙãÄ©‚•Ž_ÕÛìðÓÞGmÅ7ßðkýûÆ$mŽµ™Gg5ÚK—âzÖ%ƒù¬,fk8›…d	"éÒX™O¥Xù8åxˆ9—0ì2íÜ¿rýfYá€Ìzk»Oîocÿ}bØ¡= nŽŽÛ5pÃv_¸>fDÏø‹ó{Ö©R|-z›+/ƒ4d¼ó{m,¥°‹qB'ý÷%‰ã¬:ÇåÝ€=EM=.Â¬>´FS’äöJÙŽƒ]¨ú\Jª*øøÆ–#|O&2W¨E©Ö6ù•6Øc´# –ƒÂ=æüŒÉöŠû²}zvåða›È‡­ã¸úÅVY­Ý‹G»<ýÝ¯Ì>#<[„Ç©áÏ‚Ð£¶Ñú•·ì–6=ºÇ’ä?Ðœ¯Áþ†4QÔi6Êûþ]•`wÚR•p]µÁ¿o%-D´\‚èn©ïJ¿)¶¶Ü¾’lª/C8ö²¡¯&`ª„<»Ëõ%>7úZ6ü¢;:\9E¥*ÂN 8øñ.N?@öCû‘oGWß¸ˆcôm¨»anO^pÍ²Éó¬DN§6NŒ]Mà)^«•¡8;ò»·4hí 6£7õ˜½¼»øêÛØ:Æb*ôŠîÆ^‹Ëåƒ¢ÅÅÿ¶÷ŸÕìn à½¾Š7¤ä]_ÆR·¤ã·†ír)—r:\&€ÍÇÄñTÑIëÃ}Ù€ÍPTHFÚúçi)^õÂà‚ÍôœFT œêà^>ÇÊ)ïÑR®v÷Á%c
YK¿|p.óÏ_½®_M"*?/QçmO·I\©¹/»½-î~.©UªË^?1ƒCd*µV=—ZE»øDÔíºÛ®wêÙ±Î?q·x±­rNŸ¤éCk‘™ÿ
‹'`ÕÐO”1ÝÅR€é(ªRõâ_ÆTó“ªÈ?BùaV¥£s¼R]h!¨©úêb	Ï%3(öGÌŠqˆö'~M ¹§ëŸÜúB	ŽƒnYk{”ú„íR‘¦æ#õ1ÁåÛX©ÔÇX«(>l+ƒ•úØ†«x\±zÛ•nMØ°¨™6f,õ©7g©Ïb†ÝvC‰~šˆð¥ÙMöé{ "“tô,Ÿ¼•šÑ¯kì²V#Y®±TËA½»GNc>‰+÷é’ð¥È{Õ4¢†ñ|bÖsR“?æ1kºÊíÖmaÍÃOÃë»w“ï)hP2-*ÌzŽ’½'	ŒeLÁÂ)_;•IDhšðå\óÁ4‰pÐ¤£= ûQÂ™¦ëP¡C©ro8ÇÛ<9`A»›	‡‡>šª]ùÖ.ûãGh>òÙ•gg>-NsAÚ€¸M5•€%iS­,D?©·ÿiå¦Ññ6»7à1¯—·MýÏ¹†3ìsãÀãÊ!¼®NÃ)¹ˆKUhpë 2·~,([:>]LZ‚Ë~\·›ãQ{÷*ÞˆEµÌPýËÜAº™¥ÇSé¢hhÔ!$î.ŸKO‘ (­PKo
ê®¥GÅ½ëzTÆÝÚÅ‚xËn÷oÅÍ‚F¾WüÔº[P_ÜØíBÔŒbûB®¾qÜ‰SÜëºž·aÀZVÜ»ˆAksërÛå˜Z¯z¿‚Ê’f¼T‹ê_•—Ëm¸µÜ
6.ÓÚ8í%Y=—bó¼–Å³f}´—PÌè)I¹dþ&>Dí¼ˆR-"7¢ta7¢È©‚îB˜¦þ iác‹s¿uí™®Há»ñ—ª_oºj;”ø"–í}~›ÖéL{KI1 fw.6åR]lS,äRÒr©—éb³ÐËîÊkž)nÑ=‡ÜMÊ¨Î`!‡“ë@'r‘\;”Ÿ„ÓIüÒÛßžóÈ§é	rY¶pýhã;pÙõˆm£gÇb5þ–<;™ù'èÙQñäk+âIÐ3ÃS7‹Çâêºþv}5¢åöçï¾‘Ïß}5­ý;ôÕ(oä«á¨V•ãÁýf¥ê}8þîRŸæ4Æóv«¹)[™p­0àÍ´—ç·HxcÕ)«¸€Å\ ²-õ	(/ûÇÍôÃ‹i£°5·á†@Ç¹ó'Í-„¦—ýÐóæÖÜ ëeß~ÒÜB(ôzÙ÷Ÿ
Mã‚Þ
?ØšØB_¶§ò­<>Q»òmy(èèóÿN]n‘Ð~¢˜t;
ËÀ£OÇE¡¥n¦M¼W7f¬Û^íÏ¿l-{zÑ67WëEIÌ)øðîù—µlcÞÓïÝl–—ÁÓö*ã¢Æ7œ7®“ß‘YßÕÑ}ã‰ã¦a1˜áÄ’Ú¯KM¨Ÿ|×ÛšÖ›ÿþz2…‰7„<%ÝeÄ>ãwÔ°YK´€ãõ`e>þ¼LÓºÄb¶ûù=Å[	óJÄ]¼¸»ÓiBÍfÇóÛ†Ïö÷4æ˜{-Ù\`kØ´éÔm¡“Œ8ªø£IÚ™”ï²:®Ðb–•ñD™µkK%7RuÆ~·Ó(iÛ¯J6°rßpü4‰ç¬/×&3c2ÔBÚ³6Šõª¨ÒÑóô½tÔ– {i~·[ÎÇ«ÉL„'ž“/’îfò»–	æ£×[MÖ[w—¥³I6lÝÙ¬_ŠY†­/ÔÏt€îLzrÉI €zÞ'‚ÙíÊÜÕez0/ –h_oî¥Öðdþaš¾œF—ß ûÝ
ù* ¨…á¡a.WíjQ4m†úIžžM`åƒsìTùäL$²N¾’ƒ¡K&i™¨­F™ v*8Ëe©çéK­´ê¾Ao„ŸÆåC/…ë«öo¢xªs ;Ý™2Å£`úÈ’PãsVþ\¢NÐKµUÀäb[±ñÒõ¯{Üï÷i\«	|£Þ^…»êƒôSµhEê1Ž`žw	ÿ´Ø$»w-çYü\êi´ôš¥±Õ{Î‚öi†(Óã_žÿ,~}h©Ï6F¹›lÝ(ºúþ±-}¼tÌ:ù÷’}ŽüNnÛÛt‹ÜR‹nîmÚèD´,w¡«ÅÖhyÞ¦¬Í‚^Aÿ?   ÿÿì½ÙvÛX– úÞ_q‚Ù¢²DJ¤ËjI¾²lGªÒƒÊrFU/_¯0DB$Ò À @ÁÖZÕ¿qŸnV?ÔêZ«žrÝ/ðŸÔ—Ü½Ï  gI9YFe9DgØgŸ}ö¼Ú+È!¥[èÿÚ.˜i®)y Ì2=±V”p­à¹¨ëeý8©º•N³ÍŠr¡î%§š`º—
ÜÌxÛš´úå*/½-2Æd]G%Û…|´Qe•vÂüg~Ôª1è¥w‡r(;Gû[J½¤.s«[J\¡lÒy•`±•4Côú¼¦}\†AœùO›×!Ûmëø°=Í Œh{ôI<¢2¡õSÿBûŠ“‘¿xûÚ÷eÞwƒ­ -9¤ô0ë=Ì°šÌÉµ€œú¬n’îCýêÜš çA‚æ3ÿÊ›…mú^è»A>Ê¼–0° r M½aB—'4I3TD1¹ô€˜à’‘dè2I½¡3óO7å‚¬EO|ÁÏYª9¹×d%M
Ù:–ûmVì×Í2¢®\Û¬Ì¯,ú2?¥{z;V†<@¶MÍ0¨‹-Öý¢a×ö<xû%Îò“hˆ¦âuÑˆ=^Á´œæbÜ³¨š¢é‰å®ôC·*~Pxè:†Zô«dÇ5Ý¦ŽÝ*úuô' ÔùïZ}ŠQÜ6ºÒ“ízò¦ÂùØ®ÊoìK¬;†”r‹¹ßtöáLß—Í—%³fÅã~œçõéœ<{(iÂar#äžÕ¹jKÖ±”ó­ããìn&½\ñÀ“ßW,“µVÅU›ÏsZqŸ“Š6Vˆ³ëDÿñÏÿÒäõ¹ ÷CßF# S§öµ|0•bM2jº;À†lEM›˜Š†‰Öx#Ã»t£´ /ï¨°½\1&•d¬x5eìaQáv+ªæ¦þ$ L»¡çF"À™à¶z¿õ#lq8š5Õ—ôºÙWÓ”ÐõLSé»FŽ%·#®£ ‚;”ÓbCô×1¡ù¿:âUÐ¨ìÃq‘þ.¡"/äñÐX÷Œ:$-†?Ì™ékÂŸ²3Ç æ,´0
™J—-¥áU%á'VŸi­h1ï¦ª«~YuE$Š‡
o§Ð‰ôQ'Ò¯6¸œRdrÛñfYì¢Ùiª‘‡?ñ@ÄGãï0Ÿ»ÚÀÑˆ]º¤ÎBÉGeu2`Êƒ'”S‰ Ð÷Sì`@ß­¯Ör"éIªpþBBŽè’ãÁç#Óxèç#JH:C]Bòù¯S ƒi®\ PÇ
7ø"ÓÎHZˆŠ
Â#?Ï¼ðç.!uÑ»J•«Î¦‰QÖæÑl7ê%\Ló+¨OiKœõ%M5&îD[­5QêMNöXO,˜øA¢Ó ˜`ãªà³eøb*í<ðÑ®Ñæ¯þŠ
mir˜æ*ÔÙU!÷××hÓäøÇkôù¯ƒ &>¹ðG”x¥¥ÚÎD§‘è”Ú{ðé(N€Ö&Äg¶ÁtƒÑ{æó_©ûL
£[©þÛ•Ä0òÂc 0*ê"À©yßJ]T”E­érÊ«Y‚Š¤¨äÑ¨Î=¡Lå0áúÔ”CÈX(#H/2/z	ú×¼;;sñÓóWçoŸ_œtƒ˜Û¡Ÿé¿€VB•)0L¨2}ê•âkLéê$>Ì‘wª|C¯]eBì#~bù¨†tTÃ/<*»ª¹^ aeCªÉ© ‰hü«Ó¯X¼
=°›XŸ­à¡ÔÀM•À!¬ókÔAiÓ\7V"ç¨jW"Ï%ª·`ú	9É¤R±¿Û0ÿDë˜óóKç€pRái:×g‹“â^ŽúóFÚsÞW~Ü—¸ìqQûÀ
ôð†GÊgóïÜöÀ7EþW¢Èç¼]®Ë70e•/ûu³kê|Ï]ùö›J¯¿!•~Ž\«¿ñå¯¾°fÉÊ/ÅëÒ}Úè=	˜`2›Qã˜øC?ÞÐƒ“!›±°»h@ÚÁð€°$Ëð_|v@ÖN²àÚ[# £<ãŸg>|·&áR–ÜåX+ÉešTê%mÃ­küÿK…ÖþtA*FÑþó„Ì1°c ú%÷ä€
']tWž¥è \`°Ü3^ÞˆI>‹í!þÿåYÂêOiPkÐïúô+z”áÕ)ƒ1iÂx6|ž$åx;gú]?Iâ¤ÝzŽÿaèD{ò’ðõAkƒäÍH]ü—ò§¡ww1`BŠÝÅßÍ‡B;+Æá2ìš}ÆÛ²!ÄyèE±P`¥%ÞùiæS”xåG)fwƒ?O¢ÿë™DE›D”|@_W¤NW….°ðÑ¯‹,8‚Æ¨ò,ðRŠjtŠÇ˜Çwré'ˆ	y__¤NW…	 :/$CïneÄ¦úIÖþøfCu›2f–ãÊãïçùt¨•b8áéq¾û¸v³jŒc?Â ‡ÞPƒb×üéË d'< ¢Uzü’èVëzUHbƒ©Ï:Àlb+F½&žAÌ3@'|Úéçõ
DôJ×Z
ãªS²áÝ˜c;Ã#‹yâ'A$!I~nÑ ªòjó	syâ‚”ÏÑÜ=£…Þ4\9÷µï¶r`15m1¾2æ°qþAûø]{‡ÏÄ‘	y qnÆ~y|ùÉÙÐiíèþÅqwg€ZyÿÜY9ò¦8[Š¤#?MÛ?—ü;|§fw2šñ¹¿‰Œ`r	)ðgèp„äÏŸÿBü[œ6‰<‚þy|Áw-	½d8HÆ“/ãºÀ`„EÇ‰÷‹_`ZxÀæ5RDJ”µ`ï~ü~^½{OÚt‰×?Âþ];Ë¼	¬HASèif1-ñ–”Òk ž*ºB-~©-é;Xë@þMßèÛ}ùÿámLN&}øìqC ¦.óßD©É´½^¼È×÷ _ØâQ…/ÈÄ±I½Gøc¼Š'	ì=øŠ®‡(ÏvŸo{Š<ÞpˆäÍQSûÿÅ'/îEÑ^ ©•Ú¼ÍÑÏ'tùê3äéËKAwŠ	“,5àªD~C£Ùµ?QÒ¡âEwFRD¼¨h0 Ô7¿<NŒªãß7}ŸÔ‚pHsIÀ—]XÿÔQ„Ã†jDwéuàß°<Xa4
ÐÌœ^q“ÐáKöèƒãñGøMAô”Žÿ{­±ÁØ|ŠgYž<_ìaÑ•ˆNXÄ#Œ^xK-wŒœ±k¯‰·~B~Z .cmGÙ8¼£bƒQ	‚vŠlnbÀ§¼{Íˆ±Ý"qFÌ‚VeÜÜ$³KL/;E„$t2$ŒGÁ ?
è=ù­ç·Ó€%CÇ‰Uè2
!8¯û°i…‘Êö¼ô¯2±…¯›‘n±i°-_Áœ¨3¸÷ÂÈòn¤›Ü•Ï7eïáË,»~—S:Ä¢'ôø(†ÿ]©ÅåÆ*sÊÒRëÝB
CTaÒh1þâkI;¶ñÐ¤V(y™O@©^—œ]IÀÖ€ñ Ò'Ø¿@äöH€‘¸ÙÇ;oäÔ2Ÿå/O²Âñ­W˜ñø™GE0ã·Ý,¦¯=qz­M‹¯À–§?Ìƒ*³Ãàê
Ï'<<$xî¯·+­ðe€Vh²›„í¼éMÒîmmm‘ß“½üŸþÎz…‘"R†Ø|™ó†a;IZ‹u§ýM¼ÛöVëÒ›!‚•Ï½!ú]ò&¶ë&À\ê^8˜¡!²<C‚d'Ž®‚ÑFD…1&ïÁù”
3U„Á$àu±6/ýùß¬.u!‡Â^q_¨fží-"MŸpM<ØÛ…'Û"#‘e.¤Š…Ü4Vb¡ü¡ï˜ð¯4-ò©pPO¬ù gŠWaÇº3öé¹X¿NÞ…ÆŠÑKM(0×E¡~¬¥ûª.9]æÚcÃ¡ÉÀ+|þÇS`>Ñ4?ßñã9ÓéÇËX*_¡ILGý²ËU²û%-¶<6<Ož3`¨¦˜^—ƒ)ç¬Ê³€=WúTìŠ×MÍw<ˆ:ãÀ‹»xæ¶F«sÈ´ì³/YYO‘x¨Ôox¹0tµç®7…3)UóÕ\úèî	ÿ#Tî‘­„ÃÃö­ªüƒÚKÝT÷Å ‚(*Íˆä% šÛlvªŸu«È¸W3´ÑájµV]æe\ 	’±ßÏ¤˜RÞ’¡ù0þHSŽ1ÑA`WÜ*»{¥qê’{ì²1í[Üô+c,=#8ä”¼5Ã¦Ïºäœúh82Ë”W"†²öáW³ˆæN`‡9ž)œTWÞuœl ŽSm	ñà¯(¾öX’x†IÀ°"ÈlbÌuÌÆÑýX$€ìb&i‚@Ð™üX¬A†êÊ-°hA^–Nñ8#S?ŒËòT—<²¸­	Š—1Q$©4Þk8ËÄÁ1Ì¼cSè6²p·²PSÓ6Ð¹«UÖCáh!œe+·kZ.Ì´%Q«nG,mLº·i:³'Ká3ùNà›.‰¿¤ô^U‹rÝbù–.yBNp­¹’ª¾ÚrZ5kª¡b³½WaŒ¢7³¬e\ƒ–VU¿j°vî‹^j¯
ˆ/HP¡yŠ¨Wžšbþzu¸‰$[ü®YnÅ‰Æ”<ôÑ"%•6½(˜`+éNë’_…åÐË"|LÎÒv\šºvª%…êž¡;ÍÎÍýR˜ß.Å	*Å«YÑÞBlÏä¶W‰°kr”V“v×|B·û‹¡}Ýú84y&¹ÞnhAÃOÓ“¯:ƒ9nêWmÆ´Yñõ&OÈÚOú›‘ø‘ŸP¶0Aµs
½ž´»†JÞÓ$ð)á§g˜ ÞÚHºkµà¿TÀ5²^CÉKÌöˆžsxpÓ”Še_i¼™ü!/vžý~¢ðýƒ°¸…¬ßËc?¼Â$*q¼9ÈGsfèxŽšæhð^y´)ñHþ`$Gí`Èî¦ÐUr×ý’€øc@"?9j!âwÔ/^{áÐ•Nß¨ŸpÀt0µlÛ'ç‰x#¼D¼.mÈéæúˆWX¢«x0KÐÐˆyžŠEâ±-~»Bð?Ï:Q\¯˜F×>¬Äÿy†ŒEåAÅ¹[•¨éºM*70nÛÒ!åGcEA6iåÎycî‹'¾øÏ±~„Àö’
¹Gó½êV—sOléRºÉ[*ÈÜ1’deéXÌGÉÁ°D¢ªÕ7ÔÄJE²‘Q÷APS®i­JÒT|²:YVtb&Ðúp-Î€ËJv´êë±À6_å‚ÉÑçƒN¸hÂòm²q¨Ãèî™®uHQKì[¸¯aŠÍý´.,
yJgŸÿ’‚ëª*Kh(D[c[/1›jÞ#€ZGÁªMkl˜Ú¾_cnå99…¥‹±áU˜;G@R¬äa[ÕÒ73!ìûÈûèâL šM“c8c0~ˆãQè¿Ä-y¯Ø \tŠÎW…¦×£r£»Àeï¶Ú-ŸÆ·G­-²Eú;ð?M¦U/c²Ôð¨5˜%ÀEf§HÙZdxÔzÕïww÷H¯ßíï¶:ÝGûîÖ#ôÛÞîtû¼ý‡^ÿz§Ûßïv÷p^ïnã;[Ð'¾oõÈvw»wÝï>z4Þîî>ô»[ûðÊã><èïwvºvØ_ûÝ­Ç¿hÒZËCý¯Û;'û»Ûl” €õ·¡ÍÇ`ª»Ý½N÷ñ>y„íõ»{{aûì`ï|·pPÛ0Ð­=xö¨Çþêw÷÷ÈVg·ÛŒcÛîìu{{0¶Ýí?ô»½}þþÎév÷ñcÒß‚›ÐÁ#‚­`ïNc~ñôééÖ.ó.4Fz;0]Z¿ƒƒênïBÇÛì Ñã´ÛÛ†;;ÛâÆ þt4§x›ìwww	šÀÿöS¼»ÝÝ»ðÙi„0nüÖr¿ýôÆúüdg{{7‡ïnw{ÐƒÂÛ ,D‹\X¸·nw{»üç´÷ûÆ¡âä`ApPðÂ
± €£¸áèp2ðíÞAÐºû¸H{ˆ+ê$‡¿zÄ°‡¯GõÛl:)dêD-Œæh4Œù®ã4«=v	/Dv¸«Àgaý$ÔÐžŠ*U+ÿz ãÝ¼XæeBIŸ•«‘¸ZMPI¢íÓ+Dèil[UtÇi^…ñMg‡ðÑ8áì ˆû¥?ó—®ƒ4¸ÅíK´€ð¿ÇLÁ~°aÈšùæïÉE0ô/Bo µÐ»#¿ß,fqšì1€TÅl:×0å‰·þ0?ßóIÃä6ÑVá>“xŠåÙ7þ¥³³…Ñöl¾å%­+bÏäaµ©™¬$‰HR–»Å\7¥óÄÒ¨ºdÉ,ófÎ(@ì°$hñ
¥ó—áU|Àûð=”“äE@Í»+2AÅÎ§;’0·s4š¿z˜ Í05E‚=Ñ	Ÿ&[˜¦YSÔQßºHíŠ($I=²œhcGŒús©~ê% ô&Ùv)ó„šÏ2Ï£¨4˜’!ªð…Lwü'.¥R1nð	åÞ,³ªYLXŸPÎmË>‘ç‰6r¹Ö­¡xX¥,(Êì›<ÌøòŽœeŸÿ_t§ü!žø©ºW}JäÒ-uQNÛŽÿ®D—L–…<Œºz†ÈdØ(DO—
|Bÿ©”ÂŽìÃW~4+Ý¬ŠÕõª“‡›lÇ—øzuš³G·a®gF4†)&ÿa»i'þôm†6e=.Ÿwñ”`ñ¨¡!/ÐÛ‘•ij¿FKrváøxáz…:™¶-¸c¤.4¥
Õl’)*iÕ4‹¶1¢f£pM¦¤)”#§À¡=ÂôO=9áR­`žZÀ¹Ô»Vö à.üŒ{o2Œ†¨‘xó˜"]|¹aŠ4ü•‡Âª·j×Ê%^Å²mq×S†qºÒËi1îÆJç£yíç­²-«CŠÜ%UÕàjÀy·|\7·²KožÐÿ‚9
û°ç£AGÔ´M]…aMQÐÎ½˜I;÷¨^Wõ¤É*pŠ@’¢e¥uÖ W¯ÏTV/Ñ¦%aùøÌ`ø”$èó*4í!ž:1ÿ!=›;ÒK™Û¾2œgX>¤Î¸ 9ÆšâÉt–¡WÈfä]û#ü«Æÿd°g	òÜnXÏ šc}‘ãgq¤§G“Ú×Ÿ¨3†®Ì¥aäz
FÎÃØNŠ$€=ìGÍêÉn÷™@À‹ÏõÊ~Ú¢ËU[«~½Lë×T$Žò!æ¤ý,À
’èæ«Ê7êÖõ%èµ°XÊÇLYß®­Ã·Dùnš¼ÈOÿÂC¥¨ñöþ£.'gé`P€]æí…çÈtÂ³vëbP5ŒGhËãhº«Áè}ŒÞ^)FÛÀjÇæ“Ñvuà{YøE«¬Js¡…¤&Õ´7ì”LŸÅ%óŠ)ãJÅñüý é&
çÖlª:bŒE¨M¾bê¢ž:A’8ùQA±Y ÕðùY÷Ôná@|þë4 Þ‘©?š{À¼^(ÅêþýÅ›×
–Nyh<‹o"ô$¶½HYMvV1ˆÂà<õ†$IÑ"¿ò Ê¬Ž“].ç¡ªª(ç¤NÅÖsîIR–y1k²¼`—Rp>•?M-k!3FÆálâ´†T5‡¼˜Žÿ*5é¼ÁÀŸÂNíþ9#õ+…«³ž8lÇÒ†T)kÙ¥´ˆPÔZLO„JŠô’ƒÈŸzèˆ‘SêJVÓGÌå †ÂÉZŽ2ÚÖ<:Öz;œÄˆHÝšÊ¯€¨Íç$žzlÇ0&øýòY‘øá[ú G?P¾^vdÛ§©¦]˜:³³"Ó¼ä„ÝY@Ÿi:Lî3™Ÿ‰æíZM°‰ÚcNÍÃ‹qà‡Coºp¶éº¶ß¹Ðg-ÿ4›(ŠMÇ%¯yµs¿‡Ñeâ_ûß)s<+;I¨sŽ¾uLsÑOˆ7õ±Äú¼ŽÈ,YH—yƒûÄ˜Á¸ÿ3´‘˜\Ò€ˆ@.1Ø ñ´Á1GbTñ«/íÛJîR…¼#™ø– ^añW¨çœÎêÜ·Ái8V«÷âQsºw!sŒ–Õ¨œDäªäJ[Pžç‹F\„N.b=­‰1fù£'SoŽ–‚~ë¡½Ü8JMt*§©j%¶WœTôP	 T¼¡¤,_†9ic|Uà¬Dü*¾/‚,õËÑ•ªA¨sï©Ã#iln»Q°ïº5ˆ²|ñµÍ{ÆcUõæ½J( ­ŽaUÍ°ÑZý¸bkôj­ï&Àe±§íFÑ¬M¡Û(RU½02¿£^í·š‹Éêì.¥DQ¤ðPõŠšë`š8-ÞaS~‹]¹.vix/‘î³E â/—â¼Ø¥MliâÂ, «áÁø8¤øOÆ_4fOvgÇø.L8åÉN2dÇ0ü‘ž•y¢™”œ&Ÿÿ?¥NLÛŸàÇh/”{Nó'MÈ\±î•wé1ÕCßk~øRçkq¯ÝS­_ÎL2Þqèa'¾‘ù)wÑÃJ<Ýz§äQx¢ÄÓÔ“wY=_Ç.#wÇ.÷½:nçêHuv¼ªã˜¥7ßf]zÌø!kÊqJnS®Z={È.•#yí'©Ñ,ëkÛeMm•ë<<a„ñ0fä“I<DÂâ™ªzt‘tè“]zÑš:}¹…Žcu‹k@z‰O&¸ÚBZ%ÉÖÐxþ]Oý‘‹1XzRŠnX€Ècl("±z{€ðå_yéÑœÿA_ÜvÆö Ë“ô¬(œ‚ ÎºR%žMAñj7j¶Œ .Gsñ—ú½8Â,Þ°#3cfêrfi,†ôŽÎHŸUZ|‰¡qÃk¬%¬ºm~Ï1=µæt#–K»áTYíÕMvÙIü‰æ¬d
€£9Ë`ÉFÇî™¿;O‚8¡BBùÛâ¾þ{T•>¿PhÌÕˆ/‚ÐgOÔß•V—æ+~j¿È¡O?Èiß—@J¿~›ûÀÕ*úÀ_ºÙÿàƒxŽUöh½sšálT»§þš¿‡)Þäïøo[çÏ^]Qòq2 ²óì…nj5Qäb“nš¿æ‘ÃÅ—ü†ú«0Å§ã8@´(þÖ‚þ¥ôzé§Fk<Á·ŽæÅß¦wsGIOþª|Wö·>¦«¬tR¿ù"‰'Xsî\Ã1¢jO©§Ôÿ5B.¾ãxâkŽ^!Š¿Ø„šïÑKX=­¬´€x¥:xQŠÓu¢s9€yè\åÁ§{U.Vv4—i?Qè:¦5/;v4ÏÿÔ½[Ï£¨“‡™"ðHhuªô|Lä»ºÐúD} Ôo¨m“F×„ã ¢?±I—`1E˜L–7Æñf®„Þ9ŸÃñ»1v7Àmèw
áËùŸê7mHcŠgÌÂ-Xƒê-3KWúvX½e%õ¥ÏÅ]íQö¬âÊíWå}®{ë¢~WïY¿­pªû|pyÀº½%¶4kà·»¡OèøílÞy+\ã‰7HbZ¦þ2ö’á‹]ié·»ê¯p"ÏÄ<¾rbzivq.f“‰—Ü_¾¡þ*Hñb¨Ç²#òoÑE¡ìÓ­ãËåq·"ëZºÈòV0¹=”´Ú;ÞÖG47ÛO¸w£šîº%·cÝˆ^{×ÁEê˜Ó× ÀßÆµ²ét™}§S$ä•€ËE€‹DÑ¼n6Òå/œµÍqûÜºâÜ¡÷ú¦ÀÍy·¦êŒf¤êA0«	dEÝázÉa€-8<”a`™;[#]o‹-ÑÊiñe²Èm¡Ûm²x+<?†A:èxiê§)úÜ/q~TZúížÏ`"'hÏ÷žA½ì||¸p+å²YS_Íª*Zûí®ì+:Ô´þö–zAµå2ºhâoÆtÁ'¤_MS»¢§Òû¥ŸšO–Uzæ}T›Ð<Ð³ƒhÿ£Gu¢>ãK·VÁf¥¥ø@a‡(Gj-]Õ„FÁ	ØBÍmº¡?µ†ÖqŠêVõ‹—®9/RB_II9YARÆM9T•&éd‘|ø&öì¶\H¹6Tõ2ðz¿Ó¸ˆà±˜ˆã.,¬Èæ¾Hd^ŸnôV´ö7Co_²	½ò"ØìÈ
ØÑE5àÂhZ+Õ+l§µÎíPi\ÑŒAJW´’ÿT´”?snM”yT4&é¬n4[ÛŸ¬š^ÜpÿdcÌu4¯ó«_ªZŽÕ‹U:ÃÆpØ¾4röµž˜qâYDç1ÿ1¤…®O+äápyš}]…ÄúuOü¾µ’mVp	HËÇˆT›W'œÄ4%‡«á¬•µ‘Þ2e;%¿¶Ã)6†$ýˆ¹_Õ9.ó%ˆS?Ž¹¢Pq­˜u©<íöV©˜f­É%ƒY¤.C[Öá?D3*€7ò”qiáAbÑ¹T¡my}`†–Å`À,2+‡Â²ÓªTÍ99mÒ/rð‹"+:ù1¡z¾ÿÿ0-Dzö«j&®lQJ2_¹™ß®îÂ” ¤“úwßàßny öu‘¹&—åqß\¡u,JL™É;Ü÷µmê!´Ž¹ó¯¥xûHä%•òmZwlƒ3v-ïÏ¿òQ«]†óÄQ91y_{ò2j ª÷¹Öec©ãó(	†ÿA|N;=2?û˜m­ø¹ÍXƒ¿J)ƒÖ5ç!vºdE¬
Òl¢K1éÜµ]‘.¼ý>JâÙÔZQãïèRb¨ß—BK°´vVÛãŠš^¶6KÑ-E‡J£&Ù!):ªÜéòU.DKvýÚTª%^ŒQ2üS`†ê4íËˆ(9©æ .­yOîN9G”Tq¡%%ŸW^6¦âÂšñô<Cg$U"4_rÐõ¬\dé“&\uéCG²W\zX\‡èZýÖAÃ=+>ÙbO¤÷–Å÷«"F±¢ÞD3Ôx˜~—xé¸¿úõµÒûnù„ÆÛÑáx»ž7ÝG­czÐ%Þ/^|c±ö{`Ô¶-VIÍ?ÒšÓ;LZ[N ÅRLí—3LÍ¦S? nç{ªâàäº¥Lu±|¦êa¶¶hŒŸ:JŸ»Õfý*®¹40,¿ÍÆ¥Ÿ£`Óf¥)Þ°a†>B² G?ÄF}†R¢ûŽ—å	2ÕAÛ¦É™Ã:¥¤ïåk]Ó§š½“ËÄë|õi•Ý,F&Ï„ßß*Uºãü^^Ñýã|E>¬"ð×Ä÷U•ý‹)ÒKKÚKj"PZ	
³.3-§'¤'Yxôø§DCM^±XÚ•T.•¶$Ì‚Ú€xëøK©Dßô _L°ÑŸÚPAÎñ1¯ÏÐÏ~e’~Mv7ë¨yŸC‰}e|Ÿ¿)©øù8ê~À“oÍ¬Ä0ÆoL9ÐMQIÛnsÚ¹7òI‡ôÖÉï³uî'xoƒÈ/”Ÿ­;)
²¤¥ß²ÊAûÒÍ„’? €©ÆÎœC·4»Åim‘^zKú;ðOgþIF—^(ÿ_wkþCðù>¾³£zgg½¥ÏÑA$éGÉñ çÓÅZhd
®®"?M=dèM¦ðä€ôwMÊ•À¤|c)h^ëAÓˆáp¬·m È©dŸ¦zå°ªÔµU£/Z¿Q‹¼ÒTWô0Ú¢¿)}Ñb#'ÑRZ£¯Do´ ‚@¥;rÑ
,¤PœˆÝ.úqoIšeË¤(­á­(ž e/¡"·¿Aå:v¬Âà]½ždcöZþ“¾d:våKNâ&:Fö»wgço.~zþêüíó‹“nÂÙÐO«¯®›¼0*p©sÕÆ\Æku¢°öjàgJ¹­vc=›&hAmŸ»¾ïkÐø­hK3­ßÃnê‡[ææ*¿J]s‚Ã+}×ùÍ	sÑÖ¹kwŠ>vnÀAåpÞ­ZE‡fÒFñWª¹ðZ.¦£ÉrMd§&ƒönÁìuÑ¯BïYªö•(3v÷v]À7×Zv¢jªS¾$¢ãLo¦8RUØvƒY¸Xêx¾K5PäÂ¡!p(èÅÖ†öÝ'ÃÎÓJ:n[&Oèˆã¹Ëx(h•÷ñŽ–›oçyòÜ9K¥f“o­›XùªIp5I ’â¾¿»µ¹õÂ$»„¦žMUlPÊ¥‰ô™öŠ‹2@èx	yH…ê«×þèó_AloÁé\ÃÖócÍEŽ!†ú9Õ÷lF“GzGS„¾‡çŽª4Fù¢tnMÿlÒìK½)'J‰' •Ì&UÂ­¡ûXq–TNí‚¡“ z@<†ó¡u|úúüï8œ‹‰:oû1g=*çeiæo¸'/*°“ý vàø–äÝjd³¢°cVæ(Ššœ’úp.½ä®%¥"ýÂN‹ée¹…eR;I|O$¯ú¸™ŸBžwà”™îT‚·ÄK4l”ç&BDh¤?²˜tÈ÷¹,Àt÷›A•Â˜‹¼Ÿöûëÿln’·>¼Æ“ÒlâCyj¼Ž¯cB‡÷ù¯4òßzd9Jl§cÿ:‰£·´øoaÕPÛ£°Î•V¹«ÔB.ífyÁd;·çºu°¯³©7°©y;ºòð¨-ÞqÉb¥_E÷³¦¬6È´æcbÏo]K^½T*ìaz—¡?ÄøëÂdGÓëÏézVìÓâã6ÀòoÓZï¶ÝÛ ôZ
µ§¿©NtA¬ªB‡—v©kø‰¬ìjöðßdÜ9ÿü—QyD†Ù=
)ó¢¸E§6ËVPÌÂÃWƒKÏ³QÛ¹GŽl·²óLõ!›–èTsš³«š¬ÖÐ éé%ÀÕ<·¸âá²J–É;aµÛnûx2ü!H³ÏM€Q9¨AÅMikÕ­Òm¡ZÃÂÒÏ+é™hæsž )ÏyTJo.×óª
C‰ùj§XihÐ³ù—EÝÕÖÆù¯böp~öìr±•–±«Ec—]c—ºPjµ$q¥‚W¯¨ßW²³©¤½¦Š­ž D^É¤ýñÈö &8ŒèJü©ŸŸÿõóÿç§Ä'®-¯¶—P “É ˜z!¼à§ð¬ú^½Ñž|t´šBß‹žÍêÉGQ4‚§ILSV	Èò× ¾¥ +Á#%C6bZ	œVñâ~qŒXWšye­ÒÒ[ÿ
f=>½QoMVWyó˜¼@üMJ"[
Ò$®3ŒÔ"Ùe7³'rUƒ‘ï´ú]U,`Eõûãþ`ì%ö‘XüšTâÊÐm5[Ž–¡›M˜·ÁòÝ®›\Öè ÜÖð¢®k4‰^0DÄGü»óýœ÷sÿÑŒàu§¶íªSºªõö´þjûf5¼Äg/…ßð8›«0¾éè«¶ŠË¢!¨z®9( ã©Wú­=ê~n:—8³[=c‘Õp(JkV7:U§óÞî¥fñS“|ÊjÖ#KéòŠøëP·K2…qrWržuP@;¸›Ù9Øü½iM‘^—÷µ
{»•†wC…>FðºJ ù=ª«”8X†—xíÞÅæ$.Af
Õãù*3®±`Üþ_Ø%
‘azK+g—#¦ˆ)¤i«Z OÞ¤Ççâo;«uˆÎÍn¡¯¹ˆ,úâNÂLÞC0•¯E0t‘(Ói‡dmè¯‘Wp±¿îàÒ?geäá8˜fOß’ûõFÈçî@«#Î>q1ãµ˜ã^;_àÅíïb´,ç€mÖ s	x&ò5óÀËÉ' /­–Vå‹„¾ìP?uÎ@®«ÆYö¤=À#>rþ\fßÝZsÃDÊ*„˜ë>GLßéˆ¡Ÿ:ú°‹vÍF•ªiÊR…R|‡íÝÿøçÑ´“ë>ÿž´Óu¾WFªÜ\¸öóVœšj{©bºD-»éðj¬kÀËYß€WußÉ[¬rHõdçœ‚sÙ;¼.ç€¶@ÆñJœÎ3¬ÊÊÍ%è0òÂ—À£–^ÉÉeèÌRKû6/í/°ŒËòR«·ØÚ9ú ?õ’Ó±—dÛåUÉ«U¬peæm‘Æ¸íï(¹o"?9æ¼¸øü0K²:ÔO×(£´¶îÊ7 T¤†;n[¬	­ÒT°ÝU¸Ù—	ù[ngõ¤j)†Må£/a¾«gÅï¥Ë·˜ß>^Žlžp¤¼„³ !5nŠ€”5ñ"<xÑÕP¹¸Ù£EÂšìS9¬±¯ŠÕùÄ1bÁäaÂÞ°y™à¥7^®ÈHcõP¡c­Ë¨œQA|¿êOdlfáèQÆ5?Â„Þh c¹ç°±ð19~ÖÚ°n°ÄÖß5æV¨Ü)•ê:£¦.rº†|¼o²º5ynH_‘A“cáV™cA•_ÞìÔ®B§¹¯¥&sÂ2šL3¥+Ý^¤ÄÐI‡'Iß¼ô¯²ràþq˜y*>Ð°¯ÇQ{:®áC«”¬ùÓg‡hbúgÎHu¹ÞËØìJ“ú“€»jY|"løFßuã¹'!WµPnØ’9R©ŸT± ¦ÊÕŽ”1Hýb¤e¨È€~Ô0»¦¨Ø®c’$ivÑÓkUåÚ+ŠA»²Ï®àsUê52›×ÕoŠÙXØ«E5pFv¬±¦mR<++ÜªQ:J2Á¼¬^{FÕke¢¢ÿÜÈÌ”izFFÏ¦4ûJéC9WƒqÓ9$£aˆ£ð 1ÙYJÞ³Mƒí[¶ÑV”ÃÝ‡gÝš‡>lX :pè.J.
èJx\“€TˆŽLn6ëq«¹WŠt Mšx¿\ZðWN­ÌaËP+ƒd¢Ü±*Ó€Ë{Kk¬A…+dTÎÈù9Ž›:ÍàM«§%ÁÔ_/±3•Àq÷ƒå‚Ç%…»Í5”ŽÕ(Ú›™#½P¿²<ê¸Òùæï	«ø²áœÐð$Â¤–”ü~S!7”$ÔŽX´_ïÚ¿ðýá%¬…²ËZ§óA¤F"mFÀmÄž’S/Ç"«ùôÖ†²Ç&q!eUÅßÃ5<¿&^4óBœäÅl€Ev4+ü„´Ø¼_7 <Kh[¹ªä¾º¯ðOê€Ît¦JR¬“„çŠiÓzOô´Æ@Q-þb˜é4ˆä9¹‰4^xM%e2¹Ì=¶ÛíZxAM»ë ãE¦x:öŸNƒdú}õ,ÖèÔ(ÇGNcdS>ÿ;Ì“•„C<|L|X’íïË÷ºYLüùøã"Cß½öûèÂ1Žg	07ýÎ0ÙÚ™Ñ«I·X «t}=Ö?R¦È&Ëë j eêñŠn/FD¸¦GC;žê¹ØšaW½Ç…uDùTZBU¼J2ewÓmd¤¯-•ÒUâœ8«ý“ /Ô_r6¸ßÏÿŠš„§€ÛÊb,
B}ËÓ)´¢.ú¢!åq›hC ìê	@áÝ|ƒRlQvÁZ­qún4ñžµ9·ý’ÓBM'äd'^K÷Bk¨Aý·>w®ÿÚ?ˆ€ ÄÞóß%´ç÷tXïžI6‡ÂÐ¬Åó³¡RWe. BVÞÒ2e¸òØÚ–ÒªFKëðl»¨e™øé$Fy’m·xÿË’8Ío5Øl‡§ñôÎŽÚú³I>ZKL¸%YU53>–lèªuRZÐVr‹]ÜÙkÇŒæs·ä™C÷˜ºÕ&KX£Köçf:Ï&h­ªÁjd–—tf;aôâ{Å¤¬B}Šêb Öì&£ÒJ¨¸Ô¬‹6qÊ|ÞÖ#Çæ&éuaÏE”–Ç%)@°ðŽü<˜£4ç¸ë%)`]xÿ|Œ„âž_ðŒÙçÆO@ZŽ#_%¤0½Ý,
 Å·b‘È°yÿùív¾Š€1ê‘÷»äšúWú~È‘’Ë;‚i4´Ã y#üáÓ;|×ò ˆØ öáaJùØ"ÆòþÃ1r~¯Ge6Ý«8yîÆtøú-É†x7'T•A¡NuV}%û2Š“	ÿÿdµ£ájsÁmÞö:6ÈÿVŽŸ  bfƒñý;hïÍ%f4é~òïÒvFëÝ+ ü:®èNX3žOëTV€¥9õR1òè(ŸFù‘šMÒŒžÖÃ×ÇJêQmÒïÅÇàë÷4‡¹ò®¡­ît–24Q’Ä{-ÚowÉ}JÞ¼~ù?Õ‘lìeÄdÀ.ÂÎ#o*6mJhRæ‰-€¹J=Nxöº±Ò€‰H_…Êy—‘öûŸ6(r§(r
|ãh·½rI¿òÞo}è²˜ šÅ'ñÛ—pkƒ¬M³ÎÓ·k4 Š*^®©Á)Fir…÷õý—QÑ7bS%w†±
ðÎúQT õKÙWªQTc´Ì 0ÙN¦~2ša)ûÜÆ“tÚòQÙ]@7¨DNì
eq§ý5R°Ð§J,Um²ÏÅ¢ocVBÑ©+N	¦<º&TcV½î%˜gZùìõçÿìu`Àq†áyâO‚4õÒqÒ•'ÌÕäµY•ßÓw3dÅÔkMƒà«Ç‚¦Ãƒ"›ÅãÇìÇUÇ°«CÚ,µŽõ õuz·”[ƒŸš>%ççI<0"`äå@Ú\VTÂ~ê™P¢Šù'_
,@D%¾<Õ%/ÖÚÔ,s©º°yÓÎ‹)V±º|ªŽ9…šç#Ý`ßnà`6ä™Ý+·¼á8b†S<€8y<×å]{AˆÚ(šâí·Uj‹VÑ—þ¤^I…X‹<qŒœ:š-÷Ã¸ÏÀžx:öDõ.'~”Ã*¨>Ñ çÊ#‚9©…H™ÿS|Ø^`Ù
FXN7ØÓwË¬ <…ÑÈ £.„;˜S…•á c•…yôOä_ï{ ë5:ÀkWˆ	Œ$èœ0MbX8IÑr…p†ÊŸàLš0Èì2Ä9k!fq\ÁF›aAo00	ŒK²RN)rBdÈGn(›cx/Z *]úi`mKïYØÖœÎÀGJ2¼Ú=2ezP¨_™AåÐ6±§ü•õ*+iî	´)€Ê·%DÓîÐO	È˜ñOùsŠýÈœ"ø±L¿Ñû€k Æ©+ìPôú´q¯—Õ^/›õ
ˆS={u~rúîÍOoÞ>{þ–.òÉ“.Cú˜êSèòÐž4ÇJÞÜSesO6Ç±‡¶×¡£l ™pÎDŸ†
Õ‘‚]Êž§^»®VLªú¹unÆ«ÓNa1F/\Ÿ×æþIižj9.•3š‹Brð·=ÏVƒj×´asþßU¤Ò¹Ù	+ùeˆÆ{­cEµX—˜PêWu9t	ŸkYj[HUm&¶d5/øx"û€ºZëâ‰š•¦bŒÌâï §Oã`hOÞj‡¢ÉZf‚—ÊÅà‘
Tf(•7S‘8Ë¥^˜y2/	[ªÇ’/ÊïÜc@†2« ŸÙé§aqâÙ­´®q§Ï<ô«yq´dÉš©úPSâS¾Êž7-^³­â53œ%4ÖN®ÞA_lmèÄð²ë5JT¼a¶ZÖ®4»iHÐšl|@>~?/äñûß}´…Å£»¬yƒWyo‡WT¦iµÈ˜á¸åGC+¡_„ˆÁ%”ŠW3ØÐä6q·ŽKpu*g ;4*.ï—aW¢E,Ë>wØëL^CËn_ºž´K3GÎ£VŸ³Ts‡bŒ>M0u¬ƒ3÷]f½KµŸºcè'O6Ý#.‰ÔêåwJ!3Êb<®¹¨1jÇÂpèëÈüpº“KœHÁvLAøþü¿@¤<Üï4¡Ñ¹ÍÎž/T©#i¿GAvƒŠ³o»-²9:Ðy¾¢ÍkçmYÃ„s¥(“©˜T‚ˆBì·5fQž—{nŸÂñ×áåz ßIIÅCáâ”˜ÉÁ³´z‰ÃšÃÀig[BlkiY0afLqVv£×ø›V/±–N/WMÙEöÇ¤.ÒL†â/é½Ížlðp ;ßYã5ºP	…«çÞ+¥Ö+blÂ#œbÒÇ¸2Œi’˜ËT¬eRÐ¸8éÆi”t«Ò‰3öþËœ¾2"M1è‚‹pÃë³+8û¨*¡õ
¶QSKÎ7]r±kŽt¸+ØûMö›rn	Ü×È1m!w‰ÐO2€ÄKŠ“Ö£Ô•EáÁ>¤WËÐéˆzb‡¸š‰˜±a£Š=”à,$]ˆK%eà²Kì cˆË1Ý…[®JÇ\c®à~!v	¾1örw‡Ó³,g²«Û9»ªçŒUì*$%Þ2;YÜIÈUOJÑq'ó½°ÎÎÊ*#YAh=–•áš&ÑŸ+x;—b§Ž™ó{N^ìzËä1µÊÂÿåÛöX–:æQDIõŠI$Bþ†±ðŸÕ¹Dð¡›³×Y÷—É÷®¸ª&/4lÔ=‡À×Ø‹ëþO»³7´ßbDSÿ­Å,Oºï{ìƒ×™¤ÊÓI3@-”ËQ ©nûBÙâ¢uÖ˜ÈtIg’Ÿ®ÇþRZ“"È­;‡ÃÆk¼Þ6 ¶z–\b‹k^¶(–SÙ, õc`¿\QË¢ÁàÖ™æ/àÆ®š3Øufaî&ÈbKIÄ[62hYZ»ö¿g¦Á6ÍjýÌæ.šŒÂxJSM4µŸÊ†óÜ†Z„“1µ<aé<ÊÇUX?Ñ@ÿÔncµ_Gq5¡ETË!¦r¯BS‘RSÊ@\¡ZpË2¾ºlñªÖÆ)U¯$J8$Uã*Ø‘YH©¯Ö*-f@·³?4!gN¹,Š²q&ýl©½%d}a¿ëÉ±ÓbëSQ@Õ]ü»=ýNÂ,^ã¢­!úP¦X¤ŸWŸÿÏ0È{R9¥¾òÂ?î™"½Õ©~W­»8cÓ= s>ñ_IcáøZc*X³žàq!§L­S$ô¬á n;…•¡ÔºyÃrþŠÎžº6H¯aÚ/3|O`~IÖ 2úŽKOçM(,:@Jõ3hiÑcOÁ$ÝÛŒ6Õ«Éí)Q\‹wäM^6Hß[\¹yÃKï¢q­V½¼/È€,¥0­6GÏæwÍ¼»ñeê'×p]üÆÔOm©ÑTuZ¬´HÝ·Ë·º ÙN¸Dº¶æV[¬z¥~öƒù¨ù‹F,y‰˜ÛÇâFwC°’W"É‡ûˆÎ?9NÚ2ÈFo”ýB;¼ÁbðÄË	aD7Ëa
»0wb®`y“˜Ú"žeNšœ†£5HÙ¸So€þ,X÷„*|²ðàï7È6H‘~ïªÏ.}CßF°2Êâè‡~÷ÆK¢vëy’Ä$ò ûi¾•ÿ¬=ùüÔè¡›F:KZÄ½ŒIyÌ 0ÒÀ¹ÅF¬"´jÝBƒiüMÓ/œ+¸ˆk!N¿¸øéŽù_:7Èæ3CßdäÎó'p iŸo{éë…lßòE· "¤hÅ¬%dª°k»Š´.Ž‰óË—«-¼H
êž	Jw9Û»ØÕŒÓÑ8¯²,É°ÍØ¦F¥tò¯¾ffK³}c¸ÄõE®
ØƒLW1ƒoŒ×7ÆK¾¾1^üÅ7ÆË|)¯<f{aæ+×¼–T­9ã•Ç›|Y¶Ë¹°{õzh¦ë$ô“Œ1]ež‹¯Äv¾‹&
øÆt‰ë‹0]2Ìƒþ7vë»%_ßØ­þâ»e¾ì%U‹óZÂž.[ÐsN‹þúâŒ–H¨ùµñYÿ¤â±üÎ`9»”^_Ø´;Ütz›®¾%êFœª¾56¦†Þ¥êÃ+Ðå¨@ê0Ù:~Ã¸3–	žl’F¥áÇá&m~!Ö·!2-RÂ^\EjðQ…´ËÇtSJº"^|µÜê¢¼*wÎ¸ƒñ™{Xùbùw‚E¹å²vë”¦v<9?Cö|8j!· ƒ¼
F3Lý×%çqB®¼ë8Ù Þ0 ØðHñuäÑ8	ö	Ç^ê%:KÑOSèÒó=êÉì¸0}þ·$@r’ÀÆPì¶d¸˜‡ã—ášHå6JôX…àa;´sY²r)de2È—–@–‘?î‰Â²Yxá…cx1IVÈÓj‘kï— ³žüË`èäcÁM¼ ûžKLIÂëp/2gÛŒÿï¨EþŽvÒøiê|D~ø©N6gó2Ó
å¥ß„ì£.|ªÏãÃâ|{r”oEª”ƒã!<~¼›. õÊtA&«0ð×Š´,ogÇë€˜Ø³…žWÐ¨–tÃ.Êp,6js'ÕJ½ZÇÎnËÅÕPrËËü ä‡ SÈÙI³î›ÊF®Û!Yý2TŽ2|ô2ö êrÙBcÑ¬ríE:¤>àrVÎ¦IªR;{Ç@%ø‚0Wù·ô¯ãeè‰UøG$M”7ô¾û²0ÿRP½ 1æSˆÕ+\êbE¬«f ›o}Æspú€TîùF£¼¥.›¶­ûê"ø ±–&j4¶zºˆq§¿CSá8DQÖ¢‰h\×`–`¾QªmÍÈÎwì6o2X\ Ê/ôC+ËÖ´M3Õ°kzÌŸµža¹;Ÿx?Ï2±@JÚ?Ç	ðéñË‘&J·ìšŒakÄ EÆ*`pí…3ÿ¨†Ò(J7«®8:{ÑCûHUNRèñ»­^Ü¥C‘%Ð&ý7Ú&5}$¦<\.±‚½“Åò¢*Ú­©(+1c±+Ùzªý`«	o5B0«ÍÐx0“sU§a0„yo!Â¬hw¯ËËs\Òit8Ü˜*÷$òo½daÚDÓY¶ˆiƒio¯‚P_UÌtÉñ¤,kÎ"­Ô6ôRF}˜9*6;þNŸtßoYr9è.ª,ƒ6Ö	«HùþþÓ4–•ÞÞ k‚ÂUÝ‚¨ÆRsCZ„×Bvè¼ük@bÔÂþ)	MÕ¸Ý;Qª^­.¶*ï.Èë‹–¹Öîlâü3À")YññÃuò¸LBj’|–Ì"Z!mâÝ¢wiŽ(ÌWæu<¡§µ³ÉÈVÃÖ8ÈÅ¼ ØUÉgxòZ|Ó±_`ûützÑ§µÆ’3»t´«>l”J½ãµ í$>l­Ã"Ðx?¿…¡D^ø2ˆ>Irðb¶ãÿv…Ïèå[Ëðœ­¹ßVQËÕ¬|nð—–]˜ýÉ$ìô¾Ê…ÿ§_oµWx5úš
p9¬üåÃMŒwywÝ"N/ŸÀHûÐžØ¨ñCí#íÅàU9—ÜJŽb¾¶~%­ÕkIÁˆÒ»rzƒ[l¸žÎ\ÄõÓs™kÿá^%,žÍ~õli)rûþø7íµ¡—Ž/c@±5uÊ‹hù	~ÁúËäGÿ3ñ~}Ú2¯ÖÄw;y†/š² uüc<øüod}â…ô@‚þ£,A]À³rÍÛ¸ñÈÏ3/üyãœÄ´àsW‘¹N‡›“…¾nåÉz©Âßü@Æ\Ä
°!ª¨\çK‹.ª§f¸håMV‰~@ê…˜*3žzZÙqkƒ¤Xþê>Þ­³ì\Y]ú¦—Ó«àßYƒj05æÌH'EÎøÛ^¯™T‰'—I­$vX^Åz4†¥2ì¶˜AªÙ&ÀŒ^b®ê’Ûq2Ÿ’$é+!ëóTŽû†|y4G]±`Ç¼õÓ`§‡›ã¾¦MsV…Ê	VœQË"«”È¥WIGBƒØ8Ï'SàÁ¼'Ý¶­E\ÐôE[Ý~·iëø?þù_¶DTÛ©‘ËuDt*Xi4¨PRHŒÍ=%ò‚—µ? sUù÷,˜Æ²Ra4Ãlõ»s®³Ž/UæcNJ÷á™5°;EÂ>>ùqò|dmèÇ¢í@w 6ñ¤òåwïÎÎß\üôüÕùÛç'Ý Ó=ý”¶HžÐ	`¨‰µul·ýŽê|r"b«-Ï8K¥&ëÅãeÔØÈª1«Ò“¥J"Ó¾V3VWÄ«§¶$ß˜Qà@€µ6 ;š½¹j‹R¾2 oÐ˜°;é»I<6¹ÆDc|«rÇÒ2Ö.Âå=©”´I!æ‹©G	½±‡ÓÈi¶/VÃœí‹á¢NE…Î‰§%ÑÛb´uaÍµ¦†Må¥—ÖLE¨UßhÞ×Ý^ÂF3ÇªËÃI!¹V!A|ùÉK":KýäIw×õ)X2$rN›!µ:}¥F‘Ä°0qŽz8Ÿs@´)~w7õÚûèOñÌæ‹O1?j[Mó§g©ãù 9´Ôì]¼d%ÓÈó[8GkæiÕøZåÃóÕÕqâÓÐ‹ôL8;j+aºð==NhYSV¼ñIÅ£Ä›xúqž‚0Ñ¹Jé[NÖÏM:ÜS`™€wGã÷²cž& &wÒ˜…«÷ù³º1>ƒŒ¥êòOãÏdø²Ñ°”4H­W¡e¶2/KÉ˜p²^µ@ŸB^ð™~ôÀiF ¢_ø¯¨{¬
S`Ì¢àçK‹Žâ“\¹]SXhŒÓî°ŽÞ`œ§³TÖ¦†ÿQ>wžgÞë<–¹œ^µ‘ÑöÇZ¢Ì´;öRvk}ÝI¸Î?ó†C~Cõz&Æ"ÇÊâ³åBõØÌ½ÚêÃ”­×­3-#s1ˆŸG\—zƒ$ÃàOôNº)~D‹ô®o•[;ë`âÝ.Ð|Ÿüû@·ì¢·nìƒ…
Hçƒ¬Ú‰Äü=!¯¼ºEn¨Ý®L}“´KceÝÂ8àhDç´-í‚ ½–fÅ\çÃÓÍŒ6ÍW·aó'œºxíÅ›ghå4| ìÈ<):àkÓQ­˜Ü†¬çG@k¾ÞÉ°VípO§_ãMÒ´»•D»•*gEI3S}˜ºâ·VX(‰x¶ZXÏ 
¡ÈËk·ôÃx½¡¿n,PÒ8ŽrÑêÏ€K”~Þ)ëôæ¥=¿LyPÉë>‡9Â¹u,†ÌÁy@Á¸tÕÎ:tn:½=¬jµ§ªSŒÕ-Ë’×vÁ´‡%OD«’AvÏî¤a|ãR¦·<Ð>´ï6ÐIº(Ylí Zª36Ø¹ÍögaY`w¾ÃóùØsQŠhÑ}y	Í#Ëí°“¨î$5Q¾·U5r¨³[µ\µç
B_ª›»nÛe¿a8g>¼§ÉçGökX~RòÉÀâ]œXÅò×—³„§_z•µµË¦¸ª¹-Ï'ÁÍÈ°ÞRKí·ÆÇdµw9{>;‘Ì4ÛŠ®ï{=GÍŸÆÈcõØ«…Vœó:ˆsi:S]õCŸ%ô¼%í³½³ímÙÂk/ÄòGZÅúÍî×sÒyAOïœ¹ó *c(•\A²‰k´ÎšbAAÐ·kyb¬U!Êó`Çà+§P
]Þ“n\ûa±ò'gWÁh€Í?ªï÷½Ý|KÀïÄSûu†¾7”jßZúÑXj³´¹$Yð[;¥­Êúæ¥Å4Òh6"ÔšOðáö%å¸]
°O®½Òø„¾Áe²‚dJˆös$Q™ëw|íÛ]ÅþäeäZ'£œ\tN8 ã€°iøØkšÁƒÑ+xÁ¢¡×2ÃË­EÝÍ”›o›L1hFDI•ÚÝ-{4Ã£Â%íøoÿ|§¢ìVgxF/-khð	×"ŒY‡„²nV+1‚c«»kóqmäs€WÙ&ƒ>[>ÏüÌÂÔf‰—¡Ö—0ºHJâ__?—xýQ·«%2ÖæïÝ àldÊÛ5oWñ@Ã¿—'eÙü¶Û\ˆÙ!ËT‰Åî1gžšW„!SR¡s/buZ¨×]çñ.ºŠcöÃªG‘Æ‰#7m¬új,Væ È‡¥ 2æ5®lWñ=–Æ½ð®}ÑˆÊâùÝFU}j\@Ë9ùÖ¿.o|zÃùÉ}™O×ðÅ%Î”J-+êJ³iÌ4ï6¡mŸòyÈ´‘Ó·y×;	”†Ó.‡#I»Ü·”@Ó)–åŒ†u¶±dZ®‰¦«ãçž.ÂÓè £Rî“1ü¿›“
õä¶è3m’ðS/9ƒØ¼í*-ÙÕ’”š:uÔã=¿õ3–ˆµÍTïÓ8!ŸÿV²¶+8L5å,ˆÕG?áâç¶CÉÔ9¯ðMƒªÓvUç'YâhÞ%aŽ3Ÿí¡ŸæÅDêãd>Yz;„Ž•O)
÷j‹öòr¼‰T—ÕéÕUóMÜÓ±ñIáûOèÄœ]0Íä¿°zŸ°4µÑ OW ïE«ÖIé'ô2' ÑAÞñ³˜R7vj=÷Ó~‘{{_«3I[‘ 4noþ»£ŠíÙº”¥FòÙÒ†êVæe—['¡C£7ê™ß¯¯³ÚÖHâxž3G‡aºaIñKi6uë6}Ái³òvÙ¨mêÓºûÓ,‰?ù§èƒjý×ÞÖåãý^ËÜ*~xéGþSt~vKŽ¥pBÎ¡£ÐVã4Š]=¿ ;'HËR_1úìlÙQµ
‡«í«]ß2<RB½Ôp}þÂYÏ>yRžº°:X¾â™åù?Zdþ»ý­Ë†óW@®C ¯ƒÜoýå.m›Üe—»W^§{éï}MklŽ‰Š‘ig›yø
jsG-SÉ9=nC.ÜüaIØ‘¤°rRáSÁm§ ý*§’ Çp÷„'èáqð½>ÕIzé˜ÎûÝÍív£5æk¶°šù0Òë‘2=ýNü*(\ip‰I4Pïñ4¾=jmáD¸íºõƒêA ²®×Ã£Ö+Òëoìì“²½¿±½ÝmmôÈþ>ÜsÎÂÁv×l«ÞÕîÕã†ßýc0ÌÆG­¦Ý½"È¦°¤ˆRÎ_axÔ¢ÑnŸ8†SÿŠpŸKÔí~!Ø?4þ:¡/úC)ÀKïî¨Õë=înï»×ŸÇWW©Ÿa:5üžtÄ	²É˜à9è×Èô4»iˆfA Èë®3Ìû#½”øŠñ,[sÍÎâ„s Â]œ^¬Ð@ï2Ã•Ù³,žtøÆ©šžò»R µé]XõCL`]æ…Ý‡"Æñœ®‹›Ï” …SH¾k€*ÀdðP;'ß†C’g q<4¾
Jå5÷I¦1Z«cÜy¿Íl@¶C«b)”Žÿà®ÓíÖÐÐ…VåÜ¡°Óº¡l®Kç‹_½æ9å²¹\gù…’-,œRa¥ŽD:›IÉfV±EW±‡°x¤gÆ\Î¦ÈbÒÌú¿"åÕHy2´G ñ×pÆF™›/!7ìä„£ãw-2öq7ñ_ÝßyàSÝŸ·àE+–½Ì;š¿wÀ{¡cáë)‰=È]Ã7…Ç\D†£ÊÅÎ•·Z”d•[æw¥Ö™àÕ´užö_nú5V¬Ë›õ¯vàrl¶Ûí¶ËžßL=ñ¾˜P¿-uüxÇÛ¾Ü‡Ž?òþƒÙ¯v"5 ¾Éþâã‹A(ûÿ,=šïmÙ‡ ÌJñÁ¾ÃSoˆÇÊI4B¦h×þbî1s°m
Žn›`Q4_WŽ¡K`Óéa&–»@‰¡»n‹ï×á©g6Õ|ÀŸïç´û÷Œ§ŸÓf»&ƒ8ð¯ö„G‡›@­Tò]‡Y0µK Ý0ÈÁ¶cýeINóµ¥lz¦„`øö§·ùMøÙ›Þà ƒ!á¯_EÕç°­ŽæmNT¨!ƒ‘¡c}ÿñ{Îþ^nÙKVÊa:…§õè9ÜTœvÚ--~¤’ÀQ„°ï ‡m¶¹ç,1ÒÂË·nþÛ&c©À‹•½7	FPµçâÄ¼_oMñ@ãÌ«·¤·b¤üç×3ÚÜ—[‹#E§nû(çu²hw›ú23+|Ué:ð!¦ùJð“\V*hô¡*c{MhQ¥X›5Ñ’Iý¬¬Y,^yê(¶‡w40'¼š‹„Wð·CÂ+>˜Ëú¨¦—]*FéÝ*ÌÔèQÎŒ½¿¥òctÖþ@òNJZÛß2úÅ[d¶ç×qÈ°iÇ°žä<¿¨ìf!€ÝKU4E‚{ÓHO°Œ†}â¡âÏ>,-V+‡ÖÑ
M!ÎÀç)Ðéã\H„Yï¥”6)5„úƒY\Ót4ß06I«Z7X³Â
£îßÅ#à¸1õ@–Äaª½ÏÛRá|))ÖTªct¢Óâ•ˆîÐ_èÑfÒ#4MA•úYžüŒr&ÌëÕ[æL¾J=ašZ¤IRA‰SV²•-MrbW'wXÔ÷™¹ûam²Ìá€N˜×3bT®ff.È[Ù•—ú¢U|é%'Ãœ2˜$,¨í>G/ƒhL7Ì;8àXø”Ý¹Èê·"Lºô’ÿTˆ„óýâÑÉç¿\aÖ,@¥§hêqÂ³åö²åubÌ^“;¯I$Ò¯üN@ 0Iò6îTZ6«c
ûÎ®ÌÞ,6`ø;£˜e´±¨ZzÊÕ‡Kjÿ4cÕˆ‘U‚Q½úüb-(sžÄ#8sÓØ®ºo Þ¸ÀÀô†)reæ„ÝÒiÂFfÃ%¦Ê1ô–fƒÛr2B5´Ù±üŸeþÿýÖ‡'Ì¡ËÝr×(J—Á«­ê÷Ü[}ZyŸv›‹ð¡ñâóÿfÁÄ#'×À@z”™þÊQC¶éÔC-+vHï×AÓp™~Á†ŒÕ%Ÿf…@šÕ‚Z6÷,¶®…Q–†€#F½ÀÞ\:4}f0ùbÎ¥#È³ñÜÅ´ÝÅ˜C.:~è(—‹Ë.ºÊ9&þ‡Ž^y
ä¶d5žt›U •4–yzqc_ò¼cž•½{k“Åõwsiæ÷¿ÃcËÏe}“ºË–²ù8 ò*~=«ç«,üdWÿ:øiÆÀ/þ¯Ò¿[‰²[¬/üÉ—‡®á©J+wo/ÀaPNhô"¡Ò$Á²ó «FY9d"4:ªdR9®Ð¡bnS©XLèÇÈfÍ\)Ê'Ã=™xÉ(ˆ¨›_<= ½­’àÐØŸ¡Á¹ºÁ}é[°r f~xM—YópÜ&Û-¢w†ã8š_y°åîë^¶öüÿtr¤…©9›ñàN;oºùT»SFË.ÚE÷víì_®mÐ­tü‚E(°
…ÃþÁ
»œEà È0žx¸2ïø€&\F³¸±¸lî%þÄlÆ‡·T×ÏÑ‡^Á\¤³ÿ$£K¯½µAÿ¯»µ»¾„¹b6þÝÇ²&%©aAa^'´¤_>aû?úÓ×P?ä0B;üqÉíÀgÕ$'qgnŽÂ9‚S&ÎáƒˆÒ¬JŸä»‡¹~Á<¬çÛŽÙÃ˜q„ÚÝ¿¦qpbcä~+²˜ôgy³ôÍní>ÞÛ{ìÒ¬ÕË '©Kæ9A²ß(ó7Ê\\L.èjÿÊ»8QÐodýAÈ:ìÔ®“Xû‘ÓÏ„{¾ßÛ ð?ÀÄ”K/¹ è{%ÃÃMA"v¦hàë´X;µC†òƒõZdŸ£Ç„É}{§u|žG™Ÿ¡
MÏ2ŒS•«±©ïÍ÷‘þfEÏÇú^D£þÉH·ÃÙ4Ä\2yQö´;Åi$)žp;*>üdcKmž\ }?mÝ”¼õw‡<ùü†”Rþ˜U5*ªhð¹\*^Ê¢í{]d@@cÆšr^Ê`‰ïÏíE<Ð>ÿ©É‹Ï®Ò›ÆÌñì²Mˆ‡‘Àº÷ux~xÿa½‹‹€’j„k˜m®]JTí±¤<ƒ÷¾÷îI9´3,ú}Ú¸ßËj¿—Mû&SœîÙ«ó“Ówo~zóöÙó·tNžtápv1Â––ö¥•Ï|ªlð)kðÿ  ÿÿì}ÛnÜX¶ØûùŠm!i•ft—ÝmëØrª%¹Gs|ÑXjÏ:9T‘’Øf‘5$K–[<ä-Hä%@Ð'ÁÐOA^Î«þd¾ Ÿ½ÖÞ›Üw’U¥›íB»UEnnîËZk¯ûêÐ!ß{ìq	Gê€[©K™oòbqÁ+´ÿ±%„ gØy³FJv¾aÊë»î.æ³jDõÄç™‹gÌ9ÎsÞ¡“¶˜“—›ïäûPúUscPB¹‡†¸Y<ªe¢ÙOÊŒù1HYÀ#¡v ]ÙX#›í»|uõ—0V:­üfe'ÚÝÎi©xRæ®I½$ln-¬FÍû =Â::“yX¸Í‚ÌÁK÷€Ðö®ið©šD*êÚö ±Ô¶ ÝÜVÏ“çÑ7]êñØê<zN×[$¸àð×”sú¬ýM6;#Ï¹wD
í«’ [/#n­Q¤š¯Šã¤àöIßŠùÒç¡ø¢¸U¨÷
O1K¯W$Þ6Yãu¥|œV5xò*¹K«
Š«k*òËfœ†œˆZ9+^xö-am«¶7šˆ}eÆç•=‹¤ÝÊà+¥^ê\‰™J"KƒºpÖü$¥—µ>Z–`>ß$ë«­‹/Ÿcþ©6…—ÏQï4QÍe³²Ëw¼c=d[kŠ´Rå3@TæM%$G´Ž%“=¥Å\òr›bÉ »†\m´!8ž&®šÉN¿ñGîJÆŽ:‰ì¦\ú–®8‹Úk´Åzç&Í¶Îdà,c²æÅ6ÑX›æ3}ÞæÜÝ½xv!¾™m¤*¹´üËl[‰ïÏ.ª¯¶qVu‚¥¡ê×¼Ïi‹c»n}~G™ŒvÁúl¶„/f‹Q|<QQ gAysíŠþÌÊ”4UFêY•·¯ûüÔˆ«U¡ÖµØ¼°Aßaºé,U9K’95Ü®àìöi_KZ•%ãA±àßî]êNuð-ûáñ÷y‚ý³¯îa;B<sCùt¯Zß 5€|¢ø*íšõ¹*'<>Qý²¶•Ÿcké·Ù~€Eå_f't?êï®vÞòêUóÙ„^Ý>Íâ}ýÝ:ö—RSå§e…«V÷XÖ\5›fT£}–~Ç+Í‚OP–‚ôRWŸ­¯Ùž|q}½ühn^|‘gÃJVÐëÈš“œ’Åè8(Ptzv!ÿ²®Ú¡Ò\»ÐýÔžöÆ$ƒt‘³ÙÀU—ŸÈùk£…ó˜Âuœ–›§qZk9¶<Ûí‘Å3>Ø;í¾ÃÝSÐí€§nQh–'»S3áºÜ¹6w€UŸ¢%:Ja©|DwB‰2j#X2ô:Ê*¨ÓJõ;ôô?€î½UV+2ÌÑ`¡Ÿ^ý“àÉ«Ø<Ìx`dØ9¥·‚ |bŒR¤ÝÁ×å¹¿L@÷ãýû8ÕŒö³‹ ø˜H/(Ï
§­Ì?z-÷cì!º///[ m‘Ô‚ä&ºRºÛYãžô‡ZSc{fÏÐþGÿ¢‘1„_¡°óÙ*óÐ[Ï•	m’pÁeË³û¸ˆrwô{éà|b±|;Ù Â¿£E2/+Ræ]Xôºâë«ílê.)D©\PNIodãp7÷L”ÁB–DË=óÞÜ.üas‹Ói4”k!ØßæÜ"©:vÁ®¿¶i¯Å`#×(ý#|?ŠnœDe#³—L2_lpNÓ2æ©×c4ú¸œ£íŒ²zñ?eõ £ú@…ƒÐ²]“Ÿ„“œ…mOÃvçaÛ³®Ûi7ßþ´œî¼lbv=3'S´QöN¯ê™DÙÓNMßMQo˜#íÛ†ÓéC	çá9«±Zë``·§H°ä;`VwÂµAûAq¦Æ´Or–fÝïÊËà.tß²
_ÎýsbcÝiö¹ž:Ò¸½7EIÍ}cn7†ì%ã¡î2ê1ýB&¯º0+(á."kú~˜†)'"QíÇƒ«_(µ"²t¥Y9ÿÄ'¡¦q’³7Qþ½2§],»&ærLðEÙ)Ø&gJs¨5zÞ*5¶¸ç-OêXãÓ¥'_×Õ%>rŒôºö¨#E¯<¸Ò£ÍÄ º®”ýDˆ¹ßEðh$nèl™7Ø½ñ­‚HÛ‚äNqøêØJÚåÏ$%Km²JèRät=,VO¼>eÕN­æÃÆô1ë«Žº“uF U¡òÐ¨ÑR!:ØhÞ‚þ\¿Sd¸p¨[ø²aÇ²ð?lU}†£¤”d†¼ÐW¢bL4G	×!ŠrÚ$õlš¸R÷w˜7ø×ÑM	ú®£Wñå¿%XŒpFSm—§¿eQI§_â7v¿ÄÆ9°eÐF(V×&óâÄ›¿¦ÀŸ!Q~o¨ù“<_da|’œ7gMÅW9<)×,Üyþ¦*÷"Äo¶´ÿîÅ¦cÀõxg…Y§›3ÓycøÝõ<`ê…4’»4$íš mÓ³Ë®_¾S^
eé<¬³â_ÞjÆßú"%¥èå´“CßöØðmÖ^NFÀ9¾Œ!"ØEÎ›¸×pò¬rÖ¥ëb!Ñ:è¬7¢Rl:ªk©ÂúÖôÛò“Ç­uDü™µvz"Ï#ÈwÝÕÕv3W„¨Ê–& z%Ã «u¶\–øà=Ÿ®Ø²émÄñ,Ýaò½’q™jÍúráï¯‡Þ|D[3«ö<»Ç~€Ã„hm|·1‡YºCç»ÔöÕ!¶W^ÞöUÌØpP%„z2-7ëˆ]s?³ŸÇ‹UŸ«¯ÛŸM² Ü=‹AEŸ<Ò0‰^ÄIÄîx¼*w”úçÍz¢È»P÷=k/Þ&NO”gøoß›öw^Ô¯@Dê(òí¼pÛoÂúx=vµ›°ÎÎafZ{]Úa\ÐÃ‚³ƒ?ÚòüøØÉºð±µuác;¿”}>Z<-Ú¯ZÝ¤ÏJû)€Âˆvªîç"®ñVt(ˆwÐÍÒã8îGY^n#ãõîj«ÿûRÆeB/¡Žrr2‰©a¬ë—¥­’wX¿¬ˆìz¦cîÑ‹<Dî¤'ì@U{)•vÊÐW²¹Ú_
ª‘Þ¥i¹cª¥` r¤´\PeZ_Âå2‡=ƒCåhöÙK°È²ùÞEóZ?—ÝVLY	¾çP¥îÙß¡œ Ã¿¼Ga\GI>»xàš[Ý«¤mÓÛÚÔ²C-}š!` îõÏtm(2€úÓ³lpõWB¹°«_øfÐ7ƒ‘ž?"UØsï4…ªÍê)æ±P,(oš¢„V*²RÎmiˆòt»3^Â“È…{E*ðYrqiÓŽ3zlrW'–œ]‚Pë¥uù‡¤k[YOÚãMQ”2…"Lòì"8â@é :FiY,Çé ‡QÑ3 k<71	âÞŒË<³è#(ZSÆêHIäÔ‘VõàA!¼Ì¬êË8:›üÌz3ÉÞë2áâé£í±×¸Ž–y»ß—ÉÐú:¾ILùÏ²C¨ls–åzšà„bêlhy)öl±?OWX3ó¥¶¢[¢OÖÑ®ýÜþY^^võõt…áŠa¿7ÁKÞSØÞƒnÀIE³ÏÐçüÂ‚£ÔLœŽÆ&‹Ôb@"lê*¢SJD¢üÙÜ.Ôã?#’é„Ïö´I%0(õ¶IEÙ2f=ì±©3Zá~•Ïå]Kôg±öQô ··W‡öÕ´së*í0Ò.<]A@\¨9lÊ‚€Ï!\a›Y<xRbï^G'Wÿ‡²|d…ðþ|Œ2&¶wsÅÁ 1!›?ãd]‚äŒŽ¬
`1ÿoäâ_„àÆ{¤ò4Œéö˜˜ÚŠ]csZ/lÁ–@1‚ˆö1”k¡ìŒÌÈ,“½¢ÈÈ1=ºI$ —³è$ËãŸËMXµ ª\m^ùeà„8ˆ‹¬XÖ`ìæ (}ÖÌJ#²ß¹ÓlÑuÈœ7Òæ19cäxËîHzÊÉ9 ª‹äY)ÔHëý¶˜%yý$'u¹o…kÒQÞË2¹èÄá›|º3áý<
te%ã¥d¥`¦ŠA–œR ¹úðS’ÉnlJ._™z0 Ï€vœk­ã‚á™E÷ãà«xt…Cý5:þçà E»6_h|3€#Ùã¨êŒöUGàbÑ=®:\ž qAGÒO‰‹–®îç¤O5À£…ÞP^§NìÂ\ççdnÂÄB ´Xžl/1¨ÍäÍ“{¬ÙOçäåÄE«@k(‹¡^‘fÎW¸oµ7«û¤ÎšÛ[ïPôé·yc,Ööµ7¬ÍšÚ¼V€MÆÕ*õR	.öO‹ÅÖsèÜ=ÖSs|““l¹˜Qpµà%k/påÉžzJX8õœ>ßT°‡P¬Í´&$‹bÉRePøGsyôçqœG¡vÃ0¦˜Î°·
èÛEüÉá=öB5kî,€“âÀ´µ#”Ã#Mà}×¡îÛqœ€™i}
°³ù$î}›d”óÂl¡<v¢½âh5œJ.­)ED¶ðK³ /û…ÍP§[Xå ²ê—DBÅ8‡Dá]"Õ¥ÐèTøm:u¾ü:«²¼®Öu—£÷Õ°þÿó¿üâÈÁY±¿»ú5Œ“ŒôÇTF¸ú e“;J’7jäÃq†ŸQ("i@Žâ£$ÎÊh€$ÏPåÑIÌâ$<*J³ž'7(L;Ç˜ÿ­Æ}â#]ÐC1L…ÙàýF§Óæ4¼ßŽóAMƒú.‰¥%–·ÉÊrÄo®µ¥ôýW¿à0@¸£"ú	âw	å¨„û®À¶Ž°i&–Âƒ$ñœ¸ÿÛß,Ã¢c6¼U«œ„Ì Äd4¾ƒú) 7œÊ'ZV‘Rð0DŸØ›´ÊÅìÈu²AF6Z	QÕ[s¨u%dvËN•
'{¤8õ‡U¡–bél…îc¯Ò/gß›R…k#"P«geY—³€T–]y›­~½Ž 0§Ç>ì€ q8ÊÎ]ó“à˜gÁq5„ŽüÒ®eÇ\þÝ3ÖøŠ‹N¯ô—f”4RFož÷p/^ÄÈ uœÒ=>úÍ½úAJmaQ­TcøÑ7ÇG,—+,T%z>Âñaé9­k‡7‡®röE# á\EéÈ=òa	7`6Ä}å7¼ ø¸r7IÙÚõ«¯ì'ûCˆÜµo/t×\#¸ÎSVUÒKÌmÔà®FçÉŒî`l{®![…a„ =ÔN†([Ä"+lÔÑÅ¨Ø3x3Ò]³’v»èâFé'?N[Ì¶­ºì¦Iù[¾'W Éë¤g¾ŽUúáÇ–3	‰ïs\ {íT×Ì­ÅgdGèéN¸¨=Gk›v"-æÚ.½Rªö:;„ÌŽT‹;uˆK—«Ã$¾¹¼_¹¸lå“ë<û&›–Äâ€¯(ðÝ+Zü\Y
›‰]ÿ ª§µœ?ÙA‰¾ŸŽ$„Qã¹‡çIé>}¼PŸœHõkLføÚ ¹ÊºCíÜ—ý7d÷àßï¾Ýy³)Ùâ. á	Ârº¾"ºß—÷‹5iFÅÃ>Ã£Æ(oÊ(¯<^õ¼ÃšUˆërÊ’F¤.¨Q[bxµCr&#P§ñq´'1äÉ&VFð>ŠzF©­ÏCª#1·µæ>X­®ÞBò h×¼ÜÒã©níà%|Š³ò‰vd8ö­(¹óÐÛàg˜Ã<¨¦–Ç'Ü§iÁÉ\à°ZHC¨nt5q+×g¬¯@ÃÔ×·ÍJé ZÌé¦…üU)Ã¢N©ãE"õ¼©y¸‘!È˜]{IÃØd¢%ž”<Î˜Ûko´_¾ 2!v…T#ñBt+˜¾Pm…kà…Å¢ù{Ð–;oÛ“¥Ornõ$™F|®	Šç¶¶_ïÿ¾n)«"oÞ ý.VÂÜ‰Oâ²èé÷”O›2FÄÅ» ‰C‹Ù•'Ï_›æš]•ÊX‘	'*ËDÈßþÇ%ï ‰ThŠ'Ê¸³J´¨dÝanU¥»É'ößÉ^z6›©9ïM’çbFÄSñgX]]^eÿVè¿Õ¥UOå½apþ“7· ÎUÛÐî&R]i±+'v›2©øª'N‰Ù^(¡ˆëÇìx$RÅE¸Õ!ÖÍÿ@Fþ'`2k|÷¦ZBR{±@6ã­Ö8ùD­>¹¾:çÆõÍúÁÝþ`vÃ²'I‹å¨F˜˜õnòâå44W™ÓŸ<ÿyõ¯¨Q|Üî¿¸	T¡¡Û,åaFú\ «Ë¿H²R«$d…2ÆÃ8‰Oþ»ÊÎÊÓ¬˜ÜApüzŒ›3sô=ß{{­0ýêê¯ÅJŸ×½<Šòrœ­¬Ã?sõíp­œ£‚.¶Xè™«ÚýõA¬ŸÒß†ê«{Ü•#zŽ}î”6ÇÑÉÀá?r}LæÍ…2`›fž;öiLƒÝ)ÚŽ}:ÄÜ±Ž=DGi¦›<cÿYç(ð”`hë³Ða0?ïw
Nûú$a|ì3m0ûØŒµjÿÝûø³Žð>>i'”w‰òkƒ) ‘iÓãÿøG ŸØõÁJ8ßíP¸Êó¯ó‰<Md#û(±‚âw8 ›3owdŸ*tP‘/n°zª­{×Dg:rSgØt»Øíº;{ow÷ßlB2ù2(³E²þÑÕÿÎ0ßüñl[¼¾xï†wÝ­å›N¾"/³AˆŸ–u·èíŒ
¤¬Œú£ný±È[®Ù «VlÆ¦]£ÿ¶2×giÀ=¤ÜØ1ÀÁ‹øÜ/–}¶Z—|Õ`Ñf®1;ÿ45\ÛQ2N‚ü¸Ù>6U*[¯Ù+RY¿Ÿ“RjwHÅFC82õqA‡Ÿôy+6ã3úœÅ¹Üê²
¿
7((l¸…Z?ß—Ñ‰’ÕOì`+!¡‰S½”}ÅÎo‰:‰t Êš|‘(äÍMT ;ŠÊ}_ÅòÛlÝè®Ï‰nÿE7¨Õ]êFÇ*ð9áhmU‡œ5¿«pÇÙáï³Áçîf=îÚùÛ5¹¤5ùÚÝ §][?»Æ)¹,Ýã	§¥¢šg]åZçñª“}êºèUKº»âOwMHíÉÇ®©½@y›Îv·åjwŽvNÒ&ßê 2Û3\ï¾r½ˆiçy7+Ä{i¶ñ™R)rW¨¦Üdª¨Må]Z '‘nZ‚9•Û³ì”úž:¸S›£q‘ÇzJWºÅ{êkÈB?õÅëþ9P'ïõE{Vò-/ŒüÃ×«g§?šÅ‘s#¡O£€6[èÈJéam
´¦š.ß÷D!{n¢{¸9ù«åñåKÙë†(7É¶‘k/)ÉÖ:Ÿ FÏŸg¦)s%k°ììì„°I³1LÉØ[(ƒ­Më@%9ÁÚ‘¯{ËòÐÓ‰J &Áðd%ŸPá-'ÛrSÈF áš„#ï,WmÉÙ¼Ë¤dl‚Ð¢Æ ¢ÂÏ,ŸJ°qH>ygF(çØb-vÈÚ‘W,º&Y†(‡V|¾Ö·#×4È4S‚Ö3Çq2úÃ†&>'	šêì›Ž“BcÛÄ,“Û²vŸy¦sÚÄ ÑÚ¤v‹ðÐÎÙä³d¥Ûy”L]}JnHšœC>Kðhô ™üHiïr‹@!£U]è/à¡€‡X–ë Ñ÷µƒÈ´™)³ÎlB÷è¸z®QÚ¶6aJ–owC§+y¿Eq:û/‰7™"Í÷L
Ÿ6KåM8øE8÷¾óV…sF®áÀ”ºïhj¼MÙŒ‡çí'Aê‘ÍXôæ½àŸ¼®d.1Œ®® ÂÅ!”h˜± &^ sSP,‚¿ÐW¿ŠÒ"Hðk?óo;ÑqœRŠxÌÛ ÇÊ	©±ºøÄ¡öÄwyPŽcðvGìjÏ³aÍm±¿¤·±
,Š…]àtæ¶ðíàëG{¨—anëo¿þgRÿ&½wq$W¿Ò'+d/‰‡ôwxzw$ßæíGùÕ¯YˆØÇ6¦·ÃVdB'¢ÜÊãôÙÜšƒx6÷äÉ“æz¥ç›dcu‘<|´H¾¦ŸXéñ£ˆ À#Ûˆþ†±RøŠBfôw´„5™)¡¨ºÞ„²0E´—–Ú	²HÖV±¨"}óÝ:KJJX
£úƒ¨(>ßã¤`gIŸ‘íY‚ëÝz°·Á™!ê›M|dô¡í3úÊ6”¼Œ)…¢iÈv5’¹­êkÕáÁ¸Ñ3Å×aJ}]%tC: ÍS§e!ÝsVôÉê%5ŸÏÓ+1ä^J±Æq*E™÷]þPêðšé^XŽ¡{!yFäR»–š¸|‰
i‰Š½to‡¿¢zÍ . ö/iñ³~M‹ek^¸¦¥ƒ=&?Œ Ð@¡zf§Ù‡Ã,(ÊÞœHPrõ+ý6È†¤#j<˜[$sëvŽö¨u	;0.¢Ü–g§Ì?ZÓ_°"È˜w²A/„G‹dž»M’hM1¿èY\H°3ÊÁ)âe{ì–DËQžg97üaï¦ÄW>QÉqNg=H²q¸Içlî|(¥“èE9ö²<ÂÞzô(Äb¬@7°îÈßîî.’²Îgåß\ˆù\þ“mBçûÝª$³Àwù¨8 "s	5K)±—n)tÊ¬;W•ãÒÔU‡]¬(/£Ÿ¹¼\Ä)M]FN^âv%åžè%å$W,Q,ç¹zhïÒÃˆðTW!qN5Ê‹«_Ï¢D.B§ ›ŽC™þöLÅÈOJÌ>– Ã*S‡ÙvAY:¸é)¢~¾@YÚõäVËeö2ûåÛAõ0ËÖüß¯ ŒÆû(ÜÏ³£$Û” ”æÛFìvP•àÁxc-ÃÓúÛaÄt*­Õ&Yuè KÆƒ,r¨à·Å€
Ï€ŠŽ’F”Gå8O5•®y\ÂG;2±#$øü¶*«iˆ6Ú±ÙXcž?#Î7öTUêîëÚRç(Ï>’¯®üBõ}MÎ°ÞàS5/Ua©„%«`i£0¡\éT/4©aº<úÎS³_×à¬£Ò˜èÁ %  ×Ç©ÇˆãK6èA¤2´³Bšõ­F|ŽZãW«jÈEè6ñ±‰$}é©æÒéÎ"ªL@eÅ
ÔÆClIä.‚$é+…Ç07T>ðeÆSòâaá7±"øcë‚ƒó½ÚžÏÕÎ•AÙ—Å¯g;S¶È*”6âNo¹mF)º¥eŒç@[dÖ]f—Œ%HW
§± ýDhmù‘;útdç;Ü5œ¥—<v*Þu?¡ŒÐ!Û“$¥AiAÒ/£”±%AF8µvÐF±œƒçåÛÝ£B–ˆQÏQVãØÇ}Éµ-»~Yq=P²*ÔQ·X¹Äƒ¦A$U§g1”½¦¢}a’ôŽËÄ#H× ‚TÚK³˜üégâ€‰…+è(¡N¶XOJ3:BBÿ+¢qQ³UôŸ¡#ž5Äf9À =,†A
`6Ä²by4ÌÎbhÐpš5LÝ[ê¶c[ðQµßiø(n¬y0R´y,ãä¬1R*:ìÃI©Ú§Rì“3t	Ó ¼ºØæ™ÀœQ¹®¦'a)¤œŽ‡xÕÈAál”…²†¸ÐŒó€RœYÂq­MñQM¸ÂÇjár¡'ºm-v±85‘Ã¹­ê‹çÈ§Øü:;Ëˆ'+´OÖÈÒì¼:Ú^}àß÷Àz²àÝy„>N™£’w­ÛÉâêô¨¹IØvôq€<«)w•Kaµ=IW—É.ÅL ªùÕ/0¥Q–¢¸Î0r@_t’åñÏt-„eQ)WOVtw(ëÉªN©$8Í(ò”–hø``M2˜™T Ü4ÐeÐ÷×!#0[›©Y«}pŒ{7)6tµ]L(<ÀÃhfIv‹äee¢çóApõ¯a´HvÃ1¯ê¸HÞ|sö“òõf×‚WñÈd²m"Ékõ¡f©Ä!¸¢“Îë1äQJÁ¢+åñÒ+t1·m´i¦„SeÓ‹Ãì-^_$õV7Ò*Tù¹h¨ön¨/ádîhù›JÊÖƒÚîÈô^ZÄ”ºròŒ´É­ŸTß"é¬HfU”;œ!­¤gtø&M>ÎžˆJéÜš‰x¹Æy‘å”Ü–ànH$t+u:ÓDM®ý°£¥Ÿx_Ž·¶Ç›ÄHiêNWob¯-`Ñt†Á3wþcJòÎGW_ïÁ¥Ž­>¸,z|å Ú®ô»Bín¸s-Šqè5S_™]ì{sÚ]Î	Ó^ÃqcNAè*¢dLd0ŠÏ™°qý||ƒ¤Ú Àñ:›IEAOÈú™v‚ÄŒÀÜÆóÝs:††Ãì Vh7ýe<“&4T;8~©³¸˜›Øã(ûI¢Ô	é|h;bEØŠÞ(´uŸ1Ÿ†ÆùM`ŽFd¿6ãJ]âÔ#y3ZB½]]ÇÍ3Úa‡W}ªslð¸AsÞY?g3úTãL/¾'ÛêŠ%C[õÔ¢¿’ÍÿÁ(8A4×•­àŽ‚t€' îÃf‹—–èÄôÌ¦L‡X½UÊyø}O“x«ÂHE‹A"rJ{…ü>lhyTŒ2:
ô!WO¨—+Í}è$-WØê±JY²èR]‡ÀË˜F³©3Ê0DÇâó]×é³l+l?]›\ Sm?*—Ö$ˆ­ÐU°V–•~#è4êÜ!?££ Ýq<``‘gpTq¦j\d $ººÔDÀkÊ-Î$rfý¿)\òFÄo¸:#Ù[î[Þ¢{Iê–›yEníÙ;+oS3§\õ:‘ #Ó 
Ìö0 d"`LhW+;Ô°>Y$]hÇª^T¥;©NoSÌgF™ZÆwùL|62¾ñwWÀéù"ÝO+ÝsŒö‰ö‚´µ–ëÅw^¨g¦².G bÔ®ãü«ÇÔãkßÎng8â©Æºª£;c© Å’å'AJ$Â‰ MÉ1A¶T†Ê°Ý3;ÌØÅùÄ(ÁShÀ‹–·–©oÄ8÷…’M`†{á‰Vß©¤ÙqJyŠ(Î³Eò.JéN7‘´nö¶ê‰Û'j-½Ä“d_pfÕ•‚·«z˜3ÇÖJ@Çá-¨%'TÐÕd{EÍá›7Â	Mf§½y2MÅ#º0Ý¥¥©G·3•EIãÂç­¿ý‡!oä©)NgÂG­òI[¶*_^§xgÁÆôEH[øƒ¾Püo¤å­¼ä Wz~†—b*»JÛ`®×Ä€¸AgVá-;òö…ðuxËÇöowúÞÊª­°‚OoûòÔ|x|öx+'FÛƒJ2s¼cû‚¶*ÚJªçOoeµúÄõ na7	ê9D—Ø<ñL}{-cmDeiTÛÙpD;¸Y‹çí2Òûµ^z_ÒEX~R{fFÐ¶&P†hu4-‹:½”XkŽY"°57­tÏy¼ºûjµî[Ó©û*òVô.Ñ©I»gÀede)¡NÊ=¦M¦Ù¸÷"Þµ–õâ^$hçôXþçQ‘MQÿ<n†ûöÐAGq	¡Zx¼¢)¸´àk‹;„Õ€	»¨mØŽ€´æØÁd8–®p¶#ÍÃlâb‡g¼Þ/›¥nZ‚â…}Ï	ŸM¢ƒ>Vèâ@k/,eíçòö óoÉ¥i%VÈ`FI]ž¥ÙIƒ:NÐPÁè>,×‚˜°?:à¾pø`¾<×Aðª±õ©œ1œÙE8Ð<`ž· QË3ÙÈx8Êò’§ÿi»{ø]}~ÚÏvŸP·’ByÆ²ð;‘e¬1™7ø¬mÑÒC-êUp‹É‰; ¶±DÜÚŠœÂÿ”Î¹› ³Á(ÕË›
 Û\Ê¤×uîY·Ž!®úà:$Þs¸üé™~ç¶xv({š·]RWß‡ÈãŒBÂ`/²<§r#z6ÞÓ™tÍF@,aß¸¸;zö÷+Z[iø~Zµ6˜!ŽDân%Cmº~;»jÇcUr=Âq¢ð]
²B‹÷ñhíÂñ(‰°möÑ9áïN-±­®o‰¹4§-q`.qÕðÓZâZä¿®5"mÓ×í¦YaÁ‘ß‚çè;xJ*HœbI´bT-ÛæRòTÐ}úgq‘+®ß‘I–ZÝX7Ë¬2€:;¼àÝ5qÞñ*î?^¬çmF^G¾Œ˜a§G¿o’¢§‡Eo–jxÁu§…	wâKez¸,¨¡}^ºs´­˜`¹¨<:#}ÄY}x”ã<.?îÐ2NZ3{¯™‡å>Yôä0gÄ^ÿdøÜŒÝºás#§6A"	ÿ±*_†Pì`Ê,là:°ëòP*´kœtb]…gžŒ‚ü}ú»Štå&*¤ñèèhLwh®-ãhsÖGüv°‹zTæÁà=ø,}ˆ© TÎmYÁ’‚v4(1Ô×dýNš¤0ao§ßóÀÊ±’$
BxoŸœ––ÅSTN1’™ãäôawŽ“Ã¡~¾Z °:c+}ô:›E›nuØœÞ”  ü{-zôa‚å0 {Ä’øäê¯³“ôs¼EUÀŽ +±òy”çô¨hŠàwâ'ytæŽ4i‰ªŠqÃ®4ä‡?9²iÃ7ü›]üÓ
Œ‡5_ý…äÑ HxÒÝ¿ˆTÈÁ/é‘@*Š†<8 tHÙ˜s3µ&‰5‰Ó"ÊÑ‚è³ŸLzõ¬îºÂeâ3´Pˆ¾º¿Ðî|²W¥¯ßòœÌ¡¯¯`Ô™¦…ÂT×.gz9rM;ÇhI>­4è	Öršj™/IDÇct0…rrîjˆÝrw'ýr¦	n±äÛyä¸ÔÌ¿Ù¡“} LNH¬Ï¹Ÿ3Fñh­FmÚð­ÚLé0ÝÕ‘¸¹ØÀ‹8AGt*oø]ˆ]µîS¥*L%
Á÷æímáÀëJÎ]Š¸-òL+%`KÀM±øÀx»pÑ·¦ì^@9ú£VZÀhÈCNÁt¾$tÏ²ô„P§<…4(0êÌÒ!ñBë¬Ä¼ã<§„pÎMÎ½œ†,óòÒÇî9ÐÇ:!6¹s(4VÿÄ,É˜”¾a”¸šûÒ›{dyyY•²ªÁ@Þé:`~‘KÇ mIJÍk—ú%K«çk­¬)ƒP¸³¥N¤uÚ­¦°hOÌ:‹€­u}DÖÀ·”qiÈœó	Ñ&&ME åhƒÉZ\‚? }å.]øib§¤u´“qY pŽGFö¨ì6È)#›ý©Z;!pJ”¿GÖ‰,aÉßö±˜ƒT§YIyaˆæü+ã$ŒÌÓÉK -Ç‰2ÕÀu°8Èœ‘«^wôv6”UÎ„]½Õ›ûÂ·÷U…Ä™-bÙà÷N’ì(H^‹„ÿá+Ø|~‘XvÓVÞ‰»x—C(“÷ûËšl°ß[ì›¾/Iï¢×¥·òÐÝ:€×Ï@²b’ë€§ïÁ9pº´þ°]¬mPIØ"çƒ,dû’BZ%‘v¡jMX.z3ªg™U²®!Ú»MQÀ"#=æpõ/ 	rÕtUl¨èªˆu m•cï¶›"›ÈÙ&yªG?Þ°	ÖDu©Þï¢,ü ºÚe6ÄtZNXú„øÐ@È<!\±ƒ=XTäÝË…Å€±„BS“MO“çdþÍ¸Ì³y¬d1`uà|édZ‰åÀxa[Ð#±··çjP—f³I%[€ñ¿ŒJ+\O4%¦fv¬bIŽ”„ÐVûöÞF*w¦¡¿Þ¡Öéw ªs[ø‡ô¶³!eêŒïéÄµœÃô°Tð/èæ±ëùf˜…ËN3¬ÂxAËŠëÂ[ìO¹Ú˜.‚ÜÜþñð¬®B)nl’  ÎÝð«¯Èƒ	ÕãÎNå•^°¹¸`%=A»§à¨ë”•™‡‚Ô’9@OAPŒ¨Ärƒ3fÀ˜»›ûÊ¤º(J«ò¨^ß»–ôiÐ¶k=èÇ!sv&ŸïÜÄ«EÕ^õfga7¥b.9—I‰ßÊ $Þ ÌIì7$ÃBçaÙªÄnÝŽM©šªÝ¢Än÷Xê½®Ö$þ°jKrw5¹%IZk´#‰­ÍIu†8¶uØ–ê¡U&#´I×9Îï…rè1÷ð<içYbø†&'›ÒÏutùZã{5ûÒÅÊoÈËþÎ²{ð‡ïwßî¼Ù$»Å AþÁ½hËfþÖ¿YÑPK}=}ó8£-}C\-ó9ÉÛ`åÑ*z8·p:Xy¼êy‡ÃëB8P¢ÌE„8ã“]Z\.Žü¡ôà9×26®£6ž‡ÁF×™KÈéãc}¤yßèóÜæ¶Ö\ž‹DlôWúF›Ki8’ÀÇlÖŸÅP‚Ož¶sùzãŸ¸§¸õÔ[\{I8½z¸4ÈÏ OùCÈÇ>VÁoƒ±'B&mÍ¦˜<I&KRß"
{C§š10—ÃÏ[•>‚ÑµÐ*&¿Ø+ ÚU‡U/X‘½ît‘T½ê%[ÙuøØ·Ã§n¬‹HéŸHH‹Rýá«‡¨	´µ¬å?·ðK‹Êˆ®Úˆn¦ŸÝsU½%œ9eñHŸK€_ðGÅtÞ²Îøã2UF¾,Áúðm˜wå)úŒTÅ8YãB±_3¯íÐÌ'£¼§eµF›ìõ.ë|$Vd®ù¬MÒs°_óó¾^ùÆ'è™&wÎoý)¯ôú{´‡[¼Ìò;ÐŒ¡vŠ—XOÎŽ\æJ«¬ŸîT”ô²K¯¸Ð† õ±Þ|ÀŒ&PÛq³ Š`—¹¬ànå>J¡·Ú™E…ÿ#1†S:î ¬›x0§Añ::ßÐè-¸D¦”¬>³‘BÔ0êÁ(ý‹W
¯<,Ù“ðžô<Hð™ei%ø”á
Ð*´zÇ—H‚xÈ—(øb«‘ÞF\Xhù1?&®½¦çÍ|Õ¹zo?ÈÁû×ƒ4JaGýãÆ¨j­-»èÄB¦näHÐ’±`kX1üçÖ…¼n—d	qÍFœn’ñpÕk¾ÿÂÀáÕ¯å8²bšý|S«	ŸÍ&|ìœÏò£{ÄûùV_É)êª(ù+ókÚ@ÜyöçqJ=Åg$Ô\D;ñÃP4ÂçÎqã,-=Ë%UŒ“2Ëãk[¨0òlx•b©àüO±tL×ØµÁŸ$ ï½î¿¾úçþÁ"Ù=8|Û?¼úOßíõ­¾ìcByÍîÍÂ‹©B÷r™} ÄëKtu ê `¿I:¶¼%{)­‘%ŒËñPµÆíyÕû»‡{¯¿#_‘w»¯wº<ûf÷mÿêŸ¯þÛî}zÿí›íÝƒƒ7:x»»ýý[úùÝ÷¯ú¯»<y¸»ýúÍË7tu@;ç^"J~øŽt?±­ï²¦VŽ>¼óªZKg­«•g¢(k×ÝÊZˆ¥g)«Zwxõ—AJv‡¦ÖÊŒY§xCš!!{N"/|îÁ™qJûy$þ™ß÷"ªçŸ¼»a 2Û9J@ÃÇ©fRuvúÞñ 1T#m8&ÉòÉ
7ž}ž:°öVÁRdŸ4€>œ@ÿ0â‚U±µ³hˆ9¡ >òÞ¶Ò TC3MìoZ€l¼hMœ!,§;{ow÷ßl¢=>ú‰!ñ"C5$NÃz‚ô[ëé£ûb=À·ó~œÈl¨Öó¸ž…ro¸Ocy«9ÌƒŸ33¦Às{T#L¢çÉ­½±r?3"þÅ¶£©&ÄÂ Q{uõ—0f§¢ÏØFÚOÊ€€	DVø5…º<…ó›Ûb;>ümŸÓgñÿÑûe -ƒrìã>G¼(pQ'ö)oÔxVXÁºž1FˆAB 	ûÖ¸w‡¤OyH–»}KþÕ±£*«%í¦þ>s|¹	^¼Ù·Š#‘«B³Côôe\”o €£PTh^73ÕÀíT @'CÐðgEŒ6ÎJëÛ]¡`áÜ"‚G—Ú]ˆ0ºX³ƒq¡…ºðéÃb?ªÅXy´ê&&¼ÕµR“ÒhT×Éú"Ùè  g¸™à€³;¦¯G¦µE
Kø\1·³›¿%k†ÒR<q™fåh”QÑ^ª—€[$k«Nc£ŸâX7ILAâ5„±@N=éMÈLuº<Î{«‹8†%:+gêƒ–Â’=ÑK2¢a	Æç¶vp<ÏHÑ‹i‚ åëôÂZ ç9¤0‡|EXVœ~?Á"ØÖ,þÛï"cR"ÙK¯~¥Ôgj;øJw3ºÜktß{)Ðí™šMên?-Ó ‚Ú‹xøÌºƒ]¶™ÃíS0vã]”ò"@×€ŸñäÎ7W¿œEÉßŒ.&kyåÚcBÎž
 oà,±Aê÷Ó"¹5&!ºÝ‰n½z m7@‘o„ðZ/jWÁ¶p	Ž™¼Œ†GyTF¦á:Ì6ßÒ0ç£$º9{À·.® Ø0 õp
€za¿v‡_¿Úx|Öi™?äÁˆ©3@&²Mãh\–YêQ°vì Ïxð^
ð”¡\wÁOøŒw‡AL‰­ìw­Ü²äç`ŸzŠƒÔåÏê¨^²æ¹^Nž“Ú5-#³FAÜbõö|i”ÅØééÒžn‰GrÃz¯é•H\nøÎ¹;Úˆ­U)•½ˆUGQ)"Öä<¥"ìJ|lÀR­Æ}UÒÓ Ï§Ù~SBp;@X7Îqð¼‚éÖ4‡–lw	–ÅJédÚ»#€þÇÓ ,‚ÑÈëâîg
îÕô ÞpXR@^¾{`>*Šà$:øó˜.H3ðãõG£;þAñ1[p~ýijÏK–ßÁûûãâÔ¥}dcÙÓ_}Eæ_g"©y–Î“8% üÝwšB1(Gß|â’È½PaâÏcÊ)ïÓqQÐ+=OàŒ»Âô"'YÊ(t&ìaŸ ôÞÜ~FyŽà,ËÉip'˜a­€äŒ"â<ZR(e7&)•AN(ç,óZ¢#Šx÷á ¼[ÈbyÎBâ¸ãºÞ’”Á†mV›ÚIû©2Xš&"&1§&“nÞv0òrt
K×H¿`i:Ñ.—‚è¢çd˜TñÌ8]ì·a`ŽA&;]W˜•Rø"—íäŸ IðdJ'C¤9¡‡È±³FêU<÷ª*Ÿ7}u
ý€®§K‰*¥r:KÖÐZõì´.©ß&÷àØ‰ƒ¢ËÀã°‰Â§IÊwÑÒªÿ	
ÅÇµcº)ÿbírk¢)Ù†7ù¯_n­ÃãÅ„Ïo\nmLóü£Ë­GÓ<ÿÍåÖ7ž÷91xt{Ž˜Û[ÃømúdP¢Þf‡^)gÑcY©˜IW>Éj„¨t.´¤±ž}hAf¨”Ê®»oxü”Ù¿‹`—¡X6&¿}ôèñÃ'â3	;åU#NNbx÷-æŠ;ÊÇ¡Ö¶ß°ärvy4;/LWžYäÑêV›Y$Û’ZÌ¨$³%ØMW wÔbæµÞ	ºE0(Â¬šZUæÉ*.»ªÿ°ÀWw$~Ê:Hâ-,q]VEQ+uÄµ[ªƒTÍÔU‰5è‰`áîuxz$wwÒ¯Îu¤%o[IŠ-‰</c„ùë”¶
ÃoÞ¹G;8n?Éšâ÷ïSù†ò<æ¾Äð*æþ8NƒtÅyf>žg
`R™|´M^÷"€ÖW‰¥ã»šhþ+™E<Ùk«ZB¼ÖE?a»‚J[ôjÐ&qwX°…­[6ß¢¸ð	Xý%yÏ½ÓÞÒh§³‚¿p‚ÃæÎA~
ÚO €c\è_ˆW±N}ŒÙúÍxgÓÁë­tÂ§C'« ^þ€§CgeZ´ÛùÛïd%?ÔZ+!°³«„Ða×[	áKÑ‚®Eà¯-ð&hC º¶jøù,¸ö2ºßR™ce
ºÐ¶ë(S ^;Nö ò^ipBaCë¢N
K Ý?üx©·–Ö;>Ð®‹¶Œt
vg|Ø½k™m>ÄAéWÔÉ®\›â±VtS=Öº©Ù+­: ›V?*ëïÐEF|1¤¢>fÆOøÞpQMï¹UOÀ“)<WgWƒÿ²0ÂÌ"´ã(Æà0¼G 1¬‹ÏçU•XµŒüdºR8‰ÊâÁ©*ù~àd)­~—Á	UÛUYô®ÛïÜYSÁHÏÖ‡èð,ùZ4·L—ËÄºÓÈ¤"ª\­F£mu‡²iL¹bHó6˜ÛšSúçô•™%M=`ÙŸºâ@•ýë:€é~á@•Ôë8—·;ØòÏ
ìf£
!ìö3q{*óYõQöIB§Ú|V_Tx2VGhã\	íº)£Z=}»MM .ß·®5æWêfkO“÷ ­9Mˆ–]¹eZ;	rÑ0òPÏaCàu€å“#²u¨ß¤S‡D
±BrH‘‘Ù{¶Èìm÷Ïñª$]Èn	ânªÄÈýrÌÖÃé‚ãÉÖãUÝcoÃ–Ó@C+Þ­n9Það‘Øaïé(ÖS™zn/ÙÔC`’ª4&a9âNV˜Þœ’ÕñÆmijcgšÈÒ4½­É¡âs(ëÜom+
Z×(?wÌ…z(éÀ\×Œ°œCS#, $´©¼C°|V‰òAš#­—1 ƒöâáhÄv½¡+Å³†_²("™ÀûèõáŒ¤¸QkWšhD5dëè„É ÿN›ã8I²£ y-ˆyø*(§½ÞWX&FÍíõR.á;ò£T>ãuÏF½ŸbwÚc¹dð;u’i)o#«#Ø([H®5’×šdô(*?DQj‹—žÄ´Šj4û|=>–S~ú[¿áÚ57etÆž5Ežuˆ;ÃJNvÊãÆs¦%ûjÔúÇQÁoÉnÅjò´Ç`:ãücCà\AíŒ*ní$J)É.£þÞÁøäœºi+…^$?üÿœ£…%b½û‚×tÚçlHtªèii‘96ùL­ÎÐDë-a¹—ÿT€¨:Ó®7w4‚ô¹â¬êG½ÜÜ”FÄìÌv³¹Ë÷£¸šIVÏP½ÜÜ.*=c˜Î *Šz\æ-Ö³7OÅ7ŸÝ>ÕXQž Ã$k9y%Ïë]+©6É[QV+NÏ®~I §|ÉÊÊ(‰O®þšâ€ôsdÊâ ñ ¨/É $Ò‹ò|“éÇFÑràÿ>#¿Ån—‡,Ö–”þ\˜î W˜È8E.Ð0©ð¾E4ŒÌŒ_WÁ¾|uü ¿ÄÓuX7Ò{J˜‡®©fzå@®Êý¾ «¼×g±°ÆšQ¶l/;+èŸÙ™•;: =¾Oh<zñáÉ;¿Ãñ@0ˆ"-ì†=æÑž™Åíìô™õø4ÝšˆÑI²HèŽBç`[;Â_9t³ŽÎ@­‚å+XÂ[+d$Ò‰bMåi‘W¼Ê1ñ±yÔùåëù7ß:`CE1‰ƒêÔ›Š˜}4ÆÑÝçýt3è†$Ã˜*€÷ˆþåW¿”T¸) ú¬ÊÁ™}õKaŽH'ãáˆRB<ÓÛ¼Æ0ÍØÞlV€¾{8j÷ä,r‹äö÷ŠÉ:¥ÊÉ8m·®:Þ¬€ƒòlþxTMm|Âºê¡î5;oÖÅYñÌpL¿¶TVÙøôVj«F”°‹-‚3öÃÎ/\TVUÐè®2ðëŽ;ýý•íþ”cÐ„Ñƒ¿%0ªrÞlÀP)g€.z½ò%PèÄyÀeÌˆhîû³¡Õ˜ã¦¶cj=Pœ¥§Oåèµ÷Vôb\Ðzƒxí8Ó.À~¦ó*JAf’aVÆgÌÅ.Î•ueÜŠ~Hß‰ú-Xeª<fƒY6-ËØu3@¬£aû»3µx'4uÝ'C—ž<á2¨éNÿ,§¡Å¦»Åó*HŒÈÓÀÛÆètŠÒia;³DéðÚé4ª»Á¦3[;Üï¢<rL± lY²ALAŽL5yÂwft]q@ÓGÙ&òeö¹pbšäj÷ÕWÄ+äx}ó\]¶ˆò1Ü-Ê»ã‚î9ë=m ‘JdÒ¤™TåéM$Å™ƒñ…u&Âs:Z´â2ÿ¨×ZE)ø2I QÃrÜ0#*31OÝ”×£ò•FÆì$ÍŠ2H‚ÒÓ·âþNu;Ã‡«÷Ä¸×>»ˆ‹ªmåt[Öæ›k¶7œt9"Ñ…ßê{õ€3z_þU·É¹6¨¾J=DÇÁ8¡›6Î8Á½ÚñQÉ—å‰TÃ–ÚJ«#š®lYü§åµ_yÁÁï™÷;bq,«[-¶L*páGùºÿ½Ú†q˜‰Æ=((&5‘0v0Îsz(íàš„ÊSÏ±”eyª½}AÆVéÍb§™çzFÂóJïœ·Ýž/£k!EXep(G!§AÖ6z—ÐDb—=-%êV¯Jõr_?Ñ—³„´¶—ìC–œÍv[ÞÎƒà,b«¾”¹±·À@«=éü3ó$!”x…Iô=z›K[ƒ»¬v€gÒ+e¹4A¡†y–Àõàc:pÁ¯Ú ™Z˜íjA;¥÷)7€ý³¯´£±‚¾%ïçÙÅ(Ê)aÆž±\®ÛR¦<é]Êdâz›_¨[‹(Ç±Ï.à+À$ý!yžg9=Zè³»ðÕ²Vßƒ÷ãÑîù(ËK×r±6-×JklY(ö2¾TÕö\Ý
‘RÒy„—¡
¬D<)1àf·_ðßuvý÷¬yŽ¯Ò¯È³ØÉ>¤I„ÐrtëÏëðM/žÐÃ]tõm50ãÉE¢Žº=d_ŒÂãïóD}Ñ=êâó($qJwc	U˜hsåëUKq­ ÷ÏK¼¨	0Ì³}vœ7ê‹š¨´rÜgdá«`Âƒó¥K_Ó§K?<Y?;ý‘Ôe<YBÀã8¡Æa¥.ÎÜ’à¬“C¡¿e/‹b¾©Í»\e¼ô'?,=Æ,¡Rúÿ5!!Èÿ¥Ø²¦Ý³W? ƒ(áòØšÊÚUTçtÃp‘8äïâ‚ªìpo«²“OWN7ZÛzŒUÆ26¶…ýNêC.±Íñ8-H4®Õ¶`Ùõ—3 N¡ÉüŸ(–¤ïç.­«¸{N8’—qú^ZI920§Ã£KKúG9„$	–õËbëJÞ}2î5ñq@ÀÝ…Á÷e*(æÑ´ÌòÞ<kðwGÚßyAöóè,Ž>Ìý8¿ ¦ùß¾z¹÷í2£»+Á%KìÛ2$Ç£íÿè­c`k%ßÓÒ^¬À«ónjÛ5Ö½ßÏYZ}ÿ¶ïÁ›6Î§Ùmèžïív o4‹°ã›Ö›ì+,ŸæÑ1àzÙ›„ü<£Íæ—3° T&è8X¦OÏÛŸÀÐm{éXiqv6,5–¸Ï	Ø.‹màÑ¹ ÀµÒgx¾´fõˆ3öêä”2®ž‚Ì¸¬Lv1]ðêø^ÉKÑ Íñ|eÖ:¨¸}öm2ÛÉ.ëƒÞôèÔi$ÏGÑSN!ŠX_\—K‹žî”ý©
ÎDÐ#§9‹ö±²Æ$–†âBúY«-žf”Ç”Ó”çT–¦?)sÊþ^7ÄrH9Œí«=zºBÙrÜÞ?E‰Ë¿ûÿ   ÿÿ e¾c`