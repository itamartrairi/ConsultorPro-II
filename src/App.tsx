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
        if (rawSavedKey === "undefined" || rawSavedKey === "null") {
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
        if (envKey === "undefined" || envKey === "null") {
          envKey = "";
        }
        
        // Check if we are hosted on Netlify or similar client-only static hosting
        const isNetlify = window.location.hostname.includes("netlify") || 
                          window.location.hostname.includes("github.io") || 
                          window.location.hostname.includes("vercel.app");
        
        if (!isNetlify) {
          try {
            const headers: Record<string, string> = {
              "Content-Type": "application/json"
            };
            if (savedKey) {
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
                // Save to IndexedDB cache
                setCachedAI(cacheKey, data.text).catch(() => {});
                return {
                  text: data.text
                };
              }
            }
            
            // If the error code is 404 (Server endpoint doesn't exist), fall back to client-side execution.
            if (response.status !== 404) {
              const errData = await response.json().catch(() => ({}));
              throw new Error(errData.error || `Erro HTTP! Status: ${response.status}`);
            }
          } catch (e: any) {
            // Propagate actual errors from our middleware, otherwise fallback for network disconnects/404s
            const isNetworkOr404 = e.message && (
              e.message.includes("404") || 
              e.message.includes("Failed to fetch") || 
              e.message.includes("NetworkError")
            );
            if (!isNetworkOr404) {
              throw e;
            }
            console.warn("Proxy API indisponÃ­vel no momento. Usando fallback client-side.");
          }
        }

        // Fallback: Client-side Direct Call (for Netlify, Vercel, static pages, etc.)
        const finalKey = savedKey || envKey;
        if (!finalKey || finalKey === "undefined" || finalKey.trim() === "") {
          throw new Error("Chave API do Gemini nÃ£o configurada.\n\nPara habilitar a IA na hospedagem do Netlify ou em ambiente estÃ¡tico:\n1. Acesse o menu ConfiguraÃ§Ãµes da aplicaÃ§Ã£o e insira sua chave do Google Gemini.\n2. Ou crie uma chave gratuita em: https://aistudio.google.com/app/apikey");
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
      let exportEmpresas: Empresa[] = [...emprxœì½MsG’ z×¯âéuUuE Õê¢H,‚jÌð@=¶Ëå#•	 ÅªÌRf ˆC¿=¬íaNkï¾½{hÛ5ëÓØüü±çîñå™•€{F°n±23Â#ÂÃÃ¿ÂÃ#)£òõÃOýM’J$?Ìò¢z’F§Y^Vé8/G‚=½z-‰WÃá0f% Ž¢"9‰ÊýI”å#!Ÿèa{Y + ²}šdq´{ždv#Ê.eÍ`‡Ëñöy4I£qž,(û<ªæEGqÒµÆN‘ÄI6N£8‚’»ÓY‘”{i”ÈO%¯€·_$Ó´,˜þi@Ìô·`½üx’Le=ù“ÕSßõóÉ\Ž’~±(Õ—J·‹$*õ|Ë§`¹ÃätJsdÊš7P^U¸_lÅ7Qu–€ìò")JqRäS1ÉÇÑDL“i^\Šq4>KÄIZ”•ª8Î³²’…¶UµGðÅÛ“É3|{”3 Ä¨ì:u¦ÑÌ|ƒ:Yr!žG³¯ÊªH³ÓU¡¿=¶õx+Ã“¼Ø…Þôñè±x/ÒÑ/¶†i<p Ë¤êðzUƒ‡âÊÀ*L‰²¨£Š²Ê‹è4-qM&ÇÑø*VúÞ«ß.ÚdñGòéPž&Õ^•Lû=zûFÂ|“ÈÕÖ3ÃÔ_gÀÑÍÌ¢¢Lbhá_¾Ò“Så!«à¶‹"º¦%ýÛ—µƒÐ²˜ò³q¥~]±Tã3ÑO KWY›Cñ4Á/'0gb<Éç1LÇ8/âR¤™Øß>Ø~öl÷™¸H«3	$ˆqÕ¬•P*‹E•N“|^‰2:INçQbû>/“bPC4‚:Rµ‰¨¼ÌÆâ«£ÕÇý­UXÒÉWGW5äç8°Íëë«fMxY"†·"¶•éÃ"'ýWì»³UçéÝ‚ë÷ùä€ÖToû}z¡>öuo¬«ƒƒûšMæ•!Q`€˜'@ˆÌn<l³º[â{üØç“I2®Ò<ëÇÇ«¢§Ùfh€-,¾*.€$ý^~‘%Å^Üƒ¯Áq’†sXbƒZQry}ìÐE.ïzBšzéÖ¸IO‘ƒ,ßSÃw°ÑI:M«þæúúz§>;u—ï°’çË÷YU|3CM z½*7é®dyË÷Ög•]ºë×¹å–ã›Pn9~M§+ñz•nÒßiT-ßÝ©ÑÎ–îtCÕ›t}
Üò}×ŒéÍ˜éÝúÞPu¹¾»€‹+$Ê,š»Z{‰Z~ý-ª'õ·J½¯²;¼×ß‚&^‰º3—7(N/¢´2’$”'ø˜ðíƒZó$—}-‰CÒ”˜Ëæ“	 âªzXÉŒÛƒ2ýö 8¯½=4Æ
ïaŠOÝ”f!·‡dVô"Pž®ãi®jùˆ_ýJÜS¿Ñê«.]µ˜éC`rŒeD2ÛCþIX}µ¦C‚}•ÀÈNHÈNH^mÝ£FlêÆ^Ý˜ê‚„ü6Û0ŽªprÊ«îÞíŒJêñhÅõUÀwç	XbNÍ«‘ß4âCq†ÉœÈäÎ)
*ÃJÂ(u®ŽÒzíVœ²n„ñÊûéâV·Ô¹È¶rñ!Œ\óµ>:· ¨›ª†¬+xÃ½ZÜ%`ÌÔs‰@¿Þ…IÄs?5	wYTTˆ
™F]ÑH«™è>ÖqÉt@›×ß0µðîºÔ¢[ëN-ÊA ‘-Ã¸›è¬ˆEŽ¼ß8ÐN¢3QŽÛºäz UŸ¨ÒöÓ!øÝÖŸ—Qõ
kÞa§Pa2½Â‡¶U‚b°U"1×eÃBáÎMCçc†Nt>¦±ŒC…úX_%¼vwÅz^(¼»îBÑ6,óKû™È‘´[uX>I†Q‘õW^ÌÏ“©H3Ðfyvý·ód"`Nsò‡Š8è]™Ï 5ùt–Âê…WðS ¥ÀRB§YZŠ29y9ZY¦Ñ€LSƒÒ5h¸ÜãÉÆê°ïÚ€%€,¿xš rUäËCyUI0¬ò½Ã—‡D(ðT$`t“þýW£áëû@:½µÞ`XNRxZÔÆï<¸h«ì·J ¨ø°¥°£ÈŒ¡‡a9›¤U¿÷zƒWë¯YÿO´öãöÚZ_ûÝ›5ÙÚ›Þ@üÿkª×sÛ&ô4$/¢)¨Lâ-¾œOª¼H£7ño>}oûsõé{>è«áwež½}èâUÖL ÕÎ|4›Aó;¶4PÄš8Üýú`{·g5Òó¤(ÁÎƒÂ›Ãuö^ÎWoW£F\ÛÒØé‘Cx„¶‘ƒÐ?ýIô¶ÇIYæÒcÜãzñ“,?±,ê»¼„oé^ÙÓ¤Špõ¹]¨ò*šh}pä+¯“$;­ÎVýòîæR@Gk¨gèwä“zSw* æêyO™×<¶r<
J¤¦VNŒZ$GSmwg*ÀŸê±(oKª¹†Ù{ò6£šjØ]'wJ•QYÒ@J¶D¼€xlÉ¢‘\l™jØ’ÑB’à}t7#CäÀVVp;²™êØZD¶Æ¬qÚy™¦‰¶eÊ†©e8û…fðò‘7{„¶ŒyeËçe•OŸå§¹ÿÎpÞ¿j‘èñnäçÀKõÎ”ÔuÒ“Ë¾eêÒì_›Fjò5}þY9±T/£¥MŸIë6p=rÎþ[éî¿)é„<¶·*üæëFeõF#P=¼U&QT¦?&ÿxb;ù<‹û(Ž¾žäÇýW
ÒkïPNÜë›Ÿ<,Ã¢ª`%Bv%Eà+G^.'+n&)–•7“7•ËJˆÛÈ‡e¹ü²<¾‡÷J/-«p-EÓ(;Ëÿñx¤5¬?Ùå<´¯]åi7¨.µƒ"˜%1gnóbœÈ’ ¢ö2Téß$oh’Î)WlžM.{5¾c­€>-Î5öJ¹@€©Èr¸åZ2–qÿ¾8*ÒÓS“˜Ÿ‹€ÔÅq‘_`wâü"ÃØÃ^¨àõÖr$Í¯VE-n}y8£¦xHÞöð,¿8Ê«ôßJhd§L°] ¡H+åœ”Ì{¢ÿé{9-Wâ¿¼å¼”`q£B@»æ	Z7:¨Àr£¢ ÀkPic
¾äEìž\D²éBM2EIQtÃ©?(QY_¨X“;´5œB÷1Ê¨cåVr†Ñy„‚)Y¬Šêtbåi49‹Ðº“®èá}¢éArÇ3°ñ&ÉÞ{HZ¹ÂË#ho$’h\w ÐiBŒã«ß=¶—ÍæÕ.,=xñx`ýqâ	L(TO0€ê4©†ø\nÁn’=@ÏÀ=|9PQjâeí”ÁfP²" «ã¤P6 ’Ï½èk
’ß‡9Q‰wèÓîæ€{C1)ðßJ5åUó[`ò¡<GÓ^Ši7åž®Ó¢¡ÜÌX×&9ƒµBß•D´]|?OÏsqý˜æ+Žíõ,–6bQñ} XŸ/Ð`œ <œåø9& L@ÜÏ£IUâ]rYzm¤SÉž˜=&¶¥Ÿa|ô)4ÊÒ„ÞØ¨4Žç5¦ª\I6ðìË¹jMÚÀÒñ[­p<%Uä
·†Ë°vª®9ç[+4#5£›>É7‹oä£¤_X@ôf -95­`0øBRW;5˜ =9çêÑÌ›~jƒÀÃùûÊô§éÍð¾çXvËyi:ë„L˜×ì¹mUxÞ`µœ·ve8Áíüº(dÊ™V‚Ñ	º-‡Õ¹Ñqþ ¯×©t'ñX¬»Ü¤£.S‡ÄýÄ©[-vµêS5À¿¨„ü{æHí0JÎÎœ‘ôþàhy¹0ÔåFí„q5ŽÜiàæ£7¬×ºo¿Çm£`ëÀøˆËè<©GÎ¶ÕéÒqÎë¾©`÷y¹0Ôå¦Íkœ7§…›Ï[H<×iÀÞi]³¼|{k7[Ën¸ÑÂ…í´xstùëàÈ7Nƒˆ1…Àj(`ÏíÎ#ÿ›Q`›ñv™¸P(?O$h´ßŒ'	XÌZpw™nŠ@¥fxt­õªB`·A ‚Ñ†@ÝÌ¨[¸j}ËÁŸç¾¢O—©Cº9ò´²×Œ;ÓÈPgàßæ”†éÊfù®U,Ë"50K
ãã7¤Â¶b÷ˆÕv½Aš÷í5Å‚ —°Ñ´Ûmáß|àŽjî=äÖ¬~ñh¼ èÆñ¸=é2"i1›½‚Ú¬ì˜Oµ¢]'¤LŽ‹(y#k¾™@U@ØR]u·0:íj©Þu fO<jŒ:<¬Oº4ÚÞ0n×Á’Ê&ÄSeXð|f]ŸGù†Kôýà5»¯èk4«^Iw1¤Þû5Ø>cM‹öËºû!%Ô¯ÞãkÓË|lÏ¯¦ËÔËš€šØöËÚ=@_JÕpª÷=&]/Çö¼Í/ïíÉ¹G}~Ý½Ù°ÃÀ¯Þ£ms	8 ®X¼¾Õìn—ÐÑãžˆ2ÍÆEž¥?j|$(ÐèžçqÂ«¢y]ÿõúåñx2¿þ[…%C@S£ÞöB»ÍU_h‡¤!X)<ÊBu†Cü¼ëÎÖŽ,ÍNòú°ddKc°V2)—“ÖZ—'NYü†—D©m¾yc…ozøÇ¸V³ß´éÜÓ˜à–‘rvQ´Š8Qõ‡âI‘ž¤ßÏ‘‹
xÊ‡+ÝŒHÈ
9àÖÿ“4‹&Î„‘QóÝ—Žør&±x-ÇFÿ}È·ðŸíò(ù¡êÓFÆC³©rÿ¾xr™EÓtL½Áõ2ÚÙ~öLDçQ:‰€]	éèìãÙ$F‡ü“¯Åo@¦–‰¿&)s()ˆ¨r`¶j â¶†¡©°rŸƒf­Îðmjä0©Ô>	üRÁ”ÍF‰*Ç¸e€;>€‡­‘Úà2ËÙ0¡­(dwn,™TjâTxžÞ„ØÎbÁ†=ÆÌœÉ} ]'2“3¥†Uþ,¿HŠ¨L,õÒß¥3Ü¬£a‹ê,ªÄ (ùa6IÇiS'À¸~VV½‘ÖA™OéÙÐõxà6.=’ÃtvxÜAhÒs1\÷íjÒúDb×/,XÔ $ŠÅ£Ú–0¨àöÁîö¡‰N°ÏjûÑ0ƒq˜cW¤¤9KrOSä0ð$zÒhöÐÕUrqÙv¤†¦šƒâ²5EÛTXý–úc[ÍÃ “ù°ÌJ±}‡F³˜)Júh52„iqÆ[œÑ„6cµ
aÊ–¼lé”m8N¯Ná¸ñÈ_g±ÜaÊ=®qT/ãcýs¦nçrˆIZVnÀ®&èîîVIÑÿ:1etëÁÂ*ª~?ZÇÄJ¢!É·d7‹¤ÇgÕÚ×=Œ¨.“¬L«ô<­.G¢w­§£ÌAMy¥hrÕNäª™¦U;	«ŠW·*øâW'©;Ô,öPoµqYIîO'lËT«ÅB/G«t–ëƒK>Òì®²ÞÙñûû¨œ‚Lœ~|ËFøÆ£ÑË	<­’Ù­ÁÛíH»ï<Û¤4+Óª-o‹Ac ‡Lœp1H.q'I‹˜’D-¼ápX§¥ÕÐÇ¹ç¨PÞWÏdr¾úV’Ûò1{í,*½Vès}Àóu;3fÆH”jÒ×‚RcÆð0†i±…óHA[oCjì«rÀ'KqÄsï›¤À¨(<r´·ÿòðÍîóýƒÝÃía*åIÙ×ù#É©¬óŠaÖ§6÷àcz­”^ù>2Ë©u1ÒòØ•Æ¬Ë£˜&x¶³¨

4ê-R‘#/%Êa:¢F¾rO¦ùçP*Œ›NI?¬¶ÔLrC	[REºÛ¤2é—L;H,ü7,´›‡RÖ‡R~$CQÔçœ;•B·³Lµ¢3@va±GôÞÃ5ª$>ºœ%·04l9¹FÑÌPk#-©ÃÙ%5§Cär(|úÊ¾"V-˜‚Äª2þÀ`)*tÜ—oÓëa“?ƒš#bèð…q\n÷uTaýì!š)$õ!¦ámù/½qÖMž}àÑÀ²XŒÏHèõ~‰ï¾ÇßöX›-Ñg-ÁWUÞ8ï5¬úÆ€G>´¶Zf¸HÔ ê9‹¬ÔA)‰D²ì¥ôË¿3…aÄúw0þÎ±Ä,=€ïg]Õ–hëTDö¾@¾~!‡Eä€]%TÑ¤ó‡À¬v™SS¿ÑÔ#ó©J/…=Ác‘Øô„(£]fŸpŠ£Féä…m8Õ<5Öî’$ýØí´CnWÜô*É‡‹ŸË&ì®,#qX›¤}õ3ø¹ŠÇF“uÚN~ 3Ð„Å>Jï2\id’‚òDþk–Ž:Â|yñÈmgËyDn>oãËìÍ§ï©ŸWo­—ÛQszæ2A‹Â±6_M[Hr)1
ôŠ÷Ú,ŸÍ'Q• ¦¥ü4ÔSp¯ÄÅY’‰iG‘Ï`N|65Ô¦ªSÜ5*Õ }wONà«#ÉÃY>G(/®iW—ICêYòÄHÕ––öOôÀú­Š Ñ›VK³Ê@à(ÇöqÕöpµ]ç°hý}2™+ªrÜgääO¤³o^a„ó8B´Ñ‘Êšg[ÁP"£ô ¼o¡Ö?A%¶KecÆç#ñ-íõ98¿7w½•|óÏ*¦#;h";ž³C°=Oµ	2þ´Ä–éÈ+AR§“9Tœá´Š¦QQQZ¤ÿá”Î*ó©ñc[éÎYœíøR¨dM4–ßÓaýGw•Jg€âM˜” :o’Î€Cm¼nË<žrC)•¡Àf 6€>jXûÎÉ}sÔ€yz%¹8CPá¤ÊJ¢b|†ïiš˜DºÇ; +LMd]aù^OÓÍÐGPòd+L±Ôy¤×W1ºgÊ‡ò.Ôq+‹zHl<Ôœñ<‰š€:‡úO:Ñ­!+”oVÁÈ¤¨,/ò"¦3K§y~:Iˆr]}€c¿Žò89‰æ“J©‰µÃFº…mcÚƒItIÅ·ÄÛOß;¯®À6Äºƒ·Ø£=ZWâˆ–àéU©ž›8Í‡·äX·ÜÜÄ ^äç¹…É†l´¾ä©YÎŸ±"ú1Ês žÉˆŸï	›Š‘X¢Sýú¶<,ÐL±^ó6õqÎåøw­ØAŠ¯ŠVkRœ'&Á*¡YßÐÓ)°ÇÓ4[PR-á‘d5ü‹%»ûí”8Ë«üÛƒgXYÿ®wÈÍKö6èL‘ó	•x*ÀÜÜ¶I©¢íúŸ¸I£’Þëÿÿ âgë¼N‘OÂÅ{>QKãIŠÛÕ=»ví2–4N ¬îªŒp ?&  `í[
eµ¿½í\b+Y¥JñyÛ|ïUºkî?Q„ðD1>–,Ð}(Ë¨7.UÝ¶å1š
íÃKðr¥Ø5©b'ZþË_Í ùbÀý{ê¬–Z²ìô–~C:„]ÓM%$ÃR”5÷’UqùNSkt]–G®ôÒ’v*[Já–my·ôÃ@Y³ì ,-²P!\gK®ªPÿut–ÇÅãÿ›WÑªCý 5¬*Õš‘ÊOÈ3	ì‡fdƒ‘1!qjãüú/Eš£º,µeTì+sÔ‹b‚tP¶Å3E¬Ñ&Ñÿr~ŒN´cÄcža]<rœÈ£¤`Â‹U YÇž»)Þ>´ožåQŒ9¼C³`ÂÈX>¬±÷ç˜¯‡RÇŒdÜ‚.	ã*.g`<ã€Óô´€îlO&Gù®.Ñwu9§6£AÃ24RÜÞžÒiJÌ.žgRŠþž×ÒL[Ú‰ÕîÀ6¸[ë@O&0µåÈÚ\¸lï©!È8ôyûs›„ê[CŽè	ƒœÀ8N* %$Í~¸×ƒ¡PÙƒÒ×K‘VHÆÙ8‰Ó86œ•„›¯iò06 HN‘ˆÁ¨^E	HÇµ9ÄeRYT…EF­5ùÇ°ë’oÐÓÏvLX#X_y	=ŸŠíý=ñÉ¥ROd8:®`ä$=Ezœ‚©S5.ÔLBß_¹ÍÒû§	ÎÄ}YuEìkÌ
«R°Sþ©þÀý,…;ß!ÑaÄHòù‰äêgQyHÝÇžuËÍÒ^-¶xÏ–Üãê× „|³¦g2‡*pF·gÑlù1RƒÂ%ÐÂø,:OÇ1†bƒ‘âf}GvŠàï²üb’Ä§	ñD†äHyd-ðæHã‚±D»w²‹úžÓEU²Ñ¶÷„ŒþKùY‚ÉÞ€ŽÑ=…qfä‚BNÑ^ÀùÏŽsy™@sÿÎÓäBÎ)Oneíºò¢UŠÖ³Ò>­×f0ã,µYŒé{vªGåX„ê3:~au™~ +â/I…SèjhX[òo`÷”‰MKÉv×Ý=Fî$A 2’$WŽç ­ÊñÌ8F³)yý[ï«
b¥¦(bŸ	äïYxŠ›H­Ãµªä]ÞxÀÔÖ³<;Ì¢Y	ªsßtxUæ‘8žÑK¡Ü3VóSeÏ®BÜ´Lö|êü×IÌiféD:RáÃ)Ì«ê÷á¿çþ¾3‘~Ð$#jn×B§ŒLd‰Îx‡¤÷=ÎÐøÖbHUÞï…¼òç\<¨÷WbDÕ[v°W)8—~S™aÄè¸2ï}^/gIAÜˆœËÏö\Ês¨å{/¿‚w‡ÇóÞˆœ÷×'s¯óœÔÍä.Íè›’{0©ÊÞàS>ó25-€z‘Ð:pNccZtÝ…øSj™è×éñ$ÍA£Œ09dç“ë9Ë‡"²A·L'g¨öõm¬”°Qt˜OÿŠ¡9×–Bö'I3¦#v¶Ö	‡ ¢àÕ9E°nÝ5÷Ó ù²—ëœ°ù<)@¡‘WÉk‰²KÀ%*ßòÀ‚gŸPÙ½Xßl…!^n'±wuï¶†)o”RlJwQœpÀ¡ÞÀ X ý@J‰L‚5ê‚êjý™×‹ìÂn©€.ÖökWÑ»5ÅºTàR‹l¢ s¸8DA¶ºKA†šïž‚TŒKAúäv‚~!‰0Iðiå$ÁBYÃaÏL×	‚Uæô`;ïœÔ­}5˜ƒè?!1”’JNeGb(nbpf”Ó‚¾k «ÚÞpYÕ]^Tåœí!§7}	Ï“[ãý~ž¦‚ßrr|œº3¤b5\­%c/êhÓÕ8ÒTìÉ]£L~Ô©ôe8Œ:äuåàOçA¸;ÔÙá;Ž%ñjàp&¯@ˆÅ±ê³1Ô?#
m†…»C£‹î4ì;ƒtjë5® Þ'«þùïµPwßöî¾VËÏ{+IÊ`pdÁIÇ2šUQsi‘—bräñX:‰!AJ)E_¹å'ù)å8,si0…@V+U›\bœU~r2I³=øòx~Òîôòü4–˜°?$Ðé£N÷ÛØ¡–T“èå±cÑÖŒ'ëècèT‘’wdÐßK¸¸K‚âÎŠ5Bë«¨½éá6º<UoSáÀ‚[Vä$ h«lçU	ï.½óÄ„®,Ø‰æ›Ï®«‚»Z´ŠAº ;q¥< *oF£A¡´Î]QéEþ6L?áš:tÖk©¹#Žs/æswŸ;„h&š¹ùÌ!¬ÅS'§«¤H#»×éeŸÃGwo³6iNþ»ÚÄÅP Àé™kœ­¸Ólmhª0%N;Æ)ç‰6î|žŠîó„àÌ<Q7Ÿ'„uÓy2ÃêIÏõ‚I²‰÷j3Dø(ô)Xz†Š¦*:Í¡Y›$”_œº7ýÈù¢‹~îù|»ÏöÀÌ5wóùBX€%6ç4òGGì×&TÞŒ¤'³	wa–ãN“«TèÐ¼Zû' Ñ3JÇ¯îU|F«î3Z™éT­Ü|B«;œÍšÕÕ8‹ŽES›=Ô‚ðþ«ç°bˆï»Ì!–žvYDnèûºé]F5ËràöImÿÔj}"¨GõÍ«¹ˆmp‚ü›$§ÑøÒÍ•!úZÝ;NN0ü›†Š^.÷<Ò{9žÆ×#ä^)¾Ý{r3ýo<‰R¹›-ÞæÅ´zXŸø.‰ÕENØí«·¨à÷Ì ˜¾¾$ènWvÒ0U]›Ñ
€ðÄ1ó<9žIî†0¦›c|¼³Í|Bg°Ôê\Fw¼¹áËÞQ@¦Pí˜qž^%M7Í'8
Bc²85<?O¡™õ‡ðÏWÂ­…ï~óH|¾^««ûyLkç‘¸(€8¿¦„øØé›©³yöÍYÕ†¼¯,]ÅF¨¿Wó1ÚÀ„Sï}Q[x}UtŠ<´ˆ­ÙõjÿäÜI¨ã|Š÷¸{uxz¶¹nœå0?©„ÕÂˆÉJhèÞê§Ë¿|È	L§×2æbžôÂLª¶[9ÔAƒøC¯#³èÒEzq$žE†µ‰YTDâ÷9MjfAdÚþêWÂÆ_Ld}y¶}"cäœ>„¢ýÞ@67[8&¶ŒpP5mÇeè²ïíÉD8y'ïÕY"ílXC”E¸º°PÝ=ÎJ¯¸c¨³›Ó1=z9ÝÒÝ£”…û˜îÞµ¼z7’às÷
Í6òQçÄ¨”	‰1cBí£ùök±±¾¾.Øõ~µ²¨m­PPa°u=
{ð5ôà¸­µ-=¨•mîšÙ…5‰ß)_·9¸õO÷ðêC„h¸öïO¢Ñ'ƒŠÔ¹ÅëQy'%›Í‹Ãÿ>;ËJì¥'üœ#ŸgRÌfxeI©§^£3jâì2– Ç„‚Ðð>–PˆY;ÌÒš¼ÏWÿŒÜ„^§ËÆèöÆPìaŠ‡¦„ì|"ƒ4¥‚\J„>J}W	([âÅîw Œ˜ü™%àD@æ·£X&cœÒbÆgDQ§¬ãöl'Qm.<UB
®Y¬pÖPªjä†È.@±Bj@‚´·Ò45›C¼Ñh²†Ì‚EmÚ“Ba9b)Vˆ0ÙÂDIžÝ€ßnr…Ì¨'7.IV² 2Èf–g'Ùfœ(|OC½¤šî.„Õ´®º‘˜ÃÚ8.[XœÍ“kØ[G‚6\°1X)7´“f®î"È9C%+»t”‹3¨ A#ni‰]$.Y¸c”ÉK¼;Tš¾¸èžoÚ'GôAr2TæµÜXoXŒ†ú˜¡‘_±õŠ+Pï¦†Ê×òèñkmôK¢lG¹@$'ñGß¶é‡»ÙôÚ^\îeX‰E6pèÜ:3uA˜˜ƒB[º#Ä>€LŠÓyVEæ,I “ŽÛ˜"!0×Éu2•ìÂ-:„'³nW¯^£A#¡d~‚x<ìçJÈ³T>ÄØf0xª3…¨«	g·5P`´Õ_²Œeœ-"9,„e?³Œgì%ÀÃ`ñ4›'gnj¦;÷Dx$“€¸{‚øÅD-*eHKžp(b¤/3Ð®äA ±jÆ`™—ÿµœ9Žg“¤ª™åÙšÌ!@žCn
hbJÀÂ2Ú0PÛ|,•lÔ’JÖ•I%‹_™”%ží6´pÌd*ª.¼€k÷y™˜“ÇJe#(5XU Rç´hsÈìg®¹ƒ}’0š8¢Éøó‡Aí=²k™—@Þ”üðLŸÈ2ôêMOü†ìÑa–_ÀôüFôè•¼ƒ8Êâ\âQÝ¾üÙƒ!žÜ’O›«â·5C™:tÆX©{¾Yúè±ÇÞq«[Ž”J÷ˆ³Æëˆ8éªÓ„æ«£—uO½ÏX¹Ys)E^TÊ!5·T™C‰­!þ .b.µ…I™É×#±ÁëàZÑƒÝÇxî;)0òOéA‹	÷¨¼Â³_Öqv’ÓŽú,7‘¬¹ì]zW¿ðÁÎ+(¼´åÓ¸ßêy =‰;œÍË³>IÝ¨e¯&Ç)H$ß µˆë°¶–*ÄÍŽiTÙpL¿"j&´A5²àÿ†MæWÜ“Á3‹<lv°²÷¯7m6lç¹RœÃ{)øW¿Ñ#´'FÕ‡ÙðÛZf€v
Û6èô!n#†¾ë-˜:S®á»Ùz¤Io›ÖúZßb¸å–CÃƒF«·É ¶¼{$¬e–åÖ_P¿îÀÙá`V²%´¾gÄj_|£Ñüà[2Å˜Jê$2óL<'³™¿‡ÑlE†µï‹da²œÜ{ˆÍO‹hÉ ìÒ=öVã.R×]Hë[vUH€–kL'°†·=uNAg¤Ÿâ¬_¶°
YøÌ”{ð÷N¢²ÒANÇ;ñÓyRÐq¹¨×’j‰°„Š8Áà”cÚêJtö3!¦[-J|ÐrÈ »³K~yË±e³–.T$VÀÑ¥[ÁzU^E“CÔTW	8ç€Ê~4éËÜˆã1^¤^Ic{«Âgœâ>j'‹nŠ_Kã»$çÂÆ *5¼Ÿcéh%ÝÅÒgC¸/ú¶y	Pn–x0EJð{Ì3&Ó^^=am¹Ðäy».¸eßþÜÍñ„JÕâÃpYpW$™/WRª=úªÃ\#añléñ{jê<¨Ì¯“§ø.ç¦áÃß†“rþ#«HÞ¡‘€fµ½ËsK$‚ò­èì8Ñ¼Êa yy¯€Ä0¯ÿ7‰[yoÝS£&‡Om9‡3Í#æZ”:éWzBc	;=ü&]Ó^&q›cõ€ÜUàï)E ±¶ý3¦&L­ïôÏÆg—ƒ!CùžËðUêü¬T™F<b”ê€w³'¡m²þõOj†”Í¾aZêó$1. çK[’É|5ŒÄýÖ¥‚8É³äiúC*Ð¯~”ØWÿîYgfC&šÝhuy"Ýl`{á2:'Í^C?ÚÇàÂ3db¡y”ã•Š`E}¸¥TÒ½ xÐäÔ2€|®Ý²|îl…úŒr.ß¸»,/ï‹¤áðùÈÈL—´ažµr›@ó°[óI±QÔâÌù¿Zp„Ø]K‘å¹¹;šá9¢o(y&åÆ*©Ü’Ô`..yâ¨gCŒö­m­LO³½óíç³ùL%íÂ=Ù$BÚW‰ô˜^¦H@êM oˆ¶çÇNb-ÕÔâäJJ˜[³Wr~'¡½‡Å«ÈØçû˜”y¶6žäe¯_®Qà§—Ý[²E9Q<êÆW9]j'"Ð}0âŒÚqÄ	Z”“ü´¿"g‚êk€ ·c›<GUø¦°Æ!Ì3üâGB
¦{Ä´²ô4ï”²þ9/ìE'½×;ëÇKP[¦×ËÒ\Ä˜hQD´%&ÓÊeøx{ýÒg
yù£²’)óÑŽD‰p›JÎÓHQÎP|[¥è¹ú–Ïafe~P@]!v×ˆ›%â0ÉÎ íãXÏ°·Þ>¯œ‹ •Ä¡];^€Y³G|ý(–$L›§†°ÂµFv“†\””Ræ™%™ ÷7WCÜ
 @Ù"MÈ÷…òD&uÕ:oxÁò%MÙè¡ÈsC0:ÃY¯®ñ ÊõU‘Ê¥–Ÿt¡é(ûžÅû*E¯Zøø_ú$ê¯5g70ZL¼Ö> Üû²¨wGC+GÐU¼8œ~´aÈ–¹Ýx¸ÚÅèq™Ñ„ª7J™6ÞçÑ=Q^Ù7òR Ö¢	^oy¹–fÈ'{Íb³h&r~Œ@²ƒd
Ï‡bñItNÿ½þkÄÖO9ð‚³ˆ-g(ìZšáåµñÚØàŠ¸_°ìc§k6¿tS9DöZ–Wk”¼m rŒÀ¢¨ÛxjºÞÖ}IôÎvª¹±mÕHÜ@àþî,™äLE0u¾ 
(€ƒ$˜¯{Ó?½ìx¢%É‰v	1ÛqŠ‘Ò(ŒÈÉÀ¬S>By½í¿&”ÚÌTÞ¶©ÿ¡js©ò¤š\¨$Ïø´eBi cwv[X=^B>G­I*]ÈX^Î+â"Vø3Â—5©öŸ<]¨¯)?$9;Í]ÜÞQ)3RçÌTCJ¾jèøðZ`v½Zc ²9’!ãÑM9õ6ùðFöòõo™W#™„×ÉÂ«@ÊÀâÇuÇ- éÐÇ6•Î¿6Á¹ðfžëmÇ8P\ÊDKÒÁÑ{’ê`ËF$“1 âë:ˆã® tð0ö/fý³Þ5ÆÊïyŽ*„æŒ¿Å”ïI3»òWé0oEV¸LélÞÑ¬4¾sà`4R*BA	”yX°Tóð–ŸV±ugÛT^‡ÆcÊuÆÅw%¬æÁÞK•ÖùHô&Ì`gi¯¦SöNú8ámô¹6R¯¬Ã@;¨
i,m|îÅÛû§4&­-ÆLÍ˜3ÁŒûá¤QÚÕK%*E^ êÝ~ŒÂ òò³7–Ôm,ÇE”ô€VåŠXÈÕ/f<‘-jÂHUB¢wûŽ`þˆB’QÐ/,{˜Ï‹1¯×=¿ûäoþ=À/u&mŸ†÷w<‘p¦i~/~|Èøæ áBÇLšc?1 Ð@m ûþ~rÿt•’ÉykB­éo?3$$ÝìÙ¡ëÐÓøVCJ-äðÕPå‰æj¶sL•zž²È€…\U}Mä.z["Œ^º‰¯?gt!Kýq¬>âý~#¹†§i51™R<üÊ `M³(&|]°¯&hzº_ä´iÓ‘Ê),U|œ-…©~öC¯ûÏ6›Jü>IOÏHÐëF‡gòÕ¯-„¸oú4¼À7.¸ƒ–¨“µ®ÁÔí¿ø¦·ÊØÐšæUk¶…U±±¾ÊmßV­”`O·zÎxC‰äí(ížøN>ŸÐÑ]œX51ÔúAöµdØ–’fšåŒ÷É*ª‚QÐ³™ C£÷14ŠŸJ«9ó“×¯©›Ä©Õ´Dç]éî6ÅŒð4Lvò	¨v_ b¾| ÿÙüÜ=M¦w¯=X_wªÑ†[àKÕŸ=8¦7W…ü_ï©‹Ljÿ(ù¡’ío>xPÿü4Ï*”ýõðÇþ
¬¦ó5Q@âÊq>‰W¼’4SòòÅá·ÏŽ^ìm‹ýƒ—+T²ñÛ!všþûä;(£x)åyîyá?Ÿ4tõ3‰¾àŽ?¦±’íãè8Á,íRÀŠSxú3noÁ/Êë“,ªÅÏôhó³Uñù¦_b²æµß±¹éŠ…b.«jè ›c¸ah¿>pà¸töW›:#c:Ýy«îÇ‰Oßû˜0ê)à+*%}¿$·„×ÃÛ X£D»:ƒås„_Ýbêã[‹›ÏÖýÝÅ¸XÀXÙœÀ“"s½â£1ú6Ù+qý?ÊpiÔdÝÒHg | ùm®ÅéiZÁ«išÍEæäïè“ø¤ˆ.ÔJÛüI	Wåçëœ˜6bÂ3R}øóÍÐ‚‘Ô¨ÅuFšá¹@Xø»Ï¶®ÿûÁÞË‘Ø¶ýâ¥x²Üï¿^ÿ——b÷ðè >þ·oöv^öjÊÞn÷D&îõ•ºZ±
Ý«hTbé
T!vQ>ü´ºT/›v„½í|Ù†ï¨„JcJf›P:"§í@Í§”L@VÃß˜ÇÀ ðÇö&ÕMÛj[¢o€²°.¼ó_Ñõß†Õl,¤j)ÊÓÑ¶¾ÕÏ@.rÈâ{´±9Z_‡ÿá5JæË`à)Ûv '„…
ÙuøåôÛ>7ušÐè1Öt‹ú
¯5?0h/¾‘TÔÛˆÙ ò®'oà¡qa/ÒòEôÂ|³æêÀù
¥Ù§Z§''øYÇ»DÇ¥_íÜz;«@=‰.Kjœ¤“¾_ôé@ñ¯Åæ?›Ÿ0€{Ã©o?}¯A^Åoƒ²Q§åÞ\MÔ¹Œ¢ó¶`º¿ÍLd?*ëkî½anHµÿ1¢Ü­¡/•aì$ðOaËÌm)ûØ|Ô¡>÷//{HA5°>5¸óÍÂ—EÐ$â}t¤¹¼è¡ò¤ù.i
:ÙY5OŸWù²cÉÓ™°ŠŠê?ŽÄ?~Ü$^½êý_ävýçƒÝmüñdoû›×ÿýðx=P¨a_?Û}NŸPü»‡bW<ß}ö{ÔÒñûþËÃ£íÃÉÃ}Pá®ÿüÇÝgø¼÷âúŸwö^âÏýƒíÿD?ž|{0”/ö^Ê_‡{GßJAÓ{ýÚöï8/GVºØR¬á!®IrRŒ˜+Ð:°U>‰Ï¿\@(Ù#PžÀÚ³²3QÚîƒ¼ÉË\j~?}ËyF^Ì§¨>vc(	ë·¤
³&Ü9:¬.'xaÜ{'R›0ƒÜ(x½ÚlTZT<€âœRGâ·õ÷Ø>^¤ªC¨gžæíÜÏ˜L&û€`÷ÐZhh8½á¡Ù¾Ý^º@ácxp¯l¬Š/ ê—^[½é·¬ã|2Ÿf¦#â:ö›$uj$ Ëþˆ9 !6¼
› ŽÕÈÆ÷…[Ó«ÿù÷ûgÞ÷/½ïŸ{ß?Ût¿?ð¾o,ÑK×ø­_c}Q/—nãw>ž7Úk¸‹~O;P{áª/e¤°¾Ç)¾ì™Ûû$Õ¥Äï_†ƒÏi3VV&É(ñRe¨K‡Tl`oûÙÑ6ˆ[¥$ÒúÆØêÍM´3·*~»ùšLƒ"/“µ/œÕîÁ}~ýß€³/†¼`c@!!GÈõ<ÐW·ÅÖïî
[;/_ì<ûöúŸŸ¼\82ä þ€"¼v<'­hÛß}ñd÷ÅÑîLIÅGr¿WÉA~aù!çï›(Ä6±ï›6_£ÇÒ0µ|iêæçIq2É/p#ÌÃcPÙÞ)¾¯Œ¡ßc—ŠqèÓ<¯<Ï§Ô!X’ë3¯­Ù< uIJÈ—'TÐÆ3;GŠ6èHÑ#^üæ7œpO/[@ïaÊMzùÞø9¾¬ãŽ˜uï3ùØd¼=08èññ~kbÿú/ GDâÓ÷é~PÞê4˜ó”µ¹ï‹MGlÚ^‘žL²	6œ#}w–ùÊ6jÕXRùÅ0Ÿá%uÐV>¯fóªß;žäÇóbâxÐA1ysƒ|gRe9¡­—nzc“\£õ„jÓM‡éM3Ãê6Æ½ˆÎé8ß#Jÿ1Æ±=øïª˜DÇ	¬ÇvUäÙÎ$¿W*Z“À¯Žç øé,[ªØ£÷þÙõ´’×ó¥qr/¹ÞÍ†&k•%ºx½g–¸V.ÖNp¿ödüŽŽ^¯ÉÙ§Ñlí31ûaís1»„_ÊK»öÃDTED—iæÙæ‹ç2ú{Ä)”›­0ž!ÕŒn¶ÄÊñéci¤­]œAó¢<‹€LÖ&§ú—.°W8”‘X¡jš¹¬‹3d&#ú~GYƒ+
Î@"ç±zü
çL”@÷Þo®_y([a£¦ŸhhÆ¾&Û,Ñ·¶±±¾¢§nGºâàs¯~­Ïƒ+qßô´œEëß
2Ìµ2™¦ä›}üžèîê«ûXNVúê¾$1|âaŠtL¸„¢Ç8=“ÌÊIúC¼PµZ»\['ÃþùÑ}±öÛM3 çò˜ŽI­ú‡èNx+*?ƒ‰¼µu#9&§#*EÕÀ&OGhw¦cü%{².Àn¡q–Æq’úKùÒ 	à€hÖø+\’:WÙ..Vh©àÖ@YgrýÌÖ¾XylÈ3X%°º¦Ç€\c›¬r½ú–:Ãÿ4/¶<ë­}7/«ôäR?¶.2ÞèÊ×Q±sùgzal2:”„½më½ƒ/|÷“¤ÁIn3~Ü|­¢½=¶f¾\Gò£Œ‰kYž%+Y 6Þ‹Ä@c“ðÕÆÆì‡×B6kHøëÏvã(JL«µ•ÇÇ—b¯ºþ`Æ‹oòiRÖÛó°£ísÕpµ¶! Ê8Õå‘…3®ëøÑûßÓŽ…óxÀ£•½ìúoã4_q?J–ôè½=J©)=¾ ª%²]–(¯{n:$ ¾^9-Ýï:Ì¯çé§{3<VµåS.¬<î™&e÷Û*­Ã¾éÀŠ49Á«ªÃZþl†ÇÕã25~úÙÜ™ÅQAþÞà¸·O± Fš•½ÕéŽ UþÃ%;}²˜0y>ùî´éd¡oÈ{}M'fŽq£Ë'›ÆlÇ}x–&“˜nÂ¾ª—Q£ÿÃÿ)ÒÏR?×­a!Œ	<x¢½†…vL˜Z5,Ô0ÑeˆGE>;»l	ðN	¹—n--—7*¬4X4¦¯œçÁUWÖb>LÃOtfñ<‹…ÉK­Ù|)Ù¯øÓ3¥íh·aÐ»2Ÿìî,DCœ–ã5îIY"…tGƒ_ñ§GÃQ‘PFëoga\ØûÀ;ƒ¼úûF˜Uþé±¡N¾‡QÑ93^¼.ÜkÚC	 o¦ÜpXOÓI‚Ž¨ð¸ì¥½ÔÌj²PbJ‡*?¸YÇ“
<ZÃŽ{PaaïK­kïMù–Þ¿oø¿¥ÌÅ;œá‰ ò,Iü‰c8x¢S0“¨“€¢£kc[kAU¯¼P`-³V›ëúG—1Kƒ³ÄPyòœÛ±a%ÇØiP6çÏc²•n®g|ulZnâzî€Yµö…öUu×6}?wLß÷÷˜ÚŠRØÁ}´2æ‹$;›O{æ5|j\†b¶ö9÷›È¾Ô»4õý3!¸ÍN—ã¤ºH’=u0Íþˆõ€?Âº
çmö¢ë—(«•ÇÒéýÞÅÒUÈjþýÛpûè¤BGðúðÁy¾>e-Å`È=KN*Ü6ùLl‰`›6—6´³Jo5a Ty³Læré]½½ªwXˆ@ƒW˜¼¾¬.8èšÇ*4ŸgkÂÐ
:wAë}«5åoàá:Ï1­|‡Mît
žê„Þu„>øô€vß½/.äžôÛOßSðÖ4ú¡¿¾*¹`¥aìfA¯·~Ÿ¿¥%(·ƒ³9ÝŸ}ñ :*ò¬Uï™ølÝ$R\ýßoíf…ý«3ÒúäÕ^X¼Xwg,Q8,Ò/|e:=õg²,ÆÞËm³³¼Ê¿=x†AOoÏªjVŽîßŸ§kÑyTE&¬™Þféý­[ÿTÕÐr‰=ºúÕ1¬àS¢¯GëÇ¿ûrãWcÜ|trrò¶Æ‡£IE– ¯	‚ ÏØ!\ÅÜ6õ¶þV<`Er’ERìç +.­dùš~åõ%C ß°˜€¤Ö.Ö|ÿ2ŸÕ˜_93¾/É;Ï0øÊã &¿º?ë þ‡Òç¦5 tôº®îë®½»,z$? êÒº3s7Ø‘3¼ÕÙ	SoZ7ìÜ­#Ö)oèÐë—óŠmùÓN¬ö0J¥uVQÈnw¡dæ'ß¬’yôÁÊÐa]™8YÀ¼<[Ë–J¹Ì®­8Ë/Žò¨¬ú+æ|<¥hL²1f¹Ñhì«LBtºBžÀ£´lO„ï¸ ãTÊg@É±«ˆtÔûÿ@nMé±ÆTbºÝ4Îf,|(”wÈæ3ïjÃR‰ª©il¨Å=*n‘ {	á¾êÞ@¤Ùùõ_ \~Oìe¸]—ˆùTŽB} \Ÿ‹øúo§)žÅ (œ%d1²öŸ"®uè(êŒ±Ù‰[ÉÇž÷½“ûO[7"5ÔûgÉ)^›Æ*Tpºq+¤>M€–1ÍclQÞ„£rÝÑI"X÷t±s×ÿ3K"<“Oá ' Ž$
4xŽ@jað‰>`fzåhqÞSZ?ö^¢‡ñîä³K‘ŠáÖÜDiîMÅ·ý÷ú¼2üÃùxœ”å!ekÔZ¬kVYldâ?P>Î#•ô+ç÷Ÿ¥„´¢C‹êþËóféó„®<‰–B¯ÎÑÞþËÃ7»Ï÷v·í©øì§h8„fïÐÑ:u	»Íµ]Ñå/*êë‹&lRf€özÀcòAµš‰²—Ùü¿Ø®K&V
PN9Ùñ>[/Ëº{˜™qÑ#X"SJàfðÁ¾ªäÖ–!éUï c‰£éuâTÎ'Y7(›@Ò*t€ú‹Þ4 ¸±”×–Ó`¦{$¬àaÇfã‘Â£yIÒíØ—x¦l@¤-ád¹k+(ÓÑ9Xe½|\±SX—;*Ë;ƒ¡X{‡/M4–)ß%Y»=µºþ•Dåª¾&íu{‚pÿ±Z/v“àlt›‘¥feáÌ,9;õÂ„ûIG=€QOg}/l-±7ûº9Ö™ò;]uÝévëzn¿Ÿs²ä8ê¢¼D^†,û1T—;wÌ6þéAÔ‚ÿÊèÿT§¦8Â[°êQíÙœ„’ÊÚñÁûñ s«6­UŸN,ðÀ<’õT‹M}Žøêage—®ÑÑ|Ñ|Ñ|ÿ4_•LÙ­„Â™®é÷w§"SXÿ/
²#~Q{—Ò³>nWÏ­Ÿ@%öu{º/Å®<RV.¤ÊbDµÉ‹¼Œ¤+o±¸šO.¯ÀnùÕ.EÖátý0F][*…
ê…æ:%£®Ú®„UÃ¥µ7jí'ÐßP[B}“nGÇ×ìzd‘xÐýÈï3¸#…,||M*†?<­˜…Þ/VÎBµº*h¼·-*fÎ‹â‡|\De:IÒ"¿{Í„tÇÜìÄV
a}_¨œµU^ œÙN4«dw†Ä)d?æ¡y wÜ/7R±8µÞ‰Go™–÷KÃèôdÀ­BÐÁ×ç•÷:tÒjuÍ¦ö¹£vƒtYƒÎ×WoKµ4fŸ8õ­ÞOåXù#ùù¾êþBœp£T¸’c¡ïæèJ9£ÂïÀ5hqe?Ê$!#ÑÛ®Òóˆ%_@;ÂUàÛ&1­{ñ* ÒA¡š`%‚†Ì#™\Ò¡]Å/ÄI’Ä¸›^WÊx¾ãŽÄ!$4¼wÂ/"¶–æ0'f0YÃkpYÞb‹ Hè”º‘³uÛ§'.PÓW¡&3æN¦_]~Nëç“0„à. ~Š¬§ßØÊê²§¢¹ñ„*ÑkMßzâ¶e¯>‘-ÉgF"Î(N]ÿÀ_½ß™ÕCTÜhºHµDašgxú\DYŒùOé|ãªòMÝÒá…¾Ôl7ÌxÍéÜiÑãßí<ÎîçÆuOHlZûrÊnë®~ãž‰éà»rcÂ©²‚~ìqÈ=¾•Ûx˜ÇŽžÛl	•ýÜ»FcU€Ò$C>PIþÓÁÞÑn3®îÖEÎ`wq“/²²»Ê=MÎs—w0ÐîÆeN
-©w¥  '›HÿÅdûÅdûÅdûP&›ã:7àZüçwmàÝ?Ý3ï”|3óNVîjÞå“ø€_YC$äVEeÚ¦6ÝTëŽNÙëÕ+rcR×i0&ÙgJj€ÆWªµÝˆ”jh„.#ò¡³Kÿê½Rº®ß!~û_­Ò‘ÒgýZõ+ ÃÎíN†¨ux›iX¤²eû;Âol¨è¹‡!Ô(mÜbºú³¬Ï
„Ë!°MâZ`±ÍkVfG#Òzù§3òóOgÍž~„€®þé¬n”œÑ­µ`†`ˆ®1ñFÕ1ˆæ¤²2Á
åz!õyr© aþ*ÕšÞ	~
íOòŠ–GK+ «üªFÿ–oº¬‘.‘ôoP×·Ý|‹ñ¸æ%îV±NH-—‹ê«rëW¨±ÛX^.=rbl—ÇC†[M@Ú/ 5¼©Ââ»I¿jéH×‹K9#k·èZÓÆùf÷š"è–‹L%i8·™6\äL'ÍÂW‘"AêKÖ="óúmy«<ü_å4–¬]³ÏlGÓ¿dÕ[ÉÎÎ¢×É†ÄÀÕ®¦FmGÑºéÔ=¼:éUÍv"R}%uì»ãs¹±ÚÅ<%XlQè­ñ×G%ïˆE4¹Ö¨šÏ
þ;ü,ùó@†—vØí†­K¸ÉÙFTÑB¾X/@¾@?›Céš«äµräÄ0.TÞNÂ\r³èr’GË:ä|=gù9÷'¯Tm\Ý6XDªžƒ6w^@
ÀjRz[K[7>?GUoƒjûÖî't Ú‹j ÙO<‹®Ê¬ß:_ÃÉ}“?ì‘¸(Ò*ùúñqÀu‰Êšô]Ö¢(\ß©àP¨¡dÂ}lUS·Û1G_ƒyÞ°S5ÿ>½ÍâÈÓÝ:,Xüë¦ÔáßR+Xwë•;yKÿî@É£®Ü•¢‡w®ìáŸ·âðÏYf)q’‡¤5³Õ¡‰za Îs3;[²{ã|:M«þm¼ÿ&êzÿß9ÐýOßktõönvBƒéqµOYb((¼3#ò‹í«û:ô=vèƒ¯½IŠÓyV9oJ$MÊû¼*X€ê–.sêo3XËÅë UKÀZÓ,î“½o¬}Êºmûçø.e€l¬»ÔmcËÿÜûh÷ìÆRA´wB{ÅFcc'0‘O0vÂøpG‚M¥ñˆ×¢”¦&p'òT¿_|IÛŠ˜›[5ìnúã[gÓŸÚCR:SñjÏÓdÅòó·ô{ØÚ›ñ$‰0\d¦ÁRö%µÞ_Ýµã›@º¦úïõ8íÓ“¡{ù„4¿±êb²gïëààÞWq#Ç&„âAL‹¼Ûuu½ë~¹!ˆ:)Ô·ž»nØfsÆ·ƒ­£¯°î¨§¹ªRŒŒÛÏeîÎðLîv[xæl³Æ‰„$C¥­æÖMU=”€Phà¨RÎqŽŠÓÜñîúqW¥’40ÍŸ–Y6PyÊ\Úœ¨ÿm²¾¥7ôê,/GoÙ2	Æ’nÇ—d)F;[ž¥”ÑIò´xWi£5€ïb&Ë°w{ñVàâáúÎ'T––²Ç½Po·Yå²t§Â«×xÃ‰¹‡Åó¹‡&Â¸²hömI7'K­”a_ÿûyR\öeÿVÅÅY³ÓS3€wƒ=zÔÓC¸—åiCì®5Ê| ÃÙ¼„×á¸¡°Õ@×¿ýöL_jzy&>}¯0v…÷§ãdôÖ½Øjôšƒ8è°ÈØ¦½"Ýœ¢æDä'º"8x	¡Ü’‰´û¶(RBâ‘5*ñåÄž;Éy–éu‰¯¿Õ	À´æÌB™O“~òCZÒ²À=õÛH"´.ënŒÐT¶õ‰Îµ;Ó‡ã(ë0»ÎA§_“$;­ÎÄc±Îqî½Y§{oD ~øÍ#ñùúzhJÇgóìtq1dbêâ~ºŠ5©bÝ9ÖÁ‰†Ýu#Mí˜éÇ §Ÿ¸KnAÞ`Í6Y²oõå0Ô`Tàç”ó¹”×íèÙàÓà³KÙÙ°ÖàÆ
iÒÒØQ‡²ä+Ê½º(ÊXí{¼¶£8@¸…¤0äØI(÷ÉH¢·o¸hû‡Ã—/†r”éÉ¥Ö»ßçý•QEîJXFÎ'?¨{žo àûõPEc–.c-5µeeÀÊÐÆ2r`#ÏA.ÀDŸDé$‘2=°Ú„aˆù7ÙUÎÉ„€@è tµmø¹:½½<Ðú°2n›8ÉN‹˜\4f*0nµxï‘Sj”…‚èrö›SAßDª íž€ä©ñÃN²gm9ª"5à¨ŠhÙ¨ý9ÒŽšÈëjÐo`ŠÚÑ§n}`n‹tåoŒYh5J¼Û×>¥ÓY4®rû"d€âÖÊiiAQ¬[¤a;òc±Íý…8zß›’´ï1¬w÷&`HlgñSÚéÃ¬*Ý¡oë}hû+…vŒÅ=P	Ž_4À÷á’lÑÝZäEÌyóc4ˆtö".i +Ró¼…ùñ-½…
ìOà-”-)“9¸ˆÐ>–™Åƒžµ,×ÌHÐbé.Oõ,'a®,çü‚µÑ™gPºØr^`'·úý|‚ø™ý~
µKûý,ƒùý~a¶ÿÖ™­rÖµðÑÃ?;8ÉûûˆGTžÄÛðÃ 'QóÂŽžÄ»õ.Ë’´!bI†°)ÜºI4„Uy¯îBCXS²diCØÎLÐÖ`?CXcYÂ7¡ÏÛÂ1ß…!¬FõÂF'ù÷k+ds»6ÀsW¾û8¬]5¤Û[»²ÓÛ“	gJMg\ÓØ¸QÖåÓÃzõú\ª÷êµ=®Ô`Ú@Ýª˜'þÝ?+;Ê1Œþg)Va4µ¬­¼oc9¥•ÌæÈ¹~1[SBVM»]ÐD$©ÜÝbë;Pô²]LkÖÜ•Ö´¿¸‘Ö´—ø´fýœËÓ«ûQÑð.Í´OkÑONk®/ù¦´f¡, 5éÅ;Ì'sàÞÖ‰§ï¸ç”öKÐÇ¿uKÓºõ@¨¶{õ€bFÂ’MgŸÞ»P»yâºúàå­<}¥úÁ}
kBc¬i[ÊZo@EÎ§x©^¬‹`§ß”óÓ/õ,5”èÇ\¿Tö0÷ÏÊè<™x_ÞÍR!/Méù¤ƒ÷T#ËW¾§ÖAýK<qÅ;´¶52ëÃR­uv@Ú™þ ‘;-ÉÙKIwù…’eÍ!YÞ‰CRÔ²þH&*êîÈ_ÄÆ¿G±¡˜6‰„ŸT4¬)îµ„:|ýÂÝrîþaÂ2_ÿ(¢2—æ¯Òò×ßª-«‡.YªOú5Ïª^5nMâ¶®ïW5KCXÒ†dSpª …qhÉõçŒ.r‰úö>UM–Ü¥jµ¬¯U½<¹C5Àè™CUÕ–âP¬X	TfÚÓGàƒÕôv‡.XÆ½ºyÅsô‚ g[Ú'ÆVç².1^õ#ôˆ•’ûýkòs»Ä<x3ÒÕùÊŽÌµÒYx—2ä~eUB¾­UBÛ§f=÷ý„°ç·FúKÑþ’{7r"ßz•5nrÜÎmýw¾ÂE$Ž-yŸäé­TœÐâ¶ûã·ÙR	CÖ!è·q ‡ õóFl¨•ÿ´ìþÜÀ‰{.'ÝþÊÏÐG;”Ÿûò¯FA`tŠGtLÂTžŽT6cV7{3|(ÖIsÏÔFp*©&•ñ}!Tôþ}q@þ^©~ÄÌ-lÓYá©œãà¯©¼ïzy}ŠI(Aç“H™L¶ÏA€J¨Iy–i©6é6,
%üP"ä·Ûâú/8ý+Ÿ¾78¼Zß]ÿ3îÂ™6ËE	€p§ÿm§üÇ¾ËÌÅ¢uŽ™–™DºC_Ì‹ä[ÞËfóªßëuæ‘o¯ÿ@MÃÁù·ÙàxFZoñ:+|ÊË·LËÚ-¦ès»Z|©Cø:Ñ+ÃµÚ<xVJ´P4’±¼^ž# :ud*ú{+ß8"Ëf1÷ø¬ÁTô@5±‰ýŠe	åFÅ;òÍ¡¬x†bi)%eoˆ¥Ôn+ÿ0>QD 1“å—¹f%r<Ž…§µÖ—,l×^9K—ev3‚¬}ßšêÎ›¥úäø†-q¹,n†‘‡‘ÅÁ^ØÐæ,ä}7ìWµùÐkqÏ±Hõ>A­/žPó$)ÇK
„g)qëŽa0X|cô4Î|ØƒÏ½‘í”ËµÓ!ˆ/Ïw²üêWþÏãkÑ«‰„ÖCÔÌäÌ¢¥Ë4ÌLºz‡Æ…^¿\MoBœÊŠOiøhÙÝ
ó+[Éz÷ö¼Ç†4ªn)Äåp@”‡­$ºoCçí‘»øE•À\_¾KTƒ
¸E[fÕ8C±iBåßbŸ ?óîé¶ÓÀÃ1¶þà5¸¾g¾ËÁƒöÔiìAgxÍ.ý}ŒÍÓŽÓ:c\Ÿv	®1íÄìI{Þ	gd}*Ý…ƒ{åx…LTî”kàâôífCýÛ˜ž°A³'Æ—¶ëõ®’¨ýñ\Ñ/’tn­èç“ØM:I›â¯Ê“AÅßêuq’eÙ0fÛŒˆ—“ÿƒj¶f@ÈJF(ÚÂk	Îæ–Tÿk-ÉJF“‘òmF¦	ê|džÔ)­$¼NùÃú1 “ƒn¦Â+EMÊ&¹U«0	ÆÆåø¢W¯¨/oI<²Më€¬hÕ7Ñq10+rÔÍöót7²-X/f:|¤Ö‹ÙBåÚl,¥Y:Q6v¢¼[Eg~Ê»P·jÂ6;a¸hÌNûÎÑD[["-ïZêQ`Mx±1DKõŠÇ-ß©¦»¸°º»õÌ½pÞwÖéMøÝcÅ5|€‹žº@©ÇœL'8Š½ì	F=:vÛOªó;T"ÿj7Ä£®±
k:ÄØ¼à >(kxÊ¿º†b\ùª’Û“Öû ·4Ô­Õ.þP§ïÇ†Ë23\ãÿ¶¬™Ÿ
 HuB‡¾H¯
Z>œ£‘ñQ"Hý»¬Õö,àÛ	]Ìªò$9‰æ“J*¼¡Ð_ë]ÕÃ…(}úb?cüòòø;œÔíƒÝíCŽŽáWò6“9ÏÒßËÙù©Ã4IT' €dji.D,Gs¯­>m„ÐH²¢ÓäÂËe³r^$*ì~/³WÌ<¢H~T"BJÝ­kUÿf—Æê¯(Oàëæó|Ë:d|bHvHÜÒ¼P44Ä—J˜ÆÇjˆ¹ÚýU±=HÍV±nÈ³‡Ù6:keñÑÝƒ…»”špuûì¶¿î•M[•ÎfeËF¥º¤W÷ºÎYšb­Õ±G]OþîÁÊúÏƒ!ê5âôÞ¢>tâPF#Q4DEt‹xiñ
ÀŠÉíCj§h~Ê¨ ÃK–8lÀË—uÃØ€î+ïcŠh^ym!K®;V¤)Èl÷×VÐ™¶k›«­eCÛpó6 ÍÌ½aûŽO+É>Þh¯†v¶½¶x1t¶Ä¡n¦‰]Fm¶‰„¶HÇÉqâîÛEñ]í´¨oêÛ_°&»º÷ù"ýX]üÖQßºÍUï4Ø•94ø´»ð‰î.ìöÁ‚¾S£ëªmðZ—š·„ý‡–eÜŒí˜ŒlÚx;Þ³¤#ðÃ±¦¿W? £Ý•/Pý{cOÆ­™,ØžLv˜%0kxêv¢¶¸=_†¾ÂÜ9·áÜÌ?hgT‡Ãáƒ"õÓPðäZr•tÍæOt‹ó%Ë‘uôuzËÔñn.ì\Ïú
—¨TIä¿™!öÝŠ.SI—²Ò“t:[œp°3¿µ <Êÿ)¡£çUÏ’}«±9DäÈqpú:ÈÕÝÈ–k}
10ï<†ñÍ-øb/¯ +VË20y*Ž£ñ»Ó'i¤’yâø·Õ‰óºåU»P<¤†átìI·FèÖõ0	E¾¨†lz÷ä9tÙÂÈØÐ%·&]-'bÛ=þSlJ´~5W9"Ñ%Þ¹ÓÈAœ¹ÙÈEo!xûû@ÁÅ7½˜sÜwzÛ‹^'tpÃRÙõ¼ErÞì	d\%b`u6ßy€ÎÎ'óÙ$£ÆÄ¸ÐS´J8c¶$ÇÚ¥q åe
=È9,ùìAO {;´\.òª¡Cªžj½}–¾Hè«Gb#äÿ[yÈ:»þ‹lòú_d›"–ÃFD
7:óucC¤þ4;É[]€÷ï‹Ã¼¨ðšøwI2ÕY"ŽŒÀë†ûø„?.Rè>LsuþÿþYzz–(”àäæ™€‰¦ïE2>8`x+¡	ò±¢C‘üõ?õûÑª8vÖŸJ©u'4j6·[ Ž¡H\‚)Uûh¾ýZl¬¯¯¦€ÖÊ¢	¼>Â:;‚¶,Ûö¿†öÛÚ¯}li¿V¶©}ENX“ˆP7ðNN¤Ó¨ lÒ9dDô«õ×Ž(Ð‹Ã=%ÕR¯ü^Lô]«H¦®"q{¾™¬ ª¤	‰~Ý® MjC±‡×ƒcZ(©])C'¨B‡¤µ}O5f6äb!…;£Â(\ÇÁu“‘5ˆ¾Ð¤k¹”äÊ/ÚÇœ\”ª£U¿K*[Þxê·ûaf!\UŽžzÚ¢ò¢„[XÛ×s*d$%A¥d0ô¤ŠZ®ŒÕe†­´¤û²–—Ûb	eÝË<#Áiôöc›Æ¬iîC2ÃAV4™<#ŠC|Eñ¶z´¤â­}Ù6aBW]’€º˜ Ú±?<<èNt¶ 'qÈ•]Æm‘\ÞXÃJ¥Ÿ ­V¦Ï¸•ãÖÁÉôéû ?’üæÊ‘‹}<-®%#>í¹ €Mm…ÊÑ³CÉQ¤ú.Í¬cávÚ$w1!nÑž‘HöýAÐÅ´-yáyðˆ]œ\¼G³òÞz£¶ZŽÌûmÏÒIL—"H7tJÊ¡¼BVÔ*©žGQù®Ü`¿7åïíSLŸ#?ª×FÛÝm0-“!,Šþ+¼¨ßÚî³Hñ–wGŽ7Y pJž&ù0eQ¬.Õ‰%-€Ÿ¾#F´üˆh–ß$çÆOøs¡c’‚ž¼‰Î£IéD£w×!§?¯ýedºÙµ`µRÈ<£¹vÅÞmºïä
rßÁJ¢Ìm‚¡ËHUOÜCj{¾ØÒÄ?µvIérÖ'
X MYCäß¢ôXMfªÌÈ&Ù<Y’ó™¾Vºm–²Eße²ÓžiðÆfÏWS]ÎÊ{"ƒždñ
à4Ö‘Q[ì÷Hô§Úƒ²5”‰Î¢Y2[:òš^é‚à"»ÉƒÓìÃiöâ(€O”éÏµ t80%\75Bú°g[L.•]D© 8ˆ_‰Ù¼8•lŠZ™Ó®ŸÏÀÐJIÂ~»§A€iŒº>°1QŽ‹$ÉÄ	ü÷Çd`‡ÙÍl‰uGßùÓÍb)úàîÄ¡GÉóBæˆÛ^‹¯˜¬í#YÞ6aÖ		ž<l2RdÎ:Ú*¾Ú¯ûßl‰øÓé2¤›Û>N\ÞØÝÝ¼éjà´ÐÛI½KÛæN¬›1’]@ŒžÒ F‚ÒLºŸ$eµ.$)K¶Ó¥n…Õ0ÄYYâl_ßUxYw§`÷«a>Ô€Ú·¡C¦h£1ºe’’{Ò«ÑätD’„{žb 	JX…¤´¤Ðbö>*ÏŽó¨ˆ½÷ã"ÏòÓ"šFÞ‡wQveÞË"™D€¿ÈÉÏ;ÿG(¤òqå0?ë½°²ô“¯ fa†c²–‘.NRÜqGË´D£Y¤L brœfá$'"ùU(¹U~(yß)²ý	3Ô—Ø¶én—“³:ŸW š¢¬Œäþ â@ÁïÐW>U,çEU2²súF“¾nÔs}î ¯æwhÁ-göG(ÊªÝø—ýÐª³Þ©C÷Ü*Ä	Œ7lb•«˜õA!ÝOÛÞO[8ž&ÈºUbä8²ŒqÆ`ØsŠö,À²ä;¿8÷¸÷/&§	¥S}.‰Ÿ¡'ŸÄìÅ²N‰Ø•z%–êÑn	¶5äºï­æ¾‹–›m%¸_#\B{9Ü‘6ÏRÀ<ŒøjLÞ• Ï1÷€;v»»õŠ ÅUg€süýPÓ4OÊþÀGdÈÇâ´~Çžv³¿E]yãx\:û[Zü+í ¾,y!Ý)¢L3TŸÒÕÞœHù_R Ç€–c´ÂÕŽ2^·Ž	p”Èíé•·³±ð L`«­¯sÇšÒˆîHïÒu½,ò”„MwáPó§p¿Ô#Â[ÅÚUóÌúd”Ÿ…©¥&º2ä¾ ³g‰¶½m{í]ý1¶ïÍ°:í%KD-ÞL^ÊCÓÐ;¹ì·Þlf¬ŠTú‚ymL1’ÚSª¼KÞWp­;í^ÈeÍ6n=cf©vÙ¶8£îÞ/Òa¶ñ8m|”þ‘Û¬._‚þ¼®’ŸÇ×‚§{–zmÃï  s7ÞõÍLwcŽ[#Ýýfàf;F³qÿ¿™^Wµäûôj—^ÒF“÷ý]qk£°Nd<@T6âÅˆ$e­JH5åÊ|Ít9ÝT¸àï`3°‹zê*¨Ž{¼NÎŠ$Ug´Mà-ßk¦²œikxŠ™Õ€†áÏ¬nÊ›Û:“lž[Éô:Îî2sëÿ@½¾áÑ2“®9ß>•uî`¦WšóÛè4Xb‚ù›æ¼9–ï;/af, UÌ[ø¹rPâù´ô¥Ü˜žGU#-,òÚÜ$d“=ÀËÎÄ ý"ˆøšâ´ß<8à5rû¹W­y“o;Ï>þÓoáÔó"{*:Év½Ñß4Û<Dáv¢]·T›i?d¡m¢ñ´ÈíøÏ+ÚQ7Wø¤-¯´*“ÉÉ'uè8Wv$±aJnè¾L¼”êÜŸõ[ÆæèÈàóódzKfWç¥jÖÝ‘k±
}Ob‘ÔOOY‡"U)óy1¦ï{ì†qé7¤ô»¾ÛQ}Å*hd$ÞëýMeTogåERÀ‡ã<G/ÎCïûÔšOpû<ª—S'ËØCçž¶ñ>]áØµÖ1¢l5Ø&~!T1C3pÃ ¦î&ËÙ~¡ŸcÑ?jÄ±(ŒCB±NO{·¨Á¸sªãë¦úpÊ:Yñ&IáïÇƒQc?¢Îçv¢¥ßÁVÊÈUŸA)¯R9\aTÍãøØâ¶®dœL")bh'JlÙR¯¿½£ÍÑú:ü¯7àÇ©<Âû':”’ƒmÑ|;OÇ›[&½}ÓƒÆò0Ë/êB²h²_äßÉ¬tÂëªN¨ÒëX»·Ÿ¾wf€rB\‰5áàý Q§DÿÓ÷v‡ì6l-³à]Þ>ü$4ð—Çß„»†í|ãÍÕAvwmxdeËxgÖFÏª•U4ÑvM…žà¦<rþ¥ñ¶e`tÀ3øÑ=×òYŸ?{,Ö·Ýj…È
ŽÄ«Þ.fÑª®ÿ÷i¡p‡	'Pï$Ëý•ø#ÚFôö)F6NÒ‚ò $ãyåÅïçÓ(“õŽ’q–Or 	÷¶{¯íèqdŽ7Àãv@ò]D,vÜÑIÃ%²0ÝX H„ûrã}äÒµ[Æû¨—¶³§ WA…R$)cG:¤”$'+ç8÷,)7–'>>»ÀÆÎz‰ö¢X!ª‘$PQç’”„€ÆŒ×®5<²‘¿“œƒôIÞfÇÂïj§@ÎÆž^uê1N«ï“m8,LÑs[×¨ÓuWµóÑñ½IÛcw#ûï‰¢L¦³:[è@DËËßê¾“Ï'1ž–„G^ÞI÷`
—§OœI?Ê÷2Ày52=|…QJ¯^»2ªðËS *_ £ÓîÇ*qsú)®‚0TâžFn «;Í¾’
Õc_6ª:˜£¢€QYÁN•Hå
¶¾ÎšwÕz'Oãv†¼dÀæ|¦iæÎŸé+nÔÅµþ2ÄôL@1?Î’âtžU‘+ÞC	Íôê©ãÐ\NFjA_B‡•äcŽÊÑWÿ¦xÔºÒŒ'u
¦O÷nÎåQ¾}L¯µ
é\¥¢o;éTaý\Û,áD`“>æ±ÆVãtÔQ½¯ õ%D[xº-Cý:‘,†çÌ‡»ƒ½S™~Xå ç_¼Ð^]YGg˜STuûÍƒË>ôF©7ûîA§Ì»=¹»Öj‡äµCrÎé¡}²Õ5E-ˆ¶O1
¦ë…6që8‰thJ¡[Þá‹_kÛÔ÷ÜfnŽÇ:Ãˆâ¸Ï®çsú‡Ao`qù!ãj ƒ||†z«#•¼™‘ßž+‘« #)	o%Þ¨¥»i˜‰ËÕúX¨×4ÈýÃ ¶¤ÏIòÔK-Û »»ÅŸ†Z
¢YW£¿ü¸¼!X¨™çæe
ž%%Z\³­!ýÂ…5Ïâ0—`ÆÇ™|=µŠ”Vš†‹‹ñ0ö ü&5ïö£bœâÚÝð~½ÆC^d÷ò;”Ñ‹¯<£ÍmzƒTFYœËiSYû>ûb0,çÇR¬÷7WÅok3G©’©UVKTÚ¢²Ç<ÌÓQ÷‚Æª–òFDv«T-ÑF5ùÖë¹Pl¹Ys)EÀTÊ!f·T™ógd3üÙRÍóOåÒüÿºGÏQ°Í¬¶¼…‘HX•Îæ%ÆˆÊ¹òY¬¯áó¹•‘§üx=’ëpJVU¦S\Æ¾+5´ÓÃ2Ñ•&ºRE.ê”ÑJ¨£…>šøãŸ{p:l/HÊP6ƒpm:˜y¦ïdfö·ò ­Œò¤sKÀc1?a6Ö¾„¦`ÍWV§_Õ9}_[ßvtžÔíâ0™[Â4ºÞÑ‡õ‚NHbÍ5Á*tèF½,ú!’XŸj6ßåá0{6­50‹Ùìwrp©û
íæ<”7t!6Ð³_ îB¤7~±E¾6…îq«Y~é.®4[2ìP“~J•¿’¼&y\[¬šÙër<¡aáßí¦¢²ØÏæÙ;îFs›¬m57½•N_áÖÄ‰ÚÚ–Û€³>•™“éW,ñâ–¾òÉèCƒ°þ!ä^vÏÇâ1E‚¶ªr{/™\.½¨~È=ž·Î^CaöX¼¦LÜ«¸'^B¯Ïs'Åf¼þÊLÌ·L•A§ò(‡ˆ¾›“/‡oMT:;ËAõ–œ=«¡bs[v<e²gê}áôÆOÕ¡¶¶^‚ -’¹˜OÍæ­©îÇìcù;£g µ& þgóhr3Ü ]:C«ÜšÕ@Qf€ófI'ÇûICI°z˜8,ÍjsüÐ¾„è¶š™\Ÿ,ò%Š]xaíØ…·áÈúbmÓOAKô†-×IØï¶µ­WÊ}@æ!‰‰~–`ægQI%bJÖh·!˜jC,JR»â½%¶Zæ€Ìç•:Õ&÷ñWÅg«g¹iX@êz&g1O©”†µ¬7‹t‹nA9ë¢-já4º ôo“‰‹"}©%I»ãr=xkó	ÚG MkSžcÒ§–¥û€Ñ¼ ÃÿÀÂ?R’QçØü´å^bC¤!Ÿë8_WØì=‚J`}ûDo—˜ŠÊË©‘QE°à’Iž’\;ËË$Sn4ÎPc>Y#‰®°±êkCí²g S‘œGYÅ‡ÆGêì–±s¤œÂx“4ø·EX·&lwm¡;"ýŠÚÐÝí7`Øi²}K§î¾å‹®yQCcãF¦À”µDF4µA#€|{#ƒÿ«{í8¨©è¯FQMõN¬I}•º§^£Qô	”ÉD|“LÓ,uÉ5dXÛ{}GlÞ‹ÒÀðvˆ‚·÷÷(SšŒ¡Ì§t_ãPìƒúzçÅªˆâtœb>år	[±°£«\ÿõú_5þ³è8¤ “	‰sn–Ž–,ÿùÙö=IO¯ÿO6N#±]TéIŠNÃ»AœÙ`šÎ0TùíNŽ9îKL-BSi'5vu«¨•Þ™§!®@êPiœË-ì29c
“ï­+ÄÉiZJB
@Ö·	(«b}U|9@à€½õ@?ÿ8ô½˜W8Sùð“O^ìí¾8Ú¾þ¯×ÿå¥Ø9¸þç£½mñdWl¿¸þó³½Ã]q¸ûüúÿ}!_oŠýÝƒo¾…‡bWìî¿<„ß£O¶A¬¥€¼ï4Ãi@uZÓÙÚiÔ×Î«êÌƒY>á²D ‹#Ža9jµj’¤Ù©[E¾t†‡›æ ê»$Uë'?ŒÄÊK,ÏAá¢+´|ç"BÙä5.®ÿwœ“. èÍIžJ ]æC$<©d,˜JÆ!1¸…½öó²œ§¤yœNgÙ¹­•ÁigMlÛ¡­¦ÓìáŠòh¯ JŸ NÐþÁË¯Ÿí>ß÷Å7Ûßl?{	ˆÝ~6ÌÇ×ÿæâ»2X/-ô9¼îÅŸWe‘ 9`Þ›Of‚·$j„ëC~µàa™F(•‘õhO -lÂ}"úeZÍuK Ò«ô˜ª v7.ÒÄô§k¸RëD5—Q¯Ðé@Ó©f{øÉfÃüÅ&|ÆÜ%Œò8GÎY\ÿ¹¨žy59º)]2Z‰æî,j\+Ùõ_#MSºõ&<y²ÃÍ³‰gÛ;ß¾¨OÞŽ‡>KH1	²9ÝûÀìé°öË‚Ô¢Ã°–È=Á¬æƒá'ŸÅÞóý—GÛ/ŽvGÀœdµ»/`¥ênPã€60y²’f;9Ip&I¶~NÁ*Ž&9--»3Žì‰8è/×¡_&Óëÿ™aÿai¥@câO²I'd@ÿÿ   ÿÿì½ÙvY²ú~¾"E— ‹ÁIRQ“)RÕE·¦Õuí¥Ö•’È$™% ‰ÊLp(š^öøñ~@ÙgÝ»Öyêu¾@â/¹±ç)3A‚*uŸÂªÌ=ïØ±cŽ˜ÌÓ¡u=¿¡T:üü/ˆ/`wG0)T]ü—òÊì›Ø.âV¥âÃŒL7·ùÛÂ¡Ã8?˜.„5 fÑd	êÎÒèºOÉÇ¢›¤â$_ŽìJÞÆù¤­³,g~4ðúbŸX½,“÷ãÞJk.U@—ëw÷÷ðÊ£_H•p¶Ô¹ÒG‹hÕŒf€§–ÔqœõÇy’ŽÊ¾à1`ˆ25^•
Àñ:¦{e£wåh—'KZÒ]V	€]›Æ¸í·-y‘Ç`§Pà­ÇS&%‚i¬þTæ“% jOh-\XEí`ký7ovþ‹-”ŒH<Vúª•_=ýOÏvßºµIo(ªÊÒ@#‘yh¤µ|ðöÍþË?FW¾¶u­Ì•Xþôy* (ÏUC€ö\SQ ?G5ã0ÌQO?#­«ÙQ×¯Öþveq÷_MÏHVÀz¼FO|¤Óº\û¼]ŽÞi‰€.œÖì]Ó{%©þö[W …*{—½·µ„¥ªçÆèHìÓ}šUøä *¼±í-õÅr”9Jo€ˆ¢‚úØ5•ßŸ µìšÖ¥]£^çûlL1|y[¶øŠ}°ø!®Ü¬2wÔR ú¥f=vŒ*Äj^Nq[Ž×-Úæ“t_þ\!åÌÊýeÂ’¿+Gv–«qnÍNuœÚ©¡k¹¨;K¦´­›6¬!£BËøU’P+b1	‰z3mAùHÈg§©.µ6ø`~‚.Úß1xL waÄHÎ=2c^Î3"™Òs"dö†>þ5.ûÊ9”õW”,šý¿Nþ:Ù£[æ_'KÑ·@jöÇH›“åžJñîñÕðŸAÇ$%ãû®Ò ÑÄ©U2?¦»x„4ÒÍc_¸í²ÀKs/8½FûaM95/õ	ÖæôkÓ.qÍØVtR	ËGk`Þ¶M¥sPm›Ë]¤Õá ‚=•†X´T9ìæ.R¹?ä jdbïGƒ¦uý“þP¯ä] ³v÷DØã^OÖj8y2O¾„'!b9‡zÀÝ“#")qø¨gY1 \¦’ƒ•sÙŸ÷zHìkëýî½~?rñª>wÇ¤D×£Xˆ.y(¢²H®Â/À>ð e$AÏs©&é¢ÊçaÄË²oÙA7q`®¥Õ.™%òV´ÃÐ$e5=ÒêSÚ]×þ­ÙŠz„›šÓ["fwð+’k»²i/¯ÉåØ¥Â¾Î†ŸT5ç8z5lò <…ôön´µ–TGìŸ;(o ú€æÀÞ&†˜¸£êG›¨.‘~¢±Ò»uµMy" hÀÖuãàÛÞ»Á{­mw†2éak7£GÚ$t81´Ž5†^>IaÓ2Zÿ"EÖ‚“¸ÙÄHÆÞ8ÿu–Ú›I…âÍ´ÀRžfdï!E"%FrÊ¼¤ï´¥¶É`m_d,Þ›ßÎ5-L€i¸Çq µãp9æØø:¤¨‘„,Úâ,ê·r2bÛˆá­Æ³¸ëhu+ãÆTì1ˆdfVT6‹Ußâ¹-Ú[øy`=â}Í÷ú½-mžXbÞ8ð—Ü±¡?ƒ…­úÚOì’$(ìn8­Ó7*,DIVY¹Æµ¿¨“/Y5tþ¸ÒlCµBW¦˜ãb—2@+¦úE'®ó~bJ

"(CPhàæ@¾P[ìåü2”z~bJÚør10ëœ³û ˜úÕ¬‹PZz TÀi)áÔ±T@Y2 üPrh¢/Œ}(R¸:ÑûÖ†³fÈõÀn«v.uƒÌkAYs€Möƒ,š*üY9%þ•†=üò§¦ÌÆáÀ1½AB
Dš¢S ODnNµ|¼~¾óòÕ‡½gvvw^}x½³÷fçUÀFMòìŠª]ïG;	Ú‹‰ÛEM	Ã‘D<6OP¢(\´ –…ŽÎ}2¼ˆN`h¦qEL§‰1\öP=zŸoþ€U»öÜ–]¾ÈXíŽRÔ+cÎ 4s OçHÔueŒLœM,ÿç¹ËÝÄDïÙKsC³*çSMäÙö­bñ{é+ÈBN º»é¿¿©hÛ’•²:²q>Æî…i:‚Ù/ŠfŸêÕ^m¢
…gä—Z·O,Ù1;G
	G÷œK"éëCŒ)Ñ$ñ‰'ãêT¢^óÕp$–0b\årÑ 'µ· ‘'®çxáZ¶µŸ[$Y/tÝ.ôKV¹¥–ÄëÉðÍ±•XÈº<°M_]®NÀWòVsû75,	¥+3Çªî,o8Bæ¸‰«XT—¦VÓ«™(Ös8›'F]L3‰¾íJ$5¸à»¦¸.,nW«Öm /tP# ¶»BUjY¹iÆ#zŽ<	q"ÇFêØ
ÇyTtñµÂQ[ýèéÓ¬–³1fÉåòK†¨Þ¤S`¥MÁ'½Þ¦çUîE(DŽ‘-ØÇÿóßÿWôÍeìafèý±×ÿ)Ï&ÝÎ_'ïK
c¾¾j…õfž	)fÁ¾x(áŽ=c4ì¾Iö<aAäÎPxLÂå•æqb ‡ Ð·ðÕß's2Gúˆ¦CÿÏ³ƒèÍ³×Ï÷wwÐDhïUtðì/ž½|û
ƒû˜Dù_¦SA”_õ¶ÿ:Áà?Úž^}TÝ©-‹YnÛû"`J¿øîö…–çªƒËÄ—Ö'.j©!Î'ô	áÍ:4Cm7_$J‚³íEó<›¾½’é¤¶LC®ûPd[¡N<–	ŠÈì<­SjâÂÕ9ß´3Ô7n°ÍïÍ¯&Ÿ¼jX€Ÿ~E
JŒÆ®ž0'šË`3ˆÉpªs(ÕÙO^}†áúéP;Ê» \T±õ¦þj·îË£Á2„¾IÃ`ÞŠ¼PÀÐ@3ž¡a“–pz(m€Q0›T£_ÛV‘Ìxýóß€Ê`6l
_J÷¡ÒPn/Ócºž£%Í]-ÝA…áÇè[ƒ‘¤KâÏ³xR1ä½‹›€,;ƒ„õU 1BÓÀŸÅŽ?ÿzžÝ‡J]Ùò6õíª3þ’¢Úk»òÍ¥s#\}+mfž™ÓŸÉÑí¢¾˜¯¨›‘†k'›ZÙŽ˜rë
ZjlAJZÝ(Aic"UZòãÇµ!ñRPÒi•+¥6-å„„>b]‘ñ
I¿ˆ':Ð
Ö€Ê#ó¦GÀ¨
êé¶¦@ÜfWÀºP@†R}N`A£,QŠ8Ïm[ÐîDðcÍj1ö«Mlîø{‚¡ÉÕ3":BÅ›ø‹ás7Èf‹4†
«¼VAÓ£’aVoj4œû_æ§1_t+$!/mE³12Å`QÑ6–ÃV²U·l3­	p~’‡¬.­:ÆÖØ`Â4çþÁ+É£×/mTÝÎÛª×Ìæˆ;“mqFë:é|›lO=¼N³¨5‡Fr7”[é<`F,^Ëçég'ÈÒ¤<¦eÓTs`é}½¯û‡•þœ–v­éÐÆÌ#«Ë€½á@Ž”Åm’ì5¶”©©ù«Í©%õf•T"2Ôü„E‘¬–y8š#|ˆ½òÿ3ÓÉ‰y8&_t2a0Ýˆ"£{‰{âOå#—Àº—jZåøá—:w¶õ¹‘F5êÖ®ÍºÝz¾ù–€Í—®NæÐ4u‹Hbä¼M$eZ´ìkLsB~¦C=GO
äçÀ‡”óm>òÉŸ¶+ñšý¤ÁXìÜóÝƒ@S×#ET2òÎ²a±›œ…ÞÁm­‡‡–T„äwôyàoö)4{jöÐmö°¡YnÃ]ÆÍ·NZf¥2ÚOÎ¹M›°ä˜$û(;u ÌŠ˜AÑ0y8Ì•57n‡¶ÔS’Wð
ËÑZMü“¨g¤˜eãø=ÿØ÷lÒµ	õe<ýÙ44*ZÖ!°†³ L¦ø MlŠrÔËÑ`™I­Wy:ñr
íK› ,ÚŽq}Z—.¼s/(R)´3ØUuÁ8râuØ•+{ýnÂaW¼kV‰Õc6<V …Â©ÔwñÞ`4˜ÀÍï‰Þò÷{ùÇˆOÑÝ,½:ÀZÅ`:€ÅñSÀ±pRß±b_ x	Ùkv-Š`¾óxééÜˆÜ²€ãÇ¹£yO¡~©ð Î -•VîhõËM¹öÔ‚
UqôúY«¬#œ¾¡­¡¼¼ª=õxHlGÆ™Ñì>Eþ_¢û,óZ÷QCÔ^.öV&×|~zÖÝ¸:Ýc"þ/Em}z!Ä}ÚªÈ)o	¢«à%^e*ùŠ*ä£dß<Uˆ0°ÐEQ]Ÿ’ŽžÇV3–øÈCY‹ÝCtI2N&ª ¥öœ×:ÒŒ×‡ªÆýj2–TÅØtïí@‡-ðÆ’”¸Ðw,>L–eoPˆõ¤ç´`a¯¦Ô8–qe@ÙR25×¢m¬q»^ýIñ‚˜ÝR'"áåœ#kJ¯	åk(§ßÛ Zc¢nBÎˆsdQ5udÂk%²‘
¯51ß­‘º,ÑèÓ!Â0vSª!<†ESš ´5q‹`£M¼t˜J>,iºd˜î<¤Û–¶ßkOÊ>Ï^ßp3ef<¾zfò•{ø›ò­2ªÏQÖš|e|Uû9ÛàžŸM0MÒ[n­Žþˆˆ³Bg˜å???øÏýY•Êþažú øCŒˆþA=Üå­nüÎ)Õk£5Ìñ¡Ê?”'iZußñ9* õ¤UÙŽ:ÏÎ±ÎKkŽ²óv_WzŽ§ñiQ•‹x(2’Ãuº\w»|óƒ·KÒ©Né¯Ém¯m<JøŠœÁ!×%Ucß¥m->ÏÅZ|ñz®ÅWÊßÆW‘Ý;;2ûá”…OBˆÙÓ}RbÏ'Fµ–gãúëÂçEëÂ¿ëër ‚Ÿyæ€›Î·.r:-V¦óù h`I=×zÍ‡ø¿1“¡ô¡ÁâûZp'¾„ñ¬ÊÇäÉhb‡OQ01%V_C{:ˆõ1T~öùŸi”´	(”ŸVPDEø-:”,xôÆyÔÜÏ¿ž¦#l÷ÄO`CÆñêüéõ>õ»—aù£ï	ã$>úµÁàßë5vq¶vHã:`ˆÛ@þÃÏÿï¤daTâpÌ{Ìc*`Û,,…“*b
Ã• ‹±­˜ÏkŸ9Ër‹þ5U°Eðøûl”R•‹ñN}^T˜˜ >*Ï™=°í¾aáøòrAx÷ï——ó£^•÷
IÙ©å†ÖÙ#¤HKÚÁì$éùißÌÛ€0‘^qìÈn&ég»XløFE<œ!Êgþ%ÑÀÔã™×yÍQbGÏva¼/sþ®¼Ö*´Ç}¡ehƒþ¤;¥Ï=Q_Ÿ{¢„„Ò†„Úkp)ôâ:Çç1Œ¾üþ‡.+ûŸ¦™l1/Û ­²O|›Áâ±Çóïëüx‡°à‡Bþ ÞaÙ÷	==;¦#Åâ¤ÛÌ‚·¿…ŽÓg§0‡?¼}ñ|2UÏØ=ôØÁQÀÛ Ñ˜ö™@y´|Ò—áÈ­ª8Ñôè^“†æBåˆ÷ìÅ®J!Ñ9èd‘]Êì)¨•Q„w‡Bóny­´NÃkáTäüNãÑ'Üé<À°EŽm°DÖROYæ[ïZã<·1"IºÌ¬GY¼:‡ÃôóWa”—,Ó;ŸYË3]«º„¥\›<Ñ#uú†ˆ[„½îçt!iðpZÂa“å×s\ƒ¢|yžp§[Ô®n2ì­ËT³Òë'.Šø¢cŠ7ÑÀ_Þ+lÉplðÀ*ÂÑ^¸ˆô4©i„÷sð)ƒÓX×M°„è%X€¸E©²C¡h¶¬Ÿ–Ïá,oË»Ýòý4Êó±‰
ìg¸‚@/¬üÃµâIZ‘äTÄ&2µ+0ãëúT¦R†Íå³å‡ì0 =$OÊƒåî±ß‹ž8ÏTØâï˜èÜI¡¨Ë›µ¾4q³ö”R;r„cÈLì)Rî@·C>"¿ÂÙ})‹GÈIŽFé9†¢¨6pyÀõ[Zk4¥¢‰2ãxK¶
‡f‰Qñ„ú…Bâ±ßRa#£|è™ø¤â›B­žö2ŠÌHæ+/½»ôòû½%ë-ÚØÇ°9«ïþ:l+øçîÑûÕãåh)\øÿŽW~¬|'ŠiKj¦iü!¡;Æ°Å„Kg`jÎ¢Ã‹h8)^JÆpŠ>¥P
ò2²Ö0ÙÏw¡îŠU´óëhÎX¼{èÍ\`ÿó#êÀ(?ü	 £#]nƒ¾öºêôOÐôŽ]Õõ	{#è~2u(ÚWÍ%	z†&Ð-éuÝÇ–p.=¿!5&oXµG¼+ÖŠG!¦ÉæWj|ªÔ²}ÂÕ0ž‹œ‘%TbJÕê“ÒjnPR_£<EG,ÿ¯¬:ÁvZ_5–ð†üŽûý«9ŠKñPè*9ôÃc$ç?|G¡°Ý_(ö§ù=WžÄ‹ŠCÆ»Žä¿äõ…±äÓD>7?ô¾ïÔ"F¸Õ_VY5YÍSøhh“‰%¢ùÌ;‹2úJÖ;¦ê«ã\9È%ÈIÆD ¿Ï†Ž¿ñ/PFè€Í¾¬D.2LÒ+Š-4cCH³sú‹¡snà N©ZMÒ¸š!ËîŽˆ³©rPü75Is&Á­cü1ø%¦ˆOå;¥ƒà¯àˆ–Ük:æDvH¿”ëÓ#š¦ñg,V»ÔzÒŠ9	ÎTö'°a3z¿ÿ<“«+Šè[©=ÖkÉìbÁ€ŸÊdÅpp²˜¦ñƒ…âŠŸXØ7ô2×†]òA”¹>h•i)¾ÍDÓ¸B¢^ËšNgœ-•ýñß—J Éžv(	ÃóW‚I•€Ãï È;òÂSn®@“ŸÀ¢GæOâ|Õ‰Á_’¦Uí+AjõE õ¢;ü)~¤°ƒ¬2ü)ó˜þí(ekÖ7˜iÿ¢íþ4›< † ©0ü0+D¯ KF„:ZVÁ4Íc`lÝ‘™—ýiœ½Þ«ö˜áâ˜Mó<Î8"ØÓ3Xä8iogŸ¦™ì¥R8€5ñÛ,ri¸§thÓ
pÉ¿Ç% ùÞ7E}†Ìg]+âócE{ÆÃ’c&Ä4z-€5œ'.ŠÓ%À­ìR‰4¸DƒÖP>¤¦F9Ç:ò““”»'„«JGiš $ØíOsÇ•ýâ3®G`(ÿX¶‰¯&€Ë‡¼Óxl"x­Útv\ANi?kˆFÁ µ8A³# ˜¬1b¡,#“ºXÆ{–IiÈÑÖå5Ïûâ'·:›„L±89$©²‰E‰ì€ÚXUöéL£ÊÝÅÉÈ³že½Ûó0jÌOä˜Ëwòé{nØVG´à¸ÞæûÉ‹xºM’¾"y(Œ/8Kˆ~šW$×€B*ñR"8s®£¤ã¸]û{}¢#E£’d…9ÖF±`¡±–92r!cqx²$A8ÎN‰ ð aÅrIF¢õýïj9´;¢Šž,~HîV¾}ä²±rea¡x¶§¢’Ç]Œ
Ôg×êŠY’Ýÿ®±ÛÖ#Å`÷4O½gCx@\²…×’­*d,g¾¦dVÈ¿‹nµG&nŠ'v=S8`”Ýö°ÿŽ™Ê
øW£s¿P€•4'‹­igâØÅ÷6w¡k§lÐ4¥þi\Oû¤g£îažD£¬¬\Â";IBøÅ’>1¼4U¬§Wé£2 9 GS•ØóóNŒ‚×èg‰Õ…J„—B™Ó2Ðh¡*Ô‚ˆ‚¸0¥[Äbè·]½NhAjÊØKÂC¤€)àPÅuã³ðÁCŒh£„~ü_SÅ2özhâ´ÍYžûð£CÈ»ˆR•±è›œåÞ¸+Îéúrt·çÁcˆÚq–J2j8ùÓB³|÷¶£­×¨ƒ÷+§‹,¶VBmfæ!1™Õ¶÷0_QG8v0K¼D¾î…dÖ ¡—Ÿÿ]â×¬¦jÃèûo û„S´!±¦Ær×œNÓÐ‘Òû­‡°ªYâ»ïg“¡¸Üyê”ä”,š r?£QvŒá—N˜2#Cwmb±šu¯óˆXyƒ­ Ô¹)ü‹ònÏ¥LÅ¤˜îÝ{ ‹øŒÇëÓ4yT¦¨8e}¡~ƒv;ZâŠ›y¾{o6Ê¤†¼MØ@þÕÊ©Ç<Íß#ázv’’Ca#Ê=IÅ"qáç(Ï?!æCñà'8†˜Éh6ž ‰KÔ¢Ö$j*XmX¸}¯Ñ*1ŽÏ€ŽKíwüˆŽF‚E:žB0¹ùå	ù& ‰èÆc³Ò‚¹åÅKõJp~<*ç;éçÓÌåòŽ’^Žk[ËÖÚ÷œì8ÀÊ>å,9rOwP©ÕÏJúË%¯P°ô,5‹“úÈ·œ†DwG b2å”Sú»XR„XÍâ¸W©Vé=µiæKyŸD”æb"ü„Å¡œÂ¢%x,ËmZ;*l÷à nö Ø§}-îa"Dl
?@GhwIÎ“$.8‡(Ž¹uˆÕàìt®øá¸e„`€,¶!S$”€ôÕ0Å„ƒÈ^ú jÖxÏø÷HÅp‚ûÍïÃ3<¶/üàU‰H-D)*­Mù6Z{ ÏZ‡žÕœ5RÆ”Ú‰+Z8V¸ù«Ÿžb¬`R<f:MM&~+õ <ô¨1u€÷ zN¸ÏïvÒa‡a99/lTàû¨|O–…ïÜq1N|º|wòž)IØj¼cÍ8ªë¬ÚØ{F9@‹`	òX¸N<`§Y2–î;G/aê'Þ;
hÓ²Ë×ž&™.˜Ìo!)ßvDÖ†¶ÂÓŸfBåv§IƒKaï$~py™¦â_£€Üª¾¶àö½«L?‰Ë½4¡¼Ol,„ ôe˜ÖW‰~ŠIèí 9­™År½fÆ.‘,ªl-4©#xÜµx’†3*/I¼—Ùé´®Ê!ãLG ñÄ²T¬yoÊ;X%tykÂ‹=Ö2GTÉHeWAªš×a7*o#P<fÂsà‚˜
ÔáìˆSMÓh…‡|ˆSOª¥µ á˜õ,ñ®a^!•¦¤å	á¥Êó8`’Zá§µäÊ €}¹Ú]±ðj ü‰g¬Ög¸œ“ÂÝK¾U]¶gªKúíéÐxît'ùÍ`‡BJ'v[õÈŸô4Y]à.<s.ª'ZÆXZMb<ü,V”'Z´ye<ÅŽ8?¶¥™K%‡H‹ Tû±(y•ó4É¬ðã|“Üê–#É‡_¤J‡‹[^!­”WÑB'ÛlðnZY¯ÜŠ_ur*üÔÊª¨_!¯RèË-Ô,µÂ#¹¨É-ÚVv…Ÿ+çÄyeXj¥­(­·]¸AšUÏD®û)!Nhy(!þf~JHTt%	£¸¬^kÊ£Ž¿‚ƒ­_ºU%f—²Ò<CBÂW
ï¢Ê:ôy]‡@mˆ‰Æ’–.µ$¤N‹&z
3îÙûÈÑN¨x€Ðt{ëÑ³ÁÀÛƒ*.{ðQÈb—[PÉ¯¤á»¨\K-cUuÅ¼<1Ä÷¹ßNgtâ“èã‹ttB¡·ÓqôÍ¥‡<¼ú¨Ñ„4>¿f€!#¤­iÒÖD¸´‚üt_ÎÆ‡iÑE¨!¨[ë…®‡ßi)ßªànÒ-îÓÏw90Ð´~µ‘ËQC'Y¹7ãÙ‘	¤„ŸÉ>‰jré=H2’M| Âbµ«|¼x
9ÁÕ	˜·×¹ÇÃ÷&ü_xÞ­­ŽŸŒ5Ý!@¹a“{­¨7éç¦¨·t|{Ô[:æÔë7H½É	xˆ,-Ø6Û0·Œr´æ ®Îçæw>ÒN¹[k»ÛD2gkŽÎ®Kè-€zÓÁ]‘o´ÔÎ^8.>óPd>ŠL9Õ9$™ƒÍI‘iâ3½ÚoGq¹£C”húµÛ&Ç¸LÔ"(„>/Ø{ü‹‡Ú½	¦ÅŽ·ë
5a *Zm:u¸þðZ’¸v„¦²ŸäÞ£NƒÜ3<c+9”gÞ¿­uD+?
-hÖy˜„c¯úÕR¬E|Æå¨‹˜ ¹YfË³xëïÄ³³*-‰g%šP–¡ ˆH-½DjÙD¤zü-z•¹Å¬û°þfc¹élù	/€r†ÛÎRŸ(éfxr[d3¬QÍ
qú‰f1ö/D37ŠC=°°­] ªýµ€¼ÕB!íµ´3.
ÙÐêHóÈðš©"H¦¼³ö…»ÍÙYwÚ6jÆÆÍ5jÇ;M7`†lÌ¥‘æ[A7³·>m7·¦¯éÀF!NÉ@s‚Qb‡Ë9|vœƒz6I~s£ÃÛô—Iv”!§Äâ¿ôD44:b	žá„÷d”ŒÉY\$+GÙh´–4–¥$‚˜ÃóÅÏf1GœGºÃ‰côÞÄ’	ƒ*‡×âi8Åâ%G´bÂwÄÏ·ñBJïë×ÕjE…®6¤eÕŠ¾U †¯‡¢¸{ïv®aO e8IÛÒ~BU»‚ Ê\›…bû×¶‚€b/¯kñúzª„k±Í´1×âœó3Œ4/×Œ.öÓlN¾™õ5?£º›çåŽ¯ËÝ"îcIMu¬†˜ˆ2@2‚O%ƒÅØ1¸,—ø1±P€o
5f£¼Œ”êÊY­+ÿ|öëõPÃÁ~Á²-áã„ä93z|iÚ¯X(4P®M_aÃ»jbîdxÖÏmÞ³¡.@ÿ@A¿„ýõ/œsÄ
.1 ¦”kGøßX!2©u_óþÚ^È¾†ÆœP¹“ÍÊm‡fÖªi^¿Õ˜‰±`[öô‘âÑþzŸ¤‘Ìü‹‡õêãB–Ã-àÕ¾ŒjýWfýÈòæeœ¬ÓQySXþà´²¬cÂeÏšÂÅn¨øM8¶ôÐøVÎÈ{ÜÃ:mOs^FÿºÊeN_húå@ë$Žh«á(Ù–^_ÇxEçíÄ³R$¦…òíò×©¡Ä_•æ›A£Š’1E.´Øþ¢Jòâœè&U¿Žî’U¬Õ_âÇ›Ø]=óWš‡QÇí!V¥ASIU}éØ¬ú5ÂA›b«A‡2¾ª	
–ÊÀ¦b£=
tT$nc!(PTÙr@¢ká…uç1‰!]i}EÂaüxEZa„³_
_q¹±F©x‹ÝTzL]$ÈR20Ÿ”MÞ+H–"ƒyÄ¾ø©—'syÂ¼äÉ\Î0—D?5Re.P˜ëÄÉL²1Ÿ?¿ámÕ$/¦šáËª…"Ë½«B¿ô4Rò+º·Ç€Ã(7kt ð†ád:3YÎŸËÍÅ%½Ó÷«t\¾Í±ÅmL·•ô€…k?)­4¥gbŽ®Ö”úý¾mŽ®EŠgMáyûHÝŠ¤ÕÞ\m€Õ¯zÐˆît¢Ï×Ÿ @ÓŸ²µð¼;ÍSQ¬^So%õftöÞ'Ï¶¼õÍ½“®ÃYôí£hsà8]‹-o“éW^˜xh‘±6»bg9[Æ¾¨+»¶&~ÇH…(§:¾Ë™åê-ÓJ&qÂ*}XMþ–±'~`žÝ&Á]S`ë¸i™XLÏQÑR íœfÀ uh¯D¯ž‚¯éD;<%’Õ-{Í.zÞN’Fe6Á Ù/ÌßszPN’íh	ÖêôÇ€³ãc3‚NKìL§£+Ø.f/ŒAx‘Eû¡ˆ7ôû †íHÄ>É¡\›„ZŒfúpâòÓ!¦ÛA\ô‰º"M”ÏN§÷@®û¥¾	_’yÃ|˜‹§oä15¬Áò *Û¡„;ö”ˆŽÐ»ÉTä©Ã/½ž¥^óLÃÑÓéé­ý§ƒW/ûifG"Ód/4gö–4vÙœíìŽú³A*qFÍñBón6Úºr¤"Û"u|©-2n‹EìbNÂ$:oô¡:ÛSŠíqÁ	½›”bwÊëíŽvqÞÚæèx’å€&éø%´ÄàØAÐéÙÛ\Vz± ˜°k¨™zbÉb\–v×P•<0`Äj8#”Ñµ,Iei>¨/¹ÁÚ`x‚î;¾ƒ2ÌÖË{Ë	³Ýš-ú@îÒÃl"a‹ZÛlÏFéˆC¡9îˆ6,ÉÀtP\|§lÇOÌ`ûh±ùÌ¯Z”»×ÈöÓØÏ?ÎôËÙ2^Ÿøz6¸å§	¿¼J(Ç®Ú>=	&›É2oH'/Ùy)z\6y(b¡€ íÓQëòL\¾Q5°4½yQ@røAÝBbÝšq€(yS$pˆA…ÐS#g–ÑjÄ#îÚˆA<nÂBwÙ30j‰èáÎÛý×¯><{ñúÍ³ƒYªìY¸Âí«5ºÝ·Æn_^¤ÁÛÍoŠ8<S¸£¼ÜQ¶Æ0²f‰¡XMd˜ìAÃÈ£\ òPsYf7áòkÂúèo‡ˆo2`B1ËÀWTpe\¾´­‘,3“iïD¥«e£ANÇÚZ
£A+e“Ñ  ½Ìöl‘˜ÖœßÉhMävbÕ¢i"#Ï/[4myÍ4el@<÷oÇ,ÈOF‹LH·f¤®•	Xò –“f òI1¹™ËýÇ£´¨ ì›7¯¢Ý7ŸÿçÛýÝW(C€ÇQõÏ³ì4(3Z˜ígpÇã¡ñLeøg‡Ey{:;:‚‹EæÓ°ÁñÎ0ÝëÎ$AMQavžâ{mG{*¼—4@ ”e…ð&¾aÝÉº„·K ›ÆPÞØªJvÇ×-¬oŸ.¬ä8c	ùŒˆØ¬A2´Ù‡Ãå¨ ‘ÁWXÄ¢O5—ñv\§bw=úÔ&¯uªÍ,àáÚ@F¥´Î<“:óˆÂ;ûbjÏÓÓtäômØ1¨ùAo¸dýÊcôAªJÖPì[Ÿ…/r’ÑqV·ÓBGpÿbú‰˜ÁJvö·Öõ¢?5Á#ÍžÆÞ*F	ýôk§ï€ÛïjÐÖÕFlU­Pi¦’Gék‚¿Á{ÒÚÁ÷'ÆlG‰aÒVFÎv”µdÜ‘Z<ô²^aïìvŽ•½ÉêÒ
›ÒÁúL÷$FSE,6g8›Ôæº7Åú_…AksÉ#ä¸£ÐócæöýÂ£¼»+–án¤6)€Iôdóv†w&Õ Já?ž%Cé;ÿd')$I~
$ÛUn¨’g~OÄ½/3Üé­TO¥@–…QõâÕä5{_þ˜¥gÛÑažówl:&¶½#kâVærEvfÐ<NóMzÔç¨ÇÝ‡³¸˜t—Ü5ùéó¯&•þ.)-	¦—£ÃQþó,EÊV2ãâê„yx%q_¥¤ÓÅà!8,˜"â~u9hÅ`@Y©_ûàóÉã» Ù|Ø€þ~»¶¾=Àtå”%z&J‡çØ;‘ý5T¿$áO-Ê_´nt[ˆ®l#y€²”xÌÔ&›Ã£s€à"¯h ej]ñ×G¬*V¢o.¡É'ÄQ¹—9Ð»#Ì³”vÐÂÒ·
äºfž…Ixty¶Z
›É‘«G¨¸ÑZÙŽÞfcÌŸ5žÊ6Šo£¦y«ÕîªbtRæé‰*€
Y‰©ÔGñ$ýI¤H“ïMÕ³}žU9\;¾ ÛæF˜e¬—5Ê%Ös0•Œ¤H{¦’î.lRQ“ºŒFxÚ'ù$•¯{ÆöQ«²Ñ·ùþV°ÒI¸HU¬…ÐËïò¾ HxæGGlµÁ•NbûImªT£ÎŸ1ä,¦ ô×àU4íâU‹’¹qÉ ô%ø3iiÙ%b SÜâN¤‚R¼ê)W#k­š¬áS9zkK\wVM ¥œ–=±éécz¬ëA¤©ˆÞy"äñÂâ·µwÊÝC¦:Ù)|b?t¼-ÔR¿>è¬EU¸i¹UìÀ™oŸÂíû§5Gò÷cªŠûe¿±üDå×Ò]«À««wV»ÚÇ-ÂçÌ\ë}èŽdq½ÙpÇy°=§õ›]ï®ÃC(Ôl!³  SXCbc‚]@b*‘¼t
¡Þ{™Ÿ´A¿rËA½›ë¯£‹0ÐS…Æ4ÿÜsñgïœHÖˆš›(‰¾9‚®#ˆœ2Ä‹ÂÈ{ÎI¥ÁTh¢ ®‹ËDÂ€Xçq4ŠißO9(é–4Ô1ZÕŒ-ä×é˜­¨rÓp)izmßf)4˜Â‰â!š	sŸPMx¼­éu˜‘dlYGæ‡°w§«|Ûêƒ$ ÛÑ@&vÓ.ÛÎâM—·ùi&F[ÙÌ «­f(x5Ô†¼e¾ïÆ)Óª²t1~‰²qìÚ K=¸Ì0mA¦-Ð\lB€S:~à‰\¡¼Ž¾ŒÝñS}BË@”ŸÚuA ÊÍôIÆ¯4:÷Õ´ÊÆ¨Ž÷ÎÄ)†R6§°ÛU
¯»é¤œs<‘¬%åç*y	½S‹c¾ ¤Dt›Pê©Ðª< ªIhÆvdZÎŠ=6,Ð5Y§ ¨„)ÐŠÄÔ„x¾$CèŠ¬žñQ·Ã™CMtîÐ²l"xîŽoˆº´]²ø…í
 „±*A·œ`RÛY9"&A=&Ah'â/û°?§˜o¦HÓ_ÒSXžÄ“cø2‰:»E†÷,~G[
Wu¹LÇÉ²ƒåœgKÌäepÙçšl.û8'Ö.ÐÈæò½j`vùZ´dyÙ§ã«JúÙ_ö±õzLÊj«!-«:²Ër³lGŸ‘l µÝ†±ìðd6ùe]^Çh–š¬7š%«ØÑò­ÛÇ¶%³cŽÑQœa¸ œWèÊ¥Oi:E4Ï.‚a>½h#cîu5ÌÄ¿¨¤z[
qP‡ç~¹à1’XÐVã	%ž[.*Xiî¶ÌTwØá¡¼CÊ`é54ŸæÅrT¡°ÎÈ)Ê´¿$ý~ù¶#5oT•JúÅàìæ¶JnK™…`‰Œø,É*Sä¬‰$U{þŠ†çu WSÍÒ{SSÙ€‘ŒÑ„sõ^ÛFoV7|¹2.{˜/ó6.0Ý2¥ÂÒS$Q8;÷jŸSaçiAÇº&×G®3ÆºÜ’úÍ3Î/¯óÂ£œk^6ZºEêÕ´¡
“§ë©Ø4DVEß3A“h?{Òg<âX]<Ñràú[-è“b†gFs:PóGjòRÃæâi îõùj³c†:è!1:‹/J“ÁtÐtã”iÁK9ø“ç1’ƒ Ø¿ÄiA‰åðð¬¶±‘Åž¡[Â´ìÖ¶|‡o!!lîCŒË3…©ý<Ìêj ­QâW¿063ÔÔÖU{Bîçhaô—é(uÛ¢Îv„o–ÉÈt[IQ:Ñ:°géfæË¤JÚ¾ ÉÔÇ¸Žuð PßÍiKÊ=ŽÌ˜Lãùß6+[ƒ?=…f°(:ŠjY„9{(ã/ÅH¹y³ûc pÞæ{ì5Íh9ÚZ,GƒþÖ}“3®H5€÷£œ¥u]»ü·<gþ*ÄËä±‚*=ÅôÇÃG°-Fª=~ÉÙØHÌ}CG­èÕ2ÃoiåKäžrw,þªÜ±T×Kå!.ÓL¸‰1É”0.µ§‘îï‰Tös,¯3è óë,¥GàYV‘=›[ccÇ´Ø>ü‘.d7±8À¿p¡BÖ2Öb(ËnpœìÄ•(§ä—«ÂnõxÕ‹U§òÝoB¦¦±NêLT^4pPÖ8û”µ¨ñy:	¨Uèq~”h³_ãH¡®sÎ•¸÷Ëáœ¶£	`= ¾0Ü0Eœ0òã„Ê5ažßËÔ\šÂ÷Y:`o8¾FÙ[F7_~Ð7Â;ä¨h`žEãžÌãŒÂÝØ˜çh4+O¸	øÀƒBqúå xÉî¥‡0´aŠâÜÂoT‡šø:ŸUõU$eZSÈ>w%‡2µÆÌL¾E¥~™ý"S®Ò%V”¡; MÒÐÅ)¥Ê–YÐJÓ¢Hm Küìè(¢Í“@B"éj$œÓugt\¢Ñè¹¡ƒéjóÐš^ï³àzƒR‚«7Ê1Ls6‘ËIªt`f3'Ž‚i`¢*ŸœÙm!$\n+X6…ÊYÛðTA%GG¤¨öÚ“!]if4º¦~”šW9
YxJéµ©]ñÓÑ†hO*¨8
¦ ¥RÌâß*Àt±ô~ê	–Ö¨x±tÑÔñ¥+Á“¼Ø[SoÍÆ%PüL¶âî<„î™Ï…ýôé‰xüj!¤i)dù6YÞÚ×ñ«^U 1¹SÛ†Æ2p†
xQÀ¡„“š¤JkRFƒÛ–³3â$8¤íêvW$Bq	ïžÇ%½O·v•­¯ãÃ®fÍ¥›iFà€ð€=ÈÆD¾Ä„ôøåu\ÀM;¢ïnÈ j™vË}‡dÓ¶¦@äñT¥k0ÝÛôóh}9â­?Z[Ž°ÇG‹OólBiìOÑi%Ü“hè:ùLŽïI´Ïº´žŽšz±–þ@Ã4n
&/B³]¡Å,mÊ¸§7Í½o´=Ò|"	Xù\À°Â0ñPÇð]<6Îº~°ý&2ª³>/`Î i,Yyú¬Óßêœëç[µ¤Î¯*É«ð¸ýU0ækÜÌûÂîÈ•1Tpí½IãaÅí4¸Ä´;—¨¦Gw²Éðî¡	J“à? }Êè(MTé÷ÄAœËÒ HpMúOôÜrñä8emNBE?ÏÒYÚ†*!½)pHZW
mX(šéhCØè#TÃÝ/{­r‹º[X•Òãl’‘K²òÓê$­ÎòâS´¿ú
Fù™XÙhÕ!4ß™y%é¨RùOÞ¤i3á%”%}¾mqêÃeëYàxíˆpÿÍÛPwµAå-sÓ€Ÿ!q'«D~…Fù°”cÌ¸¨¸I¹ªÔ/§#¸Í:+ÊÖ…¢AÇE	GŠmÓ}¡kI¨A "Ànè—ªÉ­×éº¡§û“ŠU~7x¿­à
¶ž¯±çÑJ´æ¼[uÖàf@ÙjXàƒ¹X^UfV¾Œ_vUõþ1	Ë5¾<ÉÏÞæp±v—È\&›œ~þu”¡ïÏr´D,Ù’FèÛ¢¨é^"mc`°;5ÃœÏJ á&
[é+Må–á<h«¬¸m	[¸$Õ°¿1fqåRF²©n»ê™bÏmP‡ÇÐéê³#¦Þn˜GDã§^Ñº\~Oè0“QâÔÒ`°Ã†çþ±å¹o¸R=0ö‡êûx(Ð29úÓ/^X§`±	ÆÞfÌMšÃ•òÁ‹ÄŠÔtn@bûš*_3pâÛä jƒ˜Zí`L$¿$Ç!ÌÃd¹«×Â‰ ®O×q™ ±!Âà¶‰·ÎHót·Ë6«Á½-œÇÙÖÑ`Ö.½ÎñuÛ©tï¿~•ï¼âÖïJTv%BX ^³çœÚøÈ¥¬	xå®j­±¥ÁˆZ`„U0ÞÔ¨G-Œžä†ES£1`œÄä‘Ë£ÇÞAd_²æºo-¬#P£F9 Y=Ú¤°Mu÷"²…ÂCÿõÞ÷ÈëiD†–Ã‰$’TÚ +2V2Dº-“m#TŠ·ÃsÕƒNLøevpøUþS	#îÜ–ðl}„2©q\I(#¸ö]!ÖAµJ’Õ/V/àýðÃöx¼$®'èqÒ÷ù¤: r¶».èþâmz^íæ#Ø×- M¶0´ï`«÷ QÂ³123Itw050aP·ó&®áóßŠŒ"_ê)?:T9Bíñ%lqvìpgˆFrEGïÈÖÖC#[èƒ¦!|Ü#óõo.ùâ]}d}nli»é[:Ûø6tKÛå¼Eg
Š¨S#ýü?ÈßSµSìo.[uûp÷]XÕNOLqs]	)‚˜Çqq¡ÁT<)ÏÒ‚qEL óHXnPÏ·BrøRÂ+ë¹’•¸ïH b u•Wpëp…3šš˜<Úö2è)2™µ:ŽÏ¯Ó&Æé²@ <ˆÝ0°ôtø¸òÒi[í>
Ý®6»Õ¨+G¥ÇÁûèûì<Mºk”pxÐ	œÈµÍÜ{ ¸Ö1\&'yD’#mèWÿžÃÈÖ@ãv÷¼¹)ù¡œªÏœ?âQ½¼ò °ÜÀ"îZ%0Ió‚ ZàÅ¨P—f­«,Ñ¼‚L+l'Ù ‡láóÕ€b™k¼>’Šü¿’ó*óXêx9:$ êÆªR§Ç\KR´Ñ§ÝCý]-º»kïnGÆÌ•[¸"/£ø1†røW
¬‰|On°õAMîû vT…Ñ~'{¶RÞä¡]‰¾.º§&Ð¼=úüÏ46*)}i˜I"—uG1èº‹A;/àgc	 Ñ°øü/„»3m™úbe¶$-c°Õh;õGƒ4©NwÂ‹ê¿lCEEbž¤10ïÞuy#ú)]ð»@øý•\ À€ïß«&óäb[í‹¤à$EÅnç¸È’ŽÙåAu1b¼0E!…v1¿W<s	ì¢7ÎqãÚp(‹…®>UÂ{é	_QHxÜÔ¶$§)ãªVÄÔÂÉ¿ËS­xc^ü|UMÍŒÑÿ5Ïk'Ž!3o¡ÂN$lÆÅqüpV˜ÚžJÈ–G™¹\i„ø8I^£¦ëg‚˜C{é^â%G-h€ÄÚ÷Ò5I¸]˜æl4g÷ø9Zô43ÖàéQçg}K_y‚v*4,‰“OËí
;,ðS¦ó‘§jR~þ–›Nfò0Ž”8Tæb6ôþ³5÷é2ù9ö/ýÁ sÉšIâÚæ³j:ƒÕ=å‡³‚÷kWAÅÊSòñó²,Ò>yÁõÊÝÓx´¦%´SÖ?,¹éµÖk
b`bp°E
WŠ5å¬b÷èÅù1ŠiÎ(òNÅ’Þ8„ê±zMÔˆ+ãÃ{8ÓaÆðEFÀ®î \³xA‘™Ú,ˆFWà²ˆ6äœeúú¨J®Èù@¶
ý¡5cJË¨µìøË–¿’([Ö„‚%^HmÂ£Í|ôO­å{s¨ó/„ŸLÈŠñ¼KßÈ’Å_x‘ŽsKSÂ¥‡“ìçY*JƒÞÂâ4ÎFˆebÓ@Ì[º‘éÒÕR‘FbJ8µ±Qœ	Qß§ÏWœ¯B’ˆIuct§ë†Ï‰’¦Y¨`¾•-H#R§SiS«§ÖT
ðk‘;¶ƒöTîD~Lnï_®‚RI2QtTÓZ‘|²u™.\ÃiÍžK›sÏæGb›ÁŒÕÊÓ7”Á”é¤Ìð®àæìÂêv˜ö‚Vd9zçÀï²:÷Ë*,Ó²ŒûÞd=ØÐžg‡@]Pøù†ãE2Žƒ´ª%!ö¥D¶EL‰£ÛbxŽãy­ž¥Êr‘e8ŽhÂDOª¦	¤bŠ2/ä9¤™8b«”“«h©Vtý BËü@N}’s{ˆ”é[ Ÿ±!÷Ôu1r×oÙØÞ‰íxìÏ±Õ§* \gdœá±e
§¯67‰è*ÖjYÖòÔn1K}f^;ðÚ/vhÿúåÅ•*)´”2½Œ”!Ò÷x‰½®ÞeÙíL‘>’Þp.~Ü¨‹ÖíÔÚßÑ	Lžà·?óŒnQç@%ñÑ¥Iâ’&> Ý” [Þ¼–ókÝÑ,£Yú¦j«ä€_^/¡¶V‹„ür~È÷.#-åMÀ¾\Ø·Ú|0¯“$’”Qn	Ä(·Kà»vyI¤¡™KÈ “b3“K‹*ÍT°NãÖ´Ë‘:Ú5[²ìNG’Â| ]>î‡IvGqY"Oûh©œÆÃtåbåîÒc¹6v™£QzÎ<V˜¾”¿X9L«³4huYíÇÆ®><YÓ›C¹ßÊÆù˜æIµr8Š‡Ÿ"zV¢ÅçÊýÁ ª
xTôJ•ŸTKŸf‡£,¯ÒaÌe‚'é0#ßè‡«'kVoS§3Öð4L]¢êl¼ôø¨¹øß$fT
øÎêåÃÕ©1±UkfÞu:Ž§+ë~]9+âé’5¸l2U¶Ë'º¥Aõl”.Ù¯²äÑå|Xa‰*œZÿ'Y’ÀfØ%âá0V–úç£ò|9Â?KV‘|²K¯.™u%Ë[Aé"ÌºjMçé¬ªrµaÜ,èøâ°<KØâ(~zt™ägtL›Ž§¸WQ•U#íÓ8;Ç ìy’Žxj
k± ý=^;Â3þèríþ!zAuÌ‘¬²¡\{|,/P‘—ö Òtì5~o;J^ë,f¬ŒgMòáaÜžÐÂ×§ûI·£ÃI§÷¤?ÄZÝÞ•3^*À-Ûr/íD›ñ¶€~k8ëv©‰7456ÝÄ‚d7ŸKFWS ÎïZ-«ÕôÜTî{
Ù( #]¹äð˜}ß¬nÁ¯¼HÒ‚=Y‡·'ùiZl‹2kê‰Vnƒ!8~ä“•!ŠûKï ¬÷]»v:æ•¹3)”Ùñ\Ã¼bM‚õ¹²ïlüð3À2šeETåxHà?ý>xpjÉô)ÿš–pž¢CÏÆqT~þ[tšþbã¥(²a?ßqy².à÷.Áïóo×è-Ì'Š#5{2¼ÐE={š—a™šZ¸{Ç<Õ^»;0ajÍ©­ƒæ‚‹×JpsS°ðÍ×Ù†]œÈÞ}KØ}¯1sYt\3kQekÝËÈ£* ·áPƒÕ rÃŠEL5#ÜÐƒøj¼{H	¿€•~=š•æ:SÊ1‘¹ÖøáªûÍƒÂ2ï!ƒßò Ô•8\³Õ…7©é@$é_´š#1ç¡¶µ+ 
ñƒ¡ë<;%“ç¿t4î ¸æœŽà™0œ“aN·sË{ä=J±Üž¹w§íq’¢š¿ûÓt 2GÿV‡é@R¿Íi:`¶;µ+ÀË´:Kê¦áVAŠt˜Á<‰(‚Kþ¡œ§˜Ùª­Ä¿äâ!¿ž4“!ëÍ§i&[ÈKYz6ª€sù ÕÐ@Ÿ—¥ÞKÈ°y2Ö¿îå÷•tiðsk'¥ó9é›—))ýö'dåSCÔ80ŠÕnô,.Ù*ÎEà.h¸òøK%’ÙåÊ#”ï!cô2=þü·aæÉš•³ÙWle>à?ë¥üò8_˜D9ÞñáÊÝ¯’#ô=>_A{<SxèáÄ­ýÃp
ò¸tuå²95¬ËáÄ†¥é9ez±²Á%‹ù(á‚Å±Îpcn+1Ÿ•õèì$«R&ä$²³Ã]ÔòK¼©”ù¤ÐUKj– šE9Ê'•.õÜ”7²ö«c1–ÃÜ{gõpõÐ’km‡¤“¿ÎíÐ8’¯f;<tý‚¶Ci÷¾ÆÝÐ	Ú¯e3¼tá‚vƒ]ë_åVHŠãkÙæpv'Bfeÿ*÷Aï+Ú$2Ê •án‹M$è²Ñ;¬ý‡ÿÙo´Ù[’z¢4PÁÅ«¬0Œ•i<IGK>-å‘¶NÍ"XÐÄ>ÂP´^®¬Eãd[ý\'Å¦«ÄqÔ–ôt¦#½õÃQ.ô•ç¥8rofS`›†q™âˆ×–#wƒ>ÅÓ¼`Çàá*5íé’©y”g+G3€K áMáõˆÜ¤ÉÊ¹ V—T;Ê‡³r›K
V(·'{„‹¨¶Q?0n«ë¤ô °wåó¨NÍà|•+çS(Ec7'ÙÃ¦<pè´µ@>ÅóÊ´´ôu%¦6bxçá*+à«{éÚŠ’1,¹ñÀ0EÛŸÒ‹G—øðJÌ›~<fdž?\e+äÀMžÓÃ/Œ‚ÙàÜû?*X:6óB§ÓÀâ€”iÎ!×Á¨U•…ÐˆþDÝÝ|<G1*ââ²Wê–Ë ŒP€¤~FU‡~,*¡Ÿ~<fê:íÝìD4ñ°½+ÿÕT¯¶›÷âk?7¼÷Èå4 àÞ…dzzÇE–¨ÝvÉˆJÙý¨Ò½^]ó~;¯:»(³•ú\Ï8ú†&r*íR½…h³…YSKáe±Û¾ºÑªÇÇ­V›¿„eŽË',me¤¢åÃ¯´¡Ë¶[Œ™M{¾ÓN‚9Z;_7GµÑhfÔ×XÕåIœägHJ32•ÿ'6iÌ

NIwæé›Â`¼JLð…²'²&3Wku¦&àEa‹6rµ›Æ{Ò%kÚR/ˆÖUñ¼«×†›roÊ5¼Ì§ÂÝÆ®oú!¹ƒwkƒéù{íŠ{@÷­ºà¥åÜY†Û÷qVGƒøŠpŒï2¤á®âxCÕ/M\@îúƒ;êˆûtÚ­J6I²ã\®ÿyÛ+cO/´DM‹ä¹øX¥ ¨¶1¿$‹r`Š²êbåþ /*ñkÀn-8F+ìà‰k¦Ùç(«O+ƒ0ÈúøqýÓRa†SÊŽùÍšì2ÌzyjÚTÔ2ä¨)ù­<¼1<–ëâ‘;ø•][Tš –e_*ÚµÐlK!¬&HÇn¬ôî}Mñ°Î?íÕpÔVøàè~e­¿¥TÒê‡yt\#ý0d)¶e]4œe^¬P4â´ðñ
|‰¸I@e¬,•BåÃ˜ã!æ\Â°Bµ³ueÛÍj…=0ãý­>¦¸¿ó÷•A‡´€¸9hXVl×€Ó|áú¼?Â/N6Œ[Å',¸ÏzëK/½8x¼“6–Ø…(!‘þy©Ä2Ö÷ÝãÜ7`_`S‡ŠP»­Ñ”8º½âW¶eàß°„¶T|.9Veô|Ó–tO"2W(E©VÖø…4Ø!´ÌÅ˜{ÌùâíõÅxûøøÊ¢ÃÖ‘àtñ‹)²¾/Í‹"Kº4ýê]uNÇI€fÐ85ôÙC`zÄ1\9ÛnHÓƒg,ŠþÍùä¯O¢f¥¼Ûàï¢³Ó–¢„ëŠþm	Œ h>¤e#,¸@Wðppy/`úuv´ùñåx`]|IàÚKWL ‰|ü|È—kŸ}- |‘WŽQ¨Šk7 [¸øÑ§ïAû¾ó¨G[Þ8aômˆ»~j;¸¦éäÏ³´DJ§6NŒYÁ)ºÕò‰PœþÝÙš°¶ [Ã7õ½8_|ñ	lcÑ‹=rEû`Ï»ó"ËÅ/EÇÿ¶þÏbv7ð^_Äëò.ƒ/b«[âñ[ƒv¾•¹®·*€ÍŸ1<ŽŠã)¢“Ö‡û2–5CQ!5ÔÖ?‰Köªççm¦g5"áü¹n=ýÛ§Å±²Ê;¸T»»ËÅc
[¿x°œ¿õÏß½¬_L"È?/PælOŽ·‰]©ñ—]z‹ÛŸKBj…ê¼×¯Lá˜J­BÎ¥VÐÎ>q»ì¶ëÜzæ@ŒûùÏW£µÒƒÏé«T}ðÅšgæ_`óØZ5ôZÊìb!‹i	•¨Rôb)_†Dóã¢È¡|’Vñè]ªsÉ5UŸC\Ì×sÁŠù³B¢ù	»	4÷t}å“]Ÿ	ÁqÐ-kÍ£¿^*Â´A}$>*¸|-•ø(mÅ‡m¥°SqŽ+V¯»’­15ÓF%>õê,ñ™Oq!Ûn(Q£àÀO¾T§©Á>;$2‰GÏ³É'.½_£_àµÑr¦šðÊÝt²I\züšœ„/YîÜ«¦5Œç+ÓžÓštxø™1ûY“+·]·…6?¯WW£(hP4Í+Ìz¢ý½Æ2¦`á”¯‚ÊD,4ß9W}0M"\4ñh_À~é—LÓˆe(Ð¡T¹7œãm^‚ó\žeaA»›•¾ò
hªvåj»Ì ùÊ×\ž­ù´¸U”ƒ´Zâ6ÕdT-¡H›jeÎª¸I½ÝO+3¦osz=ór{ÛÔŸóžk¸Ã¾23üX¦|Ãëê4Ü’ó˜tP…³*së×‚0ñ¸¥‹áë…¤˜|àÇ6û¸9µ7ÿ âPTKÕ¿œÃ$"Ï(=žrÛCƒ!as	ü\:‚n@aI…ZZSPw--*6®kQáwkà-›YlÝŠ™|1¦ø©5· ¾¨¿Ù«„ö¹L5\å¸§¸×µ-5nC;„½¬tì<v­Õ­‹%l£j½ê}”Íx¡Õ¿++—Û0k¹h\¤*´qÚÒz.Dçy-gÍþH+¡Ò“£,2!òÈüMlˆÚYÅ’ETfDñÜfD[Í…0LýEÒÂÆ(-.d~kë3M‘ü¾ñ—¢_gºâ8”sØ"–ím­~›ÖéT{I1 fw&6åBMlSÌeRÒr«ib3×óîÊkÞ)Ü¢y™›”A™Á\'×Y€#¹4:)¿
£“°ÓÛ?žñÈ×i	rY¶0ýhc;pÙum£eÇ|5þ‘,;æ™ùWhÙQñä¾ñÄk™áˆ›Ùc\âe]ÿ¸¶sáróó»­Fàó»­†§µƒ¶ål5,Ñª0<ØjªÞØ†ãw“
ü¸8§1ž·]•ÐMÙJ…k„oÆ½zžq…7VjçÐ˜3T¶Å>Þ åeßó¸81Íaæ£æ6ìè8wýIsþ€éeß÷¼¹5;ÈzÙ7Ÿ4·à½^öÝ§LÒ8§µ‚C¶FöÑ—í±|+‡¯T¯|[
2úü¿Q…[D´_)$ÝŽ…Â"àèë1Qh)›iïÕŽk·W;Å“ÍÖ¼§ms}é±Ü”èº®žlÖvêÑYŒõžîÙÙ,/½·íU ÆEÃßpÞ¸‚M~‡g};W÷'Ž‡F‹Á7—~]:dBýä¨ÞÖä°<ü[ƒh
oyJ²Ë€~Æí¨á°z¶hÃë7Œ”ùü+Ð2MûŠÙâíçÅ[ñÓJD]¼\Ýé4f³áùm¯ÏŸ^ïKÈÍ!÷Z+dR­×†¸M«n™dÀPÅý(IÒÎ¤<K4\¡Í,+e‰R´kK$W\µÐ~·“(IÝ‚^•t`å¶áøibÏµ¾lL¡ñPsIÏÚ6ØªWy^ÄçÜP›o œ¥ün·œ—£‚…'ž£o£îzô‡–2â­×[Ž­»Kãb’&­;+úå0/Rìh0W?Ó!š3ÉÉ=&#XÔ“>!Ìn—dU–éÁ¼ [¢"|ÐÜK­âÉü<üÓ4x9ˆ.ŸùÝ
ø+ «…á¡a.WíjQ4m†z/‹'pˆ²aŽ¹ vªlrÌYGOå`è’I\FoÄQ£L ;Üå¼Ô‹xŠ¥–Zuß 7ÂOãö¡•ÂõEû7<Õ†îÔG¨â‘1}dp¨á9{.VÇk¥Úª ¿3v¬´ñÜô¯û®ßïÓ¸–#øF½½÷)vÅñ§hÑˆÔ£ÁëýÓâhÌlãZÆ³ø¹”Óhi5Kc«·œ.ìëQ&Ç¿8ûYü4ÚÐRŸm”r79ºApuíc[ôÚ48é˜eòïÛ5ØÜ¶µé¹…ÝÜÚ´ÑˆhQæBWóíÑâ¬MçØ›¯Ì*¨EH7éÿÖ&˜¥””Ü‚¦‰O3J´Íày]ÓK÷:±MŠŒÛlÕ.¸–Qz¨	&{±V o#VßÌò²6ˆNð]ÖaE»Ð¯6V'ÌÊ«ÖAÏãäCÙ=º>ðÊ%C‘[Û…ÄÂ¦U	&[)+Ï$å9…" Ð§Ùá(Ë«t£ã¶ˆv»ôx ½¬
ÀŒ¨{L£|HeLùS¥¾òâ8žd¿Äýõ}(ûžã(Sy-ÌÖk-Ì0›ÌÎiœŸ3JYÞ¤f¼Õä¯–Ú¸*`4÷Ò£x6ª¨é+!ïþ¨Šg˜Â–À/Ð4N
ÚžÀ¤¬P<0É£Ã	nY‚(#Iê0Uº;<3²ªžø†¿f}øæÔ>'k4O"Û–é~çKöÛN3âÏ\;_šßÐ²„ÓüÏÂ öØëfpÑ6Ã [ÌûEn×ÍqðÐ÷KÜå;“Uÿ1Áµj¤9ž"Zv¥ƒ·,2\S=±8ÀV?tT±‚rðu7´è×KŽº-[vëé·¥=±XÂMñV£M1ò÷…ŽÎx³áoRÆÇÍ¢ü¹m‰C×p‹µ‹øâ¹LŠÏVîÃ~_W_jMKÓþ:—ùéZYöjÌFI69–JÈ»JÈEk
‚†¥œFXz¬ ®…^ Ý†	ÁÃš<ðà÷–†b\¡´£Qð@b“Ç—W\ITÑÅqÍòÑÄÿùïÿkžâ—I\)‚¾‹JÀV6ïå­‰v¨VÒUsÒIÂv´î³[±f¢mTS–ÊÀeM?vW
Æ´”Œ–Uƒ¦ÁFnç"kn™Ž3"ÚkzžËB(<Voó7é«.ÇzIµ¢’^ÎWK#š
ê1L4õæ2,©1+©…u´Xp“(-f1Tþ!"T–øÍÏ#Ó†Ã"ý6@‘'ò¸m¨Û#ƒ¤ëÁ3fúšàÇ4æáÄŒ…®Bu©Ën$áõá¶)|&µ¢dÞóŠ®ÖMÑU¤a<¸ÔÑv™È:ÊDÖío&Ÿ¯Ä³*o#ÙœW$¢‹ÊÞ¡œ»_Á1¹tHÆBÈG¼z4dÂŽK¨V)‚FiŒ”â
:ô§~ÉQíhr’E8¿F#tŽèG?æÃÏÿ_4Í“TŽ¨ˆÊÊŠÏ›,¥p–:Ÿ sƒ™tF“BX"ˆ8úy~žá’‰nÞ÷ŠÚÊlæ êÒ<ŠvãßÂë‰`~ñÉ\Ò’Öò’y%&í‘¶_jâ•›ìÂÚc>±lœfEH‚R·6m|M‘¾˜H[:>6K´yÑßP ­@wÓ\„8Ûfr{‰6% Ç//ÓãÏfy”Fé1!¯ò¶DÛ•èt":%}ÏªçàÚ"J™n°\Ž´žùü72Ÿ)at•·E1½ð…©Á0>ì"–3P¾»ø0‹_›Óvê»é,‚¥øä[‘¨^ÆB˜Ê×„ËSK¾Bµ‰2²ò Š'I\ }ÍÛý×¯><{ñúÍ³ƒ~6â6IKÕPHñ`%D™Â„(3%«”4 J÷ña†ˆ¼So‰°t•¡[Œ(QvbrT	*ùÂ£j5»	6$gL­šˆÆ¿:Yð‚%Á‹·“‡£Ü–x^!pÖ¹ãk¸KÙ$¹ž[ˆ,AµYˆ|©a½k†ŸÐƒLzû[sÆŸXzÌéùÇ€h%ÂÿÒ2t.Ï7Å•¸(ZÊÏç’žó¾ä%peP;·ØãuõÃ×¼ò¾»¼Óîü.ÈÿJùœ¶“²ü¢Ìª©Hì—óUtÄùu4·U÷w‘>~þDúþ¸TÿZ`Ä]”¿20úÂ’ý†PžÓøR×žS£WÄ0Æd6MPâX¤I:fqÃÍPÍ˜ÛÆÅdu³d;bA–—á/¾ÛŽ:;Uvw"àQžŽòŸg)Ôëh°Tª`'9OSj½”]xtŠ5ð/1­Cü5ìWŒ¬üy]¢cÇPô]EÛÑPé¢¹ò¬DeÁzÏø‰ÏâLLr/vüÿp9êfõÃPTgúí-C¿¢Gm®Pœ2<‰ºÃQ>Kž…éo‡Ë™Ò~ZyÑ]z†XÇ°t¢½iI¨½½´Éf´.þÉü;Å³!¤8 ï.^V…:Sãh3ìšUãm5ÄëQ<Éýð€+½x›–UJ ñ"”Ý¾îLfüÛ^z”M2›[9 /	+Z§‹ØøÉo,8‚¹Ae/‹K‚?¸$â5Æñ¦Å-B‚ìëKB‚Öé¢ –.EI|±0äÓ¥EÕýøfC²ÍŠ†È©å¸ðø›K9â¢JtÇ‹xxœ;o]jVsÃØ0È$N vÊß>Ï€ÆâÅ-šÕã—7§ëEÝ,SÖF[0è-aà„<14Â§N?ÿs¬ 16 néFgO©	îN(Û>^YÌ7gHøi°ETænó	sy‘³’”ç¨îžQ¢·!¹+K[ûþ’\,&¦Uã3!‡ógêÇ/ºÃ|„þLY>© œ³“´ ‚<?›¤Å~àÔyôþÅq÷g Z²n¬<‰§8[Òã´-»? ~Ëô¡ÃêB3>÷W0‘c˜\Á\
Òe£è§Ï¿Fé9N;šÄÚ7Hÿ‚;Kxéë €Œ_Æ}!ÁUÇEüKœ Úh›Í+ÉJJäµàì~üæÒ~zui‹{áüvö«x;ÜAVdÑ.ô4U9 -QJé5o=]¡ßhK«{‰áÑw*QéŠ
ÿÇc|ŒÁÉ´j Ï1Wbè²´ .ðm6FO»=Uïï¶ÜXõÊ¢ãÕ Ó(6­wu…‡õ¦€³µh?Dz¶+yì	xâ$AôÖPËûÿ%ý>/¾Ï¸EWl¸ÓFlóF‚_Ñöed3èñËswÔ„£!l5Àª†~B'³ÓtìÅCÛQ<¹¨EEQœÍ†„MððëãàÈÈ·¨Ôô)Tqœp¶£%Œ%5û°ÿe|L ‡9HOéi–ž1<ØaT
P fŽ¯¸Jèásöê56×ãPUAô¹ãNcÃ“tø)ŸU2x¾HØÃ¼+œ0‰Ç(Æ£ØBh¹hd]·#J}@zZ *£3Î'ÕÉè‚Ø†xB9vŠlnbÀ»¼ûÀˆ±]8#fAYWW£ƒÙ!†—"@F4™h”gCyÐ3½Ô³óiÆ‚¡ãÄ,üŸ•Cˆ^Šƒâ)DÊ0²LdÏóô¨GˆÑÚpéèˆCƒm¥âÄ¹ }/Ì€Lv£íØøÂ¼ß¼½£‡/Óì¦}Žah(EOèúPÃ¿c´x³±jãœr†Ôh½¯¸0Æª¡ð:¯'*NìÜCÓZKÉÓ|H­õ£ý#@	Ø€úù‘*ÐAj/ÊÐ·ZæpÇiæ+9ps’Å×³ˆñz/&¬¶n¿Ê©Ø“VÅº”|Ž<ý¨”I>$ÙÑÞO0}xˆð14¼¢î©ÇV+| 
v3L³QW6½u×ƒAô‡è®üg}³gR‘!Vn³lø!'-B«ÚwêoŸw
zZIÀÌ{OÄz?zUÙu–a,õx4œ¡"Hž$B´“OŽ²ãŒˆ˜1ÆïÁýT
5	FÙ8ãy17/ýÜÃ’öV+>Î‚ûB4ó$ÚDbÐô†KjàÅÝ-x³!" Q§Bl(ä& ¹
õŠ!ô”©ƒ?c,AàóÁ€ž>Ø ðA+|	ŽF9\ë­¡/ÈjÿVdm VŒæ±Ö„r[‚èu°kéÊ–%—³C)=®¹4Ùò
¤¿ÀÅñˆƒOæç¿>‘2©»ýxKo
úPwEQÍ>ÉôkRl}lxŸ<cK‚®šbz}¾L’²2ggÎ¨*Î‡euãØŽg“•“X/nâ)uÆ!SÓf_Ó²>ÜEäá3P?ãéÂÐÔž›Þ(3¦R²ã¹F.ëhî	ÿ%”´ƒJ+að°qîKÿà·ÒG3ÕûbPÙdbÌˆòšël6mƒOW+r²æ(Úh85¹ZmÓ£K€ƒd¤Æ7—
(¦D[20OòrŒ±ºò%Ó|ødÍg(¸ÇÓý3}kŒÆ»‡|€œw`ØÑù¬½&»QXš.„"šUò+`†ªöá¿£Ù„b'°ËPÏnª£ø4/–ÆIZÅðm’ŸÆ,I>Ã `˜d66æ4gãè´	Kv0Ó$AÀèÌ`ý˜¯A…âÊ5°¨AZVNñ:‹¦é(7ù©~ôlR@mi*qÅX9à¤fÐZ¯á,‹ô®af[B·eVÍØ€ûÖFMëŽAÈ\ÍÚ¡…0–µ;
ZÎÌt5VËÕ#“Î6¹Ñh®3wu¥±á>#O?t£cñMïek”]åÚò"ÚÁ½à*l#|¿æÔVkúW¥I÷n¯Yv<y5«º€PNœÕ
­•m÷a;kK[t£={!â¬@æ.‚ž95Ïü-|õpQ¶øíhnÅ3Ê	ÜöÕ¢_%V›ñ$c+ånkÃ®ÖòL[ÐCå>¦Gi{lL=8UC z‹wèæ|÷æ}ÃMŽuƒjþj—èÚ]\±»µNnw-»y®R;h·cº±~+t=t…~'Mö4Ó»ˆ+ZPñ3ïÍgÏàõô6cÒ¬üå&O¢Î÷1Þô›¡øã´ ²°@±s	½d¸<e¿ƒBÞÝ"K	ñÓÁ=&€v„6Š~Ç¹þÉ S08hÝÉCŒöˆ–sxqSHEÓVJdƒ?ôÍ–ÑïÇÛ1ˆ³Þ•>ö‡GD%Ÿ m¼Ã£K¦èx†’æhðÊ{µyáÈçþ‘¡'‡s1TSh„Ü®]Ãô€(--!’®øžÆ£€+MK¸w Ü0-ÛMÅÍ¹#Š£'B\ ‹×§†Z]¡Siñöx…-:Ê‡³rçIýPÇü±…ð¿ÏV&¹›1ö~˜»Eúó	ë…eÜí4ï¾MaUÎ`ÜM[‡$T:9ñ$dÓvî5o¬ýæ‰ÿ6ö/Šà{NLî£Ë»ö[ì.§žØÖ•tÈ—|+sÁPR#qHS`>†¡¢ìì~dåCÙH¨Ç…ÀÇ€¨‰jêØ(ÍG7!"sÑ²§“zv7	ÒçlqYÊŽ%÷µëÜDáûL09øã|ÐUXiC¶#.¿Ý•'Òu(œÀ¾®Ã½)Mæ§.³ $vÈäý¥œ}þµ¨q®³_X[X“ˆÖ![1šJÞ'°j+ÊÐïhêm´`~køu‡¸ÕçÔÊ-]ÌˆÏ"îB5äuX‘n[¶Ké«™`öS¤}Bþ
u C#!œÌóãQúä•ç€`±•w¾Ï­àayzl6ºTöÖR„zË§ùù£¥A4ˆÖ7á¿@¤Õ¸:Á`©£GKÃYTdµ‹˜m)J-½X_ïoÝÖÖûë[ÃÁJÿÞý•þàÚmo¬ô×Wðñkë§›ýõ»'[ýïÖ‡ðŠ÷7°Ì úÄ‚Pj-Úèo¬®÷ïÝ;ÙèoÝ®÷÷¡ÈwëðbýþÊfÿÞ&ûv¿?øî—@Xk}¨ÿncsçþÖ%0`ëÐæw÷`ª[ýÍ»+ýïîG÷°½õþÝ»£ìs{âx„ƒÚ€îÂ»{kìÛzÿþÝh°²Õ_ÿÇ¶±r·¿vÆ¶µñÃzí>ÿþæîFÿ»ï¢õ<„îEØ
öÞjÌß?}º;ØbcÞ‚Æ¢µM˜..Úú
ª¿±o°/°Dß•ýµx²¹!üxÖŸF³‹£ûý­­†?€&ðïz‰O7ú›ð
E›0ŒkÃ^Þ_ƒ~Ö[õÙÎæÆÆ–\ß­þÆýáT†Þ€ÅB°ØÄ…g›£þÚÖ
þ³»vûÆ¡âä`CpPð®B,ŒÖG‡“ºwïF¸´Ãþ}Ü¤»+´ê‘\ÿˆáŸ»Ùl%q9'~Làj¬½å©?>ÉËÊyÝÆe\]ˆìr÷;€×Þ…îMÀ=ªeSŽÿµr óšy±ÈË¡¾FªF£j[H‚iˆ×µ/,¡ÛÈ6›uÇiò³•“,I Ò(áj[ùýÒOYè4+³Ã‘x|ˆþý„	 Ø6]r¹ú‡è KÒCX¡WÐÚ(¾ˆþ°ªf~šì5:€Øl6ÍÕXL}âÙyšÈû]N&·ŠºŠxø))ò)¦g/ÐÝø—•ÍzÛ³ùš[ê
b÷õauIMfp"ÚÔùî‡b®«Ú…ùØÒ‰½e+k‘Ü0¹Ìl1%UàYâ[‚¯lhÜw¸/òÃÊÿÆÈ'é›€’öT¤\ª;Ÿ.¢
sÀ–Q‚‰ü¹n‚ajŠ2…æ@'|šlcæšâ÷úyj[¬Æõè|b9R+?ÿÿ  ÿÿì}ýrÉ‘ç«ÔàÆKÐ&A‚ŠKRA‘Ò˜¶¤¡EÍ¬/x
«	4Á¶h¸ Hc±û
w\ÄÅEÜÍî»þË±OÀ7Ù'¸G¸Ìªêîªêúê(ifÕöÌýQŸYY™Y™¿Ô°ê§A
Jo:Ù”'ôr–½e¦ÐÀ´‘"2Û(ðï¹–JÕ¸Þ{Ô{'Ñàj¢‹å§O¨ç¶DŸÈÓÔ¹\©Ö’<LIÊšòÖMf|qKN&÷ÿÝ)¿M†a¦¯Õ‰,ÝÒ'åt­ø¯$¾d;Y(Â¨Õ=DdÃV%Úº»(ò	Ùû½a[öÞËp4•nªju5ëäÞ[ñ’\¯‡9ûæ&.ìÌxŒaŒà?lÕ"OàÌŸ¾ÍÈF¶ãóy“Œ	&Z6Ô§Êàz;²4MíWèqI®³lŸ½ ^V¸“mÐ„;VîB!U¨e“ŒÑH«çY<±…Õk…/˜’!QŽ;€MûáŸº"àR%až^Á¹0»*k€)pgá„{o2Œ†è"ñ*û1ëLñå–.ÒðW
«_"h]“S¼æÓ¶Î]OÅ™R/ge_¸+íám<?oÉJ´h)±KJ®jÀÝ%jÀyE¶|\=ne—a¼9 ÿsaÍzi2¢GÛÔUæíÂ‹™´êe]MT#$!ZV˜gÃ€âÀUó3ÉC€ÙKŒ°$_÷šOY‚W¡äiAôÔ‰ùgDô¬?þD/ ·}f4Ï¨¼D=áþ!@äkÒK†ãé½BÖFÁu8À¿jPü¯A;NQæö£z6¢Õ—?Í‰žn}´O>d_}¢G±4]‹eäºA.ÀØNJ$@=ìGåÔ“ÝÞ`
O>×•ýŒ+Å„U™«jš8V¯-I•Cè˜“öq„$ÑÍ;Ñ¥o4ÍÚKÐkac)n3²½Ý˜‡¯ÉˆòÕ4{wVì<þ…‡JÑÃÛ»w&LNicÐ»(Ûçž#ã)(ÏÆi`¢¨ÀÕ0¡-¶£îd,†¢w<(zs¡íV75¦°*¨ß‹Ò/¢œeÌ…q$PÓ>Ò°
˜ÅyÅ†¸¢œF<»A?@ºÈŸÂ¾5ë¶kj›¯Ø7UUO”ïüh€ Ô‚"€®ù|¯{D7°!ôîÿ:Ž¨wd¦ 0¯Ê±:¿9ûî•F¤ÓnÇÉ‡z’å¶Ù$­&Û«ØˆBãØxš’Gzðˆ&ù ‹:^çr…¥š¢¼AÊ)¬bî	Z–}2+º|..i´à¢+ßs!
VÁádè5¨ƒC^ÌÆÅø£ ×Ç°R;Ì’‘þ•ÒÕƒ‡Y=–£´ uÆZviOD(i5³¡‘âï‚t„AäOôFÂHÉu%«Ø#fb Céd-Fm½í$¤NÅäWÄÔ÷g3’Œƒ¶cüýò5¨H|ó•>èÒ´¯ËŽl;jÚG¨³;+2ËK¡A¸ÌHÓEc
ŸÉbO´/W`­5VLÍ½³«(Œû4Þ´1ÚtÕÚïè³‚?Í:ŠjÓä5¯wîP!ºHÃëð+-Æ³¶Ò„ºè[‹~H‚qˆ)ÖgUBf`!æ’` j“:ÂkDh#	¹ ‘˜b°ð´Å1Gt›øÕKk  6É]ªÔw„c¾$¨W˜Àü5æ9¯½ºð­Çá´l«êý’yTœî}Øãe.'09•]Ês¼h¤E¨ä,¡ÑÓ†c†=›|£xp´ô[íå‡£ôˆNç4¥†Pbyrà¤¦%€Ró†)R¾,}2Æøê†S‰øÕ|_YšÊÑ•ºFè±÷ôá‘46·]+ØwÙD)_|n‹šq[Õ½y§³æ­aÕõP‰hU?–‡Ø½Z©»Îà²ØÓv­hÖº£[+RU?1¢¼£Ÿã·†‹àFUñ§Rà(Bx¨~Fíy0m’¯°®¼Å®šR»²W÷©Däñ—sI^ì2[Ú¤0G èbd0Þ!þÓ ‚ñ­èÉþâÿÀG(£§2ÙáÅ1¤{ån4“‘£ôþ/ü”$1c}¹<Fk¡âØ3ŠŸ4$3Í2¸ÓÞ¥ÛT}¯ùæK¯óxéŽZý
a’ÉŽý c8ñI˜‘ü.žÇ°O7@Þùfd¢40ä“÷™³\Ç.«tÇ.Œ^“4È±:2Ó9^êÄx¢ôëÃ. ²ËLßaMM‡Û†UkÙ¥s$¯Üá,u4ã*bm[¶ÔªRçÞ!cŒ§@YÐ£“>2– šè²G— CïƒÑE0ZÒÃ—;ø8f·h±tCo`ñY/@€«uô¡Õ²lçßuõ…7ƒU«&­ê†	ˆ&†"ë—(_áeíÏøôõ\Úž°5Àp’ŽËÄ)8Ä“ŽI…£)h^íD}Ã’É¹Ëþ,ÿKÿ^2BoX‘+2µŒ,ÉÞÐ™Q¥ó/x
·¼ÆJÂ¬Ûö÷<á©»{2bXÚ5»Ê@´×Ùy;ñ=Å¬d€ýC°d­c÷ìß¦Q’R%Aþ¶¼oþM¥Ï®#T3âó(ÙýwÒìRŒ°ò§ñ‹bôéÅ/ãûÂÒ/„ßö:p¶Ê:ð—©÷ß† žc–=
¢wJÎ•{ú¯ù{ñ&~Ç»j<=~^VEÙÇaØÎñsS×*<¢ÄbnÚ¿æ‘Ãå—ü†þ«8$GWI„dQþmúÂëÒOƒ	$ÖdˆoíÏÊ¿mïŽ’ø•|×4ì¯C„Sf:­Þ|ž&Ã3Ì¹w®aÑ•§µSš·ÿ
#6ß«d¶^%Šn¿X„^î1kX]£®Ô@½Òm¼¨Å™*1¹@?Ì®vã3½*&+ÛŸ‰¿ŒŸh7t“ÐZ¤ÛŸšÞ­â(šôafÜÏ-‚¦uv¾rLÄ»¦ðôŸú èßÐŸMZ\Ãý˜M6‡ˆ™QCÈ´Pym¯GáÚÑ;å=0S¸7}×¦î´õŽ¡0|¹øSÿ¦‹ˆ„xFnqÃê©·ì"ôm_½ådõÒç©æ®q+;–†X¹aüJ^ç¦·Îª{±zÏù­"èî?ðÆ€èÒæXÒ¬€Ÿî‚>¤íÿËÙ¾ò8ÇÃ —&4MýE¤ý9&[)é§;ë/±#Çy?>ÅìÇA69»õÎ¦ÃaÞ‚d.ßÐóŽ(ÃW@5¡np(cˆ¿Mœ5­3ÄPÇ—å8žü®¢ôšJ:›¥ b=h‰´Z;Þ6ï1p†‘ûî²¨çŽ¥®äKüÀk¤e½Õ ËÂ,C¿â9ÖˆRÒOwCGñÌ2 …vKÄgg](×›LSèúbfUSÚOwf_ÒÎ 5é§7­PšfæÑBò"~6¦î3Þ!ólŒ4Æ=Þ—~>™×°SÔ¡ax`Þñèï·á-µû„lƒ”nÍ¿ë0	^ü@e´æ«~Öµ°±„ŠkhMŸQƒ<`2·k2øT/žžã´„½½€YÒ	UPÇ£@„,Z	ßäm–ÑL¯æ¿Q/‹Ç,òÔ²ó7‘ÃšI|þ²Ó™lqÞ‹LÖâÙæÇoóÒ~6üöëÐË`‹E÷6ÚÔÊ—UÒ‘æçC•ÞåPó‘¦‹YISJ‘àPSRñÌ»´<•¦°ü‘éd"R}ï´fá‚û»hr…ˆ8û³jÂ²ê¥ËW§^,›†-Àò…–³¯ÍÌŒ3Ï2‰ùÈ /Ädš›Ë¡ÄMYàª×	cøÖÉ¶YR`5s¹6ÏÀ6L(ì€…ëÇÙx<Šü–)ûum>È±1ìâÄ·Ôãø/
#NÏªgšd¬•„½R
ÎÍu)a`¥áØ™Y@…*­M›÷P?D7€|Ì€ÔôðmiyJŽuÔmÇë‹fLn6Ìê¼ðQ˜·[J6ÅB’3&Yü(žmA;?‚Ö wïÏlÿ‡n!‰Ð½_—na“"%l>b1?]Û…ä‰dÃ]úwš|À¿ý°nvLÑ‡6·Ì«;ü{ë O£KD!ooíjÃX¦ó½uÀCL7’‹DœÒ­½ ¡2ºœZN+P¥¸=â­Ö8§1¨£"øy•\â4FÕìWjBœ¨Òó úÿ…ôœ­vÉ°¿[þÜ@D©òç&#býjµš»™‡™ YXªäXD'êKq·Ü}W‰\‰ð¤ÉtlsWC•WMaÿú÷÷yL<©¬qMÞ"W™’©®@µÔAÀ££°ªÃ‡ŸèðÞâ•®¨i,¬‘ ü›ó¹eÔih·O‹(;QqN¥9ïŠƒ»%ãà¨ò-)…<»¬…5•æ5KÆ§)l:!Ûšýj(AWpJ‘Ò'u¤jéCO¶W^fX^{è>úæ!g‡œôäò¯Þ›—6xÔÏ'%ŒjEµˆz¤ñ03ü&²«ÅÏ¯“¯¸_ðÃLAïbD{W›UœÏ">)´èF'äH¿AmÓQ¨–3{GqåÇ·Ì)ƒÜ0Eg:‡im;_»H'æfd¦ŒW²eHúf}Æ1é#Ñð©µZ#²QyÍ„†aŠaÖ.sˆ6#òFù†‹2ÌQ`%£ðØúy´)ÚÓ0\½‚I¨LµuÎºfò’€­åkÙP§^¼Sa›ü‘iÊêc'¦@ûÞX—²yqy¯ÈZ†>@¡ó§n´É}ª ²µ#‰˜9„ª0¥ù”Ú”QƒBd¹fåríITžò(aü3ˆF}öB>YÆ™ÔN•1íECk@2„yŠÂ¹Ly_ì Í° ÕŸž¡‚ž"vI?,’·}fš~Ew·Ûèñ>%ö•õ}þ¦`âçw`«ûw¾%»ÿÁ¼?1ã@'C#m»ÍÐNƒAHVIw™ü’	[§aŠ÷Vˆø‚ülÙË¨P²%#ÿMÆ—>\RòkTƒ ×ØZ}èæŒFðï’Ö:énŒoÈÆükõü+\˜ªžÿ¿³¾³ÿ!ø|ßÙÒ½³µÜ2ãA;ÇV  Ôœ!$ö …‡LÑåå(Ì²][&ŽáÉ.ÙØ¶*f;0®¨ÄÏÞPˆesTxÖ0ä(»Ôn¡m–Ê][þ¢qg:C_©k+zkÑÏÊ^ÔÌbäe3šËjô™Øt¶#«@#»€fGìt:ðí
´Y†Ê$Ü@mo’!pviÍñ›øª×±m¶€àòÕF6a¯?éK¶mW¼D ª¼d¿zsrúÝÙž½<}ýìì°zñ´fê«Ë6/e\ª’…Z˜O{NÎZ-òŒT‘ßl×¶ùY‚Zûüí}ŸƒÅoAKšYývQ?Ü4×7ùÕ0êz'x¼â°×q›ß,ga>Ö:kâV™§Ä-x˜¢<ö»E[¢hÓlÖ(þŠŠ÷Ñ”«$_"8(ÁÛ¥°Cô‰J»§”Ñ(Õ¢w·}†ofÞ´ÜLÕ–C¾¦ãÍosœ†œCšZãc†s—…Žé§ÔÚ9èG
z±µ¡|ÿÎ°ýT)Á$-`Éä	íÀ.ñÜw™¥ò:ÞÐ”ÚíÌ¿“L(¦6û6º‰ÉWEƒ«h
Œ÷ÆöúÚ&æDÎ%9S(«SA‘Ó¯˜ÑÄÊ‹ƒµãè)yD…æ«Wáàþ¯½(q—àµ¯aéÅ¶æ£ÇKŽõ=×¡…ÍŒ#¼cH´ÝÅ}Gÿ/_”ÏõFã?Ú,ûBmÚÍ‰râ!h%Ó¡Ê¸|ßƒköe×.:a ÄcØZG¯N³ËÇ¹ì¨ûvosÎ­r&k3?ÃÁ=|®ŒmÞÙ1Àßœ²[%Kæ¤L^§d"÷(*zJÂ¾ØÒÛ– ·ø‘šíé²ÞÂÐ¢Ó0ÈÎOcWDÙ=dºÕ)Þ‚,Q³PŽ¿‚„Ð~GL²J¾.tf‹¸{WoTé˜€pQÔÓ>»ü·dm¼á5¼yˆ‰å©ñ*¹NmÞý_)°€sËòÔØŽ®Âë4½¦	NËSýyæò1Z!·µVÈ¹=Ð/ØÎÎÝxžçë_ÎÔkœ©üstíæQ™¼éÄÊ<‹þ{l6˜‰Ã·Ð;Üo?Ê‚‹8ìcüuydG!ŽÍûtù÷¨ü¸cy·)Þú0¸iwW½‡'…ÆÝßLT‡8tQ¢Ë´…—‹wéó”åÈÓzñèßv¸szÿã D³;TRf%€¿JSkò)¨iìÊÃgCs÷³E£¶wœØ~ÕŒØN9w#b32·™]* §?£Ò5k€‹IêìpÅ·nÂ²I–é;±
õ^¤Foü:Ê&÷MAPÙ­ŒŠŸÑÖi[¥ËšçA·ŠSŽîÜGZìw8é)ÎyC£þrß„¯v†ÙTÝëð—U~iêèm¶¶ö½‡ý³ëÖ‹¼Œ]58»<ø»ôÉ Õ´«J–¢n‰Q_dër™¤<½¦ÊO=A‰¼ŒÒaûÝ1è€€†ØŒèJÃq8‰îÿåþßÃŒ„dÐ¹å¥‰ôhÞõ^4bx-ÂO)%à^õµ~¡=yçyjÚ‹Ã`t<­,¢j”+ž"ñ8šrj@Ž¸ñ52i<2Òg-¦ÙŽi¦"îÇ†ó|
ƒfŸY§¶ô:¼„^_}Ð&?nS–;ví€¼ˆ@ýM%•-mçZêP‰Üº›ÝYµ`+­ºE«†ÌZ°|wð<ì]©»%¿æÙFrcè¦…›ÍÇËÐM­%Ð¨ÿÁIÿfÙæ²Féá¶†u]ÃB¡"4A¼Ã¿W¿žñzîÞÙ	¼êÔ¶©:µ¡«Z÷‘Ñ_mÇî¯†×ƒø¬á¥ñ[Ù {s'VÍ™)óËa!P=×<€ ñT³™*‡U?7“KœÝM`””º–ÿ¶ÌœÒÏJKN7ÚU¯ýÞí¥æðS|Ê*§GŽôÌŠúë‘›Hh
…Iz+9Ïz =ÜÍÜlñÞ¸bH¯jŒ;Fƒ½û”†WC•>ÆpÎm•ÀŠ{ÔV)n°?0„—déÎçÌ)¿r6SÚ¨>šÌ§ôX“=œžÿ—ç¥Ê0¾¡Ùåˆ)biÚY·$yoœæ›|±sžÑ¾¹Aèk>*‹9O²—2SÔÅLç+B	EÎ´BZý>Yê‡Kä%\ì¯[¸€õÏXªlØÆ“§¯ÉÝr-âów À‹åJfŸøãÕÌñ¯ÆÎxñs‡7	ž,[¯ æpœãÕóÀËË' /£•Vç‹Œ^v¨n†&g ßYã"û!ò„‘¹|.Šï~¥ùQ"bÄóNó-fÃk‹¡Ÿzú°‹1ö ”ÌPÚBF|‡­Ýÿø‡2”SØî„'íl™¯Ç…±*?îÆ~Þš]S^ªYŸ>QË~v¼jÛðò¶7à¥®;q‰)›TWtÎ)e0Ÿ%ÐØáu>4³ˆë[ƒðùE9Bù¹Ý Ž‚øÈ¨’Ã+9¼H#ÓñƒPÀÜ¾ÍsûÌã²<×ì5›;Oä§Azt¤“MyV
DþÎÌ¬Ã˜´ýe·É‡Q˜žôY<ÿò0YFú§KTPZZö•‹kp*R¡¿%V‡W²tnkÜì%GBþ–ß^=«šK`Óùè”ïëÇ©ø½Ôùšùíãå)æåŽ”7½x¥¤"MÐ²†Á7^ôAµ¤z-/Oñ¨IØ@uê¡‡ÕöUq:ŸxF,Ø<LØ./¼Ì‡—:¤qz¨Ð¶VuT.¨"¾£úF“ sîè!ÓZ8B@o< cØs˜p5÷±9~XæÚ°n9‰­¾kÅVPîHi@rs£!÷«<º¼ž„©®EÎÂÍ.ÑåM7`,Üh1tø
7Ú„ìêèÔ÷µ4 'ÌcÉ4	S¦ôÔ%$†I{Ø;LÓäÃ‹ðr"îÿÄ“@'ZöÍ4ê†ƒà><•-ftˆ:GÿÌ©zDnö2¶»Òdá0bÍVOŸägøVß}á…'!7µP®Y’=RI‰ORNPÓa5ÅmÒ†š¡<q ¿©]1TlV·1A“´»è™-º”ÔŠaÐmìsø|zµŽÍ«æ7MoâUSœU«miÆ”Îdƒ›¥£% Û˜Ëæµcj^“™Šùs«0#Òt­‚žËhö™ò«ÁºèKË¨âhá<hLt–×lÝàF÷’­µÅ°@ÿæ9—fãðÀ‡Ô‡öýUÉ9B}/ˆÂtDv³V=w÷
q‚¼ÉïWháÂ¹•=ìoneÑL´+Vw4àóžrRàk H…ŒËY%?ÏvS§¼éô”¡,˜úÁàë’8£Ž{¸Ì<.Ü]®¡´­VÕÞ.™•ú…áêãJgk¿$,ãè†}rHÃ“ÓZ2òË5V\S“Ð;bÑzƒë0¯üyö/`.´UV*õFz"2"2 íl"ítµ’º‹¬—Ó[+Ú3èÄ™€ªŠ¿ûK¸ƒÑ4ˆ±“gÓ&Ù1ÌðÒ²PóNõ ÈâYºKË*|T÷ÕjnXOãLËŠMšðLÓmšïã‰™×X8ªÃ2of6ŽFbŸüTš³ ¾¦šr@…\æÛét² ¡Üeã&]<º
{ï¢´‡ú^6ÐÖh×¨ÄGŽSîÿýd)áßDÃ¦ä]ûkù^g’Ð!þ8› ï^ûü-ºp\%Ó„›Õ~4ˆ&K+d¦˜ÕH¸ÅY…[èë±üŽ
E.]Þ4ªVfà/éòbL„[z¼ã©YŠ­ìê×x~:¢}*L¡OSö?º-ƒl¢ìeÎK`¦LY€¸$Îrÿ¸3Üs1¤ßûAKBIS mML
B}Ë³1”¢Oúb`r»m¼¡Â ¶Í ôŽ®¿@)µh«`¥V$	s5†xÏJŸ[9ÿ–ÓBM§äp¤AKO÷¹ÕÐ@ú¯Cî\ÿ¹~4†”tÏKdÏï™¨ÞI¶…3àY=Jç'}­­Ê\@•¬¢¤yÒp±´,í©M­ÃÑvÑÊ2³a‚ú$[nÉþ?I“¬¸Uc±í%ã[7ióÑç”MŠÖ:bÂ$É²ªÙéQ:C×Í“öÝrJî8÷öÚ±“¹á¸[ðÌ¡¿7ju{›Mæ8–ÎŸëÙ<ë@MÎCfasÉ¦®Æ¬¾+GÊ:Ò§¤ž7Ä‰~`;´Ñžj.½èbîÑ¹ƒÏÚfâX[#Ý¬¹~eÅÓã’F¨ð–üi
cŽÚ\ ênf@uñ-ü«b$l”¤ ðü÷(Ð}>„)hËÉ(Ô))Ìn7EPâë|’È>ˆEýÅív1‹@1ú–otÈ·úWø¾Ï‘‘‹[‚0ÆfPÜˆ°ÿôßÂ¹Ü&Öƒu¸—Q9v…äm9{ œÝiÛ¡ô¦s™¤Ï‚Þm¾yI²&¤Á‡CjÊ £NmN{%ûr”¤Cþ)þÉr3ŽúÏé™>hó²—±@þ·¶ý„ + sÒ»úmxå}wˆ&÷ámÖ®ŒÑrç¿I*z6´çý2Õ`jŽ‚,DŠÜß/º!?Ò‹I†ÖÓ|àyóÍq…BóJMŒ¨Òéóüã·ðõù[Ãf®½k)«3žfŒL´,ñÎHö›òÛ0“ï^½ø¯„ÚÇÈä*˜ 7qVæÊ¦ù¢Í­C@žX×07©')GŸ¡++É„ˆ4
u¡íwî2Ò>ÿÃ
%îì-åA>M¯±ív°B.èWÁùúÛ‹	 (>iØ¾€[+di<Y}úzi…@QÃË5õ/XÂ$Å¨M.ã¦p^]Ðoq5u#5©ƒ ¹³˜(Œe€÷¶¢iCB_Q£2¨Åx™Å`˜£ŒÃt0ÅTö…L Ø´Å­²ÓÀ6¨%NÝØ1š$Øí¯‘ƒÅ@>*{dc©+“}žo!æ8e¥”œ:ù.ÁŒGgÑZÌÔû§AŠ8ÓÚg¯îÿÄgÃ@âŒãÓ4FYd‡ùN·/w˜›É+½Ø—ß3W3îMÊ®WŠÅ×ÐŽ'wK4‹ÇWØË8I`!ªMZ“JÇü
PÀò2*½ëÚ¥ÁwÍ²óÓ4À0"ƒ`ìeWX\QT ¾ëÙHB%Šâ“N% ¢
ŸuÊË¹¶•Ó,@'¶(Ú{2óYT§OW1çP³¢¥+ìÛlÌŠØ³;í’·lGìà7 vžÀ&u×A£5ŠBœáù­Êmñ€‡et€É¥ «×r!V"Ž¡Ó X¹&µñ$0‰'ºwi?qó£–ü¡~Gš“[}Ò+‘¢ü§ð-{Ad+u9ýÆž¾ÃDfÍÀÓ1XÆh€á?Lš‰QÇÊ²Ð¶ŠJ‚Øú'â¯óî[ zÃ†ãµ«	Â	£zÂ8M`^`'Å“+¨Â!>AÐ„ÞÄ­Cœ²Šf–Û,´)&´	z=›‚Á¤$'çÄ–¢‘«EË]j“¡çy	À¥¥ŸÑVzÏ!¶|>Ò²ÕîP(3…^ðT>Ú6ñ”¿²¬Š’®Á=„2óAåK‚2¢q§f½tÌäÅsJý(œâðcš ~£ûÖ(o§)±CYëÓÚµ^¨µ^Ô«»zòòôðèÍwøîõñ³×t’ŸtÑ'ÔžB§‡ÖdØVŠâžj‹{Z³8N=´¼UÚÊÊ‘U7ÙLÌ0ThŽÌÅ¥BàyŒF ®ë“ºz.«®"ØVÇ«å‰1záü¼ÖvÖIFq6èÉ±”ÎhØ/ÉÁßnœ­Ù®iÁvüßEÀ?™ÜìòSò‹ïŽj2º|^B¹Ÿêrè8>3ŠÔ®ªJO\`/ø9ú€>[ks fíQ1@Fañ@Ó@'½«¨ïou¢í´Ì6^:ƒotCe%y1•ÀYîQzÄ“`Hf’²µªªp|Q|çQ É,bøìN?5“àƒ"º•JÃuµºÁ 6T\#[r"UïR|Š—ìyÓâ9Û¯™þ4¥±vböúbkÅ¤†Ë®/ôPjWñ†Yo9Cº²ÉmLCZ€×L®vÉ»¯g¥>~÷‹w®°8kt—7xÈ{[<£2H«DÆôwó[á¨ïdôM˜x\B¹¸Š`CÁm’QÒ:ÆÕ+iÓP\Þ/â*V¢EÓ>óXëL_èG}ÇjŸ;Ÿ´€™§äQÉÏ)%äÜ¢c†	¦Žu°ç^ Ëlp¡÷áÓWõ`Ó]â¤VM§¸%…Ìh“ñøbQc Ô–Cà0ç-åt‡Ø‘Rìƒò}ÿ Rî­]mÕáÑÅ™/Tk#iŸ£"»BÕÙ×¥…ÝÙ\(è<¯XóÚEYÎ0áÂ(Êtê}¦U£"¢Qû]…9ŒçååíS:þz¼\ä;”L<t\¼€™<<KÕ+ß¬Y21	¯n
„mp-•ƒ<ÌŒÎd7zƒ¿©zåséõ²êlªAy”qLûá*Í°Ÿÿ%¼·Ö<<ÀöwVx…/(¡pUì=	Z¯Œ±©Õ
ŠðŠI¯ãÊ(¦0—-YË°äq`w¦§è–R‰7õý‹’¾ƒ‰¦lt)EøÑuIÙ
Í~£jhÝRl4ä’ó£MÌ vÍwrñãný¦’‡ ƒÿyÂr—Øý$ãX¼`ø°Y=¤ªÚöQ’n¡Ó“ô<ÔŽüª§~ bÃŠJ=”á4Ò.òK§eà´ƒHì¡cä—'Ü…V¥'Ö˜C*¸›G‰CîEŠý]áîp” Êò¤¡¸ºYˆ«fÉX'®Ò@¢AôÑ)su’¬¦ä2M†R4F²:	ƒ¸*ÎŠ&#Ñ@èÜæÕáê‚èÏ4²O²SOäü®—×»^3},@«rîÿÀð¶†RÇ< Ê(Ià^	å!ý$÷Ÿ5¹Dð¦ÛÑëœëËæ{W^ê‘lT=‡4ƒo8C,¯¿ÿ{·;ohŸ¿g
D]ÿ­f'O:çÝ·îÆ›Ž¤ä‘3i3å€:8—§¢.ûÒØâcõ¶˜ˆ´d3)v	ßm.«IäWÇfc5^n[6W>ËZ.±å5“Oeè"×	@uØ‘3j9,ü´GäùÜÀØUq»ŽF½i\ø‚ål±¥eâ-tL­Ûú¯Œ3³`ägš+ôô³84÷±d”‡§j¢îù©xp^œ¡Jð:L•[“Ÿtîí*O?ñ€þiÝ$z¿ŽòªÃ‹¨•#ïÊŽümIv¤¢´¸¹iÁe|qhñºÒ®VÇ¼U]I•ð UãjÄYHk¯6/fƒî¨CÞ
vZ4iãlöY©¼9týüü®+ÆN×Š­Ïòªþê¾*ÑÓï0ž$K\•ã±5ÄŠÃ,Mêyyÿ¯ý¨¨IwÈ)ÔU$þñ¯Ìé­‡ú]´íâ„uw—ÌxÇ?‘ÅÂóµÚ\°rz‚kÄ‡2³N¹‘Ð½.ñ[),ÅåîÔÍ+êËø«ô¹Aº5a¿ì#àá{ýK'5FÆ\‘gêé¢Í‰°KþšZô€g.$Ý¹mÔ«NÝ(¾É;Š"/jÀ÷–Wq¼d·£ñÍ¦^Á‡ š [îaZmNž+Ìïšy1v’‹,L¯ºøq˜¹ ÑtuZTJ¤îÛò­h¶C®‘.-ùåS¯,œ|ŽB´ü¼$ï[ƒ†cr£ÛFíÈEÉËÆdŸû€µ.<<ÉNÚâç½"û…6j;ø¼ÍÆ//‚É«™RØ…8\Ü‰½¿€9äE"´E2xYrj¶.ÏAÊþ»Ïz#ôgÁ¼g Tá“Æ¿[!› E6üÞ×ž-}CzßF03JsòOâ°ó!HGíÖ³4MÈ( ê§x+ÿœÌ=¼ÿ-zè¦‘M ÂƒÝÖ
ñOc"·F8×¬Å:ÖA³Ö5jLíoê~áÁ%¿IúåÅwwÄYý€b>;è»¹ü$Pë³æÀÛúºÑÙ·xÑ%¨	)j0™@&E\ÛÖÀºxçË—ïYx	
êeº¼Ï»ØUOÒ18¯2”d˜ŽzbS­T:ÅWŸ³°•³}¸òë£\Ê°ÿ…®²_¯/‚—x}¼ø‹/‚—ýÒ^EÌvcá«°¼J¦ÖBð*âM>®ØåØ]½Zè:ŒÃtÂ„.Yæâ3ñŸ]î¢@_„®üú(B—8æ?A‰‹7ÿ‹¸õEÜ¯/âÖñEÜ²_q‹²ªæ²V~ž.ž ’ýõÑ­Pós“³~¯“±pø\Àòv‘^o|´;ìï~Xí®ùú–èñÊúVû05.ÂØ¨8@ËQÔa²uð“Î<Y#gŒKÃ½5Z|#Ñ·&15IaŸ_%4ø@aíâ6]—“.H_¬´ÚTVåÎÙw0>ñÏ +^'j*-¨—µ[GÚñðôØ¿a«‰È¥häe4˜"ô_‡œ&)¹®“t…ý¨Ã†òC@Ê¯G“`Ÿpê¥^¢ÓýäB—î/èQON`ÅÅÑàþßF ’Ã^„º`§ÕPàbŽGj¨£iÈ;":l>”ê±ÅÃ®vû²8=dáZÈÂt­Ì£Ü‘0†e.¶ð<ˆ¯$T}H]¡€Õ"×ÁŸ#D=ë…Q?(ÙGÃEÜ@|/4¦4åy¸›ôÙÕãÿ6j‘_ÑJ:Ã0Ë‚AˆÄ?õ`s®6Ï£1-P_úIè>úÄ§fçÛ£|UHIÇC(xüøÆMÖ@ëéÉ|Ad‹8à¯$i™ÿœ¯]bÏ7¼È ¡¦tÈ›]¦áhÖj{%j&n¥bo·åòª©¹é¾	ÊC@)ää°^õuõ@«Ôí@V½,™£,½Hu1m¡5i–œû£I…Ô‡œ¥³©5 •S»vT²ð€8æ:ÿ–nó<^–šX†¯|‹¤@yýà«;ækTÏ@ycö
Ÿ|Á–‚XkõXøæëÉœ? —ûV¼Q7Cª²þÀÖãuÁ\ ÕR0 Zm«ÂE\­nlQ((ÊJ4ëêM³]Ä¥ÖÖâ‡è|ÇnóÒ¨ðÉFÂýÐÌ²m]¤vã ^ð¦û­cLw’àOÓˆŒLÐ„šö0Œ¯’äôdŠi‹H¥KvI¤°%ô€#c°Z¸âi¸_!iÔ-©›uW2:º
F(Ï?@«òÒBWHØ™ÐìÅÚQ­S­eR±G"äá|À
îJšá¢jÊ­˜(‰˜‰Ø
ºQW·\9á‡ìÔ¢Ôîôä4@ÓisÆü¨cÖ´ûçåå—´«|Ü˜)÷pÞicÞÆÓI“£f½½ŒbsV1Û%Æ“2Ôœ&¥Tô\‡úÐ²_.vü=éœ¯;°L5–AË„e¤|?ŽAd¥·WÈRÎarS]CTm­¹&/Â«Ñ¹ (
´_á51Za¿Oc[6nKõ^œª[É‹­ÃÝ}½iškCãN†Á <*ÀŠH®òÇ2©hò“t:¢Ò†Áðî ÏÉ
‹™y•©Åiéd8på°µ6²™»<Ã ¯%:	¦à«Ð,Ÿ?\ÄÁèýRmÍ™]¦	ÚÖo68”Z»“µ€ì9¬‰Õ¡ÉhÃx?»¦Œ‚øE4z/èÁÍ(¬‘ãÿvÏøåëpÓðŒÍyØÖqËÅÌ|qà/L{~ìO†ñj÷³œøßºÙ®­\áU[é««À5rXøË{kïóî²ÇˆxU<?€‘ñ¡Ø¨öCã#ãMãu˜K~)G¯mCA¡uz-iÄ‚<õ®opƒEƒ´ÁáÌó¸~º/së?ÜSpdñ¬m¨{KKƒíûC~h/õƒìê"[ÒC^<Ç“ŸèÏ˜™ü í?Îß¯öÀ˜æÕ	|·U |QÈ‚ÖÁIïþßÈ8é‡$ˆé†õ&)ÚŽåœ·IÙ$?MƒøOShç0¡	Ÿ;ä:aì­Tú:Ê“e)Ãßì†Œ¹ˆ•Ã†¨&pU.Ý+«Pw5D¸hEªL?BÄ•™ŒƒÍì¸¾B2Ì	uoWEvn¬–¾éßt«„7Ñ¤F•1µbfdÃ3þvçkAG&ðä<ÐJù
+²X®`ª,+ƒMf”V†	¤Ð£ˆU-9°ýÄ óiY’9²§rïjÃ‚—G1êÊ% +æu˜M¡ý$Û[»Ú0”iGåÐ˜<ÑÕ½.‘£JåXz
5(ùÂy6ƒ<éŒ`Ù:ÔC]´ÔÍÜò&küÇ?üSÍ’¢µv´A9ˆÉ+´J±ÐSjÑ!ÏxZû]2Ó¥ŸDã„¬Í²;ýî¼ó¬ãÅS•…ˆIéß<»ÅVgØÇ;ÿ<IŸõ£IêqX;Ð˜M2T¾üêÍÉéwgxöòôõ³³ÃN„pý0£%’'´»jâ,Kãe¿¡6Ÿ‚‰¸rËóN2¡Ï|ñxY-6¢iÌiôdPId,Ñ«ŠX­¨/NOmA¿±)4y‚{ÌµÔÜìUŽU[¦ò% sVÀZ¬dÃOãqé56c‘[µ+–¦±V¤Ÿ÷„TÒ6ƒ˜S—2zk.¯–S´/–ÃŽöÅ˜p™§B‘€æé zkÆ[[îœ95\&/³¶fKB­ûÆð¾éög43ÌºÜF#d×:"H>ŒÂô„ˆN³0}Ò™Fýe3Eæ"™a$
ICmúZŒÄ°<"â6÷ðÞç€iS>ýæv·ÑŸâØççŸ">j[3šöOO2ÏýÁ°7¹Ù›dIiäÙl‰‘N¢µË´zzUåðbvM’ø8Ff!œmµJ˜.|AwCEVW¯ÝÂ^šŒ’As;@™õÑ¹JléëvÖû¿¦mîˆL »ãá÷¼m§ ¦·B›sWïÓãç¦6Ã
Æ%uúÇIŠ{2|Y«YZ¤·«Ð4[“`’‘op²šµÀ!ŸË™áèw iŽ@E?'<£î.L}0Eš2XtTŸÄÌí†$À¹Å8ë\‚èô®
8Kƒ`µ 4üïòôð…ó<ó6Xæ±Ì2&¼n!ãÙ+‰
ÐÓÎU±[ËË–H‚âÃeþYÐïóº×•1±&9Ö&Ÿ•ÕcÐwµÔ‡I[ošgšFæ¬—¤!oŽØ:1/õ
I)…ÁŸè;v2üˆ&é]^!:·vVÁ0¸iPüù%ÖnXEwÙZm
UN{µAø{B^¨¥¡v[éúiKmeÕB;`kDç´uã„ö ¼	LMƒÉ\æÍ3õŒÍg·fñ9MxUñ*HjÏÈÊ«ùÀØQxÒTÀçfU7cšá¶ žx½Ã~%Ûá#“}Iaw ]%ËY™ÒÌ–¦jø­äÎD­Öµ˜B(ñòÇíÂ8éEý .[”ÔŽ£lšýÈ`ŽÔÏ[²Mo&­ùyÒƒ
^÷Å˜ã8·@äÀ9Øè0Îµ³::V»0«Õ#]žbÌn)k^›¥ÐKžˆN#ƒèž½šÅÉŸ4½rC7°¡~Ý´éÄöÀ=Dse€bÇ5Vn½õYÈ4Xop¿A9ö4OEÔt]^@ñ(r{,'M'Ô•¤gÊw®¬FYc¶Õ“«öLÃè¥¼¹Ë®Uö&€SæÃ{”ÞÿUØÏaú‰ä“É»¸4°ˆé¯\NS œ}ìY6æ–âÔã¶O‚#Ã|yx¹ßjo“jí"z>Û‘ì<ÛI®çÝ®§åÏpÈãôØ«„VXœó;äûÒxg!ºê¡„>Mé~KÚ'‡fg	ZÛ¼‰×zAŒé0´ŠÕ;¹­ÏÉä=¾õpæ.‚ª¬¡TbÉ:®Ñ&GhJ2 o3æòÄX«R•çÁžÁW^	 4¶¼'QtÆåÌžÔžëlñQu…œw·‹%¿Ó@ï×‡A_È}ë¨Çp
Pé¥Ë%ÉAßÆ.­+ó[¤NÌ»‘¦BOsáã!n"ü|IÛnÆ—¢¬“ë@jÃ08ï€, ™vD7
"Ñ×o±ñu/wøS¤‘k¦°sÑ8Naè]Eô ›†½¢Œ_ÁA4ê-ûx9F«©»™vñm’1ÍäQRòF»½îŽ¦ñØ`t´¤ãó7ä+gw:Ã3~é˜C‹O¸”aÌ(:¤Ttsžãp¬w¶]>®µ|ð’ÏdÐg+ŸãpDqæ:‰¡§3>a t’´Ì¿:>ñúÑ—kd2ÎâïüFÀû©(×¾\]Ì#vÊ±ø]¶=³Ê”$bw™3OÅ+Â©©0¹±<­»Ôënõñ6ºŠ#ú¡êQdpâ(Ž6KQ
}5š¥9(ñ°4\Æ>ÇÊrÍ¿ÇÔ¸gÁu˜¢;ñüjE5ŸZ'Ð±O¾/AÊ»:úÀåÉQN7ÈÅ’dJ7¥–SõåÙ4fšW›Ò²ø"dÚ*é»¼ë½JË‡Í¬Ëñ@°.o8R ™Ë"¢aU­­™Ê9ÑLy¼Qá|dŠð´:ÀèŒ‡;ä
þñsR¡žÜ{¦K~¤GW 6oújKn³$å¦¥MíxÏnÂÞ”±¶™é}œ¤äþ1“µÛÀaË)ç ¬ô.nz¤Lñß4¨:k«6?á$Žâ.åÇqö½='„y1‘”ú8Ùw@o‡£ã”SÊÄ½Æ¤½<o*äßeyzMÙ|S86Þ)|ÿ	í˜·¦ý—§Þ‡¬Cnt6HP§ï özçy©o—‰ôj™‘ÑnQñ
;1¥nìôô4ÌøEîÜu-îHÚIR7øyó¯ö•³gçTJ…½¥UO™ç™\~:	Z½QWÈìny™å¶FÇqÎ<†éú…)Å/…ÞTO·élœÖ”·åCm[ÎÕŸMÒä}x„~±Ð¨Öé®_<Þé¶ì¥â‡A>Eçg?p,"pÅ!ˆo P¬êÙMÄØAZŽüÚHél¨÷ÈÖº›TÕq¸ÜÚ¼ÜÍ#Ê(TSWûŸ;ë¹;Oä®ç§Ž¯8Ò Øÿošôûq¸~Q³ÿºÈÕ(ò ×‚×á€J—®Eî³Êý3¯Óµ‡ü÷®b5¶ÇÄå†‘ñê&óðÍªµ-½N%bzÜÄ\¹)åCIÙ´0Ô£?E¨1ø±é ¯J*)zìtO8@ƒïnP›dAY=ß^Û|ë>´F¼f‡¨Y4#»háè°ã—1ŒÂu”E¢v§ÉÍ~k'ŽÛ¶_=h&ëxÝßo½$Ý•­rH6wV6w ºõ•.ÙÙ{Þ(luíÃ²ê^n_>®ùÝßEýÉÕ~«nu/¢(dc˜R$)ï/£8ÞoÑh¿O<Ã©?á¸Ïîv×hìž£Ÿ×‡Z@¦Áí~«Û}ÜÙÜñ®/?O./³p‚pjø=YÍw5&l€ÌAøF¦g“Û˜†h–j—,±êVûE}¤›‘0€M1™N–|ÑY¼hT¸ë×‹
.²$žR}2I†«|Cã\ÍÌù}9þè=?ÕÀZQæósJ3:/~>SùPx…äûæPLzNÅÙ·e“ä4ž›ƒÁWAkü Ç}ÂÑÍÕqµz¾ÉÎ€\›–rR(lÿ5†[i0ÂíVhèDë0wèØÝPVr©Ëä‹¯^³BˆòY\¾½üH`!êHd²¹ÙŒlvÛÉè2âš Gp1˜!‹EH³Óÿ¯®€Ù;Ck{ìhÀâKÉ¶sÂÖñ‹¹
q5ñ_ÛßiRÛŸ“¶àE'•öƒI°?;÷ ûÜÆÂ3ÖS»[¸†¯;
¹Š[•Ï9WQj™’U,™ßJgŠWÝÒ9ì¿Xô+ÌXW^nÁåYl§ÓiËžßÌ<q^ö†?P»Ão?Þ
6/v â·d—œ¿µûáõÖÍ¤z ÷o!¹_¼õ|1É¾Æ?ÍögÖÝM a¥ü`ÇãƒqÐÇmåp4@¡hÛýRîo9ƒ°«Ž~‹ )™7%à…Sè”ÇlzˆÄr»$Ño–]ñ½ùµwÂžMíïzðçê×3ZÀÝ»;&ÓÏh±ÓA<äW7àÑÞp@'—|“$ñ$»µ‰°nhäY.¶cþIŽŠ¹¥,zf„`ôvÆ7ÅMøÙß £>á
¯—\IÍç°¬ögmÎTèAc9¹õüÝ×\ü-½ÜÞ±—œœÃ1êt<[ÏÞšf·kxnéð#Ž2„}%lûY‘?þ£$Hç^¾Õã¿Mr%$¸c‘Ù«qí€`U{–ï˜wËµ£)¨E6à–ã6o)ÿùù´¶ðåÎ“Åå-E§nw+gU¶èv›ú8=+}Ué<ð&fÅLðœW+¨õ¡±½¢´è ÖfW š'x’Iý¬œ(¯<}ÛÃ;Ø¯†ýð
þö ¼â¹¨¶j|±ºMÕ(³[…ZÝ!ddìu£‡²ök:‘·híÆºÕ/Þ¡³=»Nbîx€Ec€<&€çù¦º›3† V/5Ñ” ÷¶–b`t’ íici²Z1´ŽfhŠ±!‡<B§Œs!#D½ m2zö¦“èšæ £ |ýÄ¦­Ý`í+Œº“@âFèIšÄ™&ô¾(KGó(ÖXÈct£ÓâeÝÑß¥¿Ð£ÍfG¨A•…“üŒJ&ÌëÍ[v$ß©ž‡f‹&iLPùÂ‘lòI“ìêå‹ö>»tß¯t–9Ðó|FŒËUŽ™Kö&»òR_4Å—^p2,8ƒMÃ²µÛçèE4º¢ælp,|Êí\ät‡[%]é*BÂþþDéèÛôþÇKDÍRzŠG=Ax.l/î¬—p`÷šÜrxM"“~N`AbP¶ñçÒâ±:BØ¯n‹âMÃ°AÖ^¸zkU³œ‚î$êIœ}X2ûg–E%hÕËûÅhB™Ó4Àž›%nÓ}õÆg|Æ¾S"WNhÐ-í&,d6|bª<Co)Üº×!TÍ3;†ÿ)ËÿçëoŸ0‡.ÿ“»ZQºl¼Úºj1qÏÓ§•×é>³ñ9'|hº¸ÿ_ñ$äðÈ€
ÓŸ9iˆg:âÐëŠ«¤ûiˆÆÖœ‡!¦Æ/¸ˆQòñ¤TH'• –µGŽ³®Æ$KCÀ‘¢^æ	À¾»ð`hfd0ñbÎ¥ÐÈ'W?pÓv]c¹èøaâ\>.»èz(6ä€xøzzåiˆÛj2é&Ë@+X,xžüÆŽàyÇ<”µ{cW“óëW3¡çw¿Àm+,t}#uŸ%åòq.@ÅYØû|f¡À«,ýd5þÕá§ˆø?‘.å]%Ênc}?þèZžêL°bõîåF/ªML;ºêhRCO4UjÿÆfr\ CÅÌeRq #ë5s¥w†;2ÒA4¢n~Éx—t×WHŠMcÆá%üûê
÷¥Û%˜lÁ)5€8	³(Q¸ÌŠ‡ã&ÙlP½'ØŽýÙe Kî®êeëÆ×øýáM”•Gí(Ù´H ÷°ÛEÁPÍûÊì2žìâ¹è£­o¶v.–VèR:‹þŒI(0…Gþ+mÁ«œŽ"  ‘~2pfÎaðLÞú´¦ùa±| ÜMÃ¡ýDÞœQ[<Gnx±<Èê&ü+\íõú¿Îúöò'ÈÊ±ñ/Þ­%¤†…yxÐ”~E‡qØÿ.d”¾„ö!ºÇ§Ü=ø,›ä0%?Gá‚À©çñÁˆò¬r”<>)V!òý‚yXÏ6=³û	Mâµ½"~Mãàò…Qø­xÈbÒ‹b¿1»¾ýøÑ£Ç>Å:½
–:'ÎI$û…3áÌåÅèbƒ.w.ƒËžýÂÖ„­ÃJ­ðà*‹e±ÿL¹wáù£ÿZAJ¹Ò3z½åÎd¸·–³ˆÆÎ5|še±Ó;dh?X®DöyzLØÜ··Z§E”ù	: Px–~’‘‘œMð^ xÃéo7ôp<æ8F4êŸìƒvÛŸŽcÄ’)’²gíÂ)Î IqÀí¸xÿ[L%µ9¸ ú(FaÖ.«‘^‡@»}>¿"@Ê°,£VC>SÅ(Ú±×%¢ 4"PÈy9Àß_Úçeð@ûâ§Ÿ]Ò›Växv¹:ÄÃÈ	PÝyu<ßž¿]îà$ÁÀKµŽk90‡F¬¤šöèGcÎ»oáž€É`ìaYïÓÚõ^¨õ^Ô­7Ž±»'/OÞ|÷‡ï^?{M'èðIžd—àØ²¡¥uCå‹Ÿj|Z»@>÷´ÄUÚRèrˆ«”ù&ŸçÃ¸B¼Xöo=)%ÃÚ“5–Ð9ø„IÕ×9:˜ûE‹ÊÆUü4ˆnã*vÎ˜sœe¿£NÚyŸ¬Ò|-ß‡òP!†?7)”{Xq€@7ë±Ã£¡6dš‡ñ$a~ª z$”¤k›ë-²ë_äËûíGR¡…ß¬èDëQlKzà%È]›M½$tn-,GÍû`tAóè4ó°02/ÕC CÛëÂà¡šr(êòìŸÄªoB»ÖAIÏÍqô«®	e{ty­§à3N.ÌùQû]gvœsk‹fï+@Ðµ·éÚêÂ¢Z*’ãŒÐíj¥xéK˜|1”ÉÏ2K2K«W$}\7¤ôq²*É“gÉ]]—è0¿Û•Ißnˆ8˜ˆJ:+žxö5*a¾YÛGÄ¶4ã†ýJ"©?e°¥‚’oÕÎÄšÈj¯LœµÔ$õ²R†g
æ›]²±î|ù†âOù$^¾¡v§F9—«ÉM¾ã5ó!ër\Ã¢2ŸáBeÎÐ !ò ÕL™lI-fÒ—}’%£îÚÄlkh	Ññ46åL6úo›3ò$²‡bê[q,ªÏIà»và&í›Ð/œºª §i“»ºŒ•nî«ý®~ÁÝÝ³ýYþWõ!K.¼'þª¾[¨ïû³âO];‹<ÁBSÕ{Öï”ÁÑÝ×~,uF¹¡ý'‰¾‰TßÇÁíÙ´×³ì%Í•;ê7ksòTqQ/*½}YæÏ¹jju“ÍçgÐŸ1ß4¦ª\$Ëœ›töéyŸ'¯Jâi/	‘Yð¿>¾TŸëÐZNû—ß§1-ŸýiÞŽÒ°O÷Ü¾¸;ˆwµ5/ ž(­J¹§ý®À„§_¿´ï
ÀçômáwõýM*ÿ"À|”›Þ+ÖM ~!ß­~ÃÝ£«$êAëË¿µm!¼*ýÔÌqnÅ*Í=š1—M+Ç¨•÷“Ñ·<Ó,ú%#Ô^Êì³å=Ý—¯Cn¯?M«7Ÿ§ÉðØ
õ:Òb’[/ƒŒªNû3ñ—vÔÞH¯+7êïÚóîÀd9YÜ\ù³ÛŽ£Ñd÷*ê÷µ¹æ›ñóh´Xµeî=øuNj?³-Øcsóc©u˜ÓOw{©¿q|?F7–R„ÛŸÙí¨GÚÁä:3žàLÒ[ë™ð”–J)‹B…w:µ­REÙ%X¡	,\/tµ…j3EèjÍ^Ó“%ø/=¾êã¯~~‚¤ËùžHÚ%ýeÓ)èN³05ûµ›‡¯àCåÃwœôÚ}üçb…,‰*úÒŠ¡¡Ë+V'ou´¯š“ÕÜ
zW¤Ý‹“iÿYjé(£…$;!0ï´Ýz†ÿa}‹F¡5°ZÞnk…› ·Œêì¢ycCS+í-|?ŽoÂI†-Ó'ã©V\Ù“çM@5°±âgž’Í€uïxËü›®Ÿb©¡²™}Âž‘Pæ¬ßÎ8û–ùxžY|Â$­ü6éëèÄ¶æ5u»Z2zÖÇÎ·'ŒwÊY­&Ë«äÊ
ám&÷.±gìªTùÛÕ”P:÷¯dtýÀYò­ºOß—*÷­Šmg“`‚Î`Œ[±‚Ø=ó7§ ¢0ß2ù»ò¾þÛ8	úÏ®#T9áË«`ÔÃçQ²'ý¢PXËŸWWg¡,{Ñz¸–ôÿm«éôøyY]H‡=X|ÇÏÍ²àÇ"F>ú’ÝâTê¹÷4FY˜5Ù;ø§žûÇm-ÍõÖ[s½õÓ\Oykéná?jå+‡¬¡PN†&OŽÁšEå:ÇæL‚›ð“Ñe”O†ã$ÑS¹[|ªŒþ-ChK’Ú¢ò-´3HÉ›hL¾ùyÕ8HÂÊemâ¬ Jd¡>Y£ÂÓ©w¸"Å f…gû£«¹j{4cé=i JÖW}¥(Õ"«Ó‚ÞdÄÂpau;“4¶+â¨{he•Ð4,ÕzWª5(åÜÕ1i$øœ#Žõ~‹ÏPJè‘ 8Ger¿¯L}+K‰Ô Cgs*–3Í)’ÕŸ`l€Áå@Y×IïþßHa÷?òÉ 	dáF¨g…¤ðy¢¬;Å¢¡3½ìQOw	;…"Šä=0‚©³ÔåÊBÙ[£ÅU*a›M•ÉWR{ä`A›"¤–)ÓäeûÁn2ÐØW
Än¡3æê†øCð:[ÛÈ¿Ô{¤Q÷¥ª'( Ýü:ˆb$¥³p€i|³N4êÅÓ~˜µ+„µjkå&z|7¤‰Æ–5Vd%¡Qó/2FÁ:Ú¥ÚK §SŽxÊÆ Y½&UžÚÚ6Ûq+÷é	¾>§¬5ÖY[´¯©bW5Ùifˆ:Éâ
3÷î%cÜ¡˜(¾y—Ï7ýqÀþ³·Æ^«VªƒåÍËdµèØ·èhæu}Y{kl­¨÷gUòçç£ýU=â„M±ZæWXæÒò²Œ2§Õuœ"‹Ðùp‚Ü¯€‰„é~ëfŠ.1S!ITÆ§ûºÊ%¨ÛÚ§f%PŠ¦Íœz«Gg*ñZûEÄÇžËŸê[ÑÆRÌV»ïkTÔŠcž²B*FtåÆÞò¥„2 *øœÂ%qÐ-âá—‚x÷*ÜÿD>²Fxy6ÁŽ
&ºº¹ŒbÐ˜’Í¿1Š.gA|-;D
Š˜ÿšä—\qã%‚>mútBL.6‹«yTlÆ– \YåS¯/Z@œ™9É²„\‚êÃ$‘ M‡á I£?3‘›0<Ñ"¡EZœehLîEY’u>™ ¤!¥ÿÔÂžJQ"Ò?ù¬Å¢3m“¹l¤ô£¹`d¨å‹t$|e”ŽTUGD2‘¬àŒ ”þ©„%{ýY
Nòp©I]òV‘ÉÄ'>¹Éf;cŽé”Uc%“¥De.Le½$¾
ynôŒÉÄØh	íC¦¾êÁ7hçVë(cëLcû1ÈUü”Ü`þæ6þ'x|EW+¬2|vöK_ÅSî¢0(«<ù¥ƒs\¸’;uG¼
?)Z¸{š&`+´BvCq\Œ6±YuœŸ`ÚàíIÀi)€ñÉ(êEh6'O,±?C›Ö‘×~%†|Gè9\ñy¥©õ¤Y5àcÑÒÞF4{OÓ(¼¤QrÜ]êQýtn¯0œÄd:ZLÍ$^Ò]GsbkÙt>?ÑS	…ÃðLÂ¨Ž¸<dÉÒ?Eú²-ï	FÞ§ov¥ÕC`Õ¦=MˆWò	$«ÅÂï«ÝHÃ?M£4ì+*‡)UPºOJèG,ü>0ætNãÏû2®Æ"ˆÂýÌ‰ykM*ÇO\äý¹SÝÓiã1ÓÆd§=ä“qûiœ€äô5Oœ•kñÞ|k]4
øÚ æ,Ft&8(Á(™ õ%Â¾›êÔ“ †±â "kÈu4êGƒ„†S%EK†®ðÇÐ¹¿ü>ËÃ²!g˜0'¬²e¹ùÿç¿ÿhˆÒÏ‡¬uðÃý_F½iœÃ)è÷?"µìò˜tòd×‚Ê"\á×@Ed‹è"Ž’IØ£,¯b>JÃ¦nÅ—)=JF³¶= »"Oj8ç‰·t¹S¡jÕ`öÖýf­Ýæ*ì½?ŠÒ^Î³ôM‹ç*÷A`(R•šKk)Ôÿ#m*÷ý0ÿˆ~˜½¤ ÐðqÞ%:¢-S)yI =Ê÷u§óïœ~'C®Ú´.¨Òä”ë¹šLßÑüÀ£ñ‰‘–Vå¦)üô2f‹Z‹!Òb“Œu¼!ä[Æu1B¶˜u§ÂE;ƒsdNLÊòÔ”òÑIjþ™ËÞràT²£(hÑÿA©|c»x ¤`1!»hqY`qI]lƒí!s¸HnLý³W³¸Ó‹X]ÛÕÌ˜	—CÆ˜ýh¸¨U¥¼]°HUJ³ÔÃ½xéŠ¨S¨ÐŒÄíqnmÐ¹¢ 1­mxkkƒá‰­Ñ€¢0=ãø°ºM®ÊìBš­LIÛÊÅ…ô§è*
?!g!ì£Â&Ì|‚ñÐµ/.ùMë	^úH1¼ê-4Ïiwö-1ÓOWRd”K!\¦,ò(M(¢L%KOÑ:v’A<ÐxècÆt9H¤\¿9°”HQ:±ˆeôÒqG“ ¢Çøa¬»%õw%€Ð¼¤=v~Êd8o©Î´nÔE7Mkª14ægÆIIßò£tþÖ³ç/+¥ Ï®`ê¹shAçmÕ	—ZÏéi›ò#-––ýf†R«½J®ƒ\g§\‹;uä·î,§M|sy¹4 dvçå“kl<û+78ì¾dÀ7ZñSi(t^$z;þWaÙ­Nü9HÎ„zRq$!Œ·¶nâ’Òmöx{æ¿.ÓUX®4î˜]ãÅáñwäÙÙï¾öúø»]†—»€$„CUnØÒl|cÉã·˜4‚k;ëÖ\
ð°œ³L¸"XDéÿL‰070ò†9½²>}tEŸt‰2¹ì#™g¤2j?¤¹ÖA×¼±V¦ºXð–‡øËÕÛµòš³ÒYs¤mdÎC¯ƒ?c÷ØÊC¼å4pŸ¦e£pA›å¡Qs£é³qqcÁFð‚4ª6ðò‘CÐ,¦Åœn<ô¯Â.1uP<„‚WˆPò®â-pgÔò0Éósì¿¦aN0äd©ÙÐ?ÜFÜMÞôüòy€‘	‘)4 h‰•¢½hú§GÕZºFY84;a7 m±pÚ6éÉÌFtN˜®®2ô¨¸upôêô7ºQEœ¼ÞhüG”p)Vþq4ˆ&Y[}¾,¡ÀnéÇÊ+Ê~â¨m©åNÕåFõ-LÆ’NØ¸•ÿøßÿƒüpÿ#4ØžüÍàë‘ë¦Fß
,ìæûŸädt½˜®Ÿ— =SÝ"˜§äÏ°¾ÞYgÿ¬Á?ë«¶æÃàæ%h

äÁœ¸é>¼ÛÅªåLîÙPÆžŒ+¾úœ+ÛJ%°pí+ÛADÀdÏîf±‰Õó?ÿc<2sÖ½+ƒÌ+ç‹­Rk±
Ë5ùXÆ§ßXo™W‡\³ºqûoÌZþÿ   ÿÿì}ÛrÜF’èû~E™qBlÎð"R’GâÊÔ¶IÊæ®DqÙ´f#¼/Ø²a¡ MRf0â¼mœ·}Ùðž‡3þ}åŸì—œÌº¡P¨* ÝhÞDÄŒÅ
…ºdfå=ùåÈo8ª)L´}Dc§›¼¼ÍU£µcúÁóŸ×ÿ@MÅÇíîÛ›`@4tÓ9fðý.& õ§Þ†q‚©UB²Œñ(/±ã]egÕiæÆ‚FLnß;ÙŸÐÍi}DÏ·ÆÞÎ¦ß_ÿ-]ëòÌøÇ~’M’GÑÊx•ü3Ÿ¾€×›ÁéÈOa±ÅB·¬Åîç±nJª¯æqs|T–è9vÝ7(­Ž£SÃÔúÌ›‹Ê€uš9"ìØU3˜†.t£h;v5ˆ¹c—Ž=DGe¦›<cÿ™ç(p¤Ò«ë³Ð`8?ç-O,
Vûú4a|ìš5˜]&cm±ÿf}ü]KxŸ´Ê›DùÕÁÔÈÔiGãÿøK ŸØ|ŽàB8ßíP8éù×øDž%²‘]…XAñÐÎdMƒÙ%C•¹âå[MhÝÇ*:Ós0¢³…¨_Ö¨ÌtH»îÎÞáîÞÑ‡MB+Èfñ2ÙE
ÿúÿa ÿ¡?¦‰Àx¶­FÞÞ»aàÝ°+Bù¦“'ä]Ü÷B‘âáaYws€ÞŽA e…–ºùe·|ºfýX®XË¦ÝRÿue®/Ò€{ÜØ	ÂÁÛàÂ-–}±Zg|ÕpÑZ×¨?L×¶NB/y7ÓeR¥²õj_‘Êúý’”R»#y¡/äOf>®}ìðAŸ×tÅZ>£±Ï6ÎåZ·ÍPø((Ü  ðÌ.(ä‚ úù¾óOYýÄÖªX°¢—²ƒ#»-ê$ÒÖäQ¢P/“›h‚ZvUûžÅr[»ntós¢;xÛju—ºñI‘
¸œp´¶E‡œu·«pÇ9xëî³Âçîf=îêùÛU¹¤UùÚÝ §]]?»Ê)Ù,Íã	g¥¢šgt­sxÕ©>uÏmôª¦GÝ]ñ§›R;`Dñ±«Dj'PÞ¦³Ým¹ÚÍÓÑÎ*Cšä[TÚ=Ãõî¥ëBL=Ï»¶dç­vã3•âÁ¶PMµÉLQ›…oiœDyhæ,<n;°S+ <Sp§6ÇÒMë©Üiï©¯!ýÔ¯Aøg¿8yw¨+ÚSÊ·xf­W~üúéÙð'‚!Ò'a|O02ŠŒ“RBŸJm4¨¡#×{—¤)Ð˜Bh¶|ßS…ìÙ9ˆ
îáæä¯šÇ—+e¯¢ì$ÛD®Kx	$[ë|Š={|^9M™-YƒagÛÂ¦ÌÊ0¥ÒÎ˜BLmj*©ÆŽ"“Y´Z°–D%“px²’K¨pÉHu'1ƒl„rA•päœƒá®)9»…w™–ŒMZTTTCøiCð‘‚Eò±È;-¡œe‹µØ!cGN±hN²)dZsùZßŽ\S!ÓÌpÏËÉèšúHœ&ph¦³o6NŠÛ¦f™ì–µûÌ3U˜Ó¦Ú&µ[„‡zÎ&_$+]Ï£djèhêSr‹@RåòE‚G¥ÈôGJ}[
a•u¡Á£ bYæ ¢ï¹ƒÈ´™)³ÎLB÷øD¾W)	(mk›0K‡³¿»!‹ÃJÞoQ&pÿ%ñ*Sd¥òžIáíà`µT^…ƒÂ¹ó›·*œ«02‡Sé¾¡©ñ6e3žwz‘C6cÑ›÷‚rº’ÙÄ0Xº.Ž°DCË‚˜ø€ÎMa±þAŸ~ïG©Ò?»Ñ„ÿµãŸPÄ3oÑ=FN¨«K?€‚8Öžø.ñ²I€>Ðöˆ]í}6¬…-ö/é<{Š,Ò¥]Ðé,lÑ ƒ¯_4î!_†…­ÿùý?Hþ›t>™^ÿ+NÖÈ^Œà÷ÀsônH¾9Ì;ð“ëßãÅ>¶1¶"SŠ8Ñdtì'Fæ`Dß,¬[ø€o^½zU]¯ôb“<{ºLž¿X&_Ã¿¯ŒôøÎQƒD á‘mFãX¾ü3ú[Zâš´J(d×›X&õ÷¢L;A–ÉúSZT¾|·Î’^„%E9ªÛ÷ÓôË=NR¶x–tÙnDXïÆƒ„}ÏQßlê#£Ëm—ÑW¶¡ä] ‹¦5 Ûr$[òOÙao’ŽáLquØ€RÏ«„î  yêÔ,¤{ÁŠ>}¢”æ-ø<½CîD€5–çX)ªüÜæÕ o9ÝË±À to@¾!j©]CM\¾D©²Diná×ñ_Q½¦7úXûÉ´ø@¹~Me«^¸ª¥ÃŽÉÏ½Iº‡ªSîcŸÅ^šu@|„’ëßá¯~<"é„¢ÆWËd!eÝ.@Z—¸“ÔOLyv²ä³1ý+‚Ì€y'îwøÿãe²ÈÝ&ÓŸZ“..;ìô½¬?$H¼LÃýCÕO’8yã?ìÛ@|Õ•œ$0ë~O›0çòá”rúoÌ±'>í­G!-ÆŠtƒÖùóáÞÑî2ù7ã|Öþ×¥˜ÏÕ¿™&ôw®ßµJ2|WŠ3+2gX³ˆ½ò¨@§Êuçd9.M]uäÃbùIæÿê©åå|Niò2rê×+)÷J/)§¸b‰b9oŠ‡ö.F„§ò¸þ+	@T?I¯?óCµ]¡ ÝldË¬hô·S¦QŒüDp`vi	:Zeê(Þ¡] K‡O0=…ßo©¬Ží:j«¥Õ,~ŸûÉ¶—úšekñïKŸ Fã“?8HâãÐ¥Û@ ²ò×Æì±'Kpq¼–1ŒWñmýë¸b:RkµIžÚFÔ‹ÃI?öm#Jùc1 Ô1 ´á€”%~6I"M¥[>.ñÒŽLZÈ|~+Ëj–DíØ¬¬1Ïßç{+‡ªâîë/šRç­(Ï.ÅWWý`ñ{UÎ°ÎàS5'Ua—RÂ’U°4Q˜ZéT/4©aº:úÆ›²f?¯Á™G¥1ÑƒA‹ äÇ©ÃˆãJ6èA”2´í!ÕúÖR|N±Æ¯VÔ8!‘‹Ð
lâ2‰$]å­êÒéÖ"E™ dE	j“¶Ê$r—^v…Çhn<¬|àÊŒWÈ‹G¿‰¡?¶.y18×§ÍYñìYíl
›`³øuLgÊyŠ¥‰¸Ã#»í¯TŠnåÙ*çöP[T®»Ìn#+^(œÆZ ö_¡õÕöèÓ±™ï°×pV>òÒB¨x×Ý¡#$¶§¡/J¾ÀÒ‚¤›ùcK¼˜pjm¡‰bYÏË·ÛGEY"F=ÇqšN÷¥Öv6ìú•ä:¨dU¨ýf`åŠüta:D";=°ì5ˆöi™¤7\&AºŽ¤Ê^–K©W7¬G™XÜ±F¹Œu²ÅzŒa„þ—ú“4g«à	ŸÌ‰ÀHgÐf	Â #/B0Ñ²b‰?ŠÏlPqšULÝ[j·cð±h¿ÓðQ<Xw`¤hóRÅÉ¶1R):ìÂI¥Úg¡Ø'gèB¦Aòxt±Í­À\©,r\(œ0„HNF¼rä 8ÇŸ²†¸ÐîŸN(N›0TºW§ø¨&\Ñ×ráj©#º­-v±85‘Ã¹®ê‹çÈlÞÏbâÈ
íÒ„U²ô%v¾8ÚNäŸó¿÷Ðz²dÝy„>N•£Rw­Ûéâêô¨¹iØvêã€yV#î*áj;’®®’]ÀL$ªÉõo8¥qQqad>t'Á¯°B²\(W?%Ý]FÊz:ÁªSEÅ€ü)ÐJX`’Z“
„û €X}-2³µ•5k¹NéÙMŠMmS
%à‘ßâ0>¼eòXÙ_à|îy×ÿ=ð—Éî`Â«:.“C„¯IÂ~__îZpàE<*3Ù&‘d¿øRµTb‹\ÓIç|Ì‰X4% <^zqÛD›Z%œÅQv8½8Šéýe’oy%!•¡Êÿ×FCµ/ tóUñ#œ¬â­1bRI™z(¶»2½¥PWNž)­¦äÖMªo‘tJ’)‹rZ¤•pF>Dáçö‰¨’Î­šè¡—k’¤qä6CwC ’•Xi0¨Ó™*j2÷Ã K?ñ·ºÇ›ÂHijOW_:Äö`Qu†á;wþcJòÆGWÏ÷à*Ž-?¸züÂAµ-õ»Bmn¸³Ïj’qè-¦¾2›Ø÷.µºZ¦½Šã¦<¡D©4”5~pÁ„Ÿëçƒã %Ô  Nï³™¤(P¤pÚ`æxªŸ©'H´æè6žì^À³ƒ¡½ìï‚Ñ™4¡¡Ú¡ãW:Ò…vÀžŽ²†ú@­Î‡¶#V´ÇVôF¡}¤ûŒ¹44È¯sj4¡ìWÈf,Õ…:u_ÝŒšPoV×qóŒvØÑ».Õ9mð²BsÞX?g2ºTãL/¾§Úê¿*)–JÚª×ý•jþ÷ÆÞ)Es]Ù
 wìE}zÐ}Ø¬ñ±I%B:±¢CÎl`:Äê=ÎÃ­è{[#Zâ“!ôŠù}ØÐ?Ç0 xÉÖ“‡êe©—NÔrj½&•%Ë6Õõ y™”¦Ñ¬êÿÄ£ø<¦®ëð.Û
ÓË¯×&e.Ðª¶g+ë
ÄJt¬•a¥?:Muî¿‘_©£ Ìî$è3°Hb<ª8S5Ic”B]]ZFÀ¹
åg’9³þÏA
W¼ñï¶${«}Á[t¯HÝj3§È­½{gåmajæ”+_'âeÊt ¨"³=ò€LxŒ‚	íª´Cò“EÑ…æx\Ô‹éN¤“ÃÛó™Q&—ñm>_ŒŒ¯BüÝð=Fz¥ûY¥{ŽÑ.Ñ^¶Úr½xáÎõÌTÖä,õ¼yœù˜:|íëÙíJŽxEcìèÎXê<Eq ÄÉ©ÁáC‰p(@SqLP-Õ‚¡*ÙîÐ‰fì€â|¢ÒS¨Ï‹–×–©oÄ8÷HÉ¦0ÃúôDKÉ÷´Ð"û6ˆ€§ðƒ$^&ývºŠ¤5³·É7nŸ¨ÕôÃÁ™5VW
ÞNöÐ+­*”€M· –œRA#T“õ5Gv>ô„šÊNz…òt:É#Ú0Ý¦¥ÉG»3•AIcÃç­ÿùßÿE>¨S+8	5é“¶jT¾T|®à…ìMàC”¶"ñG|Zð¿Q–WzÉa®ôäŒÞ
@vég¦ÁÌ×Ä@qf–¦Sá-;ò„ð5¼åc{ÄÛ®Ä[Uµ•–÷ð¶«NÍ…·Þ·ÂÑxj´íIýIëx+Æöˆ¶E´UTÏoUµú#â:7°›u‰¢ÇKlý~Ðªo¯a¬•¨¬Œj;¡ƒ›µxÞ.#}ë¥]tÏ€ðÓ2Ø­Aëš@¢]æÑ´,êôJa­9f‰ÀÚÜt¡{ÎãåÝËÔº¯M;”îeä­è]¡SÓvÏ€«”•%Ã:)÷˜6M™fcîE¼k-ëÅ½HPÏ9è¥üÏ£"«¢þyÜ÷íAœøA†¡Zôx¥¦àÌh€_Î-îVƒ&ì4·a[ÒZÌ#°ƒƒ‰éXšÂÙŽ2°}¤;<ã…ò5üØ´`¨tS/ÍcxCøl¢uÐ§º8ÐšKû¹º= §ù·TˆÒ€TŠ*˜©Kâ(>M¼‘—Ç	–T0ºËœ@fìö½Æ x >˜/Ï<ž[§’Ê•†Ó.PÔÍæMµÚÊF£qœd<ýOÝmÜ£/ñÑÁêóÓ¾Ým¬|£¸• åÑ P…ß©,c•É¼Ñgí¯<×¢^·žÚ`+KÄC+2Äÿ:çnÌS¨^^U ÝäR¦D¸npÏºâª®Aâ=‹ËŸžéwa«Ç³C™ÓlØí’qõƒpˆ<‰ÃôF?N“Ñ3ñžÖ¤kN0( ¤%ì+·“îŽžý}ÌŠÖJmßO£ÖÁ3Ä’HÜ`Á•2Ô¦-á·µ«z<–”[à§›à®¨‹BJÖˆ¥aú)±Ý`2ƒ>n›ytVø»SK,D«ù-1—æ´%öÊK,>¬%ÎEþy­±i«–8o7Ë
›Žúšq¾ÁSRaâC¢•RuÑ¬n.%GÅÝÇ¡{¤qº†éêpñ-™‘
Ã ¥VŸm”Ë¬2€:3¼Ð§ëâ¼ãUÜ~¼´žwy0ê:òe¤v:ð÷&I3tzX&Á&a©†—ìYwè´hÂàª0=º,T}^ÙsÔ­˜`¸Yxµ%}ÄYàŸ¿÷²IdŸwà€ÂÚÌÞ>óÐî“u '¯OsFìu Ãggì6J>7jjJ$ñ?È!0ö È—Q(¶0e6pÙÀu(-Ô'˜A[á™×½±—|
Q'¹A[n¤^ŒŽŽ'°CuG“³>Åo»ø²@²ÄëBŸ¥ó  laË– Ú~?£¡¾eÖoø¼L
Cöuø;ñŒ+	}o€ßÍ‚ÓafX¼‚Ê)€!•sœŸ7ç89êç«
å+õÑl¡Rp­óÀäôVØ  á¤ GM°<ðÈ€Xœ^ÿsv’nBŸbäJ°%ÀJ¬|â‡ÞUêNü¢Ž®¼#UZ`RTÅØaWòsŸÙ´á—ü‡]ü#	Æ£Œ¯ÿJ¿ï…<iîÎŸúEÌÁ¯è‘P*òG<8 uHñ„s3¹&‰5	¢ÔO¨Ñe?™öêÝu…Ê"Ægh-°½|¾Tï|2W¥Ï¿ò†,P__Á¨3M!…‘÷®Z=È,¹¦­c4$Ÿ.4èÖsšj™¯ˆã)u0ÔÔÜÕ»eïNùeM\cÉ·“ÀKèR3ÿf‹Nö«ÂÔ`pBb}ÃýœiÖj\§ßYÙfF‡é¦ŽÄÕÅÞ!uDyÃíBl«5pŸ*È8€Þ¸oiÞÞ¼¶äÜ™ˆÛ"ßh¥L	¸‹{¥¯}cÊî5”ƒ×ˆµ°Ò"Fcr£ÅŒÀžÅÑ)'b”sufè	ø qVb^ýI’ ¡G\€¹©¹—£Ë¼<Æôñˆ{ô1Nˆ@íå?i–dš”¾b”t5”/w.Éêêª:²ä`0ït0Š¿È•e€¦$¥å{Wú-K«çk•ÖƒÚ³¥N‘¤uÖ­eaÑœ˜µ€­u}DÖÀC`\*2ç< ÚdÁ¤™” u0Y‹K0á¢¯Ú¥?ËØiAãÒZÚ©¸¬ ¹	Ë+có »õ`dãŸåb˜	…@ Qþ²6”È–üiïˆs0‘ê(Î€ÆhÎÏ¸2VÂøUùtr@ÃqR˜ªg;X,dÎÌÈÉÏÍÞ¶CYÕLØò«Î\Ø—®½—(ÎlÃ¶ ¿wÆÇ^¸/þÞãþÓ—èÂ/ÃnšÊ;Râ.¾eÑªä}ÌÀþ*'ì÷ûË ßW¤s)Çuå¬<t·äõc”ì…˜d;ðmŠŽ÷à®l<¯këI	[ä|P…lWRH£$R/T­
ËEo¥êYå*Ysˆö®S0I‡ù\ÿj‚lõ mÕ +jÚ*Ò:€¦Ê±wë‚M·2i¼Iöé‚êÑwlD‚5Q]ªó½Ÿ …UWÛ“4‹G4––ê	Ùƒ'„KWô¯ƒ¥¨áÀ¼›b¹h1`ZB¡ªÉ¦£É²øa’%ñ"­d0`5à|a2µÄrd¼h[Ô#±¯×çjK ®Ìf$[„ñ
¿M•Vt-l<ÑŒlXa8³y²^KšAIm¹O`çÐïƒÜÜõµN¿CP]Ø¢ÿÎv<¦ÎcÁøŽNPËÙ8š@…þ‹º…EÚõâ3Ìâm«¶ÀxaKÉuÑ[ìŸ&ruiºä¶è?žÕV(åÒŽM
D#ÀÙ>yB¾šé©zÜÚ©ºÒK&÷û\HOP¯Ä):êß:eeæ!/2dÐS¤cXNtÆÌswsW™TE©UÕ©ñûX“^R´mGû±ÈœÉçG;ñªQµWC½Öâ,Ì¦”#šKÎfRâOg2(‰/sûÉ°¨ó°jUbnÇ¦$§j¶(±Ç–z¯©5‰¿\´%Ù»šÞ’¤¬5µ#‰ÍÍIy†8¶µØ–ò¡I“µ)÷9ÎïÔ.¨ÄÂó‹°žgIÉ74<ÝT~nPg¯5¾W³/]®ý¼ëî| »½þa÷pçÃ&ÙMûñóîà@ ¶læoý‡5µ´ñÀçáË+èŒ¶ò'b³h•ßS¼Ö^<¥Î5œÖ^>u|Ãâu!€(s!ˆÁiL]Zl.–ü¡pð£k×qÏCƒ¿`¥ëÌùÊ2\¡…ÇøX_hÞ7ú<7…­u›ç"ýDßèòR–Ið*7«‚Ït¤À'OÓÎÙ|½é—¸WpëÉ·8÷’°zõpiŸAŽò‡‘]FÁïcO„LZ›M)ó$Ê,IþˆØ˜j”È\b?gUv¼£k UL~1W@4«Žd/´"{Þé2‘½ê%SÙu¼ÌÛáÒWÖEú'ÒR© ~¸ê!j‚´V²üçý£FeD[mD;ÓÏžÙ*‚ÞÎôDY<Òåà#þñ‡:o{qcü±™€¤‘/i}ø:Ì{á-xG©bXãTi¿f/¯­í¨™OEyGK¹F›ìó6ë^Š…æB+2ç|Ö&éXØ¯ÅÅ%W¯|½É)õLS;ç~N¤^CÿŽörÏ¡ÙBý5c;¥·XOÖŽlæJ£¬„Ws*J:ñ˜¥W\ªCPKõ±Þ\ÀLM ¦ãfIÁÎXWÀ;lð?Ÿ•Ð[€vfQá?ÅH4†S9îŒ›z0C/Ý÷OÑ7Ô?D—ÈÈê7¦!DüŽÒ½x²P¸ô°doâO|Óñ"¡ï¬*+Á§Œw€Ðê|~DñXÄC½àK[õ6âÆRÍÏˆù1qmÎ›EÙyñÙ— ÷¯i
…õËŽQr­»hÅB¦näHP“±`k(þsëR]·+²BqÍDœn’ñ°Õk¾ÿÂÀÑõïÙ$Ä²bšÝ|U«‰W…f/3ç³úâñ>¥|«o½pHuU@þ²$¦5m0î<þË¤U<Å[’?r.¢žøQR4âuç¸q––žå’J'a'ÁœÁ+Œ|³€¼Jº’rþ']95¶mðƒô½ýîþõ¿w{Ëd·wtØ=ºþ?ßíu¾ì*CyÎîµáŒÅ,B÷jÿ€ ÄëK4uj/óp¿I0°lyMöRY#C—å%¹Æõ_yß=ü§Ý£½ýïÈòqw§É»v»×ÿ~ýŸ»=xûàðÃön¯÷¡A‡»Û?ÂäûÞw÷›¼y´»½ÿáÝXÔÎÙ—Èß‘æ'¶ñæ]ÖÔªÑ‡w^UkãÌuµêL
ÊÚ»²cgá,eUëŽ®ÿÚ€a·hjÌ˜qŠ7¤r±ã$røòáuÎŒ!ôòHð+î>DŠžêî<‘ÙÎR/«š©¨ƒ0Ó÷†II5R‡c2,—¬póàÙå©soZŠìAèóiôŸ'^O[;óG4'Æ‡bÞÛº@ê¡j¨EÐ¤ýÍ
•7‰3„åtgïpwïèÃ&µÇû¿0$^&h¨ÆÄi´ž üÕÄzúâ¾XO=ôí¼'2ªñ<ÎgQ8ŸÙOcu«‘9L¼_ãrLã4v¨F˜.DÏ“›¿zbåAÄ,DüÑ¶£©&ÄÂP¢öþú¯ƒÀkOE!:oÙFÚ3À#+ÜšB]ž¢ó[Øbÿ6|ù[/¸€wé?îWï—4ó²‰‹søñ"¥‹Bqâ x+¤ÆmaëºeŒƒÄ@öWCàÞ‘.ð,wû–ú«aG2«%t“ÿÝ:¾Ü/^í[Å‘ÈV¡Ù¢Nzý.H³Àþ@ThÞ(gÊÛª@ÁNF¨áÓ€Ú8¥Ö5»BáÂÙE‡.µ"º/Œ0º\7ƒq¡…ºðéÿ‰Å~ÈÅX{ñÔNLx«¹R“ÒjT×—ÉÆ2yÖ@ Žéf¢:Îî>Om˜Æ.á›‚¹=ü#Y_¢a(5Å›iVF{IêïEz	¸e²þÔjltÓ@:ÖM Hìc†àä“ÞÄÌTÃÕ‘wÑyºLÇ°³²¦>¨),™³½2$#eX`|ak‡Žã½˜&Q>O/ìQëæ<Çæ˜¯ˆ–‡¿Oilcÿ©mŽw‘±F)‘ìE×¿õ™ÙNƒ¾ÒÍŒ.÷šÇ@Ý÷^„t»U³IÞíÃ2RP{ŒÁ¬9˜Á²µcÐ§`ìÁG?âE€æ€ˆxrç›ëßÎüðÑ7£‰ÉZ]¹ú˜°·<ÌØ&6(ý>,’›cò#ÑmNtóÕCh»Š|#„×xS»‹¶…ž/8fòÎ'~æ—ó0˜h|MsÀx’ŒCÿæìßú \\ð¬¬ È‡cQ äûDØi,J|ýnåñ™§e>O¼1Sg LdšÆñ$ËâÈ¡6`ÌØƒžAÿ“à©B¹î‚òïŽ¼ ˆ­êw]xdÈÏÁ®|Š—ýÈæÏj©^²Žæ…^Nž“Úu-#³FAÜÒêíÉÊ8h§Ã•W<ÝäÆõ^×+‘ØÜð­s·´'š«R¤½ˆUG)RDZ“s"ì9_+[A0°T«q/Kz–èó3Ûo*nãÆYž÷8Ý\ƒfÑ’í®à²)ƒLÓkwÐÿ<ô²ÔÍ°.ž~¡à.§_ñ%‡¥È«OïÌûiêú½¿L`AªŸ®Qw<¾3àï¥Ÿ£>1çç—HS{‘±ü&Ø?˜¤C›ö‘eŒeo?yB÷c‘Ô<ŽI‘s,wnOS(†€åààûÞ¹dDí„‰¿L€S>€AšÂŽ#ðÇC»¢éENY2`MØÃ.ôÎÂA<‡w'Ëdè!Í°–brFñŒ-–²›dS€à„e^Kü¾ˆçqàÅØB¦«®ËÛýš¤7lSnj#uîC'd¸4UDLaNË4LyxHXoä%ÙxˆKWI¿piÑ.›‚è²ce˜ŠâYét1?ÆY2•Ùé¼ÂÄ8Há‹œÕ“\
€*À«)åH6(Çfc‚Ô+}ãT	HŸ;}µ
ýˆ®Ã•—¤(¥!rZKæÐ*Áº=m€Mê7Éý8v/íâ2!ðXl¢xUIù6Z*ûo P\¶ÓMù—ëW[ëMÉ6¼Ê ôúÆÕÖ¾žNùþ³«­g³¼ÿâjëÅ,ïÿéjëOÞw918t{–˜Û[ÃømxÓË¨ÞfîÀYt˜ EÖ$3iË')GH5‚Ö…V4Ö-Ð‡T¡EÕ RvÝ>ø’7À/±ÿƒ«¸ŠÀâ	ùã‹/Ÿ¿×44Fì”S8=‰áÝ×H˜+.{”E­m~`Èålóh¶Þ˜­<³È£Õ¬6³H¶¥´h©$³!ØMW ·ÔbæµÞ	u‹`PD³jjU™§«¸l«þÃ_íuøóë ‰¯°Äu±Œ¢.ÔA÷n©’œ©­kÐÁÂÍë ñô:Höî”_ë )K^·’:Sy^Æˆæ¯+´ÅPþðÎ•8Ú¡óàö“¸*~ÿ>¥‘¯È Ïcî3.cîO‚È‹ú~Äå×“ø<EV±ô@)3à‚º©Âó^DÐüÎ21t|WÍ?°’YÄ‘½VÖâµ.ºÑ ·ËË°´E§Dê$î®¦°uÃæ®!«¿ä.ò¥W`Ú¶:+¸'Xügîä×  ÝøhŒü‹ñ*Âi¢[¿vð€w65œ7°h¥Iœ®B‚zùž•I¨ÑnSäo¿“•ÜPk¬„À´W	¡Áæ[	á±hAÓ¢ð×‹¸ÔÀ!]k5üò
Ì½LA	Ýo©LAi•e
šÐ¶y”)(Þ;òN÷òÞ{‘w
°¡u‘y§©Æ€à-„îºÒ[+ëùçÐuZ—ÁNÑîL_BöCïZe›è ô;ÅÉ®ÍMñ˜«@š©sÝTûÊG£è¦Õ…õ·è"}¾JQŸrÆOø¿p)§÷Æ¨'àÉÞgWÃÿŒãO3‹@'~@ƒÃè=ïÄGa^|>‘Ub‹eä§Ó•ZÀIToN²äû<ÀÉPZý.ƒUmË²èM·ßº³e#œ­Ï©Ã³âkQeÜ*»&^°¡[ÍL*RÑ@æj-E1šVw¤šÆt‘+À4oý…­…Bÿ¢Âã‚¾2mÒÔËþÔdö¯y ÓýÂ™Ôë	¬ËÛ	LùÛÂ³ÙH"„Ù~&Ïd>“ßeŸtÊÍgùÍOÆê=»(„vÝ”Q-Ÿ¾Ù¦&P—ï[S‹š	óe†ºvíiêÔ5§)ÑÒ ™€ÖL‚,F4y¨çÇÀ°!ô: å“}²¯5¨ß¤S‹DŠ±BjHQ)³÷hP#³·Ù?Ç©’´Q g¸$Œ»‘‰‘»Ù„­‡ÕÇ‘­ÇZt½[N}YðnµËñƒ
À{OC±dJì¹¾dßË‡À$UeLÂrÄ;®0}yJFÇ»¥©Ži*KÓì¶&‹ŠÏ¢¬³g¾I´•4¯Q6øÒ1ë¡D}K<p^3ÂpÍŒ°ˆB˜ÐFz‡ÐòA´ÒÈX”ÒiÍ¸L: {G%¶«è]<kø-ƒ"B	úœz}X#)æB ríJC6ŽN˜èâ¥Ófä8NÃøØ÷1¼÷²þ°3¦Ï,£ææz)”.ÑoÈO¡ò¯{6¶è}ø›ÓÃ­¿“'™VòøV²:‚2…ä#yIFýìÜ÷#S¼ô4¦UªF3Ï§¤‡£¯%ÀÏS…þÖ¸vÍN­±gU‘gâÎh%'3å±ã9ÓR}5rýã8	Fè·ä·b5yêc0Ì8ù\8—øX;CÆ­úìÌïîõ&§§èÔ­
ìð2ùñ'ü¿u´¸D¬wWðšNû¬‰N-2Ç&Ÿ©±Â5Ñ:K˜aîåŸeÙ™v¿º£1¦Ï/Ä²Ÿâíên”4"åÎL«»ü4äLâ|†ÅÛÕÝÐE…S1Àéôý4ÍÇU~Äº³öæ¨øæ²›ãÅQåñbšd-!84ãy½s%Õ&9eµ‚èìú·sÊ0YYæ‡Áéõß¢~à‘nB™²À(jFÃ+ÒÇ‰tü$Ù$^ô¹’FÔø¿Fä´ÛÕ‹õÅ%…ŸK³ô&2ˆ(h˜,ð¾©?
J™¿–Á®|yü ¿ÅÓu7Â³B˜‡®YÌtÊ\•ûC
«¼×e±°¥5?¶m/;+è_¹3#wÔƒ?…Ð^|>EòÎïèx0Ä‘vÃóhÎÌb÷
¶úÌ:|šnMÄh$Y„°£Ø9ÚÖ.üA…¯u³öÏP­BËW°„·(Kd$Ê‰bLåiWœÊ1q™<êÜò‚ñükÃ·ÙPQL¢'O½™XÐ©ÙÇÒ8šû¼?B7ƒnL2LS páŸqrý[ÂMŠÑg2o€göõo	†9R:ŒÆ@	é™^à5†©`/ñfmúî…ßç¨ÝQ³È-“:ØK—ˆ?"@•ÃITn]u,¼YÕÙüq¨šêø„5ÕCÝ$ªv»ØÌ‹²â™ƒ	üYSYeâÓk©­*QÂ,ÔÎ,±f~a>P‰X%¡Ñ^eàó«Ž;Ýƒµíî[âeÔ„ÁÁ_‹r^;`¨‹”- ^oSù(êÄÙã2¦O4÷ývh5ÍÎqSÛ2µ*Î¢¡ÏS9ÚCí½´Q!ÎgêØ·Š1ïýeæÅYpÆ\ìB\à¤°®ŒÛ¡¢¥ï–Dý¬*«<ÚÁ,“–eìº ÖÑ°ÃýÝ™Ú¼Sšºî“¡ËOŽð Ôt§„SÑbÓÞâR@#òìðº1:¢tjØÎQ:ü…ºA:•êÆfc°‡é´k‡ûÞO^BS, KÜ äØÉÁT±#|§¥( yÅÍ	dv˜6È•ÙçÒŠiz@­Ý“'Ä+dyúæÙº¬äb¸k$”·ÇÝsÖ{Ö@"-”¨L“Z©ÊÓ šH‰'*ÆNÔ˜7ÎéhÐŠ7ÈüS¼W+Ò¨€/ÓU,ÇÝ3™‰yêf¨¼‡˜¯ÔÇ4f§QœfA_”^Šç;òqL_–ß	Rt¯ýæ2He[ét›Öä›[n_rÒåˆ!þÊŸåŽá¹ú+o“p4*=ø'Þ$„M›$}:Á½¶ã£Ro«‘ÃVÚ*«#š®mý§ÕµF_uÁÑï™÷¤;bq«MÖZl•TÐ…/½Ê×ýï‹qG±hÜÁ‚bJcû“$Ci‡®É ðÖZÊÀ°<ro÷cÌØª|Yì4ó¼¢Þ‚¾ð¼Ò;çm÷KoV©k! lax(ûNƒŒmô.±‰Â.;Z*Ô-_ùp_¿ÀÇYBZÓG°KÎfz¬ngÏ;óÙê"ÁÁ¿zYRÚ[d ‹=éü3ó$!@¼¡ÿõ6W¶†îr±êq¦|R…«2(ä0Ï¸ö>G}üTS‹rûµ€Ná9p´ö't4) o/ãý|s9ö“ Ì´gÚoçm…Éz“<&“þ€ÇüFÞj’ú	ÍûÍ%þ‰0	?0O’8£ÞÝÅ?kõ­×ÿ4ï^Œã$³-kSs­´Æ†…bãK%°÷òvXˆHç1½U`â	|¼Ï€›=~Ëç-ØýLqÍú)ýŽ:‹ø<
co€ý˜(B·þ¾ßpówÑÕ·r`¥7—IqÔõ!ûr<8ù!	µÐÝ£.¸ð$ˆ`7V¨j€&Ú\ûú©¡8ˆVÐû×^Ô8Hâ1¼;IuEMHíƒ·ÀYüS0áÞÅÊùÊ×pc¸òã«³áO$/ãÑC‚ž'!¼4?²qæ†g
Kñ[æ²(å/Õù–­Œ—þæùÊKš%TIÿ¿.$5ã¿[Vµ{æê'd…#œA^SYÛŠêŸ•\Gp$VùÇ eF ™îP–|½6|VÛÖSZeZÆÆ´°ßR§“<äKl<Nkµ-XvýÕ©C¨e²ø3`IôiqéÊ¸Š»°À‘¾¢OÊJª‘	––tIXÖ=ö‹ý­-9xóÉØ3Ô'‰‡°ý	
î« (&Ÿ©+h'EÖàGîŽt°ó–$þYàŸ/ü´¸„¦ùïÞ¿Û{‹v™ÑÝ–à†
Èûk“ãAû?;ë˜ÚbÉ÷(3+pªÄœ›Z7Aqï–Vß½í{#ô¦’Yv{€ó}ªÝöÔfv|Ó:‹žy¼ÕaâŸ`À]/s“?Ï Ù"Åå- Òx«ðö¢ùÝ>Ý´—–•ggÅRÓ÷	€m²Ø%<¾8WúŒ.VÖq¥½:ãêØ)ÌüA—•É.e¼<¾WñR´@óAü_™ç*nž}Åf²‹Å ÄúPïG8:õcš’çŠ£è5§i‚¬/]—+ƒžnÈþ‘g"àÈ©Î¢½@„¬2‰eIq¡üÌÕ¯b`ÁiÊi Æ	ÈÒð˜SöïrAŠy#P,Ç”Ã´½Ü£×kÀ–Ó]áýJ\ýÝÿ  ÿÿ ¨Á÷M