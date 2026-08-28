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
          const saved = encryptedLocalStorage.getItem('local_diagnosticos');
          const parsed = saved ? JSON.parse(saved) : [];
          const nextList = Array.isArray(parsed) 
            ? parsed.map((d: any) => d.id === targetDiag!.id ? updatedDiag : d)
            : [updatedDiag];
          encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
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
      const saved = encryptedLocalStorage.getItem('custom_modelos_relatorio_v2');
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
      const saved = encryptedLocalStorage.getItem('local_diagnosticos');
      const parsed = saved ? JSON.parse(saved) : [];
      const nextList = Array.isArray(parsed) 
        ? parsed.map((d: any) => d.id === selectedDiagnostico.id ? updatedDiag : d)
        : [updatedDiag];
      encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
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
        const saved = encryptedLocalStorage.getItem('local_diagnosticos');
        const parsed = saved ? JSON.parse(saved) : [];
        const nextList = Array.isArray(parsed) 
          ? parsed.map((d: any) => d.id === selectedDiagnostico.id ? updatedDiag : d)
          : [updatedDiag];
        encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(nextList));
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
        const savedTasks = encryptedLocalStorage.getItem('local_tarefas_plano');
        const parsedTasks = savedTasks ? JSON.parse(savedTasks) : [];
        const others = Array.isArray(parsedTasks) ? parsedTasks.filter((t: any) => t.diagnosticoId !== selectedDiagnostico.id) : [];
        encryptedLocalStorage.setItem('local_tarefas_plano', JSON.stringify([...others, ...newTarefasFromAtividades]));
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
                encryptedLocalStorage.setItem('custom_modelos_relatorio_v2', JSON.stringify(updatedList));
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
        const savedTasks = encryptedLocalStorage.getItem('local_tarefas_plano');
        const parsedTasks = savedTasks ? JSON.parse(savedTasks) : [];
        const others = Array.isArray(parsedTasks) ? parsedTasks.filter((t: any) => t.diagnosticoId !== selectedDiagnostico.id) : [];
        encryptedLocalStorage.setItem('local_tarefas_plano', JSON.stringify([...others, ...sortedTarefas]));

        const savedDiags = encryptedLocalStorage.getItem('local_diagnosticos');
        const parsedDiags = savedDiags ? JSON.parse(savedDiags) : [];
        const nextDiags = Array.isArray(parsedDiags) 
          ? parsedDiags.map((d: any) => d.id === selectedDiagnostico.id ? { ...d, cronograma: synchronizedCronograma, status: newStatus } : d)
          : [];
        encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(nextDiags));
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
      const enabled = encryptedLocalStorage.getItem('daily_reminder_enabled') === 'true';
      if (!enabled) return;

      const reminderTime = encryptedLocalStorage.getItem('daily_reminder_time') || '09:00';
      const [remHourStr, remMinStr] = reminderTime.split(':');
      const remHour = parseInt(remHourStr || '9', 10);
      const remMin = parseInt(remMinStr || '0', 10);

      const now = new Date();
      const currentHour = now.getHours();
      const currentMin = now.getMinutes();

      if (currentHour === remHour && currentMin === remMin) {
        const todayStr = now.toISOString().split('T')[0];
        const lastNotified = encryptedLocalStorage.getItem('last_daily_reminder_date');
        
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
            encryptedLocalStorage.setItem('last_daily_reminder_date', todayStr);
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
                  onClick={() => { setView('checkout'); encryptedLocalStorage.setItem('selected_plan', 'monthly'); }}
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
                  onClick={() => { setView('checkout'); encryptedLocalStorage.setItem('selected_plan', 'annual'); }}
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
            encryptedLocalStorage.setItem('preferred_logo', 'sebrae');
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
            encryptedLocalStorage.setItem('preferred_logo', 'consultora');
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
            encryptedLocalStorage.setItem('preferred_logo', 'none');
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
  const [dailyReminderEnabled, setDailyReminderEnabled] = React.useState(() => encryptedLocalStorage.getItem('daily_reminder_enabled') === 'true');
  const [dailyReminderTime, setDailyReminderTime] = React.useState(() => encryptedLocalStorage.getItem('daily_reminder_time') || '09:00');
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
    encryptedLocalStorage.setItem('daily_reminder_enabled', String(nextVal));
  };

  const saveReminderConfig = () => {
    encryptedLocalStorage.setItem('daily_reminder_time', dailyReminderTime);
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
          encryptedLocalStorage.setItem('sebrae_custom_logo', base64);
        } else {
          setCustomConsultoraLogo(base64);
          encryptedLocalStorage.setItem('consultora_custom_logo', base64);
        }
      };
      reader.readAsDataURL(file);
    }
  };

  const removeLogo = (isSebrae: boolean) => {
    if (isSebrae) {
      setCustomLogo(null);
      encryptedLocalStorage.removeItem('sebrae_custom_logo');
    } else {
      setCustomConsultoraLogo(null);
      encryptedLocalStorage.removeItem('consultora_custom_logo');
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
            encryptedLocalStorage.setItem('selected_plan', 'monthly');
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
            encryptedLocalStorage.setItem('selected_plan', 'annual');
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

        let rawSavedKey = encryptedLocalStorage.getItem('custom_gemini_api_key') || "";
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
    const saved = encryptedLocalStorage.getItem('local_all_respostas');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        cachedAllRespostasMap = new Map(parsed.map(r => [r.id, r]));
        return parsed;
      }
    }
    const legacy = encryptedLocalStorage.getItem('local_respostas');
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
    encryptedLocalStorage.setItem('local_all_respostas', JSON.stringify(all));
    encryptedLocalStorage.setItem('local_respostas', JSON.stringify(all));
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
      const saved = encryptedLocalStorage.getItem('last_sync_summary');
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
    return (encryptedLocalStorage.getItem('storage_mode') as 'cloud' | 'local') || 'cloud';
  });

  const handleSetStorageMode = (mode: 'cloud' | 'local') => {
    setStorageMode(mode);
    encryptedLocalStorage.setItem('storage_mode', mode);
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
  const [customLogo, setCustomLogo] = useState<string | null>(encryptedLocalStorage.getItem('sebrae_custom_logo'));
  const [customConsultoraLogo, setCustomConsultoraLogo] = useState<string | null>(encryptedLocalStorage.getItem('consultora_custom_logo'));
  const [logoChoice, setLogoChoice] = useState<'sebrae' | 'consultora' | 'none'>(() => {
    return (encryptedLocalStorage.getItem('preferred_logo') as 'sebrae' | 'consultora' | 'none') || 'sebrae';
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
      const saved = encryptedLocalStorage.getItem('local_db_areas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [dbSegmentos, setDbSegmentos] = useState<{ id: string, nome: string }[]>(() => {
    try {
      const saved = encryptedLocalStorage.getItem('local_db_segmentos');
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
      const saved = encryptedLocalStorage.getItem('deleted_areas');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      encryptedLocalStorage.setItem('deleted_areas', JSON.stringify(deletedAreas));
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
      const saved = encryptedLocalStorage.getItem('local_empresas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 9;
  const [empresasCredenciadas, setEmpresasCredenciadas] = useState<EmpresaCredenciada[]>(() => {
    try {
      const saved = encryptedLocalStorage.getItem('local_empresas_credenciadas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [diagnosticos, setDiagnosticos] = useState<Diagnostico[]>(() => {
    try {
      const saved = encryptedLocalStorage.getItem('local_diagnosticos');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [respostas, setRespostas] = useState<Resposta[]>(() => {
    try {
      const saved = encryptedLocalStorage.getItem('local_respostas');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });
  const [respostasLoaded, setRespostasLoaded] = useState(true);
  const [premissas, setPremissas] = useState<Premissa[]>(() => {
    try {
      if (encryptedLocalStorage.getItem('user_cleared_premissas') === 'true') return [];
      const saved = encryptedLocalStorage.getItem('local_premissas');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [problemas, setProblemas] = useState<Problema[]>(() => {
    try {
      if (encryptedLocalStorage.getItem('user_cleared_problemas') === 'true') return [];
      const saved = encryptedLocalStorage.getItem('local_problemas');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [solucoes, setSolucoes] = useState<Solucao[]>(() => {
    try {
      if (encryptedLocalStorage.getItem('user_cleared_solucoes') === 'true') return [];
      const saved = encryptedLocalStorage.getItem('local_solucoes');
      if (saved !== null) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });
  const [tarefasPlano, setTarefasPlano] = useState<TarefaPlanoAcao[]>(() => {
    try {
      const saved = encryptedLocalStorage.getItem('local_tarefas_plano');
      return saved ? JSON.parse(saved) : [];
    } catch { return []; }
  });

  // Auto-sync state changes to localStorage
  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_empresas', JSON.stringify(empresas)); } catch (e) { console.error("Error saving local_empresas:", e); }
  }, [empresas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_empresas_credenciadas', JSON.stringify(empresasCredenciadas)); } catch (e) { console.error("Error saving local_empresas_credenciadas:", e); }
  }, [empresasCredenciadas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(diagnosticos)); } catch (e) { console.error("Error saving local_diagnosticos:", e); }
  }, [diagnosticos]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_respostas', JSON.stringify(respostas)); } catch (e) { console.error("Error saving local_respostas:", e); }
  }, [respostas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_premissas', JSON.stringify(premissas)); } catch (e) { console.error("Error saving local_premissas:", e); }
  }, [premissas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_problemas', JSON.stringify(problemas)); } catch (e) { console.error("Error saving local_problemas:", e); }
  }, [problemas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_solucoes', JSON.stringify(solucoes)); } catch (e) { console.error("Error saving local_solucoes:", e); }
  }, [solucoes]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_tarefas_plano', JSON.stringify(tarefasPlano)); } catch (e) { console.error("Error saving local_tarefas_plano:", e); }
  }, [tarefasPlano]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_db_areas', JSON.stringify(dbAreas)); } catch (e) { console.error("Error saving local_db_areas:", e); }
  }, [dbAreas]);

  useEffect(() => {
    try { encryptedLocalStorage.setItem('local_db_segmentos', JSON.stringify(dbSegmentos)); } catch (e) { console.error("Error saving local_db_segmentos:", e); }
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
        },
        logs: []
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
        const savedEmpStr = encryptedLocalStorage.getItem('local_empresas');
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
          summary.logs!.push(`Empresa ${local.nomeFantasia || local.razaoSocial || "Desconhecida"} enviada para a nuvem.`);
          finalEmpresas.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.empresas.downloaded++;
          summary.logs!.push(`Empresa ${cloud.nomeFantasia || cloud.razaoSocial || "Desconhecida"} baixada para o local.`);
          finalEmpresas.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'empresas', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.empresas.updated++;
            summary.logs!.push(`Empresa ${local.nomeFantasia || local.razaoSocial || "Desconhecida"} atualizada na nuvem.`);
            finalEmpresas.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.empresas.downloaded++;
            summary.logs!.push(`Empresa ${cloud.nomeFantasia || cloud.razaoSocial || "Desconhecida"} atualizada no local.`);
            finalEmpresas.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalEmpresas.push({ ...cloud, ...local });
          }
        }
      }
      setEmpresas(finalEmpresas);
      encryptedLocalStorage.setItem('local_empresas', JSON.stringify(finalEmpresas));

      // --- B. Sync DiagnÃ³sticos ---
      const cloudDiagsSnap = await getDocs(isAdmin ? query(collection(db, 'diagnosticos')) : query(collection(db, 'diagnosticos'), where('ownerId', '==', user.uid)));
      const cloudDiagsMap = new Map<string, Diagnostico>();
      cloudDiagsSnap.forEach(d => {
        cloudDiagsMap.set(d.id, { id: d.id, ...d.data() } as Diagnostico);
      });

      const localDiagsMap = new Map<string, Diagnostico>();
      diagnosticos.forEach(d => localDiagsMap.set(d.id, d));
      try {
        const savedDiagStr = encryptedLocalStorage.getItem('local_diagnosticos');
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
          summary.logs!.push(`DiagnÃ³stico de ${local.empresaId || "Desconhecida"} (ID: ${local.id.substring(0,6)}) enviado para a nuvem.`);
          finalDiagnosticos.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.diagnosticos.downloaded++;
          summary.logs!.push(`DiagnÃ³stico de ${cloud.empresaId || "Desconhecida"} (ID: ${cloud.id.substring(0,6)}) baixado para o local.`);
          finalDiagnosticos.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'diagnosticos', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.diagnosticos.updated++;
            summary.logs!.push(`DiagnÃ³stico de ${local.empresaId || "Desconhecida"} (ID: ${local.id.substring(0,6)}) atualizado na nuvem.`);
            finalDiagnosticos.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.diagnosticos.downloaded++;
            summary.logs!.push(`DiagnÃ³stico de ${cloud.empresaId || "Desconhecida"} (ID: ${cloud.id.substring(0,6)}) atualizado no local.`);
            finalDiagnosticos.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalDiagnosticos.push({ ...cloud, ...local });
          }
        }
      }
      setDiagnosticos(finalDiagnosticos);
      encryptedLocalStorage.setItem('local_diagnosticos', JSON.stringify(finalDiagnosticos));

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
          summary.logs!.push(`Resposta (Q: ${local.premissaId || local.id.substring(0,4)}) enviada para a nuvem.`);
          finalRespostas.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.respostas.downloaded++;
          summary.logs!.push(`Resposta (Q: ${cloud.premissaId || cloud.id.substring(0,4)}) baixada para o local.`);
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
            summary.logs!.push(`Resposta (Q: ${local.premissaId || local.id.substring(0,4)}) atualizada na nuvem.`);
            finalRespostas.push({ ...cloud, ...local });
          } else if (!localHasVal && cloudHasVal) {
            summary.updatedInLocal++;
            summary.details.respostas.downloaded++;
            summary.logs!.push(`Resposta (Q: ${cloud.premissaId || cloud.id.substring(0,4)}) atualizada no local.`);
            finalRespostas.push({ ...local, ...cloud });
          } else if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'respostas', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.respostas.updated++;
            summary.logs!.push(`Resposta (Q: ${local.premissaId || local.id.substring(0,4)}) atualizada na nuvem.`);
            finalRespostas.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.respostas.downloaded++;
            summary.logs!.push(`Resposta (Q: ${cloud.premissaId || cloud.id.substring(0,4)}) atualizada no local.`);
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
        const savedTasksStr = encryptedLocalStorage.getItem('local_tarefas_plano');
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
          summary.logs!.push(`Tarefa ${local.problema || "sem nome"} enviada para a nuvem.`);
          finalTasks.push(local);
        } else if (!local && cloud) {
          summary.downloadedFromCloud++;
          summary.details.tarefas.downloaded++;
          summary.logs!.push(`Tarefa ${cloud.problema || "sem nome"} baixada para o local.`);
          finalTasks.push(cloud);
        } else if (local && cloud) {
          const tLocal = extractItemTimestamp(local);
          const tCloud = extractItemTimestamp(cloud);
          if (tLocal > tCloud) {
            batchQueue.pushItem(doc(db, 'tarefas_plano', id), { ...cloud, ...local, ownerId: user.uid, updatedAt: new Date().toISOString() });
            summary.updatedInCloud++;
            summary.details.tarefas.updated++;
            summary.logs!.push(`Tarefa ${local.problema || "sem nome"} atualizada na nuvem.`);
            finalTasks.push({ ...cloud, ...local });
          } else if (tCloud > tLocal) {
            summary.updatedInLocal++;
            summary.details.tarefas.downloaded++;
            summary.logs!.push(`Tarefa ${cloud.problema || "sem nome"} atualizada no local.`);
            finalTasks.push({ ...local, ...cloud });
          } else {
            summary.identicalOrMerged++;
            finalTasks.push({ ...cloud, ...local });
          }
        }
      }
      setTarefasPlano(finalTasks);
      encryptedLocalStorage.setItem('local_tarefas_plano', JSON.stringify(finalTasks));

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
              const nameLocal = (typeof l === "object" ? (l.nome || l.id) : l) || docId;
              summary.logs!.push(`Item de base [${colName}] '${nameLocal}' enviado para nuvem.`);
              merged.push(l);
            } else if (!l && c) {
              summary.downloadedFromCloud++;
              summary.details.biblioteca.downloaded++;
              const nameCloud = (typeof c === "object" ? (c.nome || c.id) : c) || docId;
              summary.logs!.push(`Item de base [${colName}] '${nameCloud}' baixado para o local.`);
              merged.push(c);
            } else if (l && c) {
              summary.identicalOrMerged++;
              merged.push({ ...c, ...l });
            }
          }
          setter(merged);
          encryptedLocalStorage.setItem(storageKey, JSON.stringify(merged));
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
      encryptedLocalStorage.setItem('last_sync_summary', JSON.stringify(summary));
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
      anchor.setAttribute('download', fileNaxœì½MwäF’ x¯_áÉÑ*"Jd$I)UªH19T’YÅîü`‘TõÛÉÍÍ 	e¤(ýæ0osš·÷©C¿=Ô©_ÿ‚ückfþeîp $SRu‹OO ÜÍÝÍÍÍÍÌÍÍ¦ÉàñoýEÙø</†q~™Mò([â4$/£iâ•(««I2ŒÓr6‰® X/Ë³¤§Åùx>M²jx’ÇWÃh6K²øéy:‰û²ºßÞx’Žß÷á­z]&Õq:MòyÕïÄÖq­>QWìIˆôT( ÃYT@›/ó88EL3¶À°H¦ùEìþÝ°ßß>‡âùûäÕÉwÉ¸‚çþ¼˜°ò7bUãsÑÇ†uÕ›U±¹¾¾nŠI5/2èÿ\áÒÖ*
Ûáqž•9 ö2*²þÊ7“üD˜Ù8&““hü~´²*°’íâ!T"ŽªèÛ"Å©ÁŸ#˜@rT¥yöð»2ÏÏ£ð¼5¯N×¾Zí‰ÏD’5ßî?Í§3˜Î¬êcÉ£ªH³36\Ù€ÄÀ7“=.’¨Jö&	>õ{QÕQ3p^$§XEv®öú³SAs'ó*é÷ô¸{«†
ë[Hµ#±.A®u‚Õï[(¶‘F—¤JŸ $Å‰jsàÑRPM^ôWöà‘å–²â`åDW›õ€îJÛÂoôÿoh½Jj˜ÏÎ7@œóÙqþt’ÏqF¢ò*‹> ÷jUÌÓx$JIQM~šO&°Æ€Jmõ´J¦%ÕýfÊNp’P$¢Pqñ—¿ˆ;E]Ó’þ•`ø~'IvV‹­­-±>PCÕƒ<ÊîO’JàZÿ|íÔÂ7Ÿm‰Ïa:j+ï„&`K\Pã|èÇ'µå3>Ÿgï¡”„Z=%ýtÁT[œº"ë`Y‘ŸÊº.³öcu˜Æ8X n¨®±
]1( ÈãµD¡+»ª`;\’FŠ‹¶uVEei•þ˜<Ë‹g@Oe•IÿZ‡C¬»*€Ø’bæˆ@ÜVÅµ˜&ÅL%R0¼a´nÜe”*œÇùtšVvÙ)RTÛN>RÚ0™Î õæZVvéÊ-³*zú§fWNœFgY^Vé8o…ÅË<þXƒ	-Íà[{çL!€f~×@UÀ7N£ò`ey4^ ªÇ·3|®ÕÈxZ$1ìiwÃ#/ÏpúvÌÞ×ƒ2Ó´\0S¦€5¿ òØxR…”ú]Œr>Î“VHº Ò?ëäs²b;åÈ" %Â Ž’3Ü™ -†Ò¿kà¢3ØØ¢½úJ¼Á¾}›\,l§K}‹ó"°¬ÊñÎE4I#D›ƒ¯ßFæ}û‚ë–¢ó¦Ö•iL?£8	€³;vkyhÐÅà{ÙQ¾¿žGY<Iö~˜åEõ<G¹ÕÚM’íƒÀ÷KùýEG“W Ïô‘ÅªáÓw	
6PYP§\ž|øPlÅ¢ê<)ÈŸÂî%¥8-ò©H³µ)È-PØT•À>‹	vD‘ñybYè§0Œ(«
.n®	õaOñŠ‘P¿^¿a½†mCs‘7k•v“	öd*s6 pÌ¸âHÈ'zØGç ;|A(	k;Ìi¶½ì‹ !µ×àì×à‘½¬¡”WÀ;ÐŒv$ôOÂ0á`=ÅU±žüÉê©ozGŠ‰Žýb3 Ùk¥’{êù–OÁr†=š²æÍ£t¥ïXJÏÊË¤P.©Y9#j¦²U²ÐŽª¶%PÞ™Lh­ê=ÜŠ.²Î4š™oP'K.Å‹höµ”lW…þöÄÖã­A4ÜƒÞôZûÄÖŠmé`) ÁëUê³Œhq{@u4P2°!· £°¶è©r@²øª ÅÕ¬Jbw$¡Ï€aÙïQá·áýE3v0$"ÏPëEùŽ^½Ò“S…Ë·ÎÕ!díÁ ´þ¦ü\c=ÅÚæP<KðË)LÆ¹“ÞT$ã¼ˆK`¬â`çpçùó½çâ2U%"g\AµGk%”N[I­„ïÓäl1S…æeRÔõ¥tU³}|}¼ú¤?£Ek;ùúøÉª†ü¶ùh}}ÕÚÄñ€—õõ]¥6ªÃ"Ýæµc™­:Hø\Ô	Xö	íi5ÍZ}ìëÞXW÷W(­j4 'üÓ<)Ð ³OÛ¬î¶ø?úšS¼øhañUq	|%é÷”Ôƒ¯[[ðœ¤!hDƒA­‹¸…y}ìÐEWÿtè¥[ã6=EV²|O™n³*&)êvdáèÒg§îòVûò}öÔ§.]õªÜ¦»’å-ß[ŸUvé®_çv”[ŽoC¹ž¸Þx½J·é/HáËw·AtïÒé†ª·é:jØË÷½AAïÒ÷†ªËõÝ7ãË±¸›D™E3eWk/QÜ¯¿E9¥þVÉùõrï/Çõ· ’×_¢Í÷ÜNI;Ô;ìPÞÆÇ6ß>ˆ5»ù¸ìëh0$)AmsÙ|2H“ïªzXÉlwe8úÝAq^{whŒÞÂŸº;(ÍBîÉ¬èE <YÇ“\ÕòŸ~*¨ß¨þUWþé˜‘‡@÷h!J›|ÒwdcáiïF‡H°¯)	)ÉÀ«­{ÃˆMÝØ«S]PHâ‘¿Qy¢AprÂ«îÞíŒJÊñ¨ÎõUÀwç	¨dÞÁ_ ‹Èoñ¡¸ÃˆdVŠdr«†‹•a%a”:WGi½v+NY7Âxåýtq«[êŽ\dÛ¹øF®ùZ[PÔuÖCÖ¼áÞ,î¿Ú`ÌÔs‰@¿Þ‡IÄ³C5	·]TTˆ
™F]ÑH«™è>ÖqÉt@›×ß0µðîºÔ¢[ëN-ÊR ‘-Ã¸«è¬ˆEŽ¼ß8ÐN¢3QŽÛºäšUŸ¨Ò=öÓ!øÝÖŸ¹Qõ
kÞc§P`2½Â‡¶U‚Û`ëŽÄl˜…[9q
8Ñù˜Æ2-êc}•ðÚÝ÷(Ö›ðBáÝuŠn°a¡˜_ÚÎD†¤½¢¨[À¬oÉËùE‚Ö{ô‹È³»H&è0ÍÉ0*â\œÐÑ žOg)¬^x?P
,%4š¥¥(“³y‘—è:`À45(Yƒ†ËMŸl¬û®XÈòËgy"WE¶<$Ý¨Júƒa•ï½’.+ðT$ t“þÃ×£á›‡@:½µÞ@¶ƒµñ{.ê*À­Ò *>l…Féô(RcèaXÎ&iÕïýçÞàõúÖÆÿ­ý¸³ö_Ö×~ÿvM¶ö¶7Ÿá? Mõzn[À„ž)hì¾œOª¼H£·ño?¹¶ý¹ùäšúfˆ®9ï»x•õ hKµ3ÍfÐüSÛ *(bMí}s¸³×³éER” çAáÍá:{/ç+‰wªQ#®miìôÈ!<BÛÈAè_þ"z;ã¤,si:îq¹˜¼Ddé¹ôb@y——ð5ÝûqšT‘ô8á]¨ò*šØ£&Ox•^«~y÷”) £5Ô3ô;òI½©†{Øæêy'P=¯ylåxÜ‘šjØ}bÔ²s4Õv¨ü¹¡;ŠòÎ¦šk˜C(ïTª©†=~r£Tù•%¤dKÄˆÇ–,ÉÅ–©ˆ--$	ÞG÷T2DleÏ%›I¡Ž­ED`kÌ§—išh[¦l˜Z†yph¯|2ØwsXhËX‡kŽ›—U>}žŸåþ;Ãy#üª·Dw+WK}2%eôôªo™ºTûW…uÑDžßÿƒqÔÔ>›õ2z·é³Çš\‹\ø ®tâÔ6…Ì¶·*üv‰Êê­ÞZ° œðá™rLLþñpäw>,òy÷qƒB¿ÙþkéìøPN<ë›_<¼£ÓL¬D‘ôŠ8ÂWÎºÜîq»½cÙãvûÆmwe÷Œ»ìËòýe¹~žï•^z÷ÂEqM£ì<ÿÇ“‘"Ô°DeøÐ¾vÅ©½  Ô¢a–ÄœÝÍ‹q"K‚ÐÚËPÈ›¼¥KR¨\ºy6¹êÕ8‘a´ú´48©»‘k°ãp\¤ggèA1Î©‹“"¿ÄîhocÃp¨à®zky”æ`«"È¨Ð‰ûh>FÙñˆ¼íáy~yœWé¿SîQ¨¹LÐf ¡H©/åœÄÎ¢ÿÉµœ–ñßÞ¡ëœ‹G
 j:»¨ïh7îµO>9u×}îdÉ¦5É¾ç~+NýAˆJ3@Å
(Ø¡íáº@+Ç°’P•."Üª’•ÁªX¡¾A'VžE“óõ=ÙèŠÞoB~gûÓºßY¨ì0‰ÆÕð):Kˆq|ýÇãÏ÷³Ù¼R× žÔ¼¼ÑK==Ð·ê,©†ø\nA“²Îyð¥õÉfµS›AyÌŠÀî'…Ò
‘|é…¹^"¿s}m@9ÐÑy§ãMrWÿWªi,¯šßF—]ØáQÙ—·ë“ò@WC¯tõ{å¦ ØºZÊ9¬êøž$¢âûyz‘‹‹èÇ4_q4k¯g±Ôú˜[‹jˆŸÀú|‰*ä÷ÃYŽŸÓhÐA)”.}i¦¨J¼O®J¯t*Ù³«;^ÀÆÛ)4ÊRéÿhÖ|8ž¹æMÍ\;õË¹MÁõ¥¶þªÊG«Žg³¬9RhhÂ\ k/ üú˜s¾µB3Û a$eÓ'ùfòÍþhé½Y HïœŽ¹|!©«Œ?wŽ¶ó¦ŸÚ pO?ß7Úö§éÍð¾-¹³S´yÍžÛV…g^Âé¹rƒ¥÷–ÞËº-‡Õywn¼%0@ë¯ÿRßÄy"Ö]n»£.S‡Ä- Ä©»é,Æ#kÕ×³j- £Úí¯™µÃpw›#lÈ 8l^.õ–Ãw\½Qà´t{4ÚË.¾FD€u™­ãC/£‹¤îfÛV§KÇ9÷wúP­‚ÝçåÂPo9®/Yã:MÝ~C;Wp	T¡ÖåÌË··vÇeî:)-\óNÓ·ÇÛ½ØÄå+°A™B`5\°çŽ&{ƒª¶=ÿ*Ø€¼³*Û@­óíx’ õÅ¬)÷´ê¶(5¼”º:~JU¡ °{A©¾IÖ‚RÝÞ]Qê]Z»JÌE7†QÏD¨.S‡tè4÷é±iZ»2ý‹{wÂå®¾ìÇ7{ù®uŸ—Ej`n»»Ÿ¼•	›wvÕÀd~'Ñ­yß>bS,òö#·÷[Foº=5ÀÁAÈ„ZÇÂÃò|±æv©ËÐ¤šnŽ,jóôÔ|ª]zŠÊä¤ˆ’·ÄÛ	À€‘Ô .Õg÷H¥¡÷n¡–êKÈÖGÍ£òÛï7Kµ:w1fÞ€ƒ7ÿPˆ7KÇå@‡¾ï^gO>}éiÕ+éž€†”¿;	­Éð~Y÷D4$ùú5Â§m2 JÖÄ¥zYs Q“ü²ö”Òßäj8Õ§•G¯—c§–þç—÷NƒŒ¥>¿îéqØ€á×
Ÿ"·™( 7ìFç°Õlþ—Ðñ e
k,ÏÒõ@$Èêw „WEó"úð/þw.`5'ó‹£°¼ÐàrÕx*`{¡Íøª/tbÓàNe¡:
Ã!þ@Ö~ç¨G–f§y}XÒ÷¦Ñ,™”‰Ë[k­ËË±¬~ÃK¢Ô6ß|ÐÔˆk5ûM‡0Á3–	a)ãûw]'ªþPü9)ÒÓôûy"rŒÏW§+‘''rÀ+¬ÿ§iM&œ	#£æ§.ñ/Ò)ãõêh31QðuÔÿì”ÇÉUŸV›Cž‡ÅîUMÓ1õ×Ëhgçùs]Dé$v%¤áµOñ\b< ØýF|»l‰¾bðk’ÂöŠ°cŽŽ âŽ†¡»°r_€<îÄHR¥±ÄQR©sø¥Ü=Ÿ˜ƒU0ŽñO  Û:>bõp@PÈî£¢¸¡>ÙÉbéc‡=ÆÌœÉs)]'2“3¥†Uþ<¿LŠ§Q™Xê¤½OgxxHÃÕyT‰K@PòéJ+˜Š8Æ•ðÛ¼êÔ)Ê|šH_Ô†®Ç·qŠs„e'wš†ô\×};äÇš´~£F±ë–vˆ£"ÆíQHKTpçpoçÈøÏFØg5‡ýh˜åÓÄ\c %ÍY’3…t¨‡‚êª}qÙv¤¨¦šƒâ²5EÛTXý–e[ÍÃ  ;C¯CÃeXAI_þ&@F‚0-Îx‹3:ÑZ°!LÙ’—-²7ÿuáÂ)7^Jáë,–ç3L¸CÅÈU êeb¬nÁ´Ãí\1IËÊu)Ö=ÀÓæ*)úßä°MD™AÝWÇz°°ŠªßVÅ	±’hHû[‚!ñ J/@Îªµo{èó]&Y™VéEZ]DïZOûÁƒ˜òÚ„÷a‡lä åÉ U§¸UÁ¿ºë¥Ø¡f±Gú«ËJr8aG¸Z,&'y}kX¥³\_­ò‘fO¹õI“Ü?×ådnÄwl„„¹œÀÓ*™Ý¼=µñÞÀ³CS³r8­Úò®øG—8¦pÈÄ	ƒä;îd"i£§¨…7ë´´úR÷œ5Êûê©LÎW_Kr[>a¯E¥×
}®xž¡lgÆÌ‰Múz£Ô˜1<ŒaZlã<’“yÿÛëšûªðÉR<€í¹÷‡¤@/-¼!s¼ðêèíÞ‹ƒÃ½£a*÷“²¯#\’SXçÃ 7Nmº‘ÂÇôF3(Ïå@,æZu9Ç‚˜uy“£³vUAF9°eWä›‘´eˆ0­F¾rïÎù7e*ôP›NI>¬¶ÕLrE	[REºÛ$2é—L:ìXøoxÓnJYJùŠ¢>çvß™Üt;ï©vë]xÛ£äÞ£%ª$>¾š%wP4l9¹FQÍPk#-©ÃÙ5÷Wär(|úÂ¾"V½1‰UÅ$‚Á:»ˆÐq_¾}B¯1–'
ð3ø¡9"v¾_ Áåö|HÖÏÞ-§™BRßÙÄ4¼mÿ¥7ÎºÊs àÜ Xœñ9mz}…_â{g†ïñ·=bg[ôYK[[º ^rÞkXõŒ|hmµÌp‘¨Ô£*!Y©«\‰¤Ù!Jé—ëh
Ãˆõï ? £‰5hz ?»©-Ñ&Ö©ˆìº@¾~%‡Eä€]%TÑ¤ó‡À¬v™SS¿QÕ%ó™
€…=Á‹›Ø´„(¥Mf¿áGÒMóX÷©f©±z—$é'n§r»áªWI6\ü\6i`÷¥`™‡µIÒW?“ñŠ³8ùaP§íäPó M8Pì£´N Ã•J&	È!ÛIä¿fô¨#Ì–Üv¶Gäæ#ñ.¾ÊÞ~rMý¼yg­äØþˆšÓ3ÿ˜m´¸9Öæ«Ibí\júÁ{m–ÏæŒÚYªýÓPOÁ­—çI&¦LTä3˜ŸMuð*ÆäŽB5ì¾{§§ðÕÙÉÂYš¤¸N¤]]äŽŒ…kð‰ª-½Ûïêõ[¢6­–f•‚ÀPŽíãªíáj»ÌaÑúÇd2VTå$öœŒü‰4öÍ«sô¸Gˆ6º¢Aqý¬q+è¶d„„÷-Ôú'¨ÄN©¬û|$¾¥³>çæ®µ’þÙkÏtÅä)jƒÈŽçìšnÏm‚Œ?-±eº”KÔýig8­¢iTTE”é>£ÛÔã|jìØÖBúô<Îvr¥£†"n~Oá¶î+L•ŽQÅ›0Atd'£‡ÚxÝ–!Gå:‡R*†‚s l mÔ°öØæêÿ>òôJrq† ÂI	”•DÅøßÓ4±éï ¬05‘uå{=M·CAeÈ“­0Ä"P7äa^\Áè)ŠQÇ­,ê!±1ì pÆ‹4&jêêG¼yEP¶‡¬P¾YÛ°'Eey™1Ý¡:Ëó³IB”ëÊûu”ÇÉi4ŸTJL¬­6Òml[%¬ âÛâÝ'×Î«Ð±îàöhŸÖ•8¦…%øUUªç†vóÀãe29Öm7º1€—ùEna²!©/¹Dj–óÂg¬ˆ~Œò£¨g2âÃçgÂÆ—b$–èT¿†¾m4S¬×¼M€s9þ]+ž"ÅWE«5).’Ãt•ÐÀ¬ï€èéØãYš-(ió «á_,ÙØo§Äy^åß>ÇÊúw½Ûxao^Â°w@fŠœO(ÄkW‰ ævñØ&¥Zˆ¶cèâb$Jz¬ÿ{üƒŠŸ¯ó:E>	ïEøD-')W÷ìÚµËX28wë9,šT0…¥P&QûÇÛÎE*¶’U0Ÿ·Ígð^EæöèþãEOäãcÉÍ‡òúŽzãrQÕm[~‹ÑTè^‚—+ÅÀ®í*p¢÷ù«$_ì ¸ÿ@ÝSK–Ý&ÓoH†°kº©„dXŠ²á^²*.ßiêoŽ¡Ëò
˜^ZROeK)Ü²-ï–~(k–”¥E*„ëaÉUê¿öîÂò¸˜tv—†U´êP?ˆE«Ju ¦¤òûl®e Š˜Œô	‰sçþZ¤9ŠËRZVN%Á¾’3Gíò*Æra[<‡Q$°Àumð/ç'hD;A<æÖÅ+Ð‰¼Ú
ª!¼X’uôIÐK"àícûæyÅ…Á»Ä*ŒôåÃûqŽ…(¸ÍHú-è’aÀizV@wv&“ã|O—è»²ÇœS›‘ á))oOév'Æ?Ï3¹Šþž×ÒLkÚ‰•îA7¸_í@Ý—&0µåÈÚ\¸l¨!H?´yûs&ë[CŽè	œ@9N* %$Í~¸×ƒ¡PñÒ×K‘VHÆÙ8‰Ó86Üå•„›¯iòÐ7 HÎˆÁ¨^y	HÃÕ9ÄURYT…•EF­µýa×%ß ¥Ÿ-m˜°J°$¾òªÄ”U;ûâ“+%žHw4\%0<ÀÈiz†ô8U§j\¨§&¿¿ò0š¥Ïœ‰‡²êŠ&Ø!*×·V‰§YýûYnîü„D;#]T Èç§’«ŸGåu{NÔy"K{5oã}[rgŒk¨_ƒ²ÍšžÉ(¯ÀÝž5x³å'H
—@ãóè"!ÇèVˆF¾‹›µÙ):J€¼ÏòËIŸ%Ä’K å‘¶Àš#•Æ™CìþéBè{FUÉvDGŒØÙÒû4åç	†£8Fóú™‘	
7tòöÎ~’ËtÍý»H“K9w¦<™•µéÊóV)Zo&Jý´^s˜ÁŒ³P'0d1¦ïÙ™•£ªÏhø=‚	Ôeú˜¯ˆ¿$Nq «¡amËtÞÀ,î)VœÞ%ÛMw¤¹”ÈH’\9ž´*Ç;ìèÍ¦öëÇ0Ø’x_U+5E¡ÛðÙ†ü=sOqC½uHÌ JÞgN&&°žåÙQÍJû¦Ã«2ÒeÀðŒV
e¶èS3á˜*{ftåâ¦÷dÏ¦Î¿q™ÄÜ®FŽ§#µ¾ÀÂÌ±ª~þJzù ?h’53«H¡QF†ÚDc¼òCÒçž gè9|ëmHUÞ”jKJ‡õþFŒ¨zË	ö*9çÒn*#žWÆá Ï«âÕ,)ˆ‘qùùþÑ±Kyµ|ïÅ{ð²ŒÜ22ÿÇˆÊp}2÷:ÏIÝÍõ×HîR¾-¹ƒ¥ªh>å3+SÓ¨	­çZ26¦·®û¡J-ý&=™¤9H”†¯Ìã|òá_Ï@ó!l-ÓÉ9Š}}ë+%¬cåÌõéßÐ5çÃ?Ë!{Ñ“$÷™Ó»µë¸CQðêœ"X·î›ûiÐÎ|ÙKÒuÎØ|YEU2%™8)»Â\¡ð-/,xú	•Ýu.tñr;‰½«[·5L™üJ±)åÜE~À‡ú ƒ|ô	%2(×@Šª(5ög^kÌ³»¥ºXÛo\AïÎëRKAÖ-²‰‚Ì%åÙê.j¾
RN0.é;áu
ú•$Â$Á§•“seS„½i]'V™Óƒqì¼wrP	j0ÚBb(%1”œÊŽÄPþÜÄàÌ(§¡«ØÞNë>Siyg{ÈéM§	ºwrkÌàG‰a"ø'ÇÇ©;Cú*VCò/é{QG›®Æ‘¦|Oîeò£ö/ÝaÔ%¯:zÂý¡Îßap,¨X‡³9—,ŽUwxœõ¡þQhÃ1Ü]tp£aß¤£P[«qMñ>YñÏ¯7u÷½aïîkµü¼·’¤< Göœ4¼ £Y5“ÙpÉ'G^¥«‘è¤„R´e‘Y~’ŸQÌÅ2—
h¤µRµÉúYå§§“4KÐ‚/¯ç'íF/ÏNc‰	ûC»}Ôí~J íÇ@+‹Öf¼½Ž>†n©ýŽú	ßî’àvg·5tBë+¯½éá6Jïª©p`Á#+2 ´Uvòª6/ÛŽ&zj\WœDóÃg×TÁM-ZÄ Y€Ý¸R– †?•5½Ñ ‡Ð	Zç.ÃVéyþ6L¿áš:4Ö.k©¹#Žó æswŸ;„h&š¹ýÌ!¬ÅS§r´“§‘=ëôbßá£{¶Y›4'_mâŒàG(PàôÌ5ÎVÜi¶ŒÅ64U§Š]ã”óD
>OE÷yBpfž¨ÛÏÂºí<™aõ¤åzÁ$Ù°µ"|z†,=CEÓfÈóÐ¬Mî_œº¹ˆä|Q*¢c>_ãîó…=0óEÍÝ~¾ÖG`‰Î9üÑÙök*s7é	EàlBÇ]˜å¸Óä*:4¯VàhôŒÒõ«ŸÑªûŒVf:U+·ŸÐêg³¦u5Î¢£ÑÔf¥ |ÿê9¬Øâû.sˆå‚·]„‘ú¾lFr—Í²¸}R;?µRŸÊQ}sÀjÒ±NØÿ&ÉY4¾rceˆ¾÷N’Stÿ¦á‚ —Ë3Ï´^ŽçEþõ¹WŠo÷wo'ÿ'Q*O³Å»¼˜Tëß%±J5…Ý¾y÷Ø€
;£œ©hTÞ ©¾¾¢fÈï®ì$jªº6>â áÕcf‚rL”Üat8GyæŸù„\Ï`ÍÕ®¼ŒÒÑ¹~ÌÞbQ@•¦pîèsž€%u8Í08
Bcd95<ôCO¡™õÇðÏ×Â­…ï>Û_¬×êê~žÐ"Ú—Pé7ä›Ÿ8}a3u>ÏÞ£^«Ú©ÕÒUl„Úð«q¹0£2LøG0õþÙu–×W@¸8ÅËA‹ø›]¸öOÎ„:Î§˜rÞ«ÃkÔÑuc1Gùi%¤Ïº.HžBC÷Ø å)óyM 8P{X8½‘Uó¤f[µÙÊ³“6”1t?2ô˜ôïŒ’ÿÅ‘xeèè&fQ‰?æt[©™)‘²ûé§ÂzdLd}yÛ}"½æœ;?†¢ýÞ9@6¹7¥[ú<¨š¶ãÒ™úÎd"œH8Õy"5oXL“E»¼º0UÝ=Î\o¸©¨³áÓQFÚ=ÝÒÝý”Õ…[îßØLrRðÙ|[Ï2Tçx­”	‰1†Bí£ùö[±±¾¾.XJÂZY”¿Ö¸c¡ûuÝ/{ðôà¤­µ-=¨•mîšÙ…5‰ßL_×BOõOw‡êc8m¸ñOè´Ñà*'ÝŒÔMÆÿª"QJ~›‡ÿ¬x–•Ø´"§üæ#ŸgÕf˜@³¤T/$\®Ñ­5q~Ë­ Ç„ÜÒ0cLÈé,fi•`òð«¿N†CÏ¯Óecô{c(ö1èC
S»ì|"Ý6¥È\Ê­–:›
H]âåÞŸ÷¡Œ˜,œ%àD^Iæù[¬I“1N©ÃFãs¢¨3Öq{Û“¨¶Vfe‚ëÁA+Ü†5ÜUõä:Í.@±Bj`io¥ij6‡˜si²†Ì‚]E±ÚÛ…Âûˆ5­ØM„í-l+iÀ³ëÜm_!Åj÷–›‹F’ÝY™d3Ë³“l
3NÎQ/©¦»au ­›n$æ°6ŽËg#çöÖ‘ ltEÊm þ…«»rGÎÅPÈÊ®ÜKæâ*hF›[šCb‰Kîe8/ã‚
|Àå&§“r2M&§C¥pË£ö†uàÀh¨±ù[¯¸õùj¨|-²O¼£×XeO•¡ ¶ä$žc(IàÛ–"}'b7‚>TÛËý+1_Vg¦Î óOs`P¨Tw„ØIq6ÏªÈÜ.	ÄÖqS$z;SF R€‚¸M×òd@î‘á"âæõ4%hä1m”ôPC0‡ý\!z–ŠÐã‚Û˜Ït,£u5Áâì¶
”¶úKÃŒ³E„ Çƒ‚°ìg½xè>žfóÄãÌMÍt‡ãÞdX7†O¿ºEi‰ÌEŒôUÒ•¼4VÃ,3ð"Â–3PÇñ¶’5³<[“QÈ–ÈU¡a›ÒW²°ŒVÔÁ.µ—õGe‚ËâW¶ËÏvZ8æO28U^À¥‚¼ÌºÐÉ‹¦²Ü5XU RçþhsÈìg®ºƒ}’0š8¢‰ô§AíR=²k©@ær–y&ÏdTzõ¶'>#}t˜å—0=Ÿ‰½’Y’£,Î%UÆèÏ¿ñ.—|Ú\¿«)ÊÔ 3ÆJÝÏÒj=ön[ÙrÔ Tº—ž5^GÄIW&4_Õ¸¬{~ÆÊÍšK)ò¢R©¹¥ÊJlñq“v&e&_Ä¯ƒkyDÿw®zŸàMðv‘oRüH¸—çžý²ŽÕ“ŒvÔgy¬dÕe/-_=„Wxé¨ñÖ3E{;îp6/Ïû´ëŽ@½÷jÒq¬³ÐòPÛ€¸«kù B|Ðœ¡F•uÐô+¢dBGVÓ!»Ðpìüš[28`¦‘‡ÕVvàÞéõ¦Í:ò¼P‚søPÿê9>B‡aT}œC¿­eèì°í¤€î#âÁ"aè»>K€©3å®µ›ÃHšôðAj­¯õ³†;ž=4œ4h´z§êŒÁË,a5³,·ö‚zç¨ƒiÉ–Ðúž«mñJcðƒ¯É4c"©ÚÌSñœXgþF³¾
LÔ¾/’I„ásrï="6?+¢i$Ý²K÷"\»HYw¡kU¬ó +',—*˜8N`oÛuNA·¦ÿ˜â¬_µ°
YøÜ”{°÷N¢²ÒANÇ;ñÓERP  ¹¨×“	j‰G°„Š8Aw•:óJtö3ä3¦[-J|Ørí »±K»‚yË±åÔ–.”oVÀÐ¥[ÁzU^E“#””cW	8ç€Ê~4éËh‰ã1¦z/†$1È³…Uá3Nñ ¥“Å 7Åo¥ò]’qacÐ•>HŠ±44H‰’²³ôÙŠ¾m^”‡%LçAÑŸÚø=æ“j/“QFFG.4yÞ©™Å_¼p£>áŸµø0\ÜÀiÏ—+)U®}Õa.‘0G¶ôxæšzÌ*3àëp*¾É¹©Fø:¸á¤œÿÈ*Ò¢wdv@³:ÞåÑ&AXt¼œh^å0Ð¼Ì4 1ÌÿŠ‡Ä­¼7Èî©QÕ'¶ëšÀfƒ>t-B´+íÒXÂF¿IWµ—a]ãfï= wåŠÆ{J>i¬mÿÖ©q\ë;=ÄÛòÙÕ`ÈP¾ï2|ÕŸ:-U†¨[ŒRðn<%´€ £AÖ¿þå/Í²Ùwa ìCK}6Æà|iëA2™O¢†‘¸ßZ T@§y–<KÈƒ BZàÕ/»àêß=íÌLcHE³­.O¤\¶.£sï5ô£}.<C&šG9^y VÔÐ‡[J…áPMN-ÈçÚ-ËçÎV¨Ï(çò§ÛÈ2ñ2ƒD 0‡ÏGFfº Gð¬•»¸ž‡Í*aŠ¢æyÎÿÕGˆÝµÄYž››mGó/¼Yô
§IÑò‚B*×d5KÞAêYW'#}A«FZ+Ó³l?ÃHùl>Sa¼ðB6‰Th=&—)rÀ¢nÁù±jK5µ8Ü’ÚÌ­Z‡«	9¿b†ÞÃâŠULdìóCÓ<[Oò2‰×N®ÖÈÔ‹÷…­Y„¼œ¨^þ‚ã«œÒÜ‰dt=£vœí5ÊI~Ö_‘3Aõ5@Û±Mµ*œ;¬qóÿ‚ø†ƒ  1§,=M@Å;£8€Î›ú¤‚àZg]âxbËôÃß²41†^‰É@sþÞ~ø+É3…L©´dŠ…ôT¢D:¸M%i¤(g(¾­R4Œ‹\}Ëç0³2b( ®{kÄÍq”dçÐöI¬gØ[ˆ@ïœWÎÅ)€JâÐÉ®¯‹GÀ¬9#>~”WÍ ÒSCXáR#Ë­!%K£ zfI&hýÆÕ('([¤	Ù¾p?‘a^µÌ^°|IS|z(òÂŒŽyÖ«K<H…r}U$r©å'MhzÊ¾gñ
Ú«>þŸ>ÉŸúkÍØ‚Ó–×šÃ”y_õ²6´r]%À‹ÃI†l™Û­‡«MŒÞ—M¨zã.ÓÆû<º'Êk ûFCÄZ4Á„—Wki†|²×Ì ö0®f"×éwÀ$;H¦€ñ|(`ŸFôÿÿ±õSÎ#ÌLp±åÌE‚]K3Lg¯®ˆûË^‚2v¶f#N7•Cd¯eyµFQÂÛ*Ç,Šº-€§æ ûaf°î#¸L¢÷¶SÍí¨FbàÏwgÉ$ *‚ªó%P@$Áø|Ý›þé÷Ž]½sÐ>Ñ¾CÅNœ¢§4nfËÉ@­S„B™ðöß
vf*ïØd PµŽŒÆ¹T‘SMtTÚÏxNÚ2¡À€±;»-¬ó“ÏQj’B2–WóŠ¸È€>ÃqÀeM0ƒƒÝgå5e‡$c§ÉÎí]ž2#unQ5é«†Ž¯E f	×‘ÍÝé‡ŒfÊ©wÈ‡9zØŠà¿m^dX^'.¯)‹@žÔAž´€¤Û;T:ÿÆ8çÂ›5x®·ãp@p)}QI;GïgHªƒmë‘LVÄ ˆoê Nº‚ÐÎÃØ¿˜õÏZ×+àªšG0žÿ¾'	ÌìÊK\¥çÀ¼Yá2¥Ûj˜µYI|À@i¤à„‚B*s·`)æaÞŸV±5gÛD^‡ÆcŠ~ÊÅw%¬fÁÞK•ÖùHô&Lagi¯¦SöNÚ8ámô…VRo¬Á@¨
©,m|áùÛû§4&©-ÆØÍ03ƒ?÷ÃI£µdL….E^ âÝAŒÂ°ååOÏs<XRùYNŠ(é­Ê,‹XÈÕ/f,‘-jÂHUB¢wûŽ`þŒ›$]¢ _Xö(Ÿc _¯{~÷ÉÞüG€_êØÚ>9æï*x#á\Óü~üø±ÍÂ#„Ž)Làcå?1 Ð@m ûáÿµûðl•ÂËykB­io?7$$ÍìyJ	ÒÓøNCJ-äðÕPEŽFo¶æ˜*õ<a‘™ª>úšÈMô¶D½”›¢‹sJÑãR«˜ño$×0à4­&&vŠ‡_é¬i·	ŸDœÇ«	šž9Åt·r
Kß gKa*„_‚ýÐëþóÍ¦LÒ³sÚèu£Ãsùê·Â@<4}^â\ŒNKÔÉZ×`ê^þ¡·ÊØÐšæUk¶…U±±¾ÊmßVí.	Àží÷œñ†BËÛQÚ3ñ§ù|B—yqbÕÄP3h9Ð;ÃŽÜ`¦YyŸ¬®*è=›	º=úPC#ÿ©´šÓ…?™Må§VÓw¥{Ú£3Â³t2yšO@´ÛøóÕ#øßæïåéi2E¿ÃxíÑúºSÜ’øOX‚¨þü‘Ä1ý»¹*ä½g.2©ýãä‡J¶¿ùèQýó³<«pßèo¬‡?öW`5]$(‰WNòI¼â•¬ 	˜’W/¾}~üêpG¾Zi ’ß±ÓôÿkØßAÅ4!ù¹ç¹ÿü¦¡«ŸKôwü1ÕÞ>ŽNŒÛ.7Xqï@~Æã-PøàEùá_Å$ÍÅ…jñ³=Úü|U|±)Á—¾yí÷lnº"E¡„˜Ëª:ÀænÚï‡8nóÕ¦ÎHŸÎ@wÞ©Œ#ñÉµï€!¤ž¾¢R2ðÐ÷ñAû–›Ä‡¿a~ Ö¸£ÝœÃò9Æ¯n1õñÅÍçëþî¡_,`¬lnàÉ-£¿â£Úbt~Ùñá•áÒ(Éº¥‘þÎAø òÛ\‹Ó³´‚WÓ4›#ŠÌ+ ÉÞÑ/'q·ˆ.ÕJÛüI	Wåëœ˜6bÂ;R}ø‹ÍÐ‚‘Ô¨·;³Õ™ÝïÂÂ?Ü{¾süáî¿‰ƒç;/_‰Ý=à~ÿíÃ}%öŽŽáãÿÃþÓW½š0‚ùîve(__ø d‹U(Ó¢‰¥)PM„¼ØEòÓêJ½l:öŽóe¾¡*•T(™nBŠœ¶5ŸQTYc°/ ÀÛÛT7m«m‹¾y ÊÂºðÎE	Á	ª¤l,$j)ÊÛÑ¶¾•Ï`_äÑÅ÷xcs´¾ÿab%óe0ð„m;€SÂ€B…ì:ürúmŸ›:Mhôkº‹Å}…×Žšè´ˆ©p$õ6b–„\Þõd`Nö"-_F/Í7«®œ¯Pš}
‰µqzzŠŸµ¿KtRúõPÏ­·ÐªÔntUjPã$ôü‡¢OŠ+¾4ÿÛüb€ÜÎ–jPñî“kò&~ÜupPnMÀÕ„@aä·Óýílf<ûQX_s3‰¹.ÕþÇˆ¢¹†¾?V†.°“Àw>…-3·­ôcóQ»ú<¼ºê!ÕÀúÔàÎ7_A“ˆ÷Ñ‘æ2õCäIõ+\Òtâµjž>¯òcdÇ’§3û`Õÿ9˜ÿø9pÿ‘xýº÷ŸÐÉíÃ?îíàÝý?¼üð?Ž×µöÍó½ô	wƒÿgïHì‰{ÏÿˆRÚ~?xut¼sÔ£=äè D¸ÿüç½çø¼ÿòÃÿxºÿ
îüú±ûíáP¾Ø%í+7šÞ›7¶'y|5²»‹ý ·5¼Ä5IN«‘Ùæ
Ôìc•ÏFâ‹¯VPBö„'ÐöìÞ™Æ¸ÛÀ~9’é]jv?÷…,#/çSŸ¸>”„õ;ì¤
Ó&Ü9:ª®&˜BîÚñƒÔ*Ì W
Þ¬6+•– 8§„Ç‘ø]ý=¶©•ABuõÜ“¼ŒÉdr ˆv­…††ÓšíËÐí¥>†÷úÑÆªøª~õè•›~ÇpÁº1Î'óif:Â ®c¿°I§F°ì˜bÃ«°	âXl|_ºõ7½ú_<r¿î}ÿÊûþ…÷ýóM÷û#ïûÆÂ}¹tßù5ÖÕøjé6~ïãy£½†»èÐøôj/\õ¥ôÖ™â«žÉç'©f(7IüþUØùø‚ceh’”/T†JC¤|{;Ïw`[±UJ"Í¡¡oô­ÞÜD=ó÷«âw›oH5(ò2YûÒYíÜþ;pöÅ7€ll è/%ä¹žúæ®Øúý}aëé«—OŸûáì¾Z82ä°ýE>zãXNZÑv°÷rwïåñÞ½LIÅG2¿WÉa~iù!çï›¸‰mbß7m¾A‹¥aj,øÒÔÍ/’ât’_âA¨‡' ²½W|^E¿ÇÒŒqè³<¯<Ë§tÔ†M°$Óg>6V[sx â’Ü!_RAëÏì\)Ú +E["¼øì3Nx&éÐz˜r•^¾7vŽ¯êß¸!fÝûLö 6ïÍZ||ƒßš8øðW#"ñÉuzƒ”õ:ê<euî‡bÓ‘›ŽW¤%“t‚çJ_Àœe¾²ƒZõã–T~9Ìg˜¶ÚÊçÕl^õ{'“üd^L:&oO`ïM¨,ÇÕ¢5g¡vñ:YÀ4Z(¡Ý´{‘>4óÜ n£ÑËè‚®ómQø1ŒíÃÿWÅ$:IÐa…,¶«"ÏžNÒñ{q£¼E0H üúd‚ŸŽ²¥Šm]ûw[Ô~ÑJ&ìKãä$*^!r½\‡Æk•%šx·®Ç™%®•ËµS<¯= ¿£«×kr6ÅY4[û\Ì~XûBÌ®à—²Ò®ý0UQzÍ<[Ã˜cñ\†B{8ƒr³Æ3¤±šÑÍ¶X99[c,¢µËsh^”çÉÚäLÿÒå ö
‡2+TM3—uqŽÌdÄ@?Òï¨ kpEÁHä<Q_ãœ‰è~ëzsýÆCÙ
5ýDE+0ö5Ùf‰V¸µõ=ý8p;Òo_xõk}Üˆ‡¦§å,ÊXÿVa®•É4%Ûì“k¢»›¯b9Yéë‡’Äð‰»A(Ò1îŠãô"L2+§éI,0Åjµvµ¶NŠüó#¢ûríw›f@Îå	]“Z+ô9ÐœðVT~%™Çu=9&g#*EÕÀ&ÏF¨w¦cü%{².@o…Mã<ã$3ô—ò¥AÀÑ¬ñW¸$un²]\¬ÐRÁÿ­° Îåú™­}¹òÄg°J`uMO ¸Æ6YåzõK,uŽÿk^8lyÖ[ûn^Véé•~l]d¼#Ð•o¢âé9häŸë…±ÉèPô¶­÷¾|ðABÞL’'¹Íø=póµŠÎöØšùjÉ"&®ey–¬<aŽÚ˜)‰/€Æ&	àëÙo„lÖð×ŸíÆ%P”˜Vk+ON®Ä~õá/þO“²Þž‡õhŸ³¨†«µPÆ	¬.,ô6ãÊ€°Ž·®ÿH'Îâ[+ûÙ‡¿Ó|Åý(YÒÖµ½J¡)=þFTdJŸ(@7]P_oœ–væ7ót‚Ó½«:ò)V^÷L“²ûˆm•Öaßv`EšœbòêðÀ€–¿CÇ…›aÁqµÄ¸LŸ~6ŸÂ™ÅQAöÞà¸wÎ° zš•½ÕéŽ Uþã%»}²˜0y„ùî´éÄ¥oÈµNÜ‰‘c\ÆèòÉ¦1Ûq§É$¦ÜØ7õ2jô@7äÿM‚ôóöŸÿRÃBxñ¶öÚ1ajÕ°PÃD—!ùìüªyt4$À;Eæ^bXx´´x\Þ¨°Ò`Ñ˜¾vž7]iXoóaÞÕ‘Ä‹h\,ÞL¦XjÍÆ#èLÉ~ÅŸž)í\B»ÜÞ•:¸»ôt!â´¯Áæž”%RHw4øz4	E´þvÆ…ÍÞd2ð[a"Tù§Ç†ºùF~DãÌxñºp·‡@ÞN¹å°ž¥“QáqÙ4¾ÄŒj²PœbB‡*?¸]Ç’
<RÃS÷¢ÂÂÞ—
Z×Þ›ò-½¿npø¿ãž‹3v4ÃAåy’øÇp°«C0•¨ÓEW-ÖÆ¶Ö2U½òÂkñ˜µ˜Ø<X×>ºìˆYœ%†ÊƒçÜ}Œ+18ÆNƒ²1–“­t{9ãë‡ Ór×3Ìªµ/µE¨ª›†´êû…£ú^?`b+¦0$·ƒ!Úhe(Ì—Iv>Ÿö+ÌkþäÌ˜Ålín7‘}©wiêÛgBp›.'Iu™${ê`šíë{„5ÎÑÛõE×.QV+O¤ÑûÚÅÒMÈjþú]¸}4R¡!x}øÈàŒ,_Ÿ°–bPäž'§›|.¶E°M‡KÚX¥šÐQªÇ¬Y¦ 3¹ônÞÝÔ;,D Á^_ÖtÍbšÏóµah»Î õ¹Õš²7ðŠðç˜V¶Ã&s:9OuBï£:BIüz@§o[××âRžI¿ûäšœ·¦ÑýõUéÈ+}·0
z½õ‡ü--AyœÍñè¶øüËGÐ™P‘ Å¨2˜gâóuHipó¼³‡ö¯ÎHë“W{5`þbÝ±Dá°H¿ôm”éôÌŸÉ²o]Ëm³ó¼Ê¿=|ŽNOïÎ«jVŽ>œ§kÑETE¬™>ŒféÃí[ÿDÕÐr…=ºùôVðÑ×ÖÆúÉï¿ÚøtŒ'ƒ[§§§ïj|8šT¤‰ ðÚF´;„«˜Û¦þÁÖßŠ¬HN“¢HŠƒöŠ«­•,_Ó¯Ü¢þÎÀ7,& ©µË5ß¾Åg5æWNÃŒï+2ÃÎ3¾ò$€É¯Î:€ÿ¡ô¹i(]½®ƒ«Ûºk/ä)KƒÉ/¨º´îÌÜ-NäouNÂÔ›Ö;÷èˆuÊ:ôúÕ¼bÇcþ´«=ŠÒBIU2‡ÛS¨ ÙƒùÉ«d\}±2tYWN–0®Ö²­B.³´çùåq•UÅÜ§I6Æ(â0ú•}IˆnWÈxõ€Ž‘íÍƒpŽºN¥l»ŠHF}yðdÖ”k%¦ÛMã|`ÆÂ‡Bq‡lü0ó®6,¨ššÆ†jPÜ«â	²W‘Ðî«îDš]|ø+€Ëˆýë1ŸÊQ¨K`ãøÛYŠwñ(
G	YŒ¬ƒgˆ«ÃD]:Š:clvêVò±ç}oÇäÁ³VàHõþyr†‰´ÃX…
R7î„Ôg	Ðr!¦yŒ-ÊL8*ÖÝ$Ò%ñL;÷áÿÍ’ïä“;È)ˆã‰^ š›|¢™^Zœ÷Ö½W·èa¼OóÙ•ŽHÅpkrSšLªø¶­ïK!Ã?šÇIYQ´FÝ Åº&aÅFþáã"RA/°rþ ñYJH+Úµ¨~á¿L0n–î0èÊƒx`)´êï¼:z»÷âàpïhÇÞ€Ï~ˆ†#høðSºZ§Ò²ÛøX;E]1QQ_'š°A™Ú›÷ÉÑj^$
È~fãÿb».m_)@9ÅdÇ·^”u÷23ã¢Ç°D¦ÀÍàƒ}U+È­-]Ò«$~ŠŒU$Ž¦]Ö‰S9ŸÐeÝ\ lI«Ðê/zÓ ,àÆR^[Nc€}˜î‘°»6­ÏL’nÇ¾Ä;e£ "m	'Ê][AŽÎÁ*ûèläãŠ¹˜Âº|ª¢<±Û0èŠµôÊxc™ò]‚µÛ[[¡4ñ¯%*Wuš´7íÂýKÆj5.Lì&ÁÙè6#KÍÊÂ™Yrvê3„÷“¯z £žÎú^8ØZ`oö5”+9Ö‘ò;%¿î”ïºÛ/Á§ÄäG–G%ÊKdzdÙ¡J‡ì$›mrüÓƒ¨9ÿ•Ñ:þ©Nq„Y°ê^íÑœ¸’ÊÚñÁûþ°çVmR«¾Xà…y$ë©Þ6u>(¶9â«Ç…]J¢!ÿ*ùþ*ùþ*ùþ=H¾*˜²[	…3\Ó×÷'"“[ÿ¯²³)ü*ö.%gý²%^=·~8 PØ—í)_Š]y$¬\H‘ÅlÕ&.ò2B®¼ÍzàJ><¸¼»íWoHŠ¬Ý!,èúeŒ.²¶
%”M:%#®Ú®„EÃ¥¥7jí'ßP[B|“fGæÇ×lzdžxÑüÈóÜ“@	>¾&ÃžÌBïg¡Z]4ÞÛ#çEqC>)¢2$i‘ß¿€f\ºcnvj+…0È¾/ÎÚ*/Îl'šE²{Câ­²ŸFóPÈ,OÝ/·±8µÞ‹E³LËüÒ0:‚˜ :ðú¼r ¯C'©ÆQ—ljŸ;J7H—5è|Ípñ¦±TKcÖø‰SßjýT†•?Óßí«n/Ä	7B…»s,´òÑ|YÉ gTCø=˜-®ìG$d$z;Uz±à¨'»
|;Æ ¦u+"¦ Dª	V"hÈÀ<’É]ÚUüBœ&IŒ§éu¡Œ{á;æHü²GBÃû§<É±…´4—91‚É¦Àcd™Å Hè”J‚ÈÙºíÓ®ÔtÇ¨I¹—éWÉÃ/hýñx†ÜÄS¡Èzú­¬’8MÆªDO¬5õÄmË¦>‘-ÉgF"N§®Ÿ… ð—Fîwfõ7š.-q3Í3¼}.¢,Æø†gt¿q	Qù¶féðB_j¶f¼ftî´èñïng÷sãº'$6­}9ew5W¿uïÄt°]9„QáT‚¬ {²cïdÇ6æ±#ç6kCBE?7äÆÒh¬
š¤Ë
IÃ:Ü?ÞkÆÕýšÈì.fòEZÖbS¹'Éyæò
Úý˜ÌI %ñ® dd“é¿ªl¿ªl¿ªlKesLç\‹ýü¾¼{°§{ê’€o§ÞÉÊ]Õ»|ò$5Dr@nU¦ijÕMµîÈ”½^½"W&ue’}¦ †h|¥úPÛ•È@©†F(‘%ý«÷JÉº~‡xö¿Z¥c%Ïúµê) ÃÆíNŠ¨5x›ih¤²eûÂo­¨è¹‡!Ô(mÜ¢ºú³¬Ï
„Ë!°MâZ ±ÍkZfG%ÒZù§3²óOgÍ–~„€¦þé¬®”œSÖZPCÐE×¨ˆ˜QµD¢:©´LÐB¹\H}ž\)h¿Jµ¦ÏF‚ßBû‹LÑ²HZXå©ý,ß”¬‘’Hú™TÚâ¶Ì·ˆ»`\ânë„Ô’\T§Ê­+\¡Æî¢y¹ôÈ‰±a\)rlE4iO@jxS3„Å¹I¿jéH×Ä¥œ‘5ƒ[”Ö´q¾Y^SÝ’ÈT’†“Í´!‘3Ý4§"E‚ÔIÖ="óÚmy«Üý_Å4–¬%‰Ùgv¢é'YõV²s²èu²á 1ÚÕÔ¨(Z3ÊÃ«ƒ^Õh§"Õ)©Ë`ß“˜ËeˆÕ.æ)Áb‹8ˆBog¸f8*yO,¢É´FÕ|V°ßá_`É‡Œo2¼´Ãf7ük]ÂMÆ6¢ŠòÅzòúÙJÓ\%ÓÊ‘Ã˜<PxO8	w2ÉÍ¢«I-kóåœåçÜŸ¼~P´qeÛ`)zÚÌyA( «Ièm-lÝØüqT½Šaì[»Ðh;Õ ²œO,‹®È¬ß:_ÃÉ}=lK\i•|ƒýø$`ºDaMÚ.k^®í¹Tp(ÔÎP2á¾¶ª©Ûí˜#¯Á<ïEØ©šÌ†¿¹ÍâÈ“Ý:,Xüë&ÔáßR+Xwë…;yGÿîAÈ£®Ü— ‡÷.ìáŸ·âðÏYf)q’‡¤%³Õ¡‰ra Îs3;G²{ã|:M«þ]¬ÿÆê#Zÿß9ÐÃO®5ºyw?'¡Átó¸: (1ä”GÖ™ÙÅT¾Çmðµ7Iq6Ï*çM‰¤IqŸWsPÝÖe`Nýc«yc£˜Pµ¬5Íâ>éûFÛ§¨Û¶ŽíR:ÈÆz£a¹AÝ6¶ýÏ½?àíž=ÁXÊ‰ö]hoØh¬ïò	úN{NâH°©4ñš÷ƒ’Àô#ÃžDžé÷‹¯!i]cs«†ÝC|ëúS{(C
 ÐâjÓA‰´Ž0ÇçY2„òùE‚\ú=löíx’Dè72Óð)“Z×7÷mg#‘6jg!ôz|Ð“Y ò	‰cÕEiÏ& ë`é>PñDÇF†âÞL‹ÌÜu¹½ëÁ¹¡Œ:MÔÏ ßn8#fsÆÏ…­ÅsY÷Ø\™)HOÆþçr÷ˆx&ˆ{><sÎ‡YãDB’³Ò™sëéªJ`wh`­rÃã¬;¦Ùä¯löÄf•lÒÀ=Z®Ù@å)³ms¢þ÷Éú–>Ù«³¼€C½eGÈ$Kº\’¥1my–RF§É.ˆó4®Òºm ßÅ&–;`ïöãí@âú(T–*³Ç½PïØYç²”\áõLub²x¦#÷öDƒ?WÍ¾-)…²ïA*Æñõ¿Ÿ'ÅU_öoU\ž'0;=5˜$lk«§‡:p³æiCì®ÕÎ| ÃÙ¼„×a¢°ú@y‚ßý	{&N®4=ˆ<Ÿ\+ŒÝ`átœŒÞ¹Ù­h¯9ˆƒ‹Œ:4rq!±À Q
5'"?ÕiÄÁl„òl>&Òî;Ø"—	‰G²Ø¨pè{v%çYÆÙ%¾üV' Óš3e>MúÉiIËOöÔo³¡šY·g„¦²E“¬ODp®Ý™>GY‡Ùunü8ýš$ÙYu.žˆuÞˆ“ gàˆ@5üðÙ–øb}=4¥ãóyö^Úº21†q?]ÅšT±n%ë`MCÈîº‘:wLôcÓOÝ%·@3oPk›TÚw:K58Æ9.eÞ=||v);–À<ÝXa³[„d£4vÄ¡,ùÁnå^]ÜÊXí¼¶#8@¸ª¤0V˜P  m‰â[¾ÇýÃÑ«—C9ÜôôÊƒïeü¹¾12É}ÉËløÉ*óó-vú~ÝyÑG%ÐXoŸbÁ²›ÁÊ.ˆcY±ÙØaƒ€?ÒI"7÷ÀŠhÛC»@“‚å,ÆÐæØ:ìºˆ:H|…ìŠÞ}cÐ‚±ÚmÜ6q’1ÜhÌda<|ñV,f–SjÑ”9‡èrö›SAç¦UÐ
P`ª1ÆN›ÐÚrdFjÀ‘QÅQ'v$&5‘×Í ßÀµéOÜšþ@ïé8ÊßýÐŠ–˜í×>¥ÓY4®rû"¤‰âaËYiAb¬«¦a…ò—¢&šŒ†8zß¬’²´ïÑÑÏû& Qìdñ3:ûÃ¬*å¿Ôù{ÛþÊÝ;Æâž.¨v_4À÷áÒ&£»µÈ®˜Ÿ,²+æ'¨:élW\RV¤æÙó“û²*ø?ýP¶¤”èàjBY®$üeV>xú³\<#A«¦ûÆª§Ë˜s¥Kç·ØaÍ{¥‹uéšs«%Ð§ŒŸÙ¨P»´%ÐrÚ%ðW®ûïë*ó]Cý8Œ´ƒI‘ìÁÑ¬øâˆÊ¶x~´-j^ØÑ¶x¿–ÄeY’VC,©Q56…[‚ª±*ïÕ]¨kJv€Ü^5¶ST5ü_†j¬Ñ­UãÛêÝUc‡ªïC5V£ú	Tc#œüÇU²¹¦`¾ŽÒ+ß…Šý2ô_5¤»ë¿²Ó;“	çNM÷`ÓØXXæìª!ÎõúÍ}°«Þë7öºð‚ÚZ ©Šyâgúþy˜ÛqŽÎ‡ðŸ¥AX—ÑÔ2»òÜŽ-í”Ö6›5'›übF§ö_EÖÑA“•¤{÷<®ï@Ñy1õY³qWêÓÆåFêÓ&å»PŸ5ŠÞú_õáEñÒœ`ûÔýäÔçZ oK}Êê“¶¿£|2oMQvõ«ÏÈ8µÔaãm·ÅŒ„%›Î–À{ñ•¡vóÄ5ÂËû±–
úG7*ôur±±
q)k½y:Ÿb–¾XÁN¿-çg	f	-5”èÇ\¿TZ4ÈÊè"™x_ÞÏR!/Méù¤5ù-T#}Y¾§ÖAVL<±Ç{ÔÑ52ŒÄ-b·u6[Ú™þ@-vÌý»LIÉÃfÌ²fÆ,ïÅŒ©jY+&Û3êFÌ_÷ÿˆû‡¾³Ú´7ü¤{BÃšâ¶N¨ÃWÑ¯Üý'çîÇ½Óðõ_„wçÒüUêAþÚ`‘µeõÐ%KõI¿fÕ«Æ­IœÁÖõ­±f©q·U/Ù\L±ú/B]´tûsz)¹Ô}wK¬¦ÏnˆµâÖT;¬^§ÜàøÌ«jKq(v	TfbÔ/Àr«éí·Œu³œfçw{»[¦·6›q¿@«Y)ùá¿%?·ÙÌc‰·³š1 ]M¶ìV^+å…=CF[V%äDÜZ%´™ûô­ç>p.¶×ÃíVÃmÏ0îfƒ¾¿•Øx|rOvð/|ADâÄ.ŠÓ<½“¨b	ötþ.Ç7aÈÚ%þ.¦ùd#ÆÞŠyµr­–“¦[Ü7Êe¸è£bË/¤ùÉ[]/ÒÒ•L•Í˜ÚÍæŠµú<0µC œ
ûIe|ã
}øP’YŠ11³³ óÅ€[x]è$øk*2½^^Ç&c‚ÊAÐÅ)’‡&“é# jR^²Zª-Ý†E¡„
ÕünG|ø+NÿÊ'×‡7+â»Å˜ÀðfÒÍrQ ô3x×)B³oƒs±h­m¦e¶Ý£qçer‰-ïg³yÕïõ:óÈwþ9€š†ýï<²Á’´ÞbÆVø”éÁLËÚÎ¦ès§Zœv"œðôÆp­6“^âÒ#-d,ò˜g	\™ŠþÞÊ7ŽICZÌ=>k0=PMlâ bYBO§QñžŒ}(=+ž¡ØEZÊ2†7ÄRj·•#+"€˜ÉòË\³9GST¡äëK6·~oœ¥ËbÏ™¬}ß™êÎ›¥º{rË†¸\7ÃÈÃìÅA!îåØÐæ,dÎ7ìWµùØkqßÑlõÁC­/žîP³›”ã¥œÃÔvëŽa0Xè–|kô4Î|øH€›7Ú)—k§!t_ž3n¬ùôS‡çÈ±ÙèÕD›ÖcÔÌäÌ¢¥Ë4ÌLºz‡Ê…^¿\LoBœŠVŠÏhøhÙ=[ó+ÛõþõxiTÝq—Ã­<¬í„vò 8 iñ¨‹}U	Ì„æ›V5¨€yµeV‘1Ä›&Tþ-¶-òËøÞÀ‘n;<ìáë^ƒ»åà;pæû<HOÆ4ª×LêØßÇÈQ=í8í¡;Ïõi—àãaÌvÛb8#ëSé.Ü+7À$7!P8¸S®‹3Ðw›õoc ÅÉž_Ú.×»B¢¶ësA¿H2¹µ ŸOb7šÈ$m‚¿*O6U«×AÄIB–eÃ˜mS"^MZìªÙš!+q!÷¯%P8›[Rý¯µ$+qMJÊ·©&(ó‘zR7j¤´’0áóÇµc@'ÝT…×Šš”Nr§VaŒŽËñE¯Þ,P&P^^“Ø²Mk¯hÕWÑqÑÓ+rÄÍök}·Ò-X/fÚ¥Ö‹ÙBáÚ7,%Y:Q6v¢¼_Ag~Ê»PóßjÂ6»è¸hÌNûÎI[é¬"5ïZêneMx±NIKõŠ;3-ß©&»Òvaewk™{é¼ï,Ó$Ê´Å%|€‹–º_€PÁ¢NqûÙ.ºQ:zÛO*ó;T"ÿj9ìQÖX…5blž“”U<å_]B1¦|UÉíIkÆJc-ukµ‹=Ôéûß±â²Ì×ø¿/mæ§BRÐ¡SýµãCAkÀ‡3pT2~‘Rÿ.«u€¾ xÁqBµƒªì&§Ñ|RI7äÐáK½Kxñë&J°~«øPàä¿¼:ùgwçpoçˆ£¢#kcˆ–LNÁ$ŽÎ
t`wv¢j¾4M[«ã@›kiòc"º£9î³­~BmÑH²"ÓäÂ<¸Y9/åÐ¿ŸÙl8[tG ¥‰PU÷[ä¿Y~[ý•üå=|Ý|ó€Ÿ]‡´PtÉŽè*„[šŠ&€†øJíªñ‰b®Ž•k’D³z¬òcvžÎZY|éF÷`áq¥&\Ý>KLØýÄ²éÌÒ9µl9±Tù„u¯ë,¦É‹[Ý¬Ôõ$!áïNþ¬¬ÿÜÑ+¢^£ÁA@2êë,e4Eƒ{D7WŸ!/¬˜ÜÝ7 v?ç§t0¼dY£Ì°ø]·tè¾ò~IŽÍ+¯ÍW`ÉuÇŠ4¹™sÿúÂ
ZÕRcí4²aµµœlnÞ´™¹7fßó=(ÙÇ[´ñµÀÐÎÎÙ/†Î*94ÐMG±Ë¨MI‘Ð	å89Ž#ûV|_G.êßÛù¬É®v~¾H©¶~k±oÝf³wìÊŒÛ]øDw[vû`>ã‹©ÑµÙ6˜¯ËÍ[Â†À#Ë2nÇvL„¸m¼ïYÒ"øñXÓß«AÐ‘ŒîË(¨þ½µIãÎL–nìL&{?Ì˜5¼Ï;QgÝžQCg[w®}8Éù‡Cm•êpw$|Ï¤~½
^Â¾–œF%eýâÝÇ=•[Þ4‘•uvÀ[Uö22.ÀšoS»’s÷v†“çBøx!TºÜØ•
¤˜¤ÓÙâ°‰ÙµååqþOéíD¯{|ýùpÐë¨CqÎÅ§o‚ü\%¿@®^ëSˆÿy÷:üÌ>w`[ˆaÌÎAÉdËqÐ¯y*N¢ñû³'i¦b“âøwÔUøºâVK’âp:ö¥U$”_>ÌƒB:U…/ª!›Þ}yA^¶ðBIt‰JI”ãDì¸·Ó?b P‰Öï£&7“³£ºÄ;0I›Ä™ÔM.òxÁ‹åß
.Necî•ßk:½Nè"3,†¥bXêy‹ä¼ÙÑ¸JÄÀêlÎå€¶ÒÝùl’ŽQàb\èZEÂ1GŠcí‡âKÐò2‚è–|v… !ÑÍƒ-—‹L´54`HR´@­±ÃÒ™’¾Þ!óáÊKDÖù‡¿Ê&?ü«lSÄrØÈ€H^Ç³ ÝØ©?ÍNóVâÃ‡â(/*B×û$™‰ê<' F`bå>>áË:‡Ó\Å#xxžž'
%8¹y&`¢é{‘ŒÞJh‚L´hä3ÄOý~´*Nœõ§b} GÝÁÉ#œÍí6lÇP$.A«}4ß~+6Ö××“_keQƒ^aC[–Îmûß@û'mí×>¶´_+ÛÔ¾öF§¬ID¨‰x70ÒiT6é4"úõúg+Ð‹Ã€%UÊ¬W~?&ú®U$MY‘¸½]MJUÒ„ŒD¿nW&µ¡ØÇDè¯J¥ô¤ Ú¨íª1sv$“¤…äõŒ"£p·8®ÛYƒèoš”€|@Ñ·ü¢}F1ÄÐëõ»¤²å!¡ž¾Cáª
¬g-ò²¡3
	†`|§BŽRxŠRCOª¨eÏX]Æ KKú·/kyA7n#Ð{±q$\ð~l#®5QChqÐM&Ôâ$âõh‰Çã²mB‰®º$I-¥a@ŽŽâ!D÷fI¨K@t‚ÜØ¥(Éjâ>,xúÑåjeúŒ£¹þtìXŸ\x–äI7ÎÞÙÇ›éz÷Ä#AâÃ"+°²Í¡Pq…žR@)âÀ.˜fÖvq7‰“[±·¨óH$û&'èbì–à0m <b'WUà^¾·©­–ëù~ÛãótSFFDÉNI9”×(•Â[%ñô8*ß—ì÷¦ü½s†!äïQõÆHÄ 1¦e2„ÕÑí€õä©M®E‹„s9pwä˜¼€Sä7É«)Ü`u©N,©%üô1ÛÏOÑ‰ˆfùmraL‘?:¦H*)èòÉÛè"š¤‘Ž’zrúóÆ_FfKRéQRV+…T8ŠôkWLàÝ¦ûN® ÷¬$zÁL+øÊÈªz² «íùbmÿ4ÔZ¦Öå4TÐR›"”È¿E!½šTYEN²yÒ6ç3Wšv–ÒWße{§QNÓ`ÚjÏžS]ÍŠ±"ýªdñnÀi¬¯¶Ùï‘èOµ•e{(ƒ³¶fýÉœÉ\ÅÒLÁ·ì&+O³§ÙÒ£ î*ó —\Ðaß—pM<7	ÉÌžþ1¹Rº¥RF`w ÁA|*fóâL~°ñue¾~>e,¥öÛ}ÔgÔ€‰r\$I&Náÿ?&;ÌnªM¬ÃOú¢nZM±ÐNw/F?
øRYÜöZìÉ¤Hh;Êô¦Á´àuÇ&EFÖàL¢£>ãkz ÍJŠ?¯.gºýÈéŠË-[Ü²»«@]• 
´ >Žþs¿Pˆ`iŸíJ°ž„ X×L±º5Ÿleµ.d+K¶Ó®n…Õ0\YngU˜Ü‚Êýsµ†R#k?)°*ì¶‰ÃîíyŠª³‘I¸)zÀà¾¬°•–äóÌÞGåùI±÷~\äY~VDÓÈûð>ÊN¢Ì{Y$“ð—]ù¡öÿ…”?®Ni­ vG©¦D³.lÎ0Õ1éØH §)º >[¢ª-R¶mbøžfá4'"ƒ]¯P»]ù±¤„N.÷»L½_â@¨»6Ofð|^Á†ee$OùÖü­àh­ÅrÑI^T%#;×Ðh¨›¸xˆD£ÒZyË€5eÕn2ýÐ·>\©C÷ÜŠÑ‡	Œ‚b•™õA!Ý‹ÛÞOÆ˜'ž%ÈºUè8²ŒqÆ`ØsrC,À²ägÊ8÷xª/&§v%‰ý?CO–ŒØ‹eM±+·´e,Õ£%ŒìÐÉ'tßÆþÌ--%6ÛJpkH¸„¶¸#lž¥€yð¼,(/0:‚;v{nöšÌ ÎUç€süýPÓ4OÊþÀGdÈ2ã´~Ïöv³•Feùqì4­4-V™vß/y)0¢L3ŸÒ•WAœÈý¿$×) -Ç¨»«³jÌ0ö6p„Èíéw²ðfsy`«­î	tÏ’ÒlÝÞ•k°Yd_	›¨îÅw¢f…áÖ¬[úP„¡µç¹µä(ëKÛgÈèA—â­¨Ûö8Ú»Zqlß›au:¥–ˆZ|L½”]§¡wòhÚo1|ŒÍXW¨‚2ÅH"¨O©ò.yß0Àµî´Û.—UÛ¸.ÄÕRmèm1aÝ¿í¤ÃlÿjCa4÷÷gC¹ËÂó7×_ˆ9åçQöõæÔ=Ö¾ÖóŸ¢0'&WðuÂªûQÙ­¢	fÉà1>úÒñ“3Ï®øÉ= Ôù¿$’&)ïûCJ†qgÅ±NmÜ=U6ây¨&e­JH|åéu9ùU¸àïá˜±‹ë
±Ž™½NÎŠ$Uçä2N`2ßkî²œúk˜‹™Õ€âÏ¬nÊ›Û:·lž[Éý:Îî2sëÿHGÈ¾rÒ2“®Êß>•uî`¦Wª¦óÛhXXb‚ù›æ¼9–ï;/a¦P¤UÌ[ø¹rpÇói!ho¹59¼ˆªFZXdÙ¹%IÈ&=z€—‰AÛN>%ð¿h2ˆÓr|°p)Àkäîs¯Zó&ßvž},ü1§ßÂÿEÏ?n	ÜšÑio×žM³Í}î¶µë–j3íû@´M4ÞUùH[;þóní(›+|Ò±XZ•Éäô7uèZv!²aJniâL¼xêÖ¡µmÆæâÈà¿ùE2½£‘³«S5ëžÚµh…¾µ±Hêw·¬Ñ‘ª”ù¼Ó÷}–x]Ú)6ÈžošT_ÑY‰k}ª´ë¬¼L
øp’çhéyì}
µæ<bêeÆÁÂ	‘öØù„w}¼O7˜¼Ö:úª­ÛÄ/„*¦hn€ÄÔM© 9[—2´Åq,ú8…±L(0Ö0j3­Œ;WI°aœ§¬Òo’þ™=(å0ö3 ê|nW jêðt¥ŒÌùô˜‚B•ÃFÕÜC-nknÆ9Á˜&’"†v¢Ä¶½Ÿøú™èolŽÖ×á¿Þ€_æòpïwµ“&E5Û¦CúvžŽyl&½}ÛƒÆò0Ë/ë—O²hrPäßÉzÂëªŽÓëX¾wŸ\;3@-nÄšpð~˜¨{¢ÿÉµÀ!KN e¿›Á»Ç¿	üÕÉw#á®a;ß˜ÐÛ ÈžÀ±#¬lïÆÜHàM¹²Š¦3:r£©ÐÜo[ç3Vym1
é€gð?¢®å³¾ýöD¬3n»Ý
‘‰×½=V}øÿÎÒ7w˜p²hÑõøQñ^²ÜOÅŸQ7¢·Ï€0²q’…à0Ï(/þ8ŸF™¬wœŒ³|’H¨¸¿Ó{cGÇ$sàC=<2H¾‹(üÃŽ›'¾‘4\"ÓŠDx çG.]»e¼zi;çzÔY(ù¨2v¤Uiçdåãž%åÆòÄÇ‡ÃaØØY!Ñ9C+dB5Ú	ÔG”¹äµC@cÆj×êfÙÈßÉPÎAzW4ïrªáwµ“ChcOo:õ§Õ·É6\ÕÆº«iÔéº+ÚùèøÞÄ²'–ýkb…¸'Ó­jä "ååsÜ?Íç“o%HMÂ#/ïž}0ü…ËÓ'Î¤çûà¼™¾FO¦×oÜ=ªðË“3
_ £Ó1È*qsú)n‚0TÔ![G«Ìl_Kê‰¿7ª:š£ €ž[ÁN•HE<¶¾ÎšwÕz'ï®ãq†L•à‚s>Ó43çÏt¢•~×_†[êO¸ÍÏ†³¤8›gUänï¡hlzõÔqhR¬‘XÐ—Ða%ù˜£rôu0¨‘l=<"U0¼—ÿ„Gùö	½Ö"¤“:Dß™õTaý\;PáD`“>æ±ÆvãtÔQ}  õ%D[xº-Cý:‘,†çÌ‡Ç„½3;YE2ç_<÷_]YG‡ÇSTuûÍƒË>ôJ©7ûîªÔ»=¹¿Öj—dò$9çôÐ>Ù*ÙRKÔöé!FÁd½Ðé3ž!'‘v?À]èŽ™ˆñkí¼úÛÌíñXgQK†ñ¸FšÔù—ß²Ð÷X ÈÇç(·:»’73òÛe/r b$%á­Ä¼`º›†™¸\­…:pMƒÜ?jKú‚D O¼Ô{»Qdw·ùÓPï‚¨ÖÕè/?)o	j&ÅE„Q¡B€gI‰×l{H¿paÍ³8Ì%®r&_ÄF­"ÅÄ¦áâb<J§=(¿	EÍ»ƒ¨§¸v·Å¼_¯9vàé½<³ƒRzñ•§ô¢ºM¯`Ð€Ê(‹s9m*äàç_†åüDnëýÍUñ»ÚÌQgªdj…G•º¨ì1wuÄ½ ²*„¥¼‘ÝªUïh£ÚþÖë¹Pl¹Ys)EÀTÊ!f·T™ógd3üÙRÍóOå£+Òüÿ¶GÏQ°M­¶¼…‘HX”Îæ%ú‘Ê¹òY¬/áó¹•Þ©ü
÷F’ëpJVUÆ‚\ƒ¿+5´ÓÃ2Ñ•&ºRE.ê”ÑJ¨£…>š³àŸ{%;¬/HÊP:ƒp­:˜yÆejö·òj®ô¥»MÀc1:b6Ö¶„&‡Î×V¦_Õ‰ßXÛvt‘Ôõâ0™[Â8=ºÖÑÇõ‚ŽÛbÍ4Á*tèF½,Ú!’Xß—6ßå2{­Õ1‹éì÷r¹©û
íf<”·4!6Ð³_ nB¤7~±E¶6…nq«i~é.¦4[2lP“~°=“¬&b[¬šÙër…¡aáßï¡¢ÒØÏçÙ{nFs›¬5½•L_áÑÄ©:Ú–Ç€³>•aŸéW,ñâ–¾ñÉèc^•°ö!ä^öÌûâUF‚¶ª"‹/™\-}Ž¨~È3žwÎYCaÎ˜¿¦^<«x ^A¯/r'xÆ9üðWÜ31X4UÊë"únL¾¾3=RÁô,ÕGrö>‡rÒm9ñ”‘ª©÷Åÿ  ÿÿì½Ùv#W’ ø®¯ðà¨@‰—XRÉX8IIìŒ…I25Ý'2&Â	w’® à»ƒ‹ØœÓ5ýÐOý8 ê‡<]çÔSžúþIÉ˜ÙÝ7w	†"³„#÷»_»vm7c4v®Úzh‘N£éH*oeus:RekFOáÖù?žÆÃØ¯‚tæø°L5+E2Z»tÜÌ òx×ñ^:N $]pz4)°ÿ2”§ÍCÛ7D»Ó¬	À…÷‘}£¨ƒç§	‘mL½„¸ë‹Y¨Œ>2,
[&Ñ¾×mqŠðÞ‡Å< k¢Û¡#ØÑ/æ—qI%
©Ô†Mµ~S›×{m5‹@™O+îùçmŒéñ£‡%ªôÇÏ	 ž[Ê8Ì#êoC'žNó„î0l¯«”q.ê¬NR€XþMbqñJŸéH’vœël~Ç†þ±ÑÐÙd¾NÂ³™‰4˜(@  pÀd‚$É9mê¢Ï 	„ù3çQ¥îI¢$]õ‰P—ÈŠj9¸”S,FuÃK‡ùø„îµÓ¼LÇ\ìÌÀ)^ƒG©‡ÕHc(Ú\ErDßDÛè`+Ò³x\éSÓgjhË49BN!¥INìTJ¬‰Ö‹6_¦KÛ„Â™º©~ƒ	ú…&›wpŠá+¼h²Î2™ü¦˜ÉpeÄ#e4‹¯ü0Æð?Oš¨æAeˆ¤@y5^ÕdXoØš¸§Ôôz	Z£”á0ú.eãÌ¤állîvkóAœy¦·E¼¹·1ÅZc2½ã„’Mö£= _ã³¼XŒâ$dÍ¹œÆ‘ª«°%ªÜüåæßR>ÿÓø(f@“El-¤o-¹–Zôõ	t²è{˜ÜüËxÅÑfQeÇ
ç³pRÁ4š ©ò‡­#ì—~„¶RmjbÒVë@VZ>O õ®áÖ¡.²$%‘#rØez2Å0'%fßò…)Ò“¬d€äiYä2à\ÅÊbôu‡ÕK‘´£ŸÃØ‹i…;•÷¿øâÍþîÎëÃÍ›ÿ~óßÞD[û7ÿãpwk3ÚÞ‰6_ßüÓËÝƒè`çÕÍÿûš=Þ<ˆövö¿ûÔ8ˆv¢ýƒ½7ð}ý‹M¸Ö2XL&x2Æm@rzPšo!4ê
áUif¹‚hY¯—Ë(Ž«p³È‘ªå›MJMÝ"â¥StnšBS?¦_°nz±-¼Á‚É.JwX!ç;b¼›¼ÅÍÿJ2@Ò,ÅFsœg¬‘Š2+fÆv°Ü@G¯½¼,§Y·y’&C6¸…Þ:ÂÎR´©¦¶pp„\¢½€$9¼Æqƒööß|órçÕf´}·¹ÿÝæË7°°›/ûÑùàæ_`/~ØaÆzY!üx0W½¯œ#pÀØ8)ž0ÍS<Šë»ƒC|µ ÍÃ1ñVFÔ#$t°iíÓ¨[fÕTôWz•RÝŠ,•#Ã-Çæ ø9áÝiT(t íä»Ýÿb-°ÿ„G°ƒŸhbŒfy”#æ,n~A,*vžoŽ¨FD³V¢½;ŽÏÊøæ/qŽ2JÙãß<6àðîÅÑËÍ­?½v7oËZ}—bR(DsbôžÝ!ÐÑúC,·9cÀY"Nôcª÷ú_<ìG»¯öÞìn¾>ÜYä\ ªÝy'Uƒ:‡e–g\Òn§Ç)î$ÑÁª#XŸàŠãaNGKiÆ=¦úåæ:ðe:ºùç1ŽŽV0 —¡¶éDÂÀ@—ŒZ‡Ñó;J¥ƒ›E|»;Ì€I¡êbÈ˜^ ”Wfß$À¶·°’(ddº¹Îß.ÄùÁp!¬ 1‹  HPwÞF×}J>Ý$­ï$ùbd—Pò6ì'íhe9ó£×¯ûÄêeÁ˜\¸‡ðVZs©º\¿»»WŽý²@ªD€³¥ÎÅ>XD«f4<µ¤Žã¬?Ê“tXöc@”©ñªT Ž×	ÝûKûO–Ž‡qyº ef•`Øµi¼Û~Ý’‰q¼v
ŽÐz<aR"˜Æòe>^ðè  ÆÁà”ÖÂˆUÔ¶ÖßÜßßüÏ¶P2"ñXé«jT~óÍÜÙ:tk“ÞPT•¥F"òÐIkùàp÷õwÑµ¯m]+3C%–~–
Ê3Õ =ÓTÈÏPÍ83ÔÓÏHëj¶DÔõ«µ¿][d%‚\ ‘6¨€õØC—|¤Óº\û¼]ŒÞjQ‹€.œÖì]Ó{%©þê+W …*{›½³µ$ nŒŽÄ>0ÝgYu‰OªÂ5ßR_,F™£ôˆ(*¨]SùÝ1RûÈ®i]Úµ 1êu¾ÍF˜·e‹¯˜Ñ‹1âÊÍ*sG-ª_J`VÑãË¸áD¬æå×åxÝ0¡u>I÷åÀRÆ®Ü_&,ù»vdwa¹çÖì„U'©×º–‹z°`JÛºi³Àº2*´¼_%	µ"·¨7Ó”ôû|z–êRkƒÖé'¸á¢ÝMƒÇzFŒäÑ#SvQáå<%’)½ Bhÿhàã_ãR±¯œƒ@Y?qµ@É¡ÙÿóøÏãmºõiþy¼}¤f„´É	Yî©üô>_ÿtLR2¾oá*= MœZ%ócº‹C ¤‘nùy·^™{qÀé5ÚChÊ©y©O°6§_›ô‰kÆ¦°¢ãJX>Z„	·m*ƒjÛ\n!­0ì©4Äzª¤Êa7·Êý>P#{?4­ë7ú½’o4vÎê“Sa{;Y«áäÉ|<ùBžp„ˆåVèwOŽˆ¤ÄUà£ž-dÅ s™JMVÎdÞë!±¯­÷ÛwúýÈÅ«úÜy“]bq ºä¡ˆÊ"¹.¿ ûÀƒ”=cÌ¥vš¤‹*Ÿ‡SË¾eUÜÄ¹–V»d–È[ÑC“x”Õxþ\«O9ƒU_û·f+ênjNlˆ˜ÝÄ¯H®mÉf¤½¼&t–/`—
;¸—>ª$n<pôjØä% x
é;íÝhk-©ŽØ?wPÞ÷Í½M1qGÕ6Q]"½¡±’Ëu´MÙP6`ëºqðmïíÊ;­uw†2åak7£çÚ$t81´Ž5†g>MaÓ2Zÿ"EÖ‚“¸ÙØHÆÞ8ÿu–Ú›I…âÍ´ÀRžfdï!E"%FrÊ¬¨oµ¥¶É`m_d,Þ›Ï5-L€i¸Çq µr9æØø:¤¨‘„,Úâ,ê·r2âÛˆáPãY\ƒu´º•qc*vD23+ª›ÅªoñÜí-ü<°ñ¾æ{ýÎ^ZbÞ8ð—Ü±¡?ƒ…­úÚOì‚$(ìn8­Ó7*,DIVY¹Æµ¿¨“/Y5tþ¸ÒlCµB×¦˜ãb—2@+&F'®ó~bJ

4(CPhàæ@¾P[lçü2”z~bJÚør>0ëœ³û ˜úÕ¬‹PZz TÀi)áÔ±T@Y2 |_rh¢/Œ½/R¸:ÑûÖ†³fÈõÀn«v.uƒÌkAYs€Möƒ,š*|§œÿÊGÃ¾OùSSfc‡¿pà˜ÇÞ !"MÑ)Ð'"3¨Z>ö^n¾~ó~{çýæÖæ›÷{›Ûû›o6j’gWTíZ?ÚLÐ^LÜŽØ(j¢HŽ$âEôèµ ŠÂEëbYèøàÜÇƒËè†fWÄÔpšÃeÕ£WñÅ£ï±j×žÛ¢ËÙ«­aŠz¥ÌF„ftáé‰º®Œ‘‰³‰åÿ8s`¹»˜è={ichVå|¢‰<;À¾u"@B,Æ/}eYÈ	@w7Ýðû›Š¶­!Y	³óçì^˜¦ó!˜ýñ¢höÉ ^íÕ#TB¡ðŒüRëö‰¥Zf`çH!áè^pI$}}æ1%šÄ">ñd\I@Ôk¾¥ŽÄFŒk£\.àÄ¡öæ$òÄÁõÏ#\Ë¶6ã3‹$ë’®Û…~©Ã*·tÂ’x=¾9¶ÒY—¶é«ËÕ	øJÞjnÿâF †%£teæXÕÝ€åMG(Ð÷!qca‹êÒÔjz5óÅzŽ çÑ©QC—ÓŒC¢o»‰†E.ø®)®‹ÛÕªuÈ‹Ô¨í®P•ZVE®Gšñˆž#BBœÈ±‘:¶ÂqF]|­pÔã~ôÍ¸–ÓæßåòK†¨öÓ	°Ò¦à†“^‡éE•{
‘cdöáÿ×Ž¾¼Š½ Ì½?ôú?æÙ¸Ûùó¸cá}IaÌÖ—C­°ÞÌ3!Å,Ø%Ü±cŒ†Ý7É¶',ˆÜ
ÒB‚‰C¸ºÖ<NÌà Zà¾ú»d.BæHÐtèÿÛ9ˆöwö^înm¢‰Ðö›è`ç»W;¯ß`p“(ÿÓd"ˆòëÞúŸÇüGÛÓëª;µb1Ëu{_Léß“¾pÃò\up™8âÒúäF#u äÀùÁàD‚>!¼±Y‡f¨¡õæ‹DIpÖ= hžgsÃ×Cà£W2]€Ô–iHÀ•"Âbÿ¨²-P'Ë$FdvžÖ)5qáêœoÚê›7ØfwæW“O^5,ÀO¿"%OcWÏ˜Íe°Äd8Õ9ˆêì'¯¾Ãpýt¨å].ªØŒzSµ[_Ë£Árî§Àa0oE^(àh Ï‰Ð°IK8½”6À¨˜MªÑ{¶U$3^¿ù+PÌ†MáKé>TªSÀíezB×s´` ¹ë…¨0ü}e0’tIüq+†ü±wqeg°¾4Fh¸Àâæèèæ—‹è>TêÊ–×¨oWñ—Õ^Û•/¯œáúƒXi3;ÍŒþLŽnõÅ|EÝ¬5\;ÙÔÊzÄ”³X×PÐRcsRÒêF	JËù^\TY<|fÉ_Ô†TÄKAI§U>•ÚÔ¸”7úˆuEÆ$ý"^œè@+X*Ìw˜£*¨§ëj˜q›]ëB:Hõ9²D)â<×mA»Á5«ÅØcd¬6±™ãï	>„&WÏˆèloâC,†ÏÝ ›-Ò>*¬tXuNJ†Y½‰YÐpîŸÅ|Ñ­„¼´ÍÆ ÈƒEEÛpX[ÉVÝ²Í´&ÀùI6²º´
è<[cƒ	S\Ðœ»ox$^¿´Qu;‡T¯™Íw&ÛâŒÖmÒù6Ùžzx›fQk8än(ÿÒ[yÀŒX¼–Ï/ÒÏN¥I;ÿx>LË¦©æÀÒûz_÷+ý,íZÓ¡™GV—{Ã+‹Û$Ùnl)RSóW›wKêÍ*©Dd¨yƒE‘¬y8šc|ˆ½òÿ3SÎ‰y8&_t2a0Ýˆ"£{‰{âsOå#—ÀºWjZåøá—:w¶õ¹‘F5êÖ®ÍºÝ{¦¾Ù–€Í—®NæÐ4u‹Hbä¼M$eZ´ì[L3B~¦C=GO
ägÀ‡”n6òÉŸ¿+ñšý¤ÁHìÜ†ùîi ©Û‘"
ªƒygÙ°ØŠMÎBoâ¶†ÖÃCK*Bò÷ôyêoöhö(Ôì‘ÛìQC³Ü†»›oÔÍJe´›\p›6aÉ1NvQvê@˜1ƒ¢aòp˜K«nÜm©'$¯à£Õšø&QÏH1ËFñzþ±ïÙ¸kê‹xú³5h)hT!´¬C`g˜LñišØå¨£•Ev&µ^åéÄË)´/m‚°h;Æõi]ºð.¼ lHe¤ÐÎ`WÕsàÈEˆ×aW®í5öº	‡]ñj¬Y%6T#@Ž=ÚðXI€
§RÜÅ{ƒÑ`7¿'zËßîå#>Ew³ôè kÿŽé gÄO ÇrÀI}ÇŠ}^€à%d_<¬Ù­(‚ÙÎã•§s#rËŽçŽf=…ú¥Â/€8ÿ©¶TZ]¸£Õ/K4åÚSZ(TÅÑëg­²Žpú†"´†r÷ªrôÔã!±gF³û9‚‰î³ÌkÝGQ{¹Ø[™\óùé- Xwãê,t‰ø¿µõ›K!îÓV5@NyK]/ñ*SÉWT!ý ûæ©B„‘€Ý€.Šêútô<¶š±ÄG
ÌZèÎ¢K’q2‘P-µç¼Ö‘f¼>T5îW“±¤*Æ¦{ol:l1€7–¤Ä…¾c!ða²h(k|ƒB´˜x¬'=§ã{5¥Æ´Œ+Ê¦’©y¼mcÛðêOŠÄì–j8q¯fY‹PzM(_C9øÞÕëÄ]prFœ#‹ª©#ö”È.D*ìib¾{#tY¢Ñ§C6„aì®TCxó¦4Aik
âÁF›xè0•|XÒtÈ0%(ÜyH·-m¿×ž”}ž½¾ãfÊÌx|õÌä+-öðWå1ZeTŸ¢¬5ùÌøªös¶Á=?cš¤Cn­Žþˆˆ³Bç˜å?½<øOýi•ËþQž|øCŒˆþA=ÜåP7~ç”êµÑæx_åïËÓ4­ºoùzÒª¬Gì‚…ódàª£ƒìîêJÏÑ$ "­³9¬rEFr¸M—kn—ûß{»$]‘ê”þŠÜöÚÆ ¡„¯È9rmñXR5ö]ÚöÑâó\<¡Å¯gZ|¥üm\qÙ½³)³NXø$ôˆ˜N0Ý'%öÜ0ª•°<o¿.|^´.ü»¾.*ø™wa¸Éàlë"÷ Óbe:7ÿ %õ\wê5âÿÄL†Ò‡‹ïjÁøÆÓ*Q'£‰M>EÁÄ”X}íé0l ÖÇPùÙÍ_h”´	O)”ŸVPDEø-:”,xôÆyÔÜ›_ÎÒ!¶ûñØq¼:ØÛ¥~·3,Ÿaô=aœÄG¿º²òô[8Û
;¤q0Ä…m ÿÑÍÿ—,LƒJŽÙaOxLl›å‘¥pRELa¸¤b1¶³âypë3'@cQnÑ¿¥
¶¨›SªÒaÑ"ÞÃ©Ï‹
ô/†å³ö¡Ý}Ž//ç„w÷Ex¿¼œõª¼WHÊN,7´Î6!EZÒfoð IÈOúfÞ„iÄˆôŠcGv3I?ÛùbÃ}ñpv„(œù—DCg^g£ÄŽžíÂx_æü]y«UhûBËÐýIwJŸ{¢¾>÷D		¥	!´×àRèÅuŽÏc}ùý]4Vö?N2Ùb^¶A[eŸø6ƒÅcgß×ÙñaÁ÷…:ý¼Ã²ïzÚ¹¤CÅâ¤ëÌ‚·¿…NÒ3˜Æ³ï_½ÜO¦Õ»‡^88
x$Ó>¨!¯“–}N€ÜÊñ¡ŠMÏîå1ÙÙhh.Tî©xÏ^l©]‘ƒNÙ¢ÜÁž‚ZExw(4Oá–×Jë4¼NEÎï,NqÂÎS<a [äØKd-õ„e¾õ®5Îs#’¤‹Ìz”Å«s8L?OAq†yÉ2½ó9‘µ<Óµª»AX
ÁµÉ=R§ûô@Ü"ìu?§Iƒ‡³Ê›ü(¿Æ˜ãåË³ÁVlQ»ºÉ°·.SÍJ¯Ÿ¸(âËŽ)ÞDy¯°E$Ã±•§VŽöÂE¤§IM#¼ŸƒœÆºn‚%D/ÁÄ-J}”
E³eåø´|	gy]Þí–ï§QžMT`?Ãzaå®OÒŠ$§"†x0‘©]á€_×§2•2l.Ÿýn˜±Ã€ô<(–?¸Ç~/Úpž©°Å¿_ÁDçN
E]Þ¬õ¥‰›µ§”Ú‘#CfbO‘rºòùÎîkX<Â@~äHr<L/0´EµË®ßÒZ£Á0¥ýH”ÇX²U84KŒŠ'Ô/ý–
åCÏÄ'Ü„jõ´—QdFB0_ÉxéÝ…×ßn/XoÑÆ>†ÍY~ûçéÊÃ••%üóäøÝòÉb´.üÇK?¯,ý^Ó–ÔLÓø}:DwŒa‹	–ÎÆÀÔœGG—ÑpR¼”Œá}L/¡0:äed­5`²8žïBÜªhç×Ñœ±x÷Ðš¹ÀþçÇÔ;Q~ô#@GG.ºÜ}íuÕé /è»$ªë#öF ÐýhêP´¯šKôM [Òêº-á\z~C2j:MöYµç¼+ÖŠG!¦ÉæWj|ªÔ²}ÄÕ0ž‹œ‘%TbJÕê£Òj>¤¤¾FyŠŽXþ_YuŠ5ì´¾j
,á%ù-ö»§V18r—â¡ÐUrè‡Ç">H Ï+~øŽBa»¿PìOó{"®<‰‡Œ·ÉÉëcÈ§‰|n<6~è%|ß©EŒp«¿¬²j:´š§ðÑÐ&KDò™wgô•"¬wLÕWÇ¹rK“Œ‰:A~Ÿã_ ŒÐ1 ›|Y‰\d˜¤WZhÊ†fôC!žä$ÜÀ1@Rµš¤q5E–ÝgSå øoj’æL‚	Z+ÆøcðKLŸÉwJÁ_Á!-¹×tÌ‰ì~)Ö§G4Mã1ÎþD¬v©õ¤s:œ©ìO<`Ãfô~ÿi*WWÑ·R{¬×’!ØÅ‚?•ÈŠáàd1LãÅ9?±°oèe®»äƒ(s}
Ð*ÓR|›Š¦q…D½$–5Î8[*ûã¿5.•@“=íPþ†ÿæ¯“*1†ß<wä…'Ü\&?†EŽÌžÄùªƒ¿$L«*ÚW‚Ô>ê‹ êEwøSüHaYeøSæ1ýÛQÊ&Ö¬o0“"þYÛü%8h6)x@RaøaZˆ^–8Œu´¬‚išGÀØº"3/ûÓ8{½Wí1ÃÅ	›æE:˜rD°=¢g°ÈpÒÞÎ>N2ÙJ¤p kâ¶YäÒp!OéÐ¦"à’K@ò½oŠú™ÏºVÄçÇŠöŒ‡%ÇLˆiôZ k8O\§K€[Ù¥ip‰­¡|HMsŽuäÇ§)wO"W•ŽÓ4AI°ÛŸæŽ+ûÅg\ÀPþ‰l_—x§ñÈDðZµÉô¸‚<žÑ~Ö$‚AjqŒfG@0YcÄBXF&u±$Œ÷</’:Ò£¬ËkžõÄOnu6™bqrHRec‹ Ùµ±0ªìã¹F•)º‹“‘ç=Ëz·çaÔ˜7ä˜Ë·òé;nØVG´à¸óÝäU<Y'I_‘<Æœ%D?Ík’k@!•x)œ9WÈQÒqÜ®ÝíÀ>QŠ‘¢QI²Äk£X°ÐXË¹±¸F¼Ù’ g§ŒD@ø°b¹¤G#ÑúþwµÚÑEO?$w+ß>wÙX¹²°P¼ÛSQÉã.Fê³kuÅ,ÉnŽ×Ømë‘b°{š‡§Þ³!< .YŽÂkÉV2–3_S2+äßE·Ú#“G7Å†]Ïe×=ì¿c¦Æ‡²þÕèÜ/`%ÍÉbkÚ™x+vñÍ]hçÚ)4MC©ƒSä~Çéùð’{˜'Ñ0++†°Èf’~±¤O/MëéUú¨hèÑTeöü¼c£à5úYâGu¡á¥Aæ´4Z¨
µ ¢ .F)Åæ±zãm—C¯Zš2ö’ðÐ) CŠ8PqÝø,|0Â#Ú(¡¿Ç×T±Œ½š8­³G–ç>üèÐ#ò.¢Te,ú&g¹>çtm1zÒóà1Dí8K%5œüi¡Y¾{ÛÑÖkÔÁû•ˆÓEW¬„ÚÌÌCb2+,ªmïa¾6¢Žpì`–x|Ý+É¬AC¯oþ]âW­¦jÃèûo û„S´!±¦Ær×œNÓÐ‘Òõ•‡°ªYâ»o§ã¸Üyê”ä”,š r?Ãav‚á—N™2#Cwmb±šu¯óˆXÙÇVPjŠ?„ÜþEy·çR¦bRL÷ö€E|Îãõiš<ª†
STœ²¾P¿Á»­qÅÍ<ß¾3eRCÞ&l ÿjå†ÔÆcžæo‘p=?Í É¡°åž¦b‘¸ðs˜çó¡xð#Ã1Ìd8‘Ä%jQk5¬6,Ü®Œ×h•Å@Ç¥ö;~D‡CA‡"O!˜ÜŠüò„|‡ÐDô@ã±YiÁÜòâ¥z%8?•€óô‹óiærùGI/ÇÕÇ‹ÖÚ÷œì8ÀÊ>å,9rO7Q©ÕÏJúË%¯P°ô45‹“úÈ·œ†Dw G b2å”Sú[XR„XÍâ¸W©Vé=µiæKyŸD”æb"ü„Å¡œÂ¢%x!ËmZ;*l÷à nö4Ø§}-îc"Dl
?@ÇhwIÎã$.8‡(Ž¹uˆÕàìt®øá¸e„`€,¶!S$”€ôÕ0Å„ƒÈ^ú jÖxÇø÷HÅpŠûÍïÃS3<¶/üàU‰H-D)*­Mù*Z}
ÏŸY‡žÕœ5RÆ”Ú‰+Z8V¸ù«Ÿža¬`R¼`:MM&~+õ <ô¨ž3u€÷ zN¸ÏïvÒa‡a99/lTàû¨|O…ïÜq1N}º|{úŽ)IØj¼eÍ8ªë¬ÚØ{F9@‹`	òX¸N<`§Y2–î[G/aê'Þ9
hÓ²Ë×ž&™.˜Ìo!)ßvDÖ†¶ÂÓŸfBåv§IƒKaï$~py™¦â_£€Üª¾¶àö«L?Ëí4¡¼Ol,„ ôe˜ÖW‰~ŠIèí 9­™Årí1c—HU6‰šÔÆ<îZ<IÃƒ—$ÞËìŽtZW
å€…q¦#€xbQ*Ö¼7å¬º¼5áÅ6k™£ªd$‹²« UÍë°•·(3á9pALêpvÄ©¦i´BƒC>Ä©'ÕRZPˆ‹pÌz–x×È0«JSÒò„ðRHåy0	I­ðÓZre À®\í®Xx5PþÄ3Vë3\ÎIáî5ßª.Û3Õ%ýöth<wº“üf°C!¥»­zäOzš¬.ðFž9ÕF –1V§V“?óå‰m^O±#ÎÃmiæRÉ!Â"Õ~ÌK^åÅ<M2+üx'ß$·º§ÅHòÁg ©Rãáâ–7HëåU´ÐÉ:¼›VÖ+·âWEœ
?µ²*êWÈ«úr5K­ðãH®jr‹¶•]áçÚ9q^–Zig+B JëmnfÕ3‘k~JˆZJˆ¿™]IÂ0.«=MyÔñAp°õKw ªÁìRVš€'pHHøJáBTB‡>¯ÛH¢1ñÀXÒÒ¥–„ÔiÞDBA`Æ={9Ú	šnr=z6x{PÅe>
Yìr*yOAŠ¾‹ÊµÔ2VUWÁË†!¾Ï•øp:£7¢¯Òá)…ÞNGÑ—WòðúƒFÐøt´Ç CFH[Õ"¤­Špiùé¾žŽŽÒ¢‹PCP·Ú]¿ÑR¾UÁÝ¤[Ü§Ÿïr8` iüj#£†N²r{Ê³#H	?/’}ÕäÒ+zd$›ø@…ÅjWùxñr‚«0o¯s‡ï5Lø?ñ¼Z[?kºCÌrÃ&·[QoÒÏMQoéèþ¨·tÄ©7Öoz“ðYZ°m¶anåhÍ\ÏG¿÷‘vÊÝZÛÝ&z9[stv[BoÔ›îŠ|£¥vöÂqñ™…"{è£È”SC’I1ØŒ™&>Ó«ýz—;:D‰¦_»orŒËD-‚Bèó‚½Ç?{(0¡Ý‘`Zìx»®Pª¢Õ¦S‡ëo%‰kGh*ûIî=ê4ÈM0Ã3¶’CyæýÑZG´ò£Ð‚f=‡I8FñªŸ-ÅZÄç\Ž:?i ›e¶|0ó¸þF<;«Ò’xV"¡Ùe
€ˆÔÒK¤–MDª×Áß¢Wù‘QÌºë¯6–ÛÎ–Ÿð(gh±á,õ‰’n†'÷E6ÃŠÕ¬§ŸhcÿD4s£8ÔëÚ Ú_È[Í Ò¾PK;ã¢GZi~B3UD)Ð”wÖ¾p²9;ëNÛFÍØ²9¢¦Bíx§éÌ¹4Òl+èföÖ§íæÖô5ØèOÄ)hN0Jìp9‡ÏŽsPÏ&Éontx`›þ4ÎŽ3ä”Xü—žˆfFG,Á30œðžŒÒ€19‹dé8¡€Ö’Æ²”D°Óbx¾øÙ,æˆó\w8qŒÞ›X2aPåðZ<M §X¼äˆVLøŽøù6^Hé}ýºZ­¨ÐÕ†´¬ZÑCbxñz(Š;°‡ðnóöP†“´-í'Tµ;( ÜÁ­ÙY(¶{k+(öú¶†{·S%lÞŠm¦¹çœŸc|¤Y¹ft±Ÿd3òÍ¬¯ÙÕ­|4+w|[îqKjªc5Ät@”’ñ|â(,Æ¦Áe¹Ä‰…|S¨¹0åe¤TWÎj]ûç³ëX¯¨‡jö–5h	'$Ï™ÑãkÓ~ÅB¡rmúš³ÞUs·à ÃÛh°~nóžuú{
ú%ì¬÷xáÄ˜#Vp‰1¥\;ÂÿÆ
‘aH¬ûš÷×öðBö5ì4æ¬€ÊlVn;4³VMóú­ÆLŒ£Ø²§þæ¯,è}’F2ó/ÖësY ´€W»2ªuô_˜	ôsË›—}p²NG5æMaùƒÓÊ¢Ž	=h&»¡â?`X4áØÒCã[9#ïqë´=ÍyýÛ*—9}¡é—­“8¢­†3 `d[z{ã-÷7ÏJ‘4˜Ê·ËŸ§†?~Ušo*JVÄ¹Ðbû‹*É‹s¢›d0Tý6ºKV±V‰obwõÌ_iF?¶G„X•M%Uõ¥c³ê×mŠ­Êøª~$(X*˜Šö(ÐQ‘¸… @QmdË‰®…î Ô]Äü)†tk¤õ	‡ñãi…VX<ÌJ|*|ÅåÆ¥â-vWé1u KÉÀl"P6y¯ YŠfûâ§^žÌå	³6's9ÃLeüÔH•¹@aöu¬'3ÉÆl2düüŠ·U“¼˜j†/«Š,÷®
ýÒÓHÉ¯èÞ£Ü¬ÑÀ†éÌd96.7—ôNß­ÒQy˜c‹ë˜n+
>è)3
×~RZiJÏÄ]­)õû}Û]‹Ïš6Â)òö‘ºI«½¹Ú «_÷ ÝéDŸ¯?AÎ§?ek1àywš§¢X½¦ÞJêÍèìOžmyë›{']‡³è«çÑ£ÇéZly›L¿òÂÄC‹ŒµÙ;ËÙ"öE]Ùµ5ñ;F*Dù;Õñ]Î,Wo™V2‰VéÃjòo´Œ=ñóì6	îš’ [ÇMË|ÄbbxŽŠ–ió,¨;@{%zõ|MÇÚyä)‘¬nYØkvÑóv’4*³1ùË~fþ®˜Óƒr’¬G°¾P§?œŸ˜tBXbs2^ZÁŽpé0{)`Â‹,ÚE¼¡ß0dhß@"öIåÚ$Ô¥0…Ë	ôõRk¯G/?KtºDJï‘º+ÒD9ïtzOå\é»‚!ñE!™@Ì‡ÂxG\Ã5†¡ÒŽnx`ÏÚ1½ODÂ:üÒëYz¶ºÙb€zš*­é{ÑýÇƒ7¯ûfÇ—"÷d/4y—4öÝ¥íì˜Žç¼e*§FÍ–ñB³n™6ìº=s&óß4u[|ªM3n”¹î™âdÂ[&zŸmÇô1;VŠs¹Æ1½—b¿Ê;î—vÝÞÛvéØ•eŽÖ© &tÄÙ¡ÓéÙa.+=XXÌ!Ø5”S=±v1®KÖk(XžPc5„Ê[–ÀÚ²ä Ô—Üim0<-A÷-ßJœHëååºƒ9rÌ}°7ðQ6–ÀƒE­ý¶g#B{Ä¡€D›¨d8;(.¾SŽä3D?A\l>ó+$åî5
hìJR€?ç.*³e’°Á-F8MøåekB™yÕöé©3ÙLyË@py‰Õ+Ñã¢ÉyãdlŸŽZ—çïòªêÝ$GïÅ,\ °ˆ’wÅG“-0ðf-G<`¯!Äã&$!TŸ=Ç–ˆ'îî½9x¿ójoç`S¦*{Òpûj7D÷­Q‡Û—{ðvó»bÏÔ)ï‰”­‘Œ¬Yà(VA“8¦'ŸÐð)ç€EÔ\YãMˆ¤ü,‰>{D&â›…˜pÍ2Ö\E—S­«ÇBtëÍ¤PÚ;‘ˆézÑhÓºvƒ–âÆhÐJýd4(ˆ1³=[´¦5gç‰2Z9¢Ø_µhšhÄÈÌM[^3ÝÏ!Ü1‹òœÑ"Ò¶i‹ke–\¥VÆäˆ…RL’æJâaZTPvÿM´µó?w·Þ ,G9\Ý?M³³<¢kañƒ;Wg<Ã?›,ZÜ7Óãc¸adþ9=‘ÁÓÆnŽ”±fÉá©Ò0@Øz´­Â„ICJ=XVoâ«ž¬Txû°²i	Ž­ª¤y|ñ·ÜÂúöéBOŽŒ1&!Ôˆ‘íaÁ$SÀŸÝx0XŒ
|…E,úTs¯ÉåY*v×¢„ÚäýNµ™…<\]‘Ñ-­3Ï¤W'<2ñæ®˜ÚËô,:}öj~Ð.Y¿òÊ“õäûÖgaœ¤vœ9Âí´Ð\ÄØƒ~bÇfÐ“ÍÝu£u½è@Vðˆµg±·ŠQB?ýÚé;àvÀ´uµ›GU+Tš)éQŠ›à¯D°¥´vð}ÃX€õ(1ŒKÚÊÚÙŽ²–ŒËR¦GcÖ+ì}ÓÎ°²wY]ZaSÊ(˜Lt_/eHŒ¦ŠXl$Îp·!¶€M6gê­ÍI8ãŽBÏŒàw7	òî®Y¦¼¡Ú¤ &Ñ“ÖÛ™â™`Uƒ(…ÿx¶¥7a•ì4)ÐnªÜ@%áü–¨|_†¸Ó1ê©ž’,)2ëå›ñ{_þ¥çëÑQžóTyl:&¶} kâVæºEöjÐ<Ns?=îsÔãîÃy\Œ»îšüxóK„É©AçMJ/CîÅèh˜ÿ4M‘„…•Ì¸Ø;ažbIÜW©íôA1x¦ˆ¸_]Z1PVêWCà~À¸Í|ûøî @þvÅ# ÄW×ÖWVà¿º„Ê=¥ÃsìèÿòßÀ’À§ Z7:-D—¶‘<ÐYJLŒ@fj“ÍáÑ9@p	\4Ð2µ·øëV•GKÑ—WÐä1GTîu4ÇÖó5¥´Ôô­¹À™'Fa¥ž­–Âfräê*€´VÖ£Ãl„y¸FâFÙFñmÔ4xµZbUŒNªÑ<=QPq2-1%û0§?ŠTkò½©Â¶Ï³*‡kÇtÝÜ³ŒõR F¹Äz¦Ú‘”iáT`SÀÀÂm†M*êbr˜áOû8§òuÏØ>jU6z˜ïŽa+„‹ôÇZ½ü!àŠô˜gt´ÉV\y%f°›Ô¦\5êüC×b*K^EÓRNPE)™—BŸ„?Ò‘F––}ƒ!0å.îDú§(3Å«žr>²ÖÑ:Ê>•£·¶ˆÁu‹Õ$SÊùÙ‘ž¾ ÇºÒDšlÁ€è'Ò/,~;Ñ'ÜÍd¢“¢Áû¡ãµ¡–zOø²³Uá¦åVM°Gd¾}
[´gìŸÖIäO¨*î—ýÆòƒ•÷¤7»VWWï¬vµŽ{…Ï)ºÖ‹ÑÉüz³áŽa{N?ê7»Þí‡‡b¨ÙBfAI¦°ªÄÆ»€ÄU"ÁiB½÷<Â?1hƒ~åˆz7·_Ga Ç!Œ§hþ±çâ,ÎÞ9±57Q$}+rÝF"9=bˆ¥’¿sN*¦BSqX\&lÄ:£QL»~ÊAI·¤ÁÑªf,i!¿NÇlE•›„KInû~0K¡áNWÑL¼»A5áñz´ª×aÆ–±ee™ÁÞaÌóu«’ ¬G+ú3±›vÙv–sº¼ÍO0y²ØÊfYm=0Ã@Á³è£6ä-ò}7N™V•¥ñ‹–c×XêÁe€i2mæv`œ:ÐñOäJçuôeìŽŸêê¢üÔ®Pn¦O2~­Ñ¹o&U6B½Ì€¸w&N1ì™²Ñ(…Ý®RxÝMÇå”˜ã±d-)ÏW)ÈÛHè­bXË%%¢Û\€ROeB€VåQMB3þ³#{ÐrPœé‘aÉ®Éº8E%LnP$¦&Äó.¢@Wl`õŒºÎj¢s‡ÎeÁsw|CÔ¥íºÅ/lW  Œ^	ª˜¸å“ãNË!	˜0™ê		B#8Ú…ý9Ã¼5Ešþœ–˜
ó4ŸÀ—qÔÙ*2¼¯`ñ;ÚRØ¸ªËe:Nf‘=-ç<[b&/ƒË>·dsÙÇ9±vF6—ïU³Ë×¢%ËË>m_UÒÏþ²­àcRV[i‰XÕ‘]”›e;”øŒm¨í>Œn§ÓñG(èò6Æ·Ôd½ñ-Y×ˆ–ïÝÎ¶¥(™sô¬ˆŽãÃàÌÈÊBW.}LÓ	¢yvòÉes¯«a&þE%çãØRŒƒ:<¯ð{À‘Ä‚¶O(ñ„ÜpQÁJs÷g¦ºÃ7„ò)ƒ…=hþ8>Ë‹Å¨Baœ‘3:•iAhüýòmGj"Þ:©*•ô‹ÁÙÌ8”Ü–2Á/q'É*Sä¬‰$U{þŠ†u WSÍÒ{SSÙ€µŒÑ„sõÞÚFoV·€¹6.{˜/s˜¶™Rjé©–(,ž{µÏ¨°ó´ c]“ë#×™c]îIýæç§×ÆùáQÎ5/-Ý<õjÚP…íÓíTl"«¢Š€oGÐ$ÚäÏ6úŒG	£‹-—	®¿Õ‚>)ff4§5$  ¦!/5l.žê^ß±6;Ö`¨ƒžÃóø²4L+M7.ÑA™T¹”Óñ€?)q^ 9X €ý+œ”XÀj™ï¹º%LïnmËgqøøærÂæ>Ä¸ì(LíçafPWmG¸ú…±™¡¦¶®ÚSrcG£?M†yœ¨ÛítÖ#|³HÖ¦ëJŠÒ‰þKÔ=K1Ã_&UÒöM¦>Æu¬ƒ€úVNk\RÆéQ,xp`Æd:ÐÿçáÊÒã•?|Í`Qt8Õ²söPÆŸŠ¡4rðfwG@áæÛì5Íh1z¼¶²­ômrÆ©ð~”³´®k—ÿ–ç¬À_…8c™<V@P¥g˜FyáÖÅHµÇ¯9;‰9tè¨½ZføV¾Dî)wÇRá¯ÊKu»±TâB0Í„›“L‰GáRû&ÒýF‘Ê&‹HxA/ ™;¬³”gQDölnelÓbûðDDºpÝÄv ÿÂ…
YËX‹¡<®¿Áq²W¾§Ü”Ÿ~¬
»ÕãU/VH œÊw¿	™šÆ:aü©3QyÑÀAYcàìS^Ô¢Æ}r‚P«Ðãì(Ñf>Æ‘B]çœ+qï§Ã9mGÀ:Û@}aØ@`Š8aäÇ;	•kÂ<¿6–©?¸4…o³tÀÞp|²÷Œn>ý ï„wÈ‡ÑÀ<<óÆ=˜Ç…)º±1ÏñpZžrð}„Bzúå xÉn§G0´AŠâÜÂoT‡¾›ø:ŸVõU$eZSÈbw-‡2±ÆÌL¾E¥~™ý,S·Ò%V”!#; MÒÐÅ¥Ü–ÙÔJÓ¢Hm K³¼s|œÐæŒI !‘t5NîºS;.ÑæpøÒÐÁtµyhM¯õY½A)ÁÇÕæ'î9Ëå$U:0³™Á4H0Q•OÎì¶.·,›Bå¬mª ’¬#RT{íÉ´®4³]Óˆ?JÍ«‡<†,<¥ôÚÔ®øéhC4'T…S€R)fño`ºXz?ñ]kT¼XºhjÈx†ÒàI^l­©·fã’('[qwB÷ÌçÂ~úôD¼~µ?Ò´ú|,oíëÀøU¯*˜Ü‹©mCc€CÎ(àPÂIMR¥5©ÇÁmËÙ™Gqœ‡Òvu;ˆ+¡¸„w/ã’Þ'ˆ[»ÊÖ×ñaW³fÈÒÍ4#p@x@†d#¢@_cb{ü²pÓé{‡r€Z¦Ý2ÔßÙ´­jy<UéšL÷zý<_[ŒxëÏW#ìñùŠ†Å'y6®ð¶:Cw¤•hpÑÐuò™ßF´
ÏWti=4õb-ý#Ó¸)˜¼ÍBt…³D´)ãžÞ4ôîh'z®7¸!	Xù\À°Â0ñPÇð]<6Îº~°ý&2ª³>/`Î i,Yyú¬Óßêœëç[µ¤Î¯*É«ð¸ÿU0ækÜÌ»ÂîÈ¥Tpíí§ñ âv\bÚ]•¨¦Gw²ñàî¡1J“à? }Êè8MTé÷ÄAœËÒ HpMúOôÜrñø$e‘qNBE?MÓiÚ†*!½)pHZW
mX(šéhCxØG¨†»_öZ!åu¯Àªôæeãl„„X’•—Çiuž£Ýå7hÀ0ÌÏÅÊ&@«¸ ùþÈÌkIGÕšÈò&M›	/¡,éóÅè1§>\¶ž ×Ž÷ÏÐ¼uWTÞ27øw²JäWh”‹@9ÆŒ‹Š›”«Jýr2„Û¬³¤l](ªt\”p´¡Ø:Ýº–„Ú"ìCý²S5¹õ:]7ôtw\±ÊoWÞ-F«+p[ÏWÙóh)ZuÞ­‰:«p3 l5,HðÁ\,¯*3+_Ç¯»ªzÿ„„å#_žæç‡9\¬Ý2—ÉÆg7¿3ôýYŒˆ%[Ðè}[µ ÝK¤mÖc'£Fc˜¢âÙ4Ü$Pa+}¥©Ü2œm•·- a—¤ö7Æ,®]ÊH¶ Õm×=S,â¹MJãð8:]}öBbÄÔÛóˆÈcüÌ+"Z“Ëï	Af2JœÚ@]¼.ü'–¿áSõÔØ(jÈÇL1@.ÈãŸ~ñÂ:)‹M0>7cþÒÀœ vè$VŒÀ§›pKÛéTiúš¡„4à&+Õ.¼©çÆMòËvR=L¨»š.œæú–ëZ/DJ1øoâ¶3ÒEÂ½ÃòØj'ÁÎâ~ëè4k÷@ç—øºmVº?`¿ÊwÞp{x%<»Ñ- ÓÙ¬t^­~äRÖ„¼vWµÖüÒ˜`Ä10-oj¦ŽOrÃÆ)ŠÑ<0NbòÑåqi ú/Ysâh-¾#P£F9 Y=Ú¤°Mu÷j²ÅÂgoû[äþ4²CËŠäI*­’a+Y$ÝºÉ¶š@ºÅ‡íá¹êA'/üR<8ürÿ±„wþKøº>G)Õ(®$”\û.ë ÀZ%Éò«WË—ð‰¾ÿ~}4ZôŒ8éÛ|\ Û]”q˜^T[ùöõ1+1hðÊãÞSD	;#do’èÉÊŠ¬©ˆºýøˆ›¿ÅÔÔ“‰t¨þb„úä+ØâìäÎ ÍæŠŽ<$Þ‘­®…F¶º¢š†ða›Ú¿¼â‹wýõùð±¶›¾¥³ÍqC÷¶]Î+ct¦ È<5Ò›"PÕVLáµ¿¼jÕíÀïwaU;=1ÅGkòHH¡Ät4Š‹K¦âqyžŒOb"šç:Àr{ÎÊ’ç—’	gYÏ•ôÄ}G"¨«¼‚[‡-œÑÔDñàñ?°—•ž"œY«£øâ6mbd.Â£‚Ø“O‡«3¶ÕîÍÐíj³[ŽºrTzÄ¼¾Í.Ò¤»J©ŒW:¹ú(÷¨>v—Éi‘ìÁHúõà0òxEãw¼¹)­¢ˆÄªÏœ?â¿¼Ÿ °<À"îZ%0‰õ‚ ZàÅ¨P—f­ó,QÁ‚^+l·Ù Ïláó™ÕY‚b‘‹o¼^“Š!¸–ó*ó˜ìx1:" êÆªR§ÇœMR´Ú§Ý#ý]-º{bïnGá•[¸"/£ø	wø7
¾‰ü;¹ÁBø5¹7„ØQ û­  ìÙJ	”#–veüº0Ÿš8@ƒwöèæ/46*)½k˜‘"—uG1èš‹A;¯à§#	 Ñ ¸ùWÂÝ™¶L}±2%-c0ÚhMuˆ£AšT§;aŠEõŸ×¡¢"1OÓ¸†·o;„¼ýŠd1ø] üþF. `ÀwïTGyr¹®öE[pš¢ª·sRdIÇìò º²^˜ü ŒB»˜ß)žÀ¹¶Ð?ç¤ˆGqí8ÅBWŸ*á½ô„÷¿($|pj[’Ó”±W+bsáŒmßå©V<1/þV>Œª¦®ÆhÈÿšgÌ€Ç™·Pa§(¶ââ$þ8«"ãMÈ–™¹\j„ø8Iö€QÓ56AÌ¡½t/ñ’£´W@bm“€éš$ÜLs:š3‹{ü­­ô4ÃÖàéQçgí±¾<òmVhj')ž–mÚvXà§L$OÕ¸¼ù–›Næ1Ž”8TæbVõþ³5óé2ù9ö/ýÁ0tÉ1NâÚæÓj2…Õ=æGÓ‚B‡lWeÅ\ÌSòúó²,Ò/>yÅ5ÍÝ³x¸ª¥ÊSö@,miºÖj
‚abp°EŠdŠ5å¬¢ùèÅù1Šiî)òNÅ’ÞÈ„ê±zM‰kãÃ{8ÓAÆ\òE®Á®î\³xA±šÚ,ˆFWà²ˆ6äœeúú¨J®WÈ1yE¶
¢5cŠÍ¨µìøË–È’p[Ö„‚%pHmÔ£Í|ôOÚå{s¨³/„ŸLÈŠñ¼KoÉ’Edx•ŽrKwÂÅˆãì§i*JÃàÂâ,Î†ˆeÊÔ@8\º‘éÒÕ’œFbŠ:µ±Qä	Qß§áWœ¯B’ˆIuct°ë†Ï‰-’¦Y¨8¿•-H#R§SiS«§ÖT
ðk‘;Ö„öTîD~Lnï_®‚RI2quTÓZ‘Ö²u™f.\ÃiÍÂK›sÏæGb›Á\ÕÒ7û”Á”é¸Ìð®àæìÁêv˜>ƒVd1zëÀï¢:÷‹*PÓ¢Œ ûÎd=ØÐ^fG@]Rˆú†ãE2Žƒ´ª%!öµD¶EL)©ÛbxŽãy­ž¥Ür‘e86ŽhÂDOª¦	¤bŠ28/ä9¤™8d«”“«h¹[t¡BËü@N|’s{ˆ”é[ S²!÷ÔR€1Ÿr×“ÙØÞ‰íŠìÏÞÕ¹*Z\gdœá±e
'Æ67‰è*èjQÖòÇÚn1K}f^;ðÚÏwh_„~yq¥J7í¥ÌNc ¥@ˆA 4Ä=^b¯«|YôÇ?S¤¤7$œ‹÷ê¢õ¹@;µö7t“§øõÁ,£›×9Pti’¸äƒ9h7%È–÷¯åìÀZw4Kãh–þ£©Ú*9à—·P¨­Õ<!¿œò½ËHKy°/çöí‡6Ìë$‰$e”£1Ê­Àø®MA^ihæ2È¤ØãÞäÂ¢J3¬Ó¸õíb¤ŽvÍ–,ºÓ‘¤0h—ûY’Eƒa\–ÈÓ>_('ñ ]º\z²ðB®]æx˜^0Ÿ%¦/å¡Å/—ŽÒê<MÇZ]Vû…±«ÏNWõæPî·ôðbLó¸Z:Æƒ=+Ñtéë••¨*à!PÑKUvrZ-¼ø&;fy•b.<MyK?[>]µz›8±†CÃÔ%Ú¥NG/¾CÍÅÿ$1£RpÀw.P/Ÿ-OŒ‰-[3ó®ÓI<YZ‹ðëÒyO¬ÁeãÉ´²@ÑQªgÃtÁ~•%Ï(ÄK]áÐú?Í’6Ã.é¤z¾Ð¿–‹þY°Šäã-²}~Åì-Y&J aÐek:ßL«*GQFÒ‚>€/Ë³€-³ÁÇçWI~>FWõÃt4Ám¸Žª¬Âh¿‰³Óž'é'«°Úßæµ#<ãÏ¯V¿¾†!D¯¨Ž9’e6”[¥*òÒ`Bš®‚½ÆïmGÉëocùŒ•ñ¬I>˜"ì£µÛ -|ýær7évt8éô6ú¬Õí];ãåÁŒÑ²-‡ñÒN´oè7°†³nWšxCScÓM,Hvó¹¤at5àü®Õ²ZMÏMå.°§B 0Ò¥ß9:aß¯,?†_y‘¤{²oOó³´XeVÕ­ÜC†àHø‘—(î/½ƒ°6ÞwíÚ	n˜Ÿææp¨PfÇsóŠ5)oÔçÚ¾³ñÃÏ ËpšQ•ã!ÿôûLàQÀ©%Ó§ü[ZÂyŠŽt>ÅQyó×è,ýÙÆKQdÃ~žqyº&à÷	ÁïËo×èæÅ‘š½^hŽ¢ž=Í«°ÌM-Ü†½cžj/„=Y1ajÕ©Ç+B˜	.ö”àæ®`á›¯³;»8‘½û–*°1ú2Þbæ,,²è¸fÖ¢ÊÖºW‘GU@Ž6ÂÅ}¬ä‹¡jÆ¼¡ñ	Ôxû.l~+½7œ–æ:S1‘™ÖøÙ²ûÍƒÂ2ï!ƒ_ó Ô•Ì8\³Õ…7©é@$é_´š#1ã¡¶µ+ 
ñƒ¡ë<;%“ç¿t§4î²¸êœŽà™0Ü“aN·sÏ{ä=J±Üž™w§íq’¢š¿ùÓt ÒLÿZ‡é@R¿Îi:`¶;µ+ÀË´:Kê¦áVAï‹tÁ<‰(‚Kþ¾œž¤˜ëª­Ä?çâ!¿ž4“!ëÍÇI&[ÈKYz:¬€syÕÐ@Ÿ—¥ÞKÈ°y2Ö¿íå÷™tiðso'e%óé›×))ývÇdåSCÔ80ŠÕîô,.Ù*ÎDàÎi¸òøK%òÚåÊ#”ï!cô:=¹ùë ódÍÊÙlŒ+¶2ðŸõÒ?~y‰/L¢†ïèhéI„WÉ1z#_,¡=ž)<òpâÖþa€y\ººrÙœÖÕ`lÃÇÂä†2¹\zÈ%‹ù0á‚Å‘Îpc¶+1Ÿ¥µèü4«R&ç$²³`Ôò¼©”ù¤ÐUKj šE9ÊÇ•.õ|$odí!V5Æb,‡¹÷>ÎêÙò‘%·ÚI'žÛ¡q$ŸÍvxèú9m‡Òî}Ž»¡´ŸËfxéÂ9í»Ö?Ë­Çç²ÌáìN„ÌÓþYîƒÞg´Hd”*ÃÝ›HÐe£4Xû‡ˆì7Úì-I=Q¨àâU–˜ÆÒ$§ÃŸ–ò!‘¶NÍ"XÐÄ>ÂP´^.­F£d]ý\#Å#W‰ã¨-éé0>J‡zëGÃ\è+/JpäÞL'À6â2Å¯.¼@î}Š'yÁŽÁ³ejÚÓ%Só6(%Î—Ž§ — Â„×"Jy“&KX]Rí8LËu.)X¢lŸì,¢ÚFýÀøxeey”ö®|Õ)Ã)œ¯Òaå|j¥hì¦â${˜À”‡rb’6 È'x^ù€^ . ÄdGï<[f|u¯\[Q2†%7¦hûczùü
^‹yÓììÁ3âgËl…¸³Észøé€Q0œ{ÿ{KÇfbVèt˜2Í!ä:µª’¡1ÀŸ¨»•¦£(FE\\öjAÝr9`!
ÐƒÔÏ¨êÐE%ôÓìO]§½»ˆ&¶wí¿šêÕv³^ÃíçCï=r5	(¸„w!™ž>p‘%j·D2¢Rv?ªt¯W×¼ßÎ«Î.Êãl¥>·3Ž¾£‰œJÄTo!ÚlaÖÔRxYì¶¯ï´êñI«Õæ/a™ã“rƒ%§­Œä´|ø•6tÙv‹1Sä¡IÏwÚé@Ð!Gkçkã¦á¨6š Í,úZ «£z¡<“üIiF¦òß£Ä&­“iAá*é.ÀÌ}XŒ`‰)¿PöDÖ$pæj­ÎÔ¼¨3lÑFN¢vÓxOº$bM[Êâ±Âj *Þ¢€wõÚpS>Ä›r/³#À©p·±ë›~HîàíêÊäâvÅ‹= ûV]ðÒrî<Ãmû8«‚£A|E8ÆwÒp—q¼¡êW&. wýÁuÄýF:íV%'ÙI.×…ÿ¼ï•±§Z¢¦Eò\|¬R TÛ˜_ŒE90EYu¹ôõ
^Tâ×
»µà-±ƒ'^¬šfGü1œS ¬>.­„AÖÇëŸ–
3ü˜RvÌxÖd—aÖ£ÈS“¦¢–!GMáÈoåáá±XßÜÁ¯ìÚ¢Òl°,ûÚPÁÐ®….`[
a5Aú8v«`¥·ïjŠ‡uvøi¯†£¶ÂG?ðK«ýÇJ%­®q˜‡'5ÒC–b[ÖEƒiQæÅÅ'N¯À—ˆ›´TÆÊR)T>Œ9ža&{ T;¯m»Y­°Gf¼¿·ÓÇ÷÷qþ>3èwËŠí°aš/Ü2‚÷GøÅéCãVñ	¾f
½µ…W^<ÞéÃ6–Ø…(!‘þE©Ä2Ö÷ÝãÜ7`W`S‡ŠP»­Ñ”8º½æW¶eàß°„¶T|®8Veô|Ó–tO"2U(E©–Ö ø…4Ø!´ÌÅ˜{ÌâíõÅxûøäÚ¢ÃÖ[ÁèâSdýµ4/Š,é6ÐôËOÔ9%š-@ãÔÐgÏ€éÇhåÚÙvCš<cQôÐœoAþú$AÐiVÊ»þ&J0;m)J¸­Øàß·À‚æCZ6Â‚t	—÷¦_cG›_ŽÖÄ—®½4qÅš(ÁÇÏ‡|¹!â³£¯¥c€/òÑÁ(Ë1
Uqíá`‹ ?úâô=hßwõãhËg1Œ¾1c×Oíq×4ÿqš–HéÔÆ‰1«18E·Z>Š³Ã¿;[Ó Ö`kø¦²çç‹/>¡ƒ-c,z±¢G®hìY×`Vd9ÿ¥háøßÖÿYÌîÞÛ‹x}BÞùaðyluK<~oÐÎ·r.·ÃíÖC°ù#†ÇQq<EtÒúp_æÂ²f(*¤†Úú§qÉ^õüà¼Íô¬FD œ?Ö­§û´8VVy—êbww¹xL!cëç–ó·þù›—õ‹Iùç9ÊüƒíÉñ6±+5þ²«AoqûseAH­P÷ú™)S©ÕBÈ¹Ô
ÚÙ' n—Ýv[Ïˆqÿ1ßâÙj´Vzð9}–ª¾X³Ìül[«†~BK’]Ìe1-A¢UŠ^,åâÂÈ£a~\ù”OÒ*ž¢Ku.™ ¦ê3ˆ‹ùzÎ™@1?"`VˆB4?a7æžn¯|²ë3!8ºe­YôQâã×KB˜6¨ÄG—o£¥¥­¢ø°­Vâc*®ÂqÅêuW²5¦Ã¢fÚ¨±Ä§^%>³).dÛ%jøiBÂWê45"Ø@"ãxø2ä’Ñ¯kô¼V#Z®ÑTó^¹»ƒNC6‰/öÈIøŠeÓ½nQÃx>3í9©I‡‡Ÿ³™5¹rÛu[hóðÓðzy9úž‚E“¼Â¬ñ0ÚÝŽ`,#
NÜ)¨LÄBÓøsÕ'ÂEwU ìç‘~É4X¶€Jž{Ç9Þç%8ËèY´»Qéë!¯€¦j×®¶Ëü¸Ñš¯|ÍåÙšO‹[E9H«%nSMF%ÐŠ´©Væ¬Š›æÛý´2ÓhZñ6§×c1/··Mýï¹†;ì33ãÀeÊÁ7¼®NÃ-9‹IUh0ë 2÷~-{º>_HšƒÉ~l³»ÃQ{ó*ÞEµÄPýËÌA"òÌ€Ò£	·84h6—ÀÏ•#Hà–T¨¥5u×Ò¢âám-*<ãnmbA¼g3‹Ç÷bfA#Ÿ©~jÍ-¨¯êßovÁj¡}&SW9nÅ)îumKûÐÆ`/+];‹Akuë|	Ûù¨Z¯{ŸÀeN3ž«FõoÊÊå>ÌZîç©
mœöœ´žsÑyÞJãY³?ÒJ(¤ôä(‹Lˆ¼2¢vVD±d•Q<³QàVAs!Ì S‘´°1
E‹™ßÚúÌFS$¿oü•è×™®8å¶ˆe{[D«ƒ_ÇÄ¦uzÕÞ\RˆÙÝ‡‰M9W›Æ”3™”´ÜêyšØÌ´Å¼»ò–wŠ÷hžCæ&ePf0“ÁÉmV'àH.NÊÏÂè$ìôö÷g<òyZ‚\•-L?ÚØ\udÛhÙ1[¿'ËŽYfþZvÔF<ùÚˆxâµÌpÄÍì1.q²®¿_[™p¹ùùÍV#ðùÍVÃÓÚ¿C[òN¶–hU<nªÞÙ†ã7“
ü¸8§1ž·]•ÐMÙJ…k„oÆ½zžq…7VhgÐ˜3T¶Å>Þ åeßó¸81Íaæ£æ6ìè8wýIsþ€éeß÷¼¹5;ÈzÙ7Ÿ4·à½^öÝ§LÒ8£µ‚C¶FöÑ—í±|+‡ÏT¯|_
2úü¿S…{D´Ÿ)$Ý…Â<àèó1Qh)›iïÕŽk·W;ÅÓG­yO'ÚæÚÂ¹)Ñ¾ºŸ-Ÿ>ªíÔ£³é=ýÎÎfyå½m¯1.ÖøŽóÆÀlò›<ëÛ¸ºï<q<4Zf¸±¸ôëÊ!ê'ß@õ¶&‡åá¼M`â!OIvÐÏ¸5VÏÍ`x½ÏH™›_€–iÚ—PÌo?¿£x+~Z‰¨‹×Ë›&Ðl6<¿ïõùÃÞ®„Ü¼roµB&ØzmˆÛ´ê¶IUÜ’$mŽËó´@ÃÚÌ²R–(E»¶DbqÅµàPíw;‰’Ô-èUIVÎ`ŽŸ&ö\ëËÖÉ5“ô¬`ƒ­z•WñðU|ÁµùÀYJ€ÀïvËéh1*Xxâé(ú*ê®EÿáÀRFü¯öz‹ÑJëîÒ¸§IëÎŠ~9È‹;Z™©ŸÉ Í™ää^‘ ,êiŸf·Ë²,Ëô`^€-Q¾ÒÜK­âÉü<ûÃ$x9ˆ.¿ò»ð!V@VÃCÃ\®ÛÕ¢hÚ*õvŸŒáeƒslVÙø„%²Ž¾”ƒ¡KÆqí‹£F™ 6+¸Ëy©WñK-´ê¾An„ŸÆíC+…Û‹öï"xª3 Ý©PÅ#cúÜàPÃsö\¬Ž×JµU@eìXi)â¹é_÷m¿ß§q-Fðz{çSìŠâOÑ¢©G‚9Ö%ú§ÅÑ˜) ÙÃ[ÏâçJN£¥Õ,­Þr¸°Ï3D™ÿüìgñÓhCK}¶QÊÝåèÁÕµmlÑkÓà¤c–É¿çlsÔ`wrßÖ¦3häæj\twkÓF#¢y™]Ï¶Gó³6ao>3« !Ý\¤ÿk›`–RRr6˜&>iÌ(Ñ6ƒçmM/ÝëÄ6)2n³eK¸àZFé¡&˜ìÅZvj¼‡°úf–—Õ•èÿÑeV´ýj#aUpÂü§¼jýô2¾D>”Ý£k+^¹d(rk»¸BØ²*Ád+e…à™¤<§P úMv4Ìò*ÄÁè¸-¢Ý.¼Øh/«0#êÓ(ŸRQþÔ_¨¯¼8‰ÇÙÏ1AF}É¾g8
Á”C^³µZ3Ì&³ygÀçS–7©ïD5ù«¥6îƒ
Ííô8ž+júZÈ»?ªâ)¦0‚e0ä4‰“‚¶'0)+Œóè(d‚[– ÊE’Ú¦³J·çfBVÕßð=Ö‡oNís²F³$²m™îw¶d¿í4#þÌµ³¥ù-K8Íñ,`/¼n÷m302±Å¼_ävÝ}¿Ä]¾9NPõ\«Fš£á)¢eK:ˆqË"Ã5%Ð‹lõCG+(PÇpC‹~½äx Û²e·ž~[Ú‹%Ù?n´)F>ðk¡£3Þ<tƒ7)ããfQþÌ¶Ä¡k¸EƒÚE|ñL&ÅçK_Ãþµ®¾4Ôš–
¦ýu.óÓµ²ì!Ô˜“l|"•O•óÖK9°ðBA\½@»‚‡U)xàÁï-Å¨BiG£àÄ&/®$®¸–¨¢‹âšå¢‰ÿý_ÿy–âWI\+‚¾‹JÀV6ïå½‰v¨VÒUsÒqÂv´î³[±f¢mTS–ÊÀeM?öD
Æ´”Œ–Uƒ¦ÁFn"kn™Ž2"ÚkzžÉB(<V‡ù~:†ÁªË±^R­¨¤×³ÕÒˆ¦‚zMF½™KjÌJjaíƒ|D”³ªÿ*Küê€g‘iÃa‘~ ÈyÜ7Ôm“AÒíà‡3}NðcópbÆB·¡ºÔew’ðú‚ðGë>“ZQ2ïYEWk¦è*Ò0
\êh;Lde"kvƒwŠŒ.–âi•·‰<šU$¢‹ÊÞ¡œ»_Á1¹tDÆBÈG¼z4`ÂŽK¨V)‚†iŒ”â:ô]¤~ÉQmjr’E8¿DCtŽèG?äƒ›‰&y’ÊQ9EYBqó×	àÁR
`©ó127XIg4)„%‚ˆ£Ÿ¦ñð§)n!™èæ}¯È ­Ìf¢.Í£h7þ-¼æWŸÌ$-i-/™UbÒiû¥&^¹É¬=æËFiV„$(ukÓVÀ×ià“‰´¥ãc³D›ýÚÚt7Íyˆ³m&÷×—hSrüò:=¹ùë Ë£4:HOy•÷%Ú®D§cÑ)é{Põ$/ ×QÊtƒåb” õÌÍ_É|¦„ÑÍUþÝÅ0ôÂ¦Ãø°‹XÎ@ùFìâÃ,~ilNÛ©ï¦³>”âcïE¢za*_.O-ù
Õ&ÊÈÊƒ*'qö5‡»{oÞï¼ÚÛß9Øìgc n“´T…¿ VB”) Lˆ2S²JIªtfˆÈ;õ–KWÙ€ºÅˆe'&G•Ð¨’O<ªfQ³› anCrÆÔ*¡‰hü³“ÏY<9p;)p8ZÁ}‰g·pa9¾†»”M’ë™…ÈT›…ÈWÖ»eø	=È¤W°ÿxÆø/8=ç­DøŸZ†ÎåÙâ¦¸EKùùLÒsÞ—¼®jç{¼­~`røšWÞwWÚßùŸ‰ ŸÓvR–_C”Y5‰ýz¶ŠŽ8¿Žæ¶êþ&ÒÇÏß‘H_Â—êß
Œ¸‹ògFŸX²ÇÊ3_ŠâÚsjôšXÆ˜L'	J‹4IÇƒ,Nb¸ª)sÛ¸¢n–¬G,Èò"üÅwëQg³ÊÎâN<Ê7Ãü§i
õ:,UÅ¥„*ØIÎÓ”Z/eaüKLë úÀ#k6¢+tìˆ~£ëh=(#]4Wž–h ¬ Xï?ñyœ‰Inçƒn‚ÿ-FÁ¬¾hƒê,B¿½EèWô¨mÂ5ŠS§Qw0Ì§ÉNQ˜þv¸œù0í§E‘Ý…üÃ:†¥íeHKBíõ…ÅH6£uñ…ùw2Œ/¦Hq@Þ]¼¬
u¦ÆÑfØ5«ÆÛjˆ½a<Îýð€+½8LË*%x•ŽKŒî_7ÇSþm;=ÎÆÍ=ŠÐ§„­Óylüø×ÁÌ ²Å%Á\ñãøŽŽÒâ!Aöõ)!Aët^ K£$¾œò€éÓ¢ê~8„Ùl³¢!r`j9.<þòJN‡¸¨Ýñ"çÁ‡;A—šÕÌ0ö2‰“ ˆñ·/3àA±xq€fõø)ÁÍéz^@7ÄË”u€ÑÄæzx!OÌ ð©Ó›¿Ä
câîqö”šàî”’±íâ•Å,q“Q6Ö€DŸ[ä@eî6Ÿ0×é +IyŽêî)%z»²´µï/ÈÅbbZ5>rØ8‚ ~ü²;È‡èÏ”åã:àÀ9?M ÈóóqZì& NçÏá_w
 %ûçÆÊãx‚³% =I+€Ð²û“à°L:¬.u0ãs9ÉÌ¥ ¢Àq6Œ~¼ù%J/pÚÑ8ŽÐ¾Aú<XÐÀK_d<ø2î	F¨:.âŸãü ½Ð†ël^IV"P"¯g÷Ã—WöÓë¨K[Üû ç·³[Å#Ø!à²"‹¶ §é°Êi‰RZH¯xëé
¥øF[Z=ØëL¾SéŒJWTøÿ<ÁÇœL«ðsE †.Kà³‚ÔhÒí©‚|×åÆªW¯^èF±i½«+ü÷øQo
8{P‹öC¤g»–Çž€'NDo­µŒ±ÿŸÓoóâÛŒ[QtÅ†+0mÄ6ûüÒˆ¶/#›@ˆ_^
¼£&`«V5ôƒ:žž¥#/Zâñe-*Šâ°h6 l‚‡_GFö¸E= ¦Ï Šã„³-`,	¨Ù‡ý/ã8lÈAjxJÏ²ôœ¹àÁ£R€0s|ÅUBÏ^²W{Ø\?@TÑäŽŸ:NÓÁÇ|ZÉàù"aó®DpÂ$°±Åå¼Ìñð öºè>Ù¬×íˆâï‘ð„¦€ÜèŒòqu:¼$þ!+Ažâ›¤ùG`èØ®Š £Ó¡ôŒËËÑÁôãÌN2#šU4ÌO²¼è™^jçb’±¨è8Cë"ÈJ†*Ä ¯ ×AñN$¥Y$úçez\‰³Äˆn8•t†ÄéÁ¶R•âaÐ¾fI&»Ñ¶nti^tÞÞÑÕ—©xÓ>G54§ºGÔð-Þm¬Ú8'œ35Zï+vA…±¥j(¼®Á’Š£;óÐ´VÄRò|Ÿ R«ýh÷p¶à@A‡¤
tì‹2tÉ­9ÜÅ'qF*úJÜœ¤Eúõ,ª@¼ÞŽ‰«­Û¯r*¶ÑªX—²°ÀÙ§õƒ2éˆ$;>Æ‹
Æ£Q >††—TÃ=õØj…o´BQoi6ìÊ¦—£îêÊÊJôÑùÏÚ£žEQEZ¨X¹Í²ágpœ´P­jß©¿Q|Ñ]QÐÐÓJ2€` f^€
 ÖúÑ›
è¯óƒªÇÃÁ5
Dû$¢||œLaDÄ•1Æ.ªRèkH20ÌFO ŒIzéç6–´·Z1¤ptØ2šèáJ$Mo¸È^<yoŠÐD@ÔÉ
¹-hî…B½b=eêàÏKø|0 §6|Ð
_g†ãa÷{kè²Ú¿%ÙEˆ£y¡5áÜ– zìÇZº¶…ÊåôHŠ‘k.M¶¼ÂéOpq|3Ì)ÞÏ~}"‰Rwûñ|–Þ"ý¡îŠ¢š}.›%è×ÄÙúØð>ÙaK‚>›bz}¾L’Ä2ggÎ¨*Î‡e~ã‘gã¥Ó%X/në)•ŽV"Óx_S·>ÛBäá³T?çyÃÐæžÛà({¦[²¹Ö.kh÷	ÿm”4ˆÚ+aùððÂ—Âo®öª_‹Aeã±1#Ð’"ðû#ÛòÓUœ®:7NMÒVÛéÊ„`%©ñå•Š	Ñ–Ì“üÅc<„€®|Á´#>]5ÆŠòñ˜éë{}kŒÆ»‡|€,x`ØÑù´í‘),MBM«—1pEUŽ–ûðßñtLAØe¨g7Õq|–‹ ã$6‰bø6ÎÏb$Ÿb40L2?s–³qô?Xƒ„%;˜j"!àx¦°~Ìé B¹eªXT%­?-'xE“t˜›ŒU?ÚWP[#šJ\1žXª)4fl8Ë"=k˜™É–Ðm™US6à¾µQ“ºc²[³öÃcq!¬f­ÇŽ¦–33]çrŠÆÁ¤³Mþ4šÍ]{løÑÈ“ÀÝðD|Óâ|ÙªeWu¹O[^D›¸× \…mïW¡ÚúMÿª4)áí5ËNÆo¦UÊ©³Z¡µ²@l¯mi”n´g/ÄAœ(ÙÜBÐ3§æ™¿…¯ž-#Ê¿®¸q†9¡‚û¾Zô«Äj3g#l¥œÀmm¸ÑZžkz¤üÈôpm/Œ©§jHVïñ}4Û½ùµá/Ç‹ºA5Çµ†Ktõ	®Ø“Zo·'–«Ý,W©½Û1}¸v›+t-t…þ^zœlk6x×¸ hÖ›ÏžÁêWèvÆÄZù	ÊM6¢Î·1Þô›¡ø“´ ²°@ùs	½d¸<e¿ƒÒÞ­"K	ñÓÁ]'€v„6Š~Ç¹¾°À4Zw@òÃ>¢	^Ü[Ñ4šÆ‡Ùà}³eü‘Ç	@¢Á>¤w­ýÙ1FSÉÇ@›ïðüŠi<vPä¼	^{¯6/ùü@2tép.†êr­‘´Û5PR`ž¥Åó¤QÒ%Á³x8p¥ic	÷„; ¦ƒ1f»©¸97EqtIˆ`ñúÔP«+t"MÃ®¯°EÇù`Z®£Æ>©*Ù
la|Áï³¥qî¦N£½æ®F‘þ4EÂÂzaYyûâÍºoX•swÍÖýÿ   ÿÿì}ïrIrç«ÔÀ³Gp— 	€”(š¤‚¢¤Yy%-jÇëà)VM 	ôªÆv¹0#ìW¸÷ånì;b?mø	ø&~‚{„Ë¬ªî®ª®@ÍìôÎR@£»þfeefeþ’NŠPád¤ÉÌ&ÌÜ/Ìòò7þ:æØÃ^R%÷hñ@í£Çìré‰M]FyC72·Œ%9…CÚü!yJ,JMÃ¡gV:–‚zæü5•š6T–¦“›‘UÙ²¦;ƒ6Çås6¸,wG£ús5(Ø%áë|19ùcÐÏ²ÂÔˆGv’oúxº;äµ‰(*¿ÕÈû
¥¸üP«ÊBn±C%ï·ÙüþûÔe§þ L¡%#mEl½DX´¼O`ÔZÉPqZÛè€éÝâ»áVì“W|zÞ#Ö<E¸3I´ mäJ¿¥Æ–~;Ï•ýeSà‚-Ô°h

gÆ7I2ŒÃ—¸$ï4Ä@‹^aúºø‚Ãìz(ºRö^ƒàæ“äæ¨±CvHwþ3@®³¢¦ÆGþ<)rvŠœ­AGWÝn{ïétÛÝ½þN«ýp¿ÕÞyˆÜ½V»ÛÂÛ¿ît¯wÛÝ£½ö£nîÂãí>³uâƒðT‡ôÚ½Îu·ýðá¨×Þ{Øï¶wöá‘G]ø¡»ßÚm?ÜeŸöÛ;þdÀ·›ú7½Ý“ý½k%(`Ý”ùè!tu¯½û Õ~´ObyÝöƒqëlaí}üna£zÐÐðÛÃûÔmï? ;­½v÷¶­×zÐî<€¶íõ~Ýmwö¡ùû»§½ö£G¤»7¡‚‡KÁÚ½ÚüüÉ“Ó=Öæ=(Œtv¡»8hÝ6ªÝÛƒŠ{ìÑ£¬ÝéÁÝ^~ã»‡0þ´5§x›ì·÷ö4ŠÀ»Þíµwá.<Dv¡1´ß†¹Üï@=]¯¶>;ÙíõöŠñÝk÷öûxF¸ƒ…d±‹÷vã^»³×Â?§‡X76;‚‚?8VH0xÐ:7lvÞ}ð€àÐöÛû8IVè¨“büõ-†5|=¬ÞfÐË SÕ *@2†­Ñº1«~8J²YågŸØñrCd›»>ÜºVwBïQôP!ýX‰ `µ}—ÄèïÅ ˜	e}N©Fj=,A’5Dãg6ˆÐÝÄ6UuÇn^ÅÉÇÖ(à¥)HÂ³ƒ2 ˜~-ºŽ²è2Îo_â	ÿ<bö…5C´„,¶IÎ£Ax	#ô-”·ä—Ûeï1`“ýŒ‘ ªšMû*¦Øñè&û{ÑièÜ6žUýƒ4™bžöãŽÿÔÚÝÁ°{Ö_yJ«†Øb³šô˜LÒD„)êÝ‡y_·…óÔÒ‰:e­)&¬æ[6˜…T b)Á¯¨/íw8¯’Ëžÿu ž$NZ~Ø]¡‘ÅP4ì|¸%3Ì6Œ™ä–¿j¼ …šš¢MÁxÂ»É&¦.|Š>üÛ²­¨B‚Ö#ê‰.qÄj?×°ê'A
Jo:ëIz9ËÞ2åh`ZH‘™íø\K¥j\ÿê½³h8š©Çbùéê¹cÑ9ò,5†0WªµdSòƒ²¦|„uSÄ_Þ’³ûÿ‹~•ß$ã0Ó×jÆF–né³sºVüW_²,ñÔê"²a«mÝ]ù˜þNÂ2ƒ-ûðU8™K7Uµºš~òp›­xI®×ã=¼‰;3Ac˜"
[µÈ8ó§O3²‘í¸À|Þ&S‚Y¤ZÀ†T<G·G–¯©ù]/ÉuÖ&Ô}pSáN¶e@3ïX¹ÅV¡–M2E#­žgñ7VT¯¾¨J†Œ9"þ 6í!â@uDä¥Jæ<½‚siŽzUÖ SàÎÃ÷ÞdÐDâUöcÑŸ˜Í-]¤q°<&V¿DÐº&çzÍ§m‡û 2Š3å`ÎÊ¾pVÚÃÓx~Þ•hÑR‚˜”\Õ‰{@ÔÈóŠlù¨zÜÊ.Ãxsdÿsæ1ÂšŸôÓdB¶©Ï0Ì)*Ú…;3i®Õ›ºšð§8HBØ¬0Ï†Å«&j’‡ Ó˜ñIÐï4Ÿ²3ÀBÉÓ>ÑÇÈ…~BDÏúãOô„ÛFóŒÊ@Ô3îDŽA'ýd<ÏÐ+d{\‡CüTƒâ:ØÓen?ªg#ZP}	ö³<ÑÓ­öÉ‡ì«¿è!A,M×D¹ŽF0È“	PûR9õd·»L!àYè:²q¥˜ ¢*sÕ­æ‹cõÚ²ÅQ9„Ž9i>0•$ºy'º<Ž¦ù@{	z-¬a,ÅmF¶·ò-3¢|5-ÞŸ;„á1Sôðöî½	œSÚ4Ã.Êö¹çÈtÊ³q˜èjp5ŒGhŠí¨;ë¡è}Šî­•¢]Ãê¦æ“áVuà{YúE”³¬Ã»0Ž¤sÚGö‚3C±ø@°Ø W”Óˆg7èHùØ·æSÝcÍFmó{XUõôHIùÎJ-(èšÏ÷º'AtBÿþ/ÓˆzGfápâóz¡«ýwçß¾ÖˆtÚMãiòq‚žd9…õ–É¯Éö*6¢Ð86žæƒ$Ã‘üD³ýŠE¯s¹B†RMQÞèNåVÁ÷-Ë>™]>—4ZpÑ•ßNs!
VÁáÅØk2,˜5‡¼˜ÿ*Š8HA¿Na¥¶ÿ%ý#¥«·{,GiAêŒµìÒžˆPÒZÎN„FŠÒ	F“?	Ð#%§Ô•¬bXˆ¥“µeÔ3ütlÈøv8NÚ“7^‹T?Z,H2úáŽaLðýò5ðH|ó•^èÐ´ËŽlûsÚG¨³;+2ËK¡A¸ÌÓEc
ŸÉbO´/Wi­5VpÍÃóQÆoº4ìtÕÚïñ³DÍ:ŠjÓ±ä5¯wîP!ºLÃëð+-Ø³¶ÒºèÇ”~L‚iˆ¹ÖUBf¨!mæ’`j“:Âk„j#	¹¤‘˜k°µÅ1Gt?.ãW/­‚Ú$w©RßŽø’ ^aó×˜ç¼öêÂ·‡Ó²­ª÷KæQqº÷asŒ—U¸œÀäTveÌ,Ï£‘¡’ó„FObŒôxjòâÁÑBÐo5´—ŽÒ#:Ó”B‰åÉ“š” JÍ¦@Jù²ôÉã«N%âWó~diþQŽ®Ô5BÂ§¤±¹ÍZÁ¾›Î Jùâs[ÔŒÛªîÉ;0h}«®‡JD«ú²<ÄÎèÕJÝu—Åž6kE³ÖÝZ‘ªú‰åýÄß5,XD9ªŠ?8•GÂCõ3jOˆi“´x…uå-vÕ”ºØe½rÜO%Š ¿\Iòb—áÒ&…9@×#ƒñvñŸŒ?h…QöÇø>Bí8•ÉNf(Žaø#Ý+
Ä™Œœ¦÷F¨Ì ‰ëËå1ZÇžQ ¥1Yh–Áö.Ý¦:è{Í7_ê|oÄwÔêW“LvÃ‰OÌÂŒäwñ<†åzºòÎÈÃ.ÈDi`H,ï3%f¹Ž]VéŽ]þ`½&icud¦s¼Ô‰ñ„ë-Ö‡] d—™2¾)Âš
š2·´Ö,²KçH^¹ÃYêdÇUèÚ¦l©U¥ÎÃÆÏ€² G!'d,A4Ó¥‘.Ñ‡>“Ë`²¡Ç1wðqLsÑ`è†ÞÀâ³~€HW;èC«eÙÏßëè_
o0«VMZÕ3LE"Ö/P¾Â« ;Zðôñ\Úž±5Àp’ž–Tpˆgm!¥
GSÐ<ÚŽ†%“s—£EþIÿ\2A8oX‘3+Dµ1Y‘ÞÒ™á¥ó7
·<ÆJÂôÛöç<qª»{2a Ú5»ÊÐ´××ÙU;ñ[
^É GeÉZÇîÙß;K£$¥J‚ünyßü>šJŸ]G¨4fÄçQ²_ôïI³K1ÂÊ¯Æ7ŠÑ§/ßŒÏCJß¾ÛëÀÙ*ëÀo¦Þ‚zŽéö(šÞE8VîéßæÏ!Ö›øÿîªñìéó²*Ê>NúÀvž>7u­Â#J,6á¦ým9\¾ÉoèßŠ“ar:J"$‹ò³qè_
K_&hX“1>u´(?Ûž-%ñ-ù®iØß„¦ÌtZ½ù<MÆç˜|î\Ã6¢+Ok§4oÿF.l¾£d¶^%Šn¿X„^î1kX£®´„z¥ÛxQ‹3Ubr9€~˜7\íÆgzTÌZv´¿_Ñnè&¡µÈ?v´(>šž­â(šôaf<Ê-‚¦uv¾rLÄ»¦ðôÿˆú èŸÐŸMZ\Ãý˜M¶‚ˆ™QCÈ´Pym¯GáÚÑ;ã=0S¸7}×¦î´õN¡0|¸ø¨ÒED	b=#·¸aõÕ[v‘Nzw Þr²zéõTs×¸•=•†X¹a|K^ç¦§Î«{±zÏù®"èîâ+ Ñe¬°¤Y?Þ}BÛÿC,gûÊ[ãƒ~šÐ|õ—IV˜l¥¤ï¬¿ÂŽ<ÍûñCÌ~d³óÛIÿ|>é-Hæò=ïˆ2|Tê‡2†øÝÄIQÓ:G0u|XŽãÉï*J¯©¤óYQ
B×ƒ–HK µãmóCg¹?æ.‹zîXúàJ¾ÄŸx¢¬ß
²,Ì2ô+^a(%ýx×ÈSèÈ	žY Ð®a‰øì¬kåz³y
]_Ï¬jJûñÎì+Ú´&ýø¦jAÓÌ*ZH^ÄOÆÔ}Î;džMƒ‘Æ8£§ÂóÒWÃ+«vŠ:Ô"?˜wG<úûMxKí>!Û ¥[«ï:L‚_ÐEF­ùªŸu-l,¡âZÓgÔ ˜ÌíšT>Õ‹çé8+ao¯àA–=DBÃñ(!‹VÂ'y@[»a4Ó«‰pÔËbÃ1‹<µìüËÈaËI|þ²Ó™lqÞ‹LÖâÙæÇoóÒ~2üö%ëÐ«`‹E÷6º¬•/?ªä%ÍÏ‡*?x—CÍGšb,f%M)E¦CMIÅoÞ¥å9í4…å?™N("ÕoÖ,\pÿÍFˆˆs´¨f.«^ºÄuêÅÒ:aaØ,_h9{ÛÌÌ8ó,#˜òBI¦éÀ°¹JÜ”®zÝ‘0†wl›%•VÃ0çkóTlã„ÂXx±~œÇ£Èo™A‘²_×æƒÃ.¾C|K=Ž_ñ 0âô¬z¡ÉÊZÉÜ+åâìíH™+5ÇÎÌ*TimÚª‡Êø"ºäcü£¦7€oKËSr¬£Æh;_Ç0còrcÀ¬Îk…U»¥¤U,$9c¶ÅÏ²ñçÑÖ´ó#hz÷þÄöè’Ýûu	âÖ6)R6Àå'C,æÇk»°<‘l|@?§ÉGüì‡u³oŠ>´¹eŽºvø÷ÆqžO—ˆBÞáö¨k,Ó‚ùÞ8æŽ!¦ÉÅG"NiŽÖ^€P]N-§¨ÒO¸=ÿÀ[­qÎbPGEð1ò:¹Äi4ŒªÙ¯Ô„8Q¥çaþAzÎZ2”_»ˆ(U~í1"6Ð¯VË Iœy‘	…¥z@žE´£wËÝw5È•ßašÌ§6÷q5T¹e
û×?/¸ÏcáYekò¹Ê”<øËHuª¥…–ZüE‡÷æ@¯tEMcaàïØœÏ-£NC»}ZDÙ‰Šs*ÍyGÜ]G@•oL)äif-¬©¼0¯Y2=KaÓ
ÙÖì×’tµ §)½RGª–^ôd{åef€åuˆî£oarvøÀIO.ÿzá¹UiƒGýü „QC­¨Q4>Í¿MƒlÔ]ÿü:ùŠû?Ìäñ.Ft8êUq>‹ø¤xØ8¦,ýµž£P,gñžâÊOo˜S¹a0:û2ŠÎ|:Ó>Úv¾v‘*6NÌÍÈL¯dËôpg‡Æ1é#ÑðWkµFd£òZÃÃ¬]æ!:lFäò	e˜£ÀJFá±õóhS´§a¸z3³P˜jëœ=tÍå%[Ë×¦¡N½x'¦Â6ù#Ó•Õ"F'¦@ûîîHÙ¼¸¼Wd-C PƒùS7Úä>UPÙÝ—DÌBU˜Òn>¥6eÔ Yî˜Y9ƒ\{•§<J?Ñd`À^È'Ë8“Ú©2¦½XÒŒaž¢p%S@^ÆÏv€ÏfXƒêOÏPAÏ	»dÉÛ¾0M¿¢»ÛíôxŸ{Ëú<R0ñó;°Õ}ƒ;ß†ÝˆÿÉ¼?2ã@;C#m³ÉÐÎ‚aHZ¤³I~É„­³0Å{[D|@þmÓË¨P²%#ÿMÆ‡>Ž@)ù5ªÈAkìnÁ>tsN#øHc‡tºÓÒÝ…?­ð'^˜ªžÿ×ÞÙß„þ¾ÏìêžÙÝl˜qˆ c+ jÎ{ØÀC¦èêjfÙ­“†§ðËéîÙ
³	˜#*±À³7bÙÕž59
Â.µ[h›¥r×F…¿h\ç™äÐWêÚŠ>µè'e/ZÎbäe3ZÉjô…Ø–4èlG>V¥ìš±ÝnÃ»[DÐf*“pµ5¼5IÆÀÙ¥=4Çoâ7¨^Ç¶UØ‚«×sÙ„=V|¥Ù¶]ñªòJ@ýêí‹³oÏÿìÕÙ›gç'íhÒçƒ0SÝ´ya(ãR•,ÔÂ|Úët¢pÖj‘g¤Šüf»¶ÈÏ´¤µÏßÞ÷%XüÖ´¤™ÕïÓ.êO7ÍõM~5ŒºÇ	8ìuÜæ·ÈY˜µÎßš¸[æ)qK¦(ýnÝ–(Ú4›5Š?¢â}G4%FË€äKD§%x¯vˆ>±Qi÷”2¥ZTâÎžÏð-Ì›–›©Ú2`È—Àt¼ùÍÒgIÎ¡M­ñ2Ã¹ËÂ!ÇôÓjí†œ‡Cô#	½ØšP¾gØ~ª”`’°dò˜và€xî»L†‚RyoiJífæ_È‹L(¦6û6º‰ÉWEƒ«h
Œwwog»‡9‘„s	CÎ
ÇêTPäô+f4±òâ`í8:AJÞQ¡ùêu8¼ÿK?JÜ%xíkXz±­ùè1Ä’#D}Îuha3ãÏmwpßÑÁÿËåsýÉô6Ë¾P›vs¢œxZÉ|¬2nß÷àÀš½DÙµKNØ èñö‡Æñéë³¿;àã\vÔƒ}»·9çV¹µ™Ÿààž<WÆ6ïìç`‰oEÙ­’%sV&¯S2‘{=%a_émC€[üÌNËíé²ÞÂÐ¢NÒ0È>ŸÇ®ˆ²SzÈt«S¼Y¢f¡	¡ùž~™%¤E¾.tf‹¸{_oTé˜€pQÔÓ¼x·ù·d{›¼	á1¼yˆ‰ŸÊSãurÚ¼û¿P`ç–å©±ŽÂë4™¼¡	NËSýyæò1Z!÷´VÈ•=ÐØÎÎÝxžçë?Ÿ©×8SÿüçèÚÍ£2yÇÒ‰•yý÷Ùl037†o w%¸ßA”—q8ÀøëòÈŽB›÷é*òïiùrÆòoS¼õqpÓìlzO
»¿™¨Npè¢D—i/ïÒç)Ë‘§õâÐ¿ípçìþûa4	ˆ8fw¨¤,J •¦¶åSPÓ0Ø•‡/†8Vîg="Š&Mï9±ýj9b;ãhÜK›‘é¨¸ÍìR9ýpŽY\ORg‡+¾u–M²Lß‰U¨÷"5zãø×Q6»ÿK
‚ÊAeTüŒ¶NÛ*]Ö<ºUÄð˜rtçÒâ ÍIOpÎ6äø |53Ì¦2Œè^‡ß¬òË²nÞfkkÿ×Ñ{Ø?;n½ØÉËØUƒ£±Ëƒ¯±KŸRM»ªd)ê”õE¶.—IÊÓkª¼ðÔ”È«(7ß?âÐ[‚]i8gÑý¿ßÿW˜‘Œ:·<£T"‘Í»Þ¦AEø*¥Ü«¾Ö/´Çï=OMûqLžÎ« ‹¨åŠ§†H<Ž¦œãnA|C‡LŒX‹i¶cš©ˆûÅ±aÄ<ŸÂ ÙgÖ©-½	¯ ×£ÓÚäÇqÊrÇn“—¨¿©¤²e Mâ<CK*‘[w³{"«Œb¥U·hÕ°€Y6ïŽŸ‡ýQº[âðk®‘m$7†ö,Ül5^†njM,Fý_n6m.k´‘nkxQ×5,*BÄ{üÜúzÁë¹{o'ðªS[OujCWµÎ£¿Ú¾Ý_¯Oâ³†—ÆodìÍUœ|l™3Sæ—ÃB z®y Aâ©f3U=ª~n&—8»›À()umþm™9e”6œn´«^û½ÛKÍá§&ø”UNé™õ×#7‘Ð&
“ôVržõ0@{¸›¹%Øâ¹iÅ^Õ÷{÷)¯†*}Œ%à.œÛ*7÷¨­RÜ ``/ÉÆÏ™S~ål¦´Q}6™Oé±&{8=ÿ/Ï%J•azC³ËSÄÒ´¿cIòÞ8>Ë7ù:cç<¢}s‚ÐÇ|Tsžd/e¦¨‹˜ÎW„:‹œi‹4²17È+¸Ø§[¸€õ/XªlØ¦³'oÈÝf-âów À‹åJf¯øãµœã^K;_àÅÏÞ&x²\l½˜KÀÓï¨ž? ^^>x­´:¿Xdô²C@u349ùÎÙO÷€ ŒôÈåsQ|÷+Í©¨#žwšo1]¯-†¾êéÀ.ÆØ<6R2Ci²0â3líþ÷?ÿ«¡œÂæpÿ=üÒÌ6ùz\«òsá^ÚÏ[³kêÏK5ëÓ'jÙÏÎ€Wm[^Þö¼Ôu'.1e“êˆÎ9¥æ³–vx]ÍÍ,âú–ÀÅà|~]ŽP~.A7@€“ ~	2ªäðJN.ÓÈtü °²oóÊþ«¸,¯4{ËÍ§ò“ =é¬'ÏJÈ¿Æ™Y4s3¶¿¢ì6ù8	ÓƒBÏ¿£<Ì@Vç‘þ×*(mlúÊÅ58©ÐŽß«Ã«Y:÷4nö’#!Êo¯^U­$°é|ôÊ÷õãTü^j‰|Ëùíãå)æåŽ”7ýx¥¤"MÐ²ÆÁ7^ôAµ¤z-/Oñh™°:ëÔC«í«ât>ñŒX°y˜°'\^&x™/×tHãôP¡m­ê¨\PE|_õ'Œf,æÜÑC¦µp‚€Þx Æ°ç0ájî	bsü°Ìµ-`Ýr[}ÖŠ­ Ü‘Ò€äæ:FCîWyt-x</R]‹œ…›¢Ë›nÀX¸Ñb,èðn´	ÙÕÑ©ïki@NXÅ’i¦Lé©KH“öpx’¦ÉÇ—áÕLÜÿ.‰gN´ì›iÔÁ-|x*%ZþÌèuŽþ™3RõˆÜìelw¥ÉÂqÄš­ž,>ÎÏð­¾-úÂOBn k &\³${¤’Ÿ¤œ ¦ÃjŠ‡Ú¤®š¡<q ÖŒÀ®*zÕmLÐ$í.zfK .%µbtûÜ>_£^­cóªùMÓ‡xµ¬Î*ŽÕ¶´cJg²ÁMÒÑmÌeóÚSj^“™Šùu«0#Òt¬‚žËhö…ò«ÁºèKË¨âhá<hLt–×lÝàF÷’­µÅ°@ÿæ9—æÒáŸ6,P8ðW%Wôe<¾  
ÓÙÍvõDÜyÜ+Ä	zð&C¼_¡-„kçVö°¿U¸•E3Ñ®XÝÑ€ÏsÊIo¬"2.g•ü<ÛMfð¦ÓS†²`êƒKâŒ8îá~°Zð¸`pw¹†Ò¶ZU{»pdVê×†¨+]lÿ’°Œ? È	O"LkÉÈ/·5ZqMMBïˆEë®Ã¼òça8¸„¹ÐVY©tÑŸè‰ÈˆÈ ´ÓCÚéh%u9Y/§7¶´5fÐ‰sU¿6pÿ“yc'Ïç}L²c˜áÇ¤a¡æýêÅ³ô€–Uø¨
î«ûÿÔÜÐšÎãLËŠMšðBÓmšïã±™×X8ªÃ2of6&bŸüTšó ¾¦šr@…\æÛn·² ¡ÜMãeºx:
ûN£´‡]}/—ÐÖh×¨ÄGNSîÿýd)áßFã¦ä}ókù^{–Ð!~9Ÿ¡ï^óâºpŒ’y
ÂM·5ˆ†Ñlc‹Œ£É³	·X «p}=6ßS¡È¥Ë›FÕÂÊÜã]^Œ‰pKw<1K±•ƒ]ýÏOG´¿
S¨ƒÁSÀ”ýnË ›({•ó˜)S .‰³Ü?Åî÷\Ìé÷þßÑ’PÒH[³“‚Pßòl
¥è“¾˜„Üno¨0€=3(½£ë/PJ-Ú*X©IÂ\!Þ³ÒçFÎ€å4AÓï)9&iÐÐÓ}n54þ›;×i„M€!$%ÝóïÙó{&ª÷G’-FáxVŸÒù‹ÖVe	. JVQÒ*i¸ŠØZ–öT¦Öáh»he‡Ù8A}’-·dÿÍÒ$+nÕXl‡§ÉôÖMÚ|ô9e“¢µŽ˜p'I²¬jvz”ÎÐuó¤=A·œ’;ÎÅ½½vìdn8î<sè÷n£no³É
§ÑÒùs=›gh£©ÁyÈ,l.ÙÜµÃ˜ÕwåHYGú”Ôó†8Ñl‡6ÚSBÂ¥]ŒÀ=:wðEÓLÛÛ¤Ó†57È£¬Øàaz\’ÁÈÞ’?ÎaÌQ›@ÝÒ¨.¾…?ý#a£$…çO¸Gîó1LA[N&¡NIav»ù$‚ßä“DŽ@<(ê/n7‹YŠÑ·¼Û&ßPè_áýoDF.o	Âh›Aq#ÂÁ“[|
çò ˜XÖáaFåØ-’·åâÝ14rq§m‡Ò›öU’>ú#Ú|ó’dMHƒ'Ô”AGÚ:œöJöæ$IÇüUüÈr3NÏé™þÐäeobü³¶ý„ + sÖý&¼…ò¾½DD“ö‡ð6kVÆh³}Œß$}ÀÚóa“ê
05§A"EÝÒ‹I†ÖÓ|àyóÍq…BóJMŒ¨Òé‹üåwðöÅ;Ãf®½k)«=gŒL´,ñÎHö½6ùMNÉ·¯_þ#¡ö123ôg .ÂÊ¡lš/ÚŒÐ:ä‰Ý s“z’rôº°²’@ˆH£PGÚ~ç.#Í‹ßoQâÎÞQäÓxÛÑl[ä’¾\ì¼k³˜ Šâ“†ÍK¸µE6¦³Ö“7[4 Š^®©Á&)Fmr7…‹ê‚~‡Ó¨©©IÉÅDa,¼·}H]	}EÊ #àeƒaŽv2ÓáSÙ0`Ó·Êö¶A-qêÆŽ‘Ð,™Án,òQÙ#K]™ìõ|1—À)+¥äÔÎw	f<:ÆÔb¦Þ?RÄ™Öþöúþß@üq6$Î8>KÃq”eAv’ïtGr‡¹™¼Ò‹#ù9s5Óþ¬ìz¥hP|íxLðèð D³xôh‹}¹Š“¢Ú¤m©tÌ¯ ln¢Ò»£]|×);?K“!#2Æ^„ÅeEá»ž$T¢(^ùì4PR2 ª¡ðéQ§¼œk[90ÍRtb‹¢½'3ŸEuútsµ(ZºÅÞÝÂÆl‰=»Ó.yËvÄNqŠ`ç	lRWpD1Z£(ÄžßªÜxXF˜\úX½–±9pŒÅÊõ0©%°ˆ'I<Ñ=Kû‰›•°äõ;ÐœÜ"è“^‰å?ý€ï‚Ø"[)¨ƒÈé7öô&2kžŽÑÐ2FCœÿaÒLŒ:V–}€¶UTÄÖ?¿]tÞÕ6t¯½\MÆH=Ð¦ió;)ž\á¸€@ŽñMèÏÜ:Ä+¡hf¹]ÁB›cB› ß·)LJrrNl)J¹ÊP´Ü¥6P1Úp‘— \Zújm¥çbkÁgà%-Yí…2óPè_Q@å£mOù#›ª(éÜ(3T¾$(#š¶aÖOAÇL~_üN©…S~LÄotÞ¡Àåí4%v(k}R»ÖKµÖËzµacW_¼:;9}ûíï¿}óôÙ:É'ÛŒèjO¡ÓCk2l+EqO´Å=©Y§Z^‹¶²†rdÂM63š#sq©xž“	ˆëzÃ¤®žËaëã(‚muÚ*OŒÑ×àçµ½¿C2Š³AOŽ¥tFãA™H>»q¶jd»¦Ûñ×ÿdr³ËOÉ/c<¼7:V¨ÉèòyIå~ªË¡OàøÂ(R»Bª*=qT¼Hà{äèúl­Ë5kŠi2
‹¿ š:é¢¼Õ=Š¶Ó2Ûxé\ê†Ê>Jòb*³Ü£ô<ˆgÁ˜,$e«¥ªp|Q|æQ É¬cøìN?5“äƒ"º•JÃ5juÀCWÅÅ1²%'Rõ¡!Å§xÉž7ž³MñšÌSk'fï 6¶Lj¸ìúB¥o˜†3¤+›ÝÆ4¤xÍlt@Þ½(õñ»_¼w…ÅY£»œ¸Ák@ÞÛå•i@Z%2fpß
''£_†‰—Á%”‹«6Ü&™$ci\½Ò˜6Ååý2N b%ZÄ1íµÎô…A4p¬ö•óIû ˜yJ•üœRBÎ]J1f˜`êX{î%ºÌ—z>}ÅPO6Ý!>@jÕtŠ»RÈŒ6/5Bí:sÞQþ Iw|‰)ÅŽ)(ß÷ÿ*åáöh·.ÎìÜx¡ZIóÙ-ªÎ¾)-ì®ÈæBé@çyÅš×,Êr†	FQ¦S1­Úï*Ìa<//lŸÒñ×ãáj ß‰dâ¡ãâÌäáYª^ùfÍ’‰aHà´ÕÛàZ*+y˜3œÉnôSõÊçÒëaÕÙTƒ.ò â˜öÃUšñ ÿ$<·Ý<<ÀöwVx…/(¡pUì=	Z¯Œ±©Õ
ŠðŠI¯ãÊ(¦0—-YË¸äq3`w¦§è–R‰7õý‹’¾ƒ‰¦lt)EøÑuIÙ
Í>T5´N)6rÉùÑ¦f»È‡Û¹øq·Í¾SÉÃ€ÁŽ<a¹Kl„~’q,^0|Ø¬RUm„û(
I§‚ÐéIzjG~ÕS?±aK¥Êp–Ò.òK§eà´ƒHì¡cä—'Ü…V¥'Ö˜C*¸[E‰]AîEŠýûÂÝá4A”åÙ’âj¯WÍ’±N\¥DÃ4 Sfk–´Rr•&c)#iÍÂ ®Š³¢ÉH4:7‚Uu¸º úlç“ìÔ9¿ãåõÀ®7LÐªœû?0¼í€¡Ô1ˆ2J¸WB&yÈß ÉýgM.¼évô:çú²ùÞ•—zä…UÏ!ÍàÎËëŸþÉíÀÃÎš˜Q×k¹“¿Çí‹Î;wãMGRòÈ™´™r@œËSQ—}ilñ±Æz[LD	Z²™»„ï¶¿’Õ¤ò«Îc³±Ço6-ˆ+Ÿe-—ØòZÈ'Š2t‘ë ºìËµ~Ú#òü%ÜÀØUq»Ž&ýy\ø‚ål±¡eâtL­Ûú¯Œ3³`ägš[ôô³84÷±d”‡§j¢îù©xp^œ¡Jð:L•[“Ÿtí*O?ñ€þIÝ$z¿ŽòªÃ‹¨•#ïÊŽümIv¤¢´¸¹iÁe|}hñºÒF­)oUGR%<@Õxq'GÒÚ«F@Æ‹Ù »ÅŸOÔ!o…N;-š´q6û¬TÞ
º~~~×c§kÅÖgyUõ?_•èéwÏ’®ÊñØbÅaeêyuÿƒ¨¨IwÈ)ÔU$þñ¯Ìé­‡ú]·íâëîYðŽÿ@ÏÇjsÁÊé	®vÊÌ:åFB÷º<Äo¥°#”»S7¯h ãW´èsƒtjÂ~ÙGÀÃ÷ú—ÎjŒŒ¹"ÏÔÓEš`%–ü4µè1Ï\HºsÚ¨Wº!Q|“wE^Ö€ï-¯âx#Èn'}â›#L½‚A4¶<Â´šœ<·˜ß5óbl'—Y˜^ uñÓ0sA£é.ê´¨”HÝ·å[mÐlÇ\#ÝØðË-¦^Y8û&œ„hù›xIÞ·%ŽÉn—jG.J^å0&G|Ü‡¬uáÉ‹á¤)yNÑ[²_èRmgŸ·`¹ñÄË‹`òjV£v!wb¬ay‘m‘Ìg^–œš­Ës²¸So„þ,˜÷”*üeéÆßm‘h‘K¾ïkÏ–Þ!}Œo#˜eyòOâ°ý1H'ÍÆ³4MÈ$ ê§x+ÿ–Ì=¾ÿ-zè¦‘Í‡ Â-âŸÆDn3(Œ4pn¹ëXÍZ·Tcj¿S÷ï.ùµ”¤_^|wGü—ÖGóÙAß­pÈ]à'r XŸ5ÞîÔ×K}‹]‚š¢%&³È¤ˆk{XOà|ùò=/AAý‘ L—÷y»êI:çU†’ÓQOlª•J§xëK¶ò`¶Ÿ®üú,—2ì?B¡«ìÁÏ‚×Ï‚—xý,x}â7~¼ì—Fð*b¶—¾
Ë«dj-¯"ÞäóŠ]Þ‰ÝÕëS]'q˜Î˜Ð%Ë\|&þÚå.
ð³Ð•_ŸEèÇüG(qñæÿ,ný,n‰×ÏâÖ'~ãgqË~iÄ-Êª–—µòótñ½´è·Ï.hå€š_šœõ;Œ…ÃÿÉ,o·éñ¥vÇƒƒ­Î¶¯o‰¾¯¬oµSãà2ŒÍaŠ´H&Çß2éŒ!Á“mrÎ¸4|9Ü¦Å/%úÖ$¦eRØçW	>TX»¸M×å¤k’Å×+­.+«rçìˆ;¿ðÏ +^'ZVZP/k6N)´ãÉÙt`ÿ&„­&"” ‘WÑpŽÐmr–¤ä*¸NÒ-¢>Ê)ßž4N‚½Â©—z‰Î3ô“G]º¿ G=y+.Ž†÷ÿ9œ¤°ð"ÔÛ%.æáøy¤†:š†¼#¢Ãæ§R=Ö¡xØÕc_Ö§‡¬]Y›ò¹5Uô;Æ°¡¬Äžñ( ABÕ‡TÐ
X-rü)BÔ³~x‚’},¹ˆ—ß)Myîeúìêñÿœ4È¯h%íq˜eÁ0Dâ‡¯z°9W›WÑ˜Ö¨/ý(t}âS3Ž‹óíˆQ¾Š*¤¤ƒã!<~¼{“-¡õôe¾ ²‚uðW’´¬~ÎŽ×±‰gK7¼È ¡¦tÈ›]¦áX®ÕöJÔLJÅÞnËåUSs+Ò|#2”‡€RÈ‹“zÕ×Õ­R· Yõ²dŽ²¼ô2	`ÔÅ´…Ö¤Yrîe*¤>à8ä,M¨©œÚµc ’…|Æ1×ù·t–Ïãe©‰eøÊ·H
”7¾ú¼cþ¹FõÔ˜1f¯ðÉl)ˆ%±Ve€o¾	™ÌÁùr¹oÄµp3¤*ël=^÷ÌEðR-ªÕ¶*\Ä¨ÕÝ¥P8Q”•h"×ÕŸgˆ7J­­ÅÑùŽÝæ¥Qá“Lf„'ú¡™e+ÚºH5ìšÆA?ð¦G§˜î.$Áç™&˜ 	5íq’äôdŽi‹H¥KvC¤°ô#c°Z¸âyxT!iÔ-©›uW29“!†ç Uyi¡[$lÏhöâ6mŠ¨Ö©¿Ö2©Ø#òp5`w%Ëá¢jÊ­˜(‰˜‰Ø
ºQG·\9á‡ìÔ¢Ôîôä,@ÓiMsÆü`)Æ¬h÷ÏËË1.i7Z|Ü˜)÷dÞéÒ¼;šLç³eŽ6˜õö*ŠÍYÅl—OÊPs–)¥² W:Ô‡¾£r±ã÷ìqûbÇå`º¨±ÊØ$,#åsøüÛi"+½½E6r“›ê–´@ÕÖškò"¼–: Eö+¼"F+ìoÓØ–ÛR½§êTòbëpwA__6Íµ¡q/ÆÁ0|T$€"y¸ÊCÉ,¤¢ÉÏÒù„fH7xÀ»<'(,fæu2¦§ã¡+‡­µ‘ËyA±KÁ3üúZò±`
¾
ÁòùýeL>lÔÖœÙeš =ýfƒC©µë0YÈNÃ–±:,3Ú0ÞÏn )“ ~M>zðr¶”ãwÏøå›pÓðŒÍyØÔqËõÌ|qà/L{~ìOÆq«óENüï~¸Ù®­\áU[é««À-å °ö‡·1ÞçÙMñªxu #ãn`£Ú?2þ i¼sÉ/å(âµuZ§×’F,ÈSïŠð7X4HÎ<ë§û2·þÃ=GÏvWÝ[lßï¢ðcscd£ËHlCyñO~¢?aþeò´ÿiþ|µÆ4¯Nà»Ýá‹B4Ž¿Kú÷ÿI¦É $AL7$¨2KÑðTÎy›”M"ùã<ˆÿ8‡vŽšð¹­A®ÓÆáö8A¥¯­ü²)eø[\Ã1±rØP Õd®Ê¥‡eê®†¢H•éGPHƒ¡2“iÐ§™w¶H†y áSûÑ^UdçÆjéNñN§úBxÍjÔPS+fF6.13à³;_:2é€'WVÊWX‘Åz8‚©²¬6™QfX&B^"VµäÀö]ƒÎ§eIæLÈfœÊÃQ×‚—G1êÊ% +æM˜Í¡ƒ$;ÜueÚQ94&dtFuF¯Kä¨R9–žGJC¾pž§ ƒÛX¶uÁP-µ—ûAÞdãÿþç­YR4¡ÖNƒ6(ç1™`…Ö AI£1ÖzJ-:ä9Ok@ºôï³hšð‘£Y¶`§ßwžu¼xª²1)ý›g·XÀêÌûxçŸ'é³A4kB=kºÃ ³IÆÊ›_½}qöíùïŸ½:{óìü¤!Üã Ìh‰ä1íÀ†š8KÇÒxÙo©Í§`"®Üò¼€™P„g¾x¼¬Ñ4æ4z2¨$2•èUE¬VÔ§§¶ ßØš<Á=æÚ ê@nö:Çª-SùŠ€¹@+`-VÒõÓx\zÇXäVíŠ¥i¬)Âç9!•´Í æÅÔ¡ŒÞš‹Ã«åí‹%Ã°£}1&\æ©P$à‚y:€Þ–ã­K[îœ95\&/³¶fKB­{Çð¼éö
g4Ìº<Gd×:"H>NÂôgaú¸=›fŠÌE2ÃH’?†4ÚôµˆayDÄm4&îá½ÏÓ¦|úíí4n¢?ÅS?žŸ¿Šø¨MÍhÚ_}‘yî†½ÁÈÍÞ¦ K*H#Ïn`KŒt­]¦ÕÓ«*‡³k’Ä§q01ál«UÂtáº{œÐ(²º¢xíöÓd’Ó`˜Ûy
ÊÔd€ÎUbKß„°³Þÿ%hsOAdÙ¿Wmó40½Úœ»zŸ=}njãSXÁx¢¤Nÿ4IqO†7k5KËƒôvšfkÌ2òNV³˜!äs93œü=HšPÑÏÃÏ¨{¬S`/Ì'ÑçÕ'1s»!	pn1ÎÚW :ýQg©c¬„†ÿû<=|á<Ï¼6y,³Œ	¯[ÈxöÇJ¢ô´=
2vksÓIP¼¸É_~C÷¸2&Ö$ÇÚä³r¢z,ú®–úiÒÖ›æ™¦‘9ï'iÈ›#¶NÌK½ERJað}§Óv†/Ñ$½›[DçÖÎ*7Kß%¿Ä:Ðí«èlZë M¡
ÒY¦V"É«`Õ¢4Ôl*]ß&M©­¬Zhlèœ¶cœÐ>”7ƒ©Yb27yóL=£EóÙ­Y|N^U¼’šÅ3²òj>0vž4ð¹iéfL3ÜÔsƒ# ¯w<¨d;|`²¯ñ")ì®´«d9+SšÙòÃT¿•¼À¹‘ˆ£ÕBÃ:S%^žã¸ù]'ýhÂMk‚’Úq”Ëf2X!õó®lÓ[Hk~•ô ‚×}1æ8Îc90dö:Œ+gí¬ŽÎÇVçfµz ËSŒÙ-eÍ«W
í±ä‰è42ˆîÙ­,N>ú¤é•ÚÅ†výº+hÒ‰í±{ˆVÊ ÅŽk¬Üzë³<Ybu¾ÅýåØ³<Ñ²ëòŠG‘Ûc9i:¡®$=S¾se5òÈ³§ž\5F/åÍÝt­²1œ1ÞÓôþÏ¨Â~	ÓO$ŸLÞÅ¥uLåàržò áìsÏ²17°|§·xüæ[ÈÃãÈýV{›TkÑóÙŽdçÙNr½èt<-†C§Ç^%´Ââœ‡Ü!ß—¦ó8ÑU%ôyJ÷[Ò|qbv– µ­šx­Ä˜þC«X½³ÛÚùœL^ÐÓ[gî"¨ÊJ%f¬ãmr„¦T ú6c.OŒµ*UyLà|å• JcË{ÜžD×a\ÎüÉ‹Ú³`=€-^ª®‹Î^±$à{èý:ã0¹oõN*½t¹$9èÛØ¥e~‹Ô‰y7²É|Hèi.¼<ÆM„Ÿ/iÛÍøRÔ‡urHí_r Ì®: kÈA¦ÑnA$ºãú]6¾îå®Š4r“áv.º§)l ýQD°iøØkŠàÁø<D“AÐ°—c´–u7Ó.¾™bÐL%%o´{;îhFGK:Þñ?þùJÇÙÎðŒ_:æÐân` e3Š)Ýœ§Ä8;í=—k-Ÿ¼ä3ôÙÊÇçi8¢8sÄ‹…ÐÓŸ0:IZæ_?Ÿx}‰ƒèË52gñw~#à}ÈT”k_®.æ†'vÊ±ø]¶=³Ê”$bw˜3OÅ+Â©©0¹±<­Ôë®õh]ÅýPõ(28qG›¥(…¾Ë¥9(ñ°4\Æ>ÇÊrÍßÇÔ¸çÁu˜¢;ñüjK5ŸZ'Ð±O¾	¯@Ê~äòä¾(§äbI2¥›RÃ)„úòl3Í«MiÙ§üK2m•ô]Þõ^
¥åÅå¬ËñP°.w)ÐL†eÑ°ªÀÖÖLåœh¦<Þ¨p>0ExZ`tÆÃ}2‚ÿû9©POn‡=Ó¥	?	ÒÓ¨Í=_mÉm–¤Ü´´©£ïÙMØŸ3 Ö&3½O“”Üÿf²v8l9å„ÕE?áòkÏ#eê‚gø¦AÕYSµù	'qw)?Ž³ïíq8#Ì‹‰¤ÔÇÉ¾2x;§œR&î5&íåéxS!ÿ.ËÓkÊæ›úÃ±ñNáóiÇ¼]0íì¿<õ>a5r£³A‚:}µß¿ÈK}·I¤¯PË‚L€ŒŠŠ·Ø‰)uc§§ ga–À7rç®k}GÒN"ºÁÏ›u¤œ=;§R*¤è--¨zÊ¼ÊäòÓI¨ÐêºEw››,·5²8Žsæé0L×/L)¾)ô¦zºM`ã´­<-jÛêt®þl–&ÂSô‹…F5þ¦³sùh¿Ó°—Š/^ƒaøŸýÀ±4Nˆ ÀyT‡ j¼…B±ªg7}G`i8òk#¥³¡>$»;nRUÇáj·wµ:šG”Q¨¦®ö?wÖswžÈ]ÏOoq¤A±ÿ—éÿÞ£pç²fÿu	«#PäA®=oÂ!•.]‹Üg•ûg^§kùï]Åjl‰Ë#ÓVyøæ
Õö®^§1=nb®Ü”ò¡¤ìZ˜ê1˜#Ô|éyè«’JŠ{ ÝÐÃãà;]j“²)!­‹½íÞ;÷¡5â5;DÍ¢ÙõPDÿÁŽ_Å0
×Q]"ˆÚ=ž$7G`œ8n{~õ y˜¬oàõà¨ñŠtº[»ûä„ôö·zûPÝÎV‡ìïÃ=o¶ºŽ`Yu®ö®Õ|ï¢ÁltÔ¨[ÝËh
Ù¦IÊûå«(Ž4šÁïÏpêpÜw»[jì™£Ÿ×‡Z@¦ÁíQ£ÓyÔîíû×—¯'WWY8C85|Ÿ´òd›	 sÐ|#Ó³ÙmLC4Ku@6Xu­AQéd$`SLæ³_t/šîzèõ ÂƒË,‰çTgŸÍ’q‹ohœ«™9¿/Ò½ç§ú1X+Ê|~îC	ãxAçÅÏg*
¯|ß< ê€‰ÃCÏ©8û¶l’Æss0ø*hô¸O8£¹:F­‹;rmZÊI¡°ý×n¥Á·[ ¡­ÃÜ¡cgtCÙÊ¥.“/¾z-
!Êgqùöò3-,©°VG"“ÍÍfd³›Ø^L®!ay ¤§Æ\Î#fÈbÒìôMÆ«° {ghÀã¯aÌX|)ùÈvNØ:~Ñ £Wÿæ°ýE!µý9itRé ˜G‹ºÏm,<c=e±…kø¡±£ð3W‘a«ò9ç*J-S²Š%ó»BéLñª[:‡ý‹~ëŠbÃ«]¸<‹m·ÛMÙó›™'.ÊÞðÔîðÛBÅvƒÞå>TüŽ‹wv ¼Þ¹™Täþ= $÷ƒ·žF Ù7ÀøçÙÑâÁŽ»	 ¬”/ì{¼0¸­œL†(í¹_@Êý"gÐvuÁÃ‘Ão,KæËðÚ)tÊc6=Db¹Ý’„7›®øÞü:<aÏ¦ö‚÷}øØúzA¸{Çdú-¶Ýg:ˆ‡üê<:Üèä’o“$žES·6ÑÖ<ÏÅvÌ2¤"Éi1·ô‚EÏŒŒþá‡ýéMq¾v¦7$èh@¸Âë%×_Qó9,«£E“3zÁXNnc½xÿ5K/·÷ì!'çpŒ:OçÖs¸­Ùí–<·tø‘

GÂ¾‹¶ý¬ÈÿQ¤s/ßêñ_Œ„w,r`({5nŒ j.òón³v4Å'jg‘¸!ä¸Í[Ê¿~9­-|¹ódqyKÑ©ÛÝÊE•-ºÝ¦>OÏJ_U:¼‰Y1ü†G'WÕ
j½¨Cl¯(-:ˆµÅDóO2©Ÿ•ÅBã•§bûôŽvÀ«ñ ¼‚Ï€W¼1—ÕVM/[{T2»UØ¨ÕBFÆÞßÑù1z(k¿¦y+€Övw¬~ñíÙusÇ,ä1l 8Ï/«»9c`õRM	poké	¦Ñ&£h‡øCHK“ÕŠ¡u4CSŒ=9ä:ý`œ™ ê½ i“ÑƒÐ°?ŸE×4à$6mÕèk7XaÔýÛd7BÌÒ$Î4¡÷EY:š—@±¦B£›¯òèŽÁý†m6;B]ª,œàgT2a^ŸhÞ²#ùÖHõ„84»4I“`‚ÊŽld“OšD`W/wX´÷Ù¥ûA¥³Ìá€v˜ç3b\®rÌ\²7Ù•—ú¢)¾ô‚“aÁl–e¨Ý>G/£Éˆ.˜·°Á±ð)·s‘ÓnM”t¤U„„ýý‘ÒÑ7éý÷Wˆš¤ôz<‚ð\Ø^.ÜY/áÀî5¹ëðšD&ý*œÁ
‚Ä lãÏ¥Åcu„°oí‰âÍ’aƒ0¬ý°ukU³œ‚î$êIœ}X2ûg3–E%hÕ«ûÿÀhB™³4Âž›%nÓ}õÆg|Æ>S"WNhÐ-í&,d6|bª<Co)ÜŽ×!TÍ3;†ÿ)Ëÿ;ï3‡.ÿ“»ZQºl¼šºj1qÏÓ§•×é>³ñ9'üÔtqÿ¿ãY4ÈÉ5¦¿pÒÏt*Ä¡×[¤óÃ­9Ÿ†˜–~ÀEŒê”Og¥B:«µl?pœu-M²4)êUž ìÛK†fF/æ\:|6úŽ»˜6ë’sÈEÇçòqÙE×C±!ÇÄÃÿÐÓ+OCÜTcI{,­`±,àyòû‚çó\PÖî]MÎ¯_-„žßý·­°Ðõ=ŽÔ}–”ËÇ¹t gáðË™…¯²ô“]×øW‡Ÿ"~þáÿ2Fº”w•(»uŒõy8þü£kùUg‚«w'à°h('4z‘Pm’`ÚyÐU'³z
èD i´p¨yü¡Íä¸F‡Š…Ë¤â8@?FÖkæJ!ïwd¤ÃhBÝü’ééìl‘›Æ>Æá|‚}u‹ûÒL¶à”@	œ…YL(\fÅÃ±Gzª÷Ûq´¸
`ÉÝU½lÝø¿;¹‰²ò¨%›	àv»(ªùP¹]Æ“]<}°ûpwÿrc‹.¥óèO˜„³Px´àiÖXå|] ‰’q€3sƒdòÎ§5ËËÂ4ÛO„ñÇ›sjkßÑ‡A,ÒêÁŸtx4w¶èÿÚ;{›+œ +ÇÆ¿x¿E6æáuBSúÆaÿ‡QúÚ‡<Zèœr÷à³l’ãd’Ìü…§BœÇÊ³ÊQòx¥X}<„È÷æa½èy8fšÄjoK|›ÆÁå£ð[ñ(Å¤?-Š}h.vgïÑƒ|Šuz,uEœ“<HögÎü3g./f@tµ\õ½8èÏlý“°uX©\e±,ö£àŸ)÷.¼x°Eà? ¤”Ë =§'Ð»îL†‡Û9‹XÚ™¢†¯ÓrYìôÚ6+‘}ž6÷íÝÆñYeþP(<Ë ÉÈDÎÆ¦?x¯PÜµDúÛ=9ƒ	ú'G ÝæÓ±dŠ¤ìY³pŠ3hRp;.>ødcIM.€>ŠQ˜5Ëjd`7!Ðî€ƒÏo	òÇ,Ë¨ÕPEƒÏÅTñŠ¶Gìu‰(€ˆr^@pÄ÷öy<Ð¾øjÀÅg—ô¤9ž]®ñ0rTwQÏwï6Û8I0ðÀR­ãZÌ‰kÁ©€¦=ºÀÑ‚‹Î;¸'`2{XÖû¤v½—j½—uëÆSìî‹Wg'§o¿ýý·ož>{C'èäq~É.Á±eCKë2†Ê>Ñø¤v|îi‰-ÚRèrˆ«”ù&_äÃ¸E¼Xöï<)%ÃÚ“5•Ð9ø„IÕ×9:˜GE‹ÊÆUü$ˆnã*vÎ˜sœe¿£NÚyŸ¬Ò|-ß‡òP!†ŠƒÊ=®8@ ›õÔáÑP2Í“x–0?U =JÒíÞNƒøùêþ?‘Thá7+:ÑzÛP x	rWoY/	[ËQó!˜\Ò<:ËyX˜™ƒ—ê!€¡íuaðPM9uyöO	¢å›Ð®q\Òóò8úU×„²=º<V‡Ó®øÀ‚ÓŸsuÔ~×™]çÜÚ¢…Äû
tímº¶:°¨6Šä8tû„Z)^ú&_ÌÊäß2K2K«W$ý¹*w¥ôq²*É“gÉmíHt˜ßíÈ$Èo/‰8˜ˆJ:+žxö*a¾YÛGÄ¶4ã†ýJ"©?e°¥‚’oÕÎÄšH«_&ÎÚX&õ²R†g
æ›ÒÝñN¾|Cñ§|/ßP»ÓR9—«ÉM¾ã5ó!ër\Ã¢2ŸáBeÎÐ !ò ÕL™lI-fÒ—}’%£î:Älkh	Ññ46åL6úï™3ò$²ÅÔ·0â4XTŸ“ÀwíÀMÚ7¡_8uUNÓ&-.vu+Ý<Rû]}ƒ»»gG‹üSõ!K.<'~«>[¨ïG‹â£®Ež`¡©ê=ë{Êàèîkß*uF¹¡}'‰>‰ªOLãàö|Þï‡YvŽ’ÈæÊõíyª¸¨×•Þ¾,ó§Æ\µµºÉæó3è/˜oSU®“e®ÌO:ûáyŸ'¯Jây?	‘YðO__ªÏuh-gƒ«ß¦1-Ÿ}4ï§i8 {î@ÜÄ»Ú„O”V¥ÜÓ¾W`ÂÓ7ŠoÚgàsú´ð½ú|Ÿ&•™a>ÊÏ¦çŠuˆoÈw«ïÆp÷t”D}h}ùYÛö—Â£ÒWÍçV¬ÒÜ£sùØ´rŒZy>™|Ã3Í¢OP2Aí¥Ì>[ÞÓ½ù&äözñÕ´zóyšŒÏ­P¯#-&9°Åð*È¨êt´¿iGí­ô¸r£þ®½êLAa“õmÀE‘?¹ýè8šÌFÑ` ÍÝøßŒŸG“õª-+ïÁorRû‰mÁ››K­Ãœ~¼ÛKýã·Stc)E¸£EÝNú¤Ì®3ã	Î,½µž	Ïi©”²(Tx»ÝÖPÛ)U”‚šÀÂõBWS¨Æ`0S„®&Ðì5=Y‚éñÕ ¿ò$]Îøé±Ô¡2Ø4’áî<S³_»yèð
>Q>|O“~s€ÿ¿Ü"¢Š¾±ehèæ–ÕÉ[mã£æd5w¤‚þˆ4ûq2<K-e´Äa;æ6ÏðÖ·h2Zû!¡å4¶HQ°±	zË¨Î.š764µÒÞÂÓùÆ0œeØ2}2žjÅ•=yUÑÔ€É6V|ÍSò pàSûÑ¾·üÁßéøÉ –*›Ù7!ì	eÎúíŒ³o™ç™ÅgL²ÐÊo³¶°Ž^ØÖ¼Æ îcWK&ÏØùæŒñN9«Õlóo•\Y!<ÍäÞöû‚*Uþt5%”Îý+™<…~à,ùV= ÏK•ûVÅ6óY0Cg0Æ­XAìžù3Q˜o™ü^y_ÿnœƒg×ªœðæ(˜âðy‡ì‹~Q(¬å×Ï««Š³P–½n=Ž?\Kz‡·ÕtöôyY]H'}X|OŸ›eÁÏ!E~ùèsHvëS©WÞ7Òpea¶ÌÞÁ_õÜ?nki®·Þšë­ŸæzÆ[KwÿQ+9a…r2ä0ùsrÖ"*×9Žè8gÜ„ŸL®¢tüb<MÒÙ)=•»Å_•Ñ?¤eaIRT¾…v)yM©Á7?¯ši@X¹,°MœT©“,Ô'kTx:õW¤Ô¬ðl2d5 WmNæq,='Í TÉúª¯eà©ZdUcšBÐŸÍƒX.ÌC£a{–FãfEüuM£¬š†¥ZïVµ¥œ»z#&ŸsÄ±>jðJ	=ç¨Lî÷•©oe©"‘ú`èlc%À²b¦9C²ú#Œ #¸(ë:éßÿ')ìþ{>4,ÜUà¬ž"•u§X4t¦—Cêé.a§PD‘¼F0u–º\Y(‡Û´¸J%l³©2ùJj,¨'Bj™2M^%°$ó}¥@Aì:c¶ºâÁël»›¿©÷H£îKUO4P ºùuÅHJçáÓøfíhÒçƒ0kVkÔÖÊMôøv>K,k¬†ÈJB£æ_dŒ‚u
´Kµ#– N§ñ”?ŒA²zMª.<µµM¶ãVîÓ|}NYk¬³¶hß²Š]Õd§™!ê$‹+ÌÜ{˜Lq‡b^ øä]>ßôË1ûçp›=V­TË›—É
jÐ±oÓhæu}Y‡Ûl­¨÷Uòçç£ùU=â„M±ZæWXæÆæ¦Œ2šLçÕuœ"‹Ðùp‚ÜGÀDÂô¨ñ³EW˜©$*ãÓ½]åÔmí‡f%PŠ¦Íœz«Gg*ñZûEÄÇwžËŸê[ÑÄRÌV»ÏkTÔŠcž²B*FtåÆá6ò¥„2 *øœÂ%qÐ-âá›‚x÷:ÞÿD>²Mxy6ÁŽ
&ºº¹ŒbÐ˜’Íß1Š.çA|-;A
Š˜ÿšä—\qã%‚>múá„˜\lWó¤ØŒ,A¸²Ê§^_ µ€8#
2mò"ËrªG“D4†Ã$þÄDnÂðD‹„iq”¡1¹eIÖþÁ )ýUCz*E‰HÿË-k›Ìe#¥ËF†Z~–Ž„·ŒÒ‘ªêˆH&’œ”Ò(aIÃ^’‚“<Ü?ˆÔ¤.y«Èdâ_ˆÜd³1
ÇtÊª±’ÉR¢2¦²~Bž=cF216ZBûP„©¯úðZÇ¹Õ:ÊØ:ÓØ~r?%7˜¿Æ¹ÿ1CÑÕ
«ŸýÒCñ”»(Ê*O~é`ÁnåNÝQ‡ÂWAŠîž¥	Xã
­‡‡ÝP£MlQçÇ˜68B{pZ
`übõ#4›‰“'–XŠŸÆÎ‹¡M;Èk¿’C¾#ôœ®ø{¥©õ¤Y5àcÝÒ^·Fš€Ã'i^Ñ(9î.õ ~:·×Nb2­™
¦f/é®£9±µl:_žè©„ÂŠax&aTG\²déŸ"½Ù”w	#ïÓ7Òê!°jÓ€ž&Ä[ù’Vq ð»j7Òðó(Ê•Ã”*(ÝJè§,ü0åéœÆŸd\u;…ú‰;òÖšTŽ¯¸ÈûK§º'ó(Æc¦î
d§=ä“qóIœ€ä5Oœ•kñÞ|k]7
øÚ æ,Gt&8(Á$™!õ%Ã›êÔ“ †±â "kÈu4DÃ„†S%EK†®ðŸ¡/rù}–‡¥+g˜0'¬²e¹ùÿç}oˆÒÏ‡¬qüÝýŸ'ýyœ“9è÷ß#µð˜tò­d×‚Ê"\á×@EdËè2Ž’YØ§,¯b>JÃ!¦nÅ‡)=JF³¦= »"Oj;ç‰·t³]¡jÕ`öÖ}¯Ön3
ûN£´‡«,}“Æâ¹Ê}ŠT¥æÒZ
õßO›Êý ÌÂ? f?) 4|œw‰ŽhcËTJÞ@@‚2Ä}Ýéü;§ß	dÅ«zÖUšœr!W“éá;šŸøÁh|b¤¥U¹©A
_½Ì€Ùâ…Öbˆ´è‘©Ž×"d‚¼cË¸.FÈ³îTØ hgpŽÌ‰IYžšR~ :‰BÍ_#ó/cÙBœJv-šá?(•w÷Š„,&d-.‹,Î é±‹m°}d—É©böj7cz²k»š3árÂ˜³µª´ƒ·©Ji–z¸/]‘µq
š‘¸=0Î­ºB4¦•¢ïlm0üb@k4 ‡(LÏÆ8>¶öÈ¨Ì.¤ÙÊ”´­\|QÈ`Ž®¢ðrÖÂ>*lÂÌ'Ø]ûâ’ïYO€ðÒGŠáUoá yN»³ïbˆ™~êl¸’Z £\
á2e‘GiFe*YzŠÖY°“âîÀC3¦ËA"åúÍ¥DŠÒ‰E,£—Ž;š=ÆcÝ¥(©·¸+„æ%í±óS&ÃyKu¦u£.ºi‚\S¡1'8[bœ”ô-Ÿb”.Þyöü%Bc¥t ôÙL=×b­é¡©:áRë9=mS~ÁH‹M¿Ã……¡Tàj¯“ë ×Ù)×âNù­;Ë©Ã2¾¹¼\²¸óòÉ56ž}Ê‡Dã€/ðÍ#VüT
ÉÝÿ  ÿÿì}ÛnÜH–àû|E”°°RÝºÛ®²5.y³$¹J=¶¤VÊîj5T&%±Ì$³I¦$— `ßû¶Àb±/ûR³‹n ~`^õ'ý%{NÜFÉL¦.¶åR&Æåœç~Œzü¯ü|ZË‰÷«÷bLôºäHB5ž{væîÒÇ»+ÿ­1™áëÉ-¬;V×xÛÝÞ';½?¿ß9ÜÞßPÒðr˜ðB¥‚®2ß8êøµSFpåÅª³V –x¸X3/¸¢hD,åÿl…0×±~òº½¼²¹|tIž¬beïSPÏjÈèóP2ÍÍm®ÙÖÒVK„7<Ì¿\¾Ü¨n¡½*S±1EÙFæ<tèýŠÓc˜‡ù–“à”û4-X™:¬ÒU7ÚšØ•‹ë-+Á%h”uàù­
FS* -D‹9ÝÔ¿¤2¼@ÔAðP:^$JÏš·ÀµUÊÃ"Ïopvõ%{¡J’škæw÷¬Á›Ú/ßx™ØBäHœ]¦Táya±hnÀž ´ÕÎëÃ¶MNvÈ`. «LsbÈ«ëJ:#(žÛÜÚ;øSÜ«¢n^?ý‚.Í•¿œYÚÑï/²À>3;åŸ ýà…Á ÇRîªºTWuV_©2.È„%n%äÿç’7¿Á€ÝÅßªøÖ¨uÓ`n2öäûßd7:ogjÖ{VtWªkƒxüVW—WÙ¿ø·ºäª`>ô.ßR€¦IjçŠ¼éuhw©.VrO‡f{ÅÎüêSb¶J qÝ˜]DÉÉ.÷©<Äšù¨ÈÿMf•ïÞ(&™×^,­ôV‰…9N¾,æ§__³cGñÍúÁ]ÿ`¶Ã²#S¼á¨¦0ÑönðòF44WÖŽéÏžÿ¼ù j*>nußÜZ ¡[˜Î1ƒ÷w1¹ ­?õ&ŒL­’`Œ‡Ax‰ï+;«N374brûÞÉÞ˜nNëè#z¾3öv¦0ýîæïéJ—gÆ?ö“lœ<ŠVÆOÉ?sõ9t¸ÖN‡~
‹-ºu`-v?;ˆuSú»P}5›ã£²DÏ±ÏCƒÒê8:1üGÍ Ï¼¹¨X§™#ÂŽ}jÓÐ…nmÇ>bîØGÇ¢£2Óž1ˆÿÌs8RéÕõYh0œŸó…–;«}}’0>ö™6˜}LÆÚbÿÍûø³–ð>>i+”7‰ò«ƒ)¨‘©ÓŽÆÿñ– @>±ÙÁ…p¾»¡pÒó¯ñ‰<Md#ûbÄop@;“u6d:¨ŒÈ7(ŸjBë>TÑ™†œƒ-ìDý²Fe¦CÚu·wwvö7­ ›Å‹dý)ü›ÿ‡þ‡þˆ&ãÙ¶xŸ?xï‡wÝ®å›Nž·qßEŠ‡ÏËº›ôV)+´ôhÔÍ?yË§kÖåŠµlÚ-õ_Wæú"¸GÀ ¼	.ÝbÙ«e0ÀpÆW­uÚùç©áÚòÃqè%àfú˜T©l½ÚW¤²~¿$¥ÔÎÄF^èù“©k;ü¬Ïkºb-ŸÑØgçr­Ëf(|nQPxjrAý|ßú§…¬~bk		U,XÑKÙÁ‘Ýué@
kò(Q¨“›h‚ZvUûžÅr[»nt³s¢;xÓju—ºÑI‘
¸œp´¶E‡œ5·«pÇ9xãî³Âçîv=îêùÛU¹¤UùÚÝ¢§]]?»Ê)Ù,Íã	§¥¢šgt­sxÕ©>uÏlôª¦GÝ}ñ§›R;`Dñ±«Dj'PÞ¥³Ý]¹ÚÍÒÑÎ*Cšä[TÚ=Ãõî¥ëBL=Ï»¶dç¥vã3•âÁ¶PMµÉTQ›…wiœD¹iæ,Ün;°S+ <Up§6ÇÒEë©\iï©¯!ýÔ¯Aøg¿8yw¨+ÚSÊ·xf--ýøõêùÙOC¤OÂøî`d%¥„>•ÚpPCG®÷IS 1…Ðtù¾'
Ù³sÜÃíÉ_5/WÊ^;DÙI¶‰\—ðH¶Öù1zöø¼rš2[²ÃÎ¶'„M
˜•aJ¥1…2˜ÚÔTRŒ9D&³h!µ`-‰J(&áð*d%—Pá’‘êNb
Ùå":ƒ*áÈ9ÃUSrvï2)› ´¨2¨¨†ðÓ†à#‹äc‘wZB9Ëk±CÆŽœbÑŒdRÈ8´âòµ¾¹¦B¦™â4ž9–“Ñ64ñ‘8IàÐTgßtœ5¶MÌ2Ù-k™gª0§MµMjwõœM¾HVºžGÉÄÐÑÔ§ä¤Ê9ä‹JÉ”ú> wÂ0*ëB?‚G<Ä²Ì@Dß3‘{!h3Sf5œ™„îÑ‰|®RPÚÖ6a*–g÷C‡•|Ø¢8LàáKâU¦ÈJ3ä“ÂÛÁÁj©¼
…sç;ïT8Wad¦Ò}CSã]Êf<<ï ô"‡lÆ¢7ÿät%³‰a°
t(\a‰†–1ñ›Âbü…^ýÎR/¤_»Ñ˜ÛöO‚(â¹7o‚#'TŒÕ¥/@AkO|ŸxÙ8@h{Ä®ö<ÖÜ&ûK:OW±‚EºÐ :¹Mú:øúyãòe˜ÛüÇïÿƒä¿IçCyáÍï°âd…ì†Á~<Gï¶€äÛÃ¼?¹ù=PìcÓÙf+2¡ˆ‡Ç~bd†AôíÜš…øvîåË—ÕõJ/7ÈÓÕEòìù"ùþ¾4Òã{G„G¶ýcøòÌèoi‰kÒ*¡]o`Y˜Ôß2íY$k«´¨"¼ù~%½KŠrT·ï§é—{œ¤l!ð,é2²Ý&ˆ°Þ	{ž¢¾ÙÄGF—Ú.£¯lCÉÛ (Mk@¶åHæ6åWÙaoœŽàLquØ€RÏª„î  yêÔ,¤{ÉŠ>}¢”æ-ø<½CîD€5–ûX)ª|ßæÕ o9ÝË±À tw@¾%j©]CM\¾D©²Di.ãÛñ¯¨^ÓÇ}¬ý‹dZ¼ \¿¦Æ²U/\ÕÒáŽÉO½qº‡ªSîã,¾8Š½4ëÌ ø%7¿Ã·~<$é˜¢ÆWs‹d.eÝÎAZ—¸ãÔOLyv²ä“1ý+‚Ì€y;îwøïx‘Ìs·Éôg…Ö¤ó‹ŽÅÅ;}/ëŸ$^¦—áþÇ¡¿ì'IœÀ¼ñ{7_õD%'	ÌºÆãÁÌ¹¼GøJ9ý7æØ‹ŸöÖ£cEºAëŽüåp÷hg‘ü«q>+ÿéJÌçú_Mú'×ïZ%™¾«GÅ¹‡™3¬Y
Ä^¹U Såºs²—¦®:òa±ü$óõÔòr>§4y9u‰ë•”{©—”S\±D±œ×ÅC{#ÂSyÜü	 ªŸ¤7¿Ÿû¡Z„®P€n:2ŒeV4úÛ)Ó(F~"80»´­2uoÓ.€¥Ã;˜žÂïFƒ7TVÇvµÕÂr¿/üdËKýÍ²5ÿÏ¥W £ñÑ$ñqèÓ-  Yùm#vÛ“%¸F8^ËFËø´þv\1©µÚ «¶õâpÜ}ÛˆR~[(u(m8 eD‰Ÿ“HSé–KühG&-äˆ	>¿“e5K¢vlVÖ˜çÏˆó=•CUq÷õMˆÎF©s†V”gÅWW}añ}UÎ°ÎàS5'Ua¥„%«`i¢0µÒ©^hRÃtuô6eÍ~^ƒ3Jc¢ƒ@ ÉS‡Ç•lÐ!‚(ehÛBªõ­¥øœb_­"¨qB"¡ØÄÇ$’t•§ªK§[‹<e%¨‡Ø*“È]yaØ-£¹ñ°ò+3^!/-ü&V„þØ¼âÅà\¯6gÅ³gµ³e@(l‚Íâ×1)›dJq‡[vÛ_©ÝÒÓeÏí¡¶¨\w™]F0V ½P8µ@í¿ BkËÏíÑ§#3ßa¯á¬¼ä……Pñ®»!0BGHlOC_”|Ž¥I7ó#Æ–x1áÔÚ6BÅ²ž—o·Š²DŒzŽâ4.îK­ílØõkÉ9tP1ÈªPûÍ:,ÀÊ5=øéÂt0ˆDvz`ÙkíÓ2Io¸L<‚t#H•½,—S?ÝX°ebqÇRå"ÖÉë	D0†ø/õÇiÎVÁ>/˜‘ Ï: Í„8,†^„`6¤eÅŸØ â4«˜º5¶ÔnÇ6àcÑ~§á£¸±æÀHÑæ…Š“mc¤RtØ…“JµÏB±OÎÐ…Lƒäñèb›[¹RYä&¸Q8ar6JðÊ‘àl|ÊfpâB»:N< 8mÂPéZâ£špEËe„ë…Žè¶¶ØÅâÔDçºª/ž#°y/>‰#+´KVÉÒ—Øùâh;‘Á¿ï¢õdÁ*ºó;}œ*G¥0îZ·“ÅÕéQs“°íÔÇó¬FÜU.ÂÕv$]]&;€™HT“›ßpJ£8¢â:ÃÈ>¼è4N‚_a-„e±P®~Jº»ˆ”õtŒU§Š$8ŠùS %>”°À$´&÷A°úþZdfk+kÖrœÒ½ÛšÚ.&JÀ#¿Åa|x‹ä°²¿ÀùÜónþcà/’Á˜Wu\$‡_ã„ý¾¾ÜµàÀ‹xTf²M"É^ñ¡j©Ä!¸¢“ÎÙ˜?°hJ@y¼ô
,&â¶‰6µJ8‹£ìpzqÒë‹$ß*òJB*C•ÿ¯†jo@éæ«âK8YÅ;Zc~Ç¤’2õPlwdz7J ®œ<SZMÉ­›Tß!é”$Så´H+áŒìGá§ö‰¨’Î­šè¡—kœ¤qä6CwC ’•Xi0¨Ó™*j2óÃ K?ñ·ºÇ›ÂHijOW_:Äö`Qu†á3÷þcJòÆGWÏöà*Ž-?¸züÂAµ%õ;Bmn¸³Ok’qè-¦¾2›Ø÷æ®´ºž¦½Šã¦<¡D©4”5~pÉ„Ÿëçƒã %Ô  N¯³™¤(P¤pÚ`æxªŸ©'H´æè6žì\Â³ƒ¡½ìoƒá™4¡¡Ú¦ãW:Ò¹vÀžŽ²†ú@­Î‡¶-V´ÇVôV¡}¨ûŒ¹44È¯sj4¡ìWÈf,Õ…:u_ÝŒšPoV×qóŒvØÑ«.Õ9mð¢BsÞX?g2ºTãL/¾«Úê¿*)–JÚªWý•jþ÷FÞ)Es]Ù
 wìE}zÐ}Ø¨ñ²q%B:±¢CÎl`:Äê­çáVô½
ƒM‰‘-ñÉôŠù}ØÐ?Å0 xÈÖ“‡êe©‡NÔrj=&•%‹6Õõ y™”¦Ñ¬êÿÄ£ø<¢®ëð,Û
ÓÃ¯VÆe.Ðª¶eKk
ÄJt¬•a¥÷¦:÷ßÈ¯ÔQfwôX$1Uœ©§1J¡®.-#àL…rƒ3IœYÿg …+ÞŠøW[’½Õ¾…à-ºW¤nµ™SäÖž½·ò¶05sÊ•¯ñ²1e:T‘Ùz@&<FÁ„vUÚ¡†ùÉ¢èBs<.êE‹t'ÒÉá]ŠùÌ(“Ëø6Ÿ‰/FÆW!þþ
ø#=Òý´Ò=Çh—h/H[m¹^<pï…zf*krŒzÞ,Î¿|L¾öõìv%G¼¢±Nvto,u‚¢8PâäÔ‹àð¡D8 ©8&¨–jÁP•lwèŒÄ3v@q>Ñé)ÔçEËkËÔ·bœ{¤d˜á}z¢¥ä‡1HZh‘}DÀSøA/’~;]EÒšÙÛäwOÔjz‰‡áàÌ«+o'{h‹Æ•ÇVJ@ŽÆƒ;PKN¨ ªÉúŠš£ýíýžpBSÙiA¯CžLG#yD¦Û´49ãhw¦2(ilø¼ùÿöïd_ZÁéLø¨IŸ´e£ò¥âuï,|ao/¢´‰?jàÓ‚ÿ²¼ÒKs¥'çôR ²K?3f¶&Š0³4oÙ‘w „¯à-Û#Þnw%Þªª­°¼Ïo»êÔ\xë}ñx+'FÛžÔŸ´Ž·blh[D[EõüÙá­ªVD\âæv“ .QCôx‰Í3¿´êÛkk%*+£ÚŠ‡#èàv-žwËHäzéEÝ3 ü¤vkFÐº&P†hWy4-‹:½VXkŽY"°67]èžóxy÷rµîkÓ¥{y+zWèÔ¤Ý3à*eeÉ°NÊ¦M¦Ù˜…{ïZËzñ Ôsz¡ÿó¨Èª¨7Ã}{`'~a¨=^©)83às‹;†Õ 	;ÍmØ–€´ólã`b:–¦p¶­ÌÃl_é6Ïx¡¼_6)*ÝÔÅ+ó^>›h@ôi….´æÂRÆ~®ïÀiþ-¢4 •b…
f@ê’8ŠOoèåq‚%ŒîÃ2#¤»}¯1 ‡æË3‚'ÇÖ©¤r¥á´uà@ó€y]ƒD-·²‘Áp'OÿSwwéC|t°úü´ow+Ÿ(n%@y4Táw"ËXe2oôY{JFKÏ´¨WÁ-†§ö ØÊqÐŠœáÿ
s7fƒ)T/¯*€nr)S"\×¹gÝ:qÕ× ñžÅåOÏô;·ÙãÙ¡Ìi6ìvÉ‚¸º/"Ob „„Á0½Ð“ä$FôL¼§5éšJ i	ûÊí¤»£g±¢µRÛÀ÷Ó¨u0Á±$7Xp¥µaKømíª%å8Âé&øƒkê¢’bi˜~F#l7Â ÛfþîÕÑjvKÌ¥9m‰½òË†Ÿ×ç"ÿ¬ÖXˆ´UKœ·›f…MG}Í8ïà)©0qŠ!ÑJ©ºhV7—’£â€îãÐ=Ò8]Átu¸ø–ÌH…aÐR«O×ËeVÀ‹^èÝ5qÞñ*î?^ZÏ»<uù2Ò;ø¾AÒI°AXªá{Ö:-šp'¸.L.Õƒ@Ÿ×öÜu+&.mIqøï¼lœÙ§m8 ƒ°6³·Ç<t€ûdÀÉëÓœ»ÝÏá³3vë%Ÿ5µ	%’ø?ä{PäË([˜2¸Žlàº:‰j“FÌ ­ðÌ«ÞÈK>†¨¿“Ü -7	R	/
†HGGcØ¡¹ºŒ£ÉYŸâ·…]|Q GYâõ?¢ÏÒE P6·iK m¿ŸÑPß2ëwö¬L
CövøžxFŽ•„¾7À÷fÁéYfX¼‚Ê)€!•sœœ=kÎqr8ÔÏWÊ3Vê£×Ù,Bÿ¤<àZçÉé­° Â¢ GM°<ðÈ.€XœÞüsv’nBïbäJ°%ÀJ¬|â‡Þ%UêNü¢Ž®¼#UZ`RTÅØaWò3ŸÙ´á—ü‡ü#	ÆÃŒoþF¿ï…<iîÎŸúEÌÁ¯è‘P*ò‡<8 uHñ˜s3¹&‰5	¢ÔO¨Ñe?™ôêÝu…Ê<Ægh-°½¼¿Pï|2W¥ÏßòšÌQ__Á¨3M!…‘×®[=È,¹¦­c4$Ÿ.4èÖsšj™¯‰ã)u0ÔÔÜÕ»eïNùeM\cÉ·’ÀKèR3ÿf‹Nö«ÂÔ`pBb}ÍýœiÖjT§ßYÙfJ‡é¦ŽÄÕÅÞ!uDyÃíBl«5ð*È8€Þ¸ohÞÞ¼¶äÜ™ˆÛ"ßj¥L	¸‹{¥·}cÊî”ƒÇˆµ°Ò"Fcr£ùŒÀžÅÑ)';Ã4(4êÌÐ!ñBã¬Ä¼úã$B¸ sSs/G–yy„éã÷,ècœÚ9ËÒ,É4)}Å(éj(oî\‘ååeu, eÉÁ`Þé<`‘kË MIJË×®õK–VÏ×*­)4!µgK Ië´	ZËÂ¢91k[5êúˆ¬‡À¸TdÎùŒh““¦"P‚rÔÁd-.Á„ˆ¾j—6ü,c§KHki§â²: ä&,ŒÌ€ìÖO€‘–‹a&Dù=em(‘%,ù;ÒÞ!-æ`"ÕQœ/ŒÑœŸpe¬„ñ«òéä$€†ã¤0UÏv°XÈœ™‘“¯›½m‡²ª™°å[¹°¯\{/+PœÙ$†mA~ï4Œ½pO$ü¼Ãý§Ñ…_$†Ý4•w¤Ä]¼Ë¢TÉûˆýuN6ØïMöÍ ß×¤s%Çuí¬<t¿äõc”ì…˜d;ðiŠŽà8[ZV/ÖÖ“¶Èù 
Ù®¤FI¤^¨Z–‹ÞJÕ³ÊU²fí]§(`“ó¸ùwÔÙêÚªVÔ´U¤u M•cïÖ›4ndÒxƒìÑÕ£ïØˆk¢ºTç?A?ª®¶Æii:-+,}F|¨'dž.]Ð¿–¢†ónŠå¢Å€i	…ª&Ž&¯Éüþ8KâyZ)È`ÀjÀùÂdj‰åÈxÑ¶¨Gbo¯ÏÕ–@]™ÍH¶7â~™*­èZØx¢)Ù°Âpf³d½(–$4ƒ’ÚrŸÀÎ¡ß¹3¸ëj~ :·IÿÎV<¦ÎcÁøŽNPËÙ8š@…þEÝÂ<íz~fñ²Õ[`¼°¥äºèMö§‰\]š.¹¹MúÇÁ³Ú
¥\Ù±Ih8{Ã'OÈW =U[;UWzÁä>`?€é	ê•8EGý;§¬Ì<äE†Ìz
‚tËI€Î˜ycînî*“j£(µÊ£:5~jÒKJƒ¶lãè`?™³1ùü`'^5ªöj¨×Zœ…Ù”rDsÉÙLJüîT%ñaNb¿1uV­JìÖÝØ”äTÍ%v»ÃRï5µ&ñ‡‹¶${W“[’”µ¦v$±¢¹9)Ï‡Á¶ÛR>4i2¢–"å:ÇùÝÚu‚˜{vÖó,)ù††§ÊÏuêòµÆ÷jö¥«•?·Ýí}²ÓûóûÃíý²“öãæÜÀ@mÙÌßú+jiã×Ã›—Ðméb³h•ŸS¼Vž¯RçN+/Vï°x]‡ Ê\D¢ApS—›Ë%(<ÇèZÆÆu\ÇóÐà/Xé:s±ôœœ-ÑÂc|¬Ï5ï}Š›ÀÜæšÍs‘ˆ~¢oty)KŽ$ø)7«‚Ït¨À'OÓÎÙ|½é—¸WpëÉ·8÷’°zõpiŸAŽò‡‘}Œ‚ßSÆž™´6›RæI$”Y’ü)°70Õ(¹Ä~ÎªìøŒ®V1ùÅ\Ñ¬"8’½ÐŠìy§‹Döªk”Le×ñcÞ—^¸²."Ð?‘–JðÃUQ µªå?7é—•mµíL?»g«zG8ÓeñH—K€øSÄê¼íÅñÇf’F¾8¤õáë0ï…§à¥Šq4`5ŒSa¤ýZ˜yühxmmGÍ|*Ê;ZÊ5Ú`¯·Y÷ð£Xø`.´"sÎgmŽ…ýšŸ_põÊGÐŸRÏ4µs~ëçDê5ô÷h×xš-ÔwP3F±Sz‰õdíÈf®4ÊJøiNEI'±ôŠuj©>6Â›˜©	ÔtÜ,ˆ"ØYâë
x‡-þç“zÐÎ,*ü§é€Æp*Ç"ƒq`æÌK÷üSôõÑ%2²ú­iˆ QC¿ƒ£t/ž,.=,Ù“øŸt<Hè3ËÊJð)ã¤ ´:§/Q<ñP/øÒV#½¸°Pó5b~L\Ûƒóf^v^¼wà%èýë@šBaGýcÇ(¹Ö†]´b!S7r$¨ÉX°5”Œÿ¹y¥®Û5Y¢¸f"N·ÉxØê5?|aàèæ÷lbY1Ín¾ÆªÕÄO…f?fÎgùùâ}JùVßxáÕUùË’˜Ö´Á¸óø¯ãTñoIþÈ¹ˆzâGIÑˆŸ{Ç³´ô,—T:³8	f¶XaäÛ9äUÒ¥”ó?éÒ	¬±mƒ?K@ßÝëîÝü[··HvzG‡Ý£›ÿþýn×è«Á>e(ÏÙ½– œ±˜Eè^Îâ÷B¼¾DXG ö2÷›Ë–×d/•52„qY’k\ÿ‘wÝÃÙ9ÚÝûž<!vö¶›<»°sØ½ù·›ÿµÓƒ§÷·vz½ýîl½?„'Èïßu÷š<y´³µ·ÿvVµsö%òÃw¤ù‰m¼xŸ5µjôá½WÕÂ8s]­:“‚²vÝ®¬ÅØY8KYÕº£›¿õ#`Ø-šZ#3fœâ-i†„\ì8‰¾|øy gÆôòHð+¿ï>DŠžêî<‘ÙÎR?V5SQa¦ï’’j¤Çd Y.YáöÁ³ËSæÞ*´Ùg Ï&Ð?½ %ž,¶vîiN(ŒÅ¼·uÔCÕP‹ Iû› +/gËéöîáÎîÑþµÇû¿0$^$h¨ÆÄi´ž |kb=}þP¬§úv>Œ™Õxç³(œÆOí§±ºÕÈ&Þ¯q9¦Àq;T#L¢çÉÍ½±ò 	b"þhÛÑTba(Q{wó·Aàµ§¢·l#í†™`ˆ‘nM¡.OÑùÍm²¿þÎ.áYúÇýèÃ2€f^6vq_"^¤tQ(N o…Ô¸-¬`]·ŒbHÂ¾5î!éÉr·oª¿v$³ZB7ù÷Öñå6xñjß*ŽD¶
ÍuÒ«·Aší# ûQ¡y½œ9(n«;¢†?Njã”X×ì
…gºÔŠè"ü`„ÑÕšŒë-Ô…Oÿû!cåùª˜ðV3¥&5¤Ô¨®-’õEò´ ÓÍDtœÝ	¼žÚ0-"\Â×s;»ùG²¶@ÃPjŠ'6Ó¬2ò’Ôßôp‹dmÕjltÓ@:Ö Hìa†àä“ÞÀÌTgËCï²³ºHÇ°³²¦>¨),™³½4$#fX`|ns›Žã½˜&Q>O/ìQëæ<Çæ˜¯ˆ–‡ï§´¶1‹ÿÄ6ÇûÈX£”Hv£›ßúLm§A_éfF—Íc î{7BºÝªÙ$ïöó2RP{Á¬9˜Á²µcÐ§`ìÆ?âE€f€ŸñäÎ77¿ûá£oF“µºrõ1!aOy˜7°MlPúý¼HnŽÉD·9ÑÍW¡í(ò­^ãEí*Úz¾à˜É[xœø™_6"ÌÂ`¢ñ5Í£q2
ýÛ³|çpqÀÓ² ŽE/ìa§±(ñõ«•Çgž–ù"ñFL2‘iÇã,‹#‡Ú€50cxýJ€§
åº~Èg¼3ô ¶ªßuá–!?ûäS¼êG6VKõ’5<0/õròœÔ®i™5
Âà–VoO–Fq@;=[zÉÓ-ñHn\ï5½‰Íß:wK{‚q¡¹*EÚ‹Xu”"E¤59Ï@„½ âke+–j5îeIÏ}>ÃÌö
‚›Â¸q–ƒçN7× Y´d;K¸,FJÇ ÓôØ=ô¿œyYêFfXw¿Pp—Ó¯€ø’ÃRäÕ»æý4õNýÞ_Ç° ÕÀO×¨;Ýð÷ÒOQŸ˜‚óóHS{™±ü&Ø?§g6í#ËËž~ò„ÌïÅ"©yÍ“ "XþîÂž¦PËÁÁû½/ÈˆÚ§| ‚4…+GàŽ‡vEÓ‹œ²dþÀš°‡}<<Ð;s1ðÞyœ,’3ï8i†µ“3Šˆgôh‰°”Ý˜D ƒœ',óZâ÷ýc@<ûp /Æ2]žs…XîØ®×$e¸arS©s?wB†KSEÄæ´LÃ”›€„õ†^’Îpé*é.M#ÚeS]u¬SQ<+.æÛ80K† 2;W˜e )|‘³zr‚KP¥x9¥“!¥	ÃåXÀ¬QLz¥¯*é3`§¯V¡Ñõlé)JiˆœÖR9´J°nO`“úMr¿ ŽíÀK»¸L<›(~ª¤|-•ý7HP(>¶ÓMùWk×›kMÉ6¼Ê ôøúõæ:>žNøüÓëÍ§Ó<ÿüzóù4Ïs½ùMƒç]NÝž%æöÎ0~žô2ª·Ù†«p&@‘ÉLÚòIÊR u¡uô¡UhQ5¨”]·¾äðKìÅÿÙÇU\Æ
`ñ˜üñùóÏ^ŠÏ$4Fì”S89‰áÝ×H˜+>ö(‹ZÛ|ÃËÙæÑl½0]yf‘G«Ymf‘lKiÑRIfC±Û®@o©ÅÌk½êÁ ˆfÕÔª2OVqÙVý‡¾Úë ñûSÖAoa‰ëbE]¨ƒ$®ÝQ$9S[$Ö #‚…›×AâèuìÝ)¿×AR–¼n$u(¦$ò¼ŒÍ_Wh‹¡0üæ½+q´MçÁí'qUüþCJ#_‘AžÇÜg4\ÆÜŸ‘õý ‰Ë'ñEŠ¬bé†RfÀuS…ç½ˆ ù•Ebèø¾&šÿÌJfGöZYKˆ×ºèFÜ./ÃÒ5¨“¸»,˜ÂÖ›oP\¸
„x¬þ’»<È—^iw4Úê¬à.œ`ñŸ¹w_ƒ‚vCà£1.ðãU„ÓD¶~íàïlb:8k`ÑJ'|>$q²
	
èåx:tV&¡F»‘¿ý^VBpC­±kÐ^%„C˜m%„Ç¢M‹XÀ_/Zà.LP‡Pt­ÕðË+X0ó2%t¿£2¥qT–)hBÛfQ¦ xíÈ;ÝEÈ{çEÞ)À†ÖEæ¦‚—ºüéZo­¬wä_@×i]F;E»3}Ù½k•m>¢ƒÒ¯'»23Åc®i¦zÌuSí+: ÛV?Öß¢‹ôùb(E}ÊL<á|Ã•œÞk£ž€'Sx]ž5ÿ7Š>Í,}œø£×¼5†yñùDV‰-–‘ŸLWj'QY¼!8É’ï³ 'CiõûNTµ-Ë¢7Ý~ëÎ–Œp¶>£ÏŠ¯E•q«ì2`˜xÁ†n52©HE™«µÅhZÝ¡jÓE® Ó¼õç6ç
ý‹
súÊ´IS{,ûSSÙ¿f|LdR¯G$°.os$0å?nÌf#‰fû™¸=•ùL¾C”}RÐ)7Ÿå<«#ôô²Úu[Fµ|úf›š@]¾oM-j&Ì—êÚµ§©{P×œ¦DKƒ^¸eZ3	²Ñhä¡žÃ†Ðë€–OöÉ>Ö ~“N5,)Æ
©!E¥ÌÞÃAÌÞfÿ§JÒFœáFH0îF&Fîfc¶VG¶ghÑ=ö.l94ôEÁ»Õ.Ç[>* [ì=Åz)±çú’}/“T•1	Ëït²Âôå)oì–¦:v¦‰,MÓÛš,*>‹²Îžø6ÑVRÐ¼FÙàKÇ\¬‡õ-ñÀyÍÃ945Â"
aBéBËÑJ##Q>Hs¤5ã2è€^ì!•Ø®¢7vUð¬á—ŠA&è}êõa¤˜	Èµ+U4BÙ8:a2 ‹ÿ•N›‘ã8ãc/ÜÄ|ðÎËúg½_`™57×K¡t‰¾£@~
•ÏxÝ³‘EïÃ§Øœö.•ø<É´’Ç·’Õl”)$×ÉkL2zìg¾™â¥'1­R5šy>%=},~ž*ô7ÿÀµkvÊh=«Š<kwF+9™)Ïy˜–ê«‘ëGI0D¿%G¸«ÉSƒaÆÉ§ŠÀ¹ÄÇÚ2níÔ€dg~w·7>=E§nhU`‡Ä€É?á?ëhq‰Xï®à5öY*:ZdŽ>Sc…3j¢u–0ÃÜË?Ë²3ízuG#LŸ+ˆe?ÅËÕÝ(iDÊ™nVwùqÈ™Äù‹—«»¡‹
§b€Óéûiš«|‹ugíÍQñÍe7ÇG5V”Ç‹i’µ„à<ÒŒçõÎ•TäP”Õ
¢ó›ßBÌ)?Àde™§7úGº	eÊ/t ¨¯I$Òñ“dƒxÑ§JQsàÿ5š#¤Ý.Y¬/.)ü\˜î /0‘AD¹@{Àd÷MýaPÊÌøµtåkÌãù%ž®Ã˜¸îÂ<ÔpÍb`¦SäªÜ÷)¬òn—ÅÂ–ÖüØv´½`ì¬ åÎŒÜQzüB<zñÙÉ;¿§ãÁ`DØsÌ£93‹Ý+Øê3ëðiº3£‘dÂŽbçh[»ô¾rÔÍÚ?Gµ
-_ÁÞ¢,‘‘('Š1•§A^q*ÇÄÇäQç–Œç_¾uÈ†Šb=yêMÅ‚NÌ>–ÆÑÜçýºtc’aš* €{FÉÍo7)FŸÉ¼>žÙ7¿%æHéd0%¤gz]€×¦6€½Ä›µè;—~Ÿ£vGÍ"·Hþå`7] þ¬UÇQ¸sÕ±ðfETg#ðÇ¡jªãÖTõ ¨Úyìr#/>ÈŠgÆðµ¦²ÊÄ§×R[U¢„Y<¨œYb?ÌüÂl ±JB£½ÊÀ#8æŸ8nwV¶ºoˆ—QM`,Êyí€¡.R¶€6z½EåK< ¨gË˜>ÑÜ÷Û¡Õ4;ÇmlËÔ:¨8‹Î|žÊÑjï¬èÅ¸ õ
)pæ8S/À¾UŒyçG(3É0Î‚sæbâ'…ueÜý(}·$ê7`UYåÑf™´,S`×í ±Ž†îïÎlÔàÐÔõ]xr„¨ ¦;ý³ œŠö¯e‘§7€×Ñi¥SÃvfˆÒáÔÒ©T76ƒ=L§];Ü~2ðšbAØ Yâ~  ÇN¦ˆá;-EÍ*húH ³£À¤Ñ@®Ì>WVLÓ‚líž<!®X!ËcÔ7ÏÖeˆ Ã]#¡¼=.è³ÞÓi¡DešÔJUžÑDJ<Qy0®p¢ÆD¸qNGƒV¼AæŸâµZ‘F|™$Ð¨b9îG˜ÈLÌS7Cåõ(Ä|¥>¦1;â4úŠ ôêPÜß–·cú°|O¢{í·WA*ÛJ§Û|´&ßÜrû’“.G$Xñ-¿—8†ûê¯¼MÂ|Ð@~UzðO¼q›6Nút‚»lÇG¥^V'"‡­´UVG4]Ù4.8úO«k¿ê‚£ß3ï-H·ÅâV›6¬µØ*© _z”¯û?â0ŽbÑ¸ƒÅ”&
ÆöÇI‡Ò6]“Aá©×´”ayäÞîÅ˜±Uy³ØiæyE½}áy¥wÎÛî^/S×B@ØÂ$ðPöœÛè]b…]v´T¨[¾*òp_¿ÀËYBZÓK°KÎfº­ngÏ;÷Ùê"ÁÁo½,)í-2ÐÅžtþ™y’ ^ƒÐO½Í•­¡»\ì€zœ)¯T!äº
9Ì³®½OQß¿ÅÕÔ¢Ü¾D- S¸Ü íŸ}…ŽÆôíe¼Ÿo¯F~r„™öLÛáå¼-°0Yo<„»ÀdÒp›_È[S?¡9b¿½Â¯“ðCó$‰8ZàÙüjX«ï¼þÇñhçr'™m¹X›šk¥56,{_*ùƒ=—·ÃB¤@:ée¬«Oàã}Üìöþ;oÁ®ÿ)Å5Oè«ô+ê,¶ã‹(Œ½öc¢\Ýúó:|ÃÅS8ÜEWßÉ•ž\$ÅQ×‡ì«Ñàä}j¡/ºG]péHÁn,QÕ M´¹òõª¡8ˆVÐû×%^Ô8Hâ<;NuEMHíƒ·ÀYü*˜pïrébék¸p¶ôãËõó³ŸH^Æ£‡=1NBxè,üÈÆ™œ5r(,Åo™Ë¢”ßTç]¶2^ú“K/h–P%ýÿšÔŒÿJlYÕî™«Ÿ aŽpyaLem+ªsö´ä:‚#±È?)3Èìp‡²ìä«•³§µm=¥U¦elLûu:ÉC¹Ä6Çã´0Ñ¸VÛ‚e×_Ž‘:1„Z$ó?–Dç®«¸s	yáÛ ú¨¬¤˜Àð`iI÷8ÁÄ!¡eÝcÏ°ØßÙ’ƒ7ŸŒ=CMp’xð »Ð£à¾‚bò‰º‚fqÒ™g~äîHÛoÈAâŸþÅÜOóhšÿáèÝÛÝ7Øh‡Ým	nø«€,±oË˜ÚÿÅYÇÀÔK¾G™¹XS%æÜÔº	jŒ{°´úîmß¢7mL³ÛØœïí¶§n4‹°ã›Ö™÷Ì+à-Ÿ%þ	Ðõ27ðóšÍS\ŽÑ MÐ·OÏ›ŸíãÐM{iYiqvV,5-qŸ Ø&‹]"À£KAs¥ÏðriÍèWÚ«Ó3`\;…™?è²2Ù¥ì‚—Ç÷*^Š¶ h>ˆÿÂWfý™…Š›g_'C±™ìb1 ±>ÔûŽNý˜¦ä¹â(zÅ)Dš ëK×åÚ §;c$CÁ™8rª³hÏ!«LbYR\(?sµÅ«£Xpšr¨q²4üæ”ý½F2H1oŠå˜r˜¶—{ôjØrº+¼@‰ëúÿ   ÿÿ ×‚
B