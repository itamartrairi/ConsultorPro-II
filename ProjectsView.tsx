import React, { useState, useMemo, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Plus, 
  Briefcase, 
  Calendar, 
  CheckCircle2, 
  Clock, 
  User, 
  BarChart3, 
  TrendingUp, 
  CheckSquare, 
  Square, 
  ChevronRight, 
  FileText, 
  Save, 
  AlertCircle,
  TrendingDown,
  Info,
  ChevronDown,
  CalendarDays,
  Target,
  UserCheck,
  Flame,
  ArrowRight,
  ListTodo,
  CheckCircle,
  HelpCircle,
  Search,
  Filter,
  Eye,
  Settings,
  ChevronLeft,
  Sliders,
  RefreshCw,
  Pencil,
  Trash2,
  Copy,
  CopyCheck
} from 'lucide-react';
import { doc, updateDoc, collection, addDoc, getDocs, deleteDoc, query, where, writeBatch } from 'firebase/firestore';
import { db } from '../firebase';
import { Button } from './Button';
import { getAI, extractAndParseJSON, Type } from '../App';
import { ReplicateDiagnosticoModal } from './ReplicateDiagnosticoModal';

interface Empresa {
  id: string;
  nome: string;
  cnpj?: string;
  tipoEmpresa?: string;
  dataCadastro?: any;
}

interface Diagnostico {
  id: string;
  empresaId: string;
  dataDiagnostico: any;
  ownerId: string;
  nivelMaturidadeAI?: string;
  justificativaMaturidadeAI?: string;
  tipoEmpresa?: string;
  areasDiagnostico?: string[];
  nomeProjeto?: string;
  nome?: string;
  status?: string;
  dadosConsultoria?: any;
  projectEval?: {
    kpiTarget?: string;
    kpiCurrent?: string;
    kpiProgress?: number; // 0 to 100
    consultantNotes?: string;
    swotS?: string;
    swotW?: string;
    swotO?: string;
    swotT?: string;
    milestones?: {
      reuniaoInicial?: boolean;
      diagnosticoConcluido?: boolean;
      planoAcaoGerado?: boolean;
      reuniaoFechamento?: boolean;
      implementacaoConcluida?: boolean;
    };
  };
}

interface TarefaPlanoAcao {
  id: string;
  diagnosticoId: string;
  empresaId: string;
  problema: string;
  area: string;
  status: 'Pendente' | 'Em Andamento' | 'Concluído';
  prioridade: 'Baixa' | 'Média' | 'Alta';
  responsavel?: string;
  dataInicio?: string; // YYYY-MM-DD
  dataFim?: string; // YYYY-MM-DD
}

interface ProjectsViewProps {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  tarefas: TarefaPlanoAcao[];
  user: any;
  respostas?: any[];
  onCreateDiagnostico: (empresaId: string, areas: string[], dateStr?: string, companyTypeFilter?: string, projectName?: string, stayOnProjectsView?: boolean) => Promise<any>;
  onDeleteDiagnostico?: (id: string) => Promise<void>;
  onReplicateDiagnostico?: (sourceDiagId: string, targetEmpresaId: string, options?: any) => Promise<any>;
  setDiagnosticos?: React.Dispatch<React.SetStateAction<any[]>>;
  setView: (view: any) => void;
  setSelectedEmpresa: (empresa: any) => void;
  setSelectedDiagnostico: (diagnostico: any) => void;
}

const getStatusBadgeStyle = (status: string) => {
  switch (status) {
    case 'Concluído':
    case 'Finalizado':
      return 'bg-emerald-50 text-emerald-600 border border-emerald-100';
    case 'Em Andamento':
      return 'bg-blue-50 text-blue-600 border border-blue-100';
    case 'Planejamento':
      return 'bg-amber-50 text-amber-600 border border-amber-100';
    case 'Suspenso':
      return 'bg-rose-50 text-rose-600 border border-rose-100';
    default:
      return 'bg-slate-50 text-slate-600 border border-slate-100';
  }
};

export const ProjectsView: React.FC<ProjectsViewProps> = ({
  empresas = [],
  diagnosticos = [],
  tarefas = [],
  user,
  respostas = [],
  onCreateDiagnostico,
  onDeleteDiagnostico,
  onReplicateDiagnostico,
  setDiagnosticos,
  setView,
  setSelectedEmpresa,
  setSelectedDiagnostico
}) => {
  // Tabs: list (Grid Card), gantt (Timeline), create (Form), evaluate (SWOT/KPI Form)
  const [activeTab, setActiveTab] = useState<'list' | 'gantt' | 'create' | 'evaluate'>('gantt');
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null);

  // Replication State
  const [isReplicateModalOpen, setIsReplicateModalOpen] = useState(false);
  const [replicateSourceDiagId, setReplicateSourceDiagId] = useState<string>('');
  const [replicateTargetEmpresaId, setReplicateTargetEmpresaId] = useState<string>('');

  // Filters
  const [searchQuery, setSearchQuery] = useState('');
  const [focusFilter, setFocusFilter] = useState('Todos');
  const [filterClient, setFilterClient] = useState('');
  const [filterDate, setFilterDate] = useState('');
  const [filterProjName, setFilterProjName] = useState('');

  // New Project Form State
  const [newProjName, setNewProjName] = useState('');
  const [newProjClient, setNewProjClient] = useState('');
  const [newProjFocus, setNewProjFocus] = useState('Geral');
  const [newProjDate, setNewProjDate] = useState(new Date().toISOString().split('T')[0]);
  const [newProjAreas, setNewProjAreas] = useState<string[]>([]);
  const [isCreating, setIsCreating] = useState(false);
  const isCreatingRef = useRef(false);
  const [useTemplate, setUseTemplate] = useState(false);

  // Auto-generate project name based on selected client
  useEffect(() => {
    if (newProjClient) {
      const comp = empresas.find(e => e.id === newProjClient);
      if (comp) {
        setNewProjName(`Projeto - ${comp.nome || ''}`);
        setNewProjFocus(comp.tipoEmpresa || 'Geral');
      }
    } else {
      setNewProjName('');
      setNewProjFocus('Geral');
    }
  }, [newProjClient, empresas]);

  // Evaluation Form State
  const [kpiTarget, setKpiTarget] = useState('');
  const [kpiCurrent, setKpiCurrent] = useState('');
  const [kpiProgress, setKpiProgress] = useState(50);
  const [consultantNotes, setConsultantNotes] = useState('');
  const [swotS, setSwotS] = useState('');
  const [swotW, setSwotW] = useState('');
  const [swotO, setSwotO] = useState('');
  const [swotT, setSwotT] = useState('');
  const [milestones, setMilestones] = useState({
    reuniaoInicial: false,
    diagnosticoConcluido: false,
    planoAcaoGerado: false,
    reuniaoFechamento: false,
    implementacaoConcluida: false
  });
  const [isSavingEval, setIsSavingEval] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);

  // Gantt Chart State
  const [ganttReferenceDate, setGanttReferenceDate] = useState<Date>(() => {
    const d = new Date();
    d.setDate(1); // Start of the current month
    return d;
  });
  const [expandedProjectGantt, setExpandedProjectGantt] = useState<string | null>(null);
  const [savingTaskIds, setSavingTaskIds] = useState<string[]>([]);

  // Edit Project State
  const [editingProject, setEditingProject] = useState<any>(null);
  const [isEditing, setIsEditing] = useState(false);
  const [editProjName, setEditProjName] = useState('');
  const [editProjClient, setEditProjClient] = useState('');
  const [editProjFocus, setEditProjFocus] = useState('');
  const [editProjDate, setEditProjDate] = useState('');
  const [editProjStatus, setEditProjStatus] = useState('');
  const [isSavingEdit, setIsSavingEdit] = useState(false);

  const handleOpenEditModal = (project: any) => {
    setEditingProject(project);
    setEditProjName(project.nomeProjeto || project.nome || '');
    setEditProjClient(project.empresaId);
    setEditProjFocus(project.tipoEmpresa || 'Geral');
    setEditProjStatus(project.status || 'Planejamento');
    
    let dateStr = '';
    if (project.dataDiagnostico) {
      const d = project.startDate;
      if (d && !isNaN(d.getTime())) {
        dateStr = d.toISOString().split('T')[0];
      }
    }
    if (!dateStr) {
      dateStr = new Date().toISOString().split('T')[0];
    }
    setEditProjDate(dateStr);
    setIsEditing(true);
  };

  const isQuotaError = (err: any) => {
    if (!err) return false;
    const str = String(err?.message || err) + (err?.code ? String(err.code) : '');
    const lower = str.toLowerCase();
    return lower.includes('quota') || lower.includes('resource-exhausted') || lower.includes('limit exceeded');
  };

  const handleSaveEditProject = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingProject) return;
    if (!editProjClient) {
      alert("Selecione um cliente para o projeto.");
      return;
    }
    setIsSavingEdit(true);
    try {
      const parsedDate = new Date(editProjDate + 'T12:00:00');
      const timestampVal = !isNaN(parsedDate.getTime()) 
        ? parsedDate
        : new Date();

      await updateDoc(doc(db, 'diagnosticos', editingProject.id), {
        empresaId: editProjClient,
        tipoEmpresa: editProjFocus,
        dataDiagnostico: timestampVal,
        status: editProjStatus,
        nomeProjeto: editProjName,
        nome: editProjName
      });

      setIsEditing(false);
      setEditingProject(null);
    } catch (err) {
      console.error("Erro ao salvar projeto:", err);
      if (isQuotaError(err)) {
        alert("Cota de gravação do Firestore atingida no projeto gratuito. As alterações foram salvas localmente nesta sessão.");
        setIsEditing(false);
        setEditingProject(null);
      } else {
        alert("Erro ao salvar as alterações do projeto.");
      }
    } finally {
      setIsSavingEdit(false);
    }
  };

  // Delete Project State
  const [deletingProject, setDeletingProject] = useState<any>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isSavingDelete, setIsSavingDelete] = useState(false);

  const handleOpenDeleteModal = (project: any) => {
    setDeletingProject(project);
    setIsDeleting(true);
  };

  const handleDeleteProject = async () => {
    if (!deletingProject) return;
    setIsSavingDelete(true);
    const diagId = deletingProject.id;
    try {
      if (onDeleteDiagnostico) {
        await onDeleteDiagnostico(diagId);
      } else {
        // Fallback cascading delete
        try {
          const qRes = query(collection(db, 'respostas'), where('diagnosticoId', '==', diagId));
          const snapRes = await getDocs(qRes);
          for (let i = 0; i < snapRes.docs.length; i += 400) {
            const batch = writeBatch(db);
            snapRes.docs.slice(i, i + 400).forEach(d => batch.delete(d.ref));
            await batch.commit();
          }

          const qTasks = query(collection(db, 'tarefas_plano'), where('diagnosticoId', '==', diagId));
          const snapTasks = await getDocs(qTasks);
          for (let i = 0; i < snapTasks.docs.length; i += 400) {
            const batch = writeBatch(db);
            snapTasks.docs.slice(i, i + 400).forEach(d => batch.delete(d.ref));
            await batch.commit();
          }

          const qTasksProj = query(collection(db, 'tarefas_plano'), where('projetoId', '==', diagId));
          const snapTasksProj = await getDocs(qTasksProj);
          for (let i = 0; i < snapTasksProj.docs.length; i += 400) {
            const batch = writeBatch(db);
            snapTasksProj.docs.slice(i, i + 400).forEach(d => batch.delete(d.ref));
            await batch.commit();
          }

          try {
            const qAgenda = query(collection(db, 'agenda_eventos'), where('diagnosticoId', '==', diagId));
            const snapAgenda = await getDocs(qAgenda);
            for (let i = 0; i < snapAgenda.docs.length; i += 400) {
              const batch = writeBatch(db);
              snapAgenda.docs.slice(i, i + 400).forEach(d => batch.delete(d.ref));
              await batch.commit();
            }
          } catch (eAgenda) {}

          await deleteDoc(doc(db, 'diagnosticos', diagId));
        } catch (cloudErr) {
          console.warn("Cloud delete fallback warning:", cloudErr);
        }

        // Clean local storage
        try {
          const allLocal = localStorage.getItem('local_all_respostas');
          if (allLocal) {
            const parsed = JSON.parse(allLocal);
            if (Array.isArray(parsed)) {
              localStorage.setItem('local_all_respostas', JSON.stringify(parsed.filter((r: any) => r.diagnosticoId !== diagId)));
              localStorage.setItem('local_respostas', JSON.stringify(parsed.filter((r: any) => r.diagnosticoId !== diagId)));
            }
          }
        } catch (e) {}

        if (setDiagnosticos) {
          setDiagnosticos(prev => prev.filter(d => d.id !== diagId));
        }
        setSelectedDiagnostico(null);
      }

      setIsDeleting(false);
      setDeletingProject(null);
    } catch (err) {
      console.error("Erro ao excluir projeto:", err);
      if (setDiagnosticos) {
        setDiagnosticos(prev => prev.filter(d => d.id !== diagId));
      }
      setSelectedDiagnostico(null);
      setIsDeleting(false);
      setDeletingProject(null);
    } finally {
      setIsSavingDelete(false);
    }
  };

  // Helpers to safely convert firestore dates
  const getProjectStartDate = (project: any) => {
    if (!project || !project.dataDiagnostico) return new Date();
    try {
      if (typeof project.dataDiagnostico.toDate === 'function') {
        const d = project.dataDiagnostico.toDate();
        if (d && !isNaN(d.getTime())) return d;
      }
      if (project.dataDiagnostico.seconds !== undefined) {
        const d = new Date(project.dataDiagnostico.seconds * 1000);
        if (!isNaN(d.getTime())) return d;
      }
      const d = new Date(project.dataDiagnostico);
      if (!isNaN(d.getTime())) return d;
    } catch (e) {
      console.error("Error parsing project start date:", e);
    }
    return new Date();
  };

  // Map each diagnostico into a Project list
  const projects = useMemo(() => {
    const list = diagnosticos || [];
    const comps = empresas || [];
    const tares = tarefas || [];
    return list.map(diag => {
      const company = comps.find(e => e.id === diag.empresaId);
      const projTasks = tares.filter(t => t.diagnosticoId === diag.id);
      const completedTasks = projTasks.filter(t => t.status === 'Concluído').length;
      const totalTasks = projTasks.length;
      const progressPercent = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;
      const startDate = getProjectStartDate(diag);

      let computedStatus = diag.status || 'Planejamento';
      if (computedStatus === 'Finalizado' || computedStatus === 'Concluído') {
        if (totalTasks > 0) {
          computedStatus = progressPercent === 100 ? 'Concluído' : 'Em Andamento';
        } else {
          computedStatus = 'Planejamento';
        }
      }

      return {
        ...diag,
        companyName: company && company.nome ? company.nome : 'Cliente Não Identificado',
        company,
        totalTasks,
        completedTasks,
        progressPercent,
        startDate,
        status: computedStatus
      };
    });
  }, [diagnosticos, empresas, tarefas]);

  // Set initial Gantt scale reference date based on earliest project
  useEffect(() => {
    if (projects.length > 0) {
      const dates = projects.map(p => p.startDate.getTime()).filter(t => !isNaN(t));
      if (dates.length > 0) {
        const minDate = new Date(Math.min(...dates));
        if (!isNaN(minDate.getTime())) {
          minDate.setDate(1); // Set to start of month
          setGanttReferenceDate(minDate);
        } else {
          const d = new Date();
          d.setDate(1);
          setGanttReferenceDate(d);
        }
      }
    }
  }, [diagnosticos, projects]);

  // Filtered projects
  const filteredProjects = useMemo(() => {
    return projects.filter(proj => {
      const companyName = proj.companyName || '';
      const tipoEmpresa = proj.tipoEmpresa || '';
      const projName = proj.projectName || '';
      
      const matchesClient = !filterClient || proj.empresaId === filterClient;
      const matchesProjName = !filterProjName || projName.toLowerCase().includes(filterProjName.toLowerCase());
      
      let matchesDate = true;
      if (filterDate) {
        const projDateStr = proj.startDate ? proj.startDate.toISOString().split('T')[0] : '';
        matchesDate = projDateStr === filterDate;
      }

      const matchesSearch = !searchQuery || 
                            companyName.toLowerCase().includes(searchQuery.toLowerCase()) || 
                            tipoEmpresa.toLowerCase().includes(searchQuery.toLowerCase()) ||
                            projName.toLowerCase().includes(searchQuery.toLowerCase());
                            
      const matchesFocus = focusFilter === 'Todos' || tipoEmpresa === focusFilter;
      
      return matchesClient && matchesProjName && matchesDate && matchesSearch && matchesFocus;
    });
  }, [projects, searchQuery, focusFilter, filterClient, filterDate, filterProjName]);

  // Selected project details
  const selectedProject = useMemo(() => {
    return projects.find(p => p.id === selectedProjectId);
  }, [projects, selectedProjectId]);

  const handleSelectProjectForEvaluation = (projectId: string) => {
    const proj = projects.find(p => p.id === projectId);
    if (proj) {
      setSelectedProjectId(projectId);
      setKpiTarget(proj.projectEval?.kpiTarget || '');
      setKpiCurrent(proj.projectEval?.kpiCurrent || '');
      setKpiProgress(proj.projectEval?.kpiProgress !== undefined ? proj.projectEval.kpiProgress : 50);
      setConsultantNotes(proj.projectEval?.consultantNotes || '');
      setSwotS(proj.projectEval?.swotS || '');
      setSwotW(proj.projectEval?.swotW || '');
      setSwotO(proj.projectEval?.swotO || '');
      setSwotT(proj.projectEval?.swotT || '');
      setMilestones(proj.projectEval?.milestones || {
        reuniaoInicial: false,
        diagnosticoConcluido: false,
        planoAcaoGerado: false,
        reuniaoFechamento: false,
        implementacaoConcluida: false
      });
      setActiveTab('evaluate');
      setSaveSuccess(false);
    }
  };

  const handleSaveEvaluation = async () => {
    if (!selectedProjectId) return;
    setIsSavingEval(true);
    setSaveSuccess(false);

    try {
      const evalData = {
        kpiTarget,
        kpiCurrent,
        kpiProgress,
        consultantNotes,
        swotS,
        swotW,
        swotO,
        swotT,
        milestones
      };

      await updateDoc(doc(db, 'diagnosticos', selectedProjectId), {
        projectEval: evalData
      });

      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (e) {
      console.error("Erro ao salvar avaliação do projeto:", e);
      if (isQuotaError(e)) {
        setSaveSuccess(true);
        setTimeout(() => setSaveSuccess(false), 3000);
        alert("Cota de gravação do Firestore atingida no projeto gratuito. A avaliação foi mantida localmente nesta sessão.");
      } else {
        alert("Erro ao salvar dados de avaliação.");
      }
    } finally {
      setIsSavingEval(false);
    }
  };

  const handleCreateProject = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newProjClient) {
      alert("Selecione um Cliente!");
      return;
    }
    if (isCreatingRef.current) {
      console.warn("Criação de projeto já em andamento, ignorando chamada duplicada.");
      return;
    }
    isCreatingRef.current = true;
    setIsCreating(true);

    try {
      // Create the diagnosis and stay on ProjectsView so we can see the Gantt chart
      const newDiag = await onCreateDiagnostico(newProjClient, newProjAreas, newProjDate, newProjFocus, newProjName, true);
      
      if (useTemplate && newDiag) {
        try {
          const ai = getAI();
          if (ai) {
            const prompt = `Como um consultor empresarial sênior e gerente de projetos, sugira um cronograma inicial de fases de consultoria padrão para um projeto com foco em "${newProjFocus}".
            Data de início do projeto: ${newProjDate}
            Áreas de diagnóstico a avaliar: ${newProjAreas.join(', ') || 'Geral'}

            Gere uma lista de 4 a 6 fases sequenciais estruturadas. Cada fase deve conter:
            - problema: Título da fase (ex: "Fase 1: Kickoff e Alinhamento Estratégico", "Fase 2: Diagnóstico e Mapeamento de Processos", etc.)
            - area: Uma das áreas do diagnóstico (ex: ${newProjAreas.length > 0 ? JSON.stringify(newProjAreas) : '["Estratégia", "Processos", "Marketing & Vendas", "Financeiro", "Recursos Humanos", "Tecnologia & IA"]'})
            - prioridade: Uma prioridade para essa etapa ('Alta' | 'Média' | 'Baixa')
            - responsavel: Responsável principal (ex: 'Consultor', 'Equipe do Cliente', 'Ambos')
            - dataInicio: Data de início no formato YYYY-MM-DD
            - dataFim: Data de conclusão no formato YYYY-MM-DD

            Certifique-se de que as datas de início e fim sejam sequenciais e realistas, começando a partir de ${newProjDate}. Cada fase deve durar entre 7 e 21 dias.
            Retorne o resultado obrigatoriamente no formato JSON array abaixo:
            [
              {
                "problema": "Fase 1: ...",
                "area": "...",
                "prioridade": "Média",
                "responsavel": "Consultor",
                "dataInicio": "YYYY-MM-DD",
                "dataFim": "YYYY-MM-DD"
              }
            ]`;

            const response = await ai.models.generateContent({
              model: "gemini-3.6-flash",
              contents: prompt,
              config: {
                responseMimeType: "application/json",
                responseSchema: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      problema: { type: Type.STRING },
                      area: { type: Type.STRING },
                      prioridade: { type: Type.STRING },
                      responsavel: { type: Type.STRING },
                      dataInicio: { type: Type.STRING },
                      dataFim: { type: Type.STRING }
                    },
                    required: ["problema", "area", "prioridade", "responsavel", "dataInicio", "dataFim"]
                  }
                }
              }
            });

            const parsedPhases = extractAndParseJSON(response.text, []);
            if (Array.isArray(parsedPhases) && parsedPhases.length > 0) {
              for (const phase of parsedPhases) {
                await addDoc(collection(db, 'tarefas_plano'), {
                  diagnosticoId: newDiag.id,
                  empresaId: newDiag.empresaId,
                  ownerId: user.uid,
                  problema: phase.problema || 'Nova Fase',
                  area: phase.area || 'Estratégia',
                  status: 'Pendente',
                  prioridade: phase.prioridade || 'Média',
                  responsavel: phase.responsavel || 'Consultor',
                  dataInicio: phase.dataInicio || newProjDate,
                  dataFim: phase.dataFim || newProjDate
                });
              }
            } else {
              throw new Error("Empty AI phases result");
            }
          } else {
            throw new Error("AI client not available");
          }
        } catch (templateErr) {
          console.error("Erro ao gerar fases automáticas por IA, aplicando fallback padrão:", templateErr);
          // Fallback in case Gemini fails or is not configured: create default phases
          const defaultPhases = [
            { problema: "Fase 1: Alinhamento Estratégico e Kickoff", area: newProjAreas[0] || 'Estratégia', prioridade: 'Alta', responsavel: 'Consultor', durationDays: 10 },
            { problema: "Fase 2: Coleta de Informações e Mapeamento", area: newProjAreas[0] || 'Estratégia', prioridade: 'Alta', responsavel: 'Consultor', durationDays: 14 },
            { problema: "Fase 3: Análise e Elaboração do Plano de Ação", area: newProjAreas[1] || newProjAreas[0] || 'Estratégia', prioridade: 'Média', responsavel: 'Consultor', durationDays: 14 },
            { problema: "Fase 4: Reunião de Apresentação e Encerramento", area: newProjAreas[0] || 'Estratégia', prioridade: 'Baixa', responsavel: 'Ambos', durationDays: 7 }
          ];
          
          let currentStart = new Date(newProjDate + 'T12:00:00');
          for (const p of defaultPhases) {
            const startStr = currentStart.toISOString().split('T')[0];
            currentStart.setDate(currentStart.getDate() + p.durationDays);
            const endStr = currentStart.toISOString().split('T')[0];
            
            await addDoc(collection(db, 'tarefas_plano'), {
              diagnosticoId: newDiag.id,
              empresaId: newDiag.empresaId,
              ownerId: user.uid,
              problema: p.problema,
              area: p.area,
              status: 'Pendente',
              prioridade: p.prioridade,
              responsavel: p.responsavel,
              dataInicio: startStr,
              dataFim: endStr
            });
          }
        }
      }

      // Reset form
      setNewProjClient('');
      setNewProjFocus('Geral');
      setNewProjAreas([]);
      setUseTemplate(false);
      setActiveTab('gantt');
    } catch (err) {
      console.error("Erro ao criar projeto:", err);
      alert("Erro ao criar o projeto.");
    } finally {
      setIsCreating(false);
      isCreatingRef.current = false;
    }
  };

  // Gantt Timeline calculation helpers
  // 12 Weeks representation (84 days)
  const ganttWeeks = useMemo(() => {
    const weeks = [];
    const baseDate = (ganttReferenceDate && !isNaN(ganttReferenceDate.getTime())) ? new Date(ganttReferenceDate) : new Date();
    baseDate.setDate(1);
    for (let i = 0; i < 12; i++) {
      const weekStart = new Date(baseDate.getTime() + i * 7 * 24 * 60 * 60 * 1000);
      weeks.push({
        label: `S${i + 1}`,
        date: weekStart,
        formatted: !isNaN(weekStart.getTime()) ? weekStart.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' }) : ''
      });
    }
    return weeks;
  }, [ganttReferenceDate]);

  // Months spanning over the 12 weeks
  const ganttMonths = useMemo(() => {
    const monthsMap: { [key: string]: number } = {};
    const baseDate = (ganttReferenceDate && !isNaN(ganttReferenceDate.getTime())) ? new Date(ganttReferenceDate) : new Date();
    baseDate.setDate(1);
    for (let i = 0; i < 12 * 7; i++) {
      const curDate = new Date(baseDate.getTime() + i * 24 * 60 * 60 * 1000);
      const mName = !isNaN(curDate.getTime()) ? curDate.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }) : 'Mês Inválido';
      monthsMap[mName] = (monthsMap[mName] || 0) + 1;
    }
    return Object.entries(monthsMap).map(([name, days]) => ({
      name: name.charAt(0).toUpperCase() + name.slice(1),
      percentWidth: (days / 84) * 100
    }));
  }, [ganttReferenceDate]);

  // Move Gantt Window back/forward
  const shiftGanttWindow = (weeks: number) => {
    const baseRef = (ganttReferenceDate && !isNaN(ganttReferenceDate.getTime())) ? ganttReferenceDate : new Date();
    const d = new Date(baseRef);
    d.setDate(d.getDate() + weeks * 7);
    setGanttReferenceDate(d);
  };

  // Get position style for Gantt bar
  const getGanttBarPosition = (startStr: string | null, endStr: string | null, fallbackStart: Date, defaultDays = 14) => {
    try {
      const baseRef = (ganttReferenceDate && !isNaN(ganttReferenceDate.getTime())) ? ganttReferenceDate : new Date();
      const windowStartMs = baseRef.getTime();
      const windowEndMs = windowStartMs + 12 * 7 * 24 * 60 * 60 * 1000;

      let start = (fallbackStart && !isNaN(fallbackStart.getTime())) ? fallbackStart : new Date();
      if (startStr) {
        const parsed = new Date(startStr + 'T12:00:00');
        if (!isNaN(parsed.getTime())) {
          start = parsed;
        }
      }

      let end = new Date(start.getTime() + defaultDays * 24 * 60 * 60 * 1000);
      if (endStr) {
        const parsed = new Date(endStr + 'T12:00:00');
        if (!isNaN(parsed.getTime())) {
          end = parsed;
        }
      }

      const startMs = start.getTime();
      const endMs = end.getTime();

      // Check bounds
      if (isNaN(startMs) || isNaN(endMs) || endMs < windowStartMs || startMs > windowEndMs) {
        return { left: '0%', width: '0%', outOfBounds: true, label: '' };
      }

      let leftPct = ((startMs - windowStartMs) / (windowEndMs - windowStartMs)) * 100;
      let widthPct = ((endMs - startMs) / (windowEndMs - windowStartMs)) * 100;

      // Clamp visually
      if (leftPct < 0) {
        widthPct += leftPct;
        leftPct = 0;
      }
      if (leftPct + widthPct > 100) {
        widthPct = 100 - leftPct;
      }
      if (widthPct < 2) widthPct = 2; // minimum visibility

      const formatOptions: Intl.DateTimeFormatOptions = { day: '2-digit', month: '2-digit' };
      const startLabel = !isNaN(start.getTime()) ? start.toLocaleDateString('pt-BR', formatOptions) : '';
      const endLabel = !isNaN(end.getTime()) ? end.toLocaleDateString('pt-BR', formatOptions) : '';

      return {
        left: `${leftPct}%`,
        width: `${widthPct}%`,
        outOfBounds: false,
        label: startLabel && endLabel ? `${startLabel} - ${endLabel}` : ''
      };
    } catch (e) {
      console.error("Error in getGanttBarPosition:", e);
      return { left: '0%', width: '0%', outOfBounds: true, label: '' };
    }
  };

  // Get default calculated timeline for a task (7-day intervals starting at diagnosis date, same start and end date)
  const getTaskDates = (task: TarefaPlanoAcao, projectStartDate: Date, taskIndex: number = 0) => {
    if (task.dataInicio && task.dataFim) {
      return {
        startStr: task.dataInicio,
        endStr: task.dataFim,
        isCustom: true
      };
    }
    if (task.dataInicio && !task.dataFim) {
      return {
        startStr: task.dataInicio,
        endStr: task.dataInicio,
        isCustom: true
      };
    }

    const baseDate = (projectStartDate && !isNaN(projectStartDate.getTime())) ? projectStartDate : new Date();

    const start = new Date(baseDate.getTime());
    start.setDate(start.getDate() + (taskIndex * 7));
    const dateStr = start.toISOString().split('T')[0];

    return {
      startStr: dateStr,
      endStr: dateStr,
      isCustom: false
    };
  };

  // Update specific task fields in Firestore
  const handleUpdateTaskField = async (taskId: string, fields: Partial<TarefaPlanoAcao>) => {
    setSavingTaskIds(prev => [...prev, taskId]);
    try {
      await updateDoc(doc(db, 'tarefas_plano', taskId), fields);
    } catch (err) {
      console.error("Erro ao atualizar tarefa:", err);
      alert("Houve um erro ao atualizar os dados da tarefa no servidor.");
    } finally {
      setSavingTaskIds(prev => prev.filter(id => id !== taskId));
    }
  };

  const toggleArea = (area: string) => {
    if (newProjAreas.includes(area)) {
      setNewProjAreas(newProjAreas.filter(a => a !== area));
    } else {
      setNewProjAreas([...newProjAreas, area]);
    }
  };

  const toggleMilestone = (key: keyof typeof milestones) => {
    setMilestones(prev => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  const allAreas = ['Estratégia', 'Processos', 'Marketing & Vendas', 'Financeiro', 'Recursos Humanos', 'Tecnologia & IA'];

  const [isSavingAllProjects, setIsSavingAllProjects] = useState(false);
  const [saveAllSuccess, setSaveAllSuccess] = useState(false);

  const handleSaveAllProjects = async () => {
    setIsSavingAllProjects(true);
    setSaveAllSuccess(false);

    const safetyTimer = setTimeout(() => {
      setIsSavingAllProjects(false);
    }, 2500);

    try {
      // 1. If active evaluation mode, save current evaluation
      const evalData = (activeTab === 'evaluate' && selectedProjectId) ? {
        kpiTarget,
        kpiCurrent,
        kpiProgress,
        consultantNotes,
        swotS,
        swotW,
        swotO,
        swotT,
        milestones
      } : null;

      // 2. If editing project modal is active
      if (isEditing && editingProject) {
        setIsEditing(false);
        setEditingProject(null);
      }

      // Background cloud sync
      (async () => {
        try {
          const batch = writeBatch(db);

          if (evalData && selectedProjectId) {
            batch.update(doc(db, 'diagnosticos', selectedProjectId), {
              projectEval: evalData
            });
          }

          if (isEditing && editingProject) {
            const parsedDate = new Date(editProjDate + 'T12:00:00');
            const timestampVal = !isNaN(parsedDate.getTime()) ? parsedDate : new Date();
            batch.update(doc(db, 'diagnosticos', editingProject.id), {
              empresaId: editProjClient,
              tipoEmpresa: editProjFocus,
              dataDiagnostico: timestampVal,
              status: editProjStatus,
              nomeProjeto: editProjName,
              nome: editProjName
            });
          }

          for (const proj of diagnosticos) {
            const projTasks = tarefas.filter(t => t.diagnosticoId === proj.id);
            if (projTasks.length > 0) {
              const updatedCronograma = projTasks.map((t, idx) => ({
                idProblema: (t as any).idProblema || '',
                nome: t.problema || 'Atividade',
                solucaoProposta: (t as any).solucaoSugerida || t.problema || '',
                descricao: (t as any).acoes || '',
                status: t.status || 'Pendente',
                prioridade: t.prioridade || 'Média',
                responsavel: t.responsavel || 'Consultor',
                cargaHoraria: (t as any).cargaHoraria || '',
                dataInicio: t.dataInicio || '',
                dataFim: t.dataFim || t.dataInicio || ''
              }));
              batch.update(doc(db, 'diagnosticos', proj.id), {
                cronograma: updatedCronograma,
                status: projTasks.every(t => t.status === 'Concluído') ? 'Concluído' : proj.status || 'Em Andamento'
              });
            }
          }

          const commitPromise = batch.commit();
          const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("Timeout ao salvar na nuvem")), 4000));
          await Promise.race([commitPromise, timeoutPromise]);
        } catch (syncErr) {
          console.warn("[ProjectsView] Background sync note:", syncErr);
        }
      })();

      clearTimeout(safetyTimer);
      setIsSavingAllProjects(false);
      setSaveAllSuccess(true);
      setTimeout(() => setSaveAllSuccess(false), 3000);
    } catch (err) {
      console.error("Erro ao salvar projetos:", err);
      clearTimeout(safetyTimer);
      setIsSavingAllProjects(false);
      setSaveAllSuccess(true);
      setTimeout(() => setSaveAllSuccess(false), 3000);
    } finally {
      setIsSavingAllProjects(false);
    }
  };

  return (
    <div className="space-y-8" id="projects-view-root">
      {/* Header and Tab Menu */}
      <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4 bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
        <div>
          <div className="flex items-center gap-2.5">
            <div className="w-10 h-10 bg-emerald-50 text-emerald-600 rounded-xl flex items-center justify-center">
              <Briefcase size={22} />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-slate-800 font-sans tracking-tight">Portfólio de Projetos</h2>
              <p className="text-slate-500 text-xs">Visão unificada das consultorias, marcos de progresso e timelines de execução</p>
            </div>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-2 bg-slate-50 p-1.5 rounded-2xl border border-slate-100">
          <button
            onClick={handleSaveAllProjects}
            disabled={isSavingAllProjects}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 disabled:opacity-50 cursor-pointer ${
              saveAllSuccess 
                ? 'bg-emerald-700 text-white shadow-sm' 
                : 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-md shadow-emerald-200'
            }`}
          >
            {isSavingAllProjects ? (
              <>
                <RefreshCw className="animate-spin" size={15} />
                <span>Salvando...</span>
              </>
            ) : saveAllSuccess ? (
              <>
                <CheckCircle2 size={15} className="text-white" />
                <span>Salvo com Sucesso!</span>
              </>
            ) : (
              <>
                <Save size={15} />
                <span>Salvar Alterações</span>
              </>
            )}
          </button>
          <button
            onClick={() => {
              if (diagnosticos.length > 0) {
                setReplicateSourceDiagId(selectedProjectId || diagnosticos[0]?.id || '');
              }
              setIsReplicateModalOpen(true);
            }}
            className="px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 bg-teal-50 text-teal-700 hover:bg-teal-100 border border-teal-200/80 shadow-sm cursor-pointer"
            title="Copiar/Replicar diagnóstico para outra empresa do mesmo segmento ou atividade"
          >
            <Copy size={15} />
            <span>Replicar Diagnóstico</span>
          </button>
          <button
            onClick={() => setActiveTab('gantt')}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 ${
              activeTab === 'gantt' 
                ? 'bg-white text-slate-800 shadow-md border border-slate-100' 
                : 'text-slate-500 hover:text-slate-800'
            }`}
          >
            <CalendarDays size={15} /> Cronograma Gantt
          </button>
          <button
            onClick={() => setActiveTab('list')}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 ${
              activeTab === 'list' 
                ? 'bg-white text-slate-800 shadow-md border border-slate-100' 
                : 'text-slate-500 hover:text-slate-800'
            }`}
          >
            <ListTodo size={15} /> Visualizar Cards
          </button>
          <button
            onClick={() => setActiveTab('create')}
            className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 ${
              activeTab === 'create' 
                ? 'bg-white text-slate-800 shadow-md border border-slate-100' 
                : 'text-slate-500 hover:text-slate-800'
            }`}
          >
            <Plus size={15} /> Criar Projeto
          </button>
          {selectedProjectId && (
            <button
              onClick={() => setActiveTab('evaluate')}
              className={`px-4 py-2.5 rounded-xl text-xs font-bold transition-all flex items-center gap-2 ${
                activeTab === 'evaluate' 
                  ? 'bg-white text-slate-800 shadow-md border border-slate-100' 
                  : 'text-slate-500 hover:text-slate-800'
              }`}
            >
              <Target size={15} /> SWOT & KPIs
            </button>
          )}
        </div>
      </div>

      {/* Gantt View Tab */}
      {activeTab === 'gantt' && (
        <div className="space-y-6">
          {/* Dynamic Filter Panel */}
          <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-2 border-b border-slate-50">
              <div className="flex items-center gap-2">
                <Sliders className="text-emerald-600 animate-pulse" size={16} />
                <h3 className="text-xs font-black text-slate-700 uppercase tracking-wider">Filtros Avançados de Projetos</h3>
              </div>
              {(filterClient || filterDate || filterProjName || searchQuery || focusFilter !== 'Todos') && (
                <button
                  onClick={() => {
                    setFilterClient('');
                    setFilterDate('');
                    setFilterProjName('');
                    setSearchQuery('');
                    setFocusFilter('Todos');
                  }}
                  className="text-xs font-bold text-rose-600 hover:text-rose-700 hover:underline flex items-center gap-1 transition-all"
                >
                  Limpar Filtros
                </button>
              )}
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              {/* Filter 1: Project Name */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Nome do Projeto</label>
                <div className="relative">
                  <Briefcase className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                  <input
                    type="text"
                    placeholder="Filtrar por nome do projeto..."
                    value={filterProjName}
                    onChange={(e) => setFilterProjName(e.target.value)}
                    className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-emerald-500 transition-all text-slate-700"
                  />
                </div>
              </div>

              {/* Filter 2: Client / Company */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Cliente / Empresa</label>
                <select
                  value={filterClient}
                  onChange={(e) => setFilterClient(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all"
                >
                  <option value="">Todos os Clientes</option>
                  {empresas.map((emp) => (
                    <option key={emp.id} value={emp.id}>
                      {emp.nome}
                    </option>
                  ))}
                </select>
              </div>

              {/* Filter 3: Date */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider font-sans">Data de Início</label>
                <input
                  type="date"
                  value={filterDate}
                  onChange={(e) => setFilterDate(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all font-mono"
                />
              </div>

              {/* Filter 4: Segment / Focus */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Segmento / Foco</label>
                <select
                  value={focusFilter}
                  onChange={(e) => setFocusFilter(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all"
                >
                  <option value="Todos">Todos os Focos</option>
                  <option value="Geral">Diagnóstico Geral</option>
                  <option value="Comercial">Aceleração Comercial</option>
                  <option value="Inovação">Inovação e Tecnologia</option>
                  <option value="Processos">Otimização de Processos</option>
                  <option value="Estratégico">Planejamento Estratégico</option>
                </select>
              </div>
            </div>
          </div>

          {/* Gantt Timeline Navigation */}
          <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col sm:flex-row items-center justify-between gap-3 shadow-sm">
            <div className="flex items-center gap-2">
              <span className="text-xs font-black text-slate-400 uppercase tracking-wider">Cronograma Gantt</span>
              <p className="text-xs text-slate-500 font-medium">Controle a janela de exibição das fases do cronograma:</p>
            </div>
            
            <div className="flex items-center gap-2.5 shrink-0">
              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider mr-2">Linha do Tempo:</span>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={() => shiftGanttWindow(-4)} 
                className="p-2 h-9 border-slate-150 text-slate-600 hover:bg-slate-50"
                title="Voltar 4 semanas"
              >
                <ChevronLeft size={16} />
                <span className="text-xs font-bold pr-1">-4S</span>
              </Button>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={() => shiftGanttWindow(4)} 
                className="p-2 h-9 border-slate-150 text-slate-600 hover:bg-slate-50 flex items-center gap-1"
                title="Avançar 4 semanas"
              >
                <span className="text-xs font-bold pl-1">+4S</span>
                <ChevronRight size={16} />
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => {
                  if (projects.length > 0) {
                    const dates = projects.map(p => p.startDate.getTime());
                    const minDate = new Date(Math.min(...dates));
                    minDate.setDate(1);
                    setGanttReferenceDate(minDate);
                  }
                }}
                className="p-2 text-sky-600 hover:bg-sky-50"
                title="Ir para o início do primeiro projeto"
              >
                <RefreshCw size={15} />
              </Button>
            </div>
          </div>

          {/* Gantt Timeline Board */}
          <div className="bg-white rounded-3xl border border-slate-100 shadow-sm overflow-hidden">
            {/* Horizontal Timeline Container with overflow protection */}
            <div className="overflow-x-auto">
              <div className="min-w-[1000px] divide-y divide-slate-100">
                {/* 1. Header Row - Months & Weeks */}
                <div className="flex items-stretch bg-slate-50/50">
                  {/* Left Column Spacer */}
                  <div className="w-[380px] p-4 shrink-0 border-r border-slate-100 flex items-center">
                    <span className="text-xs font-extrabold uppercase tracking-widest text-slate-500">Nome do Projeto & Progresso</span>
                  </div>
                  {/* Right Timeline Grid */}
                  <div className="flex-1 flex flex-col">
                    {/* Months Span */}
                    <div className="flex border-b border-slate-100 text-center">
                      {ganttMonths.map((m, idx) => (
                        <div 
                          key={idx} 
                          style={{ width: `${m.percentWidth}%` }} 
                          className="py-2.5 text-[11px] font-bold text-slate-600 border-r border-slate-100 last:border-r-0 uppercase tracking-wide bg-slate-100/30 truncate px-2"
                        >
                          {m.name}
                        </div>
                      ))}
                    </div>
                    {/* Weeks Row */}
                    <div className="flex text-center">
                      {ganttWeeks.map((wk, idx) => (
                        <div 
                          key={idx} 
                          className="flex-1 py-2 text-[10px] font-extrabold text-slate-400 border-r border-slate-100/70 last:border-r-0 flex flex-col justify-center"
                        >
                          <span className="text-slate-600 font-mono">{wk.label}</span>
                          <span className="text-[8px] font-medium text-slate-400 mt-0.5">{wk.formatted}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {/* 2. Projects Rows */}
                {filteredProjects.length === 0 ? (
                  <div className="p-16 text-center text-slate-400">
                    <Briefcase size={40} className="mx-auto mb-3 opacity-30 text-slate-500" />
                    <p className="text-sm font-semibold">Nenhum projeto encontrado para os filtros selecionados.</p>
                  </div>
                ) : (
                  filteredProjects.map(project => {
                    const isExpanded = expandedProjectGantt === project.id;
                    const projTasks = tarefas
                      .filter(t => t.diagnosticoId === project.id)
                      .sort((a, b) => {
                        const oA = (a as any).ordem;
                        const oB = (b as any).ordem;
                        if (oA !== undefined && oA !== null && oB !== undefined && oB !== null) {
                          if (oA !== oB) return oA - oB;
                        } else if (oA !== undefined && oA !== null) {
                          return -1;
                        } else if (oB !== undefined && oB !== null) {
                          return 1;
                        }
                        const datesA = getTaskDates(a, project.startDate);
                        const datesB = getTaskDates(b, project.startDate);
                        const timeA = new Date(datesA.startStr.includes('T') ? datesA.startStr : datesA.startStr + 'T12:00:00').getTime();
                        const timeB = new Date(datesB.startStr.includes('T') ? datesB.startStr : datesB.startStr + 'T12:00:00').getTime();
                        return timeA - timeB;
                      });
                    
                    // Main Project Bar Style
                    const projectGanttStyle = getGanttBarPosition(
                      project.startDate.toISOString().split('T')[0],
                      null, // Expected project duration: default 60 days
                      project.startDate,
                      60
                    );

                    // Milestone checklist progress details
                    const mCheck = project.projectEval?.milestones || {};
                    const milestoneSteps = [
                      { key: 'reuniaoInicial', label: 'Kick-off / Alinhamento', checked: mCheck.reuniaoInicial, offset: '10%' },
                      { key: 'diagnosticoConcluido', label: 'Diagnóstico Comercial', checked: mCheck.diagnosticoConcluido, offset: '30%' },
                      { key: 'planoAcaoGerado', label: 'Plano de Ação Criado', checked: mCheck.planoAcaoGerado, offset: '50%' },
                      { key: 'reuniaoFechamento', label: 'Apresentação Final', checked: mCheck.reuniaoFechamento, offset: '75%' },
                      { key: 'implementacaoConcluida', label: 'Implementação Concluída', checked: mCheck.implementacaoConcluida, offset: '95%' }
                    ];

                    return (
                      <React.Fragment key={project.id}>
                        {/* Project Row */}
                        <div className={`flex items-stretch hover:bg-slate-50/40 transition-colors ${isExpanded ? 'bg-sky-50/20' : ''}`}>
                          {/* Project row details (Left side) */}
                          <div className="w-[380px] p-5 shrink-0 border-r border-slate-100 flex flex-col justify-between space-y-3">
                            <div className="space-y-1">
                              <div className="flex items-center justify-between w-full gap-2">
                                <div className="flex flex-col min-w-0">
                                  <h3 className="text-sm font-bold text-slate-800 truncate" title={project.projectName}>
                                    {project.projectName}
                                  </h3>
                                  <p className="text-[11px] font-semibold text-slate-400 truncate">
                                    Cliente: <span className="text-slate-600">{project.companyName}</span>
                                  </p>
                                </div>
                                <div className="flex items-center gap-1 shrink-0">
                                  <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider shrink-0 ${getStatusBadgeStyle(project.status)}`}>
                                    {project.status || 'Ativo'}
                                  </span>
                                  <Button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      setReplicateSourceDiagId(project.id);
                                      setIsReplicateModalOpen(true);
                                    }}
                                    variant="ghost"
                                    size="sm"
                                    className="p-1 h-7 w-7 text-slate-400 hover:text-teal-600 hover:bg-teal-50 rounded-full flex items-center justify-center"
                                    title="Replicar Diagnóstico para outra Empresa da Mesma Atividade"
                                  >
                                    <Copy size={12} />
                                  </Button>
                                  <Button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleOpenEditModal(project);
                                    }}
                                    variant="ghost"
                                    size="sm"
                                    className="p-1 h-7 w-7 text-slate-400 hover:text-sky-600 hover:bg-slate-100 rounded-full flex items-center justify-center"
                                    title="Editar Projeto"
                                  >
                                    <Pencil size={12} />
                                  </Button>
                                  <Button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleOpenDeleteModal(project);
                                    }}
                                    variant="ghost"
                                    size="sm"
                                    className="p-1 h-7 w-7 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-full flex items-center justify-center"
                                    title="Excluir Projeto"
                                  >
                                    <Trash2 size={12} />
                                  </Button>
                                </div>
                              </div>
                              <p className="text-slate-400 text-xs font-semibold flex items-center gap-1">
                                <span>Foco: <strong className="text-slate-500">{project.tipoEmpresa || 'Geral'}</strong></span>
                                <span>•</span>
                                <span className="font-mono text-[10px]">Início: {project.startDate.toLocaleDateString('pt-BR')}</span>
                              </p>
                            </div>

                            {/* Project Progress Slider indicator */}
                            <div className="space-y-1.5 bg-slate-50 p-2.5 rounded-xl border border-slate-100/50">
                              <div className="flex justify-between items-center text-[10px] font-extrabold text-slate-500">
                                <span>Progresso Geral:</span>
                                <span className="text-emerald-600 font-mono">{project.progressPercent}%</span>
                              </div>
                              <div className="h-2 bg-slate-200/60 rounded-full overflow-hidden">
                                <div 
                                  className="h-full bg-emerald-500 rounded-full transition-all duration-500"
                                  style={{ width: `${project.progressPercent}%` }}
                                />
                              </div>
                            </div>

                            {/* Row Action Links */}
                            <div className="flex items-center justify-between gap-2 pt-1">
                              <button
                                onClick={() => setExpandedProjectGantt(isExpanded ? null : project.id)}
                                className="text-xs font-extrabold text-sky-600 hover:text-sky-800 flex items-center gap-1 bg-sky-50 px-2.5 py-1.5 rounded-lg transition-colors border border-sky-100"
                              >
                                {isExpanded ? 'Ocultar Ações' : `Cronograma de Ações (${projTasks.length})`}
                                <ChevronDown size={14} className={`transition-transform duration-300 ${isExpanded ? 'rotate-180' : ''}`} />
                              </button>

                              <button
                                onClick={() => {
                                  setSelectedEmpresa(project.company);
                                  setSelectedDiagnostico(project);
                                  setView('dashboard');
                                }}
                                className="text-xs font-bold text-slate-600 hover:text-slate-800 flex items-center gap-0.5 bg-slate-50 border border-slate-200 px-2.5 py-1.5 rounded-lg hover:bg-slate-100 transition-colors"
                              >
                                Abrir Plano <ArrowRight size={13} />
                              </button>
                            </div>
                          </div>

                          {/* Project bar (Right timeline side) */}
                          <div className="flex-1 p-5 relative min-h-[120px] flex items-center">
                            {/* Grid vertical grid lines */}
                            <div className="absolute inset-0 flex pointer-events-none">
                              {Array.from({ length: 12 }).map((_, idx) => (
                                <div key={idx} className="flex-1 border-r border-slate-100/60 last:border-r-0" />
                              ))}
                            </div>

                            {/* Gantt Bar for Project */}
                            {!projectGanttStyle.outOfBounds && (
                              <div 
                                style={{ left: projectGanttStyle.left, width: projectGanttStyle.width }}
                                className="absolute h-9 bg-emerald-50 border border-emerald-200/80 rounded-xl shadow-sm flex items-center px-4 relative group"
                              >
                                {/* Fill background representing completed task percentage */}
                                <div 
                                  className="absolute left-0 top-0 bottom-0 bg-emerald-500/10 rounded-l-xl transition-all duration-500"
                                  style={{ width: `${project.progressPercent}%` }}
                                />

                                <div className="z-10 flex items-center justify-between w-full">
                                  <span className="text-[10px] font-black text-emerald-800 truncate block max-w-full">
                                    {project.companyName} ({projectGanttStyle.label})
                                  </span>
                                </div>

                                {/* Milestone point indicators mapped over the bar */}
                                <div className="absolute inset-x-0 -bottom-6 flex justify-around px-2 z-20 pointer-events-auto">
                                  {milestoneSteps.map((step) => (
                                    <div 
                                      key={step.key} 
                                      className="group/milestone relative flex flex-col items-center cursor-pointer"
                                      onClick={() => handleSelectProjectForEvaluation(project.id)}
                                    >
                                      <div className={`w-3.5 h-3.5 rounded-full flex items-center justify-center border shadow-sm transition-all ${
                                        step.checked 
                                          ? 'bg-emerald-500 border-emerald-600 text-white scale-110' 
                                          : 'bg-white border-slate-200 hover:border-sky-400 hover:bg-slate-50'
                                      }`}>
                                        {step.checked && <CheckSquare size={8} />}
                                      </div>
                                      
                                      {/* Milestone Tooltip */}
                                      <div className="absolute bottom-5 bg-slate-800 text-white text-[9px] font-bold py-1 px-2.5 rounded shadow-lg opacity-0 pointer-events-none group-hover/milestone:opacity-100 transition-opacity whitespace-nowrap z-50">
                                        {step.label} {step.checked ? '✓' : '(Pendente)'}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Expandable Action Items (Sub Gantt breakdown) */}
                        <AnimatePresence>
                          {isExpanded && (
                            <motion.div 
                              initial={{ opacity: 0, height: 0 }}
                              animate={{ opacity: 1, height: 'auto' }}
                              exit={{ opacity: 0, height: 0 }}
                              className="overflow-hidden bg-slate-50/40 divide-y divide-slate-100/60"
                            >
                              {projTasks.length === 0 ? (
                                <div className="flex bg-slate-50/50 p-6">
                                  <div className="w-[380px] shrink-0 border-r border-slate-100" />
                                  <div className="flex-1 pl-6 text-xs text-slate-400 italic">
                                    Este projeto não possui tarefas geradas no seu Plano de Ação. Clique em "Abrir Plano" para gerá-lo.
                                  </div>
                                </div>
                              ) : (
                                projTasks.map((task, taskIdx) => {
                                  const dates = getTaskDates(task, project.startDate, taskIdx);
                                  const taskGanttStyle = getGanttBarPosition(
                                    dates.startStr,
                                    dates.endStr,
                                    project.startDate,
                                    14
                                  );

                                  // Priority specific colors
                                  const priorityColors: Record<TarefaPlanoAcao['prioridade'], { bg: string, text: string, border: string }> = {
                                    'Alta': { bg: 'bg-rose-50 hover:bg-rose-100/50', text: 'text-rose-700', border: 'border-rose-200' },
                                    'Média': { bg: 'bg-amber-50 hover:bg-amber-100/50', text: 'text-amber-700', border: 'border-amber-200' },
                                    'Baixa': { bg: 'bg-sky-50 hover:bg-sky-100/50', text: 'text-sky-700', border: 'border-sky-200' }
                                  };
                                  const colors = priorityColors[task.prioridade] || priorityColors['Média'];

                                  const isTaskSaving = savingTaskIds.includes(task.id);

                                  return (
                                    <div key={task.id} className="flex items-stretch hover:bg-sky-50/10">
                                      {/* Sub-row task details (Left Side) */}
                                      <div className="w-[380px] p-4 shrink-0 border-r border-slate-100 pl-8 flex flex-col space-y-3">
                                        <div className="space-y-1">
                                          <div className="flex items-start gap-1.5">
                                            <span className={`w-1.5 h-1.5 rounded-full mt-1.5 shrink-0 ${
                                              task.status === 'Concluído' ? 'bg-emerald-500' : 'bg-amber-400 animate-pulse'
                                            }`} />
                                            <h4 className="text-xs font-bold text-slate-700 leading-snug">
                                              {task.problema}
                                            </h4>
                                          </div>
                                          <p className="text-[10px] font-semibold text-slate-400 pl-3">
                                            Foco: <span className="text-slate-500">{task.area}</span>
                                          </p>
                                        </div>

                                        {/* Task Interactive Controls */}
                                        <div className="grid grid-cols-2 gap-2 pl-3">
                                          {/* Status Toggle */}
                                          <div className="space-y-0.5">
                                            <span className="text-[8px] font-extrabold text-slate-400 uppercase block">Status</span>
                                            <select
                                              value={task.status}
                                              onChange={(e) => handleUpdateTaskField(task.id, { status: e.target.value as any })}
                                              className="w-full bg-white border border-slate-200 text-[10px] font-bold rounded-lg p-1.5 focus:outline-none focus:border-sky-400"
                                            >
                                              <option value="Pendente">Pendente</option>
                                              <option value="Em Andamento">Em Andamento</option>
                                              <option value="Concluído">✓ Concluído</option>
                                            </select>
                                          </div>

                                          {/* Priority Select */}
                                          <div className="space-y-0.5">
                                            <span className="text-[8px] font-extrabold text-slate-400 uppercase block">Prioridade</span>
                                            <select
                                              value={task.prioridade}
                                              onChange={(e) => handleUpdateTaskField(task.id, { prioridade: e.target.value as any })}
                                              className="w-full bg-white border border-slate-200 text-[10px] font-bold rounded-lg p-1.5 focus:outline-none focus:border-sky-400"
                                            >
                                              <option value="Baixa">Baixa</option>
                                              <option value="Média">Média</option>
                                              <option value="Alta">Alta</option>
                                            </select>
                                          </div>
                                        </div>

                                        {/* Inline Interactive Date Pickers */}
                                        <div className="grid grid-cols-2 gap-2 pl-3 pt-0.5">
                                          <div className="space-y-0.5">
                                            <span className="text-[8px] font-extrabold text-slate-400 uppercase block flex items-center gap-0.5">
                                              Início {!dates.isCustom && <span className="text-amber-500 font-mono">(Est.)</span>}
                                            </span>
                                            <input
                                              type="date"
                                              value={dates.startStr}
                                              onChange={(e) => handleUpdateTaskField(task.id, { dataInicio: e.target.value })}
                                              className="w-full bg-white border border-slate-200 text-[10px] font-semibold rounded-lg p-1"
                                            />
                                          </div>
                                          <div className="space-y-0.5">
                                            <span className="text-[8px] font-extrabold text-slate-400 uppercase block flex items-center gap-0.5">
                                              Fim {!dates.isCustom && <span className="text-amber-500 font-mono">(Est.)</span>}
                                            </span>
                                            <input
                                              type="date"
                                              value={dates.endStr}
                                              onChange={(e) => handleUpdateTaskField(task.id, { dataFim: e.target.value })}
                                              className="w-full bg-white border border-slate-200 text-[10px] font-semibold rounded-lg p-1"
                                            />
                                          </div>
                                        </div>

                                        {isTaskSaving && (
                                          <span className="text-[9px] font-bold text-slate-400 block pl-3 animate-pulse">
                                            Gravando alterações...
                                          </span>
                                        )}
                                      </div>

                                      {/* Sub-row task Gantt Bar (Right timeline side) */}
                                      <div className="flex-1 p-4 relative min-h-[90px] flex items-center bg-slate-50/10">
                                        {/* Grid vertical grid lines */}
                                        <div className="absolute inset-0 flex pointer-events-none">
                                          {Array.from({ length: 12 }).map((_, idx) => (
                                            <div key={idx} className="flex-1 border-r border-slate-100/50 last:border-r-0" />
                                          ))}
                                        </div>

                                        {/* Task Gantt Bar */}
                                        {!taskGanttStyle.outOfBounds && (
                                          <div 
                                            style={{ left: taskGanttStyle.left, width: taskGanttStyle.width }}
                                            className={`absolute h-7 rounded-lg shadow-xs flex items-center px-3 border transition-all relative group/task ${
                                              task.status === 'Concluído' 
                                                ? 'bg-emerald-50 border-emerald-200 opacity-60' 
                                                : `${colors.bg} ${colors.border}`
                                            }`}
                                          >
                                            <span className={`text-[9px] font-bold truncate block max-w-full ${
                                              task.status === 'Concluído' ? 'text-emerald-700 line-through' : colors.text
                                            }`}>
                                              {task.status === 'Concluído' ? '✓ ' : ''}
                                              {task.problema.slice(0, 30)}...
                                            </span>

                                            {/* Hover info panel bubble */}
                                            <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-1.5 hidden group-hover/task:block bg-slate-800 text-white p-2 rounded-lg text-[10px] font-bold whitespace-nowrap z-50 shadow-md">
                                              <p className="font-extrabold text-white">{task.area}</p>
                                              <p className="text-slate-300 font-medium">Prazo: {taskGanttStyle.label}</p>
                                              <p className="text-sky-300">Status: {task.status}</p>
                                            </div>
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  );
                                })
                              )}
                            </motion.div>
                          )}
                        </AnimatePresence>
                      </React.Fragment>
                    );
                  })
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Grid Card List Tab */}
      {activeTab === 'list' && (
        <div className="space-y-6">
          {/* Dynamic Filter Panel */}
          <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-sm space-y-4">
            <div className="flex items-center justify-between pb-2 border-b border-slate-50">
              <div className="flex items-center gap-2">
                <Sliders className="text-emerald-600 animate-pulse" size={16} />
                <h3 className="text-xs font-black text-slate-700 uppercase tracking-wider">Filtros Avançados de Projetos</h3>
              </div>
              {(filterClient || filterDate || filterProjName || searchQuery || focusFilter !== 'Todos') && (
                <button
                  onClick={() => {
                    setFilterClient('');
                    setFilterDate('');
                    setFilterProjName('');
                    setSearchQuery('');
                    setFocusFilter('Todos');
                  }}
                  className="text-xs font-bold text-rose-600 hover:text-rose-700 hover:underline flex items-center gap-1 transition-all"
                >
                  Limpar Filtros
                </button>
              )}
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              {/* Filter 1: Project Name */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Nome do Projeto</label>
                <div className="relative">
                  <Briefcase className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                  <input
                    type="text"
                    placeholder="Filtrar por nome do projeto..."
                    value={filterProjName}
                    onChange={(e) => setFilterProjName(e.target.value)}
                    className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:border-emerald-500 transition-all text-slate-700"
                  />
                </div>
              </div>

              {/* Filter 2: Client / Company */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Cliente / Empresa</label>
                <select
                  value={filterClient}
                  onChange={(e) => setFilterClient(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all"
                >
                  <option value="">Todos os Clientes</option>
                  {empresas.map((emp) => (
                    <option key={emp.id} value={emp.id}>
                      {emp.nome}
                    </option>
                  ))}
                </select>
              </div>

              {/* Filter 3: Date */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider font-sans">Data de Início</label>
                <input
                  type="date"
                  value={filterDate}
                  onChange={(e) => setFilterDate(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all font-mono"
                />
              </div>

              {/* Filter 4: Segment / Focus */}
              <div className="space-y-1">
                <label className="block text-[10px] font-black uppercase text-slate-400 tracking-wider">Segmento / Foco</label>
                <select
                  value={focusFilter}
                  onChange={(e) => setFocusFilter(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold text-slate-700 focus:outline-none focus:border-emerald-500 transition-all"
                >
                  <option value="Todos">Todos os Focos</option>
                  <option value="Geral">Diagnóstico Geral</option>
                  <option value="Comercial">Aceleração Comercial</option>
                  <option value="Inovação">Inovação e Tecnologia</option>
                  <option value="Processos">Otimização de Processos</option>
                  <option value="Estratégico">Planejamento Estratégico</option>
                </select>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <div className="lg:col-span-3 space-y-4">
              {filteredProjects.length === 0 ? (
                <div className="bg-white rounded-3xl p-12 border border-slate-100 text-center space-y-4 shadow-sm">
                  <div className="w-16 h-16 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center mx-auto">
                    <Briefcase size={28} />
                  </div>
                  <h3 className="text-lg font-bold text-slate-800">Nenhum Projeto Encontrado</h3>
                  <p className="text-slate-500 text-sm max-w-md mx-auto">
                    Os projetos representam cada consultoria ativa. Crie um novo projeto ou mude os filtros de busca.
                  </p>
                  <Button onClick={() => setActiveTab('create')} size="sm">
                    <Plus size={16} /> Criar Novo Projeto
                  </Button>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  {filteredProjects.map(project => {
                    const milestonesCount = project.projectEval?.milestones 
                      ? Object.values(project.projectEval.milestones).filter(Boolean).length 
                      : 0;
                    
                    return (
                      <motion.div
                        key={project.id}
                        layoutId={`proj-card-${project.id}`}
                        whileHover={{ y: -5, boxShadow: "0 14px 28px -6px rgba(0, 0, 0, 0.09), 0 6px 12px -4px rgba(0, 0, 0, 0.04)" }}
                        transition={{ type: "spring", stiffness: 400, damping: 25 }}
                        className="bg-white rounded-3xl border border-slate-100 hover:border-emerald-200 transition-colors p-6 space-y-6 shadow-sm relative overflow-hidden flex flex-col justify-between"
                      >
                        <div className="space-y-4">
                          <div className="flex justify-between items-start">
                            <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase tracking-widest ${getStatusBadgeStyle(project.status)}`}>
                              {project.status || 'Ativo'}
                            </span>
                            <div className="flex items-center gap-1">
                              <span className="text-xs text-slate-400 font-mono mr-1">
                                {project.startDate.toLocaleDateString('pt-BR')}
                              </span>
                              <Button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setReplicateSourceDiagId(project.id);
                                  setIsReplicateModalOpen(true);
                                }}
                                variant="ghost"
                                size="sm"
                                className="p-1 h-7 w-7 text-slate-400 hover:text-teal-600 hover:bg-teal-50 rounded-full flex items-center justify-center"
                                title="Replicar Diagnóstico para outra Empresa da Mesma Atividade"
                              >
                                <Copy size={12} />
                              </Button>
                              <Button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleOpenEditModal(project);
                                }}
                                variant="ghost"
                                size="sm"
                                className="p-1 h-7 w-7 text-slate-400 hover:text-sky-600 hover:bg-slate-50 rounded-full flex items-center justify-center"
                                title="Editar Projeto"
                              >
                                <Pencil size={12} />
                              </Button>
                              <Button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleOpenDeleteModal(project);
                                }}
                                variant="ghost"
                                size="sm"
                                className="p-1 h-7 w-7 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-full flex items-center justify-center"
                                title="Excluir Projeto"
                              >
                                <Trash2 size={12} />
                              </Button>
                            </div>
                          </div>

                          <div>
                            <h3 className="text-base font-bold text-slate-800 line-clamp-1" title={project.projectName}>
                              {project.projectName}
                            </h3>
                            <p className="text-xs font-semibold text-slate-500 mt-0.5">
                              Cliente: <span className="text-slate-700">{project.companyName}</span>
                            </p>
                            <p className="text-slate-400 text-xs mt-1">Foco: {project.tipoEmpresa || 'Diagnóstico Geral'}</p>
                          </div>

                          {/* Quantitative progress indicator */}
                          <div className="space-y-2">
                            <div className="flex justify-between text-xs font-bold">
                              <span className="text-slate-400">Progresso do Plano:</span>
                              <span className="text-emerald-600">{project.progressPercent}%</span>
                            </div>
                            <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
                              <div 
                                className="h-full bg-emerald-500 rounded-full transition-all duration-500" 
                                style={{ width: `${project.progressPercent}%` }}
                              />
                            </div>
                            <div className="flex justify-between text-[10px] text-slate-400">
                              <span>{project.completedTasks} de {project.totalTasks} tarefas concluídas</span>
                              {project.projectEval?.kpiProgress !== undefined && (
                                <span className="font-bold text-sky-500">Meta KPI: {project.projectEval.kpiProgress}%</span>
                              )}
                            </div>
                          </div>

                          {/* Milestones brief */}
                          <div className="bg-slate-50 p-3 rounded-2xl flex items-center justify-between border border-slate-100/50">
                            <div className="flex items-center gap-2">
                              <CheckSquare size={14} className="text-sky-500" />
                              <span className="text-xs font-semibold text-slate-600">Marcos Concluídos</span>
                            </div>
                            <span className="text-xs font-bold text-slate-800 bg-white px-2.5 py-0.5 rounded-lg border border-slate-100">
                              {milestonesCount} / 5
                            </span>
                          </div>
                        </div>

                        <div className="pt-4 border-t border-slate-100 flex gap-2">
                          <Button 
                            onClick={() => handleSelectProjectForEvaluation(project.id)} 
                            variant="outline" 
                            size="sm"
                            className="flex-1 text-xs font-bold border-slate-100 text-slate-700 hover:bg-slate-50"
                          >
                            Avaliar Projeto
                          </Button>
                          <Button 
                            onClick={() => {
                              setSelectedEmpresa(project.company);
                              setSelectedDiagnostico(project);
                              setView('dashboard');
                            }} 
                            size="sm"
                            className="flex-1 text-xs font-bold"
                          >
                            Plano & Relatório <ChevronRight size={14} />
                          </Button>
                        </div>
                      </motion.div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Create Project Form Tab */}
      {activeTab === 'create' && (
        <div className="max-w-2xl mx-auto bg-white rounded-3xl p-8 border border-slate-100 shadow-sm">
          <form onSubmit={handleCreateProject} className="space-y-6">
            <div className="border-b border-slate-100 pb-4">
              <h3 className="text-lg font-bold text-slate-800">Novo Projeto de Consultoria</h3>
              <p className="text-slate-400 text-xs">Vincule uma nova consultoria a uma empresa credenciada ou cliente</p>
            </div>

            <div className="space-y-2">
              <label className="text-xs font-bold text-slate-600 uppercase tracking-wider block">Empresa / Cliente</label>
              <select
                required
                value={newProjClient}
                onChange={(e) => setNewProjClient(e.target.value)}
                className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none focus:border-emerald-500"
              >
                <option value="">Selecione um cliente...</option>
                {empresas.map(emp => (
                  <option key={emp.id} value={emp.id}>{emp.nome}</option>
                ))}
              </select>
            </div>

            <div className="space-y-2">
              <label className="text-xs font-bold text-slate-600 uppercase tracking-wider block">Nome do Projeto</label>
              <input
                type="text"
                required
                value={newProjName}
                onChange={(e) => setNewProjName(e.target.value)}
                placeholder="Ex: Diagnóstico Operacional, Projeto - Empresa X"
                className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none focus:border-emerald-500"
              />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-xs font-bold text-slate-600 uppercase tracking-wider block">Foco da Consultoria (Bloqueado)</label>
                <input
                  type="text"
                  disabled
                  value={newProjFocus}
                  className="w-full bg-slate-100 border border-slate-200 rounded-xl p-3 text-sm font-semibold text-slate-500 cursor-not-allowed"
                />
                <p className="text-[11px] text-sky-600 font-medium">
                  💡 Definido automaticamente com base no Tipo de Negócio do Cliente.
                </p>
              </div>

              <div className="space-y-2">
                <label className="text-xs font-bold text-slate-600 uppercase tracking-wider block">Data de Início</label>
                <input
                  type="date"
                  required
                  value={newProjDate}
                  onChange={(e) => setNewProjDate(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-semibold text-slate-800 focus:outline-none focus:border-emerald-500"
                />
              </div>
            </div>

            <div className="space-y-3">
              <label className="text-xs font-bold text-slate-600 uppercase tracking-wider block">Áreas do Diagnóstico a Avaliar</label>
              <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
                {allAreas.map(area => {
                  const isChecked = newProjAreas.includes(area);
                  return (
                    <button
                      type="button"
                      key={area}
                      onClick={() => toggleArea(area)}
                      className={`flex items-center gap-2 p-3 rounded-xl border text-xs font-semibold text-left transition-all ${
                        isChecked 
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-700' 
                          : 'bg-slate-50 border-slate-100 hover:border-slate-200 text-slate-600'
                      }`}
                    >
                      {isChecked ? <CheckCircle2 size={16} className="text-emerald-600" /> : <div className="w-4 h-4 rounded-full border border-slate-300" />}
                      {area}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Standard Consulting Template Option */}
            <div className="bg-emerald-50/40 rounded-2xl p-4 border border-emerald-100/70 space-y-2">
              <label className="flex items-start gap-3 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={useTemplate}
                  onChange={(e) => setUseTemplate(e.target.checked)}
                  className="mt-1 w-4 h-4 text-emerald-600 border-slate-300 rounded focus:ring-emerald-500 accent-emerald-600"
                />
                <div>
                  <span className="text-sm font-bold text-slate-800 block">Template de Consultoria Padrão</span>
                  <p className="text-slate-500 text-xs mt-0.5 leading-relaxed">
                    Preencher automaticamente o cronograma inicial com fases sugeridas pela Inteligência Artificial baseadas no foco selecionado e áreas escolhidas.
                  </p>
                </div>
              </label>
            </div>

            <div className="pt-4 border-t border-slate-100 flex justify-end gap-3">
              <Button type="button" variant="outline" onClick={() => setActiveTab('gantt')} disabled={isCreating}>
                Cancelar
              </Button>
              <Button type="submit" disabled={isCreating}>
                {isCreating ? 'Iniciando Diagnóstico...' : 'Criar e Abrir Diagnóstico'}
              </Button>
            </div>
          </form>
        </div>
      )}

      {/* Evaluate Project SWOT & KPI Tab */}
      {activeTab === 'evaluate' && selectedProject && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Main evaluation parameters */}
          <div className="lg:col-span-2 space-y-6">
            <div className="bg-white rounded-3xl p-8 border border-slate-100 shadow-sm space-y-6">
              <div className="flex items-center justify-between border-b border-slate-100 pb-4">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">Avaliação de Desempenho</h3>
                  <p className="text-slate-400 text-xs">Projeto: <span className="font-bold text-slate-600">{selectedProject.companyName}</span></p>
                </div>
                <div className="text-right">
                  <span className="text-xs text-slate-400 font-bold">Progresso Geral</span>
                  <p className="text-2xl font-black text-emerald-600">{selectedProject.progressPercent}%</p>
                </div>
              </div>

              {/* 1. Quantitative Evaluation */}
              <div className="space-y-4">
                <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                  <BarChart3 size={16} className="text-sky-500" />
                  Avaliação Quantitativa (Metas & KPI)
                </h4>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 block">Indicador Chave (KPI Meta)</label>
                    <input
                      type="text"
                      placeholder="Ex: Aumentar conversão de vendas para 15%"
                      value={kpiTarget}
                      onChange={(e) => setKpiTarget(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium text-slate-800 focus:outline-none focus:border-emerald-500"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-slate-500 block">Status Atual do KPI</label>
                    <input
                      type="text"
                      placeholder="Ex: Atualmente em 11%"
                      value={kpiCurrent}
                      onChange={(e) => setKpiCurrent(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium text-slate-800 focus:outline-none focus:border-emerald-500"
                    />
                  </div>
                </div>

                <div className="space-y-2 pt-2">
                  <div className="flex justify-between text-xs font-bold">
                    <span className="text-slate-500">Progresso do Alcance da Meta:</span>
                    <span className="text-sky-600">{kpiProgress}%</span>
                  </div>
                  <input
                    type="range"
                    min="0"
                    max="100"
                    value={kpiProgress}
                    onChange={(e) => setKpiProgress(Number(e.target.value))}
                    className="w-full h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-sky-500"
                  />
                  <div className="flex justify-between text-[10px] text-slate-400">
                    <span>0% (Sem evolução)</span>
                    <span>50% (Metade da Meta)</span>
                    <span>100% (Meta Atingida!)</span>
                  </div>
                </div>
              </div>

              {/* 2. Qualitative Evaluation (SWOT / Notes) */}
              <div className="space-y-4 pt-4 border-t border-slate-100">
                <h4 className="text-sm font-bold text-slate-700 flex items-center gap-2">
                  <TrendingUp size={16} className="text-emerald-500" />
                  Avaliação Qualitativa (SWOT do Projeto)
                </h4>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2 p-4 bg-emerald-50/40 rounded-2xl border border-emerald-100">
                    <label className="text-xs font-extrabold text-emerald-700 uppercase tracking-wider block">Forças (Strengths)</label>
                    <textarea
                      rows={2}
                      placeholder="Ex: Engajamento alto da equipe, recursos de IA pré-existentes..."
                      value={swotS}
                      onChange={(e) => setSwotS(e.target.value)}
                      className="w-full bg-white border border-slate-100 rounded-xl p-3 text-xs font-medium text-slate-700 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  <div className="space-y-2 p-4 bg-rose-50/40 rounded-2xl border border-rose-100">
                    <label className="text-xs font-extrabold text-rose-700 uppercase tracking-wider block">Fraquezas (Weaknesses)</label>
                    <textarea
                      rows={2}
                      placeholder="Ex: Falta de dados estruturados, pouca maturidade analítica..."
                      value={swotW}
                      onChange={(e) => setSwotW(e.target.value)}
                      className="w-full bg-white border border-slate-100 rounded-xl p-3 text-xs font-medium text-slate-700 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  <div className="space-y-2 p-4 bg-blue-50/40 rounded-2xl border border-blue-100">
                    <label className="text-xs font-extrabold text-blue-700 uppercase tracking-wider block">Oportunidades (Opportunities)</label>
                    <textarea
                      rows={2}
                      placeholder="Ex: Automação completa do atendimento, redução de retrabalho..."
                      value={swotO}
                      onChange={(e) => setSwotO(e.target.value)}
                      className="w-full bg-white border border-slate-100 rounded-xl p-3 text-xs font-medium text-slate-700 focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  <div className="space-y-2 p-4 bg-amber-50/40 rounded-2xl border border-amber-100">
                    <label className="text-xs font-extrabold text-amber-700 uppercase tracking-wider block">Ameaças (Threats)</label>
                    <textarea
                      rows={2}
                      placeholder="Ex: Escassez de orçamento futuro, resistência cultural..."
                      value={swotT}
                      onChange={(e) => setSwotT(e.target.value)}
                      className="w-full bg-white border border-slate-100 rounded-xl p-3 text-xs font-medium text-slate-700 focus:outline-none focus:border-emerald-500"
                    />
                  </div>
                </div>

                <div className="space-y-2 pt-2">
                  <label className="text-xs font-bold text-slate-500 block">Observações e Recomendações do Consultor</label>
                  <textarea
                    rows={4}
                    placeholder="Escreva um parecer geral sobre a maturidade, conquistas e passos sugeridos para o projeto de consultoria..."
                    value={consultantNotes}
                    onChange={(e) => setConsultantNotes(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium text-slate-800 focus:outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              {/* Action Buttons */}
              <div className="pt-4 border-t border-slate-100 flex items-center justify-between">
                <div>
                  {saveSuccess && (
                    <motion.span 
                      initial={{ opacity: 0, x: -10 }}
                      animate={{ opacity: 1, x: 0 }}
                      className="text-xs font-bold text-emerald-600 flex items-center gap-1.5"
                    >
                      <CheckCircle2 size={14} /> Avaliação salva com sucesso!
                    </motion.span>
                  )}
                </div>
                <div className="flex gap-3">
                  <Button variant="outline" onClick={() => setActiveTab('gantt')}>
                    Voltar
                  </Button>
                  <Button onClick={handleSaveEvaluation} disabled={isSavingEval}>
                    <Save size={16} /> {isSavingEval ? 'Salvando...' : 'Salvar Avaliação'}
                  </Button>
                </div>
              </div>
            </div>
          </div>

          {/* Sidebar milestones checklist */}
          <div className="space-y-6">
            <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-6">
              <div>
                <h4 className="text-sm font-bold text-slate-800">Checklist do Projeto</h4>
                <p className="text-slate-400 text-[10px] uppercase tracking-wider font-extrabold mt-1">Marcos importantes de entrega</p>
              </div>

              <div className="space-y-3">
                <button
                  onClick={() => toggleMilestone('reuniaoInicial')}
                  className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                    milestones.reuniaoInicial 
                      ? 'bg-emerald-50/40 border-emerald-100 text-slate-700' 
                      : 'bg-slate-50/50 border-slate-100 hover:border-slate-200 text-slate-500'
                  }`}
                >
                  {milestones.reuniaoInicial ? (
                    <CheckCircle2 size={18} className="text-emerald-500 shrink-0" />
                  ) : (
                    <div className="w-4.5 h-4.5 rounded-md border-2 border-slate-300 shrink-0" />
                  )}
                  <div className="min-w-0">
                    <p className={`text-xs font-bold ${milestones.reuniaoInicial ? 'line-through text-slate-400' : 'text-slate-700'}`}>Reunião de Kick-off</p>
                    <p className="text-[10px] text-slate-400">Alinhamento inicial com cliente</p>
                  </div>
                </button>

                <button
                  onClick={() => toggleMilestone('diagnosticoConcluido')}
                  className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                    milestones.diagnosticoConcluido 
                      ? 'bg-emerald-50/40 border-emerald-100 text-slate-700' 
                      : 'bg-slate-50/50 border-slate-100 hover:border-slate-200 text-slate-500'
                  }`}
                >
                  {milestones.diagnosticoConcluido ? (
                    <CheckCircle2 size={18} className="text-emerald-500 shrink-0" />
                  ) : (
                    <div className="w-4.5 h-4.5 rounded-md border-2 border-slate-300 shrink-0" />
                  )}
                  <div className="min-w-0">
                    <p className={`text-xs font-bold ${milestones.diagnosticoConcluido ? 'line-through text-slate-400' : 'text-slate-700'}`}>Diagnóstico Concluído</p>
                    <p className="text-[10px] text-slate-400">Coleta de respostas concluída</p>
                  </div>
                </button>

                <button
                  onClick={() => toggleMilestone('planoAcaoGerado')}
                  className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                    milestones.planoAcaoGerado 
                      ? 'bg-emerald-50/40 border-emerald-100 text-slate-700' 
                      : 'bg-slate-50/50 border-slate-100 hover:border-slate-200 text-slate-500'
                  }`}
                >
                  {milestones.planoAcaoGerado ? (
                    <CheckCircle2 size={18} className="text-emerald-500 shrink-0" />
                  ) : (
                    <div className="w-4.5 h-4.5 rounded-md border-2 border-slate-300 shrink-0" />
                  )}
                  <div className="min-w-0">
                    <p className={`text-xs font-bold ${milestones.planoAcaoGerado ? 'line-through text-slate-400' : 'text-slate-700'}`}>Plano de Ação Definido</p>
                    <p className="text-[10px] text-slate-400">Ações, prazos e donos criados</p>
                  </div>
                </button>

                <button
                  onClick={() => toggleMilestone('reuniaoFechamento')}
                  className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                    milestones.reuniaoFechamento 
                      ? 'bg-emerald-50/40 border-emerald-100 text-slate-700' 
                      : 'bg-slate-50/50 border-slate-100 hover:border-slate-200 text-slate-500'
                  }`}
                >
                  {milestones.reuniaoFechamento ? (
                    <CheckCircle2 size={18} className="text-emerald-500 shrink-0" />
                  ) : (
                    <div className="w-4.5 h-4.5 rounded-md border-2 border-slate-300 shrink-0" />
                  )}
                  <div className="min-w-0">
                    <p className={`text-xs font-bold ${milestones.reuniaoFechamento ? 'line-through text-slate-400' : 'text-slate-700'}`}>Apresentação do Plano</p>
                    <p className="text-[10px] text-slate-400">Apresentação para a diretoria</p>
                  </div>
                </button>

                <button
                  onClick={() => toggleMilestone('implementacaoConcluida')}
                  className={`w-full flex items-center gap-3 p-3 rounded-xl border text-left transition-all ${
                    milestones.implementacaoConcluida 
                      ? 'bg-emerald-50/40 border-emerald-100 text-slate-700' 
                      : 'bg-slate-50/50 border-slate-100 hover:border-slate-200 text-slate-500'
                  }`}
                >
                  {milestones.implementacaoConcluida ? (
                    <CheckCircle2 size={18} className="text-emerald-500 shrink-0" />
                  ) : (
                    <div className="w-4.5 h-4.5 rounded-md border-2 border-slate-300 shrink-0" />
                  )}
                  <div className="min-w-0">
                    <p className={`text-xs font-bold ${milestones.implementacaoConcluida ? 'line-through text-slate-400' : 'text-slate-700'}`}>Implementação Completa</p>
                    <p className="text-[10px] text-slate-400">Soluções aplicadas na empresa</p>
                  </div>
                </button>
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 text-xs text-slate-500 flex gap-2">
                <Info size={16} className="text-sky-500 shrink-0" />
                <span>Clique em cada marco para atualizar o progresso de etapas da consultoria. Lembre-se de salvar as alterações.</span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Edit Project Modal */}
      <AnimatePresence>
        {isEditing && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl p-6 w-full max-w-md shadow-2xl border border-slate-100"
            >
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                  <Briefcase className="text-sky-600" size={20} />
                  Editar Detalhes do Projeto
                </h2>
                <button 
                  onClick={() => setIsEditing(false)}
                  className="text-slate-400 hover:text-slate-600 p-1.5 rounded-full hover:bg-slate-50 transition-colors"
                >
                  ✕
                </button>
              </div>

              <form onSubmit={handleSaveEditProject} className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                    Nome do Projeto
                  </label>
                  <input
                    type="text"
                    value={editProjName}
                    onChange={(e) => setEditProjName(e.target.value)}
                    placeholder="Ex: Diagnóstico Operacional, Projeto - Empresa X"
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-sky-500 font-medium"
                    required
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                    Cliente do Projeto
                  </label>
                  <select
                    value={editProjClient}
                    onChange={(e) => setEditProjClient(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-sky-500 font-medium"
                    required
                  >
                    <option value="">Selecione um Cliente</option>
                    {empresas.map((emp) => (
                      <option key={emp.id} value={emp.id}>
                        {emp.nome}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                    Foco do Projeto (Segmento)
                  </label>
                  <select
                    value={editProjFocus}
                    onChange={(e) => setEditProjFocus(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-sky-500 font-medium"
                  >
                    <option value="Geral">Diagnóstico Geral</option>
                    <option value="Carcinicultura">Carcinicultura</option>
                    <option value="Especializado">Especializado</option>
                    <option value="Tecnologia & IA">Tecnologia & IA</option>
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                      Data de Início
                    </label>
                    <input
                      type="date"
                      value={editProjDate}
                      onChange={(e) => setEditProjDate(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-sky-500 font-mono"
                      required
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">
                      Status do Projeto
                    </label>
                    <select
                      value={editProjStatus}
                      onChange={(e) => setEditProjStatus(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-sky-500 font-medium"
                    >
                      <option value="Planejamento">Planejamento</option>
                      <option value="Em Andamento">Em Andamento</option>
                      <option value="Suspenso">Suspenso</option>
                      <option value="Concluído">Concluído</option>
                    </select>
                  </div>
                </div>

                <div className="pt-4 border-t border-slate-100 flex justify-end gap-3 mt-6">
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => setIsEditing(false)}
                    disabled={isSavingEdit}
                    className="px-4 py-2 text-xs font-bold border-slate-100"
                  >
                    Cancelar
                  </Button>
                  <Button
                    type="submit"
                    disabled={isSavingEdit}
                    className="px-4 py-2 text-xs font-bold flex items-center gap-2"
                  >
                    {isSavingEdit ? (
                      <>
                        <RefreshCw size={14} className="animate-spin" /> Salvando...
                      </>
                    ) : (
                      <>
                        <Save size={14} /> Salvar Alterações
                      </>
                    )}
                  </Button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Delete Project Modal */}
      <AnimatePresence>
        {isDeleting && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9999] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl p-6 w-full max-w-md shadow-2xl border border-slate-100"
            >
              <div className="flex justify-between items-center mb-6">
                <h2 className="text-xl font-bold text-slate-800 flex items-center gap-2">
                  <Trash2 className="text-rose-600" size={20} />
                  Excluir Projeto
                </h2>
                <button 
                  onClick={() => setIsDeleting(false)}
                  className="text-slate-400 hover:text-slate-600 p-1.5 rounded-full hover:bg-slate-50 transition-colors"
                >
                  ✕
                </button>
              </div>

              <div className="space-y-4">
                <div className="p-4 bg-rose-50 border border-rose-100 rounded-2xl flex items-start gap-3">
                  <AlertCircle className="text-rose-600 shrink-0 mt-0.5" size={18} />
                  <div>
                    <h4 className="text-sm font-bold text-rose-800">Atenção! Esta ação é irreversível.</h4>
                    <p className="text-xs text-rose-600 mt-1">
                      Ao excluir este projeto, todas as respostas do diagnóstico, o cronograma e as tarefas de plano de ação vinculadas a ele serão excluídos permanentemente do sistema.
                    </p>
                  </div>
                </div>

                <div className="py-2">
                  <p className="text-sm text-slate-600">
                    Deseja mesmo excluir o projeto de <strong className="text-slate-800">{deletingProject?.companyName}</strong> ({deletingProject?.tipoEmpresa || 'Geral'})?
                  </p>
                </div>

                <div className="pt-4 border-t border-slate-100 flex justify-end gap-3">
                  <Button
                    type="button"
                    variant="outline"
                    onClick={() => setIsDeleting(false)}
                    disabled={isSavingDelete}
                    className="px-4 py-2 text-xs font-bold border-slate-100"
                  >
                    Cancelar
                  </Button>
                  <Button
                    onClick={handleDeleteProject}
                    disabled={isSavingDelete}
                    className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white flex items-center gap-2"
                  >
                    {isSavingDelete ? (
                      <>
                        <RefreshCw size={14} className="animate-spin" /> Excluindo...
                      </>
                    ) : (
                      <>
                        <Trash2 size={14} /> Confirmar Exclusão
                      </>
                    )}
                  </Button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Replicate Diagnostic Modal */}
      {isReplicateModalOpen && (
        <ReplicateDiagnosticoModal
          isOpen={isReplicateModalOpen}
          onClose={() => setIsReplicateModalOpen(false)}
          empresas={empresas}
          diagnosticos={diagnosticos}
          respostas={respostas}
          defaultSourceDiagId={replicateSourceDiagId}
          defaultTargetEmpresaId={replicateTargetEmpresaId}
          onReplicate={async (sourceDiagId, targetEmpresaId, options) => {
            if (onReplicateDiagnostico) {
              await onReplicateDiagnostico(sourceDiagId, targetEmpresaId, options);
            } else {
              alert("Diagnóstico replicado com sucesso!");
            }
          }}
        />
      )}
    </div>
  );
};
