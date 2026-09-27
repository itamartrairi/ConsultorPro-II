import React, { useState, useEffect, useMemo } from 'react';
import { motion } from 'motion/react';
import { 
  Calendar as CalendarIcon, 
  Plus, 
  Clock, 
  MapPin, 
  Building2, 
  User, 
  ChevronLeft, 
  ChevronRight, 
  Trash2, 
  CheckCircle2, 
  XCircle, 
  AlertCircle, 
  ListTodo,
  FileText,
  Briefcase,
  ExternalLink
} from 'lucide-react';
import { 
  collection, 
  query, 
  where, 
  onSnapshot, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  orderBy 
} from '../lib/firestoreOwned';
import { db } from '../firebase';
import { Button } from './Button';

interface Empresa {
  id: string;
  nome: string;
}

interface Diagnostico {
  id: string;
  empresaId: string;
  tipoEmpresa?: string;
}

interface AgendaEvento {
  id: string;
  titulo: string;
  data: string; // YYYY-MM-DD
  horaInicio: string; // HH:MM
  horaFim: string; // HH:MM
  empresaId?: string;
  diagnosticoId?: string;
  tipo: 'Reunião' | 'Sessão de Diagnóstico' | 'Acompanhamento' | 'Feedback' | 'Outro';
  descricao?: string;
  status: 'Agendado' | 'Realizado' | 'Cancelado';
  ownerId: string;
}

interface AgendaViewProps {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  user: any;
}

export const AgendaView: React.FC<AgendaViewProps> = ({
  empresas,
  diagnosticos,
  user
}) => {
  const [eventos, setEventos] = useState<AgendaEvento[]>([]);
  const [loading, setLoading] = useState(true);

  // Calendar State
  const [currentDate, setCurrentDate] = useState(new Date());
  const [selectedDateStr, setSelectedDateStr] = useState<string>(
    new Date().toISOString().split('T')[0]
  );

  // New Event Form State
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingEventId, setEditingEventId] = useState<string | null>(null);
  
  const [formTitulo, setFormTitulo] = useState('');
  const [formDate, setFormDate] = useState('');
  const [formHoraInicio, setFormHoraInicio] = useState('09:00');
  const [formHoraFim, setFormHoraFim] = useState('10:00');
  const [formEmpresaId, setFormEmpresaId] = useState('');
  const [formDiagnosticoId, setFormDiagnosticoId] = useState('');
  const [formTipo, setFormTipo] = useState<AgendaEvento['tipo']>('Reunião');
  const [formDescricao, setFormDescricao] = useState('');
  const [formStatus, setFormStatus] = useState<AgendaEvento['status']>('Agendado');
  const [syncToGoogleOnSave, setSyncToGoogleOnSave] = useState(true);

  // Google Calendar Integration Account
  const GOOGLE_CALENDAR_EMAIL = 'itamartrairi@gmail.com';

  const getGoogleCalendarUrl = (evento: {
    titulo: string;
    data: string;
    horaInicio: string;
    horaFim: string;
    descricao?: string;
    empresaId?: string | null;
  }) => {
    try {
      const cleanDate = evento.data.replace(/-/g, '');
      const cleanStart = evento.horaInicio.replace(':', '') + '00';
      const cleanEnd = (evento.horaFim || evento.horaInicio).replace(':', '') + '00';
      const dates = `${cleanDate}T${cleanStart}/${cleanDate}T${cleanEnd}`;

      const company = empresas.find(e => e.id === evento.empresaId);
      let details = evento.descricao ? `${evento.descricao}\n\n` : '';
      if (company) {
        details += `Cliente: ${company.nome}\n`;
      }
      details += `Gerado pelo ConsultorPro II`;

      const title = company ? `${evento.titulo} - ${company.nome}` : evento.titulo;

      const url = new URL('https://calendar.google.com/calendar/render');
      url.searchParams.set('action', 'TEMPLATE');
      url.searchParams.set('text', title);
      url.searchParams.set('dates', dates);
      url.searchParams.set('details', details);
      if (company) {
        url.searchParams.set('location', company.nome);
      }
      url.searchParams.set('add', GOOGLE_CALENDAR_EMAIL);
      url.searchParams.set('authuser', GOOGLE_CALENDAR_EMAIL);

      return url.toString();
    } catch (e) {
      console.error("Erro ao gerar link do Google Calendar:", e);
      return `https://calendar.google.com/calendar/u/0/r?authuser=${GOOGLE_CALENDAR_EMAIL}`;
    }
  };

  // Load and listen to consultant events
  useEffect(() => {
    if (!user) return;

    const q = query(
      collection(db, 'agenda_eventos'), 
      where('ownerId', '==', user.uid)
    );

    const unsub = onSnapshot(q, (snap) => {
      const data = snap.docs.map(d => ({ id: d.id, ...d.data() } as AgendaEvento));
      // Sort events by date and starting time
      data.sort((a, b) => {
        const dateCompare = a.data.localeCompare(b.data);
        if (dateCompare !== 0) return dateCompare;
        return a.horaInicio.localeCompare(b.horaInicio);
      });
      setEventos(data);
      setLoading(false);
    }, (err) => {
      console.error("Erro ao escutar agenda_eventos:", err);
      setLoading(false);
    });

    return () => unsub();
  }, [user]);

  // Calendar Helpers
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();

  const daysInMonth = useMemo(() => {
    return new Date(year, month + 1, 0).getDate();
  }, [year, month]);

  const firstDayIndex = useMemo(() => {
    // 0 = Sunday, 1 = Monday, etc.
    return new Date(year, month, 1).getDay();
  }, [year, month]);

  const handlePrevMonth = () => {
    setCurrentDate(new Date(year, month - 1, 1));
  };

  const handleNextMonth = () => {
    setCurrentDate(new Date(year, month + 1, 1));
  };

  const monthNames = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ];

  const handleDayClick = (day: number) => {
    const formattedDay = day.toString().padStart(2, '0');
    const formattedMonth = (month + 1).toString().padStart(2, '0');
    const dateStr = `${year}-${formattedMonth}-${formattedDay}`;
    setSelectedDateStr(dateStr);
  };

  // Filtered events for the selected date
  const eventsForSelectedDate = useMemo(() => {
    return eventos.filter(e => e.data === selectedDateStr);
  }, [eventos, selectedDateStr]);

  // Upcoming events (today and onwards)
  const upcomingEvents = useMemo(() => {
    const todayStr = new Date().toISOString().split('T')[0];
    return eventos.filter(e => e.data >= todayStr).slice(0, 5);
  }, [eventos]);

  // Available diagnostics for selected company in form
  const filteredDiagnostics = useMemo(() => {
    if (!formEmpresaId) return [];
    return diagnosticos.filter(d => d.empresaId === formEmpresaId);
  }, [diagnosticos, formEmpresaId]);

  // Submit new or edited event
  const handleSubmitEvent = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formTitulo.trim() || !formDate) {
      alert("Título e Data são obrigatórios!");
      return;
    }

    try {
      const payload = {
        titulo: formTitulo.trim(),
        data: formDate,
        horaInicio: formHoraInicio,
        horaFim: formHoraFim,
        empresaId: formEmpresaId || null,
        diagnosticoId: formDiagnosticoId || null,
        tipo: formTipo,
        descricao: formDescricao.trim(),
        status: formStatus,
        ownerId: user.uid
      };

      if (editingEventId) {
        // Update
        await updateDoc(doc(db, 'agenda_eventos', editingEventId), payload);
      } else {
        // Create
        await addDoc(collection(db, 'agenda_eventos'), payload);
      }

      if (syncToGoogleOnSave) {
        const gcalUrl = getGoogleCalendarUrl(payload);
        window.open(gcalUrl, '_blank', 'noopener,noreferrer');
      }

      setIsModalOpen(false);
      resetForm();
    } catch (err) {
      console.error("Erro ao salvar compromisso:", err);
      alert("Erro ao salvar o compromisso.");
    }
  };

  const handleEditEvent = (evt: AgendaEvento) => {
    setEditingEventId(evt.id);
    setFormTitulo(evt.titulo);
    setFormDate(evt.data);
    setFormHoraInicio(evt.horaInicio);
    setFormHoraFim(evt.horaFim);
    setFormEmpresaId(evt.empresaId || '');
    setFormDiagnosticoId(evt.diagnosticoId || '');
    setFormTipo(evt.tipo);
    setFormDescricao(evt.descricao || '');
    setFormStatus(evt.status);
    setIsModalOpen(true);
  };

  const handleDeleteEvent = async (evtId: string) => {
    if (!window.confirm("Deseja realmente remover este compromisso da agenda?")) return;
    try {
      await deleteDoc(doc(db, 'agenda_eventos', evtId));
    } catch (err) {
      console.error("Erro ao deletar compromisso:", err);
      alert("Erro ao remover compromisso.");
    }
  };

  const handleToggleStatus = async (evt: AgendaEvento) => {
    const nextStatusMap: Record<AgendaEvento['status'], AgendaEvento['status']> = {
      'Agendado': 'Realizado',
      'Realizado': 'Cancelado',
      'Cancelado': 'Agendado'
    };
    const nextStatus = nextStatusMap[evt.status];
    try {
      await updateDoc(doc(db, 'agenda_eventos', evt.id), { status: nextStatus });
    } catch (err) {
      console.error("Erro ao atualizar status:", err);
    }
  };

  const resetForm = () => {
    setEditingEventId(null);
    setFormTitulo('');
    setFormDate(selectedDateStr);
    setFormHoraInicio('09:00');
    setFormHoraFim('10:00');
    setFormEmpresaId('');
    setFormDiagnosticoId('');
    setFormTipo('Reunião');
    setFormDescricao('');
    setFormStatus('Agendado');
  };

  const openCreateModal = () => {
    resetForm();
    setIsModalOpen(true);
  };

  return (
    <div className="space-y-8" id="agenda-view-root">
      {/* Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">Agenda do Consultor</h2>
          <p className="text-slate-500 text-sm">Organize suas reuniões, sessões de diagnóstico e feedbacks com os clientes</p>
        </div>
        <div className="flex items-center gap-3">
          <a
            href={`https://calendar.google.com/calendar/u/0/r?authuser=${GOOGLE_CALENDAR_EMAIL}`}
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-2 px-3.5 py-2 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 text-xs font-semibold shadow-sm transition-all hover:border-slate-300"
            title={`Abrir Google Agenda (${GOOGLE_CALENDAR_EMAIL})`}
          >
            <svg className="w-4 h-4 shrink-0" viewBox="0 0 24 24">
              <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
              <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
              <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" />
              <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z" />
            </svg>
            <span className="hidden sm:inline">Google Agenda</span>
            <span className="text-xs text-slate-400 font-mono hidden lg:inline">({GOOGLE_CALENDAR_EMAIL})</span>
            <ExternalLink size={13} className="text-slate-400 ml-0.5" />
          </a>
          <Button onClick={openCreateModal} className="shrink-0">
            <Plus size={18} /> Novo Compromisso
          </Button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Left column: Calendar & Quick list */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm">
            {/* Calendar Header */}
            <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-lg text-slate-800">
                {monthNames[month]} {year}
              </h3>
              <div className="flex gap-2">
                <Button variant="outline" size="sm" onClick={handlePrevMonth} className="p-2 border-slate-100">
                  <ChevronLeft size={16} />
                </Button>
                <Button variant="outline" size="sm" onClick={handleNextMonth} className="p-2 border-slate-100">
                  <ChevronRight size={16} />
                </Button>
              </div>
            </div>

            {/* Week Days */}
            <div className="grid grid-cols-7 gap-2 text-center text-xs font-bold text-slate-400 mb-2">
              <div>Dom</div>
              <div>Seg</div>
              <div>Ter</div>
              <div>Qua</div>
              <div>Qui</div>
              <div>Sex</div>
              <div>Sáb</div>
            </div>

            {/* Days Grid */}
            <div className="grid grid-cols-7 gap-2">
              {/* Empty offset slots */}
              {Array.from({ length: firstDayIndex }).map((_, idx) => (
                <div key={`empty-${idx}`} className="h-14 bg-slate-50/40 rounded-xl" />
              ))}

              {/* Month Days */}
              {Array.from({ length: daysInMonth }).map((_, idx) => {
                const dayNum = idx + 1;
                const formattedDay = dayNum.toString().padStart(2, '0');
                const formattedMonth = (month + 1).toString().padStart(2, '0');
                const dateStr = `${year}-${formattedMonth}-${formattedDay}`;

                const hasEvents = eventos.some(e => e.data === dateStr);
                const hasPendingEvents = eventos.some(e => e.data === dateStr && e.status === 'Agendado');
                const isSelected = selectedDateStr === dateStr;
                const isToday = new Date().toISOString().split('T')[0] === dateStr;

                return (
                  <button
                    key={`day-${dayNum}`}
                    onClick={() => handleDayClick(dayNum)}
                    className={`h-14 rounded-xl flex flex-col items-center justify-between p-2 text-xs font-semibold relative transition-all border ${
                      isSelected 
                        ? 'bg-emerald-600 border-emerald-600 text-white shadow-lg shadow-emerald-100' 
                        : isToday
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                          : 'bg-white border-slate-50 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    <span>{dayNum}</span>
                    {hasEvents && (
                      <div className="flex gap-1 justify-center items-center">
                        <span className={`w-1.5 h-1.5 rounded-full ${
                          isSelected 
                            ? 'bg-white' 
                            : hasPendingEvents 
                              ? 'bg-sky-500 animate-pulse' 
                              : 'bg-slate-400'
                        }`} />
                      </div>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Detailed events for selected day */}
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-4">
            <div className="border-b border-slate-50 pb-3 flex justify-between items-center">
              <h3 className="font-bold text-slate-800">
                Compromissos para {new Date(selectedDateStr + 'T12:00:00').toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long' })}
              </h3>
              <span className="text-xs text-slate-400 font-mono bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-100">
                {eventsForSelectedDate.length} evento(s)
              </span>
            </div>

            {eventsForSelectedDate.length === 0 ? (
              <div className="text-center py-10 space-y-3">
                <div className="w-12 h-12 bg-slate-50 text-slate-400 rounded-full flex items-center justify-center mx-auto">
                  <CalendarIcon size={20} />
                </div>
                <p className="text-sm text-slate-500 font-medium">Nenhum compromisso agendado para esta data.</p>
                <Button size="sm" variant="outline" onClick={openCreateModal} className="border-slate-200">
                  <Plus size={14} /> Agendar Agora
                </Button>
              </div>
            ) : (
              <div className="space-y-3">
                {eventsForSelectedDate.map(evt => {
                  const company = empresas.find(e => e.id === evt.empresaId);
                  
                  return (
                    <div 
                      key={evt.id} 
                      className={`p-4 rounded-2xl border transition-all hover:shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 ${
                        evt.status === 'Realizado' 
                          ? 'bg-slate-50/40 border-slate-100 text-slate-500' 
                          : evt.status === 'Cancelado'
                            ? 'bg-rose-50/10 border-rose-100 text-slate-400'
                            : 'bg-white border-slate-100 hover:border-slate-200'
                      }`}
                    >
                      <div className="space-y-2 flex-1">
                        <div className="flex items-center gap-2 flex-wrap">
                          <span className={`px-2 py-0.5 rounded-md text-[10px] font-black uppercase tracking-wider ${
                            evt.tipo === 'Sessão de Diagnóstico' 
                              ? 'bg-sky-100 text-sky-700' 
                              : evt.tipo === 'Acompanhamento'
                                ? 'bg-amber-100 text-amber-700'
                                : evt.tipo === 'Feedback'
                                  ? 'bg-teal-100 text-teal-700'
                                  : 'bg-slate-100 text-slate-700'
                          }`}>
                            {evt.tipo}
                          </span>
                          <span className="text-xs text-slate-400 flex items-center gap-1 font-mono font-bold">
                            <Clock size={12} /> {evt.horaInicio}h - {evt.horaFim}h
                          </span>
                          <button 
                            onClick={() => handleToggleStatus(evt)}
                            className={`px-2 py-0.5 rounded-full text-[10px] font-black uppercase cursor-pointer hover:opacity-85 ${
                              evt.status === 'Realizado' 
                                ? 'bg-emerald-100 text-emerald-700' 
                                : evt.status === 'Cancelado'
                                  ? 'bg-rose-100 text-rose-700'
                                  : 'bg-blue-100 text-blue-700'
                            }`}
                          >
                            ● {evt.status}
                          </button>
                        </div>

                        <h4 className={`text-base font-bold ${evt.status === 'Realizado' ? 'line-through text-slate-400' : 'text-slate-800'}`}>
                          {evt.titulo}
                        </h4>

                        {company && (
                          <p className="text-xs text-slate-500 flex items-center gap-1.5 font-semibold">
                            <Building2 size={13} className="text-slate-400" /> {company.nome}
                          </p>
                        )}

                        {evt.descricao && (
                          <p className="text-xs text-slate-400 leading-relaxed bg-slate-50 p-2 rounded-lg mt-1 border border-slate-100/30">
                            {evt.descricao}
                          </p>
                        )}
                      </div>

                      <div className="flex items-center gap-1 shrink-0 md:self-center self-end">
                        <a
                          href={getGoogleCalendarUrl(evt)}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-flex items-center justify-center p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-colors"
                          title="Adicionar / Ver no Google Agenda"
                        >
                          <svg className="w-4 h-4" viewBox="0 0 24 24">
                            <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                            <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                            <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" />
                            <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z" />
                          </svg>
                        </a>
                        <Button 
                          variant="ghost" 
                          size="sm" 
                          className="p-1.5 text-slate-400 hover:text-sky-600 hover:bg-slate-100"
                          onClick={() => handleEditEvent(evt)}
                        >
                          <FileText size={16} />
                        </Button>
                        <Button 
                          variant="ghost" 
                          size="sm" 
                          className="p-1.5 text-slate-400 hover:text-rose-500 hover:bg-slate-100"
                          onClick={() => handleDeleteEvent(evt.id)}
                        >
                          <Trash2 size={16} />
                        </Button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>

        {/* Right column: Upcoming agenda summary */}
        <div className="space-y-6">
          <div className="bg-white rounded-3xl p-6 border border-slate-100 shadow-sm space-y-6">
            <div>
              <h3 className="font-bold text-slate-800 flex items-center gap-2">
                <ListTodo size={18} className="text-sky-500" />
                Compromissos Futuros
              </h3>
              <p className="text-slate-400 text-xs">Visão geral cronológica de reuniões</p>
            </div>

            {upcomingEvents.length === 0 ? (
              <p className="text-xs text-slate-400 italic">Sem reuniões futuras programadas.</p>
            ) : (
              <div className="space-y-4">
                {upcomingEvents.map(evt => {
                  const company = empresas.find(e => e.id === evt.empresaId);
                  const isToday = new Date().toISOString().split('T')[0] === evt.data;
                  const dateFormatted = new Date(evt.data + 'T12:00:00').toLocaleDateString('pt-BR', { day: 'numeric', month: 'short' });

                  return (
                    <div key={evt.id} className="flex gap-4 items-start pb-4 border-b border-slate-50 last:border-b-0 last:pb-0">
                      <div className={`w-12 h-12 rounded-xl flex flex-col items-center justify-center shrink-0 border ${
                        isToday 
                          ? 'bg-emerald-50 border-emerald-200 text-emerald-700 font-bold' 
                          : 'bg-slate-50 border-slate-100 text-slate-600'
                      }`}>
                        <span className="text-[10px] uppercase font-bold tracking-widest">{isToday ? 'Hoje' : dateFormatted.split(' ')[2]}</span>
                        <span className="text-xs font-extrabold">{isToday ? evt.horaInicio : dateFormatted.split(' ')[0]}</span>
                      </div>
                      <div className="min-w-0 flex-1">
                        <h4 className="text-xs font-bold text-slate-800 truncate">{evt.titulo}</h4>
                        <p className="text-[10px] text-slate-400 mt-1 font-mono">{evt.horaInicio}h - {evt.horaFim}h</p>
                        {company && (
                          <p className="text-[10px] text-slate-500 font-semibold truncate mt-0.5">{company.nome}</p>
                        )}
                      </div>
                      <a
                        href={getGoogleCalendarUrl(evt)}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="p-1.5 text-slate-300 hover:text-emerald-600 hover:bg-slate-50 rounded-lg shrink-0 transition-colors"
                        title="Abrir no Google Agenda"
                      >
                        <ExternalLink size={13} />
                      </a>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Save Event Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <motion.div 
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white rounded-3xl p-6 border border-slate-100 shadow-2xl max-w-3xl w-full space-y-6"
          >
            <div className="border-b border-slate-100 pb-3 flex justify-between items-center">
              <h3 className="text-lg font-bold text-slate-800">
                {editingEventId ? 'Editar Compromisso' : 'Novo Compromisso'}
              </h3>
              <button 
                onClick={() => setIsModalOpen(false)}
                className="p-1 hover:bg-slate-100 text-slate-400 hover:text-slate-700 rounded-lg"
              >
                <XCircle size={20} />
              </button>
            </div>

            <form onSubmit={handleSubmitEvent} className="space-y-4">
              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-500">Título do Evento</label>
                <input
                  type="text"
                  required
                  placeholder="Ex: Alinhamento de Metas de IA"
                  value={formTitulo}
                  onChange={(e) => setFormTitulo(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium text-slate-800 focus:outline-none focus:border-emerald-500"
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Data</label>
                  <input
                    type="date"
                    required
                    value={formDate}
                    onChange={(e) => setFormDate(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Hora Início</label>
                  <input
                    type="time"
                    required
                    value={formHoraInicio}
                    onChange={(e) => setFormHoraInicio(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Hora Fim</label>
                  <input
                    type="time"
                    required
                    value={formHoraFim}
                    onChange={(e) => setFormHoraFim(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Empresa / Cliente</label>
                  <select
                    value={formEmpresaId}
                    onChange={(e) => {
                      setFormEmpresaId(e.target.value);
                      setFormDiagnosticoId('');
                    }}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  >
                    <option value="">Selecione um cliente (Opcional)...</option>
                    {empresas.map(emp => (
                      <option key={emp.id} value={emp.id}>{emp.nome}</option>
                    ))}
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Projeto / Diagnóstico</label>
                  <select
                    value={formDiagnosticoId}
                    onChange={(e) => setFormDiagnosticoId(e.target.value)}
                    disabled={!formEmpresaId}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none disabled:opacity-50"
                  >
                    <option value="">Selecione o projeto (Opcional)...</option>
                    {filteredDiagnostics.map(diag => (
                      <option key={diag.id} value={diag.id}>
                        Projeto: {diag.tipoEmpresa || 'Geral'} ({diag.id.slice(0, 5)})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Tipo de Compromisso</label>
                  <select
                    value={formTipo}
                    onChange={(e) => setFormTipo(e.target.value as any)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  >
                    <option value="Reunião">Reunião</option>
                    <option value="Sessão de Diagnóstico">Sessão de Diagnóstico</option>
                    <option value="Acompanhamento">Acompanhamento</option>
                    <option value="Feedback">Feedback</option>
                    <option value="Outro">Outro</option>
                  </select>
                </div>

                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500">Status</label>
                  <select
                    value={formStatus}
                    onChange={(e) => setFormStatus(e.target.value as any)}
                    className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                  >
                    <option value="Agendado">Agendado</option>
                    <option value="Realizado">Realizado</option>
                    <option value="Cancelado">Cancelado</option>
                  </select>
                </div>
              </div>

              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-500">Descrição / Pauta</label>
                <textarea
                  rows={2}
                  placeholder="Ex: Alinhamento das premissas de marketing e aprovação do plano..."
                  value={formDescricao}
                  onChange={(e) => setFormDescricao(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-xs font-medium text-slate-800 focus:outline-none"
                />
              </div>

              <div className="flex items-center justify-between p-3.5 bg-emerald-50/60 border border-emerald-100 rounded-2xl">
                <div className="flex items-center gap-3">
                  <div className="w-8 h-8 rounded-xl bg-white border border-emerald-200 flex items-center justify-center shadow-xs">
                    <svg className="w-4 h-4" viewBox="0 0 24 24">
                      <path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" />
                      <path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" />
                      <path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l2.85-2.22.81-.63z" />
                      <path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.84c.87-2.6 3.3-4.52 6.16-4.52z" />
                    </svg>
                  </div>
                  <div>
                    <p className="text-xs font-bold text-slate-800">Sincronizar com Google Agenda</p>
                    <p className="text-[11px] text-slate-500">Abrir pré-preenchido para <span className="font-mono font-medium text-emerald-800">{GOOGLE_CALENDAR_EMAIL}</span></p>
                  </div>
                </div>
                <label className="relative inline-flex items-center cursor-pointer">
                  <input
                    type="checkbox"
                    checked={syncToGoogleOnSave}
                    onChange={(e) => setSyncToGoogleOnSave(e.target.checked)}
                    className="sr-only peer"
                  />
                  <div className="w-9 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-emerald-600"></div>
                </label>
              </div>

              <div className="pt-4 border-t border-slate-100 flex justify-end gap-3">
                <Button type="button" variant="outline" onClick={() => setIsModalOpen(false)}>
                  Cancelar
                </Button>
                <Button type="submit">
                  {editingEventId ? 'Salvar Alterações' : 'Criar Compromisso'}
                </Button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
};
