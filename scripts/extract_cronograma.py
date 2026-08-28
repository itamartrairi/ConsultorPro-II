with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

cron_lines = lines[2653:5465]
with open('src/components/CronogramaView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Calendar, Clock, CheckCircle2, AlertCircle, Plus, Trash2, Edit2,
  Download, Sparkles, Filter, Search, Check, RefreshCw, Layers,
  ChevronRight, ArrowUpDown, ChevronDown, CheckSquare, Upload, X,
  FileText, User, Tag, ArrowLeft, ArrowRight, Play, Eye
} from 'lucide-react';
import { cn } from '../lib/utils';
import { Empresa, Diagnostico, AtividadeCronograma, Resposta, Problema, Solucao } from '../types';
import { getCorporateColor, getCorporateRgb } from '../lib/theme';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';

''')
    for l in cron_lines:
        out.write(l)
    out.write('\nexport default CronogramaView;\n')

print('CronogramaView extracted successfully!')
