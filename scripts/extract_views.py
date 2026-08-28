with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Extract RelatorioView
rel_lines = lines[5465:6640]
with open('src/components/RelatorioView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Building2, Calendar, FileText, Download, Printer, CheckCircle2,
  AlertCircle, TrendingUp, Sparkles, User, Award, Layers, PieChart,
  BarChart2, Target, Check, RefreshCw, ChevronDown, ChevronUp, Clock
} from 'lucide-react';
import { 
  ResponsiveContainer, RadialBarChart, RadialBar, PolarAngleAxis,
  RadarChart, PolarGrid, PolarRadiusAxis, Radar, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, AreaChart, Area
} from 'recharts';
import { cn } from '../lib/utils';
import { Empresa, Diagnostico, Resposta, Problema, Solucao, TarefaPlano, AtividadeCronograma } from '../types';
import { getCorporateColor, getCorporateRgb } from '../lib/theme';

''')
    for l in rel_lines:
        out.write(l)
    out.write('\nexport default RelatorioView;\n')

# Extract GestaoPlanoView
gp_lines = lines[6640:7824]
with open('src/components/GestaoPlanoView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Calendar, CheckCircle2, Clock, AlertCircle, Plus, Trash2, Edit2,
  Download, Sparkles, Filter, Search, Check, RefreshCw, Layers,
  ChevronRight, ArrowUpDown, ChevronDown, CheckSquare, Upload, X,
  FileText, User, Tag
} from 'lucide-react';
import { cn } from '../lib/utils';
import { Empresa, Diagnostico, TarefaPlano, Problema, Solucao } from '../types';
import { getCorporateColor, getCorporateRgb } from '../lib/theme';

''')
    for l in gp_lines:
        out.write(l)
    out.write('\nexport default GestaoPlanoView;\n')

print('Extracted RelatorioView and GestaoPlanoView!')
