with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Extract HomeView
home_lines = lines[9915:10186]
with open('src/components/HomeView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Building2, Activity, Calendar, ArrowRight, TrendingUp, 
  CheckCircle2, Clock, AlertCircle, FileText, Target, Users,
  Sparkles, CheckSquare, Plus, ExternalLink
} from 'lucide-react';
import { Empresa, Diagnostico, TarefaPlano } from '../types';

export interface HomeViewProps {
  user: any;
  userProfile: any;
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  tarefasPlano: TarefaPlano[];
  onNavigate: (view: string, data?: any) => void;
  onNewCompany: () => void;
  onNewDiagnosis: () => void;
  isAdmin?: boolean;
}

''')
    for l in home_lines:
        out.write(l)
    out.write('\nexport default HomeView;\n')

print('HomeView extracted successfully!')
