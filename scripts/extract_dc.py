import os

with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Extract DadosConsultoriaView
dc_lines = lines[1886:2653]
with open('src/components/DadosConsultoriaView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect } from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Building2, User, Phone, Mail, FileText, Target, Award,
  Clock, CheckCircle2, AlertCircle, Save, Sparkles, Copy,
  Calendar, Check, Download, ExternalLink, RefreshCw, Layers
} from 'lucide-react';
import { cn } from '../lib/utils';
import { Empresa, Diagnostico, DadosConsultoria } from '../types';
import { CONSULTORIA_AREAS } from '../constants/consultoriaConstants';

''')
    for l in dc_lines:
        out.write(l)
    out.write('\nexport default DadosConsultoriaView;\n')

print('DadosConsultoriaView extracted!')
