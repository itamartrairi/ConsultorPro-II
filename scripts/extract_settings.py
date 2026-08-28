with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

settings_lines = lines[8389:9259]
with open('src/components/SettingsView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect, useCallback } from 'react';
import { motion } from 'motion/react';
import { Button } from './Button';
import { Card } from './Card';
import { CorporateColorPicker } from './CorporateColorPicker';
import { 
  ImageIcon, Upload, X, Bell, Download, Sparkles, RefreshCw, 
  CheckCircle2, AlertCircle, Key, EyeOff, Eye, ExternalLink, 
  Cloud, Lock, ShieldCheck, Info 
} from 'lucide-react';
import { cn } from '../lib/utils';
import { checkCloudConnection } from '../firebase';
import { encryptedLocalStorage } from '../lib/cryptoStorage';

export interface SyncSummary {
  pushed: { empresas: number; diagnosticos: number; respostas: number; tarefas: number };
  pulled: { empresas: number; diagnosticos: number; respostas: number; tarefas: number };
  timestamp: Date;
}

''')
    for l in settings_lines:
        out.write(l)
    out.write('\nexport default SettingsView;\n')

print('Extracted SettingsView.tsx!')
