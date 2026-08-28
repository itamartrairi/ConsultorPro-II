import os

os.makedirs('src/components', exist_ok=True)
os.makedirs('src/constants', exist_ok=True)

with open('src/App.tsx', 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# Extract LandingPage
landing_lines = lines[7824:7986]
with open('src/components/LandingPage.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React from 'react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Sparkles, CheckCircle2, TrendingUp, ShieldCheck, 
  FileText, ArrowRight, Zap, Target, Users, BarChart3
} from 'lucide-react';

''')
    for l in landing_lines:
        out.write(l)
    out.write('\nexport default LandingPage;\n')

# Extract CheckoutPage
checkout_lines = lines[7986:8389]
with open('src/components/CheckoutPage.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState } from 'react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  ShieldCheck, Check, ArrowLeft, Lock, 
  CreditCard, Sparkles, AlertCircle, CheckCircle2
} from 'lucide-react';
import { cn } from '../lib/utils';

''')
    for l in checkout_lines:
        out.write(l)
    out.write('\nexport default CheckoutPage;\n')

# Extract LicenseManagementView
license_lines = lines[9259:9801]
with open('src/components/LicenseManagementView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect } from 'react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Key, Users, Plus, Trash2, CheckCircle2, AlertCircle, 
  RefreshCw, Copy, Check, Shield, Search, Filter, ShieldCheck,
  Calendar, Clock, Edit2
} from 'lucide-react';
import { cn } from '../lib/utils';
import { db } from '../firebase';
import { 
  collection, query, getDocs, addDoc, updateDoc, 
  deleteDoc, doc, Timestamp, orderBy 
} from 'firebase/firestore';

''')
    for l in license_lines:
        out.write(l)
    out.write('\nexport default LicenseManagementView;\n')

# Extract LicensingDataView
licensing_data_lines = lines[9801:9915]
with open('src/components/LicensingDataView.tsx', 'w', encoding='utf-8') as out:
    out.write('''import React, { useState, useEffect } from 'react';
import { Button } from './Button';
import { Card } from './Card';
import { 
  Building2, Users, FileSpreadsheet, Download, 
  RefreshCw, Search, CheckCircle2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { db } from '../firebase';
import { collection, query, getDocs, orderBy } from 'firebase/firestore';

''')
    for l in licensing_data_lines:
        out.write(l)
    out.write('\nexport default LicensingDataView;\n')

print('Extracted LandingPage, CheckoutPage, LicenseManagementView, LicensingDataView!')
