import React from 'react';
import { motion } from 'motion/react';
import { X } from 'lucide-react';
import { Button } from './Button';

interface ModalProps {
  title: string;
  children: React.ReactNode;
  onClose: () => void;
  onConfirm: (e: React.MouseEvent<HTMLButtonElement>) => void | Promise<any>;
  confirmText?: string;
  variant?: 'primary' | 'secondary' | 'danger' | 'outline' | 'ghost' | 'success';
  disabled?: boolean;
  size?: 'sm' | 'md' | 'lg' | 'xl' | '2xl' | '3xl' | '4xl' | '5xl' | '6xl';
  customWidth?: string;
}

const sizeClasses = {
  sm: 'max-w-sm',
  md: 'max-w-md',
  lg: 'max-w-lg',
  xl: 'max-w-xl',
  '2xl': 'max-w-2xl',
  '3xl': 'max-w-3xl',
  '4xl': 'max-w-4xl',
  '5xl': 'max-w-5xl',
  '6xl': 'max-w-6xl'
};

export const Modal = ({ 
  title, 
  children, 
  onClose, 
  onConfirm, 
  confirmText = "Confirmar", 
  variant = "primary", 
  disabled = false,
  size = 'md',
  customWidth
}: ModalProps) => (
  <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm animate-in fade-in duration-150">
    <motion.div 
      initial={{ scale: 0.96, opacity: 0 }}
      animate={{ scale: 1, opacity: 1 }}
      exit={{ scale: 0.96, opacity: 0 }}
      transition={{ duration: 0.15 }}
      style={customWidth ? { maxWidth: customWidth } : undefined}
      className={`bg-white rounded-2xl shadow-2xl border border-slate-100 ${customWidth ? '' : sizeClasses[size]} w-full overflow-hidden flex flex-col max-h-[90vh]`}
    >
      <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between shrink-0">
        <h3 className="font-bold text-lg text-slate-800">{title}</h3>
        <button 
          onClick={onClose} 
          className="p-1.5 hover:bg-slate-100 active:scale-95 rounded-full text-slate-400 hover:text-slate-600 transition-all cursor-pointer"
        >
          <X size={18} />
        </button>
      </div>
      <div className="p-6 overflow-y-auto flex-1">
        {children}
      </div>
      <div className="px-6 py-4 bg-slate-50 flex items-center justify-end gap-3 shrink-0 border-t border-slate-100/80">
        <Button variant="secondary" onClick={onClose}>Cancelar</Button>
        <Button variant={variant} onClick={onConfirm} disabled={disabled}>{confirmText}</Button>
      </div>
    </motion.div>
  </div>
);

