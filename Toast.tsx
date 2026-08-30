import React from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { CheckCircle2, AlertCircle, Info, X } from 'lucide-react';

export interface ToastItem {
  id: string;
  message: string;
  type?: 'success' | 'error' | 'info';
  title?: string;
}

interface ToastContainerProps {
  toasts: ToastItem[];
  onDismiss: (id: string) => void;
}

export const ToastContainer: React.FC<ToastContainerProps> = ({ toasts, onDismiss }) => {
  return (
    <div className="fixed top-5 right-5 z-[99999] flex flex-col gap-2.5 max-w-md w-full pointer-events-none px-4">
      <AnimatePresence>
        {toasts.map((t) => (
          <motion.div
            key={t.id}
            initial={{ opacity: 0, y: -16, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, scale: 0.9, y: -10 }}
            transition={{ duration: 0.18 }}
            className={`pointer-events-auto flex items-start gap-3 p-4 rounded-2xl shadow-xl border backdrop-blur-md ${
              t.type === 'error'
                ? 'bg-rose-900/95 text-white border-rose-700/80 shadow-rose-950/20'
                : t.type === 'info'
                ? 'bg-slate-900/95 text-white border-slate-700/80 shadow-slate-950/20'
                : 'bg-emerald-900/95 text-white border-emerald-700/80 shadow-emerald-950/20'
            }`}
          >
            <div className="shrink-0 mt-0.5">
              {t.type === 'error' ? (
                <div className="w-6 h-6 rounded-full bg-rose-500/20 flex items-center justify-center text-rose-300">
                  <AlertCircle size={16} />
                </div>
              ) : t.type === 'info' ? (
                <div className="w-6 h-6 rounded-full bg-sky-500/20 flex items-center justify-center text-sky-300">
                  <Info size={16} />
                </div>
              ) : (
                <div className="w-6 h-6 rounded-full bg-emerald-500/30 flex items-center justify-center text-emerald-300">
                  <CheckCircle2 size={16} />
                </div>
              )}
            </div>

            <div className="flex-1 min-w-0 pr-1">
              {t.title && <h4 className="text-xs font-bold uppercase tracking-wider opacity-80 mb-0.5">{t.title}</h4>}
              <p className="text-sm font-medium leading-snug whitespace-pre-line text-slate-100">{t.message}</p>
            </div>

            <button
              onClick={() => onDismiss(t.id)}
              className="shrink-0 p-1 rounded-lg hover:bg-white/10 text-white/60 hover:text-white transition-colors cursor-pointer"
            >
              <X size={14} />
            </button>
          </motion.div>
        ))}
      </AnimatePresence>
    </div>
  );
};
