import React from 'react';
import { Check, Loader2 } from 'lucide-react';
import { cn } from '../lib/utils';
import { playClickSound } from '../lib/sound';

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  children?: React.ReactNode;
  variant?: 'primary' | 'secondary' | 'danger' | 'outline' | 'ghost' | 'success';
  size?: 'sm' | 'md' | 'lg' | 'icon';
  className?: string;
  onClick?: (e: React.MouseEvent<HTMLButtonElement>) => any;
  type?: 'button' | 'submit' | 'reset';
  disabled?: boolean;
  title?: string;
  loading?: boolean;
  showFeedback?: boolean;
  successText?: string;
  signalCompleted?: boolean;
}

export const Button = ({ 
  children, 
  className, 
  variant = 'primary', 
  size = 'md', 
  onClick, 
  type = 'button', 
  disabled = false, 
  title, 
  loading = false,
  showFeedback = false,
  successText,
  signalCompleted = false,
  ...props 
}: ButtonProps) => {

  const variants: Record<string, string> = {
    primary: 'bg-emerald-600 text-white hover:bg-emerald-700 active:bg-emerald-800 shadow-sm hover:shadow active:scale-[0.98]',
    secondary: 'bg-slate-100 text-slate-700 hover:bg-slate-200 active:bg-slate-300 active:scale-[0.98]',
    danger: 'bg-rose-600 text-white hover:bg-rose-700 active:bg-rose-800 shadow-sm active:scale-[0.98]',
    outline: 'border border-slate-200 text-slate-600 hover:bg-slate-50 active:bg-slate-100 active:scale-[0.98]',
    ghost: 'text-slate-500 hover:bg-slate-100 active:bg-slate-200 active:scale-[0.98]',
    success: 'bg-emerald-600 text-white hover:bg-emerald-700 active:scale-[0.98]'
  };

  const sizes: Record<string, string> = {
    sm: 'px-3 py-1.5 text-xs rounded-lg',
    md: 'px-4 py-2 text-sm rounded-xl',
    lg: 'px-6 py-3 text-base rounded-xl',
    icon: 'p-2 rounded-xl'
  };

  const handleClick = (e: React.MouseEvent<HTMLButtonElement>) => {
    if (disabled || loading) return;

    try {
      playClickSound();
    } catch {}

    if (onClick) {
      onClick(e);
    }
  };

  return (
    <button
      className={cn(
        'font-medium transition-all duration-100 relative select-none inline-flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed disabled:active:scale-100',
        variants[variant], 
        sizes[size],
        signalCompleted && 'ring-2 ring-emerald-400 !bg-emerald-600 !text-white shadow-md',
        className
      )} 
      onClick={handleClick}
      type={type}
      disabled={disabled || loading}
      title={title}
      {...props}
    >
      {loading ? (
        <>
          <Loader2 className="w-4 h-4 animate-spin shrink-0" />
          <span className="opacity-95 text-xs">Processando...</span>
        </>
      ) : signalCompleted ? (
        <>
          <Check className="w-4 h-4 text-white stroke-[3] animate-in zoom-in-75 duration-100 shrink-0" />
          <span className="animate-in fade-in duration-100">{successText || children}</span>
        </>
      ) : (
        children
      )}
    </button>
  );
};


