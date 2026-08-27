import React from 'react';
import { cn } from '../lib/utils';

export interface CardProps extends React.HTMLAttributes<HTMLDivElement> {
  children: React.ReactNode;
  className?: string;
}

export const Card: React.FC<CardProps> = ({ children, className, ...props }) => {
  return (
    <div
      className={cn(
        "bg-white border border-slate-200/80 rounded-2xl shadow-sm transition-all",
        className
      )}
      {...props}
    >
      {children}
    </div>
  );
};
