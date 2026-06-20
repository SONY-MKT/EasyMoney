import React, { useState, useEffect, useRef } from 'react';
import { ChevronDown, CheckCircle, X } from 'lucide-react';
import { cn } from './lib/utils';

export interface SelectOption {
  value: string | number;
  label: string;
  icon?: React.ReactNode;
  description?: string;
}

export interface CustomSelectProps {
  options: SelectOption[];
  value: string | number;
  onChange: (value: any) => void;
  placeholder?: string;
  icon?: any;
  className?: string;
  modalTitle?: string;
  alignRight?: boolean;
}

export function CustomSelect({ options, value, onChange, placeholder, icon: Icon, className, modalTitle = "ជ្រើសរើស (Select)", alignRight = false }: CustomSelectProps) {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  
  const selectedOption = options.find((o) => o.value === value);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    }
    
    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isOpen]);

  return (
    <div className={cn("relative w-full", isOpen ? "z-[9999]" : "z-10")} ref={containerRef}>
      <div 
        onClick={() => setIsOpen(!isOpen)}
        className={cn("relative flex items-center justify-between cursor-pointer w-full bg-white border border-slate-200 rounded-xl px-3 py-2.5 font-bold text-slate-700 outline-none hover:border-emerald-400 focus:ring-2 focus:ring-emerald-500/20 transition-all shadow-sm appearance-none text-xs", className)}
      >
        <div className="flex items-center gap-2 truncate">
          {Icon && <Icon className="w-4 h-4 text-slate-400 shrink-0" />}
          <span className="truncate flex-1">{selectedOption ? selectedOption.label : placeholder}</span>
        </div>
        <ChevronDown className={cn("w-3.5 h-3.5 text-slate-400 shrink-0 transition-transform duration-200", isOpen && "rotate-180")} />
      </div>

      {isOpen && (
        <div className={cn("absolute top-[calc(100%+8px)] z-[100] bg-white/95 backdrop-blur-xl rounded-2xl shadow-[0_10px_40px_-10px_rgba(0,0,0,0.1)] border border-slate-100/50 overflow-hidden animate-in fade-in zoom-in-95 duration-200", alignRight ? "right-0 min-w-full w-max" : "left-0 min-w-full w-max")}>
          <div className="max-h-[300px] overflow-y-auto overscroll-contain flex flex-col p-1.5 custom-scrollbar">
            {options.map((opt) => {
              const isActive = opt.value === value;
              return (
                <div 
                  key={opt.value}
                  onClick={() => { onChange(opt.value); setIsOpen(false); }}
                  className={cn(
                    "flex items-center justify-between px-3.5 py-2.5 rounded-xl cursor-pointer transition-colors text-[11px] font-extrabold uppercase tracking-wide",
                    isActive ? "bg-emerald-50 text-emerald-700" : "text-slate-600 hover:bg-slate-50/80"
                  )}
                >
                  <div className="flex items-center gap-2.5 truncate">
                    {opt.icon && <div className={cn("flex items-center justify-center shrink-0", isActive ? "text-emerald-600" : "text-slate-400")}>{opt.icon}</div>}
                    <div className="truncate">
                      <span>{opt.label}</span>
                      {opt.description && <span className="text-[10px] text-slate-400 ml-1.5 font-normal tracking-wide hidden sm:inline-block">{opt.description}</span>}
                    </div>
                  </div>
                  {isActive && <CheckCircle className="w-3.5 h-3.5 text-emerald-500 shrink-0 ml-2" />}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
