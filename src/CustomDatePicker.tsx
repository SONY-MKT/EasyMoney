import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Calendar, ChevronDown, ChevronLeft, ChevronRight } from 'lucide-react';

export default function CustomDatePicker({ 
  value, 
  onChange, 
  language,
  todayYMD
}: { 
  value: string, 
  onChange: (d: string) => void, 
  language: string,
  todayYMD: string
}) {
  const [isOpen, setIsOpen] = useState(false);
  
  const parts = value ? value.split('-') : [];
  const initialDate = parts.length === 3 ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10)-1, parseInt(parts[2], 10)) : new Date();
  
  const [currentMonth, setCurrentMonth] = useState(new Date(initialDate.getFullYear(), initialDate.getMonth(), 1));

  // Sync current month when popup opens to show the currently selected date
  useEffect(() => {
    if (isOpen && parts.length === 3) {
      setCurrentMonth(new Date(parseInt(parts[0], 10), parseInt(parts[1], 10)-1, 1));
    }
  }, [isOpen]);

  const getDisplayDate = (dString: string) => {
    if (!dString) return '';
    try {
      const parts = dString.split('-');
      if (parts.length !== 3) return dString;
      const d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      const isToday = dString === todayYMD;
      
      let formatted = '';
      if (language === 'English') {
        formatted = d.toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' });
        return isToday ? `Today • ${formatted}` : formatted;
      } else {
        const months = ['មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];
        formatted = `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
        return isToday ? `ថ្ងៃនេះ • ${formatted}` : formatted;
      }
    } catch {
      return dString;
    }
  };

  const daysInMonth = (year: number, month: number) => new Date(year, month + 1, 0).getDate();
  const getFirstDayOfMonth = (year: number, month: number) => new Date(year, month, 1).getDay();

  const handlePrevMonth = (e: React.MouseEvent) => {
    e.stopPropagation();
    setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1));
  };

  const handleNextMonth = (e: React.MouseEvent) => {
    e.stopPropagation();
    setCurrentMonth(new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1));
  };
  
  const handleSelectDate = (day: number) => {
    const dStr = `${currentMonth.getFullYear()}-${String(currentMonth.getMonth() + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    onChange(dStr);
    setIsOpen(false);
  };

  const renderCalendar = () => {
    const year = currentMonth.getFullYear();
    const month = currentMonth.getMonth();
    const numDays = daysInMonth(year, month);
    const firstDay = getFirstDayOfMonth(year, month);
    
    const days = [];
    // empty slots
    for (let i = 0; i < firstDay; i++) {
      days.push(<div key={`empty-${i}`} className="h-9 w-9"></div>);
    }
    
    // actual days
    for (let i = 1; i <= numDays; i++) {
        const dStr = `${year}-${String(month+1).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
        const isSelected = dStr === value;
        const isToday = dStr === todayYMD;
        
        days.push(
            <button
                key={`day-${i}`}
                type="button"
                onClick={(e) => { e.preventDefault(); e.stopPropagation(); handleSelectDate(i); }}
                className={`relative h-9 w-9 flex items-center justify-center text-sm rounded-full transition-all active:scale-95 ${
                    isSelected ? 'bg-white text-[#111] font-bold shadow-sm z-10' : 
                    isToday ? 'text-white font-bold bg-[#4F5256] border border-emerald-500/50' :
                    'text-zinc-100 hover:bg-[#4F5256]'
                }`}
            >
                {i}
                {isToday && !isSelected && <span className="absolute bottom-1 w-1 h-1 bg-emerald-400 rounded-full"></span>}
            </button>
        );
    }
    return days;
  };

  const monthNames = language === 'English' 
    ? ['January','February','March','April','May','June','July','August','September','October','November','December']
    : ['មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];

  const weekDays = language === 'English'
    ? ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']
    : ['អា', 'ច', 'អ', 'ព', 'ព្រ', 'សុ', 'ស'];

  return (
    <>
      <div 
        className="date-picker-container relative group w-full h-[54px] cursor-pointer outline-none bg-transparent"
        onClick={() => setIsOpen(true)}
      >
        <div className={`absolute inset-0 flex items-center justify-between bg-[#FBFDFB] border-2 border-[#E8EEE9] hover:border-emerald-400 px-4 rounded-[20px] text-sm outline-none transition-all font-bold text-emerald-950 active:scale-[0.98] overflow-hidden shadow-sm ${isOpen ? 'ring-4 ring-emerald-500/10 border-emerald-500' : ''}`}>
          <div className="flex items-center gap-3 overflow-hidden">
            <Calendar className="w-5 h-5 text-slate-500 shrink-0" />
            <span className="truncate">{getDisplayDate(value)}</span>
          </div>
          <ChevronDown className="w-4 h-4 text-slate-400 shrink-0 transition-transform duration-300" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)' }} />
        </div>
      </div>
      
      <AnimatePresence>
        {isOpen && (
            <div className="fixed inset-0 z-[200] flex items-center justify-center font-sans p-4">
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 transition={{ duration: 0.2 }}
                 className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm"
                 onClick={() => setIsOpen(false)}
              />
              <motion.div 
                 initial={{ scale: 0.95, opacity: 0, y: 10 }}
                 animate={{ scale: 1, opacity: 1, y: 0 }}
                 exit={{ scale: 0.95, opacity: 0, y: 10 }}
                 transition={{ type: "spring", damping: 25, stiffness: 300 }}
                 className="relative w-full max-w-[340px] bg-[#3F3F3F] rounded-[28px] overflow-hidden shadow-[0_20px_60px_rgba(0,0,0,0.4)]"
              >
                  <div className="p-5">
                    {/* Header */}
                    <div className="flex items-center justify-between mb-6 text-white font-medium">
                        <button type="button" onClick={handlePrevMonth} className="p-2 -ml-2 rounded-full hover:bg-white/10 transition-colors active:scale-95 touch-manipulation">
                            <ChevronLeft className="w-5 h-5 text-zinc-300" />
                        </button>
                        <div className="text-base font-semibold text-white tracking-wide">
                            {monthNames[currentMonth.getMonth()]} {currentMonth.getFullYear()}
                        </div>
                        <button type="button" onClick={handleNextMonth} className="p-2 -mr-2 rounded-full hover:bg-white/10 transition-colors active:scale-95 touch-manipulation">
                            <ChevronRight className="w-5 h-5 text-zinc-300" />
                        </button>
                    </div>

                    {/* Weekdays */}
                    <div className="grid grid-cols-7 gap-1 mb-3">
                         {weekDays.map((d, i) => (
                             <div key={d+i} className="text-center text-[10px] font-bold text-[#A8A8A8] uppercase tracking-widest">
                                 {d}
                             </div>
                         ))}
                    </div>

                    {/* Grid */}
                    <div className="grid grid-cols-7 gap-1 place-items-center">
                         {renderCalendar()}
                    </div>
                  </div>
              </motion.div>
            </div>
        )}
      </AnimatePresence>
    </>
  );
}
