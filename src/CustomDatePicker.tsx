import React, { useState, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { Calendar, ChevronDown, ChevronLeft, ChevronRight } from 'lucide-react';

const ScrollWheel = ({ items, selectedValue, onChange, label }: any) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const itemHeight = 44; 
  const [localValue, setLocalValue] = useState(selectedValue);

  useEffect(() => {
    const idx = items.findIndex((i: any) => i.value === selectedValue);
    if (idx !== -1 && containerRef.current) {
      containerRef.current.scrollTop = idx * itemHeight;
    }
  }, []);

  const handleScroll = (e: React.UIEvent<HTMLDivElement>) => {
    const idx = Math.round(e.currentTarget.scrollTop / itemHeight);
    if (items[idx] && items[idx].value !== localValue) {
       setLocalValue(items[idx].value);
       onChange(items[idx].value);
    }
  };

  return (
    <div className="flex flex-col items-center flex-1">
      <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-3">{label}</div>
      <div className="relative w-full h-[220px] overflow-hidden rounded-[16px]">
         {/* Transparent gradients to fade top and bottom */}
         <div className="absolute inset-0 pointer-events-none z-20" style={{
           background: 'linear-gradient(to bottom, rgba(255,255,255,1) 0%, rgba(255,255,255,0) 30%, rgba(255,255,255,0) 70%, rgba(255,255,255,1) 100%)'
         }} />
         
         {/* Center highlighted row background */}
         <div className="absolute top-1/2 left-0 right-0 h-[44px] -translate-y-1/2 bg-[#F0F4F1] rounded-xl pointer-events-none z-0" />
         
         <div 
           ref={containerRef}
           className="h-full w-full overflow-y-auto snap-y snap-mandatory relative z-10 no-scrollbar"
           onScroll={handleScroll}
           style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}
         >
            <div style={{ height: 88 }} className="snap-center" /> 
            {items.map((item: any) => (
               <div 
                 key={item.value} 
                 style={{ height: 44 }}
                 className={`snap-center flex items-center justify-center transition-all duration-200 ${item.value === localValue ? 'text-[20px] font-bold text-emerald-900' : 'text-sm font-semibold text-slate-300 scale-95'}`}
               >
                  {item.label}
               </div>
            ))}
            <div style={{ height: 88 }} className="snap-center" />
         </div>
      </div>
    </div>
  );
};

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
  const [isMobile, setIsMobile] = useState(false);
  
  const parts = value ? value.split('-') : [];
  const initialDate = parts.length === 3 ? new Date(parseInt(parts[0], 10), parseInt(parts[1], 10)-1, parseInt(parts[2], 10)) : new Date();
  
  const [currentMonth, setCurrentMonth] = useState(new Date(initialDate.getFullYear(), initialDate.getMonth(), 1));

  // Mobile wheel state
  const [mYear, setMYear] = useState(initialDate.getFullYear());
  const [mMonth, setMMonth] = useState(initialDate.getMonth() + 1);
  const [mDay, setMDay] = useState(initialDate.getDate());

  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.matchMedia('(pointer: coarse)').matches || window.innerWidth < 768);
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  // Sync state when opened
  useEffect(() => {
    if (isOpen && parts.length === 3) {
      setCurrentMonth(new Date(parseInt(parts[0], 10), parseInt(parts[1], 10)-1, 1));
      setMYear(parseInt(parts[0], 10));
      setMMonth(parseInt(parts[1], 10));
      setMDay(parseInt(parts[2], 10));
    }
  }, [isOpen, value]);

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

  const shortMonthNames = language === 'English' 
    ? ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    : monthNames;

  const weekDays = language === 'English'
    ? ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa']
    : ['អា', 'ច', 'អ', 'ព', 'ព្រ', 'សុ', 'ស'];

  // Mobile Wheel Setup
  const monthsForWheel = shortMonthNames.map((m, i) => ({ label: m, value: i + 1 }));
  const yearsForWheel = Array.from({ length: 40 }, (_, i) => new Date().getFullYear() - 20 + i).map(y => ({ label: y.toString(), value: y }));
  const daysInSelMonth = daysInMonth(mYear, mMonth - 1);
  const daysForWheel = Array.from({ length: daysInSelMonth }, (_, i) => i + 1).map(d => ({ label: d.toString().padStart(2, '0'), value: d }));

  const handleMobileConfirm = () => {
      const d = mDay > daysInSelMonth ? daysInSelMonth : mDay;
      const dStr = `${mYear}-${String(mMonth).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      onChange(dStr);
      setIsOpen(false);
  };

  return (
    <>
      <style>{`
        .no-scrollbar::-webkit-scrollbar { display: none; }
      `}</style>

      <div 
        className="date-picker-container relative group w-full h-[54px] cursor-pointer outline-none bg-transparent"
        onClick={() => setIsOpen(true)}
      >
        <div className={`absolute inset-0 flex items-center justify-between bg-[#FBFDFB] border-2 border-[#E8EEE9] hover:border-emerald-400 px-4 rounded-[20px] text-sm outline-none transition-all font-bold text-emerald-950 active:scale-[0.98] overflow-hidden shadow-sm focus-within:ring-4 focus-within:ring-emerald-500/10 focus-within:border-emerald-500 ${isOpen ? 'ring-4 ring-emerald-500/10 border-emerald-500' : ''}`}>
          <div className="flex items-center gap-3 overflow-hidden">
            <Calendar className="w-5 h-5 text-slate-500 shrink-0" />
            <span className="truncate">{getDisplayDate(value)}</span>
          </div>
          <ChevronDown className="w-4 h-4 text-slate-400 shrink-0 transition-transform duration-300" style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)' }} />
        </div>
      </div>
      
      <AnimatePresence>
        {isOpen && (
            <div className={`fixed inset-0 z-[200] flex ${isMobile ? 'items-end' : 'items-center justify-center p-4'} font-sans`}>
              <motion.div 
                 initial={{ opacity: 0 }}
                 animate={{ opacity: 1 }}
                 exit={{ opacity: 0 }}
                 transition={{ duration: 0.2 }}
                 className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm"
                 onClick={() => setIsOpen(false)}
              />
              <motion.div 
                 initial={{ y: isMobile ? '100%' : 10, scale: isMobile ? 1 : 0.95, opacity: 0 }}
                 animate={{ y: 0, scale: 1, opacity: 1 }}
                 exit={{ y: isMobile ? '100%' : 10, scale: isMobile ? 1 : 0.95, opacity: 0 }}
                 transition={{ type: "spring", damping: 25, stiffness: 300 }}
                 className={`relative w-full overflow-hidden shadow-[0_20px_60px_rgba(0,0,0,0.4)] ${isMobile ? 'bg-white rounded-t-[32px] rounded-b-none pb-safe' : 'bg-[#3F3F3F] max-w-[340px] rounded-[28px]'}`}
              >
                  {isMobile ? (
                    <div className="p-6 flex flex-col items-center w-full">
                       <div className="w-12 h-1.5 bg-slate-200 rounded-full mb-6" />
                       
                       <div className="flex w-full gap-2 justify-between mb-8 px-4">
                          <ScrollWheel items={monthsForWheel} selectedValue={mMonth} onChange={setMMonth} label={language === 'English' ? 'Month' : 'ខែ'} />
                          <ScrollWheel items={daysForWheel} selectedValue={mDay} onChange={setMDay} label={language === 'English' ? 'Day' : 'ថ្ងៃ'} />
                          <ScrollWheel items={yearsForWheel} selectedValue={mYear} onChange={setMYear} label={language === 'English' ? 'Year' : 'ឆ្នាំ'} />
                       </div>

                       <div className="w-full grid grid-cols-2 gap-4">
                          <button onClick={() => setIsOpen(false)} className="py-4 rounded-2xl font-bold bg-slate-100 text-slate-600 active:scale-95 transition-all text-sm">
                             {language === 'English' ? 'Cancel' : 'បោះបង់'}
                          </button>
                          <button onClick={handleMobileConfirm} className="py-4 rounded-2xl font-bold bg-emerald-500 text-white active:scale-95 transition-all shadow-md shadow-emerald-500/20 text-sm">
                             {language === 'English' ? 'Select Date' : 'ជ្រើសរើស'}
                          </button>
                       </div>
                    </div>
                  ) : (
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
                  )}
              </motion.div>
            </div>
        )}
      </AnimatePresence>
    </>
  );
}
