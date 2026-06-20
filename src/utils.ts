import { AppConfig } from './types';
import { getConfig } from './store';

export function formatMoney(amount: number, currency: string = '៛'): string {
  const config = getConfig();
  const options: Intl.NumberFormatOptions = { 
    minimumFractionDigits: currency === '$' ? 2 : 0, 
    maximumFractionDigits: currency === '$' ? 2 : 0,
    useGrouping: config.numberFormat !== 'Plain'
  };
  
  if (currency === '$') {
    return '$' + amount.toLocaleString('en-US', options);
  } else {
    return amount.toLocaleString('en-US', options) + ' ៛';
  }
}

export function formatDate(dateString: string): string {
  if (!dateString) return '';
  
  // Clean dateString first: take only the first 10 characters (YYYY-MM-DD) if it contains "T" or space
  let cleanDate = dateString;
  if (typeof dateString === 'string') {
    if (cleanDate.includes('T')) {
      cleanDate = cleanDate.split('T')[0];
    } else if (cleanDate.includes(' ')) {
      cleanDate = cleanDate.split(' ')[0];
    }
  }

  const config = getConfig();
  const format = config.dateFormat || 'DD/MM/YYYY';
  if (format === 'YYYY-MM-DD') return cleanDate;
  
  const parts = cleanDate.split('-');
  if (parts.length !== 3) return dateString;
  const [y, m, d] = parts;
  if (format === 'DD/MM/YYYY') return `${d}/${m}/${y}`;
  if (format === 'MM/DD/YYYY') return `${m}/${d}/${y}`;
  return cleanDate;
}

export function generateYearOptions(records: any[] = [], customYears: number[] = []) {
  const currentYear = new Date().getFullYear();
  const years = new Set<number>();
  
  for (let i = currentYear - 10; i <= currentYear + 10; i++) {
    years.add(i);
  }
  
  if (customYears && customYears.length > 0) {
    customYears.forEach(y => years.add(y));
  }
  
  if (records && records.length > 0) {
    records.forEach(r => {
      if (r.date) {
        const y = parseInt(r.date.split('-')[0]);
        if (!isNaN(y)) years.add(y);
      }
    });
  }
  
  const sortedYears = Array.from(years).sort((a, b) => a - b);
  return sortedYears.map(y => ({ value: y, label: String(y) }));
}

export function exportToCSV(data: any[], filename: string) {
  if (!data || data.length === 0) return;
  
  const header = Object.keys(data[0]);
  const csvRows = [header.join(',')];
  for (const row of data) {
    const values = header.map(k => {
      const val = row[k] === null || row[k] === undefined ? '' : String(row[k]);
      return `"${val.replace(/"/g, '""')}"`;
    });
    csvRows.push(values.join(','));
  }
  const csv = csvRows.join('\n');

  const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csv], { type: 'text/csv;charset=utf-8;' }); // BOM for Excel
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", filename);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
