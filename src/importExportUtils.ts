import * as XLSX from 'xlsx-js-style';
import { FinanceRecord, User, AppConfig, Category, RecordType } from './types';
import { formatMoney, formatDate, exportToCSV } from './utils';
import { getBackupJSON, getRecords, getCategories } from './store';
import { v4 as uuidv4 } from 'uuid';

const KHMER_MONTHS = [
  'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'
];

export const getKhmerMonthYearSheetName = (month: number, year: number) => {
  const mName = KHMER_MONTHS[month - 1] || `ខែ${month}`;
  return `${mName} ${year}`;
};

export function parseExcelDate(val: any): { dateStr: string, month: number, year: number } {
  if (!val) {
    const today = new Date();
    return {
      dateStr: today.toISOString().split('T')[0],
      month: today.getMonth() + 1,
      year: today.getFullYear()
    };
  }
  
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = val.getMonth() + 1;
    const d = val.getDate();
    return {
      dateStr: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      month: m,
      year: y
    };
  }

  if (typeof val === 'number') {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    const y = date.getFullYear();
    const m = date.getMonth() + 1;
    const d = date.getDate();
    return {
      dateStr: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      month: m,
      year: y
    };
  }

  const str = String(val).trim();
  
  const dmRef = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (dmRef) {
    const d = parseInt(dmRef[1], 10);
    const m = parseInt(dmRef[2], 10);
    const y = parseInt(dmRef[3], 10);
    return {
      dateStr: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      month: m,
      year: y
    };
  }
  
  const ymRef = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (ymRef) {
    const y = parseInt(ymRef[1], 10);
    const m = parseInt(ymRef[2], 10);
    const d = parseInt(ymRef[3], 10);
    return {
      dateStr: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
      month: m,
      year: y
    };
  }

  try {
    const dObj = new Date(str);
    if (!isNaN(dObj.getTime())) {
      const y = dObj.getFullYear();
      const m = dObj.getMonth() + 1;
      const d = dObj.getDate();
      return {
        dateStr: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`,
        month: m,
        year: y
      };
    }
  } catch (e) {
    // ignore
  }

  const today = new Date();
  return {
    dateStr: today.toISOString().split('T')[0],
    month: today.getMonth() + 1,
    year: today.getFullYear()
  };
}

export const buildEasyMoneyDataRows = (records: FinanceRecord[], users: User[], config: AppConfig) => {
  return records.map((r) => ({
    'ID': r.id,
    'Date': r.date,
    'Month': r.month,
    'Year': r.year,
    'Type': r.type,
    'Category': r.category,
    'Description': r.description || '',
    'Amount': r.amount,
    'Currency': r.currency,
    'Exchange Rate': r.exchangeRate || (config.exchangeRate || 4000),
    'Amount KHR': r.amountKHR,
    'Amount USD': r.amountUSD,
    'Display Amount': r.displayAmount,
    'Payment Method': r.paymentMethod || '',
    'Family Member': users.find(u => u.id === r.memberId)?.name || 'Unknown',
    'Member ID': r.memberId,
    'Note': r.note || '',
    'Created At': r.createdAt ? new Date(r.createdAt).toISOString() : new Date().toISOString(),
    'Updated At': r.updatedAt ? new Date(r.updatedAt).toISOString() : new Date().toISOString()
  }));
};

export const buildEasyMoneyViewRows = (records: FinanceRecord[], users: User[], config: AppConfig) => {
  return records.map((r) => {
    const d = new Date(r.date);
    const dateFormatted = !isNaN(d.getTime()) ? `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}` : r.date;
    return {
      'ថ្ងៃខែ': dateFormatted,
      'ប្រភេទ': r.type,
      'ប្រភេទចំណូល/ចំណាយ': r.category,
      'ពិពណ៌នា': r.description || '',
      'ចំនួនលុយ': r.amount,
      'រូបិយប័ណ្ណ': r.currency === '$' ? 'USD' : 'KHR',
      'វិធីបង់ប្រាក់': r.paymentMethod || '',
      'សមាជិក': users.find(u => u.id === r.memberId)?.name || 'Unknown',
      'កំណត់ចំណាំ': r.note || '',
      'កូដសម្គាល់ (ID)': r.id
    };
  });
};

function styleViewSheet(ws: any) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  
  ws['!autofilter'] = { ref: ws['!ref'] as string };
  ws['!views'] = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_cell({ r: R, c: C });
      if (!ws[address]) continue;

      if (R === 0) {
        ws[address].v = String(ws[address].v);
        ws[address].s = {
          fill: { fgColor: { rgb: "D9EAD3" } }, 
          font: { bold: true, color: { rgb: "000000" }, name: "Kantumruy Pro", size: 10 },
          alignment: { vertical: "center", horizontal: "center", wrapText: true },
          border: {
            bottom: { style: "medium", color: { rgb: "B8C7B5" } }
          }
        };
      } else {
        const typeAddress = XLSX.utils.encode_cell({ r: R, c: 1 }); 
        const isInc = ws[typeAddress]?.v === 'ចំណូល';
        
        ws[address].s = {
          alignment: { vertical: "center", horizontal: "left" },
          font: { name: "Kantumruy Pro", size: 10 },
          border: {
            bottom: { style: "thin", color: { rgb: "E2E8F0" } }
          }
        };

        if (C === 1) { // 'ប្រភេទ'
          ws[address].s.fill = { fgColor: { rgb: isInc ? "E2EFDA" : "FCE4D6" } };
          ws[address].s.font.bold = true;
          ws[address].s.font.color = { rgb: isInc ? "375623" : "C65911" };
          ws[address].s.alignment.horizontal = "center";
        }

        if (C === 4) { // 'ចំនួនលុយ'
          const rCurrency = ws[XLSX.utils.encode_cell({ r: R, c: 5 })]?.v; 
          ws[address].t = 'n';
          if (rCurrency === 'USD') {
            ws[address].z = '"$"#,##0.00';
          } else {
            ws[address].z = '#,##0" ៛"';
          }
          ws[address].s.font.bold = true;
          ws[address].s.alignment.horizontal = "right";
          ws[address].s.font.color = { rgb: isInc ? "1B5E20" : "B71C1C" }; 
        }

        if (C === 0 || C === 5) { // 'ថ្ងៃខែ', 'រូបិយប័ណ្ណ'
          ws[address].s.alignment.horizontal = "center";
        }
        
        if (C === 9) { // 'កូដសម្គាល់ (ID)'
          ws[address].s.font.color = { rgb: "64748B" }; 
          ws[address].s.font.size = 8.5;
          ws[address].s.alignment.horizontal = "center";
        }
      }
    }
  }

  ws['!cols'] = [
    { wch: 15 }, // ថ្ងៃខែ
    { wch: 10 }, // ប្រភេទ
    { wch: 22 }, // ប្រភេទចំណូល/ចំណាយ
    { wch: 30 }, // ពិពណ៌នា
    { wch: 16 }, // ចំនួនលុយ
    { wch: 12 }, // រូបិយប័ណ្ណ
    { wch: 15 }, // វិធីបង់ប្រាក់
    { wch: 15 }, // សមាជិក
    { wch: 25 }, // កំណត់ចំណាំ
    { wch: 28 }  // កូដសម្គាល់ (ID)
  ];
}

export const exportEasyMoneyExcel = (records: FinanceRecord[], users: User[], config: AppConfig, filename: string) => {
  const wb = XLSX.utils.book_new();

  // Determine the sheet name dynamically based on the filtered records
  let sheetName = 'គ្រប់កំណត់ត្រា'; // Default: All records
  if (records.length > 0) {
    const firstRec = records[0];
    const allSameMonthAndYear = records.every(r => r.month === firstRec.month && r.year === firstRec.year);
    const allSameYear = records.every(r => r.year === firstRec.year);

    if (allSameMonthAndYear) {
      sheetName = getKhmerMonthYearSheetName(firstRec.month, firstRec.year);
    } else if (allSameYear) {
      sheetName = `ឆ្នាំ ${firstRec.year}`;
    }
  }

  // Create one single sheet with the viewed rows
  const viewRows = buildEasyMoneyViewRows(records, users, config);
  const ws = XLSX.utils.json_to_sheet(viewRows);
  styleViewSheet(ws);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  wb.Props = {
    Title: "EasyMoney Monthly Export",
    CreatedDate: new Date()
  };

  XLSX.writeFile(wb, filename);
};

export const exportEasyMoneyCSV = (records: FinanceRecord[], users: User[], config: AppConfig, filename: string) => {
  const dataRows = buildEasyMoneyDataRows(records, users, config);
  exportToCSV(dataRows, filename);
};

export const parseEasyMoneyExchangeFormat = (
  rows: any[], 
  users: User[], 
  config: AppConfig, 
  sessionUser: User | null
): { records: FinanceRecord[], duplicates: number } => {
  // Backwards compatibility function
  const helperWb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(helperWb, ws, 'EasyMoney_Data');
  return parseEasyMoneyExcelWorkbook(helperWb, users, config, sessionUser);
};

export const parseEasyMoneyExcelWorkbook = (
  workbook: any,
  users: User[],
  config: AppConfig,
  sessionUser: User | null
): { records: FinanceRecord[], duplicates: number } => {
  const recordsInDb = getRecords();
  const categoriesInDb = getCategories();
  const cleanRecords: FinanceRecord[] = [];
  let duplicates = 0;

  const newCategories: Category[] = [];
  const processedKeys = new Set<string>();

  workbook.SheetNames.forEach((sheetName: string) => {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws['!ref']) return;

    const rawRows: any[] = XLSX.utils.sheet_to_json(ws, { defval: '' });
    if (rawRows.length === 0) return;

    const firstRow = rawRows[0];
    const isViewFormat = ('ប្រភេទ' in firstRow && 'ចំនួនលុយ' in firstRow);
    const isDataFormat = ('Type' in firstRow && 'Amount' in firstRow);

    if (!isViewFormat && !isDataFormat) {
      return; 
    }

    rawRows.forEach((row) => {
      let id = '';
      let dateRaw: any = '';
      let type: RecordType = 'ចំណាយ';
      let categoryName = 'ផ្សេងៗ';
      let description = '';
      let amountNum = 0;
      let currency: '៛' | '$' = '$';
      let paymentMethod = 'លុយសុទ្ធ';
      let memberName = 'Admin';
      let note = '';

      if (isDataFormat) {
        id = String(row['ID'] || '');
        dateRaw = row['Date'];
        
        const rawType = String(row['Type'] || '');
        type = (rawType === 'ចំណូល' || rawType === 'Income') ? 'ចំណូល' : 'ចំណាយ';
        categoryName = String(row['Category'] || 'ផ្សេងៗ');
        description = String(row['Description'] || '');
        amountNum = Number(row['Amount']) || 0;
        
        const rawCurr = String(row['Currency'] || '');
        currency = (rawCurr === 'KHR' || rawCurr === '៛') ? '៛' : '$';
        paymentMethod = String(row['Payment Method'] || 'លុយសុទ្ធ');
        memberName = String(row['Family Member'] || 'Admin');
        note = String(row['Note'] || '');
      } else {
        id = String(row['កូដសម្គាល់ (ID)'] || row['ID'] || '');
        dateRaw = row['ថ្ងៃខែ'];
        
        const rawType = String(row['ប្រភេទ'] || '');
        type = (rawType === 'ចំណូល' || rawType === 'Income') ? 'ចំណូល' : 'ចំណាយ';
        categoryName = String(row['ប្រភេទចំណូល/ចំណាយ'] || 'ផ្សេងៗ');
        description = String(row['ពិពណ៌នា'] || '');

        let rawAmount = row['ចំនួនលុយ'];
        if (typeof rawAmount === 'number') {
          amountNum = rawAmount;
        } else if (rawAmount) {
          const cleaned = String(rawAmount).replace(/[^0-9\.]/g, '');
          amountNum = parseFloat(cleaned) || 0;
        }
        
        const rawCurr = String(row['រូបិយប័ណ្ណ'] || '');
        currency = (rawCurr === 'KHR' || rawCurr === '៛' || rawCurr === 'រៀល') ? '៛' : '$';
        paymentMethod = String(row['វិធីបង់ប្រាក់'] || 'លុយសុទ្ធ');
        memberName = String(row['សមាជិក'] || 'Admin');
        note = String(row['កំណត់ចំណាំ'] || '');
      }

      if (!dateRaw && amountNum === 0) return;

      const { dateStr, month, year } = parseExcelDate(dateRaw);

      const uniqueFileKey = id ? `id_${id}` : `sig_${dateStr}_${type}_${categoryName}_${amountNum}_${description}`;
      if (processedKeys.has(uniqueFileKey)) {
        return;
      }
      processedKeys.add(uniqueFileKey);

      const isDuplicate = recordsInDb.some((dr: FinanceRecord) => {
        if (id && dr.id === id) return true;
        if (
          dr.date === dateStr &&
          dr.amount === amountNum &&
          dr.type === type &&
          dr.category === categoryName &&
          dr.description === description
        ) return true;
        return false;
      });

      if (isDuplicate) {
        duplicates++;
        return;
      }

      if (!categoriesInDb.some(c => c.name === categoryName && c.type === type) && !newCategories.some(c => c.name === categoryName && c.type === type)) {
        newCategories.push({
          id: uuidv4(),
          type: type,
          name: categoryName
        });
      }

      const matchName = memberName.toLowerCase().trim();
      const userMatch = users.find(u => u.name.toLowerCase().trim() === matchName);
      const memberId = userMatch ? userMatch.id : (sessionUser?.id || users[0]?.id || 'unknown');

      const exRate = config.exchangeRate || 4000;
      const amountKHR = currency === '៛' ? amountNum : amountNum * exRate;
      const amountUSD = currency === '$' ? amountNum : amountNum / exRate;

      cleanRecords.push({
        id: id || uuidv4(),
        date: dateStr,
        month: month,
        year: year,
        type: type,
        category: categoryName,
        description: description,
        currency: currency,
        amount: amountNum,
        amountKHR: amountKHR,
        amountUSD: amountUSD,
        exchangeRate: exRate,
        displayAmount: currency === '$' ? `$${amountNum}` : `${amountNum} ៛`,
        paymentMethod: paymentMethod,
        memberId: memberId,
        note: note,
        createdAt: Date.now(),
        updatedAt: Date.now()
      });
    });
  });

  if (newCategories.length > 0) {
    import('./store').then(m => {
      m.saveCategories([...categoriesInDb, ...newCategories]);
    });
  }

  return { records: cleanRecords, duplicates };
};
