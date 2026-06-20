import { User, FinanceRecord, Category, AppConfig } from './types';
import { v4 as uuidv4 } from 'uuid';

const STORAGE_KEYS = {
  USERS: 'EM_USERS',
  RECORDS: 'EM_RECORDS',
  CATEGORIES: 'EM_CATEGORIES',
  CONFIG: 'EM_CONFIG',
  SESSION: 'EM_SESSION',
};

const DEFAULT_USERS: User[] = [
  { id: '1', name: 'Admin', pin: '1234', role: 'Admin', isActive: true },
  { id: '2', name: 'Member', pin: '1111', role: 'Member', isActive: true },
];

const DEFAULT_CATEGORIES: Category[] = [
  { id: uuidv4(), type: 'ចំណូល', name: 'ប្រាក់ខែ' },
  { id: uuidv4(), type: 'ចំណូល', name: 'លក់ដូរ' },
  { id: uuidv4(), type: 'ចំណូល', name: 'អាជីវកម្ម' },
  { id: uuidv4(), type: 'ចំណូល', name: 'ជំនួយ' },
  { id: uuidv4(), type: 'ចំណូល', name: 'ផ្សេងៗ' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'អាហារ' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'សាំង' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'ផ្ទះ' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'ទឹកភ្លើង' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'សុខភាព' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'កូនៗ' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'សាលារៀន' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'បំណុល' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'ទូរស័ព្ទ/Internet' },
  { id: uuidv4(), type: 'ចំណាយ', name: 'ផ្សេងៗ' },
];

const DEFAULT_CONFIG: AppConfig = {
  currency: '$',
  themeColor: 'emerald',
  appName: 'EasyMoney',
  language: 'Khmer',
  themeMode: 'Light',
  defaultMonth: new Date().getMonth() + 1,
  defaultYear: new Date().getFullYear(),
  dateFormat: 'DD/MM/YYYY',
  numberFormat: 'Comma',
  accentColor: 'teal',
  backgroundStyle: 'Plain',
  buttonGradient: true,
  cardRadius: 'rounded-[32px]',
  fontFamily: 'Inter',
  fontSize: 'text-sm',
  enableAnimation: true,
  memberViewAll: true,
  memberEditOwn: true,
  memberDeleteOwn: true,
  memberExport: true,
  memberImport: true,
  adminSettingsLock: false,
  importType: 'CSV',
  importMappingPreset: 'Default',
  duplicateCheckMode: 'DateAmountDesc',
  importPreviewLimit: 50,
  autoAssignMissingMember: true,
  autoDetectType: true,
  exportIncludeHeader: true,
  exportFileNameFormat: 'EasyMoney_[Date]'
};

function getStorage<T>(key: string, defaultVal: T): T {
  try {
    const val = localStorage.getItem(key);
    return val ? JSON.parse(val) : defaultVal;
  } catch {
    return defaultVal;
  }
}

function setStorage<T>(key: string, value: T) {
  localStorage.setItem(key, JSON.stringify(value));
}

export function initStorage() {
  if (!localStorage.getItem(STORAGE_KEYS.USERS)) {
    setStorage(STORAGE_KEYS.USERS, DEFAULT_USERS);
  }
  if (!localStorage.getItem(STORAGE_KEYS.CATEGORIES)) {
    setStorage(STORAGE_KEYS.CATEGORIES, DEFAULT_CATEGORIES);
  }
  if (!localStorage.getItem(STORAGE_KEYS.CONFIG)) {
    setStorage(STORAGE_KEYS.CONFIG, DEFAULT_CONFIG);
  } else {
    try {
      const existing = localStorage.getItem(STORAGE_KEYS.CONFIG);
      if (existing) {
        const parsed = JSON.parse(existing);
        let modified = false;
        if (parsed.currency === '៛') {
          parsed.currency = '$';
          modified = true;
        }
        if (parsed.dateFormat === 'YYYY-MM-DD') {
          parsed.dateFormat = 'DD/MM/YYYY';
          modified = true;
        }
        if (modified) {
          setStorage(STORAGE_KEYS.CONFIG, parsed);
        }
      }
    } catch (e) {
      // ignore
    }
  }
  if (!localStorage.getItem(STORAGE_KEYS.RECORDS)) {
    setStorage(STORAGE_KEYS.RECORDS, []);
  }
}

export function getUsers(): User[] {
  return getStorage(STORAGE_KEYS.USERS, DEFAULT_USERS);
}

export function saveUsers(users: User[]) {
  setStorage(STORAGE_KEYS.USERS, users);
}

export function getRecords(): FinanceRecord[] {
  return getStorage(STORAGE_KEYS.RECORDS, []);
}

export function saveRecords(records: FinanceRecord[]) {
  setStorage(STORAGE_KEYS.RECORDS, records);
}

export function getCategories(): Category[] {
  return getStorage(STORAGE_KEYS.CATEGORIES, DEFAULT_CATEGORIES);
}

export function saveCategories(categories: Category[]) {
  setStorage(STORAGE_KEYS.CATEGORIES, categories);
}

export function getConfig(): AppConfig {
  return getStorage(STORAGE_KEYS.CONFIG, DEFAULT_CONFIG);
}

export function saveConfig(config: AppConfig) {
  setStorage(STORAGE_KEYS.CONFIG, config);
}

export function getSession(): User | null {
  return getStorage<User | null>(STORAGE_KEYS.SESSION, null);
}

export function saveSession(user: User | null) {
  setStorage(STORAGE_KEYS.SESSION, user);
}

export function clearData() {
  localStorage.removeItem(STORAGE_KEYS.RECORDS);
  // Keep users, config, categories, or reset entirely? 
  // Requirements say "Clear all data with confirmation" usually means records, or full reset. Let's reset everything except session.
  localStorage.removeItem(STORAGE_KEYS.CATEGORIES);
  localStorage.removeItem(STORAGE_KEYS.USERS);
  localStorage.removeItem(STORAGE_KEYS.CONFIG);
  localStorage.removeItem(STORAGE_KEYS.SESSION);
  
  // Clear Google Sheet Sync local data as well for proper master reset
  localStorage.removeItem('EM_SYNC_MODE');
  localStorage.removeItem('EM_SYNC_URL');
  localStorage.removeItem('EM_OFFLINE_QUEUE');
  localStorage.removeItem('EM_LAST_SYNC');
  
  initStorage();
}

export function getBackupJSON() {
  return JSON.stringify({
    users: getUsers(),
    records: getRecords(),
    categories: getCategories(),
    config: getConfig(),
    syncUrl: localStorage.getItem('EM_SYNC_URL') || '',
    syncMode: localStorage.getItem('EM_SYNC_MODE') || 'local'
  }, null, 2);
}

export function restoreFromJSON(jsonString: string): boolean {
  try {
    const data = JSON.parse(jsonString);
    if (data.users && data.records && data.categories) {
      setStorage(STORAGE_KEYS.USERS, data.users);
      setStorage(STORAGE_KEYS.RECORDS, data.records);
      setStorage(STORAGE_KEYS.CATEGORIES, data.categories);
      if (data.config) setStorage(STORAGE_KEYS.CONFIG, data.config);
      
      // Auto-restore and activate Google Sheet Hybrid Sync Auto mapping
      if (data.syncUrl) {
        localStorage.setItem('EM_SYNC_URL', String(data.syncUrl).trim());
        localStorage.setItem('EM_SYNC_MODE', 'hybrid');
        localStorage.setItem('EM_JUST_RESTORED', 'true');
      } else if (localStorage.getItem('EM_SYNC_URL')) {
        // Force establish auto sync on reload
        localStorage.setItem('EM_SYNC_MODE', 'hybrid');
        localStorage.setItem('EM_JUST_RESTORED', 'true');
      } else {
        localStorage.setItem('EM_SYNC_MODE', 'hybrid');
        localStorage.setItem('EM_JUST_RESTORED', 'true');
      }
      return true;
    }
  } catch (e) {
    // console.error(e);
  }
  return false;
}
