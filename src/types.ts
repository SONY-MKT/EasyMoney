export type RecordType = 'ចំណូល' | 'ចំណាយ';

export interface User {
  id: string;
  name: string;
  pin: string;
  role: 'Admin' | 'Member';
  isActive: boolean;
}

export interface FinanceRecord {
  id: string;
  date: string;
  month: number;
  year: number;
  type: RecordType;
  category: string;
  description: string;
  currency: '៛' | '$';
  amount: number; // original input amount
  amountKHR: number;
  amountUSD: number;
  exchangeRate: number;
  displayAmount: string;
  paymentMethod: string;
  memberId: string;
  note: string;
  createdAt: number;
  updatedAt: number;
}

export interface Category {
  id: string;
  type: RecordType;
  name: string;
  nameEn?: string;
}

export interface AppConfig {
  currency: '៛' | '$';
  exchangeRate?: number;
  themeColor: string;
  customYears?: number[];
  
  // General Settings
  appName?: string;
  appLogoUrl?: string;
  defaultMonth?: number;
  defaultYear?: number;
  dateFormat?: string;
  numberFormat?: string;
  language?: 'Khmer' | 'English';
  themeMode?: 'Light' | 'Dark' | 'System';

  // Theme Settings
  accentColor?: string;
  backgroundStyle?: string;
  buttonGradient?: boolean;
  cardRadius?: string; // 'none' | 'sm' | 'md' | 'lg' | 'full'
  fontFamily?: string;
  fontSize?: string;
  enableAnimation?: boolean;
  compactMode?: boolean;
  showShadows?: boolean;

  // User Permission Settings
  memberViewAll?: boolean;
  memberEditOwn?: boolean;
  memberDeleteOwn?: boolean;
  memberExport?: boolean;
  memberImport?: boolean;
  adminSettingsLock?: boolean;

  // Import Settings
  importType?: 'CSV' | 'XLSX';
  importMappingPreset?: string;
  duplicateCheckMode?: 'ID' | 'DateAmountDesc';
  importPreviewLimit?: number;
  autoAssignMissingMember?: boolean;
  autoDetectType?: boolean;

  // Export Settings
  exportIncludeHeader?: boolean;
  exportFileNameFormat?: string;
}
