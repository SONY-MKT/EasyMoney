import React, { useEffect, useState, useMemo, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx-js-style';
import { exportEasyMoneyExcel, exportEasyMoneyCSV, parseEasyMoneyExchangeFormat, parseEasyMoneyExcelWorkbook } from './importExportUtils';
import { v4 as uuidv4 } from 'uuid';
import { CustomSelect } from './CustomSelect';
import CustomDatePicker from './CustomDatePicker';
import { 
  PieChart, List, Plus, Settings, LogOut, ArrowUpRight, ArrowDownLeft,
  ChevronRight, ChevronDown, ChevronUp, Download, Upload, Users, Tag, Trash2, AlertCircle,
  Cloud, Check, RefreshCw, Wifi, WifiOff, Database, Palette, Lock, Shield,
  Sliders, FileText, CheckCircle2, Search, Edit2, Key, ToggleLeft, ToggleRight, User as UserIcon, Info, Loader2, PlusCircle, X,
  GripVertical, TrendingUp, TrendingDown, Calendar
} from 'lucide-react';
import { 
  initStorage, getSession, saveSession, getUsers, getRecords, saveRecords, 
  getCategories, saveCategories, getConfig, clearData, getBackupJSON, restoreFromJSON,
  saveUsers, saveConfig
} from './store';
import { User, FinanceRecord, Category, AppConfig, RecordType } from './types';
import { formatMoney, formatDate, exportToCSV, generateYearOptions } from './utils';
import { cn } from './lib/utils';
import { PieChart as RePieChart, Pie, Cell, ResponsiveContainer, Tooltip as ReTooltip, LineChart, Line, XAxis, YAxis, CartesianGrid } from 'recharts';
import { 
  getSyncMode, saveSyncMode, getSyncUrl, saveSyncUrl, getOfflineQueue, addToOfflineQueue,
  testGoogleSheetConnection, syncRecordToGoogleSheet, loadRecordsFromGoogleSheet, 
  updateRecordInGoogleSheet, deleteRecordFromGoogleSheet, syncLocalDataToGoogleSheet,
  getLastSyncTime, SyncMode, syncAppConfigToGoogleSheet, syncUsersToGoogleSheet, syncCategoriesToGoogleSheet,
  loadSettingsFromGoogleSheet, loadUsersFromGoogleSheet, loadCategoriesFromGoogleSheet
} from './googleSheetSync';
import { Toaster, toast } from 'sonner';
import { motion, AnimatePresence } from 'motion/react';
import CountUp from 'react-countup';
import { DragDropContext, Droppable, Draggable, DropResult } from '@hello-pangea/dnd';

const KHMER_MONTHS = ['មករា', 'កុម្ភៈ', 'មិនា', 'មេសា', 'ឧសភា', 'មិថុនា', 'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'];

const TRANSLATIONS: Record<string, Record<string, string>> = {
  English: {
    // Bottom Nav & Tabs
    "សរុប": "Dashboard",
    "បញ្ជី": "List",
    "ទិន្នន័យ": "Database",
    "កំណត់": "Settings",

    // General Words
    "រក្សាទុក": "Save",
    "បោះបង់": "Cancel",
    "យល់ព្រម": "OK",
    "លុប": "Delete",
    "កែប្រែ": "Edit",
    "បន្ថែម": "Add",
    "ចាកចេញ": "Sign Out",
    "សូមបញ្ជាក់ (Confirm)": "Confirm Action",
    "សូមបញ្ជាក់": "Confirm Action",
    "យល់ព្រមលុប": "Yes, Delete",
    "បិទ": "Close",

    // Month Names
    "មករា": "January",
    "កុម្ភៈ": "February",
    "មិនា": "March",
    "មេសា": "April",
    "ឧសភា": "May",
    "មិថុនា": "June",
    "កក្កដា": "July",
    "សីហា": "August",
    "កញ្ញា": "September",
    "តុលា": "October",
    "វិច្ឆិកា": "November",
    "ធ្នូ": "December",

    // App Header & Welcome
    "សួស្តី": "Hello",

    // Bottom navigation
    "សមាជិក": "Members",
    "កំណត់ត្រា": "Records",

    // Dashboard View strings
    "សមតុល្យសរុប (ខែនេះ)": "Total Balance (This Month)",
    "សមតុល្យសរុប (ឆ្នាំនេះ)": "Total Balance (This Year)",
    "សមតុល្យខែនេះ": "Balance This Month",
    "ចំណូល": "Income",
    "ចំណាយ": "Expense",
    "ចំណូលសរុប": "Total Income",
    "ចំណាយសរុប": "Total Expenses",
    "របាយការណ៍សរុប": "Overview Report",
    "ស្ថិតិកំណត់ត្រា": "Record Statistics",
    "សរុបកំណត់ត្រា": "Total Records",
    "ស្ថានភាពសង្ខេប": "Summary Info",
    "ការកំណត់ និងសុវត្ថិភាព": "Settings & Security",
    "សកម្ម": "Active",
    "សកម្មភាព": "Activity",
    "ស្ងប់ស្ងាត់": "Inactive",
    "មិនទាន់មាន": "None",
    "ស្ថិតិប្រព័ន្ធទិន្នន័យ (DATABASE STATISTICS)": "Database Statistics",
    "ព័ត៌មានជំនួយ និងសុវត្ថិភាព (HELP & SAFETY TIPS)": "Help & Safety Tips",
    "រក្សាទុកបាន ១០០% ស្របច្បាប់ឯកជនភាព": "100% Private & Secure",
    "ទិន្នន័យចំណូល-ចំណាយទាំងអស់ត្រូវបានរក្សាទុកក្នុុង Local Storage នៃឧបករណ៍របស់អ្នក។ គ្មាននរណាម្នាក់អាចមើលវាបានឡើយ។": "All your income and expense data are stored in your device's local storage. Nobody else can access or view it.",

    // List View
    "តារាងចំណាយខែនេះ": "Expenses Chart (This Month)",
    "តារាងចំណាយឆ្នាំនេះ": "Expenses Chart (This Year)",
    "និន្នាការចំណាយ ៦ ខែចុងក្រោយ": "Expenses Trend (Last 6 Months)",
    "ព័ត៌មានលម្អិត": "Detail Information",
    "ចំនួន": "Amount",
    "សមតុល្យសរុប": "Total Balance",
    "បញ្ជីប្រតិបត្តិការ": "Transactions List",
    "ស្វែងរក...": "Search...",
    "ចំណាយសរុបក្នុងបញ្ជី": "Total Expenses in List",
    "ចំណូលសរុបក្នុងបញ្ជី": "Total Income in List",
    "សមតុល្យលទ្ធផល": "Result Balance",
    "ប្រភេទចំណាយ": "Expense Categories",
    "ប្រភេទចំណូល": "Income Categories",
    "សមាជិកទាំងអស់": "All Members",
    "ប្រភេទទាំងអស់": "All Categories",
    "ខែទាំងអស់": "All Months",
    "ឆ្នាំទាំងអស់": "All Years",
    "គ្មានទិន្នន័យស្របនឹងការស្វែងរករបស់អ្នកឡើយ!": "No data matches your search!",

    // Quick add floating menu
    "បន្ថែមប្រតិបត្តិការ": "Add Transaction",
    "កាលបរិច្ឆេទ": "Date",
    "ចំនួនទឹកប្រាក់": "Amount",
    "ការពិពណ៌នា": "Description",
    "ប្រភព / សមាជិក": "Source / Member",
    "ប្រភេទចំណូល/ចំណាយ": "Category Type",
    "បញ្ចូលប្រតិបត្តិការថ្មី": "Insert New Transaction",
    "កែប្រែប្រតិបត្តិការ": "Edit Transaction/Record",
    "ប្រចាំខែ": "Monthly",
    "ប្រចាំឆ្នាំ": "Yearly",
    "ចំណូលពេញឆ្នាំ": "Yearly Income",
    "ចំណូលខែនេះ": "Monthly Income",
    "ចំណាយពេញឆ្នាំ": "Yearly Expenses",
    "ចំណាយខែនេះ": "Monthly Expenses",
    "ប្រាក់សល់ប្រចាំឆ្នាំ": "Yearly Savings",
    "ប្រាក់សល់ខែនេះ": "Monthly Savings",

    // File / Data Sync view
    "ស្ដារ ឬ នាំចូល": "Restore or Import",
    "ស្ដារ ឬនាំចូល": "Restore or Import",
    "នាំចេញទិន្នន័យ": "Export Data",
    "នាំចូលទិន្នន័យ": "Import Data",
    "ទិន្នន័យម៉ាស៊ីន": "LocalStorage",
    "ស្ថានភាពសមកាលកម្ម": "Sync State",
    "របៀបសមកាលកម្ម (Sync Mode)": "Sync Mode",
    "ការភ្ជាប់ Web App URL": "Google Sheets API URL",
    "ស្ថានភាពការងារ": "Activity State",
    "ទាញយកជា EXCEL ឬ CSV": "Download as EXCEL or CSV",
    "ទាញយកជា EXCEL": "Download as EXCEL",
    "ទាញយកជា CSV": "Download as CSV",
    "ទាញយកជា JSON": "Download as JSON",
    "ស្ដារទិន្នន័យឡើងវិញ": "Restore System Backup",
    "ស្ដារពីលទ្ធផល JSON": "Restore from JSON File",
    "នាំចូលពី Excel/CSV": "Import from Excel/CSV",
    "ទាញយកទិន្នន័យ JSON": "Download JSON Backup",
    "សមកាលកម្មដោយដៃ": "Manual Sync Now",
    "ការភ្ជាប់ល្អ": "Connected",
    "គ្មានអ៊ីនធឺណិត": "Offline",
    "ម៉ោង Sync ចុងក្រោយ": "Last Sync Time",
    "ប្រភេទឯកសារ": "File Type",
    "គំរូទិន្នន័យ": "Data preview rows",

    // Category / Member View
    "សមាជិកគ្រួសារ (FAMILY ACCOUNTS)": "Family Members",
    "សមាជិកគ្រួសារ": "Family Members",
    "ប្រភេទប្រតិបត្តិការ (CATEGORY LIST)": "Categories",
    "ប្រភេទប្រតិបត្តិការ": "Categories",
    "បន្ថែមសមាជិកគ្រួសារ": "Add Family Member",
    "បន្ថែមប្រភេទចំណាយ": "Add Expense Category",
    "បន្ថែមប្រភេទចំណូល": "Add Income Category",
    "បញ្ជីចំណូល": "Income List",
    "បញ្ជីចំណាយ": "Expense List",
    "បន្ថែមប្រភេទថ្មី...": "Add new category...",
    "បន្ថែមគណនីថ្មី (Create Member account)": "Create Member Account",
    "បន្ថែមគណនីថ្មី": "Create Member Account",
    "ឈ្មោះសមាជិក": "Member Name",
    "សិទ្ធិប្រព័ន្ធ (Role)": "System Role",
    "សិទ្ធិប្រព័ន្ធ": "System Role",
    "សមាជិកធម្មតា": "Regular Member",
    "បន្ថែមគណនី": "Create Account",

    // General settings titles
    "ការកំណត់ទូទៅ (GENERAL SETTINGS)": "General Settings",
    "ការកំណត់ផ្ទៃកម្មវិធី (THEME SETTINGS)": "Theme Settings",
    "ភាសាកម្មវិធី": "App Language",
    "រូបសញ្ញា Logo (URL)": "App Logo URL",
    "ឈ្មោះកម្មវិធី": "App Name",
    "រូបប័ណ្ណលុយចម្បង (បង្ហាញ)": "Primary Currency (Display)",
    "អត្រាប្តូរប្រាក់ (១ ដុល្លារ = ?)": "Exchange Rate (1 USD = ?)",
    "អត្រាប្តូរប្រាក់": "Exchange Rate",
    "ឆ្នាំលំនាំដើម (Default Year)": "Default Year",
    "ឆ្នាំលំនាំដើម": "Default Year",
    "ខែលំនាំដើម (Default Month)": "Default Month",
    "ខែលំនាំដើម": "Default Month",
    "ទម្រង់កាលបរិច្ឆេទ": "Date Format",
    "ទម្រង់លេខសំគាល់": "Number Format",
    "ម៉ូដពណ៌ផ្ទៃ": "Theme Mode",
    "ទំហំរៀបចំ": "Layout Size",
    "ក្ដារពណ៌ចម្បង": "Primary Palette",
    "បង្ហាញស្រមោល": "Show Shadows",
    "បន្ថែមជម្រៅលើប៊ូតុង និងកាត": "Add depth to cards and buttons",
    "បើកចលនា UI (Animations)": "Enable UI Animations",
    "បង្ហាញចលនារលូនពេលប្ដូរផ្ចាំ": "Smooth transitions between views",

    // Help safety
    "គណនីរបស់ខ្ញុំ": "My Account",
    "ចាកចេញពីគណនី": "Logout from account",
    "រក្សាទុកក្នុងម៉ាស៊ីន ប៉ុន្តែសមកាលកម្មទៅ Google Sheet បរាជ័យ!": "Saved locally, but syncing to Google Sheet failed!",
    "ប្ដូររបៀបសមកាលកម្មជោគជ័យ!": "Sync mode changed successfully!",

    // Alerts, toasts, and confirmations translations
    "សូមបញ្ចូល API URL ជាមុនសិន!": "Please enter API URL first!",
    "ភ្ជាប់ Google Sheet បានជោគជ័យ": "Google Sheet connected successfully!",
    "ការតភ្ជាប់ទៅកាន់ Google Sheet បរាជ័យ!": "Connection to Google Sheet failed!",
    "កំពុងសមកាលកម្មទិន្នន័យ...": "Syncing data...",
    "សមកាលកម្មទិន្នន័យ កំណត់ត្រា និងការកំណត់បានជោគជ័យពេញលេញ!": "Data sync and configurations completed successfully!",
    "សមកាលកម្មកំណត់ត្រាជោគជ័យ ប៉ុន្តែការកំណត់មួយចំនួនបរាជ័យ": "Record sync successful, but some configurations failed",
    "សមកាលកម្មមានបញ្ហាមួយចំនួន": "Sync encountered some issues",
    "ខកខានក្នុងការ Sync៖ ": "Failed to sync: ",
    "នាំចេញ JSON Backup រួចរាល់": "JSON backup exported successfully!",
    "នាំចេញ JSON បរាជ័យ": "JSON export failed",
    "ស្តារទិន្នន័យពី JSON ជោគជ័យ កម្មវិធីកំពុងចាប់ផ្ដើមឡើងវិញ...": "JSON data restored successfully! Restarting application...",
    "បានកំណត់ Settings ទៅដើមវិញជោគជ័យ": "Settings reset to default successfully!",
    "កំពុងលុបពី Google Sheet...": "Deleting from Google Sheet...",
    "បានលុបកំណត់ត្រាទាំងអស់ (រួមទាំងក្នុង Google Sheet) រួចរាល់": "All records deleted (including Google Sheet) successfully!",
    "បានលុបក្នុងម៉ាស៊ីន ប៉ុន្តែបំពានសិទ្ធិឯកសារបញ្ជូនទៅ Google Sheet!": "Deleted locally, but sync deletion to Google Sheet failed due to permission!",
    "បានលុបកំណត់ត្រាទាំងអស់រួចរាល់": "All records deleted successfully!",
    "តើអ្នកពិតជាចង់កំណត់ការកំណត់ (Settings) ទាំងអស់ទៅដើមវិញមែនទេ?": "Are you sure you want to reset all settings to default?",
    "⚠️ គ្រោះថ្នាក់៖ តើអ្នកពិតជាចង់លុបប្រតិបត្តិការកត់ត្រាទាំងអស់មែនទេ?": "⚠️ Warning: Are you sure you want to delete all transaction records?",
    "🚨 គ្រោះថ្នាក់បំផុត៖ តើអ្នកពិតជានឹងលុបទិន្នន័យកម្មវិធីទាំងអស់មែនទេ? កម្មវិធីនឹងចាប់ផ្តើមជាថ្មីទាំងស្រុង!": "🚨 Extreme Danger: Are you sure you want to erase ALL database records? The application will start completely raw!",
    "គ្មានទិន្នន័យដើម្បីនាំចូលឡើយ!": "No data to import!",
    "បន្ថែមប្រភេទលុយបានជោគជ័យ": "Category added successfully!",
    "តើអ្នកចង់លុបប្រភេទនេះមែនទេ?": "Are you sure you want to delete this category?",
    "រក្សាទុកការកំណត់បានជោគជ័យ": "Settings saved successfully!",
    "ភ្ជាប់ជោគជ័យ! កំពុងពិនិត្យទិន្នន័យចាស់...": "Connection successful! Checking historical records...",
    "មិនមានទិន្នន័យថ្មីក្នុង Google Sheet ទេ": "No new records found in Google Sheet",
    "មិនមានទិន្នន័យក្នុង Google Sheet": "No records found in Google Sheet",
    "កំហុសក្នុងការទាញទិន្នន័យពី Google Sheet": "Failed to pull records from Google Sheet",
    "កំពុងចាប់ផ្តើមសមកាលកម្ម...": "Initiating synchronization...",
    "សមកាលកម្មទិន្នន័យបានជោគជ័យពេញលេញ!": "Data sync completed successfully!",
    "សមកាលកម្មបានជួបបញ្ហាមួយចំនួន": "Sync encountered some issues",
    "កំហុសនៃការ Sync": "Sync error",
    "តើអ្នកពិតជាចង់សម្អាត Offline Queue ដែលមិនទាន់ Sync មែនទេ?": "Are you sure you want to clear your unsynced offline queue?",
    "បានសម្អាត Queue រួចរាល់": "Queue cleared successfully!",
    "ចំនួនទឹកប្រាក់មិនត្រឹមត្រូវទេ": "Invalid currency amount entered!",
    "សូមជ្រើសរើសប្រភេទ": "Please select a category!",
    "រក្សាទុក និង Sync ជោគជ័យ": "Saved and synced successfully!",
    "សមកាលកម្មបរាជ័យ ប៉ុន្តែទិន្នន័យត្រូវបានរក្សាទុកក្នុង Local": "Sync failed, but data saved locally",
    "រក្សាទុកក្នុង Local និង Sync ជោគជ័យ": "Saved locally and synced successfully!",
    "រក្សាទុកក្នុង Local និងបានដាក់ចូល Queue សម្រាប់ Sync ពេលមាន Internet": "Saved locally and queued for offline sync",
    "រក្សាទុកបានជោគជ័យ (Local Only)": "Saved successfully! (Local storage only)",

    // Toaster default titles
    "ជោគជ័យ": "Success",
    "មានកំហុស": "Error",
    "សមកាលកម្ម": "Sync",
    "ព្រមាន": "Warning",
    "ជូនដំណឹង": "Notification",

    // Google Sheets sync additions
    "កំពុងរក្សាទុក និងសមកាលកម្មទៅ Google Sheet...": "Saving and syncing to Google Sheet...",
    "រក្សាទុកការកំណត់ និងសមកាលកម្មទៅ Google Sheet រួចរាល់!": "Settings saved and synced to Google Sheet successfully!",
    "រក្សាទុកការកំណត់ជោគជ័យ": "Settings saved successfully!",
    "កំពុងរក្សាទុក និងសមកាលកម្ម...": "Saving and syncing...",
    "រក្សាទុកការកំណត់ និងសមកាលកម្មរៀបរយ!": "Settings saved and synced!",
    "រក្សាទុកក្នុងម៉ាស៊ីន ប៉ុន្តែសមកាលកម្មបរាជ័យ!": "Saved locally, but sync failed!",
    "រក្សាទុកជោគជ័យ": "Successfully saved!"
  }
};

const t = (text: string, language: string = 'Khmer'): string => {
  let result = text;
  if (language === 'English') {
    result = TRANSLATIONS.English[text] || text;
  }
  
  if (language === 'English') {
    // Strip Khmer Unicode block: \u1780-\u17FF
    let cleaned = result.replace(/[\u1780-\u17FF]/g, '').trim();
    // Strip empty parentheses and clean spacing
    cleaned = cleaned.replace(/^\(|\)$/g, '').replace(/\s*\(\s*\)/g, '').replace(/\s+/g, ' ').trim();
    return cleaned || result;
  } else {
    // Khmer language. Strip Latin texts in parentheses, e.g. "សមាជិក (Member)" -> "សមាជិក"
    return result.replace(/\s*\([A-Za-z0-9\s,&:\/\?\-\+!.\/]+\)/gi, '').trim();
  }
};

const CATEGORY_TRANSLATIONS: Record<string, string> = {
  // ចំណូល (Income)
  'ប្រាក់ខែ': 'Salary',
  'លក់ដូរ': 'Sales',
  'អាជីវកម្ម': 'Business',
  'ជំនួយ': 'Allowance',
  'ផ្សេងៗ': 'Others',
  
  // ចំណាយ (Expense)
  'អាហារ': 'Food & Drinks',
  'សាំង': 'Fuel & Gasoline',
  'ផ្ទះ': 'Housing',
  'ទឹកភ្លើង': 'Utilities',
  'សុខភាព': 'Healthcare',
  'កូនៗ': 'Children/Kids',
  'សាលារៀន': 'Education',
  'បំណុល': 'Debt/Loans',
  'ទូរស័ព្ទ/Internet': 'Phone & Internet',
};

const translateCategory = (name: string, language: string = 'Khmer'): string => {
  if (language === 'English') {
    return CATEGORY_TRANSLATIONS[name] || name;
  }
  return name;
};

const CurrencyCounter = ({ amount, curr }: { amount: number, curr: string }) => {
  return <CountUp 
    end={amount}
    preserveValue={true} 
    duration={0.6} 
    separator="," 
    decimals={curr === '$' ? 2 : 0} 
    prefix={curr === '$' ? '$' : ''} 
    suffix={curr === '៛' ? ' ៛' : ''} 
  />
};

export default function App() {
  const [isLoaded, setIsLoaded] = useState(false);
  const [session, setSession] = useState<User | null>(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  
  // App State
  const [users, setUsers] = useState<User[]>([]);
  const [records, setRecords] = useState<FinanceRecord[]>([]);
  const [categories, setCategories] = useState<Category[]>([]);
  const [config, setConfig] = useState<AppConfig>({ currency: '$', themeColor: 'emerald' });
  const [toastMsg, setToastMsg] = useState('');
  const [confirmState, setConfirmState] = useState<{message: string, onConfirm: () => void} | null>(null);
  const [promptYearState, setPromptYearState] = useState(false);
  const [progressModal, setProgressModal] = useState<{
    isOpen: boolean;
    title: string;
    percent: number;
    statusText: string;
    subText?: string;
  } | null>(null);
  
  // Swipe State
  const touchStartRef = useRef<{x: number, y: number} | null>(null);

  const handleTouchStart = (e: React.TouchEvent) => {
    touchStartRef.current = { x: e.touches[0].clientX, y: e.touches[0].clientY };
  };

  const handleTouchEnd = (e: React.TouchEvent) => {
    if (!touchStartRef.current) return;
    const touchEndX = e.changedTouches[0].clientX;
    const touchEndY = e.changedTouches[0].clientY;
    
    const deltaX = touchEndX - touchStartRef.current.x;
    const deltaY = Math.abs(touchEndY - touchStartRef.current.y);

    if (touchStartRef.current.x < 40 && deltaX > 50 && deltaY < 60) {
      const isFromSettings = ['import', 'export', 'data', 'members', 'categories', 'googleSheetSync'].includes(activeTab);
      if (isFromSettings) {
        setActiveTab('settings');
      } else if (activeTab !== 'dashboard') {
        setActiveTab('dashboard');
      }
    }
    touchStartRef.current = null;
  };

  // Dynamic Styles
  const getDynamicStyles = () => {
    let css = ``;

    // Theme Color Remap
    if (config.themeColor && config.themeColor !== 'emerald') {
      const target = config.themeColor;
      css += `
        :root {
          --color-emerald-50: var(--color-${target}-50);
          --color-emerald-100: var(--color-${target}-100);
          --color-emerald-200: var(--color-${target}-200);
          --color-emerald-300: var(--color-${target}-300);
          --color-emerald-400: var(--color-${target}-400);
          --color-emerald-500: var(--color-${target}-500);
          --color-emerald-600: var(--color-${target}-600);
          --color-emerald-700: var(--color-${target}-700);
          --color-emerald-800: var(--color-${target}-800);
          --color-emerald-900: var(--color-${target}-900);
          --color-emerald-950: var(--color-${target}-950);
        }
      `;
    }

    // Compact Mode
    if (config.compactMode) {
      css += `
        .p-6 { padding: 16px !important; }
        .p-5 { padding: 12px !important; }
        .py-4 { padding-top: 10px !important; padding-bottom: 10px !important; }
        .gap-4 { gap: 12px !important; }
        .gap-5 { gap: 12px !important; }
        .gap-6 { gap: 16px !important; }
        .rounded-\\[32px\\] { border-radius: 24px !important; }
        .rounded-\\[24px\\] { border-radius: 16px !important; }
      `;
    }

    // Font Size
    if (config.fontSize === 'Small') {
      css += `html { font-size: 14px !important; }`;
    } else if (config.fontSize === 'Large') {
      css += `html { font-size: 18px !important; }`;
    }

    // Card Radius
    if (config.cardRadius && config.cardRadius !== 'rounded-[32px]') { // existing default check
       const radiusMap: Record<string, string> = {
         'none': '0px !important',
         'sm': '8px !important',
         'md': '16px !important',
         'lg': '24px !important',
         'full': '9999px !important',
       };
       if (radiusMap[config.cardRadius]) {
         css += `
           .rounded-\\[32px\\], .rounded-\\[24px\\], .rounded-2xl, .rounded-xl { 
             border-radius: ${radiusMap[config.cardRadius]}; 
           }
         `;
       }
    }

    // Shadows
    if (config.showShadows === false) {
      css += `
        .shadow-sm, .shadow-md, .shadow-lg, .shadow-2xl { box-shadow: none !important; }
      `;
    }

    // Animations
    if (config.enableAnimation === false) {
      css += `
        * {
          transition: none !important;
          animation: none !important;
        }
      `;
    }

    // Dark Mode (Simple Invert Filter version that preserves hues but inverts lightness)
    const isDark = config.themeMode === 'Dark' || (config.themeMode === 'System' && window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    if (isDark) {
      css += `
        html, body { background-color: #111 !important; color: #eee !important; }
        #root > div {
          filter: invert(1) hue-rotate(180deg);
          background-color: #1a1a1a !important; /* Invert of #F5F7F5 is very dark gray */
        }
        /* Re-invert images and certain colorful icons so they look normal */
        img, .avatar, .recharts-wrapper {
          filter: invert(1) hue-rotate(180deg);
        }
      `;
    }

    return css;
  };

  const askConfirm = (message: string, onConfirm: () => void) => { setConfirmState({ message, onConfirm }); };
  const [editRecord, setEditRecord] = useState<FinanceRecord | null>(null);
  const [isLoadingSync, setIsLoadingSync] = useState(false);

  useEffect(() => {
    initStorage();
    setSession(getSession());
    refreshData();
    setIsLoaded(true);
  }, []);

  useEffect(() => {
    document.title = config.appName || 'EasyMoney';
    let favicon = document.querySelector('link[rel="icon"]') as HTMLLinkElement;
    if (!favicon) {
      favicon = document.createElement('link');
      favicon.rel = 'icon';
      document.head.appendChild(favicon);
    }
    favicon.href = config.appLogoUrl || "/favicon.png";
  }, [config.appName, config.appLogoUrl]);

  useEffect(() => {
    const handleOnline = () => {
      const mode = getSyncMode();
      if (mode === 'hybrid') {
        syncLocalDataToGoogleSheet(undefined, session?.name).then(res => {
          if (res.success && res.syncedCount > 0) {
            refreshData();
            showToast(config.language === 'English'
              ? `Sync successful: sent ${res.syncedCount} pending records 📶`
              : `សមកាលកម្មទិន្នន័យជោគជ័យ: បានបញ្ជូនទិន្នន័យកកស្ទះ ${res.syncedCount} ជួរ 📶`
            );
          }
        }).catch(err => { /* Online auto-sync failed */ });
      }
    };
    window.addEventListener('online', handleOnline);
    return () => window.removeEventListener('online', handleOnline);
  }, [session, config.language]);

  useEffect(() => {
    if (isLoaded) {
      const justRestored = localStorage.getItem('EM_JUST_RESTORED') === 'true';
      const mode = getSyncMode();
      const syncUrl = getSyncUrl();

      if (justRestored && mode === 'hybrid' && syncUrl && navigator.onLine) {
        // Show automatic sync startup notification
        showToast(config.language === 'English'
          ? 'Connecting to Google Sheet and auto-syncing... 🔄'
          : 'កំពុងភ្ជាប់ទៅកាន់ Google Sheet និងសមកាលកម្មទិន្នន័យស្វ័យប្រវត្ត... 🔄',
          'info'
        );
        setTimeout(async () => {
          try {
            const curRecords = getRecords();
            const curConfig = getConfig();
            const curUsers = getUsers();
            const curCategories = getCategories();

            // Perform auto push/mapping of all recovered contents to Google Sheet
            const res = await syncLocalDataToGoogleSheet(curRecords, session?.name || 'Sys');
            await syncAppConfigToGoogleSheet(curConfig);
            await syncUsersToGoogleSheet(curUsers);
            await syncCategoriesToGoogleSheet(curCategories);

            localStorage.removeItem('EM_JUST_RESTORED');
            refreshData();
            showToast(config.language === 'English'
              ? 'Auto-connection and sync with Google Sheet successful! 📶✨'
              : 'ភ្ជាប់ស្វ័យប្រវត្តិ និងសមកាលកម្មជាមួយ Google Sheet ជោគជ័យពេញលេញ! 📶✨',
              'success'
            );
          } catch (err) {
            console.error('Auto connection post-restore failed:', err);
            localStorage.removeItem('EM_JUST_RESTORED');
          }
        }, 1000);
      } else {
        if (mode === 'hybrid' && navigator.onLine) {
          setTimeout(() => {
            syncLocalDataToGoogleSheet(undefined, session?.name).then(res => {
              if (res.success && res.syncedCount > 0) {
                refreshData();
                showToast(config.language === 'English'
                  ? `Sync successful: sent ${res.syncedCount} pending records 📶`
                  : `សមកាលកម្មជោគជ័យ: បានបញ្ជូនទិន្នន័យកកស្ទះ ${res.syncedCount} ជួរ 📶`
                );
              }
            }).catch(err => { /* Initial hybrid sync failed */ });
          }, 1500);
        }
      }
    }
  }, [isLoaded, session, config.language]);

  const refreshData = async (skipRemotePull: boolean = false) => {
    const curUsers = getUsers();
    const curCategories = getCategories();
    const curConfig = getConfig();
    setUsers(curUsers);
    setCategories(curCategories);
    setConfig(curConfig);

    // Provide instant UI response by loading local data first
    setRecords(getRecords());

    if (skipRemotePull) return;

    const mode = getSyncMode();
    const syncUrl = getSyncUrl();

    if ((mode === 'sync' || mode === 'hybrid') && syncUrl) {
      setIsLoadingSync(true);
      try {
        const [gasRecords, gasSettings, gasUsers, gasCategories] = await Promise.all([
          loadRecordsFromGoogleSheet(),
          loadSettingsFromGoogleSheet().catch(() => null),
          loadUsersFromGoogleSheet().catch(() => []),
          loadCategoriesFromGoogleSheet().catch(() => [])
        ]);

        if (gasRecords) {
          if (mode === 'hybrid') {
            const queue = getOfflineQueue();
            const gasMap = new Map(gasRecords.map((r: FinanceRecord) => [r.id, r]));
            
            for (const q of queue) {
              if (q.action === 'add' && q.record) {
                gasMap.set(q.id, q.record);
              } else if (q.action === 'update' && q.record) {
                gasMap.set(q.id, q.record);
              } else if (q.action === 'delete') {
                gasMap.delete(q.id);
              }
            }
            const mergedRecords = Array.from(gasMap.values());
            setRecords(mergedRecords);
            saveRecords(mergedRecords);
          } else {
            setRecords(gasRecords);
            saveRecords(gasRecords);
          }
        }

        if (gasSettings && Object.keys(gasSettings).length > 0) {
          const mergedConfig: any = { ...curConfig };
          Object.keys(gasSettings).forEach(k => {
            let val = gasSettings[k];
            if (val === 'true') val = true;
            else if (val === 'false') val = false;
            else if (!isNaN(Number(val)) && val !== '') val = Number(val);
            else if (typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
              try { val = JSON.parse(val); } catch(_) {}
            }
            mergedConfig[k] = val;
          });
          setConfig(mergedConfig);
          saveConfig(mergedConfig);
        }

        if (gasUsers && gasUsers.length > 0) {
          setUsers(gasUsers);
          saveUsers(gasUsers);
        }

        if (gasCategories && gasCategories.length > 0) {
          setCategories(gasCategories);
          saveCategories(gasCategories);
        }
      } catch (err: any) {
        setRecords(getRecords());
        // Simple non-blocking warning toast
        showToast(config.language === 'English'
          ? 'Sync error: using local data instead'
          : 'កំហុសសមកាលកម្ម: បានប្រើប្រាស់ទិន្នន័យ Local ជំនួស',
          'warning'
        );
      } finally {
        setIsLoadingSync(false);
      }
    } else {
      setRecords(getRecords());
    }
  };

  const showToast = (msg: string, type: 'success' | 'error' | 'info' | 'sync' | 'warning' = 'success', title?: string) => {
    // Basic auto-detection based on msg context
    if (msg.includes('កំហុស') || msg.includes('Error') || msg.includes('មិនអាច')) type = 'error';
    if (msg.includes('សមកាលកម្ម') || msg.includes('Sync') || msg.includes('បញ្ជូន')) type = 'sync';
    
    // Remove emojis for a cleaner look
    const cleanMsg = msg.replace(/[\u{1F300}-\u{1F64F}\u{1F680}-\u{1F6FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}\u{1F900}-\u{1F9FF}\u{1F1E0}-\u{1F1FF}]/gu, '').trim();

    let defaultTitle = 'ជូនដំណឹង';
    if (type === 'success') defaultTitle = 'ជោគជ័យ';
    else if (type === 'error') defaultTitle = 'មានកំហុស';
    else if (type === 'sync') defaultTitle = 'សមកាលកម្ម';
    else if (type === 'warning') defaultTitle = 'ព្រមាន';

    const finalTitle = t(title || defaultTitle, config.language);
    const finalMsg = t(cleanMsg || msg, config.language);

    let iconName = 'info';
    let colorClass = 'text-blue-500';
    let bgAccent = 'bg-blue-50';

    if (type === 'success') {
      iconName = 'check_circle';
      colorClass = 'text-emerald-500';
      bgAccent = 'bg-emerald-50';
    } else if (type === 'error') {
      iconName = 'error';
      colorClass = 'text-rose-500';
      bgAccent = 'bg-rose-50';
    } else if (type === 'sync') {
      iconName = 'cloud_sync';
      colorClass = 'text-teal-500';
      bgAccent = 'bg-teal-50';
    } else if (type === 'warning') {
      iconName = 'warning';
      colorClass = 'text-orange-500';
      bgAccent = 'bg-orange-50';
    }
    
    toast.custom((tId) => (
      <div 
        className="w-[350px] max-w-full bg-white/95 backdrop-blur-xl rounded-[20px] border border-slate-100 shadow-[0_16px_36px_rgba(15,23,42,0.06),0_4px_12px_rgba(15,23,42,0.02)] p-4 flex items-start gap-3.5 pointer-events-auto transition-all relative overflow-hidden font-sans"
      >
        <div className={cn("w-9 h-9 rounded-xl flex items-center justify-center shrink-0 border border-slate-100 shadow-[inset_0_1px_1px_rgba(255,255,255,0.4)]", bgAccent)}>
          <span className={cn("material-symbols-rounded text-[20px] select-none", colorClass)}>{iconName}</span>
        </div>
        <div className="flex-1 min-w-0 pt-0.5">
          <div className="flex items-center gap-1.5 mb-1.5 select-none font-sans">
            <h4 className="text-[12px] font-black text-slate-800 leading-none font-sans">{finalTitle}</h4>
            <span className="w-1 h-1 rounded-full bg-slate-200" />
            <span className="text-[8px] font-black text-slate-400 font-mono tracking-wider uppercase">
              {type === 'sync' ? 'REAL-TIME' : 'SYSTEM'}
            </span>
          </div>
          <p className="text-[11px] text-slate-600 font-medium leading-relaxed font-sans select-text">
            {finalMsg}
          </p>
        </div>
        <button 
          onClick={() => toast.dismiss(tId)} 
          className="w-6 h-6 rounded-lg flex items-center justify-center hover:bg-slate-50 text-slate-400 hover:text-slate-600 transition-colors shrink-0"
        >
          <span className="material-symbols-rounded text-[16px]">close</span>
        </button>
      </div>
    ), { duration: 3200 });
  };

  const handleLogin = (user: User) => {
    saveSession(user);
    setSession(user);
    setActiveTab('dashboard');
  };

  const handleLogout = () => {
    saveSession(null);
    setSession(null);
  };

  const handleDeleteRecord = (id: string) => {
    askConfirm(config.language === 'English'
      ? 'Are you sure you want to delete this record?'
      : 'តើអ្នកពិតជាចង់លុបប្រតិបត្តិការនេះមែនទេ?',
      async () => {
        // 1. Local update immediately
        const updated = records.filter(r => r.id !== id);
        saveRecords(updated);
        refreshData(true);
        showToast(config.language === 'English'
          ? 'Deleted successfully!'
          : 'បានលុបដោយជោគជ័យ!'
        );

        // 2. Sync to cloud if needed
        const mode = getSyncMode();
        if (mode === 'sync' || (mode === 'hybrid' && navigator.onLine)) {
          showToast(config.language === 'English'
            ? 'Sending deletion request to Google Sheet...'
            : 'កំពុងបញ្ជូនលុបទៅ Google Sheet...'
          );
          const ok = await deleteRecordFromGoogleSheet(id, session.name);
          if (ok) {
            showToast(config.language === 'English'
              ? 'Deleted from Google Sheet successfully! 🗑️'
              : 'បានលុបរួចរាល់ពី Google Sheet 🗑️'
            );
            refreshData(true);
          } else {
            if (mode === 'hybrid') {
              addToOfflineQueue('delete', id);
              showToast(config.language === 'English'
                ? 'No Google Sheet connection. Deletion queued offline.'
                : 'គ្មានការភ្ជាប់ចូល Google Sheet, បានបម្រុងទុក Offline'
              );
            } else {
              showToast(config.language === 'English'
                ? 'Could not delete directly from Google Sheet (Network Error)'
                : 'មិនអាចលុបដោយផ្ទាល់ពី Google Sheet ទេ (Network Error)'
              );
            }
          }
        } else if (mode === 'hybrid') {
          addToOfflineQueue('delete', id);
          showToast(config.language === 'English'
            ? 'Offline: Deletion saved to your offline queue.'
            : 'Offline: បានរក្សាទុកក្នុង Offline Queue របស់លោកអ្នក'
          );
        }
      }
    );
  };

  if (!isLoaded) {
    return (
      <div className="min-h-screen bg-[#F5F7F5] flex flex-col items-center justify-center gap-4 text-emerald-900 font-bold">
        <img src="/easymoney-logo.png" alt="Loading" className="w-24 h-24 object-contain animate-pulse mb-2 drop-shadow-xl" />
        <div className="text-sm tracking-wide">កំពុងដំណើរការ...</div>
      </div>
    );
  }

  if (!session) {
    return <LoginView users={users} onLogin={handleLogin} config={config} />;
  }

  return (
    <div 
      className="fixed inset-0 flex flex-col w-full md:max-w-xl lg:max-w-3xl xl:max-w-4xl sm:mx-auto bg-[#F5F7F5] text-[#2D3436] font-sans shadow-2xl overflow-hidden md:border-x border-[#E8EEE9]"
      onTouchStart={handleTouchStart}
      onTouchEnd={handleTouchEnd}
    >
      <style dangerouslySetInnerHTML={{ __html: getDynamicStyles() }} />
      <Toaster position="top-center" offset="16px" />
      <AnimatePresence>
        {/* Confirm Notification */}
        {confirmState && (
          <div className="fixed inset-0 z-[100] flex items-end sm:items-center justify-center sm:p-4">
            <motion.div initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} onClick={() => setConfirmState(null)} className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" />
            <motion.div initial={{ y: '100%', opacity: 0 }} animate={{ y: 0, opacity: 1 }} exit={{ y: '100%', opacity: 0 }} transition={{ type: "spring", bounce: 0, duration: 0.4 }} className="relative bg-white rounded-t-[32px] sm:rounded-[32px] p-6 w-full max-w-[340px] shadow-2xl flex flex-col gap-4 font-sans border border-[#E8EEE9] pb-safe z-10">
              <div className="w-12 h-1.5 bg-slate-200 rounded-full mx-auto mb-2 sm:hidden" />
              <div className="flex gap-4">
                <div className="w-12 h-12 bg-red-50 rounded-[20px] flex items-center justify-center text-red-500 shrink-0 border border-red-100">
                  <AlertCircle className="w-6 h-6" />
                </div>
                <div className="pt-1">
                  <h3 className="font-extrabold text-slate-800 text-sm leading-tight mb-1">សូមបញ្ជាក់ (Confirm)</h3>
                  <p className="text-xs text-slate-500 font-bold leading-relaxed">{confirmState.message}</p>
                </div>
              </div>
              <div className="flex gap-2 mt-2">
                <button onClick={() => setConfirmState(null)} className="flex-1 py-3.5 bg-[#F5F7F5] hover:bg-[#E8EEE9] text-slate-600 outline-none rounded-[16px] font-extrabold text-[11px] uppercase tracking-wider transition-colors active:scale-95">បោះបង់ (Cancel)</button>
                <button 
                  onClick={() => {
                    confirmState.onConfirm();
                    setConfirmState(null);
                  }} 
                  className="flex-1 py-3.5 bg-red-500 hover:bg-red-600 text-white outline-none rounded-[16px] font-extrabold text-[11px] uppercase tracking-wider transition-colors shadow-lg shadow-red-500/30 active:scale-95">
                  យល់ព្រមលុប
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {/* Progress Modal (Restoring or Importing Data) */}
        {progressModal && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }} 
              animate={{ opacity: 1 }} 
              exit={{ opacity: 0 }} 
              className="absolute inset-0 bg-slate-950/60 backdrop-blur-md" 
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }} 
              animate={{ scale: 1, opacity: 1 }} 
              exit={{ scale: 0.9, opacity: 0 }} 
              transition={{ type: "spring", bounce: 0.1, duration: 0.4 }} 
              className="relative bg-white rounded-[28px] p-6 w-full max-w-[340px] shadow-[0_24px_50px_rgba(0,0,0,0.2)] flex flex-col items-center text-center font-sans border border-slate-100 z-10"
            >
              <div className="relative w-20 h-20 flex items-center justify-center mb-4">
                <svg className="w-20 h-20 -rotate-90">
                  <circle cx="40" cy="40" r="34" className="stroke-slate-100 fill-none" strokeWidth="6" />
                  <circle 
                    cx="40" 
                    cy="40" 
                    r="34" 
                    className="stroke-emerald-500 fill-none transition-all duration-300" 
                    strokeWidth="6" 
                    strokeDasharray={2 * Math.PI * 34} 
                    strokeDashoffset={2 * Math.PI * 34 * (1 - (progressModal.percent || 0) / 100)} 
                    strokeLinecap="round"
                  />
                </svg>
                <div className="absolute inset-0 flex flex-col items-center justify-center">
                  <span className="text-[15px] font-black text-slate-800 leading-none">{Math.round(progressModal.percent || 0)}%</span>
                </div>
              </div>

              <h3 className="font-extrabold text-slate-800 text-[13px] tracking-wide mb-1 leading-snug">{progressModal.title}</h3>
              <p className="text-[10px] text-slate-400 font-extrabold tracking-wider uppercase mb-3.5">{progressModal.statusText}</p>
              
              {progressModal.subText && (
                <div className="px-4 py-2.5 bg-slate-50 border border-slate-100 rounded-xl text-[10px] text-slate-500 font-bold max-w-full truncate leading-relaxed">
                  {progressModal.subText}
                </div>
              )}
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {/* Year Prompt Modal */}
        {promptYearState && (
          <div className="fixed inset-0 z-[100] flex items-end sm:items-center justify-center sm:p-4">
            <motion.div initial={{opacity:0}} animate={{opacity:1}} exit={{opacity:0}} onClick={() => setPromptYearState(false)} className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" />
            <motion.div initial={{ y: '100%', opacity: 0 }} animate={{ y: 0, opacity: 1 }} exit={{ y: '100%', opacity: 0 }} transition={{ type: "spring", bounce: 0, duration: 0.4 }} className="relative bg-white rounded-t-[32px] sm:rounded-[32px] p-6 w-full max-w-[340px] shadow-2xl flex flex-col gap-4 font-sans border border-[#E8EEE9] pb-safe z-10">
              <div className="w-12 h-1.5 bg-slate-200 rounded-full mx-auto mb-2 sm:hidden" />
              <div className="text-center">
                <h3 className="font-extrabold text-slate-800 text-sm leading-tight mb-2">
                  {config.language === 'English' ? 'Add New Year' : 'បន្ថែមឆ្នាំថ្មី'}
                </h3>
                <p className="text-[11px] text-slate-500 font-bold leading-relaxed mb-4">
                  {config.language === 'English' ? 'Please enter a 4-digit year (e.g., 2030)' : 'សូមវាយបញ្ចូលឆ្នាំ ៤ខ្ទង់ (ឧ. 2030)'}
                </p>
                <form onSubmit={(e: any) => {
                  e.preventDefault();
                  const val = e.target.yearInput.value;
                  const newYear = parseInt(val);
                  if (val.length === 4 && !isNaN(newYear)) {
                    const freshConfig = { ...config, customYears: Array.from(new Set([...(config.customYears || []), newYear])) };
                    import('./store').then(m => {
                      m.saveConfig(freshConfig);
                      setConfig(freshConfig);
                      syncAppConfigToGoogleSheet(freshConfig);
                      setPromptYearState(false);
                      showToast(config.language === 'English' ? 'Year added successfully! 📅' : 'បានបន្ថែមឆ្នាំជោគជ័យ 📅');
                    });
                  } else {
                    showToast(config.language === 'English' ? 'Please enter a valid 4-digit year! ⚠️' : 'សូមបញ្ចូលឆ្នាំជាលេខ ៤ខ្ទង់ ⚠️');
                  }
                }}>
                  <input name="yearInput" type="number" min="1900" max="2100" className="w-full bg-[#F5F7F5] border-2 border-[#E8EEE9] px-4 py-3 rounded-2xl text-center text-sm font-extrabold text-slate-800 outline-none focus:border-emerald-500 transition-all mb-4" placeholder="YYYY" autoFocus />
                  <div className="flex gap-2">
                    <button type="button" onClick={() => setPromptYearState(false)} className="flex-1 py-3 bg-[#F5F7F5] hover:bg-[#E8EEE9] text-slate-600 outline-none rounded-2xl font-extrabold text-[11px] uppercase tracking-wider transition-colors active:scale-95">
                      {config.language === 'English' ? 'Cancel' : 'បោះបង់'}
                    </button>
                    <button type="submit" className="flex-1 py-3 bg-emerald-500 hover:bg-emerald-600 text-white outline-none rounded-2xl font-extrabold text-[11px] uppercase tracking-wider transition-colors shadow-lg shadow-emerald-500/30 active:scale-95">
                      {config.language === 'English' ? 'Save' : 'រក្សាទុក'}
                    </button>
                  </div>
                </form>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Header */}
      <header className="app-header w-full flex-none flex items-center justify-between px-6 bg-white border-b border-[#E8EEE9] shadow-sm z-30">
        <button onClick={() => setActiveTab('dashboard')} className="flex items-center gap-2 outline-none hover:opacity-80 transition-opacity">
          <img src={config.appLogoUrl || "/easymoney-logo.png"} alt="Logo" className="w-10 h-10 object-contain" />
          <h1 className="text-xl font-bold tracking-tight text-emerald-900">{config.appName || 'EasyMoney'}</h1>
        </button>
        <div className="flex items-center gap-4">
          <div className="text-right">
            <p className="text-[10px] text-[#636E72] uppercase tracking-widest font-semibold">
              {config.language === 'English' ? `Hello ${session.role}` : `សួស្ដី ${session.role}`}
            </p>
            <p className="text-xs font-bold text-slate-800">{session.name}</p>
          </div>
          <button onClick={handleLogout} className="w-10 h-10 rounded-full bg-[#F5F7F5] flex items-center justify-center text-emerald-700 hover:bg-emerald-100 transition-colors border border-[#E8EEE9] shadow-sm">
            <LogOut className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="flex-1 overflow-y-auto overflow-x-hidden w-full pt-safe pb-[calc(64px+env(safe-area-inset-bottom))]">
        <AnimatePresence mode="wait" initial={false}>
          <motion.div
            key={activeTab}
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -20 }}
            transition={{ duration: 0.2, ease: "easeOut" }}
            className="w-full min-h-full"
          >
            {activeTab === 'dashboard' && <DashboardView records={records} session={session} currency={config.currency} askConfirm={askConfirm} config={config} setPromptYearState={setPromptYearState} />}
            {activeTab === 'records' && <RecordsView records={records} session={session} currency={config.currency} onUpdate={() => refreshData(true)} showToast={showToast} users={users} onEdit={(r: any) => { setEditRecord(r); setActiveTab('add'); }} onDeleteTrigger={handleDeleteRecord} config={config} setPromptYearState={setPromptYearState} />}
            {activeTab === 'add' && <AddRecordView session={session} categories={categories} onUpdate={() => refreshData(true)} showToast={showToast} setActiveTab={setActiveTab} editRecord={editRecord} users={users} config={config} />}
            {activeTab === 'settings' && <SettingsView session={session} setActiveTab={setActiveTab} config={config} onUpdate={() => refreshData(true)} showToast={showToast} records={records} users={users} askConfirm={askConfirm} setPromptYearState={setPromptYearState} setProgressModal={setProgressModal} />}
            {activeTab === 'import' && <ImportView session={session} onUpdate={() => refreshData(true)} showToast={showToast} setActiveTab={setActiveTab} config={config} users={users} setProgressModal={setProgressModal} />}
            {activeTab === 'export' && <ExportView records={records} users={users} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} config={config} setPromptYearState={setPromptYearState} />}
            {activeTab === 'data' && <DataCenterView session={session} setActiveTab={setActiveTab} records={records} users={users} config={config} />}
            {activeTab === 'members' && <MembersView users={users} session={session} onUpdate={() => refreshData(true)} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} config={config} />}
            {activeTab === 'categories' && <CategoriesView categories={categories} onUpdate={() => refreshData(true)} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} config={config} />}
            {activeTab === 'googleSheetSync' && <GoogleSheetSyncView session={session} records={records} onUpdate={() => refreshData(true)} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} config={config} />}
          </motion.div>
        </AnimatePresence>
      </main>

      {/* Bottom Nav */}
      <nav 
        className="fixed bottom-0 left-0 right-0 w-full md:max-w-xl lg:max-w-3xl xl:max-w-4xl sm:mx-auto bg-white/70 backdrop-blur-2xl border-t border-[#E8EEE9]/60 z-40 shadow-[0_-8px_30px_rgba(0,0,0,0.05)]"
        style={{ paddingBottom: 'env(safe-area-inset-bottom)' }}
      >
        <div className="flex justify-around items-center h-[56px] px-2">
          <NavItem icon="pie_chart" label={t("សរុប", config.language)} isActive={activeTab === 'dashboard'} onClick={() => setActiveTab('dashboard')} />
          <NavItem icon="list_alt" label={t("បញ្ជី", config.language)} isActive={activeTab === 'records'} onClick={() => setActiveTab('records')} />
          
          <div className="flex-1 flex justify-center -mt-5">
            <button 
              onClick={() => { setEditRecord(null); setActiveTab('add'); }} 
              className="group relative flex items-center justify-center outline-none active:scale-95 transition-all duration-300 drop-shadow-[0_4px_16px_rgba(16,185,129,0.3)] hover:drop-shadow-[0_8px_20px_rgba(16,185,129,0.4)]"
            >
              <div className="relative w-[50px] h-[50px] bg-gradient-to-tr from-emerald-500 to-teal-600 group-hover:from-emerald-400 group-hover:to-teal-500 text-white rounded-[20px] rotate-3 group-hover:rotate-6 flex items-center justify-center border-4 border-white z-10 transition-all duration-300">
                <span className="material-symbols-rounded text-3xl -rotate-3 group-hover:-rotate-6 transition-all">add</span>
              </div>
            </button>
          </div>
          
          <NavItem icon="database" label={t("ទិន្នន័យ", config.language)} isActive={['data', 'import', 'export'].includes(activeTab)} onClick={() => setActiveTab('data')} />
          <NavItem icon="settings" label={t("កំណត់", config.language)} isActive={['settings', 'members', 'categories', 'googleSheetSync'].includes(activeTab)} onClick={() => setActiveTab('settings')} />
        </div>
      </nav>
    </div>
  );
}

// -------- COMMON COMPONENTS --------

const CheckCircle = ({ className }: { className?: string }) => (
  <svg className={className} xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
);

function NavItem({ icon, label, isActive, onClick }: { icon: string, label: string, isActive: boolean, onClick: () => void }) {
  return (
    <button onClick={onClick} className="relative flex flex-col items-center justify-center gap-1 transition-all duration-300 outline-none w-16 h-14 rounded-2xl active:scale-95 group">
      {isActive && (
        <motion.div layoutId="nav-pill" className="absolute inset-0 bg-emerald-50 rounded-2xl -z-10" transition={{ type: "spring", stiffness: 400, damping: 30 }} />
      )}
      <span className={cn("material-symbols-rounded transition-all duration-300", isActive ? "text-emerald-600 text-2xl" : "text-slate-400 group-hover:text-emerald-500 text-[22px]")}>
        {icon}
      </span>
      <span className={cn("text-[8px] font-extrabold tracking-wide transition-all duration-300", isActive ? "text-emerald-600" : "text-slate-400 group-hover:text-emerald-500")}>
        {label}
      </span>
    </button>
  );
}

// -------- VIEWS --------

function LoginView({ users, onLogin, config }: { users: User[], onLogin: (u: User) => void, config: AppConfig }) {
  const [selectedUserId, setSelectedUserId] = useState('');
  const [pin, setPin] = useState('');
  const [error, setError] = useState('');

  const handleLogin = () => {
    const u = users.find(u => u.id === selectedUserId);
    if (!u) { setError('សូមជ្រើសរើសឈ្មោះសមាជិក'); return; }
    if (String(u.pin).trim() !== String(pin).trim()) { setError('លេខសម្ងាត់មិនត្រឹមត្រូវទេ'); return; }
    if (!u.isActive) { setError('គណនីនេះត្រូវបានបិទ'); return; }
    onLogin(u);
  };

  return (
    <div className="h-screen w-full md:max-w-xl lg:max-w-3xl xl:max-w-4xl sm:mx-auto bg-[#F5F7F5] flex items-center justify-center p-6 text-center shadow-2xl relative overflow-hidden md:border-x border-[#E8EEE9]">
      {/* Decorative blobs */}
      <div className="absolute top-0 left-0 w-64 h-64 bg-emerald-500 rounded-full mix-blend-multiply filter blur-3xl opacity-20 -translate-x-1/2 -translate-y-1/2"></div>
      <div className="absolute bottom-0 right-0 w-64 h-64 bg-teal-500 rounded-full mix-blend-multiply filter blur-3xl opacity-20 translate-x-1/2 translate-y-1/2"></div>
      
      <div className="bg-white/80 backdrop-blur-xl rounded-[40px] shadow-2xl shadow-emerald-900/5 border border-white w-full p-8 text-center relative z-10 transition-all">
        <img src={config?.appLogoUrl || "/easymoney-logo.png"} alt="Logo" className="w-24 h-24 object-contain mx-auto mb-6 drop-shadow-xl" />
        <h1 className="text-3xl font-extrabold tracking-tight text-emerald-950 mb-2">{config?.appName || 'EasyMoney'}</h1>
        <p className="text-[10px] uppercase font-bold tracking-widest text-[#636E72] mb-10">កត់ត្រាចំណូល-ចំណាយគ្រួសារ</p>
        
        {error && <div className="mb-6 text-red-600 text-xs font-bold bg-red-50 p-4 rounded-2xl border border-red-100 animate-in fade-in zoom-in-95">{error}</div>}

        <div className="space-y-5 text-left">
          <div>
            <label className="text-[10px] uppercase font-bold tracking-widest text-[#636E72] block mb-2 ml-2">ជ្រើសរើសគណនី</label>
            <CustomSelect 
              value={selectedUserId} 
              onChange={setSelectedUserId}
              placeholder="-- សូមជ្រើសរើសគណនី --"
              icon={UserIcon}
              modalTitle="ជ្រើសរើសគណនី (Select Account)"
              options={users.filter(u => u.isActive).map(u => ({
                value: u.id,
                label: u.name,
                description: u.role === 'Admin' ? 'អ្នកគ្រប់គ្រង' : 'សមាជិក',
                icon: <UserIcon className="w-5 h-5" />
              }))}
            />
          </div>
          <div>
            <label className="text-[10px] uppercase font-bold tracking-widest text-[#636E72] block mb-2 ml-2">លេខសម្ងាត់ PIN</label>
            <input type="password" value={pin} onChange={e => setPin(e.target.value)} pattern="[0-9]*" inputMode="numeric" className="w-full px-5 py-4 rounded-3xl border-2 border-[#E8EEE9] focus:ring-4 focus:ring-emerald-500/20 focus:border-emerald-500 outline-none transition bg-white text-center tracking-[1em] font-extrabold text-xl text-emerald-950 placeholder:text-slate-300 placeholder:tracking-normal placeholder:font-bold placeholder:text-sm shadow-sm" placeholder="បញ្ចូល PIN ៤ខ្ទង់" />
          </div>
          <button onClick={handleLogin} className="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold py-5 rounded-3xl shadow-xl shadow-slate-900/20 transition mt-8 flex justify-center items-center gap-2 text-sm active:scale-[0.98]">
            ចូលប្រើប្រាស់
            <span className="material-symbols-rounded text-base text-emerald-400">arrow_forward</span>
          </button>
        </div>
      </div>
    </div>
  );
}

function DashboardView({ records, session, currency, askConfirm, config, setPromptYearState }: { records: FinanceRecord[], session: User, currency: string, askConfirm: any, config: AppConfig, setPromptYearState: any }) {
  const d = new Date();
  const [filterMonth, setFilterMonth] = useState(config.defaultMonth || (d.getMonth() + 1));
  const [filterYear, setFilterYear] = useState(config.defaultYear || d.getFullYear());
  const [summaryMode, setSummaryMode] = useState<'year' | 'month'>('month');

  const viewableRecords = useMemo(() => session.role === 'Admin' ? records : records.filter(r => r.memberId === session.id), [records, session.role, session.id]);
  const monthRecords = useMemo(() => viewableRecords.filter(r => r.month === filterMonth && r.year === filterYear), [viewableRecords, filterMonth, filterYear]);
  const yearRecords = useMemo(() => viewableRecords.filter(r => r.year === filterYear), [viewableRecords, filterYear]);

  const exRate = config.exchangeRate || 4000;
  const getAmt = useCallback((r: FinanceRecord, targetCurrency: string) => {
    const rCurr = r.currency || currency;
    if (targetCurrency === '$') {
      return r.amountUSD ?? (rCurr === '$' ? r.amount : r.amount / exRate);
    } else {
      return r.amountKHR ?? (rCurr === '៛' ? r.amount : r.amount * exRate);
    }
  }, [currency, exRate]);

  const mIncome = useMemo(() => monthRecords.filter(r => r.type === 'ចំណូល').reduce((s, r) => s + getAmt(r, currency), 0), [monthRecords, getAmt, currency]);
  const mExpense = useMemo(() => monthRecords.filter(r => r.type === 'ចំណាយ').reduce((s, r) => s + getAmt(r, currency), 0), [monthRecords, getAmt, currency]);
  const mBal = mIncome - mExpense;

  const yIncome = useMemo(() => yearRecords.filter(r => r.type === 'ចំណូល').reduce((s, r) => s + getAmt(r, currency), 0), [yearRecords, getAmt, currency]);
  const yExpense = useMemo(() => yearRecords.filter(r => r.type === 'ចំណាយ').reduce((s, r) => s + getAmt(r, currency), 0), [yearRecords, getAmt, currency]);
  const yBal = yIncome - yExpense;

  const altCurrency = currency === '$' ? '៛' : '$';
  const getAltAmount = useCallback((amount: number) => currency === '$' ? amount * exRate : amount / exRate, [currency, exRate]);

  const isYearly = summaryMode === 'year';
  const displayBal = isYearly ? yBal : mBal;
  const displayIncome = isYearly ? yIncome : mIncome;
  const displayExpense = isYearly ? yExpense : mExpense;
  const displayTitle = isYearly ? "សមតុល្យសរុប (ឆ្នាំនេះ)" : "សមតុល្យសរុប (ខែនេះ)";

  const monthlyBreakdown = useMemo(() => Array.from({ length: 12 }).map((_, idx) => {
    const mNum = idx + 1;
    const recs = yearRecords.filter(r => r.month === mNum);
    const inc = recs.filter(r => r.type === 'ចំណូល').reduce((sum, r) => sum + getAmt(r, currency), 0);
    const exp = recs.filter(r => r.type === 'ចំណាយ').reduce((sum, r) => sum + getAmt(r, currency), 0);
    const bal = inc - exp;
    return {
      monthNum: mNum,
      monthName: t(KHMER_MONTHS[idx], config.language),
      income: inc,
      expense: exp,
      balance: bal,
      hasData: recs.length > 0
    };
  }), [yearRecords, getAmt, currency, config.language]);

  const currentRecordsForChart = summaryMode === 'year' ? yearRecords : monthRecords;
  const expByCategory = useMemo(() => currentRecordsForChart.filter(r => r.type === 'ចំណាយ').reduce((acc, r) => {
    acc[r.category] = (acc[r.category] || 0) + getAmt(r, currency);
    return acc;
  }, {} as Record<string, number>), [currentRecordsForChart, getAmt, currency]);
  
  const chartData = useMemo(() => Object.keys(expByCategory).map((k, i) => ({ name: translateCategory(k, config.language), value: expByCategory[k], fill: ['#059669', '#10b981', '#34d399', '#6ee7b7', '#f43f5e', '#fb923c', '#eab308', '#6366f1'][i % 8] })), [expByCategory, config.language]);

  const trendData = useMemo(() => Array.from({ length: 6 }).map((_, i) => {
    const td = new Date(filterYear, filterMonth - 1 - i, 1);
    const mMonth = td.getMonth() + 1;
    const mYear = td.getFullYear();
    const recs = viewableRecords.filter((r:any) => r.month === mMonth && r.year === mYear);
    return {
      name: t(KHMER_MONTHS[mMonth - 1], config.language),
      ចំណាយ: recs.filter((r:any) => r.type === 'ចំណាយ').reduce((s:number, r:any) => s + getAmt(r, currency), 0),
      ចំណូល: recs.filter((r:any) => r.type === 'ចំណូល').reduce((s:number, r:any) => s + getAmt(r, currency), 0)
    };
  }).reverse(), [filterYear, filterMonth, viewableRecords, getAmt, currency, config.language]);

  return (
    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.3 }} className="p-6 space-y-6 pb-24">
      {/* Date Filters */}
      <div className="flex gap-3 relative z-50">
        <div className="flex-1">
          <CustomSelect 
            value={filterMonth} 
            onChange={setFilterMonth}
            className="px-5 py-4 rounded-3xl text-sm font-bold text-emerald-950 shadow-sm"
            modalTitle="ជ្រើសរើសខែ (Select Month)"
            options={[1,2,3,4,5,6,7,8,9,10,11,12].map(m => ({ value: m, label: t(KHMER_MONTHS[m - 1], config.language) }))}
          />
        </div>
        <div className="w-32">
          <CustomSelect 
            value={filterYear} 
            onChange={(v: any) => { if (v === 'ADD_NEW_YEAR') setPromptYearState(true); else setFilterYear(v); }}
            className="px-5 py-4 rounded-3xl text-sm font-bold text-emerald-950 shadow-sm"
            modalTitle="ជ្រើសរើសឆ្នាំ (Select Year)"
            options={[...generateYearOptions(records, config.customYears), { value: 'ADD_NEW_YEAR', label: '+ បន្ថែមឆ្នាំថ្មី' }]}
          />
        </div>
      </div>

      {/* Hero Balance Card */}
      <AnimatePresence mode="wait">
        <motion.div 
          key={`hero-${summaryMode}-${filterMonth}-${filterYear}`}
          initial={{ opacity: 0, y: 20, scale: 0.95 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          exit={{ opacity: 0, y: -20, scale: 0.95 }}
          transition={{ duration: 0.3, ease: "easeOut" }}
          className={cn("rounded-[32px] p-7 text-white shadow-2xl transition-all relative overflow-hidden", displayBal >= 0 ? "bg-gradient-to-br from-emerald-600 to-teal-800 shadow-emerald-200/50" : "bg-gradient-to-br from-red-500 to-rose-700 shadow-red-200/50")}
        >
          <div className="absolute top-0 right-0 w-48 h-48 bg-white opacity-5 rounded-full blur-3xl -translate-y-1/2 translate-x-1/4"></div>
          <div className="relative z-10">
            <div className="flex items-center gap-2 mb-1">
              <p className="text-white/80 text-[11px] font-bold tracking-widest uppercase mb-1">{t(displayTitle, config.language)}</p>
            </div>
            <div className="mb-8">
              <h2 className="text-[40px] font-extrabold tracking-tight drop-shadow-sm leading-none">
                <CurrencyCounter amount={displayBal} curr={currency} />
              </h2>
              <p className="text-white/70 font-bold mt-2 text-sm">~ <CurrencyCounter amount={getAltAmount(displayBal)} curr={altCurrency} /></p>
            </div>
            <div className="flex justify-between items-center bg-black/10 backdrop-blur-md p-4 rounded-2xl border border-white/10 shadow-inner">
              <div>
                <p className="text-white/70 text-[9px] uppercase font-extrabold tracking-widest mb-1 flex items-center gap-1"><span className="material-symbols-rounded text-[14px] text-emerald-400">trending_up</span> {t("ចំណូល", config.language)}</p>
                <p className="text-lg font-bold leading-tight"><CurrencyCounter amount={displayIncome} curr={currency} /></p>
                <p className="text-white/60 text-[10px] font-bold mt-0.5">~ <CurrencyCounter amount={getAltAmount(displayIncome)} curr={altCurrency} /></p>
              </div>
              <div className="w-px h-10 bg-white/10"></div>
              <div className="text-right">
                <p className="text-white/70 text-[9px] uppercase font-extrabold tracking-widest mb-1 flex items-center gap-1 justify-end"><span className="material-symbols-rounded text-[14px] text-red-400">trending_down</span> {t("ចំណាយ", config.language)}</p>
                <p className="text-lg font-bold leading-tight"><CurrencyCounter amount={displayExpense} curr={currency} /></p>
                <p className="text-white/60 text-[10px] font-bold mt-0.5">~ <CurrencyCounter amount={getAltAmount(displayExpense)} curr={altCurrency} /></p>
              </div>
            </div>
          </div>
        </motion.div>
      </AnimatePresence>

      {/* Chart Section */}
      {chartData.length > 0 && (
        <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.2 }} className="bg-white rounded-[32px] p-6 border border-[#E8EEE9] shadow-sm hover:shadow-md transition-shadow">
          <h4 className="text-[11px] font-extrabold tracking-widest uppercase text-emerald-900 mb-6 text-center">
            {summaryMode === 'year' ? t("តារាងចំណាយឆ្នាំនេះ", config.language) : t("តារាងចំណាយខែនេះ", config.language)}
          </h4>
          <div className="h-[200px]">
            <ResponsiveContainer width="100%" height="100%">
              <RePieChart>
                <Pie data={chartData} cx="50%" cy="50%" innerRadius={60} outerRadius={90} paddingAngle={3} dataKey="value" stroke="none" cornerRadius={6}>
                  {chartData.map((entry, index) => <Cell key={index} fill={entry.fill} />)}
                </Pie>
                <ReTooltip formatter={(value: number) => formatMoney(value, currency)} contentStyle={{ borderRadius: '20px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)', padding: '12px 16px', fontWeight: 'bold', fontSize: '14px', color: '#064e3b' }} itemStyle={{ color: '#064e3b' }} />
              </RePieChart>
            </ResponsiveContainer>
          </div>
        </motion.div>
      )}

      {/* Spending Trend Chart */}
      <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} transition={{ delay: 0.3 }} className="bg-white rounded-[32px] p-6 border border-[#E8EEE9] shadow-sm hover:shadow-md transition-shadow">
        <h4 className="text-[11px] font-extrabold tracking-widest uppercase text-emerald-900 mb-6 text-center">{t("និន្នាការចំណាយ ៦ ខែចុងក្រោយ", config.language)}</h4>
        <div className="h-[200px]">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={trendData} margin={{ top: 5, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E8EEE9" />
              <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b', fontWeight: 'bold' }} dy={10} />
              <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b', fontWeight: 'bold' }} tickFormatter={(val) => Math.abs(val) > 999 ? `${(val/1000).toFixed(1)}k` : val} />
              <ReTooltip formatter={(value: number) => formatMoney(value, currency)} contentStyle={{ borderRadius: '16px', border: '1px solid #E8EEE9', boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.05)', padding: '8px 12px', fontSize: '12px', fontWeight: 'bold', color: '#064e3b' }} />
              <Line type="monotone" dataKey="ចំណាយ" stroke="#f43f5e" strokeWidth={3} dot={{ r: 4, fill: '#f43f5e', strokeWidth: 2, stroke: '#fff' }} activeDot={{ r: 6 }} />
              <Line type="monotone" dataKey="ចំណូល" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: '#10b981', strokeWidth: 2, stroke: '#fff' }} activeDot={{ r: 6 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </motion.div>

      {/* Summary View */}
      <motion.div 
        initial={{ opacity: 0, y: 20, scale: 0.95 }}
        animate={{ opacity: 1, y: 0, scale: 1 }}
        transition={{ duration: 0.3, ease: "easeOut", delay: 0.1 }}
        className="bg-white rounded-[32px] p-7 border border-[#E8EEE9] shadow-sm"
      >
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6 border-b border-[#F0F4F1] pb-5">
          <h4 className="text-sm font-extrabold text-emerald-950 flex items-center gap-2">
            {t("របាយការណ៍សរុប", config.language)}
          </h4>
          
          <div className="flex bg-[#F5F7F5] p-1 rounded-2xl w-full sm:w-auto">
            <button 
              type="button"
              onClick={() => setSummaryMode('month')}
              className={cn("flex-1 sm:flex-none px-4 py-2 text-xs font-bold rounded-xl transition-all", summaryMode === 'month' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-550 hover:text-emerald-600")}
            >
              {t("ប្រចាំខែ", config.language)}
            </button>
            <button 
              type="button"
              onClick={() => setSummaryMode('year')}
              className={cn("flex-1 sm:flex-none px-4 py-2 text-xs font-bold rounded-xl transition-all", summaryMode === 'year' ? "bg-white text-emerald-700 shadow-sm" : "text-slate-550 hover:text-emerald-600")}
            >
              {t("ប្រចាំឆ្នាំ", config.language)}
            </button>
          </div>
        </div>

        <div>
          <AnimatePresence mode="wait">
            <motion.div 
              key={`summary-${summaryMode}-${filterMonth}-${filterYear}`}
              initial={{ opacity: 0, x: -10 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: 10 }}
              transition={{ duration: 0.2 }}
              className="space-y-4"
            >
              <div className="flex justify-between items-center">
                <span className="text-[11px] font-bold text-slate-500 uppercase tracking-widest">{summaryMode === 'year' ? t('ចំណូលពេញឆ្នាំ', config.language) : t('ចំណូលខែនេះ', config.language)}</span>
                <div className="text-right">
                  <span className="text-sm font-bold text-emerald-600 block leading-tight"><CurrencyCounter amount={summaryMode === 'year' ? yIncome : mIncome} curr={currency} /></span>
                  <span className="text-[10px] font-bold text-slate-400">~ <CurrencyCounter amount={getAltAmount(summaryMode === 'year' ? yIncome : mIncome)} curr={altCurrency} /></span>
                </div>
              </div>
              <div className="flex justify-between items-center">
                <span className="text-[11px] font-bold text-slate-500 uppercase tracking-widest">{summaryMode === 'year' ? t('ចំណាយពេញឆ្នាំ', config.language) : t('ចំណាយខែនេះ', config.language)}</span>
                <div className="text-right">
                  <span className="text-sm font-bold text-red-500 block leading-tight"><CurrencyCounter amount={summaryMode === 'year' ? yExpense : mExpense} curr={currency} /></span>
                  <span className="text-[10px] font-bold text-slate-400">~ <CurrencyCounter amount={getAltAmount(summaryMode === 'year' ? yExpense : mExpense)} curr={altCurrency} /></span>
                </div>
              </div>
              <div className="flex justify-between items-end pt-4 border-t border-[#F0F4F1] mt-2">
                <span className="text-xs font-extrabold text-emerald-950 mb-1">{summaryMode === 'year' ? t('ប្រាក់សល់ប្រចាំឆ្នាំ', config.language) : t('ប្រាក់សល់ខែនេះ', config.language)}</span>
                <div className="text-right">
                  <span className={cn("text-lg font-extrabold border-b-2 border-dashed pb-0.5 block leading-tight mb-1", (summaryMode === 'year' ? yBal : mBal) >= 0 ? "text-emerald-700 border-emerald-300" : "text-red-600 border-red-300")}><CurrencyCounter amount={summaryMode === 'year' ? yBal : mBal} curr={currency} /></span>
                  <span className="text-xs font-bold text-slate-500">~ <CurrencyCounter amount={getAltAmount(summaryMode === 'year' ? yBal : mBal)} curr={altCurrency} /></span>
                </div>
              </div>

              {/* Optional Monthly Breakdown List for Year View */}
              {summaryMode === 'year' && (
                <div className="mt-6 pt-5 border-t border-slate-100 animate-in fade-in slide-in-from-bottom-2 duration-300">
                  <p className="text-[11px] font-extrabold text-slate-400 uppercase tracking-widest mb-3 flex items-center gap-1.5 justify-center">
                    <List className="w-3.5 h-3.5 text-emerald-600 animate-pulse" />
                    <span>{config.language === 'English' ? "Monthly Breakdown" : "របាយការណ៍លម្អិតតាមខែនីមួយៗ"}</span>
                  </p>
                  <div className="space-y-2 max-h-[350px] overflow-y-auto pr-1">
                    {monthlyBreakdown.map((mb) => (
                      <div 
                        key={mb.monthNum} 
                        className={cn(
                          "flex items-center justify-between p-3 rounded-xl border transition-all duration-200",
                          mb.hasData 
                            ? "bg-slate-50/50 border-slate-100 hover:bg-slate-50 hover:border-slate-200/80" 
                            : "bg-white border-dashed border-slate-100/80 opacity-60"
                        )}
                      >
                        <div className="flex items-center gap-2">
                          <div className={cn(
                            "w-2 h-2 rounded-full",
                            mb.hasData 
                              ? (mb.balance >= 0 ? "bg-emerald-500" : "bg-rose-500") 
                              : "bg-slate-200"
                          )} />
                          <span className="text-xs font-bold text-slate-700">{mb.monthName}</span>
                        </div>
                        
                        <div className="flex items-center gap-5 text-right font-sans">
                          {mb.hasData ? (
                            <>
                              <div className="hidden xs:block">
                                <p className="text-[9px] uppercase font-bold text-slate-400">{t("ចំណូល", config.language)}</p>
                                <p className="text-xs font-bold text-emerald-600">+{formatMoney(mb.income, currency)}</p>
                              </div>
                              <div className="hidden xs:block">
                                <p className="text-[9px] uppercase font-bold text-slate-400">{t("ចំណាយ", config.language)}</p>
                                <p className="text-xs font-bold text-rose-500">-{formatMoney(mb.expense, currency)}</p>
                              </div>
                              <div>
                                <p className="text-[9px] uppercase font-bold text-slate-400">{config.language === 'English' ? "Savings" : "សល់"}</p>
                                <p className={cn(
                                  "text-xs font-extrabold",
                                  mb.balance >= 0 ? "text-emerald-700" : "text-rose-600"
                                )}>
                                  {mb.balance >= 0 ? '+' : ''}{formatMoney(mb.balance, currency)}
                                </p>
                              </div>
                            </>
                          ) : (
                            <span className="text-[10px] font-bold text-slate-400/80 italic">
                              {config.language === 'English' ? "No data" : "គ្មានទិន្នន័យ"}
                            </span>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </motion.div>
          </AnimatePresence>
        </div>
      </motion.div>
    </motion.div>
  );
}

function AddRecordView({ session, categories, onUpdate, showToast, setActiveTab, editRecord, users, config }: any) {
  const [type, setType] = useState<RecordType>(editRecord?.type || 'ចំណាយ');
  const dLocal = new Date();
  dLocal.setMinutes(dLocal.getMinutes() - dLocal.getTimezoneOffset());
  const todayYMD = dLocal.toISOString().split('T')[0];
  
  const [date, setDate] = useState(editRecord?.date || todayYMD);
  const [amount, setAmount] = useState(editRecord?.amount ? String(editRecord.amount) : '');
  const [currency, setCurrency] = useState<'៛' | '$'>(editRecord?.currency || config.currency || '$');
  const [category, setCategory] = useState(editRecord?.category || '');
  const [desc, setDesc] = useState(editRecord?.description || '');
  const [note, setNote] = useState(editRecord?.note || '');
  const [payment, setPayment] = useState(editRecord?.paymentMethod || 'សាច់ប្រាក់');
  const [memberId, setMemberId] = useState(editRecord?.memberId || session.id);
  const [isSaving, setIsSaving] = useState(false);

  const activeCategories = categories.filter((c:any) => c.type === type);

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (isSaving) return;
    if (!amount || isNaN(Number(amount)) || Number(amount) <= 0) {
      showToast(t('ចំនួនទឹកប្រាក់មិនត្រឹមត្រូវទេ', config.language), 'error'); return;
    }
    if (!category) {
      showToast(t('សូមជ្រើសរើសប្រភេទ', config.language), 'error'); return;
    }
    
    setIsSaving(true);
    const d = new Date(date);
    const targetId = editRecord?.id || uuidv4();
    const isEditing = !!editRecord;
    
    const exRate = config.exchangeRate || 4000;
    const numAmount = Number(amount);
    const amountKHR = currency === '៛' ? numAmount : Math.round(numAmount * exRate);
    const amountUSD = currency === '$' ? numAmount : Math.round((numAmount / exRate) * 100) / 100;
    const displayAmount = currency === '$' ? `$${numAmount}` : `${numAmount} ៛`;

    const recordObj: FinanceRecord = {
      id: targetId,
      date,
      month: (() => {
        const parts = date.split('-');
        return parts[1] ? parseInt(parts[1], 10) : (d.getMonth() + 1);
      })(),
      year: (() => {
        const parts = date.split('-');
        return parts[0] ? parseInt(parts[0], 10) : d.getFullYear();
      })(),
      type,
      category,
      description: desc,
      currency,
      amount: numAmount,
      amountKHR,
      amountUSD,
      displayAmount,
      exchangeRate: exRate,
      paymentMethod: payment,
      memberId,
      note,
      createdAt: editRecord?.createdAt || Date.now(),
      updatedAt: Date.now()
    };

    // 1. Always write to Local Storage first for reliable performance & hybrid fallback
    const recs = getRecords();
    if (isEditing) {
      const idx = recs.findIndex((r: any) => r.id === targetId);
      if (idx !== -1) {
        recs[idx] = recordObj;
      }
    } else {
      recs.push(recordObj);
    }
    saveRecords(recs);

    // 2. Routing based on selected sync mode
    const mode = getSyncMode();
    
    // UI update right away for responsiveness
    setAmount('');
    setDesc('');
    await onUpdate();
    setIsSaving(false);
    setActiveTab('dashboard');

    if (mode === 'sync') {
      const syncTask = isEditing
        ? updateRecordInGoogleSheet(recordObj, session.name)
        : syncRecordToGoogleSheet(recordObj, session.name);
        
      syncTask.then(ok => {
        if (ok) {
          showToast(t('រក្សាទុក និង Sync ជោគជ័យ', config.language), 'success');
        } else {
          showToast(t('សមកាលកម្មបរាជ័យ ប៉ុន្តែទិន្នន័យត្រូវបានរក្សាទុកក្នុង Local', config.language), 'error');
        }
      }).catch(() => {
        showToast(t('សមកាលកម្មបរាជ័យ ប៉ុន្តែទិន្នន័យត្រូវបានរក្សាទុកក្នុង Local', config.language), 'error');
      });
    } else if (mode === 'hybrid') {
      if (navigator.onLine) {
        const syncTask = isEditing
          ? updateRecordInGoogleSheet(recordObj, session.name)
          : syncRecordToGoogleSheet(recordObj, session.name);
          
        syncTask.then(ok => {
          if (ok) {
            showToast(t('រក្សាទុកក្នុង Local និង Sync ជោគជ័យ', config.language), 'success');
          } else {
            addToOfflineQueue(isEditing ? 'update' : 'add', targetId, recordObj);
            showToast(t('រក្សាទុកក្នុង Local និងបានដាក់ចូល Queue សម្រាប់ Sync ពេលមាន Internet', config.language), 'info');
          }
        }).catch(() => {
          addToOfflineQueue(isEditing ? 'update' : 'add', targetId, recordObj);
          showToast(t('រក្សាទុកក្នុង Local និងបានដាក់ចូល Queue សម្រាប់ Sync ពេលមាន Internet', config.language), 'info');
        });
      } else {
        addToOfflineQueue(isEditing ? 'update' : 'add', targetId, recordObj);
        showToast(t('រក្សាទុកក្នុង Local និងបានដាក់ចូល Queue សម្រាប់ Sync ពេលមាន Internet', config.language), 'info');
      }
    } else {
      showToast(t('រក្សាទុកបានជោគជ័យ (Local Only)', config.language), 'success');
    }
  };

  return (
    <motion.div initial={{ opacity: 0, scale: 0.95 }} animate={{ opacity: 1, scale: 1 }} exit={{ opacity: 0, scale: 0.95 }} transition={{ duration: 0.2 }} className="p-6 space-y-6 pb-24">
      
      
      <form onSubmit={handleSave} className="bg-white p-6 rounded-[32px] border border-[#E8EEE9] shadow-lg shadow-black/5 space-y-6 relative">
        
        {/* Type Toggle */}
        <div className="flex bg-[#F5F7F5] p-1.5 rounded-[20px]">
          <label className={cn("flex-1 text-center py-3.5 rounded-2xl cursor-pointer transition-all font-bold text-sm", type === 'ចំណូល' ? "bg-emerald-600 shadow-md text-white" : "text-slate-500 hover:text-emerald-700")}>
            <input type="radio" name="entry-type" value="ចំណូល" className="hidden" checked={type === 'ចំណូល'} onChange={() => { setType('ចំណូល'); setCategory(''); }} />
            {config.language === 'English' ? 'Income' : 'ចំណូល'}
          </label>
          <label className={cn("flex-1 text-center py-3.5 rounded-2xl cursor-pointer transition-all font-bold text-sm", type === 'ចំណាយ' ? "bg-red-500 shadow-md text-white" : "text-slate-500 hover:text-red-700")}>
            <input type="radio" name="entry-type" value="ចំណាយ" className="hidden" checked={type === 'ចំណាយ'} onChange={() => { setType('ចំណាយ'); setCategory(''); }} />
            {config.language === 'English' ? 'Expense' : 'ចំណាយ'}
          </label>
        </div>

        {/* Amount */}
        <div>
          <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
            {config.language === 'English' ? 'Amount' : 'ចំនួនទឹកប្រាក់'} <span className="text-red-400">*</span>
          </label>
          <div className="relative">
            <button 
              type="button" 
              onClick={() => setCurrency(currency === '$' ? '៛' : '$')}
              className="absolute left-3 top-1/2 -translate-y-1/2 w-12 h-12 bg-emerald-50 rounded-[18px] text-lg font-black text-emerald-600 flex items-center justify-center hover:bg-emerald-100 transition-colors shadow-sm active:scale-95 z-10"
              title={config.language === 'English' ? "Switch Currency" : "ផ្លាស់ប្ដូររូបិយប័ណ្ណ (Switch Currency)"}
            >
              {currency}
            </button>
            <input type="number" step="0.01" required value={amount} onChange={e => setAmount(e.target.value)} className={cn("w-full bg-[#FBFDFB] border-2 border-[#E8EEE9] pl-20 pr-6 py-5 rounded-[24px] text-2xl font-extrabold outline-none transition-all shadow-sm focus:bg-white text-center tracking-wider", type === 'ចំណូល' ? "text-emerald-600 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 placeholder:text-emerald-200" : "text-red-500 focus:border-red-500 focus:ring-4 focus:ring-red-500/10 placeholder:text-red-200")} placeholder="0.00" />
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
              {config.language === 'English' ? 'Date' : 'កាលបរិច្ឆេទ'} <span className="text-red-400">*</span>
            </label>
            <CustomDatePicker 
              value={date} 
              onChange={setDate} 
              language={config.language} 
              todayYMD={todayYMD} 
            />
          </div>
          <div>
            <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
              {config.language === 'English' ? 'Category' : 'ប្រភេទ'} <span className="text-red-400">*</span>
            </label>
            <CustomSelect 
              value={category} 
              onChange={setCategory}
              placeholder={config.language === 'English' ? '-- Select --' : '-- ជ្រើសរើស --'}
              className="px-4 py-[15px] h-[54px] rounded-[20px] shadow-sm bg-[#FBFDFB] border-2 border-[#E8EEE9] hover:border-emerald-400 text-sm font-bold text-emerald-950"
              modalTitle={config.language === 'English' ? 'Select Category' : 'ជ្រើសរើសប្រភេទ (Select Category)'}
              options={activeCategories.map((c:any) => ({ value: c.name, label: translateCategory(c.name, config.language) }))}
            />
          </div>
        </div>

        <div>
          <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
            {config.language === 'English' ? 'Description' : 'ពិពណ៌នា (Description)'} <span className="text-red-400">*</span>
          </label>
          <input type="text" required placeholder={config.language === 'English' ? 'e.g. Buy rice...' : 'ឧ. ទិញអង្ករ...'} value={desc} onChange={e => setDesc(e.target.value)} className="w-full bg-[#FBFDFB] border-2 border-[#E8EEE9] px-5 py-4 rounded-[20px] text-sm outline-none focus:border-slate-800 transition font-bold text-emerald-950" />
        </div>
        
        <div>
          <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
            {config.language === 'English' ? 'Note (Optional)' : 'ចំណាំ (Note) (មិនចាំបាច់)'}
          </label>
          <input type="text" placeholder={config.language === 'English' ? 'Additional notes...' : 'កំណត់ចំណាំបន្ថែម...'} value={note} onChange={e => setNote(e.target.value)} className="w-full bg-[#FBFDFB] border-2 border-[#E8EEE9] px-5 py-4 rounded-[20px] text-sm outline-none focus:border-slate-800 transition font-bold text-emerald-950" />
        </div>

        <div>
           <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
             {config.language === 'English' ? 'Payment Method' : 'វិធីទូទាត់'} <span className="text-red-400">*</span>
           </label>
           <div className="flex gap-2 bg-[#F5F7F5] p-1.5 rounded-[20px]">
              {['សាច់ប្រាក់', 'ABA', 'ការទូទាត់ផ្សេងៗ'].map(pm => (
                <button type="button" key={pm} onClick={() => setPayment(pm)} className={cn("flex-1 text-xs font-bold py-3.5 rounded-2xl transition", payment === pm ? "bg-white text-emerald-900 shadow-sm" : "text-slate-500 hover:text-slate-800")}>
                  {pm === 'សាច់ប្រាក់' ? (config.language === 'English' ? 'Cash' : 'សាច់ប្រាក់') : pm === 'ការទូទាត់ផ្សេងៗ' ? (config.language === 'English' ? 'Bank' : 'ផ្សេងៗ') : pm}
                </button>
              ))}
           </div>
        </div>

        {session.role === 'Admin' && (
        <div>
           <label className="text-[10px] uppercase font-bold tracking-widest text-slate-400 block mb-2 ml-1">
             {config.language === 'English' ? 'Record on Behalf of Family Member' : 'កត់ត្រាសម្រាប់គ្រួសារជំនួស'} <span className="text-red-400">*</span>
           </label>
           <CustomSelect 
              value={memberId} 
              onChange={setMemberId}
              className="px-4 py-4 rounded-[20px] shadow-none bg-[#FBFDFB]"
              modalTitle={config.language === 'English' ? 'Select Family Member' : 'ជ្រើសរើសសមាជិក (Select Member)'}
              options={users.map((u:any) => ({ value: u.id, label: u.name, description: u.role }))}
           />
        </div>
        )}

        <button disabled={isSaving} type="submit" className={cn("w-full text-white font-bold py-5 rounded-[24px] shadow-xl transition mt-8 flex justify-center items-center gap-2 text-sm", isSaving ? "bg-slate-500 opacity-80" : "bg-slate-900 hover:bg-slate-800 shadow-slate-900/20 active:scale-[0.98]")}>
          {isSaving ? <span className="material-symbols-rounded text-xl animate-spin">sync</span> : <span className="material-symbols-rounded text-xl">add</span>}
          {isSaving 
            ? (config.language === 'English' ? 'Saving...' : 'កំពុងរក្សាទុក...') 
            : (editRecord 
                ? (config.language === 'English' ? 'Update Record' : 'កែប្រែកំណត់ត្រា') 
                : (config.language === 'English' ? 'Create Record' : 'បង្កើតកំណត់ត្រា'))}
        </button>
      </form>
    </motion.div>
  );
}

function RecordsView({ records, session, currency, onUpdate, showToast, users, onEdit, onDeleteTrigger, config, setPromptYearState }: any) {
  const [filterType, setFilterType] = useState('all');
  const d = new Date();
  const [filterMonth, setFilterMonth] = useState(config.defaultMonth || (d.getMonth() + 1));
  const [filterYear, setFilterYear] = useState(config.defaultYear || d.getFullYear());
  const [filterMember, setFilterMember] = useState('all');
  const [searchQuery, setSearchQuery] = useState('');

  const viewableRecords = useMemo(() => session.role === 'Admin' ? records : records.filter((r:any) => r.memberId === session.id), [records, session.role, session.id]);
  
  const filtered = useMemo(() => viewableRecords
    .filter((r:any) => filterType === 'all' || r.type === filterType)
    .filter((r:any) => r.month === filterMonth && r.year === filterYear)
    .filter((r:any) => filterMember === 'all' || r.memberId === filterMember)
    .filter((r:any) => {
      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return (r.description && r.description.toLowerCase().includes(q)) || 
             (r.category && r.category.toLowerCase().includes(q)) || 
             (r.note && r.note.toLowerCase().includes(q));
    })
    .sort((a:any,b:any) => b.createdAt - a.createdAt), [viewableRecords, filterType, filterMonth, filterYear, filterMember, searchQuery]);

  const getMemberName = useCallback((id: string) => users.find((u:any) => u.id === id)?.name || id, [users]);

  const exRate = config?.exchangeRate || 4000;
  const getAmt = useCallback((r: FinanceRecord, targetCurrency: string) => {
    const rCurr = r.currency || currency;
    if (targetCurrency === '$') {
      return r.amountUSD ?? (rCurr === '$' ? r.amount : r.amount / exRate);
    } else {
      return r.amountKHR ?? (rCurr === '៛' ? r.amount : r.amount * exRate);
    }
  }, [currency, exRate]);

  const totalFilteredInc = useMemo(() => filtered.filter((r:any) => r.type==='ចំណូល').reduce((s:number,r:any)=>s+getAmt(r, currency),0), [filtered, getAmt, currency]);
  const totalFilteredExp = useMemo(() => filtered.filter((r:any) => r.type==='ចំណាយ').reduce((s:number,r:any)=>s+getAmt(r, currency),0), [filtered, getAmt, currency]);

  const altCurrency = currency === '$' ? '៛' : '$';
  const getAltAmount = useCallback((amount: number) => currency === '$' ? amount * exRate : amount / exRate, [currency, exRate]);

  return (
    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.3 }} className="p-6 space-y-5 pb-24">
      {/* Header and filters */}
      <div className="bg-white p-5 rounded-[32px] border border-[#E8EEE9] shadow-sm space-y-4 relative z-50">
        <div className="flex bg-[#F5F7F5] p-1.5 rounded-[20px] relative z-50">
          <label className={cn("flex-1 text-center py-3 rounded-2xl cursor-pointer transition-all font-bold text-[11px]", filterType === 'all' ? "bg-slate-800 shadow-md text-white" : "text-slate-500 hover:text-slate-700")}>
            <input type="radio" value="all" className="hidden" checked={filterType === 'all'} onChange={() => setFilterType('all')} />
            {t("ទាំងអស់", config.language)}
          </label>
          <label className={cn("flex-1 text-center py-3 rounded-2xl cursor-pointer transition-all font-bold text-[11px]", filterType === 'ចំណូល' ? "bg-emerald-600 shadow-md text-white" : "text-slate-500 hover:text-emerald-700")}>
            <input type="radio" value="ចំណូល" className="hidden" checked={filterType === 'ចំណូល'} onChange={() => setFilterType('ចំណូល')} />
            {t("ចំណូល", config.language)}
          </label>
          <label className={cn("flex-1 text-center py-3 rounded-2xl cursor-pointer transition-all font-bold text-[11px]", filterType === 'ចំណាយ' ? "bg-red-500 shadow-md text-white" : "text-slate-500 hover:text-red-700")}>
            <input type="radio" value="ចំណាយ" className="hidden" checked={filterType === 'ចំណាយ'} onChange={() => setFilterType('ចំណាយ')} />
            {t("ចំណាយ", config.language)}
          </label>
        </div>

        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
          <input 
            type="text" 
            placeholder={t("ស្វែងរក...", config.language)} 
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full bg-[#F5F7F5] border border-[#E8EEE9] pl-10 pr-4 py-3 rounded-2xl text-[11px] font-bold text-emerald-950 outline-none appearance-none font-sans"
          />
        </div>
        
        <div className="grid grid-cols-3 gap-2 relative z-40">
          {/* Month */}
          <CustomSelect 
            value={filterMonth} 
            onChange={setFilterMonth}
            className="px-3 py-2.5 rounded-xl text-[10px] uppercase shadow-none border-[#E8EEE9] bg-[#FBFDFB]"
            modalTitle={t("ជ្រើសរើសខែ", config.language)}
            options={[1,2,3,4,5,6,7,8,9,10,11,12].map(m => ({ value: m, label: t(KHMER_MONTHS[m - 1], config.language) }))}
          />
          
          {/* Year */}
          <CustomSelect 
            value={filterYear} 
            onChange={(v: any) => { if (v === 'ADD_NEW_YEAR') setPromptYearState(true); else setFilterYear(v); }}
            className="px-3 py-2.5 rounded-xl text-[10px] uppercase shadow-none border-[#E8EEE9] bg-[#FBFDFB]"
            modalTitle={t("ជ្រើសរើសឆ្នាំ", config.language)}
            options={[...generateYearOptions(records, config.customYears).map(opt => ({ ...opt, label: t(opt.label, config.language) })), { value: 'ADD_NEW_YEAR', label: config.language === 'English' ? '+ Add New Year' : '+ បន្ថែមឆ្នាំថ្មី' }]}
          />

          {/* Member */}
          <CustomSelect 
            value={filterMember} 
            onChange={setFilterMember}
            className="px-3 py-2.5 rounded-xl text-[10px] uppercase shadow-none border-[#E8EEE9] bg-[#FBFDFB] text-ellipsis overflow-hidden truncate"
            modalTitle={t("ជ្រើសរើសសមាជិក", config.language)}
            alignRight={true}
            options={[
              { value: 'all', label: t('សមាជិកទាំងអស់', config.language) },
              ...users.map((u: any) => ({ value: u.id, label: u.name }))
            ]}
          />
        </div>

        <div className="flex flex-col sm:flex-row sm:justify-between sm:items-center pt-4 border-t border-[#F0F4F1] gap-3">
          <span className="text-[10px] font-bold text-slate-500 uppercase bg-slate-50 px-3 py-1.5 rounded-lg inline-block w-fit">
            {config.language === 'English' 
              ? `Total of ${filtered.length} results` 
              : `លទ្ធផលសរុបមាន ${filtered.length}`}
          </span>
          <div className="flex flex-col sm:items-end gap-1.5">
            <div className="flex flex-wrap items-center gap-2">
              {(filterType === 'all' || filterType === 'ចំណូល') && (
                <div className="bg-emerald-50 text-emerald-700 px-3 py-1.5 rounded-xl border border-emerald-100 flex flex-col items-end">
                  <span className="text-[11px] font-black">+{formatMoney(totalFilteredInc, currency)}</span>
                  <span className="text-[9px] font-bold opacity-70">~ {formatMoney(getAltAmount(totalFilteredInc), altCurrency)}</span>
                </div>
              )}
              {(filterType === 'all' || filterType === 'ចំណាយ') && (
                <div className="bg-red-50 text-red-600 px-3 py-1.5 rounded-xl border border-red-100 flex flex-col items-end">
                  <span className="text-[11px] font-black">-{formatMoney(totalFilteredExp, currency)}</span>
                  <span className="text-[9px] font-bold opacity-70">~ {formatMoney(getAltAmount(totalFilteredExp), altCurrency)}</span>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="space-y-3">
        {filtered.length === 0 && (
          <div className="text-center py-16 text-slate-400 text-sm font-bold bg-white rounded-[32px] border border-dashed border-[#E8EEE9] p-6 shadow-sm">
            <div className="w-16 h-16 bg-slate-50 text-slate-300 rounded-full flex items-center justify-center mx-auto mb-4">
              <FileText className="w-8 h-8 text-slate-400" />
            </div>
            {config.language === 'English' ? "No transaction records found for this month" : "មិនមានកំណត់ត្រាប្រតិបត្តិការទេនៅក្នុងខែនេះ"}
          </div>
        )}

        {/* Desktop Table View */}
        {filtered.length > 0 && (
          <div className="hidden md:block bg-white rounded-[32px] border border-[#E8EEE9] overflow-hidden shadow-sm">
            <table className="w-full text-left text-xs text-slate-700 font-sans">
              <thead className="bg-[#F5F7F5] text-slate-500 font-extrabold uppercase tracking-widest text-[9px] border-b border-[#E8EEE9]">
                <tr>
                  <th className="px-5 py-4">{config.language === 'English' ? 'Date' : 'កាលបរិច្ឆេទ'}</th>
                  <th className="px-5 py-4">{config.language === 'English' ? 'Type' : 'ប្រភេទ'}</th>
                  <th className="px-5 py-4">{config.language === 'English' ? 'Category & Description' : 'ប្រភេទចំណាយ & ពិពណ៌នា'}</th>
                  <th className="px-5 py-4">{config.language === 'English' ? 'Amount' : 'ចំនួនទឹកប្រាក់'}</th>
                  <th className="px-5 py-4">{config.language === 'English' ? 'Member' : 'សមាជិក'}</th>
                  <th className="px-5 py-4 text-right">{config.language === 'English' ? 'Action' : 'សកម្មភាព'}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#F0F4F1]">
                <AnimatePresence mode="popLayout">
                  {filtered.map((r: any) => {
                    const isInc = r.type === 'ចំណូល';
                    return (
                      <motion.tr 
                        key={r.id} 
                        layout
                        initial={{ opacity: 0, y: -20 }} 
                        animate={{ opacity: 1, y: 0 }} 
                        exit={{ opacity: 0, y: -20, scale: 0.95 }} 
                        transition={{ duration: 0.3, type: "spring", bounce: 0.25 }}
                        className="hover:bg-slate-50/50 transition-colors"
                      >
                        <td className="px-5 py-4 font-bold">{formatDate(r.date)}</td>
                        <td className="px-5 py-4">
                          <span className={cn(
                            "px-2.5 py-1 rounded-full text-[10px] font-extrabold uppercase tracking-wider",
                            isInc ? "bg-emerald-50 text-emerald-700 border border-emerald-100" : "bg-red-50 text-red-700 border border-red-100"
                          )}>
                            {config.language === 'English' ? (isInc ? 'Income' : 'Expense') : r.type}
                          </span>
                        </td>
                        <td className="px-5 py-4">
                          <p className="font-extrabold text-slate-800 text-[13px]">{translateCategory(r.category, config.language)}</p>
                          <p className="text-slate-400 text-[11px] mt-0.5">{r.description || '—'}</p>
                        </td>
                        <td className={cn("px-5 py-4 font-extrabold text-[13px] text-right", isInc ? "text-emerald-700" : "text-red-500")}>
                          <div className="flex flex-col items-end">
                            <span>{isInc ? '+' : '-'}{formatMoney(r.amount, r.currency || currency)}</span>
                            <span className="text-[10px] text-slate-400 font-bold mt-1">~ {formatMoney(r.currency === '$' ? r.amountKHR || r.amount * 4000 : r.amountUSD || r.amount / 4000, r.currency === '$' ? '៛' : '$')}</span>
                          </div>
                        </td>
                        <td className="px-5 py-4 text-slate-500 font-bold">{getMemberName(r.memberId)}</td>
                        <td className="px-5 py-4 text-right font-sans">
                          <div className="flex gap-2 justify-end">
                            <button onClick={() => onEdit(r)} className="bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm p-2 rounded-xl hover:bg-blue-100 active:scale-95 transition-all outline-none animate-none">
                              <span className="material-symbols-rounded text-[14px]">edit</span>
                            </button>
                            <button onClick={() => onDeleteTrigger(r.id)} className="bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm p-2 rounded-xl hover:bg-red-100 active:scale-95 transition-all outline-none animate-none">
                              <span className="material-symbols-rounded text-[14px]">delete</span>
                            </button>
                          </div>
                        </td>
                      </motion.tr>
                    );
                  })}
                </AnimatePresence>
              </tbody>
            </table>
          </div>
        )}

        {/* Mobile View Card */}
        <div className="md:hidden space-y-3">
          <AnimatePresence mode="popLayout">
            {filtered.map((r:any) => {
              const isInc = r.type === 'ចំណូល';
              return (
                <motion.div 
                  key={r.id} 
                  layout
                  initial={{ opacity: 0, y: -20, scale: 0.95 }} 
                  animate={{ opacity: 1, y: 0, scale: 1 }} 
                  exit={{ opacity: 0, y: -20, scale: 0.95 }} 
                  transition={{ duration: 0.3, type: "spring", bounce: 0.25 }}
                  className={cn(
                    "bg-white p-4 rounded-[28px] border flex flex-col hover:border-emerald-200 transition-colors shadow-sm relative",
                    isInc ? "border-l-4 border-l-emerald-600 border-[#E8EEE9]" : "border-l-4 border-l-red-500 border-[#E8EEE9]"
                  )}
                >
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-3">
                      <div className={cn("w-10 h-10 rounded-[14px] flex items-center justify-center font-bold text-xs shadow-sm shadow-black/5 shrink-0 bg-white border border-[#E8EEE9]", isInc ? 'text-emerald-600' : 'text-red-500')}>
                        {isInc ? <span className="material-symbols-rounded text-base">arrow_upward</span> : <span className="material-symbols-rounded text-base">arrow_downward</span>}
                      </div>
                      <div>
                        <p className="font-extrabold text-sm text-slate-800 leading-tight">{translateCategory(r.category, config.language)}</p>
                        <p className="text-slate-400 text-[11px] font-bold tracking-tight shrink-1 max-w-[170px] truncate">{r.description || '—'}</p>
                      </div>
                    </div>
                    <div className="text-right">
                      <p className={cn("font-extrabold text-sm tracking-tight font-sans", isInc ? 'text-emerald-700' : 'text-red-500')}>
                        {isInc ? '+' : '-'}{formatMoney(r.amount, r.currency || currency)}
                      </p>
                      <p className="text-[10px] text-slate-400 font-bold mb-1">
                        ~ {formatMoney(r.currency === '$' ? r.amountKHR || r.amount * 4000 : r.amountUSD || r.amount / 4000, r.currency === '$' ? '៛' : '$')}
                      </p>
                      <span className="text-[9px] text-slate-400 font-bold tracking-wider block">{getMemberName(r.memberId)}</span>
                    </div>
                  </div>

                  <div className="flex justify-between items-center mt-3 pt-3 border-t border-[#F0F4F1]">
                    <div className="flex gap-1 font-sans">
                      <span className="bg-[#F5F7F5] px-2 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider text-slate-500">
                        {config.language === 'English' ? (r.paymentMethod === 'សាច់ប្រាក់' ? 'Cash' : r.paymentMethod === 'ការទូទាត់ផ្សេងៗ' ? 'Bank' : r.paymentMethod) : r.paymentMethod}
                      </span>
                      <span className="bg-[#F5F7F5] px-2 py-0.5 rounded text-[8px] font-bold uppercase tracking-wider text-slate-500">{formatDate(r.date)}</span>
                    </div>

                    <div className="flex gap-1.5 z-10 font-sans">
                      <button onClick={() => onEdit(r)} className="bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm px-3 py-1.5 rounded-xl hover:bg-blue-100 active:scale-95 transition-all outline-none font-extrabold text-[9px] uppercase tracking-wider flex items-center gap-1">
                        <span className="material-symbols-rounded text-[10px]">edit</span> {config.language === 'English' ? 'Edit' : 'កែប្រែ'}
                      </button>
                      <button onClick={() => onDeleteTrigger(r.id)} className="bg-red-55 text-red-650 px-3 py-1.5 rounded-xl hover:bg-red-100 active:scale-95 transition-all outline-none font-extrabold text-[9px] uppercase tracking-wider flex items-center gap-1">
                        <span className="material-symbols-rounded text-[10px]">delete</span> {config.language === 'English' ? 'Delete' : 'លុប'}
                      </button>
                    </div>
                  </div>
                </motion.div>
              );
            })}
          </AnimatePresence>
        </div>
      </div>
    </motion.div>
  );
}

function SettingsView({ session, setActiveTab, config, onUpdate, showToast, records, users, askConfirm, setPromptYearState, setProgressModal }: any) {
  const [openSection, setOpenSection] = useState<string | null>(null);
  const [appsScriptUrl, setAppsScriptUrl] = useState(getSyncUrl() || '');
  const [testingSheet, setTestingSheet] = useState(false);
  const [syncingSheet, setSyncingSheet] = useState(false);
  const [offlineQueueCount, setOfflineQueueCount] = useState(getOfflineQueue().length);

  const handleUpdateConfig = async (key: keyof AppConfig, value: any) => {
    const freshConfig = { ...config, [key]: value };
    const m = await import('./store');
    m.saveConfig(freshConfig);
    onUpdate();

    const syncMode = getSyncMode();
    const syncUrl = getSyncUrl();
    if (syncUrl && (syncMode === 'sync' || syncMode === 'hybrid')) {
      showToast(t('កំពុងរក្សាទុក និងសមកាលកម្មទៅ Google Sheet...', freshConfig.language) || 'កំពុងរក្សាទុក និងសមកាលកម្ម...', 'sync');
      const ok = await syncAppConfigToGoogleSheet(freshConfig);
      if (ok) {
        showToast(t('រក្សាទុកការកំណត់ និងសមកាលកម្មទៅ Google Sheet រួចរាល់!', freshConfig.language) || 'រក្សាទុកការកំណត់ និងសមកាលកម្មរៀបរយ!', 'success');
      } else {
        showToast(t('រក្សាទុកក្នុងម៉ាស៊ីន ប៉ុន្តែសមកាលកម្មទៅ Google Sheet បរាជ័យ!', freshConfig.language) || 'រក្សាទុកក្នុងម៉ាស៊ីន ប៉ុន្តែសមកាលកម្មបរាជ័យ!', 'warning');
      }
    } else {
      showToast(t('រក្សាទុកការកំណត់ជោគជ័យ', freshConfig.language) || 'រក្សាទុកជោគជ័យ', 'success');
    }
  };

  const handleTestSheetConnection = async () => {
    if (!appsScriptUrl) {
      showToast(t('សូមបញ្ចូល API URL ជាមុនសិន!', config.language));
      return;
    }
    setTestingSheet(true);
    try {
      const ok = await testGoogleSheetConnection(appsScriptUrl);
      if (ok) {
        saveSyncUrl(appsScriptUrl);
        showToast(t('ភ្ជាប់ Google Sheet បានជោគជ័យ', config.language));
        onUpdate();
      } else {
        showToast(t('ការតភ្ជាប់ទៅកាន់ Google Sheet បរាជ័យ!', config.language));
      }
    } catch (err: any) {
      alert((config.language === 'English' ? 'Error: ' : 'កំហុស៖ ') + (err.message || (config.language === 'English' ? 'Could not connect' : 'មិនអាចតភ្ជាប់បានឡើយ')));
    } finally {
      setTestingSheet(false);
    }
  };

  const handleManualSyncNow = async () => {
    setSyncingSheet(true);
    try {
      showToast(t('កំពុងសមកាលកម្មទិន្នន័យ...', config.language));
      const [res, settingsOk, usersOk, categoriesOk] = await Promise.all([
        syncLocalDataToGoogleSheet(records, session.name),
        syncAppConfigToGoogleSheet(config),
        syncUsersToGoogleSheet(users),
        syncCategoriesToGoogleSheet(getCategories())
      ]);

      if (res.success && settingsOk && usersOk && categoriesOk) {
        showToast(t('សមកាលកម្មទិន្នន័យ កំណត់ត្រា និងការកំណត់បានជោគជ័យពេញលេញ!', config.language));
        setOfflineQueueCount(0);
        onUpdate();
      } else if (res.success) {
        showToast(t('សមកាលកម្មកំណត់ត្រាជោគជ័យ ប៉ុន្តែការកំណត់មួយចំនួនបរាជ័យ', config.language));
        setOfflineQueueCount(0);
        onUpdate();
      } else {
        showToast(res.error || t('សមកាលកម្មមានបញ្ហាមួយចំនួន', config.language));
      }
    } catch (err: any) {
      showToast(t('ខកខានក្នុងការ Sync៖ ', config.language) + err.message);
    } finally {
      setSyncingSheet(false);
    }
  };

  const handleBackupJson = () => {
    try {
      const dataStr = getBackupJSON();
      const dataUri = 'data:application/json;charset=utf-8,'+ encodeURIComponent(dataStr);
      
      const fileName = `${config.exportFileNameFormat || 'EasyMoney_[Date]'}`.replace('[Date]', new Date().toISOString().split('T')[0]) + '.json';
      const linkElement = document.createElement('a');
      linkElement.setAttribute('href', dataUri);
      linkElement.setAttribute('download', fileName);
      linkElement.click();
      showToast(t('នាំចេញ JSON Backup រួចរាល់', config.language));
    } catch (err) {
      showToast(t('នាំចេញ JSON បរាជ័យ', config.language));
    }
  };

  const handleRestoreJson = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setProgressModal({
      isOpen: true,
      title: config.language === 'English' ? 'Restore System Backup from JSON' : 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
      percent: 0,
      statusText: config.language === 'English' ? 'Reading backup file...' : 'កំពុងអានឯកសារ...',
      subText: file.name
    });

    const reader = new FileReader();
    reader.onload = (evt) => {
      const fileContent = evt.target?.result as string;
      let progress = 0;
      
      const interval = setInterval(() => {
        progress += Math.floor(Math.random() * 15) + 8;
        if (progress >= 100) {
          progress = 100;
          clearInterval(interval);

          const success = restoreFromJSON(fileContent);
          if (success) {
            setProgressModal({
              isOpen: true,
              title: config.language === 'English' ? 'Restore System Backup from JSON' : 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
              percent: 100,
              statusText: config.language === 'English' ? 'System restored successfully!' : 'ការស្ដារទទួលបានជោគជ័យ!',
              subText: config.language === 'English' ? 'All records and settings have been restored.' : 'ទិន្នន័យទាំងអស់ត្រូវបានដំឡើងឡើងវិញពេញលេញ'
            });
            onUpdate();
            showToast(t('ស្តារទិន្នន័យពី JSON ជោគជ័យ កម្មវិធីកំពុងចាប់ផ្ដើមឡើងវិញ...', config.language));
            setTimeout(() => {
              setProgressModal(null);
              window.location.reload();
            }, 1200);
          } else {
            setProgressModal(null);
            alert(config.language === 'English' ? 'Invalid JSON system backup file!' : 'ឯកសារ JSON មិនត្រឹមត្រូវឡើយ');
          }
        } else {
          let statusLabel = config.language === 'English' ? 'Reading backup file...' : 'កំពុងអានឯកសារ...';
          if (progress > 30 && progress < 70) {
            statusLabel = config.language === 'English' ? 'Verifying transaction schema...' : 'កំពុងត្រួតពិនិត្យភាពត្រឹមត្រូវជួរឈរ...';
          } else if (progress >= 70) {
            statusLabel = config.language === 'English' ? 'Restoring filesystems...' : 'កំពុងសរសេរចូលប្រព័ន្ធ Local និងសង្គ្រោះទំហំតុ...';
          }
          setProgressModal({
            isOpen: true,
            title: config.language === 'English' ? 'Restore System Backup from JSON' : 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
            percent: progress,
            statusText: statusLabel,
            subText: file.name
          });
        }
      }, 70);
    };
    reader.readAsText(file);
  };

  const handleResetSettingsOnly = () => {
    askConfirm(t('តើអ្នកពិតជាចង់កំណត់ការកំណត់ (Settings) ទាំងអស់ទៅដើមវិញមែនទេ?', config.language), () =>  {
      import('./store').then(async m => {
        const defaultConfig: AppConfig = {
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
        m.saveConfig(defaultConfig);
        await syncAppConfigToGoogleSheet(defaultConfig);
        onUpdate();
        showToast(t('បានកំណត់ Settings ទៅដើមវិញជោគជ័យ', config.language));
        setTimeout(() => {
          window.location.reload();
        }, 1200);
      });
    });
  };

  const handleClearAllRecords = () => {
    askConfirm(t('⚠️ គ្រោះថ្នាក់៖ តើអ្នកពិតជាចង់លុបប្រតិបត្តិការកត់ត្រាទាំងអស់មែនទេ?', config.language), async () => {
      saveRecords([]);
      onUpdate();
      
      const syncMode = getSyncMode();
      const apiUrl = getSyncUrl();
      if ((syncMode === 'sync' || syncMode === 'hybrid') && apiUrl) {
        showToast(t('កំពុងលុបពី Google Sheet...', config.language));
        const res = await syncLocalDataToGoogleSheet([], session?.name);
        if (res.success) {
          showToast(t('បានលុបកំណត់ត្រាទាំងអស់ (រួមទាំងក្នុង Google Sheet) រួចរាល់', config.language), 'success');
        } else {
          showToast(t('បានលុបក្នុងម៉ាស៊ីន ប៉ុន្តែបំពានសិទ្ធិឯកសារបញ្ជូនទៅ Google Sheet!', config.language), 'error');
        }
      } else {
        showToast(t('បានលុបកំណត់ត្រាទាំងអស់រួចរាល់', config.language), 'success');
      }
    });
  };

  const handleResetAppToDefault = () => {
    askConfirm(t('🚨 គ្រោះថ្នាក់បំផុត៖ តើអ្នកពិតជានឹងលុបទិន្នន័យកម្មវិធីទាំងអស់មែនទេ? កម្មវិធីនឹងចាប់ផ្តើមជាថ្មីទាំងស្រុង!', config.language), () => {

      clearData();
      window.location.reload();
    });
  };

  const handleCurrencyToggle = () => {
    const nextCurr = config.currency === '៛' ? '$' : '៛';
    handleUpdateConfig('currency', nextCurr);
  };

  const toggleSection = (id: string) => {
    setOpenSection(openSection === id ? null : id);
  };

  const isLocked = config.adminSettingsLock && session.role !== 'Admin';

  if (isLocked) {
    return (
      <div className="p-6 text-center max-w-sm mx-auto space-y-6 pt-16 animate-in fade-in">
        <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-full flex items-center justify-center mx-auto shadow-md">
          <Lock className="w-10 h-10" />
        </div>
        <h3 className="text-emerald-950 font-extrabold text-base">ទំព័រនេះត្រូវបានចាក់សោរ</h3>
        <p className="text-xs text-slate-500 font-bold leading-relaxed">មុខងារគ្រប់គ្រងត្រូវបានចាក់សោរសម្រាប់តែសមាជិកដែលមានតួនាទីជា Admin តែប៉ុណ្ណោះ។ សូមទាក់ទងអ្នកគ្រប់គ្រងដើម្បីបើកសិទ្ធិ។</p>
        <button onClick={() => setActiveTab('dashboard')} className="w-full bg-slate-900 text-white font-bold py-3.5 rounded-[20px] text-xs transition active:scale-95 shadow-md">ត្រឡប់ទៅការិយាល័យ</button>
      </div>
    );
  }

  return (
    <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 pb-36 md:pb-28">
      {/* Grid of Admin shortcuts */}
      {session.role === 'Admin' && (
        <div className="grid grid-cols-2 gap-3 mb-2 font-sans">
          <button onClick={() => setActiveTab('members')} className="bg-white hover:bg-slate-50 border border-[#E8EEE9] p-4 rounded-[24px] flex items-center gap-3 shadow-sm transition active:scale-95 text-left outline-none">
            <div className="p-3 rounded-2xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-xl">group</span></div>
            <div>
              <p className="font-extrabold text-xs text-emerald-950">{config.language === 'English' ? "Family Accounts" : "សមាជិកគ្រួសារ"}</p>
              <p className="text-[9px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">{config.language === 'English' ? "Manage Members" : "គណនីសមាជិកគ្រួសារ"}</p>
            </div>
          </button>
          
          <button onClick={() => setActiveTab('categories')} className="bg-white hover:bg-slate-50 border border-[#E8EEE9] p-4 rounded-[24px] flex items-center gap-3 shadow-sm transition active:scale-95 text-left outline-none">
            <div className="p-3 rounded-2xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-xl">sell</span></div>
            <div>
              <p className="font-extrabold text-xs text-emerald-950">{config.language === 'English' ? "Category List" : "ប្រភេទចំណាយ"}</p>
              <p className="text-[9px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">{config.language === 'English' ? "Expense Categories" : "គណនីប្រភេទចំណាយ"}</p>
            </div>
          </button>
        </div>
      )}

      {/* Accordion Wrapper */}
      <div className="space-y-3.5 font-sans relative">
        
        {/* General Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'A' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('A')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px]">settings</span></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">
                {config.language === 'English' ? "General Settings" : "ការកំណត់ទូទៅ"}
              </span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'A' && "rotate-90")}>chevron_right</span>
          </div>
          
          {openSection === 'A' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/25">
              {/* Language */}
              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                  {config.language === 'English' ? "App Language" : "ភាសាកម្មវិធី"}
                </label>
                <CustomSelect 
                  value={config.language || 'Khmer'} 
                  onChange={(v: string) => handleUpdateConfig('language', v)}
                  className="p-3 rounded-xl text-xs font-bold shadow-none"
                  modalTitle={config.language === 'English' ? "Language" : "ភាសា"}
                  options={[
                    { value: 'Khmer', label: config.language === 'English' ? 'Khmer' : 'ភាសាខ្មែរ' },
                    { value: 'English', label: config.language === 'English' ? 'English' : 'ភាសាអង់គ្លេស' }
                  ]}
                />
              </div>

              {/* App Name */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "App Name" : "ឈ្មោះកម្មវិធី"}
                  </label>
                  <input type="text" value={config.appName || 'EasyMoney'} onChange={(e) => handleUpdateConfig('appName', e.target.value)} className="w-full bg-white border border-[#E8EEE9] p-3 rounded-xl text-xs font-bold text-emerald-950 outline-none" placeholder="EasyMoney" />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "App Logo URL" : "រូបសញ្ញា Logo"}
                  </label>
                  <input type="text" value={config.appLogoUrl || ''} onChange={(e) => handleUpdateConfig('appLogoUrl', e.target.value)} className="w-full bg-white border border-[#E8EEE9] p-3 rounded-xl text-xs font-bold text-emerald-950 outline-none" placeholder="https://..." />
                </div>
              </div>

              {/* Currency Selector */}
              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                  {config.language === 'English' ? "Default Currency" : "រូបិយប័ណ្ណបង្ហាញជាចម្បង"}
                </label>
                <div className="flex bg-[#F5F7F5] p-1 rounded-xl">
                  <button type="button" onClick={() => handleUpdateConfig('currency', '៛')} className={cn("flex-1 text-center py-2.5 rounded-lg font-bold text-xs", config.currency === '៛' ? "bg-white text-emerald-600 shadow-sm" : "text-slate-550")}>៛ Khmer Riel</button>
                  <button type="button" onClick={() => handleUpdateConfig('currency', '$')} className={cn("flex-1 text-center py-2.5 rounded-lg font-bold text-xs", config.currency === '$' ? "bg-white text-emerald-600 shadow-sm" : "text-slate-550")}>$ US Dollar</button>
                </div>
              </div>

              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                  {config.language === 'English' ? "Exchange Rate (1 USD = ? KHR)" : "អត្រាប្តូរប្រាក់ (១ ដុល្លារ = ?)"}
                </label>
                <input 
                  type="number" 
                  value={config.exchangeRate || 4000} 
                  onChange={(e) => handleUpdateConfig('exchangeRate', Number(e.target.value))}
                  className="w-full p-3 rounded-xl border border-[#E8EEE9] bg-white font-bold text-xs text-emerald-950 focus:border-emerald-400 focus:ring-4 focus:ring-emerald-500/10 outline-none"
                />
              </div>

              {/* Calendar Defaults */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Default Month" : "ខែលំនាំដើម"}
                  </label>
                  <CustomSelect 
                    value={config.defaultMonth || 1} 
                    onChange={(v: number) => handleUpdateConfig('defaultMonth', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Month" : "ខែ"}
                    options={[1,2,3,4,5,6,7,8,9,10,11,12].map(m => ({ value: m, label: config.language === 'English' ? ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'][m - 1] : KHMER_MONTHS[m - 1] }))}
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Default Year" : "ឆ្នាំលំនាំដើម"}
                  </label>
                  <CustomSelect 
                    value={config.defaultYear || 2026} 
                    onChange={(v: any) => { if (v === 'ADD_NEW_YEAR') setPromptYearState(true); else handleUpdateConfig('defaultYear', v); }}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Year" : "ឆ្នាំ"}
                    options={[...generateYearOptions(records, config.customYears).map(opt => ({ ...opt, label: t(opt.label, config.language) })), { value: 'ADD_NEW_YEAR', label: config.language === 'English' ? '+ Add New Year' : '+ បន្ថែមឆ្នាំថ្មី' }]}
                  />
                </div>
              </div>

              {/* Formats */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Date Format" : "ទម្រង់កាលបរិច្ឆេទ"}
                  </label>
                  <CustomSelect 
                    value={config.dateFormat || 'YYYY-MM-DD'} 
                    onChange={(v: string) => handleUpdateConfig('dateFormat', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Date Format" : "ទម្រង់កាលបរិច្ឆេទ"}
                    options={[
                      { value: 'YYYY-MM-DD', label: 'YYYY-MM-DD' },
                      { value: 'DD/MM/YYYY', label: 'DD/MM/YYYY' },
                      { value: 'MM/DD/YYYY', label: 'MM/DD/YYYY' }
                    ]}
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Number Format" : "ទម្រង់លេខសំគាល់"}
                  </label>
                  <CustomSelect 
                    value={config.numberFormat || 'Comma'} 
                    onChange={(v: string) => handleUpdateConfig('numberFormat', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle="ទ្រង់ទ្រាយលេខ"
                    options={[
                      { value: 'Comma', label: '1,000,000' },
                      { value: 'Plain', label: '1000000' }
                    ]}
                  />
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Theme Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'B' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('B')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px]">palette</span></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">
                {config.language === 'English' ? "Theme Settings" : "ការកំណត់រចនាប័ទ្ម"}
              </span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'B' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'B' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/25">
              {/* Theme Mode & Compact Mode */}
              <div className="grid grid-cols-2 gap-3 mb-4">
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Theme Mode" : "ទម្រង់ពណ៌"}
                  </label>
                  <CustomSelect 
                    value={config.themeMode || 'Light'} 
                    onChange={(v: string) => handleUpdateConfig('themeMode', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Theme Mode" : "ទម្រង់ពណ៌"}
                    options={[
                      { value: 'Light', label: config.language === 'English' ? 'Light Mode' : 'ម៉ូដភ្លឺ' },
                      { value: 'Dark', label: config.language === 'English' ? 'Dark Mode' : 'ម៉ូដងងឹត' },
                      { value: 'System', label: config.language === 'English' ? 'System Default' : 'តាមប្រព័ន្ធ' }
                    ]}
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">
                    {config.language === 'English' ? "Layout Size" : "ទំហំផ្ទៃកម្មវិធី"}
                  </label>
                  <CustomSelect 
                    value={config.compactMode ? 'Compact' : 'Normal'} 
                    onChange={(v: string) => handleUpdateConfig('compactMode', v === 'Compact')}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Layout Size" : "ទំហំទម្រង់"}
                    options={[
                      { value: 'Normal', label: config.language === 'English' ? 'Spacious' : 'ទំហំធម្មតា' },
                      { value: 'Compact', label: config.language === 'English' ? 'Compact' : 'ទំហំតូច' }
                    ]}
                  />
                </div>
              </div>

              {/* Primary Color Palette */}
              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-2 block">
                  {config.language === 'English' ? "Primary Palette" : "ក្ដារពណ៌ចម្បង"}
                </label>
                <div className="grid grid-cols-6 gap-2">
                  {['emerald', 'indigo', 'violet', 'blue', 'rose', 'slate'].map((col) => (
                    <button key={col} type="button" onClick={() => handleUpdateConfig('themeColor', col)} className={cn(
                      "w-full h-10 rounded-xl relative flex items-center justify-center transition-all",
                      col === 'emerald' && 'bg-emerald-600',
                      col === 'indigo' && 'bg-indigo-600',
                      col === 'violet' && 'bg-violet-600',
                      col === 'blue' && 'bg-blue-600',
                      col === 'rose' && 'bg-rose-600',
                      col === 'slate' && 'bg-slate-600'
                    )}>
                      {config.themeColor === col && <span className="material-symbols-rounded text-base text-white animate-scale-up">check</span>}
                    </button>
                  ))}
                </div>
              </div>

              {/* Toggles */}
              <div className="pt-2 border-t border-dashed border-slate-200 mt-4 space-y-1">
                <div className="flex justify-between items-center py-2">
                  <div>
                    <span className="text-xs font-bold text-slate-700 block">
                      {config.language === 'English' ? "Show Shadows" : "បង្ហាញស្រមោល"}
                    </span>
                    <span className="text-[9px] font-medium text-slate-400">
                      {config.language === 'English' ? "Add depth to cards and buttons" : "បន្ថែមជម្រៅលើប៊ូតុង និងកាត"}
                    </span>
                  </div>
                  <button type="button" onClick={() => handleUpdateConfig('showShadows', config.showShadows !== false ? false : true)} className="text-emerald-600 transition hover:scale-105 outline-none">
                    {config.showShadows !== false ? <span className="material-symbols-rounded text-[40px] text-emerald-600">toggle_on</span> : <span className="material-symbols-rounded text-[40px] text-slate-400">toggle_off</span>}
                  </button>
                </div>

                <div className="flex justify-between items-center py-2">
                  <div>
                    <span className="text-xs font-bold text-slate-700 block">
                      {config.language === 'English' ? "Enable UI Animations" : "បើកចលនា UI"}
                    </span>
                    <span className="text-[9px] font-medium text-slate-400">
                      {config.language === 'English' ? "Smooth transitions between views" : "បង្ហាញចលនារលូនពេលប្ដូរផ្ចាំ"}
                    </span>
                  </div>
                  <button type="button" onClick={() => handleUpdateConfig('enableAnimation', config.enableAnimation !== false ? false : true)} className="text-emerald-600 transition hover:scale-105 outline-none">
                    {config.enableAnimation !== false ? <span className="material-symbols-rounded text-[40px] text-emerald-600">toggle_on</span> : <span className="material-symbols-rounded text-[40px] text-slate-400">toggle_off</span>}
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Permissions Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'C' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('C')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm"><Shield className="w-4 h-4"/></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">
                {config.language === 'English' ? "Permissions Settings" : "សិទ្ធិ និងសុវត្ថិភាព"}
              </span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'C' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'C' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-3.5 bg-slate-50/25">
              <div className="flex justify-between items-center py-1.5 border-b border-slate-100 border-dashed">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Members can view all records" : "សមាជិកទាំងអស់អាចមើលកំណត់ត្រាគ្នាបាន"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "If disabled, members only see their own" : "ប្រសិនបើបិទ សមាជិកឃើញតែកំណត់ត្រាខ្លួនឯងប៉ុណ្ណោះ"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('memberViewAll', !config.memberViewAll)} className="text-emerald-600 outline-none">
                  {config.memberViewAll ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>

              <div className="flex justify-between items-center py-1.5 border-b border-slate-100 border-dashed">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Members can edit own records" : "សមាជិកអាចកែប្រែកំណត់ត្រាខ្លួនឯងបាន"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "Allows editing form operations" : "អនុញ្ញាតឱ្យកែប្រែកិច្ចការអាជីវកម្មរបស់ខ្លួនឯង"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('memberEditOwn', !config.memberEditOwn)} className="text-emerald-600 outline-none">
                  {config.memberEditOwn ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>

              <div className="flex justify-between items-center py-1.5 border-b border-slate-100 border-dashed">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Members can delete own records" : "សមាជិកអាចលុបកំណត់ត្រាខ្លួនឯងបាន"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "Allows deletions with confirmation overlay" : "អនុញ្ញាតឱ្យលុបទិន្នន័យដោយមានការផ្ទៀងផ្ទាត់"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('memberDeleteOwn', !config.memberDeleteOwn)} className="text-emerald-600 outline-none">
                  {config.memberDeleteOwn ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>

              <div className="flex justify-between items-center py-1.5 border-b border-slate-100 border-dashed">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Members can export CSV/JSON" : "សមាជិកអាចទាញយកឯកសារលទ្ធផល (Export)"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "Provides access to the Export tab" : "អាចបើកកម្មវិធីទាញយកលទ្ធផល CSV ឬ JSON"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('memberExport', !config.memberExport)} className="text-emerald-600 outline-none">
                  {config.memberExport ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>

              <div className="flex justify-between items-center py-1.5 border-b border-slate-100 border-dashed">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Members can import data" : "សមាជិកអាចបញ្ជូនទិន្នន័យចូលបាន (Import)"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "Unlocks upload options for files" : "អនុញ្ញាតឱ្យបញ្ចូលទិន្នន័យពីក្រៅឧបករណ៍"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('memberImport', !config.memberImport)} className="text-emerald-600 outline-none">
                  {config.memberImport ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>

              <div className="flex justify-between items-center py-1.5">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">
                    {config.language === 'English' ? "Admin Lock settings screen" : "មានតែ Admin ទេ ទើបអាចចូលការកំណត់នេះបាន"}
                  </span>
                  <span className="text-[10px] text-slate-400 font-bold">
                    {config.language === 'English' ? "Locked settings for non-administrators" : "ចាក់សោការកំណត់សម្រាប់សមាជិកធម្មតា"}
                  </span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('adminSettingsLock', !config.adminSettingsLock)} className="text-emerald-600 outline-none">
                  {config.adminSettingsLock ? <ToggleRight className="w-10 h-10 text-emerald-600" /> : <ToggleLeft className="w-10 h-10 text-slate-400" />}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Google Sheet Sync Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'D' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('D')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px]">database</span></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">
                {config.language === 'English' ? "Google Sheet Sync" : "សមកាលកម្ម Google Sheet"}
              </span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'D' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'D' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/25">
              {/* Mode list */}
              <div className="space-y-2.5">
                {[
                  { 
                    value: 'local', 
                    title: config.language === 'English' ? '1. Local Storage Only' : '១. រក្សាទុកក្នុងម៉ាស៊ីន', 
                    desc: config.language === 'English' ? 'Save data locally in your browser storage' : 'រក្សាទុកទិន្នន័យក្នុងម៉ាស៊ីនដោយមិនប្រើអ៊ីនធឺណិត' 
                  },
                  { 
                    value: 'sync', 
                    title: config.language === 'English' ? '2. Auto Sync to Sheets' : '២. សមកាលកម្មទៅ Sheets ចម្បង', 
                    desc: config.language === 'English' ? 'Read and write automatically via Google Web Apps Script API' : 'សរសេរ និងអានដោយស្វ័យប្រវត្តិតាមរយៈ Google Web Apps Script API' 
                  },
                  { 
                    value: 'hybrid', 
                    title: config.language === 'English' ? '3. Hybrid Sync' : '៣. សមកាលកម្មតាមលក្ខណៈគ្រោងទុក (Hybrid)', 
                    desc: config.language === 'English' ? 'Save locally first, then sync to cloud once connection is online' : 'រក្សាទុកក្នុងម៉ាស៊ីនរង់ចាំបញ្ជូនទៅ API ពេលឧបករណ៍មានអ៊ីនធឺណិត' 
                  }
                ].map((syncItem) => {
                  const mode = getSyncMode();
                  return (
                    <label key={syncItem.value} className={cn(
                      "flex items-start gap-3 p-3.5 rounded-xl border cursor-pointer transition-all",
                      mode === syncItem.value ? "border-emerald-500 bg-emerald-50/15" : "border-slate-100 bg-white"
                    )}>
                      <input 
                        type="radio" 
                        value={syncItem.value} 
                        checked={mode === syncItem.value} 
                        onChange={() => {
                          saveSyncMode(syncItem.value as any);
                          onUpdate();
                          showToast(config.language === 'English' ? "Sync mode changed successfully!" : "ប្ដូររបៀបសមកាលកម្មជោគជ័យ!", 'success');
                        }} 
                        className="mt-1 accent-emerald-600" 
                      />
                      <div>
                        <p className="font-extrabold text-xs text-slate-800 leading-tight">{syncItem.title}</p>
                        <p className="text-[10px] text-slate-400 mt-0.5 leading-relaxed">{syncItem.desc}</p>
                      </div>
                    </label>
                  );
                })}
              </div>

              {/* URL field */}
              {getSyncMode() !== 'local' && (
                <div className="space-y-3.5 pt-1">
                  <label className="text-[10px] uppercase font-bold text-slate-500 block">Google Web App URL</label>
                  <textarea 
                    rows={2} 
                    value={appsScriptUrl} 
                    onChange={(e) => {
                      const newUrl = e.target.value;
                      setAppsScriptUrl(newUrl);
                      saveSyncUrl(newUrl);
                    }}
                    onBlur={() => {
                      onUpdate();
                    }}
                    className="w-full bg-white border border-[#E8EEE9] p-3 rounded-xl text-xs font-mono select-all outline-none" 
                    placeholder="https://script.google.com/macros/s/.../exec" 
                  />
                  
                  <div className="flex gap-2">
                    <button type="button" onClick={handleTestSheetConnection} disabled={testingSheet} className="flex-1 bg-white border-2 border-emerald-100 text-emerald-800 text-[10px] py-2.5 rounded-xl font-bold flex items-center justify-center gap-1">
                      <span className="material-symbols-rounded text-[14px]">wifi</span> {testingSheet ? (config.language === 'English' ? 'Testing...' : 'កំពុងសាកល្បង...') : (config.language === 'English' ? 'Test Connection' : 'តេស្តតភ្ជាប់')}
                    </button>
                    <button type="button" onClick={handleManualSyncNow} disabled={syncingSheet} className="flex-1 bg-slate-900 text-white text-[10px] py-2.5 rounded-xl font-extrabold flex items-center justify-center gap-1">
                      <span className={cn("material-symbols-rounded text-[14px]", syncingSheet && "animate-spin")}>sync</span> {syncingSheet ? (config.language === 'English' ? 'Syncing...' : 'កំពុងសមកាល...') : (config.language === 'English' ? 'Sync Now' : 'សមកាលកម្មទិន្នន័យ')}
                    </button>
                  </div>

                  <div className="flex justify-between items-center text-[10px] text-slate-400 border-t border-slate-100 border-dashed pt-3 pb-1">
                    <span>{config.language === 'English' ? `Last sync: ${getLastSyncTime()}` : `សមកាលកម្មចុងក្រោយ៖ ${getLastSyncTime()}`}</span>
                    <span className="font-bold">{config.language === 'English' ? `Offline records: ${offlineQueueCount}` : `កំណត់ត្រារង់ចាំបញ្ជូន៖ ${offlineQueueCount}`}</span>
                  </div>

                  {/* Step-by-step Setup Guide */}
                  <div className="bg-emerald-50/25 border border-emerald-100/60 p-4 rounded-2xl space-y-2 mt-3 text-slate-700 leading-relaxed text-[11px]">
                    <div className="font-extrabold text-xs text-emerald-900 flex items-center gap-1.5 border-b border-emerald-100/40 pb-1.5">
                      <span className="material-symbols-rounded text-sm">info</span>
                      <span>
                        {config.language === 'English' ? "How to Setup Google Sheet Sync" : "របៀបរៀបចំសមកាលកម្ម Google Sheet"}
                      </span>
                    </div>
                    <ol className="list-decimal pl-4.5 space-y-1 my-0 font-medium text-slate-600">
                      <li>
                        {config.language === 'English' 
                          ? "Copy the Apps Script code from /Code.gs" 
                          : "ចម្លងកូដក្នុងឯកសារ /Code.gs"}
                      </li>
                      <li>
                        {config.language === 'English' 
                          ? "Go to Google Sheets -> Extensions -> Apps Script" 
                          : "ទៅកាន់ Google Sheets -> Extensions -> Apps Script"}
                      </li>
                      <li>
                        {config.language === 'English' 
                          ? "Paste the code, save, then click Deploy -> New Deployment" 
                          : "ផាសកូដដែលបានចម្លង រួចចុច Save រួចចុច Deploy -> New Deployment"}
                      </li>
                      <li>
                        {config.language === 'English' 
                          ? "Configure Execute as: 'Me', and Who has access: 'Anyone'. Click Deploy." 
                          : "កំណត់ Execute as: 'Me' និង Who has access: 'Anyone' រួចចុច Deploy"}
                      </li>
                      <li>
                        {config.language === 'English' 
                          ? "Authorize access, copy the generated Web App URL and paste it in the field above!" 
                          : "ផ្ដល់សិទ្ធិ (Authorize Access) រួចចម្លង Web App URL យកមកបិទភ្ជាប់ក្នុងប្រអប់ខាងលើ!"}
                      </li>
                    </ol>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Import Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'E' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('E')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px]">arrow_downward</span></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">{config.language === 'English' ? "Import Settings" : "ការកំណត់នាំចូល"}</span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'E' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'E' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/25">
              {/* Type & mapping preset */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">{config.language === 'English' ? "Default File Type" : "ប្រភេទឯកសារលំនាំដើម"}</label>
                  <CustomSelect 
                    value={config.importType || 'CSV'} 
                    onChange={(v: string) => handleUpdateConfig('importType', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "File Type" : "ប្រភេទឯកសារ (File Type)"}
                    options={[
                      { value: 'CSV', label: config.language === 'English' ? 'Comma Separated (CSV)' : 'ឯកសារ CSV (ក្បៀស)' },
                      { value: 'XLSX', label: config.language === 'English' ? 'Excel Spreadsheet (XLSX)' : 'ឯកសារ Excel (XLSX)' }
                    ]}
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">{config.language === 'English' ? "Mapping Preset" : "ទម្រង់កំណត់ទិន្នន័យ"}</label>
                  <CustomSelect 
                    value={config.importMappingPreset || 'Default'} 
                    onChange={(v: string) => handleUpdateConfig('importMappingPreset', v)}
                    className="p-3 rounded-xl text-xs font-bold shadow-none"
                    modalTitle={config.language === 'English' ? "Mapping Preset" : "គំរូទិន្នន័យ (Mapping Preset)"}
                    options={[
                      { value: 'Default', label: config.language === 'English' ? 'Khmer / English Auto Detect' : 'ស្វែងរក ខ្មែរ/អង់គ្លេស ស្វ័យប្រវត្តិ' },
                      { value: 'Custom', label: config.language === 'English' ? 'Custom Mapping Rule' : 'ក្បួនផែនទីកំណត់ដោយខ្លួនឯង' }
                    ]}
                  />
                </div>
              </div>

              {/* Row Limits */}
              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">{config.language === 'English' ? "Preview Max Rows limit" : "ដែនកំណត់ជួរទិន្នន័យមើលមុន"}</label>
                <CustomSelect 
                  value={config.importPreviewLimit || 50} 
                  onChange={(v: number) => handleUpdateConfig('importPreviewLimit', v)}
                  className="p-3 rounded-xl text-xs font-bold shadow-none"
                  modalTitle={config.language === 'English' ? "Preview Row Limit" : "ដែនកំណត់ការមើលមុន"}
                  options={[
                    { value: 10, label: config.language === 'English' ? '10 Rows preview' : 'បង្ហាញ ១០ ជួរដំបូង' },
                    { value: 25, label: config.language === 'English' ? '25 Rows preview' : 'បង្ហាញ ២៥ ជួរដំបូង' },
                    { value: 50, label: config.language === 'English' ? '50 Rows preview' : 'បង្ហាញ ៥០ ជួរដំបូង' },
                    { value: 100, label: config.language === 'English' ? '100 Rows preview' : 'បង្ហាញ ១០០ ជួរដំបូង' }
                  ]}
                />
              </div>

              {/* Toggles */}
              <div className="flex justify-between items-center py-1">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">{config.language === 'English' ? "Auto assign current user role if missing" : "បំពេញគណនីបច្ចុប្បន្នបើខ្វះព័ត៌មាន"}</span>
                  <span className="text-[10px] text-slate-400 font-bold">{config.language === 'English' ? "Fallback family user to current account" : "កំណត់សមាជិកលំនាំដើមទៅគណនីកំពុងប្រើ"}</span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('autoAssignMissingMember', !config.autoAssignMissingMember)} className="text-emerald-600 outline-none">
                  {config.autoAssignMissingMember ? <span className="material-symbols-rounded text-[40px] text-emerald-600">toggle_on</span> : <span className="material-symbols-rounded text-[40px] text-slate-400">toggle_off</span>}
                </button>
              </div>

              <div className="flex justify-between items-center py-1">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">{config.language === 'English' ? "Auto detect Income/Expense Category" : "ស្វែងរកប្រភេទចំណូល/ចំណាយដោយស្វ័យប្រវត្ត"}</span>
                  <span className="text-[10px] text-slate-400 font-bold">{config.language === 'English' ? "Analyze rows category and deduce values" : "វិភាគជួរប្រភេទចំណូល/ចំណាយ ដើម្បីទាញយកតម្លៃ"}</span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('autoDetectType', !config.autoDetectType)} className="text-emerald-600 outline-none">
                  {config.autoDetectType ? <span className="material-symbols-rounded text-[40px] text-emerald-600">toggle_on</span> : <span className="material-symbols-rounded text-[40px] text-slate-400">toggle_off</span>}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Export Settings */}
        <div className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'F' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('F')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px]">arrow_upward</span></div>
              <span className="font-extrabold text-sm text-slate-800 uppercase tracking-wider">{config.language === 'English' ? "Export Settings" : "ការកំណត់នាំចេញ"}</span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'F' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'F' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/25">
              {/* File Name Format */}
              <div>
                <label className="text-[10px] uppercase font-bold text-slate-500 mb-1.5 block">{config.language === 'English' ? "File Name Output structure" : "ទម្រង់ឈ្មោះឯកសារនាំចេញ"}</label>
                <input type="text" value={config.exportFileNameFormat || 'EasyMoney_[Date]'} onChange={(e) => handleUpdateConfig('exportFileNameFormat', e.target.value)} className="w-full bg-white border border-[#E8EEE9] p-3 rounded-xl text-xs font-mono outline-none text-slate-800" placeholder="EasyMoney_[Date]" />
                <p className="text-[9px] text-slate-400 mt-1">{config.language === 'English' ? "* Export Preview: " : "* គំរូ នាំចេញ: "}<span className="font-bold text-slate-500">{config.exportFileNameFormat?.replace('[Date]', new Date().toISOString().split('T')[0]) || 'EasyMoney'}.csv</span></p>
              </div>

              {/* Toggles */}
              <div className="flex justify-between items-center py-1 border-b border-dashed border-slate-100 pb-3">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 block">{config.language === 'English' ? "Include Header Row in spreadsheet" : "រួមបញ្ចូលចំណងជើងជួរឈរ (Header)"}</span>
                  <span className="text-[10px] text-slate-400 font-bold">{config.language === 'English' ? "Write column labels on line 1" : "សរសេរឈ្មោះជួរឈរនៅលើបន្ទាត់ទី១"}</span>
                </div>
                <button type="button" onClick={() => handleUpdateConfig('exportIncludeHeader', !config.exportIncludeHeader)} className="text-emerald-600 outline-none">
                  {config.exportIncludeHeader ? <span className="material-symbols-rounded text-[40px] text-emerald-600">toggle_on</span> : <span className="material-symbols-rounded text-[40px] text-slate-400">toggle_off</span>}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* G. Data Management */}
        <div id="settings-reset-section" className={cn("bg-white rounded-[24px] border border-[#E8EEE9] shadow-sm transition-all", openSection === 'G' ? "relative z-[30]" : "relative z-[10]")}>
          <div onClick={() => toggleSection('G')} className="p-4 flex items-center justify-between cursor-pointer hover:bg-slate-50 transition-colors">
            <div className="flex items-center gap-3">
              <div className="p-2.5 rounded-xl bg-red-50 text-red-600 border border-red-100 shadow-sm flex items-center justify-center"><span className="material-symbols-rounded text-[18px] text-red-600">warning</span></div>
              <span className="font-extrabold text-sm text-red-600 uppercase tracking-widest">{config.language === 'English' ? "SYSTEM RESET" : "លុបប្រព័ន្ធឡើងវិញ"}</span>
            </div>
            <span className={cn("material-symbols-rounded text-[18px] text-slate-400 transition-transform", openSection === 'G' && "rotate-90")}>chevron_right</span>
          </div>

          {openSection === 'G' && (
            <div className="p-5 border-t border-[#F0F4F1] space-y-4 bg-slate-50/30">
              {/* Backup & restore shortcuts */}
              <div className="grid grid-cols-2 gap-3 pb-2">
                <button id="backup-json-btn" type="button" onClick={handleBackupJson} className="bg-white hover:bg-emerald-50 border border-[#E8EEE9] hover:border-emerald-200 p-4 rounded-[20px] font-extrabold text-xs text-slate-700 flex flex-col items-center justify-center text-center gap-2 shadow-sm transition-all group outline-none">
                  <div className="w-10 h-10 rounded-full bg-emerald-100/50 flex items-center justify-center group-hover:scale-110 transition-transform">
                    <span className="material-symbols-rounded text-[20px] text-emerald-600">database</span>
                  </div>
                  <span>{config.language === 'English' ? 'Download Backup (JSON)' : 'ទាញយក Backup (JSON)'}</span>
                </button>
                <label id="restore-json-label" className="bg-white hover:bg-blue-50 border border-[#E8EEE9] hover:border-blue-200 p-4 rounded-[20px] font-extrabold text-xs text-slate-700 flex flex-col items-center justify-center text-center gap-2 shadow-sm transition-all cursor-pointer group">
                  <div className="w-10 h-10 rounded-full bg-blue-100/50 flex items-center justify-center group-hover:scale-110 transition-transform">
                    <span className="material-symbols-rounded text-[20px] text-blue-600">upload</span>
                  </div>
                  <span>{config.language === 'English' ? 'Restore Backup (JSON)' : 'ស្តារទិន្នន័យ (JSON)'}</span>
                  <input type="file" accept=".json" onChange={handleRestoreJson} className="hidden" />
                </label>
              </div>

              {/* Dangerous actions */}
              <div className="space-y-3">
                <div id="reset-settings-container" className="bg-orange-50/50 border border-orange-100 rounded-[24px] overflow-hidden shadow-sm">
                  <button id="reset-settings-btn" type="button" onClick={handleResetSettingsOnly} className="w-full text-left px-5 py-4 flex items-center justify-between hover:bg-orange-100/30 transition-colors active:bg-orange-100 outline-none">
                    <div className="flex items-start gap-3.5">
                      <div className="p-2.5 rounded-xl bg-orange-100 text-orange-600 shrink-0 mt-0.5"><span className="material-symbols-rounded text-[20px]">settings_backup_restore</span></div>
                      <div>
                        <div className="text-xs font-extrabold text-orange-950">
                          {config.language === 'English' ? "Reset Settings Only" : "កំណត់ការកំណត់ឡើងវិញ"}
                        </div>
                        <div className="text-[10px] text-orange-700/80 font-bold leading-relaxed mt-0.5">
                          {config.language === 'English' 
                            ? "Reset colors, formats, and program structures to default (retains transaction history)" 
                            : "កំណត់ពណ៌ ទ្រង់ទ្រាយ និងរាល់រចនាសម្ព័ន្ធកម្មវិធីទៅដើមវិញ (រក្សាទុកទិន្នន័យប្រតិបត្តិការ)"}
                        </div>
                      </div>
                    </div>
                  </button>
                </div>

                <div id="clear-records-container" className="bg-red-50/30 border border-red-100 rounded-[24px] overflow-hidden shadow-sm">
                  <button id="clear-records-btn" type="button" onClick={handleClearAllRecords} className="w-full text-left px-5 py-4 flex items-center justify-between hover:bg-red-50 transition-colors active:bg-red-100 outline-none">
                    <div className="flex items-start gap-3.5">
                      <div className="p-2.5 rounded-xl bg-red-100 text-red-600 shrink-0 mt-0.5"><span className="material-symbols-rounded text-[20px]">receipt_long</span></div>
                      <div>
                        <div className="text-xs font-extrabold text-red-950">
                          {config.language === 'English' ? "Clear All Transactions" : "លុបប្រតិបត្តិការកត់ត្រាទាំងអស់"}
                        </div>
                        <div className="text-[10px] text-red-700/80 font-bold leading-relaxed mt-0.5">
                          {config.language === 'English' 
                            ? "Clear all income and expense history from Local and Cloud sync (retains settings)" 
                            : "លុបរាល់ព័ត៌មានចំណូល និងចំណាយទាំងអស់ចេញពីប្រព័ន្ធ Local នឹង Cloud (រក្សារាល់ការកំណត់)"}
                        </div>
                      </div>
                    </div>
                  </button>
                </div>

                <div id="factory-reset-container" className="bg-red-600 rounded-[24px] overflow-hidden shadow-md shadow-red-200/55">
                  <button id="factory-reset-btn" type="button" onClick={handleResetAppToDefault} className="w-full text-left px-5 py-4 flex items-center justify-between hover:bg-red-700 transition-colors active:bg-red-800 outline-none">
                    <div className="flex items-start gap-3.5">
                      <div className="p-2.5 rounded-xl bg-white/20 text-white shrink-0 mt-0.5"><span className="material-symbols-rounded text-[20px]">delete_forever</span></div>
                      <div>
                        <div className="text-xs font-extrabold text-white">
                          {config.language === 'English' ? "Factory Reset System" : "លុបសម្អាតប្រព័ន្ធទាំងមូល"}
                        </div>
                        <div className="text-[10px] text-red-100 font-bold leading-relaxed mt-0.5">
                          {config.language === 'English' 
                            ? "Delete all transaction data, history, members, and configuration settings to start fresh" 
                            : "លុបរាល់ទិន្នន័យ ប្រវត្តិប្រតិបត្តិការ សមាជិក និងការកំណត់ទាំងអស់ទាំងស្រុងដើម្បីចាប់ផ្តើមសារជាថ្មី"}
                        </div>
                      </div>
                    </div>
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>

      </div>
    </div>
  );
}

function ImportView({ session, onUpdate, showToast, setActiveTab, config, users, setProgressModal }: any) {
  const [fileData, setFileData] = useState<{ headers: string[]; rows: any[] } | null>(null);
  const [mappings, setMappings] = useState<Record<string, string>>({});
  const [duplicateCheck, setDuplicateCheck] = useState<'DateAmountDesc' | 'ID'>('DateAmountDesc');
  const [isXlsx, setIsXlsx] = useState(false);
  const [importPreviewLimit, setImportPreviewLimit] = useState(config.importPreviewLimit || 50);
  const [finalRecords, setFinalRecords] = useState<FinanceRecord[]>([]);
  const [duplicatesSkipped, setDuplicatesSkipped] = useState(0);
  const [step, setStep] = useState<'upload' | 'map' | 'preview'>('upload');
  const [isEasyMoneyFormat, setIsEasyMoneyFormat] = useState(false);
  const [isDragOver, setIsDragOver] = useState(false);

  const targetFields = [
    { key: 'date', label: 'កាលបរិច្ឆេទ (Date)', khmerTips: 'ថ្ងៃខែឆ្នាំ, កាលបរិច្ឆេទ, ថ្ងៃ', required: true },
    { key: 'type', label: 'ប្រភេទ (Type)', khmerTips: 'ប្រភេទ, ចំណូល/ចំណាយ', required: true },
    { key: 'category', label: 'ប្រភេទចំណាយ (Category)', khmerTips: 'ប្រភេទចំណាយ, ប្រភេទចំណូល, ក្រុម', required: true },
    { key: 'description', label: 'ពិពណ៌នា (Description)', khmerTips: 'ពិពណ៌នា, ការពិពណ៌នា, ព័ត៌មានលម្អិត', required: false },
    { key: 'amount', label: 'ចំនួនទឹកប្រាក់ (Amount)', khmerTips: 'ចំនួនទឹកប្រាក់, ចំនួន, ទឹកប្រាក់, តម្លៃ', required: true },
    { key: 'paymentMethod', label: 'វិធីបង់ប្រាក់ (Payment Method)', khmerTips: 'វិធីបង់ប្រាក់, ទម្រង់ទូទាត់, ប្រភេទបង់ប្រាក់', required: false }
  ];

  // Helper patterns to auto-detect Khmer and English headers
  const autoDetectMap: Record<string, string[]> = {
    date: ['date', 'ថ្ងៃខែឆ្នាំ', 'កាលបរិច្ឆេទ', 'ថ្ងៃ', 'datetime', 'time'],
    type: ['type', 'ប្រភេទ', 'ចំណូល/ចំណាយ', 'ប្រតិបត្តិការ', 'transaction_type'],
    category: ['category', 'ប្រភេទចំណាយ', 'ប្រភេទចំណូល', 'ក្រុម', 'ប្រភេទប្រាក់', 'group'],
    description: ['description', 'ពិពណ៌នា', 'ការពិពណ៌នា', 'ព័ត៌មានលម្អិត', 'កំណត់សម្គាល់', 'memo', 'note'],
    amount: ['amount', 'ចំនួនទឹកប្រាក់', 'ចំនួន', 'ទឹកប្រាក់', 'តម្លៃ', 'ប្រាក់', 'value', 'price'],
    paymentMethod: ['paymentmethod', 'វិធីបង់ប្រាក់', 'ទម្រង់ទូទាត់', 'ប្រភេទបង់ប្រាក់', 'payment', 'method', 'គណនី']
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const processFile = (file: File) => {
    if (file.name.endsWith('.json')) {
      setProgressModal({
        isOpen: true,
        title: 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
        percent: 0,
        statusText: 'កំពុងអានឯកសារ...',
        subText: file.name
      });

      const reader = new FileReader();
      reader.onload = (event) => {
        const fileContent = event.target?.result as string;
        let progress = 0;
        
        const interval = setInterval(() => {
          progress += Math.floor(Math.random() * 15) + 8;
          if (progress >= 100) {
            progress = 100;
            clearInterval(interval);

            const success = restoreFromJSON(fileContent);
            if (success) {
              setProgressModal({
                isOpen: true,
                title: config.language === 'English' ? 'Restore System Backup from JSON' : 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
                percent: 100,
                statusText: config.language === 'English' ? 'System restored successfully!' : 'ការស្ដារទទួលបានជោគជ័យ!',
                subText: config.language === 'English' ? 'All records and settings have been restored.' : 'ទិន្នន័យទាំងអស់ត្រូវបានដំឡើងឡើងវិញពេញលេញ'
              });
              onUpdate();
              showToast(t('ស្តារទិន្នន័យពី JSON ជោគជ័យ កម្មវិធីកំពុងចាប់ផ្ដើមឡើងវិញ...', config.language));
              setActiveTab('dashboard');
              setTimeout(() => {
                setProgressModal(null);
                window.location.reload();
              }, 1200);
            } else {
              setProgressModal(null);
              alert(config.language === 'English' ? 'Invalid JSON system backup file!' : 'ឯកសារ JSON មិនត្រឹមត្រូវឡើយ');
            }
          } else {
            let statusLabel = config.language === 'English' ? 'Parsing JSON backup...' : 'កំពុងវិភាគ JSON Backup...';
            if (progress > 40 && progress < 80) statusLabel = config.language === 'English' ? 'Verifying authentication and accounts...' : 'កំពុងផ្ទៀងផ្ទាត់សិទ្ធិ និងគណនី...';
            else if (progress >= 80) statusLabel = config.language === 'English' ? 'Writing datasets locally...' : 'កំពុងសរសេរទិន្នន័យចូល Local...';

            setProgressModal({
              isOpen: true,
              title: config.language === 'English' ? 'Restore System Backup from JSON' : 'ស្តារទិន្នន័យពី JSON (Restore System Backup)',
              percent: progress,
              statusText: statusLabel,
              subText: file.name
            });
          }
        }, 70);
      };
      reader.readAsText(file);
      return;
    }

    const reader = new FileReader();
    const extension = file.name.split('.').pop()?.toLowerCase();
    
    if (extension === 'xlsx' || extension === 'xls') {
      setIsXlsx(true);
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          
          let isEasyMoneyExcel = false;
          for (let name of workbook.SheetNames) {
            const ws = workbook.Sheets[name];
            if (!ws || !ws['!ref']) continue;
            const sheetRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
            if (sheetRows.length > 0) {
              const headers = (sheetRows[0] as any[]).map(h => String(h || '').trim());
              if (
                (headers.includes('ប្រភេទ') && headers.includes('ចំនួនលុយ')) ||
                (headers.includes('Type') && headers.includes('Amount'))
              ) {
                isEasyMoneyExcel = true;
                break;
              }
            }
          }

          if (isEasyMoneyExcel) {
            const parsed = parseEasyMoneyExcelWorkbook(workbook, users, config, session);
            setFinalRecords(parsed.records);
            setDuplicatesSkipped(parsed.duplicates);
            
            // Generate nice preview rows from the most relevant sheet (either EasyMoney_View, គ្រប់កំណត់ត្រា or first sheet)
            const previewRows: any[] = [];
            const previewSheetName = workbook.SheetNames.find(n => n === 'EasyMoney_View' || n === 'គ្រប់កំណត់ត្រា') || workbook.SheetNames[0];
            const ws = workbook.Sheets[previewSheetName];
            if (ws && ws['!ref']) {
              const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1 });
              if (jsonData.length > 0) {
                const headers = (jsonData[0] as any[]).map(h => String(h || '').trim());
                for (let i = 1; i < jsonData.length; i++) {
                  const rowArr = jsonData[i] as any[];
                  if (!rowArr || rowArr.length === 0) continue;
                  const obj: Record<string, any> = {};
                  headers.forEach((h, idx) => {
                    obj[h] = rowArr[idx] !== undefined ? rowArr[idx] : '';
                  });
                  previewRows.push(obj);
                }
                setFileData({ headers, rows: previewRows });
              }
            }
            
            setIsEasyMoneyFormat(true);
            setStep('preview');
            return;
          }

          const firstSheet = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheet];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          if (jsonData.length === 0) {
            alert('ឯកសារ Excel នេះមិនសំបូរទិន្នន័យឡើយ!');
            return;
          }

          const headers = (jsonData[0] as any[]).map(h => String(h || '').trim());
          const rowsObjects: any[] = [];
          for (let i = 1; i < jsonData.length; i++) {
            const rowArr = jsonData[i] as any[];
            if (!rowArr || rowArr.length === 0) continue;
            const obj: Record<string, any> = {};
            headers.forEach((h, idx) => {
              obj[h] = rowArr[idx] !== undefined ? rowArr[idx] : '';
            });
            rowsObjects.push(obj);
          }

          setFileData({ headers, rows: rowsObjects });
          performAutoMapping(headers, rowsObjects);
          setStep('map');
        } catch (err: any) {
          alert('កំហុសក្នុងការអាន Excel: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setIsXlsx(false);
      reader.onload = (e) => {
        try {
          const text = e.target?.result as string;
          const rows: any[] = [];
          const lines = text.split(/\r?\n/).filter(l => l.trim());
          if (lines.length === 0) {
            alert('ឯកសារ CSV គ្មានទិន្នន័យឡើយ!');
            return;
          }

          // Parse CSV with quoted logic
          const headersLine = lines[0];
          const parseCSVLine = (lineStr: string) => {
            const result: string[] = [];
            let current = '';
            let inQuotes = false;
            for (let j = 0; j < lineStr.length; j++) {
              let char = lineStr[j];
              if (inQuotes) {
                if (char === '"') {
                  if (j + 1 < lineStr.length && lineStr[j + 1] === '"') {
                    current += '"';
                    j++;
                  } else {
                    inQuotes = false;
                  }
                } else {
                  current += char;
                }
              } else {
                if (char === '"') {
                  inQuotes = true;
                } else if (char === ',') {
                  result.push(current.trim());
                  current = '';
                } else {
                  current += char;
                }
              }
            }
            result.push(current.trim());
            return result;
          };

          const headers = parseCSVLine(headersLine).map(h => h.replace(/^"|"$/g, '').trim());
          
          for (let i = 1; i < lines.length; i++) {
            const rowArr = parseCSVLine(lines[i]);
            const obj: Record<string, any> = {};
            headers.forEach((h, idx) => {
              obj[h] = rowArr[idx] !== undefined ? rowArr[idx] : '';
            });
            rows.push(obj);
          }

          if (headers.includes('ID') && headers.includes('Display Amount') && headers.includes('Payment Method') && headers.includes('Created At')) {
            const parsed = parseEasyMoneyExchangeFormat(rows, users, config, session);
            setFinalRecords(parsed.records);
            setDuplicatesSkipped(parsed.duplicates);
            setFileData({ headers, rows });
            setIsEasyMoneyFormat(true);
            setStep('preview');
            return;
          }

          setFileData({ headers, rows });
          performAutoMapping(headers, rows);
          setStep('map');
        } catch (err: any) {
          alert('កំហុសការដោះស្រាយ CSV: ' + err.message);
        }
      };
      reader.readAsText(file);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) processFile(file);
  };

  const performAutoMapping = (headers: string[], rows: any[]) => {
    const matched: Record<string, string> = {};
    targetFields.forEach(field => {
      const candidates = autoDetectMap[field.key] || [];
      const match = headers.find(h => {
        const normalized = h.toLowerCase().trim();
        return candidates.some(c => normalized === c || normalized.includes(c));
      });
      if (match) matched[field.key] = match;
    });
    setMappings(matched);
  };

  const handleMapFieldChange = (key: string, value: string) => {
    setMappings(prev => ({ ...prev, [key]: value }));
  };

  const handleAnalyzeMapping = () => {
    if (!fileData) return;

    // Validate required fields mapping
    const missing = targetFields.filter(f => f.required && !mappings[f.key]);
    if (missing.length > 0) {
      alert(`សូមផែនទី (Map) ជួរឈរដែលចាំបាច់៖ ${missing.map(m => m.label).join(', ')}`);
      return;
    }

    const recordsInDb = getRecords();
    const cleanRecords: FinanceRecord[] = [];
    let duplicates = 0;

    const exRate = config.exchangeRate || 4000;

    fileData.rows.forEach(row => {
      const dateVal = String(row[mappings.date] || '').trim();
      const typeVal = String(row[mappings.type] || '').trim();
      const categoryVal = String(row[mappings.category] || '').trim();
      
      const rawAmount = String(row[mappings.amount] || '').trim();
      const detectedCurrency = (rawAmount.includes('$') || rawAmount.toUpperCase().includes('USD')) ? '$' : (rawAmount.includes('៛') || rawAmount.toUpperCase().includes('KHR') || rawAmount.includes('រៀល')) ? '៛' : config.currency;
      const amountVal = Math.abs(Number(rawAmount.replace(/[^0-9.-]/g, '')));
      
      const descVal = mappings.description ? String(row[mappings.description] || '').trim() : '';
      const paymentVal = mappings.paymentMethod ? String(row[mappings.paymentMethod] || '').trim() : 'លុយសុទ្ធ';

      // Safe guard values
      if (!dateVal || !typeVal || isNaN(amountVal)) return;

      // Classify Type - Auto Khmer normalizations
      let type: RecordType = 'ចំណាយ';
      if (typeVal.includes('ចំណូល') || typeVal.toLowerCase() === 'income' || typeVal.toLowerCase() === 'in') {
        type = 'ចំណូល';
      }

      // Check duplicates logic
      let isDuplicate = false;
      if (duplicateCheck === 'DateAmountDesc') {
        isDuplicate = recordsInDb.some((existing: FinanceRecord) => 
          existing.date === dateVal && 
          existing.amount === amountVal && 
          existing.description === descVal &&
          existing.type === type
        );
      }

      if (isDuplicate) {
        duplicates++;
      } else {
        const dObj = new Date(dateVal);
        const validDate = isNaN(dObj.getTime()) ? new Date() : dObj;
        
        const amountKHR = detectedCurrency === '៛' ? amountVal : amountVal * exRate;
        const amountUSD = detectedCurrency === '$' ? amountVal : amountVal / exRate;
        const displayAmount = detectedCurrency === '$' ? `$${amountVal}` : `${amountVal} ៛`;

        cleanRecords.push({
          id: uuidv4(),
          date: dateVal,
          month: validDate.getMonth() + 1,
          year: validDate.getFullYear(),
          type,
          category: categoryVal || 'ផ្សេងៗ',
          description: descVal,
          currency: detectedCurrency as '៛' | '$',
          amount: amountVal,
          amountKHR,
          amountUSD,
          displayAmount,
          exchangeRate: exRate,
          paymentMethod: paymentVal || 'ផ្សេងៗ',
          memberId: session.id, // defaults current logged-in user
          note: '',
          createdAt: Date.now(),
          updatedAt: Date.now()
        });
      }
    });

    setFinalRecords(cleanRecords);
    setDuplicatesSkipped(duplicates);
    setStep('preview');
  };

  const handleCommitUpload = () => {
    if (finalRecords.length === 0) {
      showToast(config.language === 'English' ? 'No data rows to import!' : 'គ្មានទិន្នន័យដើម្បីនាំចូលឡើយ!');
      return;
    }

    setProgressModal({
      isOpen: true,
      title: config.language === 'English' ? 'Importing Data (Importing Data Rows)' : 'នាំចូលទិន្នន័យ (Importing Data Rows)',
      percent: 0,
      statusText: config.language === 'English' ? 'Analyzing spreadsheet structure...' : 'កំពុងសិក្សារចនាសម្ព័ន្ធ...',
      subText: config.language === 'English' ? `Preparing to load ${finalRecords.length} records` : `រៀបចំបញ្ចូលចំណូល/ចំណាយចំនួន ${finalRecords.length} ជួរ`
    });

    let progress = 0;
    const interval = setInterval(() => {
      progress += Math.floor(Math.random() * 20) + 12;
      if (progress >= 100) {
        progress = 100;
        clearInterval(interval);

        const merged = [...getRecords(), ...finalRecords];
        saveRecords(merged);
        
        // Add to offline queue
        const syncMode = getSyncMode();
        if (syncMode === 'sync' || syncMode === 'hybrid') {
          finalRecords.forEach(r => addToOfflineQueue('add', r.id, r));
          showToast(config.language === 'English'
            ? `Success: Imported ${finalRecords.length} rows, sync to Google Sheet is in progress`
            : `ជោគជ័យ៖ បាននាំចូល ${finalRecords.length} ជួរ, កំពុងបញ្ចូនទៅ Google Sheet បន្តិចម្តងៗ`
          );
          import('./googleSheetSync').then(m => m.syncLocalDataToGoogleSheet(merged, session.name));
        } else {
          showToast(config.language === 'English'
            ? `Success: Imported ${finalRecords.length} rows, skipped ${duplicatesSkipped} duplicates`
            : `ជោគជ័យ៖ បាននាំចូល ${finalRecords.length} ជួរ, រំលងជាន់គ្នា ${duplicatesSkipped} ជួរ`
          );
        }

        onUpdate();
        setProgressModal(null);
        setActiveTab('dashboard');
      } else {
        const rowsProcessed = Math.floor((finalRecords.length * progress) / 100);
        setProgressModal({
          isOpen: true,
          title: config.language === 'English' ? 'Importing Data (Importing Data Rows)' : 'នាំចូលទិន្នន័យ (Importing Data Rows)',
          percent: progress,
          statusText: progress > 55
            ? (config.language === 'English' ? 'Writing transaction history locally...' : 'កំពុងបញ្ចូលទៅក្នុងឃ្លាំងទិន្នន័យម៉ាស៊ីន...')
            : (config.language === 'English' ? 'Generating transactions ledger...' : 'កំពុងបង្កើតបញ្ជីប្រតិបត្តិការ...'),
          subText: config.language === 'English'
            ? `Processed ${rowsProcessed}/${finalRecords.length} lines`
            : `បានដំណើរការប្រភព ${rowsProcessed}/${finalRecords.length} ជួរ`
        });
      }
    }, 70);
  };

  const handleFileUploadInput = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  if (step === 'map' && fileData) {
    return (
      <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 pb-28 font-sans">
        <div className="bg-white p-5 rounded-[28px] border border-[#E8EEE9] space-y-4 shadow-sm font-sans">
          <h3 className="text-[11px] font-extrabold uppercase tracking-widest text-[#636E72] pb-2 border-b border-[#F0F4F1] font-sans">
            {config.language === 'English' ? "Data mapping fields" : "ផែនទីទិន្នន័យ"}
          </h3>
          
          <div className="space-y-4 font-sans">
            {targetFields.map(target => (
              <div key={target.key} className="grid grid-cols-1 md:grid-cols-3 gap-2 items-center pb-3 border-b border-dashed border-slate-100 last:border-b-0">
                <div>
                  <span className="text-xs font-extrabold text-slate-800 flex items-center gap-1 font-sans">
                    {target.label} {target.required && <span className="text-red-500 font-black">*</span>}
                  </span>
                  {config.language !== 'English' && target.khmerTips && (
                    <span className="text-[9px] text-[#74b9ff] font-bold block leading-tight font-sans">ណែនាំ៖ {target.khmerTips}</span>
                  )}
                </div>
                <div className="md:col-span-2">
                  <CustomSelect 
                    value={mappings[target.key] || ''} 
                    onChange={(v: string) => handleMapFieldChange(target.key, v)}
                    className="px-4 py-3 rounded-xl text-xs font-bold shadow-none border-[#E8EEE9] bg-slate-50/75 font-sans"
                    modalTitle={config.language === 'English' ? `Select column for ${target.label}` : `ជ្រើសរើសជួរឈរសម្រាប់ ${target.label}`}
                    placeholder={config.language === 'English' ? "-- Skip --" : "-- បោះបង់ ឬ មិនជ្រើសរើស --"}
                    options={fileData.headers.map((h: string) => ({ value: h, label: h }))}
                  />
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Configurations inline */}
        <div className="bg-[#fbfcfa] p-5 rounded-[28px] border border-[#E8EEE9] space-y-4 shadow-sm">
          <h3 className="text-[11px] font-extrabold uppercase tracking-widest text-slate-500">ការកំណត់ជាមុន (Parsing configurations)</h3>
          
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="text-[9px] uppercase font-bold text-slate-400 mb-1 block">Duplicate Check Matcher</label>
              <CustomSelect 
                value={duplicateCheck} 
                onChange={(v: string) => setDuplicateCheck(v as 'ID' | 'DateAmountDesc')}
                className="p-3 rounded-xl text-xs font-bold shadow-none border-[#E8EEE9] bg-white"
                modalTitle="យន្តការពិនិត្យស្ទួនអត្ថបទ"
                options={[
                  { value: 'DateAmountDesc', label: 'Date, Exact Amount & Description match' },
                  { value: 'ID', label: 'Skip unique IDs comparison checks' }
                ]}
              />
            </div>
            <div>
              <label className="text-[9px] uppercase font-bold text-slate-400 mb-1 block">Limit preview capacity threshold</label>
              <CustomSelect 
                value={importPreviewLimit} 
                onChange={(v: number) => setImportPreviewLimit(v)}
                className="p-3 rounded-xl text-xs font-bold shadow-none border-[#E8EEE9] bg-white"
                modalTitle="ដែនកំណត់ការមើលមុន"
                options={[
                  { value: 25, label: '25 Rows limits preview' },
                  { value: 50, label: '50 Rows limits preview' },
                  { value: 100, label: '100 Rows limits preview' }
                ]}
              />
            </div>
          </div>
        </div>

        <button onClick={handleAnalyzeMapping} className="w-full bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-extrabold tracking-wider uppercase py-4 rounded-2xl shadow-lg transition active:scale-95 leading-none">
          វិភាគផែនទីទិន្នន័យ (Preview Analysis)
        </button>
      </div>
    );
  }

  if (step === 'preview') {
    return (
      <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 pb-28 font-sans">
        {/* Stats card */}
        <div className="grid grid-cols-2 gap-3 font-sans">
          <div className="bg-[#ebfcf3] border border-[#c1f5d6] p-4 rounded-[24px]">
            <p className="text-[9px] uppercase text-[#1e824c] font-extrabold tracking-widest block font-sans">Ready rows loaded</p>
            <p className="font-extrabold text-[#115e33] text-2xl mt-0.5 font-sans">{finalRecords.length} ជួរ</p>
          </div>
          <div className="bg-[#fdf3f2] border border-[#fbd4d0] p-4 rounded-[24px]">
            <p className="text-[9px] uppercase text-[#c0392b] font-extrabold tracking-widest block font-sans">Duplicates skipped</p>
            <p className="font-extrabold text-[#962d22] text-2xl mt-0.5 font-sans">{duplicatesSkipped} ជួរ</p>
          </div>
        </div>

        {isEasyMoneyFormat && (
          <div className="bg-emerald-50 text-emerald-800 border border-emerald-100 p-4 rounded-2xl flex items-center gap-3 mb-2 font-bold text-xs font-sans">
            <span className="material-symbols-rounded text-emerald-600 shadow-sm bg-white p-2 text-base rounded-lg border border-emerald-100">verified</span>
            <span>EasyMoney Exchange Format រកឃើញដោយស្វ័យប្រវត្តិ! (Auto-detected structure)</span>
          </div>
        )}

        <div className="bg-white p-5 rounded-[28px] border border-[#E8EEE9] space-y-4 shadow-sm font-sans">
          <h3 className="text-[11px] font-extrabold uppercase tracking-widest text-[#636E72] pb-2 border-b border-[#F0F4F1] font-sans">គំរូទិន្នន័យ ({Math.min(finalRecords.length, importPreviewLimit)} ជួរដំបូង)</h3>
          
          <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1">
            {finalRecords.slice(0, importPreviewLimit).map((r, idx) => (
              <div key={r.id} className="flex justify-between items-center text-[10px] p-3 bg-slate-50 hover:bg-slate-100/70 transition rounded-xl font-sans border border-slate-100">
                <div className="font-sans">
                  <span className="font-extrabold text-slate-800">{formatDate(r.date)}</span> — <span className="font-bold text-slate-500">{translateCategory(r.category, config.language)}</span>
                  <p className="text-[9px] text-slate-400 mt-0.5">{r.description || 'No description listed'}</p>
                </div>
                <span className={cn(
                  "font-sans text-xs font-black",
                  r.type === 'ចំណូល' ? 'text-emerald-700' : 'text-red-500'
                )}>
                  {r.type === 'ចំណូល' ? '+' : '-'}{r.amount}
                </span>
              </div>
            ))}
          </div>
        </div>

        <button onClick={handleCommitUpload} className="w-full bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-extrabold tracking-wider uppercase py-4 rounded-2xl shadow-lg transition active:scale-95 leading-none font-sans">
          យល់ព្រមនាំចូលសរុប {finalRecords.length} កំណត់ត្រា (Submit final import)
        </button>
      </div>
    );
  }

  return (
    <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 font-sans pb-24">
      <div 
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={cn(
          "bg-white p-10 rounded-[40px] border-2 border-dashed text-center relative overflow-hidden group transition-all",
          isDragOver ? "border-indigo-500 bg-indigo-50/20 scale-[0.99]" : "border-[#E8EEE9] hover:border-emerald-200"
        )}
      >
        <div className="relative z-10 font-sans">
          <div className="w-20 h-20 bg-indigo-100 text-indigo-600 rounded-full flex items-center justify-center mx-auto mb-6 shadow-sm">
            <Upload className="w-8 h-8" />
          </div>
          <p className="font-extrabold text-sm text-indigo-950 mb-1 font-sans">ជ្រើសរើសឯកសារ CSV ឬ Excel (xlsx)</p>
          <p className="text-[10px] text-indigo-900/50 font-bold uppercase tracking-widest mb-8 font-sans">ទំហំអតិបរមា 15MB សម្រាប់ការទាញចូល</p>
          
          <label className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-4 px-8 rounded-[24px] shadow-xl shadow-indigo-200 transition cursor-pointer text-sm inline-block active:scale-95 leading-none font-sans">
            ជ្រើសរើសឯកសារ (Select File)
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleFileUploadInput} />
          </label>
        </div>
      </div>
      
      <div className="bg-indigo-50 border border-indigo-100 rounded-[32px] p-6 font-sans">
        <h4 className="text-xs font-extrabold text-indigo-900 uppercase tracking-widest mb-3 font-sans">លក្ខណៈពិសេសនៃការនាំចូល (Core Importer Features)</h4>
        <ul className="text-xs font-bold text-indigo-800/80 leading-relaxed mb-3 list-inside list-disc space-y-1.5 font-sans">
          <li>គាំទ្រទាំងប្រភេទ CSV និង Excel (xlsx, xls) ពេញស្ដង់ដា</li>
          <li>សមកាលកម្មដោយស្វ័យប្រវត្តិជាមួយទិន្នន័យ Google Sheet Sync</li>
          <li>ប្រព័ន្ធ Auto-Detect columns ផែនទីភាសាខ្មែរ និងអង់គ្លេស</li>
          <li>ត្រួតពិនិត្យជួរជាន់គ្នា (Checking duplicate rows reports) មុនការយល់ព្រម</li>
        </ul>
      </div>
    </div>
  );
}

function DataCenterView({ session, setActiveTab, records = [], users = [], config }: any) {
  const isAdmin = session?.role === 'Admin';
  
  const syncMode = getSyncMode();
  const lastSync = getLastSyncTime();
  const offlineQueue = getOfflineQueue();
  
  const totalCount = records.length;
  const incomeCount = records.filter((r: any) => r.type === 'ចំណូល').length;
  const expenseCount = records.filter((r: any) => r.type === 'ចំណាយ').length;
  
  const formatLastSync = (timeStr: string | null) => {
    if (!timeStr) return config.language === 'English' ? 'None yet' : 'មិនទាន់មាន';
    try {
      const d = new Date(timeStr);
      return d.toLocaleTimeString('km-KH', { hour: '2-digit', minute: '2-digit' }) + ' ' + d.toLocaleDateString('km-KH', { month: 'short', day: 'numeric' });
    } catch (e) {
      return timeStr;
    }
  };

  return (
    <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 font-sans pb-28">
      {/* ផ្នែក ប៊ូតុងបញ្ជា */}
      <div className="grid gap-4">
        {isAdmin && (
          <button onClick={() => setActiveTab('import')} className="bg-white p-6 rounded-[32px] border border-[#E8EEE9] flex items-center justify-between hover:border-emerald-200 transition text-left group shadow-sm hover:shadow-md active:scale-[0.98]">
            <div className="flex items-center gap-5">
              <div className="w-14 h-14 bg-[#F5F7F5] flex items-center justify-center text-emerald-600 rounded-[20px] transition-colors shadow-sm group-hover:bg-emerald-50"><span className="material-symbols-rounded text-2xl">arrow_downward</span></div>
              <div>
                <p className="font-extrabold text-sm text-emerald-950 mb-1">{config.language === 'English' ? "Import Data" : "នាំចូលទិន្នន័យ"}</p>
                <p className="text-[10px] text-slate-500 tracking-widest font-bold uppercase">{config.language === 'English' ? "Import from Excel or CSV" : "ទាញយកពី Excel ឬ CSV"}</p>
              </div>
            </div>
            <ChevronRight className="w-5 h-5 text-slate-300 group-hover:text-emerald-500 transition-colors" />
          </button>
        )}

        <button onClick={() => setActiveTab('export')} className="bg-white p-6 rounded-[32px] border border-[#E8EEE9] flex items-center justify-between hover:border-emerald-200 transition text-left group shadow-sm hover:shadow-md active:scale-[0.98]">
          <div className="flex items-center gap-5">
             <div className="w-14 h-14 bg-[#F5F7F5] flex items-center justify-center text-emerald-600 rounded-[20px] transition-colors shadow-sm group-hover:bg-emerald-50"><span className="material-symbols-rounded text-2xl">arrow_upward</span></div>
            <div>
              <p className="font-extrabold text-sm text-emerald-950 mb-1">{config.language === 'English' ? "Export Data" : "នាំចេញទិន្នន័យ"}</p>
              <p className="text-[10px] text-slate-500 tracking-widest font-bold uppercase">{config.language === 'English' ? "Download as Excel or CSV" : "ទាញយកជា Excel ឬ CSV"}</p>
            </div>
          </div>
          <ChevronRight className="w-5 h-5 text-slate-300 group-hover:text-emerald-500 transition-colors" />
        </button>
      </div>

      {/* ផ្នែក ស្ថិតិប្រព័ន្ធទិន្នន័យ */}
      <div className="space-y-3">
        <h3 className="text-[10px] font-extrabold uppercase tracking-widest text-[#636E72] pl-1">
          {config.language === 'English' ? "Database Statistics" : "ស្ថិតិប្រព័ន្ធទិន្នន័យ"}
        </h3>
        
        <div className="bg-white rounded-[32px] border border-[#E8EEE9] p-6 shadow-sm space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <div className="bg-[#F5F7F5] rounded-2xl p-4 flex flex-col justify-between">
              <div className="flex justify-between items-center mb-2">
                <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                  {config.language === 'English' ? "Total Records" : "សរុបកំណត់ត្រា"}
                </span>
                <span className="w-6 h-6 bg-emerald-100 rounded-lg flex items-center justify-center text-emerald-700">
                  <Database className="w-3.5 h-3.5" />
                </span>
              </div>
              <div className="text-xl font-black text-slate-800">{totalCount}</div>
              <div className="text-[9px] font-bold text-slate-400 mt-1 uppercase tracking-wider">
                {config.language === 'English' ? `${incomeCount} Income • ${expenseCount} Expense` : `${incomeCount} ចំណូល • ${expenseCount} ចំណាយ`}
              </div>
            </div>

            <div className="bg-[#F5F7F5] rounded-2xl p-4 flex flex-col justify-between">
              <div className="flex justify-between items-center mb-2">
                <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-500">
                  {config.language === 'English' ? "Storage Status" : "ស្ថានភាពរក្សាទុក"}
                </span>
                <span className={`w-6 h-6 rounded-lg flex items-center justify-center ${syncMode === 'hybrid' ? 'bg-sky-100 text-sky-700' : 'bg-emerald-100 text-emerald-700'}`}>
                  {syncMode === 'hybrid' ? <Cloud className="w-3.5 h-3.5" /> : <Shield className="w-3.5 h-3.5" />}
                </span>
              </div>
              <div className="text-sm font-black text-slate-800">
                {syncMode === 'hybrid' ? 'Cloud Sync' : (config.language === 'English' ? 'Stored Offline' : 'រក្សាទុក Offline')}
              </div>
              <div className="text-[9px] font-bold text-slate-400 mt-1 uppercase tracking-wider">
                {syncMode === 'hybrid' ? `Sync: ${formatLastSync(lastSync)}` : (config.language === 'English' ? 'On this device' : 'នៅលើឧបករណ៍')}
              </div>
            </div>
          </div>

          {syncMode === 'hybrid' && offlineQueue.length > 0 && (
            <div className="bg-amber-50 rounded-2xl p-4 border border-amber-100 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <span className="w-8 h-8 rounded-full bg-amber-100 flex items-center justify-center text-amber-700">
                  <RefreshCw className="w-4 h-4 animate-spin-slow" />
                </span>
                <div>
                  <div className="text-xs font-extrabold text-amber-950">
                    {config.language === 'English' ? "Unsynced Data" : "ទិន្នន័យមិនទាន់សមកាលកម្ម"}
                  </div>
                  <div className="text-[9px] font-bold text-amber-600 uppercase tracking-widest">
                    {config.language === 'English' ? `${offlineQueue.length} rows pending sync` : `មាន ${offlineQueue.length} ជួររង់ចាំបញ្ជូន`}
                  </div>
                </div>
              </div>
              <button 
                onClick={() => setActiveTab('googleSheetSync')} 
                className="bg-white hover:bg-amber-100/50 text-amber-800 border border-amber-200 text-[10px] font-extrabold px-3 py-1.5 rounded-lg active:scale-95 transition"
              >
                {config.language === 'English' ? "Inspect" : "ពិនិត្យ"}
              </button>
            </div>
          )}
        </div>
      </div>

      {/* ផ្នែក សុវត្ថិភាព និងការណែនាំ */}
      <div className="space-y-3">
        <h3 className="text-[10px] font-extrabold uppercase tracking-widest text-[#636E72] pl-1">
          {config.language === 'English' ? "Help & Safety Tips" : "ព័ត៌មានជំនួយ និងសុវត្ថិភាព"}
        </h3>
        
        <div className="bg-emerald-950/5 rounded-[32px] border border-emerald-900/5 p-6 space-y-4">
          <div className="flex gap-4">
            <div className="flex-none w-10 h-10 bg-emerald-600 rounded-xl flex items-center justify-center text-white shadow-md shadow-emerald-200">
              <Shield className="w-5 h-5" />
            </div>
            <div>
              <h4 className="text-xs font-extrabold text-emerald-950 mb-1">
                {config.language === 'English' ? "100% Secure & Private" : "រក្សាទុកដោយសុវត្ថិភាពខ្ពស់"}
              </h4>
              <p className="text-[11px] font-medium text-emerald-900/70 leading-relaxed">
                {config.language === 'English' 
                  ? "All income/expense entries are saved locally on your device's browser storage. No external agents can intercept it." 
                  : "ទិន្នន័យចំណូល-ចំណាយទាំងអស់ត្រូវបានរក្សាទុកក្នុង Local Storage នៃឧបករណ៍របស់អ្នកដោយសុវត្ថិភាព។"}
              </p>
            </div>
          </div>

          <div className="flex gap-4 border-t border-dashed border-emerald-900/10 pt-4">
            <div className="flex-none w-10 h-10 bg-sky-600 rounded-xl flex items-center justify-center text-white shadow-md shadow-sky-200">
              <Info className="w-5 h-5" />
            </div>
            <div>
              <h4 className="text-xs font-extrabold text-slate-900 mb-1">
                {config.language === 'English' ? "Data Loss Advisory" : "ការណែនាំអំពីការបាត់បង់ទិន្នន័យ"}
              </h4>
              <p className="text-[11px] font-medium text-slate-600 leading-relaxed">
                {config.language === 'English' 
                  ? "If you clear browser storage or history, app data might be lost. Ensure you export a backup to Excel/CSV weekly/monthly!" 
                  : "ប្រសិនបើអ្នកលុប Cache (ប្រវត្តិ) របស់ browser កម្មវិធីអាចនឹងបាត់បង់ទិន្នន័យ។ ដើម្បីសុវត្ថិភាព អ្នកគួររៀបចំទាញយក Excel / CSV បម្រុងទុកជារៀងរាល់សប្តាហ៍ ឬខែ!"}
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ExportView({ records, users, showToast, setActiveTab, config, setPromptYearState }: any) {
  const [eMonth, setEMonth] = useState(new Date().getMonth() + 1);
  const [eYear, setEYear] = useState(config.defaultYear || new Date().getFullYear());
  const [exportFormat, setExportFormat] = useState<'XLSX' | 'CSV'>('XLSX');

  const handleExportData = (mode: 'all' | 'month' | 'year') => {
    let filtered = records;
    let title = 'All';
    if (mode === 'month') {
      filtered = records.filter((r: any) => r.month === eMonth && r.year === eYear);
      title = `${eYear}_${eMonth}`;
    } else if (mode === 'year') {
      filtered = records.filter((r: any) => r.year === eYear);
      title = `${eYear}`;
    }

    if (filtered.length === 0) {
      alert(config.language === 'English' ? "No data to export" : "មិនមានទិន្នន័យដើម្បីទាញយកទេ (No data to export)");
      return;
    }

    if (exportFormat === 'XLSX') {
      exportEasyMoneyExcel(filtered, users, config, `EasyMoney_Exchange_${title}.xlsx`);
      showToast(config.language === 'English' ? "Excel export successful!" : "ទាញយក Excel Backup បានជោគជ័យ");
    } else if (exportFormat === 'CSV') {
      exportEasyMoneyCSV(filtered, users, config, `EasyMoney_Exchange_${title}.csv`);
      showToast(config.language === 'English' ? "CSV export successful!" : "ទាញយក CSV បានជោគជ័យ");
    }
  };

  return (
    <div className="space-y-4 font-sans max-w-xl mx-auto p-4 animate-in fade-in slide-in-from-right-8">
      {/* Header */}
      <div className="flex items-center justify-between border-b border-slate-150 pb-3">
        <h3 className="font-extrabold text-sm text-slate-800 uppercase tracking-widest flex items-center gap-1.5 font-sans">
          <span className="material-symbols-rounded text-lg text-emerald-600 font-sans">arrow_upward</span>
          {config.language === 'English' ? "Export Data" : "នាំចេញទិន្នន័យ"}
        </h3>
        <button 
          onClick={() => setActiveTab('settings')} 
          className="bg-slate-100 text-slate-500 hover:text-slate-800 hover:bg-slate-200 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors font-sans"
        >
          {config.language === 'English' ? "Back to Settings" : "ត្រឡប់ក្រោយ"}
        </button>
      </div>

      <div className="bg-white rounded-[32px] border border-[#E8EEE9] p-6 shadow-sm space-y-5">
        <div className="flex gap-3">
          <div className="flex-1">
            <CustomSelect 
              value={exportFormat} 
              onChange={setExportFormat}
              className="px-4 py-3 rounded-2xl text-[11px] font-bold uppercase shadow-none border border-emerald-500 bg-emerald-50 text-emerald-800 font-sans"
              modalTitle={config.language === 'English' ? "Select Format" : "ជ្រើសរើសប្រភេទ File"}
              options={[
                { value: 'XLSX', label: 'Excel (.xlsx)' },
                { value: 'CSV', label: 'CSV' }
              ]}
            />
          </div>
          <div className="w-28">
            <CustomSelect 
              value={eMonth} 
              onChange={setEMonth}
              className="px-4 py-3 rounded-2xl text-[11px] uppercase shadow-none border-[#E8EEE9] bg-[#FBFDFB] font-sans"
              modalTitle={config.language === 'English' ? "Select Month" : "ជ្រើសរើសខែ"}
              options={[1,2,3,4,5,6,7,8,9,10,11,12].map(m => ({ value: m, label: t(KHMER_MONTHS[m - 1], config.language) }))}
            />
          </div>
          <div className="w-28">
            <CustomSelect 
              value={eYear} 
              onChange={(v: any) => { if (v === 'ADD_NEW_YEAR') setPromptYearState(true); else setEYear(v); }}
              className="px-4 py-3 rounded-2xl text-[11px] uppercase shadow-none border-[#E8EEE9] bg-[#FBFDFB] font-sans"
              modalTitle={config.language === 'English' ? "Select Year" : "ជ្រើសរើសឆ្នាំ"}
              alignRight={true}
              options={[...generateYearOptions(records, config.customYears).map(opt => ({ ...opt, label: t(opt.label, config.language) })), { value: 'ADD_NEW_YEAR', label: config.language === 'English' ? '+ Add New Year' : '+ បន្ថែមឆ្នាំថ្មី' }]}
            />
          </div>
        </div>

        <div className="grid grid-cols-2 gap-3 font-sans">
          <button onClick={() => handleExportData('month')} className="bg-emerald-50 text-emerald-700 font-bold text-[11px] py-3 rounded-xl hover:bg-emerald-100 transition active:scale-95 font-sans">
            {config.language === 'English' ? 'Download Monthly' : 'ទាញយកប្រចាំខែ'}
          </button>
          <button onClick={() => handleExportData('year')} className="bg-blue-50 text-blue-700 font-bold text-[11px] py-3 rounded-xl hover:bg-blue-100 transition active:scale-95 font-sans">
            {config.language === 'English' ? 'Download Yearly' : 'ទាញយកប្រចាំឆ្នាំ'}
          </button>
        </div>
      </div>

      <button 
        onClick={() => handleExportData('all')} 
        className="w-full bg-white border border-[#E8EEE9] rounded-[32px] p-6 hover:bg-slate-50 transition text-left group font-sans"
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-5">
            <div className="w-14 h-14 bg-[#F5F7F5] flex items-center justify-center text-emerald-600 rounded-[20px] transition-colors shadow-sm group-hover:bg-emerald-50">
              <span className="material-symbols-rounded text-2xl font-sans">arrow_upward</span>
            </div>
            <div>
              <p className="font-extrabold text-sm text-emerald-950 mb-1 font-sans">
                {config.language === 'English' ? 'Export All' : 'ទាញយកទាំងអស់'}
              </p>
              <p className="text-[10px] text-slate-500 tracking-widest font-bold uppercase font-sans">
                {config.language === 'English' ? 'Data from day one' : 'ទិន្នន័យពីថ្ងៃដំបូង'}
              </p>
            </div>
          </div>
        </div>
      </button>
    </div>
  );
}

function MembersView({ users, session, onUpdate, showToast, setActiveTab, askConfirm, config }: any) {
  const [name, setName] = useState('');
  const [pin, setPin] = useState('');
  const [role, setRole] = useState<'Member'|'Admin'>('Member');
  
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState<'all'|'active'|'disabled'>('all');
  
  const [editingUser, setEditingUser] = useState<User | null>(null);
  const [editName, setEditName] = useState('');
  const [editRole, setEditRole] = useState<'Member'|'Admin'>('Member');
  const [editPin, setEditPin] = useState('');
  const [isChangingPin, setIsChangingPin] = useState(false);
  const [isFormOpen, setIsFormOpen] = useState(false);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if (!name || pin.length < 4) {
      showToast(config.language === 'English' ? "Please enter a name and PIN of at least 4 digits!" : "សូមបញ្ចូលឈ្មោះ និង PIN យ៉ាងហោចណាស់ ៤ ខ្ទង់!");
      return;
    }
    const curr = getUsers();
    curr.push({ id: uuidv4(), name, pin, role, isActive: true });
    saveUsers(curr);
    if (getSyncMode() !== 'local') syncUsersToGoogleSheet(curr);
    onUpdate();
    setName(''); setPin(''); setRole('Member');
    setIsFormOpen(false);
    showToast(config.language === 'English' ? "Family member account registered successfully!" : "ពង្រីកគណនីគ្រួសាររួចរាល់");
  };

  const handleToggleActive = (id: string) => {
    if (id === session?.id) {
      showToast(config.language === 'English' ? "You cannot block your own account!" : "អ្នកមិនអាចបិទគណនីខ្លួនឯងបានទេ!");
      return;
    }
    const curr = getUsers();
    const u = curr.find((x: any) => x.id === id);
    if (u) {
      u.isActive = !u.isActive;
      saveUsers(curr);
      if (getSyncMode() !== 'local') syncUsersToGoogleSheet(curr);
      onUpdate();
      showToast(u.isActive 
        ? (config.language === 'English' ? 'Account activated successfully' : 'បានបើកដំណើរការគណនីនេះឡើងវិញ') 
        : (config.language === 'English' ? 'Account deactivated successfully' : 'បានបិទគណនីគ្រួសារនេះ')
      );
    }
  };

  const handleDeleteUser = (id: string, userName: string) => {
    if (id === session?.id) {
      showToast(config.language === 'English' ? "You cannot delete your own account!" : "អ្នកមិនអាចលុបគណនីខ្លួនឯងបានទេ!");
      return;
    }
    askConfirm(config.language === 'English' ? `Are you sure you want to delete family account "${userName}"?` : `តើអ្នកពិតជាចង់លុបគណនីគ្រួសារ "${userName}" មែនទេ?`, () => {

      let curr = getUsers();
      curr = curr.filter((u: any) => u.id !== id);
      saveUsers(curr);
      if (getSyncMode() !== 'local') syncUsersToGoogleSheet(curr);
      onUpdate();
      showToast(config.language === 'English' ? "Family account deleted successfully!" : "បានលុបគណនីគ្រួសារដោយជោគជ័យ");
    });
  };

  const startEditing = (u: any) => {
    setEditingUser(u);
    setEditName(u.name);
    setEditRole(u.role);
    setEditPin('');
    setIsChangingPin(false);
  };

  const handleSaveEdit = () => {
    if (!editName) {
      showToast(config.language === 'English' ? "Please enter account name!" : "សូមបញ្ចូលឈ្មោះគណនី!");
      return;
    }
    const curr = getUsers();
    const u = curr.find((x: any) => x.id === editingUser?.id);
    if (u) {
      u.name = editName;
      u.role = editRole;
      if (isChangingPin) {
        if (editPin.length < 4 || editPin.length > 6) {
          showToast(config.language === 'English' ? "PIN must be between 4 and 6 digits!" : "PIN ត្រូវមានពី ៤ ទៅ ៦ ខ្ទង់!");
          return;
        }
        u.pin = editPin;
      }
      saveUsers(curr);
      if (getSyncMode() !== 'local') syncUsersToGoogleSheet(curr);
      onUpdate();
      setEditingUser(null);
      showToast(config.language === 'English' ? "Account details modified successfully!" : "បានកែប្រែព័ត៌មានគណនីដោយជោគជ័យ");
    }
  };

  const filteredUsers = users.filter((u: any) => {
    const s = searchQuery.toLowerCase();
    const matchName = u.name.toLowerCase().includes(s);
    const matchStatus = statusFilter === 'all' ? true : (statusFilter === 'active' ? u.isActive : !u.isActive);
    return matchName && matchStatus;
  });

  return (
    <div id="members-view-container" className="p-4 sm:p-6 space-y-5 animate-in fade-in slide-in-from-right-8 pb-28 font-sans">
      {/* SEARCH AND FILTERS */}
      <div id="members-search-filter-card" className="bg-white p-4.5 rounded-[24px] border border-slate-100 space-y-3.5 shadow-[0_4px_25px_rgba(0,0,0,0.015)] font-sans">
        <div className="relative group">
          <input 
            id="members-search-input"
            type="text" 
            placeholder={config.language === 'English' ? "Search members by name..." : "ស្វែងរកសមាជិកតាមឈ្មោះ..."} 
            value={searchQuery}
            onChange={e => setSearchQuery(e.target.value)}
            className="w-full bg-slate-50/60 hover:bg-slate-50 focus:bg-white border border-slate-200/60 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/5 transition-all pl-10 pr-4 h-11 rounded-xl text-xs font-bold text-slate-800 outline-none placeholder:text-slate-400"
          />
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 group-focus-within:text-emerald-500 transition-colors" />
        </div>
        
        <div id="members-filter-segments" className="bg-slate-50 p-1.5 rounded-xl flex gap-1 border border-slate-200/30">
          {[
            { value: 'all', label: config.language === 'English' ? 'All' : 'ទាំងអស់' },
            { value: 'active', label: config.language === 'English' ? 'Active' : 'កំពុងដំណើរការ' },
            { value: 'disabled', label: config.language === 'English' ? 'Disabled' : 'បិទការប្រើប្រាស់' }
          ].map((item) => (
            <button 
              id={`filter-btn-${item.value}`}
              key={item.value} 
              onClick={() => setStatusFilter(item.value as any)}
              className={cn(
                "flex-1 text-[11px] py-2 rounded-lg font-bold transition-all duration-300 outline-none",
                statusFilter === item.value 
                  ? "bg-white text-emerald-700 shadow-xs border border-slate-200/40 font-extrabold" 
                  : "text-slate-500 hover:text-slate-800"
              )}
            >
              {item.label}
            </button>
          ))}
        </div>
      </div>

      {/* CREATE NEW USER */}
      <form id="add-member-form" onSubmit={handleAdd} className="bg-white p-4.5 rounded-[24px] border border-slate-100 shadow-[0_8px_30px_rgba(0,0,0,0.02)] font-sans relative transition-all duration-300">
        <button 
          type="button"
          onClick={() => setIsFormOpen(!isFormOpen)}
          className="w-full flex items-center justify-between outline-none cursor-pointer group/header py-1"
        >
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 bg-emerald-500/10 border border-emerald-500/15 text-emerald-600 flex items-center justify-center rounded-xl transition-all group-hover/header:bg-emerald-500/15 group-hover/header:scale-105">
              <Plus className="w-4.5 h-4.5" />
            </div>
            <div className="text-left">
              <h3 className="font-black text-[12px] uppercase tracking-wider text-slate-800 group-hover/header:text-emerald-950 transition-colors">
                {config.language === 'English' ? "Add New Account" : "បន្ថែមគណនីថ្មី"}
              </h3>
              <p className="text-[10px] text-slate-400 font-bold mt-0.5">
                {config.language === 'English' ? "Click to open registration panel" : "ចុចទីនេះដើម្បីបើកផ្ទាំងចុះឈ្មោះ"}
              </p>
            </div>
          </div>
          <div className="w-8 h-8 flex items-center justify-center rounded-xl bg-slate-50 hover:bg-slate-100 transition-colors border border-slate-200/40">
            {isFormOpen ? (
              <ChevronUp className="w-4.5 h-4.5 text-slate-500" />
            ) : (
              <ChevronDown className="w-4.5 h-4.5 text-slate-500" />
            )}
          </div>
        </button>

        <AnimatePresence initial={false}>
          {isFormOpen && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: "auto", opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              transition={{ duration: 0.25, ease: "easeInOut" }}
              className="overflow-visible"
            >
              <div className="pt-5 space-y-4 border-t border-slate-100/80 mt-4">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div className="relative group">
                    <input 
                      id="add-member-name-input"
                      required 
                      placeholder={config.language === 'English' ? "Enter name" : "វាយបញ្ចូលឈ្មោះ"} 
                      value={name} 
                      onChange={e=>setName(e.target.value)} 
                      className="w-full bg-slate-50/60 hover:bg-white focus:bg-white border border-slate-200/65 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/5 transition-all pl-10 pr-4 h-11.5 rounded-xl text-xs font-semibold text-slate-800 outline-none placeholder:text-slate-400" 
                    />
                    <UserIcon className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-emerald-500 transition-colors pointer-events-none" />
                  </div>
                  
                  <div className="relative group">
                    <input 
                      id="add-member-pin-input"
                      required 
                      placeholder={config.language === 'English' ? "PIN (4-6 digits)" : "PIN (៤-៦ ខ្ទង់)"} 
                      value={pin} 
                      onChange={e=>setPin(e.target.value)} 
                      pattern="[0-9]*" 
                      inputMode="numeric" 
                      className="w-full bg-slate-50/60 hover:bg-white focus:bg-white border border-slate-200/65 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/5 transition-all pl-10 pr-4 h-11.5 rounded-xl text-xs font-extrabold text-slate-850 outline-none text-center tracking-wider placeholder:text-slate-400" 
                    />
                    <Lock className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-emerald-500 transition-colors pointer-events-none" />
                  </div>

                  <div className="relative group">
                    <CustomSelect 
                      value={role} 
                      onChange={(v: string) => setRole(v as any)}
                      className="pl-10 pr-4 h-11.5 rounded-xl border border-slate-200/65 bg-slate-50/60 hover:bg-white focus-within:border-emerald-500 focus-within:ring-4 focus-within:ring-emerald-500/5 transition-all text-xs font-semibold text-slate-800 shadow-none py-0 flex items-center"
                      modalTitle={config.language === 'English' ? "Select Role" : "ផ្តល់សិទ្ធិ"}
                      options={[
                        { value: 'Member', label: config.language === 'English' ? 'Member' : 'សមាជិក' },
                        { value: 'Admin', label: config.language === 'English' ? 'Admin' : 'អ្នកគ្រប់គ្រង' }
                      ]}
                    />
                    <Shield className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-emerald-500 transition-colors pointer-events-none" />
                  </div>
                </div>

                <button 
                  id="add-member-submit-btn" 
                  type="submit" 
                  className="w-full bg-slate-900 hover:bg-slate-800 active:scale-[0.98] text-white font-black h-11.5 rounded-xl shadow-sm hover:shadow-md transition-all duration-300 text-xs uppercase tracking-widest flex items-center justify-center gap-2 cursor-pointer animate-in fade-in zoom-in-95"
                >
                  <PlusCircle className="w-4 h-4" />
                  {config.language === 'English' ? "Register Credentials" : "ចុះឈ្មោះចូលប្រព័ន្ធ"}
                </button>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </form>

      {/* USER CARDS LIST */}
      <div id="members-list-wrapper" className="space-y-4 font-sans">
        <div className="flex items-center justify-between px-1">
          <h4 className="text-[11px] uppercase font-bold text-slate-400 tracking-wider">
            {config.language === 'English' ? `Family Accounts (${filteredUsers.length})` : `បញ្ជីឈ្មោះគណនីគ្រួសារ (${filteredUsers.length} នាក់)`}
          </h4>
        </div>
        
        {filteredUsers.length === 0 ? (
          <div id="no-members-view" className="bg-white text-center py-12 rounded-[24px] border border-dashed border-slate-200 shadow-sm p-6">
            <Users className="w-10 h-10 text-slate-300 mx-auto mb-3" />
            <p className="text-xs text-slate-400 font-extrabold font-sans">
              {config.language === 'English' ? "No member accounts found" : "រកមិនឃើញសមាជិកណាម្នាក់ឡើយ!"}
            </p>
          </div>
        ) : (
          filteredUsers.map((u: any) => {
            return (
              <div id={`member-card-${u.id}`} key={u.id} className="bg-white p-5 rounded-[22px] border border-slate-100 hover:border-emerald-500/15 shadow-[0_4px_25px_rgba(0,0,0,0.012)] hover:shadow-[0_12px_36px_rgba(16,185,129,0.03)] hover:scale-[1.005] transition-all duration-300 flex flex-col gap-4 relative overflow-hidden group">
                <div className="flex items-start gap-4">
                  <div className={cn(
                    "w-12 h-12 rounded-xl flex items-center justify-center font-black text-[13px] text-white shadow-xs shrink-0 transition-all duration-300 group-hover:scale-105", 
                    u.role === 'Admin' ? 'bg-gradient-to-tr from-indigo-500 to-violet-600' : 'bg-gradient-to-tr from-emerald-500 to-teal-600'
                  )}>
                    {u.name.substring(0,2).toUpperCase()}
                  </div>
                  <div className="flex-1 pt-0.5">
                    <div className="flex items-center gap-2 mb-2">
                      <p className="font-extrabold text-xs text-slate-805">
                        {u.name}
                      </p>
                      {u.id === session?.id && (
                        <span className="text-[8px] bg-gradient-to-r from-amber-500 to-orange-500 text-white px-2 py-0.5 rounded-md font-black uppercase tracking-wider shadow-sm">
                          {config.language === 'English' ? "Myself" : "ខ្ញុំផ្ទាល់"}
                        </span>
                      )}
                    </div>
                    <div className="flex flex-wrap items-center gap-1.5">
                      <span className={cn(
                        "text-[9px] px-2 py-0.5 rounded-md font-bold uppercase tracking-wider flex items-center gap-1 border",
                        u.role === 'Admin' ? 'bg-indigo-50 text-indigo-700 border-indigo-100/60' : 'bg-emerald-50 text-emerald-700 border-emerald-100/60'
                      )}>
                        <Shield className="w-2.5 h-2.5" />
                        {u.role === 'Admin' ? (config.language === 'English' ? 'Admin' : 'អភិបាល') : (config.language === 'English' ? 'Member' : 'សមាជិក')}
                      </span>
                      <span className={cn(
                        "text-[9px] px-2 py-0.5 rounded-md font-bold flex items-center gap-1 border",
                        u.isActive ? "bg-emerald-50 text-emerald-700 border-emerald-100/60" : "bg-red-50 text-red-500 border-red-100/60"
                      )}>
                        <span className={cn("w-1.5 h-1.5 rounded-full animate-pulse", u.isActive ? "bg-emerald-500" : "bg-red-500")} />
                        {u.isActive ? (config.language === 'English' ? 'Active' : 'សកម្ម') : (config.language === 'English' ? 'Disabled' : 'បានបិទ')}
                      </span>
                      <span className="text-[9px] bg-slate-50 border border-slate-100/80 text-slate-400 px-2 py-0.5 rounded-md lg:rounded-lg font-mono font-bold flex items-center gap-1">
                        <Lock className="w-2.5 h-2.5" />
                        PIN: {u.pin ? "••••" : (config.language === 'English' ? "None" : "គ្មាន")}
                      </span>
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2 pt-3.5 border-t border-slate-100/80 justify-end w-full">
                  <button 
                    id={`edit-member-btn-${u.id}`}
                    type="button" 
                    onClick={() => startEditing(u)} 
                    className="p-2 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50/50 rounded-xl border border-slate-200/50 hover:border-emerald-200/50 transition duration-200 active:scale-95 flex items-center justify-center"
                    title={config.language === 'English' ? "Edit Account" : "កែសម្រួលគណនី"}
                  >
                    <Edit2 className="w-3.5 h-3.5" />
                  </button>

                  {u.id === session?.id ? (
                    <div className="flex-1 max-w-[140px] px-3.5 py-2 bg-emerald-50/70 border border-emerald-100/40 rounded-xl flex items-center justify-center gap-1.5 text-[10px] font-bold text-emerald-700 select-none">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                      {config.language === 'English' ? "Current User" : "គណនីបច្ចុប្បន្ន"}
                    </div>
                  ) : (
                    <button 
                      id={`toggle-active-btn-${u.id}`}
                      type="button" 
                      onClick={() => handleToggleActive(u.id)} 
                      className={cn(
                        "px-4.5 py-1.5 rounded-xl text-xs font-bold transition-all border active:scale-95 flex-1 max-w-[140px] text-center cursor-pointer",
                        u.isActive 
                          ? "bg-amber-50 hover:bg-amber-100 text-amber-800 border-amber-200/40" 
                          : "bg-emerald-50/80 hover:bg-emerald-100 text-emerald-800 border-emerald-250/40"
                      )}
                    >
                      {u.isActive ? (config.language === 'English' ? "Block" : "បិទគណនី") : (config.language === 'English' ? "Activate" : "បើកគណនី")}
                    </button>
                  )}

                  <button 
                    id={`delete-member-btn-${u.id}`}
                    type="button" 
                    onClick={() => handleDeleteUser(u.id, u.name)} 
                    disabled={u.id === session?.id}
                    className={cn(
                      "p-2 rounded-xl transition border active:scale-95 flex items-center justify-center",
                      u.id === session?.id 
                        ? "opacity-35 cursor-not-allowed bg-slate-50 text-slate-300 border-slate-150" 
                        : "bg-red-50 hover:bg-red-100 text-red-600 border-red-100/65 hover:border-red-200"
                    )}
                    title={config.language === 'English' ? "Delete Account" : "លុបគណនី"}
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              </div>
            );
          })
        )}
      </div>

      <AnimatePresence>
        {editingUser && (
          <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center sm:p-4">
            <motion.div 
              initial={{opacity:0}} 
              animate={{opacity:1}} 
              exit={{opacity:0}} 
              onClick={() => setEditingUser(null)} 
              className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" 
            />
            <motion.div 
              initial={{ y: '100%' }} 
              animate={{ y: 0 }} 
              exit={{ y: '100%', opacity: 0 }} 
              transition={{ type: "spring", bounce: 0, duration: 0.4 }} 
              className="relative bg-white w-full max-w-md rounded-t-[24px] sm:rounded-[24px] p-6 space-y-5 shadow-2xl font-sans pb-safe border border-slate-100"
            >
              <div className="w-12 h-1.5 bg-slate-200 rounded-full mx-auto mb-2 sm:hidden" />
              <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 className="font-extrabold text-xs text-slate-800 uppercase tracking-widest flex items-center gap-1.5">
                  <Sliders className="w-4 h-4 text-emerald-600" /> {config.language === 'English' ? "Edit Account Info" : "កែប្រែព័ត៌មានគណនី"}
                </h3>
                <button 
                  type="button" 
                  onClick={() => setEditingUser(null)} 
                  className="bg-slate-100/80 text-slate-500 hover:text-slate-800 hover:bg-slate-200 p-1.5 rounded-xl transition-colors cursor-pointer"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="text-[9px] uppercase font-bold text-slate-400 mb-1 block">{config.language === 'English' ? "Account Name (Name)" : "ឈ្មោះគណនី (Name)"}</label>
                  <input 
                    type="text" 
                    value={editName} 
                    onChange={e => setEditName(e.target.value)} 
                    className="w-full bg-slate-50/60 hover:bg-white focus:bg-white border border-slate-200/60 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/5 transition-all p-3 h-11.5 rounded-xl text-xs font-semibold text-slate-800 outline-none" 
                    placeholder={config.language === 'English' ? "Enter new name" : "វាយឈ្មោះថ្មី"}
                  />
                </div>

                <div>
                  <label className="text-[9px] uppercase font-bold text-slate-400 mb-1 block">{config.language === 'English' ? "Access Role" : "សិទ្ធិប្រើប្រាស់ (Access Role)"}</label>
                  <CustomSelect 
                    value={editRole} 
                    onChange={(v: string) => setEditRole(v as any)}
                    className="h-11.5 py-0 px-4 rounded-xl border border-slate-200/60 bg-slate-50/60 hover:bg-white transition-all font-semibold text-slate-800 shadow-none text-xs flex items-center"
                    modalTitle={config.language === 'English' ? "Select Role" : "ផ្តល់សិទ្ធិ (Select Role)"}
                    options={[
                      { value: 'Member', label: config.language === 'English' ? 'Member' : 'សមាជិក (Member)' },
                      { value: 'Admin', label: config.language === 'English' ? 'Admin' : 'អ្នកគ្រប់គ្រង (Admin)' }
                    ]}
                  />
                </div>

                <div className="pt-3 border-t border-dashed border-slate-100">
                  <div className="flex justify-between items-center pb-2">
                    <span className="text-xs font-extrabold text-slate-800">{config.language === 'English' ? "Reset PIN" : "ប្ដូរលេខកូដសម្ងាត់ (Reset PIN)"}</span>
                    <button 
                      type="button" 
                      onClick={() => setIsChangingPin(!isChangingPin)}
                      className="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-800 px-3 py-1.5 rounded-lg font-bold transition cursor-pointer"
                    >
                      {isChangingPin ? (config.language === 'English' ? "Cancel" : "បិទវិញ") : (config.language === 'English' ? "Change PIN" : "កែប្រែ PIN")}
                    </button>
                  </div>
                  
                  {isChangingPin && (
                    <motion.div initial={{ opacity: 0, height: 0 }} animate={{ opacity: 1, height: 'auto' }} className="space-y-1.5">
                      <input 
                        type="text" 
                        placeholder={config.language === 'English' ? "Enter new 4-6 digit PIN" : "វាយបញ្ចូល PIN ៤-៦ ខ្ទង់ថ្មី"} 
                        value={editPin} 
                        onChange={e => setEditPin(e.target.value)}
                        pattern="[0-9]*" 
                        inputMode="numeric"
                        className="w-full bg-[#F5F7F5] border border-slate-200 p-3 rounded-xl text-xs font-extrabold text-emerald-950 text-center tracking-widest outline-none mt-2"
                      />
                      <p className="text-[8px] text-amber-600 font-extrabold leading-normal">{config.language === 'English' ? "* PIN must be between 4 and 6 digits for secure workspace sign-in" : "* PIN ត្រូវតែមាន ៤ ទៅ ៦ ខ្ទង់សុទ្ធ ដើម្បីសុវត្ថិភាពចូលគណនី"}</p>
                    </motion.div>
                  )}
                </div>
              </div>

              <div className="flex gap-2.5 pt-2">
                <button 
                  type="button" 
                  onClick={() => setEditingUser(null)} 
                  className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-600 hover:text-slate-800 font-bold h-11.5 rounded-xl text-xs transition active:scale-95 cursor-pointer"
                >
                  {config.language === 'English' ? "Cancel" : "បោះបង់"}
                </button>
                <button 
                  type="button" 
                  onClick={handleSaveEdit} 
                  className="flex-1 bg-slate-900 hover:bg-slate-800 text-white font-bold h-11.5 rounded-xl text-xs transition shadow-sm active:scale-95 cursor-pointer"
                >
                  {config.language === 'English' ? "Save" : "រក្សាទុក"}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
function CategoriesView({ categories, onUpdate, showToast, setActiveTab, askConfirm }: any) {
  const [name, setName] = useState('');
  const [type, setType] = useState<RecordType>('ចំណាយ');
  const userConfig = getConfig();

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if(!name) return;
    const curr = getCategories();
    curr.push({ id: uuidv4(), name, type });
    saveCategories(curr);
    if (getSyncMode() !== 'local') syncCategoriesToGoogleSheet(curr);
    onUpdate();
    setName('');
    showToast(userConfig.language === 'English' ? 'Category added successfully' : 'បន្ថែមប្រភេទលុយបានជោគជ័យ', 'success');
  };

  const handleDelete = (id: string) => {
    askConfirm(userConfig.language === 'English' ? 'Are you sure you want to delete this category?' : 'តើអ្នកចង់លុបប្រភេទនេះមែនទេ?', () => {
      const curr = getCategories().filter((c:any) => c.id !== id);
      saveCategories(curr);
      if (getSyncMode() !== 'local') syncCategoriesToGoogleSheet(curr);
      onUpdate();
    });
  }

  const handleDragEnd = (result: DropResult) => {
    if (!result.destination) return;
    
    const typeCategories = categories.filter((c:any) => c.type === type);
    const otherCategories = categories.filter((c:any) => c.type !== type);
    
    const items = Array.from(typeCategories);
    const [reorderedItem] = items.splice(result.source.index, 1);
    items.splice(result.destination.index, 0, reorderedItem);
    
    const newCategories = [...otherCategories, ...items];
    saveCategories(newCategories);
    if (getSyncMode() !== 'local') syncCategoriesToGoogleSheet(newCategories);
    onUpdate();
  };

  const filteredCategories = categories.filter((c:any) => c.type === type);

  return (
    <div id="categories-view-container" className="p-4 sm:p-6 space-y-5 animate-in fade-in slide-in-from-right-8 pb-24 font-sans">
      
      {/* SWITCHER */}
      <div id="categories-switcher" className="bg-slate-50 p-1.5 rounded-xl flex gap-1 border border-slate-200/30">
        <button
          id="categories-income-tab-btn"
          type="button"
          onClick={() => setType('ចំណូល')}
          className={cn(
            "flex-1 text-center py-2.5 rounded-lg cursor-pointer transition-all duration-300 font-bold text-xs flex items-center justify-center gap-1.5 outline-none border focus:outline-none",
            type === 'ចំណូល' 
              ? "bg-white text-emerald-700 border-slate-200/40 shadow-xs font-black" 
              : "text-slate-500 border-transparent hover:text-slate-800"
          )}
        >
          <TrendingUp className="w-4 h-4 text-emerald-500" />
          <span>{userConfig.language === 'English' ? "Income" : "ចំណូល"}</span>
        </button>
        <button
          id="categories-expense-tab-btn"
          type="button"
          onClick={() => setType('ចំណាយ')}
          className={cn(
            "flex-1 text-center py-2.5 rounded-lg cursor-pointer transition-all duration-300 font-bold text-xs flex items-center justify-center gap-1.5 outline-none border focus:outline-none",
            type === 'ចំណាយ' 
              ? "bg-white text-rose-700 border-slate-200/40 shadow-xs font-black" 
              : "text-slate-500 border-transparent hover:text-slate-800"
          )}
        >
          <TrendingDown className="w-4 h-4 text-rose-500" />
          <span>{userConfig.language === 'English' ? "Expense" : "ចំណាយ"}</span>
        </button>
      </div>

      {/* ADD CATEGORY FORM */}
      <form id="add-category-form" onSubmit={handleAdd} className="bg-white p-4 rounded-[20px] border border-slate-100 shadow-[0_4px_25px_rgba(0,0,0,0.015)] font-sans relative transition-all duration-300">
        <div className="flex gap-2">
          <div className="relative flex-1 group">
            <input 
              id="category-name-input"
              required 
              placeholder={userConfig.language === 'English' ? `Add new ${type === 'ចំណូល' ? 'Income' : 'Expense'} Category...` : `បន្ថែមប្រភេទ${type}ថ្មី...`} 
              value={name} 
              onChange={e=>setName(e.target.value)} 
              className="w-full bg-slate-50/60 hover:bg-white focus:bg-white border border-slate-200/60 focus:border-indigo-500 focus:ring-4 focus:ring-indigo-500/5 transition-all pl-10 pr-4 h-11.5 rounded-xl text-xs font-semibold text-slate-800 outline-none placeholder:text-slate-400" 
            />
            <Tag className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 group-focus-within:text-indigo-500 transition-colors pointer-events-none" />
          </div>
          <button 
            id="add-category-submit"
            type="submit" 
            className={cn(
              "px-5 h-11.5 rounded-xl font-black text-xs uppercase tracking-wider flex items-center justify-center gap-1.5 cursor-pointer text-white transition-all duration-350 hover:opacity-95 active:scale-[0.98] shadow-xs shrink-0 select-none",
              type === 'ចំណូល' ? "bg-gradient-to-r from-emerald-500 to-teal-500" : "bg-gradient-to-r from-rose-500 to-red-500"
            )}
          >
            <Plus className="w-4 h-4" />
            <span>{userConfig.language === 'English' ? "Add" : "បន្ថែម"}</span>
          </button>
        </div>
      </form>

      {/* CATEGORIES LIST CONTAINER */}
      <div id="categories-list-wrapper" className="space-y-3.5 font-sans">
        <div className="flex items-center justify-between px-1">
          <h4 className="text-[11px] uppercase font-bold text-slate-400 tracking-wider flex items-center gap-1.5">
            {type === 'ចំណូល' ? (
              <>
                <TrendingUp className="w-3.5 h-3.5 text-emerald-500" />
                <span>{userConfig.language === 'English' ? `Income Categories (${filteredCategories.length})` : `បញ្ជីចំណូល (${filteredCategories.length} ប្រភេទ)`}</span>
              </>
            ) : (
              <>
                <TrendingDown className="w-3.5 h-3.5 text-rose-500" />
                <span>{userConfig.language === 'English' ? `Expense Categories (${filteredCategories.length})` : `បញ្ជីចំណាយ (${filteredCategories.length} ប្រភេទ)`}</span>
              </>
            )}
          </h4>
        </div>
        
        <DragDropContext onDragEnd={handleDragEnd}>
          <Droppable droppableId="categories-list">
            {(provided) => (
              <div {...provided.droppableProps} ref={provided.innerRef} className="space-y-2.5 bg-[#F8FAFC]/55 p-3.5 rounded-2xl border border-slate-100/70">
                {filteredCategories.map((c:any, index: number) => (
                  <Draggable key={c.id} draggableId={c.id} index={index}>
                    {(provided, snapshot) => (
                      <div 
                        id={`category-item-${c.id}`}
                        ref={provided.innerRef}
                        {...provided.draggableProps}
                        className={cn(
                          "bg-white px-4 py-3 flex justify-between items-center group transition-all duration-250 select-none",
                          snapshot.isDragging 
                            ? "shadow-lg border-indigo-200 bg-indigo-50/10 scale-[1.015] rounded-xl rotate-[0.5deg] z-50 border" 
                            : "border-slate-100 hover:border-slate-200/60 hover:shadow-[0_4px_20px_rgba(0,0,0,0.012)] rounded-xl border shadow-xs"
                        )}
                      >
                        <div className="flex items-center gap-3">
                          <div {...provided.dragHandleProps} className="text-slate-300 hover:text-slate-400 cursor-grab active:cursor-grabbing p-1.5 rounded-lg hover:bg-slate-50 transition-colors flex items-center justify-center">
                            <GripVertical className="w-3.5 h-3.5" />
                          </div>
                          <div className={cn(
                            "w-2 h-2 rounded-full",
                            type === 'ចំណូល' ? "bg-emerald-500/80" : "bg-rose-500/80"
                          )} />
                          <span className="text-xs font-bold text-slate-700">
                            {translateCategory(c.name, userConfig.language)}
                          </span>
                        </div>
                        <button 
                          id={`delete-category-btn-${c.id}`}
                          type="button" 
                          onClick={()=>handleDelete(c.id)} 
                          className="text-slate-400 hover:text-rose-500 transition-all p-2 bg-slate-50 hover:bg-rose-500/10 active:scale-95 rounded-lg flex items-center justify-center border border-slate-100/50 hover:border-rose-100/50 cursor-pointer"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    )}
                  </Draggable>
                ))}
                {provided.placeholder}
                {filteredCategories.length === 0 && (
                  <div id="no-categories-view" className="text-center py-10 rounded-xl border border-dashed border-slate-200 shadow-xs p-6 bg-white">
                    <Tag className="w-8 h-8 text-slate-300 mx-auto mb-2" />
                    <p className="text-xs font-bold text-slate-400">{userConfig.language === 'English' ? "No categories added yet" : "មិនទាន់មានប្រភេទនៅឡើយទេ"}</p>
                  </div>
                )}
              </div>
            )}
          </Droppable>
        </DragDropContext>
      </div>
    </div>
  );
}

function GoogleSheetSyncView({ session, records, onUpdate, showToast, setActiveTab, askConfirm, config }: any) {
  const [mode, setMode] = useState<SyncMode>(getSyncMode());
  const [url, setUrl] = useState(getSyncUrl());
  const [testing, setTesting] = useState(false);
  const [syncingAll, setSyncingAll] = useState(false);
  const [queueCount, setQueueCount] = useState(getOfflineQueue().length);
  const [lastSync, setLastSync] = useState(getLastSyncTime());

  const handleSaveSettings = () => {
    saveSyncMode(mode);
    saveSyncUrl(url);
    onUpdate();
    showToast(t('រក្សាទុកការកំណត់បានជោគជ័យ', config.language));
  };

  const handleTestConnection = async () => {
    if (!url) {
      showToast(t('សូមបញ្ចូល API URL ជាមុនសិន!', config.language));
      return;
    }
    setTesting(true);
    try {
      const ok = await testGoogleSheetConnection(url);
      if (ok) {
        showToast(t('ភ្ជាប់ជោគជ័យ! កំពុងពិនិត្យទិន្នន័យចាស់...', config.language), 'sync');
        saveSyncUrl(url);
        saveSyncMode(mode);
        
        try {
          const gasRecords = await loadRecordsFromGoogleSheet();
          if (gasRecords && gasRecords.length > 0) {
            let updatedRecords = [...records];
            let addedCount = 0;
            gasRecords.forEach((gasRec: any) => {
               const existsById = updatedRecords.some(r => r.id === gasRec.id);
               const existsByValue = updatedRecords.some(r => 
                  r.date === gasRec.date && 
                  String(r.amount) === String(gasRec.amount) && 
                  String(r.description || '').toLowerCase() === String(gasRec.description || '').toLowerCase()
               );
               if (!existsById && !existsByValue) {
                 updatedRecords.push(gasRec);
                 addedCount++;
               }
            });
            if (addedCount > 0) {
              updatedRecords.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
              import('./store').then(m => {
                m.saveRecords(updatedRecords);
                onUpdate();
                showToast(config.language === 'English' ? `Pulled data from Google Sheet successfully (${addedCount} rows)` : `បានទាញទិន្នន័យពី Google Sheet រួចរាល់ (${addedCount} ជួរ)`);
              });
            } else {
              showToast(t('មិនមានទិន្នន័យថ្មីក្នុង Google Sheet ទេ', config.language));
              onUpdate();
            }
          } else {
             showToast(t('មិនមានទិន្នន័យក្នុង Google Sheet', config.language));
             onUpdate();
          }
        } catch(err: any) {
           showToast(err.message || t('កំហុសក្នុងការទាញទិន្នន័យពី Google Sheet', config.language), 'error');
           onUpdate();
        }
      }
    } catch (err: any) {
      alert(err.message || (config.language === 'English' ? 'Connection Error' : 'កំហុសនៃការតភ្ជាប់'));
    } finally {
      setTesting(false);
    }
  };

  const handleManualSyncAll = async () => {
    setSyncingAll(true);
    try {
      showToast(t('កំពុងចាប់ផ្តើមសមកាលកម្ម...', config.language));
      let pulledCount = 0;
      let currentRecords = [...records];
      
      if (url) {
        try {
            const gasRecords = await loadRecordsFromGoogleSheet();
            if (gasRecords && gasRecords.length > 0) {
              gasRecords.forEach((gasRec: any) => {
                 const existsById = currentRecords.some(r => r.id === gasRec.id);
                 const existsByValue = currentRecords.some(r => 
                    r.date === gasRec.date && 
                    String(r.amount) === String(gasRec.amount) && 
                    String(r.description || '').toLowerCase() === String(gasRec.description || '').toLowerCase()
                 );
                 if (!existsById && !existsByValue) {
                   currentRecords.push(gasRec);
                   pulledCount++;
                 }
              });
              if (pulledCount > 0) {
                currentRecords.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());
                const m = await import('./store');
                m.saveRecords(currentRecords);
              }
            }
        } catch (err) {
           console.log("Could not pull records", err);
        }
      }

      const res = await syncLocalDataToGoogleSheet(currentRecords, session.name);
      await syncUsersToGoogleSheet(getUsers());
      await syncCategoriesToGoogleSheet(getCategories());
      await syncAppConfigToGoogleSheet(getConfig());
      if (res.success) {
        setQueueCount(getOfflineQueue().length);
        setLastSync(getLastSyncTime());
        if (pulledCount > 0) {
           showToast(config.language === 'English' ? `Sync successful! (Pulled ${pulledCount} rows)` : `សមកាលកម្មទិន្នន័យជោគជ័យ! (ទាញមក ${pulledCount} ជួរ)`);
        } else {
           showToast(t('សមកាលកម្មទិន្នន័យបានជោគជ័យពេញលេញ!', config.language));
        }
        onUpdate();
      } else {
        showToast(res.error || t('សមកាលកម្មបានជួបបញ្ហាមួយចំនួន', config.language));
        if (pulledCount > 0) onUpdate();
      }
    } catch (err: any) {
      showToast(err.message || t('កំហុសនៃការ Sync', config.language));
    } finally {
      setSyncingAll(false);
    }
  };

  const handleClearQueue = () => {
    askConfirm(t('តើអ្នកពិតជាចង់សម្អាត Offline Queue ដែលមិនទាន់ Sync មែនទេ?', config.language), () => {

      localStorage.removeItem('EM_OFFLINE_QUEUE');
      setQueueCount(0);
      showToast(t('បានសម្អាត Queue រួចរាល់', config.language));
    });
  };

  return (
    <div className="p-6 space-y-6 animate-in fade-in slide-in-from-right-8 pb-24">
      {/* Sync Mode Selector */}
      <div className="bg-white p-6 rounded-[32px] border border-[#E8EEE9] space-y-4 shadow-sm">
        <h3 className="text-xs font-extrabold uppercase tracking-widest text-emerald-950 mb-3">របៀបសមកាលកម្ម (Sync Mode)</h3>
        <div className="space-y-3">
          {/* Local Only */}
          <label className={cn(
            "relative flex flex-col sm:flex-row items-start sm:items-center gap-4 p-5 rounded-[24px] border-2 cursor-pointer transition-all duration-300",
            mode === 'local' 
              ? "border-emerald-500 bg-emerald-50/40 shadow-sm" 
              : "border-[#E8EEE9] bg-white hover:border-emerald-200 hover:bg-slate-50"
          )}>
            <input type="radio" value="local" checked={mode === 'local'} onChange={() => setMode('local')} className="sr-only" />
            
            <div className={cn(
              "p-3 rounded-2xl transition-colors shrink-0",
              mode === 'local' ? "bg-emerald-500 text-white shadow-md shadow-emerald-500/20" : "bg-slate-100 text-slate-500"
            )}>
              <WifiOff className="w-6 h-6" />
            </div>

            <div className="flex-1 pr-8">
              <h4 className={cn("font-bold text-sm tracking-tight mb-1", mode === 'local' ? "text-emerald-900" : "text-slate-800")}>
                1. Local Only
              </h4>
              <p className="text-xs text-slate-500 leading-relaxed font-medium">គ្មានអ៊ីនធឺណិតក៏ប្រើបាន។ រក្សាទុកក្នុងឧបករណ៍។ ចាំបាច់ត្រូវ Backup ខ្លួនឯងតាម CSV/JSON។</p>
            </div>

            {mode === 'local' && (
              <div className="absolute right-5 top-1/2 -translate-y-1/2">
                <CheckCircle2 className="w-6 h-6 text-emerald-500" />
              </div>
            )}
          </label>

          {/* Google Sheet Sync */}
          <label className={cn(
            "relative flex flex-col sm:flex-row items-start sm:items-center gap-4 p-5 rounded-[24px] border-2 cursor-pointer transition-all duration-300",
            mode === 'sync' 
              ? "border-emerald-500 bg-emerald-50/40 shadow-sm" 
              : "border-[#E8EEE9] bg-white hover:border-emerald-200 hover:bg-slate-50"
          )}>
            <input type="radio" value="sync" checked={mode === 'sync'} onChange={() => setMode('sync')} className="sr-only" />
            
            <div className={cn(
              "p-3 rounded-2xl transition-colors shrink-0",
              mode === 'sync' ? "bg-emerald-500 text-white shadow-md shadow-emerald-500/20" : "bg-slate-100 text-slate-500"
            )}>
              <Cloud className="w-6 h-6" />
            </div>

            <div className="flex-1 pr-8">
              <h4 className={cn("font-bold text-sm tracking-tight mb-1", mode === 'sync' ? "text-emerald-900" : "text-slate-800")}>
                2. Google Sheet Sync
              </h4>
              <p className="text-xs text-slate-500 leading-relaxed font-medium">ភ្ជាប់ទិន្នន័យពី Sheet ផ្ទាល់។ ទាមទារអ៊ីនធឺណិតគ្រប់ពេល។</p>
            </div>

            {mode === 'sync' && (
              <div className="absolute right-5 top-1/2 -translate-y-1/2">
                <CheckCircle2 className="w-6 h-6 text-emerald-500" />
              </div>
            )}
          </label>

          {/* Hybrid Mode */}
          <label className={cn(
            "relative flex flex-col sm:flex-row items-start sm:items-center gap-4 p-5 rounded-[24px] border-2 cursor-pointer transition-all duration-300",
            mode === 'hybrid' 
              ? "border-emerald-500 bg-emerald-50/40 shadow-sm" 
              : "border-[#E8EEE9] bg-white hover:border-emerald-200 hover:bg-slate-50"
          )}>
            <input type="radio" value="hybrid" checked={mode === 'hybrid'} onChange={() => setMode('hybrid')} className="sr-only" />
            
            <div className={cn(
              "p-3 rounded-2xl transition-colors shrink-0",
              mode === 'hybrid' ? "bg-emerald-500 text-white shadow-md shadow-emerald-500/20" : "bg-slate-100 text-slate-500"
            )}>
              <RefreshCw className="w-6 h-6" />
            </div>

            <div className="flex-1 pr-8">
              <div className="flex items-center gap-2 mb-1">
                <h4 className={cn("font-bold text-sm tracking-tight", mode === 'hybrid' ? "text-emerald-900" : "text-slate-800")}>
                  3. Hybrid
                </h4>
                <span className="bg-emerald-100 text-emerald-700 text-[10px] font-extrabold uppercase px-2 py-0.5 rounded-full">ណែនាំបំផុត</span>
              </div>
              <p className="text-xs text-slate-500 leading-relaxed font-medium">ល្អបំផុត! គ្មានអ៊ីនធឺណិតក៏ប្រើបាន និង Backup អូតូទៅ Sheet ពេលមានអ៊ីនធឺណិតវិញ។</p>
            </div>

            {mode === 'hybrid' && (
              <div className="absolute right-5 top-1/2 -translate-y-1/2">
                <CheckCircle2 className="w-6 h-6 text-emerald-500" />
              </div>
            )}
          </label>
        </div>
      </div>

      {/* URL Parameter section */}
      {mode !== 'local' && (
        <div className="bg-white p-6 sm:p-8 rounded-[32px] border border-[#E8EEE9] shadow-sm animate-in fade-in slide-in-from-bottom-4 duration-500">
          <div className="flex items-center gap-3 mb-5">
            <div className="p-2.5 rounded-2xl bg-emerald-50 text-emerald-600">
              <Cloud className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-sm font-extrabold text-emerald-950">ការភ្ជាប់ Web App URL</h3>
              <p className="text-xs text-slate-500 font-medium tracking-tight mt-0.5">ភ្ជាប់ទៅកាន់ Google Apps Script របស់អ្នក</p>
            </div>
          </div>
          
          <div className="space-y-5">
            <div>
              <label className="text-[10px] uppercase font-extrabold tracking-widest text-slate-500 block mb-2 ml-1">Google Apps Script URL <span className="text-red-500">*</span></label>
              <div className="relative group">
                <textarea 
                  rows={3}
                  placeholder="https://script.google.com/macros/s/.../exec" 
                  value={url} 
                  onChange={e => setUrl(e.target.value)} 
                  className="w-full bg-[#f8faf8] border-2 border-[#E8EEE9] px-5 py-4 rounded-[24px] text-xs font-mono outline-none focus:bg-white focus:border-emerald-500 shadow-sm transition-all text-emerald-900 resize-none group-hover:border-emerald-200"
                />
              </div>
            </div>
            
            <div className="flex flex-col sm:flex-row gap-3 pt-1">
              <button 
                type="button" 
                onClick={handleTestConnection} 
                disabled={testing || !url.trim()}
                className={cn(
                  "flex-1 flex items-center justify-center gap-2 font-extrabold text-xs py-4 rounded-[20px] transition-all active:scale-[0.98]",
                  testing ? "bg-emerald-50 text-emerald-600" : 
                  !url.trim() ? "bg-slate-50 text-slate-400 cursor-not-allowed" : 
                  "bg-emerald-50 hover:bg-emerald-100 text-emerald-700 hover:shadow-md hover:shadow-emerald-100/50"
                )}
              >
                {testing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Wifi className="w-4 h-4" />}
                {testing ? 'កំពុងតេស្ត...' : 'តេស្តការតភ្ជាប់'}
              </button>
              
              <button 
                type="button" 
                onClick={handleSaveSettings}
                disabled={!url.trim()}
                className={cn(
                  "flex-1 flex items-center justify-center gap-2 font-extrabold text-xs py-4 rounded-[20px] transition-all active:scale-[0.98]",
                  !url.trim() ? "bg-slate-100 text-slate-400 cursor-not-allowed" : "bg-emerald-950 hover:bg-emerald-900 text-white shadow-md hover:shadow-lg shadow-emerald-900/20"
                )}
              >
                <CheckCircle2 className="w-4 h-4" />
                រក្សាទុកការកំណត់
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Manual Actions and Status Dashboard */}
      {mode !== 'local' && (
        <div className="bg-white p-6 rounded-[32px] border border-[#E8EEE9] space-y-4 shadow-sm">
          <h3 className="text-xs font-extrabold uppercase tracking-widest text-emerald-950">ស្ថានភាពការងារ និង Dashboard</h3>
          <div className="space-y-3.5">
            <div className="flex justify-between items-center text-xs py-1 border-b border-dashed border-slate-100">
              <span className="text-slate-500 font-bold">ការតភ្ជាប់ចុងក្រោយបង្អស់</span>
              <span className="font-extrabold text-slate-850">{lastSync}</span>
            </div>
            
            <div className="flex justify-between items-center text-xs py-1 border-b border-dashed border-slate-100">
              <span className="text-slate-500 font-bold">ជួររង់ចាំ (Offline Queue)</span>
              <div className="flex items-center gap-2">
                <span className={cn("font-extrabold px-2.5 py-0.5 rounded-full text-[10px]", queueCount > 0 ? "bg-white text-emerald-600 border border-[#E8EEE9] shadow-sm border border-red-100" : "bg-emerald-50 text-emerald-700")}>
                  {queueCount} records
                </span>
                {queueCount > 0 && (
                  <button onClick={handleClearQueue} className="text-xs text-red-400 hover:underline">Clear</button>
                )}
              </div>
            </div>

            <div className="pt-2 flex flex-col gap-2">
              <button 
                type="button" 
                disabled={syncingAll}
                onClick={handleManualSyncAll}
                className="w-full border-2 border-[#E8EEE9] hover:bg-slate-50 active:scale-[0.98] font-bold text-xs py-4 rounded-[20px] transition-all flex items-center justify-center gap-2 text-emerald-900"
              >
                {syncingAll ? <span className="material-symbols-rounded text-base animate-spin text-emerald-600">sync</span> : <span className="material-symbols-rounded text-base text-emerald-600">sync</span>}
                សមកាលកម្មទិន្នន័យទាំងអស់ទៅ Sheet 📡
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Guide Cards */}
      <div className="bg-[#F5F7F5] p-5 rounded-[28px] border border-[#E8EEE9] space-y-3">
        <h4 className="text-[11px] font-extrabold text-emerald-900 uppercase tracking-wider flex items-center gap-1.5 pb-2 border-b border-emerald-100/30">
          <Database className="w-3.5 h-3.5" /> ការណែនាំខ្លីៗសម្រាប់ការតភ្ជាប់
        </h4>
        <ol className="text-[10px] text-[#636E72] leading-relaxed space-y-2 list-decimal pl-4 font-bold">
          <li>ដំឡើងកូដ <code className="bg-white px-2 py-0.5 rounded font-mono text-emerald-700 text-[10px]">Code.gs</code> ទៅកាន់ Apps Script នៃ Spreadsheet របស់អ្នក។</li>
          <li>ចុចរត់ <code className="bg-white px-2 py-0.5 rounded font-mono text-emerald-700 text-[10px]">setupEasyMoneySheet</code> ម្តងដើម្បីបង្កើតតារាងចាំបាច់។</li>
          <li>Deploy ជា Web App ដោយត្រៀមសិទ្ធិ Execute as <code className="bg-white px-2 py-0.5 rounded text-[10px]">Me</code> និង Who has access <code className="bg-white px-2 py-0.5 rounded text-[10px]">Anyone</code>។</li>
          <li>ចម្លង Web App URL មកបញ្ចូលក្នុងប្រអប់ខាងលើ រួចចុច Save។</li>
        </ol>
      </div>
    </div>
  );
}
