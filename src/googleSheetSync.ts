import { FinanceRecord } from './types';

export type SyncMode = 'local' | 'sync' | 'hybrid';

export interface SyncConfig {
  mode: SyncMode;
  url: string;
}

const KEYS = {
  SYNC_MODE: 'EM_SYNC_MODE',
  SYNC_URL: 'EM_SYNC_URL',
  OFFLINE_QUEUE: 'EM_OFFLINE_QUEUE',
  LAST_SYNC: 'EM_LAST_SYNC',
};

// Getters & Setters
export function getSyncMode(): SyncMode {
  return (localStorage.getItem(KEYS.SYNC_MODE) as SyncMode) || 'local';
}

export function saveSyncMode(mode: SyncMode) {
  localStorage.setItem(KEYS.SYNC_MODE, mode);
}

export function getSyncUrl(): string {
  return localStorage.getItem(KEYS.SYNC_URL) || '';
}

export function saveSyncUrl(url: string) {
  localStorage.setItem(KEYS.SYNC_URL, url.trim());
}

export interface QueueItem {
  action: 'add' | 'update' | 'delete';
  id: string;
  record?: FinanceRecord;
  timestamp: number;
}

export function getOfflineQueue(): QueueItem[] {
  try {
    const val = localStorage.getItem(KEYS.OFFLINE_QUEUE);
    return val ? JSON.parse(val) : [];
  } catch {
    return [];
  }
}

export function saveOfflineQueue(queue: QueueItem[]) {
  localStorage.setItem(KEYS.OFFLINE_QUEUE, JSON.stringify(queue));
}

export function addToOfflineQueue(action: 'add' | 'update' | 'delete', id: string, record?: FinanceRecord) {
  const queue = getOfflineQueue();
  
  // Clean redundant queue items to avoid extra requests
  // For example, if we have add and then update for the same record, we can optimize or just keep it sequentially.
  // Sequential keeping is safest, but we can deduplicate deletes of the same ID
  if (action === 'delete') {
    // If the record was added locally but never synced, and now is deleted, we can just remove the add queue item!
    const addIdx = queue.findIndex(q => q.id === id && q.action === 'add');
    if (addIdx !== -1) {
      queue.splice(addIdx, 1);
      saveOfflineQueue(queue);
      return;
    }
  }

  queue.push({
    action,
    id,
    record,
    timestamp: Date.now()
  });
  saveOfflineQueue(queue);
}

// REST Backend API Callers

// Global state on whether sync is in progress
let isSyncing = false;

// 1. Test Google Sheets Connection
export async function testGoogleSheetConnection(url?: string): Promise<boolean> {
  const apiUrl = url || getSyncUrl();
  if (!apiUrl) {
    throw new Error('សូមបញ្ចូល API Web App URL របស់ Google Apps Script របស់លោកអ្នក');
  }

  try {
    const res = await fetch(`${apiUrl}?action=test`, {
      method: 'GET',
      mode: 'cors',
      credentials: 'omit'
    });
    
    if (!res.ok) {
      throw new Error('ភ្ជាប់ទៅកាន់ Server មិនជោគជ័យឡើយ (HTTP ' + res.status + ')');
    }
    
    const textStatus = await res.text();
    let data;
    try {
      data = JSON.parse(textStatus);
    } catch {
      throw new Error('ការឆ្លើយតបពី Google Sheet API មិនទាន់ត្រឹមត្រូវ (សូមប្រាកដថាអ្នកបាន Deploy ជា Web App និងផ្តល់សិទ្ធិ "Anyone")');
    }

    if (data.status === 'success') {
      return true;
    } else {
      throw new Error(data.message || 'ភ្ជាប់ទៅកាន់ Google Sheet មិនបានជោគជ័យឡើយ');
    }
  } catch (err: any) {
    // Return Khmer-localized standard fetch failures
    let errMsg = err.message || 'កំហុសបណ្តាញអ៊ីនធឺណិត';
    if (err.message && err.message.includes('fetch')) {
      errMsg = 'មិនអាចទាក់ទង Google Sheets API បានទេ (សូមពិនិត្យអក្ខរាវិរុទ្ធ URL ឬបណ្តាញរបស់អ្នក)';
    } else if (err.toString().includes('Failed to fetch')) {
      errMsg = 'ការភ្ជាប់យឺត ឬដាច់ការភ្ជាប់ (Failed to fetch) សូមពិនិត្យមើល URL របស់លោកអ្នក';
    }
    throw new Error(errMsg);
  }
}

// 2. Syn Record directly (for add)
export async function syncRecordToGoogleSheet(record: FinanceRecord, user?: string): Promise<boolean> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'addRecord',
        data: record,
        user: user || 'Sys'
      })
    });
    
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    // console.error('Error syncing record to Google Sheet:', e);
    return false;
  }
}

// 3. Load Records from Google Sheet
export async function loadRecordsFromGoogleSheet(): Promise<FinanceRecord[]> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return [];

  const res = await fetch(`${apiUrl}?action=getRecords`, {
    method: 'GET',
    mode: 'cors'
  });
  
  if (!res.ok) {
    throw new Error('មិនអាចអានទិន្នន័យពី Google Sheet បានឡើយ (HTTP ' + res.status + ')');
  }

  const data = await res.json();
  if (data.status === 'success') {
    return data.records || [];
  } else {
    throw new Error(data.message || 'មិនអាចទទួលបានទិន្នន័យពី Google Sheet ឡើយ');
  }
}

// Load Settings from Google Sheet
export async function loadSettingsFromGoogleSheet(): Promise<any> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return null;

  const res = await fetch(`${apiUrl}?action=getSettings`, {
    method: 'GET',
    mode: 'cors'
  });
  if (!res.ok) throw new Error('មិនអាចអាន Settings ពី Google Sheet បានឡើយ (HTTP ' + res.status + ')');
  const data = await res.json();
  if (data.status === 'success') {
    return data.settings || {};
  } else {
    throw new Error(data.message || 'មិនអាចទទួលបាន Settings ពី Google Sheet ឡើយ');
  }
}

// Load Users/Family Members from Google Sheet
export async function loadUsersFromGoogleSheet(): Promise<any[]> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return [];

  const res = await fetch(`${apiUrl}?action=getUsers`, {
    method: 'GET',
    mode: 'cors'
  });
  if (!res.ok) throw new Error('មិនអាចអានបញ្ជីសមាជិកពី Google Sheet បានឡើយ (HTTP ' + res.status + ')');
  const data = await res.json();
  if (data.status === 'success') {
    return data.users || [];
  } else {
    throw new Error(data.message || 'មិនអាចទទួលបានបញ្ជីសមាជិកពី Google Sheet ឡើយ');
  }
}

// Load Categories from Google Sheet
export async function loadCategoriesFromGoogleSheet(): Promise<any[]> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return [];

  const res = await fetch(`${apiUrl}?action=getCategories`, {
    method: 'GET',
    mode: 'cors'
  });
  if (!res.ok) throw new Error('មិនអាចអានប្រភេទចំណូល/ចំណាយពី Google Sheet បានឡើយ (HTTP ' + res.status + ')');
  const data = await res.json();
  if (data.status === 'success') {
    return data.categories || [];
  } else {
    throw new Error(data.message || 'មិនអាចទទួលបានប្រភេទចំណូល/ចំណាយពី Google Sheet ឡើយ');
  }
}

// 4. Update existing record in Google Sheet
export async function updateRecordInGoogleSheet(record: FinanceRecord, user?: string): Promise<boolean> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'updateRecord',
        data: record,
        user: user || 'Sys'
      })
    });
    
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    // console.error('Error updating record in Google Sheet:', e);
    return false;
  }
}

// 5. Delete specific record in Google Sheet by ID
export async function deleteRecordFromGoogleSheet(id: string, user?: string): Promise<boolean> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'deleteRecord',
        id: id,
        user: user || 'Sys'
      })
    });
    
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    // console.error('Error deleting record from Google Sheet:', e);
    return false;
  }
}

// 6. Synchronize Offline Queue and Local data to Google Sheet
export async function syncLocalDataToGoogleSheet(localRecords?: FinanceRecord[], user?: string): Promise<{ success: boolean; syncedCount: number; error?: string }> {
  const apiUrl = getSyncUrl();
  if (!apiUrl) {
    return { success: false, syncedCount: 0, error: 'មិនទាន់ទាក់ទង URL បានទេ' };
  }

  if (isSyncing) {
    return { success: false, syncedCount: 0, error: 'ការ Sync កំពុងដំណើរការរួចហើយ...' };
  }

  isSyncing = true;
  try {
    const queue = getOfflineQueue();
    
    // Setup online check to avoid blocking
    if (!navigator.onLine) {
      isSyncing = false;
      return { success: false, syncedCount: 0, error: 'ដាច់បណ្តាញអ៊ីនធឺណិត (Offline)' };
    }

    // A. If we have items in the queue, process them sequentially
    if (queue.length > 0) {
      let successfullySynced = 0;
      const remainingQueue: QueueItem[] = [];

      for (const item of queue) {
        let ok = false;
        try {
          if (item.action === 'add' && item.record) {
            ok = await syncRecordToGoogleSheet(item.record, user);
          } else if (item.action === 'update' && item.record) {
            ok = await updateRecordInGoogleSheet(item.record, user);
          } else if (item.action === 'delete') {
            ok = await deleteRecordFromGoogleSheet(item.id, user);
          }
        } catch (err) {
          // console.error('Failed to sync queue item:', item, err);
          ok = false;
        }

        if (ok) {
          successfullySynced++;
        } else {
          remainingQueue.push(item);
        }
      }

      saveOfflineQueue(remainingQueue);
      isSyncing = false;

      // Update last sync time
      if (successfullySynced > 0) {
        localStorage.setItem(KEYS.LAST_SYNC, String(Date.now()));
      }

      return {
        success: remainingQueue.length === 0,
        syncedCount: successfullySynced,
        error: remainingQueue.length > 0 ? 'ការ Sync ខ្លះត្រូវបានបង្អង់ទុកជួរ (Queue)' : undefined
      };
    }

    // B. If no offline queue but we have local records and we want to perform a full overwrite/upload to spreadsheet
    if (localRecords !== undefined) {
      const response = await fetch(apiUrl, {
        method: 'POST',
        mode: 'cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'syncAll',
          data: localRecords,
          user: user || 'Sys'
        })
      });

      if (response.ok) {
        const resData = await response.json();
        if (resData.status === 'success') {
          localStorage.setItem(KEYS.LAST_SYNC, String(Date.now()));
          isSyncing = false;
          return { success: true, syncedCount: localRecords.length };
        }
      }
    }

    isSyncing = false;
    return { success: true, syncedCount: 0 };

  } catch (err: any) {
    isSyncing = false;
    return { success: false, syncedCount: 0, error: err.message || 'កំហុសពេលកំពុង Sync' };
  }
}

export function getLastSyncTime(): string {
  const timestamp = localStorage.getItem(KEYS.LAST_SYNC);
  if (!timestamp) return 'មិនទាន់ធ្លាប់ Sync';
  const d = new Date(Number(timestamp));
  return d.toLocaleTimeString('en-US', { hour12: false }) + ' ' + d.toLocaleDateString();
}

export async function syncAppConfigToGoogleSheet(config: any): Promise<boolean> {
  const apiUrl = getSyncUrl();
  const mode = getSyncMode();
  if (!apiUrl || mode === 'local') return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'syncSettings',
        data: config,
        user: 'Sys'
      })
    });
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    // console.warn('Failed to sync config to Google Sheet', e);
    return false;
  }
}

export async function syncUsersToGoogleSheet(users: any): Promise<boolean> {
  const apiUrl = getSyncUrl();
  const mode = getSyncMode();
  if (!apiUrl || mode === 'local') return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'syncUsers',
        data: users,
        user: 'Sys'
      })
    });
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    return false;
  }
}

export async function syncCategoriesToGoogleSheet(categories: any): Promise<boolean> {
  const apiUrl = getSyncUrl();
  const mode = getSyncMode();
  if (!apiUrl || mode === 'local') return false;

  try {
    const response = await fetch(apiUrl, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        action: 'syncCategories',
        data: categories,
        user: 'Sys'
      })
    });
    if (!response.ok) return false;
    const resData = await response.json();
    return resData.status === 'success';
  } catch (e) {
    return false;
  }
}
