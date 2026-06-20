const fs = require('fs');
let code = fs.readFileSync('./src/App.tsx', 'utf-8');

// 1. Add confirmState and askConfirm to App.tsx
code = code.replace(
  "const [toastMsg, setToastMsg] = useState('');",
  "const [toastMsg, setToastMsg] = useState('');\n  const [confirmState, setConfirmState] = useState<{message: string, onConfirm: () => void} | null>(null);\n\n  const askConfirm = (message: string, onConfirm: () => void) => { setConfirmState({ message, onConfirm }); };"
);

// 2. Pass askConfirm to the Views inside App.tsx
code = code.replace(
  "{activeTab === 'dashboard' && <DashboardView records={records} session={session} currency={config.currency} />}",
  "{activeTab === 'dashboard' && <DashboardView records={records} session={session} currency={config.currency} askConfirm={askConfirm} />}"
);
code = code.replace(
  "records={records} users={users} showToast={showToast} setActiveTab={setActiveTab} />}",
  "records={records} users={users} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} />}"
);
// Make sure all views get it if needed. Actually let's just do a blanket prop add for all activeTab matches in App
// Wait, I can just sequentially string replace the view usages.

let appReplaceStr = `
        {activeTab === 'dashboard' && <DashboardView records={records} session={session} currency={config.currency} />}
        {activeTab === 'records' && <RecordsView records={records} session={session} currency={config.currency} onUpdate={refreshData} showToast={showToast} users={users} onEdit={(r: any) => { setEditRecord(r); setActiveTab('add'); }} onDeleteTrigger={handleDeleteRecord} />}
        {activeTab === 'add' && <AddRecordView session={session} categories={categories} onUpdate={refreshData} showToast={showToast} setActiveTab={setActiveTab} editRecord={editRecord} users={users} />}
        {activeTab === 'settings' && <SettingsView session={session} setActiveTab={setActiveTab} config={config} onUpdate={refreshData} showToast={showToast} records={records} users={users} askConfirm={askConfirm} />}
        {activeTab === 'import' && <ImportView session={session} onUpdate={refreshData} showToast={showToast} setActiveTab={setActiveTab} config={config} users={users} />}
        {activeTab === 'export' && <ExportView records={records} users={users} showToast={showToast} setActiveTab={setActiveTab} />}
        {activeTab === 'data' && <DataCenterView session={session} setActiveTab={setActiveTab} />}
        {activeTab === 'members' && <MembersView users={users} session={session} onUpdate={refreshData} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} />}
        {activeTab === 'categories' && <CategoriesView categories={categories} onUpdate={refreshData} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} />}
        {activeTab === 'googleSheetSync' && <GoogleSheetSyncView session={session} records={records} onUpdate={refreshData} showToast={showToast} setActiveTab={setActiveTab} askConfirm={askConfirm} />}
`;

code = code.replace(/\{activeTab === 'dashboard'[\s\S]*?\{activeTab === 'googleSheetSync'[^}]+\}\n/, appReplaceStr.trim() + '\n');

// 3. Render the Confirm Modal in App
const modalCode = `
      {/* Confirm Notification */}
      {confirmState && (
        <div className="absolute inset-0 z-[100] bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4 min-h-screen">
          <div className="bg-white rounded-[32px] p-6 w-full max-w-[340px] shadow-2xl animate-in zoom-in-95 flex flex-col gap-4 font-sans border border-[#E8EEE9]">
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
              <button onClick={() => setConfirmState(null)} className="flex-1 py-3.5 bg-[#F5F7F5] hover:bg-[#E8EEE9] text-slate-600 outline-none rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-colors">បោះបង់ (Cancel)</button>
              <button 
                onClick={() => {
                  confirmState.onConfirm();
                  setConfirmState(null);
                }} 
                className="flex-1 py-3.5 bg-red-500 hover:bg-red-600 text-white outline-none rounded-xl font-extrabold text-[11px] uppercase tracking-wider transition-colors shadow-lg shadow-red-500/30">
                យល់ព្រមលុប (Delete)
              </button>
            </div>
          </div>
        </div>
      )}
`;

code = code.replace("{/* Toast Notification */}", modalCode + "\n      {/* Toast Notification */}");

// 4. Update the handleDeleteRecord in App
code = code.replace(
  "    if (window.confirm('តើអ្នកពិតជាចង់លុបប្រតិបត្តិការនេះមែនទេ? (Are you sure you want to delete this record?)')) {\n      const updated = records.filter(r => r.id !== id);\n      saveRecords(updated);\n      refreshData();\n      showToast('បានលុបដោយជោគជ័យ! (Deleted successfully)');\n    }",
  "    askConfirm('តើអ្នកពិតជាចង់លុបប្រតិបត្តិការនេះមែនទេ? (Are you sure you want to delete this record?)', () => {\n      const updated = records.filter(r => r.id !== id);\n      saveRecords(updated);\n      refreshData();\n      showToast('បានលុបដោយជោគជ័យ! (Deleted successfully)');\n    });"
);

// 5. Update other confirm usages

// In SettingsView:
code = code.replace(
  "if (confirm('តើអ្នកពិតជាចង់កំណត់ការកំណត់ (Settings) ទាំងអស់ទៅដើមវិញមែនទេ?'))",
  "askConfirm('តើអ្នកពិតជាចង់កំណត់ការកំណត់ (Settings) ទាំងអស់ទៅដើមវិញមែនទេ?', () => "
);
code = code.replace(
/askConfirm\('តើអ្នកពិតជាចង់កំណត់ការកំណត់ \(Settings\) ទាំងអស់ទៅដើមវិញមែនទេ\?', \(\) => \{\n\s*const defaultConf = [^;]+;\n\s*saveConfig\(defaultConf\);\n\s*onUpdate\(\);\n\s*showToast\('ការកំណត់ត្រូវបានជម្រះ! 🧹'\);\n\s*\}\)/g,
(match) => match.replace(/}$/, "});")
);

// We should use regex properly
code = code.replace(
  /if \(confirm\('តើអ្នកពិតជាចង់កំណត់ការកំណត់ \(Settings\) ទាំងអស់ទៅដើមវិញមែនទេ\?'\)\) \{([\s\S]*?showToast\('ការកំណត់ត្រូវបានជម្រះ! 🧹'\);\s*)\}/,
  "askConfirm('តើអ្នកពិតជាចង់កំណត់ការកំណត់ (Settings) ទាំងអស់ទៅដើមវិញមែនទេ?', () => {\n$1});"
);

code = code.replace(
  /if \(confirm\('⚠️ គ្រោះថ្នាក់៖ តើអ្នកពិតជាចង់លុបប្រតិបត្តិការកត់ត្រាទាំងអស់មែនទេ\? សកម្មភាពនេះមិនអាចសង្គ្រោះវិញបានឡើយ!'\)\) \{([\s\S]*?showToast\('លុបទិន្នន័យបានជោគជ័យ! 🗑️'\);\s*)\}/,
  "askConfirm('⚠️ គ្រោះថ្នាក់៖ តើអ្នកពិតជាចង់លុបប្រតិបត្តិការកត់ត្រាទាំងអស់មែនទេ? សកម្មភាពនេះមិនអាចសង្គ្រោះវិញបានឡើយ!', () => {\n$1});"
);

code = code.replace(
  /if \(confirm\('🚨 គ្រោះថ្នាក់បំផុត៖ តើអ្នកពិតជានឹងលុបទិន្នន័យកម្មវិធីទាំងអស់[^']+'\)\) \{([\s\S]*?window.location.reload\(\);\s*)\}/,
  "askConfirm('🚨 គ្រោះថ្នាក់បំផុត៖ តើអ្នកពិតជានឹងលុបទិន្នន័យកម្មវិធីទាំងអស់មែនទេ? កម្មវិធីនឹងចាប់ផ្តើមជាថ្មីទាំងស្រុង!', () => {\n$1});"
);

// Update destructured props to include askConfirm
code = code.replace(
  "function SettingsView({ session, setActiveTab, config, onUpdate, showToast, records, users }: any) {",
  "function SettingsView({ session, setActiveTab, config, onUpdate, showToast, records, users, askConfirm }: any) {"
);

code = code.replace(
  "function MembersView({ users, session, onUpdate, showToast, setActiveTab }: any) {",
  "function MembersView({ users, session, onUpdate, showToast, setActiveTab, askConfirm }: any) {"
);

code = code.replace(
  /if \(confirm\(`តើអ្នកពិតជាចង់លុបគណនីគ្រួសារ "\$\{userName\}" មែនទេ\? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានទេ។`\)\) \{([\s\S]*?showToast[^\n]+;\s*)\}/,
  "askConfirm(`តើអ្នកពិតជាចង់លុបគណនីគ្រួសារ \"${userName}\" មែនទេ?`, () => {\n$1});"
);

code = code.replace(
  "function CategoriesView({ categories, onUpdate, showToast, setActiveTab }: any) {",
  "function CategoriesView({ categories, onUpdate, showToast, setActiveTab, askConfirm }: any) {"
);

code = code.replace(
  /if\(confirm\('តើអ្នកចង់លុបប្រភេទនេះមែនទេ\?'\)\) \{([\s\S]*?onUpdate\(\);\s*)\}/,
  "askConfirm('តើអ្នកចង់លុបប្រភេទនេះមែនទេ?', () => {\n$1});"
);

code = code.replace(
  "function GoogleSheetSyncView({ session, records, onUpdate, showToast, setActiveTab }: any) {",
  "function GoogleSheetSyncView({ session, records, onUpdate, showToast, setActiveTab, askConfirm }: any) {"
);

code = code.replace(
  /if \(confirm\('តើអ្នកពិតជាចង់សម្អាត Offline Queue ដែលមិនទាន់ Sync មែនទេ\?'\)\) \{([\s\S]*?showToast[^\n]+;\s*)\}/,
  "askConfirm('តើអ្នកពិតជាចង់សម្អាត Offline Queue ដែលមិនទាន់ Sync មែនទេ?', () => {\n$1});"
);

fs.writeFileSync('./src/App.tsx', code);
