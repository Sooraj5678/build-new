import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import logoImg from '@assets/logo.png';

declare global {
  interface Window {
    electronAPI?: {
      isElectron: boolean;
      loadSettings: () => Promise<HeaderSettings>;
      saveSettings: (s: HeaderSettings) => Promise<boolean>;
      printReport: () => Promise<boolean>;
      savePDF: (html: string) => Promise<{ success: boolean; filePath?: string; error?: string }>;
    };
  }
}

// ── Types ────────────────────────────────────────────────────

interface PartyRow {
  partyName: string;
  mobile: string;
  city: string;
  amount: number;
}

interface CityGroup {
  city: string;
  parties: PartyRow[];
  total: number;
}

interface HeaderSettings {
  companyName: string;
  addressLine: string;
  mobile1: string;
  mobile2: string;
  reportTitle: string;
  fromDate: string;
  toDate: string;
}

const DEFAULT_SETTINGS: HeaderSettings = {
  companyName: '',
  addressLine: '',
  mobile1: '',
  mobile2: '',
  reportTitle: 'City Wise Outstanding Receivable',
  fromDate: '',
  toDate: '',
};

// ── Excel Parsing ────────────────────────────────────────────

function normHead(h: string) {
  return String(h).trim().toLowerCase().replace(/[\s_\-/]+/g, '');
}

const PARTY_KEYS  = ['partyname','name','party','customername','accountname','ledgername'];
const MOBILE_KEYS = ['mobile','mobilenumber','phone','phonenumber','contact','contactnumber','mob'];
const CITY_KEYS   = ['address','area','city','addressareacity','addressarea','areacity','location','place'];
const AMOUNT_KEYS = ['outstandingamount','outstanding','amount','balance','dueamount','totalamount','netamount','payable','receivable','closingbalance','duebalance'];

function findCol(headers: string[], keys: string[]): string | null {
  for (const h of headers) {
    const n = normHead(h);
    if (keys.some(k => n.includes(k) || k.includes(n))) return h;
  }
  return null;
}

function normCity(raw: unknown): string {
  if (!raw || String(raw).trim() === '') return 'UNKNOWN';
  return String(raw).trim().toUpperCase();
}

function parseExcel(buf: ArrayBuffer): PartyRow[] {
  const wb = XLSX.read(buf, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows: Record<string, unknown>[] = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (!rows.length) throw new Error('No data found in the Excel file.');

  const headers = Object.keys(rows[0]);
  const partyKey  = findCol(headers, PARTY_KEYS);
  const mobileKey = findCol(headers, MOBILE_KEYS);
  const cityKey   = findCol(headers, CITY_KEYS);
  const amountKey = findCol(headers, AMOUNT_KEYS);

  if (!partyKey)  throw new Error('Could not find Party Name column. Expected: Party Name, Name, Account Name.');
  if (!amountKey) throw new Error('Could not find Outstanding Amount column. Expected: Outstanding Amount, Balance, Amount.');

  return rows
    .map(r => ({
      partyName: String(r[partyKey] ?? '').trim(),
      mobile:    mobileKey ? String(r[mobileKey] ?? '').trim() : '',
      city:      normCity(cityKey ? r[cityKey] : null),
      amount:    parseFloat(String(r[amountKey] ?? '0').replace(/[^0-9.\-]/g, '')) || 0,
    }))
    .filter(r => r.partyName !== '');
}

function groupByCity(rows: PartyRow[]): CityGroup[] {
  const map = new Map<string, PartyRow[]>();
  for (const r of rows) {
    if (!map.has(r.city)) map.set(r.city, []);
    map.get(r.city)!.push(r);
  }
  return Array.from(map.entries())
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([city, parties]) => ({
      city, parties,
      total: parties.reduce((s, p) => s + p.amount, 0),
    }));
}

// ── Formatters ───────────────────────────────────────────────

function fmt(n: number) {
  return n.toLocaleString('en-IN', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function today() {
  return new Date().toLocaleDateString('en-IN', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// ── Report HTML (for PDF & print window) ────────────────────

function buildReportHTML(groups: CityGroup[], s: HeaderSettings): string {
  const grandTotal = groups.reduce((acc, g) => acc + g.total, 0);

  const groupsHTML = groups.map(g => `
    <div style="margin-bottom:20px;page-break-inside:avoid;">
      <div style="font-weight:bold;font-size:13px;letter-spacing:1px;font-family:'Courier New',monospace;margin-bottom:3px;">${g.city}</div>
      <table style="width:100%;border-collapse:collapse;">
        <thead>
          <tr style="background:#e8e8e8;">
            <th style="padding:3px 6px;border:1px solid #999;font-family:'Courier New',monospace;font-size:12px;text-align:left;">Party Name</th>
            <th style="padding:3px 6px;border:1px solid #999;font-family:'Courier New',monospace;font-size:12px;text-align:left;width:150px;">Mobile</th>
            <th style="padding:3px 6px;border:1px solid #999;font-family:'Courier New',monospace;font-size:12px;text-align:right;width:110px;">Amount</th>
          </tr>
        </thead>
        <tbody>
          ${g.parties.map(p => `
          <tr>
            <td style="padding:2px 6px;border:1px solid #ccc;font-family:'Courier New',monospace;font-size:12px;">${p.partyName}</td>
            <td style="padding:2px 6px;border:1px solid #ccc;font-family:'Courier New',monospace;font-size:12px;">${p.mobile}</td>
            <td style="padding:2px 6px;border:1px solid #ccc;font-family:'Courier New',monospace;font-size:12px;text-align:right;">${fmt(p.amount)}</td>
          </tr>`).join('')}
        </tbody>
        <tfoot>
          <tr style="background:#f0f0f0;">
            <td colspan="2" style="padding:3px 6px;border:1px solid #999;font-family:'Courier New',monospace;font-size:12px;font-weight:bold;">City Total</td>
            <td style="padding:3px 6px;border:1px solid #999;font-family:'Courier New',monospace;font-size:12px;font-weight:bold;text-align:right;">${fmt(g.total)}</td>
          </tr>
        </tfoot>
      </table>
    </div>`).join('');

  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${s.reportTitle || 'Outstanding Report'}</title>
<style>@page{size:A4 portrait;margin:15mm}body{font-family:'Courier New',Courier,monospace;font-size:12px;background:white;color:#000}</style>
</head><body>
  <div style="text-align:center;margin-bottom:18px;border-bottom:2px solid #333;padding-bottom:12px;">
    ${s.companyName ? `<div style="font-size:16px;font-weight:bold;letter-spacing:2px;font-family:'Courier New',monospace;">${s.companyName}</div>` : ''}
    ${s.addressLine ? `<div style="font-size:12px;font-family:'Courier New',monospace;margin-top:3px;">${s.addressLine}</div>` : ''}
    ${(s.mobile1 || s.mobile2) ? `<div style="font-size:12px;font-family:'Courier New',monospace;margin-top:2px;">${[s.mobile1 && `Mo.${s.mobile1}`, s.mobile2 && `Mo.${s.mobile2}`].filter(Boolean).join('&nbsp;&nbsp;&nbsp;')}</div>` : ''}
    <div style="font-size:14px;font-weight:bold;margin-top:8px;font-family:'Courier New',monospace;">${s.reportTitle || 'City Wise Outstanding Receivable'}</div>
    <div style="font-size:11px;margin-top:4px;font-family:'Courier New',monospace;color:#444;">
      ${(s.fromDate || s.toDate) ? `From Date: ${s.fromDate || '__/__/____'}&nbsp;&nbsp;&nbsp;To&nbsp;&nbsp;&nbsp;${s.toDate || '__/__/____'}` : `Date: ${today()}`}
    </div>
  </div>
  ${groupsHTML}
  <div style="border-top:2px solid #333;padding-top:6px;margin-top:12px;">
    <table style="width:100%;border-collapse:collapse;">
      <tr style="background:#d0d0d0;">
        <td colspan="2" style="padding:5px 6px;border:1px solid #666;font-family:'Courier New',monospace;font-size:13px;font-weight:bold;">GRAND TOTAL</td>
        <td style="padding:5px 6px;border:1px solid #666;font-family:'Courier New',monospace;font-size:13px;font-weight:bold;text-align:right;width:110px;">${fmt(grandTotal)}</td>
      </tr>
    </table>
  </div>
</body></html>`;
}

// ── App Component ────────────────────────────────────────────

const MONO: React.CSSProperties = { fontFamily: "'Courier New', Courier, monospace" };

function Field({ label, value, onChange, placeholder, type = 'text' }: {
  label: string; value: string; onChange: (v: string) => void;
  placeholder?: string; type?: string;
}) {
  return (
    <div>
      <div style={{ fontSize: '11px', color: '#666', marginBottom: '3px', ...MONO }}>{label}</div>
      <input type={type} value={value} placeholder={placeholder}
        onChange={e => onChange(e.target.value)}
        style={{ width: '100%', padding: '5px 8px', border: '1px solid #bbb', ...MONO, fontSize: '12px', background: '#fafafa', outline: 'none' }}
      />
    </div>
  );
}

export default function App() {
  const [settings, setSettings]   = useState<HeaderSettings>(DEFAULT_SETTINGS);
  const [groups, setGroups]       = useState<CityGroup[]>([]);
  const [fileName, setFileName]   = useState('');
  const [error, setError]         = useState('');
  const [loading, setLoading]     = useState(false);
  const [stats, setStats]         = useState<{ cities: number; parties: number; total: number } | null>(null);
  const [showSettings, setShowSettings] = useState(false);
  const [pdfMsg, setPdfMsg]       = useState('');
  const fileRef = useRef<HTMLInputElement>(null);
  const isElectron = !!window.electronAPI?.isElectron;

  // Load settings on mount
  useEffect(() => {
    if (isElectron) {
      window.electronAPI!.loadSettings().then(s => setSettings(s));
    } else {
      try {
        const raw = localStorage.getItem('vyapar_settings');
        if (raw) setSettings({ ...DEFAULT_SETTINGS, ...JSON.parse(raw) });
      } catch {}
    }
  }, [isElectron]);

  function persistSettings(s: HeaderSettings) {
    setSettings(s);
    if (isElectron) window.electronAPI!.saveSettings(s);
    else localStorage.setItem('vyapar_settings', JSON.stringify(s));
  }

  function setField<K extends keyof HeaderSettings>(k: K, v: string) {
    persistSettings({ ...settings, [k]: v });
  }

  // ── File handling ──────────────────────────────────────────

  function loadFile(file: File) {
    setError(''); setLoading(true); setFileName(file.name);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const rows    = parseExcel(e.target!.result as ArrayBuffer);
        const grouped = groupByCity(rows);
        setGroups(grouped);
        setStats({ cities: grouped.length, parties: rows.length, total: rows.reduce((s, r) => s + r.amount, 0) });
      } catch (err: unknown) {
        setError(err instanceof Error ? err.message : 'Failed to parse file.');
        setGroups([]); setStats(null);
      } finally { setLoading(false); }
    };
    reader.onerror = () => { setError('Could not read file.'); setLoading(false); };
    reader.readAsArrayBuffer(file);
  }

  function onFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const f = e.target.files?.[0]; if (f) loadFile(f); e.target.value = '';
  }

  function onDrop(e: React.DragEvent) {
    e.preventDefault(); const f = e.dataTransfer.files?.[0]; if (f) loadFile(f);
  }

  // ── PDF / Print ────────────────────────────────────────────

  async function handlePDF() {
    if (!groups.length) return;
    const html = buildReportHTML(groups, settings);
    if (isElectron) {
      setPdfMsg('Saving...');
      const res = await window.electronAPI!.savePDF(html);
      if (res.success) {
        setPdfMsg(`✔ Saved`);
        setTimeout(() => setPdfMsg(''), 4000);
      } else {
        setPdfMsg(res.error === 'Cancelled' ? '' : `✘ ${res.error}`);
        setTimeout(() => setPdfMsg(''), 4000);
      }
    } else {
      const w = window.open('', '_blank', 'width=900,height=700');
      if (!w) { alert('Pop-up blocked. Please allow pop-ups for this site.'); return; }
      w.document.open(); w.document.write(html); w.document.close();
      w.focus(); setTimeout(() => w.print(), 600);
    }
  }

  async function handlePrint() {
    if (!groups.length) return;
    if (isElectron) {
      await window.electronAPI!.printReport();
    } else {
      const html = buildReportHTML(groups, settings);
      const w = window.open('', '_blank', 'width=900,height=700');
      if (!w) { alert('Pop-up blocked. Please allow pop-ups for this site.'); return; }
      w.document.open(); w.document.write(html); w.document.close();
      w.focus(); setTimeout(() => w.print(), 600);
    }
  }

  function reset() { setGroups([]); setFileName(''); setError(''); setStats(null); setPdfMsg(''); }

  // ── Derived ────────────────────────────────────────────────

  const hasReport  = groups.length > 0;
  const grandTotal = groups.reduce((s, g) => s + g.total, 0);

  // ── Styles ─────────────────────────────────────────────────

  const btnBase: React.CSSProperties  = { ...MONO, border: 'none', cursor: 'pointer', fontSize: '12px', padding: '6px 16px', letterSpacing: '0.3px' };
  const btnPrimary: React.CSSProperties = { ...btnBase, background: '#1a3a5c', color: '#fff' };
  const btnGreen: React.CSSProperties  = { ...btnBase, background: '#2d5a27', color: '#fff' };
  const btnOutline: React.CSSProperties = { ...btnBase, background: '#fff', color: '#333', border: '1px solid #999' };

  // ── Render ─────────────────────────────────────────────────

  return (
    <div style={{ minHeight: '100vh', background: '#f0ede8', ...MONO }}>

      {/* ── Logo Banner ─────────────────────────────── */}
      <div style={{ background: '#fff', borderBottom: '1px solid #ddd', padding: '4px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <img src={logoImg} alt="Vyapar Outstanding Formatter"
          style={{ height: '86px', width: 'auto', objectFit: 'contain', maxWidth: '340px' }} />
        <div style={{ fontSize: '11px', color: '#888', textAlign: 'right', lineHeight: '1.6' }}>
          City-wise Outstanding Report Generator<br />
          <span style={{ fontSize: '10px' }}>Offline · Client-side · No Upload Required</span>
        </div>
      </div>

      {/* ── Action Bar ──────────────────────────────── */}
      <div style={{ background: '#1a3a5c', padding: '6px 20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '8px' }}>
        <div style={{ fontSize: '11px', color: 'rgba(255,255,255,0.6)' }}>
          {isElectron ? '🖥 Desktop App' : '🌐 Web App'}
        </div>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button onClick={() => setShowSettings(s => !s)}
            style={{ ...btnBase, background: showSettings ? '#2d5a8c' : 'transparent', border: '1px solid rgba(255,255,255,0.5)', color: '#fff', padding: '4px 14px' }}>
            ⚙ Settings
          </button>
          {hasReport && (
            <button onClick={reset}
              style={{ ...btnBase, background: 'transparent', border: '1px solid rgba(255,255,255,0.4)', color: '#fff', padding: '4px 12px' }}>
              ✕ New File
            </button>
          )}
        </div>
      </div>

      <div style={{ maxWidth: '940px', margin: '0 auto', padding: '16px' }}>

        {/* ── Settings Panel ──────────────────────── */}
        {showSettings && (
          <div style={{ background: '#fff', border: '1px solid #bbb', padding: '16px', marginBottom: '14px' }}>
            <div style={{ fontWeight: 'bold', fontSize: '13px', marginBottom: '12px', borderBottom: '1px solid #ddd', paddingBottom: '6px' }}>
              Report Header Settings
              <span style={{ fontWeight: 'normal', fontSize: '11px', color: '#888', marginLeft: '8px' }}>
                ({isElectron ? 'saved to disk — auto-loads on next open' : 'saved in browser storage'})
              </span>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
              <Field label="Company Name"    value={settings.companyName}  onChange={v => setField('companyName', v)}  placeholder="e.g. NURMOHMED SULEMANBHAI MEMAN" />
              <Field label="Address Line"    value={settings.addressLine}  onChange={v => setField('addressLine', v)}  placeholder="e.g. NATIONAL HIGHWAY ROAD THARA" />
              <Field label="Mobile Number 1" value={settings.mobile1}      onChange={v => setField('mobile1', v)}      placeholder="e.g. 9898257361" />
              <Field label="Mobile Number 2" value={settings.mobile2}      onChange={v => setField('mobile2', v)}      placeholder="e.g. 9724800088" />
              <Field label="Report Title"    value={settings.reportTitle}  onChange={v => setField('reportTitle', v)}  placeholder="City Wise Outstanding Receivable" />
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <Field label="From Date" value={settings.fromDate} onChange={v => setField('fromDate', v)} type="date" />
                <Field label="To Date"   value={settings.toDate}   onChange={v => setField('toDate', v)}   type="date" />
              </div>
            </div>
            {/* Header Preview */}
            {(settings.companyName || settings.reportTitle) && (
              <div style={{ marginTop: '12px', padding: '10px 14px', background: '#f8f8f0', border: '1px dashed #ccc', textAlign: 'center', fontSize: '12px' }}>
                {settings.companyName && <div style={{ fontWeight: 'bold', fontSize: '13px', letterSpacing: '1px' }}>{settings.companyName}</div>}
                {settings.addressLine && <div>{settings.addressLine}</div>}
                {(settings.mobile1 || settings.mobile2) && <div>{[settings.mobile1 && `Mo.${settings.mobile1}`, settings.mobile2 && `Mo.${settings.mobile2}`].filter(Boolean).join('   ')}</div>}
                <div style={{ fontWeight: 'bold', marginTop: '4px' }}>{settings.reportTitle || 'City Wise Outstanding Receivable'}</div>
                <div style={{ fontSize: '11px', color: '#555', marginTop: '3px' }}>
                  {(settings.fromDate || settings.toDate)
                    ? `From Date: ${settings.fromDate || '__/__/____'}   To   ${settings.toDate || '__/__/____'}`
                    : `Date: ${today()}`}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ── Upload Panel ────────────────────────── */}
        {!hasReport && (
          <>
            <div onDrop={onDrop} onDragOver={e => e.preventDefault()} onClick={() => fileRef.current?.click()}
              style={{ background: '#fff', border: '2px dashed #aaa', padding: '36px 30px', textAlign: 'center', marginBottom: '12px', cursor: 'pointer' }}>
              <div style={{ fontSize: '36px', marginBottom: '8px' }}>📊</div>
              <div style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '5px' }}>Upload Vyapar Excel Export</div>
              <div style={{ fontSize: '12px', color: '#666', marginBottom: '18px' }}>
                Drag & drop your .xlsx / .xls file here, or click to browse
              </div>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={onFileChange} />
              <button onClick={e => { e.stopPropagation(); fileRef.current?.click(); }} disabled={loading}
                style={{ ...btnPrimary, padding: '10px 28px', opacity: loading ? 0.7 : 1 }}>
                {loading ? 'PROCESSING...' : '📂  UPLOAD EXCEL FILE'}
              </button>
              {fileName && loading && (
                <div style={{ marginTop: '10px', fontSize: '12px', color: '#555' }}>
                  Processing <strong>{fileName}</strong>…
                </div>
              )}
            </div>
            <div style={{ background: '#fffbf0', border: '1px solid #ddd', padding: '10px 14px', fontSize: '11px', color: '#666' }}>
              <strong>Auto-detected columns:</strong> Party Name · Mobile/Phone · Address/Area/City · Outstanding Amount/Balance
            </div>
          </>
        )}

        {/* ── Error ───────────────────────────────── */}
        {error && (
          <div style={{ background: '#fef2f2', border: '1px solid #f87171', padding: '10px 14px', marginBottom: '12px', fontSize: '12px', color: '#b91c1c' }}>
            <strong>Error:</strong> {error}
            <button onClick={reset} style={{ ...btnBase, marginLeft: '12px', padding: '2px 10px', border: '1px solid #f87171', color: '#b91c1c', background: 'transparent' }}>
              Try Again
            </button>
          </div>
        )}

        {/* ── Report Action Bar ────────────────────── */}
        {hasReport && (
          <div style={{ background: '#fff', border: '1px solid #ccc', padding: '10px 14px', marginBottom: '12px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
            <div style={{ fontSize: '12px' }}>
              <span style={{ marginRight: '14px' }}>📁 <strong>{fileName}</strong></span>
              <span style={{ marginRight: '12px' }}>Cities: <strong>{stats?.cities}</strong></span>
              <span style={{ marginRight: '12px' }}>Parties: <strong>{stats?.parties}</strong></span>
              <span>Total: <strong>₹{stats ? fmt(stats.total) : 0}</strong></span>
            </div>
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              {pdfMsg && (
                <span style={{ fontSize: '11px', color: pdfMsg.startsWith('✔') ? '#2d5a27' : '#b91c1c' }}>{pdfMsg}</span>
              )}
              <button onClick={() => fileRef.current?.click()} style={btnOutline}>📂 New File</button>
              <button onClick={handlePDF}   style={btnPrimary}>📄 Save as PDF</button>
              <button onClick={handlePrint} style={btnGreen}>🖨️ Print</button>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={onFileChange} />
            </div>
          </div>
        )}

        {/* ── Report Preview ───────────────────────── */}
        {hasReport && (
          <div style={{ background: '#fff', border: '1px solid #ccc', padding: '28px 32px', fontSize: '12px', lineHeight: '1.5' }}>

            {/* Header */}
            <div style={{ textAlign: 'center', marginBottom: '16px', borderBottom: '2px solid #333', paddingBottom: '10px' }}>
              {settings.companyName && <div style={{ fontSize: '15px', fontWeight: 'bold', letterSpacing: '2px' }}>{settings.companyName}</div>}
              {settings.addressLine && <div style={{ fontSize: '12px' }}>{settings.addressLine}</div>}
              {(settings.mobile1 || settings.mobile2) && (
                <div style={{ fontSize: '12px' }}>
                  {[settings.mobile1 && `Mo.${settings.mobile1}`, settings.mobile2 && `Mo.${settings.mobile2}`].filter(Boolean).join('   ')}
                </div>
              )}
              <div style={{ fontSize: '13px', fontWeight: 'bold', marginTop: '6px' }}>
                {settings.reportTitle || 'City Wise Outstanding Receivable'}
              </div>
              <div style={{ fontSize: '11px', marginTop: '4px', color: '#555' }}>
                {(settings.fromDate || settings.toDate)
                  ? `From Date: ${settings.fromDate || '__/__/____'}   To   ${settings.toDate || '__/__/____'}`
                  : `Date: ${today()}`}
              </div>
            </div>

            {/* City Groups */}
            {groups.map(g => (
              <div key={g.city} style={{ marginBottom: '18px' }}>
                <div style={{ fontWeight: 'bold', fontSize: '13px', letterSpacing: '1px', marginBottom: '3px' }}>{g.city}</div>
                <div style={{ borderBottom: '1px solid #555', marginBottom: '3px' }} />
                <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '2px' }}>
                  <thead>
                    <tr style={{ background: '#eeeeee' }}>
                      <th style={{ padding: '3px 6px', border: '1px solid #bbb', textAlign: 'left', fontWeight: 'bold' }}>Party Name</th>
                      <th style={{ padding: '3px 6px', border: '1px solid #bbb', textAlign: 'left', fontWeight: 'bold', width: '150px' }}>Mobile</th>
                      <th style={{ padding: '3px 6px', border: '1px solid #bbb', textAlign: 'right', fontWeight: 'bold', width: '110px' }}>Amount</th>
                    </tr>
                  </thead>
                  <tbody>
                    {g.parties.map((p, i) => (
                      <tr key={i} style={{ background: i % 2 === 0 ? '#fff' : '#fafafa' }}>
                        <td style={{ padding: '2px 6px', border: '1px solid #e0e0e0' }}>{p.partyName}</td>
                        <td style={{ padding: '2px 6px', border: '1px solid #e0e0e0', color: '#555' }}>{p.mobile}</td>
                        <td style={{ padding: '2px 6px', border: '1px solid #e0e0e0', textAlign: 'right' }}>{fmt(p.amount)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: '#f0f0f0' }}>
                      <td colSpan={2} style={{ padding: '3px 6px', border: '1px solid #aaa', fontWeight: 'bold' }}>City Total</td>
                      <td style={{ padding: '3px 6px', border: '1px solid #aaa', fontWeight: 'bold', textAlign: 'right' }}>{fmt(g.total)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            ))}

            {/* Grand Total */}
            <div style={{ borderTop: '2px solid #222', paddingTop: '6px', marginTop: '10px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <tbody>
                  <tr style={{ background: '#d8d8d8' }}>
                    <td colSpan={2} style={{ padding: '5px 6px', border: '1px solid #888', fontWeight: 'bold', fontSize: '13px' }}>GRAND TOTAL</td>
                    <td style={{ padding: '5px 6px', border: '1px solid #888', fontWeight: 'bold', fontSize: '13px', textAlign: 'right', width: '110px' }}>
                      {fmt(grandTotal)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
