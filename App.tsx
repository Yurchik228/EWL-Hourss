
import React, { useState, useCallback, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { ExcelService } from './services/excelService';
import { ComparisonService } from './services/comparisonService';
import { SheetAnalysis, CalculationResult, DayRange, CalculationSummary, CalculationStats, ComparisonItem, MasterRecord } from './types';
import { 
  FileUp, 
  Calculator, 
  Download, 
  Search, 
  Settings2, 
  AlertCircle, 
  CheckCircle2, 
  Layers,
  FileSpreadsheet,
  RefreshCw,
  CalendarDays,
  ChevronDown,
  ChevronRight,
  ShieldAlert,
  Frown,
  FileCode,
  Wallet,
  Coins,
  Settings,
  ClipboardList,
  Diff,
  ArrowRightLeft,
  XCircle,
  AlertTriangle,
  FileText,
  Trash2,
  Building2,
  TrendingUp,
  LayoutGrid,
  List,
  Eraser,
  UploadCloud,
  FileBox
} from 'lucide-react';

const HOURLY_RATE = 31.40;
const CLIENT_RATE = 49.92;

const App: React.FC = () => {
  const [analyses, setAnalyses] = useState<Map<string, SheetAnalysis>>(new Map());
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [sheetSegments, setSheetSegments] = useState<Map<string, number>>(new Map());
  const [results, setResults] = useState<CalculationResult[] | null>(null);
  const [calcStats, setCalcStats] = useState<CalculationStats | null>(null);
  const [loading, setLoading] = useState(false);
  const [expandedRow, setExpandedRow] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'calc' | 'verify'>('calc');
  const [viewMode, setViewMode] = useState<'flat' | 'grouped'>('grouped');
  const [status, setStatus] = useState<{ msg: string; type: 'info' | 'error' | 'success' }>({ 
    msg: 'Wgraj pliki Excel, aby rozpocząć.', 
    type: 'info' 
  });
  
  const [systemRawData, setSystemRawData] = useState('');
  const [masterRawData, setMasterRawData] = useState('');
  const [masterFilesData, setMasterFilesData] = useState<Map<string, MasterRecord[]>>(new Map());
  const [comparisonResults, setComparisonResults] = useState<ComparisonItem[] | null>(null);

  const [dayRange, setDayRange] = useState<DayRange>({ from: 1, to: 31 });
  const [includeSheetCol, setIncludeSheetCol] = useState(false);
  const [ignoreLineCol, setIgnoreLineCol] = useState(true); 
  const [minHours, setMinHours] = useState<number>(0.1);
  const [filterInactive, setFilterInactive] = useState<boolean>(true);
  const [searchQuery, setSearchQuery] = useState('');

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('pl-PL', { style: 'currency', currency: 'PLN' }).format(value);
  };

  const handleFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    
    setLoading(true);
    setStatus({ msg: 'Przetwarzanie plików...', type: 'info' });

    try {
      const fileArray = Array.from(files) as File[];
      const newAnalysesEntries: [string, SheetAnalysis][] = [];
      const newSelectedEntries: string[] = [];
      const newSegmentsEntries: [string, number][] = [];

      await Promise.all(fileArray.map(async (file: File) => {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        for (const sheetName of wb.SheetNames) {
          const analysis = ExcelService.analyzeSheet(wb, sheetName, file.name);
          const key = `${file.name} | ${sheetName}`;
          newAnalysesEntries.push([key, analysis]);
          if (analysis.ok) { 
            newSelectedEntries.push(key); 
            newSegmentsEntries.push([key, 1]); 
          }
        }
      }));

      setAnalyses(prev => {
        const next = new Map(prev);
        newAnalysesEntries.forEach(([k, v]) => next.set(k, v));
        return next;
      });

      setSelectedSheets(prev => {
        const next = new Set(prev);
        newSelectedEntries.forEach(k => next.add(k));
        return next;
      });

      setSheetSegments(prev => {
        const next = new Map(prev);
        newSegmentsEntries.forEach(([k, v]) => {
          if (!next.has(k)) next.set(k, v);
        });
        return next;
      });

      setStatus({ msg: `Dodano ${fileArray.length} plików. Gotowe do przeliczenia.`, type: 'success' });
    } catch (err: any) {
      setStatus({ msg: `Błąd wczytywania: ${err.message}`, type: 'error' });
    } finally { 
      setLoading(false); 
      if (e.target) e.target.value = '';
    }
  }, [analyses.size]);

  const handleMasterFileUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setLoading(true);
    const newData = new Map(masterFilesData);

    try {
      // FIX: Ensure 'file' is properly typed to resolve 'Property does not exist on type unknown' errors
      const fileArray = Array.from(files) as File[];
      await Promise.all(fileArray.map(async (file: File) => {
        const data = await file.arrayBuffer();
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
        
        // Convert grid to text format compatible with parseTextData
        const text = json.map(row => row.join("\t")).join("\n");
        const parsed = ComparisonService.parseTextData(text);
        newData.set(file.name, parsed);
      }));

      setMasterFilesData(newData);
      setStatus({ msg: `Wczytano ${files.length} plików Master do weryfikacji.`, type: 'success' });
    } catch (err: any) {
      setStatus({ msg: `Błąd wczytywania plików Master: ${err.message}`, type: 'error' });
    } finally {
      setLoading(false);
      if (e.target) e.target.value = '';
    }
  }, [masterFilesData]);

  const clearAllFiles = () => {
    setAnalyses(new Map());
    setSelectedSheets(new Set());
    setSheetSegments(new Map());
    setResults(null);
    setCalcStats(null);
    setStatus({ msg: 'Wyczyszczono wszystkie pliki.', type: 'info' });
  };

  const handleCalculate = () => {
    if (!analyses.size || !selectedSheets.size) return;
    setLoading(true); 
    try {
      const summary: CalculationSummary = ExcelService.calculate(
          analyses, 
          selectedSheets, 
          sheetSegments, 
          dayRange, 
          includeSheetCol, 
          ignoreLineCol,
          minHours,
          filterInactive
      );
      setResults(summary.results); 
      setCalcStats(summary.stats);
      setStatus({ msg: `Obliczenia zakończone sukcesem!`, type: 'success' });
    } catch (err: any) { 
      setStatus({ msg: `Błąd obliczeń: ${err.message}`, type: 'error' });
    } finally { setLoading(false); }
  };

  const handleVerify = () => {
    const hasCurrentResults = results && results.length > 0;
    const hasSystemPaste = systemRawData.trim().length > 0;
    
    let sysRecords: MasterRecord[] = [];
    if (hasSystemPaste) {
      sysRecords = ComparisonService.parseTextData(systemRawData);
    } else if (hasCurrentResults) {
      sysRecords = results!.map(r => ({ fullName: `${r.Imie} ${r.Nazwisko}`.trim(), hours: r.Razem }));
    } else {
      setStatus({ msg: "Brak danych systemowych. Wygeneruj raport lub wklej dane.", type: "error" });
      return;
    }

    // Master records source: either uploaded files or pasted text
    let masterRecords: MasterRecord[] = [];
    if (masterFilesData.size > 0) {
      masterFilesData.forEach(records => {
        masterRecords = masterRecords.concat(records);
      });
    } else if (masterRawData.trim().length > 0) {
      masterRecords = ComparisonService.parseTextData(masterRawData);
    } else {
      setStatus({ msg: "Wklej dane lub wgraj pliki od partnerów (Master).", type: "error" });
      return;
    }

    setLoading(true);
    try {
      const comparison = ComparisonService.compareData(sysRecords, masterRecords);
      setComparisonResults(comparison);
      setStatus({ msg: `Audyt zakończony. Znaleziono ${comparison.filter(c => c.status !== 'ok').length} rozbieżności.`, type: 'success' });
    } catch (err: any) {
      setStatus({ msg: `Błąd weryfikacji: ${err.message}`, type: 'error' });
    } finally {
      setLoading(false);
    }
  };

  const filteredResults = useMemo(() => {
    if (!results) return [];
    const q = searchQuery.toLowerCase();
    return results.filter(r => r.Imie.toLowerCase().includes(q) || r.Nazwisko.toLowerCase().includes(q) || r.Linia.toLowerCase().includes(q));
  }, [results, searchQuery]);

  const groupedResultsByFirm = useMemo(() => {
    const groups: Record<string, CalculationResult[]> = {};
    filteredResults.forEach(r => {
      const firmName = r.Sheet.split(' | ')[0];
      if (!groups[firmName]) groups[firmName] = [];
      groups[firmName].push(r);
    });
    return groups;
  }, [filteredResults]);

  const downloadXLSX = () => {
    if (!filteredResults.length) return;
    const data = filteredResults.map(r => ({
      'Firma': r.Sheet.split(' | ')[0],
      'Linia': r.Linia,
      'Imię': r.Imie,
      'Nazwisko': r.Nazwisko,
      'Suma godzin': r.Razem,
      'Stawka (Pracownik)': HOURLY_RATE,
      'Wynagrodzenie (PLN)': Number((r.Razem * HOURLY_RATE).toFixed(2)),
      'Stawka (Klient)': CLIENT_RATE,
      'Faktura Klient (PLN)': Number((r.Razem * CLIENT_RATE).toFixed(2))
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new(); 
    XLSX.utils.book_append_sheet(wb, ws, "Raport"); 
    XLSX.writeFile(wb, "EWL_Raport_Finansowy_Firmy.xlsx");
  };

  const groupedAnalyses = useMemo(() => {
    const groups: Record<string, { key: string; analysis: SheetAnalysis }[]> = {};
    analyses.forEach((analysis, key) => {
      if (!groups[analysis.fileName]) groups[analysis.fileName] = [];
      groups[analysis.fileName].push({ key, analysis });
    });
    return groups;
  }, [analyses]);

  const renderTable = (data: CalculationResult[], title?: string) => {
    const groupHours = data.reduce((sum, r) => sum + r.Razem, 0);
    return (
      <div className="space-y-4 mb-8 last:mb-0">
        {title && (
          <div className="flex items-center justify-between bg-white/[0.03] border-l-4 border-blue-500 p-4 rounded-r-xl">
            <div>
              <h3 className="text-sm font-black text-white uppercase tracking-wider">{title}</h3>
              <p className="text-[10px] text-white/30 uppercase font-bold">Liczba osób: {data.length}</p>
            </div>
            <div className="flex gap-4">
              <div className="text-right">
                <p className="text-[9px] font-black text-white/30 uppercase">Suma Godzin</p>
                <p className="text-sm font-black text-blue-400">{groupHours.toFixed(1)}h</p>
              </div>
              <div className="text-right">
                <p className="text-[9px] font-black text-white/30 uppercase">Faktura Klient</p>
                <p className="text-sm font-black text-emerald-400">{formatCurrency(groupHours * CLIENT_RATE)}</p>
              </div>
            </div>
          </div>
        )}
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-left">
            <thead className="text-[9px] font-black text-white/30 uppercase border-b border-white/5 bg-black/20">
              <tr> 
                <th className="p-4 w-10"></th> 
                <th className="p-4">Pracownik</th> 
                <th className="p-4 text-right">Godziny</th> 
                <th className="p-4 text-right">Dla Pracownika</th> 
                <th className="p-4 text-right pr-8 text-blue-400">Dla Klienta</th> 
              </tr>
            </thead>
            <tbody className="divide-y divide-white/[0.02]">
              {data.map((row, idx) => (
                <React.Fragment key={idx}>
                  <tr onClick={() => setExpandedRow(expandedRow === `${row.Sheet}-${row.Imie}-${row.Nazwisko}` ? null : `${row.Sheet}-${row.Imie}-${row.Nazwisko}`)} className="hover:bg-white/[0.03] cursor-pointer group transition-colors">
                    <td className="p-4 text-center"> {expandedRow === `${row.Sheet}-${row.Imie}-${row.Nazwisko}` ? <ChevronDown size={14} className="text-blue-500" /> : <ChevronRight size={14} className="text-white/20" />} </td>
                    <td className="p-4">
                      <div className="text-sm font-bold text-white/90">{row.Imie} {row.Nazwisko}</div>
                      <div className="text-[8px] text-white/20 font-mono uppercase">{row.Linia || 'Brak linii'}</div>
                    </td>
                    <td className="p-4 text-sm font-black text-white/60 text-right">{row.Razem.toFixed(1)}</td>
                    <td className="p-4 text-sm font-black text-emerald-400 text-right">{formatCurrency(row.Razem * HOURLY_RATE)}</td>
                    <td className="p-4 text-sm font-black text-blue-400 text-right pr-8">{formatCurrency(row.Razem * CLIENT_RATE)}</td>
                  </tr>
                  {expandedRow === `${row.Sheet}-${row.Imie}-${row.Nazwisko}` && (
                    <tr className="bg-blue-500/[0.02]">
                      <td colSpan={5} className="p-6">
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                          <div className="space-y-4">
                            <div className="bg-black/30 p-4 rounded-xl border border-white/5 space-y-4">
                              <div className="flex items-center justify-between border-b border-white/5 pb-2">
                                <p className="text-[10px] font-black text-white/40 uppercase flex items-center gap-2"> <Wallet size={14}/> Rozliczenie </p>
                                <p className="text-[10px] font-black text-blue-400 uppercase">{row.Sheet}</p>
                              </div>
                              <div className="grid grid-cols-2 gap-4">
                                <div>
                                  <span className="text-[8px] font-black text-white/30 uppercase block mb-1">Pracownik ({HOURLY_RATE})</span>
                                  <p className="text-lg font-black text-emerald-400">{formatCurrency(row.Razem * HOURLY_RATE)}</p>
                                </div>
                                <div>
                                  <span className="text-[8px] font-black text-white/30 uppercase block mb-1">Klient ({CLIENT_RATE})</span>
                                  <p className="text-lg font-black text-blue-400">{formatCurrency(row.Razem * CLIENT_RATE)}</p>
                                </div>
                              </div>
                              <div className="pt-2 border-t border-white/5 flex justify-between items-center">
                                <span className="text-[8px] font-black text-white/30 uppercase">Zysk na godzinach:</span>
                                <span className="text-xs font-black text-indigo-400">{formatCurrency(row.Razem * (CLIENT_RATE - HOURLY_RATE))}</span>
                              </div>
                            </div>
                          </div>
                          <div className="grid grid-cols-8 gap-1 opacity-50">
                            {Array.from({ length: 31 }, (_, i) => i + 1).map(d => (
                              <div key={d} className={`p-1 border rounded text-center ${row.DailyBreakdown[d] ? 'bg-blue-500/10 border-blue-500/20 text-white' : 'border-white/5'}`}>
                                <p className="text-[6px]">{d}</p> <p className="text-[8px] font-black">{row.DailyBreakdown[d] || '-'}</p>
                              </div>
                            ))}
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  return (
    <div className="max-w-[1440px] mx-auto p-4 md:p-8 space-y-6 text-[#e8eefc]">
      <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 border-b border-white/10 pb-6">
        <div className="flex items-center gap-4">
          <div className="bg-blue-600/20 p-3 rounded-2xl border border-blue-500/20 shadow-inner">
            <Calculator className="text-blue-500" size={32} />
          </div>
          <div>
            <h1 className="text-2xl font-black flex items-center gap-2"> EWL Pay Pro <span className="text-[10px] bg-blue-600 px-2 py-0.5 rounded-full uppercase tracking-tighter">Enterprise</span> </h1>
            <p className="text-white/40 text-xs font-medium uppercase tracking-widest">Zaawansowany system rozliczeń i audytu firm</p>
          </div>
        </div>
        <div className="flex gap-3">
           <div className="bg-[#0f1726] p-1 rounded-xl border border-white/10 flex">
             <button onClick={() => setActiveTab('calc')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all flex items-center gap-2 ${activeTab === 'calc' ? 'bg-blue-600 text-white shadow-lg' : 'text-white/40 hover:text-white'}`}> <Calculator size={14} /> OBLICZENIA </button>
             <button onClick={() => setActiveTab('verify')} className={`px-4 py-2 rounded-lg text-xs font-black transition-all flex items-center gap-2 ${activeTab === 'verify' ? 'bg-orange-600 text-white shadow-lg' : 'text-white/40 hover:text-white'}`}> <ArrowRightLeft size={14} /> WERYFIKACJA </button>
           </div>
           {activeTab === 'calc' ? (
             <label className="bg-blue-600 hover:bg-blue-500 px-6 py-3 rounded-xl cursor-pointer flex items-center gap-2 font-black transition-all shadow-lg active:scale-95 text-white">
              <FileUp size={20} /> Wgraj Pliki <input type="file" className="hidden" accept=".xlsx,.xls" multiple onChange={handleFileUpload} />
             </label>
           ) : (
             <label className="bg-orange-600 hover:bg-orange-500 px-6 py-3 rounded-xl cursor-pointer flex items-center gap-2 font-black transition-all shadow-lg active:scale-95 text-white">
              <UploadCloud size={20} /> Wgraj Master <input type="file" className="hidden" accept=".xlsx,.xls" multiple onChange={handleMasterFileUpload} />
             </label>
           )}
        </div>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <div className="lg:col-span-4 space-y-6">
          <div className={`p-4 rounded-xl border flex items-center gap-3 transition-all ${status.type === 'error' ? 'bg-red-500/10 border-red-500/30 text-red-400' : status.type === 'success' ? 'bg-emerald-500/10 border-emerald-500/30 text-emerald-400' : 'bg-blue-500/10 border-blue-500/30 text-blue-400'}`}>
            {status.type === 'error' ? <AlertCircle size={20} /> : <CheckCircle2 size={20} />} <span className="text-sm font-bold">{status.msg}</span>
          </div>

          {activeTab === 'calc' ? (
            <div className="space-y-6">
              <div className="bg-[#0f1726] border border-white/10 rounded-2xl p-6 space-y-5 shadow-2xl">
                <h3 className="text-[10px] font-black uppercase tracking-widest text-white/40 flex items-center gap-2"> <Settings2 size={14} /> Parametry Rozliczenia </h3>
                
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1"> <span className="text-[10px] font-bold text-white/40 ml-1 uppercase">Dzień Od</span> <input type="number" value={dayRange.from} onChange={e => setDayRange(p => ({...p, from: +e.target.value}))} className="w-full bg-black/40 border border-white/10 rounded-xl p-3 text-sm text-white font-mono outline-none focus:border-blue-500/50" /> </div>
                  <div className="space-y-1"> <span className="text-[10px] font-bold text-white/40 ml-1 uppercase">Dzień Do</span> <input type="number" value={dayRange.to} onChange={e => setDayRange(p => ({...p, to: +e.target.value}))} className="w-full bg-black/40 border border-white/10 rounded-xl p-3 text-sm text-white font-mono outline-none focus:border-blue-500/50" /> </div>
                </div>

                <div className="space-y-3 pt-2">
                  <div className="bg-black/20 p-4 rounded-xl border border-white/5 space-y-4">
                    <p className="text-[10px] font-black text-emerald-400 uppercase tracking-widest flex items-center gap-2"> <Coins size={12}/> Finanse i Kontrahenci </p>
                    <div className="grid grid-cols-2 gap-2">
                      <div className="p-3 bg-emerald-500/5 rounded-xl border border-emerald-500/10">
                        <span className="text-[8px] font-bold text-white/40 uppercase block mb-1">Pracownik:</span>
                        <span className="text-xs font-black text-emerald-400">{HOURLY_RATE} zł</span>
                      </div>
                      <div className="p-3 bg-blue-500/5 rounded-xl border border-blue-500/10 text-right">
                        <span className="text-[8px] font-bold text-white/40 uppercase block mb-1">Klient:</span>
                        <span className="text-xs font-black text-blue-400">{CLIENT_RATE} zł</span>
                      </div>
                    </div>
                    
                    <label className="flex items-center gap-3 cursor-pointer group">
                      <input type="checkbox" checked={includeSheetCol} onChange={e => setIncludeSheetCol(e.target.checked)} className="w-4 h-4 accent-blue-500 rounded" />
                      <div className="flex flex-col"> 
                        <span className="text-xs font-bold text-white/90">Separacja projektowa</span> 
                        <span className="text-[9px] text-white/30 italic">Rozdzielaj te same osoby na arkusze</span> 
                      </div>
                    </label>

                    <label className="flex items-center gap-3 cursor-pointer group">
                      <input type="checkbox" checked={filterInactive} onChange={e => setFilterInactive(e.target.checked)} className="w-4 h-4 accent-blue-500 rounded" />
                      <div className="flex flex-col"> 
                        <span className="text-xs font-bold text-white/90">Pokaż brak aktywności</span> 
                        <span className="text-[9px] text-white/30 italic">Uwzględnij kody nieobecności</span> 
                      </div>
                    </label>
                  </div>
                </div>

                <button onClick={handleCalculate} disabled={loading || !selectedSheets.size} className="w-full py-4 bg-blue-600 hover:bg-blue-500 disabled:opacity-50 text-white font-black rounded-xl flex items-center justify-center gap-2 shadow-xl shadow-blue-500/20 transition-all active:scale-[0.98]">
                  {loading ? <RefreshCw className="animate-spin" /> : <Calculator />} GENERUJ RAPORTY FIRM
                </button>
              </div>

              <div className="bg-[#0f1726] border border-white/10 rounded-2xl p-6 max-h-[500px] overflow-hidden flex flex-col shadow-sm">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-[10px] font-black uppercase tracking-widest text-white/40 flex items-center gap-2"> <Layers size={14} /> Arkusze i Segmenty </h3>
                  {analyses.size > 0 && (
                    <button onClick={clearAllFiles} className="text-[10px] font-bold text-red-400 hover:text-red-300 flex items-center gap-1 uppercase transition-colors"> <Eraser size={12}/> Wyzyść </button>
                  )}
                </div>
                <div className="overflow-y-auto space-y-4 pr-2 custom-scrollbar">
                  {(Object.entries(groupedAnalyses) as [string, { key: string; analysis: SheetAnalysis }[]][]).map(([fileName, sheets]) => (
                    <div key={fileName} className="space-y-2">
                      <p className="text-[10px] font-black uppercase text-blue-400 border-l-2 border-blue-500 pl-2 mb-2 truncate" title={fileName}>{fileName}</p>
                      {sheets.map(({ key, analysis: a }) => (
                        <div key={key} className={`p-3 rounded-xl border transition-all ${selectedSheets.has(key) ? 'border-blue-500/40 bg-blue-500/10' : 'border-white/5 bg-white/5'}`}>
                           <div className="flex items-center justify-between mb-2">
                              <label className="flex items-center gap-3 cursor-pointer flex-1 min-w-0">
                                <input type="checkbox" checked={selectedSheets.has(key)} onChange={() => setSelectedSheets(p => { const n = new Set(p); n.has(key) ? n.delete(key) : n.add(key); return n; })} className="w-4 h-4 accent-blue-500 rounded" />
                                <span className="text-[11px] font-black truncate" title={a.sheetName}>{a.sheetName}</span>
                              </label>
                              {selectedSheets.has(key) && a.segments && a.segments.length > 1 && (
                                <div className="flex items-center gap-1 bg-black/40 px-1.5 py-0.5 rounded-lg border border-white/5">
                                  <span className="text-[8px] font-bold text-white/30 uppercase mr-1">Seg:</span>
                                  {a.segments.map(s => (
                                    <button 
                                      key={s.seg}
                                      onClick={() => setSheetSegments(prev => new Map(prev).set(key, s.seg))}
                                      className={`w-4 h-4 rounded text-[9px] font-black transition-all ${sheetSegments.get(key) === s.seg ? 'bg-blue-600 text-white' : 'text-white/20 hover:text-white'}`}
                                    >
                                      {s.seg}
                                    </button>
                                  ))}
                                </div>
                              )}
                           </div>
                           {a.segments && a.segments.length > 0 && selectedSheets.has(key) && (
                             <div className="flex items-center justify-between text-[8px] font-bold text-white/20 uppercase px-1">
                               <span>Dni: {a.segments.find(s => s.seg === (sheetSegments.get(key) || 1))?.dayMin}-{a.segments.find(s => s.seg === (sheetSegments.get(key) || 1))?.dayMax}</span>
                               <span>{a.segments.find(s => s.seg === (sheetSegments.get(key) || 1))?.countPairs} kolumn</span>
                             </div>
                           )}
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          ) : (
            <div className="bg-[#0f1726] border border-orange-500/20 rounded-2xl p-6 space-y-6 shadow-2xl">
              <div className="flex items-center justify-between border-b border-white/5 pb-4">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-orange-500/10 rounded-xl"> <ShieldAlert className="text-orange-500" size={20} /> </div>
                  <h3 className="text-xs font-black text-white/90 uppercase">Weryfikacja Partnerów</h3>
                </div>
                <button onClick={() => {setSystemRawData(''); setMasterRawData(''); setMasterFilesData(new Map()); setComparisonResults(null);}} className="p-2 hover:bg-white/5 rounded-lg text-white/20 hover:text-red-400 transition-colors"> <Eraser size={16}/> </button>
              </div>

              <div className="space-y-4">
                {masterFilesData.size > 0 && (
                  <div className="space-y-2">
                    <p className="text-[9px] font-black text-white/30 uppercase tracking-widest px-1">Wgrane pliki partnerów:</p>
                    <div className="grid gap-2">
                      {Array.from(masterFilesData.keys()).map(name => (
                        <div key={name} className="flex items-center justify-between bg-black/40 border border-white/5 p-2 rounded-lg">
                          <div className="flex items-center gap-2 overflow-hidden">
                            <FileBox size={14} className="text-orange-500 flex-shrink-0" />
                            <span className="text-[10px] font-bold truncate">{name}</span>
                          </div>
                          <span className="text-[8px] font-black text-white/20 uppercase">{masterFilesData.get(name)?.length} osób</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                <div className="space-y-2">
                  <div className="flex justify-between items-center px-1">
                    <p className="text-[10px] font-bold text-blue-400 uppercase flex items-center gap-1.5"> <FileText size={12}/> Raport z Kalkulatora </p>
                    {results && <span className="text-[8px] bg-blue-600/20 text-blue-400 px-1.5 py-0.5 rounded font-black">GOTOWY ({results.length} os.)</span>}
                  </div>
                  <textarea 
                    value={systemRawData} 
                    onChange={(e) => setSystemRawData(e.target.value)}
                    placeholder={results ? "Zostaw puste, aby użyć aktualnego raportu..." : "Wklej dane z kalkulatora..."}
                    className="w-full h-20 bg-black/40 border border-white/10 rounded-xl p-3 text-[10px] font-mono outline-none focus:border-blue-500/50 transition-all placeholder:opacity-20"
                  />
                </div>

                <div className="space-y-2">
                  <p className="text-[10px] font-bold text-orange-400 uppercase flex items-center gap-1.5 px-1"> <ClipboardList size={12}/> Tekst Master (opcja) </p>
                  <textarea 
                    value={masterRawData} 
                    onChange={(e) => setMasterRawData(e.target.value)}
                    placeholder="Wklej dane od kontrahentów (jeśli nie wgrywasz plików)..." 
                    className="w-full h-32 bg-black/40 border border-white/10 rounded-xl p-3 text-[10px] font-mono outline-none focus:border-orange-500/50 transition-all placeholder:opacity-20"
                  />
                </div>
              </div>

              <div className="p-4 bg-orange-500/5 border border-orange-500/10 rounded-xl space-y-2">
                 <div className="flex items-center gap-2 text-orange-400"> <AlertTriangle size={14}/> <span className="text-[10px] font-black uppercase">Algorytm Dopasowania</span> </div>
                 <ul className="text-[8px] text-white/40 space-y-1 ml-4 list-disc uppercase font-bold">
                   <li>Usuwa emotikony i symbole z nazwisk</li>
                   <li>Ignoruje literówki i kolejność Imie/Nazwisko</li>
                   <li>Sumuje godziny z wielu plików master</li>
                 </ul>
              </div>

              <button onClick={handleVerify} disabled={loading} className="w-full py-4 bg-orange-600 hover:bg-orange-500 text-white font-black rounded-xl flex items-center justify-center gap-2 shadow-xl shadow-orange-500/20 transition-all active:scale-[0.98]">
                {loading ? <RefreshCw className="animate-spin" /> : <Diff size={20} />} URUCHOM AUDYT KRZYŻOWY
              </button>
            </div>
          )}
        </div>

        <div className="lg:col-span-8 space-y-6">
          <div className="bg-[#0f1726] border border-white/10 rounded-2xl flex flex-col overflow-hidden shadow-xl min-h-[600px]">
            {activeTab === 'calc' ? (
              <>
                {calcStats && (
                  <div className="grid grid-cols-4 gap-px bg-white/5 border-b border-white/10">
                    <div className="p-6 bg-[#0f1726] text-center"> 
                      <p className="text-[9px] font-black text-white/30 uppercase tracking-widest mb-1">Łączne Godziny</p> 
                      <p className="text-2xl font-black text-blue-400">{calcStats.totalHours.toFixed(1)}</p> 
                    </div>
                    <div className="p-6 bg-[#0f1726] text-center border-x border-white/5"> 
                      <p className="text-[9px] font-black text-white/30 uppercase tracking-widest flex items-center justify-center gap-1 mb-1"> <Wallet size={10} /> Koszt Pracowników</p> 
                      <p className="text-2xl font-black text-emerald-400">{formatCurrency(calcStats.totalHours * HOURLY_RATE)}</p> 
                    </div>
                    <div className="p-6 bg-[#0f1726] text-center border-r border-white/5"> 
                      <p className="text-[9px] font-black text-white/30 uppercase tracking-widest flex items-center justify-center gap-1 mb-1"> <Building2 size={10} /> Faktura Klient</p> 
                      <p className="text-2xl font-black text-blue-400">{formatCurrency(calcStats.totalHours * CLIENT_RATE)}</p> 
                    </div>
                    <div className="p-6 bg-[#0f1726] text-center"> 
                      <p className="text-[9px] font-black text-white/30 uppercase tracking-widest mb-1">Firmy / Osoby</p> 
                      <p className="text-2xl font-black text-white">
                        {Object.keys(groupedResultsByFirm).length} / {calcStats.totalPeople}
                      </p> 
                    </div>
                  </div>
                )}
                <div className="p-4 border-b border-white/10 bg-white/[0.02] flex items-center justify-between gap-4">
                  <div className="relative flex-1">
                    <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-white/20" />
                    <input type="text" placeholder="Szukaj pracownika, linii lub firmy..." value={searchQuery} onChange={e => setSearchQuery(e.target.value)} className="w-full bg-black/40 border border-white/10 rounded-xl pl-10 pr-4 py-2.5 text-xs outline-none focus:border-blue-500/50" />
                  </div>
                  <div className="flex gap-2">
                    <div className="flex bg-black/40 p-1 rounded-xl border border-white/5">
                      <button onClick={() => setViewMode('grouped')} className={`p-2 rounded-lg transition-all ${viewMode === 'grouped' ? 'bg-blue-600 text-white shadow-lg' : 'text-white/20 hover:text-white'}`}> <LayoutGrid size={16}/> </button>
                      <button onClick={() => setViewMode('flat')} className={`p-2 rounded-lg transition-all ${viewMode === 'flat' ? 'bg-blue-600 text-white shadow-lg' : 'text-white/20 hover:text-white'}`}> <List size={16}/> </button>
                    </div>
                    <button onClick={downloadXLSX} disabled={!results} className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-black flex items-center gap-2 transition-all active:scale-95 shadow-lg"> <Download size={14} /> EKSPORTUJ XLSX </button>
                  </div>
                </div>
                <div className="flex-1 overflow-auto custom-scrollbar max-h-[800px] p-4">
                  {results ? (
                    viewMode === 'grouped' ? (
                      (Object.entries(groupedResultsByFirm) as [string, CalculationResult[]][]).map(([firm, data]) => (
                        <div key={firm}>
                          {renderTable(data, firm)}
                        </div>
                      ))
                    ) : (
                      renderTable(filteredResults)
                    )
                  ) : (
                    <div className="flex flex-col items-center justify-center h-full py-40 opacity-20 text-center"> 
                      <FileSpreadsheet size={64}/> 
                      <p className="mt-4 font-black uppercase text-xs tracking-widest">Wgraj pliki i kliknij "Generuj Raport"</p> 
                    </div>
                  )}
                </div>
              </>
            ) : (
              <div className="flex flex-col h-full">
                <div className="p-6 border-b border-white/5 flex items-center justify-between bg-orange-500/[0.02]">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-orange-600/20 rounded-xl border border-orange-500/20 shadow-inner"> <ArrowRightLeft className="text-orange-500" size={20} /> </div>
                    <div> <h2 className="text-sm font-black text-white uppercase tracking-widest">Audytor Partnerów</h2> <p className="text-[10px] text-white/40 font-bold uppercase">Automatyczne dopasowanie do raportu EWL</p> </div>
                  </div>
                </div>
                <div className="flex-1 overflow-auto custom-scrollbar">
                  {!comparisonResults ? (
                    <div className="flex flex-col items-center justify-center h-full py-40 text-center px-12">
                      <Diff size={64} className="mb-4 opacity-10" />
                      <p className="text-sm font-black uppercase tracking-widest opacity-30">Weryfikacja krzyżowa</p>
                      <p className="text-[10px] mt-4 max-w-sm text-white/30 leading-relaxed uppercase">Wgraj pliki master od partnerów i kliknij Audyt. System dopasuje osoby z raportu obliczeń do ich odpowiedników w plikach master, ignorując emotikony i różnice w pisowni.</p>
                    </div>
                  ) : (
                    <table className="w-full border-collapse text-left">
                      <thead className="sticky top-0 z-20 bg-[#0f1726] text-[9px] font-black text-white/30 uppercase border-b border-white/5 bg-black/20">
                        <tr> <th className="p-4">Pracownik</th> <th className="p-4 text-right">EWL Pay Pro</th> <th className="p-4 text-right">Master (Partner)</th> <th className="p-4 text-right pr-8">Rozbieżność</th> </tr>
                      </thead>
                      <tbody className="divide-y divide-white/[0.02]">
                        {comparisonResults.map((row, idx) => (
                          <tr key={idx} className={`hover:bg-white/[0.03] transition-colors ${row.status === 'ok' ? 'opacity-40' : 'bg-orange-500/[0.05]'}`}>
                            <td className="p-4 flex items-center gap-3">
                              {row.status === 'ok' ? <CheckCircle2 size={16} className="text-emerald-500" /> : <AlertTriangle size={16} className="text-orange-500" />}
                              <div className="flex flex-col min-w-0"> 
                                <span className="text-xs font-black text-white/90 truncate">{row.name}</span> 
                                <div className="flex items-center gap-2">
                                  <span className="text-[8px] font-black text-white/20 uppercase tracking-tighter">
                                    {row.status === 'ok' ? 'Zgodny' : row.status === 'missing_master' ? 'Brak w Master' : row.status === 'missing_system' ? 'Brak w Systemie' : 'Niezgodność'}
                                  </span>
                                  {row.similarity !== undefined && row.similarity > 0 && row.similarity < 1 && (
                                    <span className={`text-[8px] font-black px-1 rounded ${row.similarity >= 0.8 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-orange-500/10 text-orange-500'}`}>
                                      {Math.round(row.similarity * 100)}% Match
                                    </span>
                                  )}
                                </div>
                              </div>
                            </td>
                            <td className="p-4 text-xs font-mono text-right text-blue-400">{row.systemHours.toFixed(1)}h</td>
                            <td className="p-4 text-xs font-mono text-right text-orange-400">{row.masterHours.toFixed(1)}h</td>
                            <td className={`p-4 text-sm font-black text-right pr-8 ${row.diff === 0 ? 'text-emerald-500' : 'text-red-500'}`}> {row.diff > 0 ? `+${row.diff.toFixed(1)}` : row.diff.toFixed(1)} </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default App;
