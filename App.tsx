
import React, { useState } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Cell 
} from 'recharts';
import { 
  Upload, FileCheck, Download, Calendar, 
  ChevronRight, Database, BarChart3, Package, Info, AlertCircle, RefreshCw, Table, ShieldCheck
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { SIHRecord, SalesRecord, FreezerRecord, PlanningDay, PlanningResults } from './types';
import { runPlanningLogic, deduceSKUType } from './services/plannerEngine';
import { exportToExcel } from './services/excelExport';

const App: React.FC = () => {
  const [sihData, setSihData] = useState<SIHRecord[] | null>(null);
  const [salesData, setSalesData] = useState<SalesRecord[] | null>(null);
  const [freezerData, setFreezerData] = useState<FreezerRecord[] | null>(null);
  const [planningDay, setPlanningDay] = useState<PlanningDay>(PlanningDay.MON_FRI);
  const [results, setResults] = useState<PlanningResults | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const normalizeKeys = (data: any[], type: 'SIH' | 'Sales' | 'Freezer'): any[] => {
    if (data.length === 0) return [];
    
    // Get headers of first object
    const headers = Object.keys(data[0]);

    // Detect column mapping for this dataset
    const mapping: Record<string, string> = {};

    headers.forEach(h => {
      const clean = h.toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');

      // Store ID
      if (
        clean === 'store' || clean === 'storeid' || clean === 'storename' || 
        clean === 'outlet' || clean === 'outletname' || clean === 'customer' || clean === 'customername' ||
        clean === 'site' || clean === 'sitename' || clean === 'hub' || clean === 'hubname' ||
        clean === 'dc' || clean === 'dcname' || clean === 'distributioncenter' ||
        clean.includes('store') || clean.includes('outlet') || clean.includes('customer') || clean.includes('site') || clean.includes('hub')
      ) {
        mapping[h] = 'StoreID';
      }
      // SKU Type
      else if (
        clean === 'skutype' || clean === 'itemtype' || clean === 'type' || clean === 'category' ||
        clean === 'itemcategory' || clean === 'productcategory' || clean === 'group' || clean === 'subcategory' ||
        clean.includes('type') || clean.includes('category') || clean.includes('group') || clean.includes('class')
      ) {
        mapping[h] = 'SKUType';
      }
      // SKU
      else if (
        clean === 'sku' || clean === 'skuname' || clean === 'skucode' ||
        clean === 'item' || clean === 'itemname' || clean === 'itemcode' ||
        clean === 'product' || clean === 'productname' || clean === 'productcode' ||
        clean === 'article' || clean === 'articlename' || clean === 'articlecode' ||
        clean === 'material' || clean === 'materialname' || clean === 'materialcode' ||
        clean.includes('sku') || clean.includes('item') || clean.includes('product') || clean.includes('article') || clean.includes('material')
      ) {
        mapping[h] = 'SKU';
      }
      // City
      else if (
        clean === 'city' || clean === 'location' || clean === 'region' || clean === 'state' || clean === 'zone' ||
        clean.includes('city') || clean.includes('location') || clean.includes('region') || clean.includes('state') || clean.includes('zone')
      ) {
        mapping[h] = 'City';
      }
      // StockLevel / Units / Capacity based on type
      else if (type === 'SIH' && (
        clean === 'sih' || clean === 'sihqty' || clean === 'stockinhand' || clean === 'stock' || clean === 'stocklevel' ||
        clean === 'closingstock' || clean === 'openingstock' || clean === 'closingsih' || clean === 'qtyonhand' || clean === 'onhand' ||
        clean.includes('sih') || clean.includes('stock') || clean.includes('onhand') || clean.includes('closing') || clean.includes('opening') ||
        clean.includes('physic') || clean === 'qty' || clean === 'quantity' || clean === 'units' || clean.includes('hand')
      )) {
        mapping[h] = 'StockLevel';
      }
      else if (type === 'Sales' && (
        clean === 'totalquantity' || clean === 'totalqty' || clean === 'units' || clean === 'qty' || clean === 'quantity' ||
        clean === 'unitssold' || clean === 'salesqty' || clean === 'salesquantity' || clean === 'sumofqty' || clean === 'qtysold' || clean === 'volume' ||
        clean.includes('sold') || clean.includes('sale') || clean.includes('qty') || clean.includes('quant') || clean.includes('unit') || clean.includes('volume')
      )) {
        mapping[h] = 'Units';
      }
      else if (type === 'Sales' && (
        clean === 'date' || clean === 'salesdate' || clean === 'invoicedate' || clean === 'transactiondate' || clean === 'postingdate' ||
        clean.includes('date') || clean.includes('time')
      )) {
        mapping[h] = 'Date';
      }
      else if (type === 'Freezer' && (
        clean === 'maxunitsallmix' || clean === 'freezercapacity' || clean === 'capacity' || clean === 'maxcapacity' || clean === 'maxunits' ||
        clean === 'freezersize' || clean === 'freezervolume' || clean === 'chestfreezercapacity' || clean === 'holdingcapacity' || clean === 'maxstockcapacity' ||
        clean.includes('allmix') || clean.includes('all mix') || clean.includes('all-mix') ||
        ((clean.includes('capac') || clean.includes('max') || clean.includes('size') || clean.includes('hold') || clean.includes('volume')) &&
         !clean.includes('cup') && !clean.includes('cone') && !clean.includes('bar') && !clean.includes('cass') && !clean.includes('tub') && !clean.includes('stick'))
      )) {
        mapping[h] = 'Capacity';
      }
    });

    // Fallbacks if mapping failed for required fields
    if (type === 'SIH' && !Object.values(mapping).includes('StockLevel')) {
      const numericCol = headers.find(h => {
        const clean = h.toLowerCase().trim();
        return clean.includes('qty') || clean.includes('quantity') || clean.includes('stock') || clean.includes('sih') || clean.includes('units') || clean.includes('balance');
      });
      if (numericCol) mapping[numericCol] = 'StockLevel';
    }

    if (type === 'Sales' && !Object.values(mapping).includes('Units')) {
      const numericCol = headers.find(h => {
        const clean = h.toLowerCase().trim();
        return clean.includes('units') || clean.includes('qty') || clean.includes('quantity') || clean.includes('sold') || clean.includes('sales');
      });
      if (numericCol) mapping[numericCol] = 'Units';
    }

    if (type === 'Freezer' && !Object.values(mapping).includes('Capacity')) {
      const numericCol = headers.find(h => {
        const clean = h.toLowerCase().trim();
        const hasCategory = clean.includes('cup') || clean.includes('cone') || clean.includes('bar') || clean.includes('cass') || clean.includes('tub') || clean.includes('stick');
        return !hasCategory && (clean.includes('capacity') || clean.includes('max') || clean.includes('units') || clean.includes('size'));
      });
      if (numericCol) mapping[numericCol] = 'Capacity';
    }

    return data.map(row => {
      const normalizedRow: any = {
        StoreID: '',
        SKU: '',
        SKUType: 'Unknown',
        City: 'Unknown',
        StockLevel: 0,
        Units: 0,
        Capacity: 0,
        Date: ''
      };

      if (type === 'Freezer') {
        normalizedRow.typeCapacities = {};
      }

      headers.forEach(h => {
        const mappedKey = mapping[h];
        let value = row[h];

        if (mappedKey) {
          if (mappedKey === 'StoreID' || mappedKey === 'SKU' || mappedKey === 'City' || mappedKey === 'SKUType' || mappedKey === 'Date') {
            value = value !== undefined && value !== null ? String(value).trim() : '';
            normalizedRow[mappedKey] = value;
          } else if (['StockLevel', 'Units', 'Capacity'].includes(mappedKey)) {
            const parsed = parseFloat(String(value).replace(/[^0-9.-]+/g, ""));
            value = isNaN(parsed) ? 0 : parsed;
            normalizedRow[mappedKey] = value;
          }
        }

        if (type === 'Freezer') {
          const cleanHeader = h.toString().toLowerCase().trim();
          const parsedVal = parseFloat(String(value).replace(/[^0-9.-]+/g, ""));
          const numericValue = isNaN(parsedVal) ? 0 : parsedVal;

          if (cleanHeader.includes('cup') || cleanHeader.includes('cups')) {
            normalizedRow.typeCapacities.cup = numericValue;
          } else if (cleanHeader.includes('cone') || cleanHeader.includes('cones')) {
            normalizedRow.typeCapacities.cone = numericValue;
          } else if (cleanHeader.includes('bar') || cleanHeader.includes('bars')) {
            normalizedRow.typeCapacities.bar = numericValue;
          } else if (cleanHeader.includes('cass') || cleanHeader.includes('cassata')) {
            normalizedRow.typeCapacities.cassata = numericValue;
          } else if (cleanHeader.includes('tub') || cleanHeader.includes('tubs')) {
            normalizedRow.typeCapacities.tub = numericValue;
          } else if (cleanHeader.includes('stick') || cleanHeader.includes('sticks')) {
            normalizedRow.typeCapacities.stick = numericValue;
          } else if (cleanHeader.includes('all mix') || cleanHeader.includes('allmix') || cleanHeader.includes('all-mix')) {
            normalizedRow.typeCapacities.allmix = numericValue;
          }
        }
      });

      if (type !== 'Freezer' && normalizedRow.SKU) {
        normalizedRow.SKUType = deduceSKUType(normalizedRow.SKU, normalizedRow.SKUType);
      }

      return normalizedRow;
    });
  };

  const handleFileUpload = (type: 'SIH' | 'Sales' | 'Freezer', file: File) => {
    setError(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawJson = XLSX.utils.sheet_to_json(ws);
        
        if (rawJson.length === 0) {
          setError(`The ${type} file appears to be empty.`);
          return;
        }

        const normalized = normalizeKeys(rawJson, type);
        const firstRow = normalized[0];
        if (type === 'SIH' && !firstRow.StoreID) setError("SIH file missing 'Store Name' column.");
        if (type === 'Sales' && !firstRow.StoreID) setError("Sales file missing 'Store Name' column.");
        if (type === 'Freezer' && !firstRow.StoreID) setError("Freezer file missing 'Store Name' column.");

        if (type === 'SIH') setSihData(normalized as SIHRecord[]);
        if (type === 'Sales') setSalesData(normalized as SalesRecord[]);
        if (type === 'Freezer') setFreezerData(normalized as FreezerRecord[]);
      } catch (err) {
        setError(`Failed to parse ${type} file.`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleGeneratePlan = () => {
    if (!sihData || !salesData || !freezerData) {
      setError("Missing required files.");
      return;
    }

    setIsLoading(true);
    setError(null);
    
    setTimeout(() => {
      try {
        const calculatedResults = runPlanningLogic(sihData, salesData, freezerData, planningDay);
        if (calculatedResults.storeSummaries.length === 0) {
          setError("No matching Store Names found.");
          setResults(null);
        } else {
          setResults(calculatedResults);
        }
      } catch (err) {
        setError("Planning logic error.");
        console.error(err);
      } finally {
        setIsLoading(false);
      }
    }, 600);
  };

  const loadSampleData = () => {
    const sampleStores = ["Store Alpha", "Store Beta", "Store Gamma"];
    const sampleCities = ["Mumbai", "Delhi", "Bangalore"];
    const sampleSKUs = ["Vanilla Tub 500ml", "Choco Cone", "Strawberry Stick"];
    const sampleSKUTypes = ["Tub", "Cone", "Stick"];
    
    const sampleSIH = [];
    const sampleSales = [];
    const sampleFreezer = [];

    sampleStores.forEach((s, idx) => {
      const city = sampleCities[idx];
      sampleFreezer.push({ StoreID: s, City: city, Capacity: 1000 });
      sampleSKUs.forEach((sku, sIdx) => {
        const skuType = sampleSKUTypes[sIdx];
        sampleSIH.push({ StoreID: s, City: city, SKU: sku, SKUType: skuType, StockLevel: 50 });
        for (let i = 0; i < 15; i++) {
          sampleSales.push({ StoreID: s, City: city, SKU: sku, SKUType: skuType, Units: Math.floor(Math.random() * 10), Date: `2024-05-${i+1}` });
        }
      });
    });

    setSihData(sampleSIH as any);
    setSalesData(sampleSales as any);
    setFreezerData(sampleFreezer as any);
    setError(null);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-8">
      <header className="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-start md:items-center mb-8 gap-4 border-b border-slate-800 pb-6">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-cyan-500/10 rounded-xl flex items-center justify-center border border-cyan-500/20">
            <Package className="w-8 h-8 text-cyan-400" />
          </div>
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white">GZ Tier 1 Dispatch Planner</h1>
            <p className="text-slate-400 text-sm">Automated Supply Chain Optimization</p>
          </div>
        </div>
        
        <div className="flex flex-wrap items-center gap-3">
          <button onClick={loadSampleData} className="px-4 py-2 rounded-lg bg-slate-900 border border-slate-700 text-slate-400 hover:text-white transition-all text-sm flex items-center gap-2">
            <RefreshCw className="w-4 h-4" /> Load Sample
          </button>
          <div className="flex items-center gap-1 bg-slate-900 p-1 rounded-xl border border-slate-800">
            <button onClick={() => setPlanningDay(PlanningDay.MON_FRI)} className={`px-4 py-2 rounded-lg transition-all text-sm font-medium ${planningDay === PlanningDay.MON_FRI ? 'bg-cyan-600 text-white shadow-lg' : 'text-slate-400'}`}>Mon-Fri</button>
            <button onClick={() => setPlanningDay(PlanningDay.SATURDAY)} className={`px-4 py-2 rounded-lg transition-all text-sm font-medium ${planningDay === PlanningDay.SATURDAY ? 'bg-cyan-600 text-white shadow-lg' : 'text-slate-400'}`}>Saturday</button>
          </div>
        </div>
      </header>

      {error && (
        <div className="max-w-7xl mx-auto mb-6 p-4 bg-red-500/10 border border-red-500/20 rounded-xl flex items-center gap-3 text-red-400">
          <AlertCircle className="w-5 h-5 flex-shrink-0" />
          <p className="text-sm font-medium">{error}</p>
        </div>
      )}

      <main className="max-w-7xl mx-auto grid grid-cols-1 lg:grid-cols-12 gap-8">
        <div className="lg:col-span-4 space-y-6">
          <section className="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-xl overflow-hidden relative">
            <h2 className="text-lg font-semibold mb-6 flex items-center gap-2"><Database className="w-5 h-5 text-cyan-400" /> Upload Data Center</h2>
            <div className="space-y-4">
              <FileUploadCard label="SIH (Stock In Hand)" onUpload={(f) => handleFileUpload('SIH', f)} hasData={!!sihData} recordCount={sihData?.length} expectedHeaders="Store Name, SKU Name, SIH QTY, SKU Type, City" />
              <FileUploadCard label="Sales (15-day History)" onUpload={(f) => handleFileUpload('Sales', f)} hasData={!!salesData} recordCount={salesData?.length} expectedHeaders="Store Name, SKU Name, Total Quantity, City" />
              <FileUploadCard label="Freezer Capacity" onUpload={(f) => handleFileUpload('Freezer', f)} hasData={!!freezerData} recordCount={freezerData?.length} expectedHeaders="Store Name, City, Max Units - All Mix" />
            </div>
            <button onClick={handleGeneratePlan} disabled={isLoading || !sihData || !salesData || !freezerData} className="w-full mt-8 py-4 px-6 bg-cyan-600 hover:bg-cyan-500 disabled:bg-slate-800 rounded-xl font-bold text-lg transition-all flex items-center justify-center gap-3">
              {isLoading ? <div className="w-6 h-6 border-2 border-white/30 border-t-white rounded-full animate-spin" /> : <>Generate Dispatch Plan <ChevronRight className="w-5 h-5" /></>}
            </button>
          </section>

          {(sihData || salesData || freezerData) && (
            <section className="bg-slate-900 rounded-2xl p-6 border border-slate-800 shadow-xl space-y-4 animate-in fade-in duration-300">
              <h3 className="text-sm font-bold text-slate-300 tracking-wider uppercase flex items-center gap-2">
                <ShieldCheck className="w-4 h-4 text-emerald-400" /> Data Integrity Check
              </h3>
              
              <div className="space-y-3 divide-y divide-slate-800">
                {sihData && (
                  <div className="pt-2 text-xs text-slate-400">
                    <p className="font-bold text-white flex justify-between mb-1">
                      <span>Inventory (SIH):</span>
                      <span className="text-emerald-400">{sihData.length.toLocaleString()} rows</span>
                    </p>
                    <div className="grid grid-cols-2 gap-x-2 gap-y-1 mt-2 text-[11px]">
                      <div>Unique Stores: <strong className="text-slate-100">{new Set(sihData.map(s => s.StoreID.toLowerCase().trim())).size}</strong></div>
                      <div>Total Stock: <strong className="text-emerald-400 font-bold">{sihData.reduce((acc, s) => acc + (Number(s.StockLevel) || 0), 0).toLocaleString()}</strong></div>
                      <div>Avg/Row: <strong className="text-slate-100">{sihData.length ? Math.round(sihData.reduce((acc, s) => acc + (Number(s.StockLevel) || 0), 0) / sihData.length) : 0}</strong></div>
                      <div>Valid (&gt;0): <strong className="text-slate-100">{sihData.filter(s => (Number(s.StockLevel) || 0) > 0).length}</strong></div>
                    </div>
                  </div>
                )}

                {salesData && (
                  <div className="pt-3 text-xs text-slate-400">
                    <p className="font-bold text-white flex justify-between mb-1">
                      <span>Sales File:</span>
                      <span className="text-emerald-400">{salesData.length.toLocaleString()} rows</span>
                    </p>
                    <div className="grid grid-cols-2 gap-x-2 gap-y-1 mt-2 text-[11px]">
                      <div>Total Sold: <strong className="text-emerald-400 font-bold">{salesData.reduce((acc, s) => acc + (Number(s.Units) || 0), 0).toLocaleString()}</strong></div>
                      <div>Avg Sold/Row: <strong className="text-slate-100">{salesData.length ? (salesData.reduce((acc, s) => acc + (Number(s.Units) || 0), 0) / salesData.length).toFixed(1) : '0'}</strong></div>
                    </div>
                  </div>
                )}

                {freezerData && (
                  <div className="pt-3 text-xs text-slate-400 space-y-2">
                    <p className="font-bold text-white flex justify-between">
                      <span>Freezer Capacity File:</span>
                      <span className="text-emerald-400 font-bold">{freezerData.reduce((acc, s) => acc + (Number(s.Capacity) || 0), 0).toLocaleString()} Max Units</span>
                    </p>
                    <div className="text-[11px] space-y-1">
                      <div>Total Stores: <strong className="text-slate-100">{freezerData.length}</strong></div>
                      {freezerData[0]?.typeCapacities && Object.keys(freezerData[0].typeCapacities).length > 0 && (
                        <div className="text-slate-400 mt-1.5 pt-1.5 border-t border-slate-850">
                          <span className="text-slate-300 font-semibold block mb-1 text-[10px] uppercase tracking-wider">Category Max Capacities (First Store):</span>
                          <div className="flex flex-wrap gap-x-2 gap-y-1">
                            {Object.entries(freezerData[0].typeCapacities).map(([category, value]) => (
                              <span key={category} className="bg-slate-800 text-slate-200 px-1.5 py-0.5 rounded text-[10px] font-mono">
                                {category.charAt(0).toUpperCase() + category.slice(1)}: {value}
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </section>
          )}
        </div>

        <div className="lg:col-span-8">
          {!results ? (
            <div className="h-full min-h-[600px] flex flex-col items-center justify-center bg-slate-900/20 rounded-3xl border border-slate-800 border-dashed text-slate-500 p-8 text-center">
              <BarChart3 className="w-12 h-12 text-slate-600 mb-8" />
              <h3 className="text-2xl font-bold text-slate-200 mb-2">Analysis Engine Standby</h3>
              <p className="max-w-md text-slate-500">Upload your operational data to begin the dispatch optimization.</p>
            </div>
          ) : (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-6 duration-700">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <MetricCard label="Average Fill Rate" value={`${results.overallFillRate.toFixed(1)}%`} status={results.overallFillRate > 90 ? "success" : "warning"} />
                <MetricCard label="Dispatch Volume" value={results.dispatchPlan.reduce((acc, p) => acc + p.FinalDispatch, 0).toLocaleString()} sub="Total units recommended" />
                <MetricCard label="Operational Scope" value={results.storeSummaries.length.toString()} sub="Active stores mapped" />
              </div>
              <div className="bg-slate-900 rounded-3xl p-8 border border-slate-800 shadow-2xl">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-8 gap-4">
                  <h3 className="text-xl font-bold text-white flex items-center gap-3"><BarChart3 className="w-6 h-6 text-cyan-400" /> Utilization Dashboard</h3>
                  <button onClick={() => exportToExcel(results)} className="flex items-center gap-3 px-6 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl font-bold transition-all shadow-xl shadow-emerald-900/30">
                    <Download className="w-5 h-5" /> Export Excel Plan
                  </button>
                </div>
                <div className="h-[350px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={results.storeSummaries.slice(0, 20)}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                      <XAxis dataKey="StoreID" stroke="#475569" fontSize={11} dy={10} />
                      <YAxis stroke="#475569" fontSize={11} domain={[0, 100]} tickFormatter={(val) => `${val}%`} />
                      <Tooltip contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #334155', borderRadius: '16px' }} />
                      <Bar dataKey="FillPercentage" radius={[6, 6, 0, 0]} name="Fill Rate">
                        {results.storeSummaries.slice(0, 20).map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={entry.FillPercentage >= 95 ? '#10b981' : '#0ea5e9'} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

const FileUploadCard: React.FC<{ label: string; onUpload: (f: File) => void; hasData: boolean; recordCount?: number; expectedHeaders: string; }> = ({ label, onUpload, hasData, recordCount, expectedHeaders }) => (
  <div className={`p-4 rounded-2xl border-2 transition-all ${hasData ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-slate-800 bg-slate-900/40'}`}>
    <div className="flex justify-between items-center">
      <div className="flex items-center gap-3">
        <div className={`w-12 h-12 rounded-xl flex items-center justify-center ${hasData ? 'bg-emerald-500/10' : 'bg-slate-800'}`}>
          {hasData ? <FileCheck className="w-6 h-6 text-emerald-400" /> : <Upload className="w-6 h-6 text-slate-400" />}
        </div>
        <div>
          <p className={`text-sm font-bold ${hasData ? 'text-emerald-400' : 'text-slate-200'}`}>{label}</p>
          <p className="text-[10px] text-slate-500 uppercase tracking-tight">{hasData ? `${recordCount?.toLocaleString()} rows loaded` : `Need: ${expectedHeaders}`}</p>
        </div>
      </div>
      <label className="cursor-pointer">
        <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files?.[0]; if (f) onUpload(f); }} />
        <span className="px-4 py-2 rounded-xl text-xs font-black uppercase bg-slate-800 text-slate-300 hover:bg-slate-700 transition-all">{hasData ? 'Change' : 'Upload'}</span>
      </label>
    </div>
  </div>
);

const MetricCard: React.FC<{ label: string; value: string; sub?: string; status?: 'success' | 'warning' | 'neutral'; }> = ({ label, value, sub, status = 'neutral' }) => (
  <div className="bg-slate-900 rounded-3xl p-6 border border-slate-800">
    <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] mb-3">{label}</p>
    <h4 className="text-4xl font-black text-white tracking-tighter">{value}</h4>
    {sub && <p className="text-[10px] mt-2 text-slate-500">{sub}</p>}
  </div>
);

const HealthItem: React.FC<{ label: string; count: number }> = ({ label, count }) => (
  <div className="flex justify-between items-center text-xs">
    <span className="text-slate-500">{label}</span>
    <span className="text-emerald-400 font-bold">{count}</span>
  </div>
);

export default App;
