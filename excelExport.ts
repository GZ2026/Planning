
import * as XLSX from 'xlsx';
import { PlanningResults } from '../types';

export const exportToExcel = (results: PlanningResults) => {
  const wb = XLSX.utils.book_new();

  // Tab 1: Dispatch_Plan
  // Headers: City, Store Name, SKU Type, SKU, Reorder Units
  const planData = results.dispatchPlan.map(item => ({
    'City': item.City,
    'Store Name': item.StoreID,
    'SKU Type': item.SKUType,
    'SKU': item.SKU,
    'Reorder Units': item.FinalDispatch
  }));
  const wsPlan = XLSX.utils.json_to_sheet(planData);
  XLSX.utils.book_append_sheet(wb, wsPlan, "Dispatch_Plan");

  // Tab 2: Store_Summary
  // Headers: Store Name, Opening SIH, Core Reorder, Top-Up Units, 15-day Forecast, Final Freezer Fill %
  const summaryData = results.storeSummaries.map(s => {
    const totalUnits = s.InitialStock + s.CoreReorder + s.TopUpUnits;
    
    return {
      'Store Name': s.StoreID,
      'Opening SIH': s.InitialStock,
      'Core Reorder': s.CoreReorder,
      'Top-Up Units': s.TopUpUnits,
      'Total Units (SIH+Reorder+TopUp)': totalUnits,
      'Freezer Capacity (All Mix)': s.Capacity,
      '15-day Forecast': Math.round(s.Forecast15Day),
      'Final Freezer Fill %': `${s.FillPercentage.toFixed(2)}%`
    };
  });
  const wsSummary = XLSX.utils.json_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, "Store_Summary");

  // Generate filename
  const dateStr = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `GZ_Tier1_Dispatch_Plan_${dateStr}.xlsx`);
};
