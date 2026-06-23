import { 
  SIHRecord, SalesRecord, FreezerRecord, 
  PlanningDay, PlanningResults, DispatchPlanItem, StoreSummary 
} from '../types';

// Standarized cleanup function for strings to tolerate spacing, casing, and trailing tabs from Excel uploads
const cleanKey = (val: string): string => {
  return (val || '').toString().trim().replace(/\s+/g, ' ').toLowerCase();
};

// Deduce SKU category name from actual SKU name
export const deduceSKUType = (skuName: string, existingType?: string): string => {
  if (existingType && existingType !== 'Unknown' && existingType.trim() !== '') {
    const cleanExist = existingType.trim().toLowerCase();
    if (cleanExist.includes('cup')) return 'Cup';
    if (cleanExist.includes('cone')) return 'Cone';
    if (cleanExist.includes('bar')) return 'Bar';
    if (cleanExist.includes('cassata') || cleanExist.includes('cass')) return 'Cassata';
    if (cleanExist.includes('tub')) return 'Tub';
    if (cleanExist.includes('stick')) return 'Stick';
    return existingType.trim();
  }
  const clean = (skuName || '').toLowerCase();
  if (clean.includes('cup')) return 'Cup';
  if (clean.includes('cone')) return 'Cone';
  if (clean.includes('bar')) return 'Bar';
  if (clean.includes('cassata') || clean.includes('cass')) return 'Cassata';
  if (clean.includes('tub')) return 'Tub';
  if (clean.includes('stick')) return 'Stick';
  return 'Unknown';
};

// Map any SKU type to the clean category key
export const getSKUTypeCapacityKey = (skuType: string): string => {
  const clean = (skuType || '').toLowerCase().trim();
  if (clean.includes('cup')) return 'cup';
  if (clean.includes('cone')) return 'cone';
  if (clean.includes('bar')) return 'bar';
  if (clean.includes('cass')) return 'cassata';
  if (clean.includes('tub')) return 'tub';
  if (clean.includes('stick')) return 'stick';
  return 'allmix';
};

// Retrieve capacity of SKU category in a dynamic way
export const getStoreSKUTypeCapacity = (freezer: FreezerRecord, skuType: string): number => {
  const key = getSKUTypeCapacityKey(skuType);
  const typeCapacities = freezer.typeCapacities || {};
  let cap = typeCapacities[key] || 0;
  if (!cap || cap <= 0) {
    // Fallback to "allmix", or default total Capacity from sheet
    cap = typeCapacities['allmix'] || freezer.Capacity || 0;
  }
  return cap || 550; // default fallback of 550 if none found
};

export const runPlanningLogic = (
  sihData: SIHRecord[],
  salesData: SalesRecord[],
  freezerData: FreezerRecord[],
  planningDay: PlanningDay
): PlanningResults => {
  const leadTime = planningDay === PlanningDay.MON_FRI ? 11 : 12;

  // Pre-process Sales Data with case-invariant and spacing-invariant keys
  const salesMap = new Map<string, number[]>(); 
  salesData.forEach(row => {
    const storeID = cleanKey(row.StoreID);
    const sku = cleanKey(row.SKU);
    if (!storeID || !sku) return;
    const key = `${storeID}|${sku}`;
    if (!salesMap.has(key)) salesMap.set(key, []);
    salesMap.get(key)!.push(Number(row.Units) || 0);
  });

  const avgSalesMap = new Map<string, number>();
  salesMap.forEach((units, key) => {
    const sum = units.reduce((acc, val) => acc + (Number(val) || 0), 0);
    // Divide by 15 days of history
    avgSalesMap.set(key, sum / 15);
  });

  // Index Freezer and SIH data to clean keys
  const freezerMap = new Map<string, FreezerRecord>(); // cleanKey(StoreID) -> FreezerRecord
  freezerData.forEach(f => {
    const key = cleanKey(f.StoreID);
    if (key) freezerMap.set(key, f);
  });

  const sihStoreKeys = new Set(sihData.map(s => cleanKey(s.StoreID)).filter(Boolean));
  // Intersect case-insensitively and spacing-insensitively
  const validStoreKeys = Array.from(freezerMap.keys()).filter(key => sihStoreKeys.has(key));

  const dispatchPlan: DispatchPlanItem[] = [];

  validStoreKeys.forEach(key => {
    const freezerInfo = freezerMap.get(key)!;
    const storeID = freezerInfo.StoreID; // Keep original display name
    const storeCapacity = Number(freezerInfo.Capacity) || 550;
    const storeCity = freezerInfo.City || "Unknown";
    
    // Filter SIH records that match this store
    const storeSIH = sihData.filter(s => cleanKey(s.StoreID) === key);
    
    const storeDispatchCandidates: DispatchPlanItem[] = [];
    
    // Calculate total starting stock (sum of all SKUs' SIH)
    const totalSIH = storeSIH.reduce((acc, item) => acc + (Number(item.StockLevel) || 0), 0);
    const availableSpace = Math.max(0, storeCapacity - totalSIH);
    
    storeSIH.forEach(item => {
      const cleanSKU = cleanKey(item.SKU);
      const avgSales = Number(avgSalesMap.get(`${key}|${cleanSKU}`)) || 0;
      const forecast = avgSales * leadTime;
      const stockLevel = Number(item.StockLevel) || 0;
      const rawReorder = Math.max(0, Math.ceil(forecast - stockLevel));
      const deducedType = deduceSKUType(item.SKU, item.SKUType);

      storeDispatchCandidates.push({
        StoreID: storeID,
        City: storeCity,
        SKU: item.SKU,
        SKUType: deducedType,
        SIH: stockLevel,
        DailyAvgSales: avgSales,
        Forecast: forecast,
        RawReorder: rawReorder,
        ScaledDispatch: 0,
        TopUpAmount: 0,
        FinalDispatch: 0,
        FinalStock: stockLevel
      });
    });

    const totalRawReorder = storeDispatchCandidates.reduce((acc, c) => acc + c.RawReorder, 0);

    // Core Reorder calculation: fit within available physical space
    if (totalRawReorder <= availableSpace) {
      storeDispatchCandidates.forEach(c => {
        c.ScaledDispatch = c.RawReorder;
      });
    } else {
      const totalVelocity = storeDispatchCandidates.reduce((acc, c) => acc + c.DailyAvgSales, 0);
      
      storeDispatchCandidates.forEach(c => {
        const vel = c.DailyAvgSales;
        const weight = totalVelocity > 0 ? vel / totalVelocity : 1 / storeDispatchCandidates.length;
        c.ScaledDispatch = Math.max(0, Math.min(c.RawReorder, Math.floor(availableSpace * weight)));
      });

      // Distribute any remaining availableSpace due to rounding issues
      let currentDispatched = storeDispatchCandidates.reduce((acc, c) => acc + c.ScaledDispatch, 0);
      const unmetCandidates = [...storeDispatchCandidates]
        .filter(c => c.ScaledDispatch < c.RawReorder)
        .sort((a, b) => b.DailyAvgSales - a.DailyAvgSales);

      for (const item of unmetCandidates) {
        if (currentDispatched < availableSpace) {
          item.ScaledDispatch += 1;
          currentDispatched += 1;
        }
      }
    }

    // Top-Up Units: fill any remaining physical capacity with fastest-moving products
    const totalDispatched = storeDispatchCandidates.reduce((acc, c) => acc + c.ScaledDispatch, 0);
    const remainingTopUpSpace = Math.max(0, storeCapacity - totalSIH - totalDispatched);

    const topUpCandidates = [...storeDispatchCandidates]
      .filter(c => c.DailyAvgSales > 0)
      .sort((a, b) => b.DailyAvgSales - a.DailyAvgSales);

    const topUpMap = new Map<string, number>();
    if (remainingTopUpSpace > 0 && topUpCandidates.length > 0) {
      let spaceFilled = 0;
      while (spaceFilled < remainingTopUpSpace) {
        let incrementedAny = false;
        for (const item of topUpCandidates) {
          if (spaceFilled < remainingTopUpSpace) {
            topUpMap.set(item.SKU, (topUpMap.get(item.SKU) || 0) + 1);
            spaceFilled += 1;
            incrementedAny = true;
          }
        }
        if (!incrementedAny) break;
      }
    }

    // Finalize Candidate Fields
    storeDispatchCandidates.forEach(item => {
      const topUpAmount = topUpMap.get(item.SKU) || 0;
      item.TopUpAmount = topUpAmount;
      item.FinalDispatch = item.ScaledDispatch + topUpAmount;
      item.FinalStock = item.SIH + item.FinalDispatch;
      dispatchPlan.push(item);
    });
  });

  const storeSummaries: StoreSummary[] = validStoreKeys.map(key => {
    const freezer = freezerMap.get(key)!;
    const storeID = freezer.StoreID; // Keep original display name
    const storeItems = dispatchPlan.filter(p => cleanKey(p.StoreID) === key);
    const capacity = Number(freezer.Capacity) || 550;
    
    const initialStock = storeItems.reduce((acc, i) => acc + (Number(i.SIH) || 0), 0);
    const coreReorder = storeItems.reduce((acc, i) => acc + (Number(i.ScaledDispatch) || 0), 0);
    const topUpUnits = storeItems.reduce((acc, i) => acc + (Number(i.TopUpAmount) || 0), 0);
    const totalDispatch = coreReorder + topUpUnits;
    const forecast15Day = storeItems.reduce((acc, i) => acc + ((Number(i.DailyAvgSales) || 0) * 15), 0);
    const finalStock = initialStock + totalDispatch;

    // Calculate actual total physical units divided by overall freezer capacity
    const fillPercentage = capacity > 0 ? Math.min(100, (finalStock / capacity) * 100) : 100;

    return {
      StoreID: storeID,
      City: freezer.City || "Unknown",
      Capacity: capacity, // Keep Nominal All Mix Capacity
      InitialStock: initialStock,
      CoreReorder: coreReorder,
      TopUpUnits: topUpUnits,
      Forecast15Day: forecast15Day,
      FinalStock: finalStock,
      FillPercentage: fillPercentage
    };
  });

  const overallFillRate = storeSummaries.length > 0 ? storeSummaries.reduce((acc, s) => acc + (Number(s.FillPercentage) || 0), 0) / storeSummaries.length : 0;

  return { dispatchPlan, storeSummaries, overallFillRate };
};
