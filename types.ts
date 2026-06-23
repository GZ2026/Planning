
export interface SIHRecord {
  StoreID: string;
  City: string;
  SKU: string;
  SKUType: string;
  StockLevel: number;
}

export interface SalesRecord {
  StoreID: string;
  City: string;
  SKU: string;
  SKUType: string;
  Date: string;
  Units: number;
}

export interface FreezerRecord {
  StoreID: string;
  City: string;
  Capacity: number;
  typeCapacities?: Record<string, number>;
}

export interface DispatchPlanItem {
  StoreID: string;
  City: string;
  SKU: string;
  SKUType: string;
  SIH: number;
  DailyAvgSales: number;
  Forecast: number;
  RawReorder: number;
  ScaledDispatch: number;
  TopUpAmount: number;
  FinalDispatch: number;
  FinalStock: number;
}

export interface StoreSummary {
  StoreID: string;
  City: string;
  Capacity: number;
  InitialStock: number;
  CoreReorder: number;
  TopUpUnits: number;
  Forecast15Day: number;
  FinalStock: number;
  FillPercentage: number;
}

export interface PlanningResults {
  dispatchPlan: DispatchPlanItem[];
  storeSummaries: StoreSummary[];
  overallFillRate: number;
}

export enum PlanningDay {
  MON_FRI = 'Mon-Fri',
  SATURDAY = 'Saturday'
}
