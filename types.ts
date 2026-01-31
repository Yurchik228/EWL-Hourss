
export interface DayPair {
  day: number;
  colA: number;
  colB: number;
  seg?: number;
}

export interface Segment {
  seg: number;
  dayMin: number;
  dayMax: number;
  countPairs: number;
}

export interface NameColumns {
  colLine: number;
  colLast: number;
  colFirst: number;
}

export interface SheetAnalysis {
  ok: boolean;
  err?: string;
  sheetName: string;
  fileName: string;
  grid?: any[][];
  shiftRow?: number;
  namesHeaderRow?: number;
  nameCols?: NameColumns;
  dataStartRow?: number;
  pairs?: DayPair[];
  segments?: Segment[];
}

export interface CalculationResult {
  Sheet: string;
  Linia: string;
  Imie: string;
  Nazwisko: string;
  Razem: number;
  DailyBreakdown: Record<number, number>;
  AbsenceBreakdown: Record<string, number>; 
}

export interface MasterRecord {
  fullName: string;
  hours: number;
}

export interface ComparisonItem {
  name: string;
  systemHours: number;
  masterHours: number;
  diff: number;
  status: 'ok' | 'mismatch' | 'missing_master' | 'missing_system';
  similarity?: number;
}

export interface CalculationStats {
  totalHours: number;
  totalPeople: number;
  skippedRows: number;
  skippedReasons: {
    emptyName: number;
    zeroHours: number;
    invalidFormat: number;
    inactive: number;
  };
}

export interface CalculationSummary {
  results: CalculationResult[];
  stats: CalculationStats;
  sheetDailyTotals: Record<string, Record<number, number>>;
}

export interface DayRange {
  from: number;
  to: number;
}
