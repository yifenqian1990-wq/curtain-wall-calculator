export type MaterialType = 'profile' | 'panel' | 'steel' | 'gasket' | 'fastener' | 'auxiliary' | 'other';

export interface ColumnMappings {
  categoryCol: number | null;
  nameCol: number;
  specCol: number;
  materialCol: number | null;
  qtyCol: number;
  lenCol: number | null;
  unitCol: number;
  remarkCol: number | null;
}

export interface RangeConfiguration {
  profileRange: string; // Dynamic cell coordinates like C8:L57
  steelRange: string;   // e.g. C58:L67
  panelRange: string;   // e.g. C68:L97
  profileKeyword?: string; // B-Column search keyword for profiles, default "铝型材"
  steelKeyword?: string;   // B-Column search keyword for steel components, default "钢件"
  panelKeyword?: string;   // B-Column search keyword for panels, default "面板"
}

export interface COMConnectionState {
  isConnected: boolean;
  provider: 'excel' | 'wps' | 'none';
  host: string;
  port: number;
  activeWorkbookName: string;
  isSending: boolean;
}

export interface AuxiliaryLedgerItemRule {
  id: string;
  name: string;
  expression: string;
  isActive: boolean;
  description?: string;
  isExclusive?: boolean;
  lengthAdjustment?: number; // 加长尺寸 (mm)
}

export interface AuxiliaryLedgerItem {
  id: string;
  category: 'gasket' | 'fastener' | 'auxiliary';
  name: string;
  drawingNo: string;       // 加工图号
  materialModel: string;   // 材质及型号
  size: string;            // 尺寸
  unit: string;            // 单位
  remark: string;          // 备注
  position: string;        // 位置
  rules?: AuxiliaryLedgerItemRule[];
}

export interface ProcessedItem {
  rowIndex: number;
  originalCategory: string; // From the Excel merged cells or row values
  category: MaterialType;   // Our unified category
  
  // Clean original Excel columns (C to L)
  name: string;             // Col C: 名称
  drawingNo: string;        // Col D: 加工图号
  materialModel: string;    // Col E: 材质及型号
  cuttingSize: string;      // Col F: 下料尺寸 (L) (mm)
  unitUsage: string;        // Col G: 单樘用量
  qty: number;              // Col H: 总计
  unit: string;             // Col I: 单位
  l1: string;               // Col J: L1
  l2: string;               // Col K: L2
  remark: string;           // Col L: 备注
  
  // Legacy / compatible properties for formula engine
  spec: string;             // Compatibility field
  material: string;         // Compatibility field
  len: number;              // Compatibility field (parsed meters)

  // For calculated rows
  formula?: string; // Active calculation expression
  calculatedQty?: number;
  isCalculated?: boolean;
  matchLog?: string[]; // Tracing which source materials were matched
}

export interface FormulaRule {
  id: string;
  name: string;
  targetCategory: 'gasket' | 'fastener' | 'auxiliary';
  targetKeyword: string; // Keyword to match item name or spec
  targetSpecKeyword?: string; // Optional specification keyword to match
  expression: string; // The formula, e.g. "SUM_LEN('H2-JT') * 1.05"
  description: string;
}

export interface CalculationResult {
  rowIndex: number;
  originalName: string;
  originalSpec: string;
  formula: string;
  calculatedValue: number;
  log: string[];
}

export interface AIAnalysisResponse {
  headerRowIndex: number;
  columnMappings: {
    categoryCol: number | null;
    nameCol: number | null;
    specCol: number | null;
    materialCol: number | null;
    qtyCol: number | null;
    lenCol: number | null;
    unitCol: number | null;
    remarkCol: number | null;
  };
  rowClassifications: {
    rowIndex: number;
    type: MaterialType;
    name: string;
    spec: string;
    qty: number;
    len: number;
    unit: string;
    material: string;
    recommendedFormula?: string;
  }[];
}

export interface SteelPanelItemDetail {
  unitNo: string;
  rowSeq: string | number;
  categoryName: string; // '钢件' | '面板'
  name: string;
  drawingNo: string;
  materialModel: string;
  cuttingSize: string;
  unitUsage: string | number;
  qty: number;
  unit: string;
  l1: string;
  l2: string;
  remark: string;
}


