import React, { useState, useEffect, useMemo } from "react";
import XLSX from "xlsx-js-style";
import { 
  BarChart3, FileSpreadsheet, Download, RefreshCw, 
  Search, Layers, CheckCircle2, AlertCircle, Sparkles, 
  Table, ChevronRight, Filter, PieChart, Box, Settings,
  X, FileCheck
} from "lucide-react";
import { ProcessedItem, RangeConfiguration, AuxiliaryLedgerItem, SteelPanelItemDetail } from "../types";
import { parseExcelRange, resolveRangeByBColumnKeyword, sanitizeExcelValue, evaluateFormula, isAuxiliaryItem } from "../utils/calcEngine";

interface ProfileItemDetail {
  unitNo: string;
  rowSeq: string | number;
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

interface ProfileSummaryItem {
  id: string;
  seq: number;
  name: string;
  drawingNo: string;
  materialModel: string;
  cuttingSize: string;
  totalQty: number;
  unit: string;
  l1: string;
  l2: string;
  unitCount: number;
  unitsInvolvedList: string[];
  remark: string;
}

interface XuKaSummaryItem {
  seq: number;
  name: string;
  drawingNo: string;
  materialModel: string;
  cuttingSize: string;
  totalQty: number;
  unit: string;
  l1: string;
  l2: string;
  remark: string;
}

interface AuxSummaryItem {
  id: string;
  seq: number;
  category: string;
  name: string;
  drawingNo: string;
  materialModel: string;
  size: string;
  totalQty: number;
  unit: string;
  unitCount: number;
  unitsInvolvedList: string[];
  remark: string;
}

interface DataSummaryModuleProps {
  workbook: XLSX.WorkBook | null;
  excelSheets: string[];
  wsConnected: boolean;
  pullLoading: boolean;
  ranges: RangeConfiguration;
  ledger: AuxiliaryLedgerItem[];
  readSheetPromise?: (sheetName: string) => Promise<any>;
}

const DEFAULT_F2_FORMULA = "=@INDEX('D:\\在建工程\\@202604同济益田广场\\下单\\王成\\02.型材\\[型材下单汇总.xlsx]型材'!$F$4:$F$10000,MATCH(C2,'D:\\在建工程\\@202604同济益田广场\\下单\\王成\\02.型材\\[型材下单汇总.xlsx]型材'!$C$4:$C$10000,0))";

// Helper function to extract unit frame count (樘数/单元数量) for a sheet from metadata rows or Catalog index sheet
const extractSheetUnitCount = (
  sheetName: string, 
  sheetRows: string[][], 
  catalogUnitMap: Record<string, number>
): number => {
  // 1. Check Catalog index sheet map first
  if (catalogUnitMap[sheetName] && catalogUnitMap[sheetName] > 0) {
    return catalogUnitMap[sheetName];
  }
  const cleanSheetName = sheetName.replace(/\s+/g, "").toLowerCase();
  for (const [k, v] of Object.entries(catalogUnitMap)) {
    const cleanK = k.replace(/\s+/g, "").toLowerCase();
    if (cleanK === cleanSheetName || cleanSheetName.includes(cleanK) || cleanK.includes(cleanSheetName)) {
      if (v > 0) return v;
    }
  }

  // Keywords indicating block/unit quantity in engineering sheets
  const quantityKeywords = [
    "板块数量", "工程数量", "单元数量", "樘数", "板块数", "总数量", 
    "图纸数量", "重复数量", "工程套数", "板块件数", "单元数", "套数", "数量"
  ];

  // 2. Scan top 15 metadata rows cell by cell
  for (let r = 0; r < Math.min(15, sheetRows.length); r++) {
    const row = sheetRows[r];
    if (!row) continue;

    for (let c = 0; c < row.length; c++) {
      const cellVal = String(row[c] || "").trim();
      if (!cellVal) continue;

      const matchedKw = quantityKeywords.find(kw => cellVal.includes(kw));
      if (matchedKw) {
        // A) Check if cellVal itself contains the number (e.g. "板块数量：12", "板块数量 12", "工程数量12樘")
        const inCellMatch = cellVal.match(/(?:板块数量|工程数量|单元数量|樘数|板块数|总数量|图纸数量|重复数量|工程套数|板块件数|单元数|套数|数量)\s*[:：=]?\s*(\d+(?:\.\d+)?)/i);
        if (inCellMatch && inCellMatch[1]) {
          const num = Math.round(parseFloat(inCellMatch[1]));
          if (num > 0) return num;
        }

        // B) Check adjacent right cells (col + 1, col + 2, col + 3)
        for (let offset = 1; offset <= 3; offset++) {
          if (c + offset < row.length) {
            const adjVal = String(row[c + offset] || "").trim();
            if (adjVal) {
              const adjMatch = adjVal.match(/^(\d+(?:\.\d+)?)\s*(?:樘|套|个|块|面|幅|组)?$/);
              if (adjMatch && adjMatch[1]) {
                const num = Math.round(parseFloat(adjMatch[1]));
                if (num > 0) return num;
              }
            }
          }
        }

        // C) Check cell directly below (row + 1, col)
        if (r + 1 < sheetRows.length && sheetRows[r + 1][c]) {
          const belowVal = String(sheetRows[r + 1][c] || "").trim();
          const belowMatch = belowVal.match(/^(\d+(?:\.\d+)?)\s*(?:樘|套|个|块|面|幅|组)?$/);
          if (belowMatch && belowMatch[1]) {
            const num = Math.round(parseFloat(belowMatch[1]));
            if (num > 0) return num;
          }
        }
      }
    }

    // D) Fallback regex on whole row string
    const rowStr = row.join(" ");
    const keyMatch = rowStr.match(/(?:板块数量|工程数量|单元数量|樘数|板块数|总数量|图纸数量|重复数量|工程套数|板块件数|单元数|套数|数量)\s*[:：=]?\s*(\d+(?:\.\d+)?)/i);
    if (keyMatch && keyMatch[1]) {
      const num = Math.round(parseFloat(keyMatch[1]));
      if (num > 0) return num;
    }

    const unitMatch = rowStr.match(/(\d+(?:\.\d+)?)\s*(?:樘|套|个|块|面)/);
    if (unitMatch && unitMatch[1]) {
      const num = Math.round(parseFloat(unitMatch[1]));
      if (num > 0) return num;
    }
  }

  // Default to 1 if no explicit unit quantity is found
  return 1;
};

// Helper function to safely format cutting size without destroying dimension strings like "1200*1500" or "300*120"
const formatCuttingSize = (val: string): string => {
  if (!val) return "";
  const str = String(val).trim();
  if (str.startsWith("-")) return "";

  // If it's a multi-dimensional string with delimiters (*, x, X, ×), preserve the dimension format
  if (/[*xX×]/.test(str)) {
    return str.replace(/\s+/g, "");
  }

  // Single numeric length (e.g. "2620" or "2620.5")
  const num = parseFloat(str.replace(/[^\d.]/g, ""));
  if (!isNaN(num)) {
    return Number.isInteger(num) ? String(num) : String(Math.round(num * 10) / 10);
  }
  return str;
};

export default function DataSummaryModule({
  workbook,
  excelSheets,
  wsConnected,
  pullLoading,
  ranges,
  ledger,
  readSheetPromise,
}: DataSummaryModuleProps) {
  const [activeSubTab, setActiveSubTab] = useState<'profiles_detail' | 'profiles_summary' | 'xuka_summary' | 'aux_summary' | 'steel_panel_detail'>('profiles_detail');
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [selectedSheetsState, setSelectedSheetsState] = useState<Record<string, boolean>>({});
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [scanProgress, setScanProgress] = useState<{ current: number; total: number; currentSheet: string }>({ current: 0, total: 0, currentSheet: "" });

  // Custom persistent formula for F2 in Profile Summary sheet
  const [f2Formula, setF2Formula] = useState<string>(() => {
    return localStorage.getItem("profile_summary_f2_formula") ?? DEFAULT_F2_FORMULA;
  });

  const handleFormulaChange = (val: string) => {
    setF2Formula(val);
    localStorage.setItem("profile_summary_f2_formula", val);
  };

  // Extracted raw and aggregated data state
  const [profileDataList, setProfileDataList] = useState<ProfileItemDetail[]>([]);
  const [steelPanelDataList, setSteelPanelDataList] = useState<SteelPanelItemDetail[]>([]);
  const [profileSummaryList, setProfileSummaryList] = useState<ProfileSummaryItem[]>([]);
  const [xuKaSummaryList, setXuKaSummaryList] = useState<XuKaSummaryItem[]>([]);
  const [auxSummaryList, setAuxSummaryList] = useState<AuxSummaryItem[]>([]);
  const [scannedUnitsCount, setScannedUnitsCount] = useState<number>(0);
  const [showCompletionModal, setShowCompletionModal] = useState<boolean>(false);
  const [lastCompletedTime, setLastCompletedTime] = useState<string | null>(null);

  // Compute combined allSheetNames prioritizing comprehensive excelSheets list over single-sheet workbook
  const allSheetNames = useMemo(() => {
    if (excelSheets && excelSheets.length > 0) {
      if (!workbook || workbook.SheetNames.length <= 1) {
        return excelSheets;
      }
      return Array.from(new Set([...excelSheets, ...workbook.SheetNames]));
    }
    return workbook ? workbook.SheetNames : [];
  }, [workbook, excelSheets]);

  // Initialize list of valid sheets (excluding "目录")
  useEffect(() => {
    const initial: Record<string, boolean> = {};
    allSheetNames.forEach(sheet => {
      const isDirectory = sheet.includes("目录") || sheet.toLowerCase().includes("index");
      initial[sheet] = !isDirectory;
    });
    setSelectedSheetsState(initial);
  }, [allSheetNames]);

  const toggleAllSheets = (select: boolean) => {
    const updated: Record<string, boolean> = {};
    allSheetNames.forEach(s => {
      updated[s] = select;
    });
    setSelectedSheetsState(updated);
  };

  const toggleExcludeDirectorySheets = () => {
    const updated: Record<string, boolean> = {};
    allSheetNames.forEach(s => {
      const isDirectory = s.includes("目录") || s.toLowerCase().includes("index");
      updated[s] = !isDirectory;
    });
    setSelectedSheetsState(updated);
  };

  const toggleSingleSheet = (sheetName: string) => {
    setSelectedSheetsState(prev => ({
      ...prev,
      [sheetName]: !prev[sheetName]
    }));
  };

  // Main aggregation algorithm
  const runAggregation = async () => {
    setIsScanning(true);
    const allSheets = allSheetNames;
    let activeSheets = allSheets.filter(s => selectedSheetsState[s]);

    if (activeSheets.length === 0) {
      // Fallback: If selectedSheetsState hasn't flushed state yet on mount, default to non-directory sheets
      const nonDirectorySheets = allSheets.filter(s => !s.includes("目录") && !s.toLowerCase().includes("index"));
      if (nonDirectorySheets.length > 0) {
        activeSheets = nonDirectorySheets;
      } else if (allSheets.length > 0) {
        activeSheets = allSheets;
      } else {
        alert("请至少勾选一个有效工作表进行数据汇总统计！");
        setIsScanning(false);
        return;
      }
    }

    // Step A: Pre-scan Catalog / Index sheets (00-工程图纸目录 or 目录) to build a mapping of sheetName -> unitCount (工程数量/樘数)
    const catalogUnitMap: Record<string, number> = {};
    const catalogSheets = allSheets.filter(s => s.includes("目录") || s.toLowerCase().includes("index"));

    for (const catSheet of catalogSheets) {
      let catRows: string[][] = [];
      if (workbook && workbook.Sheets[catSheet]) {
        const jsonRows = XLSX.utils.sheet_to_json<string[]>(workbook.Sheets[catSheet], { header: 1, defval: "" });
        catRows = jsonRows.map(r => (Array.isArray(r) ? r.map(c => sanitizeExcelValue(c)) : []));
      } else if (wsConnected && readSheetPromise) {
        try {
          const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error("读取目录超时 (5s)")), 5000));
          const res: any = await Promise.race([readSheetPromise(catSheet), timeoutPromise]);
          if (res && res.success && res.rows) {
            catRows = res.rows.map((r: any[]) => r.map((c: any) => sanitizeExcelValue(c)));
          }
        } catch (err) {
          console.warn(`Skip catalog sheet ${catSheet}:`, err);
        }
      }

      // Find header row and quantity column index in catalog sheet
      let qtyColIdx = -1;
      let nameColIdx = -1;

      for (let r = 0; r < Math.min(10, catRows.length); r++) {
        const row = catRows[r];
        if (!row) continue;
        for (let c = 0; c < row.length; c++) {
          const cellStr = String(row[c] || "").trim();
          if (["数量", "工程数量", "樘数", "板块数量", "套数", "件数", "块数", "单元数量"].some(k => cellStr.includes(k))) {
            qtyColIdx = c;
          }
          if (["图号", "图纸编号", "板块名称", "单元名称", "图纸名称", "部件名称", "板块型号", "名称"].some(k => cellStr.includes(k))) {
            nameColIdx = c;
          }
        }
        if (qtyColIdx !== -1) break;
      }

      catRows.forEach(row => {
        if (!row || row.length < 2) return;
        const rowText = row.join(" ");

        for (const targetSheet of activeSheets) {
          const cleanTarget = targetSheet.replace(/\s+/g, "").toLowerCase();
          const matchesSheet = rowText.toLowerCase().includes(cleanTarget) || 
                               rowText.replace(/\s+/g, "").toLowerCase().includes(cleanTarget) ||
                               (nameColIdx >= 0 && row[nameColIdx] && row[nameColIdx].replace(/\s+/g, "").toLowerCase().includes(cleanTarget));

          if (matchesSheet) {
            // If explicit quantity column index was found, use it!
            if (qtyColIdx >= 0 && row[qtyColIdx]) {
              const cellVal = String(row[qtyColIdx]).trim();
              const m = cellVal.match(/(\d+(?:\.\d+)?)/);
              if (m && parseFloat(m[1]) > 0) {
                catalogUnitMap[targetSheet] = Math.round(parseFloat(m[1]));
                continue;
              }
            }

            // Fallback: scan backwards from end of row to avoid matching drawing numbers or sizes
            for (let c = row.length - 1; c >= 2; c--) {
              const cellVal = String(row[c]).trim();
              if (!cellVal) continue;
              // Skip dimensions like 1200*2400 or drawing numbers like DWG-01 or long text
              if (/[*xX×]/.test(cellVal) || /[a-zA-Z]/.test(cellVal) || cellVal.length > 6) continue;
              const m = cellVal.match(/^(\d+(?:\.\d+)?)\s*(?:樘|套|个|块|面)?$/);
              if (m && parseFloat(m[1]) > 0) {
                catalogUnitMap[targetSheet] = Math.round(parseFloat(m[1]));
                break;
              }
            }
          }
        }
      });
    }

    const rawProfiles: ProfileItemDetail[] = [];
    const rawSteelPanels: SteelPanelItemDetail[] = [];
    const auxMap: Record<string, {
      category: string;
      name: string;
      drawingNo: string;
      materialModel: string;
      size: string;
      unit: string;
      totalQty: number;
      unitsSet: Set<string>;
      remarksSet: Set<string>;
    }> = {};

    let processedCount = 0;

    for (const sheetName of activeSheets) {
      processedCount++;
      setScanProgress({ current: processedCount, total: activeSheets.length, currentSheet: sheetName });

      let sheetRows: string[][] = [];
      let ws: XLSX.WorkSheet | null = null;

      if (workbook && workbook.Sheets[sheetName]) {
        ws = workbook.Sheets[sheetName];
        const jsonRows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1, defval: "" });
        sheetRows = jsonRows.map(r => (Array.isArray(r) ? r.map(c => sanitizeExcelValue(c)) : []));
      } else if (wsConnected && readSheetPromise) {
        try {
          const res = await readSheetPromise(sheetName);
          if (res && res.success && res.rows) {
            sheetRows = res.rows.map((r: any[]) => r.map((c: any) => sanitizeExcelValue(c)));
            ws = XLSX.utils.aoa_to_sheet(sheetRows);
          }
        } catch (err) {
          console.error(`Failed to read sheet ${sheetName}:`, err);
        }
      }

      if (!sheetRows || sheetRows.length === 0) continue;

      // Extract Sheet Unit Count (樘数 / 单元工程数量)
      const sheetUnitCount = extractSheetUnitCount(sheetName, sheetRows, catalogUnitMap);

      // Extract Unit Serial Number (单元编号)
      let unitNo = sheetName;
      for (let r = 0; r < Math.min(5, sheetRows.length); r++) {
        const rowStr = sheetRows[r].join(" ");
        if (rowStr.includes("板块型号:") || rowStr.includes("单元编号:")) {
          const match = rowStr.match(/(?:板块型号|单元编号)\s*[:：]\s*([A-Za-z0-9_#-]+)/);
          if (match && match[1]) {
            unitNo = `${sheetName} (${match[1]})`;
            break;
          }
        }
      }

      // 1. Robust Section Range Resolution supporting broad keyword variants
      const profileKeywords = Array.from(new Set([ranges.profileKeyword, "铝型材", "型材", "铝材", "铝  型  材", "型  材", "铝合金型材"].filter(Boolean) as string[]));
      const steelKeywords = Array.from(new Set([ranges.steelKeyword, "钢件", "钢件部分", "钢  件", "钢卡", "钢结构", "钢   件"].filter(Boolean) as string[]));
      const panelKeywords = Array.from(new Set([ranges.panelKeyword, "面板", "玻璃面板", "面  板", "玻璃", "采光面板"].filter(Boolean) as string[]));

      let pResolved: { startRow: number; endRow: number } | null = null;
      for (const pk of profileKeywords) {
        pResolved = resolveRangeByBColumnKeyword(ws, sheetRows, pk);
        if (pResolved) break;
      }

      let sResolved: { startRow: number; endRow: number } | null = null;
      for (const sk of steelKeywords) {
        sResolved = resolveRangeByBColumnKeyword(ws, sheetRows, sk);
        if (sResolved) break;
      }

      let paResolved: { startRow: number; endRow: number } | null = null;
      for (const pak of panelKeywords) {
        paResolved = resolveRangeByBColumnKeyword(ws, sheetRows, pak);
        if (paResolved) break;
      }

      // Check for auxiliary material section starting rows
      const gasketKeywords = ["型材气密胶条类", "胶条", "胶条类", "气密胶条", "三元乙丙", "密封胶条"];
      const fastenerKeywords = ["不锈钢紧固件类", "紧固件", "紧固件类", "五金螺栓", "螺栓", "紧固件及螺栓"];
      const auxKeywords = ["结构密封胶及辅材类", "辅材", "辅材类", "辅料", "密封胶及辅材", "辅助材料"];

      let gasketResolved: { startRow: number; endRow: number } | null = null;
      for (const kw of gasketKeywords) {
        gasketResolved = resolveRangeByBColumnKeyword(ws, sheetRows, kw);
        if (gasketResolved) break;
      }

      let fastenerResolved: { startRow: number; endRow: number } | null = null;
      for (const kw of fastenerKeywords) {
        fastenerResolved = resolveRangeByBColumnKeyword(ws, sheetRows, kw);
        if (fastenerResolved) break;
      }

      let auxResolved: { startRow: number; endRow: number } | null = null;
      for (const kw of auxKeywords) {
        auxResolved = resolveRangeByBColumnKeyword(ws, sheetRows, kw);
        if (auxResolved) break;
      }

      let earliestAuxStart = 9999;
      if (gasketResolved && gasketResolved.startRow < earliestAuxStart) earliestAuxStart = gasketResolved.startRow;
      if (fastenerResolved && fastenerResolved.startRow < earliestAuxStart) earliestAuxStart = fastenerResolved.startRow;
      if (auxResolved && auxResolved.startRow < earliestAuxStart) earliestAuxStart = auxResolved.startRow;

      // Scan rows to find any auxiliary section headers in Column B or A
      for (let r = 0; r < sheetRows.length; r++) {
        const rVal = sheetRows[r];
        if (!rVal) continue;
        const bVal = rVal[1] ? String(rVal[1]).trim() : "";
        const aVal = rVal[0] ? String(rVal[0]).trim() : "";
        const combined = (bVal + " " + aVal).replace(/\s+/g, "").toLowerCase();
        const hasAuxKeyword = [
          "胶条", "紧固件", "辅材", "辅料", "密封胶", "五金", "螺栓", "epdm", "配件", "三元乙丙", "双面贴"
        ].some(k => combined.includes(k));
        if (hasAuxKeyword && r < earliestAuxStart) {
          earliestAuxStart = r;
        }
      }

      const pRange = pResolved ? { startRow: pResolved.startRow, endRow: pResolved.endRow } : parseExcelRange(ranges.profileRange);
      const sRange = sResolved ? { startRow: sResolved.startRow, endRow: sResolved.endRow } : parseExcelRange(ranges.steelRange);
      const paRange = paResolved ? { startRow: paResolved.startRow, endRow: paResolved.endRow } : parseExcelRange(ranges.panelRange);

      // Capping heuristic: Profiles, steel parts, and panels never continue after auxiliary sections begin
      if (earliestAuxStart !== 9999) {
        if (pRange && pRange.endRow >= earliestAuxStart) {
          pRange.endRow = earliestAuxStart - 1;
        }
        if (sRange && sRange.endRow >= earliestAuxStart) {
          sRange.endRow = earliestAuxStart - 1;
        }
        if (paRange && paRange.endRow >= earliestAuxStart) {
          paRange.endRow = earliestAuxStart - 1;
        }
      }

      const sheetExtractedItems: ProcessedItem[] = [];

      for (let rIdx = 0; rIdx < sheetRows.length; rIdx++) {
        const row = sheetRows[rIdx];
        if (!row || row.filter(c => c && String(c).trim() !== "").length === 0) continue;

        const name = sanitizeExcelValue(row[2]);
        const drawingNo = sanitizeExcelValue(row[3]);
        const materialModel = sanitizeExcelValue(row[4]);
        const remark = sanitizeExcelValue(row[11]);

        // Determine category: exact range check OR heuristic fallback if ranges are unresolvable
        let isProfileRow = false;
        let isSteelRow = false;
        let isPanelRow = false;

        // Force false for profiles, steel parts, and panels if row index is on/after earliest auxiliary section
        const isPastAuxiliaryStart = earliestAuxStart !== 9999 && rIdx >= earliestAuxStart;

        if (!isPastAuxiliaryStart) {
          if (pRange && rIdx >= pRange.startRow && rIdx <= pRange.endRow) {
            isProfileRow = true;
          } else if (sRange && rIdx >= sRange.startRow && rIdx <= sRange.endRow) {
            isSteelRow = true;
          } else if (paRange && rIdx >= paRange.startRow && rIdx <= paRange.endRow) {
            isPanelRow = true;
          } else {
            // Heuristic fallback if ranges are not explicitly resolved
            const rowStr = row.join(" ");
            const isAuxHeader = ["胶条", "紧固件", "辅材", "五金", "密封胶", "二、", "三、", "四、"].some(k => rowStr.includes(k));
            const isFooter = ["页脚", "备注:", "注:", "编制:", "审核:"].some(k => rowStr.includes(k));
            const isHeader = ["名称", "品名", "大类", "铝型材", "型材部分"].some(k => rowStr.includes(k));

            if (!isAuxHeader && !isFooter && !isHeader && rIdx >= 5) {
              if (sRange && rIdx >= sRange.startRow) {
                isSteelRow = true;
              } else if (paRange && rIdx >= paRange.startRow) {
                isPanelRow = true;
              } else {
                isProfileRow = true;
              }
            }
          }
        }

        if (!name || name === "名称" || name === "品名" || name === "大类" || name === "页脚" || name === "备注") continue;

        const rawCutSize = String(sanitizeExcelValue(row[5])).trim();
        const cuttingSize = formatCuttingSize(rawCutSize);
        const unitUsage = sanitizeExcelValue(row[6]);
        const qtyStr = sanitizeExcelValue(row[7]);

        let unitUsageNum = parseFloat(unitUsage.replace(/[^\d.]/g, "")) || 0;
        const rawQtyNum = parseFloat(qtyStr.replace(/[^\d.]/g, "")) || 0;

        // If unit usage is missing or 0 but material name and cutting size exist, default unit usage to 1 per unit
        if (unitUsageNum === 0 && rawQtyNum === 0 && (cuttingSize || materialModel)) {
          unitUsageNum = 1;
        }

        // Calculate total quantity accurately:
        // 1. If total quantity column (Col H) is filled with total > 0, use it.
        // 2. If Col H is empty/0 OR equals single unit usage (Col G) while sheetUnitCount > 1, multiply unitUsage * sheetUnitCount
        let qty = rawQtyNum;
        if (qty === 0 && unitUsageNum > 0) {
          qty = unitUsageNum * sheetUnitCount;
        } else if (qty > 0 && qty === unitUsageNum && sheetUnitCount > 1) {
          qty = unitUsageNum * sheetUnitCount;
        }
        qty = Math.round(qty * 100) / 100;

        const unit = sanitizeExcelValue(row[8]);
        const l1 = sanitizeExcelValue(row[9]);
        const l2 = sanitizeExcelValue(row[10]);

        let len = 0;
        if (cuttingSize) {
          const cleanCutSize = cuttingSize.replace(/[^\d.]/g, "");
          const sizeInMm = parseFloat(cleanCutSize) || 0;
          len = sizeInMm > 0 ? sizeInMm / 1000 : 0;
        }

        if (isProfileRow) {
          rawProfiles.push({
            unitNo,
            rowSeq: rawProfiles.filter(p => p.unitNo === unitNo).length + 1,
            name,
            drawingNo,
            materialModel,
            cuttingSize,
            unitUsage: unitUsage || (unitUsageNum > 0 ? String(unitUsageNum) : "1"),
            qty,
            unit: unit || "支",
            l1,
            l2,
            remark
          });
        }

        if (isSteelRow || isPanelRow) {
          rawSteelPanels.push({
            unitNo,
            rowSeq: rawSteelPanels.filter(p => p.unitNo === unitNo).length + 1,
            categoryName: isSteelRow ? "钢件" : "面板",
            name,
            drawingNo,
            materialModel,
            cuttingSize,
            unitUsage: unitUsage || (unitUsageNum > 0 ? String(unitUsageNum) : "1"),
            qty,
            unit: unit || (isSteelRow ? "个" : "㎡"),
            l1,
            l2,
            remark
          });
        }

        // Keep item for aux calculation engine
        sheetExtractedItems.push({
          rowIndex: rIdx,
          originalCategory: isProfileRow ? '型材' : (isSteelRow ? '钢件' : (isPanelRow ? '面板' : '其它')),
          category: isProfileRow ? 'profile' : (isSteelRow ? 'steel' : (isPanelRow ? 'panel' : 'other')),
          name,
          drawingNo,
          materialModel,
          cuttingSize,
          unitUsage,
          qty,
          unit,
          l1,
          l2,
          remark,
          spec: materialModel,
          material: materialModel,
          len
        });
      }

      // 2. Auxiliary materials aggregation for this sheet
      // Option A: Extract raw auxiliary items directly present in raw sheet under "胶条", "紧固件", "辅材"
      const rawExtractedAuxKeys = new Set<string>();

      const extractRawAux = (range: { startRow: number; endRow: number } | null, categoryLabel: string) => {
        if (!range) return;
        for (let r = range.startRow; r <= range.endRow; r++) {
          const row = sheetRows[r];
          if (!row) continue;
          const name = sanitizeExcelValue(row[2]);
          if (!name || name === "名称" || name === "品名") continue;

          const drawingNo = sanitizeExcelValue(row[3]);
          const materialModel = sanitizeExcelValue(row[4]);
          const size = sanitizeExcelValue(row[5]);
          
          const rawQtyH = parseFloat(sanitizeExcelValue(row[7]).replace(/[^\d.]/g, "")) || 0;
          const rawQtyG = parseFloat(sanitizeExcelValue(row[6]).replace(/[^\d.]/g, "")) || 0;

          let qtyVal = rawQtyH;
          if (qtyVal === 0 && rawQtyG > 0) {
            qtyVal = rawQtyG * sheetUnitCount;
          } else if (qtyVal > 0 && qtyVal === rawQtyG && sheetUnitCount > 1) {
            qtyVal = rawQtyG * sheetUnitCount;
          }

          const unit = sanitizeExcelValue(row[8]) || "项";
          const remark = sanitizeExcelValue(row[11]);

          if (qtyVal > 0) {
            const auxKey = `${categoryLabel}_${name.trim()}_${drawingNo.trim()}_${materialModel.trim()}_${size.trim()}_${unit.trim()}`;
            rawExtractedAuxKeys.add(`${name.trim()}|||${drawingNo.trim()}`);

            if (!auxMap[auxKey]) {
              auxMap[auxKey] = {
                category: categoryLabel,
                name,
                drawingNo,
                materialModel,
                size,
                unit,
                totalQty: 0,
                unitsSet: new Set(),
                remarksSet: new Set()
              };
            }
            auxMap[auxKey].totalQty += qtyVal;
            auxMap[auxKey].unitsSet.add(unitNo);
            if (remark) auxMap[auxKey].remarksSet.add(remark);
          }
        }
      };

      extractRawAux(gasketResolved, "型材气密胶条类");
      extractRawAux(fastenerResolved, "不锈钢紧固件类");
      extractRawAux(auxResolved, "结构密封胶及辅材类");

      // Option B: Calculated auxiliary materials using active ledger rules if configured
      if (ledger && ledger.length > 0) {
        ledger.forEach(item => {
          const activeRules = (item.rules || []).filter(r => r.isActive);
          if (activeRules.length > 0) {
            let unitUsage = 0;
            activeRules.forEach(rule => {
              const { value } = evaluateFormula(rule.expression, sheetExtractedItems, {}, rule.lengthAdjustment || 0, sheetRows);
              unitUsage += value;
            });
            
            if (unitUsage > 0) {
              // Calculate total quantity across all units (樘数) in this sheet
              const sheetTotalQty = Math.round(unitUsage * sheetUnitCount * 100) / 100;
              const catLabel = item.category === 'gasket' ? '型材气密胶条类' : (item.category === 'fastener' ? '不锈钢紧固件类' : '结构密封胶及辅材类');
              const auxKey = `${catLabel}_${item.name.trim()}_${item.drawingNo.trim()}_${item.materialModel.trim()}_${item.size.trim()}_${item.unit.trim()}`;
              
              // Prevent double counting if the item was already extracted from raw sheet rows
              const rawKey = `${item.name.trim()}|||${item.drawingNo.trim()}`;
              if (!rawExtractedAuxKeys.has(rawKey)) {
                if (!auxMap[auxKey]) {
                  auxMap[auxKey] = {
                    category: catLabel,
                    name: item.name,
                    drawingNo: item.drawingNo,
                    materialModel: item.materialModel,
                    size: item.size,
                    unit: item.unit,
                    totalQty: 0,
                    unitsSet: new Set(),
                    remarksSet: new Set()
                  };
                }
                auxMap[auxKey].totalQty += sheetTotalQty;
                auxMap[auxKey].unitsSet.add(unitNo);
                if (item.remark) auxMap[auxKey].remarksSet.add(item.remark);
              }
            }
          }
        });
      }
    }

    // Process Profile Summary Table (Grouping by 材质及型号 & 下料尺寸)
    const profileSummaryMap: Record<string, {
      namesSet: Set<string>;
      drawingNosSet: Set<string>;
      materialModel: string;
      cuttingSize: string;
      unit: string;
      totalQty: number;
      unitsSet: Set<string>;
      remarksSet: Set<string>;
      l1Set: Set<string>;
      l2Set: Set<string>;
    }> = {};

    rawProfiles.forEach(p => {
      // Grouping key strictly groups by Material & Model (材质及型号) and Cut Size (下料尺寸) for nesting/cutting optimization
      const cleanModel = (p.materialModel || "").trim();
      const cleanCutSize = (p.cuttingSize || "").trim();
      const groupKey = `${cleanModel.toLowerCase()}|||${cleanCutSize.toLowerCase()}`;

      if (!profileSummaryMap[groupKey]) {
        profileSummaryMap[groupKey] = {
          namesSet: new Set(),
          drawingNosSet: new Set(),
          materialModel: cleanModel,
          cuttingSize: cleanCutSize,
          unit: p.unit || "支",
          totalQty: 0,
          unitsSet: new Set(),
          remarksSet: new Set(),
          l1Set: new Set(),
          l2Set: new Set()
        };
      }

      if (p.name) profileSummaryMap[groupKey].namesSet.add(p.name);
      if (p.drawingNo) profileSummaryMap[groupKey].drawingNosSet.add(p.drawingNo);
      profileSummaryMap[groupKey].totalQty += p.qty;
      profileSummaryMap[groupKey].unitsSet.add(p.unitNo);
      if (p.remark) profileSummaryMap[groupKey].remarksSet.add(p.remark);
      if (p.l1) profileSummaryMap[groupKey].l1Set.add(p.l1);
      if (p.l2) profileSummaryMap[groupKey].l2Set.add(p.l2);
    });

    const summaryProfilesList: ProfileSummaryItem[] = Object.values(profileSummaryMap).map((item, idx) => ({
      id: `psum-${idx}`,
      seq: idx + 1,
      name: Array.from(item.namesSet).find(n => n && n.trim() !== "") || "型材",
      drawingNo: Array.from(item.drawingNosSet).join("；"),
      materialModel: item.materialModel,
      cuttingSize: item.cuttingSize,
      totalQty: Math.round(item.totalQty * 100) / 100,
      unit: item.unit,
      l1: Array.from(item.l1Set).join("；"),
      l2: Array.from(item.l2Set).join("；"),
      unitCount: item.unitsSet.size,
      unitsInvolvedList: Array.from(item.unitsSet),
      remark: Array.from(item.remarksSet).join("；")
    }));

    // Process XuKa Table (Grouping strictly by 加工图号、材质及型号、下料尺寸、L1、L2、备注)
    const xuKaMap: Record<string, {
      firstName: string;
      drawingNo: string;
      materialModel: string;
      cuttingSize: string;
      unit: string;
      l1: string;
      l2: string;
      remark: string;
      totalQty: number;
    }> = {};

    rawProfiles.forEach(p => {
      const drawingNo = (p.drawingNo || "").trim();
      const materialModel = (p.materialModel || "").trim();
      const cuttingSize = (p.cuttingSize || "").trim();
      const l1 = (p.l1 || "").trim();
      const l2 = (p.l2 || "").trim();
      const remark = (p.remark || "").trim();
      const name = (p.name || "").trim();
      const unit = (p.unit || "支").trim();

      // Group key strictly consists of the 6 required fields
      const groupKey = [drawingNo.toLowerCase(), materialModel.toLowerCase(), cuttingSize.toLowerCase(), l1.toLowerCase(), l2.toLowerCase(), remark.toLowerCase()].join("|||");

      if (!xuKaMap[groupKey]) {
        xuKaMap[groupKey] = {
          firstName: name,
          drawingNo,
          materialModel,
          cuttingSize,
          unit,
          l1,
          l2,
          remark,
          totalQty: 0
        };
      } else if (!xuKaMap[groupKey].firstName && name) {
        xuKaMap[groupKey].firstName = name;
      }

      xuKaMap[groupKey].totalQty += p.qty;
    });

    const summaryXuKaList: XuKaSummaryItem[] = Object.values(xuKaMap).map((item, idx) => ({
      seq: idx + 1,
      name: item.firstName || "型材",
      drawingNo: item.drawingNo,
      materialModel: item.materialModel,
      cuttingSize: item.cuttingSize,
      totalQty: Math.round(item.totalQty * 100) / 100,
      unit: item.unit,
      l1: item.l1,
      l2: item.l2,
      remark: item.remark
    }));

    // Process Auxiliary Summary Table
    const auxCategoryOrderMap: Record<string, number> = {
      "型材气密胶条类": 1,
      "不锈钢紧固件类": 2,
      "结构密封胶及辅材类": 3,
    };

    const rawAuxList = Object.values(auxMap);
    rawAuxList.sort((a, b) => {
      const orderA = auxCategoryOrderMap[a.category] ?? 99;
      const orderB = auxCategoryOrderMap[b.category] ?? 99;
      if (orderA !== orderB) {
        return orderA - orderB;
      }
      return a.name.localeCompare(b.name, "zh-CN");
    });

    const summaryAuxList: AuxSummaryItem[] = rawAuxList.map((item, idx) => ({
      id: `asum-${idx}`,
      seq: idx + 1,
      category: item.category,
      name: item.name,
      drawingNo: item.drawingNo,
      materialModel: item.materialModel,
      size: item.size,
      totalQty: Math.round(item.totalQty * 100) / 100,
      unit: item.unit,
      unitCount: item.unitsSet.size,
      unitsInvolvedList: Array.from(item.unitsSet),
      remark: Array.from(item.remarksSet).join("；")
    }));

    setProfileDataList(rawProfiles);
    setSteelPanelDataList(rawSteelPanels);
    setProfileSummaryList(summaryProfilesList);
    setXuKaSummaryList(summaryXuKaList);
    setAuxSummaryList(summaryAuxList);
    setScannedUnitsCount(activeSheets.length);
    setIsScanning(false);
    setLastCompletedTime(new Date().toLocaleTimeString());
    setShowCompletionModal(true);
  };

  // Run initial scan once when component mounts or sheets loaded
  useEffect(() => {
    if ((workbook || excelSheets.length > 0) && profileDataList.length === 0 && !isScanning) {
      runAggregation();
    }
  }, [workbook, excelSheets]);

  // Helper function to build styled worksheet with auto column widths, auto filter, and centered text
  const createStyledWorksheet = (rows: any[][], rowHeightPt?: number): XLSX.WorkSheet => {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    if (!rows || rows.length === 0) return ws;

    if (rowHeightPt) {
      ws["!rows"] = rows.map(() => ({ hpt: rowHeightPt }));
    }

    const colCount = rows[0].length;
    const colWidths: number[] = new Array(colCount).fill(0);

    // 1. Calculate auto column width for each column based on content
    rows.forEach((row) => {
      row.forEach((val, colIdx) => {
        if (colIdx >= colCount) return;
        const strVal = val !== null && val !== undefined ? String(val) : "";
        const lines = strVal.split("\n");
        let maxLineLen = 0;
        lines.forEach((line) => {
          let len = 0;
          for (let i = 0; i < line.length; i++) {
            const code = line.charCodeAt(i);
            // Chinese / full-width characters count as ~2.1 width, ASCII as ~1.15
            len += code > 255 ? 2.1 : 1.15;
          }
          if (len > maxLineLen) maxLineLen = len;
        });

        if (maxLineLen > colWidths[colIdx]) {
          colWidths[colIdx] = maxLineLen;
        }
      });
    });

    // Set column widths with extra padding for filter button (+5 units) and clamp min/max width
    ws["!cols"] = colWidths.map((w) => ({
      wch: Math.min(Math.max(Math.ceil(w) + 5, 12), 65),
    }));

    // 2. Apply cell styles (centering, fonts, borders, header background)
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1");
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const isHeader = R === 0;
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[cellRef]) {
          ws[cellRef] = { t: "s", v: "" };
        }

        ws[cellRef].s = {
          alignment: {
            horizontal: "center",
            vertical: "center",
            wrapText: true,
          },
          font: {
            name: "Microsoft YaHei",
            sz: isHeader ? 10.5 : 10,
            bold: isHeader,
            color: { rgb: isHeader ? "1E293B" : "0F172A" },
          },
          border: {
            top: { style: "thin", color: { rgb: "CBD5E1" } },
            bottom: { style: "thin", color: { rgb: "CBD5E1" } },
            left: { style: "thin", color: { rgb: "CBD5E1" } },
            right: { style: "thin", color: { rgb: "CBD5E1" } },
          },
          fill: isHeader
            ? {
                fgColor: { rgb: "F1F5F9" },
              }
            : undefined,
        };
      }
    }

    // 3. Set autofilter for header row
    if (ws["!ref"]) {
      ws["!autofilter"] = { ref: ws["!ref"] };
    }

    return ws;
  };

  // One-Click Export to 1 Excel file with 5 worksheets
  const handleExportAllToExcel = () => {
    if (profileDataList.length === 0 && steelPanelDataList.length === 0 && profileSummaryList.length === 0) {
      alert("无可导出的数据！请先点击『扫描并汇总所有工作表数据』按钮进行计算。");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      // Table 1 Sheet: 型材数据列表
      const sheet1Headers = [
        "单元编号", "序号", "名称", "加工图号", "材质及型号", "下料尺寸 (L) (mm)", "单樘用量", "总计数量", "单位", "L1", "L2", "备注"
      ];
      const sheet1Rows = [
        sheet1Headers,
        ...profileDataList.map((p) => [
          p.unitNo,
          p.rowSeq,
          p.name,
          p.drawingNo,
          p.materialModel,
          p.cuttingSize,
          p.unitUsage,
          p.qty,
          p.unit,
          p.l1,
          p.l2,
          p.remark
        ])
      ];
      const ws1 = createStyledWorksheet(sheet1Rows);
      XLSX.utils.book_append_sheet(wb, ws1, "型材数据列表");

      // Table 2 Sheet: 型材汇总表(套裁用)
      const sheet2Headers = [
        "序号", "名称", "材质及型号", "下料尺寸 (L) (mm)", "汇总总数量", "颜色"
      ];
      const sheet2Rows = [
        sheet2Headers,
        ...profileSummaryList.map(ps => [
          ps.seq,
          ps.name,
          ps.materialModel,
          ps.cuttingSize,
          ps.totalQty,
          ""
        ])
      ];
      const ws2 = createStyledWorksheet(sheet2Rows);

      // 在F2单元格填入用户设置的公式（型材汇总表 套裁用）
      if (f2Formula && f2Formula.trim()) {
        const formulaStr = f2Formula.trim();
        const cleanFormula = formulaStr.startsWith("=") ? formulaStr.substring(1) : formulaStr;
        const targetRef = "F2";
        if (!ws2[targetRef]) {
          ws2[targetRef] = { t: "s", v: "" };
        }
        ws2[targetRef].f = cleanFormula;

        // 确保 Range 涵盖 F2 (行 2，列 F) 并应用样式
        const range2 = XLSX.utils.decode_range(ws2["!ref"] || "A1");
        if (range2.e.r < 1) range2.e.r = 1;
        if (range2.e.c < 5) range2.e.c = 5;
        ws2["!ref"] = XLSX.utils.encode_range(range2);

        if (ws2[targetRef] && !ws2[targetRef].s) {
          ws2[targetRef].s = {
            alignment: { horizontal: "center", vertical: "center", wrapText: true },
            font: { name: "Microsoft YaHei", sz: 10, color: { rgb: "0F172A" } },
            border: {
              top: { style: "thin", color: { rgb: "CBD5E1" } },
              bottom: { style: "thin", color: { rgb: "CBD5E1" } },
              left: { style: "thin", color: { rgb: "CBD5E1" } },
              right: { style: "thin", color: { rgb: "CBD5E1" } },
            }
          };
        }
      }

      XLSX.utils.book_append_sheet(wb, ws2, "型材汇总表(套裁用)");

      // Table 3 Sheet: 序卡 (汇总依据：加工图号、材质及型号、下料尺寸、L1、L2、备注)
      const sheetXuKaHeaders = [
        "序号", "名称", "加工图号", "材质及型号", "下料尺寸 (L) (mm)", "汇总总数量", "单位", "L1", "L2", "备注"
      ];
      const sheetXuKaRows = [
        sheetXuKaHeaders,
        ...xuKaSummaryList.map(xk => [
          xk.seq,
          xk.name,
          xk.drawingNo,
          xk.materialModel,
          xk.cuttingSize,
          xk.totalQty,
          xk.unit,
          xk.l1,
          xk.l2,
          xk.remark
        ])
      ];
      const wsXuKa = createStyledWorksheet(sheetXuKaRows);
      XLSX.utils.book_append_sheet(wb, wsXuKa, "序卡");

      // Table 4 Sheet: 辅材汇总表
      const sheet3Headers = [
        "序号", "辅材类别", "名称", "加工图号", "材质及型号", "规格/尺寸", "汇总总数量", "单位", "备注"
      ];
      const sheet3Rows = [
        sheet3Headers,
        ...auxSummaryList.map(as => [
          as.seq,
          as.category,
          as.name,
          as.drawingNo,
          as.materialModel,
          as.size,
          as.totalQty,
          as.unit,
          as.remark
        ])
      ];
      const ws3 = createStyledWorksheet(sheet3Rows, 30);
      XLSX.utils.book_append_sheet(wb, ws3, "辅材汇总表");

      // Table 5 Sheet: 钢件与面板明细表
      const sheet5Headers = [
        "单元编号", "序号", "类别", "名称", "加工图号", "材质及型号", "下料尺寸 (L) (mm)", "单樘用量", "总计数量", "单位", "L1", "L2", "备注"
      ];
      const sheet5Rows = [
        sheet5Headers,
        ...steelPanelDataList.map((p) => [
          p.unitNo,
          p.rowSeq,
          p.categoryName,
          p.name,
          p.drawingNo,
          p.materialModel,
          p.cuttingSize,
          p.unitUsage,
          p.qty,
          p.unit,
          p.l1,
          p.l2,
          p.remark
        ])
      ];
      const ws5 = createStyledWorksheet(sheet5Rows);
      XLSX.utils.book_append_sheet(wb, ws5, "钢件与面板明细表");

      // Download file
      XLSX.writeFile(wb, "细目铝型材及辅材汇总表.xlsx");
      alert("🎉 导出成功！五张汇总表（型材数据列表、型材汇总表(套裁用)、序卡、辅材汇总表、钢件与面板明细表）已一键打包写入一个 Excel 文件并自动下载！");
    } catch (err: any) {
      alert(`导出失败: ${err.message}`);
    }
  };

  // Search filtering
  const filteredProfilesDetail = useMemo(() => {
    if (!searchQuery.trim()) return profileDataList;
    const q = searchQuery.toLowerCase();
    return profileDataList.filter(p => 
      p.unitNo.toLowerCase().includes(q) ||
      p.name.toLowerCase().includes(q) ||
      p.drawingNo.toLowerCase().includes(q) ||
      p.materialModel.toLowerCase().includes(q) ||
      p.cuttingSize.toLowerCase().includes(q) ||
      p.remark.toLowerCase().includes(q)
    );
  }, [profileDataList, searchQuery]);

  const filteredProfilesSummary = useMemo(() => {
    if (!searchQuery.trim()) return profileSummaryList;
    const q = searchQuery.toLowerCase();
    return profileSummaryList.filter(ps => 
      ps.name.toLowerCase().includes(q) ||
      ps.drawingNo.toLowerCase().includes(q) ||
      ps.materialModel.toLowerCase().includes(q) ||
      ps.cuttingSize.toLowerCase().includes(q) ||
      ps.remark.toLowerCase().includes(q)
    );
  }, [profileSummaryList, searchQuery]);

  const filteredXuKaSummary = useMemo(() => {
    if (!searchQuery.trim()) return xuKaSummaryList;
    const q = searchQuery.toLowerCase();
    return xuKaSummaryList.filter(xk => 
      xk.name.toLowerCase().includes(q) ||
      xk.drawingNo.toLowerCase().includes(q) ||
      xk.materialModel.toLowerCase().includes(q) ||
      xk.cuttingSize.toLowerCase().includes(q) ||
      xk.l1.toLowerCase().includes(q) ||
      xk.l2.toLowerCase().includes(q) ||
      xk.remark.toLowerCase().includes(q)
    );
  }, [xuKaSummaryList, searchQuery]);

  const filteredAuxSummary = useMemo(() => {
    if (!searchQuery.trim()) return auxSummaryList;
    const q = searchQuery.toLowerCase();
    return auxSummaryList.filter(as => 
      as.category.toLowerCase().includes(q) ||
      as.name.toLowerCase().includes(q) ||
      as.drawingNo.toLowerCase().includes(q) ||
      as.materialModel.toLowerCase().includes(q) ||
      as.size.toLowerCase().includes(q) ||
      as.remark.toLowerCase().includes(q)
    );
  }, [auxSummaryList, searchQuery]);

  const filteredSteelPanelsDetail = useMemo(() => {
    if (!searchQuery.trim()) return steelPanelDataList;
    const q = searchQuery.toLowerCase();
    return steelPanelDataList.filter(p => 
      p.unitNo.toLowerCase().includes(q) ||
      p.categoryName.toLowerCase().includes(q) ||
      p.name.toLowerCase().includes(q) ||
      p.drawingNo.toLowerCase().includes(q) ||
      p.materialModel.toLowerCase().includes(q) ||
      p.cuttingSize.toLowerCase().includes(q) ||
      p.remark.toLowerCase().includes(q)
    );
  }, [steelPanelDataList, searchQuery]);

  // Overall metric totals
  const totalProfilePieces = useMemo(() => {
    return profileDataList.reduce((acc, p) => acc + (p.qty || 0), 0);
  }, [profileDataList]);

  return (
    <div className="space-y-6" id="module_data_summary">
      {/* Completion Confirmation Banner */}
      {lastCompletedTime && (
        <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 text-xs text-emerald-900 shadow-2xs">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-emerald-100 border border-emerald-200 text-emerald-700 flex items-center justify-center shrink-0">
              <CheckCircle2 className="w-5 h-5" />
            </div>
            <div>
              <div className="font-extrabold text-emerald-800 flex items-center gap-2">
                <span>数据汇总表格生成完毕！</span>
                <span className="text-[10px] bg-emerald-100/80 text-emerald-800 px-2 py-0.5 rounded font-mono font-normal">
                  生成时间: {lastCompletedTime}
                </span>
              </div>
              <p className="text-[11px] text-emerald-700 mt-0.5">
                已成功扫描 <strong className="font-extrabold">{scannedUnitsCount}</strong> 个单元工作表，生成型材明细 <strong>{profileDataList.length}</strong> 行 ({totalProfilePieces}支)、型材套裁汇总 <strong>{profileSummaryList.length}</strong> 项、序卡 <strong>{xuKaSummaryList.length}</strong> 项、辅材汇总 <strong>{auxSummaryList.length}</strong> 项、钢件与面板明细 <strong>{steelPanelDataList.length}</strong> 行。
              </p>
            </div>
          </div>
          <button
            onClick={() => setShowCompletionModal(true)}
            className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-3.5 py-1.5 rounded-lg text-xs transition-colors shrink-0 cursor-pointer shadow-2xs flex items-center gap-1"
          >
            <FileCheck className="w-3.5 h-3.5" />
            <span>查看确认详情</span>
          </button>
        </div>
      )}

      {/* Overview Metric Banner Cards */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">已统计单元工作表</p>
            <h3 className="text-2xl font-black text-slate-800 mt-1">{scannedUnitsCount} <span className="text-xs font-normal text-slate-500">个单元</span></h3>
          </div>
          <div className="w-11 h-11 bg-indigo-50 border border-indigo-100 rounded-lg flex items-center justify-center text-indigo-600">
            <Layers className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">铝型材明细总条数</p>
            <h3 className="text-2xl font-black text-slate-800 mt-1">{profileDataList.length} <span className="text-xs font-normal text-slate-500">行 (共{totalProfilePieces}支)</span></h3>
          </div>
          <div className="w-11 h-11 bg-emerald-50 border border-emerald-100 rounded-lg flex items-center justify-center text-emerald-600">
            <Table className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">钢件与面板明细条数</p>
            <h3 className="text-2xl font-black text-slate-800 mt-1">{steelPanelDataList.length} <span className="text-xs font-normal text-slate-500">行</span></h3>
          </div>
          <div className="w-11 h-11 bg-orange-50 border border-orange-100 rounded-lg flex items-center justify-center text-orange-600">
            <Box className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">去重铝型材型号与下料组</p>
            <h3 className="text-2xl font-black text-slate-800 mt-1">{profileSummaryList.length} <span className="text-xs font-normal text-slate-500">类</span></h3>
          </div>
          <div className="w-11 h-11 bg-amber-50 border border-amber-100 rounded-lg flex items-center justify-center text-amber-600">
            <PieChart className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex items-center justify-between">
          <div>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-wider">去重辅材与胶条五金组</p>
            <h3 className="text-2xl font-black text-slate-800 mt-1">{auxSummaryList.length} <span className="text-xs font-normal text-slate-500">项</span></h3>
          </div>
          <div className="w-11 h-11 bg-purple-50 border border-purple-100 rounded-lg flex items-center justify-center text-purple-600">
            <Box className="w-5 h-5" />
          </div>
        </div>
      </div>

      {/* Control Action Bar */}
      <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-indigo-600 rounded-lg flex items-center justify-center text-white font-bold shadow-sm">
            <BarChart3 className="w-5 h-5" />
          </div>
          <div>
            <h2 className="text-base font-bold text-slate-800 flex items-center gap-2">
              第5模块：全区数据汇总与一键导出
              <span className="bg-indigo-50 text-indigo-700 text-[10px] font-bold px-2 py-0.5 rounded border border-indigo-150">
                支持工作簿下全部 Sheet 跨表统筹
              </span>
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              轮询统计 Excel 工作簿中所有单元图纸工作表（自动排除目录），将全区铝型材及辅材归集去重并一键导出 Excel。
            </p>
          </div>
        </div>

        <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
          <button
            onClick={runAggregation}
            disabled={isScanning || pullLoading}
            className="flex-1 md:flex-initial bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] text-white text-xs font-bold px-4 py-2.5 rounded-lg border border-indigo-700 transition-all flex items-center justify-center gap-2 cursor-pointer shadow-xs disabled:opacity-50 select-none"
          >
            <RefreshCw className={`w-4 h-4 ${isScanning ? "animate-spin" : ""}`} />
            <span>{isScanning ? `正在扫描汇总 (${scanProgress.current}/${scanProgress.total})...` : "扫描并汇总所有选中的工作表"}</span>
          </button>

          <button
            onClick={handleExportAllToExcel}
            disabled={isScanning || (profileDataList.length === 0 && auxSummaryList.length === 0)}
            className="flex-1 md:flex-initial bg-emerald-600 hover:bg-emerald-700 active:scale-[0.98] text-white text-xs font-extrabold px-5 py-2.5 rounded-lg transition-all flex items-center justify-center gap-2 cursor-pointer shadow-sm shadow-emerald-100 disabled:opacity-50 select-none"
          >
            <Download className="w-4 h-4" />
            <span>一键导出汇总 Excel (包含4张表)</span>
          </button>
        </div>
      </div>

      {/* Export Formula Configuration Card */}
      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-2xs space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-2 pb-1 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <Settings className="w-4 h-4 text-indigo-600" />
            <span className="text-xs font-bold text-slate-800">
              导出公式设置 — 型材汇总表(套裁用) F2 单元格公式
            </span>
            <span className="text-[11px] text-slate-400 font-normal">
              (数据导出时将自动写入 F2，已本地缓存)
            </span>
          </div>
          {f2Formula !== DEFAULT_F2_FORMULA && (
            <button
              onClick={() => handleFormulaChange(DEFAULT_F2_FORMULA)}
              className="text-[11px] text-indigo-600 hover:text-indigo-800 font-medium hover:underline cursor-pointer"
            >
              恢复默认公式
            </button>
          )}
        </div>
        <div className="flex items-center gap-2">
          <input
            type="text"
            value={f2Formula}
            onChange={(e) => handleFormulaChange(e.target.value)}
            placeholder="请输入 F2 单元格中使用的导出公式，如 =@INDEX(...)"
            className="flex-1 text-xs font-mono bg-slate-50 border border-slate-300 rounded-lg px-3 py-2 text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:bg-white transition-all"
          />
        </div>
        <p className="text-[11px] text-slate-500">
          💡 导出 Excel 时，系统将在『型材汇总表(套裁用)』工作表中的 <span className="font-mono font-bold text-slate-700">F2</span> 单元格自动填充此设置的公式。
        </p>
      </div>

      {/* Worksheet Scope Management Panel */}
      <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-2xs space-y-3">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-150">
          <div className="flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-indigo-600" />
            <span className="text-xs font-bold text-slate-800">
              工作簿工作表范围选择清单 (已选中 {Object.values(selectedSheetsState).filter(Boolean).length} / {allSheetNames.length} 个Sheet)
            </span>
            <span className="text-[11px] text-slate-500 font-normal">
              — 点击标签可勾选/取消对应工作表
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => toggleAllSheets(true)}
              className="text-[11px] font-bold text-slate-600 hover:text-indigo-600 bg-slate-100 hover:bg-indigo-50 px-2.5 py-1 rounded border border-slate-200 transition-colors cursor-pointer select-none"
            >
              勾选全部 ({allSheetNames.length})
            </button>
            <button
              onClick={toggleExcludeDirectorySheets}
              className="text-[11px] font-bold text-indigo-700 hover:text-indigo-800 bg-indigo-50 hover:bg-indigo-100 px-2.5 py-1 rounded border border-indigo-200 transition-colors cursor-pointer select-none"
            >
              重置 (仅选中单元细目表，排除目录)
            </button>
            <button
              onClick={() => toggleAllSheets(false)}
              className="text-[11px] font-bold text-slate-500 hover:text-rose-600 bg-slate-100 hover:bg-rose-50 px-2.5 py-1 rounded border border-slate-200 transition-colors cursor-pointer select-none"
            >
              取消全部
            </button>
          </div>
        </div>

        {/* Scan Progress Bar */}
        {isScanning && (
          <div className="p-3 bg-indigo-50/80 rounded-lg border border-indigo-100 space-y-2">
            <div className="flex justify-between items-center text-xs text-indigo-900 font-bold">
              <span className="flex items-center gap-2">
                <RefreshCw className="w-3.5 h-3.5 animate-spin text-indigo-600" />
                正在跨工作表检索数据: {scanProgress.currentSheet || "准备中..."}
              </span>
              <span className="font-mono">{scanProgress.current} / {scanProgress.total}</span>
            </div>
            <div className="w-full bg-indigo-200/60 rounded-full h-2 overflow-hidden">
              <div 
                className="bg-indigo-600 h-2 rounded-full transition-all duration-300"
                style={{ width: `${Math.round((scanProgress.current / (scanProgress.total || 1)) * 100)}%` }}
              />
            </div>
          </div>
        )}

        {/* Sheet badges list */}
        <div className="flex flex-wrap gap-2 pt-1">
          {allSheetNames.length === 0 ? (
            <div className="text-xs text-slate-400 italic py-2">
              未检索到 Excel 工作表列表。请先通过第 1 模块选择或连接 Excel 文件。
            </div>
          ) : (
            allSheetNames.map(sheet => {
              const isSelected = !!selectedSheetsState[sheet];
              const isDirectory = sheet.includes("目录") || sheet.toLowerCase().includes("index");

              return (
                <button
                  key={sheet}
                  onClick={() => toggleSingleSheet(sheet)}
                  className={`text-xs font-mono font-medium px-3 py-1.5 rounded-lg border transition-all flex items-center gap-2 cursor-pointer select-none ${
                    isSelected
                      ? "bg-indigo-50 text-indigo-900 border-indigo-300 font-bold shadow-2xs"
                      : "bg-slate-50 text-slate-400 border-slate-200 hover:bg-slate-100 opacity-60 line-through"
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={isSelected}
                    onChange={() => {}} // handled by button onClick
                    className="w-3.5 h-3.5 rounded text-indigo-600 focus:ring-0 cursor-pointer pointer-events-none"
                  />
                  <span>{sheet}</span>
                  {isDirectory && (
                    <span className="text-[10px] bg-amber-100 text-amber-800 font-sans px-1.5 py-0.2 rounded font-normal no-underline">
                      目录索引
                    </span>
                  )}
                </button>
              );
            })
          )}
        </div>
      </div>

      {/* Main Table Sub-Tabs & Filtering Section */}
      <div id="summary_tables_tabs" className="bg-white border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
        {/* Table Nav Tabs & Search bar */}
        <div className="p-4 border-b border-slate-200 bg-slate-50/80 flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
          <div className="flex flex-wrap items-center gap-2">
            {[
              { id: 'profiles_detail', name: '表一：型材数据列表 (按单元顺序)', count: profileDataList.length },
              { id: 'profiles_summary', name: '表二：型材汇总表(套裁用)', count: profileSummaryList.length },
              { id: 'xuka_summary', name: '表三：序卡 (按加工图号+材质+尺寸+L1/L2+备注汇总)', count: xuKaSummaryList.length },
              { id: 'aux_summary', name: '表四：辅材汇总表 (去重复汇总)', count: auxSummaryList.length },
              { id: 'steel_panel_detail', name: '表五：钢件与面板明细表 (按单元顺序)', count: steelPanelDataList.length },
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveSubTab(tab.id as any)}
                className={`px-3.5 py-2 rounded-lg text-xs font-bold transition-all border flex items-center gap-2 cursor-pointer select-none ${
                  activeSubTab === tab.id
                    ? "bg-indigo-600 text-white border-indigo-700 shadow-xs"
                    : "bg-white text-slate-600 hover:bg-slate-100 border-slate-200"
                }`}
              >
                <span>{tab.name}</span>
                <span className={`px-1.5 py-0.2 rounded-full text-[10px] font-mono ${
                  activeSubTab === tab.id ? "bg-indigo-700 text-indigo-100" : "bg-slate-100 text-slate-600"
                }`}>
                  {tab.count}
                </span>
              </button>
            ))}
          </div>

          <div className="relative w-full lg:w-72">
            <Search className="w-3.5 h-3.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="搜索名称、图号、材质或单元号..."
              className="w-full bg-white border border-slate-200 rounded-lg pl-8 pr-3 py-1.5 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-indigo-500"
            />
          </div>
        </div>

        {/* Sub-Tab 1: 型材数据列表 (Profile Data List) */}
        {activeSubTab === 'profiles_detail' && (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 border-collapse">
              <thead className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-3 px-4 w-36">单元编号</th>
                  <th className="py-3 px-3 w-16 text-center">序号</th>
                  <th className="py-3 px-4 min-w-[140px]">名称</th>
                  <th className="py-3 px-4 min-w-[120px]">加工图号</th>
                  <th className="py-3 px-4 min-w-[150px]">材质及型号</th>
                  <th className="py-3 px-4 text-right min-w-[130px]">下料尺寸 (L) (mm)</th>
                  <th className="py-3 px-3 text-center w-24">单樘用量</th>
                  <th className="py-3 px-4 text-right font-black text-indigo-900 w-28">总计数量</th>
                  <th className="py-3 px-3 text-center w-16">单位</th>
                  <th className="py-3 px-3 text-right w-20">L1</th>
                  <th className="py-3 px-3 text-right w-20">L2</th>
                  <th className="py-3 px-4 min-w-[150px]">备注</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredProfilesDetail.length === 0 ? (
                  <tr>
                    <td colSpan={12} className="py-12 text-center text-slate-400 italic">
                      {isScanning ? "正在扫描工作表数据中..." : "暂无型材数据。请确保连通数据源并点击上方『扫描并汇总所有工作表数据』。"}
                    </td>
                  </tr>
                ) : (
                  filteredProfilesDetail.map((p, idx) => (
                    <tr key={idx} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-4 font-bold text-indigo-700 font-mono bg-indigo-50/30 border-r border-slate-100">
                        {p.unitNo}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-400 font-mono">{p.rowSeq}</td>
                      <td className="py-2.5 px-4 font-bold text-slate-900">{p.name}</td>
                      <td className="py-2.5 px-4 font-mono text-slate-700 bg-slate-50/50">{p.drawingNo || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-700">{p.materialModel || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-800">{p.cuttingSize || "-"}</td>
                      <td className="py-2.5 px-3 text-center font-mono text-slate-600">{p.unitUsage || "1"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-black text-indigo-700 text-sm bg-indigo-50/20">
                        {p.qty}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-500">{p.unit || "支"}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-500">{p.l1 || "-"}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-500">{p.l2 || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-500 truncate max-w-xs">{p.remark || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Sub-Tab 2: 型材汇总表 (Profile Summary Table) */}
        {activeSubTab === 'profiles_summary' && (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 border-collapse">
              <thead className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-3 px-3 w-16 text-center">序号</th>
                  <th className="py-3 px-4 min-w-[150px]">名称</th>
                  <th className="py-3 px-4 min-w-[160px] bg-amber-50/50 text-amber-900">材质及型号</th>
                  <th className="py-3 px-4 text-right min-w-[150px] bg-amber-50/50 text-amber-900">下料尺寸 (L) (mm)</th>
                  <th className="py-3 px-4 text-right font-black text-emerald-800 w-32 bg-emerald-50/50">汇总总数量</th>
                  <th className="py-3 px-4 min-w-[120px]">颜色</th>
                  <th className="py-3 px-4 min-w-[150px]">汇总备注</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredProfilesSummary.length === 0 ? (
                  <tr>
                    <td colSpan={7} className="py-12 text-center text-slate-400 italic">
                      {isScanning ? "正在计算汇总数据中..." : "暂无型材汇总数据。"}
                    </td>
                  </tr>
                ) : (
                  filteredProfilesSummary.map((ps) => (
                    <tr key={ps.id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400">{ps.seq}</td>
                      <td className="py-2.5 px-4 font-bold text-slate-900">{ps.name}</td>
                      <td className="py-2.5 px-4 font-semibold text-slate-800 bg-amber-50/20">{ps.materialModel || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-900 bg-amber-50/20">{ps.cuttingSize || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-black text-emerald-700 text-sm bg-emerald-50/30">
                        {ps.totalQty}
                      </td>
                      <td className="py-2.5 px-4 text-slate-500 font-mono text-xs truncate max-w-[180px]" title={ps.seq === 1 ? f2Formula : ""}>
                        {ps.seq === 1 ? (f2Formula || "-") : "-"}
                      </td>
                      <td className="py-2.5 px-4 text-slate-500 truncate max-w-xs">{ps.remark || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Sub-Tab 3: 序卡 (XuKa Summary Table) */}
        {activeSubTab === 'xuka_summary' && (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 border-collapse">
              <thead className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-3 px-3 w-16 text-center">序号</th>
                  <th className="py-3 px-4 min-w-[140px]">名称</th>
                  <th className="py-3 px-4 min-w-[130px] bg-blue-50/50 text-blue-900">加工图号</th>
                  <th className="py-3 px-4 min-w-[150px] bg-blue-50/50 text-blue-900">材质及型号</th>
                  <th className="py-3 px-4 text-right min-w-[140px] bg-blue-50/50 text-blue-900">下料尺寸 (L) (mm)</th>
                  <th className="py-3 px-4 text-right font-black text-blue-900 w-28 bg-blue-50/70">汇总总数量</th>
                  <th className="py-3 px-3 text-center w-16">单位</th>
                  <th className="py-3 px-3 text-right w-20">L1</th>
                  <th className="py-3 px-3 text-right w-20">L2</th>
                  <th className="py-3 px-4 min-w-[150px]">备注</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredXuKaSummary.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="py-12 text-center text-slate-400 italic">
                      {isScanning ? "正在计算序卡数据中..." : "暂序卡汇总数据。"}
                    </td>
                  </tr>
                ) : (
                  filteredXuKaSummary.map((xk) => (
                    <tr key={xk.seq} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400">{xk.seq}</td>
                      <td className="py-2.5 px-4 font-bold text-slate-900">{xk.name}</td>
                      <td className="py-2.5 px-4 font-mono font-semibold text-blue-900 bg-blue-50/20">{xk.drawingNo || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-800 bg-blue-50/20">{xk.materialModel || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-900 bg-blue-50/20">{xk.cuttingSize || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-black text-blue-700 text-sm bg-blue-50/40">
                        {xk.totalQty}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-500">{xk.unit || "支"}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-600">{xk.l1 || "-"}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-600">{xk.l2 || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-500 truncate max-w-xs">{xk.remark || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Sub-Tab 4: 辅材汇总表 (Auxiliary Summary Table) */}
        {activeSubTab === 'aux_summary' && (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 border-collapse">
              <thead className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-3 px-3 w-16 text-center">序号</th>
                  <th className="py-3 px-4 min-w-[140px]">辅材大类</th>
                  <th className="py-3 px-4 min-w-[160px]">名称</th>
                  <th className="py-3 px-4 min-w-[130px]">加工图号</th>
                  <th className="py-3 px-4 min-w-[150px]">材质及型号</th>
                  <th className="py-3 px-4 min-w-[140px]">规格 / 尺寸</th>
                  <th className="py-3 px-4 text-right font-black text-purple-900 w-32 bg-purple-50/50">汇总总数量</th>
                  <th className="py-3 px-3 text-center w-16">单位</th>
                  <th className="py-3 px-4 min-w-[150px]">备注</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredAuxSummary.length === 0 ? (
                  <tr>
                    <td colSpan={9} className="py-12 text-center text-slate-400 italic">
                      {isScanning ? "正在计算辅材汇总数据中..." : "暂无辅材汇总数据。"}
                    </td>
                  </tr>
                ) : (
                  filteredAuxSummary.map((as) => (
                    <tr key={as.id} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-3 text-center font-mono text-slate-400">{as.seq}</td>
                      <td className="py-2.5 px-4 font-bold text-purple-800">
                        <span className="bg-purple-50 border border-purple-150 text-purple-700 px-2 py-0.5 rounded text-[11px]">
                          {as.category}
                        </span>
                      </td>
                      <td className="py-2.5 px-4 font-bold text-slate-900">{as.name}</td>
                      <td className="py-2.5 px-4 font-mono text-slate-700">{as.drawingNo || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-700">{as.materialModel || "-"}</td>
                      <td className="py-2.5 px-4 font-mono text-slate-800">{as.size || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-black text-purple-700 text-sm bg-purple-50/20">
                        {as.totalQty}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-500 font-medium">{as.unit || "项"}</td>
                      <td className="py-2.5 px-4 text-slate-500 truncate max-w-xs">{as.remark || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Sub-Tab 5: 钢件与面板明细表 (Steel & Panel Detail List) */}
        {activeSubTab === 'steel_panel_detail' && (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 border-collapse">
              <thead className="bg-slate-100/90 border-b border-slate-200 text-slate-600 font-bold uppercase tracking-wider">
                <tr>
                  <th className="py-3 px-4 w-36">单元编号</th>
                  <th className="py-3 px-3 w-16 text-center">序号</th>
                  <th className="py-3 px-4 text-center w-20">类别</th>
                  <th className="py-3 px-4 min-w-[140px]">名称</th>
                  <th className="py-3 px-4 min-w-[120px]">加工图号</th>
                  <th className="py-3 px-4 min-w-[150px]">材质及型号</th>
                  <th className="py-3 px-4 text-right min-w-[130px]">下料尺寸 (L) (mm)</th>
                  <th className="py-3 px-3 text-center w-24">单樘用量</th>
                  <th className="py-3 px-4 text-right font-black text-indigo-900 w-28">总计数量</th>
                  <th className="py-3 px-3 text-center w-16">单位</th>
                  <th className="py-3 px-3 text-right w-20">L1</th>
                  <th className="py-3 px-3 text-right w-20">L2</th>
                  <th className="py-3 px-4 min-w-[150px]">备注</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredSteelPanelsDetail.length === 0 ? (
                  <tr>
                    <td colSpan={13} className="py-12 text-center text-slate-400 italic">
                      {isScanning ? "正在扫描工作表数据中..." : "暂无钢件与面板明细数据。请确保连通数据源并点击上方『扫描并汇总所有工作表数据』。"}
                    </td>
                  </tr>
                ) : (
                  filteredSteelPanelsDetail.map((p, idx) => (
                    <tr key={idx} className="hover:bg-slate-50/80 transition-colors">
                      <td className="py-2.5 px-4 font-bold text-indigo-700 font-mono bg-indigo-50/30 border-r border-slate-100">
                        {p.unitNo}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-400 font-mono">{p.rowSeq}</td>
                      <td className="py-2.5 px-4 text-center font-bold">
                        <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${
                          p.categoryName === '钢件'
                            ? "bg-amber-100/80 text-amber-800 border border-amber-200"
                            : "bg-purple-100/80 text-purple-800 border border-purple-200"
                        }`}>
                          {p.categoryName}
                        </span>
                      </td>
                      <td className="py-2.5 px-4 font-bold text-slate-900">{p.name}</td>
                      <td className="py-2.5 px-4 font-mono text-slate-700 bg-slate-50/50">{p.drawingNo || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-700">{p.materialModel || "-"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-bold text-slate-800">{p.cuttingSize || "-"}</td>
                      <td className="py-2.5 px-3 text-center font-mono text-slate-600">{p.unitUsage || "1"}</td>
                      <td className="py-2.5 px-4 text-right font-mono font-black text-indigo-700 text-sm bg-indigo-50/20">
                        {p.qty}
                      </td>
                      <td className="py-2.5 px-3 text-center text-slate-500">{p.unit}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-500">{p.l1 || "-"}</td>
                      <td className="py-2.5 px-3 text-right font-mono text-slate-500">{p.l2 || "-"}</td>
                      <td className="py-2.5 px-4 text-slate-500 truncate max-w-xs">{p.remark || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Generation Complete Confirmation Modal */}
      {showCompletionModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-200">
          <div className="bg-white rounded-2xl shadow-2xl border border-slate-100 max-w-lg w-full p-6 space-y-5 relative">
            <button 
              onClick={() => setShowCompletionModal(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-100 transition-colors cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>

            <div className="flex items-center gap-3.5">
              <div className="w-12 h-12 rounded-2xl bg-emerald-100 border border-emerald-200 text-emerald-600 flex items-center justify-center shrink-0 shadow-xs">
                <CheckCircle2 className="w-7 h-7" />
              </div>
              <div>
                <h3 className="text-base font-black text-slate-800">数据汇总表格生成完毕！</h3>
                <p className="text-xs text-slate-500 mt-0.5">全区单元图纸工作表数据已顺利解析、归集与去重</p>
              </div>
            </div>

            <div className="bg-slate-50 rounded-xl p-4 border border-slate-200/80 space-y-3">
              <div className="text-xs font-bold text-slate-700 flex items-center justify-between pb-2 border-b border-slate-200">
                <span className="flex items-center gap-1.5">
                  <FileCheck className="w-4 h-4 text-emerald-600" />
                  <span>汇总统计汇总表清单</span>
                </span>
                <span className="text-emerald-700 bg-emerald-100/80 text-[11px] px-2 py-0.5 rounded font-mono font-bold">
                  生成成功 ({lastCompletedTime})
                </span>
              </div>
              
              <div className="grid grid-cols-2 gap-2.5 text-xs">
                <div className="bg-white p-2.5 rounded-lg border border-slate-150">
                  <span className="text-slate-400 block text-[11px]">统计单元工作表</span>
                  <span className="text-slate-800 font-extrabold text-sm">{scannedUnitsCount} <span className="text-xs font-normal text-slate-500">个</span></span>
                </div>
                <div className="bg-white p-2.5 rounded-lg border border-slate-150">
                  <span className="text-slate-400 block text-[11px]">型材数据明细列表</span>
                  <span className="text-slate-800 font-extrabold text-sm">{profileDataList.length} <span className="text-xs font-normal text-slate-500">行 ({totalProfilePieces}支)</span></span>
                </div>
                <div className="bg-white p-2.5 rounded-lg border border-slate-150">
                  <span className="text-slate-400 block text-[11px]">型材汇总表 (套裁用)</span>
                  <span className="text-slate-800 font-extrabold text-sm">{profileSummaryList.length} <span className="text-xs font-normal text-slate-500">类</span></span>
                </div>
                <div className="bg-white p-2.5 rounded-lg border border-slate-150">
                  <span className="text-slate-400 block text-[11px]">序卡 (套裁汇总)</span>
                  <span className="text-slate-800 font-extrabold text-sm">{xuKaSummaryList.length} <span className="text-xs font-normal text-slate-500">项</span></span>
                </div>
                <div className="bg-white p-2.5 rounded-lg border border-slate-150 col-span-2">
                  <span className="text-slate-400 block text-[11px]">辅材汇总表 (胶条/紧固件/密封胶)</span>
                  <span className="text-slate-800 font-extrabold text-sm">{auxSummaryList.length} <span className="text-xs font-normal text-slate-500">项</span></span>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3 pt-1">
              <button
                type="button"
                onClick={() => {
                  setShowCompletionModal(false);
                  const el = document.getElementById("summary_tables_tabs");
                  if (el) el.scrollIntoView({ behavior: "smooth" });
                }}
                className="flex-1 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold py-2.5 px-4 rounded-xl transition-colors cursor-pointer text-center"
              >
                查看汇总表格
              </button>
              <button
                type="button"
                onClick={() => {
                  setShowCompletionModal(false);
                  handleExportAllToExcel();
                }}
                className="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold py-2.5 px-4 rounded-xl shadow-xs transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>一键导出 Excel</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
