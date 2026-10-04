import React, { useState } from "react";
import { AuxiliaryLedgerItem, AuxiliaryLedgerItemRule, ProcessedItem } from "../types";
import { Layers, Plus, Trash, Clipboard, FileText, CheckCircle2, AlertCircle, HelpCircle, Upload, Download, Edit3, X, Sparkles, Sliders, Search, Compass, Check, Copy, Ruler, History, RotateCcw, ChevronDown, ChevronUp } from "lucide-react";
import { matchSearchQuery, evaluateFormula } from "../utils/calcEngine";

// Globals to cache the hot-connection file handle even if the component is re-rendered or remounted
let cachedFileHandle: any = null;
let cachedFileName: string = "";

interface ExpressionSegment {
  type: 'text' | 'token';
  text: string;
  keyword?: string;
  category?: string;
  searchField?: string;
  targetField?: string;
  isExactMatch?: boolean;
  startIndex: number;
  endIndex: number;
}

const parseExpression = (expr: string): ExpressionSegment[] => {
  if (!expr) return [];
  const segments: ExpressionSegment[] = [];
  const rowValRegex = /ROW_VAL\((["'])(.*?)\1,\s*(["'])(.*?)\3,\s*(["'])(.*?)\5,\s*(["'])(.*?)\7(?:,\s*(["'])(true|false)\9)?\)/gi;
  let lastIndex = 0;
  let match;

  while ((match = rowValRegex.exec(expr)) !== null) {
    const matchIndex = match.index;
    const fullCall = match[0];

    if (matchIndex > lastIndex) {
      segments.push({
        type: 'text',
        text: expr.slice(lastIndex, matchIndex),
        startIndex: lastIndex,
        endIndex: matchIndex
      });
    }

    segments.push({
      type: 'token',
      text: fullCall,
      category: match[2],
      searchField: match[4],
      keyword: match[6],
      targetField: match[8] || 'cuttingSize',
      isExactMatch: match[10] === 'true',
      startIndex: matchIndex,
      endIndex: matchIndex + fullCall.length
    });

    lastIndex = rowValRegex.lastIndex;
  }

  if (lastIndex < expr.length) {
    segments.push({
      type: 'text',
      text: expr.slice(lastIndex),
      startIndex: lastIndex,
      endIndex: expr.length
    });
  }

  return segments;
};

function parseClipboardData(text: string): string[][] {
  let delimiter = '\t';
  if (text.indexOf('\t') === -1) {
    if (text.indexOf(';') !== -1) {
      delimiter = ';';
    } else if (text.indexOf('，') !== -1) {
      delimiter = '，';
    } else if (text.indexOf(',') !== -1) {
      delimiter = ',';
    }
  }

  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentCell = "";
  let inQuotes = false;
  let i = 0;

  while (i < text.length) {
    const char = text[i];
    
    if (inQuotes) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          currentCell += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        currentCell += char;
        i++;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
        i++;
      } else if (char === delimiter) {
        currentRow.push(currentCell);
        currentCell = "";
        i++;
      } else if (char === '\r') {
        if (text[i + 1] === '\n') {
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = "";
          i += 2;
        } else {
          currentRow.push(currentCell);
          rows.push(currentRow);
          currentRow = [];
          currentCell = "";
          i++;
        }
      } else if (char === '\n') {
        currentRow.push(currentCell);
        rows.push(currentRow);
        currentRow = [];
        currentCell = "";
        i++;
      } else {
        currentCell += char;
        i++;
      }
    }
  }

  if (currentRow.length > 0 || currentCell !== "") {
    currentRow.push(currentCell);
    rows.push(currentRow);
  }

  return rows.filter(row => {
    return row.length > 1 || (row.length === 1 && row[0].trim() !== "");
  });
}

interface AuxMaterialLedgerProps {
  ledger: AuxiliaryLedgerItem[];
  items: ProcessedItem[];
  projectVars?: Record<string, any>;
  rawRows?: string[][];
  onAddLedgerItem: (item: AuxiliaryLedgerItem) => void;
  onBulkAddLedgerItems: (items: AuxiliaryLedgerItem[]) => void;
  onRemoveLedgerItem: (id: string) => void;
  onClearLedger: () => void;
  onUpdateLedgerItem: (item: AuxiliaryLedgerItem) => void;
  onOverwriteLedgerItems?: (items: AuxiliaryLedgerItem[]) => void;
  ledgerHistory?: Array<{
    id: string;
    timestamp: string;
    description: string;
    itemCount: number;
    data: AuxiliaryLedgerItem[];
  }>;
  onRestoreHistory?: (historyId: string) => void;
}

export default function AuxMaterialLedger({
  ledger,
  items,
  projectVars = {},
  rawRows,
  onAddLedgerItem,
  onBulkAddLedgerItems,
  onRemoveLedgerItem,
  onClearLedger,
  onUpdateLedgerItem,
  onOverwriteLedgerItems,
  ledgerHistory = [],
  onRestoreHistory,
}: AuxMaterialLedgerProps) {
  // Calculate temporary unit usage results for ledger table display (not saved to file)
  const calculatedMap = React.useMemo(() => {
    const map = new Map<string, { unitUsage: number; isCalculated: boolean; rulesCount: number }>();

    ledger.forEach(item => {
      const itemRules = item.rules || [];
      const activeRules = itemRules.filter(r => r.isActive);
      if (activeRules.length === 0) {
        map.set(item.id, { unitUsage: 0, isCalculated: false, rulesCount: 0 });
        return;
      }

      let unitUsage = 0;
      const evaluatedRules = activeRules.map(rule => {
        const { value } = evaluateFormula(rule.expression, items, projectVars, rule.lengthAdjustment || 0, rawRows);
        return { rule, value };
      });

      const matchedExclusive = evaluatedRules.find(er => er.rule.isExclusive && er.value !== 0);

      evaluatedRules.forEach(({ rule, value }) => {
        const isIgnored = !!matchedExclusive && matchedExclusive.rule.id !== rule.id;
        if (!isIgnored) {
          unitUsage += value;
        }
      });

      const COUNT_UNITS = ['个', '支', '块', '只', '套'];
      const cleanUnit = item.unit ? item.unit.trim() : '';
      const isCountUnit = COUNT_UNITS.includes(cleanUnit);
      if (isCountUnit) {
        unitUsage = Math.ceil(unitUsage);
      } else if (item.category === 'fastener') {
        unitUsage = Math.round(unitUsage);
      }

      map.set(item.id, { unitUsage, isCalculated: true, rulesCount: activeRules.length });
    });

    return map;
  }, [ledger, items, projectVars, rawRows]);

  const [activeInputTab, setActiveInputTab] = useState<'manual' | 'paste' | null>(null);
  const [showImportDropdown, setShowImportDropdown] = useState(false);
  const [importMode, setImportMode] = useState<'append' | 'overwrite'>('append');
  const dropdownRef = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setShowImportDropdown(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);
  
  // Token inline editing states
  const [editingToken, setEditingToken] = useState<{
    ruleId: string;
    segmentIndex: number;
    category: string;
    searchField: string;
    keyword: string;
    targetField: string;
    isExactMatch?: boolean;
    originalExpression: string;
  } | null>(null);

  const handleTokenClick = (ruleId: string, startIndex: number, endIndex: number) => {
    const listTextarea = document.getElementById(`rule-textarea-list-${ruleId}`) as HTMLTextAreaElement;
    const editTextarea = document.getElementById(`rule-textarea-edit-${ruleId}`) as HTMLTextAreaElement;

    if (listTextarea) {
      listTextarea.focus();
      listTextarea.setSelectionRange(startIndex, endIndex);
    }
    if (editTextarea) {
      editTextarea.focus();
      editTextarea.setSelectionRange(startIndex, endIndex);
    }
  };

  const renderFormulaVisualizer = (ruleId: string, expression: string) => {
    const segments = parseExpression(expression);
    
    if (segments.length === 0) {
      return (
        <span className="text-slate-400 italic text-[11px] select-none">
          公式为空（请在右侧点击“插入参数”或直接在下方编写数学公式）
        </span>
      );
    }

    return (
      <div className="flex flex-wrap items-center gap-1.5 p-2 px-2.5 bg-slate-100 border border-slate-200 rounded-lg min-h-[38px] w-full text-xs font-semibold font-mono text-slate-800 break-all select-none shadow-inner">
        {segments.map((segment, idx) => {
          if (segment.type === 'text') {
            return (
              <span key={idx} className="text-slate-600 font-bold whitespace-pre-wrap select-text">
                {segment.text}
              </span>
            );
          } else {
            return (
              <span
                key={idx}
                onClick={(e) => {
                  e.stopPropagation();
                  handleTokenClick(ruleId, segment.startIndex || 0, segment.endIndex || 0);
                }}
                onDoubleClick={(e) => {
                  e.stopPropagation();
                  setEditingToken({
                    ruleId,
                    segmentIndex: idx,
                    category: segment.category || 'profile',
                    searchField: segment.searchField || 'name',
                    keyword: segment.keyword || '',
                    targetField: segment.targetField || 'cuttingSize',
                    isExactMatch: segment.isExactMatch,
                    originalExpression: expression
                  });
                }}
                className="inline-flex items-center gap-0.5 px-2 py-0.5 rounded border border-cyan-300 bg-cyan-50/80 hover:bg-cyan-100 hover:text-cyan-950 transition-colors cursor-pointer font-bold select-none text-[11px] text-cyan-800 shadow-xs"
                title="单击定位并选中对应的公式段，双击自定义修改该公式的提取规则"
              >
                &lt;{segment.isExactMatch ? '【精确】' : ''}{segment.keyword}{segment.targetField === 'cuttingSize' ? '长度' : segment.targetField === 'cuttingWidth' ? '宽' : segment.targetField === 'cuttingHeight' ? '高' : segment.targetField === 'unitUsage' ? '数量' : ''}&gt;
              </span>
            );
          }
        })}
      </div>
    );
  };

  // File Handle state for hot connection
  const [fileHandle, setFileHandle] = useState<any>(cachedFileHandle);
  const [connectedFileName, setConnectedFileName] = useState<string>(cachedFileName);

  // Sync back to local file system dynamically whenever ledger list is updated
  React.useEffect(() => {
    async function syncToFile() {
      if (!fileHandle) return;
      try {
        // Query permissions first to handle secure sandboxed origins
        if (fileHandle.queryPermission) {
          const permission = await fileHandle.queryPermission({ mode: "readwrite" });
          if (permission !== "granted") {
            const request = await fileHandle.requestPermission({ mode: "readwrite" });
            if (request !== "granted") {
              console.warn("Permission to write file was denied by the user/browser configuration.");
              return;
            }
          }
        }
        const writable = await fileHandle.createWritable();
        await writable.write(JSON.stringify(ledger, null, 2));
        await writable.close();
        console.log("🟢 [台账热连接] 算料结构/特征条目更新，已动态保存回写至:", fileHandle.name);
      } catch (err: any) {
        console.error("🔴 [台账热连接] 动态保存至本地文件失败:", err);
      }
    }
    syncToFile();
  }, [ledger, fileHandle]);
  
  // Manual form state
  const [category, setCategory] = useState<'gasket' | 'fastener' | 'auxiliary'>('gasket');
  const [name, setName] = useState('');
  const [drawingNo, setDrawingNo] = useState('');
  const [materialModel, setMaterialModel] = useState('');
  const [size, setSize] = useState('');
  const [unit, setUnit] = useState('米');
  const [remark, setRemark] = useState('');
  const [position, setPosition] = useState('');

  // Editing state for the merged rules dialog window
  const [editingItem, setEditingItem] = useState<AuxiliaryLedgerItem | null>(null);
  const [isRecycleBinOpen, setIsRecycleBinOpen] = useState(false);

  // States for the popup lengthening dialog
  const [lengthenModalOpen, setLengthenModalOpen] = useState(false);
  const [lengthenRuleId, setLengthenRuleId] = useState<string | null>(null);
  const [lengthenValue, setLengthenValue] = useState<string>("");

  // States for the popup parameter insertion dialog
  const [paramModalOpen, setParamModalOpen] = useState(false);
  const [targetRuleId, setTargetRuleId] = useState<string | null>(null);
  const [paramCategory, setParamCategory] = useState<string>('profile');
  const [searchHeader, setSearchHeader] = useState<string>('name');
  const [searchKeyword, setSearchKeyword] = useState('');
  const [isSearchExactMatch, setIsSearchExactMatch] = useState(false);
  const [selectedRowId, setSelectedRowId] = useState<number | null>(null);
  const [isAllAttributesExpanded, setIsAllAttributesExpanded] = useState(false);

  const handleStartEdit = (item: AuxiliaryLedgerItem) => {
    setEditingItem({
      ...item,
      rules: item.rules ? item.rules.map(r => ({ ...r })) : []
    });
  };

  const handleUpdateEditingField = (field: keyof AuxiliaryLedgerItem, value: any) => {
    if (!editingItem) return;
    setEditingItem(prev => prev ? { ...prev, [field]: value } : null);
  };

  const handleUpdateRuleField = (ruleId: string, ruleField: keyof AuxiliaryLedgerItemRule, value: any) => {
    if (!editingItem) return;
    setEditingItem(prev => {
      if (!prev) return null;
      const updatedRules = (prev.rules || []).map(r => 
        r.id === ruleId ? { ...r, [ruleField]: value } : r
      );
      return { ...prev, rules: updatedRules };
    });
  };

  const handleAddRuleToEditing = () => {
    if (!editingItem) return;
    const isGasket = editingItem.category === "gasket";
    const newRule: AuxiliaryLedgerItemRule = {
      id: `rule-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      name: `算料规则_${(editingItem.rules || []).length + 1}`,
      expression: "",
      isActive: true,
      description: "",
      lengthAdjustment: isGasket ? 50 : undefined
    };
    setEditingItem(prev => {
      if (!prev) return null;
      return {
        ...prev,
        rules: [...(prev.rules || []), newRule]
      };
    });
  };

  const handleDeleteRule = (ruleId: string) => {
    if (!editingItem) return;
    setEditingItem(prev => {
      if (!prev) return null;
      return {
        ...prev,
        rules: (prev.rules || []).filter(r => r.id !== ruleId)
      };
    });
  };

  const handleDuplicateRule = (rule: AuxiliaryLedgerItemRule) => {
    if (!editingItem) return;
    const copiedRule: AuxiliaryLedgerItemRule = {
      ...rule,
      id: `rule-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      name: `${rule.name}_副本`,
    };
    setEditingItem(prev => {
      if (!prev) return null;
      return {
        ...prev,
        rules: [...(prev.rules || []), copiedRule]
      };
    });
  };

  const handleSaveEdit = () => {
    if (!editingItem) return;
    onUpdateLedgerItem(editingItem);
    setEditingItem(null);
  };

  const handleSelectCellData = (fieldKey: string, rawVal: string, forceNumeric: boolean, correspondingUnitUsage?: string) => {
    let insertToken = "";
    
    if (forceNumeric) {
      const trimmed = rawVal.trim().replace(/,/g, '');
      const match = trimmed.match(/[-+]?[0-9]*\.?[0-9]+/);
      if (match) {
        let numericVal = Number(match[0]);
        // Convert profile and steel/fastener millimeters standard sizes to meters (divide by 1000)
        if ((paramCategory === "profile" || paramCategory === "steel" || paramCategory === "fastener") && ["cuttingSize", "l1", "l2"].includes(fieldKey)) {
          numericVal = numericVal / 1000;
        }
        
        // If the parameter is cuttingSize, multiply by corresponding unitUsage (单樘用量)
        if (fieldKey === "cuttingSize") {
          const usageStr = correspondingUnitUsage || "1";
          const usageMatch = usageStr.trim().replace(/,/g, '').match(/[-+]?[0-9]*\.?[0-9]+/);
          const usageVal = usageMatch ? Number(usageMatch[0]) : 1;
          insertToken = `(${numericVal} * ${usageVal})`;
        } else {
          insertToken = String(numericVal);
        }
      } else {
        insertToken = "0";
      }
    } else {
      if (fieldKey === "cuttingSize") {
        const usageStr = correspondingUnitUsage || "1";
        const usageMatch = usageStr.trim().replace(/,/g, '').match(/[-+]?[0-9]*\.?[0-9]+/);
        const usageVal = usageMatch ? Number(usageMatch[0]) : 1;
        insertToken = `'${rawVal.replace(/'/g, "\\'")}' * ${usageVal}`;
      } else {
        insertToken = `'${rawVal.replace(/'/g, "\\'")}'`;
      }
    }

    if (!editingItem || !targetRuleId) return;

    setEditingItem(prev => {
      if (!prev || !prev.rules) return prev;
      const updatedRules = prev.rules.map(r => {
        if (r.id === targetRuleId) {
          const currentExp = r.expression.trim();
          let separator = "";
          if (currentExp !== "") {
            const lastChar = currentExp[currentExp.length - 1];
            if (!["+", "-", "*", "/", "(", ")"].includes(lastChar)) {
              separator = " + ";
            } else {
              separator = " ";
            }
          }
          return {
            ...r,
            expression: currentExp + separator + insertToken
          };
        }
        return r;
      });
      return { ...prev, rules: updatedRules };
    });
  };

  const handleSelectDynamicRowVal = (fieldKey: string, matchedRowName: string) => {
    const category = paramCategory;
    const searchField = searchHeader;
    const keyword = searchKeyword.trim() || matchedRowName.trim();
    const targetField = fieldKey;

    let insertToken = "";
    if (["cuttingSize", "cuttingWidth", "cuttingHeight"].includes(fieldKey)) {
      insertToken = `(ROW_VAL('${category}', '${searchField}', '${keyword}', '${targetField}', '${isSearchExactMatch}') * ROW_VAL('${category}', '${searchField}', '${keyword}', 'unitUsage', '${isSearchExactMatch}'))`;
    } else {
      insertToken = `ROW_VAL('${category}', '${searchField}', '${keyword}', '${targetField}', '${isSearchExactMatch}')`;
    }

    if (!editingItem || !targetRuleId) return;

    setEditingItem(prev => {
      if (!prev || !prev.rules) return prev;
      const updatedRules = prev.rules.map(r => {
        if (r.id === targetRuleId) {
          const currentExp = r.expression.trim();
          let separator = "";
          if (currentExp !== "") {
            const lastChar = currentExp[currentExp.length - 1];
            if (!["+", "-", "*", "/", "(", ")"].includes(lastChar)) {
              separator = " + ";
            } else {
              separator = " ";
            }
          }
          return {
            ...r,
            expression: currentExp + separator + insertToken
          };
        }
        return r;
      });
      return { ...prev, rules: updatedRules };
    });
  };

  const handleInsertFormulaText = (text: string) => {
    if (!editingItem || !targetRuleId) return;
    setEditingItem(prev => {
      if (!prev || !prev.rules) return prev;
      const updatedRules = prev.rules.map(r => {
        if (r.id === targetRuleId) {
          const currentExp = r.expression.trim();
          let separator = "";
          if (currentExp !== "") {
            const lastChar = currentExp[currentExp.length - 1];
            if (!["+", "-", "*", "/", "(", ")"].includes(lastChar)) {
              separator = " + ";
            } else {
              separator = " ";
            }
          }
          return {
            ...r,
            expression: currentExp + separator + text
          };
        }
        return r;
      });
      return { ...prev, rules: updatedRules };
    });
  };

  // Bulk paste state
  const [pasteCategory, setPasteCategory] = useState<'gasket' | 'fastener' | 'auxiliary'>('gasket');
  const [pasteText, setPasteText] = useState('');
  const [pasteFeedback, setPasteFeedback] = useState<string | null>(null);

  const handleManualAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if (!name.trim()) return;

    const newItem: AuxiliaryLedgerItem = {
      id: `ledger-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
      category,
      name: name.trim(),
      drawingNo: drawingNo.trim(),
      materialModel: materialModel.trim(),
      size: size.trim(),
      unit: unit.trim(),
      remark: remark.trim(),
      position: position.trim(),
    };

    onAddLedgerItem(newItem);
    
    // Clear major inputs
    setName('');
    setDrawingNo('');
    setMaterialModel('');
    setSize('');
    setRemark('');
    setPosition('');
  };

  const handlePasteImport = () => {
    if (!pasteText.trim()) {
      setPasteFeedback("请在文本域内粘贴从 Excel 中复制的物料行数据。");
      return;
    }

    const rows = parseClipboardData(pasteText);
    const parsedItems: AuxiliaryLedgerItem[] = [];
    let errorCount = 0;

    rows.forEach((cells) => {
      // Filter out empty rows or headers
      const joined = cells.join('');
      if (!joined.trim()) return;
      if (joined.includes("名称") || joined.includes("加工图号") || joined.includes("单樘用量") || joined.includes("算料逻辑归属")) {
        return;
      }

      if (cells.length > 0 && cells[0].trim()) {
        const cLen = cells.length;
        // Detect if the first cell is purely numeric and the row is long enough, indicating a '序号' (Index) column is present
        const hasIndexColumn = /^\d+$/.test(cells[0].trim()) && cLen >= 4;
        const offset = hasIndexColumn ? 1 : 0;

        const pName = cells[offset + 0]?.trim() || "未命名辅料";
        const pDrawingNo = cells[offset + 1]?.trim() || "—";
        const pMaterialModel = cells[offset + 2]?.trim() || "—";
        const pSize = cells[offset + 3]?.trim() || "—";
        
        // Use the standard 10+ column layout from user's Excel template
        // ["(序号)", "名称", "加工图号", "材质及型号", "下料尺寸", "单樘用量", "总计", "单位", "L1", "L2", "备注", "算料逻辑归属"]
        const pUnit = cells[offset + 6]?.trim() || (pasteCategory === "gasket" ? "米" : "只");
        const pRemark = cells[offset + 9]?.trim() || "";
        const pPosition = ""; // Force clear
        
        let determinedCategory = pasteCategory;

        // Auto-detect category if the 11th column is provided
        if (cells[offset + 10]) {
          const catStr = cells[offset + 10].trim();
          if (catStr.includes("胶条")) {
            determinedCategory = "gasket";
          } else if (catStr.includes("紧固") || catStr.includes("螺栓") || catStr.includes("五金")) {
            determinedCategory = "fastener";
          } else if (catStr.includes("密") || catStr.includes("胶") || catStr.includes("垫") || catStr.includes("辅材")) {
            determinedCategory = "auxiliary";
          }
        }

        const item: AuxiliaryLedgerItem = {
          id: `ledger-${Date.now()}-${Math.floor(Math.random() * 10000)}`,
          category: determinedCategory,
          name: pName,
          drawingNo: pDrawingNo,
          materialModel: pMaterialModel,
          size: pSize,
          unit: pUnit,
          remark: pRemark,
          position: pPosition, // Cleared/empty as requested!
        };
        parsedItems.push(item);
      } else {
        errorCount++;
      }
    });

    if (parsedItems.length > 0) {
      if (importMode === 'overwrite') {
        // Keep other categories' items, overwrite only items of the current active/imported categories
        const categoriesToOverwrite = new Set<string>([pasteCategory]);
        parsedItems.forEach(item => categoriesToOverwrite.add(item.category));
        const preservedItems = ledger.filter(item => !categoriesToOverwrite.has(item.category));
        const finalLedger = [...preservedItems, ...parsedItems];

        if (onOverwriteLedgerItems) {
          onOverwriteLedgerItems(finalLedger);
        } else {
          onClearLedger();
          onBulkAddLedgerItems(finalLedger);
        }
        setPasteFeedback(`成功覆盖导入当前品类的 ${parsedItems.length} 行物料（其他品类物料未受影响），且已自动将“单樘用量”与“总计”列清空！`);
      } else {
        onBulkAddLedgerItems(parsedItems);
        setPasteFeedback(`成功追加导入 ${parsedItems.length} 行物料，且已自动将“单樘用量”与“总计”列清空！`);
      }
      setPasteText('');
    } else {
      setPasteFeedback("未能成功解析出有效的物料，请确认格式是否正确。");
    }
  };

  const loadDemoItems = () => {
    const demos: AuxiliaryLedgerItem[] = [
      {
        id: "demo-1",
        category: "gasket",
        name: "中空玻璃外侧气密胶条",
        drawingNo: "H2-JT19",
        materialModel: "三元乙丙橡胶 EPDM",
        size: "15mm宽 槽深6",
        unit: "米",
        remark: "室内室外槽压嵌配套使用",
        position: "玻璃周长槽位"
      },
      {
        id: "demo-2",
        category: "fastener",
        name: "横梁固定不锈钢自攻螺栓",
        drawingNo: "ST6.3*32",
        materialModel: "304不锈钢 盘头",
        size: "M6.3x32-35",
        unit: "只",
        remark: "每个立柱横梁结合处配置4颗螺钉",
        position: "龙骨结合角码孔"
      },
      {
        id: "demo-3",
        category: "auxiliary",
        name: "结构缝耐候聚氨酯密封胶",
        drawingNo: "W-SH35",
        materialModel: "单组份防霉型密封胶",
        size: "300ml 支装",
        unit: "支",
        remark: "根据外围玻璃结合缝总周长自动套数",
        position: "玻璃拼缝"
      }
    ];
    onBulkAddLedgerItems(demos);
  };

  const handleImportWithFileSystemAccess = async () => {
    if (typeof (window as any).showOpenFilePicker === "function") {
      try {
        const [handle] = await (window as any).showOpenFilePicker({
          types: [{
            description: "配比辅料 JSON 配置文件",
            accept: {
              "application/json": [".json"]
            }
          }],
          multiple: false
        });
        const file = await handle.getFile();
        const text = await file.text();
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) {
          const validated = parsed.filter(item => item && typeof item === "object" && "name" in item) as AuxiliaryLedgerItem[];
          if (validated.length > 0) {
            const mapped = validated.map(item => ({
              ...item,
              id: item.id || `ledger-imported-${Date.now()}-${Math.floor(Math.random() * 10000)}`
            }));
            
            // Assign cache handles before adding bulk items to state
            cachedFileHandle = handle;
            cachedFileName = handle.name;
            setFileHandle(handle);
            setConnectedFileName(handle.name);
            
            onBulkAddLedgerItems(mapped);
            setPasteFeedback(`✨ 已成功开启热连接！已导入并关联本地文件 [${handle.name}]，台账所有修改/增删将实时流式动态保存回写至该文件。`);
          } else {
            setPasteFeedback("导入失败：JSON 文件中未包含任何合规的辅材台账项目形式。");
          }
        } else {
          setPasteFeedback("导入失败：JSON 数据的最外层结构必须为数组格式。");
        }
      } catch (err: any) {
        if (err.name !== "AbortError") {
          setPasteFeedback(`文件读取关联失败: ${err.message || err}`);
        }
      }
    } else {
      // Fallback to native upload
      document.getElementById("btn-native-file-upload")?.click();
    }
  };

  const handleImportData = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileReader = new FileReader();
    const files = e.target.files;
    if (!files || files.length === 0) return;

    fileReader.onload = event => {
      try {
        const parsed = JSON.parse(event.target?.result as string);
        if (Array.isArray(parsed)) {
          const validated = parsed.filter(item => item && typeof item === 'object' && 'name' in item) as AuxiliaryLedgerItem[];
          if (validated.length > 0) {
            const mapped = validated.map(item => ({
              ...item,
              id: item.id || `ledger-imported-${Date.now()}-${Math.floor(Math.random() * 10000)}`
            }));
            onBulkAddLedgerItems(mapped);
            setPasteFeedback(`成功导入 ${mapped.length} 项辅材料数据！(提示：单次导入未建立热连接，如需修改自动同步，推荐点击 [热连接导入台账])`);
          } else {
            setPasteFeedback("导入失败：JSON 文件中未包含有效的辅材台账记录。");
          }
        } else {
          setPasteFeedback("导入失败：JSON 数据的根节点必须是个数组。");
        }
      } catch (error) {
        setPasteFeedback("解析 JSON 失败：请确保文件格式符合 standard JSON 规范。");
      }
    };
    fileReader.readAsText(files[0]);
    e.target.value = "";
  };

  const handleExportData = () => {
    try {
      if (ledger.length === 0) {
        setPasteFeedback("当前台账没有任何数据可以导出，请先新增或载入常规数据。");
        return;
      }
      const dataStr = JSON.stringify(ledger, null, 2);
      const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(dataStr);

      const exportFileDefaultName = `auxiliary_material_ledger_${new Date().toISOString().slice(0, 10)}.json`;

      const linkElement = document.createElement('a');
      linkElement.setAttribute('href', dataUri);
      linkElement.setAttribute('download', exportFileDefaultName);
      linkElement.click();

      setPasteFeedback(`台账数据导出成功！已保存为 ${exportFileDefaultName}`);
    } catch (err) {
      setPasteFeedback("配置数据导出失败！");
    }
  };

  const filteredItems = React.useMemo(() => {
    if (!items || items.length === 0) return [];
    const results = items.filter(item => {
      // 1. Filter by category
      if (item.category !== paramCategory) return false;
      
      // 2. Filter by search column keyword (logical and wildcard matches)
      if (!searchKeyword.trim()) return true;
      
      let valToMatch = "";
      if (searchHeader === "name") valToMatch = item.name || "";
      else if (searchHeader === "drawingNo") valToMatch = item.drawingNo || "";
      else if (searchHeader === "materialModel") valToMatch = item.materialModel || "";
      else if (searchHeader === "cuttingSize") valToMatch = item.cuttingSize || "";
      else if (searchHeader === "unitUsage") valToMatch = item.unitUsage || "";
      else if (searchHeader === "l1") valToMatch = item.l1 || "";
      else if (searchHeader === "l2") valToMatch = item.l2 || "";
      else if (searchHeader === "remark") valToMatch = item.remark || "";
      
      return matchSearchQuery(valToMatch, searchKeyword, isSearchExactMatch);
    });

    if (results.length === 0 && searchKeyword.trim()) {
      const kw = searchKeyword.trim();
      const dummyItem: ProcessedItem = {
        rowIndex: -1,
        originalCategory: paramCategory,
        category: paramCategory as any,
        name: searchHeader === "name" ? kw : "未匹配(空)",
        drawingNo: searchHeader === "drawingNo" ? kw : "空",
        materialModel: searchHeader === "materialModel" ? kw : "空",
        cuttingSize: `<${kw}长度>`,
        unitUsage: `<${kw}数量>`,
        qty: 0,
        unit: "",
        l1: searchHeader === "l1" ? kw : "",
        l2: searchHeader === "l2" ? kw : "",
        remark: searchHeader === "remark" ? kw : "",
        spec: "",
        material: "",
        len: 0
      };
      results.push(dummyItem);
    }

    return results;
  }, [items, paramCategory, searchHeader, searchKeyword, isSearchExactMatch]);

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="module_aux_ledger">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Layers className="w-5 h-5 text-indigo-600" />
            配比辅料及紧固件台账定义
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            定义需要运算定额的辅材（胶条、螺母螺栓、海绵垫及结构耐候密封胶）及加工规格图号，后续配比引擎可秒级推算其用量。
          </p>
          {connectedFileName ? (
            <div className="mt-2.5 inline-flex items-center gap-2 px-3 py-1.5 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs shadow-xs animate-pulse select-none">
              <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
              <span className="font-bold">🟢 热连接已建立: 正在实时流式同步至 [{connectedFileName}]</span>
              <button 
                type="button"
                onClick={() => {
                  cachedFileHandle = null;
                  cachedFileName = "";
                  setFileHandle(null);
                  setConnectedFileName("");
                  setPasteFeedback("热连接已主动断开，台账修改仅缓存在浏览器中。");
                }}
                className="ml-2 text-[10px] text-emerald-600 hover:text-emerald-950 hover:bg-emerald-100/50 px-1.5 py-0.5 rounded border border-emerald-200 font-bold transition-all cursor-pointer"
              >
                断开热连接
              </button>
            </div>
          ) : (
            <div className="mt-2 text-[10px] text-slate-400 flex items-center gap-1">
              <span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>
              <span>暂无活动热连接。推荐点击 [导入台账数据] 选择 [建立本地文件热连接] 以开启实时流式动态同步。</span>
            </div>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {ledger.length === 0 && (
            <button
              onClick={loadDemoItems}
              className="text-xs bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-2 rounded transition-colors flex items-center gap-1.5 shadow-xs cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>采用示例辅材台账</span>
            </button>
          )}
          
          <div className="relative" ref={dropdownRef}>
            <button
              type="button"
              onClick={() => setShowImportDropdown(prev => !prev)}
              className="text-xs bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-2 rounded transition-colors flex items-center gap-1.5 cursor-pointer shadow-xs shadow-indigo-100"
              title="载入台账配置文件并选择是否建立实时流式热保存连接"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>导入台账数据</span>
              <span className="text-[9px] opacity-80 select-none">▼</span>
            </button>
            {showImportDropdown && (
              <div className="absolute left-0 mt-1 w-64 bg-white border border-slate-200 rounded-lg shadow-lg z-50 text-slate-700 overflow-hidden font-medium text-xs">
                <button
                  type="button"
                  onClick={async () => {
                    setShowImportDropdown(false);
                    await handleImportWithFileSystemAccess();
                  }}
                  className="w-full text-left px-4 py-2.5 hover:bg-indigo-50 hover:text-indigo-700 flex flex-col gap-0.5 border-b border-slate-100 transition-colors cursor-pointer"
                >
                  <span className="font-bold flex items-center gap-1 text-slate-900">
                    <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                    建立本地文件热连接
                  </span>
                  <span className="text-[10px] text-slate-400 font-normal">
                    选择 JSON 配置文件，所有对台账的增删改均会实时回写原文件。
                  </span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setShowImportDropdown(false);
                    document.getElementById("btn-native-file-upload")?.click();
                  }}
                  className="w-full text-left px-4 py-2.5 hover:bg-indigo-50 hover:text-indigo-700 flex flex-col gap-0.5 transition-colors cursor-pointer"
                >
                  <span className="font-bold flex items-center gap-1 text-slate-900">
                    <span className="w-1.5 h-1.5 rounded-full bg-slate-400"></span>
                    仅单次常规导入
                  </span>
                  <span className="text-[10px] text-slate-400 font-normal">
                    仅导入台账数据到浏览器，后续修改不覆盖原文件。
                  </span>
                </button>
              </div>
            )}
          </div>
          <input
            id="btn-native-file-upload"
            type="file"
            accept=".json"
            onChange={handleImportData}
            className="hidden"
          />

          <button
            onClick={handleExportData}
            title="导出当前台账全量设定数据为 JSON 文件"
            className="text-xs bg-slate-50 hover:bg-slate-100 text-slate-700 border border-slate-300 font-bold px-3 py-2 rounded transition-colors flex items-center gap-1.5 cursor-pointer"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>导出台账数据</span>
          </button>

          <button
            onClick={onClearLedger}
            className="text-xs bg-rose-50 hover:bg-rose-100 text-rose-700 font-bold px-3 py-2 rounded transition-colors border border-rose-200 cursor-pointer"
          >
            清空台账
          </button>
        </div>
      </div>

      {/* Input panel tabs */}
      <div className="flex flex-col gap-6 ">
        {/* Establish registry section */}
        <div className="border border-slate-200 rounded-xl overflow-hidden bg-white shadow-xs max-w-4xl">
          <div className="flex bg-slate-50 border-b border-slate-200">
            <button
              id="btn_tab_manual"
              type="button"
              onClick={() => { setActiveInputTab(prev => prev === 'manual' ? null : 'manual'); setPasteFeedback(null); }}
              className={`flex-1 py-3.5 text-xs font-bold transition-all text-center border-r border-slate-200 cursor-pointer flex items-center justify-center gap-2 hover:bg-slate-100/60 ${
                activeInputTab === 'manual' ? "bg-white text-indigo-700 border-t-2 border-t-indigo-600 font-bold" : "text-slate-500 hover:text-slate-800"
              }`}
            >
              <span>手动逐行向台账录入物料</span>
              {activeInputTab === 'manual' ? (
                <span className="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-150 px-2 py-0.5 rounded-full font-bold flex items-center gap-0.5">
                  已展开 <span className="text-[8px]">▲</span>
                </span>
              ) : (
                <span className="text-[10px] bg-slate-100 text-slate-505 border border-slate-200 px-2 py-0.5 rounded-full font-semibold flex items-center gap-0.5">
                  点击展开 <span className="text-[8px]">▼</span>
                </span>
              )}
            </button>
            <button
              id="btn_tab_paste"
              type="button"
              onClick={() => { setActiveInputTab(prev => prev === 'paste' ? null : 'paste'); setPasteFeedback(null); }}
              className={`flex-1 py-3.5 text-xs font-bold transition-all text-center cursor-pointer flex items-center justify-center gap-2 hover:bg-slate-100/60 ${
                activeInputTab === 'paste' ? "bg-white text-indigo-700 border-t-2 border-t-indigo-600 font-bold" : "text-slate-500 hover:text-slate-800"
              }`}
            >
              <span>从 Excel 复制一键粘贴导入</span>
              {activeInputTab === 'paste' ? (
                <span className="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-150 px-2 py-0.5 rounded-full font-bold flex items-center gap-0.5">
                  已展开 <span className="text-[8px]">▲</span>
                </span>
              ) : (
                <span className="text-[10px] bg-slate-100 text-slate-505 border border-slate-200 px-2 py-0.5 rounded-full font-semibold flex items-center gap-0.5">
                  点击展开 <span className="text-[8px]">▼</span>
                </span>
              )}
            </button>
          </div>

          {activeInputTab !== null && (
            <div className="p-5 bg-white border-t border-slate-100 transition-all duration-300">
              {activeInputTab === 'manual' ? (
                <form onSubmit={handleManualAdd} className="space-y-3.5 text-xs">
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">辅料归并品类</label>
                      <select
                        value={category}
                        onChange={(e) => setCategory(e.target.value as any)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2 cursor-pointer font-semibold text-slate-800"
                      >
                        <option value="gasket">型材气密胶条类</option>
                        <option value="fastener">不锈钢紧固五金螺栓类</option>
                        <option value="auxiliary">结构密封胶及辅助垫块类</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">辅料名称 <span className="text-rose-500">*</span></label>
                      <input
                        type="text"
                        required
                        value={name}
                        onChange={(e) => setName(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2 font-medium"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">工艺加工图号</label>
                      <input
                        type="text"
                        value={drawingNo}
                        onChange={(e) => setDrawingNo(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2 font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">材质牌号及型号</label>
                      <input
                        type="text"
                        value={materialModel}
                        onChange={(e) => setMaterialModel(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-3 gap-3">
                    <div className="col-span-2">
                      <label className="block text-slate-500 font-bold mb-1">下料尺寸 (L) (MM)</label>
                      <input
                        type="text"
                        value={size}
                        onChange={(e) => setSize(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">单位</label>
                      <input
                        type="text"
                        value={unit}
                        onChange={(e) => setUnit(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 gap-3">
                    <div>
                      <label className="block text-slate-500 font-bold mb-1">备注说明 (备注)</label>
                      <input
                        type="text"
                        value={remark}
                        onChange={(e) => setRemark(e.target.value)}
                        className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2"
                      />
                    </div>
                  </div>

                  <button
                    type="submit"
                    className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 px-3 rounded select-none cursor-pointer transition-colors flex items-center justify-center gap-1.5 shadow-xs shadow-indigo-100"
                  >
                    <Plus className="w-4 h-4" />
                    <span>添加至辅料列表清单</span>
                  </button>
                </form>
              ) : (
                <div className="space-y-4 text-xs">
                  <div className="bg-amber-50 border border-amber-200 p-3 rounded text-slate-700 mb-2 leading-relaxed flex gap-2">
                    <HelpCircle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-bold block">剪贴板格式指南 (Excel 字段对齐):</span>
                      <span className="text-[10px] text-slate-500 block">
                        直接从 Excel 中选中多行并复制，内容序列需为：
                      </span>
                      <code className="text-[9px] bg-white border border-slate-200 px-1.5 py-0.5 rounded font-mono text-rose-600 inline-block mt-1 font-bold">
                        [名称] [加工图号] [材质型号] [外形尺寸] [单位] [备注说明] [安装位置]
                      </code>
                    </div>
                  </div>

                  <div>
                    <label className="block text-slate-500 font-bold mb-1">粘贴导入物料品类</label>
                    <select
                      value={pasteCategory}
                      onChange={(e) => setPasteCategory(e.target.value as any)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2 cursor-pointer font-semibold text-slate-800"
                    >
                      <option value="gasket">型材气密胶条类</option>
                      <option value="fastener">不锈钢紧固五金螺栓类</option>
                      <option value="auxiliary">结构密封胶及辅助垫块类</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-500 font-bold mb-1">在此粘贴 Excel 数据块 (支持 Tab/分号 分割)</label>
                    <textarea
                      rows={6}
                      value={pasteText}
                      onChange={(e) => setPasteText(e.target.value)}
                      placeholder="例如: 
 EP01双向密封条	H2-JT04	三元乙丙EPDM	12*10mm	米	打胶槽密封	玻璃腔体槽
 横架十字自攻螺栓	ST4.8*13	304不锈钢铆钉	M4.8x13	只	龙骨固定螺纹件	压料板交汇"
                      className="w-full bg-slate-50/50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-2.5 font-mono text-[10px] leading-relaxed"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-505 font-bold mb-1.5">数据导入模式</label>
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        type="button"
                        onClick={() => setImportMode('append')}
                        className={`flex items-center justify-center gap-1.5 py-2 px-3 rounded-lg border text-xs font-bold transition-all cursor-pointer ${
                          importMode === 'append'
                            ? 'bg-indigo-50 border-indigo-300 text-indigo-700 shadow-xs'
                            : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        <span className={`w-2 h-2 rounded-full ${importMode === 'append' ? 'bg-indigo-600 animate-pulse' : 'bg-slate-400'}`}></span>
                        追加导入 (保留现有)
                      </button>
                      <button
                        type="button"
                        onClick={() => setImportMode('overwrite')}
                        className={`flex items-center justify-center gap-1.5 py-2 px-3 rounded-lg border text-xs font-bold transition-all cursor-pointer ${
                          importMode === 'overwrite'
                            ? 'bg-amber-50 border-amber-300 text-amber-700 shadow-xs'
                            : 'bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100'
                        }`}
                      >
                        <span className={`w-2 h-2 rounded-full ${importMode === 'overwrite' ? 'bg-amber-600 animate-pulse' : 'bg-slate-400'}`}></span>
                        覆盖导入 (清空现有)
                      </button>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handlePasteImport}
                    className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-3 rounded select-none cursor-pointer transition-colors flex items-center justify-center gap-1.5 shadow-xs shadow-indigo-100"
                  >
                    <Clipboard className="w-4 h-4" />
                    <span>执行剪贴板解析导入</span>
                  </button>
                </div>
              )}

              {pasteFeedback && (
                <div className={`mt-4 p-3 rounded border text-[11px] leading-relaxed font-bold flex gap-1.5 items-start ${
                  pasteFeedback.includes("成功") 
                    ? "bg-emerald-50 border-emerald-200 text-emerald-800" 
                    : "bg-red-50 border-red-200 text-red-800"
                }`}>
                  {pasteFeedback.includes("成功") ? (
                    <CheckCircle2 className="w-4.5 h-4.5 text-emerald-600 shrink-0" />
                  ) : (
                    <AlertCircle className="w-4.5 h-4.5 text-red-500 shrink-0" />
                  )}
                  <span>{pasteFeedback}</span>
                </div>
              )}
            </div>
          )}
        </div>

        {/* 变动回收站 (Change Recycle Bin) */}
        <div className="border border-slate-200 rounded-xl overflow-hidden bg-white shadow-xs max-w-4xl p-5" id="ledger_recycle_bin">
          <div 
            className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 cursor-pointer select-none"
            onClick={() => setIsRecycleBinOpen(!isRecycleBinOpen)}
          >
            <div className="flex items-center gap-2">
              <History className="w-4 h-4 text-amber-500" />
              <h3 className="text-sm font-bold text-slate-800 flex items-center gap-1.5">
                辅材台账变动回收站
                <span className="text-[10px] bg-amber-50 text-amber-700 px-2 py-0.5 rounded-full border border-amber-200 font-bold">
                  保留最近 {ledgerHistory.length} / 5 次历史修改
                </span>
              </h3>
              {isRecycleBinOpen ? (
                <ChevronUp className="w-4 h-4 text-slate-400" />
              ) : (
                <ChevronDown className="w-4 h-4 text-slate-400" />
              )}
            </div>
            <div className="flex items-center gap-2">
              {ledgerHistory.length > 0 && (
                <span className="text-[10px] text-slate-400 font-medium">
                  {isRecycleBinOpen ? "点击折叠历史记录" : "点击展开查看历史备份"}
                </span>
              )}
              <button
                type="button"
                className="px-2.5 py-1 bg-slate-50 hover:bg-slate-100 text-slate-700 rounded text-xs font-bold transition-all border border-slate-200 shrink-0"
              >
                {isRecycleBinOpen ? "收起" : "展开"}
              </button>
            </div>
          </div>
          
          {isRecycleBinOpen && (
            <div className="mt-4 pt-4 border-t border-slate-100">
              {ledgerHistory.length === 0 ? (
                <div className="bg-slate-50/50 border border-dashed border-slate-200 rounded-lg p-4 text-center text-[11px] text-slate-400">
                  暂无历史变动备份。当您对辅材台账执行添加、编辑、删除、清空或导入操作时，系统会自动记录前一版本的快照。
                </div>
              ) : (
                <div className="space-y-2">
                  {ledgerHistory.map((history, idx) => (
                    <div 
                      key={history.id} 
                      className="flex flex-col sm:flex-row sm:items-center justify-between p-3 rounded-lg border border-slate-150 bg-slate-50/40 hover:bg-slate-50 transition-colors gap-2"
                    >
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-[10px] bg-indigo-50 border border-indigo-100 text-indigo-700 px-2 py-0.5 rounded font-mono font-bold">
                          版本 {idx + 1} ({history.timestamp})
                        </span>
                        <span className="text-xs font-bold text-slate-700">
                          {history.description}
                        </span>
                        <span className="text-[10px] text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded font-medium">
                          快照包含: {history.itemCount} 项物料
                        </span>
                      </div>
                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          if (onRestoreHistory) {
                            onRestoreHistory(history.id);
                            alert(`已成功还原历史修改版本 (${history.timestamp})！`);
                          }
                        }}
                        className="flex items-center justify-center gap-1.5 px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 active:scale-[0.98] text-[11px] font-bold rounded border border-indigo-200 cursor-pointer transition-all select-none whitespace-nowrap self-end sm:self-auto"
                        title="将当前辅材料台账的数据内容重置还原为该历史版本"
                      >
                        <RotateCcw className="w-3.5 h-3.5" />
                        <span>恢复此版本</span>
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>

        {/* Catalog Table Registry section */}
        <div className="w-full flex flex-col">
          <div className="border border-slate-200 rounded-xl overflow-hidden shadow-xs flex-1 bg-white">
            <div className="p-4 bg-slate-50/80 border-b border-slate-200 flex justify-between items-center">
              <span className="font-bold text-xs text-slate-700 uppercase tracking-widest flex items-center gap-1">
                <FileText className="w-4 h-4 text-indigo-600" />
                当前已建立的辅材料台账清单 ({ledger.length} 个登记品项)
              </span>
            </div>

            <div className="overflow-x-auto overflow-y-auto max-h-[500px]">
              <table className="min-w-full divide-y divide-slate-100 text-[11px]">
                <thead className="bg-slate-50/50 text-slate-550 font-bold uppercase tracking-wider text-[10px]">
                  <tr>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">名称</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">加工图号</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">材质及型号</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">下料尺寸 (L) (MM)</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">单樘用量</th>
                    <th className="px-3 py-2.5 text-right border-b border-slate-200">总计</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">单位</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">L1</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">L2</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">备注</th>
                    <th className="px-3 py-2.5 text-left border-b border-slate-200">算料逻辑归属</th>
                    <th className="px-3 py-2.5 text-center w-12 border-b border-slate-200">操作</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-slate-700">
                  {ledger.length === 0 ? (
                    <tr>
                      <td colSpan={12} className="px-4 py-12 text-center text-slate-400">
                        <div className="flex flex-col items-center justify-center gap-3">
                          <p className="text-xs text-slate-500 font-medium">尚未建立任何辅料项，请手动录入、复制粘贴，或点击下方按钮采用示例台账。</p>
                          <button
                            type="button"
                            onClick={loadDemoItems}
                            className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-4 py-2 rounded-lg text-xs shadow-xs transition-all cursor-pointer flex items-center gap-1.5"
                          >
                            <Sparkles className="w-4 h-4" />
                            <span>采用示例辅材台账</span>
                          </button>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    ledger.map((item) => {
                      const getFriendlyCategory = (cat: string) => {
                        if (cat === 'gasket') return '型材气密胶条类';
                        if (cat === 'fastener') return '不锈钢紧固五金螺栓类';
                        if (cat === 'auxiliary') return '结构密封胶及辅助垫块类';
                        return '其它辅材类';
                      };

                      return (
                        <tr key={item.id} className="hover:bg-indigo-50/10 transition-colors">
                          <td className="px-3 py-2.5 font-bold text-slate-900">{item.name}</td>
                          <td className="px-3 py-2.5 font-mono text-indigo-700 font-semibold">{item.drawingNo}</td>
                          <td className="px-3 py-2.5 text-slate-655 font-semibold">{item.materialModel}</td>
                          <td className="px-3 py-2.5 font-mono text-slate-500">{item.size || "—"}</td>
                          <td className="px-3 py-2.5 text-slate-700 font-medium whitespace-nowrap">
                            {(() => {
                              const calcInfo = calculatedMap.get(item.id);
                              const isCalculated = calcInfo && calcInfo.isCalculated;
                              const COUNT_UNITS = ['个', '支', '块', '只', '套'];
                              const cleanUnit = item.unit ? item.unit.trim() : '';
                              const isCountUnit = COUNT_UNITS.includes(cleanUnit);

                              const formattedVal = calcInfo ? (
                                isCountUnit
                                  ? Math.ceil(calcInfo.unitUsage)
                                  : Number(calcInfo.unitUsage.toFixed(3))
                              ) : 0;

                              return isCalculated ? (
                                <div className="flex flex-col gap-0.5">
                                  <span className="font-mono text-indigo-700 font-bold text-xs bg-indigo-50 px-2 py-0.5 rounded border border-indigo-150 inline-block w-fit">
                                    {formattedVal}
                                  </span>
                                  <span className="text-[10px] text-slate-400">
                                    {calcInfo.rulesCount}条规则套算 ({item.rules?.length || 0}条配置){item.position ? ` · ${item.position}` : ''}
                                  </span>
                                </div>
                              ) : (
                                <div className="flex flex-col gap-0.5">
                                  {item.rules && item.rules.length > 0 ? (
                                    <span className="bg-slate-100 text-slate-500 px-2 py-0.5 rounded text-[9px] font-bold inline-block w-fit">
                                      {item.rules.filter(r => r.isActive).length}条算法 (待导入料单)
                                    </span>
                                  ) : (
                                    <span className="text-slate-400 font-mono text-xs">—</span>
                                  )}
                                  {item.position && <span className="text-[10px] text-slate-400">{item.position}</span>}
                                </div>
                              );
                            })()}
                          </td>
                          <td className="px-3 py-2.5 text-right font-mono text-slate-450 font-bold">—</td>
                          <td className="px-3 py-2.5 font-bold text-slate-500">{item.unit}</td>
                          <td className="px-3 py-2.5 font-mono text-slate-400">—</td>
                          <td className="px-3 py-2.5 font-mono text-slate-400">—</td>
                          <td className="px-3 py-2.5 text-slate-500 truncate" title={item.remark}>{item.remark || "—"}</td>
                          <td className="px-3 py-2.5 text-slate-600 font-semibold whitespace-nowrap">
                            <span className="bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-bold">
                              {getFriendlyCategory(item.category)}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 text-center whitespace-nowrap">
                            <div className="flex items-center justify-center gap-1.5">
                              <button
                                onClick={() => handleStartEdit(item)}
                                className="text-slate-405 hover:text-indigo-600 hover:bg-indigo-50 p-1.5 rounded transition-all cursor-pointer"
                                title="编辑辅材规格与单樘算法"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => onRemoveLedgerItem(item.id)}
                                className="text-slate-405 hover:text-rose-600 hover:bg-rose-50 p-1.5 rounded transition-all cursor-pointer"
                                title="从台账中除外"
                              >
                                <Trash className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>

      {/* Edit Modal Dialog for Ledger Item and Nested Formulas */}
      {editingItem && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xl max-w-2xl w-full overflow-hidden flex flex-col max-h-[90vh]">
            
            {/* Modal Header */}
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-2">
                <Sliders className="w-5 h-5 text-indigo-600" />
                <div>
                  <h3 className="text-sm font-bold text-slate-800">编辑辅配料台账与单樘用量算法</h3>
                  <p className="text-[10px] text-slate-400 mt-0.5">更改材质详情并可在线定义多条并行的原材料算量映射公式</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setEditingItem(null)}
                className="text-slate-400 hover:text-slate-705 hover:bg-slate-100 p-1.5 rounded-lg transition-all"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Modal Scrollable Body */}
            <div className="p-5 overflow-y-auto space-y-4 text-xs flex-1">
              
              {/* Part 1: Standard ledger record information */}
              <div className="space-y-2">
                <h4 className="font-bold text-indigo-700 text-[11px] uppercase tracking-wider flex items-center gap-1 border-b border-slate-100 pb-1">
                  <FileText className="w-3.5 h-3.5" />
                  辅配料基础规格定义
                </h4>
                
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-450 font-bold mb-1">大类分类归属</label>
                    <select
                      value={editingItem.category}
                      onChange={(e) => handleUpdateEditingField('category', e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 font-semibold text-slate-800"
                    >
                      <option value="gasket">型材气密胶条类</option>
                      <option value="fastener">不锈钢紧固五金螺栓类</option>
                      <option value="auxiliary">结构密封胶及辅助垫块类</option>
                    </select>
                  </div>

                  <div>
                    <label className="block text-slate-450 font-bold mb-1">辅料名称</label>
                    <input
                      type="text"
                      value={editingItem.name}
                      onChange={(e) => handleUpdateEditingField('name', e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800 font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-450 font-bold mb-1">加工图号</label>
                    <input
                      type="text"
                      value={editingItem.drawingNo}
                      onChange={(e) => handleUpdateEditingField('drawingNo', e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800 font-mono font-semibold"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-450 font-bold mb-1">材质及型号</label>
                    <input
                      type="text"
                      value={editingItem.materialModel}
                      onChange={(e) => handleUpdateEditingField('materialModel', e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-450 font-bold mb-1">尺寸规格</label>
                    <input
                      type="text"
                      value={editingItem.size}
                      onChange={(e) => handleUpdateEditingField('size', e.target.value)}
                      className="w-full bg-slate-550 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800 font-mono"
                    />
                  </div>

                  <div>
                    <label className="block text-slate-450 font-bold mb-1">核算计量单位</label>
                    <input
                      type="text"
                      value={editingItem.unit}
                      onChange={(e) => handleUpdateEditingField('unit', e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800"
                    />
                  </div>

                  <div className="col-span-2 grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-slate-450 font-bold mb-1">备注说明</label>
                      <input
                        type="text"
                        value={editingItem.remark}
                        onChange={(e) => handleUpdateEditingField('remark', e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800"
                      />
                    </div>
                    <div>
                      <label className="block text-slate-450 font-bold mb-1">安装位置</label>
                      <input
                        type="text"
                        value={editingItem.position}
                        onChange={(e) => handleUpdateEditingField('position', e.target.value)}
                        className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded p-1.5 text-slate-800"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {/* Part 2: Rules and formulas nested inside editingItem */}
              <div className="space-y-3 pt-2">
                <div className="flex justify-between items-center border-b border-slate-100 pb-1.5">
                  <h4 className="font-bold text-indigo-700 text-[11px] uppercase tracking-wider flex items-center gap-1">
                    <Sparkles className="w-3.5 h-3.5" />
                    单樘用量智能计算套算规则 ({editingItem.rules?.length || 0} 条分项算法)
                  </h4>
                  <button
                    type="button"
                    onClick={handleAddRuleToEditing}
                    className="bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 font-bold px-2 py-0.5 rounded text-[10px] flex items-center gap-1 transition-all cursor-pointer select-none"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    增加换算公式规则
                  </button>
                </div>

                <div className="space-y-3.5 max-h-[220px] overflow-y-auto pr-1">
                  {!editingItem.rules || editingItem.rules.length === 0 ? (
                    <div className="p-6 text-center border border-dashed border-slate-200 rounded-lg text-slate-400 italic leading-relaxed">
                      暂无针对该辅材项的动态算料 rules。
                      <br className="mb-1" />
                      点击右上角“增加换算公式规则”，设置自动套用型材、面板或五金数量公式！
                    </div>
                  ) : (
                    editingItem.rules.map((rule) => (
                      <div key={rule.id} className="p-3 bg-slate-50 rounded-lg border border-slate-200/80 space-y-2">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2 flex-1">
                            <input
                              type="text"
                              value={rule.name}
                              onChange={(e) => handleUpdateRuleField(rule.id, 'name', e.target.value)}
                              placeholder="算料规则名称，例如: 型材H2-JT19配套等"
                              className="bg-white border border-slate-200 rounded px-2.5 py-1 text-xs font-bold text-slate-800 flex-1 focus:border-indigo-500"
                            />
                            
                            <button
                              type="button"
                              onClick={() => {
                                setTargetRuleId(rule.id);
                                setSelectedRowId(null);
                                setSearchKeyword("");
                                setIsAllAttributesExpanded(false);
                                setParamModalOpen(true);
                              }}
                              className="bg-indigo-50 hover:bg-indigo-150 border border-indigo-200 text-indigo-700 font-bold px-2 py-1 rounded text-[10px] flex items-center gap-1 transition-all cursor-pointer shrink-0"
                              title="在公式命名的右侧一键唤起料单列表插入参数"
                            >
                              <Plus className="w-3.5 h-3.5" />
                              <span>插入参数</span>
                            </button>
                          </div>
                          
                          <div className="flex items-center gap-1.5 shrink-0">
                            <button
                              type="button"
                              onClick={() => {
                                setLengthenRuleId(rule.id);
                                setLengthenValue(String(rule.lengthAdjustment || 0));
                                setLengthenModalOpen(true);
                              }}
                              className={`px-2 py-0.5 rounded text-[10px] font-bold transition-all flex items-center gap-1 border cursor-pointer ${
                                rule.lengthAdjustment
                                  ? "bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100 font-extrabold"
                                  : "bg-slate-100 text-slate-450 border-slate-250 hover:bg-slate-200"
                              }`}
                              title="加长：点击弹出加长尺寸输入框，将每一个匹配到的型材长度、玻璃/面板尺寸都加长一定长度带入公式计算"
                            >
                              <Ruler className={`w-3.5 h-3.5 ${rule.lengthAdjustment ? "text-amber-600 animate-pulse" : "text-slate-400"}`} />
                              <span>{rule.lengthAdjustment ? `加长: +${rule.lengthAdjustment}mm` : "加长"}</span>
                            </button>

                            <button
                              type="button"
                              onClick={() => handleUpdateRuleField(rule.id, 'isActive', !rule.isActive)}
                              className={`px-2 py-0.5 rounded text-[10px] font-bold transition-all flex items-center gap-1 border cursor-pointer ${
                                rule.isActive
                                  ? "bg-emerald-50 text-emerald-700 border-emerald-300"
                                  : "bg-slate-100 text-slate-450 border-slate-250 hover:bg-slate-200"
                              }`}
                            >
                              <span className={`w-1.5 h-1.5 rounded-full ${rule.isActive ? "bg-emerald-500 animate-pulse" : "bg-slate-400"}`} />
                              {rule.isActive ? "启用" : "暂停"}
                            </button>

                            <button
                              type="button"
                              onClick={() => handleUpdateRuleField(rule.id, 'isExclusive', !rule.isExclusive)}
                              className={`px-2 py-0.5 rounded text-[10px] font-bold transition-all flex items-center gap-1 border cursor-pointer ${
                                rule.isExclusive
                                  ? "bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100"
                                  : "bg-slate-100 text-slate-450 border-slate-250 hover:bg-slate-200"
                              }`}
                              title="排它模式：开启此选项后，若该公式计算结果不为0，则该辅材下的其它普通（未开启排它）的算料公式不参与计算 (重置为0)。"
                            >
                              <span className={`w-1.5 h-1.5 rounded-full ${rule.isExclusive ? "bg-amber-500 animate-pulse" : "bg-slate-400"}`} />
                              {rule.isExclusive ? "排它：开" : "排它：关"}
                            </button>

                            <button
                              type="button"
                              onClick={() => handleDuplicateRule(rule)}
                              className="text-slate-400 hover:text-indigo-600 hover:bg-slate-100 p-1 rounded transition-all cursor-pointer"
                              title="复制此算料公式规则"
                            >
                              <Copy className="w-3.5 h-3.5" />
                            </button>

                            <button
                              type="button"
                              onClick={() => handleDeleteRule(rule.id)}
                              className="text-slate-400 hover:text-rose-600 hover:bg-rose-55 p-1 rounded transition-all cursor-pointer"
                              title="删除此公式规则"
                            >
                              <Trash className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>

                        <div className="space-y-1.5 text-xs text-slate-700">
                          <div className="flex gap-2 items-start">
                            <div className="font-mono text-[10px] text-slate-400 shrink-0 select-none mt-2">公式预览:</div>
                            <div className="flex-1">
                              {renderFormulaVisualizer(rule.id, rule.expression)}
                              <p className="text-[9px] text-slate-400 mt-1 select-none">
                                💡 上方青色 &lt;搜索词&gt; 代表一串后台动态取数公式。<b>双击标签</b>可直接修改公式参数。
                              </p>
                            </div>
                          </div>
                          
                          <div className="flex gap-2 items-start">
                            <div className="font-mono text-[10px] text-slate-400 shrink-0 select-none mt-1.5">编辑表达式:</div>
                          <div className="flex flex-col gap-2 flex-1">
                            <textarea
                              id={`rule-textarea-list-${rule.id}`}
                              rows={15}
                              value={rule.expression}
                              onChange={(e) => handleUpdateRuleField(rule.id, 'expression', e.target.value)}
                              placeholder="如: SUM_LEN('H2-JT19') * 1.0, 亦可通过上面插入参数直接构筑"
                              className="bg-white border border-slate-200 rounded-md px-2.5 py-1.5 text-xs font-semibold font-mono text-slate-850 w-full focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 resize-y min-h-[260px] overflow-y-auto"
                            />
                            
                            <div className="p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-slate-600 text-[10.5px] flex items-center gap-1.5">
                              <HelpCircle className="w-4 h-4 text-indigo-500 shrink-0" />
                              <span>💡 公式 AND/OR 且/或逻辑编写、模糊匹配规则已统一置于应用页面最底部，可随时折叠展开查看。</span>
                            </div>
                          </div>
                          </div>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>

            {/* Modal Actions Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 shrink-0">
              <button
                type="button"
                onClick={() => setEditingItem(null)}
                className="bg-white text-slate-700 hover:bg-slate-100 border border-slate-300 px-4 py-2 rounded text-xs font-bold transition-all cursor-pointer"
              >
                取消
              </button>
              <button
                type="button"
                onClick={handleSaveEdit}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded text-xs font-bold transition-all shadow-md cursor-pointer"
              >
                确认修改并应用 (Apply Changes)
              </button>
            </div>

          </div>
        </div>
      )}

      {/* Lengthen Modal Popup */}
      {lengthenModalOpen && lengthenRuleId && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-[100] p-4">
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xl p-5 w-full max-w-sm space-y-4">
            <div className="flex justify-between items-center border-b border-slate-100 pb-2.5 select-none">
              <h3 className="text-xs font-bold text-slate-800 flex items-center gap-1.5 font-sans">
                <Ruler className="w-4 h-4 text-indigo-600" />
                <span>设置此规则计算加长尺寸</span>
              </h3>
              <button
                type="button"
                onClick={() => {
                  setLengthenModalOpen(false);
                  setLengthenRuleId(null);
                }}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-md hover:bg-slate-50 transition-all cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="space-y-2">
              <p className="text-[11px] text-slate-500 leading-relaxed font-semibold">
                请输入加长尺寸（单位：毫米 mm）。匹配到的所有<b>型材长度</b>、<b>玻璃/面板尺寸规格</b>都会在公式计算前累加该长度。
              </p>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  value={lengthenValue}
                  onChange={(e) => setLengthenValue(e.target.value)}
                  placeholder="如：50"
                  className="flex-1 bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 rounded-lg p-2 text-slate-800 font-mono text-sm font-bold text-center"
                  autoFocus
                />
                <span className="text-xs font-bold text-slate-400 select-none">mm</span>
              </div>
            </div>
            <div className="flex justify-end gap-2 border-t border-slate-100 pt-3 select-none">
              <button
                type="button"
                onClick={() => {
                  handleUpdateRuleField(lengthenRuleId, 'lengthAdjustment', 0);
                  setLengthenModalOpen(false);
                  setLengthenRuleId(null);
                }}
                className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer border border-slate-200"
              >
                清除加长 (Reset)
              </button>
              <button
                type="button"
                onClick={() => {
                  const val = parseFloat(lengthenValue) || 0;
                  handleUpdateRuleField(lengthenRuleId, 'lengthAdjustment', val);
                  setLengthenModalOpen(false);
                  setLengthenRuleId(null);
                }}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer shadow-sm shadow-indigo-100"
              >
                确认加长 (Confirm)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Token inline editor modal */}
      {editingToken && (
        <div className="fixed inset-0 z-60 bg-slate-900/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xl max-w-sm w-full overflow-hidden flex flex-col">
            <div className="p-4 bg-indigo-50 border-b border-indigo-100 flex justify-between items-center select-none">
              <h3 className="font-bold text-xs text-indigo-900 flex items-center gap-1.5">
                <Sparkles className="w-4 h-4 text-indigo-600" />
                <span>修改提取公式和检索条件</span>
              </h3>
              <button
                type="button"
                onClick={() => setEditingToken(null)}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg transition-colors cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <div className="p-4 space-y-3.5 text-xs font-semibold text-slate-700">
              <div className="space-y-1">
                <label className="text-[11px] text-slate-400 block font-normal">🎯 对应的搜索词 (Keyword)</label>
                <input
                  type="text"
                  value={editingToken.keyword}
                  onChange={(e) => setEditingToken(prev => prev ? { ...prev, keyword: e.target.value } : null)}
                  className="w-full bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 transition-all font-mono"
                  placeholder="如: H2-JT19"
                />
              </div>

              <div className="space-y-1">
                <label className="text-[11px] text-slate-400 block font-normal">📁 检索归属分类 (Category)</label>
                <select
                  value={editingToken.category}
                  onChange={(e) => setEditingToken(prev => prev ? { ...prev, category: e.target.value } : null)}
                  className="w-full bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 transition-all font-semibold"
                >
                  <option value="profile">型材类 (profile)</option>
                  <option value="panel">面板类 (panel)</option>
                  <option value="hardware">五金类 (hardware)</option>
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[11px] text-slate-400 block font-normal font-mono">🔍 检索关键词匹配列 (Search Field)</label>
                <select
                  value={editingToken.searchField}
                  onChange={(e) => setEditingToken(prev => prev ? { ...prev, searchField: e.target.value } : null)}
                  className="w-full bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 transition-all font-mono font-semibold"
                >
                  <option value="name font-sans">名称 (name)</option>
                  <option value="drawingNo">图纸号/图号 (drawingNo)</option>
                  <option value="materialModel">材质型号 (materialModel)</option>
                  <option value="cuttingSize">下料尺寸 (cuttingSize)</option>
                  <option value="remark">备注 (remark)</option>
                </select>
              </div>

              <div className="space-y-1">
                <label className="text-[11px] text-slate-400 block font-normal font-mono">📊 目标数据提取源列 (Target Value Field)</label>
                <select
                  value={editingToken.targetField}
                  onChange={(e) => setEditingToken(prev => prev ? { ...prev, targetField: e.target.value } : null)}
                  className="w-full bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 text-xs font-bold text-slate-900 focus:bg-white focus:outline-none focus:ring-1 focus:ring-indigo-500 focus:border-indigo-500 transition-all font-mono font-semibold"
                >
                  <option value="cuttingSize">下料尺寸 (cuttingSize)</option>
                  <option value="cuttingWidth">下料宽 (cuttingWidth)</option>
                  <option value="cuttingHeight">下料高 (cuttingHeight)</option>
                  <option value="unitUsage">单樘用量 (unitUsage)</option>
                  <option value="l1">起步尺寸 L1 (l1)</option>
                  <option value="l2">收尾尺寸 L2 (l2)</option>
                  <option value="drawingNo">图号 (drawingNo)</option>
                  <option value="materialModel">材质型号 (materialModel)</option>
                  <option value="remark">备注 (remark)</option>
                </select>
              </div>

              <div className="pt-2">
                <label className="flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={editingToken.isExactMatch || false}
                    onChange={(e) => setEditingToken(prev => prev ? { ...prev, isExactMatch: e.target.checked } : null)}
                    className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 bg-white w-3.5 h-3.5"
                  />
                  <span className="text-[11px] font-bold text-slate-600">强制全字精确匹配 (Exact Match)</span>
                </label>
              </div>
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 text-xs shrink-0 select-none">
              <button
                type="button"
                onClick={() => setEditingToken(null)}
                className="bg-white text-slate-700 hover:bg-slate-100 border border-slate-300 px-4 py-2 rounded font-bold transition-all cursor-pointer"
              >
                取消
              </button>
              <button
                type="button"
                onClick={() => {
                  const updatedTokenString = `ROW_VAL('${editingToken.category}', '${editingToken.searchField}', '${editingToken.keyword}', '${editingToken.targetField}', '${editingToken.isExactMatch ? 'true' : 'false'}')`;
                  const currentExpr = editingToken.originalExpression;
                  const segments = parseExpression(currentExpr);
                  
                  const updatedSegments = segments.map((seg, idx) => {
                    if (idx === editingToken.segmentIndex) {
                      return updatedTokenString;
                    }
                    return seg.type === 'token' 
                      ? (seg.isExactMatch !== undefined 
                          ? `ROW_VAL('${seg.category}', '${seg.searchField}', '${seg.keyword}', '${seg.targetField}', '${seg.isExactMatch ? 'true' : 'false'}')`
                          : `ROW_VAL('${seg.category}', '${seg.searchField}', '${seg.keyword}', '${seg.targetField}')`)
                      : seg.text;
                  });
                  
                  const newExpression = updatedSegments.join('');
                  handleUpdateRuleField(editingToken.ruleId, 'expression', newExpression);
                  setEditingToken(null);
                }}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded font-bold transition-all shadow-md cursor-pointer"
              >
                保存修改
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Parameter Selection and Fuzzy Insertion Dialog */}
      {paramModalOpen && (
        <div className="fixed inset-0 z-55 overflow-y-auto bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl border border-slate-200 shadow-2xl max-w-3xl w-full overflow-hidden flex flex-col max-h-[85vh]">
            
            {/* Header */}
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex justify-between items-center shrink-0">
              <div className="flex items-center gap-2">
                <Compass className="w-5 h-5 text-indigo-600 animate-spin-slow" />
                <div>
                  <h3 className="text-sm font-bold text-slate-800">料单位置参数一键智能引入工具</h3>
                  <p className="text-[10px] text-slate-400 mt-0.5">选取“料单导入与区间映射”中具体工艺行的动态数据，秒级套进算料公式</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setParamModalOpen(false)}
                className="text-slate-400 hover:text-slate-800 hover:bg-slate-100 p-1.5 rounded-lg transition-all"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Scrollable grid body of selectors */}
            <div className="p-5 overflow-y-auto space-y-4 flex-1 text-xs">
              
              {/* Step 1: Material logic category select */}
              <div className="space-y-1.5">
                <span className="block text-[11px] font-bold text-indigo-700 uppercase tracking-wide">
                  1. 选择大类逻辑归属 (对应导入料单中的品类)
                </span>
                
                <div className="grid grid-cols-3 gap-1.5">
                  {[
                    { val: 'profile', label: '型材龙骨类' },
                    { val: 'panel', label: '面板玻璃类' },
                    { val: 'steel', label: '五金角码类' },
                  ].map(catItem => {
                    const isSelected = paramCategory === catItem.val;
                    return (
                      <button
                        key={catItem.val}
                        type="button"
                        onClick={() => {
                          setParamCategory(catItem.val);
                          setSelectedRowId(null);
                        }}
                        className={`p-2 rounded border text-left transition-all font-bold cursor-pointer ${
                          isSelected
                            ? "bg-indigo-600 border-indigo-600 text-white shadow-sm"
                            : "bg-slate-50 border-slate-200 text-slate-700 hover:bg-slate-100"
                        }`}
                      >
                        <div className="text-[10px] opacity-90">{catItem.label}</div>
                        <div className="text-[8px] font-mono opacity-60 uppercase">{catItem.val}</div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Step 2: Determine Header & Query Fuzzy search keyword */}
              <div className="space-y-2 pt-1">
                <span className="block text-[11px] font-bold text-indigo-700 uppercase tracking-wide">
                  2. 定位定位列表头 (选择查找字段并模糊匹配)
                </span>
                
                <div className="flex flex-col sm:flex-row gap-2 items-stretch sm:items-center">
                  <div className="flex items-center gap-1 shrink-0">
                    <span className="text-slate-400 font-medium">搜索表头:</span>
                    <select
                      value={searchHeader}
                      onChange={(e) => {
                        setSearchHeader(e.target.value);
                        setSelectedRowId(null);
                      }}
                      className="bg-white border border-slate-200 rounded px-2.5 py-1.5 text-xs font-bold text-slate-700 focus:border-indigo-500 cursor-pointer"
                    >
                      <option value="name">名称 (Name)</option>
                      <option value="drawingNo">加工图号 (Drawing No)</option>
                      <option value="materialModel">材质及型号 (Model)</option>
                      <option value="cuttingSize">下料尺寸 (Size)</option>
                      <option value="unitUsage">单樘用量 (Unit Usage)</option>
                      <option value="l1">L1 尺寸</option>
                      <option value="l2">L2 尺寸</option>
                      <option value="remark">备注说明 (Remarks)</option>
                    </select>
                  </div>

                  <div className="relative flex-1">
                    <Search className="absolute left-2.5 top-2.5 w-3.5 h-3.5 text-slate-400" />
                    <input
                      type="text"
                      value={searchKeyword}
                      onChange={(e) => {
                        setSearchKeyword(e.target.value);
                        setSelectedRowId(null);
                      }}
                      placeholder="在此输入模糊拼音或规格文字进行自动过滤匹配..."
                      className="w-full bg-slate-50 border border-slate-200 focus:border-indigo-500 focus:bg-white rounded pl-8 pr-3 py-1.5 text-xs text-slate-800 font-medium"
                    />
                  </div>
                  <label className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 cursor-pointer hover:bg-slate-100 transition-colors shrink-0">
                    <input
                      type="checkbox"
                      checked={isSearchExactMatch}
                      onChange={(e) => {
                        setIsSearchExactMatch(e.target.checked);
                        setSelectedRowId(null);
                      }}
                      className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 bg-white w-3.5 h-3.5"
                    />
                    <span className="text-xs font-bold text-slate-600">精确匹配</span>
                  </label>
                </div>
              </div>

                             {/* Step 3: List of matched rows in spreadsheet */}
              <div className="space-y-2 pt-1">
                <div className="flex justify-between items-center">
                  <span className="text-[11px] font-bold text-indigo-700 uppercase tracking-wide">
                    3. 核对确定目标物理行清单 ({filteredItems.length} 行符合条件)
                  </span>
                  {selectedRowId !== null && (
                    <span className="text-[10px] text-emerald-600 font-bold bg-emerald-50 px-2 py-0.5 rounded border border-emerald-200">
                      ★ 已锁定第 {selectedRowId + 1} 行
                    </span>
                  )}
                </div>

                <div className="border border-slate-200 rounded-lg overflow-hidden bg-slate-50 max-h-[160px] overflow-y-auto">
                  {filteredItems.length === 0 ? (
                    <div className="p-8 text-center text-slate-400 italic">
                      未检索到匹配的料单行项目。请切换大类或调整匹配关键字！
                    </div>
                  ) : (
                    <table className="min-w-full divide-y divide-slate-100 text-[10px] bg-white">
                      <thead className="bg-slate-50 text-slate-500 font-bold sticky top-0 bg-slate-100">
                        <tr>
                          <th className="px-2 py-1.5 text-left w-10">行号</th>
                          <th className="px-2 py-1.5 text-left">名称</th>
                          <th className="px-2 py-1.5 text-left">加工图号</th>
                          <th className="px-2 py-1.5 text-left">材质牌号</th>
                          <th className="px-2 py-1.5 text-left">下料尺寸</th>
                          <th className="px-2 py-1.5 text-left">L1/L2</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredItems.map((row, idx) => {
                          const isSelected = selectedRowId === idx;
                          return (
                            <tr
                              key={idx}
                              onClick={() => setSelectedRowId(idx)}
                              className={`cursor-pointer hover:bg-indigo-50/40 transition-colors ${
                                isSelected ? "bg-indigo-50 font-bold text-indigo-900" : "text-slate-700"
                              }`}
                            >
                              <td className="px-2 py-1.5 font-mono text-slate-400">{idx + 1}</td>
                              <td className="px-2 py-1.5 truncate max-w-[120px]" title={row.name}>{row.name}</td>
                              <td className="px-2 py-1.5 font-mono truncate max-w-[100px]" title={row.drawingNo}>{row.drawingNo || "—"}</td>
                              <td className="px-2 py-1.5 truncate max-w-[90px]" title={row.materialModel}>{row.materialModel || "—"}</td>
                              <td className="px-2 py-1.5 font-mono">{row.cuttingSize || "—"}</td>
                              <td className="px-2 py-1.5 font-mono">
                                {row.l1 !== undefined || row.l2 !== undefined ? `${row.l1 || 0}/${row.l2 || 0}` : "—"}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>

              {/* Step 4: Click a row to see individual parameters */}
              {selectedRowId !== null && filteredItems[selectedRowId] ? (() => {
                const row = filteredItems[selectedRowId];
                
                // Construct baseParameterFields
                const baseParameterFields: { label: string; key: string; val: any }[] = [
                  { label: "名称", key: "name", val: row.name },
                  { label: "加工图号", key: "drawingNo", val: row.drawingNo },
                  { label: "材质型号", key: "materialModel", val: row.materialModel },
                ];

                const cuttingSizeVal = row.cuttingSize || "";
                const csMatch = cuttingSizeVal.match(/([\d.]+)\*([\d.]+)/);
                if (csMatch) {
                  baseParameterFields.push({ label: "下料宽(W)", key: "cuttingWidth", val: csMatch[1] });
                  baseParameterFields.push({ label: "下料高(H)", key: "cuttingHeight", val: csMatch[2] });
                } else {
                  baseParameterFields.push({ label: "下料尺寸", key: "cuttingSize", val: cuttingSizeVal });
                }

                baseParameterFields.push(
                  { label: "单樘用量", key: "unitUsage", val: row.unitUsage },
                  { label: "L1", key: "l1", val: row.l1 },
                  { label: "L2", key: "l2", val: row.l2 },
                  { label: "备注", key: "remark", val: row.remark },
                  { label: "算料逻辑归属", key: "category", val: row.category }
                );

                const coreKeys = ["cuttingSize", "cuttingWidth", "cuttingHeight", "unitUsage", "l1", "l2"];
                const visibleFields = isAllAttributesExpanded 
                  ? baseParameterFields 
                  : baseParameterFields.filter(f => coreKeys.includes(f.key));

                return (
                  <div className="space-y-3 pt-2 border-t border-slate-100">
                    <span className="block text-[11px] font-bold text-indigo-700 uppercase tracking-wide">
                      4. 匹配出的属性值清单 (点击按钮一键引入)
                    </span>
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                      {visibleFields.map(field => {
                        const displayVal = field.val || "—";
                        const docVal = String(displayVal).trim();
                        const isEmpty = !docVal || docVal === "—" || docVal === "";
                        
                        return (
                          <div
                            key={field.key}
                            className={`p-2.5 rounded-lg border flex flex-col justify-between transition-all bg-white shadow-2xs border-slate-200 hover:border-indigo-400 hover:bg-indigo-50/20 active:scale-[0.98]`}
                          >
                            <div className="flex justify-between items-center">
                              <span className="text-[10px] text-slate-400 font-bold">
                                {field.label}
                                {((paramCategory === "profile" && ["cuttingSize", "l1", "l2"].includes(field.key)) || ["cuttingWidth", "cuttingHeight"].includes(field.key)) && (
                                  <span className="text-[8px] text-amber-600 font-bold bg-amber-50 px-1 py-0.5 rounded border border-amber-100 ml-1" title="型材默认毫米(mm)或宽高尺寸，在配套算料时会自动折算为米(m)">mm→m</span>
                                )}
                              </span>
                              <span className="text-[8px] font-mono text-slate-300 uppercase">{field.key}</span>
                            </div>
                            <div className={"text-[11px] font-bold text-slate-800 truncate mt-1 select-all" + (isEmpty ? " text-slate-300 italic" : "")} title={docVal}>
                              {docVal}
                            </div>
                            
                            <div className="mt-2 flex flex-col gap-1 pt-1.5 border-t border-dashed border-slate-100">
                              <button
                                type="button"
                                onClick={() => handleSelectDynamicRowVal(field.key, row.name)}
                                className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-1 px-1.5 rounded text-[8.5px] transition-all cursor-pointer flex items-center justify-center gap-0.5"
                                title="在计算辅材时，动态到所选大类匹配名称等条件的行中读取这一列的数据进行公式计算。此功能无需手工写死固定数据。"
                              >
                                <Sparkles className="w-2.5 h-2.5 shrink-0" />
                                <span>✨ 动态公式参数</span>
                              </button>
                              
                              {!isEmpty && (
                                <div className="flex gap-1 justify-between">
                                  {["l1", "l2", "cuttingSize", "cuttingWidth", "cuttingHeight", "unitUsage"].includes(field.key) && String(docVal).match(/[0-9]/) && (
                                    <button
                                      type="button"
                                      onClick={() => handleSelectCellData(field.key, docVal, true, row.unitUsage)}
                                      className="bg-slate-100 hover:bg-slate-250 text-slate-700 font-semibold px-1 py-0.5 rounded text-[8px] transition-all cursor-pointer flex-1 text-center"
                                      title="仅提取该单元格内的数字作为固定死数值插入公式 (如 1500)"
                                    >
                                      固化数字
                                    </button>
                                  )}
                                  <button
                                    type="button"
                                    onClick={() => handleSelectCellData(field.key, docVal, false, row.unitUsage)}
                                    className="bg-slate-100 hover:bg-slate-250 text-slate-700 font-semibold px-1 py-0.5 rounded text-[8px] transition-all cursor-pointer flex-1 text-center"
                                    title="按原始字符串值作为死参数插入公式 (如 'H2-JT')"
                                  >
                                    固化字符串
                                  </button>
                                </div>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                    
                    <div className="flex justify-center pt-1.5">
                      <button
                        type="button"
                        onClick={() => setIsAllAttributesExpanded(!isAllAttributesExpanded)}
                        className="inline-flex items-center gap-1.5 px-4 py-1.5 bg-slate-50 hover:bg-slate-100 text-slate-700 font-bold text-[10.5px] rounded-lg border border-slate-250 transition-all cursor-pointer select-none active:scale-[0.98]"
                      >
                        {isAllAttributesExpanded ? (
                          <>
                            <ChevronUp className="w-3.5 h-3.5 text-slate-500" />
                            <span>收起其它非核心属性</span>
                          </>
                        ) : (
                          <>
                            <ChevronDown className="w-3.5 h-3.5 text-slate-500" />
                            <span>展开显示全部属性 (包含名称、图号、备注、大类等)</span>
                          </>
                        )}
                      </button>
                    </div>
                  </div>
                );
              })() : (
                <div className="p-6 text-center border border-dashed border-slate-200 bg-slate-50/50 rounded-lg text-slate-400 select-none">
                  ⚡ 请在第 3 步中鼠标点击选中一行，或在此输入搜索词后直接选取生成动态公式。
                </div>
              )}

              {/* Formula Preview bar */}
              {editingItem && targetRuleId && (
                <div className="space-y-2">
                  <div className="p-3 bg-indigo-50 border border-indigo-100 rounded-lg space-y-1">
                    <div className="flex justify-between items-center select-none">
                      <span className="text-[10px] text-indigo-700 font-bold">🎯 当前公式规则和表达式编辑预览:</span>
                      <button
                        type="button"
                        onClick={() => {
                          setEditingItem(prev => {
                            if (!prev || !prev.rules) return prev;
                            return {
                              ...prev,
                              rules: prev.rules.map(r => r.id === targetRuleId ? { ...r, expression: "" } : r)
                            };
                          });
                        }}
                        className="text-rose-600 hover:bg-rose-100 px-1.5 py-0.5 rounded text-[9px] font-bold transition-all cursor-pointer"
                      >
                        清空规则公式
                      </button>
                    </div>
                    <div className="space-y-2">
                      <div className="border border-indigo-200/50 rounded-lg bg-indigo-50/30 p-2">
                        <div className="text-[10px] text-indigo-600/85 mb-1.5 select-none font-bold">✨ 实时预览 & 连带修改:</div>
                        {renderFormulaVisualizer(targetRuleId, editingItem.rules?.find(r => r.id === targetRuleId)?.expression || "")}
                        <p className="text-[9px] text-slate-400 mt-1 select-none">
                          💡 上方的青色 &lt;搜索词&gt; 模块支持<b>双击一键编辑修改取数规则</b>。
                        </p>
                      </div>

                      <div className="bg-white/80 border border-indigo-150/85 rounded-lg p-2.5 select-none shadow-xs">
                        <div className="text-[10px] text-indigo-750 font-bold mb-1.5 flex items-center gap-1">
                          <Sparkles className="w-3.5 h-3.5 text-indigo-500 shrink-0" />
                          <span>工程参量与板块匹配便捷插入 (若匹配成功，算料结果返回 1，否则为 0):</span>
                        </div>
                        <div className="flex flex-wrap gap-1.5">
                          {[
                            { label: "板宽 W", insert: "W", desc: "板块宽度工程参量 (米)" },
                            { label: "板高 H", insert: "H", desc: "板块高度工程参量 (米)" },
                            { label: "板副宽 W1", insert: "W1", desc: "板块副宽工程参量 (米)" },
                            { label: "匹配编号 'D01'", insert: "MATCH('板块编号', 'D01')", desc: "当板块编号为 D01 时计算生效" },
                            { label: "匹配型号 'UNIT'", insert: "MATCH('板块型号', 'UNIT-140-1500')", desc: "当板块型号包含特定特征时生效" },
                            { label: "宽 W >= 1.5m", insert: "MATCH('W', '>=1.5')", desc: "当板宽大于等于 1.5 米时生效" },
                            { label: "高 H >= 2.0m", insert: "MATCH('H', '>=2.0')", desc: "当板高大于等于 2.0 米时生效" },
                            { label: "副宽 W1 == 0.5m", insert: "MATCH('W1', '==0.5')", desc: "当副宽等于 0.5 米时生效" },
                          ].map((btn, idx) => (
                            <button
                              key={idx}
                              type="button"
                              onClick={() => handleInsertFormulaText(btn.insert)}
                              className="bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 font-bold px-2 py-1 rounded text-[10px] transition-all cursor-pointer select-none active:scale-95"
                              title={btn.desc}
                            >
                              {btn.label}
                            </button>
                          ))}
                        </div>
                      </div>

                      <div className="relative">
                        <textarea
                          id={`rule-textarea-edit-${targetRuleId}`}
                          rows={15}
                          value={editingItem.rules?.find(r => r.id === targetRuleId)?.expression || ""}
                          onChange={(e) => {
                            if (targetRuleId) {
                              handleUpdateRuleField(targetRuleId, 'expression', e.target.value);
                            }
                          }}
                          placeholder="公式表达式为空，请点击上面的参数卡片，或直接在此处输入/编辑并微调公式..."
                          className="w-full font-mono font-bold text-xs bg-white text-indigo-900 px-3 py-2 border-2 border-indigo-200 rounded-lg outline-none focus:border-indigo-500 focus:ring-2 focus:ring-indigo-100 transition-all placeholder:text-slate-400 placeholder:font-normal resize-y min-h-[260px] overflow-y-auto"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex items-start gap-1.5 text-[10px] bg-amber-50/70 border border-amber-200 text-amber-800 p-2.5 rounded-lg leading-relaxed select-none">
                    <span className="font-bold shrink-0 mt-0.5">💡 单位自动折算优势:</span>
                    <span>检测到算料规则的计量常在“米 (m)”中执行，对于 <b>型材龙骨类 (profile)</b> 的长度相关字段（下料尺寸、L1、L2），以及<b>提取下料宽高 (W/H) </b>的尺寸数值，不管是作为<b>固化参数</b>插入还是<b>动态逻辑算料 (ROW_VAL)</b>，系统检测到毫米 (mm) 单位后都会<b>自动除以 1000 转换为米</b>计算，无需额外手动配置额外换算。</span>
                  </div>
                </div>
              )}

            </div>

            {/* Actions Footer */}
            <div className="p-3 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 shrink-0">
              <button
                type="button"
                onClick={() => setParamModalOpen(false)}
                className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded text-xs font-bold transition-all shadow-md cursor-pointer"
              >
                完成插入并关闭窗口
              </button>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}
