import React, { useState, useMemo } from "react";
import { AuxiliaryLedgerItem, ProcessedItem, COMConnectionState } from "../types";
import { evaluateFormula, simplifyFormulaDisplay } from "../utils/calcEngine";
import { 
  Calculator, Sparkles, CheckCircle2, ChevronRight, HelpCircle, 
  ChevronUp, ChevronDown, Radio, BarChart3, CloudLightning, 
  FileSpreadsheet, ClipboardCheck, Search, Database, Code, 
  Info, Layers, Minimize2, Maximize2
} from "lucide-react";
import * as XLSX from "xlsx";

interface AuxCalculationResultProps {
  ledger: AuxiliaryLedgerItem[];
  items: ProcessedItem[]; // Already parsed by range
  onExportExcel: (calculatedItems: any[]) => void;
  syncStatus: 'idle' | 'linking' | 'sending' | 'success';
  syncProgress?: number;
  onTriggerSync: (items: any[]) => void;
  sheetName?: string;
  projectVars?: Record<string, any>;
  excelSheets?: string[];
  onSelectSheet?: (sheet: string) => void;
  rawRows?: string[][];
}

const formatDecimals = (val: string) => {
  const num = parseFloat(val);
  return isNaN(num) ? val : Number(num.toFixed(3)).toString();
};

const formatMathExpr = (expr: string) => {
  return expr.replace(/[\d.]+/g, (match) => {
    if (!match.includes('.')) return match;
    const num = parseFloat(match);
    if (isNaN(num)) return match;
    if (num > 0 && num < 0.001) return match;
    return Number(num.toFixed(3)).toString();
  });
};

const parseVariablesFromLogs = (logs: string[]) => {
  const vars: Array<{ name: string; formula: string; value: string; detail: string; type: string }> = [];
  let mathEval: { original: string; evaluated: string; result: string } | null = null;

  logs.forEach(log => {
    // 1. SUM_LEN
    const sumLenMatch = log.match(/SUM_LEN\("(.*?)"\).*?总长度:\s*([\d.]+)\s*m/i);
    if (sumLenMatch) {
      vars.push({
        name: "型材长",
        formula: `SUM_LEN('${sumLenMatch[1]}')`,
        value: `${formatDecimals(sumLenMatch[2])} 米`,
        detail: log.split(' | ')[0] || log,
        type: 'SUM_LEN'
      });
      return;
    }
    // 2. SUM_QTY
    const sumQtyMatch = log.match(/SUM_QTY\("(.*?)"\).*?总数量:\s*([\d.]+)/i);
    if (sumQtyMatch) {
      vars.push({
        name: "总数量",
        formula: `SUM_QTY('${sumQtyMatch[1]}')`,
        value: `${formatDecimals(sumQtyMatch[2])} 个`,
        detail: log.split(' | ')[0] || log,
        type: 'SUM_QTY'
      });
      return;
    }
    // 3. SUM_PERIMETER
    const sumPerimMatch = log.match(/SUM_PERIMETER\("(.*?)"\).*?总周长:\s*([\d.]+)\s*m/i);
    if (sumPerimMatch) {
      vars.push({
        name: "面板周长",
        formula: `SUM_PERIMETER('${sumPerimMatch[1]}')`,
        value: `${formatDecimals(sumPerimMatch[2])} 米`,
        detail: log.split(' | ')[0] || log,
        type: 'SUM_PERIMETER'
      });
      return;
    }
    // 4. SUM_AREA
    const sumAreaMatch = log.match(/SUM_AREA\("(.*?)"\).*?总面积:\s*([\d.]+)\s*㎡/i);
    if (sumAreaMatch) {
      vars.push({
        name: "面板面积",
        formula: `SUM_AREA('${sumAreaMatch[1]}')`,
        value: `${formatDecimals(sumAreaMatch[2])} ㎡`,
        detail: log.split(' | ')[0] || log,
        type: 'SUM_AREA'
      });
      return;
    }
    // 5. ROW_VAL
    const rowValMatch = log.match(/ROW_VAL\("(.*?)",\s*"(.*?)",\s*"(.*?)",\s*"(.*?)"\).*?累加汇总(?:得|体积):\s*([-+]?[\d.]+)/i);
    if (rowValMatch) {
      const catFriendly = rowValMatch[1] === 'profile' ? '型材' : rowValMatch[1] === 'panel' ? '面板' : '钢件';
      const fieldFriendly = rowValMatch[4] === 'cuttingSize' || rowValMatch[4] === 'cuttingWidth' || rowValMatch[4] === 'cuttingHeight' ? '尺寸' : rowValMatch[4];
      vars.push({
        name: `${catFriendly}物理量`,
        formula: `ROW_VAL('${rowValMatch[1]}', '${rowValMatch[2]}', '${rowValMatch[3]}', '${rowValMatch[4]}')`,
        value: `${formatDecimals(rowValMatch[5])}`,
        detail: `匹配「${rowValMatch[3]}」的${fieldFriendly}`,
        type: 'ROW_VAL'
      });
      return;
    }

    // 6. Project Parameters
    const projectVarMatch = log.match(/提取工程参量参数 '(.+?)' = ([-+]?[\d.]+)/i);
    if (projectVarMatch) {
      vars.push({
        name: `常量参数 ${projectVarMatch[1]}`,
        formula: String(projectVarMatch[1]),
        value: `${formatDecimals(projectVarMatch[2])}`,
        detail: `全局工程参数变量提取`,
        type: 'VAR_PARAM'
      });
      return;
    }

    // Cell References
    const cellRefMatch = log.match(/提取引用单元格 \[(.+?)\] 的数据 = ([-+]?[\d.]+)/i);
    if (cellRefMatch) {
      vars.push({
        name: `单元格引用 ${cellRefMatch[1]}`,
        formula: String(cellRefMatch[1]),
        value: `${formatDecimals(cellRefMatch[2])}`,
        detail: `从当前工作表提取单元格值`,
        type: 'CELL_REF'
      });
      return;
    }
    
    const cellRefFailMatch = log.match(/提取引用单元格 \[(.+?)\] 失败/i);
    if (cellRefFailMatch) {
      vars.push({
        name: `单元格引用 ${cellRefFailMatch[1]}`,
        formula: String(cellRefFailMatch[1]),
        value: `失败 (0)`,
        detail: `工作表数据未加载或行列越界`,
        type: 'CELL_REF'
      });
      return;
    }

    // Parse the final mathematical statement
    const mathMatch = log.match(/解析表达式:\s*(.*?)\s*->\s*计算表达式:\s*(.*?)\s*->\s*结果:\s*([-+]?[\d.]+)/i);
    if (mathMatch) {
      mathEval = {
        original: mathMatch[1],
        evaluated: formatMathExpr(mathMatch[2]),
        result: formatDecimals(mathMatch[3])
      };
    }
  });

  return { vars, mathEval };
};

export default function AuxCalculationResult({
  ledger,
  items,
  onExportExcel,
  syncStatus,
  syncProgress = 0,
  onTriggerSync,
  sheetName = "",
  projectVars = {},
  excelSheets = [],
  onSelectSheet,
  rawRows,
}: AuxCalculationResultProps) {
  const [activeTab, setActiveTab] = useState<'all' | 'gasket' | 'fastener' | 'auxiliary'>('all');
  const [expandedRows, setExpandedRows] = useState<Record<string, boolean>>({});

  // Real-time floating panel selector and visibility states
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [hasInitializedSelection, setHasInitializedSelection] = useState(false);
  const [panelOpen, setPanelOpen] = useState(false);
  const [panelTab, setPanelTab] = useState<'rules' | 'variables' | 'logs'>('rules');

  // Seek the second table (usually Steel, 'steel') to compute unitPanelsCount (樘数 = 总计 / 单樘用量)
  const unitPanelsInfo = useMemo(() => {
    const k3Value = projectVars?.panelsCount !== undefined ? projectVars.panelsCount : 0;
    if (k3Value > 0) {
      return {
        count: k3Value,
        sourceName: "（未检测到有效数据，采用固定参数导入项）",
        sourceQty: k3Value,
        sourceUsage: 1,
        found: true,
        tableLabel: "固定参数导入项"
      };
    }
    return {
      count: 1, // Default to 1 if no valid K3 value
      sourceName: "（未检测到有效数据，采用默认演示系数）",
      sourceQty: 1,
      sourceUsage: 1,
      found: false,
      tableLabel: "（默认常数）"
    };
  }, [items, projectVars]);

  const unitPanelsCount = unitPanelsInfo.count;

  // Core Math Calculation: Combine Ledger + Rules + Range Parsed Items!
  const calculatedLedger = useMemo(() => {
    return ledger.map(item => {
      const itemRules = item.rules || [];
      const activeRules = itemRules.filter(r => r.isActive);

      if (activeRules.length > 0) {
        let calculatedUnitUsage = 0;
        const totalLogs: string[] = [];

        // Pre-evaluate all active rules to check for exclusive results
        const evaluatedRules = activeRules.map(rule => {
          const { value, logs } = evaluateFormula(rule.expression, items, projectVars, rule.lengthAdjustment || 0, rawRows);
          return { rule, value, logs };
        });

        const matchedExclusiveRule = evaluatedRules.find(er => er.rule.isExclusive && er.value !== 0);

        evaluatedRules.forEach(({ rule, value, logs }) => {
          const simplifiedExpr = simplifyFormulaDisplay(rule.expression);
          const simplifiedLogs = logs.map(log => simplifyFormulaDisplay(log));

          const isIgnored = !!matchedExclusiveRule && matchedExclusiveRule.rule.id !== rule.id;

          if (isIgnored) {
            totalLogs.push(`【配量单樘公式: ${rule.name} (表达式: ${simplifiedExpr})】: 由于排它规则[${matchedExclusiveRule.rule.name}]计算结果不为0，此规则已被忽略不予计算。`);
          } else {
            calculatedUnitUsage += value;
            totalLogs.push(`【配量单樘公式: ${rule.name} (表达式: ${simplifiedExpr})】${rule.isExclusive ? ' [排它] ' : ''}: 计算单樘用量 = ${value}。 (源日志: ${simplifiedLogs.join(" | ")})`);
          }
        });

        // Apply precision rules based on category and unit
        let finalUnitUsage = calculatedUnitUsage;
        let finalQty = calculatedUnitUsage * unitPanelsCount;
        
        const COUNT_UNITS = ['个', '支', '块', '只', '套'];
        const cleanUnit = item.unit ? item.unit.trim() : '';
        const isCountUnit = COUNT_UNITS.includes(cleanUnit);

        if (isCountUnit) {
          finalUnitUsage = Math.ceil(finalUnitUsage);
          finalQty = Math.ceil(finalQty);
        } else if (item.category === 'fastener') {
          finalUnitUsage = Math.round(finalUnitUsage);
          finalQty = Math.round(finalQty);
        } else {
          finalUnitUsage = Math.round(finalUnitUsage * 150) / 150; // Keep compatible precision
          finalUnitUsage = Math.round(finalUnitUsage * 100) / 100;
          finalQty = Math.round(finalQty * 100) / 100;
        }

        const decimalPlaces = (item.category === 'fastener' || isCountUnit) ? 0 : 2;
        totalLogs.push(`【总计折算核算】: 总计 = 单樘用量 (${finalUnitUsage.toFixed(decimalPlaces)}) * 单元板块樘数 (${unitPanelsCount}) = ${finalQty.toFixed(decimalPlaces)}。`);

        const usedRules = matchedExclusiveRule ? [matchedExclusiveRule.rule] : activeRules;

        return {
          ...item,
          isCalculated: true,
          calculatedUnitUsage: finalUnitUsage, 
          calculatedQty: finalQty,       
          formula: usedRules.map(r => r.expression).join(" + "),
          matchLog: totalLogs,
          evaluatedRulesDetails: evaluatedRules.map(({ rule, value, logs }) => {
            const isIgnored = !!matchedExclusiveRule && matchedExclusiveRule.rule.id !== rule.id;
            return {
              name: rule.name,
              expression: rule.expression,
              isExclusive: !!rule.isExclusive,
              isIgnored,
              value: isIgnored ? 0 : value,
              logs
            };
          })
        };
      }

      return {
        ...item,
        isCalculated: false,
        calculatedUnitUsage: 0,
        calculatedQty: 0,
        formula: "",
        matchLog: [`[提示] 该辅材品类目前未配置/未启用任何单樘换算计算规则。单樘用量及总计已自动置空。`],
        evaluatedRulesDetails: []
      };
    });
  }, [ledger, items, unitPanelsCount]);

  React.useEffect(() => {
    if (calculatedLedger.length > 0 && !hasInitializedSelection) {
      setSelectedIds(new Set(calculatedLedger.map(i => i.id)));
      setHasInitializedSelection(true);
    }
  }, [calculatedLedger, hasInitializedSelection]);

  const selectedItems = useMemo(() => {
    return calculatedLedger.filter(item => selectedIds.has(item.id));
  }, [calculatedLedger, selectedIds]);

  const aggregatedData = useMemo(() => {
    const rules: Array<{ itemName: string; ruleName: string; expression: string; isExclusive: boolean; isIgnored: boolean; value: number }> = [];
    const allLogs: string[] = [];
    
    selectedItems.forEach(item => {
      if (item.evaluatedRulesDetails) {
        item.evaluatedRulesDetails.forEach((rd: any) => {
          rules.push({
            itemName: item.name,
            ruleName: rd.name,
            expression: rd.expression,
            isExclusive: !!rd.isExclusive,
            isIgnored: !!rd.isIgnored,
            value: rd.value
          });
        });
      }
      
      const matchLogs = item.matchLog || [];
      matchLogs.forEach(log => {
        if (!log.includes('总计折算核算') && !log.includes('未启用任何单樘换算')) {
          allLogs.push(`[${item.name}] ${log}`);
        }
      });
    });

    const { vars } = parseVariablesFromLogs(allLogs);

    // Deduplicate variables
    const uniqueVars: typeof vars = [];
    const seenFormulas = new Set<string>();
    vars.forEach(v => {
      const key = `${v.formula}::${v.value}`;
      if (!seenFormulas.has(key)) {
        seenFormulas.add(key);
        uniqueVars.push(v);
      }
    });

    return {
      rules,
      variables: uniqueVars,
      allLogs,
    };
  }, [selectedItems]);

  const toggleExpandRow = (id: string) => {
    setExpandedRows(prev => ({ ...prev, [id]: !prev[id] }));
  };

  const visibleItems = useMemo(() => {
    if (activeTab === 'all') return calculatedLedger;
    return calculatedLedger.filter(i => i.category === activeTab);
  }, [calculatedLedger, activeTab]);

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="module_calc_totals">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4 border-b border-slate-100 pb-5">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Calculator className="w-5 h-5 text-indigo-600" />
            辅材物料大底算料汇总及直达填回 (Final Quant Summary)
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            系统读取第二模块区间分割的原材料算量基准，映射第三模块注册的辅具台账目，最终在第五模块内完成高精度统筹和逻辑闭环。
          </p>
        </div>

        <div className="flex flex-col gap-2 w-full md:w-auto">
          <div className="flex flex-wrap gap-2">
            <button
              onClick={() => onExportExcel(calculatedLedger)}
              disabled={calculatedLedger.length === 0}
              className="flex-1 md:flex-none justify-center bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-4 py-2.5 rounded transition-all text-xs flex items-center gap-1.5 shadow-sm cursor-pointer disabled:opacity-40"
            >
              <FileSpreadsheet className="w-4 h-4" />
              一键写回原表导出 Excel
            </button>

            <button
              onClick={() => onTriggerSync(calculatedLedger)}
              disabled={calculatedLedger.length === 0 || syncStatus === 'linking' || syncStatus === 'sending'}
              className="flex-1 md:flex-none relative overflow-hidden justify-center bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2.5 rounded transition-all text-xs flex items-center gap-1.5 shadow-xs cursor-pointer disabled:opacity-40"
            >
              {syncStatus === 'sending' && (
                <div 
                  className="absolute left-0 top-0 bottom-0 bg-emerald-500 transition-all duration-300 z-0"
                  style={{ width: `${syncProgress}%` }}
                ></div>
              )}
              <span className="relative z-10 flex items-center gap-1.5">
                <Radio className={`w-4 h-4 ${syncStatus !== "idle" && syncStatus !== "success" ? "animate-pulse text-emerald-100" : ""}`} />
                {syncStatus === 'sending' ? `COM 填回中 ${syncProgress}%` : '一键 COM 填回'}
              </span>
            </button>
          </div>
        </div>
      </div>

      {/* Dynamic Unit Panels Count Banner */}
      <div className="mb-6 p-4 rounded-xl border bg-indigo-50/40 border-indigo-100 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
        <div className="space-y-1">
          <div className="flex items-center gap-1.5 text-indigo-900 font-bold text-sm">
            <Sparkles className="w-4 h-4 text-indigo-600 animate-pulse" />
            <span>智能单元板块大样标称樘数计算中心</span>
          </div>
          <div className="text-slate-600 leading-relaxed">
            依据<b>工程固定参数</b>静态直取：
            从坐标 K3 获取 -&gt;
            单元板块樘数 <span className="font-mono text-emerald-700 font-bold text-xs bg-emerald-50 px-1.5 py-0.5 rounded border border-emerald-100">{unitPanelsCount} 樘</span>
          </div>
        </div>
        <div className="shrink-0 bg-white border border-indigo-100 px-3.5 py-1.5 rounded-lg text-center shadow-2xs">
          <span className="text-[9px] text-slate-400 block font-bold uppercase tracking-wider">单元物理板块樘数</span>
          <div className="flex items-baseline justify-center gap-2 mt-1">
            <span className="text-xl font-black text-indigo-900 tracking-tight flex items-center bg-slate-50 hover:bg-slate-100 rounded border border-slate-200 px-1.5 py-0.5 cursor-pointer">
              <select
                value={sheetName}
                onChange={(e) => onSelectSheet?.(e.target.value)}
                className="bg-transparent text-sm font-black text-indigo-900 tracking-tight border-none p-0 focus:ring-0 focus:outline-none cursor-pointer max-w-[150px] truncate"
              >
                {excelSheets && excelSheets.length > 0 ? (
                  excelSheets.map((sh, idx) => (
                    <option key={idx} value={sh} className="text-xs font-sans text-slate-800 bg-white">
                      {sh}
                    </option>
                  ))
                ) : (
                  <option value={sheetName} className="text-xs font-sans text-slate-800 bg-white">
                    {sheetName || "工作表"}
                  </option>
                )}
              </select>
            </span>
            <span className="text-sm font-bold text-slate-400">=</span>
            <span className="text-lg font-extrabold font-mono text-indigo-700">{unitPanelsCount} <span className="text-xs font-sans text-slate-500 font-normal">樘</span></span>
          </div>
        </div>
      </div>

      {/* Filter and selector tab panel info */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-4.5 gap-3">
        <div className="flex gap-1 bg-slate-100 p-0.5 rounded border border-slate-200">
          <button
            onClick={() => setActiveTab('all')}
            className={`px-3 py-1.5 rounded text-xs font-bold transition-colors cursor-pointer ${
              activeTab === 'all' ? "bg-white text-slate-900 shadow-xs border border-slate-200/50 font-bold" : "text-slate-500 hover:text-slate-800"
            }`}
          >
            全部大类 ({calculatedLedger.length})
          </button>
          <button
            onClick={() => setActiveTab('gasket')}
            className={`px-3 py-1.5 rounded text-xs font-bold transition-colors cursor-pointer ${
              activeTab === 'gasket' ? "bg-white text-emerald-700 shadow-xs border border-slate-200/50 font-bold" : "text-slate-500 hover:text-emerald-700"
            }`}
          >
            气密胶条 ({calculatedLedger.filter(i => i.category === 'gasket').length})
          </button>
          <button
            onClick={() => setActiveTab('fastener')}
            className={`px-3 py-1.5 rounded text-xs font-bold transition-colors cursor-pointer ${
              activeTab === 'fastener' ? "bg-white text-amber-700 shadow-xs border border-slate-200/50 font-bold" : "text-slate-500 hover:text-amber-700"
            }`}
          >
            五金紧固件 ({calculatedLedger.filter(i => i.category === 'fastener').length})
          </button>
          <button
            onClick={() => setActiveTab('auxiliary')}
            className={`px-3 py-1.5 rounded text-xs font-bold transition-colors cursor-pointer ${
              activeTab === 'auxiliary' ? "bg-white text-indigo-700 shadow-xs border border-slate-200/50 font-bold" : "text-slate-500 hover:text-indigo-700"
            }`}
          >
            胶水辅材 ({calculatedLedger.filter(i => i.category === 'auxiliary').length})
          </button>
        </div>

        <div className="text-[11px] text-slate-500 font-semibold bg-slate-50 border px-2.5 py-1 rounded">
          当前待重算源行数: <span className="text-indigo-600 font-mono font-bold">{items.length} 组下料行</span>
        </div>
      </div>

      {/* Aggregate Output Table */}
      <div className="border border-slate-200 rounded overflow-hidden shadow-xs overflow-x-auto">
        <table className="min-w-full divide-y divide-slate-200 text-[11px]">
          <thead className="bg-slate-50 text-slate-550 font-bold tracking-wider text-[10px] uppercase">
            <tr>
              <th className="px-3 py-3 text-center w-10 border-b border-slate-200">
                <input
                  type="checkbox"
                  checked={visibleItems.length > 0 && visibleItems.every(item => selectedIds.has(item.id))}
                  onChange={() => {
                    const allSelected = visibleItems.every(item => selectedIds.has(item.id));
                    const next = new Set(selectedIds);
                    if (allSelected) {
                      visibleItems.forEach(item => next.delete(item.id));
                    } else {
                      visibleItems.forEach(item => next.add(item.id));
                    }
                    setSelectedIds(next);
                  }}
                  className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer"
                  title="全选/全不选当前页"
                />
              </th>
              <th className="px-3 py-3 text-left w-12 border-b border-slate-200">序号</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">名称</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">加工图号</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">材质及型号</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">下料尺寸 (L) (MM)</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">单樘用量</th>
              <th className="px-4 py-3 text-right font-bold text-indigo-700 bg-indigo-50/15 border-b border-slate-200">总计</th>
              <th className="px-3 py-3 text-left w-12 border-b border-slate-200">单位</th>
              <th className="px-3 py-3 text-left w-12 border-b border-slate-200">L1</th>
              <th className="px-3 py-3 text-left w-12 border-b border-slate-200">L2</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">备注</th>
              <th className="px-3 py-3 text-left border-b border-slate-200">算料逻辑归属</th>
              <th className="px-3 py-3 text-center w-20 border-b border-slate-200">核算审计链</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-slate-100 text-slate-700">
            {visibleItems.length === 0 ? (
              <tr>
                <td colSpan={14} className="px-4 py-12 text-center text-slate-400 italic">
                  此类别暂无建立辅材台账目，请在第三模块注册辅材！
                </td>
              </tr>
            ) : (
              visibleItems.map((item, idx) => {
                const isExpanded = !!expandedRows[item.id];
                
                const getFriendlyCategory = (cat: string) => {
                  if (cat === 'gasket') return '型材气密胶条类';
                  if (cat === 'fastener') return '不锈钢紧固五金螺栓类';
                  if (cat === 'auxiliary') return '结构密封胶及辅助垫块类';
                  return '其它辅材类';
                };

                return (
                  <React.Fragment key={item.id}>
                    <tr className={`hover:bg-indigo-50/15 transition-colors ${item.isCalculated ? "bg-indigo-50/5" : ""} ${selectedIds.has(item.id) ? "bg-indigo-50/10 font-medium" : "opacity-85"}`}>
                      <td className="px-3 py-3 text-center">
                        <input
                          type="checkbox"
                          checked={selectedIds.has(item.id)}
                          onChange={() => {
                            const next = new Set(selectedIds);
                            if (next.has(item.id)) {
                              next.delete(item.id);
                            } else {
                              next.add(item.id);
                            }
                            setSelectedIds(next);
                          }}
                          className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer"
                        />
                      </td>
                      <td className="px-3 py-3 font-mono text-slate-400">{idx + 1}</td>
                      <td className="px-3 py-3 font-bold text-slate-900">{item.name}</td>
                      <td className="px-3 py-3 font-mono text-indigo-600 font-semibold">{item.drawingNo}</td>
                      <td className="px-3 py-3 text-slate-600 font-medium">
                        {item.materialModel}
                      </td>
                      <td className="px-3 py-3 font-mono text-slate-500">
                        {item.size || "—"}
                      </td>
                      <td className="px-3 py-3 text-slate-900 font-mono font-bold">
                        {item.isCalculated ? item.calculatedUnitUsage?.toFixed((item.category === 'fastener' || (item.unit && (item.unit.trim() === '个' || item.unit.trim() === '支'))) ? 0 : 2) : ((item.category === 'fastener' || (item.unit && (item.unit.trim() === '个' || item.unit.trim() === '支'))) ? "0" : "0.00")}
                      </td>
                      <td className="px-4 py-3 text-right bg-rose-50/15 font-bold text-xs text-rose-700 font-mono">
                        {item.isCalculated ? item.calculatedQty?.toFixed((item.category === 'fastener' || (item.unit && (item.unit.trim() === '个' || item.unit.trim() === '支'))) ? 0 : 2) : ((item.category === 'fastener' || (item.unit && (item.unit.trim() === '个' || item.unit.trim() === '支'))) ? "0" : "0.00")}
                      </td>
                      <td className="px-3 py-3 font-bold text-slate-500">{item.unit}</td>
                      <td className="px-3 py-3 text-slate-405 font-mono text-slate-400">—</td>
                      <td className="px-3 py-3 text-slate-405 font-mono text-slate-400">—</td>
                      <td className="px-3 py-3 text-slate-500 max-w-[150px] truncate" title={item.remark}>
                        {item.remark || "—"}
                      </td>
                      <td className="px-3 py-3 text-slate-600 font-semibold">
                        <span className="bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-bold">
                          {getFriendlyCategory(item.category)}
                        </span>
                      </td>
                      <td className="px-3 py-3 text-center">
                        <button
                          onClick={() => toggleExpandRow(item.id)}
                          className={`text-[10px] px-2 py-1 rounded border transition-all flex items-center gap-1 mx-auto ${
                            isExpanded
                              ? "bg-indigo-100 border-indigo-200 text-indigo-700 font-bold"
                              : "bg-white border-slate-200 text-slate-500 hover:bg-slate-50 cursor-pointer font-semibold"
                          }`}
                        >
                          <span>溯源</span>
                          {isExpanded ? <ChevronUp className="w-3 h-3 text-indigo-500" /> : <ChevronDown className="w-3 h-3 text-slate-400" />}
                        </button>
                      </td>
                    </tr>

                     {isExpanded && (
                      <tr className="bg-slate-50/50">
                        <td colSpan={14} className="px-8 py-4 w-full border-t border-b border-indigo-150 bg-indigo-50/5">
                          <div className="space-y-4 max-w-6xl mx-auto text-left">
                            {/* Title Banner */}
                            <div className="flex items-center justify-between border-b border-slate-200 pb-2">
                              <h4 className="font-bold text-xs text-indigo-950 tracking-wide flex items-center gap-1.5">
                                <Sparkles className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                                辅材「{item.name}」精细算料逻辑与公式因子审计明细 (Formula Factor Analysis)
                              </h4>
                              <div className="text-[10px] text-slate-500 font-semibold bg-slate-100 px-2 py-0.5 rounded flex items-baseline gap-1">
                                单元物理型材板块数: <span className="font-bold text-indigo-700 text-[11px] uppercase ml-0.5 mr-0.5">{sheetName || "工作表"}</span> <span className="text-slate-400">=</span> <span className="font-bold text-indigo-700 font-mono">{unitPanelsCount} 樘</span>
                              </div>
                            </div>

                            {/* Two Column Grid */}
                            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                              {/* Left Board: Applied Formulas */}
                              <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs space-y-3">
                                <div className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest flex items-center gap-1.5">
                                  <FileSpreadsheet className="w-3.5 h-3.5 text-indigo-600" />
                                  1. 调用的算料预设公式 (Applied Formulas)
                                </div>

                                {item.evaluatedRulesDetails && item.evaluatedRulesDetails.length > 0 ? (
                                  <div className="space-y-2.5">
                                    {item.evaluatedRulesDetails.map((ruleDetail: any, rIdx: number) => (
                                      <div key={rIdx} className={`p-3 rounded-lg border text-[11px] space-y-2 transition-all ${
                                        ruleDetail.isIgnored 
                                          ? "bg-slate-50 border-slate-200 opacity-60" 
                                          : "bg-indigo-50/20 border-indigo-100"
                                      }`}>
                                        <div className="flex items-center justify-between">
                                          <div className="font-bold text-slate-800 flex items-center gap-1.5">
                                            <span className={`w-1.5 h-1.5 rounded-full ${ruleDetail.isIgnored ? "bg-slate-450" : "bg-indigo-500"}`} />
                                            {ruleDetail.name}
                                          </div>
                                          <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold ${
                                            ruleDetail.isIgnored 
                                              ? "bg-slate-200 text-slate-650" 
                                              : ruleDetail.isExclusive 
                                                ? "bg-amber-100 text-amber-800 border border-amber-200 animate-pulse" 
                                                : "bg-indigo-100 text-indigo-800"
                                          }`}>
                                            {ruleDetail.isIgnored ? "由于排它公式生效此项已忽略" : ruleDetail.isExclusive ? "排它独占模式" : "普通累加模式"}
                                          </span>
                                        </div>

                                        <div className="space-y-1 text-slate-600 font-mono">
                                          <div className="flex items-start gap-1">
                                            <span className="text-slate-400 font-sans shrink-0">设计原型算式:</span>
                                            <span className="font-semibold text-slate-900 break-all bg-slate-50 px-1 py-0.5 rounded text-[10px]">{ruleDetail.expression}</span>
                                          </div>
                                          <div className="flex items-start gap-1">
                                            <span className="text-slate-400 font-sans shrink-0">简化易读算式:</span>
                                            <span className="text-slate-700 bg-white border border-slate-200/80 px-1 py-0.5 rounded text-[10px] font-sans font-medium break-all">{simplifyFormulaDisplay(ruleDetail.expression)}</span>
                                          </div>
                                        </div>

                                        {!ruleDetail.isIgnored && (
                                          <div className="flex justify-between items-center bg-white/80 px-2.5 py-1.5 rounded border border-indigo-100 font-bold text-indigo-950 font-sans mt-2">
                                            <span>贡献单樘用量:</span>
                                            <span className="font-mono text-indigo-700 text-xs">
                                              +{ruleDetail.value.toFixed(item.category === 'fastener' ? 0 : 2)} <span className="text-[10px] font-sans text-slate-500 font-normal">{item.unit}</span>
                                            </span>
                                          </div>
                                        )}
                                      </div>
                                    ))}
                                  </div>
                                ) : (
                                  <div className="p-4 text-center text-slate-400 italic border border-dashed border-slate-200 rounded-lg bg-slate-50/50">
                                    此条目未启用或未匹配任何计算公式。
                                  </div>
                                )}
                              </div>

                              {/* Right Board: Factors Explained */}
                              <div className="bg-white p-3.5 rounded-xl border border-slate-200 shadow-2xs space-y-3">
                                <div className="text-[10px] font-extrabold text-slate-500 uppercase tracking-widest flex items-center gap-1.5">
                                  <ClipboardCheck className="w-3.5 h-3.5 text-emerald-600" />
                                  2. 算式变量的具体解析值 (Variables Resolved)
                                </div>

                                {item.evaluatedRulesDetails && item.evaluatedRulesDetails.length > 0 && (
                                  <div className="space-y-3">
                                    {(() => {
                                      const allLogs = item.evaluatedRulesDetails
                                        .filter((rd: any) => !rd.isIgnored)
                                        .flatMap((rd: any) => rd.logs || []);

                                      const { vars, mathEval } = parseVariablesFromLogs(allLogs);

                                      return (
                                        <div className="space-y-3">
                                          {vars.length > 0 ? (
                                            <div className="space-y-1.5 max-h-[160px] overflow-y-auto pr-1">
                                              {vars.map((v, vIdx) => (
                                                <div key={vIdx} className="flex flex-col sm:flex-row sm:items-center justify-between p-2.5 bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-lg text-[11px] gap-2">
                                                  <div className="space-y-1">
                                                    <div className="flex items-center gap-1">
                                                      <span className="bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold px-1.5 py-0.5 rounded text-[9px] uppercase tracking-wide">
                                                        {v.name}
                                                      </span>
                                                      <code className="font-mono text-slate-900 font-bold">{v.formula}</code>
                                                    </div>
                                                    <div className="text-[10px] text-slate-500 font-sans">{v.detail}</div>
                                                  </div>
                                                  <div className="text-right sm:border-l sm:border-slate-200 sm:pl-3 shrink-0 flex flex-col justify-center">
                                                    <div className="font-extrabold font-mono text-emerald-700 text-xs">{v.value}</div>
                                                  </div>
                                                </div>
                                              ))}
                                            </div>
                                          ) : (
                                            <div className="p-4 text-xs text-slate-400 italic text-center border border-dashed border-slate-200 rounded-lg bg-slate-50/50">
                                              未检测到公式内部动态搜索变量，该公式通过纯常数直接进行评估。
                                            </div>
                                          )}

                                          {/* Algebraic substitution math walkthrough */}
                                          {mathEval && (
                                            <div className="bg-amber-50/40 border border-amber-100 p-3 rounded-lg text-[11px] space-y-2 font-mono text-slate-700">
                                              <div className="text-[10px] font-sans font-extrabold text-amber-900 flex items-center gap-1">
                                                <span>★ 算术代数求值过程 (Algebraic Substitution Pathway)</span>
                                              </div>
                                              <div className="space-y-1.5 bg-white/60 p-2 rounded border border-amber-200/50">
                                                <div className="flex justify-between border-b border-dashed border-slate-200/80 pb-1">
                                                  <span className="text-slate-450 font-sans">设计公式</span>
                                                  <span className="font-semibold text-slate-800">{simplifyFormulaDisplay(mathEval.original)}</span>
                                                </div>
                                                <div className="flex justify-between border-b border-dashed border-slate-200/80 pb-1">
                                                  <span className="text-slate-450 font-sans">因式代入 (Values Replaced)</span>
                                                  <span className="text-amber-800 font-bold">{mathEval.evaluated}</span>
                                                </div>
                                                <div className="flex justify-between pt-0.5">
                                                  <span className="text-slate-450 font-sans">单樘运算贡献值 (Calculated Sum)</span>
                                                  <span className="font-extrabold text-rose-700 text-xs">{Number(mathEval.result).toFixed(3)}</span>
                                                </div>
                                              </div>
                                            </div>
                                          )}
                                        </div>
                                      );
                                    })()}
                                  </div>
                                )}
                              </div>
                            </div>

                            {/* Raw Evaluation Logs */}
                            <div className="space-y-1.5">
                              <div className="text-[9px] font-extrabold text-slate-400 uppercase tracking-widest">
                                完整溯源日志审计 (Full Raw Calculation Logs)
                              </div>
                              <div className="bg-slate-900 text-slate-200 p-3.5 rounded-lg border border-slate-800 font-mono text-[10.5px] space-y-1 max-h-[100px] overflow-y-auto leading-relaxed">
                                {item.matchLog && item.matchLog.length > 0 ? (
                                  item.matchLog.map((log, lIdx) => (
                                    <div key={lIdx} className="flex gap-2">
                                      <span className="text-indigo-400 shrink-0 select-none">[{lIdx+1}]</span>
                                      <span className="text-slate-300 word-break-all text-left">{log}</span>
                                    </div>
                                  ))
                                ) : (
                                  <div className="text-slate-500 italic">正在查询料单物理特征参数中...</div>
                                )}
                              </div>
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Floating Minimize Trigger / Action Board */}
      {!panelOpen ? (
        <button
          onClick={() => setPanelOpen(true)}
          className="fixed bottom-6 right-6 z-50 bg-indigo-600 hover:bg-indigo-700 text-white font-bold p-3 px-4.5 rounded-full shadow-2xl flex items-center gap-2 hover:scale-105 active:scale-95 transition-all cursor-pointer border border-indigo-400/30 animate-bounce"
        >
          <Layers className="w-4 h-4" />
          <span className="text-xs tracking-wide">🔮 辅材算料因子审计透视 ({selectedIds.size}项选中)</span>
        </button>
      ) : (
        <div className="fixed bottom-6 right-6 z-50 w-[480px] max-w-[calc(100vw-2rem)] bg-white/95 backdrop-blur-md rounded-xl border border-slate-200/80 shadow-2xl flex flex-col overflow-hidden max-h-[480px]">
          {/* Header */}
          <div className="bg-indigo-950 text-white px-4 py-3 flex items-center justify-between border-b border-indigo-900 shrink-0">
            <div className="flex items-center gap-1.5">
              <Sparkles className="w-4 h-4 text-indigo-400 animate-pulse" />
              <div className="text-left">
                <h3 className="font-bold text-xs tracking-wide">全局算料因子透视仪 (Audit Panel)</h3>
                <p className="text-[9.5px] text-indigo-300">已聚合 {selectedIds.size} 项所选辅材的高精因式结构</p>
              </div>
            </div>
            <div className="flex items-center gap-1">
              <button
                onClick={() => setPanelOpen(false)}
                className="p-1 hover:bg-slate-800 rounded text-slate-300 hover:text-white transition-colors cursor-pointer"
                title="最小化面板"
              >
                <Minimize2 className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          {/* Quick Selection Toolbar inside panel */}
          <div className="bg-slate-50 border-b border-slate-200 px-3 py-1.5 flex items-center justify-between text-[10px] text-slate-500 font-medium shrink-0">
            <div className="text-left">
              已选辅材项: <span className="font-mono font-bold text-slate-800">{selectedIds.size}</span> / <span className="font-mono">{calculatedLedger.length} 项全类</span>
            </div>
            <div className="flex gap-2">
              <button 
                onClick={() => setSelectedIds(new Set(calculatedLedger.map(i => i.id)))}
                className="text-indigo-600 hover:text-indigo-800 font-bold hover:underline cursor-pointer"
              >
                全选
              </button>
              <span className="text-slate-300">|</span>
              <button 
                onClick={() => setSelectedIds(new Set())}
                className="text-slate-500 hover:text-slate-700 font-bold hover:underline cursor-pointer"
              >
                清空
              </button>
            </div>
          </div>

          {/* Tabs */}
          <div className="bg-white border-b border-slate-200 px-2 flex gap-1 shrink-0">
            <button
              onClick={() => setPanelTab('rules')}
              className={`px-3 py-2 text-[11px] font-bold border-b-2 flex items-center gap-1 cursor-pointer transition-all ${
                panelTab === 'rules' 
                  ? "border-indigo-650 text-indigo-650" 
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <Database className="w-3.5 h-3.5" />
              公式/排它规则 ({aggregatedData.rules.length})
            </button>
            <button
              onClick={() => setPanelTab('variables')}
              className={`px-3 py-2 text-[11px] font-bold border-b-2 flex items-center gap-1 cursor-pointer transition-all ${
                panelTab === 'variables' 
                  ? "border-emerald-600 text-emerald-600" 
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <Code className="w-3.5 h-3.5" />
              因数数值溯源 ({aggregatedData.variables.length})
            </button>
            <button
              onClick={() => setPanelTab('logs')}
              className={`px-3 py-2 text-[11px] font-bold border-b-2 flex items-center gap-1 cursor-pointer transition-all ${
                panelTab === 'logs' 
                  ? "border-slate-800 text-slate-800" 
                  : "border-transparent text-slate-500 hover:text-slate-800"
              }`}
            >
              <Search className="w-3.5 h-3.5" />
              算料逻辑审计流 ({aggregatedData.allLogs.length})
            </button>
          </div>

          {/* Scrollable Aggregated Content Panels */}
          <div className="flex-1 overflow-y-auto p-3 bg-slate-50/50 space-y-3">
            {selectedIds.size === 0 ? (
              <div className="py-12 px-4 text-center space-y-2">
                <Info className="w-8 h-8 text-slate-400 mx-auto" />
                <p className="text-xs text-slate-500 font-medium">请勾选主表行中的复选框，</p>
                <p className="text-[10px] text-slate-400">实时调取所选项目的高精计算规则及审计因子。</p>
              </div>
            ) : (
              <>
                {panelTab === 'rules' && (
                  <div className="space-y-2 text-left">
                    {aggregatedData.rules.map((rule, idx) => (
                      <div key={idx} className={`p-2.5 rounded-lg border text-[11px] space-y-1 bg-white ${
                        rule.isIgnored ? "border-slate-150 opacity-60 bg-slate-100/50" : "border-indigo-100 shadow-2xs"
                      }`}>
                        <div className="flex items-center justify-between">
                          <div className="font-bold text-slate-800 flex items-center gap-1">
                            <span className="text-[9.5px] bg-slate-155 text-slate-600 font-bold px-1 py-0.5 rounded mr-1 max-w-[100px] truncate">{rule.itemName}</span>
                            <span>{rule.ruleName}</span>
                          </div>
                          <span className={`text-[8.5px] px-1 py-0.2 rounded font-bold ${
                            rule.isIgnored 
                              ? "bg-slate-200 text-slate-500" 
                              : rule.isExclusive 
                                ? "bg-amber-100 text-amber-800 border border-amber-200" 
                                : "bg-indigo-100 text-indigo-800"
                          }`}>
                            {rule.isIgnored ? "由于排它公式此项已忽略" : rule.isExclusive ? "排它独占" : "普通累加"}
                          </span>
                        </div>
                        <div className="text-[10.5px] font-mono text-slate-600 space-y-0.5 bg-slate-50/60 p-1.5 rounded border border-slate-100">
                          <div className="text-slate-400 text-[8.5px]">公式物理算式表达式</div>
                          <div className="font-semibold text-indigo-950 break-all">{rule.expression}</div>
                          <div className="text-slate-500 text-[9.5px] font-sans">
                            {simplifyFormulaDisplay(rule.expression)}
                          </div>
                        </div>
                        {!rule.isIgnored && (
                          <div className="text-right text-[10px] text-indigo-700 font-bold font-mono">
                            单樘贡献: +{rule.value.toFixed(3)}
                          </div>
                        )}
                      </div>
                    ))}

                    {/* 算料规则单位换算与参量单元格匹配说明 (小型化提示) */}
                    <div className="mt-4 p-2.5 bg-slate-50 border border-slate-200 rounded-lg text-slate-600 text-[10.5px] flex items-center gap-1.5 shadow-2xs">
                      <HelpCircle className="w-4 h-4 text-indigo-500 shrink-0" />
                      <span>💡 算料公式物理参量与单位自动折算说明已置于应用页面最底部，可随时折叠展开查阅。</span>
                    </div>
                  </div>
                )}

                {panelTab === 'variables' && (
                  <div className="space-y-2 text-left">
                    {aggregatedData.variables.length === 0 ? (
                      <div className="text-center py-8 text-slate-400 italic text-[11px]">
                        未检测到包含变量公式，均为常数运算。
                      </div>
                    ) : (
                      aggregatedData.variables.map((v, idx) => (
                        <div key={idx} className="p-2.5 rounded-lg border border-slate-200 bg-white shadow-2xs text-[11px] space-y-1.5">
                          <div className="flex items-center justify-between">
                            <span className="bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold px-1.5 py-0.5 rounded text-[8.5px] uppercase tracking-wider">
                              {v.name}
                            </span>
                            <code className="font-mono text-[10px] font-bold text-slate-700">{v.formula}</code>
                          </div>
                          <div className="flex justify-between items-center bg-slate-50 p-1.5 rounded">
                            <span className="text-[10px] text-slate-500 font-sans">{v.detail}</span>
                            <span className="font-mono font-extrabold text-emerald-700 text-md">{v.value}</span>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                )}

                {panelTab === 'logs' && (
                  <div className="bg-slate-900 rounded-lg p-2.5 border border-slate-800 font-mono text-[9.5px] text-slate-200 space-y-1 max-h-[300px] overflow-y-auto">
                    {aggregatedData.allLogs.map((log, idx) => (
                      <div key={idx} className="flex items-start gap-1.5 leading-relaxed">
                        <span className="text-indigo-400 shrink-0">[{idx + 1}]</span>
                        <span className="text-slate-350 text-left break-all select-all">{log}</span>
                      </div>
                    ))}
                  </div>
                )}
              </>
            )}
          </div>

          {/* Quick Metrics Footer */}
          <div className="bg-slate-100 border-t border-slate-200/80 px-4 py-2.5 flex items-center justify-between shrink-0 text-[10px] text-slate-500">
            <span className="font-semibold text-slate-500">
              因子面板聚合计数实时监控：
            </span>
            <div className="flex gap-2 text-slate-700 font-mono">
              <span className="bg-indigo-50 px-2 py-0.5 border border-indigo-200 shadow-3xs font-sans rounded">
                计算公式: <span className="font-bold text-indigo-700">{aggregatedData.rules.filter(r => !r.isIgnored).length} 项</span>
              </span>
              <span className="bg-emerald-50 px-2 py-0.5 border border-emerald-200 shadow-3xs font-sans rounded">
                分析因子: <span className="font-bold text-emerald-700">{aggregatedData.variables.length} 个</span>
              </span>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
