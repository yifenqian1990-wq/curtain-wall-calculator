import React, { useState, useEffect } from "react";
import { ProcessedItem, FormulaRule, ColumnMappings } from "../types";
import { evaluateFormula, simplifyFormulaDisplay } from "../utils/calcEngine";
import * as XLSX from "xlsx";
import { RefreshCw, FileSpreadsheet, CheckCircle2, ChevronDown, ChevronUp, Radio, Network, Sparkles, Send } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";

interface CalculatorPreviewProps {
  originalWorkbook: XLSX.WorkBook | null;
  sheetName: string | null;
  items: ProcessedItem[];
  formulas: FormulaRule[];
  mappings: ColumnMappings;
  originalRawRows: string[][];
  projectVars?: Record<string, number>;
}

export default function CalculatorPreview({
  originalWorkbook,
  sheetName,
  items,
  formulas,
  mappings,
  originalRawRows,
  projectVars = {}
}: CalculatorPreviewProps) {
  const [calculatedItems, setCalculatedItems] = useState<ProcessedItem[]>([]);
  const [activeTab, setActiveTab] = useState<'all' | 'gasket' | 'fastener' | 'auxiliary'>('all');
  const [expandedRowLogs, setExpandedRowLogs] = useState<Record<number, boolean>>({});
  const [isExporting, setIsExporting] = useState(false);
  const [syncStatus, setSyncStatus] = useState<'idle' | 'linking' | 'sending' | 'success'>('idle');
  const [syncLogs, setSyncLogs] = useState<string[]>([]);

  // Trigger calculations whenever items or formulas change
  useEffect(() => {
    const computed = items.map(item => {
      // Find matching formula rule
      if (item.category === 'gasket' || item.category === 'fastener' || item.category === 'auxiliary') {
        const matchingRule = formulas.find(rule => {
          const catMatches = rule.targetCategory === item.category;
          const nameMatches = item.name.toLowerCase().includes(rule.targetKeyword.toLowerCase());
          const specMatches = rule.targetSpecKeyword 
            ? item.spec.toLowerCase().includes(rule.targetSpecKeyword.toLowerCase()) 
            : true;
          return catMatches && nameMatches && specMatches;
        });

        if (matchingRule) {
          const { value, logs } = evaluateFormula(matchingRule.expression, items, projectVars, 0, originalRawRows);
          return {
            ...item,
            formula: matchingRule.expression,
            calculatedQty: value,
            isCalculated: true,
            matchLog: logs
          };
        }
      }
      return item;
    });

    setCalculatedItems(computed);
  }, [items, formulas]);

  const toggleRowLog = (rowIndex: number) => {
    setExpandedRowLogs(prev => ({ ...prev, [rowIndex]: !prev[rowIndex] }));
  };

  const filterItems = calculatedItems.filter(item => {
    // Only show items that are calculated targets
    const isTarget = item.category === 'gasket' || item.category === 'fastener' || item.category === 'auxiliary';
    if (!isTarget) return false;
    
    if (activeTab === 'all') return true;
    return item.category === activeTab;
  });

  // Method 1: Export as processed in-place Excel sheet (Actual file writer)
  const handleExportExcel = () => {
    if (!originalWorkbook || !sheetName) {
      alert("请先导入项目 Excel 工作表才能进行一键填回导出哦！");
      return;
    }

    setIsExporting(true);
    
    try {
      // Clone workbook and worksheet in memory
      const wb = { ...originalWorkbook };
      const ws = wb.Sheets[sheetName];
      if (!ws) throw new Error("无法读取到对应的工作表内存指针");

      let filledCount = 0;

      // Locate the mapped quantity / length column
      const qtyColIndex = mappings.qtyCol;

      filterItems.forEach(item => {
        if (item.isCalculated && item.calculatedQty !== undefined) {
          // Address of the cell using sheetjs utilities
          const cellAddress = XLSX.utils.encode_cell({ r: item.rowIndex, c: qtyColIndex });
          const cell = ws[cellAddress];

          if (cell) {
            cell.v = item.calculatedQty;
            cell.t = "n"; // format as numeric type
            if (cell.w) delete cell.w; // Clear cached string representation
          } else {
            // Allocate new cell
            ws[cellAddress] = { t: "n", v: item.calculatedQty };
          }
          filledCount++;
        }
      });

      // Write in Excel file bytes
      XLSX.writeFile(wb, `单元加工材料细目表_自动套算填回.xlsx`);

      // Simple brief browser notification
      alert(`🎉 填回成功！已将计算得出的辅材与紧固件含量准确写回原表，共更新了 ${filledCount} 个相关单元格，并开始下载更新后的 Excel 文件。`);
    } catch (err: any) {
      console.error(err);
      alert(`无法填入 Excel: ${err.message}`);
    } finally {
      setIsExporting(false);
    }
  };

  // Method 2: Synchronize to Mock Local office instance via "Excel Data Connection API"
  const handleMockDirectSync = () => {
    setSyncStatus('linking');
    setSyncLogs([]);
    
    // Simulate real-time Local Excel WebSocket handshaking
    const steps = [
      { t: 0, text: "正在侦测本地 Excel / WPS Office 宿主运行环境..." },
      { t: 600, text: "连接成功: 幕墙工程 Excel 数据总线通道 (PORT 3000 IPC Bridge) 已建立" },
      { t: 1200, text: "正在匹配 Excel 活动工作簿名称: 《单元加工材料细目表》..." },
      { t: 1800, text: "活动表：'单元分块明细表' (识别有 138 行数据已加载，正在同步列绑定)" },
      { t: 2400, text: "定位输出列 A1 网格(图表辅料填数区) - [E列计设计量, H列计采购量] 结构匹配" },
      { t: 3000, text: "正在写入计算值...... (写入 胶条量 及 紧固件) ..." },
      { t: 3500, text: "同步成功：共计填入 14 行辅材、紧固件与海绵的最终数量！" },
    ];

    steps.forEach(step => {
      setTimeout(() => {
        setSyncLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${step.text}`]);
        if (step.t === 3500) {
          setSyncStatus('success');
        }
      }, step.t);
    });
  };

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="calcs_dashboard">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center mb-6 gap-4 border-b border-gray-50 pb-5">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Radio className="w-5 h-5 text-emerald-500 animate-pulse" />
            第四步：核算细目成果及 Excel 精准填回 Sync
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            系统根据套算规则，对型材对应的气密隔音胶条、结构耐候密封胶、自攻螺母角码等进行了高精度材料算量配比。
          </p>
        </div>

        <div className="flex flex-wrap gap-3 w-full lg:w-auto">
          {/* Real local sheet output action */}
          <button
            onClick={handleExportExcel}
            className="flex-1 lg:flex-none justify-center bg-indigo-600 hover:bg-indigo-700 text-white font-semibold px-4 py-2.5 rounded transition-all text-xs flex items-center gap-2 shadow-sm cursor-pointer"
            disabled={calculatedItems.length === 0 || isExporting}
          >
            {isExporting ? <RefreshCw className="w-4 h-4 animate-spin" /> : <FileSpreadsheet className="w-4 h-4" />}
            一键写入单元格并导出原表
          </button>
        </div>
      </div>

      {/* Excel sync status banner if active */}
      <AnimatePresence>
        {syncStatus !== 'idle' && (
          <motion.div
            initial={{ opacity: 0, y: -10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className={`p-5 rounded border mb-6 text-xs ${
              syncStatus === 'success'
                ? "bg-emerald-50 border-emerald-200 text-emerald-800"
                : "bg-indigo-50 border-indigo-100 text-indigo-800"
            }`}
          >
            <div className="flex justify-between items-center mb-3">
              <div className="flex items-center gap-2 font-bold uppercase tracking-wider text-[10px]">
                <Network className={`w-4 h-4 text-indigo-600 ${syncStatus !== 'success' ? 'animate-spin' : ''}`} />
                <span>LOCAL OFFICE INSTANCE DISCOVERY IPC (PORT 3000 BRIDGING LINK)</span>
              </div>
              <button
                onClick={() => setSyncStatus('idle')}
                className="text-gray-400 hover:text-gray-600 font-bold"
              >
                关闭
              </button>
            </div>
            
            <div className="bg-slate-900 text-emerald-400 font-mono p-3.5 rounded space-y-1.5 max-h-[140px] overflow-y-auto mb-3 text-[10px] border border-slate-800 shadow-inner leading-relaxed select-all">
              {syncLogs.length === 0 ? (
                <div className="text-slate-500 italic">SYSTEM READY: WAITING FOR INITIATION HANDSHAKE</div>
              ) : (
                syncLogs.map((log, index) => (
                  <div key={index} className="leading-relaxed">{log}</div>
                ))
              )}
            </div>

            {syncStatus === 'success' && (
              <div className="flex items-center gap-1.5 text-xs font-bold text-emerald-800 bg-emerald-50 p-2.5 rounded border border-emerald-200 animate-pulse mt-2">
                <CheckCircle2 className="w-4.5 h-4.5 text-emerald-600 shrink-0" />
                恭喜！辅材用量算料值已一键穿透直达本地 Excel 运行端，对应单元格高亮填入完毕。
              </div>
            )}
          </motion.div>
        )}
      </AnimatePresence>

      {/* Selector tab filter and stats bar */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-4.5 gap-3">
        <div className="flex gap-1 bg-slate-100 p-1 rounded border border-slate-200">
          <button
            onClick={() => setActiveTab('all')}
            className={`px-3 py-1.5 rounded text-xs font-semibold transition-colors cursor-pointer ${
              activeTab === 'all' ? "bg-white text-slate-900 shadow-xs border border-slate-200" : "text-slate-500 hover:text-slate-800"
            }`}
          >
            全部核算明细 ({calculatedItems.filter(i => i.category === 'gasket' || i.category === 'fastener' || i.category === 'auxiliary').length})
          </button>
          <button
            onClick={() => setActiveTab('gasket')}
            className={`px-3 py-1.5 rounded text-xs font-semibold transition-colors cursor-pointer ${
              activeTab === 'gasket' ? "bg-white text-emerald-700 shadow-xs border border-slate-200" : "text-slate-500 hover:text-emerald-700"
            }`}
          >
            型材胶条 ({calculatedItems.filter(i => i.category === 'gasket').length})
          </button>
          <button
            onClick={() => setActiveTab('fastener')}
            className={`px-3 py-1.5 rounded text-xs font-semibold transition-colors cursor-pointer ${
              activeTab === 'fastener' ? "bg-white text-amber-700 shadow-xs border border-slate-200" : "text-slate-500 hover:text-amber-700"
            }`}
          >
            五金紧固件 ({calculatedItems.filter(i => i.category === 'fastener').length})
          </button>
          <button
            onClick={() => setActiveTab('auxiliary')}
            className={`px-3 py-1.5 rounded text-xs font-semibold transition-colors cursor-pointer ${
              activeTab === 'auxiliary' ? "bg-white text-indigo-700 shadow-xs border border-slate-200" : "text-slate-500 hover:text-indigo-700"
            }`}
          >
            结构辅助件 ({calculatedItems.filter(i => i.category === 'auxiliary').length})
          </button>
        </div>

        <div className="text-xs text-slate-500 font-medium">
          锁定计算填回列: <span className="bg-amber-50 text-amber-800 border border-amber-200 px-2 py-1 rounded font-mono font-bold">L_Col {String.fromCharCode(65 + mappings.qtyCol)}</span>
        </div>
      </div>

      <div className="border border-slate-200 rounded overflow-hidden shadow-xs">
        <table className="min-w-full divide-y divide-slate-200 text-xs">
          <thead className="bg-slate-50/80 text-slate-500 font-bold tracking-wider text-[10px] uppercase">
            <tr>
              <th className="px-5 py-3 text-left w-12 border-b border-slate-200">行号</th>
              <th className="px-5 py-3 text-left border-b border-slate-200">品名/零件及装配件名称</th>
              <th className="px-5 py-3 text-left border-b border-slate-200">规格型号及加工图号</th>
              <th className="px-5 py-3 text-left border-b border-slate-200">原默认值</th>
              <th className="px-5 py-3 text-left border-b border-slate-200">智能套算计算式</th>
              <th className="px-5 py-3 text-right font-bold text-indigo-700 bg-indigo-50/10 border-b border-slate-200">系统换算定额</th>
              <th className="px-5 py-3 text-left w-12 border-b border-slate-200">单位</th>
              <th className="px-5 py-3 text-center w-28 border-b border-slate-200">逻辑审计链</th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-slate-100 text-slate-700">
            {filterItems.length === 0 ? (
              <tr>
                <td colSpan={8} className="px-5 py-10 text-center text-slate-400 italic">
                  您上传的表格中，当前筛选的算量类别中暂无需要用公式自动算出的空白项条目。
                </td>
              </tr>
            ) : (
              filterItems.map(item => {
                const isExpanded = expandedRowLogs[item.rowIndex] || false;
                
                return (
                  <React.Fragment key={item.rowIndex}>
                    <tr className={`hover:bg-indigo-50/20 transition-colors ${item.isCalculated ? 'bg-indigo-50/10' : ''}`}>
                      <td className="px-5 py-3 font-mono text-slate-400">
                        {item.rowIndex + 1}
                      </td>
                      <td className="px-5 py-3 font-bold text-slate-900">
                        {item.name}
                      </td>
                      <td className="px-5 py-3 font-mono text-slate-600 bg-slate-50/40">
                        {item.spec}
                      </td>
                      <td className="px-5 py-3 text-slate-400 line-through font-mono">
                        {item.originalCategory ? item.qty : "空" }
                      </td>
                      <td className="px-5 py-3 text-indigo-600 font-mono text-[11px]">
                        {item.formula ? (
                          <span className="bg-indigo-50 border border-indigo-100/60 px-1.5 py-0.5 rounded text-[10px] font-bold cursor-help" title={`完整公式: ${item.formula}`}>
                            {simplifyFormulaDisplay(item.formula)}
                          </span>
                        ) : (
                          <span className="text-slate-300 italic text-[10px]">未匹配到套算公式</span>
                        )}
                      </td>
                      <td className="px-5 py-3 text-right bg-indigo-50/20 font-bold text-sm text-indigo-700 font-mono">
                        {item.isCalculated ? item.calculatedQty?.toFixed(2) : "—"}
                      </td>
                      <td className="px-5 py-3 font-medium text-slate-500">
                        {item.unit}
                      </td>
                      <td className="px-5 py-3 text-center">
                        <button
                          onClick={() => toggleRowLog(item.rowIndex)}
                          className={`text-[11px] font-bold px-2.5 py-1 rounded border transition-all flex items-center gap-1 mx-auto hover:bg-slate-50 cursor-pointer ${
                            isExpanded 
                              ? "bg-indigo-50 border-indigo-200 text-indigo-700 font-bold" 
                              : "bg-white border-slate-200 text-slate-600"
                          }`}
                        >
                          <span>审计展开</span>
                          {isExpanded ? <ChevronUp className="w-3 h-3 text-indigo-500" /> : <ChevronDown className="w-3 h-3 text-slate-400" />}
                        </button>
                      </td>
                    </tr>
                    
                    {/* Log details trace row */}
                    {isExpanded && (
                      <tr>
                        <td colSpan={8} className="bg-slate-50/80 px-8 py-4 border-t border-b border-slate-200">
                          <div className="text-xs text-slate-800">
                            <div className="font-bold text-[10px] text-slate-500 uppercase tracking-widest mb-2 flex items-center gap-1">
                              <Sparkles className="w-3.5 h-3.5 text-indigo-600" />
                              算定逻辑溯源核算审计链 (Calculated Quant Auditing Logs)
                            </div>
                            <div className="bg-white p-3.5 rounded border border-slate-200 font-mono text-[11px] space-y-1.5 text-slate-600 leading-relaxed max-w-4xl">
                              {item.matchLog && item.matchLog.length > 0 ? (
                                item.matchLog.map((log, lIdx) => (
                                  <div key={lIdx} className="border-l-2 border-indigo-200 pl-3.5">
                                    {log}
                                  </div>
                                ))
                              ) : (
                                <div className="text-slate-400 italic font-sans">该辅助类物料未能检索到可套数的参考型材尺寸、截面尺寸，请在上方校验规则设置。</div>
                              )}
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
    </div>
  );
}
