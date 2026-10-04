import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { RangeConfiguration, ProcessedItem } from "../types";
import { parseExcelRange, resolveRangeByBColumnKeyword } from "../utils/calcEngine";
import { Grid, Eye, Shuffle, CheckCircle, AlertTriangle, Blocks, Search, Sparkles } from "lucide-react";

interface ExcelRangeSelectorProps {
  ranges: RangeConfiguration;
  onUpdateRanges: (newRanges: RangeConfiguration) => void;
  itemsCount: number;
  items: ProcessedItem[];
  workbook: XLSX.WorkBook | null;
  sheetName: string;
  rawRows: string[][];
}

export default function ExcelRangeSelector({
  ranges,
  onUpdateRanges,
  itemsCount,
  items,
  workbook,
  sheetName,
  rawRows,
}: ExcelRangeSelectorProps) {
  const [profileInput, setProfileInput] = useState(ranges.profileRange);
  const [steelInput, setSteelInput] = useState(ranges.steelRange);
  const [panelInput, setPanelInput] = useState(ranges.panelRange);

  // B-Column search keyword states linked to the configurations
  const [profileKeyword, setProfileKeyword] = useState(ranges.profileKeyword || "铝  型  材");
  const [steelKeyword, setSteelKeyword] = useState(ranges.steelKeyword || "钢   件");
  const [panelKeyword, setPanelKeyword] = useState(ranges.panelKeyword || "面   板");

  // Detection messages state for in-card quick visual confirmation or warning
  const [detectionStatus, setDetectionStatus] = useState<Record<string, { success: boolean; msg: string } | null>>({
    profile: null,
    steel: null,
    panel: null,
  });

  useEffect(() => {
    setProfileInput(ranges.profileRange);
    setSteelInput(ranges.steelRange);
    setPanelInput(ranges.panelRange);
    
    setProfileKeyword(ranges.profileKeyword || "铝  型  材");
    setSteelKeyword(ranges.steelKeyword || "钢   件");
    setPanelKeyword(ranges.panelKeyword || "面   板");
  }, [ranges]);

  const handleApply = () => {
    onUpdateRanges({
      profileRange: profileInput,
      steelRange: steelInput,
      panelRange: panelInput,
      profileKeyword,
      steelKeyword,
      panelKeyword,
    });
  };

  // Parse ranges to show feedback to the designer
  const pRange = parseExcelRange(profileInput);
  const sRange = parseExcelRange(steelInput);
  const paRange = parseExcelRange(panelInput);

  // Helper to trigger smart detection on a single category group
  const handleSmartDetectGroup = (groupType: 'profile' | 'steel' | 'panel') => {
    let kw = '';
    if (groupType === 'profile') kw = profileKeyword;
    else if (groupType === 'steel') kw = steelKeyword;
    else if (groupType === 'panel') kw = panelKeyword;

    if (!kw.trim()) {
      alert("请输入匹配关键词（如'型材'、'钢件'），以便检索！");
      return;
    }

    const ws = workbook ? workbook.Sheets[sheetName] : null;
    const resolved = resolveRangeByBColumnKeyword(ws, rawRows, kw);

    if (resolved) {
      // Form coordinates from Column C to L
      const rangeStr = `C${resolved.startRow + 1}:L${resolved.endRow + 1}`;
      
      if (groupType === 'profile') {
        setProfileInput(rangeStr);
      } else if (groupType === 'steel') {
        setSteelInput(rangeStr);
      } else if (groupType === 'panel') {
        setPanelInput(rangeStr);
      }

      setDetectionStatus(prev => ({
        ...prev,
        [groupType]: {
          success: true,
          msg: `🔍 B列匹配成功: C${resolved.startRow + 1}:L${resolved.endRow + 1}`
        }
      }));
    } else {
      setDetectionStatus(prev => ({
        ...prev,
        [groupType]: {
          success: false,
          msg: `⚠️ B列/A列中未检索到 "${kw}" 关键词`
        }
      }));
    }
  };

  // Global Master trigger that processes all 3 groups sequentially using keywords
  const handleSmartDetectAll = () => {
    const ws = workbook ? workbook.Sheets[sheetName] : null;
    let successCount = 0;
    const missing: string[] = [];
    const nextRanges = { 
      profile: profileInput,
      steel: steelInput,
      panel: panelInput,
    };

    const targets = [
      { type: 'profile' as const, label: 'A组型材', val: profileKeyword, setInput: setProfileInput },
      { type: 'steel' as const, label: 'B组钢件', val: steelKeyword, setInput: setSteelInput },
      { type: 'panel' as const, label: 'C组面板', val: panelKeyword, setInput: setPanelInput },
    ];

    targets.forEach(item => {
      if (!item.val.trim()) {
        missing.push(`${item.label} (关键字未填写)`);
        return;
      }
      const resolved = resolveRangeByBColumnKeyword(ws, rawRows, item.val);
      if (resolved) {
        const rangeStr = `C${resolved.startRow + 1}:L${resolved.endRow + 1}`;
        item.setInput(rangeStr);
        nextRanges[item.type] = rangeStr;
        successCount++;

        setDetectionStatus(prev => ({
          ...prev,
          [item.type]: {
            success: true,
            msg: `🔍 B列匹配成功: ${rangeStr}`
          }
        }));
      } else {
        missing.push(`${item.label} ("${item.val}")`);
        setDetectionStatus(prev => ({
          ...prev,
          [item.type]: {
            success: false,
            msg: `⚠️ B列/A列中未发现该名称`
          }
        }));
      }
    });

    if (successCount > 0) {
      // Sync configurations immediately with parent
      onUpdateRanges({
        profileRange: nextRanges.profile,
        steelRange: nextRanges.steel,
        panelRange: nextRanges.panel,
        profileKeyword,
        steelKeyword,
        panelKeyword,
      });

      alert(`✨ B列大类关键字智检区间成功！\n成功识别并更正 ${successCount} 个坐标区间。\n${
        missing.length > 0 ? `未匹配组: ${missing.join(", ")}` : "全部大类分区已精确锁定。"
      }`);
    } else {
      alert(`⚠️ 未能根据B列大类名称检索到任何区间。\n未匹配信息: ${missing.join(", ")}\n请核对大类名称，或直接手动指定坐标。`);
    }
  };

  // Get active items count in each parsed range
  const profileItemsCount = items.filter(i => i.category === 'profile').length;
  const steelItemsCount = items.filter(i => i.category === 'steel').length;
  const panelItemsCount = items.filter(i => i.category === 'panel').length;

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="module_range_selector">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Grid className="w-5 h-5 text-indigo-600" />
            工程单元区间引用与物料分割导入
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            支持两种划分分割带方式：一、手动指定区间引用（如 C8:L57）；二、通过 B 列合并大类单元格或内容（如'铝型材'）智能检索。
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={handleSmartDetectAll}
            className="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-4 py-2.5 rounded transition-all flex items-center gap-1.5 cursor-pointer shadow-sm shadow-emerald-100"
            title="读取当前活动工作表B列的分类大类合并单元格，智能对齐并计算三组数据首尾坐标范围"
          >
            <Sparkles className="w-3.5 h-3.5" />
            B列大类名称智能识别所有区间
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Profile Range Card */}
        <div className="bg-slate-50/50 p-5 rounded border border-slate-200 hover:border-indigo-200 transition-colors flex flex-col justify-between space-y-4">
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="font-bold text-xs text-indigo-700 uppercase tracking-wider bg-indigo-50 px-2 py-1 rounded">
                A 组：铝型材数据范围
              </span>
              <span className="text-[10px] text-slate-450 font-mono font-bold">
                {profileItemsCount} / {itemsCount} 行
              </span>
            </div>
            
            <div className="space-y-3">
              <div>
                <label className="block text-[11px] font-bold text-slate-500 mb-1.5">
                  工作表单元引用坐标
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={profileInput}
                    onChange={(e) => setProfileInput(e.target.value)}
                    placeholder="如 C8:L57"
                    className="w-full bg-white border border-slate-200 rounded px-3 py-2 text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
              </div>

              <div className="text-[11px] text-slate-500 space-y-1 bg-white p-2.5 rounded border border-slate-150">
                {pRange ? (
                  <div className="flex items-center gap-1.5 text-emerald-700 font-medium">
                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" />
                    <span>
                      引用成立：Excel 行 {pRange.startRow + 1} 至 {pRange.endRow + 1} (共 {pRange.endRow - pRange.startRow + 1} 行)
                    </span>
                  </div>
                ) : (
                  <div className="flex items-center gap-1.5 text-rose-600 font-medium">
                    <AlertTriangle className="w-3.5 h-3.5 text-rose-500 flex-shrink-0" />
                    <span>坐标格式不合法，请使用 A1:B10 格式</span>
                  </div>
                )}
                <p className="text-[10px] text-slate-400">
                  系统在此坐标段内过滤物料品名非空行，自动归档为“铝合金型材原材”进行长度配套折算。
                </p>
              </div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-200 space-y-2">
            <label className="block text-[11px] font-bold text-slate-500 flex justify-between items-center">
              <span>B列合并单元格 / 大类关键字</span>
              <span className="text-[9px] text-slate-400 font-normal">支持空格、子串匹配</span>
            </label>
            <div className="flex gap-1.5">
              <input
                type="text"
                value={profileKeyword}
                onChange={(e) => setProfileKeyword(e.target.value)}
                placeholder="铝型材 过滤名"
                className="w-full bg-white border border-slate-200 rounded px-2.5 py-1.5 text-xs text-slate-700 font-semibold focus:outline-none focus:border-indigo-500"
              />
              <button
                type="button"
                onClick={() => handleSmartDetectGroup('profile')}
                className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-[11px] font-bold px-2 rounded flex items-center gap-1 cursor-pointer whitespace-nowrap transition-colors"
              >
                <Search className="w-3 h-3 text-indigo-500" />
                检索定位
              </button>
            </div>
            {detectionStatus.profile && (
              <div className={`text-[10px] px-2 py-1 rounded border leading-tight ${
                detectionStatus.profile.success 
                  ? 'bg-emerald-50/50 border-emerald-100 text-emerald-700' 
                  : 'bg-rose-50/40 border-rose-100 text-rose-500'
              }`}>
                {detectionStatus.profile.msg}
              </div>
            )}
          </div>
        </div>

        {/* Steel Range Card */}
        <div className="bg-slate-50/50 p-5 rounded border border-slate-200 hover:border-indigo-200 transition-colors flex flex-col justify-between space-y-4">
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="font-bold text-xs text-amber-700 uppercase tracking-wider bg-amber-50 px-2 py-1 rounded">
                B 组：钢构件/五金钢件范围
              </span>
              <span className="text-[10px] text-slate-450 font-mono font-bold">
                {steelItemsCount} / {itemsCount} 行
              </span>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-[11px] font-bold text-slate-500 mb-1.5">
                  工作表单元引用坐标
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={steelInput}
                    onChange={(e) => setSteelInput(e.target.value)}
                    placeholder="如 C58:L67"
                    className="w-full bg-white border border-slate-200 rounded px-3 py-2 text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
              </div>

              <div className="text-[11px] text-slate-500 space-y-1 bg-white p-2.5 rounded border border-slate-150">
                {sRange ? (
                  <div className="flex items-center gap-1.5 text-emerald-700 font-medium">
                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" />
                    <span>
                      引用成立：Excel 行 {sRange.startRow + 1} 至 {sRange.endRow + 1} (共 {sRange.endRow - sRange.startRow + 1} 行)
                    </span>
                  </div>
                ) : (
                  <div className="flex items-center gap-1.5 text-rose-600 font-medium">
                    <AlertTriangle className="w-3.5 h-3.5 text-rose-500 flex-shrink-0" />
                    <span>坐标格式不合法，请使用 A1:B10 格式</span>
                  </div>
                )}
                <p className="text-[10px] text-slate-400">
                  本段将锁定挂件、角码及转接钢板数据。通过 SUM_QTY 统计其总配数来进行辅件精细配套。
                </p>
              </div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-200 space-y-2">
            <label className="block text-[11px] font-bold text-slate-500 flex justify-between items-center">
              <span>B列合并单元格 / 大类关键字</span>
              <span className="text-[9px] text-slate-400 font-normal">支持空格、子串匹配</span>
            </label>
            <div className="flex gap-1.5">
              <input
                type="text"
                value={steelKeyword}
                onChange={(e) => setSteelKeyword(e.target.value)}
                placeholder="钢件 过滤名"
                className="w-full bg-white border border-slate-200 rounded px-2.5 py-1.5 text-xs text-slate-700 font-semibold focus:outline-none focus:border-indigo-500"
              />
              <button
                type="button"
                onClick={() => handleSmartDetectGroup('steel')}
                className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-[11px] font-bold px-2 rounded flex items-center gap-1 cursor-pointer whitespace-nowrap transition-colors"
              >
                <Search className="w-3 h-3 text-indigo-500" />
                检索定位
              </button>
            </div>
            {detectionStatus.steel && (
              <div className={`text-[10px] px-2 py-1 rounded border leading-tight ${
                detectionStatus.steel.success 
                  ? 'bg-emerald-50/50 border-emerald-100 text-emerald-700' 
                  : 'bg-rose-50/40 border-rose-100 text-rose-500'
              }`}>
                {detectionStatus.steel.msg}
              </div>
            )}
          </div>
        </div>

        {/* Panel Range Card */}
        <div className="bg-slate-50/50 p-5 rounded border border-slate-200 hover:border-indigo-200 transition-colors flex flex-col justify-between space-y-4">
          <div className="space-y-4">
            <div className="flex justify-between items-center">
              <span className="font-bold text-xs text-emerald-700 uppercase tracking-wider bg-emerald-50 px-2 py-1 rounded">
                C 组：幕墙玻璃/装饰面板范围
              </span>
              <span className="text-[10px] text-slate-450 font-mono font-bold">
                {panelItemsCount} / {itemsCount} 行
              </span>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-[11px] font-bold text-slate-500 mb-1.5">
                  工作表单元引用坐标
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={panelInput}
                    onChange={(e) => setPanelInput(e.target.value)}
                    placeholder="如 C68:L97"
                    className="w-full bg-white border border-slate-200 rounded px-3 py-2 text-xs font-mono font-bold text-slate-800 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500"
                  />
                </div>
              </div>

              <div className="text-[11px] text-slate-500 space-y-1 bg-white p-2.5 rounded border border-slate-150">
                {paRange ? (
                  <div className="flex items-center gap-1.5 text-emerald-700 font-medium">
                    <CheckCircle className="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" />
                    <span>
                      引用成立：Excel 行 {paRange.startRow + 1} 至 {paRange.endRow + 1} (共 {paRange.endRow - paRange.startRow + 1} 行)
                    </span>
                  </div>
                ) : (
                  <div className="flex items-center gap-1.5 text-rose-600 font-medium">
                    <AlertTriangle className="w-3.5 h-3.5 text-rose-500 flex-shrink-0" />
                    <span>坐标格式不合法，请使用 A1:B10 格式</span>
                  </div>
                )}
                <p className="text-[10px] text-slate-400">
                  本段获取中空钢化玻璃、石材或铝单板。用于提取宽、高、单体面积的复杂汇总套算。
                </p>
              </div>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-200 space-y-2">
            <label className="block text-[11px] font-bold text-slate-500 flex justify-between items-center">
              <span>B列合并单元格 / 大类关键字</span>
              <span className="text-[9px] text-slate-400 font-normal">支持空格、子串匹配</span>
            </label>
            <div className="flex gap-1.5">
              <input
                type="text"
                value={panelKeyword}
                onChange={(e) => setPanelKeyword(e.target.value)}
                placeholder="面板 过滤名"
                className="w-full bg-white border border-slate-200 rounded px-2.5 py-1.5 text-xs text-slate-700 font-semibold focus:outline-none focus:border-indigo-500"
              />
              <button
                type="button"
                onClick={() => handleSmartDetectGroup('panel')}
                className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-[11px] font-bold px-2 rounded flex items-center gap-1 cursor-pointer whitespace-nowrap transition-colors"
              >
                <Search className="w-3 h-3 text-indigo-500" />
                检索定位
              </button>
            </div>
            {detectionStatus.panel && (
              <div className={`text-[10px] px-2 py-1 rounded border leading-tight ${
                detectionStatus.panel.success 
                  ? 'bg-emerald-50/50 border-emerald-100 text-emerald-700' 
                  : 'bg-rose-50/40 border-rose-100 text-rose-500'
              }`}>
                {detectionStatus.panel.msg}
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="mt-4 flex flex-wrap gap-2 items-center text-xs">
        <span className="text-slate-400 font-bold uppercase tracking-wider text-[10px] flex items-center gap-1">
          <Blocks className="w-3.5 h-3.5 text-indigo-500" />
          常用幕墙工程表模板区间预设一键套用:
        </span>
        <button
          onClick={() => {
            setProfileInput("C8:L12");
            setSteelInput("C16:L19");
            setPanelInput("C22:L25");
            onUpdateRanges({ 
              profileRange: "C8:L12", 
              steelRange: "C16:L19", 
              panelRange: "C22:L25",
              profileKeyword,
              steelKeyword,
              panelKeyword,
            });
          }}
          className="bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-bold px-2.5 py-1 rounded border border-emerald-200 cursor-pointer text-[11px]"
        >
          双向示范料单切片段 (C8:L12 C16:L19 C22:L25)
        </button>

        <button
          onClick={() => {
            setProfileInput("C8:L57");
            setSteelInput("C58:L67");
            setPanelInput("C68:L97");
            onUpdateRanges({ 
              profileRange: "C8:L57", 
              steelRange: "C58:L67", 
              panelRange: "C68:L97",
              profileKeyword,
              steelKeyword,
              panelKeyword,
            });
          }}
          className="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-semibold px-2.5 py-1 rounded border border-indigo-200 cursor-pointer text-[11px]"
        >
          标准版-50型材-20钢件-30面板 (C8:L57 C58:L67 C68:L97)
        </button>

        <button
          onClick={() => {
            setProfileInput("C5:L30");
            setSteelInput("C31:L45");
            setPanelInput("C46:L80");
            onUpdateRanges({ 
              profileRange: "C5:L30", 
              steelRange: "C31:L45", 
              panelRange: "C46:L80",
              profileKeyword,
              steelKeyword,
              panelKeyword,
            });
          }}
          className="bg-slate-100 hover:bg-slate-200 text-slate-705 px-2.5 py-1 rounded border border-slate-200 cursor-pointer font-medium text-[11px]"
        >
          高密度简版 (C5:L30 C31:L45 C46:L80)
        </button>
      </div>
    </div>
  );
}
