import React, { useState, useEffect, useMemo } from "react";
import { 
  Play, RefreshCw, CheckCircle2, AlertCircle, Server, 
  ListTodo, FileSpreadsheet, Activity, ChevronRight, Square 
} from "lucide-react";

interface BatchImportFillbackProps {
  excelSheets: string[];
  wsConnected: boolean;
  batchActive: boolean;
  batchStopping?: boolean;
  batchProgress: {
    currentSheetName: string;
    step: 'idle' | 'reading' | 'calculating' | 'writing' | 'success' | 'error' | 'stopping';
    completedCount: number;
    totalCount: number;
    logs: Array<{ sheetName: string; status: 'success' | 'failed' | 'stopped'; message: string }>;
  };
  onStartBatch: (selectedSheets: string[]) => void;
  onStopBatch: () => void;
  onRefreshSheets: () => void;
  pullLoading: boolean;
}

export default function BatchImportFillback({
  excelSheets,
  wsConnected,
  batchActive,
  batchStopping = false,
  batchProgress,
  onStartBatch,
  onStopBatch,
  onRefreshSheets,
  pullLoading,
}: BatchImportFillbackProps) {
  const [selectedSheetsState, setSelectedSheetsState] = useState<Record<string, boolean>>({});

  // Initialize selected sheets list when excelSheets changes
  useEffect(() => {
    const initial: Record<string, boolean> = {};
    excelSheets.forEach(sheet => {
      // Default to checked unless sheet name is "目录"
      initial[sheet] = sheet !== "目录";
    });
    setSelectedSheetsState(initial);
  }, [excelSheets]);

  const sheetsToProcess = useMemo(() => {
    return excelSheets.filter(sheet => selectedSheetsState[sheet]);
  }, [excelSheets, selectedSheetsState]);

  const handleToggleSheet = (sheetName: string) => {
    if (batchActive) return;
    setSelectedSheetsState(prev => ({
      ...prev,
      [sheetName]: !prev[sheetName]
    }));
  };

  const handleSelectAll = (checked: boolean) => {
    if (batchActive) return;
    const next: Record<string, boolean> = {};
    excelSheets.forEach(sheet => {
      next[sheet] = checked && sheet !== "目录";
    });
    setSelectedSheetsState(next);
  };

  const handleTriggerBatch = () => {
    if (sheetsToProcess.length === 0) {
      alert("请至少勾选一个待处理的工作表！");
      return;
    }
    onStartBatch(sheetsToProcess);
  };

  const getStepDescription = (step: string) => {
    switch (step) {
      case "reading":
        return "正在从 Excel 极速提取物理下料行...";
      case "calculating":
        return "正在运行规则计算引擎，统筹配套辅料...";
      case "writing":
        return "正在通过 COM 穿透对目标工作表高亮回填...";
      case "stopping":
        return "用户已请求停止，正在安全中止后续任务...";
      case "success":
        return "当前工作表回填完毕！";
      case "error":
        return "处理遇到意外中断";
      default:
        return "准备就绪";
    }
  };

  const progressPercent = batchProgress.totalCount > 0 
    ? Math.round((batchProgress.completedCount / batchProgress.totalCount) * 100) 
    : 0;

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6" id="module_batch_import_fillback">
      {/* Title Header */}
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center border-b border-slate-100 pb-5 gap-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <ListTodo className="w-5 h-5 text-indigo-600 animate-pulse" />
            批量工作表导入与算料填回 (Batch Sheet Processor)
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            一键全自动循环运行：按顺序读取每个工作表的型材/钢件/面板，计算辅配料后自动填回原表，免除手动翻页与重复操作。
          </p>
        </div>

        <button
          onClick={onRefreshSheets}
          disabled={!wsConnected || batchActive || pullLoading}
          className="flex items-center gap-1.5 px-3.5 py-2 border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-bold rounded-lg transition-all cursor-pointer disabled:opacity-40"
        >
          <RefreshCw className={`w-3.5 h-3.5 ${pullLoading ? "animate-spin" : ""}`} />
          刷新 Excel 工作表列表
        </button>
      </div>

      {/* Connection State Warning */}
      {!wsConnected && (
        <div className="bg-amber-50 border border-amber-200 p-4.5 rounded-xl flex items-start gap-3.5 text-left shadow-2xs">
          <AlertCircle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
          <div className="space-y-1">
            <h4 className="text-xs font-bold text-amber-900">本地 OLE COM 网桥未连通</h4>
            <p className="text-[11px] text-amber-700/85 leading-relaxed">
              批量导入填回操作依赖于本地桌面端 Excel/WPS。请确保在 <b>“本地 OLE COM 网桥连接”</b> 菜单中下载并运行了 <code>com_bridge.py</code> 脚本，启动本地双向配对。
            </p>
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left Side: Worksheets Selection Checkbox List */}
        <div className="lg:col-span-5 border border-slate-200 rounded-xl overflow-hidden flex flex-col h-[400px]">
          <div className="bg-slate-50 border-b border-slate-200 px-4 py-3 flex items-center justify-between shrink-0">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 text-slate-500" />
              <span className="text-xs font-bold text-slate-700">可读取的工作表列表</span>
            </div>
            <span className="text-[10px] bg-slate-200 text-slate-600 px-2 py-0.5 rounded font-mono font-bold">
              共 {excelSheets.length} 个
            </span>
          </div>

          {excelSheets.length > 0 ? (
            <div className="flex-1 overflow-y-auto p-4 space-y-2">
              {/* Select All */}
              <label className="flex items-center gap-3 p-2.5 rounded-lg hover:bg-slate-50 border border-transparent transition-colors cursor-pointer select-none">
                <input
                  type="checkbox"
                  disabled={batchActive}
                  checked={excelSheets.length > 0 && excelSheets.every(s => s === "目录" || selectedSheetsState[s])}
                  onChange={(e) => handleSelectAll(e.target.checked)}
                  className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-4 h-4 cursor-pointer disabled:opacity-50"
                />
                <span className="text-xs font-extrabold text-slate-800">默认全选 (排除“目录”工作表)</span>
              </label>

              <hr className="border-slate-100 my-1" />

              {excelSheets.map((sheet, idx) => (
                <div 
                  key={idx}
                  onClick={() => handleToggleSheet(sheet)}
                  className={`flex items-center justify-between p-2.5 rounded-lg border text-xs transition-all select-none cursor-pointer ${
                    selectedSheetsState[sheet]
                      ? "bg-indigo-50/20 border-indigo-150"
                      : "bg-white border-slate-150 hover:bg-slate-50/50"
                  } ${sheet === "目录" ? "opacity-60" : ""}`}
                >
                  <div className="flex items-center gap-3">
                    <input
                      type="checkbox"
                      checked={!!selectedSheetsState[sheet]}
                      disabled={batchActive || sheet === "目录"}
                      onChange={() => {}} // Handled by outer click
                      className="rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 w-3.5 h-3.5 cursor-pointer disabled:opacity-50"
                    />
                    <span className={`font-medium ${selectedSheetsState[sheet] ? "text-indigo-900 font-bold" : "text-slate-600"}`}>
                      {sheet}
                    </span>
                  </div>
                  {sheet === "目录" && (
                    <span className="text-[9px] bg-slate-100 border border-slate-200 text-slate-400 px-1.5 py-0.5 rounded">
                      排除不处理
                    </span>
                  )}
                </div>
              ))}
            </div>
          ) : (
            <div className="flex-1 flex flex-col items-center justify-center p-6 text-center text-slate-400 bg-slate-50/50 space-y-3">
              <FileSpreadsheet className="w-10 h-10 text-slate-300 animate-pulse" />
              <div className="max-w-xs space-y-1">
                <p className="text-xs font-bold text-slate-600">未检测到工作表列表</p>
                <p className="text-[11px] text-slate-400">请确保本地 Excel/WPS 已连通网桥并打开活跃料单，点击上方“刷新 Excel 工作表列表”抓取列表页签。</p>
              </div>
            </div>
          )}
        </div>

        {/* Right Side: Running Dashboard Progress & Logs */}
        <div className="lg:col-span-7 border border-slate-200 rounded-xl overflow-hidden flex flex-col h-[400px]">
          <div className="bg-indigo-950 border-b border-indigo-900 px-4 py-3 flex items-center justify-between shrink-0">
            <div className="flex items-center gap-2 text-white">
              <Activity className="w-4 h-4 text-indigo-400 animate-pulse" />
              <span className="text-xs font-bold">自动化流水线面板 (Pipeline Status)</span>
            </div>
            {batchActive && (
              <div className="flex items-center gap-2">
                <span className={`text-[10px] font-mono font-bold px-2 py-0.5 rounded border ${
                  batchStopping 
                    ? "text-amber-300 bg-amber-950/60 border-amber-800/50" 
                    : "text-emerald-400 bg-emerald-950/50 border-emerald-800/40"
                }`}>
                  {batchStopping ? "STOPPING..." : "RUNNING"}
                </span>
                <button
                  onClick={onStopBatch}
                  disabled={batchStopping}
                  id="btn_stop_batch_header"
                  className="flex items-center gap-1 bg-rose-600/90 hover:bg-rose-600 active:scale-95 text-white text-[11px] font-bold px-2.5 py-1 rounded transition-all cursor-pointer disabled:opacity-50 shadow-xs shadow-rose-950"
                  title="立即停止后续工作表填回"
                >
                  <Square className="w-3 h-3 fill-current" />
                  <span>{batchStopping ? "正在停止..." : "停止填回"}</span>
                </button>
              </div>
            )}
          </div>

          <div className="flex-1 p-5 flex flex-col justify-between overflow-hidden bg-[#070b19]">
            {/* Upper half: Live progress */}
            <div className="space-y-4 shrink-0">
              <div className="flex items-center justify-between text-xs text-indigo-200 font-semibold">
                <span>流水线进度:</span>
                <span className="font-mono text-emerald-400 font-bold">
                  {batchProgress.completedCount} / {batchProgress.totalCount} 工作表 ({progressPercent}%)
                </span>
              </div>

              {/* Progress Bar */}
              <div className="w-full bg-slate-800 rounded-full h-3 overflow-hidden border border-slate-700">
                <div 
                  className="bg-indigo-500 h-full rounded-full transition-all duration-300"
                  style={{ width: `${progressPercent}%` }}
                ></div>
              </div>

              {/* Current Active Step */}
              {batchActive ? (
                <div className="bg-indigo-900/30 border border-indigo-800/50 p-3.5 rounded-lg text-left text-xs text-white flex items-center gap-3">
                  <RefreshCw className="w-4 h-4 text-indigo-400 animate-spin shrink-0" />
                  <div className="space-y-0.5 flex-1">
                    <p className="font-bold flex items-center gap-1.5">
                      <span>正在处理:</span>
                      <span className="text-indigo-300 font-mono underline">{batchProgress.currentSheetName}</span>
                    </p>
                    <p className="text-[11px] text-indigo-200/80">{getStepDescription(batchProgress.step)}</p>
                  </div>
                </div>
              ) : (
                <div className="bg-slate-900/40 border border-slate-800/85 p-3.5 rounded-lg text-left text-xs text-slate-400 italic">
                  等待启动批量填回。请勾选左侧的工作表，然后点击下方启动按钮。
                </div>
              )}
            </div>

            {/* Middle: Log Output Area */}
            <div className="flex-1 overflow-y-auto my-4 bg-slate-950/60 p-3.5 rounded-lg border border-slate-800 font-mono text-[10.5px] text-slate-300 space-y-2 leading-relaxed">
              <div className="text-slate-550 border-b border-slate-800 pb-1.5 uppercase font-bold tracking-wider text-[9px] flex justify-between">
                <span>[批处理日志跟踪]</span>
                <span>{batchProgress.logs.length} 条记录</span>
              </div>
              
              {batchProgress.logs.length === 0 ? (
                <div className="text-slate-600 italic py-6 text-center">暂无日志。启动后会在这里显示每张图纸回写报告。</div>
              ) : (
                batchProgress.logs.map((log, idx) => (
                  <div key={idx} className="flex items-start gap-2 text-left">
                    <span className="text-slate-500">[{idx + 1}]</span>
                    <span className="font-bold font-sans text-[11px] min-w-[100px] text-slate-200">
                      [{log.sheetName}]
                    </span>
                    {log.status === "success" ? (
                      <span className="text-emerald-400 font-bold">✔ 填回成功</span>
                    ) : log.status === "stopped" ? (
                      <span className="text-amber-400 font-bold">⏸ 已停止</span>
                    ) : (
                      <span className="text-rose-400 font-bold">❌ 填回失败</span>
                    )}
                    <span className="text-slate-400 border-l border-slate-800 pl-2">
                      {log.message}
                    </span>
                  </div>
                ))
              )}
            </div>

            {/* Lower half: Trigger Action */}
            <div className="pt-2 shrink-0 border-t border-slate-800/50 flex flex-col sm:flex-row justify-between items-center gap-3">
              <div className="text-[10px] text-slate-500 font-bold text-left self-start sm:self-center">
                已勾选待填回: <span className="font-mono text-indigo-400 text-xs font-bold">{sheetsToProcess.length}张工作表</span> / {excelSheets.length}个总类
              </div>

              <div className="flex items-center gap-2.5 w-full sm:w-auto">
                {batchActive ? (
                  <button
                    onClick={onStopBatch}
                    disabled={batchStopping}
                    id="btn_stop_batch"
                    className="w-full sm:w-auto bg-rose-600 hover:bg-rose-700 active:scale-[0.98] disabled:opacity-50 text-white text-xs font-extrabold px-6 py-3 rounded-lg transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-rose-900/20 shrink-0 select-none uppercase tracking-wider"
                  >
                    <Square className="w-4 h-4 fill-white" />
                    <span>{batchStopping ? "正在停止填回..." : "停止填回"}</span>
                  </button>
                ) : (
                  <button
                    disabled={!wsConnected || sheetsToProcess.length === 0}
                    onClick={handleTriggerBatch}
                    id="btn_start_batch"
                    className="w-full sm:w-auto bg-indigo-600 hover:bg-indigo-700 hover:scale-[1.02] disabled:opacity-40 disabled:scale-100 disabled:hover:bg-indigo-600 active:scale-[0.98] text-white text-xs font-extrabold px-6 py-3 rounded-lg transition-all flex items-center justify-center gap-2 cursor-pointer shadow-md shadow-indigo-900/10 shrink-0 select-none uppercase tracking-wider"
                  >
                    <Play className="w-4 h-4" />
                    <span>一键启动批量导入填回</span>
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
