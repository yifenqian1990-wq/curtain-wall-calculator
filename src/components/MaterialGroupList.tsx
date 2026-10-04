import React, { useState } from "react";
import { ProcessedItem, ColumnMappings, MaterialType } from "../types";
import { Layers, ChevronDown, ChevronUp, Edit2, AlertCircle, RotateCcw, Link } from "lucide-react";

interface MaterialGroupListProps {
  items: ProcessedItem[];
  mappings: ColumnMappings;
  availableColumns: string[];
  onUpdateMappings: (mappings: ColumnMappings) => void;
  onUpdateItemCategory: (rowIndex: number, newCategory: MaterialType) => void;
}

export default function MaterialGroupList({
  items,
  mappings,
  availableColumns,
  onUpdateMappings,
  onUpdateItemCategory,
}: MaterialGroupListProps) {
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({
    other: true,
  });

  const toggleGroup = (group: string) => {
    setCollapsedGroups(prev => ({ ...prev, [group]: !prev[group] }));
  };

  const GROUP_TITLES: Record<MaterialType, string> = {
    profile: "型材材料 (龙骨 / 立柱 / 横梁)",
    panel: "面板材料 (玻璃 / 保温板 / 铝单板 / 石材)",
    steel: "钢材构件 (钢角码 / 节点锚件 / 钢垫块 / 支撑)",
    gasket: "胶条与密封条 (EPDM胶条 / 气密隔离条)",
    fastener: "精密紧固件 (固定自攻钉 / 挂件螺母 / 膨胀栓)",
    auxiliary: "辅料材料 (耐候密封胶 / 止水海绵 / 尼龙隔震垫)",
    other: "其它材料或杂项 (说明行 / 表头 / 未识别)"
  };

  const getGroupBadgeColor = (group: MaterialType) => {
    switch (group) {
      case 'profile': return 'bg-cyan-50 border-cyan-100 text-cyan-800';
      case 'panel': return 'bg-indigo-50 border-indigo-100 text-indigo-800';
      case 'steel': return 'bg-orange-50 border-orange-100 text-orange-800';
      case 'gasket': return 'bg-emerald-50 border-emerald-100 text-emerald-800';
      case 'fastener': return 'bg-amber-50 border-amber-100 text-amber-800';
      case 'auxiliary': return 'bg-blue-50 border-blue-100 text-blue-800';
      default: return 'bg-gray-50 border-gray-100 text-gray-500';
    }
  };

  // Group items
  const groupedItems: Record<MaterialType, ProcessedItem[]> = {
    profile: [],
    panel: [],
    steel: [],
    gasket: [],
    fastener: [],
    auxiliary: [],
    other: []
  };

  items.forEach(item => {
    groupedItems[item.category].push(item);
  });

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="elements_list">
      <div className="flex flex-col xl:flex-row justify-between items-start xl:items-center mb-6 gap-4 border-b border-slate-100 pb-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Layers className="w-5 h-5 text-indigo-600" />
            工程料单材料提取核算明细 (型材、钢件与面板)
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            系统已依据您填写的 Excel 坐标引用区间「直接读取、过滤并渲染」活跃工作表的条目。型材组、钢件组与面板组数据列表如下。
          </p>
        </div>
      </div>

      {/* Accordion list of groups */}
      <div className="space-y-4">
        {(["profile", "steel", "panel"] as const).map(groupKey => {
          const groupItems = groupedItems[groupKey];
          const isCollapsed = collapsedGroups[groupKey] || false;
          const badgeColor = getGroupBadgeColor(groupKey);

          return (
            <div key={groupKey} className="border border-slate-200 rounded overflow-hidden shadow-xs">
              {/* Accordion Hook */}
              <div
                onClick={() => toggleGroup(groupKey)}
                className="bg-slate-50 hover:bg-slate-100/90 px-5 py-3.5 flex items-center justify-between cursor-pointer transition-colors border-b border-slate-150"
              >
                <div className="flex items-center gap-3">
                  <span className={`px-2.5 py-0.5 text-[10px] font-bold rounded border ${badgeColor}`}>
                    {groupItems.length} 行数据
                  </span>
                  <span className="font-bold text-slate-800 text-sm tracking-tight">{GROUP_TITLES[groupKey]}</span>
                </div>
                {isCollapsed ? <ChevronDown className="w-4 h-4 text-slate-400" /> : <ChevronUp className="w-4 h-4 text-slate-400" />}
              </div>

              {!isCollapsed && (
                <div className="overflow-x-auto">
                  <table className="min-w-full divide-y divide-slate-200 text-xs text-left">
                    <thead className="bg-slate-50/80 text-slate-500 font-bold tracking-wider text-[10px] uppercase">
                      <tr>
                        <th className="px-4 py-2.5 w-12 border-b border-slate-200">行号</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">名称</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">加工图号</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">材质及型号</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">下料尺寸 (L) (mm)</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">单樘用量</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">总计</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">单位</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">L1</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">L2</th>
                        <th className="px-4 py-2.5 border-b border-slate-200">备注</th>
                        <th className="px-4 py-2.5 text-center w-32 border-b border-slate-200">算料逻辑归属</th>
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-slate-100 text-slate-700">
                      {groupItems.length === 0 ? (
                        <tr>
                          <td colSpan={12} className="px-4 py-8 text-center text-slate-400 italic">
                            当前核算类别中暂无相关的材料清单条目
                          </td>
                        </tr>
                      ) : (
                        groupItems.map(item => (
                          <tr key={item.rowIndex} className="hover:bg-indigo-50/20 transition-colors">
                            <td className="px-4 py-2.5 font-mono text-slate-400">
                              {item.rowIndex + 1}
                            </td>
                            <td className="px-4 py-2.5 font-bold text-slate-900 border-r border-slate-100">
                              {item.name || <span className="text-slate-300 italic">(空)</span>}
                            </td>
                            <td className="px-4 py-2.5 font-mono text-indigo-700 bg-slate-50/50">
                              {item.drawingNo || <span className="text-slate-300 italic">—</span>}
                            </td>
                            <td className="px-4 py-2.5 text-slate-600 font-medium bg-slate-50/40">
                              {item.materialModel || <span className="text-slate-300 italic">—</span>}
                            </td>
                            <td className="px-4 py-2.5 font-mono text-rose-600 font-bold bg-rose-50/10">
                              {item.cuttingSize || <span className="text-slate-300 italic">—</span>}
                            </td>
                            <td className="px-4 py-2.5 font-mono text-slate-600">
                              {item.unitUsage || <span className="text-slate-300 italic">—</span>}
                            </td>
                            <td className="px-4 py-2.5 font-mono font-bold text-slate-900 bg-amber-50/20">
                              {item.qty}
                            </td>
                            <td className="px-4 py-2.5 font-semibold text-slate-500">
                              {item.unit || "—"}
                            </td>
                            <td className="px-4 py-2.5 font-mono text-slate-500">
                              {item.l1 || "—"}
                            </td>
                            <td className="px-4 py-2.5 font-mono text-slate-500">
                              {item.l2 || "—"}
                            </td>
                            <td className="px-4 py-2.5 text-slate-500 truncate max-w-[150px]" title={item.remark}>
                              {item.remark || "—"}
                            </td>
                            <td className="px-4 py-2.5 text-center">
                              <select
                                value={item.category}
                                onChange={e => onUpdateItemCategory(item.rowIndex, e.target.value as MaterialType)}
                                className="bg-white border border-slate-200 rounded p-1 text-[10px] font-semibold text-slate-700 focus:outline-none focus:border-indigo-500 focus:ring-1 focus:ring-indigo-500 cursor-pointer"
                              >
                                <option value="profile">型材材料</option>
                                <option value="panel">面板材料</option>
                                <option value="steel">钢材锚件</option>
                                <option value="other">忽略不归类</option>
                              </select>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
