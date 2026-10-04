import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import ExcelComConnector from "./components/ExcelComConnector";
import ExcelRangeSelector from "./components/ExcelRangeSelector";
import AuxMaterialLedger from "./components/AuxMaterialLedger";
import AuxCalculationResult from "./components/AuxCalculationResult";
import MaterialGroupList from "./components/MaterialGroupList";
import BatchImportFillback from "./components/BatchImportFillback";
import DataSummaryModule from "./components/DataSummaryModule";
import { 
  ColumnMappings, ProcessedItem, MaterialType, 
  RangeConfiguration, COMConnectionState, AuxiliaryLedgerItem 
} from "./types";
import { classifyRowHeuristic, parseExcelRange, evaluateFormula, sanitizeExcelValue, resolveRangeByBColumnKeyword, simplifyFormulaDisplay, isAuxiliaryItem } from "./utils/calcEngine";
import { Layers, FileSpreadsheet, Sparkles, HelpCircle, CheckCircle2, Server, Settings, Calculator, BookOpen, RefreshCw, ChevronDown, ChevronUp, Link, Unlink, BarChart3 } from "lucide-react";

// Initial state for Module 3 auxiliary materials pre-set template
const DEFAULT_LEDGER: AuxiliaryLedgerItem[] = [
  {
    id: "l-gasket-jt19",
    category: "gasket",
    name: "H2-JT19主龙骨密封胶条",
    drawingNo: "H2-JT19",
    materialModel: "三元乙丙 EPDM 开槽嵌条",
    size: "15mm宽 槽深6",
    unit: "米",
    remark: "根据横梁立柱型材单重等长配套",
    position: "立柱护边嵌槽",
    rules: [
      {
        id: "rule-g-1",
        name: "胶条H2-JT19长度等比配套规则(1.0倍)",
        expression: "SUM_LEN('H2-JT19') * 1.0",
        isActive: true,
        description: "按型材H2-JT19的主龙骨下料长度1.0倍计算配套胶条",
        lengthAdjustment: 50
      }
    ]
  },
  {
    id: "l-gasket-jt04",
    category: "gasket",
    name: "H2-JT04中空玻璃室内外双侧胶条",
    drawingNo: "H2-JT04",
    materialModel: "三元乙丙高回弹中空空腔条",
    size: "12*10 黑色双向",
    unit: "米",
    remark: "打胶槽内外一式两份，自动翻倍算米数",
    position: "玻璃腔外侧拼条",
    rules: [
      {
        id: "rule-g-2",
        name: "内外双向胶条2.0倍配套公式",
        expression: "SUM_LEN('H2-JT04') * 2.0",
        isActive: true,
        description: "室内和室外拼缝处双侧排布，长度进行2.0倍套算",
        lengthAdjustment: 50
      }
    ]
  },
  {
    id: "l-fastener-screw",
    category: "fastener",
    name: "盘头十字不锈钢自攻钉",
    drawingNo: "ST4.8*13",
    materialModel: "不锈钢 SUS304 钻尾",
    size: "M4.8x13-14",
    unit: "只",
    remark: "固定主次接口龙骨角码：单角码配4只",
    position: "龙骨交界承载角码",
    rules: [
      {
        id: "rule-f-1",
        name: "固定角码自攻钉(4只/角码)",
        expression: "SUM_QTY('Q235-K11') * 4",
        isActive: true,
        description: "每个连接用角码，固定需配备4只自攻螺钉"
      }
    ]
  },
  {
    id: "l-aux-sealant",
    category: "auxiliary",
    name: "幕墙单组份硅酮耐候密封胶",
    drawingNo: "耐候密封胶",
    materialModel: "硅酮中性耐候防霉35级",
    size: "300ml 优质支装",
    unit: "支",
    remark: "外层周缝密封打胶量：估算因子0.15L/m",
    position: "室外层玻璃与型材搭接",
    rules: [
      {
        id: "rule-a-1",
        name: "耐候胶按玻璃面板周长配套估算",
        expression: "SUM_PERIMETER('玻璃') * 0.15",
        isActive: true,
        description: "按中空玻璃周长、缝深与缝宽估算用量 (体积因子0.15L/米)"
      }
    ]
  },
  {
    id: "l-aux-insulation",
    category: "auxiliary",
    name: "龙骨空心防火保温岩棉块",
    drawingNo: "岩棉",
    materialModel: "超细绝热防火玻璃棉",
    size: "100mm厚 100Kg/m3",
    unit: "㎡",
    remark: "根据配套大平面中空保温区域等比覆盖",
    position: "主龙骨防火封层",
    rules: [
      {
        id: "rule-a-2",
        name: "岩棉按玻璃面板面积覆盖估算(损耗1.05)",
        expression: "SUM_AREA('玻璃') * 1.05",
        isActive: true,
        description: "按玻璃采光面板尺寸计岩棉防火防护铺设总面积(损耗系数1.05)"
      }
    ]
  }
];

const DEFAULT_RAW_ROWS: string[][] = [
  // Row 1 to 5 (General engineering metadata headers)
  ["", "", "单元板块料单明细表 (项目部示范数据)", "", "", "", "", "", "", "", "", ""],
  ["", "", "项目名称: 示例超高层幕墙幕结构工程", "", "", "", "编制日期: 2026-06-06", "", "", "", "", ""],
  ["", "", "板块型号: UNIT-140-1500 (标准层板块)", "", "", "", "填料审核: OLE_COM 桥接写入", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "【型材导入段 (配置: C8:L12)】", "", "", "", "【钢卡导入段 (配置: C16:L19)】", "", "", "【玻璃面板段 (配置: C22:L25)】", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  // Row 7 (Index 6, matching Title Headers)
  ["大类", "序号", "名称", "加工图号", "材质及型号", "下料尺寸 (L) (mm)", "单樘用量", "总计", "单位", "L1", "L2", "备注"],
  // Row 8 (Index 7, Profile Item 1)
  ["型材", "1", "公立柱", "H2-JT01", "铝合金6063-T6", "2620", "1", "12", "支", "2620", "", "左侧立柱主龙骨"],
  // Row 9 (Index 8, Profile Item 2)
  ["型材", "2", "母立柱", "H2-JT02", "铝合金6063-T6", "3940", "1", "8", "支", "3940", "", "右侧立柱主龙骨"],
  // Row 10 (Index 9, Profile Item 3)
  ["型材", "3", "上横梁", "H2-JT03", "铝合金6063-T6", "1450", "2", "16", "支", "", "", "顶部横向框"],
  // Row 11 (Index 10, Profile Item 4)
  ["型材", "4", "下横梁", "H2-JT04", "铝合金6063-T6", "1450", "2", "16", "支", "", "", "底部横向框"],
  // Row 12 (Index 11, Profile Item 5)
  ["型材", "5", "中间横梁", "H2-JT05", "铝合金6063-T6", "1420", "1", "8", "支", "", "", "加强层框体"],
  // Row 13 to 15 (Spacers)
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  // Row 16 (Index 15, Steel Item 1)
  ["钢件", "1", "热浸镀锌钢牛腿", "Q235-K11", "Q235B", "300*120", "2", "16", "个", "", "", "主锚固角码"],
  // Row 17 (Index 16, Steel Item 2)
  ["钢件", "2", "固定挂件组件", "SUS304-T2", "不锈钢304", "", "4", "12", "套", "", "", "挂扣式定位挂件"],
  // Row 18 (Index 17, Steel Item 3)
  ["钢件", "3", "底座防震钢垫片", "Q235-DP3", "Q235B", "100*100*6", "4", "32", "片", "", "", "垫块找平"],
  // Row 19 (Index 18, Steel Item 4)
  ["钢件", "4", "龙骨槽钢支架", "SUS304-Z1", "不锈钢304", "", "1", "8", "支", "", "", "横梁连墙支架"],
  // Row 20 to 21 (Spacers)
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  // Row 22 (Index 21, Panel Item 1)
  ["面板", "1", "中空钢化胶片玻璃", "GL-01", "6+12A+6", "1200*1500", "1", "24", "㎡", "1200", "1500", "主采光面透光玻璃"],
  // Row 23 (Index 22, Panel Item 2)
  ["面板", "2", "背衬保温岩棉块", "IN-01", "防火高密度岩棉", "100mm厚", "1", "15", "㎡", "", "", "防火横梁区密封岩棉"],
  // Row 24 (Index 23, Panel Item 3)
  ["面板", "3", "层间磨砂防爆玻璃", "GL-02", "8mm超白钢化", "1200*800", "1", "12", "㎡", "1200", "800", "层间封孔腰线"],
  // Row 25 (Index 24, Panel Item 4)
  ["面板", "4", "副框铝单板", "AL-01", "铝单板3.0mm厚", "1180*450", "1", "12", "㎡", "1180", "450", "边侧收口装饰"],
  ["", "", "", "", "", "", "", "", "", "", "", ""]
];

const DEFAULT_CORNER_RAW_ROWS: string[][] = [
  ["", "", "转角单元板块料单明细表 (D02转角单元)", "", "", "", "", "", "", "", "", ""],
  ["", "", "项目名称: 示例超高层幕墙幕结构工程", "", "", "", "编制日期: 2026-06-06", "", "", "", "", ""],
  ["", "", "板块型号: UNIT-180-CORNER (转角层板块)", "", "", "", "填料审核: OLE_COM 桥接写入", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "【型材导入段 (配置: C8:L12)】", "", "", "", "【钢卡导入段 (配置: C16:L19)】", "", "", "【玻璃面板段 (配置: C22:L25)】", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["大类", "序号", "名称", "加工图号", "材质及型号", "下料尺寸 (L) (mm)", "单樘用量", "总计", "单位", "L1", "L2", "备注"],
  ["型材", "1", "公立柱", "H2-JT01", "铝合金6063-T6", "2620", "1", "6", "支", "2620", "", "左侧立柱主龙骨"],
  ["型材", "2", "转角主立柱", "H2-JT08", "铝合金6063-T6", "3940", "1", "6", "支", "3940", "", "90度转角特制主柱"],
  ["型材", "3", "上横梁", "H2-JT03", "铝合金6063-T6", "1450", "2", "12", "支", "", "", "顶部横向框"],
  ["型材", "4", "下横梁", "H2-JT04", "铝合金6063-T6", "1450", "2", "12", "支", "", "", "底部横向框"],
  ["型材", "5", "拐角收口柱", "H2-JT09", "铝合金6063-T6", "2620", "1", "6", "支", "2620", "", "转角外饰扣槽型材"],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["钢件", "1", "热浸镀锌钢牛腿", "Q235-K11", "Q235B", "300*120", "2", "12", "个", "", "", "主锚固角码"],
  ["钢件", "2", "固定挂件组件", "SUS304-T2", "不锈钢304", "", "4", "8", "套", "", "", "挂扣式定位挂件"],
  ["钢件", "3", "底座防震钢垫片", "Q235-DP3", "Q235B", "100*100*6", "4", "24", "片", "", "", "垫块找平"],
  ["钢件", "4", "龙骨槽钢支架", "SUS304-Z1", "不锈钢304", "", "1", "6", "支", "", "", "横梁连墙支架"],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["面板", "1", "中空钢化胶片玻璃", "GL-01", "6+12A+6", "1200*1500", "1", "18", "㎡", "1200", "1500", "主采光面透光玻璃"],
  ["面板", "2", "背衬保温岩棉块", "IN-01", "防火高密度岩棉", "100mm厚", "1", "12", "㎡", "", "", "防火横梁区密封岩棉"],
  ["面板", "3", "层间磨砂防爆玻璃", "GL-02", "8mm超白钢化", "1200*800", "1", "8", "㎡", "1200", "800", "层间封孔腰线"],
  ["面板", "4", "副框铝单板", "AL-01", "铝单板3.0mm厚", "1180*450", "1", "8", "㎡", "1180", "450", "边侧收口装饰"],
  ["", "", "", "", "", "", "", "", "", "", "", ""]
];

const buildDefaultWorkbook = () => {
  const wb = XLSX.utils.book_new();

  // Sheet 0: Catalog Index (目录)
  const catalogRows = [
    ["", "工程图纸与单元细目表索引目录", "", "", ""],
    ["序号", "单元图号", "单元类型描述", "工程数量", "备注"],
    ["1", "D01-标准1500单元细目表", "1500mm宽标准幕墙单元", "12 樘", "主楼标准层"],
    ["2", "D02-转角1800单元细目表", "1800mm宽90度转角单元", "6 樘", "主楼四角转角层"],
  ];
  const wsIndex = XLSX.utils.aoa_to_sheet(catalogRows);
  XLSX.utils.book_append_sheet(wb, wsIndex, "00-工程图纸目录");

  // Sheet 1: D01 Standard Unit
  const ws1 = XLSX.utils.aoa_to_sheet(DEFAULT_RAW_ROWS);
  ws1['!cols'] = [
    { wch: 10 }, { wch: 8 }, { wch: 18 }, { wch: 15 }, { wch: 20 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 25 }
  ];
  XLSX.utils.book_append_sheet(wb, ws1, "D01-标准1500单元细目表");

  // Sheet 2: D02 Corner Unit
  const ws2 = XLSX.utils.aoa_to_sheet(DEFAULT_CORNER_RAW_ROWS);
  ws2['!cols'] = [
    { wch: 10 }, { wch: 8 }, { wch: 18 }, { wch: 15 }, { wch: 20 }, { wch: 22 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 25 }
  ];
  XLSX.utils.book_append_sheet(wb, ws2, "D02-转角1800单元细目表");

  return wb;
};

export default function App() {
  const [activeTab, setActiveTab ] = useState<string>("bridge");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetName, setSheetName] = useState<string>("");
  const [rawRows, setRawRows] = useState<string[][]>([]);
  
  // Batch processing states
  const [batchActive, setBatchActive] = useState<boolean>(false);
  const [batchStopping, setBatchStopping] = useState<boolean>(false);
  const stopBatchRef = React.useRef<boolean>(false);
  const [batchProgress, setBatchProgress] = useState<{
    currentSheetName: string;
    step: 'idle' | 'reading' | 'calculating' | 'writing' | 'success' | 'error' | 'stopping';
    completedCount: number;
    totalCount: number;
    logs: Array<{ sheetName: string; status: 'success' | 'failed' | 'stopped'; message: string }>;
  }>({
    currentSheetName: "",
    step: 'idle',
    completedCount: 0,
    totalCount: 0,
    logs: []
  });

  const wsResolveRef = React.useRef<((data: any) => void) | null>(null);
  const wsWriteResolveRef = React.useRef<{ resolve: (val: any) => void; reject: (err: Error) => void; resetTimeout?: () => void } | null>(null);
  const wsSelectionResolveRef = React.useRef<((data: any) => void) | null>(null);

  const handleFetchSelection = () => {
    return new Promise<{ success: boolean; address: string; rows: string[][] }>((resolve, reject) => {
      if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
        reject(new Error("本地 Python OLE COM 网桥尚未连通！请在第 1 模块【COM网桥配对】中核对连接状态。"));
        return;
      }
      
      let timeoutTimer: NodeJS.Timeout | null = null;
      
      wsSelectionResolveRef.current = (data) => {
        if (timeoutTimer) clearTimeout(timeoutTimer);
        if (data.success) {
          resolve({
            success: true,
            address: data.address,
            rows: data.rows
          });
        } else {
          reject(new Error(data.message || "未能成功抓取 Excel 选中区域，请检查单元格是否正处于编辑状态。"));
        }
      };
      
      timeoutTimer = setTimeout(() => {
        if (wsSelectionResolveRef.current) {
          wsSelectionResolveRef.current = null;
          reject(new Error("读取 Excel 选中区域超时 (10秒)。请在 Excel/WPS 中按下 Enter 或 Esc 退出单元格编辑模式，然后再度尝试获取！"));
        }
      }, 10000);
      
      localWs.send(JSON.stringify({
        action: "read_selection"
      }));
    });
  };

  // Python COM Syncing States
  const [localWsPort, setLocalWsPort] = useState<number>(8001);
  const [wsConnected, setWsConnected] = useState<boolean>(false);
  const [localWs, setLocalWs] = useState<WebSocket | null>(null);
  const [comBridgeStatus, setComBridgeStatus] = useState<string>("disconnected");
  const [comBridgeMessage, setComBridgeMessage] = useState<string>("等待开启本地 Python OLE COM 网桥...");
  
  // Direct Excel OLE Pulling States
  const [pullLoading, setPullLoading] = useState<boolean>(false);
  const [pullError, setPullError] = useState<string | null>(null);
  const [importStatusMessage, setImportStatusMessage] = useState<string | null>(null);
  const [excelSheets, setExcelSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [isDocOpen, setIsDocOpen] = useState(false);

  // Dynamic column mapping indicators
  const [mappings, setMappings] = useState<ColumnMappings>({
    categoryCol: null,
    nameCol: 2,       // Col C
    specCol: 3,       // Col D
    materialCol: 4,   // Col E
    lenCol: 5,        // Col F
    qtyCol: 7,        // Col H
    unitCol: 8,       // Col I
    remarkCol: 11,    // Col L
  });

  // State configurations for the 5 User Modules
  // Module 1 state
  const [comConnection, setComConnection] = useState<COMConnectionState>({
    isConnected: false,
    provider: 'wps',
    host: 'localhost',
    port: 3000,
    activeWorkbookName: '单元加工材料细目表_标准模板.xlsx',
    isSending: false,
  });
  const [syncStatus, setSyncStatus] = useState<'idle' | 'linking' | 'sending' | 'success'>('idle');
  const [syncProgress, setSyncProgress] = useState<number>(0);

  // Module 2 state: Range-based coordinate references
  const [ranges, setRanges] = useState<RangeConfiguration>({
    profileRange: 'C8:L12', // Matches profiles row 8 to row 12 in DEFAULT_RAW_ROWS (0-indexed 7-11)
    steelRange: 'C16:L19',   // Matches steel row 16 to row 19 (0-indexed 15-18)
    panelRange: 'C22:L25',   // Matches panel row 22 to row 25 (0-indexed 21-24)
    profileKeyword: '铝  型  材',
    steelKeyword: '钢   件',
    panelKeyword: '面   板',
  });

  // Module 3 state: Master ledger targets list
  // Module 3 ledger targets list
  const [ledger, setLedger] = useState<AuxiliaryLedgerItem[]>([]);
  const [ledgerHistory, setLedgerHistory] = useState<Array<{
    id: string;
    timestamp: string;
    description: string;
    itemCount: number;
    data: AuxiliaryLedgerItem[];
  }>>([]);

  const updateLedgerWithHistory = (
    newLedgerOrFn: AuxiliaryLedgerItem[] | ((prev: AuxiliaryLedgerItem[]) => AuxiliaryLedgerItem[]),
    description: string
  ) => {
    setLedger(prev => {
      const resolvedNew = typeof newLedgerOrFn === 'function' ? newLedgerOrFn(prev) : newLedgerOrFn;
      
      // Save snapshot of current (prev) ledger before the modification
      const timestamp = new Date().toLocaleTimeString('zh-CN', { hour12: false });
      setLedgerHistory(prevHistory => {
        const newHistoryItem = {
          id: `hist-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
          timestamp,
          description,
          itemCount: prev.length,
          data: [...prev]
        };
        const updatedHistory = [newHistoryItem, ...prevHistory];
        return updatedHistory.slice(0, 5); // Keep recent 5
      });

      return resolvedNew;
    });
  };

  const handleRestoreHistory = (historyId: string) => {
    const record = ledgerHistory.find(h => h.id === historyId);
    if (!record) return;

    const timestamp = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    
    setLedgerHistory(prevHistory => {
      const newHistoryItem = {
        id: `hist-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
        timestamp,
        description: `恢复历史版本: ${record.description}`,
        itemCount: ledger.length,
        data: [...ledger]
      };
      const updatedHistory = [newHistoryItem, ...prevHistory.filter(h => h.id !== historyId)];
      return updatedHistory.slice(0, 5);
    });

    setLedger(record.data);
  };

  const [projectVars, setProjectVars] = useState<Record<string, any>>({ W: 0, H: 0, W1: 0, plateNo: "", plateCode: "", plateModel: "", sheetName: "" });

  const [items, setItems] = useState<ProcessedItem[]>([]);
  const [availableColumns, setAvailableColumns] = useState<string[]>([]);
  const [aiResponse, setAiResponse] = useState<any | null>(null);
  const [userDisconnected, setUserDisconnected] = useState<boolean>(false);

  // Connect to the local OLE/COM Python WebSocket Server
  useEffect(() => {
    if (userDisconnected) {
      setWsConnected(false);
      setComBridgeStatus("disconnected");
      setComBridgeMessage("已手动断开网桥连接。点击右上方【连接网桥】按钮可重新激活本地双向 COM 网桥并一键联动。");
      setComConnection(prev => ({
        ...prev,
        isConnected: false,
        isSending: false
      }));
      return;
    }

    let active = true;
    let socket: WebSocket | null = null;
    let reconnectTimer: NodeJS.Timeout | null = null;

    function connect() {
      if (!active || userDisconnected) return;
      
      setWsConnected(false);
      setComBridgeStatus("disconnected");
      setComBridgeMessage(`正在向本机 ws://127.0.0.1:${localWsPort} 建立 WebSocket 握手...`);

      try {
        socket = new WebSocket(`ws://127.0.0.1:${localWsPort}`);
        setLocalWs(socket);

        socket.onopen = () => {
          if (!active || userDisconnected) return;
          setWsConnected(true);
          setComBridgeStatus("idle");
          setComBridgeMessage("本地 Python OLE/COM 双向网桥已成功配对并连接！等待在网页端一键触发一键填料写回。");
          setComConnection(prev => ({
            ...prev,
            isConnected: true
          }));
        };

        socket.onmessage = (event) => {
          if (!active) return;
          try {
            const data = JSON.parse(event.data);
            if (data.type === "progress") {
              setSyncProgress(data.progress);
              setComBridgeMessage(`正在精准回填单元格... [${data.current} / ${data.total}] (${data.progress}%)`);
              // 收到写入进度回报，重置超时计时器，防止大批量写入时前端提前误判超时
              if (wsWriteResolveRef.current?.resetTimeout) {
                wsWriteResolveRef.current.resetTimeout();
              }
            }
            
            if (data.type === "status") {
              setComBridgeStatus(data.status);
              setComBridgeMessage(data.message);
              
              if (data.status === "error") {
                setPullLoading(false);
                setSyncStatus("idle");
                if (wsWriteResolveRef.current) {
                  wsWriteResolveRef.current.reject(new Error(data.message || "写入错误"));
                  wsWriteResolveRef.current = null;
                }
              }
              
              const activeLine = data.status === "writing" || data.status === "success" || data.status === "idle" || data.status === "reading";
              setComConnection(prev => ({
                ...prev,
                isConnected: activeLine,
                isSending: data.status === "writing"
              }));
              
              if (data.status === "success") {
                setSyncStatus("success");
                setSyncProgress(100);
                if (wsWriteResolveRef.current) {
                  wsWriteResolveRef.current.resolve(data);
                  wsWriteResolveRef.current = null;
                }
                setTimeout(() => {
                  setSyncStatus('idle');
                  setSyncProgress(0);
                }, 3000);
              }
              if (data.status === "writing") {
                setSyncStatus("sending");
                setSyncProgress(0);
              }
            }
            
            if (data.type === "read_response") {
              setPullLoading(false);
              
              const isBatchWaiting = !!wsResolveRef.current;
              if (wsResolveRef.current) {
                wsResolveRef.current(data);
                wsResolveRef.current = null;
              }

              if (data.success && data.rows) {
                try {
                  const sanitizedRows = data.rows.map((row: any[]) =>
                    row.map((cell: any) => sanitizeExcelValue(cell))
                  );
                  const ws = XLSX.utils.aoa_to_sheet(sanitizedRows);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, data.sheetName || "Sheet1");
                  
                  importExcelRows(wb, data.sheetName || "Excel提取物料表", sanitizedRows);

                  if (data.sheetsList && Array.isArray(data.sheetsList)) {
                    setExcelSheets(data.sheetsList);
                    setSelectedSheet(data.sheetName || data.sheetsList[0] || "");
                  }
                } catch (e: any) {
                  if (!isBatchWaiting) {
                    alert(`解析本地 Excel 数据并生成工作表格式出错: ${e.message}`);
                  }
                }
              } else {
                if (!isBatchWaiting) {
                  alert(`从本地 Excel 读取料单失败: ${data.message || "未知错误"}`);
                }
              }
            }

            if (data.type === "selection_response") {
              if (wsSelectionResolveRef.current) {
                wsSelectionResolveRef.current(data);
                wsSelectionResolveRef.current = null;
              }
            }
          } catch (err) {
            console.error("Failed to parse websocket message:", err);
          }
        };

        socket.onclose = () => {
          if (!active) return;
          setWsConnected(false);
          setPullLoading(false);
          setComBridgeStatus("disconnected");
          setComBridgeMessage(`与本机 ws://127.0.0.1:${localWsPort} 的连接已断开。`);
          setComConnection(prev => ({
            ...prev,
            isConnected: false,
            isSending: false
          }));

          if (wsResolveRef.current) {
            wsResolveRef.current({ success: false, message: "网桥连接断开" });
            wsResolveRef.current = null;
          }
          if (wsWriteResolveRef.current) {
            wsWriteResolveRef.current.reject(new Error("网桥连接断开"));
            wsWriteResolveRef.current = null;
          }

          if (!userDisconnected) {
            reconnectTimer = setTimeout(connect, 4000);
          }
        };

        socket.onerror = () => {
          if (socket) socket.close();
        };
      } catch (err) {
        console.error(err);
        if (!userDisconnected) {
          reconnectTimer = setTimeout(connect, 4000);
        }
      }
    }

    connect();

    return () => {
      active = false;
      if (socket) {
        socket.close();
      }
      if (reconnectTimer) {
        clearTimeout(reconnectTimer);
      }
    };
  }, [localWsPort, userDisconnected]);

  // Initialize available columns dropdown labels when raw rows loader completes
  useEffect(() => {
    if (rawRows.length > 0) {
      let bestHeaderRowIndex = 0;
      let maxNonEmpty = 0;

      for (let i = 0; i < Math.min(10, rawRows.length); i++) {
        const row = rawRows[i];
        const count = row.filter(cell => cell.trim() !== "").length;
        if (count > maxNonEmpty) {
          maxNonEmpty = count;
          bestHeaderRowIndex = i;
        }
      }

      const bestHeaderRow = rawRows[bestHeaderRowIndex] || [];
      const labels = bestHeaderRow.map((cell, idx) => {
        const trimmed = cell.trim();
        return trimmed ? `${trimmed} (${String.fromCharCode(65 + idx)}列)` : `未使用列 ${String.fromCharCode(65 + idx)}`;
      });
      setAvailableColumns(labels);
    }
  }, [rawRows]);

  // Extract items from rawRows based strictly on coordinate ranges (e.g. C8:L57)
  useEffect(() => {
    if (rawRows.length === 0) return;

    const pRange = parseExcelRange(ranges.profileRange);
    const sRange = parseExcelRange(ranges.steelRange);
    const paRange = parseExcelRange(ranges.panelRange);

    const processedList: ProcessedItem[] = [];

    for (let rIdx = 0; rIdx < rawRows.length; rIdx++) {
      const row = rawRows[rIdx];
      if (!row || row.filter(c => c && String(c).trim() !== "").length === 0) {
        continue; // Skip empty rows
      }

      const name = sanitizeExcelValue(row[2]);
      const drawingNo = sanitizeExcelValue(row[3]);
      const materialModel = sanitizeExcelValue(row[4]);
      const remark = sanitizeExcelValue(row[11]);

      // Exact coordinate range checks
      let category: MaterialType | null = null;
      if (pRange && rIdx >= pRange.startRow && rIdx <= pRange.endRow) {
        category = 'profile';
      } else if (sRange && rIdx >= sRange.startRow && rIdx <= sRange.endRow) {
        category = 'steel';
      } else if (paRange && rIdx >= paRange.startRow && rIdx <= paRange.endRow) {
        category = 'panel';
      }

      // If a row does not fall within the defined bounds of profiles, steel, or panels, skip it
      if (!category) {
        continue;
      }

      // Skip empty spacer rows or headers within the range if name is blank or is a header label
      if (!name || name === "名称" || name === "品名") {
        continue;
      }
      let cuttingSize = String(sanitizeExcelValue(row[5])).trim();
      if (cuttingSize.startsWith("-") || parseFloat(cuttingSize) < 0) {
        cuttingSize = "";
      }
      const unitUsage = sanitizeExcelValue(row[6]);

      // Parse quantity and lengths safely after applying sanitization
      const qtyStr = sanitizeExcelValue(row[7]);
      const rawQty = qtyStr.replace(/[^\d.]/g, "");
      const qty = parseFloat(rawQty) || 0;

      const unit = sanitizeExcelValue(row[8]);
      const l1 = sanitizeExcelValue(row[9]);
      const l2 = sanitizeExcelValue(row[10]);

      // Compatibility mappings for calculations
      const spec = materialModel;
      const material = materialModel;
      
      // Convert cuttingSize in mm (e.g. "2620") to meters (e.g. 2.62) for compatibility with formula engine (SUM_LEN)
      let len = 0;
      if (cuttingSize) {
        const cleanCutSize = cuttingSize.replace(/[^\d.]/g, "");
        const sizeInMm = parseFloat(cleanCutSize) || 0;
        len = sizeInMm > 0 ? sizeInMm / 1000 : 0;
      }

      processedList.push({
        rowIndex: rIdx,
        originalCategory: category === 'profile' ? '型材' : (category === 'steel' ? '钢件' : '面板'),
        category,
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
        spec,
        material,
        len
      });
    }

    setItems(processedList);
  }, [rawRows, ranges]);

  // High-performance direct OLE COM extraction alignment service - Direct and Fast with No AI
  const importExcelRows = (wb: XLSX.WorkBook, sheet: string, rows: string[][]) => {
    setPullLoading(true);
    setPullError(null);
    
    // Sanitize every cell of the input rows to clear excel errors and float zero noise
    const sanitizedInputRows = rows.map(r => r.map(c => sanitizeExcelValue(c)));

    const getExcelCellVal = (r: number, c: number) => {
      if (!sanitizedInputRows) return 0;
      if (r >= sanitizedInputRows.length) return 0;
      if (c >= sanitizedInputRows[r].length) return 0;
      const val = sanitizedInputRows[r][c] || "";
      const num = parseFloat(String(val).replace(/[^\d.-]/g, ''));
      return isNaN(num) ? 0 : num;
    }

    // Extract plateNo from sheet name (e.g. "D01-标准1500单元细目表" -> "D01")
    const sheetMatch = sheet.match(/^([A-Za-z0-9]+)/);
    const parsedPlateNo = sheetMatch ? sheetMatch[1] : sheet;

    // Extract plateModel from cell C3 (Row 2, Col 2) which is usually "板块型号: UNIT-140-1500 (标准层板块)"
    let parsedPlateModel = "";
    if (sanitizedInputRows && sanitizedInputRows[2] && sanitizedInputRows[2][2]) {
      const cellC3 = String(sanitizedInputRows[2][2]);
      const modelMatch = cellC3.match(/(?:型号|编号)\s*[:：]\s*([A-Za-z0-9_\-]+)/i);
      if (modelMatch) {
        parsedPlateModel = modelMatch[1];
      } else {
        parsedPlateModel = cellC3.replace(/^(板块型号|板块编号|型号|编号)\s*[:：]\s*/, "").trim();
      }
    }

    setProjectVars({
      W: getExcelCellVal(3, 7) / 1000,  // H4
      H: getExcelCellVal(3, 8) / 1000,  // I4
      W1: getExcelCellVal(1, 12) / 1000, // M2
      panelsCount: getExcelCellVal(2, 10), // K3
      plateNo: parsedPlateNo,
      plateCode: parsedPlateModel || parsedPlateNo,
      plateModel: parsedPlateModel,
      sheetName: sheet
    });

    // Setup immediate fixed column mappings matching original structures
    const finalMappings: ColumnMappings = {
      categoryCol: null,
      nameCol: 2,       // Col C
      specCol: 3,       // Col D
      materialCol: 4,   // Col E
      lenCol: 5,        // Col F
      qtyCol: 7,        // Col H
      unitCol: 8,       // Col I
      remarkCol: 11,    // Col L
    };

    // Auto-detect and locate column B merged/flat cell coordinate offsets
    const ws = wb.Sheets[sheet] || null;
    const pk = ranges.profileKeyword || "铝  型  材";
    const sk = ranges.steelKeyword || "钢   件";
    const pak = ranges.panelKeyword || "面   板";

    const pResolved = resolveRangeByBColumnKeyword(ws, sanitizedInputRows, pk);
    const sResolved = resolveRangeByBColumnKeyword(ws, sanitizedInputRows, sk);
    const paResolved = resolveRangeByBColumnKeyword(ws, sanitizedInputRows, pak);

    let nextProfileRange = ranges.profileRange;
    let nextSteelRange = ranges.steelRange;
    let nextPanelRange = ranges.panelRange;
    const detectMsgs: string[] = [];

    if (pResolved) {
      nextProfileRange = `C${pResolved.startRow + 1}:L${pResolved.endRow + 1}`;
      detectMsgs.push(`型材 C${pResolved.startRow + 1}:L${pResolved.endRow + 1}`);
    }
    if (sResolved) {
      nextSteelRange = `C${sResolved.startRow + 1}:L${sResolved.endRow + 1}`;
      detectMsgs.push(`钢件 C${sResolved.startRow + 1}:L${sResolved.endRow + 1}`);
    }
    if (paResolved) {
      nextPanelRange = `C${paResolved.startRow + 1}:L${paResolved.endRow + 1}`;
      detectMsgs.push(`面板 C${paResolved.startRow + 1}:L${paResolved.endRow + 1}`);
    }

    if (pResolved || sResolved || paResolved) {
      setRanges({
        profileRange: nextProfileRange,
        steelRange: nextSteelRange,
        panelRange: nextPanelRange,
        profileKeyword: pk,
        steelKeyword: sk,
        panelKeyword: pak
      });
    }

    setWorkbook(wb);
    setSheetName(sheet);
    setRawRows(sanitizedInputRows);
    setMappings(finalMappings);
    
    // Automatically capture workbook name into COM state as default target
    setComConnection(prev => ({
      ...prev,
      activeWorkbookName: "幕墙加工明细配套料单_" + sheet + ".xlsx"
    }));

    setPullLoading(false);
    const detectFeedback = detectMsgs.length > 0
      ? `自动识别B列合并单元锁定坐标: ${detectMsgs.join(" | ")}`
      : "未识别到分类，支持人工微调。";
    setImportStatusMessage(`成功拉取 [${sheet}] 工作表（${rows.length}行）。${detectFeedback}`);
    
    setTimeout(() => {
      setImportStatusMessage(null);
    }, 4500);
  };

  const handlePullFromExcel = (overrideSheetName?: string) => {
    if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
      alert("本地 OLE COM 网桥未连通！请配对启动『本地 OLE COM 网桥连接』菜单中的 Python 桥接程序以直接提取料单数据。");
      return;
    }
    const targetSheet = overrideSheetName || selectedSheet;
    setPullLoading(true);
    setPullError(null);
    setComBridgeStatus("reading");
    setComBridgeMessage(targetSheet 
      ? `正在从本地读取指定的 [${targetSheet}] 工作表中，请稍候...` 
      : "正在向本地 Excel/WPS 表格发送拉取命令，抓取当前活动工作表的单元格中..."
    );
    
    try {
      localWs.send(JSON.stringify({
        action: "read_excel",
        sheetName: targetSheet || ""
      }));
    } catch (err: any) {
      setPullLoading(false);
      setPullError(err.message || "发送拉取指令失败");
      setComBridgeStatus("error");
      setComBridgeMessage(`发送读取命令出错: ${err.message}`);
    }
  };

  const handleUpdateItemCategory = (rowIndex: number, newCategory: MaterialType) => {
    setItems(prev => prev.map(item => item.rowIndex === rowIndex ? { ...item, category: newCategory } : item));
  };

  // Module 3 ledger handlers
  const handleAddLedgerItem = (item: AuxiliaryLedgerItem) => {
    updateLedgerWithHistory(prev => [...prev, item], `手动添加: ${item.name}`);
  };

  const handleBulkAddLedgerItems = (newItems: AuxiliaryLedgerItem[]) => {
    updateLedgerWithHistory(prev => {
      // Filter out duplicates with identical names or drawings
      const filtered = prev.filter(p => !newItems.some(n => n.name === p.name && n.drawingNo === p.drawingNo));
      return [...filtered, ...newItems];
    }, `批量导入/载入辅料 (${newItems.length}项)`);
  };

  const handleRemoveLedgerItem = (id: string) => {
    const item = ledger.find(i => i.id === id);
    const itemName = item ? item.name : id;
    updateLedgerWithHistory(prev => prev.filter(i => i.id !== id), `删除物料: ${itemName}`);
  };

  const handleClearLedger = () => {
    updateLedgerWithHistory([], "清空辅材料台账");
  };

  const handleUpdateLedgerItem = (updated: AuxiliaryLedgerItem) => {
    updateLedgerWithHistory(prev => prev.map(item => item.id === updated.id ? updated : item), `编辑修改: ${updated.name}`);
  };

  const handleOverwriteLedgerItems = (newItems: AuxiliaryLedgerItem[]) => {
    updateLedgerWithHistory(newItems, `覆盖导入辅料 (${newItems.length}项)`);
  };

  const handleUpdateMappings = (newMappings: ColumnMappings) => {
    setMappings(newMappings);
  };

  // Module 5 COM bridging trigger using OLE Python Local WebSocket Bridge
  const handleTriggerComWriteBack = (calculatedResultsList: any[]) => {
    if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
      alert("本地 COM 网桥尚未连接，请双击运行 com_bridge.py 并确保状态为『准备就绪』！");
      return;
    }

    setSyncStatus('linking');
    setComConnection(prev => ({ ...prev, isSending: true }));

    try {
      if (!workbook || !sheetName) throw new Error("缺少本地工作表实例，无法进行区域映射计算");
      const wb = Object.assign({}, workbook);
      const ws = wb.Sheets[sheetName];

      const cellUpdates: { r: number, c: number, v: any, bg?: number }[] = [];

      const colStart = 2; // C col
      const colEnd = 11; // L col

      const extractUpdatesForRegion = (keyword: string, category: string) => {
        const range = resolveRangeByBColumnKeyword(ws, rawRows, keyword);
        if (!range) return;

        const targetItems = calculatedResultsList.filter(item => 
          item.category === category && item.isCalculated && (item.calculatedUnitUsage || 0) > 0
        );

        let currentRowOffset = 0;

        for (let r = range.startRow; r <= range.endRow; r++) {
          if (currentRowOffset < targetItems.length) {
            const item = targetItems[currentRowOffset];

            // Mapping: C=Name, D=Drawing, E=Material, F=Size, G=UnitUsage, H=Qty, I=Unit, J=L1, K=L2, L=Remark
            cellUpdates.push({ r, c: colStart, v: item.name || "" });
            cellUpdates.push({ r, c: colStart + 1, v: item.drawingNo || "" });
            cellUpdates.push({ r, c: colStart + 2, v: item.materialModel || "" });
            cellUpdates.push({ r, c: colStart + 3, v: item.cuttingSize || "" });
            cellUpdates.push({ r, c: colStart + 4, v: item.calculatedUnitUsage, bg: 65535 }); 
            cellUpdates.push({ r, c: colStart + 5, v: `=G${r + 1}*K3`, bg: 65535 });
            cellUpdates.push({ r, c: colStart + 6, v: item.unit || "" });
            cellUpdates.push({ r, c: colStart + 7, v: item.l1 || "" });
            cellUpdates.push({ r, c: colStart + 8, v: item.l2 || "" });
            cellUpdates.push({ r, c: colStart + 9, v: item.remark || "" });

            currentRowOffset++;
          } else {
            // 超出算料有效项的剩余行，清空旧数据
            for (let c = colStart; c <= colEnd; c++) {
              cellUpdates.push({ r, c, v: "" });
            }
          }
        }
      };

      extractUpdatesForRegion("胶条", "gasket");
      extractUpdatesForRegion("紧固件", "fastener");
      extractUpdatesForRegion("辅材", "auxiliary");

      if (cellUpdates.length === 0) {
        throw new Error("无法在当前表格中检索到「胶条」「紧固件」「辅材」等区域用于填回。请确保读取了正确的表格区段。");
      }

      localWs.send(JSON.stringify({
        action: "write_cells",
        excelFileName: sheetName || comConnection.activeWorkbookName || "excel_file",
        updates: cellUpdates,
      }));
    } catch (err: any) {
      console.error("Local WS send error or mapping error:", err);
      setSyncStatus('idle');
      setComConnection(prev => ({ ...prev, isSending: false }));
      alert(`填回预处理失败：${err.message}`);
    }
  };

  // Real offline write and workbook file download callback
  const handleExportFinishedExcel = (calculatedResultsList: any[]) => {
    if (!workbook || !sheetName) return;

    try {
      // Create cloned copy
      const wb = { ...workbook };
      const ws = wb.Sheets[sheetName];
      if (!ws) throw new Error("无法连接至数据源表格内存槽");

      let filledCount = 0;

      const fillRegion = (keyword: string, category: string) => {
        const range = resolveRangeByBColumnKeyword(ws, rawRows, keyword);
        if (!range) return;

        // Filter items: only those calculated, belonging to this category, with unit usage > 0
        const targetItems = calculatedResultsList.filter(item => 
          item.category === category && item.isCalculated && (item.calculatedUnitUsage || 0) > 0
        );

        // Assume standard offsets based on B column merged fields
        // Column mapping: C(2)=名称, D(3)=加工图号, E(4)=材质及型号, F(5)=下料尺寸, G(6)=单樘用量, H(7)=总计, I(8)=单位, J(9)=L1, K(10)=L2, L(11)=备注
        const colStart = 2; // C col
        const colEnd = 11; // L col

        let currentRowOffset = 0;

        for (let r = range.startRow; r <= range.endRow; r++) {
          // Clear current row cells
          for (let c = colStart; c <= colEnd; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (ws[cellRef]) {
               ws[cellRef] = { t: "s", v: "" };
            }
          }

          // Fill with target items sequentially
          if (currentRowOffset < targetItems.length) {
            const item = targetItems[currentRowOffset];

            const writeMappings: Array<{ c: number; v: any; t: string; f?: string }> = [
              { c: colStart, v: item.name || "", t: "s" },                          // 名称
              { c: colStart + 1, v: item.drawingNo || "", t: "s" },                 // 加工图号
              { c: colStart + 2, v: item.materialModel || "", t: "s" },             // 材质及型号
              { c: colStart + 3, v: item.cuttingSize || "", t: "s" },               // 下料尺寸
              { c: colStart + 4, v: item.calculatedUnitUsage, t: "n" },             // 单樘用量
              { c: colStart + 5, v: item.calculatedQty, t: "n", f: `G${r + 1}*K3` }, // 总计 (动态公式 =单樘用量*K3)
              { c: colStart + 6, v: item.unit || "", t: "s" },                      // 单位
              { c: colStart + 7, v: item.l1 || "", t: "s" },                        // L1
              { c: colStart + 8, v: item.l2 || "", t: "s" },                        // L2
              { c: colStart + 9, v: item.remark || "", t: "s" },                    // 备注
            ];

            writeMappings.forEach(mapping => {
              const cellRef = XLSX.utils.encode_cell({ r, c: mapping.c });
              if (mapping.f) {
                ws[cellRef] = { t: mapping.t, f: mapping.f, v: mapping.v };
              } else {
                ws[cellRef] = { t: mapping.t, v: mapping.v };
              }
            });

            filledCount++;
            currentRowOffset++;
          }
        }
      };

      fillRegion("胶条", "gasket");
      fillRegion("紧固件", "fastener");
      fillRegion("辅材", "auxiliary");

      // Generate the standalone "辅材配套台账清单" sheet as requested by the user's header specification
      const exportRows = [
        ["名称", "加工图号", "材质及型号", "下料尺寸 (L) (MM)", "单樘用量", "总计", "单位", "L1", "L2", "备注", "算料逻辑归属"]
      ];

      calculatedResultsList.forEach(item => {
        const getFriendlyCategory = (cat: string) => {
          if (cat === 'gasket') return '型材气密胶条类';
          if (cat === 'fastener') return '不锈钢紧固五金螺栓类';
          if (cat === 'auxiliary') return '结构密封胶及辅助垫块类';
          return '其它辅材类';
        };

        exportRows.push([
          item.name || "",
          item.drawingNo || "",
          item.materialModel || "",
          item.size || "",
          item.isCalculated ? String(item.calculatedUnitUsage?.toFixed(4)) : "0.0000",
          item.isCalculated ? String(item.calculatedQty?.toFixed(2)) : "0.00",
          item.unit || "",
          "",
          "",
          item.remark || "",
          getFriendlyCategory(item.category)
        ]);
      });

      const summaryWs = XLSX.utils.aoa_to_sheet(exportRows);
      
      // Auto width formatting for columns
      summaryWs['!cols'] = [
        { wch: 24 }, { wch: 15 }, { wch: 22 }, { wch: 18 }, { wch: 15 }, { wch: 12 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 25 }, { wch: 22 }
      ];

      // Remove sheet if already exists to overwrite, then append
      if (wb.SheetNames.includes("辅材配套台账清单")) {
        wb.Sheets["辅材配套台账清单"] = summaryWs;
      } else {
        XLSX.utils.book_append_sheet(wb, summaryWs, "辅材配套台账清单");
      }

      XLSX.writeFile(wb, `双向COM联动算料填回表_已写回_${sheetName}.xlsx`);
      alert(`🎉 原件写回并附加「辅材配套台账清单」工作表处理成功！已成功定位原表并精准写入 ${filledCount} 行，已自动开始下载。`);
    } catch (err: any) {
      alert(`Excel 写入失败: ${err.message}`);
    }
  };

  // Helper action to restore preset workbook, ledger and rules data for evaluation
  const handleLoadBuiltInDemoData = () => {
    const wb = buildDefaultWorkbook();
    setWorkbook(wb);
    const defaultSheet = wb.SheetNames.find(s => !s.includes("目录")) || wb.SheetNames[0] || "D01-标准1500单元细目表";
    setSheetName(defaultSheet);
    setRawRows(DEFAULT_RAW_ROWS);
    setExcelSheets(wb.SheetNames);
    setSelectedSheet(defaultSheet);
    updateLedgerWithHistory(DEFAULT_LEDGER, "载入内置演示台账");
    setMappings({
      categoryCol: null,
      nameCol: 2,       // Col C
      specCol: 3,       // Col D
      materialCol: 4,   // Col E
      lenCol: 5,        // Col F
      qtyCol: 7,        // Col H
      unitCol: 8,       // Col I
      remarkCol: 11,    // Col L
    });
    setRanges({
      profileRange: 'C8:L12',
      steelRange: 'C16:L19',
      panelRange: 'C22:L25',
      profileKeyword: '铝  型  材',
      steelKeyword: '钢   件',
      panelKeyword: '面   板',
    });
  };

  // Initial load on mount (keep state clean until user explicitly clicks "采用示例")
  useEffect(() => {
    // Intentionally empty: do not automatically load sample unit data and auxiliary material ledger on startup
  }, []);

  const handleGlobalRefreshExcel = () => {
    if (wsConnected && localWs && localWs.readyState === WebSocket.OPEN) {
      handlePullFromExcel();
    } else if (sheetName === "单元细目表_演示模板") {
      handleLoadBuiltInDemoData();
      alert("已重新载入并刷新内置演示 Excel 数据！");
    } else {
      handlePullFromExcel();
    }
  };

  const calculateResultsForSheet = (
    targetSheetName: string,
    sheetRawRows: string[][],
    ledgerRules: AuxiliaryLedgerItem[],
    rangeConfig: RangeConfiguration
  ) => {
    const sanitizedRows = sheetRawRows.map(r => r.map(c => sanitizeExcelValue(c)));

    const getExcelCellVal = (r: number, c: number) => {
      if (r >= sanitizedRows.length) return 0;
      if (c >= sanitizedRows[r].length) return 0;
      const val = sanitizedRows[r][c] || "";
      const num = parseFloat(String(val).replace(/[^\d.-]/g, ''));
      return isNaN(num) ? 0 : num;
    };

    const sheetMatch = targetSheetName.match(/^([A-Za-z0-9]+)/);
    const parsedPlateNo = sheetMatch ? sheetMatch[1] : targetSheetName;

    let parsedPlateModel = "";
    if (sanitizedRows && sanitizedRows[2] && sanitizedRows[2][2]) {
      const cellC3 = String(sanitizedRows[2][2]);
      const modelMatch = cellC3.match(/(?:型号|编号)\s*[:：]\s*([A-Za-z0-9_\-]+)/i);
      if (modelMatch) {
        parsedPlateModel = modelMatch[1];
      } else {
        parsedPlateModel = cellC3.replace(/^(板块型号|板块编号|型号|编号)\s*[:：]\s*/, "").trim();
      }
    }

    const pVars = {
      W: getExcelCellVal(3, 7) / 1000,
      H: getExcelCellVal(3, 8) / 1000,
      W1: getExcelCellVal(1, 12) / 1000,
      panelsCount: getExcelCellVal(2, 10),
      plateNo: parsedPlateNo,
      plateCode: parsedPlateModel || parsedPlateNo,
      plateModel: parsedPlateModel,
      sheetName: targetSheetName
    };

    const unitPanelsCount = pVars.panelsCount > 0 ? pVars.panelsCount : 1;
    const ws = XLSX.utils.aoa_to_sheet(sanitizedRows);

    const pk = rangeConfig.profileKeyword || "铝  型  材";
    const sk = rangeConfig.steelKeyword || "钢   件";
    const pak = rangeConfig.panelKeyword || "面   板";

    const pResolved = resolveRangeByBColumnKeyword(ws, sanitizedRows, pk);
    const sResolved = resolveRangeByBColumnKeyword(ws, sanitizedRows, sk);
    const paResolved = resolveRangeByBColumnKeyword(ws, sanitizedRows, pak);

    const sheetRanges = { ...rangeConfig };
    if (pResolved) sheetRanges.profileRange = `C${pResolved.startRow + 1}:L${pResolved.endRow + 1}`;
    if (sResolved) sheetRanges.steelRange = `C${sResolved.startRow + 1}:L${sResolved.endRow + 1}`;
    if (paResolved) sheetRanges.panelRange = `C${paResolved.startRow + 1}:L${paResolved.endRow + 1}`;

    const pRange = parseExcelRange(sheetRanges.profileRange);
    const sRange = parseExcelRange(sheetRanges.steelRange);
    const paRange = parseExcelRange(sheetRanges.panelRange);

    const extractedItems: ProcessedItem[] = [];

    for (let rIdx = 0; rIdx < sanitizedRows.length; rIdx++) {
      const row = sanitizedRows[rIdx];
      if (!row || row.filter(c => c && String(c).trim() !== "").length === 0) continue;

      const name = sanitizeExcelValue(row[2]);
      const drawingNo = sanitizeExcelValue(row[3]);
      const materialModel = sanitizeExcelValue(row[4]);
      const remark = sanitizeExcelValue(row[11]);

      let category: MaterialType | null = null;
      if (pRange && rIdx >= pRange.startRow && rIdx <= pRange.endRow) {
        category = 'profile';
      } else if (sRange && rIdx >= sRange.startRow && rIdx <= sRange.endRow) {
        category = 'steel';
      } else if (paRange && rIdx >= paRange.startRow && rIdx <= paRange.endRow) {
        category = 'panel';
      }

      if (!category) continue;

      if (!name || name === "名称" || name === "品名") continue;
      let cuttingSize = String(sanitizeExcelValue(row[5])).trim();
      if (cuttingSize.startsWith("-") || parseFloat(cuttingSize) < 0) {
        cuttingSize = "";
      }
      const unitUsage = sanitizeExcelValue(row[6]);

      const qtyStr = sanitizeExcelValue(row[7]);
      const rawQty = qtyStr.replace(/[^\d.]/g, "");
      const qty = parseFloat(rawQty) || 0;

      const unit = sanitizeExcelValue(row[8]);
      const l1 = sanitizeExcelValue(row[9]);
      const l2 = sanitizeExcelValue(row[10]);

      let len = 0;
      if (cuttingSize) {
        const cleanCutSize = cuttingSize.replace(/[^\d.]/g, "");
        const sizeInMm = parseFloat(cleanCutSize) || 0;
        len = sizeInMm > 0 ? sizeInMm / 1000 : 0;
      }

      extractedItems.push({
        rowIndex: rIdx,
        originalCategory: category === 'profile' ? '型材' : (category === 'steel' ? '钢件' : '面板'),
        category,
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

    const calculatedResults = ledgerRules.map(item => {
      const itemRules = item.rules || [];
      const activeRules = itemRules.filter(r => r.isActive);

      if (activeRules.length > 0) {
        let calculatedUnitUsage = 0;
        const totalLogs: string[] = [];

        const evaluatedRules = activeRules.map(rule => {
          const { value, logs } = evaluateFormula(rule.expression, extractedItems, pVars, rule.lengthAdjustment || 0, sanitizedRows);
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
          finalUnitUsage = Math.round(finalUnitUsage * 150) / 150;
          finalUnitUsage = Math.round(finalUnitUsage * 100) / 100;
          finalQty = Math.round(finalQty * 100) / 100;
        }

        const decPlaces = (item.category === 'fastener' || isCountUnit) ? 0 : 2;
        totalLogs.push(`【总计折算核算】: 总计 = 单樘用量 (${finalUnitUsage.toFixed(decPlaces)}) * 单元板块樘数 (${unitPanelsCount}) = ${finalQty.toFixed(decPlaces)}。`);

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

    return {
      calculatedResults,
      projectVars: pVars,
      items: extractedItems,
      rawRows: sanitizedRows,
      ws
    };
  };

  const buildCellUpdatesForSheet = (
    ws: XLSX.WorkSheet,
    parsedRawRows: string[][],
    calculatedResultsList: any[],
    targetSheetName: string
  ) => {
    const cellUpdates: { r: number, c: number, v: any, bg?: number }[] = [];
    const colStart = 2; // C col
    const colEnd = 11; // L col

    const extractUpdatesForRegion = (keyword: string, category: string) => {
      const range = resolveRangeByBColumnKeyword(ws, parsedRawRows, keyword);
      if (!range) return;

      const targetItems = calculatedResultsList.filter(item => 
        item.category === category && item.isCalculated && (item.calculatedUnitUsage || 0) > 0
      );

      let currentRowOffset = 0;

      for (let r = range.startRow; r <= range.endRow; r++) {
        if (currentRowOffset < targetItems.length) {
          const item = targetItems[currentRowOffset];

          cellUpdates.push({ r, c: colStart, v: item.name || "" });
          cellUpdates.push({ r, c: colStart + 1, v: item.drawingNo || "" });
          cellUpdates.push({ r, c: colStart + 2, v: item.materialModel || "" });
          cellUpdates.push({ r, c: colStart + 3, v: item.cuttingSize || "" });
          cellUpdates.push({ r, c: colStart + 4, v: item.calculatedUnitUsage, bg: 65535 }); 
          cellUpdates.push({ r, c: colStart + 5, v: `=G${r + 1}*K3`, bg: 65535 });
          cellUpdates.push({ r, c: colStart + 6, v: item.unit || "" });
          cellUpdates.push({ r, c: colStart + 7, v: item.l1 || "" });
          cellUpdates.push({ r, c: colStart + 8, v: item.l2 || "" });
          cellUpdates.push({ r, c: colStart + 9, v: item.remark || "" });

          currentRowOffset++;
        } else {
          // 超出有效物料的行，清空旧数据
          for (let c = colStart; c <= colEnd; c++) {
            cellUpdates.push({ r, c, v: "" });
          }
        }
      }
    };

    extractUpdatesForRegion("胶条", "gasket");
    extractUpdatesForRegion("紧固件", "fastener");
    extractUpdatesForRegion("辅材", "auxiliary");

    return cellUpdates;
  };

  const readSheetPromise = (targetSheet: string) => {
    return new Promise<any>((resolve, reject) => {
      if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
        reject(new Error("本地 WebSocket 网桥连接已断开"));
        return;
      }

      let timeoutTimer: NodeJS.Timeout | null = null;

      wsResolveRef.current = (response) => {
        if (timeoutTimer) {
          clearTimeout(timeoutTimer);
          timeoutTimer = null;
        }
        resolve(response);
      };

      timeoutTimer = setTimeout(() => {
        if (wsResolveRef.current) {
          wsResolveRef.current = null;
          reject(new Error(`读取工作表 [${targetSheet}] 超时 (15秒)`));
        }
      }, 15000);

      localWs.send(JSON.stringify({
        action: "read_excel",
        sheetName: targetSheet
      }));
    });
  };

  const writeSheetPromise = (targetSheet: string, cellUpdates: any[]) => {
    return new Promise<void>((resolve, reject) => {
      if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
        reject(new Error("本地 WebSocket 网桥连接已断开"));
        return;
      }

      let timeoutTimer: NodeJS.Timeout | null = null;
      // 动态根据回写单元格数量调整超时时间（最少 60 秒，每 100 单元格加 15 秒）
      const timeoutMs = Math.max(60000, cellUpdates.length * 150);

      const resetTimer = () => {
        if (timeoutTimer) {
          clearTimeout(timeoutTimer);
        }
        timeoutTimer = setTimeout(() => {
          if (wsWriteResolveRef.current) {
            wsWriteResolveRef.current = null;
            reject(new Error(`回写工作表 [${targetSheet}] 超时 (${Math.round(timeoutMs / 1000)}秒内未完成)`));
          }
        }, timeoutMs);
      };

      wsWriteResolveRef.current = {
        resolve: () => {
          if (timeoutTimer) {
            clearTimeout(timeoutTimer);
            timeoutTimer = null;
          }
          resolve();
        },
        reject: (err) => {
          if (timeoutTimer) {
            clearTimeout(timeoutTimer);
            timeoutTimer = null;
          }
          reject(err);
        },
        resetTimeout: () => {
          // 实时收到进度广播时续期超时时间，防止大批量数据回填时被前端过早熔断
          resetTimer();
        }
      };

      resetTimer();

      localWs.send(JSON.stringify({
        action: "write_cells",
        excelFileName: targetSheet,
        updates: cellUpdates,
      }));
    });
  };

  const executeVbPromise = (vbCode: string) => {
    return new Promise<void>((resolve, reject) => {
      if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
        reject(new Error("本地 WebSocket 网桥连接已断开"));
        return;
      }

      let timeoutTimer: NodeJS.Timeout | null = null;

      wsWriteResolveRef.current = {
        resolve: () => {
          if (timeoutTimer) {
            clearTimeout(timeoutTimer);
            timeoutTimer = null;
          }
          resolve();
        },
        reject: (err) => {
          if (timeoutTimer) {
            clearTimeout(timeoutTimer);
            timeoutTimer = null;
          }
          reject(err);
        }
      };

      timeoutTimer = setTimeout(() => {
        if (wsWriteResolveRef.current) {
          wsWriteResolveRef.current = null;
          reject(new Error("执行 VB 脚本超时 (30秒)"));
        }
      }, 30000);

      localWs.send(JSON.stringify({
        action: "execute_vb",
        code: vbCode,
      }));
    });
  };

  const handleStopBatch = () => {
    if (!batchActive || stopBatchRef.current) return;
    stopBatchRef.current = true;
    setBatchStopping(true);
    setBatchProgress(prev => ({
      ...prev,
      step: 'stopping',
      logs: [
        ...prev.logs,
        {
          sheetName: prev.currentSheetName || "批量任务",
          status: 'stopped',
          message: "收到用户停止指令，正在安全中止后续工作表填回..."
        }
      ]
    }));
  };

  const startBatchProcess = async (selectedSheets: string[]) => {
    if (!wsConnected || !localWs || localWs.readyState !== WebSocket.OPEN) {
      alert("本地 COM 网桥未连通！请配对启动『本地 OLE COM 网桥连接』后再进行批量填回操作。");
      return;
    }

    stopBatchRef.current = false;
    setBatchStopping(false);
    setBatchActive(true);
    setBatchProgress({
      currentSheetName: "",
      step: 'idle',
      completedCount: 0,
      totalCount: selectedSheets.length,
      logs: []
    });

    const logsList: Array<{ sheetName: string; status: 'success' | 'failed' | 'stopped'; message: string }> = [];

    for (let i = 0; i < selectedSheets.length; i++) {
      if (stopBatchRef.current) {
        logsList.push({
          sheetName: selectedSheets[i],
          status: 'stopped',
          message: `用户手动停止：已跳过此工作表及后续所有工作表。`
        });
        break;
      }

      const sheet = selectedSheets[i];
      
      setBatchProgress(prev => ({
        ...prev,
        currentSheetName: sheet,
        step: 'reading',
        completedCount: i
      }));

      try {
        const readRes = await readSheetPromise(sheet);
        if (!readRes.success) {
          throw new Error(readRes.message || "读取料单失败");
        }

        if (stopBatchRef.current) {
          logsList.push({
            sheetName: sheet,
            status: 'stopped',
            message: `用户手动停止：已在计算回填前终止。`
          });
          break;
        }
        
        setBatchProgress(prev => ({
          ...prev,
          step: 'calculating'
        }));

        const { calculatedResults, ws, rawRows: parsedRawRows } = calculateResultsForSheet(
          sheet,
          readRes.rows,
          ledger,
          ranges
        );

        const cellUpdates = buildCellUpdatesForSheet(ws, parsedRawRows, calculatedResults, sheet);
        
        if (cellUpdates.length === 0) {
          throw new Error("无法在当前表格中检索到「胶条」「紧固件」「辅材」等区域用于填回。请确保读取了正确的表格区段。");
        }

        if (stopBatchRef.current) {
          logsList.push({
            sheetName: sheet,
            status: 'stopped',
            message: `用户手动停止：已在写入表格前终止。`
          });
          break;
        }

        setBatchProgress(prev => ({
          ...prev,
          step: 'writing'
        }));

        await writeSheetPromise(sheet, cellUpdates);

        logsList.push({
          sheetName: sheet,
          status: 'success',
          message: `成功计算并填回了 ${calculatedResults.filter(r => r.isCalculated && r.calculatedQty > 0).length} 项辅料五金配套用量。`
        });

      } catch (err: any) {
        console.error(`Batch processing error for ${sheet}:`, err);
        logsList.push({
          sheetName: sheet,
          status: 'failed',
          message: err.message || "未知错误"
        });
      }

      setBatchProgress(prev => ({
        ...prev,
        logs: [...logsList],
        completedCount: i + 1
      }));

      if (stopBatchRef.current) {
        break;
      }
    }

    const wasStopped = stopBatchRef.current;
    stopBatchRef.current = false;
    setBatchStopping(false);
    setBatchActive(false);

    setBatchProgress(prev => ({
      ...prev,
      step: 'idle',
      logs: [...logsList]
    }));

    if (wasStopped) {
      const successCount = logsList.filter(l => l.status === 'success').length;
      alert(`🛑 批量导入填回已停止！\n已成功处理 ${successCount} 个工作表，已取消后续未执行工作表。`);
      return;
    }

    const failedList = logsList.filter(l => l.status === 'failed');
    if (failedList.length === 0) {
      alert(`🎉 批量导入填回操作全部完成！共成功处理 ${logsList.length} 个工作表。`);
    } else if (failedList.length === logsList.length) {
      alert(`❌ 批量导入填回全部失败！原因:\n${failedList.map(f => `[${f.sheetName}]: ${f.message}`).join("\n")}`);
    } else {
      alert(`⚠️ 批量导入填回部分完成！\n成功: ${logsList.length - failedList.length} 个工作表\n失败: ${failedList.length} 个工作表。\n\n具体失败原因如下:\n${failedList.map(f => `[${f.sheetName}]: ${f.message}`).join("\n")}`);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-950 font-sans flex flex-col justify-between" id="app_frame">
      <div className="flex-1">
        {/* Balanced Architectural Header */}
        <header className="h-16 flex items-center justify-between px-6 border-b border-slate-200 bg-white sticky top-0 z-30 shadow-xs">
          <div className="flex items-center gap-4">
            <div className="w-10 h-10 bg-indigo-600 rounded flex items-center justify-center text-white font-bold shadow-md">
              <FileSpreadsheet className="w-6 h-6" />
            </div>
            <div>
              <h1 className="text-base sm:text-lg font-bold tracking-tight text-slate-800 flex items-center gap-2">
                幕墙工程辅材/五金件智能计算填回系统
                <span className="bg-indigo-50 border border-indigo-100 text-[9px] text-indigo-700 font-bold px-2 py-0.5 rounded flex items-center gap-1">
                  <Sparkles className="w-2.5 h-2.5" />
                  COM双向穿透版 v3.2
                </span>
              </h1>
            </div>
          </div>
          
          <div className="flex items-center gap-3">
            <button
              onClick={handleGlobalRefreshExcel}
              disabled={pullLoading}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.98] text-white text-xs font-bold rounded-lg shadow-sm transition-all cursor-pointer disabled:opacity-50 disabled:cursor-not-allowed select-none"
              id="refresh_excel_btn"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${pullLoading ? 'animate-spin' : ''}`} />
              <span>刷新excel数据</span>
            </button>

            {wsConnected ? (
              <button
                onClick={() => {
                  setUserDisconnected(true);
                  if (localWs) {
                    localWs.close();
                  }
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-rose-600 hover:bg-rose-700 active:scale-[0.98] text-white text-xs font-bold rounded-lg shadow-sm transition-all cursor-pointer select-none"
                id="disconnect_excel_btn"
                title="主动断开与本地 Python 网桥的连接，释放 WPS/Excel 进程占用"
              >
                <Unlink className="w-3.5 h-3.5" />
                <span>断开连接</span>
              </button>
            ) : (
              <button
                onClick={() => {
                  setUserDisconnected(false);
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 hover:bg-indigo-700 active:scale-[0.98] text-white text-xs font-bold rounded-lg shadow-sm transition-all cursor-pointer select-none"
                id="connect_excel_btn"
                title="重新建立与本地 Python 网桥的 WebSocket 联通通道"
              >
                <Link className="w-3.5 h-3.5" />
                <span>连接网桥</span>
              </button>
            )}

            {excelSheets.length > 0 && (
              <div className="relative inline-block">
                <select
                  value={selectedSheet}
                  onChange={(e) => {
                    const val = e.target.value;
                    setSelectedSheet(val);
                    handlePullFromExcel(val);
                  }}
                  disabled={pullLoading}
                  className="appearance-none bg-indigo-50 hover:bg-indigo-100 text-indigo-800 border border-indigo-200 rounded-lg pl-3 pr-8 py-1.5 text-xs font-bold transition-all cursor-pointer disabled:opacity-50 select-none focus:outline-none focus:ring-1 focus:ring-indigo-500 font-sans"
                  style={{
                    backgroundImage: `url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='%234f46e5' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'><polyline points='6 9 12 15 18 9'></polyline></svg>")`,
                    backgroundRepeat: 'no-repeat',
                    backgroundPosition: 'right 8px center',
                    backgroundSize: '14px'
                  }}
                  title="工作表名称列表"
                >
                  {excelSheets.map((sh, idx) => (
                    <option key={idx} value={sh}>
                      📄 {sh}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {workbook ? (
              <div className="flex items-center px-4 py-1.5 bg-green-50 border border-green-200 rounded text-xs font-bold text-green-700 shadow-sm animate-pulse">
                <span className="w-2 h-2 bg-green-500 rounded-full mr-2"></span>
                Excel 数据表已加载 ({sheetName})
              </div>
            ) : (
              <div className="flex items-center px-4 py-1.5 bg-slate-50 border border-slate-200 rounded text-xs font-medium text-slate-400">
                暂无活动数据表
              </div>
            )}
          </div>
        </header>

        {/* Top Navigation Menu Bar */}
        <div className="bg-white border-b border-slate-200 sticky top-16 z-20 shadow-xs">
          <div className="max-w-7xl mx-auto px-6">
            <nav className="flex flex-wrap gap-2 py-3" aria-label="Tabs">
              {[
                { id: "bridge", name: "本地 OLE COM 网桥连接", icon: Server },
                { id: "import", name: "料单导入与区间映射", icon: FileSpreadsheet },
                { id: "ledger", name: "配套五金与辅材台账", icon: Layers },
                { id: "calculation", name: "智能算料与一键填回", icon: Calculator },
                { id: "batch", name: "批量导入填回", icon: Sparkles },
                { id: "summary", name: "数据汇总", icon: BarChart3 },
              ].map((tab) => {
                const Icon = tab.icon;
                const isActive = activeTab === tab.id;
                return (
                  <button
                    key={tab.id}
                    onClick={() => setActiveTab(tab.id)}
                    className={`flex items-center gap-2 px-4 py-2.5 rounded-lg text-xs font-bold transition-all border cursor-pointer select-none ${
                      isActive
                        ? "bg-indigo-600 text-white border-indigo-700 shadow-sm"
                        : "bg-slate-50 text-slate-600 hover:text-slate-900 hover:bg-slate-100 border-slate-200"
                    }`}
                  >
                    <Icon className="w-4 h-4 shrink-0" />
                    <span>{tab.name}</span>
                  </button>
                );
              })}
            </nav>
          </div>
        </div>

        {/* Main Tab Views container */}
        <main className="max-w-7xl mx-auto px-6 mt-6 pb-16">
          
          {activeTab === "bridge" && (
            <div className="space-y-6">
              <ExcelComConnector
                connection={comConnection}
                localWsPort={localWsPort}
                setLocalWsPort={setLocalWsPort}
                wsConnected={wsConnected}
                comBridgeStatus={comBridgeStatus}
                comBridgeMessage={comBridgeMessage}
                onSendSimulation={() => {
                  handleTriggerComWriteBack(ledger);
                }}
              />
              
              {/* OLE COM bridge state reminder in bridge tab */}
              {wsConnected ? (
                <div className="bg-emerald-50 border border-emerald-200 px-5 py-4 rounded-xl flex items-center gap-3 text-left">
                  <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                  <div>
                    <h4 className="text-xs font-bold text-emerald-850">COM/OLE 网桥已安全连通</h4>
                    <p className="text-[11px] text-emerald-700/80 mt-0.5">无缝对接本地表格！请前往第 2 菜单中配置坐标引用区间并直接提取您的钢铝型材与面板数据。</p>
                  </div>
                </div>
              ) : (
                <div className="bg-slate-50 border border-slate-200 p-4 rounded-xl flex items-center gap-3 text-left">
                  <span className="w-2 h-2 bg-slate-300 rounded-full animate-ping shrink-0" />
                  <div className="space-y-0.5">
                    <h3 className="text-xs font-bold text-slate-700">COM 心跳网桥后台等待接入中...</h3>
                    <p className="text-[11px] text-slate-400">请先下载网桥程序并于本地电脑执行 <code>python com_bridge.py</code>，启动监听后将自动配对并加载。您也可以在此期间进行公式套算调试！</p>
                  </div>
                </div>
              )}
            </div>
          )}

          {activeTab === "import" && (
            <div className="space-y-6">
              {/* OLE COM direct pull client widget */}
              {wsConnected ? (
                <div className="bg-emerald-50 border border-emerald-250 p-5 rounded-xl shadow-xs flex flex-col md:flex-row justify-between items-start md:items-center gap-4 animate-fade-in text-left shadow-emerald-50">
                  <div className="space-y-1 text-left flex-1">
                    <h3 className="text-sm font-bold text-emerald-800 flex items-center gap-2">
                      <span className="w-2.5 h-2.5 bg-emerald-500 rounded-full animate-pulse" />
                      本地 OLE/COM 双向穿透网桥 [已成功连接]
                    </h3>
                    <p className="text-xs text-emerald-700/80">
                      通过下方填写分组坐标区间 (如格式 <code>C8:L57</code>)，即可免除拖拽，直接按选区截取、分类加载下方表格！
                    </p>
                    {importStatusMessage && (
                      <p className="text-xs text-indigo-700 font-bold animate-pulse mt-1">
                        🚀 {importStatusMessage}
                      </p>
                    )}
                  </div>
                  
                  <div className="flex flex-col sm:flex-row items-center gap-2.5 shrink-0 w-full md:w-auto">
                    {excelSheets.length > 0 && (
                      <div className="flex items-center gap-1.5 bg-white border border-emerald-250 rounded-lg px-2.5 py-1.5 shadow-2xs w-full sm:w-auto">
                        <span className="text-[10px] font-bold text-slate-400 shrink-0 uppercase">工作表:</span>
                        <select
                          value={selectedSheet}
                          onChange={(e) => {
                            const val = e.target.value;
                            setSelectedSheet(val);
                            handlePullFromExcel(val);
                          }}
                          disabled={pullLoading}
                          className="bg-transparent text-xs font-bold text-indigo-700 border-none p-0 focus:ring-0 focus:outline-none max-w-[130px] truncate cursor-pointer font-sans"
                        >
                          {excelSheets.map((sh, idx) => (
                            <option key={idx} value={sh}>{sh}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <button
                      disabled={pullLoading}
                      onClick={() => handlePullFromExcel()}
                      className="bg-emerald-600 hover:bg-emerald-700 hover:scale-[1.02] text-white text-xs font-bold px-4 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 cursor-pointer shadow-sm shadow-emerald-100 disabled:opacity-50 shrink-0 select-none w-full sm:w-auto"
                    >
                      {pullLoading ? (
                        <RefreshCw className="w-4 h-4 animate-spin" />
                      ) : (
                        <Server className="w-4 h-4" />
                      )}
                      {pullLoading ? "正在拉取料单..." : "一键导入本地 Excel 当前活动料单"}
                    </button>
                  </div>
                </div>
              ) : (
                <div className="bg-slate-50 border border-slate-200 p-5 rounded-xl flex flex-col md:flex-row justify-between items-start md:items-center gap-3 text-left">
                  <div className="space-y-0.5">
                    <h3 className="text-xs font-bold text-slate-705 flex items-center gap-1.5 leading-snug text-slate-700">
                      💡 提示：当前正处于「模拟示范数据」运行环境中
                    </h3>
                    <p className="text-[11px] text-slate-400 leading-relaxed">
                      当您在第一个菜单中启动 <code>python com_bridge.py</code> 连接本地 OLE 时，此面板将自动升级为【Excel 高亮即时直读窗口】，实时通过坐标区间提取渲染！
                    </p>
                  </div>
                </div>
              )}

              {workbook ? (
                <div className="space-y-6">
                  <ExcelRangeSelector
                    ranges={ranges}
                    onUpdateRanges={(newRanges) => setRanges(newRanges)}
                    itemsCount={items.length}
                    items={items}
                    workbook={workbook}
                    sheetName={sheetName}
                    rawRows={rawRows}
                  />

                  <MaterialGroupList
                    items={items}
                    mappings={mappings}
                    availableColumns={availableColumns}
                    onUpdateMappings={handleUpdateMappings}
                    onUpdateItemCategory={handleUpdateItemCategory}
                  />
                </div>
              ) : (
                <div className="bg-white p-8 rounded-xl border border-slate-200 shadow-xs text-xs text-center text-slate-500 py-16 flex flex-col items-center justify-center gap-4 font-sans">
                  <div className="max-w-md space-y-1.5">
                    <p className="font-bold text-slate-700 text-sm">暂无活动料单数据 (No active workbook loaded)</p>
                    <p className="text-slate-400">应用已被默认清空。您可以开启 COM 穿透连接后直接拉取本地运行的 Excel 画布数据，或一键加载系统预设的工程细目与辅材台账对照示例进行即时测试算量！</p>
                  </div>
                  <button 
                    onClick={handleLoadBuiltInDemoData}
                    className="bg-indigo-600 hover:bg-indigo-700 hover:scale-[1.01] active:scale-[0.99] text-white text-xs font-bold px-4 py-2.5 rounded-lg transition-all flex items-center justify-center gap-2 cursor-pointer shadow-sm shadow-indigo-100"
                  >
                    <BookOpen className="w-4 h-4" />
                    采用示例 (加载演示单元数据与辅材台账)
                  </button>
                </div>
              )}
            </div>
          )}

          {activeTab === "ledger" && (
            <AuxMaterialLedger
              ledger={ledger}
              items={items}
              projectVars={projectVars}
              rawRows={rawRows}
              onAddLedgerItem={handleAddLedgerItem}
              onBulkAddLedgerItems={handleBulkAddLedgerItems}
              onRemoveLedgerItem={handleRemoveLedgerItem}
              onClearLedger={handleClearLedger}
              onUpdateLedgerItem={handleUpdateLedgerItem}
              onOverwriteLedgerItems={handleOverwriteLedgerItems}
              ledgerHistory={ledgerHistory}
              onRestoreHistory={handleRestoreHistory}
            />
          )}

          {activeTab === "calculation" && (
            <AuxCalculationResult
              ledger={ledger}
              items={items}
              onExportExcel={handleExportFinishedExcel}
              syncStatus={syncStatus}
              syncProgress={syncProgress}
              onTriggerSync={handleTriggerComWriteBack}
              sheetName={sheetName}
              projectVars={projectVars}
              excelSheets={excelSheets}
              onSelectSheet={(sh) => {
                setSelectedSheet(sh);
                handlePullFromExcel(sh);
              }}
              rawRows={rawRows}
            />
          )}

          {activeTab === "batch" && (
            <BatchImportFillback
              excelSheets={excelSheets}
              wsConnected={wsConnected}
              batchActive={batchActive}
              batchStopping={batchStopping}
              batchProgress={batchProgress}
              onStartBatch={startBatchProcess}
              onStopBatch={handleStopBatch}
              onRefreshSheets={() => handlePullFromExcel("")}
              pullLoading={pullLoading}
            />
          )}

          {activeTab === "summary" && (
            <DataSummaryModule
              workbook={workbook}
              excelSheets={excelSheets}
              wsConnected={wsConnected}
              pullLoading={pullLoading}
              ranges={ranges}
              ledger={ledger}
              readSheetPromise={readSheetPromise}
            />
          )}

          {/* Unified Collapsible Documentation Panel at the bottom of the page */}
          <div className="mt-8 border border-slate-200 rounded-xl bg-white shadow-2xs overflow-hidden max-w-7xl mx-auto text-left">
            <div 
              className="flex items-center justify-between px-5 py-4 bg-slate-50 border-b border-slate-200 cursor-pointer select-none hover:bg-slate-100/80 transition-colors"
              onClick={() => setIsDocOpen(!isDocOpen)}
            >
              <div className="flex items-center gap-2 text-indigo-900 font-bold text-sm">
                <BookOpen className="w-4 h-4 text-indigo-600" />
                <span>📖 算料参数单位说明与公式逻辑运算方法官方手册</span>
                <span className="text-[10px] bg-indigo-50 border border-indigo-150 text-indigo-700 px-2.5 py-0.5 rounded-full font-bold">
                  点击折叠/展开查阅
                </span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[11px] text-slate-400 font-medium">
                  {isDocOpen ? "点击折叠隐藏" : "点击展开查看完整算料逻辑"}
                </span>
                {isDocOpen ? (
                  <ChevronUp className="w-4 h-4 text-slate-500" />
                ) : (
                  <ChevronDown className="w-4 h-4 text-slate-500" />
                )}
              </div>
            </div>

            {isDocOpen && (
              <div className="p-6 space-y-6 text-xs text-slate-700 leading-relaxed bg-white animate-fade-in">
                {/* Section 1 */}
                <div className="space-y-2">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-150 pb-2">
                    <span className="inline-block w-1.5 h-3.5 bg-indigo-500 rounded-sm"></span>
                    一、算料公式物理参量与单位自动折算说明
                  </h4>
                  <div className="space-y-2 text-slate-600 pl-3">
                    <p>
                      • <b>参数自动换算为米 (m)</b>：公式内匹配到的参数（如 <code>ROW_VAL</code> 提取的型材与五金角码类下料尺寸、L1、L2 等）、所有<b>直接引用的 Excel 单元格</b>（如 <code>H4</code>、<code>I4</code> 等）、以及内置的 <b>W、H、W1</b> 等工程参量，均已被系统<b>自动从毫米 (mm) 等比折算为米 (m)</b> 参与运算。用户编写公式时可直接使用，无需手工除以 1000。
                    </p>
                    <p>
                      • <b>加长按钮联动生效</b>：若对规则设置了“加长尺寸”，则在公式计算时，该加长量（mm）<b>同样会自动应用到所有直接引用的单元格及匹配尺寸参数中</b>，先进行加长叠加，再整体折算为米进行计算。
                    </p>
                    <p>
                      • <b>内置工程参量单元格对应表</b>（用于提醒核对）：
                    </p>
                    <div className="mt-1.5 pl-3 border-l-2 border-indigo-300 font-mono text-[11px] text-slate-850 space-y-1 bg-slate-50 py-2 rounded max-w-2xl">
                      <div>- 内置 <b>W</b>：对应主表 <b>H4</b> 单元格（幕墙分格宽度，已自动换算成米）</div>
                      <div>- 内置 <b>H</b>：对应主表 <b>I4</b> 单元格（幕墙分格高度，已自动换算成米）</div>
                      <div>- 内置 <b>W1</b> (或 H1)：对应主表 <b>M2</b> 单元格（次要宽度参量，已自动换算成米）</div>
                      <div>- 内置 <b>panelsCount</b>：对应主表 <b>K3</b> 单元格（分格樘数）</div>
                    </div>
                    <p className="text-slate-400 text-[10.5px] italic">
                      * 请注意核对 Excel 原表中上述对应单元格数据的准确性，以确保算料精准。
                    </p>
                  </div>
                </div>

                {/* Section 2 - Core Built-in Formulas */}
                <div className="space-y-2">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-150 pb-2">
                    <span className="inline-block w-1.5 h-3.5 bg-indigo-500 rounded-sm"></span>
                    二、系统内置核心算料函数参考手册
                  </h4>
                  <div className="space-y-4 text-slate-600 pl-3">
                    <p>
                      系统提供了 4 个预置公式函数，用于从 Excel 原始明细表、型材分类表或五金参数表中动态查询并汇总下料数据：
                    </p>

                    <div className="overflow-hidden border border-slate-150 rounded-lg">
                      <table className="w-full text-left border-collapse bg-slate-50/50">
                        <thead>
                          <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-150">
                            <th className="p-2.5 text-[11px] w-1/4">函数名称及语法</th>
                            <th className="p-2.5 text-[11px] w-2/5">功能描述及运算规则</th>
                            <th className="p-2.5 text-[11px]">经典应用场景举例</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-150 text-[11px]">
                          <tr>
                            <td className="p-2.5 font-mono text-indigo-700 font-bold bg-white">
                              SUM_LEN('检索词')
                            </td>
                            <td className="p-2.5 text-slate-600 bg-white">
                              自动检索匹配明细行，累加所有匹配行的<b>“下料尺寸” × “数量/件数”</b>。
                              <br/>
                              <span className="text-amber-600 font-semibold">* 注意：提取结果会自动从毫米(mm)折算为米(m)输出。</span>
                            </td>
                            <td className="p-2.5 text-slate-500 bg-white">
                              <code className="text-slate-800 font-mono">SUM_LEN('立柱')</code>
                              <br/>
                              汇总所有匹配到“立柱”的材料的总切料长度（米）。
                            </td>
                          </tr>
                          <tr>
                            <td className="p-2.5 font-mono text-indigo-700 font-bold bg-white">
                              SUM_QTY('检索词')
                            </td>
                            <td className="p-2.5 text-slate-600 bg-white">
                              自动检索匹配明细行，累加所有匹配行的<b>“数量/件数”</b>（即不乘尺寸，纯求和数量）。
                            </td>
                            <td className="p-2.5 text-slate-500 bg-white">
                              <code className="text-slate-800 font-mono">SUM_QTY('角码')</code>
                              <br/>
                              汇总所有匹配到“角码”的配件的总支数/件数。
                            </td>
                          </tr>
                          <tr>
                            <td className="p-2.5 font-mono text-indigo-700 font-bold bg-white">
                              ROW_VAL('类别', '搜索列', '检索词', '目标列', '是否精确')
                            </td>
                            <td className="p-2.5 text-slate-600 bg-white">
                              用于去<b>参数账册表</b>中精准搜索指定行，并提取该行指定属性列的单值数值（如型材下料尺寸、壁厚、米重、角码间距等）。
                            </td>
                            <td className="p-2.5 text-slate-500 bg-white">
                              <code className="text-slate-800 font-mono">ROW_VAL('型材', '型材规格', 'H2-JT19', 'cuttingSize')</code>
                              <br/>
                              从“型材”分类账册中，查找型材规格为 H2-JT19 的那一行的下料尺寸（换算成米）。
                            </td>
                          </tr>
                          <tr>
                            <td className="p-2.5 font-mono text-indigo-700 font-bold bg-white">
                              MATCH('参数名', '模式')
                            </td>
                            <td className="p-2.5 text-slate-600 bg-white">
                              <b>条件匹配函数</b>。在参数满足指定条件或编号时返回 <code className="text-emerald-600 font-bold">1</code>，否则返回 <code className="text-slate-400">0</code>。
                              可以作为乘数因子实现逻辑分支条件触发。
                            </td>
                            <td className="p-2.5 text-slate-500 bg-white">
                              <code className="text-slate-800 font-mono">MATCH('W', '&gt;=1.5') * 2</code>
                              <br/>
                              若分格宽度大于等于1.5米，此项计算用量，否则用量直接归零。
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>

                {/* Section 3 */}
                <div className="space-y-2">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-150 pb-2">
                    <span className="inline-block w-1.5 h-3.5 bg-indigo-500 rounded-sm"></span>
                    三、公式与检索词逻辑表达式编写指南 (AND / OR)
                  </h4>
                  <div className="space-y-2 text-slate-600 pl-3">
                    <p>
                      算料取数函数（如 <code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">SUM_LEN</code>、<code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">SUM_QTY</code> 以及动态参数 <code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">ROW_VAL</code> 的搜索词）原生支持<b>组合逻辑运算和模糊通配符</b>：
                    </p>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
                      <div className="bg-emerald-50/40 border border-emerald-100 p-3.5 rounded-lg">
                        <span className="font-bold text-emerald-800 flex items-center gap-1 text-[13px] mb-1">🤝 逻辑且 (AND)</span>
                        <p className="text-slate-600 mb-2">同时满足多个关键词的行才会被匹配并带入计算。</p>
                        <div className="font-mono text-slate-500 text-[11px]">可用符号：<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">&amp;&amp;</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">&amp;</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">且</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">and</code> (不区分大小写)</div>
                        <div className="mt-2.5 font-semibold text-slate-800 bg-white p-1.5 rounded border border-emerald-200">例：<code className="text-indigo-700 font-mono">SUM_LEN('H2-JT19 &amp; 立柱')</code></div>
                        <div className="text-slate-500 text-[10.5px] mt-1">解释：匹配名称或规格中<b>同时包含</b> "H2-JT19" 且包含 "立柱" 的行。</div>
                      </div>

                      <div className="bg-amber-50/40 border border-amber-100 p-3.5 rounded-lg">
                        <span className="font-bold text-amber-800 flex items-center gap-1 text-[13px] mb-1">🔀 逻辑或 (OR)</span>
                        <p className="text-slate-600 mb-2">满足任意一个关键词的行都会被匹配并带入计算。</p>
                        <div className="font-mono text-slate-500 text-[11px]">可用符号：<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">||</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">|</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">或</code>、<code className="font-bold text-slate-700 bg-white px-1 py-0.5 rounded border">or</code> (不区分大小写)</div>
                        <div className="mt-2.5 font-semibold text-slate-800 bg-white p-1.5 rounded border border-amber-200">例：<code className="text-indigo-700 font-mono">SUM_LEN('H2-JT19 | H2-JT20')</code></div>
                        <div className="text-slate-500 text-[10.5px] mt-1">解释：匹配名称或规格中包含 "H2-JT19" <b>或者</b>包含 "H2-JT20" 的行。</div>
                      </div>
                    </div>

                    <div className="bg-blue-50/30 border border-blue-100 p-3.5 rounded-lg mt-3">
                      <span className="font-bold text-blue-800 flex items-center gap-1 mb-1">✨ 模糊通配符与高级嵌套</span>
                      <ul className="list-disc pl-5 space-y-1 text-slate-600">
                        <li><b>通配符：</b><code className="font-mono bg-white px-1 rounded border text-slate-700">*</code> 代表任意多个字符，<code className="font-mono bg-white px-1 rounded border text-slate-700">?</code> 代表单个字符。</li>
                        <li><b>示例：</b><code className="font-mono text-indigo-700">SUM_LEN('H2-JT*')</code> 将完美匹配 H2-JT19、H2-JT20、H2-JT-A 等所有前缀。</li>
                        <li><b>多级嵌套复合逻辑：</b>可以使用括号分级，例如：<code className="font-mono text-indigo-700">SUM_LEN('(H2-JT19 &amp; 立柱) | (H2-JT20 &amp; 横梁)')</code></li>
                      </ul>
                    </div>
                  </div>
                </div>

                {/* Section 4 */}
                <div className="space-y-2">
                  <h4 className="text-sm font-bold text-slate-800 flex items-center gap-1.5 border-b border-slate-150 pb-2">
                    <span className="inline-block w-1.5 h-3.5 bg-indigo-500 rounded-sm"></span>
                    四、板块工程参量及板块编号匹配逻辑说明 (MATCH)
                  </h4>
                  <div className="space-y-2 text-slate-600 pl-3">
                    <p>
                      为了支持针对特定板块大小（如大板块、异形板块、超高超宽板块）或特定板块编号进行精细化的辅材算量条件触发，系统内置了 <code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">MATCH</code> (或 <code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">MATCH_PARAM</code> / <code className="font-mono bg-slate-100 px-1 py-0.5 rounded border text-indigo-700">MATCH_PLATE</code>) 函数。
                    </p>
                    <p className="font-semibold text-slate-800">
                      🎯 核心规则：MATCH 函数在匹配成功时输出 1，匹配失败时输出 0。用户可以将其用作乘数因子，实现条件化智能算料：
                    </p>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-3">
                      <div className="bg-indigo-50/40 border border-indigo-100 p-3.5 rounded-lg">
                        <span className="font-bold text-indigo-800 flex items-center gap-1 text-[13px] mb-1">📏 工程参量阈值匹配</span>
                        <p className="text-slate-600 mb-2">用于针对宽、高或副宽数值的数学比较大小，触发对应的用量逻辑。</p>
                        <div className="font-mono text-[10.5px] text-slate-600 space-y-1.5">
                          <div>• <b>板宽 H4</b>：使用 <code>MATCH('W', '&gt;=1.5')</code></div>
                          <div>• <b>板高 I4</b>：使用 <code>MATCH('H', '&gt;=2.0')</code></div>
                          <div>• <b>副宽 M2</b>：使用 <code>MATCH('W1', '==0.5')</code></div>
                          <div className="text-slate-400 font-sans italic text-[10px] mt-1">* 提示：匹配阈值单位均为<b>米 (m)</b>。比较操作符支持 <code>&gt;=</code>, <code>&lt;=</code>, <code>&gt;</code>, <code>&lt;</code>, <code>==</code>, <code>!=</code></div>
                        </div>
                        <div className="mt-2.5 font-semibold text-slate-800 bg-white p-1.5 rounded border border-indigo-200">
                          例：<code className="text-indigo-700 font-mono">MATCH('W', '&gt;=1.5') * 2</code>
                        </div>
                        <div className="text-slate-500 text-[10.5px] mt-1">解释：当分格板宽大于等于 1.5 米时，计算结果为 1 * 2 = 2；否则计算结果为 0 * 2 = 0。</div>
                      </div>

                      <div className="bg-sky-50/40 border border-sky-100 p-3.5 rounded-lg">
                        <span className="font-bold text-sky-800 flex items-center gap-1 text-[13px] mb-1">🏷️ 板块编号/型号字符串匹配</span>
                        <p className="text-slate-600 mb-2">用于针对当前的板块编号、板块型号或工作表名称进行模糊模式匹配。</p>
                        <div className="font-mono text-[10.5px] text-slate-600 space-y-1.5">
                          <div>• <b>编号匹配</b>：使用 <code>MATCH('板块编号', 'D01')</code></div>
                          <div>• <b>型号匹配</b>：使用 <code>MATCH('板块型号', 'UNIT*')</code> (支持通配符)</div>
                          <div>• <b>工作表名</b>：使用 <code>MATCH('sheetName', 'D*标准*')</code> (多重逻辑)</div>
                          <div className="text-slate-400 font-sans italic text-[10px] mt-1">* 提示：板块编号从 Excel 工作表名前缀提取，板块型号从 C3 单元格提取。</div>
                        </div>
                        <div className="mt-2.5 font-semibold text-slate-800 bg-white p-1.5 rounded border border-sky-200">
                          例：<code className="text-indigo-700 font-mono">MATCH('板块编号', 'D01 | D02') * SUM_QTY('螺栓')</code>
                        </div>
                        <div className="text-slate-500 text-[10.5px] mt-1">解释：仅当板块编号为 D01 或 D02 时，累加螺栓用量才生效，其它板块直接不计算螺栓用量。</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </main>
      </div>

      {/* High Contrast Precise Metric footer */}
      <footer className="h-10 bg-slate-100 border-t border-slate-200 text-slate-500 flex items-center justify-between px-6 text-xs font-sans">
        <div>
          <span>幕墙工程加工型材辅材智能配套计算与 COM 一键写回系统</span>
        </div>
        <div className="font-medium select-none text-right">
          © 幕墙设计工程计算套件 COM 穿透版
        </div>
      </footer>
    </div>
  );
}
