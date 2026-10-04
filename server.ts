import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

// Lazy initialize Gemini client
let aiClient: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI {
  if (!aiClient) {
    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      throw new Error("GEMINI_API_KEY environment variable is not set. Please add it via Settings > Secrets.");
    }
    aiClient = new GoogleGenAI({
      apiKey: apiKey,
      httpOptions: {
        headers: {
          "User-Agent": "aistudio-build",
        },
      },
    });
  }
  return aiClient;
}

app.use(express.json({ limit: "50mb" }));

// API: Analyze sheet column mappings, group rows, and recommend formula rules using Gemini
app.post("/api/analyze-sheet", async (req, res) => {
  try {
    const { rows, filename } = req.body;
    if (!rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: "Invalid rows data format" });
    }

    // Prepare a condensed sample of the spreadsheet to fit in token limit and maintain speed
    // Take the first 150 rows as sample context which is perfect
    const sampleRows = rows.slice(0, 150).map((row, idx) => ({
      index: idx,
      cells: row.map((cell: any) => (cell !== null && cell !== undefined) ? String(cell).trim() : ""),
    }));

    const client = getGeminiClient();

    const prompt = `You are an expert curtain wall (幕墙) engineering software assistant. 
Our user uploaded a curtain wall unit processing detailing list (材料加工细目/排料表) Excel sheet named "${filename || 'document.xlsx'}".
We need to analyze this file's structure. Here is a sample of the first 150 rows. 
Some rows represent columns header, some represent Profile (型材), Panel (面板/玻璃), Steel (钢件), Gaskets (胶条), Fasteners (紧固件/螺栓/螺钉), or Auxiliary Materials (辅材 as single-side tape, sealant, fire-insulation rockwool, spacer blocks).

Please assist us in:
1. Identifying which row index represents the Column Header row.
2. Mapping columns to standard field properties:
   - "category": Row group category (e.g. 型材, 面板, 钢件, 胶条, 紧固件, 辅材) - if merged visually, it might appear in a dedicated column.
   - "name": Item/Material name (e.g., 胶条, 钢垫板, 玻璃, 自攻螺钉, 密封胶).
   - "spec": Structural dimensions/drawing code (e.g., H2-JT19, J9A-JT01, 100*30*5, ST4.8*13).
   - "material": Material (e.g., 硬质橡胶, 尼龙, Q235B, 钢化夹胶玻璃).
   - "qty": Total numbers/quantity (e.g., 24, 28, 4.00).
   - "len": Individual or total length (e.g., 2.62, 13.72, 149) or empty.
   - "unit": Units (e.g., 米, 个, m2, ml, kg, 套, 只).
   - "remark": Usage or remarks (e.g., 护边胶条, 上极梁胶条, 镀锌板固定).
3. Analyzing and categorizing every row in the sample into: "profile", "panel", "steel", "gasket", "fastener", "auxiliary", or "other/header".
4. Recommending calculation formulas for each "gasket", "fastener", or "auxiliary" rows.
   A formula should refer to:
   - SUM_LEN(profile_spec_prefix) - e.g. SUM_LEN("H2-JT")
   - SUM_QTY(steel_spec) - e.g. SUM_QTY("100*30*5")
   - SUM_QTY(panel_name_or_spec) - e.g. SUM_QTY("玻璃")
   - SUM_PERIMETER(panel_spec) - perimeter of panel: 2 * (W + H)
   - SUM_AREA(panel_spec) - area of panel: W * H
   - Or a constant multipliers (coefficient context, e.g. length of profile * 1.05 for wastage).

Return the analyzed results strictly in JSON with this schema:
{
  "headerRowIndex": number (the 0-based row index containing the column headers),
  "columnMappings": {
    "categoryCol": number | null (0-based column index),
    "nameCol": number | null,
    "specCol": number | null,
    "materialCol": number | null,
    "qtyCol": number | null,
    "lenCol": number | null,
    "unitCol": number | null,
    "remarkCol": number | null
  },
  "rowClassifications": [
    {
      "rowIndex": number,
      "type": "profile" | "panel" | "steel" | "gasket" | "fastener" | "auxiliary" | "other",
      "name": "extracted name",
      "spec": "extracted spec drawing code",
      "qty": number,
      "len": number,
      "unit": "unit value",
      "material": "material value",
      "recommendedFormula": "recommended formula string (or empty if main material)"
    }
  ]
}

Please make your assessment highly professional, analyzing curtain wall profiles (like H2-JT) and fasteners specs.
Here is the sheet sample rows JSON:
${JSON.stringify(sampleRows)}`;

    const response = await client.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            headerRowIndex: {
              type: Type.INTEGER,
              description: "The row index where columns headers (名称, 规格, 数量, 备注, etc.) are located",
            },
            columnMappings: {
              type: Type.OBJECT,
              properties: {
                categoryCol: { type: Type.INTEGER, description: "Column index for category/section column if any" },
                nameCol: { type: Type.INTEGER, description: "Column index for Name/名称" },
                specCol: { type: Type.INTEGER, description: "Column index for Specification/Model/规格/型号" },
                materialCol: { type: Type.INTEGER, description: "Column index for Material/材质" },
                qtyCol: { type: Type.INTEGER, description: "Column index for Quantity/数量" },
                lenCol: { type: Type.INTEGER, description: "Column index for Length/长度 if any" },
                unitCol: { type: Type.INTEGER, description: "Column index for Unit/单位" },
                remarkCol: { type: Type.INTEGER, description: "Column index for Remarks/备注" },
              },
              required: ["nameCol", "specCol", "qtyCol", "unitCol"],
            },
            rowClassifications: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: {
                  rowIndex: { type: Type.INTEGER },
                  type: { type: Type.STRING, description: "One of profile, panel, steel, gasket, fastener, auxiliary, other" },
                  name: { type: Type.STRING },
                  spec: { type: Type.STRING },
                  qty: { type: Type.NUMBER },
                  len: { type: Type.NUMBER },
                  unit: { type: Type.STRING },
                  material: { type: Type.STRING },
                  recommendedFormula: { type: Type.STRING, description: "Suggested engineering formula using SUM_LEN, SUM_QTY, etc. based on matching materials" },
                },
                required: ["rowIndex", "type", "name"],
              },
            },
          },
          required: ["headerRowIndex", "columnMappings", "rowClassifications"],
        },
      },
    });

    const parsedData = JSON.parse(response.text || "{}");
    res.json(parsedData);
  } catch (error: any) {
    console.error("AI Analysis error:", error);
    res.status(500).json({ error: error.message || "Failed to analyze spreadsheet" });
  }
});

// ==========================================
// COM SYNCING APIS FOR LOCAL PYTHON BRIDGE
// ==========================================
interface SyncSession {
  code: string;
  items: any[];
  excelFileName: string;
  qtyCol: number;
  status: 'idle' | 'pending' | 'writing' | 'success' | 'error';
  message: string;
  updatedAt: number;
}

const syncSessions = new Map<string, SyncSession>();

// Auto-cleanup stale sessions (older than 2 hours)
setInterval(() => {
  const now = Date.now();
  for (const [code, session] of syncSessions.entries()) {
    if (now - session.updatedAt > 2 * 60 * 60 * 1000) {
      syncSessions.delete(code);
    }
  }
}, 10 * 60 * 1000);

// Initialize a sync session and return a connection code
app.post("/api/sync/init", (req, res) => {
  const code = "COM-" + Math.floor(1000 + Math.random() * 9000);
  const session: SyncSession = {
    code,
    items: [],
    excelFileName: "",
    qtyCol: 5,
    status: 'idle',
    message: "建立通道成功。请复制下方的本地程序配对口令或直接运行 Python 脚本进行穿透连接。",
    updatedAt: Date.now(),
  };
  syncSessions.set(code, session);
  res.json({ code });
});

// Push calculations from Web App page
app.post("/api/sync/push", (req, res) => {
  const { code, items, excelFileName, qtyCol } = req.body;
  if (!code) {
    return res.status(400).json({ error: "连接口令不能为空" });
  }

  const session = syncSessions.get(code);
  if (!session) {
    return res.status(404).json({ error: "口令已过期或不存在，请刷新重试" });
  }

  session.items = items || [];
  session.excelFileName = excelFileName || "";
  session.qtyCol = qtyCol !== undefined ? qtyCol : 5;
  session.status = 'pending';
  session.message = "云端算料计算完成！数据负载就绪，等待本地 Python 客户端拉取写入...";
  session.updatedAt = Date.now();

  res.json({ status: "success", code });
});

// Pull command payload (Called by Python program)
app.get("/api/sync/pull/:code", (req, res) => {
  const { code } = req.params;
  const session = syncSessions.get(code);

  if (!session) {
    return res.status(404).json({ error: "口令口令无效，未在服务器索引到您的同步通道" });
  }

  // Update session info to show PyBridge active connection
  if (session.status === 'idle') {
    session.message = "本地 Python 客户端已成功连接！通道开启常置监听，等待您一键发送写回指令...";
    session.updatedAt = Date.now();
  }

  res.json({
    code: session.code,
    items: session.items,
    excelFileName: session.excelFileName,
    qtyCol: session.qtyCol,
    status: session.status,
    message: session.message,
  });
});

// Notify write status (Called by Python program)
app.post("/api/sync/status/:code", (req, res) => {
  const { code } = req.params;
  const { status, message } = req.body;

  const session = syncSessions.get(code);
  if (!session) {
    return res.status(404).json({ error: "口令失效" });
  }

  session.status = status;
  session.message = message || "";
  session.updatedAt = Date.now();

  res.json({ status: "success" });
});

// Read write status in real time (Called by React Web UI polling)
app.get("/api/sync/status/:code", (req, res) => {
  const { code } = req.params;
  const session = syncSessions.get(code);

  if (!session) {
    return res.status(404).json({ error: "COM 对话已断开或在云端被清理" });
  }

  res.json({
    code: session.code,
    status: session.status,
    message: session.message,
    updatedAt: session.updatedAt,
  });
});

// Helper for Gemini call with rotation & automatic failover
async function callGeminiWithRotation(keys: string[], model: string, contents: any, config: any) {
  const triedKeys = keys.length > 0 ? keys : [process.env.GEMINI_API_KEY || ""];
  const errors: string[] = [];

  for (let i = 0; i < triedKeys.length; i++) {
    const key = triedKeys[i];
    if (!key) {
      errors.push(`第 ${i + 1} 个密钥为空，跳过`);
      continue;
    }
    try {
      const client = new GoogleGenAI({
        apiKey: key,
        httpOptions: {
          headers: {
            "User-Agent": "aistudio-build",
          },
        },
      });
      const response = await client.models.generateContent({
        model: model || "gemini-3.6-flash",
        contents: contents,
        config: config,
      });
      return {
        text: response.text || "",
        successIndex: keys.length > 0 ? i : -1,
        keyStatuses: triedKeys.map((k, idx) => ({
          index: idx,
          status: idx < i ? "exhausted" : (idx === i ? "active" : "untested"),
          error: idx < i ? errors[idx] : null,
        })),
      };
    } catch (err: any) {
      console.error(`Gemini Key ${i} 尝试失败:`, err.message || err);
      errors.push(err.message || String(err));
    }
  }
  throw new Error(`所有可用的 Gemini 密钥均已尝试但均告失败。错误信息汇总：\n` + errors.map((e, idx) => `[密钥 #${idx + 1}]: ${e}`).join("\n"));
}

// Helper for OpenAI/DeepSeek call with rotation & automatic failover
async function callOpenAIWithRotation(keys: string[], baseUrl: string, model: string, messages: any[], temperature: number) {
  const isDeepSeek = baseUrl.includes("deepseek");
  const triedKeys = keys.length > 0 ? keys : [(isDeepSeek ? process.env.DEEPSEEK_API_KEY : process.env.OPENAI_API_KEY) || ""];
  const errors: string[] = [];

  for (let i = 0; i < triedKeys.length; i++) {
    const key = triedKeys[i];
    if (!key) {
      errors.push(`第 ${i + 1} 个密钥为空，跳过`);
      continue;
    }
    try {
      const response = await fetch(`${baseUrl}/chat/completions`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Authorization": `Bearer ${key}`,
        },
        body: JSON.stringify({
          model: model,
          messages: messages,
          temperature: temperature,
        }),
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errText}`);
      }

      const resData: any = await response.json();
      const text = resData.choices?.[0]?.message?.content || "";
      return {
        text: text,
        successIndex: keys.length > 0 ? i : -1,
        keyStatuses: triedKeys.map((k, idx) => ({
          index: idx,
          status: idx < i ? "exhausted" : (idx === i ? "active" : "untested"),
          error: idx < i ? errors[idx] : null,
        })),
      };
    } catch (err: any) {
      console.error(`API Key ${i} 在 ${baseUrl} 尝试失败:`, err.message || err);
      errors.push(err.message || String(err));
    }
  }
  throw new Error(`所有可用的 ${model} 密钥均已尝试但均告失败。错误信息汇总：\n` + errors.map((e, idx) => `[密钥 #${idx + 1}]: ${e}`).join("\n"));
}

// Office AI Chat endpoint
app.post("/api/office-ai/chat", async (req, res) => {
  try {
    const {
      provider,
      model,
      prompt,
      history,
      systemInstruction,
      temperature,
      keys,
      selectionData,
      customUrl,
    } = req.body;

    let contextPrompt = "";
    if (selectionData) {
      contextPrompt += `【Excel 选中单元格区域内容及上下文结构】:\n区域范围: ${selectionData.address}\n选区数据:\n${JSON.stringify(selectionData.rows, null, 2)}\n\n`;
    }
    contextPrompt += `【用户提问/算料代码生成指令】: ${prompt}`;

    if (provider === "gemini") {
      const contents: any[] = [];
      if (history && Array.isArray(history)) {
        history.forEach((msg: any) => {
          const role = msg.role === "assistant" || msg.role === "model" ? "model" : "user";
          contents.push({
            role: role,
            parts: [{ text: msg.text || msg.content || "" }],
          });
        });
      }
      contents.push({
        role: "user",
        parts: [{ text: contextPrompt }],
      });

      const config: any = {
        temperature: temperature !== undefined ? Number(temperature) : 0.7,
      };
      if (systemInstruction) {
        config.systemInstruction = systemInstruction;
      }

      const result = await callGeminiWithRotation(keys || [], model, contents, config);
      return res.json(result);
    } else {
      let baseUrl = "https://api.openai.com/v1";
      if (provider === "deepseek") {
        baseUrl = "https://api.deepseek.com";
      } else if (provider === "custom" && customUrl) {
        baseUrl = customUrl.replace(/\/chat\/completions$/, "");
      }

      const messages: any[] = [];
      if (systemInstruction) {
        messages.push({ role: "system", content: systemInstruction });
      }
      if (history && Array.isArray(history)) {
        history.forEach((msg: any) => {
          const role = msg.role === "assistant" || msg.role === "model" ? "assistant" : "user";
          messages.push({
            role: role,
            content: msg.text || msg.content || "",
          });
        });
      }
      messages.push({ role: "user", content: contextPrompt });

      const result = await callOpenAIWithRotation(
        keys || [],
        baseUrl,
        model,
        messages,
        temperature !== undefined ? Number(temperature) : 0.7
      );
      return res.json(result);
    }
  } catch (error: any) {
    console.error("Office AI Chat processing failure:", error);
    res.status(500).json({ error: error.message || "无法处理 AI 聊天生成请求" });
  }
});

// Serve Vite-built web application
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
