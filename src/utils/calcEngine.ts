import { ProcessedItem, FormulaRule, ColumnMappings } from "../types";
import * as XLSX from "xlsx";

/**
 * Evaluates whether a target text matches a query supporting AND/OR logic and wildcards.
 * Supports:
 * - OR logic: |, ||, "或", "or" (case-insensitive)
 * - AND logic: &, &&, "且", "and" (case-insensitive)
 * - Wildcards: * (match zero or more chars), ? (match exactly one char)
 */
export function matchSearchQuery(text: string, query: string, isExactMatch?: boolean): boolean {
  if (!query || !query.trim()) return true;
  if (!text) return false;

  const cleanText = text.trim();

  if (isExactMatch) {
    return cleanText.toLowerCase() === query.trim().toLowerCase();
  }

  // Split by OR delimiters
  const orParts = query.split(/\s*(?:\|\||\||或|\b[Oo][Rr]\b)\s*/);

  // If any OR section is fully satisfied, then it matches
  return orParts.some(orPart => {
    if (!orPart.trim()) return false;
    
    // Split the OR part by AND delimiters
    const andParts = orPart.split(/\s*(?:&&|&|且|\b[Aa][Nn][Dd]\b)\s*/);
    
    // All AND parts must be satisfied
    return andParts.every(andPart => {
      const term = andPart.trim();
      if (!term) return true; // Ignore trailing/leading empty terms safely

      // Check if this term uses wildcards
      if (term.includes('*') || term.includes('?')) {
        // Convert wildcard pattern to regex
        // Escape regex characters except * and ?
        const escaped = term.replace(/[.+^${}()|[\]\\]/g, '\\$&');
        const regexStr = '^' + escaped.replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
        const regex = new RegExp(regexStr, 'i');
        return regex.test(cleanText);
      } else {
        // Fall back to simple substring match
        return cleanText.toLowerCase().includes(term.toLowerCase());
      }
    });
  });
}

/**
 * Simplifies technical formula logic strings (like SUM_LEN, ROW_VAL, etc.) into a modern Chinese readable format.
 */
export function simplifyFormulaDisplay(expr: string): string {
  if (!expr) return "";
  let simplified = expr;

  // Replace SUM_LEN('xxx') or SUM_LEN("xxx") with 型材长('xxx')
  simplified = simplified.replace(/SUM_LEN\((["'])(.*?)\1\)/gi, "型材长('$2')");

  // Replace SUM_QTY('xxx') with 数量('xxx')
  simplified = simplified.replace(/SUM_QTY\((["'])(.*?)\1\)/gi, "数量('$2')");

  // Replace SUM_PERIMETER('xxx') with 面板周长('xxx')
  simplified = simplified.replace(/SUM_PERIMETER\((["'])(.*?)\1\)/gi, "面板周长('$2')");

  // Replace SUM_AREA('xxx') with 面板面积('xxx')
  simplified = simplified.replace(/SUM_AREA\((["'])(.*?)\1\)/gi, "面板面积('$2')");

  // Replace ROW_VAL('profile', 'name', '公立柱', 'cuttingSize', 'true') with 型材「公立柱」的尺寸 or similar
  const rowValRegex = /ROW_VAL\((["'])(.*?)\1,\s*(["'])(.*?)\3,\s*(["'])(.*?)\5,\s*(["'])(.*?)\7(?:,\s*(["'])(true|false)\9)?\)/gi;
  simplified = simplified.replace(rowValRegex, (match, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10) => {
    let categoryName = '项目';
    const cat = String(p2).toLowerCase();
    if (cat === 'profile') categoryName = '型材';
    else if (cat === 'panel') categoryName = '面板';
    else if (cat === 'steel') categoryName = '钢件';

    let targetFieldName = p8;
    if (p8 === 'cuttingSize' || p8 === 'cuttingWidth' || p8 === 'cuttingHeight') targetFieldName = '尺寸';
    else if (p8 === 'unitUsage') targetFieldName = '用量';
    else if (p8 === 'l1') targetFieldName = 'L1值';
    else if (p8 === 'l2') targetFieldName = 'L2值';

    const exactTag = p10 === 'true' ? '【精确】' : '';
    return `${categoryName}${exactTag}「${p6}」的${targetFieldName}`;
  });

  // Simplify MATCH('W', '>=1.5') or MATCH('板块编号', 'D01') to 匹配(W = >=1.5) or 匹配(板块编号 = D01)
  const matchSimplifyRegex = /MATCH(?:_PARAM|_PLATE)?\(\s*(["'])(.*?)\1\s*,\s*(["'])(.*?)\3\s*\)/gi;
  simplified = simplified.replace(matchSimplifyRegex, (match, quote1, paramName, quote2, pattern) => {
    return `匹配(${paramName} = ${pattern})`;
  });

  const matchSimplifyRegexNoQuotes = /MATCH(?:_PARAM|_PLATE)?\(\s*(["'])(.*?)\1\s*,\s*([-+]?[\d.]+)\s*\)/gi;
  simplified = simplified.replace(matchSimplifyRegexNoQuotes, (match, quote1, paramName, val) => {
    return `匹配(${paramName} = ${val})`;
  });

  return simplified;
}

/**
 * Checks if a trimmed cell value in column A or B represents a valid category name label
 * (e.g. "铝型材", "钢件", "面板") rather than standard item serialization numbers ("1", "1.1", "序号", "A-1" etc.)
 */
export function isCategoryLabel(val: string): boolean {
  if (!val) return false;
  const clean = val.replace(/\s+/g, "").trim();
  if (!clean) return false;

  // If it's a pure number or sequence of digits (e.g. "1", "12", "03", "57")
  if (/^\d+(\.\d+)?$/.test(clean)) return false;

  // If it's a common table header or serial prefix (case insensitive)
  const lowercase = clean.toLowerCase();
  if (
    lowercase === "序号" || 
    lowercase === "no" || 
    lowercase === "no." || 
    lowercase === "id" || 
    lowercase === "index" ||
    lowercase === "名称" ||
    lowercase === "品名"
  ) {
    return false;
  }

  // If it consists entirely of letters, digits, and basic punctuation / serial symbols
  // and is very short, e.g. "1-1", "a", "c8", "group1", "no.1"
  if (/^[a-zA-Z0-9.\-()（）]+$/.test(clean) && clean.length <= 4) {
    return false;
  }

  return true;
}

/**
 * Searches the worksheet and rawRows for a keyword within Column B (index 1) or Column A (index 0), returning
 * the row boundaries (0-indexed start and end rows) for the matching segment.
 * It supports both merged cells (ws['!merges']) and flat consecutive row scans with serial filter safety.
 */
export function resolveRangeByBColumnKeyword(
  ws: XLSX.WorkSheet | null,
  rawRows: string[][],
  keyword: string
): { startRow: number; endRow: number } | null {
  if (!keyword) return null;
  const cleanKeyword = keyword.replace(/\s+/g, "").toLowerCase();

  let startRow = -1;
  let mergeEndRow = -1;

  // 1. Try merges from SheetJS worksheet to locate the starting row
  if (ws && ws["!merges"]) {
    for (const merge of ws["!merges"]) {
      // Check if merge overlaps Column A (index 0) or Column B (index 1)
      if (merge.s.c <= 1 && merge.e.c >= 0) {
        let foundMatch = false;
        for (let r = merge.s.r; r <= merge.e.r; r++) {
          for (let c = merge.s.c; c <= merge.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            const cell = ws[cellRef];
            if (cell && (cell.v !== undefined || cell.w !== undefined)) {
              const cellVal = String(cell.v || cell.w || "").replace(/\s+/g, "").toLowerCase();
              if (cellVal && (cellVal.includes(cleanKeyword) || cleanKeyword.includes(cellVal))) {
                foundMatch = true;
                break;
              }
            }
          }
          if (foundMatch) break;
        }

        if (foundMatch) {
          startRow = merge.s.r;
          mergeEndRow = merge.e.r;
          break; // Stop looking, we found the matching merge block
        }
      }
    }
  }

  // 2. Fallback: Search in rawRows directly if no merge cell was matched
  if (startRow === -1) {
    for (let r = 0; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (row) {
        if (row[1]) {
          const val = String(row[1]).replace(/\s+/g, "").toLowerCase();
          if (val && (val.includes(cleanKeyword) || cleanKeyword.includes(val))) {
            startRow = r;
            break;
          }
        }
        if (row[0]) {
          const val = String(row[0]).replace(/\s+/g, "").toLowerCase();
          if (val && (val.includes(cleanKeyword) || cleanKeyword.includes(val))) {
            startRow = r;
            break;
          }
        }
      }
    }
  }

  // 3. Scan downward from startRow to find the true end of the category's items/rows
  if (startRow !== -1) {
    if (mergeEndRow !== -1) {
      return { startRow, endRow: mergeEndRow };
    }
    let endRow = startRow;

    for (let r = startRow + 1; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row) continue;

      const bVal = row[1] ? String(row[1]).trim() : "";
      const aVal = row[0] ? String(row[0]).trim() : "";

      // If B or A has a new non-empty value that qualifies as a category label (not a serial number),
      // and it does NOT contain/match the current category keyword, then a new section has started. We break!
      if (bVal !== "" || aVal !== "") {
        const bIsCategory = isCategoryLabel(bVal);
        const aIsCategory = isCategoryLabel(aVal);

        let shouldStop = false;

        if (bVal !== "" && bIsCategory) {
          const bValLower = bVal.replace(/\s+/g, "").toLowerCase();
          const bMatches = bValLower.includes(cleanKeyword) || cleanKeyword.includes(bValLower);
          
          // Only break if we are outside the matched merge area
          if (!bMatches && r > mergeEndRow) {
            shouldStop = true;
          }
        }

        if (aVal !== "" && aIsCategory) {
          const aValLower = aVal.replace(/\s+/g, "").toLowerCase();
          const aMatches = aValLower.includes(cleanKeyword) || cleanKeyword.includes(aValLower);
          if (!aMatches && r > mergeEndRow) {
            shouldStop = true;
          }
        }

        if (shouldStop) {
          break;
        }
      }

      // If the row is not completely empty (checking if there is a name at col C / Col Index 2,
      // or drawingNo at col D, or any basic content)
      const rowName = row[2] ? String(row[2]).trim() : "";
      const rowDrawing = row[3] ? String(row[3]).trim() : "";
      if (rowName || rowDrawing || bVal || aVal) {
        endRow = Math.max(endRow, r);
      }
    }
    return { startRow, endRow };
  }

  return null;
}

/**
 * Clears excel error values (e.g., #N/A, #VALUE!, spill errors) and negative numbers or zero values by converting them to ""
 */
export function sanitizeExcelValue(val: any): string {
  if (val === undefined || val === null) return "";

  // Treat actual negative numbers as empty cells (e.g. array overflow errors or formula errors)
  if (typeof val === "number" && val < 0) {
    return "";
  }

  const str = String(val).trim();
  const lower = str.toLowerCase();

  // If a string represents a negative number (starts with "-" followed by digit or is a negative number)
  if (str.startsWith("-")) {
    return "";
  }

  // Common Excel error patterns (case-insensitive) including spill / calc / array overflow errors
  if (
    lower === "#n/a" ||
    lower === "#ref!" ||
    lower === "#value!" ||
    lower === "#div/0!" ||
    lower === "#num!" ||
    lower === "#name?" ||
    lower === "#null!" ||
    lower === "#spill!" ||
    lower === "#spill" ||
    lower === "#calc!" ||
    lower.includes("#n/a") ||
    lower.includes("spill") ||
    lower.includes("#value") ||
    lower.includes("#ref") ||
    lower.includes("#calc")
  ) {
    return "";
  }

  // Check if value is a pure 0, 0.0, 0.00, etc.
  const num = Number(str);
  if (!isNaN(num) && num === 0) {
    return "";
  }

  return str;
}

/**
 * Parses coordinate ranges like "C8:L57" into 0-indexed start and end row numbers.
 * Returns null if the format is invalid.
 */
export function parseExcelRange(rangeStr: string): { startRow: number; endRow: number } | null {
  if (!rangeStr) return null;
  const match = rangeStr.toUpperCase().trim().match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;
  
  const startRow = parseInt(match[2], 10) - 1;
  const endRow = parseInt(match[4], 10) - 1;
  
  return {
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow)
  };
}

/**
 * Parses coordinate cell references like "A3" or "G15" into 0-indexed row and col numbers.
 * Returns null if format is invalid.
 */
export function parseCellCoordinate(coord: string): { row: number; col: number } | null {
  if (!coord) return null;
  const match = coord.toUpperCase().trim().match(/^([A-Z]{1,3})(\d+)$/);
  if (!match) return null;
  const colStr = match[1];
  const rowStr = match[2];
  
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col = col - 1; // 0-indexed
  
  const row = parseInt(rowStr, 10) - 1; // 0-indexed
  
  return { row, col };
}

/**
 * Heuristically classifies a row based on keyword matches in name, spec, or cell text
 */
export function classifyRowHeuristic(
  name: string,
  spec: string,
  categoryText: string,
  remark: string
): 'profile' | 'panel' | 'steel' | 'gasket' | 'fastener' | 'auxiliary' | 'other' {
  const combined = `${name} ${spec} ${categoryText} ${remark}`.toLowerCase();
  
  if (combined.includes("型材") || combined.includes("立柱") || combined.includes("横梁") || combined.includes("铝材") || /^h\d+-/i.test(spec)) {
    return "profile";
  }
  if (combined.includes("玻璃") || combined.includes("面板") || combined.includes("饰面板") || combined.includes("岩棉板") || combined.includes("铝单板") || combined.includes("石材")) {
    return "panel";
  }
  if (combined.includes("钢件") || combined.includes("钢板") || combined.includes("角码") || combined.includes("挂件") || combined.includes("槽钢") || combined.includes("预埋件")) {
    return "steel";
  }
  if (combined.includes("胶条") || combined.includes("密封条") || combined.includes("三元乙丙") || combined.includes("硅橡胶条")) {
    return "gasket";
  }
  if (combined.includes("螺栓") || combined.includes("自攻钉") || combined.includes("铆钉") || combined.includes("垫圈") || combined.includes("螺母") || combined.includes("紧固件") || combined.includes("丝杆")) {
    return "fastener";
  }
  if (combined.includes("胶") || combined.includes("密封胶") || combined.includes("单面贴") || combined.includes("海绵") || combined.includes("双组份") || combined.includes("垫块") || combined.includes("发泡棒") || combined.includes("止水海绵") || combined.includes("辅材") || combined.includes("保温塞")) {
    return "auxiliary";
  }
  
  return "other";
}

/**
 * Extract numerical dimensions from common specification string formats like "1200*1500", "100*30*5", "30*25"
 * Returns an array of numbers representing size dimensions.
 */
export function extractDimensions(spec: string): number[] {
  if (!spec) return [];
  // Standardize delimiters: replace "x", "X", ",", "，" with "*" and extract sequence of numbers (floats/ints)
  const normalized = spec.replace(/[xX,，]/g, '*');
  const matches = normalized.match(/\d+(\.\d+)?/g);
  if (!matches) return [];
  return matches.map(Number);
}

/**
 * Calculates panel perimeter in Meters from spec dimensions (assuming mm units) times quantity
 */
export function calculatePanelPerimeter(spec: string, qty: number, lengthAdjustment: number = 0): { perimeter: number; log: string } {
  const dims = extractDimensions(spec);
  if (dims.length >= 2) {
    const w = dims[0] + lengthAdjustment;
    const h = dims[1] + lengthAdjustment;
    // Perimeter = 2 * (w + h) in mm, converted to meters: 2 * (w + h) / 1000
    const itemPerimeter = (2 * (w + h)) / 1000;
    const total = itemPerimeter * qty;
    const adjStr = lengthAdjustment > 0 ? ` (已加长${lengthAdjustment}mm -> ${w}x${h}mm)` : "";
    return {
      perimeter: total,
      log: `[规格: ${dims[0]}x${dims[1]}mm${adjStr}, 单个周长: ${itemPerimeter.toFixed(3)}m, 数量: ${qty} -> 总周长: ${total.toFixed(3)}m]`
    };
  }
  return { perimeter: 0, log: `[规格 '${spec}' 未能提取足额尺寸维度计算周长]` };
}

/**
 * Calculates panel area in Square Meters from spec dimensions (assuming mm units) times quantity
 */
export function calculatePanelArea(spec: string, qty: number, lengthAdjustment: number = 0): { area: number; log: string } {
  const dims = extractDimensions(spec);
  if (dims.length >= 2) {
    const w = dims[0] + lengthAdjustment;
    const h = dims[1] + lengthAdjustment;
    // Area = w * h in mm2, converted to m2: (w * h) / 1,000,000
    const itemArea = (w * h) / 1000000;
    const total = itemArea * qty;
    const adjStr = lengthAdjustment > 0 ? ` (已加长${lengthAdjustment}mm -> ${w}x${h}mm)` : "";
    return {
      area: total,
      log: `[规格: ${dims[0]}x${dims[1]}mm${adjStr}, 单个面积: ${itemArea.toFixed(4)}㎡, 数量: ${qty} -> 总面积: ${total.toFixed(4)}㎡]`
    };
  }
  return { area: 0, log: `[规格 '${spec}' 未能提取足额尺寸维度计算面积]` };
}

/**
 * Helper to retrieve parameter values for MATCH calculations
 */
export function getParamValue(paramName: string, vars: Record<string, any>): any {
  if (!vars) return undefined;
  const name = paramName.trim().toLowerCase();
  if (name === "w" || name === "宽度" || name === "板宽" || name === "板块宽度") {
    return vars.W;
  }
  if (name === "h" || name === "高度" || name === "板高" || name === "板块高度") {
    return vars.H;
  }
  if (name === "w1" || name === "副宽" || name === "板副宽" || name === "板块副宽") {
    return vars.W1;
  }
  if (
    name === "plateno" ||
    name === "platecode" ||
    name === "platemodel" ||
    name === "板块编号" ||
    name === "板块型号" ||
    name === "编号" ||
    name === "型号" ||
    name === "工作表" ||
    name === "sheetname"
  ) {
    return vars.plateNo || vars.plateCode || vars.plateModel || vars.sheetName || "";
  }
  if (vars[paramName] !== undefined) {
    return vars[paramName];
  }
  const exactKey = Object.keys(vars).find(k => k.toLowerCase() === name);
  if (exactKey) {
    return vars[exactKey];
  }
  return undefined;
}

/**
 * Helper to check pattern match (numeric comparisons or wildcard/logic strings)
 */
export function checkParamMatch(paramValue: any, pattern: string): boolean {
  if (paramValue === undefined || paramValue === null) return false;
  const cleanPattern = pattern.trim();
  
  // Check if pattern is a numeric comparison like >=1.5 or <2.0
  const compMatch = cleanPattern.match(/^([>=<!]+)\s*([-+]?[\d.]+)/);
  if (compMatch) {
    const op = compMatch[1];
    const val = parseFloat(compMatch[2]);
    const numParam = parseFloat(String(paramValue));
    if (isNaN(numParam) || isNaN(val)) return false;
    
    if (op === ">=") return numParam >= val;
    if (op === "<=") return numParam <= val;
    if (op === ">") return numParam > val;
    if (op === "<") return numParam < val;
    if (op === "==" || op === "=") return numParam === val;
    if (op === "!=") return numParam !== val;
  }
  
  // Otherwise use advanced matchSearchQuery with wildcard and OR/AND support
  return matchSearchQuery(String(paramValue), cleanPattern, false);
}

/**
 * Evaluates a mathematical formula expression against list of materials
 */
export function evaluateFormula(
  expression: string,
  allItems: ProcessedItem[],
  projectVars: Record<string, any> = {},
  lengthAdjustment: number = 0,
  rawRows?: string[][]
): { value: number; logs: string[] } {
  const logs: string[] = [];
  
  if (lengthAdjustment > 0) {
    logs.push(`⚙️ 启用加长调整: 所有匹配的型材及面板长度加长 ${lengthAdjustment}mm`);
  }

  // Standardize full-width and Chinese characters/punctuation to standard half-width ASCII counterparts
  const standardizedBase = (expression || "")
    .replace(/[\uff08\u3010\u3008\u300a\u300c\u300e\u3014（【\[\{]/g, "(")
    .replace(/[\uff09\u3011\u3009\u300b\u300d\u300f\u3015）】\]\}]/g, ")")
    .replace(/[\u00d7×*]/g, "*")
    .replace(/[\u00f7÷\/／]/g, "/")
    .replace(/[\uff0b＋+]/g, "+")
    .replace(/[\uff0d－\-]/g, "-")
    .replace(/，/g, ",")
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'");

  let evaluatedExpr = standardizedBase;

  // 0. Solve MATCH('param', 'pattern') -> once matched successfully, output is 1, else 0
  const matchPattern = /MATCH(?:_PARAM|_PLATE)?\(\s*(["'])(.*?)\1\s*,\s*(["']?)(.*?)\3\s*\)/gi;
  if (matchPattern.test(evaluatedExpr)) {
    evaluatedExpr = evaluatedExpr.replace(matchPattern, (fullCall, quote1, paramName, quote2, pattern) => {
      const paramValue = getParamValue(paramName, projectVars);
      const isMatched = checkParamMatch(paramValue, pattern);
      const resultValue = isMatched ? 1 : 0;
      
      logs.push(`评估匹配函数 ${fullCall}: 提取参数 '${paramName}' = ${paramValue}, 匹配模式 '${pattern}' -> ${isMatched ? "成功 (1)" : "失败 (0)"}`);
      return String(resultValue);
    });
  }

  // Substitute project variables (e.g., W, H, W1)
  Object.keys(projectVars).forEach(key => {
    // Only match standalone variable names with word boundaries to avoid replacing "W" inside "WITH" for example.
    const regex = new RegExp(`\\b${key}\\b`, 'gi');
    if (regex.test(evaluatedExpr)) {
      evaluatedExpr = evaluatedExpr.replace(regex, String(projectVars[key]));
      logs.push(`提取工程参量参数 '${key}' = ${projectVars[key]}`);
    }
  });

  // Substitute direct cell references (e.g., A3, B4, G15)
  // Pattern matches single or double quoted strings first so they are skipped.
  // Group 1 matches column letters, Group 2 matches row number.
  const cellRefPattern = /"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|\b([A-Z]{1,3})(\d+)\b/gi;
  evaluatedExpr = evaluatedExpr.replace(cellRefPattern, (match, colLetters, rowNumber) => {
    if (colLetters === undefined || rowNumber === undefined) {
      return match; // Quoted string literal, return unchanged
    }
    const cellRef = match.toUpperCase();
    let cellValue = 0;
    
    if (rawRows && rawRows.length > 0) {
      const coord = parseCellCoordinate(cellRef);
      if (coord) {
        const { row, col } = coord;
        if (row >= 0 && row < rawRows.length && col >= 0 && col < rawRows[row].length) {
          const rawVal = rawRows[row][col];
          if (rawVal !== undefined && rawVal !== null) {
            const cleaned = String(rawVal).trim().replace(/,/g, '');
            const numMatch = cleaned.match(/[-+]?[0-9]*\.?[0-9]+/);
            if (numMatch) {
              cellValue = Number(numMatch[0]);
            }
          }
        }
      }
      
      const originalVal = cellValue;
      let adjLog = "";
      if (lengthAdjustment > 0) {
        cellValue += lengthAdjustment;
        adjLog = ` [含加长+${lengthAdjustment}mm -> ${cellValue}mm]`;
      }
      const convertedVal = cellValue / 1000;
      logs.push(`提取引用单元格 [${cellRef}] 的数据 = ${originalVal}mm${adjLog} -> 自动换算为米 = ${convertedVal}m`);
      cellValue = convertedVal;
    } else {
      logs.push(`提取引用单元格 [${cellRef}] 失败 (工作表数据未加载)，默认值为 0`);
    }
    
    return String(cellValue);
  });

  // 1. Solve SUM_LEN(profile_spec_prefix)
  // Matches SUM_LEN("...") or SUM_LEN('...')
  const sumLenRegex = /SUM_LEN\((["'])(.*?)\1\)/gi;
  let lenMatch;
  while ((lenMatch = sumLenRegex.exec(standardizedBase)) !== null) {
    const fullCall = lenMatch[0];
    const prefix = lenMatch[2].toLowerCase().trim();
    
    // Find all profiles, steel, or fastener items matching spec prefix (e.g. H2-JT, Q235) or match name containing prefix
    const matchedProfiles = allItems.filter(item => 
      (item.category === "profile" || item.category === "steel" || item.category === "fastener") && 
      (matchSearchQuery(item.spec, prefix) || matchSearchQuery(item.name, prefix))
    );
    
    let sumLength = 0;
    const details: string[] = [];
    
    matchedProfiles.forEach(p => {
      // Total profile length = single length * qty (if single length exists, otherwise we default single length to 1 and qty represents total length, e.g. 13.72m)
      let rowLen = p.len > 0 ? p.len : 1;
      const rowQty = p.qty > 0 ? p.qty : 1;
      
      let itemTotalLen = 0;
      let adjStr = "";
      if (p.len > 0) {
        if (lengthAdjustment > 0) {
          rowLen = rowLen + (lengthAdjustment / 1000);
          adjStr = ` [含加长+${lengthAdjustment}mm]`;
        }
        itemTotalLen = rowLen * rowQty;
      } else {
        // If no explicit len column, qty might be the meter length. If there is a length adjustment, apply it per unit
        let rawSingleLen = 1;
        if (lengthAdjustment > 0) {
          rawSingleLen = 1 + (lengthAdjustment / 1000);
          adjStr = ` [含无长型材加算+${lengthAdjustment}mm]`;
        }
        itemTotalLen = rawSingleLen * rowQty;
      }
      sumLength += itemTotalLen;
      
      details.push(`${p.name}(${p.spec}): 单长${rowLen.toFixed(3)}m${adjStr} x 数量${rowQty} = ${itemTotalLen.toFixed(2)}m (行 ${p.rowIndex + 1})`);
    });

    if (matchedProfiles.length > 0) {
      logs.push(`SUM_LEN("${prefix}") 匹配到 ${matchedProfiles.length} 行型材: ${details.join(", ")} | 总长度: ${sumLength.toFixed(2)}m`);
    } else {
      logs.push(`SUM_LEN("${prefix}") 未能匹配到任何型材`);
    }
    
    // Safety check to avoid division-by-zero or NaN issues, replace in formula evaluation string
    evaluatedExpr = evaluatedExpr.replace(fullCall, String(sumLength));
  }

  // 2. Solve SUM_QTY(keyword)
  // Matches SUM_QTY("...") or SUM_QTY('...')
  const sumQtyRegex = /SUM_QTY\((["'])(.*?)\1\)/gi;
  let qtyMatch;
  while ((qtyMatch = sumQtyRegex.exec(standardizedBase)) !== null) {
    const fullCall = qtyMatch[0];
    const keyword = qtyMatch[2].toLowerCase().trim();
    
    // Match anywhere in code, name, spec, or remarks
    const matchedItems = allItems.filter(item => 
      matchSearchQuery(item.name, keyword) || 
      matchSearchQuery(item.spec, keyword) ||
      matchSearchQuery(item.material, keyword)
    );
    
    let sumQuantity = 0;
    const details: string[] = [];
    
    matchedItems.forEach(i => {
      sumQuantity += i.qty;
      details.push(`${i.name}(${i.spec}): 数量${i.qty} (行 ${i.rowIndex + 1})`);
    });

    if (matchedItems.length > 0) {
      logs.push(`SUM_QTY("${keyword}") 匹配到 ${matchedItems.length} 个项目: ${details.join(", ")} | 总数量: ${sumQuantity}`);
    } else {
      logs.push(`SUM_QTY("${keyword}") 未能匹配到任何项目`);
    }
    
    evaluatedExpr = evaluatedExpr.replace(fullCall, String(sumQuantity));
  }

  // 3. Solve SUM_PERIMETER(keyword)
  const sumPerimeterRegex = /SUM_PERIMETER\((["'])(.*?)\1\)/gi;
  let perimMatch;
  while ((perimMatch = sumPerimeterRegex.exec(standardizedBase)) !== null) {
    const fullCall = perimMatch[0];
    const keyword = perimMatch[2].toLowerCase().trim();
    
    const matchedPanels = allItems.filter(item => 
      item.category === "panel" && 
      (matchSearchQuery(item.name, keyword) || matchSearchQuery(item.spec, keyword))
    );
    
    let totalPerimeter = 0;
    const details: string[] = [];
    
    matchedPanels.forEach(p => {
      const { perimeter, log } = calculatePanelPerimeter(p.spec, p.qty, lengthAdjustment);
      totalPerimeter += perimeter;
      if (perimeter > 0) {
        details.push(`${p.name}(${p.spec}): ${log} (行 ${p.rowIndex + 1})`);
      }
    });

    if (matchedPanels.length > 0) {
      logs.push(`SUM_PERIMETER("${keyword}") 匹配到 ${matchedPanels.length} 个面板: ${details.join("; ")} | 总周长: ${totalPerimeter.toFixed(2)}m`);
    } else {
      logs.push(`SUM_PERIMETER("${keyword}") 未能匹配到任何面板`);
    }

    evaluatedExpr = evaluatedExpr.replace(fullCall, String(totalPerimeter));
  }

  // 4. Solve SUM_AREA(keyword)
  const sumAreaRegex = /SUM_AREA\((["'])(.*?)\1\)/gi;
  let areaMatch;
  while ((areaMatch = sumAreaRegex.exec(standardizedBase)) !== null) {
    const fullCall = areaMatch[0];
    const keyword = areaMatch[2].toLowerCase().trim();
    
    const matchedPanels = allItems.filter(item => 
      item.category === "panel" && 
      (matchSearchQuery(item.name, keyword) || matchSearchQuery(item.spec, keyword))
    );
    
    let totalArea = 0;
    const details: string[] = [];
    
    matchedPanels.forEach(p => {
      const { area, log } = calculatePanelArea(p.spec, p.qty, lengthAdjustment);
      totalArea += area;
      if (area > 0) {
        details.push(`${p.name}(${p.spec}): ${log} (行 ${p.rowIndex + 1})`);
      }
    });

    if (matchedPanels.length > 0) {
      logs.push(`SUM_AREA("${keyword}") 匹配到 ${matchedPanels.length} 面面板: ${details.join("; ")} | 总面积: ${totalArea.toFixed(2)}㎡`);
    } else {
      logs.push(`SUM_AREA("${keyword}") 未能匹配到任何面板`);
    }

    evaluatedExpr = evaluatedExpr.replace(fullCall, String(totalArea));
  }

  // 5. Solve ROW_VAL('category', 'searchField', 'keyword', 'targetField', 'isExactMatch')
  // Matches ROW_VAL('profile', 'name', '公立柱', 'cuttingSize', 'true') with single or double quotes
  const rowValRegex = /ROW_VAL\((["'])(.*?)\1,\s*(["'])(.*?)\3,\s*(["'])(.*?)\5,\s*(["'])(.*?)\7(?:,\s*(["'])(true|false)\9)?\)/gi;
  
  // Pre-scan all ROW_VAL calls to check if they share the exact same selector signature.
  // If they do, we should evaluate the entire expression row-by-row on the matched items and sum the results,
  // preventing math distortion from pre-summing individual components (e.g. perimeters/areas of multiple glass items).
  const freshRowValRegex = new RegExp(rowValRegex.source, "gi");
  const rowValCalls: {
    fullCall: string;
    category: string;
    searchField: string;
    keyword: string;
    targetField: string;
    isExactMatch: boolean;
  }[] = [];

  let rMatch;
  while ((rMatch = freshRowValRegex.exec(evaluatedExpr)) !== null) {
    rowValCalls.push({
      fullCall: rMatch[0],
      category: rMatch[2],
      searchField: rMatch[4],
      keyword: rMatch[6].toLowerCase().trim(),
      targetField: rMatch[8],
      isExactMatch: rMatch[10] === 'true',
    });
  }

  const allSameSelector = rowValCalls.length > 0 && rowValCalls.every(call => 
    call.category === rowValCalls[0].category &&
    call.searchField === rowValCalls[0].searchField &&
    call.keyword === rowValCalls[0].keyword &&
    call.isExactMatch === rowValCalls[0].isExactMatch
  );

  if (allSameSelector) {
    const firstCall = rowValCalls[0];
    const category = firstCall.category;
    const searchField = firstCall.searchField;
    const keyword = firstCall.keyword;
    const isExactMatch = firstCall.isExactMatch;

    // Find all matching rows in list
    const matchedItems = allItems.filter(item => {
      if (item.category !== category) return false;
      let valToMatch = "";
      if (searchField === "name") valToMatch = item.name || "";
      else if (searchField === "drawingNo") valToMatch = item.drawingNo || "";
      else if (searchField === "materialModel") valToMatch = item.materialModel || "";
      else if (searchField === "cuttingSize") valToMatch = item.cuttingSize || "";
      else if (searchField === "unitUsage") valToMatch = item.unitUsage || "";
      else if (searchField === "l1") valToMatch = item.l1 || "";
      else if (searchField === "l2") valToMatch = item.l2 || "";
      else if (searchField === "remark") valToMatch = item.remark || "";
      return matchSearchQuery(valToMatch, keyword, isExactMatch);
    });

    if (matchedItems.length > 0) {
      let totalValueSum = 0;
      const rowDetails: string[] = [];
      const rowExprs: string[] = [];

      matchedItems.forEach(item => {
        let itemExpr = evaluatedExpr;

        rowValCalls.forEach(call => {
          let rawVal = "";
          const targetField = call.targetField;
          if (targetField === "name") rawVal = item.name;
          else if (targetField === "drawingNo") rawVal = item.drawingNo;
          else if (targetField === "materialModel") rawVal = item.materialModel;
          else if (targetField === "cuttingSize" || targetField === "cuttingWidth" || targetField === "cuttingHeight") {
            const cs = item.cuttingSize || "";
            if (targetField === "cuttingSize") {
              rawVal = cs;
            } else {
              const match = cs.match(/([-+]?[\d.]+)\s*[*xX×]\s*([-+]?[\d.]+)/);
              if (match) {
                rawVal = targetField === "cuttingWidth" ? match[1] : match[2];
              } else {
                rawVal = cs;
              }
            }
          }
          else if (targetField === "unitUsage") rawVal = item.unitUsage;
          else if (targetField === "l1") rawVal = item.l1;
          else if (targetField === "l2") rawVal = item.l2;
          else if (targetField === "remark") rawVal = item.remark;
          else if (targetField === "category") rawVal = item.category;

          // Extract numeric value
          let rowNumVal = 0;
          if (rawVal) {
            const cleaned = rawVal.trim().replace(/,/g, '');
            const numMatch = cleaned.match(/[-+]?[0-9]*\.?[0-9]+/);
            if (numMatch) {
              rowNumVal = Number(numMatch[0]);
            }
          }

          let adjLog = "";
          if (lengthAdjustment > 0 && ["cuttingSize", "cuttingWidth", "cuttingHeight", "l1", "l2"].includes(targetField)) {
            const orig = rowNumVal;
            rowNumVal += lengthAdjustment;
            adjLog = `[加长加算: ${orig}mm + ${lengthAdjustment}mm -> ${rowNumVal}mm]`;
          }

          let convStr = "";
          // Unit conversion: if category is profile, steel, or fastener (五金角码/钢件类), or extracting cuttingWidth/cuttingHeight (usually in mm), convert to meters
          const isLengthCategory = category === "profile" || category === "steel" || category === "fastener";
          if ((isLengthCategory && ["cuttingSize", "l1", "l2"].includes(targetField)) || ["cuttingWidth", "cuttingHeight"].includes(targetField)) {
            const original = rowNumVal;
            rowNumVal = rowNumVal / 1000;
            convStr = ` (提取${original}mm${adjLog ? " " + adjLog : ""}→折算${Number(rowNumVal.toFixed(3))}m)`;
          } else if (adjLog) {
            convStr = ` (${adjLog})`;
          }

          // Replace this specific call in itemExpr with the resolved value
          // We use replaceAll because there might be multiple identical calls in the expression
          itemExpr = itemExpr.replaceAll(call.fullCall, String(Number(rowNumVal.toFixed(3))));
        });

        // Now evaluate itemExpr algebraically
        try {
          const sanitized = itemExpr.replace(/[^0-9.+\/*()\s-]/g, "");
          // Safely evaluate standard math formula
          // eslint-disable-next-line no-eval
          const rowResult = Function(`"use strict"; return (${sanitized})`)();
          if (!isNaN(rowResult) && isFinite(rowResult)) {
            totalValueSum += rowResult;
            rowDetails.push(`[行${item.rowIndex + 1}《${item.name || "未命名"}》: 结果 ${rowResult.toFixed(3)}]`);
          } else {
            rowDetails.push(`[行${item.rowIndex + 1}《${item.name || "未命名"}》: 结果非有效数]`);
          }
        } catch (err: any) {
          rowDetails.push(`[行${item.rowIndex + 1}《${item.name || "未命名"}》: 计算错误: ${err.message}]`);
        }

        rowExprs.push(itemExpr);
      });

      const exactStr = isExactMatch ? ' 【精确】' : '';
      const adjText = lengthAdjustment > 0 ? ` (已融合加长调整: +${lengthAdjustment}mm)` : "";
      logs.push(`ROW_VAL("${category}", "${searchField}", "${keyword}")${exactStr}${adjText} 检测到同源多参引用，启用【同源逐行算定汇总】。共匹配到 ${matchedItems.length} 行: ${rowDetails.join(", ")} | 累计总和: ${Number(totalValueSum.toFixed(3))}`);
      
      // Build mathematically equivalent combined expression to preserve detailed traceability of each row's values
      evaluatedExpr = rowExprs.map(expr => `(${expr})`).join(" + ");
    } else {
      const exactStr = isExactMatch ? ' 【精确】' : '';
      logs.push(`ROW_VAL("${category}", "${searchField}", "${keyword}")${exactStr} 未能在分类 "${category}" 的料单中匹配到任何行，结果设为 0`);
      evaluatedExpr = "0";
    }
  } else {
    // Original fallback logic
    let valMatch;
    while ((valMatch = rowValRegex.exec(standardizedBase)) !== null) {
      const fullCall = valMatch[0];
      const category = valMatch[2];
      const searchField = valMatch[4];
      const keyword = valMatch[6].toLowerCase().trim();
      const targetField = valMatch[8];
      const isExactMatch = valMatch[10] === 'true';
      
      // Find all matching rows in list
      const matchedItems = allItems.filter(item => {
        if (item.category !== category) return false;
        let valToMatch = "";
        if (searchField === "name") valToMatch = item.name || "";
        else if (searchField === "drawingNo") valToMatch = item.drawingNo || "";
        else if (searchField === "materialModel") valToMatch = item.materialModel || "";
        else if (searchField === "cuttingSize") valToMatch = item.cuttingSize || "";
        else if (searchField === "unitUsage") valToMatch = item.unitUsage || "";
        else if (searchField === "l1") valToMatch = item.l1 || "";
        else if (searchField === "l2") valToMatch = item.l2 || "";
        else if (searchField === "remark") valToMatch = item.remark || "";
        return matchSearchQuery(valToMatch, keyword, isExactMatch);
      });
      
      let resolvedValue = 0;
      if (matchedItems.length > 0) {
        const detailLogs: string[] = [];
        let totalSum = 0;
        
        matchedItems.forEach(item => {
          let rawVal = "";
          if (targetField === "name") rawVal = item.name;
          else if (targetField === "drawingNo") rawVal = item.drawingNo;
          else if (targetField === "materialModel") rawVal = item.materialModel;
          else if (targetField === "cuttingSize" || targetField === "cuttingWidth" || targetField === "cuttingHeight") {
            const cs = item.cuttingSize || "";
            if (targetField === "cuttingSize") {
              rawVal = cs;
            } else {
              const match = cs.match(/([-+]?[\d.]+)\s*[*xX×]\s*([-+]?[\d.]+)/);
              if (match) {
                rawVal = targetField === "cuttingWidth" ? match[1] : match[2];
              } else {
                rawVal = cs;
              }
            }
          }
          else if (targetField === "unitUsage") rawVal = item.unitUsage;
          else if (targetField === "l1") rawVal = item.l1;
          else if (targetField === "l2") rawVal = item.l2;
          else if (targetField === "remark") rawVal = item.remark;
          else if (targetField === "category") rawVal = item.category;
          
          // Extract numeric value
          let rowNumVal = 0;
          if (rawVal) {
            const cleaned = rawVal.trim().replace(/,/g, '');
            const numMatch = cleaned.match(/[-+]?[0-9]*\.?[0-9]+/);
            if (numMatch) {
              rowNumVal = Number(numMatch[0]);
            }
          }
          
          let adjLog = "";
          if (lengthAdjustment > 0 && ["cuttingSize", "cuttingWidth", "cuttingHeight", "l1", "l2"].includes(targetField)) {
            const orig = rowNumVal;
            rowNumVal += lengthAdjustment;
            adjLog = `[加长加算: ${orig}mm + ${lengthAdjustment}mm -> ${rowNumVal}mm]`;
          }

          let convStr = "";
          // Unit conversion: if category is profile, steel, or fastener (五金角码/钢件类), or extracting cuttingWidth/cuttingHeight (usually in mm), convert to meters
          const isLengthCategory = category === "profile" || category === "steel" || category === "fastener";
          if ((isLengthCategory && ["cuttingSize", "l1", "l2"].includes(targetField)) || ["cuttingWidth", "cuttingHeight"].includes(targetField)) {
            const original = rowNumVal;
            rowNumVal = rowNumVal / 1000;
            convStr = ` (提取${original}mm${adjLog ? " " + adjLog : ""}→折算${Number(rowNumVal.toFixed(3))}m)`;
          } else if (adjLog) {
            convStr = ` (${adjLog})`;
          }
          
          totalSum += rowNumVal;
          detailLogs.push(`[第${item.rowIndex + 1}行: 主键:${item.name || ""} 物理量:${rawVal || "0"}${convStr}]`);
        });
        
        resolvedValue = totalSum;
        const exactStr = isExactMatch ? ' 【精确】' : '';
        const adjText = lengthAdjustment > 0 ? ` (已加长: +${lengthAdjustment}mm)` : "";
        logs.push(`ROW_VAL("${category}", "${searchField}", "${keyword}", "${targetField}", "${isExactMatch}")${exactStr}${adjText} 共检索到 ${matchedItems.length} 行匹配, 原始参数分布: ${detailLogs.join(", ")}, 累加汇总体积: ${Number(resolvedValue.toFixed(3))}`);
      } else {
        const exactStr = isExactMatch ? ' 【精确】' : '';
        logs.push(`ROW_VAL("${category}", "${searchField}", "${keyword}", "${targetField}", "${isExactMatch}")${exactStr} 未能在分类 "${category}" 的料单中匹配到任何 "${searchField}" 为 "${keyword}" 的行`);
      }
      
      evaluatedExpr = evaluatedExpr.replace(fullCall, String(Number(resolvedValue.toFixed(3))));
    }
  }

  // Evaluate the standard algebraic expression safely
  try {
    // Sanitize mathematical characters: only allow numbers, operators, standard parentheses, and whitespace
    const sanitized = evaluatedExpr.replace(/[^0-9.+\/*()\s-]/g, "");
    if (!sanitized.trim()) {
      return { value: 0, logs: [...logs, "计算表达式解析为空"] };
    }
    
    // Safely evaluate standard math formula
    // eslint-disable-next-line no-eval
    const result = Function(`"use strict"; return (${sanitized})`)();
    
    if (isNaN(result) || !isFinite(result)) {
      return { value: 0, logs: [...logs, `计算出错: 结果不是有效数字 (解析式: ${sanitized})`] };
    }

    return {
      value: Math.round(result * 1000) / 1000, // Round to 3 decimals
      logs: [...logs, `解析表达式: ${expression} -> 计算表达式: ${evaluatedExpr} -> 结果: ${result.toFixed(3)}`]
    };
  } catch (err: any) {
    return {
      value: 0,
      logs: [...logs, `公式解析/数学计算错误: ${err.message} (解析度: ${evaluatedExpr})`]
    };
  }
}

/**
 * Default formula templates that will automatically calculate auxiliary materials based on standard industry logic
 */
export const DEFAULT_FORMULAS: FormulaRule[] = [
  {
    id: "gasket-jt19",
    name: "胶条H2-JT19公式",
    targetCategory: "gasket",
    targetKeyword: "H2-JT19",
    expression: "SUM_LEN('H2-JT19') * 1.0",
    description: "按照型材H2-JT19的长度等比例配套胶条数量"
  },
  {
    id: "gasket-common",
    name: "胶条长度等比配套",
    targetCategory: "gasket",
    targetKeyword: "胶条",
    expression: "SUM_LEN('H2-JT') * 1.0",
    description: "按照包含'H2-JT'型材的总长度配套胶条长(系数1.0)"
  },
  {
    id: "fastener-beam",
    name: "自攻钉横梁配套公式",
    targetCategory: "fastener",
    targetKeyword: "ST6.3",
    expression: "SUM_QTY('自攻钉') * 1.0",
    description: "自攻钉按匹配基数等比套算"
  },
  {
    id: "fastener-screw",
    name: "盘头自攻钉配套公式",
    targetCategory: "fastener",
    targetKeyword: "ST4.8*13",
    expression: "SUM_QTY('角码') * 4",
    description: "固定角码紧固件计算: 每个角码需配备4只自攻针"
  },
  {
    id: "gasket-dual-run",
    name: "中空玻璃室内外双向胶条",
    targetCategory: "gasket",
    targetKeyword: "H2-JT04",
    expression: "SUM_LEN('H2-JT04') * 2.0",
    description: "中空玻璃打胶槽室内与室外一式两份，需要2倍配套"
  },
  {
    id: "aux-sealant-perimeter",
    name: "耐候密封胶按面板周长统计",
    targetCategory: "auxiliary",
    targetKeyword: "耐候密封胶",
    expression: "SUM_PERIMETER('玻璃') * 0.15",
    description: "按中空玻璃周长缝深缝宽计算总用量 (体积估算系数0.15L/m)"
  },
  {
    id: "aux-insulation",
    name: "保温岩棉按面板面积配套",
    targetCategory: "auxiliary",
    targetKeyword: "岩棉",
    expression: "SUM_AREA('玻璃') * 1.05",
    description: "按玻璃面板尺寸计岩棉防护面积(损耗系数1.05)"
  },
  {
    id: "aux-water-foam",
    name: "止水海绵数量配套",
    targetCategory: "auxiliary",
    targetKeyword: "止水海绵",
    expression: "SUM_QTY('上极梁') * 1",
    description: "上楣止水海绵，每个上楣梁位置配套1块海绵"
  },
  {
    id: "aux-nylon-bracket",
    name: "尼龙垫圈固定螺栓配套",
    targetCategory: "auxiliary",
    targetKeyword: "尼龙垫",
    expression: "SUM_QTY('挂件') * 2",
    description: "挂件螺丝垫片计算：每个挂件紧固处需要2个尼龙垫"
  }
];

/**
 * Robust check to determine if an item is an auxiliary material (fastener, gasket, sealant/glue, etc.)
 * based on its name, specification, or remark.
 */
export function isAuxiliaryItem(name: string, materialModel: string, remark?: string): boolean {
  const n = (name || "").replace(/\s+/g, "").toLowerCase();
  const m = (materialModel || "").replace(/\s+/g, "").toLowerCase();
  const r = (remark || "").replace(/\s+/g, "").toLowerCase();

  const auxKeywords = [
    "胶条", "epdm", "密封条", "气密胶条", "三元乙丙", "双面贴", "双面胶", "发泡双面贴", "泡沫棒", "泡沫条", "聚乙烯泡沫",
    "螺栓", "螺丝", "螺母", "螺帽", "自攻钉", "自攻", "自钻", "螺钉", "垫圈", "弹垫", "平垫", "弹簧垫圈", "平垫圈", "拉铆钉", "铆钉", "销轴",
    "锚栓", "膨胀螺栓", "膨胀管", "膨胀螺丝", "螺杆", "植筋胶", "发泡剂", "发泡胶", "胶水", "美纹纸", "清毒剂", "美纹胶", "胶带",
    "密封胶", "耐候胶", "结构胶", "耐候硅酮", "硅酮密封胶", "底漆", "清洗剂"
  ];

  for (const kw of auxKeywords) {
    if (n.includes(kw) || m.includes(kw) || r.includes(kw)) {
      return true;
    }
  }

  if (
    n.includes("bolt") || n.includes("screw") || n.includes("gasket") || n.includes("sealant") ||
    m.includes("bolt") || m.includes("screw") || m.includes("gasket") || m.includes("sealant")
  ) {
    return true;
  }

  return false;
}

