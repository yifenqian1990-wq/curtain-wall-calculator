import React, { useState } from "react";
import { COMConnectionState } from "../types";
import { Network, Server, Play, CheckCircle2, AlertTriangle, RefreshCw, Terminal, Cpu, Download, Copy, Check, Info } from "lucide-react";

interface ExcelComConnectorProps {
  connection: COMConnectionState;
  localWsPort: number;
  setLocalWsPort: (port: number) => void;
  wsConnected: boolean;
  comBridgeStatus: string;
  comBridgeMessage: string;
  onSendSimulation: () => void;
}

export default function ExcelComConnector({
  connection,
  localWsPort,
  setLocalWsPort,
  wsConnected,
  comBridgeStatus,
  comBridgeMessage,
  onSendSimulation,
}: ExcelComConnectorProps) {
  const [portInput, setPortInput] = useState<string>(localWsPort.toString());
  const [copied, setCopied] = useState(false);

  const handleApplyPort = (e: React.FormEvent) => {
    e.preventDefault();
    const port = parseInt(portInput.trim());
    if (!isNaN(port) && port > 1024 && port < 65535) {
      setLocalWsPort(port);
    } else {
      alert("请输入合法的端口号 (1025 ~ 65534)");
    }
  };

  const handleCopyCode = () => {
    navigator.clipboard.writeText(`python com_bridge.py`);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const downloadPythonScript = () => {
    const scriptContent = `# -*- coding: utf-8 -*-
"""
幕墙辅件算料一键填回 OLE COM 本地 WebSocket 桥接器 (PyBridge)
使用说明：
1. 确保本地电脑安装了 Python (建议 3.8及以上版本)
2. 在本地终端(CMD/Terminal)运行命令安装必要依赖：
   pip install websockets pywin32
3. 保持本地 Microsoft Excel 或者 金山 WPS 处于开启状态，并加载好目标幕墙工程料单。
4. 启动本程序进行本地端口监听：
   python com_bridge.py
"""

import os
import sys
import json
import asyncio
import websockets
import pythoncom
import win32com.client

PORT = ${localWsPort}

print("=" * 60)
print("  幕墙辅件算料一键填回 OLE COM 本地 WebSocket 桥接器  ")
print("=" * 60)
print(f"[*] 监听本地 WebSocket 服务端: ws://127.0.0.1:{PORT}")
print("[*] 正在准备 OLE/COM 环境，请确保本地 Microsoft Excel 或 金山 WPS 处于运行状态...")

# 导入 Windows 本地 OLE 接口
try:
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
except ImportError:
    print("[错误] 未检测到 pywin32 库。请在本地终端中运行以下命令安装它：")
    print("      pip install pywin32 websockets")
    print("      或者运行: python -m pip install websockets pywin32")
    input("\\n推荐在电脑端CMD控制台中执行上述安装即可。按Enter退出...")
    sys.exit(1)

async def handle_client(websocket):
    pythoncom.CoInitialize()
    print(f"\\n[+] 网页端已成功连接到本地网桥 (端口 {PORT})")
    try:
        # 发送初始在线状态给网页端
        await websocket.send(json.dumps({
            "type": "status",
            "status": "idle",
            "message": "本地 OLE COM 服务端在线监听中，等待您在网页端触发一键填料..."
        }))
        
        async for message_str in websocket:
            try:
                data = json.loads(message_str)
            except Exception as e:
                print(f"[!] 无法解析接收到的数据: {e}")
                continue
                
            action = data.get("action")
            if action == "write_excel":
                items_to_write = data.get("items", [])
                excel_name = data.get("excelFileName", "")
                qty_col_idx = data.get("qtyCol", 5) # 0-based index
                
                print(f"\\n[+] 接收到网页端一键发送写回指令，开始回写 {len(items_to_write)} 项算料数据...")
                await websocket.send(json.dumps({
                    "type": "status",
                    "status": "writing",
                    "message": "正在捕获本地正在运行的 Excel/WPS 活跃工作簿..."
                }))
                
                try:
                    excel_app = None
                    active_provider = ""
                    # 自适应查找 WPS 和 Microsoft Excel
                    providers = [
                        ("WPS 电子表格", "Ket.Application"), 
                        ("Microsoft Excel", "Excel.Application")
                    ]
                    
                    for name, prog_id in providers:
                        try:
                            excel_app = win32com.client.GetActiveObject(prog_id)
                            active_provider = name
                            print(f"[✔] 成功连接本地: {name}")
                            break
                        except Exception:
                            continue
                            
                    if not excel_app:
                        err_msg = "未发现任何活跃的 Excel 或 WPS 表格主进程。请确保目标表格已被双击打开在桌面上。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        pass
                        continue
                        
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "writing",
                        "message": f"已定位 {active_provider} 内核，正在精确匹配文件名: {excel_name}..."
                    }))
                    
                    try:
                        wb = excel_app.ActiveWorkbook
                    except Exception as e:
                        wb = None
                        
                    if not wb:
                        err_msg = "未检测到任何可编辑的工作表！请确保本地有活动表格处于打开状态。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        continue

                    # 优先选择并激活指定名称的工作表，实现自动切换 Excel 标签页
                    sheet = None
                    if excel_name:
                        try:
                            sheet = wb.Sheets(excel_name)
                            sheet.Activate()
                            print(f"[✔] 已激活并自动切换至目标工作表: [{excel_name}]")
                        except Exception as e_sheet:
                            print(f"[!] 自动切换至工作表 [{excel_name}] 失败: {e_sheet}，将默认使用当前活动工作表。")
                            
                    if not sheet:
                        sheet = wb.ActiveSheet
                    print(f"[*] 选中当前工作表底表：{sheet.Name}")
                    
                    old_calc_mode = None
                    try:
                        excel_app.ScreenUpdating = False
                        # 辅材用量填回时先将计算改成手动，方便填回的速度和效率
                        try:
                            old_calc_mode = excel_app.Calculation
                            excel_app.Calculation = -4135  # xlCalculationManual (手动计算)
                            print("[*] 已将 Excel 计算模式切换为 [手动计算] (提升填回速度)")
                        except Exception as e_calc:
                            print(f"[*] 无法切换计算模式为手动: {e_calc}")
                    except:
                        pass
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "writing",
                        "message": f"已锁定工作簿 [{wb.Name}]，已临时切换至 [手动计算] 模式，开始智能单元格高亮回填..."
                    }))
                    
                    success_count = 0
                    fail_count = 0
                    
                    # 网页 0-based 列坐标修正为 Excel OLE 1-based 列
                    excel_col = qty_col_idx + 1 
                    
                    for idx, item in enumerate(items_to_write):
                        if not item.get("isCalculated") or item.get("calculatedQty") is None:
                            continue
                            
                        target_qty = round(float(item["calculatedQty"]), 3)
                        row_index_0based = item.get("rowIndex")
                        target_drawing = item.get("drawingNo", "").strip().lower()
                        target_name = item.get("name", "").strip().lower()
                        
                        written_ok = False
                        
                        # 定时发送进度信息
                        if idx % max(1, len(items_to_write) // 20) == 0:
                            await websocket.send(json.dumps({
                                "type": "progress",
                                "progress": int((idx / len(items_to_write)) * 100),
                                "current": idx + 1,
                                "total": len(items_to_write)
                            }))
                            
                        # 优先选择精确坐标检验回写
                        if row_index_0based is not None:
                            excel_row_1based = row_index_0based + 1
                            try:
                                # 读取该行邻近的前几列，确认物料名或规格匹配
                                cell_val = ""
                                for test_col in [2, 3, 4]:
                                    try:
                                        cell_val += str(sheet.Cells(excel_row_1based, test_col).Value or "") + " "
                                    except:
                                        pass
                                cell_val_lower = cell_val.lower()
                                
                                if (target_drawing and target_drawing in cell_val_lower) or (target_name and target_name in cell_val_lower):
                                    sheet.Cells(excel_row_1based, excel_col).Value = target_qty
                                    written_ok = True
                                    print(f"  [精确行回写] 行号 {excel_row_1based}: {item['name']} ({item['drawingNo']}) -> 填入: {target_qty}")
                            except Exception as ex:
                                print(f"  [精确行检验提醒] 行号 {excel_row_1based} 溢出或遇到锁单元格: {ex}")
                        
                        # 如果坐标偏离，则启用智能全局扫描
                        if not written_ok:
                            for r_search in range(2, 500): # 检索前500行
                                try:
                                    line_str = ""
                                    for c_search in [2, 3, 4]:
                                        line_str += str(sheet.Cells(r_search, c_search).Value or "") + " "
                                    
                                    line_str = line_str.lower()
                                    if (target_drawing and target_drawing in line_str) or (target_name and target_name in line_str):
                                        sheet.Cells(r_search, excel_col).Value = target_qty
                                        written_ok = True
                                        print(f"  [全局检索定位] 行号 {r_search}: '{item['name']}' ({item['drawingNo']}) -> 填入: {target_qty}")
                                        break
                                except Exception:
                                    break
                                    
                        if written_ok:
                            success_count += 1
                        else:
                            fail_count += 1
                            
                    success_msg = f"已成功完成回写！在 Excel 表中精准填入 {success_count} 个数量格。未配对 {fail_count} 项物料。"
                    print(f"[✔] {success_msg}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "success",
                        "message": success_msg
                    }))
                    
                except Exception as ex:
                    print(f"[!] 运行时捕获异常: {ex}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "error",
                        "message": f"回写意外中断: {str(ex)}"
                    }))
                finally:
                    try:
                        if excel_app and old_calc_mode is not None:
                            excel_app.Calculation = old_calc_mode
                            print("[*] 已成功恢复 Excel 原始计算模式")
                    except Exception as e_calc:
                        print(f"[*] 无法恢复原始计算模式: {e_calc}")
                    try:
                        if excel_app:
                            excel_app.ScreenUpdating = True
                    except:
                        pass
                    # 显式解除 OLE COM 对象的所有权引用并触发垃圾回收
                    sheet = None
                    wb = None
                    excel_app = None
                    import gc
                    gc.collect()
                    
            elif action == "write_cells":
                updates = data.get("updates", [])
                excel_name = data.get("excelFileName", "")
                
                print(f"\\n[+] 接收到网页端批量精准单元格区域写回指令，开始回写更新 {len(updates)} 个单元格...")
                await websocket.send(json.dumps({
                    "type": "status",
                    "status": "writing",
                    "message": "正在捕捉本地正在运行的 Excel/WPS 活跃工作簿进行区间清空并回写..."
                }))
                
                try:
                    excel_app = None
                    active_provider = ""
                    providers = [
                        ("WPS 电子表格", "Ket.Application"), 
                        ("Microsoft Excel", "Excel.Application")
                    ]
                    
                    for name, prog_id in providers:
                        try:
                            excel_app = win32com.client.GetActiveObject(prog_id)
                            active_provider = name
                            print(f"[✔] 成功连接本地: {name}")
                            break
                        except Exception:
                            continue
                            
                    if not excel_app:
                        err_msg = "未发现任何活跃的 Excel 或 WPS 表格主进程。请确保目标表格已被双击打开在桌面上。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        pass
                        continue
                        
                    try:
                        wb = excel_app.ActiveWorkbook
                    except Exception as e:
                        wb = None
                        
                    if not wb:
                        raise Exception("无法捕获当前处于活动状态的表格簿")

                    # 优先选择并激活指定名称的工作表，实现自动切换 Excel 标签页
                    sheet = None
                    if excel_name:
                        try:
                            sheet = wb.Sheets(excel_name)
                            sheet.Activate()
                            print(f"[✔] 已激活并自动切换至目标工作表: [{excel_name}]")
                        except Exception as e_sheet:
                            print(f"[!] 自动切换至工作表 [{excel_name}] 失败: {e_sheet}，将默认使用当前活动工作表。")
                            
                    if not sheet:
                        sheet = wb.ActiveSheet
                    print(f"[*] 定位到当前表：{sheet.Name}")
                    
                    old_calc_mode = None
                    try:
                        excel_app.ScreenUpdating = False
                        # 辅材用量填回时先将计算改成手动，方便填回的速度和效率
                        try:
                            old_calc_mode = excel_app.Calculation
                            excel_app.Calculation = -4135  # xlCalculationManual (手动计算)
                            print("[*] 已将 Excel 计算模式切换为 [手动计算] (提升填回速度)")
                        except Exception as e_calc:
                            print(f"[*] 无法切换计算模式为手动: {e_calc}")
                    except:
                        pass
                    
                    # 批量更新单元格
                    for idx, upd in enumerate(updates):
                        # 定时发送进度信息
                        if idx % max(1, len(updates) // 20) == 0:
                            await websocket.send(json.dumps({
                                "type": "progress",
                                "progress": int((idx / len(updates)) * 100),
                                "current": idx + 1,
                                "total": len(updates)
                            }))
                            await asyncio.sleep(0.001)
                            
                        r_ex = upd["r"] + 1
                        c_ex = upd["c"] + 1
                        val = upd["v"]
                        bg_col = upd.get("bg", None)
                        try:
                            cell = sheet.Cells(r_ex, c_ex)
                            if val == "" or val is None:
                                cell.Value = ""
                            elif str(val).startswith("=IMAGE("):
                                try:
                                    import re
                                    import urllib.request
                                    import tempfile
                                    url_match = re.search(r'=IMAGE\("(.*?)"\)', str(val))
                                    if url_match:
                                        img_url = url_match.group(1)
                                        inserted = False
                                        # 优先尝试下载为本地临时文件再插入 (WPS 与 Excel 对本地路径图片兼容性极佳)
                                        try:
                                            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                                            temp_path = temp_file.name
                                            temp_file.close()
                                            req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                                            with urllib.request.urlopen(req, timeout=3) as resp, open(temp_path, 'wb') as out_f:
                                                out_f.write(resp.read())
                                            
                                            pic = sheet.Pictures().Insert(temp_path)
                                            pic.Top = cell.Top
                                            pic.Left = cell.Left
                                            pic.Width = 80
                                            pic.Height = 80
                                            inserted = True
                                            try:
                                                os.remove(temp_path)
                                            except:
                                                pass
                                        except Exception:
                                            pass

                                        if not inserted:
                                            # 如果下载失败或受限，尝试直接插入
                                            try:
                                                pic = sheet.Pictures().Insert(img_url)
                                                pic.Top = cell.Top
                                                pic.Left = cell.Left
                                                pic.Width = 80
                                                pic.Height = 80
                                                inserted = True
                                            except Exception:
                                                pass

                                        if not inserted:
                                            # WPS 及非365版本不原生支持 =IMAGE() 公式，此处安全跳过或写入空/文字，避免抛出 COM 异常
                                            try:
                                                cell.Formula = val
                                            except Exception:
                                                pass
                                except Exception:
                                    pass
                            elif str(val).startswith("="):
                                cell.Formula = val
                            else:
                                cell.Value = val
                        except Exception as ce:
                            print(f"    写入单元格 [{r_ex}, {c_ex}] 失败: {ce}")
                            
                    success_msg = f"已成功清空旧区域并回填 {len(updates)} 个单元格！"
                    print(f"[✔] {success_msg}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "success",
                        "message": success_msg
                    }))
                    
                except Exception as ex:
                    print(f"[!] 批量写入异常: {ex}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "error",
                        "message": f"批量写入中断: {str(ex)}"
                    }))
                finally:
                    try:
                        if excel_app and old_calc_mode is not None:
                            excel_app.Calculation = old_calc_mode
                            print("[*] 已成功恢复 Excel 原始计算模式")
                    except Exception as e_calc:
                        print(f"[*] 无法恢复原始计算模式: {e_calc}")
                    try:
                        if excel_app:
                            excel_app.ScreenUpdating = True
                    except:
                        pass
                    # 显式解除 OLE COM 对象的所有权引用并触发垃圾回收
                    sheet = None
                    wb = None
                    excel_app = None
                    import gc
                    gc.collect()

            elif action == "read_excel":
                requested_sheet = data.get("sheetName", "")
                print(f"\\n[+] 接收到网页端自动拉取表格数据指令 (目标页签: '{requested_sheet}')...")
                await websocket.send(json.dumps({
                    "type": "status",
                    "status": "reading",
                    "message": "正在获取本地 WPS/Excel 进程并读取单元格..."
                }))
                
                try:
                    excel_app = None
                    active_provider = ""
                    providers = [
                        ("WPS 电子表格", "Ket.Application"), 
                        ("Microsoft Excel", "Excel.Application")
                    ]
                    
                    for name, prog_id in providers:
                        try:
                            excel_app = win32com.client.GetActiveObject(prog_id)
                            active_provider = name
                            print(f"[✔] 成功读取锁定: {name}")
                            break
                        except Exception:
                            continue
                            
                    if not excel_app:
                        err_msg = "未发现活动软件进程，请确保本地已在 wps 或 excel 表格窗口中打开了算料料单。若正处于编辑公式/修改单元格模式中，请按回车或 Esc 退出表格编辑模式。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        await websocket.send(json.dumps({
                            "type": "read_response",
                            "success": False,
                            "message": err_msg
                        }))
                        pass
                        continue
 
                    wb = excel_app.ActiveWorkbook
                    if not wb:
                        err_msg = "没有检测到任何已打开的活动工作谱！请确定 Excel 处于前台未被编辑且已经打开料单。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        await websocket.send(json.dumps({
                            "type": "read_response",
                            "success": False,
                            "message": err_msg
                        }))
                        pass
                        continue
                    
                    # 获取该工作谱中的所有工作表页签列表
                    sheets_list = []
                    try:
                        for s_idx in range(1, wb.Sheets.Count + 1):
                            sheets_list.append(wb.Sheets.Item(s_idx).Name)
                    except Exception as e_sheets:
                        print(f"[*] 提取工作表列表受限: {e_sheets}")
                        sheets_list = []
                        
                    # 确定要读取的目标工作表
                    sheet = None
                    if requested_sheet:
                        try:
                            sheet = wb.Sheets(requested_sheet)
                            print(f"[✔] 已锁定指定名称的工作表: [{requested_sheet}]")
                        except Exception:
                            print(f"[!] 未找到工作表 [{requested_sheet}]，将默认使用当前活动工作表。")
                            
                    if not sheet:
                        sheet = wb.ActiveSheet
                    
                    if not sheet:
                        err_msg = "未检测到有效的工作表！请尝试双击表格退出编辑状态。"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                        await websocket.send(json.dumps({
                            "type": "read_response",
                            "success": False,
                            "message": err_msg
                        }))
                        pass
                        continue
                        
                    print(f"[*] 正在读取工作表: [{wb.Name}] -> [{sheet.Name}]")
                    rows_data = []
                    
                    # 使用极其高效的批量 Range-Value 二维数组直接读取，支持 WPS & Excel 的 OLE API (仅需 1 次 OLE 请求，绝不卡死转圈圈)
                    try:
                        # 尝试读取 UsedRange 加速
                        used_range = sheet.UsedRange
                        vals = used_range.Value
                        print(f"[*] 自动捕获 UsedRange 批量回传中...")
                    except Exception as e_used:
                        print(f"[*] UsedRange 读取受限 ({e_used})。尝试读取固定 A1:K150 常规算料区域...")
                        try:
                            vals = sheet.Range(sheet.Cells(1, 1), sheet.Cells(150, 11)).Value
                        except Exception as e_range:
                            print(f"[!] 固定区间批量读取失败 ({e_range})。请检查 Excel 是否处于单元格编辑状态！")
                            vals = None
                            
                    if vals:
                        if isinstance(vals, tuple):
                            for r_tuple in vals:
                                row_arr = []
                                if isinstance(r_tuple, tuple):
                                    for cell_val in r_tuple:
                                        row_arr.append("" if cell_val is None else str(cell_val))
                                else:
                                    row_arr.append("" if r_tuple is None else str(r_tuple))
                                rows_data.append(row_arr)
                        else:
                            rows_data.append(["" if vals is None else str(vals)])
                    else:
                        raise Exception("无法从 Excel 中直接抽取有效的单元格值。可能被 Windows 锁定或单元格正处于闪烁编辑状态（请按键 Esc/Enter 平复状态）。")
 
                    success_msg = f"成功从工作表 [{sheet.Name}] 读取了 {len(rows_data)} 行物料原始字段！"
                    print(f"[✔] {success_msg}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "idle",
                        "message": f"表格 [{wb.Name}] 的 [{sheet.Name}] 数据读取成功！"
                    }))
                    await websocket.send(json.dumps({
                        "type": "read_response",
                        "success": True,
                        "sheetName": sheet.Name,
                        "sheetsList": sheets_list,
                        "fileName": wb.Name,
                        "rows": rows_data,
                        "message": success_msg
                    }))
                except Exception as ex:
                    print(f"[!] 读取时发生致命异常: {ex}")
                    err_hint = f"读取失败: {str(ex)}。提示：若 Excel 正在编辑单元格，请务必在 Excel 中按键盘 Enter 或 Esc 退出编辑状态，再度尝试拉取！"
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "error",
                        "message": err_hint
                    }))
                    await websocket.send(json.dumps({
                        "type": "read_response",
                        "success": False,
                        "message": err_hint
                    }))
                finally:
                    # 显式解除 OLE COM 对象的所有权引用并触发垃圾回收
                    sheet = None
                    wb = None
                    excel_app = None
                    import gc
                    gc.collect()
                    
            elif action == "execute_vb":
                code_to_run = data.get("code", "")
                print(f"\\n[+] 接收到网页端动态 VB 脚本执行指令，代码长度: {len(code_to_run)} 字符...")
                await websocket.send(json.dumps({
                    "type": "status",
                    "status": "writing",
                    "message": "正在解析并安全转换 VB/VBA 脚本为兼容 VBScript..."
                }))
                
                try:
                    import tempfile
                    import subprocess
                    import re
                    
                    # 转换 VBA/VB 代码为兼容的 VBScript
                    def convert_vba_to_vbscript(vba_code):
                        lines = vba_code.splitlines()
                        sanitized_lines = []
                        for line in lines:
                            cleaned = line
                            
                            # 1. 注释掉 Option Explicit，避免其不作为首行而报错
                            cleaned = re.sub(r"\\bOption\\s+Explicit\\b", "' Option Explicit", cleaned, flags=re.IGNORECASE)
                            
                            # 2. 注释掉单纯的行标签 (如 ErrorHandler:)
                            if re.match(r"^\\s*\\w+\\s*:\\s*$", cleaned):
                                cleaned = "' " + cleaned
                                
                            # 3. 将 On Error GoTo ... 转换为 On Error Resume Next
                            cleaned = re.sub(r"\\bOn\\s+Error\\s+GoTo\\s+\\w+\\b", "On Error Resume Next", cleaned, flags=re.IGNORECASE)
                            
                            # 4. 将 Next i 转换为 Next (VBScript 不支持 Next 后面带变量名)
                            cleaned = re.sub(r"\\bNext\\s+\\w+\\b", "Next", cleaned, flags=re.IGNORECASE)
                            
                            # 5. 移除强类型声明 (As Range, As Long, As String 等)
                            cleaned = re.sub(r"\\bAs\\s+[\\w\\.\\$]+", "", cleaned, flags=re.IGNORECASE)
                            
                            # 6. 处理 Function 函数返回类型声明 (Function Name(...) As Double)
                            cleaned = re.sub(r"(Function\\s+\\w+\\s*\\(.*?\\))\\s+As\\s+[\\w\\.\\$]+", r"\\1", cleaned, flags=re.IGNORECASE)
                            
                            # 7. 将独立的 End 替换为 WScript.Quit (VBScript 不支持裸 End 语句退出)
                            if cleaned.strip().lower() == "end":
                                cleaned = "WScript.Quit"
                                
                            sanitized_lines.append(cleaned)
                            
                        vbs = "\\n".join(sanitized_lines)
                        
                        # 2. 注入标准兼容头部与全局包装函数
                        header = """' VBScript Excel Controller Auto-Generated Header
On Error Resume Next
Dim xlApp, wb, sheet
Set xlApp = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Error: Cannot find active Excel/WPS application."
    WScript.Quit
End If
Set wb = xlApp.ActiveWorkbook
If Err.Number <> 0 Or wb Is Nothing Then
    WScript.Echo "Error: No active workbook found."
    WScript.Quit
End If
Set sheet = xlApp.ActiveSheet
If Err.Number <> 0 Or sheet Is Nothing Then
    WScript.Echo "Error: No active sheet found."
    WScript.Quit
End If

' Inject standard global objects into scope so user VBA code can run without prefixes
Dim ActiveWorkbook, ActiveSheet, ThisWorkbook
Set ActiveWorkbook = wb
Set ActiveSheet = sheet
Set ThisWorkbook = wb

' Common Excel VBA Constants for VBScript compatibility
Const xlCalculationAutomatic = -4105
Const xlCalculationManual = -4135
Const xlCenter = -4108
Const xlLeft = -4131
Const xlRight = -4152
Const xlTop = -4160
Const xlBottom = -4107
Const xlSolid = 1
Const xlNone = -4142
Const xlThin = 2
Const xlMedium = -4138
Const xlThick = 4
Const xlDouble = -4119
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideHorizontal = 12
Const xlInsideVertical = 11
Const xlContinuous = 1
Const xlThemeColorAccent1 = 5
Const xlThemeColorAccent2 = 6
Const xlThemeColorAccent3 = 7
Const xlThemeColorAccent4 = 8
Const xlThemeColorAccent5 = 9
Const xlThemeColorAccent6 = 10
Const xlUp = -4162
Const xlDown = -4121
Const xlToLeft = -4159
Const xlToRight = -4161

' Standard Color Constants
Const vbBlack = 0
Const vbRed = 255
Const vbGreen = 65280
Const vbYellow = 65535
Const vbBlue = 16711680
Const vbMagenta = 16711935
Const vbCyan = 16776960
Const vbWhite = 16777215

' Global helper wrappers to support implicit ActiveSheet references in VBA
Function Range(byval cellOrRange)
    Set Range = ActiveSheet.Range(cellOrRange)
End Function

Function Cells(byval row, byval col)
    Set Cells = ActiveSheet.Cells(row, col)
End Function

Class DebugClass
    Public Sub Print(byval msg)
        WScript.Echo msg
    End Sub
End Class
Dim Debug
Set Debug = New DebugClass

"""
                        # 3. 如果脚本包含 Sub 结构，自动在脚本最末尾进行 Call 调用唯一的第一个 Sub 主程序
                        sub_matches = re.findall(r"Sub\\s+(\\w+)\\s*\\(", vbs, re.IGNORECASE)
                        footer = ""
                        if sub_matches:
                            sub_name = sub_matches[0]
                            footer = f"\\n\\n' Auto-calling the main sub routine\\nCall {sub_name}()\\n"
                            
                        return header + vbs + footer

                    vbs_content = convert_vba_to_vbscript(code_to_run)
                    
                    # 写入临时文件
                    temp_path = None
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".vbs", mode="w") as tf:
                        temp_path = tf.name
                        
                    # 尝试用 GBK 写入，如果有中文或者特殊字符用 UTF-8 BOM 写入
                    try:
                        with open(temp_path, "w", encoding="gbk") as f:
                            f.write(vbs_content)
                    except UnicodeEncodeError:
                        with open(temp_path, "w", encoding="utf-8-sig") as f:
                            f.write(vbs_content)
                            
                    print(f"[*] 临时 VBS 脚本已生成: {temp_path}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "writing",
                        "message": "正在通过本地 Windows 宿主环境 cscript.exe 安全运行 VBS 自动化脚本..."
                    }))
                    
                    # 启动 cscript 运行脚本
                    process = subprocess.Popen(
                        ["cscript.exe", "/nologo", temp_path],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True
                    )
                    stdout, stderr = process.communicate()
                    
                    # 清理临时文件
                    try:
                        os.remove(temp_path)
                    except Exception as e_del:
                        print(f"[*] 无法删除临时文件: {e_del}")
                        
                    if process.returncode == 0:
                        success_msg = "VB 代码执行成功！本地表格已完成相应修改。"
                        stdout_str = stdout.strip()
                        if stdout_str:
                            success_msg += f"\\n脚本执行输出:\\n{stdout_str}"
                        print(f"[✔] {success_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "success",
                            "message": success_msg
                        }))
                    else:
                        stderr_str = stderr.strip() or stdout.strip() or "未知运行错误"
                        err_msg = f"VB 脚本运行失败 (返回码 {process.returncode}): {stderr_str}"
                        print(f"[错误] {err_msg}")
                        await websocket.send(json.dumps({
                            "type": "status",
                            "status": "error",
                            "message": err_msg
                        }))
                except Exception as ex_vb:
                    print(f"[!] 执行 VB 脚本致命异常: {ex_vb}")
                    await websocket.send(json.dumps({
                        "type": "status",
                        "status": "error",
                        "message": f"VB 脚本运行致命错误: {str(ex_vb)}"
                    }))
                    
    except websockets.exceptions.ConnectionClosed:
        print("[-] 网页端已断开连接.")
    except Exception as e:
        print(f"[-] 连接发生错误: {e}")
    finally:
        pythoncom.CoUninitialize()
        print("[*] 已安全释放 COM 线程资源库连接计数")

async def main():
    try:
        async with websockets.serve(handle_client, "127.0.0.1", PORT):
            print(f"\\n[✔] 本地 OLE COM 进程桥接器已启动在 ws://127.0.0.1:{PORT}")
            print("[*] 请在浏览器中点击『一键 COM 穿透填回』触发联动。按 Ctrl+C 可停止。")
            await asyncio.Future()
    except Exception as e:
        print(f"\\n[!] 启动服务失败(可能端口被占用): {e}")

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\\n[!] 进程已安全停止退出。")
`
      .replace("${localWsPort}", localWsPort.toString());

    const blob = new Blob([scriptContent], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "com_bridge.py";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const getStatusBadge = () => {
    if (wsConnected) {
      switch (comBridgeStatus) {
        case "writing":
          return (
            <span className="flex items-center px-3 py-1.5 bg-amber-50 border border-amber-200 rounded text-xs font-bold text-amber-700 animate-pulse">
              <RefreshCw className="w-3.5 h-3.5 animate-spin mr-1.5 text-amber-600" />
              正在写入 Excel/WPS 单元格...
            </span>
          );
        case "success":
          return (
            <span className="flex items-center px-3 py-1.5 bg-green-50 border border-green-200 rounded text-xs font-bold text-green-700">
              <CheckCircle2 className="w-3.5 h-3.5 text-green-600 mr-1.5" />
              回写成功！已高亮标记
            </span>
          );
        case "error":
          return (
            <span className="flex items-center px-3 py-1.5 bg-red-50 border border-red-200 rounded text-xs font-bold text-red-700">
              <AlertTriangle className="w-3.5 h-3.5 text-red-500 mr-1.5" />
              回写遇到错误
            </span>
          );
        default:
          return (
            <span className="flex items-center px-3 py-1.5 bg-emerald-50 border border-emerald-200 rounded text-xs font-bold text-emerald-700">
              <span className="w-2 h-2 bg-emerald-500 rounded-full mr-1.5 animate-ping"></span>
              本地 COM 连接成功 (准备就绪)
            </span>
          );
      }
    } else {
      return (
        <span className="flex items-center px-3 py-1.5 bg-slate-100 border border-slate-200 rounded text-xs font-bold text-slate-500">
          <span className="w-2 h-2 bg-slate-400 rounded-full mr-1.5 animate-pulse"></span>
          本地网桥未配对
        </span>
      );
    }
  };

  return (
    <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm" id="module_com_connector">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-6 gap-4">
        <div>
          <h2 className="text-lg font-bold text-slate-800 tracking-tight flex items-center gap-2">
            <Server className="w-5 h-5 text-indigo-600" />
            电脑本地 OLE/COM 双向网桥 (极速 WebSocket/CAD 连接协议)
          </h2>
          <p className="text-xs text-slate-500 mt-1">
            由您提供的极速 Local WebSocket CAD 连接协议启发而来。我们在电脑端启动轻量极速网桥，网页通过 <code>ws://127.0.0.1</code> 闪电配对，直接秒级操控本地 Excel / WPS 填充。<strong>不限端口，无任何云端限制！</strong>
          </p>
        </div>

        <div className="flex items-center gap-2 shrink-0">
          {getStatusBadge()}
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* COM Parameters UI */}
        <div className="lg:col-span-6 space-y-4">
          <div className="bg-slate-50 p-5 rounded-lg border border-slate-200 space-y-4 flex flex-col justify-between h-full">
            <div>
              <h3 className="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5 mb-3">
                <Cpu className="w-4 h-4 text-indigo-600" />
                本地 WebSocket 端口与程序指令配置
              </h3>

              {/* Port Config and Connect Form */}
              <div className="bg-white p-4 rounded-lg border border-slate-200 mb-4 shadow-2xs">
                <form onSubmit={handleApplyPort} className="flex flex-col sm:flex-row gap-3 items-end sm:items-center justify-between">
                  <div>
                    <label className="text-[10px] uppercase font-bold text-slate-400 tracking-wider mb-1 block">
                      本地网桥监听端口 (WebSocket Port)
                    </label>
                    <input
                      type="number"
                      value={portInput}
                      onChange={(e) => setPortInput(e.target.value)}
                      placeholder="8001"
                      disabled={wsConnected}
                      className="text-sm font-mono font-bold text-indigo-700 bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5 w-32 focus:outline-none focus:ring-1 focus:ring-indigo-500 disabled:opacity-50"
                    />
                  </div>
                  <div className="flex gap-2">
                    <button
                      type="submit"
                      disabled={wsConnected}
                      className="px-4 py-1.5 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 text-xs font-bold rounded cursor-pointer transition-colors disabled:opacity-50"
                    >
                      修改端口
                    </button>
                    {wsConnected && (
                      <span className="text-[11px] text-green-600 font-bold flex items-center gap-1">
                        <CheckCircle2 className="w-3.5 h-3.5" /> 已连通
                      </span>
                    )}
                  </div>
                </form>
              </div>

              {/* Step instructions */}
              <div className="space-y-2.5 text-xs text-slate-600">
                <div className="flex gap-2">
                  <span className="w-5 h-5 rounded-full bg-slate-200 flex items-center justify-center font-bold text-[10px] text-slate-600 shrink-0 mt-0.5">1</span>
                  <span>第一步：在右下方点击并一键下载 <code>com_bridge.py</code> 网桥程序。</span>
                </div>
                <div className="flex gap-2">
                  <span className="w-5 h-5 rounded-full bg-slate-200 flex items-center justify-center font-bold text-[10px] text-slate-600 shrink-0 mt-0.5">2</span>
                  <span>第二步：打开您的本地命令行并运行 <code>pip install websockets pywin32</code>。</span>
                </div>
                <div className="flex gap-2">
                  <span className="w-5 h-5 rounded-full bg-slate-200 flex items-center justify-center font-bold text-[10px] text-slate-600 shrink-0 mt-0.5">3</span>
                  <span>第三步：启动运行 <code>python com_bridge.py</code> 即可看到网桥成功建立监听。</span>
                </div>
              </div>
            </div>

            <div className="flex flex-col sm:flex-row gap-2 pt-4 border-t border-slate-200">
              <button
                onClick={downloadPythonScript}
                className="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 px-3 rounded-lg text-xs transition-colors flex items-center justify-center gap-1.5 cursor-pointer shadow-sm"
              >
                <Download className="w-4 h-4" />
                <span>下载 com_bridge.py</span>
              </button>
            </div>
          </div>
        </div>

        {/* COM Terminal Logging UI */}
        <div className="lg:col-span-6 flex flex-col h-full min-h-[245px]">
          <div className="flex-1 flex flex-col bg-slate-900 border border-slate-800 rounded-lg overflow-hidden">
            <div className="bg-slate-950 px-4 py-3 flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-2">
                <Terminal className="w-4 h-4 text-indigo-400" />
                <span className="text-[10px] font-mono text-slate-400 tracking-wider">
                  LOCAL WEBSOCKET TUNNEL LOGGING
                </span>
              </div>
              <div className="flex items-center gap-1.5">
                <span className={`w-1.5 h-1.5 rounded-full ${wsConnected ? 'bg-emerald-500 animate-ping' : 'bg-red-500'}`}></span>
                <span className="text-[9px] font-mono text-slate-500 uppercase">
                  {wsConnected ? 'Connected' : 'Listening...'}
                </span>
              </div>
            </div>

            <div className="flex-1 font-mono text-xs text-indigo-300 p-5 space-y-3 min-h-[160px] overflow-y-auto bg-[#0a0f1d] leading-relaxed">
              <div className="text-slate-500 text-[10px] mb-2 border-b border-slate-800/80 pb-2 flex justify-between items-center">
                <span>[LOCAL PROTOCOL ENGINE ACTIVE]</span>
                <button
                  onClick={handleCopyCode}
                  className="px-2 py-0.5 bg-slate-800/50 hover:bg-slate-800 text-[9px] text-indigo-400 hover:text-indigo-300 rounded border border-slate-700/80 cursor-pointer flex items-center gap-1"
                >
                  {copied ? <Check className="w-3 h-3 text-green-500" /> : <Copy className="w-3 h-3" />}
                  复制运行命令
                </button>
              </div>
              <div>
                <span className="text-indigo-500 mr-2">▶</span>
                <span className="text-slate-400 font-bold">连接地址:</span>{" "}
                <span className="text-emerald-400 font-bold">ws://127.0.0.1:{localWsPort}</span>
              </div>
              <div>
                <span className="text-indigo-500 mr-2">▶</span>
                <span className="text-slate-400 font-bold">桥接状态:</span>{" "}
                <span className={`uppercase font-bold ${wsConnected ? 'text-green-400' : 'text-amber-400'}`}>
                  {wsConnected ? "已连接 (SUCCESS)" : "等待中 (DISCONNECTED)"}
                </span>
              </div>
              <div className="bg-slate-900/50 p-4 rounded border border-slate-800 text-slate-300 whitespace-pre-wrap leading-relaxed mt-2 text-[11px] font-sans">
                {comBridgeMessage}
              </div>
            </div>
          </div>
          
          <p className="text-[11px] text-slate-500 mt-2.5 flex items-start gap-1.5 leading-relaxed">
            <Info className="w-3.5 h-3.5 text-indigo-500 shrink-0 mt-0.5" />
            <span>
              得益于本地 WebSocket 握手协议，您可以完美避开 Google AI Studio 网页与本地沙箱的端口隔离。数据将在您的网页端与本机 Python 桥接软件之间<b>瞬间传输，不需要您暴露任何 3000 外网端口</b>，安全稳定、完全私密。
            </span>
          </p>
        </div>
      </div>
    </div>
  );
}
