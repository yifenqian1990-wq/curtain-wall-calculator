# 幕墙单元辅材自动计算器

一键导入幕墙加工细目 Excel，智能识别型材、面板和钢件，按计算公式自动计算辅材与紧固件含量，支持结果一键填回原表与导出。

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![React](https://img.shields.io/badge/React-19-61dafb?logo=react&logoColor=white)](https://react.dev/)
[![Vite](https://img.shields.io/badge/Vite-6-646cff?logo=vite&logoColor=white)](https://vitejs.dev/)
[![TypeScript](https://img.shields.io/badge/TypeScript-5-3178c6?logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![Deploy](https://github.com/yifenqian1990-wq/curtain-wall-calculator/actions/workflows/deploy.yml/badge.svg)](https://github.com/yifenqian1990-wq/curtain-wall-calculator/actions)
[![Live](https://img.shields.io/website?url=https%3A%2F%2Fyifenqian1990-wq.github.io%2Fcurtain-wall-calculator%2F&label=online)](https://yifenqian1990-wq.github.io/curtain-wall-calculator/)

## 🚀 在线体验

**https://yifenqian1990-wq.github.io/curtain-wall-calculator/**

打开即用，导入与计算全在浏览器本地完成，不上传任何业务数据。

## ✨ 功能特性

- **Excel 导入**：读取幕墙单元加工材料细目表（.xlsx），智能识别型材、面板、钢件行
- **辅材自动计算**：按内置公式计算辅材与紧固件含量，生成汇总与台账
- **一键填回**：通过本地桥接器将计算结果写回你电脑上打开的 Excel / WPS（见下方说明）
- **导出**：计算结果导出为 Excel
- **纯本地运行**：所有数据只在浏览器本地处理，保护业务数据隐私

## 💻 本地运行

环境要求：Node.js ≥ 20

```bash
npm install
npm run dev      # http://localhost:3000
```

生产构建：

```bash
npm run build    # 产物在 dist/，可直接用任意静态服务器托管
```

## 🔌 一键填回（本地桥接器）

"一键填回"需要配合你电脑上运行的桥接程序，将网页计算结果写入本地 Excel / WPS：

1. 电脑安装 Python 3.8+，执行 `pip install websockets pywin32`
2. 打开目标 Excel / WPS 文件
3. 运行 `python com_bridge.py`（仓库根目录）
4. 在网页的填回面板中连接 `ws://127.0.0.1:端口`，即可推送填回

该功能为纯本地通信，不经过任何云端服务器。

## 🛠️ 技术栈

React 19 + Vite 6 + TypeScript + xlsx

## 📦 部署

本仓库通过 GitHub Pages 自动发布：推送到 `main` 分支后，GitHub Actions 自动构建并上线，全程无需人工干预。

## ❓ 常见问题

**Q: 导入的 Excel 没有识别出数据？**
A: 请确认表格为 `.xlsx` 格式，且包含"单元加工材料细目表"标准模板的列结构；可先用仓库内的标准模板测试。

**Q: 一键填回连不上本地桥接器？**
A: 确认 `com_bridge.py` 正在运行、Excel / WPS 已打开目标文件，且网页填回面板中的端口与桥接器监听端口一致。

**Q: 计算结果和手工算的有差异？**
A: 请检查识别到的型材/面板/钢件分类是否正确，分类错误会导致公式套用错误；可在预览中手动调整分类后重新计算。

## 📄 许可证

本项目采用 [MIT](LICENSE) 开源许可证。
