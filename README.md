# ExcelEditTool

ExcelEditTool 是一个本地 Excel 汇总处理工具。当前版本只保留 Tauri 桌面客户端，Excel 读取、预览、筛选、汇总和导出都在本机完成，不再包含 Python 后端、pywebview 或 PyInstaller 打包链。

## 功能

- 选择本地 `.xlsx` 文件并预览原始数据。
- 支持多个 sheet 切换，每个 sheet 可独立生成汇总结果。
- 支持自定义表头行、分组列、求和列。
- 原始数据支持表头下拉多选过滤，汇总时使用过滤后的数据。
- 原始数据和汇总结果支持双击单元格编辑。
- 汇总结果支持拖拽调整列顺序，复制和导出会按调整后的顺序输出。
- 支持带格式复制汇总结果到 Excel。
- 导出文件命名规则：`原Excel文件名_汇总.xlsx`。

## 下载

Windows 安装包请到 GitHub Releases 下载：

- [Latest Release](https://github.com/XingAur/ExcelEditTool/releases/latest)
- 安装包文件名：`ExcelEditTool_Tauri_Setup.exe`

当前安装包配置：

- 默认安装目录：`D:\Programs\ExcelEditTool`
- 安装完成后不自动启动程序
- 不随包携带 Python 或固定版 WebView2 Runtime
- 不在安装时下载或更新 WebView2，依赖系统已有 Microsoft Edge WebView2 Runtime

## 技术栈

- React + TypeScript + Vite
- Fluent UI React v9
- ExcelJS
- Tauri 2
- Vitest + ESLint

## 目录

```text
frontend/                 React 前端与 Tauri 客户端
frontend/src/             Excel 处理和界面代码
frontend/src-tauri/       Tauri 配置、Rust 入口和 NSIS 安装器模板
scripts/package_tauri_client.ps1
```

## 开发

```powershell
cd frontend
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

启动 Tauri 开发客户端：

```powershell
cd frontend
npm run tauri
```

## 验证

```powershell
cd frontend
npm test
npm run lint
npm run build
```

## 打包

在项目根目录运行：

```powershell
.\scripts\package_tauri_client.ps1
```

输出文件：

```text
release\ExcelEditTool_Tauri_Setup.exe
```
