# ExcelEditTool

ExcelEditTool 是一个面向财务人员的本地 Excel 汇总处理工具，重点解决“同型号、同单价、同单位等维度合并数量和金额”这类日常表格整理工作。

## 主要功能

- 选择本地 `.xlsx` 文件并预览原始数据。
- 支持多个 sheet 切换，每个 sheet 可独立生成汇总结果。
- 支持自定义表头行、分组列、求和列。
- 原始数据支持表头下拉多选过滤，汇总时使用过滤后的数据。
- 原始数据和汇总结果都支持双击单元格编辑，导出使用编辑后的数据。
- 汇总结果支持拖拽调整列顺序，复制和导出都会按调整后的顺序输出。
- 原始数据/汇总结果、分组列/求和列区域支持拖拽调整显示比例。
- 支持带格式复制汇总结果到 Excel。
- 导出文件名规则：`原Excel文件名_汇总.xlsx`。

## 技术栈

- 前端：React + TypeScript + Vite + Fluent UI React v9
- 后端：Python + FastAPI + openpyxl
- 桌面客户端：pywebview + PyInstaller
- 测试：Vitest、pytest
- Windows 安装包：IExpress + WebView2 Evergreen Bootstrapper

## 目录结构

```text
backend/       FastAPI 接口、Excel 预览和导出逻辑
desktop/       pywebview 桌面入口
frontend/      React 前端界面
packaging/     图标、客户端 README、安装脚本
scripts/       打包脚本
```

## 开发环境

安装 Python 依赖：

```powershell
python -m pip install -r requirements.txt
```

安装前端依赖：

```powershell
cd frontend
npm install
```

启动后端：

```powershell
python -m uvicorn backend.app.main:app --host 127.0.0.1 --port 8000
```

启动前端：

```powershell
cd frontend
npm run dev -- --host 127.0.0.1 --port 5173
```

浏览器访问 `http://127.0.0.1:5173/`。

## 验证

```powershell
cd frontend
npm test
npm run lint
npm run build

cd ..
python -m pytest backend\tests desktop\tests -q
```

## 打包

生成免安装 zip 客户端：

```powershell
.\scripts\package_client.ps1
```

输出：

- `dist\ExcelEditTool\ExcelEditTool.exe`
- `release\ExcelEditTool.zip`

生成 Windows 安装包：

```powershell
.\scripts\package_installer.ps1
```

输出：

- `release\ExcelEditTool_Setup.exe`

`package_installer.ps1` 会先调用 `package_client.ps1` 生成客户端，再下载 Microsoft Edge WebView2 Evergreen Bootstrapper，并打入安装包。

安装包默认安装到 `D:\ExcelEditTool`；如果当前电脑没有 D 盘，则回退到当前用户的本地程序目录。

## WebView2 说明

pywebview 在 Windows 上需要 Microsoft Edge WebView2 Runtime。安装包版本会在安装时检测系统是否已有 WebView2 Runtime，缺失时通过 Microsoft 官方 Bootstrapper 静默安装。

当前方案使用 Evergreen Bootstrapper，优点是安装包体积小；首次安装缺少 Runtime 时需要联网。如果要做完全离线分发，可以改为随包附带 Evergreen Standalone Installer，安装包体积会明显变大。

## 提交前建议

提交代码前建议至少执行：

```powershell
cd frontend
npm test
npm run lint

cd ..
python -m pytest backend\tests desktop\tests -q
```
