# ExcelEditTool

本地 Excel 表格处理工具，面向财务人员的日常汇总场景。

支持上传 `.xlsx` 文件，切换 sheet 预览数据，双击单元格编辑原表，按自定义分组列和求和列生成汇总，继续编辑汇总结果，带格式复制到 Excel，以及导出汇总文件。

导出文件名规则：`{原文件名}_汇总.xlsx`

## 技术栈

- 前端：React + TypeScript + Vite + Fluent UI React v9
- 后端：Python + FastAPI + openpyxl
- 桌面客户端：pywebview + PyInstaller
- 测试：pytest（后端）、Vitest（前端）

## 开发环境启动

**后端：**

```powershell
python -m uvicorn backend.app.main:app --host 127.0.0.1 --port 8000
```

**前端：**

```powershell
cd frontend
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

浏览器访问 `http://127.0.0.1:5173/`

## 验证

```powershell
# 后端测试
python -m pytest backend/tests -q

# 前端测试 & 检查
cd frontend
npm test
npm run lint
npm run build
```

## 打包 Windows 客户端

```powershell
.\scripts\package_client.ps1
```

输出：
- `dist\ExcelEditTool\ExcelEditTool.exe`
- `release\ExcelEditTool.zip`
