# ExcelEditTool 前端

这是本地 Excel 表格处理工具的 React 前端。界面使用 Fluent UI React v9，视觉风格参考新版 Microsoft 365 Excel。

## 开发启动

在项目根目录先启动后端：

```powershell
python -m uvicorn backend.app.main:app --host 127.0.0.1 --port 8000
```

再进入 `frontend` 目录启动前端：

```powershell
npm install
npm run dev -- --host 127.0.0.1 --port 5173
```

打开 `http://127.0.0.1:5173/` 使用。

## 验证

```powershell
npm test
npm run lint
npm run build
```
