# ExcelEditTool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Excel summary tool for finance users with sheet preview, configurable grouping and summing, styled result preview, formatted clipboard copy, and export to `原Excel文件名_汇总.xlsx`.

**Architecture:** A React/Vite frontend talks to a local FastAPI backend. The backend owns Excel parsing, style extraction, grouping, and `.xlsx` generation with openpyxl. The frontend owns the Office-like interaction surface, table preview, configuration controls, clipboard HTML, and download flow.

**Tech Stack:** React, TypeScript, Vite, Fluent UI React v9, FastAPI, pydantic, openpyxl, pytest, Vitest.

---

## File Structure

- `backend/app/main.py`: FastAPI app and HTTP endpoints.
- `backend/app/models.py`: Request and response models.
- `backend/app/excel_processor.py`: Workbook loading, preview extraction, grouping, style mapping, export.
- `backend/tests/test_excel_processor.py`: Tests for grouping, warnings, style propagation, and export naming.
- `frontend/src/App.tsx`: Main UI composition and workflow state.
- `frontend/src/api.ts`: Backend API client.
- `frontend/src/clipboard.ts`: HTML and plain text clipboard generation.
- `frontend/src/types.ts`: Shared frontend types.
- `frontend/src/styles.css`: Office-inspired application styles and table rendering.
- `frontend/src/clipboard.test.ts`: Clipboard HTML tests.
- Root project files for Python and Node setup.

## Tasks

- [ ] Create project scaffolding and dependency manifests.
- [ ] Write failing backend tests for workbook preview, grouping, warnings, and export name.
- [ ] Implement backend Excel processor to pass tests.
- [ ] Add FastAPI endpoints for upload, preview, summarize, and export.
- [ ] Write failing frontend tests for clipboard HTML/plain text.
- [ ] Implement clipboard serialization.
- [ ] Build Fluent UI frontend workflow.
- [ ] Run backend and frontend verification.
- [ ] Start local services and provide URLs.
