"use client";

import { ChangeEvent, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  cleanAllSheets,
  columnName,
  Row,
  sheetRows,
  transformBroadcast,
  transformExternalSales,
  transformMomoPchome,
  transformOms,
  TransformResult
} from "@/lib/etl";

type LogKind = "" | "ok" | "warn";

type LogState = {
  message: string;
  kind: LogKind;
};

const actions = [
  {
    id: "momo-pchome",
    title: "MOMO / PCHOME 銷貨單",
    desc: "篩選客戶並合併訂單品號"
  },
  {
    id: "broadcast",
    title: "輔翼轉 ERP 轉播單",
    desc: "依欄名轉成 ERP 轉播單格式"
  },
  {
    id: "external-sales",
    title: "Shopee / Yahoo 銷貨單",
    desc: "外站特殊格式與序號整理"
  },
  {
    id: "oms",
    title: "銷貨格式轉 OMS",
    desc: "轉出 OMS 匯入欄位格式"
  },
  {
    id: "clean-all",
    title: "清理所有工作表",
    desc: "刪除前 6 列、過濾列與欄位"
  }
] as const;

type ActionId = (typeof actions)[number]["id"];

export default function Home() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [selectedSheetName, setSelectedSheetName] = useState("");
  const [originalFileName, setOriginalFileName] = useState("");
  const [fileMeta, setFileMeta] = useState("尚未選擇檔案");
  const [log, setLog] = useState<LogState>({
    message: "請先選擇 Excel 檔案。",
    kind: ""
  });

  const sheetNames = workbook?.SheetNames ?? [];
  const ready = Boolean(workbook && selectedSheetName);

  const previewRows = useMemo(() => {
    if (!workbook || !selectedSheetName) return [];
    const sheet = workbook.Sheets[selectedSheetName];
    if (!sheet) return [];
    return sheetRows(sheet);
  }, [workbook, selectedSheetName]);

  const previewInfo = useMemo(() => {
    const maxRows = Math.min(previewRows.length, 30);
    const maxCols = Math.min(
      Math.max(...previewRows.slice(0, maxRows).map((row) => row.length), 0),
      20
    );

    return {
      rows: previewRows.slice(0, maxRows),
      maxCols,
      note: selectedSheetName && previewRows.length ? `${selectedSheetName}，共 ${previewRows.length} 列` : "最多顯示前 30 列、20 欄"
    };
  }, [previewRows, selectedSheetName]);

  function updateLog(message: string, kind: LogKind = "") {
    setLog({ message, kind });
  }

  function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;

    setOriginalFileName(file.name);
    setWorkbook(null);
    setSelectedSheetName("");
    updateLog("讀取 Excel 檔案中…");
    setFileMeta(`${file.name} (${Math.round(file.size / 1024).toLocaleString()} KB)`);

    const reader = new FileReader();
    reader.onload = (loadEvent) => {
      try {
        const result = loadEvent.target?.result;
        if (!(result instanceof ArrayBuffer)) {
          throw new Error("無法讀取檔案內容。");
        }

        const data = new Uint8Array(result);
        const nextWorkbook = XLSX.read(data, { type: "array", cellDates: true });
        if (!nextWorkbook.SheetNames.length) {
          setWorkbook(null);
          setSelectedSheetName("");
          updateLog("檔案內沒有可用工作表。", "warn");
          return;
        }

        setWorkbook(nextWorkbook);
        setSelectedSheetName(nextWorkbook.SheetNames[0]);
        updateLog(`已載入 ${file.name}\n請選擇工作表與要執行的 ETL 程式。`, "ok");
      } catch (err) {
        console.error(err);
        setWorkbook(null);
        setSelectedSheetName("");
        updateLog(`讀取檔案失敗：${errorMessage(err)}`, "warn");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function getSelectedSheet() {
    if (!workbook) {
      updateLog("尚未載入檔案。", "warn");
      return null;
    }

    const sheet = workbook.Sheets[selectedSheetName];
    if (!selectedSheetName || !sheet) {
      updateLog("請先選擇工作表。", "warn");
      return null;
    }

    return sheet;
  }

  function runAction(actionId: ActionId) {
    if (actionId === "clean-all") {
      if (!workbook) {
        updateLog("尚未載入檔案", "warn");
        return;
      }

      runTransform("刪除所有工作表前 5 列並過濾指定列處理中…", () => cleanAllSheets(workbook, originalFileName));
      return;
    }

    const sheet = getSelectedSheet();
    if (!sheet) return;

    if (actionId === "momo-pchome") {
      runTransform("功能一（MOMO 銷貨單 + 合併）處理中…", () => transformMomoPchome(sheet, originalFileName), "功能一發生錯誤");
    }

    if (actionId === "broadcast") {
      runTransform("功能二（轉播單）處理中…", () => transformBroadcast(sheet, originalFileName), "轉播單轉換發生錯誤");
    }

    if (actionId === "external-sales") {
      runTransform("功能三（外站特殊格式）處理中…", () => transformExternalSales(sheet, originalFileName), "SHOPEE&YAHOO&迪卡農轉檔發生錯誤");
    }

    if (actionId === "oms") {
      runTransform("功能四（匯入格式 → OMS）處理中…", () => transformOms(sheet, originalFileName), "功能四發生錯誤");
    }
  }

  function runTransform(loadingMessage: string, transform: () => TransformResult, errorPrefix = "刪除前列發生錯誤") {
    updateLog(loadingMessage);

    try {
      const result = transform();
      XLSX.writeFile(result.workbook, result.fileName);
      updateLog(result.message, "ok");
    } catch (err) {
      console.error(err);
      updateLog(`${errorPrefix}：${errorMessage(err)}`, "warn");
    }
  }

  return (
    <main className="app">
      <div className="topbar">
        <div>
          <h1>Excel ETL 工具</h1>
          <p className="subtitle">匯入 Excel、選擇工作表，再執行需要的轉換並下載結果。</p>
        </div>
      </div>

      <div className="layout">
        <section className="panel controls" aria-label="ETL 控制區">
          <div className="field">
            <label htmlFor="fileInput">Excel 檔案</label>
            <div className="filebox">
              <input id="fileInput" type="file" accept=".xlsx,.xls,.xlsm,.csv" onChange={handleFileChange} />
              <div className="meta">{fileMeta}</div>
            </div>
          </div>

          <div className="field">
            <label htmlFor="sheetSelect">工作表</label>
            <select
              id="sheetSelect"
              disabled={!workbook}
              value={selectedSheetName}
              onChange={(event) => setSelectedSheetName(event.target.value)}
            >
              {!workbook && <option value="">請先匯入檔案</option>}
              {workbook && sheetNames.length === 0 && <option value="">沒有工作表</option>}
              {sheetNames.map((sheetName) => (
                <option key={sheetName} value={sheetName}>
                  {sheetName}
                </option>
              ))}
            </select>
          </div>

          <div className="field">
            <label>ETL 程式</label>
            <div className="actions">
              {actions.map((action, index) => (
                <button
                  key={action.id}
                  className="action"
                  type="button"
                  disabled={action.id === "clean-all" ? !workbook : !ready}
                  onClick={() => runAction(action.id)}
                >
                  <span className="icon">{index + 1}</span>
                  <span>
                    <span className="action-title">{action.title}</span>
                    <span className="action-desc">{action.desc}</span>
                  </span>
                </button>
              ))}
            </div>
          </div>
        </section>

        <section className="workspace">
          <div className="panel status" data-kind={log.kind}>
            {log.message}
          </div>
          <section className="panel preview" aria-label="工作表預覽">
            <div className="preview-head">
              <h2 className="preview-title">資料預覽</h2>
              <span className="preview-note">{previewInfo.note}</span>
            </div>
            <div className="table-wrap">
              <PreviewTable rows={previewInfo.rows} maxCols={previewInfo.maxCols} hasWorkbook={Boolean(workbook)} />
            </div>
          </section>
        </section>
      </div>
    </main>
  );
}

function PreviewTable({
  rows,
  maxCols,
  hasWorkbook
}: {
  rows: Row[];
  maxCols: number;
  hasWorkbook: boolean;
}) {
  if (!hasWorkbook) {
    return <div className="empty">尚無資料可預覽</div>;
  }

  if (!rows.length) {
    return <div className="empty">這個工作表沒有資料</div>;
  }

  return (
    <table>
      <thead>
        <tr>
          {Array.from({ length: maxCols }, (_, index) => (
            <th key={columnName(index)}>{columnName(index)}</th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((row, rowIndex) => (
          <tr key={rowIndex}>
            {Array.from({ length: maxCols }, (_, colIndex) => (
              <td key={`${rowIndex}-${colIndex}`}>{formatCell(row[colIndex])}</td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function formatCell(value: Row[number]) {
  if (value === undefined || value === null) return "";
  return String(value);
}

function errorMessage(err: unknown) {
  return err instanceof Error ? err.message : String(err);
}
