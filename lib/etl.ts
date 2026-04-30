import * as XLSX from "xlsx";

export type CellValue = string | number | boolean | Date | null | undefined;
export type Row = CellValue[];
export type Aoa = Row[];

export type TransformResult = {
  workbook: XLSX.WorkBook;
  fileName: string;
  message: string;
};

export function val(row: Row | undefined, idx: number) {
  const v = row && row[idx];
  return v === undefined || v === null ? "" : String(v).trim();
}

export function getBaseName(fileName: string) {
  return String(fileName || "輸出").replace(/\.[^/.]+$/, "");
}

export function columnName(index: number) {
  let n = index + 1;
  let name = "";
  while (n > 0) {
    const mod = (n - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    n = Math.floor((n - mod) / 26);
  }
  return name;
}

export function sheetRows(sheet: XLSX.WorkSheet, defval: CellValue = "") {
  return XLSX.utils.sheet_to_json<Row>(sheet, {
    header: 1,
    raw: false,
    defval
  });
}

function newWorkbook(sheetName: string, rows: Aoa) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return workbook;
}

export function transformMomoPchome(sheet: XLSX.WorkSheet, originalFileName: string): TransformResult {
  const rows = sheetRows(sheet);
  if (!rows || rows.length <= 1) {
    throw new Error("來源資料不足。");
  }

  const data = rows.slice(1);
  const headers = [
    "銷貨日期", "客戶代號(固定)", "部門代號(固定)", "通路訂單編號",
    "備註", "品號", "數量", "金額", "通路訂單序號(不填)",
    "庫別", "發票地址一(固定)", "客戶全名(固定)"
  ];

  const momo: Aoa = [headers];
  let momoCount = 0;

  data.forEach((row) => {
    const customerCode = val(row, 16);
    const isMomo = customerCode === "27365925" || customerCode === "27365925DS";
    const isNew = customerCode === "16606102";

    if (isMomo || isNew) {
      const orderNo = isNew ? val(row, 4) : val(row, 11);

      momo.push([
        "",
        customerCode,
        val(row, 17),
        orderNo,
        val(row, 3),
        val(row, 9),
        val(row, 13),
        val(row, 15),
        "",
        val(row, 20),
        val(row, 19),
        val(row, 18)
      ]);
      momoCount++;
    }
  });

  const merged: Aoa = [headers];
  const dict: Record<string, number> = {};

  for (let i = 1; i < momo.length; i++) {
    const row = momo[i];
    const orderNo = row[3];
    const itemNo = row[5];
    const key = `${orderNo}||${itemNo}`;

    if (!dict[key]) {
      dict[key] = merged.length;
      merged.push([...row]);
    } else {
      const idx = dict[key];
      const oldQty = Number(merged[idx][6] || 0);
      const newQty = Number(row[6] || 0);
      merged[idx][6] = oldQty + newQty;

      const oldAmt = Number(merged[idx][7] || 0);
      const newAmt = Number(row[7] || 0);
      merged[idx][7] = oldAmt + newAmt;
    }
  }

  const fileName = `MOMO&PCHOME銷貨單${getBaseName(originalFileName)}.xlsx`;
  return {
    workbook: newWorkbook("合併後資料", merged),
    fileName,
    message:
      "MOMO銷貨單完成！\n" +
      `原始篩選資料：${momoCount} 筆\n` +
      `銷貨單資料：${merged.length - 1} 筆\n` +
      `已下載：${fileName}`
  };
}

export function transformBroadcast(sheet: XLSX.WorkSheet, originalFileName: string): TransformResult {
  const rows = sheetRows(sheet);
  if (!rows || rows.length <= 1) {
    throw new Error("來源資料不足，無法轉換。");
  }

  const header = rows[0];
  const data = rows.slice(1);

  function findCol(keyword: string) {
    const normalizedKeyword = keyword.replace(/ /g, "");
    for (let i = 0; i < header.length; i++) {
      const h = String(header[i] || "").replace(/[\n ]/g, "");
      if (h.includes(normalizedKeyword)) return i + 1;
    }
    return 0;
  }

  const c撥出日期 = findCol("撥出日期");
  const c單號 = findCol("單號");
  const c撥出倉庫 = findCol("撥出倉庫");
  const c撥入倉庫 = findCol("撥入倉庫");
  const c商品代號 = findCol("商品代號");
  const c商品名稱 = findCol("商品名稱");
  const c撥出數量 = findCol("撥出數量");

  const required: Array<[string, number]> = [
    ["撥出日期", c撥出日期],
    ["單號", c單號],
    ["撥出倉庫", c撥出倉庫],
    ["撥入倉庫", c撥入倉庫],
    ["商品代號", c商品代號],
    ["商品名稱", c商品名稱],
    ["撥出數量", c撥出數量]
  ];
  for (const [name, col] of required) {
    if (col === 0) {
      throw new Error(`來源缺少必要欄位：「${name}」`);
    }
  }

  const COL = {
    單別: 2,
    單號: 3,
    部門代號: 5,
    備註: 6,
    廠別代號: 9,
    單據日期: 15,
    確認者: 16,
    簽核狀態碼: 18,
    運輸方式: 21,
    BN單別: 66,
    BP序號: 68,
    品號: 69,
    品名: 70,
    數量: 72,
    單位: 73,
    轉出庫: 78,
    轉入庫: 79
  };
  const maxCol = 79;
  const aoa: Aoa = [];
  const title = new Array<CellValue>(maxCol).fill("");

  title[COL.單別 - 1] = "單別";
  title[COL.單號 - 1] = "單號";
  title[COL.部門代號 - 1] = "部門代號";
  title[COL.備註 - 1] = "備註";
  title[COL.廠別代號 - 1] = "廠別代號";
  title[COL.單據日期 - 1] = "單據日期";
  title[COL.確認者 - 1] = "確認者";
  title[COL.簽核狀態碼 - 1] = "簽核狀態碼";
  title[COL.運輸方式 - 1] = "運輸方式";
  title[COL.BN單別 - 1] = "單別";
  title[COL.BP序號 - 1] = "序號";
  title[COL.品號 - 1] = "品號";
  title[COL.品名 - 1] = "品名";
  title[COL.數量 - 1] = "數量";
  title[COL.單位 - 1] = "單位";
  title[COL.轉出庫 - 1] = "轉出庫";
  title[COL.轉入庫 - 1] = "轉入庫";

  aoa.push(title);

  let serial = 1;
  for (let i = 0; i < data.length; i++) {
    if (!val(data[i], c商品代號 - 1)) continue;

    const newRow = new Array<CellValue>(maxCol).fill("");

    newRow[COL.單別 - 1] = "122";
    newRow[COL.部門代號 - 1] = "6999";
    newRow[COL.廠別代號 - 1] = "001";
    newRow[COL.確認者 - 1] = "E11075";
    newRow[COL.簽核狀態碼 - 1] = "N";
    newRow[COL.運輸方式 - 1] = "5";

    newRow[COL.單據日期 - 1] = data[i][c撥出日期 - 1];
    newRow[COL.備註 - 1] = data[i][c單號 - 1];
    newRow[COL.品號 - 1] = data[i][c商品代號 - 1];
    newRow[COL.品名 - 1] = data[i][c商品名稱 - 1];
    newRow[COL.數量 - 1] = data[i][c撥出數量 - 1];
    newRow[COL.轉出庫 - 1] = data[i][c撥出倉庫 - 1];
    newRow[COL.轉入庫 - 1] = data[i][c撥入倉庫 - 1];

    newRow[COL.BN單別 - 1] = "122";
    newRow[COL.BP序號 - 1] = serial++;

    aoa.push(newRow);
  }

  const fileName = `轉播單_${getBaseName(originalFileName)}.xlsx`;
  return {
    workbook: newWorkbook("轉播單", aoa),
    fileName,
    message: `轉播單轉換完成！\n已下載：${fileName}\n總筆數：${aoa.length - 1}`
  };
}

export function transformExternalSales(sheet: XLSX.WorkSheet, originalFileName: string): TransformResult {
  const rows = sheetRows(sheet);
  if (!rows || rows.length <= 1) {
    throw new Error("來源資料不足。");
  }

  const data = rows.slice(1);
  const targetCustomers = ["56801904A", "56801904D", "27240313A", "54880333", "53048094"];

  const tempRows = data
    .filter((r) => targetCustomers.includes(val(r, 16)))
    .map((r) => {
      const cust = val(r, 16);
      const useE = ["56801904A", "56801904D", "54880333", "53048094"].includes(cust);
      const orderNo = useE ? val(r, 4) : val(r, 11);

      let pn = val(r, 9);
      if (pn === "seller_discount") pn = "100004-000-000";
      if (pn === "SHIPFEE") pn = "100001-000-000";
      if (["platform_coin", "Discount"].includes(pn)) return null;

      return {
        客戶代號: cust,
        部門代號: val(r, 17),
        通路訂單編號: orderNo,
        備註: val(r, 3),
        品號: pn,
        數量: pn === "100004-000-000" || pn === "100001-000-000" ? 0 : val(r, 13),
        金額: val(r, 15),
        庫別: val(r, 20),
        發票地址: val(r, 19),
        客戶全名: val(r, 18),
        序號: 0
      };
    })
    .filter((row): row is NonNullable<typeof row> => row !== null);

  const groups: Record<string, number> = {};
  tempRows.forEach((r) => {
    const key = `${r.通路訂單編號}|||${r.備註}`;
    groups[key] = (groups[key] || 0) + 1;
    r.序號 = groups[key];
  });

  tempRows.sort((a, b) => {
    if (a.通路訂單編號 < b.通路訂單編號) return -1;
    if (a.通路訂單編號 > b.通路訂單編號) return 1;
    return a.序號 - b.序號;
  });

  const out: Aoa = [[
    "銷貨日期", "客戶代號(固定)", "部門代號(固定)", "通路訂單編號",
    "備註", "品號", "數量", "金額", "通路訂單序號(要填)",
    "庫別", "發票地址一(固定)", "客戶全名(固定)"
  ]];

  tempRows.forEach((r) => {
    out.push([
      "",
      r.客戶代號,
      r.部門代號,
      r.通路訂單編號,
      r.備註,
      r.品號,
      r.數量,
      r.金額,
      r.序號,
      r.庫別,
      r.發票地址,
      r.客戶全名
    ]);
  });

  const fileName = `SHOPEE&YAHOO銷貨單${getBaseName(originalFileName)}.xlsx`;
  return {
    workbook: newWorkbook("外站轉檔", out),
    fileName,
    message: `SHOPEE&YAHOO&迪卡農轉檔完成！\n已下載：${fileName}\n總筆數：${tempRows.length}`
  };
}

export function transformOms(sheet: XLSX.WorkSheet, originalFileName: string): TransformResult {
  const rows = sheetRows(sheet);
  if (rows.length < 3) {
    throw new Error("來源資料太少，無法轉換。");
  }

  function findCol(keyword: string) {
    for (let r = 0; r < 10 && r < rows.length; r++) {
      const row = rows[r];
      for (let c = 0; c < row.length; c++) {
        if (row[c] && row[c]?.toString().includes(keyword)) {
          return c + 1;
        }
      }
    }
    return 0;
  }

  function dateToYYYYMMDD(v: CellValue) {
    if (!v) return "";
    if (!Number.isNaN(Date.parse(String(v)))) {
      const d = new Date(String(v));
      return (
        d.getFullYear() +
        String(d.getMonth() + 1).padStart(2, "0") +
        String(d.getDate()).padStart(2, "0")
      );
    }
    const digits = v.toString().replace(/\D/g, "");
    return digits.length >= 8 ? digits.slice(0, 8) : digits;
  }

  function lastDataRow(colIndexes: number[]) {
    let last = 0;
    colIndexes.forEach((col) => {
      if (!col) return;
      for (let r = rows.length - 1; r >= 0; r--) {
        const v = rows[r][col - 1];
        if (v !== "" && v !== undefined && v !== null) {
          last = Math.max(last, r + 1);
          break;
        }
      }
    });
    return last;
  }

  const col品名 = findCol("品名");
  const col品號 = findCol("品號");
  const col數量 = findCol("訂單數量");
  const col單價 = findCol("單價");
  const col金額 = findCol("金額");

  const col單號 = findCol("單號");
  const col序號 = findCol("序號");
  const col單別 = findCol("單別");
  const col單據日期 = findCol("單據日期") || findCol("訂單日期");

  const col賣場編號 = findCol("賣場編號");
  const col收件人姓名 = findCol("收件人姓名");
  const col收件人電話 = findCol("收件人電話");
  const col收件人地址 = findCol("收件人地址");

  if (!col品名 || !col品號 || !col數量 || !col單價 || !col金額) {
    throw new Error("找不到必要欄位：品名 / 品號 / 訂單數量 / 單價 / 金額");
  }

  const lastRow = lastDataRow([col品名, col品號, col數量, col單價, col金額]);
  if (lastRow < 3) {
    throw new Error("來源資料不足，無法轉換。");
  }

  const omsHeaders = [
    "賣場訂單編號","賣場編號","訂購人姓名","收件人姓名","訂購人電話","收件人電話",
    "收件人郵遞區號","收件人地址","訂單成立日","商品名稱","商品貨號","實售價格",
    "商品單價","購買量","訂單總金額","代收款","訂單備註","付款方式","配送方式","規格",
    "交易序號","超商門市代碼","超商名稱","配送編號","收件國家","收件省分","收件城市",
    "群品","批號","儲位","效期","上架日","季節"
  ];

  const out: Aoa = [omsHeaders];
  const dictLine: Record<string, number> = {};
  let outRow = 2;

  for (let r = 3; r <= lastRow; r++) {
    const row = rows[r - 1];
    const 品名 = row[col品名 - 1];
    const 品號 = row[col品號 - 1];
    const 數量 = row[col數量 - 1];

    if (!品名 && !品號 && !數量) continue;

    const output = new Array<CellValue>(omsHeaders.length).fill("");

    output[9] = 品名;
    output[10] = 品號;

    if (col賣場編號) output[1] = row[col賣場編號 - 1];
    if (col收件人姓名) output[2] = row[col收件人姓名 - 1];
    if (col收件人姓名) output[3] = row[col收件人姓名 - 1];
    if (col收件人電話) output[4] = row[col收件人電話 - 1];
    if (col收件人電話) output[5] = row[col收件人電話 - 1];
    if (col收件人地址) output[7] = row[col收件人地址 - 1];

    output[11] = 0;
    output[12] = "";
    output[13] = 數量;
    output[14] = 0;

    let seq: CellValue = "";
    const raw序號 = col序號 ? row[col序號 - 1] : "";
    const raw單號 = col單號 ? row[col單號 - 1] : "";

    if (raw序號) seq = raw序號;
    else if (raw單號) seq = raw單號;
    else {
      const key = String(raw單號 || "");
      dictLine[key] = (dictLine[key] || 0) + 1;
      seq = dictLine[key];
    }
    output[20] = seq;

    const prefix = col單別 ? row[col單別 - 1] : "";
    const dt = col單據日期 ? dateToYYYYMMDD(row[col單據日期 - 1]) : "";
    output[0] = `${prefix}${dt}${seq}`;

    output[18] = "99";
    output[19] = "104";

    out.push(output);
    outRow++;
  }

  const fileName = `OMS轉出格式_${getBaseName(originalFileName)}.xlsx`;
  return {
    workbook: newWorkbook("轉出OMS格式", out),
    fileName,
    message: `銷貨格式轉OMS格式完成！\n共輸出 ${outRow - 2} 筆資料。\n已下載：${fileName}`
  };
}

export function cleanAllSheets(workbook: XLSX.WorkBook, originalFileName: string): TransformResult {
  const outWb = XLSX.utils.book_new();

  workbook.SheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetRows(sheet);

    if (!rows || rows.length === 0) {
      XLSX.utils.book_append_sheet(outWb, XLSX.utils.aoa_to_sheet([]), sheetName);
      return;
    }

    let newRows = rows.slice(5);

    newRows = newRows.filter((row) => {
      const colA = row[0] !== undefined && row[0] !== null ? String(row[0]).trim() : "";
      const colS = row[18] !== undefined && row[18] !== null ? String(row[18]).trim() : "";

      const removeAValues = [
        "銷售員小計：",
        "現金：",
        "信用卡：",
        "有價券類：",
        "其他：",
        "補登金額："
      ];

      if (removeAValues.includes(colA)) return false;
      if (colS === "尾款") return false;

      return true;
    });

    newRows = newRows.map((row) => {
      const copied = [...row];

      if (copied[14] !== undefined && copied[14] !== "") {
        const num = Number(String(copied[14]).replace(/,/g, ""));
        copied[14] = Number.isNaN(num) ? copied[14] : num;
      }

      if (copied[15] !== undefined && copied[15] !== "") {
        const num = Number(String(copied[15]).replace(/,/g, ""));
        copied[15] = Number.isNaN(num) ? copied[15] : num;
      }

      return copied;
    });

    newRows = newRows.map((row) => {
      return row.filter((_cell, idx) => {
        const removeIndexes = [1, 3, 4, 5, 6, 7, 8, 12, 13, 17];
        return !removeIndexes.includes(idx);
      });
    });

    const safeSheetName = sheetName.slice(0, 31) || "Sheet";
    XLSX.utils.book_append_sheet(outWb, XLSX.utils.aoa_to_sheet(newRows), safeSheetName);
  });

  const fileName = `過濾後_${getBaseName(originalFileName)}.xlsx`;
  return {
    workbook: outWb,
    fileName,
    message:
      "處理完成！\n" +
      "已刪除每個工作表前 5 列\n" +
      "已刪除 A 欄為指定小計或付款文字的列\n" +
      "已刪除 S 欄為「尾款」的列\n" +
      `已下載：${fileName}`
  };
}
