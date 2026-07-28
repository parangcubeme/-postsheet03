/// <reference lib="webworker" />
import * as XLSX from "xlsx";
import type { Row, WorkerResponse } from "./types";

const PRODUCT_HEADERS = ["상품명", "제품명", "상품이름", "품명", "상품 명"];
const PRICE_HEADERS = ["판매가", "판매가격", "상품가격", "상품 판매가", "공급가", "공급가격", "판매단가", "단가", "기준가격", "원가", "매입가", "소비자가", "정상가", "가격"];
const CODE_HEADERS = ["판매자 상품코드", "판매자상품코드", "상품코드", "자체상품코드", "관리코드", "품목코드", "판매자코드"];

function cleanKey(value: unknown): string {
  return String(value ?? "").replace(/[\s\n\r_\-()\[\]\/]/g, "").toLowerCase();
}

function normalize(value: unknown): string {
  return String(value ?? "").trim();
}

function scoreHeader(row: unknown[]): number {
  const expected = [...PRODUCT_HEADERS, ...PRICE_HEADERS, ...CODE_HEADERS].map(cleanKey);
  return row
    .map(cleanKey)
    .filter(Boolean)
    .reduce((score, cell) => score + (expected.some((item) => cell === item || cell.includes(item) || item.includes(cell)) ? 1 : 0), 0);
}

function findBestSheet(workbook: XLSX.WorkBook): { sheetName: string; matrix: unknown[][]; headerIndex: number; score: number } {
  let best = { sheetName: workbook.SheetNames[0] ?? "Sheet1", matrix: [] as unknown[][], headerIndex: 0, score: -1 };

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
    const scanLimit = Math.min(matrix.length, 50);
    for (let index = 0; index < scanLimit; index += 1) {
      const score = scoreHeader(matrix[index] ?? []);
      if (score > best.score) best = { sheetName, matrix, headerIndex: index, score };
    }
  }

  return best;
}

function uniqueHeaders(values: unknown[]): string[] {
  const counts = new Map<string, number>();
  return values.map((value, index) => {
    const base = normalize(value) || `열${index + 1}`;
    const count = counts.get(base) ?? 0;
    counts.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

self.onmessage = (event: MessageEvent<ArrayBuffer>) => {
  try {
    const workbook = XLSX.read(event.data, { type: "array", cellDates: false });
    const { sheetName, matrix, headerIndex, score } = findBestSheet(workbook);
    if (!matrix.length || score < 1) {
      const response: WorkerResponse = { ok: false, message: "상품명 또는 가격 제목 행을 찾지 못했습니다." };
      self.postMessage(response);
      return;
    }

    const headers = uniqueHeaders(matrix[headerIndex] ?? []);
    const rows: Row[] = matrix
      .slice(headerIndex + 1)
      .map((values) => Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""])))
      .filter((row) => Object.values(row).some((value) => normalize(value)));

    const response: WorkerResponse = { ok: true, rows, sheetName, headerRow: headerIndex + 1 };
    self.postMessage(response);
  } catch (error) {
    const response: WorkerResponse = {
      ok: false,
      message: error instanceof Error ? `엑셀 분석 오류: ${error.message}` : "엑셀 파일을 읽지 못했습니다.",
    };
    self.postMessage(response);
  }
};

export {};
