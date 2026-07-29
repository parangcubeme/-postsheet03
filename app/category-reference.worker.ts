/// <reference lib="webworker" />

import type { Row } from "./types";
import { makeCategoryReferences } from "./category-reference";

function clean(value: unknown): string {
  return String(value ?? "").trim();
}

function cleanKey(value: unknown): string {
  return clean(value).replace(/[\s\n\r_\-()\[\]\/·>]/g, "").toLowerCase();
}

function detectHeaderRow(rows: unknown[][]): number {
  const keywords = [
    "카테고리명", "카테고리", "분류명", "세분류", "소분류", "최종카테고리",
    "카테고리코드", "esm카테고리코드", "a노출코드", "g노출코드", "옥션노출코드", "g마켓노출코드",
  ].map(cleanKey);

  let bestIndex = 0;
  let bestScore = -1;
  rows.slice(0, 60).forEach((row, index) => {
    const keys = row.map(cleanKey);
    const score = keywords.reduce((sum, keyword) => sum + (keys.includes(keyword) ? 1 : 0), 0);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });
  return bestIndex;
}

self.onmessage = async (event: MessageEvent<ArrayBuffer>) => {
  try {
    const XLSX = await import("xlsx");
    const workbook = XLSX.read(event.data, { type: "array", raw: false });
    const allRows: Row[] = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
      if (!matrix.length) continue;
      const headerIndex = detectHeaderRow(matrix);
      const headers = matrix[headerIndex].map((value, index) => clean(value) || `열${index + 1}`);
      for (const values of matrix.slice(headerIndex + 1)) {
        const row: Row = {};
        headers.forEach((header, index) => {
          row[header] = values[index] ?? "";
        });
        if (Object.values(row).some((value) => clean(value))) allRows.push(row);
      }
    }

    const references = makeCategoryReferences(allRows);
    self.postMessage({ ok: true, references, rowCount: allRows.length, sheetCount: workbook.SheetNames.length });
  } catch (error) {
    self.postMessage({
      ok: false,
      message: error instanceof Error ? `카테고리 파일 분석 오류: ${error.message}` : "카테고리 파일을 읽지 못했습니다.",
    });
  }
};

export {};
