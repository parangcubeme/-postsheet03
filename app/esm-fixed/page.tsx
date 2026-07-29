"use client";

import { useMemo, useRef, useState } from "react";
import { makeEsmRows, validateProducts } from "../excel";
import type { Product, WorkerResponse } from "../types";

const MAX_PRODUCTS_PER_FILE = 500;
const DATA_START_ROW_INDEX = 7; // 엑셀 8행
const TARGET_SHEET_NAME = "NEW 일반상품";

export default function EsmFixedPage() {
  const workerRef = useRef<Worker | null>(null);
  const [products, setProducts] = useState<Product[]>([]);
  const [sourceName, setSourceName] = useState("");
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("상품 파일과 공식 ESM 템플릿을 업로드해 주세요.");

  const validationMessages = useMemo(() => validateProducts(products), [products]);

  async function uploadProducts(file?: File) {
    if (!file) return;

    workerRef.current?.terminate();
    setBusy(true);
    setSourceName(file.name);
    setStatus("상품 파일을 분석하고 있습니다.");

    try {
      const worker = new Worker(new URL("../excel.worker.ts", import.meta.url), { type: "module" });
      workerRef.current = worker;

      worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
        const response = event.data;
        if (!response.ok) {
          setStatus(response.message);
          setBusy(false);
          worker.terminate();
          workerRef.current = null;
          return;
        }

        setProducts(response.products);
        setStatus(`${response.products.length.toLocaleString()}개 상품을 읽었습니다.`);
        setBusy(false);
        worker.terminate();
        workerRef.current = null;
      };

      worker.onerror = () => {
        setStatus("상품 파일 분석 중 오류가 발생했습니다.");
        setBusy(false);
        worker.terminate();
        workerRef.current = null;
      };

      const buffer = await file.arrayBuffer();
      worker.postMessage(buffer, [buffer]);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "상품 파일을 읽지 못했습니다.");
      setBusy(false);
    }
  }

  async function downloadWithOfficialTemplate() {
    if (!products.length) {
      setStatus("상품 파일을 먼저 업로드해 주세요.");
      return;
    }
    if (!templateFile) {
      setStatus("ESM에서 내려받은 공식 일반상품 일괄등록 템플릿을 업로드해 주세요.");
      return;
    }
    if (!auctionId.trim() || !gmarketId.trim()) {
      setStatus("옥션 ID와 G마켓 ID를 입력해 주세요.");
      return;
    }
    if (validationMessages.length) {
      setStatus(`필수값 오류 ${validationMessages.length}종이 남아 있어 다운로드를 중지했습니다.`);
      return;
    }

    setBusy(true);
    setStatus("공식 ESM 템플릿의 1~7행을 보존하고 8행부터 상품을 입력하고 있습니다.");

    try {
      const XLSX = await import("xlsx");
      const templateBuffer = await templateFile.arrayBuffer();
      const chunks = Array.from(
        { length: Math.ceil(products.length / MAX_PRODUCTS_PER_FILE) },
        (_, index) => products.slice(index * MAX_PRODUCTS_PER_FILE, (index + 1) * MAX_PRODUCTS_PER_FILE),
      );

      for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex += 1) {
        const workbook = XLSX.read(templateBuffer.slice(0), {
          type: "array",
          cellStyles: true,
          cellDates: true,
          cellFormula: true,
          bookVBA: true,
        });

        const sheetName = workbook.SheetNames.includes(TARGET_SHEET_NAME)
          ? TARGET_SHEET_NAME
          : workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) throw new Error("공식 템플릿 시트를 찾지 못했습니다.");

        const rows = makeEsmRows(chunks[chunkIndex], auctionId.trim(), gmarketId.trim());
        const bodyRows = rows.slice(DATA_START_ROW_INDEX);

        const originalRange = worksheet["!ref"] ? XLSX.utils.decode_range(worksheet["!ref"]) : null;
        if (originalRange) {
          for (let row = DATA_START_ROW_INDEX; row <= originalRange.e.r; row += 1) {
            for (let col = 0; col <= originalRange.e.c; col += 1) {
              delete worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
            }
          }
        }

        XLSX.utils.sheet_add_aoa(worksheet, bodyRows, { origin: { r: DATA_START_ROW_INDEX, c: 0 } });

        const finalColumnCount = Math.max(originalRange?.e.c ?? 0, bodyRows[0]?.length ? bodyRows[0].length - 1 : 0);
        worksheet["!ref"] = XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: DATA_START_ROW_INDEX + bodyRows.length - 1, c: finalColumnCount },
        });

        const suffix = chunks.length > 1 ? `_${chunkIndex + 1}` : "";
        XLSX.writeFile(workbook, `ESM_전체상품_${chunks[chunkIndex].length}개${suffix}.xlsx`, {
          bookType: "xlsx",
          cellStyles: true,
          bookVBA: true,
        });
      }

      setStatus(`${products.length.toLocaleString()}개 상품을 공식 템플릿 기준으로 저장했습니다.`);
    } catch (error) {
      setStatus(error instanceof Error ? `다운로드 오류: ${error.message}` : "ESM 파일 생성에 실패했습니다.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main style={{ maxWidth: 920, margin: "40px auto", padding: 24, fontFamily: "sans-serif" }}>
      <h1>ESM 공식 템플릿 변환</h1>
      <p>새 통합문서를 만들지 않고 공식 ESM 파일의 1~7행과 시트 구조를 유지한 채 8행부터 상품을 입력합니다.</p>

      <section style={{ display: "grid", gap: 16, marginTop: 24 }}>
        <label>
          <strong>1. 상품 엑셀</strong><br />
          <input type="file" accept=".xlsx,.xls,.csv" disabled={busy} onChange={(event) => void uploadProducts(event.currentTarget.files?.[0])} />
        </label>

        <label>
          <strong>2. 공식 ESM 일반상품 템플릿</strong><br />
          <input type="file" accept=".xlsx,.xlsm" disabled={busy} onChange={(event) => setTemplateFile(event.currentTarget.files?.[0] ?? null)} />
        </label>

        <label>
          <strong>3. 옥션 판매자 ID</strong><br />
          <input value={auctionId} onChange={(event) => setAuctionId(event.target.value)} />
        </label>

        <label>
          <strong>4. G마켓 판매자 ID</strong><br />
          <input value={gmarketId} onChange={(event) => setGmarketId(event.target.value)} />
        </label>

        <div>
          <div>상품 파일: {sourceName || "미업로드"}</div>
          <div>공식 템플릿: {templateFile?.name || "미업로드"}</div>
          <div>상품 수: {products.length.toLocaleString()}개</div>
          <div>검사 오류: {validationMessages.length.toLocaleString()}종</div>
        </div>

        {validationMessages.length > 0 && (
          <ul style={{ color: "#b42318" }}>
            {validationMessages.map((message) => <li key={message}>{message}</li>)}
          </ul>
        )}

        <button type="button" disabled={busy || !products.length || !templateFile} onClick={() => void downloadWithOfficialTemplate()}>
          공식 ESM 템플릿으로 다운로드
        </button>

        <div role="status">{busy ? "처리 중 · " : ""}{status}</div>
      </section>
    </main>
  );
}
