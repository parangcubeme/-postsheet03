"use client";

import { useMemo, useRef, useState } from "react";
import { applyCategoryReferences } from "../category-reference";
import type { CategoryReference } from "../category-reference";
import { hasCompleteEsmCategory, makeEsmRows } from "../excel";
import type { Product, WorkerResponse } from "../types";

type CategoryWorkerResponse =
  | { ok: true; references: CategoryReference[]; rowCount: number; sheetCount: number }
  | { ok: false; message: string };

export default function CategoryAutoPage() {
  const categoryWorkerRef = useRef<Worker | null>(null);
  const productWorkerRef = useRef<Worker | null>(null);
  const [references, setReferences] = useState<CategoryReference[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [categoryFileName, setCategoryFileName] = useState("");
  const [productFileName, setProductFileName] = useState("");
  const [status, setStatus] = useState("1단계: 카테고리 목록 엑셀을 올려주세요.");
  const [busy, setBusy] = useState(false);
  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");

  const pending = useMemo(
    () => products.filter((product) => !hasCompleteEsmCategory(product)).length,
    [products],
  );
  const matched = products.length - pending;

  async function uploadCategory(file?: File) {
    if (!file) return;
    categoryWorkerRef.current?.terminate();
    setBusy(true);
    setCategoryFileName(file.name);
    setStatus("카테고리 엑셀 전체를 읽고 있습니다.");

    const worker = new Worker(new URL("../category-reference.worker.ts", import.meta.url), { type: "module" });
    categoryWorkerRef.current = worker;
    worker.onmessage = (event: MessageEvent<CategoryWorkerResponse>) => {
      const response = event.data;
      if (!response.ok) {
        setStatus(response.message);
      } else {
        setReferences(response.references);
        setProducts((current) => {
          if (!current.length) return current;
          return applyCategoryReferences(current, response.references).products;
        });
        setStatus(`카테고리 기준 ${response.references.length.toLocaleString()}개를 읽었습니다. 이제 상품 엑셀을 올려주세요.`);
      }
      setBusy(false);
      worker.terminate();
      categoryWorkerRef.current = null;
    };
    worker.onerror = () => {
      setStatus("카테고리 파일을 읽는 중 오류가 발생했습니다.");
      setBusy(false);
      worker.terminate();
      categoryWorkerRef.current = null;
    };
    const buffer = await file.arrayBuffer();
    worker.postMessage(buffer, [buffer]);
  }

  async function uploadProducts(file?: File) {
    if (!file) return;
    if (!references.length) {
      setStatus("먼저 카테고리 목록 엑셀을 올려주세요.");
      return;
    }

    productWorkerRef.current?.terminate();
    setBusy(true);
    setProductFileName(file.name);
    setStatus("상품을 읽고 카테고리 엑셀 전체와 비교해 초벌 분류하고 있습니다.");

    const worker = new Worker(new URL("../excel.worker.ts", import.meta.url), { type: "module" });
    productWorkerRef.current = worker;
    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const response = event.data;
      if (!response.ok) {
        setStatus(response.message);
      } else {
        const result = applyCategoryReferences(response.products, references);
        setProducts(result.products);
        setStatus(`상품 ${result.products.length.toLocaleString()}개 중 ${result.matched.toLocaleString()}개를 카테고리 엑셀로 자동분류했습니다. 확인 필요 ${result.pending.toLocaleString()}개입니다.`);
      }
      setBusy(false);
      worker.terminate();
      productWorkerRef.current = null;
    };
    worker.onerror = () => {
      setStatus("상품 파일 분석 중 오류가 발생했습니다.");
      setBusy(false);
      worker.terminate();
      productWorkerRef.current = null;
    };
    const buffer = await file.arrayBuffer();
    worker.postMessage(buffer, [buffer]);
  }

  async function download() {
    if (!products.length) return;
    if (!auctionId.trim() || !gmarketId.trim()) {
      setStatus("옥션 판매자 ID와 G마켓 판매자 ID를 입력해 주세요.");
      return;
    }
    const XLSX = await import("xlsx");
    const worksheet = XLSX.utils.aoa_to_sheet(makeEsmRows(products, auctionId.trim(), gmarketId.trim()));
    worksheet["!cols"] = Array.from({ length: 64 }, () => ({ wch: 16 }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "NEW 일반상품");
    XLSX.writeFile(workbook, `ESM_자동분류_${products.length}개.xlsx`);
    setStatus("자동분류 결과 엑셀 다운로드를 시작했습니다.");
  }

  return (
    <main className="wrap">
      <section className="hero">
        <span className="badge">CATEGORY AUTO</span>
        <h1>카테고리 엑셀 자동분류</h1>
        <p>카테고리 목록 전체를 기준으로 상품명을 초벌 분류하고, 확신이 낮은 상품만 확인 대상으로 남깁니다.</p>
      </section>

      <section className="panel">
        <div className="form-grid">
          <label className="upload-box">
            <strong>1. 카테고리 목록 엑셀</strong>
            <span>{categoryFileName || ".xls 또는 .xlsx"}</span>
            <input type="file" accept=".xls,.xlsx" disabled={busy} onChange={(event) => void uploadCategory(event.currentTarget.files?.[0])} />
          </label>
          <label className="upload-box">
            <strong>2. 상품 엑셀</strong>
            <span>{productFileName || "카테고리 파일을 먼저 올린 뒤 상품 파일 업로드"}</span>
            <input type="file" accept=".xls,.xlsx,.csv" disabled={busy || !references.length} onChange={(event) => void uploadProducts(event.currentTarget.files?.[0])} />
          </label>
        </div>

        <div className="form-grid seller-fields" style={{ marginTop: 18 }}>
          <label className="field"><span>옥션 판매자 ID</span><input value={auctionId} onChange={(event) => setAuctionId(event.target.value)} /></label>
          <label className="field"><span>G마켓 판매자 ID</span><input value={gmarketId} onChange={(event) => setGmarketId(event.target.value)} /></label>
        </div>

        {products.length > 0 && (
          <>
            <div className={pending ? "status summary-status" : "success-box"}>
              전체 {products.length.toLocaleString()}개 · 자동분류 {matched.toLocaleString()}개 · 확인 필요 {pending.toLocaleString()}개
            </div>
            <div className="review-table-wrap" style={{ marginTop: 18 }}>
              <table className="review-table">
                <thead><tr><th>No.</th><th>상품명</th><th>분류결과</th><th>ESM</th><th>G마켓</th><th>옥션</th></tr></thead>
                <tbody>
                  {products.slice(0, 100).map((product, index) => {
                    const needsReview = !hasCompleteEsmCategory(product);
                    return (
                      <tr key={product.id} className={needsReview ? "row-needs-review" : "row-ready"}>
                        <td>{index + 1}</td>
                        <td>{product.productName}</td>
                        <td className={needsReview ? "cell-error" : "ready-cell"}>{product.categoryGroup}</td>
                        <td className={!product.categoryCode ? "cell-error" : ""}>{product.categoryCode || "확인 필요"}</td>
                        <td className={!product.gmarketExposureCode ? "cell-error" : ""}>{product.gmarketExposureCode || "확인 필요"}</td>
                        <td className={!product.auctionExposureCode ? "cell-error" : ""}>{product.auctionExposureCode || "확인 필요"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            <div className="actions">
              <button type="button" className="primary" disabled={busy} onClick={() => void download()}>자동분류 ESM 한 파일 다운로드</button>
            </div>
          </>
        )}

        <div className="status" role="status">{busy && <span className="status-dot" />}{status}</div>
      </section>
    </main>
  );
}
