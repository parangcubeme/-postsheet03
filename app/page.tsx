"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import {
  digits,
  getProductIssues,
  groupProducts,
  hasCompleteEsmCategory,
  hasCompleteNotice,
  hasCompleteOrigin,
  hasCompleteShipping,
  makeEsmRows,
  naverHeaders,
  validateProducts,
} from "./excel";
import type {
  MarketTab,
  Product,
  ProductGroup,
  StepId,
  WorkerResponse,
} from "./types";

const steps: { id: StepId; label: string }[] = [
  { id: "category", label: "1. 카테고리" },
  { id: "notice", label: "2. 고시정보" },
  { id: "origin", label: "3. 원산지" },
  { id: "shipping", label: "4. 배송정책" },
  { id: "price", label: "5. 가격" },
  { id: "review", label: "6. 최종검사·다운로드" },
];

const GROUP_PAGE_SIZE = 30;
const REVIEW_PAGE_SIZE = 50;

function groupByName(products: Product[], getName: (product: Product) => string): ProductGroup[] {
  const map = new Map<string, Product[]>();
  for (const product of products) {
    const name = getName(product);
    const current = map.get(name);
    if (current) current.push(product);
    else map.set(name, [product]);
  }
  return [...map.entries()].map(([name, items]) => ({ name, items }));
}

function makeEditorGroups(products: Product[], step: StepId): ProductGroup[] {
  if (step === "category") {
    return groupByName(
      products.filter((product) => !hasCompleteEsmCategory(product)),
      (product) => product.categoryGroup || "미분류 제품군",
    );
  }

  if (step === "notice") {
    return groupByName(
      products.filter((product) => !hasCompleteNotice(product)),
      (product) => `${product.categoryGroup} / ${product.productGroupCode || "상품군 미입력"} / ${product.noticeTemplateCode || "고시 미입력"}`,
    );
  }

  if (step === "shipping") {
    return groupByName(products, (product) => {
      const values = [
        product.departureCode,
        product.shippingPolicyNumber,
        product.returnAddressCode,
        product.auctionShippingPolicy,
        product.gmarketShippingPolicy,
        product.courierCode,
      ].filter(Boolean);
      return values.length ? values.join(" / ") : "배송정보 미입력";
    });
  }

  return groupProducts(products, step);
}

function groupNeedsReview(group: ProductGroup, step: StepId): boolean {
  if (step === "category") return group.items.some((product) => !hasCompleteEsmCategory(product));
  if (step === "notice") return group.items.some((product) => !hasCompleteNotice(product));
  if (step === "origin") return group.items.some((product) => !hasCompleteOrigin(product));
  if (step === "shipping") return group.items.some((product) => !hasCompleteShipping(product));
  return false;
}

function formatMoney(value: number): string {
  return Number.isFinite(value) ? `${Math.round(value).toLocaleString()}원` : "-";
}

export default function Page() {
  const workerRef = useRef<Worker | null>(null);
  const [market, setMarket] = useState<MarketTab>("smartstore");
  const [products, setProducts] = useState<Product[]>([]);
  const [sourceRowCount, setSourceRowCount] = useState(0);
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("상품 엑셀을 업로드해 주세요.");
  const [busy, setBusy] = useState(false);

  const [feeRate, setFeeRate] = useState(6);
  const [smartMargin, setSmartMargin] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [smartRound, setSmartRound] = useState(100);
  const [smartCategory, setSmartCategory] = useState("50001770");
  const [smartCourier, setSmartCourier] = useState("CJGLS");
  const [asPhone, setAsPhone] = useState("");

  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");
  const [step, setStep] = useState<StepId>("category");
  const [selectedProductId, setSelectedProductId] = useState("");
  const [groupPage, setGroupPage] = useState(0);
  const [reviewPage, setReviewPage] = useState(0);
  const [esmMargin, setEsmMargin] = useState(0);
  const [esmRound, setEsmRound] = useState(100);

  useEffect(() => {
    return () => workerRef.current?.terminate();
  }, []);

  const groups = useMemo(() => makeEditorGroups(products, step), [products, step]);
  const categoryManualCount = useMemo(
    () => products.filter((product) => !hasCompleteEsmCategory(product)).length,
    [products],
  );
  const noticeManualCount = useMemo(
    () => products.filter((product) => !hasCompleteNotice(product)).length,
    [products],
  );
  const categoryReadyCount = products.length - categoryManualCount;
  const noticeReadyCount = products.length - noticeManualCount;
  const selectedGroupIndex = selectedProductId
    ? groups.findIndex((group) => group.items.some((item) => item.id === selectedProductId))
    : groups.length
      ? 0
      : -1;
  const currentGroup = selectedGroupIndex >= 0 ? groups[selectedGroupIndex] : groups[0];
  const visibleGroups = useMemo(
    () => groups.slice(groupPage * GROUP_PAGE_SIZE, groupPage * GROUP_PAGE_SIZE + GROUP_PAGE_SIZE),
    [groups, groupPage],
  );
  const validationMessages = useMemo(() => validateProducts(products), [products]);
  const reviewPageCount = Math.max(1, Math.ceil(products.length / REVIEW_PAGE_SIZE));
  const visibleReviewProducts = useMemo(
    () => products.slice(reviewPage * REVIEW_PAGE_SIZE, reviewPage * REVIEW_PAGE_SIZE + REVIEW_PAGE_SIZE),
    [products, reviewPage],
  );

  useEffect(() => {
    if (selectedGroupIndex < 0) return;
    const nextPage = Math.floor(selectedGroupIndex / GROUP_PAGE_SIZE);
    setGroupPage((current) => (current === nextPage ? current : nextPage));
  }, [selectedGroupIndex]);

  useEffect(() => {
    if (reviewPage < reviewPageCount) return;
    setReviewPage(Math.max(0, reviewPageCount - 1));
  }, [reviewPage, reviewPageCount]);

  async function uploadFile(file?: File) {
    if (!file) return;

    workerRef.current?.terminate();
    setBusy(true);
    setFileName(file.name);
    setStatus("엑셀을 분석하고 ESM 전용 형식에 맞춰 옮기고 있습니다. 탭과 입력창은 계속 사용할 수 있습니다.");

    try {
      const worker = new Worker(new URL("./excel.worker.ts", import.meta.url), { type: "module" });
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

        const categoryPending = response.products.filter((product) => !hasCompleteEsmCategory(product)).length;
        const noticePending = response.products.filter((product) => !hasCompleteNotice(product)).length;
        setProducts(response.products);
        setSourceRowCount(response.rows.length);
        setStep("category");
        setSelectedProductId("");
        setGroupPage(0);
        setReviewPage(0);
        setStatus(
          `${response.sheetName} 시트 ${response.headerRow}행을 제목으로 인식했습니다. ${response.products.length}개 상품을 ESM 형식으로 옮겼습니다. 카테고리 확인 ${categoryPending}개, 고시정보 확인 ${noticePending}개입니다.`,
        );
        setBusy(false);
        worker.terminate();
        workerRef.current = null;
      };

      worker.onerror = () => {
        setStatus("엑셀 분석 중 오류가 발생했습니다. 오류 내용을 확인할 수 있도록 파일 형식과 제목 행을 다시 검사해 주세요.");
        setBusy(false);
        worker.terminate();
        workerRef.current = null;
      };

      const buffer = await file.arrayBuffer();
      worker.postMessage(buffer, [buffer]);
    } catch (error) {
      setStatus(error instanceof Error ? `파일 준비 오류: ${error.message}` : "파일을 읽지 못했습니다.");
      setBusy(false);
      workerRef.current?.terminate();
      workerRef.current = null;
    }
  }

  function selectStep(nextStep: StepId) {
    setStep(nextStep);
    setSelectedProductId("");
    setGroupPage(0);
    if (nextStep === "review") setReviewPage(0);
  }

  function patchCurrentGroup(patch: Partial<Product>) {
    if (!currentGroup) return;
    const anchorId = currentGroup.items[0]?.id ?? "";
    const ids = new Set(currentGroup.items.map((item) => item.id));
    setSelectedProductId(anchorId);
    setProducts((current) => current.map((product) => (ids.has(product.id) ? { ...product, ...patch } : product)));
    setStatus(`${currentGroup.items.length}개 상품에 적용했습니다.`);
  }

  function applyEsmPrices() {
    const unit = Math.max(1, esmRound);
    setProducts((current) =>
      current.map((product) => ({
        ...product,
        finalPrice:
          esmMargin === 0
            ? product.basePrice
            : Math.ceil((product.basePrice * (1 + esmMargin / 100)) / unit) * unit,
      })),
    );
    setStatus(`${products.length}개 상품의 ESM 판매가를 적용했습니다.`);
  }

  async function downloadSmartStore() {
    if (!products.length) return;
    setBusy(true);
    setStatus("스마트스토어 엑셀을 만들고 있습니다.");

    try {
      const XLSX = await import("xlsx");
      const unit = Math.max(1, smartRound);
      const body = products.map((product) => {
        const salePrice =
          Math.ceil(((product.basePrice + extraCost) * (1 + (feeRate + smartMargin) / 100)) / unit) * unit;
        const shippingFee = Math.max(0, product.shippingFee);
        const returnFee = shippingFee || 3000;
        const multipleOrigins = ["Y", "복수원산지"].includes(product.multipleOrigins) ? "Y" : "N";

        return [
          product.sellerCode,
          product.categoryCode || smartCategory,
          product.productName,
          "신상품",
          salePrice,
          product.vatType || "과세상품",
          product.stock,
          product.mainImage,
          product.additionalImage,
          product.detailHtml,
          product.originRegionCode || "03",
          multipleOrigins,
          product.originDirect,
          "Y",
          "택배, 소포, 등기",
          smartCourier,
          shippingFee > 0 ? "유료" : "무료",
          shippingFee,
          returnFee,
          returnFee * 2,
          "N",
          asPhone,
          "판매자에게 문의하시거나 A/S 연락처로 문의해 주세요.",
          "Y",
          "N",
        ];
      });

      const worksheet = XLSX.utils.aoa_to_sheet([naverHeaders, ...body]);
      worksheet["!cols"] = naverHeaders.map(() => ({ wch: 18 }));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "일괄등록");
      XLSX.writeFile(workbook, "postsheet03_스마트스토어.xlsx");
      setStatus("스마트스토어 엑셀 다운로드를 시작했습니다.");
    } catch (error) {
      setStatus(error instanceof Error ? `다운로드 오류: ${error.message}` : "엑셀을 만들지 못했습니다.");
    } finally {
      setBusy(false);
    }
  }

  async function downloadEsmWorkbook() {
    if (!products.length) return;
    if (!auctionId.trim() || !gmarketId.trim()) {
      setStatus("옥션 판매자 ID와 G마켓 판매자 ID를 먼저 입력해 주세요.");
      return;
    }
    if (validationMessages.length) {
      setStatus("빨간색으로 표시된 미입력 항목을 수정한 뒤 최종 엑셀을 다운로드해 주세요.");
      return;
    }

    setBusy(true);
    setStatus(`${products.length}개 상품을 한 개의 ESM 엑셀 파일로 만들고 있습니다.`);

    try {
      const XLSX = await import("xlsx");
      const worksheet = XLSX.utils.aoa_to_sheet(makeEsmRows(products, auctionId.trim(), gmarketId.trim()));
      worksheet["!cols"] = Array.from({ length: 64 }, (_, index) => ({ wch: index === 4 || index === 27 ? 42 : 16 }));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "NEW 일반상품");
      XLSX.writeFile(workbook, `ESM_전체상품_${products.length}개.xlsx`);
      setStatus(`${products.length}개 상품이 들어간 ESM 엑셀 한 파일 다운로드를 시작했습니다.`);
    } catch (error) {
      setStatus(error instanceof Error ? `다운로드 오류: ${error.message}` : "ESM 엑셀을 만들지 못했습니다.");
    } finally {
      setBusy(false);
    }
  }

  function renderGroupEditor() {
    if (!currentGroup) return <div className="empty-box">수정할 상품이 없습니다.</div>;
    const first = currentGroup.items[0];

    if (step === "category") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <p className="description">제공된 카테고리 ESM 파일의 상품명과 코드를 먼저 자동 적용했습니다. 남은 제품군만 ESM → G마켓 → 옥션 순서로 입력해 주세요.</p>
          <div className="form-grid">
            <ApplyField label="1. ESM 카테고리 코드" value={first.categoryCode} numeric invalid={!first.categoryCode} onApply={(value) => patchCurrentGroup({ categoryCode: digits(value) })} />
            <ApplyField label="2. G마켓 노출코드" value={first.gmarketExposureCode} numeric invalid={!first.gmarketExposureCode} onApply={(value) => patchCurrentGroup({ gmarketExposureCode: digits(value) })} />
            <ApplyField label="3. 옥션 노출코드" value={first.auctionExposureCode} numeric invalid={!first.auctionExposureCode} onApply={(value) => patchCurrentGroup({ auctionExposureCode: digits(value) })} />
          </div>
        </GroupEditor>
      );
    }

    if (step === "notice") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <p className="description">카테고리 기준 파일에서 확인된 상품군·고시정보는 자동 입력했습니다. 빨간 항목만 제품군에 맞게 수정해 주세요.</p>
          <div className="form-grid">
            <ApplyField label="상품군 코드" value={first.productGroupCode} numeric invalid={!first.productGroupCode} onApply={(value) => patchCurrentGroup({ productGroupCode: digits(value) })} />
            <ApplyField label="상품고시정보 템플릿코드" value={first.noticeTemplateCode} numeric invalid={!first.noticeTemplateCode} onApply={(value) => patchCurrentGroup({ noticeTemplateCode: digits(value) })} />
          </div>
        </GroupEditor>
      );
    }

    if (step === "origin") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <div className="form-grid">
            <ApplyField label="원산지 상품타입" value={first.originProductType} invalid={!first.originProductType} onApply={(value) => patchCurrentGroup({ originProductType: value })} />
            <ApplyField label="원산지 지역타입" value={first.originRegionType} invalid={!first.originRegionType} onApply={(value) => patchCurrentGroup({ originRegionType: value })} />
            <ApplyField label="원산지 지역코드" value={first.originRegionCode} numeric invalid={!first.originRegionCode} onApply={(value) => patchCurrentGroup({ originRegionCode: digits(value) })} />
            <ApplyField label="복수원산지 여부" value={first.multipleOrigins} invalid={!first.multipleOrigins} onApply={(value) => patchCurrentGroup({ multipleOrigins: value })} />
          </div>
        </GroupEditor>
      );
    }

    if (step === "shipping") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <div className="form-grid">
            <ApplyField label="출하지 코드" value={first.departureCode} numeric invalid={!first.departureCode} onApply={(value) => patchCurrentGroup({ departureCode: digits(value) })} />
            <ApplyField label="배송정책번호" value={first.shippingPolicyNumber} numeric invalid={!first.shippingPolicyNumber} onApply={(value) => patchCurrentGroup({ shippingPolicyNumber: digits(value) })} />
            <ApplyField label="반품·교환 주소 코드" value={first.returnAddressCode} numeric invalid={!first.returnAddressCode} onApply={(value) => patchCurrentGroup({ returnAddressCode: digits(value) })} />
            <ApplyField label="옥션 발송정책" value={first.auctionShippingPolicy} numeric invalid={!first.auctionShippingPolicy} onApply={(value) => patchCurrentGroup({ auctionShippingPolicy: digits(value) })} />
            <ApplyField label="G마켓 발송정책" value={first.gmarketShippingPolicy} numeric invalid={!first.gmarketShippingPolicy} onApply={(value) => patchCurrentGroup({ gmarketShippingPolicy: digits(value) })} />
            <ApplyField label="택배사코드" value={first.courierCode} numeric invalid={!first.courierCode} onApply={(value) => patchCurrentGroup({ courierCode: digits(value) })} />
            <ApplyField label="반품·교환 배송비" value={first.returnShippingFee} numeric invalid={!(first.returnShippingFee >= 0)} onApply={(value) => patchCurrentGroup({ returnShippingFee: Number(value) || 0 })} />
          </div>
        </GroupEditor>
      );
    }

    return null;
  }

  return (
    <main className="wrap">
      <section className="hero">
        <span className="badge">POSTSHEET03</span>
        <h1>상품 엑셀 변환</h1>
        <p>업로드한 상품 정보를 ESM 전용 엑셀의 맞는 열로 옮기고, 수정이 필요한 항목만 확인한 뒤 한 파일로 다운로드합니다.</p>
        <div className="hero-points">
          <span>서버 저장 없음</span>
          <span>ZIP 사용 안 함</span>
          <span>ESM 한 엑셀 파일</span>
        </div>
      </section>

      <section className="panel">
        <div className="tabs" aria-label="판매처 선택">
          <button type="button" className={market === "smartstore" ? "active" : ""} onClick={() => setMarket("smartstore")}>스마트스토어</button>
          <button type="button" className={market === "esm" ? "active" : ""} onClick={() => setMarket("esm")}>ESM · 옥션/G마켓</button>
        </div>

        <label className="upload-box">
          <strong>상품 일괄목록 엑셀 업로드</strong>
          <span>.xlsx, .xls, .csv · 안내문과 빈 행이 있어도 제목 행을 찾고 ESM 전용 열로 옮깁니다.</span>
          <input
            type="file"
            accept=".xlsx,.xls,.csv"
            disabled={busy}
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              void uploadFile(file);
              event.currentTarget.value = "";
            }}
          />
        </label>

        {fileName && (
          <div className="file-summary">
            <strong>{fileName}</strong>
            <span>원본 행 {sourceRowCount.toLocaleString()}개 · 변환 가능 상품 {products.length.toLocaleString()}개</span>
          </div>
        )}

        {market === "smartstore" ? (
          <section className="market-section">
            <div className="section-heading">
              <div>
                <span className="eyebrow">SMART STORE</span>
                <h2>스마트스토어 설정</h2>
              </div>
              <span>{products.length.toLocaleString()}개 상품</span>
            </div>

            <div className="form-grid">
              <NumberField label="네이버 수수료율 (%)" value={feeRate} onChange={setFeeRate} />
              <NumberField label="추가 마진율 (%)" value={smartMargin} onChange={setSmartMargin} />
              <NumberField label="상품당 추가비용" value={extraCost} onChange={setExtraCost} />
              <NumberField label="판매가 올림 단위" value={smartRound} onChange={setSmartRound} min={1} />
              <TextField label="기본 카테고리코드" value={smartCategory} onChange={setSmartCategory} />
              <TextField label="택배사코드" value={smartCourier} onChange={setSmartCourier} />
              <TextField label="A/S 전화번호" value={asPhone} onChange={setAsPhone} placeholder="연락처 입력" />
            </div>

            <div className="actions">
              <button type="button" className="primary" disabled={busy || products.length === 0} onClick={() => void downloadSmartStore()}>
                스마트스토어 엑셀 다운로드
              </button>
            </div>
          </section>
        ) : (
          <section className="market-section">
            <div className="section-heading">
              <div>
                <span className="eyebrow">ESM PLUS</span>
                <h2>옥션·G마켓 설정</h2>
              </div>
              <span>{products.length.toLocaleString()}개 상품</span>
            </div>

            <div className="form-grid seller-fields">
              <TextField label="옥션 판매자 ID" value={auctionId} onChange={setAuctionId} />
              <TextField label="G마켓 판매자 ID" value={gmarketId} onChange={setGmarketId} />
            </div>

            <nav className="stepbar" aria-label="ESM 작업 단계">
              {steps.map((item) => (
                <button key={item.id} type="button" className={step === item.id ? "active" : ""} onClick={() => selectStep(item.id)}>
                  {item.label}
                </button>
              ))}
            </nav>

            {products.length > 0 && (
              <div className={categoryManualCount || noticeManualCount ? "status summary-status" : "success-box"}>
                카테고리 자동·기존 입력 {categoryReadyCount.toLocaleString()}개 / 직접 확인 {categoryManualCount.toLocaleString()}개 · 고시정보 입력 {noticeReadyCount.toLocaleString()}개 / 직접 확인 {noticeManualCount.toLocaleString()}개
              </div>
            )}

            {products.length === 0 ? (
              <div className="empty-box">먼저 상품 엑셀을 업로드해 주세요.</div>
            ) : step === "category" && groups.length === 0 ? (
              <div className="success-box">모든 상품의 ESM·G마켓·옥션 카테고리가 입력되었습니다. 다음 단계로 이동해 주세요.</div>
            ) : step === "notice" && groups.length === 0 ? (
              <div className="success-box">모든 상품의 상품군 코드와 고시정보 템플릿코드가 입력되었습니다. 다음 단계로 이동해 주세요.</div>
            ) : ["category", "notice", "origin", "shipping"].includes(step) ? (
              <div className="group-workspace">
                <aside className="group-list">
                  <div className="group-list-head">
                    <strong>{groups.length.toLocaleString()}개 묶음</strong>
                    <span>{step === "category" ? `${categoryManualCount.toLocaleString()}개 확인` : step === "notice" ? `${noticeManualCount.toLocaleString()}개 확인` : `${products.length.toLocaleString()}개 상품`}</span>
                  </div>
                  {visibleGroups.map((group) => {
                    const selected = (currentGroup?.name ?? "") === group.name;
                    const needsReview = groupNeedsReview(group, step);
                    return (
                      <button
                        key={group.name}
                        type="button"
                        className={`${selected ? "selected" : ""}${needsReview ? " needs-review" : ""}`.trim()}
                        onClick={() => setSelectedProductId(group.items[0]?.id ?? "")}
                      >
                        <strong>{group.name}</strong>
                        <span>{group.items.length.toLocaleString()}개</span>
                      </button>
                    );
                  })}
                  {groups.length > GROUP_PAGE_SIZE && (
                    <div className="pager">
                      <button type="button" disabled={groupPage === 0} onClick={() => setGroupPage((page) => Math.max(0, page - 1))}>이전</button>
                      <span>{groupPage + 1} / {Math.ceil(groups.length / GROUP_PAGE_SIZE)}</span>
                      <button type="button" disabled={(groupPage + 1) * GROUP_PAGE_SIZE >= groups.length} onClick={() => setGroupPage((page) => page + 1)}>다음</button>
                    </div>
                  )}
                </aside>
                <div>{renderGroupEditor()}</div>
              </div>
            ) : step === "price" ? (
              <div className="editor-card">
                <div className="editor-heading">
                  <div>
                    <span className="eyebrow">PRICE</span>
                    <h3>ESM 전체 가격 적용</h3>
                  </div>
                  <span>{products.length.toLocaleString()}개</span>
                </div>
                <div className="form-grid">
                  <NumberField label="마진율 (%)" value={esmMargin} onChange={setEsmMargin} />
                  <NumberField label="판매가 올림 단위" value={esmRound} onChange={setEsmRound} min={1} />
                </div>
                <div className="actions">
                  <button type="button" className="primary" onClick={applyEsmPrices}>전체 상품 가격 적용</button>
                </div>
              </div>
            ) : (
              <div className="review-stack">
                <section className="editor-card">
                  <div className="editor-heading">
                    <div>
                      <span className="eyebrow">FINAL CHECK</span>
                      <h3>최종검사</h3>
                    </div>
                    <span>{validationMessages.length ? `${validationMessages.length}종 확인 필요` : "전체 항목 확인"}</span>
                  </div>
                  {validationMessages.length ? (
                    <ul className="warning-list">
                      {validationMessages.map((message) => <li key={message}>{message}</li>)}
                    </ul>
                  ) : (
                    <div className="success-box">모든 상품의 필수항목이 입력되었습니다.</div>
                  )}
                </section>

                <section className="editor-card viewer-card">
                  <div className="editor-heading">
                    <div>
                      <span className="eyebrow">ESM FILE VIEWER</span>
                      <h3>최종 엑셀 목록 미리보기</h3>
                    </div>
                    <span>{products.length.toLocaleString()}개 상품</span>
                  </div>
                  <p className="description">빨간 칸은 수정이 필요한 값입니다. 최대 50개씩 보여주며 최종 다운로드에는 전체 상품이 한 엑셀 파일에 들어갑니다.</p>
                  <div className="review-table-wrap">
                    <table className="review-table">
                      <thead>
                        <tr>
                          <th>No.</th>
                          <th>상품명</th>
                          <th>ESM 카테고리</th>
                          <th>G마켓</th>
                          <th>옥션</th>
                          <th>상품군</th>
                          <th>고시정보</th>
                          <th>원산지</th>
                          <th>배송정책</th>
                          <th>판매가</th>
                          <th>검사결과</th>
                        </tr>
                      </thead>
                      <tbody>
                        {visibleReviewProducts.map((product, index) => {
                          const issues = getProductIssues(product);
                          const originText = [product.originProductType, product.originRegionType, product.originRegionCode, product.multipleOrigins].filter(Boolean).join(" / ");
                          const shippingText = [product.departureCode, product.shippingPolicyNumber, product.returnAddressCode, product.auctionShippingPolicy, product.gmarketShippingPolicy, product.courierCode].filter(Boolean).join(" / ");
                          return (
                            <tr key={product.id} className={issues.length ? "row-needs-review" : "row-ready"}>
                              <td>{reviewPage * REVIEW_PAGE_SIZE + index + 1}</td>
                              <td className={!product.productName ? "cell-error" : ""}>{product.productName || "미입력"}</td>
                              <td className={!product.categoryCode ? "cell-error" : ""}>{product.categoryCode || "미입력"}</td>
                              <td className={!product.gmarketExposureCode ? "cell-error" : ""}>{product.gmarketExposureCode || "미입력"}</td>
                              <td className={!product.auctionExposureCode ? "cell-error" : ""}>{product.auctionExposureCode || "미입력"}</td>
                              <td className={!product.productGroupCode ? "cell-error" : ""}>{product.productGroupCode || "미입력"}</td>
                              <td className={!product.noticeTemplateCode ? "cell-error" : ""}>{product.noticeTemplateCode || "미입력"}</td>
                              <td className={!hasCompleteOrigin(product) ? "cell-error" : ""}>{originText || "미입력"}</td>
                              <td className={!hasCompleteShipping(product) ? "cell-error" : ""}>{shippingText || "미입력"}</td>
                              <td className={product.finalPrice > 0 ? "" : "cell-error"}>{formatMoney(product.finalPrice)}</td>
                              <td className={issues.length ? "cell-error issue-cell" : "ready-cell"}>{issues.length ? issues.join(", ") : "확인완료"}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  {reviewPageCount > 1 && (
                    <div className="pager viewer-pager">
                      <button type="button" disabled={reviewPage === 0} onClick={() => setReviewPage((page) => Math.max(0, page - 1))}>이전 50개</button>
                      <span>{reviewPage + 1} / {reviewPageCount}</span>
                      <button type="button" disabled={reviewPage + 1 >= reviewPageCount} onClick={() => setReviewPage((page) => page + 1)}>다음 50개</button>
                    </div>
                  )}
                </section>

                <section className="editor-card final-download-card">
                  <div>
                    <span className="eyebrow">ONE EXCEL FILE</span>
                    <h3>최종 ESM 엑셀 다운로드</h3>
                    <p className="description">카테고리별 파일 분리나 ZIP 압축 없이 전체 상품을 공식 ESM 형식의 한 엑셀 파일로 만듭니다.</p>
                  </div>
                  <button type="button" className="primary" disabled={busy || validationMessages.length > 0} onClick={() => void downloadEsmWorkbook()}>
                    전체 {products.length.toLocaleString()}개 상품 한 엑셀 다운로드
                  </button>
                </section>
              </div>
            )}
          </section>
        )}

        <div className="status" role="status" aria-live="polite">
          {busy && <span className="status-dot" aria-hidden="true" />}
          {status}
        </div>
        <div className="privacy">원본 엑셀, 변환 결과와 개인정보는 서버에 저장하지 않습니다.</div>
      </section>
    </main>
  );
}

function TextField({
  label,
  value,
  onChange,
  placeholder,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <input value={value} placeholder={placeholder} onChange={(event) => onChange(event.target.value)} />
    </label>
  );
}

function NumberField({
  label,
  value,
  onChange,
  min,
}: {
  label: string;
  value: number;
  onChange: (value: number) => void;
  min?: number;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      <input type="number" min={min} value={value} onChange={(event) => onChange(Number(event.target.value))} />
    </label>
  );
}

function ApplyField({
  label,
  value,
  onApply,
  numeric = false,
  invalid = false,
}: {
  label: string;
  value: string | number;
  onApply: (value: string) => void;
  numeric?: boolean;
  invalid?: boolean;
}) {
  const [draft, setDraft] = useState(String(value ?? ""));

  useEffect(() => {
    setDraft(String(value ?? ""));
  }, [value]);

  return (
    <label className={`field apply-field${invalid ? " invalid" : ""}`}>
      <span>{label}{invalid ? " · 수정 필요" : ""}</span>
      <input type={numeric ? "number" : "text"} value={draft} onChange={(event) => setDraft(event.target.value)} />
      <button type="button" onClick={() => onApply(draft)}>이 묶음에 적용</button>
    </label>
  );
}

function GroupEditor({ groupName, items, children }: { groupName: string; items: Product[]; children: ReactNode }) {
  return (
    <section className="editor-card">
      <div className="editor-heading">
        <div>
          <span className="eyebrow">GROUP EDIT</span>
          <h3>{groupName}</h3>
        </div>
        <span>{items.length.toLocaleString()}개 상품</span>
      </div>
      <details className="sample-products">
        <summary>상품명 확인</summary>
        <ul>
          {items.slice(0, 10).map((item) => <li key={item.id}>{item.productName}</li>)}
        </ul>
      </details>
      <div className="editor-fields">{children}</div>
    </section>
  );
}
