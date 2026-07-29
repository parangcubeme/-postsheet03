"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { ReactNode } from "react";
import {
  digits,
  groupProducts,
  hasCompleteEsmCategory,
  makeDownloadPages,
  makeEsmRows,
  naverHeaders,
  validateProducts,
} from "./excel";
import type {
  DownloadPage,
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

function safeFileName(value: string): string {
  return value.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "_");
}

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
    const unresolved = products.filter((product) => !hasCompleteEsmCategory(product));
    return groupByName(unresolved, (product) => product.categoryGroup || "미분류 제품군");
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
  const categoryReadyCount = products.length - categoryManualCount;
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
  const downloadPages = useMemo<DownloadPage[]>(() => makeDownloadPages(products), [products]);
  const validationMessages = useMemo(() => validateProducts(products), [products]);

  useEffect(() => {
    if (selectedGroupIndex < 0) return;
    const nextPage = Math.floor(selectedGroupIndex / GROUP_PAGE_SIZE);
    setGroupPage((current) => (current === nextPage ? current : nextPage));
  }, [selectedGroupIndex]);

  async function uploadFile(file?: File) {
    if (!file) return;

    workerRef.current?.terminate();
    setBusy(true);
    setFileName(file.name);
    setStatus("엑셀을 분석하고 있습니다. 탭과 입력창은 계속 사용할 수 있습니다.");

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

        const manualCount = response.products.filter((product) => !hasCompleteEsmCategory(product)).length;
        const readyCount = response.products.length - manualCount;
        setProducts(response.products);
        setSourceRowCount(response.rows.length);
        setStep("category");
        setSelectedProductId("");
        setGroupPage(0);
        setStatus(
          `${response.sheetName} 시트 ${response.headerRow}행을 제목으로 인식했습니다. ${response.products.length}개 상품을 읽었습니다. 카테고리 자동·기존 입력 ${readyCount}개, 직접 확인 ${manualCount}개입니다.`,
        );
        setBusy(false);
        worker.terminate();
        workerRef.current = null;
      };

      worker.onerror = () => {
        setStatus("엑셀 분석 중 오류가 발생했습니다. 다른 형식의 파일로 다시 시도해 주세요.");
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
          product.multipleOrigins === "단일원산지" ? "N" : "Y",
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

  async function downloadEsmPage(page: DownloadPage) {
    if (!auctionId.trim() || !gmarketId.trim()) {
      setStatus("옥션 판매자 ID와 G마켓 판매자 ID를 먼저 입력해 주세요.");
      return;
    }

    const missingCategoryCount = page.items.filter((product) => !hasCompleteEsmCategory(product)).length;
    if (missingCategoryCount) {
      setStatus(`이 파일에 ESM·G마켓·옥션 카테고리가 덜 입력된 상품이 ${missingCategoryCount}개 있습니다. 1단계에서 먼저 확인해 주세요.`);
      return;
    }

    setBusy(true);
    setStatus(`${page.items.length}개 상품의 ESM 엑셀을 만들고 있습니다.`);

    try {
      const XLSX = await import("xlsx");
      const worksheet = XLSX.utils.aoa_to_sheet(makeEsmRows(page.items, auctionId.trim(), gmarketId.trim()));
      worksheet["!cols"] = Array.from({ length: 64 }, () => ({ wch: 16 }));
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "NEW 일반상품");
      const name = safeFileName(page.category);
      XLSX.writeFile(workbook, `ESM_카테고리_${name}_${page.page}페이지_${page.items.length}개.xlsx`);
      setStatus("ESM 엑셀 다운로드를 시작했습니다.");
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
          <p className="description">카테고리는 ESM → G마켓 → 옥션 순서로 입력합니다. 이미 자동 입력된 값은 그대로 두고 빈 값만 채우면 됩니다.</p>
          <div className="form-grid">
            <ApplyField
              label="1. ESM 카테고리 코드"
              value={first.categoryCode}
              numeric
              onApply={(value) => {
                const categoryCode = digits(value);
                patchCurrentGroup({
                  categoryCode,
                  categoryGroup: categoryCode ? `수동분류 · ${categoryCode}` : "미분류 제품군",
                });
              }}
            />
            <ApplyField label="2. G마켓 노출코드" value={first.gmarketExposureCode} numeric onApply={(value) => patchCurrentGroup({ gmarketExposureCode: digits(value) })} />
            <ApplyField label="3. 옥션 노출코드" value={first.auctionExposureCode} numeric onApply={(value) => patchCurrentGroup({ auctionExposureCode: digits(value) })} />
            <ApplyField label="상품군 코드" value={first.productGroupCode} numeric onApply={(value) => patchCurrentGroup({ productGroupCode: digits(value) })} />
          </div>
        </GroupEditor>
      );
    }

    if (step === "notice") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <ApplyField
            label="상품고시정보 템플릿코드"
            value={first.noticeTemplateCode}
            numeric
            onApply={(value) => patchCurrentGroup({ noticeTemplateCode: digits(value) })}
          />
        </GroupEditor>
      );
    }

    if (step === "origin") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <div className="form-grid">
            <ApplyField label="원산지 상품타입" value={first.originProductType} onApply={(value) => patchCurrentGroup({ originProductType: value })} />
            <ApplyField label="원산지 지역타입" value={first.originRegionType} onApply={(value) => patchCurrentGroup({ originRegionType: value })} />
            <ApplyField label="원산지 지역코드" value={first.originRegionCode} numeric onApply={(value) => patchCurrentGroup({ originRegionCode: digits(value) })} />
            <ApplyField label="복수원산지 여부" value={first.multipleOrigins} onApply={(value) => patchCurrentGroup({ multipleOrigins: value })} />
          </div>
        </GroupEditor>
      );
    }

    if (step === "shipping") {
      return (
        <GroupEditor groupName={currentGroup.name} items={currentGroup.items}>
          <div className="form-grid">
            <ApplyField label="출하지 코드" value={first.departureCode} numeric onApply={(value) => patchCurrentGroup({ departureCode: digits(value) })} />
            <ApplyField label="배송정책번호" value={first.shippingPolicyNumber} numeric onApply={(value) => patchCurrentGroup({ shippingPolicyNumber: digits(value) })} />
            <ApplyField label="반품·교환 주소 코드" value={first.returnAddressCode} numeric onApply={(value) => patchCurrentGroup({ returnAddressCode: digits(value) })} />
            <ApplyField label="옥션 발송정책" value={first.auctionShippingPolicy} numeric onApply={(value) => patchCurrentGroup({ auctionShippingPolicy: digits(value) })} />
            <ApplyField label="G마켓 발송정책" value={first.gmarketShippingPolicy} numeric onApply={(value) => patchCurrentGroup({ gmarketShippingPolicy: digits(value) })} />
            <ApplyField label="택배사코드" value={first.courierCode} numeric onApply={(value) => patchCurrentGroup({ courierCode: digits(value) })} />
            <ApplyField label="반품·교환 배송비" value={first.returnShippingFee} numeric onApply={(value) => patchCurrentGroup({ returnShippingFee: Number(value) || 0 })} />
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
        <p>스마트스토어와 ESM 옥션·G마켓 등록 파일을 브라우저에서 변환합니다.</p>
        <div className="hero-points">
          <span>서버 저장 없음</span>
          <span>ZIP 사용 안 함</span>
          <span>ESM 최대 500개씩 분리</span>
        </div>
      </section>

      <section className="panel">
        <div className="tabs" aria-label="판매처 선택">
          <button type="button" className={market === "smartstore" ? "active" : ""} onClick={() => setMarket("smartstore")}>스마트스토어</button>
          <button type="button" className={market === "esm" ? "active" : ""} onClick={() => setMarket("esm")}>ESM · 옥션/G마켓</button>
        </div>

        <label className="upload-box">
          <strong>상품 일괄목록 엑셀 업로드</strong>
          <span>.xlsx, .xls, .csv · 안내문과 빈 행이 있어도 제목 행을 탐색합니다.</span>
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
              <div className={categoryManualCount ? "status" : "success-box"}>
                카테고리 3종 자동·기존 입력 {categoryReadyCount.toLocaleString()}개 · 직접 확인 {categoryManualCount.toLocaleString()}개 · 입력 순서 ESM → G마켓 → 옥션
              </div>
            )}

            {products.length === 0 ? (
              <div className="empty-box">먼저 상품 엑셀을 업로드해 주세요.</div>
            ) : step === "category" && groups.length === 0 ? (
              <div className="success-box">모든 상품의 ESM·G마켓·옥션 카테고리가 입력되었습니다. 다음 단계로 이동해 주세요.</div>
            ) : ["category", "notice", "origin", "shipping"].includes(step) ? (
              <div className="group-workspace">
                <aside className="group-list">
                  <div className="group-list-head">
                    <strong>{step === "category" ? `${groups.length.toLocaleString()}개 직접 확인 묶음` : `${groups.length.toLocaleString()}개 묶음`}</strong>
                    <span>{step === "category" ? `${categoryManualCount.toLocaleString()}개 상품` : `${products.length.toLocaleString()}개 상품`}</span>
                  </div>
                  {visibleGroups.map((group) => (
                    <button
                      key={group.name}
                      type="button"
                      className={(currentGroup?.name ?? "") === group.name ? "selected" : ""}
                      onClick={() => setSelectedProductId(group.items[0]?.id ?? "")}
                    >
                      <strong>{group.name}</strong>
                      <span>{group.items.length.toLocaleString()}개</span>
                    </button>
                  ))}
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
              <div className="review-grid">
                <section className="editor-card">
                  <div className="editor-heading">
                    <div>
                      <span className="eyebrow">FINAL CHECK</span>
                      <h3>최종검사</h3>
                    </div>
                    <span>{validationMessages.length ? `${validationMessages.length}개 확인 필요` : "필수항목 확인"}</span>
                  </div>
                  {validationMessages.length ? (
                    <ul className="warning-list">
                      {validationMessages.map((message) => <li key={message}>{message}</li>)}
                    </ul>
                  ) : (
                    <div className="success-box">현재 자동검사에서 누락된 필수항목이 없습니다.</div>
                  )}
                </section>

                <section className="editor-card">
                  <div className="editor-heading">
                    <div>
                      <span className="eyebrow">DOWNLOAD</span>
                      <h3>카테고리별 개별 다운로드</h3>
                    </div>
                    <span>{downloadPages.length.toLocaleString()}개 파일</span>
                  </div>
                  <p className="description">카테고리별 최대 500개씩 나눕니다. ZIP 압축은 사용하지 않습니다.</p>
                  <div className="download-list">
                    {downloadPages.map((page) => {
                      const categoryIncomplete = page.items.some((product) => !hasCompleteEsmCategory(product));
                      return (
                        <button key={page.key} type="button" disabled={busy || categoryIncomplete} onClick={() => void downloadEsmPage(page)}>
                          <strong>카테고리 {page.category}</strong>
                          <span>{categoryIncomplete ? "ESM·G마켓·옥션 코드 확인 필요" : `${page.page}페이지 · ${page.items.length.toLocaleString()}개 다운로드`}</span>
                        </button>
                      );
                    })}
                  </div>
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
}: {
  label: string;
  value: string | number;
  onApply: (value: string) => void;
  numeric?: boolean;
}) {
  const [draft, setDraft] = useState(String(value ?? ""));

  useEffect(() => {
    setDraft(String(value ?? ""));
  }, [value]);

  return (
    <label className="field apply-field">
      <span>{label}</span>
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
