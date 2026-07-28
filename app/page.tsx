"use client";

import { useMemo, useState } from "react";

type Row = Record<string, unknown>;
type Market = "smartstore" | "esm";
type Product = {
  sellerCode: string;
  name: string;
  price: number;
  stock: number;
  image: string;
  detail: string;
  shippingFee: number;
  vat: string;
  category: string;
};

const clean = (v: unknown) => String(v ?? "").trim();
const key = (v: unknown) => clean(v).replace(/[\s_\-()/\[\]]/g, "").toLowerCase();
const num = (v: unknown) => {
  const n = Number(clean(v).replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
};
const aliases = {
  name: ["상품명", "제품명", "품명", "상품이름"],
  code: ["판매자상품코드", "판매자 상품코드", "상품코드", "관리코드"],
  price: ["판매가", "판매가격", "상품가격", "공급가", "원가", "가격"],
  stock: ["재고수량", "재고", "수량", "판매가능수량"],
  image: ["대표이미지", "대표이미지URL", "메인이미지", "기본이미지", "이미지URL"],
  detail: ["상세설명", "상품설명", "상품상세설명", "상세HTML"],
  shipping: ["기본배송비", "배송비"],
  vat: ["부가세", "과세구분", "부가세유형"],
  category: ["카테고리코드", "카테고리 코드"],
};

function pick(row: Row, names: string[]) {
  const entries = Object.entries(row);
  for (const name of names) {
    const found = entries.find(([k]) => key(k) === key(name));
    if (found && clean(found[1])) return found[1];
  }
  return "";
}

function toProducts(rows: Row[]): Product[] {
  return rows.map((row, i) => ({
    sellerCode: clean(pick(row, aliases.code)) || `P${i + 1}`,
    name: clean(pick(row, aliases.name)),
    price: num(pick(row, aliases.price)),
    stock: num(pick(row, aliases.stock)) || 99999,
    image: clean(pick(row, aliases.image)),
    detail: clean(pick(row, aliases.detail)),
    shippingFee: num(pick(row, aliases.shipping)),
    vat: clean(pick(row, aliases.vat)) || "과세상품",
    category: clean(pick(row, aliases.category)),
  })).filter(p => p.name && p.price > 0);
}

export default function Page() {
  const [market, setMarket] = useState<Market>("smartstore");
  const [rows, setRows] = useState<Row[]>([]);
  const [products, setProducts] = useState<Product[]>([]);
  const [status, setStatus] = useState("상품 엑셀을 업로드해 주세요.");
  const [busy, setBusy] = useState(false);
  const [feeRate, setFeeRate] = useState(6);
  const [marginRate, setMarginRate] = useState(30);
  const [extraCost, setExtraCost] = useState(0);
  const [roundUnit, setRoundUnit] = useState(100);
  const [smartCategory, setSmartCategory] = useState("50001770");
  const [courier, setCourier] = useState("CJGLS");
  const [asPhone, setAsPhone] = useState("01027483227");
  const [auctionId, setAuctionId] = useState("");
  const [gmarketId, setGmarketId] = useState("");

  const categoryPages = useMemo(() => {
    const map = new Map<string, Product[]>();
    products.forEach(p => {
      const k = p.category || "미분류";
      map.set(k, [...(map.get(k) ?? []), p]);
    });
    return [...map.entries()].flatMap(([category, items]) =>
      Array.from({ length: Math.ceil(items.length / 500) }, (_, page) => ({
        category,
        page: page + 1,
        items: items.slice(page * 500, page * 500 + 500),
      }))
    );
  }, [products]);

  async function upload(file?: File) {
    if (!file) return;
    setBusy(true);
    setStatus("엑셀을 읽고 있습니다.");
    try {
      const XLSX = await import("xlsx");
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: false });
      const candidates = [...aliases.name, ...aliases.price].map(key);
      let headerIndex = 0;
      let best = -1;
      matrix.slice(0, 30).forEach((row, i) => {
        const score = row.map(key).filter(v => candidates.includes(v)).length;
        if (score > best) { best = score; headerIndex = i; }
      });
      const headers = (matrix[headerIndex] ?? []).map((v, i) => clean(v) || `열${i + 1}`);
      const parsedRows = matrix.slice(headerIndex + 1)
        .map(values => Object.fromEntries(headers.map((h, i) => [h, values[i] ?? ""])))
        .filter(row => Object.values(row).some(v => clean(v)));
      const parsed = toProducts(parsedRows);
      setRows(parsedRows);
      setProducts(parsed);
      setStatus(`${parsed.length}개 상품을 읽었습니다.`);
    } catch {
      setStatus("파일을 읽지 못했습니다. 엑셀 형식을 확인해 주세요.");
    } finally {
      setBusy(false);
    }
  }

  async function downloadSmart() {
    if (!rows.length) return;
    setBusy(true);
    try {
      const XLSX = await import("xlsx");
      const headers = ["판매자 상품코드","카테고리코드","상품명","상품상태","판매가","부가세","재고수량","대표이미지","추가이미지","상세설명","원산지코드","복수원산지여부","원산지 직접입력","미성년자 구매","배송방법","택배사코드","배송비유형","기본배송비","반품배송비","교환배송비","별도설치비","A/S 전화번호","A/S 안내","구매평 노출여부","알림받기 동의 고객 전용 여부"];
      const body = products.map(p => {
        const sale = Math.ceil(((p.price + extraCost) * (1 + (feeRate + marginRate) / 100)) / Math.max(1, roundUnit)) * Math.max(1, roundUnit);
        return [p.sellerCode, p.category || smartCategory, p.name, "신상품", sale, p.vat, p.stock, p.image, "", p.detail, "03", "N", "", "Y", "택배, 소포, 등기", courier, p.shippingFee > 0 ? "유료" : "무료", p.shippingFee, p.shippingFee || 3000, (p.shippingFee || 3000) * 2, "N", asPhone, "판매자에게 문의해 주세요.", "Y", "N"];
      });
      const ws = XLSX.utils.aoa_to_sheet([headers, ...body]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "일괄등록");
      XLSX.writeFile(wb, "postsheet03_스마트스토어.xlsx");
      setStatus("스마트스토어 파일 다운로드를 시작했습니다.");
    } finally { setBusy(false); }
  }

  async function downloadEsm(category: string, page: number, items: Product[]) {
    setBusy(true);
    try {
      const XLSX = await import("xlsx");
      const headers = ["노출 사이트","A ID","G ID","상품명","카테고리 코드","A 판매가","G 판매가","A 재고","G 재고","기본이미지","상품상세설명"];
      const body = items.map(p => ["옥션/G마켓", auctionId, gmarketId, p.name, p.category, p.price, p.price, p.stock, p.stock, p.image, p.detail]);
      const ws = XLSX.utils.aoa_to_sheet([headers, ...body]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "NEW 일반상품");
      XLSX.writeFile(wb, `ESM_${category}_${page}페이지_${items.length}개.xlsx`);
      setStatus("ESM 파일 다운로드를 시작했습니다.");
    } finally { setBusy(false); }
  }

  return (
    <main className="wrap">
      <section className="hero">
        <span className="badge">POSTSHEET03</span>
        <h1>상품 엑셀 변환</h1>
        <p>스마트스토어와 ESM 파일을 변환합니다. ZIP은 사용하지 않으며 엑셀 기능은 업로드·다운로드할 때만 불러옵니다.</p>
      </section>

      <section className="panel">
        <div className="tabs">
          <button className={market === "smartstore" ? "active" : ""} onClick={() => setMarket("smartstore")}>스마트스토어</button>
          <button className={market === "esm" ? "active" : ""} onClick={() => setMarket("esm")}>ESM · 옥션/G마켓</button>
        </div>

        <div className="grid">
          <label className="full">상품 일괄목록 엑셀 업로드
            <input type="file" accept=".xlsx,.xls,.csv" disabled={busy} onChange={e => upload(e.target.files?.[0])} />
          </label>

          {market === "smartstore" ? <>
            <label>네이버 수수료율 (%)<input type="number" value={feeRate} onChange={e => setFeeRate(Number(e.target.value))} /></label>
            <label>추가 마진율 (%)<input type="number" value={marginRate} onChange={e => setMarginRate(Number(e.target.value))} /></label>
            <label>상품당 추가비용<input type="number" value={extraCost} onChange={e => setExtraCost(Number(e.target.value))} /></label>
            <label>판매가 올림 단위<input type="number" value={roundUnit} onChange={e => setRoundUnit(Number(e.target.value))} /></label>
            <label>기본 카테고리코드<input value={smartCategory} onChange={e => setSmartCategory(e.target.value)} /></label>
            <label>택배사코드<input value={courier} onChange={e => setCourier(e.target.value)} /></label>
            <label>A/S 전화번호<input value={asPhone} onChange={e => setAsPhone(e.target.value)} /></label>
            <div className="actions full"><button disabled={busy || !products.length} onClick={downloadSmart}>스마트스토어 엑셀 다운로드</button></div>
          </> : <>
            <label>옥션 판매자 ID<input value={auctionId} onChange={e => setAuctionId(e.target.value)} /></label>
            <label>G마켓 판매자 ID<input value={gmarketId} onChange={e => setGmarketId(e.target.value)} /></label>
            <div className="summary full">
              <strong>{products.length}개 상품 · {categoryPages.length}개 다운로드 파일</strong>
              <div className="download-list">
                {categoryPages.map(p => <button key={`${p.category}-${p.page}`} disabled={busy} onClick={() => downloadEsm(p.category, p.page, p.items)}>{p.category} · {p.page}페이지 · {p.items.length}개 다운로드</button>)}
              </div>
            </div>
          </>}
        </div>

        <div className="status">{busy ? "처리 중 · " : ""}{status}</div>
        <div className="privacy">원본 엑셀과 개인정보는 서버에 저장하지 않습니다.</div>
      </section>
    </main>
  );
}
