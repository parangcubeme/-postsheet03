import type { DownloadPage, Product, ProductGroup, Row, StepId } from "./types";

export const aliases = {
  productName: ["상품명", "제품명", "상품이름", "품명", "상품 명"],
  sellerCode: ["판매자 상품코드", "판매자상품코드", "상품코드", "자체상품코드", "관리코드", "품목코드", "판매자코드"],
  basePrice: ["판매가", "A 판매가", "G 판매가", "판매가격", "상품가격", "상품 판매가", "공급가", "공급가격", "판매단가", "단가", "기준가격", "원가", "매입가", "소비자가", "정상가", "가격"],
  stock: ["재고수량", "재고", "수량", "판매가능수량", "A 재고", "G 재고"],
  mainImage: ["대표이미지", "대표이미지URL", "이미지1", "메인이미지", "이미지URL", "기본이미지"],
  detailHtml: ["상세설명", "상품설명", "상세HTML", "상품상세", "상세페이지", "상세이미지", "상품상세설명"],
  shippingFee: ["기본배송비", "배송비", "반품/교환 배송비"],
  vatType: ["부가세", "과세구분", "부가세유형", "부가세여부"],
  originCode: ["원산지코드", "원산지 지역코드"],
  originDirect: ["원산지 직접입력", "원산지직접입력", "원산지", "원산지명"],
};

export const naverHeaders = [
  "판매자 상품코드", "카테고리코드", "상품명", "상품상태", "판매가", "부가세", "재고수량", "대표이미지", "추가이미지", "상세설명",
  "원산지코드", "복수원산지여부", "원산지 직접입력", "미성년자 구매", "배송방법", "택배사코드", "배송비유형", "기본배송비",
  "반품배송비", "교환배송비", "별도설치비", "A/S 전화번호", "A/S 안내", "구매평 노출여부", "알림받기 동의 고객 전용 여부",
];

export const esmHeaders = [
  "노출\n사이트", "A ID", "G ID", "상품명", "A프로모션 문구", "G프로모션 문구", "G 영문", "G 중문", "카테고리 템플릿 코드", "카테고리 코드",
  "A 노출코드", "G 노출코드", "판매기간", "A 판매가", "G 판매가", "A 할인유형", "A 할인가", "G 할인유형", "G 할인가", "A 재고", "G 재고",
  "옵션\n타입", "옵션명", "옵션\n입력값", "기본이미지", "추가이미지", "상품상세설명", "배송정보 \n템플릿 코드", "배송방법", "출하지 코드",
  "배송정책번호", "반품/교환\n주소 코드", "A 발송정책", "G 발송정책", "택배사\n코드", "반품/교환\n배송비", "상품군\n코드",
  "상품고시정보\n템플릿코드", "인증타입", "인증품목선택", "인증코드", "인증타입", "인증품목선택", "인증코드", "병행수입여부", "인증타입",
  "인증품목선택", "인증코드", "병행수입여부", "인증타입", "승인/신고번호", "원산지\n상품타입", "원산지\n지역타입", "원산지\n지역코드",
  "복수\n원산지여부", "사은품/덤 \n템플릿 코드", "사은품", "덤", "소비기한", "제조일자", "청소년구매\n불가여부", "부가세여부", "선물하기상품",
];

type CategoryPreset = {
  label: string;
  pattern: RegExp;
  esmCategoryCode: string;
  gmarketExposureCode: string;
  auctionExposureCode: string;
  productGroupCode: string;
  noticeTemplateCode: string;
};

// 성원님이 제공한 실제 ESM 전용 엑셀 21개 상품에서 확인한 분류 조합입니다.
// 입력 의미는 ESM → G마켓 → 옥션 순서이며, 공식 ESM 출력 열에는 옥션(A) → G마켓(G) 순서로 배치합니다.
const categoryPresets: CategoryPreset[] = [
  {
    label: "메이크업 스패출러",
    pattern: /스패출러|스파츌라|스패츌러|파운데이션\s*믹싱|메이크업\s*스푼/i,
    esmCategoryCode: "00090004000100150000",
    gmarketExposureCode: "100000005200001764300009393",
    auctionExposureCode: "18650400",
    productGroupCode: "35",
    noticeTemplateCode: "239479",
  },
  {
    label: "파우치·이너백·소품정리",
    pattern: /파우치|이너백|백인백|내부가방|화장품정리|메이크업\s*박스|브러쉬\s*보관함|브러쉬파우치|보석함|쥬얼리\s*정리|미니백|고데기\s*보관|고데기\s*파우치|케이블\s*정리|생리대\s*수납/i,
    esmCategoryCode: "00090004000100150000",
    gmarketExposureCode: "100000049200000787300009543",
    auctionExposureCode: "30110400",
    productGroupCode: "5",
    noticeTemplateCode: "239479",
  },
];

export function cleanKey(value: unknown): string {
  return String(value ?? "").replace(/[\s\n\r_\-()\[\]\/]/g, "").toLowerCase();
}

export function normalize(value: unknown): string {
  return String(value ?? "").trim();
}

export function digits(value: unknown): string {
  return normalize(value).replace(/\D/g, "");
}

export function numberValue(value: unknown): number {
  const number = Number(String(value ?? "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

export function pickExact(row: Row, name: string): unknown {
  const target = cleanKey(name);
  const found = Object.entries(row).find(([key]) => cleanKey(key) === target);
  return found?.[1] ?? "";
}

export function pick(row: Row, names: string[]): unknown {
  for (const name of names) {
    const value = pickExact(row, name);
    if (normalize(value)) return value;
  }
  return "";
}

function detectCategoryPreset(name: string): CategoryPreset | undefined {
  return categoryPresets.find((preset) => preset.pattern.test(name));
}

function detectGroup(name: string): { group: string; productGroupCode: string } {
  const text = name.toLowerCase();
  if (/티셔츠|셔츠|바지|원피스|의류/.test(text)) return { group: "의류 확인 필요", productGroupCode: "1" };
  if (/신발|구두|운동화|샌들/.test(text)) return { group: "신발 확인 필요", productGroupCode: "2" };
  if (/가방|백팩|파우치/.test(text)) return { group: "가방·파우치 확인 필요", productGroupCode: "3" };
  if (/만두|과자|라면|커피|음료|빵|떡/.test(text)) return { group: "가공식품 확인 필요", productGroupCode: "21" };
  return { group: "미분류 제품군", productGroupCode: "35" };
}

export function hasCompleteEsmCategory(product: Product): boolean {
  return Boolean(product.categoryCode && product.gmarketExposureCode && product.auctionExposureCode);
}

export function hasCompleteNotice(product: Product): boolean {
  return Boolean(product.productGroupCode && product.noticeTemplateCode);
}

export function hasCompleteOrigin(product: Product): boolean {
  return Boolean(product.originProductType && product.originRegionType && product.originRegionCode && product.multipleOrigins);
}

export function hasCompleteShipping(product: Product): boolean {
  return Boolean(
    product.departureCode
      && product.shippingPolicyNumber
      && product.returnAddressCode
      && product.auctionShippingPolicy
      && product.gmarketShippingPolicy
      && product.courierCode,
  );
}

export function getProductIssues(product: Product): string[] {
  const issues: string[] = [];
  if (!product.productName) issues.push("상품명");
  if (!product.categoryCode) issues.push("ESM 카테고리");
  if (!product.gmarketExposureCode) issues.push("G마켓 카테고리");
  if (!product.auctionExposureCode) issues.push("옥션 카테고리");
  if (!product.productGroupCode) issues.push("상품군 코드");
  if (!product.noticeTemplateCode) issues.push("고시정보 코드");
  if (!product.mainImage) issues.push("대표이미지");
  if (!product.detailHtml) issues.push("상세설명");
  if (!hasCompleteOrigin(product)) issues.push("원산지");
  if (!hasCompleteShipping(product)) issues.push("배송정책");
  if (!(product.finalPrice > 0)) issues.push("판매가");
  return issues;
}

export function makeProducts(rows: Row[]): Product[] {
  const products: Product[] = [];
  rows.forEach((raw, index) => {
    const productName = normalize(pick(raw, aliases.productName));
    const basePrice = numberValue(pick(raw, aliases.basePrice));
    if (!productName || basePrice <= 0) return;

    const preset = detectCategoryPreset(productName);
    const detected = detectGroup(productName);
    const sourceCategoryCode = digits(pickExact(raw, "카테고리코드"));
    const sourceAuctionCode = digits(pickExact(raw, "A 노출코드"));
    const sourceGmarketCode = digits(pickExact(raw, "G 노출코드"));
    const categoryCode = sourceCategoryCode || preset?.esmCategoryCode || "";
    const gmarketExposureCode = sourceGmarketCode || preset?.gmarketExposureCode || "";
    const auctionExposureCode = sourceAuctionCode || preset?.auctionExposureCode || "";
    const categoryComplete = Boolean(categoryCode && gmarketExposureCode && auctionExposureCode);
    const hadSourceCategory = Boolean(sourceCategoryCode || sourceAuctionCode || sourceGmarketCode);
    const categoryGroup = categoryComplete
      ? hadSourceCategory
        ? `기존분류 · ${categoryCode}`
        : `자동분류 · ${preset?.label ?? categoryCode}`
      : preset
        ? `부분 자동분류 · ${preset.label}`
        : detected.group;
    const shippingGroup = ["출하지 코드", "배송정책번호", "반품/교환 주소 코드", "A 발송정책", "G 발송정책"]
      .map((key) => normalize(pickExact(raw, key)))
      .filter(Boolean)
      .join(" / ") || "배송정보 미입력";

    products.push({
      id: `P${index + 1}`,
      raw,
      productName,
      sellerCode: normalize(pick(raw, aliases.sellerCode)) || `P${index + 1}`,
      basePrice,
      finalPrice: basePrice,
      stock: numberValue(pick(raw, aliases.stock)) || 99999,
      mainImage: normalize(pick(raw, aliases.mainImage)),
      additionalImage: normalize(pickExact(raw, "추가이미지")) || normalize(pickExact(raw, "이미지2")),
      detailHtml: normalize(pick(raw, aliases.detailHtml)),
      shippingFee: numberValue(pick(raw, aliases.shippingFee)),
      vatType: normalize(pick(raw, aliases.vatType)),
      originDirect: normalize(pick(raw, aliases.originDirect)),
      categoryGroup,
      categoryCode,
      auctionExposureCode,
      gmarketExposureCode,
      productGroupCode: digits(pickExact(raw, "상품군 코드")) || preset?.productGroupCode || detected.productGroupCode,
      noticeTemplateCode: digits(pickExact(raw, "상품고시정보 템플릿코드")) || preset?.noticeTemplateCode || "",
      originProductType: normalize(pickExact(raw, "원산지 상품타입")) || "해당없음",
      originRegionType: normalize(pickExact(raw, "원산지 지역타입")) || "알수없음",
      originRegionCode: digits(pick(raw, aliases.originCode)),
      multipleOrigins: normalize(pickExact(raw, "복수 원산지여부")) || "N",
      shippingGroup,
      departureCode: digits(pickExact(raw, "출하지 코드")),
      shippingPolicyNumber: digits(pickExact(raw, "배송정책번호")),
      returnAddressCode: digits(pickExact(raw, "반품/교환 주소 코드")),
      auctionShippingPolicy: digits(pickExact(raw, "A 발송정책")),
      gmarketShippingPolicy: digits(pickExact(raw, "G 발송정책")),
      courierCode: digits(pickExact(raw, "택배사 코드")) || "10013",
      returnShippingFee: numberValue(pickExact(raw, "반품/교환 배송비")) || 2500,
    });
  });
  return products;
}

export function groupProducts(products: Product[], step: StepId): ProductGroup[] {
  const map = new Map<string, Product[]>();
  for (const product of products) {
    const key = step === "category"
      ? product.categoryGroup
      : step === "notice"
        ? `${product.categoryGroup} / ${product.productGroupCode || "상품군 미입력"} / ${product.noticeTemplateCode || "고시 미입력"}`
        : step === "origin"
          ? product.originRegionCode || "원산지 미입력"
          : step === "shipping"
            ? product.shippingGroup
            : "전체";
    const current = map.get(key);
    if (current) current.push(product);
    else map.set(key, [product]);
  }
  return [...map.entries()].map(([name, items]) => ({ name, items }));
}

export function makeDownloadPages(products: Product[]): DownloadPage[] {
  const byCategory = new Map<string, Product[]>();
  for (const product of products) {
    const category = product.categoryCode || "미분류";
    const current = byCategory.get(category);
    if (current) current.push(product);
    else byCategory.set(category, [product]);
  }
  const pages: DownloadPage[] = [];
  for (const [category, items] of byCategory) {
    for (let index = 0; index < items.length; index += 500) {
      pages.push({
        key: `${category}-${index}`,
        category,
        page: Math.floor(index / 500) + 1,
        items: items.slice(index, index + 500),
      });
    }
  }
  return pages;
}

export function validateProducts(products: Product[]): string[] {
  const counts = new Map<string, number>();
  for (const product of products) {
    for (const issue of getProductIssues(product)) {
      counts.set(issue, (counts.get(issue) ?? 0) + 1);
    }
  }
  return [...counts.entries()].map(([issue, count]) => `${issue} 미입력 ${count}개`);
}

export function makeEsmRows(products: Product[], auctionId: string, gmarketId: string): unknown[][] {
  const top: unknown[][] = [
    ["▶가이드 바로가기", null, " ※ 문서버전 : NEW 2.0", null, "필독▶ 일반배송 전용 파일입니다. 상품정보는 8행부터 입력됩니다.", ...Array(59).fill(null)],
    [null, "상품기본정보", ...Array(26).fill(null), "배송정보", ...Array(8).fill(null), "상품고시정보", ...Array(25).fill(null)],
    [null, ...esmHeaders],
    [null, ...esmHeaders.map(() => "")],
    [null, ...esmHeaders.map(() => "")],
    [null, ...esmHeaders.map(() => "")],
    [null, ...esmHeaders.map(() => "")],
  ];

  const body = products.map((product) => {
    const row = Array(64).fill("");
    row[1] = "옥션/G마켓";
    row[2] = auctionId;
    row[3] = gmarketId;
    row[4] = product.productName;
    row[10] = product.categoryCode;
    row[11] = product.auctionExposureCode;
    row[12] = product.gmarketExposureCode;
    row[13] = "무제한";
    row[14] = product.finalPrice;
    row[15] = product.finalPrice;
    row[20] = product.stock;
    row[21] = product.stock;
    row[22] = "미사용";
    row[25] = product.mainImage;
    row[26] = product.additionalImage;
    row[27] = product.detailHtml;
    row[29] = "일반택배";
    row[30] = product.departureCode;
    row[31] = product.shippingPolicyNumber;
    row[32] = product.returnAddressCode;
    row[33] = product.auctionShippingPolicy;
    row[34] = product.gmarketShippingPolicy;
    row[35] = product.courierCode;
    row[36] = product.returnShippingFee;
    row[37] = product.productGroupCode;
    row[38] = product.noticeTemplateCode;
    row[39] = "인증대상아님";
    row[42] = "인증대상아님";
    row[45] = "해당사항없음";
    row[46] = "인증대상아님";
    row[49] = "해당사항없음";
    row[50] = "인증대상아님";
    row[52] = product.originProductType;
    row[53] = product.originRegionType;
    row[54] = product.originRegionCode;
    row[55] = product.multipleOrigins;
    row[61] = "구매가능";
    row[62] = product.vatType.includes("면세") ? "면세상품" : "과세상품";
    row[63] = "가능";
    return row;
  });
  return [...top, ...body];
}
