import type { Product, Row } from "./types";

export type CategoryReference = {
  label: string;
  keywords: string[];
  categoryCode: string;
  auctionExposureCode: string;
  gmarketExposureCode: string;
  productGroupCode: string;
  noticeTemplateCode: string;
};

const aliases = {
  label: ["카테고리명", "카테고리", "분류명", "세분류", "소분류", "최종카테고리", "상품분류", "카테고리경로"],
  categoryCode: ["카테고리코드", "ESM카테고리코드", "ESM 카테고리 코드", "ESM코드"],
  auctionExposureCode: ["A노출코드", "A 노출코드", "옥션노출코드", "옥션 노출코드"],
  gmarketExposureCode: ["G노출코드", "G 노출코드", "G마켓노출코드", "G마켓 노출코드"],
  productGroupCode: ["상품군코드", "상품군 코드"],
  noticeTemplateCode: ["상품고시정보템플릿코드", "상품고시정보 템플릿코드", "고시정보코드", "고시 템플릿코드"],
};

function clean(value: unknown): string {
  return String(value ?? "").trim();
}

function cleanKey(value: unknown): string {
  return clean(value).replace(/[\s\n\r_\-()\[\]\/·>]/g, "").toLowerCase();
}

function digits(value: unknown): string {
  return clean(value).replace(/\D/g, "");
}

function pick(row: Row, names: string[]): unknown {
  const entries = Object.entries(row);
  for (const name of names) {
    const target = cleanKey(name);
    const found = entries.find(([key]) => cleanKey(key) === target);
    if (found && clean(found[1])) return found[1];
  }
  return "";
}

const stopWords = new Set([
  "정품", "국내", "해외", "무료배송", "당일배송", "신상", "인기", "추천", "대용량", "고급", "세트", "묶음", "옵션",
  "색상", "랜덤", "남성", "여성", "공용", "미니", "휴대용", "가정용", "업소용", "1개", "2개", "3개",
]);

export function normalizeCategoryText(value: unknown): string {
  return clean(value)
    .toLowerCase()
    .replace(/\[[^\]]*\]|\([^)]*\)/g, " ")
    .replace(/[0-9]+(?:ml|g|kg|cm|mm|개|입|매|p|pcs)?/gi, " ")
    .replace(/[^0-9a-z가-힣]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function tokens(value: unknown): string[] {
  return [...new Set(
    normalizeCategoryText(value)
      .split(" ")
      .filter((token) => token.length >= 2 && !stopWords.has(token)),
  )];
}

export function makeCategoryReferences(rows: Row[]): CategoryReference[] {
  const references: CategoryReference[] = [];
  const seen = new Set<string>();

  for (const row of rows) {
    const label = clean(pick(row, aliases.label));
    const categoryCode = digits(pick(row, aliases.categoryCode));
    const auctionExposureCode = digits(pick(row, aliases.auctionExposureCode));
    const gmarketExposureCode = digits(pick(row, aliases.gmarketExposureCode));
    if (!label || !categoryCode || !auctionExposureCode || !gmarketExposureCode) continue;

    const key = `${categoryCode}|${auctionExposureCode}|${gmarketExposureCode}`;
    const keywordValues = Object.values(row)
      .map(clean)
      .filter((value) => value && value.length <= 120)
      .flatMap(tokens);
    const keywords = [...new Set([...tokens(label), ...keywordValues])];

    if (seen.has(`${key}|${normalizeCategoryText(label)}`)) continue;
    seen.add(`${key}|${normalizeCategoryText(label)}`);
    references.push({
      label,
      keywords,
      categoryCode,
      auctionExposureCode,
      gmarketExposureCode,
      productGroupCode: digits(pick(row, aliases.productGroupCode)),
      noticeTemplateCode: digits(pick(row, aliases.noticeTemplateCode)),
    });
  }

  return references;
}

function scoreReference(productName: string, reference: CategoryReference): number {
  const productText = normalizeCategoryText(productName);
  const labelText = normalizeCategoryText(reference.label);
  if (!productText || !labelText) return 0;
  if (productText === labelText) return 1000;
  if (productText.includes(labelText)) return 850 + Math.min(100, labelText.length);

  const productTokens = new Set(tokens(productText));
  const labelTokens = tokens(labelText);
  const keywordMatches = reference.keywords.filter((keyword) => productTokens.has(keyword));
  const labelMatches = labelTokens.filter((keyword) => productTokens.has(keyword));
  const longest = [...keywordMatches, ...labelMatches].reduce((max, token) => Math.max(max, token.length), 0);

  return labelMatches.length * 90 + keywordMatches.length * 20 + longest * 4;
}

export function applyCategoryReferences(
  products: Product[],
  references: CategoryReference[],
): { products: Product[]; matched: number; pending: number } {
  let matched = 0;

  const next = products.map((product) => {
    if (product.categoryCode && product.auctionExposureCode && product.gmarketExposureCode) return product;

    const ranked = references
      .map((reference) => ({ reference, score: scoreReference(product.productName, reference) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score);

    const best = ranked[0];
    const second = ranked[1];
    const confident = Boolean(best && best.score >= 120 && (!second || best.score - second.score >= 25));
    if (!best || !confident) return { ...product, categoryGroup: `자동분류 확인 필요 · ${product.categoryGroup}` };

    matched += 1;
    return {
      ...product,
      categoryGroup: `카테고리 엑셀 자동분류 · ${best.reference.label}`,
      categoryCode: product.categoryCode || best.reference.categoryCode,
      auctionExposureCode: product.auctionExposureCode || best.reference.auctionExposureCode,
      gmarketExposureCode: product.gmarketExposureCode || best.reference.gmarketExposureCode,
      productGroupCode: product.productGroupCode || best.reference.productGroupCode,
      noticeTemplateCode: product.noticeTemplateCode || best.reference.noticeTemplateCode,
    };
  });

  return {
    products: next,
    matched,
    pending: next.filter((product) => !(product.categoryCode && product.auctionExposureCode && product.gmarketExposureCode)).length,
  };
}
