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

type PreparedReference = {
  reference: CategoryReference;
  labelText: string;
  labelTokens: string[];
};

function scorePreparedReference(productText: string, productTokens: Set<string>, prepared: PreparedReference): number {
  if (!productText || !prepared.labelText) return 0;
  if (productText === prepared.labelText) return 1000;
  if (productText.includes(prepared.labelText)) return 850 + Math.min(100, prepared.labelText.length);

  let keywordMatches = 0;
  let labelMatches = 0;
  let longest = 0;

  for (const keyword of prepared.reference.keywords) {
    if (!productTokens.has(keyword)) continue;
    keywordMatches += 1;
    if (keyword.length > longest) longest = keyword.length;
  }
  for (const keyword of prepared.labelTokens) {
    if (!productTokens.has(keyword)) continue;
    labelMatches += 1;
    if (keyword.length > longest) longest = keyword.length;
  }

  return labelMatches * 90 + keywordMatches * 20 + longest * 4;
}

export function applyCategoryReferences(
  products: Product[],
  references: CategoryReference[],
): { products: Product[]; matched: number; pending: number } {
  let matched = 0;

  const prepared: PreparedReference[] = references.map((reference) => ({
    reference,
    labelText: normalizeCategoryText(reference.label),
    labelTokens: tokens(reference.label),
  }));

  const keywordIndex = new Map<string, number[]>();
  prepared.forEach((item, index) => {
    const searchTokens = new Set([...item.labelTokens, ...item.reference.keywords]);
    for (const token of searchTokens) {
      const list = keywordIndex.get(token);
      if (list) list.push(index);
      else keywordIndex.set(token, [index]);
    }
  });

  const next = products.map((product) => {
    if (product.categoryCode && product.auctionExposureCode && product.gmarketExposureCode) return product;

    const productText = normalizeCategoryText(product.productName);
    const productTokenList = tokens(productText);
    const productTokenSet = new Set(productTokenList);
    const candidateIds = new Set<number>();

    for (const token of productTokenList) {
      const ids = keywordIndex.get(token);
      if (!ids) continue;
      for (const id of ids) candidateIds.add(id);
    }

    if (!candidateIds.size) {
      return { ...product, categoryGroup: `자동분류 확인 필요 · ${product.categoryGroup}` };
    }

    let best: { prepared: PreparedReference; score: number } | undefined;
    let secondScore = 0;

    for (const id of candidateIds) {
      const item = prepared[id];
      const score = scorePreparedReference(productText, productTokenSet, item);
      if (score <= 0) continue;

      if (!best || score > best.score) {
        secondScore = best?.score ?? secondScore;
        best = { prepared: item, score };
      } else if (score > secondScore) {
        secondScore = score;
      }
    }

    const confident = Boolean(best && best.score >= 120 && (secondScore === 0 || best.score - secondScore >= 25));
    if (!best || !confident) {
      return { ...product, categoryGroup: `자동분류 확인 필요 · ${product.categoryGroup}` };
    }

    matched += 1;
    const selected = best.prepared.reference;
    return {
      ...product,
      categoryGroup: `카테고리 엑셀 자동분류 · ${selected.label}`,
      categoryCode: product.categoryCode || selected.categoryCode,
      auctionExposureCode: product.auctionExposureCode || selected.auctionExposureCode,
      gmarketExposureCode: product.gmarketExposureCode || selected.gmarketExposureCode,
      productGroupCode: product.productGroupCode || selected.productGroupCode,
      noticeTemplateCode: product.noticeTemplateCode || selected.noticeTemplateCode,
    };
  });

  return {
    products: next,
    matched,
    pending: next.filter((product) => !(product.categoryCode && product.auctionExposureCode && product.gmarketExposureCode)).length,
  };
}
