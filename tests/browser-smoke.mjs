import assert from "node:assert/strict";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { chromium } from "playwright";
import * as XLSX from "xlsx";

const baseURL = process.env.TEST_BASE_URL || "http://127.0.0.1:3000";
const outputDir = path.resolve("test-results");
await mkdir(outputDir, { recursive: true });

const commonHeaders = [
  "상품명",
  "판매자 상품코드",
  "판매가",
  "재고수량",
  "대표이미지",
  "상세설명",
  "카테고리코드",
  "G 노출코드",
  "A 노출코드",
  "상품군 코드",
  "상품고시정보 템플릿코드",
  "원산지 상품타입",
  "원산지 지역타입",
  "원산지코드",
  "복수 원산지여부",
  "출하지 코드",
  "배송정책번호",
  "반품/교환 주소 코드",
  "A 발송정책",
  "G 발송정책",
  "택배사 코드",
  "반품/교환 배송비",
];

const autoRows = [
  ["자동 카테고리 테스트"],
  [],
  commonHeaders,
  [
    "여행용 화장품 파우치 이너백",
    "AUTO-1",
    12000,
    10,
    "https://example.com/auto-1.jpg",
    "<p>파우치</p>",
    "",
    "",
    "",
    "",
    "",
    "해당없음",
    "알수없음",
    "04",
    "N",
    "23962201",
    "60785",
    "47397781",
    "14",
    "15",
    "10013",
    "20000",
  ],
  [
    "메이크업 스패출러 파운데이션 믹싱 스파츌라",
    "AUTO-2",
    9000,
    10,
    "https://example.com/auto-2.jpg",
    "<p>스패출러</p>",
    "",
    "",
    "",
    "",
    "",
    "해당없음",
    "알수없음",
    "04",
    "N",
    "23962201",
    "60785",
    "47397781",
    "14",
    "15",
    "10013",
    "20000",
  ],
  [
    "분류 규칙에 없는 테스트 제품",
    "AUTO-3",
    15000,
    10,
    "https://example.com/auto-3.jpg",
    "<p>미분류</p>",
    "",
    "",
    "",
    "",
    "",
    "해당없음",
    "알수없음",
    "04",
    "N",
    "23962201",
    "60785",
    "47397781",
    "14",
    "15",
    "10013",
    "20000",
  ],
];
const autoPath = path.join(outputDir, "auto-category-products.xlsx");
const autoBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(autoBook, XLSX.utils.aoa_to_sheet(autoRows), "상품목록");
XLSX.writeFile(autoBook, autoPath);

const rows = [
  ["상품 일괄등록 테스트 파일"],
  [],
  commonHeaders,
];
for (let index = 1; index <= 1200; index += 1) {
  const productName = index <= 400
    ? `기존 분류 테스트 상품 ${index}`
    : index <= 800
      ? `티셔츠 테스트 상품 ${index}`
      : `만두 테스트 상품 ${index}`;
  const categoryCode = index <= 400 ? "50000001" : "";
  const gmarketCode = index <= 400 ? "100000049200000787300009543" : "";
  const auctionCode = index <= 400 ? "30110400" : "";

  rows.push([
    productName,
    `TEST-${index}`,
    1000 + index,
    10,
    `https://example.com/${index}.jpg`,
    `<p>테스트 상세 ${index}</p>`,
    categoryCode,
    gmarketCode,
    auctionCode,
    "5",
    "239479",
    "해당없음",
    "알수없음",
    "04",
    "N",
    "23962201",
    "60785",
    "47397781",
    "14",
    "15",
    "10013",
    "20000",
  ]);
}
const samplePath = path.join(outputDir, "sample-products.xlsx");
const sampleBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(sampleBook, XLSX.utils.aoa_to_sheet(rows), "상품목록");
XLSX.writeFile(sampleBook, samplePath);

const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ acceptDownloads: true });
const browserErrors = [];
page.on("pageerror", (error) => browserErrors.push(`pageerror: ${error.message}`));
page.on("console", (message) => {
  if (message.type() === "error") browserErrors.push(`console: ${message.text()}`);
});

try {
  await page.goto(baseURL, { waitUntil: "networkidle", timeout: 60_000 });
  await page.getByRole("heading", { name: "상품 엑셀 변환" }).waitFor({ state: "visible" });

  await page.getByRole("button", { name: "ESM · 옥션/G마켓" }).click();
  await page.getByRole("heading", { name: "옥션·G마켓 설정" }).waitFor({ state: "visible" });
  await page.getByRole("button", { name: "스마트스토어" }).click();
  await page.getByRole("heading", { name: "스마트스토어 설정" }).waitFor({ state: "visible" });

  const phoneInput = page.locator("label.field").filter({ hasText: "A/S 전화번호" }).locator("input");
  await phoneInput.fill("01012345678");
  assert.equal(await phoneInput.inputValue(), "01012345678");

  // 제공한 ESM 카테고리 파일의 파우치/스패출러 규칙을 자동 적용하고, 미분류만 빨간색으로 남기는지 확인합니다.
  await page.locator('input[type="file"]').setInputFiles(autoPath);
  await page.getByRole("button", { name: "ESM · 옥션/G마켓" }).click();
  await page.getByText(/3개 상품을 ESM 형식으로 옮겼습니다/).waitFor({ state: "visible", timeout: 60_000 });
  await page.getByText(/카테고리 자동·기존 입력 2개 \/ 직접 확인 1개/).waitFor({ state: "visible" });
  await page.getByText(/고시정보 입력 2개 \/ 직접 확인 1개/).waitFor({ state: "visible" });
  await page.locator(".group-list>button.needs-review").waitFor({ state: "visible" });
  const categoryLabels = await page.locator(".editor-fields label.apply-field > span").allTextContents();
  assert.deepEqual(categoryLabels.map((value) => value.replace(" · 수정 필요", "")).slice(0, 3), [
    "1. ESM 카테고리 코드",
    "2. G마켓 노출코드",
    "3. 옥션 노출코드",
  ]);

  await page.getByRole("button", { name: "6. 최종검사·다운로드" }).click();
  await page.locator(".review-table").waitFor({ state: "visible" });
  assert.equal(await page.locator(".review-table tbody tr").count(), 3);
  assert.ok(await page.locator(".review-table .cell-error").count() > 0);

  // 대량 파일에서는 이미 입력된 400개를 제외하고 두 제품군만 카테고리를 수정합니다.
  await page.locator('input[type="file"]').setInputFiles(samplePath);
  await page.getByText(/1200개 상품을 ESM 형식으로 옮겼습니다/).waitFor({ state: "visible", timeout: 60_000 });
  await page.getByText(/카테고리 자동·기존 입력 400개 \/ 직접 확인 800개/).waitFor({ state: "visible" });
  await page.getByRole("button", { name: "1. 카테고리" }).click();
  await page.getByText("2개 묶음").waitFor({ state: "visible" });

  const applyCurrentCategoryTriple = async () => {
    const values = [
      ["1. ESM 카테고리 코드", "50000001"],
      ["2. G마켓 노출코드", "100000049200000787300009543"],
      ["3. 옥션 노출코드", "30110400"],
    ];
    for (const [label, value] of values) {
      const field = page.locator("label.apply-field").filter({ hasText: label });
      const input = field.locator("input");
      await input.fill(value);
      assert.equal(await input.inputValue(), value);
      await field.getByRole("button", { name: "이 묶음에 적용" }).click();
      await page.waitForTimeout(100);
    }
  };

  await page.locator(".group-list button").filter({ hasText: "의류 확인 필요" }).click();
  await page.locator(".editor-heading h3").filter({ hasText: "의류 확인 필요" }).waitFor({ state: "visible" });
  await applyCurrentCategoryTriple();
  await page.getByText(/카테고리 자동·기존 입력 800개 \/ 직접 확인 400개/).waitFor({ state: "visible" });

  await page.locator(".group-list button").filter({ hasText: "가공식품 확인 필요" }).click();
  await page.locator(".editor-heading h3").filter({ hasText: "가공식품 확인 필요" }).waitFor({ state: "visible" });
  await applyCurrentCategoryTriple();
  await page.getByText(/카테고리 자동·기존 입력 1,200개 \/ 직접 확인 0개/).waitFor({ state: "visible" });
  await page.getByText(/모든 상품의 ESM·G마켓·옥션 카테고리가 입력되었습니다/).waitFor({ state: "visible" });

  await page.getByRole("button", { name: "2. 고시정보" }).click();
  await page.getByText(/모든 상품의 상품군 코드와 고시정보 템플릿코드가 입력되었습니다/).waitFor({ state: "visible" });
  for (const label of ["3. 원산지", "4. 배송정책", "5. 가격"]) {
    await page.getByRole("button", { name: label }).click();
  }

  await page.getByRole("button", { name: "스마트스토어" }).click();
  const smartDownloadPromise = page.waitForEvent("download", { timeout: 60_000 });
  await page.getByRole("button", { name: "스마트스토어 엑셀 다운로드" }).click();
  const smartDownload = await smartDownloadPromise;
  const smartPath = path.join(outputDir, smartDownload.suggestedFilename());
  await smartDownload.saveAs(smartPath);
  const smartBook = XLSX.read(await readFile(smartPath));
  const smartMatrix = XLSX.utils.sheet_to_json(smartBook.Sheets[smartBook.SheetNames[0]], { header: 1 });
  assert.equal(smartMatrix[0][2], "상품명");
  assert.equal(smartMatrix.length, 1201);

  await page.getByRole("button", { name: "ESM · 옥션/G마켓" }).click();
  await page.locator("label.field").filter({ hasText: "옥션 판매자 ID" }).locator("input").fill("auction-test");
  await page.locator("label.field").filter({ hasText: "G마켓 판매자 ID" }).locator("input").fill("gmarket-test");
  await page.getByRole("button", { name: "6. 최종검사·다운로드" }).click();
  await page.getByText(/모든 상품의 필수항목이 입력되었습니다/).waitFor({ state: "visible" });
  assert.equal(await page.locator(".review-table tbody tr").count(), 50);
  assert.equal(await page.locator(".review-table .cell-error").count(), 0);

  const esmDownloadPromise = page.waitForEvent("download", { timeout: 60_000 });
  await page.getByRole("button", { name: /전체 1,200개 상품 한 엑셀 다운로드/ }).click();
  const esmDownload = await esmDownloadPromise;
  assert.equal(esmDownload.suggestedFilename(), "ESM_전체상품_1200개.xlsx");
  const esmPath = path.join(outputDir, esmDownload.suggestedFilename());
  await esmDownload.saveAs(esmPath);

  const esmBook = XLSX.read(await readFile(esmPath));
  assert.deepEqual(esmBook.SheetNames, ["NEW 일반상품"]);
  const esmMatrix = XLSX.utils.sheet_to_json(esmBook.Sheets["NEW 일반상품"], { header: 1, defval: "" });
  assert.equal(esmMatrix.length, 1207);
  assert.equal(esmMatrix[7][4], "기존 분류 테스트 상품 1");
  assert.equal(esmMatrix[7][2], "auction-test");
  assert.equal(esmMatrix[7][3], "gmarket-test");
  assert.equal(esmMatrix[7][10], "50000001");
  assert.equal(esmMatrix[7][11], "30110400");
  assert.equal(esmMatrix[7][12], "100000049200000787300009543");
  assert.equal(esmMatrix[1206][4], "만두 테스트 상품 1200");

  // 프로그램이 만든 한 파일을 다시 올려도 전체 1,200개를 읽어야 합니다.
  await page.locator('input[type="file"]').setInputFiles(esmPath);
  await page.getByText(/1200개 상품을 ESM 형식으로 옮겼습니다/).waitFor({ state: "visible", timeout: 60_000 });
  await page.getByText(/카테고리 자동·기존 입력 1,200개 \/ 직접 확인 0개/).waitFor({ state: "visible" });
  await page.locator(".file-summary").getByText(/변환 가능 상품 1,200개/).waitFor({ state: "visible" });

  assert.deepEqual(browserErrors, [], `브라우저 오류가 발생했습니다:\n${browserErrors.join("\n")}`);
  await writeFile(path.join(outputDir, "browser-smoke.txt"), "PASS\n", "utf8");
  console.log("PASS: ESM 형식 자동 이동, 카테고리·고시 미입력 빨간 표시, 최종 뷰어와 한 엑셀 다운로드를 확인했습니다.");
} catch (error) {
  await page.screenshot({ path: path.join(outputDir, "failure.png"), fullPage: true });
  const details = error instanceof Error ? error.stack || error.message : String(error);
  await writeFile(path.join(outputDir, "failure.txt"), `${details}\n`, "utf8");
  throw error;
} finally {
  await browser.close();
}
