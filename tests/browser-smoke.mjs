import assert from "node:assert/strict";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { chromium } from "playwright";
import * as XLSX from "xlsx";

const baseURL = process.env.TEST_BASE_URL || "http://127.0.0.1:3000";
const outputDir = path.resolve("test-results");
await mkdir(outputDir, { recursive: true });

const rows = [
  ["상품 일괄등록 테스트 파일"],
  [],
  ["상품명", "판매자 상품코드", "판매가", "재고수량", "대표이미지", "상세설명", "카테고리코드", "원산지코드", "출하지 코드", "배송정책번호", "반품/교환 주소 코드", "A 발송정책", "G 발송정책", "상품고시정보 템플릿코드"],
];
for (let index = 1; index <= 1200; index += 1) {
  const productName = index <= 400
    ? `테스트 상품 ${index}`
    : index <= 800
      ? `티셔츠 테스트 상품 ${index}`
      : `만두 테스트 상품 ${index}`;
  const categoryCode = index <= 400 ? "50000001" : "";

  rows.push([
    productName,
    `TEST-${index}`,
    1000 + index,
    10,
    `https://example.com/${index}.jpg`,
    `<p>테스트 상세 ${index}</p>`,
    categoryCode,
    "03",
    "100001",
    "200001",
    "300001",
    "400001",
    "500001",
    "239479",
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

  await page.locator('input[type="file"]').setInputFiles(samplePath);
  await page.getByRole("button", { name: "ESM · 옥션/G마켓" }).click();
  await page.getByRole("heading", { name: "옥션·G마켓 설정" }).waitFor({ state: "visible", timeout: 5_000 });
  await page.getByText(/1200개 상품을 읽었습니다/).waitFor({ state: "visible", timeout: 60_000 });

  await page.getByRole("button", { name: "1. 카테고리" }).click();
  await page.getByText("3개 묶음").waitFor({ state: "visible" });

  const categoryField = page.locator("label.apply-field").filter({ hasText: "카테고리 코드" });

  await page.locator(".group-list button").filter({ hasText: "의류 예상" }).click();
  await categoryField.locator("input").fill("50000001");
  await categoryField.getByRole("button", { name: "이 묶음에 적용" }).click();
  await page.locator(".editor-heading h3").filter({ hasText: "카테고리 50000001" }).waitFor({ state: "visible" });
  await page.locator(".editor-heading").getByText("800개 상품").waitFor({ state: "visible" });
  assert.equal(await page.locator(".group-list button.selected strong").textContent(), "카테고리 50000001");

  await page.locator(".group-list button").filter({ hasText: "가공식품 예상" }).click();
  await categoryField.locator("input").fill("50000001");
  await categoryField.getByRole("button", { name: "이 묶음에 적용" }).click();
  await page.getByText("1개 묶음").waitFor({ state: "visible" });
  await page.locator(".editor-heading").getByText("1,200개 상품").waitFor({ state: "visible" });
  assert.equal(await page.locator(".group-list button.selected strong").textContent(), "카테고리 50000001");

  for (const label of ["2. 고시정보", "3. 원산지", "4. 배송정책", "5. 가격", "6. 최종검사·다운로드"]) {
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
  await page.getByText("3개 파일").waitFor({ state: "visible" });

  const esmDownloadPromise = page.waitForEvent("download", { timeout: 60_000 });
  await page.locator(".download-list button").first().click();
  const esmDownload = await esmDownloadPromise;
  const esmPath = path.join(outputDir, esmDownload.suggestedFilename());
  await esmDownload.saveAs(esmPath);
  const esmBook = XLSX.read(await readFile(esmPath));
  const esmMatrix = XLSX.utils.sheet_to_json(esmBook.Sheets[esmBook.SheetNames[0]], { header: 1, defval: "" });
  assert.equal(esmMatrix[7][4], "테스트 상품 1");
  assert.equal(esmMatrix[7][2], "auction-test");
  assert.equal(esmMatrix[7][3], "gmarket-test");

  assert.deepEqual(browserErrors, [], `브라우저 오류가 발생했습니다:\n${browserErrors.join("\n")}`);
  await writeFile(path.join(outputDir, "browser-smoke.txt"), "PASS\n", "utf8");
  console.log("PASS: 그룹 선택 유지, 같은 카테고리 병합, 화면 클릭, 엑셀 업로드와 다운로드를 확인했습니다.");
} catch (error) {
  await page.screenshot({ path: path.join(outputDir, "failure.png"), fullPage: true });
  throw error;
} finally {
  await browser.close();
}
