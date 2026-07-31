"use client";

import { useMemo, useState } from "react";

type Workbook = import("xlsx").WorkBook;
type Worksheet = import("xlsx").WorkSheet;

type ColumnOption = {
  index: number;
  letter: string;
  label: string;
  numericCount: number;
};

type PreviewRow = {
  row: number;
  before: number;
  after: number;
};

const PRICE_WORDS = ["가격", "판매가", "소비자가", "공급가", "원가", "단가", "금액", "price", "cost"];

function readNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const cleaned = value.replace(/,/g, "").replace(/[^0-9.-]/g, "").trim();
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function roundPrice(value: number, unit: number, method: "ceil" | "round" | "floor"): number {
  const safeUnit = Math.max(1, unit);
  if (method === "floor") return Math.floor(value / safeUnit) * safeUnit;
  if (method === "round") return Math.round(value / safeUnit) * safeUnit;
  return Math.ceil(value / safeUnit) * safeUnit;
}

function calculatePrice(
  current: number,
  feeRate: number,
  marginRate: number,
  consumerRatio: number,
  unit: number,
  method: "ceil" | "round" | "floor",
): number {
  const denominator = 1 - feeRate / 100 - marginRate / 100;
  if (denominator <= 0) throw new Error("수수료율과 마진율의 합은 100%보다 작아야 합니다.");
  const ratio = Math.max(0, consumerRatio) / 100;
  return roundPrice((current / denominator) * ratio, unit, method);
}

function getHeaderText(sheet: Worksheet, rowIndex: number, columnIndex: number): string {
  const XLSX = requireXlsx();
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  const cell = sheet[address];
  return String(cell?.v ?? "").trim();
}

let xlsxModule: typeof import("xlsx") | null = null;
function requireXlsx(): typeof import("xlsx") {
  if (!xlsxModule) throw new Error("엑셀 모듈이 아직 준비되지 않았습니다.");
  return xlsxModule;
}

export default function Page() {
  const [workbook, setWorkbook] = useState<Workbook | null>(null);
  const [sourceFile, setSourceFile] = useState<File | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);
  const [feeRate, setFeeRate] = useState(0);
  const [marginRate, setMarginRate] = useState(0);
  const [consumerRatio, setConsumerRatio] = useState(100);
  const [roundUnit, setRoundUnit] = useState(100);
  const [roundMethod, setRoundMethod] = useState<"ceil" | "round" | "floor">("ceil");
  const [status, setStatus] = useState("가격을 변경할 엑셀 파일을 업로드해 주세요.");
  const [busy, setBusy] = useState(false);

  const currentSheet = workbook && sheetName ? workbook.Sheets[sheetName] : null;

  const columns = useMemo<ColumnOption[]>(() => {
    if (!currentSheet || !currentSheet["!ref"]) return [];
    const XLSX = requireXlsx();
    const range = XLSX.utils.decode_range(currentSheet["!ref"]);
    const headerIndex = Math.max(0, headerRow - 1);
    const result: ColumnOption[] = [];

    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const label = getHeaderText(currentSheet, headerIndex, column) || `${XLSX.utils.encode_col(column)}열`;
      let numericCount = 0;
      for (let row = headerIndex + 1; row <= range.e.r; row += 1) {
        const cell = currentSheet[XLSX.utils.encode_cell({ r: row, c: column })];
        if (readNumber(cell?.v) !== null) numericCount += 1;
      }
      if (numericCount > 0) {
        result.push({ index: column, letter: XLSX.utils.encode_col(column), label, numericCount });
      }
    }
    return result;
  }, [currentSheet, headerRow]);

  const preview = useMemo<PreviewRow[]>(() => {
    if (!currentSheet || !currentSheet["!ref"] || selectedColumn === null) return [];
    const XLSX = requireXlsx();
    const range = XLSX.utils.decode_range(currentSheet["!ref"]);
    const rows: PreviewRow[] = [];
    try {
      for (let row = Math.max(range.s.r, headerRow); row <= range.e.r && rows.length < 20; row += 1) {
        const cell = currentSheet[XLSX.utils.encode_cell({ r: row, c: selectedColumn })];
        const before = readNumber(cell?.v);
        if (before === null) continue;
        rows.push({
          row: row + 1,
          before,
          after: calculatePrice(before, feeRate, marginRate, consumerRatio, roundUnit, roundMethod),
        });
      }
    } catch {
      return [];
    }
    return rows;
  }, [currentSheet, selectedColumn, headerRow, feeRate, marginRate, consumerRatio, roundUnit, roundMethod]);

  async function uploadFile(file?: File) {
    if (!file) return;
    setBusy(true);
    setStatus("원본 통합문서를 읽고 있습니다.");

    try {
      const XLSX = await import("xlsx");
      xlsxModule = XLSX;
      const buffer = await file.arrayBuffer();
      const nextWorkbook = XLSX.read(buffer, {
        type: "array",
        cellStyles: true,
        cellDates: true,
        cellFormula: true,
        cellNF: true,
        cellText: true,
        bookVBA: true,
      });

      setWorkbook(nextWorkbook);
      setSourceFile(file);
      setSheetName(nextWorkbook.SheetNames[0] ?? "");
      setHeaderRow(1);
      setSelectedColumn(null);
      setStatus(`${file.name}을 읽었습니다. 가격 열을 선택해 주세요.`);
    } catch (error) {
      setWorkbook(null);
      setSourceFile(null);
      setStatus(error instanceof Error ? `파일 읽기 오류: ${error.message}` : "엑셀 파일을 읽지 못했습니다.");
    } finally {
      setBusy(false);
    }
  }

  function autoSelectPriceColumn() {
    const matched = columns.find((column) =>
      PRICE_WORDS.some((word) => column.label.toLowerCase().includes(word.toLowerCase())),
    );
    const fallback = [...columns].sort((a, b) => b.numericCount - a.numericCount)[0];
    const target = matched ?? fallback;
    if (!target) {
      setStatus("숫자가 입력된 열을 찾지 못했습니다.");
      return;
    }
    setSelectedColumn(target.index);
    setStatus(`${target.letter}열 · ${target.label}을 가격 열로 선택했습니다.`);
  }

  async function applyAndDownload() {
    if (!workbook || !sourceFile || !currentSheet || selectedColumn === null) {
      setStatus("파일과 가격 열을 먼저 선택해 주세요.");
      return;
    }

    setBusy(true);
    setStatus("원본 형식을 유지하면서 선택한 가격 셀만 변경하고 있습니다.");

    try {
      const XLSX = requireXlsx();
      const denominator = 1 - feeRate / 100 - marginRate / 100;
      if (denominator <= 0) throw new Error("수수료율과 마진율의 합은 100%보다 작아야 합니다.");
      if (!currentSheet["!ref"]) throw new Error("선택한 시트에 데이터가 없습니다.");

      const range = XLSX.utils.decode_range(currentSheet["!ref"]);
      let changed = 0;

      for (let row = Math.max(range.s.r, headerRow); row <= range.e.r; row += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: selectedColumn });
        const cell = currentSheet[address];
        if (!cell || cell.f) continue;
        const current = readNumber(cell.v);
        if (current === null || current < 0) continue;

        cell.v = calculatePrice(current, feeRate, marginRate, consumerRatio, roundUnit, roundMethod);
        cell.t = "n";
        changed += 1;
      }

      if (!changed) throw new Error("선택한 열에서 변경할 숫자 가격을 찾지 못했습니다.");

      const originalName = sourceFile.name;
      const dot = originalName.lastIndexOf(".");
      const baseName = dot > 0 ? originalName.slice(0, dot) : originalName;
      const extension = dot > 0 ? originalName.slice(dot + 1).toLowerCase() : "xlsx";
      const supportedType = ["xlsx", "xlsm", "xlsb", "xls"].includes(extension) ? extension : "xlsx";
      const outputName = `${baseName}_가격수정.${supportedType}`;

      XLSX.writeFile(workbook, outputName, {
        bookType: supportedType as "xlsx" | "xlsm" | "xlsb" | "xls",
        cellStyles: true,
        bookVBA: true,
      });
      setStatus(`${changed.toLocaleString()}개 가격 셀을 변경해 다운로드했습니다.`);
    } catch (error) {
      setStatus(error instanceof Error ? `가격 변경 오류: ${error.message}` : "가격을 변경하지 못했습니다.");
    } finally {
      setBusy(false);
    }
  }

  const selected = columns.find((column) => column.index === selectedColumn);

  return (
    <main style={{ minHeight: "100vh", background: "#f6f7fb", padding: "32px 18px", color: "#172033" }}>
      <section style={{ maxWidth: 980, margin: "0 auto" }}>
        <header style={{ marginBottom: 24 }}>
          <span style={{ display: "inline-block", padding: "6px 10px", borderRadius: 999, background: "#e9efff", fontWeight: 700, fontSize: 13 }}>가격 일괄수정</span>
          <h1 style={{ margin: "12px 0 8px", fontSize: 34 }}>엑셀 원본 형식 그대로, 가격만 변경</h1>
          <p style={{ margin: 0, color: "#596579", lineHeight: 1.7 }}>새 양식으로 변환하지 않습니다. 업로드한 통합문서에서 선택한 가격 열의 숫자 셀만 수수료율·마진율·소비자가 비율에 따라 일괄 수정합니다.</p>
        </header>

        <section style={cardStyle}>
          <label style={uploadStyle}>
            <strong style={{ fontSize: 18 }}>엑셀 파일 업로드</strong>
            <span style={{ color: "#69758a" }}>.xlsx, .xlsm, .xlsb, .xls 파일</span>
            <input type="file" accept=".xlsx,.xlsm,.xlsb,.xls" disabled={busy} onChange={(event) => void uploadFile(event.currentTarget.files?.[0])} />
          </label>

          {workbook && (
            <div style={{ display: "grid", gap: 18, marginTop: 24 }}>
              <div style={gridStyle}>
                <Field label="시트 선택">
                  <select value={sheetName} onChange={(event) => { setSheetName(event.target.value); setSelectedColumn(null); }} style={inputStyle}>
                    {workbook.SheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
                  </select>
                </Field>

                <Field label="제목 행 번호">
                  <input type="number" min={1} value={headerRow} onChange={(event) => { setHeaderRow(Math.max(1, Number(event.target.value) || 1)); setSelectedColumn(null); }} style={inputStyle} />
                </Field>

                <Field label="가격 열 선택">
                  <select value={selectedColumn ?? ""} onChange={(event) => setSelectedColumn(event.target.value === "" ? null : Number(event.target.value))} style={inputStyle}>
                    <option value="">가격 열을 선택하세요</option>
                    {columns.map((column) => (
                      <option key={column.index} value={column.index}>{column.letter}열 · {column.label} · 숫자 {column.numericCount}개</option>
                    ))}
                  </select>
                </Field>

                <div style={{ display: "flex", alignItems: "end" }}>
                  <button type="button" onClick={autoSelectPriceColumn} style={secondaryButtonStyle}>가격 열 자동 찾기</button>
                </div>
              </div>

              <div style={gridStyle}>
                <NumberField label="수수료율 (%)" value={feeRate} onChange={setFeeRate} />
                <NumberField label="마진율 (%)" value={marginRate} onChange={setMarginRate} />
                <NumberField label="소비자가 비율 (%)" value={consumerRatio} onChange={setConsumerRatio} min={0} />
                <Field label="가격 단위">
                  <select value={roundUnit} onChange={(event) => setRoundUnit(Number(event.target.value))} style={inputStyle}>
                    <option value={1}>1원</option>
                    <option value={10}>10원</option>
                    <option value={100}>100원</option>
                    <option value={1000}>1,000원</option>
                  </select>
                </Field>
                <Field label="단위 처리">
                  <select value={roundMethod} onChange={(event) => setRoundMethod(event.target.value as "ceil" | "round" | "floor")} style={inputStyle}>
                    <option value="ceil">올림</option>
                    <option value="round">반올림</option>
                    <option value="floor">내림</option>
                  </select>
                </Field>
              </div>

              <div style={{ padding: 16, borderRadius: 12, background: "#f1f4fa", lineHeight: 1.7 }}>
                <strong>계산 방식</strong><br />
                변경가격 = 현재가격 ÷ (1 - 수수료율 - 마진율) × 소비자가 비율
                <br /><span style={{ color: "#657086" }}>소비자가 비율을 적용하지 않으려면 100%로 두면 됩니다.</span>
              </div>

              {selected && (
                <section>
                  <h2 style={{ fontSize: 20, marginBottom: 10 }}>변경 미리보기 · {selected.letter}열 {selected.label}</h2>
                  <div style={{ overflowX: "auto", border: "1px solid #dfe4ee", borderRadius: 12 }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 480 }}>
                      <thead><tr><th style={thStyle}>행</th><th style={thStyle}>변경 전</th><th style={thStyle}>변경 후</th></tr></thead>
                      <tbody>
                        {preview.map((item) => (
                          <tr key={item.row}><td style={tdStyle}>{item.row}</td><td style={tdStyle}>{item.before.toLocaleString()}원</td><td style={tdStyle}>{item.after.toLocaleString()}원</td></tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              <button type="button" disabled={busy || selectedColumn === null} onClick={() => void applyAndDownload()} style={{ ...primaryButtonStyle, opacity: busy || selectedColumn === null ? 0.5 : 1 }}>
                원본 형식으로 가격 수정 파일 다운로드
              </button>
            </div>
          )}

          <div role="status" aria-live="polite" style={{ marginTop: 20, padding: 14, borderRadius: 10, background: busy ? "#fff5d9" : "#eef7f0", fontWeight: 600 }}>
            {busy ? "처리 중 · " : ""}{status}
          </div>
          <p style={{ marginBottom: 0, color: "#6d7789", fontSize: 13 }}>파일은 브라우저 안에서만 처리하며 서버에 저장하지 않습니다. 수식 셀은 안전을 위해 변경하지 않습니다.</p>
        </section>
      </section>
    </main>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <label style={{ display: "grid", gap: 7 }}><span style={{ fontWeight: 700 }}>{label}</span>{children}</label>;
}

function NumberField({ label, value, onChange, min }: { label: string; value: number; onChange: (value: number) => void; min?: number }) {
  return <Field label={label}><input type="number" min={min} value={value} onChange={(event) => onChange(Number(event.target.value))} style={inputStyle} /></Field>;
}

const cardStyle: React.CSSProperties = { background: "white", border: "1px solid #e1e5ed", borderRadius: 18, padding: 24, boxShadow: "0 12px 35px rgba(34, 48, 74, 0.07)" };
const uploadStyle: React.CSSProperties = { display: "grid", gap: 10, padding: 22, border: "2px dashed #aeb9cb", borderRadius: 14, cursor: "pointer" };
const gridStyle: React.CSSProperties = { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 14 };
const inputStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", minHeight: 44, border: "1px solid #cfd6e2", borderRadius: 9, padding: "9px 11px", fontSize: 15, background: "white" };
const primaryButtonStyle: React.CSSProperties = { minHeight: 50, border: 0, borderRadius: 11, background: "#315efb", color: "white", fontSize: 16, fontWeight: 800, cursor: "pointer" };
const secondaryButtonStyle: React.CSSProperties = { ...primaryButtonStyle, width: "100%", background: "#344054", minHeight: 44, fontSize: 14 };
const thStyle: React.CSSProperties = { textAlign: "left", padding: 12, background: "#f4f6fa", borderBottom: "1px solid #dfe4ee" };
const tdStyle: React.CSSProperties = { padding: 12, borderBottom: "1px solid #edf0f5" };
