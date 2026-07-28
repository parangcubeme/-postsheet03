import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "POSTSHEET03 상품 변환기",
  description: "스마트스토어와 ESM 상품 엑셀 변환기",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
