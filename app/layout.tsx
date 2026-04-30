import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel ETL 工具",
  description: "匯入 Excel、選擇工作表，再執行需要的轉換並下載結果。"
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-Hant">
      <body>{children}</body>
    </html>
  );
}
