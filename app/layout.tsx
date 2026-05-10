import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "康乃尔薪酬核算系统 V2.0",
  description: "康乃尔业务人员月度薪酬核算 — 基于 2026 年薪酬考核方案 V2.0",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
