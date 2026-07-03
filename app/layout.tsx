import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "薪酬核算系统",
  description: "业务人员月度薪酬自动核算工具",
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
