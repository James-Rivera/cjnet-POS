import type { Metadata } from "next";
import { Albert_Sans } from "next/font/google";
import "./globals.css";

const albertSans = Albert_Sans({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700"],
  variable: "--font-albert-sans",
});

export const metadata: Metadata = {
  title: {
    default: "CJNET POS",
    template: "%s | CJNET POS",
  },
  description: "Standalone cashier system for CJNET computer shop services.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className={albertSans.variable}>{children}</body>
    </html>
  );
}
