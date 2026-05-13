import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Excel Mapper",
  description: "Map source Excel data into a target template with Test Script IQ."
};

export default function RootLayout({
  children
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className="dark">
      <body>{children}</body>
    </html>
  );
}
