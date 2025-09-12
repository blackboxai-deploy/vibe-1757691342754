import type { Metadata } from "next";
import { Inter } from "next/font/google";
import "./globals.css";

const inter = Inter({ subsets: ["latin"] });

export const metadata: Metadata = {
  title: "Generador de Anexo Práctico - Sociedad de Beneficencia de Puno",
  description: "Sistema para generar archivos Excel con formatos del Reglamento Interno de Negocios",
  keywords: "excel, beneficencia, puno, reglamento, anexo práctico",
  authors: [{ name: "Sociedad de Beneficencia de Puno" }],
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="es">
      <head>
        <link rel="icon" href="/favicon.ico" />
      </head>
      <body className={`${inter.className} antialiased bg-gray-50`}>
        <div className="min-h-screen">
          {children}
        </div>
      </body>
    </html>
  );
}