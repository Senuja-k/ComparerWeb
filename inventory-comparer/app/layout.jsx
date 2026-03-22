import "./globals.css";

export const metadata = {
  title: "Comparer Dashboard",
  description: "Inventory Comparison Tool",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body suppressHydrationWarning className="min-h-screen flex flex-col bg-gradient-to-b from-[#f8fafc] to-[#f0f5ff] font-sans text-[#1e1e1e]">{children}</body>
    </html>
  );
}
