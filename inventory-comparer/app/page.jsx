"use client";

import Link from "next/link";

const tools = [
  {
    key: "sku",
    href: "/sku-comparer",
    icon: "🔍",
    title: "SKU & Product Matcher",
    description: "Match and compare product SKUs across different inventory systems",
    color: "#4a6dff",
  },
  {
    key: "price",
    href: "/price-comparer",
    icon: "💰",
    title: "Price Difference Checker",
    description: "Identify price differences and inconsistencies across platforms",
    color: "#28b692",
  },
  {
    key: "loyalty",
    href: "/loyalty-comparer",
    icon: "👑",
    title: "Loyalty Comparer",
    description: "Compare loyalty status for customers across companies.",
    color: "#ff6b6b",
  },
  {
    key: "po-stock",
    href: "/po-stock-tally",
    icon: "📊",
    title: "PO-Stock Tally",
    description: "Compare purchase orders with stock levels and identify discrepancies",
    color: "#ffa726",
  },
  {
    key: "sales",
    href: "/sales-report",
    icon: "📈",
    title: "Sales Report",
    description: "Generate monthly sales reports and track merchant performance targets",
    color: "#e040fb",
  },
];

export default function Home() {
  return (
    <>
      <div className="pt-10 pb-5 text-center">
        <h1 className="text-4xl font-bold bg-gradient-to-r from-[#4a6dff] to-[#28b692] bg-clip-text text-transparent">
          Comparer Dashboard
        </h1>
        <p className="text-sm text-[#787878] mt-2">
          Select a comparison tool to begin inventory analysis
        </p>
      </div>

      <div className="flex-1 flex justify-center items-center px-5 pb-5">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-5 max-w-[1100px] w-full">
          {tools.map((tool) => (
            <Link
              key={tool.key}
              href={tool.href}
              className="bg-white border-3 rounded-2xl shadow-sm p-5 flex items-center justify-between h-[120px] cursor-pointer transition-all hover:-translate-y-1 hover:shadow-md no-underline"
              style={{ borderColor: tool.color }}
            >
              <div className="max-w-[75%]">
                <h2 className="text-lg font-semibold mb-2 flex items-center gap-2 text-[#1e1e1e]">
                  <span>{tool.icon}</span> {tool.title}
                </h2>
                <p className="text-[#787878] text-sm leading-relaxed">{tool.description}</p>
              </div>
              <div className="text-xl font-bold" style={{ color: tool.color }}>→</div>
            </Link>
          ))}
        </div>
      </div>

      <footer className="p-4 text-center text-[#aaa] text-sm">
        Comparer Dashboard — Inventory Comparison Tool
      </footer>
    </>
  );
}
