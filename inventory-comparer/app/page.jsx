"use client";

import { useEffect, useState, useRef } from "react";
import { useRouter } from "next/navigation";
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
  {
    key: "continue-deny",
    href: "/continue-deny",
    icon: "🚦",
    title: "Continue/Deny Checker",
    description: "Flag SKUs with available stock but a Deny inventory policy in Cosmetics.lk",
    color: "#e53935",
  },
  {
    key: "cosmetics-stock",
    href: "/cosmetics-stock-comparer",
    icon: "🏪",
    title: "Stock Comparer",
    description: "Find Cosmetics.lk products with zero stock and check if other shops have them",
    color: "#00897b",
  },
];

export default function Home() {
  const router = useRouter();
  const [glitch, setGlitch] = useState(false);
  const bufRef = useRef("");

  useEffect(() => {
    const onKey = (e) => {
      // Only append printable single chars
      if (e.key.length === 1) {
        bufRef.current = (bufRef.current + e.key.toLowerCase()).slice(-8);
        if (bufRef.current.endsWith("games")) {
          bufRef.current = "";
          setGlitch(true);
          setTimeout(() => router.push("/games"), 1050);
        }
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [router]);

  return (
    <div style={glitch ? { animation: "g-shake 0.5s ease-in-out" } : {}}>
      {glitch && (
        <div className="fixed inset-0 z-50 pointer-events-none overflow-hidden">
          {/* Dark takeover */}
          <div className="absolute inset-0" style={{ background: "#06060f", animation: "g-darken 0.3s ease-out forwards" }} />

          {/* RGB chromatic split — red */}
          <div className="absolute inset-0" style={{ background: "rgba(255,0,60,0.28)", mixBlendMode: "screen", animation: "g-rgb-r 0.75s ease-in-out forwards" }} />
          {/* RGB chromatic split — blue */}
          <div className="absolute inset-0" style={{ background: "rgba(0,120,255,0.28)", mixBlendMode: "screen", animation: "g-rgb-b 0.75s ease-in-out forwards" }} />

          {/* Scanline sweep beam */}
          <div className="absolute left-0 right-0 h-[3px]" style={{ background: "rgba(255,255,255,0.55)", boxShadow: "0 0 18px rgba(255,255,255,0.7), 0 0 40px rgba(74,109,255,0.5)", animation: "g-scanline 0.55s 0.05s linear forwards" }} />

          {/* Glitch bars */}
          <div className="absolute left-0 right-0 h-7" style={{ top: "27%", background: "rgba(74,109,255,0.2)", animation: "g-bar 0.55s 0.08s ease-out forwards" }} />
          <div className="absolute left-0 right-0 h-3" style={{ top: "54%", background: "rgba(224,64,251,0.18)", animation: "g-bar 0.45s 0.16s ease-out forwards" }} />
          <div className="absolute left-0 right-0 h-2" style={{ top: "73%", background: "rgba(0,230,118,0.15)", animation: "g-bar 0.4s 0.24s ease-out forwards" }} />

          {/* Centre arcade reveal */}
          <div className="absolute inset-0 flex flex-col items-center justify-center" style={{ animation: "g-text-in 0.65s 0.22s ease-out both" }}>
            <div className="text-[80px] leading-none mb-4">🕹️</div>
            <h2
              className="text-[64px] font-black tracking-[0.22em] leading-none"
              style={{
                background: "linear-gradient(90deg,#4a6dff,#e040fb,#00e676,#ffe000)",
                WebkitBackgroundClip: "text",
                WebkitTextFillColor: "transparent",
                filter: "drop-shadow(0 0 24px rgba(74,109,255,0.9)) drop-shadow(0 0 48px rgba(224,64,251,0.6))",
              }}
            >
              ARCADE
            </h2>
            <p className="text-xs tracking-[0.55em] mt-5 font-mono" style={{ color: "#4a6dff", opacity: 0.85 }}>
              ENTERING SECRET ROOM...
            </p>
          </div>

          {/* Scanlines texture */}
          <div className="absolute inset-0" style={{ backgroundImage: "repeating-linear-gradient(0deg,transparent,transparent 2px,rgba(0,0,0,0.18) 2px,rgba(0,0,0,0.18) 4px)", animation: "g-darken 0.3s ease-out forwards" }} />
        </div>
      )}
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
    </div>
  );
}
