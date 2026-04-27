"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function CosmeticsStockComparerPage() {
  const [inventoryFile, setInventoryFile] = useState([]);
  const [status, setStatus] = useState("Ready to process file");
  const [progress, setProgress] = useState(0);

  const canGenerate = inventoryFile.length > 0;

  async function handleGenerate() {
    const formData = new FormData();
    formData.append("inventoryFile", inventoryFile[0]);

    setStatus("Processing... Please wait.");
    setProgress(30);

    try {
      const res = await fetch("/api/cosmetics-stock-comparer/generate", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error(errText || "Failed to generate report");
      }

      const blob = await res.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "Cosmetics_Stock_Report.xlsx";
      link.click();

      setProgress(100);
      setStatus("✅ Report generated successfully!");
    } catch (err) {
      setProgress(0);
      setStatus("❌ Error: " + err.message);
    }
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center p-8">
      <Link href="/" className="self-start text-[#00897b] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">🏪 Cosmetics Stock Comparer</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drop the <b className="text-[#00897b]">Inventory Status</b> file below.
        <br />
        Products at <b>0 or negative</b> stock on Cosmetics.lk will be checked against
        other shops in priority order.
      </p>

      <div className="w-full max-w-[600px] mb-8">
        <DropZone
          title="Inventory Status File"
          icon="📦"
          description="Drop the inventory status Excel file here (.xlsx, .xls)"
          files={inventoryFile}
          onFilesChange={setInventoryFile}
          multiple={false}
          accentColor="#00897b"
        />
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#00897b] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#00695c] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        🚀 Generate Stock Report
      </button>

      <div className="w-full max-w-[600px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#00897b] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
