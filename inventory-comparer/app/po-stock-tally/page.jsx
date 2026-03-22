"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function POStockTallyPage() {
  const [poFiles, setPoFiles] = useState([]);
  const [saFiles, setSaFiles] = useState([]);
  const [excludeIds, setExcludeIds] = useState([]);
  const [inputValue, setInputValue] = useState("");
  const [status, setStatus] = useState("Ready to process files");
  const [progress, setProgress] = useState(0);

  const canGenerate = poFiles.length > 0 && saFiles.length > 0;

  function addTags(raw) {
    const values = raw.split(",").map((v) => v.trim()).filter(Boolean);
    const newIds = values.filter((v) => !excludeIds.includes(v));
    if (newIds.length > 0) setExcludeIds([...excludeIds, ...newIds]);
    setInputValue("");
  }

  function handleKeyDown(e) {
    if (e.key === "Enter") {
      e.preventDefault();
      addTags(inputValue);
    }
  }

  function handlePaste(e) {
    e.preventDefault();
    addTags(e.clipboardData.getData("text"));
  }

  function removeTag(id) {
    setExcludeIds(excludeIds.filter((i) => i !== id));
  }

  async function handleGenerate() {
    const formData = new FormData();
    poFiles.forEach((f) => formData.append("purchaseOrderFiles", f));
    saFiles.forEach((f) => formData.append("stockAdjustmentFiles", f));
    excludeIds.forEach((id) => formData.append("excludeSAIds", id));

    setStatus("Processing tally... Please wait.");
    setProgress(0);

    try {
      const res = await fetch("/api/po-stock/generate", {
        method: "POST",
        body: formData,
      });
      if (!res.ok) throw new Error("Failed to generate report");

      const blob = await res.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "PO_Stock_Tally_Report.xlsx";
      link.click();

      setProgress(100);
      setStatus("✅ Tally report generated successfully!");
    } catch (err) {
      setProgress(0);
      setStatus("❌ Error: " + err.message);
    }
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center p-8">
      <Link href="/" className="self-start text-[#ffa726] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">📊 PO-Stock Tally</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drag your <b className="text-[#ffa726]">Purchase Order Files</b> on the left and{" "}
        <b className="text-[#4a6dff]">Stock Adjustment Files</b> on the right.
        <br />
        Compare purchase orders with stock levels and identify discrepancies.
      </p>

      {/* Exclude SA IDs */}
      <div className="w-full max-w-[1100px] mb-5">
        <label className="font-bold block mb-1 text-sm">Exclude Stock Adjustment IDs:</label>
        <div className="flex flex-wrap gap-1 border border-[#dee2e6] rounded-lg px-3 py-1.5 bg-white">
          {excludeIds.map((id) => (
            <span
              key={id}
              className="bg-[#ffa726] text-white rounded-xl px-2.5 py-1 text-[13px] flex items-center"
            >
              {id}
              <span
                className="ml-1.5 cursor-pointer font-bold"
                onClick={() => removeTag(id)}
              >
                ✖
              </span>
            </span>
          ))}
          <input
            className="flex-1 border-none outline-none text-sm p-1 min-w-[120px]"
            placeholder="Type SA IDs and press Enter or paste comma-separated IDs"
            value={inputValue}
            onChange={(e) => setInputValue(e.target.value)}
            onKeyDown={handleKeyDown}
            onPaste={handlePaste}
          />
        </div>
      </div>

      <div className="flex gap-6 w-full max-w-[1100px] mb-10">
        <DropZone
          title="Purchase Order Files"
          icon="📋"
          description="Drop multiple purchase order Excel files here (.xlsx, .xls)"
          files={poFiles}
          onFilesChange={setPoFiles}
          accentColor="#ffa726"
        />
        <DropZone
          title="Stock Adjustment Files"
          icon="📦"
          description="Drop multiple stock adjustment Excel files here (.xlsx, .xls)"
          files={saFiles}
          onFilesChange={setSaFiles}
          accentColor="#4a6dff"
        />
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#ffa726] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#ff9800] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        🚀 Generate Tally Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#ffa726] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
