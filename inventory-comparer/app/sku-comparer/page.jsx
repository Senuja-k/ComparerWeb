"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function SKUComparerPage() {
  const [locationFiles, setLocationFiles] = useState([]);
  const [unlistedFiles, setUnlistedFiles] = useState([]);
  const [ogfActive, setOgfActive] = useState(false);
  const [status, setStatus] = useState("Ready to process files");
  const [progress, setProgress] = useState(0);

  const canGenerate = locationFiles.length > 0;

  async function handleGenerate() {
    const formData = new FormData();
    locationFiles.forEach((f) => formData.append("locationFiles", f));
    if (unlistedFiles.length > 0) {
      unlistedFiles.forEach((f) => formData.append("unlistedFiles", f));
    }
    formData.append("ogfRulesChecked", String(ogfActive));

    setStatus("Processing comparison... Please wait.");
    setProgress(0);

    try {
      const res = await fetch("/api/comparer/generate", {
        method: "POST",
        body: formData,
      });
      if (!res.ok) throw new Error("Failed to generate report");

      const blob = await res.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "SKU_Comparison_Report.xlsx";
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
      <Link href="/" className="self-start text-[#4a6dff] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">📊 SKU Comparer</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drag your <b className="text-[#4a6dff]">Location Files</b> on the left and{" "}
        <b className="text-[#28b692]">Unlisted Files</b> on the right.
        <br />
        Compare product SKUs across different locations.
      </p>

      <div className="flex gap-6 w-full max-w-[1100px] mb-6">
        <DropZone
          title="Location Files"
          icon="📦"
          description="Drop multiple location Excel files here (.xlsx, .xls)"
          files={locationFiles}
          onFilesChange={setLocationFiles}
          accentColor="#4a6dff"
        />
        <DropZone
          title="Unlisted Files"
          icon="📝"
          description="Drop multiple unlisted Excel files here (.xlsx, .xls)"
          files={unlistedFiles}
          onFilesChange={setUnlistedFiles}
          accentColor="#28b692"
        />
      </div>

      <label className="mb-5 flex items-center gap-2 cursor-pointer">
        <input
          type="checkbox"
          checked={ogfActive}
          onChange={(e) => setOgfActive(e.target.checked)}
        />
        Activate OGF Rule
      </label>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#4a6dff] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#2f4cff] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        🚀 Generate Comparison Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#4a6dff] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
