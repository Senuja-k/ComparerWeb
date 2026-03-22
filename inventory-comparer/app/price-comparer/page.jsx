"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function PriceComparerPage() {
  const [referenceFiles, setReferenceFiles] = useState([]);
  const [locationFiles, setLocationFiles] = useState([]);
  const [status, setStatus] = useState("Ready to process files");
  const [progress, setProgress] = useState(0);

  const canGenerate = referenceFiles.length > 0 && locationFiles.length > 0;

  async function handleGenerate() {
    const formData = new FormData();
    formData.append("referenceFile", referenceFiles[0]);
    locationFiles.forEach((f) => formData.append("locationFiles", f));

    setStatus("Processing price comparison... Please wait.");
    setProgress(0);

    try {
      const res = await fetch("/api/price/generatePrice", {
        method: "POST",
        body: formData,
      });

      const contentType = res.headers.get("Content-Type");
      if (contentType && (contentType.includes("application/json") || contentType.includes("text/plain"))) {
        const text = await res.text();
        throw new Error(text || "Failed to generate report");
      }

      const blob = await res.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "Price_Report.xlsx";
      link.click();

      setProgress(100);
      setStatus("✅ Price report generated successfully!");
    } catch (err) {
      setProgress(0);
      setStatus("❌ Error occurred: " + err.message);
    }
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center p-8">
      <Link href="/" className="self-start text-[#28a745] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">💰 Price Comparer</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drag your <b className="text-[#28a745]">Reference Price File</b> on the left and{" "}
        <b className="text-[#ff6b6b]">Location Files</b> on the right.
        <br />
        Compare price discrepancies across different inventory locations.
      </p>

      <div className="flex gap-6 w-full max-w-[1100px] mb-10">
        <DropZone
          title="Reference Price File"
          icon="🔑"
          description="Drop the primary Excel file here (.xlsx, .xls)"
          files={referenceFiles}
          onFilesChange={setReferenceFiles}
          multiple={false}
          accentColor="#28a745"
        />
        <DropZone
          title="Location Files"
          icon="🗺️"
          description="Drop multiple location Excel files here (.xlsx, .xls)"
          files={locationFiles}
          onFilesChange={setLocationFiles}
          accentColor="#28a745"
        />
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#28a745] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#218838] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        📈 Generate Price Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#28a745] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
