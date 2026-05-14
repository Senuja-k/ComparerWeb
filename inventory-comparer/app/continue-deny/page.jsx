"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function ContinueDenyPage() {
  const [cosmeticsFile, setCosmeticsFile] = useState([]);
  const [supplementFile, setSupplementFile] = useState([]);
  const [status, setStatus] = useState("Ready to process files");
  const [progress, setProgress] = useState(0);

  const canGenerate = cosmeticsFile.length > 0 && supplementFile.length > 0;

  async function handleGenerate() {
    const formData = new FormData();
    formData.append("cosmeticsFile", cosmeticsFile[0]);
    formData.append("supplementFile", supplementFile[0]);

    setStatus("Processing... Please wait.");
    setProgress(30);

    try {
      const res = await fetch("/api/continue-deny/generate", {
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
      link.download = "Continue_Deny_Report.xlsx";
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
      <Link href="/" className="self-start text-[#e53935] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">🚦 Continue/Deny Checker</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drop the <b className="text-[#e53935]">Cosmetics.lk</b> file on the left and the{" "}
        <b className="text-[#f57c00]">SupplementVault</b> file on the right.
        <br />
        SKUs with available stock but a <b>Deny</b> inventory policy will be flagged.
      </p>

      <div className="flex gap-6 w-full max-w-[1100px] mb-8">
        <DropZone
          title="Cosmetics.lk File"
          icon="💄"
          description="Drop the Cosmetics.lk file here (.xlsx, .xls, .csv)"
          files={cosmeticsFile}
          onFilesChange={setCosmeticsFile}
          multiple={false}
          accept=".xlsx,.xls,.csv"
          accentColor="#e53935"
        />
        <DropZone
          title="SupplementVault File"
          icon="💊"
          description="Drop the SupplementVault file here (.xlsx, .xls, .csv)"
          files={supplementFile}
          onFilesChange={setSupplementFile}
          multiple={false}
          accept=".xlsx,.xls,.csv"
          accentColor="#f57c00"
        />
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#e53935] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#c62828] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        🚀 Generate Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#e53935] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
