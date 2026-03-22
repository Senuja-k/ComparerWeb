"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../components/DropZone";

export default function LoyaltyComparerPage() {
  const [referenceFiles, setReferenceFiles] = useState([]);
  const [locationFiles, setLocationFiles] = useState([]);
  const [status, setStatus] = useState("Ready to process loyalty files");
  const [progress, setProgress] = useState(0);

  const canGenerate = referenceFiles.length > 0 && locationFiles.length > 0;

  async function handleGenerate() {
    const formData = new FormData();
    formData.append("referenceFile", referenceFiles[0]);
    locationFiles.forEach((f) => formData.append("locationsFiles", f));

    setStatus("Processing loyalty comparison... Please wait.");
    setProgress(0);

    try {
      const res = await fetch("/api/loyalty/generateLoyalty", {
        method: "POST",
        body: formData,
      });
      if (!res.ok) throw new Error("Failed to generate report");

      const blob = await res.blob();
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "Loyalty_Report.xlsx";
      link.click();

      setProgress(100);
      setStatus("✅ Loyalty report generated successfully!");
    } catch (err) {
      setProgress(0);
      setStatus("❌ Error occurred: " + err.message);
    }
  }

  return (
    <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center p-8">
      <Link href="/" className="self-start text-[#9333ea] hover:underline text-sm mb-4">
        ← Back to Dashboard
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-2">👑 Loyalty Comparer</h1>
      <p className="text-center text-[#787878] max-w-[650px] mb-10">
        Drag your <b className="text-[#9333ea]">Reference Loyalty File</b> on the left and{" "}
        <b className="text-[#ff9f1c]">Location Loyalty Files</b> on the right.
        <br />
        Compare loyalty programs, points systems, and customer rewards across locations.
      </p>

      <div className="flex gap-6 w-full max-w-[1100px] mb-10">
        <DropZone
          title="Reference Loyalty File"
          icon="👑"
          description="Drop a single .xlsx or .xls file here"
          files={referenceFiles}
          onFilesChange={setReferenceFiles}
          multiple={false}
          accentColor="#9333ea"
        />
        <DropZone
          title="Location Loyalty Files"
          icon="🏪"
          description="Drop multiple .xlsx or .xls files here"
          files={locationFiles}
          onFilesChange={setLocationFiles}
          accentColor="#ff9f1c"
        />
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#9333ea] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#7a1bbf] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        👑 Generate Loyalty Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#9333ea] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
