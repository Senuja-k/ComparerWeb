"use client";

import { useState } from "react";
import Link from "next/link";
import DropZone from "../../components/DropZone";

export default function CreatedDatePage() {
  const [orderFiles, setOrderFiles] = useState([]);
  const [couponFiles, setCouponFiles] = useState([]);
  const [targetFiles, setTargetFiles] = useState([]);
  const [daysOnline, setDaysOnline] = useState("");
  const [daysOutlet, setDaysOutlet] = useState("");
  const [totalDays, setTotalDays] = useState("");
  const [reportDay, setReportDay] = useState("");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [status, setStatus] = useState("Ready to process files");
  const [progress, setProgress] = useState(0);

  const canGenerate =
    orderFiles.length >= 2 &&
    couponFiles.length > 0 &&
    targetFiles.length > 0 &&
    daysOnline &&
    daysOutlet &&
    totalDays &&
    reportDay &&
    startDate &&
    endDate;

  async function handleGenerate() {
    const formData = new FormData();
    orderFiles.forEach((f) => formData.append("orderFiles", f));
    formData.append("couponFile", couponFiles[0]);
    formData.append("targetFile", targetFiles[0]);
    formData.append("daysRemainingOnline", daysOnline);
    formData.append("daysRemainingOutlet", daysOutlet);
    formData.append("totalDays", totalDays);
    formData.append("reportDay", reportDay);
    formData.append("startDate", startDate);
    formData.append("endDate", endDate);

    setStatus("Processing... Please wait.");
    setProgress(30);

    try {
      const res = await fetch("/api/supplement-vault/generate", {
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
      link.download = "SupplementVault_Sales_Report.xlsx";
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
      <Link
        href="/supplement-vault"
        className="self-start text-[#7c4dff] hover:underline text-sm mb-4"
      >
        ← Back to SupplementVault
      </Link>
      <h1 className="text-3xl font-bold text-[#1e1e1e] mb-1">
        💊 SupplementVault.lk — Created Date
      </h1>
      <p className="text-center text-[#787878] max-w-[700px] mb-6">
        Upload Order Reports, Merchant Coupon Codes, and the Target Table to generate your
        monthly sales report.
      </p>

      {/* Order Reports */}
      <div className="w-full max-w-[1100px]">
        <p className="text-xs font-semibold text-[#555] uppercase tracking-wide mb-2">
          Order Reports
        </p>
      </div>
      <div className="flex gap-5 w-full max-w-[1100px] mb-4">
        <DropZone
          title="Order Report Files"
          icon="📦"
          description="Drop the Origins and SupplementVault order report Excel files here (.xlsx, .xls, .csv)"
          files={orderFiles}
          onFilesChange={setOrderFiles}
          accept=".xlsx,.xls,.csv"
          accentColor="#e040fb"
        />
      </div>

      {/* Coupon & Target */}
      <div className="w-full max-w-[1100px]">
        <p className="text-xs font-semibold text-[#555] uppercase tracking-wide mb-2">
          Merchant Coupon Codes &amp; Target Table
        </p>
      </div>
      <div className="flex gap-5 w-full max-w-[1100px] mb-4">
        <div className="max-w-[350px] flex-1">
          <DropZone
            title="Merchant Coupon Codes"
            icon="🏷️"
            description="Drop the Coupon Code file here (.xlsx, .xls, .csv)"
            files={couponFiles}
            onFilesChange={setCouponFiles}
            multiple={false}
            accept=".xlsx,.xls,.csv"
            accentColor="#e040fb"
          />
        </div>
        <DropZone
          title="Target Table"
          icon="🎯"
          description="Drop the Target Excel file containing the monthly target table (.xlsx, .xls)"
          files={targetFiles}
          onFilesChange={setTargetFiles}
          multiple={false}
          accentColor="#e040fb"
        />
      </div>

      {/* Parameters */}
      <div className="w-full max-w-[1100px]">
        <p className="text-xs font-semibold text-[#555] uppercase tracking-wide mb-2">
          Report Parameters
        </p>
      </div>
      <div className="flex gap-4 w-full max-w-[1100px] mb-4 flex-wrap">
        <div className="flex flex-col flex-1 min-w-[200px]">
          <label className="text-[13px] font-semibold text-[#555] mb-1">Start Date (Created at)</label>
          <input
            type="date"
            className="px-3 py-2.5 border border-[#dee2e6] rounded-lg text-sm outline-none focus:border-[#e040fb]"
            value={startDate}
            onChange={(e) => setStartDate(e.target.value)}
          />
        </div>
        <div className="flex flex-col flex-1 min-w-[200px]">
          <label className="text-[13px] font-semibold text-[#555] mb-1">End Date (Created at)</label>
          <input
            type="date"
            className="px-3 py-2.5 border border-[#dee2e6] rounded-lg text-sm outline-none focus:border-[#e040fb]"
            value={endDate}
            onChange={(e) => setEndDate(e.target.value)}
          />
        </div>
      </div>
      <div className="flex gap-4 w-full max-w-[1100px] mb-5 flex-wrap">
        {[
          { label: "Days Remaining (Online Merchants)", value: daysOnline, setter: setDaysOnline, ph: "e.g. 12" },
          { label: "Days Remaining (Outlet Merchants)", value: daysOutlet, setter: setDaysOutlet, ph: "e.g. 12" },
          { label: "Total Days in Month", value: totalDays, setter: setTotalDays, ph: "e.g. 31" },
          { label: "Report Day (Day count)", value: reportDay, setter: setReportDay, ph: "e.g. 19" },
        ].map((f) => (
          <div key={f.label} className="flex flex-col flex-1 min-w-[200px]">
            <label className="text-[13px] font-semibold text-[#555] mb-1">{f.label}</label>
            <input
              type="number"
              min={f.label.includes("Total") || f.label.includes("Report") ? 1 : 0}
              placeholder={f.ph}
              className="px-3 py-2.5 border border-[#dee2e6] rounded-lg text-sm outline-none focus:border-[#e040fb]"
              value={f.value}
              onChange={(e) => f.setter(e.target.value)}
            />
          </div>
        ))}
      </div>

      <button
        disabled={!canGenerate}
        onClick={handleGenerate}
        className="bg-[#e040fb] text-white font-bold text-base border-none rounded-[20px] px-10 py-4 cursor-pointer transition-colors hover:bg-[#c030dd] disabled:bg-[#c8c8c8] disabled:cursor-not-allowed"
      >
        🚀 Generate Sales Report
      </button>

      <div className="w-full max-w-[1100px] mt-5 text-center">
        <div className="w-full h-5 bg-[#e9ecef] rounded-[10px] overflow-hidden mb-2">
          <div
            className="h-full bg-[#e040fb] transition-all duration-300"
            style={{ width: `${progress}%` }}
          />
        </div>
        <p className="text-sm text-[#787878]">{status}</p>
      </div>
    </div>
  );
}
