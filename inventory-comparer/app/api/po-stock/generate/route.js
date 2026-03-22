import { NextResponse } from "next/server";
import { generateReport } from "@/lib/poStockTallyLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();

    const poEntries = formData.getAll("purchaseOrderFiles");
    const saEntries = formData.getAll("stockAdjustmentFiles");
    const excludeIds = formData.getAll("excludeSAIds");

    if (!poEntries || poEntries.length === 0) {
      return NextResponse.json({ error: "At least one purchase order file is required" }, { status: 400 });
    }
    if (!saEntries || saEntries.length === 0) {
      return NextResponse.json({ error: "At least one stock adjustment file is required" }, { status: 400 });
    }

    const poFiles = await Promise.all(
      poEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const saFiles = await Promise.all(
      saEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const reportBuffer = await generateReport(poFiles, saFiles, excludeIds);

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="PO_Stock_Tally_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("PO Stock Tally error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
