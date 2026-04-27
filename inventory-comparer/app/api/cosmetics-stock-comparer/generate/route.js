import { NextResponse } from "next/server";
import { generateCosmeticsStockReport } from "@/lib/cosmeticsStockComparerLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();
    const inventoryEntry = formData.get("inventoryFile");

    if (!inventoryEntry) {
      return NextResponse.json({ error: "Inventory file is required" }, { status: 400 });
    }

    const inventoryFile = {
      name: inventoryEntry.name,
      buffer: Buffer.from(await inventoryEntry.arrayBuffer()),
    };

    const reportBuffer = await generateCosmeticsStockReport(inventoryFile);

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="Cosmetics_Stock_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("Cosmetics Stock Comparer error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
