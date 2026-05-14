import { NextResponse } from "next/server";
import { generateSkuReport } from "@/lib/skuComparerLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();

    const locationEntries = formData.getAll("locationFiles");
    const unlistedEntries = formData.getAll("unlistedFiles");
    const ogfRulesChecked = formData.get("ogfRulesChecked") === "true";

    if (!locationEntries || locationEntries.length === 0) {
      return NextResponse.json({ error: "At least one location file is required" }, { status: 400 });
    }

    const locationFiles = await Promise.all(
      locationEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const unlistedFiles = await Promise.all(
      unlistedEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const reportBuffer = await generateSkuReport(locationFiles, unlistedFiles, ogfRulesChecked);

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="SKU_Comparison_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("SKU Comparer error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
