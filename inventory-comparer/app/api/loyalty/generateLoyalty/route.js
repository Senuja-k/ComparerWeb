import { NextResponse } from "next/server";
import { generateLoyaltyReport } from "@/lib/loyaltyComparerLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();

    const referenceEntry = formData.get("referenceFile");
    const locationEntries = formData.getAll("locationsFiles");

    if (!referenceEntry) {
      return NextResponse.json({ error: "Reference file is required" }, { status: 400 });
    }
    if (!locationEntries || locationEntries.length === 0) {
      return NextResponse.json({ error: "At least one location file is required" }, { status: 400 });
    }

    const referenceFile = {
      name: referenceEntry.name,
      buffer: Buffer.from(await referenceEntry.arrayBuffer()),
    };

    const locationFiles = await Promise.all(
      locationEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const reportBuffer = await generateLoyaltyReport(referenceFile, locationFiles);

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="Loyalty_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("Loyalty Comparer error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
