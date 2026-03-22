import { NextResponse } from "next/server";
import { generateSupplementVaultReport } from "@/lib/supplementVaultLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();

    const orderEntries = formData.getAll("orderFiles");
    const couponEntry = formData.get("couponFile");
    const targetEntry = formData.get("targetFile");

    const daysRemainingOnline = Number(formData.get("daysRemainingOnline") || 0);
    const daysRemainingOutlet = Number(formData.get("daysRemainingOutlet") || 0);
    const totalDays = Number(formData.get("totalDays") || 0);
    const reportDay = Number(formData.get("reportDay") || 0);

    const fulfilledFrom = formData.get("fulfilledFrom");
    const fulfilledTo = formData.get("fulfilledTo");

    if (!orderEntries || orderEntries.length < 2) {
      return NextResponse.json({ error: "At least 2 order files are required" }, { status: 400 });
    }
    if (!couponEntry) {
      return NextResponse.json({ error: "Coupon file is required" }, { status: 400 });
    }
    if (!targetEntry) {
      return NextResponse.json({ error: "Target file is required" }, { status: 400 });
    }

    const orderFiles = await Promise.all(
      orderEntries.map(async (f) => ({
        name: f.name,
        buffer: Buffer.from(await f.arrayBuffer()),
      }))
    );

    const couponFile = {
      name: couponEntry.name,
      buffer: Buffer.from(await couponEntry.arrayBuffer()),
    };

    const targetFile = {
      name: targetEntry.name,
      buffer: Buffer.from(await targetEntry.arrayBuffer()),
    };

    const reportBuffer = await generateSupplementVaultReport({
      orderFiles,
      couponFile,
      targetFile,
      daysRemainingOnline,
      daysRemainingOutlet,
      totalDays,
      reportDay,
      fulfilledFrom: fulfilledFrom || undefined,
      fulfilledTo: fulfilledTo || undefined,
    });

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="SupplementVault_Sales_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("Supplement Vault error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
