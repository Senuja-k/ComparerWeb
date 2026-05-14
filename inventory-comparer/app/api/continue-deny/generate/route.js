import { NextResponse } from "next/server";
import { generateContinueDenyReport } from "@/lib/continueDenyLogic";

export async function POST(request) {
  try {
    const formData = await request.formData();

    const cosmeticsEntry = formData.get("cosmeticsFile");
    const supplementEntry = formData.get("supplementFile");

    if (!cosmeticsEntry) {
      return NextResponse.json({ error: "Cosmetics file is required" }, { status: 400 });
    }
    if (!supplementEntry) {
      return NextResponse.json({ error: "Supplement file is required" }, { status: 400 });
    }

    const cosmeticsFile = {
      name: cosmeticsEntry.name,
      buffer: Buffer.from(await cosmeticsEntry.arrayBuffer()),
    };

    const supplementFile = {
      name: supplementEntry.name,
      buffer: Buffer.from(await supplementEntry.arrayBuffer()),
    };

    const reportBuffer = await generateContinueDenyReport(cosmeticsFile, supplementFile);

    return new NextResponse(new Uint8Array(reportBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="Continue_Deny_Report.xlsx"',
      },
    });
  } catch (err) {
    console.error("Continue/Deny error:", err);
    return new NextResponse(err.message || "Internal server error", { status: 500 });
  }
}
