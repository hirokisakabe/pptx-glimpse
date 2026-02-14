import { NextRequest, NextResponse } from "next/server";
import { convertPptxToSvg } from "pptx-glimpse";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof Blob)) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    const slides = await convertPptxToSvg(uint8Array);

    return NextResponse.json({ slides });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return NextResponse.json({ error: `Conversion failed: ${message}` }, { status: 500 });
  }
}
