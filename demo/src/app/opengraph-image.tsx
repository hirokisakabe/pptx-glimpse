import { ImageResponse } from "next/og";

export const alt = "pptx-glimpse – PPTX to SVG Converter";
export const size = { width: 1200, height: 630 };
export const contentType = "image/png";

export default function OgImage() {
  return new ImageResponse(
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        width: "100%",
        height: "100%",
        background: "linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%)",
        color: "#fff",
        fontFamily: "sans-serif",
      }}
    >
      <div style={{ fontSize: 72, fontWeight: 700, marginBottom: 16 }}>pptx-glimpse</div>
      <div style={{ fontSize: 32, color: "#94a3b8" }}>PPTX to SVG Converter</div>
      <div style={{ fontSize: 24, color: "#64748b", marginTop: 24 }}>
        Open-source TypeScript library
      </div>
    </div>,
    { ...size },
  );
}
