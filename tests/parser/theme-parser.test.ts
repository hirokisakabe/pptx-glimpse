import { describe, it, expect, vi } from "vitest";
import { parseTheme } from "../../src/parser/theme-parser.js";

const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

describe("parseTheme", () => {
  const theme = parseTheme(themeXml);

  it("parses color scheme", () => {
    expect(theme.colorScheme.dk1).toBe("#000000");
    expect(theme.colorScheme.lt1).toBe("#FFFFFF");
    expect(theme.colorScheme.accent1).toBe("#4472C4");
    expect(theme.colorScheme.accent2).toBe("#ED7D31");
    expect(theme.colorScheme.hlink).toBe("#0563C1");
  });

  it("parses font scheme", () => {
    expect(theme.fontScheme.majorFont).toBe("Calibri Light");
    expect(theme.fontScheme.minorFont).toBe("Calibri");
  });

  it("does not warn for valid XML", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    parseTheme(themeXml);
    expect(warnSpy).not.toHaveBeenCalled();
    warnSpy.mockRestore();
  });

  it("warns and returns defaults when theme root is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('Theme: missing root element "theme"'),
    );
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
  });

  it("warns and returns defaults when themeElements is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:fmtScheme name="Office"/>
      </a:theme>
    `;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("Theme: themeElements not found"));
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
  });

  it("warns when colorScheme is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:fontScheme name="Office">
            <a:majorFont><a:latin typeface="Arial"/></a:majorFont>
            <a:minorFont><a:latin typeface="Arial"/></a:minorFont>
          </a:fontScheme>
        </a:themeElements>
      </a:theme>
    `;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("Theme: colorScheme not found"));
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Arial");
    warnSpy.mockRestore();
  });

  it("warns when fontScheme is missing", () => {
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:themeElements>
          <a:clrScheme name="Office">
            <a:dk1><a:srgbClr val="111111"/></a:dk1>
            <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="222222"/></a:dk2>
            <a:lt2><a:srgbClr val="333333"/></a:lt2>
            <a:accent1><a:srgbClr val="444444"/></a:accent1>
            <a:accent2><a:srgbClr val="555555"/></a:accent2>
            <a:accent3><a:srgbClr val="666666"/></a:accent3>
            <a:accent4><a:srgbClr val="777777"/></a:accent4>
            <a:accent5><a:srgbClr val="888888"/></a:accent5>
            <a:accent6><a:srgbClr val="999999"/></a:accent6>
            <a:hlink><a:srgbClr val="AAAAAA"/></a:hlink>
            <a:folHlink><a:srgbClr val="BBBBBB"/></a:folHlink>
          </a:clrScheme>
        </a:themeElements>
      </a:theme>
    `;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("Theme: fontScheme not found"));
    expect(result.colorScheme.dk1).toBe("#111111");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
  });
});
