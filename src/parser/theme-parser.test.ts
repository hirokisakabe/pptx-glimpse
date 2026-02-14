import { describe, it, expect, vi } from "vitest";
import { parseTheme } from "./theme-parser.js";
import { initWarningLogger } from "../warning-logger.js";

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
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `<other/>`;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(
      expect.stringContaining('missing root element "theme" in XML'),
    );
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns and returns defaults when themeElements is missing", () => {
    initWarningLogger("debug");
    const warnSpy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const xml = `
      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:fmtScheme name="Office"/>
      </a:theme>
    `;
    const result = parseTheme(xml);

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("themeElements not found"));
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns when colorScheme is missing", () => {
    initWarningLogger("debug");
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

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("colorScheme not found"));
    expect(result.colorScheme.dk1).toBe("#000000");
    expect(result.fontScheme.majorFont).toBe("Arial");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("warns when fontScheme is missing", () => {
    initWarningLogger("debug");
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

    expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining("fontScheme not found"));
    expect(result.colorScheme.dk1).toBe("#111111");
    expect(result.fontScheme.majorFont).toBe("Calibri");
    warnSpy.mockRestore();
    initWarningLogger("off");
  });

  it("parses fmtScheme with fillStyleLst, lnStyleLst, effectStyleLst, bgFillStyleLst", () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
      <a:themeElements>
        <a:clrScheme name="Office">
          <a:dk1><a:srgbClr val="000000"/></a:dk1>
          <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
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
        <a:fmtScheme name="Office">
          <a:fillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
            <a:gradFill rotWithShape="1">
              <a:gsLst>
                <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/></a:schemeClr></a:gs>
                <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="75000"/></a:schemeClr></a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
          </a:fillStyleLst>
          <a:lnStyleLst>
            <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
            <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
            <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
          </a:lnStyleLst>
          <a:effectStyleLst>
            <a:effectStyle><a:effectLst/></a:effectStyle>
            <a:effectStyle>
              <a:effectLst>
                <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                  <a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>
                </a:outerShdw>
              </a:effectLst>
            </a:effectStyle>
          </a:effectStyleLst>
          <a:bgFillStyleLst>
            <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          </a:bgFillStyleLst>
        </a:fmtScheme>
      </a:themeElements>
    </a:theme>`;
    const result = parseTheme(xml);

    expect(result.fmtScheme).toBeDefined();
    const fmt = result.fmtScheme!;

    // fillStyleLst: solidFill + gradFill = 2 fills
    expect(fmt.fillStyles).toHaveLength(2);
    expect(fmt.fillStyles[0].type).toBe("solid");
    expect(fmt.fillStyles[1].type).toBe("gradient");

    // lnStyleLst: 3 lines
    expect(fmt.lnStyles).toHaveLength(3);
    expect(fmt.lnStyles[0].width).toBe(6350);
    expect(fmt.lnStyles[1].width).toBe(12700);
    expect(fmt.lnStyles[2].width).toBe(19050);

    // effectStyleLst: 2 effect styles
    expect(fmt.effectStyles).toHaveLength(2);
    expect(fmt.effectStyles[0]).toBeNull(); // empty effectLst
    expect(fmt.effectStyles[1]).not.toBeNull();
    expect(fmt.effectStyles[1]?.outerShadow).toBeDefined();

    // bgFillStyleLst: 1 fill
    expect(fmt.bgFillStyles).toHaveLength(1);
    expect(fmt.bgFillStyles[0].type).toBe("solid");
  });

  it("returns undefined fmtScheme when empty", () => {
    // The default themeXml has an empty fmtScheme
    expect(theme.fmtScheme).toBeUndefined();
  });
});
