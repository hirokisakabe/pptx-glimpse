import { describe, expect, it } from "vitest";

// Import via the actual public surface (`@pptx-glimpse/document`).
import {
  asPt,
  clearParagraphProperties,
  clearTextRunProperties,
  createComputedView,
  findParagraphBySourceHandle,
  findTextRunBySourceHandle,
  readPptx,
  replaceParagraphPlainText,
  replaceTextRunPlainText,
  setParagraphProperties,
  setTextRunProperties,
  writePptx,
} from "../index.js";
import {
  buildMultipleTextEditFixture,
  buildNumericLikeTextFixture,
  buildSlotHandleTextEditFixture,
  buildTextEditFixture,
  buildTextEditFixtureFromSlide,
  decoder,
  findShapeByName,
  firstParagraph,
  firstRun,
  firstShape,
  getEntry,
  shapeAt,
} from "./write-pptx.test-helpers.js";

describe("writePptx - one plain text-run edit", () => {
  it("keeps numeric-like text strings in the source and computed view", () => {
    const source = readPptx(buildNumericLikeTextFixture());
    const shape = findShapeByName(source, "Numeric Text");
    const computed = createComputedView(source);
    const computedShape = computed.slides[0]?.elements.find(
      (element) => element.kind === "shape" && element.sourceNode.name === "Numeric Text",
    );

    expect(shape.textBody?.paragraphs[0]?.runs.map((run) => run.text)).toEqual([
      "007",
      "1e5",
      "12.50",
    ]);
    expect(
      computedShape?.kind === "shape"
        ? computedShape.textBody?.paragraphs[0]?.runs.map((run) => run.text)
        : undefined,
    ).toEqual(["007", "1e5", "12.50"]);
  });

  it("Existing text run can be identified with stable source handle", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shape:10:p:0:r:0",
      orderingSlot: 0,
    });
    expect(findTextRunBySourceHandle(source, run.handle!)).toBe(run);
  });

  it("Apply plain text replacement to PptxSourceModel source and reflect in PPTX after write", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const run = firstRun(source);

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited text");
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(run.text).toBe("Original");
    expect(firstRun(edited).text).toBe("Edited text");
    expect(editedRun.text).toBe("Edited text");
    expect(firstParagraph(reread).runs[1].text).toBe(" Keep ");
  });

  it("does not transform unrelated numeric-like runs when writing a dirty slide", () => {
    const source = readPptx(buildNumericLikeTextFixture());
    const editRun = findShapeByName(source, "Edit Me").textBody?.paragraphs[0]?.runs[0];
    if (editRun?.handle === undefined) throw new Error("edit run handle not found");

    const edited = replaceTextRunPlainText(source, editRun.handle, "Dirty");
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(slideXml).toContain("<a:t>007</a:t>");
    expect(slideXml).toContain("<a:t>1e5</a:t>");
    expect(slideXml).toContain("<a:t>12.50</a:t>");
    expect(
      findShapeByName(reread, "Numeric Text").textBody?.paragraphs[0]?.runs.map((run) => run.text),
    ).toEqual(["007", "1e5", "12.50"]);
    expect(findShapeByName(reread, "Edit Me").textBody?.paragraphs[0]?.runs[0]?.text).toBe("Dirty");
  });

  it("Writes dirty slide XML with a single XML declaration", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const slideXml = decoder.decode(getEntry(writePptx(edited), "ppt/slides/slide1.xml"));

    expect(slideXml.match(/<\?xml/g)).toHaveLength(1);
  });

  it("Preserving run / paragraph formatting and unrelated package material", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const edited = replaceTextRunPlainText(source, firstRun(source).handle!, "Edited text");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const run = firstRun(reread);
    const paragraph = firstParagraph(reread);

    expect(paragraph.properties).toEqual({ align: "center" });
    expect(run.properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(run.rawSidecars?.map((sidecar) => sidecar.node.name) ?? []).not.toContain("a:ea");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Runs without shape ids can also be reflected in dirty slide XML using shapeSlot handle.", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const run = firstRun(source);

    expect(run.handle).toMatchObject({
      partPath: "ppt/slides/slide1.xml",
      nodeId: "text:shapeSlot:1:p:0:r:0",
      orderingSlot: 0,
    });

    const edited = replaceTextRunPlainText(source, run.handle!, "Edited via slot");
    expect(firstRun(readPptx(writePptx(edited))).text).toBe("Edited via slot");
  });
});

describe("writePptx - multiple text edits", () => {
  it("Applies multiple text-run replacements across different shapes and paragraphs in one write.", () => {
    const source = readPptx(buildMultipleTextEditFixture());
    const firstShapeSecondParagraphRun = shapeAt(source, 0).textBody!.paragraphs[1].runs[0];
    const secondShapeRun = shapeAt(source, 1).textBody!.paragraphs[0].runs[0];

    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, firstShapeSecondParagraphRun.handle!, "Edited paragraph"),
      secondShapeRun.handle!,
      "Edited other shape",
    );
    const reread = readPptx(writePptx(edited));

    expect(shapeAt(reread, 0).textBody!.paragraphs[0].runs[0].text).toBe("First paragraph");
    expect(shapeAt(reread, 0).textBody!.paragraphs[1].runs[0].text).toBe("Edited paragraph");
    expect(shapeAt(reread, 1).textBody!.paragraphs[0].runs[0].text).toBe("Edited other shape");
  });

  it("Rejects conflicting duplicate edits for the same text run.", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);
    const edited = replaceTextRunPlainText(
      replaceTextRunPlainText(source, run.handle!, "First edit"),
      run.handle!,
      "Second edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run edits/);
  });

  it("Rejects conflicting text-run and paragraph replacements for the same paragraph.", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      replaceTextRunPlainText(source, paragraph.runs[0].handle!, "Run edit"),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run and paragraph edits/);
  });
});

describe("writePptx - text run property edits", () => {
  it("Sets all supported text run properties and persists them after write/read", () => {
    const source = readPptx(buildTextEditFixture());
    const run = firstRun(source);

    const edited = setTextRunProperties(source, run.handle!, {
      bold: false,
      italic: false,
      underline: true,
      fontSize: asPt(32),
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    const reread = readPptx(writePptx(edited));
    const editedRun = firstRun(reread);

    expect(firstRun(edited).properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00aa44" },
      typeface: "Liberation Sans",
    });
    expect(editedRun.properties).toMatchObject({
      bold: false,
      italic: false,
      underline: true,
      fontSize: 32,
      color: { kind: "srgb", hex: "00AA44" },
      typeface: "Liberation Sans",
      typefaceEa: "Noto Sans JP",
    });
  });

  it("Clears supported text run properties without removing unrelated rPr attributes or children", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="70" name="Decorated"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr lang="en-US" b="1" i="1" u="sng" sz="2400" spc="120">` +
          `<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill>` +
          `<a:latin typeface="Aptos"/><a:effectLst/><a:ea typeface="Noto Sans JP"/>` +
          `</a:rPr><a:t>Original</a:t></a:r>` +
          `<a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:rPr><a:t>Untouched</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    const edited = clearTextRunProperties(source, firstRun(source).handle!, [
      "bold",
      "italic",
      "underline",
      "fontSize",
      "color",
      "typeface",
    ]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    const properties = firstRun(reread).properties;
    expect(properties).toMatchObject({ typefaceEa: "Noto Sans JP" });
    expect(properties?.bold).toBeUndefined();
    expect(properties?.italic).toBeUndefined();
    expect(properties?.underline).toBeUndefined();
    expect(properties?.fontSize).toBeUndefined();
    expect(properties?.color).toBeUndefined();
    expect(properties?.typeface).toBeUndefined();
    expect(slideXml).toContain('lang="en-US"');
    expect(slideXml).toContain('spc="120"');
    expect(slideXml).toContain("<a:effectLst");
    expect(slideXml).toContain('<a:ea typeface="Noto Sans JP"');
    expect(firstParagraph(reread).runs[1].properties).toMatchObject({
      bold: true,
      color: { kind: "srgb", hex: "00FF00" },
    });
  });

  it("Creates rPr when setting properties on a run without existing properties", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = setTextRunProperties(source, firstRun(source).handle!, {
      bold: true,
      fontSize: asPt(18),
      color: { kind: "srgb", hex: "112233" },
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:rPr")).toBeLessThan(slideXml.indexOf("<a:t>Original</a:t>"));
    expect(firstRun(readPptx(output)).properties).toMatchObject({
      bold: true,
      fontSize: 18,
      color: { kind: "srgb", hex: "112233" },
    });
  });

  it("Does not dirty a run when clearing properties that are already absent", () => {
    const source = readPptx(buildSlotHandleTextEditFixture());
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["bold"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(edited).toBe(source);
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Removes an empty latin run property element when clearing typeface", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="71" name="Typeface"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:r><a:rPr><a:latin typeface="Aptos"/></a:rPr><a:t>Original</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const edited = clearTextRunProperties(source, firstRun(source).handle!, ["typeface"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml).not.toContain("<a:latin");
    expect(slideXml).not.toContain("<a:rPr");
    expect(firstRun(readPptx(output)).properties).toBeUndefined();
  });

  it("Rejects invalid direct text run property helper input", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;

    expect(() => setTextRunProperties(source, handle, { fontSize: asPt(0) })).toThrow(
      /fontSize must be a finite positive pt value/,
    );
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      setTextRunProperties(source, handle, { strikethrough: true }),
    ).toThrow(/unsupported text run property 'strikethrough'/);
    expect(() =>
      // @ts-expect-error exercises runtime validation for JS callers.
      clearTextRunProperties(source, handle, ["strikethrough"]),
    ).toThrow(/unsupported text run property 'strikethrough'/);
  });

  it("Rejects no-op text run property edits constructed directly in an edit journal", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = {
      ...source,
      edits: [
        {
          kind: "updateTextRunProperties",
          handle: firstRun(source).handle!,
        },
      ],
    } satisfies typeof source;

    expect(() => writePptx(edited)).toThrow(/must set or clear at least one property/);
  });

  it("Applies text and property edits to the same run in one write", () => {
    const source = readPptx(buildTextEditFixture());
    const handle = firstRun(source).handle!;
    const edited = setTextRunProperties(
      replaceTextRunPlainText(source, handle, "Edited property text"),
      handle,
      { underline: true, color: { kind: "srgb", hex: "336699" } },
    );
    const reread = readPptx(writePptx(edited));

    expect(firstRun(reread).text).toBe("Edited property text");
    expect(firstRun(reread).properties).toMatchObject({
      underline: true,
      color: { kind: "srgb", hex: "336699" },
    });
  });

  it("Rejects conflicting text run property and paragraph replacements for the same paragraph", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);
    const edited = replaceParagraphPlainText(
      setTextRunProperties(source, paragraph.runs[0].handle!, { bold: false }),
      paragraph.handle!,
      "Paragraph edit",
    );

    expect(() => writePptx(edited)).toThrow(/conflicting text run properties and paragraph edits/);
  });
});

describe("writePptx - paragraph property edits", () => {
  it("Sets paragraph alignment, bullet, and level and persists them after write/read", () => {
    const source = readPptx(buildTextEditFixture());
    const paragraph = firstParagraph(source);

    const edited = setParagraphProperties(source, paragraph.handle!, {
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstParagraph(edited).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(firstParagraph(reread).properties).toMatchObject({
      align: "right",
      level: 2,
      bullet: { type: "char", char: "\u2022" },
    });
    expect(slideXml).toContain('<a:pPr algn="r" lvl="2">');
    expect(slideXml).toContain('<a:buChar char="\u2022"');
  });

  it("Writes explicit buNone when removing bullets and supports auto-number bullets", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="72" name="Paragraph props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p>` +
          `<a:p><a:r><a:t>Numbered</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    const first = firstParagraph(source);
    const second = firstShape(source).textBody!.paragraphs[1];
    const edited = setParagraphProperties(
      setParagraphProperties(source, first.handle!, { bullet: { type: "none" } }),
      second.handle!,
      { level: 1, bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 } },
    );
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstShape(reread).textBody!.paragraphs[0].properties).toMatchObject({
      bullet: { type: "none" },
    });
    expect(firstShape(reread).textBody!.paragraphs[1].properties).toMatchObject({
      level: 1,
      bullet: { type: "autoNum", scheme: "alphaLcParenR", startAt: 3 },
    });
    expect(slideXml).toContain("<a:buNone");
    expect(slideXml).toContain('<a:buAutoNum type="alphaLcParenR" startAt="3"');
  });

  it("Clears only requested paragraph properties and preserves unedited paragraphs", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="73" name="Preserve paragraph props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr algn="ctr" lvl="2"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>Edit</a:t></a:r></a:p>` +
          `<a:p><a:pPr algn="r" lvl="1"><a:buChar char="&#x25E6;"/></a:pPr><a:r><a:t>Keep</a:t></a:r></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const edited = clearParagraphProperties(source, firstParagraph(source).handle!, ["align"]);
    const output = writePptx(edited);
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));
    const reread = readPptx(output);

    expect(firstParagraph(reread).properties).toMatchObject({
      level: 2,
      bullet: { type: "char", char: "\u2022" },
      lineSpacing: { type: "pct", value: 90000 },
    });
    expect(firstParagraph(reread).properties?.align).toBeUndefined();
    expect(firstShape(reread).textBody!.paragraphs[1].properties).toMatchObject({
      align: "right",
      level: 1,
      bullet: { type: "char", char: "\u25E6" },
    });
    expect(slideXml).toContain('<a:lnSpc><a:spcPct val="90000"');
    expect(slideXml).toContain('<a:pPr algn="r" lvl="1"');
  });

  it("Rejects no-op paragraph property edits constructed directly in an edit journal", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = {
      ...source,
      edits: [
        {
          kind: "updateParagraphProperties",
          handle: firstParagraph(source).handle!,
        },
      ],
    } satisfies typeof source;

    expect(() => writePptx(edited)).toThrow(/must set or clear at least one property/);
  });
});

describe("writePptx - paragraph text replacement", () => {
  it("Normalizes a multi-run paragraph to one run using the first run properties.", () => {
    const input = buildTextEditFixture();
    const source = readPptx(input);
    const paragraph = firstParagraph(source);

    expect(findParagraphBySourceHandle(source, paragraph.handle!)).toBe(paragraph);

    const edited = replaceParagraphPlainText(source, paragraph.handle!, "Paragraph replacement");
    const output = writePptx(edited);
    const reread = readPptx(output);
    const editedParagraph = firstParagraph(reread);

    expect(firstParagraph(edited).runs).toHaveLength(1);
    expect(editedParagraph.runs).toHaveLength(1);
    expect(editedParagraph.runs[0].text).toBe("Paragraph replacement");
    expect(editedParagraph.runs[0].properties).toMatchObject({
      bold: true,
      italic: true,
      fontSize: 24,
      typeface: "Aptos",
      typefaceEa: "Noto Sans JP",
      color: { kind: "srgb", hex: "FF0000" },
    });
    expect(decoder.decode(getEntry(output, "ppt/slides/slide1.xml"))).not.toContain(" Keep ");
    expect(getEntry(output, "docProps/custom.xml")).toEqual(getEntry(input, "docProps/custom.xml"));
    expect(getEntry(output, "ppt/media/image1.png")).toEqual(
      getEntry(input, "ppt/media/image1.png"),
    );
  });

  it("Round-trips paragraph replacement text with significant surrounding whitespace.", () => {
    const source = readPptx(buildTextEditFixture());
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, " Trimmed ");
    const reread = readPptx(writePptx(edited));

    expect(firstParagraph(reread).runs).toHaveLength(1);
    expect(firstRun(reread).text).toBe(" Trimmed ");
  });

  it("Keeps replacement run before endParaRPr.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="60" name="End para props"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p><a:pPr algn="ctr"/><a:r><a:t>Before</a:t></a:r><a:endParaRPr lang="ja-JP"/></a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );
    const output = writePptx(
      replaceParagraphPlainText(source, firstParagraph(source).handle!, "After"),
    );
    const slideXml = decoder.decode(getEntry(output, "ppt/slides/slide1.xml"));

    expect(slideXml.indexOf("<a:r>")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:endParaRPr")).toBeGreaterThan(-1);
    expect(slideXml.indexOf("<a:r>")).toBeLessThan(slideXml.indexOf("<a:endParaRPr"));
    expect(firstRun(readPptx(output)).text).toBe("After");
  });

  it("Rejects paragraph replacement for interleaved bullet paragraphs split by the reader.", () => {
    const source = readPptx(
      buildTextEditFixtureFromSlide(
        `<p:sp><p:nvSpPr><p:cNvPr id="61" name="Interleaved bullets"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
          `<p:spPr><a:prstGeom prst="rect"/></p:spPr>` +
          `<p:txBody><a:bodyPr/><a:lstStyle/>` +
          `<a:p>` +
          `<a:pPr><a:buChar char="&#x2022;"/></a:pPr><a:r><a:t>One</a:t></a:r><a:br/>` +
          `<a:pPr><a:buChar char="&#x25E6;"/></a:pPr><a:r><a:t>Two</a:t></a:r>` +
          `</a:p>` +
          `</p:txBody></p:sp>`,
      ),
    );

    expect(firstShape(source).textBody!.paragraphs).toHaveLength(2);
    const edited = replaceParagraphPlainText(source, firstParagraph(source).handle!, "After");

    expect(() => writePptx(edited)).toThrow(/interleaved bullet paragraph/);
  });
});
