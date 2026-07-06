import {
  buildPptx,
  gridPosition,
  NS,
  outlineXml,
  REL_TYPES,
  savePptx,
  shapeXml,
  slideRelsXml,
  solidFillXml,
  textBodyXmlHelper,
  wrapSlideXml,
} from "../fixture-builder.js";

async function createTextFixture(): Promise<void> {
  // Slide 1: Text formatting
  let id = 2;
  const textShapes1: string[] = [];
  const textTests = [
    { label: "Bold", bold: true },
    { label: "Italic", italic: true },
    { label: "Underline", underline: true },
    { label: "Strike", strikethrough: true },
    { label: "Small (12pt)", fontSize: 12 },
    { label: "Large (36pt)", fontSize: 36 },
    { label: "Red Text", color: "FF0000" },
    { label: "Blue Text", color: "0000FF" },
  ];
  textTests.forEach((t, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const pos = gridPosition(col, row, 4, 2);
    textShapes1.push(
      shapeXml(id++, `text-${t.label}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        textBodyXml: textBodyXmlHelper(t.label, {
          bold: t.bold,
          italic: t.italic,
          underline: t.underline,
          strikethrough: t.strikethrough,
          fontSize: t.fontSize ?? 18,
          color: t.color ?? "333333",
        }),
      }),
    );
  });
  const slide1 = wrapSlideXml(textShapes1.join("\n"));

  // Slide 2: Alignment, line spacing, autofit, wrapping
  id = 2;
  const textShapes2: string[] = [];
  const alignTests = [
    { label: "Left Align", align: "l" },
    { label: "Center", align: "ctr" },
    { label: "Right Align", align: "r" },
    { label: "Top Anchor", align: "ctr", anchor: "t" },
    { label: "Line Spacing 200%", lineSpacing: 200000 },
    {
      label: "AutoFit Text That Should Shrink Down",
      normAutofit: { fontScale: 50000 },
    },
    {
      label: "Fixed Line Spacing 21pt With Text That Wraps Across Lines",
      lineSpacingPts: 2100,
    },
  ];
  const alignRows = Math.ceil(alignTests.length / 3);
  alignTests.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const pos = gridPosition(col, row, 3, alignRows);
    textShapes2.push(
      shapeXml(id++, `text-${t.label}`, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        textBodyXml: textBodyXmlHelper(t.label, {
          fontSize: 18,
          color: "333333",
          align: t.align,
          anchor: t.anchor,
          lineSpacing: t.lineSpacing,
          lineSpacingPts: t.lineSpacingPts,
          normAutofit: t.normAutofit,
        }),
      }),
    );
  });
  const slide2 = wrapSlideXml(textShapes2.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "text.pptx");
}

// --- 4. Transform ---

function bulletParagraphsXml(
  items: {
    text: string;
    bulletXml: string;
    marL?: number;
    indent?: number;
    lvl?: number;
  }[],
  opts?: { anchor?: string },
): string {
  const anchor = opts?.anchor ?? "t";
  const paragraphs = items
    .map((item) => {
      const marL = item.marL ?? 342900;
      const indent = item.indent ?? -342900;
      const lvl = item.lvl ?? 0;
      return `<a:p>
      <a:pPr lvl="${lvl}" marL="${marL}" indent="${indent}">
        ${item.bulletXml}
      </a:pPr>
      <a:r>
        <a:rPr lang="en-US" sz="1400">
          <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
        </a:rPr>
        <a:t>${item.text}</a:t>
      </a:r>
    </a:p>`;
    })
    .join("\n");

  return `<p:txBody>
  <a:bodyPr anchor="${anchor}"/>
  <a:lstStyle/>
  ${paragraphs}
</p:txBody>`;
}

async function createBulletsFixture(): Promise<void> {
  // Slide 1: Bullet characters (buChar)
  let id = 2;
  const shapes1: string[] = [];

  // buChar - standard bullet
  const buCharItems = [
    { text: "First bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
    { text: "Second bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
    { text: "Third bullet item", bulletXml: `<a:buChar char="\u2022"/>` },
  ];
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "buChar-standard", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buCharItems),
    }),
  );

  // buChar - dash bullet
  const buDashItems = [
    { text: "Dash item A", bulletXml: `<a:buChar char="-"/>` },
    { text: "Dash item B", bulletXml: `<a:buChar char="-"/>` },
  ];
  const pos2 = gridPosition(1, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "buChar-dash", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buDashItems),
    }),
  );

  // buAutoNum - arabicPeriod
  const buNumItems = [
    {
      text: "Numbered one",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
    {
      text: "Numbered two",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
    {
      text: "Numbered three",
      bulletXml: `<a:buAutoNum type="arabicPeriod"/>`,
    },
  ];
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "buAutoNum-arabic", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buNumItems),
    }),
  );

  // buAutoNum - alphaLcPeriod
  const buAlphaItems = [
    {
      text: "Alpha item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
    {
      text: "Beta item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
    {
      text: "Gamma item",
      bulletXml: `<a:buAutoNum type="alphaLcPeriod"/>`,
    },
  ];
  const pos4 = gridPosition(1, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "buAutoNum-alpha", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buAlphaItems),
    }),
  );

  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: buNone, buFont, mixed
  id = 2;
  const shapes2: string[] = [];

  // buNone
  const buNoneItems = [
    { text: "No bullet here", bulletXml: `<a:buNone/>` },
    { text: "Also no bullet", bulletXml: `<a:buNone/>` },
  ];
  const pos5 = gridPosition(0, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "buNone", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buNoneItems),
    }),
  );

  // buFont + buChar (custom font bullet)
  const buFontItems = [
    {
      text: "Custom font bullet",
      bulletXml: `<a:buFont typeface="Arial"/><a:buChar char="\u25A0"/>`,
    },
    {
      text: "Another custom",
      bulletXml: `<a:buFont typeface="Arial"/><a:buChar char="\u25A0"/>`,
    },
  ];
  const pos6 = gridPosition(1, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "buFont-custom", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buFontItems),
    }),
  );

  // romanUcPeriod numbering
  const buRomanItems = [
    {
      text: "Roman one",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman two",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman three",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
    {
      text: "Roman four",
      bulletXml: `<a:buAutoNum type="romanUcPeriod"/>`,
    },
  ];
  const pos7 = gridPosition(0, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "buAutoNum-roman", {
      preset: "rect",
      x: pos7.x,
      y: pos7.y,
      cx: pos7.w,
      cy: pos7.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buRomanItems),
    }),
  );

  // Colored bullet
  const buColorItems = [
    {
      text: "Red bullet",
      bulletXml: `<a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buChar char="\u2022"/>`,
    },
    {
      text: "Blue bullet",
      bulletXml: `<a:buClr><a:srgbClr val="0000FF"/></a:buClr><a:buChar char="\u2022"/>`,
    },
    {
      text: "Green bullet",
      bulletXml: `<a:buClr><a:srgbClr val="00AA00"/></a:buClr><a:buChar char="\u2022"/>`,
    },
  ];
  const pos8 = gridPosition(1, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "buClr-colored", {
      preset: "rect",
      x: pos8.x,
      y: pos8.y,
      cx: pos8.w,
      cy: pos8.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: bulletParagraphsXml(buColorItems),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));

  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "bullets.pptx");
}

// --- 13. Flowchart Shapes ---

function multiRunTextBodyXml(
  paragraphs: {
    runs: {
      text: string;
      fontSize?: number;
      bold?: boolean;
      color?: string;
      lang?: string;
      baseline?: number;
    }[];
    align?: string;
    spcBef?: { pts?: number; pct?: number };
    spcAft?: { pts?: number; pct?: number };
  }[],
  opts?: { anchor?: string; wrap?: string },
): string {
  const anchor = opts?.anchor ?? "t";
  const wrap = opts?.wrap ? ` wrap="${opts.wrap}"` : "";
  const parasXml = paragraphs
    .map((para) => {
      const algn = para.align ? ` algn="${para.align}"` : "";
      let spcBefXml = "";
      if (para.spcBef?.pts !== undefined) {
        spcBefXml = `<a:spcBef><a:spcPts val="${para.spcBef.pts}"/></a:spcBef>`;
      } else if (para.spcBef?.pct !== undefined) {
        spcBefXml = `<a:spcBef><a:spcPct val="${para.spcBef.pct}"/></a:spcBef>`;
      }
      let spcAftXml = "";
      if (para.spcAft?.pts !== undefined) {
        spcAftXml = `<a:spcAft><a:spcPts val="${para.spcAft.pts}"/></a:spcAft>`;
      } else if (para.spcAft?.pct !== undefined) {
        spcAftXml = `<a:spcAft><a:spcPct val="${para.spcAft.pct}"/></a:spcAft>`;
      }
      const pPrContent = spcBefXml + spcAftXml;
      const runsXml = para.runs
        .map((r) => {
          const sz = r.fontSize ? ` sz="${r.fontSize * 100}"` : ` sz="1400"`;
          const b = r.bold ? ` b="1"` : "";
          const lang = r.lang ?? "en-US";
          const fillColor = r.color ?? "000000";
          const bl = r.baseline !== undefined ? ` baseline="${r.baseline}"` : "";
          return `<a:r>
        <a:rPr lang="${lang}"${sz}${b}${bl}>
          <a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>
        </a:rPr>
        <a:t>${r.text}</a:t>
      </a:r>`;
        })
        .join("\n    ");
      return `<a:p>
    <a:pPr${algn}>${pPrContent}</a:pPr>
    ${runsXml}
  </a:p>`;
    })
    .join("\n  ");

  return `<p:txBody>
  <a:bodyPr anchor="${anchor}"${wrap}/>
  <a:lstStyle/>
  ${parasXml}
</p:txBody>`;
}

async function createWordWrapFixture(): Promise<void> {
  // Slide 1: Basic word wrap scenarios
  let id = 2;
  const shapes1: string[] = [];

  // 1. Long English text in normal-width shape
  const longEnText =
    "The quick brown fox jumps over the lazy dog. This is a long sentence that should wrap across multiple lines within the shape boundary.";
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes1.push(
    shapeXml(id++, "long-en-text", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml([{ runs: [{ text: longEnText, fontSize: 14 }] }], {
        anchor: "t",
      }),
    }),
  );

  // 2. Long text in narrow shape
  const pos2 = { x: pos1.x + pos1.w + 200000, y: pos1.y, w: 1500000, h: pos1.h };
  shapes1.push(
    shapeXml(id++, "narrow-shape", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("E8F4FD"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [{ runs: [{ text: "Narrow shape forces frequent word wrapping here.", fontSize: 14 }] }],
        { anchor: "t" },
      ),
    }),
  );

  // 3. wrap="none" (no wrapping)
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "no-wrap", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("FFF3E0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "This text has wrap=none so it should not wrap at the shape boundary.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t", wrap: "none" },
      ),
    }),
  );

  // 4. Japanese text wrapping
  const pos4 = gridPosition(1, 1, 2, 2);
  shapes1.push(
    shapeXml(id++, "japanese-text", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F3E5F5"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "日本語のテキストは文字単位で折り返されます。長い文章を図形の中に配置した場合の表示を確認します。",
                fontSize: 14,
                lang: "ja-JP",
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  const slide1 = wrapSlideXml(shapes1.join("\n"));

  // Slide 2: Advanced word wrap scenarios
  id = 2;
  const shapes2: string[] = [];

  // 1. Mixed font sizes in a single paragraph
  const pos5 = gridPosition(0, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "mixed-font-sizes", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("E8F5E9"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              { text: "Large ", fontSize: 28, bold: true, color: "1565C0" },
              { text: "and small ", fontSize: 12, color: "333333" },
              { text: "mixed ", fontSize: 20, color: "C62828" },
              { text: "in one paragraph that wraps across lines.", fontSize: 14, color: "333333" },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 2. Multiple paragraphs
  const pos6 = gridPosition(1, 0, 2, 2);
  shapes2.push(
    shapeXml(id++, "multi-paragraph", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("FFF8E1"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          { runs: [{ text: "First paragraph with enough text to wrap.", fontSize: 14 }] },
          {
            runs: [{ text: "Second paragraph also wraps within the shape.", fontSize: 14 }],
            align: "ctr",
          },
          {
            runs: [{ text: "Third paragraph right-aligned.", fontSize: 14 }],
            align: "r",
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 3. Text overflow (long text in small shape)
  const pos7 = {
    x: gridPosition(0, 1, 2, 2).x,
    y: gridPosition(0, 1, 2, 2).y,
    w: 2000000,
    h: 800000,
  };
  shapes2.push(
    shapeXml(id++, "text-overflow", {
      preset: "rect",
      x: pos7.x,
      y: pos7.y,
      cx: pos7.w,
      cy: pos7.h,
      fillXml: solidFillXml("FFEBEE"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "This text is too long for the small shape and will overflow beyond the visible area of the shape boundary.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 4. Mixed CJK and Latin text
  const pos8 = gridPosition(1, 1, 2, 2);
  shapes2.push(
    shapeXml(id++, "mixed-cjk-latin", {
      preset: "rect",
      x: pos8.x,
      y: pos8.y,
      cx: pos8.w,
      cy: pos8.h,
      fillXml: solidFillXml("E0F2F1"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [
              {
                text: "English and 日本語 mixed text. テキストの折り返しが正しく動作するか確認します。Word wrap test.",
                fontSize: 14,
              },
            ],
          },
        ],
        { anchor: "t" },
      ),
    }),
  );

  const slide2 = wrapSlideXml(shapes2.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [
      { xml: slide1, rels },
      { xml: slide2, rels },
    ],
  });
  savePptx(buffer, "word-wrap.pptx");
}

// --- 18. Background blipFill ---

async function createTextDecorationFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  const testCases = [
    {
      label: "Superscript",
      paragraphs: [
        {
          runs: [
            { text: "E = mc", fontSize: 20 },
            { text: "2", fontSize: 14, baseline: 30000 },
          ],
        },
      ],
    },
    {
      label: "Subscript",
      paragraphs: [
        {
          runs: [
            { text: "H", fontSize: 20 },
            { text: "2", fontSize: 14, baseline: -25000 },
            { text: "O", fontSize: 20 },
          ],
        },
      ],
    },
    {
      label: "Mixed",
      paragraphs: [
        {
          runs: [
            { text: "x", fontSize: 18 },
            { text: "n", fontSize: 12, baseline: 30000 },
            { text: " + y", fontSize: 18 },
            { text: "m", fontSize: 12, baseline: -25000 },
          ],
        },
      ],
    },
    {
      label: "Multi-line",
      paragraphs: [
        {
          runs: [
            { text: "Line 1 with ", fontSize: 16 },
            { text: "super", fontSize: 12, baseline: 30000 },
          ],
        },
        {
          runs: [
            { text: "Line 2 with ", fontSize: 16 },
            { text: "sub", fontSize: 12, baseline: -25000 },
          ],
        },
      ],
    },
  ];

  testCases.forEach((tc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const pos = gridPosition(col, row, 2, 2);
    shapes.push(
      shapeXml(id++, tc.label, {
        preset: "rect",
        x: pos.x,
        y: pos.y,
        cx: pos.w,
        cy: pos.h,
        fillXml: solidFillXml("F0F0F0"),
        outlineXml: outlineXml(12700, "CCCCCC"),
        textBodyXml: multiRunTextBodyXml(tc.paragraphs, { anchor: "ctr" }),
      }),
    );
  });

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "text-decoration.pptx");
}

// --- 21. Slide Size 4:3 ---

async function createThemeFontFixture(): Promise<void> {
  const customTheme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${NS.a}" name="Office Theme">
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
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="Yu Gothic"/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Yu Mincho"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;

  // Text using theme font references
  const shapes = [
    // +mj-lt (major latin)
    shapeXml(2, "MajorLatin", {
      preset: "rect",
      x: 200000,
      y: 200000,
      cx: 4000000,
      cy: 1200000,
      fillXml: solidFillXml("E8EAF6"),
      textBodyXml: `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="2000">
              <a:latin typeface="+mj-lt"/>
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>Major Latin (+mj-lt)</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    }),
    // +mn-lt (minor latin)
    shapeXml(3, "MinorLatin", {
      preset: "rect",
      x: 4800000,
      y: 200000,
      cx: 4000000,
      cy: 1200000,
      fillXml: solidFillXml("E3F2FD"),
      textBodyXml: `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="2000">
              <a:latin typeface="+mn-lt"/>
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>Minor Latin (+mn-lt)</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    }),
    // +mj-ea (major east asian)
    shapeXml(4, "MajorEA", {
      preset: "rect",
      x: 200000,
      y: 1800000,
      cx: 4000000,
      cy: 1200000,
      fillXml: solidFillXml("FFF3E0"),
      textBodyXml: `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="2000">
              <a:ea typeface="+mj-ea"/>
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>Major EA (+mj-ea)</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    }),
    // +mn-ea (minor east asian)
    shapeXml(5, "MinorEA", {
      preset: "rect",
      x: 4800000,
      y: 1800000,
      cx: 4000000,
      cy: 1200000,
      fillXml: solidFillXml("E8F5E9"),
      textBodyXml: `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="2000">
              <a:ea typeface="+mn-ea"/>
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>Minor EA (+mn-ea)</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    }),
    // Explicit font (not a theme reference)
    shapeXml(6, "ExplicitFont", {
      preset: "rect",
      x: 200000,
      y: 3400000,
      cx: 8600000,
      cy: 1200000,
      fillXml: solidFillXml("F3E5F5"),
      textBodyXml: `<p:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p>
          <a:pPr algn="ctr"/>
          <a:r>
            <a:rPr lang="en-US" sz="2000">
              <a:latin typeface="Arial"/>
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>Explicit Font (Arial)</a:t>
          </a:r>
        </a:p>
      </p:txBody>`,
    }),
  ];

  const slideXml = wrapSlideXml(shapes.join("\n"));
  const slideRels = slideRelsXml();

  const buffer = await buildPptx({
    slides: [{ xml: slideXml, rels: slideRels }],
    themeXml: customTheme,
  });

  savePptx(buffer, "theme-fonts.pptx");
}

// ============================================================
// Text Style Inheritance
// ============================================================

async function createTextStyleInheritanceFixture(): Promise<void> {
  // Slide master: txStyles (titleStyle: 36pt+white, bodyStyle: 24pt+white, otherStyle: 14pt+white)
  const customSlideMaster = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr><a:defRPr sz="3600"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
      <a:lvl2pPr><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:lvl1pPr><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:lvl1pPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>`;

  // defaultTextStyle: 12pt to defRPr of lvl1pPr
  const defaultTextStyleXml = `<a:lvl1pPr><a:defRPr sz="1200"/></a:lvl1pPr>`;

  const slideRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_TYPES.slideLayout}" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

  // Shape 1: title placeholder (no fontSize -> 36pt from txStyles.titleStyle)
  const shape1 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="1143000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>Title (36pt from txStyles)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  // Shape 2: body placeholder (no fontSize -> 24pt from txStyles.bodyStyle)
  const shape2 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="8229600" cy="1143000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>Body (24pt from txStyles)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  // Shape 3: Normal shape (no fontSize -> 14pt from txStyles.otherStyle)
  const shape3 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Other"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="2971800"/><a:ext cx="3810000" cy="762000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="E7E6E6"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:t>Other (14pt from otherStyle)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  // Shape 4: Specify fontSize directly in rPr with title placeholder (20pt, takes precedence over txStyles)
  const shape4 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="Title Direct"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="3962400"/><a:ext cx="3810000" cy="762000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:rPr sz="2000"/><a:t>Title direct 20pt</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  // Shape 5: body level 1 (no fontSize -> 20pt from txStyles.bodyStyle.lvl2pPr)
  const shape5 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="6" name="Body Level2"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4876800" y="2971800"/><a:ext cx="3810000" cy="1752600"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:pPr lvl="0"/><a:r><a:t>Body L1 (24pt)</a:t></a:r></a:p>
    <a:p><a:pPr lvl="1"/><a:r><a:t>Body L2 (20pt)</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  // Shape 6: Specify color directly to rPr with body placeholder (red, takes precedence over white of txStyles)
  const shape6 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="7" name="Body Direct Color"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="3"/></p:nvPr></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="4876800" y="3962400"/><a:ext cx="3810000" cy="762000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Body direct red</a:t></a:r></a:p>
  </p:txBody>
</p:sp>`;

  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="333333"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${shape1}
      ${shape2}
      ${shape3}
      ${shape4}
      ${shape5}
      ${shape6}
    </p:spTree>
  </p:cSld>
</p:sld>`;

  const buffer = await buildPptx({
    slides: [{ xml: slideXml, rels: slideRels }],
    slideMasterXml: customSlideMaster,
    defaultTextStyleXml,
  });
  savePptx(buffer, "text-style-inheritance.pptx");
}

// --- Z-order mixed (cross-type element ordering) ---

async function createParagraphSpacingFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // 1. spaceBefore (pts)
  const pos1 = gridPosition(0, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "space-before-pts", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          { runs: [{ text: "Paragraph 1", fontSize: 14 }] },
          {
            runs: [{ text: "spaceBefore 12pt", fontSize: 14 }],
            spcBef: { pts: 1200 },
          },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 2. spaceAfter (pts)
  const pos2 = gridPosition(1, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "space-after-pts", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [{ text: "spaceAfter 12pt", fontSize: 14 }],
            spcAft: { pts: 1200 },
          },
          { runs: [{ text: "Paragraph 2", fontSize: 14 }] },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 3. spaceBefore (pct) - 50% of font size
  const pos3 = gridPosition(2, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "space-before-pct", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          { runs: [{ text: "Paragraph 1", fontSize: 14 }] },
          {
            runs: [{ text: "spcBef 50%", fontSize: 14 }],
            spcBef: { pct: 50000 },
          },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 4. spaceAfter (pct) - 100% of font size
  const pos4 = gridPosition(0, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "space-after-pct", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [{ text: "spcAft 100%", fontSize: 14 }],
            spcAft: { pct: 100000 },
          },
          { runs: [{ text: "Paragraph 2", fontSize: 14 }] },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 5. max(spaceAfter, spaceBefore) - spaceAfter wins
  const pos5 = gridPosition(1, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "max-space-after-wins", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          {
            runs: [{ text: "spcAft 20pt", fontSize: 14 }],
            spcAft: { pts: 2000 },
          },
          {
            runs: [{ text: "spcBef 5pt", fontSize: 14 }],
            spcBef: { pts: 500 },
          },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  // 6. Both spaceBefore and spaceAfter on same paragraph
  const pos6 = gridPosition(2, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "both-before-after", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: multiRunTextBodyXml(
        [
          { runs: [{ text: "Paragraph 1", fontSize: 14 }] },
          {
            runs: [{ text: "spcBef+spcAft 10pt", fontSize: 14 }],
            spcBef: { pts: 1000 },
            spcAft: { pts: 1000 },
          },
          { runs: [{ text: "Paragraph 3", fontSize: 14 }] },
        ],
        { anchor: "t" },
      ),
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "paragraph-spacing.pptx");
}

// --- Text Advanced (field codes, line breaks, tab stops) ---

async function createTextAdvancedFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // Shape 1: Field code (slide number) - Mix of text run + field code
  const pos1 = gridPosition(0, 0, 3, 3);
  shapes.push(
    shapeXml(id++, "field-slidenum", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("E8F0FE"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="ctr"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr algn="ctr"/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Slide </a:t>
    </a:r>
    <a:fld type="slidenum" uuid="{B5A3C44A-1234-5678-9ABC-DEF012345678}">
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </a:rPr>
      <a:t>1</a:t>
    </a:fld>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t> of 10</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 2: Date field code
  const pos2 = gridPosition(1, 0, 3, 3);
  shapes.push(
    shapeXml(id++, "field-date", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("FEF3E8"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="ctr"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr algn="ctr"/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Date: </a:t>
    </a:r>
    <a:fld type="datetime1" uuid="{C6B4D55B-2345-6789-ABCD-EF0123456789}">
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
      </a:rPr>
      <a:t>2024-01-15</a:t>
    </a:fld>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 3: Multiple fields interleaved with text
  const pos3 = gridPosition(2, 0, 3, 3);
  shapes.push(
    shapeXml(id++, "field-multi", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("E8FEE8"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="ctr"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr algn="ctr"/>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Page </a:t>
    </a:r>
    <a:fld type="slidenum" uuid="{A1111111-1111-1111-1111-111111111111}">
      <a:rPr lang="en-US" sz="1200" b="1">
        <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
      </a:rPr>
      <a:t>1</a:t>
    </a:fld>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t> | </a:t>
    </a:r>
    <a:fld type="datetime1" uuid="{B2222222-2222-2222-2222-222222222222}">
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>
      </a:rPr>
      <a:t>Jan 15</a:t>
    </a:fld>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 4: Line break (br)
  const pos4 = gridPosition(0, 1, 3, 3);
  shapes.push(
    shapeXml(id++, "line-break", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Line One</a:t>
    </a:r>
    <a:br>
      <a:rPr lang="en-US" sz="1400"/>
    </a:br>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </a:rPr>
      <a:t>Line Two</a:t>
    </a:r>
    <a:br>
      <a:rPr lang="en-US" sz="1400"/>
    </a:br>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill>
      </a:rPr>
      <a:t>Line Three</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 5: Tab stops
  const pos5 = gridPosition(1, 1, 3, 3);
  shapes.push(
    shapeXml(id++, "tab-stops", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("F0F0F0"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr>
      <a:tabLst>
        <a:tab pos="914400" algn="l"/>
        <a:tab pos="2743200" algn="r"/>
      </a:tabLst>
    </a:pPr>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Name&#x9;Value&#x9;Total</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr>
      <a:tabLst>
        <a:tab pos="914400" algn="l"/>
        <a:tab pos="2743200" algn="r"/>
      </a:tabLst>
    </a:pPr>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </a:rPr>
      <a:t>Item A&#x9;100&#x9;$500</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 6: numCol (multi-column)
  const pos6 = gridPosition(2, 1, 3, 3);
  shapes.push(
    shapeXml(id++, "num-col", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("FEE8FE"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t" numCol="2"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>This text is in a two-column layout. The text should wrap within the narrower column width.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // Shape 7: field + br combined
  const pos7 = gridPosition(0, 2, 3, 3);
  shapes.push(
    shapeXml(id++, "field-br-combo", {
      preset: "rect",
      x: pos7.x,
      y: pos7.y,
      cx: pos7.w,
      cy: pos7.h,
      fillXml: solidFillXml("E8E8FE"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Title</a:t>
    </a:r>
    <a:br>
      <a:rPr lang="en-US" sz="1200"/>
    </a:br>
    <a:r>
      <a:rPr lang="en-US" sz="1200">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
      </a:rPr>
      <a:t>Slide </a:t>
    </a:r>
    <a:fld type="slidenum" uuid="{D4444444-4444-4444-4444-444444444444}">
      <a:rPr lang="en-US" sz="1200" b="1">
        <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
      </a:rPr>
      <a:t>1</a:t>
    </a:fld>
  </a:p>
</p:txBody>`,
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "text-advanced.pptx");
}

async function createShrinkToFitFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // 1. normAutofit (without fontScale) - Case where text protrudes -> dynamic reduction
  const pos1 = gridPosition(0, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "shrink-overflow", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("F0F0F0"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="3600">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Long text shrunk to fit within this shape.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 2. normAutofit (without fontScale) - case where text fits -> no scaling
  const pos2 = gridPosition(1, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "shrink-fits", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("E8F4FD"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Short text</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 3. normAutofit + fontSize unspecified - reduction at default font size
  const pos3 = gridPosition(2, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "shrink-default-fontsize", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("E8FDE8"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Text without explicit font size that overflows and should be auto-shrunk by normAutofit.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 4. noAutofit - Leave the text as it is even if it extends
  const pos4 = gridPosition(0, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "no-autofit-overflow", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("FDE8E8"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="3600">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Long text that overflows without shrinking.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 5. normAutofit + multiple paragraphs
  const pos5 = gridPosition(1, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "shrink-multi-para", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("FFF0F5"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:normAutofit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>First paragraph.</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Second paragraph.</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Third paragraph.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 6. normAutofit + fontScale preset - cases where further dynamic scaling is required
  const pos6 = gridPosition(2, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "shrink-with-fontscale", {
      preset: "rect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("F0E8FD"),
      outlineXml: outlineXml(12700, "999999"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:normAutofit fontScale="80000"/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="3600">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Text with fontScale 80% that still overflows and needs further shrinking.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "shrink-to-fit.pptx");
}

async function createSpAutofitFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // 1. spAutofit - Case where the text extends -> the shape is enlarged
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes.push(
    shapeXml(id++, "sp-autofit-overflow", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("E8F0FE"),
      outlineXml: outlineXml(12700, "4472C4"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:spAutoFit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="3600">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Long text that causes the shape to grow taller automatically.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 2. spAutofit - Case where text fits -> shape remains the same
  const pos2 = gridPosition(1, 0, 2, 2);
  shapes.push(
    shapeXml(id++, "sp-autofit-fits", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("E8FDE8"),
      outlineXml: outlineXml(12700, "70AD47"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:spAutoFit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Short text</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 3. spAutofit + multiple paragraphs
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes.push(
    shapeXml(id++, "sp-autofit-multi-para", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("FFF0E0"),
      outlineXml: outlineXml(12700, "ED7D31"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t"><a:spAutoFit/></a:bodyPr>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>First paragraph with enough text to overflow.</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Second paragraph adds more content.</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Third paragraph for extra height.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "sp-autofit.pptx");
}

// --- Style Reference ---

async function createStyleReferenceFixture(): Promise<void> {
  // Custom theme with fmtScheme containing fill/line/effect styles
  const styleRefTheme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${NS.a}" name="Office Theme">
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
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
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

  const shapes: string[] = [];

  // Shape 1: fillRef=1 (solid fill via style), lnRef=1 (thin line via style)
  const p1 = gridPosition(0, 0, 3, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="StyleRef Solid"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p1.x}" y="${p1.y}"/><a:ext cx="${p1.w}" cy="${p1.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  ${textBodyXmlHelper("fillRef=1, lnRef=1", { fontSize: 12 })}
</p:sp>`);

  // Shape 2: fillRef=2 (gradient fill via style), lnRef=2 (thick line via style)
  const p2 = gridPosition(1, 0, 3, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="StyleRef Gradient"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p2.x}" y="${p2.y}"/><a:ext cx="${p2.w}" cy="${p2.h}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef>
    <a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>
    <a:effectRef idx="1"><a:schemeClr val="accent2"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  ${textBodyXmlHelper("fillRef=2, lnRef=2, effectRef=1", { fontSize: 10 })}
</p:sp>`);

  // Shape 3: Direct fill overrides style ref (spPr has solidFill, style has fillRef)
  const p3 = gridPosition(2, 0, 3, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Direct Override"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p3.x}" y="${p3.y}"/><a:ext cx="${p3.w}" cy="${p3.h}"/></a:xfrm>
    <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FF6384"/></a:solidFill>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent3"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent3"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent3"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  ${textBodyXmlHelper("Direct fill override", { fontSize: 10 })}
</p:sp>`);

  // Shape 4: lnRef=0 (no line), fillRef=1 with accent4
  const p4 = gridPosition(0, 1, 3, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="No Line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p4.x}" y="${p4.y}"/><a:ext cx="${p4.w}" cy="${p4.h}"/></a:xfrm>
    <a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="0"><a:schemeClr val="accent4"/></a:lnRef>
    <a:fillRef idx="1"><a:schemeClr val="accent4"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent4"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  ${textBodyXmlHelper("lnRef=0, fillRef=1", { fontSize: 10 })}
</p:sp>`);

  // Shape 5: fillRef=0 (no fill), lnRef=3 (thick line) with accent5
  const p5 = gridPosition(1, 1, 3, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="6" name="No Fill Thick Line"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p5.x}" y="${p5.y}"/><a:ext cx="${p5.w}" cy="${p5.h}"/></a:xfrm>
    <a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="3"><a:schemeClr val="accent5"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent5"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent5"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
  ${textBodyXmlHelper("fillRef=0, lnRef=3", { fontSize: 10 })}
</p:sp>`);

  // Connector with style reference
  const p6 = gridPosition(2, 1, 3, 2);
  shapes.push(`<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="7" name="Styled Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${p6.x}" y="${p6.y}"/><a:ext cx="${p6.w}" cy="${p6.h}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:style>
    <a:lnRef idx="2"><a:schemeClr val="accent6"/></a:lnRef>
    <a:fillRef idx="0"><a:schemeClr val="accent6"/></a:fillRef>
    <a:effectRef idx="0"><a:schemeClr val="accent6"/></a:effectRef>
    <a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>
  </p:style>
</p:cxnSp>`);

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({
    slides: [{ xml: slide, rels }],
    themeXml: styleRefTheme,
  });
  savePptx(buffer, "style-reference.pptx");
}

async function createShapeHyperlinkTextOutlineFixture(): Promise<void> {
  const margin = 457200; // 0.5 inch
  const shapeW = 8229600; // 8.5 inch
  const shapeH = 1371600; // 1.5 inch

  // Shape 1: Shape-level hyperlink (entire shape is clickable)
  const shape1 = `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="Shape 1">
      <a:hlinkClick r:id="rId2"/>
    </p:cNvPr>
    <p:cNvSpPr/><p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="1800">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        </a:rPr>
        <a:t>Shape-level hyperlink (click entire shape)</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  // Shape 2: Text outline (black stroke on red text)
  const shape2 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Shape 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin + shapeH + margin}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="3600">
          <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
          <a:ln w="12700">
            <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          </a:ln>
        </a:rPr>
        <a:t>Text Outline</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  // Shape 3: Text outline with thick stroke (white text with blue outline)
  const shape3 = `<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Shape 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${margin}" y="${margin + (shapeH + margin) * 2}"/><a:ext cx="${shapeW}" cy="${shapeH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="ctr"/>
      <a:r>
        <a:rPr lang="en-US" sz="4800">
          <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
          <a:ln w="25400">
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          </a:ln>
        </a:rPr>
        <a:t>Thick Outline</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  const slide = wrapSlideXml([shape1, shape2, shape3].join("\n"));
  const rels = slideRelsXml([
    {
      id: "rId2",
      type: REL_TYPES.hyperlink,
      target: "https://example.com/shape-link",
      targetMode: "External",
    },
  ]);

  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "shape-hyperlink-text-outline.pptx");
}

async function createVerticalTextFixture(): Promise<void> {
  let id = 2;
  const shapes: string[] = [];

  // 1. vert="vert" (90° CW) - basic vertical writing
  const pos1 = gridPosition(0, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-90cw", {
      preset: "rect",
      x: pos1.x,
      y: pos1.y,
      cx: pos1.w,
      cy: pos1.h,
      fillXml: solidFillXml("E8F0FE"),
      outlineXml: outlineXml(12700, "4472C4"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t" vert="vert"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1800">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Vertical Text (90 CW)</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 2. vert="vert270" (90° CCW)
  const pos2 = gridPosition(1, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-270ccw", {
      preset: "rect",
      x: pos2.x,
      y: pos2.y,
      cx: pos2.w,
      cy: pos2.h,
      fillXml: solidFillXml("FDE8E8"),
      outlineXml: outlineXml(12700, "ED7D31"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t" vert="vert270"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1800">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Vertical Text (270 CCW)</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 3. vert="eaVert" (East Asian vertical)
  const pos3 = gridPosition(2, 0, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-ea", {
      preset: "rect",
      x: pos3.x,
      y: pos3.y,
      cx: pos3.w,
      cy: pos3.h,
      fillXml: solidFillXml("E8FDE8"),
      outlineXml: outlineXml(12700, "70AD47"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t" vert="eaVert"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="ja-JP" sz="1800">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>縦書きテスト</a:t>
    </a:r>
  </a:p>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="ja-JP" sz="1400">
        <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
        <a:latin typeface="Calibri"/>
        <a:ea typeface="Meiryo"/>
      </a:rPr>
      <a:t>ABC漢字DEF</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 4. vert="vert" + anchor="ctr" (center alignment)
  const pos4 = gridPosition(0, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-center", {
      preset: "rect",
      x: pos4.x,
      y: pos4.y,
      cx: pos4.w,
      cy: pos4.h,
      fillXml: solidFillXml("FFF0E0"),
      outlineXml: outlineXml(12700, "FFC000"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="ctr" vert="vert"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr algn="ctr"/>
    <a:r>
      <a:rPr lang="en-US" sz="2400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Centered</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 5. vert="vert" + multiple line text wrapping
  const pos5 = gridPosition(1, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-wrap", {
      preset: "rect",
      x: pos5.x,
      y: pos5.y,
      cx: pos5.w,
      cy: pos5.h,
      fillXml: solidFillXml("E8E0F8"),
      outlineXml: outlineXml(12700, "9966FF"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="t" vert="vert"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1400">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>This is a longer text that should wrap in vertical mode.</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  // 6. vert="vert" + anchor="b" (bottom alignment)
  const pos6 = gridPosition(2, 1, 3, 2);
  shapes.push(
    shapeXml(id++, "vert-text-bottom", {
      preset: "roundRect",
      x: pos6.x,
      y: pos6.y,
      cx: pos6.w,
      cy: pos6.h,
      fillXml: solidFillXml("F0E8E8"),
      outlineXml: outlineXml(12700, "954F72"),
      textBodyXml: `<p:txBody>
  <a:bodyPr anchor="b" vert="vert"/>
  <a:lstStyle/>
  <a:p>
    <a:pPr/>
    <a:r>
      <a:rPr lang="en-US" sz="1800">
        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
      </a:rPr>
      <a:t>Bottom Aligned</a:t>
    </a:r>
  </a:p>
</p:txBody>`,
    }),
  );

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "vertical-text.pptx");
}

// --- Charts 3D Fallback ---

async function createMultiLangFontFixture(): Promise<void> {
  const shapes: string[] = [];

  // Shape 1 (row0, col0): Latin+CJK mixed text with different fonts
  const pos1 = gridPosition(0, 0, 2, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Mixed"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${pos1.x}" y="${pos1.y}"/><a:ext cx="${pos1.w}" cy="${pos1.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F0F0F0"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1600">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:latin typeface="Liberation Sans"/>
          <a:ea typeface="Noto Sans CJK JP"/>
        </a:rPr>
        <a:t>Hello World 日本語テスト</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`);

  // Shape 2 (row0, col1): CJK only
  const pos2 = gridPosition(1, 0, 2, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="CJK Only"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${pos2.x}" y="${pos2.y}"/><a:ext cx="${pos2.w}" cy="${pos2.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="E8F4FD"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="ja-JP" sz="1600">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:latin typeface="Liberation Sans"/>
          <a:ea typeface="Noto Sans CJK JP"/>
        </a:rPr>
        <a:t>日本語テスト</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`);

  // Shape 3 (row1, col0): Latin only
  const pos3 = gridPosition(0, 1, 2, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Latin Only"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${pos3.x}" y="${pos3.y}"/><a:ext cx="${pos3.w}" cy="${pos3.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="FDF5E6"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1600">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:latin typeface="Liberation Serif"/>
          <a:ea typeface="Noto Sans CJK JP"/>
        </a:rPr>
        <a:t>Hello World</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`);

  // Shape 4 (row1, col1): Same CJK-capable font for both (no split needed)
  const pos4 = gridPosition(1, 1, 2, 2);
  shapes.push(`<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="Same Font"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${pos4.x}" y="${pos4.y}"/><a:ext cx="${pos4.w}" cy="${pos4.h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F0F0E0"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr"/>
    <a:lstStyle/>
    <a:p>
      <a:r>
        <a:rPr lang="en-US" sz="1600">
          <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
          <a:latin typeface="Noto Sans CJK JP"/>
          <a:ea typeface="Noto Sans CJK JP"/>
        </a:rPr>
        <a:t>Test テスト (same font)</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`);

  const slide = wrapSlideXml(shapes.join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slide, rels }] });
  savePptx(buffer, "multi-lang-font.pptx");
}

// --- Placeholder Inheritance Extended ---

async function createInterleavedBulletPprFixture(): Promise<void> {
  const bulletItems = [
    { label: "Product", desc: "AI dashboard beta release" },
    { label: "Sales", desc: "ARR target 800M" },
    { label: "Team", desc: "Hire 15 engineers" },
    { label: "Partners", desc: "3 new contracts" },
  ];

  // Alternating pPr/r pattern: buChar pPr + bold r + buNone pPr + normal r (with \n)
  const interleavedRuns = bulletItems
    .map((item, i) => {
      const trailingNewline = i < bulletItems.length - 1 ? "\n" : "";
      return `
          <a:pPr algn="l" marL="342900" indent="-342900">
            <a:lnSpc><a:spcPct val="130000"/></a:lnSpc>
            <a:buSzPct val="100000"/>
            <a:buChar char="&#x2022;"/>
          </a:pPr>
          <a:r>
            <a:rPr lang="en-US" sz="1600" b="1">
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>${item.label}</a:t>
          </a:r>
          <a:pPr algn="l" indent="0" marL="0">
            <a:lnSpc><a:spcPct val="130000"/></a:lnSpc>
            <a:buNone/>
          </a:pPr>
          <a:r>
            <a:rPr lang="en-US" sz="1600">
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>: ${item.desc}${trailingNewline}</a:t>
          </a:r>`;
    })
    .join("");

  const titleShape = `<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="274638"/><a:ext cx="8229600" cy="400000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="l"/>
      <a:r>
        <a:rPr lang="en-US" sz="2000" b="1">
          <a:solidFill><a:srgbClr val="333333"/></a:solidFill>
        </a:rPr>
        <a:t>Interleaved pPr/r bullet list</a:t>
      </a:r>
    </a:p>
  </p:txBody>
</p:sp>`;

  const bulletShape = `<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="Bullets"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="800000"/><a:ext cx="8229600" cy="2400000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="F5F5F5"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
    <a:lstStyle/>
    <a:p>${interleavedRuns}
      <a:endParaRPr lang="en-US" sz="1600"/>
    </a:p>
  </p:txBody>
</p:sp>`;

  const slideXml = wrapSlideXml([titleShape, bulletShape].join("\n"));
  const rels = slideRelsXml();
  const buffer = await buildPptx({ slides: [{ xml: slideXml, rels }] });
  savePptx(buffer, "interleaved-bullet-ppr.pptx");
}

export const textFixtureCreators: FixtureCreatorMap = {
  "text.pptx": createTextFixture,
  "bullets.pptx": createBulletsFixture,
  "word-wrap.pptx": createWordWrapFixture,
  "text-decoration.pptx": createTextDecorationFixture,
  "theme-fonts.pptx": createThemeFontFixture,
  "text-style-inheritance.pptx": createTextStyleInheritanceFixture,
  "paragraph-spacing.pptx": createParagraphSpacingFixture,
  "text-advanced.pptx": createTextAdvancedFixture,
  "shrink-to-fit.pptx": createShrinkToFitFixture,
  "sp-autofit.pptx": createSpAutofitFixture,
  "style-reference.pptx": createStyleReferenceFixture,
  "shape-hyperlink-text-outline.pptx": createShapeHyperlinkTextOutlineFixture,
  "vertical-text.pptx": createVerticalTextFixture,
  "multi-lang-font.pptx": createMultiLangFontFixture,
  "interleaved-bullet-ppr.pptx": createInterleavedBulletPprFixture,
};
