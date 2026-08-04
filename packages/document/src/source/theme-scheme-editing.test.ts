import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";

// Exercise the supported package-root boundary.
import {
  addShape,
  asEmu,
  asPartPath,
  createComputedView,
  createPptx,
  type PptxSourceModel,
  readPptx,
  type SourceHandle,
  updateThemeScheme,
  writePptx,
} from "../index.js";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const themePath = "ppt/theme/theme1.xml";

describe("updateThemeScheme", () => {
  it("writes, reads, computes, and preserves a field-level existing-theme edit", () => {
    const input = existingThemeFixture();
    const source = readPptx(input);
    const handle = requireThemeHandle(source);
    const originalTheme = source.themes[0];
    const originalMasterRelationships = requireEntry(
      unzipSync(input),
      "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    );

    const colors = {
      dk1: "010101",
      lt1: "F1F1F1",
      dk2: "020202",
      lt2: "E2E2E2",
      accent1: "123456",
      accent2: "234567",
      accent3: "345678",
      accent4: "456789",
      accent5: "56789A",
      accent6: "6789AB",
      hlink: "789ABC",
      folHlink: "89ABCD",
    } as const;
    const edited = updateThemeScheme(source, handle, {
      colorScheme: colors,
      fontScheme: {
        major: {
          latin: "Brand Display",
          eastAsian: "Noto Sans CJK JP",
          complexScript: "Noto Sans Arabic",
        },
        minor: {
          latin: "Brand Text",
          eastAsian: "Yu Gothic",
          complexScript: "Arial",
        },
      },
    });

    expect(source.themes[0]).toBe(originalTheme);
    expect(source.edits).toBeUndefined();
    expect(edited.themes[0]).toMatchObject({
      partPath: themePath,
      handle,
      colorScheme: { colors: { accent1: { kind: "srgb", hex: "123456" } } },
      fontScheme: {
        majorLatin: "Brand Display",
        majorEastAsian: "Noto Sans CJK JP",
        majorComplexScript: "Noto Sans Arabic",
        minorLatin: "Brand Text",
        minorEastAsian: "Yu Gothic",
        minorComplexScript: "Arial",
      },
    });

    const output = writePptx(edited);
    const outputEntries = unzipSync(output);
    const inputEntries = unzipSync(input);
    expect(Object.keys(outputEntries).sort()).toEqual(Object.keys(inputEntries).sort());
    expect(requireEntry(outputEntries, "ppt/slideMasters/_rels/slideMaster1.xml.rels")).toEqual(
      originalMasterRelationships,
    );
    const themeXml = decoder.decode(requireEntry(outputEntries, themePath));
    const inputThemeXml = decoder.decode(requireEntry(inputEntries, themePath));
    for (const [slot, hex] of Object.entries(colors)) {
      expect(themeXml).toContain(`<a:${slot}><a:srgbClr val="${hex}"/></a:${slot}>`);
    }
    expect(extractElement(themeXml, "a:fmtScheme")).toBe(
      extractElement(inputThemeXml, "a:fmtScheme"),
    );
    expect(themeXml).toContain('<x:preserve value="unknown-theme-xml"/>');
    expect(themeXml).toContain('<a:font script="Jpan" typeface="Preserved Japanese"/>');

    const reread = readPptx(output);
    expect(reread.slideMasters[0]?.themePartPath).toBe(themePath);
    expect(reread.themes[0]?.partPath).toBe(themePath);
    expect(reread.themes[0]?.colorScheme?.colors.accent1).toEqual({
      kind: "srgb",
      hex: "123456",
    });
    expect(reread.themes[0]?.fontScheme).toMatchObject({
      majorLatin: "Brand Display",
      minorLatin: "Brand Text",
      majorJapanese: "Preserved Japanese",
    });
    expect(createComputedView(reread).slides[0]?.elements[0]).toMatchObject({
      kind: "shape",
      fill: { kind: "solid", color: { hex: "#123456" } },
      textBody: {
        paragraphs: [
          {
            runs: [
              {
                properties: {
                  typeface: "Brand Display",
                  typefaceEa: "Noto Sans CJK JP",
                  typefaceCs: "Noto Sans Arabic",
                },
              },
            ],
          },
        ],
      },
    });
  });

  it("preserves every omitted color and font field", () => {
    const source = readPptx(existingThemeFixture());
    const edited = updateThemeScheme(source, requireThemeHandle(source), {
      colorScheme: { accent1: "abcdef" },
      fontScheme: { major: { latin: "Only Changed" } },
    });
    expect(edited.themes[0]?.colorScheme?.colors).toMatchObject({
      accent1: { kind: "srgb", hex: "ABCDEF" },
      accent2: source.themes[0]?.colorScheme?.colors.accent2,
      dk1: source.themes[0]?.colorScheme?.colors.dk1,
    });
    expect(edited.themes[0]?.fontScheme).toMatchObject({
      majorLatin: "Only Changed",
      majorEastAsian: source.themes[0]?.fontScheme?.majorEastAsian,
      minorLatin: source.themes[0]?.fontScheme?.minorLatin,
      minorComplexScript: source.themes[0]?.fontScheme?.minorComplexScript,
    });
    const reread = readPptx(writePptx(edited));
    expect(reread.themes[0]?.colorScheme?.colors).toMatchObject({
      accent1: { kind: "srgb", hex: "ABCDEF" },
      accent2: source.themes[0]?.colorScheme?.colors.accent2,
      dk1: source.themes[0]?.colorScheme?.colors.dk1,
    });
    expect(reread.themes[0]?.fontScheme).toMatchObject({
      majorLatin: "Only Changed",
      majorEastAsian: source.themes[0]?.fontScheme?.majorEastAsian,
      minorLatin: source.themes[0]?.fontScheme?.minorLatin,
      minorComplexScript: source.themes[0]?.fontScheme?.minorComplexScript,
    });
  });

  it("rejects missing or ambiguous handles, invalid inputs, and unsupported raw structures", () => {
    const source = readPptx(existingThemeFixture());
    const handle = requireThemeHandle(source);
    const assertAtomicFailure = (
      candidate: PptxSourceModel,
      operation: () => unknown,
      message: RegExp,
    ) => {
      const themes = candidate.themes;
      const edits = candidate.edits;
      expect(operation).toThrow(message);
      expect(candidate.themes).toBe(themes);
      expect(candidate.edits).toBe(edits);
    };

    const missingHandle: SourceHandle = { partPath: asPartPath("ppt/theme/missing.xml") };
    assertAtomicFailure(
      source,
      () =>
        updateThemeScheme(source, missingHandle, {
          colorScheme: { accent1: "123456" },
        }),
      /theme handle was not found/,
    );
    const ambiguous = { ...source, themes: [source.themes[0], source.themes[0]] };
    assertAtomicFailure(
      ambiguous,
      () => updateThemeScheme(ambiguous, handle, { colorScheme: { accent1: "123456" } }),
      /theme handle is ambiguous/,
    );
    assertAtomicFailure(
      source,
      () =>
        updateThemeScheme(source, handle, {
          // @ts-expect-error Runtime validation covers untyped callers.
          colorScheme: { accent1: "12345" },
        }),
      /6-digit hex color/,
    );
    assertAtomicFailure(
      source,
      () => updateThemeScheme(source, handle, { fontScheme: { major: { latin: "" } } }),
      /latin must be a non-empty string/,
    );
    assertAtomicFailure(
      source,
      () => updateThemeScheme(source, handle, {}),
      /at least one color or font field/,
    );

    const files = unzipSync(existingThemeFixture());
    const malformedTheme = decoder
      .decode(requireEntry(files, themePath))
      .replace(
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>',
        '<a:accent1><a:srgbClr val="4472C4"/><a:sysClr val="windowText"/></a:accent1>',
      );
    const unsupported = readPptx(
      zipSync({ ...files, [themePath]: encoder.encode(malformedTheme) }),
    );
    assertAtomicFailure(
      unsupported,
      () =>
        updateThemeScheme(unsupported, requireThemeHandle(unsupported), {
          colorScheme: { accent1: "123456" },
        }),
      /must contain exactly one supported color/,
    );

    const missingFontFiles = unzipSync(existingThemeFixture());
    const missingFontTheme = decoder
      .decode(requireEntry(missingFontFiles, themePath))
      .replace('<a:latin typeface="Original Display"/>', "");
    const unsupportedFont = readPptx(
      zipSync({ ...missingFontFiles, [themePath]: encoder.encode(missingFontTheme) }),
    );
    assertAtomicFailure(
      unsupportedFont,
      () =>
        updateThemeScheme(unsupportedFont, requireThemeHandle(unsupportedFont), {
          fontScheme: { major: { latin: "Brand Display" } },
        }),
      /exactly one a:latin element/,
    );
  });
});

function existingThemeFixture(): Uint8Array {
  let source = createPptx({
    theme: {
      fontScheme: {
        major: {
          latin: "Original Display",
          eastAsian: "Original EA",
          complexScript: "Original CS",
        },
        minor: {
          latin: "Original Text",
          eastAsian: "Original Minor EA",
          complexScript: "Original Minor CS",
        },
      },
    },
  });
  const slideHandle = source.slides[0]?.handle;
  if (slideHandle === undefined) throw new Error("fixture slide handle is missing");
  source = addShape(source, slideHandle, {
    geometry: { kind: "preset", preset: "rect" },
    offsetX: asEmu(100000),
    offsetY: asEmu(100000),
    width: asEmu(1000000),
    height: asEmu(500000),
    fill: { kind: "solid", color: { kind: "srgb", hex: "FFFFFF" } },
    paragraphs: [{ runs: [{ text: "Theme", properties: { fontFace: "+mj-lt" } }] }],
  });
  const files = unzipSync(writePptx(source));
  const themeXml = decoder
    .decode(requireEntry(files, themePath))
    .replace(
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:test:preserve"',
    )
    .replace(
      '<a:ea typeface="Original EA"/>',
      '<a:ea typeface="Original EA"/><a:font script="Jpan" typeface="Preserved Japanese"/>',
    )
    .replace("</a:theme>", '<x:preserve value="unknown-theme-xml"/></a:theme>');
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = decoder
    .decode(requireEntry(files, slidePath))
    .replace('<a:srgbClr val="FFFFFF"/>', '<a:schemeClr val="accent1"/>')
    .replace(
      '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-lt"/><a:cs typeface="+mj-lt"/>',
      '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/>',
    );
  return zipSync({
    ...files,
    [themePath]: encoder.encode(themeXml),
    [slidePath]: encoder.encode(slideXml),
  });
}

function requireThemeHandle(source: PptxSourceModel): SourceHandle {
  const handle = source.themes[0]?.handle;
  if (handle === undefined) throw new Error("fixture theme handle is missing");
  return handle;
}

function requireEntry(files: Record<string, Uint8Array>, path: string): Uint8Array {
  const value = files[path];
  if (value === undefined) throw new Error(`fixture entry '${path}' is missing`);
  return value;
}

function extractElement(xml: string, qualifiedName: string): string {
  const match = new RegExp(`<${qualifiedName}\\b[\\s\\S]*?</${qualifiedName}>`).exec(xml)?.[0];
  if (match === undefined) throw new Error(`fixture element '${qualifiedName}' is missing`);
  return match;
}
