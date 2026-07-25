import {
  asPartPath,
  asPt,
  asRawSidecarId,
  asSourceNodeId,
  type SourceHandle,
  type SourceTextBody,
} from "@pptx-glimpse/document";
import { describe, expect, it } from "vitest";

import {
  proseMirrorDocJsonToEditorCommands,
  proseMirrorDocJsonToTextBody,
  textBodyToProseMirrorDocJson,
} from "./prosemirror-text-body-compat.js";

const paragraphHandle = handle("paragraph-1", 0);
const firstRunHandle = handle("run-1", 1);
const secondRunHandle = handle("run-2", 2);
const textBody: SourceTextBody = {
  handle: handle("text-body-1", 0),
  properties: { anchor: "middle", marginLeft: 100 },
  paragraphs: [
    {
      handle: paragraphHandle,
      properties: { align: "left", level: 0 },
      runs: [
        {
          kind: "textRun",
          handle: firstRunHandle,
          text: "One",
          properties: { bold: true, fontSize: asPt(24), typeface: "Aptos" },
        },
        {
          kind: "textRun",
          handle: secondRunHandle,
          text: " two",
          properties: { italic: true },
        },
      ],
    },
  ],
};

describe("ProseMirror text body compatibility", () => {
  it("round-trips source paragraphs, runs, handles, and properties", () => {
    const roundTripped = proseMirrorDocJsonToTextBody(
      textBody,
      textBodyToProseMirrorDocJson(textBody),
    );

    expect(roundTripped).toEqual(textBody);
  });

  it("converts compatible text and paragraph property edits to editor commands", () => {
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const paragraph = docJson.content?.[0];
    const firstRun = paragraph?.content?.[0];
    const secondRun = paragraph?.content?.[1];
    if (paragraph === undefined || firstRun === undefined || secondRun === undefined) {
      throw new Error("compatibility fixture is incomplete");
    }

    const commands = proseMirrorDocJsonToEditorCommands(textBody, {
      type: "doc",
      content: [
        {
          ...paragraph,
          attrs: {
            ...paragraph.attrs,
            properties: { align: "right", level: 0 },
          },
          content: [{ ...firstRun, text: "Edited" }, secondRun],
        },
      ],
    });

    expect(commands).toEqual([
      {
        kind: "setParagraphProperties",
        handle: paragraphHandle,
        properties: { align: "right" },
      },
      {
        kind: "replaceTextRunPlainText",
        handle: firstRunHandle,
        text: "Edited",
      },
    ]);
  });

  it("falls back to paragraph replacement when run order changes", () => {
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const paragraph = docJson.content?.[0];
    const firstRun = paragraph?.content?.[0];
    const secondRun = paragraph?.content?.[1];
    if (paragraph === undefined || firstRun === undefined || secondRun === undefined) {
      throw new Error("compatibility fixture is incomplete");
    }

    expect(
      proseMirrorDocJsonToEditorCommands(textBody, {
        type: "doc",
        content: [{ ...paragraph, content: [secondRun, firstRun] }],
      }),
    ).toEqual([
      {
        kind: "replaceParagraphPlainText",
        handle: paragraphHandle,
        text: " twoOne",
      },
    ]);
  });

  it("preserves run handles when an edit changes text across multiple runs", () => {
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const paragraph = docJson.content?.[0];
    const firstRun = paragraph?.content?.[0];
    const secondRun = paragraph?.content?.[1];
    if (paragraph === undefined || firstRun === undefined || secondRun === undefined) {
      throw new Error("compatibility fixture is incomplete");
    }

    expect(
      proseMirrorDocJsonToEditorCommands(textBody, {
        type: "doc",
        content: [
          {
            ...paragraph,
            content: [
              { ...firstRun, text: "OnX" },
              { ...secondRun, text: "two" },
            ],
          },
        ],
      }),
    ).toEqual([
      {
        kind: "replaceTextRunPlainText",
        handle: firstRunHandle,
        text: "OnX",
      },
      {
        kind: "replaceTextRunPlainText",
        handle: secondRunHandle,
        text: "two",
      },
    ]);
  });

  it("converts removed and added paragraph properties to clear and set commands", () => {
    const docJson = textBodyToProseMirrorDocJson(textBody);
    const paragraph = docJson.content?.[0];
    if (paragraph === undefined) throw new Error("compatibility fixture is incomplete");

    expect(
      proseMirrorDocJsonToEditorCommands(textBody, {
        type: "doc",
        content: [
          {
            ...paragraph,
            attrs: {
              ...paragraph.attrs,
              properties: { bullet: { type: "char", char: "\u2022" } },
            },
          },
        ],
      }),
    ).toEqual([
      {
        kind: "clearParagraphProperties",
        handle: paragraphHandle,
        properties: ["align", "level"],
      },
      {
        kind: "setParagraphProperties",
        handle: paragraphHandle,
        properties: { bullet: { type: "char", char: "\u2022" } },
      },
    ]);
  });

  it.each([
    { text: "", rawSidecars: undefined, expected: /empty text runs/ },
    { text: "\n", rawSidecars: undefined, expected: /unsupported run-like/ },
    {
      text: "One",
      rawSidecars: [{ id: asRawSidecarId("raw-1"), node: { name: "a:br" } }],
      expected: /unsupported run-like/,
    },
  ] as const)("rejects unsupported run-like content: %#", ({ text, rawSidecars, expected }) => {
    expect(() =>
      textBodyToProseMirrorDocJson({
        ...textBody,
        paragraphs: [
          {
            ...textBody.paragraphs[0],
            runs: [
              {
                ...textBody.paragraphs[0].runs[0],
                text,
                ...(rawSidecars !== undefined ? { rawSidecars } : {}),
              },
            ],
          },
        ],
      }),
    ).toThrow(expected);
  });
});

function handle(nodeId: string, orderingSlot: number): SourceHandle {
  return {
    partPath: asPartPath("ppt/slides/slide1.xml"),
    nodeId: asSourceNodeId(nodeId),
    orderingSlot,
  };
}
