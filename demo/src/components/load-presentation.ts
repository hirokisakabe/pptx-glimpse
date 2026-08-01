import {
  convertPptxToSvg,
  createPptxEditorSession,
  type FontBuffer,
  type PptxEditorSession,
  type SlideSvg,
} from "pptx-glimpse";

export type DemoMode = "view" | "edit";

export type LoadedPresentation =
  | {
      readonly mode: "view";
      readonly file: File;
      readonly fileName: string;
      readonly fonts: readonly FontBuffer[];
      readonly slides: readonly SlideSvg[];
    }
  | {
      readonly mode: "edit";
      readonly file: File;
      readonly fileName: string;
      readonly fonts: readonly FontBuffer[];
      readonly editor: PptxEditorSession;
    };

interface PresentationFile {
  readonly name: string;
  arrayBuffer(): Promise<ArrayBuffer>;
}

interface PresentationLoaderDependencies<
  Slide,
  Editor extends { readonly slides: readonly unknown[] },
> {
  convertToSvg(input: Uint8Array, fonts: readonly FontBuffer[]): Promise<readonly Slide[]>;
  createEditor(input: Uint8Array, fonts: readonly FontBuffer[]): Promise<Editor>;
}

type PresentationLoadResult<FileType, Slide, Editor> =
  | {
      readonly mode: "view";
      readonly file: FileType;
      readonly fileName: string;
      readonly fonts: readonly FontBuffer[];
      readonly slides: readonly Slide[];
    }
  | {
      readonly mode: "edit";
      readonly file: FileType;
      readonly fileName: string;
      readonly fonts: readonly FontBuffer[];
      readonly editor: Editor;
    };

export async function loadPresentation(
  file: File,
  mode: DemoMode,
  fontFiles: readonly File[],
): Promise<LoadedPresentation> {
  return loadPresentationWith(file, mode, fontFiles, {
    convertToSvg: async (input, fonts) => {
      const report = await convertPptxToSvg(input, {
        fonts: [...fonts],
        skipSystemFonts: true,
      });
      return report.slides;
    },
    createEditor: (input, fonts) =>
      createPptxEditorSession(input, {
        fonts: [...fonts],
        skipSystemFonts: true,
        textOutput: "text",
      }),
  });
}

export async function loadPresentationWith<
  FileType extends PresentationFile,
  Slide,
  Editor extends { readonly slides: readonly unknown[] },
>(
  file: FileType,
  mode: DemoMode,
  fontFiles: readonly PresentationFile[],
  dependencies: PresentationLoaderDependencies<Slide, Editor>,
): Promise<PresentationLoadResult<FileType, Slide, Editor>> {
  const [pptxArrayBuffer, fonts] = await Promise.all([
    file.arrayBuffer(),
    readFontBuffers(fontFiles),
  ]);
  const pptxBytes = new Uint8Array(pptxArrayBuffer);

  if (mode === "view") {
    const slides = await dependencies.convertToSvg(pptxBytes, fonts);
    assertSlidesPresent(slides.length);
    return { mode, file, fileName: file.name, fonts, slides };
  }

  const editor = await dependencies.createEditor(pptxBytes, fonts);
  assertSlidesPresent(editor.slides.length);
  return { mode, file, fileName: file.name, fonts, editor };
}

async function readFontBuffers(files: readonly PresentationFile[]): Promise<FontBuffer[]> {
  return Promise.all(
    files.map(async (file) => ({
      name: file.name.replace(/\.(?:ttf|otf|ttc)$/i, ""),
      data: await file.arrayBuffer(),
    })),
  );
}

function assertSlidesPresent(slideCount: number): void {
  if (slideCount === 0) throw new Error("No slides found in the selected file");
}
