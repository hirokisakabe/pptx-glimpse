# Writing and round-trip preservation

Use `writePptx(source)` to serialize a `PptxSourceModel` to PPTX bytes:

```ts
import { readFile, writeFile } from "node:fs/promises";
import { readPptx, writePptx } from "@pptx-glimpse/document";

const source = readPptx(await readFile("input.pptx"));
const output = writePptx(source);

await writeFile("round-trip.pptx", output);
```

The output is a `Uint8Array`.

## Structural preservation

`writePptx` targets structural round-trip preservation, not byte-for-byte equality. ZIP metadata,
XML attribute order, namespace-prefix placement, and defaulted OOXML values may differ. The writer
regenerates package bookkeeping such as content types and relationships, while media bytes,
unknown parts, and other unedited package material prefer the raw bytes retained by the reader.

Supported editing and authoring operations append edit records to the typed source model. These
records define dirty parts or presentation-topology operations. The writer validates the edits,
serializes only those dirty scopes, and copies untouched raw parts. This connects the model's two
representations:

1. Typed nodes provide the public inspection and editing surface.
2. Raw sidecars and package parts preserve material outside that typed surface.
3. Edit records identify where typed changes make the raw representation dirty.
4. The writer merges supported dirty edits and preserves the rest structurally.

If required preserved bytes are missing or a dirty edit cannot be safely serialized, `writePptx`
throws instead of silently rebuilding or dropping unsupported content.

## Responsibility boundary

The writer owns PPTX package serialization only. It does not compute effective inherited values,
render slides, choose fonts, or emit SVG/PNG. Use a [computed view](./computed-view.md) for resolved
document semantics. Upper layers may consume the written bytes, but the document writer has no
dependency on those consumers.

Import `writePptx` and `WritePptxOutput` from `@pptx-glimpse/document`. Dirty-scope serializers,
raw replacement mechanisms, topology patchers, and XML helpers are internal APIs.

See the [feature support matrix](./feature-support.md) for preservation guarantees by element type.
