# Document Feature Support

This matrix describes the public `@pptx-glimpse/document` surface. It is intentionally
separate from SVG/PNG rendering fidelity: a value can be available in the document model
without being rendered by `pptx-glimpse`.

## Status legend

- **S — supported**: a typed public workflow exists for this stage and is covered by an
  implementation test. This does not mean that every OOXML variant is supported.
- **△ — partial**: typed support exists, but only for the constraints described below.
- **P — preserved**: there is no typed workflow for this stage, but unchanged source material is
  retained as a raw sidecar or raw package part for structural round trips.
- **— — unsupported**: neither a typed workflow nor a preservation guarantee is currently
  documented and tested.

“Existing edit” means changing or deleting an element already present in an input PPTX. Adding a
new element to a slide loaded from an existing PPTX uses the same authoring operations as the
from-scratch writer and is therefore represented in the “from-scratch writer” column instead.

“Round-trip preservation” is structural, not byte-for-byte. `P` in that column means no-edit,
opaque preservation only; `S` means supported edits or typed authoring are also written and
reread in tests. See [Read and Write Round Trip](../README.md#read-and-write-round-trip).

## Matrix

| PowerPoint element      | Reader | Computed view | From-scratch writer | Existing edit | Round-trip preservation |
| ----------------------- | :----: | :-----------: | :-----------------: | :-----------: | :---------------------: |
| Text                    |   S    |       S       |          S          |       △       |            S            |
| Shape                   |   △    |       △       |          △          |       △       |            S            |
| Picture                 |   △    |       S       |          △          |       △       |            S            |
| Connector               |   △    |       S       |          △          |       △       |            S            |
| Table                   |   △    |       S       |          △          |       —       |            S            |
| Chart                   |   △    |       △       |          △          |       —       |            S            |
| Group                   |   △    |       S       |          —          |       —       |            P            |
| Background              |   △    |       △       |          △          |       △       |            S            |
| Master / layout / theme |   △    |       △       |          △          |       —       |            S            |

## Constraints and evidence

| Element                 | Current boundary                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Text                    | The reader and computed view expose typed paragraphs, runs, body properties, and the supported style cascade. Authoring covers text boxes and shape/table text with the typed formatting inputs shown in [From-Scratch PPTX](../README.md#from-scratch-pptx). Existing edits are limited to plain run/paragraph text, selected run properties, and paragraph alignment/level/bullet properties; they are exercised in [writer edit tests](../src/writer/write-pptx.test.ts).                                                                                                                                                                      |
| Shape                   | Common transforms, preset/custom geometry, fills, outlines, text, and selected effects are typed; unmodeled DrawingML remains raw. Authoring accepts the inputs defined by `AddShapeInput`, while existing edits are limited to top-level transform, fill, outline, and deletion. Nested group children and `mc:AlternateContent` fallback nodes cannot be edited through those helpers. See [shape editing](../src/source/shape-editing.ts) and [typed reader tests](../src/reader/slide-reader.test.ts).                                                                                                                                        |
| Picture                 | Typed reading/computation resolves embedded media and supported picture/effect properties. Authoring is limited to PNG/JPEG bytes and the typed crop/effect inputs. Existing replacement is a same-format media-byte swap and can affect other pictures that share the media part. See [picture authoring](../src/source/picture-authoring.ts) and [image replacement](../src/source/image-replacement.ts).                                                                                                                                                                                                                                       |
| Connector               | Typed source/computed nodes retain supported transforms, geometry, outlines, connection sites, and arrow endpoints. Authoring uses the connector presets, endpoint forms, and horizontal/vertical transform flips accepted by `AddConnectorInput`. Existing connectors can use the supported outline and delete operations, but there is no endpoint/geometry editing API. See [shape editing](../src/source/shape-editing.ts) and [writer edit tests](../src/writer/write-pptx.test.ts).                                                                                                                                                         |
| Table                   | Native tables have typed rows, columns, cells, text, fills, borders, margins, merges, and hyperlinks for the implemented subset. Authoring creates new tables, but existing table cells or structure cannot be edited through the public API. The `S` preservation status covers authored-table write/reread tests; existing table XML is only claimed as opaque no-edit preservation. See [table authoring](../src/source/table-authoring.ts) and [reader tests](../src/reader/slide-reader.test.ts).                                                                                                                                            |
| Chart                   | The reader/computed view supports typed chart relationships and data projections for bar, line, pie, doughnut, area, scatter, bubble, radar, stock, surface, and of-pie charts; other chart XML remains preserved package material. Authoring is narrower: `addChart` creates bar, line, pie, area, doughnut, or radar charts with an editable embedded workbook and the typed formatting implemented in [chart authoring](../src/source/chart-authoring.ts). The `S` preservation status covers authored-chart write/reread tests; existing chart data or formatting cannot be edited.                                                           |
| Group                   | Group transforms, children, fills, and selected effects are recursively projected into the typed source/computed models. There is no public group authoring or editing operation. Existing group content is only claimed as opaque no-edit preservation. See [computed view types](../src/computed/pptx-computed-view.ts) and [reader tests](../src/reader/slide-reader.test.ts).                                                                                                                                                                                                                                                                 |
| Background              | Slide/layout/master fallback and the implemented fill/style-reference subset are typed; unmodeled backgrounds stay raw. Authoring/editing is limited to replacing a slide background with the supported solid, linear/radial gradient, PNG, or JPEG forms. Master/layout backgrounds can be initialized by `createPptx`, but there is no existing master/layout background edit. See [background authoring](../src/source/slide-background-authoring.ts) and [master/layout authoring](../README.md#slide-master-and-layout-authoring).                                                                                                           |
| Master / layout / theme | The reader and computed view follow the slide → layout → master → theme chain and expose the implemented background, placeholder, color-map, color/font scheme, and text-style semantics. From-scratch creation has one configurable initial master/layout and a generated theme; it does not provide arbitrary theme authoring or an API for additional masters/layouts. The `S` preservation status covers that authored template chain; existing template/theme parts are not editable and are retained as opaque package material. See [computed view](../src/computed/create-computed-view.ts) and [builder](../src/builder/create-pptx.ts). |

The raw preservation hooks and dependency boundaries behind the final column are documented on
[`PptxSourceModel`](../src/source/pptx-source-model.ts) and in the
[`writePptx` implementation](../src/writer/write-pptx.ts). If a capability is not supported by a
public root export and confirmed by implementation tests, this table must not mark it `S`.
