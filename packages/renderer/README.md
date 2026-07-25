# @pptx-glimpse/renderer

> **Private internal package.** This workspace is not published for direct installation and is not
> a supported public API. Install [`pptx-glimpse`](https://www.npmjs.com/package/pptx-glimpse)
> instead.

`@pptx-glimpse/renderer` owns the display-oriented model and the rendering machinery consumed by
the public `pptx-glimpse` package:

- renderer model contracts for slides, shapes, text, tables, charts, images, fills, lines, and
  effects;
- SVG generation and optional SVG-to-PNG conversion;
- font discovery, mapping, measurement, text wrapping, and text-to-path conversion;
- renderer-specific units, fallbacks, and warning collection.

The package accepts a render-ready model. It does not read or write PPTX packages, understand
`PptxSourceModel`, manage editor commands or history, or depend on
`@pptx-glimpse/document`/`@pptx-glimpse/editor`. The adapter in `packages/core` converts the
document computed view into the renderer model, keeping the dependency boundary one way.

Internal exports and model contracts may change without notice. Public consumers should use the
rendering and font APIs exported from
[`pptx-glimpse`](https://github.com/hirokisakabe/pptx-glimpse/blob/main/packages/core/README.md).

See the [root README](https://github.com/hirokisakabe/pptx-glimpse#readme) for public package
choices and the [architecture overview](../../docs/architecture/overview.md) for repository
package boundaries.

## License

MIT
