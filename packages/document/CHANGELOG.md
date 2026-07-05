# @pptx-glimpse/document

## 0.2.0

### Minor Changes

- f627f71: Add headless text run formatting edits for bold, italic, underline, font size, direct sRGB color, and latin typeface.
- 3d61817: Add headless image media replacement for existing pic shapes, limited to same-format media byte swaps.
- 020f949: Add PptxSourceModel editing and writer support for inserting native PowerPoint connector shapes.
- 8904a5c: Expose `p:sldLayout@show` as `SourceSlideLayout.show` for detecting hidden slide layouts.

### Patch Changes

- f0136a9: Add headless slide duplicate/delete editing support with package relationship, content type, and ID management.
- c57532b: Add headless empty slide creation from a slide layout, including writer package bookkeeping and editor-core command support.
- 69ae720: Add PptxSourceModel writer operations for adding text boxes and deleting top-level shapes.

## 0.1.0

### Minor Changes

- b32b8a8: Publish the document package as an installable public 0.x package with README guidance.
