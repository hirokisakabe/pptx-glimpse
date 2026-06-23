---
"pptx-glimpse": major
---

Switch public SVG/PNG conversion defaults to the CleanDoc document path.

The CleanDoc path intentionally starts with a narrower rendering subset. Tables,
charts, SmartArt, groups, connectors, effects, and other raw shape-tree nodes can
be omitted from public conversion output until their CleanDoc support lands.
CJK text can emit `document-render.cjk-font-context-unsupported` while East Asian
and complex-script theme font context is still being moved into CleanDoc.
