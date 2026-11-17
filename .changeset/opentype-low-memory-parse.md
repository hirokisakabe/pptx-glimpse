---
"pptx-glimpse": patch
---

Parse font buffers with opentype.js `lowMemory` mode. Glyph outlines and metrics are now decoded on first use instead of materializing every glyph at parse time, which cuts retained heap for a ~65k-glyph CJK font from ~220 MiB to ~30 MiB per font. Rendering output (text measurement, line wrapping and path outlines) is unchanged. Applies to both caller-supplied `fonts` and system font scanning.
