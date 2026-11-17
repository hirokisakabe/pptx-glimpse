---
"pptx-glimpse": patch
---

Render `headEnd` / `tailEnd` arrow markers on `<p:sp>` shapes. `renderConnector` already wired markers up, but `renderShape` did not, so a shape with line geometry and an `<a:ln>` carrying arrow endpoints was drawn without arrows and without a diagnostic. Multi-path custom geometry is excluded because markers are inheritable and would be drawn on every subpath.
