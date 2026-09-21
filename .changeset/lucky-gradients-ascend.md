---
"pptx-glimpse": patch
---

Sort gradient stops by position when adapting a computed view to the renderer model. OOXML does not require `<a:gs pos>` to be listed in ascending order, but SVG clamps a gradient `<stop>` whose offset is below the previous one, which rendered such gradients with a wrong colour band.
