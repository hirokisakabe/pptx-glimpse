---
"pptx-glimpse": patch
---

Fix the direction of the `headEnd` arrow marker. With `orient="auto"` the marker's +x axis follows the path direction, which at the start vertex points into the line, so `marker-start` was rendered pointing inwards. The start marker is now mirrored horizontally and anchored with `refX="0"`, making it symmetric with `marker-end`.
