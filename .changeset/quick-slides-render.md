---
"pptx-glimpse": patch
---

Editor command、undo、redo の適用時に、影響範囲を安全に特定できる場合は対象 slide だけを再レンダリングして SVG cache を更新するようにしました。特定できない場合は全 slide を再レンダリングします。
