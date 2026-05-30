---
"pptx-glimpse": patch
---

OOXML Strict 形式のSmartArt（Diagram）URIに対応。Transitional URI（`http://schemas.openxmlformats.org/drawingml/2006/diagram`）のみを直接比較していた判定を、Strict URI（`http://purl.oclc.org/ooxml/drawingml/diagram`）にも対応した allowlist 判定に変更。
