---
"@pptx-glimpse/document": patch
"@pptx-glimpse/editor": patch
"pptx-glimpse": patch
---

共有 media part の画像置換を copy-on-write 化し、選択した picture だけを新しい media と relationship に切り替える。
