---
"pptx-glimpse": patch
"@pptx-glimpse/document": minor
---

Unify new-content edit XML generation at edit time: `addTextBox` / `addConnector` now finalize their shape XML fragment on the edit record and derive the in-memory shape from it, and `addEmptySlideFromLayout` / `duplicateSlide` assign the new `p:sldId` numeric id at edit time. The writer no longer generates new-content XML and only applies insertion positions. The `addTextBox` / `addConnector` / `addEmptySlideFromLayout` / `duplicateSlide` edit record shapes changed accordingly.
