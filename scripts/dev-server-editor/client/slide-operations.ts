export const DEV_EDITOR_SLIDE_OPERATIONS_SCRIPT = `    function duplicateSlide(index) {
      if (activeTextEditor) {
        setEditorMessage("Finish text editing before slide operations", true);
        return;
      }
      var slide = slides[index];
      if (!slide || !slide.handle) {
        setEditorMessage("Slide handle is unavailable", true);
        return;
      }
      postJson("/api/editor/command", {
        command: {
          kind: "duplicateSlide",
          handle: slide.handle
        }
      })
        .then(function (data) {
          applyEditorResponse(data, index + 1);
          setEditorMessage("Duplicated slide", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function deleteSlide(index) {
      if (activeTextEditor) {
        setEditorMessage("Finish text editing before slide operations", true);
        return;
      }
      if (slideCount <= 1) {
        setEditorMessage("Cannot delete the last slide", true);
        return;
      }
      var slide = slides[index];
      if (!slide || !slide.handle) {
        setEditorMessage("Slide handle is unavailable", true);
        return;
      }
      postJson("/api/editor/command", {
        command: {
          kind: "deleteSlide",
          handle: slide.handle
        }
      })
        .then(function (data) {
          applyEditorResponse(data, Math.min(index, (data.slides || slides).length - 1));
          setEditorMessage("Deleted slide", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

`;
