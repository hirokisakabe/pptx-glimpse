export const DEV_EDITOR_SHAPE_OPTIONS_SCRIPT = `    function loadShapeOptions(slideNumber) {
      var requestId = ++shapeRequestId;
      fetch("/api/editor/shapes?slide=" + slideNumber)
        .then(function (res) {
          if (!res.ok) throw new Error("shape request failed");
          return res.json();
        })
        .then(function (data) {
          if (requestId !== shapeRequestId || slideNumber !== currentIndex + 1) return;
          shapeOptions = (data.shapes || []).filter(function (shape) {
            return shape.handle && shape.bounds && (shape.editableTransform || shape.editableImageReplacement);
          });
          textRunOptions = [];
          data.shapes.forEach(function (shape) {
            (shape.textRuns || []).forEach(function (run, index) {
              textRunOptions.push({
                label: (shape.name || shape.id) + " / " + (index + 1),
                text: run.text,
                handle: run.handle
              });
            });
          });
          var select = document.getElementById("text-run-select");
          select.innerHTML = textRunOptions
            .map(function (option, index) {
              return '<option value="' + index + '">' + escapeHtmlClient(option.label) + '</option>';
            })
            .join("");
          select.disabled = textRunOptions.length === 0;
          document.getElementById("apply-text-button").disabled = textRunOptions.length === 0;
          syncTextInput();
          syncSelection();
          syncImageReplacementInput();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function syncTextInput() {
      var select = document.getElementById("text-run-select");
      var input = document.getElementById("text-run-input");
      var option = textRunOptions[Number(select.value || "0")];
      input.value = option ? option.text : "";
      input.disabled = !option;
    }

`;

export const DEV_EDITOR_PANEL_COMMANDS_SCRIPT = `    function addTextBox() {
      if (activeTextEditor) {
        commitTextEditor()
          .then(addTextBox)
          .catch(function () {});
        return;
      }
      postJson("/api/editor/add-text-box", { slide: currentIndex + 1 })
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Text box added", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function addConnector() {
      if (activeTextEditor) {
        commitTextEditor()
          .then(addConnector)
          .catch(function () {});
        return;
      }
      postJson("/api/editor/add-connector", { slide: currentIndex + 1 })
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Connector added", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function deleteSelectedShape() {
      if (activeTextEditor || !selectedShape || !selectedShape.handle || !selectedShape.editableDelete) {
        return;
      }
      postJson("/api/editor/command", {
        command: {
          kind: "deleteShape",
          handle: selectedShape.handle
        }
      })
        .then(function (data) {
          selectedShape = null;
          selectedShapeKey = null;
          applyEditorResponse(data);
          setEditorMessage("Deleted", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function isTypingTarget(target) {
      return target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target instanceof HTMLSelectElement ||
        (target instanceof HTMLElement && target.isContentEditable);
    }

`;

interface DevEditorImageReplacementScriptOptions {
  readonly maxImageReplacementBytes: number;
}

export function createDevEditorImageReplacementScript(
  options: DevEditorImageReplacementScriptOptions,
): string {
  return `    function syncImageReplacementInput() {
      var input = document.getElementById("image-replacement-input");
      if (!input) return;
      var replacement = selectedShape && selectedShape.editableImageReplacement;
      input.disabled = !replacement;
      input.value = "";
      if (replacement) {
        input.setAttribute("accept", replacement.accept);
        input.title =
          "Replace " + replacement.mediaPartPath + " with another " + replacement.contentType + " image";
      } else {
        input.removeAttribute("accept");
        input.title = "Select an image shape to replace its media";
      }
    }

    function replaceSelectedImage(file) {
      if (!selectedShape || !selectedShape.handle || !selectedShape.editableImageReplacement) {
        setEditorMessage("Select an image shape before replacing media.", true);
        return Promise.resolve();
      }
      if (file.size > ${String(options.maxImageReplacementBytes)}) {
        setEditorMessage("Replacement image is too large.", true);
        return Promise.resolve();
      }
      var handle = selectedShape.handle;
      var input = document.getElementById("image-replacement-input");
      if (input) input.disabled = true;
      return file.arrayBuffer()
        .then(function (buffer) {
          return postJson("/api/editor/command", {
            command: {
              kind: "replaceImage",
              handle: handle,
              bytes: Array.from(new Uint8Array(buffer))
            }
          });
        })
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage(imageReplacementMessage(data), false);
        });
    }

    function imageReplacementMessage(data) {
      var warnings = (data && data.warnings) || [];
      var shared = warnings.find(function (warning) {
        return warning.code === "shared-media-part";
      });
      if (!shared) return "Image replaced";
      return (
        "Image replaced; shared media part affects " +
        shared.referenceCount +
        " pictures: " +
        shared.mediaPartPath
      );
    }

`;
}
