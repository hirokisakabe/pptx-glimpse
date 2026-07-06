export const DEV_EDITOR_BOOTSTRAP_SCRIPT = `    renderThumbnails();
    updateHistory(editorHistory);
    loadShapeOptions(1);

    document.getElementById("text-run-select").addEventListener("change", syncTextInput);
    document.getElementById("image-replacement-input").addEventListener("change", function () {
      var input = document.getElementById("image-replacement-input");
      var file = input && input.files ? input.files[0] : null;
      if (!file) return;
      replaceSelectedImage(file)
        .catch(function (err) {
          setEditorMessage(err.message, true);
        })
        .finally(function () {
          syncImageReplacementInput();
        });
    });
    document.getElementById("apply-text-button").addEventListener("click", function () {
      var option = textRunOptions[Number(document.getElementById("text-run-select").value || "0")];
      if (!option) return;
      postJson("/api/editor/command", {
        command: {
          kind: "replaceTextRunPlainText",
          handle: option.handle,
          text: document.getElementById("text-run-input").value
        }
      })
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("add-text-box-button").addEventListener("click", addTextBox);
    document.getElementById("delete-shape-button").addEventListener("click", deleteSelectedShape);
    document.getElementById("undo-button").addEventListener("click", function () {
      postJson("/api/editor/undo")
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Undone", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("redo-button").addEventListener("click", function () {
      postJson("/api/editor/redo")
        .then(function (data) {
          applyEditorResponse(data);
          setEditorMessage("Redone", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });
    document.getElementById("save-button").addEventListener("click", function () {
      postJson("/api/editor/save")
        .then(function (data) {
          updateHistory(data.history);
          setEditorMessage("Saved: " + data.path, false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    });

    // WebSocket for live reload
    function connect() {
      var ws = new WebSocket("ws://" + location.host);
      var status = document.getElementById("status");

      ws.onopen = function () {
        status.textContent = "Connected";
        status.className = "";
      };
      ws.onclose = function () {
        status.textContent = "Disconnected - reconnecting...";
        status.className = "error";
        setTimeout(connect, 2000);
      };
      ws.onmessage = function (event) {
        var data = JSON.parse(event.data);
        if (data.type === "rendering") {
          status.textContent = "Re-rendering...";
          status.className = "rendering";
        } else if (data.type === "reload") {
          status.textContent = "Updating...";
          status.className = "rendering";
          location.reload();
        } else if (data.type === "error") {
          status.textContent = "Error: " + data.message;
          status.className = "error";
        }
      };
    }
    connect();

    // Initial: resize the main SVG
    (function () {
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
    })();

    // Keyboard navigation
    document.addEventListener("keydown", function (e) {
      if (e.key === "Delete" && !isTypingTarget(e.target)) {
        if (selectedShape && selectedShape.editableDelete) {
          e.preventDefault();
          deleteSelectedShape();
        }
        return;
      }
      if (e.key === "ArrowLeft" && currentIndex > 0) selectSlide(currentIndex - 1);
      if (e.key === "ArrowRight" && currentIndex < slideCount - 1)
        selectSlide(currentIndex + 1);
    });
`;
