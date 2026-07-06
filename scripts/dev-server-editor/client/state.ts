interface DevEditorInitialStateScriptOptions {
  readonly slides: readonly unknown[];
  readonly slideCount: number;
  readonly emuPerPixel: number;
}

export function createDevEditorInitialStateScript(
  options: DevEditorInitialStateScriptOptions,
): string {
  return `    var currentIndex = 0;
    var slideCount = ${String(options.slideCount)};
    var slides = ${jsonForScript(options.slides)};
    var editorHistory = { canUndo: false, canRedo: false, undoDepth: 0, redoDepth: 0 };
    var textRunOptions = [];
    var shapeOptions = [];
    var selectedShapeKey = null;
    var selectedShape = null;
    var dragState = null;
    var activeTextEditor = null;
    var shapeRequestId = 0;
    var EMU_PER_PIXEL = ${String(options.emuPerPixel)};

`;
}

export const DEV_EDITOR_CORE_SCRIPT = `    function selectSlide(index) {
      if (activeTextEditor) {
        commitTextEditor()
          .then(function () {
            selectSlide(index);
          })
          .catch(function () {
            // Keep the editor open; commitTextEditor already reported the failure.
          });
        return;
      }
      var slideChanged = currentIndex !== index;
      currentIndex = index;
      shapeOptions = [];
      textRunOptions = [];
      selectedShape = null;
      if (slideChanged) selectedShapeKey = null;
      syncImageReplacementInput();
      var thumbs = document.querySelectorAll(".thumbnail");
      for (var i = 0; i < thumbs.length; i++) {
        if (i === index) {
          thumbs[i].classList.add("active");
        } else {
          thumbs[i].classList.remove("active");
        }
      }
      document.getElementById("slide-container").innerHTML =
        slides[index] ? slides[index].svg : "<p>No slides</p>";
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
      document.getElementById("info").textContent =
        "Slide " + (index + 1) + " / " + slideCount;
      renderSelectionOverlay();
      loadShapeOptions(index + 1);
    }

    function renderThumbnails() {
      var sidebar = document.getElementById("sidebar");
      sidebar.innerHTML = slides
        .map(function (s, i) {
          return '<div class="thumbnail' + (i === currentIndex ? ' active' : '') +
            '" data-index="' + i + '">' +
            '<div class="thumb-label">' +
            '<span class="thumb-title">Slide ' + s.slideNumber + '</span>' +
            '<span class="thumb-actions">' +
            '<button class="thumb-action" data-testid="duplicate-slide-' + i + '" data-action="duplicate" data-index="' + i + '" type="button" title="Duplicate slide">D</button>' +
            '<button class="thumb-action" data-testid="delete-slide-' + i + '" data-action="delete" data-index="' + i + '" type="button" title="Delete slide"' +
            (slideCount <= 1 ? ' disabled' : '') + '>X</button>' +
            '</span>' +
            '</div>' +
            '<div class="thumb-svg">' + s.svg + '</div>' +
            '</div>';
        })
        .join("");

      var thumbs = document.querySelectorAll(".thumbnail");
      for (var i = 0; i < thumbs.length; i++) {
        (function (idx) {
          thumbs[idx].addEventListener("click", function () {
            selectSlide(idx);
          });
        })(i);
      }

      Array.prototype.forEach.call(sidebar.querySelectorAll("[data-action]"), function (button) {
        button.addEventListener("mousedown", function (event) {
          event.preventDefault();
        });
        button.addEventListener("click", function (event) {
          event.preventDefault();
          event.stopPropagation();
          var index = Number(button.getAttribute("data-index") || "-1");
          var action = button.getAttribute("data-action");
          if (action === "duplicate") duplicateSlide(index);
          if (action === "delete") deleteSlide(index);
        });
      });
    }

    function updateHistory(history) {
      editorHistory = history || editorHistory;
      document.getElementById("undo-button").disabled = !editorHistory.canUndo;
      document.getElementById("redo-button").disabled = !editorHistory.canRedo;
      updateSelectedShapeActions();
    }

    function updateSelectedShapeActions() {
      document.getElementById("delete-shape-button").disabled =
        activeTextEditor || !selectedShape || !selectedShape.editableDelete;
    }

    function shapeKey(shape) {
      return shape && shape.handle ? handleKey(shape.handle) : "";
    }

    function handleKey(handle) {
      return [
        handle.partPath || "",
        handle.nodeId || "",
        handle.relationshipId || "",
        handle.orderingSlot == null ? "" : String(handle.orderingSlot)
      ].join("\\u0000");
    }

    function syncSelection(selection) {
      selectedShapeKey = selection && selection.shapeHandle ? handleKey(selection.shapeHandle) : selectedShapeKey;
      selectedShape = null;
      if (selectedShapeKey) {
        for (var i = 0; i < shapeOptions.length; i++) {
          if (shapeKey(shapeOptions[i]) === selectedShapeKey && shapeOptions[i].bounds) {
            selectedShape = cloneShape(shapeOptions[i]);
            break;
          }
        }
      }
      if (!selectedShape) selectedShapeKey = null;
      syncImageReplacementInput();
      renderSelectionOverlay();
      updateSelectedShapeActions();
    }

    function cloneShape(shape) {
      return {
        id: shape.id,
        kind: shape.kind,
        name: shape.name,
        handle: shape.handle,
        editableDelete: shape.editableDelete === true,
        editableTransform: shape.editableTransform,
        editableTextBody: shape.editableTextBody,
        editableImageReplacement: shape.editableImageReplacement,
        bounds: {
          x: shape.bounds.x,
          y: shape.bounds.y,
          width: shape.bounds.width,
          height: shape.bounds.height
        }
      };
    }

    function setEditorMessage(message, isError) {
      var element = document.getElementById("editor-message");
      element.textContent = message;
      element.style.color = isError ? "#fca5a5" : "#94a3b8";
    }

`;

export const DEV_EDITOR_RESPONSE_SCRIPT = `    function applyEditorResponse(data, preferredIndex) {
      slides = data.slides || slides;
      slideCount = slides.length;
      updateHistory(data.history);
      if (data.selection && data.selection.shapeHandle) {
        selectedShapeKey = handleKey(data.selection.shapeHandle);
      }
      renderThumbnails();
      var nextIndex = preferredIndex == null ? currentIndex : preferredIndex;
      selectSlide(Math.min(Math.max(nextIndex, 0), Math.max(slides.length - 1, 0)));
    }

`;

function jsonForScript(value: unknown): string {
  return JSON.stringify(value).replace(/</g, "\\u003c");
}
