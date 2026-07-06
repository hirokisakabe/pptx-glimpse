export const DEV_EDITOR_SELECTION_SCRIPT = `    function renderSelectionOverlay() {
      var container = document.getElementById("slide-container");
      var renderedSvg = container.querySelector("svg:not(#selection-overlay)");
      var existing = document.getElementById("selection-overlay");
      if (existing) existing.remove();
      if (!renderedSvg) return;

      var viewBox = renderedSvg.getAttribute("viewBox") || "0 0 960 540";
      var overlay = document.createElementNS("http://www.w3.org/2000/svg", "svg");
      overlay.setAttribute("id", "selection-overlay");
      overlay.setAttribute("viewBox", viewBox);
      overlay.setAttribute("data-testid", "selection-overlay");

      shapeOptions.forEach(function (shape, index) {
        if (!shape.bounds) return;
        var hit = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        hit.setAttribute("class", "shape-hit-area");
        hit.setAttribute("data-testid", "shape-hit-area");
        hit.setAttribute("data-shape-index", String(index));
        setRectAttributes(hit, shape.bounds);
        hit.addEventListener("pointerdown", function (event) {
          selectShape(shape, event);
        });
        hit.addEventListener("mousedown", function (event) {
          if (activeTextEditor) return;
          if (event.detail >= 2 && shape.editableTextBody) {
            event.preventDefault();
            openTextEditor(shape);
          }
        });
        hit.addEventListener("dblclick", function (event) {
          event.preventDefault();
          event.stopPropagation();
          openTextEditor(shape);
        });
        overlay.appendChild(hit);
      });

      if (selectedShape && selectedShape.bounds) {
        var box = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        box.setAttribute("class", "selection-box");
        box.setAttribute("data-testid", "selection-box");
        setRectAttributes(box, selectedShape.bounds);
        overlay.appendChild(box);

        if (selectedShape.editableTransform) {
          ["nw", "ne", "sw", "se"].forEach(function (handle) {
            var point = handlePoint(selectedShape.bounds, handle);
            var rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
            rect.setAttribute("class", "selection-handle");
            rect.setAttribute("data-testid", "selection-handle-" + handle);
            rect.setAttribute("data-handle", handle);
            rect.setAttribute("x", String(point.x - 4));
            rect.setAttribute("y", String(point.y - 4));
            rect.setAttribute("width", "8");
            rect.setAttribute("height", "8");
            rect.addEventListener("pointerdown", function (event) {
              event.preventDefault();
              event.stopPropagation();
              beginDrag("resize", handle, event);
            });
            overlay.appendChild(rect);
          });
        }
      }

      container.appendChild(overlay);
      positionActiveTextEditor();
      updateSelectedShapeActions();
    }

    function setRectAttributes(rect, bounds) {
      rect.setAttribute("x", String(bounds.x));
      rect.setAttribute("y", String(bounds.y));
      rect.setAttribute("width", String(bounds.width));
      rect.setAttribute("height", String(bounds.height));
    }

    function handlePoint(bounds, handle) {
      var right = bounds.x + bounds.width;
      var bottom = bounds.y + bounds.height;
      if (handle === "nw") return { x: bounds.x, y: bounds.y };
      if (handle === "ne") return { x: right, y: bounds.y };
      if (handle === "sw") return { x: bounds.x, y: bottom };
      return { x: right, y: bottom };
    }

    function selectShape(shape, event) {
      if (activeTextEditor) {
        commitTextEditor()
          .then(function () {
            selectShapeAfterCommit(shape);
          })
          .catch(function () {});
        return;
      }
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      syncImageReplacementInput();
      if (shape.editableTransform) beginDrag("move", null, event);
      window.setTimeout(renderSelectionOverlay, 0);
      postJson("/api/editor/select", { handle: shape.handle })
        .then(function (data) {
          updateHistory(data.history);
          if (data.selection && data.selection.shapeHandle) {
            selectedShapeKey = handleKey(data.selection.shapeHandle);
          }
          syncImageReplacementInput();
          renderSelectionOverlay();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function selectShapeAfterCommit(shape) {
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      syncImageReplacementInput();
      renderSelectionOverlay();
      postJson("/api/editor/select", { handle: shape.handle })
        .then(function (data) {
          updateHistory(data.history);
          if (data.selection && data.selection.shapeHandle) {
            selectedShapeKey = handleKey(data.selection.shapeHandle);
          }
          syncImageReplacementInput();
          renderSelectionOverlay();
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

`;

export const DEV_EDITOR_TRANSFORM_SCRIPT = `    function beginDrag(kind, handle, event) {
      if (!selectedShape || !selectedShape.bounds) return;
      var overlay = document.getElementById("selection-overlay");
      if (!overlay) return;
      var point = eventPoint(overlay, event);
      if (!point) return;
      dragState = {
        kind: kind,
        handle: handle,
        pointerId: event.pointerId,
        startPoint: point,
        startBounds: {
          x: selectedShape.bounds.x,
          y: selectedShape.bounds.y,
          width: selectedShape.bounds.width,
          height: selectedShape.bounds.height
        }
      };
      try {
        overlay.setPointerCapture(event.pointerId);
      } catch (_error) {
        // Pointer capture can be unavailable after a fast re-render; window listeners still cover drag.
      }
      window.addEventListener("pointermove", updateDrag);
      window.addEventListener("pointerup", finishDrag);
      window.addEventListener("pointercancel", cancelDrag);
    }

    function updateDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      var overlay = document.getElementById("selection-overlay");
      if (!overlay) return;
      var point = eventPoint(overlay, event);
      if (!point) return;
      var dx = point.x - dragState.startPoint.x;
      var dy = point.y - dragState.startPoint.y;
      selectedShape.bounds =
        dragState.kind === "move"
          ? movedBounds(dragState.startBounds, dx, dy)
          : resizedBounds(dragState.startBounds, dragState.handle, dx, dy);
      renderSelectionOverlay();
    }

    function finishDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      var nextBounds = selectedShape.bounds;
      var previousDrag = dragState;
      dragState = null;
      detachDragListeners();
      applyShapeBoundsEdit(previousDrag.startBounds, nextBounds).catch(function (err) {
        setEditorMessage(err.message, true);
        selectedShape.bounds = previousDrag.startBounds;
        renderSelectionOverlay();
      });
    }

    function cancelDrag(event) {
      if (!dragState || event.pointerId !== dragState.pointerId || !selectedShape) return;
      selectedShape.bounds = dragState.startBounds;
      dragState = null;
      detachDragListeners();
      renderSelectionOverlay();
    }

    function detachDragListeners() {
      window.removeEventListener("pointermove", updateDrag);
      window.removeEventListener("pointerup", finishDrag);
      window.removeEventListener("pointercancel", cancelDrag);
    }

    function eventPoint(svg, event) {
      var matrix = svg.getScreenCTM();
      if (!matrix) return null;
      var point = svg.createSVGPoint();
      point.x = event.clientX;
      point.y = event.clientY;
      return point.matrixTransform(matrix.inverse());
    }

    function movedBounds(bounds, dx, dy) {
      return {
        x: bounds.x + dx,
        y: bounds.y + dy,
        width: bounds.width,
        height: bounds.height
      };
    }

    function resizedBounds(bounds, handle, dx, dy) {
      var minSize = 8;
      var right = bounds.x + bounds.width;
      var bottom = bounds.y + bounds.height;
      var next = {
        x: bounds.x,
        y: bounds.y,
        width: bounds.width,
        height: bounds.height
      };
      if (handle === "nw" || handle === "sw") {
        next.x = Math.min(bounds.x + dx, right - minSize);
        next.width = right - next.x;
      }
      if (handle === "ne" || handle === "se") {
        next.width = Math.max(minSize, bounds.width + dx);
      }
      if (handle === "nw" || handle === "ne") {
        next.y = Math.min(bounds.y + dy, bottom - minSize);
        next.height = bottom - next.y;
      }
      if (handle === "sw" || handle === "se") {
        next.height = Math.max(minSize, bounds.height + dy);
      }
      return next;
    }

    function applyShapeBoundsEdit(startBounds, nextBounds) {
      if (!selectedShape || !selectedShape.handle || !selectedShape.editableTransform) {
        return Promise.resolve();
      }
      var handle = selectedShape.handle;
      var changed =
        Math.round(startBounds.x) !== Math.round(nextBounds.x) ||
        Math.round(startBounds.y) !== Math.round(nextBounds.y) ||
        Math.round(startBounds.width) !== Math.round(nextBounds.width) ||
        Math.round(startBounds.height) !== Math.round(nextBounds.height);
      if (!changed) return Promise.resolve();
      return postJson("/api/editor/command", {
        command: {
          kind: "setShapeTransform",
          handle: handle,
          offsetX: Math.round(nextBounds.x * EMU_PER_PIXEL),
          offsetY: Math.round(nextBounds.y * EMU_PER_PIXEL),
          width: Math.round(nextBounds.width * EMU_PER_PIXEL),
          height: Math.round(nextBounds.height * EMU_PER_PIXEL)
        }
      }).then(function (data) {
        if (data) {
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        }
      });
    }

`;
