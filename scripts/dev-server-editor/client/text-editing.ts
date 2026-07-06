export const DEV_EDITOR_TEXT_EDITING_SCRIPT = `    function openTextEditor(shape) {
      if (!shape || !shape.handle || !shape.bounds || !shape.editableTextBody) return;
      if (activeTextEditor) return;
      if (dragState) {
        dragState = null;
        detachDragListeners();
      }
      closeTextEditor();
      selectedShapeKey = shapeKey(shape);
      selectedShape = cloneShape(shape);
      renderSelectionOverlay();

      var container = document.getElementById("slide-container");
      var overlay = document.createElement("div");
      overlay.id = "text-editor-overlay";
      overlay.setAttribute("data-testid", "text-editor-overlay");
      overlay.dataset.shapeKey = selectedShapeKey || "";
      overlay.appendChild(createTextRunFormatToolbar());
      overlay.appendChild(createTextEditorContent(shape.editableTextBody.docJson));

      var actions = document.createElement("div");
      actions.className = "text-editor-actions";
      var done = document.createElement("button");
      done.type = "button";
      done.textContent = "Done";
      done.setAttribute("data-testid", "text-editor-done");
      done.addEventListener("click", function () {
        commitTextEditor().catch(function () {});
      });
      actions.appendChild(done);
      overlay.appendChild(actions);

      overlay.addEventListener("keydown", function (event) {
        if (event.isComposing || event.keyCode === 229) return;
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          commitTextEditor().catch(function () {});
        }
        if (event.key === "Escape") {
          event.preventDefault();
          closeTextEditor();
        }
      });
      overlay.addEventListener("focusout", function (event) {
        window.setTimeout(function () {
          if (!activeTextEditor) return;
          if (event.relatedTarget && activeTextEditor.element.contains(event.relatedTarget)) return;
          commitTextEditor().catch(function () {});
        }, 0);
      });

      container.appendChild(overlay);
      activeTextEditor = {
        element: overlay,
        shape: cloneShape(shape),
        originalDocJson: shape.editableTextBody.docJson,
        selectedRunElement: null,
        committing: false,
        commitPromise: null
      };
      updateSelectedShapeActions();
      positionActiveTextEditor();
      var firstRun = overlay.querySelector(".text-editor-run");
      if (firstRun) {
        setActiveTextRunElement(firstRun);
        firstRun.focus();
        selectElementContents(firstRun);
      }
    }

    function createTextRunFormatToolbar() {
      var toolbar = document.createElement("div");
      toolbar.className = "text-run-format-toolbar";
      toolbar.setAttribute("data-testid", "text-run-format-toolbar");

      [["bold", "B"], ["italic", "I"], ["underline", "U"]].forEach(function (item) {
        var button = document.createElement("button");
        button.type = "button";
        button.textContent = item[1];
        button.dataset.property = item[0];
        button.setAttribute("data-testid", "text-run-format-" + item[0]);
        button.setAttribute("aria-pressed", "false");
        button.addEventListener("click", function () {
          toggleActiveTextRunBooleanProperty(item[0]);
        });
        toolbar.appendChild(button);
      });

      var size = document.createElement("input");
      size.type = "number";
      size.min = "1";
      size.step = "0.5";
      size.placeholder = "pt";
      size.setAttribute("data-testid", "text-run-format-font-size");
      size.addEventListener("change", function () {
        var value = Number(size.value);
        if (size.value.trim() === "") {
          applyActiveTextRunPropertyClear(["fontSize"]);
        } else if (Number.isFinite(value) && value > 0) {
          applyActiveTextRunPropertySet({ fontSize: value });
        } else {
          setEditorMessage("font size must be positive", true);
        }
      });
      toolbar.appendChild(size);

      var clearSize = clearButton("font-size", function () {
        applyActiveTextRunPropertyClear(["fontSize"]);
      });
      toolbar.appendChild(clearSize);

      var color = document.createElement("input");
      color.type = "color";
      color.setAttribute("data-testid", "text-run-format-color");
      color.addEventListener("change", function () {
        applyActiveTextRunPropertySet({
          color: { kind: "srgb", hex: color.value.replace(/^#/, "").toUpperCase() }
        });
      });
      toolbar.appendChild(color);

      var clearColor = clearButton("color", function () {
        applyActiveTextRunPropertyClear(["color"]);
      });
      toolbar.appendChild(clearColor);

      var typeface = document.createElement("input");
      typeface.type = "text";
      typeface.placeholder = "Font";
      typeface.setAttribute("data-testid", "text-run-format-typeface");
      typeface.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
          event.preventDefault();
          applyTypefaceInput(typeface);
        }
      });
      typeface.addEventListener("change", function () {
        applyTypefaceInput(typeface);
      });
      toolbar.appendChild(typeface);

      var clearTypeface = clearButton("typeface", function () {
        applyActiveTextRunPropertyClear(["typeface"]);
      });
      toolbar.appendChild(clearTypeface);

      window.setTimeout(refreshTextRunFormatToolbar, 0);
      return toolbar;
    }

    function clearButton(name, onClick) {
      var button = document.createElement("button");
      button.type = "button";
      button.textContent = "x";
      button.setAttribute("data-testid", "text-run-format-clear-" + name);
      button.addEventListener("click", onClick);
      return button;
    }

    function applyTypefaceInput(input) {
      var value = input.value.trim();
      if (value.length === 0) {
        applyActiveTextRunPropertyClear(["typeface"]);
        return;
      }
      applyActiveTextRunPropertySet({ typeface: value });
    }

    function setActiveTextRunElement(element) {
      if (!activeTextEditor) return;
      activeTextEditor.selectedRunElement = element;
      refreshTextRunFormatToolbar();
    }

    function activeTextRunInfo() {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return null;
      var element = activeTextEditor.selectedRunElement;
      var paragraphIndex = Number(element.dataset.paragraphIndex || "-1");
      var runIndex = Number(element.dataset.runIndex || "-1");
      var paragraph = (activeTextEditor.originalDocJson.content || [])[paragraphIndex];
      var textNode = paragraph && (paragraph.content || [])[runIndex];
      var mark = textNode && (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      var attrs = mark && mark.attrs ? mark.attrs : {};
      return {
        handle: attrs.handle || null,
        properties: attrs.properties || {}
      };
    }

    function refreshTextRunFormatToolbar() {
      if (!activeTextEditor) return;
      var toolbar = activeTextEditor.element.querySelector('[data-testid="text-run-format-toolbar"]');
      if (!toolbar) return;
      var info = activeTextRunInfo();
      var properties = info && info.properties ? info.properties : {};
      var disabled = !info || !info.handle;
      setToolbarControl(toolbar, "bold", disabled, Boolean(properties.bold));
      setToolbarControl(toolbar, "italic", disabled, Boolean(properties.italic));
      setToolbarControl(toolbar, "underline", disabled, Boolean(properties.underline));
      var size = toolbar.querySelector('[data-testid="text-run-format-font-size"]');
      if (size) {
        size.disabled = disabled;
        size.value = properties.fontSize == null ? "" : String(properties.fontSize);
      }
      var color = toolbar.querySelector('[data-testid="text-run-format-color"]');
      if (color) {
        color.disabled = disabled;
        color.value =
          properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string"
            ? "#" + properties.color.hex
            : "#000000";
      }
      var typeface = toolbar.querySelector('[data-testid="text-run-format-typeface"]');
      if (typeface) {
        typeface.disabled = disabled;
        typeface.value = typeof properties.typeface === "string" ? properties.typeface : "";
      }
      Array.prototype.forEach.call(toolbar.querySelectorAll('[data-testid^="text-run-format-clear-"]'), function (control) {
        control.disabled = disabled;
      });
    }

    function setToolbarControl(toolbar, property, disabled, pressed) {
      var control = toolbar.querySelector('[data-testid="text-run-format-' + property + '"]');
      if (!control) return;
      control.disabled = disabled;
      control.setAttribute("aria-pressed", pressed ? "true" : "false");
    }

    function toggleActiveTextRunBooleanProperty(property) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      if ((info.properties || {})[property] === true) {
        applyActiveTextRunPropertyClear([property]);
        return;
      }
      var properties = {};
      properties[property] = true;
      applyActiveTextRunPropertySet(properties);
    }

    function applyActiveTextRunPropertySet(properties) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      applyActiveTextRunPropertyCommand({
        kind: "setTextRunProperties",
        handle: info.handle,
        properties: properties
      });
    }

    function applyActiveTextRunPropertyClear(properties) {
      var info = activeTextRunInfo();
      if (!info || !info.handle) {
        setEditorMessage("No editable text run selected", true);
        return;
      }
      applyActiveTextRunPropertyCommand({
        kind: "clearTextRunProperties",
        handle: info.handle,
        properties: properties
      });
    }

    function applyActiveTextRunPropertyCommand(command) {
      postJson("/api/editor/command", { command: command })
        .then(function (data) {
          applyEditorResponseBehindTextEditor(data);
          patchActiveTextRunProperties(command);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          setEditorMessage(err.message, true);
        });
    }

    function applyEditorResponseBehindTextEditor(data) {
      slides = data.slides || slides;
      slideCount = slides.length;
      updateHistory(data.history);
      if (data.selection && data.selection.shapeHandle) {
        selectedShapeKey = handleKey(data.selection.shapeHandle);
      }
      renderThumbnails();
      replaceCurrentRenderedSvg();
      renderSelectionOverlay();
      loadShapeOptions(currentIndex + 1);
    }

    function replaceCurrentRenderedSvg() {
      var container = document.getElementById("slide-container");
      var previous = container.querySelector("svg:not(#selection-overlay)");
      var slide = slides[currentIndex];
      if (!slide) return;
      var wrapper = document.createElement("div");
      wrapper.innerHTML = slide.svg;
      var next = wrapper.querySelector("svg");
      if (!next) return;
      next.removeAttribute("width");
      next.removeAttribute("height");
      next.style.width = "100%";
      next.style.height = "auto";
      if (previous) {
        previous.replaceWith(next);
      } else {
        container.insertBefore(next, container.firstChild);
      }
    }

    function patchActiveTextRunProperties(command) {
      var textNode = activeTextRunTextNode();
      if (!textNode) return;
      var mark = (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      if (!mark) return;
      var attrs = mark.attrs || {};
      var properties = { ...(attrs.properties || {}) };
      if (command.kind === "setTextRunProperties") {
        Object.keys(command.properties || {}).forEach(function (property) {
          properties[property] = command.properties[property];
        });
      }
      if (command.kind === "clearTextRunProperties") {
        (command.properties || []).forEach(function (property) {
          delete properties[property];
        });
      }
      attrs.properties = Object.keys(properties).length > 0 ? properties : null;
      mark.attrs = attrs;
      refreshTextRunFormatToolbar();
      syncActiveTextRunStyle(properties);
    }

    function activeTextRunTextNode() {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return null;
      var element = activeTextEditor.selectedRunElement;
      var paragraphIndex = Number(element.dataset.paragraphIndex || "-1");
      var runIndex = Number(element.dataset.runIndex || "-1");
      var paragraph = (activeTextEditor.originalDocJson.content || [])[paragraphIndex];
      return paragraph && (paragraph.content || [])[runIndex] ? paragraph.content[runIndex] : null;
    }

    function syncActiveTextRunStyle(properties) {
      if (!activeTextEditor || !activeTextEditor.selectedRunElement) return;
      var run = activeTextEditor.selectedRunElement;
      run.style.fontWeight = properties.bold === true ? "700" : "";
      run.style.fontStyle = properties.italic === true ? "italic" : "";
      run.style.textDecoration = properties.underline === true ? "underline" : "";
      run.style.fontSize = properties.fontSize != null ? String(properties.fontSize) + "pt" : "";
      run.style.fontFamily = properties.typeface
        ? '"' + String(properties.typeface).replace(/"/g, "") + '"'
        : "";
      run.style.color =
        properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string"
          ? "#" + properties.color.hex
          : "";
    }

    function createTextEditorContent(docJson) {
      var body = document.createElement("div");
      body.setAttribute("data-testid", "text-editor-content");
      (docJson.content || []).forEach(function (paragraph, paragraphIndex) {
        var paragraphElement = document.createElement("div");
        paragraphElement.className = "text-editor-paragraph";
        paragraphElement.dataset.paragraphIndex = String(paragraphIndex);
        (paragraph.content || []).forEach(function (textNode, runIndex) {
          var run = document.createElement("span");
          run.className = "text-editor-run";
          run.setAttribute("data-testid", "text-editor-run");
          run.contentEditable = "true";
          run.dataset.paragraphIndex = String(paragraphIndex);
          run.dataset.runIndex = String(runIndex);
          run.textContent = textNode.text || "";
          applyTextRunStyle(run, textNode);
          run.addEventListener("focus", function () {
            setActiveTextRunElement(run);
          });
          run.addEventListener("pointerdown", function () {
            setActiveTextRunElement(run);
          });
          run.addEventListener("beforeinput", function (event) {
            if (event.inputType === "insertParagraph" || event.inputType === "insertLineBreak") {
              event.preventDefault();
              commitTextEditor().catch(function () {});
            }
          });
          run.addEventListener("paste", function (event) {
            event.preventDefault();
            var text = (event.clipboardData ? event.clipboardData.getData("text/plain") : "")
              .replace(/\\r?\\n/g, " ");
            document.execCommand("insertText", false, text);
          });
          paragraphElement.appendChild(run);
        });
        body.appendChild(paragraphElement);
      });
      return body;
    }

    function applyTextRunStyle(run, textNode) {
      var mark = (textNode.marks || []).find(function (candidate) {
        return candidate.type === "pptxRun";
      });
      var properties = mark && mark.attrs && mark.attrs.properties ? mark.attrs.properties : {};
      if (properties.bold === true) run.style.fontWeight = "700";
      if (properties.italic === true) run.style.fontStyle = "italic";
      if (properties.underline === true) run.style.textDecoration = "underline";
      if (properties.fontSize != null) run.style.fontSize = String(properties.fontSize) + "pt";
      if (properties.typeface) run.style.fontFamily = '"' + String(properties.typeface).replace(/"/g, "") + '"';
      if (properties.color && properties.color.kind === "srgb" && typeof properties.color.hex === "string") {
        run.style.color = "#" + properties.color.hex;
      }
    }

    function selectElementContents(element) {
      var range = document.createRange();
      range.selectNodeContents(element);
      var selection = window.getSelection();
      if (!selection) return;
      selection.removeAllRanges();
      selection.addRange(range);
    }

    function positionActiveTextEditor() {
      if (!activeTextEditor || !activeTextEditor.shape || !activeTextEditor.shape.bounds) return;
      var container = document.getElementById("slide-container");
      var renderedSvg = container.querySelector("svg:not(#selection-overlay)");
      if (!renderedSvg) return;
      var svgRect = renderedSvg.getBoundingClientRect();
      var containerRect = container.getBoundingClientRect();
      var viewBox = parseViewBox(renderedSvg.getAttribute("viewBox") || "0 0 960 540");
      var bounds = activeTextEditor.shape.bounds;
      var left = svgRect.left - containerRect.left + ((bounds.x - viewBox.x) / viewBox.width) * svgRect.width;
      var top = svgRect.top - containerRect.top + ((bounds.y - viewBox.y) / viewBox.height) * svgRect.height;
      var width = (bounds.width / viewBox.width) * svgRect.width;
      var height = (bounds.height / viewBox.height) * svgRect.height;
      activeTextEditor.element.style.left = left + "px";
      activeTextEditor.element.style.top = top + "px";
      activeTextEditor.element.style.width = width + "px";
      activeTextEditor.element.style.height = height + "px";
    }

    function parseViewBox(value) {
      var parts = String(value).trim().split(/\\s+/).map(Number);
      return {
        x: Number.isFinite(parts[0]) ? parts[0] : 0,
        y: Number.isFinite(parts[1]) ? parts[1] : 0,
        width: Number.isFinite(parts[2]) && parts[2] > 0 ? parts[2] : 960,
        height: Number.isFinite(parts[3]) && parts[3] > 0 ? parts[3] : 540
      };
    }

    function textEditorDocJson() {
      if (!activeTextEditor) return null;
      var original = activeTextEditor.originalDocJson;
      var paragraphs = [];
      (original.content || []).forEach(function (paragraph, paragraphIndex) {
        var paragraphElement = activeTextEditor.element.querySelector(
          '.text-editor-paragraph[data-paragraph-index="' + paragraphIndex + '"]'
        );
        var content = [];
        var emptyRunCount = 0;
        (paragraph.content || []).forEach(function (textNode, runIndex) {
          var runElement = paragraphElement
            ? paragraphElement.querySelector('.text-editor-run[data-run-index="' + runIndex + '"]')
            : null;
          var text = runElement ? runElement.textContent || "" : textNode.text || "";
          if (text.length === 0) {
            emptyRunCount += 1;
            return;
          }
          content.push({
            type: "text",
            text: text,
            marks: textNode.marks || []
          });
        });
        if (emptyRunCount > 0 && content.length > 0) {
          throw new Error("Clearing individual runs in multi-run text is unsupported.");
        }
        paragraphs.push({
          type: "paragraph",
          attrs: paragraph.attrs || {},
          content: content
        });
      });
      return { type: "doc", content: paragraphs };
    }

    function commitTextEditor() {
      if (!activeTextEditor) return Promise.resolve();
      var editor = activeTextEditor;
      if (editor.committing) return editor.commitPromise || Promise.resolve();
      var docJson = textEditorDocJson();
      if (!docJson) return Promise.resolve();
      editor.committing = true;
      editor.commitPromise = postJson("/api/editor/text-body", {
        handle: editor.shape.handle,
        docJson: docJson
      })
        .then(function (data) {
          closeTextEditor(editor);
          applyEditorResponse(data);
          setEditorMessage("Applied", false);
        })
        .catch(function (err) {
          editor.committing = false;
          editor.commitPromise = null;
          setEditorMessage(err.message, true);
          throw err;
        });
      return editor.commitPromise;
    }

    function closeTextEditor(editor) {
      var target = editor || activeTextEditor;
      if (!target) return;
      if (editor && activeTextEditor !== editor) return;
      if (!editor && target.committing) return;
      target.element.remove();
      if (activeTextEditor === target) activeTextEditor = null;
      updateSelectedShapeActions();
    }

`;
