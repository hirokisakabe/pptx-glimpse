"use client";

import type {
  PptxEditorSlideLayoutCatalogEntry,
  PptxEditorSlideMasterCatalogEntry,
  SourceHandle,
} from "pptx-glimpse";
import {
  memo,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  useSyncExternalStore,
} from "react";

import { LayoutPickerAddLifecycle } from "./layout-picker-add-lifecycle.js";
import {
  type LayoutPreviewLoader,
  type LayoutPreviewState,
  LayoutPreviewStore,
} from "./layout-preview-store.js";

interface EditorLayoutPickerProps {
  readonly busy: boolean;
  readonly catalog: readonly PptxEditorSlideMasterCatalogEntry[];
  readonly currentLayoutPartPath?: string;
  readonly interactionScope: object;
  readonly onAdd: (layout: PptxEditorSlideLayoutCatalogEntry) => Promise<boolean>;
  readonly previewLayout: LayoutPreviewLoader;
}

/** Ordered layout catalog with session-local, non-mutating SVG preview state. */
export function EditorLayoutPicker({
  busy,
  catalog,
  currentLayoutPartPath,
  interactionScope,
  onAdd,
  previewLayout,
}: EditorLayoutPickerProps) {
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const optionRefs = useRef(new Map<string, HTMLButtonElement>());
  const [open, setOpen] = useState(false);
  const [adding, setAdding] = useState(false);
  const [selectedKey, setSelectedKey] = useState<string | null>(null);
  const [addLifecycle] = useState(() => new LayoutPickerAddLifecycle(interactionScope));
  const previewStore = useMemo(
    () => new LayoutPreviewStore(previewLayout),
    [interactionScope, previewLayout],
  );
  useSyncExternalStore(previewStore.subscribe, previewStore.getVersion, previewStore.getVersion);
  const layouts = useMemo(() => catalog.flatMap((master) => master.layouts), [catalog]);
  const currentLayout = layouts.find((layout) => layout.handle.partPath === currentLayoutPartPath);
  const selectedLayout =
    layouts.find((layout) => handleKey(layout.handle) === selectedKey) ??
    currentLayout ??
    layouts[0];

  useLayoutEffect(() => {
    addLifecycle.activate(interactionScope);
    setOpen(false);
    setAdding(false);
    setSelectedKey(null);
    optionRefs.current.clear();
    return () => addLifecycle.invalidate();
  }, [addLifecycle, interactionScope]);

  useEffect(() => () => previewStore.dispose(), [previewStore]);

  useEffect(() => {
    if (open) previewStore.load(layouts.map((layout) => layout.handle));
  }, [layouts, open, previewStore]);

  const close = (restoreFocus: boolean) => {
    setOpen(false);
    if (restoreFocus) window.setTimeout(() => triggerRef.current?.focus(), 0);
  };

  const openPicker = () => {
    const initial = currentLayout ?? layouts[0];
    if (initial === undefined) return;
    const key = handleKey(initial.handle);
    setSelectedKey(key);
    setOpen(true);
    window.setTimeout(() => optionRefs.current.get(key)?.focus(), 0);
  };

  const selectRelative = (key: string, delta: number | "first" | "last") => {
    const currentIndex = layouts.findIndex((layout) => handleKey(layout.handle) === key);
    if (currentIndex < 0) return;
    const nextIndex =
      delta === "first"
        ? 0
        : delta === "last"
          ? layouts.length - 1
          : (currentIndex + delta + layouts.length) % layouts.length;
    const next = layouts[nextIndex];
    if (next === undefined) return;
    const nextKey = handleKey(next.handle);
    setSelectedKey(nextKey);
    optionRefs.current.get(nextKey)?.focus();
  };

  return (
    <div className="layout-picker">
      <button
        ref={triggerRef}
        aria-expanded={open}
        aria-haspopup="dialog"
        data-testid="new-slide-menu-button"
        disabled={busy || adding || layouts.length === 0}
        type="button"
        onClick={() => (open ? close(false) : openPicker())}
      >
        New slide
      </button>
      {open ? (
        <div
          aria-label="Choose a slide layout"
          className="layout-picker-popover"
          data-testid="layout-picker"
          role="dialog"
          onKeyDown={(event) => {
            if (event.key !== "Escape") return;
            event.preventDefault();
            close(true);
          }}
        >
          <div className="layout-picker-heading">
            <div>
              <strong>New slide</strong>
              <span>Choose a layout</span>
            </div>
            <button
              aria-label="Close layout picker"
              disabled={busy || adding}
              type="button"
              onClick={() => close(true)}
            >
              ×
            </button>
          </div>
          <div className="layout-catalog" role="radiogroup" aria-label="Slide layouts">
            {catalog.map((master, masterIndex) => {
              const masterName = master.name?.trim() || `Master ${(masterIndex + 1).toString()}`;
              return (
                <section className="layout-master-group" key={handleKey(master.handle)}>
                  <h3>{masterName}</h3>
                  <div className="layout-master-options">
                    {master.layouts.map((layout, layoutIndex) => {
                      const key = handleKey(layout.handle);
                      const name = layout.name?.trim() || `Layout ${(layoutIndex + 1).toString()}`;
                      const selected =
                        selectedLayout !== undefined && key === handleKey(selectedLayout.handle);
                      const current = layout.handle.partPath === currentLayoutPartPath;
                      const preview = previewStore.get(layout.handle);
                      return (
                        <button
                          ref={(element) => {
                            if (element === null) optionRefs.current.delete(key);
                            else optionRefs.current.set(key, element);
                          }}
                          aria-checked={selected}
                          aria-label={`${masterName}, ${name}${layout.hidden ? ", hidden" : ""}${current ? ", current layout" : ""}`}
                          className={`layout-option${selected ? " selected" : ""}`}
                          data-layout-part-path={layout.handle.partPath}
                          data-testid="layout-option"
                          key={key}
                          role="radio"
                          tabIndex={selected ? 0 : -1}
                          type="button"
                          disabled={busy || adding}
                          onClick={() => setSelectedKey(key)}
                          onKeyDown={(event) => {
                            const direction =
                              event.key === "ArrowRight" || event.key === "ArrowDown"
                                ? 1
                                : event.key === "ArrowLeft" || event.key === "ArrowUp"
                                  ? -1
                                  : event.key === "Home"
                                    ? "first"
                                    : event.key === "End"
                                      ? "last"
                                      : undefined;
                            if (direction === undefined) return;
                            event.preventDefault();
                            selectRelative(key, direction);
                          }}
                        >
                          <MemoizedLayoutThumbnail name={name} preview={preview} />
                          <span className="layout-option-name">{name}</span>
                          <span className="layout-option-metadata">
                            {layout.type ?? "Unspecified type"}
                            {layout.hidden ? <em>Hidden</em> : null}
                            {current ? <em>Current</em> : null}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </section>
              );
            })}
          </div>
          <div className="layout-picker-actions">
            <span aria-live="polite">
              {selectedLayout === undefined
                ? "No layout available"
                : `${selectedLayout.name?.trim() || "Unnamed layout"} selected`}
            </span>
            <button
              data-testid="add-slide-from-layout"
              disabled={busy || adding || selectedLayout === undefined}
              type="button"
              onClick={() => {
                if (selectedLayout === undefined) return;
                const completion = addLifecycle.capture();
                setAdding(true);
                void onAdd(selectedLayout).then(
                  (added) => {
                    if (!addLifecycle.isCurrent(completion)) return;
                    setAdding(false);
                    if (added) {
                      setOpen(false);
                      window.setTimeout(() => {
                        if (addLifecycle.isCurrent(completion)) triggerRef.current?.focus();
                      }, 0);
                    }
                  },
                  () => {
                    if (addLifecycle.isCurrent(completion)) setAdding(false);
                  },
                );
              }}
            >
              {adding ? "Adding…" : "Add slide"}
            </button>
          </div>
        </div>
      ) : null}
    </div>
  );
}

export function LayoutThumbnail({
  name,
  preview,
}: {
  readonly name: string;
  readonly preview?: LayoutPreviewState;
}) {
  if (preview?.status === "ready") {
    return (
      <img
        alt={`${name} preview`}
        className="layout-thumbnail-preview"
        data-thumbnail-state="ready"
        src={`data:image/svg+xml;charset=utf-8,${encodeURIComponent(preview.svg)}`}
      />
    );
  }
  const loading = preview === undefined || preview.status === "loading";
  const message = loading ? "Loading preview…" : preview.message;
  return (
    <span
      aria-label={`${name}, ${message}`}
      className={`layout-thumbnail-fallback${loading ? " is-loading" : ""}`}
      data-thumbnail-state={loading ? "loading" : "fallback"}
      role="img"
    >
      <span aria-hidden="true" />
      <small>{message}</small>
    </span>
  );
}

const MemoizedLayoutThumbnail = memo(LayoutThumbnail);

function handleKey(handle: SourceHandle): string {
  return [
    handle.partPath ?? "",
    handle.nodeId ?? "",
    handle.relationshipId ?? "",
    handle.orderingSlot === undefined ? "" : String(handle.orderingSlot),
    JSON.stringify(handle.rawSidecarIds ?? []),
  ].join("\u0000");
}
