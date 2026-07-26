import Link from "next/link";

const paths = [
  {
    marker: "SVG",
    title: "Render presentations",
    description: "Convert a complete deck or selected slides to SVG or PNG.",
    href: "/docs/rendering",
    action: "Open guide",
  },
  {
    marker: "EDIT",
    title: "Build an editing flow",
    description: "Read, inspect, edit, rerender, undo, redo, and save through one session.",
    href: "/docs/editing",
    action: "Open guide",
  },
  {
    marker: "API",
    title: "Look up an API",
    description: "Find complete signatures, options, return values, runtime notes, and errors.",
    href: "/docs/api",
    action: "Open reference",
  },
] as const;

export function DocsHero({
  title,
  description,
}: {
  readonly title: string;
  readonly description: string;
}) {
  return (
    <header className="docs-mdx-hero">
      <h1>{title}</h1>
      <p>{description}</p>
    </header>
  );
}

export function DocsPathGrid() {
  return (
    <div className="docs-path-grid docs-path-grid-new">
      {paths.map((path) => (
        <Link className="docs-path" href={path.href} key={path.href}>
          <span>{path.marker}</span>
          <h3>{path.title}</h3>
          <p>{path.description}</p>
          <strong>{path.action} →</strong>
        </Link>
      ))}
    </div>
  );
}

export function DocsWorkflow() {
  return (
    <ol className="docs-process">
      <li>
        <span>01</span>
        <div>
          <strong>Open</strong>
          <p>Read PPTX bytes from a file, request, object store, or another source.</p>
        </div>
      </li>
      <li>
        <span>02</span>
        <div>
          <strong>Render or edit</strong>
          <p>Convert slides directly, or keep an editor session for repeated changes.</p>
        </div>
      </li>
      <li>
        <span>03</span>
        <div>
          <strong>Use the output</strong>
          <p>Embed SVG, write PNG, or save updated PPTX bytes.</p>
        </div>
      </li>
    </ol>
  );
}

export function DocsPackageStack() {
  return (
    <ol className="docs-package-stack" aria-label="Packages from highest to lowest level">
      <li>
        <code>pptx-glimpse</code>
        <span>render + high-level editing</span>
      </li>
      <li>
        <code>@pptx-glimpse/editor</code>
        <span>commands + selection + history</span>
      </li>
      <li>
        <code>@pptx-glimpse/document</code>
        <span>OOXML source + computed view + writing</span>
      </li>
    </ol>
  );
}

export function DocsNextLinks() {
  return (
    <div className="docs-link-list">
      <Link href="/docs/rendering">
        <strong>Control rendering</strong>
        <span>Selected slides, output size, fonts, SVG text, and browser PNG.</span>
      </Link>
      <Link href="/docs/editing">
        <strong>Build an editing flow</strong>
        <span>Inspect shapes, apply commands, manage history, and save PPTX bytes.</span>
      </Link>
      <Link href="/docs/packages">
        <strong>Work below the high-level API</strong>
        <span>Read, author, or edit documents with the lower-level packages.</span>
      </Link>
    </div>
  );
}
