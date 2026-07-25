import type { ReactNode } from "react";
import Link from "next/link";

export interface DocsTocItem {
  readonly href: `#${string}`;
  readonly label: string;
}

interface DocsPageProps {
  readonly eyebrow: string;
  readonly title: string;
  readonly description: string;
  readonly toc: readonly DocsTocItem[];
  readonly children: ReactNode;
}

export function DocsPage({ eyebrow, title, description, toc, children }: DocsPageProps) {
  return (
    <>
      <article className="docs-article">
        <header className="docs-article-header">
          <p className="eyebrow">{eyebrow}</p>
          <h1>{title}</h1>
          <p>{description}</p>
        </header>
        <div className="docs-prose">{children}</div>
      </article>

      <aside className="docs-toc">
        <p>On this page</p>
        <nav aria-label="On this page">
          <ul>
            {toc.map((item) => (
              <li key={item.href}>
                <a href={item.href}>{item.label}</a>
              </li>
            ))}
          </ul>
        </nav>
      </aside>
    </>
  );
}

export function DocsCallout({
  title,
  children,
}: {
  readonly title: string;
  readonly children: ReactNode;
}) {
  return (
    <aside className="docs-callout">
      <strong>{title}</strong>
      <div>{children}</div>
    </aside>
  );
}

export function DocsPager({
  previous,
  next,
}: {
  readonly previous?: { readonly href: string; readonly label: string };
  readonly next?: { readonly href: string; readonly label: string };
}) {
  return (
    <nav className="docs-pager" aria-label="Documentation pages">
      {previous ? (
        <Link href={previous.href}>
          <span>Previous</span>
          {previous.label}
        </Link>
      ) : (
        <span />
      )}
      {next ? (
        <Link href={next.href}>
          <span>Next</span>
          {next.label}
        </Link>
      ) : (
        <span />
      )}
    </nav>
  );
}
