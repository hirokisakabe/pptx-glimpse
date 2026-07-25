import { Children, cloneElement, isValidElement, type ReactNode } from "react";
import { useMDXComponents } from "nextra-theme-docs";
import { codeToHtml } from "shiki";

export interface DocsTocItem {
  readonly href: `#${string}`;
  readonly label: string;
}

interface DocsPageProps {
  readonly title: string;
  readonly description: string;
  readonly filePath: `src/app/docs/${string}/page.tsx`;
  readonly toc: readonly DocsTocItem[];
  readonly children: ReactNode;
}

async function highlightCodeBlocks(node: ReactNode): Promise<ReactNode> {
  if (Array.isArray(node)) {
    return Promise.all(Children.toArray(node).map(highlightCodeBlocks));
  }

  if (!isValidElement<{ children?: ReactNode }>(node)) {
    return node;
  }

  if (node.type === "pre") {
    const code = node.props.children;

    if (
      isValidElement<{ children?: ReactNode }>(code) &&
      code.type === "code" &&
      typeof code.props.children === "string"
    ) {
      const html = await codeToHtml(code.props.children, {
        lang: "typescript",
        theme: "github-dark",
      });

      return (
        <div
          className="docs-code-block"
          dangerouslySetInnerHTML={{ __html: html }}
          key={node.key}
        />
      );
    }
  }

  if (node.props.children === undefined) {
    return node;
  }

  const children = await Promise.all(
    Children.toArray(node.props.children).map(highlightCodeBlocks),
  );

  return cloneElement(node, undefined, children);
}

export async function DocsPage({ title, description, filePath, toc, children }: DocsPageProps) {
  const Wrapper = useMDXComponents().wrapper;
  const highlightedChildren = await highlightCodeBlocks(children);

  if (!Wrapper) {
    throw new Error("Nextra documentation wrapper is unavailable");
  }

  return (
    <Wrapper
      toc={toc.map((item) => ({
        depth: 2,
        id: item.href.slice(1),
        value: item.label,
      }))}
      metadata={{
        title,
        description,
        filePath,
      }}
      sourceCode=""
    >
      <div className="docs-article">
        <header className="docs-article-header">
          <h1>{title}</h1>
          <p>{description}</p>
        </header>
        <div className="docs-prose">{highlightedChildren}</div>
      </div>
    </Wrapper>
  );
}

export function DocsPager({
  previous: _previous,
  next: _next,
}: {
  readonly previous?: { readonly href: string; readonly label: string };
  readonly next?: { readonly href: string; readonly label: string };
}) {
  return null;
}
