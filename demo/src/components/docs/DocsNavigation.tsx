"use client";

import Link from "next/link";
import { usePathname } from "next/navigation";
import { DOCS_NAVIGATION } from "@/lib/docs-navigation";

function NavigationLinks() {
  const pathname = usePathname();

  return (
    <nav className="docs-navigation" aria-label="Documentation">
      {DOCS_NAVIGATION.map((group) => (
        <section className="docs-navigation-group" key={group.label}>
          <p>{group.label}</p>
          <ul>
            {group.items.map((item) => {
              const isCurrent = pathname === item.href;

              return (
                <li key={item.href}>
                  <Link href={item.href} aria-current={isCurrent ? "page" : undefined}>
                    {item.label}
                  </Link>
                </li>
              );
            })}
          </ul>
        </section>
      ))}
    </nav>
  );
}

export function DocsNavigation() {
  return (
    <>
      <aside className="docs-sidebar">
        <Link className="docs-sidebar-title" href="/docs">
          <span aria-hidden="true">▱</span>
          Documentation
        </Link>
        <NavigationLinks />
      </aside>

      <details className="docs-mobile-navigation">
        <summary>Browse documentation</summary>
        <NavigationLinks />
      </details>
    </>
  );
}
