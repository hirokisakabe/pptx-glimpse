import type { ReactNode } from "react";
import { DocsNavigation } from "@/components/docs/DocsNavigation";

export default function DocsLayout({ children }: { readonly children: ReactNode }) {
  return (
    <main className="docs-shell">
      <DocsNavigation />
      {children}
    </main>
  );
}
