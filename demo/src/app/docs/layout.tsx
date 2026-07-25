import type { ReactNode } from "react";
import { getPageMap } from "nextra/page-map";
import { Layout, Navbar } from "nextra-theme-docs";
import { BrandMark } from "@/components/BrandMark";
import { GITHUB_REPOSITORY_URL } from "@/lib/constants";

function DocsLogo() {
  return (
    <span className="docs-brand">
      <BrandMark />
      <span>pptx-glimpse</span>
    </span>
  );
}

export default async function DocsLayout({ children }: { readonly children: ReactNode }) {
  const navbar = (
    <Navbar logo={<DocsLogo />} logoLink="/docs" projectLink={GITHUB_REPOSITORY_URL} />
  );

  return (
    <div className="docs-site">
      <Layout
        navbar={navbar}
        pageMap={await getPageMap("/docs")}
        docsRepositoryBase={`${GITHUB_REPOSITORY_URL}/tree/main/demo`}
        sidebar={{ defaultMenuCollapseLevel: 1, toggleButton: true }}
        toc={{ title: "On this page", backToTop: "Back to top" }}
        feedback={{ content: "Documentation feedback", labels: "documentation" }}
        editLink="Edit this page on GitHub"
      >
        {children}
      </Layout>
    </div>
  );
}
