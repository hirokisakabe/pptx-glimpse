import { GitHubIcon } from "nextra/icons";
import { BrandMark } from "@/components/BrandMark";
import { GITHUB_REPOSITORY_URL } from "@/lib/constants";

export function SiteHeader() {
  return (
    <header className="site-header">
      <a className="site-mark" href="/" aria-label="pptx-glimpse demo home">
        <BrandMark />
        <span>pptx-glimpse</span>
      </a>
      <nav className="site-nav" aria-label="Primary navigation">
        <a href="/docs">Documentation</a>
        <a
          className="site-nav-repository"
          href={GITHUB_REPOSITORY_URL}
          target="_blank"
          rel="noreferrer"
          aria-label="GitHub repository"
        >
          <GitHubIcon height="24" />
        </a>
      </nav>
    </header>
  );
}
