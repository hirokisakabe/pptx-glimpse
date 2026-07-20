import Link from "next/link";

export function SiteHeader() {
  return (
    <header className="site-header">
      <Link className="site-mark" href="/" aria-label="pptx-glimpse demo home">
        <span className="site-mark-icon" aria-hidden="true">
          P
        </span>
        <span>pptx-glimpse</span>
      </Link>
      <nav className="site-nav" aria-label="Primary navigation">
        <Link href="/">Demo</Link>
        <Link href="/docs">Documentation</Link>
        <a href="https://github.com/hirokisakabe/pptx-glimpse">GitHub</a>
      </nav>
    </header>
  );
}
