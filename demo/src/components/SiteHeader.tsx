export function SiteHeader() {
  return (
    <header className="site-header">
      <a className="site-mark" href="/" aria-label="pptx-glimpse demo home">
        <span className="site-mark-icon" aria-hidden="true">
          P
        </span>
        <span>pptx-glimpse</span>
      </a>
      <nav className="site-nav" aria-label="Primary navigation">
        <a href="/">Demo</a>
        <a href="/docs">Documentation</a>
        <a href="https://github.com/hirokisakabe/pptx-glimpse" target="_blank" rel="noreferrer">
          GitHub
        </a>
      </nav>
    </header>
  );
}
