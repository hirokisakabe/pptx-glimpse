export const DOCS_NAVIGATION = [
  {
    label: "Start here",
    items: [
      { href: "/docs", label: "Overview" },
      { href: "/docs/getting-started", label: "Getting started" },
    ],
  },
  {
    label: "Guides",
    items: [
      { href: "/docs/rendering", label: "Render presentations" },
      { href: "/docs/editing", label: "Build an editing flow" },
      { href: "/docs/browser", label: "Run in the browser" },
      { href: "/docs/nodejs", label: "Run in Node.js" },
    ],
  },
  {
    label: "Reference",
    items: [
      { href: "/docs/api", label: "High-level API" },
      { href: "/docs/packages", label: "Packages" },
    ],
  },
] as const;
