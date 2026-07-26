export const DOCS_NAVIGATION = [
  {
    label: "Start here",
    items: [
      { href: "/docs", label: "Overview" },
      { href: "/docs/getting-started", label: "Getting started" },
      { href: "/docs/why", label: "Why pptx-glimpse" },
    ],
  },
  {
    label: "Guides",
    items: [
      { href: "/docs/rendering", label: "Rendering presentations" },
      { href: "/docs/editing", label: "Editing presentations" },
      { href: "/docs/fonts", label: "Using fonts" },
      { href: "/docs/browser", label: "Browser usage" },
      { href: "/docs/nodejs", label: "Node.js usage" },
    ],
  },
  {
    label: "Reference",
    items: [
      { href: "/docs/api", label: "High-level API" },
      { href: "/docs/api-reference", label: "Generated API reference" },
      { href: "/docs/feature-support", label: "Feature support" },
      { href: "/docs/packages", label: "Choosing a package" },
    ],
  },
] as const;
