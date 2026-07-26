# @pptx-glimpse/mcp

Node.js-only [Model Context Protocol](https://modelcontextprotocol.io/) server that previews
PPTX slides as PNG images. It is a thin integration layer over the public `pptx-glimpse`
conversion API and communicates over stdio.

## Requirements

- Node.js 22 or later
- A local PPTX file readable by the MCP server process

## Configuration

Use the package directly with `npx`; no global installation is required.

```json
{
  "mcpServers": {
    "pptx-glimpse": {
      "command": "npx",
      "args": ["-y", "@pptx-glimpse/mcp"]
    }
  }
}
```

The server exposes one tool:

- `preview_pptx`
  - `filePath`: path to a local PPTX file
  - `slides` (optional): array of 1-based slide numbers

Each converted slide is returned as an `image/png` content item. `structuredContent.slides`
maps every image's content index to its original slide number and dimensions.
`structuredContent.diagnostics` and `structuredContent.supportCoverage` summarize rendering
support.

Example arguments:

```json
{
  "filePath": "/absolute/path/to/deck.pptx",
  "slides": [1, 3]
}
```

Only stdio transport is supported. The package reads local files but does not modify them.
For predictable local-server resource usage, previews do not scan operating-system font
directories.
