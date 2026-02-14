import { watch } from "fs";
import { createServer } from "http";
import { basename, resolve } from "path";
import { execFile } from "child_process";
import { promisify } from "util";
import { WebSocketServer, type WebSocket } from "ws";

const DEFAULT_PORT = 3000;
const DEBOUNCE_MS = 300;
const WATCH_DIR = resolve("src");
const RENDER_TIMEOUT_MS = 30_000;
const MAX_BUFFER = 50 * 1024 * 1024;

const execFileAsync = promisify(execFile);

interface SlideSvg {
  slideNumber: number;
  svg: string;
}

// --- Rendering via child process ---

async function renderSlides(pptxPath: string): Promise<SlideSvg[]> {
  const workerPath = resolve("scripts/dev-server-render.ts");
  const { stdout } = await execFileAsync("npx", ["tsx", workerPath, pptxPath], {
    maxBuffer: MAX_BUFFER,
    timeout: RENDER_TIMEOUT_MS,
  });
  return JSON.parse(stdout) as SlideSvg[];
}

// --- WebSocket ---

function broadcast(wss: WebSocketServer, data: unknown): void {
  const msg = JSON.stringify(data);
  for (const client of wss.clients) {
    if (client.readyState === 1) {
      client.send(msg);
    }
  }
}

// --- File watcher ---

function watchSourceFiles(onChange: () => void): void {
  let timeout: ReturnType<typeof setTimeout> | null = null;

  watch(WATCH_DIR, { recursive: true }, (_event, filename) => {
    if (!filename || !filename.endsWith(".ts")) return;
    if (filename.endsWith(".test.ts")) return;

    console.log(`Change detected: ${filename}`);

    if (timeout) clearTimeout(timeout);
    timeout = setTimeout(() => {
      timeout = null;
      onChange();
    }, DEBOUNCE_MS);
  });
}

// --- HTML template ---

function generateHtml(slides: SlideSvg[], pptxName: string): string {
  const thumbnailsHtml = slides
    .map(
      (s, i) =>
        `<div class="thumbnail${i === 0 ? " active" : ""}" data-index="${i}">` +
        `<div class="thumb-label">Slide ${String(s.slideNumber)}</div>` +
        `<div class="thumb-svg">${s.svg}</div>` +
        `</div>`,
    )
    .join("");

  const firstSvg = slides.length > 0 ? slides[0].svg : "<p>No slides</p>";

  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>pptx-glimpse dev - ${escapeHtml(pptxName)}</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background: #1a1a2e;
      color: #e0e0e0;
    }
    #header {
      padding: 12px 20px;
      background: #16213e;
      display: flex;
      align-items: center;
      justify-content: space-between;
    }
    #header h1 { font-size: 14px; font-weight: 600; color: #a0a0c0; }
    #status { font-size: 12px; color: #4caf50; }
    #status.rendering { color: #ff9800; }
    #status.error { color: #f44336; }
    #main { display: flex; height: calc(100vh - 48px); }
    #sidebar {
      width: 180px;
      overflow-y: auto;
      background: #16213e;
      padding: 8px;
      border-right: 1px solid #2a2a4a;
    }
    .thumbnail {
      margin-bottom: 8px;
      padding: 4px;
      border: 2px solid transparent;
      border-radius: 4px;
      cursor: pointer;
      background: #fff;
    }
    .thumbnail.active { border-color: #4472c4; }
    .thumbnail:hover { border-color: #6090d0; }
    .thumb-label {
      font-size: 10px;
      color: #888;
      text-align: center;
      padding: 2px 0;
      background: #16213e;
    }
    .thumb-svg svg { width: 100%; height: auto; display: block; }
    #viewer {
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      overflow: auto;
    }
    #slide-container {
      background: #fff;
      border-radius: 4px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      max-width: 100%;
      max-height: 100%;
    }
    #slide-container svg { display: block; width: 100%; height: auto; }
    #info {
      padding: 4px 20px;
      background: #16213e;
      font-size: 11px;
      color: #666;
      text-align: center;
    }
  </style>
</head>
<body>
  <div id="header">
    <h1>pptx-glimpse dev &mdash; ${escapeHtml(pptxName)}</h1>
    <span id="status">Connected</span>
  </div>
  <div id="main">
    <div id="sidebar">${thumbnailsHtml}</div>
    <div id="viewer">
      <div id="slide-container">${firstSvg}</div>
    </div>
  </div>
  <div id="info">Slide 1 / ${String(slides.length)}</div>
  <script>
    var currentIndex = 0;
    var slideCount = ${String(slides.length)};

    function selectSlide(index) {
      currentIndex = index;
      var thumbs = document.querySelectorAll(".thumbnail");
      for (var i = 0; i < thumbs.length; i++) {
        if (i === index) {
          thumbs[i].classList.add("active");
        } else {
          thumbs[i].classList.remove("active");
        }
      }
      var thumb = thumbs[index];
      var svgHtml = thumb.querySelector(".thumb-svg").innerHTML;
      document.getElementById("slide-container").innerHTML = svgHtml;
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
      document.getElementById("info").textContent =
        "Slide " + (index + 1) + " / " + slideCount;
    }

    // Click handlers for thumbnails
    var thumbs = document.querySelectorAll(".thumbnail");
    for (var i = 0; i < thumbs.length; i++) {
      (function (idx) {
        thumbs[idx].addEventListener("click", function () {
          selectSlide(idx);
        });
      })(i);
    }

    // WebSocket for live reload
    function connect() {
      var ws = new WebSocket("ws://" + location.host);
      var status = document.getElementById("status");

      ws.onopen = function () {
        status.textContent = "Connected";
        status.className = "";
      };
      ws.onclose = function () {
        status.textContent = "Disconnected - reconnecting...";
        status.className = "error";
        setTimeout(connect, 2000);
      };
      ws.onmessage = function (event) {
        var data = JSON.parse(event.data);
        if (data.type === "rendering") {
          status.textContent = "Re-rendering...";
          status.className = "rendering";
        } else if (data.type === "reload") {
          status.textContent = "Updating...";
          status.className = "rendering";
          location.reload();
        } else if (data.type === "error") {
          status.textContent = "Error: " + data.message;
          status.className = "error";
        }
      };
    }
    connect();

    // Initial: resize the main SVG
    (function () {
      var svg = document.querySelector("#slide-container svg");
      if (svg) {
        svg.removeAttribute("width");
        svg.removeAttribute("height");
        svg.style.width = "100%";
        svg.style.height = "auto";
      }
    })();

    // Keyboard navigation
    document.addEventListener("keydown", function (e) {
      if (e.key === "ArrowLeft" && currentIndex > 0) selectSlide(currentIndex - 1);
      if (e.key === "ArrowRight" && currentIndex < slideCount - 1)
        selectSlide(currentIndex + 1);
    });
  </script>
</body>
</html>`;
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// --- Main ---

async function main(): Promise<void> {
  const pptxPath = process.argv[2];
  if (!pptxPath) {
    console.error("Usage: npm run dev -- <pptx-file> [--port <port>]");
    process.exit(1);
  }

  const portArgIdx = process.argv.indexOf("--port");
  const port = portArgIdx !== -1 ? Number(process.argv[portArgIdx + 1]) : DEFAULT_PORT;

  const resolvedPath = resolve(pptxPath);
  const pptxName = basename(resolvedPath);

  console.log(`Loading: ${resolvedPath}`);

  let slides = await renderSlides(resolvedPath);
  console.log(`Rendered ${String(slides.length)} slide(s)`);

  const server = createServer((req, res) => {
    if (req.url === "/" || req.url === "/index.html") {
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(generateHtml(slides, pptxName));
    } else {
      res.writeHead(404);
      res.end("Not Found");
    }
  });

  const wss = new WebSocketServer({ server });
  wss.on("connection", (_ws: WebSocket) => {
    console.log("Browser connected");
  });

  server.listen(port, () => {
    console.log(`Dev server running at http://localhost:${String(port)}`);
    console.log(`Watching: ${WATCH_DIR}`);
  });

  let rendering = false;

  watchSourceFiles(() => {
    if (rendering) return;
    rendering = true;

    console.log("Re-rendering...");
    broadcast(wss, { type: "rendering" });

    renderSlides(resolvedPath)
      .then((result) => {
        slides = result;
        console.log(`Re-rendered ${String(slides.length)} slide(s)`);
        broadcast(wss, { type: "reload" });
      })
      .catch((err: unknown) => {
        const message = err instanceof Error ? err.message : String(err);
        console.error(`Render error: ${message}`);
        broadcast(wss, { type: "error", message });
      })
      .finally(() => {
        rendering = false;
      });
  });
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
