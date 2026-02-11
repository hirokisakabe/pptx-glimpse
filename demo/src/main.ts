import { Buffer } from "buffer";
(globalThis as unknown as { Buffer: typeof Buffer }).Buffer = Buffer;

async function main() {
  const { setupUI } = await import("./ui.js");
  setupUI();
}

main().catch(console.error);
