async function main() {
  const { setupUI } = await import("./ui.js");
  setupUI();
}

main().catch(console.error);
