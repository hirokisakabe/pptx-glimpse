import { writeFile } from "node:fs/promises";

import { createAuthoringIntegrationFixture } from "../e2e/fixtures/authoring-integration.js";

const outputUrl = new URL("../shared-fixtures/authoring-integration.pptx", import.meta.url);

await writeFile(outputUrl, createAuthoringIntegrationFixture());
console.log(`Wrote ${outputUrl.pathname}`);
