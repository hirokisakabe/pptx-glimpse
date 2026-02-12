import { readFileSync } from "fs";
import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

function prettyPrintXml(rawXml: string): string {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    preserveOrder: true,
    trimValues: false,
  });
  const parsed = parser.parse(rawXml);

  const builder = new XMLBuilder({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    preserveOrder: true,
    format: true,
    indentBy: "  ",
    suppressEmptyNode: false,
  });
  return builder.build(parsed) as string;
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function printXmlFile(zip: JSZip, path: string): Promise<void> {
  const file = zip.file(path);
  if (!file) {
    console.error(`Error: ${path} not found in archive`);
    process.exit(1);
  }
  const xml = await file.async("string");
  console.log(prettyPrintXml(xml));
}

async function handleRels(zip: JSZip): Promise<void> {
  const relFiles: string[] = [];
  zip.forEach((path, entry) => {
    if (!entry.dir && path.endsWith(".rels")) {
      relFiles.push(path);
    }
  });
  relFiles.sort();

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    isArray: (_name: string, jpath: string) => {
      return jpath.split(".").pop() === "Relationship";
    },
  });

  for (const relPath of relFiles) {
    const file = zip.file(relPath)!;
    const xml = await file.async("string");
    const parsed = parser.parse(xml);
    const relationships = parsed?.Relationships?.Relationship;
    if (!relationships || relationships.length === 0) continue;

    console.log(`\n=== ${relPath} ===`);
    for (const rel of relationships) {
      const id = rel["@_Id"] ?? "";
      const type = (rel["@_Type"] ?? "").split("/").pop() ?? "";
      const target = rel["@_Target"] ?? "";
      console.log(`  ${id}  ${type}  → ${target}`);
    }
  }
}

async function handleTree(zip: JSZip): Promise<void> {
  const entries: { path: string; size: number }[] = [];

  const promises: Promise<void>[] = [];
  zip.forEach((path, entry) => {
    if (entry.dir) return;
    promises.push(
      entry.async("nodebuffer").then((buf) => {
        entries.push({ path, size: buf.length });
      }),
    );
  });
  await Promise.all(promises);

  entries.sort((a, b) => a.path.localeCompare(b.path));

  const maxPathLen = Math.max(...entries.map((e) => e.path.length));

  for (const entry of entries) {
    const padding = " ".repeat(maxPathLen - entry.path.length + 2);
    console.log(`${entry.path}${padding}${formatBytes(entry.size)}`);
  }

  const totalSize = entries.reduce((sum, e) => sum + e.size, 0);
  console.log(`\n${entries.length} files, ${formatBytes(totalSize)} total`);
}

async function handleGrep(zip: JSZip, searchTerm: string): Promise<void> {
  const xmlFiles: string[] = [];
  zip.forEach((path, entry) => {
    if (!entry.dir && (path.endsWith(".xml") || path.endsWith(".rels"))) {
      xmlFiles.push(path);
    }
  });
  xmlFiles.sort();

  let matchCount = 0;

  for (const filePath of xmlFiles) {
    const file = zip.file(filePath)!;
    const rawContent = await file.async("string");

    if (!rawContent.includes(searchTerm)) continue;

    const pretty = prettyPrintXml(rawContent);
    const lines = pretty.split("\n");

    const matches: { lineNum: number; line: string }[] = [];
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes(searchTerm)) {
        matches.push({ lineNum: i + 1, line: lines[i].trimEnd() });
      }
    }

    if (matches.length > 0) {
      console.log(
        `\n=== ${filePath} (${matches.length} match${matches.length > 1 ? "es" : ""}) ===`,
      );
      for (const m of matches) {
        const maxLen = 200;
        const display = m.line.length > maxLen ? m.line.substring(0, maxLen) + "..." : m.line;
        console.log(`  L${m.lineNum}: ${display}`);
      }
      matchCount += matches.length;
    }
  }

  if (matchCount === 0) {
    console.log(`No matches found for "${searchTerm}"`);
  } else {
    console.log(`\n${matchCount} match${matchCount > 1 ? "es" : ""} total`);
  }
}

function printUsage(): void {
  console.error("Usage: npm run inspect -- <file.pptx> <command>");
  console.error("");
  console.error("Commands:");
  console.error("  slide<N>       Pretty-print slide N XML (e.g., slide1)");
  console.error("  theme          Pretty-print theme XML");
  console.error("  master         Pretty-print slide master XML");
  console.error("  layout<N>      Pretty-print slide layout N XML (e.g., layout1)");
  console.error("  presentation   Pretty-print presentation XML");
  console.error("  rels           Show all relationships");
  console.error("  tree           File tree with sizes");
  console.error("  grep <term>    Search for term in all XML files");
}

async function main(): Promise<void> {
  const filePath = process.argv[2];
  const command = process.argv[3];

  if (!filePath || !command) {
    printUsage();
    process.exit(1);
  }

  const input = readFileSync(filePath);
  const zip = await JSZip.loadAsync(input);

  const slideMatch = command.match(/^slide(\d+)$/);
  const layoutMatch = command.match(/^layout(\d+)$/);

  if (slideMatch) {
    await printXmlFile(zip, `ppt/slides/slide${slideMatch[1]}.xml`);
  } else if (command === "theme") {
    await printXmlFile(zip, "ppt/theme/theme1.xml");
  } else if (command === "master") {
    await printXmlFile(zip, "ppt/slideMasters/slideMaster1.xml");
  } else if (layoutMatch) {
    await printXmlFile(zip, `ppt/slideLayouts/slideLayout${layoutMatch[1]}.xml`);
  } else if (command === "presentation") {
    await printXmlFile(zip, "ppt/presentation.xml");
  } else if (command === "rels") {
    await handleRels(zip);
  } else if (command === "tree") {
    await handleTree(zip);
  } else if (command === "grep") {
    const searchTerm = process.argv[4];
    if (!searchTerm) {
      console.error("Error: grep command requires a search term");
      console.error("Usage: npm run inspect -- <file.pptx> grep <term>");
      process.exit(1);
    }
    await handleGrep(zip, searchTerm);
  } else {
    console.error(`Unknown command: ${command}`);
    printUsage();
    process.exit(1);
  }
}

main().catch(console.error);
