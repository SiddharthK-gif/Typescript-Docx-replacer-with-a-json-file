import * as fs from "fs";
import * as path from "path";
import { replaceTemplateVariables } from "./replacer";

// ─── Types ────────────────────────────────────────────────────────────────────

export type JsonValue = string | number | boolean | null | JsonObject | JsonArray;
export interface JsonObject { [key: string]: JsonValue }
export type JsonArray = JsonValue[];

export interface ReplaceOptions {
  templatePath: string;
  dataPath: string;
  outputPath?: string;
  delimiters?: { start: string; end: string };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

export function flattenObject(obj: JsonObject, prefix = "", result: JsonObject = {}): JsonObject {
  for (const [key, value] of Object.entries(obj)) {
    const flatKey = prefix ? `${prefix}.${key}` : key;
    if (value !== null && typeof value === "object" && !Array.isArray(value)) {
      result[flatKey] = value;
      flattenObject(value as JsonObject, flatKey, result);
    } else {
      result[flatKey] = value;
    }
    if (!prefix) result[key] = value;
  }
  return result;
}

// ─── CLI entry point ──────────────────────────────────────────────────────────

function printUsage(): void {
  console.log(`
Usage:
  ts-node src/index.ts <template.docx> <data.json> [output.docx]

Arguments:
  template.docx   Path to the Word document with {{placeholders}}
  data.json       Path to the JSON file with replacement values
  output.docx     (optional) Where to write the filled document
                  Defaults to <template>-output.docx in the same folder

Template syntax (inside the .docx, using default {{ }} delimiters):
  {{firstName}}               Simple value
  {{address.city}}            Nested object — dot notation is auto-resolved
  {{#items}}…{{/items}}       Loop over an array ({{name}}, {{.}}, etc.)
  {{#isActive}}…{{/isActive}} Conditional — renders when truthy
  {{^isActive}}…{{/isActive}} Inverted conditional — renders when falsy

Examples:
  ts-node src/index.ts template.docx data.json
  ts-node src/index.ts template.docx data.json filled.docx
`);
}

function main(): void {
  const args = process.argv.slice(2);

  if (args.length < 2 || args.includes("--help") || args.includes("-h")) {
    printUsage();
    process.exit(args.length < 2 ? 1 : 0);
  }

  const [templatePath, dataPath, outputPath] = args;

  try {
    // Step 1 — replace template variables, get back the rendered doc
    const doc = replaceTemplateVariables({ templatePath, dataPath, outputPath });

    // Step 2 — rezip the modified XML back into a .docx buffer
    const outBuffer = doc.getZip().generate({ type: "nodebuffer" });

    // Step 3 — write the buffer to disk
    const absTemplate = path.resolve(templatePath);
    const resolvedOutput =
      outputPath ??
      path.join(
        path.dirname(absTemplate),
        path.basename(absTemplate, ".docx") + "-output.docx"
      );
    fs.writeFileSync(resolvedOutput, outBuffer);

    console.log(`✅ Done! Output written to:\n   ${path.resolve(resolvedOutput)}`);
  } catch (err) {
    console.error(`❌ Error: ${(err as Error).message}`);
    process.exit(1);
  }
}

main();
