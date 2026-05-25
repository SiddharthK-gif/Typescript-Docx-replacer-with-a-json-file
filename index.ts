import * as fs from "fs";
import * as path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

// ─── Types ────────────────────────────────────────────────────────────────────

type JsonValue = string | number | boolean | null | JsonObject | JsonArray;
interface JsonObject { [key: string]: JsonValue }
type JsonArray = JsonValue[];

interface ReplaceOptions {
  /** Path to the .docx template file */
  templatePath: string;
  /** Path to the .json data file */
  dataPath: string;
  /** Output path for the filled .docx (defaults to <template>-output.docx) */
  outputPath?: string;
  /**
   * Delimiter syntax for template variables.
   * Defaults to { start: "{{", end: "}}" }
   */
  delimiters?: { start: string; end: string };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Flattens a nested JSON object so that dot-notation keys work in templates.
 *
 * { address: { city: "SF" } }  →  { "address.city": "SF", address: { city: "SF" } }
 *
 * The original nested structure is preserved alongside flat keys so both
 * {{address.city}} and {{#address}}{{city}}{{/address}} work.
 */
function flattenObject(
  obj: JsonObject,
  prefix = "",
  result: JsonObject = {}
): JsonObject {
  for (const [key, value] of Object.entries(obj)) {
    const flatKey = prefix ? `${prefix}.${key}` : key;

    if (
      value !== null &&
      typeof value === "object" &&
      !Array.isArray(value)
    ) {
      // Keep the nested object for {{#scope}} blocks
      result[flatKey] = value;
      // Also recurse to create dot-notation flat keys
      flattenObject(value as JsonObject, flatKey, result);
    } else {
      result[flatKey] = value;
    }

    // Always keep top-level key as-is for loop/conditional blocks
    if (!prefix) {
      result[key] = value;
    }
  }
  return result;
}

// ─── Core function ────────────────────────────────────────────────────────────

/**
 * Reads a .docx template and a JSON data file, replaces every
 * {{variable}} placeholder with the matching value from the JSON,
 * and writes the result to disk.
 *
 * Supported template syntax (with default {{ }} delimiters):
 *   {{name}}                 Simple scalar value
 *   {{address.city}}         Nested object via dot notation (auto-flattened)
 *   {{#items}}…{{/items}}    Loop over an array  (use {{.}} for scalar items)
 *   {{#flag}}…{{/flag}}      Conditional block (renders when truthy)
 *   {{^flag}}…{{/flag}}      Inverted conditional (renders when falsy / empty)
 */
export function replaceTemplateVariables(options: ReplaceOptions): string {
  const { templatePath, dataPath, delimiters } = options;

  const absTemplate = path.resolve(templatePath);
  const absData     = path.resolve(dataPath);

  if (!fs.existsSync(absTemplate)) throw new Error(`Template file not found: ${absTemplate}`);
  if (!fs.existsSync(absData))     throw new Error(`Data file not found: ${absData}`);

  const outputPath =
    options.outputPath ??
    path.join(
      path.dirname(absTemplate),
      path.basename(absTemplate, ".docx") + "-output.docx"
    );

  // Load and parse template
  const zip = new PizZip(fs.readFileSync(absTemplate));
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks:    true,
    delimiters:    delimiters ?? { start: "{{", end: "}}" },
  });

  // Parse JSON data
  let rawData: JsonObject;
  try {
    rawData = JSON.parse(fs.readFileSync(absData, "utf-8")) as JsonObject;
  } catch (err) {
    throw new Error(`Invalid JSON in ${absData}: ${(err as Error).message}`);
  }

  // Flatten nested objects so dot-notation tags work
  const data = flattenObject(rawData);

  // Render template
  try {
    doc.render(data);
  } catch (err: unknown) {
    const e = err as { properties?: { errors?: Array<{ message?: string }> } };
    if (e?.properties?.errors?.length) {
      const details = e.properties.errors.map((x) => x.message).join("\n  ");
      throw new Error(`Template rendering failed:\n  ${details}`);
    }
    throw err;
  }

  const outBuffer = doc.getZip().generate({ type: "nodebuffer" });
  fs.writeFileSync(outputPath, outBuffer);
  return path.resolve(outputPath);
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
  {{firstName}}              Simple value
  {{address.city}}           Nested object — dot notation is auto-resolved
  {{#items}}…{{/items}}      Loop over an array ({{name}}, {{.}}, etc.)
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
    const out = replaceTemplateVariables({ templatePath, dataPath, outputPath });
    console.log(`✅ Done! Output written to:\n   ${out}`);
  } catch (err) {
    console.error(`❌ Error: ${(err as Error).message}`);
    process.exit(1);
  }
}

main();
