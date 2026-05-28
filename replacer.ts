import * as fs from "fs";
import * as path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { ReplaceOptions, JsonObject, flattenObject } from "./index";

export function replaceTemplateVariables(options: ReplaceOptions): Docxtemplater {
  const { templatePath, dataPath, delimiters } = options;

  const absTemplate = path.resolve(templatePath);
  const absData     = path.resolve(dataPath);

  if (!fs.existsSync(absTemplate)) throw new Error(`Template file not found: ${absTemplate}`);
  if (!fs.existsSync(absData))     throw new Error(`Data file not found: ${absData}`);

  const blob = fs.readFileSync(absTemplate);
  const zip  = new PizZip(blob);
  const doc  = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks:    true,
    delimiters:    delimiters ?? { start: "{{", end: "}}" },
  });

  let rawData: JsonObject;
  try {
    rawData = JSON.parse(fs.readFileSync(absData, "utf-8")) as JsonObject;
  } catch (err) {
    throw new Error(`Invalid JSON in ${absData}: ${(err as Error).message}`);
  }

  const data = flattenObject(rawData);

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

  return doc;
}
