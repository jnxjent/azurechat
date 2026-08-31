/**
 * Validates the generated PPTX (ZIP) package without invoking PowerPoint.
 * B4 candidates are accepted only when all internal package references resolve.
 */

import JSZip from "jszip";
import path from "path";

export type PptxValidationResult = {
  valid: boolean;
  errors: string[];
  warnings: string[];
};

export async function validateGeneratedPptx(
  buffer: Buffer,
  options: { expectedSlideCount?: number; context?: string } = {}
): Promise<PptxValidationResult> {
  const errors: string[] = [];
  const warnings: string[] = [];

  if (!buffer || buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    errors.push("Not a valid ZIP file (magic bytes mismatch)");
    return { valid: false, errors, warnings };
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (error) {
    errors.push(`ZIP parse failed: ${error instanceof Error ? error.message : String(error)}`);
    return { valid: false, errors, warnings };
  }

  const fileNames = Object.keys(zip.files);
  for (const requiredFile of [
    "[Content_Types].xml",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
  ]) {
    if (!zip.files[requiredFile]) errors.push(`Missing required file: ${requiredFile}`);
  }
  if (errors.length > 0) return { valid: false, errors, warnings };

  const slideFiles = fileNames.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  if (slideFiles.length === 0) {
    errors.push("No slide XML files found in ppt/slides/");
    return { valid: false, errors, warnings };
  }
  if (
    options.expectedSlideCount !== undefined &&
    slideFiles.length !== options.expectedSlideCount
  ) {
    errors.push(
      `Slide count mismatch: expected ${options.expectedSlideCount}, found ${slideFiles.length}`
    );
  }

  // A dangling internal relationship may cause PowerPoint to repair or reject a deck.
  const relationshipFiles = fileNames.filter((name) => name.endsWith(".rels"));
  for (const relationshipFile of relationshipFiles) {
    try {
      const content = await zip.files[relationshipFile].async("string");
      const relationshipTags = content.match(/<Relationship\b[^>]*>/g) ?? [];
      for (const tag of relationshipTags) {
        if (/\bTargetMode="External"/i.test(tag)) continue;

        const targetMatch = tag.match(/\bTarget="([^"]+)"/i);
        if (!targetMatch) {
          errors.push(`Relationship without Target in ${relationshipFile}`);
          continue;
        }

        let target = targetMatch[1].replace(/\\/g, "/").split("#", 1)[0];
        try {
          target = decodeURIComponent(target);
        } catch {
          // Keep malformed percent escapes intact; the existence check will reject them.
        }

        const markerIndex = relationshipFile.indexOf("/_rels/");
        const ownerDir = relationshipFile === "_rels/.rels"
          ? ""
          : markerIndex >= 0
            ? relationshipFile.slice(0, markerIndex)
            : path.posix.dirname(relationshipFile);
        const resolvedTarget = target.startsWith("/")
          ? path.posix.normalize(target.slice(1))
          : path.posix.normalize(path.posix.join(ownerDir, target));

        if (!zip.files[resolvedTarget]) {
          errors.push(
            `Relationship target not found: ${relationshipFile} -> ${resolvedTarget}`
          );
        }
      }
    } catch (error) {
      errors.push(
        `Could not parse relationship file ${relationshipFile}: ${
          error instanceof Error ? error.message : String(error)
        }`
      );
    }
  }

  try {
    const presentationXml = await zip.files["ppt/presentation.xml"].async("string");
    const slideSize = presentationXml.match(/<p:sldSz[^>]+cx="(\d+)"[^>]+cy="(\d+)"/);
    if (slideSize) {
      const cx = Number.parseInt(slideSize[1], 10);
      const cy = Number.parseInt(slideSize[2], 10);
      const ratio = cx / cy;
      if (Math.abs(ratio - 16 / 9) > 0.05) {
        warnings.push(
          `Slide size ratio ${ratio.toFixed(3)} is not 16:9 (cx=${cx} cy=${cy})`
        );
      }
    } else {
      warnings.push("Could not find slide size in presentation.xml");
    }
  } catch {
    warnings.push("Could not parse presentation.xml for slide size");
  }

  const invalidNumberPattern = /\b(?:NaN|Infinity|-Infinity)\b/;
  const negativeExtentPattern = /\b(?:cx|cy)="-\d+"/;
  const srgbColorPattern = /<a:srgbClr\b[^>]*\bval="([^"]+)"/g;
  for (const slideFile of slideFiles) {
    try {
      const content = await zip.files[slideFile].async("string");
      if (invalidNumberPattern.test(content)) {
        errors.push(`Invalid value (NaN/Infinity) in ${slideFile}`);
      }
      if (negativeExtentPattern.test(content)) {
        errors.push(`Negative width/height extent in ${slideFile}`);
      }
      let colorMatch: RegExpExecArray | null;
      while ((colorMatch = srgbColorPattern.exec(content)) !== null) {
        if (!/^[0-9a-f]{6}$/i.test(colorMatch[1])) {
          errors.push(`Invalid sRGB color "${colorMatch[1]}" in ${slideFile}`);
        }
      }
    } catch (error) {
      errors.push(
        `Could not read ${slideFile}: ${error instanceof Error ? error.message : String(error)}`
      );
    }
  }

  const valid = errors.length === 0;
  const level = valid ? "log" : "warn";
  console[level](
    `[ppt-package] context=${options.context ?? "unspecified"} ` +
    `valid=${valid} errors=${errors.length} warnings=${warnings.length}`
  );
  return { valid, errors, warnings };
}
