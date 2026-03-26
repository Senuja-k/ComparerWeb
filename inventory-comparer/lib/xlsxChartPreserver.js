import JSZip from 'jszip';

/**
 * ExcelJS does not preserve charts, drawings, or media when loading and saving
 * an xlsx file. This utility restores those elements from the original file
 * into the ExcelJS output by merging at the ZIP level.
 *
 * Instead of assuming sheet index == file name, we build a name→file mapping
 * from the workbook.xml in both ZIPs to correctly pair original sheets with
 * output sheets even when ExcelJS renumbers them.
 *
 * @param {Buffer} originalBuffer - The original xlsx file buffer (with charts)
 * @param {Buffer} excelJsOutputBuffer - The ExcelJS-generated output buffer (charts lost)
 * @param {number} originalSheetCount - Number of sheets in the original file
 * @returns {Promise<Buffer>} - Merged buffer with charts restored
 */
export async function restoreChartsFromOriginal(originalBuffer, excelJsOutputBuffer, originalSheetCount) {
  const origZip = await JSZip.loadAsync(originalBuffer);
  const outZip = await JSZip.loadAsync(excelJsOutputBuffer);

  // 1. Collect chart/drawing/media files from the original
  const chartRelatedFiles = new Map();
  const origPaths = [];
  origZip.forEach((path) => origPaths.push(path));

  for (const path of origPaths) {
    const file = origZip.file(path);
    if (!file || file.dir) continue;

    const isChartRelated =
      path.startsWith('xl/charts/') ||
      path.startsWith('xl/drawings/') ||
      path.startsWith('xl/media/') ||
      path.match(/^xl\/charts\/_rels\//) ||
      path.match(/^xl\/drawings\/_rels\//);

    if (isChartRelated) {
      chartRelatedFiles.set(path, await file.async('nodebuffer'));
    }
  }

  // If no chart-related files exist, nothing to restore
  if (chartRelatedFiles.size === 0) {
    return excelJsOutputBuffer;
  }

  // 2. Build name → sheetN.xml mapping for both original and output
  const origSheetMap = await buildSheetNameMap(origZip);
  const outSheetMap = await buildSheetNameMap(outZip);

  // 3. Copy chart/drawing/media files into the output ZIP
  for (const [path, content] of chartRelatedFiles) {
    outZip.file(path, content);
  }

  // 4. For each original sheet that has drawing relationships,
  //    find the matching output sheet by NAME and restore the references
  for (const [sheetName, origSheetFile] of origSheetMap) {
    const outSheetFile = outSheetMap.get(sheetName);
    if (!outSheetFile) continue; // sheet was removed or renamed; skip

    const origRelsPath = origSheetFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
    const origRelsFile = origZip.file(origRelsPath);
    if (!origRelsFile) continue;

    const origRelsContent = await origRelsFile.async('string');
    const drawingRels = extractDrawingRelationships(origRelsContent);
    if (drawingRels.length === 0) continue;

    const outRelsPath = outSheetFile.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';
    const outRelsFileObj = outZip.file(outRelsPath);
    let drawingRId;

    if (outRelsFileObj) {
      let outRelsContent = await outRelsFileObj.async('string');
      // Check if drawing rels already exist
      if (outRelsContent.includes('/drawing"')) continue;

      const maxId = findMaxRelationshipId(outRelsContent);
      let nextId = maxId + 1;

      for (const rel of drawingRels) {
        const newRId = `rId${nextId}`;
        const newRel = rel.replace(/Id="rId\d+"/, `Id="${newRId}"`);
        outRelsContent = outRelsContent.replace('</Relationships>', `  ${newRel}\n</Relationships>`);
        drawingRId = newRId;
        nextId++;
      }
      outZip.file(outRelsPath, outRelsContent);
    } else {
      let relsContent = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      relsContent += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
      let nextId = 1;

      for (const rel of drawingRels) {
        const newRId = `rId${nextId}`;
        const newRel = rel.replace(/Id="rId\d+"/, `Id="${newRId}"`);
        relsContent += `  ${newRel}\n`;
        drawingRId = newRId;
        nextId++;
      }
      relsContent += '</Relationships>';
      outZip.file(outRelsPath, relsContent);
    }

    // Add <drawing> element back into the output sheet XML
    if (drawingRId) {
      const outSheetXmlFile = outZip.file(outSheetFile);
      if (outSheetXmlFile) {
        let sheetContent = await outSheetXmlFile.async('string');
        if (!sheetContent.includes('<drawing ') && !sheetContent.includes('<drawing>')) {
          sheetContent = insertDrawingElement(sheetContent, drawingRId);
          outZip.file(outSheetFile, sheetContent);
        }
      }
    }

    // Restore <legacyDrawing> if present in original but missing in output
    const origSheetXml = origZip.file(origSheetFile);
    if (origSheetXml) {
      const origContent = await origSheetXml.async('string');
      const outSheetXmlFile = outZip.file(outSheetFile);
      if (outSheetXmlFile) {
        let outContent = await outSheetXmlFile.async('string');
        const legacyMatch = origContent.match(/<legacyDrawing[^/]*\/>/);
        if (legacyMatch && !outContent.includes('<legacyDrawing')) {
          outContent = outContent.replace('</worksheet>', `${legacyMatch[0]}\n</worksheet>`);
          outZip.file(outSheetFile, outContent);
        }
      }
    }
  }

  // 5. Merge [Content_Types].xml — add chart/drawing/media content types from original
  const origCTFile = origZip.file('[Content_Types].xml');
  const outCTFile = outZip.file('[Content_Types].xml');

  if (origCTFile && outCTFile) {
    const origCT = await origCTFile.async('string');
    let outCT = await outCTFile.async('string');

    // Extract Override entries for charts, drawings, media
    const overrideRegex = /<Override[^>]*PartName="\/xl\/(charts|drawings|media)[^"]*"[^>]*\/>/gi;
    let match;
    while ((match = overrideRegex.exec(origCT)) !== null) {
      if (!outCT.includes(match[0])) {
        outCT = outCT.replace('</Types>', `  ${match[0]}\n</Types>`);
      }
    }

    // Also add Default entries for chart-related extensions if missing
    const defaultRegex = /<Default[^>]*Extension="(emf|wmf|vml)"[^>]*\/>/gi;
    while ((match = defaultRegex.exec(origCT)) !== null) {
      if (!outCT.includes(match[0])) {
        outCT = outCT.replace('</Types>', `  ${match[0]}\n</Types>`);
      }
    }

    outZip.file('[Content_Types].xml', outCT);
  }

  return Buffer.from(await outZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}


/**
 * Build a map of sheet name → sheet XML path (e.g. "xl/worksheets/sheet3.xml")
 * by reading workbook.xml and workbook.xml.rels from the ZIP.
 */
async function buildSheetNameMap(zip) {
  const nameToPath = new Map();

  const wbFile = zip.file('xl/workbook.xml');
  const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (!wbFile || !wbRelsFile) return nameToPath;

  const wbXml = await wbFile.async('string');
  const relsXml = await wbRelsFile.async('string');

  // Parse relationships: rId -> target path
  const rIdToTarget = new Map();
  const relRegex = /<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"[^>]*\/>/gi;
  let m;
  while ((m = relRegex.exec(relsXml)) !== null) {
    // Target is relative to xl/, e.g. "worksheets/sheet1.xml"
    rIdToTarget.set(m[1], 'xl/' + m[2]);
  }

  // Parse sheet entries: <sheet name="..." sheetId="..." r:id="rIdN"/>
  const sheetRegex = /<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"[^>]*\/>/gi;
  while ((m = sheetRegex.exec(wbXml)) !== null) {
    const name = m[1];
    const rId = m[2];
    const target = rIdToTarget.get(rId);
    if (target) {
      nameToPath.set(name, target);
    }
  }

  return nameToPath;
}


/**
 * Extract Relationship elements that reference drawings from a .rels XML string.
 */
function extractDrawingRelationships(relsXml) {
  const results = [];
  const regex = /<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/gi;
  let match;
  while ((match = regex.exec(relsXml)) !== null) {
    results.push(match[0]);
  }
  return results;
}


/**
 * Find the highest rId number in a .rels XML string.
 */
function findMaxRelationshipId(relsXml) {
  const regex = /Id="rId(\d+)"/g;
  let maxId = 0;
  let match;
  while ((match = regex.exec(relsXml)) !== null) {
    maxId = Math.max(maxId, parseInt(match[1]));
  }
  return maxId;
}


/**
 * Insert a <drawing> element into a worksheet XML at the correct position.
 * Per OOXML spec, <drawing> appears after most content elements but before
 * <tableParts> and <extLst>.
 */
function insertDrawingElement(sheetXml, rId) {
  const drawingTag = `<drawing r:id="${rId}"/>`;

  // Insert before these elements if they exist (they come after <drawing> in the spec)
  const insertBeforePatterns = [
    '<tableParts',
    '<extLst',
    '</worksheet>',
  ];

  for (const pattern of insertBeforePatterns) {
    const idx = sheetXml.indexOf(pattern);
    if (idx !== -1) {
      return sheetXml.slice(0, idx) + drawingTag + '\n' + sheetXml.slice(idx);
    }
  }

  // Fallback: insert before closing tag
  return sheetXml.replace('</worksheet>', drawingTag + '\n</worksheet>');
}
