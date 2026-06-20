import { useState, useCallback, useEffect, useRef } from "react";
import JSZip from "jszip";

/* ═══════════════════════════════════════════════════════════════════════════
   XML / FILE READING
   ═══════════════════════════════════════════════════════════════════════════ */

function getFileBytes() {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error("File read timed out — try again")), 90000);
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4194304 },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          clearTimeout(timeout);
          return reject(new Error(result.error.message));
        }
        const file = result.value;
        const sliceCount = file.sliceCount;
        console.log(`File has ${sliceCount} slices`);
        const slices = [];
        function getSlice(index) {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status === Office.AsyncResultStatus.Failed) {
              clearTimeout(timeout);
              file.closeAsync();
              return reject(new Error(sliceResult.error.message));
            }
            slices.push(sliceResult.value.data);
            if (index < sliceCount - 1) {
              if (index % 10 === 0) console.log(`Slice ${index}/${sliceCount}`);
              getSlice(index + 1);
            } else {
              clearTimeout(timeout);
              file.closeAsync();
              const byteArrays = slices.map(s =>
                s instanceof ArrayBuffer ? new Uint8Array(s) :
                s instanceof Uint8Array  ? s :
                new Uint8Array(s)
              );
              const total = byteArrays.reduce((n, s) => n + s.length, 0);
              const combined = new Uint8Array(total);
              let offset = 0;
              for (const arr of byteArrays) { combined.set(arr, offset); offset += arr.length; }
              resolve(combined);
            }
          });
        }
        getSlice(0);
      }
    );
  });
}

function parseXml(xmlString) {
  return new DOMParser().parseFromString(xmlString, "application/xml");
}

function resolveThemeColor(ref, themeColors) {
  const map = {
    dk1: "dark1", dk2: "dark2", lt1: "light1", lt2: "light2",
    accent1: "accent1", accent2: "accent2", accent3: "accent3",
    accent4: "accent4", accent5: "accent5", accent6: "accent6",
    hlink: "hyperlink", folHlink: "followedHyperlink",
  };
  return themeColors[map[ref] || ref] || null;
}

function emuToInches(emu) {
  return parseFloat((parseInt(emu, 10) / 914400).toFixed(3));
}

/* ── Parse theme XML ────────────────────────────────────────────────────── */

function parseThemeXml(xml) {
  const doc = parseXml(xml);
  const colors = {};
  const colorMap = {
    dk1: "dark1", dk2: "dark2", lt1: "light1", lt2: "light2",
    accent1: "accent1", accent2: "accent2", accent3: "accent3",
    accent4: "accent4", accent5: "accent5", accent6: "accent6",
  };
  for (const [tag, name] of Object.entries(colorMap)) {
    const el = doc.getElementsByTagNameNS("*", tag)[0];
    if (el) {
      const srgb = el.getElementsByTagNameNS("*", "srgbClr")[0];
      const sys  = el.getElementsByTagNameNS("*", "sysClr")[0];
      if (srgb) colors[name] = "#" + srgb.getAttribute("val");
      else if (sys) colors[name] = "#" + (sys.getAttribute("lastClr") || "000000");
    }
  }
  const majorEl = doc.getElementsByTagNameNS("*", "majorFont")[0];
  const minorEl = doc.getElementsByTagNameNS("*", "minorFont")[0];
  return {
    colors,
    fonts: {
      heading: majorEl?.getElementsByTagNameNS("*", "latin")[0]?.getAttribute("typeface") || null,
      body:    minorEl?.getElementsByTagNameNS("*", "latin")[0]?.getAttribute("typeface") || null,
    },
  };
}

/* ── Parse slide master XML ─────────────────────────────────────────────── */

function parseMasterXml(xml, theme) {
  const doc = parseXml(xml);
  const placeholders = [];
  for (const sp of doc.getElementsByTagNameNS("*", "sp")) {
    const ph     = sp.getElementsByTagNameNS("*", "ph")[0];
    const phType = ph?.getAttribute("type") || "body";
    const phIdx  = ph?.getAttribute("idx") || "0";

    const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
    const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const position = off && ext ? {
      left:   emuToInches(off.getAttribute("x")),
      top:    emuToInches(off.getAttribute("y")),
      width:  emuToInches(ext.getAttribute("cx")),
      height: emuToInches(ext.getAttribute("cy")),
    } : null;

    const txBody   = sp.getElementsByTagNameNS("*", "txBody")[0];
    const lstStyle = txBody?.getElementsByTagNameNS("*", "lstStyle")[0];
    const lvl1pPr  = lstStyle?.getElementsByTagNameNS("*", "lvl1pPr")[0];
    const defRPr   = lvl1pPr?.getElementsByTagNameNS("*", "defRPr")[0];
    const firstPara = txBody?.getElementsByTagNameNS("*", "p")[0];
    const pPr       = firstPara?.getElementsByTagNameNS("*", "pPr")[0];
    const firstRPr  = firstPara?.getElementsByTagNameNS("*", "r")[0]?.getElementsByTagNameNS("*", "rPr")[0];
    const rPr = defRPr || firstRPr;

    // Font name
    let fontName = null;
    if (rPr) {
      const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
      const tf = latin?.getAttribute("typeface");
      if (tf?.startsWith("+mj")) fontName = theme.fonts.heading;
      else if (tf?.startsWith("+mn")) fontName = theme.fonts.body;
      else if (tf) fontName = tf;
    }
    if (!fontName) fontName = (phType === "title" || phType === "ctrTitle") ? theme.fonts.heading : theme.fonts.body;

    // Font size
    let fontSize = null;
    if (defRPr) { const sz = defRPr.getAttribute("sz"); if (sz) fontSize = parseInt(sz, 10) / 100; }
    if (!fontSize) fontSize = (phType === "title" || phType === "ctrTitle") ? 36 : 18;

    // Colour
    let color = null;
    if (rPr) {
      const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
      if (solidFill) {
        const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
        const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
        const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
        if (srgb)    color = "#" + srgb.getAttribute("val");
        else if (sys) color = "#" + (sys.getAttribute("lastClr") || "000000");
        else if (scheme) color = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
      }
    }
    if (!color) color = (phType === "title" || phType === "ctrTitle") ? (theme.colors.dark1 || "#000000") : (theme.colors.dark2 || theme.colors.dark1 || "#000000");

    // Bold
    let bold = null;
    if (rPr) { const b = rPr.getAttribute("b"); bold = b === "1" || b === "true"; }
    if (bold === null) bold = (phType === "title" || phType === "ctrTitle");

    // Alignment
    let alignment = "left";
    const algn = pPr?.getAttribute("algn") || lvl1pPr?.getAttribute("algn");
    if (algn === "ctr") alignment = "center";
    else if (algn === "r") alignment = "right";
    else if (algn === "just") alignment = "justify";

    // Fill
    let masterFill = null;
    const spPrEl = sp.getElementsByTagNameNS("*", "spPr")[0];
    if (spPrEl) {
      if (spPrEl.getElementsByTagNameNS("*", "noFill")[0]) {
        masterFill = "none";
      } else {
        const solidFill = spPrEl.getElementsByTagNameNS("*", "solidFill")[0];
        if (solidFill) {
          const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
          const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
          if (srgb) masterFill = "#" + srgb.getAttribute("val").toUpperCase();
          else if (scheme) {
            const resolved = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
            masterFill = resolved || ("theme:" + scheme.getAttribute("val"));
          }
        }
      }
    }

    // Paragraph formatting per level
    const paraFormat = {};
    if (lstStyle) {
      for (let lvl = 1; lvl <= 9; lvl++) {
        const pPrEl = lstStyle.getElementsByTagNameNS("*", `lvl${lvl}pPr`)[0];
        if (!pPrEl) continue;
        const spcBef    = pPrEl.getElementsByTagNameNS("*", "spcBef")[0];
        const spcAft    = pPrEl.getElementsByTagNameNS("*", "spcAft")[0];
        const spcLin    = pPrEl.getElementsByTagNameNS("*", "lnSpc")[0];
        const buNone    = pPrEl.getElementsByTagNameNS("*", "buNone")[0];
        const buChar    = pPrEl.getElementsByTagNameNS("*", "buChar")[0];
        const buFont    = pPrEl.getElementsByTagNameNS("*", "buFont")[0];
        const buAutoNum = pPrEl.getElementsByTagNameNS("*", "buAutoNum")[0];
        paraFormat[lvl] = {
          indent: pPrEl.getAttribute("indent") ? parseInt(pPrEl.getAttribute("indent")) : null,
          marL:   pPrEl.getAttribute("marL")   ? parseInt(pPrEl.getAttribute("marL"))   : null,
          spcBef: spcBef?.getElementsByTagNameNS("*", "spcPts")[0]?.getAttribute("val")
                    ? parseInt(spcBef.getElementsByTagNameNS("*", "spcPts")[0].getAttribute("val")) / 100 : null,
          spcAft: spcAft?.getElementsByTagNameNS("*", "spcPts")[0]?.getAttribute("val")
                    ? parseInt(spcAft.getElementsByTagNameNS("*", "spcPts")[0].getAttribute("val")) / 100 : null,
          spcLin: spcLin?.getElementsByTagNameNS("*", "spcPct")[0]?.getAttribute("val")
                    ? parseInt(spcLin.getElementsByTagNameNS("*", "spcPct")[0].getAttribute("val")) / 1000 : null,
          bullet: buNone ? "none" : buChar ? buChar.getAttribute("char") : buAutoNum ? "autonumber" : null,
          bulletFont: buFont?.getAttribute("typeface") || null,
        };
      }
    }

    placeholders.push({ type: phType, idx: phIdx, font: { name: fontName, size: fontSize, color, bold }, alignment, position, fill: masterFill, paraFormat });
  }
  return placeholders;
}

/* ── Parse slide XML ────────────────────────────────────────────────────── */

function parseSlideXml(xml, theme, masterPlaceholders, layoutPositions = {}) {
  const doc = parseXml(xml);
  const shapes = [];

  function extractColorFromEl(el) {
    if (!el) return null;
    const solidFill = el.getElementsByTagNameNS("*", "solidFill")[0];
    if (!solidFill) return null;
    const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
    const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
    const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
    if (srgb)   return "#" + srgb.getAttribute("val");
    if (sys)    return "#" + (sys.getAttribute("lastClr") || "000000");
    if (scheme) return resolveThemeColor(scheme.getAttribute("val"), theme.colors);
    return null;
  }

  for (const sp of doc.getElementsByTagNameNS("*", "sp")) {
    const nvSpPr = sp.getElementsByTagNameNS("*", "nvSpPr")[0];
    const cNvPr  = nvSpPr?.getElementsByTagNameNS("*", "cNvPr")[0];
    const ph     = sp.getElementsByTagNameNS("*", "ph")[0];

    const id      = cNvPr?.getAttribute("id") || "";
    const name    = cNvPr?.getAttribute("name") || "";
    const phType  = ph?.getAttribute("type") || "body";
    const phIdx   = ph?.getAttribute("idx") || "0";

    const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
    const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const position = off && ext ? {
      left:   emuToInches(off.getAttribute("x")),
      top:    emuToInches(off.getAttribute("y")),
      width:  emuToInches(ext.getAttribute("cx")),
      height: emuToInches(ext.getAttribute("cy")),
    } : (layoutPositions[`${phType}:${phIdx}`] || layoutPositions["body:0"] || null);

    // Shape fill
    let shapeFill = null, shapeBorder = null;
    const spPr = sp.getElementsByTagNameNS("*", "spPr")[0];
    if (spPr) {
      const direct = Array.from(spPr.childNodes);
      const noFillEl    = direct.find(n => n.localName === "noFill");
      const solidFillEl = direct.find(n => n.localName === "solidFill");
      const gradFillEl  = direct.find(n => n.localName === "gradFill");
      if (noFillEl) {
        shapeFill = "none";
      } else if (solidFillEl) {
        const srgb   = solidFillEl.getElementsByTagNameNS("*", "srgbClr")[0];
        const scheme = solidFillEl.getElementsByTagNameNS("*", "schemeClr")[0];
        const sys    = solidFillEl.getElementsByTagNameNS("*", "sysClr")[0];
        if (srgb)    shapeFill = "#" + srgb.getAttribute("val").toUpperCase();
        else if (sys) shapeFill = "#" + (sys.getAttribute("lastClr") || "000000").toUpperCase();
        else if (scheme) {
          const resolved = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
          shapeFill = resolved ? resolved.toUpperCase() : ("theme:" + scheme.getAttribute("val"));
        }
      } else if (gradFillEl) {
        const firstStop = gradFillEl.getElementsByTagNameNS("*", "gs")[0];
        const srgb = firstStop?.getElementsByTagNameNS("*", "srgbClr")[0];
        shapeFill = srgb ? "#" + srgb.getAttribute("val").toUpperCase() + " (gradient)" : "gradient";
      }
      const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
      if (ln) {
        if (ln.getElementsByTagNameNS("*", "noFill")[0]) {
          shapeBorder = "none";
        } else {
          const lnSolid = ln.getElementsByTagNameNS("*", "solidFill")[0];
          if (lnSolid) {
            const srgb   = lnSolid.getElementsByTagNameNS("*", "srgbClr")[0];
            const scheme = lnSolid.getElementsByTagNameNS("*", "schemeClr")[0];
            const sys    = lnSolid.getElementsByTagNameNS("*", "sysClr")[0];
            if (srgb)    shapeBorder = "#" + srgb.getAttribute("val").toUpperCase();
            else if (sys) shapeBorder = "#" + (sys.getAttribute("lastClr") || "000000").toUpperCase();
            else if (scheme) {
              const resolved = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
              shapeBorder = resolved ? resolved.toUpperCase() : ("theme:" + scheme.getAttribute("val"));
            }
          }
        }
      }
    }

    const txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
    if (!txBody) continue;
    const allRuns  = Array.from(txBody.getElementsByTagNameNS("*", "r"));
    const textContent = allRuns.map(r => r.getElementsByTagNameNS("*", "t")[0]?.textContent || "").join("");
    if (!textContent.trim() && allRuns.length === 0) continue;

    const firstRun = allRuns[0];
    const rPr     = firstRun?.getElementsByTagNameNS("*", "rPr")[0];
    const allParas = Array.from(txBody.getElementsByTagNameNS("*", "p"));

    let fontName = null, fontSize = null, color = null, bold = null, italic = null;

    for (const run of allRuns) { const c = extractColorFromEl(run.getElementsByTagNameNS("*", "rPr")[0]); if (c) { color = c; break; } }
    if (!color) { for (const para of allParas) { const c = extractColorFromEl(para.getElementsByTagNameNS("*", "pPr")[0]?.getElementsByTagNameNS("*", "defRPr")[0]); if (c) { color = c; break; } } }
    if (!color) { const c = extractColorFromEl(txBody.getElementsByTagNameNS("*", "lstStyle")[0]?.getElementsByTagNameNS("*", "lvl1pPr")[0]?.getElementsByTagNameNS("*", "defRPr")[0]); if (c) color = c; }

    if (rPr) {
      const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
      const tf = latin?.getAttribute("typeface");
      if (tf?.startsWith("+mj")) fontName = theme.fonts.heading;
      else if (tf?.startsWith("+mn")) fontName = theme.fonts.body;
      else if (tf) fontName = tf;
      const sz = rPr.getAttribute("sz");
      if (sz) fontSize = parseInt(sz, 10) / 100;
      const b = rPr.getAttribute("b");
      bold   = b === "1" || b === "true" ? true : b === "0" || b === "false" ? false : null;
      const i = rPr.getAttribute("i");
      italic = i === "1" || i === "true" ? true : i === "0" || i === "false" ? false : null;
    }

    const firstPara = txBody.getElementsByTagNameNS("*", "p")[0];
    const pPrEl = firstPara?.getElementsByTagNameNS("*", "pPr")[0];
    let alignment = "left";
    const algn = pPrEl?.getAttribute("algn");
    if (algn === "ctr") alignment = "center";
    else if (algn === "r") alignment = "right";

    const masterPh = masterPlaceholders.find(p => p.type === phType) ||
                     masterPlaceholders.find(p => p.type === "body") ||
                     masterPlaceholders.find(p => p.type !== "title" && p.type !== "ctrTitle") ||
                     null;
    const layoutTargetPos = layoutPositions[`${phType}:${phIdx}`] || layoutPositions[`${phType}:0`] || masterPh?.position || null;

    shapes.push({
      id, name, phType, phIdx, position, shapeFill, shapeBorder,
      textContent: textContent.substring(0, 100),
      current: {
        fontName:  fontName  || "(inherited)",
        fontSize:  fontSize  || "(inherited)",
        color:     color     || "(inherited)",
        bold:      bold      !== null ? bold   : "(inherited)",
        italic:    italic    !== null ? italic : "(inherited)",
        alignment,
      },
      masterTarget: masterPh ? {
        fontName:   masterPh.font.name,
        fontSize:   masterPh.font.size,
        color:      masterPh.font.color,
        bold:       masterPh.font.bold,
        alignment:  masterPh.alignment,
        position:   layoutTargetPos,
        fill:       masterPh.fill || "none",
        paraFormat: masterPh.paraFormat || {},
      } : null,
    });
  }

  // ── Parse tables (graphicFrame elements containing tbl) ──────────────────
  for (const gf of doc.getElementsByTagNameNS("*", "graphicFrame")) {
    const tbl = gf.getElementsByTagNameNS("*", "tbl")[0];
    if (!tbl) continue;

    const nvGfSpPr = gf.getElementsByTagNameNS("*", "nvGraphicFramePr")[0];
    const cNvPr   = nvGfSpPr?.getElementsByTagNameNS("*", "cNvPr")[0];
    const id   = cNvPr?.getAttribute("id") || "";
    const name = cNvPr?.getAttribute("name") || "";

    const xfrm = gf.getElementsByTagNameNS("*", "xfrm")[0];
    const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const position = off && ext ? {
      left:   emuToInches(off.getAttribute("x")),
      top:    emuToInches(off.getAttribute("y")),
      width:  emuToInches(ext.getAttribute("cx")),
      height: emuToInches(ext.getAttribute("cy")),
    } : null;
    if (!position) continue;

    // Collect font info from first cell that has a run
    let fontName = null, fontSize = null, color = null, bold = null, italic = null;
    const allRuns = Array.from(tbl.getElementsByTagNameNS("*", "r"));
    const firstRun = allRuns[0];
    if (firstRun) {
      const rPr = firstRun.getElementsByTagNameNS("*", "rPr")[0];
      if (rPr) {
        const szAttr = rPr.getAttribute("sz");
        if (szAttr) fontSize = parseInt(szAttr, 10) / 100;
        const bAttr = rPr.getAttribute("b");
        if (bAttr !== null) bold = bAttr === "1" || bAttr === "true";
        const iAttr = rPr.getAttribute("i");
        if (iAttr !== null) italic = iAttr === "1" || iAttr === "true";
        const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
        if (solidFill) {
          const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
          const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
          const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
          if (srgb)    color = "#" + srgb.getAttribute("val").toUpperCase();
          else if (sys) color = "#" + (sys.getAttribute("lastClr") || "000000").toUpperCase();
          else if (scheme) color = resolveThemeColor(scheme.getAttribute("val"), theme.colors) || null;
        }
        const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
        if (latin) fontName = latin.getAttribute("typeface") || null;
      }
    }
    const textContent = allRuns.map(r => r.getElementsByTagNameNS("*", "t")[0]?.textContent || "").join(" ").substring(0, 100);

    shapes.push({
      id, name,
      phType: "table", phIdx: "0",
      position, shapeFill: null, shapeBorder: null,
      isTable: true,
      textContent,
      current: {
        fontName:  fontName  || "(inherited)",
        fontSize:  fontSize  || "(inherited)",
        color:     color     || "(inherited)",
        bold:      bold      !== null ? bold   : "(inherited)",
        italic:    italic    !== null ? italic : "(inherited)",
        alignment: null,
      },
      masterTarget: null,
    });
  }

  // ── Parse groups (grpSp elements) — treat as single shape by bounding box ─
  // Only parse top-level groups (direct children of spTree), not nested ones
  const spTree = doc.getElementsByTagNameNS("*", "spTree")[0];
  if (spTree) {
    for (const grp of Array.from(spTree.childNodes).filter(n => n.localName === "grpSp")) {
      const nvGrpSpPr = grp.getElementsByTagNameNS("*", "nvGrpSpPr")[0];
      const cNvPr     = nvGrpSpPr?.getElementsByTagNameNS("*", "cNvPr")[0];
      const id   = cNvPr?.getAttribute("id") || "";
      const name = cNvPr?.getAttribute("name") || "";

      // Group transform gives the bounding box
      const grpSpPr = grp.getElementsByTagNameNS("*", "grpSpPr")[0];
      const xfrm = grpSpPr?.getElementsByTagNameNS("*", "xfrm")[0];
      const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
      const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
      const position = off && ext ? {
        left:   emuToInches(off.getAttribute("x")),
        top:    emuToInches(off.getAttribute("y")),
        width:  emuToInches(ext.getAttribute("cx")),
        height: emuToInches(ext.getAttribute("cy")),
      } : null;
      if (!position) continue;

      // Collect font info from first run inside the group
      let fontName = null, fontSize = null, color = null, bold = null, italic = null;
      const allRuns = Array.from(grp.getElementsByTagNameNS("*", "r"));
      const firstRun = allRuns[0];
      if (firstRun) {
        const rPr = firstRun.getElementsByTagNameNS("*", "rPr")[0];
        if (rPr) {
          const szAttr = rPr.getAttribute("sz");
          if (szAttr) fontSize = parseInt(szAttr, 10) / 100;
          const bAttr = rPr.getAttribute("b");
          if (bAttr !== null) bold = bAttr === "1" || bAttr === "true";
          const iAttr = rPr.getAttribute("i");
          if (iAttr !== null) italic = iAttr === "1" || iAttr === "true";
          const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
          if (solidFill) {
            const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
            const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
            const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
            if (srgb)    color = "#" + srgb.getAttribute("val").toUpperCase();
            else if (sys) color = "#" + (sys.getAttribute("lastClr") || "000000").toUpperCase();
            else if (scheme) color = resolveThemeColor(scheme.getAttribute("val"), theme.colors) || null;
          }
          const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
          if (latin) fontName = latin.getAttribute("typeface") || null;
        }
      }
      const textContent = allRuns.map(r => r.getElementsByTagNameNS("*", "t")[0]?.textContent || "").join(" ").substring(0, 100);

      shapes.push({
        id, name,
        phType: "group", phIdx: "0",
        position, shapeFill: null, shapeBorder: null,
        isGroup: true,
        textContent,
        current: {
          fontName:  fontName  || "(inherited)",
          fontSize:  fontSize  || "(inherited)",
          color:     color     || "(inherited)",
          bold:      bold      !== null ? bold   : "(inherited)",
          italic:    italic    !== null ? italic : "(inherited)",
          alignment: null,
        },
        masterTarget: null,
      });
    }
  }

  return shapes;
}

async function readAllMasters(zip) {
  const relsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const masters = [];

  if (relsFile) {
    const relsXml = await relsFile.async("string");
    const relsDoc = parseXml(relsXml);
    for (const rel of relsDoc.getElementsByTagNameNS("*", "Relationship")) {
      const target = rel.getAttribute("Target") || "";
      const match = target.match(/slideMasters\/slideMaster(\d+)\.xml/);
      if (!match) continue;
      const masterIndex = parseInt(match[1], 10);
      const masterPath  = `ppt/slideMasters/slideMaster${masterIndex}.xml`;
      const masterRelsPath = `ppt/slideMasters/_rels/slideMaster${masterIndex}.xml.rels`;
      let theme = { colors: {}, fonts: { heading: null, body: null } };
      const masterRelsFile = zip.file(masterRelsPath);
      if (masterRelsFile) {
        const masterRelsXml = await masterRelsFile.async("string");
        const masterRelsDoc = parseXml(masterRelsXml);
        for (const mRel of masterRelsDoc.getElementsByTagNameNS("*", "Relationship")) {
          const themeMatch = (mRel.getAttribute("Target") || "").match(/\.\.\/theme\/theme(\d+)\.xml/);
          if (themeMatch) {
            const themeFile = zip.file(`ppt/theme/theme${themeMatch[1]}.xml`);
            if (themeFile) theme = parseThemeXml(await themeFile.async("string"));
            break;
          }
        }
      }
      const masterFile = zip.file(masterPath);
      if (!masterFile) continue;
      const masterXml = await masterFile.async("string");
      const masterDoc = parseXml(masterXml);
      const masterName = masterDoc.getElementsByTagNameNS("*", "cSld")[0]?.getAttribute("name") || `Master ${masterIndex}`;
      masters.push({
        index: masterIndex,
        name: masterName,
        theme,
        placeholders: parseMasterXml(masterXml, theme),
        headingFont: theme.fonts.heading,
        bodyFont: theme.fonts.body,
        colors: Object.entries(theme.colors).filter(([, v]) => v).slice(0, 5).map(([, v]) => v),
      });
    }
  }

  // Count slides per master — reject imported masters used by < 20% of slides vs dominant
  const masterSlideCounts = {};
  for (const m of masters) masterSlideCounts[m.index] = 0;
  for (const slideFile of zip.file(/^ppt\/slides\/slide\d+\.xml$/)) {
    const slideNum = slideFile.name.match(/slide(\d+)\.xml/)?.[1];
    if (!slideNum) continue;
    const relsFile = zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`);
    if (!relsFile) continue;
    try {
      const relsXml = await relsFile.async("string");
      const relsDoc = parseXml(relsXml);
      for (const rel of relsDoc.getElementsByTagNameNS("*", "Relationship")) {
        const layoutMatch = (rel.getAttribute("Target") || "").match(/slideLayouts\/slideLayout(\d+)\.xml/);
        if (!layoutMatch) continue;
        const layoutRelsFile = zip.file(`ppt/slideLayouts/_rels/slideLayout${layoutMatch[1]}.xml.rels`);
        if (!layoutRelsFile) continue;
        const layoutRelsXml = await layoutRelsFile.async("string");
        for (const lrel of parseXml(layoutRelsXml).getElementsByTagNameNS("*", "Relationship")) {
          const masterMatch = (lrel.getAttribute("Target") || "").match(/slideMasters\/slideMaster(\d+)\.xml/);
          if (masterMatch) {
            const idx = parseInt(masterMatch[1], 10);
            if (masterSlideCounts[idx] !== undefined) masterSlideCounts[idx]++;
            break;
          }
        }
        break;
      }
    } catch (e) { /* skip */ }
  }
  console.log("Master slide counts:", masterSlideCounts);
  const maxCount = Math.max(...Object.values(masterSlideCounts), 0);
  const filtered = masters.filter(m => {
    const count = masterSlideCounts[m.index] || 0;
    const keep = maxCount === 0 || count >= Math.max(1, maxCount * 0.2);
    if (!keep) console.log(`Rejecting imported master ${m.index} ("${m.name}") — ${count} slide(s) vs dominant ${maxCount}`);
    return keep;
  });
  return filtered.length > 0 ? filtered : masters;
}

/* ── Phase 1: read file bytes + discover masters ────────────────────────── */

async function readPptxFile() {
  const bytes = await getFileBytes();
  const zip = await JSZip.loadAsync(bytes);
  const masters = await readAllMasters(zip);
  return { zip, masters };
}

/* ── Phase 2: read slide data using chosen master ───────────────────────── */

async function readSlideWithMaster(zip, masters, chosenMasterIndex, selectedSlideIndex) {
  const slideFile = zip.file(`ppt/slides/slide${selectedSlideIndex}.xml`);
  if (!slideFile) throw new Error(`Slide ${selectedSlideIndex} not found in file`);
  const slideXml = await slideFile.async("string");

  // Always use master 1's layouts — ignore whatever imported layout the slide references
  let layoutPositions = {};

  const masterRelsFile = zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels");
  if (masterRelsFile) {
    const masterRelsXml = await masterRelsFile.async("string");
    const masterRels = parseXml(masterRelsXml).getElementsByTagNameNS("*", "Relationship");
    const titlePositions = [];
    let firstLayoutDone = false;

    for (const rel of masterRels) {
      const layoutMatch = (rel.getAttribute("Target") || "").match(/slideLayouts\/slideLayout(\d+)\.xml/);
      if (!layoutMatch) continue;
      const layoutFile = zip.file(`ppt/slideLayouts/slideLayout${layoutMatch[1]}.xml`);
      if (!layoutFile) continue;
      const layoutDoc = parseXml(await layoutFile.async("string"));
      let layoutTitlePos = null, hasBody = false;

      for (const sp of layoutDoc.getElementsByTagNameNS("*", "sp")) {
        const ph = sp.getElementsByTagNameNS("*", "ph")[0];
        if (!ph) continue;
        const phType = ph.getAttribute("type") || "body";
        const phIdx  = ph.getAttribute("idx") || "0";
        const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
        const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
        const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
        if (!off || !ext) continue;
        const pos = {
          left: emuToInches(off.getAttribute("x")),   top:    emuToInches(off.getAttribute("y")),
          width: emuToInches(ext.getAttribute("cx")), height: emuToInches(ext.getAttribute("cy")),
        };
        if (!firstLayoutDone && phType !== "title" && phType !== "ctrTitle") layoutPositions[`${phType}:${phIdx}`] = pos;
        if (phType === "title" || phType === "ctrTitle") layoutTitlePos = pos;
        if (phType === "body" || phType === "obj" || phIdx === "1") hasBody = true;
      }
      firstLayoutDone = true;
      if (layoutTitlePos && hasBody) titlePositions.push(layoutTitlePos);
    }

    // Title position: master placeholder is most authoritative, fall back to layout vote
    const masterTitlePh = masters.find(m => m.index === 1)?.placeholders?.find(p => p.type === "title" || p.type === "ctrTitle");
    if (masterTitlePh?.position) {
      layoutPositions["title:0"] = masterTitlePh.position;
      console.log(`Title position from master: ${JSON.stringify(masterTitlePh.position)}`);
    } else if (titlePositions.length > 0) {
      const freq = {};
      for (const p of titlePositions) {
        const k = `${p.left.toFixed(2)},${p.top.toFixed(2)},${p.width.toFixed(2)},${p.height.toFixed(2)}`;
        freq[k] = (freq[k] || 0) + 1;
      }
      const topKey = Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
      const [left, top, width, height] = topKey.split(",").map(Number);
      layoutPositions["title:0"] = { left, top, width, height };
      console.log(`Title position from ${titlePositions.length} layouts: ${topKey}`);
    }
  }

  const master = masters.find(m => m.index === chosenMasterIndex) || masters[0];
  const slideShapes = parseSlideXml(slideXml, master.theme, master.placeholders, layoutPositions);
  console.log(`Using master ${master.index} ("${master.name}") for slide ${selectedSlideIndex}`);
  console.log("Layout positions:", JSON.stringify(layoutPositions));
  return { theme: master.theme, masterPlaceholders: master.placeholders, masterName: master.name, slideShapes, layoutPositions, slideIndex: selectedSlideIndex };
}

/* ═══════════════════════════════════════════════════════════════════════════
   OFFICE JS — slide index + apply fixes
   ═══════════════════════════════════════════════════════════════════════════ */

function getSelectedSlideIndex() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) return reject(new Error(result.error.message));
      resolve(result.value.slides[0].index);
    });
  });
}

// Captures the currently selected shape on the slide, for manual title override
async function captureSelectedShape() {
  return PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedShapes();
    sel.load("items");
    await ctx.sync();
    if (sel.items.length === 0) return null;
    const shape = sel.items[0];
    shape.load(["id", "name"]);
    await ctx.sync();
    return { id: shape.id, name: shape.name };
  });
}

/* ── Colour utilities ────────────────────────────────────────────────────── */

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return { r: parseInt(h.substring(0, 2), 16), g: parseInt(h.substring(2, 4), 16), b: parseInt(h.substring(4, 6), 16) };
}

function colourDistance(hex1, hex2) {
  const a = hexToRgb(hex1), b = hexToRgb(hex2);
  return Math.sqrt(2 * Math.pow(a.r-b.r, 2) + 4 * Math.pow(a.g-b.g, 2) + 3 * Math.pow(a.b-b.b, 2));
}

function hexLuminance(hex) {
  const h = hex.replace("#", "");
  const toLinear = c => c <= 0.03928 ? c/12.92 : Math.pow((c+0.055)/1.055, 2.4);
  return 0.2126*toLinear(parseInt(h.slice(0,2),16)/255) + 0.7152*toLinear(parseInt(h.slice(2,4),16)/255) + 0.0722*toLinear(parseInt(h.slice(4,6),16)/255);
}

function snapToThemeColor(hex, themeColors) {
  if (!hex || hex === "none" || hex.startsWith("theme:")) return hex;
  let nearest = null, nearestDist = Infinity;
  for (const [, themeHex] of Object.entries(themeColors)) {
    if (!themeHex) continue;
    const dist = colourDistance(hex, themeHex);
    if (dist < nearestDist) { nearestDist = dist; nearest = themeHex; }
  }
  return nearest || hex;
}

/* ── Apply fixes via Office JS ──────────────────────────────────────────── */

async function applyFixes(slideIndex, fixes, themeColors = {}) {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    const slide = slides.items[slideIndex - 1];
    const shapes = slide.shapes;
    shapes.load("items");
    await ctx.sync();
    for (const shape of shapes.items) shape.load(["id", "name", "left", "top", "width", "height"]);
    await ctx.sync();

    const originalPositions = new Map();
    for (const s of shapes.items) originalPositions.set(String(s.id), { left: s.left, top: s.top, width: s.width, height: s.height });

    for (const fix of fixes) {
      const lookupId = fix._slideShape?.id || fix.shapeId;
      const target = lookupId
        ? shapes.items.find(s => String(s.id) === String(lookupId))
        : shapes.items.find(s => s.name === fix.shapeName);
      console.log(`Fix for "${fix.shapeName}" → ${target ? "FOUND" : "NOT FOUND"}`);
      if (!target) continue;

      if (fix.position) {
        const inchToPt = 72;
        const { left, top, width, height } = fix.position;
        const orig = originalPositions.get(String(target.id)) || {};
        if (left   !== undefined && left   >= 0  && left   < 13)  { console.log(`  left: ${orig.left} → ${left * inchToPt}`);  target.left   = left   * inchToPt; }
        if (top    !== undefined && top    >= 0  && top    < 8)   { console.log(`  top: ${orig.top} → ${top * inchToPt}`);     target.top    = top    * inchToPt; }
        if (width  !== undefined && width  > 0.5 && width  <= 14) { target.width  = width  * inchToPt; }
        if (height !== undefined && height > 0.1 && height <= 8)  { target.height = height * inchToPt; }
      }

      if (fix.fill !== undefined && fix.fill !== null) {
        try {
          if (fix.fill === "none" && fix.shapeFill && fix.shapeFill !== "none") { target.fill.clear(); console.log(`  ✓ Fill cleared`); }
          else if (fix.fill?.startsWith("#")) { const snapped = snapToThemeColor(fix.fill, themeColors); target.fill.setSolidColor(snapped.replace("#", "")); }
        } catch (e) { console.log(`  Error setting fill:`, e.message); }
      }

      if (fix.border !== undefined && fix.border !== null) {
        try {
          if (fix.border === "none" && fix.shapeBorder && fix.shapeBorder !== "none") {
            const isExactTheme = Object.values(themeColors).some(c => c && c.toUpperCase() === fix.shapeBorder.toUpperCase());
            const snapped = snapToThemeColor(fix.shapeBorder, themeColors);
            if (isExactTheme) { /* keep */ }
            else if (snapped !== fix.shapeBorder) { target.lineFormat.color = snapped.replace("#", ""); target.lineFormat.visible = true; }
            else { target.lineFormat.visible = false; }
          }
        } catch (e) { console.log(`  Error setting border:`, e.message); }
      }

      if (fix.font || fix.alignment) {
        try {
          if (fix._slideShape?.isTable) {
            // Tables: apply font/colour to every cell
            const table = target.table;
            table.load("rowCount,columnCount");
            await ctx.sync();
            for (let r = 0; r < table.rowCount; r++) {
              for (let c = 0; c < table.columnCount; c++) {
                const cell = table.getCell(r, c);
                const tr = cell.textFrame.textRange;
                try {
                  if (fix.font?.name)  tr.font.name = fix.font.name;
                  if (fix.font?.size)  tr.font.size = fix.font.size;
                  if (fix.font?.color) {
                    let textColor = snapToThemeColor(fix.font.color, themeColors);
                    tr.font.color = textColor.replace("#", "");
                  }
                } catch (e) { /* empty cell */ }
              }
            }
            await ctx.sync();
          } else if (fix._slideShape?.isGroup) {
            // Groups: apply font/colour to all child shapes
            try {
              const groupShapes = target.shapes;
              groupShapes.load("items");
              await ctx.sync();
              for (const child of groupShapes.items) {
                try {
                  const tr = child.textFrame.textRange;
                  if (fix.font?.name)  tr.font.name = fix.font.name;
                  if (fix.font?.size)  tr.font.size = fix.font.size;
                  if (fix.font?.color) {
                    let textColor = snapToThemeColor(fix.font.color, themeColors);
                    tr.font.color = textColor.replace("#", "");
                  }
                  await ctx.sync();
                } catch (e) { /* no text */ }
              }
            } catch (e) { /* no group */ }
          } else {
            const tr = target.textFrame.textRange;
            tr.load(["text"]);
            await ctx.sync();
            if (fix.font) {
              if (fix.font.name)  tr.font.name = fix.font.name;
              if (fix.font.size)  tr.font.size = fix.font.size;
              if (fix.font.color) {
                let textColor = snapToThemeColor(fix.font.color, themeColors);
                const fill = fix.shapeFill;
                if (fill && fill !== "none") {
                  const contrast = (Math.max(hexLuminance(fill), hexLuminance(textColor)) + 0.05) / (Math.min(hexLuminance(fill), hexLuminance(textColor)) + 0.05);
                  if (contrast < 3) textColor = hexLuminance(fill) > 0.179 ? "#000000" : "#FFFFFF";
                }
                tr.font.color = textColor.replace("#", "");
              }
            }
            if (fix.alignment) {
              const alignMap = { left: PowerPoint.ParagraphHorizontalAlignment.left, center: PowerPoint.ParagraphHorizontalAlignment.center, right: PowerPoint.ParagraphHorizontalAlignment.right };
              if (alignMap[fix.alignment]) tr.paragraphFormat.horizontalAlignment = alignMap[fix.alignment];
            }
            await ctx.sync();
          }
        } catch (e) { console.log(`  Error applying font:`, e.message); }
      }
    }
    await ctx.sync();
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════════════════════════════════════ */

function LogLine({ entry }) {
  const color = entry.msg.startsWith("✓") ? "#4ade80" : entry.msg.startsWith("✗") ? "#f87171" : "#94a3b8";
  return (
    <div style={{ display: "flex", gap: 8, marginBottom: 3 }}>
      <span style={{ fontSize: 9, color: "#475569", fontFamily: "monospace", flexShrink: 0 }}>{entry.time}</span>
      <span style={{ fontSize: 10, color, fontFamily: "monospace" }}>{entry.msg}</span>
    </div>
  );
}

function ThemeCard({ theme, masterPlaceholders }) {
  if (!theme) return null;
  const colors = Object.entries(theme.colors).filter(([, v]) => v).slice(0, 6);
  return (
    <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>Detected from file</div>
      <div style={{ display: "flex", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
        {theme.fonts.heading && <span style={{ fontSize: 10, background: "#eff6ff", color: "#1e40af", padding: "2px 8px", borderRadius: 20, border: "1px solid #bfdbfe" }}>Aa {theme.fonts.heading}</span>}
        {theme.fonts.body && theme.fonts.body !== theme.fonts.heading && <span style={{ fontSize: 10, background: "#f5f3ff", color: "#6d28d9", padding: "2px 8px", borderRadius: 20, border: "1px solid #ddd6fe" }}>Aa {theme.fonts.body}</span>}
      </div>
      <div style={{ display: "flex", gap: 4, marginBottom: 8, flexWrap: "wrap" }}>
        {colors.map(([k, v]) => (
          <div key={k} title={`${k}: ${v}`} style={{ display: "flex", alignItems: "center", gap: 3 }}>
            <div style={{ width: 14, height: 14, borderRadius: 3, background: v, border: "1px solid rgba(0,0,0,0.1)" }} />
            <span style={{ fontSize: 9, color: "#6b7280", fontFamily: "monospace" }}>{v}</span>
          </div>
        ))}
      </div>
      <div style={{ fontSize: 10, color: "#9ca3af" }}>{masterPlaceholders.length} master placeholder{masterPlaceholders.length !== 1 ? "s" : ""} · read from XML</div>
    </div>
  );
}

function FixBadge({ count }) {
  return (
    <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "10px 14px", display: "flex", alignItems: "center", gap: 10 }}>
      <span style={{ fontSize: 20 }}>✓</span>
      <div>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#166534" }}>Cleanup complete</div>
        <div style={{ fontSize: 11, color: "#15803d" }}>{count} shape{count !== 1 ? "s" : ""} reformatted · use Ctrl+Z to undo</div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   MAIN APP
   ═══════════════════════════════════════════════════════════════════════════ */

export default function App() {
  const [status, setStatus]               = useState("idle");
  const [log, setLog]                     = useState([]);
  const [fixCount, setFixCount]           = useState(0);
  const [error, setError]                 = useState(null);
  const [detectedTheme, setDetectedTheme] = useState(null);
  const [detectedMaster, setDetectedMaster] = useState([]);

  // Manual title override, per slide, for this session only — { [slideIndex]: { id, name } }
  const [titleOverrides, setTitleOverrides] = useState({});
  const [overrideStatus, setOverrideStatus] = useState(null); // transient feedback message

  // Cached file data — loaded once in background, reused for every Fix click
  const [fileReady, setFileReady]   = useState(false);
  const [fileError, setFileError]   = useState(null);
  const cachedZip     = useRef(null);
  const cachedMasters = useRef(null);

  // Load the file in the background as soon as the add-in opens
  useEffect(() => {
    let cancelled = false;
    async function loadFile() {
      try {
        const { zip, masters } = await readPptxFile();
        if (cancelled) return;
        cachedZip.current     = zip;
        cachedMasters.current = masters;
        setFileReady(true);
      } catch (e) {
        if (cancelled) return;
        setFileError(e.message);
      }
    }
    loadFile();
    return () => { cancelled = true; };
  }, []);

  const addLog = (msg) => setLog(l => [...l, {
    time: new Date().toLocaleTimeString("en-US", { hour12: false, hour: "2-digit", minute: "2-digit", second: "2-digit" }),
    msg,
  }]);

  const handleSetTitleOverride = useCallback(async () => {
    setOverrideStatus(null);
    try {
      const slideIndex = await getSelectedSlideIndex();
      const shape = await captureSelectedShape();
      if (!shape) {
        setOverrideStatus({ ok: false, msg: "No shape selected — click a shape on the slide first" });
        return;
      }
      setTitleOverrides(prev => ({ ...prev, [slideIndex]: shape }));
      setOverrideStatus({ ok: true, msg: `"${shape.name}" set as title for slide ${slideIndex}` });
    } catch (e) {
      setOverrideStatus({ ok: false, msg: e.message });
    }
  }, []);

  const handleCleanup = useCallback(async () => {
    setStatus("running");
    setLog([]);
    setError(null);
    setFixCount(0);
    setDetectedTheme(null);
    setDetectedMaster([]);

    try {
      addLog("Reading selected slide…");
      const slideIndex = await getSelectedSlideIndex();
      const slideW = 13.33, slideH = 7.5; // standard WIDE layout inches
      addLog(`Slide ${slideIndex} selected`);

      // Always read the file fresh so we get current slide positions.
      // Masters are cached since they don't change between runs.
      addLog("Reading .pptx file…");
      let zip, masters;
      ({ zip, masters } = await readPptxFile());
      if (cachedMasters.current) {
        masters = cachedMasters.current; // reuse parsed masters, fresh zip
      } else {
        cachedMasters.current = masters;
      }

      if (masters.length === 0) throw new Error("No slide masters found in this file");
      const primaryMaster = masters[0];
      const pptxData = await readSlideWithMaster(zip, masters, primaryMaster.index, slideIndex);
      setDetectedTheme(pptxData.theme);
      setDetectedMaster(pptxData.masterPlaceholders);

      // Apply manual title override for this slide, if one was set this session
      const override = titleOverrides[slideIndex];
      if (override) {
        const alreadyTitle = pptxData.slideShapes.find(s => s.phType === "title" || s.phType === "ctrTitle");
        const target = pptxData.slideShapes.find(s => String(s.id) === String(override.id))
          || pptxData.slideShapes.find(s => s.name === override.name);
        if (target) {
          if (alreadyTitle && alreadyTitle !== target) alreadyTitle.phType = "body"; // demote previous auto-detected title
          target.phType = "title";
          addLog(`Using manual title override: "${target.name}"`);
        } else {
          addLog(`⚠ Title override shape not found on this slide (looked for id="${override.id}" name="${override.name}") — using automatic detection`);
        }
      }

      // ─── PRE-STEP: Strip paragraph overrides directly in zip XML ────────────
      addLog("Resetting paragraph formatting…");
      try {
        const slidePath = `ppt/slides/slide${slideIndex}.xml`;
        const slideZipFile = zip.file(slidePath);
        if (slideZipFile) {
          let xml = await slideZipFile.async("string");
          const before = xml;
          xml = xml.replace(/<a:buNone\s*\/>/g, "");
          xml = xml.replace(/<a:buChar[^/]*\/>/g, "");
          xml = xml.replace(/<a:buFont[^/>]*(?:\/>|>[\s\S]*?<\/a:buFont>)/g, "");
          xml = xml.replace(/<a:buClr>[\s\S]*?<\/a:buClr>/g, "");
          xml = xml.replace(/<a:buSzPct[^/]*\/>/g, "");
          xml = xml.replace(/<a:buSzPts[^/]*\/>/g, "");
          xml = xml.replace(/<a:buAutoNum[^/]*\/>/g, "");
          xml = xml.replace(/(<a:pPr[^>]*?)\s+indent="[^"]*"/g, "$1");
          xml = xml.replace(/(<a:pPr[^>]*?)\s+marL="[^"]*"/g, "$1");
          if (xml !== before) {
            zip.file(slidePath, xml);
            addLog("✓ Paragraph formatting reset — saving…");
            const newPptxBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
            await new Promise((resolve, reject) => {
              Office.context.document.setFileAsync(newPptxBytes.buffer, { fileType: Office.FileType.Compressed }, (result) => {
                result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result.error?.message || "setFileAsync failed"));
              });
            });
          } else {
            addLog("No paragraph overrides found");
          }
        }
      } catch (e) {
        addLog(`⚠ Paragraph reset skipped: ${e.message}`);
      }

      const dupIndex = slideIndex; // working on original — Ctrl+Z to undo
      const themeColors = pptxData.theme.colors;
      const themeColorList = Object.values(themeColors).filter(v => v);
      let totalFixes = 0;

      // ─── STEP 1: Title — snap to master position and font ───────────────────
      const titleShape    = pptxData.slideShapes.find(s => s.phType === "title" || s.phType === "ctrTitle");
      const titleMaster   = pptxData.masterPlaceholders.find(p => p.type === "title" || p.type === "ctrTitle");
      const layoutTitlePos = pptxData.layoutPositions?.["title:0"];
      const targetTitlePos = layoutTitlePos || titleMaster?.position;

      if (titleShape && targetTitlePos) {
        const cur = titleShape.position;
        const posNeedsFix = !cur ||
          Math.abs(cur.left   - targetTitlePos.left)   > targetTitlePos.width  * 0.005 ||
          Math.abs(cur.top    - targetTitlePos.top)    > targetTitlePos.height * 0.005 ||
          Math.abs(cur.width  - targetTitlePos.width)  > targetTitlePos.width  * 0.005 ||
          Math.abs(cur.height - targetTitlePos.height) > targetTitlePos.height * 0.005;
        const fontNeedsFix = titleShape.current.fontName !== "(inherited)" && titleShape.current.fontName !== titleShape.masterTarget?.fontName;
        const fillNeedsFix = titleShape.shapeFill && titleShape.shapeFill !== "none" && titleShape.masterTarget?.fill === "none";
        if (posNeedsFix || fontNeedsFix || fillNeedsFix) {
          addLog("Step 1: Title position & font…");
          await applyFixes(dupIndex, [{
            shapeName: titleShape.name, shapeId: titleShape.id, _slideShape: titleShape,
            shapeFill: fillNeedsFix ? "none" : (titleShape.shapeFill || null),
            ...(posNeedsFix  ? { position: targetTitlePos } : {}),
            ...(fontNeedsFix ? { font: { name: titleShape.masterTarget?.fontName } } : {}),
          }], themeColors);
          totalFixes++;
        }
      }

      // ─── STEP 2: Fonts — correct font name, normalise sizes, expand text boxes ─
      addLog("Step 2: Fonts…");
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const slide  = slides.items[dupIndex - 1];
        const shapes = slide.shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name", "type"]);
        await ctx.sync();

        const nonTitleSizes = pptxData.slideShapes
          .filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && typeof ss.current.fontSize === "number")
          .map(ss => ss.current.fontSize);
        const sizeFreq = nonTitleSizes.reduce((acc, s) => { acc[s] = (acc[s] || 0) + 1; return acc; }, {});
        const normalisedSize = nonTitleSizes.length > 0 ? parseInt(Object.entries(sizeFreq).sort((a, b) => b[1] - a[1])[0][0]) : null;

        for (const ss of pptxData.slideShapes) {
          if (ss.isTable) {
            const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
            if (!os) continue;
            try {
              const table = os.table;
              table.load("rowCount,columnCount");
              await ctx.sync();
              const masterFont = pptxData.masterPlaceholders.find(p => p.type === "body")?.font;
              const cells = [];
              for (let r = 0; r < table.rowCount; r++)
                for (let c = 0; c < table.columnCount; c++) {
                  const cell = table.getCell(r, c);
                  const tr = cell.textFrame.textRange;
                  tr.font.load(["name", "size"]);
                  cells.push({ tr });
                }
              await ctx.sync();
              for (const { tr } of cells) {
                try {
                  if (masterFont?.name && tr.font.name !== masterFont.name) tr.font.name = masterFont.name;
                  if (normalisedSize && typeof ss.current.fontSize === "number" && Math.abs(ss.current.fontSize - normalisedSize) <= 3) tr.font.size = normalisedSize;
                } catch (e) { /* empty cell */ }
              }
              await ctx.sync();
            } catch (e) { /* no table */ }
            continue;
          }
          if (ss.isGroup) {
            const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
            if (!os) continue;
            try {
              const masterFont = pptxData.masterPlaceholders.find(p => p.type === "body")?.font;
              const groupShapes = os.shapes;
              groupShapes.load("items");
              await ctx.sync();
              const trs = [];
              for (const child of groupShapes.items) {
                try { const tr = child.textFrame.textRange; tr.font.load(["name", "size"]); trs.push({ tr }); } catch (e) { /* no text */ }
              }
              await ctx.sync();
              for (const { tr } of trs) {
                try {
                  if (masterFont?.name && tr.font.name !== masterFont.name) tr.font.name = masterFont.name;
                  if (normalisedSize && typeof ss.current.fontSize === "number" && Math.abs(ss.current.fontSize - normalisedSize) <= 3) tr.font.size = normalisedSize;
                } catch (e) { /* no text */ }
              }
              await ctx.sync();
            } catch (e) { /* no group */ }
            continue;
          }
          const osShape = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!osShape) continue;
          try {
            const tr = osShape.textFrame.textRange;
            tr.font.load(["name", "size"]);
            try { await ctx.sync(); } catch (e) { continue; }
            const isTitle = ss.phType === "title" || ss.phType === "ctrTitle";
            if (isTitle) continue;
            const bodyFont = pptxData.masterPlaceholders.find(p => p.type === "body")?.font;
            const targetFont = ss.masterTarget?.fontName || bodyFont?.name;
            if (!targetFont) continue;
            let changed = false;
            if (tr.font.name !== targetFont) { tr.font.name = targetFont; changed = true; }
            const currentSize = typeof ss.current.fontSize === "number" ? ss.current.fontSize : tr.font.size;
            if (normalisedSize && currentSize !== null && Math.abs(currentSize - normalisedSize) <= 3 && tr.font.size !== normalisedSize) { tr.font.size = normalisedSize; changed = true; }
            if (changed) { await ctx.sync(); totalFixes++; }
          } catch (e) { /* shape may not support font ops */ }
        }

        // Expand text boxes to fit content
        for (const ss of pptxData.slideShapes) {
          if (!ss.position || !ss.textContent) continue;
          if (ss.isTable || ss.isGroup) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          // Skip shapes without text frames (connectors, images, etc.)
          if (os.type && !["GeometricShape", "Placeholder"].some(t => os.type.toString().includes(t))) continue;
          try {
            os.load(["left", "top", "width", "height"]);
            os.textFrame.load("autoSizeSetting");
            try { await ctx.sync(); } catch (e) { continue; }
            os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
            await ctx.sync();
            os.load(["width", "height"]);
            await ctx.sync();
            const neededW = os.width / 72, neededH = os.height / 72;
            os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            await ctx.sync();
            const { left: origLeft, top: origTop, width: origW, height: origH } = ss.position;
            const stepW = slideW / 100, stepH = slideH / 100;
            if (neededW <= origW + stepW * 0.1 && neededH <= origH + stepH * 0.1) continue;
            const maxW = slideW - origLeft, maxH = slideH - origTop;
            let curW = origW, curH = origH;
            while ((curW < neededW - stepW * 0.1 || curH < neededH - stepH * 0.1) && (curW < maxW - stepW * 0.1 || curH < maxH - stepH * 0.1)) {
              if (curW < neededW - stepW * 0.1 && curW < maxW) curW = Math.min(curW + stepW, maxW);
              if (curH < neededH - stepH * 0.1 && curH < maxH) curH = Math.min(curH + stepH, maxH);
            }
            let overlaps = false;
            for (const other of pptxData.slideShapes) {
              if (String(other.id) === String(ss.id) || !other.position) continue;
              const o = other.position;
              if (origLeft < o.left + o.width && origLeft + curW > o.left && origTop < o.top + o.height && origTop + curH > o.top) { overlaps = true; break; }
            }
            if (overlaps || origLeft + curW > slideW + stepW * 0.5 || origTop + curH > slideH + stepH * 0.5) {
              os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;
            } else {
              os.width = curW * 72; os.height = curH * 72;
            }
            await ctx.sync();
            totalFixes++;
          } catch (e) { /* ignore */ }
        }

        // Normalise similarly-sized shape groups
        const textShapes = pptxData.slideShapes.filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && ss.position && typeof ss.current.fontSize === "number" && !ss.isTable && !ss.isGroup);
        for (let i = 0; i < textShapes.length; i++) {
          const a = textShapes[i];
          const group = [a];
          for (let j = 0; j < textShapes.length; j++) {
            if (i === j) continue;
            const b = textShapes[j];
            if (Math.abs(a.position.width  - b.position.width)  / a.position.width  <= 0.15 &&
                Math.abs(a.position.height - b.position.height) / a.position.height <= 0.15) group.push(b);
          }
          if (group.length < 2) continue;
          const freq = group.reduce((acc, s) => { acc[s.current.fontSize] = (acc[s.current.fontSize]||0)+1; return acc; }, {});
          const groupSize = parseInt(Object.entries(freq).sort((a, b) => b[1]-a[1])[0][0]);
          for (const ss of group) {
            if (ss.current.fontSize === groupSize || Math.abs(ss.current.fontSize - groupSize) > 3) continue;
            const os = shapes.items.find(s => String(s.id) === String(ss.id));
            if (!os) continue;
            try { const tr = os.textFrame.textRange; tr.font.load("size"); await ctx.sync(); tr.font.size = groupSize; await ctx.sync(); totalFixes++; } catch (e) { /* ignore */ }
          }
        }
      });

      // ─── STEP 3: Colours — snap text to nearest theme colour ──────────────────
      addLog("Step 3: Colours…");
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const shapes = slides.items[dupIndex - 1].shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name", "type"]);
        await ctx.sync();

        // All shapes: snap font colour to nearest theme colour — iterate Office.js shapes directly
        const themeColorValues = Object.values(themeColors).filter(Boolean);
        const primaryThemeColor = themeColorValues[0];
        // All shapes: batch load colours via XML-matched IDs, then snap to nearest theme colour
        const colorJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (ss.isTable || ss.isGroup) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try { const tr = os.textFrame.textRange; tr.font.load("color"); colorJobs.push({ tr, ss }); } catch (e) { /* no text frame */ }
        }
        try { await ctx.sync(); } catch (e) { /* ignore */ }
        for (const { tr, ss } of colorJobs) {
          try {
            const rawColor = tr.font.color;
            const shapeColor = rawColor ? `#${rawColor}` : null;
            // If Office.js returns null, fall back to XML-parsed colour
            const xmlColor = ss.current?.color && ss.current.color !== "(inherited)" && !ss.current.color.startsWith("theme:") ? ss.current.color : null;
            const effectiveColor = (shapeColor && shapeColor !== "#null" && shapeColor !== "#") ? shapeColor : xmlColor;
            let targetColor;
            if (!effectiveColor) {
              targetColor = primaryThemeColor;
            } else {
              const nearest = snapToThemeColor(effectiveColor, themeColors);
              targetColor = nearest.toLowerCase() === effectiveColor.toLowerCase() ? null : nearest;
            }
            if (!targetColor) continue;
            tr.font.color = targetColor.replace("#", "");
            totalFixes++;
          } catch (e) { /* skip */ }
        }
        try { await ctx.sync(); } catch (e) { /* ignore */ }

        // All shapes: snap fill colour and border/line colour to nearest theme colour
        const fillJobs = [];
        for (const os of shapes.items) {
          try {
            os.load(["name", "type", "left", "top", "width", "height"]);
            os.fill.load(["type", "color"]);
            fillJobs.push({ os, kind: "fill" });
          } catch (e) { /* no fill */ }
          try { os.lineFormat.load(["color", "visible"]); fillJobs.push({ os, kind: "line" }); } catch (e) { /* no line */ }
        }
        try { await ctx.sync(); } catch (e) { /* ignore */ }
        // Identify which live shapes have text, to test whether a colourless shape sits behind one
        const textBearingBoxes = [];
        for (const os of shapes.items) {
          try {
            const tr = os.textFrame.textRange;
            tr.load("text");
            textBearingBoxes.push({ os, tr });
          } catch (e) { /* no text frame */ }
        }
        try { await ctx.sync(); } catch (e) { /* ignore */ }
        const textBoxRects = textBearingBoxes
          .filter(({ tr }) => { try { return tr.text && tr.text.trim().length > 0; } catch (e) { return false; } })
          .map(({ os }) => ({ left: os.left, top: os.top, width: os.width, height: os.height }));
        const isContainerOf = (container, rects) => rects.some(inner =>
          inner.left >= container.left - 0.5 && inner.left + inner.width  <= container.left + container.width  + 0.5 &&
          inner.top  >= container.top  - 0.5 && inner.top  + inner.height <= container.top  + container.height + 0.5
        );
        let colorRotationIndex = 0;
        for (const { os, kind } of fillJobs) {
          try {
            if (kind === "fill") {
              const fillTypeStr = String(os.fill.type).toLowerCase();
              if (fillTypeStr !== "solid") continue; // skip none/gradient/picture fills
              const liveFill = os.fill.color ? (os.fill.color.startsWith("#") ? os.fill.color : `#${os.fill.color}`) : null;
              if (liveFill) {
                const nearest = snapToThemeColor(liveFill, themeColors);
                if (nearest.toLowerCase() === liveFill.toLowerCase()) continue;
                os.fill.setSolidColor(nearest.replace("#", ""));
                totalFixes++;
              } else {
                // Colour comes from a style/theme reference Office.js can't read directly
                const shapeRect = { left: os.left, top: os.top, width: os.width, height: os.height };
                const behindTextBox = isContainerOf(shapeRect, textBoxRects);
                if (behindTextBox) {
                  // This shape contains a text box within its bounds — it's a background panel — clear the fill
                  os.fill.clear();
                  totalFixes++;
                } else if (themeColorValues.length > 0) {
                  // Standalone or overlapping decorative shape (e.g. icon) — assign a theme colour, rotating through the palette
                  const chosen = themeColorValues[colorRotationIndex % themeColorValues.length];
                  colorRotationIndex++;
                  os.fill.setSolidColor(chosen.replace("#", ""));
                  totalFixes++;
                }
              }
            } else {
              if (!os.lineFormat.visible) continue;
              const cur = os.lineFormat.color ? (os.lineFormat.color.startsWith("#") ? os.lineFormat.color : `#${os.lineFormat.color}`) : null;
              if (!cur) continue;
              const nearest = snapToThemeColor(cur, themeColors);
              if (nearest.toLowerCase() === cur.toLowerCase()) continue;
              os.lineFormat.color = nearest.replace("#", "");
              totalFixes++;
            }
          } catch (e) { addLog(`⚠ write failed for "${os.name}": ${e.message}`); }
        }
        try { await ctx.sync(); } catch (e) { addLog(`⚠ sync after fill write failed: ${e.message}`); }
        for (const ss of pptxData.slideShapes) {
          if (!ss.isTable) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          try {
            const table = os.table;
            table.load("rowCount,columnCount");
            await ctx.sync();
            const cellJobs = [];
            for (let r = 0; r < table.rowCount; r++)
              for (let c = 0; c < table.columnCount; c++) {
                const cell = table.getCell(r, c);
                const tr = cell.textFrame.textRange;
                tr.font.load("color");
                cellJobs.push({ tr });
              }
            await ctx.sync();
            for (const { tr } of cellJobs) {
              try {
                const cur = tr.font.color ? `#${tr.font.color}` : null;
                if (!cur || cur === "#null" || cur === "#") continue;
                if (themeColorList.some(c => c && cur.toLowerCase() === c.toLowerCase())) continue;
                const nearest = snapToThemeColor(cur, themeColors);
                if (nearest.toLowerCase() !== cur.toLowerCase()) { tr.font.color = nearest.replace("#", ""); totalFixes++; }
              } catch (e) { /* empty cell */ }
            }
            await ctx.sync();
          } catch (e) { /* no table */ }
        }

        // Groups: batch load all child colours, sync once, write, sync once
        for (const ss of pptxData.slideShapes) {
          if (!ss.isGroup) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          try {
            const groupShapes = os.shapes;
            groupShapes.load("items");
            await ctx.sync();
            const trs = [];
            for (const child of groupShapes.items) {
              try { const tr = child.textFrame.textRange; tr.font.load("color"); trs.push({ tr }); } catch (e) { /* no text */ }
            }
            await ctx.sync();
            for (const { tr } of trs) {
              try {
                const cur = tr.font.color ? `#${tr.font.color}` : null;
                if (!cur || cur === "#null" || cur === "#") continue;
                if (themeColorList.some(c => c && cur.toLowerCase() === c.toLowerCase())) continue;
                const nearest = snapToThemeColor(cur, themeColors);
                if (nearest.toLowerCase() !== cur.toLowerCase()) { tr.font.color = nearest.replace("#", ""); totalFixes++; }
              } catch (e) { /* no text */ }
            }
            await ctx.sync();
          } catch (e) { /* no group */ }
        }
      });

      // ─── STEP 4: Grid pipeline (looped until stable) ────────────────────────
      let skipAlignment = false;
      {
        const nonTitleShapes = pptxData.slideShapes.filter(s =>
          s.phType !== "title" && s.phType !== "ctrTitle" && s.phType !== "sldNum" && s.phType !== "ftr" && s.position
        );
        const SHAPE_COMPLEXITY_LIMIT = 25;
        const tooManyShapes = nonTitleShapes.length > SHAPE_COMPLEXITY_LIMIT;
        if (tooManyShapes) {
          addLog(`Slide has ${nonTitleShapes.length} shapes (limit ${SHAPE_COMPLEXITY_LIMIT}) — skipping position/size alignment, keeping fonts, colours and title position`);
        }

        // Check original (pre-fix) positions for any overlap — if found, skip alignment entirely
        let hasOverlap = false;
        for (let i = 0; i < nonTitleShapes.length && !hasOverlap; i++) {
          for (let j = i + 1; j < nonTitleShapes.length; j++) {
            const a = nonTitleShapes[i].position, b = nonTitleShapes[j].position;
            const overlapX = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left);
            const overlapY = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top);
            if (overlapX > 0 && overlapY > 0) { hasOverlap = true; break; }
          }
        }
        if (hasOverlap && !tooManyShapes) {
          addLog(`Slide has overlapping shapes — skipping position/size alignment, keeping fonts, colours and title position`);
        }
        skipAlignment = tooManyShapes || hasOverlap;

        if (nonTitleShapes.length > 0 && !tooManyShapes && !hasOverlap && targetTitlePos) {
          const GRID = 50, MAX_CELLS = 5;
          const areaLeft = targetTitlePos.left, areaRight = targetTitlePos.left + targetTitlePos.width;
          const areaTop  = targetTitlePos.top + targetTitlePos.height + slideH * 0.013, areaBottom = slideH * 0.987;
          const cellW = (areaRight - areaLeft) / GRID, cellH = (areaBottom - areaTop) / GRID;
          const q = v => Math.round(v * 10000) / 10000; // 4dp precision to avoid float drift
          const snapX = v => { const n = Math.round((v - areaLeft) / cellW); return q(areaLeft + n * cellW); };
          const snapY = v => { const n = Math.round((v - areaTop)  / cellH); return q(areaTop  + n * cellH); };
          const clamp = (snapped, orig, cell) => Math.abs(snapped - orig) <= MAX_CELLS * cell ? snapped : orig;

          const gridFixes = [];
          const gridFixMap = new Map();
          const positions = nonTitleShapes.map(s => ({ ...s.position }));

          const recordFix = (idx) => {
            const s = nonTitleShapes[idx], p = positions[idx];
            const key = String(s.id);
            const ex = gridFixMap.get(key);
            if (ex) { ex.position = { ...p }; }
            else {
              const fix = { shapeName: s.name, shapeId: s.id, _slideShape: s, shapeFill: s.shapeFill||null, position: { ...p } };
              gridFixMap.set(key, fix);
              gridFixes.push(fix);
            }
          };

          const sameH = (a, b) => Math.abs(a.height - b.height) <= cellH;
          const sameW = (a, b) => Math.abs(a.width  - b.width)  <= cellW;

          // ── Step 1: Snap every shape to nearest grid point (once, up front) ──
          for (let i = 0; i < nonTitleShapes.length; i++) {
            const orig = nonTitleShapes[i].position;
            const newLeft   = clamp(snapX(orig.left),              orig.left,              cellW);
            const newTop    = clamp(snapY(orig.top),               orig.top,               cellH);
            const newRight  = clamp(snapX(orig.left + orig.width), orig.left + orig.width, cellW);
            const newBottom = clamp(snapY(orig.top  + orig.height),orig.top  + orig.height,cellH);
            const newWidth  = Math.max(newRight - newLeft, cellW);
            const newHeight = Math.max(newBottom - newTop, cellH);
            if (Math.abs(newLeft-orig.left) > cellW*0.1 || Math.abs(newTop-orig.top) > cellH*0.1 ||
                Math.abs(newWidth-orig.width) > cellW*0.1 || Math.abs(newHeight-orig.height) > cellH*0.1) {
              positions[i] = { left: newLeft, top: newTop, width: newWidth, height: newHeight };
              recordFix(i);
            }
          }

          // ── Steps 2–5 loop until no more changes ─────────────────────────────
          for (let pass = 0; pass < 5; pass++) {
            let changed = false;

            // Step 2: Match dimensions — where width or height is within 1 cell,
            //         snap both to the larger value
            for (let i = 0; i < nonTitleShapes.length; i++) {
              for (let j = i + 1; j < nonTitleShapes.length; j++) {
                const a = positions[i], b = positions[j];
                if (Math.abs(a.width - b.width) <= cellW * 2 && Math.abs(a.width - b.width) > cellW * 0.1) {
                  const w = Math.max(a.width, b.width);
                  positions[i].width = w; positions[j].width = w;
                  recordFix(i); recordFix(j); changed = true;
                }
                if (Math.abs(a.height - b.height) <= cellH * 2 && Math.abs(a.height - b.height) > cellH * 0.1) {
                  const h = Math.max(a.height, b.height);
                  positions[i].height = h; positions[j].height = h;
                  recordFix(i); recordFix(j); changed = true;
                }
              }
            }

            // Step 3: Resolve text-box overlaps — move one cell at a time
            // Skip pairs where one shape is fully contained inside the other (intentional layering, e.g. icons)
            const isContained = (x, y) =>
              x.left >= y.left - cellW * 0.1 && x.left + x.width  <= y.left + y.width  + cellW * 0.1 &&
              x.top  >= y.top  - cellH * 0.1 && x.top  + x.height <= y.top  + y.height + cellH * 0.1;
            for (let i = 0; i < nonTitleShapes.length; i++) {
              for (let j = i + 1; j < nonTitleShapes.length; j++) {
                const a = positions[i], b = positions[j];
                if (isContained(a, b) || isContained(b, a)) continue;
                const overlapX = Math.min(a.left+a.width, b.left+b.width) - Math.max(a.left, b.left);
                const overlapY = Math.min(a.top+a.height, b.top+b.height) - Math.max(a.top,  b.top);
                if (overlapX > cellW * 0.1 && overlapY > cellH * 0.1) {
                  // Move j one cell in the axis of least overlap
                  if (overlapX <= overlapY) { positions[j].left += cellW; }
                  else                      { positions[j].top  += cellH; }
                  recordFix(j); changed = true;
                }
              }
            }

            // Step 4: Align edges — within 2 grid cells AND same height/width
            for (let i = 0; i < nonTitleShapes.length; i++) {
              for (let j = i + 1; j < nonTitleShapes.length; j++) {
                const a = positions[i], b = positions[j];
                // Align tops if within 2 vertical cells — no height check needed
                if (Math.abs(a.top - b.top) <= cellH * 2 && Math.abs(a.top - b.top) > cellH * 0.1) {
                  const t = q(Math.min(a.top, b.top));
                  positions[i].top = t; positions[j].top = t;
                  recordFix(i); recordFix(j); changed = true;
                }
                // Align lefts if within 2 horizontal cells and same width
                if (sameW(a, b) && Math.abs(a.left - b.left) <= cellW * 2 && Math.abs(a.left - b.left) > cellW * 0.1) {
                  const l = q(Math.min(a.left, b.left));
                  positions[i].left = l; positions[j].left = l;
                  recordFix(i); recordFix(j); changed = true;
                }
                // Align bottom edges if within 2 vertical cells and same height
                const aBotE = a.top + a.height, bBotE = b.top + b.height;
                if (sameH(a, b) && Math.abs(aBotE - bBotE) <= cellH * 2 && Math.abs(aBotE - bBotE) > cellH * 0.1) {
                  const bot = q(Math.min(aBotE, bBotE));
                  positions[i].top = q(bot - a.height); positions[j].top = q(bot - b.height);
                  recordFix(i); recordFix(j); changed = true;
                }
                // Align right edges if within 2 horizontal cells and same width
                const aRightE = a.left + a.width, bRightE = b.left + b.width;
                if (sameW(a, b) && Math.abs(aRightE - bRightE) <= cellW * 2 && Math.abs(aRightE - bRightE) > cellW * 0.1) {
                  const right = q(Math.min(aRightE, bRightE));
                  positions[i].left = q(right - a.width); positions[j].left = q(right - b.width);
                  recordFix(i); recordFix(j); changed = true;
                }
              }
            }

            // Step 5: Distribute spacing — groups of 2+ that are aligned and
            //         same size get evenly spaced
            for (let i = 0; i < nonTitleShapes.length; i++) {
              // Find all shapes with same height and tops aligned with i
              const hRow = [i];
              for (let j = 0; j < nonTitleShapes.length; j++) {
                if (j === i) continue;
                const a = positions[i], b = positions[j];
                if (sameH(a, b) && Math.abs(a.top - b.top) <= cellH * 0.1) hRow.push(j);
              }
              if (hRow.length >= 2) {
                const sorted = hRow.sort((x, y) => positions[x].left - positions[y].left);
                const lmost = positions[sorted[0]].left;
                const rmost = positions[sorted[sorted.length-1]].left + positions[sorted[sorted.length-1]].width;
                const totalW = sorted.reduce((s, idx) => s + positions[idx].width, 0);
                const gap = (rmost - lmost - totalW) / (sorted.length - 1);
                if (gap > 0) {
                  let cursor = lmost;
                  for (const idx of sorted) {
                    if (Math.abs(positions[idx].left - cursor) > cellW * 0.1) {
                      positions[idx].left = cursor; recordFix(idx); changed = true;
                    }
                    cursor += positions[idx].width + gap;
                  }
                }
              }

              // Find all shapes with same width and lefts aligned with i
              const wCol = [i];
              for (let j = 0; j < nonTitleShapes.length; j++) {
                if (j === i) continue;
                const a = positions[i], b = positions[j];
                if (sameW(a, b) && Math.abs(a.left - b.left) <= cellW * 0.1) wCol.push(j);
              }
              if (wCol.length >= 2) {
                const sorted = wCol.sort((x, y) => positions[x].top - positions[y].top);
                const tmost = positions[sorted[0]].top;
                const bmost = positions[sorted[sorted.length-1]].top + positions[sorted[sorted.length-1]].height;
                const totalH = sorted.reduce((s, idx) => s + positions[idx].height, 0);
                const gap = (bmost - tmost - totalH) / (sorted.length - 1);
                if (gap > 0) {
                  let cursor = tmost;
                  for (const idx of sorted) {
                    if (Math.abs(positions[idx].top - cursor) > cellH * 0.1) {
                      positions[idx].top = cursor; recordFix(idx); changed = true;
                    }
                    cursor += positions[idx].height + gap;
                  }
                }
              }
            }

            if (!changed) break;
          }

          if (gridFixes.length > 0) {
            addLog(`Step 4: Aligning ${gridFixes.length} shape(s)…`);
            await applyFixes(dupIndex, gridFixes, themeColors);
            totalFixes += gridFixes.length;
          }
        }
      }

      // ─── ENFORCE: title position always wins ────────────────────────────────
      if (titleShape && targetTitlePos) {
        await applyFixes(dupIndex, [{
          shapeName: titleShape.name, shapeId: titleShape.id, _slideShape: titleShape,
          shapeFill: titleShape.shapeFill || null, position: targetTitlePos,
        }], themeColors);
      }

      // ─── SAFETY NET: push any remaining text-box overlaps apart ─────────────
      if (!skipAlignment) {
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const shapes = slides.items[dupIndex - 1].shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name", "left", "top", "width", "height", "type"]);
        await ctx.sync();
        // Only consider text boxes (type 1 = text box, 14 = placeholder)
        const live = shapes.items
          .filter(s => s.width > 0 && s.height > 0)
          .map(s => ({ s, left: s.left/72, top: s.top/72, width: s.width/72, height: s.height/72 }));
        // Grid cell sizes based on slide area
        const gW = slideW / 50, gH = slideH / 50;
        let changed = true;
        const isContainedLive = (x, y) =>
          x.left >= y.left - gW * 0.5 && x.left + x.width  <= y.left + y.width  + gW * 0.5 &&
          x.top  >= y.top  - gH * 0.5 && x.top  + x.height <= y.top  + y.height + gH * 0.5;
        for (let iter = 0; iter < 8 && changed; iter++) {
          changed = false;
          for (let i = 0; i < live.length; i++) {
            for (let j = i + 1; j < live.length; j++) {
              const a = live[i], b = live[j];
              if (isContainedLive(a, b) || isContainedLive(b, a)) continue;
              const overX = a.left < b.left + b.width  - gW && a.left + a.width  > b.left + gW;
              const overY = a.top  < b.top  + b.height - gH && a.top  + a.height > b.top  + gH;
              if (!overX || !overY) continue;
              const lower = a.top >= b.top ? a : b;
              const upper = a.top >= b.top ? b : a;
              const newTop = upper.top + upper.height + gH;
              if (newTop + lower.height <= slideH) { lower.top = newTop; lower.s.top = newTop * 72; changed = true; }
            }
          }
        }
        await ctx.sync();
      });
      }

      addLog(`✓ Done — ${totalFixes} fix${totalFixes !== 1 ? "es" : ""} applied`);
      setFixCount(totalFixes);
      setStatus("done");
    } catch (err) {
      setError(err.message);
      addLog("✗ " + err.message);
      setStatus("error");
    }
  }, [titleOverrides]);

  const isRunning = status === "running";

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#f8f9fb", minHeight: "100vh", display: "flex", flexDirection: "column", maxWidth: 340, margin: "0 auto" }}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg) } }
        @keyframes fadeIn { from{opacity:0;transform:translateY(5px)} to{opacity:1;transform:translateY(0)} }
        .btn:hover:not(:disabled) { background: #174f8a !important; transform: translateY(-1px); box-shadow: 0 6px 20px rgba(31,92,158,0.4) !important; }
        .btn:disabled { opacity: 0.6; cursor: not-allowed; }
      `}</style>

      <div style={{ background: "#111111", padding: "18px 16px 14px", color: "#fff" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>✦</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>SlideLint</div>
            <div style={{ fontSize: 10, opacity: 0.75 }}>Fixes your slides</div>
          </div>
          <div style={{ marginLeft: "auto" }}>
            <span style={{ fontSize: 10, opacity: 0.7 }}>
              {status === "idle" && "Ready"}
              {status === "running" && "Working…"}
              {status === "done" && "✓ Done"}
              {status === "error" && "✗ Error"}
            </span>
          </div>
        </div>
      </div>

      <div style={{ flex: 1, padding: 16, display: "flex", flexDirection: "column", gap: 12 }}>

        {/* File loading status */}
        {!fileReady && !fileError && (
          <div style={{ background: "#fff", borderRadius: 8, border: "1px solid #e5e7eb", padding: "10px 14px", fontSize: 11, color: "#6b7280", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ width: 10, height: 10, border: "2px solid #d1d5db", borderTop: "2px solid #6b7280", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block", flexShrink: 0 }} />
            Loading template in background…
          </div>
        )}
        {fileReady && status === "idle" && (
          <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "8px 12px", fontSize: 11, color: "#166534", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span>✓ Template ready: <strong>{cachedMasters.current?.[0]?.name}</strong></span>
            <button onClick={async () => { setFileReady(false); setFileError(null); try { const { zip, masters } = await readPptxFile(); cachedZip.current = zip; cachedMasters.current = masters; setFileReady(true); } catch (e) { setFileError(e.message); } }}
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#15803d", textDecoration: "underline", padding: 0 }}>↺ Reload</button>
          </div>
        )}
        {fileError && (
          <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "8px 12px", fontSize: 11, color: "#991b1b", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span>⚠ {fileError}</span>
            <button onClick={async () => { setFileReady(false); setFileError(null); try { const { zip, masters } = await readPptxFile(); cachedZip.current = zip; cachedMasters.current = masters; setFileReady(true); } catch (e) { setFileError(e.message); } }}
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#991b1b", textDecoration: "underline", padding: 0 }}>↺ Retry</button>
          </div>
        )}

        {detectedTheme && <ThemeCard theme={detectedTheme} masterPlaceholders={detectedMaster} />}

        {status === "idle" && !detectedTheme && (
          <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>What cleanup does</div>
            {["Snaps title to master position & size", "Fixes fonts to match the template", "Normalises font sizes", "Resets colours to theme palette", "Aligns shapes to a clean grid"].map(text => (
              <div key={text} style={{ display: "flex", gap: 8, marginBottom: 5, alignItems: "flex-start" }}>
                <span style={{ fontSize: 11, color: "#111111", flexShrink: 0, width: 16, textAlign: "center" }}>✓</span>
                <span style={{ fontSize: 11, color: "#374151", lineHeight: 1.4 }}>{text}</span>
              </div>
            ))}
          </div>
        )}

        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={handleSetTitleOverride} disabled={isRunning}
            style={{ flex: 1, padding: "8px 0", background: "#fff", color: "#374151", border: "1px solid #e5e7eb", borderRadius: 8, fontSize: 11, fontWeight: 600, cursor: "pointer" }}>
            Use selected shape as title
          </button>
        </div>
        {overrideStatus && (
          <div style={{ background: overrideStatus.ok ? "#f0fdf4" : "#fef2f2", border: `1px solid ${overrideStatus.ok ? "#bbf7d0" : "#fecaca"}`, borderRadius: 8, padding: "8px 12px", fontSize: 11, color: overrideStatus.ok ? "#166534" : "#991b1b" }}>
            {overrideStatus.msg}
          </div>
        )}

        <button className="btn" onClick={handleCleanup} disabled={isRunning}
          style={{ width: "100%", padding: "14px 0", background: status === "done" ? "#15803d" : "#111111", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "all 0.2s ease", boxShadow: "0 4px 14px rgba(0,0,0,0.28)" }}>
          {isRunning ? (
            <><span style={{ width: 14, height: 14, border: "2px solid rgba(255,255,255,0.3)", borderTop: "2px solid #fff", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block" }} />Working…</>
          ) : status === "done" ? "✓ Done — clean another?" : "SlideLint"}
        </button>

        {status === "done" && fixCount > 0  && <FixBadge count={fixCount} />}
        {status === "done" && fixCount === 0 && <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#166534" }}>✓ Slide already matches the master — no changes needed.</div>}

        {error && <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "10px 12px", fontSize: 11, color: "#991b1b" }}><strong>Error:</strong> {error}</div>}

        {log.length > 0 && (
          <div style={{ background: "#0f172a", borderRadius: 10, padding: "10px 12px" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#64748b", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 6 }}>Activity</div>
            {log.map((entry, i) => <LogLine key={i} entry={entry} />)}
          </div>
        )}
      </div>

      <div style={{ padding: "10px 16px", borderTop: "1px solid #e5e7eb", background: "#fff", display: "flex", justifyContent: "space-between" }}>
        <span style={{ fontSize: 9, color: "#9ca3af", fontFamily: "monospace" }}>v3.0.0</span>
        <span style={{ fontSize: 9, color: "#9ca3af" }}>PowerPoint Add-in</span>
      </div>
    </div>
  );
}
