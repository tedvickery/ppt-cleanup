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

/* ── Colour / font extraction helpers ───────────────────────────────────── */

// Extract a resolved colour from a single rPr/defRPr element, or null if none
function extractColourFromRPr(rPr, theme) {
  if (!rPr) return null;
  const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
  if (!solidFill) return null;
  const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
  const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
  const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
  if (srgb)    return "#" + srgb.getAttribute("val");
  if (sys)     return "#" + (sys.getAttribute("lastClr") || "000000");
  if (scheme)  return resolveThemeColor(scheme.getAttribute("val"), theme.colors) || null;
  return null;
}

// Extract a font name from a single rPr/defRPr element, resolving +mj/+mn theme references
function extractFontFromRPr(rPr, theme) {
  if (!rPr) return null;
  const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
  const tf = latin?.getAttribute("typeface");
  if (!tf) return null;
  if (tf.startsWith("+mj")) return theme.fonts.heading;
  if (tf.startsWith("+mn")) return theme.fonts.body;
  return tf;
}

// Walk the full OOXML inheritance chain for colour: run → para defRPr → lstStyle lvl1pPr → fallback
function getEffectiveColour(runRPr, paraDefRPr, lvl1DefRPr, theme, isTitle) {
  return (
    extractColourFromRPr(runRPr,    theme) ||
    extractColourFromRPr(paraDefRPr, theme) ||
    extractColourFromRPr(lvl1DefRPr, theme) ||
    (isTitle ? (theme.colors.dark1 || "#000000") : (theme.colors.dark2 || theme.colors.dark1 || "#000000"))
  );
}

// Walk the full OOXML inheritance chain for font name
function getEffectiveFont(runRPr, paraDefRPr, lvl1DefRPr, theme, isTitle) {
  return (
    extractFontFromRPr(runRPr,    theme) ||
    extractFontFromRPr(paraDefRPr, theme) ||
    extractFontFromRPr(lvl1DefRPr, theme) ||
    (isTitle ? theme.fonts.heading : theme.fonts.body)
  );
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

    const txBody    = sp.getElementsByTagNameNS("*", "txBody")[0];
    const lstStyle  = txBody?.getElementsByTagNameNS("*", "lstStyle")[0];
    const lvl1pPr   = lstStyle?.getElementsByTagNameNS("*", "lvl1pPr")[0];
    const lvl1DefRPr = lvl1pPr?.getElementsByTagNameNS("*", "defRPr")[0];
    const firstPara  = txBody?.getElementsByTagNameNS("*", "p")[0];
    const pPr        = firstPara?.getElementsByTagNameNS("*", "pPr")[0];
    const paraDefRPr = pPr?.getElementsByTagNameNS("*", "defRPr")[0];
    const runRPr     = firstPara?.getElementsByTagNameNS("*", "r")[0]?.getElementsByTagNameNS("*", "rPr")[0];
    const isTitle    = phType === "title" || phType === "ctrTitle";

    // Font name — full inheritance walk: run → para defRPr → lstStyle lvl1pPr → theme
    const fontName = getEffectiveFont(runRPr, paraDefRPr, lvl1DefRPr, theme, isTitle);

    // Font size — check lvl1DefRPr first, then run rPr, then defaults
    let fontSize = null;
    const szEl = lvl1DefRPr || runRPr;
    if (szEl) { const sz = szEl.getAttribute("sz"); if (sz) fontSize = parseInt(sz, 10) / 100; }
    if (!fontSize) fontSize = isTitle ? 36 : 18;

    // Colour — full inheritance walk: run → para defRPr → lstStyle lvl1pPr → theme
    const color = getEffectiveColour(runRPr, paraDefRPr, lvl1DefRPr, theme, isTitle);

    // Bold
    let bold = null;
    const boldEl = runRPr || lvl1DefRPr;
    if (boldEl) { const b = boldEl.getAttribute("b"); bold = b === "1" || b === "true"; }
    if (bold === null) bold = isTitle;

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

    // Text body padding from bodyPr (in EMU — convert to points: 1pt = 12700 EMU)
    const bodyPr = txBody?.getElementsByTagNameNS("*", "bodyPr")[0];
    const DEFAULT_PADDING = { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 };
    const padding = {
      left:   bodyPr?.getAttribute("lIns") != null ? parseInt(bodyPr.getAttribute("lIns"), 10) / 12700 : DEFAULT_PADDING.left,
      right:  bodyPr?.getAttribute("rIns") != null ? parseInt(bodyPr.getAttribute("rIns"), 10) / 12700 : DEFAULT_PADDING.right,
      top:    bodyPr?.getAttribute("tIns") != null ? parseInt(bodyPr.getAttribute("tIns"), 10) / 12700 : DEFAULT_PADDING.top,
      bottom: bodyPr?.getAttribute("bIns") != null ? parseInt(bodyPr.getAttribute("bIns"), 10) / 12700 : DEFAULT_PADDING.bottom,
    };

    placeholders.push({ type: phType, idx: phIdx, font: { name: fontName, size: fontSize, color, bold }, alignment, position, fill: masterFill, paraFormat, padding });
  }

  // Fallback: if no placeholders found, try reading font from <p:txStyles> which some
  // modern masters use instead of placeholder shapes to define default text formatting
  if (placeholders.length === 0) {
    const txStyles = doc.getElementsByTagNameNS("*", "txStyles")[0];
    if (txStyles) {
      const styleMap = { titleStyle: "title", bodyStyle: "body", otherStyle: "body" };
      for (const [elName, phType] of Object.entries(styleMap)) {
        const styleEl = txStyles.getElementsByTagNameNS("*", elName)[0];
        if (!styleEl) continue;
        const lvl1pPr    = styleEl.getElementsByTagNameNS("*", "lvl1pPr")[0];
        const lvl1DefRPr = lvl1pPr?.getElementsByTagNameNS("*", "defRPr")[0];
        const fontName   = getEffectiveFont(null, null, lvl1DefRPr, theme, phType === "title");
        const color      = getEffectiveColour(null, null, lvl1DefRPr, theme, phType === "title");
        let fontSize = null;
        if (lvl1DefRPr) { const sz = lvl1DefRPr.getAttribute("sz"); if (sz) fontSize = parseInt(sz, 10) / 100; }
        if (!fontSize) fontSize = phType === "title" ? 36 : 18;
        if (!placeholders.find(p => p.type === phType)) {
          placeholders.push({ type: phType, idx: "0", font: { name: fontName, size: fontSize, color, bold: phType === "title" }, alignment: "left", position: null, fill: null, paraFormat: {} });
        }
      }
    }
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

    // Shape fill — check spPr first, then fall back to p:style fillRef
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
    // If spPr has no fill, check p:style > a:fillRef — this is how icon/preset shapes store their colour
    if (!shapeFill) {
      const style = sp.getElementsByTagNameNS("*", "style")[0];
      const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
      if (fillRef) {
        const scheme = fillRef.getElementsByTagNameNS("*", "schemeClr")[0];
        const srgb   = fillRef.getElementsByTagNameNS("*", "srgbClr")[0];
        const sys    = fillRef.getElementsByTagNameNS("*", "sysClr")[0];
        if (srgb)    shapeFill = "#" + srgb.getAttribute("val").toUpperCase();
        else if (sys) shapeFill = "#" + (sys.getAttribute("lastClr") || "000000").toUpperCase();
        else if (scheme) {
          const resolved = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
          shapeFill = resolved ? resolved.toUpperCase() : ("theme:" + scheme.getAttribute("val"));
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
    // Walk inheritance chain for font name if not found on run rPr
    if (!fontName) {
      for (const para of allParas) {
        const defRPr = para.getElementsByTagNameNS("*", "pPr")[0]?.getElementsByTagNameNS("*", "defRPr")[0];
        const tf = defRPr?.getElementsByTagNameNS("*", "latin")[0]?.getAttribute("typeface");
        if (tf) { fontName = tf.startsWith("+mj") ? theme.fonts.heading : tf.startsWith("+mn") ? theme.fonts.body : tf; break; }
      }
    }
    if (!fontName) {
      const lvl1DefRPr = txBody.getElementsByTagNameNS("*", "lstStyle")[0]?.getElementsByTagNameNS("*", "lvl1pPr")[0]?.getElementsByTagNameNS("*", "defRPr")[0];
      const tf = lvl1DefRPr?.getElementsByTagNameNS("*", "latin")[0]?.getAttribute("typeface");
      if (tf) fontName = tf.startsWith("+mj") ? theme.fonts.heading : tf.startsWith("+mn") ? theme.fonts.body : tf;
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
    const masterRefs = [];
    for (const rel of relsDoc.getElementsByTagNameNS("*", "Relationship")) {
      const target = rel.getAttribute("Target") || "";
      const match = target.match(/slideMasters\/slideMaster(\d+)\.xml/);
      if (match) masterRefs.push(parseInt(match[1], 10));
    }

    await Promise.all(masterRefs.map(async (masterIndex) => {
      const masterPath     = `ppt/slideMasters/slideMaster${masterIndex}.xml`;
      const masterRelsPath = `ppt/slideMasters/_rels/slideMaster${masterIndex}.xml.rels`;
      let theme = { colors: {}, fonts: { heading: null, body: null } };
      const [masterRelsFile, masterFile] = [zip.file(masterRelsPath), zip.file(masterPath)];
      if (!masterFile) return;

      const [masterRelsXml, masterXml] = await Promise.all([
        masterRelsFile ? masterRelsFile.async("string") : Promise.resolve(null),
        masterFile.async("string"),
      ]);

      if (masterRelsXml) {
        for (const mRel of parseXml(masterRelsXml).getElementsByTagNameNS("*", "Relationship")) {
          const themeMatch = (mRel.getAttribute("Target") || "").match(/\.\.\/theme\/theme(\d+)\.xml/);
          if (themeMatch) {
            const themeFile = zip.file(`ppt/theme/theme${themeMatch[1]}.xml`);
            if (themeFile) theme = parseThemeXml(await themeFile.async("string"));
            break;
          }
        }
      }

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
    }));

    masters.sort((a, b) => a.index - b.index);
  }

  // Count slides per master — parallelised with cached layout rels
  const masterSlideCounts = {};
  for (const m of masters) masterSlideCounts[m.index] = 0;
  const slideFiles = zip.file(/^ppt\/slides\/slide\d+\.xml$/);
  const layoutRelsCache = {};

  await Promise.all(slideFiles.map(async (slideFile) => {
    const slideNum = slideFile.name.match(/slide(\d+)\.xml/)?.[1];
    if (!slideNum) return;
    const relsFile = zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`);
    if (!relsFile) return;
    try {
      const relsXml = await relsFile.async("string");
      const relsDoc = parseXml(relsXml);
      for (const rel of relsDoc.getElementsByTagNameNS("*", "Relationship")) {
        const layoutMatch = (rel.getAttribute("Target") || "").match(/slideLayouts\/slideLayout(\d+)\.xml/);
        if (!layoutMatch) break;
        const layoutNum = layoutMatch[1];
        if (!layoutRelsCache[layoutNum]) {
          const layoutRelsFile = zip.file(`ppt/slideLayouts/_rels/slideLayout${layoutNum}.xml.rels`);
          layoutRelsCache[layoutNum] = layoutRelsFile ? await layoutRelsFile.async("string") : null;
        }
        const layoutRelsXml = layoutRelsCache[layoutNum];
        if (!layoutRelsXml) break;
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
  }));
  const maxCount = Math.max(...Object.values(masterSlideCounts), 0);
  const filtered = masters.filter(m => {
    const count = masterSlideCounts[m.index] || 0;
    return maxCount === 0 || count >= Math.max(1, maxCount * 0.2);
  });
  const dominantMasters = filtered.length > 0 ? filtered : masters;
  const dominantMasterIndex = dominantMasters[0]?.index ?? 1;
  return { masters: dominantMasters, dominantMasterIndex };
}

/* ── Phase 1: read file bytes + discover masters ────────────────────────── */

async function readPptxFile() {
  const bytes = await getFileBytes();
  const zip = await JSZip.loadAsync(bytes);
  const { masters, dominantMasterIndex } = await readAllMasters(zip);
  return { zip, masters, dominantMasterIndex };
}

/* ── Phase 2: read slide data using chosen master ───────────────────────── */

async function readSlideWithMaster(zip, masters, chosenMasterIndex, selectedSlideIndex) {
  const slideFile = zip.file(`ppt/slides/slide${selectedSlideIndex}.xml`);
  if (!slideFile) throw new Error(`Slide ${selectedSlideIndex} not found in file`);
  const slideXml = await slideFile.async("string");

  // Use the chosen master's layouts for title position
  let layoutPositions = {};
  const masterRelsFile = zip.file(`ppt/slideMasters/_rels/slideMaster${chosenMasterIndex}.xml.rels`);
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
      let layoutTitlePos = null, hasBody = false, isCtrTitle = false;

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
        if (phType === "title" || phType === "ctrTitle") { layoutTitlePos = pos; isCtrTitle = phType === "ctrTitle"; }
        if (phType === "body" || phType === "obj") hasBody = true;
      }
      firstLayoutDone = true;
      if (layoutTitlePos && hasBody && !isCtrTitle) titlePositions.push(layoutTitlePos);
    }

    // Simple mode vote — most common title position across all content layouts wins
    if (titlePositions.length > 0) {
      const freq = {};
      for (const p of titlePositions) {
        const k = `${p.left.toFixed(2)},${p.top.toFixed(2)},${p.width.toFixed(2)},${p.height.toFixed(2)}`;
        freq[k] = (freq[k] || 0) + 1;
      }
      const topKey = Object.entries(freq).sort((a, b) => b[1] - a[1])[0][0];
      const [left, top, width, height] = topKey.split(",").map(Number);
      layoutPositions["title:0"] = { left, top, width, height };
    }
  }

  const master = masters.find(m => m.index === chosenMasterIndex) || masters[0];
  const slideShapes = parseSlideXml(slideXml, master.theme, master.placeholders, layoutPositions);
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

    for (const fix of fixes) {
      const lookupId = fix._slideShape?.id || fix.shapeId;
      const target = lookupId
        ? shapes.items.find(s => String(s.id) === String(lookupId))
        : shapes.items.find(s => s.name === fix.shapeName);
      if (!target) continue;

      if (fix.position) {
        const inchToPt = 72;
        const { left, top, width, height } = fix.position;
        if (left   !== undefined && left   >= -2 && left   < 15)  target.left   = left   * inchToPt;
        if (top    !== undefined && top    >= -2 && top    < 10)  target.top    = top    * inchToPt;
        if (width  !== undefined && width  >= 0  && width  <= 15) target.width  = width  * inchToPt;
        if (height !== undefined && height >= 0  && height <= 10) target.height = height * inchToPt;
      }

      if (fix.fill !== undefined && fix.fill !== null) {
        try {
          if (fix.fill === "none" && fix.shapeFill && fix.shapeFill !== "none") target.fill.clear();
          else if (fix.fill?.startsWith("#")) { const snapped = snapToThemeColor(fix.fill, themeColors); target.fill.setSolidColor(snapped.replace("#", "")); }
        } catch (e) { /* skip */ }
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
        } catch (e) { /* skip */ }
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
            try {
              const groupShapes = target.shapes;
              groupShapes.load("items");
              await ctx.sync();
              for (const child of groupShapes.items) {
                try {
                  const tr = child.textFrame.textRange;
                  if (fix.font?.name)  tr.font.name = fix.font.name;
                  if (fix.font?.size)  tr.font.size = fix.font.size;
                  if (fix.font?.color) { tr.font.color = snapToThemeColor(fix.font.color, themeColors).replace("#", ""); }
                } catch (e) { /* no text */ }
              }
              await ctx.sync();
            } catch (e) { /* no group */ }
          } else {
            const tr = target.textFrame.textRange;
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
            if (fix.padding) {
              const tf = target.textFrame;
              if (fix.padding.left   != null) tf.leftMargin   = fix.padding.left;
              if (fix.padding.right  != null) tf.rightMargin  = fix.padding.right;
              if (fix.padding.top    != null) tf.topMargin    = fix.padding.top;
              if (fix.padding.bottom != null) tf.bottomMargin = fix.padding.bottom;
            }
            if (fix.verticalAlignment) {
              const vaMap = {
                top:    PowerPoint.TextVerticalAlignment.top,
                middle: PowerPoint.TextVerticalAlignment.middle,
                bottom: PowerPoint.TextVerticalAlignment.bottom,
              };
              if (vaMap[fix.verticalAlignment]) target.textFrame.verticalAlignment = vaMap[fix.verticalAlignment];
            }
            if (fix.alignment) {
              const alignMap = { left: PowerPoint.ParagraphHorizontalAlignment.left, center: PowerPoint.ParagraphHorizontalAlignment.center, right: PowerPoint.ParagraphHorizontalAlignment.right };
              if (alignMap[fix.alignment]) tr.paragraphFormat.horizontalAlignment = alignMap[fix.alignment];
            }
            await ctx.sync();
          }
        } catch (e) { /* skip */ }
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
  const colors = Object.entries(theme.colors).filter(([, v]) => v);
  return (
    <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>Detected from file</div>
      <div style={{ display: "flex", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
        {theme.fonts.heading && <span style={{ fontSize: 10, background: "#eff6ff", color: "#1e40af", padding: "2px 8px", borderRadius: 20, border: "1px solid #bfdbfe" }}>Aa {theme.fonts.heading}</span>}
        {theme.fonts.body && theme.fonts.body !== theme.fonts.heading && <span style={{ fontSize: 10, background: "#f5f3ff", color: "#6d28d9", padding: "2px 8px", borderRadius: 20, border: "1px solid #ddd6fe" }}>Aa {theme.fonts.body}</span>}
      </div>
      <div style={{ display: "flex", gap: 3, marginBottom: 8, flexWrap: "wrap" }}>
        {colors.map(([k, v]) => (
          <div key={k} title={`${k}: ${v}`}
            style={{ width: 18, height: 18, borderRadius: 4, background: v, border: "1px solid rgba(0,0,0,0.12)", flexShrink: 0 }} />
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
  const [masterWarning, setMasterWarning] = useState(false);
  const [colorWarning, setColorWarning]   = useState(null);

  // Manual title override, per slide, for this session only — { [slideIndex]: { id, name } }
  const [titleOverrides, setTitleOverrides] = useState({});
  const [overrideStatus, setOverrideStatus] = useState(null); // transient feedback message

  // Cached file data — loaded once in background, reused for every Fix click
  const [fileReady, setFileReady]   = useState(false);
  const [fileError, setFileError]   = useState(null);
  const cachedZip       = useRef(null);
  const cachedMasters   = useRef(null);
  const cachedDominantMaster = useRef(1);
  const cachedTemplateShapes = useRef(null); // pre-computed once from reference slides
  const cachedPptxData  = useRef({}); // keyed by slideIndex — slide shapes cached between runs

  // Load the file in the background as soon as the add-in opens
  useEffect(() => {
    let cancelled = false;
    async function loadFile() {
      try {
        const { zip, masters, dominantMasterIndex } = await readPptxFile();
        if (cancelled) return;
        cachedZip.current          = zip;
        cachedMasters.current      = masters;
        cachedDominantMaster.current = dominantMasterIndex;
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

  // ── Rebuild slide: fresh shell from layout + copied content shapes ────────
  const handleRebuildSlide = useCallback(async () => {
    setStatus("running");
    setLog([]);
    setError(null);
    addLog("Rebuilding slide from master…");
    try {
      const zip = cachedZip.current;
      if (!zip) throw new Error("No file loaded — run cleanup first to load the file");

      // Get current slide index
      const slideIndex = await getSelectedSlideIndex();
      const slidePath = `ppt/slides/slide${slideIndex}.xml`;
      const slideRelsPath = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;

      const slideFile = zip.file(slidePath);
      const slideRelsFile = zip.file(slideRelsPath);
      if (!slideFile || !slideRelsFile) throw new Error("Slide files not found in ZIP");

      const slideXml = await slideFile.async("string");
      const slideRelsXml = await slideRelsFile.async("string");

      // Find which layout this slide references
      const slideRelsDoc = parseXml(slideRelsXml);
      const rels = slideRelsDoc.getElementsByTagNameNS("*", "Relationship");
      let layoutRId = null, layoutPath = null;
      for (const rel of rels) {
        const type = rel.getAttribute("Type") || "";
        if (type.includes("slideLayout")) {
          layoutRId = rel.getAttribute("Id");
          const target = rel.getAttribute("Target") || "";
          layoutPath = target.startsWith("../") ? "ppt/" + target.slice(3) : "ppt/slides/" + target;
          break;
        }
      }
      if (!layoutPath) throw new Error("Could not find layout reference in slide rels");
      addLog(`Layout: ${layoutPath}`);

      // Extract content shapes from current slide (non-placeholder shapes only)
      const slideDoc = parseXml(slideXml);
      const spTree = slideDoc.getElementsByTagNameNS("*", "spTree")[0];
      if (!spTree) throw new Error("No spTree found in slide");

      // Collect content nodes: shapes without ph, pictures, graphic frames, groups
      const contentNodes = [];
      for (const child of Array.from(spTree.childNodes)) {
        const localName = child.localName;
        if (!localName) continue;
        if (localName === "sp") {
          // Only include if it has NO placeholder (ph) element
          const ph = child.getElementsByTagNameNS?.("*", "ph")?.[0];
          if (!ph) contentNodes.push(child);
        } else if (["pic", "graphicFrame", "grpSp", "sp"].includes(localName)) {
          contentNodes.push(child);
        }
      }
      addLog(`Found ${contentNodes.length} content shapes to carry over`);

      // Build fresh slide XML — minimal shell referencing same layout
      // The layout inheritance brings in footer, page number, date placeholders automatically
      const freshSlideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      ${contentNodes.map(n => {
        const s = new XMLSerializer();
        return s.serializeToString(n);
      }).join("\n      ")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;

      zip.file(slidePath, freshSlideXml);

      // Save back to PowerPoint
      const newPptxBytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
      await new Promise((resolve, reject) => {
        Office.context.document.setFileAsync(newPptxBytes.buffer, { fileType: Office.FileType.Compressed }, (result) => {
          result.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(result.error?.message || "setFileAsync failed"));
        });
      });

      // Invalidate cache for this slide
      if (cachedPptxData.current[slideIndex]) delete cachedPptxData.current[slideIndex];

      addLog(`✓ Slide rebuilt — ${contentNodes.length} shapes carried over`);
      setStatus("done");
    } catch (e) {
      addLog(`⚠ Rebuild failed: ${e.message}`);
      setStatus("idle");
    }
  }, [addLog]);

  const handleCleanup = useCallback(async () => {
    setStatus("running");
    setLog([]);
    setError(null);
    setFixCount(0);
    setMasterWarning(false);
    setColorWarning(null);
    setDetectedTheme(null);
    setDetectedMaster([]);

    try {
      addLog("Reading selected slide…");
      const slideIndex = await getSelectedSlideIndex();
      const slideW = 13.33, slideH = 7.5; // standard WIDE layout inches
      addLog(`Slide ${slideIndex} selected`);

      // Use cached zip and masters — re-read only if the cache is empty (first run or refresh)
      addLog("Reading .pptx file…");
      let zip = cachedZip.current;
      let masters = cachedMasters.current;
      if (!zip || !masters) {
        let dominantMasterIndex;
        ({ zip, masters, dominantMasterIndex } = await readPptxFile());
        cachedZip.current          = zip;
        cachedMasters.current      = masters;
        cachedDominantMaster.current = dominantMasterIndex;
      }

      if (masters.length === 0) throw new Error("No slide masters found in this file");
      const primaryMaster = masters[0];

      // Use cached pptxData for this slide — re-parse only if slide changed since last run
      let pptxData = cachedPptxData.current[slideIndex];
      if (!pptxData) {
        pptxData = await readSlideWithMaster(zip, masters, primaryMaster.index, slideIndex);
        cachedPptxData.current[slideIndex] = pptxData;
      } else {
        addLog(`Using cached template data for slide ${slideIndex}`);
      }
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

      // If no title shape is detected (by placeholder type or manual override), warn and stop —
      // the tool can't snap the title without knowing which shape it is
      const resolvedTitle = pptxData.slideShapes.find(s => s.phType === "title" || s.phType === "ctrTitle");
      if (!resolvedTitle) {
        setError("No title box found on this slide. Select the title shape and click \"Use selected shape as title\", then try again.");
        setStatus("idle");
        return;
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
          // Strip run-level latin font overrides — these take priority over paragraph-level
          // font writes from Office.js and prevent the master font from being applied correctly
          xml = xml.replace(/<a:latin[^>]*typeface="[^"]*"[^>]*\/>/g, "");
          xml = xml.replace(/<a:latin[^>]*typeface="[^"]*"[^>]*>[\s\S]*?<\/a:latin>/g, "");
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
          Math.abs(cur.left - targetTitlePos.left) > 0.05 ||
          Math.abs(cur.top  - targetTitlePos.top)  > 0.05;
        const headingFontForTitle = pptxData.theme.fonts.heading;
        const titleFontSize = titleMaster?.font?.size || null;
        const fontNeedsFix = headingFontForTitle &&
          titleShape.current.fontName !== "(inherited)" &&
          titleShape.current.fontName !== headingFontForTitle;
        const sizNeedsFix = titleFontSize && titleShape.current.fontSize &&
          Math.abs(titleShape.current.fontSize - titleFontSize) > 1;
        const fillNeedsFix = titleShape.shapeFill && titleShape.shapeFill !== "none" && titleShape.masterTarget?.fill === "none";
        const primaryThemeColorForTitle = themeColorList[0];
        const titleColorNeedsFix = primaryThemeColorForTitle &&
          titleShape.current.color !== "(inherited)" &&
          titleShape.current.color?.toLowerCase() !== primaryThemeColorForTitle.toLowerCase();
        if (posNeedsFix || fontNeedsFix || sizNeedsFix || fillNeedsFix || titleColorNeedsFix) {
          addLog("Step 1: Title position & font…");
          try {
            const posFix = posNeedsFix ? {
              left:   targetTitlePos.left,
              top:    targetTitlePos.top,
              width:  cur?.width  || targetTitlePos.width,
              height: cur?.height || targetTitlePos.height,
            } : undefined;
            await applyFixes(dupIndex, [{
              shapeName: titleShape.name, shapeId: titleShape.id, _slideShape: titleShape,
              shapeFill: fillNeedsFix ? "none" : (titleShape.shapeFill || null),
              ...(posNeedsFix ? { position: posFix } : {}),
              ...(fontNeedsFix || sizNeedsFix || titleColorNeedsFix ? { font: {
                ...(fontNeedsFix ? { name: headingFontForTitle } : {}),
                ...(sizNeedsFix ? { size: titleFontSize } : {}),
                ...(titleColorNeedsFix ? { color: primaryThemeColorForTitle } : {}),
              } } : {}),
              ...(titleMaster?.padding ? { padding: titleMaster.padding } : { padding: { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 } }),
              verticalAlignment: "top",
            }], themeColors);
            totalFixes++;

            // ── Squeeze all content shapes to fit new available space ──────────
            if (posNeedsFix && cur) {
              const SLIDE_W = 13.33, SLIDE_H = 7.5;
              const titleH = cur.height || 0.5;

              const oldTitleBottom = cur.top  + titleH;
              const newTitleBottom = targetTitlePos.top + titleH;
              const oldAvailV = SLIDE_H - oldTitleBottom;
              const newAvailV = SLIDE_H - newTitleBottom;
              const oldAvailH = SLIDE_W - cur.left;
              const newAvailH = SLIDE_W - targetTitlePos.left;

              const ratioV = (oldAvailV > 0 && newAvailV > 0) ? newAvailV / oldAvailV : 1;
              const ratioH = (oldAvailH > 0 && newAvailH > 0) ? newAvailH / oldAvailH : 1;
              const needsV = Math.abs(ratioV - 1) > 0.005;
              const needsH = Math.abs(ratioH - 1) > 0.005;

              if (needsV || needsH) {
                addLog(`  Squeezing content: V=${ratioV.toFixed(3)} H=${ratioH.toFixed(3)}`);
                await PowerPoint.run(async (ctx) => {
                  const slide = ctx.presentation.slides.getItemAt(dupIndex - 1);
                  const allShapes = slide.shapes;
                  allShapes.load("items");
                  await ctx.sync();
                  for (const s of allShapes.items) s.load(["id", "name", "left", "top", "width", "height"]);
                  await ctx.sync();

                  const inchToPt = 72;
                  const titleId = String(titleShape.id);
                  const ratioAbove = (cur.top > 0 && targetTitlePos.top > 0) ? targetTitlePos.top / cur.top : 1;
                  let squeezed = 0;
                  for (const s of allShapes.items) {
                    if (String(s.id) === titleId) continue;
                    const oldLeft   = s.left   / inchToPt;
                    const oldTop    = s.top     / inchToPt;
                    const oldWidth  = s.width   / inchToPt;
                    const oldHeight = s.height  / inchToPt;
                    const isAboveTitle = (oldTop + oldHeight) <= cur.top;
                    if (isAboveTitle && Math.abs(ratioAbove - 1) > 0.005) {
                      // Scale top proportionally within the space above the title
                      s.left = (targetTitlePos.left + (oldLeft - cur.left) * ratioH) * inchToPt;
                      s.top  = (oldTop * ratioAbove) * inchToPt;
                      s.width  = oldWidth  * ratioH * inchToPt;
                      s.height = oldHeight * inchToPt; // preserve height for shapes above
                    } else {
                      s.left   = (targetTitlePos.left + (oldLeft - cur.left) * ratioH) * inchToPt;
                      s.top    = (newTitleBottom      + (oldTop  - oldTitleBottom) * ratioV) * inchToPt;
                      s.width  = oldWidth  * ratioH * inchToPt;
                      s.height = oldHeight * ratioV * inchToPt;
                    }
                    squeezed++;
                  }
                  await ctx.sync();
                  addLog(`  ✓ Squeezed ${squeezed} shapes`);
                });
              }
            }
          } catch (e) { addLog(`⚠ Step 1 error: ${e.message}`); }
        }
      }

      // ─── STEPS 2 + 3: Fonts, sizes, colours — single PowerPoint.run ────────────
      addLog("Step 2: Fonts & colours…");
      const usedColors = new Set(); // track which theme colours are used — for ranking check
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

        // ── Font step: use theme fonts directly — most reliable source ──────────
        const nonTitleSizes = pptxData.slideShapes
          .filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && typeof ss.current.fontSize === "number")
          .map(ss => ss.current.fontSize);
        const sizeFreq = nonTitleSizes.reduce((acc, s) => { acc[s] = (acc[s] || 0) + 1; return acc; }, {});
        const normalisedSize = nonTitleSizes.length > 0 ? parseInt(Object.entries(sizeFreq).sort((a, b) => b[1] - a[1])[0][0]) : null;
        const bodyFont = pptxData.theme.fonts.body;   // minor font = body text
        const headingFont = pptxData.theme.fonts.heading; // major font = titles

        // Batch: load all regular shape text ranges — font name write is unconditional, size from XML
        const fontJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (ss.isTable || ss.isGroup || ss.phType === "title" || ss.phType === "ctrTitle") continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try { const tr = os.textFrame.textRange; fontJobs.push({ tr, ss }); } catch (e) { /* no text frame */ }
        }
        // Batch: load all table dimensions at once
        const tableJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (!ss.isTable) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try { const table = os.table; table.load("rowCount,columnCount"); tableJobs.push({ table, ss }); } catch (e) { /* no table */ }
        }
        // Batch: load all group children at once
        const groupJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (!ss.isGroup) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try { os.shapes.load("items"); groupJobs.push({ os, ss }); } catch (e) { /* no group */ }
        }
        await ctx.sync(); // ONE sync for all shape/table/group loads

        // Build table cell and group child jobs (no loads needed — font name write is unconditional)
        const tableCellJobs = [];
        for (const { table, ss } of tableJobs) {
          try {
            for (let r = 0; r < table.rowCount; r++)
              for (let c = 0; c < table.columnCount; c++)
                tableCellJobs.push({ tr: table.getCell(r, c).textFrame.textRange, ss });
          } catch (e) { /* empty table */ }
        }
        const groupTrJobs = [];
        for (const { os, ss } of groupJobs) {
          try {
            for (const child of os.shapes.items) {
              try { groupTrJobs.push({ tr: child.textFrame.textRange, ss }); } catch (e) { /* no text */ }
            }
          } catch (e) { /* no children */ }
        }

        // Write all font/size changes — no syncs needed
        for (const { tr, ss } of [...fontJobs, ...tableCellJobs, ...groupTrJobs]) {
          try {
            if (bodyFont) { tr.font.name = bodyFont; totalFixes++; }
            const currentSize = typeof ss.current.fontSize === "number" ? ss.current.fontSize : null;
            if (normalisedSize && currentSize !== null && Math.abs(currentSize - normalisedSize) <= 3) { tr.font.size = normalisedSize; totalFixes++; }
          } catch (e) { /* skip */ }
        }

        // Normalise similarly-sized shapes — O(n) frequency map approach
        const sizeGroups = new Map();
        for (const ss of pptxData.slideShapes.filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && ss.position && typeof ss.current.fontSize === "number" && !ss.isTable && !ss.isGroup)) {
          const sizeKey = `${Math.round(ss.position.width * 10)}_${Math.round(ss.position.height * 10)}`;
          if (!sizeGroups.has(sizeKey)) sizeGroups.set(sizeKey, []);
          sizeGroups.get(sizeKey).push(ss);
        }
        for (const group of sizeGroups.values()) {
          if (group.length < 2) continue;
          const freq = group.reduce((acc, s) => { acc[s.current.fontSize] = (acc[s.current.fontSize]||0)+1; return acc; }, {});
          const groupSize = parseInt(Object.entries(freq).sort((a, b) => b[1]-a[1])[0][0]);
          for (const ss of group) {
            if (ss.current.fontSize === groupSize || Math.abs(ss.current.fontSize - groupSize) > 3) continue;
            const os = shapes.items.find(s => String(s.id) === String(ss.id));
            if (!os) continue;
            try { os.textFrame.textRange.font.size = groupSize; totalFixes++; } catch (e) { /* ignore */ }
          }
        }
        await ctx.sync(); // ONE sync for all font/size writes

        // ── Colour step ──────────────────────────────────────────────────────────
        addLog("Step 3: Colours…");
        let bgColor = "#FFFFFF";

        // Read master background from XML — most reliable source
        try {
          const masterXmlFile = zip.file(`ppt/slideMasters/slideMaster${pptxData.theme ? cachedMasters.current?.[0]?.index ?? 1 : 1}.xml`);
          if (masterXmlFile) {
            const masterDoc = parseXml(await masterXmlFile.async("string"));
            const bg = masterDoc.getElementsByTagNameNS("*", "bg")[0];
            const solidFill = bg?.getElementsByTagNameNS("*", "solidFill")[0];
            if (solidFill) {
              const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
              const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
              const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
              if (srgb)    bgColor = "#" + srgb.getAttribute("val");
              else if (sys) bgColor = "#" + (sys.getAttribute("lastClr") || "FFFFFF");
              else if (scheme) bgColor = resolveThemeColor(scheme.getAttribute("val"), pptxData.theme.colors) || "#FFFFFF";
            }
          }
        } catch (e) { /* use default */ }

        try {
          const slideBg = slides.items[dupIndex - 1].background;
          slideBg.load("isMasterBackgroundFollowed");
          await ctx.sync();
          if (!slideBg.isMasterBackgroundFollowed) {
            slideBg.reset();
            await ctx.sync();
            addLog(`Background reset to master`);
          }
        } catch (e) {
          // followsMasterBackground not supported — snap to nearest theme colour
          try {
            const slideBgFill = slides.items[dupIndex - 1].background.fill;
            slideBgFill.load("type");
            await ctx.sync();
            if (String(slideBgFill.type).toLowerCase() === "solid") {
              const solid = slideBgFill.getSolidColorOrNullObject ? slideBgFill.getSolidColorOrNullObject() : null;
              if (solid) {
                solid.load("color");
                await ctx.sync();
                if (!solid.isNullObject && solid.color) {
                  const slideColor = solid.color.startsWith("#") ? solid.color : `#${solid.color}`;
                  const nearest = snapToThemeColor(slideColor, themeColors);
                  if (nearest.toLowerCase() !== slideColor.toLowerCase()) {
                    slideBgFill.setSolidColor(nearest.replace("#", ""));
                    await ctx.sync();
                    addLog(`Background snapped: ${slideColor} → ${nearest}`);
                  }
                }
              }
            }
          } catch (e2) { /* not supported */ }
        }

        // Build colour pools
        const themeColorsNoBg = themeColors;
        const themeColorValues = Object.values(themeColors).filter(Boolean);
        const fontPrimaryColor = themeColorValues[0];
        const fontThemeColors = themeColors;

        // Batch load all font colours and fill colours at once
        const colorJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (ss.isTable || ss.isGroup) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try { const tr = os.textFrame.textRange; tr.font.load("color"); colorJobs.push({ tr, ss }); } catch (e) { /* no text frame */ }
        }
        const fillJobs = [];
        for (const os of shapes.items) {
          try { os.fill.load(["type", "color", "foregroundColor"]); fillJobs.push({ os, kind: "fill" }); } catch (e) { /* no fill */ }
          try { os.lineFormat.load(["color", "visible"]); fillJobs.push({ os, kind: "line" }); } catch (e) { /* no line */ }
        }
        // Load table colour jobs
        const tableColorJobs = [];
        for (const ss of pptxData.slideShapes) {
          if (!ss.isTable) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          try { const table = os.table; table.load("rowCount,columnCount"); tableColorJobs.push({ table, ss }); } catch (e) { /* no table */ }
        }
        await ctx.sync(); // ONE sync for all colour reads (font + fill + table dims)

        // Load table cell colours now we know dimensions
        const tableCellColorJobs = [];
        for (const { table, ss } of tableColorJobs) {
          try {
            for (let r = 0; r < table.rowCount; r++)
              for (let c = 0; c < table.columnCount; c++) {
                const cell = table.getCell(r, c);
                const tr = cell.textFrame.textRange;
                tr.font.load("color");
                cell.fill.load(["type", "color", "foregroundColor"]);
                tableCellColorJobs.push({ tr, cell });
              }
          } catch (e) { /* empty table */ }
        }
        // Load group child colours (groupJobs already loaded children above)
        const groupColorJobs = [];
        for (const { os } of groupJobs) {
          try {
            for (const child of os.shapes.items) {
              try { const tr = child.textFrame.textRange; tr.font.load("color"); groupColorJobs.push({ tr }); } catch (e) { /* no text */ }
            }
          } catch (e) { /* no children */ }
        }
        await ctx.sync(); // ONE sync for table cell / group child colour reads

        // Write all font colours
        for (const { tr, ss } of colorJobs) {
          try {
            const rawColor = tr.font.color;
            const officeColor = (rawColor && rawColor !== "null") ? (rawColor.startsWith("#") ? rawColor : `#${rawColor}`) : null;
            const xmlExplicit = ss.current?.color && ss.current.color !== "(inherited)" && !ss.current.color.startsWith("theme:") ? ss.current.color : null;
            const masterColor = ss.masterTarget?.color && ss.masterTarget.color !== "(inherited)" ? ss.masterTarget.color : null;
            const effectiveColor = officeColor || xmlExplicit || masterColor;
            let targetColor;
            if (!effectiveColor) { targetColor = fontPrimaryColor; }
            else { const nearest = snapToThemeColor(effectiveColor, fontThemeColors); targetColor = nearest.toLowerCase() === effectiveColor.toLowerCase() ? null : nearest; }
            if (!targetColor) continue;
            tr.font.color = targetColor.replace("#", ""); totalFixes++;
            usedColors.add(targetColor.toUpperCase());
          } catch (e) { /* skip */ }
        }

        // Write all fill/border colours = themeColorValues.length > 0 ? themeColorValues : Object.values(themeColors).filter(Boolean);
        for (const { os, kind } of fillJobs) {
          try {
            if (kind === "fill") {
              const fillType = String(os.fill.type).toLowerCase();
              if (fillType !== "solid") { if (fillType !== "noFill" && fillType !== "null" && fillType !== "undefined") addLog(`  Skip fill: type=${fillType} name=${os.name}`); continue; }
              const liveColor = os.fill.color ? (os.fill.color.startsWith("#") ? os.fill.color : `#${os.fill.color}`) : null;
              const liveFg    = os.fill.foregroundColor ? (os.fill.foregroundColor.startsWith("#") ? os.fill.foregroundColor : `#${os.fill.foregroundColor}`) : null;
              const ss = pptxData.slideShapes.find(s => String(s.id) === String(os.id) || s.name === os.name);
              const xmlFill = ss?.shapeFill && !ss.shapeFill.startsWith("theme:") && ss.shapeFill !== "none" && !ss.shapeFill.includes("gradient") ? ss.shapeFill : null;
              const effectiveFill = liveColor || liveFg || xmlFill;
              if (effectiveFill) {
                // We know the current colour — snap to nearest theme colour
                if (fillPool.some(c => c.toLowerCase() === effectiveFill.toLowerCase())) continue; // already correct
                const nearest = snapToThemeColor(effectiveFill, themeColorsNoBg);
                os.fill.setSolidColor(nearest.replace("#", "")); totalFixes++;
                usedColors.add(nearest.toUpperCase());
              } else {
                // Colour genuinely unresolvable — random theme colour
                if (fillPool.length > 0) { os.fill.setSolidColor(fillPool[Math.floor(Math.random() * fillPool.length)].replace("#", "")); totalFixes++; }
              }
            } else {
              if (!os.lineFormat.visible) continue;
              const cur = os.lineFormat.color ? (os.lineFormat.color.startsWith("#") ? os.lineFormat.color : `#${os.lineFormat.color}`) : null;
              if (!cur) continue;
              const nearest = snapToThemeColor(cur, themeColorsNoBg);
              if (nearest.toLowerCase() !== cur.toLowerCase()) { os.lineFormat.color = nearest.replace("#", ""); totalFixes++; }
            }
          } catch (e) { /* skip */ }
        }

        // Write table and group cell/child colours
        for (const { tr } of [...tableCellColorJobs, ...groupColorJobs]) {
          try {
            const cur = tr.font.color ? `#${tr.font.color}` : null;
            if (!cur) continue;
            const nearest = snapToThemeColor(cur, themeColorsNoBg);
            if (nearest.toLowerCase() !== cur.toLowerCase()) { tr.font.color = nearest.replace("#", ""); totalFixes++; }
          } catch (e) { /* skip */ }
        }

        await ctx.sync(); // final write sync
      });




      addLog(`✓ Done — ${totalFixes} fix${totalFixes !== 1 ? "es" : ""} applied`);

      // ── Colour ranking check ─────────────────────────────────────────────
      // Warn if a lower-ranked theme colour is used but a higher-ranked one isn't
      const rankedColors = Object.values(themeColors).filter(Boolean).map(c => c.toUpperCase());
      const usedRanks = rankedColors.map((c, i) => usedColors.has(c) ? i : -1).filter(i => i >= 0);
      if (usedRanks.length > 0) {
        const highestUsed = Math.max(...usedRanks);
        const unusedAbove = rankedColors.slice(0, highestUsed).filter(c => !usedColors.has(c));
        if (unusedAbove.length > 0) {
          setColorWarning(`Colour ${highestUsed + 1} is used but ${unusedAbove.length} higher-priority colour${unusedAbove.length > 1 ? "s" : ""} ${unusedAbove.length > 1 ? "are" : "is"} unused — consider reviewing theme colour usage`);
        }
      }
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
            <div style={{ fontWeight: 700, fontSize: 15 }}>SnapBack</div>
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

        {/* 1. SnapBack button */}
        <button className="btn" onClick={handleCleanup} disabled={isRunning}
          style={{ width: "100%", padding: "14px 0", background: status === "done" ? "#15803d" : "#111111", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "all 0.2s ease", boxShadow: "0 4px 14px rgba(0,0,0,0.28)" }}>
          {isRunning ? (
            <><span style={{ width: 14, height: 14, border: "2px solid rgba(255,255,255,0.3)", borderTop: "2px solid #fff", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block" }} />Working…</>
          ) : status === "done" ? "✓ Done — clean another?" : "SnapBack"}
        </button>

        {status === "done" && fixCount > 0  && <FixBadge count={fixCount} />}
        {status === "done" && fixCount === 0 && <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#166534" }}>✓ Slide already matches the master — no changes needed.</div>}
        {status === "done" && masterWarning && (
          <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#92400e" }}>
            ⚠ Looks like this slide might not be aligned to the master template. Consider copying the contents of this slide to a new blank slide.
          </div>
        )}
        {status === "done" && colorWarning && (
          <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#92400e" }}>
            ⚠ {colorWarning}
          </div>
        )}
        {error && <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "10px 12px", fontSize: 11, color: "#991b1b" }}><strong>Error:</strong> {error}</div>}

        {/* 2. Select shape as title */}
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

        {/* 3. Template ready / file status */}
        {!fileReady && !fileError && (
          <div style={{ background: "#fff", borderRadius: 8, border: "1px solid #e5e7eb", padding: "10px 14px", fontSize: 11, color: "#6b7280", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ width: 10, height: 10, border: "2px solid #d1d5db", borderTop: "2px solid #6b7280", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block", flexShrink: 0 }} />
            Loading template in background…
          </div>
        )}
        {fileReady && status === "idle" && (
          <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "8px 12px", fontSize: 11, color: "#166534", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span>✓ Template ready: <strong>{cachedMasters.current?.[0]?.name}</strong></span>
            <button onClick={async () => { setFileReady(false); setFileError(null); try { const { zip, masters, dominantMasterIndex } = await readPptxFile(); cachedZip.current = zip; cachedMasters.current = masters; cachedDominantMaster.current = dominantMasterIndex; cachedPptxData.current = {}; cachedTemplateShapes.current = null; setFileReady(true); } catch (e) { setFileError(e.message); } }}
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#15803d", textDecoration: "underline", padding: 0 }}>↺ Reload</button>
          </div>
        )}
        {fileError && (
          <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "8px 12px", fontSize: 11, color: "#991b1b", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <span>⚠ {fileError}</span>
            <button onClick={async () => { setFileReady(false); setFileError(null); try { const { zip, masters, dominantMasterIndex } = await readPptxFile(); cachedZip.current = zip; cachedMasters.current = masters; cachedDominantMaster.current = dominantMasterIndex; cachedPptxData.current = {}; cachedTemplateShapes.current = null; setFileReady(true); } catch (e) { setFileError(e.message); } }}
              style={{ background: "none", border: "none", cursor: "pointer", fontSize: 11, color: "#991b1b", textDecoration: "underline", padding: 0 }}>↺ Retry</button>
          </div>
        )}

        {detectedTheme && <ThemeCard theme={detectedTheme} masterPlaceholders={detectedMaster} />}

        {/* 4. Description */}
        {status === "idle" && !detectedTheme && (
          <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>What cleanup does</div>
            {["Snaps title to master position & size", "Fixes fonts to match the template", "Normalises font sizes", "Snaps colours to theme palette", "Corrects shape fill colours"].map(text => (
              <div key={text} style={{ display: "flex", gap: 8, marginBottom: 5, alignItems: "flex-start" }}>
                <span style={{ fontSize: 11, color: "#111111", flexShrink: 0, width: 16, textAlign: "center" }}>✓</span>
                <span style={{ fontSize: 11, color: "#374151", lineHeight: 1.4 }}>{text}</span>
              </div>
            ))}
          </div>
        )}

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
