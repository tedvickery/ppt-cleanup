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
    const runs = txBody.getElementsByTagNameNS("*", "r");
    const textContent = Array.from(runs).map(r => r.getElementsByTagNameNS("*", "t")[0]?.textContent || "").join("");
    if (!textContent.trim() && runs.length === 0) continue;

    const allRuns  = Array.from(txBody.getElementsByTagNameNS("*", "r"));
    const firstRun = allRuns[0];
    const rPr     = firstRun?.getElementsByTagNameNS("*", "rPr")[0];
    const allParas = Array.from(txBody.getElementsByTagNameNS("*", "p"));

    let fontName = null, fontSize = null, color = null, bold = null, italic = null;

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
                     (phType === "body" ? masterPlaceholders.find(p => p.type === "body") : null);
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
  return shapes;
}

/* ── Read all masters ───────────────────────────────────────────────────── */

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

async function duplicateSlide(slideIndex) {
  // Work on original — user can Ctrl+Z to undo
  return slideIndex;
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

function snapToThemeColor(hex, themeColors, threshold = 80) {
  if (!hex || hex === "none" || hex.startsWith("theme:")) return hex;
  let nearest = null, nearestDist = Infinity;
  for (const [, themeHex] of Object.entries(themeColors)) {
    if (!themeHex) continue;
    const dist = colourDistance(hex, themeHex);
    if (dist < nearestDist) { nearestDist = dist; nearest = themeHex; }
  }
  if (nearestDist <= threshold) { console.log(`  Snapping ${hex} → ${nearest}`); return nearest; }
  return hex;
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
      addLog(`Slide ${slideIndex} selected`);

      // Use cached file data if available, otherwise read now
      let zip, masters;
      if (cachedZip.current && cachedMasters.current) {
        zip     = cachedZip.current;
        masters = cachedMasters.current;
        addLog(`Using cached template: "${masters[0]?.name}"`);
      } else {
        addLog("Reading .pptx file…");
        ({ zip, masters } = await readPptxFile());
        cachedZip.current     = zip;
        cachedMasters.current = masters;
        addLog(`Found ${masters.length} master${masters.length !== 1 ? "s" : ""}`);
        if (masters.length > 1) addLog(`(${masters.length - 1} imported master${masters.length > 2 ? "s" : ""} ignored)`);
      }

      if (masters.length === 0) throw new Error("No slide masters found in this file");
      const primaryMaster = masters[0];
      const pptxData = await readSlideWithMaster(zip, masters, primaryMaster.index, slideIndex);
      setDetectedTheme(pptxData.theme);
      setDetectedMaster(pptxData.masterPlaceholders);

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
          Math.abs(cur.left   - targetTitlePos.left)   > 0.01 ||
          Math.abs(cur.top    - targetTitlePos.top)    > 0.01 ||
          Math.abs(cur.width  - targetTitlePos.width)  > 0.01 ||
          Math.abs(cur.height - targetTitlePos.height) > 0.01;
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

      // ─── STEP 2: Fonts — correct font name, normalise sizes, expand text boxes
      addLog("Step 2: Fonts…");
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const slide  = slides.items[dupIndex - 1];
        const shapes = slide.shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name"]);
        await ctx.sync();

        const nonTitleSizes = pptxData.slideShapes
          .filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && typeof ss.current.fontSize === "number")
          .map(ss => ss.current.fontSize);
        const sizeFreq = nonTitleSizes.reduce((acc, s) => { acc[s] = (acc[s] || 0) + 1; return acc; }, {});
        const normalisedSize = nonTitleSizes.length > 0 ? parseInt(Object.entries(sizeFreq).sort((a, b) => b[1] - a[1])[0][0]) : null;

        for (const ss of pptxData.slideShapes) {
          if (!ss.masterTarget) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id)) || shapes.items.find(s => s.name === ss.name);
          if (!os) continue;
          try {
            const tr = os.textFrame.textRange;
            tr.font.load(["name", "size"]);
            await ctx.sync();
            const isTitle       = ss.phType === "title" || ss.phType === "ctrTitle";
            const wrongFont     = ss.current.fontName !== "(inherited)" && ss.current.fontName !== ss.masterTarget.fontName;
            const inheritedFont = ss.current.fontName === "(inherited)";
            const mixedSize     = tr.font.size === null;
            const currentSize   = typeof ss.current.fontSize === "number" ? ss.current.fontSize : null;
            const sizesDiffer   = !isTitle && normalisedSize && currentSize !== null && currentSize !== normalisedSize && Math.abs(currentSize - normalisedSize) <= 3;
            const needsFontFix  = wrongFont || (inheritedFont && ss.masterTarget.fontName);
            const needsSizeFix  = mixedSize || sizesDiffer;
            const needsFillReset = ss.shapeFill && ss.shapeFill !== "none" && ss.masterTarget?.fill === "none";
            if (!needsFontFix && !needsSizeFix && !needsFillReset) continue;
            if (needsFillReset) { try { os.fill.clear(); await ctx.sync(); } catch (e) { /* ignore */ } }
            if (needsFontFix) {
              tr.font.name = ss.masterTarget.fontName;
              if (normalisedSize && currentSize !== null && Math.abs(currentSize - normalisedSize) <= 3) tr.font.size = normalisedSize;
            }
            if (!needsFontFix && needsSizeFix) {
              tr.font.size = isTitle ? ss.masterTarget.fontSize : (normalisedSize || ss.masterTarget.fontSize);
            }
            if (mixedSize && !needsFontFix) tr.font.size = normalisedSize || ss.masterTarget.fontSize;
            await ctx.sync();
            totalFixes++;
          } catch (e) { /* shape may not support font ops */ }
        }

        // Expand text boxes to fit content
        const slideW = 13.33, slideH = 7.5;
        for (const ss of pptxData.slideShapes) {
          if (!ss.position || !ss.textContent) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          try {
            os.load(["left", "top", "width", "height"]);
            os.textFrame.load("autoSizeSetting");
            await ctx.sync();
            os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeShapeToFitText;
            await ctx.sync();
            os.load(["width", "height"]);
            await ctx.sync();
            const neededW = os.width / 72, neededH = os.height / 72;
            os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
            await ctx.sync();
            const { left: origLeft, top: origTop, width: origW, height: origH } = ss.position;
            if (neededW <= origW + 0.01 && neededH <= origH + 0.01) continue;
            const stepW = slideW / 100, stepH = slideH / 100;
            const maxW = slideW - origLeft, maxH = slideH - origTop;
            let curW = origW, curH = origH;
            while ((curW < neededW - 0.01 || curH < neededH - 0.01) && (curW < maxW - 0.01 || curH < maxH - 0.01)) {
              if (curW < neededW - 0.01 && curW < maxW) curW = Math.min(curW + stepW, maxW);
              if (curH < neededH - 0.01 && curH < maxH) curH = Math.min(curH + stepH, maxH);
            }
            let overlaps = false;
            for (const other of pptxData.slideShapes) {
              if (String(other.id) === String(ss.id) || !other.position) continue;
              const o = other.position;
              if (origLeft < o.left + o.width && origLeft + curW > o.left && origTop < o.top + o.height && origTop + curH > o.top) { overlaps = true; break; }
            }
            if (overlaps || origLeft + curW > slideW + 0.05 || origTop + curH > slideH + 0.05) {
              os.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeTextToFitShape;
            } else {
              os.width = curW * 72; os.height = curH * 72;
            }
            await ctx.sync();
            totalFixes++;
          } catch (e) { /* ignore */ }
        }

        // Normalise similarly-sized shape groups
        const textShapes = pptxData.slideShapes.filter(ss => ss.phType !== "title" && ss.phType !== "ctrTitle" && ss.position && typeof ss.current.fontSize === "number");
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

      // ─── STEP 3: Colours — snap text to first theme colour ─────────────────
      addLog("Step 3: Colours…");
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const shapes = slides.items[dupIndex - 1].shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name"]);
        await ctx.sync();
        for (const ss of pptxData.slideShapes) {
          if (!ss.masterTarget) continue;
          const os = shapes.items.find(s => String(s.id) === String(ss.id));
          if (!os) continue;
          try {
            const tr = os.textFrame.textRange;
            tr.font.load("color");
            await ctx.sync();
            const liveColor = tr.font.color ? `#${tr.font.color}` : null;
            const cur = liveColor || ss.current.color;
            if (!cur || cur === "(inherited)" || cur === "#null" || cur === "#") continue;
            const isThemeColor = themeColorList.some(c => c && cur && c.toLowerCase() === cur.toLowerCase());
            if (isThemeColor) continue;
            if (ss.masterTarget?.color && ss.masterTarget.color !== "(inherited)" && cur.toLowerCase() === ss.masterTarget.color.toLowerCase()) continue;
            const firstThemeColor = themeColorList[0];
            if (!firstThemeColor || cur.toLowerCase() === firstThemeColor.toLowerCase()) continue;
            let textColor = firstThemeColor;
            const fill = ss.shapeFill;
            if (fill && fill !== "none" && !fill.startsWith("theme:")) {
              const contrast = (Math.max(hexLuminance(fill), hexLuminance(firstThemeColor)) + 0.05) / (Math.min(hexLuminance(fill), hexLuminance(firstThemeColor)) + 0.05);
              if (contrast < 3) textColor = hexLuminance(fill) > 0.179 ? "#000000" : "#FFFFFF";
            }
            tr.font.color = textColor.replace("#", "");
            await ctx.sync();
            totalFixes++;
          } catch (e) { /* no text */ }
        }
      });

      // ─── STEP 4: Grid alignment + overlap fix ───────────────────────────────
      {
        const nonTitleShapes = pptxData.slideShapes.filter(s =>
          s.phType !== "title" && s.phType !== "ctrTitle" && s.phType !== "sldNum" && s.phType !== "ftr" && s.position
        );
        if (nonTitleShapes.length > 0 && targetTitlePos) {
          const GRID = 100, MAX_CELLS = 5;
          const areaLeft = targetTitlePos.left, areaRight = targetTitlePos.left + targetTitlePos.width;
          const areaTop  = targetTitlePos.top + targetTitlePos.height + 0.1, areaBottom = 7.4;
          const cellW = (areaRight - areaLeft) / GRID, cellH = (areaBottom - areaTop) / GRID;
          const snapX = v => areaLeft + Math.round((v - areaLeft) / cellW) * cellW;
          const snapY = v => areaTop  + Math.round((v - areaTop)  / cellH) * cellH;
          const clamp = (snapped, orig, cell) => Math.abs(snapped - orig) <= MAX_CELLS * cell ? snapped : orig;

          const gridFixes = [];
          const positions = nonTitleShapes.map(s => ({ ...s.position }));

          for (let i = 0; i < nonTitleShapes.length; i++) {
            const s = nonTitleShapes[i], orig = s.position;
            const newLeft   = clamp(snapX(orig.left),              orig.left,              cellW);
            const newTop    = clamp(snapY(orig.top),               orig.top,               cellH);
            const newRight  = clamp(snapX(orig.left + orig.width), orig.left + orig.width, cellW);
            const newBottom = clamp(snapY(orig.top  + orig.height),orig.top  + orig.height,cellH);
            const newWidth  = Math.max(newRight - newLeft, cellW);
            const newHeight = Math.max(newBottom - newTop, cellH);
            if (Math.abs(newLeft-orig.left)>0.001 || Math.abs(newTop-orig.top)>0.001 || Math.abs(newWidth-orig.width)>0.001 || Math.abs(newHeight-orig.height)>0.001) {
              positions[i] = { left: newLeft, top: newTop, width: newWidth, height: newHeight };
              gridFixes.push({ shapeName: s.name, shapeId: s.id, _slideShape: s, shapeFill: s.shapeFill||null, position: { left: newLeft, top: newTop, width: newWidth, height: newHeight } });
            }
          }

          // Align similarly-sized shape edges
          for (let i = 0; i < nonTitleShapes.length; i++) {
            for (let j = i + 1; j < nonTitleShapes.length; j++) {
              const a = positions[i], b = positions[j];
              const wSim = Math.abs(a.width  - b.width)  / a.width  <= 0.15;
              const hSim = Math.abs(a.height - b.height) / a.height <= 0.15;
              if (!wSim && !hSim) continue;
              if (wSim && Math.abs(a.left - b.left) / a.width <= 0.15) { const avg = (a.left + b.left) / 2; positions[i].left = avg; positions[j].left = avg; }
              if (hSim && Math.abs(a.top  - b.top)  / a.height <= 0.15) { const avg = (a.top  + b.top)  / 2; positions[i].top  = avg; positions[j].top  = avg; }
              if (wSim && Math.abs((a.left+a.width) - (b.left+b.width)) / a.width <= 0.15) {
                const avg = ((a.left+a.width) + (b.left+b.width)) / 2;
                positions[i].left = avg - a.width; positions[j].left = avg - b.width;
              }
              if (hSim && Math.abs((a.top+a.height) - (b.top+b.height)) / a.height <= 0.15) {
                const avg = ((a.top+a.height) + (b.top+b.height)) / 2;
                positions[i].top = avg - a.height; positions[j].top = avg - b.height;
              }
            }
          }
          for (let i = 0; i < nonTitleShapes.length; i++) {
            const s = nonTitleShapes[i], orig = s.position, p = positions[i];
            if (Math.abs(p.left-orig.left) > 0.001 || Math.abs(p.top-orig.top) > 0.001) {
              const ex = gridFixes.find(f => String(f.shapeId) === String(s.id));
              if (ex) { ex.position.left = p.left; ex.position.top = p.top; }
              else gridFixes.push({ shapeName: s.name, shapeId: s.id, _slideShape: s, shapeFill: s.shapeFill||null, position: { ...p } });
            }
          }

          // Proximity snap: if any two shapes' edges are within 1 grid cell of each other, snap them
          // This catches column headers/boxes that are nearly aligned but not within dimension tolerance
          for (let i = 0; i < nonTitleShapes.length; i++) {
            for (let j = i + 1; j < nonTitleShapes.length; j++) {
              const a = positions[i], b = positions[j];
              const snapAndRecord = (axis, val) => {
                positions[i][axis] = val; positions[j][axis] = val;
                for (const idx of [i, j]) {
                  const s = nonTitleShapes[idx];
                  const ex = gridFixes.find(f => String(f.shapeId) === String(s.id));
                  if (ex) ex.position[axis] = val;
                  else gridFixes.push({ shapeName: s.name, shapeId: s.id, _slideShape: s, shapeFill: s.shapeFill||null, position: { ...positions[idx] } });
                }
              };
              // Snap tops if within 1 grid cell
              if (Math.abs(a.top - b.top) > 0.001 && Math.abs(a.top - b.top) <= cellH) {
                snapAndRecord("top", (a.top + b.top) / 2);
              }
              // Snap lefts if within 1 grid cell
              if (Math.abs(a.left - b.left) > 0.001 && Math.abs(a.left - b.left) <= cellW) {
                snapAndRecord("left", (a.left + b.left) / 2);
              }
              // Snap bottom edges if within 1 grid cell
              const aBot = a.top + a.height, bBot = b.top + b.height;
              if (Math.abs(aBot - bBot) > 0.001 && Math.abs(aBot - bBot) <= cellH) {
                const avg = (aBot + bBot) / 2;
                snapAndRecord("top", avg - a.height);
                positions[j].top = avg - b.height;
                const exj = gridFixes.find(f => String(f.shapeId) === String(nonTitleShapes[j].id));
                if (exj) exj.position.top = avg - b.height;
              }
              // Snap right edges if within 1 grid cell
              const aRight = a.left + a.width, bRight = b.left + b.width;
              if (Math.abs(aRight - bRight) > 0.001 && Math.abs(aRight - bRight) <= cellW) {
                const avg = (aRight + bRight) / 2;
                snapAndRecord("left", avg - a.width);
                positions[j].left = avg - b.width;
                const exj = gridFixes.find(f => String(f.shapeId) === String(nonTitleShapes[j].id));
                if (exj) exj.position.left = avg - b.width;
              }
            }
          }


          const distributed = new Set();
          for (let i = 0; i < nonTitleShapes.length; i++) {
            if (distributed.has(i)) continue;
            const group = [i];
            for (let j = i + 1; j < nonTitleShapes.length; j++) {
              const a = positions[i], b = positions[j];
              if (Math.abs(a.width-b.width)/a.width<=0.10 && Math.abs(a.height-b.height)/a.height<=0.10) group.push(j);
            }
            if (group.length < 3) continue;
            group.forEach(idx => distributed.add(idx));
            const gp = group.map(idx => positions[idx]);
            const topsAligned  = gp.every(p => Math.abs(p.top  - gp[0].top)  / gp[0].height <= 0.15);
            const leftsAligned = gp.every(p => Math.abs(p.left - gp[0].left) / gp[0].width  <= 0.15);
            if (topsAligned) {
              const sorted = [...group].sort((x, y) => positions[x].left - positions[y].left);
              const lmost = positions[sorted[0]].left, rmost = positions[sorted[sorted.length-1]].left + positions[sorted[sorted.length-1]].width;
              const gap = (rmost - lmost - sorted.reduce((s, idx) => s + positions[idx].width, 0)) / (sorted.length - 1);
              let cursor = lmost;
              for (const idx of sorted) {
                if (Math.abs(positions[idx].left - cursor) > 0.001) {
                  positions[idx].left = cursor;
                  const ex = gridFixes.find(f => String(f.shapeId) === String(nonTitleShapes[idx].id));
                  if (ex) ex.position.left = cursor;
                  else gridFixes.push({ shapeName: nonTitleShapes[idx].name, shapeId: nonTitleShapes[idx].id, _slideShape: nonTitleShapes[idx], shapeFill: nonTitleShapes[idx].shapeFill||null, position: { ...positions[idx] } });
                }
                cursor += positions[idx].width + gap;
              }
            }
            if (leftsAligned) {
              const sorted = [...group].sort((x, y) => positions[x].top - positions[y].top);
              const tmost = positions[sorted[0]].top, bmost = positions[sorted[sorted.length-1]].top + positions[sorted[sorted.length-1]].height;
              const gap = (bmost - tmost - sorted.reduce((s, idx) => s + positions[idx].height, 0)) / (sorted.length - 1);
              let cursor = tmost;
              for (const idx of sorted) {
                if (Math.abs(positions[idx].top - cursor) > 0.001) {
                  positions[idx].top = cursor;
                  const ex = gridFixes.find(f => String(f.shapeId) === String(nonTitleShapes[idx].id));
                  if (ex) ex.position.top = cursor;
                  else gridFixes.push({ shapeName: nonTitleShapes[idx].name, shapeId: nonTitleShapes[idx].id, _slideShape: nonTitleShapes[idx], shapeFill: nonTitleShapes[idx].shapeFill||null, position: { ...positions[idx] } });
                }
                cursor += positions[idx].height + gap;
              }
            }
          }

          // Resolve overlaps
          for (let i = 0; i < nonTitleShapes.length; i++) {
            for (let j = i + 1; j < nonTitleShapes.length; j++) {
              const a = positions[i], b = positions[j];
              const overlapX = Math.min(a.left+a.width, b.left+b.width) - Math.max(a.left, b.left);
              const overlapY = Math.min(a.top+a.height, b.top+b.height) - Math.max(a.top,  b.top);
              if (overlapX > 0.01 && overlapY > 0.01) {
                if (overlapX + cellW <= overlapY + cellH) positions[j].left += overlapX + cellW;
                else positions[j].top  += overlapY + cellH;
                const ex = gridFixes.find(f => String(f.shapeId) === String(nonTitleShapes[j].id));
                if (ex) ex.position = { ...positions[j] };
                else gridFixes.push({ shapeName: nonTitleShapes[j].name, shapeId: nonTitleShapes[j].id, _slideShape: nonTitleShapes[j], shapeFill: nonTitleShapes[j].shapeFill||null, position: { ...positions[j] } });
              }
            }
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

      // ─── SAFETY NET: push any remaining overlaps apart ───────────────────────
      await PowerPoint.run(async (ctx) => {
        const slides = ctx.presentation.slides;
        slides.load("items");
        await ctx.sync();
        const shapes = slides.items[dupIndex - 1].shapes;
        shapes.load("items");
        await ctx.sync();
        for (const s of shapes.items) s.load(["id", "name", "left", "top", "width", "height"]);
        await ctx.sync();
        const live = shapes.items.map(s => ({ s, left: s.left/72, top: s.top/72, width: s.width/72, height: s.height/72 })).filter(s => s.width > 0.1 && s.height > 0.1);
        let changed = true;
        for (let iter = 0; iter < 8 && changed; iter++) {
          changed = false;
          for (let i = 0; i < live.length; i++) {
            for (let j = i + 1; j < live.length; j++) {
              const a = live[i], b = live[j];
              const overX = a.left < b.left + b.width - 0.05 && a.left + a.width > b.left + 0.05;
              const overY = a.top  < b.top  + b.height - 0.05 && a.top  + a.height > b.top  + 0.05;
              if (!overX || !overY) continue;
              const lower = a.top >= b.top ? a : b;
              const upper = a.top >= b.top ? b : a;
              const newTop = upper.top + upper.height + 0.1;
              if (newTop + lower.height <= 7.5) { lower.top = newTop; lower.s.top = newTop * 72; changed = true; }
            }
          }
        }
        await ctx.sync();
      });

      addLog(`✓ Done — ${totalFixes} fix${totalFixes !== 1 ? "es" : ""} applied`);
      setFixCount(totalFixes);
      setStatus("done");
    } catch (err) {
      setError(err.message);
      addLog("✗ " + err.message);
      setStatus("error");
    }
  }, []);

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
            <div style={{ fontWeight: 700, fontSize: 15 }}>Plz fix thx</div>
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

        <button className="btn" onClick={handleCleanup} disabled={isRunning}
          style={{ width: "100%", padding: "14px 0", background: status === "done" ? "#15803d" : "#111111", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "all 0.2s ease", boxShadow: "0 4px 14px rgba(0,0,0,0.28)" }}>
          {isRunning ? (
            <><span style={{ width: 14, height: 14, border: "2px solid rgba(255,255,255,0.3)", borderTop: "2px solid #fff", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block" }} />Working…</>
          ) : status === "done" ? "✓ Done — clean another?" : "Plz fix thx"}
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
