import { useState, useCallback } from "react";
import JSZip from "jszip";

const STORAGE_KEY = "ppt_cleanup_api_key";

/* ═══════════════════════════════════════════════════════════════════════════
   XML READING — get the raw .pptx bytes from Office, unzip, parse XML
   ═══════════════════════════════════════════════════════════════════════════ */

function getFileBytes() {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      reject(new Error("File read timed out — try again"));
    }, 30000);

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
              // PowerPoint Online returns slices as regular arrays, not ArrayBuffers
              // Handle both formats
              const byteArrays = slices.map(s => {
                if (s instanceof ArrayBuffer) return new Uint8Array(s);
                if (s instanceof Uint8Array) return s;
                // Plain array of numbers
                return new Uint8Array(s);
              });
              const total = byteArrays.reduce((n, s) => n + s.length, 0);
              const combined = new Uint8Array(total);
              let offset = 0;
              for (const arr of byteArrays) {
                combined.set(arr, offset);
                offset += arr.length;
              }
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

// Resolve a theme colour reference (e.g. "dk1", "lt1", "accent1") to hex
function resolveThemeColor(ref, themeColors) {
  const map = {
    dk1: "dark1", dk2: "dark2",
    lt1: "light1", lt2: "light2",
    accent1: "accent1", accent2: "accent2",
    accent3: "accent3", accent4: "accent4",
    accent5: "accent5", accent6: "accent6",
    hlink: "hyperlink", folHlink: "followedHyperlink",
  };
  return themeColors[map[ref] || ref] || null;
}

// Parse <a:srgbClr val="RRGGBB"> or <a:sysClr> or <a:schemeClr> into hex
function extractColor(el, themeColors) {
  if (!el) return null;
  const srgb = el.querySelector("srgbClr");
  if (srgb) return "#" + srgb.getAttribute("val");
  const sys = el.querySelector("sysClr");
  if (sys) return "#" + (sys.getAttribute("lastClr") || "000000");
  const scheme = el.querySelector("schemeClr");
  if (scheme) return resolveThemeColor(scheme.getAttribute("val"), themeColors);
  return null;
}

// Extract font size from <a:sz val="2400"> (hundredths of a point)
function extractFontSize(el) {
  if (!el) return null;
  const sz = el.querySelector("sz");
  if (sz) return parseInt(sz.getAttribute("val"), 10) / 100;
  return null;
}

// Extract font name from <a:latin typeface="Calibri"> or <a:ea> / <a:cs>
function extractFontName(el) {
  if (!el) return null;
  const latin = el.querySelector("latin");
  if (latin) {
    const tf = latin.getAttribute("typeface");
    // +mj-lt = major (heading) font, +mn-lt = minor (body) font — resolved later
    if (tf && !tf.startsWith("+")) return tf;
  }
  return null;
}

function isBold(el) {
  if (!el) return null;
  const rPr = el.querySelector("rPr");
  if (!rPr) return null;
  const b = rPr.getAttribute("b");
  return b === "1" || b === "true" ? true : b === "0" || b === "false" ? false : null;
}

// pt to inches
function emuToInches(emu) {
  return parseFloat((parseInt(emu, 10) / 914400).toFixed(3));
}

/* ── Parse theme1.xml ───────────────────────────────────────────────────── */

function parseThemeXml(xml) {
  const doc = parseXml(xml);

  // Colour scheme
  const colors = {};
  const colorMap = {
    dk1: "dark1", dk2: "dark2", lt1: "light1", lt2: "light2",
    accent1: "accent1", accent2: "accent2", accent3: "accent3",
    accent4: "accent4", accent5: "accent5", accent6: "accent6",
  };
  for (const [tag, name] of Object.entries(colorMap)) {
    // Try namespace-aware and plain
    const el = doc.querySelector(`${tag}, [*|tag="${tag}"]`) ||
               doc.getElementsByTagNameNS("*", tag)[0];
    if (el) {
      const srgb = el.querySelector("srgbClr") ||
                   el.getElementsByTagNameNS("*", "srgbClr")[0];
      const sys  = el.querySelector("sysClr") ||
                   el.getElementsByTagNameNS("*", "sysClr")[0];
      if (srgb) colors[name] = "#" + srgb.getAttribute("val");
      else if (sys) colors[name] = "#" + (sys.getAttribute("lastClr") || "000000");
    }
  }

  // Font scheme
  const majorEl = doc.getElementsByTagNameNS("*", "majorFont")[0];
  const minorEl = doc.getElementsByTagNameNS("*", "minorFont")[0];
  const majorLatin = majorEl?.getElementsByTagNameNS("*", "latin")[0];
  const minorLatin = minorEl?.getElementsByTagNameNS("*", "latin")[0];

  return {
    colors,
    fonts: {
      heading: majorLatin?.getAttribute("typeface") || null,
      body: minorLatin?.getAttribute("typeface") || null,
    },
  };
}

/* ── Parse slideMaster1.xml ─────────────────────────────────────────────── */

function parseMasterXml(xml, theme) {
  const doc = parseXml(xml);
  const placeholders = [];

  // sp = shape elements
  const shapes = doc.getElementsByTagNameNS("*", "sp");
  for (const sp of shapes) {
    // ph = placeholder element
    const ph = sp.getElementsByTagNameNS("*", "ph")[0];
    const phType = ph?.getAttribute("type") || "body";
    const phIdx  = ph?.getAttribute("idx") || "0";

    // Position & size from spPr/xfrm
    const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
    const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];

    const position = off && ext ? {
      left:   emuToInches(off.getAttribute("x")),
      top:    emuToInches(off.getAttribute("y")),
      width:  emuToInches(ext.getAttribute("cx")),
      height: emuToInches(ext.getAttribute("cy")),
    } : null;

    // Font from txBody > lstStyle > lvl1pPr / defRPr
    // or directly on the paragraph run
    const txBody    = sp.getElementsByTagNameNS("*", "txBody")[0];
    const lstStyle  = txBody?.getElementsByTagNameNS("*", "lstStyle")[0];
    const lvl1pPr   = lstStyle?.getElementsByTagNameNS("*", "lvl1pPr")[0];
    const defRPr    = lvl1pPr?.getElementsByTagNameNS("*", "defRPr")[0];

    // Also check direct paragraph runs
    const firstPara = txBody?.getElementsByTagNameNS("*", "p")[0];
    const pPr       = firstPara?.getElementsByTagNameNS("*", "pPr")[0];
    const firstRPr  = firstPara?.getElementsByTagNameNS("*", "r")[0]
                        ?.getElementsByTagNameNS("*", "rPr")[0];

    const rPr = defRPr || firstRPr;

    // Font name — check rPr first, then fall back to theme fonts
    let fontName = null;
    if (rPr) {
      const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
      const tf = latin?.getAttribute("typeface");
      if (tf && tf.startsWith("+mj")) fontName = theme.fonts.heading;
      else if (tf && tf.startsWith("+mn")) fontName = theme.fonts.body;
      else if (tf) fontName = tf;
    }
    // Default by placeholder type
    if (!fontName) {
      fontName = (phType === "title" || phType === "ctrTitle")
        ? theme.fonts.heading
        : theme.fonts.body;
    }

    // Font size
    let fontSize = null;
    if (rPr) fontSize = extractFontSize({ querySelector: (s) => rPr.getElementsByTagNameNS("*", s.replace("a:", ""))[0] });
    if (!fontSize && defRPr) {
      const sz = defRPr.getAttribute("sz");
      if (sz) fontSize = parseInt(sz, 10) / 100;
    }
    if (!fontSize) fontSize = (phType === "title" || phType === "ctrTitle") ? 36 : 18;

    // Colour
    let color = null;
    if (rPr) {
      const solidFill = rPr.getElementsByTagNameNS("*", "solidFill")[0];
      if (solidFill) {
        const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
        const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
        const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
        if (srgb)   color = "#" + srgb.getAttribute("val");
        else if (sys) color = "#" + (sys.getAttribute("lastClr") || "000000");
        else if (scheme) color = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
      }
    }
    if (!color) {
      color = (phType === "title" || phType === "ctrTitle")
        ? (theme.colors.dark1 || "#000000")
        : (theme.colors.dark2 || theme.colors.dark1 || "#000000");
    }

    // Bold
    let bold = null;
    if (rPr) {
      const b = rPr.getAttribute("b");
      bold = b === "1" || b === "true";
    }
    if (bold === null) bold = (phType === "title" || phType === "ctrTitle");

    // Alignment
    let alignment = "left";
    const algn = pPr?.getAttribute("algn") || lvl1pPr?.getAttribute("algn");
    if (algn === "ctr") alignment = "center";
    else if (algn === "r") alignment = "right";
    else if (algn === "just") alignment = "justify";

    // Master placeholder fill
    let masterFill = null;
    const spPrEl = sp.getElementsByTagNameNS("*", "spPr")[0];
    if (spPrEl) {
      const noFill = spPrEl.getElementsByTagNameNS("*", "noFill")[0];
      if (noFill) {
        masterFill = "none";
      } else {
        const solidFill = spPrEl.getElementsByTagNameNS("*", "solidFill")[0];
        if (solidFill) {
          const srgb = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
          const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
          if (srgb) masterFill = "#" + srgb.getAttribute("val").toUpperCase();
          else if (scheme) {
            const resolved = resolveThemeColor(scheme.getAttribute("val"), theme.colors);
            masterFill = resolved || ("theme:" + scheme.getAttribute("val"));
          }
        }
      }
    }

    placeholders.push({
      type: phType,
      idx: phIdx,
      font: { name: fontName, size: fontSize, color, bold },
      alignment,
      position,
      fill: masterFill,
    });
  }

  return placeholders;
}

/* ── Parse slide XML for current slide shapes ───────────────────────────── */

function parseSlideXml(xml, theme, masterPlaceholders, layoutPositions = {}) {
  const doc = parseXml(xml);
  const shapes = [];

  const spEls = doc.getElementsByTagNameNS("*", "sp");
  // Log ALL shape-like elements to find non-sp shapes
  const allShapeTypes = ["sp", "pic", "cxnSp", "graphicFrame", "grpSp"];
  const foundElements = {};
  for (const type of allShapeTypes) {
    const els = doc.getElementsByTagNameNS("*", type);
    if (els.length > 0) {
      foundElements[type] = Array.from(els).map(el => {
        const cNvPr = el.getElementsByTagNameNS("*", "cNvPr")[0];
        const solidFill = el.getElementsByTagNameNS("*", "solidFill")[0];
        const srgb = solidFill?.getElementsByTagNameNS("*", "srgbClr")[0];
        return `${cNvPr?.getAttribute("name")} fill=${srgb ? "#"+srgb.getAttribute("val") : "none"}`;
      });
    }
  }
  console.log("All shape elements:", JSON.stringify(foundElements));
  console.log("All sp elements in slide:", Array.from(spEls).map(sp => {
    const cNvPr = sp.getElementsByTagNameNS("*", "cNvPr")[0];
    const ph = sp.getElementsByTagNameNS("*", "ph")[0];
    return `${cNvPr?.getAttribute("name")} [ph=${ph ? "yes" : "no"}]`;
  }));
  for (const sp of spEls) {
    const nvSpPr = sp.getElementsByTagNameNS("*", "nvSpPr")[0];
    const cNvPr  = nvSpPr?.getElementsByTagNameNS("*", "cNvPr")[0];
    const ph     = sp.getElementsByTagNameNS("*", "ph")[0];

    const id   = cNvPr?.getAttribute("id") || "";
    const name = cNvPr?.getAttribute("name") || "";
    const phType = ph?.getAttribute("type") || "body";
    const phIdx  = ph?.getAttribute("idx") || "0";

    // Position — use slide XML first, fall back to layout position
    const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
    const off  = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext  = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const position = off && ext ? {
      left:   emuToInches(off.getAttribute("x")),
      top:    emuToInches(off.getAttribute("y")),
      width:  emuToInches(ext.getAttribute("cx")),
      height: emuToInches(ext.getAttribute("cy")),
    } : (layoutPositions[`${phType}:${phIdx}`] || layoutPositions[`body:0`] || null);
    let shapeFill = null;
    let shapeBorder = null;
    const spPr = sp.getElementsByTagNameNS("*", "spPr")[0];
    if (spPr) {
      // Only check direct children of spPr for fill (not descendants which could be border fills)
      const directChildren = Array.from(spPr.childNodes);
      const noFillEl = directChildren.find(n => n.localName === "noFill");
      const solidFillEl = directChildren.find(n => n.localName === "solidFill");
      const gradFillEl = directChildren.find(n => n.localName === "gradFill");

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
        // Gradient fill — get first stop colour as representative
        const firstStop = gradFillEl.getElementsByTagNameNS("*", "gs")[0];
        const srgb = firstStop?.getElementsByTagNameNS("*", "srgbClr")[0];
        if (srgb) shapeFill = "#" + srgb.getAttribute("val").toUpperCase() + " (gradient)";
        else shapeFill = "gradient";
      }

      // Border/outline
      const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
      if (ln) {
        const lnNoFill = ln.getElementsByTagNameNS("*", "noFill")[0];
        if (lnNoFill) {
          shapeBorder = "none";
        } else {
          const lnSolidFill = ln.getElementsByTagNameNS("*", "solidFill")[0];
          if (lnSolidFill) {
            const srgb   = lnSolidFill.getElementsByTagNameNS("*", "srgbClr")[0];
            const scheme = lnSolidFill.getElementsByTagNameNS("*", "schemeClr")[0];
            const sys    = lnSolidFill.getElementsByTagNameNS("*", "sysClr")[0];
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

    // Text content
    const txBody = sp.getElementsByTagNameNS("*", "txBody")[0];
    if (!txBody) continue;

    // Collect all text
    const runs = txBody.getElementsByTagNameNS("*", "r");
    const textContent = Array.from(runs).map(r => {
      const t = r.getElementsByTagNameNS("*", "t")[0];
      return t?.textContent || "";
    }).join("");

    if (!textContent.trim() && runs.length === 0) continue;

    // Get font from runs — scan all runs AND paragraph-level overrides for explicit colours
    const allRuns = Array.from(txBody.getElementsByTagNameNS("*", "r"));
    const firstRun = allRuns[0];
    const rPr = firstRun?.getElementsByTagNameNS("*", "rPr")[0];
    const allParas = Array.from(txBody.getElementsByTagNameNS("*", "p"));

    let fontName = null, fontSize = null, color = null, bold = null;

    // Helper to extract colour from a solidFill element's parent
    function extractColorFromEl(el) {
      if (!el) return null;
      const solidFill = el.getElementsByTagNameNS("*", "solidFill")[0];
      if (!solidFill) return null;
      const srgb   = solidFill.getElementsByTagNameNS("*", "srgbClr")[0];
      const scheme = solidFill.getElementsByTagNameNS("*", "schemeClr")[0];
      const sys    = solidFill.getElementsByTagNameNS("*", "sysClr")[0];
      if (srgb)    return "#" + srgb.getAttribute("val");
      if (sys)     return "#" + (sys.getAttribute("lastClr") || "000000");
      if (scheme)  return resolveThemeColor(scheme.getAttribute("val"), theme.colors);
      return null;
    }

    // Priority 1: explicit colour on runs (most specific)
    for (const run of allRuns) {
      const rp = run.getElementsByTagNameNS("*", "rPr")[0];
      const c = extractColorFromEl(rp);
      if (c) { color = c; break; }
    }

    // Priority 2: paragraph-level defRPr colour
    if (!color) {
      for (const para of allParas) {
        const pPrEl = para.getElementsByTagNameNS("*", "pPr")[0];
        const defRPr = pPrEl?.getElementsByTagNameNS("*", "defRPr")[0];
        const c = extractColorFromEl(defRPr);
        if (c) { color = c; break; }
      }
    }

    // Priority 3: txBody lstStyle defRPr colour
    if (!color) {
      const lstStyle = txBody.getElementsByTagNameNS("*", "lstStyle")[0];
      const lvl1pPr = lstStyle?.getElementsByTagNameNS("*", "lvl1pPr")[0];
      const defRPr = lvl1pPr?.getElementsByTagNameNS("*", "defRPr")[0];
      const c = extractColorFromEl(defRPr);
      if (c) color = c;
    }

    if (rPr) {
      // Font name
      const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
      const tf = latin?.getAttribute("typeface");
      if (tf && tf.startsWith("+mj")) fontName = theme.fonts.heading;
      else if (tf && tf.startsWith("+mn")) fontName = theme.fonts.body;
      else if (tf) fontName = tf;

      // Size
      const sz = rPr.getAttribute("sz");
      if (sz) fontSize = parseInt(sz, 10) / 100;

      // Bold + Italic
      const b = rPr.getAttribute("b");
      bold = b === "1" || b === "true" ? true : b === "0" || b === "false" ? false : null;
    }

    // Also check italic on first run
    let italic = null;
    if (rPr) {
      const i = rPr.getAttribute("i");
      italic = i === "1" || i === "true" ? true : i === "0" || i === "false" ? false : null;
    }

    // Paragraph alignment
    const firstPara = txBody.getElementsByTagNameNS("*", "p")[0];
    const pPr = firstPara?.getElementsByTagNameNS("*", "pPr")[0];
    let alignment = "left";
    const algn = pPr?.getAttribute("algn");
    if (algn === "ctr") alignment = "center";
    else if (algn === "r") alignment = "right";

    // Find matching master placeholder for inherited values
    const masterPh = masterPlaceholders.find(
      (p) => p.type === phType || (phType === "body" && p.idx === phIdx)
    ) || masterPlaceholders.find((p) => p.type === "body");

    // Use layout position (per phIdx) as the target position — more accurate than master
    const layoutTargetPos = layoutPositions[`${phType}:${phIdx}`] || layoutPositions[`${phType}:0`] || masterPh?.position || null;

    shapes.push({
      id,
      name,
      phType,
      phIdx,
      position,
      shapeFill,
      shapeBorder,
      textContent: textContent.substring(0, 100),
      current: {
        fontName: fontName || "(inherited)",
        fontSize: fontSize || "(inherited)",
        color: color || "(inherited)",
        bold: bold !== null ? bold : "(inherited)",
        italic: italic !== null ? italic : "(inherited)",
        alignment,
      },
      masterTarget: masterPh ? {
        fontName: masterPh.font.name,
        fontSize: masterPh.font.size,
        color: masterPh.font.color,
        bold: masterPh.font.bold,
        alignment: masterPh.alignment,
        position: layoutTargetPos,
        fill: masterPh.fill || "none",
      } : null,
    });
  }

  return shapes;
}

/* ── Read all masters from the file for the picker ──────────────────────── */

async function readAllMasters(zip) {
  // presentation.xml.rels tells us exactly which masters exist and in what order
  const relsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const masters = [];

  if (relsFile) {
    const relsXml = await relsFile.async("string");
    const relsDoc = parseXml(relsXml);
    const rels = relsDoc.getElementsByTagNameNS("*", "Relationship");

    for (const rel of rels) {
      const target = rel.getAttribute("Target") || "";
      const match = target.match(/slideMasters\/slideMaster(\d+)\.xml/);
      if (!match) continue;

      const masterIndex = parseInt(match[1], 10);
      const masterPath = `ppt/slideMasters/slideMaster${masterIndex}.xml`;

      // Each master references its own theme via its own rels file
      const masterRelsPath = `ppt/slideMasters/_rels/slideMaster${masterIndex}.xml.rels`;
      let theme = { colors: {}, fonts: { heading: null, body: null } };

      const masterRelsFile = zip.file(masterRelsPath);
      if (masterRelsFile) {
        const masterRelsXml = await masterRelsFile.async("string");
        const masterRelsDoc = parseXml(masterRelsXml);
        const masterRels = masterRelsDoc.getElementsByTagNameNS("*", "Relationship");
        for (const mRel of masterRels) {
          const mTarget = mRel.getAttribute("Target") || "";
          const themeMatch = mTarget.match(/\.\.\/theme\/theme(\d+)\.xml/);
          if (themeMatch) {
            const themeFile = zip.file(`ppt/theme/theme${themeMatch[1]}.xml`);
            if (themeFile) {
              const themeXml = await themeFile.async("string");
              theme = parseThemeXml(themeXml);
            }
            break;
          }
        }
      }

      // Parse master name from XML (stored in <p:cSld name="...">)
      const masterFile = zip.file(masterPath);
      if (!masterFile) continue;
      const masterXml = await masterFile.async("string");
      const masterDoc = parseXml(masterXml);
      const cSld = masterDoc.getElementsByTagNameNS("*", "cSld")[0];
      const masterName = cSld?.getAttribute("name") || `Master ${masterIndex}`;

      const placeholders = parseMasterXml(masterXml, theme);

      masters.push({
        index: masterIndex,
        name: masterName,
        theme,
        placeholders,
        // Summary info for the picker UI
        headingFont: theme.fonts.heading,
        bodyFont: theme.fonts.body,
        colors: Object.entries(theme.colors).filter(([, v]) => v).slice(0, 5).map(([, v]) => v),
      });
    }
  }

  // Fallback: if rels parsing failed, try master1 directly
  if (masters.length === 0) {
    const masterFile = zip.file("ppt/slideMasters/slideMaster1.xml");
    if (masterFile) {
      const themeFile = zip.file("ppt/theme/theme1.xml");
      const theme = themeFile
        ? parseThemeXml(await themeFile.async("string"))
        : { colors: {}, fonts: { heading: null, body: null } };
      const masterXml = await masterFile.async("string");
      masters.push({
        index: 1,
        name: "Master 1",
        theme,
        placeholders: parseMasterXml(masterXml, theme),
        headingFont: theme.fonts.heading,
        bodyFont: theme.fonts.body,
        colors: Object.entries(theme.colors).filter(([, v]) => v).slice(0, 5).map(([, v]) => v),
      });
    }
  }

  return masters;
}

/* ── Main XML reader — ties it all together ─────────────────────────────── */

// Phase 1: just read the file bytes and discover all masters
async function readPptxFile() {
  const bytes = await getFileBytes();
  const zip = await JSZip.loadAsync(bytes);
  const masters = await readAllMasters(zip);
  return { zip, masters };
}

// Phase 2: read slide data using the chosen master
async function readSlideWithMaster(zip, masters, chosenMasterIndex, selectedSlideIndex) {
  const master = masters.find((m) => m.index === chosenMasterIndex) || masters[0];

  const slideFile = zip.file(`ppt/slides/slide${selectedSlideIndex}.xml`);
  if (!slideFile) throw new Error(`Slide ${selectedSlideIndex} not found in file`);
  const slideXml = await slideFile.async("string");

  // Also read the slide's layout XML to get fallback positions
  // The slide rels file tells us which layout this slide uses
  const slideRelsFile = zip.file(`ppt/slides/_rels/slide${selectedSlideIndex}.xml.rels`);
  let layoutPositions = {}; // phIdx/phType → position

  if (slideRelsFile) {
    const slideRelsXml = await slideRelsFile.async("string");
    const relsDoc = parseXml(slideRelsXml);
    const rels = relsDoc.getElementsByTagNameNS("*", "Relationship");
    for (const rel of rels) {
      const target = rel.getAttribute("Target") || "";
      const match = target.match(/slideLayouts\/slideLayout(\d+)\.xml/);
      if (match) {
        const layoutFile = zip.file(`ppt/slideLayouts/slideLayout${match[1]}.xml`);
        if (layoutFile) {
          const layoutXml = await layoutFile.async("string");
          const layoutDoc = parseXml(layoutXml);
          const layoutShapes = layoutDoc.getElementsByTagNameNS("*", "sp");
          for (const sp of layoutShapes) {
            const ph = sp.getElementsByTagNameNS("*", "ph")[0];
            if (!ph) continue;
            const phType = ph.getAttribute("type") || "body";
            const phIdx = ph.getAttribute("idx") || "0";
            const xfrm = sp.getElementsByTagNameNS("*", "xfrm")[0];
            const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
            const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
            if (off && ext) {
              const key = `${phType}:${phIdx}`;
              layoutPositions[key] = {
                left:   emuToInches(off.getAttribute("x")),
                top:    emuToInches(off.getAttribute("y")),
                width:  emuToInches(ext.getAttribute("cx")),
                height: emuToInches(ext.getAttribute("cy")),
              };
            }
          }
        }
        break;
      }
    }
  }

  const slideShapes = parseSlideXml(slideXml, master.theme, master.placeholders, layoutPositions);
  console.log("Layout positions found:", JSON.stringify(layoutPositions));
  console.log("Slide shape positions:", slideShapes.map(s => `${s.name}: ${JSON.stringify(s.position)}`));

  return {
    theme: master.theme,
    masterPlaceholders: master.placeholders,
    masterName: master.name,
    slideShapes,
    layoutPositions,
    slideIndex: selectedSlideIndex,
  };
}

/* ═══════════════════════════════════════════════════════════════════════════
   OFFICE JS — get selected slide index + apply fixes (writes stay via JS API)
   ═══════════════════════════════════════════════════════════════════════════ */

function getSelectedSlideIndex() {
  return new Promise((resolve, reject) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.SlideRange,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed)
          return reject(new Error(result.error.message));
        resolve(result.value.slides[0].index);
      }
    );
  });
}

async function duplicateSlide(slideIndex) {
  // PowerPoint Online doesn't support reliable slide duplication via JS API
  // Work on the original slide — user can Ctrl+Z to undo
  return slideIndex;
}

/* ── Colour snapping — find nearest theme colour within threshold ─────────── */

function hexToRgb(hex) {
  const h = hex.replace("#", "");
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16),
  };
}

function colourDistance(hex1, hex2) {
  const a = hexToRgb(hex1), b = hexToRgb(hex2);
  // Weighted Euclidean distance (human perception weighting)
  return Math.sqrt(
    2 * Math.pow(a.r - b.r, 2) +
    4 * Math.pow(a.g - b.g, 2) +
    3 * Math.pow(a.b - b.b, 2)
  );
}

// Snap a hex colour to the nearest theme colour if within threshold (max ~80 on 0-765 scale)
function snapToThemeColor(hex, themeColors, threshold = 80) {
  if (!hex || hex === "none" || hex.startsWith("theme:")) return hex;
  let nearest = null, nearestDist = Infinity;
  for (const [, themeHex] of Object.entries(themeColors)) {
    if (!themeHex) continue;
    const dist = colourDistance(hex, themeHex);
    if (dist < nearestDist) {
      nearestDist = dist;
      nearest = themeHex;
    }
  }
  if (nearestDist <= threshold) {
    console.log(`  Snapping ${hex} → ${nearest} (distance: ${nearestDist.toFixed(0)})`);
    return nearest;
  }
  return hex; // too different from any theme colour — return as-is
}

async function applyFixes(slideIndex, fixes, themeColors = {}) {
  const themeColorValues = new Set(
    Object.values(themeColors).filter(Boolean).map(v => v.toUpperCase().replace("#", ""))
  );

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

    console.log("Shapes on slide:", shapes.items.map(s => s.name));
    console.log("Fixes from Claude:", JSON.stringify(fixes));

    for (const fix of fixes) {
      // shapeIndex refers to our pptxData.slideShapes array order
      // Use the shape's id/name from that array to find the correct Office JS shape
      const slideShape = fix.shapeIndex !== undefined ? fix._slideShape : null;
      const lookupName = slideShape?.name || fix.shapeName;
      const lookupId = slideShape?.id || fix.shapeId;
      const target = lookupId
        ? (shapes.items.find(s => s.id === lookupId) || shapes.items.find(s => s.name === lookupName))
        : shapes.items.find(s => s.name === lookupName || s.id === fix.shapeId);
      console.log(`Fix for "${lookupName}" (idx:${fix.shapeIndex}) → ${target ? `FOUND: ${target.name}` : "NOT FOUND"}`);
      if (!target) continue;

      // Position
      if (fix.position) {
        const inchToPt = 72;
        const { left, top, width, height } = fix.position;
        if (left   !== undefined && left   > 0   && left   < 10)  { console.log(`  left: ${target.left} → ${left * inchToPt}`);  target.left   = left   * inchToPt; }
        if (top    !== undefined && top    > 0   && top    < 7.5) { console.log(`  top: ${target.top} → ${top * inchToPt}`);    target.top    = top    * inchToPt; }
        if (width  !== undefined && width  > 0.5 && width  <= 10) { target.width  = width  * inchToPt; }
        if (height !== undefined && height > 0.1 && height <= 7.5){ target.height = height * inchToPt; }
      }

      // Fill — keep theme colours, snap near-theme, clear non-theme
      if (fix.fill !== undefined && fix.fill !== null) {
        try {
          if (fix.fill === "none" && fix.shapeFill && fix.shapeFill !== "none") {
            const isExactTheme = Object.values(themeColors).some(c => c && c.toUpperCase() === fix.shapeFill.toUpperCase());
            const snapped = snapToThemeColor(fix.shapeFill, themeColors);
            const didSnap = snapped !== fix.shapeFill;
            if (isExactTheme) {
              console.log(`  Keeping exact theme fill (${fix.shapeFill})`);
            } else if (didSnap) {
              target.fill.setSolidColor(snapped.replace("#", ""));
              console.log(`  ✓ Fill snapped ${fix.shapeFill} → ${snapped}`);
            } else {
              target.fill.clear();
              console.log(`  ✓ Non-theme fill cleared (was ${fix.shapeFill})`);
            }
          } else if (fix.fill && fix.fill !== "none") {
            const snapped = snapToThemeColor(fix.fill, themeColors);
            target.fill.setSolidColor(snapped.replace("#", ""));
            console.log(`  ✓ Fill set to ${snapped}`);
          }
        } catch (e) { console.log(`  Error setting fill:`, e.message); }
      }

      // Border — same logic as fill
      if (fix.border !== undefined && fix.border !== null) {
        try {
          if (fix.border === "none" && fix.shapeBorder && fix.shapeBorder !== "none") {
            const isExactTheme = Object.values(themeColors).some(c => c && c.toUpperCase() === fix.shapeBorder.toUpperCase());
            const snapped = snapToThemeColor(fix.shapeBorder, themeColors);
            if (isExactTheme) {
              console.log(`  Keeping exact theme border (${fix.shapeBorder})`);
            } else if (snapped !== fix.shapeBorder) {
              target.lineFormat.color = snapped.replace("#", "");
              target.lineFormat.visible = true;
              console.log(`  ✓ Border snapped ${fix.shapeBorder} → ${snapped}`);
            } else {
              target.lineFormat.visible = false;
              console.log(`  ✓ Non-theme border removed (was ${fix.shapeBorder})`);
            }
          } else if (fix.border && fix.border !== "none") {
            const snapped = snapToThemeColor(fix.border, themeColors);
            target.lineFormat.color = snapped.replace("#", "");
            target.lineFormat.visible = true;
          }
        } catch (e) { console.log(`  Error setting border:`, e.message); }
      }

      // Font, colour, alignment
      if (fix.font || fix.alignment) {
        try {
          const tf = target.textFrame;
          const tr = tf.textRange;
          tr.load(["text"]);
          await ctx.sync();

          if (fix.font) {
            if (fix.font.name)  tr.font.name = fix.font.name;
            if (fix.font.size)  tr.font.size = fix.font.size;
            if (fix.font.color) {
              const snapped = snapToThemeColor(fix.font.color, themeColors);
              tr.font.color = snapped.replace("#", "");
            }
            if (fix.font.bold    !== undefined) tr.font.bold   = fix.font.bold;
            if (fix.font.italic  !== undefined) tr.font.italic = fix.font.italic;
          }

          if (fix.alignment) {
            const alignMap = { left: PowerPoint.ParagraphHorizontalAlignment.left, center: PowerPoint.ParagraphHorizontalAlignment.center, right: PowerPoint.ParagraphHorizontalAlignment.right };
            if (alignMap[fix.alignment]) tr.paragraphFormat.horizontalAlignment = alignMap[fix.alignment];
          }

          await ctx.sync();
          console.log(`  ✓ Font/alignment applied to "${target.name}"`);
        } catch (e) { console.log(`  Error applying font:`, e.message); }
      }
    }
    await ctx.sync();
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   AI — build prompt from XML data and call Claude
   ═══════════════════════════════════════════════════════════════════════════ */

function buildPrompt(pptxData) {
  const { theme, masterPlaceholders, slideShapes } = pptxData;

  const themeSection = `━━━ THEME ━━━
FONTS: heading="${theme.fonts.heading}" body="${theme.fonts.body}"
COLOURS: ${Object.entries(theme.colors).filter(([,v])=>v).map(([k,v])=>`${k}=${v}`).join(" ")}`;

  const masterSection = `━━━ MASTER PLACEHOLDERS ━━━
${masterPlaceholders.map(p => `  [${p.type}] font="${p.font.name}" size=${p.font.size}pt color=${p.font.color} align=${p.alignment}`).join("\n")}`;

  const phTypeCounts = {};
  slideShapes.forEach(s => { phTypeCounts[s.phType] = (phTypeCounts[s.phType] || 0) + 1; });

  const slideSection = slideShapes.map((s, i) => {
    const layoutPos = s.masterTarget?.position;
    const showPos = layoutPos && phTypeCounts[s.phType] === 1;
    return `  Shape #${i} "${s.name}" [${s.phType}]
    font="${s.current.fontName}" size=${s.current.fontSize}pt color=${s.current.color} bold=${s.current.bold} italic=${s.current.italic} align=${s.current.alignment} fill=${s.shapeFill||"none"} border=${s.shapeBorder||"none"}${s.position ? ` pos=(${s.position.left}",${s.position.top}") size=(${s.position.width}"×${s.position.height}")` : ""}
    → target font="${s.masterTarget?.fontName}" size=${s.masterTarget?.fontSize}pt color=${s.masterTarget?.color} align=${s.masterTarget?.alignment} fill=${s.masterTarget?.fill||"none"}${showPos ? ` pos=(${layoutPos.left}",${layoutPos.top}") size=(${layoutPos.width}"×${layoutPos.height}")` : ""}
    text: "${s.textContent}"`;
  }).join("\n\n");

  return `You are a PowerPoint formatting expert. Analyse these slide shapes and return fixes.

${themeSection}

${masterSection}

━━━ SLIDE SHAPES ━━━
${slideSection}

RULES — return ONLY a JSON array, no markdown:
Each item: {"shapeIndex":N,"font":{"name":"...","size":N,"color":"#hex","bold":true/false,"italic":true/false},"alignment":"left|center|right","fill":"#hex or none","border":"#hex or none"}
Use the Shape # number as shapeIndex.

FONT rules:
- Each shape has a "→ target" line — use EXACTLY that target font name and size, not the master placeholder defaults
- ANY shape whose current font ≠ target font MUST be fixed
- Set bold=false and italic=false unless the target says otherwise
- "(inherited)" for the current font still requires a fix if the target font is explicit

COLOUR rules:
- Text colour: if "(inherited)" or matches target → omit color from fix
- If text colour is clearly off-brand → set it to the target colour for that placeholder

FILL/BORDER rules:
- If current fill is an exact theme colour → omit fill from fix (keep it)
- If current fill is NOT a theme colour → include "fill":"none" to remove it
- Theme colours are: #000000 #44546A #FFFFFF #E7E6E6 #55BC7E #FFC330 #BE80FF #FF8345 #FF70BF #60A2F5

ALIGNMENT: set to match master target. Every shape with a non-template font MUST appear in output.`;
}

async function callClaude(pptxData, apiKey) {
  const prompt = buildPrompt(pptxData);
  console.log("=== PROMPT TO CLAUDE ===\n", prompt);
  const response = await fetch("/api/claude", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      system: "You are a PowerPoint formatting assistant. Return ONLY valid JSON arrays. No markdown fences, no explanation, no preamble.",
      messages: [{ role: "user", content: prompt }],
    }),
  });

  if (!response.ok) {
    const err = await response.json();
    throw new Error(err.error?.message || `API error ${response.status}`);
  }

  const data = await response.json();
  const raw = data.content?.find((b) => b.type === "text")?.text || "[]";
  return JSON.parse(raw.replace(/```json|```/g, "").trim());
}

/* ═══════════════════════════════════════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════════════════════════════════════ */

function ApiKeyScreen({ onSave }) {
  const [val, setVal] = useState("");
  const [show, setShow] = useState(false);
  const [err, setErr] = useState(null);

  const handleSave = () => {
    const trimmed = val.trim();
    if (!trimmed.startsWith("sk-ant-")) {
      setErr("Key should start with sk-ant-  —  check you copied it correctly");
      return;
    }
    localStorage.setItem(STORAGE_KEY, trimmed);
    onSave(trimmed);
  };

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#f8f9fb", minHeight: "100vh", display: "flex", flexDirection: "column", maxWidth: 340, margin: "0 auto" }}>
      <div style={{ background: "linear-gradient(135deg,#1F5C9E,#2E86C1)", padding: "18px 16px 14px", color: "#fff" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>✦</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>Slide Cleanup</div>
            <div style={{ fontSize: 10, opacity: 0.75 }}>Setup — one time only</div>
          </div>
        </div>
      </div>
      <div style={{ flex: 1, padding: 16, display: "flex", flexDirection: "column", gap: 14 }}>
        <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "14px" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#111827", marginBottom: 6 }}>Enter your Anthropic API key</div>
          <div style={{ fontSize: 11, color: "#6b7280", lineHeight: 1.5, marginBottom: 14 }}>
            Saved only on your computer. Get yours free at{" "}
            <span style={{ color: "#1F5C9E", textDecoration: "underline", cursor: "pointer" }}
              onClick={() => window.open("https://console.anthropic.com/settings/keys", "_blank")}>
              console.anthropic.com
            </span>
          </div>
          <div style={{ position: "relative", marginBottom: 10 }}>
            <input
              type={show ? "text" : "password"}
              value={val}
              onChange={(e) => { setVal(e.target.value); setErr(null); }}
              placeholder="sk-ant-api03-…"
              style={{ width: "100%", padding: "9px 36px 9px 10px", borderRadius: 7, border: `1px solid ${err ? "#fca5a5" : "#d1d5db"}`, fontSize: 12, fontFamily: "monospace", outline: "none", boxSizing: "border-box", background: err ? "#fef2f2" : "#fff" }}
            />
            <button onClick={() => setShow(s => !s)}
              style={{ position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", fontSize: 13, color: "#9ca3af", padding: 0 }}>
              {show ? "🙈" : "👁"}
            </button>
          </div>
          {err && <div style={{ fontSize: 11, color: "#dc2626", marginBottom: 8 }}>{err}</div>}
          <button onClick={handleSave} disabled={!val.trim()}
            style={{ width: "100%", padding: "11px 0", background: "#1F5C9E", color: "#fff", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: val.trim() ? "pointer" : "not-allowed", opacity: val.trim() ? 1 : 0.5 }}>
            Save & continue →
          </button>
        </div>
        <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 8, padding: "10px 12px", fontSize: 11, color: "#92400e" }}>
          🔒 Stored in your browser only — never leaves your machine except to call Anthropic directly.
        </div>
      </div>
    </div>
  );
}

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
  const colors = Object.entries(theme.colors).filter(([,v]) => v).slice(0, 6);
  return (
    <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>
        Detected from file
      </div>
      <div style={{ display: "flex", gap: 6, marginBottom: 8, flexWrap: "wrap" }}>
        {theme.fonts.heading && (
          <span style={{ fontSize: 10, background: "#eff6ff", color: "#1e40af", padding: "2px 8px", borderRadius: 20, border: "1px solid #bfdbfe" }}>
            Aa {theme.fonts.heading}
          </span>
        )}
        {theme.fonts.body && theme.fonts.body !== theme.fonts.heading && (
          <span style={{ fontSize: 10, background: "#f5f3ff", color: "#6d28d9", padding: "2px 8px", borderRadius: 20, border: "1px solid #ddd6fe" }}>
            Aa {theme.fonts.body}
          </span>
        )}
      </div>
      <div style={{ display: "flex", gap: 4, marginBottom: 8, flexWrap: "wrap" }}>
        {colors.map(([k, v]) => (
          <div key={k} title={`${k}: ${v}`} style={{ display: "flex", alignItems: "center", gap: 3 }}>
            <div style={{ width: 14, height: 14, borderRadius: 3, background: v, border: "1px solid rgba(0,0,0,0.1)" }} />
            <span style={{ fontSize: 9, color: "#6b7280", fontFamily: "monospace" }}>{v}</span>
          </div>
        ))}
      </div>
      <div style={{ fontSize: 10, color: "#9ca3af" }}>
        {masterPlaceholders.length} master placeholder{masterPlaceholders.length !== 1 ? "s" : ""} · read from XML
      </div>
    </div>
  );
}

function MasterPicker({ masters, onSelect }) {
  const [chosen, setChosen] = useState(masters[0]?.index);

  return (
    <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "14px", animation: "fadeIn 0.3s ease" }}>
      <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 4 }}>
        Multiple masters found
      </div>
      <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 12 }}>
        This file has {masters.length} slide masters — pick the one to clean up to:
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 14 }}>
        {masters.map((m) => {
          const isChosen = chosen === m.index;
          return (
            <div
              key={m.index}
              onClick={() => setChosen(m.index)}
              style={{
                padding: "10px 12px",
                borderRadius: 8,
                border: `2px solid ${isChosen ? "#1F5C9E" : "#e5e7eb"}`,
                background: isChosen ? "#eff6ff" : "#fafafa",
                cursor: "pointer",
                transition: "all 0.15s ease",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                <div style={{
                  width: 16, height: 16, borderRadius: "50%",
                  border: `2px solid ${isChosen ? "#1F5C9E" : "#d1d5db"}`,
                  background: isChosen ? "#1F5C9E" : "#fff",
                  flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center",
                }}>
                  {isChosen && <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#fff" }} />}
                </div>
                <span style={{ fontSize: 12, fontWeight: 600, color: "#111827" }}>
                  {m.name || `Master ${m.index}`}
                </span>
              </div>

              {/* Font badges */}
              <div style={{ display: "flex", gap: 5, marginBottom: 6, paddingLeft: 24, flexWrap: "wrap" }}>
                {m.headingFont && (
                  <span style={{ fontSize: 9, background: "#eff6ff", color: "#1e40af", padding: "1px 6px", borderRadius: 20, border: "1px solid #bfdbfe" }}>
                    Aa {m.headingFont}
                  </span>
                )}
                {m.bodyFont && m.bodyFont !== m.headingFont && (
                  <span style={{ fontSize: 9, background: "#f5f3ff", color: "#6d28d9", padding: "1px 6px", borderRadius: 20, border: "1px solid #ddd6fe" }}>
                    Aa {m.bodyFont}
                  </span>
                )}
              </div>

              {/* Colour swatches */}
              {m.colors.length > 0 && (
                <div style={{ display: "flex", gap: 3, paddingLeft: 24 }}>
                  {m.colors.map((hex, i) => (
                    <div key={i} title={hex} style={{
                      width: 14, height: 14, borderRadius: 3,
                      background: hex, border: "1px solid rgba(0,0,0,0.1)",
                    }} />
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>

      <button
        onClick={() => onSelect(chosen)}
        style={{
          width: "100%", padding: "11px 0", background: "#1F5C9E", color: "#fff",
          border: "none", borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: "pointer",
        }}
      >
        Use this master →
      </button>
    </div>
  );
}

function FixBadge({ count }) {
  return (
    <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "10px 14px", display: "flex", alignItems: "center", gap: 10 }}>
      <span style={{ fontSize: 20 }}>✓</span>
      <div>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#166534" }}>Cleanup complete</div>
        <div style={{ fontSize: 11, color: "#15803d" }}>
          {count} shape{count !== 1 ? "s" : ""} reformatted · original preserved
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   MAIN APP
   ═══════════════════════════════════════════════════════════════════════════ */

export default function App() {
  const [status, setStatus] = useState("idle");
  const [log, setLog] = useState([]);
  const [fixCount, setFixCount] = useState(0);
  const [error, setError] = useState(null);
  const [detectedTheme, setDetectedTheme] = useState(null);
  const [detectedMaster, setDetectedMaster] = useState([]);
  const [apiKey, setApiKey] = useState(() => localStorage.getItem(STORAGE_KEY) || "");
  const [showSettings, setShowSettings] = useState(false);

  // Multi-master state
  const [availableMasters, setAvailableMasters] = useState([]);
  const [pendingZip, setPendingZip] = useState(null);
  const [pendingSlideIndex, setPendingSlideIndex] = useState(null);
  const [showMasterPicker, setShowMasterPicker] = useState(false);

  if (!apiKey) return <ApiKeyScreen onSave={setApiKey} />;

  if (showSettings) return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#f8f9fb", minHeight: "100vh", maxWidth: 340, margin: "0 auto" }}>
      <div style={{ background: "linear-gradient(135deg,#1F5C9E,#2E86C1)", padding: "14px 16px", color: "#fff", display: "flex", alignItems: "center", gap: 10 }}>
        <button onClick={() => setShowSettings(false)} style={{ background: "rgba(255,255,255,0.2)", border: "none", color: "#fff", borderRadius: 6, padding: "4px 10px", cursor: "pointer", fontSize: 12 }}>← Back</button>
        <span style={{ fontWeight: 700, fontSize: 14 }}>API Key Settings</span>
      </div>
      <div style={{ padding: 16 }}>
        <ApiKeyScreen onSave={(k) => { setApiKey(k); setShowSettings(false); }} />
      </div>
    </div>
  );

  const addLog = (msg) => setLog(l => [...l, {
    time: new Date().toLocaleTimeString("en-US", { hour12: false, hour: "2-digit", minute: "2-digit", second: "2-digit" }),
    msg,
  }]);

  // Phase 2: run cleanup with a chosen master
  const runCleanupWithMaster = async (zip, masters, chosenMasterIndex, slideIndex) => {
    setShowMasterPicker(false);
    setStatus("running");

    try {
      addLog(`Using master: "${masters.find(m => m.index === chosenMasterIndex)?.name || chosenMasterIndex}"`);
      addLog("Reading slide shapes from XML…");
      const pptxData = await readSlideWithMaster(zip, masters, chosenMasterIndex, slideIndex);
      setDetectedTheme(pptxData.theme);
      setDetectedMaster(pptxData.masterPlaceholders);
      addLog(`Theme: "${pptxData.theme.fonts.heading}" / "${pptxData.theme.fonts.body}"`);
      addLog(`${pptxData.masterPlaceholders.length} placeholders · ${pptxData.slideShapes.length} slide shapes`);

      addLog("Duplicating slide…");
      const dupIndex = await duplicateSlide(slideIndex);
      addLog(dupIndex === slideIndex
        ? "⚠ Running on original slide — use Ctrl+Z to undo if needed"
        : `Duplicate created at position ${dupIndex}`);

      addLog("Sending to Claude for analysis…");
      const fixes = await callClaude(pptxData, apiKey);
      addLog(`${fixes.length} fix${fixes.length !== 1 ? "es" : ""} identified`);

      if (fixes.length === 0) {
        addLog("✓ Slide already matches master theme");
        setFixCount(0);
        setStatus("done");
        return;
      }

      addLog("Applying fixes…");
      // Enrich fixes using shapeIndex (Claude returns the index from the prompt)
      const enrichedFixes = fixes.map(fix => {
        const shape = fix.shapeIndex !== undefined
          ? pptxData.slideShapes[fix.shapeIndex]
          : pptxData.slideShapes.find(s => s.name === fix.shapeName);
        return {
          ...fix,
          shapeName: shape?.name || fix.shapeName,
          shapeFill: shape?.shapeFill || null,
          shapeBorder: shape?.shapeBorder || null,
          _slideShape: shape || null,
        };
      });
      await applyFixes(dupIndex, enrichedFixes, pptxData.theme.colors);

      // Apply layout positions — use shape id for duplicate-named shapes
      const shapeNameCounts = {};
      pptxData.slideShapes.forEach(s => { shapeNameCounts[s.name] = (shapeNameCounts[s.name] || 0) + 1; });
      const themeColorSet = new Set(Object.values(pptxData.theme.colors).filter(Boolean).map(c => c.toUpperCase()));
      const positionFixes = pptxData.slideShapes
        .filter(s => s.masterTarget?.position && s.position)
        .filter(s => {
          const t = s.masterTarget.position;
          const c = s.position;
          const needsMove = Math.abs(t.left - c.left) > 0.15 ||
                 Math.abs(t.top - c.top) > 0.15 ||
                 Math.abs(t.width - c.width) > 0.15 ||
                 Math.abs(t.height - c.height) > 0.15;
          if (!needsMove) return false;
          if (s.shapeFill && s.shapeFill !== "none") {
            const isExactTheme = themeColorSet.has(s.shapeFill.replace("#","").toUpperCase());
            const fontAlreadyCorrect = s.current.fontName === s.masterTarget?.fontName ||
                                       s.current.fontName === "(inherited)";
            if (isExactTheme && fontAlreadyCorrect) return false;
          }
          return true;
        })
        .map(s => ({ shapeName: s.name, shapeId: s.id, position: s.masterTarget.position }));

      if (positionFixes.length > 0) {
        addLog(`Applying ${positionFixes.length} layout position fix(es)…`);
        await applyFixes(dupIndex, positionFixes, pptxData.theme.colors);
      }
      addLog(`✓ Done — ${fixes.length} shape${fixes.length !== 1 ? "s" : ""} reformatted`);
      setFixCount(fixes.length);
      setStatus("done");
    } catch (err) {
      setError(err.message);
      addLog("✗ " + err.message);
      setStatus("error");
    }
  };

  // Phase 1: read file, discover masters, show picker if needed
  const handleCleanup = useCallback(async () => {
    setStatus("running");
    setLog([]);
    setError(null);
    setFixCount(0);
    setDetectedTheme(null);
    setDetectedMaster([]);
    setShowMasterPicker(false);
    setAvailableMasters([]);
    setPendingZip(null);
    setPendingSlideIndex(null);

    try {
      addLog("Reading selected slide…");
      const slideIndex = await getSelectedSlideIndex();
      addLog(`Slide ${slideIndex} selected`);

      addLog("Reading .pptx file (this may take a moment)…");
      const { zip, masters } = await readPptxFile();
      addLog(`Found ${masters.length} slide master${masters.length !== 1 ? "s" : ""}`);

      if (masters.length === 0) throw new Error("No slide masters found in this file");

      if (masters.length === 1) {
        // Only one master — skip picker, go straight to cleanup
        await runCleanupWithMaster(zip, masters, masters[0].index, slideIndex);
      } else {
        // Multiple masters — show picker
        setAvailableMasters(masters);
        setPendingZip(zip);
        setPendingSlideIndex(slideIndex);
        setStatus("picking");
        setShowMasterPicker(true);
        addLog("Multiple masters detected — please choose one");
      }
    } catch (err) {
      setError(err.message);
      addLog("✗ " + err.message);
      setStatus("error");
    }
  }, [apiKey]);

  const isRunning = status === "running";

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#f8f9fb", minHeight: "100vh", display: "flex", flexDirection: "column", maxWidth: 340, margin: "0 auto" }}>
      <style>{`
        @keyframes spin { to { transform: rotate(360deg) } }
        @keyframes fadeIn { from{opacity:0;transform:translateY(5px)} to{opacity:1;transform:translateY(0)} }
        .btn:hover:not(:disabled) { background: #174f8a !important; transform: translateY(-1px); box-shadow: 0 6px 20px rgba(31,92,158,0.4) !important; }
        .btn:disabled { opacity: 0.6; cursor: not-allowed; }
      `}</style>

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg,#1F5C9E,#2E86C1)", padding: "18px 16px 14px", color: "#fff" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "rgba(255,255,255,0.2)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>✦</div>
          <div>
            <div style={{ fontWeight: 700, fontSize: 15 }}>Slide Cleanup</div>
            <div style={{ fontSize: 10, opacity: 0.75 }}>AI · Reads live from .pptx XML</div>
          </div>
          <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 10, opacity: 0.7 }}>
              {status === "idle" && "Ready"}
              {status === "running" && "Working…"}
              {status === "done" && "✓ Done"}
              {status === "error" && "✗ Error"}
            </span>
            <button onClick={() => setShowSettings(true)} title="API key settings"
              style={{ background: "rgba(255,255,255,0.15)", border: "none", color: "#fff", borderRadius: 6, width: 26, height: 26, cursor: "pointer", fontSize: 13, display: "flex", alignItems: "center", justifyContent: "center" }}>⚙</button>
          </div>
        </div>
      </div>

      <div style={{ flex: 1, padding: 16, display: "flex", flexDirection: "column", gap: 12 }}>

        {/* Master picker — shown when multiple masters detected */}
        {showMasterPicker && availableMasters.length > 1 && (
          <MasterPicker
            masters={availableMasters}
            onSelect={(chosenIndex) =>
              runCleanupWithMaster(pendingZip, availableMasters, chosenIndex, pendingSlideIndex)
            }
          />
        )}

        {/* Theme card once detected */}
        {detectedTheme && !showMasterPicker && <ThemeCard theme={detectedTheme} masterPlaceholders={detectedMaster} />}}

        {/* Info card when idle */}
        {status === "idle" && !detectedTheme && (
          <div style={{ background: "#fff", borderRadius: 10, border: "1px solid #e5e7eb", padding: "12px 14px" }}>
            <div style={{ fontSize: 10, fontWeight: 700, color: "#9ca3af", letterSpacing: "0.08em", textTransform: "uppercase", marginBottom: 8 }}>What cleanup does</div>
            {[
              ["⧉", "Duplicates the slide — original stays untouched"],
              ["◎", "Reads theme + master directly from the .pptx XML"],
              ["Aa", "Works on any file you open — no setup needed"],
              ["✦", "Reformats every shape to exactly match the master"],
            ].map(([icon, text]) => (
              <div key={text} style={{ display: "flex", gap: 8, marginBottom: 5, alignItems: "flex-start" }}>
                <span style={{ fontSize: 11, color: "#1F5C9E", flexShrink: 0, width: 16, textAlign: "center" }}>{icon}</span>
                <span style={{ fontSize: 11, color: "#374151", lineHeight: 1.4 }}>{text}</span>
              </div>
            ))}
          </div>
        )}

        {/* Button — hidden while picker is showing */}
        {!showMasterPicker && (
        <button className="btn" onClick={handleCleanup} disabled={isRunning}
          style={{ width: "100%", padding: "14px 0", background: status === "done" ? "#15803d" : "#1F5C9E", color: "#fff", border: "none", borderRadius: 10, fontSize: 14, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, transition: "all 0.2s ease", boxShadow: "0 4px 14px rgba(31,92,158,0.28)" }}>
          {isRunning ? (
            <>
              <span style={{ width: 14, height: 14, border: "2px solid rgba(255,255,255,0.3)", borderTop: "2px solid #fff", borderRadius: "50%", animation: "spin 0.8s linear infinite", display: "inline-block" }} />
              Working…
            </>
          ) : status === "done" ? "✓ Done — clean another?" : "✦ Cleanup Slide"}
        </button>
        )}

        {status === "done" && fixCount > 0 && <FixBadge count={fixCount} />}
        {status === "done" && fixCount === 0 && (
          <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 8, padding: "10px 14px", fontSize: 11, color: "#166534" }}>
            ✓ Slide already matches the master — no changes needed.
          </div>
        )}

        {error && (
          <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "10px 12px", fontSize: 11, color: "#991b1b" }}>
            <strong>Error:</strong> {error}
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
        <span style={{ fontSize: 9, color: "#9ca3af", fontFamily: "monospace" }}>v2.1.0 · claude-sonnet-4</span>
        <span style={{ fontSize: 9, color: "#9ca3af" }}>PowerPoint Add-in</span>
      </div>
    </div>
  );
}
