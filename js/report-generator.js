// Function to fetch Word templates (.docx) with cache-busting
async function loadTemplate(url) {
  const res = await fetch(url + '?t=' + Date.now());
  if (!res.ok) {
    throw new Error("Template not found: " + url + " (Status: " + res.status + ")");
  }
  return await res.arrayBuffer();
}

// Convert Excel dates, booleans, and values
function formatVal(v) {
  if (v === undefined || v === null) return "";

  // 1. JavaScript Date object handling
  if (v instanceof Date && !isNaN(v.getTime())) {
    const d = String(v.getDate()).padStart(2, '0');
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const y = v.getFullYear();
    return `${d}/${m}/${y}`;
  }

  const s = String(v).trim();
  if (s.toLowerCase() === "true") return "OK";
  if (s.toLowerCase() === "false") return "N/A";

  // 2. Excel numeric serial date fallback
  if (/^\d{5}$/.test(s)) {
    const num = parseInt(s, 10);
    if (num > 30000 && num < 60000) {
      const parsedDate = new Date((num - 25569) * 86400 * 1000);
      const d = String(parsedDate.getDate()).padStart(2, '0');
      const m = String(parsedDate.getMonth() + 1).padStart(2, '0');
      const y = parsedDate.getFullYear();
      return `${d}/${m}/${y}`;
    }
  }

  return s;
}

function randRange(min, max) {
  if (min > max) {
    const temp = min;
    min = max;
    max = temp;
  }
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function rawTagParser(tag) {
  const key = tag.trim();
  return {
    get: function(scope) {
      if (key === '.') return scope;
      return (scope && scope[key] !== undefined && scope[key] !== null) ? scope[key] : '';
    }
  };
}

function renderDocx(templateBuffer, tags) {
  let zip;
  try {
    zip = new window.PizZip(templateBuffer);
  } catch (e) {
    throw new Error("Unable to parse Word template: " + e.message);
  }

  try {
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: '{{', end: '}}' },
      parser: rawTagParser,
      nullGetter: () => ""
    });
    doc.render(tags);
    return doc.getZip().generate({ type: 'uint8array' });
  } catch (err) {
    if (err.properties && err.properties.errors) {
      const details = err.properties.errors.map((e, i) => {
        const t = e.properties?.tag ? ` [Tag: ${e.properties.tag}]` : '';
        const msg = e.properties?.explanation || e.message;
        return `#${i + 1}:${t} ${msg}`;
      }).join('\n');
      throw new Error("Template Syntax Error:\n" + details);
    }
    throw err;
  }
}

// CT Document Generation
function buildSingleCTDoc(folder, templateBytes, generalTags, rowTags, ctPrmTag, iCP, rowNameA, onLog) {
  const tags = Object.assign({}, generalTags, rowTags);

  const ctP = tags['CT_P'] || '';
  const ctP2 = tags['CT_P2'] || '';
  const ctS = tags['CT_S'] || '1';
  const noc = parseInt(tags['noc'] || '1', 10);

  tags['CTR'] = ctP ? `${ctP}/${ctS}` : '';
  if (noc === 1) {
    tags['CTR2'] = '-';
  } else if (noc === 2) {
    tags['CTR2'] = ctP2 ? `${ctP2}/${ctS}` : '-';
  } else {
    tags['CTR2'] = '-';
  }

  let rngVal = parseInt(tags['range'] || '100', 10);
  if (isNaN(rngVal) || rngVal <= 0) rngVal = 100;

  const centers = [2500, 5000, 7500, 10000];
  const prefixes = ["Rp", "Yp", "Bp"];

  for (let j = 0; j < 4; j++) {
    const base = centers[j];
    const minVal = base - rngVal;
    const maxVal = base + Math.floor(rngVal / 2);

    for (let p = 0; p < 3; p++) {
      let strVal = "";
      let attempts = 0;
      do {
        const rnd = randRange(minVal, maxVal) / 10000;
        strVal = rnd.toFixed(3);
        attempts++;
      } while (strVal.endsWith("0") && attempts < 50);

      if (strVal.endsWith("0")) {
        strVal = strVal.slice(0, -1) + String(randRange(1, 9));
      }

      tags[`${prefixes[p]}${j + 1}`] = strVal;
    }
  }

  const ctPrm = parseInt(tags[ctPrmTag] || '0', 10);
  if (!isNaN(ctPrm) && ctPrm > 0) {
    tags["CT1"] = String(Math.floor(ctPrm * 0.25));
    tags["CT2"] = String(Math.floor(ctPrm * 0.5));
    tags["CT3"] = String(Math.floor(ctPrm * 0.75));
    tags["CT4"] = String(ctPrm);
  } else {
    tags["CT1"] = "-"; tags["CT2"] = "-"; tags["CT3"] = "-"; tags["CT4"] = "-";
  }

  const prps = (tags['prps'] || '').trim().toUpperCase();
  const LP = noc * 2;
  let CP = 1;

  if (prps === "METERING") {
    CP = 1;
    if (iCP) CP = CP + 1;
    tags["CT11"] = tags["CT1"]; tags["CT21"] = tags["CT2"]; tags["CT31"] = tags["CT3"]; tags["CT41"] = tags["CT4"];
    tags["Rp11"] = "-"; tags["Rp21"] = "-"; tags["Rp31"] = "-"; tags["Rp41"] = "-";
    tags["Yp11"] = "-"; tags["Yp21"] = "-"; tags["Yp31"] = "-"; tags["Yp41"] = "-";
    tags["Bp11"] = "-"; tags["Bp21"] = "-"; tags["Bp31"] = "-"; tags["Bp41"] = "-";

  } else if (prps === "PROTECTION") {
    CP = noc + 1;
    if (iCP) CP = CP + 1;
    tags["CT11"] = "-"; tags["CT21"] = "-"; tags["CT31"] = "-"; tags["CT41"] = "-";
    tags["Rp11"] = tags["Rp1"]; tags["Rp21"] = tags["Rp2"]; tags["Rp31"] = tags["Rp3"]; tags["Rp41"] = tags["Rp4"];
    tags["Yp11"] = tags["Yp1"]; tags["Yp21"] = tags["Yp2"]; tags["Yp31"] = tags["Yp3"]; tags["Yp41"] = tags["Yp4"];
    tags["Bp11"] = tags["Bp1"]; tags["Bp21"] = tags["Bp2"]; tags["Bp31"] = tags["Bp3"]; tags["Bp41"] = tags["Bp4"];
  }

  tags["sht"] = `${CP} OF ${LP}`;

  const outBytes = renderDocx(templateBytes, tags);
  const shtNo = String(tags["sht"]).replace(/\s+/g, '');
  const cleanName = rowNameA.replace(/\s+/g, '_');
  const filename = `${cleanName}_${shtNo}.docx`;

  folder.file(filename, outBytes);
  if (onLog) onLog(`- Created CT/${filename}`);
}

// Master Orchestration Function
async function generateAllReports(parsedWorkbook, options, onLog) {
  // Pre-flight validation
  if (!parsedWorkbook.Sheets['General']) {
    throw new Error("Missing required 'General' sheet. Operation stopped.");
  }

  if (options.runMv && !parsedWorkbook.Sheets['MV_Panel']) {
    throw new Error("Missing 'MV_Panel' sheet required for MV Panel Report.");
  }
  if (options.runCt && !parsedWorkbook.Sheets['CT']) {
    throw new Error("Missing 'CT' sheet required for CT Report.");
  }
  if (options.runBbDlro && !parsedWorkbook.Sheets['BB_CR']) {
    throw new Error("Missing 'BB_CR' sheet required for Busbar DLRO Report.");
  }
  if (options.runBbIr && !parsedWorkbook.Sheets['BB_IR']) {
    throw new Error("Missing 'BB_IR' sheet required for Busbar IR & HiPot Report.");
  }
  if (options.runVt && !parsedWorkbook.Sheets['VT']) {
    throw new Error("Missing 'VT' sheet required for VT Report.");
  }

  const zipOut = new JSZip();

  // 1. General Tags
  const generalTags = {};
  const genData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['General'], { header: 1 });
  if (genData.length >= 3) {
    const tagRow = genData[1];
    const valRow = genData[2];
    for (let c = 0; c < tagRow.length; c++) {
      const tag = String(tagRow[c] || '').trim();
      if (tag) generalTags[tag] = formatVal(valRow[c]);
    }
  }

  // 2. MV_Panel Processing
  if (options.runMv) {
    onLog("Loading MVPanel.docx template...");
    const mvTemplate = await loadTemplate('./Templates/MVPanel.docx');
    onLog("Generating MV_Panel reports...");

    const mvData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['MV_Panel'], { header: 1 });
    if (mvData.length >= 3) {
      const tagRow = mvData[1];
      const folder = zipOut.folder("MV_Panel");

      for (let r = 2; r < mvData.length; r++) {
        const row = mvData[r];
        if (!row || row.length === 0) continue;
        const rowName = String(row[1] || row[0] || '').trim();
        if (!rowName || rowName.toLowerCase() === 'cub') continue;

        const tags = Object.assign({}, generalTags);
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) tags[tag] = formatVal(row[c]);
        }

        const outBytes = renderDocx(mvTemplate, tags);
        const fileName = rowName.replace(/\s+/g, '_');
        folder.file(`${fileName}.docx`, outBytes);
        onLog(`- Created MV_Panel/${fileName}.docx`);
      }
    }
  }

  // 3. CT Processing
  if (options.runCt) {
    onLog("Loading Ct1.docx template...");
    const ctTemplate = await loadTemplate('./Templates/Ct1.docx');
    onLog("Generating CT reports...");

    const ctData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['CT'], { header: 1 });
    if (ctData.length >= 3) {
      const tagRow = ctData[1];
      const folder = zipOut.folder("CT");

      for (let r = 2; r < ctData.length; r++) {
        const row = ctData[r];
        if (!row || row.length === 0) continue;
        const rowName = String(row[1] || row[0] || '').trim();
        if (!rowName || rowName.toLowerCase() === 'cub') continue;

        const baseRowTags = {};
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) baseRowTags[tag] = formatVal(row[c]);
        }

        if (!baseRowTags['prps']) continue;
        const noc = parseInt(baseRowTags['noc'] || '1', 10);

        buildSingleCTDoc(folder, ctTemplate, generalTags, baseRowTags, "CT_P", false, rowName, onLog);
        if (noc === 2) {
          buildSingleCTDoc(folder, ctTemplate, generalTags, baseRowTags, "CT_P2", true, rowName, onLog);
        }
      }
    }
  }

  // 4. Busbar DLRO (BB_CR sheet -> Templates/BbDlro.docx -> BUSBAR folder)
  if (options.runBbDlro) {
    onLog("Loading BbDlro.docx template...");
    const dlroTemplate = await loadTemplate('./Templates/BbDlro.docx');
    onLog("Generating Busbar DLRO reports...");

    const dlroData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['BB_CR'], { header: 1 });
    if (dlroData.length >= 3) {
      const tagRow = dlroData[1];
      const folder = zipOut.folder("BUSBAR");

      for (let r = 2; r < dlroData.length; r++) {
        const row = dlroData[r];
        if (!row || row.length === 0) continue;
        const rowName = String(row[1] || row[0] || '').trim();
        if (!rowName) continue;

        const tags = Object.assign({}, generalTags);
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) {
            const raw = formatVal(row[c]);
            tags[tag] = raw === "" ? "-" : raw;
          }
        }

        // Get default rating from sheet/general tags (e.g. 400)
        const ratingVal = tags['r_i'] || '';

        // Dynamically check j1 to j11; blank out all fields from the first empty joint onward
        let blankRemaining = false;
        for (let j = 1; j <= 11; j++) {
          const jKey = `j${j}`;
          const val = tags[jKey];

          if (blankRemaining || !val || val === "-" || val.trim() === "") {
            blankRemaining = true;
            tags[`j${j}`] = "";
            tags[`r_i${j}`] = "";
            tags[`r${j}`] = "";
            tags[`y${j}`] = "";
            tags[`b${j}`] = "";
          } else {
            // Assign specific row rating if available, or fall back to the general rating
            tags[`r_i${j}`] = tags[`r_i${j}`] || ratingVal;
          }
        }

        const outBytes = renderDocx(dlroTemplate, tags);
        const fileName = rowName.replace(/\s+/g, '_');
        folder.file(`${fileName}_DLRO.docx`, outBytes);
        onLog(`- Created BUSBAR/${fileName}_DLRO.docx`);
      }
    }
  }

  // 5. Busbar IR & HiPot (BB_IR sheet -> Templates/BbIr.docx -> BUSBAR folder)
  if (options.runBbIr) {
    onLog("Loading BbIr.docx template...");
    const irTemplate = await loadTemplate('./Templates/BbIr.docx');
    onLog("Generating Busbar IR & HiPot reports...");

    const irData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['BB_IR'], { header: 1 });
    if (irData.length >= 3) {
      const tagRow = irData[1];
      const folder = zipOut.folder("BUSBAR");

      for (let r = 2; r < irData.length; r++) {
        const row = irData[r];
        if (!row || row.length === 0) continue;
        const rowName = String(row[1] || row[0] || '').trim();
        if (!rowName) continue;

        const tags = Object.assign({}, generalTags);
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) {
            const raw = formatVal(row[c]);
            tags[tag] = raw === "" ? "-" : raw;
          }
        }

        const outBytes = renderDocx(irTemplate, tags);
        const fileName = rowName.replace(/\s+/g, '_');
        folder.file(`${fileName}_IR.docx`, outBytes);
        onLog(`- Created BUSBAR/${fileName}_IR.docx`);
      }
    }
  }

  // 6. VT Report (VT sheet -> Templates/Vt.docx -> VT folder)
  if (options.runVt) {
    onLog("Loading Vt.docx template...");
    const vtTemplate = await loadTemplate('./Templates/Vt.docx');
    onLog("Generating VT reports...");

    const vtData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['VT'], { header: 1 });
    if (vtData.length >= 3) {
      const tagRow = vtData[1];
      const folder = zipOut.folder("VT");

      for (let r = 2; r < vtData.length; r++) {
        const row = vtData[r];
        if (!row || row.length === 0) continue;
        const rowName = String(row[1] || row[0] || '').trim();
        if (!rowName) continue;

        const tags = Object.assign({}, generalTags);
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) {
            const raw = formatVal(row[c]);
            tags[tag] = raw === "" ? "-" : raw;
          }
        }

        const outBytes = renderDocx(vtTemplate, tags);
        const fileName = rowName.replace(/\s+/g, '_');
        folder.file(`${fileName}.docx`, outBytes);
        onLog(`- Created VT/${fileName}.docx`);
      }
    }
  }

  onLog("Bundling files into ZIP...");
  return await zipOut.generateAsync({ type: 'blob' });
}