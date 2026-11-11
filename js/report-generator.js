// Fetch template buffer with cache-busting
async function loadTemplate(url) {
  const res = await fetch(url + '?t=' + Date.now());
  if (!res.ok) throw new Error("Template not found: " + url + " (Status: " + res.status + ")");
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

  // 2. Excel serial number date handling
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
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// Treat tags as raw dictionary keys (allows hyphens, underscores)
function rawTagParser(tag) {
  const key = tag.trim();
  return {
    get: function(scope) {
      if (key === '.') return scope;
      return (scope && scope[key] !== undefined && scope[key] !== null) ? scope[key] : '';
    }
  };
}

// Render docx with double braces {{tag}}
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

// Generate single CT docx
function buildSingleCTDoc(folder, templateBytes, generalTags, rowTags, ctPrmTag, iCP, row, onLog) {
  const tags = Object.assign({}, generalTags, rowTags);

  const rngVal = parseInt(tags['range'] || '0', 10);
  const centers = [2500, 5000, 7500, 10000];
  const prefixes = ["Rp", "Yp", "Bp"];

  for (let j = 0; j < 4; j++) {
    const base = centers[j];
    const minVal = base - rngVal;
    const maxVal = base + Math.floor(rngVal / 2);

    for (let p = 0; p < 3; p++) {
      let strVal = "";
      do {
        const rnd = randRange(minVal, maxVal) / 10000;
        strVal = rnd.toFixed(3);
      } while (strVal.endsWith("0"));

      tags[`${prefixes[p]}${j + 1}`] = strVal;
    }
  }

  const ctPrm = parseInt(tags[ctPrmTag] || '0', 10);
  if (!isNaN(ctPrm)) {
    tags["CT1"] = String(Math.floor(ctPrm * 0.25));
    tags["CT2"] = String(Math.floor(ctPrm * 0.5));
    tags["CT3"] = String(Math.floor(ctPrm * 0.75));
    tags["CT4"] = String(ctPrm);
  }

  const prps = tags['prps'];
  const noc = parseInt(tags['noc'] || '1', 10);
  const LP = noc * 2;
  let CP = 1;

  if (prps === "METERING") {
    CP = 1;
    if (iCP) CP += 1;
    tags["CT11"] = tags["CT1"]; tags["CT21"] = tags["CT2"]; tags["CT31"] = tags["CT3"]; tags["CT41"] = tags["CT4"];
    tags["Rp11"] = "-"; tags["Rp21"] = "-"; tags["Rp31"] = "-"; tags["Rp41"] = "-";
    tags["Yp11"] = "-"; tags["Yp21"] = "-"; tags["Yp31"] = "-"; tags["Yp41"] = "-";
    tags["Bp11"] = "-"; tags["Bp21"] = "-"; tags["Bp31"] = "-"; tags["Bp41"] = "-";
  } else if (prps === "PROTECTION") {
    CP = noc + 1;
    if (iCP) CP += 1;
    tags["CT11"] = "-"; tags["CT21"] = "-"; tags["CT31"] = "-"; tags["CT41"] = "-";
    tags["Rp11"] = tags["Rp1"]; tags["Rp21"] = tags["Rp2"]; tags["Rp31"] = tags["Rp3"]; tags["Rp41"] = tags["Rp4"];
    tags["Yp11"] = tags["Yp1"]; tags["Yp21"] = tags["Yp2"]; tags["Yp31"] = tags["Yp3"]; tags["Yp41"] = tags["Yp4"];
    tags["Bp11"] = tags["Bp1"]; tags["Bp21"] = tags["Bp2"]; tags["Bp31"] = tags["Bp3"]; tags["Bp41"] = tags["Bp4"];
  }

  tags["sht"] = `${CP} OF ${LP}`;

  const outBytes = renderDocx(templateBytes, tags);
  const nameA = String(row[1] || row[0]).replace(/\s+/g, '_');
  const shtNo = String(tags["sht"]).replace(/\s+/g, '');
  const filename = `${nameA}_${shtNo}.docx`;

  folder.file(filename, outBytes);
  if (onLog) onLog(`- Created CT/${filename}`);
}

// Orchestrator function
async function generateAllReports(parsedWorkbook, runMv, runCt, onLog) {
  const zipOut = new JSZip();

  // 1. General Tags
  const generalTags = {};
  if (parsedWorkbook.Sheets['General']) {
    const genData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['General'], { header: 1 });
    if (genData.length >= 3) {
      const tagRow = genData[1];
      const valRow = genData[2];
      for (let c = 0; c < tagRow.length; c++) {
        const tag = String(tagRow[c] || '').trim();
        if (tag) generalTags[tag] = formatVal(valRow[c]);
      }
    }
  }

  // 2. MV_Panel
  if (runMv && parsedWorkbook.Sheets['MV_Panel']) {
    onLog("Loading MVPanel.docx template...");
    const mvTemplate = await loadTemplate('./Templates/MVPanel.docx');
    onLog("Generating MV_Panel reports...");

    const mvData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['MV_Panel'], { header: 1 });
    if (mvData.length >= 3) {
      const tagRow = mvData[1];
      const folder = zipOut.folder("MV_Panel");

      for (let r = 2; r < mvData.length; r++) {
        const row = mvData[r];
        if (!row || row.length === 0 || !row[0]) continue;

        const tags = Object.assign({}, generalTags);
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) tags[tag] = formatVal(row[c]);
        }

        const outBytes = renderDocx(mvTemplate, tags);
        const nameA = String(row[1] || row[0]).replace(/\s+/g, '_');
        folder.file(`${nameA}.docx`, outBytes);
        onLog(`- Created MV_Panel/${nameA}.docx`);
      }
    }
  }

  // 3. CT
  if (runCt && parsedWorkbook.Sheets['CT']) {
    onLog("Loading Ct1.docx template...");
    const ctTemplate = await loadTemplate('./Templates/Ct1.docx');
    onLog("Generating CT reports...");

    const ctData = XLSX.utils.sheet_to_json(parsedWorkbook.Sheets['CT'], { header: 1 });
    if (ctData.length >= 3) {
      const tagRow = ctData[1];
      const folder = zipOut.folder("CT");

      for (let r = 2; r < ctData.length; r++) {
        const row = ctData[r];
        if (!row || row.length === 0 || !row[0]) continue;

        const baseRowTags = {};
        for (let c = 0; c < tagRow.length; c++) {
          const tag = String(tagRow[c] || '').trim();
          if (tag) baseRowTags[tag] = formatVal(row[c]);
        }

        const noc = parseInt(baseRowTags['noc'] || '1', 10);

        buildSingleCTDoc(folder, ctTemplate, generalTags, baseRowTags, "CT_P", false, row, onLog);

        if (noc === 2) {
          buildSingleCTDoc(folder, ctTemplate, generalTags, baseRowTags, "CT_P2", true, row, onLog);
        }
      }
    }
  }

  onLog("Bundling files into ZIP...");
  return await zipOut.generateAsync({ type: 'blob' });
}