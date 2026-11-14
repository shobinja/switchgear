document.addEventListener('DOMContentLoaded', () => {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  const fileInput = document.getElementById('excelFile');
  const dropArea = document.getElementById('dropArea');
  const fileLabel = document.getElementById('fileLabel');
  const processBtn = document.getElementById('processBtn');
  const logDiv = document.getElementById('log');

  const checkboxes = {
    mv: document.getElementById('chkMv'),
    ct: document.getElementById('chkCt'),
    bbDlro: document.getElementById('chkBbDlro'),
    bbIr: document.getElementById('chkBbIr'),
    vt: document.getElementById('chkVt')
  };

  const labels = {
    mv: document.getElementById('lblMv'),
    ct: document.getElementById('lblCt'),
    bbDlro: document.getElementById('lblBbDlro'),
    bbIr: document.getElementById('lblBbIr'),
    vt: document.getElementById('lblVt')
  };

  let selectedFile = null;
  let parsedWorkbook = null;

  function updateButtonState() {
    const hasSelection = Object.values(checkboxes).some(cb => cb.checked && !cb.disabled);
    processBtn.disabled = !selectedFile || !hasSelection;
  }

  Object.values(checkboxes).forEach(cb => {
    cb.addEventListener('change', updateButtonState);
  });

  async function handleFile(f) {
    selectedFile = f;
    fileLabel.innerText = "Selected: " + f.name;
    logDiv.style.display = 'none';
    logDiv.textContent = '';

    try {
      const data = await selectedFile.arrayBuffer();
      parsedWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

      // Enable/disable based on available sheets, but keep all unchecked by default
      const sheetAvailability = {
        mv: !!parsedWorkbook.Sheets['MV_Panel'],
        ct: !!parsedWorkbook.Sheets['CT'],
        bbDlro: !!parsedWorkbook.Sheets['BB_CR'],
        bbIr: !!parsedWorkbook.Sheets['BB_IR'],
        vt: !!parsedWorkbook.Sheets['VT']
      };

      for (const key in checkboxes) {
        const available = sheetAvailability[key];
        checkboxes[key].checked = false; // Always unchecked by default
        checkboxes[key].disabled = !available;
        labels[key].classList.toggle('disabled', !available);
        labels[key].title = available ? "" : `Required sheet for this report is missing in workbook`;
      }

      updateButtonState();
    } catch (e) {
      log("Error reading workbook: " + e.message);
    }
  }

  dropArea.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
  });
  dropArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropArea.style.borderColor = '#2563eb';
  });
  dropArea.addEventListener('dragleave', () => {
    dropArea.style.borderColor = '#cbd5e1';
  });
  dropArea.addEventListener('drop', (e) => {
    e.preventDefault();
    dropArea.style.borderColor = '#cbd5e1';
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
  });

  function log(msg) {
    logDiv.style.display = 'block';
    logDiv.textContent += msg + '\n';
    logDiv.scrollTop = logDiv.scrollHeight;
  }

  processBtn.addEventListener('click', async () => {
    if (!selectedFile || !parsedWorkbook) return;

    processBtn.disabled = true;
    logDiv.textContent = '';
    log("Validating sheets and processing...");

    try {
      const zipBlob = await generateAllReports(
        parsedWorkbook,
        {
          runMv: checkboxes.mv.checked,
          runCt: checkboxes.ct.checked,
          runBbDlro: checkboxes.bbDlro.checked,
          runBbIr: checkboxes.bbIr.checked,
          runVt: checkboxes.vt.checked
        },
        log
      );

      saveAs(zipBlob, "Switchgear_Reports.zip");
      log("Done! Reports downloaded successfully.");
    } catch (err) {
      log("Error: " + err.message);
      console.error(err);
    } finally {
      processBtn.disabled = false;
    }
  });
});