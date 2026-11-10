document.addEventListener('DOMContentLoaded', () => {
  // Service Worker for offline PWA functionality
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => {});
  }

  const fileInput = document.getElementById('excelFile');
  const dropArea = document.getElementById('dropArea');
  const fileLabel = document.getElementById('fileLabel');
  const processBtn = document.getElementById('processBtn');
  const logDiv = document.getElementById('log');

  const chkMv = document.getElementById('chkMv');
  const chkCt = document.getElementById('chkCt');
  const lblMv = document.getElementById('lblMv');
  const lblCt = document.getElementById('lblCt');

  let selectedFile = null;
  let parsedWorkbook = null;

  function updateButtonState() {
    const hasSelection = (chkMv.checked && !chkMv.disabled) || (chkCt.checked && !chkCt.disabled);
    processBtn.disabled = !selectedFile || !hasSelection;
  }

  chkMv.addEventListener('change', updateButtonState);
  chkCt.addEventListener('change', updateButtonState);

  async function handleFile(f) {
    selectedFile = f;
    fileLabel.innerText = "Selected: " + f.name;
    logDiv.style.display = 'none';
    logDiv.textContent = '';

    try {
      const data = await selectedFile.arrayBuffer();
      parsedWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

      const hasMv = !!parsedWorkbook.Sheets['MV_Panel'];
      const hasCt = !!parsedWorkbook.Sheets['CT'];

      chkMv.checked = hasMv;
      chkMv.disabled = !hasMv;
      lblMv.classList.toggle('disabled', !hasMv);
      lblMv.title = hasMv ? "" : "MV_Panel sheet not found in workbook";

      chkCt.checked = hasCt;
      chkCt.disabled = !hasCt;
      lblCt.classList.toggle('disabled', !hasCt);
      lblCt.title = hasCt ? "" : "CT sheet not found in workbook";

      updateButtonState();
    } catch (e) {
      log("Error reading workbook structure: " + e.message);
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
    log("Processing reports...");

    try {
      const zipBlob = await generateAllReports(
        parsedWorkbook,
        chkMv.checked,
        chkCt.checked,
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