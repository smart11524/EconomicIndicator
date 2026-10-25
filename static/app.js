/**
 * PetroEcon Analytics - Client Controller
 * Manages file uploads, run triggers, UI updates, chart display,
 * and saving the complete output ZIP to user-selected locations.
 */

document.addEventListener("DOMContentLoaded", () => {
  // Elements
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const selectedFileInfo = document.getElementById("selectedFileInfo");
  const selectedFileName = document.getElementById("selectedFileName");
  const selectedFileSize = document.getElementById("selectedFileSize");
  const btnRemoveFile = document.getElementById("btnRemoveFile");

  const btnToggleConfig = document.getElementById("btnToggleConfig");
  const configPanel = document.getElementById("configPanel");
  const cfgDiscountRate = document.getElementById("cfgDiscountRate");
  const cfgCapitalBudget = document.getElementById("cfgCapitalBudget");
  const cfgEconomicCutoff = document.getElementById("cfgEconomicCutoff");
  const cfgPriceEscalation = document.getElementById("cfgPriceEscalation");

  const btnRunAnalysis = document.getElementById("btnRunAnalysis");
  const btnSampleRun = document.getElementById("btnSampleRun");
  const runLoader = document.getElementById("runLoader");
  const loaderStatus = document.getElementById("loaderStatus");
  const errorBanner = document.getElementById("errorBanner");
  const errorMessage = document.getElementById("errorMessage");

  const resultsSection = document.getElementById("resultsSection");
  const resultsTabs = document.getElementById("resultsTabs");

  // Export Buttons
  const btnSaveZipPicker = document.getElementById("btnSaveZipPicker");
  const btnSaveToFolderModal = document.getElementById("btnSaveToFolderModal");
  const btnDirectDownloadZip = document.getElementById("btnDirectDownloadZip");
  const btnDownloadExcel = document.getElementById("btnDownloadExcel");

  // Save Modal Elements
  const saveFolderModal = document.getElementById("saveFolderModal");
  const btnCloseSaveModal = document.getElementById("btnCloseSaveModal");
  const btnCancelSaveModal = document.getElementById("btnCancelSaveModal");
  const btnConfirmSaveToFolder = document.getElementById("btnConfirmSaveToFolder");
  const inputTargetFolder = document.getElementById("inputTargetFolder");
  const inputCustomZipName = document.getElementById("inputCustomZipName");
  const saveModalStatus = document.getElementById("saveModalStatus");
  const presetButtonsGroup = document.getElementById("presetButtonsGroup");

  let currentFile = null;
  let currentRunData = null;
  let serverPresetPaths = {};

  // Fetch server preset paths on startup
  fetch("/api/preset-paths")
    .then(r => r.json())
    .then(paths => {
      serverPresetPaths = paths;
    })
    .catch(() => {});

  // ----------------------------------------------------
  // Dropzone & File Selection Handlers
  // ----------------------------------------------------
  dropzone.addEventListener("click", (e) => {
    if (e.target !== btnRemoveFile && !btnRemoveFile.contains(e.target)) {
      fileInput.click();
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      handleFileSelected(e.target.files[0]);
    }
  });

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  btnRemoveFile.addEventListener("click", (e) => {
    e.stopPropagation();
    currentFile = null;
    fileInput.value = "";
    selectedFileInfo.style.display = "none";
    dropzone.querySelector(".dropzone-content").style.display = "block";
    btnRunAnalysis.disabled = true;
  });

  function handleFileSelected(file) {
    const validExtensions = [".xlsx", ".xls", ".xlsm"];
    const ext = "." + file.name.split(".").pop().toLowerCase();
    if (!validExtensions.includes(ext)) {
      showError(`Invalid file format '${ext}'. Please choose an Excel workbook (.xlsx, .xls, .xlsm).`);
      return;
    }

    currentFile = file;
    selectedFileName.textContent = file.name;
    selectedFileSize.textContent = formatBytes(file.size);
    selectedFileInfo.style.display = "flex";
    dropzone.querySelector(".dropzone-content").style.display = "none";
    btnRunAnalysis.disabled = false;
    hideError();
  }

  // ----------------------------------------------------
  // Configuration Accordion Toggle
  // ----------------------------------------------------
  btnToggleConfig.addEventListener("click", () => {
    btnToggleConfig.classList.toggle("active");
    if (configPanel.style.display === "block") {
      configPanel.style.display = "none";
    } else {
      configPanel.style.display = "block";
    }
  });

  // ----------------------------------------------------
  // Run Execution
  // ----------------------------------------------------
  btnRunAnalysis.addEventListener("click", () => {
    if (!currentFile) return;

    const formData = new FormData();
    formData.append("file", currentFile);

    if (cfgDiscountRate.value) formData.append("discount_rate", cfgDiscountRate.value);
    if (cfgCapitalBudget.value) formData.append("capital_budget", cfgCapitalBudget.value);
    formData.append("economic_cutoff", cfgEconomicCutoff.checked);
    formData.append("price_escalation", cfgPriceEscalation.checked);

    executeRun("/api/upload-and-run", formData);
  });

  btnSampleRun.addEventListener("click", () => {
    const payload = {
      economic_cutoff: cfgEconomicCutoff.checked,
      price_escalation: cfgPriceEscalation.checked
    };
    if (cfgDiscountRate.value) payload.discount_rate = cfgDiscountRate.value;
    if (cfgCapitalBudget.value) payload.capital_budget = cfgCapitalBudget.value;

    executeRun("/api/run-sample", JSON.stringify(payload), true);
  });

  function executeRun(endpoint, body, isJson = false) {
    hideError();
    setLoading(true, "Evaluating economic indicators...", "Parsing dataset, computing project cash flows and generating high-res figures...");

    const headers = {};
    if (isJson) {
      headers["Content-Type"] = "application/json";
    }

    fetch(endpoint, {
      method: "POST",
      headers: headers,
      body: body
    })
      .then(async (res) => {
        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error || "An error occurred during evaluation.");
        }
        return data;
      })
      .then((data) => {
        currentRunData = data;
        renderResults(data);
        setLoading(false);
        resultsSection.style.display = "block";
        resultsSection.scrollIntoView({ behavior: "smooth" });
      })
      .catch((err) => {
        setLoading(false);
        showError(err.message);
      });
  }

  function setLoading(loading, statusText, subText) {
    if (loading) {
      runLoader.style.display = "flex";
      loaderStatus.textContent = statusText || "Processing...";
      loaderSubtext.textContent = subText || "Please wait while calculations run.";
      btnRunAnalysis.disabled = true;
      btnSampleRun.disabled = true;
    } else {
      runLoader.style.display = "none";
      btnRunAnalysis.disabled = !currentFile;
      btnSampleRun.disabled = false;
    }
  }

  function showError(msg) {
    errorMessage.textContent = msg;
    errorBanner.style.display = "flex";
    errorBanner.scrollIntoView({ behavior: "smooth" });
  }

  function hideError() {
    errorBanner.style.display = "none";
  }

  // ----------------------------------------------------
  // Render Results to UI
  // ----------------------------------------------------
  function renderResults(data) {
    const p = data.project;
    const bm = data.base_metrics;

    // 1. Executive Banner
    document.getElementById("resProjectName").textContent = `${p.name} Field Development Study`;
    document.getElementById("resProjectMeta").textContent = 
      `${p.life} Years Life | ${(p.discount_rate * 100).toFixed(1)}% Discount Rate | ${(p.tax_rate * 100).toFixed(0)}% Tax | Depr: ${p.dep_name} (${p.dep_years}y)`;

    document.getElementById("resRecommendation").textContent = data.recommendation;
    const isBudgeted = data.final_comparison && data.final_comparison.some(c => c["Within budget"] === "NO");
    document.getElementById("resRecommendationReason").textContent = isBudgeted
      ? `Optimal case constrained by capital budget, confirmed through incremental (Delta) analysis.`
      : `NPV-maximizing, incremental-analysis-confirmed development alternative for maximum enterprise value creation.`;

    // 2. Export Toolbar Links
    btnDirectDownloadZip.href = `/api/download-zip/${data.run_id}`;
    btnDownloadExcel.href = `/api/download-results-excel/${data.run_id}`;
    inputCustomZipName.value = data.zip_filename;

    // 3. KPI Grid
    document.getElementById("kpiNpv").textContent = `$${bm.npv.toFixed(2)} MM`;
    const kpiNpvStatus = document.getElementById("kpiNpvStatus");
    if (bm.npv > 0) {
      kpiNpvStatus.textContent = "✓ Value Accretive (Accept)";
      kpiNpvStatus.style.color = "var(--accent-emerald)";
    } else {
      kpiNpvStatus.textContent = "✗ Value Destructive (Reject)";
      kpiNpvStatus.style.color = "var(--accent-rose)";
    }

    document.getElementById("kpiIrr").textContent = bm.irr_text;
    document.getElementById("kpiMirr").textContent = `MIRR: ${(bm.mirr * 100).toFixed(2)}%`;

    document.getElementById("kpiPayback").textContent = bm.payback !== null ? `${bm.payback.toFixed(2)} yrs` : "Never";
    document.getElementById("kpiDiscPayback").textContent = bm.disc_payback !== null ? `Disc. Payback: ${bm.disc_payback.toFixed(2)} yrs` : "Disc. Payback: Never";

    document.getElementById("kpiPir").textContent = bm.pir.toFixed(3);
    document.getElementById("kpiPi").textContent = `PI: ${bm.pi.toFixed(3)} | Undisc: ${bm.pir_undiscounted.toFixed(2)}`;

    document.getElementById("kpiCapex").textContent = `$${bm.total_capex.toFixed(2)} MM`;
    document.getElementById("kpiWells").textContent = `${p.wells} Wells Initial`;

    document.getElementById("kpiWinner").textContent = data.recommendation;
    const winnerRow = data.alternatives.find(a => a.Alternative === data.recommendation);
    document.getElementById("kpiWinnerNpv").textContent = winnerRow ? `NPV: $${winnerRow["NPV ($MM)"].toFixed(2)} MM` : "";

    // 4. Screening Table
    const tbodyScreening = document.querySelector("#tableScreening tbody");
    tbodyScreening.innerHTML = "";
    data.screening_table.forEach(row => {
      const isOk = row["OK?"] === "OK";
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${row.Indicator}</strong></td>
        <td class="cell-numeric">${row.Value}</td>
        <td>${row.Criterion}</td>
        <td class="cell-badge"><span class="badge ${isOk ? 'badge-success' : 'badge-danger'}">${row["OK?"]}</span></td>
      `;
      tbodyScreening.appendChild(tr);
    });

    // 5. Capex Breakdown Table
    const tbodyCapex = document.querySelector("#tableCapex tbody");
    tbodyCapex.innerHTML = "";
    for (const [k, v] of Object.entries(data.capex_breakdown)) {
      const tr = document.createElement("tr");
      const isTotal = k.includes("TOTAL");
      tr.innerHTML = `
        <td>${isTotal ? `<strong>${k}</strong>` : k}</td>
        <td class="cell-numeric">${isTotal ? `<strong>$${v.toFixed(2)}</strong>` : `$${v.toFixed(2)}`}</td>
      `;
      tbodyCapex.appendChild(tr);
    }

    // 6. Workbook Verification Table
    const verCol = document.getElementById("verificationCol");
    const tbodyVer = document.querySelector("#tableVerification tbody");
    tbodyVer.innerHTML = "";
    if (data.verification && data.verification.length > 0) {
      verCol.style.display = "block";
      data.verification.forEach(row => {
        const tr = document.createElement("tr");
        const match = row.Match === "YES";
        tr.innerHTML = `
          <td><strong>${row.Indicator}</strong></td>
          <td class="cell-numeric">${typeof row.Python === 'number' ? row.Python.toFixed(3) : row.Python}</td>
          <td class="cell-numeric">${typeof row.Workbook === 'number' ? row.Workbook.toFixed(3) : row.Workbook}</td>
          <td class="cell-numeric">${typeof row.Difference === 'number' ? (row.Difference > 0 ? '+' : '') + row.Difference.toFixed(4) : '-'}</td>
          <td class="cell-badge"><span class="badge ${match ? 'badge-success' : 'badge-danger'}">${row.Match}</span></td>
        `;
        tbodyVer.appendChild(tr);
      });
    } else {
      verCol.style.display = "none";
    }

    // 7. Cash Flow Table
    const tbodyCf = document.querySelector("#tableCashFlow tbody");
    tbodyCf.innerHTML = "";
    data.cash_flow_preview.forEach(row => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${row.Year}</td>
        <td>${row.ProjYear}</td>
        <td class="cell-numeric">${row.Oil_MMB.toFixed(3)}</td>
        <td class="cell-numeric">${row.Gas_BCF.toFixed(3)}</td>
        <td class="cell-numeric">${row.Revenue.toFixed(3)}</td>
        <td class="cell-numeric">${row.CAPEX > 0 ? row.CAPEX.toFixed(3) : '-'}</td>
        <td class="cell-numeric">${row.OPEX.toFixed(3)}</td>
        <td class="cell-numeric">${row.Depreciation.toFixed(3)}</td>
        <td class="cell-numeric">${row.Tax.toFixed(3)}</td>
        <td class="cell-numeric" style="color: ${row.NCF >= 0 ? 'var(--accent-emerald)' : 'var(--accent-rose)'}; font-weight: 600;">${row.NCF.toFixed(3)}</td>
        <td class="cell-numeric" style="font-weight: 600;">${row.Cum_NCF.toFixed(3)}</td>
      `;
      tbodyCf.appendChild(tr);
    });

    // 8. Alternatives Table
    const tbodyAlts = document.querySelector("#tableAlternatives tbody");
    tbodyAlts.innerHTML = "";
    data.alternatives.forEach(row => {
      const tr = document.createElement("tr");
      const isWinner = row.Alternative === data.recommendation;
      if (isWinner) tr.style.backgroundColor = "rgba(6, 182, 212, 0.1)";
      tr.innerHTML = `
        <td><strong>${row.Alternative}</strong> ${isWinner ? '<span class="badge badge-accent">Selected</span>' : ''}</td>
        <td class="cell-numeric">${row.Wells}</td>
        <td class="cell-numeric">$${row["Capex ($MM)"].toFixed(2)}</td>
        <td class="cell-numeric" style="color: ${row["NPV ($MM)"] > 0 ? 'var(--accent-emerald)' : 'var(--accent-rose)'}; font-weight: 700;">$${row["NPV ($MM)"].toFixed(2)}</td>
        <td class="cell-numeric">${row.IRR}</td>
        <td class="cell-numeric">${typeof row.MIRR === 'number' ? (row.MIRR * 100).toFixed(1) + '%' : row.MIRR}</td>
        <td class="cell-numeric">${row["Payback (yr)"] !== null ? row["Payback (yr)"].toFixed(2) : '-'}</td>
        <td class="cell-numeric">${row.PI.toFixed(2)}</td>
        <td class="cell-numeric">${row.PIR.toFixed(2)}</td>
      `;
      tbodyAlts.appendChild(tr);
    });

    // 9. Incremental Analysis Table
    const tbodyInc = document.querySelector("#tableIncremental tbody");
    tbodyInc.innerHTML = "";
    data.incremental_analysis.forEach(row => {
      const isTake = row.Decision.includes("TAKE");
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${row.Increment}</strong></td>
        <td class="cell-numeric">$${row["Delta Capex ($MM)"].toFixed(2)}</td>
        <td class="cell-numeric" style="color: ${row["Delta NPV ($MM)"] > 0 ? 'var(--accent-emerald)' : 'var(--accent-rose)'}; font-weight: 700;">$${row["Delta NPV ($MM)"].toFixed(2)}</td>
        <td class="cell-numeric">${row["Delta IRR"]}</td>
        <td class="cell-numeric">${typeof row["Delta MIRR"] === 'number' ? (row["Delta MIRR"] * 100).toFixed(1) + '%' : row["Delta MIRR"]}</td>
        <td class="cell-numeric">${isNaN(row["Delta PIR"]) ? '-' : row["Delta PIR"].toFixed(2)}</td>
        <td class="cell-badge"><span class="badge ${isTake ? 'badge-success' : 'badge-danger'}">${row.Decision}</span></td>
      `;
      tbodyInc.appendChild(tr);
    });

    // 10. Sensitivity Table
    const tbodySens = document.querySelector("#tableSensitivity tbody");
    tbodySens.innerHTML = "";
    data.sensitivity.forEach(row => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${row.Variable}</strong></td>
        <td>${row["Low case"]}</td>
        <td class="cell-numeric">$${row["NPV @ low"].toFixed(2)}</td>
        <td>${row["High case"]}</td>
        <td class="cell-numeric">$${row["NPV @ high"].toFixed(2)}</td>
        <td class="cell-numeric" style="color: var(--accent-amber); font-weight: 700;">$${row["Range ($MM)"].toFixed(2)}</td>
      `;
      tbodySens.appendChild(tr);
    });

    // 11. Startup Delay Table
    const tbodyStartup = document.querySelector("#tableStartupDelay tbody");
    tbodyStartup.innerHTML = "";
    data.startup_delay.forEach(row => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><strong>${row["Delay (years)"]} yr</strong></td>
        <td class="cell-numeric" style="color: ${row["NPV ($MM)"] > 0 ? 'var(--accent-emerald)' : 'var(--accent-rose)'}; font-weight: 700;">$${row["NPV ($MM)"].toFixed(2)}</td>
        <td class="cell-numeric">${row.IRR}</td>
        <td class="cell-numeric">${typeof row.MIRR === 'number' ? (row.MIRR * 100).toFixed(1) + '%' : row.MIRR}</td>
        <td class="cell-numeric">${row["Payback from today (yr)"] !== null ? row["Payback from today (yr)"].toFixed(2) : '-'}</td>
        <td class="cell-numeric">${row["Payback from start-up (yr)"] !== null ? row["Payback from start-up (yr)"].toFixed(2) : '-'}</td>
        <td class="cell-numeric">${row.PIR.toFixed(3)}</td>
      `;
      tbodyStartup.appendChild(tr);
    });

    // 12. Load Embedded Figures
    const timestamp = Date.now();
    const figUrl = (name) => `/api/figures/${data.run_id}/${name}?t=${timestamp}`;

    document.getElementById("figDiscountAndPv").src = figUrl("fig_01_discount_and_pv_flows.png");
    document.getElementById("figNpvProfile").src = figUrl("fig_03_npv_profile.png");
    document.getElementById("figPaybackCurves").src = figUrl("fig_04_payback_curves.png");
    document.getElementById("figNpvComponents").src = figUrl("fig_02_npv_component_buildup.png");
    document.getElementById("figAlternatives").src = figUrl("fig_06_alternatives_comparison.png");
    document.getElementById("figSensitivity").src = figUrl("fig_07_sensitivity_spider_tornado.png");
    document.getElementById("figStartupDelay").src = figUrl("fig_05_startup_delay.png");

    // 13. Populate Gallery Grid
    const galleryContainer = document.getElementById("galleryContainer");
    galleryContainer.innerHTML = "";
    const figNames = [
      { id: "fig_01_discount_and_pv_flows.png", title: "Discount Factors & PV Cash Flows" },
      { id: "fig_02_npv_component_buildup.png", title: "NPV Component Build-up" },
      { id: "fig_03_npv_profile.png", title: "NPV Profile & IRR Roots" },
      { id: "fig_04_payback_curves.png", title: "Cumulative Payback Curves" },
      { id: "fig_05_startup_delay.png", title: "Startup Delay Valuation Risk" },
      { id: "fig_06_alternatives_comparison.png", title: "Alternatives & Capital Efficiency" },
      { id: "fig_07_sensitivity_spider_tornado.png", title: "Sensitivity Spider & Tornado Charts" }
    ];

    figNames.forEach(item => {
      const url = figUrl(item.id);
      const card = document.createElement("div");
      card.className = "gallery-item";
      card.innerHTML = `
        <a href="${url}" target="_blank" title="View Full High-Resolution Image">
          <img src="${url}" alt="${item.title}" class="chart-img">
        </a>
        <div class="gallery-item-body">
          <span class="gallery-item-title">${item.title}</span>
          <a href="${url}" download="${item.id}" class="btn btn-outline text-sm">Download PNG</a>
        </div>
      `;
      galleryContainer.appendChild(card);
    });

    // 14. Populate Execution Log
    document.getElementById("runLogViewer").textContent = data.log || "No log generated.";
  }

  // ----------------------------------------------------
  // Tabs Navigation
  // ----------------------------------------------------
  resultsTabs.addEventListener("click", (e) => {
    if (e.target.classList.contains("tab-btn")) {
      const targetId = e.target.getAttribute("data-tab");
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

      e.target.classList.add("active");
      const targetContent = document.getElementById(targetId);
      if (targetContent) targetContent.classList.add("active");
    }
  });

  // ----------------------------------------------------
  // SAVE ZIP TO USER-SELECTED LOCATION (File System Access API)
  // ----------------------------------------------------
  btnSaveZipPicker.addEventListener("click", async () => {
    if (!currentRunData) return;

    const zipFilename = currentRunData.zip_filename;
    const downloadUrl = `/api/download-zip/${currentRunData.run_id}`;

    // Modern Chrome / Edge native "Save As" file picker
    if ("showSaveFilePicker" in window) {
      try {
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: zipFilename,
          types: [{
            description: "ZIP Archive (.zip)",
            accept: { "application/zip": [".zip"] }
          }]
        });

        // Fetch blob and write to selected handle
        const response = await fetch(downloadUrl);
        const blob = await response.blob();
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();

        alert(`✓ Output ZIP successfully saved to your selected computer location:\n${fileHandle.name}`);
        return;
      } catch (err) {
        if (err.name === "AbortError") {
          // User clicked Cancel in dialog
          return;
        }
        console.warn("showSaveFilePicker failed or unsupported, falling back to modal...", err);
      }
    }

    // Fallback: Open directory input modal
    openSaveModal();
  });

  // ----------------------------------------------------
  // SAVE TO LOCAL DIRECTORY MODAL
  // ----------------------------------------------------
  btnSaveToFolderModal.addEventListener("click", () => {
    openSaveModal();
  });

  function openSaveModal() {
    if (!currentRunData) return;
    saveFolderModal.style.display = "flex";
    saveModalStatus.style.display = "none";
    inputCustomZipName.value = currentRunData.zip_filename;

    // Set default folder to Desktop or Project Folder if empty
    if (!inputTargetFolder.value) {
      inputTargetFolder.value = serverPresetPaths["Desktop"] || serverPresetPaths["Project Folder"] || "";
    }
  }

  function closeSaveModal() {
    saveFolderModal.style.display = "none";
  }

  btnCloseSaveModal.addEventListener("click", closeSaveModal);
  btnCancelSaveModal.addEventListener("click", closeSaveModal);

  // Preset button clicks
  presetButtonsGroup.addEventListener("click", (e) => {
    const btn = e.target.closest(".btn-preset");
    if (!btn) return;
    const presetKey = btn.getAttribute("data-path");
    if (serverPresetPaths[presetKey]) {
      inputTargetFolder.value = serverPresetPaths[presetKey];
    }
  });

  // Confirm save to local folder
  btnConfirmSaveToFolder.addEventListener("click", () => {
    if (!currentRunData) return;
    const targetDir = inputTargetFolder.value.trim();
    const customName = inputCustomZipName.value.trim();

    if (!targetDir) {
      showModalStatus("Please enter or select a destination folder path.", "error");
      return;
    }

    btnConfirmSaveToFolder.disabled = true;
    btnConfirmSaveToFolder.textContent = "Saving...";

    fetch("/api/save-zip-to-folder", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        run_id: currentRunData.run_id,
        target_dir: targetDir,
        custom_name: customName
      })
    })
      .then(async (r) => {
        const res = await r.json();
        if (!r.ok) throw new Error(res.error || "Failed to save ZIP to folder.");
        return res;
      })
      .then((data) => {
        showModalStatus(`✓ Saved successfully!\nLocation: ${data.saved_path}\nSize: ${formatBytes(data.size_bytes)}`, "success");
        btnConfirmSaveToFolder.disabled = false;
        btnConfirmSaveToFolder.textContent = "Save ZIP to this Location";
      })
      .catch((err) => {
        showModalStatus(`Error: ${err.message}`, "error");
        btnConfirmSaveToFolder.disabled = false;
        btnConfirmSaveToFolder.textContent = "Save ZIP to this Location";
      });
  });

  function showModalStatus(text, type) {
    saveModalStatus.textContent = text;
    saveModalStatus.className = `modal-alert ${type}`;
    saveModalStatus.style.display = "block";
  }

  // Helper: Format Bytes
  function formatBytes(bytes) {
    if (!bytes || bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  }
});
