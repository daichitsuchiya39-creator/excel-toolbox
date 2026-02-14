const { invoke } = window.__TAURI__.core;
const { open, save } = window.__TAURI__.dialog;

const state = {
  pick: {
    filePath: "",
    sheets: [],
  },
  merge: {
    paths: [],
  },
  macro: {
    filePath: "",
  },
};

const ui = {};

function setStatus(element, message, tone = "muted") {
  element.textContent = message;
  element.dataset.tone = tone;
}

function updatePickMeta() {
  if (!state.pick.filePath) {
    ui.pickSelectedFile.textContent = "Not Selected";
    ui.pickFileStatus.textContent = "Please select a file";
    ui.pickRunExtract.disabled = true;
    return;
  }
  ui.pickSelectedFile.textContent = state.pick.filePath;
  ui.pickFileStatus.textContent = "Sheet list loaded";
  ui.pickRunExtract.disabled = false;
}

function updateMergeMeta() {
  if (state.merge.paths.length < 2) {
    ui.mergeSelectedFiles.textContent = "Not Selected";
    ui.mergeFileStatus.textContent = "Please select 2 or more files";
    ui.mergeRun.disabled = true;
    return;
  }
  const display = state.merge.paths.length <= 3
    ? state.merge.paths.join(", ")
    : `${state.merge.paths.length} files selected`;
  ui.mergeSelectedFiles.textContent = display;
  ui.mergeFileStatus.textContent = "Ready to merge";
  ui.mergeRun.disabled = false;
}

function updateMacroMeta() {
  if (!state.macro.filePath) {
    ui.macroSelectedFile.textContent = "Not Selected";
    ui.macroFileStatus.textContent = "Please select an .xlsm file";
    ui.macroRun.disabled = true;
    return;
  }
  ui.macroSelectedFile.textContent = state.macro.filePath;
  ui.macroFileStatus.textContent = "Ready to remove macros";
  ui.macroRun.disabled = false;
}

function renderSheets() {
  ui.sheetList.innerHTML = "";
  ui.sheetCount.textContent = `${state.pick.sheets.length} sheets`;

  if (state.pick.sheets.length === 0) {
    const empty = document.createElement("div");
    empty.className = "sheet-empty";
    empty.textContent = "No sheets found";
    ui.sheetList.appendChild(empty);
    return;
  }

  state.pick.sheets.forEach((name) => {
    const label = document.createElement("label");
    label.className = "sheet-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = name;

    const span = document.createElement("span");
    span.textContent = name;

    label.appendChild(checkbox);
    label.appendChild(span);
    ui.sheetList.appendChild(label);
  });
}

function updateMode() {
  const mode = document.querySelector("input[name='mode']:checked").value;
  const isKeyword = mode === "keyword";
  ui.keywordArea.classList.toggle("hidden", !isKeyword);
  ui.selectionArea.classList.toggle("hidden", isKeyword);
}

function setActiveTool(tool) {
  ui.toolTabs.forEach((tab) => {
    const isActive = tab.dataset.tool === tool;
    tab.classList.toggle("is-active", isActive);
    tab.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  ui.toolPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.toolPanel === tool);
  });
}

async function handleSelectFile() {
  setStatus(ui.pickResult, "Selecting file...", "muted");
  const filePath = await open({
    multiple: false,
    filters: [{ name: "Excel", extensions: ["xlsx", "xlsm"] }],
  });

  if (!filePath) {
    setStatus(ui.pickResult, "File selection cancelled.", "muted");
    return;
  }

  state.pick.filePath = filePath;
  setStatus(ui.pickResult, "Loading sheets...", "muted");

  try {
    const sheets = await invoke("load_sheets", { path: state.pick.filePath });
    state.pick.sheets = sheets;
    renderSheets();
    updatePickMeta();
    setStatus(ui.pickResult, "Ready. Choose extraction method and run.", "success");
  } catch (error) {
    setStatus(ui.pickResult, `Error: ${error}`, "error");
  }
}

function getSelectedSheets() {
  return Array.from(ui.sheetList.querySelectorAll("input[type='checkbox']"))
    .filter((input) => input.checked)
    .map((input) => input.value);
}

async function handleExtract() {
  if (!state.pick.filePath) {
    setStatus(ui.pickResult, "Please select an Excel file first.", "error");
    return;
  }

  const mode = document.querySelector("input[name='mode']:checked").value;
  const fileName = state.pick.filePath.split(/[/\\\\]/).pop() || "sheetpic.xlsx";
  const baseName = fileName.replace(/\.(xlsx|xlsm)$/i, "");

  const defaultName =
    mode === "keyword"
      ? `extracted_${ui.keywordInput.value || "keyword"}.xlsx`
      : `extracted_${baseName}.xlsx`;

  const outputPath = await save({
    defaultPath: defaultName,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });

  if (!outputPath) {
    setStatus(ui.pickResult, "Save location cancelled.", "muted");
    return;
  }

  setStatus(ui.pickResult, "Extracting...", "muted");

  try {
    let count = 0;
    if (mode === "keyword") {
      const keyword = ui.keywordInput.value.trim();
      if (!keyword) {
        setStatus(ui.pickResult, "Please enter a keyword.", "error");
        return;
      }
      count = await invoke("extract_by_keyword", {
        path: state.pick.filePath,
        keyword,
        outputPath,
      });
    } else {
      const selected = getSelectedSheets();
      if (selected.length === 0) {
        setStatus(ui.pickResult, "Please select sheets to extract.", "error");
        return;
      }
      count = await invoke("extract_by_selection", {
        path: state.pick.filePath,
        sheets: selected,
        outputPath,
      });
    }

    setStatus(ui.pickResult, `Complete: ${count} sheet(s) extracted.`, "success");
  } catch (error) {
    setStatus(ui.pickResult, `Error: ${error}`, "error");
  }
}

async function handleMergeSelectFiles() {
  setStatus(ui.mergeResult, "Selecting files...", "muted");
  const paths = await open({
    multiple: true,
    filters: [{ name: "Excel", extensions: ["xlsx", "xlsm"] }],
  });

  if (!paths || paths.length === 0) {
    setStatus(ui.mergeResult, "File selection cancelled.", "muted");
    return;
  }

  state.merge.paths = paths;
  updateMergeMeta();
  setStatus(ui.mergeResult, "Ready to merge. Specify save location.", "success");
}

async function handleMergeRun() {
  if (state.merge.paths.length < 2) {
    setStatus(ui.mergeResult, "Please select 2 or more files.", "error");
    return;
  }

  const outputPath = await save({
    defaultPath: "merged_file.xlsx",
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });

  if (!outputPath) {
    setStatus(ui.mergeResult, "Save location cancelled.", "muted");
    return;
  }

  setStatus(ui.mergeResult, "Merging...", "muted");

  try {
    const totalSheets = await invoke("merge_workbooks", {
      paths: state.merge.paths,
      outputPath,
    });
    setStatus(
      ui.mergeResult,
      `Complete: ${totalSheets} sheet(s) merged.`,
      "success"
    );
  } catch (error) {
    setStatus(ui.mergeResult, `Error: ${error}`, "error");
  }
}

async function handleMacroSelectFile() {
  setStatus(ui.macroResult, "Selecting file...", "muted");
  const filePath = await open({
    multiple: false,
    filters: [{ name: "Excel Macro", extensions: ["xlsm"] }],
  });

  if (!filePath) {
    setStatus(ui.macroResult, "File selection cancelled.", "muted");
    return;
  }

  state.macro.filePath = filePath;
  updateMacroMeta();
  setStatus(ui.macroResult, "Ready. Specify save location and run.", "success");
}

async function handleMacroRun() {
  if (!state.macro.filePath) {
    setStatus(ui.macroResult, "Please select an .xlsm file.", "error");
    return;
  }

  const fileName =
    state.macro.filePath.split(/[/\\\\]/).pop() || "macro.xlsm";
  const baseName = fileName.replace(/\.xlsm$/i, "");
  const outputPath = await save({
    defaultPath: `${baseName}(macros_removed).xlsx`,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });

  if (!outputPath) {
    setStatus(ui.macroResult, "Save location cancelled.", "muted");
    return;
  }

  setStatus(ui.macroResult, "Removing macros...", "muted");

  try {
    await invoke("remove_macro", {
      path: state.macro.filePath,
      outputPath,
    });
    setStatus(ui.macroResult, "Complete: Saved as .xlsx.", "success");
  } catch (error) {
    setStatus(ui.macroResult, `Error: ${error}`, "error");
  }
}

window.addEventListener("DOMContentLoaded", () => {
  ui.toolTabs = Array.from(document.querySelectorAll(".tool-tab"));
  ui.toolPanels = Array.from(document.querySelectorAll(".tool-panel"));
  ui.pickSelectFile = document.querySelector("#pick-select-file");
  ui.pickSelectedFile = document.querySelector("#pick-selected-file");
  ui.pickFileStatus = document.querySelector("#pick-file-status");
  ui.modeInputs = document.querySelectorAll("input[name='mode']");
  ui.keywordArea = document.querySelector("#keyword-area");
  ui.keywordInput = document.querySelector("#keyword-input");
  ui.selectionArea = document.querySelector("#selection-area");
  ui.sheetList = document.querySelector("#sheet-list");
  ui.sheetCount = document.querySelector("#sheet-count");
  ui.pickRunExtract = document.querySelector("#pick-run-extract");
  ui.pickResult = document.querySelector("#pick-result");

  ui.mergeSelectFiles = document.querySelector("#merge-select-files");
  ui.mergeSelectedFiles = document.querySelector("#merge-selected-files");
  ui.mergeFileStatus = document.querySelector("#merge-file-status");
  ui.mergeRun = document.querySelector("#merge-run");
  ui.mergeResult = document.querySelector("#merge-result");

  ui.macroSelectFile = document.querySelector("#macro-select-file");
  ui.macroSelectedFile = document.querySelector("#macro-selected-file");
  ui.macroFileStatus = document.querySelector("#macro-file-status");
  ui.macroRun = document.querySelector("#macro-run");
  ui.macroResult = document.querySelector("#macro-result");

  updateMode();
  updatePickMeta();
  updateMergeMeta();
  updateMacroMeta();
  setActiveTool("pick");

  ui.toolTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      setActiveTool(tab.dataset.tool);
    });
  });

  ui.pickSelectFile.addEventListener("click", handleSelectFile);
  ui.pickRunExtract.addEventListener("click", handleExtract);
  ui.modeInputs.forEach((input) => input.addEventListener("change", updateMode));

  ui.mergeSelectFiles.addEventListener("click", handleMergeSelectFiles);
  ui.mergeRun.addEventListener("click", handleMergeRun);

  ui.macroSelectFile.addEventListener("click", handleMacroSelectFile);
  ui.macroRun.addEventListener("click", handleMacroRun);
});
