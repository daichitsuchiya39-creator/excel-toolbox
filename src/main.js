const { invoke } = window.__TAURI__.core;
const { open, save } = window.__TAURI__.dialog;

const state = {
  filePath: "",
  sheets: [],
};

const ui = {};

function setStatus(message, tone = "muted") {
  ui.result.textContent = message;
  ui.result.dataset.tone = tone;
}

function updateFileMeta() {
  if (!state.filePath) {
    ui.selectedFile.textContent = "未選択";
    ui.fileStatus.textContent = "ファイルを選択してください";
    ui.runExtract.disabled = true;
    return;
  }
  ui.selectedFile.textContent = state.filePath;
  ui.fileStatus.textContent = "シート一覧を取得しました";
  ui.runExtract.disabled = false;
}

function renderSheets() {
  ui.sheetList.innerHTML = "";
  ui.sheetCount.textContent = `${state.sheets.length} 件`;

  if (state.sheets.length === 0) {
    const empty = document.createElement("div");
    empty.className = "sheet-empty";
    empty.textContent = "シートが見つかりませんでした";
    ui.sheetList.appendChild(empty);
    return;
  }

  state.sheets.forEach((name) => {
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

async function handleSelectFile() {
  setStatus("ファイルを選択しています...", "muted");
  const filePath = await open({
    multiple: false,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });

  if (!filePath) {
    setStatus("ファイル選択をキャンセルしました。", "muted");
    return;
  }

  state.filePath = filePath;
  setStatus("シートを読み込み中...", "muted");

  try {
    const sheets = await invoke("load_sheets", { path: state.filePath });
    state.sheets = sheets;
    renderSheets();
    updateFileMeta();
    setStatus("準備完了。抽出方法を選んで実行できます。", "success");
  } catch (error) {
    setStatus(`エラー: ${error}`, "error");
  }
}

function getSelectedSheets() {
  return Array.from(ui.sheetList.querySelectorAll("input[type='checkbox']"))
    .filter((input) => input.checked)
    .map((input) => input.value);
}

async function handleExtract() {
  if (!state.filePath) {
    setStatus("先にExcelファイルを選択してください。", "error");
    return;
  }

  const mode = document.querySelector("input[name='mode']:checked").value;
  const fileName = state.filePath.split(/[/\\\\]/).pop() || "sheetpic.xlsx";
  const baseName = fileName.replace(/\.xlsx$/i, "");

  const defaultName =
    mode === "keyword"
      ? `抽出_${ui.keywordInput.value || "keyword"}.xlsx`
      : `抽出_${baseName}.xlsx`;

  const outputPath = await save({
    defaultPath: defaultName,
    filters: [{ name: "Excel", extensions: ["xlsx"] }],
  });

  if (!outputPath) {
    setStatus("保存先の指定をキャンセルしました。", "muted");
    return;
  }

  setStatus("抽出を実行中...", "muted");

  try {
    let count = 0;
    if (mode === "keyword") {
      const keyword = ui.keywordInput.value.trim();
      if (!keyword) {
        setStatus("キーワードを入力してください。", "error");
        return;
      }
      count = await invoke("extract_by_keyword", {
        path: state.filePath,
        keyword,
        outputPath,
      });
    } else {
      const selected = getSelectedSheets();
      if (selected.length === 0) {
        setStatus("抽出するシートを選択してください。", "error");
        return;
      }
      count = await invoke("extract_by_selection", {
        path: state.filePath,
        sheets: selected,
        outputPath,
      });
    }

    setStatus(`完了: ${count} 件のシートを抽出しました。`, "success");
  } catch (error) {
    setStatus(`エラー: ${error}`, "error");
  }
}

window.addEventListener("DOMContentLoaded", () => {
  ui.selectFile = document.querySelector("#select-file");
  ui.selectedFile = document.querySelector("#selected-file");
  ui.fileStatus = document.querySelector("#file-status");
  ui.modeInputs = document.querySelectorAll("input[name='mode']");
  ui.keywordArea = document.querySelector("#keyword-area");
  ui.keywordInput = document.querySelector("#keyword-input");
  ui.selectionArea = document.querySelector("#selection-area");
  ui.sheetList = document.querySelector("#sheet-list");
  ui.sheetCount = document.querySelector("#sheet-count");
  ui.runExtract = document.querySelector("#run-extract");
  ui.result = document.querySelector("#result");

  updateMode();
  ui.selectFile.addEventListener("click", handleSelectFile);
  ui.runExtract.addEventListener("click", handleExtract);
  ui.modeInputs.forEach((input) => input.addEventListener("change", updateMode));
});
