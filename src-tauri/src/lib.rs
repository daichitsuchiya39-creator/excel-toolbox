use std::collections::HashSet;

use umya_spreadsheet::{reader::xlsx as xlsx_reader, writer::xlsx as xlsx_writer, Spreadsheet};

fn get_sheet_names(book: &Spreadsheet) -> Vec<String> {
    book.get_sheet_collection_no_check()
        .iter()
        .map(|sheet| sheet.get_name().to_string())
        .collect()
}

fn write_selected_sheets(
    source_path: &str,
    selected: &[String],
    output_path: &str,
) -> Result<usize, String> {
    let mut book = xlsx_reader::read(source_path).map_err(|e| e.to_string())?;
    let all_names = get_sheet_names(&book);
    let selected_set: HashSet<String> = selected.iter().cloned().collect();
    let selected_existing: Vec<String> = all_names
        .iter()
        .filter(|name| selected_set.contains(*name))
        .cloned()
        .collect();

    if selected_existing.is_empty() {
        return Err("対象のシートが見つかりませんでした".to_string());
    }

    for name in all_names {
        if !selected_set.contains(&name) {
            book.remove_sheet_by_name(&name)
                .map_err(|e| e.to_string())?;
        }
    }

    xlsx_writer::write(&book, output_path).map_err(|e| e.to_string())?;
    Ok(selected_existing.len())
}

#[tauri::command]
fn load_sheets(path: String) -> Result<Vec<String>, String> {
    let book = xlsx_reader::read(&path).map_err(|e| e.to_string())?;
    Ok(get_sheet_names(&book))
}

#[tauri::command]
fn extract_by_keyword(path: String, keyword: String, output_path: String) -> Result<usize, String> {
    let book = xlsx_reader::read(&path).map_err(|e| e.to_string())?;
    let selected: Vec<String> = get_sheet_names(&book)
        .into_iter()
        .filter(|name| name.contains(&keyword))
        .collect();

    if selected.is_empty() {
        return Err("該当するシートがありませんでした".to_string());
    }

    write_selected_sheets(&path, &selected, &output_path)
}

#[tauri::command]
fn extract_by_selection(
    path: String,
    sheets: Vec<String>,
    output_path: String,
) -> Result<usize, String> {
    if sheets.is_empty() {
        return Err("シートが選択されていません".to_string());
    }

    write_selected_sheets(&path, &sheets, &output_path)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            load_sheets,
            extract_by_keyword,
            extract_by_selection
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
