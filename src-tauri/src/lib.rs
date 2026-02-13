use std::collections::{HashMap, HashSet};
use std::path::Path;

use umya_spreadsheet::{reader::xlsx as xlsx_reader, writer::xlsx as xlsx_writer, Spreadsheet};

fn get_sheet_names(book: &Spreadsheet) -> Vec<String> {
    book.get_sheet_collection_no_check()
        .iter()
        .map(|sheet| sheet.get_name().to_string())
        .collect()
}

fn get_base_filename(path: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("workbook")
        .to_string()
}

fn take_chars(input: &str, count: usize) -> String {
    input.chars().take(count).collect()
}

fn char_len(input: &str) -> usize {
    input.chars().count()
}

fn normalize_sheet_name(name: &str) -> String {
    if char_len(name) > 31 {
        let mut trimmed = take_chars(name, 28);
        trimmed.push_str("...");
        trimmed
    } else {
        name.to_string()
    }
}

fn unique_sheet_name(base: &str, counts: &mut HashMap<String, usize>) -> String {
    let mut name = normalize_sheet_name(base);
    if let Some(count) = counts.get_mut(&name) {
        *count += 1;
        let suffix = format!("_{}", count);
        let max_base = 31usize.saturating_sub(suffix.len());
        let mut base_trim = name.clone();
        if char_len(&base_trim) > max_base {
            base_trim = take_chars(&base_trim, max_base);
        }
        name = format!("{}{}", base_trim, suffix);
    } else {
        counts.insert(name.clone(), 0);
    }
    name
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

#[tauri::command]
fn merge_workbooks(paths: Vec<String>, output_path: String) -> Result<usize, String> {
    if paths.len() < 2 {
        return Err("2つ以上のファイルを選択してください".to_string());
    }

    let mut merged_book = umya_spreadsheet::new_file_empty_worksheet();
    let mut sheet_name_counts: HashMap<String, usize> = HashMap::new();
    let mut total_sheets = 0usize;

    for path in paths {
        let book = xlsx_reader::read(&path).map_err(|e| e.to_string())?;
        let base_name = get_base_filename(&path);

        for sheet in book.get_sheet_collection_no_check() {
            let original_name = sheet.get_name();
            let base_sheet = format!("{}_{}", base_name, original_name);
            let new_name = unique_sheet_name(&base_sheet, &mut sheet_name_counts);

            let mut cloned = sheet.clone();
            cloned.set_name(new_name);
            merged_book
                .add_sheet(cloned)
                .map_err(|e| e.to_string())?;
            total_sheets += 1;
        }
    }

    if total_sheets == 0 {
        return Err("マージ対象のシートが見つかりませんでした".to_string());
    }

    xlsx_writer::write(&merged_book, &output_path).map_err(|e| e.to_string())?;
    Ok(total_sheets)
}

#[tauri::command]
fn remove_macro(path: String, output_path: String) -> Result<(), String> {
    let mut book = xlsx_reader::read(&path).map_err(|e| e.to_string())?;
    book.remove_macros_code();
    xlsx_writer::write(&book, &output_path).map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            load_sheets,
            extract_by_keyword,
            extract_by_selection,
            merge_workbooks,
            remove_macro
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
