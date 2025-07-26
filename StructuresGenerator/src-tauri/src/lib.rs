use calamine::Reader;
use serde::Serialize;
use std::process::Command;
use tauri::{command, AppHandle};

#[derive(Serialize)]
struct FileDialogResult {
    canceled: bool,
    file_path: Option<String>,
    sheet_names: Option<Vec<String>>,
    error: Option<String>,
}

#[command]
fn select_excel_file(_app: AppHandle) -> FileDialogResult {
    let dialog = rfd::FileDialog::new()
        .add_filter("Excel Files", &["xlsx", "xls"])
        .pick_file();

    match dialog {
        Some(path_buf) => {
            let path = path_buf.to_string_lossy().to_string();

            // Leemos el contenido usando crate `calamine`
            match calamine::open_workbook_auto(&path) {
                Ok(workbook) => {
                    let sheet_names: Vec<String> = workbook.sheet_names().to_vec();
                    FileDialogResult {
                        canceled: false,
                        file_path: Some(path),
                        sheet_names: Some(sheet_names),
                        error: None,
                    }
                }
                Err(e) => FileDialogResult {
                    canceled: false,
                    file_path: Some(path),
                    sheet_names: None,
                    error: Some(format!("Error al leer Excel: {}", e)),
                },
            }
        }
        None => FileDialogResult {
            canceled: true,
            file_path: None,
            sheet_names: None,
            error: None,
        },
    }
}

#[command]
fn process_excel(file_path: String, sheet_name: String) -> Result<serde_json::Value, String> {
    println!("📥 Ejecutando script Python con:");
    println!("  ➤ file_path = {}", file_path);
    println!("  ➤ sheet_name = {}", sheet_name);

    let output = Command::new("python")
        .arg("../main.py")
        .arg(file_path)
        .arg(sheet_name)
        .output()
        .map_err(|e| e.to_string())?;

    // if !output.status.success() {
    //     return Err(String::from_utf8_lossy(&output.stderr).to_string());
    // }

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        println!("❌ STDERR:\n{}", stderr);
        return Err(format!("Error del script Python:\n{}", stderr));
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    println!("✅ STDOUT:\n{}", stdout);

    serde_json::from_str(&stdout).map_err(|e| e.to_string())
}

#[tauri::command]
fn abrir_carpeta_con_archivo(_path: String) -> Result<(), String> {
    // Hardcodeamos aquí la ruta del archivo generado
    let path_str = "C:\\Users\\Quique\\Documents\\github\\StructuresGenerator_Tauri\\StructuresGenerator\\generated";
    let ruta = std::path::Path::new(path_str);

    if !ruta.exists() {
        return Err(format!("El archivo no existe en: {}", path_str));
    }

    #[cfg(target_os = "windows")]
    {
        use std::process::Command;
        Command::new("explorer")
            .arg(ruta)
            .spawn()
            .map_err(|e| format!("Error al abrir el explorador: {}", e))?;
    }

    Ok(())
}


#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![select_excel_file, process_excel, abrir_carpeta_con_archivo])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
