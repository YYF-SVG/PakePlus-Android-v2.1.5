use tauri::{Manager, Window};
use std::fs;
use std::path::PathBuf;
use base64::{engine::general_purpose::STANDARD, Engine as _};
use chrono::Local;

// 命令：保存Excel文件
#[tauri::command]
async fn save_excel_file(window: Window, data: String) -> Result<(), String> {
    // 使用原生对话框让用户选择保存位置
    let file_path = match tauri::api::dialog::FileDialogBuilder::new()
        .set_title("保存Excel文件")
        .add_filter("Excel文件", &["xlsx", "xls"])
        .set_file_name(&format!("车辆费用记录_{}.xlsx", Local::now().format("%Y%m%d")))
        .save_file()
    {
        Some(path) => path,
        None => return Err("用户取消了保存".into()),
    };

    // 解码Base64数据
    let file_data = match STANDARD.decode(&data) {
        Ok(data) => data,
        Err(e) => return Err(format!("数据解码失败: {}", e)),
    };

    // 写入文件
    match fs::write(&file_path, file_data) {
        Ok(_) => Ok(()),
        Err(e) => Err(format!("文件写入失败: {}", e)),
    }
}

// 命令：打开并读取Excel文件
#[tauri::command]
async fn open_excel_file(window: Window) -> Result<String, String> {
    // 使用原生对话框让用户选择文件
    let file_path = match tauri::api::dialog::FileDialogBuilder::new()
        .set_title("选择Excel文件")
        .add_filter("Excel文件", &["xlsx", "xls"])
        .pick_file()
    {
        Some(path) => path,
        None => return Err("用户取消了选择".into()),
    };

    // 读取文件
    let file_data = match fs::read(&file_path) {
        Ok(data) => data,
        Err(e) => return Err(format!("文件读取失败: {}", e)),
    };

    // 编码为Base64返回
    Ok(STANDARD.encode(&file_data))
}

// 主函数
fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![save_excel_file, open_excel_file])
        .run(tauri::generate_context!())
        .expect("运行Tauri应用时出错");
}