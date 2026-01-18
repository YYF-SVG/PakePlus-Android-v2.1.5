use tauri::{Manager, Window};
use std::fs;
use std::path::PathBuf;
use base64::{engine::general_purpose::STANDARD, Engine as _};
use calamine::{Reader, Xlsx, open_workbook, DataType};
use serde::{Deserialize, Serialize};
use serde_json::json;
use chrono::Local;

// 错误类型
#[derive(Debug, thiserror::Error)]
enum AppError {
    #[error("用户取消操作")]
    UserCanceled,
    #[error("文件错误: {0}")]
    FileError(String),
    #[error("Excel解析错误: {0}")]
    ExcelError(String),
    #[error("未知错误: {0}")]
    Unknown(String),
}

// 转换为前端可用的错误
impl serde::Serialize for AppError {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: serde::ser::Serializer,
    {
        serializer.serialize_str(self.to_string().as_ref())
    }
}

// 导入Excel命令
#[tauri::command]
async fn import_excel_file(window: Window) -> Result<serde_json::Value, AppError> {
    // 1. 打开文件选择对话框
    let file_path = tauri::api::dialog::FileDialogBuilder::new()
        .set_title("选择Excel文件")
        .add_filter("Excel文件", &["xlsx", "xls"])
        .pick_file()
        .ok_or(AppError::UserCanceled)?;

    println!("开始解析文件: {:?}", file_path);

    // 2. 使用calamine打开Excel文件
    let mut excel: Xlsx<_> = open_workbook(&file_path)
        .map_err(|e| AppError::ExcelError(format!("无法打开Excel文件: {}", e)))?;

    // 3. 准备结果
    let mut charging_data = Vec::new();
    let mut parking_data = Vec::new();

    // 4. 检查所有工作表
    for sheet_name in excel.sheet_names() {
        println!("处理工作表: {}", sheet_name);
        
        if let Ok(range) = excel.worksheet_range(&sheet_name) {
            let rows: Vec<_> = range.rows().collect();
            
            if rows.is_empty() {
                continue;
            }

            // 获取表头
            let headers: Vec<String> = rows[0]
                .iter()
                .map(|cell| cell.to_string().trim().to_string())
                .collect();

            // 处理数据行
            for row in rows.iter().skip(1) {
                let mut row_data = serde_json::Map::new();
                
                for (col_idx, cell) in row.iter().enumerate() {
                    if col_idx < headers.len() {
                        let header = &headers[col_idx];
                        let value = match cell {
                            DataType::String(s) => json!(s.trim()),
                            DataType::Float(f) => json!(f),
                            DataType::Int(i) => json!(i),
                            DataType::Bool(b) => json!(b),
                            DataType::DateTime(d) => {
                                let dt = chrono::NaiveDateTime::from_timestamp_opt((d * 86400.0) as i64, 0)
                                    .unwrap_or_else(|| chrono::NaiveDateTime::from_timestamp_opt(0, 0).unwrap());
                                json!(dt.format("%Y-%m-%d").to_string())
                            }
                            DataType::Empty => json!(null),
                            _ => json!(cell.to_string()),
                        };
                        row_data.insert(header.clone(), value);
                    }
                }

                // 根据表头内容判断是充电记录还是停车记录
                let row_str = format!("{:?}", row_data);
                if row_str.contains("充电") || row_str.contains("里程") || row_str.contains("电费") {
                    charging_data.push(json!(row_data));
                } else if row_str.contains("停车") || row_str.contains("费用") {
                    parking_data.push(json!(row_data));
                } else if !row_data.is_empty() {
                    // 默认作为充电记录
                    charging_data.push(json!(row_data));
                }
            }
        }
    }

    // 5. 返回结果
    Ok(json!({
        "success": true,
        "message": format!("成功导入数据。充电记录: {} 条, 停车记录: {} 条", charging_data.len(), parking_data.len()),
        "chargingData": charging_data,
        "parkingData": parking_data,
        "totalRecords": charging_data.len() + parking_data.len()
    }))
}

// 保存Excel命令
#[tauri::command]
async fn save_excel_file(window: Window, excel_data: String, file_name: Option<String>) -> Result<(), AppError> {
    // 1. 打开保存对话框
    let default_name = file_name.unwrap_or_else(|| {
        format!("汽车费用记录_{}.xlsx", Local::now().format("%Y%m%d_%H%M%S"))
    });

    let file_path = tauri::api::dialog::FileDialogBuilder::new()
        .set_title("保存Excel文件")
        .set_file_name(&default_name)
        .add_filter("Excel文件", &["xlsx", "xls"])
        .save_file()
        .ok_or(AppError::UserCanceled)?;

    // 2. 解码Base64数据
    let file_bytes = STANDARD.decode(&excel_data)
        .map_err(|e| AppError::FileError(format!("数据解码失败: {}", e)))?;

    // 3. 写入文件
    fs::write(&file_path, file_bytes)
        .map_err(|e| AppError::FileError(format!("文件写入失败: {}", e)))?;

    println!("文件保存成功: {:?}", file_path);
    Ok(())
}

// 主函数
fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            import_excel_file,
            save_excel_file
        ])
        .setup(|app| {
            println!("汽车费用追踪应用启动");
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("运行Tauri应用时出错");
}