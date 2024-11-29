use actix_web::{web, App, HttpResponse, HttpServer, Result};
use actix_multipart::Multipart;
use futures::StreamExt;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Write, Read};
use tempfile::TempDir;
use zip::{write::FileOptions, ZipWriter};
use std::io::Cursor;
use std::path::{Path, PathBuf};
use chrono::Local;
use calamine::{open_workbook, Reader, DataType, Xlsx};
use rust_xlsxwriter::{Workbook, Format, Color};
use actix_cors::Cors;
use log::info;
use env_logger;
use num_cpus;

#[derive(Debug, Serialize, Deserialize)]
struct ProcessingResult {
    replaced_count: usize,
    filename: String,
}

// Structure to represent a text segment with its coloring information
#[derive(Clone)]
struct TextSegment {
    text: String,
    is_replaced: bool,
}

// Function to split a string with the replaced part
fn split_and_replace(original: &str, find: &str, replace: &str) -> Vec<TextSegment> {
    let mut result = Vec::new();
    let mut current = original;
    let mut found_replacement = false;

    while let Some(pos) = current.find(find) {
        // Add the part before the find string (if any)
        if pos > 0 {
            result.push(TextSegment {
                text: current[..pos].to_string(),
                is_replaced: false
            });
        }

        // Add the replaced part
        result.push(TextSegment {
            text: replace.to_string(),
            is_replaced: true
        });

        found_replacement = true;

        // Move to the part after the find string
        current = &current[pos + find.len()..];
    }

    // Add any remaining part
    if !current.is_empty() {
        result.push(TextSegment {
            text: current.to_string(),
            is_replaced: false
        });
    }

    // If no replacements were made, return the original string
    if !found_replacement {
        result = vec![TextSegment {
            text: original.to_string(),
            is_replaced: false
        }];
    }

    result
}

async fn process_excel(mut payload: Multipart) -> Result<HttpResponse> {
    let temp_dir = TempDir::new().map_err(actix_web::error::ErrorInternalServerError)?;
    let mut files_to_process: Vec<(PathBuf, String)> = Vec::new();
    let mut find_text = String::new();
    let mut replace_text = String::new();

    while let Some(item) = payload.next().await {
        let mut field = item?;
        let content_disposition = field.content_disposition();

        if let Some(name) = content_disposition.get_name() {
            match name {
                "find" => {
                    let mut content = Vec::new();
                    while let Some(chunk) = field.next().await {
                        content.extend_from_slice(&chunk?);
                    }
                    find_text = String::from_utf8(content).map_err(actix_web::error::ErrorInternalServerError)?;
                },
                "replace" => {
                    let mut content = Vec::new();
                    while let Some(chunk) = field.next().await {
                        content.extend_from_slice(&chunk?);
                    }
                    replace_text = String::from_utf8(content).map_err(actix_web::error::ErrorInternalServerError)?;
                },
                "files[]" => {
                    let filename = content_disposition
                        .get_filename()
                        .unwrap_or("unknown.xlsx")
                        .to_string();

                    let filepath = temp_dir.path().join(&filename);
                    let mut f = File::create(&filepath).map_err(actix_web::error::ErrorInternalServerError)?;

                    while let Some(chunk) = field.next().await {
                        f.write_all(&chunk?).map_err(actix_web::error::ErrorInternalServerError)?;
                    }

                    files_to_process.push((filepath, filename));
                },
                _ => {}
            }
        }
    }

    info!("Files to process: {:?}", files_to_process);

    let mut zip_buffer = Vec::new();
    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_buffer));
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o755);

        for (file_path, original_filename) in &files_to_process {
            // Open the workbook
            let mut workbook: Xlsx<_> = open_workbook(&file_path)
                .map_err(actix_web::error::ErrorInternalServerError)?;

            // Get the first worksheet (as per requirements)
            let sheet = workbook.worksheet_range_at(0)
                .unwrap_or_else(|| panic!("No worksheet found"))
                .map_err(actix_web::error::ErrorInternalServerError)?;

            // Prepare output filename with timestamp
            let timestamp = Local::now().format("%m%d%y%H%M%S");
            let original_stem = Path::new(&original_filename).file_stem().unwrap_or_default().to_string_lossy().to_string();
            let new_filename = format!(
                "{}Replace0-{}.xlsx",
                original_stem,
                timestamp
            );
            let _processed_path = temp_dir.path().join(&new_filename);

            // Create a new workbook for output
            let mut xlsx_workbook = Workbook::new();
            let worksheet = xlsx_workbook.add_worksheet();

            let mut total_replacements = 0;

            // Create formats
            let default_format = Format::new();
            let red_format = Format::new()
                .set_font_color(Color::Red)
                .set_bold();

            // Process each cell in the first sheet
            for (row_idx, row) in sheet.rows().enumerate() {
                for (col_idx, cell) in row.iter().enumerate() {
                    match cell {
                        DataType::String(s) => {
                            if s.contains(&find_text) {
                                total_replacements += 1;

                                // Split the string into parts to be colored
                                let parts = split_and_replace(s, &find_text, &replace_text);

                                // For cells with replacements, create a rich text format
                                let mut rich_text: Vec<(&Format, &str)> = Vec::new();
                                for part in &parts {
                                    let format = if part.is_replaced {
                                        &red_format
                                    } else {
                                        &default_format
                                    };

                                    rich_text.push((format, part.text.as_str()));
                                }

                                // Write the rich text to the cell
                                worksheet.write_rich_string(
                                    row_idx as u32,
                                    col_idx as u16,
                                    &rich_text
                                ).map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                            } else {
                                worksheet.write_string(row_idx as u32, col_idx as u16, s.trim())
                                    .map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                            }
                        },
                        DataType::Float(n) => {
                            worksheet.write_number(row_idx as u32, col_idx as u16, *n)
                                .map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                        },
                        DataType::Int(n) => {
                            worksheet.write_number(row_idx as u32, col_idx as u16, *n as f64)
                                .map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                        },
                        DataType::DateTime(dt) => {
                            worksheet.write_string(row_idx as u32, col_idx as u16, &dt.to_string())
                                .map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                        },
                        _ => {
                            worksheet.write_string(row_idx as u32, col_idx as u16, "")
                                .map_err(|e| actix_web::error::ErrorInternalServerError(e))?;
                        }
                    }
                }
            }

            let final_filename = format!(
                "{}Replace{}-{}.xlsx",
                original_stem,
                total_replacements,
                timestamp
            );
            let final_processed_path = temp_dir.path().join(&final_filename);

            xlsx_workbook.save(&final_processed_path).map_err(actix_web::error::ErrorInternalServerError)?;

            // Add processed file to zip
            zip.start_file(&final_filename, options)
                .map_err(actix_web::error::ErrorInternalServerError)?;

            let mut content = Vec::new();
            File::open(&final_processed_path)
                .map_err(actix_web::error::ErrorInternalServerError)?
                .read_to_end(&mut content)
                .map_err(actix_web::error::ErrorInternalServerError)?;

            zip.write_all(&content)
                .map_err(actix_web::error::ErrorInternalServerError)?;
        }

        zip.finish().map_err(actix_web::error::ErrorInternalServerError)?;
    }

    Ok(HttpResponse::Ok()
        .content_type("application/zip")
        .body(zip_buffer))
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    env_logger::init();
    let host = "127.0.0.1";
    let port = 5000;
    let num_workers = num_cpus::get();

    println!("Server starting at http://{}:{}", host, port);

    HttpServer::new(|| {
        App::new()
            .wrap(Cors::default()
                .allow_any_origin()
                .allow_any_method()
                .allow_any_header())
            .service(
                web::resource("/process-excel")
                    .route(web::post().to(process_excel))
            )
    })
        .workers(num_workers)
        .backlog(1024)
        .bind((host, port))?
        .run()
        .await
}