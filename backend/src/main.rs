use actix_web::{web, App, HttpRequest, HttpResponse, HttpServer, Result, ResponseError};
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
use calamine::{open_workbook, Reader, DataType, Xlsx, XlsxError};
use rust_xlsxwriter::{Workbook, Format, Color};
use actix_cors::Cors;
use log::info;
use env_logger;
use num_cpus;
use swagger_ui::{Assets, Config, Spec, swagger_spec_file};
use mime_guess::from_path;
use postcard::to_stdvec;
use rayon::prelude::*;
use std::sync::Arc;
use std::sync::Mutex;
use std::fmt;

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

#[derive(Debug)]
enum ProcessFileError {
    IoError(std::io::Error),
    CalamineError(XlsxError),
    XlsxWriterError(rust_xlsxwriter::XlsxError),
    ZipError(zip::result::ZipError),
}

impl From<std::io::Error> for ProcessFileError {
    fn from(err: std::io::Error) -> ProcessFileError {
        ProcessFileError::IoError(err)
    }
}

impl From<XlsxError> for ProcessFileError {
    fn from(err: XlsxError) -> ProcessFileError {
        ProcessFileError::CalamineError(err)
    }
}

impl From<rust_xlsxwriter::XlsxError> for ProcessFileError {
    fn from(err: rust_xlsxwriter::XlsxError) -> ProcessFileError {
        ProcessFileError::XlsxWriterError(err)
    }
}

impl From<zip::result::ZipError> for ProcessFileError {
    fn from(err: zip::result::ZipError) -> ProcessFileError {
        ProcessFileError::ZipError(err)
    }
}

impl fmt::Display for ProcessFileError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ProcessFileError::IoError(err) => write!(f, "IO error: {}", err),
            ProcessFileError::CalamineError(err) => write!(f, "Calamine error: {}", err),
            ProcessFileError::XlsxWriterError(err) => write!(f, "XlsxWriter error: {}", err),
            ProcessFileError::ZipError(err) => write!(f, "Zip error: {}", err),
        }
    }
}

impl ResponseError for ProcessFileError {
    fn error_response(&self) -> HttpResponse {
        HttpResponse::InternalServerError().finish()
    }
}

fn process_file(
    file_path: &PathBuf,
    original_filename: &String,
    find_text: &String,
    replace_text: &String,
    temp_dir: &TempDir,
) -> Result<(PathBuf, usize), ProcessFileError> {
    // Open the workbook
    let mut workbook: Xlsx<_> = open_workbook(&file_path)?;

    // Get the first worksheet (as per requirements)
    let sheet = workbook.worksheet_range_at(0)
        .unwrap_or_else(|| panic!("No worksheet found"))?;

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
                    if s.contains(&*find_text) {
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
                        )?;
                    } else {
                        worksheet.write_string(row_idx as u32, col_idx as u16, s.trim())?;
                    }
                },
                DataType::Float(n) => {
                    worksheet.write_number(row_idx as u32, col_idx as u16, *n)?;
                },
                DataType::Int(n) => {
                    worksheet.write_number(row_idx as u32, col_idx as u16, *n as f64)?;
                },
                DataType::DateTime(dt) => {
                    worksheet.write_string(row_idx as u32, col_idx as u16, &dt.to_string())?;
                },
                _ => {
                    worksheet.write_string(row_idx as u32, col_idx as u16, "")?;
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

    xlsx_workbook.save(&final_processed_path)?;

    Ok((final_processed_path, total_replacements))
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
    let mut total_replacements = 0;
    let original_filename = Arc::new(Mutex::new(String::new()));

    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_buffer));
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o755);

        let results: Result<Vec<_>, ProcessFileError> = files_to_process.par_iter().map(|(file_path, filename)| {
            let mut original_filename = original_filename.lock().unwrap();
            *original_filename = filename.clone();
            process_file(file_path, filename, &find_text, &replace_text, &temp_dir)
        }).collect();

        for (final_processed_path, replacements) in results? {
            total_replacements += replacements;

            // Add processed file to zip
            zip.start_file(final_processed_path.file_name().unwrap().to_str().unwrap(), options)
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

    // Serialize the result using Postcard
    let processing_result = ProcessingResult {
        replaced_count: total_replacements,
        filename: original_filename.lock().unwrap().clone(),
    };
    let _serialized_result = to_stdvec(&processing_result).map_err(actix_web::error::ErrorInternalServerError)?;

    Ok(HttpResponse::Ok()
        .content_type("application/zip")
        .body(zip_buffer))
}

async fn swagger_ui_handler(req: HttpRequest) -> HttpResponse {
    let path = req.match_info().get("tail").unwrap_or("");
    match path {
        "" => {
            let _spec: Spec = swagger_spec_file!("../docs/openapi.yaml");
            let _config = Config {
                url: "/api-docs/openapi.yaml".to_string(),
                ..Default::default()
            };
            let html = format!(
                r#"
                <!DOCTYPE html>
                <html lang="en">
                  <head>
                    <meta charset="UTF-8">
                    <title>Swagger UI</title>
                    <link rel="stylesheet" type="text/css" href="/swagger-ui/swagger-ui.css" />
                    <link rel="icon" type="image/png" href="/swagger-ui/favicon-32x32.png" sizes="32x32" />
                    <link rel="icon" type="image/png" href="/swagger-ui/favicon-16x16.png" sizes="16x16" />
                  </head>
                  <body>
                    <div id="swagger-ui"></div>
                    <script src="/swagger-ui/swagger-ui-bundle.js"></script>
                    <script src="/swagger-ui/swagger-ui-standalone-preset.js"></script>
                    <script>
                    window.onload = function() {{
                      const ui = SwaggerUIBundle({{
                        url: "/docs/openapi.yaml",
                        dom_id: '#swagger-ui',
                        deepLinking: true,
                        presets: [
                          SwaggerUIBundle.presets.apis,
                          SwaggerUIStandalonePreset
                        ],
                        plugins: [
                          SwaggerUIBundle.plugins.DownloadUrl
                        ],
                        layout: "StandaloneLayout"
                      }})
                      window.ui = ui
                    }}
                  </script>
                  </body>
                </html>
                "#
            );
            HttpResponse::Ok()
                .content_type("text/html")
                .body(html)
        }
        _ => {
            if let Some(content) = Assets::get(path) {
                HttpResponse::Ok()
                    .content_type(from_path(path).first_or_octet_stream())
                    .body(content)
            } else {
                HttpResponse::NotFound().finish()
            }
        }
    }
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    env_logger::init();
    let host = "127.0.0.1";
    let port = 5000;
    let num_workers = num_cpus::get();

    println!("Server starting at http://{}:{}", host, port);
    println!("Access Swagger UI at http://{}:{}/swagger-ui/", host, port);

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
            .service(
                web::resource("/swagger-ui/{tail:.*}")
                    .route(web::get().to(swagger_ui_handler))
            )
            .route("/docs/openapi.yaml", web::get().to(|| async {
                HttpResponse::Ok()
                    .content_type("application/yaml")
                    .body(include_str!("../docs/openapi.yaml"))
            }))
    })
        .workers(num_workers)
        .backlog(1024)
        .bind((host, port))?
        .run()
        .await
}
