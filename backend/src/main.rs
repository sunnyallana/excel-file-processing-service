use actix_web::{web, App, HttpRequest, HttpResponse, HttpServer, Result, ResponseError};
use actix_multipart::Multipart;
use futures::StreamExt;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Write, Read};
use tempfile::TempDir;
use zip::{read::ZipArchive, write::FileOptions, ZipWriter};
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
use postcard::to_vec;
use std::sync::Arc;
use std::sync::Mutex;
use std::fmt;
use tokio::task;
use tokio::task::JoinError;
use futures::future::join_all;
use actix_files as fs;

// Optimize the processing result structure for more efficient serialization
#[derive(Debug, Serialize, Deserialize, Default)]
struct ProcessingResultBatch {
    total_replaced_count: usize,
    processed_files: Vec<ProcessingResult>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
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
    JoinError(JoinError),
    PostcardError(postcard::Error),
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

impl From<JoinError> for ProcessFileError {
    fn from(err: JoinError) -> ProcessFileError {
        ProcessFileError::JoinError(err)
    }
}

impl From<postcard::Error> for ProcessFileError {
    fn from(err: postcard::Error) -> ProcessFileError {
        ProcessFileError::PostcardError(err)
    }
}

impl fmt::Display for ProcessFileError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ProcessFileError::IoError(err) => write!(f, "IO error: {}", err),
            ProcessFileError::CalamineError(err) => write!(f, "Calamine error: {}", err),
            ProcessFileError::XlsxWriterError(err) => write!(f, "XlsxWriter error: {}", err),
            ProcessFileError::ZipError(err) => write!(f, "Zip error: {}", err),
            ProcessFileError::JoinError(err) => write!(f, "Join error: {}", err),
            ProcessFileError::PostcardError(err) => write!(f, "Postcard error: {}", err),
        }
    }
}

impl ResponseError for ProcessFileError {
    fn error_response(&self) -> HttpResponse {
        HttpResponse::InternalServerError().finish()
    }
}

async fn process_file(
    file_path: PathBuf,
    original_filename: String,
    find_text: String,
    replace_text: String,
    temp_dir: Arc<TempDir>,
) -> Result<ProcessingResult, ProcessFileError> {
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

    Ok(ProcessingResult {
        replaced_count: total_replacements,
        filename: final_filename,
    })
}

async fn process_excel(mut payload: Multipart) -> Result<HttpResponse> {
    let temp_dir = Arc::new(TempDir::new().map_err(actix_web::error::ErrorInternalServerError)?);
    let mut files_to_process: Vec<(PathBuf, String)> = Vec::new();
    let mut find_text = String::new();
    let mut replace_text = String::new();
    let mut zip_file_path = None;

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
                "file" => {
                    let filename = content_disposition
                        .get_filename()
                        .unwrap_or("uploaded_files.zip")
                        .to_string();

                    let filepath = temp_dir.path().join(&filename);
                    let mut f = File::create(&filepath).map_err(actix_web::error::ErrorInternalServerError)?;

                    while let Some(chunk) = field.next().await {
                        f.write_all(&chunk?).map_err(actix_web::error::ErrorInternalServerError)?;
                    }

                    zip_file_path = Some(filepath);
                },
                _ => {}
            }
        }
    }

    if zip_file_path.is_none() {
        return Err(actix_web::error::ErrorBadRequest("No file uploaded"));
    }

    let zip_file_path = zip_file_path.unwrap();

    // Decompress the ZIP file
    let file = File::open(&zip_file_path).map_err(actix_web::error::ErrorInternalServerError)?;
    let mut archive = ZipArchive::new(file).map_err(actix_web::error::ErrorInternalServerError)?;
    
    for i in 0..archive.len() {
        let mut zip_file = archive.by_index(i).map_err(actix_web::error::ErrorInternalServerError)?;
        let outpath = match zip_file.enclosed_name() {
            Some(path) => temp_dir.path().join(path),
            None => continue,
        };

        if outpath.extension().map_or(false, |ext| ext == "xls" || ext == "xlsx") {
            let mut outfile = File::create(&outpath).map_err(actix_web::error::ErrorInternalServerError)?;
            std::io::copy(&mut zip_file, &mut outfile).map_err(actix_web::error::ErrorInternalServerError)?;
            files_to_process.push((outpath, zip_file.name().to_string()));
        }
    }

    info!("Files to process: {:?}", files_to_process);

    let mut zip_buffer = Vec::new();
    let processing_result_batch = Arc::new(Mutex::new(ProcessingResultBatch::default()));

    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_buffer));
        let options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o755);

        // Create a vector of tuples with cloned data
        let processed_files: Vec<_> = files_to_process.into_iter().map(|(file_path, filename)| {
            (file_path, filename, find_text.clone(), replace_text.clone())
        }).collect();

        // Use futures to process files
        let futures: Vec<_> = processed_files.into_iter().map(|(file_path, filename, find_text, replace_text)| {
            let processing_result_batch = Arc::clone(&processing_result_batch);
            let temp_dir = Arc::clone(&temp_dir);

            task::spawn(async move {
                let result = process_file(file_path, filename.clone(), find_text, replace_text, temp_dir).await;

                // If processing is successful, update the batch result
                if let Ok(ref processed_file) = result {
                    let mut batch = processing_result_batch.lock().unwrap();
                    batch.total_replaced_count += processed_file.replaced_count;
                    batch.processed_files.push(processed_file.clone());
                }

                result
            })
        }).collect();

        let results = join_all(futures).await;

        for result in results {
            match result {
                Ok(Ok(processed_file)) => {
                    // Add processed file to zip
                    let final_processed_path = temp_dir.path().join(&processed_file.filename);

                    zip.start_file(&processed_file.filename, options)
                        .map_err(actix_web::error::ErrorInternalServerError)?;

                    let mut content = Vec::new();
                    File::open(&final_processed_path)
                        .map_err(actix_web::error::ErrorInternalServerError)?
                        .read_to_end(&mut content)
                        .map_err(actix_web::error::ErrorInternalServerError)?;

                    zip.write_all(&content)
                        .map_err(actix_web::error::ErrorInternalServerError)?;
                },
                Ok(Err(e)) => return Err(ProcessFileError::from(e).into()),
                Err(e) => return Err(ProcessFileError::from(e).into()),
            }
        }

        zip.finish().map_err(actix_web::error::ErrorInternalServerError)?;
    }

    // Serialize the entire batch result using Postcard's more efficient serialization
    let batch_result = processing_result_batch.lock().unwrap();
    let _serialized_result: heapless::Vec<u8, 1024> = to_vec(&*batch_result)
        .map_err(actix_web::error::ErrorInternalServerError)?;

    // Generate a timestamp for the filename
    let timestamp = Local::now().format("%Y%m%d%H%M%S");

    // Optional: You can add the serialized metadata as a separate file in the zip or
    // send it as a custom header if needed
    // For this example, we'll just return the zip with the processed files

    Ok(HttpResponse::Ok()
        .content_type("application/zip")
        .append_header(("Content-Disposition", format!("attachment; filename=\"processed_files_{}.zip\"", timestamp)))
        .append_header(("X-Processed-Files-Count", batch_result.processed_files.len().to_string()))
        .append_header(("X-Total-Replacements", batch_result.total_replaced_count.to_string()))
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

    println!("Access Application at http://{}:{}/frontend", host, port);
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
            .service(
                fs::Files::new("/frontend", "../../frontend").index_file("index.html")
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