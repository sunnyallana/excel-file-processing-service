use actix_web::{web, App, HttpRequest, HttpResponse, HttpServer, Result, ResponseError};
use actix_multipart::Multipart;
use futures::StreamExt;
use serde::{Deserialize, Serialize};
use std::fs::File;
use std::io::{Write, Read, Cursor};
use tempfile::TempDir;
use zip::{read::ZipArchive, write::FileOptions, ZipWriter, result::ZipError};
use std::path::{Path, PathBuf};
use chrono::Local;
use calamine::{open_workbook, Reader, DataType, Xlsx, Error as CalamineError};
use rust_xlsxwriter::{Workbook, Format, Color};
use actix_cors::Cors;
use num_cpus;
use swagger_ui::{Assets, Config, Spec, swagger_spec_file};
use mime_guess::from_path;
use std::string::FromUtf8Error;
use actix_files as fs;
use std::fmt;

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

struct TextSegment {
    text: String,
    is_replaced: bool,
}

fn split_and_replace(original: &str, find: &str, replace: &str) -> Vec<TextSegment> {
    let mut result = Vec::new();
    let mut current = original;
    let mut found_replacement = false;

    while let Some(pos) = current.find(find) {
        if pos > 0 {
            result.push(TextSegment {
                text: current[..pos].to_string(),
                is_replaced: false
            });
        }

        result.push(TextSegment {
            text: replace.to_string(),
            is_replaced: true
        });

        found_replacement = true;
        current = &current[pos + find.len()..];
    }

    if !current.is_empty() {
        result.push(TextSegment {
            text: current.to_string(),
            is_replaced: false
        });
    }

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
    CalamineError(CalamineError),
    XlsxWriterError(rust_xlsxwriter::XlsxError),
    ZipError(ZipError),
    Utf8Error(FromUtf8Error),
    CalamineXlsxError(calamine::XlsxError),
}

impl fmt::Display for ProcessFileError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ProcessFileError::IoError(err) => write!(f, "IO error: {}", err),
            ProcessFileError::CalamineError(err) => write!(f, "Calamine error: {}", err),
            ProcessFileError::XlsxWriterError(err) => write!(f, "XlsxWriter error: {}", err),
            ProcessFileError::ZipError(err) => write!(f, "Zip error: {}", err),
            ProcessFileError::Utf8Error(err) => write!(f, "Utf8 error: {}", err),
            ProcessFileError::CalamineXlsxError(err) => write!(f, "Calamine Xlsx error: {}", err),
        }
    }
}

impl From<std::io::Error> for ProcessFileError {
    fn from(err: std::io::Error) -> ProcessFileError {
        ProcessFileError::IoError(err)
    }
}

impl From<CalamineError> for ProcessFileError {
    fn from(err: CalamineError) -> ProcessFileError {
        ProcessFileError::CalamineError(err)
    }
}

impl From<rust_xlsxwriter::XlsxError> for ProcessFileError {
    fn from(err: rust_xlsxwriter::XlsxError) -> ProcessFileError {
        ProcessFileError::XlsxWriterError(err)
    }
}

impl From<ZipError> for ProcessFileError {
    fn from(err: ZipError) -> ProcessFileError {
        ProcessFileError::ZipError(err)
    }
}

impl From<FromUtf8Error> for ProcessFileError {
    fn from(err: FromUtf8Error) -> ProcessFileError {
        ProcessFileError::Utf8Error(err)
    }
}

impl From<calamine::XlsxError> for ProcessFileError {
    fn from(err: calamine::XlsxError) -> ProcessFileError {
        ProcessFileError::CalamineXlsxError(err)
    }
}

impl ResponseError for ProcessFileError {
    fn error_response(&self) -> HttpResponse {
        HttpResponse::InternalServerError().finish()
    }
}

// Wrap ZipError in a custom error type that implements ResponseError
#[derive(Debug)]
struct CustomZipError(ZipError);

impl fmt::Display for CustomZipError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Zip error: {}", self.0)
    }
}

impl ResponseError for CustomZipError {
    fn error_response(&self) -> HttpResponse {
        HttpResponse::InternalServerError().finish()
    }
}

impl From<ZipError> for CustomZipError {
    fn from(err: ZipError) -> CustomZipError {
        CustomZipError(err)
    }
}

async fn process_file(
    file_path: PathBuf,
    original_filename: String,
    find_text: String,
    replace_text: String,
    temp_dir: &TempDir,
) -> Result<ProcessingResult, ProcessFileError> {
    let mut workbook: Xlsx<_> = open_workbook(&file_path).map_err(ProcessFileError::from)?;
    let sheet = workbook.worksheet_range_at(0).unwrap_or_else(|| panic!("No worksheet found")).map_err(ProcessFileError::from)?;

    let timestamp = Local::now().format("%m%d%y%H%M%S");
    let original_stem = Path::new(&original_filename).file_stem().unwrap_or_default().to_string_lossy().to_string();
    let new_filename = format!("{}Replace0-{}.xlsx", original_stem, timestamp);
    let _processed_path = temp_dir.path().join(&new_filename);

    let mut xlsx_workbook = Workbook::new();
    let worksheet = xlsx_workbook.add_worksheet();

    let mut total_replacements = 0;

    let default_format = Format::new();
    let red_format = Format::new().set_font_color(Color::Red).set_bold();

    let mut headers_found = false;
    for (row_idx, row) in sheet.rows().enumerate() {
        if headers_found {
            break;
        }
        for (col_idx, cell) in row.iter().enumerate() {
            match cell {
                DataType::String(s) => {
                    if s.contains(&*find_text) {
                        total_replacements += 1;

                        let parts = split_and_replace(s, &find_text, &replace_text);
                        let mut rich_text: Vec<(&Format, &str)> = Vec::new();
                        for part in &parts {
                            let format = if part.is_replaced { &red_format } else { &default_format };
                            rich_text.push((format, part.text.as_str()));
                        }

                        worksheet.write_rich_string(row_idx as u32, col_idx as u16, &rich_text)?;
                    } else {
                        worksheet.write_string(row_idx as u32, col_idx as u16, s.trim())?;
                    }

                    // Check if the cell contains a header
                    if s.trim().to_lowercase().contains("header") {
                        headers_found = true;
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

    let final_filename = format!("{}Replace{}-{}.xlsx", original_stem, total_replacements, timestamp);
    let final_processed_path = temp_dir.path().join(&final_filename);

    xlsx_workbook.save(&final_processed_path)?;

    Ok(ProcessingResult {
        replaced_count: total_replacements,
        filename: final_filename,
    })
}

async fn process_excel(mut payload: Multipart) -> Result<HttpResponse> {
    let temp_dir = TempDir::new()?;
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
                    find_text = String::from_utf8(content).map_err(ProcessFileError::from)?;
                },
                "replace" => {
                    let mut content = Vec::new();
                    while let Some(chunk) = field.next().await {
                        content.extend_from_slice(&chunk?);
                    }
                    replace_text = String::from_utf8(content).map_err(ProcessFileError::from)?;
                },
                "file" => {
                    let filename = content_disposition.get_filename().unwrap_or("uploaded_files.zip").to_string();
                    let filepath = temp_dir.path().join(&filename);
                    let mut f = File::create(&filepath)?;

                    while let Some(chunk) = field.next().await {
                        f.write_all(&chunk?)?;
                    }

                    zip_file_path = Some(filepath);
                },
                _ => {}
            }
        }
    }

    if zip_file_path.is_none() {
        return Err(actix_web::error::ErrorBadRequest("No file uploaded").into());
    }

    let zip_file_path = zip_file_path.unwrap();
    let file = File::open(&zip_file_path)?;
    let mut archive = ZipArchive::new(file).map_err(CustomZipError::from)?;

    for i in 0..archive.len() {
        let mut zip_file = archive.by_index(i).map_err(CustomZipError::from)?;
        let outpath = match zip_file.enclosed_name() {
            Some(path) => temp_dir.path().join(path),
            None => continue,
        };

        if outpath.extension().map_or(false, |ext| ext == "xls" || ext == "xlsx") {
            let mut outfile = File::create(&outpath)?;
            std::io::copy(&mut zip_file, &mut outfile)?;
            files_to_process.push((outpath, zip_file.name().to_string()));
        }
    }

    let mut zip_buffer = Vec::new();
    let mut processing_result_batch = ProcessingResultBatch::default();

    {
        let mut zip = ZipWriter::new(Cursor::new(&mut zip_buffer));
        let options = FileOptions::default().compression_method(zip::CompressionMethod::Deflated).unix_permissions(0o755);

        for (file_path, filename) in files_to_process {
            let result = process_file(file_path, filename.clone(), find_text.clone(), replace_text.clone(), &temp_dir).await;

            if let Ok(processed_file) = result {
                processing_result_batch.total_replaced_count += processed_file.replaced_count;
                processing_result_batch.processed_files.push(processed_file.clone());

                let final_processed_path = temp_dir.path().join(&processed_file.filename);
                zip.start_file(&processed_file.filename, options).map_err(CustomZipError::from)?;

                let mut content = Vec::new();
                File::open(&final_processed_path)?.read_to_end(&mut content)?;
                zip.write_all(&content)?;
            }
        }

        zip.finish().map_err(CustomZipError::from)?;
    }

    let timestamp = Local::now().format("%Y%m%d%H%M%S");

    Ok(HttpResponse::Ok()
        .content_type("application/zip")
        .append_header(("Content-Disposition", format!("attachment; filename=\"processed_files_{}.zip\"", timestamp)))
        .append_header(("X-Processed-Files-Count", processing_result_batch.processed_files.len().to_string()))
        .append_header(("X-Total-Replacements", processing_result_batch.total_replaced_count.to_string()))
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