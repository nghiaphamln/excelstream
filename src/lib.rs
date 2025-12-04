//! # excelstream
//!
//! A high-performance Rust library for streaming Excel import/export operations.
//!
//! ## Features
//!
//! - **Streaming Read**: Read large Excel files without loading entire file into memory
//! - **Streaming Write**: Write millions of rows with constant ~80MB memory usage
//! - **Formula Support**: Write Excel formulas that calculate correctly
//! - **High Performance**: 30K-45K rows/sec throughput with true streaming
//! - **Better Errors**: Context-rich error messages with debugging info
//! - **Multiple Formats**: Support for XLSX, XLS, ODS formats
//! - **Type Safety**: Strong typing with Rust's type system
//! - **Zero-copy**: Minimize memory allocations where possible
//!
//! ## Quick Start
//!
//! ### Reading Excel Files (Streaming)
//!
//! ```rust,no_run
//! use excelstream::reader::ExcelReader;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let mut reader = ExcelReader::open("data.xlsx")?;
//!
//! for row_result in reader.rows("Sheet1")? {
//!     let row = row_result?;
//!     println!("Row: {:?}", row);
//! }
//! # Ok(())
//! # }
//! ```
//!
//! ### Writing Excel Files (Streaming)
//!
//! ```rust,no_run
//! use excelstream::writer::ExcelWriter;
//! use excelstream::types::CellValue;
//!
//! # fn main() -> Result<(), Box<dyn std::error::Error>> {
//! let mut writer = ExcelWriter::new("output.xlsx")?;
//!
//! // Write header
//! writer.write_header(&["Name", "Age", "City"])?;
//!
//! // Write data rows with typed values
//! writer.write_row_typed(&[
//!     CellValue::String("Alice".to_string()),
//!     CellValue::Int(30),
//!     CellValue::String("New York".to_string()),
//! ])?;
//!
//! // Write with formulas
//! writer.write_row_typed(&[
//!     CellValue::String("Total".to_string()),
//!     CellValue::Formula("=COUNT(B2:B10)".to_string()),
//!     CellValue::Empty,
//! ])?;
//!
//! writer.save()?;
//! # Ok(())
//! # }
//! ```

pub mod error;
pub mod fast_writer;
pub mod reader;
pub mod types;
pub mod writer;

pub use error::{ExcelError, Result};
pub use reader::ExcelReader;
pub use types::{Cell, CellStyle, CellValue, Row, StyledCell};
pub use writer::ExcelWriter;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_library_imports() {
        // Test that all public types are accessible
        let _ = std::marker::PhantomData::<ExcelError>;
        let _ = std::marker::PhantomData::<ExcelReader>;
        let _ = std::marker::PhantomData::<ExcelWriter>;
    }
}

/// Reorder XLSX ZIP file to have [Content_Types].xml first (required by Office spec)
/// 
/// This function ensures Office Open XML compliance by reordering internal files.
/// Some tools like Excel Online are strict about this ordering.
/// 
/// # Arguments
/// * `file_path` - Path to the .xlsx file to fix
/// 
/// # Errors
/// Returns error if file cannot be read or written
pub fn fix_xlsx_zip_order<P: AsRef<std::path::Path>>(file_path: P) -> Result<()> {
    use std::io::Write;
    use zip::ZipArchive;
    
    let file_path = file_path.as_ref();
    let temp_path = file_path.with_extension("tmp.xlsx");
    
    // Define correct Office Open XML order
    const CORRECT_ORDER: &[&str] = &[
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
        "xl/theme/theme1.xml",
        "xl/styles.xml",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/sharedStrings.xml",
        "docProps/core.xml",
        "docProps/app.xml",
    ];
    
    // Read all files from original
    let file = std::fs::File::open(file_path)?;
    let mut archive = ZipArchive::new(file)?;
    let mut files_data: std::collections::HashMap<String, Vec<u8>> = std::collections::HashMap::new();
    
    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();
        let mut data = Vec::new();
        std::io::Read::read_to_end(&mut file, &mut data)?;
        files_data.insert(name, data);
    }
    
    // Write in correct order
    let temp_file = std::fs::File::create(&temp_path)?;
    let mut zip_writer = zip::ZipWriter::new(temp_file);
    let options = zip::write::FileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    
    // First add in correct order
    for fname in CORRECT_ORDER.iter() {
        if let Some(data) = files_data.remove(*fname) {
            zip_writer.start_file(*fname, options)?;
            zip_writer.write_all(&data)?;
        }
    }
    
    // Then add any remaining files (for multi-sheet support)
    for (fname, data) in files_data.iter() {
        zip_writer.start_file(fname, options)?;
        zip_writer.write_all(data)?;
    }
    
    zip_writer.finish()?;
    
    // Replace original with temp
    std::fs::rename(&temp_path, file_path)?;
    
    Ok(())
}
