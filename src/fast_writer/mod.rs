//! Fast Excel writer optimized for streaming
//!
//! This module provides a high-performance Excel writer that focuses on:
//! - Minimal memory allocations (2.7 MB for any file size with ZeroTempWorkbook)
//! - Direct XML generation
//! - Optimized ZIP compression
//! - Streaming-first design

pub mod memory;
pub mod shared_strings;
pub mod streaming_zip_reader;
pub mod streaming_zip_writer;
pub mod ultra_low_memory;
pub mod worksheet;
pub mod xml_writer;
pub mod zero_temp_workbook;

use crate::error::Result;
use std::path::Path;

pub use memory::{create_workbook_auto, create_workbook_with_profile, MemoryProfile};
pub use ultra_low_memory::UltraLowMemoryWorkbook;
pub use worksheet::FastWorksheet;
pub use zero_temp_workbook::ZeroTempWorkbook;

/// Create a fast Excel writer optimized for large datasets
///
/// # Examples
///
/// ```no_run
/// use excelstream::fast_writer::UltraLowMemoryWorkbook;
///
/// let mut workbook = UltraLowMemoryWorkbook::new("output.xlsx")?;
/// workbook.add_worksheet("Sheet1")?;
///
/// workbook.write_row(&["Name", "Age", "Email"])?;
/// workbook.write_row(&["Alice", "30", "alice@example.com"])?;
///
/// workbook.close()?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub fn create_fast_writer<P: AsRef<Path>>(path: P) -> Result<UltraLowMemoryWorkbook> {
    UltraLowMemoryWorkbook::new(path)
}
