//! Ultra-low memory workbook - wrapper around ZeroTempWorkbook

use super::zero_temp_workbook::ZeroTempWorkbook;
use crate::error::Result;
use crate::types::{CellValue, ProtectionOptions};
use std::path::Path;

pub struct UltraLowMemoryWorkbook {
    inner: ZeroTempWorkbook,
    compression_level: u32,
}

impl UltraLowMemoryWorkbook {
    pub fn new<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::with_compression(path, 6)
    }

    pub fn with_compression<P: AsRef<Path>>(path: P, compression_level: u32) -> Result<Self> {
        let inner = ZeroTempWorkbook::new(
            path.as_ref().to_str().unwrap_or("output.xlsx"),
            compression_level.min(9),
        )?;

        Ok(UltraLowMemoryWorkbook {
            inner,
            compression_level: compression_level.min(9),
        })
    }

    pub fn protect_sheet(&mut self, _options: ProtectionOptions) -> Result<()> {
        Ok(())
    }

    pub fn add_worksheet(&mut self, name: &str) -> Result<()> {
        self.inner.add_worksheet(name)
    }

    pub fn write_row(&mut self, values: &[&str]) -> Result<()> {
        self.inner.write_row(values)
    }

    pub fn write_row_typed(&mut self, values: &[CellValue]) -> Result<()> {
        let string_values: Vec<String> = values
            .iter()
            .map(|v| match v {
                CellValue::String(s) => s.clone(),
                CellValue::Int(i) => i.to_string(),
                CellValue::Float(f) => f.to_string(),
                CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
                CellValue::DateTime(dt) => dt.to_string(), // Excel serial date number
                CellValue::Error(e) => e.clone(),
                CellValue::Formula(f) => f.clone(),
                CellValue::Empty => String::new(),
            })
            .collect();

        let refs: Vec<&str> = string_values.iter().map(|s| s.as_str()).collect();
        self.inner.write_row(&refs)
    }

    pub fn write_row_styled(&mut self, _values: &[crate::types::StyledCell]) -> Result<()> {
        // TODO: Implement styling in ZeroTempWorkbook
        // For now, just write the cell values without styling
        let string_values: Vec<String> = _values
            .iter()
            .map(|styled| match &styled.value {
                CellValue::String(s) => s.clone(),
                CellValue::Int(i) => i.to_string(),
                CellValue::Float(f) => f.to_string(),
                CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_string(),
                CellValue::DateTime(dt) => dt.to_string(),
                CellValue::Error(e) => e.clone(),
                CellValue::Formula(f) => f.clone(),
                CellValue::Empty => String::new(),
            })
            .collect();

        let refs: Vec<&str> = string_values.iter().map(|s| s.as_str()).collect();
        self.inner.write_row(&refs)
    }

    pub fn set_compression_level(&mut self, level: u32) {
        self.compression_level = level.min(9);
    }

    pub fn compression_level(&self) -> u32 {
        self.compression_level
    }

    pub fn close(self) -> Result<()> {
        self.inner.close()
    }

    // Stub methods for API compatibility
    pub fn set_column_width(&mut self, _col: u32, _width: f64) -> Result<()> {
        // TODO: Implement in ZeroTempWorkbook
        Ok(())
    }

    pub fn set_next_row_height(&mut self, _height: f64) -> Result<()> {
        // TODO: Implement in ZeroTempWorkbook
        Ok(())
    }

    pub fn set_flush_interval(&mut self, _interval: u32) {
        // Not applicable for ZeroTempWorkbook (always streaming)
    }

    pub fn set_max_buffer_size(&mut self, _size: usize) {
        // Not applicable for ZeroTempWorkbook (uses fixed 4KB buffer)
    }
}
