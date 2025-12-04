//! Test Windows compatibility - minimal file to test on Windows Excel

use excelstream::writer::ExcelWriter;
use excelstream::types::CellValue;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = ExcelWriter::new("windows_test_minimal.xlsx")?;
    
    // Minimal data - just 2 cells
    writer.write_row(&["Hello"])?;
    
    writer.save()?;
    
    println!("✅ Created windows_test_minimal.xlsx - single cell");
    
    Ok(())
}
