//! Test Windows compatibility - minimal file to test on Windows Excel

use excelstream::writer::ExcelWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create with custom compression level (1 = fast, 6 = balanced, 9 = max)
    let mut writer = ExcelWriter::with_compression("windows_test_minimal.xlsx", 1)?;

    // Or use default (level 6) then change
    // let mut writer = ExcelWriter::new("windows_test_minimal.xlsx")?;
    // writer.set_compression_level(1);

    // Check current compression level
    println!("Compression level: {}", writer.compression_level());

    // Minimal data - just 2 cells
    writer.write_row(["Hello"])?;

    writer.save()?;

    println!("✅ Created windows_test_minimal.xlsx - single cell with compression level 1");

    Ok(())
}
