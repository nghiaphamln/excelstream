//! Test Windows compatibility with multiple sheets

use excelstream::writer::ExcelWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = ExcelWriter::new("windows_test_multisheet.xlsx")?;
    
    // Sheet 1
    writer.write_row(&["Sheet1", "Data"])?;
    
    // Sheet 2
    writer.add_sheet("Sheet2")?;
    writer.write_row(&["Sheet2", "Data"])?;
    
    // Sheet 3
    writer.add_sheet("Sheet3")?;
    writer.write_row(&["Sheet3", "Data"])?;
    
    writer.save()?;
    
    println!("✅ Created windows_test_multisheet.xlsx with 3 sheets");
    
    Ok(())
}
