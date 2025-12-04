//! Test edge case - workbook with 5 sheets

use excelstream::writer::ExcelWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = ExcelWriter::new("windows_test_5sheets.xlsx")?;

    // Create 5 sheets
    for i in 1..=5 {
        if i > 1 {
            writer.add_sheet(&format!("Sheet{}", i))?;
        }
        writer.write_row(&[&format!("Data in Sheet {}", i)])?;
    }

    writer.save()?;

    println!("✅ Created windows_test_5sheets.xlsx with 5 sheets");

    Ok(())
}
