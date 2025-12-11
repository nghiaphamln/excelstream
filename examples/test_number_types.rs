//! Test to verify that numbers are properly formatted in Excel
//!
//! This example creates an Excel file with different data types and
//! demonstrates that numbers (Int and Float) are formatted as numeric types,
//! not as text strings.

use excelstream::{
    types::{CellStyle, CellValue},
    ExcelWriter,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = ExcelWriter::new("test_number_types.xlsx")?;

    // Write header with bold style
    writer.write_header_bold(["Description", "Value", "Type", "Expected Format"])?; // Test various number formats
    writer.write_row_typed(&[
        CellValue::String("Small integer".to_string()),
        CellValue::Int(123),
        CellValue::String("Int".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    writer.write_row_typed(&[
        CellValue::String("Large integer".to_string()),
        CellValue::Int(1234567890),
        CellValue::String("Int".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    writer.write_row_typed(&[
        CellValue::String("Negative number".to_string()),
        CellValue::Int(-9876),
        CellValue::String("Int".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    writer.write_row_typed(&[
        CellValue::String("Float with decimals".to_string()),
        CellValue::Float(123.456),
        CellValue::String("Float".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    writer.write_row_typed(&[
        CellValue::String("Small decimal".to_string()),
        CellValue::Float(0.00123),
        CellValue::String("Float".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    writer.write_row_typed(&[
        CellValue::String("Scientific notation".to_string()),
        CellValue::Float(1.23e10),
        CellValue::String("Float".to_string()),
        CellValue::String("Number (right-aligned)".to_string()),
    ])?;

    // String number for comparison
    writer.write_row_typed(&[
        CellValue::String("String number".to_string()),
        CellValue::String("12345".to_string()),
        CellValue::String("String".to_string()),
        CellValue::String("Text (left-aligned)".to_string()),
    ])?;

    // Test with number formatting styles
    writer.write_row_typed(&[CellValue::Empty])?; // Empty row
    writer.write_row(["Testing with number styles:"])?;

    writer.write_row_styled(&[
        (
            CellValue::String("Integer with comma".to_string()),
            CellStyle::Default,
        ),
        (CellValue::Int(1234567), CellStyle::NumberInteger),
        (
            CellValue::String("NumberInteger style".to_string()),
            CellStyle::Default,
        ),
        (
            CellValue::String("1,234,567".to_string()),
            CellStyle::Default,
        ),
    ])?;

    writer.write_row_styled(&[
        (
            CellValue::String("Decimal format".to_string()),
            CellStyle::Default,
        ),
        (CellValue::Float(1234.5678), CellStyle::NumberDecimal),
        (
            CellValue::String("NumberDecimal style".to_string()),
            CellStyle::Default,
        ),
        (
            CellValue::String("1,234.57".to_string()),
            CellStyle::Default,
        ),
    ])?;

    writer.write_row_styled(&[
        (
            CellValue::String("Currency format".to_string()),
            CellStyle::Default,
        ),
        (CellValue::Float(1234.56), CellStyle::NumberCurrency),
        (
            CellValue::String("NumberCurrency style".to_string()),
            CellStyle::Default,
        ),
        (
            CellValue::String("$1,234.56".to_string()),
            CellStyle::Default,
        ),
    ])?;

    writer.write_row_styled(&[
        (
            CellValue::String("Percentage format".to_string()),
            CellStyle::Default,
        ),
        (CellValue::Float(0.95), CellStyle::NumberPercentage),
        (
            CellValue::String("NumberPercentage style".to_string()),
            CellStyle::Default,
        ),
        (CellValue::String("95.00%".to_string()), CellStyle::Default),
    ])?;

    writer.save()?;

    println!("\n✓ Excel file created: test_number_types.xlsx");
    println!("\nHow to verify:");
    println!("1. Open the file in Excel/LibreOffice");
    println!("2. Check column B (Value) alignment:");
    println!("   - Rows 2-7: Numbers should be RIGHT-aligned");
    println!("   - Row 8: String number should be LEFT-aligned");
    println!("3. Click on cells in column B to see the type in the formula bar");
    println!("4. Try to use SUM() function - it should work on numeric cells only");
    println!("\nExpected behavior:");
    println!("- CellValue::Int and CellValue::Float → Formatted as numbers (t=\"n\" in XML)");
    println!("- CellValue::String → Formatted as text (t=\"inlineStr\" in XML)");

    Ok(())
}
