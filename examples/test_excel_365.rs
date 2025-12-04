use excelstream::fast_writer::FastWorkbook;
use excelstream::CellValue;
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("🧪 Testing ExcelStream v0.4.2 for Excel 365 Compatibility");
    println!("{}", "=".repeat(70));
    
    // Test 1: Single sheet with basic data
    println!("\n📝 Test 1: Creating single sheet file...");
    let mut wb1 = FastWorkbook::new("test_excel365_single.xlsx")?;
    wb1.add_worksheet("Sales Data")?;
    
    wb1.write_row(&["Product", "Quantity", "Price", "Total"])?;
    
    wb1.write_row(&["Product 1", "10", "99.99", "=B2*C2"])?;
    wb1.write_row(&["Product 2", "20", "99.99", "=B3*C3"])?;
    wb1.write_row(&["Product 3", "30", "99.99", "=B4*C4"])?;
    wb1.write_row(&["Product 4", "40", "99.99", "=B5*C5"])?;
    wb1.write_row(&["Product 5", "50", "99.99", "=B6*C6"])?;
    wb1.write_row(&["Product 6", "60", "99.99", "=B7*C7"])?;
    wb1.write_row(&["Product 7", "70", "99.99", "=B8*C8"])?;
    wb1.write_row(&["Product 8", "80", "99.99", "=B9*C9"])?;
    wb1.write_row(&["Product 9", "90", "99.99", "=B10*C10"])?;
    wb1.write_row(&["Product 10", "100", "99.99", "=B11*C11"])?;
    
    wb1.close()?;
    println!("✅ Created: test_excel365_single.xlsx (1 sheet, 11 rows)");
    
    // Test 2: Multiple sheets
    println!("\n📊 Test 2: Creating multi-sheet file...");
    let mut wb2 = FastWorkbook::new("test_excel365_multi.xlsx")?;
    
    // Sheet 1: Q1 Data
    wb2.add_worksheet("Q1 Sales")?;
    wb2.write_row(&["Month", "Revenue"])?;
    wb2.write_row(&["January", "50000"])?;
    wb2.write_row(&["February", "55000"])?;
    wb2.write_row(&["March", "60000"])?;
    
    // Sheet 2: Q2 Data
    wb2.add_worksheet("Q2 Sales")?;
    wb2.write_row(&["Month", "Revenue"])?;
    wb2.write_row(&["April", "65000"])?;
    wb2.write_row(&["May", "70000"])?;
    wb2.write_row(&["June", "75000"])?;
    
    // Sheet 3: Summary
    wb2.add_worksheet("Summary")?;
    wb2.write_row(&["Quarter", "Total Revenue"])?;
    wb2.write_row(&["Q1", "165000"])?;
    wb2.write_row(&["Q2", "210000"])?;
    
    wb2.close()?;
    println!("✅ Created: test_excel365_multi.xlsx (3 sheets)");
    
    // Test 3: Large dataset
    println!("\n📈 Test 3: Creating larger dataset...");
    let mut wb3 = FastWorkbook::new("test_excel365_large.xlsx")?;
    wb3.add_worksheet("Large Data")?;
    
    wb3.write_row(&["ID", "Name", "Value", "Status"])?;
    
    for i in 1..=1000 {
        let status = if i % 2 == 0 { "Active" } else { "Pending" };
        wb3.write_row(&[
            &i.to_string(),
            &format!("Item-{:04}", i),
            &format!("{:.1}", i as f64 * 1.5),
            status,
        ])?;
    }
    
    wb3.close()?;
    println!("✅ Created: test_excel365_large.xlsx (1 sheet, 1001 rows)");
    
    println!();
    println!("{}", "=".repeat(70));
    println!("🎯 TEST FILES READY FOR EXCEL 365");
    println!("{}", "=".repeat(70));
    println!("\n📋 Test Instructions:");
    println!("1. Copy these files to Windows:");
    println!("   - test_excel365_single.xlsx");
    println!("   - test_excel365_multi.xlsx");
    println!("   - test_excel365_large.xlsx");
    println!("\n2. Open each file in Excel 365");
    println!("\n3. Check for:");
    println!("   ❌ No '[Repaired]' in title bar");
    println!("   ❌ No 'PROTECTED VIEW' warning");
    println!("   ❌ No repair dialog on open");
    println!("   ✅ All sheets appear correctly");
    println!("   ✅ Data displays properly");
    println!("   ✅ Formulas calculate correctly");
    println!("\n4. Expected Results (v0.4.2 fixes applied):");
    println!("   ✅ Files open cleanly without errors");
    println!("   ✅ No hardcoded 3-sheet references");
    println!("   ✅ Dynamic generation working correctly");
    println!("\n💡 If you still see '[Repaired]' error, please:");
    println!("   - Screenshot the error message");
    println!("   - Check File > Info for details");
    println!("   - Report back what Excel says");
    println!();
    
    Ok(())
}
