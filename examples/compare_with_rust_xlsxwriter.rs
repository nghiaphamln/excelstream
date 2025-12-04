//! Compare output with rust_xlsxwriter library
//!
//! This creates identical files using both libraries to compare structure

use excelstream::writer::ExcelWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Generating ExcelStream file ===");

    // Create with excelstream
    let mut writer = ExcelWriter::new("output_excelstream.xlsx")?;
    writer.write_row(&["Name", "Age", "City"])?;
    writer.write_row(&["Alice", "30", "NYC"])?;
    writer.write_row(&["Bob", "25", "LA"])?;
    writer.save()?;

    println!("✅ Created output_excelstream.xlsx");

    // Try to use rust_xlsxwriter if available
    #[cfg(feature = "compare-rust-xlsxwriter")]
    {
        println!("\n=== Generating rust_xlsxwriter file ===");
        use rust_xlsxwriter::*;

        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        worksheet.write_string(0, 0, "Name")?;
        worksheet.write_string(0, 1, "Age")?;
        worksheet.write_string(0, 2, "City")?;

        worksheet.write_string(1, 0, "Alice")?;
        worksheet.write_string(1, 1, "30")?;
        worksheet.write_string(1, 2, "NYC")?;

        worksheet.write_string(2, 0, "Bob")?;
        worksheet.write_string(2, 1, "25")?;
        worksheet.write_string(2, 2, "LA")?;

        workbook.save("output_rust_xlsxwriter.xlsx")?;

        println!("✅ Created output_rust_xlsxwriter.xlsx");

        // Compare
        println!("\n=== Running Comparison ===");
        std::process::Command::new("python3")
            .arg("-c")
            .arg(
                r#"
import zipfile, hashlib

def compare_files(f1, f2):
    with zipfile.ZipFile(f1) as z1, zipfile.ZipFile(f2) as z2:
        files1 = sorted(z1.namelist())
        files2 = sorted(z2.namelist())
        
        print(f"\n{f1}: {len(files1)} files")
        print(f"{f2}: {len(files2)} files")
        
        common = set(files1) & set(files2)
        only1 = set(files1) - set(files2)
        only2 = set(files2) - set(files1)
        
        if only1:
            print(f"\nOnly in {f1}: {only1}")
        if only2:
            print(f"\nOnly in {f2}: {only2}")
        
        print(f"\nComparing {len(common)} common files:")
        differences = []
        for fname in sorted(common):
            d1 = z1.read(fname)
            d2 = z2.read(fname)
            if d1 == d2:
                print(f"  ✓ {fname} - IDENTICAL")
            else:
                m1 = hashlib.md5(d1).hexdigest()[:8]
                m2 = hashlib.md5(d2).hexdigest()[:8]
                print(f"  ✗ {fname} - DIFFERENT ({m1} vs {m2})")
                differences.append(fname)
        
        if differences:
            print(f"\n⚠️  {len(differences)} files differ:")
            for f in differences:
                print(f"    - {f}")
        else:
            print("\n✅ All common files are IDENTICAL!")

compare_files('output_excelstream.xlsx', 'output_rust_xlsxwriter.xlsx')
"#,
            )
            .status()
            .expect("Failed to run comparison");
    }

    #[cfg(not(feature = "compare-rust-xlsxwriter"))]
    {
        println!("\n⚠️  rust_xlsxwriter not available for comparison");
        println!("To enable: add rust_xlsxwriter as dev-dependency with feature flag");
    }

    Ok(())
}
