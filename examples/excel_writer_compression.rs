//! Example: Using ExcelWriter with compression level configuration

use excelstream::writer::ExcelWriter;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== ExcelWriter Compression Level Examples ===\n");

    // Method 1: Set compression during construction
    println!("1. Using with_compression():");
    let mut writer1 = ExcelWriter::with_compression("excel_fast.xlsx", 1)?;
    writer1.write_row(["Name", "Age", "City"])?;
    for i in 0..1000 {
        writer1.write_row([&format!("User {}", i), "30", "NYC"])?;
    }
    writer1.save()?;
    println!("   ✓ Created excel_fast.xlsx with compression level 1");

    // Method 2: Set compression after construction
    println!("\n2. Using set_compression_level():");
    let mut writer2 = ExcelWriter::new("excel_custom.xlsx")?;
    println!("   Default compression: {}", writer2.compression_level());

    writer2.set_compression_level(3); // Moderate compression
    println!("   Changed to: {}", writer2.compression_level());

    writer2.write_row(["Name", "Age", "City"])?;
    for i in 0..1000 {
        writer2.write_row([&format!("User {}", i), "30", "NYC"])?;
    }
    writer2.save()?;
    println!("   ✓ Created excel_custom.xlsx with compression level 3");

    // Method 3: Different levels for different scenarios
    println!("\n3. Compression level recommendations:");

    // Development - fast compression
    let mut dev_writer = ExcelWriter::with_compression("excel_dev.xlsx", 1)?;
    dev_writer.write_row(["Test", "Data"])?;
    dev_writer.save()?;
    println!("   ✓ Development (level 1): excel_dev.xlsx");

    // Production - balanced
    let mut prod_writer = ExcelWriter::with_compression("excel_prod.xlsx", 6)?;
    prod_writer.write_row(["Test", "Data"])?;
    prod_writer.save()?;
    println!("   ✓ Production (level 6): excel_prod.xlsx");

    // Archive - maximum compression
    let mut archive_writer = ExcelWriter::with_compression("excel_archive.xlsx", 9)?;
    archive_writer.write_row(["Test", "Data"])?;
    archive_writer.save()?;
    println!("   ✓ Archive (level 9): excel_archive.xlsx");

    // Method 4: Environment-based configuration
    println!("\n4. Environment-based compression:");

    #[cfg(debug_assertions)]
    let compression = 1; // Fast for debug builds
    #[cfg(not(debug_assertions))]
    let compression = 6; // Balanced for release builds

    let mut env_writer = ExcelWriter::with_compression("excel_env.xlsx", compression)?;
    println!(
        "   Using compression level {} for {} build",
        compression,
        if cfg!(debug_assertions) {
            "DEBUG"
        } else {
            "RELEASE"
        }
    );

    env_writer.write_row(["Test", "Data"])?;
    env_writer.save()?;
    println!("   ✓ Created excel_env.xlsx");

    // File size comparison
    println!("\n5. File size comparison:");
    let files = ["excel_dev.xlsx", "excel_prod.xlsx", "excel_archive.xlsx"];

    for file in &files {
        if let Ok(metadata) = std::fs::metadata(file) {
            println!("   {}: {} bytes", file, metadata.len());
        }
    }

    println!("\n✅ All examples completed successfully!");
    Ok(())
}
