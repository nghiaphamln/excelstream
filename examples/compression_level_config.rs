//! Example: Configuring compression level for UltraLowMemoryWorkbook

use excelstream::fast_writer::UltraLowMemoryWorkbook;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Compression Level Configuration Examples ===\n");

    // Method 1: Set compression level during construction
    println!("1. Setting compression level during construction:");

    // Fast compression (level 1) - good for development/testing
    let mut wb_fast = UltraLowMemoryWorkbook::with_compression("output_fast.xlsx", 1)?;
    wb_fast.add_worksheet("FastSheet")?;
    for i in 0..1000 {
        wb_fast.write_row(&["Data", &i.to_string(), "Value"])?;
    }
    wb_fast.close()?;
    println!("   ✓ Created output_fast.xlsx with compression level 1 (fast)");

    // Balanced compression (level 6) - recommended for production
    let mut wb_balanced = UltraLowMemoryWorkbook::with_compression("output_balanced.xlsx", 6)?;
    wb_balanced.add_worksheet("BalancedSheet")?;
    for i in 0..1000 {
        wb_balanced.write_row(&["Data", &i.to_string(), "Value"])?;
    }
    wb_balanced.close()?;
    println!("   ✓ Created output_balanced.xlsx with compression level 6 (balanced)");

    // Method 2: Set compression level after construction
    println!("\n2. Setting compression level after construction:");

    let mut wb = UltraLowMemoryWorkbook::new("output_configurable.xlsx")?;

    // Check current compression level
    println!("   Default compression level: {}", wb.compression_level());

    // Change to fast compression for development
    wb.set_compression_level(1);
    println!(
        "   Changed compression level to: {}",
        wb.compression_level()
    );

    wb.add_worksheet("ConfigSheet")?;
    for i in 0..1000 {
        wb.write_row(&["Data", &i.to_string(), "Value"])?;
    }
    wb.close()?;
    println!("   ✓ Created output_configurable.xlsx with custom compression level");

    // Method 3: Different compression levels for different scenarios
    println!("\n3. Compression level recommendations:");
    println!("   Level 0: No compression - fastest, ~280MB for 1M rows");
    println!("   Level 1: Fast - very fast, ~31MB for 1M rows (good for dev)");
    println!("   Level 3: Moderate - fast, ~22MB for 1M rows");
    println!("   Level 6: Balanced - ~18MB for 1M rows (RECOMMENDED for production)");
    println!("   Level 9: Maximum - slowest, ~18MB for 1M rows");

    // Example: Development vs Production
    println!("\n4. Development vs Production example:");

    #[cfg(debug_assertions)]
    let compression = 1; // Fast for development
    #[cfg(not(debug_assertions))]
    let compression = 6; // Balanced for production

    let mut wb_env = UltraLowMemoryWorkbook::with_compression("output_env.xlsx", compression)?;
    println!(
        "   Using compression level {} for {} build",
        compression,
        if cfg!(debug_assertions) {
            "DEBUG"
        } else {
            "RELEASE"
        }
    );

    wb_env.add_worksheet("EnvSheet")?;
    for i in 0..1000 {
        wb_env.write_row(&["Data", &i.to_string(), "Value"])?;
    }
    wb_env.close()?;
    println!("   ✓ Created output_env.xlsx");

    println!("\n✓ All examples completed successfully!");
    Ok(())
}
