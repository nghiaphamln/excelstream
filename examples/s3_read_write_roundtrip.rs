//! Example: S3 Write/Read roundtrip test
//!
//! Demonstrates uploading an Excel file to S3, then reading it back
//! to verify data integrity.
//!
//! Prerequisites:
//! 1. AWS credentials configured
//! 2. S3 bucket with write permissions
//!
//! Run with:
//! ```bash
//! export AWS_S3_BUCKET="your-test-bucket"
//! export AWS_REGION="us-east-1"  # Optional
//!
//! cargo run --example s3_read_write_roundtrip --features cloud-s3
//! ```

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::{S3ExcelReader, S3ExcelWriter};

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🔄 S3 Write/Read Roundtrip Test\n");
    println!("================================\n");

    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "my-excel-test".to_string());
    let key = "test/roundtrip_test.xlsx";
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "us-east-1".to_string());

    println!("📍 S3 Location: s3://{}/{}", bucket, key);
    println!("🌍 Region: {}\n", region);

    // Phase 1: Write to S3
    println!("📝 Phase 1: Writing test data to S3...");
    println!("---------------------------------------");
    let mut writer = S3ExcelWriter::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    writer.write_header_bold(["ID", "Product", "Quantity", "Price"])?;

    let test_data = vec![
        vec!["1", "Laptop", "10", "999.99"],
        vec!["2", "Mouse", "50", "29.99"],
        vec!["3", "Keyboard", "30", "79.99"],
        vec!["4", "Monitor", "15", "299.99"],
        vec!["5", "Webcam", "25", "89.99"],
    ];

    for row in &test_data {
        writer.write_row(row)?;
    }

    writer.save().await?;
    println!("✅ Write complete: {} rows + header\n", test_data.len());

    // Phase 2: Read from S3
    println!("📖 Phase 2: Reading back from S3...");
    println!("------------------------------------");
    let mut reader = S3ExcelReader::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    println!("✅ File downloaded and opened\n");

    // Verify data
    println!("🔍 Phase 3: Verifying data integrity...");
    println!("---------------------------------------");
    let mut read_rows = Vec::new();

    for row_result in reader.rows("Sheet1")? {
        let row = row_result?;
        read_rows.push(row.to_strings());
    }

    // Skip header (first row)
    let header = &read_rows[0];
    let data_rows = &read_rows[1..];

    // Verify header
    let expected_header = vec!["ID", "Product", "Quantity", "Price"];
    if header == &expected_header {
        println!("✅ Header matches: {:?}", header);
    } else {
        println!("❌ Header mismatch!");
        println!("   Expected: {:?}", expected_header);
        println!("   Actual:   {:?}", header);
    }

    // Verify row count
    if data_rows.len() == test_data.len() {
        println!("✅ Row count matches: {}", data_rows.len());
    } else {
        println!(
            "❌ Row count mismatch: expected {}, got {}",
            test_data.len(),
            data_rows.len()
        );
    }

    // Verify each row
    let mut all_match = true;
    for (i, (expected, actual)) in test_data.iter().zip(data_rows.iter()).enumerate() {
        if expected == actual {
            println!("  Row {}: ✅ Match - {:?}", i + 1, actual);
        } else {
            println!("  Row {}: ❌ Mismatch", i + 1);
            println!("    Expected: {:?}", expected);
            println!("    Actual:   {:?}", actual);
            all_match = false;
        }
    }

    println!();
    if all_match {
        println!("🎉 Roundtrip test PASSED! All data verified successfully.");
    } else {
        println!("❌ Roundtrip test FAILED! Data mismatch detected.");
        return Err("Data verification failed".into());
    }

    println!("\n💡 Tip: The temp file was automatically cleaned up after reading.");
    println!("   Memory usage: ~4-6 MB (SST in RAM only)");

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("❌ This example requires the 'cloud-s3' feature.");
    eprintln!("\nRun with:");
    eprintln!("  cargo run --example s3_read_write_roundtrip --features cloud-s3");
    eprintln!("\nEnvironment variables:");
    eprintln!("  export AWS_S3_BUCKET=your-test-bucket");
    eprintln!("  export AWS_REGION=us-east-1  # Optional");
    std::process::exit(1);
}
