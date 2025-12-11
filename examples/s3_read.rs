//! Example: Read Excel file from Amazon S3
//!
//! This example demonstrates how to read Excel files directly from S3
//! using streaming to minimize memory usage.
//!
//! Prerequisites:
//! 1. AWS credentials configured (via ~/.aws/credentials or environment)
//! 2. S3 bucket with Excel file
//!
//! Run with:
//! ```bash
//! # Set environment variables
//! export AWS_S3_BUCKET="your-bucket-name"
//! export AWS_S3_KEY="path/to/file.xlsx"
//! export AWS_REGION="us-east-1"  # Optional, defaults to us-east-1
//!
//! # Or use test credentials
//! export AWS_ACCESS_KEY_ID="your-access-key"
//! export AWS_SECRET_ACCESS_KEY="your-secret-key"
//!
//! # Run example
//! cargo run --example s3_read --features cloud-s3
//! ```

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::S3ExcelReader;

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("📖 ExcelStream S3 Reader Example\n");
    println!("=================================\n");

    // Configuration from environment
    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "my-excel-reports".to_string());
    let key = std::env::var("AWS_S3_KEY")
        .unwrap_or_else(|_| "reports/monthly_sales_2024.xlsx".to_string());
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "us-east-1".to_string());

    println!("📍 Reading from: s3://{}/{}", bucket, key);
    println!("🌍 Region: {}\n", region);

    // Build S3 reader (downloads file)
    println!("⏳ Downloading and opening Excel file...");
    let mut reader = S3ExcelReader::builder()
        .bucket(&bucket)
        .key(&key)
        .region(&region)
        .build()
        .await?;

    println!("✅ File opened successfully\n");

    // List all sheets
    let sheet_names = reader.sheet_names();
    println!("📋 Found {} sheets:", sheet_names.len());
    for (i, name) in sheet_names.iter().enumerate() {
        println!("  {}. {}", i + 1, name);
    }
    println!();

    // Read first sheet
    if let Some(first_sheet) = sheet_names.first() {
        println!("📊 Reading sheet: {}\n", first_sheet);

        let mut row_count = 0;
        for row_result in reader.rows(first_sheet)? {
            let row = row_result?;
            row_count += 1;

            // Print first 5 rows
            if row_count <= 5 {
                println!("Row {}: {:?}", row.index + 1, row.to_strings());
            }

            // Progress indicator for large files
            if row_count % 1000 == 0 {
                println!("  ... processed {} rows", row_count);
            }
        }

        println!("\n✅ Processed {} total rows", row_count);
    }

    println!("\n🎉 S3 read complete!");

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("❌ This example requires the 'cloud-s3' feature.");
    eprintln!("\nRun with:");
    eprintln!("  cargo run --example s3_read --features cloud-s3");
    eprintln!("\nEnvironment variables:");
    eprintln!("  export AWS_S3_BUCKET=your-bucket-name");
    eprintln!("  export AWS_S3_KEY=path/to/file.xlsx");
    eprintln!("  export AWS_REGION=us-east-1  # Optional");
    std::process::exit(1);
}
