//! Example: Verify S3 uploaded file by downloading and reading it
//!
//! This downloads the Excel file from S3 and displays its contents
//! to verify the upload worked correctly.

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::S3ExcelReader;

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🔍 Verifying S3 Excel Upload\n");

    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "lune-nonprod".to_string());
    let key = "reports/monthly_sales_2024.xlsx";
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "ap-southeast-1".to_string());

    println!("📥 Downloading from s3://{}/{}", bucket, key);

    let mut reader = S3ExcelReader::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    println!("✅ Downloaded successfully\n");

    // Show sheet names
    println!("📋 Sheets: {:?}\n", reader.sheet_names());

    // Read first sheet
    println!("📊 Reading data from Sheet1:\n");
    let mut row_count = 0;
    for row in reader.rows("Sheet1")? {
        let row = row?;
        if row_count < 5 {
            println!("   Row {}: {:?}", row.index, row.to_strings());
        }
        row_count += 1;
    }

    println!("\n✅ Total rows: {}", row_count);
    println!("\n🎉 Verification successful! File uploaded and readable.");

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("❌ This example requires the 'cloud-s3' feature.");
    std::process::exit(1);
}
