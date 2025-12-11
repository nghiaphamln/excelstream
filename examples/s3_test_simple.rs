//! Simple S3 write then read test

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::{S3ExcelReader, S3ExcelWriter};

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "lune-nonprod".to_string());
    let key = "reports/test_reader.xlsx";
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "ap-southeast-1".to_string());

    println!("🚀 S3 Write/Read Test\n");
    println!("📍 Location: s3://{}/{}", bucket, key);
    println!("🌍 Region: {}\n", region);

    // STEP 1: Write
    println!("📝 STEP 1: Writing to S3...");
    let mut writer = S3ExcelWriter::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    writer.write_header_bold(["ID", "Name", "Value"])?;
    writer.write_row(["1", "Test A", "100"])?;
    writer.write_row(["2", "Test B", "200"])?;
    writer.write_row(["3", "Test C", "300"])?;

    writer.save().await?;
    println!("✅ Write completed\n");

    // STEP 2: Read
    println!("📖 STEP 2: Reading from S3...");

    let reader_result = S3ExcelReader::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await;

    match reader_result {
        Ok(mut reader) => {
            println!("✅ Reader initialized successfully!\n");

            let sheets = reader.sheet_names();
            println!("📋 Sheets found: {:?}", sheets);

            println!("\n📊 Reading data:");
            for row_result in reader.rows("Sheet1")? {
                let row = row_result?;
                println!("  {:?}", row.to_strings());
            }

            println!("\n🎉 Test PASSED!");
        }
        Err(e) => {
            println!("❌ Reader failed!");
            println!("   Error: {:?}", e);
            println!("   Message: {}", e);

            // Try to get more details
            if let excelstream::ExcelError::IoError(io_err) = &e {
                println!("   IO Error kind: {:?}", io_err.kind());
                if let Some(inner) = io_err.get_ref() {
                    println!("   Inner error: {:?}", inner);
                }
            }

            return Err(e.into());
        }
    }

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("This example requires the 'cloud-s3' feature");
    std::process::exit(1);
}
