//! Simple S3 read test with detailed error reporting

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::S3ExcelReader;

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🧪 S3ExcelReader Test\n");

    let bucket = "lune-nonprod";
    let key = "test/roundtrip_test.xlsx";
    let region = "ap-southeast-1";

    println!("📍 Location: s3://{}/{}", bucket, key);
    println!("🌍 Region: {}\n", region);

    println!("⏳ Building S3ExcelReader...");
    let reader_result = S3ExcelReader::builder()
        .bucket(bucket)
        .key(key)
        .region(region)
        .build()
        .await;

    match reader_result {
        Ok(mut reader) => {
            println!("✅ Reader built successfully!");
            println!("📋 Sheets: {:?}\n", reader.sheet_names());

            println!("📖 Reading rows...");
            let mut count = 0;
            for row_result in reader.rows("Sheet1")? {
                let row = row_result?;
                count += 1;
                if count <= 10 {
                    println!("  Row {}: {:?}", count, row.to_strings());
                }
            }
            println!("\n✅ Read {} rows successfully!", count);
        }
        Err(e) => {
            println!("❌ Error building reader:");
            println!("   Error type: {:?}", e);
            println!("   Error message: {}", e);
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
