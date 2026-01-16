//! Example: Stream Excel file directly to Amazon S3 (TRUE STREAMING - NO TEMP FILES!)
//!
//! This example demonstrates how to generate Excel files and upload them
//! directly to S3 using s-zip's cloud support - NO temporary files needed!
//!
//! Benefits:
//! - ✅ ZERO disk usage (perfect for Lambda/containers)
//! - ✅ Works in read-only filesystems
//! - ✅ Constant ~4 KB memory usage for buffering
//! - ✅ TRUE streaming using s-zip's S3ZipWriter
//! - ✅ Multipart upload handled by s-zip
//!
//! Prerequisites:
//! 1. AWS credentials configured (via ~/.aws/credentials or environment variables)
//! 2. S3 bucket exists with proper permissions
//!
//! Run with:
//! ```bash
//! export AWS_ACCESS_KEY_ID="your-key"
//! export AWS_SECRET_ACCESS_KEY="your-secret"
//! export AWS_REGION="ap-southeast-1"
//! export AWS_S3_BUCKET="your-bucket"
//! cargo run --example s3_streaming --features cloud-s3
//! ```

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::S3ExcelWriter;

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🚀 ExcelStream S3 Direct Streaming Example (v0.16.0)");
    println!("   Using s-zip 0.8.0 cloud support - NO TEMP FILES!\n");

    // Configuration
    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "my-excel-reports".to_string());
    let key = "reports/monthly_sales_2024.xlsx";
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "us-east-1".to_string());

    println!("📍 Target: s3://{}/{}", bucket, key);
    println!("🌍 Region: {}\n", region);

    // Create S3 Excel writer - streams directly to S3!
    println!("⏳ Creating S3 Excel writer...");
    let mut writer = S3ExcelWriter::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    println!("✅ S3 writer initialized (using s-zip's S3ZipWriter)\n");

    // Write header
    println!("📝 Writing header row...");
    writer
        .write_header_bold(["Month", "Product", "Sales", "Profit"])
        .await?;

    // Generate sample data
    println!("📊 Writing sales data...");
    let months = ["January", "February", "March", "April", "May", "June"];
    let products = ["Laptop", "Phone", "Tablet", "Monitor", "Keyboard"];

    let mut row_count = 0;
    for month in &months {
        for product in &products {
            let sales = (row_count * 1000 + 5000) as f64;
            let profit = sales * 0.25;

            let sales_str = format!("{:.2}", sales);
            let profit_str = format!("{:.2}", profit);

            writer
                .write_row([*month, *product, &sales_str, &profit_str])
                .await?;

            row_count += 1;
        }
    }

    println!("✅ Wrote {} rows\n", row_count);

    // Upload to S3
    println!("☁️  Streaming to S3 (completing multipart upload)...");
    writer.save().await?;

    println!("✅ Upload complete!\n");
    println!("🎉 File available at: s3://{}/{}", bucket, key);
    println!("\n💡 Features:");
    println!("   ✅ ZERO disk usage (no temp files!)");
    println!("   ✅ Constant ~4 KB memory for buffering");
    println!("   ✅ Uses s-zip 0.8.0 cloud support");
    println!("   ✅ Async streaming with tokio");
    println!("\n🔍 Verify with:");
    println!("   aws s3 ls s3://{}/{}", bucket, key);
    println!(
        "   aws s3 cp s3://{}/{} . && unzip -l {}",
        bucket,
        key,
        key.split('/').next_back().unwrap()
    );

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("❌ This example requires the 'cloud-s3' feature.");
    eprintln!("\nRun with:");
    eprintln!("  cargo run --example s3_streaming --features cloud-s3");
    eprintln!("\nMake sure you have AWS credentials configured:");
    eprintln!("  export AWS_ACCESS_KEY_ID=your_key");
    eprintln!("  export AWS_SECRET_ACCESS_KEY=your_secret");
    eprintln!("  export AWS_S3_BUCKET=your-bucket-name");
    std::process::exit(1);
}
