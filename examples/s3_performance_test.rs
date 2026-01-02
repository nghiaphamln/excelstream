//! Example: S3 Performance Test - Large dataset streaming
//!
//! This tests memory usage and performance when streaming large Excel files to S3.
//! Monitor with: /usr/bin/time -v to see peak memory usage

#[cfg(feature = "cloud-s3")]
use excelstream::cloud::S3ExcelWriter;

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("🚀 S3 Performance Test - Large Dataset Streaming\n");

    let bucket = std::env::var("AWS_S3_BUCKET").unwrap_or_else(|_| "lune-nonprod".to_string());
    let key = "reports/performance_test_large.xlsx";
    let region = std::env::var("AWS_REGION").unwrap_or_else(|_| "ap-southeast-1".to_string());

    // Test parameters
    let num_rows: u32 = std::env::var("TEST_ROWS")
        .ok()
        .and_then(|s| s.parse().ok())
        .unwrap_or(100_000); // 100K rows by default

    println!("📊 Test Configuration:");
    println!("   Target: s3://{}/{}", bucket, key);
    println!("   Region: {}", region);
    println!("   Rows: {}", num_rows);
    println!("   Columns: 10\n");

    let start = std::time::Instant::now();

    // Create S3 Excel writer
    println!("⏳ Initializing S3 writer...");
    let mut writer = S3ExcelWriter::builder()
        .bucket(&bucket)
        .key(key)
        .region(&region)
        .build()
        .await?;

    println!("✅ Writer initialized\n");

    // Write header
    println!("📝 Writing header...");
    writer
        .write_header_bold([
            "ID",
            "Name",
            "Email",
            "City",
            "Country",
            "Age",
            "Salary",
            "Department",
            "Join Date",
            "Status",
        ])
        .await?;

    // Write data in batches with progress
    println!("📊 Streaming {} rows to S3...", num_rows);
    let batch_size = 10_000;
    let mut rows_written = 0;

    for batch_start in (0..num_rows).step_by(batch_size as usize) {
        let batch_end = (batch_start + batch_size).min(num_rows);

        for i in batch_start..batch_end {
            let id = format!("{}", i + 1);
            let name = format!("User{}", i + 1);
            let email = format!("user{}@example.com", i + 1);
            let city = format!("City{}", (i % 100) + 1);
            let country = format!("Country{}", (i % 20) + 1);
            let age = format!("{}", 20 + (i % 50));
            let salary = format!("{:.2}", 30000.0 + (i as f64 * 100.0));
            let dept = format!("Dept{}", (i % 10) + 1);
            let date = format!("2024-{:02}-{:02}", (i % 12) + 1, (i % 28) + 1);
            let status = if i % 3 == 0 { "Active" } else { "Inactive" };

            writer
                .write_row([
                    id.as_str(),
                    name.as_str(),
                    email.as_str(),
                    city.as_str(),
                    country.as_str(),
                    age.as_str(),
                    salary.as_str(),
                    dept.as_str(),
                    date.as_str(),
                    status,
                ])
                .await?;

            rows_written += 1;
        }

        // Progress update
        let percent = (rows_written as f64 / num_rows as f64 * 100.0) as u32;
        println!(
            "   Progress: {}% ({}/{} rows)",
            percent, rows_written, num_rows
        );
    }

    println!("\n☁️  Finalizing upload (completing multipart)...");
    writer.save().await?;

    let elapsed = start.elapsed();

    println!("\n✅ Upload Complete!\n");
    println!("📈 Performance Results:");
    println!("   Total rows: {}", num_rows);
    println!("   Time: {:.2}s", elapsed.as_secs_f64());
    println!(
        "   Throughput: {:.0} rows/sec",
        num_rows as f64 / elapsed.as_secs_f64()
    );
    println!("\n💾 Memory Usage:");
    println!("   Expected peak: ~10-15 MB (constant, regardless of file size)");
    println!("   No temp files used!");
    println!("\n🔍 Verify with:");
    println!("   Run: cargo run --example s3_verify --features cloud-s3");

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("❌ This example requires the 'cloud-s3' feature.");
    eprintln!("\nRun with:");
    eprintln!(
        "  /usr/bin/time -v cargo run --example s3_performance_test --features cloud-s3 --release"
    );
    eprintln!("\nOptional environment variables:");
    eprintln!("  TEST_ROWS=100000  # Number of rows to generate (default: 100,000)");
    std::process::exit(1);
}
