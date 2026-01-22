//! Multi-Cloud S3 Configuration Example
//!
//! This example demonstrates how to use S3ExcelWriter with different cloud providers
//! using explicit credentials via AWS SDK Client.
//!
//! Supported providers:
//! - AWS S3
//! - MinIO (self-hosted S3-compatible)
//! - Cloudflare R2
//! - DigitalOcean Spaces
//! - Backblaze B2
//! - Any S3-compatible storage
//!
//! Set environment variables for each provider:
//!
//! AWS S3:
//! export AWS_ACCESS_KEY_ID=xxx
//! export AWS_SECRET_ACCESS_KEY=xxx
//! export AWS_REGION=us-east-1
//! export AWS_BUCKET=my-bucket
//!
//! MinIO:
//! export MINIO_ACCESS_KEY=minioadmin
//! export MINIO_SECRET_KEY=minioadmin
//! export MINIO_ENDPOINT=http://localhost:9000
//! export MINIO_BUCKET=test-bucket
//!
//! Cloudflare R2:
//! export R2_ACCESS_KEY=xxx
//! export R2_SECRET_KEY=xxx
//! export R2_ENDPOINT=https://xxx.r2.cloudflarestorage.com
//! export R2_BUCKET=my-bucket
//!
//! Run:
//! cargo run --example multi_cloud_config --features cloud-s3

use aws_sdk_s3::{config::Credentials, Client};
use excelstream::cloud::{S3ExcelReader, S3ExcelWriter};
use excelstream::error::Result;
use s_zip::cloud::S3ZipWriter;

// ============================================================================
// Helper Functions - Create S3 Clients with Explicit Credentials
// ============================================================================

async fn create_aws_client(access_key: &str, secret_key: &str, region: &str) -> Client {
    let creds = Credentials::new(access_key, secret_key, None, None, "aws");
    let config = aws_sdk_s3::Config::builder()
        .credentials_provider(creds)
        .region(aws_sdk_s3::config::Region::new(region.to_string()))
        .build();

    Client::from_conf(config)
}

async fn create_minio_client(access_key: &str, secret_key: &str, endpoint: &str) -> Client {
    let creds = Credentials::new(access_key, secret_key, None, None, "minio");
    let config = aws_sdk_s3::Config::builder()
        .credentials_provider(creds)
        .endpoint_url(endpoint)
        .region(aws_sdk_s3::config::Region::new("us-east-1"))
        .force_path_style(true) // Required for MinIO
        .build();

    Client::from_conf(config)
}

async fn create_r2_client(access_key: &str, secret_key: &str, endpoint: &str) -> Client {
    let creds = Credentials::new(access_key, secret_key, None, None, "r2");
    let config = aws_sdk_s3::Config::builder()
        .credentials_provider(creds)
        .endpoint_url(endpoint)
        .region(aws_sdk_s3::config::Region::new("auto"))
        .build();

    Client::from_conf(config)
}

async fn create_spaces_client(access_key: &str, secret_key: &str, region: &str) -> Client {
    let endpoint = format!("https://{}.digitaloceanspaces.com", region);
    let creds = Credentials::new(access_key, secret_key, None, None, "spaces");
    let config = aws_sdk_s3::Config::builder()
        .credentials_provider(creds)
        .endpoint_url(endpoint)
        .region(aws_sdk_s3::config::Region::new(region.to_string()))
        .build();

    Client::from_conf(config)
}

async fn create_b2_client(
    access_key: &str,
    secret_key: &str,
    endpoint: &str,
    region: &str,
) -> Client {
    let creds = Credentials::new(access_key, secret_key, None, None, "b2");
    let config = aws_sdk_s3::Config::builder()
        .credentials_provider(creds)
        .endpoint_url(endpoint)
        .region(aws_sdk_s3::config::Region::new(region.to_string()))
        .build();

    Client::from_conf(config)
}

// ============================================================================
// Main Examples
// ============================================================================

#[tokio::main]
async fn main() -> Result<()> {
    println!("🌐 Multi-Cloud S3 Configuration Examples\n");

    let mut tasks = vec![];

    // Try AWS S3
    if let (Ok(key), Ok(secret)) = (
        std::env::var("AWS_ACCESS_KEY_ID"),
        std::env::var("AWS_SECRET_ACCESS_KEY"),
    ) {
        println!("🌩️  AWS S3 configured");
        let region = std::env::var("AWS_REGION").unwrap_or("us-east-1".to_string());
        let bucket = std::env::var("AWS_BUCKET").unwrap_or("my-bucket".to_string());
        tasks.push(tokio::spawn(async move {
            upload_to_aws(&key, &secret, &region, &bucket).await
        }));
    } else {
        println!("⚠️  AWS S3 not configured (set AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY)");
    }

    // Try MinIO
    if let (Ok(key), Ok(secret)) = (
        std::env::var("MINIO_ACCESS_KEY"),
        std::env::var("MINIO_SECRET_KEY"),
    ) {
        println!("📦 MinIO configured");
        let endpoint =
            std::env::var("MINIO_ENDPOINT").unwrap_or("http://localhost:9000".to_string());
        let bucket = std::env::var("MINIO_BUCKET").unwrap_or("test-bucket".to_string());
        tasks.push(tokio::spawn(async move {
            upload_to_minio(&key, &secret, &endpoint, &bucket).await
        }));
    } else {
        println!("⚠️  MinIO not configured (set MINIO_ACCESS_KEY, MINIO_SECRET_KEY)");
    }

    // Try Cloudflare R2
    if let (Ok(key), Ok(secret), Ok(endpoint)) = (
        std::env::var("R2_ACCESS_KEY"),
        std::env::var("R2_SECRET_KEY"),
        std::env::var("R2_ENDPOINT"),
    ) {
        println!("☁️  Cloudflare R2 configured");
        let bucket = std::env::var("R2_BUCKET").unwrap_or("my-bucket".to_string());
        tasks.push(tokio::spawn(async move {
            upload_to_r2(&key, &secret, &endpoint, &bucket).await
        }));
    } else {
        println!(
            "⚠️  Cloudflare R2 not configured (set R2_ACCESS_KEY, R2_SECRET_KEY, R2_ENDPOINT)"
        );
    }

    if tasks.is_empty() {
        println!("\n❌ No cloud providers configured!");
        println!("\n💡 Set environment variables for at least one provider:");
        println!("   AWS: AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY");
        println!("   MinIO: MINIO_ACCESS_KEY, MINIO_SECRET_KEY");
        println!("   R2: R2_ACCESS_KEY, R2_SECRET_KEY, R2_ENDPOINT");
        return Ok(());
    }

    println!(
        "\n🚀 Uploading to {} provider(s) simultaneously...\n",
        tasks.len()
    );

    // Wait for all uploads to complete
    let mut results = Vec::new();
    for task in tasks {
        results.push(task.await);
    }

    let mut success_count = 0;
    let mut error_count = 0;

    for result in results {
        match result {
            Ok(Ok(())) => success_count += 1,
            Ok(Err(e)) => {
                error_count += 1;
                eprintln!("  ❌ Upload failed: {}", e);
            }
            Err(e) => {
                error_count += 1;
                eprintln!("  ❌ Task panicked: {}", e);
            }
        }
    }

    println!("\n📊 Summary:");
    println!("  ✅ Successful: {}", success_count);
    println!("  ❌ Failed: {}", error_count);

    Ok(())
}

// ============================================================================
// Upload Functions for Each Provider
// ============================================================================

async fn upload_to_aws(
    access_key: &str,
    secret_key: &str,
    region: &str,
    bucket: &str,
) -> Result<()> {
    println!("  📤 Uploading to AWS S3...");

    let client = create_aws_client(access_key, secret_key, region).await;
    let s3_writer = S3ZipWriter::new(client, bucket, "multi-cloud/aws-report.xlsx")
        .await
        .map_err(|e| {
            excelstream::error::ExcelError::IoError(std::io::Error::other(e.to_string()))
        })?;

    let mut writer = S3ExcelWriter::from_s3_writer(s3_writer);

    writer
        .write_header_bold(&["Provider", "Region", "Status"])
        .await?;
    writer.write_row(&["AWS S3", region, "Active"]).await?;
    writer
        .write_row(&["Multi-Cloud", "Global", "Enabled"])
        .await?;

    writer.save().await?;
    println!("  ✅ Successfully uploaded to AWS S3");

    Ok(())
}

async fn upload_to_minio(
    access_key: &str,
    secret_key: &str,
    endpoint: &str,
    bucket: &str,
) -> Result<()> {
    println!("  📤 Uploading to MinIO...");

    let client = create_minio_client(access_key, secret_key, endpoint).await;
    let s3_writer = S3ZipWriter::new(client, bucket, "multi-cloud/minio-report.xlsx")
        .await
        .map_err(|e| {
            excelstream::error::ExcelError::IoError(std::io::Error::other(e.to_string()))
        })?;

    let mut writer = S3ExcelWriter::from_s3_writer(s3_writer);

    writer
        .write_header_bold(&["Provider", "Endpoint", "Status"])
        .await?;
    writer.write_row(&["MinIO", endpoint, "Active"]).await?;
    writer
        .write_row(&["Self-Hosted", "Local", "Running"])
        .await?;

    writer.save().await?;
    println!("  ✅ Successfully uploaded to MinIO");

    Ok(())
}

async fn upload_to_r2(
    access_key: &str,
    secret_key: &str,
    endpoint: &str,
    bucket: &str,
) -> Result<()> {
    println!("  📤 Uploading to Cloudflare R2...");

    let client = create_r2_client(access_key, secret_key, endpoint).await;
    let s3_writer = S3ZipWriter::new(client, bucket, "multi-cloud/r2-report.xlsx")
        .await
        .map_err(|e| {
            excelstream::error::ExcelError::IoError(std::io::Error::other(e.to_string()))
        })?;

    let mut writer = S3ExcelWriter::from_s3_writer(s3_writer);

    writer
        .write_header_bold(&["Provider", "Feature", "Status"])
        .await?;
    writer
        .write_row(&["Cloudflare R2", "Zero Egress", "Enabled"])
        .await?;
    writer
        .write_row(&["Global Network", "Edge", "Active"])
        .await?;

    writer.save().await?;
    println!("  ✅ Successfully uploaded to Cloudflare R2");

    Ok(())
}

// ============================================================================
// Additional Example Functions
// ============================================================================

#[allow(dead_code)]
async fn example_digitalocean_spaces() -> Result<()> {
    let client = create_spaces_client("DO_KEY", "DO_SECRET", "nyc3").await;
    let s3_writer = S3ZipWriter::new(client, "my-spaces", "report.xlsx")
        .await
        .map_err(|e| {
            excelstream::error::ExcelError::IoError(std::io::Error::other(e.to_string()))
        })?;

    let mut writer = S3ExcelWriter::from_s3_writer(s3_writer);
    writer.write_header_bold(&["Service", "Region"]).await?;
    writer.write_row(&["DigitalOcean Spaces", "NYC3"]).await?;
    writer.save().await?;

    Ok(())
}

#[allow(dead_code)]
async fn example_backblaze_b2() -> Result<()> {
    let client = create_b2_client(
        "B2_KEY",
        "B2_SECRET",
        "https://s3.us-west-004.backblazeb2.com",
        "us-west-004",
    )
    .await;

    let s3_writer = S3ZipWriter::new(client, "my-b2-bucket", "backup.xlsx")
        .await
        .map_err(|e| {
            excelstream::error::ExcelError::IoError(std::io::Error::other(e.to_string()))
        })?;

    let mut writer = S3ExcelWriter::from_s3_writer(s3_writer);
    writer.write_header_bold(&["Service", "Cost"]).await?;
    writer.write_row(&["Backblaze B2", "Low"]).await?;
    writer.save().await?;

    Ok(())
}

// ============================================================================
// Reader Examples
// ============================================================================

#[allow(dead_code)]
async fn example_read_from_aws() -> Result<()> {
    let client = create_aws_client("ACCESS_KEY", "SECRET_KEY", "us-east-1").await;

    let mut reader = S3ExcelReader::from_s3_client(client, "my-bucket", "data.xlsx").await?;

    println!("Sheets: {:?}", reader.sheet_names());

    for row in reader.rows("Sheet1")? {
        let row = row?;
        println!("{:?}", row.to_strings());
    }

    Ok(())
}

#[allow(dead_code)]
async fn example_read_from_minio() -> Result<()> {
    let client = create_minio_client("minioadmin", "minioadmin", "http://localhost:9000").await;

    let mut reader = S3ExcelReader::from_s3_client(client, "test-bucket", "data.xlsx").await?;

    println!("Reading from MinIO...");
    for row in reader.rows("Sheet1")? {
        let row = row?;
        println!("{:?}", row.to_strings());
    }

    Ok(())
}

#[allow(dead_code)]
async fn example_read_from_r2() -> Result<()> {
    let client = create_r2_client(
        "R2_KEY",
        "R2_SECRET",
        "https://account.r2.cloudflarestorage.com",
    )
    .await;

    let mut reader = S3ExcelReader::from_s3_client(client, "my-r2-bucket", "report.xlsx").await?;

    println!("Reading from Cloudflare R2...");
    for row in reader.rows("Sheet1")? {
        let row = row?;
        println!("{:?}", row.to_strings());
    }

    Ok(())
}
