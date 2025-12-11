//! List files in S3 bucket to find existing files

#[cfg(feature = "cloud-s3")]
#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    use aws_config;
    use aws_sdk_s3;

    let bucket = "lune-nonprod";
    let region = "ap-southeast-1";

    println!("📦 Listing files in bucket: {}\n", bucket);

    // Initialize AWS SDK
    let region_provider = aws_sdk_s3::config::Region::new(region.to_string());
    let config = aws_config::defaults(aws_config::BehaviorVersion::latest())
        .region(region_provider)
        .load()
        .await;
    let s3_client = aws_sdk_s3::Client::new(&config);

    // List objects
    let resp = s3_client
        .list_objects_v2()
        .bucket(bucket)
        .prefix("") // List all
        .max_keys(100)
        .send()
        .await?;

    let contents = resp.contents();
    if !contents.is_empty() {
        println!("Found {} objects:\n", contents.len());
        for obj in contents {
            if let (Some(key), Some(size)) = (obj.key(), obj.size()) {
                println!("  📄 {} ({} bytes)", key, size);
            }
        }
    } else {
        println!("No objects found");
    }

    Ok(())
}

#[cfg(not(feature = "cloud-s3"))]
fn main() {
    eprintln!("This example requires the 'cloud-s3' feature");
    std::process::exit(1);
}
