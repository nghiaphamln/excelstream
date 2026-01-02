# excelstream

🦀 **High-performance streaming Excel & CSV library for Rust with constant memory usage**

[![Rust](https://img.shields.io/badge/rust-1.70%2B-orange.svg)](https://www.rust-lang.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/KSD-CO/excelstream/workflows/Rust/badge.svg)](https://github.com/KSD-CO/excelstream/actions)

## ✨ Highlights

- 📊 **XLSX & CSV Support** - Read/write Excel and CSV files
- 📉 **Constant Memory** - ~3-35 MB regardless of file size
- ☁️ **Cloud Streaming** - Direct S3 uploads with ZERO temp files
- ⚡ **High Performance** - 94K rows/sec (S3), 1.2M rows/sec (CSV)
- 🔄 **True Streaming** - Process files row-by-row, no buffering
- 🐳 **Production Ready** - Works in 256 MB containers

## 🔥 What's New in v0.14.0

**TRUE S3 Streaming** - Zero temp files, async API, constant memory!

```rust
use excelstream::cloud::S3ExcelWriter;

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = S3ExcelWriter::builder()
        .bucket("my-bucket")
        .key("report.xlsx")
        .region("us-east-1")
        .build()
        .await?;

    writer.write_header_bold(["Month", "Sales"]).await?;
    writer.write_row(["January", "50000"]).await?;
    writer.save().await?; // ✅ Streams directly to S3!
    Ok(())
}
```

**Performance:**
- **500K rows** → 34 MB peak memory, 94K rows/sec
- **ZERO temp files** → Works in read-only filesystems (Lambda!)
- **Breaking change:** S3 methods now async (add `.await`)

[See full changelog](CHANGELOG.md) | [Migration guide](#migration-from-v013)

---

## 📦 Quick Start

### Installation

```toml
[dependencies]
excelstream = "0.14"

# Optional features
excelstream = { version = "0.14", features = ["cloud-s3"] }  # S3 support
```

### Write Excel (Local)

```rust
use excelstream::ExcelWriter;

let mut writer = ExcelWriter::new("output.xlsx")?;

// Write 1M rows with only 3 MB memory!
writer.write_header_bold(&["ID", "Name", "Amount"])?;
for i in 1..=1_000_000 {
    writer.write_row(&[&i.to_string(), "Item", "1000"])?;
}
writer.save()?;
```

### Read Excel (Streaming)

```rust
use excelstream::ExcelReader;

let mut reader = ExcelReader::open("large.xlsx")?;

// Process 1 GB file with only 12 MB memory!
for row in reader.rows("Sheet1")? {
    let row = row?;
    println!("{:?}", row.to_strings());
}
```

### S3 Streaming (v0.14+)

```rust
use excelstream::cloud::S3ExcelWriter;

#[tokio::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut writer = S3ExcelWriter::builder()
        .bucket("reports")
        .key("sales.xlsx")
        .build()
        .await?;

    writer.write_header_bold(["Date", "Revenue"]).await?;
    writer.write_row(["2024-01-01", "125000"]).await?;
    writer.save().await?;  // Streams to S3, no disk!
    Ok(())
}
```

[More examples →](examples/)

---

## 🎯 Why ExcelStream?

**The Problem:** Traditional libraries load entire files into memory

```rust
// ❌ Traditional: 1 GB file = 1+ GB RAM (OOM in containers!)
let workbook = Workbook::new("huge.xlsx")?;
```

**The Solution:** True streaming with constant memory

```rust
// ✅ ExcelStream: 1 GB file = 12 MB RAM
let mut reader = ExcelReader::open("huge.xlsx")?;
for row in reader.rows("Sheet1")? { /* streaming! */ }
```

### Performance Comparison

| Operation | Traditional | ExcelStream | Improvement |
|-----------|-------------|-------------|-------------|
| Write 1M rows | 100+ MB | **2.7 MB** | **97% less memory** |
| Read 1GB file | ❌ Crash | **12 MB** | Works! |
| S3 upload 500K rows | Temp file | **34 MB** | **Zero disk** |
| K8s pod (256MB) | ❌ OOMKilled | ✅ Works | Production ready |

---

## ☁️ Cloud Features

### S3 Direct Streaming (v0.14)

Upload Excel files directly to S3 with **ZERO temp files**:

```bash
cargo add excelstream --features cloud-s3
```

**Performance (Real AWS S3):**

| Dataset | Memory | Throughput | Temp Files |
|---------|--------|------------|------------|
| 10K rows | 15 MB | 11K rows/s | **ZERO** ✅ |
| 100K rows | 23 MB | 45K rows/s | **ZERO** ✅ |
| 500K rows | 34 MB | 94K rows/s | **ZERO** ✅ |

Perfect for:
- ✅ AWS Lambda (read-only filesystem)
- ✅ Docker containers (no disk space)
- ✅ Kubernetes CronJobs (limited memory)

[See S3 performance details →](PERFORMANCE_S3.md)

### HTTP Streaming

Stream Excel files directly to web responses:

```rust
use excelstream::cloud::HttpExcelWriter;

async fn download() -> impl IntoResponse {
    let mut writer = HttpExcelWriter::new();
    writer.write_row(&["Data"])?;
    ([(header::CONTENT_TYPE, "application/vnd....")], writer.finish()?)
}
```

[HTTP streaming guide →](examples/http_streaming.rs)

---

## 📊 CSV Support

**13.5x faster** than Excel for CSV workloads:

```rust
use excelstream::csv::CsvWriter;

let mut writer = CsvWriter::new("data.csv")?;
writer.write_row(&["A", "B", "C"])?;  // 1.2M rows/sec!
writer.save()?;
```

**Features:**
- ✅ Zstd compression (`.csv.zst` - 2.9x smaller)
- ✅ Auto-detection (`.csv`, `.csv.gz`, `.csv.zst`)
- ✅ Streaming (< 5 MB memory)

[CSV examples →](examples/csv_write.rs)

---

## 🚀 Use Cases

### 1. Large File Processing

```rust
// Process 500 MB Excel with only 25 MB RAM
let mut reader = ExcelReader::open("customers.xlsx")?;
for row in reader.rows("Sales")? {
    // Process row-by-row, constant memory!
}
```

### 2. Database Exports

```rust
// Export 1M database rows to Excel
let mut writer = ExcelWriter::new("export.xlsx")?;
let rows = db.query("SELECT * FROM large_table")?;
for row in rows {
    writer.write_row(&[row.get(0), row.get(1)])?;
}
writer.save()?;  // Only 3 MB memory used!
```

### 3. Cloud Pipelines

```rust
// Lambda function: DB → Excel → S3
let mut writer = S3ExcelWriter::builder()
    .bucket("data-lake").key("export.xlsx").build().await?;

let rows = db.query_stream("SELECT * FROM events").await?;
while let Some(row) = rows.next().await {
    writer.write_row(row).await?;
}
writer.save().await?;  // No temp files, no disk!
```

---

## 📚 Documentation

- [API Docs](https://docs.rs/excelstream) - Full API reference
- [Examples](examples/) - Code examples for all features
- [CHANGELOG](CHANGELOG.md) - Version history
- [Performance](PERFORMANCE_S3.md) - Detailed benchmarks

### Key Topics

- [Excel Writing](examples/basic_write.rs) - Basic & advanced writing
- [Excel Reading](examples/basic_read.rs) - Streaming read
- [S3 Streaming](examples/s3_streaming.rs) - Cloud uploads
- [CSV Support](examples/csv_write.rs) - CSV operations
- [Styling](examples/cell_formatting.rs) - Cell formatting & colors

---

## 🔧 Features

| Feature | Description |
|---------|-------------|
| `default` | Core Excel/CSV with Zstd compression |
| `cloud-s3` | S3 direct streaming (async) |
| `cloud-http` | HTTP response streaming |
| `serde` | Serde serialization support |
| `parallel` | Parallel processing with Rayon |

---

## ⚡ Performance

**Memory Usage (Constant):**
- Excel write: **2.7 MB** (any size)
- Excel read: **10-12 MB** (any size)
- S3 streaming: **30-35 MB** (any size)
- CSV write: **< 5 MB** (any size)

**Throughput:**
- Excel write: 42K rows/sec
- Excel read: 50K rows/sec
- S3 streaming: 94K rows/sec
- CSV write: 1.2M rows/sec

---

## 🛠️ Migration from v0.13

**S3ExcelWriter** is now async:

```rust
// OLD (v0.13 - sync)
writer.write_row(&["a", "b"])?;

// NEW (v0.14 - async)
writer.write_row(["a", "b"]).await?;
```

All other APIs unchanged!

---

## 📋 Requirements

- Rust 1.70+
- Optional: AWS credentials for S3 features

---

## 🤝 Contributing

Contributions welcome! Please read [CONTRIBUTING.md](CONTRIBUTING.md).

---

## 📄 License

MIT License - See [LICENSE](LICENSE) for details

---

## 🙏 Credits

- Built with [s-zip](https://crates.io/crates/s-zip) for streaming ZIP
- AWS SDK for Rust
- All contributors and users!

---

**Need help?** [Open an issue](https://github.com/KSD-CO/excelstream/issues) | **Questions?** [Discussions](https://github.com/KSD-CO/excelstream/discussions)
