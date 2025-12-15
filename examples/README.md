# ExcelStream Examples

**15 carefully curated examples** demonstrating ExcelStream's core features.

## 📚 Getting Started (4 examples)

### Basic Operations
- **`basic_read.rs`** - Read Excel files, iterate rows and cells
- **`basic_write.rs`** - Write Excel files with simple API
- **`streaming_read.rs`** - Stream large files with constant memory
- **`streaming_write.rs`** - Write large files with constant memory

```bash
# Quick start
cargo run --example basic_write
cargo run --example basic_read
cargo run --example streaming_write --release
cargo run --example streaming_read --release
```

## 🎨 Common Features (5 examples)

### Formatting & Layout
- **`multi_sheet.rs`** - Multiple worksheets in one workbook
- **`cell_formatting.rs`** - Cell styles, colors, borders, number formats
- **`column_width_row_height.rs`** - Set column widths and row heights
- **`worksheet_protection.rs`** - Password protect worksheets
- **`csv_to_excel.rs`** - Convert CSV to XLSX

```bash
cargo run --example multi_sheet
cargo run --example cell_formatting
cargo run --example worksheet_protection
```

## ⚡ Performance & Benchmarks (3 examples)

### Memory & Speed Tests
- **`memory_benchmark.rs`** - Write benchmark (1M rows)
- **`memory_benchmark_read.rs`** - Read benchmark (1M rows)
- **`writers_comparison.rs`** - Compare write_row() vs write_row_typed() vs styled()

```bash
# WARNING: These create large files (180+ MB each)
cargo run --example memory_benchmark --release
cargo run --example memory_benchmark_read --release
cargo run --example writers_comparison --release
```

**Expected Results:**
- Write: 40K-45K rows/sec
- Read: 35K-40K rows/sec  
- Memory: 2-3 MB constant
- File size: ~180 MB (1M rows × 30 cols)

## ☁️ Cloud & Database (2 examples)

### Advanced Integrations
- **`s3_streaming.rs`** - Direct streaming to/from AWS S3
- **`postgres_to_excel_advanced.rs`** - Export PostgreSQL to Excel

```bash
# Requires optional features
cargo run --example s3_streaming --features cloud-s3
cargo run --example postgres_to_excel_advanced --features postgres-async
```

## 🚀 Advanced Features (1 example)

### v0.10+ Features
- **`incremental_append.rs`** - Append rows to existing files (10-100x faster)

```bash
cargo run --example incremental_append
```

## 📊 Performance Summary

Based on `writers_comparison.rs` (1M rows × 30 columns):

| Method | Speed | Use Case |
|--------|-------|----------|
| write_row() | 42,557 rows/sec | Simple text export |
| write_row_typed() | 36,178 rows/sec | Excel formulas, calculations |
| write_row_styled() | 33,044 rows/sec | Formatted reports |
| UltraLowMemoryWorkbook | 43,839 rows/sec | Lowest memory (2-3 MB) |

All methods maintain **constant 2-3 MB memory** regardless of file size.

## 🗂️ Example Categories

```
examples/
├── Basic (4)           - Getting started, essential operations
├── Features (5)        - Common use cases (styling, multi-sheet, etc.)
├── Benchmarks (3)      - Performance testing and comparison
├── Cloud (2)           - S3, PostgreSQL integrations
└── Advanced (1)        - Incremental append (v0.10+)
```

## 💡 Quick Tips

**For Learning:**
1. Start with `basic_write.rs` and `basic_read.rs`
2. Then try `streaming_write.rs` for large files
3. Explore `cell_formatting.rs` for styled output

**For Testing Performance:**
1. Run `memory_benchmark.rs` to verify write performance
2. Run `memory_benchmark_read.rs` to verify read performance
3. Check memory usage with `/usr/bin/time -v` on Linux

**For Production:**
1. Use `streaming_write.rs` pattern for large datasets
2. Use `write_row_typed()` if users need Excel formulas
3. Use `incremental_append.rs` for log-style data

## 📝 Notes

- All benchmarks should be run with `--release` flag
- Large benchmarks create 180+ MB files in current directory
- Cloud examples require AWS credentials or PostgreSQL setup
- See main README.md for detailed feature documentation

## 🔗 Related

- [Main Documentation](../README.md)
- [CHANGELOG](../CHANGELOG.md)
- [Performance Results](../PERFORMANCE_RESULTS_V0.8.0.md)
