# Release Notes - v0.9.0

## 🎉 Major Feature: Zero-Temp Streaming Architecture

**84% Memory Reduction** - Revolutionary streaming write implementation!

### Breakthrough Performance

| Metric | v0.8.0 (temp files) | v0.9.0 (streaming) | Improvement |
|--------|---------------------|-------------------|-------------|
| **Memory** | 17 MB | **2.7 MB** | **84% reduction** ✅ |
| **Speed** | ~1400ms | ~1500ms | Comparable |
| **File Size** | 15 MB | 16 MB | +7% (acceptable) |
| **Temp Files** | Yes | **No** ✅ | Zero disk I/O |

### What's New?

#### ZeroTempWorkbook API

Stream XML directly into ZIP compressor - no intermediate files!

```rust
use excelstream::fast_writer::ZeroTempWorkbook;

// Create workbook with compression level 6
let mut writer = ZeroTempWorkbook::new("output.xlsx", 6)?;

// Add worksheet
writer.add_worksheet("Sales")?;

// Write millions of rows with only 2.7 MB RAM!
for i in 0..10_000_000 {
    writer.write_row(&[
        &i.to_string(),
        &format!("Product {}", i),
        &(i as f64 * 99.99).to_string(),
    ])?;
}

// Close and finalize
writer.close()?;
```

### Technical Architecture

#### StreamingZipWriter
- Custom ZIP writer with on-the-fly compression
- Data descriptor mode (bit 3) - no seeking needed
- Writes CRC32/sizes after compressed data
- Pure streaming - no temp storage

#### ZeroTempWorkbook
- Streams XML generation directly into compressor
- 4KB XML buffer (reused per row)
- Minimal allocations
- Direct write to final .xlsx file

#### Memory Flow
```
Row Data → 4KB XML Buffer → DEFLATE Encoder → ZIP Entry → Final .xlsx
                ↑                                              
                └─── Reused (no accumulation!)
```

### Performance Validation

Tested with:
- ✅ 1 million rows: **2.75 MB RAM**, 1506ms
- ✅ 10 million rows: **2.7 MB RAM**, ~15 seconds
- ✅ File validity: Verified with Excel and `unzip -t`
- ✅ VmHWM cross-verification: Consistent across tools

### Migration Guide

#### Option A: Use New ZeroTempWorkbook (Recommended)
```rust
// Old (v0.8.0) - 17 MB RAM
use excelstream::fast_writer::UltraLowMemoryWorkbook;
let mut wb = UltraLowMemoryWorkbook::new("output.xlsx")?;

// New (v0.9.0) - 2.7 MB RAM
use excelstream::fast_writer::ZeroTempWorkbook;
let mut wb = ZeroTempWorkbook::new("output.xlsx", 6)?; // level 6 = balanced
```

#### Option B: Keep Using UltraLowMemoryWorkbook
- Still available for backward compatibility
- Uses temp files (17 MB RAM)
- More mature, more features
- Choose if you need stability over memory savings

### Trade-offs

| Aspect | v0.9.0 Streaming | v0.8.0 Temp Files |
|--------|------------------|-------------------|
| Memory | **2.7 MB** ✅ | 17 MB |
| Temp Files | None ✅ | Yes (~file size) |
| File Size | 16 MB | 15 MB |
| Speed | ~1500ms | ~1400ms |
| Maturity | New | Stable |
| Features | Basic | Advanced |

### When to Use What?

**Use ZeroTempWorkbook (v0.9.0) if:**
- Memory is critical (< 128 MB pods)
- Processing millions of rows
- No disk space for temp files
- Simple write operations (no complex formatting)

**Use UltraLowMemoryWorkbook (v0.8.0) if:**
- Need advanced features (cell styling, formulas)
- File size matters more than memory
- Proven stability required
- Memory budget > 20 MB

### What's Next?

v0.10.0 planned features:
- Add cell styling to ZeroTempWorkbook
- Formula support in streaming mode
- Multi-sheet parallel writing
- Auto-compression level selection
- Benchmark against other Rust libraries

### Breaking Changes

None! This is a pure addition:
- `ZeroTempWorkbook` is a new API
- `UltraLowMemoryWorkbook` unchanged
- All existing code continues to work

### Acknowledgments

This breakthrough was achieved through:
1. Recognizing that algorithm optimization hit a ceiling (17 MB)
2. Pivoting to architectural solution (eliminate temp files)
3. Implementing custom streaming ZIP writer
4. Validating with real-world datasets

The key insight: **Sometimes the best optimization is removing work entirely, not doing work faster.**

---

**Full Changelog**: See `CHANGELOG.md` for complete details.

**Documentation**: Updated `README.md` with v0.9.0 examples.

**Try it now**:
```bash
cargo update
cargo add excelstream@0.9.0
```
