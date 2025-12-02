# ExcelStream Improvement Plan

This document outlines the planned improvements for the excelstream library based on comprehensive code review.

## Current Status (v0.2.0)

**Strengths:**
- ✅ Excellent performance (21-47% faster than rust_xlsxwriter)
- ✅ True streaming with constant ~80MB memory usage
- ✅ Good test coverage (18 unit tests + 7 integration tests)
- ✅ Comprehensive documentation
- ✅ Rich examples (21 examples)

**Areas for Improvement:**
- Code quality issues (11 clippy warnings)
- Missing features (formatting, formulas, cell merging)
- API ergonomics could be improved
- Some error handling improvements needed

---

## PHASE 1 - Immediate Fixes (v0.2.1)

**Target: Fix critical code quality and add basic missing features**

### 1.1 Code Quality Fixes ✓

- [x] Fix unused `mut` in [worksheet.rs:227](src/fast_writer/worksheet.rs#L227)
- [x] Fix needless borrow in [reader.rs:71,104,121](src/reader.rs)
- [x] Fix unnecessary cast in [reader.rs:141](src/reader.rs#L141)
- [x] Fix needless borrows in writer.rs tests
- [x] Fix PI constant usage in [writer.rs:367](src/writer.rs#L367)

### 1.2 Documentation Fixes ✓

- [x] Fix package name in [lib.rs:1](src/lib.rs#L1) (rust-excelize → excelstream)

### 1.3 Error Handling Cleanup ✓

- [x] Remove unused `XlsxWriterError` variant from [error.rs](src/error.rs)
- [x] Clean up outdated error documentation

### 1.4 Basic Formatting Support

**Priority: HIGH**

- [ ] Implement bold header formatting
  - Add `Format` struct with basic properties (bold, italic)
  - Modify FastWorkbook to support styles.xml generation
  - Update `write_header()` to apply bold formatting

- [ ] Implement column width support
  - Add column width tracking to FastWorksheet
  - Generate proper `<col>` elements in worksheet XML
  - Make `set_column_width()` functional (currently no-op)

### 1.5 Testing

- [x] Verify all clippy warnings are resolved
- [x] Run full test suite
- [ ] Add tests for new formatting features

**Estimated Time:** 2-4 hours
**Complexity:** Low-Medium

---

## PHASE 2 - Short Term (v0.2.2)

**Target: Essential Excel features**

### 2.1 Formula Support

```rust
pub enum CellValue {
    Formula(String),  // Add this variant
    // ... existing variants
}

impl ExcelWriter {
    pub fn write_formula(&mut self, col: u32, formula: &str) -> Result<()>;
}
```

### 2.2 Cell Merging

```rust
impl ExcelWriter {
    pub fn merge_range(&mut self, start_row: u32, start_col: u32,
                       end_row: u32, end_col: u32, content: &str) -> Result<()>;
}
```

### 2.3 Improved Error Messages

```rust
#[error("Sheet '{sheet}' not found. Available sheets: {available}")]
SheetNotFound { sheet: String, available: String },

#[error("Failed to write row {row} to sheet '{sheet}': {source}")]
WriteRowError {
    row: u32,
    sheet: String,
    source: Box<ExcelError>,
},
```

### 2.4 Additional Tests

- [ ] Edge case tests (empty strings, long strings, special characters)
- [ ] XML escaping tests
- [ ] Excel limits tests (max rows, max columns)
- [ ] Unicode sheet names tests

### 2.5 Dependency Updates

- [ ] Update calamine to latest version
- [ ] Review and update other dependencies

**Estimated Time:** 1 week
**Complexity:** Medium

---

## PHASE 3 - Medium Term (v0.3.0)

**Target: Advanced styling and performance**

### 3.1 Cell Formatting & Styling API

```rust
pub struct CellStyle {
    pub font: FontStyle,
    pub fill: Option<FillStyle>,
    pub border: Option<BorderStyle>,
    pub alignment: Option<Alignment>,
    pub number_format: Option<String>,
}

pub struct FontStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: Color,
    pub size: f64,
    pub name: String,
}

impl ExcelWriter {
    pub fn write_cell_with_style(&mut self, row: u32, col: u32,
                                  value: &CellValue, style: &CellStyle) -> Result<()>;
}
```

### 3.2 Parallel Reading Support

```rust
#[cfg(feature = "parallel")]
impl ExcelReader {
    pub fn rows_parallel(&mut self, sheet_name: &str) -> Result<ParRowIterator>;
}
```

### 3.3 Data Validation

```rust
pub enum DataValidation {
    List(Vec<String>),
    Integer { min: i64, max: i64 },
    Decimal { min: f64, max: f64 },
    Date { min: DateTime, max: DateTime },
    Custom(String),
}

impl ExcelWriter {
    pub fn add_data_validation(&mut self, range: Range,
                                validation: DataValidation) -> Result<()>;
}
```

### 3.4 Ergonomic API Improvements

```rust
// Macro for easy row creation
#[macro_export]
macro_rules! row {
    ($($val:expr),* $(,)?) => {
        vec![$(CellValue::from($val)),*]
    };
}

// Builder pattern for CellValue
impl CellValue {
    pub fn string(s: impl Into<String>) -> Self;
    pub fn int(i: impl Into<i64>) -> Self;
    pub fn float(f: impl Into<f64>) -> Self;
}

// Iterator-based batch operations
pub fn write_rows_typed_iter<I>(&mut self, rows: I) -> Result<()>
where
    I: Iterator<Item = Vec<CellValue>>;
```

### 3.5 Performance Optimizations

- [ ] Pre-allocated string buffers in XML writer
- [ ] Buffer reuse to reduce allocations
- [ ] Benchmark and profile critical paths

**Estimated Time:** 3-4 weeks
**Complexity:** High

---

## PHASE 4 - Long Term (v0.4.0+)

**Target: Advanced Excel features**

### 4.1 Conditional Formatting

```rust
pub enum ConditionalFormat {
    ColorScale {
        min_color: Color,
        mid_color: Option<Color>,
        max_color: Color,
    },
    DataBar {
        color: Color,
        show_value: bool,
    },
    IconSet {
        icons: IconSetType,
        reverse: bool,
    },
    CellValue {
        operator: ComparisonOperator,
        value: CellValue,
        format: CellStyle,
    },
}

impl ExcelWriter {
    pub fn add_conditional_format(&mut self, range: &str,
                                   format: ConditionalFormat) -> Result<()>;
}
```

### 4.2 Charts

```rust
pub enum ChartType {
    Line,
    Column,
    Bar,
    Pie,
    Scatter,
    Area,
}

pub struct Chart {
    chart_type: ChartType,
    series: Vec<ChartSeries>,
    title: Option<String>,
    x_axis: AxisOptions,
    y_axis: AxisOptions,
}

impl ExcelWriter {
    pub fn insert_chart(&mut self, sheet: &str, row: u32, col: u32,
                        chart: &Chart) -> Result<()>;
}
```

### 4.3 Images

```rust
impl ExcelWriter {
    pub fn insert_image(&mut self, sheet: &str, row: u32, col: u32,
                        path: &str) -> Result<()>;
    pub fn insert_image_with_options(&mut self, sheet: &str, row: u32, col: u32,
                                      path: &str, options: ImageOptions) -> Result<()>;
}
```

### 4.4 Rich Text

```rust
pub struct RichText {
    runs: Vec<TextRun>,
}

pub struct TextRun {
    text: String,
    font: FontStyle,
}

impl ExcelWriter {
    pub fn write_rich_text(&mut self, row: u32, col: u32,
                           rich_text: &RichText) -> Result<()>;
}
```

### 4.5 Worksheet Protection

```rust
pub struct ProtectionOptions {
    pub password: Option<String>,
    pub select_locked_cells: bool,
    pub select_unlocked_cells: bool,
    pub format_cells: bool,
    pub format_columns: bool,
    pub format_rows: bool,
}

impl ExcelWriter {
    pub fn protect_sheet(&mut self, options: ProtectionOptions) -> Result<()>;
}
```

**Estimated Time:** 8-12 weeks
**Complexity:** Very High

---

## PHASE 5 - Repository & Publishing

### 5.1 CI/CD Setup

```yaml
# .github/workflows/ci.yml
- Automated testing on push/PR
- Clippy checks
- Format checks
- Benchmark tracking
- Documentation deployment
```

### 5.2 Additional Badges

```markdown
[![Crates.io](https://img.shields.io/crates/v/excelstream.svg)]
[![Documentation](https://docs.rs/excelstream/badge.svg)]
[![Downloads](https://img.shields.io/crates/d/excelstream.svg)]
[![CI](https://github.com/KSD-CO/excelstream/workflows/CI/badge.svg)]
```

### 5.3 Documentation Improvements

- [ ] Create CHANGELOG.md
- [ ] Add CONTRIBUTING.md guidelines
- [ ] API documentation examples
- [ ] Migration guides for major versions
- [ ] Performance tuning guide

### 5.4 Community

- [ ] Set up issue templates
- [ ] PR templates
- [ ] Code of conduct
- [ ] Security policy

**Estimated Time:** 1-2 weeks
**Complexity:** Low

---

## Testing Strategy

### Unit Tests
- Test each module independently
- Cover edge cases and error conditions
- Test public APIs

### Integration Tests
- Test full read/write workflows
- Test multi-sheet operations
- Test large dataset handling

### Property-Based Tests
```rust
use proptest::prelude::*;

proptest! {
    #[test]
    fn roundtrip_arbitrary_data(rows: Vec<Vec<String>>) {
        // Write and read back, should match
    }
}
```

### Performance Tests
- Benchmark critical operations
- Memory usage tests
- Streaming validation tests

### Compatibility Tests
- Test Excel compatibility
- Test LibreOffice compatibility
- Test different Excel versions

---

## Performance Goals

### Current Performance (v0.2.0)
- ExcelWriter.write_row(): 36,870 rows/s
- ExcelWriter.write_row_typed(): 42,877 rows/s
- FastWorkbook direct: 44,753 rows/s
- Memory: ~80MB constant

### Target Performance (v0.3.0+)
- Maintain or improve write speeds
- Keep memory usage under 100MB for streaming
- Parallel reading: 2-4x speedup on multi-core systems
- Zero-copy optimizations where possible

---

## Breaking Changes Policy

### Semantic Versioning
- Patch (0.2.x): Bug fixes, no API changes
- Minor (0.x.0): New features, backward compatible
- Major (x.0.0): Breaking API changes

### Deprecation Strategy
- Deprecate old APIs in minor version
- Keep deprecated APIs for at least one minor version
- Document migration path clearly
- Remove in next major version

---

## Success Metrics

### Code Quality
- Zero clippy warnings with `-D warnings`
- Test coverage > 80%
- All examples working
- Documentation for all public APIs

### Performance
- Faster than rust_xlsxwriter for all operations
- Memory usage stays constant for streaming
- No performance regressions

### Community
- GitHub stars growth
- crates.io downloads
- Issue response time < 48 hours
- Regular releases (monthly for active development)

---

## Dependencies Philosophy

### Core Dependencies (minimal)
- calamine: Excel reading
- zip: ZIP compression
- thiserror: Error handling

### Optional Dependencies
- serde: Serialization support
- rayon: Parallel processing
- chrono: Date/time handling (for examples)

### Dev Dependencies
- tempfile: Testing
- criterion: Benchmarking
- proptest: Property-based testing
- rust_xlsxwriter: Comparison benchmarks only

---

## Notes

- Maintain backward compatibility within minor versions
- Keep streaming as the core feature
- Performance is a key differentiator
- Memory efficiency is non-negotiable
- Excel compatibility must be validated
- Documentation is as important as code

---

**Last Updated:** 2024-12-02
**Version:** 0.2.0
**Next Milestone:** v0.2.1 (Phase 1 completion)
