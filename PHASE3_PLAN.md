# Phase 3: Cell Formatting & Styling
**Version**: v0.3.0
**Status**: Planning
**Date**: 2024-12-02

## Overview
Add cell formatting capabilities to excelstream while maintaining streaming performance and constant memory usage.

## Current State Analysis

### Existing Implementation
- **FastWorkbook**: Streaming writer with ~80MB constant memory
- **SharedStrings**: String deduplication system in place
- **Basic styles.xml**: Minimal default styles (1 font, 2 fills, 1 border)
- **Cell writing**: Supports typed values (Int, Float, Bool, String, Formula)
- **No style support**: All cells use default formatting (style index 0)

### Key Files
- `src/fast_writer/workbook.rs` - Lines 387-410: write_styles() method
- `src/fast_writer/worksheet.rs` - Lines 114-208: write_row_typed() method
- `src/writer.rs` - Lines 145-163: ExcelWriter.write_row_typed()
- `src/types.rs` - CellValue enum definition

## Excel Styles Format

### How Excel Stores Styles (styles.xml)
```xml
<styleSheet>
  <!-- Number formats: currency, percentage, dates -->
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="$#,##0.00"/>
  </numFmts>

  <!-- Font definitions: bold, italic, size, color -->
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>  <!-- Bold -->
    <font><i/><sz val="11"/><name val="Calibri"/></font>  <!-- Italic -->
  </fonts>

  <!-- Fill patterns: background colors -->
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
  </fills>

  <!-- Border styles -->
  <borders count="2">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
  </borders>

  <!-- Cell formats (combinations of fonts, fills, borders) -->
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>  <!-- Default -->
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>  <!-- Bold -->
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>  <!-- Red bg -->
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- Currency -->
  </cellXfs>
</styleSheet>
```

### Cell Style Reference
Cells reference a cellXfs index via the "s" attribute:
```xml
<c r="A1" s="1"><v>Hello</v></c>  <!-- Uses cellXfs[1] = bold -->
<c r="B1" s="2"><v>World</v></c>  <!-- Uses cellXfs[2] = red background -->
```

## Implementation Strategy

### Approach: Pre-defined Styles (Recommended for v0.3.0)

**Rationale:**
- ✅ Simpler implementation
- ✅ Predictable memory usage (fixed number of styles)
- ✅ Covers 80% of use cases
- ✅ Fast - no dynamic style tracking needed
- ✅ Easy to extend later

**Trade-offs:**
- ❌ Limited flexibility (fixed set of styles)
- ❌ Can't create arbitrary color combinations
- ✅ BUT: Can be extended in v0.3.1+ with dynamic styles

### Pre-defined Style Set

```rust
pub enum CellStyle {
    Default,           // Index 0: No formatting
    HeaderBold,        // Index 1: Bold text (for headers)
    NumberInteger,     // Index 2: Integer format with thousand separator
    NumberDecimal,     // Index 3: 2 decimal places
    NumberCurrency,    // Index 4: $#,##0.00
    NumberPercentage,  // Index 5: 0.00%
    DateDefault,       // Index 6: MM/DD/YYYY
    DateTimestamp,     // Index 7: MM/DD/YYYY HH:MM:SS
    TextBold,          // Index 8: Bold for emphasis
    TextItalic,        // Index 9: Italic for notes
    HighlightYellow,   // Index 10: Yellow background
    HighlightGreen,    // Index 11: Green background
    HighlightRed,      // Index 12: Red background
    BorderThin,        // Index 13: Thin borders all sides
}
```

Total: 14 predefined styles (including default)

### API Design

#### Option 1: Separate write method with styles
```rust
writer.write_row_styled(&[
    (CellValue::String("Total".into()), CellStyle::HeaderBold),
    (CellValue::Float(1234.56), CellStyle::NumberCurrency),
    (CellValue::Int(95), CellStyle::NumberPercentage),
])?;
```

#### Option 2: Style embedded in CellValue (NOT RECOMMENDED)
```rust
// BAD: Mixes data with presentation
CellValue::StyledString("Total".into(), CellStyle::HeaderBold)
```

#### Option 3: Dedicated header method
```rust
writer.write_header_bold(&["Name", "Amount", "Percent"])?;
writer.write_row_typed(&[
    CellValue::String("Alice".into()),
    CellValue::Float(1234.56),
    CellValue::Int(95),
])?;
```

**Recommendation**: **Option 1** (write_row_styled) + **Option 3** (convenience methods)
- Option 1: Full flexibility
- Option 3: Ergonomic for common cases

### Implementation Plan

#### Step 1: Create Style Types (types.rs)
```rust
/// Cell style presets
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellStyle {
    Default,
    HeaderBold,
    NumberInteger,
    NumberDecimal,
    NumberCurrency,
    NumberPercentage,
    DateDefault,
    DateTimestamp,
    TextBold,
    TextItalic,
    HighlightYellow,
    HighlightGreen,
    HighlightRed,
    BorderThin,
}

impl CellStyle {
    /// Get the style index for XML
    pub fn index(&self) -> u32 {
        *self as u32
    }
}

/// Styled cell value (value + style)
#[derive(Debug, Clone)]
pub struct StyledCell {
    pub value: CellValue,
    pub style: CellStyle,
}

impl StyledCell {
    pub fn new(value: CellValue, style: CellStyle) -> Self {
        StyledCell { value, style }
    }

    pub fn default_style(value: CellValue) -> Self {
        StyledCell { value, style: CellStyle::Default }
    }
}
```

#### Step 2: Update worksheet.rs - Add write_row_styled
```rust
impl<W: Write> FastWorksheet<W> {
    pub fn write_row_styled(&mut self, cells: &[StyledCell]) -> Result<()> {
        self.cell_ref.next_row();
        self.row_count += 1;

        self.xml_writer.start_element("row")?;
        self.xml_writer.attribute_int("r", self.row_count as i64)?;
        self.xml_writer.close_start_tag()?;

        for cell in cells {
            let cell_ref = self.cell_ref.next_cell();
            let style_index = cell.style.index();

            match &cell.value {
                CellValue::String(s) => {
                    let string_index = self.shared_strings.add_string(s);
                    self.xml_writer.start_element("c")?;
                    self.xml_writer.attribute("r", &cell_ref)?;
                    self.xml_writer.attribute("t", "s")?;
                    if style_index > 0 {
                        self.xml_writer.attribute_int("s", style_index as i64)?;
                    }
                    // ... write value ...
                }
                CellValue::Int(n) => {
                    self.xml_writer.start_element("c")?;
                    self.xml_writer.attribute("r", &cell_ref)?;
                    if style_index > 0 {
                        self.xml_writer.attribute_int("s", style_index as i64)?;
                    }
                    // ... write value ...
                }
                // ... handle other types ...
            }
        }

        self.xml_writer.end_element("row")?;
        Ok(())
    }
}
```

#### Step 3: Generate Complete styles.xml (workbook.rs)
Replace the write_styles() method with full implementation:

```rust
fn write_styles(&mut self) -> Result<()> {
    let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

<!-- Number Formats -->
<numFmts count="5">
  <numFmt numFmtId="164" formatCode="#,##0"/>
  <numFmt numFmtId="165" formatCode="#,##0.00"/>
  <numFmt numFmtId="166" formatCode="$#,##0.00"/>
  <numFmt numFmtId="167" formatCode="0.00%"/>
  <numFmt numFmtId="168" formatCode="MM/DD/YYYY HH:MM:SS"/>
</numFmts>

<!-- Fonts -->
<fonts count="3">
  <font><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/></font>
  <font><i/><sz val="11"/><name val="Calibri"/></font>
</fonts>

<!-- Fills -->
<fills count="5">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
</fills>

<!-- Borders -->
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color auto="1"/></left>
    <right style="thin"><color auto="1"/></right>
    <top style="thin"><color auto="1"/></top>
    <bottom style="thin"><color auto="1"/></bottom>
  </border>
</borders>

<cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>

<!-- Cell Formats -->
<cellXfs count="14">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>  <!-- 0: Default -->
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>  <!-- 1: HeaderBold -->
  <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 2: NumberInteger -->
  <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 3: NumberDecimal -->
  <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 4: NumberCurrency -->
  <xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 5: NumberPercentage -->
  <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 6: DateDefault (MM/DD/YYYY) -->
  <xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>  <!-- 7: DateTimestamp -->
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>  <!-- 8: TextBold -->
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>  <!-- 9: TextItalic -->
  <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>  <!-- 10: HighlightYellow -->
  <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>  <!-- 11: HighlightGreen -->
  <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>  <!-- 12: HighlightRed -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>  <!-- 13: BorderThin -->
</cellXfs>

</styleSheet>"#;
    self.zip.write_all(xml.as_bytes())?;
    Ok(())
}
```

#### Step 4: Update ExcelWriter API (writer.rs)
```rust
impl ExcelWriter {
    /// Write row with styled cells
    pub fn write_row_styled(&mut self, cells: &[(CellValue, CellStyle)]) -> Result<()> {
        let styled_cells: Vec<StyledCell> = cells
            .iter()
            .map(|(value, style)| StyledCell::new(value.clone(), *style))
            .collect();
        self.inner.write_row_styled(&styled_cells)?;
        self.current_row += 1;
        Ok(())
    }

    /// Write header row with bold formatting
    pub fn write_header_bold<I, S>(&mut self, headers: I) -> Result<()>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        let cells: Vec<_> = headers
            .into_iter()
            .map(|h| (CellValue::String(h.as_ref().to_string()), CellStyle::HeaderBold))
            .collect();
        self.write_row_styled(&cells)
    }

    /// Convenience: Write row with all cells using same style
    pub fn write_row_with_style(&mut self, values: &[CellValue], style: CellStyle) -> Result<()> {
        let cells: Vec<_> = values
            .iter()
            .map(|v| (v.clone(), style))
            .collect();
        self.write_row_styled(&cells)
    }
}
```

#### Step 5: Update workbook.rs write_row for styled cells
Need to refactor workbook.rs to support styled cells in the direct write_row method.

### Implementation Steps Summary

1. ✅ **types.rs**: Add CellStyle enum and StyledCell struct
2. ✅ **fast_writer/worksheet.rs**: Add write_row_styled() method
3. ✅ **fast_writer/workbook.rs**:
   - Update write_styles() with full style definitions
   - Add write_row_styled() wrapper
4. ✅ **writer.rs**: Add convenience methods (write_row_styled, write_header_bold)
5. ✅ **tests**: Add comprehensive tests for all styles
6. ✅ **examples**: Create styling examples
7. ✅ **docs**: Update README with styling guide

## Testing Strategy

### Unit Tests
```rust
#[test]
fn test_cell_style_indices() {
    assert_eq!(CellStyle::Default.index(), 0);
    assert_eq!(CellStyle::HeaderBold.index(), 1);
    assert_eq!(CellStyle::NumberCurrency.index(), 4);
}

#[test]
fn test_write_styled_cells() {
    let mut workbook = FastWorkbook::new("test.xlsx")?;
    workbook.add_worksheet("Test")?;

    workbook.write_row_styled(&[
        StyledCell::new(CellValue::String("Total".into()), CellStyle::HeaderBold),
        StyledCell::new(CellValue::Float(1234.56), CellStyle::NumberCurrency),
    ])?;

    workbook.close()?;
    // Verify file opens in Excel and formatting is correct
}
```

### Integration Tests
- Write file with all 14 styles, open in Excel, verify appearance
- Test mixing styled and unstyled rows
- Test large file with styled cells (1M rows) - verify memory stays ~80MB
- Test header_bold convenience method

### Performance Tests
- Benchmark: Write 1M rows with styles vs without styles
- Expected: < 5% performance degradation
- Memory: Should stay ~80MB constant

## Documentation Updates

### README.md additions
```markdown
### Cell Formatting

#### Bold Headers
writer.write_header_bold(&["Name", "Amount", "Status"])?;

#### Styled Cells
use excelstream::types::{CellValue, CellStyle};

writer.write_row_styled(&[
    (CellValue::String("Total".into()), CellStyle::HeaderBold),
    (CellValue::Float(1234.56), CellStyle::NumberCurrency),
    (CellValue::Int(95), CellStyle::NumberPercentage),
])?;

#### Available Styles
- **CellStyle::HeaderBold** - Bold text for headers
- **CellStyle::NumberCurrency** - $#,##0.00
- **CellStyle::NumberPercentage** - 0.00%
- **CellStyle::NumberInteger** - #,##0
- **CellStyle::NumberDecimal** - #,##0.00
- **CellStyle::DateDefault** - MM/DD/YYYY
- **CellStyle::HighlightYellow** - Yellow background
- **CellStyle::TextBold** - Bold emphasis
- **CellStyle::TextItalic** - Italic notes
```

### Example File
Create `examples/cell_formatting.rs` demonstrating all styles

## Future Enhancements (v0.3.1+)

### Dynamic Style Builder
For users who need custom styles beyond the 14 presets:

```rust
let custom_style = StyleBuilder::new()
    .bold()
    .italic()
    .font_size(14)
    .background_color(Color::RGB(255, 200, 100))
    .border_all(BorderStyle::Medium)
    .build();

writer.write_row_with_custom_style(&cells, custom_style)?;
```

**Implementation approach:**
- StyleManager tracks unique style combinations (HashMap)
- Dynamically build styles.xml based on used styles
- More complex but fully flexible

### Additional Features
- Column width support
- Row height support
- Cell alignment (left, center, right)
- Text wrapping
- Merged cells
- Conditional formatting
- Data validation

## Success Criteria

✅ **Functional:**
- All 14 predefined styles work correctly in Excel
- write_header_bold() creates bold headers
- Styled cells display correct formatting

✅ **Performance:**
- Memory stays ~80MB constant (< 100MB for 1M rows)
- Speed within 5% of current performance
- No regressions in existing functionality

✅ **Compatibility:**
- Files open correctly in Excel 2016+
- Files open correctly in LibreOffice Calc
- Files open correctly in Google Sheets

✅ **Testing:**
- 15+ new tests covering all styles
- Integration test with real Excel file verification
- Performance benchmark showing < 5% overhead

## Timeline Estimate
- **Step 1-2**: 2-3 hours (types + worksheet)
- **Step 3**: 1-2 hours (styles.xml generation)
- **Step 4-5**: 2-3 hours (API + workbook integration)
- **Testing**: 2-3 hours
- **Documentation**: 1-2 hours
- **Total**: 8-13 hours

## Questions for User

1. **Scope**: Are the 14 predefined styles sufficient for v0.3.0? Or need dynamic styles immediately?
2. **Priority**: Which styles are most important? (e.g., HeaderBold, NumberCurrency, Percentage)
3. **API preference**: Do you prefer `write_row_styled()` or separate methods like `write_header_bold()`?
4. **Colors**: Are the 3 highlight colors (Yellow, Green, Red) sufficient?
5. **Column width**: Should this be included in Phase 3 or deferred to Phase 4?

## Conclusion

This plan provides a **pragmatic, incremental approach** to adding styling support:
- ✅ Ships quickly with 14 common styles
- ✅ Maintains streaming performance
- ✅ Easy to extend later with dynamic styles
- ✅ Backward compatible - existing code continues to work

**Recommendation**: Proceed with this plan for v0.3.0, then add dynamic StyleBuilder in v0.3.1 based on user feedback.
