//! Zero-temp-file workbook - streams XML directly into ZIP compressor
//!
//! Expected memory: 8-12 MB (vs 17MB with temp files)

use super::shared_strings::SharedStrings;
use super::streaming_zip_writer::StreamingZipWriter;
use crate::error::Result;
use crate::types::ProtectionOptions;

/// Workbook that streams XML directly into compressor (no temp files)
pub struct ZeroTempWorkbook {
    zip_writer: Option<StreamingZipWriter>,
    worksheets: Vec<String>,
    worksheet_count: u32,
    current_row: u32,
    max_col: u32,
    xml_buffer: Vec<u8>,
    #[allow(dead_code)]
    shared_strings: SharedStrings,
    #[allow(dead_code)]
    protection: Option<ProtectionOptions>,
    in_worksheet: bool,
}

impl ZeroTempWorkbook {
    pub fn new(path: &str, compression_level: u32) -> Result<Self> {
        let zip_writer = StreamingZipWriter::new(path, compression_level)?;

        Ok(Self {
            zip_writer: Some(zip_writer),
            worksheets: Vec::new(),
            worksheet_count: 0,
            current_row: 0,
            max_col: 0,
            xml_buffer: Vec::with_capacity(4096),
            shared_strings: SharedStrings::new(),
            protection: None,
            in_worksheet: false,
        })
    }

    pub fn add_worksheet(&mut self, name: &str) -> Result<()> {
        // Finish previous worksheet if any
        self.finish_current_worksheet()?;

        self.worksheet_count += 1;
        self.worksheets.push(name.to_string());
        self.current_row = 0;
        self.max_col = 0;
        // Reset protection for new worksheet
        self.protection = None;

        // Start new worksheet entry in ZIP
        let entry_name = format!("xl/worksheets/sheet{}.xml", self.worksheet_count);
        self.zip_writer.as_mut().unwrap().start_entry(&entry_name)?;

        // Write worksheet XML header
        let header = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>"#;

        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(header.as_bytes())?;
        self.in_worksheet = true;

        Ok(())
    }

    pub fn protect_sheet(&mut self, options: ProtectionOptions) -> Result<()> {
        self.protection = Some(options);
        Ok(())
    }

    pub fn write_row(&mut self, values: &[&str]) -> Result<()> {
        if !self.in_worksheet {
            return Err(crate::error::ExcelError::WriteError(
                "No worksheet started".to_string(),
            ));
        }

        self.current_row += 1;
        self.max_col = self.max_col.max(values.len() as u32);

        // Build row XML in buffer
        self.xml_buffer.clear();
        self.xml_buffer.extend_from_slice(b"<row r=\"");
        self.xml_buffer
            .extend_from_slice(self.current_row.to_string().as_bytes());
        self.xml_buffer.extend_from_slice(b"\">");

        for (col_idx, value) in values.iter().enumerate() {
            let col_letter = Self::column_letter(col_idx as u32 + 1);
            self.xml_buffer.extend_from_slice(b"<c r=\"");
            self.xml_buffer.extend_from_slice(col_letter.as_bytes());
            self.xml_buffer
                .extend_from_slice(self.current_row.to_string().as_bytes());

            if value.is_empty() {
                self.xml_buffer.extend_from_slice(b"\"/>");
            } else {
                // Always treat as string to preserve leading zeros and exact formatting
                // Users should use write_row_typed() if they want numeric types
                self.xml_buffer
                    .extend_from_slice(b"\" t=\"inlineStr\"><is><t>");
                Self::write_escaped(&mut self.xml_buffer, value);
                self.xml_buffer.extend_from_slice(b"</t></is></c>");
            }
        }

        self.xml_buffer.extend_from_slice(b"</row>");

        // Stream to compressor immediately
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(&self.xml_buffer)?;

        Ok(())
    }

    /// Write a row with cell styling
    pub fn write_row_styled(&mut self, cells: &[crate::types::StyledCell]) -> Result<()> {
        if !self.in_worksheet {
            return Err(crate::error::ExcelError::WriteError(
                "No worksheet started".to_string(),
            ));
        }

        self.current_row += 1;
        self.max_col = self.max_col.max(cells.len() as u32);

        // Build row XML in buffer
        self.xml_buffer.clear();
        self.xml_buffer.extend_from_slice(b"<row r=\"");
        self.xml_buffer
            .extend_from_slice(self.current_row.to_string().as_bytes());
        self.xml_buffer.extend_from_slice(b"\">");

        for (col_idx, styled_cell) in cells.iter().enumerate() {
            let col_letter = Self::column_letter(col_idx as u32 + 1);
            let value = &styled_cell.value;
            let style_id = styled_cell.style.index();

            self.xml_buffer.extend_from_slice(b"<c r=\"");
            self.xml_buffer.extend_from_slice(col_letter.as_bytes());
            self.xml_buffer
                .extend_from_slice(self.current_row.to_string().as_bytes());
            self.xml_buffer.extend_from_slice(b"\"");

            // Add style attribute if not default
            if style_id > 0 {
                self.xml_buffer.extend_from_slice(b" s=\"");
                self.xml_buffer
                    .extend_from_slice(style_id.to_string().as_bytes());
                self.xml_buffer.extend_from_slice(b"\"");
            }

            // Write cell value based on type
            match value {
                crate::types::CellValue::Empty => {
                    self.xml_buffer.extend_from_slice(b"/>");
                }
                crate::types::CellValue::Int(i) => {
                    self.xml_buffer.extend_from_slice(b" t=\"n\"><v>");
                    self.xml_buffer.extend_from_slice(i.to_string().as_bytes());
                    self.xml_buffer.extend_from_slice(b"</v></c>");
                }
                crate::types::CellValue::Float(f) => {
                    self.xml_buffer.extend_from_slice(b" t=\"n\"><v>");
                    self.xml_buffer.extend_from_slice(f.to_string().as_bytes());
                    self.xml_buffer.extend_from_slice(b"</v></c>");
                }
                crate::types::CellValue::Bool(b) => {
                    self.xml_buffer.extend_from_slice(b" t=\"b\"><v>");
                    self.xml_buffer
                        .extend_from_slice(if *b { b"1" } else { b"0" });
                    self.xml_buffer.extend_from_slice(b"</v></c>");
                }
                crate::types::CellValue::String(s) => {
                    self.xml_buffer
                        .extend_from_slice(b" t=\"inlineStr\"><is><t>");
                    Self::write_escaped(&mut self.xml_buffer, s);
                    self.xml_buffer.extend_from_slice(b"</t></is></c>");
                }
                crate::types::CellValue::Formula(f) => {
                    self.xml_buffer.extend_from_slice(b"><f>");
                    Self::write_escaped(&mut self.xml_buffer, f);
                    self.xml_buffer.extend_from_slice(b"</f></c>");
                }
                crate::types::CellValue::DateTime(dt) => {
                    // Excel date serial number
                    self.xml_buffer.extend_from_slice(b" t=\"n\"><v>");
                    self.xml_buffer.extend_from_slice(dt.to_string().as_bytes());
                    self.xml_buffer.extend_from_slice(b"</v></c>");
                }
                crate::types::CellValue::Error(e) => {
                    self.xml_buffer.extend_from_slice(b" t=\"e\"><v>");
                    Self::write_escaped(&mut self.xml_buffer, e);
                    self.xml_buffer.extend_from_slice(b"</v></c>");
                }
            }
        }

        self.xml_buffer.extend_from_slice(b"</row>");

        // Stream to compressor immediately
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(&self.xml_buffer)?;

        Ok(())
    }

    fn finish_current_worksheet(&mut self) -> Result<()> {
        if self.in_worksheet {
            // Close sheetData
            self.zip_writer
                .as_mut()
                .unwrap()
                .write_data(b"</sheetData>")?;

            // Add sheetProtection if present
            if let Some(ref prot) = self.protection {
                let mut protection_xml = String::from("<sheetProtection sheet=\"1\"");

                // Add password hash if present
                if let Some(ref hash) = prot.password_hash {
                    protection_xml.push_str(&format!(" password=\"{}\"", hash));
                }

                // For Excel protection:
                // - If field = false (don't allow), we don't set attribute (default is protected)
                // - If field = true (allow), we set attribute = "0" (not protected)

                if prot.select_locked_cells {
                    protection_xml.push_str(" selectLockedCells=\"0\"");
                }
                if prot.select_unlocked_cells {
                    protection_xml.push_str(" selectUnlockedCells=\"0\"");
                }
                if prot.format_cells {
                    protection_xml.push_str(" formatCells=\"0\"");
                }
                if prot.format_columns {
                    protection_xml.push_str(" formatColumns=\"0\"");
                }
                if prot.format_rows {
                    protection_xml.push_str(" formatRows=\"0\"");
                }
                if prot.insert_columns {
                    protection_xml.push_str(" insertColumns=\"0\"");
                }
                if prot.insert_rows {
                    protection_xml.push_str(" insertRows=\"0\"");
                }
                if prot.delete_columns {
                    protection_xml.push_str(" deleteColumns=\"0\"");
                }
                if prot.delete_rows {
                    protection_xml.push_str(" deleteRows=\"0\"");
                }
                if prot.sort {
                    protection_xml.push_str(" sort=\"0\"");
                }
                if prot.auto_filter {
                    protection_xml.push_str(" autoFilter=\"0\"");
                }

                protection_xml.push_str("/>");

                self.zip_writer
                    .as_mut()
                    .unwrap()
                    .write_data(protection_xml.as_bytes())?;
            }

            // Close worksheet
            self.zip_writer
                .as_mut()
                .unwrap()
                .write_data(b"</worksheet>")?;
            self.in_worksheet = false;
        }
        Ok(())
    }

    pub fn close(mut self) -> Result<()> {
        // Finish current worksheet
        self.finish_current_worksheet()?;

        // Write all other required ZIP entries
        self.write_content_types()?;
        self.write_rels()?;
        self.write_workbook()?;
        self.write_workbook_rels()?;
        self.write_styles()?;
        self.write_shared_strings()?;
        self.write_app_props()?;
        self.write_core_props()?;

        // Finish ZIP
        self.zip_writer.take().unwrap().finish()?;

        Ok(())
    }

    fn write_content_types(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("[Content_Types].xml")?;
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>"#,
        );

        for i in 1..=self.worksheet_count {
            xml.push_str(&format!(
                r#"
<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                i
            ));
        }

        xml.push_str("\n</Types>");
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_rels(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("_rels/.rels")?;
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#;
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_workbook(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("xl/workbook.xml")?;
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>"#,
        );

        for (i, name) in self.worksheets.iter().enumerate() {
            xml.push_str(&format!(
                r#"
<sheet name="{}" sheetId="{}" r:id="rId{}"/>"#,
                name,
                i + 1,
                i + 1
            ));
        }

        xml.push_str("\n</sheets>\n</workbook>");
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_workbook_rels(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("xl/_rels/workbook.xml.rels")?;
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
        );

        for i in 1..=self.worksheet_count {
            xml.push_str(&format!(
                r#"
<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#,
                i, i
            ));
        }

        xml.push_str(&format!(
            r#"
<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#,
            self.worksheet_count + 1,
            self.worksheet_count + 2
        ));

        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_styles(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("xl/styles.xml")?;
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="0"/>
<fonts count="3">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><sz val="11"/><name val="Calibri"/></font>
<font><i/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="5">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
</borders>
<cellXfs count="14">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="5" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
</cellXfs>
</styleSheet>"#;
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_shared_strings(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("xl/sharedStrings.xml")?;
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>
"#;
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_app_props(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("docProps/app.xml")?;
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>ExcelStream</Application>
</Properties>"#;
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn write_core_props(&mut self) -> Result<()> {
        self.zip_writer
            .as_mut()
            .unwrap()
            .start_entry("docProps/core.xml")?;
        let xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>ExcelStream</dc:creator>
</cp:coreProperties>"#;
        self.zip_writer
            .as_mut()
            .unwrap()
            .write_data(xml.as_bytes())?;
        Ok(())
    }

    fn column_letter(n: u32) -> String {
        let mut result = String::new();
        let mut n = n;
        while n > 0 {
            let rem = (n - 1) % 26;
            result.insert(0, (b'A' + rem as u8) as char);
            n = (n - 1) / 26;
        }
        result
    }

    fn write_escaped(buffer: &mut Vec<u8>, s: &str) {
        for c in s.chars() {
            match c {
                '&' => buffer.extend_from_slice(b"&amp;"),
                '<' => buffer.extend_from_slice(b"&lt;"),
                '>' => buffer.extend_from_slice(b"&gt;"),
                '"' => buffer.extend_from_slice(b"&quot;"),
                '\'' => buffer.extend_from_slice(b"&apos;"),
                _ => {
                    let mut buf = [0; 4];
                    buffer.extend_from_slice(c.encode_utf8(&mut buf).as_bytes());
                }
            }
        }
    }
}
