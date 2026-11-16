// pivot_cache - A module for creating the Excel pivotCacheDefinition.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

use std::io::Cursor;

use crate::pivot_table::PivotTableSource;
use crate::utility;
use crate::xmlwriter::{
    escape_attribute_escapes, xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag,
    xml_start_tag,
};
use crate::{ColNum, XlsxError};

// A struct to represent the pivotCacheDefinition.xml file. The cache doesn't
// contain any records and is refreshed by Excel when the file is loaded.
pub(crate) struct PivotCache {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) source: PivotTableSource,
    pub(crate) field_names: Vec<String>,
}

impl PivotCache {
    // Create a new PivotCache struct instance.
    pub(crate) fn new(source: PivotTableSource) -> PivotCache {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotCache {
            writer,
            source,
            field_names: vec![],
        }
    }

    // Validate that the header cells in the source data are non-blank and
    // unique (case-insensitively) since Excel uses them as the pivot table
    // field names.
    pub(crate) fn validate_field_names(&self) -> Result<(), XlsxError> {
        let mut seen_names = vec![];

        for (col_offset, name) in self.field_names.iter().enumerate() {
            let cell = utility::row_col_to_cell(
                self.source.first_row,
                self.source.first_col + col_offset as ColNum,
            );

            if name.is_empty() {
                return Err(XlsxError::PivotTableError(format!(
                    "Header cell '{}!{cell}' in the source data is blank",
                    self.source.sheet_name
                )));
            }

            let lowercase_name = name.to_lowercase();
            if seen_names.contains(&lowercase_name) {
                return Err(XlsxError::PivotTableError(format!(
                    "Header '{name}' in cell '{}!{cell}' is duplicated in the source data",
                    self.source.sheet_name
                )));
            }

            seen_names.push(lowercase_name);
        }

        Ok(())
    }

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the pivotCacheDefinition element.
        self.write_pivot_cache_definition();

        // Write the cacheSource element.
        self.write_cache_source();

        // Write the cacheFields element.
        self.write_cache_fields();

        // Close the pivotCacheDefinition tag.
        xml_end_tag(&mut self.writer, "pivotCacheDefinition");
    }

    // Write the <pivotCacheDefinition> element.
    fn write_pivot_cache_definition(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let xmlns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        let attributes = [
            ("xmlns", xmlns.to_string()),
            ("xmlns:r", xmlns_rel.to_string()),
            ("saveData", "0".to_string()),
            ("refreshOnLoad", "1".to_string()),
            ("createdVersion", "3".to_string()),
            ("refreshedVersion", "3".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("recordCount", self.source.num_records().to_string()),
        ];

        xml_start_tag(&mut self.writer, "pivotCacheDefinition", &attributes);
    }

    // Write the <cacheSource> element.
    fn write_cache_source(&mut self) {
        let attributes = [("type", "worksheet")];

        xml_start_tag(&mut self.writer, "cacheSource", &attributes);

        // Write the worksheetSource element.
        self.write_worksheet_source();

        xml_end_tag(&mut self.writer, "cacheSource");
    }

    // Write the <worksheetSource> element.
    fn write_worksheet_source(&mut self) {
        let range = utility::cell_range(
            self.source.first_row,
            self.source.first_col,
            self.source.last_row,
            self.source.last_col,
        );

        let attributes = [
            ("ref", range),
            (
                "sheet",
                escape_attribute_escapes(&self.source.sheet_name).into_owned(),
            ),
        ];

        xml_empty_tag(&mut self.writer, "worksheetSource", &attributes);
    }

    // Write the <cacheFields> element.
    fn write_cache_fields(&mut self) {
        let attributes = [("count", self.field_names.len().to_string())];

        xml_start_tag(&mut self.writer, "cacheFields", &attributes);

        for name in self.field_names.clone() {
            // Write the cacheField element.
            self.write_cache_field(&name);
        }

        xml_end_tag(&mut self.writer, "cacheFields");
    }

    // Write the <cacheField> element.
    fn write_cache_field(&mut self, name: &str) {
        let attributes = [
            ("name", escape_attribute_escapes(name).into_owned()),
            ("numFmtId", "0".to_string()),
        ];

        xml_start_tag(&mut self.writer, "cacheField", &attributes);

        // Write the sharedItems element.
        xml_empty_tag_only(&mut self.writer, "sharedItems");

        xml_end_tag(&mut self.writer, "cacheField");
    }
}
