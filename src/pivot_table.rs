// pivot_table - A module for creating the Excel pivotTable.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::fmt;
use std::io::Cursor;

use crate::utility;
use crate::worksheet::{COL_MAX, ROW_MAX};
use crate::xmlwriter::{
    escape_attribute_escapes, xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag,
    xml_start_tag,
};
use crate::{ColNum, Format, RowNum, XlsxError};

/// The `PivotTable` struct represents a worksheet pivot table.
///
/// This feature is a work in progress.
///
/// # Limitations
///
/// Pivot tables in Excel use a cache of the source data. This data is filtered
/// and summarized to create the pivot table. The `rust_xlsxwriter` library
/// doesn't store the cached data, so the pivot table is written without its
/// data. When Excel opens the file, it will build the cache and the pivot table
/// from the source data. This leads to the following limitations:
///
/// - Applications that don't refresh pivot tables when the file is opened will
///   show an empty pivot table. This also applies to pivot tables on protected
///   worksheets.
/// - Excel will ask the end user to save the file even if no visible changes
///   were made.
///
/// The overall limitation can't easily be fixed without replicating most of the
/// logic of the Excel Pivot Table engine.
///
#[derive(Clone)]
pub struct PivotTable {
    pub(crate) writer: Cursor<Vec<u8>>,

    pub(crate) name: String,
    pub(crate) cache_id: u32,
    pub(crate) source: PivotTableSource,
    pub(crate) field_names: Vec<String>,

    pub(crate) row: RowNum,
    pub(crate) col: ColNum,

    row_fields: Vec<String>,
    column_fields: Vec<String>,
    filter_fields: Vec<String>,
    pub(crate) value_fields: Vec<PivotTableValueField>,

    style: PivotTableStyle,
    layout: PivotTableLayout,
    show_row_grand_totals: bool,
    show_column_grand_totals: bool,
}

impl PivotTable {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    /// Create a new `PivotTable` struct instance.
    ///
    #[allow(clippy::new_without_default)]
    pub fn new() -> PivotTable {
        let writer = Cursor::new(Vec::with_capacity(2048));

        PivotTable {
            writer,
            name: String::new(),
            cache_id: 0,
            source: PivotTableSource::default(),
            field_names: vec![],
            row: 0,
            col: 0,
            row_fields: vec![],
            column_fields: vec![],
            filter_fields: vec![],
            value_fields: vec![],
            style: PivotTableStyle::Light16,
            layout: PivotTableLayout::Compact,
            show_row_grand_totals: true,
            show_column_grand_totals: true,
        }
    }

    /// Set the worksheet range of the source data for the pivot table.
    ///
    /// The first row of the range should contain the field names.
    ///
    /// # Parameters
    ///
    /// - `sheet_name`: The worksheet name of the source data.
    /// - `first_row`: The first row of the range. (All zero indexed.)
    /// - `first_col`: The first column of the range.
    /// - `last_row`: The last row of the range.
    /// - `last_col`: The last column of the range.
    ///
    pub fn set_data_source(
        mut self,
        sheet_name: &str,
        first_row: RowNum,
        first_col: ColNum,
        last_row: RowNum,
        last_col: ColNum,
    ) -> PivotTable {
        self.source = PivotTableSource {
            sheet_name: sheet_name.to_string(),
            first_row,
            first_col,
            last_row,
            last_col,
            field_names: std::mem::take(&mut self.source.field_names),
        };

        self
    }

    /// Set the field names of the source data.
    ///
    /// The field names are normally read from the header row of the source
    /// data when the workbook is saved. This isn't possible for a
    /// `constant_memory` or `low_memory` worksheet since the header row will
    /// have been written to disk, so the field names must be set explicitly.
    /// On other worksheets they are checked against the header row when the
    /// workbook is saved.
    ///
    /// # Parameters
    ///
    /// - `field_names`: The header names of the source data columns.
    ///
    pub fn set_field_names(mut self, field_names: &[impl AsRef<str>]) -> PivotTable {
        self.source.field_names = field_names
            .iter()
            .map(|name| name.as_ref().to_string())
            .collect();

        self
    }

    /// Add a row field to the pivot table.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This should match a header in the
    ///   first row of the source data.
    ///
    pub fn add_row_field(mut self, name: &str) -> PivotTable {
        self.row_fields.push(name.to_string());
        self
    }

    /// Add a value field to the pivot table.
    ///
    /// # Parameters
    ///
    /// - `value_field`: A [`PivotTableValueField`] struct reference.
    ///
    pub fn add_value_field(mut self, value_field: &PivotTableValueField) -> PivotTable {
        self.value_fields.push(value_field.clone());
        self
    }

    /// Add a column field to the pivot table.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This should match a header in the
    ///   first row of the source data.
    ///
    pub fn add_column_field(mut self, name: &str) -> PivotTable {
        self.column_fields.push(name.to_string());
        self
    }

    /// Add a filter field to the pivot table.
    ///
    /// The filter fields are displayed above the pivot table, with a blank row
    /// between them and the pivot table body.
    ///
    /// # Parameters
    ///
    /// - `name`: The name of the field. This should match a header in the
    ///   first row of the source data.
    ///
    pub fn add_filter_field(mut self, name: &str) -> PivotTable {
        self.filter_fields.push(name.to_string());
        self
    }

    /// Set the style of the pivot table.
    ///
    /// The default style is [`PivotTableStyle::Light16`].
    ///
    /// # Parameters
    ///
    /// - `style`: A [`PivotTableStyle`] enum value.
    ///
    pub fn set_style(mut self, style: PivotTableStyle) -> PivotTable {
        self.style = style;
        self
    }

    /// Set the report layout of the pivot table.
    ///
    /// The default layout is [`PivotTableLayout::Compact`].
    ///
    /// # Parameters
    ///
    /// - `layout`: A [`PivotTableLayout`] enum value.
    ///
    pub fn set_layout(mut self, layout: PivotTableLayout) -> PivotTable {
        self.layout = layout;
        self
    }

    /// Turn on/off the grand totals for the pivot table rows.
    ///
    /// The row grand totals are displayed in a "Grand Total" column to the
    /// right of the pivot table. They are on by default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn set_row_grand_totals(mut self, enable: bool) -> PivotTable {
        self.show_row_grand_totals = enable;
        self
    }

    /// Turn on/off the grand totals for the pivot table columns.
    ///
    /// The column grand totals are displayed in a "Grand Total" row at the
    /// bottom of the pivot table. They are on by default.
    ///
    /// # Parameters
    ///
    /// - `enable`: Turn the property on/off. It is on by default.
    ///
    pub fn set_column_grand_totals(mut self, enable: bool) -> PivotTable {
        self.show_column_grand_totals = enable;
        self
    }

    // Validate the pivot table data source and fields:
    //
    // - The source must be set and have at least one data row.
    // - Any field names must match the number of source columns.
    // - A field can only be used once in the row, column and filter fields,
    //   but it can also be a value field.
    // - Custom value field names must be unique, ignoring case.
    //
    pub(crate) fn validate(&self) -> Result<(), XlsxError> {
        let source = &self.source;

        if source.sheet_name.is_empty() {
            return Err(XlsxError::PivotTableError(
                "The pivot table data source must be set with set_data_source()".to_string(),
            ));
        }

        if source.last_row >= ROW_MAX || source.last_col >= COL_MAX {
            return Err(XlsxError::RowColumnLimitError);
        }

        if source.first_row > source.last_row || source.first_col > source.last_col {
            return Err(XlsxError::RowColumnOrderError);
        }

        // The source data needs a header row and at least one row of data.
        if source.first_row == source.last_row {
            return Err(XlsxError::PivotTableError(
                "The pivot table data source must have a header row and at least one data row"
                    .to_string(),
            ));
        }

        // Any user defined field names must match the number of columns.
        if !source.field_names.is_empty()
            && source.field_names.len() != source.num_fields() as usize
        {
            return Err(XlsxError::PivotTableError(format!(
                "The number of field names ({}) doesn't match the number of data source columns ({})",
                source.field_names.len(),
                source.num_fields()
            )));
        }

        let mut seen_names = vec![];

        for name in self
            .row_fields
            .iter()
            .chain(&self.column_fields)
            .chain(&self.filter_fields)
        {
            if seen_names.contains(&name) {
                return Err(XlsxError::PivotTableError(format!(
                    "Field '{name}' is used more than once in the row, column or filter fields"
                )));
            }

            seen_names.push(name);
        }

        let value_names = self.value_field_names();
        for (value_field, name) in self.value_fields.iter().zip(&value_names) {
            if !value_field.custom_name.is_empty()
                && value_names
                    .iter()
                    .filter(|other| other.to_lowercase() == name.to_lowercase())
                    .count()
                    > 1
            {
                return Err(XlsxError::PivotTableError(format!(
                    "Value field name '{name}' is used more than once"
                )));
            }
        }

        Ok(())
    }

    // Validate that the pivot table field names match the field names in the
    // header row of the source data. This can only be done when the workbook
    // is saved since the header data may not have been written before then.
    // A custom value field name also can't be the same as a header name,
    // ignoring case.
    pub(crate) fn validate_field_names(&self) -> Result<(), XlsxError> {
        for name in self
            .row_fields
            .iter()
            .chain(&self.column_fields)
            .chain(&self.filter_fields)
            .chain(
                self.value_fields
                    .iter()
                    .map(|value_field| &value_field.field),
            )
        {
            if self.field_index(name).is_none() {
                return Err(XlsxError::PivotTableError(format!(
                    "Field '{name}' doesn't match a header in the source data"
                )));
            }
        }

        for value_field in &self.value_fields {
            let name = &value_field.custom_name;
            if !name.is_empty()
                && self
                    .field_names
                    .iter()
                    .any(|field_name| field_name.to_lowercase() == name.to_lowercase())
            {
                return Err(XlsxError::PivotTableError(format!(
                    "Value field name '{name}' is the same as a header in the source data"
                )));
            }
        }

        Ok(())
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the pivotTableDefinition element.
        self.write_pivot_table_definition();

        // Write the location element.
        self.write_location();

        // Write the pivotFields element.
        self.write_pivot_fields();

        let has_fields = !self.row_fields.is_empty()
            || !self.column_fields.is_empty()
            || !self.value_fields.is_empty();

        if !self.row_fields.is_empty() {
            // Write the rowFields element.
            self.write_axis_fields("rowFields", &self.row_fields.clone());
        }

        if has_fields && (self.row_fields.is_empty() || self.show_column_grand_totals) {
            // Write the rowItems element.
            self.write_axis_items("rowItems", !self.row_fields.is_empty());
        }

        // Multiple value fields are displayed in columns via a virtual "Values"
        // field that is added to the column fields.
        let has_values_field = self.value_fields.len() > 1;

        if !self.column_fields.is_empty() || has_values_field {
            // Write the colFields element.
            self.write_col_fields(has_values_field);
        }

        if has_fields {
            if has_values_field {
                // With column fields the value items are all grand totals.
                if self.column_fields.is_empty() || self.show_row_grand_totals {
                    // Write the colItems element.
                    self.write_value_col_items();
                }
            } else if self.column_fields.is_empty() || self.show_row_grand_totals {
                // Write the colItems element.
                self.write_axis_items("colItems", !self.column_fields.is_empty());
            }
        }

        if !self.filter_fields.is_empty() {
            // Write the pageFields element.
            self.write_page_fields();
        }

        if !self.value_fields.is_empty() {
            // Write the dataFields element.
            self.write_data_fields();
        }

        // Write the pivotTableStyleInfo element.
        self.write_pivot_table_style_info();

        // Close the pivotTableDefinition tag.
        xml_end_tag(&mut self.writer, "pivotTableDefinition");
    }

    // Write the <pivotTableDefinition> element.
    fn write_pivot_table_definition(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        let mut attributes = vec![
            ("xmlns", xmlns.to_string()),
            ("name", escape_attribute_escapes(&self.name).into_owned()),
            ("cacheId", self.cache_id.to_string()),
            ("applyNumberFormats", "0".to_string()),
            ("applyBorderFormats", "0".to_string()),
            ("applyFontFormats", "0".to_string()),
            ("applyPatternFormats", "0".to_string()),
            ("applyAlignmentFormats", "0".to_string()),
            ("applyWidthHeightFormats", "1".to_string()),
            ("dataCaption", "Values".to_string()),
            ("updatedVersion", "3".to_string()),
            ("minRefreshableVersion", "3".to_string()),
            ("showCalcMbrs", "0".to_string()),
            ("useAutoFormatting", "1".to_string()),
        ];

        // The Excel UI and XML names are reversed. The UI row grand totals are
        // the "Grand Total" column, which is the XML colGrandTotals, and vice
        // versa.
        if !self.show_column_grand_totals {
            attributes.push(("rowGrandTotals", "0".to_string()));
        }

        if !self.show_row_grand_totals {
            attributes.push(("colGrandTotals", "0".to_string()));
        }

        attributes.extend([
            ("itemPrintTitles", "1".to_string()),
            ("createdVersion", "3".to_string()),
            ("indent", "0".to_string()),
        ]);

        match self.layout {
            PivotTableLayout::Compact => {
                attributes.push(("outline", "1".to_string()));
                attributes.push(("outlineData", "1".to_string()));
            }
            PivotTableLayout::Outline => {
                attributes.push(("compact", "0".to_string()));
                attributes.push(("outline", "1".to_string()));
                attributes.push(("outlineData", "1".to_string()));
                attributes.push(("compactData", "0".to_string()));
            }
            PivotTableLayout::Tabular => {
                attributes.push(("compact", "0".to_string()));
                attributes.push(("compactData", "0".to_string()));
            }
        }

        attributes.push(("multipleFieldFilters", "0".to_string()));

        xml_start_tag(&mut self.writer, "pivotTableDefinition", &attributes);
    }

    // Write the <location> element.
    fn write_location(&mut self) {
        let mut first_header_row = 1;
        // There is a header row for each column field, including the virtual
        // "Values" field used for multiple value fields.
        let mut num_col_fields = self.column_fields.len();
        if self.value_fields.len() > 1 {
            num_col_fields += 1;
        }
        let mut first_data_row = 1 + num_col_fields;
        // In the compact layout the row fields share a column. In the outline
        // and tabular layouts each row field has its own column.
        let first_data_col = if self.row_fields.is_empty() && self.column_fields.is_empty() {
            0
        } else if matches!(self.layout, PivotTableLayout::Compact) {
            1
        } else {
            self.row_fields.len().max(1)
        };

        // A pivot table with only filter fields has no header or data rows.
        if self.row_fields.is_empty()
            && self.column_fields.is_empty()
            && self.value_fields.is_empty()
            && !self.filter_fields.is_empty()
        {
            first_header_row = 0;
            first_data_row = 0;
        }

        let num_filters = self.filter_fields.len() as RowNum;

        let mut attributes = vec![
            ("ref", utility::row_col_to_cell(self.body_row(), self.col)),
            ("firstHeaderRow", first_header_row.to_string()),
            ("firstDataRow", first_data_row.to_string()),
            ("firstDataCol", first_data_col.to_string()),
        ];

        if num_filters > 0 {
            attributes.push(("rowPageCount", num_filters.to_string()));
            attributes.push(("colPageCount", "1".to_string()));
        }

        xml_empty_tag(&mut self.writer, "location", &attributes);
    }

    // Write the <pivotFields> element.
    fn write_pivot_fields(&mut self) {
        let count = self.source.num_fields();
        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "pivotFields", &attributes);

        for index in 0..count as usize {
            // Write the pivotField element.
            self.write_pivot_field(index);
        }

        xml_end_tag(&mut self.writer, "pivotFields");
    }

    // Write the <pivotField> element.
    fn write_pivot_field(&mut self, index: usize) {
        let is_row_field = self
            .row_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_column_field = self
            .column_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_filter_field = self
            .filter_fields
            .iter()
            .any(|name| self.field_index(name) == Some(index));

        let is_data_field = self
            .value_fields
            .iter()
            .any(|value_field| self.field_index(&value_field.field) == Some(index));

        let mut attributes = vec![];

        if is_row_field {
            attributes.push(("axis", "axisRow"));
        } else if is_column_field {
            attributes.push(("axis", "axisCol"));
        } else if is_filter_field {
            attributes.push(("axis", "axisPage"));
        }

        if is_data_field {
            attributes.push(("dataField", "1"));
        }

        match self.layout {
            PivotTableLayout::Compact => {}
            PivotTableLayout::Outline => {
                attributes.push(("compact", "0"));
            }
            PivotTableLayout::Tabular => {
                attributes.push(("compact", "0"));
                attributes.push(("outline", "0"));
            }
        }

        attributes.push(("showAll", "0"));

        if is_row_field || is_column_field || is_filter_field {
            xml_start_tag(&mut self.writer, "pivotField", &attributes);

            // Write the items element.
            self.write_items();

            xml_end_tag(&mut self.writer, "pivotField");
        } else {
            xml_empty_tag(&mut self.writer, "pivotField", &attributes);
        }
    }

    // Write the <items> element. The items are generated by Excel when the
    // pivot table is refreshed so we only write the default item.
    fn write_items(&mut self) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, "items", &attributes);

        // Write the item element.
        let attributes = [("t", "default")];
        xml_empty_tag(&mut self.writer, "item", &attributes);

        xml_end_tag(&mut self.writer, "items");
    }

    // Write the <rowFields> or <colFields> element.
    fn write_axis_fields(&mut self, tag: &str, names: &[String]) {
        let attributes = [("count", names.len().to_string())];

        xml_start_tag(&mut self.writer, tag, &attributes);

        for name in names {
            if let Some(index) = self.field_index(name) {
                // Write the field element.
                let attributes = [("x", index.to_string())];
                xml_empty_tag(&mut self.writer, "field", &attributes);
            }
        }

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <colFields> element, with the virtual "Values" field (index
    // -2) after the column fields if there are multiple value fields.
    fn write_col_fields(&mut self, has_values_field: bool) {
        let mut indices: Vec<String> = self
            .column_fields
            .iter()
            .filter_map(|name| self.field_index(name))
            .map(|index| index.to_string())
            .collect();

        if has_values_field {
            indices.push("-2".to_string());
        }

        let attributes = [("count", indices.len().to_string())];

        xml_start_tag(&mut self.writer, "colFields", &attributes);

        for index in indices {
            // Write the field element.
            let attributes = [("x", index)];
            xml_empty_tag(&mut self.writer, "field", &attributes);
        }

        xml_end_tag(&mut self.writer, "colFields");
    }

    // Write the <colItems> element for multiple value fields. If there are no
    // column fields Excel writes an item for each value field. Otherwise we
    // write a grand total for each value field.
    fn write_value_col_items(&mut self) {
        let count = self.value_fields.len();
        let has_column_fields = !self.column_fields.is_empty();
        let attributes = [("count", count.to_string())];

        xml_start_tag(&mut self.writer, "colItems", &attributes);

        for index in 0..count {
            let mut attributes = vec![];

            if has_column_fields {
                attributes.push(("t", "grand".to_string()));
            }

            if index > 0 {
                attributes.push(("i", index.to_string()));
            }

            // Write the i element.
            xml_start_tag(&mut self.writer, "i", &attributes);

            // Write the x element.
            if has_column_fields || index == 0 {
                xml_empty_tag_only(&mut self.writer, "x");
            } else {
                let attributes = [("v", index.to_string())];
                xml_empty_tag(&mut self.writer, "x", &attributes);
            }

            xml_end_tag(&mut self.writer, "i");
        }

        xml_end_tag(&mut self.writer, "colItems");
    }

    // Write the <rowItems> or <colItems> element. The items are generated by
    // Excel when the pivot table is refreshed so we only write the grand
    // total, if the axis has fields, or an empty item.
    fn write_axis_items(&mut self, tag: &str, has_fields: bool) {
        let attributes = [("count", "1")];

        xml_start_tag(&mut self.writer, tag, &attributes);

        // Write the i element.
        if has_fields {
            let attributes = [("t", "grand")];
            xml_start_tag(&mut self.writer, "i", &attributes);
            xml_empty_tag_only(&mut self.writer, "x");
            xml_end_tag(&mut self.writer, "i");
        } else {
            xml_empty_tag_only(&mut self.writer, "i");
        }

        xml_end_tag(&mut self.writer, tag);
    }

    // Write the <pageFields> element.
    fn write_page_fields(&mut self) {
        let attributes = [("count", self.filter_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "pageFields", &attributes);

        for name in self.filter_fields.clone() {
            if let Some(index) = self.field_index(&name) {
                // Write the pageField element.
                let attributes = [("fld", index.to_string()), ("hier", "-1".to_string())];
                xml_empty_tag(&mut self.writer, "pageField", &attributes);
            }
        }

        xml_end_tag(&mut self.writer, "pageFields");
    }

    // Write the <dataFields> element.
    fn write_data_fields(&mut self) {
        let attributes = [("count", self.value_fields.len().to_string())];

        xml_start_tag(&mut self.writer, "dataFields", &attributes);

        let names = self.value_field_names();
        for (value_field, name) in self.value_fields.clone().iter().zip(names) {
            // Write the dataField element.
            self.write_data_field(value_field, &name);
        }

        xml_end_tag(&mut self.writer, "dataFields");
    }

    // Write the <dataField> element.
    fn write_data_field(&mut self, value_field: &PivotTableValueField, name: &str) {
        let Some(index) = self.field_index(&value_field.field) else {
            return;
        };

        let mut attributes = vec![
            ("name", escape_attribute_escapes(name).into_owned()),
            ("fld", index.to_string()),
        ];

        if !matches!(value_field.function, PivotTableFunction::Sum) {
            attributes.push(("subtotal", value_field.function.to_string()));
        }

        attributes.push(("baseField", "0".to_string()));
        attributes.push(("baseItem", "0".to_string()));

        if value_field.num_format_index > 0 {
            attributes.push(("numFmtId", value_field.num_format_index.to_string()));
        }

        xml_empty_tag(&mut self.writer, "dataField", &attributes);
    }

    // Get the first row of the pivot table body. The (row, col) position is the
    // top left of the pivot table including any filter fields. The body is
    // below the filter fields and a blank row.
    pub(crate) fn body_row(&self) -> RowNum {
        let num_filters = self.filter_fields.len() as RowNum;

        if num_filters == 0 {
            self.row
        } else {
            self.row + num_filters + 1
        }
    }

    // Check if a cell is in the known part of the pivot table: the anchor
    // column from the anchor down to the first body row.
    pub(crate) fn overlaps_anchor(&self, row: RowNum, col: ColNum) -> bool {
        col == self.col && (self.row..=self.body_row()).contains(&row)
    }

    // Get the display names of the value fields. These are either the custom
    // name or a default name like "Sum of Volume". Excel makes duplicate
    // default names unique by appending a number, such as "Sum of Volume2".
    // The names are compared ignoring case.
    fn value_field_names(&self) -> Vec<String> {
        let mut names: Vec<String> = vec![];
        let contains = |names: &[String], name: &str| {
            names
                .iter()
                .any(|other| other.to_lowercase() == name.to_lowercase())
        };

        for value_field in &self.value_fields {
            let mut name = if value_field.custom_name.is_empty() {
                format!(
                    "{} of {}",
                    value_field.function.caption(),
                    value_field.field
                )
            } else {
                value_field.custom_name.clone()
            };

            if value_field.custom_name.is_empty() && contains(&names, &name) {
                let mut suffix = 2;
                while contains(&names, &format!("{name}{suffix}")) {
                    suffix += 1;
                }
                name = format!("{name}{suffix}");
            }

            names.push(name);
        }

        names
    }

    // Get the index of a field in the source data from its header name.
    fn field_index(&self, name: &str) -> Option<usize> {
        self.field_names
            .iter()
            .position(|field_name| field_name == name)
    }

    // Write the <pivotTableStyleInfo> element.
    fn write_pivot_table_style_info(&mut self) {
        let mut attributes = vec![];

        if !matches!(self.style, PivotTableStyle::None) {
            attributes.push(("name", self.style.to_string()));
        }

        attributes.push(("showRowHeaders", "1".to_string()));
        attributes.push(("showColHeaders", "1".to_string()));
        attributes.push(("showRowStripes", "0".to_string()));
        attributes.push(("showColStripes", "0".to_string()));
        attributes.push(("showLastColumn", "1".to_string()));

        xml_empty_tag(&mut self.writer, "pivotTableStyleInfo", &attributes);
    }
}

// -----------------------------------------------------------------------
// PivotTableValueField
// -----------------------------------------------------------------------

/// The `PivotTableValueField` struct represents a value field in a pivot
/// table.
///
/// This feature is a work in progress.
///
#[derive(Clone)]
pub struct PivotTableValueField {
    field: String,
    function: PivotTableFunction,
    custom_name: String,
    pub(crate) format: Option<Format>,
    pub(crate) num_format_index: u16,
}

impl PivotTableValueField {
    /// Create a new `PivotTableValueField` struct instance.
    ///
    /// # Parameters
    ///
    /// - `field`: The name of the source data field to summarize. This
    ///   should match a header in the first row of the source data.
    ///
    pub fn new(field: &str) -> PivotTableValueField {
        PivotTableValueField {
            field: field.to_string(),
            function: PivotTableFunction::Sum,
            custom_name: String::new(),
            format: None,
            num_format_index: 0,
        }
    }

    /// Set the function used to summarize the value field.
    ///
    /// The default function is [`PivotTableFunction::Sum`].
    ///
    /// # Parameters
    ///
    /// - `function`: A [`PivotTableFunction`] enum value.
    ///
    pub fn set_function(mut self, function: PivotTableFunction) -> PivotTableValueField {
        self.function = function;
        self
    }

    /// Set a custom name for the value field.
    ///
    /// The default name is the function and field name, such as "Sum of
    /// Volume".
    ///
    /// # Parameters
    ///
    /// - `name`: The custom name of the value field.
    ///
    pub fn set_custom_name(mut self, name: impl Into<String>) -> PivotTableValueField {
        self.custom_name = name.into();
        self
    }

    /// Set the number format for the value field.
    ///
    /// # Parameters
    ///
    /// - `num_format`: The number format string, such as `"#,##0.00"`. See
    ///   [`Format::set_num_format()`](crate::Format::set_num_format).
    ///
    pub fn set_num_format(mut self, num_format: impl Into<String>) -> PivotTableValueField {
        // The number format is stored internally as a cell Format so that it
        // shares the cell number format indices.
        self.format = Some(Format::new().set_num_format(num_format));
        self
    }
}

// -----------------------------------------------------------------------
// PivotTableLayout
// -----------------------------------------------------------------------

/// The `PivotTableLayout` enum defines the report layout of a pivot table.
///
/// These are the layouts in the Excel "Report Layout" menu.
///
#[derive(Clone)]
pub enum PivotTableLayout {
    /// Show the pivot table in compact form, with the row fields in one
    /// column. This is the default.
    Compact,

    /// Show the pivot table in outline form, with each row field in its own
    /// column and the items of the next field below the item in the current
    /// field.
    Outline,

    /// Show the pivot table in tabular form, with each row field in its own
    /// column.
    Tabular,
}

// -----------------------------------------------------------------------
// PivotTableStyle
// -----------------------------------------------------------------------

/// The `PivotTableStyle` enum defines the built-in Excel pivot table styles.
///
/// These are the styles in the Excel "Pivot Table Styles" gallery.
///
#[derive(Clone)]
pub enum PivotTableStyle {
    /// No pivot table style.
    None,

    /// Pivot Style Light 1.
    Light1,

    /// Pivot Style Light 2.
    Light2,

    /// Pivot Style Light 3.
    Light3,

    /// Pivot Style Light 4.
    Light4,

    /// Pivot Style Light 5.
    Light5,

    /// Pivot Style Light 6.
    Light6,

    /// Pivot Style Light 7.
    Light7,

    /// Pivot Style Light 8.
    Light8,

    /// Pivot Style Light 9.
    Light9,

    /// Pivot Style Light 10.
    Light10,

    /// Pivot Style Light 11.
    Light11,

    /// Pivot Style Light 12.
    Light12,

    /// Pivot Style Light 13.
    Light13,

    /// Pivot Style Light 14.
    Light14,

    /// Pivot Style Light 15.
    Light15,

    /// Pivot Style Light 16.
    Light16,

    /// Pivot Style Light 17.
    Light17,

    /// Pivot Style Light 18.
    Light18,

    /// Pivot Style Light 19.
    Light19,

    /// Pivot Style Light 20.
    Light20,

    /// Pivot Style Light 21.
    Light21,

    /// Pivot Style Light 22.
    Light22,

    /// Pivot Style Light 23.
    Light23,

    /// Pivot Style Light 24.
    Light24,

    /// Pivot Style Light 25.
    Light25,

    /// Pivot Style Light 26.
    Light26,

    /// Pivot Style Light 27.
    Light27,

    /// Pivot Style Light 28.
    Light28,

    /// Pivot Style Medium 1.
    Medium1,

    /// Pivot Style Medium 2.
    Medium2,

    /// Pivot Style Medium 3.
    Medium3,

    /// Pivot Style Medium 4.
    Medium4,

    /// Pivot Style Medium 5.
    Medium5,

    /// Pivot Style Medium 6.
    Medium6,

    /// Pivot Style Medium 7.
    Medium7,

    /// Pivot Style Medium 8.
    Medium8,

    /// Pivot Style Medium 9.
    Medium9,

    /// Pivot Style Medium 10.
    Medium10,

    /// Pivot Style Medium 11.
    Medium11,

    /// Pivot Style Medium 12.
    Medium12,

    /// Pivot Style Medium 13.
    Medium13,

    /// Pivot Style Medium 14.
    Medium14,

    /// Pivot Style Medium 15.
    Medium15,

    /// Pivot Style Medium 16.
    Medium16,

    /// Pivot Style Medium 17.
    Medium17,

    /// Pivot Style Medium 18.
    Medium18,

    /// Pivot Style Medium 19.
    Medium19,

    /// Pivot Style Medium 20.
    Medium20,

    /// Pivot Style Medium 21.
    Medium21,

    /// Pivot Style Medium 22.
    Medium22,

    /// Pivot Style Medium 23.
    Medium23,

    /// Pivot Style Medium 24.
    Medium24,

    /// Pivot Style Medium 25.
    Medium25,

    /// Pivot Style Medium 26.
    Medium26,

    /// Pivot Style Medium 27.
    Medium27,

    /// Pivot Style Medium 28.
    Medium28,

    /// Pivot Style Dark 1.
    Dark1,

    /// Pivot Style Dark 2.
    Dark2,

    /// Pivot Style Dark 3.
    Dark3,

    /// Pivot Style Dark 4.
    Dark4,

    /// Pivot Style Dark 5.
    Dark5,

    /// Pivot Style Dark 6.
    Dark6,

    /// Pivot Style Dark 7.
    Dark7,

    /// Pivot Style Dark 8.
    Dark8,

    /// Pivot Style Dark 9.
    Dark9,

    /// Pivot Style Dark 10.
    Dark10,

    /// Pivot Style Dark 11.
    Dark11,

    /// Pivot Style Dark 12.
    Dark12,

    /// Pivot Style Dark 13.
    Dark13,

    /// Pivot Style Dark 14.
    Dark14,

    /// Pivot Style Dark 15.
    Dark15,

    /// Pivot Style Dark 16.
    Dark16,

    /// Pivot Style Dark 17.
    Dark17,

    /// Pivot Style Dark 18.
    Dark18,

    /// Pivot Style Dark 19.
    Dark19,

    /// Pivot Style Dark 20.
    Dark20,

    /// Pivot Style Dark 21.
    Dark21,

    /// Pivot Style Dark 22.
    Dark22,

    /// Pivot Style Dark 23.
    Dark23,

    /// Pivot Style Dark 24.
    Dark24,

    /// Pivot Style Dark 25.
    Dark25,

    /// Pivot Style Dark 26.
    Dark26,

    /// Pivot Style Dark 27.
    Dark27,

    /// Pivot Style Dark 28.
    Dark28,
}

impl fmt::Display for PivotTableStyle {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::None => write!(f, ""),
            Self::Light1 => write!(f, "PivotStyleLight1"),
            Self::Light2 => write!(f, "PivotStyleLight2"),
            Self::Light3 => write!(f, "PivotStyleLight3"),
            Self::Light4 => write!(f, "PivotStyleLight4"),
            Self::Light5 => write!(f, "PivotStyleLight5"),
            Self::Light6 => write!(f, "PivotStyleLight6"),
            Self::Light7 => write!(f, "PivotStyleLight7"),
            Self::Light8 => write!(f, "PivotStyleLight8"),
            Self::Light9 => write!(f, "PivotStyleLight9"),
            Self::Light10 => write!(f, "PivotStyleLight10"),
            Self::Light11 => write!(f, "PivotStyleLight11"),
            Self::Light12 => write!(f, "PivotStyleLight12"),
            Self::Light13 => write!(f, "PivotStyleLight13"),
            Self::Light14 => write!(f, "PivotStyleLight14"),
            Self::Light15 => write!(f, "PivotStyleLight15"),
            Self::Light16 => write!(f, "PivotStyleLight16"),
            Self::Light17 => write!(f, "PivotStyleLight17"),
            Self::Light18 => write!(f, "PivotStyleLight18"),
            Self::Light19 => write!(f, "PivotStyleLight19"),
            Self::Light20 => write!(f, "PivotStyleLight20"),
            Self::Light21 => write!(f, "PivotStyleLight21"),
            Self::Light22 => write!(f, "PivotStyleLight22"),
            Self::Light23 => write!(f, "PivotStyleLight23"),
            Self::Light24 => write!(f, "PivotStyleLight24"),
            Self::Light25 => write!(f, "PivotStyleLight25"),
            Self::Light26 => write!(f, "PivotStyleLight26"),
            Self::Light27 => write!(f, "PivotStyleLight27"),
            Self::Light28 => write!(f, "PivotStyleLight28"),
            Self::Medium1 => write!(f, "PivotStyleMedium1"),
            Self::Medium2 => write!(f, "PivotStyleMedium2"),
            Self::Medium3 => write!(f, "PivotStyleMedium3"),
            Self::Medium4 => write!(f, "PivotStyleMedium4"),
            Self::Medium5 => write!(f, "PivotStyleMedium5"),
            Self::Medium6 => write!(f, "PivotStyleMedium6"),
            Self::Medium7 => write!(f, "PivotStyleMedium7"),
            Self::Medium8 => write!(f, "PivotStyleMedium8"),
            Self::Medium9 => write!(f, "PivotStyleMedium9"),
            Self::Medium10 => write!(f, "PivotStyleMedium10"),
            Self::Medium11 => write!(f, "PivotStyleMedium11"),
            Self::Medium12 => write!(f, "PivotStyleMedium12"),
            Self::Medium13 => write!(f, "PivotStyleMedium13"),
            Self::Medium14 => write!(f, "PivotStyleMedium14"),
            Self::Medium15 => write!(f, "PivotStyleMedium15"),
            Self::Medium16 => write!(f, "PivotStyleMedium16"),
            Self::Medium17 => write!(f, "PivotStyleMedium17"),
            Self::Medium18 => write!(f, "PivotStyleMedium18"),
            Self::Medium19 => write!(f, "PivotStyleMedium19"),
            Self::Medium20 => write!(f, "PivotStyleMedium20"),
            Self::Medium21 => write!(f, "PivotStyleMedium21"),
            Self::Medium22 => write!(f, "PivotStyleMedium22"),
            Self::Medium23 => write!(f, "PivotStyleMedium23"),
            Self::Medium24 => write!(f, "PivotStyleMedium24"),
            Self::Medium25 => write!(f, "PivotStyleMedium25"),
            Self::Medium26 => write!(f, "PivotStyleMedium26"),
            Self::Medium27 => write!(f, "PivotStyleMedium27"),
            Self::Medium28 => write!(f, "PivotStyleMedium28"),
            Self::Dark1 => write!(f, "PivotStyleDark1"),
            Self::Dark2 => write!(f, "PivotStyleDark2"),
            Self::Dark3 => write!(f, "PivotStyleDark3"),
            Self::Dark4 => write!(f, "PivotStyleDark4"),
            Self::Dark5 => write!(f, "PivotStyleDark5"),
            Self::Dark6 => write!(f, "PivotStyleDark6"),
            Self::Dark7 => write!(f, "PivotStyleDark7"),
            Self::Dark8 => write!(f, "PivotStyleDark8"),
            Self::Dark9 => write!(f, "PivotStyleDark9"),
            Self::Dark10 => write!(f, "PivotStyleDark10"),
            Self::Dark11 => write!(f, "PivotStyleDark11"),
            Self::Dark12 => write!(f, "PivotStyleDark12"),
            Self::Dark13 => write!(f, "PivotStyleDark13"),
            Self::Dark14 => write!(f, "PivotStyleDark14"),
            Self::Dark15 => write!(f, "PivotStyleDark15"),
            Self::Dark16 => write!(f, "PivotStyleDark16"),
            Self::Dark17 => write!(f, "PivotStyleDark17"),
            Self::Dark18 => write!(f, "PivotStyleDark18"),
            Self::Dark19 => write!(f, "PivotStyleDark19"),
            Self::Dark20 => write!(f, "PivotStyleDark20"),
            Self::Dark21 => write!(f, "PivotStyleDark21"),
            Self::Dark22 => write!(f, "PivotStyleDark22"),
            Self::Dark23 => write!(f, "PivotStyleDark23"),
            Self::Dark24 => write!(f, "PivotStyleDark24"),
            Self::Dark25 => write!(f, "PivotStyleDark25"),
            Self::Dark26 => write!(f, "PivotStyleDark26"),
            Self::Dark27 => write!(f, "PivotStyleDark27"),
            Self::Dark28 => write!(f, "PivotStyleDark28"),
        }
    }
}

// -----------------------------------------------------------------------
// PivotTableFunction
// -----------------------------------------------------------------------

/// The `PivotTableFunction` enum defines the functions used to summarize
/// pivot table value fields.
///
/// These are the functions in the Excel "Summarize value field by" list.
///
#[derive(Clone)]
pub enum PivotTableFunction {
    /// The sum of the values. This is the default.
    Sum,

    /// The number of values. This includes text values.
    Count,

    /// The average of the values.
    Average,

    /// The largest value.
    Max,

    /// The smallest value.
    Min,

    /// The product of the values.
    Product,

    /// The number of numeric values.
    CountNumbers,

    /// The standard deviation of a sample of the values.
    StdDev,

    /// The standard deviation of the population of the values.
    StdDevP,

    /// The variance of a sample of the values.
    Var,

    /// The variance of the population of the values.
    VarP,
}

impl PivotTableFunction {
    // The function name used in the default value field caption, such as
    // "Sum of Volume".
    fn caption(&self) -> &str {
        match self {
            Self::Sum => "Sum",
            Self::Max => "Max",
            Self::Min => "Min",
            Self::Var => "Var",
            Self::VarP => "Varp",
            Self::StdDev => "StdDev",
            Self::Average => "Average",
            Self::Product => "Product",
            Self::StdDevP => "StdDevp",
            Self::Count | Self::CountNumbers => "Count",
        }
    }
}

impl fmt::Display for PivotTableFunction {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Sum => write!(f, "sum"),
            Self::Max => write!(f, "max"),
            Self::Min => write!(f, "min"),
            Self::Var => write!(f, "var"),
            Self::VarP => write!(f, "varp"),
            Self::Count => write!(f, "count"),
            Self::StdDev => write!(f, "stdDev"),
            Self::Average => write!(f, "average"),
            Self::Product => write!(f, "product"),
            Self::StdDevP => write!(f, "stdDevp"),
            Self::CountNumbers => write!(f, "countNums"),
        }
    }
}

// -----------------------------------------------------------------------
// PivotTableSource
// -----------------------------------------------------------------------

// The worksheet range of the source data for a pivot table. Pivot tables with
// the same source share a pivot cache.
#[derive(Clone, Default, PartialEq)]
pub(crate) struct PivotTableSource {
    pub(crate) sheet_name: String,
    pub(crate) first_row: RowNum,
    pub(crate) first_col: ColNum,
    pub(crate) last_row: RowNum,
    pub(crate) last_col: ColNum,
    pub(crate) field_names: Vec<String>,
}

impl PivotTableSource {
    // The number of fields (columns) in the source data.
    pub(crate) fn num_fields(&self) -> ColNum {
        self.last_col - self.first_col + 1
    }

    // The number of data records (rows) in the source data, excluding the
    // header row.
    pub(crate) fn num_records(&self) -> RowNum {
        self.last_row - self.first_row
    }
}
