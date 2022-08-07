// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::cmp;
use std::collections::HashMap;
use std::io::Write;
use std::mem;

use crate::error::XlsxError;
use crate::format::Format;
use crate::shared_strings_table::SharedStringsTable;
use crate::utility;
use crate::xmlwriter::XMLWriter;

pub type RowNum = u32;
pub type ColNum = u16;

const ROW_MAX: RowNum = 1_048_576;
const COL_MAX: ColNum = 16_384;

pub struct Worksheet {
    pub writer: XMLWriter,
    pub name: String,
    pub selected: bool,
    table: HashMap<RowNum, HashMap<ColNum, CellType>>,
    col_names: HashMap<ColNum, String>,
    dimensions: WorksheetDimensions,
    uses_string_table: bool,
}

impl Worksheet {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Worksheet struct.
    pub(crate) fn new(name: String) -> Worksheet {
        let writer = XMLWriter::new();
        let table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();
        let col_names: HashMap<ColNum, String> = HashMap::new();

        // Initialize the min and max dimensions with their opposite value.
        let dimensions = WorksheetDimensions {
            row_min: ROW_MAX,
            col_min: COL_MAX,
            row_max: 0,
            col_max: 0,
        };

        Worksheet {
            writer,
            name,
            selected: false,
            table,
            col_names,
            dimensions,
            uses_string_table: false,
        }
    }

    // Set the worksheet name instead of the default Sheet1, Sheet2, etc.
    pub fn set_name(&mut self, name: &str) -> &mut Worksheet {
        self.name = name.to_string();
        self
    }

    // Write a formatted number to a cell.
    pub fn write_number<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
        format: &Format,
    ) -> Result<(), XlsxError>
    where
        T: Into<f64>,
    {
        // Store the cell data.
        self.store_number(row, col, number.into(), Some(format))?;

        Ok(())
    }

    // Write an unformatted number to a cell.
    pub fn write_number_only<T>(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: T,
    ) -> Result<(), XlsxError>
    where
        T: Into<f64>,
    {
        // Store the cell data.
        self.store_number(row, col, number.into(), None)?;

        Ok(())
    }

    // Write a formatted string to a cell.
    pub fn write_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: &Format,
    ) -> Result<(), XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string, Some(format))?;

        Ok(())
    }

    // Write a unformatted string to a cell.
    pub fn write_string_only(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
    ) -> Result<(), XlsxError> {
        // Store the cell data.
        self.store_string(row, col, string, None)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Property getters.
    // -----------------------------------------------------------------------

    pub(crate) fn uses_string_table(&self) -> bool {
        self.uses_string_table
    }

    // -----------------------------------------------------------------------
    // Crate level helper methods.
    // -----------------------------------------------------------------------

    // Store a number cell.
    fn store_number(
        &mut self,
        row: RowNum,
        col: ColNum,
        number: f64,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColRange);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => format.xf_index(),
            None => 0,
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::Number { number, xf_index };

        self.insert_cell(row, col, cell);

        Ok(())
    }

    // Writer a unformatted string to a cell.
    pub fn store_string(
        &mut self,
        row: RowNum,
        col: ColNum,
        string: &str,
        format: Option<&Format>,
    ) -> Result<(), XlsxError> {
        // Check row and col are in the allowed range.
        if !self.check_dimensions(row, col) {
            return Err(XlsxError::RowColRange);
        }

        // Get the index of the format object, if any.
        let xf_index = match format {
            Some(format) => format.xf_index(),
            None => 0,
        };

        // Create the appropriate cell type to hold the data.
        let cell = CellType::String {
            string: string.to_string(),
            xf_index,
        };

        self.insert_cell(row, col, cell);
        self.uses_string_table = true;

        Ok(())
    }

    // Insert a cell value into the worksheet table data structure.
    fn insert_cell(&mut self, row: RowNum, col: ColNum, cell: CellType) {
        match self.table.get_mut(&row) {
            Some(columns) => {
                // The row already exists. Insert/replace column value.
                columns.insert(col, cell);
            }
            None => {
                // The row doesn't exist, create a new row with columns and insert
                // the cell value.
                let mut columns: HashMap<ColNum, CellType> = HashMap::new();
                columns.insert(col, cell);
                self.table.insert(row, columns);
            }
        }
    }

    // Check that row and col are within the allowed Excel range and store max
    // and min values for use in other methods/elements.
    fn check_dimensions(&mut self, row: RowNum, col: ColNum) -> bool {
        // Check that the row an column number are withing Excel's ranges.
        if row >= ROW_MAX {
            return false;
        }
        if col >= COL_MAX {
            return false;
        }

        // Store any changes in worksheet dimensions.
        self.dimensions.row_min = cmp::min(self.dimensions.row_min, row);
        self.dimensions.col_min = cmp::min(self.dimensions.col_min, col);
        self.dimensions.row_max = cmp::max(self.dimensions.row_max, row);
        self.dimensions.col_max = cmp::max(self.dimensions.col_max, col);

        true
    }

    // Cached/faster version of utility.col_to_name() to use in the inner loop.
    fn col_to_name(&mut self, col_num: ColNum) -> String {
        if let Some(col_name) = self.col_names.get(&col_num) {
            col_name.clone()
        } else {
            let col_name = utility::col_to_name(col_num);
            self.col_names.insert(col_num, col_name.clone());
            col_name
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self, string_table: &mut SharedStringsTable) {
        self.writer.xml_declaration();

        // Write the worksheet element.
        self.write_worksheet();

        // Write the dimension element.

        self.write_dimension();

        // Write the sheetViews element.
        self.write_sheet_views();

        // Write the sheetFormatPr element.
        self.write_sheet_format_pr();

        // Write the sheetData element.
        self.write_sheet_data(string_table);

        // Write the pageMargins element.
        self.write_page_margins();

        // Close the worksheet tag.
        self.writer.xml_end_tag("worksheet");
    }

    // Write the <worksheet> element.
    fn write_worksheet(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let xmlns_r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("worksheet", &attributes);
    }

    // Write the <dimension> element.
    fn write_dimension(&mut self) {
        let mut attributes = vec![];
        let mut range = "A1".to_string();

        if !self.table.is_empty() {
            range = utility::cell_range(
                self.dimensions.row_min,
                self.dimensions.col_min,
                self.dimensions.row_max,
                self.dimensions.col_max,
            );
        }

        attributes.push(("ref", range));

        self.writer.xml_empty_tag_attr("dimension", &attributes);
    }

    // Write the <sheetViews> element.
    fn write_sheet_views(&mut self) {
        self.writer.xml_start_tag("sheetViews");

        // Write the sheetView element.
        self.write_sheet_view();

        self.writer.xml_end_tag("sheetViews");
    }

    // Write the <sheetView> element.
    fn write_sheet_view(&mut self) {
        let mut attributes = vec![];

        if self.selected {
            attributes.push(("tabSelected", "1".to_string()));
        }

        attributes.push(("workbookViewId", "0".to_string()));

        self.writer.xml_empty_tag_attr("sheetView", &attributes);
    }

    // Write the <sheetFormatPr> element.
    fn write_sheet_format_pr(&mut self) {
        let attributes = vec![("defaultRowHeight", "15".to_string())];

        self.writer.xml_empty_tag_attr("sheetFormatPr", &attributes);
    }

    // Write the <sheetData> element.
    fn write_sheet_data(&mut self, string_table: &mut SharedStringsTable) {
        if self.table.is_empty() {
            self.writer.xml_empty_tag("sheetData");
        } else {
            self.writer.xml_start_tag("sheetData");
            self.write_data_table(string_table);
            self.writer.xml_end_tag("sheetData");
        }
    }

    // Write the <pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = vec![
            ("left", "0.7".to_string()),
            ("right", "0.7".to_string()),
            ("top", "0.75".to_string()),
            ("bottom", "0.75".to_string()),
            ("header", "0.3".to_string()),
            ("footer", "0.3".to_string()),
        ];

        self.writer.xml_empty_tag_attr("pageMargins", &attributes);
    }

    // Write out all the row and cell data in the worksheet data table.
    fn write_data_table(&mut self, string_table: &mut SharedStringsTable) {
        let spans = self.calculate_spans();

        // Swap out the worksheet data table so we can iterate over it and still
        // call self.write_xml() methods.
        // TODO. check efficiency of this and/or alternatives.
        let mut temp_table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();
        mem::swap(&mut temp_table, &mut self.table);

        for row_num in self.dimensions.row_min..=self.dimensions.row_max {
            let span_index = row_num / 16;
            let span = spans.get(&span_index);

            match temp_table.get(&row_num) {
                Some(columns) => {
                    self.write_row(row_num, span);

                    for col_num in self.dimensions.col_min..=self.dimensions.col_max {
                        match columns.get(&col_num) {
                            Some(cell) => match cell {
                                CellType::Number { number, xf_index } => {
                                    self.write_number_cell(row_num, col_num, number, xf_index)
                                }
                                CellType::String { string, xf_index } => {
                                    let string_index = string_table.get_shared_string_index(string);
                                    self.write_string_cell(
                                        row_num,
                                        col_num,
                                        &string_index,
                                        xf_index,
                                    );
                                }
                            },
                            _ => continue,
                        }
                    }

                    self.writer.xml_end_tag("row");
                }
                _ => continue,
            }
        }
    }

    // Calculate the "spans" attribute of the <row> tag. This is an XLSX
    // optimization and isn't strictly required. However, it makes comparing
    // files easier. The span is the same for each block of 16 rows.
    fn calculate_spans(&mut self) -> HashMap<u32, String> {
        let mut spans: HashMap<RowNum, String> = HashMap::new();
        let mut span_min = COL_MAX;
        let mut span_max = 0;

        for row_num in self.dimensions.row_min..=self.dimensions.row_max {
            if let Some(columns) = self.table.get(&row_num) {
                for col_num in self.dimensions.col_min..=self.dimensions.col_max {
                    match columns.get(&col_num) {
                        Some(_) => {
                            if span_min == COL_MAX {
                                span_min = col_num;
                                span_max = col_num;
                            } else {
                                span_min = cmp::min(span_min, col_num);
                                span_max = cmp::max(span_max, col_num);
                            }
                        }
                        _ => continue,
                    }
                }
            }

            // Store the span range for each block or 16 rows.
            if (row_num + 1) % 16 == 0 || row_num == self.dimensions.row_max {
                let span_index = row_num / 16;
                span_min += 1;
                span_max += 1;
                let span_range = format!("{}:{}", span_min, span_max);
                spans.insert(span_index, span_range);
                span_min = COL_MAX;
            }
        }

        spans
    }

    // Write the <row> element.
    fn write_row(&mut self, row_num: RowNum, span: Option<&String>) {
        let row_num = format!("{}", row_num + 1);
        let mut attributes = vec![("r", row_num)];

        if let Some(span_range) = span {
            attributes.push(("spans", span_range.clone()))
        }

        self.writer.xml_start_tag_attr("row", &attributes);
    }

    // Write the <c> element for a number.
    fn write_number_cell(&mut self, row: RowNum, col: ColNum, number: &f64, xf_index: &u32) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{}><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            number
        )
        .expect("Couldn't write to file");
    }

    // Write the <c> element for a string.
    fn write_string_cell(&mut self, row: RowNum, col: ColNum, string_index: &u32, xf_index: &u32) {
        let col_name = self.col_to_name(col);
        let mut style = String::from("");

        if *xf_index > 0 {
            style = format!(r#" s="{}""#, *xf_index);
        }

        write!(
            &mut self.writer.xmlfile,
            r#"<c r="{}{}"{} t="s"><v>{}</v></c>"#,
            col_name,
            row + 1,
            style,
            string_index
        )
        .expect("Couldn't write to file");
    }
}

// -----------------------------------------------------------------------
// Helper enums/structs
// -----------------------------------------------------------------------

struct WorksheetDimensions {
    row_min: RowNum,
    col_min: ColNum,
    row_max: RowNum,
    col_max: ColNum,
}
enum CellType {
    Number { number: f64, xf_index: u32 },
    String { string: String, xf_index: u32 },
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::SharedStringsTable;
    use super::Worksheet;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;
    use std::collections::HashMap;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet::new("".to_string());
        let mut string_table = SharedStringsTable::new();

        worksheet.selected = true;

        worksheet.assemble_xml_file(&mut string_table);

        let got = worksheet.writer.read_to_string();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <dimension ref="A1"/>
              <sheetViews>
                <sheetView tabSelected="1" workbookViewId="0"/>
              </sheetViews>
              <sheetFormatPr defaultRowHeight="15"/>
              <sheetData/>
              <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
            </worksheet>
            "#,
        );

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_1() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (0..17).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:16".to_string()), (1, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_2() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (1..18).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:15".to_string()), (1, "16:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_3() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (2..19).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:14".to_string()), (1, "15:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_4() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (3..20).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:13".to_string()), (1, "14:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_5() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (4..21).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:12".to_string()), (1, "13:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_6() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (5..22).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:11".to_string()), (1, "12:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_7() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (6..23).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:10".to_string()), (1, "11:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_8() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (7..24).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:9".to_string()), (1, "10:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_9() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (8..25).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:8".to_string()), (1, "9:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_10() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (9..26).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:7".to_string()), (1, "8:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_11() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (10..27).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:6".to_string()), (1, "7:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_12() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (11..28).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:5".to_string()), (1, "6:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_13() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (12..29).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:4".to_string()), (1, "5:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_14() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (13..30).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:3".to_string()), (1, "4:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_15() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (14..31).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:2".to_string()), (1, "3:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_16() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (15..32).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(0, "1:1".to_string()), (1, "2:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_17() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number_only(row_num, col_num as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "1:16".to_string()), (2, "17:17".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }

    #[test]
    fn test_calculate_spans_18() {
        let mut worksheet = Worksheet::new("".to_string());

        for (col_num, row_num) in (16..33).enumerate() {
            worksheet
                .write_number_only(row_num, (col_num + 1) as u16, 1.0)
                .unwrap();
        }

        let expected = HashMap::from([(1, "2:17".to_string()), (2, "18:18".to_string())]);
        let got = worksheet.calculate_spans();

        assert_eq!(got, expected);
    }
}
