// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::cmp;
use std::collections::HashMap;
use std::mem;

use crate::utility;
use crate::xmlwriter::XMLWriter;

pub type RowNum = u32;
pub type ColNum = u16;

const ROW_MAX: RowNum = 1048576;
const COL_MAX: ColNum = 16384;

pub struct Worksheet {
    pub writer: XMLWriter,
    pub name: String,
    pub selected: bool,
    table: HashMap<RowNum, HashMap<ColNum, CellType>>,
    dimensions: WorksheetDimensions,
}

impl<'a> Worksheet {
    //
    // Public (and crate public) methods.
    //

    // Create a new Worksheet struct.
    pub(crate) fn new(name: String) -> Worksheet {
        let writer = XMLWriter::new();
        let table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();

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
            dimensions,
        }
    }

    // Set the worksheet name instead of the default Sheet1, Sheet2, etc.
    pub fn set_name(&mut self, name: &str) -> &mut Worksheet {
        self.name = name.to_string();
        self
    }

    // Writer a number to a cell.
    pub fn write_number(&mut self, row: RowNum, col: ColNum, number: f64) {
        if !self.check_dimensions(row, col) {
            return;
        }

        let cell = CellType::Number { number };
        self.insert_cell(row, col, cell)
    }

    //
    // Crate level helper methods.
    //

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

    //
    // XML assembly methods.
    //

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
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
        self.write_sheet_data();

        // Write the pageMargins element.
        self.write_page_margins();

        // Close the worksheet tag.
        self.writer.xml_end_tag("worksheet");
    }

    // Write the <worksheet> element.
    fn write_worksheet(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("worksheet", &attributes);
    }

    // Write the <dimension> element.
    fn write_dimension(&mut self) {
        let mut attributes = vec![];
        let mut range = String::from("A1");

        if !self.table.is_empty() {
            range = utility::cell_range(
                self.dimensions.row_min,
                self.dimensions.col_min,
                self.dimensions.row_max,
                self.dimensions.col_max,
            );
        }

        attributes.push(("ref", range.as_str()));

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
            attributes.push(("tabSelected", "1"));
        }

        attributes.push(("workbookViewId", "0"));

        self.writer.xml_empty_tag_attr("sheetView", &attributes);
    }

    // Write the <sheetFormatPr> element.
    fn write_sheet_format_pr(&mut self) {
        let attributes = vec![("defaultRowHeight", "15")];

        self.writer.xml_empty_tag_attr("sheetFormatPr", &attributes);
    }

    // Write the <sheetData> element.
    fn write_sheet_data(&mut self) {
        if self.table.is_empty() {
            self.writer.xml_empty_tag("sheetData");
        } else {
            self.writer.xml_start_tag("sheetData");
            self.writer_data_table();
            self.writer.xml_end_tag("sheetData");
        }
    }

    // Write the <pageMargins> element.
    fn write_page_margins(&mut self) {
        let attributes = vec![
            ("left", "0.7"),
            ("right", "0.7"),
            ("top", "0.75"),
            ("bottom", "0.75"),
            ("header", "0.3"),
            ("footer", "0.3"),
        ];

        self.writer.xml_empty_tag_attr("pageMargins", &attributes);
    }

    // Write out all the row and cell data in the worksheet data table.
    fn writer_data_table(&mut self) {
        // Swap out the worksheet data table so we can iterate over it and still
        // call self.write_xml() methods.
        // TODO. check efficiency of this and/or alternatives.
        let mut temp_table: HashMap<RowNum, HashMap<ColNum, CellType>> = HashMap::new();
        mem::swap(&mut temp_table, &mut self.table);

        for row_num in self.dimensions.row_min..=self.dimensions.row_max {
            match temp_table.get(&row_num) {
                Some(columns) => {
                    self.write_row(row_num);

                    for col_num in self.dimensions.col_min..=self.dimensions.col_max {
                        match columns.get(&col_num) {
                            Some(cell) => match cell {
                                CellType::Number { number } => {
                                    self.write_number_cell(row_num, col_num, number)
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

    // Write the <row> element.
    fn write_row(&mut self, row_num: RowNum) {
        let row_num = format!("{}", row_num + 1);
        let attributes = vec![("r", row_num.as_str()), ("spans", "1:3")];

        self.writer.xml_start_tag_attr("row", &attributes);
    }

    // Write the <c> element for a number.
    fn write_number_cell(&mut self, row: RowNum, col: ColNum, number: &f64) {
        let range = utility::rowcol_to_cell(row, col);
        let attributes = vec![("r", range.as_str())];

        self.writer.xml_start_tag_attr("c", &attributes);

        self.write_number_value(format!("{}", number).as_str());
        self.writer.xml_end_tag("c");
    }

    // Write the <v> element.
    fn write_number_value(&mut self, value: &str) {
        self.writer.xml_data_element("v", value);
    }
}

//
// Helper enums/structs
//

struct WorksheetDimensions {
    row_min: RowNum,
    col_min: ColNum,
    row_max: RowNum,
    col_max: ColNum,
}
enum CellType {
    Number { number: f64 },
}

//
// Tests.
//
#[cfg(test)]
mod tests {

    use super::Worksheet;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet::new("".to_string());

        worksheet.selected = true;

        worksheet.assemble_xml_file();

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
}
