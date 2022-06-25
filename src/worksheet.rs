// worksheet - A module for creating the Excel sheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Worksheet {
    pub writer: XMLWriter,
}

impl Worksheet {
    // Create a new Worksheet struct.
    pub fn new() -> Worksheet {
        let writer = XMLWriter::new();

        Worksheet { writer }
    }

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
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
        let attributes = vec![("ref", "A1")];

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
        let attributes = vec![("tabSelected", "1"), ("workbookViewId", "0")];

        self.writer.xml_empty_tag_attr("sheetView", &attributes);
    }

    // Write the <sheetFormatPr> element.
    fn write_sheet_format_pr(&mut self) {
        let attributes = vec![("defaultRowHeight", "15")];

        self.writer.xml_empty_tag_attr("sheetFormatPr", &attributes);
    }

    // Write the <sheetData> element.
    fn write_sheet_data(&mut self) {
        self.writer.xml_empty_tag("sheetData");
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
}

#[cfg(test)]
mod tests {

    use super::Worksheet;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut worksheet = Worksheet::new();

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
