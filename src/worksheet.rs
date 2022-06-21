// worksheet - A module for creating the Excel Worksheet.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Worksheet<'a> {
    pub writer: &'a mut XMLWriter<'a>,
}

impl<'a> Worksheet<'a> {
    // Create a new Worksheet struct.
    pub fn new(writer: &'a mut XMLWriter<'a>) -> Worksheet<'a> {
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
    use super::XMLWriter;

    use pretty_assertions::assert_eq;
    use std::fs::File;
    use std::io::{Read, Seek, SeekFrom};
    use tempfile::tempfile;

    // Convert XML string/doc into a vector for comparison testing.
    pub fn xml_to_vec(xml_string: &str) -> Vec<String> {
        let mut xml_elements: Vec<String> = Vec::new();
        let re = regex::Regex::new(r">\s*<").unwrap();
        let tokens: Vec<&str> = re.split(xml_string).collect();

        for token in &tokens {
            let mut element = token.trim().to_string();

            // Add back the removed brackets.
            if !element.starts_with('<') {
                element = format!("<{}", element);
            }
            if !element.ends_with('>') {
                element = format!("{}>", element);
            }

            xml_elements.push(element);
        }
        xml_elements
    }

    // Test helper to read xml data back from a filehandle.
    fn read_xmlfile_data(tempfile: &mut File) -> String {
        let mut got = String::new();
        tempfile.seek(SeekFrom::Start(0)).unwrap();
        tempfile.read_to_string(&mut got).unwrap();
        got
    }

    #[test]
    fn test_assemble() {
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        let mut worksheet = Worksheet::new(&mut writer);

        worksheet.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
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
