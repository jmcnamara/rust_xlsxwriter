// workbook - A module for creating the Excel Workbook.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct Workbook<'a> {
    pub writer: &'a mut XMLWriter<'a>,
}

impl<'a> Workbook<'a> {
    // Create a new Workbook struct.
    pub fn new(writer: &'a mut XMLWriter<'a>) -> Workbook<'a> {
        Workbook { writer }
    }

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the workbook element.
        self.write_workbook();

        // Write the fileVersion element.
        self.write_file_version();

        // Write the workbookPr element.
        self.write_workbook_pr();

        // Write the bookViews element.
        self.write_book_views();

        // Write the sheets element.
        self.write_sheets();

        // Write the calcPr element.
        self.write_calc_pr();

        // Close the workbook tag.
        self.writer.xml_end_tag("workbook");
    }

    // Write the <workbook> element.
    fn write_workbook(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let xmlns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("workbook", &attributes);
    }

    // Write the <fileVersion> element.
    fn write_file_version(&mut self) {
        let attributes = vec![
            ("appName", "xl"),
            ("lastEdited", "4"),
            ("lowestEdited", "4"),
            ("rupBuild", "4505"),
        ];

        self.writer.xml_empty_tag_attr("fileVersion", &attributes);
    }

    // Write the <workbookPr> element.
    fn write_workbook_pr(&mut self) {
        let attributes = vec![("defaultThemeVersion", "124226")];

        self.writer.xml_empty_tag_attr("workbookPr", &attributes);
    }

    // Write the <bookViews> element.
    fn write_book_views(&mut self) {
        self.writer.xml_start_tag("bookViews");

        // Write the workbookView element.
        self.write_workbook_view();

        self.writer.xml_end_tag("bookViews");
    }

    // Write the <workbookView> element.
    fn write_workbook_view(&mut self) {
        let attributes = vec![
            ("xWindow", "240"),
            ("yWindow", "15"),
            ("windowWidth", "16095"),
            ("windowHeight", "9660"),
        ];

        self.writer.xml_empty_tag_attr("workbookView", &attributes);
    }

    // Write the <sheets> element.
    fn write_sheets(&mut self) {
        self.writer.xml_start_tag("sheets");

        // Write the sheet element.
        self.write_sheet();

        self.writer.xml_end_tag("sheets");
    }

    // Write the <sheet> element.
    fn write_sheet(&mut self) {
        let attributes = vec![("name", "Sheet1"), ("sheetId", "1"), ("r:id", "rId1")];

        self.writer.xml_empty_tag_attr("sheet", &attributes);
    }

    // Write the <calcPr> element.
    fn write_calc_pr(&mut self) {
        let attributes = vec![("calcId", "124519"), ("fullCalcOnLoad", "1")];

        self.writer.xml_empty_tag_attr("calcPr", &attributes);
    }
}

#[cfg(test)]
mod tests {

    use super::Workbook;
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

        let mut workbook = Workbook::new(&mut writer);

        workbook.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
              <workbookPr defaultThemeVersion="124226"/>
              <bookViews>
                <workbookView xWindow="240" yWindow="15" windowWidth="16095" windowHeight="9660"/>
              </bookViews>
              <sheets>
                <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
              </sheets>
              <calcPr calcId="124519" fullCalcOnLoad="1"/>
            </workbook>
            "#,
        );

        assert_eq!(got, expected);
    }
}
