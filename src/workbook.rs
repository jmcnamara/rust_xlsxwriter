// workbook - A module for creating the Excel workbook.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::collections::HashMap;

use crate::format::Format;
use crate::packager::Packager;
use crate::packager::PackagerOptions;
use crate::shared_strings_table::SharedStringsTable;
use crate::worksheet::Worksheet;
use crate::xmlwriter::XMLWriter;

pub struct Workbook<'a> {
    pub writer: XMLWriter,
    filename: &'a str,
    worksheets: Vec<Worksheet>,
    xf_formats: Vec<Format>,
    xf_indices: HashMap<String, u32>,
    font_count: u32,
}

impl<'a> Workbook<'a> {
    // -----------------------------------------------------------------------
    // Public (and crate public) methods.
    // -----------------------------------------------------------------------

    // Create a new Workbook struct.
    pub fn new(filename: &'a str) -> Workbook {
        let writer = XMLWriter::new();
        let default_format = Format::new();
        let xf_indices = HashMap::from([(default_format.get_format_key(), 0)]);

        Workbook {
            writer,
            filename,
            worksheets: vec![],
            xf_formats: vec![default_format],
            xf_indices,
            font_count: 0,
        }
    }

    // Prototype function for adding worksheets.
    pub fn add_worksheet(&mut self) -> &mut Worksheet {
        let sheet_name = format!("Sheet{}", self.worksheets.len() + 1);

        let worksheet = Worksheet::new(sheet_name);
        self.worksheets.push(worksheet);
        let worksheet = self.worksheets.last_mut().unwrap();

        worksheet
    }

    pub fn add_format(&mut self) -> Format {
        Format::new()
    }

    // Set the index for the format.
    pub fn register_format(&mut self, format: &mut Format) {
        let format_key = format.get_format_key();

        match self.xf_indices.get_mut(&format_key) {
            Some(xf_index) => {
                format.set_xf_index(*xf_index);
            }
            None => {
                let xf_index = self.xf_formats.len() as u32;
                self.xf_formats.push(format.clone());
                format.set_xf_index(xf_index);

                self.xf_indices.insert(format_key, xf_index);
            }
        }
    }

    // Assemble the xlsx file and close it.
    pub fn close(&mut self) {
        // Ensure that there is at least one worksheet in the workbook.
        if self.worksheets.is_empty() {
            self.add_worksheet();
        }
        // Ensure one sheet is selected.
        self.worksheets[0].selected = true;

        // Prepare the formats for writing with styles.rs.
        self.prepare_format_properties();

        // Create the Packager object that will assemble the zip/xlsx file.
        let mut packager = Packager::new(self.filename);
        let mut package_options = PackagerOptions::new();

        // Set some of the packager options.
        package_options.num_worksheets = self.worksheets.len() as u16;
        for worksheet in self.worksheets.iter() {
            package_options.worksheet_names.push(worksheet.name.clone())
        }

        // Update and write the share string table.
        let string_table = self.update_shared_strings();
        if string_table.unique_count > 0 {
            packager.write_shared_strings_file(string_table);
            package_options.has_sst_table = true;
        }

        // Start the zip/xlsx container.
        packager.create_root_files(&package_options);

        // Write the styles.xml file to the zip/xlsx container.

        packager.create_styles_file(&self.xf_formats, self.font_count);

        // Write the workbook to the zip/xlsx container.
        packager.write_workbook_file(self);

        // Write the worksheets to the zip/xlsx container.
        for (index, worksheet) in self.worksheets.iter_mut().enumerate() {
            packager.write_worksheet_file(worksheet, index + 1);
        }

        // Write the docProp files to the zip/xlsx container.
        packager.create_doc_prop_files(&package_options);

        // Close and write the final zip/xlsx container.
        packager.close();
    }

    // -----------------------------------------------------------------------
    // Internal function/methods.
    // -----------------------------------------------------------------------

    // Iterate through the worksheets and assign a string index for each unique string.
    fn update_shared_strings(&mut self) -> SharedStringsTable {
        let mut string_table = SharedStringsTable::new();

        for worksheet in self.worksheets.iter_mut() {
            worksheet.update_shared_strings(&mut string_table);
        }

        string_table
    }

    // Prepare all Format properties prior to passing them to styles.rs.
    fn prepare_format_properties(&mut self) {
        // Set the font index for the format objects.
        self.prepare_fonts();
    }

    // Set the font index for the format objects.
    fn prepare_fonts(&mut self) {
        let mut font_count = 0;
        let mut font_indices: HashMap<String, u32> = HashMap::new();

        for xf_format in &mut self.xf_formats {
            let font_key = xf_format.get_font_key();

            match font_indices.get(&font_key) {
                Some(font_index) => {
                    xf_format.set_font_index(*font_index, false);
                }
                None => {
                    font_indices.insert(font_key, font_count);
                    xf_format.set_font_index(font_count, true);
                    font_count += 1;
                }
            }
        }
        self.font_count = font_count;
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
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
        let xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main".to_string();
        let xmlns_r =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:r", xmlns_r)];

        self.writer.xml_start_tag_attr("workbook", &attributes);
    }

    // Write the <fileVersion> element.
    fn write_file_version(&mut self) {
        let attributes = vec![
            ("appName", "xl".to_string()),
            ("lastEdited", "4".to_string()),
            ("lowestEdited", "4".to_string()),
            ("rupBuild", "4505".to_string()),
        ];

        self.writer.xml_empty_tag_attr("fileVersion", &attributes);
    }

    // Write the <workbookPr> element.
    fn write_workbook_pr(&mut self) {
        let attributes = vec![("defaultThemeVersion", "124226".to_string())];

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
            ("xWindow", "240".to_string()),
            ("yWindow", "15".to_string()),
            ("windowWidth", "16095".to_string()),
            ("windowHeight", "9660".to_string()),
        ];

        self.writer.xml_empty_tag_attr("workbookView", &attributes);
    }

    // Write the <sheets> element.
    fn write_sheets(&mut self) {
        self.writer.xml_start_tag("sheets");

        let mut worksheet_names = vec![];
        for worksheet in self.worksheets.iter() {
            worksheet_names.push(worksheet.name.clone());
        }

        for (index, name) in worksheet_names.iter().enumerate() {
            // Write the sheet element.
            self.write_sheet(name, (index + 1) as u16);
        }

        self.writer.xml_end_tag("sheets");
    }

    // Write the <sheet> element.
    fn write_sheet(&mut self, name: &str, index: u16) {
        //let name = name;
        let sheet_id = format!("{}", index);
        let ref_id = format!("rId{}", index);

        let attributes = vec![
            ("name", name.to_string()),
            ("sheetId", sheet_id),
            ("r:id", ref_id),
        ];

        self.writer.xml_empty_tag_attr("sheet", &attributes);
    }

    // Write the <calcPr> element.
    fn write_calc_pr(&mut self) {
        let attributes = vec![
            ("calcId", "124519".to_string()),
            ("fullCalcOnLoad", "1".to_string()),
        ];

        self.writer.xml_empty_tag_attr("calcPr", &attributes);
    }
}

#[cfg(test)]
mod tests {

    use super::Workbook;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut workbook = Workbook::new("test.xlsx");
        workbook.add_worksheet();

        workbook.assemble_xml_file();

        let got = workbook.writer.read_to_string();
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
