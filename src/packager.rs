// packager - A library for assembling xml files into an Excel XLSX file.
//
// This module is used in conjunction by rust_xlsxwriter to create an Excel XLSX
// container file.
//
// From Wikipedia: The Open Packaging Conventions (OPC) is a container-file
// technology initially created by Microsoft to store a combination of XML and
// non-XML files that together form a single entity such as an Open XML Paper
// Specification (OpenXPS) document.
// http://en.wikipedia.org/wiki/Open_Packaging_Conventions.
//
// At its simplest an Excel XLSX file contains the following elements::
//
//      ____ [Content_Types].xml
//     |
//     |____ docProps
//     | |____ app.xml
//     | |____ core.xml
//     |
//     |____ xl
//     | |____ workbook.xml
//     | |____ worksheets
//     | | |____ sheet1.xml
//     | |
//     | |____ styles.xml
//     | |
//     | |____ theme
//     | | |____ theme1.xml
//     | |
//     | |_____rels
//     | |____ workbook.xml.rels
//     |
//     |_____rels
//       |____ .rels
//
// The Packager class coordinates the classes that represent the elements of the
// package and writes them into the XLSX file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::relationship::Relationship;
use crate::shared_strings::SharedStrings;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::theme::Theme;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;

use std::io::Write;
use zip::write::FileOptions;
use zip::{DateTime, ZipWriter};

pub struct Packager {
    zip: ZipWriter<std::fs::File>,
    zip_options: FileOptions,
}

impl Packager {
    // Create a new Packager struct.
    pub fn new(filename: &str) -> Packager {
        let path = std::path::Path::new(filename);
        let file = std::fs::File::create(&path).unwrap();

        let zip = zip::ZipWriter::new(file);

        let zip_options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600)
            .last_modified_time(DateTime::default())
            .large_file(false);

        Packager { zip, zip_options }
    }

    // Create the component xml files and add them to the zip/xlsx container.
    pub fn create_xlsx(&mut self) {
        self.write_content_types_file();
        self.write_root_rels_file();
        self.write_workbook_rels_file();
        self.write_worksheet_files();
        self.write_workbook_file();
        self.write_shared_strings_file();
        self.write_styles_file();
        self.write_theme_file();
        self.write_core_file();
        self.write_app_file();

        // Close the zip file.
        self.zip.finish().unwrap();
    }

    // Write the [ContentTypes].xml file.
    fn write_content_types_file(&mut self) {
        let mut content_types = ContentTypes::new();

        content_types.add_worksheet_name("sheet1");
        content_types.add_share_strings();

        self.zip
            .start_file("[Content_Types].xml", self.zip_options)
            .unwrap();

        content_types.assemble_xml_file();
        let buffer = content_types.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the root level _rels/.rels xml file.
    fn write_root_rels_file(&mut self) {
        let mut rels = Relationship::new();

        rels.add_document_relationship("/officeDocument", "xl/workbook.xml");
        rels.add_package_relationship("/metadata/core-properties", "docProps/core.xml");
        rels.add_document_relationship("/extended-properties", "docProps/app.xml");

        self.zip
            .start_file("_rels/.rels", self.zip_options)
            .unwrap();

        rels.assemble_xml_file();
        let buffer = rels.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the workbook level workbook.xml.rels xml file.
    fn write_workbook_rels_file(&mut self) {
        let mut rels = Relationship::new();
        let worksheet_count = 1;

        for worksheet_index in 1..=worksheet_count {
            rels.add_document_relationship(
                "/worksheet",
                format!("worksheets/sheet{}.xml", worksheet_index).as_str(),
            );
        }

        rels.add_document_relationship("/theme", "theme/theme1.xml");
        rels.add_document_relationship("/styles", "styles.xml");

        rels.add_document_relationship("/sharedStrings", "sharedStrings.xml");

        self.zip
            .start_file("xl/_rels/workbook.xml.rels", self.zip_options)
            .unwrap();

        rels.assemble_xml_file();
        let buffer = rels.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the worksheet xml files.
    fn write_worksheet_files(&mut self) {
        let worksheet_count = 1;

        for worksheet_index in 1..=worksheet_count {
            let mut worksheet = Worksheet::new();

            let filename = format!("xl/worksheets/sheet{}.xml", worksheet_index);

            self.zip.start_file(filename, self.zip_options).unwrap();

            worksheet.assemble_xml_file();
            let buffer = worksheet.writer.read_to_buffer();
            self.zip.write_all(&*buffer).unwrap();
        }
    }

    // Write the workbook.xml file.
    fn write_workbook_file(&mut self) {
        let mut workbook = Workbook::new("");

        self.zip
            .start_file("xl/workbook.xml", self.zip_options)
            .unwrap();

        workbook.assemble_xml_file();
        let buffer = workbook.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the sharedStrings.xml file.
    fn write_shared_strings_file(&mut self) {
        let mut string_table = SharedStringsTable::new();
        let mut shared_strings = SharedStrings::new();
        string_table.get_shared_string_index("Hello");

        self.zip
            .start_file("xl/sharedStrings.xml", self.zip_options)
            .unwrap();

        shared_strings.assemble_xml_file(&string_table);
        let buffer = shared_strings.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the styles.xml file.
    fn write_styles_file(&mut self) {
        let mut styles = Styles::new();

        self.zip
            .start_file("xl/styles.xml", self.zip_options)
            .unwrap();

        styles.assemble_xml_file();
        let buffer = styles.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the theme.xml file.
    fn write_theme_file(&mut self) {
        let mut theme = Theme::new();

        self.zip
            .start_file("xl/theme/theme1.xml", self.zip_options)
            .unwrap();

        theme.assemble_xml_file();
        let buffer = theme.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the core.xml file.
    fn write_core_file(&mut self) {
        let mut core = Core::new();

        self.zip
            .start_file("docProps/core.xml", self.zip_options)
            .unwrap();

        core.assemble_xml_file();
        let buffer = core.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }
    // Write the app.xml file.
    fn write_app_file(&mut self) {
        let mut app = App::new();

        app.add_heading_pair("Worksheets", 1);
        app.add_part_name("Sheet1");

        self.zip
            .start_file("docProps/app.xml", self.zip_options)
            .unwrap();

        app.assemble_xml_file();
        let buffer = app.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }
}
