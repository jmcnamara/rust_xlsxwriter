// packager - A library for assembling xml files into an Excel XLSX file.

// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

// This library is used by rust_xlsxwriter to create an Excel XLSX container
// file.
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
// The Packager struct coordinates the classes that represent the elements of
// the package and writes them into the XLSX file.

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::error::XlsxError;
use crate::format::Format;
use crate::metadata::Metadata;
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

// Packager struct to assembler the xlsx file.
pub struct Packager {
    zip: ZipWriter<std::fs::File>,
    zip_options: FileOptions,
}

impl Packager {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Packager struct.
    pub(crate) fn new(filename: &str) -> Result<Packager, XlsxError> {
        let path = std::path::Path::new(filename);

        let file = match std::fs::File::create(&path) {
            Ok(file) => file,
            Err(e) => {
                return Err(XlsxError::IoError(format!(
                    "Error creating output file {}: {}",
                    filename, e
                )))
            }
        };

        let zip = zip::ZipWriter::new(file);

        let zip_options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600)
            .last_modified_time(DateTime::default())
            .large_file(false);

        Ok(Packager { zip, zip_options })
    }

    // Create the root and xl/ component xml files and add them to the zip/xlsx
    // container.
    pub(crate) fn create_root_files(&mut self, options: &PackagerOptions) {
        self.write_content_types_file(options);
        self.write_root_rels_file();
        self.write_workbook_rels_file(options);
        self.write_theme_file();
    }

    // Create the styles.xml file and add it to the zip/xlsx container.
    pub(crate) fn create_styles_file(
        &mut self,
        xf_formats: &Vec<Format>,
        font_count: u16,
        fill_count: u16,
        border_count: u16,
        num_format_count: u16,
    ) {
        self.write_styles_file(
            xf_formats,
            font_count,
            fill_count,
            border_count,
            num_format_count,
        );
    }

    // Create the docProps component xml files and add them to the zip/xlsx
    // container.
    pub(crate) fn create_doc_prop_files(&mut self, options: &PackagerOptions) {
        self.write_core_file();
        self.write_app_file(options);

        if options.has_dynamic_arrays {
            self.write_metadata_file();
        }
    }

    // Close the zip file.
    pub(crate) fn close(&mut self) {
        self.zip.finish().unwrap();
    }

    // Write the [ContentTypes].xml file.
    fn write_content_types_file(&mut self, options: &PackagerOptions) {
        let mut content_types = ContentTypes::new();

        for i in 0..options.num_worksheets {
            content_types.add_worksheet_name(format!("sheet{}", i + 1).as_str());
        }

        if options.has_sst_table {
            content_types.add_share_strings();
        }

        if options.has_dynamic_arrays {
            content_types.add_metadata();
        }

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
    fn write_workbook_rels_file(&mut self, options: &PackagerOptions) {
        let mut rels = Relationship::new();

        for worksheet_index in 1..=options.num_worksheets {
            rels.add_document_relationship(
                "/worksheet",
                format!("worksheets/sheet{}.xml", worksheet_index).as_str(),
            );
        }

        rels.add_document_relationship("/theme", "theme/theme1.xml");
        rels.add_document_relationship("/styles", "styles.xml");

        if options.has_sst_table {
            rels.add_document_relationship("/sharedStrings", "sharedStrings.xml");
        }

        if options.has_dynamic_arrays {
            rels.add_document_relationship("/sheetMetadata", "metadata.xml");
        }

        self.zip
            .start_file("xl/_rels/workbook.xml.rels", self.zip_options)
            .unwrap();

        rels.assemble_xml_file();
        let buffer = rels.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write a worksheet xml file.
    pub(crate) fn write_worksheet_file(
        &mut self,
        worksheet: &mut Worksheet,
        index: usize,
        string_table: &mut SharedStringsTable,
    ) {
        let filename = format!("xl/worksheets/sheet{}.xml", index);

        self.zip.start_file(filename, self.zip_options).unwrap();

        worksheet.assemble_xml_file(string_table);
        let buffer = worksheet.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the workbook.xml file.
    pub(crate) fn write_workbook_file(&mut self, workbook: &mut Workbook) {
        self.zip
            .start_file("xl/workbook.xml", self.zip_options)
            .unwrap();

        workbook.assemble_xml_file();
        let buffer = workbook.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the sharedStrings.xml file.
    pub fn write_shared_strings_file(&mut self, string_table: &SharedStringsTable) {
        let mut shared_strings = SharedStrings::new();

        self.zip
            .start_file("xl/sharedStrings.xml", self.zip_options)
            .unwrap();

        shared_strings.assemble_xml_file(string_table);
        let buffer = shared_strings.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the styles.xml file.
    fn write_styles_file(
        &mut self,
        xf_formats: &Vec<Format>,
        font_count: u16,
        fill_count: u16,
        border_count: u16,
        num_format_count: u16,
    ) {
        let mut styles = Styles::new(
            xf_formats,
            font_count,
            fill_count,
            border_count,
            num_format_count,
        );

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
    fn write_app_file(&mut self, options: &PackagerOptions) {
        let mut app = App::new();

        app.add_heading_pair("Worksheets", options.num_worksheets);

        for sheet_name in &options.worksheet_names {
            app.add_part_name(sheet_name);
        }

        self.zip
            .start_file("docProps/app.xml", self.zip_options)
            .unwrap();

        app.assemble_xml_file();
        let buffer = app.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }

    // Write the metadata.xml file.
    fn write_metadata_file(&mut self) {
        let mut metadata = Metadata::new();

        self.zip
            .start_file("xl/metadata.xml", self.zip_options)
            .unwrap();

        metadata.assemble_xml_file();
        let buffer = metadata.writer.read_to_buffer();
        self.zip.write_all(&*buffer).unwrap();
    }
}

// Internal struct to pass options to the Packager struct.
pub struct PackagerOptions {
    pub has_sst_table: bool,
    pub has_dynamic_arrays: bool,
    pub num_worksheets: u16,
    pub worksheet_names: Vec<String>,
}

impl PackagerOptions {
    // Create a new PackagerOptions struct.
    pub(crate) fn new() -> PackagerOptions {
        PackagerOptions {
            has_sst_table: false,
            has_dynamic_arrays: false,
            num_worksheets: 0,
            worksheet_names: vec![],
        }
    }
}
