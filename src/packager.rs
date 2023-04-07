// packager - A library for assembling xml files into an Excel xlsx file.

// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

// This library is used by rust_xlsxwriter to create an Excel xlsx container
// file.
//
// From Wikipedia: The Open Packaging Conventions (OPC) is a container-file
// technology initially created by Microsoft to store a combination of XML and
// non-XML files that together form a single entity such as an Open XML Paper
// Specification (OpenXPS) document.
// http://en.wikipedia.org/wiki/Open_Packaging_Conventions.
//
// At its simplest an Excel xlsx file contains the following elements:
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
// the package and writes them into the xlsx file.

use std::collections::HashSet;
use std::io::{Seek, Write};

use zip::write::FileOptions;
use zip::{DateTime, ZipWriter};

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::custom::Custom;
use crate::error::XlsxError;
use crate::metadata::Metadata;
use crate::relationship::Relationship;
use crate::shared_strings::SharedStrings;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::theme::Theme;
use crate::vml::Vml;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;
use crate::{DocProperties, NUM_IMAGE_FORMATS};

// Packager struct to assembler the xlsx file.
pub struct Packager<W: Write + Seek> {
    zip: ZipWriter<W>,
    zip_options: FileOptions,
}

impl<W: Write + Seek> Packager<W> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Packager struct.
    pub(crate) fn new(writer: W) -> Result<Packager<W>, XlsxError> {
        let zip = zip::ZipWriter::new(writer);

        let zip_options = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600)
            .last_modified_time(DateTime::default())
            .large_file(false);

        Ok(Packager { zip, zip_options })
    }

    // Write the xml files that make up the xlsx OPC package.
    pub(crate) fn assemble_file(
        &mut self,
        workbook: &mut Workbook,
        options: &PackagerOptions,
    ) -> Result<(), XlsxError> {
        // Write the sub-component files.
        self.write_content_types_file(options)?;
        self.write_root_rels_file(options)?;
        self.write_workbook_rels_file(options)?;
        self.write_theme_file()?;
        self.write_styles_file(workbook)?;
        self.write_workbook_file(workbook)?;

        // Write the worksheets and update the shared string table at the same time.
        let mut string_table = SharedStringsTable::new();
        for (index, worksheet) in workbook.worksheets.iter_mut().enumerate() {
            self.write_worksheet_file(worksheet, index + 1, &mut string_table)?;
            if worksheet.has_relationships() {
                self.write_worksheet_rels_file(worksheet, index + 1)?;
            }
        }

        if options.has_sst_table {
            self.write_shared_strings_file(&string_table)?;
        }

        self.write_core_file(options)?;
        self.write_app_file(options)?;
        self.write_custom_file(options)?;

        self.write_drawing_files(workbook)?;
        self.write_vml_files(workbook)?;
        self.write_image_files(workbook)?;
        self.write_chart_files(workbook)?;

        let mut image_index = 1;
        let mut vml_index = 1;

        for worksheet in &mut workbook.worksheets {
            if !worksheet.drawing_relationships.is_empty() {
                self.write_drawing_rels_file(&worksheet.drawing_relationships, image_index)?;
                image_index += 1;
            }
            if !worksheet.vml_drawing_relationships.is_empty() {
                self.write_vml_drawing_rels_file(&worksheet.vml_drawing_relationships, vml_index)?;
                vml_index += 1;
            }
        }

        if options.has_dynamic_arrays {
            self.write_metadata_file()?;
        }

        // Close the zip file.
        self.zip.finish()?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Internal function/methods.
    // -----------------------------------------------------------------------

    // Write the [ContentTypes].xml file.
    fn write_content_types_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut content_types = ContentTypes::new();

        for i in 0..options.num_worksheets {
            content_types.add_worksheet_name(i + 1);
        }

        for i in 0..options.num_drawings {
            content_types.add_drawing_name(i + 1);
        }

        for i in 0..options.num_charts {
            content_types.add_chart_name(i + 1);
        }

        if options.has_sst_table {
            content_types.add_share_strings();
        }

        if options.has_dynamic_arrays {
            content_types.add_metadata();
        }

        if options.has_vml {
            content_types.add_default(
                "vml",
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
            );
        }

        // Add content types for image formats used in the workbook.
        if options.image_types[1] {
            content_types.add_default("png", "image/png");
        }
        if options.image_types[2] {
            content_types.add_default("jpeg", "image/jpeg");
        }
        if options.image_types[3] {
            content_types.add_default("gif", "image/gif");
        }
        if options.image_types[4] {
            content_types.add_default("bmp", "image/bmp");
        }

        if !options.properties.custom_properties.is_empty() {
            content_types.add_custom_properties();
        }

        self.zip
            .start_file("[Content_Types].xml", self.zip_options)?;

        content_types.assemble_xml_file();
        self.zip.write_all(content_types.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the root level _rels/.rels xml file.
    fn write_root_rels_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        rels.add_document_relationship("officeDocument", "xl/workbook.xml", "");
        rels.add_package_relationship("metadata/core-properties", "docProps/core.xml");
        rels.add_document_relationship("extended-properties", "docProps/app.xml", "");

        if !options.properties.custom_properties.is_empty() {
            rels.add_document_relationship("custom-properties", "docProps/custom.xml", "");
        }

        self.zip.start_file("_rels/.rels", self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the workbook level workbook.xml.rels xml file.
    fn write_workbook_rels_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        for worksheet_index in 1..=options.num_worksheets {
            rels.add_document_relationship(
                "worksheet",
                format!("worksheets/sheet{worksheet_index}.xml").as_str(),
                "",
            );
        }

        rels.add_document_relationship("theme", "theme/theme1.xml", "");
        rels.add_document_relationship("styles", "styles.xml", "");

        if options.has_sst_table {
            rels.add_document_relationship("sharedStrings", "sharedStrings.xml", "");
        }

        if options.has_dynamic_arrays {
            rels.add_document_relationship("sheetMetadata", "metadata.xml", "");
        }

        self.zip
            .start_file("xl/_rels/workbook.xml.rels", self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a worksheet xml file.
    pub(crate) fn write_worksheet_file(
        &mut self,
        worksheet: &mut Worksheet,
        index: usize,
        string_table: &mut SharedStringsTable,
    ) -> Result<(), XlsxError> {
        let filename = format!("xl/worksheets/sheet{index}.xml");

        self.zip.start_file(filename, self.zip_options)?;

        worksheet.assemble_xml_file(string_table);
        self.zip.write_all(worksheet.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a worksheet rels file.
    pub(crate) fn write_worksheet_rels_file(
        &mut self,
        worksheet: &Worksheet,
        index: usize,
    ) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        for relationship in &worksheet.hyperlink_relationships {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        for relationship in &worksheet.drawing_object_relationships {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        let filename = format!("xl/worksheets/_rels/sheet{index}.xml.rels");

        self.zip.start_file(filename, self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a drawing rels file.
    pub(crate) fn write_drawing_rels_file(
        &mut self,
        relationships: &[(String, String, String)],
        index: usize,
    ) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        for relationship in relationships.iter() {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        let filename = format!("xl/drawings/_rels/drawing{index}.xml.rels");

        self.zip.start_file(filename, self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a vmlDrawing rels file.
    pub(crate) fn write_vml_drawing_rels_file(
        &mut self,
        relationships: &[(String, String, String)],
        index: usize,
    ) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        for relationship in relationships.iter() {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        let filename = format!("xl/drawings/_rels/vmlDrawing{index}.vml.rels");

        self.zip.start_file(filename, self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the workbook.xml file.
    pub(crate) fn write_workbook_file(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        self.zip.start_file("xl/workbook.xml", self.zip_options)?;

        workbook.assemble_xml_file();
        self.zip.write_all(workbook.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the sharedStrings.xml file.
    pub fn write_shared_strings_file(
        &mut self,
        string_table: &SharedStringsTable,
    ) -> Result<(), XlsxError> {
        let mut shared_strings = SharedStrings::new();

        self.zip
            .start_file("xl/sharedStrings.xml", self.zip_options)?;

        shared_strings.assemble_xml_file(string_table);
        self.zip
            .write_all(shared_strings.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the styles.xml file.
    fn write_styles_file(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut styles = Styles::new(
            &workbook.xf_formats,
            workbook.font_count,
            workbook.fill_count,
            workbook.border_count,
            workbook.num_formats.clone(),
            workbook.has_hyperlink_style,
            false,
        );

        self.zip.start_file("xl/styles.xml", self.zip_options)?;

        styles.assemble_xml_file();
        self.zip.write_all(styles.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the theme.xml file.
    fn write_theme_file(&mut self) -> Result<(), XlsxError> {
        let mut theme = Theme::new();

        self.zip
            .start_file("xl/theme/theme1.xml", self.zip_options)?;

        theme.assemble_xml_file();
        self.zip.write_all(theme.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the core.xml file.
    fn write_core_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut core = Core::new();
        core.properties = options.properties.clone();

        self.zip.start_file("docProps/core.xml", self.zip_options)?;

        core.assemble_xml_file();
        self.zip.write_all(core.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the custom.xml file.
    fn write_custom_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        if options.properties.custom_properties.is_empty() {
            return Ok(());
        }

        let mut custom = Custom::new();
        custom.properties = options.properties.clone();

        self.zip
            .start_file("docProps/custom.xml", self.zip_options)?;

        custom.assemble_xml_file();
        self.zip.write_all(custom.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the app.xml file.
    fn write_app_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut app = App::new();
        app.properties = options.properties.clone();
        app.doc_security = options.doc_security;

        app.add_heading_pair("Worksheets", options.num_worksheets);

        for sheet_name in &options.worksheet_names {
            app.add_part_name(sheet_name);
        }

        if !options.defined_names.is_empty() {
            app.add_heading_pair("Named Ranges", options.defined_names.len() as u16);

            for defined_name in &options.defined_names {
                app.add_part_name(defined_name);
            }
        }

        self.zip.start_file("docProps/app.xml", self.zip_options)?;

        app.assemble_xml_file();
        self.zip.write_all(app.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the metadata.xml file.
    fn write_metadata_file(&mut self) -> Result<(), XlsxError> {
        let mut metadata = Metadata::new();

        self.zip.start_file("xl/metadata.xml", self.zip_options)?;

        metadata.assemble_xml_file();
        self.zip.write_all(metadata.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the drawing files.
    fn write_drawing_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        for worksheet in &mut workbook.worksheets {
            if !worksheet.drawing.drawings.is_empty() {
                let filename = format!("xl/drawings/drawing{index}.xml");
                self.zip.start_file(filename, self.zip_options)?;

                worksheet.drawing.assemble_xml_file();
                self.zip
                    .write_all(worksheet.drawing.writer.xmlfile.get_ref())?;
                index += 1;
            }
        }

        Ok(())
    }

    // Write the vml files.
    fn write_vml_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        for worksheet in &mut workbook.worksheets {
            if worksheet.has_header_footer_images() {
                let filename = format!("xl/drawings/vmlDrawing{index}.vml");
                self.zip.start_file(filename, self.zip_options)?;

                let mut vml = Vml::new();
                vml.header_images
                    .append(&mut worksheet.header_footer_vml_info);
                vml.data_id = index;
                vml.shape_id = 1024 * index;
                vml.assemble_xml_file();

                self.zip.write_all(vml.writer.xmlfile.get_ref())?;
                index += 1;
            }
        }

        Ok(())
    }

    // Write the image files.
    fn write_image_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        let mut unique_worksheet_images = HashSet::new();
        let mut unique_header_footer_images = HashSet::new();

        for worksheet in &mut workbook.worksheets {
            for image in worksheet.images.values() {
                if !unique_worksheet_images.contains(&image.hash) {
                    let filename =
                        format!("xl/media/image{index}.{}", image.image_type.extension());
                    self.zip.start_file(filename, self.zip_options)?;

                    self.zip.write_all(&image.data)?;
                    unique_worksheet_images.insert(image.hash);
                    index += 1;
                }
            }
            if worksheet.has_header_footer_images() {
                for image in worksheet.header_footer_images.clone().into_iter().flatten() {
                    if !unique_header_footer_images.contains(&image.hash) {
                        let filename =
                            format!("xl/media/image{index}.{}", image.image_type.extension());
                        self.zip.start_file(filename, self.zip_options)?;

                        self.zip.write_all(&image.data)?;
                        unique_header_footer_images.insert(image.hash);
                        index += 1;
                    }
                }
            }
        }

        Ok(())
    }

    // Write the chart files.
    fn write_chart_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;

        for worksheet in &mut workbook.worksheets {
            for chart in worksheet.charts.values_mut() {
                let filename = format!("xl/charts/chart{index}.xml");
                self.zip.start_file(filename, self.zip_options)?;
                chart.assemble_xml_file();
                self.zip.write_all(chart.writer.xmlfile.get_ref())?;
                index += 1;
            }
        }

        Ok(())
    }
}

// Internal struct to pass options to the Packager struct.
pub(crate) struct PackagerOptions {
    pub(crate) has_sst_table: bool,
    pub(crate) has_dynamic_arrays: bool,
    pub(crate) has_vml: bool,
    pub(crate) num_worksheets: u16,
    pub(crate) num_drawings: u16,
    pub(crate) num_charts: u16,
    pub(crate) doc_security: u8,
    pub(crate) worksheet_names: Vec<String>,
    pub(crate) defined_names: Vec<String>,
    pub(crate) image_types: [bool; NUM_IMAGE_FORMATS],
    pub(crate) properties: DocProperties,
}

impl PackagerOptions {
    // Create a new PackagerOptions struct.
    pub(crate) fn new() -> PackagerOptions {
        PackagerOptions {
            has_sst_table: false,
            has_dynamic_arrays: false,
            has_vml: false,
            num_worksheets: 0,
            num_drawings: 0,
            num_charts: 0,
            doc_security: 0,
            worksheet_names: vec![],
            defined_names: vec![],
            image_types: [false; NUM_IMAGE_FORMATS],
            properties: DocProperties::new(),
        }
    }
}
