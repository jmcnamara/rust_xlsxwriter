// packager - A library for assembling xml files into an Excel xlsx file.

// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

// This library is used by `rust_xlsxwriter` to create an Excel xlsx container
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

#[cfg(not(target_arch = "wasm32"))]
use std::thread;

use zip::write::SimpleFileOptions;
use zip::{DateTime, ZipWriter};

use crate::app::App;
use crate::content_types::ContentTypes;
use crate::core::Core;
use crate::custom::Custom;
use crate::error::XlsxError;
use crate::metadata::Metadata;
use crate::relationship::Relationship;
use crate::rich_value::RichValue;
use crate::rich_value_rel::RichValueRel;
use crate::rich_value_structure::RichValueStructure;
use crate::rich_value_types::RichValueTypes;
use crate::shared_strings::SharedStrings;
use crate::shared_strings_table::SharedStringsTable;
use crate::styles::Styles;
use crate::theme::Theme;
use crate::vml::Vml;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;
use crate::{Comment, DocProperties, NUM_IMAGE_FORMATS};

// Packager struct to assembler the xlsx file.
pub struct Packager<W: Write + Seek> {
    zip: ZipWriter<W>,
    zip_options: SimpleFileOptions,
    zip_options_for_binary_files: SimpleFileOptions,
}

impl<W: Write + Seek + Send> Packager<W> {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Packager struct.
    pub(crate) fn new(writer: W) -> Packager<W> {
        let zip = zip::ZipWriter::new(writer);

        let zip_options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600)
            .last_modified_time(DateTime::default())
            .large_file(false);

        let zip_options_for_binary_files =
            zip_options.compression_method(zip::CompressionMethod::Stored);

        Packager {
            zip,
            zip_options,
            zip_options_for_binary_files,
        }
    }

    // Write the xml files that make up the xlsx OPC package.
    pub(crate) fn assemble_file(
        mut self,
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

        // Update the shared string table in each worksheet.
        let mut string_table = SharedStringsTable::new();
        for worksheet in &mut workbook.worksheets {
            worksheet.update_string_table_ids(&mut string_table);
        }

        // Assemble, but don't write, the worksheet files in parallel. These are
        // generally the largest files and the threading can help performance if
        // there are multiple large worksheets.
        #[cfg(not(target_arch = "wasm32"))]
        thread::scope(|scope| {
            for worksheet in &mut workbook.worksheets {
                scope.spawn(|| {
                    worksheet.assemble_xml_file();
                });
            }
        });

        // For wasm targets don't use threading.
        #[cfg(target_arch = "wasm32")]
        for worksheet in &mut workbook.worksheets {
            worksheet.assemble_xml_file();
        }

        // Write the worksheet file and and associated rel files.
        for (index, worksheet) in workbook.worksheets.iter_mut().enumerate() {
            self.write_worksheet_file(worksheet, index + 1)?;
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
        self.write_comment_files(workbook)?;
        self.write_image_files(workbook)?;
        self.write_chart_files(workbook)?;
        self.write_table_files(workbook)?;
        self.write_vba_project(workbook)?;

        let mut image_index = 1;

        for worksheet in &mut workbook.worksheets {
            if !worksheet.drawing_relationships.is_empty() {
                self.write_drawing_rels_file(&worksheet.drawing_relationships, image_index)?;
                image_index += 1;
            }
        }

        if options.has_metadata {
            self.write_metadata_file(options)?;
        }

        if options.has_embedded_images {
            self.write_rich_value_rels_file(workbook)?;
            self.write_rich_value_files(workbook, options)?;
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

        // Change the workbook application types based on whether it is an xlsx
        // or xlsm file.
        if options.is_xlsm_file {
            content_types.add_override(
                "/xl/workbook.xml",
                "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
            );

            if options.has_vba_signature {
                content_types.add_override(
                    "/xl/vbaProjectSignature.bin",
                    "application/vnd.ms-office.vbaProjectSignature",
                );
            }
        } else {
            content_types.add_override(
                "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
            );
        }

        for i in 0..options.num_worksheets {
            content_types.add_worksheet_name(i + 1);
        }

        for i in 0..options.num_drawings {
            content_types.add_drawing_name(i + 1);
        }

        for i in 0..options.num_charts {
            content_types.add_chart_name(i + 1);
        }

        for i in 0..options.num_tables {
            content_types.add_table_name(i + 1);
        }

        for i in 0..options.num_comments {
            content_types.add_comments_name(i + 1);
        }

        if options.has_sst_table {
            content_types.add_share_strings();
        }

        if options.has_metadata {
            content_types.add_metadata();
        }

        if options.has_embedded_images {
            content_types.add_rich_value();
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

        if options.is_xlsm_file {
            content_types.add_default("bin", "application/vnd.ms-office.vbaProject");
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

        if options.has_metadata {
            rels.add_document_relationship("sheetMetadata", "metadata.xml", "");
        }

        if options.is_xlsm_file {
            rels.add_office_relationship("2006", "vbaProject", "vbaProject.bin", "");
        }

        if options.has_embedded_images {
            rels.add_office_relationship(
                "2022/10",
                "richValueRel",
                "richData/richValueRel.xml",
                "",
            );

            rels.add_office_relationship("2017/06", "rdRichValue", "richData/rdrichvalue.xml", "");

            rels.add_office_relationship(
                "2017/06",
                "rdRichValueStructure",
                "richData/rdrichvaluestructure.xml",
                "",
            );

            rels.add_office_relationship(
                "2017/06",
                "rdRichValueTypes",
                "richData/rdRichValueTypes.xml",
                "",
            );
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
    ) -> Result<(), XlsxError> {
        let filename = format!("xl/worksheets/sheet{index}.xml");
        self.zip.start_file(filename, self.zip_options)?;
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

        for relationship in &worksheet.table_relationships {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        for relationship in &worksheet.comment_relationships {
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

        for relationship in relationships {
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
        index: u32,
    ) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        for relationship in relationships {
            rels.add_document_relationship(&relationship.0, &relationship.1, &relationship.2);
        }

        let filename = format!("xl/drawings/_rels/vmlDrawing{index}.vml.rels");

        self.zip.start_file(filename, self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a richValueRel.xml.rels file.
    pub(crate) fn write_rich_value_rels_file(
        &mut self,
        workbook: &mut Workbook,
    ) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        let mut index = 1;
        for image in &workbook.embedded_images {
            let target = format!("../media/image{index}.{}", image.image_type.extension());
            rels.add_document_relationship("image", &target, "");
            index += 1;
        }

        let filename = "xl/richData/_rels/richValueRel.xml.rels";

        self.zip.start_file(filename, self.zip_options)?;

        rels.assemble_xml_file();
        self.zip.write_all(rels.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write a vbaProject.bin.rels file.
    pub(crate) fn write_vba_project_rels_file(&mut self) -> Result<(), XlsxError> {
        let mut rels = Relationship::new();

        rels.add_office_relationship("2006", "vbaProjectSignature", "vbaProjectSignature.bin", "");

        let filename = "xl/_rels/vbaProject.bin.rels";

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
    pub(crate) fn write_shared_strings_file(
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
            &workbook.dxf_formats,
            workbook.font_count,
            workbook.fill_count,
            workbook.border_count,
            workbook.num_formats.clone(),
            workbook.has_hyperlink_style,
            workbook.has_comments,
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

        let mut num_worksheets = 0;

        for sheet_name in &options.worksheet_names {
            // Ignore veryHidden worksheets
            if !sheet_name.is_empty() {
                app.add_part_name(sheet_name);
                num_worksheets += 1;
            }
        }

        app.add_heading_pair("Worksheets", num_worksheets);

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
    fn write_metadata_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut metadata = Metadata::new();
        metadata.has_dynamic_functions = options.has_dynamic_functions;
        metadata.has_embedded_images = options.has_embedded_images;
        metadata.num_embedded_images = options.num_embedded_images;

        self.zip.start_file("xl/metadata.xml", self.zip_options)?;

        metadata.assemble_xml_file();
        self.zip.write_all(metadata.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the various RichValue files.
    fn write_rich_value_files(
        &mut self,
        workbook: &Workbook,
        options: &PackagerOptions,
    ) -> Result<(), XlsxError> {
        self.write_rich_value_file(workbook)?;
        self.write_rich_value_types_file()?;
        self.write_rich_value_structure_file(options)?;
        self.write_rich_value_rel_file(options)?;

        Ok(())
    }

    // Write the rdrichvalue.xml file.
    fn write_rich_value_file(&mut self, workbook: &Workbook) -> Result<(), XlsxError> {
        let mut rich_value = RichValue::new(&workbook.embedded_images);

        self.zip
            .start_file("xl/richData/rdrichvalue.xml", self.zip_options)?;

        rich_value.assemble_xml_file();
        self.zip.write_all(rich_value.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the rdRichValueTypes.xml file.
    fn write_rich_value_types_file(&mut self) -> Result<(), XlsxError> {
        let mut rich_value_types = RichValueTypes::new();

        self.zip
            .start_file("xl/richData/rdRichValueTypes.xml", self.zip_options)?;

        rich_value_types.assemble_xml_file();
        self.zip
            .write_all(rich_value_types.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the rdrichvaluestructure.xml file.
    fn write_rich_value_structure_file(
        &mut self,
        options: &PackagerOptions,
    ) -> Result<(), XlsxError> {
        let mut rich_value_structure = RichValueStructure::new();
        rich_value_structure.has_embedded_image_descriptions =
            options.has_embedded_image_descriptions;

        self.zip
            .start_file("xl/richData/rdrichvaluestructure.xml", self.zip_options)?;

        rich_value_structure.assemble_xml_file();
        self.zip
            .write_all(rich_value_structure.writer.xmlfile.get_ref())?;

        Ok(())
    }

    // Write the richValueRel.xml file.
    fn write_rich_value_rel_file(&mut self, options: &PackagerOptions) -> Result<(), XlsxError> {
        let mut rich_value_rel = RichValueRel::new();
        rich_value_rel.num_embedded_images = options.num_embedded_images;

        self.zip
            .start_file("xl/richData/richValueRel.xml", self.zip_options)?;

        rich_value_rel.assemble_xml_file();
        self.zip
            .write_all(rich_value_rel.writer.xmlfile.get_ref())?;

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

    // Write the comment files.
    fn write_comment_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        for worksheet in &mut workbook.worksheets {
            if !worksheet.notes.is_empty() {
                let filename = format!("xl/comments{index}.xml");
                self.zip.start_file(filename, self.zip_options)?;

                let mut comment = Comment::new();
                comment.notes = worksheet.notes.clone();
                comment.note_authors = worksheet.note_authors.keys().cloned().collect();

                comment.assemble_xml_file();

                self.zip.write_all(comment.writer.xmlfile.get_ref())?;
                index += 1;
            }
        }

        Ok(())
    }

    // Write the vml files.
    fn write_vml_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        let mut header_data_id = 1;
        for worksheet in &mut workbook.worksheets {
            if worksheet.has_header_footer_images() || worksheet.has_vml {
                if worksheet.has_vml {
                    let filename = format!("xl/drawings/vmlDrawing{index}.vml");
                    self.zip.start_file(filename, self.zip_options)?;

                    let mut vml = Vml::new();
                    vml.buttons.append(&mut worksheet.buttons_vml_info);
                    vml.comments.append(&mut worksheet.comments_vml_info);

                    vml.data_id.clone_from(&worksheet.vml_data_id);

                    vml.shape_id = worksheet.vml_shape_id;
                    vml.assemble_xml_file();

                    self.zip.write_all(vml.writer.xmlfile.get_ref())?;
                    index += 1;
                }

                if worksheet.has_header_footer_images() {
                    let filename = format!("xl/drawings/vmlDrawing{index}.vml");
                    self.zip.start_file(filename, self.zip_options)?;

                    let mut vml = Vml::new();
                    vml.header_images
                        .append(&mut worksheet.header_footer_vml_info);

                    vml.data_id = format!("{header_data_id}");
                    vml.shape_id = 1024 * header_data_id;
                    header_data_id += 1;

                    vml.assemble_xml_file();

                    self.zip.write_all(vml.writer.xmlfile.get_ref())?;

                    // The rels file index must match the vmlDrawing file index.
                    self.write_vml_drawing_rels_file(&worksheet.vml_drawing_relationships, index)?;

                    index += 1;
                }
            }
        }

        Ok(())
    }

    // Write the image files.
    fn write_image_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;
        let mut unique_worksheet_images = HashSet::new();
        let mut unique_header_footer_images = HashSet::new();

        for image in &workbook.embedded_images {
            let filename = format!("xl/media/image{index}.{}", image.image_type.extension());
            self.zip
                .start_file(filename, self.zip_options_for_binary_files)?;

            self.zip.write_all(&image.data)?;
            index += 1;
        }

        for worksheet in &mut workbook.worksheets {
            for image in worksheet.images.values() {
                if !unique_worksheet_images.contains(&image.hash) {
                    let filename =
                        format!("xl/media/image{index}.{}", image.image_type.extension());
                    self.zip
                        .start_file(filename, self.zip_options_for_binary_files)?;

                    self.zip.write_all(&image.data)?;
                    unique_worksheet_images.insert(image.hash.clone());
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

    // Write the table files.
    fn write_table_files(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        let mut index = 1;

        for worksheet in &mut workbook.worksheets {
            for table in &mut worksheet.tables {
                let filename = format!("xl/tables/table{index}.xml");
                self.zip.start_file(filename, self.zip_options)?;
                table.assemble_xml_file();
                self.zip.write_all(table.writer.xmlfile.get_ref())?;
                index += 1;
            }
        }

        Ok(())
    }

    // Write the vba project file.
    fn write_vba_project(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        if !workbook.is_xlsm_file {
            return Ok(());
        }

        let filename = "xl/vbaProject.bin";
        self.zip
            .start_file(filename, self.zip_options_for_binary_files)?;
        self.zip.write_all(&workbook.vba_project)?;

        // Write the VBA signature file, if present.
        self.write_vba_signature(workbook)
    }

    // Write the vba signature file.
    fn write_vba_signature(&mut self, workbook: &mut Workbook) -> Result<(), XlsxError> {
        if workbook.vba_signature.is_empty() {
            return Ok(());
        }

        let filename = "xl/vbaProjectSignature.bin";
        self.zip
            .start_file(filename, self.zip_options_for_binary_files)?;
        self.zip.write_all(&workbook.vba_signature)?;

        // Write the associated .rels file.
        self.write_vba_project_rels_file()?;

        Ok(())
    }
}

// Internal struct to pass options to the Packager struct.
pub(crate) struct PackagerOptions {
    pub(crate) has_sst_table: bool,
    pub(crate) has_metadata: bool,
    pub(crate) has_dynamic_functions: bool,
    pub(crate) has_embedded_images: bool,
    pub(crate) has_vml: bool,
    pub(crate) is_xlsm_file: bool,
    pub(crate) has_vba_signature: bool,
    pub(crate) num_worksheets: u16,
    pub(crate) num_drawings: u16,
    pub(crate) num_charts: u16,
    pub(crate) num_tables: u16,
    pub(crate) num_comments: u16,
    pub(crate) doc_security: u8,
    pub(crate) worksheet_names: Vec<String>,
    pub(crate) defined_names: Vec<String>,
    pub(crate) image_types: [bool; NUM_IMAGE_FORMATS],
    pub(crate) properties: DocProperties,
    pub(crate) num_embedded_images: u32,
    pub(crate) has_embedded_image_descriptions: bool,
}

impl PackagerOptions {
    // Create a new PackagerOptions struct.
    pub(crate) fn new() -> PackagerOptions {
        PackagerOptions {
            has_sst_table: false,
            has_metadata: false,
            has_dynamic_functions: false,
            has_embedded_images: false,
            has_vml: false,
            is_xlsm_file: false,
            has_vba_signature: false,
            num_worksheets: 0,
            num_drawings: 0,
            num_charts: 0,
            num_tables: 0,
            num_comments: 0,
            doc_security: 0,
            worksheet_names: vec![],
            defined_names: vec![],
            image_types: [false; NUM_IMAGE_FORMATS],
            properties: DocProperties::new(),
            num_embedded_images: 0,
            has_embedded_image_descriptions: false,
        }
    }
}
