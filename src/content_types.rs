// content_types - A module for creating the Excel [Content_Types].xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use crate::xmlwriter::XMLWriter;

pub struct ContentTypes {
    pub(crate) writer: XMLWriter,
    defaults: Vec<(String, String)>,
    overrides: Vec<(String, String)>,
}

impl ContentTypes {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new ContentTypes struct.
    pub(crate) fn new() -> ContentTypes {
        let writer = XMLWriter::new();

        ContentTypes {
            writer,
            defaults: vec![
                (
                    "rels".to_string(),
                    "application/vnd.openxmlformats-package.relationships+xml".to_string(),
                ),
                ("xml".to_string(), "application/xml".to_string()),
            ],

            overrides: vec![
                (
                    "/docProps/app.xml".to_string(),
                    "application/vnd.openxmlformats-officedocument.extended-properties+xml"
                        .to_string(),
                ),
                (
                    "/docProps/core.xml".to_string(),
                    "application/vnd.openxmlformats-package.core-properties+xml".to_string(),
                ),
                (
                    "/xl/styles.xml".to_string(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
                        .to_string(),
                ),
                (
                    "/xl/theme/theme1.xml".to_string(),
                    "application/vnd.openxmlformats-officedocument.theme+xml".to_string(),
                ),
            ],
        }
    }

    // Add elements to the ContentTypes defaults.
    pub(crate) fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .push((extension.to_string(), content_type.to_string()));
    }

    // Add elements to the ContentTypes overrides.
    pub(crate) fn add_override(&mut self, part_name: &str, content_type: &str) {
        self.overrides
            .push((part_name.to_string(), content_type.to_string()));
    }

    // Add the name of a worksheet to the ContentTypes overrides.
    pub(crate) fn add_worksheet_name(&mut self, index: u16) {
        let content_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        let part_name = format!("/xl/worksheets/sheet{index}.xml");

        self.add_override(&part_name, content_type);
    }

    // Add the name of a drawing to the ContentTypes overrides.
    pub(crate) fn add_drawing_name(&mut self, index: u16) {
        let content_type = "application/vnd.openxmlformats-officedocument.drawing+xml";
        let part_name = format!("/xl/drawings/drawing{index}.xml");

        self.add_override(&part_name, content_type);
    }

    // Add the name of a chart to the ContentTypes overrides.
    pub(crate) fn add_chart_name(&mut self, index: u16) {
        let content_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        let part_name = format!("/xl/charts/chart{index}.xml");

        self.add_override(&part_name, content_type);
    }

    // Add the name of a table to the ContentTypes overrides.
    pub(crate) fn add_table_name(&mut self, index: u16) {
        let content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        let part_name = format!("/xl/tables/table{index}.xml");

        self.add_override(&part_name, content_type);
    }

    // Add the name of a comment file to the ContentTypes overrides.
    pub(crate) fn add_comments_name(&mut self, index: u16) {
        let content_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        let part_name = format!("/xl/comments{index}.xml");

        self.add_override(&part_name, content_type);
    }

    // Add the sharedStrings link to the ContentTypes overrides.
    pub(crate) fn add_share_strings(&mut self) {
        self.add_override(
            "/xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        );
    }

    // Add the metadata file to the ContentTypes overrides.
    pub(crate) fn add_metadata(&mut self) {
        self.add_override(
            "/xl/metadata.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
        );
    }

    // Add the rdRichValue* files to the ContentTypes overrides.
    pub(crate) fn add_rich_value(&mut self) {
        self.add_override(
            "/xl/richData/rdRichValueTypes.xml",
            "application/vnd.ms-excel.rdrichvaluetypes+xml",
        );
        self.add_override(
            "/xl/richData/rdrichvalue.xml",
            "application/vnd.ms-excel.rdrichvalue+xml",
        );
        self.add_override(
            "/xl/richData/rdrichvaluestructure.xml",
            "application/vnd.ms-excel.rdrichvaluestructure+xml",
        );
        self.add_override(
            "/xl/richData/richValueRel.xml",
            "application/vnd.ms-excel.richvaluerel+xml",
        );
    }

    // Add the custom properties to the ContentTypes overrides.
    pub(crate) fn add_custom_properties(&mut self) {
        self.add_override(
            "/docProps/custom.xml",
            "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        );
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the Types element.
        self.write_types();

        // Write the Default element.
        self.write_defaults();

        // Write the Override element.
        self.write_overrides();

        // Close the Types tag.
        self.writer.xml_end_tag("Types");
    }

    // Write the <Types> element.
    fn write_types(&mut self) {
        let xmlns = "http://schemas.openxmlformats.org/package/2006/content-types";
        let attributes = [("xmlns", xmlns)];

        self.writer.xml_start_tag("Types", &attributes);
    }
    // Write all the <Default> elements.
    fn write_defaults(&mut self) {
        for pair in self.defaults.clone() {
            self.write_default(pair.0, pair.1);
        }
    }

    // Write the <Default> element.
    fn write_default(&mut self, extension: String, content_type: String) {
        let attributes = [("Extension", extension), ("ContentType", content_type)];

        self.writer.xml_empty_tag("Default", &attributes);
    }

    // Write all the <Default> elements.
    fn write_overrides(&mut self) {
        for pair in self.overrides.clone() {
            self.write_override(pair.0, pair.1);
        }
    }

    // Write the <Override> element.
    fn write_override(&mut self, part_name: String, content_type: String) {
        let attributes = [("PartName", part_name), ("ContentType", content_type)];

        self.writer.xml_empty_tag("Override", &attributes);
    }
}
