// app - A module for creating the Excel app.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use crate::{xmlwriter::XMLWriter, DocProperties};

pub struct App {
    pub(crate) writer: XMLWriter,
    heading_pairs: Vec<(String, u16)>,
    table_parts: Vec<String>,
    pub(crate) doc_security: u8,
    pub(crate) properties: DocProperties,
}

impl App {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new App struct.
    pub(crate) fn new() -> App {
        let writer = XMLWriter::new();

        App {
            writer,
            heading_pairs: vec![],
            table_parts: vec![],
            doc_security: 0,
            properties: DocProperties::new(),
        }
    }

    // Add a non-default heading pair to the file.
    pub(crate) fn add_heading_pair(&mut self, key: &str, value: u16) {
        self.heading_pairs.push((key.to_string(), value));
    }

    // Add a non-default part name to the file.
    pub(crate) fn add_part_name(&mut self, part_name: &str) {
        self.table_parts.push(part_name.to_string());
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the Properties element.
        self.write_properties();

        // Write the Application element.
        self.write_application();

        // Write the DocSecurity element.
        self.write_doc_security();

        // Write the ScaleCrop element.
        self.write_scale_crop();

        // Write the HeadingPairs element.
        self.write_heading_pairs();

        // Write the TitlesOfParts element.
        self.write_titles_of_parts();

        // Write the Manager element.
        self.write_manager();

        // Write the Company element.
        self.write_company();

        // Write the LinksUpToDate element.
        self.write_links_up_to_date();

        // Write the SharedDoc element.
        self.write_shared_doc();

        // Write the HyperlinkBase element.
        self.write_hyperlink_base();

        // Write the HyperlinksChanged element.
        self.write_hyperlinks_changed();

        // Write the AppVersion element.
        self.write_app_version();

        // Close the Properties tag.
        self.writer.xml_end_tag("Properties");
    }

    // Write the <Properties> element.
    fn write_properties(&mut self) {
        let xmlns =
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties".to_string();
        let xmlns_vt =
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:vt", xmlns_vt)];

        self.writer.xml_start_tag("Properties", &attributes);
    }

    // Write the <Application> element.
    fn write_application(&mut self) {
        self.writer
            .xml_data_element_only("Application", "Microsoft Excel");
    }

    // Write the <DocSecurity> element.
    fn write_doc_security(&mut self) {
        self.writer
            .xml_data_element_only("DocSecurity", &self.doc_security.to_string());
    }

    // Write the <ScaleCrop> element.
    fn write_scale_crop(&mut self) {
        self.writer.xml_data_element_only("ScaleCrop", "false");
    }

    // Write the <HeadingPairs> element.
    fn write_heading_pairs(&mut self) {
        self.writer.xml_start_tag_only("HeadingPairs");

        // Write the vt:vector element for headings.
        self.write_heading_vector();

        self.writer.xml_end_tag("HeadingPairs");
    }

    // Write the <vt:vector> element.
    fn write_heading_vector(&mut self) {
        let size = self.heading_pairs.len() * 2;
        let size = size.to_string();
        let attributes = vec![("size", size), ("baseType", "variant".to_string())];

        self.writer.xml_start_tag("vt:vector", &attributes);

        for heading_pair in self.heading_pairs.clone() {
            self.writer.xml_start_tag_only("vt:variant");
            self.write_vt_lpstr(&heading_pair.0);
            self.writer.xml_end_tag("vt:variant");

            self.writer.xml_start_tag_only("vt:variant");
            self.write_vt_i4(heading_pair.1);
            self.writer.xml_end_tag("vt:variant");
        }

        self.writer.xml_end_tag("vt:vector");
    }

    // Write the <TitlesOfParts> element.
    fn write_titles_of_parts(&mut self) {
        self.writer.xml_start_tag_only("TitlesOfParts");

        self.write_title_parts_vector();

        self.writer.xml_end_tag("TitlesOfParts");
    }

    // Write the <vt:vector> element.
    fn write_title_parts_vector(&mut self) {
        let size = self.table_parts.len();
        let size = size.to_string();
        let attributes = vec![("size", size), ("baseType", String::from("lpstr"))];

        self.writer.xml_start_tag("vt:vector", &attributes);

        for part_name in self.table_parts.clone() {
            self.write_vt_lpstr(&part_name);
        }

        self.writer.xml_end_tag("vt:vector");
    }

    // Write the <vt:lpstr> element.
    fn write_vt_lpstr(&mut self, data: &str) {
        self.writer.xml_data_element_only("vt:lpstr", data);
    }

    // Write the <vt:i4> element.
    fn write_vt_i4(&mut self, count: u16) {
        self.writer
            .xml_data_element_only("vt:i4", &count.to_string());
    }

    // Write the <Manager> element.
    fn write_manager(&mut self) {
        if !self.properties.manager.is_empty() {
            self.writer
                .xml_data_element_only("Manager", &self.properties.manager);
        }
    }

    // Write the <Company> element.
    fn write_company(&mut self) {
        self.writer
            .xml_data_element_only("Company", &self.properties.company);
    }

    // Write the <LinksUpToDate> element.
    fn write_links_up_to_date(&mut self) {
        self.writer.xml_data_element_only("LinksUpToDate", "false");
    }

    // Write the <SharedDoc> element.
    fn write_shared_doc(&mut self) {
        self.writer.xml_data_element_only("SharedDoc", "false");
    }

    // Write the <HyperlinkBase> element.
    fn write_hyperlink_base(&mut self) {
        if !self.properties.hyperlink_base.is_empty() {
            self.writer
                .xml_data_element_only("HyperlinkBase", &self.properties.hyperlink_base);
        }
    }

    // Write the <HyperlinksChanged> element.
    fn write_hyperlinks_changed(&mut self) {
        self.writer
            .xml_data_element_only("HyperlinksChanged", "false");
    }

    // Write the <AppVersion> element.
    fn write_app_version(&mut self) {
        self.writer.xml_data_element_only("AppVersion", "12.0000");
    }
}
