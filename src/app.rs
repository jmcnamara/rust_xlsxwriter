// app - A module for creating the Excel app.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{
    xml_data_element_only, xml_declaration, xml_end_tag, xml_start_tag, xml_start_tag_only,
};
use crate::DocProperties;

pub struct App {
    pub(crate) writer: Cursor<Vec<u8>>,
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
        let writer = Cursor::new(Vec::with_capacity(2048));

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
        if value > 0 {
            self.heading_pairs.push((key.to_string(), value));
        }
    }

    // Add a non-default part name to the file.
    pub(crate) fn add_part_name(&mut self, part_name: &str) {
        self.table_parts.push(part_name.to_string());
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and generate the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

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
        xml_end_tag(&mut self.writer, "Properties");
    }

    // Write the <Properties> element.
    fn write_properties(&mut self) {
        let xmlns =
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties".to_string();
        let xmlns_vt =
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes".to_string();

        let attributes = vec![("xmlns", xmlns), ("xmlns:vt", xmlns_vt)];

        xml_start_tag(&mut self.writer, "Properties", &attributes);
    }

    // Write the <Application> element.
    fn write_application(&mut self) {
        xml_data_element_only(&mut self.writer, "Application", "Microsoft Excel");
    }

    // Write the <DocSecurity> element.
    fn write_doc_security(&mut self) {
        xml_data_element_only(
            &mut self.writer,
            "DocSecurity",
            &self.doc_security.to_string(),
        );
    }

    // Write the <ScaleCrop> element.
    fn write_scale_crop(&mut self) {
        xml_data_element_only(&mut self.writer, "ScaleCrop", "false");
    }

    // Write the <HeadingPairs> element.
    fn write_heading_pairs(&mut self) {
        xml_start_tag_only(&mut self.writer, "HeadingPairs");

        // Write the vt:vector element for heading pairs.
        self.write_heading_vector();

        xml_end_tag(&mut self.writer, "HeadingPairs");
    }

    // Write the <vt:vector> element for heading pairs.
    fn write_heading_vector(&mut self) {
        let size = self.heading_pairs.len() * 2;
        let size = size.to_string();
        let attributes = vec![("size", size), ("baseType", "variant".to_string())];

        xml_start_tag(&mut self.writer, "vt:vector", &attributes);

        for heading_pair in self.heading_pairs.clone() {
            xml_start_tag_only(&mut self.writer, "vt:variant");
            self.write_vt_lpstr(&heading_pair.0);
            xml_end_tag(&mut self.writer, "vt:variant");

            xml_start_tag_only(&mut self.writer, "vt:variant");
            self.write_vt_i4(heading_pair.1);
            xml_end_tag(&mut self.writer, "vt:variant");
        }

        xml_end_tag(&mut self.writer, "vt:vector");
    }

    // Write the <TitlesOfParts> element.
    fn write_titles_of_parts(&mut self) {
        xml_start_tag_only(&mut self.writer, "TitlesOfParts");

        // Write the vt:vector element for title parts.
        self.write_title_parts_vector();

        xml_end_tag(&mut self.writer, "TitlesOfParts");
    }

    // Write the <vt:vector> element for title parts.
    fn write_title_parts_vector(&mut self) {
        let size = self.table_parts.len();
        let size = size.to_string();
        let attributes = vec![("size", size), ("baseType", String::from("lpstr"))];

        xml_start_tag(&mut self.writer, "vt:vector", &attributes);

        for part_name in self.table_parts.clone() {
            self.write_vt_lpstr(&part_name);
        }

        xml_end_tag(&mut self.writer, "vt:vector");
    }

    // Write the <vt:lpstr> element.
    fn write_vt_lpstr(&mut self, data: &str) {
        xml_data_element_only(&mut self.writer, "vt:lpstr", data);
    }

    // Write the <vt:i4> element.
    fn write_vt_i4(&mut self, count: u16) {
        xml_data_element_only(&mut self.writer, "vt:i4", &count.to_string());
    }

    // Write the <Manager> element.
    fn write_manager(&mut self) {
        if !self.properties.manager.is_empty() {
            xml_data_element_only(&mut self.writer, "Manager", &self.properties.manager);
        }
    }

    // Write the <Company> element.
    fn write_company(&mut self) {
        xml_data_element_only(&mut self.writer, "Company", &self.properties.company);
    }

    // Write the <LinksUpToDate> element.
    fn write_links_up_to_date(&mut self) {
        xml_data_element_only(&mut self.writer, "LinksUpToDate", "false");
    }

    // Write the <SharedDoc> element.
    fn write_shared_doc(&mut self) {
        xml_data_element_only(&mut self.writer, "SharedDoc", "false");
    }

    // Write the <HyperlinkBase> element.
    fn write_hyperlink_base(&mut self) {
        if !self.properties.hyperlink_base.is_empty() {
            xml_data_element_only(
                &mut self.writer,
                "HyperlinkBase",
                &self.properties.hyperlink_base,
            );
        }
    }

    // Write the <HyperlinksChanged> element.
    fn write_hyperlinks_changed(&mut self) {
        xml_data_element_only(&mut self.writer, "HyperlinksChanged", "false");
    }

    // Write the <AppVersion> element.
    fn write_app_version(&mut self) {
        xml_data_element_only(&mut self.writer, "AppVersion", "12.0000");
    }
}
