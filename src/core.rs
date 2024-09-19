// core - A module for creating the Excel core.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{
    xml_data_element, xml_data_element_only, xml_declaration, xml_end_tag, xml_start_tag,
};
use crate::DocProperties;

pub struct Core {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) properties: DocProperties,
}

impl Core {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Core struct.
    pub(crate) fn new() -> Core {
        let writer = Cursor::new(Vec::with_capacity(2048));

        Core {
            writer,
            properties: DocProperties::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    // Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        // Write the cp:coreProperties element.
        self.write_cp_core_properties();

        // Write the dc:title element.
        self.write_dc_title();

        // Write the dc:subject element.
        self.write_dc_subject();
        // Write the dc:creator element.
        self.write_dc_creator();

        // Write the cp:keywords element.
        self.write_cp_keywords();

        // Write the dc:description element.
        self.write_dc_description();

        // Write the cp:lastModifiedBy element.
        self.write_cp_last_modified_by();

        // Write the dcterms:created element.
        self.write_dcterms_created();

        // Write the dcterms:modified element.
        self.write_dcterms_modified();

        // Write the cp:category element.
        self.write_cp_category();

        // Write the cp:contentStatus element.
        self.write_cp_content_status();

        // Close the coreProperties tag.
        xml_end_tag(&mut self.writer, "cp:coreProperties");
    }

    // Write the <cp:coreProperties> element.
    fn write_cp_core_properties(&mut self) {
        let xmlns_cp =
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties".to_string();
        let xmlns_dc = "http://purl.org/dc/elements/1.1/".to_string();
        let xmlns_dcterms = "http://purl.org/dc/terms/".to_string();
        let xmlns_dcmitype = "http://purl.org/dc/dcmitype/".to_string();
        let xmlns_xsi = "http://www.w3.org/2001/XMLSchema-instance".to_string();

        let attributes = [
            ("xmlns:cp", xmlns_cp),
            ("xmlns:dc", xmlns_dc),
            ("xmlns:dcterms", xmlns_dcterms),
            ("xmlns:dcmitype", xmlns_dcmitype),
            ("xmlns:xsi", xmlns_xsi),
        ];

        xml_start_tag(&mut self.writer, "cp:coreProperties", &attributes);
    }

    // Write the <dc:title> element.
    fn write_dc_title(&mut self) {
        if !self.properties.title.is_empty() {
            xml_data_element_only(&mut self.writer, "dc:title", &self.properties.title);
        }
    }

    // Write the <dc:subject> element.
    fn write_dc_subject(&mut self) {
        if !self.properties.subject.is_empty() {
            xml_data_element_only(&mut self.writer, "dc:subject", &self.properties.subject);
        }
    }

    // Write the <dc:creator> element.
    fn write_dc_creator(&mut self) {
        xml_data_element_only(&mut self.writer, "dc:creator", &self.properties.author);
    }

    // Write the <cp:keywords> element.
    fn write_cp_keywords(&mut self) {
        if !self.properties.keywords.is_empty() {
            xml_data_element_only(&mut self.writer, "cp:keywords", &self.properties.keywords);
        }
    }

    // Write the <dc:description> element.
    fn write_dc_description(&mut self) {
        if !self.properties.comment.is_empty() {
            xml_data_element_only(&mut self.writer, "dc:description", &self.properties.comment);
        }
    }

    // Write the <cp:lastModifiedBy> element.
    fn write_cp_last_modified_by(&mut self) {
        xml_data_element_only(
            &mut self.writer,
            "cp:lastModifiedBy",
            &self.properties.author,
        );
    }

    // Write the <dcterms:created> element.
    fn write_dcterms_created(&mut self) {
        let attributes = [("xsi:type", "dcterms:W3CDTF")];
        let datetime = self.properties.creation_time.clone();

        xml_data_element(&mut self.writer, "dcterms:created", &datetime, &attributes);
    }

    // Write the <dcterms:modified> element.
    fn write_dcterms_modified(&mut self) {
        let attributes = [("xsi:type", "dcterms:W3CDTF")];
        let datetime = self.properties.creation_time.clone();

        xml_data_element(&mut self.writer, "dcterms:modified", &datetime, &attributes);
    }

    // Write the <cp:category> element.
    fn write_cp_category(&mut self) {
        if !self.properties.category.is_empty() {
            xml_data_element_only(&mut self.writer, "cp:category", &self.properties.category);
        }
    }

    // Write the <cp:contentStatus> element.
    fn write_cp_content_status(&mut self) {
        if !self.properties.status.is_empty() {
            xml_data_element_only(
                &mut self.writer,
                "cp:contentStatus",
                &self.properties.status,
            );
        }
    }
}
