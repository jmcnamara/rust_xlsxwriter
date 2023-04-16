// core - A module for creating the Excel core.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::{xmlwriter::XMLWriter, DocProperties};

pub struct Core {
    pub(crate) writer: XMLWriter,
    pub(crate) properties: DocProperties,
}

impl Core {
    // -----------------------------------------------------------------------
    // Crate public methods.
    // -----------------------------------------------------------------------

    // Create a new Core struct.
    pub(crate) fn new() -> Core {
        let writer = XMLWriter::new();

        Core {
            writer,
            properties: DocProperties::new(),
        }
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

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
        self.writer.xml_end_tag("cp:coreProperties");
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

        self.writer.xml_start_tag("cp:coreProperties", &attributes);
    }

    // Write the <dc:title> element.
    fn write_dc_title(&mut self) {
        if !self.properties.title.is_empty() {
            self.writer
                .xml_data_element_only("dc:title", &self.properties.title);
        }
    }

    // Write the <dc:subject> element.
    fn write_dc_subject(&mut self) {
        if !self.properties.subject.is_empty() {
            self.writer
                .xml_data_element_only("dc:subject", &self.properties.subject);
        }
    }

    // Write the <dc:creator> element.
    fn write_dc_creator(&mut self) {
        self.writer
            .xml_data_element_only("dc:creator", &self.properties.author);
    }

    // Write the <cp:keywords> element.
    fn write_cp_keywords(&mut self) {
        if !self.properties.keywords.is_empty() {
            self.writer
                .xml_data_element_only("cp:keywords", &self.properties.keywords);
        }
    }

    // Write the <dc:description> element.
    fn write_dc_description(&mut self) {
        if !self.properties.comment.is_empty() {
            self.writer
                .xml_data_element_only("dc:description", &self.properties.comment);
        }
    }

    // Write the <cp:lastModifiedBy> element.
    fn write_cp_last_modified_by(&mut self) {
        self.writer
            .xml_data_element_only("cp:lastModifiedBy", &self.properties.author);
    }

    // Write the <dcterms:created> element.
    fn write_dcterms_created(&mut self) {
        let attributes = [("xsi:type", "dcterms:W3CDTF")];
        let datetime = self
            .properties
            .creation_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element("dcterms:created", &datetime, &attributes);
    }

    // Write the <dcterms:modified> element.
    fn write_dcterms_modified(&mut self) {
        let attributes = [("xsi:type", "dcterms:W3CDTF")];

        let datetime = self
            .properties
            .creation_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element("dcterms:modified", &datetime, &attributes);
    }

    // Write the <cp:category> element.
    fn write_cp_category(&mut self) {
        if !self.properties.category.is_empty() {
            self.writer
                .xml_data_element_only("cp:category", &self.properties.category);
        }
    }

    // Write the <cp:contentStatus> element.
    fn write_cp_content_status(&mut self) {
        if !self.properties.status.is_empty() {
            self.writer
                .xml_data_element_only("cp:contentStatus", &self.properties.status);
        }
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use crate::core::Core;
    use crate::{test_functions::xml_to_vec, DocProperties};
    use chrono::{TimeZone, Utc};

    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let date = Utc.with_ymd_and_hms(2010, 1, 1, 0, 0, 0).unwrap();
        let properties = DocProperties::new()
            .set_author("A User")
            .set_creation_datetime(&date);

        let mut core = Core::new();
        core.properties = properties;

        core.assemble_xml_file();

        let got = core.writer.read_to_str();
        let got = xml_to_vec(got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
              <dc:creator>A User</dc:creator>
              <cp:lastModifiedBy>A User</cp:lastModifiedBy>
              <dcterms:created xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:created>
              <dcterms:modified xsi:type="dcterms:W3CDTF">2010-01-01T00:00:00Z</dcterms:modified>
            </cp:coreProperties>
            "#,
        );

        assert_eq!(expected, got);
    }
}
