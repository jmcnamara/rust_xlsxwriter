// core - A module for creating the Excel core.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;
use chrono::{DateTime, Utc};

pub struct Core {
    pub(crate) writer: XMLWriter,
    author: String,
    creation_time: DateTime<Utc>,
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
            author: "".to_string(),
            creation_time: Utc::now(),
        }
    }

    // Temporary function for testing. This will be replaced with full property
    // handling later.
    #[allow(dead_code)]
    pub(crate) fn set_author(&mut self, author: &str) {
        self.author = author.to_string();
    }

    // Temporary function for testing. This will be replaced with full property
    // handling later.
    #[allow(dead_code)]
    pub(crate) fn set_creation_time(&mut self, create_time: DateTime<Utc>) {
        self.creation_time = create_time;
    }

    // -----------------------------------------------------------------------
    // XML assembly methods.
    // -----------------------------------------------------------------------

    //  Assemble and write the XML file.
    pub(crate) fn assemble_xml_file(&mut self) {
        self.writer.xml_declaration();

        // Write the cp:coreProperties element.
        self.write_cp_core_properties();

        // Write the dc:creator element.
        self.write_dc_creator();

        // Write the cp:lastModifiedBy element.
        self.write_cp_last_modified_by();

        // Write the dcterms:created element.
        self.write_dcterms_created();

        // Write the dcterms:modified element.
        self.write_dcterms_modified();

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

        let attributes = vec![
            ("xmlns:cp", xmlns_cp),
            ("xmlns:dc", xmlns_dc),
            ("xmlns:dcterms", xmlns_dcterms),
            ("xmlns:dcmitype", xmlns_dcmitype),
            ("xmlns:xsi", xmlns_xsi),
        ];

        self.writer
            .xml_start_tag_attr("cp:coreProperties", &attributes);
    }

    // Write the <dc:creator> element.
    fn write_dc_creator(&mut self) {
        self.writer.xml_data_element("dc:creator", &self.author);
    }

    // Write the <cp:lastModifiedBy> element.
    fn write_cp_last_modified_by(&mut self) {
        self.writer
            .xml_data_element("cp:lastModifiedBy", &self.author);
    }

    // Write the <dcterms:created> element.
    fn write_dcterms_created(&mut self) {
        let attributes = vec![("xsi:type", "dcterms:W3CDTF".to_string())];
        let datetime = self
            .creation_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element_attr("dcterms:created", &datetime, &attributes);
    }

    // Write the <dcterms:modified> element.
    fn write_dcterms_modified(&mut self) {
        let attributes = vec![("xsi:type", "dcterms:W3CDTF".to_string())];

        let datetime = self
            .creation_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element_attr("dcterms:modified", &datetime, &attributes);
    }
}

// -----------------------------------------------------------------------
// Tests.
// -----------------------------------------------------------------------
#[cfg(test)]
mod tests {

    use super::Core;
    use crate::test_functions::xml_to_vec;
    use chrono::{TimeZone, Utc};

    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut core = Core::new();

        core.set_author("A User");
        core.set_creation_time(Utc.ymd(2010, 1, 1).and_hms(0, 0, 0));

        core.assemble_xml_file();

        let got = core.writer.read_to_string();
        let got = xml_to_vec(&got);

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

        assert_eq!(got, expected);
    }
}
