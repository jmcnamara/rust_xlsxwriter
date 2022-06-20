// core - A module for creating the Excel Core.xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;
use chrono::DateTime;
use chrono::Utc;

pub struct Core<'a> {
    pub writer: &'a mut XMLWriter<'a>,
    author: String,
    create_time: DateTime<Utc>,
}

impl<'a> Core<'a> {
    // Create a new Core struct.
    pub fn new(writer: &'a mut XMLWriter<'a>) -> Core<'a> {
        Core {
            writer,
            author: String::from(""),
            create_time: Utc::now(),
        }
    }

    // Temporary function for testing. This will be replaced with full property
    // handling later.
    #[allow(dead_code)]
    pub fn set_author(&mut self, author: &str) {
        self.author = author.to_string();
    }

    // Temporary function for testing. This will be replaced with full property
    // handling later.
    #[allow(dead_code)]
    pub fn set_create_time(&mut self, create_time: DateTime<Utc>) {
        self.create_time = create_time;
    }

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
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
        let xmlns_cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        let xmlns_dc = "http://purl.org/dc/elements/1.1/";
        let xmlns_dcterms = "http://purl.org/dc/terms/";
        let xmlns_dcmitype = "http://purl.org/dc/dcmitype/";
        let xmlns_xsi = "http://www.w3.org/2001/XMLSchema-instance";

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
        let attributes = vec![("xsi:type", "dcterms:W3CDTF")];
        let datetime = self
            .create_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element_attr("dcterms:created", &datetime, &attributes);
    }

    // Write the <dcterms:modified> element.
    fn write_dcterms_modified(&mut self) {
        let attributes = vec![("xsi:type", "dcterms:W3CDTF")];

        let datetime = self
            .create_time
            .to_rfc3339_opts(chrono::SecondsFormat::Secs, true);

        self.writer
            .xml_data_element_attr("dcterms:modified", &datetime, &attributes);
    }
}

#[cfg(test)]
mod tests {

    use super::Core;
    use super::XMLWriter;

    use chrono::TimeZone;
    use chrono::Utc;
    use pretty_assertions::assert_eq;
    use std::fs::File;
    use std::io::{Read, Seek, SeekFrom};
    use tempfile::tempfile;

    // Convert XML string/doc into a vector for comparison testing.
    pub fn xml_to_vec(xml_string: &str) -> Vec<String> {
        let mut xml_elements: Vec<String> = Vec::new();
        let re = regex::Regex::new(r">\s*<").unwrap();
        let tokens: Vec<&str> = re.split(xml_string).collect();

        for token in &tokens {
            let mut element = token.trim().to_string();

            // Add back the removed brackets.
            if !element.starts_with('<') {
                element = format!("<{}", element);
            }
            if !element.ends_with('>') {
                element = format!("{}>", element);
            }

            xml_elements.push(element);
        }
        xml_elements
    }

    // Test helper to read xml data back from a filehandle.
    fn read_xmlfile_data(tempfile: &mut File) -> String {
        let mut got = String::new();
        tempfile.seek(SeekFrom::Start(0)).unwrap();
        tempfile.read_to_string(&mut got).unwrap();
        got
    }

    #[test]
    fn test_assemble() {
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        let mut core = Core::new(&mut writer);

        core.set_author("A User");
        core.set_create_time(Utc.ymd(2010, 1, 1).and_hms(0, 0, 0));

        core.assemble_xml_file();

        let got = read_xmlfile_data(&mut tempfile);
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
