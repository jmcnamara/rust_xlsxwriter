// content_types - A module for creating the Excel [Content_Types].xml file.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use crate::xmlwriter::XMLWriter;

pub struct ContentTypes {
    pub writer: XMLWriter,
    defaults: Vec<(String, String)>,
    overrides: Vec<(String, String)>,
}

impl ContentTypes {
    // Create a new ContentTypes struct.
    pub fn new() -> ContentTypes {
        let writer = XMLWriter::new();

        ContentTypes {
            writer,
            defaults: vec![
                (
                    String::from("rels"),
                    String::from("application/vnd.openxmlformats-package.relationships+xml"),
                ),
                (String::from("xml"), String::from("application/xml")),
            ],

            overrides: vec![
                (
                    String::from("/docProps/app.xml"),
                    String::from(
                        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                    ),
                ),
                (
                    String::from("/docProps/core.xml"),
                    String::from("application/vnd.openxmlformats-package.core-properties+xml"),
                ),
                (
                    String::from("/xl/styles.xml"),
                    String::from(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                    ),
                ),
                (
                    String::from("/xl/theme/theme1.xml"),
                    String::from("application/vnd.openxmlformats-officedocument.theme+xml"),
                ),
                (
                    String::from("/xl/workbook.xml"),
                    String::from("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"),
                ),
            ],
        }
    }

    // Add elements to the ContentTypes defaults.
    pub fn add_default(&mut self, extension: &str, content_type: &str) {
        self.defaults
            .push((extension.to_string(), content_type.to_string()));
    }

    // Add elements to the ContentTypes overrides.
    fn add_override(&mut self, part_name: &str, content_type: &str) {
        self.overrides
            .push((part_name.to_string(), content_type.to_string()));
    }

    // Add the name of a worksheet to the ContentTypes overrides.
    pub fn add_worksheet_name(&mut self, worksheet_name: &str) {
        let content_type =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        let part_name = format!("/xl/worksheets/{}.xml", worksheet_name);

        self.add_override(&part_name, content_type);
    }

    // Add the sharedStrings link to the ContentTypes overrides.
    pub fn add_share_strings(&mut self) {
        self.add_override(
            "/xl/sharedStrings.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        );
    }

    //  Assemble and write the XML file.
    pub fn assemble_xml_file(&mut self) {
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
        let attributes = vec![("xmlns", xmlns)];

        self.writer.xml_start_tag_attr("Types", &attributes);
    }
    // Write all the <Default> elements.
    fn write_defaults(&mut self) {
        for pair in self.defaults.clone() {
            self.write_default(pair.0, pair.1);
        }
    }

    // Write the <Default> element.
    fn write_default(&mut self, extension: String, content_type: String) {
        let attributes = vec![
            ("Extension", extension.as_str()),
            ("ContentType", content_type.as_str()),
        ];

        self.writer.xml_empty_tag_attr("Default", &attributes);
    }

    // Write all the <Default> elements.
    fn write_overrides(&mut self) {
        for pair in self.overrides.clone() {
            self.write_override(pair.0, pair.1);
        }
    }

    // Write the <Override> element.
    fn write_override(&mut self, part_name: String, content_type: String) {
        let attributes = vec![
            ("PartName", part_name.as_str()),
            ("ContentType", content_type.as_str()),
        ];

        self.writer.xml_empty_tag_attr("Override", &attributes);
    }
}

#[cfg(test)]
mod tests {

    use super::ContentTypes;
    use crate::test_functions::xml_to_vec;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_assemble() {
        let mut content_types = ContentTypes::new();

        content_types.add_default("jpeg", "image/jpeg");
        content_types.add_worksheet_name("sheet1");
        content_types.add_share_strings();
        content_types.assemble_xml_file();

        let got = content_types.writer.read_to_string();
        let got = xml_to_vec(&got);

        let expected = xml_to_vec(
            r#"
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">

              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="jpeg" ContentType="image/jpeg"/>

              <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
              <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
              <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
              <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
              <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
             </Types>
                "#,
        );

        assert_eq!(got, expected);
    }
}
