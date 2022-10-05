// xmlwriter - a module for writing XML in the same format and with the same
// escaping as used by Excel in xlsx xml files. This is a base "class" or set of
// functionality for all of the other xml writing structs.
//
// SPDX-License-Identifier: MIT OR Apache-2.0 Copyright 2022, John McNamara,
// jmcnamara@cpan.org

use std::fs::File;
use std::io::{BufWriter, Read, Seek, Write};

use tempfile::tempfile;

pub struct XMLWriter {
    pub(crate) xmlfile: BufWriter<File>,
}

impl Default for XMLWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl XMLWriter {
    // Create a new XMLWriter struct to write XML to a given filehandle.
    pub(crate) fn new() -> XMLWriter {
        let xmlfile = tempfile().unwrap();
        let xmlfile = BufWriter::new(xmlfile);
        XMLWriter { xmlfile }
    }

    // Helper function for tests to read xml data back from the xml filehandle.
    #[allow(dead_code)]
    pub(crate) fn read_to_string(&mut self) -> String {
        let mut xml_string = String::new();
        self.xmlfile.rewind().unwrap();
        self.xmlfile
            .get_ref()
            .read_to_string(&mut xml_string)
            .unwrap();
        xml_string
    }

    // Return data to write to xlsx/zip file member.
    pub(crate) fn read_to_buffer(&mut self) -> Vec<u8> {
        let mut buffer = Vec::new();
        self.xmlfile.rewind().unwrap();
        self.xmlfile.get_ref().read_to_end(&mut buffer).unwrap();
        buffer
    }

    // Write an XML file declaration.
    pub(crate) fn xml_declaration(&mut self) {
        writeln!(
            &mut self.xmlfile,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        )
        .expect("Couldn't write to file");
    }

    // Write an XML start tag without attributes.
    pub(crate) fn xml_start_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, r"<{}>", tag).expect("Couldn't write to file");
    }

    // Write an XML start tag with attributes.
    pub(crate) fn xml_start_tag_attr(&mut self, tag: &str, attributes: &Vec<(&str, String)>) {
        let mut attribute_str = "".to_string();

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(&attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}>", tag, attribute_str).expect("Couldn't write to file");
    }

    // Write an XML end tag.
    pub(crate) fn xml_end_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, r"</{}>", tag).expect("Couldn't write to file");
    }

    // Write an empty XML tag without attributes.
    pub(crate) fn xml_empty_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, r"<{}/>", tag).expect("Couldn't write to file");
    }

    // Write an empty XML tag with attributes.
    pub(crate) fn xml_empty_tag_attr(&mut self, tag: &str, attributes: &Vec<(&str, String)>) {
        let mut attribute_str = "".to_string();

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(&attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}/>", tag, attribute_str).expect("Couldn't write to file");
    }

    // Write an XML element containing data without attributes.
    pub(crate) fn xml_data_element(&mut self, tag: &str, data: &str) {
        write!(
            &mut self.xmlfile,
            r"<{}>{}</{}>",
            tag,
            escape_data(data),
            tag
        )
        .expect("Couldn't write to file");
    }
    // Write an XML element containing data with attributes.
    pub(crate) fn xml_data_element_attr(
        &mut self,
        tag: &str,
        data: &str,
        attributes: &Vec<(&str, String)>,
    ) {
        let mut attribute_str = "".to_string();

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(&attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(
            &mut self.xmlfile,
            r"<{}{}>{}</{}>",
            tag,
            attribute_str,
            escape_data(data),
            tag
        )
        .expect("Couldn't write to file");
    }

    // Optimized tag writer for shared strings `<si>` elements.
    pub(crate) fn xml_si_element(&mut self, string: &str, preserve_whitespace: bool) {
        if preserve_whitespace {
            write!(
                &mut self.xmlfile,
                r#"<si><t xml:space="preserve">{}</t></si>"#,
                escape_data(string)
            )
            .expect("Couldn't write to file");
        } else {
            write!(&mut self.xmlfile, "<si><t>{}</t></si>", escape_data(string))
                .expect("Couldn't write to file");
        }
    }

    // Write the theme string to the theme file.
    pub(crate) fn write_theme(&mut self, theme: &str) {
        writeln!(&mut self.xmlfile, "{}", theme).expect("Couldn't write to file");
    }
}

// Escape XML characters in attributes.
fn escape_attributes(attribute: &str) -> String {
    attribute
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\n', "&#xA;")
}

// Escape XML characters in data sections of tags.  Note, this
// is different from escape_attributes() because double quotes
// and newline are not escaped by Excel.
pub(crate) fn escape_data(attribute: &str) -> String {
    attribute
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {

    use super::XMLWriter;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_xml_declaration() {
        let expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

        let mut writer = XMLWriter::new();
        writer.xml_declaration();

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag_without_attributes() {
        let expected = "<foo>";

        let mut writer = XMLWriter::new();
        writer.xml_start_tag("foo");

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }
    #[test]
    fn test_xml_start_tag_without_attributes_implicit() {
        let expected = "<foo>";
        let attributes = vec![];

        let mut writer = XMLWriter::new();
        writer.xml_start_tag_attr("foo", &attributes);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag_with_attributes() {
        let expected = r#"<foo span="8" baz="7">"#;
        let attributes = vec![("span", "8".to_string()), ("baz", "7".to_string())];

        let mut writer = XMLWriter::new();
        writer.xml_start_tag_attr("foo", &attributes);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_end_tag() {
        let expected = "</foo>";

        let mut writer = XMLWriter::new();

        writer.xml_end_tag("foo");

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag() {
        let expected = "<foo/>";

        let mut writer = XMLWriter::new();

        writer.xml_empty_tag("foo");

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag_with_attributes() {
        let expected = r#"<foo span="8"/>"#;
        let attributes = vec![("span", "8".to_string())];

        let mut writer = XMLWriter::new();

        writer.xml_empty_tag_attr("foo", &attributes);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element() {
        let expected = r#"<foo>bar</foo>"#;

        let mut writer = XMLWriter::new();
        writer.xml_data_element("foo", "bar");

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_attributes() {
        let expected = r#"<foo span="8">bar</foo>"#;
        let attributes = vec![("span", "8".to_string())];

        let mut writer = XMLWriter::new();
        writer.xml_data_element_attr("foo", "bar", &attributes);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_escapes() {
        let expected = r#"<foo span="8">&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8".to_string())];

        let mut writer = XMLWriter::new();
        writer.xml_data_element_attr("foo", "&<>\"", &attributes);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_si_element() {
        let expected = "<si><t>foo</t></si>";

        let mut writer = XMLWriter::new();
        writer.xml_si_element("foo", false);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_si_element_whitespace() {
        let expected = r#"<si><t xml:space="preserve">    foo</t></si>"#;

        let mut writer = XMLWriter::new();
        writer.xml_si_element("    foo", true);

        let got = writer.read_to_string();
        assert_eq!(got, expected);
    }
}
