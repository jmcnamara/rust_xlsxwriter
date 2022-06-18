// xmlwriter - a module for writing XML in the same format and with
// the same escaping as used by Excel in xlsx xml files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
// Copyright 2022, John McNamara, jmcnamara@cpan.org

use std::fs::File;
use std::io::Write;

pub struct XMLWriter<'a> {
    xmlfile: &'a File,
}

impl<'a> XMLWriter<'a> {
    // Create a new XMLWriter struct to write XML to a given filehandle.
    pub fn new(xmlfile: &File) -> XMLWriter {
        XMLWriter { xmlfile }
    }

    // Write an XML file declaration.
    pub fn xml_declaration(&mut self) {
        writeln!(
            &mut self.xmlfile,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
        )
        .expect("Couldn't write to file");
    }

    // Write an XML start tag with attributes.
    pub fn xml_start_tag(&mut self, tag: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}>", tag, attribute_str).expect("Couldn't write to file");
    }

    // Write an XML end tag.
    pub fn xml_end_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, r"</{}>", tag).expect("Couldn't write to file");
    }

    // Write an empty XML tag with attributes.
    #[allow(dead_code)]
    pub fn xml_empty_tag(&mut self, tag: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
            attribute_str.push_str(&pair);
        }

        write!(&mut self.xmlfile, r"<{}{}/>", tag, attribute_str).expect("Couldn't write to file");
    }

    // Write an XML element containing data with optional attributes.
    #[allow(dead_code)]
    pub fn xml_data_element(&mut self, tag: &str, data: &str, attributes: &Vec<(&str, &str)>) {
        let mut attribute_str = String::from("");

        for attribute in attributes {
            let pair = format!(r#" {}="{}""#, attribute.0, escape_attributes(attribute.1));
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
    pub fn xml_si_element(&mut self, string: &str, preserve_whitespace: bool) {
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
fn escape_data(attribute: &str) -> String {
    attribute
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

#[cfg(test)]
mod tests {

    use super::XMLWriter;
    use pretty_assertions::assert_eq;
    use std::fs::File;
    use std::io::{Read, Seek, SeekFrom};
    use tempfile::tempfile;

    fn read_xmlfile_data(tempfile: &mut File) -> String {
        let mut got = String::new();
        tempfile.seek(SeekFrom::Start(0)).unwrap();
        tempfile.read_to_string(&mut got).unwrap();
        got
    }

    #[test]
    fn test_xml_declaration() {
        let expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_declaration();

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag() {
        let expected = "<foo>";
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_start_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_start_tag_with_attributes() {
        let expected = r#"<foo span="8" baz="7">"#;
        let attributes = vec![("span", "8"), ("baz", "7")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_start_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_end_tag() {
        let expected = "</foo>";

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_end_tag("foo");

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag() {
        let expected = "<foo/>";
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_empty_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_empty_tag_with_attributes() {
        let expected = r#"<foo span="8"/>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_empty_tag("foo", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element() {
        let expected = r#"<foo>bar</foo>"#;
        let attributes = vec![];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "bar", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_attributes() {
        let expected = r#"<foo span="8">bar</foo>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "bar", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_data_element_with_escapes() {
        let expected = r#"<foo span="8">&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8")];

        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_data_element("foo", "&<>\"", &attributes);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_si_element() {
        let expected = "<si><t>foo</t></si>";
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_si_element("foo", false);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }

    #[test]
    fn test_xml_si_element_whitespace() {
        let expected = r#"<si><t xml:space="preserve">    foo</t></si>"#;
        let mut tempfile = tempfile().unwrap();
        let mut writer = XMLWriter::new(&tempfile);

        writer.xml_si_element("    foo", true);

        let got = read_xmlfile_data(&mut tempfile);
        assert_eq!(got, expected);
    }
}
