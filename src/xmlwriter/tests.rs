// xmlwriter unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod xmlwriter_tests {

    use core::str;
    use std::io::Cursor;

    use crate::xmlwriter::{
        self, xml_data_element, xml_data_element_only, xml_declaration, xml_empty_tag,
        xml_empty_tag_only, xml_end_tag, xml_si_element, xml_start_tag, xml_start_tag_only,
    };

    use pretty_assertions::assert_eq;

    #[test]
    fn test_xml_declaration() {
        let mut writer = Cursor::new(Vec::<u8>::with_capacity(2048));

        let expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

        xml_declaration(&mut writer);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_start_tag_without_attributes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = "<foo>";

        xml_start_tag_only(&mut writer, "foo");

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }
    #[test]
    fn test_xml_start_tag_without_attributes_implicit() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = "<foo>";
        let attributes: Vec<(&str, &str)> = vec![];

        xml_start_tag(&mut writer, "foo", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_start_tag_with_attributes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo span="8" baz="7">"#;
        let attributes = vec![("span", "8"), ("baz", "7")];

        xml_start_tag(&mut writer, "foo", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_end_tag() {
        let expected = "</foo>";

        let mut writer = Cursor::new(Vec::with_capacity(2048));

        xml_end_tag(&mut writer, "foo");

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_empty_tag() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = "<foo/>";

        xml_empty_tag_only(&mut writer, "foo");

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_empty_tag_with_attributes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo span="8"/>"#;
        let attributes = [("span", "8")];

        xml_empty_tag(&mut writer, "foo", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo>bar</foo>"#;

        xml_data_element_only(&mut writer, "foo", "bar");

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_attributes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo span="8">bar</foo>"#;
        let attributes = [("span", "8")];

        xml_data_element(&mut writer, "foo", "bar", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_escapes() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo span="8">&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8")];

        xml_data_element(&mut writer, "foo", "&<>\"", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_escapes_non_ascii() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<foo span="8" text="Ы&amp;&lt;&gt;&quot;&#xA;">Ы&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8"), ("text", "Ы&<>\"\n")];

        xml_data_element(&mut writer, "foo", "Ы&<>\"", &attributes);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_si_element() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = "<si><t>foo</t></si>";

        xml_si_element(&mut writer, "foo", false);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_si_element_whitespace() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = r#"<si><t xml:space="preserve">    foo</t></si>"#;

        xml_si_element(&mut writer, "    foo", true);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_si_escape_with_unicode() {
        let mut writer = Cursor::new(Vec::with_capacity(2048));

        let expected = "<si><t>_x_1_½</t></si>";

        xml_si_element(&mut writer, "_x_1_½", false);

        let got = xmlwriter::cursor_to_str(&writer);
        assert_eq!(expected, got);
    }
}
