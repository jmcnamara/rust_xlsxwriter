// Xmlwriter unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod xmlwriter_tests {

    use crate::xmlwriter::XMLWriter;
    use pretty_assertions::assert_eq;

    #[test]
    fn test_xml_declaration() {
        let expected = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

        let mut writer = XMLWriter::default();
        writer.xml_declaration();

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_start_tag_without_attributes() {
        let expected = "<foo>";

        let mut writer = XMLWriter::new();
        writer.xml_start_tag_only("foo");

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }
    #[test]
    fn test_xml_start_tag_without_attributes_implicit() {
        let expected = "<foo>";
        let attributes: Vec<(&str, &str)> = vec![];

        let mut writer = XMLWriter::new();
        writer.xml_start_tag("foo", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_start_tag_with_attributes() {
        let expected = r#"<foo span="8" baz="7">"#;
        let attributes = vec![("span", "8"), ("baz", "7")];

        let mut writer = XMLWriter::new();
        writer.xml_start_tag("foo", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_end_tag() {
        let expected = "</foo>";

        let mut writer = XMLWriter::new();

        writer.xml_end_tag("foo");

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_empty_tag() {
        let expected = "<foo/>";

        let mut writer = XMLWriter::new();

        writer.xml_empty_tag_only("foo");

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_empty_tag_with_attributes() {
        let expected = r#"<foo span="8"/>"#;
        let attributes = [("span", "8")];

        let mut writer = XMLWriter::new();

        writer.xml_empty_tag("foo", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element() {
        let expected = r#"<foo>bar</foo>"#;

        let mut writer = XMLWriter::new();
        writer.xml_data_element_only("foo", "bar");

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_attributes() {
        let expected = r#"<foo span="8">bar</foo>"#;
        let attributes = [("span", "8")];

        let mut writer = XMLWriter::new();
        writer.xml_data_element("foo", "bar", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_escapes() {
        let expected = r#"<foo span="8">&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8")];

        let mut writer = XMLWriter::new();
        writer.xml_data_element("foo", "&<>\"", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_data_element_with_escapes_non_ascii() {
        let expected = r#"<foo span="8" text="Ы&amp;&lt;&gt;&quot;&#xA;">Ы&amp;&lt;&gt;"</foo>"#;
        let attributes = vec![("span", "8"), ("text", "Ы&<>\"\n")];

        let mut writer = XMLWriter::new();
        writer.xml_data_element("foo", "Ы&<>\"", &attributes);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_si_element() {
        let expected = "<si><t>foo</t></si>";

        let mut writer = XMLWriter::new();
        writer.xml_si_element("foo", false);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }

    #[test]
    fn test_xml_si_element_whitespace() {
        let expected = r#"<si><t xml:space="preserve">    foo</t></si>"#;

        let mut writer = XMLWriter::new();
        writer.xml_si_element("    foo", true);

        let got = writer.read_to_str();
        assert_eq!(expected, got);
    }
}
