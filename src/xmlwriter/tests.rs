// xmlwriter unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod xmlwriter_tests {

    use core::str;
    use std::io::Cursor;

    use crate::xmlwriter::{
        self, escape_xml_data, escape_xml_escapes, xml_data_element, xml_data_element_only,
        xml_declaration, xml_empty_tag, xml_empty_tag_only, xml_end_tag, xml_si_element,
        xml_start_tag, xml_start_tag_only,
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

    #[test]
    fn test_escape_xml_escapes_no_escape_needed() {
        // Strings without _x should pass through unchanged
        assert_eq!(escape_xml_escapes("hello world"), "hello world");
    }

    #[test]
    fn test_escape_xml_escapes_incomplete_pattern_no_trailing_underscore() {
        // _x followed by 4 hex digits but no trailing underscore should not be escaped
        let input = "/path/to/kernel32_x64d8.pyd";
        assert_eq!(escape_xml_escapes(input), input);
    }

    #[test]
    fn test_escape_xml_escapes_insufficient_length() {
        // _x at end of string with insufficient characters
        let input = "test_x";
        assert_eq!(escape_xml_escapes(input), input);

        let input = "test_x00";
        assert_eq!(escape_xml_escapes(input), input);
    }

    #[test]
    fn test_escape_xml_escapes_complete_pattern() {
        // _xHHHH_ pattern should be escaped
        let input = "test_x0000_end";
        assert_eq!(escape_xml_escapes(input), "test_x005F_x0000_end");
    }

    #[test]
    fn test_escape_xml_escapes_valid_hex_required() {
        // _xGGGG_ with non-hex characters should not be escaped
        let input = "test_xGGGG_end";
        assert_eq!(escape_xml_escapes(input), input);

        // Mixed valid/invalid hex
        let input = "test_x00GG_end";
        assert_eq!(escape_xml_escapes(input), input);
    }

    #[test]
    fn test_escape_xml_escapes_case_insensitive_hex() {
        // Both uppercase and lowercase hex should be escaped
        let input = "test_xABCD_end";
        assert_eq!(escape_xml_escapes(input), "test_x005F_xABCD_end");

        let input = "test_xabcd_end";
        assert_eq!(escape_xml_escapes(input), "test_x005F_xabcd_end");

        let input = "test_xAbCd_end";
        assert_eq!(escape_xml_escapes(input), "test_x005F_xAbCd_end");
    }

    #[test]
    fn test_escape_xml_escapes_multiple_patterns() {
        // Multiple _xHHHH_ patterns should all be escaped
        let input = "_x0000_middle_x0001_";
        assert_eq!(
            escape_xml_escapes(input),
            "_x005F_x0000_middle_x005F_x0001_"
        );
    }

    #[test]
    fn test_escape_xml_escapes_real_world_path_no_escape() {
        // Real-world path that should not be escaped (no trailing underscore)
        let input = "/path/file_x64d8f123xc8c311aa.pyd";
        assert_eq!(escape_xml_escapes(input), input);
    }

    #[test]
    fn test_escape_xml_escapes_real_world_path_with_escape() {
        // Real-world path that should be escaped (has _xHHHH_ pattern)
        let input = "/path/file_x64d8_name.pyd";
        assert_eq!(escape_xml_escapes(input), "/path/file_x005F_x64d8_name.pyd");
    }

    #[test]
    fn test_escape_xml_escapes_at_string_boundaries() {
        // Pattern at start
        let input = "_x0000_rest";
        assert_eq!(escape_xml_escapes(input), "_x005F_x0000_rest");

        // Pattern at end
        let input = "start_x0000_";
        assert_eq!(escape_xml_escapes(input), "start_x005F_x0000_");

        // Pattern is entire string
        let input = "_x0000_";
        assert_eq!(escape_xml_escapes(input), "_x005F_x0000_");
    }

    #[test]
    fn test_escape_xml_escapes_multibyte_characters() {
        // Ensure multibyte UTF-8 characters don't break the logic
        let input = "日本語_x0000_テスト";
        assert_eq!(escape_xml_escapes(input), "日本語_x005F_x0000_テスト");

        // Multibyte chars that might look like hex bytes
        let input = "test_xäöüü_end"; // Not valid hex
        assert_eq!(escape_xml_escapes(input), input);
    }

    #[test]
    fn test_escape_xml_escapes_adjacent_patterns() {
        // Two patterns right next to each other
        let input = "_x0000__x0001_";
        assert_eq!(escape_xml_escapes(input), "_x005F_x0000__x005F_x0001_");
    }

    #[test]
    fn test_escape_xml_data_fffe_and_ffff() {
        // U+FFFE and U+FFFF are Unicode non-characters that should be escaped
        // These are invalid in XML and Excel escapes them with _xHHHH_ format
        let input_fffe = "test\u{FFFE}end";
        assert_eq!(escape_xml_data(input_fffe), "test_xFFFE_end");

        let input_ffff = "test\u{FFFF}end";
        assert_eq!(escape_xml_data(input_ffff), "test_xFFFF_end");

        // Both in the same string
        let input_both = "start\u{FFFE}middle\u{FFFF}end";
        assert_eq!(escape_xml_data(input_both), "start_xFFFE_middle_xFFFF_end");
    }
}
