// xmlwriter - a module for writing XML in the same format and with the same
// escaping as used by Excel in xlsx xml files. This is a base "class" or set of
// functionality for all of the other xml writing structs.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2024, John McNamara, jmcnamara@cpan.org

mod tests;

use std::borrow::Cow;
use std::io::{Cursor, Write};
use std::str;

pub(crate) const XML_WRITE_ERROR: &str = "Couldn't write to xml file";

#[derive(Clone)]
pub struct XMLWriter {
    pub(crate) xmlfile: Cursor<Vec<u8>>,
}

impl Default for XMLWriter {
    fn default() -> Self {
        Self::new()
    }
}

// Base XML writing struct used to write xlsx sub-component xml file.
impl XMLWriter {
    pub(crate) fn new() -> XMLWriter {
        let buf: Vec<u8> = Vec::with_capacity(2048);
        let xmlfile = Cursor::new(buf);

        XMLWriter { xmlfile }
    }

    // Helper function to read back stored xml data for tests.
    #[allow(dead_code)]
    pub(crate) fn read_to_str(&mut self) -> &str {
        str::from_utf8(self.xmlfile.get_ref()).unwrap()
    }

    pub(crate) fn read_to_string(&mut self) -> String {
        str::from_utf8(self.xmlfile.get_ref()).unwrap().to_string()
    }

    // Reset the memory buffer, usually between saves.
    pub(crate) fn reset(&mut self) {
        self.xmlfile.get_mut().clear();
        self.xmlfile.set_position(0);
    }

    // Write an XML file declaration.
    pub(crate) fn xml_declaration(&mut self) {
        self.xmlfile
            .write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
            .expect(XML_WRITE_ERROR);
    }

    // Write an XML start tag without attributes.
    pub(crate) fn xml_start_tag_only(&mut self, tag: &str) {
        write!(&mut self.xmlfile, "<{tag}>").expect(XML_WRITE_ERROR);
    }

    // Write an XML start tag with attributes.
    pub(crate) fn xml_start_tag<T>(&mut self, tag: &str, attributes: &[T])
    where
        T: IntoAttribute,
    {
        write!(&mut self.xmlfile, "<{tag}").expect(XML_WRITE_ERROR);

        for attribute in attributes {
            attribute.write_to(&mut self.xmlfile);
        }

        self.xmlfile.write_all(b">").expect(XML_WRITE_ERROR);
    }

    // Write an XML end tag.
    pub(crate) fn xml_end_tag(&mut self, tag: &str) {
        write!(&mut self.xmlfile, "</{tag}>").expect(XML_WRITE_ERROR);
    }

    // Write an empty XML tag without attributes.
    pub(crate) fn xml_empty_tag_only(&mut self, tag: &str) {
        write!(&mut self.xmlfile, "<{tag}/>").expect(XML_WRITE_ERROR);
    }

    // Write an empty XML tag with attributes.
    pub(crate) fn xml_empty_tag<T>(&mut self, tag: &str, attributes: &[T])
    where
        T: IntoAttribute,
    {
        write!(&mut self.xmlfile, "<{tag}").expect(XML_WRITE_ERROR);

        for attribute in attributes {
            attribute.write_to(&mut self.xmlfile);
        }

        self.xmlfile.write_all(b"/>").expect(XML_WRITE_ERROR);
    }

    // Write an XML element containing data without attributes.
    pub(crate) fn xml_data_element_only(&mut self, tag: &str, data: &str) {
        write!(
            &mut self.xmlfile,
            "<{}>{}</{}>",
            tag,
            escape_xml_data(data),
            tag
        )
        .expect(XML_WRITE_ERROR);
    }

    // Write an XML element containing data with attributes.
    pub(crate) fn xml_data_element<T>(&mut self, tag: &str, data: &str, attributes: &[T])
    where
        T: IntoAttribute,
    {
        write!(&mut self.xmlfile, "<{tag}").expect(XML_WRITE_ERROR);

        for attribute in attributes {
            attribute.write_to(&mut self.xmlfile);
        }

        write!(&mut self.xmlfile, ">{}</{}>", escape_xml_data(data), tag).expect(XML_WRITE_ERROR);
    }

    // Optimized tag writer for shared strings <si> elements.
    pub(crate) fn xml_si_element(&mut self, string: &str, preserve_whitespace: bool) {
        if preserve_whitespace {
            write!(
                &mut self.xmlfile,
                r#"<si><t xml:space="preserve">{}</t></si>"#,
                escape_xml_data(&escape_xml_escapes(string))
            )
            .expect(XML_WRITE_ERROR);
        } else {
            write!(
                &mut self.xmlfile,
                "<si><t>{}</t></si>",
                escape_xml_data(&escape_xml_escapes(string))
            )
            .expect(XML_WRITE_ERROR);
        }
    }

    // Write <si> element for rich strings.
    pub(crate) fn xml_rich_si_element(&mut self, string: &str) {
        write!(&mut self.xmlfile, r#"<si>{string}</si>"#).expect(XML_WRITE_ERROR);
    }

    // Write the theme string to the theme file.
    pub(crate) fn write_theme(&mut self, theme: &str) {
        writeln!(&mut self.xmlfile, "{theme}").expect(XML_WRITE_ERROR);
    }

    // Write a string with escaped XML data.
    pub(crate) fn xml_raw_string(&mut self, data: &str) {
        self.xmlfile
            .write_all(data.as_bytes())
            .expect(XML_WRITE_ERROR);
    }
}

// Escape XML characters in attributes.
pub(crate) fn escape_attributes(attribute: &str) -> Cow<str> {
    escape_string(attribute, match_attribute_html_char)
}

// Escape XML characters in data sections of tags.
pub(crate) fn escape_xml_data(data: &str) -> Cow<str> {
    escape_string(data, match_xml_char)
}

// Escape non-url characters in a hyperlink/url.
pub(crate) fn escape_url(data: &str) -> Cow<str> {
    escape_string(data, match_url_char)
}

// -----------------------------------------------------------------------
// Helper functions. Mainly for string escaping.
// -----------------------------------------------------------------------

// Match function for escape_attributes().
fn match_attribute_html_char(ch: char) -> Option<&'static str> {
    match ch {
        '&' => Some("&amp;"),
        '"' => Some("&quot;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),
        '\n' => Some("&#xA;"),
        _ => None,
    }
}

// Match function for escape_xml_data().
//
// Note, this is different from match_attribute_html_char() because double
// quotes and newline are not escaped by Excel.
//
// We need to mimic Excel by escaping control and non-printing characters in the
// range '\x00' - '\x1F'.
fn match_xml_char(ch: char) -> Option<&'static str> {
    match ch {
        // Standard XML escapes.
        '&' => Some("&amp;"),
        '<' => Some("&lt;"),
        '>' => Some("&gt;"),

        // Excel escapes control characters and other non-printing characters in
        // the range '\x00' - '\x1F' with _xHHHH_.
        '\x00' => Some("_x0000_"),
        '\x01' => Some("_x0001_"),
        '\x02' => Some("_x0002_"),
        '\x03' => Some("_x0003_"),
        '\x04' => Some("_x0004_"),
        '\x05' => Some("_x0005_"),
        '\x06' => Some("_x0006_"),
        '\x07' => Some("_x0007_"),
        '\x08' => Some("_x0008_"),
        // No escape required for '\x09' = '\t'
        // No escape required for '\x0A' = '\n'
        '\x0B' => Some("_x000B_"),
        '\x0C' => Some("_x000C_"),
        '\x0D' => Some("_x000D_"),
        '\x0E' => Some("_x000E_"),
        '\x0F' => Some("_x000F_"),
        '\x10' => Some("_x0010_"),
        '\x11' => Some("_x0011_"),
        '\x12' => Some("_x0012_"),
        '\x13' => Some("_x0013_"),
        '\x14' => Some("_x0014_"),
        '\x15' => Some("_x0015_"),
        '\x16' => Some("_x0016_"),
        '\x17' => Some("_x0017_"),
        '\x18' => Some("_x0018_"),
        '\x19' => Some("_x0019_"),
        '\x1A' => Some("_x001A_"),
        '\x1B' => Some("_x001B_"),
        '\x1C' => Some("_x001C_"),
        '\x1D' => Some("_x001D_"),
        '\x1E' => Some("_x001E_"),
        '\x1F' => Some("_x001F_"),

        _ => None,
    }
}

// Match the url characters that Excel escapes.
fn match_url_char(ch: char) -> Option<&'static str> {
    match ch {
        '%' => Some("%25"),
        '"' => Some("%22"),
        ' ' => Some("%20"),
        '<' => Some("%3c"),
        '>' => Some("%3e"),
        '[' => Some("%5b"),
        ']' => Some("%5d"),
        '^' => Some("%5e"),
        '`' => Some("%60"),
        '{' => Some("%7b"),
        '}' => Some("%7d"),
        _ => None,
    }
}

// Generic escape function with function pointer for the required handler.
fn escape_string<F>(original: &str, char_handler: F) -> Cow<str>
where
    F: FnOnce(char) -> Option<&'static str> + Copy,
{
    for (i, ch) in original.char_indices() {
        if char_handler(ch).is_some() {
            let mut escaped_string = original[..i].to_string();
            let remaining = &original[i..];
            escaped_string.reserve(remaining.len());

            for ch in remaining.chars() {
                match char_handler(ch) {
                    Some(escaped_char) => escaped_string.push_str(escaped_char),
                    None => escaped_string.push(ch),
                };
            }

            return Cow::Owned(escaped_string);
        }
    }

    Cow::Borrowed(original)
}

// Excel escapes control characters with _xHHHH_, see match_xml_char() above. As
// a result it also escapes any literal strings of that type by encoding the
// leading underscore. So  "_x0000_" -> _x005F_x0000_.
fn escape_xml_escapes(original: &str) -> Cow<str> {
    if !original.contains("_x00") {
        return Cow::Borrowed(original);
    }

    let string_end = original.len();
    let escape_length = "_x0000_".len();
    let mut escaped_string = original.to_string();

    // Match from right so we can escape target string at the same indices.
    let matches: Vec<_> = original.rmatch_indices("_x00").collect();

    for (index, _) in matches {
        if index + escape_length > string_end {
            continue;
        }

        // Check that the digits in _xABCD_ are a valid hex code.
        if original[index + 2..index + 6]
            .chars()
            .all(|c| c.is_ascii_hexdigit())
        {
            escaped_string.replace_range(index..index, "_x005F");
        }
    }

    if escaped_string == original {
        return Cow::Borrowed(original);
    }

    Cow::Owned(escaped_string)
}

// Trait to write attribute tuple values to an XML file.
pub(crate) trait IntoAttribute {
    fn write_to(&self, xmlfile: &mut Cursor<Vec<u8>>);
}

impl IntoAttribute for (&str, &str) {
    fn write_to(&self, xmlfile: &mut Cursor<Vec<u8>>) {
        write!(xmlfile, r#" {}="{}""#, self.0, escape_attributes(self.1)).expect(XML_WRITE_ERROR);
    }
}

impl IntoAttribute for (&str, String) {
    fn write_to(&self, xmlfile: &mut Cursor<Vec<u8>>) {
        write!(xmlfile, r#" {}="{}""#, self.0, escape_attributes(&self.1)).expect(XML_WRITE_ERROR);
    }
}
